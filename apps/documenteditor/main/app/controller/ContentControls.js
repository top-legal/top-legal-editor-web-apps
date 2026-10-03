/*
 * top.legal — host-driven content controls (Word template fields)
 * ------------------------------------------------------------------
 * Lets the embedding page link text in the document to one of its input fields or conditions, by
 * wrapping it in a Word content control (w:sdt) whose TAG names the field. The document itself is
 * the record of what is linked: tags survive save, reload, download and a round trip through Word,
 * so the host keeps no separate mapping.
 *
 *   tl:f:<inputFieldID>              the text is replaced by the field's value when drafting
 *   tl:c:<inputFieldID>=<optionKey>  the wrapped text is kept only when that option is chosen
 *   tl:c:<inputFieldID>=<k1>|<k2>    ... when ANY of those options is chosen (AND = nest controls)
 *   tl:t:<conditionalTextID>         the text is replaced by a playbook conditional text's output
 *                                    (resolved by the host, nested conditions included)
 *   tl:s:internal | tl:s:external    a party's signature area (block); the host turns it into the
 *                                    signing anchor when a contract is drafted
 *
 * Signature areas can also be DRAGGED in from the host page: a drag carrying
 * `application/x-tl-signature` ({ tag, alias }) is caught here before the SDK sees it, the cursor is
 * put at the drop point by replaying the click there, and the control is inserted at it.
 *
 * A condition may be inline (a phrase inside a sentence) or block (whole paragraphs / a clause).
 *
 * This controller knows nothing about fields, playbooks or drafting. It wraps, lists, selects and
 * unwraps controls whose tag starts with `tl:`, and reports precisely what happened.
 *
 * Same startup discipline as DocumentEdits.js — read the note on onLaunch there. Nothing here may
 * throw into the editor's startup, the app:ready handler defers to a later tick, and the
 * registration sits LAST in app.js. Failing alone is always the correct outcome: the host's panel
 * simply never sees 'ready' and says the editor cannot link fields.
 *
 * FILL is the exception to one-call-per-operation: drafting sets many fields at once, so it works on
 * the document model inside ONE action (one lock check over every affected paragraph, one undo
 * step, one co-editing change) instead of calling asc_SetContentControlText per control — which is
 * exactly the pattern the spike showed silently dropping all but the first.
 *
 * ONE OPERATION PER MESSAGE. Measured in the co-editing spike: several asc_* content-control calls
 * in the same tick silently drop all but the first (the co-editing lock check is asynchronous). The
 * host sends user actions, which are naturally seconds apart, and never batches them.
 *
 * Wire protocol (both directions carry __tl):
 *   app    -> editor : { __tl:'tl-office-fields', type:'ping' }
 *   app    -> editor : { __tl:'tl-office-fields', type:'list',    requestId }
 *   app    -> editor : { __tl:'tl-office-fields', type:'insert',  requestId, tag, alias, block? }
 *   app    -> editor : { __tl:'tl-office-fields', type:'select',  requestId, id }
 *   app    -> editor : { __tl:'tl-office-fields', type:'unlink',  requestId, id }
 *   app    -> editor : { __tl:'tl-office-fields', type:'selection', requestId }
 *   app    -> editor : { __tl:'tl-office-fields', type:'colors', requestId, field:[r,g,b], text:[r,g,b], condition:[r,g,b] }
 *                       Shading per kind (opaque: a docx keeps no alpha, so the host pre-blends). Applied to
 *                       new controls and, when editable, once to existing ones whose shading differs.
 *   app    -> editor : { __tl:'tl-office-fields', type:'fill', requestId, values:{ <fieldKey> | t:<conditionalTextID>: <text> | { text, lines, list, paras, clear } } }
 *                       An inline control takes `text` (or `lines`, joined by line breaks); a BLOCK control given a `list` becomes one
 *                       bullet paragraph per entry (a multiple-choice answer; numbered with `numbered:true`), styled like the
 *                       template's first paragraph in that block. `paras` does the same without bullets (a
 *                       conditional text's paragraphs); `clear` empties the control (its alias shows).
 *   app    -> editor : { __tl:'tl-office-fields', type:'applyConditions', requestId, answers:{ <fieldID>: [<optionKey>] }, dropTags?:[<tag>] }
 *                       After removing, the numbering is healed (see _numberingPlan): result carries `renumbered`.
 *   app    -> editor : { __tl:'tl-office-fields', type:'sections', requestId, show:{ <tag>: true|false } }
 *                       Live dynamic sections while drafting (`tl:c:s.<conditionID>=show` blocks): hide empties
 *                       the block, show puts its paragraphs back. Each section's paragraphs are stored once, the
 *                       first time this is called, inside the document (custom XML part), so a reload or another
 *                       editor can still bring a hidden one back. Numbering is healed after every change.
 *                       Tags left out are untouched. The host refills fields afterwards.
 *   app    -> editor : { __tl:'tl-office-fields', type:'numberingCheck', requestId }
 *                       Which numbering heals when a section goes: result carries lists, typed and
 *                       unrecognized (up to 10 previews of numbered-looking paragraphs that will not be renumbered).
 *   editor -> app    : { __tl:'tl-office-fields', type:'ready', version, canEdit }
 *   editor -> app    : { __tl:'tl-office-fields', type:'result', requestId, ok, reason?, id?, text?, controls? }
 *
 * `controls` is [{ id, tag, alias, block, text }] for every tl: control, in document order.
 */
define([
    'core'
], function () {
    'use strict';

    DE.Controllers = DE.Controllers || {};

    var CHANNEL = 'tl-office-fields';
    var SUPPORTED_VERSION = 1;
    var TAG_PREFIX = 'tl:';
    // Tags reach the file and the host; anything outside this shape is refused rather than
    // sanitised. Field ids and option keys are uuid/uniqid-like.
    // A field key may name a PART of a structured field (`__my_company.name`,
    // `<fieldID>.full_address.city`), hence the dot.
    var OPTION_KEY = '[A-Za-z0-9_ .-]{1,100}';
    var TAG_RE = new RegExp('^tl:(f:[A-Za-z0-9_.-]{1,160}|c:[A-Za-z0-9_.-]{1,160}=' + OPTION_KEY + '(\\|' + OPTION_KEY + '){0,19}|t:[A-Za-z0-9_.-]{1,160}|s:(internal|external))$');
    var SIGNATURE_TAG_RE = /^tl:s:(internal|external)$/;
    var DRAG_TYPE = 'application/x-tl-signature';
    var MAX_ALIAS = 120;
    // Text preview per control in a list — enough for the panel, never the whole clause.
    var MAX_TEXT = 200;
    // Shading so linked text is visible at rest, not only when the cursor is inside it.
    var HIGHLIGHT = [255, 236, 179];
    // Frame colour of a condition (shown on hover/focus), so it reads apart from a field.
    var CONDITION_COLOR = [230, 81, 0];
    // Tag prefix -> key of the host's `colors` message.
    var KINDS = { 'tl:f:': 'field', 'tl:t:': 'text', 'tl:c:': 'condition' };

    DE.Controllers.ContentControls = Backbone.Controller.extend(_.extend({
        models: [],
        collections: [],
        views: [],

        initialize: function () {
            try {
                this.addListeners({});
            } catch (e) { /* inert rather than fatal */ }
        },

        onLaunch: function () {
            var me = this;
            me.api = null;
            me.appOptions = null;
            me._started = false;
            try {
                if (!window.Common || !Common.NotificationCenter || typeof Common.NotificationCenter.on !== 'function') return;
                Common.NotificationCenter.on('app:ready', function (appOptions) {
                    try { me.appOptions = appOptions || {}; } catch (e) { /* not worth failing over */ }
                    setTimeout(function () { me._start(); }, 0);
                });
            } catch (e) { /* no bridge this session */ }
        },

        _start: function () {
            var me = this;
            if (me._started) return;
            me._started = true;
            try {
                me.api = me.getApplication().getController('Main').api;
            } catch (e) {
                me.api = null;
            }
            try {
                me._listen();
                me._listenDrop();
                me._postReady();
            } catch (e) { /* the host simply never sees 'ready' */ }
        },

        // ===============================================================
        // Host messaging (duplicated from DocumentEdits on purpose — see the note there)
        // ===============================================================

        parentOrigin: function () {
            if (this._parentOrigin !== undefined) return this._parentOrigin;
            var p = null;
            try {
                p = new URLSearchParams(window.location.search).get('parentOrigin');
                p = p ? new URL(p).origin : null;
            } catch (e) { p = null; }
            this._parentOrigin = p || null;
            return this._parentOrigin;
        },

        _post: function (msg) {
            try {
                var origin = this.parentOrigin();
                if (!origin || window.parent === window) return;
                msg.__tl = CHANNEL;
                window.parent.postMessage(msg, origin);
            } catch (e) { /* nothing else is affected */ }
        },

        _canEdit: function () {
            var o = this.appOptions || {};
            return !!(this.api && o.isEdit);
        },

        _postReady: function () {
            this._post({ type: 'ready', version: SUPPORTED_VERSION, canEdit: this._canEdit() });
        },

        _listen: function () {
            var me = this;
            if (me._listening) return;
            me._listening = true;
            window.addEventListener('message', function (e) {
                try {
                    if (!me.parentOrigin() || e.origin !== me.parentOrigin()) return;
                    if (e.source !== window.parent) return;
                    var d = e.data;
                    if (!d || d.__tl !== CHANNEL) return;
                    if (d.type === 'ping') { me._postReady(); return; }
                    me._handle(d);
                } catch (err) { /* the host times out; nothing else is affected */ }
            });
        },

        _result: function (requestId, body) {
            var msg = _.extend({ type: 'result', requestId: requestId }, body);
            // Every answer carries the current list, so the host never has to ask twice.
            try { msg.controls = this._list(); } catch (e) { msg.controls = []; }
            this._post(msg);
        },

        _handle: function (d) {
            var me = this;
            var requestId = d.requestId;
            if (!me.api) { me._result(requestId, { ok: false, reason: 'notReady' }); return; }
            me._ensureHighlight();
            switch (d.type) {
                case 'list': me._result(requestId, { ok: true }); return;
                case 'selection': me._result(requestId, { ok: true, text: me._selectedText() }); return;
                case 'insert': me._result(requestId, me._insert(d)); return;
                case 'select': me._result(requestId, me._select(d.id)); return;
                case 'unlink': me._result(requestId, me._unlink(d.id)); return;
                case 'remove': me._result(requestId, me._remove(d.id)); return;
                case 'colors': me._result(requestId, me._setColors(d)); return;
                case 'fill': me._result(requestId, me._fill(d.values)); return;
                case 'applyConditions': me._result(requestId, me._applyConditions(d.answers, d.dropTags)); return;
                case 'sections': me._result(requestId, me._sections(d.show)); return;
                case 'numberingCheck': me._result(requestId, me._numberingCheck()); return;
                default: me._result(requestId, { ok: false, reason: 'unsupported' });
            }
        },

        // ===============================================================
        // Content controls
        // ===============================================================

        _doc: function () {
            return this.api && this.api.WordControl && this.api.WordControl.m_oLogicDocument;
        },

        // Set once per session: shading is a document-wide display option, not part of any control.
        _ensureHighlight: function () {
            if (this._highlighted) return;
            try {
                this.api.asc_SetGlobalContentControlShowHighlight(true, HIGHLIGHT[0], HIGHLIGHT[1], HIGHLIGHT[2]);
                this._highlighted = true;
            } catch (e) { /* cosmetic */ }
        },

        _kindOf: function (tag) {
            var prefix = typeof tag === 'string' ? tag.slice(0, 5) : '';
            return KINDS[prefix] || null;
        },

        /** The host's [r,g,b] for this tag's kind, validated; null = keep the editor default. */
        _colorFor: function (tag) {
            var kind = this._kindOf(tag);
            var c = kind && this._colors ? this._colors[kind] : null;
            if (!_.isArray(c) || c.length < 3) return null;
            for (var i = 0; i < 3; i++) if (typeof c[i] !== 'number' || c[i] < 0 || c[i] > 255) return null;
            return [c[0] | 0, c[1] | 0, c[2] | 0];
        },

        /**
         * Store the host's shading per kind and bring existing controls in line, in ONE action. Only
         * controls whose shading differs are touched, so a template opened a second time is unchanged.
         */
        _setColors: function (d) {
            var me = this;
            me._colors = { field: d.field, text: d.text, condition: d.condition };
            if (!me._canEdit()) return { ok: true, recolored: 0 };
            var doc = me._doc();
            var todo = [];
            me._ownControls().forEach(function (cc) {
                var c = me._colorFor(cc.GetTag());
                if (!c || typeof cc.setShdColor !== 'function') return;
                var cur = cc.getShdColor && cc.getShdColor();
                if (cur && cur.r === c[0] && cur.g === c[1] && cur.b === c[2] && cur.a === 255) return;
                todo.push({ cc: cc, color: c });
            });
            if (!todo.length) return { ok: true, recolored: 0 };
            try {
                var locked = doc.Document_Is_SelectionLocked(AscCommon.changestype_None, {
                    Type: AscCommon.changestype_2_ElementsArray_and_Type,
                    Elements: todo.map(function (x) { return x.cc.IsBlockLevel() ? x.cc : x.cc.GetParagraph(); }).filter(Boolean),
                    CheckType: AscCommon.changestype_Paragraph_Content,
                });
                // Cosmetic: a co-editor holding the text just means the old colour stays for now.
                if (locked) return { ok: true, recolored: 0 };
                doc.StartAction(AscDFH.historydescription_Document_SetContentControlText);
                todo.forEach(function (x) {
                    x.cc.setShdColor(new AscWord.CDocumentColorA(x.color[0], x.color[1], x.color[2], 255));
                    if (typeof x.cc.SetColor === 'function') x.cc.SetColor(new AscWord.CDocumentColor(x.color[0], x.color[1], x.color[2]));
                });
                doc.Recalculate();
                doc.UpdateInterface();
                doc.FinalizeAction();
                return { ok: true, recolored: todo.length };
            } catch (e) {
                return { ok: true, recolored: 0 };
            }
        },

        _ownControls: function () {
            var doc = this._doc();
            if (!doc || typeof doc.GetAllContentControls !== 'function') return [];
            return doc.GetAllContentControls().filter(function (cc) {
                var tag = cc && typeof cc.GetTag === 'function' ? cc.GetTag() : '';
                return typeof tag === 'string' && tag.indexOf(TAG_PREFIX) === 0;
            });
        },

        _list: function () {
            return this._ownControls().map(function (cc) {
                var text = '';
                try { text = String(cc.GetInnerText ? cc.GetInnerText() : '').slice(0, MAX_TEXT); } catch (e) { text = ''; }
                return {
                    id: cc.GetId(),
                    tag: cc.GetTag(),
                    alias: cc.GetAlias ? cc.GetAlias() : '',
                    block: typeof cc.IsBlockLevel === 'function' ? cc.IsBlockLevel() : false,
                    text: text,
                };
            });
        },

        _selectedText: function () {
            // bClearText=true returns null once the selection spans paragraphs, so a multi-paragraph
            // selection read as "nothing selected". bClearText=false joins paragraphs with \r\n.
            try {
                var text = this.api.asc_GetSelectedText(false);
                return String(text || '').replace(/[\r\n]+$/, '').slice(0, MAX_TEXT);
            } catch (e) { return ''; }
        },

        _findOwn: function (id) {
            if (typeof id !== 'string') return null;
            var found = null;
            this._ownControls().forEach(function (cc) { if (cc.GetId() === id) found = cc; });
            return found;
        },

        /**
         * Wrap the current selection (or insert an empty control at the cursor) with a tagged
         * control. Inline for a field value; block for a condition, which must be able to remove
         * whole paragraphs.
         */
        _insert: function (d) {
            if (!this._canEdit()) return { ok: false, reason: 'readOnly' };
            var tag = typeof d.tag === 'string' ? d.tag : '';
            if (!TAG_RE.test(tag)) return { ok: false, reason: 'badTag' };
            var alias = String(d.alias || '').slice(0, MAX_ALIAS);
            try {
                var pr = new AscCommon.CContentControlPr();
                pr.put_Tag(tag);
                pr.put_Alias(alias);
                pr.put_Appearance(Asc.c_oAscSdtAppearance ? Asc.c_oAscSdtAppearance.Frame : 1);
                if (alias && typeof pr.put_PlaceholderText === 'function') pr.put_PlaceholderText(alias);
                var shade = this._colorFor(tag);
                if (shade) {
                    // Applied by SetContentControlPr on creation; opaque, so it survives the docx.
                    pr.ShdColor = { r: shade[0], g: shade[1], b: shade[2], a: 255 };
                    if (typeof pr.put_Color === 'function') pr.put_Color(shade[0], shade[1], shade[2]);
                } else if (tag.indexOf('tl:c:') === 0 && typeof pr.put_Color === 'function') {
                    pr.put_Color(CONDITION_COLOR[0], CONDITION_COLOR[1], CONDITION_COLOR[2]);
                }
                var type = d.block ? Asc.c_oAscSdtLevelType.Block : Asc.c_oAscSdtLevelType.Inline;
                var created = this.api.asc_AddContentControl(type, pr);
                // null = the selection was locked by a co-editor, or could not take a control
                // (e.g. inside another control of the wrong level). Say so; never guess.
                if (!created) return { ok: false, reason: 'locked' };
                return { ok: true, id: created.get_InternalId ? created.get_InternalId() : undefined };
            } catch (e) {
                return { ok: false, reason: 'error' };
            }
        },

        _select: function (id) {
            var cc = this._findOwn(id);
            if (!cc) return { ok: false, reason: 'notFound' };
            try {
                var doc = this._doc();
                doc.RemoveSelection();
                cc.SelectContentControl();
                doc.UpdateSelection();
                doc.UpdateInterface();
                // Scroll the selection into view.
                if (typeof doc.private_UpdateCursorXY === 'function') doc.private_UpdateCursorXY(true, true);
                return { ok: true };
            } catch (e) {
                return { ok: false, reason: 'error' };
            }
        },

        // Remove deletes the control WITH its content (a signature area has no text worth keeping).
        _remove: function (id) {
            if (!this._canEdit()) return { ok: false, reason: 'readOnly' };
            var cc = this._findOwn(id);
            if (!cc) return { ok: false, reason: 'notFound' };
            try {
                this.api.asc_RemoveContentControl(cc.GetId());
                return { ok: !this._findOwn(id), reason: this._findOwn(id) ? 'locked' : undefined };
            } catch (e) {
                return { ok: false, reason: 'error' };
            }
        },

        // ===============================================================
        // Drag and drop from the host
        // ===============================================================

        _isOurDrag: function (e) {
            var types = e && e.dataTransfer && e.dataTransfer.types;
            if (!types) return false;
            for (var i = 0; i < types.length; i++) if (types[i] === DRAG_TYPE) return true;
            return false;
        },

        /**
         * Capture phase on the window, so the SDK's own drop handling (which would paste the drag's
         * text) never sees our drags. Anything else passes through untouched.
         */
        _listenDrop: function () {
            var me = this;
            if (me._dropListening) return;
            me._dropListening = true;
            var over = function (e) {
                if (!me._isOurDrag(e)) return;
                e.preventDefault();
                e.stopPropagation();
                try { e.dataTransfer.dropEffect = me._canEdit() ? 'copy' : 'none'; } catch (err) { /* cosmetic */ }
            };
            window.addEventListener('dragenter', over, true);
            window.addEventListener('dragover', over, true);
            window.addEventListener('drop', function (e) {
                if (!me._isOurDrag(e)) return;
                e.preventDefault();
                e.stopPropagation();
                var d = null;
                try { d = JSON.parse(e.dataTransfer.getData(DRAG_TYPE) || 'null'); } catch (err) { d = null; }
                if (!d || typeof d.tag !== 'string' || !SIGNATURE_TAG_RE.test(d.tag)) return;
                me._dropAt(e, d);
            }, true);
        },

        /**
         * Put the cursor where the drag was released by replaying a click on the element under it —
         * the SDK's own hit-testing, so page, zoom and scroll are handled as for any click — then
         * insert. The insert waits a beat: the click's selection update must land first.
         */
        _dropAt: function (e, d) {
            var me = this;
            var result = function (body) { me._post(_.extend({ type: 'dropped', tag: d.tag }, body)); };
            if (!me._canEdit()) { result({ ok: false, reason: 'readOnly' }); return; }
            try {
                var target = document.elementFromPoint(e.clientX, e.clientY) || e.target;
                var init = {
                    bubbles: true, cancelable: true, view: window, detail: 1, button: 0, buttons: 1,
                    clientX: e.clientX, clientY: e.clientY, screenX: e.screenX, screenY: e.screenY,
                };
                target.dispatchEvent(new MouseEvent('mousedown', init));
                init.buttons = 0;
                target.dispatchEvent(new MouseEvent('mouseup', init));
            } catch (err) { /* insert at the current cursor instead */ }
            setTimeout(function () {
                var r = me._insert({ tag: d.tag, alias: d.alias, block: true });
                result(r);
                try { me._post({ type: 'result', requestId: 'drop', ok: r.ok, controls: me._list() }); } catch (err) { /* next poll */ }
            }, 60);
        },

        // Unlink keeps the text and drops only the wrapper — the reverse of insert.
        _unlink: function (id) {
            if (!this._canEdit()) return { ok: false, reason: 'readOnly' };
            var cc = this._findOwn(id);
            if (!cc) return { ok: false, reason: 'notFound' };
            try {
                this.api.asc_RemoveContentControlWrapper(cc.GetId());
                return { ok: !this._findOwn(id), reason: this._findOwn(id) ? 'locked' : undefined };
            } catch (e) {
                return { ok: false, reason: 'error' };
            }
        },

        /**
         * Put field values into every `tl:f:<fieldID>` control. Unchanged text is skipped, so a
         * host that re-sends the whole answer set on every keystroke produces no churn. An empty
         * value leaves the control as it is (the template's own text stays as the placeholder).
         */
        _fill: function (values) {
            if (!this._canEdit()) return { ok: false, reason: 'readOnly' };
            if (!values || typeof values !== 'object') return { ok: false, reason: 'badValues' };
            var doc = this._doc();
            var todo = [];
            this._ownControls().forEach(function (cc) {
                var tag = cc.GetTag();
                var key;
                if (tag.indexOf('tl:f:') === 0) key = tag.slice(5);
                else if (tag.indexOf('tl:t:') === 0) key = 't:' + tag.slice(5);
                else return;
                var v = values[key];
                var isBlock = typeof cc.IsBlockLevel === 'function' && cc.IsBlockLevel();
                var current = '';
                try { current = String(cc.GetInnerText()); } catch (e) { current = ''; }
                if (v && v.clear === true) {
                    // Decided, and nothing to show: empty the control (its alias placeholder shows).
                    if (!(cc.IsPlaceHolder && cc.IsPlaceHolder()) && current.replace(/[\r\n\s]/g, '')) todo.push({ cc: cc, clear: true, block: isBlock });
                    return;
                }
                var text = typeof v === 'string' ? v : (v && typeof v.text === 'string' ? v.text : '');
                // `paras`: plain paragraphs in a block control (a conditional text), no bullets.
                var paras = isBlock && v && _.isArray(v.paras) ? v.paras.filter(function (x) { return typeof x === 'string'; }) : null;
                if (paras && paras.length) {
                    var normP = current.split(/\r?\n/).map(function (l) { return l.trim(); }).join('\n').replace(/\n+$/, '');
                    if (normP !== paras.map(function (x) { return x.trim(); }).join('\n')) {
                        todo.push({ cc: cc, list: paras.slice(0, 200).map(function (x) { return x.slice(0, 5000); }), plain: true });
                    }
                    return;
                }
                var list = v && _.isArray(v.list) ? v.list.filter(function (x) { return typeof x === 'string' && x; }) : null;
                // `lines`: an address block and the like — line breaks INSIDE the field's paragraph.
                var lines = v && _.isArray(v.lines) ? v.lines.filter(function (x) { return typeof x === 'string' && x; }) : null;
                if (isBlock && !list && text) list = [text];
                if (isBlock ? !(list && list.length) : !text) return;
                if (isBlock) {
                    var numbered = !!(v && v.numbered === true);
                    // Compare item by item, ignoring the list glyphs and paragraph marks — but a switch
                    // between bullets and numbers (the glyph kind) still refills.
                    var rows = current.split(/\r?\n/).filter(function (l) { return l.trim(); });
                    var norm = rows.map(function (l) { return l.replace(/^[^\t]*\t/, '').trim(); }).join('\n');
                    var glyph = rows.length && /^[^\t]*\t/.test(rows[0]) ? rows[0].split('\t')[0] : null;
                    var sameKind = glyph !== null && /\d/.test(glyph) === numbered;
                    if (norm === list.join('\n') && sameKind) return;
                    todo.push({ cc: cc, list: list.slice(0, 200).map(function (x) { return x.slice(0, 2000); }), numbered: numbered });
                } else if (lines && lines.length > 1) {
                    if (current.replace(/\r/g, '') !== lines.join('\n')) todo.push({ cc: cc, lines: lines.slice(0, 50) });
                } else if (current !== text) {
                    todo.push({ cc: cc, text: text.slice(0, 5000) });
                }
            });
            if (!todo.length) return { ok: true, filled: 0 };
            try {
                var elements = todo.map(function (x) { return x.list || (x.clear && x.block) ? x.cc : x.cc.GetParagraph(); }).filter(Boolean);
                var locked = doc.Document_Is_SelectionLocked(AscCommon.changestype_None, {
                    Type: AscCommon.changestype_2_ElementsArray_and_Type,
                    Elements: elements,
                    CheckType: AscCommon.changestype_Paragraph_Content,
                });
                if (locked) return { ok: false, reason: 'locked' };
                doc.StartAction(AscDFH.historydescription_Document_SetContentControlText);
                var levels = {};
                todo.forEach(function (x) {
                    if (x.clear) {
                        if (x.block) new AscBuilder.ApiBlockLvlSdt(x.cc).GetContent().RemoveAllElements();
                        else x.cc.MakeSingleRunElement(true);
                        return;
                    }
                    if (x.list) {
                        var kind = x.numbered ? 'numbered' : 'bullet';
                        // One numbering per numbered list, so each restarts at 1; bullets can share.
                        var level = null;
                        if (!x.plain) level = kind === 'bullet' && levels.bullet ? levels.bullet : new AscBuilder.ApiDocument(doc).CreateNumbering(kind).GetLevel(0);
                        if (kind === 'bullet') levels.bullet = level;
                        this._fillList(x.cc, x.list, level);
                        return;
                    }
                    if (x.cc.IsPlaceHolder && x.cc.IsPlaceHolder()) x.cc.ReplacePlaceHolderWithContent();
                    var run = x.cc.MakeSingleRunElement(true);
                    if (!run) return;
                    if (x.lines) {
                        var apiRun = new AscBuilder.ApiRun(run);
                        x.lines.forEach(function (line, i) {
                            if (i) apiRun.AddLineBreak();
                            apiRun.AddText(line.slice(0, 2000));
                        });
                        return;
                    }
                    run.AddText(x.text);
                }, this);
                doc.Recalculate();
                doc.UpdateInterface();
                doc.FinalizeAction();
                return { ok: true, filled: todo.length };
            } catch (e) {
                return { ok: false, reason: 'error' };
            }
        },

        /**
         * Replace a block control's paragraphs with one bullet paragraph per entry (or plain
         * paragraphs when `bulletLevel` is null). Each copies the
         * block's first paragraph (paragraph style + the formatting of its first run), so the list
         * reads like the template text it replaces. Runs inside _fill's action.
         */
        _fillList: function (cc, items, bulletLevel) {
            if (cc.IsPlaceHolder && cc.IsPlaceHolder()) cc.ReplacePlaceHolderWithContent();
            var content = new AscBuilder.ApiBlockLvlSdt(cc).GetContent();
            var proto = content.GetElement(0);
            var paras = items.map(function (text) {
                var p = proto.Copy();
                for (var i = p.GetElementsCount() - 1; i > 0; i--) p.RemoveElement(i);
                var run = p.GetElement(0);
                if (run && typeof run.ClearContent === 'function') {
                    run.ClearContent();
                    run.AddText(text);
                } else {
                    p.AddText(text);
                }
                if (bulletLevel) p.SetNumbering(bulletLevel);
                return p;
            });
            content.RemoveAllElements();
            paras.forEach(function (p) { content.Push(p); });
            // RemoveAllElements leaves one empty paragraph behind (a document content is never
            // empty); drop it now that the list is in.
            if (content.GetElementsCount() > paras.length) content.RemoveElement(0);
        },

        /**
         * Finish-drafting step for conditions: a `tl:c:<fieldID>=<k1>|<k2>` control none of whose
         * options is among the answer's keys is deleted with its content; one with a chosen option is
         * kept. Fields with no answer at all are left untouched, so an unanswered question never
         * silently deletes a clause. Nested conditions: an inner control inside a dropped outer one
         * goes with it (skipped here, its id no longer resolves).
         */
        _applyConditions: function (answers, dropTags) {
            if (!this._canEdit()) return { ok: false, reason: 'readOnly' };
            if (!answers || typeof answers !== 'object') return { ok: false, reason: 'badAnswers' };
            var doc = this._doc();
            var drop = [];
            // Tags the host resolved to nothing (a conditional text with no text for the answer).
            var dropSet = {};
            if (_.isArray(dropTags)) dropTags.forEach(function (t) { if (typeof t === 'string' && TAG_RE.test(t)) dropSet[t] = true; });
            this._ownControls().forEach(function (cc) {
                if (dropSet[cc.GetTag()]) { drop.push(cc); return; }
                // A dynamic section (`tl:c:s.<conditionID>=show`) is answered like any condition: the
                // host evaluates its condition and sends `s.<conditionID>: ['show' | 'hide']`.
                var m = /^tl:c:([^=]+)=(.+)$/.exec(cc.GetTag());
                if (!m) return;
                var chosen = answers[m[1]];
                if (!_.isArray(chosen)) return;
                var keys = m[2].split('|');
                if (!keys.some(function (k) { return chosen.indexOf(k) !== -1; })) drop.push(cc);
            });
            if (!drop.length) return { ok: true, removed: 0 };
            // How the numbering counts BEFORE anything goes; healed against it once the blocks are gone.
            var plan = null;
            try { plan = this._numberingPlan(); } catch (e) { plan = null; }
            try {
                var locked = doc.Document_Is_SelectionLocked(AscCommon.changestype_None, {
                    Type: AscCommon.changestype_2_ElementsArray_and_Type,
                    Elements: drop.map(function (cc) { return cc.IsBlockLevel() ? cc : cc.GetParagraph(); }).filter(Boolean),
                    CheckType: AscCommon.changestype_ContentControl_Remove,
                });
                if (locked) return { ok: false, reason: 'locked' };
                doc.StartAction(AscDFH.historydescription_Document_RemoveContentControl);
                drop.forEach(function (cc) {
                    // An outer drop earlier in document order has already taken this one with it.
                    var liveIds = {};
                    doc.GetAllContentControls().forEach(function (c) { liveIds[c.GetId()] = true; });
                    if (liveIds[cc.GetId()]) doc.RemoveContentControl(cc.GetId());
                });
                var renumbered = 0;
                // Drafting is over: the stored section copies have no further use.
                try { this._dropSectionStore(); } catch (e) { /* a leftover part is harmless */ }
                // Same undo step as the removal. A failure leaves the numbering as Word would show it.
                try { if (plan) renumbered = this._healNumbering(plan); } catch (e) { renumbered = -1; }
                doc.Recalculate();
                doc.UpdateInterface();
                doc.FinalizeAction();
                return { ok: true, removed: drop.length, renumbered: renumbered };
            } catch (e) {
                return { ok: false, reason: 'error' };
            }
        },

        // ===============================================================
        // Live dynamic sections
        // ===============================================================
        //
        // A dynamic section is a block control tagged `tl:c:s.<conditionID>=show`. While drafting, the
        // host says which sections its answers currently show. Hiding empties the block down to one
        // collapsed, unnumbered paragraph (an empty block would show its placeholder); showing puts the
        // section's original paragraphs back. The originals are serialized (document builder JSON) the
        // first time `sections` runs — the drafted document is then still an untouched copy of the
        // template — and kept in a custom XML part, so they survive a reload and reach co-editors.
        // Finishing (applyConditions) removes hidden blocks for good and drops the part.

        _sectionPartNs: 'urn:top-legal:dynamic-sections',

        _isSectionTag: function (tag) { return typeof tag === 'string' && tag.indexOf('tl:c:s.') === 0; },

        /** The stored copies: { <tag>: { json, hidden } }, read from the custom XML part once per session. */
        _sectionStore: function () {
            if (this._sections_) return this._sections_;
            var store = {};
            try {
                var text = this._sectionPartText();
                var m = /<data>([^<]*)<\/data>/.exec(text || '');
                if (m) store = JSON.parse(decodeURIComponent(escape(atob(m[1])))) || {};
            } catch (e) { store = {}; }
            this._sections_ = store;
            return store;
        },

        /** The stored part's XML, or ''. */
        _sectionPartText: function () {
            var mgr = this._doc().getCustomXmlManager();
            if (!mgr) return '';
            for (var i = 0; i < mgr.getCount(); i++) {
                var text = '';
                try { text = mgr.getCustomXMLString(mgr.getCustomXml(i)) || ''; } catch (e) { text = ''; }
                if (text.indexOf(this._sectionPartNs) !== -1) return text;
            }
            return '';
        },

        _saveSectionStore: function () {
            var doc = this._doc();
            var mgr = doc.getCustomXmlManager();
            if (!mgr) return;
            this._dropSectionPart();
            var data = btoa(unescape(encodeURIComponent(JSON.stringify(this._sections_ || {}))));
            mgr.createCustomXml('<tlSections xmlns="' + this._sectionPartNs + '"><data>' + data + '</data></tlSections>');
        },

        _dropSectionPart: function () {
            var mgr = this._doc().getCustomXmlManager();
            if (!mgr) return;
            for (var i = mgr.getCount() - 1; i >= 0; i--) {
                var xml = mgr.getCustomXml(i);
                var text = '';
                try { text = mgr.getCustomXMLString(xml) || ''; } catch (e) { text = ''; }
                if (text.indexOf(this._sectionPartNs) !== -1) mgr.deleteExactXml(xml.itemId || (xml.getUid && xml.getUid()), xml.prefix);
            }
        },

        _dropSectionStore: function () {
            this._dropSectionPart();
            this._sections_ = {};
        },

        /** Exact copies of each section's elements, this session only (styles and lists by identity). */
        _copies: function () {
            if (!this._copies_) this._copies_ = {};
            return this._copies_;
        },

        _copyElements: function (apiContent) {
            var out = [];
            for (var i = 0; i < apiContent.GetElementsCount(); i++) out.push(apiContent.GetElement(i).Copy());
            return out;
        },

        /**
         * Fresh elements to put back: copies of this session's exact copy, or — after a reload — the
         * stored JSON (styles are re-applied by name afterwards; its lists come back as copies, which
         * the numbering heal then brings in line).
         */
        _restoredElements: function (tag, entry) {
            var exact = this._copies()[tag];
            if (exact) return exact.map(function (el) { return el.Copy(); });
            var restored = AscBuilder.Api.FromJSON(entry.json);
            var n = restored && typeof restored.GetElementsCount === 'function' ? restored.GetElementsCount() : 0;
            var out = [];
            // Copies: the parsed elements still belong to the reader's scratch content.
            for (var i = 0; i < n; i++) out.push(restored.GetElement(i).Copy());
            return out;
        },

        /** Paragraph style NAMES in order (style ids are not stable across sessions). */
        _styleNames: function (cc) {
            var doc = this._doc();
            var paras = [];
            cc.GetAllParagraphs({ All: true }, paras);
            return paras.map(function (p) {
                var id = p.Style_Get && p.Style_Get();
                var st = id && doc.Styles.Get(id);
                return st ? st.GetName() : null;
            });
        },

        _applyStyleNames: function (cc, names) {
            if (!Array.isArray(names)) return;
            var doc = this._doc();
            var paras = [];
            cc.GetAllParagraphs({ All: true }, paras);
            paras.forEach(function (p, i) {
                var id = names[i] ? doc.Styles.GetStyleIdByName(names[i]) : null;
                if (id && typeof p.Style_Add === 'function') p.Style_Add(id, true);
            });
        },

        /** One collapsed, unnumbered, empty paragraph: what a hidden section leaves behind until finishing. */
        _collapsedParagraph: function () {
            var para = AscBuilder.Api.CreateParagraph();
            para.SetSpacingBefore(0);
            para.SetSpacingAfter(0);
            para.SetSpacingLine(1, 'exact');
            para.SetNumbering(null);
            para.SetFontSize(1);
            return para;
        },

        _sections: function (show) {
            if (!this._canEdit()) return { ok: false, reason: 'readOnly' };
            if (!show || typeof show !== 'object') return { ok: false, reason: 'badShow' };
            var me = this;
            var doc = me._doc();
            var store = me._sectionStore();
            var stored = false;
            var todo = [];
            me._ownControls().forEach(function (cc) {
                var tag = cc.GetTag();
                if (!me._isSectionTag(tag) || !(cc.IsBlockLevel && cc.IsBlockLevel())) return;
                var entry = store[tag];
                if (!entry) {
                    // First sight: the block still holds the template's paragraphs.
                    var content = new AscBuilder.ApiBlockLvlSdt(cc).GetContent();
                    entry = store[tag] = { json: content.ToJSON(true, false), styles: me._styleNames(cc), hidden: false };
                    me._copies()[tag] = me._copyElements(content);
                    stored = true;
                }
                if (typeof show[tag] !== 'boolean') return;
                if (show[tag] === !entry.hidden) return;
                todo.push({ cc: cc, tag: tag, entry: entry, show: show[tag] });
            });
            if (!todo.length) {
                if (stored) {
                    try { doc.StartAction(AscDFH.historydescription_Document_AddContentControl); me._saveSectionStore(); doc.FinalizeAction(); } catch (e) { /* kept in memory */ }
                }
                return { ok: true, changed: 0 };
            }
            try {
                var locked = doc.Document_Is_SelectionLocked(AscCommon.changestype_None, {
                    Type: AscCommon.changestype_2_ElementsArray_and_Type,
                    Elements: todo.map(function (x) { return x.cc; }),
                    CheckType: AscCommon.changestype_ContentControl_Properties,
                });
                if (locked) return { ok: false, reason: 'locked' };
                var plan = null;
                try { plan = me._numberingPlan(); } catch (e) { plan = null; }
                doc.StartAction(AscDFH.historydescription_Document_AddContentControl);
                todo.forEach(function (x) {
                    var content = new AscBuilder.ApiBlockLvlSdt(x.cc).GetContent();
                    if (x.show) {
                        var elements = me._restoredElements(x.tag, x.entry);
                        if (!elements.length) return;
                        content.RemoveAllElements();
                        elements.forEach(function (el) { content.Push(el); });
                        // RemoveAllElements leaves one empty paragraph behind; the section starts after it.
                        if (content.GetElementsCount() > elements.length) content.RemoveElement(0);
                        if (!me._copies()[x.tag]) me._applyStyleNames(x.cc, x.entry.styles);
                    } else {
                        content.RemoveAllElements();
                        content.Push(me._collapsedParagraph());
                        if (content.GetElementsCount() > 1) content.RemoveElement(0);
                    }
                    x.entry.hidden = !x.show;
                });
                var renumbered = 0;
                try { if (plan) renumbered = me._healNumbering(plan); } catch (e) { renumbered = -1; }
                me._saveSectionStore();
                doc.Recalculate();
                doc.UpdateInterface();
                doc.FinalizeAction();
                return { ok: true, changed: todo.length, renumbered: renumbered };
            } catch (e) {
                try { doc.FinalizeAction(); } catch (e2) { /* no open action */ }
                return { ok: false, reason: 'error', detail: String(e && e.message || e).slice(0, 200) };
            }
        },

        // ===============================================================
        // Numbering heal
        // ===============================================================
        //
        // Templates often restart numbering by hand: "§ 1" in every part, "(1)" in every clause is
        // done by switching paragraphs onto another list (or a list with a start override) rather
        // than by a restart rule. Counters are kept per abstract list, so once a clause is removed,
        // a later run of paragraphs that continued an earlier counter shows a gap or a repeat
        // (§ 1 2 3 5 6 7). Measured on real templates in Word and Euro-Office alike.
        //
        // The plan reads, before removal, what each kind of numbered paragraph (paragraph style +
        // level) counts within: which other kind restarts it (a part restarts its §, a § its (n)),
        // or none (it counts through the whole document). A kind whose numbers do not follow one
        // such rule is left alone. After removal, every paragraph whose number no longer matches its
        // position gets what Euro-Office's own "Restart numbering" does: a copy of its list (same
        // formatting, so nothing looks different) starting at the expected value, for it and the
        // following paragraphs of the same kind, scope and list.

        /** Every numbered body paragraph in document order: { para, kind, numId, lvl, value }. */
        _numberedParagraphs: function () {
            var doc = this._doc();
            var paras = [];
            (doc.Content || []).forEach(function (el) { if (el && typeof el.GetAllParagraphs === 'function') el.GetAllParagraphs({ All: true }, paras); });
            var out = [];
            paras.forEach(function (p) {
                if (!p || typeof p.GetNumPr !== 'function') return;
                var numPr = p.GetNumPr();
                if (!numPr || numPr.NumId === undefined || numPr.NumId === null || numPr.NumId === '0') return;
                var parent = p.GetParent && p.GetParent();
                if (!parent || typeof parent.CalculateNumberingValues !== 'function' || p.GetIndex() === -1) return;
                var lvl = numPr.Lvl || 0;
                var info = parent.CalculateNumberingValues(p, numPr);
                var value = info && typeof info[lvl] === 'number' ? info[lvl] : null;
                if (value === null) return;
                var style = typeof p.Style_Get === 'function' ? p.Style_Get() : null;
                var num = doc.GetNumbering().GetNum(numPr.NumId);
                var kind = style ? 's:' + style + '|' + lvl : 'n:' + (num ? num.GetAbstractNumId() : numPr.NumId) + '|' + lvl;
                out.push({ para: p, kind: kind, numId: numPr.NumId, lvl: lvl, value: value });
            });
            return out;
        },

        /** Per kind: { scope: <kind that restarts it> | null, first: <its first value> }. */
        _numberingPlan: function () {
            var items = this._numberedParagraphs();
            var kinds = _.uniq(items.map(function (x) { return x.kind; }));
            var counts = _.countBy(items, 'kind');
            var plan = {};
            // Does kind k number first, first+1, ... restarting after every `scope` paragraph?
            var follows = function (k, scope, first) {
                var prev = null;
                var restart = true;
                for (var i = 0; i < items.length; i++) {
                    var x = items[i];
                    if (scope !== null && x.kind === scope) { restart = true; continue; }
                    if (x.kind !== k) continue;
                    var expected = restart ? first : prev + 1;
                    if (x.value !== expected) return false;
                    prev = x.value;
                    restart = false;
                }
                return true;
            };
            kinds.forEach(function (k) {
                var first = _.find(items, function (x) { return x.kind === k; }).value;
                if (follows(k, null, first)) { plan[k] = { scope: null, first: first }; return; }
                // The nearest kind that explains every restart: of those that do, the most frequent
                // (a clause restarts its "a." list, not the part the clause sits in).
                var scopes = kinds.filter(function (o) { return o !== k && follows(k, o, first); });
                scopes.sort(function (a, b) { return counts[b] - counts[a]; });
                if (scopes.length) plan[k] = { scope: scopes[0], first: first };
            });
            // Numbers typed as text ("1.", "1.2", "(3)") are healed separately; see _typedPlan.
            Object.defineProperty(plan, 'typed', { value: this._typedPlan(), enumerable: false });
            return plan;
        },

        /** Bring every planned kind back to first, first+1, ... per scope. Returns the paragraphs re-listed. */
        _healNumbering: function (plan) {
            var doc = this._doc();
            var numbering = doc.GetNumbering();
            var changed = 0;
            // Each pass fixes the first wrong paragraph (and its run); bounded by the paragraph count.
            for (var pass = 0; pass < 500; pass++) {
                var items = this._numberedParagraphs();
                var wrong = null;
                var run = [];
                Object.keys(plan).some(function (k) {
                    var rule = plan[k];
                    var pos = 0;
                    for (var i = 0; i < items.length; i++) {
                        var x = items[i];
                        if (rule.scope !== null && x.kind === rule.scope) { pos = 0; if (wrong) break; continue; }
                        if (x.kind !== k) continue;
                        if (wrong) {
                            // The run: same kind, same scope, still on the list the wrong one was on.
                            if (x.numId !== wrong.numId) break;
                            run.push(x);
                            continue;
                        }
                        pos += 1;
                        if (x.value !== rule.first + pos - 1) {
                            wrong = x;
                            wrong.expected = rule.first + pos - 1;
                            run.push(x);
                        }
                    }
                    return !!wrong;
                });
                if (!wrong) return changed + this._healTyped(plan.typed);
                var num = numbering.GetNum(wrong.numId);
                if (!num) return changed;
                // Exactly Euro-Office's "Restart numbering": a copy of the list, started where it should.
                var copy = num.Copy();
                var lvl = copy.GetLvl(wrong.lvl).Copy();
                lvl.Start = wrong.expected;
                copy.SetLvl(lvl, wrong.lvl);
                run.forEach(function (x) { x.para.SetNumPr(copy.GetId(), x.lvl); });
                changed += run.length;
            }
            return changed + this._healTyped(plan.typed);
        },

        // ---------------------------------------------------------------
        // Typed numbers
        // ---------------------------------------------------------------
        //
        // Many templates type their numbers ("1. Definitions", "1.2  The receiving party", "§ 3",
        // "(a)", "II."). Those are plain text, so nothing renumbers them. A paragraph without list
        // numbering whose text starts with one of TYPED_STYLES followed by white space is a typed
        // number of that style. Dotted decimals ("1.", "1.2", "1.2.3") are a hierarchy: a kind per
        // depth, each number its parent's current number plus the next value at its own depth. Every
        // other style is a flat kind with the list rule (counts through, or restarts after the
        // nearest kind that explains every restart). A kind whose numbers do not follow its rule in
        // the document as it stands is left alone. Healing rewrites only the number itself, in
        // place, so the formatting of the run it starts in is kept.

        _romanValue: function (r) {
            var map = { i: 1, v: 5, x: 10, l: 50, c: 100 };
            var s = r.toLowerCase();
            var total = 0;
            for (var i = 0; i < s.length; i++) {
                var v = map[s[i]];
                var next = map[s[i + 1]] || 0;
                if (!v) return null;
                total += v < next ? -v : v;
            }
            return total;
        },

        _romanText: function (n, upper) {
            var parts = [[100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
            var out = '';
            parts.forEach(function (p) { while (n >= p[0]) { out += p[1]; n -= p[0]; } });
            return upper ? out.toUpperCase() : out;
        },

        /**
         * Flat typed styles, tried in order: [kind, regex (group 1 = the number), value(text), text(value)].
         * The number's position is where group 1 sits; everything around it is kept as typed.
         */
        _typedStyles: function () {
            var me = this;
            var WS = '(?=[ \\t\\u00a0\\u2002\\u2003])';
            var dec = function (t) { return +t; };
            var decText = function (v) { return String(v); };
            var letter = function (t) { return t.toLowerCase().charCodeAt(0) - 96; };
            var letterLow = function (v) { return v >= 1 && v <= 26 ? String.fromCharCode(96 + v) : null; };
            var letterUp = function (v) { return v >= 1 && v <= 26 ? String.fromCharCode(64 + v) : null; };
            var roman = function (t) { return me._romanValue(t); };
            var romanLow = function (v) { return me._romanText(v, false); };
            var romanUp = function (v) { return me._romanText(v, true); };
            return [
                // "§ 3", "Art. 3", "Artikel 3", "Ziffer 3", "Section 3", "Clause 3", ... (optional trailing dot)
                ['w', new RegExp('^(?:§§?|Art\\.|Artikel|Ziffer|Ziff\\.|Nr\\.|Abschnitt|Teil|Section|Clause|Article|Part)[ \\u00a0]*(\\d{1,3})\\.?' + WS), dec, decText],
                ['p', new RegExp('^\\((\\d{1,3})\\)' + WS), dec, decText],
                ['r', new RegExp('^(\\d{1,3})\\)' + WS), dec, decText],
                ['pri', new RegExp('^\\(([ivx]{1,6})\\)' + WS), roman, romanLow],
                ['pa', new RegExp('^\\(([a-z])\\)' + WS), letter, letterLow],
                ['ra', new RegExp('^([a-z])\\)' + WS), letter, letterLow],
                ['ri', new RegExp('^([IVX]{1,6})\\.' + WS), roman, romanUp],
                ['ua', new RegExp('^([A-Z])[.)]' + WS), letter, letterUp],
            ];
        },

        _dottedRe: /^(\d{1,3}(?:\.\d{1,3}){0,3})(\.?)(?=[ \t   ])/,

        /** Leading characters of a paragraph with where each sits: [{ run, pos, ch }]. Stops at a non-run. */
        _leadingChars: function (p, max) {
            // sdkjs element type id of a run (word/Editor/Paragraph/RunContent/Types.js); module-scoped there.
            var RUN = 0x0027;
            var out = [];
            var content = p.Content || [];
            for (var i = 0; i < content.length && out.length < max; i++) {
                var run = content[i];
                // Only plain runs: a number inside a field, link or nested control is left alone.
                if (!run || run.Type !== RUN) return out;
                for (var j = 0; j < run.Content.length && out.length < max; j++) {
                    var item = run.Content[j];
                    var ch = null;
                    if (item.IsText && item.IsText()) ch = String.fromCharCode(item.Value);
                    else if (item.IsSpace && item.IsSpace()) ch = ' ';
                    else if (item.IsTab && item.IsTab()) ch = '\t';
                    if (ch === null) return out;
                    out.push({ run: run, pos: j, ch: ch });
                }
            }
            return out;
        },

        /** Every body paragraph without list numbering, in order: { para, lead, chars }. */
        _plainParagraphs: function () {
            var me = this;
            var doc = me._doc();
            var paras = [];
            (doc.Content || []).forEach(function (el) { if (el && typeof el.GetAllParagraphs === 'function') el.GetAllParagraphs({ All: true }, paras); });
            return paras.filter(function (p) { return p && p.GetIndex() !== -1 && !(p.GetNumPr && p.GetNumPr()); }).map(function (p) {
                var chars = me._leadingChars(p, 24);
                return { para: p, chars: chars, lead: chars.map(function (c) { return c.ch; }).join('') };
            });
        },

        /**
         * Typed-number paragraphs in order. Dotted: { kind:'d<depth>', comps, at, len }. Flat:
         * { kind, value, at, len, text(v) }. `at`/`len` locate the number in the leading chars.
         */
        _typedParagraphs: function () {
            var me = this;
            var styles = me._typedStyles();
            var out = [];
            me._plainParagraphs().forEach(function (x) {
                var m = me._dottedRe.exec(x.lead);
                if (m) {
                    var comps = m[1].split('.').map(Number);
                    out.push({ para: x.para, chars: x.chars, kind: 'd' + comps.length, comps: comps, at: 0, len: m[1].length });
                    return;
                }
                for (var i = 0; i < styles.length; i++) {
                    var st = styles[i];
                    var f = st[1].exec(x.lead);
                    if (!f) continue;
                    var value = st[2](f[1]);
                    if (!value) continue;
                    out.push({ para: x.para, chars: x.chars, kind: st[0], value: value, at: f[0].indexOf(f[1]), len: f[1].length, text: st[3] });
                    return;
                }
            });
            return out;
        },

        _typedPlan: function () {
            var items = this._typedParagraphs();
            var rules = {};
            var counts = _.countBy(items, 'kind');
            // Dotted kinds: each number is (parent's current numbers) + (next at its depth).
            var ok = {};
            var first = {};
            var cur = [];
            items.forEach(function (x) {
                if (x.kind[0] !== 'd') return;
                var level = x.comps.length;
                if (!(x.kind in ok)) { ok[x.kind] = true; first[x.kind] = x.comps[level - 1]; }
                var parentOk = level === 1 || (cur.length >= level - 1 && x.comps.slice(0, level - 1).every(function (v, i) { return v === cur[i]; }));
                var sameParent = cur.length >= level && x.comps.slice(0, level - 1).every(function (v, i) { return v === cur[i]; });
                var expected = sameParent ? cur[level - 1] + 1 : first[x.kind];
                if (!parentOk || x.comps[level - 1] !== expected) ok[x.kind] = false;
                cur = x.comps.slice();
            });
            Object.keys(ok).forEach(function (k) { if (ok[k] && counts[k] > 1) rules[k] = { first: first[k] }; });
            // A deeper level is only safe when its parent level is too.
            Object.keys(rules).forEach(function (k) {
                for (var l = 1; l < +k.slice(1); l++) if (!rules['d' + l] && counts['d' + l]) delete rules[k];
            });
            // Flat kinds: first, first+1, ... counting through, or restarting after the nearest kind.
            var kinds = _.uniq(items.map(function (x) { return x.kind; }));
            kinds.filter(function (k) { return k[0] !== 'd' && counts[k] > 1; }).forEach(function (k) {
                var kFirst = _.find(items, function (x) { return x.kind === k; }).value;
                var follows = function (scope) {
                    var prev = null;
                    var restart = true;
                    for (var i = 0; i < items.length; i++) {
                        var x = items[i];
                        if (scope && x.kind === scope) { restart = true; continue; }
                        if (x.kind !== k) continue;
                        if (x.value !== (restart ? kFirst : prev + 1)) return false;
                        prev = x.value;
                        restart = false;
                    }
                    return true;
                };
                if (follows(null)) { rules[k] = { first: kFirst, scope: null }; return; }
                var scopes = kinds.filter(function (o) { return o !== k && follows(o); });
                scopes.sort(function (a, b) { return counts[b] - counts[a]; });
                if (scopes.length) rules[k] = { first: kFirst, scope: scopes[0] };
            });
            return rules;
        },

        /** Rewrite typed numbers that no longer match their position. Returns the paragraphs changed. */
        _healTyped: function (rules) {
            if (!rules || !Object.keys(rules).length) return 0;
            var me = this;
            var items = me._typedParagraphs();
            var changed = 0;
            var cur = [];
            var flat = {};
            items.forEach(function (x) {
                // Any paragraph restarts the flat kinds scoped to its kind.
                Object.keys(rules).forEach(function (k) { if (rules[k].scope === x.kind) flat[k] = 0; });
                var want;
                var have;
                if (x.kind[0] === 'd') {
                    var level = x.comps.length;
                    if (!rules[x.kind]) { cur = x.comps.slice(); return; }
                    var next = cur.length >= level ? cur[level - 1] + 1 : rules[x.kind].first;
                    var comps = cur.slice(0, level - 1);
                    while (comps.length < level - 1) comps.push(1);
                    comps.push(next);
                    cur = comps;
                    want = comps.join('.');
                    have = x.comps.join('.');
                } else {
                    var rule = rules[x.kind];
                    if (!rule) return;
                    flat[x.kind] = (flat[x.kind] || 0) + 1;
                    want = x.text(rule.first + flat[x.kind] - 1);
                    have = x.text(x.value);
                    if (!want) return;
                }
                if (want === have) return;
                me._replaceLeading(x.chars, x.at, x.len, want);
                changed += 1;
            });
            return changed;
        },

        /** Replace chars[at .. at+len) with `text`, inside the run the first of them sits in. */
        _replaceLeading: function (chars, at, len, text) {
            var first = chars[at];
            for (var i = at + len - 1; i >= at; i--) chars[i].run.RemoveFromContent(chars[i].pos, 1, true);
            first.run.AddText(text, first.pos);
        },

        /**
         * For the template editor: which numbering will follow a removed section. `lists`/`typed`:
         * kinds that heal; `unrecognized`: paragraphs that look numbered but follow no rule, with a
         * short preview each (they keep their numbers when something before them goes).
         */
        _numberingCheck: function () {
            var plan = this._numberingPlan();
            var typedItems = this._typedParagraphs();
            var unrecognized = [];
            typedItems.forEach(function (x) { if (!plan.typed[x.kind] && unrecognized.length < 10) unrecognized.push(x.para.GetText().slice(0, 60)); });
            var listKinds = _.uniq(this._numberedParagraphs().map(function (x) { return x.kind; }));
            listKinds.forEach(function (k) { if (!plan[k] && unrecognized.length < 10) unrecognized.push('(list) ' + k); });
            // Number-like starts that matched no known style at all.
            var known = {};
            typedItems.forEach(function (x) { known[x.para.GetId()] = true; });
            this._plainParagraphs().forEach(function (x) {
                if (known[x.para.GetId()] || unrecognized.length >= 10) return;
                if (/^(?:\d{1,3}[.)\-:]|[§(]\s*\d)/.test(x.lead)) unrecognized.push(x.para.GetText().slice(0, 60));
            });
            return {
                ok: true,
                lists: Object.keys(plan).length,
                typed: Object.keys(plan.typed || {}).length,
                unrecognized: unrecognized,
            };
        },
    }, {}));
});
