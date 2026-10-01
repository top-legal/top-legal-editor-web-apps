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
    var TAG_RE = new RegExp('^tl:(f:[A-Za-z0-9_.-]{1,160}|c:[A-Za-z0-9_.-]{1,160}=' + OPTION_KEY + '(\\|' + OPTION_KEY + '){0,19}|t:[A-Za-z0-9_.-]{1,160})$');
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
                case 'colors': me._result(requestId, me._setColors(d)); return;
                case 'fill': me._result(requestId, me._fill(d.values)); return;
                case 'applyConditions': me._result(requestId, me._applyConditions(d.answers, d.dropTags)); return;
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
            try { return String(this.api.asc_GetSelectedText(true) || '').slice(0, MAX_TEXT); } catch (e) { return ''; }
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
                if (!wrong) return changed;
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
            return changed;
        },
    }, {}));
});
