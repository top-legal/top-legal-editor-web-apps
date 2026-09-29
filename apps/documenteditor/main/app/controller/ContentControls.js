/*
 * top.legal — host-driven content controls (Word template fields)
 * ------------------------------------------------------------------
 * Lets the embedding page link text in the document to one of its input fields or conditions, by
 * wrapping it in a Word content control (w:sdt) whose TAG names the field. The document itself is
 * the record of what is linked: tags survive save, reload, download and a round trip through Word,
 * so the host keeps no separate mapping.
 *
 *   tl:f:<inputFieldID>              the text is replaced by the field's value when drafting
 *   tl:c:<inputFieldID>=<optionKey>  the wrapped block is kept only when that option is chosen
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
 *   app    -> editor : { __tl:'tl-office-fields', type:'fill', requestId, values:{ <fieldID>: <text> } }
 *   app    -> editor : { __tl:'tl-office-fields', type:'applyConditions', requestId, answers:{ <fieldID>: [<optionKey>] } }
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
    var TAG_RE = /^tl:(f:[A-Za-z0-9_-]{1,100}|c:[A-Za-z0-9_-]{1,100}=[A-Za-z0-9_ .-]{1,100})$/;
    var MAX_ALIAS = 120;
    // Text preview per control in a list — enough for the panel, never the whole clause.
    var MAX_TEXT = 200;
    // Shading so linked text is visible at rest, not only when the cursor is inside it.
    var HIGHLIGHT = [255, 236, 179];

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
                case 'fill': me._result(requestId, me._fill(d.values)); return;
                case 'applyConditions': me._result(requestId, me._applyConditions(d.answers)); return;
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
                if (tag.indexOf('tl:f:') !== 0 || typeof cc.IsInlineLevel !== 'function' || !cc.IsInlineLevel()) return;
                var v = values[tag.slice(5)];
                if (typeof v !== 'string' || !v) return;
                var current = '';
                try { current = String(cc.GetInnerText()); } catch (e) { current = ''; }
                if (current !== v) todo.push({ cc: cc, text: v.slice(0, 5000) });
            });
            if (!todo.length) return { ok: true, filled: 0 };
            try {
                var paragraphs = todo.map(function (x) { return x.cc.GetParagraph(); }).filter(Boolean);
                var locked = doc.Document_Is_SelectionLocked(AscCommon.changestype_None, {
                    Type: AscCommon.changestype_2_ElementsArray_and_Type,
                    Elements: paragraphs,
                    CheckType: AscCommon.changestype_Paragraph_Content,
                });
                if (locked) return { ok: false, reason: 'locked' };
                doc.StartAction(AscDFH.historydescription_Document_SetContentControlText);
                todo.forEach(function (x) {
                    if (x.cc.IsPlaceHolder && x.cc.IsPlaceHolder()) x.cc.ReplacePlaceHolderWithContent();
                    var run = x.cc.MakeSingleRunElement(true);
                    if (run) run.AddText(x.text);
                });
                doc.Recalculate();
                doc.UpdateInterface();
                doc.FinalizeAction();
                return { ok: true, filled: todo.length };
            } catch (e) {
                return { ok: false, reason: 'error' };
            }
        },

        /**
         * Finish-drafting step for conditions: a `tl:c:<fieldID>=<optionKey>` block whose option is
         * not among the answer's keys is deleted with its content; one whose option was chosen is
         * kept. Fields with no answer at all are left untouched, so an unanswered question never
         * silently deletes a clause.
         */
        _applyConditions: function (answers) {
            if (!this._canEdit()) return { ok: false, reason: 'readOnly' };
            if (!answers || typeof answers !== 'object') return { ok: false, reason: 'badAnswers' };
            var doc = this._doc();
            var drop = [];
            this._ownControls().forEach(function (cc) {
                var m = /^tl:c:([^=]+)=(.+)$/.exec(cc.GetTag());
                if (!m) return;
                var chosen = answers[m[1]];
                if (!_.isArray(chosen)) return;
                if (chosen.indexOf(m[2]) === -1) drop.push(cc);
            });
            if (!drop.length) return { ok: true, removed: 0 };
            try {
                var locked = doc.Document_Is_SelectionLocked(AscCommon.changestype_None, {
                    Type: AscCommon.changestype_2_ElementsArray_and_Type,
                    Elements: drop.map(function (cc) { return cc.IsBlockLevel() ? cc : cc.GetParagraph(); }).filter(Boolean),
                    CheckType: AscCommon.changestype_ContentControl_Remove,
                });
                if (locked) return { ok: false, reason: 'locked' };
                doc.StartAction(AscDFH.historydescription_Document_RemoveContentControl);
                drop.forEach(function (cc) { doc.RemoveContentControl(cc.GetId()); });
                doc.Recalculate();
                doc.UpdateInterface();
                doc.FinalizeAction();
                return { ok: true, removed: drop.length };
            } catch (e) {
                return { ok: false, reason: 'error' };
            }
        },
    }, {}));
});
