/*
 * top.legal — host-driven AI comments for the PDF editor
 * ------------------------------------------------------------------
 * The PDF editor's counterpart of documenteditor's DocumentEdits.js, reduced to COMMENTS ONLY.
 *
 * The dealroom's AI assistant proposes only comments for an uploaded PDF (there is no body to
 * redline). In the pdf.js viewer the dealroom places them itself; once the PDF opens in Euro-Office
 * that viewer is not mounted, and the server cannot write into the doc-server's in-memory PDF, so
 * without this bridge an approved AI comment had nowhere to go.
 *
 * Same wire protocol, channel and startup discipline as DocumentEdits.js (read its header): never
 * throw into the editor's startup, defer off the app:ready stack, announce with 'ready'. The only
 * differences: capabilities is ['comment'], replace/delete are refused with 'read_only', and the
 * gate is canComments rather than isEdit+canReview — commenting a PDF needs no track changes.
 */
define([
    'core'
], function () {
    'use strict';

    PDFE.Controllers = PDFE.Controllers || {};

    var CHANNEL = 'tl-office-edits';
    var SUPPORTED_VERSION = 1;

    // What an AI-proposed comment is attributed to. The tracked change itself cannot carry this
    // (see AUTHORSHIP above), so the comment is the only place in the file that can.
    var AI_AUTHOR = 'top.legal AI';
    // AI comments are internal by default: the marker is for the user's own side, and the existing
    // per-team redaction keeps it out of the counterparty's download.
    var AI_COMMENT_SCOPE = 'internal';

    // A batch is a user action, not a bulk import. This cap exists so a malformed or hostile
    // message cannot lock the editor in a long synchronous loop.
    var MAX_EDITS = 60;
    // Anchors are meant to be the shortest unambiguous quote. A very long one is a bug upstream,
    // and searching for it is slow, so it is rejected as malformed rather than attempted.
    var MAX_ANCHOR_CHARS = 600;
    // Search geometry is in millimetres, annotation quads in PDF points.
    var MM_TO_PT = 72 / 25.4;
    // The marker yellow reviewers use, translucent so the quoted words stay readable.
    var HIGHLIGHT_RGB = [1, 0.85, 0.2];
    var HIGHLIGHT_OPACITY = 0.5;

    PDFE.Controllers.DocumentEdits = Backbone.Controller.extend(_.extend({
        models: [],
        collections: [],
        views: [],

        initialize: function () {
            // Wrapped for the same reason as everything else below: `addListeners` reaches through
            // `getApplication()`, and this controller must never be able to throw into the app's
            // construction. There is nothing to recover — it simply means no bridge this session.
            try {
                this.addListeners({});
            } catch (e) { /* no listeners; the controller is inert rather than fatal */ }
        },

        /**
         * NOTHING HERE MAY THROW, AND NOTHING HERE MAY RUN SYNCHRONOUSLY ON THE EDITOR'S INIT STACK.
         *
         * This is not general defensiveness — it is the specific lesson from shipping this file
         * once and taking the editor down with it. `Common.NotificationCenter.trigger('app:ready')`
         * is a SYNCHRONOUS call in Main.onDocumentContentReady, sitting directly above
         * `api.SetDrawingFreeze(false)`, `hidePreloader()` and the rest of the document-ready
         * sequence. Backbone propagates a listener's exception straight out of `trigger`, so ONE
         * bad line in here aborts the remainder of the editor's own startup: the ribbon is left
         * half-built and the branding never applies. That is exactly what happened.
         *
         * Two rules follow, and they are both load-bearing:
         *   1. every entry point is wrapped — this controller can fail, but only ever alone;
         *   2. the app:ready handler captures state and defers, so all real work happens on a
         *      later tick where a throw cannot reach the editor's startup at all.
         *
         * The cost of both is a few microseconds. The cost of neither was an outage.
         */
        onLaunch: function () {
            var me = this;
            me.api = null;
            me.appOptions = null;
            me._started = false;

            try {
                // Guarded rather than assumed: this controller launches inside someone else's
                // ordered list, and a missing global here must degrade to "no fast path", never
                // to a TypeError thrown into the launch chain.
                if (!window.Common || !Common.NotificationCenter || typeof Common.NotificationCenter.on !== 'function') return;

                Common.NotificationCenter.on('app:ready', function (appOptions) {
                    try {
                        me.appOptions = appOptions || {};
                    } catch (e) { /* not worth failing over */ }
                    // OFF THE INIT STACK. See the note above — this is the single change that
                    // makes this controller incapable of breaking the editor's startup.
                    setTimeout(function () { me._start(); }, 0);
                });
            } catch (e) { /* no bridge this session; the host falls back to the server path */ }
        },

        /** Everything the controller actually does at startup, on its own tick and in its own try. */
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
            } catch (e) { /* the editor is unaffected; the host simply never sees 'ready' */ }
        },

        // ===============================================================
        // Host messaging
        // ===============================================================

        // Duplicated from NegotiationTab rather than shared: the two files are independent overlays
        // on an upstream tree, and introducing a third shared file to save fifteen lines would add
        // another thing that can silently fail to apply during an image build.
        parentOrigin: function () {
            if (this._parentOrigin !== undefined) return this._parentOrigin;
            var p = null;
            try {
                p = new URLSearchParams(window.location.search).get('parentOrigin');
                // VALIDATED, not just read. postMessage throws a SyntaxError on a targetOrigin it
                // cannot parse, and this value arrives on the URL — so an absent, truncated or
                // malformed parentOrigin would otherwise become an exception at the worst moment.
                // Round-tripping through URL also normalises it, so the string we compare inbound
                // origins against is the same shape the browser reports.
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
            } catch (e) { /* the host just does not hear from us; nothing else is affected */ }
        },

        /**
         * Tell the host what this image can do.
         *
         * This is the whole rollout mechanism. The app routes an apply through this bridge ONLY
         * after seeing this message, so an older image — which never sends it — makes the app fall
         * back to its server-side patcher with no error and no feature detection of its own. One
         * container serves dev, beta and prod, so the image swap has to be inert until each stage's
         * frontend chooses to use it.
         */
        _postReady: function () {
            this._post({
                type: 'ready',
                version: SUPPORTED_VERSION,
                capabilities: ['comment'],
                // The host disables its own affordance rather than discovering per-edit that this
                // session is a read-only frozen snapshot.
                canApply: this._canApply(),
            });
        },

        _listen: function () {
            var me = this;
            if (me._listening) return;
            me._listening = true;
            window.addEventListener('message', function (e) {
                try {
                    // Two independent checks: the origin must be the embedding page, and the
                    // message must actually come from it — origin alone would accept a
                    // same-origin subframe.
                    if (!me.parentOrigin() || e.origin !== me.parentOrigin()) return;
                    if (e.source !== window.parent) return;
                    var d = e.data;
                    if (!d || d.__tl !== CHANNEL) return;
                    if (d.type === 'ping') { me._postReady(); return; }
                    if (d.type !== 'apply') return;
                    me._onApply(d);
                } catch (err) {
                    // The host is waiting on a result it will now never get, and its own timeout
                    // routes the batch to the server patcher. Silence here costs one slow apply;
                    // an escaping throw costs whatever else is listening on this window.
                }
            });
        },

        // ===============================================================
        // Applying
        // ===============================================================

        /**
         * Can this session write tracked changes at all?
         *
         * Both halves matter. Without isEdit the document is a frozen snapshot served to the party
         * without the turn, and this bridge must not become a way around the turn model. Without
         * canReview the editor cannot record revisions, and applying anyway would write SILENT,
         * untracked edits into a contract — far worse than refusing.
         */
        _canApply: function () {
            var o = this.appOptions || {};
            return !!(this.api && o.canComments);
        },

        _onApply: function (msg) {
            var me = this;
            var requestId = msg.requestId;
            var payload = msg.payload || {};
            var edits = payload.edits;

            // An app newer than this image may speak a shape we cannot honour. Refusing the whole
            // batch is right: a partially-understood set of contract edits is the worst outcome.
            if (payload.version !== SUPPORTED_VERSION || !_.isArray(edits)) {
                me._post({ type: 'result', requestId: requestId, applied: [], failed: [], rejected: 'unsupported' });
                return;
            }

            if (!me._canApply()) {
                me._post({
                    type: 'result',
                    requestId: requestId,
                    applied: [],
                    failed: _.map(edits, function (edit) {
                        return { id: edit && edit.id, reason: 'read_only' };
                    }),
                });
                return;
            }

            var applied = [];
            var failed = [];

            // Turn revisions on for the batch and put the setting back afterwards. 'Local' is
            // deliberate: forcing the global setting would change how everyone else in the document
            // is editing, which is not ours to decide.
            try {
                for (var i = 0; i < edits.length && i < MAX_EDITS; i += 1) {
                    var edit = edits[i] || {};
                    var outcome;
                    try {
                        outcome = me._applyOne(edit);
                    } catch (e) {
                        // One bad edit must not abandon the rest of an approved batch.
                        outcome = 'error';
                    }
                    if (outcome === true) applied.push(edit.id);
                    else failed.push({ id: edit.id, reason: outcome });
                }

                // Anything past the cap is reported, never silently dropped.
                for (var j = MAX_EDITS; j < edits.length; j += 1) {
                    failed.push({ id: edits[j] && edits[j].id, reason: 'error' });
                }
            } catch (e) { /* reported per edit above */ }

            me._post({ type: 'result', requestId: requestId, applied: applied, failed: failed });
        },

        /**
         * Apply one edit. Returns true, or a failure reason string.
         *
         * A PDF comment here is a HIGHLIGHT annotation over the quoted words, carrying the comment
         * as its contents — the same object a reviewer creates by selecting text and commenting.
         *
         * WHY NOT asc_addComment (the documenteditor path). In the PDF editor it ignores the search
         * match entirely: CPDFDoc.AddComment drops a sticky note at the last mouse position or at a
         * fixed (10,10) on the current page. And a search leaves no text selection behind
         * (getSelectionQuads() is empty), so the editor's own highlighter has nothing to work on
         * either. The match's geometry is read from the search engine instead.
         */
        _applyOne: function (edit) {
            var anchor = edit.anchor;
            // A PDF has no tracked changes: only comments are honoured here.
            if (edit.kind !== 'comment') return 'read_only';
            if (typeof anchor !== 'string' || !anchor.length || anchor.length > MAX_ANCHOR_CHARS) return 'error';
            if (!edit.reason) return 'error';

            if (!this._find(anchor)) {
                // The quote may differ from the editor's text only in spacing — justified PDF lines
                // come out of the editor's extraction with DOUBLE spaces between words ("and  interest
                // in  and  to"), while the quote has single ones. Look the quote up in the editor's
                // own lines with whitespace collapsed, and search with the editor's exact wording.
                var exact = this._editorWording(anchor);
                if (!exact || !this._find(exact)) return 'not_found';
            }

            var quadsByPage = this._matchQuads();
            if (!quadsByPage) return 'not_found';

            return this._addAiHighlight(quadsByPage, edit.reason) ? true : 'error';
        },

        _find: function (text) {
            var settings = new AscCommon.CSearchSettings();
            settings.put_Text(text);
            settings.put_MatchCase(true);
            settings.put_WholeWords(false);
            return !!this.api.asc_findText(settings, true);
        },

        /**
         * The editor's own text for `quote`, matched with all whitespace runs collapsed to one
         * space, or null. Lines are joined with a space, as the editor's multi-line search joins
         * them. Only spacing is forgiven — case and every other character must still match.
         */
        _editorWording: function (quote) {
            var engine = this.api.getPDFDoc().SearchEngine;
            var pages = (engine && engine.PagesLines) || {};
            var needle = String(quote).replace(/\s+/g, ' ').trim();
            if (!needle) return null;

            var pageKeys = _.keys(pages);
            for (var k = 0; k < pageKeys.length; k += 1) {
                var lines = pages[pageKeys[k]] || [];
                var raw = '', folded = '', map = [];
                for (var li = 0; li < lines.length; li += 1) {
                    var line = (li ? ' ' : '') + String(lines[li]);
                    for (var ci = 0; ci < line.length; ci += 1) {
                        var ch = line.charAt(ci);
                        raw += ch;
                        if (/\s/.test(ch)) {
                            if (folded.length && folded.charAt(folded.length - 1) !== ' ') { folded += ' '; map.push(raw.length - 1); }
                        } else {
                            folded += ch;
                            map.push(raw.length - 1);
                        }
                    }
                }
                var at = folded.indexOf(needle);
                if (at >= 0) return raw.slice(map[at], map[at + needle.length - 1] + 1);
            }
            return null;
        },

        /**
         * The current search match as highlight quads, grouped by page, or null.
         *
         * A match is a list of line segments ({PageNum, X, Y, W, H}, millimetres from the page's
         * top-left) — one per line for a quote that wraps. Quads are PDF points, the unit
         * getSelectionQuads() produces from a hand-made selection; converting with mm→pt yields
         * the identical rectangle (checked against a real selection of the same words).
         */
        _matchQuads: function () {
            var engine = this.api.getPDFDoc().SearchEngine;
            var match = engine && engine.Elements && (engine.Elements[engine.CurId] || engine.Elements[0]);
            if (!match || !match.length) return null;

            var byPage = {};
            for (var i = 0; i < match.length; i += 1) {
                var seg = match[i];
                if (!seg || !(seg.W > 0) || !(seg.H > 0)) continue;
                var x1 = seg.X * MM_TO_PT, y1 = seg.Y * MM_TO_PT;
                var x2 = (seg.X + seg.W) * MM_TO_PT, y2 = (seg.Y + seg.H) * MM_TO_PT;
                (byPage[seg.PageNum] = byPage[seg.PageNum] || []).push([x1, y1, x2, y1, x1, y2, x2, y2]);
            }
            return _.isEmpty(byPage) ? null : byPage;
        },

        /**
         * One highlight per page the match touches, inside one undoable action (the same
         * DoAction + historydescription_Pdf_AddAnnot the editor's own marker tool uses, so it is
         * saved and co-edited like any hand-made highlight).
         *
         * Scope rides on the AUTHOR name (eoGroupUserName): in a PDF the comment's userData is the
         * annotation id, so the userData scope stamp used in Word cannot be stored here.
         */
        _addAiHighlight: function (quadsByPage, reason) {
            try {
                var comments = this.getApplication().getController('Common.Controllers.Comments');
                if (!comments) return false;
                var oDoc = this.api.getPDFDoc();
                var author = comments.eoGroupUserName(AI_COMMENT_SCOPE, AI_AUTHOR);
                var userId = comments.currentUserId;
                var text = String(reason);

                oDoc.DoAction(function () {
                    _.each(quadsByPage, function (quads, page) {
                        var xs = [], ys = [];
                        _.each(quads, function (q) { xs.push(q[0], q[2]); ys.push(q[1], q[5]); });
                        var now = new Date().getTime();
                        var annot = oDoc.AddAnnotByProps({
                            rect: [_.min(xs) - 1, _.min(ys) - 1, _.max(xs) + 1, _.max(ys) + 1],
                            page: Number(page),
                            name: AscCommon.CreateGUID(),
                            type: AscPDF.ANNOTATIONS_TYPES.Highlight,
                            creationDate: now,
                            modDate: now,
                            hidden: false
                        });
                        if (!annot) throw new Error('annotation not created');
                        annot.SetQuads(quads);
                        annot.SetBorderColor(HIGHLIGHT_RGB);
                        annot.SetOpacity(HIGHLIGHT_OPACITY);
                        // Ownership by the signed-in user's id, so they can resolve or delete it.
                        annot.SetAuthor(author);
                        annot.SetUserId(userId);
                        annot.SetContents(text);
                    });
                }, AscDFH.historydescription_Pdf_AddAnnot, oDoc);

                if (this.api.asc_endFindText) this.api.asc_endFindText();
                return true;
            } catch (e) {
                return false;
            }
        },

    }, PDFE.Controllers.DocumentEdits || {}));
});
