/*
 * top.legal — host-driven document edits (in-session redlines)
 * ------------------------------------------------------------------
 * Applies a batch of anchored text edits, as TRACKED CHANGES, on behalf of the embedding dealroom.
 * The dealroom's AI assistant proposes changes, the user approves them there, and the approved set
 * arrives here over postMessage.
 *
 * WHY THIS EXISTS. The dealroom used to apply approved changes by patching the docx in S3 and then
 * remounting the editor with a fresh document key — the doc-server holds the authoritative document
 * in memory, so a write underneath it is invisible until the document is reloaded. That reload is
 * not a cosmetic problem: it interrupts whoever is reading, loses their place, and briefly shows a
 * loading state in the middle of a negotiation. Applying the same edits THROUGH the live session
 * removes it, and is more correct besides — the edits sync to every participant, land in the
 * document's own history, are individually undoable, and S3 never diverges from what is on screen.
 *
 * NO UI. This controller renders nothing and owns no ribbon panel. That is why, unlike
 * NegotiationTab, it needs no overlay in Toolbar.js: it takes the api off the Main controller when
 * 'app:ready' fires, which keeps this customization to exactly two touch points — this file and the
 * app.js registration. Fewer overlay points, less to re-apply on every image build.
 *
 * IT MUST NEVER TAKE THE EDITOR WITH IT. The first version of this file did: `app:ready` is
 * triggered SYNCHRONOUSLY from Main.onDocumentContentReady, immediately above
 * `api.SetDrawingFreeze(false)` and `hidePreloader()`, and Backbone lets a listener's exception
 * escape `trigger` — so a single bad line here aborted the rest of the editor's startup, leaving a
 * half-built ribbon and unapplied branding. Every entry point below is therefore wrapped, the
 * app:ready handler defers its real work to a later tick, and the registration sits LAST in app.js's
 * controller list. Nothing depends on this controller having launched, so failing alone is always
 * the correct outcome — the dealroom simply keeps using the server patcher.
 *
 * WHAT THIS DELIBERATELY DOES NOT DECIDE. It has no idea what a contract, a playbook or an approval
 * is. It receives text anchors and replacements, reports precisely what it could and could not do,
 * and never guesses. The dealroom keeps its server-side patcher for everything reported as failed,
 * so a refusal here costs the user nothing.
 *
 * AUTHORSHIP. A tracked change made in-session is authored by the SIGNED-IN USER, and this cannot be
 * changed: the co-editing identity is fixed at document load from the signed config, and there is no
 * per-change author API. That is accepted — the user approved the change, so they own it. The AI
 * provenance is carried instead by an internal-scoped comment attributed to 'top.legal AI', which
 * the existing per-team redaction strips from anything the counterparty downloads.
 *
 * Wire protocol (both directions carry __tl, so unrelated traffic is ignored cheaply):
 *   app    -> editor : { __tl:'tl-office-edits', type:'apply', requestId, payload:{ version, edits:[
 *                         { id, kind:'replace'|'delete'|'comment', anchor, replacement?, reason? } ] } }
 *   editor -> app    : { __tl:'tl-office-edits', type:'ready', version, capabilities:[...] }
 *   editor -> app    : { __tl:'tl-office-edits', type:'result', requestId,
 *                         applied:[id], failed:[{ id, reason }] }
 *
 * THE HOST GUARANTEES ANCHOR UNIQUENESS. It must resolve every anchor against the current document
 * text, server-side, before sending the batch — the editor's search API cannot answer "how many
 * times does this occur" synchronously (the total arrives on an async callback), so this controller
 * cannot re-check it. See the note in _applyOne.
 *
 * Failure reasons are a closed set, because the app branches on them:
 *   not_found        the anchor is not in the document
 *   ambiguous        reported by the HOST's own pre-check, never raised here
 *   bad_replacement  the replacement contains something the editor refuses to insert
 *   read_only        this session cannot edit, or cannot track changes
 *   error            an unexpected throw, reported rather than swallowed
 */
define([
    'core'
], function () {
    'use strict';

    DE.Controllers = DE.Controllers || {};

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

    DE.Controllers.DocumentEdits = Backbone.Controller.extend(_.extend({
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
                capabilities: ['replace', 'delete', 'comment'],
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
            return !!(this.api && o.isEdit && o.canReview);
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
            var previousTrack = null;
            try {
                previousTrack = me.api.asc_GetLocalTrackRevisions();
            } catch (e) { previousTrack = null; }

            try {
                me.api.asc_SetLocalTrackRevisions(true);

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
            } finally {
                try {
                    me.api.asc_SetLocalTrackRevisions(previousTrack === null ? null : !!previousTrack);
                } catch (e) { /* restoring a UI preference must not mask the result */ }
            }

            me._post({ type: 'result', requestId: requestId, applied: applied, failed: failed });
        },

        /**
         * Apply one edit. Returns true, or a failure reason string.
         *
         * ORDER MATTERS: find, then comment, then replace.
         *
         * The find leaves the match SELECTED, which is what both the comment and the replace act
         * on. Commenting before the replace is deliberate — it anchors to a range we know exists at
         * that moment, and the editor carries the anchor across the subsequent edit. Doing it the
         * other way round would mean commenting on whatever the replace happened to leave selected,
         * which is not specified anywhere we control.
         */
        _applyOne: function (edit) {
            var kind = edit.kind;
            var anchor = edit.anchor;

            if (kind !== 'replace' && kind !== 'delete' && kind !== 'comment') return 'error';
            if (typeof anchor !== 'string' || !anchor.length || anchor.length > MAX_ANCHOR_CHARS) return 'error';

            var replacement = kind === 'replace' ? edit.replacement : '';
            if (kind === 'replace' && typeof replacement !== 'string') return 'error';

            // Ask the editor whether it will accept the replacement BEFORE touching the document.
            // Some strings are rejected outright, and finding that out after a successful search
            // would leave the selection moved for no reason.
            if (kind === 'replace' && replacement.length) {
                var err = null;
                try {
                    err = this.api.asc_GetErrorForReplaceString(replacement);
                } catch (e) { err = null; }
                if (err) return 'bad_replacement';
            }

            var settings = new AscCommon.CSearchSettings();
            settings.put_Text(anchor);
            // Case-sensitive, and NOT whole-words: an anchor is a verbatim quote from the document,
            // and it routinely starts or ends mid-word ("...ing the Term" ).
            settings.put_MatchCase(true);
            settings.put_WholeWords(false);

            // AMBIGUITY IS NOT CHECKED HERE, AND CANNOT BE.
            //
            // asc_findText selects the next match and reports only whether it found one. The match
            // TOTAL arrives separately and asynchronously, through the 'asc_onSetSearchCurrent'
            // callback (see Search.js: onUpdateSearchCurrent(current, all)) — so there is no
            // synchronous way to ask "how many times does this occur" at the moment of applying.
            //
            // Uniqueness is therefore the HOST's guarantee, established before the batch is sent.
            // That is the better place for it regardless: the dealroom resolves every anchor
            // server-side against the same accepted document text the assistant read, using a
            // tolerant matcher that already rejects ambiguous anchors. Re-deciding it here with a
            // weaker API would duplicate that logic badly.
            //
            // The residual risk is document DRIFT: an anchor unique when the host resolved it could
            // have gained a second occurrence by the time the user approves. That window is
            // seconds, and it is why 'ambiguous' stays in the reason vocabulary below — the host
            // reports its own pre-check failures with it, so the protocol does not have to change
            // if we later add a client-side guard on the async callback.
            if (!this.api.asc_findText(settings, true)) return 'not_found';

            /**
             * A COMMENT-ONLY EDIT MUST NOT REPORT SUCCESS WITHOUT WRITING ONE. The note IS the
             * change here, so there is nothing to be best-effort about: no text, or a comment the
             * editor refused, is a failure the host can retry through the server patcher.
             *
             * The observed bug was the opposite — a comment-only proposal arrived with no text and
             * was reported applied, so comments appeared only when a redline happened to be applied
             * alongside them.
             */
            if (kind === 'comment') {
                if (!edit.reason) return 'error';
                return this._addAiComment(edit.reason) ? true : 'error';
            }

            /**
             * For a REDLINE the comment stays best-effort, deliberately: the change is what the user
             * approved, and a failed annotation must not sink it. Added before the replace so it
             * anchors to the ORIGINAL wording, which survives as a strikethrough once the deletion
             * is tracked.
             */
            if (edit.reason) this._addAiComment(edit.reason);

            /**
             * RE-FIND BEFORE REPLACING. Not defensive — required.
             *
             * asc_replaceText acts on the CURRENT search result, and adding the comment above moves
             * the selection, which silently invalidates it. Without this second find the replace
             * does nothing at all AND STILL RETURNS TRUTHY: the observed failure was a document
             * carrying the assistant's comments with none of its redlines, reported to the user as
             * fully applied. Verified locally — original text still present, replacement absent,
             * asc_HaveRevisionsChanges false, and the batch reported as applied.
             *
             * Comment first and then re-find (rather than replacing first) so the comment anchors to
             * the ORIGINAL wording, which survives as a strikethrough once the deletion is tracked.
             * That matches what the server-side patcher does, so both paths read the same in Word.
             */
            if (!this.api.asc_findText(settings, true)) return 'not_found';

            /**
             * IGNORE WHAT asc_replaceText RETURNS. It is not a success flag — measured in the local
             * editor, it returned TRUE for a replace that provably did nothing (stale search state)
             * and FALSE for one that worked, leaving the replacement in the document with the
             * revision recorded. Branching on it gets the answer backwards either way.
             *
             * 'delete' is a replace with nothing: with revisions on the editor records a deletion
             * rather than removing the text.
             */
            this.api.asc_replaceText(settings, kind === 'delete' ? '' : replacement, false);

            /**
             * Success is confirmed from the DOCUMENT instead: the replacement has to be findable.
             * This is what turns "reported applied but nothing happened" — a document carrying the
             * assistant's comments and none of its redlines — into an honest 'not_found' that the
             * host retries through the server patcher.
             *
             * A delete cannot be verified this way: with revisions on the text survives as a tracked
             * strikethrough, so searching for it proves nothing either way. It is reported as
             * applied on the strength of the find above, which is the weaker guarantee of the two.
             */
            if (kind === 'replace') {
                var check = new AscCommon.CSearchSettings();
                check.put_Text(replacement);
                check.put_MatchCase(true);
                check.put_WholeWords(false);
                if (!this.api.asc_findText(check, true)) return 'not_found';
            }

            return true;
        },

        /**
         * Attach the AI's reasoning to the current selection, attributed to the assistant.
         *
         * Scope handling is NOT optional decoration. The fork mirrors a comment's team scope into
         * the author's username group so OnlyOffice's native commentGroups filtering can show each
         * team only what it should see, and tl-office redacts by the same scope on download. Setting
         * a plain author name here would overwrite that group and leak internal comments into the
         * counterparty's copy. eoGroupUserName takes the display name as its second argument
         * precisely so both can be expressed at once.
         *
         * Returns whether a comment was actually written. A redline's caller ignores it — a failed
         * annotation must not sink the change it annotates — but a comment-only edit depends on it,
         * because there the comment IS the change.
         */
        _addAiComment: function (reason) {
            try {
                var comments = this.getApplication().getController('Common.Controllers.Comments');
                if (!comments) return false;

                // Mirrors Comments.js's own buildCommentData(): the document editor needs the Word
                // subclass, and the base class is only a fallback. That helper is module-private
                // there, so the two-line check is repeated rather than reached for.
                var comment = (typeof Asc.asc_CCommentDataWord !== 'undefined')
                    ? new Asc.asc_CCommentDataWord(null)
                    : new Asc.asc_CCommentData(null);
                comment.asc_putText(String(reason));
                comment.asc_putTime(comments.utcDateToString(new Date()));
                comment.asc_putOnlyOfficeTime(comments.ooDateToString(new Date()));
                // Attribution by NAME, ownership by ID. The signed-in user's id is used on purpose:
                // it is what lets them resolve or delete the comment. An invented author id would
                // make the assistant's own annotations undeletable by the person reviewing them.
                comment.asc_putUserId(comments.currentUserId);
                comment.asc_putSolved(false);
                comment.asc_putUserData(comments.eoEncodeScope(comment.asc_getUserData(), AI_COMMENT_SCOPE));
                comment.asc_putUserName(comments.eoGroupUserName(AI_COMMENT_SCOPE, AI_AUTHOR));

                this.api.asc_addComment(comment);
                return true;
            } catch (e) {
                return false;
            }
        },

    }, DE.Controllers.DocumentEdits || {}));
});
