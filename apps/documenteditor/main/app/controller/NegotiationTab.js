/*
 * top.legal — host-driven ribbon tabs (native)
 * ------------------------------------------------------------------
 * One or more ribbon tabs that render whatever the host dealroom tells them to, and report
 * activations back. They contain NO business logic: no turns, no modes, no contracts, and nothing
 * that knows what any particular tab is for.
 *
 * The filename and the controller id are historical — this began as a single "Negotiation" tab.
 * They are kept because renaming them would touch app.js and Toolbar.js for no functional gain,
 * and every file this fork adds has to be re-applied as an overlay on each image build.
 *
 * Why: this file lives in the DocumentServer image, and changing that image costs a ~26 minute
 * CodeBuild plus a container swap that interrupts dev, beta and prod together (one container
 * serves all three). So the image gets a fixed widget vocabulary — chip, button, segmented,
 * select — and everything that actually changes (labels, ordering, which controls exist, when
 * they are disabled, translations) arrives at runtime as a descriptor from the app, shipped with
 * an ordinary frontend deploy. A rebuild is then only needed to add a widget TYPE.
 *
 * The app remains the single implementation of the business logic: this tab emits intents
 * ('finishTurn'), never outcomes, and the dealroom's existing mutations run unchanged. That is
 * what keeps this tab and the dealroom's own MUI strip from drifting — they are two renderings
 * of one state machine.
 *
 * Deliberately NOT gated on config.isEdit (unlike the Style tab). In turn-based negotiation the
 * party without the turn is served a READ-ONLY frozen snapshot, and that is exactly the person
 * who needs to see whose turn it is.
 *
 * Wire protocol (both directions carry __tl so unrelated traffic is ignored cheaply):
 *   app  -> tab : { __tl:'tl-office-ribbon', type:'descriptor',
 *                   payload:{ version, tabs:[{id,label,controls}], tab, controls, style?, css? } }
 *   tab  -> app : { __tl:'tl-office-ribbon', type:'ready' }
 *   tab  -> app : { __tl:'tl-office-ribbon', type:'action', id, value? }
 *
 * CONTEXT MENU. `contextMenu: [{id, label}]` (optional, additive — no version bump) declares an
 * entry for the right-click menu, shown under "Add comment". The first three entries are used. A click
 * is reported as an ordinary `action` with that id, so the host handles it like a ribbon button.
 *
 * ANCHORS (same channel, `anchors:*` types) — text ranges the host can attach things to, used for
 * tasks. See the "Anchors" section below for the messages and why they are built this way.
 *
 * TABS. `tabs` is the current shape; `tab`/`controls` are the original single-tab shape and are
 * still accepted, because an app deployed against this image may predate the change. The first
 * entry in `tabs` is the PRIMARY tab: Toolbar.js builds its panel during toolbar construction,
 * before any descriptor can have arrived. Every later entry is created lazily, the first time a
 * descriptor names it — Mixtbar.addTab is pure DOM plus a config splice and re-syncs its own
 * element caches, so it is safe to call after the toolbar is up.
 *
 * Tabs are never removed once created. A descriptor that stops naming a tab simply leaves it
 * empty: tearing tabs out of a live ribbon risks orphaning the active panel, and no host has
 * wanted it.
 */
define([
    'core'
], function () {
    'use strict';

    DE.Controllers = DE.Controllers || {};

    var CHANNEL = 'tl-office-ribbon';
    var SUPPORTED_VERSION = 1;
    // The tab Toolbar.js builds during toolbar construction. Must match tabs[0].id from the host.
    var PRIMARY_TAB_ID = 'negotiation';
    // Where the primary tab sits in Mixtbar's tab list. Toolbar.js inserts it with after=1, and
    // addTab splices at after+1, so it lands at index 2 and extra tabs follow at 2, 3, ...
    //
    // This mirrors an insertion index rather than reading one back, because Mixtbar keeps its tab
    // config in a closure. If the assumption is ever wrong (Style absent in a read-only session,
    // say), addTab walks back to the nearest real tab and inserts there: the extra tab lands in a
    // different POSITION, never in a broken state.
    var PRIMARY_TAB_AFTER = 1;

    DE.Controllers.NegotiationTab = Backbone.Controller.extend(_.extend({
        models: [],
        collections: [],
        views: [],

        descriptor: null,

        initialize: function () {
            this.addListeners({});
        },

        onLaunch: function () {
            this._panel = null;
        },

        setApi: function (api) {
            this.api = api;
            // Called from Toolbar.js mid-construction: an exception here would abort the toolbar
            // build, so anchors fail alone (see DocumentEdits.js for the incident that taught this).
            try { this._bindAnchorEvents(); } catch (e) { /* anchors unavailable, editor unaffected */ }
            return this;
        },

        setConfig: function (config) {
            this.toolbar = config.toolbar;
            // config.toolbar is the Toolbar CONTROLLER; its own `.toolbar` is the Mixtbar view
            // that owns addTab/setVisible. Reached defensively: a future upstream reshuffle here
            // must cost us extra tabs, not the primary one.
            this.toolbarView = (config.toolbar && config.toolbar.toolbar) || null;
            this.appConfig = config.mode;
            return this;
        },

        // ===============================================================
        // Host messaging
        // ===============================================================

        // DocsAPI puts the embedding page's origin on the editor URL, so we can pin the target
        // rather than posting to '*' — the descriptor carries contract state and this frame is
        // cross-origin from the dealroom.
        parentOrigin: function () {
            if (this._parentOrigin !== undefined) return this._parentOrigin;
            var p = null;
            try {
                p = new URLSearchParams(window.location.search).get('parentOrigin');
            } catch (e) { p = null; }
            this._parentOrigin = p || null;
            return this._parentOrigin;
        },

        contextMenuItem: function (index) {
            var list = this.descriptor && this.descriptor.contextMenu;
            var c = Array.isArray(list) ? list[index || 0] : null;
            if (!c || typeof c.id !== 'string' || typeof c.label !== 'string' || !c.label) return null;
            return { id: c.id, label: c.label };
        },

        contextMenuAction: function (id) {
            if (typeof id === 'string' && id) this._post({ type: 'action', id: id });
        },

        /*
         * SEARCH (`search:find`) — "show this passage": the host sends candidate strings, most
         * specific first (the full quote, then shorter fallbacks such as the date inside it). The
         * first one the document contains is found with the editor's own Find, which selects the
         * match and scrolls to it. Replies `search:found { requestId, ok, text? }`.
         */
        _onSearchFind: function (d) {
            var ok = false, hit;
            try {
                var candidates = _.isArray(d.candidates) ? d.candidates : [];
                if (this.api) {
                    for (var i = 0; i < candidates.length && !ok; i++) {
                        var text = candidates[i];
                        if (typeof text !== 'string' || !text.trim()) continue;
                        var settings = new AscCommon.CSearchSettings();
                        settings.put_Text(text.trim());
                        settings.put_MatchCase(false);
                        settings.put_WholeWords(false);
                        if (this.api.asc_findText(settings, true)) { ok = true; hit = text; }
                    }
                }
            } catch (e) { ok = false; }
            if (typeof d.requestId === 'string') this._post({ type: 'search:found', requestId: d.requestId, ok: ok, text: hit });
        },

        _post: function (msg) {
            var origin = this.parentOrigin();
            if (!origin || window.parent === window) return;
            msg.__tl = CHANNEL;
            window.parent.postMessage(msg, origin);
        },

        _listen: function () {
            var me = this;
            if (me._listening) return;
            me._listening = true;
            window.addEventListener('message', function (e) {
                // Two independent checks: the origin must be the embedding page, and the message
                // must actually come from it — origin alone would accept a same-origin subframe.
                if (!me.parentOrigin() || e.origin !== me.parentOrigin()) return;
                if (e.source !== window.parent) return;
                var d = e.data;
                if (!d || d.__tl !== CHANNEL) return;
                if (typeof d.type === 'string' && d.type.indexOf('anchors:') === 0) {
                    me._onAnchorMessage(d);
                    return;
                }
                if (d.type === 'search:find') {
                    me._onSearchFind(d);
                    return;
                }
                if (d.type !== 'descriptor') return;
                var payload = d.payload;
                // An app newer than this image may speak a shape we cannot draw. Declining is
                // better than rendering it half-right: the dealroom strip is still there.
                if (!payload || payload.version !== SUPPORTED_VERSION) return;
                me.descriptor = payload;
                me._render();
            });
        },

        // ===============================================================
        // Panel (ribbon content)
        // ===============================================================
        /**
         * Build the PRIMARY tab's panel. Called once by Toolbar.js while the toolbar is being
         * constructed, long before any descriptor exists — which is exactly why the primary tab
         * cannot be descriptor-declared like the others.
         */
        createToolbarPanel: function () {
            var me = this;
            me._injectStyles();
            me._panels = {};
            me._panel = me._makePanel(PRIMARY_TAB_ID);
            me._render();       // empty state until the host sends a descriptor
            me._listen();
            // The host may have been ready long before this tab was built (and rebuilds it on
            // every remount), so ask rather than wait to be told.
            me._post({ type: 'ready' });
            return me._panel;
        },

        /**
         * The panel markup for one tab. `id` is constrained by the host to [a-z][a-z0-9]* — it is
         * interpolated into markup and, in the app's own CSS, into an attribute selector.
         */
        _makePanel: function (id) {
            var $panel = $(
                '<section class="panel" data-tab="' + id + '" role="tabpanel" aria-labelledby="' + id + '">' +
                    '<div class="group eo-neg-group">' +
                        '<div class="eo-neg-controls"></div>' +
                    '</div>' +
                '</section>'
            );
            this._panels[id] = { $panel: $panel, $controls: $panel.find('.eo-neg-controls') };
            if (id === PRIMARY_TAB_ID) this.$controls = this._panels[id].$controls;
            return $panel;
        },

        /**
         * Create a ribbon tab that did not exist at toolbar-build time.
         *
         * Wrapped in try/catch on purpose: Mixtbar is upstream code we overlay rather than own, and
         * an extra tab failing to appear must never take the editor — or the primary tab — with it.
         */
        _ensureTab: function (id, label, position) {
            var me = this;
            if (me._panels[id]) return true;
            if (!me.toolbarView || typeof me.toolbarView.addTab !== 'function') return false;

            try {
                var $panel = me._makePanel(id);
                me.toolbarView.addTab(
                    { action: id, caption: label || id, layoutname: 'toolbar-' + id, dataHintTitle: (label || id).charAt(0) },
                    $panel,
                    position
                );
                if (typeof me.toolbarView.setVisible === 'function') me.toolbarView.setVisible(id, true);
                return true;
            } catch (e) {
                delete me._panels[id];
                return false;
            }
        },

        getButtons: function () { return []; },

        _injectStyles: function () {
            if (document.getElementById('eo-neg-tab-css')) return;
            // Intentionally small. The buttons themselves are upstream's `btn-toolbar x-huge`, so
            // their sizing, hover, active and disabled states — and every UI theme — come from
            // buttons.less. Restyling them here is what made the tab look foreign in the first
            // place, so this sheet now only covers what upstream has no rule for: our SVG glyph,
            // the status chip, and group layout.
            var css =
                // flex:1 on BOTH is what makes margin-left:auto work on the status chip. Without it
                // these shrink-wrap their content, there is no free space in the row, and the chip
                // just sits next to the last button instead of at the edge.
                '.eo-neg-group{display:flex;align-items:stretch;height:100%;flex:1 1 auto;overflow:hidden;}' +
                '.eo-neg-controls{display:flex;align-items:stretch;overflow-x:auto;overflow-y:hidden;width:100%;flex:1 1 auto;height:100%;box-sizing:border-box;}' +
                '.eo-neg-seg-group{display:flex;align-items:stretch;}' +
                // Matches the sprite glyphs upstream draws in .inner-box-icon on an x-huge button.
                // 20px inside upstream's 28px .inner-box-icon box, exactly like the sprite glyphs
                // (buttons.less: div.inner-box-icon{height:28px} with a 20px icon). At 28px the
                // glyph filled the whole box and pushed the caption out of the panel, which is
                // why the labels were clipped.
                '.eo-neg-ico{width:var(--eo-neg-icon-size,20px);height:var(--eo-neg-icon-size,20px);display:block;margin:0 auto;}' +
                // The chip is a status readout, not a control, so it is centred against the tall
                // button row rather than stretched to it.
                // margin-left:auto pushes the status to the far right of the ribbon row; the app also
                // emits it last so nothing sits between it and the edge.
                '.eo-neg-chip{align-self:center;margin-left:auto;flex:0 0 auto;display:inline-flex;align-items:center;' +
                    'font-size:11px;font-weight:bold;border-radius:999px;padding:3px 12px;margin-right:12px;white-space:nowrap;}' +
                '.eo-neg-chip.green{background:#3DBD7D;color:#fff;}' +
                '.eo-neg-chip.grey{background:#eceff1;color:#4a5568;}' +
                '.eo-neg-chip.amber{background:#ffaa00;color:#fff;}' +
                '.eo-neg-sel{align-self:center;flex:0 0 auto;display:inline-flex;align-items:center;gap:6px;font-size:11px;color:#909090;margin:0 8px;}' +
                '.eo-neg-sel select{font-size:11px;padding:4px 6px;border:1px solid #cfcfcf;border-radius:3px;background:#fff;color:#363636;max-width:200px;}' +
                '.eo-neg-empty{align-self:center;font-size:11px;color:#909090;padding:0 8px;}';
            var st = document.createElement('style');
            st.id = 'eo-neg-tab-css';
            st.type = 'text/css';
            st.innerHTML = css;
            document.getElementsByTagName('head')[0].appendChild(st);
        },

        // ===============================================================
        // Rendering — one branch per widget type. Adding a type here is the
        // only change that requires rebuilding the image.
        // ===============================================================
        /**
         * Normalise whichever descriptor shape arrived into one list of tabs.
         *
         * `tabs` is the current shape. `tab`/`controls` is the original single-tab one, still
         * accepted because the app and this image are deployed independently — one container
         * serves dev, beta and prod, so there is always a window where an older app is talking to
         * a newer image.
         */
        _tabsOf: function (d) {
            if (!d) return [];
            if (Array.isArray(d.tabs) && d.tabs.length) {
                return d.tabs.filter(function (tabSpec) {
                    // The id reaches markup and an attribute selector; anything else is dropped
                    // rather than sanitised, so a malformed descriptor cannot inject either.
                    return tabSpec && typeof tabSpec.id === 'string' && /^[a-z][a-z0-9]*$/i.test(tabSpec.id);
                });
            }
            return [{ id: PRIMARY_TAB_ID, label: null, controls: d.controls || [] }];
        },

        _render: function () {
            var me = this;
            if (!me._panels || !me.$controls) return;

            me._applyStyle();

            var tabs = me._tabsOf(me.descriptor);
            var extraIndex = 0;

            tabs.forEach(function (tabSpec) {
                if (tabSpec.id !== PRIMARY_TAB_ID) {
                    extraIndex += 1;
                    // Created once, then only re-rendered. A tab the host stops sending is left
                    // in place and simply empties — see the TABS note in the file header.
                    if (!me._ensureTab(tabSpec.id, tabSpec.label, PRIMARY_TAB_AFTER + extraIndex)) return;
                }
                var slot = me._panels[tabSpec.id];
                if (slot) me._renderControls(slot.$controls, tabSpec.controls || []);
                me._syncTabCaption(tabSpec.id, tabSpec.label);
            });

            // Tabs the descriptor no longer names keep their place but show the empty state, so a
            // host that drops a tab mid-session leaves nothing stale on screen.
            Object.keys(me._panels).forEach(function (id) {
                var stillNamed = tabs.some(function (tabSpec) { return tabSpec.id === id; });
                if (!stillNamed) me._renderControls(me._panels[id].$controls, []);
            });
        },

        /**
         * Keep a tab's caption in step with the descriptor. The primary tab is built by Toolbar.js
         * with the literal caption 'Negotiation' before any descriptor exists, and extra tabs keep
         * the label they were created with — so without this the primary tab was never translated.
         * textContent (via .text), never markup: labels are host-supplied.
         */
        _syncTabCaption: function (id, label) {
            if (typeof label !== 'string' || !label) return;
            var $a = $('.toolbar .tabs a[data-tab="' + id + '"]');
            if (!$a.length || $a.text() === label) return;
            $a.text(label).attr('data-title', label);
        },

        /** One branch per widget type. Adding a TYPE here is the only change that needs a rebuild. */
        _renderControls: function ($controls, controls) {
            var me = this;
            $controls.empty();

            if (!controls.length) {
                $controls.append($('<span class="eo-neg-empty"></span>').text('—'));
                return;
            }

            controls.forEach(function (c) {
                if (!c || !c.id || !c.type) return;
                var $el = null;
                if (c.type === 'chip')           $el = me._chip(c);
                else if (c.type === 'button')    $el = me._button(c);
                else if (c.type === 'segmented') $el = me._segmented(c);
                else if (c.type === 'select')    $el = me._select(c);
                if ($el) $controls.append($el);
            });
        },

        /**
         * Styling supplied by the HOST, so visual tweaks stop costing an image rebuild.
         *
         * Every size/spacing change so far has meant a ~25 min CodeBuild plus a container swap that
         * interrupts all three stages. Reading them from the descriptor instead makes them ordinary
         * frontend changes: instant locally, a pm2 reload to beta, and independent per stage.
         *
         * Two channels. `style` is a map of tokens applied as CSS custom properties on our own
         * container — the safe, expected path, and the base sheet already reads them. `css` is an
         * escape hatch for the unforeseen; it is injected into a single <style> element we own and
         * rewrite, never appended to, so repeated renders cannot pile up sheets. Selectors there
         * should stay scoped to .eo-neg-controls: nothing stops a stray rule reaching the rest of
         * the editor, and that would be a nasty thing to debug from a screenshot.
         */
        _applyStyle: function () {
            var d = this.descriptor || {};
            var panels = this._panels || {};
            if (d.style && typeof d.style === 'object') {
                // Applied to EVERY host panel, not just the primary one: the tokens describe the
                // shared widget vocabulary, so a button in one tab must not size differently
                // from the same button in another.
                Object.keys(panels).forEach(function (id) {
                    var el = panels[id].$controls[0];
                    if (!el) return;
                    Object.keys(d.style).forEach(function (k) {
                        // Token names are constrained so a descriptor cannot set arbitrary inline CSS.
                        if (!/^[a-z0-9-]+$/i.test(k)) return;
                        el.style.setProperty('--eo-neg-' + k, String(d.style[k]));
                    });
                });
            }
            var host = document.getElementById('eo-neg-host-css');
            if (!host) {
                host = document.createElement('style');
                host.id = 'eo-neg-host-css';
                host.type = 'text/css';
                document.getElementsByTagName('head')[0].appendChild(host);
            }
            host.innerHTML = (typeof d.css === 'string') ? d.css : '';
        },

        // textContent everywhere, never .html(): labels are host-supplied strings and some are
        // user-authored (workflow names), so they must never be parsed as markup.
        _chip: function (c) {
            var tone = (c.tone === 'green' || c.tone === 'amber') ? c.tone : 'grey';
            var $chip = $('<span class="eo-neg-chip ' + tone + '"></span>').text(c.label || '');
            if (c.hint) $chip.attr('title', c.hint);
            return $chip;
        },

        /**
         * A ribbon button in the editor's OWN idiom: icon above, caption below, no border, using
         * upstream's `btn-toolbar x-huge` markup (templateHugeCaption in component/Button.js, and
         * the `btn-slot text x-huge` slots in Toolbar.template).
         *
         * Reusing upstream classes rather than styling bespoke buttons means this tab inherits
         * hover, active, disabled and every UI theme automatically, instead of drifting the first
         * time someone switches to a dark theme. The only substitution is the glyph: upstream puts
         * a sprite <i class="icon btn-..."> inside .inner-box-icon, we put our SVG there, because
         * the sprite has no entry for these actions.
         */
        _ribbonButton: function (opts) {
            var $slot = $('<div class="btn-slot text x-huge"></div>');
            var $btn = $('<button type="button" class="btn btn-toolbar x-huge"></button>');
            var $icon = $('<div class="inner-box-icon"></div>');
            var ico = this._icon(opts.icon);
            if (ico) $icon.append(ico);
            $btn.append($icon);
            $btn.append($('<div class="inner-box-caption"></div>').append($('<span class="caption"></span>').text(opts.label || '')));
            // Native browser tooltip. Upstream's own buttons use data-hint + Common.UI.Tooltip,
            // but that needs a component instance per button; title costs nothing, needs no
            // wiring, and the text is host-supplied so it is already translated.
            if (opts.hint) $btn.attr('title', opts.hint);
            if (opts.active) $btn.addClass('active');
            if (opts.disabled) $btn.addClass('disabled').prop('disabled', true);
            $btn.on('click', function () {
                if (opts.disabled) return;
                opts.onClick();
            });
            $slot.append($btn);
            return $slot;
        },

        _button: function (c) {
            var me = this;
            return me._ribbonButton({
                label: c.label,
                icon: c.icon,
                hint: c.hint,
                disabled: !!(c.disabled || c.busy),
                onClick: function () { me._post({ type: 'action', id: c.id }); },
            });
        },

        /**
         * Icons arrive as GEOMETRY — { path, viewBox } — never as names: this file cannot import
         * the app's icon set, and names would mean the image carrying a catalogue only a rebuild
         * could extend.
         *
         * The viewBox must come from the descriptor rather than be assumed here. Material Symbols
         * exports use a 960-unit `0 -960 960 960` box while Material Icons use `0 0 24 24`; a
         * hardcoded box renders the other set as an empty square. Defaulting to 24 keeps older
         * descriptors that sent a bare path string working.
         *
         * currentColor makes the glyph follow the button's own text colour, so it tracks hover,
         * the active state and every UI theme without any rule of ours.
         */
        _icon: function (icon) {
            if (!icon) return null;
            var pathData = (typeof icon === 'string') ? icon : icon.path;
            var viewBox = (typeof icon === 'object' && icon.viewBox) ? icon.viewBox : '0 0 24 24';
            if (!pathData || typeof pathData !== 'string') return null;
            var svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
            svg.setAttribute('viewBox', viewBox);
            svg.setAttribute('class', 'eo-neg-ico');
            svg.setAttribute('aria-hidden', 'true');
            var path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
            // setAttribute, not innerHTML — host-supplied, must never be parsed as markup.
            path.setAttribute('d', pathData);
            path.setAttribute('fill', 'currentColor');
            svg.appendChild(path);
            return svg;
        },

        // A run of ribbon buttons with the current one `.active` — how the editor expresses a
        // mutually exclusive choice everywhere else (Display Mode, alignment). A bordered
        // segmented pill would have been the only control of its kind in the whole ribbon.
        _segmented: function (c) {
            var me = this;
            var $wrap = $('<div class="eo-neg-seg-group"></div>');
            (c.options || []).forEach(function (o) {
                $wrap.append(me._ribbonButton({
                    label: o.label,
                    icon: o.icon,
                    hint: o.hint,
                    active: o.value === c.value,
                    disabled: !!(c.disabled || o.disabled),
                    onClick: function () {
                        if (o.value === c.value) return;
                        me._post({ type: 'action', id: c.id, value: o.value });
                    },
                }));
            });
            $wrap.append($('<div class="separator long"></div>'));
            return $wrap;
        },

        _select: function (c) {
            var me = this;
            var $wrap = $('<span class="eo-neg-sel"></span>');
            if (c.label) $wrap.append($('<span></span>').text(c.label));
            var $sel = $('<select></select>');
            if (c.disabled) $sel.prop('disabled', true);
            (c.options || []).forEach(function (o) {
                var $o = $('<option></option>').attr('value', o.value).text(o.badge ? (o.label + ' • ' + o.badge) : o.label);
                if (o.value === c.value) $o.prop('selected', true);
                $sel.append($o);
            });
            $sel.on('change', function () {
                me._post({ type: 'action', id: c.id, value: $sel.val() });
            });
            $wrap.append($sel);
            return $wrap;
        },

        // ===============================================================
        // Anchors — host-owned text ranges (tasks)
        // ===============================================================
        /*
         * An anchor is a HIDDEN bookmark (the leading "_" keeps it out of this editor's and Word's
         * bookmark lists). It is document content: it syncs to co-editors, is saved in the docx,
         * follows its text through edits, and depends on no comment — so a task created from a
         * comment stays on its text when the comment is deleted.
         *
         * The HIGHLIGHT is not document content. It is drawn with the runs' CollaborativeMarks, the
         * layer behind "changes by other users", which is per client: no history point, nothing in
         * the change stream, nothing in the saved file. Measured 2026-09-26: a second session received
         * the bookmark and zero marks. The host sends `anchors:set` to internal users only, so the
         * counterparty never sees a highlight.
         *
         * The engine clears marks on runs another user edited (on save), and text typed inside a range
         * starts unmarked, so the highlight is re-applied after every change.
         *
         *   app -> tab : anchors:set      { anchors:[{name, color?}] }  highlight exactly these
         *   app -> tab : anchors:capture  { requestId }                  remember the selection
         *   tab -> app : anchors:captured { requestId, text, canAnchor }
         *   app -> tab : anchors:create   { name, requestId | commentGuid }
         *   tab -> app : anchors:created  { name, ok, text?, reason? }
         *   app -> tab : anchors:goto     { name }                       select + scroll to it
         *   app -> tab : anchors:remove   { name }
         *   tab -> app : anchors:clicked  { name }                       a click landed in one
         *
         * Names come from the host and end up in the document, so they are restricted to
         * `_tl` + [A-Za-z0-9_-]. Everything here is wrapped: a highlight must never take the editor
         * down with it.
         */
        _anchorNameOk: function (name) {
            return typeof name === 'string' && /^_tl[A-Za-z0-9_-]{1,80}$/.test(name);
        },

        _logicDoc: function () {
            var wc = this.api && this.api.WordControl;
            return (wc && wc.m_oLogicDocument) || null;
        },

        // Creating or removing a bookmark is an edit. A read-only session (turn-based, without the
        // turn) can still highlight and navigate, but cannot anchor.
        _canEditDoc: function () {
            var ld = this._logicDoc();
            try { return !!(ld && ld.CanEdit()); } catch (e) { return false; }
        },

        _bindAnchorEvents: function () {
            var me = this;
            if (me._anchorEventsBound || !me.api || typeof me.api.asc_registerCallback !== 'function') return;
            me._anchorEventsBound = true;
            me._anchors = [];
            me._anchorRuns = {};   // run id -> { name, rgb } as of the last apply
            me._captures = {};
            var reapply = function () { me._scheduleAnchorApply(); };
            ['asc_onDocumentContentReady', 'asc_onDocumentChanged', 'asc_onBookmarksUpdate',
                'asc_onCollaborativeChanges', 'asc_onApplyChanges', 'asc_onUndoRedoInCollaboration']
                .forEach(function (ev) { me.api.asc_registerCallback(ev, reapply); });
            // A pointer event rather than a selection event: only a click should open a task, never
            // the caret arriving in a range by keyboard. pointerup, not mouseup — the editor handles
            // pointer events on #id_viewer_overlay and the compatibility mouse events never arrive.
            document.addEventListener('pointerup', function (e) {
                if (!me._anchors.length) return;
                var host = document.getElementById('editor_sdk');
                if (!host || !host.contains(e.target)) return;
                setTimeout(function () { me._checkAnchorClick(); }, 0);
            }, true);
        },

        _onAnchorMessage: function (d) {
            var me = this;
            try {
                if (d.type === 'anchors:set') {
                    me._anchors = (Array.isArray(d.anchors) ? d.anchors : [])
                        .filter(function (a) { return a && me._anchorNameOk(a.name); })
                        .slice(0, 500);
                    me._applyAnchors();
                } else if (d.type === 'anchors:capture') {
                    me._captureSelection(d.requestId);
                } else if (d.type === 'anchors:create') {
                    me._createAnchor(d);
                } else if (d.type === 'anchors:goto') {
                    me._gotoAnchor(d.name);
                } else if (d.type === 'anchors:remove') {
                    me._removeAnchor(d.name);
                }
            } catch (e) { /* see the section header */ }
        },

        /**
         * Snapshot the selection when the host starts creating something, and anchor to THAT later.
         * Focus moves into the host's dialog in between, and the snapshot is what the user pointed
         * at when they clicked — not wherever the caret happens to be when the dialog is saved.
         */
        _captureSelection: function (requestId) {
            var me = this, ld = me._logicDoc();
            if (typeof requestId !== 'string' || !requestId) return;
            var text = '', state = null;
            if (ld && ld.IsSelectionUse() && !ld.IsSelectionEmpty()) {
                text = ld.GetSelectedText(false) || '';
                state = ld.GetSelectionState();
            }
            var keys = Object.keys(me._captures);
            if (keys.length > 20) delete me._captures[keys[0]];   // abandoned dialogs
            me._captures[requestId] = state;
            me._post({ type: 'anchors:captured', requestId: requestId, text: String(text).slice(0, 500), canAnchor: !!state && me._canEditDoc() });
        },

        _createAnchor: function (d) {
            var me = this, ld = me._logicDoc();
            var reply = function (ok, extra) { me._post(_.extend({ type: 'anchors:created', name: d.name, ok: ok }, extra || {})); };
            if (!me._anchorNameOk(d.name)) return;
            var captured = null;
            if (typeof d.requestId === 'string') {
                captured = me._captures[d.requestId] || null;
                delete me._captures[d.requestId];
            }
            if (!ld) return reply(false, { reason: 'notReady' });
            if (!me._canEditDoc()) return reply(false, { reason: 'readOnly' });

            if (captured) {
                ld.SetSelectionState(captured);
            } else if (typeof d.commentGuid === 'string' && d.commentGuid) {
                var cid = ld.Comments.GetCommentIdByGuid(d.commentGuid);
                var comment = cid ? ld.Comments.Get_ById(cid) : null;
                // A document-level comment has no text to anchor to.
                if (!comment || comment.IsGlobalComment() || false === comment.SelectCommentText()) return reply(false, { reason: 'noRange' });
            } else {
                return reply(false, { reason: 'noRange' });
            }
            if (!ld.IsSelectionUse() || ld.IsSelectionEmpty()) {
                ld.RemoveSelection();
                return reply(false, { reason: 'noRange' });
            }

            var text = ld.GetSelectedText(false) || '';
            // AddBookmark runs its own lock check and history action, so this is an ordinary edit that
            // co-editors receive. Its only failure mode is that lock, and it reports it by not adding.
            ld.AddBookmark(d.name);
            var chars = ld.GetBookmarksManager().GetBookmarkByName(d.name);
            // Collapse to the end of the new range: the selection we set was ours, not the user's.
            if (chars) chars[1].GoToBookmark(); else ld.RemoveSelection();
            ld.UpdateSelection();
            ld.UpdateInterface();
            if (!chars) return reply(false, { reason: 'locked' });
            reply(true, { text: String(text).slice(0, 500) });
            me._scheduleAnchorApply();
        },

        _gotoAnchor: function (name) {
            var ld = this._logicDoc();
            if (!ld || !this._anchorNameOk(name) || !ld.GetBookmarksManager().GetBookmarkByName(name)) return;
            this._suppressClickUntil = Date.now() + 600;   // selecting it is not the user clicking it
            ld.GoToBookmark(name, true);
            ld.UpdateSelection();
            ld.UpdateInterface();
        },

        _removeAnchor: function (name) {
            var ld = this._logicDoc();
            if (!ld || !this._anchorNameOk(name) || !this._canEditDoc()) return;
            if (!ld.GetBookmarksManager().GetBookmarkByName(name)) return;
            ld.RemoveBookmark(name);
        },

        _scheduleAnchorApply: function () {
            var me = this;
            if (me._anchorTimer) clearTimeout(me._anchorTimer);
            me._anchorTimer = setTimeout(function () {
                me._anchorTimer = null;
                try { me._applyAnchors(); } catch (e) { /* see the section header */ }
            }, 250);
        },

        /**
         * The runs a bookmark covers, in document order. Inserting a bookmark splits runs at its
         * edges, so whole runs are exactly the range. The end marker carries no usable name, so the
         * walk matches the two marker OBJECTS the manager returns.
         */
        _runsInBookmark: function (ld, name) {
            var chars = ld.GetBookmarksManager().GetBookmarkByName(name);
            if (!chars) return [];
            var startPara = chars[0].GetParagraph(), endPara = chars[1].GetParagraph();
            if (!startPara || !endPara) return [];
            var paras = [startPara];
            if (startPara !== endPara) {
                var all = ld.GetAllParagraphs({ OnlyMainDocument: true, All: true }) || [];
                var s = all.indexOf(startPara), e = all.indexOf(endPara);
                if (s < 0 || e < s) return [];   // header/footnote, or not in the main flow
                paras = all.slice(s, e + 1);
            }
            var out = [], inside = false;
            var walk = function (content) {
                for (var i = 0; i < content.length; i++) {
                    var el = content[i];
                    if (!el) continue;
                    if (el === chars[0]) { inside = true; continue; }
                    if (el === chars[1]) { inside = false; continue; }
                    if (el instanceof AscWord.ParaRun) { if (inside) out.push(el); continue; }
                    if (Array.isArray(el.Content)) walk(el.Content);   // hyperlinks, inline content controls
                }
            };
            paras.forEach(function (p) { walk(p.Content); });
            return out;
        },

        _rgb: function (color) {
            var m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(color || '');
            // Default: a soft amber that reads as "attention" on white and is distinct from every
            // co-editor colour the engine assigns.
            return m ? [parseInt(m[1], 16), parseInt(m[2], 16), parseInt(m[3], 16)] : [255, 224, 138];
        },

        // Remove only the ranges in our colour — the same run can carry a real co-editor's marks.
        _unmarkRun: function (run, rgb) {
            var ranges = run.CollaborativeMarks.Ranges;
            for (var i = ranges.length - 1; i >= 0; i--) {
                var c = ranges[i].Color;
                if (c && c.r === rgb[0] && c.g === rgb[1] && c.b === rgb[2]) ranges.splice(i, 1);
            }
        },

        _applyAnchors: function () {
            var me = this, ld = me._logicDoc();
            if (!ld || !window.AscWord || !AscWord.CDocumentColor || !AscWord.ParaRun) return;
            var next = {}, changed = false;
            (me._anchors || []).forEach(function (a) {
                var rgb = me._rgb(a.color);
                me._runsInBookmark(ld, a.name).forEach(function (run) {
                    next[run.Id] = { name: a.name, rgb: rgb };
                    var len = run.Content.length;
                    var covered = run.CollaborativeMarks.Ranges.some(function (r) {
                        return r.PosS <= 0 && r.PosE >= len && r.Color && r.Color.r === rgb[0] && r.Color.g === rgb[1] && r.Color.b === rgb[2];
                    });
                    if (covered) return;
                    // Typing inside a marked run leaves our colour in pieces around the new text;
                    // replace them with one full-width range.
                    me._unmarkRun(run, rgb);
                    run.CollaborativeMarks.Add(0, len, new AscWord.CDocumentColor(rgb[0], rgb[1], rgb[2]));
                    changed = true;
                });
            });
            // Runs we marked last time that are no longer ours: the task closed or the anchor went.
            Object.keys(me._anchorRuns || {}).forEach(function (runId) {
                if (next[runId]) return;
                var run = AscCommon.g_oTableId.Get_ById(runId);
                if (run && run.CollaborativeMarks) { me._unmarkRun(run, me._anchorRuns[runId].rgb); changed = true; }
            });
            me._anchorRuns = next;
            if (changed) {
                ld.DrawingDocument.ClearCachePages();
                ld.DrawingDocument.FirePaint();
            }
        },

        _checkAnchorClick: function () {
            var me = this, ld = me._logicDoc();
            if (!ld || Date.now() < (me._suppressClickUntil || 0)) return;
            try {
                // A drag-selection that happens to start in a range is not a click on it.
                if (ld.IsSelectionUse() && !ld.IsSelectionEmpty()) return;
                var p = ld.GetCurrentParagraph();
                if (!p || !p.GetClassByPos) return;
                var run = p.GetClassByPos(p.Get_ParaContentPos(false, false));
                var hit = run && me._anchorRuns[run.Id];
                if (hit) me._post({ type: 'anchors:clicked', name: hit.name });
            } catch (e) { /* see the section header */ }
        },
    }, DE.Controllers.NegotiationTab || {}));
});
