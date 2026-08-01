/*
 * top.legal — Negotiation tab (native)
 * ------------------------------------------------------------------
 * A "Negotiation" ribbon tab that renders whatever the host dealroom tells it to, and reports
 * activations back. It contains NO negotiation logic: no turns, no modes, no contracts.
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
 *   app  -> tab : { __tl:'tl-office-ribbon', type:'descriptor', payload:{version,tab,controls} }
 *   tab  -> app : { __tl:'tl-office-ribbon', type:'ready' }
 *   tab  -> app : { __tl:'tl-office-ribbon', type:'action', id, value? }
 */
define([
    'core'
], function () {
    'use strict';

    DE.Controllers = DE.Controllers || {};

    var CHANNEL = 'tl-office-ribbon';
    var SUPPORTED_VERSION = 1;

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
            return this;
        },

        setConfig: function (config) {
            this.toolbar = config.toolbar;
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
                if (!d || d.__tl !== CHANNEL || d.type !== 'descriptor') return;
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
        createToolbarPanel: function () {
            var me = this;
            me._injectStyles();
            me._panel = $(
                '<section class="panel" data-tab="negotiation" role="tabpanel" aria-labelledby="negotiation">' +
                    '<div class="group eo-neg-group">' +
                        '<div class="eo-neg-controls" id="eo-neg-controls"></div>' +
                    '</div>' +
                '</section>'
            );
            me.$controls = me._panel.find('#eo-neg-controls');
            me._render();       // empty state until the host sends a descriptor
            me._listen();
            // The host may have been ready long before this tab was built (and rebuilds it on
            // every remount), so ask rather than wait to be told.
            me._post({ type: 'ready' });
            return me._panel;
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
                '.eo-neg-group{display:flex;align-items:stretch;height:100%;overflow:hidden;}' +
                '.eo-neg-controls{display:flex;align-items:stretch;overflow-x:auto;overflow-y:hidden;max-width:100%;height:100%;box-sizing:border-box;}' +
                '.eo-neg-seg-group{display:flex;align-items:stretch;}' +
                // Matches the sprite glyphs upstream draws in .inner-box-icon on an x-huge button.
                '.eo-neg-ico{width:20px;height:20px;display:block;margin:0 auto;}' +
                // The chip is a status readout, not a control, so it is centred against the tall
                // button row rather than stretched to it.
                '.eo-neg-chip{align-self:center;flex:0 0 auto;display:inline-flex;align-items:center;' +
                    'font-size:11px;font-weight:bold;border-radius:999px;padding:3px 10px;margin:0 8px;white-space:nowrap;}' +
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
        _render: function () {
            var me = this;
            if (!me.$controls) return;
            me.$controls.empty();

            var controls = (me.descriptor && me.descriptor.controls) || [];
            if (!controls.length) {
                me.$controls.append($('<span class="eo-neg-empty"></span>').text('—'));
                return;
            }

            controls.forEach(function (c) {
                if (!c || !c.id || !c.type) return;
                var $el = null;
                if (c.type === 'chip')           $el = me._chip(c);
                else if (c.type === 'button')    $el = me._button(c);
                else if (c.type === 'segmented') $el = me._segmented(c);
                else if (c.type === 'select')    $el = me._select(c);
                if ($el) me.$controls.append($el);
            });
        },

        // textContent everywhere, never .html(): labels are host-supplied strings and some are
        // user-authored (workflow names), so they must never be parsed as markup.
        _chip: function (c) {
            var tone = (c.tone === 'green' || c.tone === 'amber') ? c.tone : 'grey';
            return $('<span class="eo-neg-chip ' + tone + '"></span>').text(c.label || '');
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
                disabled: !!(c.disabled || c.busy),
                onClick: function () { me._post({ type: 'action', id: c.id }); },
            });
        },

        // Icons arrive as raw SVG path geometry on a 24x24 viewBox, never as names: this file
        // cannot import the app's icon set, and shipping names would mean the image carries a
        // fixed catalogue that only a rebuild could extend. currentColor makes the glyph follow
        // the button's own text colour, including the selected state.
        _icon: function (pathData) {
            if (!pathData || typeof pathData !== 'string') return null;
            var svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
            svg.setAttribute('viewBox', '0 0 24 24');
            svg.setAttribute('class', 'eo-neg-ico');
            svg.setAttribute('aria-hidden', 'true');
            var path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
            // setAttribute, not innerHTML — this is host-supplied and must never be parsed as markup.
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
    }, DE.Controllers.NegotiationTab || {}));
});
