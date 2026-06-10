/*
 * top.legal — Style tab (native)
 * ------------------------------------------------------------------
 * A dedicated "Style" ribbon tab that lists style presets directly in
 * the tab as clickable cards (apply on click), plus New / Edit / Delete
 * and a "None" reset. Presets live in the integration backend (the same
 * GET/PUT /styles endpoints the old Style-packages plugin used).
 *
 * Apply is fully NATIVE: it runs the document-builder script through
 * this.api.callCommand(fn, isRecalculate, cb) — the editor's own
 * "internal analog for plugins" (sdkjs common/apiBase.js). No plugin
 * iframe, no side panel. The builder body (GetStyle/GetTextPr/...) is
 * ported verbatim from the proven plugin.
 */
define([
    'core'
], function () {
    'use strict';

    DE.Controllers = DE.Controllers || {};

    DE.Controllers.StyleTab = Backbone.Controller.extend(_.extend({
        models: [],
        collections: [],
        views: [],

        // ---- state ----------------------------------------------------
        state: { defaultId: null, packages: [], selectedByDoc: {}, originalsByDoc: {} },
        selectedId: null,
        currentDoc: '',
        captureDone: false,
        originalSnapshot: null,
        editingId: null,
        previewTimer: null,
        editBaseline: null,

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

        // Backend base: editor is served by the Document Server (docs.top.legal),
        // backend sits behind /api on the same host. Dev falls back to :3001.
        backendBase: function () {
            var o = window.location.origin;
            return (o.indexOf('localhost') >= 0 || o.indexOf('127.0.0.1') >= 0) ? 'http://localhost:3001' : o + '/api';
        },

        userId: function () { return 'demo-user'; },  // no auth yet (#27/#28)

        // ===============================================================
        // Panel (ribbon content)
        // ===============================================================
        createToolbarPanel: function () {
            var me = this;
            me._injectStyles();
            me._panel = $(
                '<section class="panel" data-tab="style" role="tabpanel" aria-labelledby="style">' +
                    '<div class="group eo-style-group">' +
                        '<div class="eo-style-cards" id="eo-style-cards"></div>' +
                    '</div>' +
                '</section>'
            );
            me.$cards = me._panel.find('#eo-style-cards');
            me._renderCards();              // placeholder until data loads
            me._loadCollection();           // async fetch → re-render
            return me._panel;
        },

        getButtons: function () { return []; },

        _injectStyles: function () {
            if (document.getElementById('eo-style-tab-css')) return;
            var css =
                '.eo-style-group{display:flex;align-items:center;height:100%;overflow:hidden;}' +
                '.eo-style-cards{display:flex;align-items:center;gap:6px;overflow-x:auto;overflow-y:hidden;max-width:100%;padding:3px 4px;height:100%;box-sizing:border-box;}' +
                '.eo-style-card{position:relative;flex:0 0 auto;width:116px;height:46px;border:1px solid #cfcfcf;border-radius:3px;background:#fff;cursor:pointer;display:flex;align-items:center;padding:4px 7px;box-sizing:border-box;user-select:none;}' +
                '.eo-style-card:hover{border-color:#7d858c;}' +
                '.eo-style-card.selected{border-color:#446995;box-shadow:0 0 0 1px #446995 inset;}' +
                '.eo-style-card .sw{flex:0 0 auto;width:28px;height:28px;border-radius:3px;display:flex;align-items:center;justify-content:center;font-size:15px;line-height:1;margin-right:7px;}' +
                '.eo-style-card .meta{flex:1 1 auto;min-width:0;}' +
                '.eo-style-card .nm{font-size:11px;font-weight:bold;color:#363636;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}' +
                '.eo-style-card .sub{font-size:10px;color:#909090;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}' +
                '.eo-style-card .row-btns{position:absolute;top:2px;right:2px;display:none;gap:2px;}' +
                '.eo-style-card:hover .row-btns{display:flex;}' +
                '.eo-style-card .row-btns button{width:16px;height:16px;border:none;background:rgba(255,255,255,.85);border-radius:2px;cursor:pointer;font-size:10px;line-height:14px;padding:0;color:#444;}' +
                '.eo-style-card .row-btns button:hover{background:#e0e0e0;}' +
                '.eo-style-card.none .sw{background:#f0f0f0;color:#888;}' +
                '.eo-style-card.add{justify-content:center;color:#446995;font-size:12px;font-weight:bold;width:88px;border-style:dashed;}' +
                '.eo-style-card .badge{display:inline-block;font-size:8px;background:#446995;color:#fff;border-radius:2px;padding:0 3px;margin-left:4px;vertical-align:middle;}';
            var st = document.createElement('style');
            st.id = 'eo-style-tab-css';
            st.type = 'text/css';
            st.innerHTML = css;
            document.getElementsByTagName('head')[0].appendChild(st);
        },

        _renderCards: function () {
            var me = this;
            if (!me.$cards) return;
            me.$cards.empty();

            // None
            var none = $('<div class="eo-style-card none' + (me.selectedId === '__none__' ? ' selected' : '') + '">' +
                '<div class="sw">&#8709;</div><div class="meta"><div class="nm">None</div><div class="sub">reset to normal</div></div></div>');
            none.on('click', function () {
                me.selectedId = '__none__';
                me._persistSelected();
                me._renderCards();
                me._applyPackage({ id: '__none__', name: 'None' });
            });
            me.$cards.append(none);

            // Presets
            (me.state.packages || []).forEach(function (pkg) {
                var h1 = (pkg.styles && pkg.styles['Heading 1']) || {};
                var body = (pkg.styles && pkg.styles['Normal']) || {};
                var card = $('<div class="eo-style-card' + (pkg.id === me.selectedId ? ' selected' : '') + '"></div>');
                var sw = $('<div class="sw">Aa</div>');
                sw.css({ background: (h1.color || '#666') + '22', color: h1.color || '#666', fontFamily: h1.font || 'serif' });
                var meta = $('<div class="meta"></div>');
                var nm = $('<div class="nm"></div>').text(pkg.name).attr('title', pkg.name);
                if (pkg.id === me.state.defaultId) nm.append(' ').append($('<span class="badge">default</span>'));
                card.attr('title', pkg.name);
                var sub = $('<div class="sub"></div>').text((body.font || '?') + ' · ' + (body.size || '?') + 'pt');
                meta.append(nm).append(sub);
                var btns = $('<div class="row-btns"></div>');
                var edit = $('<button title="Edit">&#9998;</button>').on('click', function (e) { e.stopPropagation(); me._openEditor(pkg.id); });
                var del = $('<button title="Delete">&#128465;</button>').on('click', function (e) { e.stopPropagation(); me._deletePackage(pkg.id); });
                btns.append(edit).append(del);
                card.append(sw).append(meta).append(btns);
                card.on('click', function () {
                    me.selectedId = pkg.id;
                    me._persistSelected();
                    me._renderCards();
                    me._applyPackage(pkg);
                });
                me.$cards.append(card);
            });

            // + New
            var add = $('<div class="eo-style-card add" title="New preset">+ New</div>');
            add.on('click', function () { me._openEditor(null); });
            me.$cards.append(add);
        },

        // ===============================================================
        // Backend
        // ===============================================================
        _loadCollection: function () {
            var me = this;
            fetch(me.backendBase() + '/styles?userId=' + encodeURIComponent(me.userId()))
                .then(function (r) { if (!r.ok) throw new Error('load ' + r.status); return r.json(); })
                .then(function (col) {
                    me.state = {
                        defaultId: col.defaultId || null,
                        selectedByDoc: col.selectedByDoc || {},
                        originalsByDoc: col.originalsByDoc || {},
                        packages: col.packages || []
                    };
                    if (!me.state.packages.length) {
                        me.state.packages = me._starterPackages();
                        me.state.defaultId = 'corporate';
                        me._saveCollection();
                    }
                    me.currentDoc = (me.appConfig && me.appConfig.title) || (me.api && me.api.asc_getDocumentName && me.api.asc_getDocumentName()) || '';
                    me.selectedId = (me.currentDoc && me.state.selectedByDoc[me.currentDoc]) || null;
                    if (me.currentDoc && me.state.originalsByDoc[me.currentDoc]) {
                        me.originalSnapshot = me.state.originalsByDoc[me.currentDoc];
                        me.captureDone = true;
                    }
                    me._renderCards();
                })
                .catch(function () { /* leave None/+New visible */ });
        },

        _saveCollection: function () {
            var me = this;
            return fetch(me.backendBase() + '/styles?userId=' + encodeURIComponent(me.userId()), {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(me.state)
            }).then(function (r) { if (!r.ok) throw new Error('save ' + r.status); return r.json(); });
        },

        _persistSelected: function () {
            if (!this.state.selectedByDoc) this.state.selectedByDoc = {};
            if (this.currentDoc) this.state.selectedByDoc[this.currentDoc] = this.selectedId;
            this._saveCollection().catch(function () {});
        },

        _persistOriginal: function () {
            if (!this.currentDoc || !this.originalSnapshot) return;
            if (!this.state.originalsByDoc) this.state.originalsByDoc = {};
            if (!this.state.originalsByDoc[this.currentDoc]) {
                this.state.originalsByDoc[this.currentDoc] = this.originalSnapshot;
                this._saveCollection().catch(function () {});
            }
        },

        _pkgById: function (id) {
            var p = this.state.packages || [];
            for (var i = 0; i < p.length; i++) if (p[i].id === id) return p[i];
            return null;
        },

        _deletePackage: function (id) {
            var me = this;
            me.state.packages = (me.state.packages || []).filter(function (p) { return p.id !== id; });
            if (me.state.defaultId === id) me.state.defaultId = me.state.packages[0] ? me.state.packages[0].id : null;
            if (me.selectedId === id) me.selectedId = null;
            me._saveCollection().then(function () { me._renderCards(); }).catch(function () {});
        },

        // ===============================================================
        // Apply (NATIVE builder via this.api.callCommand)
        // ===============================================================
        _applyPackage: function (pkg, onDone) {
            var me = this;
            if (!me.api || !me.api.callCommand) return;
            if (!me.captureDone) {
                me.captureDone = true;
                me._captureOriginal(function () { me._reallyApply(pkg, onDone); });
            } else {
                me._reallyApply(pkg, onDone);
            }
        },

        _readStyles: function (cb) {
            // Read-only builder script: snapshot the named-style definitions.
            this.api.callCommand(function () {
                var doc = Api.GetDocument();
                var names = ["Normal", "Title", "Heading 1", "Heading 2", "Heading 3"];
                function toHex(x) { return ("0" + ((x || 0) & 255).toString(16)).slice(-2); }
                function g(fn, dflt) { try { var v = fn(); return (v === undefined || v === null) ? dflt : v; } catch (e) { return dflt; } }
                function d20(x) { return (x === undefined || x === null) ? null : x / 20; }  // twips→pt, null stays null (inherit)
                var out = {};
                names.forEach(function (n) {
                    try {
                        var s = doc.GetStyle(n);
                        if (!s) return;
                        var tp = s.GetTextPr(), pp = s.GetParaPr();
                        var hex = null, col = g(function () { return tp.GetColor(); }, null);
                        if (col) { try { var rgb = col.GetRGB(); hex = "#" + toHex(rgb.r) + toHex(rgb.g) + toHex(rgb.b); } catch (e) { } }
                        var rule = g(function () { return pp.GetSpacingLineRule(); }, null);
                        var lv = g(function () { return pp.GetSpacingLineValue(); }, 240);
                        out[n] = {
                            font: g(function () { return tp.GetFontFamily(); }, null),
                            size: g(function () { return tp.GetFontSize(); }, null),
                            color: hex,
                            bold: !!g(function () { return tp.GetBold(); }, false),
                            italic: !!g(function () { return tp.GetItalic(); }, false),
                            align: g(function () { return pp.GetJc(); }, null),
                            spacingBefore: d20(g(function () { return pp.GetSpacingBefore(); }, null)),
                            spacingAfter: d20(g(function () { return pp.GetSpacingAfter(); }, null)),
                            lineSpacing: rule === "auto" ? (lv || 240) / 240 : null
                        };
                    } catch (e) { }
                });
                return JSON.stringify(out);
            }, false, function (ret) {   // isRecalculate=false (read-only)
                var o = null;
                if (ret) { try { var p = JSON.parse(ret); if (p && Object.keys(p).length) o = p; } catch (e) { } }
                cb(o);
            });
        },

        _captureOriginal: function (done) {
            var me = this;
            me._readStyles(function (o) {
                if (o) me.originalSnapshot = o;
                me._persistOriginal();
                if (done) done();
            });
        },

        _reallyApply: function (pkg, onDone) {
            var me = this;
            var styles = (pkg.id === '__none__') ? (me.originalSnapshot || me._normalDefault()) : pkg.styles;
            window.Asc = window.Asc || {};
            window.Asc.scope = window.Asc.scope || {};
            window.Asc.scope.eoStyle = { styles: styles };
            me.api.callCommand(function () {
                var styles = Asc.scope.eoStyle.styles;
                var doc = Api.GetDocument();
                function hexToRgb(h) {
                    h = String(h || '').replace('#', '');
                    return { r: parseInt(h.substr(0, 2), 16) || 0, g: parseInt(h.substr(2, 2), 16) || 0, b: parseInt(h.substr(4, 2), 16) || 0 };
                }
                for (var name in styles) {
                    if (!styles.hasOwnProperty(name)) continue;
                    var d = styles[name];
                    var style = doc.GetStyle(name);
                    if (!style) continue;
                    var tp = style.GetTextPr();
                    if (d.font) tp.SetFontFamily(d.font);
                    if (d.size) tp.SetFontSize(d.size);
                    if (d.color) { var c = hexToRgb(d.color); tp.SetColor(c.r, c.g, c.b, false); }
                    else if (d.color === null) tp.SetColor(0, 0, 0, true);
                    tp.SetBold(!!d.bold);
                    tp.SetItalic(!!d.italic);
                    var pp = style.GetParaPr();
                    if (d.align) pp.SetJc(d.align);
                    if (d.spacingBefore != null) pp.SetSpacingBefore(Math.round(d.spacingBefore * 20));
                    if (d.spacingAfter != null) pp.SetSpacingAfter(Math.round(d.spacingAfter * 20));
                    if (d.lineSpacing) pp.SetSpacingLine(Math.round(d.lineSpacing * 240), 'auto');
                }
            }, true, function () { if (onDone) onDone(); });   // isRecalculate=true → reflow after restyle
        },

        _normalDefault: function () {
            return {
                'Normal':    { font: 'Calibri', size: 11, color: null, bold: false, italic: false, align: 'left', spacingBefore: 0, spacingAfter: 8, lineSpacing: 1.08 },
                'Title':     { font: 'Calibri Light', size: 28, color: null, bold: false, spacingBefore: 0, spacingAfter: 0 },
                'Heading 1': { font: 'Calibri Light', size: 16, color: '#2F5496', bold: false, spacingBefore: 12, spacingAfter: 0 },
                'Heading 2': { font: 'Calibri Light', size: 13, color: '#2F5496', bold: false, spacingBefore: 2, spacingAfter: 0 },
                'Heading 3': { font: 'Calibri Light', size: 12, color: '#1F3763', bold: false, spacingBefore: 2, spacingAfter: 0 }
            };
        },

        _starterPackages: function () {
            return [
                { id: 'corporate', name: 'Corporate', styles: {
                    'Normal':    { font: 'Calibri', size: 11, color: '#222222', lineSpacing: 1.15, spacingAfter: 8 },
                    'Title':     { font: 'Calibri Light', size: 28, color: '#1F4E79', bold: true },
                    'Heading 1': { font: 'Calibri Light', size: 18, color: '#1F4E79', bold: true, spacingBefore: 12, spacingAfter: 4 },
                    'Heading 2': { font: 'Calibri Light', size: 14, color: '#2E74B5', bold: true, spacingBefore: 10, spacingAfter: 4 },
                    'Heading 3': { font: 'Calibri', size: 12, color: '#2E74B5', bold: true, spacingBefore: 8, spacingAfter: 2 } } },
                { id: 'legal-serif', name: 'Legal Serif', styles: {
                    'Normal':    { font: 'Times New Roman', size: 12, color: '#000000', lineSpacing: 1.5, spacingAfter: 6 },
                    'Title':     { font: 'Georgia', size: 24, color: '#1a1a1a', bold: true },
                    'Heading 1': { font: 'Georgia', size: 16, color: '#1a1a1a', bold: true, spacingBefore: 12, spacingAfter: 4 },
                    'Heading 2': { font: 'Georgia', size: 14, color: '#333333', bold: true, spacingBefore: 10, spacingAfter: 4 },
                    'Heading 3': { font: 'Times New Roman', size: 13, color: '#333333', bold: true, spacingBefore: 8, spacingAfter: 2 } } },
                { id: 'modern', name: 'Modern', styles: {
                    'Normal':    { font: 'Verdana', size: 10, color: '#2d2d2d', lineSpacing: 1.5, spacingAfter: 10 },
                    'Title':     { font: 'Arial', size: 26, color: '#0b5c4f', bold: true },
                    'Heading 1': { font: 'Arial', size: 17, color: '#0b5c4f', bold: true, spacingBefore: 12, spacingAfter: 4 },
                    'Heading 2': { font: 'Arial', size: 13, color: '#12866f', bold: true, spacingBefore: 10, spacingAfter: 4 },
                    'Heading 3': { font: 'Arial', size: 11, color: '#12866f', bold: true, spacingBefore: 8, spacingAfter: 2 } } }
            ];
        },

        // ===============================================================
        // Edit / New dialog
        // ===============================================================
        _openEditor: function (id) {
            var me = this;
            me.editingId = id || null;
            me._previewApplied = false;
            // Stop the editor from consuming keystrokes while the dialog is open.
            if (me.api && me.api.asc_enableKeyEvents) me.api.asc_enableKeyEvents(false);
            var pkg = id ? me._pkgById(id) : null;
            var n = (pkg && pkg.styles['Normal']) || {};
            var t = (pkg && pkg.styles['Title']) || {};
            var h1 = (pkg && pkg.styles['Heading 1']) || {};
            var h2 = (pkg && pkg.styles['Heading 2']) || {};
            var h3 = (pkg && pkg.styles['Heading 3']) || {};

            // remember current styling so Cancel restores the live preview
            me.editBaseline = null;
            me._readStyles(function (o) { me.editBaseline = o; });

            var v = function (x, d) { return (x == null ? d : x); };
            var body =
                '<div class="eo-style-form" style="padding:8px 4px;">' +
                  '<div class="form-row"><label>Name</label><input type="text" id="eo-f-name" value="' + me._esc(pkg ? pkg.name : '') + '"></div>' +
                  '<div class="form-sec">Body</div>' +
                  '<div class="form-row"><label>Font</label><select id="eo-f-body-font">' + me._fontOptions(v(n.font, 'Calibri')) + '</select></div>' +
                  '<div class="form-row"><label>Size</label><input type="number" id="eo-f-body-size" value="' + v(n.size, 11) + '"></div>' +
                  '<div class="form-row"><label>Color</label><input type="color" id="eo-f-body-color" value="' + (n.color || '#222222') + '"></div>' +
                  '<div class="form-row"><label>Line spacing</label><input type="number" step="0.05" id="eo-f-line" value="' + v(n.lineSpacing, 1.15) + '"></div>' +
                  '<div class="form-row"><label>Space after (pt)</label><input type="number" id="eo-f-after" value="' + v(n.spacingAfter, 8) + '"></div>' +
                  '<div class="form-sec">Headings</div>' +
                  '<div class="form-row"><label>Heading font</label><select id="eo-f-head-font">' + me._fontOptions(v(h1.font, 'Calibri Light')) + '</select></div>' +
                  '<div class="form-row"><label>Heading color</label><input type="color" id="eo-f-head-color" value="' + (h1.color || '#1F4E79') + '"></div>' +
                  '<div class="form-row"><label>Title size</label><input type="number" id="eo-f-title-size" value="' + v(t.size, 28) + '"></div>' +
                  '<div class="form-row"><label>H1 size</label><input type="number" id="eo-f-h1-size" value="' + v(h1.size, 18) + '"></div>' +
                  '<div class="form-row"><label>H2 size</label><input type="number" id="eo-f-h2-size" value="' + v(h2.size, 14) + '"></div>' +
                  '<div class="form-row"><label>H3 size</label><input type="number" id="eo-f-h3-size" value="' + v(h3.size, 12) + '"></div>' +
                '</div>';

            me._injectFormCss();
            var win = new Common.UI.Window({
                title: pkg ? 'Edit preset' : 'New preset',
                width: 340,
                height: 560,
                cls: 'modal-dlg',
                buttons: [
                    { value: 'ok', caption: 'Save', primary: true },
                    { value: 'cancel', caption: 'Cancel' }
                ],
                // NOTE: the body comes from the `tpl` option (Window.js renders
                // <div class="body"><%= tpl %>…</div>); the footer is auto-built
                // from `buttons`. Don't pass `template` — render() ignores it.
                tpl: '<div class="eo-style-scroll" style="max-height:430px;overflow-y:auto;padding:0 6px;">' + body + '</div>'
            });
            win.show();

            // A raw Common.UI.Window doesn't auto-wire footer buttons (that wiring
            // lives only in the alert/confirm factory), so do it ourselves. Closing
            // via X/Esc reverts the live preview, like Cancel.
            var resolved = false;
            var finish = function (result) {
                if (resolved) return; resolved = true;
                if (result === 'ok') me._saveFromForm(win); else me._cancelEdit();
                if (me.api && me.api.asc_enableKeyEvents) me.api.asc_enableKeyEvents(true);  // hand input back to the doc
            };
            win.$window.find('.footer .dlg-btn').on('click', function (e) {
                finish(e.currentTarget.getAttribute('result'));
                win.close();
            });
            win.on('close', function () { finish('cancel'); });

            // live preview on input (debounced)
            var ids = ['eo-f-name', 'eo-f-body-font', 'eo-f-body-size', 'eo-f-body-color', 'eo-f-line', 'eo-f-after',
                'eo-f-head-font', 'eo-f-head-color', 'eo-f-title-size', 'eo-f-h1-size', 'eo-f-h2-size', 'eo-f-h3-size'];
            ids.forEach(function (fid) {
                var el = document.getElementById(fid);
                if (el) { el.addEventListener('input', function () { me._schedulePreview(); }); el.addEventListener('change', function () { me._schedulePreview(); }); }
            });
        },

        // Build <option> markup for a font <select> from the editor's font list.
        _fontOptions: function (selected) {
            var names = this._fontList();
            if (selected && names.indexOf(selected) < 0) names.unshift(selected);
            var out = '';
            for (var i = 0; i < names.length; i++) {
                var n = names[i];
                out += '<option value="' + this._esc(n) + '"' + (n === selected ? ' selected' : '') + '>' + this._esc(n) + '</option>';
            }
            return out;
        },

        _fontList: function () {
            if (this._fontNames) return this._fontNames.slice();
            var names = [];
            try {
                var list = this.api && this.api.pluginMethod_GetFontList && this.api.pluginMethod_GetFontList();
                if (list && list.length) {
                    var seen = {};
                    for (var i = 0; i < list.length; i++) {
                        var nm = list[i] && list[i].m_wsFontName;
                        if (nm && !seen[nm]) { seen[nm] = 1; names.push(nm); }
                    }
                    names.sort();
                }
            } catch (e) { }
            if (!names.length) {
                names = ['Arial', 'Calibri', 'Calibri Light', 'Cambria', 'Georgia', 'Times New Roman', 'Verdana', 'Tahoma', 'Trebuchet MS', 'Courier New'];
            }
            this._fontNames = names;
            return names.slice();
        },

        _injectFormCss: function () {
            if (document.getElementById('eo-style-form-css')) return;
            var css =
                '.eo-style-form .form-row{display:flex;align-items:center;margin:5px 0;}' +
                '.eo-style-form .form-row label{flex:0 0 120px;font-size:11px;color:#444;}' +
                '.eo-style-form .form-row input,.eo-style-form .form-row select{flex:1 1 auto;height:24px;border:1px solid #cfcfcf;border-radius:2px;padding:0 6px;box-sizing:border-box;background:#fff;}' +
                '.eo-style-form .form-row input[type=color]{padding:1px;height:24px;}' +
                '.eo-style-form .form-sec{font-size:11px;font-weight:bold;color:#446995;margin:10px 0 4px;border-bottom:1px solid #eee;}';
            var st = document.createElement('style');
            st.id = 'eo-style-form-css'; st.type = 'text/css'; st.innerHTML = css;
            document.getElementsByTagName('head')[0].appendChild(st);
        },

        _formVal: function (id, dflt) { var el = document.getElementById(id); var v = parseFloat(el && el.value); return isFinite(v) ? v : dflt; },
        _formStr: function (id, dflt) { var el = document.getElementById(id); return (el && el.value) || dflt; },

        _buildStylesFromForm: function () {
            var bodyFont = this._formStr('eo-f-body-font', 'Calibri');
            var bodyColor = this._formStr('eo-f-body-color', '#222222');
            var headFont = this._formStr('eo-f-head-font', 'Calibri Light');
            var headColor = this._formStr('eo-f-head-color', '#1F4E79');
            return {
                'Normal':    { font: bodyFont, size: this._formVal('eo-f-body-size', 11), color: bodyColor, lineSpacing: this._formVal('eo-f-line', 1.15), spacingAfter: this._formVal('eo-f-after', 8) },
                'Title':     { font: headFont, size: this._formVal('eo-f-title-size', 28), color: headColor, bold: true },
                'Heading 1': { font: headFont, size: this._formVal('eo-f-h1-size', 18), color: headColor, bold: true, spacingBefore: 12, spacingAfter: 4 },
                'Heading 2': { font: headFont, size: this._formVal('eo-f-h2-size', 14), color: headColor, bold: true, spacingBefore: 10, spacingAfter: 4 },
                'Heading 3': { font: headFont, size: this._formVal('eo-f-h3-size', 12), color: headColor, bold: true, spacingBefore: 8, spacingAfter: 2 }
            };
        },

        _schedulePreview: function () {
            var me = this;
            if (me.previewTimer) clearTimeout(me.previewTimer);
            me.previewTimer = setTimeout(function () {
                me.previewTimer = null;
                me._previewApplied = true;
                // The recalc after callCommand makes the editor reclaim canvas focus,
                // which would send the next keystrokes into the document. Remember the
                // focused field + caret and restore them once the preview is applied.
                var ae = document.activeElement;
                var aeId = ae && ae.id, ss = null, se = null;
                try { ss = ae.selectionStart; se = ae.selectionEnd; } catch (e) { }
                me._applyPackage({ id: '__preview__', name: 'Preview', styles: me._buildStylesFromForm() }, function () {
                    if (me.api && me.api.asc_enableKeyEvents) me.api.asc_enableKeyEvents(false);
                    var el = aeId && document.getElementById(aeId);
                    if (el) { try { el.focus(); if (ss != null && el.setSelectionRange) el.setSelectionRange(ss, se); } catch (e) { } }
                });
            }, 350);
        },

        _saveFromForm: function (win) {
            var me = this;
            if (me.previewTimer) { clearTimeout(me.previewTimer); me.previewTimer = null; }
            var name = (me._formStr('eo-f-name', '') || '').trim() || 'Untitled';
            var styles = me._buildStylesFromForm();
            var pkg;
            if (me.editingId && me._pkgById(me.editingId)) {
                pkg = me._pkgById(me.editingId);
                pkg.name = name; pkg.styles = styles;
            } else {
                pkg = { id: 'pkg-' + Date.now().toString(36) + Math.floor(Math.random() * 1e4).toString(36), name: name, styles: styles };
                me.state.packages.push(pkg);
                if (!me.state.defaultId) me.state.defaultId = pkg.id;
            }
            me.selectedId = pkg.id;
            me._persistSelected();
            me._saveCollection().then(function () { me._renderCards(); }).catch(function () {});
            me._applyPackage(pkg);
        },

        _cancelEdit: function () {
            var me = this;
            if (me.previewTimer) { clearTimeout(me.previewTimer); me.previewTimer = null; }
            if (!me._previewApplied) return;   // nothing was previewed → leave the document untouched
            var baseline = me.editBaseline;
            if (!baseline) {
                if (me.editingId && me._pkgById(me.editingId)) baseline = me._pkgById(me.editingId).styles;
                else baseline = me.originalSnapshot || me._normalDefault();
            }
            if (baseline) me._applyPackage({ id: '__baseline__', name: 'Reverted', styles: baseline });
        },

        _esc: function (s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;'); }

    }, DE.Controllers.StyleTab || {}));
});
