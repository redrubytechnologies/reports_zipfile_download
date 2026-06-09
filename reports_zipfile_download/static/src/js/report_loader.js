/** @odoo-module **/

import { onMounted } from "@odoo/owl";
import { useService } from "@web/core/utils/hooks";
import { FormController } from "@web/views/form/form_controller";
import { patch } from "@web/core/utils/patch";

// ─── Persistent background downloader ────────────────────────────────────────
// Uses sessionStorage so pending downloads survive SPA navigation within the tab.

const STORAGE_KEY = 'odoo_pending_report_downloads';
const POLL_MS     = 2000;
const TIMEOUT_MS  = 10 * 60 * 1000; // 10 minutes

function _getPending() {
    try { return JSON.parse(sessionStorage.getItem(STORAGE_KEY) || '[]'); }
    catch(e) { return []; }
}

function _setPending(list) {
    sessionStorage.setItem(STORAGE_KEY, JSON.stringify(list));
}

function _addPending(wizardId) {
    const list = _getPending().filter(p => p.id !== wizardId);
    list.push({ id: wizardId, start: Date.now() });
    _setPending(list);
}

async function _readWizard(wizardId) {
    const res = await fetch('/web/dataset/call_kw', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            jsonrpc: '2.0',
            id: Math.random(),
            method: 'call',
            params: {
                model: 'reports.wizard',
                method: 'read',
                args: [[wizardId], ['pdf_file', 'pdf_filename', 'excel_file', 'filename']],
                kwargs: { context: {} },
            },
        }),
    });
    const data = await res.json();
    return data.result && data.result[0] ? data.result[0] : null;
}

function _triggerDownload(wizardId, filename) {
    const url = `/web/content?model=reports.wizard&id=${wizardId}&field=pdf_file&download=true&filename=${encodeURIComponent(filename)}`;
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { if (a.parentNode) a.parentNode.removeChild(a); }, 200);
    _showDoneToast(filename);
}

function _showErrorToast(message) {
    if (!document.getElementById('_rpt_toast_style')) {
        const s = document.createElement('style');
        s.id = '_rpt_toast_style';
        s.textContent = `
            @keyframes _rpt_in{from{opacity:0;transform:translateX(60px)}to{opacity:1;transform:translateX(0)}}
            @keyframes _rpt_out{from{opacity:1}to{opacity:0}}
            ._rpt_toast{position:fixed;top:180px;right:20px;z-index:99999;color:#fff;
                padding:18px 20px;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.28);
                font-size:15px;font-family:sans-serif;max-width:380px;display:flex;
                align-items:flex-start;gap:12px;animation:_rpt_in .3s ease}
            ._rpt_toast.out{animation:_rpt_out .4s ease forwards}
        `;
        document.head.appendChild(s);
    }
    const id = '_rpt_toast_err_' + Date.now();
    const div = document.createElement('div');
    div.className = '_rpt_toast';
    div.id = id;
    div.style.background = '#dc3545';
    div.innerHTML =
        `<span style="font-size:22px;line-height:1.2">✕</span>` +
        `<div><b style="display:block;margin-bottom:4px;font-size:15px">PDF Generation Failed</b>` +
        `<span style="font-size:12px;opacity:.9;word-break:break-all">${message.replace(/^ERROR:\s*/i,'')}</span></div>` +
        `<span onclick="document.getElementById('${id}').remove()" ` +
        `style="cursor:pointer;margin-left:auto;font-size:17px;opacity:.7;flex-shrink:0">✕</span>`;
    document.body.appendChild(div);
    setTimeout(() => {
        div.classList.add('out');
        setTimeout(() => { if (div.parentNode) div.parentNode.removeChild(div); }, 400);
    }, 10000);
}

function _showDoneToast(filename) {
    if (!document.getElementById('_rpt_toast_style')) {
        const s = document.createElement('style');
        s.id = '_rpt_toast_style';
        s.textContent = `
            @keyframes _rpt_in{from{opacity:0;transform:translateX(60px)}to{opacity:1;transform:translateX(0)}}
            @keyframes _rpt_out{from{opacity:1}to{opacity:0}}
            ._rpt_toast{position:fixed;top:180px;right:20px;z-index:99999;background:#28a745;color:#fff;
                padding:18px 20px;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.28);
                font-size:15px;font-family:sans-serif;max-width:340px;display:flex;
                align-items:flex-start;gap:12px;animation:_rpt_in .3s ease}
            ._rpt_toast.out{animation:_rpt_out .4s ease forwards}
        `;
        document.head.appendChild(s);
    }
    const id = '_rpt_toast_' + Date.now();
    const div = document.createElement('div');
    div.className = '_rpt_toast';
    div.id = id;
    div.innerHTML =
        `<span style="font-size:22px;line-height:1.2">✓</span>` +
        `<div><b style="display:block;margin-bottom:4px;font-size:15px">${filename.endsWith('.zip') ? 'ZIP' : 'PDF'} Ready — Downloading</b>` +
        `<span style="font-size:12px;opacity:.85;word-break:break-all">${filename}</span></div>` +
        `<span onclick="document.getElementById('${id}').remove()" ` +
        `style="cursor:pointer;margin-left:auto;font-size:17px;opacity:.7;flex-shrink:0">✕</span>`;
    document.body.appendChild(div);
}

// Start once at module load — persists across SPA navigation
(function _startGlobalPoller() {
    async function poll() {
        const pending = _getPending();
        if (pending.length) {
            const remaining = [];
            for (const item of pending) {
                if (Date.now() - item.start > TIMEOUT_MS) continue; // expired
                try {
                    const rec = await _readWizard(item.id);
                    const pdfErr  = rec && rec.pdf_filename  && String(rec.pdf_filename).startsWith('ERROR:');
                    const xlsErr  = rec && rec.filename      && String(rec.filename).startsWith('ERROR:');
                    if (rec && rec.pdf_file) {
                        // PDF ready — auto-download
                        _triggerDownload(item.id, rec.pdf_filename || 'report.pdf');
                    } else if (pdfErr) {
                        // Background thread wrote an error — show red toast
                        _showErrorToast(rec.pdf_filename);
                    } else if (xlsErr) {
                        _showErrorToast(rec.filename);
                    } else {
                        remaining.push(item); // not ready yet — keep polling
                    }
                } catch(e) {
                    remaining.push(item); // network hiccup — keep trying
                }
            }
            _setPending(remaining);
        }
        setTimeout(poll, POLL_MS);
    }
    setTimeout(poll, POLL_MS);
})();
// ─────────────────────────────────────────────────────────────────────────────

patch(FormController.prototype, {
    setup() {
        super.setup(...arguments);
        this.orm = useService("orm");
        onMounted(() => { this._setupReportLoaderObserver(); });
    },

    _getOverlay() { return document.getElementById('report_loader_overlay'); },
    _showOverlay() { const ov = this._getOverlay(); if (ov) ov.classList.add('active'); },
    _hideOverlay() { const ov = this._getOverlay(); if (ov) ov.classList.remove('active'); },

    _setupReportLoaderObserver() {
        const self = this;

        function tryAttach() {
            const btn = document.querySelector('button[name="generate_report"]');
            if (!btn || btn.dataset.loaderAttached) return;
            btn.dataset.loaderAttached = 'true';
            btn.addEventListener('click', function () {
                // The 5 PDF print reports (quotation_print_report, dc_print_report,
                // po_print_report, invoice_print_report, installation_print_report)
                // run in a background thread and auto-download via the global poller —
                // skip the blocking loader for them.
                const reportType = self.model && self.model.root && self.model.root.data && self.model.root.data.report_type;
                const isPrintReport = reportType && reportType.includes('print_report');

                if (!isPrintReport) {
                    self._showOverlay();
                    self._startHideWatchers();
                }

                // Queue wizard ID for the global background poller (PDF downloads).
                // Use the model's resId — reliable even for new transient records
                // whose ID has not yet appeared in the URL hash.
                const wizardId = self.model && self.model.root && self.model.root.resId;
                if (wizardId) _addPending(wizardId);
            });
        }

        const domObserver = new MutationObserver(() => {
            const btn = document.querySelector('button[name="generate_report"]');
            if (btn && !btn.dataset.loaderAttached) tryAttach();
        });
        domObserver.observe(document.body, { childList: true, subtree: true });
        tryAttach();
    },

    _startHideWatchers() {
        const self = this;
        let hidden = false;
        const origFetch = window.fetch;
        let safetyTimer = null;

        const cleanup = () => {
            if (safetyTimer) clearTimeout(safetyTimer);
            window.removeEventListener('blur', onBlur);
            window.removeEventListener('focus', onFocus);
            // Restore fetch only if it's still our patched version
            if (window.fetch !== origFetch) window.fetch = origFetch;
        };

        const hide = () => {
            if (hidden) return;
            hidden = true;
            cleanup();
            self._hideOverlay();
        };

        // Safety: always hide after 10 minutes
        safetyTimer = setTimeout(hide, 10 * 60 * 1000);

        // Odoo 17 uses fetch (not XHR) for RPC calls.
        // Intercept fetch to detect when the generate_* call completes.
        window.fetch = function (...args) {
            const url = typeof args[0] === 'string' ? args[0]
                      : (args[0] && args[0].url) ? args[0].url : '';
            const isRpc = url.includes('/web/dataset/call_button') ||
                          url.includes('/web/dataset/call_kw');
            const promise = origFetch.apply(this, args);

            if (isRpc && !hidden) {
                let isGenerateCall = false;
                try {
                    const body = (args[1] && args[1].body) || '';
                    const parsed = JSON.parse(body);
                    const method = (parsed && parsed.params && parsed.params.method) || '';
                    isGenerateCall = method === 'generate_report' || /^generate_/.test(method);
                } catch (e) {}

                if (isGenerateCall) {
                    // When generate_* RPC returns, allow ~1.2 s for Odoo to
                    // process the returned action and re-render the form.
                    promise
                        .then(() => setTimeout(hide, 1200))
                        .catch(hide);
                }
            }

            return promise;
        };

        // If the browser opens a "Save file" dialog the window loses focus.
        let blurTime = null;
        const onBlur  = () => { blurTime = Date.now(); };
        const onFocus = () => { if (blurTime) hide(); };
        window.addEventListener('blur', onBlur);
        window.addEventListener('focus', onFocus);
    },
});
