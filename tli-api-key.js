/**
 * TLI API Key Manager
 * -------------------
 * Shared across all Learning Design Suite tool pages.
 * Stores the TAMU AI Chat API key in sessionStorage (tab lifetime)
 * or localStorage (persistent) based on user preference.
 *
 * On load:
 *   1. Reads stored key and fills #apiKey input if present.
 *   2. Injects a small "key" icon button in the nav bar.
 *   3. Provides a settings modal for entering / clearing the key.
 *
 * On key change (input blur):
 *   Saves to the chosen storage automatically.
 */
(function () {
    'use strict';

    var STORAGE_KEY = 'tli_api_key';
    var PREF_KEY   = 'tli_api_key_persist'; // 'local' or 'session'

    /* ── Storage helpers ────────────────────────────── */
    function getPersistMode() {
        try { return localStorage.getItem(PREF_KEY) || 'session'; } catch(e) { return 'session'; }
    }

    function setPersistMode(mode) {
        try { localStorage.setItem(PREF_KEY, mode); } catch(e) {}
    }

    function getStore() {
        return getPersistMode() === 'local' ? localStorage : sessionStorage;
    }

    function readKey() {
        // Check both stores; prefer whichever has a value
        try {
            return localStorage.getItem(STORAGE_KEY) || sessionStorage.getItem(STORAGE_KEY) || '';
        } catch(e) { return ''; }
    }

    function saveKey(key) {
        var store = getStore();
        try {
            store.setItem(STORAGE_KEY, key);
            // Clear the other store so there's only one copy
            var other = (store === localStorage) ? sessionStorage : localStorage;
            other.removeItem(STORAGE_KEY);
        } catch(e) {}
    }

    function clearKey() {
        try { localStorage.removeItem(STORAGE_KEY); } catch(e) {}
        try { sessionStorage.removeItem(STORAGE_KEY); } catch(e) {}
    }

    /* ── Auto-fill on page load ─────────────────────── */
    function autoFill() {
        var input = document.getElementById('apiKey');
        if (input && !input.value) {
            input.value = readKey();
        }
    }

    /* ── Save on input change ───────────────────────── */
    function watchInput() {
        var input = document.getElementById('apiKey');
        if (!input) return;
        input.addEventListener('change', function () {
            var val = input.value.trim();
            if (val) saveKey(val);
        });
    }

    /* ── Modal markup ───────────────────────────────── */
    function injectModal() {
        if (document.getElementById('tliKeyModal')) return;

        var overlay = document.createElement('div');
        overlay.id = 'tliKeyModal';
        overlay.setAttribute('role', 'dialog');
        overlay.setAttribute('aria-label', 'API Key Settings');
        overlay.style.cssText = 'display:none;position:fixed;inset:0;z-index:9999;background:rgba(0,0,0,0.45);align-items:center;justify-content:center;';

        var stored = readKey();
        var persist = getPersistMode() === 'local';

        overlay.innerHTML =
            '<div style="background:#fff;border-radius:4px;max-width:440px;width:90%;padding:28px 28px 24px;position:relative;box-shadow:0 8px 32px rgba(0,0,0,0.18);">' +
                '<button id="tliKeyClose" style="position:absolute;top:12px;right:14px;background:none;border:none;font-size:20px;cursor:pointer;color:#666;line-height:1;" aria-label="Close">&times;</button>' +
                '<h3 style="font-family:Oswald,Arial,sans-serif;font-size:18px;font-weight:600;text-transform:uppercase;letter-spacing:0.3px;color:#313131;margin:0 0 4px;">API Key Settings</h3>' +
                '<p style="font-size:13px;color:#6B6B6B;margin:0 0 18px;line-height:1.5;">Enter your TAMU AI Chat API key once. It will auto-fill on every tool page.</p>' +
                '<label for="tliKeyInput" style="font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:0.5px;color:#500000;display:block;margin-bottom:6px;">API Key</label>' +
                '<input id="tliKeyInput" type="password" value="' + stored.replace(/"/g, '&quot;') + '" placeholder="sk-..." style="width:100%;padding:10px 12px;border:1px solid #ccc;border-radius:2px;font-size:14px;font-family:Work Sans,Arial,sans-serif;box-sizing:border-box;">' +
                '<div style="margin-top:12px;display:flex;align-items:center;gap:8px;">' +
                    '<input id="tliKeyPersist" type="checkbox"' + (persist ? ' checked' : '') + ' style="margin:0;">' +
                    '<label for="tliKeyPersist" style="font-size:13px;color:#4A4A4A;cursor:pointer;">Remember across browser sessions</label>' +
                '</div>' +
                '<p style="font-size:11px;color:#999;margin:6px 0 0 26px;line-height:1.4;">When unchecked, the key is cleared when you close this browser tab.</p>' +
                '<div style="margin-top:20px;display:flex;gap:10px;justify-content:flex-end;">' +
                    '<button id="tliKeyClear" style="padding:8px 18px;border:1px solid #ccc;border-radius:2px;background:#fff;font-size:13px;font-weight:600;cursor:pointer;color:#666;">Clear Key</button>' +
                    '<button id="tliKeySave" style="padding:8px 22px;border:none;border-radius:2px;background:#500000;color:#fff;font-size:13px;font-weight:600;cursor:pointer;">Save</button>' +
                '</div>' +
            '</div>';

        document.body.appendChild(overlay);

        /* ── Modal events ───── */
        document.getElementById('tliKeyClose').addEventListener('click', closeModal);
        overlay.addEventListener('click', function (e) { if (e.target === overlay) closeModal(); });

        document.getElementById('tliKeySave').addEventListener('click', function () {
            var val = document.getElementById('tliKeyInput').value.trim();
            var persistChecked = document.getElementById('tliKeyPersist').checked;
            setPersistMode(persistChecked ? 'local' : 'session');
            if (val) {
                saveKey(val);
            } else {
                clearKey();
            }
            // Update the page's #apiKey input
            var pageInput = document.getElementById('apiKey');
            if (pageInput) pageInput.value = val;
            closeModal();
            updateNavIcon();
        });

        document.getElementById('tliKeyClear').addEventListener('click', function () {
            clearKey();
            document.getElementById('tliKeyInput').value = '';
            var pageInput = document.getElementById('apiKey');
            if (pageInput) pageInput.value = '';
            closeModal();
            updateNavIcon();
        });
    }

    function openModal() {
        var modal = document.getElementById('tliKeyModal');
        if (!modal) return;
        // Refresh stored value
        document.getElementById('tliKeyInput').value = readKey();
        document.getElementById('tliKeyPersist').checked = (getPersistMode() === 'local');
        modal.style.display = 'flex';
        document.getElementById('tliKeyInput').focus();
    }

    function closeModal() {
        var modal = document.getElementById('tliKeyModal');
        if (modal) modal.style.display = 'none';
    }

    /* ── Nav icon ───────────────────────────────────── */
    function injectNavIcon() {
        // Look for either .nav-inner or .lds-nav-inner
        var nav = document.querySelector('.nav-inner') || document.querySelector('.lds-nav-inner');
        if (!nav || document.getElementById('tliKeyBtn')) return;

        var btn = document.createElement('button');
        btn.id = 'tliKeyBtn';
        btn.setAttribute('aria-label', 'API Key Settings');
        btn.title = 'API Key Settings';
        btn.style.cssText = 'margin-left:auto;background:none;border:1px solid #ccc;border-radius:2px;padding:6px 12px;cursor:pointer;display:flex;align-items:center;gap:6px;font-size:12px;font-weight:600;font-family:Work Sans,Arial,sans-serif;color:#500000;text-transform:uppercase;letter-spacing:0.3px;white-space:nowrap;';
        updateBtnContent(btn);
        btn.addEventListener('click', openModal);
        nav.appendChild(btn);
    }

    function updateBtnContent(btn) {
        if (!btn) btn = document.getElementById('tliKeyBtn');
        if (!btn) return;
        var hasKey = !!readKey();
        btn.innerHTML = (hasKey ? '&#128274; ' : '&#128275; ') + '<span>API Key</span>';
        btn.style.borderColor = hasKey ? '#2D6A4F' : '#ccc';
        btn.style.color = hasKey ? '#2D6A4F' : '#500000';
    }

    function updateNavIcon() {
        updateBtnContent();
    }

    /* ── Keyboard shortcut: Escape closes modal ────── */
    document.addEventListener('keydown', function (e) {
        if (e.key === 'Escape') closeModal();
    });

    /* ── Init ───────────────────────────────────────── */
    function init() {
        autoFill();
        watchInput();
        injectModal();
        injectNavIcon();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
