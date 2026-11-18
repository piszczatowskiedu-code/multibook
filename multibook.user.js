// ==UserScript==
// @name         MultiKsięgarnia - Panel produktów
// @namespace    http://tampermonkey.net/
// @version         4.11.10
// @description  Panel boczny + fiszka + sync + edycja + okładka + filtr + podgląd okładki. Obsługa: TaniaKsiążka, Platon, Bonito, Tantis, Świat Książki, Dressler, Empik, Ateneum
// @author       You
// @match        https://www.taniaksiazka.pl/*
// @match        https://taniaksiazka.pl/*
// @match        https://platon.com.pl/*
// @match        https://www.platon.com.pl/*
// @match        https://bonito.pl/*
// @match        https://www.bonito.pl/*
// @match        https://tantis.pl/*
// @match        https://www.tantis.pl/*
// @match        https://swiatksiazki.pl/*
// @match        https://www.swiatksiazki.pl/*
// @match        https://dressler.com.pl/*
// @match        https://www.dressler.com.pl/*
// @match        https://www.empik.com/*
// @match        https://empik.com/*
// @match        https://www.ateneum.net.pl/*
// @match        https://ateneum.net.pl/*
// @match        https://www.ateneum.pl/*
// @match        https://ateneum.pl/*
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_addStyle
// @grant        GM_setClipboard
// @grant        GM_addValueChangeListener
// @grant        GM_xmlhttpRequest
// @connect      *
// @require      https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
// @run-at       document-idle
// ==/UserScript==

(function() {
    'use strict';

    // Nie uruchamiaj w iframe (np. podpowiedzi wyszukiwarki Bonito)
    if (window.top !== window.self) return;
    if (document.getElementById('tk-root')) return;

    // ========== STYLE ==========
    const PANEL_CSS = `
        /* ===== Izolacja od stylów strony (ScandiPWA / Magento itd.) ===== */
        #tk-panel,
        #tk-panel *,
        #tk-panel *::before,
        #tk-panel *::after {
            box-sizing: border-box !important;
            font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important;
            letter-spacing: normal !important;
            text-transform: none !important;
            text-shadow: none !important;
            box-shadow: none;
            max-width: none;
        }
        #tk-panel {
            font-size: 13px !important;
            line-height: 1.4 !important;
            color: var(--text) !important;
            -webkit-font-smoothing: antialiased;
        }
        #tk-panel img {
            max-width: none !important;
            max-height: none !important;
            width: auto;
            height: auto;
            display: inline-block;
            vertical-align: middle;
            object-fit: initial;
        }
        #tk-panel img.tk-cover {
            width: 48px !important;
            height: 66px !important;
            max-width: 48px !important;
            max-height: 66px !important;
            min-width: 48px !important;
            min-height: 66px !important;
            object-fit: cover !important;
            display: block !important;
            flex-shrink: 0 !important;
        }
        #tk-panel .tk-cover-placeholder {
            width: 48px !important;
            height: 66px !important;
            min-width: 48px !important;
            flex-shrink: 0 !important;
        }
        #tk-panel a {
            text-decoration: none !important;
            color: inherit;
        }
        #tk-panel button {
            font-family: inherit !important;
            font-size: inherit;
            line-height: normal !important;
            appearance: none !important;
            -webkit-appearance: none !important;
        }
        #tk-panel ul, #tk-panel li {
            list-style: none !important;
            margin: 0;
            padding: 0;
        }
        #tk-panel h1, #tk-panel h2, #tk-panel h3 {
            font-size: inherit;
            font-weight: inherit;
            margin: 0;
            padding: 0;
            line-height: inherit;
        }
        #tk-panel input,
        #tk-panel textarea {
            font-family: inherit !important;
            font-size: 12.5px !important;
            line-height: normal !important;
        }

        #tk-panel, .tk-grid-modal {
            --bg: #ffffff;
            --bg-header: #f8fafc;
            --bg-list: #f1f5f9;
            --bg-item: #ffffff;
            --border: #e2e8f0;
            --border-item: #e2e8f0;
            --text: #1e293b;
            --text-title: #0f172a;
            --text-muted: #64748b;
            --text-label: #94a3b8;
            --text-value: #334155;
            --count-bg: #e2e8f0;
            --count-text: #475569;
            --icon-color: #94a3b8;
            --scrollbar: #cbd5e1;
            --remove-hover-bg: #fee2e2;
            --empty: #94a3b8;
            --accent: #3b82f6;
            --edit-bg: #e2e8f0;
            --edit-border: #93c5fd;
            --edit-btn: #64748b;
            --pencil: #475569;
            --pencil-hover: #1e293b;
            --row-hover-bg: #f1f5f9;
        }
        #tk-panel.dark, .tk-grid-modal.dark {
            --bg: #0f1117;
            --bg-header: #131722;
            --bg-list: #0f1117;
            --bg-item: #161b26;
            --border: #1e2433;
            --border-item: #1e293b;
            --text: #e2e8f0;
            --text-title: #f1f5f9;
            --text-muted: #94a3b8;
            --text-label: #64748b;
            --text-value: #cbd5e1;
            --count-bg: #1e293b;
            --count-text: #94a3b8;
            --icon-color: #64748b;
            --scrollbar: #334155;
            --remove-hover-bg: #1e293b;
            --empty: #475569;
            --accent: #60a5fa;
            --edit-bg: #1e293b;
            --edit-border: #3b82f6;
            --edit-btn: #94a3b8;
            --pencil: #94a3b8;
            --pencil-hover: #e2e8f0;
            --row-hover-bg: #1e293b;
        }
        #tk-panel {
            position: fixed !important;
            top: 0 !important;
            right: 0 !important;
            left: auto !important;
            width: 400px !important;
            height: 100vh !important;
            max-height: 100vh !important;
            background: var(--bg) !important;
            color: var(--text) !important;
            z-index: 2147483000 !important;
            font-size: 13px !important;
            box-shadow: -6px 0 24px rgba(0,0,0,0.15) !important;
            display: flex !important;
            flex-direction: column !important;
            transition: transform 0.3s ease, background 0.3s ease, color 0.3s ease;
            border-left: 1px solid var(--border) !important;
            overflow: hidden !important;
        }
        #tk-panel.hidden { transform: translateX(100%); }
        #tk-toggle {
            position: fixed !important;
            top: 50% !important;
            right: 0 !important;
            transform: translateY(-50%) !important;
            background: #3b82f6 !important;
            color: #fff !important;
            border: none !important;
            border-radius: 10px 0 0 10px !important;
            padding: 14px 9px !important;
            cursor: pointer !important;
            z-index: 2147483001 !important;
            font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important;
            font-weight: 700 !important;
            font-size: 13px !important;
            font-style: normal !important;
            text-transform: none !important;
            text-shadow: none !important;
            box-shadow: -2px 0 10px rgba(0,0,0,0.25) !important;
            transition: right 0.3s ease, background 0.15s !important;
            writing-mode: vertical-rl !important;
            text-orientation: mixed !important;
            letter-spacing: 0.5px !important;
            margin: 0 !important;
            width: auto !important;
            height: auto !important;
            min-width: 0 !important;
            min-height: 0 !important;
            max-width: none !important;
            max-height: none !important;
            line-height: 1.2 !important;
            appearance: none !important;
            -webkit-appearance: none !important;
        }
        #tk-toggle:hover { background: #60a5fa !important; }
        /* Gdy panel otwarty – przycisk przy jego lewej krawędzi */
        #tk-panel:not(.hidden) ~ #tk-toggle {
            right: 400px !important;
        }
        .tk-theme-bar { display: none; }
        .tk-theme-label {
            display: flex;
            align-items: center;
            gap: 5px;
            font-weight: 600;
            letter-spacing: 0.2px;
            font-size: 12px;
            color: var(--text-muted);
            white-space: nowrap;
        }
        .tk-theme-label svg {
            width: 14px;
            height: 14px;
            flex-shrink: 0;
        }
        .tk-theme-switch {
            position: relative;
            display: inline-block;
            width: 42px;
            height: 22px;
            flex-shrink: 0;
        }
        .tk-theme-switch input { opacity: 0; width: 0; height: 0; }
        .tk-theme-slider {
            position: absolute;
            cursor: pointer;
            inset: 0;
            background: #cbd5e1;
            border-radius: 999px;
            transition: background 0.25s;
        }
        .tk-theme-slider::before {
            content: "";
            position: absolute;
            height: 16px;
            width: 16px;
            left: 3px;
            top: 3px;
            background: #fff;
            border-radius: 50%;
            transition: transform 0.25s;
            box-shadow: 0 1px 3px rgba(0,0,0,0.2);
        }
        .tk-theme-switch input:checked + .tk-theme-slider { background: #475569; }
        .tk-theme-switch input:checked + .tk-theme-slider::before { transform: translateX(20px); }
        .tk-theme-icon { width: 14px; height: 14px; }
        .tk-header {
            padding: 12px 14px;
            background: var(--bg-header);
            border-bottom: 1px solid var(--border);
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 10px;
        }
        .tk-header-left {
            display: flex;
            align-items: center;
            gap: 8px;
            min-width: 0;
            flex: 1;
        }
        .tk-header-right {
            display: flex;
            align-items: center;
            gap: 8px;
            flex-shrink: 0;
        }
        .tk-header-icon {
            width: 22px;
            height: 22px;
            flex-shrink: 0;
            color: var(--accent);
        }
        .tk-header h2 {
            margin: 0;
            font-size: 15px;
            font-weight: 650;
            color: var(--text-title);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .tk-count {
            background: var(--count-bg);
            color: var(--count-text);
            padding: 3px 9px;
            border-radius: 999px;
            font-size: 12px;
            font-weight: 600;
            flex-shrink: 0;
        }
        .tk-actions {
            padding: 10px 10px;
            display: flex;
            flex-wrap: nowrap;
            gap: 5px;
            border-bottom: 1px solid var(--border);
            background: var(--bg-header);
        }
        .tk-btn {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 4px;
            padding: 7px 8px;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            font-weight: 600;
            font-size: 12px;
            transition: opacity 0.15s, transform 0.1s;
            white-space: nowrap;
            line-height: 1.2;
            flex: 1 1 auto;
            min-width: 0;
        }
        .tk-btn:hover { opacity: 0.9; }
        .tk-btn:active { transform: scale(0.97); }
        .tk-btn:disabled { opacity: 0.4; cursor: not-allowed; }
        .tk-btn svg { width: 14px; height: 14px; flex-shrink: 0; }
        .tk-btn-primary { background: #22c55e; color: #052e16; }
        .tk-btn-copy { background: #eab308; color: #422006; }
        .tk-btn-fiszka { background: #a855f7; color: #faf5ff; }
        .tk-btn-danger { background: #64748b; color: #f8fafc; }
        .tk-list {
            flex: 1;
            overflow-y: auto;
            padding: 12px;
            background: var(--bg-list);
        }
        .tk-list::-webkit-scrollbar { width: 6px; }
        .tk-list::-webkit-scrollbar-track { background: transparent; }
        .tk-list::-webkit-scrollbar-thumb { background: var(--scrollbar); border-radius: 3px; }
        .tk-item {
            background: var(--bg-item);
            border-radius: 12px;
            padding: 12px 14px;
            margin-bottom: 10px;
            position: relative;
            border: 1px solid var(--border-item);
            transition: background 0.35s ease, border-color 0.35s ease, box-shadow 0.35s ease;
        }
        .tk-item.tk-item-new {
            border-color: #22c55e;
            box-shadow: 0 0 0 2px rgba(34, 197, 94, 0.4);
            background: color-mix(in srgb, #22c55e 12%, var(--bg-item, var(--bg)));
        }
        .tk-item.tk-item-dup {
            border-color: #f59e0b;
            box-shadow: 0 0 0 2px rgba(245, 158, 11, 0.45);
            background: color-mix(in srgb, #f59e0b 12%, var(--bg-item, var(--bg)));
            animation: tk-dup-pulse 0.9s ease 2;
        }
        @keyframes tk-dup-pulse {
            0%, 100% { box-shadow: 0 0 0 2px rgba(245, 158, 11, 0.35); }
            50% { box-shadow: 0 0 0 5px rgba(245, 158, 11, 0.2); }
        }
        .tk-item-top {
            display: flex;
            align-items: center;
            justify-content: flex-start;
            flex-wrap: wrap;
            gap: 6px;
            margin-top: 2px;
            margin-bottom: 0;
            padding-right: 28px;
        }
        .tk-item-source {
            display: inline-block;
            font-size: 10px;
            font-weight: 700;
            padding: 2px 7px;
            border-radius: 5px;
            letter-spacing: 0.3px;
            text-transform: uppercase;
        }
        .tk-source-taniaksiazka { background: #A94F2B; color: #FFF5EF; }
        .tk-source-platon { background: #315B86; color: #EFF6FF; }
        .tk-source-bonito { background: #E87524; color: #FFF7ED; }
        .tk-source-tantis { background: #16877D; color: #ECFDF8; }
        .tk-source-swiatksiazki { background: #B42332; color: #FFF1F2; }
        .tk-source-dressler { background: #88702F; color: #FFFAE8; }
        .tk-source-empik { background: #292929; color: #F4D27A; }
        .tk-source-ateneum { background: #1B4F72; color: #EAF2F8; }
        .tk-item-title {
            font-weight: 650;
            color: var(--text-title);
            margin: 0;
            line-height: 1.35;
            font-size: 14px;
            display: block;
            max-width: 100%;
        }
        .tk-item-title-text {
            display: inline;
            word-break: break-word;
        }
        /* Ołówek w tym samym spanie co tytuł – na końcu tekstu */
        .tk-item-title-text .tk-edit-btn {
            display: inline-flex;
            vertical-align: text-bottom;
            margin-left: 3px;
            position: relative;
            top: 1px;
        }
        .tk-item-remove {
            position: absolute;
            top: 6px;
            right: 6px;
            width: 28px;
            height: 28px;
            display: flex;
            align-items: center;
            justify-content: center;
            background: transparent;
            border: none;
            color: var(--text-label);
            cursor: pointer;
            font-size: 26px;
            font-weight: 500;
            line-height: 1;
            padding: 0;
            border-radius: 6px;
            z-index: 20;
            transition: color 0.15s, background 0.15s;
        }
        .tk-item-remove:hover {
            background: var(--remove-hover-bg);
            color: #f43f5e;
        }
        .tk-item-row {
            display: flex;
            align-items: flex-start;
            gap: 8px;
            margin: 5px 0;
            font-size: 12.5px;
            line-height: 1.35;
        }
        .tk-item-icon {
            width: 15px;
            height: 15px;
            flex-shrink: 0;
            color: var(--icon-color);
            margin-top: 1px;
        }
        .tk-item-label {
            color: var(--text-label);
            min-width: 88px;
            flex-shrink: 0;
        }
        .tk-item-value-wrap {
            flex: 1;
            min-width: 0;
            display: flex;
            align-items: flex-start;
            gap: 4px;
        }
        .tk-item-value {
            color: var(--text-value);
            word-break: break-word;
            flex: 1;
            min-width: 0;
        }
        .tk-edit-btn {
            flex-shrink: 0;
            width: 20px;
            height: 20px;
            display: flex;
            align-items: center;
            justify-content: center;
            background: transparent;
            border: none;
            border-radius: 4px;
            color: var(--pencil);
            cursor: pointer;
            padding: 0;
            opacity: 0.75;
            transition: opacity 0.15s, background 0.15s, color 0.15s;
        }
        .tk-item-row:hover .tk-edit-btn,
        .tk-item-title:hover .tk-edit-btn { opacity: 0.95; }
        .tk-edit-btn:hover {
            opacity: 1 !important;
            background: var(--edit-bg);
            color: var(--pencil-hover);
        }
        .tk-edit-btn svg { width: 12px; height: 12px; }

        /* Podświetlenie całego wiersza przy hover na ołówku */
        .tk-item-row {
            border-radius: 6px;
            margin-left: -4px;
            margin-right: -4px;
            padding-left: 4px;
            padding-right: 4px;
            transition: background 0.12s ease;
        }
        .tk-item-row:has(.tk-edit-btn:hover),
        .tk-item-row:hover {
            background: var(--row-hover-bg);
        }
        .tk-item-title {
            border-radius: 6px;
            transition: background 0.12s ease;
        }
        .tk-item-title:has(.tk-edit-btn:hover) {
            background: var(--row-hover-bg);
        }
        .tk-edit-input {
            flex: 1 1 auto;
            min-width: 0;
            width: 100%;
            background: var(--edit-bg);
            border: 1px solid var(--edit-border);
            border-radius: 6px;
            color: var(--text);
            font-size: 12.5px;
            font-family: inherit;
            padding: 4px 8px;
            outline: none;
            line-height: 1.35;
            box-sizing: border-box;
        }
        .tk-edit-input:focus {
            box-shadow: 0 0 0 2px rgba(59, 130, 246, 0.25);
        }
        /* Edycja tytułu – pełna szerokość kolumny tekstu (obok okładki) */
        .tk-item-title[data-editing] {
            display: flex !important;
            align-items: center !important;
            gap: 4px !important;
            width: 100% !important;
            max-width: 100% !important;
            margin: 0 !important;
        }
        .tk-item-title[data-editing] .tk-edit-input {
            flex: 1 1 auto !important;
            width: 100% !important;
            min-width: 0 !important;
            font-size: 13.5px !important;
            font-weight: 600 !important;
        }
        .tk-item-value-wrap[data-editing] {
            display: flex !important;
            align-items: center !important;
            gap: 4px !important;
            flex: 1 1 auto !important;
            min-width: 0 !important;
            width: 100% !important;
        }
        .tk-item-value-wrap[data-editing] .tk-edit-input {
            flex: 1 1 auto !important;
            min-width: 0 !important;
        }
        .tk-edit-actions {
            display: flex;
            gap: 2px;
            flex-shrink: 0;
        }
        .tk-edit-action {
            width: 22px;
            height: 22px;
            display: flex;
            align-items: center;
            justify-content: center;
            background: transparent;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            color: var(--edit-btn);
            padding: 0;
            font-size: 13px;
            line-height: 1;
        }
        .tk-edit-action:hover { background: var(--edit-bg); }
        .tk-edit-action.save:hover { color: #22c55e; }
        .tk-edit-action.cancel:hover { color: #f43f5e; }
        .tk-edit-action.clear:hover { color: #f59e0b; }
        .tk-empty {
            text-align: center;
            color: var(--empty);
            padding: 48px 20px;
            font-size: 13px;
            line-height: 1.5;
        }
        .tk-toast {
            position: fixed !important;
            bottom: 24px !important;
            right: 450px !important;
            background: #22c55e !important;
            color: #052e16 !important;
            padding: 11px 16px !important;
            border-radius: 10px !important;
            font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important;
            font-weight: 600 !important;
            font-size: 13px !important;
            font-style: normal !important;
            line-height: 1.35 !important;
            letter-spacing: normal !important;
            text-transform: none !important;
            text-shadow: none !important;
            z-index: 2147483002 !important;
            box-shadow: 0 6px 20px rgba(0,0,0,0.25) !important;
            animation: tk-fadein 0.25s ease;
        }
        @keyframes tk-fadein {
            from { opacity: 0; transform: translateY(8px); }
            to { opacity: 1; transform: translateY(0); }
        }
        body:has(#tk-panel.hidden) .tk-toast { right: 20px; }

        /* Search */
        .tk-search-wrap {
            padding: 8px 12px;
            background: var(--bg-header);
            border-bottom: 1px solid var(--border);
        }
        .tk-search {
            width: 100%;
            box-sizing: border-box;
            background: var(--bg-item);
            border: 1px solid var(--border-item);
            border-radius: 8px;
            color: var(--text);
            font-size: 12.5px;
            font-family: inherit;
            padding: 8px 12px 8px 34px;
            outline: none;
        }
        .tk-search:focus {
            border-color: var(--accent);
            box-shadow: 0 0 0 2px rgba(59, 130, 246, 0.2);
        }
        .tk-search::placeholder { color: var(--text-label); }
        .tk-search-box {
            position: relative;
        }
        .tk-search-icon {
            position: absolute;
            left: 10px;
            top: 50%;
            transform: translateY(-50%);
            width: 15px;
            height: 15px;
            color: var(--icon-color);
            pointer-events: none;
        }

        /* Header: cover + badge + title */
        #tk-panel .tk-item-header {
            display: flex !important;
            flex-direction: row !important;
            gap: 10px !important;
            align-items: flex-start !important;
            margin-bottom: 10px !important;
            padding-right: 28px !important;
            width: 100% !important;
        }
        #tk-panel .tk-cover {
            width: 48px !important;
            height: 66px !important;
            max-width: 48px !important;
            max-height: 66px !important;
            flex: 0 0 48px !important;
            border-radius: 6px !important;
            object-fit: cover !important;
            background: var(--border-item) !important;
            border: 1px solid var(--border-item) !important;
            display: block !important;
        }
        #tk-panel .tk-cover-placeholder {
            width: 48px !important;
            height: 66px !important;
            flex: 0 0 48px !important;
            border-radius: 6px !important;
            border: 1px solid var(--border-item) !important;
            background: var(--bg-list) !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            color: var(--icon-color) !important;
        }
        #tk-panel .tk-cover-placeholder svg {
            width: 20px !important;
            height: 20px !important;
            opacity: 0.5;
        }
        #tk-panel .tk-item-header-text {
            flex: 1 1 auto !important;
            min-width: 0 !important;
            display: flex !important;
            flex-direction: column !important;
            gap: 4px !important;
        }
        #tk-panel .tk-item-source {
            display: inline-block !important;
            font-size: 10px !important;
            font-weight: 700 !important;
            padding: 2px 7px !important;
            border-radius: 5px !important;
            letter-spacing: 0.3px !important;
            text-transform: uppercase !important;
            writing-mode: horizontal-tb !important;
            white-space: nowrap !important;
            line-height: 1.3 !important;
        }

        /* Source badge + external link */
        .tk-ext-link {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 20px;
            height: 20px;
            border-radius: 4px;
            color: var(--icon-color);
            text-decoration: none;
            transition: color 0.15s, background 0.15s;
        }
        .tk-ext-link:hover {
            color: var(--accent);
            background: var(--edit-bg);
        }
        .tk-ext-link svg {
            width: 13px;
            height: 13px;
        }
        .tk-search-box.active .tk-search {
            border-color: var(--accent);
            background: color-mix(in srgb, var(--accent) 12%, var(--bg-item));
            box-shadow: 0 0 0 2px rgba(59, 130, 246, 0.3);
            font-weight: 600;
        }
        .tk-search-box.active .tk-search-icon { color: var(--accent); }
        .tk-filter-info { display: none; }
        .tk-filter-info.on {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 8px;
            margin-top: 8px;
            padding: 7px 10px;
            border-radius: 8px;
            border: 1px solid var(--accent);
            background: color-mix(in srgb, var(--accent) 14%, var(--bg-item));
            color: var(--text-title);
            font-size: 12px;
            font-weight: 600;
            line-height: 1.35;
        }
        .tk-filter-text { min-width: 0; word-break: break-word; }
        .tk-filter-note {
            display: block;
            margin-top: 2px;
            font-size: 11px;
            font-weight: 400;
            color: var(--text-muted);
        }
        .tk-filter-clear {
            flex-shrink: 0;
            padding: 4px 9px;
            border: none;
            border-radius: 6px;
            background: var(--accent);
            color: #fff;
            font-size: 11.5px;
            font-weight: 700;
            cursor: pointer;
            white-space: nowrap;
        }
        .tk-filter-clear:hover { opacity: 0.88; }
        .tk-count.filtered { background: var(--accent); color: #fff; }

        .tk-btn-grid { background: #0ea5e9; color: #f0f9ff; }
        .tk-help-pill-grid { background: #0ea5e9; color: #f0f9ff; }

        /* ===== Siatka produktów (modal 95%) ===== */
        .tk-grid-modal {
            position: fixed;
            inset: 0;
            z-index: 2147483001;
            display: flex;
            align-items: center;
            justify-content: center;
            background: rgba(0, 0, 0, 0.6);
            pointer-events: auto;
            box-sizing: border-box;
            font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
            font-size: 13px;
            line-height: 1.4;
            color: var(--text);
            animation: tk-fadein 0.2s ease;
        }
        .tk-grid-modal *, .tk-grid-modal *::before, .tk-grid-modal *::after { box-sizing: border-box; }
        .tk-grid-modal button, .tk-grid-modal input { font-family: inherit; }
        .tk-grid-dialog {
            width: 95vw;
            height: 95vh;
            display: flex;
            flex-direction: column;
            overflow: hidden;
            background: var(--bg);
            color: var(--text);
            border: 1px solid var(--border);
            border-radius: 14px;
            box-shadow: 0 16px 48px rgba(0, 0, 0, 0.4);
        }
        .tk-grid-top {
            display: flex;
            align-items: center;
            flex-wrap: wrap;
            gap: 10px 16px;
            padding: 12px 16px;
            background: var(--bg-header);
            border-bottom: 1px solid var(--border);
        }
        .tk-grid-title { display: flex; align-items: baseline; gap: 10px; }
        .tk-grid-title h3 { margin: 0; font-size: 16px; font-weight: 650; color: var(--text-title); }
        .tk-grid-count { color: var(--text-muted); font-size: 12px; }
        .tk-grid-search-box { position: relative; flex: 1 1 260px; max-width: 420px; }
        .tk-grid-toggle {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            font-size: 12.5px;
            font-weight: 600;
            color: var(--text-value);
            cursor: pointer;
            user-select: none;
        }
        .tk-grid-close {
            margin-left: auto;
            width: 32px;
            height: 32px;
            padding: 0;
            border: none;
            border-radius: 8px;
            background: transparent;
            color: var(--text-muted);
            font-size: 24px;
            line-height: 1;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .tk-grid-close:hover { background: var(--edit-bg); color: var(--text-title); }
        .tk-grid-summary {
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            gap: 6px;
            padding: 8px 16px;
            background: var(--bg-header);
            border-bottom: 1px solid var(--border);
            color: var(--text-muted);
            font-size: 12px;
        }
        .tk-gchip {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 999px;
            border: 1px solid rgba(239, 68, 68, 0.4);
            background: rgba(239, 68, 68, 0.14);
            color: #dc2626;
            font-size: 11.5px;
            font-weight: 600;
        }
        .tk-grid-modal.dark .tk-gchip { color: #fca5a5; }
        .tk-gchip.ok {
            border-color: rgba(34, 197, 94, 0.4);
            background: rgba(34, 197, 94, 0.15);
            color: #16a34a;
        }
        .tk-grid-modal.dark .tk-gchip.ok { color: #86efac; }
        .tk-grid-body {
            flex: 1;
            overflow: auto;
            display: grid;
            /* max 5 kolumn; na węższych ekranach mniej (minmax 260px) */
            grid-template-columns: repeat(auto-fill, minmax(max(260px, calc((100% - 4 * 12px) / 5)), 1fr));
            align-content: start;
            gap: 12px;
            padding: 14px;
            background: var(--bg-list);
        }
        .tk-gcard {
            position: relative;
            padding: 12px;
            background: var(--bg-item);
            border: 1px solid var(--border-item);
            border-radius: 12px;
        }
        .tk-gcard.tk-gcard-incomplete { border-color: rgba(239, 68, 68, 0.55); }
        .tk-gcard-head { display: flex; gap: 10px; margin-bottom: 8px; padding-right: 26px; }
        .tk-gcover {
            display: block;
            flex: 0 0 64px;
            width: 64px;
            height: 90px;
            object-fit: cover;
            border: 1px solid var(--border-item);
            border-radius: 6px;
            background: var(--bg-list);
            cursor: zoom-in;
        }
        .tk-gcover-ph {
            flex: 0 0 64px;
            width: 64px;
            height: 90px;
            display: flex;
            align-items: center;
            justify-content: center;
            border: 1px solid var(--border-item);
            border-radius: 6px;
            background: var(--bg-list);
            color: var(--icon-color);
        }
        .tk-gcover-ph svg { width: 22px; height: 22px; opacity: 0.5; }
        .tk-gcard-headtext { flex: 1; min-width: 0; display: flex; flex-direction: column; gap: 5px; }
        .tk-gbadge { padding: 2px 7px; border-radius: 5px; font-size: 10.5px; font-weight: 700; white-space: nowrap; }
        .tk-gbadge.bad { background: #ef4444; color: #fff; }
        .tk-gbadge.ok { background: #22c55e; color: #052e16; }
        .tk-gcard .tk-item-row { margin: 3px -4px; font-size: 12px; }
        .tk-gcard .tk-item-label { min-width: 78px; }
        .tk-gcard .tk-item-row.tk-missing {
            background: rgba(239, 68, 68, 0.1);
        }
        .tk-dash { opacity: 0.4; }
        .tk-missing .tk-dash { opacity: 1; color: #ef4444; font-weight: 700; }
        .tk-grid-empty { grid-column: 1 / -1; padding: 48px 20px; text-align: center; color: var(--empty); }

        /* ===== Lightbox (podgląd okładki) ===== */
        #tk-panel .tk-cover[src] { cursor: zoom-in !important; }
        #tk-panel .tk-cover[src]:hover { opacity: 0.85; }

        .tk-lightbox {
            position: fixed;
            inset: 0;
            z-index: 2147483002;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            gap: 12px;
            padding: 24px;
            box-sizing: border-box;
            background: rgba(0, 0, 0, 0.82);
            pointer-events: auto;
            cursor: zoom-out;
            animation: tk-fadein 0.2s ease;
            font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
        }
        .tk-lightbox img {
            max-width: 90vw;
            max-height: 82vh;
            width: auto;
            height: auto;
            object-fit: contain;
            background: #fff;
            border-radius: 8px;
            box-shadow: 0 10px 40px rgba(0, 0, 0, 0.5);
            cursor: default;
        }
        .tk-lightbox-caption {
            color: #f1f5f9;
            font-size: 14px;
            font-weight: 600;
            text-align: center;
            max-width: 80vw;
        }
        .tk-lightbox-close {
            position: absolute;
            top: 16px;
            right: 20px;
            width: 40px;
            height: 40px;
            border: none;
            border-radius: 50%;
            background: rgba(255, 255, 255, 0.15);
            color: #fff;
            font-size: 28px;
            line-height: 1;
            cursor: pointer;
            transition: background 0.15s;
        }
        .tk-lightbox-close:hover { background: rgba(255, 255, 255, 0.3); }

        /* ===== Help button + modal ===== */
        .tk-help-btn {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 22px;
            height: 22px;
            padding: 0;
            margin: 0;
            border: none;
            border-radius: 50%;
            background: transparent;
            color: var(--text-muted);
            cursor: pointer;
            flex-shrink: 0;
            transition: color 0.15s, background 0.15s;
        }
        .tk-help-btn:hover {
            color: var(--accent);
            background: var(--edit-bg);
        }
        .tk-help-btn svg {
            width: 16px;
            height: 16px;
        }
        .tk-help-modal {
            position: fixed;
            inset: 0;
            z-index: 2147483003;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 24px;
            box-sizing: border-box;
            background: rgba(0, 0, 0, 0.55);
            pointer-events: auto;
            animation: tk-fadein 0.2s ease;
            font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        }
        .tk-help-dialog {
            background: var(--bg);
            color: var(--text);
            border: 1px solid var(--border);
            border-radius: 14px;
            box-shadow: 0 16px 48px rgba(0, 0, 0, 0.3);
            width: min(720px, 96vw);
            max-height: 85vh;
            overflow: auto;
            padding: 20px 22px 18px;
            position: relative;
        }
        .tk-help-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px 28px;
            align-items: start;
        }
        @media (max-width: 640px) {
            .tk-help-grid { grid-template-columns: 1fr; }
        }
        .tk-help-col-about {
            border-left: 1px solid var(--border);
            padding-left: 22px;
        }
        @media (max-width: 640px) {
            .tk-help-col-about {
                border-left: none;
                padding-left: 0;
                border-top: 1px solid var(--border);
                padding-top: 16px;
            }
        }
        .tk-help-about-list {
            list-style: none;
            margin: 0;
            padding: 0;
            display: flex;
            flex-direction: column;
            gap: 8px;
        }
        .tk-help-about-list li {
            font-size: 12.5px;
            line-height: 1.45;
            color: var(--text);
            padding-left: 14px;
            position: relative;
        }
        .tk-help-about-list li::before {
            content: '•';
            position: absolute;
            left: 0;
            color: var(--accent, #3b82f6);
            font-weight: 700;
        }
        .tk-help-version {
            margin-top: 14px;
            font-size: 11px;
            color: var(--text-muted);
        }
        .tk-help-dialog h3 {
            margin: 0 28px 10px 0;
            font-size: 15px;
            font-weight: 650;
            color: var(--text-title);
        }
        .tk-help-dialog h3.tk-help-section {
            margin-top: 18px;
            margin-bottom: 8px;
            padding-top: 14px;
            border-top: 1px solid var(--border);
            font-size: 14px;
        }
        .tk-help-dialog p {
            margin: 0 0 12px;
            font-size: 12.5px;
            color: var(--text-muted);
            line-height: 1.45;
        }
        .tk-help-list {
            list-style: none;
            margin: 0;
            padding: 0;
            display: flex;
            flex-direction: column;
            gap: 8px;
        }
        .tk-help-list li {
            display: flex;
            align-items: flex-start;
            gap: 10px;
            font-size: 13px;
            font-weight: 600;
            color: var(--text-title);
        }
        .tk-help-list .tk-item-source {
            flex-shrink: 0;
        }
        .tk-help-actions {
            list-style: none;
            margin: 0;
            padding: 0;
            display: flex;
            flex-direction: column;
            gap: 10px;
        }
        .tk-help-actions li {
            display: flex;
            flex-direction: column;
            gap: 2px;
        }
        .tk-help-actions .tk-help-btn-name {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            font-size: 12.5px;
            font-weight: 650;
            color: var(--text-title);
        }
        .tk-help-actions .tk-help-btn-name .tk-help-pill {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 6px;
            font-size: 11px;
            font-weight: 700;
            letter-spacing: 0.2px;
        }
        .tk-help-pill-load { background: #22c55e; color: #052e16; }
        .tk-help-pill-copy { background: #eab308; color: #422006; }
        .tk-help-pill-fiszka { background: #a855f7; color: #faf5ff; }
        .tk-help-pill-clear { background: #64748b; color: #f8fafc; }
        .tk-help-actions .tk-help-btn-desc {
            font-size: 12px;
            font-weight: 400;
            color: var(--text-muted);
            line-height: 1.4;
            padding-left: 2px;
        }
        .tk-help-close {
            position: absolute;
            top: 12px;
            right: 12px;
            width: 28px;
            height: 28px;
            border: none;
            border-radius: 6px;
            background: transparent;
            color: var(--text-muted);
            font-size: 20px;
            line-height: 1;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 0;
        }
        .tk-help-close:hover {
            background: var(--edit-bg);
            color: var(--text-title);
        }
    `;

    // ========== STAN ==========
    let products = GM_getValue('tk_products', []);
    let panelVisible = GM_getValue('tk_panel_visible', true);
    let darkTheme = GM_getValue('tk_dark_theme', false);
    let filterQuery = GM_getValue('tk_filter', '');
    let filterPersistTimer = null;
    let uiRoot = null; // ShadowRoot – izolacja CSS od strony

    function ui(sel) { return uiRoot ? uiRoot.querySelector(sel) : null; }
    function uiAll(sel) { return uiRoot ? uiRoot.querySelectorAll(sel) : []; }


    const FISZKA_FIELDS = [
        'PEŁNY TYTUŁ I PODTYTUŁ', 'DATA PREMIERY (DZIEŃ/MIESIĄC/ROK)', 'INDEKS GOLD',
        'CENA BRUTTO', 'VAT', 'CENA NETTO', 'CENA DRUKOWANA TAK/NIE',
        'RODZAJ PRODUKTU (KSIĄŻKA / AUDIOBOOK)', 'ROK WYDANIA', 'JĘZYK WYDANIA',
        'JĘZYK ORYGINAŁU', 'GRUPA TEMATYCZNA', 'PRZEDZIAŁ WIEKOWY',
        'WERSJA ELEKTRONICZNA (EBOOK) TAK/NIE', 'WERSJA AUDIO TAK/NIE',
        'LEKTOR (W PRZYPADKU AUDIOBOOKA)', 'AUTOR I', 'AUTOR II', 'AUTOR III',
        'WYDAWCA', 'TŁUMACZ I', 'TŁUMACZ II', 'TYTUŁ ORYGINAŁU', 'ISBN', 'EAN',
        'NUMER WYDANIA', 'NAZWA CYKLU/SERII', 'OPRAWA (MK/TW)', 'ILOŚĆ STRON', 'OPIS'
    ];

    const ICONS = {
        box: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/></svg>`,
        download: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>`,
        copy: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>`,
        file: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>`,
        trash: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>`,
        sun: `<svg class="tk-theme-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></svg>`,
        pencil: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 3a2.85 2.85 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z"/><path d="m15 5 4 4"/></svg>`,
        search: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>`,
        external: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>`,
        image: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/></svg>`,
        user: `<svg class="tk-item-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>`,
        tag: `<svg class="tk-item-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20.59 13.41l-7.17 7.17a2 2 0 0 1-2.83 0L2 12V2h10l8.59 8.59a2 2 0 0 1 0 2.82z"/><line x1="7" y1="7" x2="7.01" y2="7"/></svg>`,
        money: `<svg class="tk-item-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></svg>`,
        percent: `<svg class="tk-item-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="19" y1="5" x2="5" y2="19"/><circle cx="6.5" cy="6.5" r="2.5"/><circle cx="17.5" cy="17.5" r="2.5"/></svg>`,
        calendar: `<svg class="tk-item-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>`,
        barcode: `<svg class="tk-item-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 5v14M8 5v14M12 5v14M17 5v14M21 5v14"/></svg>`,
        book: `<svg class="tk-item-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/></svg>`,
        layers: `<svg class="tk-item-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/></svg>`,
        maximize: `<svg class="tk-item-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M8 3H5a2 2 0 0 0-2 2v3m18 0V5a2 2 0 0 0-2-2h-3m0 18h3a2 2 0 0 0 2-2v-3M3 16v3a2 2 0 0 0 2 2h3"/></svg>`,
        folder: `<svg class="tk-item-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>`,
        grid: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>`,
        help: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>`,
        lang: `<svg class="tk-item-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/></svg>`
    };

    function toast(msg, duration = 2500) {
        const root = uiRoot || document;
        const existing = (uiRoot ? uiRoot.querySelector('.tk-toast') : document.querySelector('.tk-toast'));
        if (existing) existing.remove();
        const el = document.createElement('div');
        el.className = 'tk-toast';
        el.textContent = msg;
        if (uiRoot) uiRoot.appendChild(el);
        else document.body.appendChild(el);
        setTimeout(() => el.remove(), duration);
    }

    /** Przy wielu seriach (przecinek / średnik / „ / ”) zostaw tylko pierwszą. */
    function normalizeSeries(raw) {
        if (!raw) return '';
        let s = String(raw).replace(/\s+/g, ' ').trim();
        if (!s) return '';
        const parts = s.split(/\s*[;|]\s*|\s+\/\s+|\s*,\s+/)
            .map(p => p.trim())
            .filter(Boolean);
        if (parts.length <= 1) return s;
        for (const p of parts) {
            if (p.length >= 2 && !/^seri[ae]$/i.test(p)) return p;
        }
        return parts[0];
    }

    function normalizeBinding(raw) {
        if (!raw) return '';
        const val = String(raw).toLowerCase().replace(/\s+/g, ' ').trim();
        if (val.includes('tward') || val.includes('hardcover') || val.includes('półtward') || val.includes('poltward')) return 'Twarda';
        if (val.includes('miękk') || val.includes('miekk') || val.includes('broszurow') || val.includes('broszur') || val.includes('soft') || val.includes('paperback')) return 'Miekka';
        return raw.trim();
    }

/** Gdy podano kilka języków („polski, angielski”), zostaje tylko pierwszy. */
    function normalizeLanguage(raw) {
        if (!raw) return '';
        const first = String(raw).split(/\s*(?:,|;|\/|\s+i\s+|\s+oraz\s+)\s*/i)[0];
        return first.replace(/\s+/g, ' ').trim();
    }

    function bindingToMK_TW(binding) {
        const b = normalizeBinding(binding);
        if (b === 'Twarda') return 'TW';
        if (b === 'Miekka') return 'MK';
        return '';
    }

    /** Zwraca {d,m,y} lub null z różnych formatów daty premiery */
    function parsePremiereParts(raw) {
        if (!raw) return null;
        const s = String(raw).trim();
        let m;
        // YYYY-MM-DD or YYYY.MM.DD or YYYY/MM/DD
        m = s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);
        if (m) return { y: m[1], m: m[2].padStart(2, '0'), d: m[3].padStart(2, '0') };
        // DD.MM.YYYY or DD/MM/YYYY or DD-MM-YYYY
        m = s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})/);
        if (m) return { d: m[1].padStart(2, '0'), m: m[2].padStart(2, '0'), y: m[3] };
        // ISO with time
        m = s.match(/^(\d{4})-(\d{2})-(\d{2})T/);
        if (m) return { y: m[1], m: m[2], d: m[3] };
        return null;
    }

    function formatPremiereDisplay(raw) {
        const p = parsePremiereParts(raw);
        if (!p) return raw || '';
        return `${p.d}.${p.m}.${p.y}`;
    }

    function formatPremiereFiszka(raw) {
        const p = parsePremiereParts(raw);
        if (!p) return raw || '';
        return `${p.d}/${p.m}/${p.y}`;
    }

    /**
     * Rok wydania: tylko z ostatnich 100 lat (i max +10 lat w przyszłość).
     * Odrzuca absurdy typu 1900, 0, 9999.
     */
    function sanitizeYear(y) {
        if (y == null || y === '') return '';
        const m = String(y).match(/(19|20)\d{2}/);
        if (!m) return '';
        const n = parseInt(m[0], 10);
        const now = new Date().getFullYear();
        if (n < now - 100 || n > now + 10) return '';
        return String(n);
    }

    function escapeHtml(str) {
        return (str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    /** Szuka jawnej stawki VAT na stronie (np. „VAT 23%”, „(+ 23% VAT)”). */
    function detectVatFromPage() {
        const vatBox = document.querySelector('.c-item-vatr, [class*="vatr"]');
        if (vatBox) {
            const mBox = vatBox.textContent.match(/(\d{1,2})\s*%/);
            if (mBox) {
                const v = parseInt(mBox[1], 10);
                if ([0, 5, 8, 23].includes(v)) return String(v);
            }
        }
        const text = document.body ? document.body.innerText : '';
        const m = text.match(/\(\+\s*(\d{1,2})\s*%\s*VAT\)/i) ||
                  text.match(/\bVAT\s*[:=]?\s*(\d{1,2})\s*%/i) ||
                  text.match(/stawka\s+(?:podatku\s+)?VAT\s*[:=]?\s*(\d{1,2})/i) ||
                  text.match(/(\d{1,2})\s*%\s*VAT/i);
        if (m) {
            const v = parseInt(m[1], 10);
            if ([0, 5, 8, 23].includes(v)) return String(v);
        }
        return '';
    }

    /**
     * Czy produkt wygląda na książkę / publikację (VAT 5%),
     * a nie zabawkę, grę, gadżet itd.
     */
    function isLikelyBookProduct(ctx = {}) {
        const cats = String(ctx.categories || '');
        // ścieżka z „Książki” / „Literatura” = książka (nawet bez ceny)
        if (/książk|ksiazka|komiks|literatura|e-?book|audiobook|lektur/i.test(cats)) return true;

        const hay = [
            cats, ctx.title, ctx.url, ctx.form, ctx.binding,
            typeof location !== 'undefined' ? location.pathname : ''
        ].join(' ').toLowerCase();

        const nonBookRe = /zabawk|puzzle(?!\s*book)|figurk|pluszak|klock|lego\b|gry\s+plansz|gra\s+plansz|gry\s+komputer|konsol|elektronik|odzież|ubrani|koszulk|kubek|puzzle\s*3d|akcesori/i;
        const bookRe = /książk|ksiazka|komiks|e-?book|audiobook|literatura|powieść|powiesc|poradnik|leksykon|słownik|slownik|podręcznik|podrecznik|baśni|basni|album\s+ilustrow/i;

        if (nonBookRe.test(hay) && !bookRe.test(hay)) return false;
        if (bookRe.test(hay)) return true;
        if (/ksiazka-p|,ksiazka|\/ksiazki|\/book|ebooki/i.test(ctx.url || (typeof location !== 'undefined' ? location.pathname : ''))) return true;
        if (ctx.form && /^książka$|^ksiazka$/i.test(String(ctx.form).trim())) return true;
        if (ctx.pages && String(ctx.pages).match(/\d+/) && ctx.binding) return true;
        return false;
    }

    /** brutto + opcjonalna stawka → { cenaBrutto, cenaNetto, vat }
     *  VAT może być ustawiony także bez ceny (np. książka z ścieżki). */
    function priceWithVat(brutto, vatRate) {
        const vatStr = (vatRate !== '' && vatRate != null) ? String(vatRate) : '';
        if (brutto == null || isNaN(brutto) || brutto <= 0) {
            return { cenaBrutto: '', cenaNetto: '', vat: vatStr };
        }
        const cenaBrutto = Number(brutto).toFixed(2);
        if (!vatStr) return { cenaBrutto, cenaNetto: '', vat: '' };
        const rate = 1 + (parseInt(vatStr, 10) / 100);
        if (!rate || rate <= 1) return { cenaBrutto, cenaNetto: '', vat: vatStr };
        return { cenaBrutto, vat: vatStr, cenaNetto: (brutto / rate).toFixed(2) };
    }

    /**
     * Ustal VAT: 1) jawna stawka ze strony, 2) 5% gdy ścieżka/heurystyka = książka, 3) puste.
     * VAT zapisujemy nawet bez ceny brutto.
     */
    function resolvePriceAndVat(brutto, ctx = {}) {
        let vat = detectVatFromPage();
        if (!vat && /książk|ksiazka|komiks|literatura|e-?book|audiobook|lektur/i.test(String(ctx.categories || ''))) {
            vat = '5';
        }
        if (!vat && isLikelyBookProduct(ctx)) vat = '5';
        return priceWithVat(brutto, vat);
    }

    function sourceLabel(source) {
        const map = {
            platon: 'Platon',
            taniaksiazka: 'TaniaKsiążka',
            bonito: 'Bonito',
            tantis: 'Tantis',
            swiatksiazki: 'Świat Książki',
            dressler: 'Dressler',
            empik: 'Empik',
            ateneum: 'Ateneum'
        };
        return map[source] || source || '';
    }

    // ========== ADAPTER: TANIAKSIĄŻKA ==========
    function isTaniaKsiazkaProductPage() {
        return !!document.querySelector('#DetailsModule') ||
               !!document.querySelector('#accordion-content__szczegóły') ||
               !!document.querySelector('li.display-detail') ||
               !!document.querySelector('h1.product-info-title');
    }

    function getTaniaDetail(label) {
        const items = document.querySelectorAll('#DetailsModule li.display-detail, .bullet li.display-detail, li.display-detail');
        for (const li of items) {
            const text = li.textContent.replace(/\s+/g, ' ').trim();
            if (text.toLowerCase().startsWith(label.toLowerCase() + ':') || text.toLowerCase().startsWith(label.toLowerCase() + ' ')) {
                const strong = li.querySelector('strong');
                if (strong) return strong.textContent.trim();
                const parts = text.split(':');
                if (parts.length > 1) return parts.slice(1).join(':').trim();
            }
        }
        return '';
    }

    function getTaniaAuthors() {
        const fromDetails = getTaniaDetail('Autorzy') || getTaniaDetail('Autor');
        if (fromDetails) return fromDetails.replace(/\s+/g, ' ').trim();
        const productAuthorRoot = document.querySelector('.product-info-author, .product-authors, .product-info__author, #DetailsModule');
        if (productAuthorRoot) {
            const links = productAuthorRoot.querySelectorAll('a');
            if (links.length > 0) {
                const names = Array.from(links).map(a => a.textContent.trim()).filter(Boolean);
                const unique = [...new Set(names)];
                if (unique.length) return unique.join(', ');
            }
            const text = productAuthorRoot.textContent.replace(/\s+/g, ' ').trim();
            const cleaned = text.replace(/^(autorzy?|autor)\s*[:：]?\s*/i, '').trim();
            if (cleaned) return cleaned;
        }
        return '';
    }

    function getTaniaYear() {
        const raw = getTaniaDetail('Rok wydania') || getTaniaDetail('Rok');
        return sanitizeYear(raw);
    }

    function getTaniaTitle() {
        const h1 = document.querySelector('h1.product-info-title');
        if (h1) return h1.textContent.trim();
        return getTaniaDetail('Tytuł') || (document.querySelector('h1')?.textContent.trim() || '');
    }

    function getTaniaBreadcrumbs() {
        const path = document.querySelector('ul#path.breadcrumb__path, ul.breadcrumb__path, ul#path');
        if (!path) return '';
        const items = path.querySelectorAll('li');
        const parts = [];
        items.forEach((li, index) => {
            if (li.classList.contains('home')) return;
            const a = li.querySelector('a');
            let text = a ? a.textContent.trim() : li.textContent.trim();
            text = text.replace(/\s+/g, ' ').trim();
            if (index === items.length - 1 && !a) return;
            if (text && text.toLowerCase() !== 'strona główna') parts.push(text);
        });
        return parts.join(' > ');
    }

    function getTaniaRetailPrice() {
        const amountEl = document.querySelector('.p-retail-info-amount');
        if (!amountEl) return null;
        const text = amountEl.textContent.replace(/\s+/g, ' ').trim();
        const match = text.match(/([\d\s]+[,.]?\d*)/);
        if (!match) return null;
        const numStr = match[1].replace(/\s/g, '').replace(',', '.');
        const price = parseFloat(numStr);
        return isNaN(price) ? null : price;
    }

    function isTaniaInKsiazkiSection() {
        const cats = getTaniaBreadcrumbs().toLowerCase();
        return cats.includes('książki') || cats.includes('ksiazki');
    }

    function extractTaniaKsiazka() {
        const title = getTaniaTitle();
        if (!title) { toast('Nie znaleziono tytułu produktu (TaniaKsiążka)'); return null; }
        const brutto = getTaniaRetailPrice();
        let vat = '', cenaBrutto = '', cenaNetto = '';
        if (brutto !== null && isTaniaInKsiazkiSection()) {
            vat = '5';
            cenaBrutto = brutto.toFixed(2);
            cenaNetto = (brutto / 1.05).toFixed(2);
        }
        return {
            id: Date.now() + Math.random().toString(36).slice(2, 7),
            source: 'taniaksiazka', title,
            author: getTaniaAuthors(),
            publisher: getTaniaDetail('Wydawnictwo'),
            series: getTaniaDetail('Seria') || (function() {
                const links = document.querySelectorAll('#DetailsModule a[href*="/seria/"], .details-section a[href*="/seria/"]');
                const names = Array.from(links).map(a => a.textContent.replace(/\s+/g, ' ').trim()).filter(Boolean);
                return [...new Set(names)].join(', ');
            })(),
            binding: normalizeBinding(getTaniaDetail('Oprawa')),
            year: getTaniaYear(),
            pages: getTaniaDetail('Ilość stron'),
            language: getTaniaDetail('Język'),
            originalTitle: getTaniaDetail('Oryginalny tytuł'),
            translator: getTaniaDetail('Tłumacz'),
            edition: getTaniaDetail('Nr wydania'),
            isbn: getTaniaDetail('ISBN') || getTaniaDetail('Model'),
            ean: getTaniaDetail('EAN') || getTaniaDetail('ISBN') || getTaniaDetail('Model'),
            categories: getTaniaBreadcrumbs(),
            premiereDate: (function() {
                const m = document.querySelector('meta[property="product:release_date"]');
                return m && m.content ? m.content.trim().slice(0, 10) : '';
            })(),
            cover: (function() {
                const img = document.querySelector('.product-gallery img, .product-image img, .gallery-image img, #product-image img, .p-image img, img.product-img');
                if (img && img.src) return img.src;
                const og = document.querySelector('meta[property="og:image"]');
                if (og && og.content) return og.content;
                return '';
            })(),
            vat, cenaBrutto, cenaNetto,
            url: location.href,
            addedAt: new Date().toISOString()
        };
    }

    // ========== ADAPTER: PLATON ==========
    function isPlatonProductPage() {
        return !!document.querySelector('h1') &&
               (!!document.querySelector('.c-item-authors') ||
                !!document.querySelector('script[type="application/ld+json"]') ||
                !!document.querySelector('.c-item-attr') ||
                !!document.querySelector('.c-item-info'));
    }

    function getPlatonJsonLd() {
        const scripts = document.querySelectorAll('script[type="application/ld+json"]');
        for (const s of scripts) {
            try {
                const data = JSON.parse(s.textContent);
                if (data['@type'] === 'Product') return data;
            } catch (e) {}
        }
        return null;
    }

    function getPlatonAttr(label) {
        const items = document.querySelectorAll('ul.c-item-attr li');
        for (const li of items) {
            const spans = li.querySelectorAll(':scope > span');
            if (spans.length >= 1) {
                const labelText = spans[0].textContent.trim().toLowerCase();
                if (labelText === label.toLowerCase()) {
                    const linkSpan = li.querySelector('a span');
                    if (linkSpan) return linkSpan.textContent.trim();
                    if (spans.length >= 2) return spans[1].textContent.trim();
                    const full = li.textContent.replace(/\s+/g, ' ').trim();
                    const parts = full.split(spans[0].textContent.trim());
                    if (parts.length > 1) return parts.slice(1).join('').trim();
                }
            }
        }
        return '';
    }

// ZASTĄP całą dotychczasową funkcję extractPlaton() tym kodem (wersja z pobieraniem serii).

    function extractPlaton() {
        const json = getPlatonJsonLd();
        let title = json?.name || document.querySelector('h1')?.textContent.trim() || '';
        if (!title) { toast('Nie znaleziono tytułu produktu (Platon)'); return null; }
        let author = '';
        if (json?.author) {
            const authors = Array.isArray(json.author) ? json.author : [json.author];
            author = authors.map(a => a.name).filter(Boolean).join(', ');
        }
        if (!author) {
            const authorEls = document.querySelectorAll('.c-item-authors a.c-item-author span, .c-item-authors .c-item-author span');
            author = Array.from(authorEls).map(el => el.textContent.trim()).filter(Boolean).join(', ');
        }
        let publisher = json?.brand?.name || getPlatonAttr('Wydawnictwo');

        // ===== NOWE: seria z listy atrybutów =====
        let series = getPlatonAttr('Seria') || getPlatonAttr('Serie') ||
                     getPlatonAttr('Seria wydawnicza') || getPlatonAttr('Cykl') || '';
        if (!series) {
            // fallback: link do serii w liście atrybutów
            const sl = document.querySelector('ul.c-item-attr a[href*="seri"] span, ul.c-item-attr a[href*="seri"]');
            if (sl) series = sl.textContent.replace(/\s+/g, ' ').trim();
        }
        series = normalizeSeries(series);
        // =========================================

        const ean = json?.sku || getPlatonAttr('Kod EAN') || '';
        const isbn = ean;
        const pages = getPlatonAttr('Strony');
        const bindingRaw = getPlatonAttr('Okładka');
        const premiere = getPlatonAttr('Data premiery');
        const year = premiere ? premiere.slice(0, 4) : '';
        let categories = '';
        const breadcrumbItems = document.querySelectorAll('.c-breadcrumb .breadcrumbs__item a span, .c-breadcrumb .breadcrumbs__item > span, .c-breadcrumbs a span');
        const parts = Array.from(breadcrumbItems).map(el => el.textContent.trim()).filter(t => t && t.toLowerCase() !== 'strona główna');
        if (parts.length) categories = [...new Set(parts)].join(' > ');
        // Ceny Platon (hurt): „Sugerowana cena detaliczna (netto)” + „(+ X% VAT)”
        let cenaNetto = '', cenaBrutto = '', vat = detectVatFromPage();
        let nettoNum = null;
        let bruttoNum = null;

        // 1) Sugerowana detaliczna NETTO (najpewniejsza)
        const priceBlocks = document.querySelectorAll('.c-item-price-cont, .c-item-prices-cont > div');
        for (const block of priceBlocks) {
            const lab = (block.querySelector('.c-item-price-label') || block).textContent || '';
            if (/sugerowana\s+cena\s+detaliczna\s*\(?\s*netto/i.test(lab) ||
                (/netto/i.test(lab) && /sugerowana|detaliczna/i.test(lab))) {
                const valEl = block.querySelector('.c-item-price .value, .value');
                if (valEl) {
                    const n = parseFloat(valEl.textContent.replace(/\s/g, '').replace(',', '.'));
                    if (!isNaN(n) && n > 0) { nettoNum = n; break; }
                }
            }
        }
        // 2) Fallback: tekst strony
        if (nettoNum == null) {
            const m = document.body.innerText.match(
                /Sugerowana\s+cena\s+detaliczna\s*\(?\s*netto\)?\s*:?\s*([\d\s]+[,.]\d{2})/i
            );
            if (m) {
                const n = parseFloat(m[1].replace(/\s/g, '').replace(',', '.'));
                if (!isNaN(n)) nettoNum = n;
            }
        }
        // 3) JSON-LD / stary selektor
        if (json?.offers?.price) {
            const price = parseFloat(json.offers.price);
            if (!isNaN(price)) bruttoNum = price;
        }
        if (bruttoNum == null) {
            const priceEl = document.querySelector('.c-item-price .value');
            if (priceEl && nettoNum == null) {
                const price = parseFloat(priceEl.textContent.replace(/\s/g, '').replace(',', '.'));
                if (!isNaN(price)) bruttoNum = price;
            }
        }

        // Ustal VAT: strona → książka 5% → zabawki/gry 23%
        if (!vat && isLikelyBookProduct({ categories, title, url: location.href, pages, binding: bindingRaw })) {
            vat = '5';
        }
        if (!vat && /zabawk|gry|gra\b/i.test(categories + ' ' + title)) {
            vat = '23';
        }

        if (nettoNum != null) {
            cenaNetto = nettoNum.toFixed(2);
            if (vat) {
                const rate = 1 + (parseInt(vat, 10) / 100);
                cenaBrutto = (nettoNum * rate).toFixed(2);
            } else if (bruttoNum != null) {
                cenaBrutto = bruttoNum.toFixed(2);
            }
        } else if (bruttoNum != null) {
            const resolved = priceWithVat(bruttoNum, vat);
            cenaBrutto = resolved.cenaBrutto;
            cenaNetto = resolved.cenaNetto;
            vat = resolved.vat || vat;
        }

        return {
            id: Date.now() + Math.random().toString(36).slice(2, 7),
            source: 'platon', title, author, publisher,
            series,                                   // <- było: series: ''
            binding: normalizeBinding(bindingRaw), year, pages,
            language: '', originalTitle: '', translator: '', edition: '',
            isbn, ean, categories, premiereDate: premiere,
            cover: (function() {
                if (json && json.image) {
                    const img = Array.isArray(json.image) ? json.image[0] : json.image;
                    if (img) return img;
                }
                const el = document.querySelector('.c-selected-photo img, .c-item-photo-cont img, .csB2BItemPhotos img');
                if (el && el.src) return el.src;
                const og = document.querySelector('meta[property="og:image"]');
                if (og && og.content) return og.content;
                return '';
            })(),
            vat, cenaBrutto, cenaNetto,
            url: location.href, addedAt: new Date().toISOString()
        };
    }


    // ========== ADAPTER: BONITO ==========
    function isBonitoProductPage() {
        return location.pathname.includes('/produkt/') ||
               !!document.querySelector('script[type="application/ld+json"]');
    }

    function getJsonLdProductAny() {
        const scripts = document.querySelectorAll('script[type="application/ld+json"]');
        for (const s of scripts) {
            try {
                const data = JSON.parse(s.textContent);
                if (data['@type'] === 'Product') return data;
                if (Array.isArray(data['@graph'])) {
                    const p = data['@graph'].find(x => x['@type'] === 'Product');
                    if (p) return p;
                }
                if (Array.isArray(data)) {
                    const p = data.find(x => x && x['@type'] === 'Product');
                    if (p) return p;
                }
            } catch (e) {}
        }
        return null;
    }

    function propsMap(additionalProperty) {
        const map = {};
        (additionalProperty || []).forEach(p => {
            if (p && p.name) map[p.name] = p.value || '';
        });
        return map;
    }

    /** Wszystkie wartości danej etykiety z additionalProperty (np. wielu Author) */
    function propsAll(additionalProperty, names) {
        const wanted = (Array.isArray(names) ? names : [names]).map(n => n.toLowerCase());
        const out = [];
        (additionalProperty || []).forEach(p => {
            if (!p || !p.name) return;
            if (!wanted.includes(String(p.name).toLowerCase())) return;
            const v = (p.value || '').toString().trim();
            if (v) out.push(v);
        });
        return out;
    }

    /**
     * Składa listę autorów:
     * - usuwa duplikaty
     * - jeśli jest „praca zbiorowa” ORAZ inni autorzy → pomiń „praca zbiorowa”
     * - jeśli tylko „praca zbiorowa” → zostaw
     */
    function normalizeAuthorsList(authors) {
        const list = (Array.isArray(authors) ? authors : String(authors || '').split(/[,;]/))
            .map(a => String(a).replace(/\s+/g, ' ').trim())
            .filter(Boolean);
        const unique = [];
        const seen = new Set();
        for (const a of list) {
            const key = a.toLowerCase();
            if (seen.has(key)) continue;
            seen.add(key);
            unique.push(a);
        }
        const isCollective = (a) => /^praca\s+zbiorowa$/i.test(a);
        const hasNamed = unique.some(a => !isCollective(a));
        const filtered = hasNamed ? unique.filter(a => !isCollective(a)) : unique;
        return filtered.join(', ');
    }

    function extractBonito() {
        const json = getJsonLdProductAny();
        if (!json || !json.name) { toast('Nie znaleziono produktu (Bonito)'); return null; }
        const props = propsMap(json.additionalProperty);
        const title = (json.name || '').trim();
        const author = (props['Autor'] || '').trim();
        const publisher = (props['Wydawnictwo'] || (json.brand && json.brand.name) || '').trim();
        const series = (props['Seria wydawnicza'] || '').trim();
        const year = (props['Rok wydania'] || '').trim();
        const binding = normalizeBinding(props['Oprawa'] || '');
        const pages = (props['Liczba stron'] || '').trim();
        const isbn = (props['Numer ISBN'] || json.gtin13 || json.isbn || '').toString().trim();
        const ean = (props['Kod paskowy (EAN)'] || json.gtin13 || isbn).toString().trim();
        const premiere = (props['Data premiery'] || '').trim();
        const categories = (json.category || '').toString().trim();
        let cover = '';
        if (json.image) cover = Array.isArray(json.image) ? json.image[0] : json.image;
        if (!cover) {
            const og = document.querySelector('meta[property="og:image"]');
            if (og && og.content) cover = og.content;
        }
        let bruttoNum = null;
        if (json.offers) {
            const off = Array.isArray(json.offers) ? json.offers[0] : json.offers;
            if (off && off.price != null) {
                const price = parseFloat(String(off.price).replace(',', '.'));
                if (!isNaN(price)) bruttoNum = price;
            }
        }
        const { cenaBrutto, cenaNetto, vat } = resolvePriceAndVat(bruttoNum, {
            categories, title, url: location.href, pages, binding
        });
        return {
            id: Date.now() + Math.random().toString(36).slice(2, 7),
            source: 'bonito', title, author, publisher, series,
            binding, year, pages, language: '',
            originalTitle: '', translator: '', edition: '',
            isbn, ean, categories, premiereDate: premiere, cover,
            vat, cenaBrutto, cenaNetto,
            url: location.href, addedAt: new Date().toISOString()
        };
    }

    // ========== ADAPTER: TANTIS ==========
    function isTantisProductPage() {
        return /\/p\d+/.test(location.pathname) ||
               /-p\d+/.test(location.pathname) ||
               !!document.querySelector('script[type="application/ld+json"]');
    }

    /** Czyta wartość z listy cech (ul.-features) po etykiecie, np. "Serie" */
    function getTantisFeature(labels) {
        const wanted = (Array.isArray(labels) ? labels : [labels]).map(l => l.toLowerCase());
        // Preferuj cechy w kontekście produktu, nie z całej strony
        const root = document.querySelector('#main-product, .product-main, .ui-product-main, main') || document;
        const items = root.querySelectorAll('ul.-features li');
        for (const li of items) {
            const labEl = li.querySelector(':scope > span');
            if (!labEl) continue;
            const lab = labEl.textContent.replace(/\s+/g, ' ').trim().replace(/:$/, '').toLowerCase();
            if (!wanted.includes(lab)) continue;
            const links = li.querySelectorAll('a');
            if (links.length) {
                return Array.from(links).map(a => a.textContent.replace(/\s+/g, ' ').trim()).filter(Boolean).join(', ');
            }
            const clone = li.cloneNode(true);
            const sp = clone.querySelector(':scope > span');
            if (sp) sp.remove();
            return clone.textContent.replace(/\s+/g, ' ').trim();
        }
        return '';
    }

    function extractTantis() {
        const json = getJsonLdProductAny();
        if (!json || !json.name) { toast('Nie znaleziono produktu (Tantis)'); return null; }
        const props = propsMap(json.additionalProperty);
        const title = (json.name || '').trim();
        // Autorzy: WSZYSTKIE wpisy Author z JSON-LD (propsMap brał tylko ostatni)
        let authorList = propsAll(json.additionalProperty, ['Author', 'Autor']);
        if (!authorList.length && json.author) {
            if (typeof json.author === 'string') authorList = [json.author.trim()];
            else if (Array.isArray(json.author)) {
                authorList = json.author.map(a => (typeof a === 'string' ? a : (a.name || ''))).filter(Boolean);
            } else if (json.author.name) authorList = [String(json.author.name).trim()];
        }
        if (!authorList.length) {
            // Linki /autor/ tylko w wąskim bloku nagłówka produktu
            const scope = document.querySelector(
                '.product-main-kv, .ui-product-kv, #main-product .ui-product-title, .product-header'
            );
            if (scope) {
                const names = Array.from(scope.querySelectorAll('a[href*="/autor/"]'))
                    .map(a => a.textContent.replace(/\s+/g, ' ').trim())
                    .filter(n => n && !/^zobacz\s/i.test(n));
                if (names.length && names.length <= 15) authorList = names;
            }
        }
        const author = normalizeAuthorsList(authorList);
        const publisher = ((json.brand && json.brand.name) || props['Wydawnictwo'] || '').trim();
        const language = (props['Język'] || '').trim();
        const pages = (props['Liczba stron'] || '').trim();
        const binding = normalizeBinding(props['Typ okładki'] || props['Oprawa'] || '');
        const isbn = (props['ISBN'] || json.gtin13 || '').toString().trim();
        const ean = (json.gtin13 || isbn).toString().trim();
        let year = '', premiere = '';
        if (json.releaseDate) {
            const d = String(json.releaseDate);
            year = d.slice(0, 4);
            premiere = d.slice(0, 10);
        }
        let categories = '';
        if (Array.isArray(json.category)) categories = json.category.join(' > ');
        else categories = (json.category || '').toString().trim();
        let cover = '';
        if (json.image) cover = Array.isArray(json.image) ? json.image[0] : json.image;
        if (!cover) {
            const og = document.querySelector('meta[property="og:image"]');
            if (og && og.content && !/logo/i.test(og.content)) cover = og.content;
        }
        let bruttoNum = null;
        // Preferuj cenę okładkową (sugerowaną przez wydawcę), a nie cenę sklepu
        {
            const priceRe = /Cena\s+okładkowa:?\s*([\d\s]+[,.]\d{2})/i;
            const moreEl = document.querySelector('.-more-prices');
            let m = moreEl ? moreEl.textContent.match(priceRe) : null;
            if (!m) {
                const side = document.querySelector('#main-product-sidebar, .ui-product-sidebar');
                if (side) m = side.textContent.match(priceRe);
            }
            if (!m) m = document.body.innerText.match(priceRe);
            if (m) {
                const price = parseFloat(m[1].replace(/\s/g, '').replace(',', '.'));
                if (!isNaN(price)) bruttoNum = price;
            }
        }
        const { cenaBrutto, cenaNetto, vat } = resolvePriceAndVat(bruttoNum, {
            categories, title, url: location.href, pages, binding
        });
        // Brak ceny okładkowej → nie podstawiaj ceny sklepowej z JSON-LD
        return {
            id: Date.now() + Math.random().toString(36).slice(2, 7),
            source: 'tantis', title, author, publisher,
            series: (function() {
                // Tylko dane produktu – NIE linki /serie/ z footera / polecanych
                let s = (props['Seria'] || props['Seria wydawnicza'] || '').toString().trim();
                if (!s) s = getTantisFeature(['Serie', 'Seria']);
                // Ogranicz do bloku głównego produktu (jeśli jest)
                if (!s) {
                    const scope = document.querySelector(
                        '#main-product, .product-main, .ui-product-main, [data-product], .product-details'
                    );
                    if (scope) {
                        const links = scope.querySelectorAll('a[href*="/serie/"]');
                        const names = Array.from(links).map(a => a.textContent.replace(/\s+/g, ' ').trim()).filter(Boolean);
                        if (names.length && names.length <= 3) s = [...new Set(names)].join(', ');
                    }
                }
                return s;
            })(),
            binding, year, pages, language,
            originalTitle: '', translator: '', edition: (props['Wydanie'] || '').trim(),
            isbn, ean, categories, premiereDate: premiere, cover,
            vat, cenaBrutto, cenaNetto,
            url: location.href, addedAt: new Date().toISOString()
        };
    }

    // ========== ADAPTER: ŚWIAT KSIĄŻKI ==========
    function isSwiatKsiazkiProductPage() {
        return location.pathname.includes('-ksiazka.html') ||
               !!document.querySelector('.ProductAttributes-AttributeLabel') ||
               !!document.querySelector('h1');
    }

    function extractSwiatKsiazki() {
        // Książki: .ProductAttributes-*; gry/zabawki często: .ProductActions-Attribute*
        const getAttr = (label) => {
            const want = label.toLowerCase().replace(/:$/, '').trim();
            const labels = document.querySelectorAll(
                '.ProductAttributes-AttributeLabel, .ProductActions-AttributeLabel, ' +
                '[class*="AttributeLabel"]'
            );
            for (const el of labels) {
                const t = el.textContent.replace(/\s+/g, ' ').trim().replace(/:$/, '').toLowerCase();
                if (t === want || t.startsWith(want + ' ') || t === want + ':') {
                    let val = el.parentElement && el.parentElement.querySelector(
                        '.ProductAttributes-AttributeValue, .ProductActions-AttributeValue, [class*="AttributeValue"]'
                    );
                    if (!val) {
                        // etykieta i wartość jako rodzeństwo
                        let sib = el.nextElementSibling;
                        while (sib && !/AttributeValue|value/i.test(sib.className || '')) sib = sib.nextElementSibling;
                        val = sib;
                    }
                    if (val) return val.textContent.replace(/\s+/g, ' ').trim();
                }
            }
            // fallback: „Data wydania: 22.10.2026” w tekście karty produktu
            const root = document.querySelector('.ProductActions, .ProductActionsOne, article') || document.body;
            const re = new RegExp(label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s*:\\s*([^\\n<]+)', 'i');
            const m = (root.innerText || '').match(re);
            if (m) return m[1].replace(/\s+/g, ' ').trim();
            return '';
        };
        let title = '';
        const h1 = document.querySelector('h1');
        if (h1) title = h1.textContent.replace(/\s+/g, ' ').trim();
        if (!title) title = getAttr('Tytuł');
        if (!title) {
            const og = document.querySelector('meta[property="og:title"]');
            if (og) title = og.content.replace(/\s*-\s*Książka.*/i, '').trim();
        }
        if (!title) { toast('Nie znaleziono produktu (Świat Książki)'); return null; }

        let author = '';
        const authorA = document.querySelector('.ProductActions-Author a');
        if (authorA) author = authorA.textContent.replace(/\s+/g, ' ').trim();
        if (!author) {
            const m = document.body.innerText.match(/Autor:\s*([^\n]+)/);
            if (m) author = m[1].replace(/\s+/g, ' ').trim();
        }

        let publisher = '';
        const mPub = document.body.innerText.match(/Wydawca:\s*([^\n]+)/);
        if (mPub) publisher = mPub[1].replace(/\s+/g, ' ').trim();
        // Gry: „Producent: Rebel”
        if (!publisher) {
            const prodA = document.querySelector('.ProductActions a[href*="producent"], .ProductActions a[href*="brand"]');
            if (prodA) publisher = prodA.textContent.replace(/\s+/g, ' ').trim();
        }
        if (!publisher) publisher = getAttr('Producent') || getAttr('Wydawca') || '';
        if (!publisher) {
            const mProd = document.body.innerText.match(/Producent:\s*([^\n]+)/);
            if (mProd) publisher = mProd[1].replace(/\s+/g, ' ').trim().split(/\s{2,}/)[0];
        }

        const pages = getAttr('Liczba stron');
        const binding = normalizeBinding(getAttr('Typ okładki') || getAttr('Oprawa') || '');
        const premiere = getAttr('Data wydania') || getAttr('Data premiery');
        let year = sanitizeYear(premiere ? (premiere.match(/(19|20)\d{2}/) || [])[0] : '');

        let isbn = '';
        const isbnMeta = document.querySelector('meta[property="isbn"], meta[name="isbn"]');
        if (isbnMeta) isbn = (isbnMeta.content || '').trim();
        if (!isbn) isbn = getAttr('EAN') || getAttr('ISBN') || getAttr('Kod EAN') || '';
        isbn = String(isbn).replace(/[-\s]/g, '');
        if (!/^\d{8,14}$/.test(isbn)) isbn = '';
        const ean = isbn;

        let cover = '';
        const og = document.querySelector('meta[property="og:image"]');
        if (og) cover = (og.content || '').replace(/&amp;/g, '&');
        if (!cover) {
            const img = document.querySelector('.ProductGallery img, .ProductGallery-SliderImage img');
            if (img) cover = img.src || '';
        }

        // TYLKO cena sugerowana / okładkowa – BEZ „Nasza cena” / ceny sklepu
        let bruttoNum = null;
        {
            const mainPrice = document.querySelector(
                '.ProductActions .ProductPrice, .ProductActions-Price, ' +
                '.ProductActionsOne .ProductPrice, [class*="ProductActions"] .ProductPrice'
            ) || document.querySelector('.ProductPrice:not(.ProductCard-Price)');

            // 1) przekreślona / „wysoka” cena przy produkcie (typowo okładkowa)
            const highCandidates = [];
            if (mainPrice) {
                mainPrice.querySelectorAll(
                    '.ProductPrice-HighPrice, [class*="HighPrice"], [class*="OldPrice"], ' +
                    '[class*="RegularPrice"], s, del'
                ).forEach(el => highCandidates.push(el));
            }
            document.querySelectorAll('.ProductActions .ProductPrice-HighPrice, .ProductActionsOne .ProductPrice-HighPrice')
                .forEach(el => highCandidates.push(el));

            for (const el of highCandidates) {
                const hm = (el.textContent || '').replace(/\s/g, ' ').match(/([\d]+[,.][\d]{2})/);
                if (!hm) continue;
                const price = parseFloat(hm[1].replace(',', '.'));
                if (!isNaN(price) && price >= 1 && price < 10000) { bruttoNum = price; break; }
            }

            // 2) jawny tekst „cena sugerowana / okładkowa / rekomendowana”
            if (bruttoNum == null) {
                const scope = document.querySelector('.ProductActions, .ProductActionsOne') || document.body;
                const els = scope.querySelectorAll('span, div, p, li, s, del');
                for (const el of els) {
                    if (el.closest('.ProductCard, [class*="ProductList"], [class*="Carousel"], footer')) continue;
                    const t = (el.textContent || '').replace(/\s+/g, ' ').trim();
                    if (t.length > 160 || el.children.length > 6) continue;
                    if (!/sugerowan|okładkow|okladkow|rekomendowan|cena\s+wydawc/i.test(t)) continue;
                    const m = t.match(/([\d\s]+[,.]\d{2})\s*zł/i) || t.match(/([\d]+[,.]\d{2})/);
                    if (!m) continue;
                    const price = parseFloat(m[1].replace(/\s/g, '').replace(',', '.'));
                    if (!isNaN(price) && price >= 1 && price < 10000) { bruttoNum = price; break; }
                }
            }
            // brak fallbacku na itemprop / meta / „Nasza cena”
        }
        // categories liczone niżej – wstępna heurystyka z URL/tytułu; dociągniemy po breadcrumbs
        let { cenaBrutto, cenaNetto, vat } = resolvePriceAndVat(bruttoNum, {
            title, url: location.href, pages, binding
        });

        let series = getAttr('Seria') || getAttr('Serie') || '';
        if (!series) {
            // Blok akcji produktu: „Seria: Nazwa”
            const actionRoot = document.querySelector('.ProductActions, .ProductActionsOne, article.ProductActionsOne');
            const scope = actionRoot || document.body;
            const allEls = scope.querySelectorAll('a, span, div, p, li');
            for (const el of allEls) {
                const t = el.textContent.replace(/\s+/g, ' ').trim();
                // krótki węzeł zaczynający się od Seria:/Serie:
                if (/^Seri[ae]\s*:/i.test(t) && t.length < 120) {
                    const link = el.querySelector('a');
                    if (link) {
                        series = link.textContent.replace(/\s+/g, ' ').trim();
                    } else {
                        series = t.replace(/^Seri[ae]\s*:\s*/i, '').trim();
                    }
                    if (series) break;
                }
            }
        }
        if (!series) {
            const m = document.body.innerText.match(/Seri[ae]\s*:\s*([^\n]+)/i);
            if (m) series = m[1].replace(/\s+/g, ' ').trim();
        }

        // Breadcrumbs → ścieżka kategorii (bez „Strona główna” i tytułu produktu)
        let categories = '';
        const crumbNames = document.querySelectorAll(
            '.Breadcrumbs-List .Breadcrumb [itemprop="name"], ' +
            'ul.Breadcrumbs-List li.Breadcrumb span[itemprop="name"], ' +
            'nav[aria-label*="breadcrumb" i] [itemprop="name"]'
        );
        if (crumbNames.length) {
            const parts = Array.from(crumbNames)
                .map(el => el.textContent.replace(/\s+/g, ' ').trim())
                .filter(t => t && !/^strona główna$/i.test(t));
            // Ostatni element to zwykle tytuł produktu – pomiń
            if (parts.length > 1) parts.pop();
            categories = parts.join(' > ');
        }
        if (!categories) {
            const links = document.querySelectorAll('.Breadcrumbs-List .Breadcrumb-Link, .Breadcrumbs a.Breadcrumb-Link');
            const parts = Array.from(links)
                .map(a => {
                    const n = a.querySelector('[itemprop="name"]');
                    return (n ? n.textContent : a.textContent).replace(/\s+/g, ' ').trim();
                })
                .filter(t => t && !/^strona główna$/i.test(t));
            if (parts.length > 1) parts.pop();
            categories = parts.join(' > ');
        }

        // dociągnij VAT z pełnymi kategoriami
        ({ cenaBrutto, cenaNetto, vat } = resolvePriceAndVat(bruttoNum, {
            categories, title, url: location.href, pages, binding
        }));

        return {
            id: Date.now() + Math.random().toString(36).slice(2, 7),
            source: 'swiatksiazki', title, author, publisher, series,
            binding, year, pages, language: '',
            originalTitle: '', translator: '', edition: '',
            isbn, ean, categories, premiereDate: premiere, cover,
            vat, cenaBrutto, cenaNetto,
            url: location.href, addedAt: new Date().toISOString()
        };
    }


    // ========== ADAPTER: DRESSLER ==========
    function isDresslerProductPage() {
        return location.pathname.includes('/product/') ||
               !!document.querySelector('.single-prod-info') ||
               (!!document.querySelector('h1') && !!document.querySelector('.single-prod-num'));
    }

    function getDresslerDetail(label) {
        const items = document.querySelectorAll('.single-prod-info ul li');
        const want = label.toLowerCase();
        for (const li of items) {
            const span = li.querySelector('span');
            if (!span) continue;
            const lab = span.textContent.replace(/\s+/g, ' ').trim().replace(/:$/, '').toLowerCase();
            if (lab === want) {
                // value = rest of li without the label span
                const clone = li.cloneNode(true);
                const sp = clone.querySelector('span');
                if (sp) sp.remove();
                return clone.textContent.replace(/\s+/g, ' ').trim();
            }
        }
        return '';
    }

    function extractDressler() {
        const titleEl = document.querySelector('article header h1, section article header h1, h1');
        const title = titleEl ? titleEl.textContent.replace(/\s+/g, ' ').trim() : '';
        if (!title) { toast('Nie znaleziono produktu (Dressler)'); return null; }

        const author = getDresslerDetail('Autor');
        const publisher = getDresslerDetail('Wydawnictwo');
        const year = getDresslerDetail('Rok wydania');
        let isbn = getDresslerDetail('ISBN').replace(/-/g, '');
        let ean = getDresslerDetail('EAN').replace(/-/g, '');
        if (!ean) ean = isbn;
        if (!isbn) isbn = ean;

        const pages = getDresslerDetail('Strony');
        const binding = normalizeBinding(getDresslerDetail('Oprawa'));
        const premiere = getDresslerDetail('Data premiery');
        const series = getDresslerDetail('Serie') || getDresslerDetail('Seria');

        // Kategorie: links + text joined
        let categories = '';
        const items = document.querySelectorAll('.single-prod-info ul li');
        for (const li of items) {
            const span = li.querySelector('span');
            if (!span) continue;

            if (/kategorie/i.test(span.textContent.trim())) {
                const links = li.querySelectorAll('a');
                if (links.length) {
                    categories = Array.from(links)
                        .map(a => a.textContent.replace(/\s+/g, ' ').trim())
                        .filter(Boolean)
                        .join(' > ');
                } else {
                    const clone = li.cloneNode(true);
                    const sp = clone.querySelector('span');
                    if (sp) sp.remove();
                    categories = clone.textContent
                        .replace(/\s+/g, ' ')
                        .replace(/\s*>\s*/g, ' > ')
                        .trim();
                }
                break;
            }
        }

        // Cena: zawsze cena detaliczna / okładkowa wyłącznie z bloku .single-prod-num.
        // Nie czytamy .new-price z rekomendacji, karuzeli ani innych sekcji strony.
        let cenaBrutto = '', cenaNetto = '', vat = '';

        const getDresslerRetailPrice = () => {
            const priceRoot = document.querySelector('.single-prod-num');
            if (!priceRoot) return null;

            // innerHTML zachowuje również komentarze HTML,
            // w których Dressler pozostawia cenę detaliczną po zalogowaniu.
            const html = priceRoot.innerHTML;

            const matches = html.matchAll(
                /<div\s+class=["']new-price["'][^>]*>\s*([\d\s.,]+)\s*zł\s*<\/div>/gi
            );

            let highest = null;

            for (const match of matches) {
                const raw = match[1]
                    .replace(/\s/g, '')
                    .replace(/\.(?=\d{3}\b)/g, '')
                    .replace(',', '.');

                const price = parseFloat(raw);
                if (!isNaN(price) && price > 0) {
                    if (highest === null || price > highest) highest = price;
                }
            }

            return highest;
        };

        const retailPrice = getDresslerRetailPrice();

        // VAT wyłącznie z bloku cenowego produktu
        const vatEl = priceRootVat();
        if (vatEl) {
            const vm =
                vatEl.textContent.match(/(\d{1,2})\s*%?\s*VAT/i) ||
                vatEl.textContent.match(/VAT\s*:?\s*(\d{1,2})/i);

            if (vm) {
                const value = parseInt(vm[1], 10);
                if ([0, 5, 8, 23].includes(value)) vat = String(value);
            }
        }

        function priceRootVat() {
            const root = document.querySelector('.single-prod-num');
            return root ? root.querySelector('.vat-info') : null;
        }

        if (retailPrice !== null) {
            const resolved = resolvePriceAndVat(retailPrice, {
                categories,
                title,
                url: location.href,
                pages,
                binding
            });

            cenaBrutto = resolved.cenaBrutto;
            cenaNetto = resolved.cenaNetto;
            vat = resolved.vat;
        }

        // Okładka – wyłącznie główna galeria aktualnego produktu Dresslera
        let cover = '';

        const coverLink = document.querySelector(
            '.single-prod .product-gallery li a[href]'
        );

        const coverImg = document.querySelector(
            '.single-prod .product-gallery li img'
        );

        if (coverLink?.href) {
            cover = coverLink.href;
        } else if (coverImg?.src) {
            cover = coverImg.src;
        }

        return {
            id: Date.now() + Math.random().toString(36).slice(2, 7),
            source: 'dressler',
            title,
            author,
            publisher,
            series,
            binding,
            year,
            pages,
            language: '',
            originalTitle: '',
            translator: '',
            edition: '',
            isbn,
            ean,
            categories,
            premiereDate: premiere,
            cover,
            vat,
            cenaBrutto,
            cenaNetto,
            url: location.href,
            addedAt: new Date().toISOString()
        };
    }

    // ========== ADAPTER: EMPIK ==========
    function isEmpikProductPage() {
        // URL typu ...,p123456,ksiazka-p  lub  ...,p123456,ebooki-i-mp3-p
        if (/\/p\d+[,/]/.test(location.pathname) || /\/p\d+$/.test(location.pathname)) return true;
        if (document.querySelector('h1') && (
            document.querySelector('[data-ta="product-title"]') ||
            document.body.innerText.includes('Informacje szczegółowe') ||
            (document.querySelector('table') && /ID produktu/i.test(document.body.innerText))
        )) return true;
        return false;
    }

    function isValidEANCode(str) {
        if (!str) return false;
        return /^\d{8,14}$/.test(String(str).trim());
    }

    /** EAN z GraphQL / drugiego skryptu Empik / DOM */
    function getEmpikEAN() {
        // 1) Shared state z innego userscriptu (Empik - Zapisz grafiki)
        try {
            const shared = (typeof unsafeWindow !== 'undefined' ? unsafeWindow : window).__TM_EMPIK_EAN__;
            if (shared && isValidEANCode(shared.value)) return String(shared.value).trim();
        } catch (e) {}
        try {
            if (window.__TM_EMPIK_EAN__ && isValidEANCode(window.__TM_EMPIK_EAN__.value)) {
                return String(window.__TM_EMPIK_EAN__.value).trim();
            }
        } catch (e) {}

        // 2) Atrybuty w DOM
        let ean = getEmpikDetail('EAN') || getEmpikDetail('Kod EAN') || getEmpikDetail('ISBN') || '';
        ean = String(ean).replace(/[-\s]/g, '');
        if (isValidEANCode(ean)) return ean;

        // 3) Meta / JSON-LD
        const isbnMeta = document.querySelector('meta[property="isbn"], meta[name="isbn"], meta[property="product:retailer_item_id"]');
        if (isbnMeta && isValidEANCode(isbnMeta.content)) return String(isbnMeta.content).trim();

        const scripts = document.querySelectorAll('script[type="application/ld+json"]');
        for (const s of scripts) {
            try {
                const data = JSON.parse(s.textContent);
                const candidates = [];
                const walk = (o) => {
                    if (!o || typeof o !== 'object') return;
                    if (o.gtin13) candidates.push(o.gtin13);
                    if (o.gtin) candidates.push(o.gtin);
                    if (o.isbn) candidates.push(o.isbn);
                    if (o.ean) candidates.push(o.ean);
                    if (Array.isArray(o)) o.forEach(walk);
                    else Object.values(o).forEach(walk);
                };
                walk(data);
                for (const c of candidates) {
                    const v = String(c).replace(/[-\s]/g, '');
                    if (isValidEANCode(v)) return v;
                }
            } catch (e) {}
        }

        return '';
    }

    /** Rozwiń sekcję szczegółów Empik (lazy / collapsible) */
    function expandEmpikDetails() {
        const clickables = Array.from(document.querySelectorAll('button, a, [role="button"], span, div'));
        const expandRe = /pokaż\s+(wszystkie|więcej|szczegóły)|rozwiń|więcej\s+informacji|informacje\s+szczegółowe/i;
        for (const el of clickables) {
            const t = (el.textContent || '').replace(/\s+/g, ' ').trim();
            if (!t || t.length > 60) continue;
            if (!expandRe.test(t)) continue;
            if (el.closest('nav, footer, header, [class*="menu"]')) continue;
            try { el.click(); } catch (e) {}
        }
        document.querySelectorAll('details:not([open])').forEach(d => {
            const sum = (d.querySelector('summary')?.textContent || '').toLowerCase();
            if (/szczegół|atrybut|informac|specyfik/i.test(sum) || !sum) {
                try { d.open = true; } catch (e) {}
            }
        });
        const anchor = Array.from(document.querySelectorAll('h2, h3, [class*="Attribute"], [data-ta*="detail"]'))
            .find(el => /informacje\s+szczegółowe|szczegóły\s+produktu|atrybuty/i.test(el.textContent || ''));
        if (anchor) {
            try { anchor.scrollIntoView({ block: 'center', behavior: 'instant' }); } catch (e) {
                try { anchor.scrollIntoView(true); } catch (e2) {}
            }
        }
    }

    /** Szczegóły „głębokie”: język / wymiary / oprawa – często po „Pokaż wszystkie” */
    function isEmpikDetailsExpanded() {
        const text = document.body.innerText || '';
        return /Język wydania|Liczba stron|Wysokość\s*\[?\s*mm|Okładka\s*:|Forma\s*:|Numer wydania/i.test(text);
    }

    /** Czy strona Empik ma już „mięso” produktu (nie sam skeleton) */
    function isEmpikDataReady() {
        if (!isEmpikProductPage()) return false;
        const titleEl = document.querySelector('h1[data-ta="product-title"], h1');
        const title = titleEl ? titleEl.textContent.replace(/\s+/g, ' ').trim() : '';
        if (!title || title.length < 2) return false;
        const hasDetails = /Informacje szczegółowe|ID produktu|Wydawnictwo|Liczba stron|Data premiery/i.test(document.body.innerText || '');
        const hasEan = !!getEmpikEAN();
        const hasAuthor = !!document.querySelector('a[href*="/autor/"], a[href*="/creator/"], [data-ta*="author"]');
        return hasDetails && (hasEan || hasAuthor);
    }

/** Szczegóły są kompletne, gdy widać końcowe wiersze listy (Indeks / EAN) – wcześniejsze wiersze są wtedy już wyrenderowane. */
    function isEmpikDetailsComplete() {
        const text = document.body.innerText || '';
        return /(Indeks|EAN)\s*:?\s*\d{5,}/i.test(text);
    }

    /** Podpis pól szczegółowych – gdy przestaje się zmieniać, strona doładowała dane. */
    function empikDetailsSignature() {
        return ['Język wydania', 'Liczba stron', 'Wysokość', 'Głębokość', 'Szerokość', 'Okładka', 'Forma', 'Numer wydania']
            .map(l => getEmpikDetail(l)).join('|');
    }

    function getEmpikDetail(label) {
        // 1) Tabela / lista atrybutów „Informacje szczegółowe”
        const rows = document.querySelectorAll(
            'table tr, [class*="Attribute"] tr, [class*="details"] tr, [class*="attribute"] li, ' +
            'dl > div, dl > dt, .productAttributes li, [data-ta*="attribute"], ' +
            '[class*="ProductAttribute"], [class*="product-attribute"]'
        );
        const want = label.toLowerCase().replace(/:$/, '').trim();
        for (const row of rows) {
            // wiersz tabeli / flex: etykieta + wartość
            const cells = row.querySelectorAll(':scope > td, :scope > th, :scope > dt, :scope > dd, :scope > span, :scope > div, :scope > p');
            if (cells.length >= 2) {
                const lab = (cells[0].textContent || '').replace(/\s+/g, ' ').trim().replace(/:$/, '').toLowerCase();
                // dokładne lub „Wysokość [mm]” zaczyna się od „wysokość”
                if (lab === want || lab.startsWith(want) || lab.replace(/\s*\[.*?\]\s*/g, '').trim() === want) {
                    return (cells[cells.length - 1].textContent || '').replace(/\s+/g, ' ').trim();
                }
            }
            // dt/dd osobno
            if (row.tagName === 'DT') {
                const lab = row.textContent.replace(/\s+/g, ' ').trim().replace(/:$/, '').toLowerCase();
                if (lab === want || lab.startsWith(want) || lab.replace(/\s*\[.*?\]\s*/g, '').trim() === want) {
                    const dd = row.nextElementSibling;
                    if (dd && dd.tagName === 'DD') return dd.textContent.replace(/\s+/g, ' ').trim();
                }
            }
        }
        // 2) Fallback – tekst strony (np. „Wydawnictwo: Xxx” lub „Wysokość [mm]: 195”)
        const re = new RegExp(
            label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') +
            '(?:\\s*\\[.*?\\])?\\s*:\\s*([^\\n]+)',
            'i'
        );
        const m = document.body.innerText.match(re);
        return m ? m[1].replace(/\s+/g, ' ').trim() : '';
    }

    function getEmpikSuggestedPrice() {
        const mainRoot =
            document.querySelector(
                '[data-ta="product-main"], [data-ta="product-price"], main, ' +
                '[class*="ProductDetails"], [class*="product-details"], ' +
                '[class*="ProductMain"], article[class*="Product"]'
            ) || document.querySelector('h1[data-ta="product-title"], h1')?.closest('section, article, main, div') ||
            document.body;

        const isOutsideMain = (el) => {
            if (!el || !mainRoot.contains(el)) return true;
            return !!el.closest(
                '[class*="recomm"], [class*="Recomm"], [class*="similar"], [class*="Similar"], ' +
                '[class*="carousel"], [class*="Carousel"], [class*="slider"], [class*="Slider"], ' +
                '[class*="sponsor"], [data-ta*="recommend"], [data-ta*="similar"], ' +
                '[aria-label*="podobn" i], [aria-label*="Polecane" i]'
            );
        };

        const candidates = Array.from(mainRoot.querySelectorAll('span, div, p, li, td, strong, em'))
            .filter(el => {
                if (isOutsideMain(el)) return false;
                const t = (el.textContent || '').replace(/\s+/g, ' ').trim();
                return /cena\s+sugerowana/i.test(t) && t.length < 160 && el.children.length < 8;
            });
        for (const el of candidates) {
            const t = el.textContent.replace(/\s+/g, ' ');
            const m = t.match(/([\d\s]+[,.]\d{2})\s*zł\s*[-–—]?\s*cena\s+sugerowana/i) ||
                      t.match(/cena\s+sugerowana[^0-9]{0,40}([\d\s]+[,.]\d{2})/i) ||
                      t.match(/([\d\s]+[,.]\d{2})/);
            if (m) {
                const price = parseFloat(m[1].replace(/\s/g, '').replace(',', '.'));
                if (!isNaN(price) && price >= 1 && price < 10000) return price;
            }
        }
        return null;
    }

    /** Ścieżka kategorii Empik: bierze najdłuższą z kilku źródeł (DOM, meta). */
    function getEmpikCategories() {
        const clean = (t) => String(t || '').replace(/\s+/g, ' ').trim();
        const junk = /^(strona główna|empik|home|[>\/›»|·•-])$/i;
        const titleEl = document.querySelector('h1[data-ta="product-title"], h1');
        const productTitle = clean(titleEl && titleEl.textContent).toLowerCase();

        const norm = (arr) => {
            const out = [];
            for (const raw of arr) {
                const t = clean(raw);
                if (!t || junk.test(t)) continue;
                if (out.length && out[out.length - 1].toLowerCase() === t.toLowerCase()) continue;
                out.push(t);
            }
            // ostatni element pomijamy tylko, gdy to tytuł produktu
            if (out.length && productTitle && out[out.length - 1].toLowerCase() === productTitle) out.pop();
            return out;
        };

        // Liście kontenera = wszystkie okruszki, także ostatni, który NIE jest linkiem
        const leaves = (root) => Array.from(root.querySelectorAll('*'))
            .filter(el => el.children.length === 0 && clean(el.textContent))
            .map(el => el.textContent);

        const candidates = [];

        // 1) DOM: kontener okruszków
        const seen = new Set();
        const roots = [];
        const anchor = document.querySelector('a[data-ta="breadcrumb-category"]');
        if (anchor) {
            const c = anchor.closest('nav, ol, ul, [class*="readcrumb"]');
            if (c) roots.push(c);
        }
        document.querySelectorAll(
            'nav[aria-label*="breadcrumb" i], [data-ta*="breadcrumb"], [class*="readcrumb"]'
        ).forEach(el => roots.push(el));
        for (const r of roots) {
            if (seen.has(r)) continue;
            seen.add(r);
            if (r.closest('footer, [class*="recomm"], [class*="similar"], [class*="carousel"]')) continue;
            if (clean(r.textContent).length > 300) continue;
            const parts = norm(leaves(r));
            if (parts.length) candidates.push(parts);
        }

        // 2) meta product:category_tree
        const tree = document.querySelector('meta[property="product:category_tree"]');
        if (tree && tree.content) candidates.push(norm(tree.content.split(/\s*>\s*/)));

        // 3) meta category1/2
        const c1 = document.querySelector('meta[property="product:category1"]');
        const c2 = document.querySelector('meta[property="product:category2"]');
        const metaParts = norm([c1 && c1.content, c2 && c2.content]);
        if (metaParts.length) candidates.push(metaParts);

        if (!candidates.length) return '';
        candidates.sort((a, b) => b.length - a.length);
        return candidates[0].join(' > ');
    }

    function extractEmpik() {
        const titleEl = document.querySelector('h1[data-ta="product-title"], h1');
        let title = titleEl ? titleEl.textContent.replace(/\s+/g, ' ').trim() : '';
        if (!title) {
            const og = document.querySelector('meta[property="og:title"]');
            if (og) title = og.content.replace(/\s*[|\-–]\s*Empik.*$/i, '').trim();
        }
        if (!title) { toast('Nie znaleziono tytułu produktu (Empik)'); return null; }

        // Autor – często „Nazwisko Imię”
        let author = getEmpikDetail('Autor') || getEmpikDetail('Autorzy');
        if (!author) {
            const authorLinks = document.querySelectorAll('a[href*="/autor/"], a[href*="/creator/"], [data-ta*="author"] a');
            const names = Array.from(authorLinks).map(a => a.textContent.replace(/\s+/g, ' ').trim()).filter(Boolean);
            if (names.length && names.length <= 8) author = [...new Set(names)].join(', ');
        }
        author = normalizeAuthorsList(author);

        const publisher = getEmpikDetail('Wydawnictwo') || getEmpikDetail('Producent') || '';
        const series = getEmpikDetail('Seria') || getEmpikDetail('Serie') || '';
        const translator = getEmpikDetail('Tłumaczenie') || getEmpikDetail('Tłumacz') || '';
        const language = getEmpikDetail('Język wydania') || getEmpikDetail('Język') || '';
        const originalTitle = getEmpikDetail('Tytuł oryginalny') || getEmpikDetail('Tytuł oryginału') || '';
        const pages = getEmpikDetail('Liczba stron') || getEmpikDetail('Strony') || '';
        const edition = getEmpikDetail('Numer wydania') || getEmpikDetail('Wydanie') || '';
        let premiere = getEmpikDetail('Data premiery') || getEmpikDetail('Data wydania') || '';
        let year = getEmpikDetail('Rok wydania') || '';
        if (!year && premiere) {
            const ym = premiere.match(/(19|20)\d{2}/);
            if (ym) year = ym[0];
        }

        // Oprawa – „Okładka: miękka, barwione brzegi”
        let bindingRaw = getEmpikDetail('Okładka') || getEmpikDetail('Oprawa') || getEmpikDetail('Typ okładki') || '';
        if (!bindingRaw) {
            const m = document.body.innerText.match(/okładka\s*:?\s*(miękka|twarda|broszurowa|kartonowa)/i) ||
                      document.body.innerText.match(/(miękka|twarda)\s*,?\s*barwione/i);
            if (m) bindingRaw = m[1] || m[0];
        }
        const binding = normalizeBinding(bindingRaw);


        // EAN (ISBN tylko jako źródło kodu)
        const ean = getEmpikEAN();
        const isbn = ean;

        // Kategorie (breadcrumbs)
        const categories = getEmpikCategories();

        // Okładka
        let cover = '';
        const ogImg = document.querySelector('meta[property="og:image"]');
        if (ogImg && ogImg.content && !/logo|placeholder/i.test(ogImg.content)) cover = ogImg.content;
        if (!cover) {
            const img = document.querySelector(
                'img[src*="ecsmedia.pl"], [data-ta*="gallery"] img, .productGallery img, [class*="gallery"] img[src*="ecsmedia"]'
            );
            if (img && img.src) cover = img.src;
        }

        // Forma produktu (książka / zabawka…)
        const form = getEmpikDetail('Forma') || '';

        // Cena – TYLKO „cena sugerowana” z głównej karty produktu (nie z „Podobne”)
        const bruttoNum = getEmpikSuggestedPrice();
        const { cenaBrutto, cenaNetto, vat } = resolvePriceAndVat(bruttoNum, {
            categories, title, url: location.href, pages, binding, form
        });

        return {
            id: Date.now() + Math.random().toString(36).slice(2, 7),
            source: 'empik',
            title,
            author,
            publisher,
            series,
            binding,
            year,
            pages,
            language,
            originalTitle,
            translator,
            edition,
            isbn,
            ean,
            categories,
            premiereDate: premiere,
            cover,
            vat,
            cenaBrutto,
            cenaNetto,
            url: location.href,
            addedAt: new Date().toISOString()
        };
    }

    // ========== ADAPTER: ATENEUM ==========
    function isAteneumProductPage() {
        if (/pozycja\.php/i.test(location.pathname) && /[?&]id=\d+/.test(location.search)) return true;
        if (document.getElementById('product_name') && document.getElementById('product_id')) return true;
        if (document.querySelector('.product-property-name') && document.querySelector('#product_name')) return true;
        return false;
    }

    function getAteneumDetail(label) {
        const want = label.toLowerCase().replace(/:$/, '').trim();
        const names = document.querySelectorAll('.product-property-name');
        for (const nameEl of names) {
            const lab = (nameEl.textContent || '').replace(/\s+/g, ' ').trim().replace(/:$/, '').toLowerCase();
            if (lab === want || lab.startsWith(want)) {
                // wartość w .col-6 > .product-property w tym samym .row
                const row = nameEl.closest('.row');
                if (row) {
                    const valEl = row.querySelector('.product-property');
                    if (valEl) return valEl.textContent.replace(/\s+/g, ' ').trim();
                }
            }
        }
        return '';
    }

    function extractAteneum() {
        const titleEl = document.getElementById('product_name') || document.querySelector('h5#product_name, h1, h5.text-black');
        let title = titleEl ? titleEl.textContent.replace(/\s+/g, ' ').trim() : '';
        if (!title) {
            const og = document.querySelector('meta[property="og:title"]');
            if (og) title = og.content.replace(/\s*[-–|]\s*Dystrybucja Ateneum.*$/i, '').trim();
        }
        if (!title) { toast('Nie znaleziono tytułu produktu (Ateneum)'); return null; }

        let author = getAteneumDetail('Autor');
        author = normalizeAuthorsList(author);
        const publisher = getAteneumDetail('Wydawnictwo') || getAteneumDetail('Producent') || '';
        const series = getAteneumDetail('Seria') || getAteneumDetail('Cykl') || '';
        const premiere = getAteneumDetail('Data premiery') || '';
        let year = sanitizeYear(getAteneumDetail('Rok wydania') || (premiere ? premiere.slice(0, 4) : ''));
        const binding = normalizeBinding(getAteneumDetail('Oprawa') || '');
        const pages = (getAteneumDetail('Strony') || getAteneumDetail('Liczba stron') || '').replace(/[^\d]/g, '');
        let ean = (getAteneumDetail('EAN') || getAteneumDetail('ISBN') || '').replace(/[-\s]/g, '');
        if (!/^\d{8,14}$/.test(ean)) ean = '';
        const isbn = ean;

        // Cena sugerowana brutto + stawka VAT ze strony
        let bruttoNum = null;
        const bruttoRaw = getAteneumDetail('Cena sugerowana brutto') || getAteneumDetail('Cena brutto') || '';
        if (bruttoRaw) {
            const n = parseFloat(bruttoRaw.replace(/\s/g, '').replace(',', '.').replace(/[^\d.]/g, ''));
            if (!isNaN(n) && n > 0) bruttoNum = n;
        }
        let vat = '';
        const vatRaw = getAteneumDetail('Stawka vat') || getAteneumDetail('Stawka VAT') || getAteneumDetail('VAT') || '';
        const vm = vatRaw.match(/(\d{1,2})/);
        if (vm && [0, 5, 8, 23].includes(parseInt(vm[1], 10))) vat = vm[1];
        if (!vat) vat = detectVatFromPage();
        // categories liczone niżej – wstępna heurystyka; po zbudowaniu ścieżki dociągamy VAT 5%
        if (!vat && isLikelyBookProduct({ title, url: location.href, pages, binding })) vat = '5';
        let { cenaBrutto, cenaNetto, vat: vatOut } = priceWithVat(bruttoNum, vat);
        vat = vatOut || vat;

        // Okładka – ten sam host co strona.
        // og:image bywa https://www.ateneum.pl/... przy wejściu z ateneum.net.pl → martwy obraz.
        const absUrl = (u) => {
            if (!u) return '';
            const s = String(u).trim();
            if (!s || /load\.gif|no_photo|placeholder|logo\.svg/i.test(s)) return '';
            try {
                const url = new URL(s, location.href);
                if (/\/phttw\//i.test(url.pathname) && /ateneum\.(pl|net\.pl)$/i.test(url.hostname)) {
                    url.protocol = location.protocol;
                    url.hostname = location.hostname;
                }
                return url.href;
            } catch (e) { return s; }
        };
        let cover = '';
        // 1) obraz na stronie (ten sam origin)
        const img = document.querySelector(
            '#product-image-container img.product-card-image, #product-image-container img, ' +
            'img.product-card-image, img[src*="/phttw/"], img[data-src*="/phttw/"]'
        );
        if (img) {
            cover = absUrl(img.getAttribute('data-src')) || absUrl(img.currentSrc) ||
                    absUrl(img.getAttribute('src')) || absUrl(img.src);
        }
        // 2) id → /phttw/{id}_large.jpg na bieżącym hostcie
        if (!cover) {
            const pidEl = document.getElementById('product_id');
            const pid = (pidEl && pidEl.value) || (location.search.match(/[?&]id=(\d+)/) || [])[1];
            if (pid) cover = location.origin + '/phttw/' + pid + '_large.jpg';
        }
        // 3) og:image z przepisaniem hosta
        if (!cover) {
            const og = document.querySelector('meta[property="og:image"], meta[property="og:image:secure_url"]');
            if (og && og.content) cover = absUrl(og.content);
        }

        // Kategorie – #category-path-container (Ateneum: Literatura > Fantastyka, Fantasy)
        let categories = '';
        const pathLinks = document.querySelectorAll(
            '#category-path-container .category-path a, #category-path-container a, p.category-path a'
        );
        if (pathLinks.length) {
            const parts = Array.from(pathLinks).map(a => a.textContent.replace(/\s+/g, ' ').trim())
                .filter(t => t && !/^strona główna$/i.test(t));
            if (parts.length) categories = parts.join(' > ');
        }
        if (!categories) {
            const pathP = document.querySelector('#category-path-container .category-path, p.category-path');
            if (pathP) {
                categories = pathP.textContent.replace(/\s+/g, ' ').replace(/\s*>\s*/g, ' > ').trim();
            }
        }

        // ścieżka z książką → 5% VAT (nawet bez ceny), chyba że strona podała inną stawkę
        if (!vat) {
            const r = resolvePriceAndVat(bruttoNum, { categories, title, url: location.href, pages, binding });
            vat = r.vat || '';
            if (r.cenaBrutto) { cenaBrutto = r.cenaBrutto; cenaNetto = r.cenaNetto; }
        }

        return {
            id: Date.now() + Math.random().toString(36).slice(2, 7),
            source: 'ateneum',
            title,
            author,
            publisher,
            series,
            binding,
            year,
            pages,
            language: '',
            originalTitle: '',
            translator: '',
            edition: '',
            isbn,
            ean,
            categories,
            premiereDate: premiere,
            cover,
            vat,
            cenaBrutto,
            cenaNetto,
            url: location.href,
            addedAt: new Date().toISOString()
        };
    }

    const SITES = {
        'www.taniaksiazka.pl': { name: 'TaniaKsiążka', isProductPage: isTaniaKsiazkaProductPage, extract: extractTaniaKsiazka },
        'taniaksiazka.pl':     { name: 'TaniaKsiążka', isProductPage: isTaniaKsiazkaProductPage, extract: extractTaniaKsiazka },
        'platon.com.pl':       { name: 'Platon', isProductPage: isPlatonProductPage, extract: extractPlaton },
        'www.platon.com.pl':   { name: 'Platon', isProductPage: isPlatonProductPage, extract: extractPlaton },
        'bonito.pl':           { name: 'Bonito', isProductPage: isBonitoProductPage, extract: extractBonito },
        'www.bonito.pl':       { name: 'Bonito', isProductPage: isBonitoProductPage, extract: extractBonito },
        'tantis.pl':           { name: 'Tantis', isProductPage: isTantisProductPage, extract: extractTantis },
        'www.tantis.pl':       { name: 'Tantis', isProductPage: isTantisProductPage, extract: extractTantis },
        'swiatksiazki.pl':     { name: 'Świat Książki', isProductPage: isSwiatKsiazkiProductPage, extract: extractSwiatKsiazki },
        'www.swiatksiazki.pl': { name: 'Świat Książki', isProductPage: isSwiatKsiazkiProductPage, extract: extractSwiatKsiazki },
        'dressler.com.pl':     { name: 'Dressler', isProductPage: isDresslerProductPage, extract: extractDressler },
        'www.dressler.com.pl': { name: 'Dressler', isProductPage: isDresslerProductPage, extract: extractDressler },
        'www.empik.com':       { name: 'Empik', isProductPage: isEmpikProductPage, extract: extractEmpik },
        'empik.com':           { name: 'Empik', isProductPage: isEmpikProductPage, extract: extractEmpik },
        'www.ateneum.net.pl':  { name: 'Ateneum', isProductPage: isAteneumProductPage, extract: extractAteneum },
        'ateneum.net.pl':      { name: 'Ateneum', isProductPage: isAteneumProductPage, extract: extractAteneum },
        'www.ateneum.pl':      { name: 'Ateneum', isProductPage: isAteneumProductPage, extract: extractAteneum },
        'ateneum.pl':          { name: 'Ateneum', isProductPage: isAteneumProductPage, extract: extractAteneum }
    };

    function getCurrentSite() { return SITES[location.hostname] || null; }

/** Ustawia filtr i synchronizuje go z pozostałymi otwartymi panelami. */
    function setFilter(value, { render = true } = {}) {
        filterQuery = String(value || '');
        const inp = ui('#tk-search');
        if (inp && inp.value !== filterQuery) inp.value = filterQuery;
        const gInp = ui('#tk-grid-search');
        if (gInp && gInp.value !== filterQuery) gInp.value = filterQuery;
        clearTimeout(filterPersistTimer);
        filterPersistTimer = setTimeout(() => GM_setValue('tk_filter', filterQuery), 150);
        if (render) renderList();
    }

    /** Zmiana filtra z innej karty – bez zapisu zwrotnego (żeby nie robić pętli). */
    function applyRemoteFilter(value) {
        const v = String(value || '');
        if (v === filterQuery) return;
        filterQuery = v;
        const inp = ui('#tk-search');
        if (inp && inp.value !== v) inp.value = v;
        const gInp = ui('#tk-grid-search');
        if (gInp && gInp.value !== v) gInp.value = v;
        if (ui('.tk-edit-input')) scheduleRender(); else renderList();
    }

    if (typeof GM_addValueChangeListener === 'function') {
        GM_addValueChangeListener('tk_filter', (name, oldV, newV, remote) => {
            if (remote) applyRemoteFilter(newV);
        });
    }

/** Jednorazowo poprawia zapisane już pozycje z wieloma językami. */
    function migrateLanguages() {
        const cur = GM_getValue('tk_products', []);
        let changed = false;
        cur.forEach(p => {
            const n = normalizeLanguage(p.language);
            if (n !== (p.language || '')) { p.language = n; changed = true; }
        });
        if (changed) {
            GM_setValue('tk_products', cur);
            products = cur;
            renderList();
            updateCount();
        }
    }

    function loadProductsFromStorage() {
        applyRemoteFilter(GM_getValue('tk_filter', ''));
        const fresh = GM_getValue('tk_products', []);
        covers = GM_getValue('tk_covers', {});
        if (JSON.stringify(fresh) === JSON.stringify(products)) return;
        if (ui('.tk-edit-input')) return;
        products = fresh;
        renderList();
        updateCount();
    }

    function saveProducts(highlightId = null) {
        GM_setValue('tk_products', products);
        renderList(highlightId);
        updateCount();
    }

    if (typeof GM_addValueChangeListener === 'function') {
        GM_addValueChangeListener('tk_products', (name, oldValue, newValue, remote) => {
            if (remote) {
                products = Array.isArray(newValue) ? newValue : [];
                renderList();
                updateCount();
            }
        });
    }

    document.addEventListener('visibilitychange', () => {
        if (document.visibilityState === 'visible') loadProductsFromStorage();
    });
    window.addEventListener('focus', loadProductsFromStorage);

    function absolutizeCoverUrl(u) {
        if (!u) return '';
        const s = String(u).trim();
        if (!s || /load\.gif|no_photo|placeholder|data:image\/svg/i.test(s)) return '';
        try { return new URL(s, location.href).href; } catch (e) { return s; }
    }

    // ========== OKŁADKI: miniatury 200 px zapisywane przy „Załaduj” ==========
    const COVER_MAX = 200;
    const COVER_QUALITY = 0.82;
    let covers = GM_getValue('tk_covers', {});   // { [productId]: dataURL }
    const coverFailed = new Set();               // porażki tylko w tej sesji strony
    let renderTimer = null;

    function scheduleRender() {
        clearTimeout(renderTimer);
        renderTimer = setTimeout(() => { if (!ui('.tk-edit-input')) renderList(); }, 300);
    }

    function fetchImageBlob(url) {
        return new Promise((resolve, reject) => {
            if (typeof GM_xmlhttpRequest !== 'function') return reject(new Error('no GM_xmlhttpRequest'));
            GM_xmlhttpRequest({
                method: 'GET', url, responseType: 'blob', timeout: 20000,
                headers: { Accept: 'image/*', Referer: location.href },
                onload(res) {
                    const b = res.response;
                    if (res.status >= 200 && res.status < 300 && b && b.size > 0) resolve(b);
                    else reject(new Error('bad response ' + res.status));
                },
                onerror() { reject(new Error('network')); },
                ontimeout() { reject(new Error('timeout')); }
            });
        });
    }

    function blobToDataUrl(blob) {
        return new Promise((resolve, reject) => {
            const r = new FileReader();
            r.onload = () => resolve(r.result);
            r.onerror = () => reject(r.error);
            r.readAsDataURL(blob);
        });
    }

    // createImageBitmap nie podlega CSP strony i odrzuci HTML podstawiony zamiast obrazu
    async function makeThumbDataUrl(blob) {
        const bmp = await createImageBitmap(blob);
        const scale = Math.min(1, COVER_MAX / Math.max(bmp.width, bmp.height));
        const w = Math.max(1, Math.round(bmp.width * scale));
        const h = Math.max(1, Math.round(bmp.height * scale));
        const c = document.createElement('canvas');
        c.width = w; c.height = h;
        const ctx = c.getContext('2d');
        ctx.fillStyle = '#fff';               // JPEG nie ma przezroczystości
        ctx.fillRect(0, 0, w, h);
        ctx.drawImage(bmp, 0, 0, w, h);
        if (bmp.close) bmp.close();
        return c.toDataURL('image/jpeg', COVER_QUALITY);
    }

    async function ensureCoverThumb(p) {
        if (!p || !p.id || !p.cover) return false;
        covers = GM_getValue('tk_covers', {});
        if (covers[p.id]) return true;
        if (coverFailed.has(p.id)) return false;
        try {
            const blob = await fetchImageBlob(p.cover);
            const thumb = await makeThumbDataUrl(blob);
            // produkt mógł zostać usunięty w trakcie pobierania
            if (!GM_getValue('tk_products', []).some(x => x.id === p.id)) return false;
            const cur = GM_getValue('tk_covers', {});   // merge – bez nadpisywania innych kart
            cur[p.id] = thumb;
            GM_setValue('tk_covers', cur);
            covers = cur;
            return true;
        } catch (e) {
            coverFailed.add(p.id);
            console.warn('[MultiKsiegarnia] Okładka nie pobrana:', p.cover, e);
            return false;
        }
    }

    let backfillRunning = false;
    /** Dociąga miniatury dla starych pozycji – po jednej, z przerwą */
    async function backfillCovers() {
        if (backfillRunning) return;
        backfillRunning = true;
        try {
            for (const p of [...products]) {
                if (!p.cover || covers[p.id] || coverFailed.has(p.id)) continue;
                if (await ensureCoverThumb(p)) scheduleRender();
                await new Promise(r => setTimeout(r, 250));
            }
        } finally { backfillRunning = false; }
    }

    function dropCovers(ids) {
        const cur = GM_getValue('tk_covers', {});
        let changed = false;
        ids.forEach(id => { if (id in cur) { delete cur[id]; changed = true; } });
        if (changed) GM_setValue('tk_covers', cur);
        covers = cur;
    }

    if (typeof GM_addValueChangeListener === 'function') {
        GM_addValueChangeListener('tk_covers', (name, oldV, newV, remote) => {
            if (remote) { covers = newV || {}; scheduleRender(); }
        });
    }

    function focusProductInList(productId, { duplicate = false } = {}) {
        if (!productId) return;
        // wyczyść filtr, żeby pozycja była na liście
        if (filterQuery) setFilter('', { render: false });
        renderList(null);
        const list = ui('.tk-list');
        if (!list) return;
        const item = list.querySelector(`.tk-item[data-id="${CSS.escape(String(productId))}"]`);
        if (!item) return;

        item.classList.remove('tk-item-new', 'tk-item-dup');
        void item.offsetWidth;
        item.classList.add(duplicate ? 'tk-item-dup' : 'tk-item-new');

        // scroll wewnątrz .tk-list (to ten kontener ma overflow-y: auto)
        const targetTop = item.offsetTop - (list.clientHeight / 2) + (item.offsetHeight / 2);
        list.scrollTo({
            top: Math.max(0, targetTop),
            behavior: 'smooth'
        });

        setTimeout(() => {
            item.classList.remove('tk-item-dup', 'tk-item-new');
        }, 2200);
    }

    function commitProduct(product) {
        if (!product) return;

        // Walidacja liczby stron – usuń brakujące lub nieprawidłowe wartości
        if (product.pages !== undefined && product.pages !== null) {
            const pagesMatch = String(product.pages).match(/\d+/);
            const pagesCount = pagesMatch ? parseInt(pagesMatch[0], 10) : 0;

            product.pages = pagesCount > 0 ? String(pagesCount) : '';
        }

        // Normalizacja nazw autorów – wspólna dla wszystkich księgarni
        if (product.author) {
            product.author = String(product.author)
                .replace(/\s+/g, ' ')
                .replace(/\s*\.(?=\s*,)/g, '')  // usuń błędną kropkę przed przecinkiem
                .replace(/\s+\.$/, '')          // usuń samotną kropkę na końcu
                .replace(/\s+([,.])/g, '$1')
                .replace(/\s*,\s*/g, ', ')
                .trim();
        }

        product.year = sanitizeYear(product.year);
        product.language = normalizeLanguage(product.language);
        product.series = normalizeSeries(product.series);
        if (product.cover) product.cover = absolutizeCoverUrl(product.cover);
        products = GM_getValue('tk_products', []);
        const code = product.ean || product.isbn || '';
        const existing = products.find(p => {
            const c = p.ean || p.isbn || '';
            return code && c && code === c;
        });
        if (existing) {
            toast('Ten produkt jest już na liście');
            focusProductInList(existing.id, { duplicate: true });
            return;
        }
        products.unshift(product);
        saveProducts(product.id);
        toast('Dodano: ' + product.title);
        ensureCoverThumb(product).then(ok => { if (ok) scheduleRender(); });
    }

    function setLoadBusy(busy, label) {
        const btn = ui('#tk-load');
        if (!btn) return;
        btn.disabled = !!busy;
        if (busy) {
            btn.dataset.prevHtml = btn.innerHTML;
            btn.innerHTML = (label || 'Czekam…');
        } else if (btn.dataset.prevHtml) {
            btn.innerHTML = btn.dataset.prevHtml;
            delete btn.dataset.prevHtml;
        }
    }

    /**
     * Empik ładuje atrybuty (EAN, wydawnictwo…) asynchronicznie.
     * Czekamy aż strona będzie „gotowa”, max ~4 s, potem i tak ekstrahujemy.
     */
    function waitForEmpikReady(maxMs = 12000) {
        return new Promise((resolve) => {
            const start = Date.now();
            let baseReadyAt = 0;
            let lastExpand = 0;
            let lastSig = null;
            let sigAt = start;
            const tick = () => {
                const now = Date.now();
                const complete = isEmpikDetailsComplete();
                // rozwijaj, dopóki nie widać końca listy szczegółów
                if (!complete && now - lastExpand > 1200) {
                    expandEmpikDetails();
                    lastExpand = now;
                }
                if (complete || isEmpikDetailsExpanded()) {
                    const sig = empikDetailsSignature();
                    if (sig !== lastSig) { lastSig = sig; sigAt = now; }
                }
                const dataReady = isEmpikDataReady();
                const base = dataReady && complete;
                if (base && !baseReadyAt) baseReadyAt = now;
                const stable = base && now - sigAt >= 700;

                if (stable && getEmpikSuggestedPrice() != null && getEmpikCategories()) return resolve(true);
                // produkt może nie mieć ceny sugerowanej – daj jej (i ścieżce kategorii) 3 s po reszcie danych
                if (stable && now - baseReadyAt >= 3000) return resolve(true);
                // brak wierszy Indeks/EAN w szczegółach – wystarczy rozwinięta i ustabilizowana lista
                if (dataReady && isEmpikDetailsExpanded() && now - start >= 6000 && now - sigAt >= 1500) return resolve(true);
                if (now - start >= maxMs) return resolve(dataReady);
                setTimeout(tick, 250);
            };
            tick();
        });
    }

        let lastLoadAt = 0;
    /** Minimalny odstęp między „Załaduj” – ogranicza wpadanie w anti-bot (Tantis itd.) */
    const LOAD_COOLDOWN_MS = {
        'Tantis': 3500,
        'Empik': 1500,
        'default': 2000
    };

    function isAntiBotPage() {
        const t = ((document.body && document.body.innerText) || '').slice(0, 8000);
        return /Weryfikacja bezpieczeństwa|nietypowo dużo żądań|nie jesteś robotem|potwierdź, że nie jesteś|cf-browser-verification|challenge-platform|Just a moment…|Checking your browser|Access denied|are you a robot/i.test(t);
    }

    async function addCurrentProduct() {
        const site = getCurrentSite();
        if (!site) { toast('Ta strona nie jest obsługiwana'); return; }

        if (isAntiBotPage()) {
            toast('Sklep wykrył zbyt wiele żądań. Odśwież stronę, potwierdź weryfikację i odczekaj chwilę przed kolejnym „Załaduj”.', 5500);
            return;
        }

        if (!site.isProductPage()) { toast('To nie jest strona produktu (' + site.name + ')'); return; }

        // cooldown między załadowaniami (szczególnie przy wielu kartach)
        const cooldown = LOAD_COOLDOWN_MS[site.name] || LOAD_COOLDOWN_MS.default;
        const wait = cooldown - (Date.now() - lastLoadAt);
        if (wait > 0) {
            setLoadBusy(true, 'Pauza ' + Math.ceil(wait / 1000) + 's…');
            try {
                await new Promise(r => setTimeout(r, wait));
            } finally {
                setLoadBusy(false);
            }
        }
        lastLoadAt = Date.now();

        // Empik: rozwiń szczegóły + poczekaj na atrybuty / EAN
        if (site.name === 'Empik') {
            setLoadBusy(true, 'Czekam na dane…');
            try {
                await waitForEmpikReady(12000);
            } finally {
                setLoadBusy(false);
            }
        }

        let product = null;
        try {
            product = site.extract();
        } catch (err) {
            console.error('[MultiKsiegarnia] Błąd adaptera ' + site.name, err);
            toast('Błąd odczytu strony (' + site.name + '): ' + (err && err.message ? err.message : err), 4000);
            return;
        }
        if (!product) return;

        // Empik: doczytaj, jeśli brakuje EAN, ceny lub języka (mogły się nie zdążyć załadować)
        if (site.name === 'Empik') {
            const score = (p) => [p.ean || p.isbn, p.cenaBrutto, p.language, p.pages, p.categories]
                .filter(Boolean).length;
            const missingKey = (p) => !(p.ean || p.isbn) || !p.cenaBrutto || !p.language;
            if (missingKey(product)) {
                setLoadBusy(true, 'Doczytuję dane…');
                try {
                    for (let i = 0; i < 3 && missingKey(product); i++) {
                        await new Promise(r => setTimeout(r, 1200));
                        expandEmpikDetails();
                        const again = site.extract();
                        if (again && score(again) >= score(product)) product = again;
                    }
                } finally { setLoadBusy(false); }
            }
        }

        commitProduct(product);
    }

    function removeProduct(id) {
        products = GM_getValue('tk_products', []);
        products = products.filter(p => p.id !== id);
        dropCovers([id]);
        saveProducts();
        toast('Usunięto');
    }

    function clearAll() {
        products = GM_getValue('tk_products', []);
        if (products.length === 0) return;
        if (!confirm('Na pewno usunąć wszystkie produkty z listy?')) return;
        products = [];
        GM_setValue('tk_covers', {}); covers = {};
        saveProducts();
        toast('Wyczyszczono listę');
    }

    function updateProductField(productId, field, value) {
        products = GM_getValue('tk_products', []);
        const idx = products.findIndex(p => p.id === productId);
        if (idx === -1) return;

        let v = value;
        // Normalizacja cen: zamień przecinek na kropkę, usuń "zł"
        if (field === 'cenaBrutto' || field === 'cenaNetto') {
            v = String(v).replace(/\s*zł\s*/gi, '').replace(',', '.').trim();
            const num = parseFloat(v);
            v = isNaN(num) ? '' : num.toFixed(2);
        }
        if (field === 'vat') {
            v = String(v).replace(/%/g, '').trim();
        }
        if (field === 'binding') {
            v = normalizeBinding(v);
        }
        if (field === 'year') {
            v = sanitizeYear(v);
        }
        if (field === 'language') {
            v = normalizeLanguage(v);
        }
        if (field === 'series') {
            v = normalizeSeries(v);
        }

        products[idx][field] = v;
        GM_setValue('tk_products', products);
        renderList();
    }

    function exportFiszka() {
        products = GM_getValue('tk_products', []);
        if (products.length === 0) { toast('Brak produktów do eksportu'); return; }
        const wb = XLSX.utils.book_new();
        products.forEach((p, idx) => {
            const bindingCode = bindingToMK_TW(p.binding);
            const values = {
                'PEŁNY TYTUŁ I PODTYTUŁ': p.title || '',
                'DATA PREMIERY (DZIEŃ/MIESIĄC/ROK)': formatPremiereFiszka(p.premiereDate),
                'INDEKS GOLD': '',
                'CENA BRUTTO': p.cenaBrutto ? p.cenaBrutto.replace('.', ',') : '',
                'VAT': p.vat ? p.vat + '%' : '',
                'CENA NETTO': p.cenaNetto ? p.cenaNetto.replace('.', ',') : '',
                'CENA DRUKOWANA TAK/NIE': 'NIE',
                'RODZAJ PRODUKTU (KSIĄŻKA / AUDIOBOOK)': 'KSIĄŻKA',
                'ROK WYDANIA': p.year || '',
                'JĘZYK WYDANIA': p.language || 'Polski',
                'JĘZYK ORYGINAŁU': p.language || 'Polski',
                'GRUPA TEMATYCZNA': p.categories || '',
                'PRZEDZIAŁ WIEKOWY': '',
                'WERSJA ELEKTRONICZNA (EBOOK) TAK/NIE': 'NIE',
                'WERSJA AUDIO TAK/NIE': 'NIE',
                'LEKTOR (W PRZYPADKU AUDIOBOOKA)': '',
                'AUTOR I': p.author || '',
                'AUTOR II': '',
                'AUTOR III': '',
                'WYDAWCA': p.publisher || '',
                'TŁUMACZ I': p.translator || '',
                'TŁUMACZ II': '',
                'TYTUŁ ORYGINAŁU': p.originalTitle || p.title || '',
                'ISBN': p.ean || p.isbn || '',
                'EAN': p.ean || p.isbn || '',
                'NUMER WYDANIA': p.edition || '1',
                'NAZWA CYKLU/SERII': p.series || '',
                'OPRAWA (MK/TW)': bindingCode,
                'ILOŚĆ STRON': p.pages || '',
                'OPIS': ''
            };
            const data = FISZKA_FIELDS.map(label => [label, values[label] || '']);
            const ws = XLSX.utils.aoa_to_sheet([['FORMULARZ DANYCH PRODUKTU: KSIĄŻKA/AUDIOBOOK', ''], ...data]);
            ws['!cols'] = [{ wch: 45 }, { wch: 50 }];
            let sheetName = (p.title || `Produkt ${idx + 1}`).replace(/[:\\\/\?\*\[\]]/g, '').replace(/\s+/g, ' ').trim().substring(0, 28) || `Produkt_${idx + 1}`;
            let finalName = sheetName;
            let counter = 1;
            while (wb.SheetNames.includes(finalName)) {
                const suffix = `_${counter}`;
                finalName = sheetName.substring(0, 31 - suffix.length) + suffix;
                counter++;
            }
            XLSX.utils.book_append_sheet(wb, ws, finalName);
        });
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `fiszki_${new Date().toISOString().slice(0,10)}.xlsx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        toast(`Wygenerowano ${products.length} fiszek`);
    }

    function copyToClipboard() {
        products = GM_getValue('tk_products', []);
        if (products.length === 0) { toast('Brak produktów do skopiowania'); return; }
        // Kolumny TSV (0-based): … AI = indeks 34 → data premiery; na końcu ścieżka (bez domeny/źródła)
        const lines = products.map(p => [
            p.ean || p.isbn || '', p.title || '', '', p.series || '', p.publisher || '',
            '', '', '', '', '', '', '', '', '', '', '',
            p.vat || '',
            p.cenaNetto ? p.cenaNetto.replace('.', ',') : '',
            p.cenaBrutto ? p.cenaBrutto.replace('.', ',') : '',
            '', '', '', p.year || '', '', p.author || '', p.pages || '',
            normalizeBinding(p.binding || ''), p.language || '',
            // 28–33 puste
            '', '', '', '', '', '',
            // 34 = AI → data premiery
            formatPremiereDisplay(p.premiereDate) || '',
            // 35–42 puste
            '', '', '', '', '', '', '', '',
            // 43 ścieżka kategorii (bez domeny / source)
            p.categories || ''
        ].join('\t'));
        const text = lines.join('\n');
        if (typeof GM_setClipboard === 'function') {
            GM_setClipboard(text, 'text');
            toast('Skopiowano ' + products.length + ' produktów');
        } else {
            navigator.clipboard.writeText(text).then(() => toast('Skopiowano ' + products.length + ' produktów'))
                .catch(() => {
                    const ta = document.createElement('textarea');
                    ta.value = text;
                    ta.style.position = 'fixed';
                    ta.style.left = '-9999px';
                    document.body.appendChild(ta);
                    ta.select();
                    document.execCommand('copy');
                    document.body.removeChild(ta);
                    toast('Skopiowano ' + products.length + ' produktów');
                });
        }
    }

    // ========== EDYCJA ==========
    function startEdit(wrapEl, productId, field, currentValue) {
        (uiRoot ? uiRoot.querySelectorAll('.tk-edit-input') : document.querySelectorAll('.tk-edit-input')).forEach(inp => {
            const parent = inp.closest('[data-editing]');
            if (parent && parent !== wrapEl) {
                const cancelBtn = parent.querySelector('.tk-edit-action.cancel');
                if (cancelBtn) cancelBtn.click();
            }
        });

        const displayValue = currentValue || '';
        wrapEl.setAttribute('data-editing', '1');
        wrapEl.innerHTML = `
            <input class="tk-edit-input" type="text" value="${escapeHtml(displayValue)}" />
            <div class="tk-edit-actions">
                <button class="tk-edit-action save" title="Zapisz (Enter)">✓</button>
                <button class="tk-edit-action clear" title="Wyczyść">⊘</button>
                <button class="tk-edit-action cancel" title="Anuluj (Esc)">✕</button>
            </div>
        `;

        const input = wrapEl.querySelector('.tk-edit-input');
        const btnSave = wrapEl.querySelector('.tk-edit-action.save');
        const btnClear = wrapEl.querySelector('.tk-edit-action.clear');
        const btnCancel = wrapEl.querySelector('.tk-edit-action.cancel');
        input.focus();
        input.select();

        let finished = false;
        const finish = (newVal, cancelled) => {
            if (finished) return;
            finished = true;
            document.removeEventListener('mousedown', onOutside);
            if (!cancelled) {
                updateProductField(productId, field, newVal);
            } else {
                renderList();
            }
        };

        btnSave.addEventListener('click', (e) => { e.stopPropagation(); finish(input.value.trim(), false); });
        btnClear.addEventListener('click', (e) => { e.stopPropagation(); input.value = ''; input.focus(); });
        btnCancel.addEventListener('click', (e) => { e.stopPropagation(); finish(null, true); });

        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') { e.preventDefault(); finish(input.value.trim(), false); }
            else if (e.key === 'Escape') { e.preventDefault(); finish(null, true); }
        });

        const onOutside = (e) => {
            const path = e.composedPath ? e.composedPath() : [e.target];
            if (!path.includes(wrapEl)) {
                if (document.body.contains(input)) finish(input.value.trim(), false);
            }
        };
        setTimeout(() => document.addEventListener('mousedown', onOutside), 0);
    }

    // ========== SIATKA (przegląd wielu produktów + kontrola braków) ==========
    // Pola, których brak oznaczamy jako „brak” (edytuj listę wg potrzeb eksportu).
    const REQUIRED_FIELDS = [
        ['author', 'Autor'], ['publisher', 'Wydawnictwo'], ['ean', 'EAN'],
        ['cenaBrutto', 'Cena brutto'], ['cenaNetto', 'Cena netto'], ['vat', 'VAT'],
        ['year', 'Rok'], ['premiereDate', 'Premiera'], ['pages', 'Strony'],
        ['binding', 'Oprawa'], ['categories', 'Ścieżka']
    ];
    let gridOnlyMissing = false;

    function fieldRaw(p, field) {
        const v = field === 'ean' ? (p.ean || p.isbn) : p[field];
        return v == null ? '' : String(v).trim();
    }

    function productMissing(p) {
        return REQUIRED_FIELDS.filter(([f]) => !fieldRaw(p, f)).map(([, label]) => label);
    }

    function closeGrid() {
        const m = ui('.tk-grid-modal');
        if (m) m.remove();
    }

    function gridRow(icon, label, display, p, field) {
        const required = REQUIRED_FIELDS.some(([f]) => f === field);
        const missing = required && !fieldRaw(p, field);
        return `
            <div class="tk-item-row${missing ? ' tk-missing' : ''}">
                ${icon}
                <span class="tk-item-label">${label}</span>
                <div class="tk-item-value-wrap" data-id="${p.id}" data-field="${field}">
                    <span class="tk-item-value">${display ? escapeHtml(display) : '<span class="tk-dash">—</span>'}</span>
                    <button class="tk-edit-btn" title="Edytuj">${ICONS.pencil}</button>
                </div>
            </div>`;
    }

    function renderGrid() {
        const modal = ui('.tk-grid-modal');
        if (!modal) return;
        const body = modal.querySelector('.tk-grid-body');
        const prevScroll = body.scrollTop;

        // statystyka braków – dla całej listy (eksport obejmuje wszystko)
        const byField = {};
        let incomplete = 0;
        products.forEach(p => {
            const miss = productMissing(p);
            if (miss.length) incomplete++;
            miss.forEach(m => { byField[m] = (byField[m] || 0) + 1; });
        });
        const summary = modal.querySelector('.tk-grid-summary');
        if (!products.length) {
            summary.innerHTML = '';
        } else if (!incomplete) {
            summary.innerHTML = `<span class="tk-gchip ok">✓ Wszystkie ${products.length} pozycji kompletne</span>`;
        } else {
            summary.innerHTML = `<span>Z brakami: <b>${incomplete}</b> z ${products.length}.</span>` +
                Object.entries(byField).map(([k, n]) => `<span class="tk-gchip">${escapeHtml(k)}: ${n}</span>`).join('');
        }

        let visible = getFilteredProducts();
        if (gridOnlyMissing) visible = visible.filter(p => productMissing(p).length);
        modal.querySelector('.tk-grid-count').textContent =
            `Widoczne ${visible.length} z ${products.length}`;
        const sb = modal.querySelector('.tk-search-box');
        if (sb) sb.classList.toggle('active', !!filterQuery.trim());

        if (!visible.length) {
            body.innerHTML = `<div class="tk-grid-empty">${products.length
                ? 'Brak pozycji spełniających filtr.' : 'Brak zapisanych produktów.'}</div>`;
            return;
        }

        body.innerHTML = visible.map(p => {
            const miss = productMissing(p);
            const thumb = covers[p.id] || '';
            const coverHtml = thumb
                ? `<img class="tk-gcover" src="${thumb}" alt="" />`
                : `<div class="tk-gcover-ph">${ICONS.image}</div>`;
            const linkHtml = p.url
                ? `<a class="tk-ext-link" href="${escapeHtml(p.url)}" target="_blank" rel="noopener noreferrer" title="Otwórz stronę produktu">${ICONS.external}</a>`
                : '';
            const badge = miss.length
                ? `<span class="tk-gbadge bad" title="${escapeHtml('Brakuje: ' + miss.join(', '))}">Braki: ${miss.length}</span>`
                : `<span class="tk-gbadge ok">Komplet</span>`;
            return `
            <div class="tk-gcard${miss.length ? ' tk-gcard-incomplete' : ''}" data-id="${p.id}">
                <button class="tk-item-remove" title="Usuń">×</button>
                <div class="tk-gcard-head">
                    ${coverHtml}
                    <div class="tk-gcard-headtext">
                        <div class="tk-item-top">
                            <span class="tk-item-source tk-source-${p.source || 'unknown'}">${escapeHtml(sourceLabel(p.source))}</span>
                            ${linkHtml}${badge}
                        </div>
                        <div class="tk-item-title" data-id="${p.id}" data-field="title">
                            <span class="tk-item-title-text">${escapeHtml(p.title)}<button class="tk-edit-btn" title="Edytuj tytuł">${ICONS.pencil}</button></span>
                        </div>
                    </div>
                </div>
                ${gridRow(ICONS.user, 'Autor', p.author, p, 'author')}
                ${gridRow(ICONS.tag, 'Wydawnictwo', p.publisher, p, 'publisher')}
                ${gridRow(ICONS.layers, 'Seria', p.series, p, 'series')}
                ${gridRow(ICONS.money, 'Cena brutto', p.cenaBrutto ? p.cenaBrutto.replace('.', ',') + ' zł' : '', p, 'cenaBrutto')}
                ${gridRow(ICONS.money, 'Cena netto', p.cenaNetto ? p.cenaNetto.replace('.', ',') + ' zł' : '', p, 'cenaNetto')}
                ${gridRow(ICONS.percent, 'VAT', p.vat ? p.vat + '%' : '', p, 'vat')}
                ${gridRow(ICONS.calendar, 'Rok', p.year, p, 'year')}
                ${gridRow(ICONS.calendar, 'Premiera', formatPremiereDisplay(p.premiereDate), p, 'premiereDate')}
                ${gridRow(ICONS.barcode, 'EAN', p.ean || p.isbn || '', p, 'ean')}
                ${gridRow(ICONS.book, 'Strony', p.pages, p, 'pages')}
                ${gridRow(ICONS.layers, 'Oprawa', normalizeBinding(p.binding || ''), p, 'binding')}
                ${gridRow(ICONS.lang, 'Język', p.language, p, 'language')}
                ${gridRow(ICONS.folder, 'Ścieżka', p.categories, p, 'categories')}
            </div>`;
        }).join('');

        body.querySelectorAll('.tk-item-remove').forEach(btn => {
            btn.addEventListener('click', () => removeProduct(btn.closest('.tk-gcard').dataset.id));
        });
        body.querySelectorAll('img.tk-gcover').forEach(img => {
            img.addEventListener('click', () => {
                const id = img.closest('.tk-gcard').dataset.id;
                const product = products.find(p => p.id === id);
                openLightbox(img.src, product ? product.title : '', product && product.cover);
            });
        });
        body.querySelectorAll('.tk-item-value-wrap .tk-edit-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const wrap = btn.closest('.tk-item-value-wrap');
                const product = products.find(p => p.id === wrap.dataset.id);
                if (!product) return;
                const field = wrap.dataset.field;
                startEdit(wrap, product.id, field, field === 'ean' ? (product.ean || product.isbn || '') : (product[field] || ''));
            });
        });
        body.querySelectorAll('.tk-item-title .tk-edit-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const titleEl = btn.closest('.tk-item-title');
                const product = products.find(p => p.id === titleEl.dataset.id);
                if (!product) return;
                startEdit(titleEl, product.id, 'title', product.title || '');
            });
        });

        body.scrollTop = prevScroll;
    }

    function openGrid() {
        if (!uiRoot) return;
        closeGrid();
        products = GM_getValue('tk_products', products);
        covers = GM_getValue('tk_covers', covers);
        const modal = document.createElement('div');
        modal.className = 'tk-grid-modal';
        if (darkTheme) modal.classList.add('dark');
        modal.innerHTML = `
            <div class="tk-grid-dialog" role="dialog" aria-label="Siatka produktów">
                <div class="tk-grid-top">
                    <div class="tk-grid-title"><h3>Produkty</h3><span class="tk-grid-count"></span></div>
                    <div class="tk-search-box tk-grid-search-box">
                        <span class="tk-search-icon">${ICONS.search}</span>
                        <input type="search" class="tk-search" id="tk-grid-search" placeholder="Filtruj listę (tytuł, autor, EAN…)" value="${escapeHtml(filterQuery)}" />
                    </div>
                    <label class="tk-grid-toggle"><input type="checkbox" id="tk-grid-missing" /> Tylko z brakami</label>
                    <button type="button" class="tk-grid-close" title="Zamknij (Esc)">×</button>
                </div>
                <div class="tk-grid-summary"></div>
                <div class="tk-grid-body"></div>
            </div>`;
        modal.addEventListener('click', (e) => {
            if (e.target === modal || e.target.closest('.tk-grid-close')) closeGrid();
        });
        uiRoot.appendChild(modal);
        const gs = modal.querySelector('#tk-grid-search');
        gs.addEventListener('input', (e) => setFilter(e.target.value));
        const gm = modal.querySelector('#tk-grid-missing');
        gm.checked = gridOnlyMissing;
        gm.addEventListener('change', () => { gridOnlyMissing = gm.checked; renderGrid(); });
        renderGrid();
    }

    // ========== LIGHTBOX ==========
    function closeLightbox() {
        const lb = ui('.tk-lightbox');
        if (lb) lb.remove();
    }

    function openLightbox(src, caption, fullUrl) {
        if (!uiRoot || !src) return;
        closeLightbox();
        const lb = document.createElement('div');
        lb.className = 'tk-lightbox';
        lb.innerHTML = `
            <button class="tk-lightbox-close" title="Zamknij (Esc)">×</button>
            <img src="${escapeHtml(src)}" alt="" />
            ${caption ? `<div class="tk-lightbox-caption">${escapeHtml(caption)}</div>` : ''}
        `;
        // klik w tło lub X zamyka; klik w samo zdjęcie nie
        lb.addEventListener('click', (e) => {
            if (e.target.tagName !== 'IMG') closeLightbox();
        });
        uiRoot.appendChild(lb);

        // pełny obraz w tle; przy porażce zostaje miniatura
        if (fullUrl && /^https?:/i.test(fullUrl)) {
            fetchImageBlob(fullUrl).then(blobToDataUrl).then(d => {
                const probe = new Image();
                probe.onload = () => {
                    const im = lb.querySelector('img');
                    if (lb.isConnected && im) im.src = d;
                };
                probe.src = d;
            }).catch(() => {});
        }
    }

    // ========== HELP MODAL ==========
    function closeHelpModal() {
        const m = ui('.tk-help-modal');
        if (m) m.remove();
    }

    function openHelpModal() {
        if (!uiRoot) return;
        closeHelpModal();
        closeLightbox();
        // unikalne nazwy z SITES
        const names = [];
        const seen = new Set();
        Object.values(SITES).forEach(s => {
            if (s && s.name && !seen.has(s.name)) {
                seen.add(s.name);
                names.push(s.name);
            }
        });
        const sourceKey = {
            'TaniaKsiążka': 'taniaksiazka',
            'Platon': 'platon',
            'Bonito': 'bonito',
            'Tantis': 'tantis',
            'Świat Książki': 'swiatksiazki',
            'Dressler': 'dressler',
            'Empik': 'empik',
            'Ateneum': 'ateneum'
        };
        const listHtml = names.map(n => {
            const key = sourceKey[n] || n.toLowerCase().replace(/\s+/g, '');
            return `<li><span class="tk-item-source tk-source-${key}">${escapeHtml(n)}</span><span>${escapeHtml(n)}</span></li>`;
        }).join('');

        const modal = document.createElement('div');
        modal.className = 'tk-help-modal';
        // dziedzicz motyw panelu
        if (darkTheme) modal.classList.add('dark');
        const version = (typeof GM_info !== 'undefined' && GM_info.script && GM_info.script.version)
            ? GM_info.script.version
            : '4.11.2';
        modal.innerHTML = `
            <div class="tk-help-dialog" role="dialog" aria-label="Pomoc">
                <button class="tk-help-close" title="Zamknij (Esc)">×</button>
                <div class="tk-help-grid">
                    <div class="tk-help-col">
                        <h3>Obsługiwane strony</h3>
                        <p>Wejdź na stronę produktu w jednym z poniższych sklepów i kliknij „Załaduj”.</p>
                        <ul class="tk-help-list">${listHtml}</ul>
                        <h3 class="tk-help-section">Przyciski panelu</h3>
                        <ul class="tk-help-actions">
                            <li>
                                <span class="tk-help-btn-name"><span class="tk-help-pill tk-help-pill-load">Załaduj</span></span>
                                <span class="tk-help-btn-desc">Pobiera dane produktu z aktualnej strony i dodaje go do listy (tytuł, autor, ceny, EAN, okładka itd.).</span>
                            </li>
                            <li>
                                <span class="tk-help-btn-name"><span class="tk-help-pill tk-help-pill-copy">Kopiuj</span></span>
                                <span class="tk-help-btn-desc">Kopiuje listę w formacie TSV (tabulatory) – do wklejenia w Excel / Google Sheets.</span>
                            </li>
                            <li>
                                <span class="tk-help-btn-name"><span class="tk-help-pill tk-help-pill-fiszka">Fiszki</span></span>
                                <span class="tk-help-btn-desc">Excel z osobnym arkuszem (fiszką) dla każdego produktu.</span>
                            </li>
                            <li>
                                <span class="tk-help-btn-name"><span class="tk-help-pill tk-help-pill-grid">Siatka</span></span>
                                <span class="tk-help-btn-desc">Otwiera duży podgląd (95% strony) z kartami wszystkich produktów. Brakujące pola są podświetlone na czerwono, można je edytować w miejscu.</span>
                            </li>
                            <li>
                                <span class="tk-help-btn-name"><span class="tk-help-pill tk-help-pill-clear">Wyczyść</span></span>
                                <span class="tk-help-btn-desc">Usuwa wszystkie produkty z listy (po potwierdzeniu).</span>
                            </li>
                        </ul>
                    </div>
                    <div class="tk-help-col tk-help-col-about">
                        <h3>About – co robi aplikacja</h3>
                        <p>MultiKsięgarnia to panel boczny (userscript) do zbierania danych produktowych ze sklepów i hurtowni książkowych.</p>
                        <ul class="tk-help-about-list">
                            <li>Scrapuje kartę produktu: tytuł, autor, wydawnictwo, seria, EAN, strony, oprawa, język, ścieżka kategorii, data premiery, rok.</li>
                            <li>Pobiera cenę brutto (preferuje cenę sugerowaną przez wydawcę), liczy netto i ustala VAT.</li>
                            <li>VAT 5% dla książek (ścieżka / heurystyka); jawna stawka ze strony ma pierwszeństwo (np. 23% zabawki).</li>
                            <li>Zapisuje miniaturę okładki; przy blokadzie cross-origin (np. na Ateneum) dociąga ją przez Tampermonkey.</li>
                            <li>Lista produktów jest wspólna między sklepami (synchronizacja w przeglądarce).</li>
                            <li>Edycja pól w panelu, filtr listy, podgląd okładki (lightbox), motyw jasny/ciemny.</li>
                            <li>Siatka produktów z podświetleniem braków przed eksportem.</li>
                            <li>Deduplikacja po EAN – ten sam produkt nie doda się drugi raz.</li>
                            <li>Eksport: schowek TSV oraz fiszki produktowe (XLSX).</li>
                            <li>Na Empiku czeka na dociągnięcie szczegółów (EAN, język, wymiary) przed odczytem.</li>
                            <li>Rok wydania filtrowany (ostatnie 100 lat + rozsądna przyszłość).</li>
                        </ul>
                        <div class="tk-help-version">Wersja skryptu: ${escapeHtml(version)}</div>
                    </div>
                </div>
            </div>
        `;
        modal.addEventListener('click', (e) => {
            if (e.target === modal || e.target.classList.contains('tk-help-close')) {
                closeHelpModal();
            }
        });
        // zastosuj zmienne CSS motywu (dialog używa var(--bg) itd.)
        const panel = ui('#tk-panel');
        if (panel) {
            modal.style.setProperty('--bg', getComputedStyle(panel).getPropertyValue('--bg'));
            modal.style.setProperty('--text', getComputedStyle(panel).getPropertyValue('--text'));
            modal.style.setProperty('--text-title', getComputedStyle(panel).getPropertyValue('--text-title'));
            modal.style.setProperty('--text-muted', getComputedStyle(panel).getPropertyValue('--text-muted'));
            modal.style.setProperty('--border', getComputedStyle(panel).getPropertyValue('--border'));
            modal.style.setProperty('--edit-bg', getComputedStyle(panel).getPropertyValue('--edit-bg'));
        }
        uiRoot.appendChild(modal);
    }

    // Esc zamyka modal / lightbox (capture, żeby nie kolidowało z Esc w edycji pól)
    document.addEventListener('keydown', (e) => {
        if (e.key !== 'Escape') return;
        if (ui('.tk-help-modal')) {
            e.stopPropagation();
            closeHelpModal();
            return;
        }
        if (ui('.tk-lightbox')) {
            e.stopPropagation();
            closeLightbox();
            return;
        }
        if (ui('.tk-grid-modal') && !ui('.tk-edit-input')) {
            // Esc w niepustym polu szukania ma je tylko wyczyścić
            const inner = (e.composedPath && e.composedPath()[0]) || e.target;
            if (inner && inner.id === 'tk-grid-search' && inner.value) return;
            e.stopPropagation();
            closeGrid();
        }
    }, true);

    function updateCount() {
        const el = ui('.tk-count');
        if (!el) return;
        const active = !!filterQuery.trim();
        el.textContent = active ? `${getFilteredProducts().length}/${products.length}` : products.length;
        el.classList.toggle('filtered', active);
        el.title = active ? 'Widoczne / wszystkie (filtr aktywny)' : 'Liczba produktów';
    }

    function row(icon, label, value, productId, field) {
        const display = value || '';
        return `
            <div class="tk-item-row">
                ${icon}
                <span class="tk-item-label">${label}</span>
                <div class="tk-item-value-wrap" data-id="${productId}" data-field="${field}">
                    <span class="tk-item-value">${display ? escapeHtml(display) : '<span style="opacity:0.35">—</span>'}</span>
                    <button class="tk-edit-btn" title="Edytuj">${ICONS.pencil}</button>
                </div>
            </div>`;
    }

    function getFilteredProducts() {
        if (!filterQuery.trim()) return products;
        const q = filterQuery.trim().toLowerCase();
        return products.filter(p => {
            const hay = [
                p.title, p.author, p.publisher, p.ean || p.isbn,
                p.series, p.categories, p.year, p.language,
                sourceLabel(p.source)
            ].join(' ').toLowerCase();
            return hay.includes(q);
        });
    }

    function renderList(highlightId = null) {
        if (ui('.tk-grid-modal')) renderGrid();
        const list = ui('.tk-list');
        if (!list) return;

        const filtered = getFilteredProducts();
        const active = !!filterQuery.trim();
        const infoEl = ui('.tk-filter-info');
        if (infoEl) {
            infoEl.classList.toggle('on', active);
            infoEl.innerHTML = active
                ? `<span class="tk-filter-text">Filtr aktywny: „${escapeHtml(filterQuery.trim())}” · widoczne ${filtered.length} z ${products.length}` +
                  `<span class="tk-filter-note">Kopiuj i Fiszki obejmują całą listę.</span></span>` +
                  `<button type="button" class="tk-filter-clear">Wyczyść filtr</button>`
                : '';
        }
        const boxEl = ui('.tk-search-box');
        if (boxEl) boxEl.classList.toggle('active', active);
        updateCount();

        if (products.length === 0) {
            list.innerHTML = `<div class="tk-empty">Brak zapisanych produktów.<br>Wejdź na stronę produktu<br>i kliknij „Załaduj”.</div>`;
            return;
        }

        if (filtered.length === 0) {
            list.innerHTML = `<div class="tk-empty">Brak wyników dla „${escapeHtml(filterQuery)}”.</div>`;
            return;
        }

        list.innerHTML = filtered.map(p => {
            const sourceClass = 'tk-source-' + (p.source || 'unknown');
            const thumb = covers[p.id] || '';
            const coverHtml = thumb
                ? `<img class="tk-cover" src="${thumb}" alt="" />`
                : `<div class="tk-cover-placeholder">${ICONS.image}</div>`;
            const linkHtml = p.url
                ? `<a class="tk-ext-link" href="${escapeHtml(p.url)}" target="_blank" rel="noopener noreferrer" title="Otwórz stronę produktu">${ICONS.external}</a>`
                : '';

            return `
            <div class="tk-item${p.id === highlightId ? ' tk-item-new' : ''}" data-id="${p.id}">
                <button class="tk-item-remove" title="Usuń">×</button>
                <div class="tk-item-header">
                    ${coverHtml}
                    <div class="tk-item-header-text">
                        <div class="tk-item-top">
                            <span class="tk-item-source ${sourceClass}">${escapeHtml(sourceLabel(p.source))}</span>
                            ${linkHtml}
                        </div>
                        <div class="tk-item-title" data-id="${p.id}" data-field="title">
                            <span class="tk-item-title-text">${escapeHtml(p.title)}<button class="tk-edit-btn" title="Edytuj tytuł">${ICONS.pencil}</button></span>
                        </div>
                    </div>
                </div>
                <div class="tk-item-fields">
                    ${row(ICONS.user, 'Autor', p.author, p.id, 'author')}
                    ${row(ICONS.tag, 'Wydawnictwo', p.publisher, p.id, 'publisher')}
                    ${row(ICONS.layers, 'Seria', p.series, p.id, 'series')}
                    ${row(ICONS.money, 'Cena brutto', p.cenaBrutto ? p.cenaBrutto.replace('.', ',') + ' zł' : '', p.id, 'cenaBrutto')}
                    ${row(ICONS.money, 'Cena netto', p.cenaNetto ? p.cenaNetto.replace('.', ',') + ' zł' : '', p.id, 'cenaNetto')}
                    ${row(ICONS.percent, 'VAT', p.vat ? p.vat + '%' : '', p.id, 'vat')}
                    ${row(ICONS.calendar, 'Rok', p.year, p.id, 'year')}
                    ${row(ICONS.calendar, 'Premiera', formatPremiereDisplay(p.premiereDate), p.id, 'premiereDate')}
                    ${row(ICONS.barcode, 'EAN', p.ean || p.isbn || '', p.id, 'ean')}
                    ${row(ICONS.book, 'Strony', p.pages, p.id, 'pages')}
                    ${row(ICONS.layers, 'Oprawa', normalizeBinding(p.binding || ''), p.id, 'binding')}
                    ${row(ICONS.lang, 'Język', p.language, p.id, 'language')}
                    ${row(ICONS.folder, 'Ścieżka', p.categories, p.id, 'categories')}
                </div>
            </div>`;
        }).join('');

        list.querySelectorAll('.tk-item-remove').forEach(btn => {
            btn.addEventListener('click', (e) => {
                removeProduct(e.target.closest('.tk-item').dataset.id);
            });
        });

        // Kliknięcie w miniaturkę → podgląd okładki w modalu
        list.querySelectorAll('img.tk-cover').forEach(img => {
            img.addEventListener('click', (e) => {
                e.stopPropagation();
                const id = img.closest('.tk-item').dataset.id;
                const product = products.find(p => p.id === id);
                openLightbox(img.src, product ? product.title : '', product && product.cover);
            });
        });

        list.querySelectorAll('.tk-item-value-wrap .tk-edit-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const wrap = btn.closest('.tk-item-value-wrap');
                const productId = wrap.dataset.id;
                const field = wrap.dataset.field;
                const product = products.find(p => p.id === productId);
                if (!product) return;
                startEdit(wrap, productId, field, product[field] || '');
            });
        });

        list.querySelectorAll('.tk-item-title .tk-edit-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const titleEl = btn.closest('.tk-item-title');
                const productId = titleEl.dataset.id;
                const product = products.find(p => p.id === productId);
                if (!product) return;
                startEdit(titleEl, productId, 'title', product.title || '');
            });
        });

        if (highlightId) {
            const newItem = list.querySelector(`.tk-item[data-id="${CSS.escape(String(highlightId))}"]`);
            if (newItem) {
                // nowe na górze – delikatny scroll + podświetlenie
                newItem.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
                setTimeout(() => newItem.classList.remove('tk-item-new'), 1600);
            }
        }
    }

    function applyTheme() {
        const panel = ui('#tk-panel');
        if (!panel) return;
        panel.classList.toggle('dark', darkTheme);
        const gm = ui('.tk-grid-modal');
        if (gm) gm.classList.toggle('dark', darkTheme);
    }

    function createPanel() {
        if (document.getElementById('tk-root')) return;
        // Host + Shadow DOM → pełna izolacja od CSS strony
        const host = document.createElement('div');
        host.id = 'tk-root';
        Object.assign(host.style, {
            all: 'initial',
            position: 'fixed',
            top: '0',
            left: '0',
            width: '0',
            height: '0',
            overflow: 'visible',
            zIndex: '2147483000',
            pointerEvents: 'none'
        });
        document.body.appendChild(host);
        uiRoot = host.attachShadow({ mode: 'open' });

        const styleEl = document.createElement('style');
        styleEl.textContent = PANEL_CSS + `
            /* Host children need pointer events */
            #tk-panel, #tk-toggle, .tk-toast { pointer-events: auto; }
        `;
        uiRoot.appendChild(styleEl);

        const panel = document.createElement('div');
        panel.id = 'tk-panel';
        if (!panelVisible) panel.classList.add('hidden');
        if (darkTheme) panel.classList.add('dark');

        const site = getCurrentSite();
        const siteName = site ? site.name : 'Multi';

        panel.innerHTML = `
            <div class="tk-header">
                <div class="tk-header-left">
                    <span class="tk-count">${products.length}</span>
                    <h2>Produkty (${siteName})</h2>
                </div>
                <div class="tk-header-right">
                    <span class="tk-theme-label">${ICONS.sun} Motyw</span>
                    <label class="tk-theme-switch" title="Przełącz jasny/ciemny motyw">
                        <input type="checkbox" id="tk-theme-toggle" ${darkTheme ? 'checked' : ''}>
                        <span class="tk-theme-slider"></span>
                    </label>
                    <button type="button" class="tk-help-btn" id="tk-help" title="Obsługiwane strony">${ICONS.help}</button>
                </div>
            </div>
            <div class="tk-actions">
                <button class="tk-btn tk-btn-primary" id="tk-load" title="Załaduj produkt z tej strony">${ICONS.download} Załaduj</button>
                <button class="tk-btn tk-btn-copy" id="tk-copy" title="Kopiuj do schowka (TSV)">${ICONS.copy} Kopiuj</button>
                <button class="tk-btn tk-btn-fiszka" id="tk-fiszka" title="Eksport fiszek">${ICONS.file} Fiszki</button>
                <button class="tk-btn tk-btn-grid" id="tk-grid" title="Podgląd siatki produktów">${ICONS.grid} Siatka</button>
                <button class="tk-btn tk-btn-danger" id="tk-clear" title="Wyczyść listę">${ICONS.trash} Wyczyść</button>
            </div>
            <div class="tk-search-wrap">
                <div class="tk-search-box">
                    <span class="tk-search-icon">${ICONS.search}</span>
                    <input type="search" class="tk-search" id="tk-search" placeholder="Filtruj listę (tytuł, autor, EAN…)" value="${escapeHtml(filterQuery)}" />
                </div>
                <div class="tk-filter-info"></div>
            </div>
            <div class="tk-list"></div>
        `;

        const toggle = document.createElement('button');
        toggle.id = 'tk-toggle';
        toggle.textContent = panelVisible ? 'Ukryj panel' : 'Pokaż panel';
        toggle.title = 'Pokaż / ukryj panel';

        uiRoot.appendChild(panel);
        uiRoot.appendChild(toggle);

        ui('#tk-load').addEventListener('click', addCurrentProduct);
        ui('#tk-copy').addEventListener('click', copyToClipboard);
        ui('#tk-fiszka').addEventListener('click', exportFiszka);
        ui('#tk-grid').addEventListener('click', openGrid);
        ui('#tk-clear').addEventListener('click', clearAll);

        ui('#tk-theme-toggle').addEventListener('change', (e) => {
            darkTheme = e.target.checked;
            GM_setValue('tk_dark_theme', darkTheme);
            applyTheme();
        });

        ui('#tk-help').addEventListener('click', (e) => {
            e.stopPropagation();
            openHelpModal();
        });

        ui('#tk-search').addEventListener('input', (e) => {
            setFilter(e.target.value);
        });
        ui('.tk-search-wrap').addEventListener('click', (e) => {
            if (e.target.closest('.tk-filter-clear')) {
                setFilter('');
                const inp = ui('#tk-search');
                if (inp) inp.focus();
            }
        });

        toggle.addEventListener('click', () => {
            panelVisible = !panelVisible;
            panel.classList.toggle('hidden', !panelVisible);
            toggle.textContent = panelVisible ? 'Ukryj panel' : 'Pokaż panel';
            GM_setValue('tk_panel_visible', panelVisible);
        });

        renderList();
        migrateLanguages();
        setTimeout(backfillCovers, 1500);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', createPanel);
    } else {
        createPanel();
    }
})();