// ==UserScript==
// @name         Adriana Matos Advocacia — Sub-Inbox Gmail
// @namespace    https://adrianamatos.adv.br/
// @version      1.3.2
// @description  Painel de sub-inboxes por município injetado no Gmail — Adriana Matos Advocacia
// @author       Carlos Daniel
// @match        https://mail.google.com/*
// @match        https://mail.google.com/mail/*
// @updateURL    https://raw.githubusercontent.com/escritorioadrianamatosadv-dev/am-userscript/main/script.user.js
// @downloadURL  https://raw.githubusercontent.com/escritorioadrianamatosadv-dev/am-userscript/main/script.user.js
// @grant        GM_xmlhttpRequest
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        unsafeWindow
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @connect      googleapis.com
// @connect      *
// @run-at       document-end
// ==/UserScript==

(function () {
  'use strict';
  const SCRIPT_RAW_URL = 'https://raw.githubusercontent.com/escritorioadrianamatosadv-dev/am-userscript/main/script.user.js';
const VERSAO_ATUAL = GM_info.script.version;

GM_xmlhttpRequest({
  method: 'GET',
  url: SCRIPT_RAW_URL + '?_t=' + Date.now(),
  onload: function(res) {
    const match = res.responseText.match(/@version\s+([\d.]+)/);
    if (!match) return;
    if (match[1] === VERSAO_ATUAL) return;
    const banner = document.createElement('div');
    banner.style.cssText = 'position:fixed;bottom:20px;right:20px;background:#1c1c1e;color:#fff;padding:14px 18px;border-radius:12px;z-index:99999;font-size:13px;font-family:-apple-system,sans-serif;box-shadow:0 8px 30px rgba(0,0,0,0.4);border:1px solid rgba(52,199,89,0.3);display:flex;align-items:center;gap:12px;';
    banner.innerHTML = `<span style="color:#34c759;font-size:18px;">↑</span><span>Nova versão! <b>${VERSAO_ATUAL} → ${match[1]}</b></span><a href="${SCRIPT_RAW_URL}" target="_blank" style="background:#34c759;color:#fff;padding:5px 12px;border-radius:8px;text-decoration:none;font-weight:700;font-size:12px;">Atualizar</a><span style="cursor:pointer;color:#666;font-size:16px;" onclick="this.parentElement.remove()">✕</span>`;
    document.body.appendChild(banner);
    setTimeout(() => { if (banner.parentElement) banner.remove(); }, 15000);
  }
});

  const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwp3qu9wUSx-0Dqa4YhZEdEPBi4AIcbuwn1X7UMNkrvPNz7rokaa1YIA9ZnvIV5SEjA/exec';
    const SYNC_INTERVAL_MS = 15000;
    GM_setValue('subinboxData', null);
    GM_setValue('subinboxes', null);

  const COR_MAP = {
    graphite: '#444444', azul: '#4986e7', teal: '#2da2bb',
    verde: '#16a765',    roxo: '#a479e2', vermelho: '#cc3a21',
    laranja: '#ffad47',  rosa: '#f691b3', ciano: '#b3efd3', dourado: '#c6aa18',
    grafite: '#444444',  sage: '#b3efd3'
  };

  let state = {
    subinboxData: {},
    movimentacoes: [],
    lastSync: null,
    expandidos: {},
    abaAtiva: {}
  };

  let injecting        = false;
  let syncTimer        = null;
  let navTimer         = null;
  let lastURL          = '';
  let lastNonThreadURL = '';

  // ════════════════════════════════════════════════════════════════
  // CSS — PAINEL PRINCIPAL
  // ════════════════════════════════════════════════════════════════
  const CSS = `
#am-root {
  isolation: isolate;
  --sys-bg: #f2f2f7; --sys-bg2: #e5e5ea; --sys-bg3: #d1d1d6;
  --sys-card: #ffffff; --sys-card-blur: rgba(255,255,255,0.82);
  --dark-1: #1c1c1e; --dark-2: #2c2c2e; --dark-3: #3a3a3c; --dark-4: #48484a;
  --dark-sep: rgba(255,255,255,0.08);
  --dt-1: #ffffff; --dt-2: rgba(255,255,255,0.75); --dt-3: rgba(255,255,255,0.45); --dt-4: rgba(255,255,255,0.25);
  --lt-1: #1c1c1e; --lt-2: #3c3c43; --lt-3: #8e8e93; --lt-4: #c7c7cc;
  --sep: rgba(60,60,67,0.13);
  --green: #34c759; --green-dark: #28a745; --green-bg: rgba(52,199,89,0.12); --green-border: rgba(52,199,89,0.3);
  --blue: #007aff; --blue-bg: rgba(0,122,255,0.10);
  --red: #ff3b30; --red-bg: rgba(255,59,48,0.10);
  --amber: #ff9f0a; --amber-bg: rgba(255,159,10,0.12);
  --r-sm: 8px; --r-md: 13px; --r-lg: 18px; --r-xl: 22px;
  --shadow-xs: 0 1px 2px rgba(0,0,0,0.06);
  --shadow-sm: 0 2px 8px rgba(0,0,0,0.08), 0 0 0 0.5px rgba(0,0,0,0.04);
  --shadow-md: 0 4px 16px rgba(0,0,0,0.10), 0 0 0 0.5px rgba(0,0,0,0.06);
  width: 100%; font-family: -apple-system, BlinkMacSystemFont, 'SF Pro Text', 'Google Sans', Roboto, sans-serif;
  font-size: 13px; display: block; position: static; z-index: auto;
  -webkit-font-smoothing: antialiased;
}
.am-loading { display:flex; align-items:center; gap:8px; padding:12px 20px; background:var(--dark-1); }
.am-loading-dot { width:6px; height:6px; border-radius:50%; background:var(--green); animation:am-blink 1.2s infinite; }
.am-loading-txt { font-size:12px; color:var(--dt-3); font-weight:500; }
.am-wrap { background:var(--dark-1); width:100%; max-height:460px; overflow-y:auto; overflow-x:hidden; border-bottom:1px solid rgba(0,0,0,0.2); position:relative; z-index:1; }
.am-wrap::-webkit-scrollbar { width:4px; }
.am-wrap::-webkit-scrollbar-track { background:transparent; }
.am-wrap::-webkit-scrollbar-thumb { background:var(--dark-3); border-radius:4px; }
.am-brand-row { display:flex; align-items:center; gap:10px; padding:10px 20px; background:rgba(28,28,30,0.95); backdrop-filter:blur(20px) saturate(180%); -webkit-backdrop-filter:blur(20px) saturate(180%); border-bottom:1px solid var(--dark-sep); position:sticky; top:0; z-index:10; }
.am-brand-icon { font-size:14px; }
.am-brand-name { font-size:12px; font-weight:700; letter-spacing:0.4px; color:var(--dt-1); flex:1; display:flex; align-items:center; gap:10px; }
.am-brand-name::before { content:''; display:inline-block; width:3px; height:14px; background:var(--green); border-radius:2px; flex-shrink:0; }
.am-brand-sync { font-size:11px; color:var(--green); font-weight:600; background:var(--green-bg); padding:3px 10px; border-radius:20px; border:1px solid var(--green-border); display:flex; align-items:center; gap:5px; }
.am-si-list { padding:10px 14px; display:flex; flex-direction:column; gap:8px; }
.am-si { background:#3a3a3c; border-radius:var(--r-lg); overflow:hidden; box-shadow:0 2px 8px rgba(0,0,0,0.18), inset 0 0.5px 0 rgba(255,255,255,0.10); border:1px solid rgba(0,0,0,0.07); transition:box-shadow 0.2s; }
.am-si:hover { box-shadow:0 4px 16px rgba(0,0,0,0.22), inset 0 0.5px 0 rgba(255,255,255,0.13); }
.am-si-head { display:flex; align-items:center; gap:10px; padding:12px 16px; cursor:pointer; background:transparent; user-select:none; transition:background 0.15s; border-radius:var(--r-lg); }
.am-si-head:hover { background:rgba(255,255,255,0.05); }
.am-si--open .am-si-head { border-radius:var(--r-lg) var(--r-lg) 0 0; background:#3a3a3c; border-bottom:1px solid var(--dark-sep); }
.am-si-dot { width:9px; height:9px; border-radius:50%; flex-shrink:0; box-shadow:0 0 6px currentColor; }
.am-si-label { font-size:13px; font-weight:600; color:var(--dt-1); letter-spacing:0.1px; flex:1; }
.am-si-meta { display:flex; align-items:center; gap:6px; font-size:11px; color:var(--dt-3); }
.am-meta-sep { color:var(--dt-4); }
.am-badge { background:var(--amber); color:#fff; font-size:10px; font-weight:800; padding:2px 8px; border-radius:20px; margin-left:4px; box-shadow:0 2px 6px rgba(255,159,10,0.5); animation:am-badge-pulse 2s infinite; }
@keyframes am-badge-pulse { 0%,100% { box-shadow:0 2px 6px rgba(255,159,10,0.5); } 50% { box-shadow:0 2px 12px rgba(255,159,10,0.8); } }
.am-si-chevron { font-size:9px; color:var(--dt-3); margin-left:2px; transition:transform 0.2s; }
.am-si--open .am-si-chevron { transform:rotate(180deg); }
.am-si-body { background:#f0f0f2; }
.am-tabs { display:flex; align-items:center; background:#ffffff; border-bottom:1px solid #e0e0e3; padding:0 14px; gap:0; }
.am-tab { background:none; border:none; border-bottom:2px solid transparent; color:#9e9ea3; font-size:11px; font-weight:600; letter-spacing:0.3px; padding:9px 12px; cursor:pointer; text-transform:uppercase; transition:color 0.12s, border-color 0.12s; white-space:nowrap; font-family:inherit; }
.am-tab:hover { color:#3c3c43; }
.am-tab--on { color:var(--green-dark); border-bottom-color:var(--green-dark); }
.am-tabs-info { margin-left:auto; font-size:10px; color:#b0b0b5; }
.am-tab-badge { display:inline-flex; align-items:center; justify-content:center; background:var(--red); color:#fff; font-size:9px; font-weight:800; min-width:16px; height:16px; padding:0 4px; border-radius:20px; margin-left:5px; vertical-align:middle; box-shadow:0 1px 4px rgba(255,59,48,0.4); }
.am-tab-content { max-height:300px; overflow-y:auto; padding:8px 10px; display:flex; flex-direction:column; gap:6px; background:#f0f0f2; }
.am-tab-content::-webkit-scrollbar { width:3px; }
.am-tab-content::-webkit-scrollbar-track { background:transparent; }
.am-tab-content::-webkit-scrollbar-thumb { background:#c0c0c5; border-radius:4px; }
.am-row { display:flex; align-items:flex-start; gap:11px; padding:11px 14px; border-radius:var(--r-md); background:linear-gradient(180deg,#ffffff 0%,#f5f5f7 100%); border:1px solid rgba(0,0,0,0.09); border-bottom:2px solid rgba(0,0,0,0.13); box-shadow:0 1px 4px rgba(0,0,0,0.06), inset 0 1px 0 rgba(255,255,255,0.95); transition:box-shadow 0.15s, transform 0.1s; text-decoration:none; color:inherit; }
.am-row--link { cursor:pointer; }
.am-row--link:hover { background:linear-gradient(180deg,#f8f8fa 0%,#efeff1 100%); border-bottom-color:rgba(0,0,0,0.18); box-shadow:0 1px 3px rgba(0,0,0,0.07), inset 0 1px 0 rgba(255,255,255,0.80) !important; transform:translateY(1px); text-decoration:none; }
.am-row--link:hover .am-open-icon { opacity:0.8; transform:translateX(2px); }
.am-row:hover { box-shadow:0 3px 10px rgba(0,0,0,0.08); }
.am-row--unread { background:#ffffff; border-left:3px solid var(--green); padding-left:11px; box-shadow:0 1px 6px rgba(52,199,89,0.12); }
.am-row--unread:hover { box-shadow:0 4px 12px rgba(52,199,89,0.18) !important; }
.am-avatar { width:32px; height:32px; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:12px; font-weight:700; color:#fff; flex-shrink:0; margin-top:1px; box-shadow:0 2px 6px rgba(0,0,0,0.3); }
.am-avatar-disp { background:var(--blue-bg); font-size:16px; color:var(--blue); border:1px solid rgba(0,122,255,0.2); }
.am-row-info { flex:1; min-width:0; }
.am-row-top { display:flex; align-items:center; gap:6px; flex-wrap:wrap; }
.am-row-nome { font-size:13px; font-weight:600; color:#1c1c1e; }
.am-role-tag { font-size:10px; font-weight:700; padding:2px 8px; border-radius:20px; text-transform:uppercase; letter-spacing:0.3px; flex-shrink:0; }
.am-role-g { background:var(--green-bg); color:var(--green); }
.am-role-ex { background:var(--red-bg); color:var(--red); }
.am-role-disp { background:var(--blue-bg); color:var(--blue); }
.am-unread-dot { width:7px; height:7px; background:var(--green); border-radius:50%; flex-shrink:0; box-shadow:0 0 6px var(--green); animation:am-blink 2s infinite; }
.am-row-data { font-size:11px; color:#6e6e78; margin-left:auto; white-space:nowrap; flex-shrink:0; }
.am-row-sub { font-size:12px; color:#4a4a52; margin-top:3px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width:480px; }
.am-row-status { font-size:11px; color:var(--green); margin-top:3px; font-weight:600; }
.am-row-hint { color:var(--green); }
.am-open-icon { font-size:13px; color:#aeaeb2; flex-shrink:0; align-self:center; opacity:0.4; transition:opacity 0.15s, transform 0.15s; }
.am-row--unread .am-open-icon { opacity:0.7; color:var(--green); }
.am-empty { padding:24px 16px; font-size:13px; color:#8e8e93; text-align:center; font-weight:500; }
.am-si-foot { display:flex; align-items:center; justify-content:space-between; padding:8px 16px 12px; background:#2c2c2e; }
.am-foot-gestores { display:flex; flex-wrap:wrap; gap:5px; flex:1; }
.am-foot-link { font-size:11px; color:var(--green); text-decoration:none; white-space:nowrap; margin-left:12px; flex-shrink:0; font-weight:600; }
.am-foot-link:hover { text-decoration:underline; }
.am-chip { display:inline-flex; align-items:center; gap:4px; border:1px solid rgba(255,255,255,0.18); border-radius:20px; padding:3px 9px 3px 6px; font-size:11px; color:var(--dt-2); background:rgba(255,255,255,0.08); white-space:nowrap; transition:background 0.12s; }
.am-chip:hover { background:rgba(255,255,255,0.14); }
.am-chip-dot { width:6px; height:6px; border-radius:50%; flex-shrink:0; }
.am-chip-role { font-size:9px; font-weight:700; padding:1px 5px; border-radius:20px; }
.am-chip-g { background:var(--green-bg); color:var(--green); }
.am-chip-ex { background:var(--red-bg); color:var(--red); }
.am-mov { padding:10px 10px 12px; }
.am-mov-card { display:flex; gap:13px; align-items:flex-start; background:linear-gradient(180deg,#ffffff 0%,#f4f4f6 100%); border:1px solid rgba(0,0,0,0.10); border-bottom:2px solid rgba(0,0,0,0.14); border-radius:var(--r-lg); padding:14px 16px; margin-bottom:8px; text-decoration:none; color:inherit; box-shadow:0 2px 6px rgba(0,0,0,0.08), 0 0 0 0.5px rgba(0,0,0,0.04), inset 0 1px 0 rgba(255,255,255,0.95); transition:background 0.15s, transform 0.1s, box-shadow 0.15s; }
.am-mov-card--link { cursor:pointer; }
.am-mov-card--link:hover { background:linear-gradient(180deg,#f8f8fa 0%,#efeff1 100%); border-color:rgba(0,0,0,0.13); border-bottom-color:rgba(0,0,0,0.18); transform:translateY(1px); box-shadow:0 1px 3px rgba(0,0,0,0.07), inset 0 1px 0 rgba(255,255,255,0.80); text-decoration:none; }
.am-mov-card--link:hover .am-open-icon--mov { opacity:0.9; }
.am-mov-card--unread { background:rgba(52,199,89,0.09); border-color:rgba(52,199,89,0.3); border-left:4px solid var(--green); padding-left:12px; }
.am-mov-card--unread:hover { background:rgba(52,199,89,0.13); border-color:rgba(52,199,89,0.45); }
.am-mov-icon { font-size:22px; flex-shrink:0; margin-top:2px; }
.am-mov-card-info { flex:1; min-width:0; }
.am-mov-tipo { font-size:10px; text-transform:uppercase; letter-spacing:0.8px; font-weight:700; color:#5a5a60; margin-bottom:6px; display:flex; align-items:center; gap:7px; }
.am-pulse { font-size:10px; font-weight:800; color:var(--green); background:var(--green-bg); padding:2px 8px; border-radius:20px; animation:am-blink 1.6s infinite; }
.am-mov-nome { font-size:15px; font-weight:700; margin-bottom:3px; color:#1c1c1e; }
.am-mov-assunto { font-size:12px; color:#4a4a52; margin-bottom:4px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.am-mov-data { font-size:11px; color:#6e6e78; margin-bottom:4px; }
.am-mov-hint { font-size:11px; color:var(--green); font-weight:600; }
.am-open-icon--mov { font-size:14px; color:#c7c7cc; align-self:center; flex-shrink:0; opacity:0.3; transition:opacity 0.15s; }
.am-mov-card--unread .am-open-icon--mov { opacity:0.6; color:var(--green); }
.am-mov-hist-title { font-size:10px; text-transform:uppercase; letter-spacing:0.9px; color:#aeaeb2; margin:4px 0 6px; font-weight:700; }
.am-mov-hist-row { display:flex; align-items:center; gap:8px; padding:7px 10px; margin:3px 0; border-radius:var(--r-sm); font-size:12px; color:#6e6e73; background:#e8e8ed; text-decoration:none; transition:background 0.1s; }
.am-mov-hist--link:hover { background:#dcdce0; color:#3c3c43; text-decoration:none; }
.am-mov-hist--unread { color:#1c1c1e; font-weight:600; }
.am-mov-hist-icon { font-size:10px; color:#aeaeb2; flex-shrink:0; }
.am-mov-hist-nome { flex:1; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.am-mov-hist-data { font-size:11px; color:#aeaeb2; flex-shrink:0; }
.am-hist-dot { color:var(--green); font-size:9px; }
@keyframes am-blink { 0%,100% { opacity:1; } 50% { opacity:0.25; } }
.am-contatos-grupo { padding:0; }
.am-contatos-grupo-label { font-size:10px; font-weight:700; text-transform:uppercase; letter-spacing:1px; color:#aeaeb2; padding:10px 14px 5px; }
#am-back-bar { display:flex; align-items:center; gap:8px; padding:7px 16px; background:rgba(28,28,30,0.92); backdrop-filter:blur(20px) saturate(180%); -webkit-backdrop-filter:blur(20px) saturate(180%); border-bottom:1px solid var(--dark-sep); position:sticky; top:0; z-index:999; }
#am-back-bar button { background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.12); border-radius:var(--r-sm); color:var(--dt-2); cursor:pointer; font-size:12px; font-weight:600; padding:5px 13px; transition:background 0.12s, color 0.12s; font-family:inherit; letter-spacing:0.2px; backdrop-filter:blur(8px); }
#am-back-bar button:hover { background:rgba(255,255,255,0.14); color:var(--dt-1); }
#am-back-primary { color:var(--green) !important; border-color:var(--green-border) !important; background:var(--green-bg) !important; }
#am-back-primary:hover { background:rgba(52,199,89,0.18) !important; }
.am-foot-del { background:none; border:1px solid transparent; cursor:pointer; font-size:13px; padding:3px 7px; border-radius:var(--r-sm); color:var(--red); opacity:0.3; transition:opacity 0.15s, background 0.15s; margin-left:6px; flex-shrink:0; }
.am-foot-del:hover { opacity:1; background:var(--red-bg); border-color:rgba(255,59,48,0.3); }
`;

  // ════════════════════════════════════════════════════════════════
  // CSS — PAINEL COMPOSE
  // ════════════════════════════════════════════════════════════════
  const CSS_COMPOSE = `
#am-compose-panel {
  position: fixed;
  z-index: 2147483647;
  width: 300px;
  max-height: 580px;
  background: #1c1c1e;
  border-radius: 14px;
  box-shadow: 0 8px 40px rgba(0,0,0,0.55), 0 0 0 1px rgba(255,255,255,0.07);
  font-family: -apple-system, BlinkMacSystemFont, 'Google Sans', Roboto, sans-serif;
  font-size: 13px;
  color: #fff;
  display: flex;
  flex-direction: column;
  overflow: hidden;
  transition: opacity 0.18s, transform 0.18s;
  -webkit-font-smoothing: antialiased;
}
#am-compose-panel.am-cp-hidden {
  opacity: 0;
  pointer-events: none;
  transform: scale(0.96) translateY(6px);
}
#am-cp-header {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 11px 14px 10px;
  background: rgba(255,255,255,0.04);
  border-bottom: 1px solid rgba(255,255,255,0.07);
  flex-shrink: 0;
}
#am-cp-header-title {
  font-size: 12px;
  font-weight: 700;
  color: rgba(255,255,255,0.9);
  flex: 1;
  letter-spacing: 0.3px;
  display: flex;
  align-items: center;
  gap: 7px;
}
#am-cp-header-title::before {
  content: '';
  display: inline-block;
  width: 3px;
  height: 13px;
  background: #34c759;
  border-radius: 2px;
  flex-shrink: 0;
}
#am-cp-close {
  background: none;
  border: none;
  color: rgba(255,255,255,0.35);
  font-size: 16px;
  cursor: pointer;
  padding: 2px 5px;
  border-radius: 6px;
  line-height: 1;
  transition: color 0.12s, background 0.12s;
}
#am-cp-close:hover { color: rgba(255,255,255,0.8); background: rgba(255,255,255,0.08); }
#am-cp-body {
  padding: 10px 12px 4px;
  overflow-y: auto;
  flex: 1;
}
#am-cp-body::-webkit-scrollbar { width: 3px; }
#am-cp-body::-webkit-scrollbar-track { background: transparent; }
#am-cp-body::-webkit-scrollbar-thumb { background: #3a3a3c; border-radius: 4px; }
.am-cp-section-label {
  font-size: 9.5px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 1.2px;
  color: rgba(255,255,255,0.3);
  padding: 8px 2px 4px;
  display: block;
}
.am-cp-select, .am-cp-input {
  width: 100%;
  background: #2c2c2e;
  border: 1px solid rgba(255,255,255,0.10);
  border-radius: 8px;
  color: rgba(255,255,255,0.88);
  font-size: 13px;
  font-family: inherit;
  padding: 8px 10px;
  margin-bottom: 2px;
  outline: none;
  transition: border-color 0.15s, background 0.15s;
  -webkit-appearance: none;
  appearance: none;
  box-sizing: border-box;
}
.am-cp-select:focus, .am-cp-input:focus {
  border-color: rgba(52,199,89,0.5);
  background: #333335;
}
.am-cp-select:disabled { opacity: 0.4; cursor: not-allowed; }
.am-cp-select option { background: #2c2c2e; }
.am-cp-email-hint {
  font-size: 11px;
  color: rgba(52,199,89,0.8);
  padding: 2px 4px 6px;
  font-weight: 500;
}
#am-cp-footer {
  padding: 10px 12px 8px;
  border-top: 1px solid rgba(255,255,255,0.06);
  flex-shrink: 0;
}
#am-cp-injetar {
  width: 100%;
  background: #34c759;
  color: #fff;
  border: none;
  border-radius: 8px;
  font-size: 13px;
  font-weight: 700;
  font-family: inherit;
  padding: 10px 14px;
  cursor: pointer;
  transition: background 0.15s, opacity 0.15s;
  letter-spacing: 0.2px;
}
#am-cp-injetar:hover:not(:disabled) { background: #28a745; }
#am-cp-injetar:disabled {
  opacity: 0.4;
  cursor: not-allowed;
  background: #3a3a3c;
  color: rgba(255,255,255,0.4);
}
#am-cp-status {
  padding: 6px 12px 10px;
  font-size: 11px;
  color: rgba(255,255,255,0.4);
  text-align: center;
  min-height: 28px;
  display: flex;
  align-items: center;
  justify-content: center;
}
#am-cp-status.am-cp-status-ok  { color: #34c759; }
#am-cp-status.am-cp-status-err { color: #ff3b30; }
`;

  // ════════════════════════════════════════════════════════════════
  // UTILITÁRIOS
  // ════════════════════════════════════════════════════════════════
  function injetarCSS() {
    if (!document.getElementById('am-styles')) {
      const s = document.createElement('style');
      s.id = 'am-styles';
      s.textContent = CSS;
      document.head.appendChild(s);
    }
    if (!document.getElementById('am-compose-styles')) {
      const s2 = document.createElement('style');
      s2.id = 'am-compose-styles';
      s2.textContent = CSS_COMPOSE;
      document.head.appendChild(s2);
    }
  }

  function storageGet(keys) {
    return new Promise(resolve => {
      const result = {};
      for (const k of keys) {
        const raw = GM_getValue(k, null);
        if (raw !== null) {
          try { result[k] = JSON.parse(raw); } catch(e) { result[k] = raw; }
        }
      }
      resolve(result);
    });
  }

  function storageSet(obj) {
    return new Promise(resolve => {
      for (const [k, v] of Object.entries(obj)) GM_setValue(k, JSON.stringify(v));
      resolve();
    });
  }

  function fetchJson(url) {
    return new Promise((resolve, reject) => {
      GM_xmlhttpRequest({
        method: 'GET',
        url: url + (url.includes('?') ? '&_t=' : '?_t=') + Date.now(),
        onload: res => {
          try { resolve(JSON.parse(res.responseText)); }
          catch(e) { reject(new Error('JSON parse error')); }
        },
        onerror: () => reject(new Error('Network error')),
        ontimeout: () => reject(new Error('Timeout'))
      });
    });
  }

  // ── NOVO: POST para ações que gravam dados ──
 function postJson(url, body) {
  return new Promise((resolve, reject) => {
    const params = Object.entries(body)
      .map(([k, v]) => encodeURIComponent(k) + '=' + encodeURIComponent(
        typeof v === 'object' ? JSON.stringify(v) : String(v)
      ))
      .join('&');
    GM_xmlhttpRequest({
      method: 'GET',
      url: url + (url.includes('?') ? '&' : '?') + params + '&_t=' + Date.now(),
      onload: res => {
        try { resolve(JSON.parse(res.responseText)); }
        catch(e) { reject(new Error('JSON parse error: ' + res.responseText.slice(0,80))); }
      },
      onerror: () => reject(new Error('Network error')),
      ontimeout: () => reject(new Error('Timeout'))
    });
  });
}

  let _ttPolicy = null;
  function getTTPolicy() {
    if (_ttPolicy) return _ttPolicy;
    try {
      const win = unsafeWindow || window;
      if (win.trustedTypes && win.trustedTypes.createPolicy) {
        _ttPolicy = win.trustedTypes.createPolicy('am-policy', {
          createHTML: s => s, createScript: s => s, createScriptURL: s => s
        });
      }
    } catch(e) {}
    return _ttPolicy;
  }

  function setHTML(el, html) {
    const p = getTTPolicy();
    if (p) el.innerHTML = p.createHTML(html);
    else el.innerHTML = html;
  }

  function esc(s) {
    return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }
  const esc2 = esc;

  function iniciais(nome) {
    const p = String(nome||'?').trim().split(' ');
    return (p[0][0] + (p[1] ? p[1][0] : '')).toUpperCase();
  }

  function contarGestoresAtivos(gs) { return gs.filter(g => g.tipo !== 'ex-gestor').length; }

  function toComparable(s) {
    if (!s) return '';
    const p = s.trim().split(' '), d = (p[0]||'').split('/');
    if (d.length !== 3) return s;
    return d[2]+'/'+d[1]+'/'+d[0]+' '+(p[1]||'00:00');
  }

  function formatHora(s) { const m = String(s).match(/(\d{2}:\d{2})/); return m ? m[1] : s; }

  function CSS_escape(s) { return String(s).replace(/([^\w-])/g, '\\$1'); }

  // ════════════════════════════════════════════════════════════════
  // SYNC + DADOS
  // ════════════════════════════════════════════════════════════════
  async function sync() {
    try {
      const [movs, gestoresRaw, subinboxesRaw] = await Promise.all([
        fetchJson(WEB_APP_URL + '?acao=listarMovimentacoes'),
        fetchJson(WEB_APP_URL + '?acao=listarGestores').catch(() => []),
        fetchJson(WEB_APP_URL + '?acao=listarSubInboxes').catch(() => [])
      ]);

      const res = await storageGet(['lidosConfirmados']);
      const lidosConfirmados = new Set((res.lidosConfirmados || []).map(String));
      const munMap = {};

      subinboxesRaw.forEach(si => {
        const key = String(si.ID || '').trim();
        if (!key) return;
        munMap[key] = {
          siId: key, municipio: String(si.Nome || ''), cor: String(si.Cor || 'azul'),
          municipioLabel: String(si.Municipio || si.Nome || ''), segmento: String(si.Segmento || 'PM'),
          gestores: [], respostas: [], disparos: [], novos: 0
        };
      });

      if (!subinboxesRaw.length) {
        gestoresRaw.forEach(g => {
          const mun = (g.Municipio || g.municipio || 'Sem Município').trim();
          if (!munMap[mun]) munMap[mun] = { siId: mun, municipio: mun, cor: 'azul', municipioLabel: mun, gestores: [], respostas: [], disparos: [], novos: 0 };
          const emailKey = (g.Email || g.email || '').toLowerCase().trim();
          if (emailKey && !munMap[mun].gestores.find(x => x.email === emailKey)) {
            munMap[mun].gestores.push({ nome: g.Nome || emailKey, email: emailKey, cargo: g.Status || '', tipo: ((g.Status||'')==='Ex-Gestor'||(g.Status||'')==='Ex-Prefeito') ? 'ex-gestor' : 'gestor', corLabel: g.CorLabel||'azul', subInbox: '' });
          }
        });
      }

      const semSubInbox = { siId: '__sem__', municipio: 'Sem sub-inbox', cor: 'graphite', municipioLabel: '', gestores: [], respostas: [], disparos: [], novos: 0 };
      const emailParaSubInbox = {};

      gestoresRaw.forEach(g => {
        const emailKey = (g.Email || '').toLowerCase().trim();
        const subInboxId = String(g.SubInbox || '').trim();
        if (!emailKey) return;
        if (subInboxId) emailParaSubInbox[emailKey] = subInboxId;
        if (subinboxesRaw.length) {
          const bucket = (subInboxId && munMap[subInboxId]) ? munMap[subInboxId] : semSubInbox;
          if (!bucket.gestores.find(x => x.email === emailKey)) {
            bucket.gestores.push({ nome: g.Nome||emailKey, email: emailKey, cargo: g.Status||'', corLabel: g.CorLabel||'azul', subInbox: subInboxId, tipo: (g.Status==='Ex-Gestor'||g.Status==='Ex-Prefeito'||g.Status==='Ex-Presidente CM') ? 'ex-gestor' : 'gestor' });
          }
        }
      });

      movs.forEach(m => {
        const emailKey = (m.de || '').toLowerCase().trim();
        let bucket;
        if (subinboxesRaw.length) {
          const subInboxId = emailParaSubInbox[emailKey] || '';
          if (!subInboxId || !munMap[subInboxId]) return;
          bucket = munMap[subInboxId];
        } else {
          const mun = (m.municipio || '').trim();
          if (!mun || !munMap[mun]) return;
          bucket = munMap[mun];
        }
        if (m.tipo === 'resposta') {
          const itemId = String(m.id || m.threadId || '');
          const lidoFinal = m.lido || lidosConfirmados.has(itemId);
          bucket.respostas.push({ ...m, lido: lidoFinal, municipio: bucket.siId });
          if (!lidoFinal) bucket.novos++;
        } else {
          bucket.disparos.push(m);
        }
      });

      Object.values(munMap).forEach(bucket => {
        bucket.gestores.sort((a, b) => {
          if (a.tipo !== b.tipo) return a.tipo === 'gestor' ? -1 : 1;
          return a.nome.localeCompare(b.nome, 'pt-BR');
        });
      });

      const lastSync = new Date().toLocaleString('pt-BR');
      await storageSet({ subinboxData: munMap, movimentacoes: movs, subinboxes: subinboxesRaw, lastSync });
      state.subinboxData  = munMap;
      state.movimentacoes = movs;
      state.lastSync      = lastSync;

      // Atualiza cache do compose se já estava carregado
      _gestoresCache   = gestoresRaw;
      _subinboxesCache = subinboxesRaw;

      const root = document.getElementById('am-root');
      if (root) renderRoot(root);
    } catch(e) {
      console.warn('[Adriana] Sync error:', e.message);
    }
  }

  async function carregarDados() {
    const res = await storageGet(['subinboxData', 'movimentacoes', 'lastSync']);
    state.subinboxData  = res.subinboxData  || {};
    state.movimentacoes = res.movimentacoes || [];
    state.lastSync      = res.lastSync      || null;
  }

  function iniciarSyncPeriodico() {
    if (syncTimer) clearInterval(syncTimer);
    syncTimer = setInterval(() => sync(), SYNC_INTERVAL_MS);
  }

  // ════════════════════════════════════════════════════════════════
  // NAVEGAÇÃO
  // ════════════════════════════════════════════════════════════════
  function iniciarNavWatch() {
    if (navTimer) clearInterval(navTimer);
    navTimer = setInterval(() => {
      const cur = location.href;
      if (cur !== lastURL) {
        const anteriorEhPrimary = /\/#inbox\/?$/.test(lastURL) || lastURL === 'https://mail.google.com/mail/u/0/';
        if (lastURL && !anteriorEhPrimary) lastNonThreadURL = lastURL;
        lastURL = cur;
        injecting = false;
        removerPainel();
        setTimeout(tentarInjetar, 400);
      } else if (!document.getElementById('am-root')) {
        tentarInjetar();
      }
      gerenciarBotaoVoltar();
    }, 600);
  }

  function ehThreadAberta(url) {
    return /\/#(?:inbox|sent|label|all|search)\/[A-Za-z0-9_\-:%.+]+/.test(url) &&
           !/\/#(?:inbox|sent|label|search)\/?$/.test(url);
  }

  function gerenciarBotaoVoltar() {
    const url = location.href;
    const estaEmPrimary = (/\/#inbox\/?$/.test(url) || url === 'https://mail.google.com/mail/u/0/') && !ehThreadAberta(url);
    const estaEmThread  = ehThreadAberta(url);
    const estaEmStarred = /\/#starred/.test(url)  && !estaEmThread;
    const estaEmSent    = /\/#sent\/?$/.test(url) && !estaEmThread;
    const estaEmTrash   = /\/#trash/.test(url)    && !estaEmThread;
    const estaEmSpam    = /\/#spam/.test(url)     && !estaEmThread;
    const estaEmAll     = /\/#all\/?$/.test(url)  && !estaEmThread;
    const estaEmSearch  = /\/#search\//.test(url) && !estaEmThread;
    const estaEmLabel   = /\/#label\//.test(url)  && !estaEmThread;
    const precisaBarra  = estaEmThread || estaEmStarred || estaEmSent || estaEmTrash || estaEmSpam || estaEmAll || estaEmSearch || estaEmLabel;
    const existente = document.getElementById('am-back-bar');

    if (!precisaBarra || estaEmPrimary) { if (existente) existente.remove(); return; }
    if (existente) return;

    const threadContainer = document.querySelector('[role="main"]');
    if (!threadContainer) return;

    let contexto = '';
    if (estaEmThread)  contexto = 'Email aberto';
    if (estaEmStarred) contexto = 'Com Estrela';
    if (estaEmSent)    contexto = 'Enviados';
    if (estaEmTrash)   contexto = 'Lixeira';
    if (estaEmSpam)    contexto = 'Spam';
    if (estaEmAll)     contexto = 'Todos os emails';
    if (estaEmSearch)  contexto = 'Resultado de busca';
    if (estaEmLabel)   contexto = 'Label';

    const bar = document.createElement('div');
    bar.id = 'am-back-bar';
    const btnPrev = document.createElement('button');
    btnPrev.id = 'am-back-prev';
    btnPrev.textContent = '<- Voltar';
    const ctx = document.createElement('span');
    ctx.style.cssText = 'font-size:11px;color:#555;flex:1;padding:0 10px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;';
    ctx.textContent = contexto;
    const btnPrimary = document.createElement('button');
    btnPrimary.id = 'am-back-primary';
    btnPrimary.textContent = 'Primary';
    bar.appendChild(btnPrev); bar.appendChild(ctx); bar.appendChild(btnPrimary);
    threadContainer.insertBefore(bar, threadContainer.firstChild);
    btnPrev.addEventListener('click', () => { if (lastNonThreadURL && lastNonThreadURL !== location.href) location.href = lastNonThreadURL; else history.back(); });
    btnPrimary.addEventListener('click', () => { location.href = 'https://mail.google.com/mail/u/0/#inbox'; });
  }

  function removerPainel() {
    const el  = document.getElementById('am-root');     if (el)  el.remove();
    const bar = document.getElementById('am-back-bar'); if (bar) bar.remove();
  }

  function tentarInjetar() {
    if (injecting) return;
    injecting = true;
    let tries = 0;
    function loop() {
      tries++;
      if (tries > 50) { injecting = false; return; }
      const existing = document.getElementById('am-root');
      if (existing) { renderRoot(existing); injecting = false; return; }
      const main = document.querySelector('[role="main"]');
      if (!main) { setTimeout(loop, 200); return; }
      const barraAbas = encontrarBarraDeAbas(main);
      if (!barraAbas) { setTimeout(loop, 200); return; }
      injetarDepoisDasAbas(barraAbas);
      injecting = false;
    }
    loop();
  }

  function encontrarBarraDeAbas(main) {
    const aKh = main.querySelector('.aKh');
    if (aKh) return aKh;
    const aKk = main.querySelector('table.aKk, TABLE.aKk, .aKk');
    if (aKk) {
      let el = aKk;
      while (el && el.parentElement !== main) el = el.parentElement;
      if (el) return el;
      return aKk.parentElement || aKk;
    }
    const tablist = main.querySelector('[role="tablist"]');
    if (tablist) {
      let el = tablist;
      while (el && el.parentElement !== main) el = el.parentElement;
      if (el) return el;
    }
    return null;
  }

  function injetarDepoisDasAbas(barraAbas) {
    const pai = barraAbas.parentElement;
    if (!pai) return;
    const root = document.createElement('div');
    root.id = 'am-root';
    root.style.cssText = 'width:100%;display:block;';
    const proximo = barraAbas.nextSibling;
    if (proximo) pai.insertBefore(root, proximo);
    else pai.appendChild(root);
    renderRoot(root);
  }

  // ════════════════════════════════════════════════════════════════
  // RENDER PAINEL PRINCIPAL
  // ════════════════════════════════════════════════════════════════
  function renderRoot(root) {
    const html = buildPainelHTML();
    if (root._lastHtml === html) return;
    root._lastHtml = html;
    setHTML(root, html);
    bindEventos(root);
  }

  function buildPainelHTML() {
    const chaves = Object.keys(state.subinboxData).sort((a, b) => {
  const da = state.subinboxData[a], db = state.subinboxData[b];
  if (a === '__sem__') return 1; if (b === '__sem__') return -1;
  const na = da.novos || 0, nb = db.novos || 0;
  if (nb !== na) return nb - na;
  // Agrupa por município, depois PM antes de CM
  const munA = (da.municipio || '').toLowerCase();
  const munB = (db.municipio || '').toLowerCase();
  if (munA !== munB) return munA.localeCompare(munB, 'pt-BR');
  const segA = (da.segmento || 'PM').toUpperCase();
  const segB = (db.segmento || 'PM').toUpperCase();
  if (segA !== segB) return segA === 'PM' ? -1 : 1;
  return (da.municipio || '').localeCompare(db.municipio || '', 'pt-BR');
});

    if (!chaves.length && !state.lastSync) {
      return `<div class="am-loading"><span class="am-loading-dot"></span><span class="am-loading-txt">Adriana Matos Advocacia — sincronizando dados...</span></div>`;
    }
    if (!chaves.length) {
      return `<div class="am-wrap"><div class="am-brand-row"><span class="am-brand-icon">Lei</span><span class="am-brand-name">Adriana Matos Advocacia — Gerenciador de Emails</span><span class="am-brand-sync">${state.lastSync ? 'OK ' + formatHora(state.lastSync) : ''}</span></div><div class="am-empty" style="padding:24px 20px;">Nenhuma sub-inbox cadastrada. <a href="${WEB_APP_URL}" target="_blank" style="color:#4986e7;">Abrir painel</a></div></div>`;
    }
    return `<div class="am-wrap"><div class="am-brand-row"><span class="am-brand-icon">Lei</span><span class="am-brand-name">Adriana Matos Advocacia — Gerenciador de Emails</span><span class="am-brand-sync">${state.lastSync ? 'OK ' + formatHora(state.lastSync) : ''}</span></div><div class="am-si-list">${chaves.map(buildSubInbox).join('')}</div></div>`;
  }

  function buildSubInbox(chave) {
    const d = state.subinboxData[chave];
    if (!d) return '';
    const expanded = !!state.expandidos[chave];
    const aba = state.abaAtiva[chave] || 'movimentacao';
    const novos = d.novos || 0;
    const cor = COR_MAP[d.cor] || (d.gestores.length ? (COR_MAP[d.gestores[0].corLabel] || COR_MAP.azul) : COR_MAP.azul);
    const segBadge = (d.segmento === 'CM') ? 'CM' : 'PM';
    const labelExibida = d.municipio
      ? (d.municipioLabel && d.municipioLabel !== d.municipio ? `${d.municipio.toUpperCase()} · ${segBadge} · ${d.municipioLabel}` : `${d.municipio.toUpperCase()} · ${segBadge}`)
      : `${chave.toUpperCase()} · ${segBadge}`;

    return `<div class="am-si ${expanded ? 'am-si--open' : ''}" data-mun="${esc(chave)}">
    <div class="am-si-head" data-toggle="${esc(chave)}">
      <span class="am-si-dot" style="background:${cor}"></span>
      <span class="am-si-label">SUB-INBOX · ${esc(labelExibida)}</span>
      <span class="am-si-meta">
        <span class="am-meta-item">${d.respostas.length} resp</span>
        <span class="am-meta-sep">·</span>
        <span class="am-meta-item">${d.disparos.length} disp</span>
        <span class="am-meta-sep">·</span>
        <span class="am-meta-item">${d.gestores.length} contato${d.gestores.length !== 1 ? 's' : ''}</span>
        ${novos ? `<span class="am-badge">${novos} novo${novos > 1 ? 's' : ''}</span>` : ''}
      </span>
      <span class="am-si-chevron">${expanded ? '▲' : '▼'}</span>
    </div>
    ${expanded ? `<div class="am-si-body">
      <div class="am-tabs">
        <button class="am-tab ${aba==='movimentacao'?'am-tab--on':''}" data-aba="movimentacao" data-mun="${esc(chave)}">ULTIMA MOV.</button>
        <button class="am-tab ${aba==='respostas'?'am-tab--on':''}" data-aba="respostas" data-mun="${esc(chave)}">RESPOSTAS (${d.respostas.length})${d.respostas.filter(r=>!r.lido).length ? `<span class="am-tab-badge">${d.respostas.filter(r=>!r.lido).length}</span>` : ''}</button>
        <button class="am-tab ${aba==='disparos'?'am-tab--on':''}" data-aba="disparos" data-mun="${esc(chave)}">DISPAROS (${d.disparos.length})</button>
        <button class="am-tab ${aba==='contatos'?'am-tab--on':''}" data-aba="contatos" data-mun="${esc(chave)}">CONTATOS (${d.gestores.length})</button>
        <span class="am-tabs-info">${contarGestoresAtivos(d.gestores)} ativo(s)</span>
      </div>
      <div class="am-tab-content">
        ${aba==='movimentacao' ? buildMovimentacao(d) : ''}
        ${aba==='respostas' ? buildRespostas(d.respostas) : ''}
        ${aba==='disparos' ? buildDisparos(d.disparos) : ''}
        ${aba==='contatos' ? buildContatos(d.gestores, cor) : ''}
      </div>
      <div class="am-si-foot">
        <span class="am-foot-gestores">${buildGestoresChips(d.gestores)}</span>
        <a class="am-foot-link" href="${WEB_APP_URL}" target="_blank">Abrir painel</a>
        ${chave !== '__sem__' ? `<button class="am-foot-del" data-del-si="${esc(chave)}" data-del-nome="${esc(d.municipio||chave)}" title="Excluir esta sub-inbox">Del</button>` : ''}
      </div>
    </div>` : ''}
  </div>`;
  }

  function buildContatos(gestores, corSubInbox) {
    if (!gestores.length) return '<div class="am-empty">Nenhum contato vinculado a esta sub-inbox.</div>';
    const ativos = gestores.filter(g => g.tipo !== 'ex-gestor');
    const exs    = gestores.filter(g => g.tipo === 'ex-gestor');
    function renderGrupo(lista, titulo) {
      if (!lista.length) return '';
      return `<div class="am-contatos-grupo"><div class="am-contatos-grupo-label">${titulo}</div>${lista.map(g => {
        const cor = COR_MAP[g.corLabel] || COR_MAP.azul;
        const isEx = g.tipo === 'ex-gestor';
        return `<div class="am-row"><span class="am-avatar" style="background:${cor}">${iniciais(g.nome)}</span><div class="am-row-info"><div class="am-row-top"><span class="am-row-nome">${esc(g.nome)}</span><span class="am-role-tag ${isEx?'am-role-ex':'am-role-g'}">${esc(g.cargo||g.status||(isEx?'Ex-Gestor':'—'))}</span></div><div class="am-row-sub" style="color:#555;">${esc(g.email)}</div></div><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${cor};flex-shrink:0;align-self:center;margin-left:4px;" title="${esc(g.corLabel)}"></span></div>`;
      }).join('')}</div>`;
    }
    return renderGrupo(ativos, 'Ativos') + renderGrupo(exs, 'Ex-gestores / Inativos');
  }

  function buildGestoresChips(gestores) {
    return gestores.map(g => {
      const cor = COR_MAP[g.corLabel] || COR_MAP.azul;
      const isEx = g.tipo === 'ex-gestor';
      const cargo = g.cargo || g.status || (isEx ? 'Ex' : '—');
      return `<span class="am-chip" style="border-color:${cor}" title="${esc(g.email)}"><span class="am-chip-dot" style="background:${cor}"></span>${esc(g.nome)}<span class="am-chip-role ${isEx?'am-chip-ex':'am-chip-g'}">${esc(cargo)}</span></span>`;
    }).join('');
  }

  function buildRespostas(respostas) {
    if (!respostas.length) return '<div class="am-empty">Nenhuma resposta registrada.</div>';
    return respostas.slice(0, 15).map(r => {
      const cor = COR_MAP[r.corLabel] || COR_MAP.azul;
      const isEx = r.statusGestor === 'Ex-Gestor';
      const naoLido = !r.lido;
      const href = r.threadUrl || '#';
      return `<a class="am-row ${naoLido?'am-row--unread':''} am-row--link" href="${esc(href)}" target="_blank" data-marcar-lido="${esc(r.id||r.threadId||'')}" data-mun-lido="${esc(r.municipio||'')}" title="Abrir no Gmail">
        <span class="am-avatar" style="background:${cor}">${iniciais(r.nome)}</span>
        <div class="am-row-info"><div class="am-row-top"><span class="am-row-nome">${esc(r.nome)}</span><span class="am-role-tag ${isEx?'am-role-ex':'am-role-g'}">${isEx?'EX-GESTOR':'GESTOR'}</span>${naoLido?'<span class="am-unread-dot"></span>':''}<span class="am-row-data">${esc(r.data)}</span></div><div class="am-row-sub">${esc(r.assunto)}</div>${naoLido?'<div class="am-row-status">Nao lido · <span class="am-row-hint">Clique para abrir no Gmail</span></div>':''}</div>
        <span class="am-open-icon">-></span>
      </a>`;
    }).join('');
  }

  function buildDisparos(disparos) {
    if (!disparos.length) return '<div class="am-empty">Nenhum disparo enviado.</div>';
    return disparos.slice(0, 15).map(d => {
      const href = d.threadUrl || '#';
      const tag = href !== '#' ? 'a' : 'div';
      const attrs = href !== '#' ? `href="${esc(href)}" target="_blank" class="am-row am-row--link"` : `class="am-row"`;
      return `<${tag} ${attrs}><span class="am-avatar am-avatar-disp">E</span><div class="am-row-info"><div class="am-row-top"><span class="am-row-nome">${esc(d.nome)}</span><span class="am-role-tag am-role-disp">DISPARO</span><span class="am-row-data">${esc(d.data)}</span></div><div class="am-row-sub">${esc(d.assunto)}</div></div>${href!=='#'?'<span class="am-open-icon">-></span>':''}</${tag}>`;
    }).join('');
  }

  function buildMovimentacao(d) {
    const todos = [
      ...d.respostas.map(r => ({ ...r, _t: 'resposta', _siId: d.siId })),
      ...d.disparos.map(x  => ({ ...x, _t: 'disparo',  _siId: d.siId }))
    ].sort((a, b) => toComparable(b.data).localeCompare(toComparable(a.data)));

    if (!todos.length) return '<div class="am-empty">Sem movimentacoes.</div>';
    const ul = todos[0];
    const naoLido = ul._t === 'resposta' && !ul.lido;
    const cor = COR_MAP[ul.corLabel] || COR_MAP.azul;
    const href = ul.threadUrl || '#';

    return `<div class="am-mov">
      <a class="am-mov-card ${naoLido?'am-mov-card--unread':''} am-mov-card--link" href="${esc(href)}" target="_blank" data-marcar-lido="${esc(ul.id||ul.threadId||'')}" data-mun-lido="${esc(ul._siId||ul.municipio||'')}">
        <span class="am-mov-icon">${ul._t==='resposta'?'R':'E'}</span>
        <div class="am-mov-card-info">
          <div class="am-mov-tipo">${ul._t==='resposta'?'Resposta recebida':'Disparo enviado'}${naoLido?'<span class="am-pulse">NAO LIDO</span>':''}</div>
          <div class="am-mov-nome" style="color:${cor}">${esc(ul.nome)}</div>
          <div class="am-mov-assunto">${esc(ul.assunto)}</div>
          <div class="am-mov-data">${esc(ul.data)}</div>
          ${naoLido?'<div class="am-mov-hint">Clique para abrir no Gmail</div>':''}
        </div>
        <span class="am-open-icon am-open-icon--mov">-></span>
      </a>
      ${todos.length > 1 ? '<div class="am-mov-hist-title">Historico recente</div>' : ''}
      ${todos.slice(1, 6).map(m => {
        const mnl = m._t === 'resposta' && !m.lido;
        const mhref = m.threadUrl || '#';
        return `<a class="am-mov-hist-row ${mnl?'am-mov-hist--unread':''} ${m.threadUrl?'am-mov-hist--link':''}" href="${esc(mhref)}" target="_blank" data-marcar-lido="${esc(m.id||m.threadId||'')}" data-mun-lido="${esc(m._siId||m.municipio||'')}">
          <span class="am-mov-hist-icon">${m._t==='resposta'?'R':'E'}</span>
          <span class="am-mov-hist-nome">${esc(m.nome)}</span>
          <span class="am-mov-hist-data">${esc(m.data)}</span>
          ${mnl?'<span class="am-hist-dot">*</span>':''}
        </a>`;
      }).join('')}
    </div>`;
  }

  function bindEventos(root) {
    if (root._amClickHandler) root.removeEventListener('click', root._amClickHandler);
    root._amClickHandler = e => {
      const abaEl = e.target.closest('[data-aba]');
      if (abaEl) { e.preventDefault(); e.stopPropagation(); state.abaAtiva[abaEl.dataset.mun] = abaEl.dataset.aba; reRender(abaEl.dataset.mun); return; }
      const linkEl = e.target.closest('[data-marcar-lido]');
      if (linkEl) { const id = linkEl.dataset.marcarLido; const mun = linkEl.dataset.munLido; if (id && mun) marcarLidoLocal(id, mun); return; }
      const toggleEl = e.target.closest('[data-toggle]');
      if (toggleEl) {
        if (e.target.closest('a, button, [data-aba], [data-marcar-lido], .am-si-body')) return;
        const mun = toggleEl.dataset.toggle;
        state.expandidos[mun] = !state.expandidos[mun];
        reRender(mun);
      }
      const delEl = e.target.closest('[data-del-si]');
      if (delEl) {
        e.preventDefault(); e.stopPropagation();
        const siId = delEl.dataset.delSi; const siNome = delEl.dataset.delNome;
        if (!confirm(`Excluir a sub-inbox "${siNome}"?`)) return;
        excluirSubInboxExtensao(siId);
      }
    };
    root.addEventListener('click', root._amClickHandler);
  }

  async function excluirSubInboxExtensao(siId) {
    try {
      const json = await fetchJson(WEB_APP_URL + '?acao=removerSubInbox&id=' + encodeURIComponent(siId));
      if (json && json.sucesso === false) { alert('Erro ao excluir: ' + (json.mensagem || 'falha')); return; }
      delete state.subinboxData[siId];
      const res = await storageGet(['subinboxData']);
      const sd = res.subinboxData || {};
      delete sd[siId];
      await storageSet({ subinboxData: sd });
      const root = document.getElementById('am-root');
      if (root) { root._lastHtml = null; renderRoot(root); }
      sync();
    } catch(e) { alert('Erro de comunicacao: ' + e.message); }
  }

  function marcarLidoLocal(id, mun) {
    if (!id || !mun) return;
    const d = state.subinboxData[mun];
    if (!d) return;
    let mudou = false;
    d.respostas.forEach(r => { const rid = String(r.id || r.threadId || ''); if (rid === id && !r.lido) { r.lido = true; mudou = true; } });
    if (!mudou) return;
    d.novos = d.respostas.filter(r => !r.lido).length;
    state.movimentacoes.forEach(m => { if (String(m.id || m.threadId || '') === id) m.lido = true; });
    reRender(mun);
    storageGet(['lidosConfirmados']).then(res => {
      const arr = res.lidosConfirmados || [];
      if (!arr.includes(id)) arr.push(id);
      storageSet({ lidosConfirmados: arr });
    });
    storageGet(['subinboxData']).then(res => {
      const sd = res.subinboxData || {};
      if (!sd[mun]) return;
      sd[mun].respostas.forEach(r => { if (String(r.id || r.threadId || '') === id) r.lido = true; });
      sd[mun].novos = sd[mun].respostas.filter(r => !r.lido).length;
      storageSet({ subinboxData: sd });
    });
    // ── USA postJson em vez de GM_xmlhttpRequest direto ──
    postJson(WEB_APP_URL, { acao: 'marcarLido', id }).catch(() => {});
  }

  function reRender(mun) {
    const el = document.querySelector(`[data-mun="${CSS_escape(mun)}"]`);
    if (!el) { const root = document.getElementById('am-root'); if (root) { root._lastHtml = null; renderRoot(root); } return; }
    const tmp = document.createElement('div');
    setHTML(tmp, buildSubInbox(mun));
    const novo = tmp.firstElementChild;
    if (novo) setHTML(el, novo.innerHTML);
    const root = document.getElementById('am-root');
    if (root) root._lastHtml = null;
  }

  // ════════════════════════════════════════════════════════════════
  // COMPOSE — DADOS E CACHE
  // ════════════════════════════════════════════════════════════════
  const TIPOS_EMAIL = [
    { id: 'pre_aviso',           icon: '⏰', nome: 'Pré-aviso de Vencimento',   desc: 'Lembrete antecipado de vencimento' },
    { id: 'cobranca_formal',     icon: '📄', nome: 'Cobrança de Honorários',    desc: 'NF + Boleto — cobrança formal' },
    { id: 'nota_fiscal_avulsa',  icon: '🧾', nome: 'Nota Fiscal Avulsa',        desc: 'Envio de nota fiscal específica' },
    { id: 'envio_documento',     icon: '📁', nome: 'Solicitação de Documentos', desc: 'Pedido de documentação ao gestor' },
    { id: 'informativo',         icon: '📢', nome: 'Informativo Jurídico',      desc: 'Comunicado ou atualização jurídica' },
    { id: 'boletim_informativo', icon: '📰', nome: 'Boletim Informativo',       desc: 'Boletim mensal para o município' },
  ];

  let _templatesCache    = null;
  let _templatesFetching = false;
  let _templatesCbs      = [];
  let _configCache       = null;
  let _gestoresCache     = null;
  let _subinboxesCache   = null;

  function buscarTemplates(cb) {
    if (_templatesCache) { cb(null, _templatesCache); return; }
    _templatesCbs.push(cb);
    if (_templatesFetching) return;
    _templatesFetching = true;
    fetchJson(WEB_APP_URL + '?acao=listarTemplates')
      .then(lista => {
        if (!Array.isArray(lista)) throw new Error('Resposta inválida');
        _templatesCache = lista;
        _templatesFetching = false;
        const cbs = _templatesCbs.slice(); _templatesCbs = [];
        cbs.forEach(f => f(null, lista));
      })
      .catch(err => {
        _templatesFetching = false;
        const cbs = _templatesCbs.slice(); _templatesCbs = [];
        cbs.forEach(f => f(err, null));
      });
  }

  function buscarConfigs(cb) {
    if (_configCache) { cb(null, _configCache); return; }
    fetchJson(WEB_APP_URL + '?acao=listarConfiguracoes')
      .then(cfg => { _configCache = cfg; cb(null, cfg); })
      .catch(err => cb(err, null));
  }

  function buscarGestoresESubinboxes(cb) {
    if (_gestoresCache && _subinboxesCache) { cb(null, _gestoresCache, _subinboxesCache); return; }
    Promise.all([
      fetchJson(WEB_APP_URL + '?acao=listarGestores'),
      fetchJson(WEB_APP_URL + '?acao=listarSubInboxes').catch(() => [])
    ]).then(([gestores, subinboxes]) => {
      _gestoresCache   = Array.isArray(gestores)   ? gestores   : [];
      _subinboxesCache = Array.isArray(subinboxes) ? subinboxes : [];
      cb(null, _gestoresCache, _subinboxesCache);
    }).catch(err => cb(err, null, null));
  }

  // ── Helpers ───────────────────────────────────────────────────
  function hojeInputDate() {
    const d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
  }

  function inputDateParaBR(s) {
    if (!s) return '';
    const p = s.split('-');
    if (p.length !== 3) return s;
    return p[2] + '/' + p[1] + '/' + p[0];
  }

  function substituirVariaveis(texto, vars) {
    let s = texto;
    Object.entries(vars).forEach(([k, v]) => { s = s.split(k).join(v || ''); });
    return s;
  }

  // ── Estado do painel ─────────────────────────────────────────
 let _cpState = {
  gestores: [], subinboxes: [], configs: {},
  municipioSel: '', contatoSel: null,
  tipoSel: 'pre_aviso', dataVenc: '', diasAnt: '3',
  horario: (() => {
    const d = new Date();
    return String(d.getHours()).padStart(2,'0') + ':' + String(d.getMinutes()).padStart(2,'0');
  })(),
  observacao: '', msg: '', _jaInjetou: false
};
  // ════════════════════════════════════════════════════════════════
  // COMPOSE — DETECÇÃO DO EDITOR
  // ════════════════════════════════════════════════════════════════
  function encontrarEditorCompose() {
    const doc = (typeof unsafeWindow !== 'undefined' && unsafeWindow.document) || document;
    const direto = doc.querySelector('.Am.aiL.Al.editable');
    if (direto) return direto;
    const fallbacks = [
      '[contenteditable="true"][aria-label*="Message Body"]',
      '[contenteditable="true"][aria-label*="Corpo"]',
      '[contenteditable="true"][g_editable="true"]',
    ];
    for (const sel of fallbacks) {
      const el = doc.querySelector(sel);
      if (el) { const r = el.getBoundingClientRect(); if (r.width > 200 && r.height > 50) return el; }
    }
    return null;
  }

  function encontrarContainerCompose(editor) {
    if (!editor) return null;
    const doc = (typeof unsafeWindow !== 'undefined' && unsafeWindow.document) || document;
    let el = editor.parentElement;
    while (el && el !== doc.body) {
      if (el.classList.contains('nH') && el.classList.contains('Hd')) return el;
      el = el.parentElement;
    }
    el = editor.parentElement;
    while (el && el !== doc.body) {
      const cs = getComputedStyle(el);
      if ((cs.position === 'fixed' || cs.position === 'absolute') && el.offsetHeight > 200) return el;
      el = el.parentElement;
    }
    return null;
  }

  // ════════════════════════════════════════════════════════════════
  // COMPOSE — BOTÃO NA TOOLBAR
  // ════════════════════════════════════════════════════════════════
  function injetarBotaoNoCompose(composeEl) {
    if (!composeEl) return;
    if (composeEl.querySelector('#am-compose-btn')) return;

    const toolbar =
      composeEl.querySelector('.btC') ||
      composeEl.querySelector('.aZ.I5') ||
      composeEl.querySelector('.aZ') ||
      composeEl.querySelector('[role="toolbar"]') ||
      composeEl.querySelector('.aDh');

    if (!toolbar) {
      setTimeout(() => {
        if (_composeEl && !_composeEl.querySelector('#am-compose-btn')) injetarBotaoNoCompose(_composeEl);
      }, 800);
      return;
    }

    const btn = document.createElement('button');
    btn.id = 'am-compose-btn';
    btn.title = 'Templates — Adriana Matos Advocacia';
    btn.style.cssText = [
      'display:inline-flex','align-items:center','gap:5px',
      'padding:4px 10px','margin-left:6px',
      'background:#1c1c1e','color:#34c759',
      'border:1px solid rgba(52,199,89,0.35)',
      'border-radius:6px','font-size:12px','font-weight:700',
      'cursor:pointer','font-family:inherit','letter-spacing:0.2px',
      'vertical-align:middle','transition:background 0.12s,border-color 0.12s',
      'flex-shrink:0',
    ].join(';');
    btn.textContent = '⬇ Templates';

    btn.addEventListener('mouseenter', () => { btn.style.background = '#2c2c2e'; btn.style.borderColor = 'rgba(52,199,89,0.6)'; });
    btn.addEventListener('mouseleave', () => { btn.style.background = '#1c1c1e'; btn.style.borderColor = 'rgba(52,199,89,0.35)'; });
    btn.addEventListener('click', e => {
      e.preventDefault(); e.stopPropagation();
      const doc2 = (typeof unsafeWindow !== 'undefined' && unsafeWindow.document) || document;
      const panel = doc2.getElementById('am-compose-panel');
      if (panel && !panel.classList.contains('am-cp-hidden')) {
        ocultarPainelCompose();
      } else {
        const r = (_composeEl ? _composeEl.getBoundingClientRect() : null) || btn.getBoundingClientRect();
        mostrarPainelCompose(r);
      }
    });

    toolbar.appendChild(btn);
  }

  // ════════════════════════════════════════════════════════════════
  // COMPOSE — OBSERVER
  // ════════════════════════════════════════════════════════════════
  let _composeAberto   = false;
  let _composeEl       = null;
  let _composeRect     = null;
  let _composeObserver = null;
  let _composePolling  = null;

  function verificarEstadoCompose() {
    const editor = encontrarEditorCompose();
    if (editor) {
      const container = encontrarContainerCompose(editor);
      const rect = (container || editor).getBoundingClientRect();
      if (!_composeAberto) {
        _composeAberto = true;
        _composeEl     = container || editor;
        _composeRect   = rect;
        buscarTemplates(() => {});
        buscarGestoresESubinboxes(() => {});
        buscarConfigs(() => {});
        setTimeout(() => { if (_composeEl) injetarBotaoNoCompose(_composeEl); }, 300);
        window.addEventListener('resize', _aoRedimensionar);
      } else {
        if (_composeEl) {
          const novoRect = _composeEl.getBoundingClientRect();
          if (novoRect.left !== _composeRect.left || novoRect.top !== _composeRect.top) {
            _composeRect = novoRect;
            const panel = document.getElementById('am-compose-panel');
            if (panel && !panel.classList.contains('am-cp-hidden')) posicionarPainelCompose(_composeRect);
          }
        }
      }
    } else {
      if (_composeAberto) {
        _composeAberto = false;
        _composeEl     = null;
        _composeRect   = null;
        ocultarPainelCompose();
        window.removeEventListener('resize', _aoRedimensionar);
      }
    }
  }

  function _aoRedimensionar() {
    if (_composeEl) { _composeRect = _composeEl.getBoundingClientRect(); posicionarPainelCompose(_composeRect); }
  }

  function iniciarObserverCompose() {
    if (_composeObserver) return;
    const doc = (typeof unsafeWindow !== 'undefined' && unsafeWindow.document) || document;
    _composeObserver = new MutationObserver(() => verificarEstadoCompose());
    _composeObserver.observe(doc.body, { childList: true, subtree: true });
    _composePolling = setInterval(verificarEstadoCompose, 500);
    verificarEstadoCompose();
  }

  // ════════════════════════════════════════════════════════════════
  // COMPOSE — PAINEL
  // ════════════════════════════════════════════════════════════════
  function criarPainelCompose() {
    const doc = (typeof unsafeWindow !== 'undefined' && unsafeWindow.document) || document;
    let panel = doc.getElementById('am-compose-panel');
    if (panel) return panel;
    panel = doc.createElement('div');
    panel.id = 'am-compose-panel';
    panel.className = 'am-cp-hidden';
    panel.style.setProperty('z-index', '2147483647', 'important');
    panel.style.setProperty('position', 'fixed', 'important');
    doc.body.appendChild(panel);
    doc.addEventListener('mousedown', e => {
      const p2  = doc.getElementById('am-compose-panel');
      const btn = doc.getElementById('am-compose-btn');
      if (p2 && !p2.contains(e.target) && (!btn || !btn.contains(e.target))) ocultarPainelCompose();
    }, true);
    return panel;
  }

  function posicionarPainelCompose(rect) {
    const doc = (typeof unsafeWindow !== 'undefined' && unsafeWindow.document) || document;
    const panel = doc.getElementById('am-compose-panel');
    if (!panel) return;
    const vw = window.innerWidth, vh = window.innerHeight;
    const panelW = 300, gap = 8;
    const panelH = Math.min(580, vh - 32);
    let left = rect.left - panelW - gap;
    let top  = rect.top;
    if (left < 8) left = rect.right + gap;
    if (left + panelW > vw - 8) { left = Math.max(8, rect.left); top = rect.top - panelH - gap; }
    if (top + panelH > vh - 8) top = vh - panelH - 8;
    if (top < 8) top = 8;
    panel.style.left      = left + 'px';
    panel.style.top       = top  + 'px';
    panel.style.maxHeight = panelH + 'px';
    panel.style.width     = panelW + 'px';
  }

 function mostrarPainelCompose(rect) {
    const panel = criarPainelCompose();
    posicionarPainelCompose(rect);
    panel.classList.remove('am-cp-hidden');

    // Atualiza horário para o momento atual ao abrir (se o usuário não tiver mudado manualmente)
    if (!_cpState._horarioManual) {
      const d = new Date();
      _cpState.horario = String(d.getHours()).padStart(2,'0') + ':' + String(d.getMinutes()).padStart(2,'0');
    }

    setPainelStatus('Carregando contatos...', '');
    renderPainelCompose();
    buscarGestoresESubinboxes((err, gestores, subinboxes) => {
      if (err) { setPainelStatus('Erro ao carregar: ' + err.message, 'err'); return; }
      _cpState.gestores   = gestores   || [];
      _cpState.subinboxes = subinboxes || [];
      buscarConfigs(() => {});
      buscarTemplates(() => {});
      if (!_cpState.dataVenc) _cpState.dataVenc = hojeInputDate();
      _cpState.msg = '';
      renderPainelCompose();
    });
  }

  function ocultarPainelCompose() {
    const doc = (typeof unsafeWindow !== 'undefined' && unsafeWindow.document) || document;
    const panel = doc.getElementById('am-compose-panel');
    if (panel) panel.classList.add('am-cp-hidden');
  }

  function setPainelStatus(msg, tipo) {
    const doc = (typeof unsafeWindow !== 'undefined' && unsafeWindow.document) || document;
    const el = doc.getElementById('am-cp-status');
    if (!el) return;
    el.textContent = msg;
    el.className = tipo === 'ok' ? 'am-cp-status-ok' : tipo === 'err' ? 'am-cp-status-err' : '';
  }

 function renderPainelCompose() {
    const doc = (typeof unsafeWindow !== 'undefined' && unsafeWindow.document) || document;
    const panel = doc.getElementById('am-compose-panel');
    if (!panel) return;
    const s = _cpState;

    const municipiosSet = new Set();
    s.gestores.forEach(g => { if (g.Municipio) municipiosSet.add(g.Municipio); });
    s.subinboxes.forEach(si => { if (si.Municipio) municipiosSet.add(si.Municipio); });
    const municipios = [...municipiosSet].sort();
    const contatosFiltrados = s.municipioSel ? s.gestores.filter(g => g.Municipio === s.municipioSel) : [];

    const tipo       = s.tipoSel;
    const isPreAviso = tipo === 'pre_aviso';
    const temDias    = isPreAviso;
    const temObs     = ['nota_fiscal_avulsa','envio_documento','informativo','boletim_informativo'].includes(tipo);
    const labelData  = isPreAviso ? 'Data de Vencimento' : 'Data do Disparo';
    const labelObs   = tipo === 'envio_documento' ? 'Documento solicitado'
                     : tipo === 'nota_fiscal_avulsa' ? 'Descrição da NF'
                     : 'Assunto / Observação';

    // Calcula se o horário escolhido é futuro para mostrar botão correto
    const agora = new Date();
    const [hh, mm] = (s.horario || '08:00').split(':').map(Number);
    const partes = (s.dataVenc || hojeInputDate()).split('-');
    const dataDisparoObj = partes.length === 3
      ? new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, parseInt(partes[2]), hh, mm, 0)
      : agora;
    const isFuturo = dataDisparoObj > agora;

    const html = `
<div id="am-cp-header">
  <span id="am-cp-header-title">Adriana Matos — Disparo</span>
  <button id="am-cp-close" title="Fechar">✕</button>
</div>
<div id="am-cp-body">
  <span class="am-cp-section-label">Município</span>
  <select class="am-cp-select" id="am-cp-municipio">
    <option value="">Selecione o município...</option>
    ${municipios.map(m => `<option value="${esc2(m)}" ${s.municipioSel===m?'selected':''}>${esc2(m)}</option>`).join('')}
  </select>
  <span class="am-cp-section-label">Contato</span>
  <select class="am-cp-select" id="am-cp-contato" ${!s.municipioSel?'disabled':''}>
    <option value="">Selecione o contato...</option>
    ${contatosFiltrados.map(g => {
      const sel = s.contatoSel && String(s.contatoSel.ID) === String(g.ID) ? 'selected' : '';
      return `<option value="${esc2(String(g.ID))}" ${sel}>${esc2(g.Nome)} — ${esc2(g.Status||'')}</option>`;
    }).join('')}
  </select>
  ${s.contatoSel ? `<div class="am-cp-email-hint">📧 ${esc2(s.contatoSel.Email)}</div>` : ''}
  <span class="am-cp-section-label">Tipo de Email</span>
  <select class="am-cp-select" id="am-cp-tipo">
    ${TIPOS_EMAIL.map(t => `<option value="${t.id}" ${s.tipoSel===t.id?'selected':''}>${t.icon} ${esc2(t.nome)}</option>`).join('')}
  </select>
  <span class="am-cp-section-label">${labelData}</span>
  <input class="am-cp-input" type="date" id="am-cp-data" value="${s.dataVenc || hojeInputDate()}"/>
  ${temDias ? `<span class="am-cp-section-label">Dias de antecedência</span><input class="am-cp-input" type="number" id="am-cp-dias" value="${s.diasAnt||'3'}" min="1" max="30" style="width:80px;"/>` : ''}
  <span class="am-cp-section-label">Horário do disparo</span>
  <input class="am-cp-input" type="time" id="am-cp-horario" value="${s.horario||'08:00'}"/>
  ${temObs ? `<span class="am-cp-section-label">${labelObs}</span><input class="am-cp-input" type="text" id="am-cp-obs" value="${esc2(s.observacao)}" placeholder="Ex: NF 001/2025"/>` : ''}
</div>
<div id="am-cp-footer" style="display:flex;flex-direction:column;gap:6px;padding:10px 12px 8px;border-top:1px solid rgba(255,255,255,0.06);">
  <button id="am-cp-injetar" ${!s.contatoSel?'disabled':''} style="width:100%;background:#34c759;color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:700;font-family:inherit;padding:10px 14px;cursor:pointer;">
    ${!s.contatoSel ? '⬇ Selecione o contato' : '⬇ Injetar template no email'}
  </button>
  <button id="am-cp-agendar" ${!s.contatoSel?'disabled':''} style="width:100%;background:${isFuturo?'#007aff':'#3a3a3c'};color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:700;font-family:inherit;padding:10px 14px;cursor:pointer;">
    ${!s.contatoSel ? '📅 Selecione o contato' : isFuturo ? '📅 Confirmar agendamento' : '📅 Só agendar (sem injetar)'}
  </button>
</div>
<div id="am-cp-status">${s.msg}</div>
`;

    const policy = getTTPolicy();
    if (policy) panel.innerHTML = policy.createHTML(html);
    else panel.innerHTML = html;

    const q = id => doc.getElementById(id);

    q('am-cp-close').onclick = () => ocultarPainelCompose();

    q('am-cp-municipio').onchange = function() {
      _cpState.municipioSel = this.value;
      _cpState.contatoSel   = null;
      renderPainelCompose();
    };

    q('am-cp-contato').onchange = function() {
      const id = this.value;
      _cpState.contatoSel = _cpState.gestores.find(g => String(g.ID) === id) || null;
      if (_cpState.contatoSel) preencherDestinatario(_cpState.contatoSel.Email);
      renderPainelCompose();
    };

    q('am-cp-tipo').onchange = function() {
      _cpState.tipoSel = this.value;
      // ── Troca ao vivo: reinjecta imediatamente se já havia injetado ──
      if (_cpState._jaInjetou && _cpState.contatoSel) {
        executarInjecaoSilenciosa();
      } else {
        renderPainelCompose();
      }
    };

    q('am-cp-data').onchange = function() { _cpState.dataVenc = this.value; renderPainelCompose(); };
    if (q('am-cp-dias')) q('am-cp-dias').onchange = function() { _cpState.diasAnt = this.value; };
   q('am-cp-horario').onchange = function() {
  _cpState.horario = this.value;
  _cpState._horarioManual = true;
  renderPainelCompose();
};
    if (q('am-cp-obs')) q('am-cp-obs').oninput = function() { _cpState.observacao = this.value; };

    if (q('am-cp-injetar')) q('am-cp-injetar').onclick = () => executarInjecao(false);
    if (q('am-cp-agendar')) q('am-cp-agendar').onclick = () => executarInjecao(true);
  }

  // Injeta o template no editor sem registrar no WebApp (para troca ao vivo)
  function executarInjecaoSilenciosa() {
    const s = _cpState;
    if (!s.contatoSel) return;

    const doc = (typeof unsafeWindow !== 'undefined' && unsafeWindow.document) || document;
    const dataEl = doc.getElementById('am-cp-data');
    const diasEl = doc.getElementById('am-cp-dias');
    const obsEl  = doc.getElementById('am-cp-obs');
    const dataVenc  = (dataEl ? dataEl.value : '') || s.dataVenc || hojeInputDate();
    const diasAnt   = (diasEl ? diasEl.value : '') || s.diasAnt  || '3';
    const obs       = (obsEl  ? obsEl.value  : '') || s.observacao || '';
    const dataVencBR = inputDateParaBR(dataVenc);

    Promise.all([
      new Promise((res, rej) => buscarTemplates((e, d) => e ? rej(e) : res(d))),
      new Promise((res, rej) => buscarConfigs((e, d)   => e ? rej(e) : res(d))),
    ]).then(([templates, configs]) => {
      const template = templates.find(t => String(t.tipo||t[0]||'').trim() === s.tipoSel);
      if (!template) return;

      const imagemUrl = String(configs.imagem_topo_url || '');
      const vars = {
        '{{nome_gestor}}':         s.contatoSel.Nome       || '',
        '{{municipio}}':           s.contatoSel.Municipio  || '',
        '{{data_vencimento}}':     dataVencBR,
        '{{dias_restantes}}':      diasAnt,
        '{{descricao_documento}}': obs || 'Documento',
        '{{assunto_informativo}}': obs || 'Informativo',
        '{{nome_escritorio}}':     String(configs.nome_escritorio || 'Adriana Matos Advocacia'),
        '{{oab}}':                 String(configs.oab      || ''),
        '{{telefone}}':            String(configs.telefone || ''),
        '{{imagem_topo_url}}':     imagemUrl,
      };

      let corpoFinal = substituirVariaveis(String(template.corpo||template[2]||''), vars);
      if (!imagemUrl) {
        corpoFinal = corpoFinal.replace(/<img[^>]*onerror[^>]*\/?>/gi, '');
        corpoFinal = corpoFinal.replace(/<img[^>]*>/gi, '');
      }

      const editor = encontrarEditorCompose();
      if (!editor) return;
      const win = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
      const policy = getTTPolicy();
      if (policy) { editor.innerHTML = policy.createHTML(corpoFinal); }
      else {
        const parsed = new win.DOMParser().parseFromString(corpoFinal, 'text/html');
        while (editor.firstChild) editor.removeChild(editor.firstChild);
        Array.from(parsed.body.childNodes).forEach(n => editor.appendChild(win.document.importNode(n, true)));
      }
      ['input','keyup','change'].forEach(ev => {
        try { editor.dispatchEvent(new win.Event(ev, { bubbles: true })); } catch(e) {}
      });
      setPainelStatus('↺ Template atualizado: ' + (TIPOS_EMAIL.find(t=>t.id===s.tipoSel)||{}).nome, 'ok');
    }).catch(() => {});

    renderPainelCompose();
  }

  // ════════════════════════════════════════════════════════════════
  // COMPOSE — PREENCHER DESTINATÁRIO
  // ════════════════════════════════════════════════════════════════
  function preencherDestinatario(email) {
    if (!email) return;
    const doc = (typeof unsafeWindow !== 'undefined' && unsafeWindow.document) || document;
    const win = (typeof unsafeWindow !== 'undefined' ? unsafeWindow : window);
    const toInputs = doc.querySelectorAll('input[name="to"], input[aria-label*="Para"], input[aria-label*="To"]');
    toInputs.forEach(inp => {
      if (inp.value.includes(email)) return;
      inp.focus();
      inp.value = email;
      inp.dispatchEvent(new win.KeyboardEvent('keydown', { bubbles: true, key: 'Enter', keyCode: 13 }));
      inp.dispatchEvent(new win.Event('input', { bubbles: true }));
      inp.dispatchEvent(new win.KeyboardEvent('keydown', { bubbles: true, key: 'Enter', keyCode: 13 }));
    });
  }

  // ════════════════════════════════════════════════════════════════
  // COMPOSE — INJEÇÃO DO TEMPLATE
  // ════════════════════════════════════════════════════════════════
  function executarInjecao(apenasAgendar) {
    const s = _cpState;
    if (!s.contatoSel) { setPainelStatus('Selecione um contato primeiro.', 'err'); return; }

    const doc = (typeof unsafeWindow !== 'undefined' && unsafeWindow.document) || document;
    const dataEl = doc.getElementById('am-cp-data');
    const diasEl = doc.getElementById('am-cp-dias');
    const obsEl  = doc.getElementById('am-cp-obs');
    const dataVenc  = (dataEl ? dataEl.value : '') || s.dataVenc || hojeInputDate();
    const diasAnt   = (diasEl ? diasEl.value : '') || s.diasAnt  || '3';
    const obs       = (obsEl  ? obsEl.value  : '') || s.observacao || '';
    const horario   = s.horario || '08:00';
    const dataVencBR = inputDateParaBR(dataVenc);

    if (apenasAgendar) {
      // ── Só agenda no WebApp, não mexe no Gmail ──
      const btnAg = doc.getElementById('am-cp-agendar');
      if (btnAg) { btnAg.disabled = true; btnAg.textContent = '⏳ Agendando...'; }

      postJson(WEB_APP_URL, {
        acao: 'agendarDisparo',
        gestorId:          s.contatoSel.ID        || '',
        nomeGestor:        s.contatoSel.Nome       || '',
        municipio:         s.contatoSel.Municipio  || '',
        emailDestinatario: s.contatoSel.Email      || '',
        tipoEmail:         s.tipoSel,
        dataVencimento:    dataVenc,
        diasAntecedencia:  diasAnt,
        horarioDisparo:    horario,
        observacao:        obs,
        recorrente:        false,
        diaDoMes:          ''
      }).then(res => {
        const dataF = dataVenc.split('-').reverse().join('/');
        setPainelStatus(`✓ Agendado para ${dataF} às ${horario} — sistema envia automaticamente.`, 'ok');
        _cpState.msg = '✓ Agendado — ' + s.contatoSel.Nome;
        const btnA = doc.getElementById('am-cp-agendar');
        if (btnA) { btnA.disabled = false; btnA.textContent = '📅 Agendar outro'; }
        sync();
      }).catch(err => {
        setPainelStatus('Erro ao agendar: ' + err.message, 'err');
        const btnA = doc.getElementById('am-cp-agendar');
        if (btnA) { btnA.disabled = false; btnA.textContent = '📅 Confirmar agendamento'; }
      });
      return;
    }

    // ── Injeção imediata no Gmail — NÃO registra no WebApp ainda ──
    setPainelStatus('Buscando template...', '');
    const btnInj = doc.getElementById('am-cp-injetar');
    if (btnInj) { btnInj.disabled = true; btnInj.textContent = '⏳ Aguarde...'; }

    Promise.all([
      new Promise((res, rej) => buscarTemplates((e, d) => e ? rej(e) : res(d))),
      new Promise((res, rej) => buscarConfigs((e, d)   => e ? rej(e) : res(d))),
    ]).then(([templates, configs]) => {
      const template = templates.find(t => String(t.tipo||t[0]||'').trim() === s.tipoSel);
      if (!template) throw new Error('Template não encontrado: ' + s.tipoSel);

      const imagemUrl = String(configs.imagem_topo_url || '');
      const vars = {
        '{{nome_gestor}}':         s.contatoSel.Nome       || '',
        '{{municipio}}':           s.contatoSel.Municipio  || '',
        '{{data_vencimento}}':     dataVencBR,
        '{{dias_restantes}}':      diasAnt,
        '{{descricao_documento}}': obs || 'Documento',
        '{{assunto_informativo}}': obs || 'Informativo',
        '{{nome_escritorio}}':     String(configs.nome_escritorio || 'Adriana Matos Advocacia'),
        '{{oab}}':                 String(configs.oab      || ''),
        '{{telefone}}':            String(configs.telefone || ''),
        '{{imagem_topo_url}}':     imagemUrl,
      };

      const assuntoFinal = substituirVariaveis(String(template.assunto||template[1]||''), vars);
      let   corpoFinal   = substituirVariaveis(String(template.corpo  ||template[2]||''), vars);
      if (!imagemUrl) {
        corpoFinal = corpoFinal.replace(/<img[^>]*onerror[^>]*\/?>/gi, '');
        corpoFinal = corpoFinal.replace(/<img[^>]*>/gi, '');
      }

      const editor = encontrarEditorCompose();
      if (!editor) throw new Error('Editor do Gmail não encontrado. Abra o Compose.');

      const win       = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
      const targetDoc = win.document;
      const container = encontrarContainerCompose(editor);

      const assuntoInput =
        (container||targetDoc).querySelector('input.aoT') ||
        (container||targetDoc).querySelector('[name="subjectbox"]') ||
        (container||targetDoc).querySelector('input[aria-label*="Assunto"]') ||
        (container||targetDoc).querySelector('input[aria-label*="Subject"]');

      if (assuntoInput && !assuntoInput.value.trim()) {
        assuntoInput.value = assuntoFinal;
        assuntoInput.dispatchEvent(new win.Event('input',  { bubbles: true }));
        assuntoInput.dispatchEvent(new win.Event('change', { bubbles: true }));
      }

      preencherDestinatario(s.contatoSel.Email);
      editor.focus();

      try {
        const policy = getTTPolicy();
        if (policy) { editor.innerHTML = policy.createHTML(corpoFinal); }
        else {
          const parsed = new win.DOMParser().parseFromString(corpoFinal, 'text/html');
          while (editor.firstChild) editor.removeChild(editor.firstChild);
          Array.from(parsed.body.childNodes).forEach(n => editor.appendChild(targetDoc.importNode(n, true)));
        }
      } catch(eInj) { throw new Error('Erro ao injetar corpo: ' + eInj.message); }

      setTimeout(() => {
        ['input','keyup','change'].forEach(ev => {
          try { editor.dispatchEvent(new win.Event(ev, { bubbles: true })); } catch(e3) {}
        });

        _cpState._jaInjetou = true;
        setPainelStatus('✓ Template injetado — envie pelo Gmail e clique em "Registrar envio".', 'ok');
        _cpState.msg = '✓ ' + s.contatoSel.Nome;

        // ── Troca botão injetar por "Registrar envio" SEM recriar o painel ──
        const btnI = doc.getElementById('am-cp-injetar');
        if (btnI) {
          btnI.disabled = false;
          btnI.textContent = '✓ Registrar envio no sistema';
          btnI.style.background = '#007aff';
          // Remove handlers antigos clonando o botão
          const btnClone = btnI.cloneNode(true);
          btnI.parentNode.replaceChild(btnClone, btnI);
          btnClone.addEventListener('click', () => {
            btnClone.disabled = true;
            btnClone.textContent = '⏳ Registrando...';
            postJson(WEB_APP_URL, {
              acao:              'agendarDisparo',
              gestorId:          s.contatoSel.ID        || '',
              nomeGestor:        s.contatoSel.Nome       || '',
              municipio:         s.contatoSel.Municipio  || '',
              emailDestinatario: s.contatoSel.Email      || '',
              tipoEmail:         s.tipoSel,
              dataVencimento:    dataVenc,
              diasAntecedencia:  diasAnt,
              horarioDisparo:    horario,
              observacao:        obs,
              recorrente:        false,
              diaDoMes:          ''
            }).then(() => {
              setPainelStatus('✓ Envio registrado no sistema!', 'ok');
              _cpState._jaInjetou = false;
              _cpState._horarioManual = false;
              sync();
              renderPainelCompose();
            }).catch(() => {
              setPainelStatus('Erro ao registrar envio.', 'err');
              btnClone.disabled = false;
              btnClone.textContent = '✓ Registrar envio no sistema';
            });
          });
        }
      }, 80);


    }).catch(err => {
      setPainelStatus('Erro: ' + err.message, 'err');
      const btnI = doc.getElementById('am-cp-injetar');
      if (btnI) { btnI.disabled = false; btnI.textContent = '⬇ Injetar template no email'; }
    });
  }

  // ════════════════════════════════════════════════════════════════
  // BOOT
  // ════════════════════════════════════════════════════════════════
  async function boot() {
    injetarCSS();
    await carregarDados();
    lastURL = location.href;
    tentarInjetar();
    iniciarNavWatch();
    iniciarSyncPeriodico();
    iniciarObserverCompose();
    sync();
  }

  boot();

})();
