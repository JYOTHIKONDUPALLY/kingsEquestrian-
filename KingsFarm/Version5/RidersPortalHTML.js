// ============================================================
// KINGS EQUESTRIAN — NEW SYSTEM
// File: 8_RiderPortalHTML.gs
// Changes:
//   Change 1: Time slots in 30-min increments
//   Change 6: Multi-profile picker when phone matches multiple riders
//   Change 7: Participants field shown only for One-Time services
//   Change 8: Allow backdating — removed min-date restrictions
// ============================================================

function getRiderPortalHtml() {
  var servicesList = [];
  try { servicesList = getServicesList(); } catch(e) { Logger.log('getServicesList: ' + e); }

  var safeServices = servicesList.map(function(s) {
    return {
      name    : String(s.name  || '').replace(/"/g, '&quot;'),
      price   : Number(s.price || 0),
      type    : String(s.type  || 'Regular').replace(/"/g, '&quot;'),
      // Change 7: show pax ONLY for One-Time type
      showPax : true
    };
  });
  var servicesJson = JSON.stringify(safeServices);
  var portalCfg = JSON.stringify({
    upiId         : CONFIG.UPI_ID || '',
    businessName  : CONFIG.BUSINESS_NAME || 'KingsEquestrian',
    advanceAmount : Number(CONFIG.ADVANCE_BOOKING_AMOUNT) || 1000
  });
  var appUiVersion = String((CONFIG && CONFIG.APP_UI_VERSION) || '');
  var buildStamp = String(Date.now());
  var buildLabel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  return _portalHTML(servicesJson, portalCfg, appUiVersion, buildStamp, buildLabel);
}

function _portalHTML(servicesJson, portalCfg, appUiVersion, buildStamp, buildLabel) {
  appUiVersion = String(appUiVersion || '');
  buildStamp = String(buildStamp || Date.now());
  buildLabel = String(buildLabel || '');

  var css = ''
    + '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}'
    + ':root{'
    + '--ink:#0a1f16;--forest:#0f3526;--pine:#1a5c3a;--sage:#2e8a5c;'
    + '--fern:#4aab7a;--mint:#8fd4b0;--mist:#c5eada;--dew:#e8f7f0;'
    + '--parchment:#f5f8f5;--white:#ffffff;'
    + '--gold:#b8860b;--gold-pale:#fdf6e3;--gold-border:#e8d48a;'
    + '--red:#b91c1c;--red-pale:#fef2f2;'
    + '--border:rgba(42,120,80,0.13);--border-md:rgba(42,120,80,0.22);--muted:#7a9a7e;'
    + '--shadow-xs:0 1px 3px rgba(10,31,22,.07);--shadow-sm:0 2px 8px rgba(10,31,22,.1);--shadow-md:0 12px 40px rgba(10,31,22,.22);'
    + '--r:14px;--r-sm:10px}'
    + 'html,body{min-height:100%;-webkit-font-smoothing:antialiased}'
    + 'body{font-family:"DM Sans",sans-serif;background:var(--parchment);color:var(--ink)}'
    // login
    + '#login-screen{min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:2rem 1.25rem 4rem;background:var(--forest);position:relative;overflow:hidden}'
    + '.login-ring{position:absolute;border-radius:50%;border:1px solid rgba(143,212,176,.1);pointer-events:none}'
    + '.login-card{width:100%;max-width:380px;background:var(--white);border-radius:20px;padding:2rem 1.75rem 1.75rem;box-shadow:0 6px 20px rgba(10,31,22,.14);position:relative;z-index:1}'
    + '.brand-block{text-align:center;margin-bottom:1.75rem}'
    + '.brand-icon{width:64px;height:64px;border-radius:16px;background:var(--forest);display:flex;align-items:center;justify-content:center;font-size:30px;margin:0 auto 1rem;border:2px solid rgba(143,212,176,.3)}'
    + '.brand-title{font-family:"Playfair Display",serif;font-size:26px;font-weight:600;color:var(--ink);line-height:1.1;margin-bottom:4px}'
    + '.brand-sub{font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:#7a9a7e;font-weight:500}'
    + '.fg{margin-bottom:14px}'
    + '.fl{display:block;font-size:11px;font-weight:600;letter-spacing:.07em;text-transform:uppercase;color:var(--sage);margin-bottom:6px}'
    + '.fi{width:100%;padding:11px 14px;border:1.5px solid var(--border-md);border-radius:var(--r-sm);font-family:"DM Sans",sans-serif;font-size:14px;color:var(--ink);background:var(--parchment);outline:none;-webkit-appearance:none;transition:all .15s}'
    + '.fi:focus{border-color:var(--sage);box-shadow:0 0 0 3px rgba(46,138,92,.1);background:var(--white)}'
    + '.fi::placeholder{color:#b0c8b8}'
    + 'select.fi{cursor:pointer}'
    + '.btn-p{width:100%;padding:13px;border-radius:var(--r-sm);background:var(--forest);color:var(--mist);border:none;font-family:"DM Sans",sans-serif;font-size:14px;font-weight:600;cursor:pointer;margin-top:4px;transition:all .15s}'
    + '.btn-p:disabled{opacity:.5;cursor:default}'
    + '.login-err{font-size:12px;color:var(--red);text-align:center;margin-top:10px;padding:9px 12px;background:var(--red-pale);border-radius:8px;border:1px solid #fecaca;display:none}'
    + '.login-note{font-size:11px;color:#9aaa9e;text-align:center;margin-top:10px;line-height:1.7}'
    + '.build-lbl{font-size:10px;color:#9aaa9e;text-align:center;margin-top:8px;letter-spacing:.04em}'
    + '.stale-banner{display:none;background:#92400e;color:#fffbeb;padding:10px 14px;font-size:12px;font-weight:600;text-align:center;border-bottom:1px solid #f59e0b}'
    + '.stale-banner.on{display:block}'
    + '.stale-banner button{margin-left:8px;padding:6px 12px;border:0;border-radius:8px;background:#fffbeb;color:#92400e;font:inherit;font-weight:700;cursor:pointer}'
    // Change 6: profile picker
    + '#profile-picker{display:none;min-height:100vh;background:var(--forest);align-items:center;justify-content:center;padding:2rem 1.25rem;flex-direction:column}'
    + '.picker-card{width:100%;max-width:400px;background:var(--white);border-radius:20px;padding:1.75rem;box-shadow:0 6px 20px rgba(10,31,22,.14)}'
    + '.picker-title{font-family:"Playfair Display",serif;font-size:20px;font-weight:600;color:var(--ink);margin-bottom:6px;text-align:center}'
    + '.picker-sub{font-size:12px;color:#7a9a7e;text-align:center;margin-bottom:20px}'
    + '.profile-btn{width:100%;background:var(--dew);border:1.5px solid var(--mist);border-radius:12px;padding:14px 16px;margin-bottom:10px;cursor:pointer;text-align:left;font-family:"DM Sans",sans-serif;transition:all .15s}'
    + '.profile-btn:hover{background:var(--mist);border-color:var(--fern)}'
    + '.profile-name{font-size:15px;font-weight:600;color:var(--pine);margin-bottom:3px}'
    + '.profile-meta{font-size:11px;color:#7a9a7e}'
    + '.btn-back{background:none;border:1px solid rgba(143,212,176,.3);color:var(--mist);font-size:11px;padding:5px 14px;border-radius:20px;cursor:pointer;font-family:"DM Sans",sans-serif;margin-top:6px;width:100%}'
    // dashboard
    + '#dashboard{display:none;min-height:100vh}'
    + '.dh{background:var(--forest);padding:14px 16px;display:flex;align-items:center;gap:12px;position:sticky;top:0;z-index:100;border-bottom:1px solid rgba(143,212,176,.1)}'
    + '.av{width:38px;height:38px;border-radius:11px;background:linear-gradient(135deg,var(--sage),var(--fern));display:flex;align-items:center;justify-content:center;font-family:"Playfair Display",serif;font-size:14px;font-weight:700;color:#fff;flex-shrink:0;border:1.5px solid rgba(143,212,176,.3)}'
    + '.dh-info{min-width:0;flex:1}'
    + '.dh-name{font-family:"Playfair Display",serif;font-size:16px;font-weight:600;color:var(--mist);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}'
    + '.dh-sub{font-size:10px;color:var(--mint);margin-top:1px}'
    + '.btn-lo{background:rgba(143,212,176,.1);border:1px solid rgba(143,212,176,.2);color:var(--mint);padding:5px 11px;border-radius:20px;font-size:11px;font-family:"DM Sans",sans-serif;cursor:pointer;flex-shrink:0}'
    // banner
    + '.info-banner{background:linear-gradient(135deg,var(--forest) 0%,var(--pine) 100%);padding:14px 16px 16px}'
    + '.info-ke{font-family:"Playfair Display",serif;font-size:18px;font-weight:600;color:var(--mist);margin-bottom:2px}'
    + '.info-svc{font-size:11px;color:var(--mint);margin-bottom:12px}'
    + '.info-stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:8px}'
    + '.istat{background:rgba(143,212,176,.1);border:1px solid rgba(143,212,176,.18);border-radius:10px;padding:9px 8px;text-align:center}'
    + '.istat-val{font-family:"Playfair Display",serif;font-size:22px;font-weight:700;color:var(--mint);line-height:1}'
    + '.istat-lbl{font-size:9px;color:rgba(143,212,176,.7);text-transform:uppercase;letter-spacing:.08em;margin-top:3px}'
    // tabs
    + '.tabs{background:var(--white);display:flex;border-bottom:1.5px solid var(--border);position:sticky;top:66px;z-index:99}'
    + '.tb{flex:1;padding:12px 6px;background:none;border:none;border-bottom:2.5px solid transparent;font-family:"DM Sans",sans-serif;font-size:12px;color:#8aaa8e;cursor:pointer;font-weight:500;transition:all .15s;display:flex;align-items:center;justify-content:center;gap:5px}'
    + '.tb.on{color:var(--pine);border-bottom-color:var(--pine);font-weight:600}'
    + '.tc{display:none;padding:16px 14px}.tc.on{display:block}'
    + '.tab-heading{font-family:"Playfair Display",serif;font-size:22px;font-weight:600;color:var(--ink);margin-bottom:2px}'
    + '.tab-sub{font-size:12px;color:#7a9a7e;margin-bottom:14px}'
    + '.sec-div{display:flex;align-items:center;gap:8px;margin:16px 0 10px;color:#b0c8b8;font-size:10px;letter-spacing:.1em;text-transform:uppercase;font-weight:600}'
    + '.sec-div::before,.sec-div::after{content:"";flex:1;height:1px;background:var(--border)}'
    // session cards
    + '.sc{background:var(--white);border:1px solid var(--border);border-radius:var(--r);padding:14px;margin-bottom:9px;box-shadow:var(--shadow-xs);overflow:hidden;position:relative}'
    + '.sc::before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:var(--border)}'
    + '.sc.s-up::before{background:var(--fern)}'
    + '.sc.s-att::before{background:var(--sage)}'
    + '.sc.s-ns::before{background:var(--red)}'
    + '.sc.s-rs::before{background:var(--gold)}'
    + '.si{padding-left:10px}'
    + '.sc-top{display:flex;justify-content:space-between;align-items:flex-start;gap:8px;margin-bottom:6px}'
    + '.sc-svc{font-size:14px;font-weight:600;color:var(--ink)}'
    + '.sc-dt{font-size:12px;color:var(--sage);margin-top:2px}'
    + '.sc-tm{font-size:11px;color:#7a9a7e;margin-top:3px}'
    + '.bdg{font-size:10px;font-weight:600;padding:3px 9px;border-radius:20px;flex-shrink:0;letter-spacing:.04em}'
    + '.b-up{background:var(--dew);color:var(--pine);border:1px solid var(--mist)}'
    + '.b-dn{background:#f3f4f6;color:#6b7280;border:1px solid #e5e7eb}'
    + '.b-ns{background:var(--red-pale);color:var(--red);border:1px solid #fecaca}'
    + '.b-rs{background:var(--gold-pale);color:var(--gold);border:1px solid var(--gold-border)}'
    // reschedule
    + '.rs-trigger{margin-top:10px;display:inline-flex;align-items:center;gap:6px;font-size:11px;font-weight:600;color:var(--sage);background:var(--dew);border:1px solid var(--mist);padding:5px 12px;border-radius:20px;cursor:pointer}'
    + '.rs-panel{display:none;margin-top:10px;padding:14px;background:var(--parchment);border:1px solid var(--border-md);border-radius:var(--r-sm)}'
    + '.rs-panel.open{display:block}'
    + '.rs-from{font-size:12px;color:var(--sage);margin-bottom:12px;padding:8px 10px;background:var(--dew);border-radius:8px;border:1px solid var(--mist)}'
    // book
    + '.mode-toggle{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:16px}'
    + '.mode-btn{padding:10px 8px;border-radius:var(--r-sm);border:1.5px solid var(--border-md);background:var(--white);font-family:"DM Sans",sans-serif;font-size:12px;font-weight:500;color:#7a9a7e;cursor:pointer;text-align:center;display:flex;flex-direction:column;align-items:center;gap:4px}'
    + '.mode-btn .mi{font-size:20px}'
    + '.mode-btn.active{border-color:var(--sage);background:var(--dew);color:var(--pine);font-weight:600}'
    + '.slot-card{background:var(--white);border:1px solid var(--border);border-radius:var(--r);padding:14px;margin-bottom:10px;position:relative}'
    + '.slot-num{font-size:10px;font-weight:700;color:var(--sage);text-transform:uppercase;letter-spacing:.08em;margin-bottom:12px;display:flex;align-items:center;gap:6px}'
    + '.sn-badge{background:var(--pine);color:var(--mist);font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px}'
    + '.rm-slot{position:absolute;right:12px;top:12px;background:none;border:1px solid #fecaca;color:var(--red);font-size:12px;cursor:pointer;padding:3px 8px;border-radius:6px;font-family:"DM Sans",sans-serif}'
    + '.add-btn{width:100%;padding:11px;border:2px dashed var(--mist);background:transparent;color:var(--sage);border-radius:var(--r-sm);font-family:"DM Sans",sans-serif;font-size:13px;font-weight:500;cursor:pointer;margin-bottom:10px;display:flex;align-items:center;justify-content:center;gap:6px}'
    // recurring
    + '.rec-card{background:var(--white);border:1px solid var(--border);border-radius:var(--r);padding:16px;margin-bottom:10px}'
    + '.day-picker{display:flex;gap:6px;flex-wrap:wrap;margin:10px 0}'
    + '.day-chip{width:40px;height:40px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:600;cursor:pointer;border:1.5px solid var(--border-md);background:var(--white);color:#7a9a7e;flex-shrink:0}'
    + '.day-chip.sel{background:var(--pine);border-color:var(--pine);color:#fff}'
    + '.pat-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin:10px 0}'
    + '.pat-btn{padding:10px;border-radius:var(--r-sm);border:1.5px solid var(--border-md);background:var(--white);color:#7a9a7e;font-family:"DM Sans",sans-serif;font-size:11px;font-weight:500;cursor:pointer;text-align:center}'
    + '.pat-btn.sel{border-color:var(--sage);background:var(--dew);color:var(--pine);font-weight:600}'
    + '.prev-wrap{display:flex;flex-wrap:wrap;gap:5px;margin-top:10px}'
    + '.prev-chip{font-size:11px;padding:3px 9px;background:var(--dew);color:var(--pine);border:1px solid var(--mist);border-radius:20px}'
    + '.prev-more{font-size:11px;padding:3px 9px;background:var(--parchment);color:#7a9a7e;border:1px solid var(--border);border-radius:20px}'
    // submit / result
    + '.btn-sub{width:100%;padding:13px;border-radius:var(--r-sm);background:var(--forest);color:var(--mist);border:none;font-family:"DM Sans",sans-serif;font-size:14px;font-weight:600;cursor:pointer;margin-top:4px}'
    + '.btn-sub:disabled{opacity:.5;cursor:default}'
    + '.result{margin-top:10px;padding:10px 14px;border-radius:var(--r-sm);font-size:12px;font-weight:500;display:none}'
    + '.r-ok{background:#d1fae5;color:#065f46;border:1px solid #6ee7b7}'
    + '.r-err{background:var(--red-pale);color:var(--red);border:1px solid #fecaca}'
    // payments
    + '.pay-card{background:var(--white);border:1px solid var(--border);border-radius:var(--r);padding:14px;margin-bottom:9px;box-shadow:var(--shadow-xs)}'
    + '.pay-top{display:flex;justify-content:space-between;align-items:flex-start;gap:10px}'
    + '.pay-amt{font-family:"Playfair Display",serif;font-size:24px;font-weight:700;color:var(--pine);line-height:1}'
    + '.pay-meta{font-size:11px;color:#7a9a7e;margin-top:5px}'
    + '.pay-txn{font-size:10px;color:#b0c8b8;margin-top:3px}'
    + '.pay-cta{display:flex;justify-content:center;margin-top:16px}'
    + '.btn-pay{display:inline-flex;align-items:center;justify-content:center;gap:8px;background:var(--forest);color:var(--mist);padding:12px 24px;text-decoration:none;border:none;border-radius:var(--r-sm);font-size:13px;font-weight:600;font-family:"DM Sans",sans-serif;cursor:pointer}'
    + '.btn-pay:disabled{opacity:.55;cursor:not-allowed}'
    + '.pay-overlay{position:fixed;inset:0;z-index:200;background:rgba(20,40,15,.55);display:none;align-items:flex-end;justify-content:center;padding:0}'
    + '.pay-overlay.on{display:flex}'
    + '.pay-modal{background:#fff;width:100%;max-width:520px;max-height:92vh;overflow:auto;border-radius:18px 18px 0 0;padding:18px 16px 28px;box-shadow:var(--shadow-md);position:relative;z-index:1}'
    + '@media(min-width:640px){.pay-overlay{align-items:center;padding:18px}.pay-modal{border-radius:16px;max-height:90vh}}'
    + '.pay-modal h3{font-family:"Playfair Display",serif;font-size:20px;color:var(--ink);margin:0 0 4px}'
    + '.pay-modal .pay-sub{font-size:12px;color:var(--muted);margin-bottom:14px;line-height:1.45}'
    + '.pay-grid{display:grid;gap:10px}'
    + '.pay-grid .fl{display:block;font-size:11px;font-weight:600;color:var(--muted);margin-bottom:4px;text-transform:uppercase;letter-spacing:.04em}'
    + '.pay-grid .fi,.pay-grid select.fi,.pay-grid textarea.fi{width:100%;padding:11px 12px;border:1px solid var(--border);border-radius:10px;font:inherit;background:#fff;color:var(--ink)}'
    + '.pay-grid .fi[readonly]{background:#f3f6f2;color:#4b5563}'
    + '.pay-req{color:var(--red);font-weight:700}'
    + '.pay-shot{border:1px dashed var(--border-md);border-radius:12px;padding:12px;background:#f8faf6}'
    + '.pay-shot-preview{display:none;width:100%;max-height:180px;object-fit:contain;border-radius:8px;margin-top:8px;background:#fff}'
    + '.pay-qr-box{text-align:center;padding:12px;border:1px solid var(--border);border-radius:12px;background:#f8faf6;margin-bottom:4px}'
    + '.pay-qr-box img{width:180px;max-width:70%;display:block;margin:8px auto}'
    + '.pay-actions{display:flex;gap:8px;margin-top:14px}'
    + '.pay-actions .btn-pay{flex:1}'
    + '.pay-actions .btn-ghost{flex:0 0 auto;padding:12px 14px;border-radius:var(--r-sm);border:1px solid var(--border);background:#fff;font:inherit;font-weight:600;color:var(--muted);cursor:pointer}'
    + '.btn-ghost{padding:10px 12px;border-radius:var(--r-sm);border:1px solid var(--border);background:#fff;font:inherit;font-weight:600;color:var(--muted);cursor:pointer}'
    + '.pay-msg{font-size:12px;margin-top:10px;line-height:1.4}'
    + '.pay-msg.err{color:#991b1b}.pay-msg.ok{color:#166534}'
    // toast / misc
    + '#toast{position:fixed;bottom:20px;left:50%;transform:translateX(-50%) translateY(70px);background:var(--forest);color:var(--mist);font-size:12px;font-weight:600;padding:10px 20px;border-radius:100px;opacity:0;transition:all .25s;pointer-events:none;white-space:nowrap;z-index:9999;border:1px solid rgba(143,212,176,.25)}'
    + '#toast.show{opacity:1;transform:translateX(-50%) translateY(0)}'
    + '.empty-st{text-align:center;padding:30px 16px;color:#b0c8b8;font-size:13px;line-height:2}'
    + '.pfooter{text-align:center;padding:16px;font-size:10px;color:#b0c8b8;border-top:1px solid var(--border);letter-spacing:.04em}'
    + '.hint{font-size:10px;color:#b0c8b8;margin-top:4px}'
    // Change 7: pax-row hidden by default, shown for One-Time
    + '.pax-row{display:block}';

  var jsLines = [
    'var RD = null;',
    'var bookMode = "single";',
    'var slotCount = 1;',
    'var recurSelDays = [];',
    'var recurPattern = "";',
    'var SERVICES = ' + servicesJson + ';',
    'var PORTAL_CFG = ' + portalCfg + ';',
    'var DAYS = ["Su","Mo","Tu","We","Th","Fr","Sa"];',
    'var pendingPhone = "";',  // Change 6: remember phone for profile picker
    'var PAY_FORM = { files: [] };',
    'var PAY_QR_TIMER = null;',
    'var APP_UI_VERSION = ' + JSON.stringify(appUiVersion) + ';',
    'window.__APP_UI_VERSION__ = APP_UI_VERSION;',
    'window.__APP_BUILD__ = ' + JSON.stringify(buildStamp) + ';',
    'window.__APP_BUILD_LABEL__ = ' + JSON.stringify(buildLabel) + ';',
    '',
    'function forceFreshAppReload() {',
    '  try {',
    '    var live = window.__LIVE_APP_UI_VERSION__ || APP_UI_VERSION || String(Date.now());',
    '    sessionStorage.setItem("ke_reload_" + live, "1");',
    '    var u = new URL(location.href);',
    '    u.searchParams.set("_v", live);',
    '    u.searchParams.set("_cb", String(Date.now()));',
    '    location.replace(u.toString());',
    '  } catch (e) { location.reload(); }',
    '}',
    'function checkStaleAppUi() {',
    '  if (typeof google === "undefined" || !google.script || !google.script.run) return;',
    '  google.script.run.withSuccessHandler(function(res) {',
    '    var live = String((res && res.version) || "").trim();',
    '    var page = String(APP_UI_VERSION || "").trim();',
    '    window.__LIVE_APP_UI_VERSION__ = live;',
    '    if (!live || !page || live === page) return;',
    '    var tried = ""; try { tried = sessionStorage.getItem("ke_reload_" + live) || ""; } catch (e) {}',
    '    if (!tried) { try { sessionStorage.setItem("ke_reload_" + live, "1"); } catch (e2) {} forceFreshAppReload(); return; }',
    '    var ban = document.getElementById("stale-banner"); if (ban) ban.classList.add("on");',
    '  }).withFailureHandler(function(){}).getAppUiVersion();',
    '}',
    'checkStaleAppUi();',

    'function esc(v) {',
    '  return String(v || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");',
    '}',

    'function toast(msg) {',
    '  var t = document.getElementById("toast");',
    '  t.textContent = msg;',
    '  t.classList.add("show");',
    '  setTimeout(function() { t.classList.remove("show"); }, 2800);',
    '}',

    'function showResult(id, ok, msg) {',
    '  var el = document.getElementById(id);',
    '  if (!el) return;',
    '  el.textContent = msg;',
    '  el.className = "result " + (ok ? "r-ok" : "r-err");',
    '  el.style.display = "block";',
    '}',

    // ── login ──────────────────────────────────────────────
    'function doLogin() {',
    '  var id = document.getElementById("inp-id").value.trim();',
    '  if (!id) { showErr("Please enter your phone or KE Number."); return; }',
    '  var btn = document.getElementById("btn-login");',
    '  btn.textContent = "Looking up..."; btn.disabled = true;',
    '  document.getElementById("login-err").style.display = "none";',
    '  pendingPhone = id;',
    '  google.script.run',
    '    .withSuccessHandler(function(data) {',
    '      btn.textContent = "View My Rides"; btn.disabled = false;',
    '      if (!data || !data.found) { showErr(data ? data.error : "Not found."); return; }',
    // Change 6: handle multi-profile
    '      if (data.multiProfile) { showProfilePicker(data.profiles); return; }',
    '      RD = data;',
    '      try { localStorage.setItem("KE_ID", RD.keNo); } catch(e) {}',
    '      renderDash();',
    '    })',
    '    .withFailureHandler(function(e) {',
    '      btn.textContent = "View My Rides"; btn.disabled = false;',
    '      showErr("Something went wrong. Try again.");',
    '    })',
    '    .getRiderData(id);',
    '}',

    'function showErr(msg) {',
    '  var el = document.getElementById("login-err");',
    '  el.textContent = msg; el.style.display = "block";',
    '}',

    // Change 6: profile picker
    'function showProfilePicker(profiles) {',
    '  document.getElementById("login-screen").style.display = "none";',
    '  var pp = document.getElementById("profile-picker");',
    '  pp.style.display = "flex";',
    '  var list = profiles.map(function(p, i) {',
    '    var initials = p.name.split(" ").map(function(w){return w[0]||"";}).join("").toUpperCase().slice(0,2) || "KE";',
    '    return "<button class=\\"profile-btn\\" data-keno=\\"" + esc(p.keNo) + "\\">"',
    '      + "<div class=\\"profile-name\\">" + esc(initials) + " " + esc(p.name) + "</div>"',
    '      + "<div class=\\"profile-meta\\">" + esc(p.keNo) + (p.services ? " &nbsp;&middot;&nbsp; " + esc(p.services) : "") + "</div>"',
    '      + "</button>";',
    '  }).join("");',
    '  document.getElementById("profile-list").innerHTML = list;',
    '  document.querySelectorAll(".profile-btn").forEach(function(btn) {',
    '    btn.addEventListener("click", function() {',
    '      var keNo = btn.getAttribute("data-keno");',
    '      selectProfile(keNo);',
    '    });',
    '  });',
    '}',

    'function selectProfile(keNo) {',
    '  google.script.run',
    '    .withSuccessHandler(function(data) {',
    '      if (!data || !data.found) { alert("Profile not found. Try again."); return; }',
    '      RD = data;',
    '      try { localStorage.setItem("KE_ID", RD.keNo); } catch(e) {}',
    '      document.getElementById("profile-picker").style.display = "none";',
    '      renderDash();',
    '    })',
    '    .withFailureHandler(function(e) { alert("Error: " + e.message); })',
    '    .getRiderData(keNo);',
    '}',

    'function backToLogin() {',
    '  document.getElementById("profile-picker").style.display = "none";',
    '  document.getElementById("login-screen").style.display = "flex";',
    '}',

    'function doLogout() {',
    '  RD = null; slotCount = 1; bookMode = "single"; PAY_FORM.files = [];',
    '  try { closePortalPaymentForm(); } catch(e) {}',
    '  try { localStorage.removeItem("KE_ID"); } catch(e) {}',
    '  document.getElementById("dashboard").style.display = "none";',
    '  document.getElementById("profile-picker").style.display = "none";',
    '  document.getElementById("login-screen").style.display = "flex";',
    '  document.getElementById("inp-id").value = "";',
    '  document.getElementById("login-err").style.display = "none";',
    '  window.scrollTo(0, 0);',
    '}',

    // auto-restore by KE No (always unique)
    '(function() {',
    '  var saved = ""; try { saved = localStorage.getItem("KE_ID") || ""; } catch(e) {}',
    '  if (!saved) return;',
    '  document.getElementById("inp-id").value = saved;',
    '  var btn = document.getElementById("btn-login");',
    '  btn.textContent = "Restoring..."; btn.disabled = true;',
    '  google.script.run',
    '    .withSuccessHandler(function(data) {',
    '      btn.textContent = "View My Rides"; btn.disabled = false;',
    '      if (!data || !data.found || data.multiProfile) { try { localStorage.removeItem("KE_ID"); } catch(e) {} return; }',
    '      RD = data; renderDash();',
    '    })',
    '    .withFailureHandler(function() { btn.textContent = "View My Rides"; btn.disabled = false; })',
    '    .getRiderData(saved);',
    '})();',

    // ── dashboard ──────────────────────────────────────────
    'function renderDash() {',
    '  document.getElementById("login-screen").style.display = "none";',
    '  document.getElementById("dashboard").style.display = "block";',
    '  var initials = RD.name.split(" ").map(function(w) { return w[0] || ""; }).join("").toUpperCase().slice(0,2) || "KE";',
    '  document.getElementById("d-av").textContent = initials;',
    '  document.getElementById("d-name").textContent = RD.name;',
    '  document.getElementById("d-phone").textContent = RD.keNo + " \u00b7 " + RD.phone;',
    '  document.getElementById("d-keno").textContent = RD.name;',
    '  document.getElementById("d-svc").textContent = RD.services || "";',
    '  var upcoming = (RD.sessions || []).filter(function(s) { return s.isFuture && s.attendance !== "Present"; }).length;',
    '  document.getElementById("d-attended").textContent = RD.classesAttended || 0;',
    '  document.getElementById("d-noshow").textContent = RD.noShowCount || 0;',
    '  document.getElementById("d-upcoming").textContent = upcoming;',
    '  document.getElementById("d-participants").textContent = RD.totalParticipants || 0;',
    '  document.getElementById("d-payments").textContent = (RD.payments || []).length;',
    '  renderSessions(); renderBook(); renderPayments();',
    '  window.scrollTo(0, 0);',
    '}',

    // ── tab switch ─────────────────────────────────────────
    'function kTab(name) {',
    '  document.querySelectorAll(".tb").forEach(function(b) { b.classList.remove("on"); });',
    '  document.querySelectorAll(".tc").forEach(function(c) { c.classList.remove("on"); });',
    '  var btn = document.querySelector(".tb[data-tab=\\"" + name + "\\"]");',
    '  if (btn) btn.classList.add("on");',
    '  var tc = document.getElementById("tc-" + name);',
    '  if (tc) tc.classList.add("on");',
    '}',

    // ── sessions ───────────────────────────────────────────
    'function renderSessions() {',
    '  var sessions = RD.sessions || [];',
    '  var upcoming = sessions.filter(function(s) { return s.isFuture; });',
    '  var past     = sessions.filter(function(s) { return !s.isFuture; });',
    '  var h = "";',
    '  h += "<div class=\\"tab-heading\\">My Sessions</div>";',
    '  h += "<div class=\\"tab-sub\\">" + sessions.length + " total &middot; " + (RD.classesAttended || 0) + " class units attended (each 30 min = 1 unit)</div>";',
    '  if (upcoming.length) { h += "<div class=\\"sec-div\\">Upcoming</div>"; h += upcoming.map(buildSessCard).join(""); }',
    '  if (past.length)     { h += "<div class=\\"sec-div\\">Past</div>";     h += past.map(buildSessCard).join(""); }',
    '  if (!sessions.length) h += "<div class=\\"empty-st\\">No sessions yet. Use the Book tab!</div>";',
    '  document.getElementById("tc-sessions").innerHTML = h;',
    // Change 8: REMOVED min-date restriction on reschedule date pickers — backdating now allowed
    '  document.querySelectorAll(".rs-trigger").forEach(function(el) {',
    '    el.addEventListener("click", function() {',
    '      var rid = el.getAttribute("data-rid");',
    '      var panel = document.getElementById("rsp-" + rid);',
    '      if (panel) panel.classList.toggle("open");',
    '    });',
    '  });',
    '  document.querySelectorAll(".rs-submit").forEach(function(el) {',
    '    el.addEventListener("click", function() { submitResched(parseInt(el.getAttribute("data-idx"), 10)); });',
    '  });',
    '}',

    'function buildSessCard(s) {',
    '  var att = s.attendance || "", status = s.status || "";',
    '  var bCls = att === "Present" ? "b-dn" : att === "No-Show" ? "b-ns" : status === "Rescheduled" ? "b-rs" : "b-up";',
    '  var lbl  = att === "Present" ? "Attended" : att === "No-Show" ? "No-Show" : status || "Scheduled";',
    '  var cSt  = att === "Present" ? "s-att" : att === "No-Show" ? "s-ns" : status === "Rescheduled" ? "s-rs" : s.isFuture ? "s-up" : "";',
    '  var canR = s.isFuture && att !== "Present" && status !== "Completed" && status !== "Cancelled";',
    '  var rHtml = "";',
    '  if (canR) {',
    '    var rid = s.rowIndex;',
    '    rHtml += "<div class=\\"rs-trigger\\" data-rid=\\"" + rid + "\\">Reschedule</div>";',
    '    rHtml += "<div class=\\"rs-panel\\" id=\\"rsp-" + rid + "\\">";',
    '    rHtml += "<div class=\\"rs-from\\">Moving: <strong>" + esc(s.date) + " &middot; " + esc(s.timeSlot || "TBD") + "</strong></div>";',
    // Change 8: no min attribute on reschedule date input — backdating allowed
    '    rHtml += "<label class=\\"fl\\">New Date</label><input type=\\"date\\" class=\\"fi rs-date\\" id=\\"rsd-" + rid + "\\" style=\\"margin-bottom:10px\\">";',
    '    rHtml += "<label class=\\"fl\\">New Time</label><select class=\\"fi\\" id=\\"rst-" + rid + "\\" style=\\"margin-bottom:10px\\"><option value=\\"\\">Same (" + esc(s.timeSlot || "TBD") + ")</option>" + buildTimeOpts() + "</select>";',
    '    rHtml += "<label class=\\"fl\\">Reason</label><input type=\\"text\\" class=\\"fi\\" id=\\"rsr-" + rid + "\\" placeholder=\\"Optional...\\" style=\\"margin-bottom:10px\\">";',
    '    rHtml += "<button class=\\"btn-sub rs-submit\\" data-idx=\\"" + rid + "\\">Confirm Reschedule</button>";',
    '    rHtml += "<div id=\\"rsr-result-" + rid + "\\" class=\\"result\\"></div></div>";',
    '  }',
    '  return "<div class=\\"sc " + cSt + "\\">"',
    '    + "<div class=\\"si\\">"',
    '    + "<div class=\\"sc-top\\"><div>"',
    '    + "<div class=\\"sc-svc\\">" + esc(s.service) + "</div>"',
    '    + "<div class=\\"sc-dt\\">&#128197; " + esc(s.date || "TBD") + "</div></div>"',
    '    + "<span class=\\"bdg " + bCls + "\\">" + lbl + "</span>"',
    '    + "</div>"',
    '    + (s.timeSlot ? "<div class=\\"sc-tm\\">&#128336; " + esc(s.timeSlot) + "</div>" : "")',
    '    + "<div class=\\"sc-tm\\">&#128101; " + (Number(s.participants) || 1) + " participant" + ((Number(s.participants) || 1) !== 1 ? "s" : "") + "</div>"',
    '    + rHtml + "</div></div>";',
    '}',

    'function submitResched(rowIndex) {',
    '  var newDate = (document.getElementById("rsd-" + rowIndex) || {}).value || "";',
    '  var newTime = (document.getElementById("rst-" + rowIndex) || {}).value || "";',
    '  var reason  = (document.getElementById("rsr-" + rowIndex) || {}).value || "";',
    '  if (!newDate) { toast("Please pick a new date"); return; }',
    '  var btn = document.querySelector(".rs-submit[data-idx=\\"" + rowIndex + "\\"]");',
    '  if (btn) { btn.textContent = "Rescheduling..."; btn.disabled = true; }',
    '  google.script.run',
    '    .withSuccessHandler(function(res) {',
    '      if (btn) { btn.textContent = "Confirm Reschedule"; btn.disabled = false; }',
    '      showResult("rsr-result-" + rowIndex, res.success, res.message || res.error);',
    '      if (res.success) {',
    '        toast("Rescheduled!");',
    '        google.script.run.withSuccessHandler(function(d) { if (d && d.found && !d.multiProfile) { RD = d; renderSessions(); } }).getRiderData(RD.keNo);',
    '      }',
    '    })',
    '    .withFailureHandler(function(e) {',
    '      if (btn) { btn.textContent = "Confirm Reschedule"; btn.disabled = false; }',
    '      showResult("rsr-result-" + rowIndex, false, "Error: " + e.message);',
    '    })',
    '    .rescheduleSession(RD.keNo, rowIndex, newDate, newTime, reason);',
    '}',

    // ── book ───────────────────────────────────────────────
    'function renderBook() {',
    '  slotCount = 1; bookMode = "single"; recurSelDays = []; recurPattern = "";',
    '  var h = "<div class=\\"tab-heading\\">Book Sessions</div>";',
    '  h += "<div class=\\"tab-sub\\">Pick specific dates or a weekly pattern for a whole month.</div>";',
    '  h += "<div class=\\"mode-toggle\\">";',
    '  h += "<button class=\\"mode-btn active\\" id=\\"mbtn-single\\" data-mode=\\"single\\"><span class=\\"mi\\">&#128197;</span>Single / Multi<br><span style=\\"font-size:10px;font-weight:400;color:#7a9a7e\\">Specific dates</span></button>";',
    '  h += "<button class=\\"mode-btn\\" id=\\"mbtn-recurring\\" data-mode=\\"recurring\\"><span class=\\"mi\\">&#128260;</span>Recurring<br><span style=\\"font-size:10px;font-weight:400;color:#7a9a7e\\">Weekly pattern</span></button>";',
    '  h += "</div>";',
    '  h += "<div id=\\"area-single\\">" + buildSingleArea() + "</div>";',
    '  h += "<div id=\\"area-recur\\" style=\\"display:none\\">" + buildRecurArea() + "</div>";',
    '  h += "<button class=\\"btn-sub\\" id=\\"btn-book\\">Submit Sessions</button>";',
    '  h += "<div id=\\"book-result\\" class=\\"result\\"></div>";',
    '  document.getElementById("tc-book").innerHTML = h;',
    // Change 8: REMOVED min-date enforcement on single slot date pickers — backdating now allowed
    '  var now = new Date();',
    '  var mv = now.getFullYear() + "-" + String(now.getMonth() + 1).padStart(2, "0");',
    '  var mi = document.getElementById("recur-month"); if (mi) { mi.value = mv; mi.min = mv; }',
    '  document.querySelectorAll(".mode-btn").forEach(function(btn) {',
    '    btn.addEventListener("click", function() { switchMode(btn.getAttribute("data-mode")); });',
    '  });',
    '  var as = document.getElementById("btn-add-slot");',
    '  if (as) as.addEventListener("click", addSlot);',
    '  var bs = document.getElementById("btn-book");',
    '  if (bs) bs.addEventListener("click", submitBookings);',
    '  document.querySelectorAll(".pat-btn").forEach(function(btn) {',
    '    btn.addEventListener("click", function() { selectPattern(btn.getAttribute("data-pat")); });',
    '  });',
    '  document.querySelectorAll(".day-chip").forEach(function(chip) {',
    '    chip.addEventListener("click", function() { toggleDay(chip); });',
    '  });',
    '  wireSlotSvcChange(1);',
    '}',

    'function buildSingleArea() {',
    '  return "<div id=\\"slots-wrap\\">" + buildSlotCard(1) + "</div>"',
    '    + "<button class=\\"add-btn\\" id=\\"btn-add-slot\\">+ Add Another Session</button>";',
    '}',

    'function buildSlotCard(n) {',
    '  var opts = SERVICES.map(function(s, i) {',
    '    return "<option value=\\"" + i + "\\">" + esc(s.name) + (s.type ? " (" + esc(s.type) + ")" : "") + "</option>";',
    '  }).join("");',
    '  var h = "<div class=\\"slot-card\\" id=\\"slot-" + n + "\\">";',
    '  h += "<div class=\\"slot-num\\"><span class=\\"sn-badge\\">" + n + "</span> Session " + n + "</div>";',
    '  if (n > 1) h += "<button class=\\"rm-slot\\" data-slot=\\"" + n + "\\">Remove</button>";',
    '  h += "<label class=\\"fl\\">Service</label>";',
    '  h += "<select class=\\"fi\\" id=\\"svc-" + n + "\\" data-slot=\\"" + n + "\\" style=\\"margin-bottom:10px\\">";',
    '  h += "<option value=\\"\\">Select a service...</option>" + opts + "</select>";',
    '  h += "<label class=\\"fl\\">Date</label>";',
    // Change 8: no min attribute on date input — backdating allowed
    '  h += "<input type=\\"date\\" class=\\"fi sd-input\\" id=\\"sdate-" + n + "\\" style=\\"margin-bottom:10px\\">";',
    '  h += "<label class=\\"fl\\">Time Slot</label>";',
    '  h += "<select class=\\"fi\\" id=\\"stime-" + n + "\\" style=\\"margin-bottom:10px\\"><option value=\\"\\">Select time...</option>" + buildTimeOpts() + "</select>";',
    // Change 7: pax only for One-Time
    '  h += "<div class=\\"pax-row\\" id=\\"pax-" + n + "\\">";',
    '  h += "<label class=\\"fl\\">No. of Participants</label>";',
    '  h += "<input type=\\"number\\" class=\\"fi\\" id=\\"spax-" + n + "\\" value=\\"1\\" min=\\"1\\" max=\\"20\\" style=\\"margin-bottom:10px\\">";',
    '  h += "<p class=\\"hint\\">Include yourself in the count.</p></div>";',
    '  h += "</div>";',
    '  return h;',
    '}',

    // Change 7: show pax only for One-Time services
    'function wireSlotSvcChange(n) {',
    '  var sel = document.getElementById("svc-" + n);',
    '  if (!sel) return;',
    '  var updatePaxDisplay = function() {',
    '    var pr = document.getElementById("pax-" + n);',
    '    if (pr) pr.style.display = "block";',
    '  };',
    '  sel.addEventListener("change", updatePaxDisplay);',
    '  updatePaxDisplay();',
    '}',

    'function addSlot() {',
    '  slotCount++;',
    '  var wrap = document.getElementById("slots-wrap");',
    '  var div = document.createElement("div");',
    '  div.innerHTML = buildSlotCard(slotCount);',
    '  wrap.appendChild(div.firstChild);',
    // Change 8: REMOVED min-date setting when adding new slot — backdating allowed
    '  wireSlotSvcChange(slotCount);',
    '  var rb = document.querySelector("#slot-" + slotCount + " .rm-slot");',
    '  if (rb) rb.addEventListener("click", function() { document.getElementById("slot-" + slotCount).remove(); });',
    '}',

    'function buildRecurArea() {',
    '  var opts = SERVICES.map(function(s, i) {',
    '    return "<option value=\\"" + i + "\\">" + esc(s.name) + (s.type ? " (" + esc(s.type) + ")" : "") + "</option>";',
    '  }).join("");',
    '  var dayChips = DAYS.map(function(d, i) {',
    '    return "<div class=\\"day-chip\\" data-day=\\"" + i + "\\">" + d + "</div>";',
    '  }).join("");',
    '  var pats = [',
    '    {id:"thu", label:"All Thursdays"},',
    '    {id:"fri", label:"All Fridays"},',
    '    {id:"sat", label:"All Saturdays"},',
    '    {id:"sun", label:"All Sundays"},',
    '    {id:"wknd", label:"All Weekends"},',
    '    {id:"wkdy", label:"Weekdays (Mon-Fri)"},',
    '    {id:"cust", label:"Custom Days"}',
    '  ];',
    '  var patBtns = pats.map(function(p) {',
    '    return "<button class=\\"pat-btn\\" data-pat=\\"" + p.id + "\\">" + p.label + "</button>";',
    '  }).join("");',
    '  var h = "<div class=\\"rec-card\\">";',
    '  h += "<label class=\\"fl\\">Service</label><select class=\\"fi\\" id=\\"rsvc\\" style=\\"margin-bottom:10px\\"><option value=\\"\\">Select...</option>" + opts + "</select>";',
    '  h += "<label class=\\"fl\\">Month</label><input type=\\"month\\" class=\\"fi\\" id=\\"recur-month\\" style=\\"margin-bottom:10px\\">";',
    '  h += "<label class=\\"fl\\">Time Slot</label><select class=\\"fi\\" id=\\"rtime\\" style=\\"margin-bottom:12px\\"><option value=\\"\\">Select time...</option>" + buildTimeOpts() + "</select>";',
    // Change 7: pax hidden for recur too unless One-Time (wire on svc change)
    '  h += "<div class=\\"pax-row\\" id=\\"recur-pax\\"><label class=\\"fl\\">No. of Participants</label><input type=\\"number\\" class=\\"fi\\" id=\\"rpax\\" value=\\"1\\" min=\\"1\\" max=\\"20\\" style=\\"margin-bottom:10px\\"></div>";',
    '  h += "<label class=\\"fl\\" style=\\"margin-bottom:8px\\">Day Pattern</label>";',
    '  h += "<div class=\\"pat-grid\\">" + patBtns + "</div>";',
    '  h += "<div id=\\"cust-days\\" style=\\"display:none\\"><label class=\\"fl\\" style=\\"margin-top:8px\\">Select Days</label><div class=\\"day-picker\\">" + dayChips + "</div></div>";',
    '  h += "<div id=\\"recur-preview\\"></div>";',
    '  h += "</div>";',
    '  return h;',
    '}',

    'function switchMode(mode) {',
    '  bookMode = mode;',
    '  document.querySelectorAll(".mode-btn").forEach(function(b) { b.classList.remove("active"); });',
    '  var active = document.getElementById("mbtn-" + mode);',
    '  if (active) active.classList.add("active");',
    '  document.getElementById("area-single").style.display  = (mode === "single")    ? "block" : "none";',
    '  document.getElementById("area-recur").style.display   = (mode === "recurring") ? "block" : "none";',
    '}',

    'function toggleDay(chip) {',
    '  chip.classList.toggle("sel");',
    '  recurSelDays = [];',
    '  document.querySelectorAll(".day-chip.sel").forEach(function(c) {',
    '    recurSelDays.push(parseInt(c.getAttribute("data-day"), 10));',
    '  });',
    '  recurPattern = "cust";',
    '  updateRecurPreview();',
    '}',

    'function selectPattern(pat) {',
    '  document.querySelectorAll(".pat-btn").forEach(function(b) { b.classList.remove("sel"); });',
    '  var el = document.querySelector(".pat-btn[data-pat=\\"" + pat + "\\"]");',
    '  if (el) el.classList.add("sel");',
    '  recurPattern = pat;',
    '  var cd = document.getElementById("cust-days");',
    '  if (pat === "cust") {',
    '    if (cd) cd.style.display = "block";',
    '  } else {',
    '    if (cd) cd.style.display = "none";',
    '    document.querySelectorAll(".day-chip").forEach(function(c) { c.classList.remove("sel"); });',
    '    recurSelDays = [];',
    '  }',
    '  updateRecurPreview();',
    '}',

    'function getRecurDates() {',
    '  var mi = document.getElementById("recur-month");',
    '  if (!mi || !mi.value) return [];',
    '  var parts = mi.value.split("-");',
    '  var year  = parseInt(parts[0], 10);',
    '  var month = parseInt(parts[1], 10) - 1;',
    '  var tgt = [];',
    '  if      (recurPattern === "thu")  tgt = [4];',
    '  else if (recurPattern === "fri")  tgt = [5];',
    '  else if (recurPattern === "sat")  tgt = [6];',
    '  else if (recurPattern === "sun")  tgt = [0];',
    '  else if (recurPattern === "wknd") tgt = [0, 6];',
    '  else if (recurPattern === "wkdy") tgt = [1, 2, 3, 4, 5];',
    '  else if (recurPattern === "cust") tgt = recurSelDays.slice();',
    '  else return [];',
    // Change 8: REMOVED future-only filter — all dates in month included (past and future)
    '  var dates = [], d = new Date(year, month, 1);',
    '  while (d.getMonth() === month) {',
    '    if (tgt.indexOf(d.getDay()) > -1) { dates.push(new Date(d)); }',
    '    d.setDate(d.getDate() + 1);',
    '  }',
    '  return dates;',
    '}',

    'function updateRecurPreview() {',
    '  var el = document.getElementById("recur-preview");',
    '  if (!el) return;',
    '  var dates = getRecurDates();',
    '  if (!dates.length) { el.innerHTML = ""; return; }',
    '  var shown = dates.slice(0, 5), extra = dates.length - shown.length;',
    '  var chips = shown.map(function(dt) {',
    '    return "<span class=\\"prev-chip\\">" + dt.toLocaleDateString("en-IN", {weekday:"short",day:"numeric",month:"short"}) + "</span>";',
    '  }).join("");',
    '  if (extra > 0) chips += "<span class=\\"prev-more\\">+" + extra + " more</span>";',
    '  el.innerHTML = "<div style=\\"margin-top:10px\\"><div style=\\"font-size:10px;font-weight:600;color:var(--sage);text-transform:uppercase;letter-spacing:.08em;margin-bottom:6px\\">" + dates.length + " session" + (dates.length !== 1 ? "s" : "") + " will be booked</div><div class=\\"prev-wrap\\">" + chips + "</div></div>";',
    '}',

    'function submitBookings() {',
    '  var requests = [];',
    '  if (bookMode === "single") {',
    '    for (var i = 1; i <= slotCount; i++) {',
    '      if (!document.getElementById("slot-" + i)) continue;',
    '      var sv = document.getElementById("svc-"   + i);',
    '      var dt = document.getElementById("sdate-" + i);',
    '      var tm = document.getElementById("stime-" + i);',
    '      var px = document.getElementById("spax-"  + i);',
    '      var sIdx = sv ? sv.value : "";',
    '      var date = dt ? dt.value : "";',
    '      if (!sIdx || !date) { toast("Fill service and date for all sessions"); return; }',
    '      var svc = SERVICES[parseInt(sIdx, 10)];',
    // Change 7: pax only for One-Time
    '      var pax = (svc && svc.showPax && px) ? parseInt(px.value || 1, 10) : 1;',
    '      requests.push({ service: svc.name, date: date, timeSlot: tm ? tm.value : "", participants: pax });',
    '    }',
    '  } else {',
    '    var rsi = document.getElementById("rsvc");',
    '    var rti = document.getElementById("rtime");',
    '    var rpi = document.getElementById("rpax");',
    '    var rSvcIdx = rsi ? rsi.value : "";',
    '    if (!rSvcIdx) { toast("Please select a service"); return; }',
    '    if (!recurPattern) { toast("Please select a day pattern"); return; }',
    '    var rdates = getRecurDates();',
    '    if (!rdates.length) { toast("No dates in selected month"); return; }',
    '    var rsvc = SERVICES[parseInt(rSvcIdx, 10)];',
    '    var rpax = (rsvc && rsvc.showPax && rpi) ? parseInt(rpi.value || 1, 10) : 1;',
    '    rdates.forEach(function(d) {',
    '      requests.push({ service: rsvc.name, date: d.toISOString().split("T")[0], timeSlot: rti ? rti.value : "", participants: rpax });',
    '    });',
    '  }',
    '  if (!requests.length) { toast("No sessions to book"); return; }',
    '  var btn = document.getElementById("btn-book");',
    '  if (btn) { btn.textContent = "Booking " + requests.length + " session" + (requests.length !== 1 ? "s" : "") + "..."; btn.disabled = true; }',
    '  google.script.run',
    '    .withSuccessHandler(function(res) {',
    '      if (btn) { btn.textContent = "Submit Sessions"; btn.disabled = false; }',
    '      showResult("book-result", res.success, res.message || res.error || "Done");',
    '      if (res.success) {',
    '        toast("Booked " + res.added + " session" + (res.added !== 1 ? "s" : "") + "!");',
    '        google.script.run.withSuccessHandler(function(d) { if (d && d.found && !d.multiProfile) { RD = d; renderDash(); } }).getRiderData(RD.keNo);',
    '      }',
    '    })',
    '    .withFailureHandler(function(e) {',
    '      if (btn) { btn.textContent = "Submit Sessions"; btn.disabled = false; }',
    '      showResult("book-result", false, "Error: " + e.message);',
    '    })',
    '    .bookMultipleSessions(RD.keNo, requests);',
    '}',

    // ── payments ───────────────────────────────────────────
    'function renderPayments() {',
    '  var payments = RD.payments || [];',
    '  var h = "<div class=\\"tab-heading\\">My Payments</div>";',
    '  h += "<div class=\\"tab-sub\\">" + payments.length + " transaction" + (payments.length !== 1 ? "s" : "") + " recorded</div>";',
    '  if (!payments.length) {',
    '    h += "<div class=\\"empty-st\\">No payments yet. Payments appear after your receipt is generated.</div>";',
    '  } else {',
    '    h += payments.map(function(p) {',
    '      var r = "<div class=\\"pay-card\\"><div class=\\"pay-top\\">";',
    '      r += "<div><div class=\\"pay-amt\\">&#8377;" + Number(p.amount).toLocaleString("en-IN") + "</div>";',
    '      if (p.payDate) r += "<div class=\\"pay-meta\\">&#128197; " + esc(p.payDate) + "</div>";',
    '      r += "</div><div style=\\"text-align:right\\">";',
    '      if (p.receiptNo) r += "<div class=\\"pay-meta\\">Receipt: " + esc(p.receiptNo) + "</div>";',
    '      if (p.paidOn) r += "<div style=\\"font-size:10px;color:#b0c8b8;margin-top:4px\\">" + esc(p.paidOn) + "</div>";',
    '      r += "</div></div>";',
    '      if (p.txnRef) r += "<div class=\\"pay-txn\\">Txn: " + esc(p.txnRef) + "</div>";',
    '      r += "</div>";',
    '      return r;',
    '    }).join("");',
    '  }',
    '  h += "<div class=\\"pay-cta\\"><button type=\\"button\\" class=\\"btn-pay\\" id=\\"btn-make-payment\\">+ Make a Payment</button></div>";',
    '  document.getElementById("tc-payments").innerHTML = h;',
    '  var makePay = document.getElementById("btn-make-payment");',
    '  if (makePay) makePay.addEventListener("click", openPortalPaymentForm);',
    '}',

    'function payUpiLink(amount) {',
    '  var amt = Number(amount);',
    '  var link = "upi://pay?pa=" + encodeURIComponent(PORTAL_CFG.upiId || "")',
    '    + "&pn=" + encodeURIComponent(PORTAL_CFG.businessName || "KingsEquestrian")',
    '    + "&cu=INR"',
    '    + "&tn=" + encodeURIComponent((RD && RD.keNo) || "KE");',
    '  if (amt > 0) link += "&am=" + amt;',
    '  return link;',
    '}',

    'function openPortalPaymentForm() {',
    '  if (!RD) { toast("Please sign in first"); return; }',
    '  PAY_FORM.files = [];',
    '  document.getElementById("pay-ke").value = RD.keNo || "";',
    '  document.getElementById("pay-phone").value = RD.phone || "";',
    '  document.getElementById("pay-amount").value = "";'
    '  document.getElementById("pay-txn").value = "";',
    '  document.getElementById("pay-pan").value = "";',
    '  var today = new Date();',
    '  var ymd = today.getFullYear() + "-" + String(today.getMonth()+1).padStart(2,"0") + "-" + String(today.getDate()).padStart(2,"0");',
    '  document.getElementById("pay-date").value = ymd;',
    '  document.getElementById("pay-shot").value = "";',
    '  document.getElementById("pay-shot-name").textContent = "No file chosen";',
    '  var prev = document.getElementById("pay-shot-preview"); if (prev) { prev.style.display="none"; prev.removeAttribute("src"); }',
    '  var msg = document.getElementById("pay-msg"); msg.className = "pay-msg"; msg.textContent = "";',
    '  document.getElementById("pay-sub").textContent = "Please fill the payment details for verification for " + (RD.name || RD.keNo) + ".";',
    '  document.getElementById("pay-overlay").classList.add("on");',
    '  google.script.run.withSuccessHandler(function(res){',
    '    if (!res || !res.success) return;',
    '    if (res.phone && !document.getElementById("pay-phone").value) document.getElementById("pay-phone").value = res.phone;',
    '    if (res.pan) document.getElementById("pay-pan").value = res.pan;',
    '  }).withFailureHandler(function(){}).getPortalPaymentPrefill(RD.keNo);',
    '  refreshPortalPaymentQr();',
    '}',

    'function closePortalPaymentForm() {',
    '  document.getElementById("pay-overlay").classList.remove("on");',
    '}',

    'function refreshPortalPaymentQr() {',
    '  if (PAY_QR_TIMER) clearTimeout(PAY_QR_TIMER);',
    '  PAY_QR_TIMER = setTimeout(function(){',
    '    var wrap = document.getElementById("pay-qr-wrap");',
    '    var amt = Number(document.getElementById("pay-amount").value || 0);',
    '    if (!(amt > 0)) { wrap.style.display = "none"; return; }',
    '    wrap.style.display = "block";',
    '    document.getElementById("pay-upi").textContent = PORTAL_CFG.upiId || "";',
    '    var img = document.getElementById("pay-qr");',
    '    if (img) img.src = "https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=" + encodeURIComponent(payUpiLink(amt));',
    '  }, 350);',
    '}',

    'function onPayScreenshotChosen(input) {',
    '  var files = Array.prototype.slice.call((input && input.files) || [], 0, 5);',
    '  PAY_FORM.files = [];',
    '  var nameEl = document.getElementById("pay-shot-name");',
    '  var prev = document.getElementById("pay-shot-preview");',
    '  if (!files.length) { if (nameEl) nameEl.textContent = "No file chosen"; if (prev) prev.style.display="none"; return; }',
    '  if (nameEl) nameEl.textContent = files.map(function(f){ return f.name; }).join(", ");',
    '  var i = 0;',
    '  function next() {',
    '    if (i >= files.length) {',
    '      var firstImg = PAY_FORM.files.filter(function(f){ return String(f.mimeType||"").indexOf("image/")===0; })[0];',
    '      if (prev) {',
    '        if (firstImg) { prev.src = "data:" + firstImg.mimeType + ";base64," + firstImg.data; prev.style.display = "block"; }',
    '        else { prev.style.display = "none"; }',
    '      }',
    '      return;',
    '    }',
    '    var file = files[i++];',
    '    var isImg = file.type && file.type.indexOf("image/") === 0;',
    '    var isPdf = (file.type === "application/pdf") || /\\.pdf$/i.test(file.name);',
    '    if (!isImg && !isPdf) { toast("Please choose an image or PDF"); input.value=""; PAY_FORM.files=[]; if (nameEl) nameEl.textContent="No file chosen"; return; }',
    '    var reader = new FileReader();',
    '    reader.onload = function() {',
    '      if (isPdf) {',
    '        var dataUrl = String(reader.result || "");',
    '        var comma = dataUrl.indexOf(",");',
    '        PAY_FORM.files.push({ name: file.name, mimeType: "application/pdf", data: comma >= 0 ? dataUrl.slice(comma + 1) : dataUrl });',
    '        next();',
    '        return;',
    '      }',
    '      var img = new Image();',
    '      img.onload = function() {',
    '        var max = 1280, w = img.width, h = img.height;',
    '        if (w > max || h > max) { var s = Math.min(max/w, max/h); w = Math.round(w*s); h = Math.round(h*s); }',
    '        var canvas = document.createElement("canvas"); canvas.width = w; canvas.height = h;',
    '        canvas.getContext("2d").drawImage(img, 0, 0, w, h);',
    '        var dataUrl = canvas.toDataURL("image/jpeg", 0.82);',
    '        PAY_FORM.files.push({ name: file.name, mimeType: "image/jpeg", data: dataUrl.split(",")[1] });',
    '        next();',
    '      };',
    '      img.onerror = function(){ toast("Could not read image"); };',
    '      img.src = reader.result;',
    '    };',
    '    reader.onerror = function(){ toast("Could not read file"); };',
    '    reader.readAsDataURL(file);',
    '  }',
    '  next();',
    '}',

    'function submitPortalPaymentForm() {',
    '  if (!RD) return;',
    '  var msg = document.getElementById("pay-msg");',
    '  var btn = document.getElementById("pay-submit");',
    '  var keNo = document.getElementById("pay-ke").value.trim();',
    '  var phone = document.getElementById("pay-phone").value.trim();',
    '  var amount = Number(document.getElementById("pay-amount").value || 0);',
    '  var payDate = document.getElementById("pay-date").value || "";',
    '  var txnRef = document.getElementById("pay-txn").value.trim();',
    '  var pan = document.getElementById("pay-pan").value.trim();',
    '  if (!keNo) { msg.className="pay-msg err"; msg.textContent="Please enter your Registration Number provided in your email."; return; }',
    '  if (!phone) { msg.className="pay-msg err"; msg.textContent="Please enter your registered phone number."; return; }',
    '  if (!(amount > 0)) { msg.className="pay-msg err"; msg.textContent="Please enter the exact amount paid."; return; }',
    '  if (!PAY_FORM.files.length) { msg.className="pay-msg err"; msg.textContent="Please upload a clear screenshot or receipt of the completed payment."; return; }',
    '  if (!pan) { msg.className="pay-msg err"; msg.textContent="PAN / Aadhaar number is required for issuing official receipts."; return; }',
    '  if (btn) { btn.disabled = true; btn.textContent = "Submitting..."; }',
    '  msg.className = "pay-msg"; msg.textContent = "Uploading payment details...";',
    '  google.script.run',
    '    .withSuccessHandler(function(res) {',
    '      if (btn) { btn.disabled = false; btn.textContent = "Submit payment"; }',
    '      if (!res || !res.success) { msg.className="pay-msg err"; msg.textContent = (res && (res.error || res.message)) || "Payment failed."; return; }',
    '      msg.className = "pay-msg ok"; msg.textContent = res.message || "Payment submitted.";',
    '      toast("Payment submitted");',
    '      google.script.run.withSuccessHandler(function(d){',
    '        if (d && d.found && !d.multiProfile) { RD = d; }',
    '        closePortalPaymentForm();',
    '        renderPayments();',
    '        if (RD) document.getElementById("d-payments").textContent = (RD.payments || []).length;',
    '      }).getRiderData(RD.keNo);',
    '    })',
    '    .withFailureHandler(function(e) {',
    '      if (btn) { btn.disabled = false; btn.textContent = "Submit payment"; }',
    '      msg.className = "pay-msg err"; msg.textContent = "Error: " + (e && e.message ? e.message : "Please try again.");',
    '    })',
    '    .submitPaymentFromPortal({ keNo: keNo, phone: phone, amount: amount, payDate: payDate, txnRef: txnRef, pan: pan, files: PAY_FORM.files });',
    '}',

    // ── time options — Change 1: 30-min slots ───────────────
    'function buildTimeOpts() {',
    '  function p(n) { return (n < 10 ? "0" : "") + n; }',
    '  function grp(label, sh, sm, eh) {',
    '    var html = "<optgroup label=\\"" + label + "\\">";',
    '    var cur = sh * 60 + sm, end = eh * 60;',
    '    while (cur + 30 <= end) {',
    '      var s1 = Math.floor(cur / 60), m1 = cur % 60;',
    '      var s2 = Math.floor((cur + 30) / 60), m2 = (cur + 30) % 60;',
    '      html += "<option>" + p(s1) + ":" + p(m1) + " - " + p(s2) + ":" + p(m2) + "</option>";',
    '      cur += 30;',   // 30-min increments
    '    }',
    '    return html + "</optgroup>";',
    '  }',
    '  return grp("Morning", 6, 0, 12) + grp("Afternoon", 13, 0, 19);',
    '}'
  ];

  var js = jsLines.join('\n');

  var html = '<!DOCTYPE html>'
    + '<html lang="en"><head>'
    + '<meta charset="UTF-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">'
    + '<meta name="apple-mobile-web-app-capable" content="yes">'
    + '<meta name="theme-color" content="#0a1f16">'
    + '<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">'
    + '<meta http-equiv="Pragma" content="no-cache">'
    + '<meta http-equiv="Expires" content="0">'
    + '<script>window.__APP_BUILD__=' + JSON.stringify(buildStamp) + ';window.__APP_BUILD_LABEL__=' + JSON.stringify(buildLabel) + ';window.__APP_UI_VERSION__=' + JSON.stringify(appUiVersion) + ';</script>'
    + '<link rel="preconnect" href="https://fonts.googleapis.com">'
    + '<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@500;600;700&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">'
    + '<title>My Rides &middot; Kings Equestrian</title>'
    + '<style>' + css + '</style>'
    + '</head><body>'

    // LOGIN
    + '<div id="login-screen">'
    +   '<div class="login-ring" style="width:320px;height:320px;top:-80px;right:-80px"></div>'
    +   '<div class="login-ring" style="width:200px;height:200px;bottom:60px;left:-60px"></div>'
    +   '<div class="login-card">'
    +     '<div class="brand-block">'
    +       '<div class="brand-icon">&#128052;</div>'
    +       '<div class="brand-title">My Rides</div>'
    +       '<div class="brand-sub">Kings Equestrian</div>'
    +     '</div>'
    +     '<div class="fg">'
    +       '<label class="fl" for="inp-id">Phone or KE Number</label>'
    +       '<input type="tel" class="fi" id="inp-id" placeholder="e.g. 9876543210 or KE240101..." maxlength="20">'
    +     '</div>'
    +     '<button class="btn-p" id="btn-login">View My Rides</button>'
    +     '<div class="login-err" id="login-err"></div>'
    +     '<p class="login-note">Enter your registered phone or KE Number.<br>No password needed.</p>'
    +     '<div class="build-lbl">build ' + buildLabel + (appUiVersion ? ' · v' + appUiVersion : '') + '</div>'
    +   '</div>'
    + '</div>'

    // Change 6: PROFILE PICKER SCREEN
    + '<div id="profile-picker" style="display:none;min-height:100vh;background:var(--forest);align-items:center;justify-content:center;padding:2rem 1.25rem;flex-direction:column">'
    +   '<div class="picker-card">'
    +     '<div class="brand-icon" style="margin:0 auto 1rem">&#128101;</div>'
    +     '<div class="picker-title">Choose a Profile</div>'
    +     '<div class="picker-sub">Multiple profiles found for this number.<br>Select a family member to view their bookings.</div>'
    +     '<div id="profile-list"></div>'
    +     '<button class="btn-back" id="btn-back-login">&#8592; Use a different number</button>'
    +   '</div>'
    + '</div>'

    // DASHBOARD
    + '<div id="dashboard">'
    +   '<div class="dh">'
    +     '<div class="av" id="d-av">KE</div>'
    +     '<div class="dh-info">'
    +       '<div class="dh-name" id="d-name"></div>'
    +       '<div class="dh-sub" id="d-phone"></div>'
    +     '</div>'
    +     '<button class="btn-lo" id="btn-logout">Sign out</button>'
    +   '</div>'
    +   '<div id="stale-banner" class="stale-banner">A newer version of My Rides is available. <button type="button" id="stale-reload-btn">Load update</button></div>'
    +   '<div class="info-banner">'
    +     '<div class="info-ke" id="d-keno"></div>'
    +     '<div class="info-svc" id="d-svc"></div>'
    +     '<div class="info-stats">'
    +       '<div class="istat"><div class="istat-val" id="d-attended">0</div><div class="istat-lbl">Class Units</div></div>'
    +       '<div class="istat"><div class="istat-val" id="d-noshow">0</div><div class="istat-lbl">No-Shows</div></div>'
    +       '<div class="istat"><div class="istat-val" id="d-upcoming">0</div><div class="istat-lbl">Upcoming</div></div>'
    +       '<div class="istat"><div class="istat-val" id="d-participants">0</div><div class="istat-lbl">Total Participants</div></div>'
    +       '<div class="istat"><div class="istat-val" id="d-payments">0</div><div class="istat-lbl">Payments</div></div>'
    +     '</div>'
    +   '</div>'
    +   '<div class="tabs">'
    +     '<button class="tb on" data-tab="sessions">&#128197; Sessions</button>'
    +     '<button class="tb" data-tab="book">+ Book</button>'
    +     '<button class="tb" data-tab="payments">&#128179; Payments</button>'
    +   '</div>'
    +   '<div id="tc-sessions" class="tc on"></div>'
    +   '<div id="tc-book" class="tc"></div>'
    +   '<div id="tc-payments" class="tc"></div>'
    +   '<div class="pfooter">Kings Equestrian Foundation &middot; Karnataka &middot; +91-9980895533<br>build ' + buildLabel + (appUiVersion ? ' · v' + appUiVersion : '') + '</div>'
    + '</div>'

    + '<div id="toast"></div>'

    + '<div id="pay-overlay" class="pay-overlay" onclick="if(event.target===this)closePortalPaymentForm()">'
    +   '<div class="pay-modal" role="dialog" aria-modal="true">'
    +     '<h3>Make a payment</h3>'
    +     '<div class="pay-sub" id="pay-sub">Please fill the payment details for verification. Receipt will be emailed after verification.</div>'
    +     '<div id="pay-qr-wrap" class="pay-qr-box" style="display:none">'
    +       '<div style="font-size:12px;color:#476d59">Scan UPI QR (optional)</div>'
    +       '<img id="pay-qr" alt="UPI QR">'
    +       '<div style="font-size:12px;font-weight:600;color:var(--pine)">UPI: <span id="pay-upi"></span></div>'
    +     '</div>'
    +     '<div class="pay-grid">'
    +       '<div><label class="fl" for="pay-ke">Registration No</label><input class="fi" id="pay-ke" readonly><div class="hint">Please Enter your Registration Number provided in your email</div></div>'
    +       '<div><label class="fl" for="pay-phone">Phone number <span class="pay-req">*</span></label><input class="fi" id="pay-phone" inputmode="tel"><div class="hint">Please enter your registered Phone Number</div></div>'
    +       '<div><label class="fl" for="pay-amount">Amount Paid (₹) <span class="pay-req">*</span></label><input class="fi" id="pay-amount" type="number" min="1" step="1" inputmode="decimal"><div class="hint">Please Enter the exact Amount Paid</div></div>'
    +       '<div class="pay-shot">'
    +         '<label class="fl" for="pay-shot">Screenshot <span class="pay-req">*</span></label>'
    +         '<input type="file" id="pay-shot" accept="image/*,.pdf,application/pdf" multiple style="display:none">'
    +         '<button type="button" class="btn-ghost" id="pay-shot-btn" style="width:100%">Choose screenshot</button>'
    +         '<div class="hint" id="pay-shot-name">No file chosen</div>'
    +         '<img class="pay-shot-preview" id="pay-shot-preview" alt="Screenshot preview">'
    +         '<div class="hint">Upload a clear screenshot or receipt of the completed payment. Upload up to 5 supported files.</div>'
    +       '</div>'
    +       '<div><label class="fl" for="pay-date">Payment Date</label><input class="fi" id="pay-date" type="date"><div class="hint">Select the date when the payment was made</div></div>'
    +       '<div><label class="fl" for="pay-txn">Transaction Reference Number</label><input class="fi" id="pay-txn" maxlength="80" placeholder="UPI / bank reference"><div class="hint">If payment was made via UPI (PhonePe, Google Pay, Paytm, etc.), please enter the transaction ID / reference number.</div></div>'
    +       '<div><label class="fl" for="pay-pan">PAN / Aadhaar Number <span class="pay-req">*</span></label><input class="fi" id="pay-pan" maxlength="20" placeholder="PAN or Aadhaar"><div class="hint">Details are required for issuing official payment receipts and for compliance with income tax regulations. This information will be kept confidential and will not be shared with any third party.</div></div>'
    +     '</div>'
    +     '<div id="pay-msg" class="pay-msg"></div>'
    +     '<div class="pay-actions">'
    +       '<button type="button" class="btn-ghost" id="pay-cancel">Back</button>'
    +       '<button type="button" class="btn-pay" id="pay-submit">Submit payment</button>'
    +     '</div>'
    +   '</div>'
    + '</div>'

    + '<script>' + js + '</script>'

    + '<script>'
    + 'document.getElementById("inp-id").addEventListener("keydown", function(e) { if (e.key === "Enter") doLogin(); });'
    + 'document.getElementById("btn-login").addEventListener("click", doLogin);'
    + 'document.getElementById("btn-logout").addEventListener("click", doLogout);'
    + 'document.getElementById("btn-back-login").addEventListener("click", backToLogin);'
    + 'document.querySelectorAll(".tb").forEach(function(btn) {'
    +   'btn.addEventListener("click", function() { kTab(btn.getAttribute("data-tab")); });'
    + '});'
    + 'document.getElementById("pay-cancel").addEventListener("click", closePortalPaymentForm);'
    + 'document.getElementById("pay-submit").addEventListener("click", submitPortalPaymentForm);'
    + 'document.getElementById("pay-amount").addEventListener("input", refreshPortalPaymentQr);'
    + 'document.getElementById("pay-shot-btn").addEventListener("click", function(){ document.getElementById("pay-shot").click(); });'
    + 'document.getElementById("pay-shot").addEventListener("change", function(){ onPayScreenshotChosen(this); });'
    + 'var staleBtn=document.getElementById("stale-reload-btn"); if(staleBtn) staleBtn.addEventListener("click", forceFreshAppReload);'
    + '</script>'

    + '</body></html>';

  return html;
}