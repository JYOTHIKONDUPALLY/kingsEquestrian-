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
  var paymentLink  = CONFIG.PAYMENT_FORM_BASE_URL || '#';
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
  var appUiVersion = String((CONFIG && CONFIG.APP_UI_VERSION) || '');

  return _portalHTML(paymentLink, servicesJson, appUiVersion);
}

function _portalHTML(payLink, servicesJson, appUiVersion) {
  var logoUrl = (typeof CONFIG !== 'undefined' && CONFIG.LOGO_URL) ? CONFIG.LOGO_URL : '';
  appUiVersion = String(appUiVersion || '');

  var css = ''
    + '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}'
    + ':root{'
    // Brand palette from the Kings Equestrian logo: green horseshoe (primary),
    // maroon horse (accent), gold crown (highlight). Shared with the attendance app.
    + '--ink:#152410;--forest:#14330f;--pine:#1f4617;--sage:#2e7d46;'
    + '--fern:#4aa06e;--mint:#9fd3b8;--mist:#cbe8d6;--dew:#eaf3ec;'
    + '--parchment:#f4f7f2;--white:#ffffff;'
    + '--maroon:#6b1a2a;--maroon-soft:#fbeef0;'
    + '--gold:#a9781a;--gold-pale:#fbf3db;--gold-border:#e9d59a;'
    + '--muted:#6f7b6a;'
    + '--red:#b91c1c;--red-pale:#fef2f2;'
    + '--border:rgba(31,70,23,0.14);--border-md:rgba(31,70,23,0.24);'
    + '--shadow-xs:0 1px 3px rgba(20,40,15,.07);--shadow-sm:0 3px 12px rgba(20,40,15,.11);--shadow-md:0 12px 34px rgba(20,40,15,.16);'
    + '--r:14px;--r-sm:10px}'
    + 'html,body{min-height:100%;-webkit-font-smoothing:antialiased}'
    + 'body{font-family:"Instrument Sans",sans-serif;background:var(--parchment);color:var(--ink)}'
    // login
    + '#login-screen{min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:2rem 1.25rem 4rem;background:linear-gradient(160deg,var(--forest),var(--pine) 55%,#42101a);position:relative;overflow:hidden}'
    + '.login-ring{position:absolute;border-radius:50%;border:1px solid rgba(201,162,39,.15);pointer-events:none}'
    + '.login-card{width:100%;max-width:380px;background:var(--white);border-radius:20px;padding:2rem 1.75rem 1.75rem;box-shadow:var(--shadow-md);position:relative;z-index:1}'
    + '.brand-block{text-align:center;margin-bottom:1.75rem}'
    + '.brand-icon{width:72px;height:72px;border-radius:16px;background:var(--white);display:flex;align-items:center;justify-content:center;margin:0 auto 1rem;border:2px solid var(--gold-border);padding:6px;box-shadow:var(--shadow-xs)}'
    + '.brand-icon img{width:100%;height:100%;object-fit:contain;border-radius:10px}'
    + '.brand-title{font-family:"Syne",sans-serif;font-size:26px;font-weight:700;color:var(--ink);line-height:1.1;margin-bottom:4px}'
    + '.brand-sub{font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);font-weight:500}'
    + '.fg{margin-bottom:14px}'
    + '.fl{display:block;font-size:11px;font-weight:600;letter-spacing:.07em;text-transform:uppercase;color:var(--sage);margin-bottom:6px}'
    + '.fi{width:100%;padding:11px 14px;border:1.5px solid var(--border-md);border-radius:var(--r-sm);font-family:"Instrument Sans",sans-serif;font-size:14px;color:var(--ink);background:var(--parchment);outline:none;-webkit-appearance:none;transition:all .15s}'
    + '.fi:focus{border-color:var(--sage);box-shadow:0 0 0 3px rgba(46,125,70,.12);background:var(--white)}'
    + '.fi::placeholder{color:#a8b8a8}'
    + 'select.fi{cursor:pointer}'
    + '.btn-p{width:100%;padding:13px;border-radius:var(--r-sm);background:linear-gradient(120deg,var(--forest),var(--pine));color:#fff;border:none;font-family:"Instrument Sans",sans-serif;font-size:14px;font-weight:600;cursor:pointer;margin-top:4px;transition:all .15s;box-shadow:var(--shadow-xs)}'
    + '.btn-p:disabled{opacity:.5;cursor:default}'
    + '.login-err{font-size:12px;color:var(--red);text-align:center;margin-top:10px;padding:9px 12px;background:var(--red-pale);border-radius:8px;border:1px solid #fecaca;display:none}'
    + '.login-note{font-size:11px;color:var(--muted);text-align:center;margin-top:10px;line-height:1.7}'
    // Change 6: profile picker
    + '#profile-picker{display:none;min-height:100vh;background:linear-gradient(160deg,var(--forest),var(--pine) 55%,#42101a);align-items:center;justify-content:center;padding:2rem 1.25rem;flex-direction:column}'
    + '.picker-card{width:100%;max-width:400px;background:var(--white);border-radius:20px;padding:1.75rem;box-shadow:var(--shadow-md)}'
    + '.picker-title{font-family:"Syne",sans-serif;font-size:20px;font-weight:700;color:var(--ink);margin-bottom:6px;text-align:center}'
    + '.picker-sub{font-size:12px;color:var(--muted);text-align:center;margin-bottom:20px}'
    + '.profile-btn{width:100%;background:var(--dew);border:1.5px solid var(--mist);border-radius:12px;padding:14px 16px;margin-bottom:10px;cursor:pointer;text-align:left;font-family:"Instrument Sans",sans-serif;transition:all .15s}'
    + '.profile-btn:hover{background:var(--mist);border-color:var(--fern)}'
    + '.profile-name{font-size:15px;font-weight:600;color:var(--pine);margin-bottom:3px}'
    + '.profile-meta{font-size:11px;color:var(--muted)}'
    + '.btn-back{background:none;border:1px solid rgba(143,212,176,.3);color:var(--mist);font-size:11px;padding:5px 14px;border-radius:20px;cursor:pointer;font-family:"Instrument Sans",sans-serif;margin-top:6px;width:100%}'
    // dashboard
    + '#dashboard{display:none;min-height:100vh}'
    + '.dh{background:linear-gradient(120deg,var(--forest),var(--pine));padding:14px 16px;display:flex;align-items:center;gap:12px;position:sticky;top:0;z-index:100;border-bottom:2px solid var(--gold)}'
    + '.av{width:38px;height:38px;border-radius:11px;background:linear-gradient(135deg,var(--maroon),#42101a);display:flex;align-items:center;justify-content:center;font-family:"Syne",sans-serif;font-size:14px;font-weight:700;color:#fff;flex-shrink:0;border:1.5px solid var(--gold-border)}'
    + '.dh-info{min-width:0;flex:1}'
    + '.dh-name{font-family:"Syne",sans-serif;font-size:16px;font-weight:600;color:var(--mist);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}'
    + '.dh-sub{font-size:10px;color:var(--mint);margin-top:1px}'
    + '.btn-lo{background:rgba(143,212,176,.1);border:1px solid rgba(143,212,176,.2);color:var(--mint);padding:5px 11px;border-radius:20px;font-size:11px;font-family:"Instrument Sans",sans-serif;cursor:pointer;flex-shrink:0}'
    // banner
    + '.info-banner{background:linear-gradient(135deg,var(--forest) 0%,var(--pine) 100%);padding:14px 16px 16px;border-bottom:2px solid var(--gold)}'
    + '.info-ke{font-family:"Syne",sans-serif;font-size:18px;font-weight:600;color:var(--mist);margin-bottom:2px}'
    + '.info-svc{font-size:11px;color:var(--mint);margin-bottom:12px}'
    + '.info-stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:8px}'
    + '.istat{background:rgba(143,212,176,.1);border:1px solid rgba(143,212,176,.18);border-radius:10px;padding:9px 8px;text-align:center}'
    + '.istat-val{font-family:"Syne",sans-serif;font-size:22px;font-weight:700;color:var(--mint);line-height:1}'
    + '.istat-lbl{font-size:9px;color:rgba(143,212,176,.7);text-transform:uppercase;letter-spacing:.08em;margin-top:3px}'
    // tabs
    + '.stale-banner{display:none;background:#92400e;color:#fffbeb;padding:10px 14px;font-size:12px;font-weight:600;text-align:center;border-bottom:1px solid #f59e0b}'
    + '.stale-banner.on{display:block}'
    + '.stale-banner button{margin-left:8px;padding:6px 12px;border:0;border-radius:8px;background:#fffbeb;color:#92400e;font:inherit;font-weight:700;cursor:pointer}'
    + '.tabs{background:var(--white);display:flex;overflow-x:auto;border-bottom:1.5px solid var(--border);position:sticky;top:66px;z-index:9;scrollbar-width:none}.tabs::-webkit-scrollbar{display:none}'
    + '.tb{flex:0 0 auto;min-width:92px;padding:12px 8px;background:none;border:none;border-bottom:2.5px solid transparent;font-family:"Instrument Sans",sans-serif;font-size:12px;color:var(--muted);cursor:pointer;font-weight:500;transition:all .15s;display:flex;align-items:center;justify-content:center;gap:5px;white-space:nowrap}'
    + '.tb.on{color:var(--pine);border-bottom-color:var(--pine);font-weight:600}'
    + '.tc{display:none;padding:16px 14px}.tc.on{display:block}'
    + '.tab-heading{font-family:"Syne",sans-serif;font-size:22px;font-weight:600;color:var(--ink);margin-bottom:2px}'
    + '.tab-sub{font-size:12px;color:var(--muted);margin-bottom:14px}'
    + '.sec-div{display:flex;align-items:center;gap:8px;margin:16px 0 10px;color:#b0c8b8;font-size:10px;letter-spacing:.1em;text-transform:uppercase;font-weight:600}'
    + '.sec-div::before,.sec-div::after{content:"";flex:1;height:1px;background:var(--border)}'
    + '.cur-summary{background:var(--white);border:1px solid var(--border);border-radius:var(--r);padding:14px;margin-bottom:12px;box-shadow:var(--shadow-xs)}'
    + '.cur-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:10px}'
    + '.cur-kpi{background:var(--dew);border:1px solid var(--mist);border-radius:10px;padding:8px;text-align:center}'
    + '.cur-kpi-val{font-family:"Syne",sans-serif;font-size:20px;font-weight:700;color:var(--pine)}'
    + '.cur-kpi-lbl{font-size:10px;color:#6f8f7b;letter-spacing:.06em;text-transform:uppercase}'
    + '.cur-progress{height:8px;background:#e6f3ec;border-radius:999px;overflow:hidden;margin-top:10px}'
    + '.cur-progress-bar{height:100%;background:linear-gradient(90deg,var(--fern),var(--sage));border-radius:999px}'
    + '.cur-acc{background:var(--white);border:1px solid var(--border);border-radius:12px;margin-bottom:8px;overflow:hidden;box-shadow:var(--shadow-xs)}'
    + '.cur-acc-hd{width:100%;display:flex;align-items:center;gap:10px;padding:12px 14px;background:linear-gradient(135deg,var(--dew),var(--white));border:none;border-bottom:1px solid transparent;cursor:pointer;font-family:"Instrument Sans",sans-serif;text-align:left;transition:background .15s}'
    + '.cur-acc.open .cur-acc-hd{border-bottom-color:var(--border);background:var(--dew)}'
    + '.cur-acc-hd:focus{outline:2px solid var(--sage);outline-offset:-2px}'
    + '.cur-acc-title{font-size:14px;font-weight:700;color:var(--pine);flex:1;min-width:0}'
    + '.cur-acc-meta{font-size:11px;color:#6f8f7b;white-space:nowrap}'
    + '.cur-acc-chev{font-size:10px;color:var(--sage);transition:transform .2s ease;flex-shrink:0;line-height:1}'
    + '.cur-acc.open .cur-acc-chev{transform:rotate(-180deg)}'
    + '.cur-acc-bd{display:none;padding:8px 10px 12px;background:var(--parchment)}'
    + '.cur-acc.open .cur-acc-bd{display:block}'
    + '.cur-item{background:var(--white);border:1px solid var(--border);border-radius:12px;padding:12px;margin-bottom:8px}'
    + '.cur-item:last-child{margin-bottom:2px}'
    + '.cur-item.pass{border-color:#9addb5;background:#f3fcf6}'
    + '.cur-h{display:flex;justify-content:space-between;gap:8px;align-items:flex-start}'
    + '.cur-title{font-size:13px;font-weight:600;color:var(--ink)}'
    + '.cur-lv{font-size:10px;color:#7a9a7e;margin-top:2px}'
    + '.cur-obj{font-size:11px;color:#446b56;margin-top:7px;line-height:1.45}'
    + '.cur-ex{font-size:11px;color:#5a7f6a;margin-top:5px;line-height:1.45}'
    + '.badge-pass{background:#dcfce7;color:#166534;border:1px solid #9addb5}'
    + '.badge-cur{background:#e0f2fe;color:#0f4c81;border:1px solid #b8dffd}'
    + '.perf-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:10px}'
    + '.perf-box{background:var(--white);border:1px solid var(--border);border-radius:10px;padding:10px}'
    + '.perf-lbl{font-size:10px;color:#7a9a7e;text-transform:uppercase;letter-spacing:.08em}'
    + '.perf-val{font-family:"Syne",sans-serif;font-size:20px;font-weight:700;color:var(--pine)}'
    + '.trend-row{display:flex;justify-content:space-between;align-items:center;padding:7px 0;border-bottom:1px solid var(--border)}'
    + '.trend-row:last-child{border-bottom:none}'
    + '.ref-box{background:var(--white);border:1px solid var(--border);border-radius:12px;padding:12px;margin-bottom:10px}'
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
    + '.sc-tm{font-size:11px;color:var(--muted);margin-top:3px}'
    + '.sc-class{font-size:12px;color:var(--pine);font-weight:600;margin-top:4px}'
    + '.sc-class-btn{display:inline-flex;align-items:center;gap:5px;margin-top:6px;padding:6px 10px;border-radius:8px;border:1.5px solid #93c5fd;background:#eff6ff;color:#1d4ed8;font-size:12px;font-weight:600;text-decoration:none;max-width:100%;line-height:1.3}'
    + '.sc-class-btn:hover,.sc-class-btn:focus{background:#dbeafe;color:#1e3a8a;border-color:#60a5fa;outline:none}'
    + '.cur-doc-btn{display:inline-flex;align-items:center;gap:4px;margin-top:6px;padding:5px 9px;border-radius:7px;border:1px solid #93c5fd;background:#eff6ff;color:#1d4ed8;font-size:11px;font-weight:600;text-decoration:none}'
    + '.cur-doc-btn:hover,.cur-doc-btn:focus{background:#dbeafe;color:#1e3a8a;border-color:#60a5fa;outline:none}'
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
    + '.mode-btn{padding:10px 8px;border-radius:var(--r-sm);border:1.5px solid var(--border-md);background:var(--white);font-family:"Instrument Sans",sans-serif;font-size:12px;font-weight:500;color:#7a9a7e;cursor:pointer;text-align:center;display:flex;flex-direction:column;align-items:center;gap:4px}'
    + '.mode-btn .mi{font-size:20px}'
    + '.mode-btn.active{border-color:var(--sage);background:var(--dew);color:var(--pine);font-weight:600}'
    + '.slot-card{background:var(--white);border:1px solid var(--border);border-radius:var(--r);padding:14px;margin-bottom:10px;position:relative}'
    + '.slot-num{font-size:10px;font-weight:700;color:var(--sage);text-transform:uppercase;letter-spacing:.08em;margin-bottom:12px;display:flex;align-items:center;gap:6px}'
    + '.sn-badge{background:var(--pine);color:var(--mist);font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px}'
    + '.rm-slot{position:absolute;right:12px;top:12px;background:none;border:1px solid #fecaca;color:var(--red);font-size:12px;cursor:pointer;padding:3px 8px;border-radius:6px;font-family:"Instrument Sans",sans-serif}'
    + '.add-btn{width:100%;padding:11px;border:2px dashed var(--mist);background:transparent;color:var(--sage);border-radius:var(--r-sm);font-family:"Instrument Sans",sans-serif;font-size:13px;font-weight:500;cursor:pointer;margin-bottom:10px;display:flex;align-items:center;justify-content:center;gap:6px}'
    // recurring
    + '.rec-card{background:var(--white);border:1px solid var(--border);border-radius:var(--r);padding:16px;margin-bottom:10px}'
    + '.day-picker{display:flex;gap:6px;flex-wrap:wrap;margin:10px 0}'
    + '.day-chip{width:40px;height:40px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:600;cursor:pointer;border:1.5px solid var(--border-md);background:var(--white);color:#7a9a7e;flex-shrink:0}'
    + '.day-chip.sel{background:var(--pine);border-color:var(--pine);color:#fff}'
    + '.pat-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin:10px 0}'
    + '.pat-btn{padding:10px;border-radius:var(--r-sm);border:1.5px solid var(--border-md);background:var(--white);color:#7a9a7e;font-family:"Instrument Sans",sans-serif;font-size:11px;font-weight:500;cursor:pointer;text-align:center}'
    + '.pat-btn.sel{border-color:var(--sage);background:var(--dew);color:var(--pine);font-weight:600}'
    + '.prev-wrap{display:flex;flex-wrap:wrap;gap:5px;margin-top:10px}'
    + '.prev-chip{font-size:11px;padding:3px 9px;background:var(--dew);color:var(--pine);border:1px solid var(--mist);border-radius:20px}'
    + '.prev-more{font-size:11px;padding:3px 9px;background:var(--parchment);color:#7a9a7e;border:1px solid var(--border);border-radius:20px}'
    // submit / result
    + '.btn-sub{width:100%;padding:13px;border-radius:var(--r-sm);background:var(--forest);color:var(--mist);border:none;font-family:"Instrument Sans",sans-serif;font-size:14px;font-weight:600;cursor:pointer;margin-top:4px}'
    + '.btn-sub:disabled{opacity:.5;cursor:default}'
    + '.result{margin-top:10px;padding:10px 14px;border-radius:var(--r-sm);font-size:12px;font-weight:500;display:none}'
    + '.r-ok{background:#d1fae5;color:#065f46;border:1px solid #6ee7b7}'
    + '.r-err{background:var(--red-pale);color:var(--red);border:1px solid #fecaca}'
    // payments
    + '.pay-card{background:var(--white);border:1px solid var(--border);border-radius:var(--r);padding:14px;margin-bottom:9px;box-shadow:var(--shadow-xs)}'
    + '.pay-top{display:flex;justify-content:space-between;align-items:flex-start;gap:10px}'
    + '.pay-amt{font-family:"Syne",sans-serif;font-size:24px;font-weight:700;color:var(--pine);line-height:1}'
    + '.pay-meta{font-size:11px;color:#7a9a7e;margin-top:5px}'
    + '.pay-txn{font-size:10px;color:#b0c8b8;margin-top:3px}'
    + '.pay-cta{display:flex;justify-content:center;margin-top:16px}'
    + '.btn-pay{display:inline-flex;align-items:center;justify-content:center;gap:8px;background:var(--forest);color:var(--mist);padding:12px 24px;text-decoration:none;border:none;border-radius:var(--r-sm);font-size:13px;font-weight:600;font-family:"Instrument Sans",sans-serif;cursor:pointer}'
    + '.btn-pay:disabled{opacity:.55;cursor:not-allowed}'
    + '.pay-overlay{position:fixed;inset:0;z-index:200;background:rgba(20,40,15,.55);display:none;align-items:flex-end;justify-content:center;padding:0}'
    + '.pay-overlay.on{display:flex}'
    + '.pay-modal{background:#fff;width:100%;max-width:520px;max-height:92vh;overflow:auto;border-radius:18px 18px 0 0;padding:18px 16px 28px;box-shadow:var(--shadow-md);position:relative;z-index:1}'
    + '@media(min-width:640px){.pay-overlay{align-items:center;padding:18px}.pay-modal{border-radius:16px;max-height:90vh}}'
    + '.pay-modal h3{font-family:"Syne",sans-serif;font-size:20px;color:var(--ink);margin:0 0 4px}'
    + '.pay-modal .pay-sub{font-size:12px;color:var(--muted);margin-bottom:14px;line-height:1.45}'
    + '.pay-grid{display:grid;gap:10px}'
    + '.pay-grid .fl{display:block;font-size:11px;font-weight:600;color:var(--muted);margin-bottom:4px;text-transform:uppercase;letter-spacing:.04em}'
    + '.pay-grid .fi,.pay-grid select.fi,.pay-grid textarea.fi{width:100%;padding:11px 12px;border:1px solid var(--border);border-radius:10px;font:inherit;background:#fff;color:var(--ink)}'
    + '.pay-grid .fi[readonly]{background:#f3f6f2;color:#4b5563}'
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
    + '.pay-choice-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}'
    + '.pay-choice-grid.three{grid-template-columns:1fr 1fr 1fr}'
    + '.pay-choice{appearance:none;-webkit-appearance:none;display:flex;flex-direction:column;align-items:flex-start;gap:2px;width:100%;text-align:left;padding:12px 11px;border:1.5px solid var(--border-md);border-radius:12px;background:#f8faf6;color:var(--ink);font:inherit;cursor:pointer;transition:border-color .15s,background .15s,box-shadow .15s}'
    + '.pay-choice .pc-title{font-size:13px;font-weight:700;color:var(--ink);line-height:1.2}'
    + '.pay-choice .pc-sub{font-size:10px;color:var(--muted);line-height:1.35}'
    + '.pay-choice.on{border-color:var(--pine);background:#eaf5ee;box-shadow:0 0 0 1px var(--pine)}'
    + '.pay-choice.on .pc-title{color:var(--pine)}'
    + '.pay-choice:disabled{opacity:.55;cursor:not-allowed}'
    + '.pay-choice:disabled:not(.on){background:#f3f4f3}'
    + '.pay-radio-list{display:flex;flex-direction:column;gap:8px}'
    + '.pay-radio{display:flex;align-items:flex-start;gap:10px;padding:12px 12px;border:1.5px solid var(--border-md);border-radius:12px;background:#f8faf6;cursor:pointer}'
    + '.pay-radio:has(input:checked){border-color:var(--pine);background:#eaf5ee;box-shadow:0 0 0 1px var(--pine)}'
    + '.pay-radio input[type=radio]{margin-top:2px;width:18px;height:18px;accent-color:var(--pine);flex-shrink:0}'
    + '.pay-radio .pc-title{font-size:13px;font-weight:700;color:var(--ink);line-height:1.25;display:block}'
    + '.pay-radio .pc-sub{font-size:11px;color:var(--muted);line-height:1.35;display:block;margin-top:2px}'
    + '.pay-radio:has(input:disabled){opacity:.6;cursor:not-allowed}'
    // shop + orders
    + '.shop-note{background:#eff6ff;color:#1e3a8a;border:1px solid #bfdbfe;border-radius:10px;padding:10px 12px;font-size:11px;line-height:1.5;margin-bottom:12px}'
    + '.shop-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,280px));gap:11px;justify-content:start}'
    + '.prod-card{background:#fff;border:1px solid var(--border);border-radius:14px;overflow:hidden;box-shadow:var(--shadow-xs);display:flex;flex-direction:column;width:100%;max-width:280px}'
    + '.prod-img{width:100%;aspect-ratio:1.25/1;object-fit:cover;background:var(--dew)}'
    + '.prod-body{padding:11px;display:flex;flex-direction:column;gap:7px;flex:1}'
    + '.prod-name{font-size:14px;font-weight:700;color:var(--ink)}'
    + '.prod-opt{font-size:10px;color:var(--sage);text-transform:uppercase;letter-spacing:.07em}'
    + '.prod-price{font-family:"Syne",sans-serif;font-size:18px;font-weight:700;color:var(--pine)}'
    + '.prod-unpriced{font-size:11px;color:var(--gold)}'
    + '.kit-includes{font-size:11px;color:var(--muted);line-height:1.5;background:var(--dew);padding:8px;border-radius:8px}'
    + '.prod-control{width:100%;padding:8px;border:1px solid var(--border-md);border-radius:8px;font:12px \"Instrument Sans\",sans-serif;background:#fff}'
    + '.prod-qty{width:55px}'
    + '.prod-add{width:100%;padding:9px;border:0;border-radius:8px;background:var(--pine);color:#fff;font:600 12px \"Instrument Sans\",sans-serif;cursor:pointer;margin-top:auto}.prod-add:disabled{opacity:.45}'
    + '.cart-box{margin-top:16px;background:#fff;border:1.5px solid var(--gold-border);border-radius:14px;padding:14px}'
    + '.cart-line{display:flex;justify-content:space-between;gap:8px;padding:9px 0;border-bottom:1px solid var(--border);font-size:12px}.cart-line:last-child{border:0}'
    + '.cart-total{display:flex;justify-content:space-between;font-size:18px;font-weight:700;color:var(--pine);padding:12px 0}'
    + '.cart-remove{background:none;border:0;color:var(--red);font-size:11px;cursor:pointer}'
    + '.checkout-box{text-align:center;background:#fff;border:1px solid var(--border);border-radius:14px;padding:16px;margin-top:14px}'
    + '.checkout-qr{width:210px;max-width:80%;margin:12px auto;display:block}'
    + '.order-card{background:#fff;border:1px solid var(--border);border-radius:14px;padding:14px;margin-bottom:11px;box-shadow:var(--shadow-xs)}'
    + '.order-top{display:flex;justify-content:space-between;gap:8px;align-items:flex-start}.order-id{font-family:\"DM Mono\",monospace;font-size:11px;color:var(--muted)}'
    + '.order-status{font-size:10px;font-weight:700;border-radius:20px;padding:4px 9px;background:var(--dew);color:var(--pine)}'
    + '.order-items{font-size:12px;line-height:1.55;margin:10px 0;color:var(--ink)}'
    + '.order-meta{font-size:11px;color:var(--muted);margin-top:4px}'
    + '.order-confirm{margin-top:12px;padding:12px;background:#eff6ff;border-radius:10px;color:#1e3a8a;font-size:12px}'
    + '.order-confirm-actions{display:flex;gap:8px;margin-top:9px}.order-confirm-actions button{flex:1;padding:8px;border-radius:8px;border:1px solid #93c5fd;background:#fff;color:#1d4ed8;font-weight:600;cursor:pointer}'
    + '@media(max-width:390px){.shop-grid{grid-template-columns:1fr}}'
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
    'var PAYMENT_LINK = ' + JSON.stringify(payLink) + ';',
    'var LOGO_URL = ' + JSON.stringify(logoUrl) + ';',
    'var DAYS = ["Su","Mo","Tu","We","Th","Fr","Sa"];',
    'var pendingPhone = "";',  // Change 6: remember phone for profile picker
    'var SHOP_CATALOG = [];',
    'var SHOP_CART = [];',
    'var SHOP_ORDERS = [];',
    'var SHOP_PAYMENT_SUMMARY = null;',
    'var SHOP_LOADED = false;',
    'var SHOP_ORDERS_LOADED = false;',
    'var PAY_FORM = { paymentFor: "Riding Classes", amount: "", screenshotData: "", screenshotName: "" };',
    'var PAY_QR_TIMER = null;',
    'var APP_UI_VERSION = ' + JSON.stringify(appUiVersion) + ';',
    'window.__APP_UI_VERSION__ = APP_UI_VERSION;',
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
    '',
    '/** Drive /file/d/.../view links are HTML pages — never use them as <img src>. */',
    'function shopImageSrc(url) {',
    '  url = String(url || "").trim();',
    '  if (!url) return LOGO_URL || "";',
    '  if (url.indexOf("data:image/") === 0) return url;',
    '  var m = url.match(/\\/file\\/d\\/([a-zA-Z0-9_-]+)/) || url.match(/[?&]id=([a-zA-Z0-9_-]+)/);',
    '  if (m && (/\\/view/i.test(url) || /usp=drivesdk/i.test(url) || /\\/file\\/d\\//i.test(url))',
    '      && url.indexOf("thumbnail") < 0 && url.indexOf("uc?") < 0 && url.indexOf("export=view") < 0) {',
    '    return "https://drive.google.com/thumbnail?id=" + m[1] + "&sz=w800";',
    '  }',
    '  return url;',
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

    'function toggleCurAcc(btn) {',
    '  var acc = btn && btn.closest ? btn.closest(".cur-acc") : null;',
    '  if (!acc) return;',
    '  acc.classList.toggle("open");',
    '  btn.setAttribute("aria-expanded", acc.classList.contains("open") ? "true" : "false");',
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
    '  RD = null; slotCount = 1; bookMode = "single"; SHOP_CART = []; SHOP_ORDERS = []; SHOP_PAYMENT_SUMMARY = null; SHOP_LOADED = false; SHOP_ORDERS_LOADED = false;',
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
    '  renderSessions(); renderCurriculum(); renderBook(); renderPayments();',
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
    '  if (name === "shop") loadShop();',
    '  if (name === "orders") loadShopOrders();',
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
    '  var clsTxt = "";',
    '  if (s.level || s.classNumber) { clsTxt = (s.level || "") + (s.classNumber ? (s.level ? " \\u00b7 " : "") + "Class " + s.classNumber : "") + (s.classTitle ? ": " + s.classTitle : ""); }',
    '  var classBlock = "";',
    '  if (clsTxt) {',
    '    if (s.docLink) {',
    '      classBlock = "<a class=\\"sc-class-btn\\" href=\\"" + esc(s.docLink) + "\\" target=\\"_blank\\" rel=\\"noopener noreferrer\\" title=\\"Open class document\\">&#128196; " + esc(clsTxt) + "</a>";',
    '    } else {',
    '      classBlock = "<div class=\\"sc-class\\">&#127942; " + esc(clsTxt) + "</div>";',
    '    }',
    '  }',
    '  return "<div class=\\"sc " + cSt + "\\">"',
    '    + "<div class=\\"si\\">"',
    '    + "<div class=\\"sc-top\\"><div>"',
    '    + "<div class=\\"sc-svc\\">" + esc(s.service) + "</div>"',
    '    + "<div class=\\"sc-dt\\">&#128197; " + esc(s.date || "TBD") + "</div></div>"',
    '    + "<span class=\\"bdg " + bCls + "\\">" + lbl + "</span>"',
    '    + "</div>"',
    '    + classBlock',
    '    + (s.timeSlot ? "<div class=\\"sc-tm\\">&#128336; " + esc(s.timeSlot) + "</div>" : "")',
    '    + (s.trainerLabel ? "<div class=\\"sc-tm\\">&#129485; " + esc(s.trainerLabel) + "</div>" : "")',
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

    // ── curriculum + performance ───────────────────────────
    'function renderCurriculum() {',
    '  var cur = RD.curriculum || { items: [], passedCount: 0, totalCount: 0 };',
    '  var perf = RD.performanceChart || null;',
    '  var total = Number(cur.totalCount || 0);',
    '  var passed = Number(cur.passedCount || 0);',
    '  var pct = total ? Math.round((passed / total) * 100) : 0;',
    '  var h = "<div class=\\"tab-heading\\">Curriculum Progress</div>";',
    '  h += "<div class=\\"tab-sub\\">Detailed class-by-class journey from your CURRICULUM tab</div>";',
    '  h += "<div class=\\"cur-summary\\">";',
    '  h += "<div><strong>Current:</strong> " + esc(cur.currentLevel || "-") + " · Class " + esc(cur.currentClassNumber || "-") + (cur.currentTitle ? " · " + esc(cur.currentTitle) : "") + "</div>";',
    '  var curDoc = "";',
    '  (cur.items || []).forEach(function(it){',
    '    if (String(it.level||"") === String(cur.currentLevel||"") && String(it.classNumber||"") === String(cur.currentClassNumber||"") && it.docLink) curDoc = it.docLink;',
    '  });',
    '  if (curDoc) h += "<div style=\\"margin-top:8px\\"><a class=\\"cur-doc-btn\\" href=\\"" + esc(curDoc) + "\\" target=\\"_blank\\" rel=\\"noopener noreferrer\\">&#128196; Open current class document</a></div>";',
    '  h += "<div class=\\"cur-progress\\"><div class=\\"cur-progress-bar\\" style=\\"width:" + pct + "%\\"></div></div>";',
    '  h += "<div class=\\"cur-grid\\">";',
    '  h += "<div class=\\"cur-kpi\\"><div class=\\"cur-kpi-val\\">" + passed + "/" + total + "</div><div class=\\"cur-kpi-lbl\\">Classes Passed</div></div>";',
    '  h += "<div class=\\"cur-kpi\\"><div class=\\"cur-kpi-val\\">" + pct + "%</div><div class=\\"cur-kpi-lbl\\">Completion</div></div>";',
    '  h += "</div>";',
    '  if (perf && perf.radar) {',
    '    h += "<div class=\\"perf-grid\\">";',
    '    h += "<div class=\\"perf-box\\"><div class=\\"perf-lbl\\">Safety</div><div class=\\"perf-val\\">" + Number(perf.radar.safety || 0).toFixed(2) + "</div></div>";',
    '    h += "<div class=\\"perf-box\\"><div class=\\"perf-lbl\\">Riding</div><div class=\\"perf-val\\">" + Number(perf.radar.riding || 0).toFixed(2) + "</div></div>";',
    '    h += "<div class=\\"perf-box\\"><div class=\\"perf-lbl\\">Knowledge</div><div class=\\"perf-val\\">" + Number(perf.radar.knowledge || 0).toFixed(2) + "</div></div>";',
    '    h += "<div class=\\"perf-box\\"><div class=\\"perf-lbl\\">Attitude</div><div class=\\"perf-val\\">" + Number(perf.radar.attitude || 0).toFixed(2) + "</div></div>";',
    '    h += "</div>";',
    '    if ((perf.trend || []).length) {',
    '      h += "<div class=\\"perf-box\\" style=\\"margin-top:8px\\"><div class=\\"perf-lbl\\" style=\\"margin-bottom:6px\\">Recent Trend</div>";',
    '      perf.trend.forEach(function(t){ h += "<div class=\\"trend-row\\"><span>" + esc(t.date || "") + "</span><strong>" + Number(t.avg || 0).toFixed(2) + "</strong></div>"; });',
    '      h += "</div>";',
    '    }',
    '  }',
    '  h += "</div>";',
    '  var items = cur.items || [];',
    '  if (!items.length) {',
    '    h += "<div class=\\"empty-st\\">Curriculum rows not found. Please check sheet name CURRICULUM/circulum.</div>";',
    '    document.getElementById("tc-curriculum").innerHTML = h;',
    '    return;',
    '  }',
    '  var byLevel = {};',
    '  var levelOrder = [];',
    '  var seenLv = {};',
    '  items.forEach(function(it) {',
    '    var lv = String(it.level || "").trim() || "—";',
    '    if (!seenLv[lv]) { seenLv[lv] = true; levelOrder.push(lv); }',
    '    if (!byLevel[lv]) byLevel[lv] = [];',
    '    byLevel[lv].push(it);',
    '  });',
    '  levelOrder.forEach(function(lv) {',
    '    var sub = byLevel[lv] || [];',
    '    var st = (cur.levels || []).filter(function(x) { return String(x.level || "") === String(lv); })[0];',
    '    var passedN = st ? Number(st.passed || 0) : sub.filter(function(i) { return i.passed; }).length;',
    '    var totalN = st ? Number(st.total || 0) : sub.length;',
    '    var stTxt = totalN ? (passedN + "/" + totalN + " passed") : "";',
    '    var isOpen = String(lv) === String(cur.currentLevel || "");',
    '    h += "<div class=\\"cur-acc" + (isOpen ? " open" : "") + "\\">";',
    '    h += "<button type=\\"button\\" class=\\"cur-acc-hd\\" aria-expanded=\\"" + (isOpen ? "true" : "false") + "\\" onclick=\\"toggleCurAcc(this)\\">";',
    '    h += "<span class=\\"cur-acc-title\\">" + esc(lv) + "</span>";',
    '    h += "<span class=\\"cur-acc-meta\\">" + esc(stTxt) + "</span>";',
    '    h += "<span class=\\"cur-acc-chev\\" aria-hidden=\\"true\\">▼</span></button>";',
    '    h += "<div class=\\"cur-acc-bd\\">";',
    '    sub.forEach(function(it) {',
    '      var badge = it.passed ? "<span class=\\"bdg badge-pass\\">Passed</span>" : ((String(it.level||"") === String(cur.currentLevel||"") && String(it.classNumber||"") === String(cur.currentClassNumber||"")) ? "<span class=\\"bdg badge-cur\\">Current</span>" : "<span class=\\"bdg b-up\\">Pending</span>");',
    '      h += "<div class=\\"cur-item " + (it.passed ? "pass" : "") + "\\">";',
    '      h += "<div class=\\"cur-h\\"><div><div class=\\"cur-title\\">" + esc(it.title || ("Class " + it.classNumber)) + "</div><div class=\\"cur-lv\\">Class " + esc(it.classNumber || "") + "</div></div>" + badge + "</div>";',
    '      if (it.objective) h += "<div class=\\"cur-obj\\"><strong>Objective:</strong> " + esc(it.objective) + "</div>";',
    '      if (it.exercise) h += "<div class=\\"cur-ex\\"><strong>Exercise:</strong> " + esc(it.exercise) + "</div>";',
    '      if (it.criteria) h += "<div class=\\"cur-ex\\"><strong>Criteria:</strong> " + esc(it.criteria) + "</div>";',
    '      if (it.docLink) h += "<a class=\\"cur-doc-btn\\" href=\\"" + esc(it.docLink) + "\\" target=\\"_blank\\" rel=\\"noopener noreferrer\\">&#128196; Open class document</a>";',
    '      h += "</div>";',
    '    });',
    '    h += "</div></div>";',
    '  });',
    '  document.getElementById("tc-curriculum").innerHTML = h;',
    '}',

    // ── book ───────────────────────────────────────────────
    'function renderBook() {',
    '  slotCount = 1; bookMode = "single"; recurSelDays = []; recurPattern = "";',
    '  var h = "<div class=\\"tab-heading\\">Book Sessions</div>";',
    '  h += "<div class=\\"tab-sub\\">Pick specific dates or a weekly pattern for a whole month.</div>";',
    '  if (RD.nextSuggestion) {',
    '    h += "<div class=\\"ref-box\\"><div style=\\"font-size:11px;color:#6f8f7b;text-transform:uppercase;letter-spacing:.06em\\">Suggested Next Class</div>";',
    '    h += "<div style=\\"font-weight:600;margin-top:4px\\">" + esc(RD.nextSuggestion.level || "") + " · Class " + esc(RD.nextSuggestion.classNumber || "") + " · " + esc(RD.nextSuggestion.title || "") + "</div>";',
    '    if (RD.nextSuggestion.objective) h += "<div style=\\"font-size:11px;color:#476d59;margin-top:6px\\"><strong>Objective:</strong> " + esc(RD.nextSuggestion.objective) + "</div>";',
    '    if (RD.nextSuggestion.docLink) h += "<div style=\\"margin-top:8px\\"><a class=\\"cur-doc-btn\\" href=\\"" + esc(RD.nextSuggestion.docLink) + "\\" target=\\"_blank\\" rel=\\"noopener noreferrer\\">&#128196; Open class document</a></div>";',
    '    h += "</div>";',
    '  }',
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

    '/** YYYY-MM-DD using local calendar date — never use toISOString() (UTC shifts day for IST etc.). */',
    'function formatLocalYMD(dt) {',
    '  if (!dt || !(dt instanceof Date) || isNaN(dt.getTime())) return "";',
    '  var y = dt.getFullYear();',
    '  var m = String(dt.getMonth() + 1).padStart(2, "0");',
    '  var day = String(dt.getDate()).padStart(2, "0");',
    '  return y + "-" + m + "-" + day;',
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
    '      requests.push({ service: rsvc.name, date: formatLocalYMD(d), timeSlot: rti ? rti.value : "", participants: rpax });',
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

    // ── shop ───────────────────────────────────────────────
    'function loadShop() {',
    '  if (SHOP_LOADED) { renderShop(); return; }',
    '  document.getElementById("tc-shop").innerHTML = "<div class=\\"empty-st\\">Loading riding essentials...</div>";',
    '  google.script.run',
    '    .withSuccessHandler(function(res) {',
    '      if (!res || !res.success) { document.getElementById("tc-shop").innerHTML = "<div class=\\"empty-st\\">" + esc((res&&res.message)||"Could not load shop.") + "</div>"; return; }',
    '      SHOP_CATALOG = res.products || []; SHOP_LOADED = true; renderShop();',
    '    })',
    '    .withFailureHandler(function(e) { document.getElementById("tc-shop").innerHTML = "<div class=\\"empty-st\\">Error: " + esc(e.message) + "</div>"; })',
    '    .getShopCatalog();',
    '}',
    '',
    'function renderShop() {',
    '  var h = "<div class=\\"tab-heading\\">Riding Essentials</div>";',
    '  h += "<div class=\\"tab-sub\\">Select your product, size and quantity.</div>";',
    '  h += "<div class=\\"shop-note\\">Unknown sizes can be coordinated in the WhatsApp group. Products without a configured price cannot be ordered yet.</div>";',
    '  if (!SHOP_CATALOG.length) h += "<div class=\\"empty-st\\">No active products configured.</div>";',
    '  h += "<div class=\\"shop-grid\\">";',
    '  SHOP_CATALOG.forEach(function(p, i) {',
    '    var sizeCtl = "";',
    '    if (p.sizeMode === "select") {',
    '      sizeCtl = "<select class=\\"prod-control\\" id=\\"shop-size-" + i + "\\"><option value=\\"\\">Select size</option>" + (p.sizes||[]).map(function(s){return "<option value=\\"" + esc(s) + "\\">" + esc(s) + "</option>";}).join("") + "</select>";',
    '    } else {',
    '      sizeCtl = "<input class=\\"prod-control\\" id=\\"shop-size-" + i + "\\" maxlength=\\"80\\" placeholder=\\"" + (p.category === "Kit" ? "Helmet / body / pants / shoe sizes" : (p.sizeMode === "uk-size" ? "Enter UK shoe size" : "Enter size")) + "\\">";',
    '    }',
    '    h += "<div class=\\"prod-card\\"><img class=\\"prod-img\\" src=\\"" + esc(shopImageSrc(p.imageUrl||"")) + "\\" alt=\\"" + esc(p.product) + "\\" loading=\\"lazy\\" onerror=\\"this.onerror=null;this.style.objectFit=\'contain\';this.src=\\\'" + esc(LOGO_URL||"") + "\\\'\\">";',
    '    h += "<div class=\\"prod-body\\"><div><div class=\\"prod-name\\">" + esc(p.product) + "</div><div class=\\"prod-opt\\">" + esc(p.option||p.category||"") + "</div></div>";',
    '    if((p.includedItems||[]).length) h += "<div class=\\"kit-includes\\"><strong>Includes:</strong><br>" + p.includedItems.map(esc).join("<br>") + "</div>";',
    '    h += p.orderable ? "<div class=\\"prod-price\\">&#8377;" + Number(p.price).toLocaleString("en-IN") + "</div>" : "<div class=\\"prod-unpriced\\">Price being updated</div>";',
    '    h += sizeCtl + "<div><label class=\\"hint\\">Quantity</label><br><input type=\\"number\\" class=\\"prod-control prod-qty\\" id=\\"shop-qty-" + i + "\\" min=\\"1\\" max=\\"10\\" value=\\"1\\"></div>";',
    '    h += "<button class=\\"prod-add\\" data-shop-add=\\"" + i + "\\"" + (p.orderable?"":" disabled") + ">Add to cart</button></div></div>";',
    '  });',
    '  h += "</div>";',
    '  if (SHOP_CART.length) {',
    '    var total = 0; h += "<div class=\\"cart-box\\"><div class=\\"tab-heading\\" style=\\"font-size:17px\\">Your Cart</div>";',
    '    SHOP_CART.forEach(function(x, i) { total += x.price*x.quantity; h += "<div class=\\"cart-line\\"><div><strong>" + esc(x.product) + (x.option&&x.option!=="Standard"?" — "+esc(x.option):"") + "</strong><br><span class=\\"hint\\">Size " + esc(x.size) + " · Qty " + x.quantity + "</span>" + ((x.includedItems||[]).length?"<div class=\\"hint\\">Includes: "+x.includedItems.map(esc).join(", ")+"</div>":"") + "</div><div style=\\"text-align:right\\">&#8377;" + Number(x.price*x.quantity).toLocaleString("en-IN") + "<br><button class=\\"cart-remove\\" data-cart-remove=\\"" + i + "\\">Remove</button></div></div>"; });',
    '    h += "<div class=\\"cart-total\\"><span>Total</span><span>&#8377;" + Number(total).toLocaleString("en-IN") + "</span></div><button class=\\"btn-sub\\" id=\\"shop-checkout\\">Place Order</button><div id=\\"shop-result\\" class=\\"result\\"></div></div>";',
    '  }',
    '  var root = document.getElementById("tc-shop"); root.innerHTML = h;',
    '  root.querySelectorAll("[data-shop-add]").forEach(function(b){ b.addEventListener("click",function(){ addShopProduct(Number(b.getAttribute("data-shop-add"))); }); });',
    '  root.querySelectorAll("[data-cart-remove]").forEach(function(b){ b.addEventListener("click",function(){ SHOP_CART.splice(Number(b.getAttribute("data-cart-remove")),1); renderShop(); }); });',
    '  var checkout = document.getElementById("shop-checkout"); if (checkout) checkout.addEventListener("click", placeShopOrderFromCart);',
    '}',
    '',
    'function addShopProduct(i) {',
    '  var p = SHOP_CATALOG[i]; if (!p || !p.orderable) return;',
    '  var size = String((document.getElementById("shop-size-" + i)||{}).value||"").trim();',
    '  var qty = Math.floor(Number((document.getElementById("shop-qty-" + i)||{}).value||1));',
    '  if (!size) { toast("Select or enter a size"); return; }',
    '  if (qty < 1 || qty > 10) { toast("Quantity must be 1 to 10"); return; }',
    '  var opt = String(p.option||"").trim().toLowerCase();',
    '  var existing = SHOP_CART.filter(function(x){return x.productId===p.productId && String(x.option||"").trim().toLowerCase()===opt && x.size.toLowerCase()===size.toLowerCase();})[0];',
    '  if (existing) existing.quantity = Math.min(10, existing.quantity + qty);',
    '  else SHOP_CART.push({productId:p.productId,product:p.product,option:p.option,size:size,quantity:qty,price:Number(p.price),includedItems:p.includedItems||[]});',
    '  toast("Added to cart"); renderShop();',
    '}',
    '',
    'function placeShopOrderFromCart() {',
    '  if (!RD || !SHOP_CART.length) return;',
    '  var btn = document.getElementById("shop-checkout"); if (btn) { btn.disabled=true; btn.textContent="Placing order..."; }',
    '  var requestId = "SHOP-" + Date.now() + "-" + Math.floor(Math.random()*100000);',
    '  google.script.run',
    '    .withSuccessHandler(function(res) {',
    '      if (btn) { btn.disabled=false; btn.textContent="Place Order"; }',
    '      if (!res || !res.success) { showResult("shop-result",false,(res&&res.message)||"Order failed"); return; }',
    '      SHOP_CART=[]; SHOP_ORDERS_LOADED=false; renderShop();',
    '      var o=res.order||{}; var due=Number(res.paymentDue==null?o.total:res.paymentDue); var box="<div class=\\"checkout-box\\"><div class=\\"tab-heading\\">Order placed</div><div class=\\"order-id\\">" + esc(o.orderId||"") + "</div><div class=\\"prod-price\\" style=\\"margin-top:10px\\">" + (due>0?"Total shopping balance: &#8377;"+due.toLocaleString("en-IN"):"Covered by existing shopping payment credit") + "</div>";',
    '      if (res.qrUrl) box += "<img class=\\"checkout-qr\\" src=\\"" + esc(res.qrUrl) + "\\" alt=\\"Payment QR\\">";',
    '      box += "<div class=\\"order-meta\\">UPI: " + esc(res.upiId||"") + "<br>Reference: " + esc(o.upiReference||o.orderId||"") + "</div>";',
    '      if(due>0) box += "<button type=\\"button\\" class=\\"btn-pay\\" style=\\"margin-top:12px\\" id=\\"shop-pay-now\\">Make shopping payment</button><div class=\\"shop-note\\" style=\\"margin-top:12px\\">Pay here in My Rides. Choose amount for your shopping balance. Payments apply across all your orders.</div>";',
    '      box += "</div>";',
    '      document.getElementById("tc-shop").insertAdjacentHTML("beforeend",box); toast("Order placed");',
    '      var payBtn=document.getElementById("shop-pay-now");',
    '      if(payBtn) payBtn.addEventListener("click", function(){ openPortalPaymentForm({paymentFor:"Shopping Kit / Equipment", amount:due}); });',
    '    })',
    '    .withFailureHandler(function(e) { if(btn){btn.disabled=false;btn.textContent="Place Order";} showResult("shop-result",false,"Error: "+e.message); })',
    '    .placeShopOrder({keNo:RD.keNo,token:RD.shopToken,clientRequestId:requestId,items:SHOP_CART.map(function(x){return {productId:x.productId,option:x.option||"",size:x.size,quantity:x.quantity};})});',
    '}',
    '',
    'function loadShopOrders(force) {',
    '  if (SHOP_ORDERS_LOADED && !force) { renderShopOrders(); return; }',
    '  document.getElementById("tc-orders").innerHTML="<div class=\\"empty-st\\">Loading orders...</div>";',
    '  google.script.run.withSuccessHandler(function(res){',
    '    if(!res||!res.success){document.getElementById("tc-orders").innerHTML="<div class=\\"empty-st\\">"+esc((res&&res.message)||"Could not load orders.")+"</div>";return;}',
    '    SHOP_ORDERS=res.orders||[]; SHOP_PAYMENT_SUMMARY=res.paymentSummary||null; SHOP_ORDERS_LOADED=true; renderShopOrders();',
    '  }).withFailureHandler(function(e){document.getElementById("tc-orders").innerHTML="<div class=\\"empty-st\\">Error: "+esc(e.message)+"</div>";}).getShopOrdersForRider(RD.keNo,RD.shopToken);',
    '}',
    '',
    'function renderShopOrders() {',
    '  var h="<div class=\\"tab-heading\\">My Orders</div><div class=\\"tab-sub\\">Shopping payments are applied across all orders.</div>";',
    '  if(SHOP_PAYMENT_SUMMARY){var s=SHOP_PAYMENT_SUMMARY;h+="<div class=\\"cart-box\\" style=\\"margin-top:0;margin-bottom:12px\\"><div class=\\"cart-line\\"><span>Total ordered</span><strong>&#8377;"+Number(s.totalOrdered||0).toLocaleString("en-IN")+"</strong></div><div class=\\"cart-line\\"><span>Shopping payments</span><strong>&#8377;"+Number(s.totalPaid||0).toLocaleString("en-IN")+"</strong></div><div class=\\"cart-total\\"><span>"+(Number(s.extraPaid||0)>0?"Extra paid":"Remaining")+"</span><span>&#8377;"+Number((s.extraPaid||0)>0?s.extraPaid:s.balance||0).toLocaleString("en-IN")+"</span></div></div>";}',
    '  if(!SHOP_ORDERS.length) h+="<div class=\\"empty-st\\">No shop orders yet.</div>";',
    '  SHOP_ORDERS.forEach(function(o){',
    '    h+="<div class=\\"order-card\\"><div class=\\"order-top\\"><div><div class=\\"order-id\\">"+esc(o.orderId)+"</div><div class=\\"order-meta\\">"+esc(o.createdAt||"")+"</div></div><span class=\\"order-status\\">"+esc(o.orderStatus)+"</span></div>";',
    '    h+="<div class=\\"order-items\\">"+(o.items||[]).map(function(x){var d=x.deliveredAt?" &#10003; Delivered "+esc(String(x.deliveredAt).slice(0,10))+(x.deliveredBy?" by "+esc(x.deliveredBy):""):" · Awaiting delivery";return esc(x.product)+(x.option&&x.option!=="Standard"?" — "+esc(x.option):"")+" · "+esc(x.size)+" × "+Number(x.quantity||1)+d;}).join("<br>")+"</div>";',
    '    h+="<div class=\\"cart-total\\" style=\\"font-size:15px;padding:6px 0\\"><span>Total</span><span>&#8377;"+Number(o.total||0).toLocaleString("en-IN")+"</span></div>";',
    '    h+="<div class=\\"order-meta\\">Payment: "+esc(o.paymentStatus)+" · Paid &#8377;"+Number(o.paidAmount||0).toLocaleString("en-IN")+" · Balance &#8377;"+Number(o.balance||0).toLocaleString("en-IN")+(o.deliveredAt?"<br>Delivered: "+esc(o.deliveredAt):"")+"</div>";',
    '    if(o.deliveryNotes) h+="<div class=\\"shop-note\\" style=\\"margin-top:9px\\">"+esc(o.deliveryNotes)+"</div>";',
    '    if(o.paymentStatus!=="Paid") h+="<div style=\\"margin-top:10px\\"><button type=\\"button\\" class=\\"btn-pay\\" data-shop-pay=\\""+Number(o.balance||0)+"\\">Make shopping payment</button><div class=\\"hint\\">Pay the remaining shopping balance in My Rides.</div></div>";',
    '    h+="</div>";',
    '  });',
    '  document.getElementById("tc-orders").innerHTML=h;',
    '  document.querySelectorAll("[data-shop-pay]").forEach(function(b){',
    '    b.addEventListener("click", function(){',
    '      openPortalPaymentForm({paymentFor:"Shopping Kit / Equipment", amount:Number(b.getAttribute("data-shop-pay")||0)});',
    '    });',
    '  });',
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
    '      if (p.paymentType) r += "<div class=\\"pay-meta\\">" + esc(p.paymentType) + "</div>";',
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
    '  if (makePay) makePay.addEventListener("click", function(){ openPortalPaymentForm({paymentFor:"Riding Classes"}); });',
    '}',
    '',
    // ── in-portal payment form ───────────────────────────────
    'function openPortalPaymentForm(opts) {',
    '  opts = opts || {};',
    '  if (!RD) { toast("Please sign in first"); return; }',
    '  PAY_FORM.paymentFor = opts.paymentFor || "Riding Classes";',
    '  PAY_FORM.amount = opts.amount != null && opts.amount !== "" ? Number(opts.amount) : "";',
    '  PAY_FORM.screenshotData = "";',
    '  PAY_FORM.screenshotName = "";',
    '  document.getElementById("pay-ke").value = RD.keNo || "";',
    '  document.getElementById("pay-phone").value = RD.phone || "";',
    '  document.getElementById("pay-amount").value = PAY_FORM.amount !== "" ? PAY_FORM.amount : "";',
    '  document.getElementById("pay-txn").value = "";',
    '  document.getElementById("pay-pan").value = RD.pan || "";',
    '  setPayChoice("pay-mode", "UPI");',
    '  setPayForRadio(PAY_FORM.paymentFor);',
    '  var isShop = String(PAY_FORM.paymentFor).indexOf("Shop") === 0;',
    '  document.querySelectorAll("input[name=\\"pay-for-radio\\"]").forEach(function(r){ r.disabled = isShop; });',
    '  var today = new Date();',
    '  var ymd = today.getFullYear() + "-" + String(today.getMonth()+1).padStart(2,"0") + "-" + String(today.getDate()).padStart(2,"0");',
    '  document.getElementById("pay-date").value = ymd;',
    '  document.getElementById("pay-shot-name").textContent = "No file chosen";',
    '  var prev = document.getElementById("pay-shot-preview"); if (prev) { prev.style.display="none"; prev.removeAttribute("src"); }',
    '  document.getElementById("pay-msg").className = "pay-msg"; document.getElementById("pay-msg").textContent = "";',
    '  document.getElementById("pay-sub").textContent = isShop',
    '    ? "Shopping payment for " + (RD.name || RD.keNo) + ". Amount can cover one or more orders."',
    '    : "Riding / class fee payment for " + (RD.name || RD.keNo) + ".";',
    '  document.getElementById("pay-qr-wrap").style.display = "none";',
    '  document.getElementById("pay-overlay").classList.add("on");',
    '  google.script.run.withSuccessHandler(function(res){',
    '    if (res && res.success) {',
    '      if (res.phone) document.getElementById("pay-phone").value = res.phone;',
    '      if (res.pan) document.getElementById("pay-pan").value = res.pan;',
    '    }',
    '  }).getPortalPaymentPrefill(RD.keNo, RD.shopToken);',
    '  refreshPortalPaymentQr();',
    '}',
    '',
    'function setPayForRadio(value) {',
    '  var val = String(value || "");',
    '  var radios = document.querySelectorAll("input[name=\\"pay-for-radio\\"]");',
    '  var matched = "Riding Classes";',
    '  radios.forEach(function(r){',
    '    var on = false;',
    '    if (val.indexOf("Shop") === 0) on = String(r.value).indexOf("Shop") === 0;',
    '    else on = String(r.value).indexOf("Riding") === 0;',
    '    r.checked = on;',
    '    if (on) matched = r.value;',
    '  });',
    '  document.getElementById("pay-for").value = matched;',
    '  PAY_FORM.paymentFor = matched;',
    '}',
    '',
    'function onPayForRadioChange() {',
    '  var checked = document.querySelector("input[name=\\"pay-for-radio\\"]:checked");',
    '  var val = checked ? checked.value : "Riding Classes";',
    '  document.getElementById("pay-for").value = val;',
    '  PAY_FORM.paymentFor = val;',
    '}',
    '',
    'function setPayChoice(fieldId, value) {',
    '  var input = document.getElementById(fieldId);',
    '  if (!input) return;',
    '  var val = String(value || "");',
    '  var buttons = document.querySelectorAll("[data-pay-mode]");',
    '  var matched = "";',
    '  buttons.forEach(function(b){',
    '    var v = b.getAttribute("data-pay-mode") || "";',
    '    var on = v === val;',
    '    b.classList.toggle("on", on);',
    '    if (on) matched = v;',
    '  });',
    '  input.value = matched || val;',
    '}',
    '',
    'function onPayChoiceClick(btn) {',
    '  if (!btn || btn.disabled) return;',
    '  if (btn.hasAttribute("data-pay-mode")) setPayChoice("pay-mode", btn.getAttribute("data-pay-mode"));',
    '}',
    '',
    'function closePortalPaymentForm() {',
    '  document.getElementById("pay-overlay").classList.remove("on");',
    '}',
    '',
    'function refreshPortalPaymentQr() {',
    '  if (PAY_QR_TIMER) clearTimeout(PAY_QR_TIMER);',
    '  PAY_QR_TIMER = setTimeout(function(){',
    '    if (!RD) return;',
    '    var amt = Number(document.getElementById("pay-amount").value || 0);',
    '    var wrap = document.getElementById("pay-qr-wrap");',
    '    if (!(amt > 0)) { wrap.style.display = "none"; return; }',
    '    google.script.run.withSuccessHandler(function(res){',
    '      if (!res || !res.success) { wrap.style.display = "none"; return; }',
    '      wrap.style.display = "block";',
    '      document.getElementById("pay-upi").textContent = res.upiId || "";',
    '      var img = document.getElementById("pay-qr"); if (img) img.src = res.qrUrl || "";',
    '    }).getPortalPaymentQr(RD.keNo, RD.shopToken, amt);',
    '  }, 350);',
    '}',
    '',
    'function onPayScreenshotChosen(input) {',
    '  var file = input && input.files && input.files[0];',
    '  PAY_FORM.screenshotData = "";',
    '  PAY_FORM.screenshotName = "";',
    '  var nameEl = document.getElementById("pay-shot-name");',
    '  var prev = document.getElementById("pay-shot-preview");',
    '  if (!file) { if (nameEl) nameEl.textContent = "No file chosen"; if (prev) prev.style.display="none"; return; }',
    '  if (!file.type || file.type.indexOf("image/") !== 0) { toast("Please choose an image screenshot"); input.value=""; return; }',
    '  if (nameEl) nameEl.textContent = file.name;',
    '  var reader = new FileReader();',
    '  reader.onload = function() {',
    '    var img = new Image();',
    '    img.onload = function() {',
    '      var max = 1280, w = img.width, h = img.height;',
    '      if (w > max || h > max) { var s = Math.min(max/w, max/h); w = Math.round(w*s); h = Math.round(h*s); }',
    '      var canvas = document.createElement("canvas"); canvas.width = w; canvas.height = h;',
    '      canvas.getContext("2d").drawImage(img, 0, 0, w, h);',
    '      PAY_FORM.screenshotData = canvas.toDataURL("image/jpeg", 0.82);',
    '      PAY_FORM.screenshotName = file.name;',
    '      if (prev) { prev.src = PAY_FORM.screenshotData; prev.style.display = "block"; }',
    '    };',
    '    img.onerror = function(){ toast("Could not read image"); };',
    '    img.src = reader.result;',
    '  };',
    '  reader.onerror = function(){ toast("Could not read file"); };',
    '  reader.readAsDataURL(file);',
    '}',
    '',
    'function submitPortalPaymentForm() {',
    '  if (!RD) return;',
    '  var msg = document.getElementById("pay-msg");',
    '  var btn = document.getElementById("pay-submit");',
    '  var amount = Number(document.getElementById("pay-amount").value || 0);',
    '  var payDate = document.getElementById("pay-date").value || "";',
    '  var mode = document.getElementById("pay-mode").value || "";',
    '  var paymentFor = document.getElementById("pay-for").value || PAY_FORM.paymentFor;',
    '  var pan = document.getElementById("pay-pan").value || "";',
    '  var txnRef = document.getElementById("pay-txn").value || "";',
    '  var phone = document.getElementById("pay-phone").value || RD.phone || "";',
    '  if (!(amount > 0)) { msg.className="pay-msg err"; msg.textContent="Enter a valid amount."; return; }',
    '  if (!payDate) { msg.className="pay-msg err"; msg.textContent="Select the payment date."; return; }',
    '  if (!mode) { msg.className="pay-msg err"; msg.textContent="Select mode of payment."; return; }',
    '  if (!PAY_FORM.screenshotData) { msg.className="pay-msg err"; msg.textContent="Upload a payment screenshot."; return; }',
    '  if (btn) { btn.disabled = true; btn.textContent = "Submitting..."; }',
    '  msg.className = "pay-msg"; msg.textContent = "Uploading payment details...";',
    '  google.script.run',
    '    .withSuccessHandler(function(res) {',
    '      if (btn) { btn.disabled = false; btn.textContent = "Submit payment"; }',
    '      if (!res || !res.success) { msg.className="pay-msg err"; msg.textContent = (res && res.message) || "Payment failed."; return; }',
    '      msg.className = "pay-msg ok"; msg.textContent = res.message || "Payment submitted.";',
    '      toast("Payment submitted");',
    '      SHOP_ORDERS_LOADED = false;',
    '      google.script.run.withSuccessHandler(function(d){',
    '        if (d && d.found && !d.multiProfile) { RD = d; }',
    '        closePortalPaymentForm();',
    '        if (document.getElementById("tc-orders") && document.getElementById("tc-orders").classList.contains("on")) loadShopOrders(true);',
    '        else if (document.getElementById("tc-payments") && document.getElementById("tc-payments").classList.contains("on")) renderPayments();',
    '        else kTab("orders");',
    '      }).getRiderData(RD.keNo);',
    '    })',
    '    .withFailureHandler(function(e) {',
    '      if (btn) { btn.disabled = false; btn.textContent = "Submit payment"; }',
    '      msg.className = "pay-msg err"; msg.textContent = "Error: " + e.message;',
    '    })',
    '    .submitPortalPayment({',
    '      keNo: RD.keNo,',
    '      token: RD.shopToken,',
    '      phone: phone,',
    '      amount: amount,',
    '      payDate: payDate,',
    '      txnRef: txnRef,',
    '      pan: pan,',
    '      mode: mode,',
    '      paymentFor: paymentFor,',
    '      base64Data: PAY_FORM.screenshotData,',
    '      mimeType: "image/jpeg"',
    '    });',
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
    + '<meta name="theme-color" content="#14330f">'
    + '<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">'
    + '<meta http-equiv="Pragma" content="no-cache">'
    + '<meta http-equiv="Expires" content="0">'
    + '<script>window.__APP_BUILD__="' + Date.now() + '";window.__APP_UI_VERSION__=' + JSON.stringify(appUiVersion) + ';</script>'
    + '<link rel="preconnect" href="https://fonts.googleapis.com">'
    + '<link href="https://fonts.googleapis.com/css2?family=Syne:wght@600;700;800&family=Instrument+Sans:wght@400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">'
    + '<title>My Rides &middot; Kings Equestrian · ' + (CONFIG.LOCATION_CITY || 'Hyderabad') + '</title>'
    + '<style>' + css + '</style>'
    + '</head><body>'

    // LOGIN
    + '<div id="login-screen">'
    +   '<div class="login-ring" style="width:320px;height:320px;top:-80px;right:-80px"></div>'
    +   '<div class="login-ring" style="width:200px;height:200px;bottom:60px;left:-60px"></div>'
    +   '<div class="login-card">'
    +     '<div class="brand-block">'
    +       '<div class="brand-icon"><img src="' + logoUrl + '" alt="Kings Equestrian"></div>'
    +       '<div class="brand-title">My Rides</div>'
    +       '<div class="brand-sub">Kings Equestrian · ' + (typeof schoolLocationShort_ === 'function' ? schoolLocationShort_() : 'Hyderabad') + '</div>'
    +     '</div>'
    +     '<div class="fg">'
    +       '<label class="fl" for="inp-id">Phone or KE Number</label>'
    +       '<input type="tel" class="fi" id="inp-id" placeholder="e.g. 9876543210 or KE240101..." maxlength="20">'
    +     '</div>'
    +     '<button class="btn-p" id="btn-login">View My Rides</button>'
    +     '<div class="login-err" id="login-err"></div>'
    +     '<p class="login-note">Enter your registered phone or KE Number.<br>No password needed.</p>'
    +   '</div>'
    + '</div>'

    // Change 6: PROFILE PICKER SCREEN
    + '<div id="profile-picker">'
    +   '<div class="picker-card">'
    +     '<div class="brand-icon" style="margin:0 auto 1rem"><img src="' + logoUrl + '" alt="Kings Equestrian"></div>'
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
    +     '<button class="tb" data-tab="curriculum">&#128218; Curriculum</button>'
    +     '<button class="tb" data-tab="book">+ Book</button>'
    +     '<button class="tb" data-tab="shop">&#128722; Shop</button>'
    +     '<button class="tb" data-tab="orders">&#128230; Orders</button>'
    +     '<button class="tb" data-tab="payments">&#128179; Payments</button>'
    +   '</div>'
    +   '<div id="tc-sessions" class="tc on"></div>'
    +   '<div id="tc-curriculum" class="tc"></div>'
    +   '<div id="tc-book" class="tc"></div>'
    +   '<div id="tc-shop" class="tc"></div>'
    +   '<div id="tc-orders" class="tc"></div>'
    +   '<div id="tc-payments" class="tc"></div>'
    +   '<div class="pfooter">Kings Equestrian Foundation &middot; ' + (typeof schoolLocationShort_ === 'function' ? schoolLocationShort_() : 'Hyderabad') + ' &middot; ' + (CONFIG.CONTACT_PHONE || '+91-9980895533') + '</div>'
    + '</div>'

    + '<div id="toast"></div>'

    + '<div id="pay-overlay" class="pay-overlay" onclick="if(event.target===this)closePortalPaymentForm()">'
    +   '<div class="pay-modal" role="dialog" aria-modal="true">'
    +     '<h3>Make a payment</h3>'
    +     '<div class="pay-sub" id="pay-sub">Submit payment details here. Receipt will be emailed after verification.</div>'
    +     '<div id="pay-qr-wrap" class="pay-qr-box" style="display:none">'
    +       '<div style="font-size:12px;color:#476d59">Scan UPI QR (optional)</div>'
    +       '<img id="pay-qr" alt="UPI QR">'
    +       '<div style="font-size:12px;font-weight:600;color:var(--pine)">UPI: <span id="pay-upi"></span></div>'
    +     '</div>'
    +     '<div class="pay-grid">'
    +       '<div><label class="fl" for="pay-ke">Registration No (KE No.)</label><input class="fi" id="pay-ke" readonly></div>'
    +       '<div><label class="fl" for="pay-phone">Phone number</label><input class="fi" id="pay-phone" inputmode="tel"></div>'
    +       '<div><label class="fl" for="pay-amount">Amount (₹)</label><input class="fi" id="pay-amount" type="number" min="1" step="1" inputmode="decimal"></div>'
    +       '<div><label class="fl" for="pay-date">Payment date</label><input class="fi" id="pay-date" type="date"></div>'
    +       '<div class="pay-shot">'
    +         '<label class="fl" for="pay-shot">Screenshot</label>'
    +         '<input type="file" id="pay-shot" accept="image/*" style="display:none">'
    +         '<button type="button" class="btn-ghost" id="pay-shot-btn" style="width:100%">Choose screenshot</button>'
    +         '<div class="hint" id="pay-shot-name">No file chosen</div>'
    +         '<img class="pay-shot-preview" id="pay-shot-preview" alt="Screenshot preview">'
    +       '</div>'
    +       '<div><label class="fl" for="pay-txn">Transaction reference <span style="font-weight:400;text-transform:none">(optional)</span></label><input class="fi" id="pay-txn" maxlength="80" placeholder="UPI / bank reference"></div>'
    +       '<div><label class="fl" for="pay-pan">PAN / Aadhaar</label><input class="fi" id="pay-pan" maxlength="20" placeholder="Auto-filled if we have it"></div>'
    +       '<div><label class="fl">Mode of payment</label>'
    +         '<input type="hidden" id="pay-mode" value="UPI">'
    +         '<div class="pay-choice-grid three" id="pay-mode-choices">'
    +           '<button type="button" class="pay-choice on" data-pay-mode="UPI"><span class="pc-title">UPI</span><span class="pc-sub">GPay / PhonePe</span></button>'
    +           '<button type="button" class="pay-choice" data-pay-mode="Bank transfer"><span class="pc-title">Bank</span><span class="pc-sub">NEFT / IMPS</span></button>'
    +           '<button type="button" class="pay-choice" data-pay-mode="Cheque"><span class="pc-title">Cheque</span><span class="pc-sub">Physical</span></button>'
    +         '</div></div>'
    +       '<div><label class="fl">Payment for</label>'
    +         '<input type="hidden" id="pay-for" value="Riding Classes">'
    +         '<div class="pay-radio-list" id="pay-for-choices" role="radiogroup" aria-label="Payment for">'
    +           '<label class="pay-radio">'
    +             '<input type="radio" name="pay-for-radio" value="Riding Classes" checked>'
    +             '<span><span class="pc-title">Riding classes</span><span class="pc-sub">Fees &amp; sessions</span></span>'
    +           '</label>'
    +           '<label class="pay-radio">'
    +             '<input type="radio" name="pay-for-radio" value="Shopping Kit / Equipment">'
    +             '<span><span class="pc-title">Shopping kit / equipment</span><span class="pc-sub">Gear &amp; shop orders</span></span>'
    +           '</label>'
    +         '</div></div>'
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
    + 'document.querySelectorAll("[data-pay-mode]").forEach(function(b){'
    +   'b.addEventListener("click", function(){ onPayChoiceClick(b); });'
    + '});'
    + 'document.querySelectorAll("input[name=\\"pay-for-radio\\"]").forEach(function(r){'
    +   'r.addEventListener("change", onPayForRadioChange);'
    + '});'
    + 'var staleBtn=document.getElementById("stale-reload-btn"); if(staleBtn) staleBtn.addEventListener("click", forceFreshAppReload);'
    + '</script>'

    + '</body></html>';

  return html;
}