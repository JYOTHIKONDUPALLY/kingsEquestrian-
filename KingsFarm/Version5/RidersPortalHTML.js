// ============================================================
// KINGS EQUESTRIAN — NEW SYSTEM
// File: 8_RiderPortalHTML.gs
//
// ARCHITECTURE NOTE:
//   All JavaScript is stored as an array of plain GAS strings
//   (one JS statement per array element) then joined with \n.
//   This means:
//     - No \uXXXX escapes inside GAS strings (use HTML entities
//       in the HTML part, plain chars in the JS part)
//     - No inline onclick="fn('arg')" -- all handlers use
//       data-* attributes read inside the function
//     - No nested quote hell
// ============================================================

function getRiderPortalHtml() {
  var paymentLink = CONFIG.PAYMENT_FORM_LINK || '#';
  var servicesList = [];
  try { servicesList = getServicesList(); } catch(e) { Logger.log('getServicesList: ' + e); }

  // Build safe JSON for services - sanitise any quotes in names
  var safeServices = servicesList.map(function(s) {
    return {
      name : String(s.name  || '').replace(/"/g, '&quot;'),
      price: Number(s.price || 0),
      type : String(s.type  || 'Regular').replace(/"/g, '&quot;'),
      pax  : String(s.type  || '').toLowerCase() === 'group'
    };
  });
  var servicesJson = JSON.stringify(safeServices);

  return _portalHTML(paymentLink, servicesJson);
}

// ─────────────────────────────────────────────────────────────
function _portalHTML(payLink, servicesJson) {

  // ── CSS (plain string, no JS inside) ──────────────────────
  var css = ''
    + '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}'
    + ':root{'
    + '--ink:#0a1f16;--forest:#0f3526;--pine:#1a5c3a;--sage:#2e8a5c;'
    + '--fern:#4aab7a;--mint:#8fd4b0;--mist:#c5eada;--dew:#e8f7f0;'
    + '--parchment:#f5f8f5;--white:#ffffff;'
    + '--gold:#b8860b;--gold-pale:#fdf6e3;--gold-border:#e8d48a;'
    + '--red:#b91c1c;--red-pale:#fef2f2;'
    + '--border:rgba(42,120,80,0.13);--border-md:rgba(42,120,80,0.22);'
    + '--shadow-xs:0 1px 3px rgba(10,31,22,.07);--shadow-sm:0 2px 8px rgba(10,31,22,.1);'
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
    + '.info-stats{display:grid;grid-template-columns:repeat(3,1fr);gap:8px}'
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
    + '.btn-pay{display:inline-flex;align-items:center;gap:8px;background:var(--forest);color:var(--mist);padding:12px 24px;text-decoration:none;border-radius:var(--r-sm);font-size:13px;font-weight:600;font-family:"DM Sans",sans-serif}'
    // toast / misc
    + '#toast{position:fixed;bottom:20px;left:50%;transform:translateX(-50%) translateY(70px);background:var(--forest);color:var(--mist);font-size:12px;font-weight:600;padding:10px 20px;border-radius:100px;opacity:0;transition:all .25s;pointer-events:none;white-space:nowrap;z-index:9999;border:1px solid rgba(143,212,176,.25)}'
    + '#toast.show{opacity:1;transform:translateX(-50%) translateY(0)}'
    + '.empty-st{text-align:center;padding:30px 16px;color:#b0c8b8;font-size:13px;line-height:2}'
    + '.pfooter{text-align:center;padding:16px;font-size:10px;color:#b0c8b8;border-top:1px solid var(--border);letter-spacing:.04em}'
    + '.hint{font-size:10px;color:#b0c8b8;margin-top:4px}'
    + '.pax-row{display:none}';

  // ── JavaScript lines (each element = one statement or block) ──
  // Rules:
  //   1. Only double quotes inside strings
  //   2. No \u escapes — use actual Unicode characters (GAS handles UTF-8 fine)
  //   3. No inline onclick with string args — use data-* + addEventListener
  //   4. HTML fragments built with double-quoted attribute values
  var jsLines = [
    // ── data ──────────────────────────────────────────────
    'var RD = null;',
    'var bookMode = "single";',
    'var slotCount = 1;',
    'var recurSelDays = [];',
    'var recurPattern = "";',
    'var SERVICES = ' + servicesJson + ';',
    'var PAYMENT_LINK = ' + JSON.stringify(payLink) + ';',
    'var DAYS = ["Su","Mo","Tu","We","Th","Fr","Sa"];',

    // ── esc ────────────────────────────────────────────────
    'function esc(v) {',
    '  return String(v || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");',
    '}',

    // ── toast ──────────────────────────────────────────────
    'function toast(msg) {',
    '  var t = document.getElementById("toast");',
    '  t.textContent = msg;',
    '  t.classList.add("show");',
    '  setTimeout(function() { t.classList.remove("show"); }, 2800);',
    '}',

    // ── showResult ─────────────────────────────────────────
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
    '  google.script.run',
    '    .withSuccessHandler(function(data) {',
    '      btn.textContent = "View My Rides"; btn.disabled = false;',
    '      if (!data || !data.found) { showErr(data ? data.error : "Not found."); return; }',
    '      RD = data;',
    '      try { localStorage.setItem("KE_ID", id); } catch(e) {}',
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

    'function doLogout() {',
    '  RD = null; slotCount = 1; bookMode = "single";',
    '  try { localStorage.removeItem("KE_ID"); } catch(e) {}',
    '  document.getElementById("dashboard").style.display = "none";',
    '  document.getElementById("login-screen").style.display = "flex";',
    '  document.getElementById("inp-id").value = "";',
    '  document.getElementById("login-err").style.display = "none";',
    '  window.scrollTo(0, 0);',
    '}',

    // auto-restore
    '(function() {',
    '  var saved = ""; try { saved = localStorage.getItem("KE_ID") || ""; } catch(e) {}',
    '  if (!saved) return;',
    '  document.getElementById("inp-id").value = saved;',
    '  var btn = document.getElementById("btn-login");',
    '  btn.textContent = "Restoring..."; btn.disabled = true;',
    '  google.script.run',
    '    .withSuccessHandler(function(data) {',
    '      btn.textContent = "View My Rides"; btn.disabled = false;',
    '      if (!data || !data.found) { try { localStorage.removeItem("KE_ID"); } catch(e) {} return; }',
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
    '  document.getElementById("d-upcoming").textContent = upcoming;',
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
    '  h += "<div class=\\"tab-sub\\">" + sessions.length + " total &middot; " + (RD.classesAttended || 0) + " attended</div>";',
    '  if (upcoming.length) { h += "<div class=\\"sec-div\\">Upcoming</div>"; h += upcoming.map(buildSessCard).join(""); }',
    '  if (past.length)     { h += "<div class=\\"sec-div\\">Past</div>";     h += past.map(buildSessCard).join(""); }',
    '  if (!sessions.length) h += "<div class=\\"empty-st\\">No sessions yet. Use the Book tab!</div>";',
    '  document.getElementById("tc-sessions").innerHTML = h;',
    '  // set min dates on reschedule inputs',
    '  var tmr = new Date(); tmr.setDate(tmr.getDate() + 1);',
    '  var md = tmr.toISOString().split("T")[0];',
    '  document.querySelectorAll(".rs-date").forEach(function(el) { el.min = md; });',
    '  // wire reschedule toggles',
    '  document.querySelectorAll(".rs-trigger").forEach(function(el) {',
    '    el.addEventListener("click", function() {',
    '      var rid = el.getAttribute("data-rid");',
    '      var panel = document.getElementById("rsp-" + rid);',
    '      if (panel) panel.classList.toggle("open");',
    '    });',
    '  });',
    '  // wire reschedule submit buttons',
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
    '    + (s.participants > 1 ? "<div class=\\"sc-tm\\">&#128101; " + s.participants + " participants</div>" : "")',
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
    '        google.script.run.withSuccessHandler(function(d) { if (d && d.found) { RD = d; renderSessions(); } }).getRiderData(RD.keNo);',
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
    '  // set min dates',
    '  var tmr = new Date(); tmr.setDate(tmr.getDate() + 1);',
    '  var md = tmr.toISOString().split("T")[0];',
    '  document.querySelectorAll(".sd-input").forEach(function(el) { el.min = md; });',
    '  // default month',
    '  var now = new Date();',
    '  var mv = now.getFullYear() + "-" + String(now.getMonth() + 1).padStart(2, "0");',
    '  var mi = document.getElementById("recur-month"); if (mi) { mi.value = mv; mi.min = mv; }',
    '  // wire mode buttons',
    '  document.querySelectorAll(".mode-btn").forEach(function(btn) {',
    '    btn.addEventListener("click", function() { switchMode(btn.getAttribute("data-mode")); });',
    '  });',
    '  // wire add-slot',
    '  var as = document.getElementById("btn-add-slot");',
    '  if (as) as.addEventListener("click", addSlot);',
    '  // wire submit',
    '  var bs = document.getElementById("btn-book");',
    '  if (bs) bs.addEventListener("click", submitBookings);',
    '  // wire pattern buttons',
    '  document.querySelectorAll(".pat-btn").forEach(function(btn) {',
    '    btn.addEventListener("click", function() { selectPattern(btn.getAttribute("data-pat")); });',
    '  });',
    '  // wire day chips',
    '  document.querySelectorAll(".day-chip").forEach(function(chip) {',
    '    chip.addEventListener("click", function() { toggleDay(chip); });',
    '  });',
    '  // wire svc change for pax visibility',
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
    '  h += "<input type=\\"date\\" class=\\"fi sd-input\\" id=\\"sdate-" + n + "\\" style=\\"margin-bottom:10px\\">";',
    '  h += "<label class=\\"fl\\">Time Slot</label>";',
    '  h += "<select class=\\"fi\\" id=\\"stime-" + n + "\\" style=\\"margin-bottom:10px\\"><option value=\\"\\">Select time...</option>" + buildTimeOpts() + "</select>";',
    '  h += "<div class=\\"pax-row\\" id=\\"pax-" + n + "\\">";',
    '  h += "<label class=\\"fl\\">Participants</label>";',
    '  h += "<input type=\\"number\\" class=\\"fi\\" id=\\"spax-" + n + "\\" value=\\"1\\" min=\\"1\\" max=\\"20\\" style=\\"margin-bottom:10px\\">";',
    '  h += "<p class=\\"hint\\">Include yourself in the count.</p></div>";',
    '  h += "</div>";',
    '  return h;',
    '}',

    'function wireSlotSvcChange(n) {',
    '  var sel = document.getElementById("svc-" + n);',
    '  if (!sel) return;',
    '  sel.addEventListener("change", function() {',
    '    var idx = parseInt(sel.value, 10);',
    '    var pr  = document.getElementById("pax-" + n);',
    '    if (!pr) return;',
    '    pr.style.display = (!isNaN(idx) && SERVICES[idx] && SERVICES[idx].pax) ? "block" : "none";',
    '  });',
    '}',

    'function addSlot() {',
    '  slotCount++;',
    '  var wrap = document.getElementById("slots-wrap");',
    '  var div = document.createElement("div");',
    '  div.innerHTML = buildSlotCard(slotCount);',
    '  wrap.appendChild(div.firstChild);',
    '  var tmr = new Date(); tmr.setDate(tmr.getDate() + 1);',
    '  var nd = document.getElementById("sdate-" + slotCount);',
    '  if (nd) nd.min = tmr.toISOString().split("T")[0];',
    '  wireSlotSvcChange(slotCount);',
    '  // wire remove button',
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
    '  h += "<div class=\\"pax-row\\" id=\\"recur-pax\\"><label class=\\"fl\\">Participants</label><input type=\\"number\\" class=\\"fi\\" id=\\"rpax\\" value=\\"1\\" min=\\"1\\" max=\\"20\\" style=\\"margin-bottom:10px\\"></div>";',
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
    '  var today = new Date(); today.setHours(0,0,0,0);',
    '  var dates = [], d = new Date(year, month, 1);',
    '  while (d.getMonth() === month) {',
    '    if (tgt.indexOf(d.getDay()) > -1) { var c = new Date(d); if (c > today) dates.push(c); }',
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
    '      var pax = (svc && svc.pax && px) ? parseInt(px.value || 1, 10) : 1;',
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
    '    if (!rdates.length) { toast("No upcoming dates in selected month"); return; }',
    '    var rsvc = SERVICES[parseInt(rSvcIdx, 10)];',
    '    var rpax = (rsvc && rsvc.pax && rpi) ? parseInt(rpi.value || 1, 10) : 1;',
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
    '        google.script.run.withSuccessHandler(function(d) { if (d && d.found) { RD = d; renderDash(); } }).getRiderData(RD.keNo);',
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
    '  h += "<div class=\\"pay-cta\\"><a class=\\"btn-pay\\" href=\\"" + PAYMENT_LINK + "\\" target=\\"_blank\\">+ Make a Payment</a></div>";',
    '  document.getElementById("tc-payments").innerHTML = h;',
    '}',

    // ── time options ───────────────────────────────────────
    'function buildTimeOpts() {',
    '  function p(n) { return (n < 10 ? "0" : "") + n; }',
    '  function grp(label, sh, sm, eh) {',
    '    var html = "<optgroup label=\\"" + label + "\\">";',
    '    var cur = sh * 60 + sm, end = eh * 60;',
    '    while (cur + 30 <= end) {',
    '      var s1 = Math.floor(cur / 60), m1 = cur % 60;',
    '      var s2 = Math.floor((cur + 30) / 60), m2 = (cur + 30) % 60;',
    '      html += "<option>" + p(s1) + ":" + p(m1) + " - " + p(s2) + ":" + p(m2) + "</option>";',
    '      cur += 30;',
    '    }',
    '    return html + "</optgroup>";',
    '  }',
    '  return grp("Morning", 6, 30, 12) + grp("Afternoon", 14, 30, 19);',
    '}'
  ];

  var js = jsLines.join('\n');

  // ── HTML skeleton ────────────────────────────────────────
  var html = '<!DOCTYPE html>'
    + '<html lang="en"><head>'
    + '<meta charset="UTF-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">'
    + '<meta name="apple-mobile-web-app-capable" content="yes">'
    + '<meta name="theme-color" content="#0a1f16">'
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
    +       '<input type="tel" class="fi" id="inp-id" placeholder="e.g. 9876543210 or KER1001" maxlength="20">'
    +     '</div>'
    +     '<button class="btn-p" id="btn-login">View My Rides</button>'
    +     '<div class="login-err" id="login-err"></div>'
    +     '<p class="login-note">Enter your registered phone or KE Number.<br>No password needed.</p>'
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
    +   '<div class="info-banner">'
    +     '<div class="info-ke" id="d-keno"></div>'
    +     '<div class="info-svc" id="d-svc"></div>'
    +     '<div class="info-stats">'
    +       '<div class="istat"><div class="istat-val" id="d-attended">0</div><div class="istat-lbl">Attended</div></div>'
    +       '<div class="istat"><div class="istat-val" id="d-upcoming">0</div><div class="istat-lbl">Upcoming</div></div>'
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
    +   '<div class="pfooter">Kings Equestrian Foundation &middot; Karnataka &middot; +91-9980895533</div>'
    + '</div>'

    + '<div id="toast"></div>'

    // SCRIPT — completely clean, no inline handlers
    + '<script>' + js + '</script>'

    // Wire static event listeners after DOM is ready
    + '<script>'
    + 'document.getElementById("inp-id").addEventListener("keydown", function(e) { if (e.key === "Enter") doLogin(); });'
    + 'document.getElementById("btn-login").addEventListener("click", doLogin);'
    + 'document.getElementById("btn-logout").addEventListener("click", doLogout);'
    + 'document.querySelectorAll(".tb").forEach(function(btn) {'
    +   'btn.addEventListener("click", function() { kTab(btn.getAttribute("data-tab")); });'
    + '});'
    + '</script>'

    + '</body></html>';

  return html;
}