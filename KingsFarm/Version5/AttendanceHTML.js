// ============================================================
// KINGS EQUESTRIAN — NEW SYSTEM
// File: 7_AttendanceHTML.gs
// Premium Attendance PWA — dark green theme, Space Mono stats
// OPTIMIZED: background preload of riders/transactions
// ============================================================

function getAttendanceAppHtml() {

  var css = [
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}',
    ':root{',
    '  --green-darkest:#061a12;--green-dark:#0d2b1f;--green-mid:#174d36;',
    '  --green-base:#1f6b4a;--green-muted:#2e8a60;--green-soft:#4fae82;',
    '  --green-pale:#9fd4bb;--green-ghost:#d4efe4;--green-tint:#f0faf5;',
    '  --gold:#c9a84c;--gold-light:#f0d98a;--gold-tint:#fdf9ec;',
    '  --red:#c0392b;--red-light:#fde8e6;',
    '  --surface:#f7faf8;--white:#ffffff;',
    '  --border:rgba(31,107,74,0.12);--border-strong:rgba(31,107,74,0.22);',
    '  --text-primary:#0d2b1f;--text-secondary:#3a6652;--text-muted:#7aaa92;',
    '  --shadow-sm:0 1px 4px rgba(13,43,31,0.08);',
    '  --shadow-md:0 4px 16px rgba(13,43,31,0.12);',
    '}',
    'body{font-family:Outfit,sans-serif;background:var(--surface);min-height:100vh;color:var(--text-primary)}',

    // Header
    'header{background:var(--green-dark);color:#fff;padding:0 16px;display:flex;align-items:center;gap:12px;position:sticky;top:0;z-index:200;height:60px;border-bottom:1px solid rgba(255,255,255,0.06)}',
    '.logo-wrap{width:36px;height:36px;border-radius:10px;background:var(--green-mid);border:1.5px solid rgba(201,168,76,0.4);display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0}',
    '.header-text h1{font-size:15px;font-weight:700;letter-spacing:.02em;color:#fff}',
    '.header-text p{font-size:11px;color:var(--green-pale)}',
    '.header-right{margin-left:auto;display:flex;align-items:center;gap:8px}',
    '.hbtn{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.1);color:#fff;width:34px;height:34px;border-radius:9px;cursor:pointer;font-size:15px;display:flex;align-items:center;justify-content:center;transition:background .15s;font-family:Outfit,sans-serif}',
    '.hbtn:hover{background:rgba(255,255,255,.16)}',

    // Date strip
    '.date-strip{background:var(--green-dark);padding:0 16px 12px;display:flex;gap:6px;position:sticky;top:60px;z-index:199}',
    '.dtab{flex:1;padding:7px 8px;text-align:center;font-size:12px;font-weight:500;color:var(--green-pale);cursor:pointer;border-radius:8px;border:1px solid transparent;transition:all .18s;font-family:Outfit,sans-serif}',
    '.dtab.active{background:var(--green-base);color:#fff;border-color:var(--green-muted);font-weight:600}',
    '.dtab:hover:not(.active){background:rgba(255,255,255,.07)}',

    // Nav tabs
    '.nav-tabs{background:var(--white);display:flex;border-bottom:1.5px solid var(--border);position:sticky;top:108px;z-index:198}',
    '.ntab{flex:1;padding:12px 8px;background:none;border:none;border-bottom:2.5px solid transparent;font-family:Outfit,sans-serif;font-size:13px;color:var(--text-muted);cursor:pointer;font-weight:500;transition:all .18s;display:flex;align-items:center;justify-content:center;gap:6px}',
    '.ntab.active{color:var(--green-base);border-bottom-color:var(--green-base);font-weight:600}',

    // Stats
    '.stats-bar{background:linear-gradient(135deg,var(--green-dark) 0%,var(--green-mid) 100%);padding:14px 16px;display:flex;gap:0}',
    '.stat-item{flex:1;text-align:center;position:relative}',
    '.stat-item+.stat-item::before{content:"";position:absolute;left:0;top:20%;height:60%;width:1px;background:rgba(255,255,255,.12)}',
    '.stat-num{font-size:26px;font-weight:700;line-height:1;color:#fff;font-family:"Space Mono",monospace}',
    '.stat-lbl{font-size:10px;color:var(--green-pale);margin-top:3px;text-transform:uppercase;letter-spacing:.07em;font-weight:500}',

    // Content
    '.content{padding:14px;max-width:760px;margin:0 auto}',

    // Session card
    '.session-card{background:var(--white);border-radius:14px;margin-bottom:11px;box-shadow:var(--shadow-sm);border:1px solid var(--border);overflow:hidden}',
    '.sc-accent{height:4px;background:var(--border)}',
    '.session-card.present .sc-accent{background:linear-gradient(90deg,#28a745,#4fae82)}',
    '.session-card.no-show .sc-accent{background:linear-gradient(90deg,#c0392b,#e05a4a)}',
    '.sc-body{padding:14px 15px}',
    '.sc-top{display:flex;align-items:flex-start;justify-content:space-between;gap:10px;margin-bottom:8px}',
    '.sc-name{font-size:16px;font-weight:700;color:var(--text-primary);line-height:1.2}',
    '.sc-badge{font-size:10px;font-weight:700;padding:3px 9px;border-radius:20px;flex-shrink:0;text-transform:uppercase;letter-spacing:.05em}',
    '.badge-pending{background:var(--gold-tint);color:var(--gold);border:1px solid #e8d48a}',
    '.badge-present{background:#e8f9f0;color:#1a7a44;border:1px solid #a3dfbe}',
    '.badge-noshow{background:var(--red-light);color:var(--red);border:1px solid #f0c0bb}',
    '.sc-meta{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:10px}',
    '.sc-chip{display:flex;align-items:center;gap:4px;font-size:12px;color:var(--text-secondary);background:var(--green-tint);padding:3px 9px;border-radius:20px;border:1px solid var(--border)}',
    '.sc-service{font-size:12px;color:var(--text-muted);margin-bottom:10px}',
    '.attended-row{display:flex;align-items:center;gap:6px;font-size:11px;color:var(--green-muted);font-weight:600;margin-bottom:10px}',
    '.att-dot{width:7px;height:7px;border-radius:50%;background:var(--green-soft);flex-shrink:0}',

    // Attendance buttons
    '.att-buttons{display:flex;gap:7px;margin-bottom:10px}',
    '.att-btn{flex:1;padding:9px 10px;border-radius:10px;border:1.5px solid;font-size:12px;font-weight:600;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:5px;transition:all .15s;font-family:Outfit,sans-serif}',
    '.att-btn:disabled{opacity:.4;cursor:default}',
    '.btn-present{background:#f0faf5;border-color:#a3dfbe;color:#1a7a44}',
    '.btn-present.active{background:#28a745;border-color:#28a745;color:#fff}',
    '.btn-noshow{background:var(--red-light);border-color:#f0b8b3;color:var(--red)}',
    '.btn-noshow.active{background:var(--red);border-color:var(--red);color:#fff}',

    // Transaction panel
    '.txn-toggle-btn{width:100%;background:var(--green-tint);border:1px solid var(--border);border-radius:10px;padding:9px 13px;display:flex;align-items:center;justify-content:space-between;cursor:pointer;font-family:Outfit,sans-serif;transition:background .15s}',
    '.txn-toggle-btn:hover{background:var(--green-ghost)}',
    '.txn-toggle-left{display:flex;align-items:center;gap:8px}',
    '.txn-icon-wrap{width:28px;height:28px;background:var(--green-ghost);border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:13px}',
    '.txn-label{font-size:12px;font-weight:600;color:var(--green-base)}',
    '.txn-sublabel{font-size:10px;color:var(--text-muted);margin-top:1px}',
    '.txn-count-badge{font-size:13px;font-weight:700;color:var(--green-base);font-family:"Space Mono",monospace;background:var(--green-ghost);padding:3px 10px;border-radius:20px;border:1px solid var(--border-strong)}',
    '.txn-arrow{font-size:10px;color:var(--text-muted);transition:transform .2s}',
    '.txn-arrow.open{transform:rotate(90deg)}',

    // Transaction drawer
    '.txn-drawer{display:none;border:1px solid var(--border);border-radius:12px;margin-top:8px;overflow:hidden}',
    '.txn-drawer.open{display:block}',
    '.txn-drawer-header{background:var(--green-dark);padding:10px 14px;display:flex;align-items:center;justify-content:space-between}',
    '.txn-drawer-title{font-size:11px;font-weight:700;color:var(--green-pale);text-transform:uppercase;letter-spacing:.08em}',
    '.txn-list{background:var(--white)}',
    '.txn-row{display:flex;align-items:center;padding:11px 14px;border-bottom:1px solid var(--border);gap:10px}',
    '.txn-row:last-child{border-bottom:none}',
    '.txn-info{flex:1}',
    '.txn-desc{font-size:13px;font-weight:600;color:var(--text-primary)}',
    '.txn-sub{font-size:10px;color:var(--text-muted);margin-top:1px;font-family:"Space Mono",monospace}',
    '.txn-right{text-align:right}',
    '.txn-amt{font-size:14px;font-weight:700;color:var(--green-base);font-family:"Space Mono",monospace}',
    '.txn-date{font-size:10px;color:var(--text-muted);margin-top:1px}',
    '.txn-status-dot{width:7px;height:7px;border-radius:50%;background:#28a745;flex-shrink:0}',

    // Riders tab
    '.search-wrap{margin-bottom:12px}',
    '.search-input{width:100%;padding:10px 14px 10px 38px;border:1.5px solid var(--border-strong);border-radius:12px;font-size:13px;font-family:Outfit,sans-serif;outline:none;background:var(--white);color:var(--text-primary);transition:border-color .15s;background-image:url("data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'16\' height=\'16\' viewBox=\'0 0 24 24\' fill=\'none\' stroke=\'%237aaa92\' stroke-width=\'2\'%3E%3Ccircle cx=\'11\' cy=\'11\' r=\'8\'/%3E%3Cpath d=\'m21 21-4.35-4.35\'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:12px center}',
    '.search-input:focus{border-color:var(--green-base)}',
    '.rider-card{background:var(--white);border-radius:14px;padding:14px;margin-bottom:10px;box-shadow:var(--shadow-sm);border:1px solid var(--border)}',
    '.rider-top{display:flex;align-items:center;gap:12px;margin-bottom:10px}',
    '.rider-avatar{width:44px;height:44px;border-radius:12px;background:var(--green-ghost);border:1.5px solid var(--border-strong);display:flex;align-items:center;justify-content:center;font-size:15px;font-weight:700;color:var(--green-base);flex-shrink:0}',
    '.rider-name{font-size:15px;font-weight:700;color:var(--text-primary)}',
    '.rider-phone{font-size:12px;color:var(--text-secondary);margin-top:2px}',
    '.rider-ke{margin-left:auto;font-size:11px;font-weight:700;font-family:"Space Mono",monospace;background:var(--green-dark);color:var(--green-pale);padding:4px 10px;border-radius:8px;flex-shrink:0}',
    '.rider-chips{display:flex;flex-wrap:wrap;gap:5px;margin-bottom:10px}',
    '.r-chip{font-size:11px;padding:3px 9px;border-radius:20px;font-weight:500}',
    '.chip-service{background:var(--green-ghost);color:var(--green-base);border:1px solid var(--border)}',
    '.chip-attended{background:#e8f9f0;color:#1a7a44;border:1px solid #a3dfbe}',
    '.rider-txn{border-top:1px solid var(--border);padding-top:10px;margin-top:4px}',
    '.rider-txn-title{font-size:10px;font-weight:700;color:var(--text-muted);text-transform:uppercase;letter-spacing:.07em;margin-bottom:6px}',
    '.rider-txn-row{display:flex;justify-content:space-between;align-items:center;font-size:12px;padding:5px 0;border-bottom:1px solid var(--green-tint)}',
    '.rider-txn-row:last-child{border-bottom:none}',
    '.rtxn-desc{color:var(--text-secondary)}',
    '.rtxn-amt{font-weight:700;color:var(--green-base);font-family:"Space Mono",monospace}',

    // Custom date
    '.custom-date-wrap{padding:8px 14px;background:var(--white);border-bottom:1px solid var(--border);display:none}',
    'input[type=date]{border:1.5px solid var(--border-strong);border-radius:9px;padding:7px 12px;font-size:12px;background:var(--white);color:var(--text-primary);font-family:Outfit,sans-serif}',

    // Preload indicator
    '.preload-badge{display:inline-flex;align-items:center;gap:5px;font-size:10px;color:var(--green-pale);background:rgba(255,255,255,.08);padding:3px 9px;border-radius:20px;border:1px solid rgba(255,255,255,.1)}',
    '.preload-dot{width:6px;height:6px;border-radius:50%;background:var(--gold);animation:pulse 1.2s ease-in-out infinite}',
    '.preload-dot.done{background:#4fae82;animation:none}',
    '@keyframes pulse{0%,100%{opacity:1}50%{opacity:.3}}',

    // Toast
    '#toast{position:fixed;bottom:20px;left:50%;transform:translateX(-50%) translateY(60px);background:var(--green-dark);color:#fff;font-size:12px;font-weight:600;padding:10px 20px;border-radius:100px;opacity:0;transition:all .25s;pointer-events:none;white-space:nowrap;z-index:9999;border:1px solid var(--green-mid)}',
    '#toast.show{opacity:1;transform:translateX(-50%) translateY(0)}',

    '.empty{text-align:center;padding:40px 20px;color:var(--text-muted);font-size:13px;line-height:1.9}',
    '.empty-icon{font-size:38px;margin-bottom:12px}'
  ].join('\n');

  // ── JavaScript ─────────────────────────────────────────
  var js = ''
    + 'var currentDate="today",currentMain="sessions",sessions=[],allRiders=[],ridersLoaded=false,ridersLoading=false;\n'

    + 'function switchDate(tab,el){'
    +   'currentDate=tab;'
    +   'document.querySelectorAll(".dtab").forEach(function(t){t.classList.remove("active")});'
    +   'el.classList.add("active");'
    +   'document.getElementById("customDateWrap").style.display=(tab==="custom")?"block":"none";'
    +   'if(currentMain==="sessions")loadAttendance();'
    + '}\n'

    + 'function switchMain(tab,el){'
    +   'currentMain=tab;'
    +   'document.querySelectorAll(".ntab").forEach(function(t){t.classList.remove("active")});'
    +   'el.classList.add("active");'
    +   'document.getElementById("tab-sessions").style.display=tab==="sessions"?"block":"none";'
    +   'document.getElementById("tab-transactions").style.display=tab==="transactions"?"block":"none";'
    +   'document.getElementById("tab-riders").style.display=tab==="riders"?"block":"none";'
    +   'document.getElementById("statsBar").style.display=tab==="sessions"?"flex":"none";'
    // If already loaded, render immediately — no waiting
    +   'if(tab==="riders"){'
    +     'if(ridersLoaded){renderRiders(allRiders);}'
    +     'else if(!ridersLoading){loadRiders();}'
    +   '}'
    +   'if(tab==="transactions"){'
    +     'if(ridersLoaded){renderTransactions(allRiders);}'
    +     'else if(!ridersLoading){loadRiders();}'
    +   '}'
    + '}\n'

    + 'function getDateParam(){'
    +   'if(currentDate==="custom"){var v=document.getElementById("customDate").value;return v||"today";}'
    +   'return currentDate;'
    + '}\n'

    + 'function loadAttendance(){'
    +   'document.getElementById("sessionList").innerHTML="<div class=\\"empty\\"><div class=\\"empty-icon\\">&#9203;</div><p>Loading sessions&hellip;</p></div>";'
    +   'document.getElementById("statsBar").style.display="none";'
    +   'google.script.run'
    +     '.withSuccessHandler(renderSessions)'
    +     '.withFailureHandler(function(e){document.getElementById("sessionList").innerHTML="<div class=\\"empty\\"><p>"+e.message+"</p></div>";})'
    +     '.getSessionsForDate_Fast(getDateParam());'
    + '}\n'

    + 'function esc(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}\n'

    + 'function renderSessions(data){'
    +   'sessions=data||[];'
    +   'var list=document.getElementById("sessionList");'
    +   'if(!sessions.length){'
    +     'list.innerHTML="<div class=\\"empty\\"><div class=\\"empty-icon\\">&#128052;</div><p>No sessions scheduled</p></div>";'
    +     'document.getElementById("statsBar").style.display="none";return;'
    +   '}'
    +   'var total=sessions.length;'
    +   'var present=sessions.filter(function(s){return s.attendance==="Present";}).length;'
    +   'var noShow=sessions.filter(function(s){return s.attendance==="No-Show";}).length;'
    +   'var unmarked=sessions.filter(function(s){return !s.attendance;}).length;'
    +   'document.getElementById("sTotal").textContent=total;'
    +   'document.getElementById("sPresent").textContent=present;'
    +   'document.getElementById("sNoShow").textContent=noShow;'
    +   'document.getElementById("sUnmarked").textContent=unmarked;'
    +   'document.getElementById("statsBar").style.display="flex";'
    +   'list.innerHTML=sessions.map(buildSessionCard).join("");'
    +   'updateHeaderDate();'
    + '}\n'

    + 'function buildSessionCard(s,idx){'
    +   'var isPresent=s.attendance==="Present";'
    +   'var isNoShow=s.attendance==="No-Show";'
    +   'var isMarked=isPresent||isNoShow;'
    +   'var cardCls="session-card"+(isPresent?" present":isNoShow?" no-show":"");'
    +   'var badgeCls=isPresent?"badge-present":isNoShow?"badge-noshow":"badge-pending";'
    +   'var badgeTxt=isPresent?"&#10003; Present":isNoShow?"&#10007; No-Show":"Pending";'
    +   'var txnRows="";'
    +   'if(s.payments&&s.payments.length>0){'
    +     'txnRows=s.payments.map(function(p){'
    +       'return "<div class=\\"txn-row\\">"'
    +         '+"<div class=\\"txn-info\\"><div class=\\"txn-desc\\">"+esc(p.receiptNo||"Receipt")+"</div>"'
    +         '+"<div class=\\"txn-sub\\">"+esc(p.txnRef||"")+"</div></div>"'
    +         '+"<div class=\\"txn-right\\"><div class=\\"txn-amt\\">&#8377;"+Number(p.amount).toLocaleString("en-IN")+"</div>"'
    +         '+"<div class=\\"txn-date\\">"+esc(p.payDate||"")+"</div></div>"'
    +         '+"<div class=\\"txn-status-dot\\"></div>"'
    +         '+"</div>";'
    +     '}).join("");'
    +   '}else{'
    +     'txnRows="<div style=\\"padding:12px 14px;font-size:12px;color:var(--text-muted);text-align:center\\">No payment records found</div>";'
    +   '}'
    +   'var dis=isMarked?" disabled":"";'
    +   'var initP=isPresent?" active":"";'
    +   'var initN=isNoShow?" active":"";'
    +   'var btnP="<button class=\\"att-btn btn-present"+initP+"\\""+dis+" data-idx=\\""+idx+"\\" data-status=\\"Present\\" onclick=\\"markAtt(this)\\">&#10003; Present</button>";'
    +   'var btnN="<button class=\\"att-btn btn-noshow"+initN+"\\""+dis+" data-idx=\\""+idx+"\\" data-status=\\"No-Show\\" onclick=\\"markAtt(this)\\">&#10007; No-Show</button>";'
    +   'var txnPanel="<button class=\\"txn-toggle-btn\\" data-idx=\\""+idx+"\\" onclick=\\"toggleTxn(this)\\">"'
    +     '+"<div class=\\"txn-toggle-left\\">"'
    +     '+"<div class=\\"txn-icon-wrap\\">&#128179;</div>"'
    +     '+"<div><div class=\\"txn-label\\">Payments</div>"'
    +     '+"<div class=\\"txn-sublabel\\">"+(s.payments?s.payments.length:0)+" record"+((!s.payments||s.payments.length!==1)?"s":"")+"</div></div>"'
    +     '+"</div>"'
    +     '+"<div style=\\"display:flex;align-items:center;gap:6px\\">"'
    +     '+"<span class=\\"txn-count-badge\\">"+(s.payments?s.payments.length:0)+"</span>"'
    +     '+"<span class=\\"txn-arrow\\" id=\\"txn-arrow-"+idx+"\\">&#9654;</span>"'
    +     '+"</div>"'
    +     '+"</button>"'
    +     '+"<div class=\\"txn-drawer\\" id=\\"txn-drawer-"+idx+"\\">"'
    +     '+"<div class=\\"txn-drawer-header\\"><div class=\\"txn-drawer-title\\">Transaction History</div></div>"'
    +     '+"<div class=\\"txn-list\\">"+txnRows+"</div>"'
    +     '+"</div>";'
    +   'return "<div class=\\""+cardCls+"\\" id=\\"card-"+idx+"\\">"'
    +     '+"<div class=\\"sc-accent\\"></div><div class=\\"sc-body\\">"'
    +     '+"<div class=\\"sc-top\\">"'
    +     '+"<div class=\\"sc-name\\">"+esc(s.name)+(s.participants>1?" &times;"+s.participants:"")+"</div>"'
    +     '+"<span class=\\"sc-badge "+badgeCls+"\\">"+badgeTxt+"</span>"'
    +     '+"</div>"'
    +     '+"<div class=\\"sc-meta\\">"'
    +     '+"<span class=\\"sc-chip\\">&#128336; "+esc(s.timeSlot||"TBD")+"</span>"'
    +     '+"<span class=\\"sc-chip\\">&#128220; "+esc(s.keNo)+"</span>"'
    +     '+"<span class=\\"sc-chip\\">&#128222; "+esc(s.phone)+"</span>"'
    +     '+"</div>"'
    +     '+"<div class=\\"sc-service\\">"+esc(s.service)+"</div>"'
    +     '+"<div class=\\"attended-row\\"><div class=\\"att-dot\\"></div>"+s.classesAttended+" class"+(s.classesAttended!==1?"es":"")+" attended</div>"'
    +     '+"<div class=\\"att-buttons\\">"+btnP+btnN+"</div>"'
    +     '+txnPanel'
    +     '+"</div></div>";'
    + '}\n'

    + 'function markAtt(el){'
    +   'var idx=parseInt(el.getAttribute("data-idx"),10);'
    +   'var status=el.getAttribute("data-status");'
    +   'var s=sessions[idx];'
    +   's.attendance=status;'
    +   'var row=el.closest(".att-buttons");'
    +   'if(row){row.querySelectorAll(".att-btn").forEach(function(b){b.disabled=true;});}'
    +   'el.classList.add("active");'
    +   'var card=document.getElementById("card-"+idx);'
    +   'if(card)card.className="session-card "+(status==="Present"?"present":"no-show");'
    +   'var badge=card?card.querySelector(".sc-badge"):null;'
    +   'if(badge){badge.className="sc-badge "+(status==="Present"?"badge-present":"badge-noshow");badge.innerHTML=status==="Present"?"&#10003; Present":"&#10007; No-Show";}'
    +   'updateStatCounts();'
    +   'showToast(status==="Present"?"&#10003; Marked Present":"&#10007; Marked No-Show");'
    +   'google.script.run'
    +     '.withSuccessHandler(function(res){'
    +       'if(!res||!res.success){showToast("Error: "+(res?res.error:"Unknown"));return;}'
    +       'if(res.emailSent){showToast("📧 Email sent successfully");}'
    +       'else{showToast("⚠️ Attendance updated (no email sent)");}'
    +     '})'
    +     '.withFailureHandler(function(e){showToast("Error: "+e.message);})'
    +     '.saveAttendance(s.rowIndex,status,null);'
    + '}\n'

    + 'function updateStatCounts(){'
    +   'var p=sessions.filter(function(x){return x.attendance==="Present";}).length;'
    +   'var n=sessions.filter(function(x){return x.attendance==="No-Show";}).length;'
    +   'var u=sessions.filter(function(x){return !x.attendance;}).length;'
    +   'document.getElementById("sPresent").textContent=p;'
    +   'document.getElementById("sNoShow").textContent=n;'
    +   'document.getElementById("sUnmarked").textContent=u;'
    + '}\n'

    + 'function toggleTxn(el){'
    +   'var idx=parseInt(el.getAttribute("data-idx"),10);'
    +   'var drawer=document.getElementById("txn-drawer-"+idx);'
    +   'var arrow=document.getElementById("txn-arrow-"+idx);'
    +   'var open=drawer.classList.toggle("open");'
    +   'if(arrow){arrow.innerHTML=open?"&#9660;":"&#9654;";arrow.className="txn-arrow"+(open?" open":"");}'+
    '}\n'

    // ── OPTIMIZED loadRiders ─────────────────────────────
    // ridersLoading flag prevents duplicate in-flight calls.
    // Both txnPageContent and riderList get error messages on failure.
    // After load, if the user is on either tab, render immediately.
    + 'function loadRiders(){'
    +   'if(ridersLoading)return;'
    +   'ridersLoading=true;'
    +   'var riderEl=document.getElementById("riderList");'
    +   'var txnEl=document.getElementById("txnPageContent");'
    +   'if(riderEl)riderEl.innerHTML="<div class=\\"empty\\"><div class=\\"empty-icon\\">&#9203;</div><p>Loading&hellip;</p></div>";'
    +   'if(txnEl)txnEl.innerHTML="<div class=\\"empty\\"><div class=\\"empty-icon\\">&#9203;</div><p>Loading&hellip;</p></div>";'
    +   'google.script.run'
    +     '.withSuccessHandler(function(data){'
    +       'allRiders=data||[];'
    +       'ridersLoaded=true;'
    +       'ridersLoading=false;'
    +       'setPreloadDone();'
    // Always render both — whichever tab is active will show, hidden ones are ready
    +       'renderRiders(allRiders);'
    +       'renderTransactions(allRiders);'
    +     '})'
    +     '.withFailureHandler(function(e){'
    +       'ridersLoading=false;'
    +       'var msg="<div class=\\"empty\\"><p>Error: "+e.message+"</p></div>";'
    +       'if(riderEl)riderEl.innerHTML=msg;'
    +       'if(txnEl)txnEl.innerHTML=msg;'
    +     '})'
    +     '.getAllRidersWithStats_Fast();'
    + '}\n'

    + 'function filterRiders(query){'
    +   'if(!query||!query.trim()){renderRiders(allRiders);return;}'
    +   'var q=query.trim().toLowerCase();'
    +   'renderRiders(allRiders.filter(function(r){'
    +     'return r.name.toLowerCase().indexOf(q)>-1||String(r.phone||"").indexOf(q)>-1||String(r.keNo||"").toLowerCase().indexOf(q)>-1;'
    +   '}));'
    + '}\n'

    + 'function renderRiders(list){'
    +   'var el=document.getElementById("riderList");'
    +   'if(!list||!list.length){el.innerHTML="<div class=\\"empty\\"><div class=\\"empty-icon\\">&#128269;</div><p>No riders match</p></div>";return;}'
    +   'el.innerHTML=list.map(function(r){'
    +     'var initials=r.name.split(" ").map(function(w){return w[0]||"";}).slice(0,2).join("").toUpperCase();'
    +     'var payRows=r.payments&&r.payments.length'
    +       '?r.payments.map(function(p){'
    +           'return "<div class=\\"rider-txn-row\\"><span class=\\"rtxn-desc\\">"+esc(p.txnRef||p.receiptNo||"Payment")+" &middot; "+esc(p.payDate||"")+"</span>"'
    +             '+"<span class=\\"rtxn-amt\\">&#8377;"+Number(p.amount).toLocaleString("en-IN")+"</span></div>";'
    +         '}).join("")'
    +       ':"<div style=\\"font-size:11px;color:var(--text-muted)\\">No payments</div>";'
    +     'var nextChip=r.nextSession'
    +       '?"<span class=\\"r-chip\\" style=\\"background:var(--gold-tint);color:#7a5a0a;border:1px solid #e0c97a\\">&#128197; "+esc(r.nextSession.date)+(r.nextSession.timeSlot?" &middot; "+esc(r.nextSession.timeSlot):"")+"</span>"'
    +       ':"";'
    +     'return "<div class=\\"rider-card\\">"'
    +       '+"<div class=\\"rider-top\\">"'
    +       '+"<div class=\\"rider-avatar\\">"+initials+"</div>"'
    +       '+"<div><div class=\\"rider-name\\">"+esc(r.name)+"</div><div class=\\"rider-phone\\">&#128222; "+esc(r.phone)+"</div></div>"'
    +       '+"<div class=\\"rider-ke\\">"+esc(r.keNo)+"</div>"'
    +       '+"</div>"'
    +       '+"<div class=\\"rider-chips\\">"'
    +       '+"<span class=\\"r-chip chip-service\\">"+esc(r.services)+"</span>"'
    +       '+"<span class=\\"r-chip chip-attended\\">&#10003; "+r.classesAttended+" attended</span>"'
    +       '+nextChip+"</div>"'
    +       '+"<div class=\\"rider-txn\\"><div class=\\"rider-txn-title\\">Payments</div>"+payRows+"</div>"'
    +       '+"</div>";'
    +   '}).join("");'
    + '}\n'

    + 'function renderTransactions(riders){'
    +   'var el=document.getElementById("txnPageContent");'
    +   'if(!el)return;'
    +   'var todayIso=(function(){var d=new Date();var off=d.getTimezoneOffset();return new Date(d.getTime()-off*60000).toISOString().slice(0,10);})();'
    +   'var existingInput=document.getElementById("txnDateFilter");'
    +   'var selectedDate=(existingInput&&existingInput.value)?existingInput.value:todayIso;'
    +   'var allPay=[];'
    +   '(riders||[]).forEach(function(r){(r.payments||[]).forEach(function(p){var filterDate=String(p.filterDate||"").trim();if(!filterDate){var raw=p.payDate||p.paidOn||"";var dt=new Date(raw);if(!isNaN(dt.getTime())){var off=dt.getTimezoneOffset();filterDate=new Date(dt.getTime()-off*60000).toISOString().slice(0,10);}}allPay.push({name:r.name,keNo:r.keNo,amount:p.amount,payDate:p.payDate,txnRef:p.txnRef,receiptNo:p.receiptNo,paidOn:p.paidOn,filterDate:filterDate});});});'
    +   'var filtered=allPay.filter(function(t){return !selectedDate||t.filterDate===selectedDate;});'
    +   'filtered.sort(function(a,b){return String(b.filterDate||"").localeCompare(String(a.filterDate||""))||String(b.paidOn||"").localeCompare(String(a.paidOn||""));});'
    +   'var prettyDate=selectedDate?new Date(selectedDate+"T00:00:00").toLocaleDateString("en-IN",{day:"numeric",month:"short",year:"numeric"}):"All dates";'
    +   'var h="<div style=\\"background:var(--white);border:1px solid var(--border);border-radius:14px;overflow:hidden;box-shadow:var(--shadow-sm)\\">";'
    +   'h+="<div style=\\"background:var(--green-dark);padding:12px 16px;display:flex;align-items:flex-end;justify-content:space-between;gap:12px;flex-wrap:wrap\\">"'
    +     '+"<div><div style=\\"font-size:12px;font-weight:700;color:var(--green-pale);text-transform:uppercase;letter-spacing:.08em\\">Transactions</div>"'
    +     '+"<div style=\\"font-size:11px;color:var(--green-pale);margin-top:2px\\">Showing "+esc(prettyDate)+"</div></div>"'
    +     '+"<div style=\\"display:flex;flex-direction:column;gap:4px;align-items:flex-end\\">"'
    +     '+"<label for=\\"txnDateFilter\\" style=\\"font-size:10px;color:var(--green-pale);text-transform:uppercase;letter-spacing:.06em\\">Payment Date</label>"'
    +     '+"<input type=\\"date\\" id=\\"txnDateFilter\\" value=\\""+selectedDate+"\\" onchange=\\"renderTransactions(allRiders)\\" style=\\"border:1px solid rgba(255,255,255,.18);border-radius:8px;padding:7px 10px;font-size:12px;background:#fff;color:var(--text-primary);min-width:150px\\">"'
    +     '+"</div></div>";'
    +   'h+="<div style=\\"padding:10px 16px;font-size:11px;color:var(--text-secondary);background:var(--green-tint);border-bottom:1px solid var(--border)\\">"+filtered.length+" record"+(filtered.length!==1?"s":"")+" found</div>";'
    +   'if(!filtered.length){h+="<div class=\\"empty\\"><div class=\\"empty-icon\\">&#128179;</div><p>No transactions found for the selected payment date</p></div></div>";el.innerHTML=h;return;}'
    +   'filtered.forEach(function(t){'
    +     'h+="<div style=\\"display:flex;align-items:center;padding:11px 14px;border-bottom:1px solid var(--border);gap:10px\\">"'
    +       '+"<div style=\\"flex:1\\">"'
    +       '+"<div style=\\"font-size:13px;font-weight:600;color:var(--text-primary)\\">"+esc(t.name)+"</div>"'
    +       '+"<div style=\\"font-size:10px;color:var(--text-muted);font-family:Space Mono,monospace;margin-top:1px\\">"+esc(t.txnRef||"")+(t.receiptNo?" &middot; "+esc(t.receiptNo):"")+"</div>"'
    +       '+"<div style=\\"font-size:11px;color:var(--text-secondary);margin-top:2px\\">"+esc(t.keNo)+"</div>"'
    +       '+"</div>"'
    +       '+"<div style=\\"text-align:right\\">"'
    +       '+"<div style=\\"font-size:14px;font-weight:700;color:var(--green-base);font-family:Space Mono,monospace\\">&#8377;"+Number(t.amount).toLocaleString("en-IN")+"</div>"'
    +       '+"<div style=\\"font-size:10px;color:var(--text-muted);margin-top:1px\\">"+esc(t.payDate||t.paidOn||"")+"</div>"'
    +       '+"</div>"'
    +       '+"<div style=\\"width:7px;height:7px;border-radius:50%;background:#28a745;flex-shrink:0\\"></div>"'
    +       '+"</div>";'
    +   '});'
    +   'h+="</div>";'
    +   'el.innerHTML=h;'
    + '}\n'

    + 'function showToast(msg){'
    +   'var t=document.getElementById("toast");t.innerHTML=msg;t.classList.add("show");'
    +   'setTimeout(function(){t.classList.remove("show");},2500);'
    + '}\n'

    + 'function updateHeaderDate(){'
    +   'var d=new Date();'
    +   'document.getElementById("headerDate").textContent=d.toLocaleDateString("en-IN",{weekday:"long",day:"numeric",month:"short",year:"numeric"});'
    + '}\n'

    // Preload indicator helpers
    + 'function setPreloadDone(){'
    +   'var dot=document.getElementById("preloadDot");'
    +   'var lbl=document.getElementById("preloadLbl");'
    +   'if(dot){dot.className="preload-dot done";}'
    +   'if(lbl){lbl.textContent="Ready";}'
    + '}\n'

    // Start sessions load, then kick off riders in background after a short delay
    + 'updateHeaderDate();'
    + 'loadAttendance();'
    + 'setTimeout(function(){if(!ridersLoaded&&!ridersLoading){loadRiders();}},1200);\n';

  // ── Assemble HTML ───────────────────────────────────────
  return '<!DOCTYPE html>\n'
    + '<html lang="en">\n'
    + '<head>\n'
    + '<meta charset="UTF-8">\n'
    + '<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">\n'
    + '<meta name="apple-mobile-web-app-capable" content="yes">\n'
    + '<meta name="theme-color" content="#0d2b1f">\n'
    + '<link rel="preconnect" href="https://fonts.googleapis.com">\n'
    + '<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&family=Space+Mono:wght@400;700&display=swap" rel="stylesheet">\n'
    + '<title>KE Attendance</title>\n'
    + '<style>' + css + '</style>\n'
    + '</head>\n<body>\n'

    + '<header>'
    +   '<div class="logo-wrap">&#128052;</div>'
    +   '<div class="header-text"><h1>KE Attendance</h1><p id="headerDate">Loading&hellip;</p></div>'
    +   '<div class="header-right">'
    // Preload indicator — shows loading dot while riders fetch in background
    +     '<div class="preload-badge">'
    +       '<div class="preload-dot" id="preloadDot"></div>'
    +       '<span id="preloadLbl" style="font-size:10px">Syncing</span>'
    +     '</div>'
    +     '<button class="hbtn" onclick="loadAttendance()" title="Refresh">&#8635;</button>'
    +   '</div>'
    + '</header>\n'

    + '<div class="date-strip">'
    +   '<div class="dtab active" onclick="switchDate(\'today\',this)">Today</div>'
    +   '<div class="dtab" onclick="switchDate(\'tomorrow\',this)">Tomorrow</div>'
    +   '<div class="dtab" onclick="switchDate(\'custom\',this)">&#128197; Custom</div>'
    + '</div>\n'

    + '<div class="custom-date-wrap" id="customDateWrap">'
    +   '<input type="date" id="customDate" onchange="loadAttendance()">'
    + '</div>\n'

    + '<div class="nav-tabs">'
    +   '<button class="ntab active" onclick="switchMain(\'sessions\',this)">&#128203; Sessions</button>'
    +   '<button class="ntab" onclick="switchMain(\'transactions\',this)">&#128179; Transactions</button>'
    +   '<button class="ntab" onclick="switchMain(\'riders\',this)">&#127939; Riders</button>'
    + '</div>\n'

    + '<div id="statsBar" class="stats-bar" style="display:none">'
    +   '<div class="stat-item"><div class="stat-num" id="sTotal">0</div><div class="stat-lbl">Total</div></div>'
    +   '<div class="stat-item"><div class="stat-num" id="sPresent">0</div><div class="stat-lbl">Present</div></div>'
    +   '<div class="stat-item"><div class="stat-num" id="sNoShow">0</div><div class="stat-lbl">No-Show</div></div>'
    +   '<div class="stat-item"><div class="stat-num" id="sUnmarked">0</div><div class="stat-lbl">Unmarked</div></div>'
    + '</div>\n'

    + '<div class="content">'
    +   '<div id="tab-sessions"><div id="sessionList"><div class="empty"><div class="empty-icon">&#9203;</div><p>Loading&hellip;</p></div></div></div>'
    +   '<div id="tab-transactions" style="display:none"><div id="txnPageContent"><div class="empty"><div class="empty-icon">&#128179;</div><p>Loading&hellip;</p></div></div></div>'
    +   '<div id="tab-riders" style="display:none">'
    +     '<div class="search-wrap"><input type="search" class="search-input" id="riderSearch" placeholder="Search name, phone or KE No&hellip;" oninput="filterRiders(this.value)"></div>'
    +     '<div id="riderList"><div class="empty"><div class="empty-icon">&#128014;</div><p>Loading&hellip;</p></div></div>'
    +   '</div>'
    + '</div>\n'

    + '<div id="toast"></div>\n'
    + '<script>\n' + js + '</script>\n'
    + '</body>\n</html>';
}