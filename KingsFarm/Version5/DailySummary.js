// ============================================================
// KINGS EQUESTRIAN — NEW SYSTEM
// File: 5_DailySummary.gs
// Daily nightly admin summary email + PDF
// Change 2: first admin email = To, all others = CC
//           trigger should run at 21:00 (9 PM)
// ============================================================

function sendDailyAdminSummary() {
  try {
    Logger.log('=== Daily Admin Summary START ===');
    const tz       = Session.getScriptTimeZone();
    const today    = new Date();
    const tomorrow = new Date(); tomorrow.setDate(today.getDate() + 1);

    const todayLbl  = Utilities.formatDate(today,    tz, 'EEEE, dd MMM yyyy');
    const tmrwLbl   = Utilities.formatDate(tomorrow, tz, 'EEEE, dd MMM yyyy');

    const todaySessions    = getSessionsForDate('today');
    const tomorrowSessions = getSessionsForDate('tomorrow');
    const newBookings      = _getNewBookingsLast24h();
    const newRiders        = _getNewRidersToday();

    Logger.log('Today: ' + todaySessions.length + ' | Tomorrow: ' + tomorrowSessions.length + ' | New bookings: ' + newBookings.length + ' | New riders: ' + newRiders.length);

   const pdfBlob  = _buildSummaryPDF(todaySessions, tomorrowSessions, newBookings, newRiders, todayLbl, tmrwLbl);
    const driveUrl = _storeSummaryInDrive(pdfBlob, today);

    const htmlBody = _buildSummaryEmail(todaySessions, tomorrowSessions, newBookings, newRiders, todayLbl, tmrwLbl, driveUrl);

    // ── Change 2: first admin = To, rest = CC ──────────────
    const adminEmails = getAdminEmails();
    if (!adminEmails.length) { Logger.log('No admin emails — skipping send'); return; }

    const primaryAdmin = adminEmails[0];
    const ccAdmins     = adminEmails.slice(1);

try{
    GmailApp.sendEmail(
      primaryAdmin,
      'KE Nightly Summary: ' + todayLbl,
      '',
      {
        htmlBody    : htmlBody,
        attachments : [pdfBlob],
        cc          : ccAdmins.join(','),
        name        : 'Kings Equestrian System'
      }
    );
      logEmail('NightlySummary', primaryAdmin, ccAdmins.join(','), 'KE Nightly Summary: ' + todayLbl, '', 'Sent', '');
} catch (e) {
  logEmailFailed('NightlySummary', primaryAdmin, ccAdmins.join(','), 'KE Nightly Summary: ' + todayLbl, '', String(e));
}

    Logger.log('Nightly summary sent to: ' + primaryAdmin + (ccAdmins.length ? ' (CC: ' + ccAdmins.join(', ') + ')' : ''));
  } catch (err) {
    Logger.log('sendDailyAdminSummary ERROR: ' + err + '\n' + err.stack);
  }
}

function testSendDailySummaryNow() {
  sendDailyAdminSummary();
  SpreadsheetApp.getUi().alert('✅ Nightly summary sent. Check admin inboxes.');
}

function testDailySummaryDryRun() {
  const today    = getSessionsForDate('today');
  const tomorrow = getSessionsForDate('tomorrow');
  const newB     = _getNewBookingsLast24h();
  const newR     = _getNewRidersToday();
  Logger.log('DRY RUN: today=' + today.length + ', tomorrow=' + tomorrow.length + ', newBookings=' + newB.length + ', newRiders=' + newR.length);
  Logger.log(JSON.stringify(today.slice(0,2), null, 2));
  SpreadsheetApp.getUi().alert('Dry run complete. Check Apps Script logs.');
}

// ────────────────────────────────────────────────────────────
//  NEW BOOKINGS IN LAST 24 HOURS
// ────────────────────────────────────────────────────────────

function _getNewBookingsLast24h() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);
    if (!sheet) return [];
    const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - 1);
    const data   = sheet.getDataRange().getValues();
    const out    = [];
    for (let i = 1; i < data.length; i++) {
      const ts = data[i][CONFIG.BOOKING_COLS.TIMESTAMP];
      if (!ts || new Date(ts) < cutoff) continue;
      out.push({
        name    : data[i][CONFIG.BOOKING_COLS.NAME]     || '',
        keNo    : data[i][CONFIG.BOOKING_COLS.KE_NO]    || '',
        services: data[i][CONFIG.BOOKING_COLS.SERVICES] || '',
        phone   : data[i][CONFIG.BOOKING_COLS.PHONE]    || '',
        time    : fmtDateTime(ts),
         prefDate: data[i][CONFIG.BOOKING_COLS.PREF_DATE] || '',   // ← ADD
  prefTime: data[i][CONFIG.BOOKING_COLS.PREF_TIME] || ''    // ← ADD
      });
    }
    return out;
  } catch (e) {
    Logger.log('_getNewBookingsLast24h error: ' + e);
    return [];
  }
}

// ────────────────────────────────────────────────────────────
//  NEW RIDERS REGISTERED TODAY
// ────────────────────────────────────────────────────────────

function _getNewRidersToday() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
    if (!sheet) return [];
    const tz     = Session.getScriptTimeZone();
    const todayD = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const data   = sheet.getDataRange().getValues();
    const out    = [];
    for (let i = 1; i < data.length; i++) {
      const reg = data[i][CONFIG.RIDER_COLS.REGISTERED];
      if (!reg) continue;
      const regD = Utilities.formatDate(new Date(reg), tz, 'yyyy-MM-dd');
      if (regD === todayD) {
        out.push({
          keNo    : data[i][CONFIG.RIDER_COLS.KE_NO]      || '',
          name    : data[i][CONFIG.RIDER_COLS.NAME]        || '',
          phone   : String(data[i][CONFIG.RIDER_COLS.PHONE] || ''),
          services: data[i][CONFIG.RIDER_COLS.SERVICES]    || '',
          prefDate: data[i][CONFIG.RIDER_COLS.PREF_DATE]    || '',   // ← ADD
  prefTime: data[i][CONFIG.RIDER_COLS.PREF_TIME]    || ''    // ← ADD
        });
      }
    }
    return out;
  } catch (e) {
    Logger.log('_getNewRidersToday error: ' + e);
    return [];
  }
}

// ────────────────────────────────────────────────────────────
//  HTML EMAIL BUILDER
// ────────────────────────────────────────────────────────────

function _buildSummaryEmail(todaySess, tomorrowSess, newBookings, newRiders, todayLbl, tmrwLbl, driveUrl) {
  const present  = todaySess.filter(s => s.attendance === 'Present').length;
  const noShow   = todaySess.filter(s => s.attendance === 'No-Show').length;
  const unmarked = todaySess.filter(s => !s.attendance).length;

  function statBox(num, label, bg, color, icon) {
    return `<div style="flex:1;min-width:90px;background:${bg};border-radius:10px;padding:13px 8px;text-align:center">
      <div style="font-size:18px;margin-bottom:3px">${icon || ''}</div>
      <div style="font-size:26px;font-weight:700;color:${color}">${num}</div>
      <div style="font-size:11px;color:#666;margin-top:2px">${label}</div></div>`;
  }

  function sessionRow(s, idx) {
    const attBg = s.attendance === 'Present' ? '#d4edda' : s.attendance === 'No-Show' ? '#f8d7da' : '#f8f9fa';
    const attTxt = s.attendance || 'Unmarked';
    const slots = _calculate30MinBlocks ? _calculate30MinBlocks(s.timeSlot || '') : 1;
    return `<tr style="background:${idx%2===0?'#fff':'#fafafa'}">
      <td style="padding:9px 11px;font-weight:600">${s.timeSlot || '—'}</td>
      <td style="padding:9px 11px">${s.name}${s.participants > 1 ? ' ×'+s.participants : ''}</td>
      <td style="padding:9px 11px;font-size:12px;color:#555">${s.service}</td>
      <td style="padding:9px 11px;font-size:12px">${s.phone}</td>
      <td style="padding:9px 11px;background:${attBg};font-weight:600;font-size:12px">${attTxt}</td>
      <td style="padding:9px 11px;font-size:11px;color:#888">${slots} × 30min</td>
      <td style="padding:9px 11px;font-size:11px;color:#888">${s.source === 'rider-portal' ? 'Portal' : 'Form'}</td>
    </tr>`;
  }

  function tomorrowRow(s, idx) {
    return `<tr style="background:${idx%2===0?'#fff':'#fafafa'}">
      <td style="padding:9px 11px;font-weight:600">${s.timeSlot || '—'}</td>
      <td style="padding:9px 11px">${s.name}${s.participants > 1 ? ' ×'+s.participants : ''}</td>
      <td style="padding:9px 11px;font-size:12px;color:#555">${s.service}</td>
      <td style="padding:9px 11px;font-size:12px">${s.phone}</td>
      <td style="padding:9px 11px;font-size:11px;color:#888">${s.keNo}</td>
    </tr>`;
  }

  function newBookingRow(b, idx) {
    return `<tr style="background:${idx%2===0?'#fff':'#fafafa'}">
      <td style="padding:8px 10px">${b.name}</td>
      <td style="padding:8px 10px;font-weight:600;color:#1f4e3d">${b.keNo}</td>
      <td style="padding:8px 10px;font-size:12px">${b.services}</td>
      <td style="padding:8px 10px;font-size:12px">${b.phone}</td>
       <td style="padding:8px 10px;font-size:12px;color:#1f4e3d;font-weight:600">${formatPrefDate(b.prefDate)}</td>
    <td style="padding:8px 10px;font-size:12px;color:#555">${formatPrefTime(b.prefTime)}</td>
      <td style="padding:8px 10px;font-size:11px;color:#888">${b.time}</td>
    </tr>`;
  }

 function newRiderRow(r, idx) {
  return `<tr style="background:${idx%2===0?'#fff':'#fafafa'}">
    <td style="padding:8px 10px;font-weight:600;color:#0c5460">${r.keNo}</td>
    <td style="padding:8px 10px">${r.name}</td>
    <td style="padding:8px 10px;font-size:12px">${r.phone}</td>
    <td style="padding:8px 10px;font-size:12px">${r.services}</td>
    <td style="padding:8px 10px;font-size:12px;color:#1f4e3d;font-weight:600">${formatPrefDate(r.prefDate)}</td>
    <td style="padding:8px 10px;font-size:12px;color:#555">${formatPrefTime(r.prefTime)}</td>
  </tr>`;
}

  const noSess  = (cols) => `<tr><td colspan="${cols}" style="text-align:center;color:#999;padding:18px">None</td></tr>`;

  // Calculate total 30-min blocks attended today
  const totalBlocks = todaySess.filter(s => s.attendance === 'Present').reduce((sum, s) => {
    return sum + (_calculate30MinBlocks ? _calculate30MinBlocks(s.timeSlot || '') : 1);
  }, 0);

  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="font-family:'Segoe UI',Arial,sans-serif;background:#f4f6f4;margin:0;padding:0;color:#333">
<div style="max-width:860px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 14px rgba(0,0,0,.1)">
  <div style="background:linear-gradient(135deg,#1f4e3d,#4f9c7a);padding:26px 30px;display:flex;align-items:center;gap:16px">
    <img src="https://drive.google.com/uc?export=view&id=1EAkJ8_EeOVmpX3L1RGLi8b9amX5wuLhb"
   style="width:72px;height:72px;border-radius:50%;border:3px solid #000;margin-bottom:12px"> 
    <div>
      <h1 style="margin:0;color:#fff;font-size:20px"> Nightly Schedule Report</h1>
      <p style="margin:4px 0 0;color:rgba(255,255,255,.88);font-size:13px">Kings Equestrian Foundation — Admin Summary</p>
      <p style="margin:2px 0 0;color:rgba(255,255,255,.7);font-size:11px">${todayLbl}</p>
    </div>
  </div>
  <div style="padding:26px 30px">

    <!-- Stats row 1: Today -->
    <h3 style="color:#1f4e3d;margin:0 0 10px;font-size:14px;text-transform:uppercase;letter-spacing:.06em"> Today's Summary</h3>
    <div style="display:flex;gap:10px;margin-bottom:22px;flex-wrap:wrap">
      ${statBox(todaySess.length, 'Total Today', '#f0f4f0', '#1f4e3d', '')}
      ${statBox(present,          'Present',      '#d4edda', '#155724', '✅')}
      ${statBox(noShow,           'No-Show',      '#f8d7da', '#721c24', '❌')}
      ${statBox(unmarked,         'Unmarked',     '#fff3cd', '#856404', '⏳')}
      ${statBox(totalBlocks,      '30-min Blocks','#e8f5e9', '#2e7d32', '')}
    </div>

    <!-- Stats row 2: Overview -->
    <div style="display:flex;gap:10px;margin-bottom:26px;flex-wrap:wrap">
      ${statBox(newBookings.length, 'New Bookings (24h)', '#d1ecf1', '#0c5460', '')}
      ${statBox(newRiders.length,   'New Riders Today',   '#e8daef', '#6c3483', '')}
      ${statBox(tomorrowSess.length,'Booked Tomorrow',    '#fef9e7', '#7d6608', '')}
    </div>

    <!-- New Riders -->
    ${newRiders.length ? `
    <h2 style="color:#6c3483;border-bottom:3px solid #6c3483;padding-bottom:7px;margin-bottom:14px;font-size:17px"> New Riders Today (${newRiders.length})</h2>
    <div style="overflow-x:auto;margin-bottom:26px">
      <table style="width:100%;border-collapse:collapse;font-size:13px;min-width:420px">
       <thead><tr style="background:#6c3483;color:#fff">
  <th style="padding:9px 10px;text-align:left">KE No</th>
  <th style="padding:9px 10px;text-align:left">Name</th>
  <th style="padding:9px 10px;text-align:left">Phone</th>
  <th style="padding:9px 10px;text-align:left">Service</th>
</tr></thead>
        <tbody>${newRiders.map(newRiderRow).join('')}</tbody>
      </table>
    </div>` : ''}

    <!-- New Bookings -->
    <h2 style="color:#0c5460;border-bottom:3px solid #0c5460;padding-bottom:7px;margin-bottom:14px;font-size:17px"> New Bookings (Last 24 Hours)</h2>
    <div style="overflow-x:auto;margin-bottom:26px">
      <table style="width:100%;border-collapse:collapse;font-size:13px;min-width:500px">
        <thead><tr style="background:#0c5460;color:#fff">
  <th style="padding:9px 10px;text-align:left">Name</th>
  <th style="padding:9px 10px;text-align:left">KE No</th>
  <th style="padding:9px 10px;text-align:left">Service</th>
  <th style="padding:9px 10px;text-align:left">Phone</th>
  <th style="padding:9px 10px;text-align:left">Booked For (Date)</th>
  <th style="padding:9px 10px;text-align:left">Time Slot</th>
  <th style="padding:9px 10px;text-align:left">Booked At</th>
</tr></thead>
        <tbody>${newBookings.length ? newBookings.map(newBookingRow).join('') : noSess(7)}</tbody>
      </table>
    </div>

    <!-- Today -->
    <h2 style="color:#1f4e3d;border-bottom:3px solid #1f4e3d;padding-bottom:7px;margin-bottom:14px;font-size:17px"> Today — ${todayLbl}</h2>
    <div style="overflow-x:auto;margin-bottom:26px">
      <table style="width:100%;border-collapse:collapse;font-size:13px;min-width:600px">
        <thead><tr style="background:#1f4e3d;color:#fff">
          <th style="padding:9px 11px;text-align:left">Time</th>
          <th style="padding:9px 11px;text-align:left">Rider</th>
          <th style="padding:9px 11px;text-align:left">Service</th>
          <th style="padding:9px 11px;text-align:left">Phone</th>
          <th style="padding:9px 11px;text-align:left">Attendance</th>
          <th style="padding:9px 11px;text-align:left">Class Units</th>
          <th style="padding:9px 11px;text-align:left">Source</th>
        </tr></thead>
        <tbody>${todaySess.length ? todaySess.map(sessionRow).join('') : noSess(7)}</tbody>
      </table>
    </div>

    <!-- Tomorrow -->
    <h2 style="color:#2c5f2d;border-bottom:3px solid #2c5f2d;padding-bottom:7px;margin-bottom:14px;font-size:17px"> Tomorrow — ${tmrwLbl} (${tomorrowSess.length} booked)</h2>
    <div style="overflow-x:auto;margin-bottom:16px">
      <table style="width:100%;border-collapse:collapse;font-size:13px;min-width:480px">
        <thead><tr style="background:#2c5f2d;color:#fff">
          <th style="padding:9px 11px;text-align:left">Time</th>
          <th style="padding:9px 11px;text-align:left">Rider</th>
          <th style="padding:9px 11px;text-align:left">Service</th>
          <th style="padding:9px 11px;text-align:left">Phone</th>
          <th style="padding:9px 11px;text-align:left">KE No</th>
        </tr></thead>
        <tbody>${tomorrowSess.length ? tomorrowSess.map((s,i)=>tomorrowRow(s,i)).join('') : noSess(5)}</tbody>
      </table>
    </div>

    ${driveUrl ? `<div style="background:#e8f5e9;border-left:4px solid #4caf50;padding:13px;border-radius:6px;font-size:12px"><strong>PDF saved to Drive:</strong> <a href="${driveUrl}" style="color:#1f4e3d">${driveUrl}</a></div>` : ''}
    <p style="font-size:11px;color:#999;margin-top:20px">Auto-generated nightly by Kings Equestrian booking system.</p>
  </div>
  <div style="background:#1f4e3d;color:#fff;padding:16px 30px;text-align:center;font-size:12px">
    <strong>Kings Equestrian Foundation</strong> | Karnataka, India | +91-9980895533 | info@kingsequestrian.com
  </div>
</div></body></html>`;
}

// ────────────────────────────────────────────────────────────
//  PDF BUILDER
// ────────────────────────────────────────────────────────────
function _buildSummaryPDF(todaySess, tomorrowSess, newBookings, newRiders, todayLbl, tmrwLbl) {
  const tz         = Session.getScriptTimeZone();
  const reportDate = Utilities.formatDate(new Date(), tz, 'dd MMM yyyy HH:mm');

  function tRows(sessions) {
    if (!sessions.length) return '<tr><td colspan="5" style="text-align:center;color:#aaa">No sessions</td></tr>';
    return sessions.map(s => `<tr>
      <td>${s.timeSlot||'—'}</td><td>${s.name}${s.participants>1?' ×'+s.participants:''}</td>
      <td style="font-size:11px">${s.service}</td><td>${s.phone}</td>
      <td>${s.attendance||'Unmarked'}</td></tr>`).join('');
  }


function rRows(riders) {
    if (!riders.length) return '<tr><td colspan="6" style="text-align:center;color:#aaa">None</td></tr>';
    return riders.map(r => `<tr>
      <td style="font-weight:600;color:#6c3483">${r.keNo}</td>
      <td>${r.name}</td>
      <td>${r.phone}</td>
      <td style="font-size:11px">${r.services}</td>
      <td style="font-weight:600;color:#1f4e3d">${formatPrefDate(r.prefDate)}</td>
      <td>${formatPrefTime(r.prefTime)}</td>
    </tr>`).join('');
  }


function nRows(bs) {
    if (!bs.length) return '<tr><td colspan="6" style="text-align:center;color:#aaa">None</td></tr>';
    return bs.map(b => `<tr>
      <td>${b.name}</td>
      <td>${b.keNo}</td>
      <td style="font-size:11px">${b.services}</td>
      <td>${b.phone}</td>
      <td style="font-weight:600;color:#1f4e3d">${formatPrefDate(b.prefDate)}</td>
      <td>${formatPrefTime(b.prefTime)}</td>
    </tr>`).join('');
  }

  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>
@page{size:A4 landscape;margin:14mm}
body{font-family:Arial,sans-serif;font-size:12px;color:#222}
h1{font-size:17px;color:#1f4e3d;margin:0 0 3px}
h2{font-size:13px;color:#1f4e3d;margin:16px 0 7px;border-bottom:2px solid #1f4e3d;padding-bottom:3px}
.meta{font-size:10px;color:#888;margin-bottom:14px}
table{width:100%;border-collapse:collapse;margin-bottom:16px}
th{background:#1f4e3d;color:#fff;padding:7px 9px;text-align:left;font-size:11px}
td{padding:6px 9px;border-bottom:1px solid #e8e8e8;font-size:11px}
tr:nth-child(even) td{background:#f9f9f9}
.stats{display:flex;gap:8px;margin-bottom:14px}
.s{background:#f0f4f0;border-radius:6px;padding:7px 10px;text-align:center;flex:1}
.sn{font-size:19px;font-weight:bold;color:#1f4e3d}
.sl{font-size:10px;color:#666}
footer{margin-top:16px;font-size:10px;color:#aaa;text-align:center;border-top:1px solid #eee;padding-top:8px}
</style></head><body>
<h1>Kings Equestrian Foundation — Nightly Schedule Summary</h1>
<div class="meta">Generated: ${reportDate}</div>
<div class="stats">
  <div class="s"><div class="sn">${newBookings.length}</div><div class="sl">New Bookings (24h)</div></div>
  <div class="s"><div class="sn">${todaySess.length}</div><div class="sl">Today Total</div></div>
  <div class="s"><div class="sn">${todaySess.filter(s=>s.attendance==='Present').length}</div><div class="sl">Present</div></div>
  <div class="s"><div class="sn">${todaySess.filter(s=>s.attendance==='No-Show').length}</div><div class="sl">No-Show</div></div>
  <div class="s"><div class="sn">${todaySess.filter(s=>!s.attendance).length}</div><div class="sl">Unmarked</div></div>
  <div class="s"><div class="sn">${tomorrowSess.length}</div><div class="sl">Tomorrow</div></div>
</div>

<h2>New Riders Today (${newRiders.length})</h2>
<table><thead><tr><th>KE No</th><th>Name</th><th>Phone</th><th>Service</th><th>Booked For</th><th>Time Slot</th></tr></thead>
<tbody>${rRows(newRiders)}</tbody></table>

<h2>New Bookings (Last 24h)</h2>
<table><thead><tr><th>Name</th><th>KE No</th><th>Service</th><th>Phone</th><th>Booked For</th><th>Time Slot</th></tr></thead>
<tbody>${nRows(newBookings)}</tbody></table>

<h2>Today — ${todayLbl}</h2>
<table><thead><tr><th>Time</th><th>Rider</th><th>Service</th><th>Phone</th><th>Attendance</th></tr></thead>
<tbody>${tRows(todaySess)}</tbody></table>

<h2>Tomorrow — ${tmrwLbl}</h2>
<table><thead><tr><th>Time</th><th>Rider</th><th>Service</th><th>Phone</th><th>Attendance</th></tr></thead>
<tbody>${tRows(tomorrowSess)}</tbody></table>

<div class="footer">Kings Equestrian Foundation | Karnataka, India | +91-9980895533</div>
</body></html>`;

  const tmp  = DriveApp.createFile('daily_sum_' + Date.now() + '.html', html, MimeType.HTML);
  const blob = tmp.getAs('application/pdf');
  const ds   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  blob.setName('KE_Daily_Schedule_' + ds + '.pdf');
  tmp.setTrashed(true);
  return blob;
}

// ────────────────────────────────────────────────────────────
//  STORE SUMMARY IN DRIVE
// ────────────────────────────────────────────────────────────

function _storeSummaryInDrive(pdfBlob, date) {
  try {
    let main = DriveApp.getFoldersByName('Kings Equestrian Receipts');
    main     = main.hasNext() ? main.next() : DriveApp.createFolder('Kings Equestrian Receipts');
    let sub  = main.getFoldersByName('Daily Summaries');
    sub      = sub.hasNext()  ? sub.next()  : main.createFolder('Daily Summaries');
    const ds = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const f  = sub.createFile(pdfBlob);
    f.setName('KE_Daily_Schedule_' + ds + '.pdf');
    return f.getUrl();
  } catch (e) {
    Logger.log('_storeSummaryInDrive error: ' + e);
    return null;
  }
}