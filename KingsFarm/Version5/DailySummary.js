// ============================================================
// KINGS EQUESTRIAN — NEW SYSTEM
// File: 5_DailySummary.gs
// Daily 7 AM admin summary email + PDF
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

    Logger.log('Today: ' + todaySessions.length + ' | Tomorrow: ' + tomorrowSessions.length + ' | New: ' + newBookings.length);

    const pdfBlob  = _buildSummaryPDF(todaySessions, tomorrowSessions, newBookings, todayLbl, tmrwLbl);
    const driveUrl = _storeSummaryInDrive(pdfBlob, today);

    const htmlBody = _buildSummaryEmail(todaySessions, tomorrowSessions, newBookings, todayLbl, tmrwLbl, driveUrl);

    const adminEmails = getAdminEmails();
    if (!adminEmails.length) { Logger.log('No admin emails — skipping send'); return; }

    GmailApp.sendEmail(
      adminEmails.join(','),
      'KE Daily Schedule: ' + todayLbl,
      '',
      {
        htmlBody    : htmlBody,
        attachments : [pdfBlob],
        name        : 'Kings Equestrian System'
      }
    );

    Logger.log('Daily summary sent to: ' + adminEmails.join(', '));
  } catch (err) {
    Logger.log('sendDailyAdminSummary ERROR: ' + err + '\n' + err.stack);
  }
}

function testSendDailySummaryNow() {
  sendDailyAdminSummary();
  SpreadsheetApp.getUi().alert('✅ Daily summary sent. Check admin inboxes.');
}

function testDailySummaryDryRun() {
  const today    = getSessionsForDate('today');
  const tomorrow = getSessionsForDate('tomorrow');
  const newB     = _getNewBookingsLast24h();
  Logger.log('DRY RUN: today=' + today.length + ', tomorrow=' + tomorrow.length + ', newBookings=' + newB.length);
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
        time    : fmtDateTime(ts)
      });
    }
    return out;
  } catch (e) {
    Logger.log('_getNewBookingsLast24h error: ' + e);
    return [];
  }
}

// ────────────────────────────────────────────────────────────
//  HTML EMAIL BUILDER
// ────────────────────────────────────────────────────────────

function _buildSummaryEmail(todaySess, tomorrowSess, newBookings, todayLbl, tmrwLbl, driveUrl) {
  const present  = todaySess.filter(s => s.attendance === 'Present').length;
  const noShow   = todaySess.filter(s => s.attendance === 'No-Show').length;
  const unmarked = todaySess.filter(s => !s.attendance).length;

  function statBox(num, label, bg, color) {
    return `<div style="flex:1;min-width:90px;background:${bg};border-radius:10px;padding:13px 8px;text-align:center">
      <div style="font-size:26px;font-weight:700;color:${color}">${num}</div>
      <div style="font-size:11px;color:#666;margin-top:2px">${label}</div></div>`;
  }

  function sessionRow(s, idx) {
    const attBg = s.attendance === 'Present' ? '#d4edda'
                : s.attendance === 'No-Show'  ? '#f8d7da'
                : '#f8f9fa';
    return `<tr style="background:${idx%2===0?'#fff':'#fafafa'}">
      <td style="padding:9px 11px;font-weight:600">${s.timeSlot || '—'}</td>
      <td style="padding:9px 11px">${s.name}${s.participants > 1 ? ' ×'+s.participants : ''}</td>
      <td style="padding:9px 11px;font-size:12px;color:#555">${s.service}</td>
      <td style="padding:9px 11px;font-size:12px">${s.phone}</td>
      <td style="padding:9px 11px;background:${attBg};font-weight:600;font-size:12px">${s.attendance || 'Unmarked'}</td>
      <td style="padding:9px 11px;font-size:11px;color:#888">${s.source === 'rider-portal' ? '🌐 Portal' : '📝 Form'}</td>
    </tr>`;
  }

  function newBookingRow(b, idx) {
    return `<tr style="background:${idx%2===0?'#fff':'#fafafa'}">
      <td style="padding:8px 10px">${b.name}</td>
      <td style="padding:8px 10px;font-weight:600;color:#1f4e3d">${b.keNo}</td>
      <td style="padding:8px 10px;font-size:12px">${b.services}</td>
      <td style="padding:8px 10px;font-size:12px">${b.phone}</td>
      <td style="padding:8px 10px;font-size:11px;color:#888">${b.time}</td>
    </tr>`;
  }

  const noSessions = '<tr><td colspan="6" style="text-align:center;color:#999;padding:18px">No sessions scheduled</td></tr>';
  const noNew      = '<tr><td colspan="5" style="text-align:center;color:#999;padding:18px">No new bookings in last 24 hours</td></tr>';

  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="font-family:'Segoe UI',Arial,sans-serif;background:#f4f6f4;margin:0;padding:0;color:#333">
<div style="max-width:820px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 14px rgba(0,0,0,.1)">
  <div style="background:linear-gradient(135deg,#1f4e3d,#4f9c7a);padding:26px 30px;display:flex;align-items:center;gap:16px">
    <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" style="width:60px;height:60px;border-radius:50%;border:3px solid rgba(255,255,255,.35)">
    <div>
      <h1 style="margin:0;color:#fff;font-size:20px">Daily Schedule Report</h1>
      <p style="margin:4px 0 0;color:rgba(255,255,255,.88);font-size:13px">Kings Equestrian Foundation — Admin Summary</p>
    </div>
  </div>
  <div style="padding:26px 30px">

    <!-- Stats -->
    <div style="display:flex;gap:10px;margin-bottom:22px;flex-wrap:wrap">
      ${statBox(todaySess.length, 'Today Total', '#f0f4f0', '#1f4e3d')}
      ${statBox(present,  'Present',  '#d4edda', '#155724')}
      ${statBox(noShow,   'No-Show',  '#f8d7da', '#721c24')}
      ${statBox(unmarked, 'Unmarked', '#fff3cd', '#856404')}
      ${statBox(newBookings.length, 'New (24h)', '#d1ecf1', '#0c5460')}
    </div>

    <!-- New Bookings -->
    <h2 style="color:#0c5460;border-bottom:3px solid #0c5460;padding-bottom:7px;margin-bottom:14px;font-size:17px">🆕 New Bookings (Last 24 Hours)</h2>
    <div style="overflow-x:auto;margin-bottom:26px">
      <table style="width:100%;border-collapse:collapse;font-size:13px;min-width:500px">
        <thead><tr style="background:#0c5460;color:#fff">
          <th style="padding:9px 10px;text-align:left">Name</th>
          <th style="padding:9px 10px;text-align:left">KE No</th>
          <th style="padding:9px 10px;text-align:left">Service</th>
          <th style="padding:9px 10px;text-align:left">Phone</th>
          <th style="padding:9px 10px;text-align:left">Booked At</th>
        </tr></thead>
        <tbody>${newBookings.length ? newBookings.map(newBookingRow).join('') : noNew}</tbody>
      </table>
    </div>

    <!-- Today -->
    <h2 style="color:#1f4e3d;border-bottom:3px solid #1f4e3d;padding-bottom:7px;margin-bottom:14px;font-size:17px">📅 Today — ${todayLbl}</h2>
    <div style="overflow-x:auto;margin-bottom:26px">
      <table style="width:100%;border-collapse:collapse;font-size:13px;min-width:560px">
        <thead><tr style="background:#1f4e3d;color:#fff">
          <th style="padding:9px 11px;text-align:left">Time</th>
          <th style="padding:9px 11px;text-align:left">Rider</th>
          <th style="padding:9px 11px;text-align:left">Service</th>
          <th style="padding:9px 11px;text-align:left">Phone</th>
          <th style="padding:9px 11px;text-align:left">Attendance</th>
          <th style="padding:9px 11px;text-align:left">Source</th>
        </tr></thead>
        <tbody>${todaySess.length ? todaySess.map(sessionRow).join('') : noSessions}</tbody>
      </table>
    </div>

    <!-- Tomorrow -->
    <h2 style="color:#2c5f2d;border-bottom:3px solid #2c5f2d;padding-bottom:7px;margin-bottom:14px;font-size:17px">📅 Tomorrow — ${tmrwLbl}</h2>
    <div style="overflow-x:auto;margin-bottom:16px">
      <table style="width:100%;border-collapse:collapse;font-size:13px;min-width:560px">
        <thead><tr style="background:#2c5f2d;color:#fff">
          <th style="padding:9px 11px;text-align:left">Time</th>
          <th style="padding:9px 11px;text-align:left">Rider</th>
          <th style="padding:9px 11px;text-align:left">Service</th>
          <th style="padding:9px 11px;text-align:left">Phone</th>
          <th style="padding:9px 11px;text-align:left">Attendance</th>
          <th style="padding:9px 11px;text-align:left">Source</th>
        </tr></thead>
        <tbody>${tomorrowSess.length ? tomorrowSess.map((s,i)=>sessionRow(s,i)).join('') : noSessions}</tbody>
      </table>
    </div>

    ${driveUrl ? `<div style="background:#e8f5e9;border-left:4px solid #4caf50;padding:13px;border-radius:6px;font-size:12px">📁 <strong>PDF saved to Drive:</strong> <a href="${driveUrl}" style="color:#1f4e3d">${driveUrl}</a></div>` : ''}
    <p style="font-size:11px;color:#999;margin-top:20px">Auto-generated by Kings Equestrian booking system.</p>
  </div>
  <div style="background:#1f4e3d;color:#fff;padding:16px 30px;text-align:center;font-size:12px">
    <strong>Kings Equestrian Foundation</strong> | Karnataka, India | +91-9980895533 | info@kingsequestrian.com
  </div>
</div></body></html>`;
}

// ────────────────────────────────────────────────────────────
//  PDF BUILDER
// ────────────────────────────────────────────────────────────

function _buildSummaryPDF(todaySess, tomorrowSess, newBookings, todayLbl, tmrwLbl) {
  const tz         = Session.getScriptTimeZone();
  const reportDate = Utilities.formatDate(new Date(), tz, 'dd MMM yyyy HH:mm');

  function tRows(sessions) {
    if (!sessions.length) return '<tr><td colspan="5" style="text-align:center;color:#aaa">No sessions</td></tr>';
    return sessions.map(s => `<tr>
      <td>${s.timeSlot||'—'}</td><td>${s.name}${s.participants>1?' ×'+s.participants:''}</td>
      <td style="font-size:11px">${s.service}</td><td>${s.phone}</td>
      <td>${s.attendance||'Unmarked'}</td></tr>`).join('');
  }

  function nRows(bs) {
    if (!bs.length) return '<tr><td colspan="4" style="text-align:center;color:#aaa">None</td></tr>';
    return bs.map(b => `<tr><td>${b.name}</td><td>${b.keNo}</td><td style="font-size:11px">${b.services}</td><td>${b.phone}</td></tr>`).join('');
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
<h1>Kings Equestrian Foundation — Daily Schedule</h1>
<div class="meta">Generated: ${reportDate}</div>
<div class="stats">
  <div class="s"><div class="sn">${newBookings.length}</div><div class="sl">New (24h)</div></div>
  <div class="s"><div class="sn">${todaySess.length}</div><div class="sl">Today</div></div>
  <div class="s"><div class="sn">${todaySess.filter(s=>s.attendance==='Present').length}</div><div class="sl">Present</div></div>
  <div class="s"><div class="sn">${todaySess.filter(s=>s.attendance==='No-Show').length}</div><div class="sl">No-Show</div></div>
  <div class="s"><div class="sn">${tomorrowSess.length}</div><div class="sl">Tomorrow</div></div>
</div>

<h2>New Bookings (Last 24h)</h2>
<table><thead><tr><th>Name</th><th>KE No</th><th>Service</th><th>Phone</th></tr></thead>
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