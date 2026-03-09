// ============================================================
// KINGS EQUESTRIAN - EXTENDED FEATURES MODULE
// File: KingsEquestrian_NewFeatures.gs
//
// Features:
//   1. Calendar event creation after booking (with guest + reminders)
//   2. Daily Operations Attendance Web App (mobile-friendly PWA)
//   3. Daily Admin Summary Email (PDF + Drive storage)
//
// HOW TO USE:
//   - Add this file as a new script file in your Apps Script project
//   - The main CONFIG object is shared from your existing Code.gs file
//   - Run setupNewFeaturesTriggers() ONCE to register all time-based triggers
//   - Run deployAttendanceApp() to get the Web App URL for mobile icon
// ============================================================


// ─────────────────────────────────────────────
// SECTION 1 — CALENDAR EVENT ON BOOKING SAVE
// ─────────────────────────────────────────────

/**
 * Creates a Google Calendar event immediately after a booking form is submitted.
 * Called from onBookingFormSubmit() in your main Code.gs — replace the existing
 * createCalendarEvent() call with this enhanced version.
 *
 * @param {Object} bookingData  - { name, email, phone, services, date, timeSlots,
 *                                  reference, participants }
 * @returns {string|null}  Calendar event ID, or null on failure
 */
function createBookingCalendarEvent(bookingData) {
    try {
        const calendar = CalendarApp.getDefaultCalendar();

        // ── Parse date ──────────────────────────────────────────────
        const date = new Date(bookingData.date);
        if (isNaN(date.getTime())) {
            Logger.log('Invalid booking date: ' + bookingData.date);
            return null;
        }

        // ── Parse first time slot ───────────────────────────────────
        const timeSlots  = String(bookingData.timeSlots || '').split(',');
        const firstSlot  = timeSlots[0].trim();
        const timeParts  = firstSlot.match(/(\d+):(\d+)\s*(AM|PM)/i);

        if (!timeParts) {
            Logger.log('Could not parse time slot: ' + firstSlot);
            return null;
        }

        let hours   = parseInt(timeParts[1]);
        const mins  = parseInt(timeParts[2]);
        const period = timeParts[3].toUpperCase();

        if (period === 'PM' && hours !== 12) hours += 12;
        if (period === 'AM' && hours === 12) hours = 0;

        const startTime = new Date(date);
        startTime.setHours(hours, mins, 0, 0);

        // Each slot = 30 min
        const endTime = new Date(startTime);
        endTime.setMinutes(endTime.getMinutes() + (timeSlots.length * 30));

        // ── Build event description ─────────────────────────────────
        const participants  = bookingData.participants || 1;
        const participantTx = participants > 1 ? ` (${participants} participants)` : '';
        const description   = [
            `📋 Booking Reference : ${bookingData.reference}`,
            `🏇 Service           : ${bookingData.services}`,
            `👥 Participants      : ${participants}`,
            `📞 Phone             : ${bookingData.phone}`,
            `✉️  Email             : ${bookingData.email}`,
            '',
            'Please arrive 15 minutes before your scheduled time.',
            'Wear comfortable clothing and closed-toe shoes.',
            '',
            '─────────────────────────────',
            'Kings Equestrian Foundation',
            'Karnataka, India',
            '+91-9980895533 | info@kingsequestrian.com'
        ].join('\n');

        // ── Create event with guest invite ──────────────────────────
        const event = calendar.createEvent(
            `Kings Equestrian — ${bookingData.name}${participantTx} (${bookingData.reference})`,
            startTime,
            endTime,
            {
                description  : description,
                location     : 'Kings Equestrian Foundation, Karnataka, India',
                guests       : bookingData.email,   // customer added as guest (can see event)
                sendInvites  : false                // NO invite email — calendar handles reminders only
            }
        );

        // ── Reminders ───────────────────────────────────────────────
        event.removeAllReminders();
        event.addEmailReminder(1440);   // 24 hours before  — email
        event.addEmailReminder(60);     //  1 hour  before  — email
        event.addPopupReminder(30);     // 30 min   before  — popup (staff)

        // ── Colour-code by service ──────────────────────────────────
        const serviceKey = String(bookingData.services || '').toLowerCase();
        let colour = CalendarApp.EventColor.CYAN;
        if (serviceKey.includes('trek'))         colour = CalendarApp.EventColor.GREEN;
        else if (serviceKey.includes('photo'))   colour = CalendarApp.EventColor.YELLOW;
        else if (serviceKey.includes('camp'))    colour = CalendarApp.EventColor.ORANGE;
        else if (serviceKey.includes('lesson') ||
                 serviceKey.includes('riding'))  colour = CalendarApp.EventColor.BLUE;
        event.setColor(colour);

        // ── Persist event ID back to Booking Sheet ──────────────────
        saveCalendarEventIdToSheet(bookingData.reference, event.getId());

        Logger.log(`Calendar event created: ${event.getId()} for ${bookingData.reference}`);
        return event.getId();

    } catch (error) {
        Logger.log('Error creating calendar event: ' + error);
        return null;
    }
}

/**
 * Writes the Google Calendar event ID into the Booking sheet
 * (adds a new column "Calendar Event ID" if not already present).
 */
function saveCalendarEventIdToSheet(reference, eventId) {
    try {
        const ss           = SpreadsheetApp.getActiveSpreadsheet();
        const bookingSheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);
        if (!bookingSheet) return;

        // Ensure header exists — find or create "Calendar Event ID" column
        const headers    = bookingSheet.getRange(1, 1, 1, bookingSheet.getLastColumn()).getValues()[0];
        let calColIndex  = headers.indexOf('Calendar Event ID');

        if (calColIndex === -1) {
            calColIndex = bookingSheet.getLastColumn(); // 0-based index
            bookingSheet.getRange(1, calColIndex + 1).setValue('Calendar Event ID');
        }

        // Find the row matching this reference
        const data = bookingSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][CONFIG.BOOKING_COLS.REFERENCE] || '').trim() === String(reference).trim()) {
                bookingSheet.getRange(i + 1, calColIndex + 1).setValue(eventId);
                Logger.log(`Calendar event ID saved at row ${i + 1}`);
                return;
            }
        }
    } catch (err) {
        Logger.log('Could not save calendar event ID: ' + err);
    }
}


// ─────────────────────────────────────────────
// SECTION 2 — DAILY OPERATIONS ATTENDANCE APP
// ─────────────────────────────────────────────

/**
 * doGet() serves the attendance web app.
 * Deploy as Web App → "Anyone" can access (share URL as phone icon).
 * Staff open it, see today's bookings, tap Present / No-Show.
 */
function doGet(e) {
    const template = HtmlService.createTemplate(getAttendanceAppHtml());
    return template
        .evaluate()
        .setTitle('KE Attendance')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Called by the web app (google.script.run) to get today's + tomorrow's bookings.
 * @param {string} dateStr  'YYYY-MM-DD' or 'today' / 'tomorrow'
 * @returns {Array}  Array of booking objects
 */
function getBookingsForDate(dateStr) {
    const ss           = SpreadsheetApp.getActiveSpreadsheet();
    const bookingSheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);
    if (!bookingSheet) return [];

    const tz   = Session.getScriptTimeZone();
    let target;
    if (dateStr === 'today') {
        target = new Date();
    } else if (dateStr === 'tomorrow') {
        target = new Date();
        target.setDate(target.getDate() + 1);
    } else {
        target = new Date(dateStr);
    }

    const targetStr = Utilities.formatDate(target, tz, 'yyyy-MM-dd');

    const paymentSheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
    // Build a set of references that have confirmed payment
    const paidRefs = new Set();
    if (paymentSheet) {
        const pData = paymentSheet.getDataRange().getValues();
        for (let i = 1; i < pData.length; i++) {
            if (String(pData[i][CONFIG.PAYMENT_COLS.RECEIPT_SENT] || '').toLowerCase() === 'yes') {
                paidRefs.add(String(pData[i][CONFIG.PAYMENT_COLS.REGISTRATION_NO] || '').trim());
            }
        }
    }

    const data     = bookingSheet.getDataRange().getValues();
    const headers  = bookingSheet.getRange(1, 1, 1, bookingSheet.getLastColumn()).getValues()[0];

    // Find attendance column index (0-based)
    let attColIndex = headers.indexOf('Attendance');
    if (attColIndex === -1) {
        attColIndex = bookingSheet.getLastColumn();
        bookingSheet.getRange(1, attColIndex + 1).setValue('Attendance');
        headers.push('Attendance');
    }

    let notesColIndex = headers.indexOf('Staff Notes');
    if (notesColIndex === -1) {
        notesColIndex = bookingSheet.getLastColumn();
        bookingSheet.getRange(1, notesColIndex + 1).setValue('Staff Notes');
        headers.push('Staff Notes');
    }

    const results = [];
    for (let i = 1; i < data.length; i++) {
        const row         = data[i];
        const prefDate    = row[CONFIG.BOOKING_COLS.PREFERRED_SERVICE_DATE];
        if (!prefDate) continue;

        let rowDateStr;
        try {
            rowDateStr = Utilities.formatDate(new Date(prefDate), tz, 'yyyy-MM-dd');
        } catch (e) { continue; }

        if (rowDateStr !== targetStr) continue;

        const ref       = String(row[CONFIG.BOOKING_COLS.REFERENCE] || '').trim();
        const isPaid    = paidRefs.has(ref);

        results.push({
            rowIndex    : i + 1,            // 1-based sheet row
            name        : row[CONFIG.BOOKING_COLS.NAME] || '',
            phone       : row[CONFIG.BOOKING_COLS.PHONE_NUMBER] || '',
            services    : row[CONFIG.BOOKING_COLS.OUR_SERVICES] || '',
            timeSlot    : row[CONFIG.BOOKING_COLS.PREFERRED_TIME_SLOT] || '',
            participants: row[CONFIG.BOOKING_COLS.NUMBER_OF_PARTICIPANTS] || 1,
            reference   : ref,
            paymentStatus: isPaid ? 'Paid' : 'Pending',
            attendance  : row[attColIndex] || '',
            notes       : row[notesColIndex] || ''
        });
    }

    // Sort by time slot
    results.sort((a, b) => a.timeSlot.localeCompare(b.timeSlot));
    return results;
}

/**
 * Called by web app to save attendance mark + optional staff note.
 * @param {number} rowIndex   Sheet row (1-based)
 * @param {string} status     'Present' | 'No-Show' | 'Rescheduled' | ''
 * @param {string} note       Optional staff note
 */
function saveAttendance(rowIndex, status, note) {
    try {
        const ss           = SpreadsheetApp.getActiveSpreadsheet();
        const bookingSheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);
        if (!bookingSheet) return { success: false, error: 'Sheet not found' };

        const headers     = bookingSheet.getRange(1, 1, 1, bookingSheet.getLastColumn()).getValues()[0];
        let attColIndex   = headers.indexOf('Attendance');
        if (attColIndex === -1) {
            attColIndex = bookingSheet.getLastColumn();
            bookingSheet.getRange(1, attColIndex + 1).setValue('Attendance');
        }

        let notesColIndex = headers.indexOf('Staff Notes');
        if (notesColIndex === -1) {
            notesColIndex = bookingSheet.getLastColumn();
            bookingSheet.getRange(1, notesColIndex + 1).setValue('Staff Notes');
        }

        // Write attendance
        const attCell = bookingSheet.getRange(rowIndex, attColIndex + 1);
        attCell.setValue(status);

        // Colour-code
        if (status === 'Present') {
            attCell.setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
        } else if (status === 'No-Show') {
            attCell.setBackground('#f8d7da').setFontColor('#721c24').setFontWeight('bold');
        } else if (status === 'Rescheduled') {
            attCell.setBackground('#fff3cd').setFontColor('#856404').setFontWeight('bold');
        } else {
            attCell.setBackground('#ffffff').setFontColor('#333333').setFontWeight('normal');
        }

        // Write notes
        if (note !== undefined && note !== null) {
            bookingSheet.getRange(rowIndex, notesColIndex + 1).setValue(note);
        }

        Logger.log(`Attendance saved: Row ${rowIndex} → ${status}`);
        return { success: true };

    } catch (err) {
        Logger.log('saveAttendance error: ' + err);
        return { success: false, error: err.message };
    }
}

/**
 * Returns the inline HTML for the attendance PWA.
 * Served by doGet() above.
 */
function getAttendanceAppHtml() {
    return `
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="KE Attendance">
<meta name="application-name" content="KE Attendance">
<meta name="theme-color" content="#1f4e3d">
<meta name="msapplication-TileColor" content="#1f4e3d">
<meta name="msapplication-TileImage" content="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg">
<!-- iOS home screen icon -->
<link rel="apple-touch-icon" href="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg">
<link rel="apple-touch-icon" sizes="152x152" href="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg">
<link rel="apple-touch-icon" sizes="167x167" href="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg">
<link rel="apple-touch-icon" sizes="180x180" href="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg">
<!-- Android / Chrome home screen icon via inline manifest blob -->
<script>
(function(){
  const manifest = {
    name: "KE Attendance",
    short_name: "KE Attendance",
    description: "Kings Equestrian Foundation — Staff Attendance",
    start_url: window.location.href,
    display: "standalone",
    background_color: "#1f4e3d",
    theme_color: "#1f4e3d",
    orientation: "portrait",
    icons: [
      { src: "https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg", sizes: "192x192", type: "image/jpeg", purpose: "any maskable" },
      { src: "https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg", sizes: "512x512", type: "image/jpeg", purpose: "any maskable" }
    ]
  };
  const blob = new Blob([JSON.stringify(manifest)], {type:'application/json'});
  const url  = URL.createObjectURL(blob);
  const link = document.createElement('link');
  link.rel   = 'manifest';
  link.href  = url;
  document.head.appendChild(link);
})();
</script>
<title>KE Attendance</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Segoe UI',sans-serif;background:#f0f4f0;min-height:100vh}
  header{background:linear-gradient(135deg,#1f4e3d,#4f9c7a);color:#fff;padding:16px 20px;display:flex;align-items:center;gap:12px;position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,.2)}
  header img{width:42px;height:42px;border-radius:50%;border:2px solid rgba(255,255,255,.4)}
  header h1{font-size:18px;font-weight:700}
  header p{font-size:12px;opacity:.85}
  .tabs{display:flex;background:#fff;border-bottom:2px solid #e0e0e0;position:sticky;top:74px;z-index:99}
  .tab{flex:1;padding:12px 8px;text-align:center;font-size:13px;font-weight:600;color:#666;cursor:pointer;border-bottom:3px solid transparent;transition:all .2s}
  .tab.active{color:#1f4e3d;border-bottom-color:#1f4e3d;background:#f9fffe}
  .content{padding:16px;max-width:700px;margin:0 auto}
  .summary-bar{background:#1f4e3d;color:#fff;border-radius:12px;padding:14px 18px;margin-bottom:16px;display:flex;gap:16px;flex-wrap:wrap}
  .stat{text-align:center;flex:1;min-width:60px}
  .stat-num{font-size:22px;font-weight:700}
  .stat-lbl{font-size:11px;opacity:.8;margin-top:2px}
  .card{background:#fff;border-radius:14px;padding:16px;margin-bottom:12px;box-shadow:0 2px 8px rgba(0,0,0,.07);border-left:4px solid #ccc;transition:box-shadow .2s}
  .card.present{border-left-color:#28a745}
  .card.no-show{border-left-color:#dc3545}
  .card.rescheduled{border-left-color:#ffc107}
  .card:active{box-shadow:0 4px 16px rgba(0,0,0,.15)}
  .card-top{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:8px}
  .name{font-size:16px;font-weight:700;color:#1f4e3d}
  .badge{font-size:11px;padding:3px 9px;border-radius:20px;font-weight:600}
  .badge-paid{background:#d4edda;color:#155724}
  .badge-pending{background:#fff3cd;color:#856404}
  .time-slot{font-size:13px;color:#555;margin-bottom:4px}
  .service{font-size:12px;color:#888;margin-bottom:8px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .ref{font-size:11px;color:#aaa}
  .btn-row{display:flex;gap:8px;margin-top:10px;flex-wrap:wrap}
  .btn{flex:1;min-width:70px;padding:9px 4px;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;transition:all .15s;display:flex;align-items:center;justify-content:center;gap:4px}
  .btn-present{background:#d4edda;color:#155724}
  .btn-present.active,
  .btn-present:active{background:#28a745;color:#fff}
  .btn-noshow{background:#f8d7da;color:#721c24}
  .btn-noshow.active,
  .btn-noshow:active{background:#dc3545;color:#fff}
  .btn-resched{background:#fff3cd;color:#856404}
  .btn-resched.active,
  .btn-resched:active{background:#ffc107;color:#333}
  .btn-clear{background:#f0f0f0;color:#666}
  .btn-clear:active{background:#ccc}
  .note-area{width:100%;margin-top:8px;padding:8px 10px;border:1px solid #ddd;border-radius:8px;font-size:13px;font-family:inherit;resize:vertical;min-height:50px;display:none}
  .note-area.show{display:block}
  .save-note-btn{display:none;margin-top:6px;padding:7px 16px;background:#1f4e3d;color:#fff;border:none;border-radius:8px;font-size:13px;cursor:pointer;font-weight:600}
  .save-note-btn.show{display:inline-block}
  .empty{text-align:center;padding:40px 20px;color:#999}
  .empty-icon{font-size:48px;margin-bottom:12px}
  .loading{text-align:center;padding:40px 20px;color:#1f4e3d;font-size:14px}
  .toast{position:fixed;bottom:24px;left:50%;transform:translateX(-50%);background:#1f4e3d;color:#fff;padding:10px 22px;border-radius:24px;font-size:13px;font-weight:600;z-index:9999;opacity:0;transition:opacity .3s;pointer-events:none;white-space:nowrap}
  .toast.show{opacity:1}
  .refresh-btn{background:rgba(255,255,255,.2);border:none;color:#fff;width:36px;height:36px;border-radius:50%;cursor:pointer;font-size:18px;margin-left:auto;display:flex;align-items:center;justify-content:center}
  .refresh-btn:active{background:rgba(255,255,255,.35)}
  .date-nav{display:flex;align-items:center;gap:10px;margin-bottom:14px;background:#fff;border-radius:12px;padding:10px 14px}
  .date-nav button{background:#f0f4f0;border:none;padding:6px 14px;border-radius:8px;cursor:pointer;font-size:13px;font-weight:600;color:#1f4e3d}
  .date-nav button:active{background:#1f4e3d;color:#fff}
  .date-label{flex:1;text-align:center;font-size:14px;font-weight:700;color:#1f4e3d}
  input[type=date]{border:1px solid #ddd;border-radius:8px;padding:6px 10px;font-size:13px;color:#333}
</style>
</head>
<body>

<header>
  <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg" alt="KE">
  <div>
    <h1>KE Attendance</h1>
    <p id="headerDate">Loading…</p>
  </div>
  <button class="refresh-btn" onclick="loadData()" title="Refresh">⟳</button>
</header>

<div class="tabs">
  <div class="tab active" onclick="switchTab('today',this)">Today</div>
  <div class="tab" onclick="switchTab('tomorrow',this)">Tomorrow</div>
  <div class="tab" onclick="switchTab('custom',this)">Date</div>
</div>

<div class="content">
  <div id="customDatePicker" style="display:none;margin-bottom:10px">
    <input type="date" id="customDate" onchange="loadData()">
  </div>
  <div id="summaryBar" class="summary-bar" style="display:none">
    <div class="stat"><div class="stat-num" id="statTotal">0</div><div class="stat-lbl">Total</div></div>
    <div class="stat"><div class="stat-num" id="statPresent">0</div><div class="stat-lbl">Present</div></div>
    <div class="stat"><div class="stat-num" id="statNoShow">0</div><div class="stat-lbl">No-Show</div></div>
    <div class="stat"><div class="stat-num" id="statPending">0</div><div class="stat-lbl">Unmarked</div></div>
  </div>
  <div id="bookingList"><div class="loading">⏳ Loading bookings…</div></div>
</div>

<div class="toast" id="toast"></div>

<script>
let currentTab = 'today';
let bookings   = [];

function switchTab(tab, el) {
  currentTab = tab;
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  el.classList.add('active');
  document.getElementById('customDatePicker').style.display = (tab === 'custom') ? 'block' : 'none';
  loadData();
}

function getDateParam() {
  if (currentTab === 'custom') {
    const v = document.getElementById('customDate').value;
    return v || 'today';
  }
  return currentTab;
}

function loadData() {
  document.getElementById('bookingList').innerHTML = '<div class="loading">⏳ Loading bookings…</div>';
  document.getElementById('summaryBar').style.display = 'none';
  
  const param = getDateParam();
  google.script.run
    .withSuccessHandler(renderBookings)
    .withFailureHandler(err => {
      document.getElementById('bookingList').innerHTML =
        '<div class="empty"><div class="empty-icon">⚠️</div><p>' + err.message + '</p></div>';
    })
    .getBookingsForDate(param);
}

function renderBookings(data) {
  bookings = data;
  const list = document.getElementById('bookingList');
  
  if (!data || data.length === 0) {
    list.innerHTML = '<div class="empty"><div class="empty-icon">🐴</div><p>No bookings found for this date</p></div>';
    document.getElementById('summaryBar').style.display = 'none';
    return;
  }
  
  // Update summary
  const total   = data.length;
  const present = data.filter(b => b.attendance === 'Present').length;
  const noshow  = data.filter(b => b.attendance === 'No-Show').length;
  const pending = data.filter(b => !b.attendance).length;
  document.getElementById('statTotal').textContent   = total;
  document.getElementById('statPresent').textContent = present;
  document.getElementById('statNoShow').textContent  = noshow;
  document.getElementById('statPending').textContent = pending;
  document.getElementById('summaryBar').style.display = 'flex';
  
  // Build cards
  list.innerHTML = data.map((b, idx) => {
    const attClass = b.attendance === 'Present' ? 'present'
                   : b.attendance === 'No-Show'  ? 'no-show'
                   : b.attendance === 'Rescheduled' ? 'rescheduled' : '';
    return \`
    <div class="card \${attClass}" id="card-\${idx}">
      <div class="card-top">
        <div>
          <div class="name">🧑 \${b.name} \${b.participants > 1 ? '(×'+b.participants+')' : ''}</div>
          <div class="time-slot">🕐 \${b.timeSlot || 'Time TBD'}</div>
          <div class="service" title="\${b.services}">\${b.services}</div>
        </div>
        <span class="badge \${b.paymentStatus === 'Paid' ? 'badge-paid' : 'badge-pending'}">\${b.paymentStatus}</span>
      </div>
      <div class="ref">📋 \${b.reference} &nbsp;|&nbsp; 📞 \${b.phone}</div>
      <div class="btn-row">
        <button class="btn btn-present \${b.attendance==='Present'?'active':''}"
          onclick="markAttendance(\${idx},'Present')">✅ Present</button>
        <button class="btn btn-noshow \${b.attendance==='No-Show'?'active':''}"
          onclick="markAttendance(\${idx},'No-Show')">❌ No-Show</button>
        <button class="btn btn-resched \${b.attendance==='Rescheduled'?'active':''}"
          onclick="markAttendance(\${idx},'Rescheduled')">🔄 Resched</button>
        <button class="btn btn-clear" onclick="markAttendance(\${idx},'')">✕</button>
      </div>
      <textarea class="note-area \${b.notes?'show':''}" id="note-\${idx}"
        placeholder="Staff note (optional)…">\${b.notes||''}</textarea>
      <button class="save-note-btn \${b.notes?'show':''}" onclick="saveNote(\${idx})">💾 Save Note</button>
    </div>\`;
  }).join('');
  
  updateHeaderDate();
}

function markAttendance(idx, status) {
  const b = bookings[idx];
  
  // Optimistic UI update
  b.attendance = status;
  const card = document.getElementById('card-\${idx}');
  if (card) {
    card.className = 'card ' + (status === 'Present' ? 'present' : status === 'No-Show' ? 'no-show' : status === 'Rescheduled' ? 'rescheduled' : '');
  }
  renderBookings(bookings);
  
  // Toggle note area
  const noteEl = document.getElementById('note-' + idx);
  const saveBtn = noteEl ? noteEl.nextElementSibling : null;
  if (status && noteEl) { noteEl.classList.add('show'); if(saveBtn) saveBtn.classList.add('show'); }
  
  // Save to sheet
  google.script.run
    .withSuccessHandler(() => showToast(status ? '✅ Marked: ' + status : '↩️ Cleared'))
    .withFailureHandler(err => showToast('❌ Error: ' + err.message))
    .saveAttendance(b.rowIndex, status, null);
}

function saveNote(idx) {
  const b    = bookings[idx];
  const note = document.getElementById('note-' + idx).value;
  google.script.run
    .withSuccessHandler(() => showToast('💾 Note saved'))
    .withFailureHandler(err => showToast('❌ ' + err.message))
    .saveAttendance(b.rowIndex, b.attendance, note);
}

function showToast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 2200);
}

function updateHeaderDate() {
  const now = new Date();
  const opts = { weekday:'long', day:'numeric', month:'short' };
  document.getElementById('headerDate').textContent =
    now.toLocaleDateString('en-IN', opts);
}

// Initial load
updateHeaderDate();
loadData();
</script>
</body>
</html>`;
}


// ─────────────────────────────────────────────
// SECTION 3 — DAILY ADMIN SUMMARY EMAIL
// ─────────────────────────────────────────────

/**
 * Main function triggered daily (e.g. 7 AM) by a time-based trigger.
 * Fetches today's + tomorrow's schedules, builds PDF + HTML email, sends to admins.
 */
function sendDailyAdminSummary() {
    try {
        Logger.log('=== Starting Daily Admin Summary ===');

        const tz      = Session.getScriptTimeZone();
        const today   = new Date();
        const tomorrow = new Date(); tomorrow.setDate(today.getDate() + 1);

        const todayStr    = Utilities.formatDate(today,    tz, 'yyyy-MM-dd');
        const tomorrowStr = Utilities.formatDate(tomorrow, tz, 'yyyy-MM-dd');
        const todayLabel  = Utilities.formatDate(today,    tz, 'EEEE, dd MMM yyyy');
        const tmrwLabel   = Utilities.formatDate(tomorrow, tz, 'EEEE, dd MMM yyyy');

        const todayBookings    = getBookingsForDate('today');
        const tomorrowBookings = getBookingsForDate('tomorrow');

        Logger.log(`Today: ${todayBookings.length} | Tomorrow: ${tomorrowBookings.length}`);

        // Build PDF
        const pdfBlob = buildDailySummaryPDF(todayBookings, tomorrowBookings, todayLabel, tmrwLabel);

        // Store in Drive
        const driveUrl = storeSummaryInDrive(pdfBlob, today);

        // Build HTML email body
        const htmlBody = buildDailySummaryEmail(
            todayBookings, tomorrowBookings, todayLabel, tmrwLabel, driveUrl
        );

        // Get admin emails
        const adminEmails = getAdminEmails();
        if (!adminEmails || adminEmails.length === 0) {
            Logger.log('No admin emails found — skipping send');
            return;
        }

        const subject = `📋 Kings Equestrian | Daily Schedule — ${todayLabel}`;

        MailApp.sendEmail({
            to          : adminEmails.join(','),
            subject     : subject,
            htmlBody    : htmlBody,
            attachments : [pdfBlob],
            name        : 'Kings Equestrian System'
        });

        Logger.log('Daily summary sent to: ' + adminEmails.join(', '));

    } catch (err) {
        Logger.log('Error in sendDailyAdminSummary: ' + err);
        Logger.log(err.stack);
    }
}

/**
 * Reads admin email addresses from the Mail Info sheet.
 * Looks for rows where type column contains 'Admin' or 'Daily Summary'.
 */
function getAdminEmails() {
    try {
        const ss            = SpreadsheetApp.getActiveSpreadsheet();
        const mailInfoSheet = ss.getSheetByName(CONFIG.SHEETS.MAIL_INFO);
        if (!mailInfoSheet) return [];

        const data   = mailInfoSheet.getDataRange().getValues();
        const emails = [];
        for (let i = 1; i < data.length; i++) {
            const email = data[i][0];
            const type  = String(data[i][1] || '').toLowerCase();
            if (email && (type.includes('admin') || type.includes('daily summary'))) {
                emails.push(email);
            }
        }
        return emails;
    } catch (e) {
        Logger.log('getAdminEmails error: ' + e);
        return [];
    }
}

/**
 * Builds a nicely-formatted HTML email body for the daily summary.
 */
function buildDailySummaryEmail(todayBookings, tomorrowBookings, todayLabel, tmrwLabel, driveUrl) {
    function buildTableRows(bookings) {
        if (!bookings || bookings.length === 0) {
            return '<tr><td colspan="6" style="text-align:center;color:#999;padding:20px;">No bookings scheduled</td></tr>';
        }
        return bookings.map((b, idx) => {
            const attColor = b.attendance === 'Present' ? '#d4edda'
                           : b.attendance === 'No-Show'  ? '#f8d7da'
                           : b.attendance === 'Rescheduled' ? '#fff3cd'
                           : '#f8f9fa';
            const paidBadge = b.paymentStatus === 'Paid'
                ? '<span style="background:#d4edda;color:#155724;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600">Paid</span>'
                : '<span style="background:#fff3cd;color:#856404;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600">Pending</span>';
            return `
            <tr style="background:${idx % 2 === 0 ? '#fff' : '#fafafa'}">
                <td style="padding:10px 12px;font-weight:600;white-space:nowrap">${b.timeSlot || '—'}</td>
                <td style="padding:10px 12px">${b.name}${b.participants > 1 ? ' ×' + b.participants : ''}</td>
                <td style="padding:10px 12px;font-size:12px;color:#555">${b.services}</td>
                <td style="padding:10px 12px">${paidBadge}</td>
                <td style="padding:10px 12px;font-size:12px">${b.phone}</td>
                <td style="padding:10px 12px;background:${attColor};font-weight:600;font-size:12px">${b.attendance || 'Unmarked'}</td>
            </tr>`;
        }).join('');
    }

    const todayStats = buildStatsHtml(todayBookings);

    return `
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="font-family:'Segoe UI',sans-serif;background:#f4f6f4;margin:0;padding:0;color:#333">
<div style="max-width:800px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 12px rgba(0,0,0,.1)">

  <!-- Header -->
  <div style="background:linear-gradient(135deg,#1f4e3d,#4f9c7a);padding:28px 30px;color:#fff;display:flex;align-items:center;gap:18px">
    <img src="https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg"
         style="width:64px;height:64px;border-radius:50%;border:3px solid rgba(255,255,255,.4)" alt="KE">
    <div>
      <h1 style="margin:0;font-size:22px">Daily Schedule Report</h1>
      <p style="margin:6px 0 0;opacity:.9;font-size:14px">Kings Equestrian Foundation — Admin Summary</p>
    </div>
  </div>

  <div style="padding:28px 30px">

    <!-- Today stats -->
    <div style="display:flex;gap:12px;margin-bottom:24px;flex-wrap:wrap">
      ${todayStats}
    </div>

    <!-- Today's Schedule -->
    <h2 style="color:#1f4e3d;border-bottom:3px solid #1f4e3d;padding-bottom:8px;margin-bottom:16px;font-size:18px">
      📅 Today — ${todayLabel}
    </h2>
    <div style="overflow-x:auto;margin-bottom:28px">
      <table style="width:100%;border-collapse:collapse;font-size:13px;min-width:550px">
        <thead>
          <tr style="background:#1f4e3d;color:#fff">
            <th style="padding:10px 12px;text-align:left">Time</th>
            <th style="padding:10px 12px;text-align:left">Rider / Group</th>
            <th style="padding:10px 12px;text-align:left">Service</th>
            <th style="padding:10px 12px;text-align:left">Payment</th>
            <th style="padding:10px 12px;text-align:left">Phone</th>
            <th style="padding:10px 12px;text-align:left">Attendance</th>
          </tr>
        </thead>
        <tbody>${buildTableRows(todayBookings)}</tbody>
      </table>
    </div>

    <!-- Tomorrow's Schedule -->
    <h2 style="color:#2c5f2d;border-bottom:3px solid #2c5f2d;padding-bottom:8px;margin-bottom:16px;font-size:18px">
      📅 Tomorrow — ${tmrwLabel}
    </h2>
    <div style="overflow-x:auto;margin-bottom:28px">
      <table style="width:100%;border-collapse:collapse;font-size:13px;min-width:550px">
        <thead>
          <tr style="background:#2c5f2d;color:#fff">
            <th style="padding:10px 12px;text-align:left">Time</th>
            <th style="padding:10px 12px;text-align:left">Rider / Group</th>
            <th style="padding:10px 12px;text-align:left">Service</th>
            <th style="padding:10px 12px;text-align:left">Payment</th>
            <th style="padding:10px 12px;text-align:left">Phone</th>
            <th style="padding:10px 12px;text-align:left">Attendance</th>
          </tr>
        </thead>
        <tbody>${buildTableRows(tomorrowBookings)}</tbody>
      </table>
    </div>

    <!-- Drive link -->
    ${driveUrl ? `
    <div style="background:#e8f5e9;border-left:4px solid #4caf50;padding:16px;border-radius:6px;margin-top:8px;font-size:13px">
      📁 <strong>PDF also saved to Drive:</strong> 
      <a href="${driveUrl}" style="color:#1f4e3d">${driveUrl}</a>
    </div>` : ''}

    <p style="font-size:12px;color:#999;margin-top:24px">
      This is an automated report generated by Kings Equestrian booking system.
    </p>
  </div>

  <div style="background:#1f4e3d;color:#fff;padding:18px 30px;text-align:center;font-size:12px">
    <strong>Kings Equestrian Foundation</strong> | Karnataka, India<br>
    +91-9980895533 | info@kingsequestrian.com
  </div>
</div>
</body>
</html>`;
}

function buildStatsHtml(bookings) {
    const total    = bookings.length;
    const present  = bookings.filter(b => b.attendance === 'Present').length;
    const noshow   = bookings.filter(b => b.attendance === 'No-Show').length;
    const pending  = bookings.filter(b => !b.attendance).length;
    const revenue  = bookings.filter(b => b.paymentStatus === 'Paid').length;

    function statBox(num, label, bg, color) {
        return `<div style="flex:1;min-width:90px;background:${bg};border-radius:10px;padding:14px 10px;text-align:center">
          <div style="font-size:24px;font-weight:700;color:${color}">${num}</div>
          <div style="font-size:11px;color:#666;margin-top:3px">${label}</div>
        </div>`;
    }

    return statBox(total,   'Total',    '#f0f4f0', '#1f4e3d') +
           statBox(present, 'Present',  '#d4edda', '#155724') +
           statBox(noshow,  'No-Show',  '#f8d7da', '#721c24') +
           statBox(pending, 'Unmarked', '#fff3cd', '#856404') +
           statBox(revenue, 'Paid',     '#d1ecf1', '#0c5460');
}

/**
 * Creates a PDF version of the daily summary using an HTML-to-PDF approach.
 */
function buildDailySummaryPDF(todayBookings, tomorrowBookings, todayLabel, tmrwLabel) {
    function tableRowsPlain(bookings) {
        if (!bookings || bookings.length === 0) return '<tr><td colspan="5" style="text-align:center;color:#999">No bookings</td></tr>';
        return bookings.map(b =>
            `<tr>
               <td>${b.timeSlot || '—'}</td>
               <td>${b.name}${b.participants > 1 ? ' ×' + b.participants : ''}</td>
               <td style="font-size:11px">${b.services}</td>
               <td>${b.paymentStatus}</td>
               <td>${b.phone}</td>
             </tr>`
        ).join('');
    }

    const tz         = Session.getScriptTimeZone();
    const reportDate = Utilities.formatDate(new Date(), tz, 'dd MMM yyyy HH:mm');

    const html = `
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  @page{size:A4 landscape;margin:15mm}
  body{font-family:Arial,sans-serif;font-size:12px;color:#222}
  h1{font-size:18px;color:#1f4e3d;margin:0 0 4px}
  h2{font-size:14px;color:#1f4e3d;margin:18px 0 8px;border-bottom:2px solid #1f4e3d;padding-bottom:4px}
  .meta{font-size:11px;color:#888;margin-bottom:18px}
  table{width:100%;border-collapse:collapse;margin-bottom:20px}
  th{background:#1f4e3d;color:#fff;padding:8px 10px;text-align:left;font-size:11px}
  td{padding:7px 10px;border-bottom:1px solid #e0e0e0;font-size:11px}
  tr:nth-child(even) td{background:#f9f9f9}
  .footer{margin-top:20px;font-size:10px;color:#aaa;text-align:center;border-top:1px solid #eee;padding-top:10px}
  .stats{display:flex;gap:10px;margin-bottom:16px}
  .stat{background:#f0f4f0;border-radius:6px;padding:8px 12px;text-align:center;flex:1}
  .stat-n{font-size:20px;font-weight:bold;color:#1f4e3d}
  .stat-l{font-size:10px;color:#666}
</style>
</head>
<body>
<h1>Kings Equestrian Foundation — Daily Schedule</h1>
<div class="meta">Generated: ${reportDate}</div>

<div class="stats">
  <div class="stat"><div class="stat-n">${todayBookings.length}</div><div class="stat-l">Today Total</div></div>
  <div class="stat"><div class="stat-n">${todayBookings.filter(b=>b.paymentStatus==='Paid').length}</div><div class="stat-l">Paid</div></div>
  <div class="stat"><div class="stat-n">${todayBookings.filter(b=>b.attendance==='Present').length}</div><div class="stat-l">Present</div></div>
  <div class="stat"><div class="stat-n">${tomorrowBookings.length}</div><div class="stat-l">Tomorrow Total</div></div>
</div>

<h2>Today — ${todayLabel}</h2>
<table>
  <thead><tr><th>Time</th><th>Name</th><th>Service</th><th>Payment</th><th>Phone</th></tr></thead>
  <tbody>${tableRowsPlain(todayBookings)}</tbody>
</table>

<h2>Tomorrow — ${tmrwLabel}</h2>
<table>
  <thead><tr><th>Time</th><th>Name</th><th>Service</th><th>Payment</th><th>Phone</th></tr></thead>
  <tbody>${tableRowsPlain(tomorrowBookings)}</tbody>
</table>

<div class="footer">Kings Equestrian Foundation | Karnataka, India | +91-9980895533</div>
</body>
</html>`;

    const tempFile = DriveApp.createFile(
        `daily_summary_temp_${new Date().getTime()}.html`, html, MimeType.HTML
    );
    const pdfBlob = tempFile.getAs('application/pdf');
    const tz2     = Session.getScriptTimeZone();
    const dateStr = Utilities.formatDate(new Date(), tz2, 'yyyy-MM-dd');
    pdfBlob.setName(`KE_Daily_Schedule_${dateStr}.pdf`);
    tempFile.setTrashed(true);
    return pdfBlob;
}

/**
 * Saves the daily summary PDF to Google Drive inside "Kings Farm Receipts/Daily Summaries".
 * @returns {string} Public/Drive URL of the saved file
 */
function storeSummaryInDrive(pdfBlob, date) {
    try {
        const mainFolderName = 'Kings Farm Receipts';
        const subFolderName  = 'Daily Summaries';

        let mainFolder = DriveApp.getFoldersByName(mainFolderName);
        mainFolder = mainFolder.hasNext() ? mainFolder.next() : DriveApp.createFolder(mainFolderName);

        let subFolder  = mainFolder.getFoldersByName(subFolderName);
        subFolder = subFolder.hasNext() ? subFolder.next() : mainFolder.createFolder(subFolderName);

        const tz      = Session.getScriptTimeZone();
        const dateStr = Utilities.formatDate(date, tz, 'yyyy-MM-dd');
        const file    = subFolder.createFile(pdfBlob);
        file.setName(`KE_Daily_Schedule_${dateStr}.pdf`);
        file.setDescription(`Auto-generated daily schedule for ${dateStr}`);

        Logger.log('Summary PDF saved to Drive: ' + file.getUrl());
        return file.getUrl();
    } catch (err) {
        Logger.log('Error saving summary to Drive: ' + err);
        return null;
    }
}


// ─────────────────────────────────────────────
// SECTION 4 — TRIGGER SETUP & WEB APP DEPLOY
// ─────────────────────────────────────────────

/**
 * Run this ONCE from the Apps Script editor to register all new triggers.
 * Existing triggers are preserved — only new ones are added.
 */
function setupNewFeaturesTriggers() {
    const ui = SpreadsheetApp.getUi();

    try {
        const ss              = SpreadsheetApp.getActiveSpreadsheet();
        const existingTriggers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());

        // Daily Admin Summary — 7:00 AM every day
        if (!existingTriggers.includes('sendDailyAdminSummary')) {
            ScriptApp.newTrigger('sendDailyAdminSummary')
                .timeBased()
                .everyDays(1)
                .atHour(7)
                .create();
            Logger.log('Trigger added: sendDailyAdminSummary @ 7 AM daily');
        }

        ui.alert(
            '✅ New Triggers Registered!',
            '• Daily Admin Summary email — every day at 7:00 AM\n\n' +
            'Note: The Attendance Web App does NOT need a trigger.\n' +
            'Go to Deploy → New Deployment → Web App to publish it.\n' +
            'Share the URL with staff — they can add it as a home-screen icon.',
            ui.ButtonSet.OK
        );

    } catch (err) {
        ui.alert('Error setting up triggers: ' + err.message);
    }
}

/**
 * Manual test — call from editor to preview today's summary without emailing.
 */
function testDailySummaryDryRun() {
    const today    = new Date();
    const tomorrow = new Date(); tomorrow.setDate(today.getDate() + 1);
    const tz       = Session.getScriptTimeZone();

    const todayBookings    = getBookingsForDate('today');
    const tomorrowBookings = getBookingsForDate('tomorrow');

    Logger.log('=== DRY RUN DAILY SUMMARY ===');
    Logger.log('Today bookings    : ' + todayBookings.length);
    Logger.log('Tomorrow bookings : ' + tomorrowBookings.length);
    Logger.log(JSON.stringify(todayBookings, null, 2));
}

/**
 * Manual test — call from editor to send the summary RIGHT NOW to admins.
 */
function testSendDailySummaryNow() {
    sendDailyAdminSummary();
    SpreadsheetApp.getUi().alert('✅ Daily summary sent. Check admin inboxes.');
}

// ─────────────────────────────────────────────
// SECTION 5 — MENU ADDITIONS
// (Merge these items into your existing onOpen() in Code.gs)
// ─────────────────────────────────────────────

/**
 * Call this from within your existing onOpen() in Code.gs by adding:
 *
 *   addExtendedMenuItems(menu);
 *
 * after the existing .addItem() calls, before .addToUi().
 *
 * OR simply add these items directly to your onOpen() menu chain.
 */
function addExtendedMenuItems(menu) {
    menu
        .addSeparator()
        .addItem('📅 Send Daily Summary Now',    'testSendDailySummaryNow')
        .addItem('🧪 Test Summary (Dry Run)',     'testDailySummaryDryRun')
        .addItem('⚙️  Setup New Features Triggers', 'setupNewFeaturesTriggers');
}