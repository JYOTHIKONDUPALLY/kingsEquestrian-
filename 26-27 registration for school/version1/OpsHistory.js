// ============================================================
// KINGS EQUESTRIAN — UNIFIED OPS HISTORY + WEEKLY ADMIN REPORT
// ============================================================

function _histYmd_(value) {
  if (typeof _stockYmd_ === 'function') return _stockYmd_(value);
  if (typeof _horseYmd_ === 'function') return _horseYmd_(value);
  if (typeof _groomerYmd_ === 'function') return _groomerYmd_(value);
  if (!value) return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var m = String(value).match(/^(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : '';
}

function _histFromFeed_() {
  var out = [];
  try {
    if (typeof _ensureStockSheets_ !== 'function') return out;
    var sheets = _ensureStockSheets_();
    if (sheets.feedMovements.getLastRow() < 2) return out;
    var mc = CONFIG.FEED_MOVEMENT_COLS;
    var data = sheets.feedMovements.getDataRange().getValues();
    // Name lookup
    var names = {};
    if (sheets.feed.getLastRow() > 1) {
      var fd = sheets.feed.getDataRange().getValues();
      for (var f = 1; f < fd.length; f++) {
        names[String(fd[f][CONFIG.FEED_COLS.ITEM_ID] || '').trim()] = String(fd[f][CONFIG.FEED_COLS.ITEM_NAME] || '').trim();
      }
    }
    for (var i = 1; i < data.length; i++) {
      var id = String(data[i][mc.ITEM_ID] || '').trim();
      var type = String(data[i][mc.TYPE] || '').trim();
      var horse = mc.REF_HORSE_NAME != null ? String(data[i][mc.REF_HORSE_NAME] || '').trim() : '';
      out.push({
        module: 'feed',
        type: type,
        date: _histYmd_(data[i][mc.DATE]),
        title: (names[id] || id) + ' · ' + type,
        details: (Number(data[i][mc.QTY] || 0) + ' · ' + String(data[i][mc.NOTES] || '').trim()
          + (horse ? ' · ' + horse : '')).replace(/^ · | · $/g, ''),
        who: String(data[i][mc.RECORDED_BY] || '').trim(),
        ref: id,
        sortKey: String(data[i][mc.RECORDED_AT] || '')
      });
    }
  } catch (ignore) {}
  return out;
}

function _histFromTack_() {
  var out = [];
  try {
    if (typeof _ensureStockSheets_ !== 'function') return out;
    var sheets = _ensureStockSheets_();
    if (sheets.tackMovements.getLastRow() < 2) return out;
    var mc = CONFIG.TACK_MOVEMENT_COLS;
    var data = sheets.tackMovements.getDataRange().getValues();
    var names = {};
    if (sheets.tack.getLastRow() > 1) {
      var td = sheets.tack.getDataRange().getValues();
      for (var t = 1; t < td.length; t++) {
        names[String(td[t][CONFIG.TACK_COLS.ITEM_ID] || '').trim()] = String(td[t][CONFIG.TACK_COLS.ITEM_NAME] || '').trim();
      }
    }
    for (var i = 1; i < data.length; i++) {
      var id = String(data[i][mc.ITEM_ID] || '').trim();
      var type = String(data[i][mc.TYPE] || '').trim();
      var from = String(data[i][mc.FROM_LOCATION] || '').trim();
      var to = String(data[i][mc.TO_LOCATION] || '').trim();
      out.push({
        module: 'tack',
        type: type,
        date: _histYmd_(data[i][mc.DATE]),
        title: (names[id] || id) + ' · ' + type,
        details: (Number(data[i][mc.QTY] || 0) + ' · ' + String(data[i][mc.REASON] || '').trim()
          + (from && to ? ' · ' + from + ' → ' + to : '')).replace(/^ · | · $/g, ''),
        who: String(data[i][mc.RECORDED_BY] || '').trim(),
        ref: id,
        sortKey: String(data[i][mc.RECORDED_AT] || '')
      });
    }
  } catch (ignore) {}
  return out;
}

function _histFromHorses_() {
  var out = [];
  try {
    if (typeof _ensureHorseCareSheets_ !== 'function') return out;
    var sheets = _ensureHorseCareSheets_();
    if (sheets.activity.getLastRow() < 2) return out;
    var c = CONFIG.HORSE_ACTIVITY_COLS;
    var data = sheets.activity.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      out.push({
        module: 'horses',
        type: String(data[i][c.ACTIVITY_TYPE] || '').trim(),
        date: _histYmd_(data[i][c.DATE]),
        title: String(data[i][c.HORSE_NAME] || data[i][c.HORSE_ID] || '').trim()
          + ' · ' + String(data[i][c.TITLE] || '').trim(),
        details: String(data[i][c.DETAILS] || '').trim(),
        who: String(data[i][c.RECORDED_BY] || '').trim(),
        ref: String(data[i][c.HORSE_ID] || '').trim(),
        sortKey: String(data[i][c.RECORDED_AT] || '')
      });
    }
  } catch (ignore) {}
  return out;
}

function _histFromStaff_() {
  var out = [];
  try {
    if (typeof _ensureGroomerSheets_ !== 'function') return out;
    var env = _ensureGroomerSheets_();
    if (env.attendance.getLastRow() > 1) {
      var ac = CONFIG.GROOMER_ATTENDANCE_COLS;
      var data = env.attendance.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var status = String(data[i][ac.STATUS] || '').trim();
        if (status !== 'Absent' && status !== 'Leave') continue;
        out.push({
          module: 'staff',
          type: status,
          date: _histYmd_(data[i][ac.DATE]),
          title: String(data[i][ac.NAME] || data[i][ac.STAFF_ID] || '').trim() + ' · ' + status,
          details: String(data[i][ac.NOTES] || '').trim(),
          who: String(data[i][ac.MARKED_BY] || '').trim(),
          ref: String(data[i][ac.STAFF_ID] || '').trim(),
          sortKey: String(data[i][ac.MARKED_AT] || '')
        });
      }
    }
    if (env.leaves.getLastRow() > 1) {
      var lc = CONFIG.GROOMER_LEAVE_COLS;
      var leaves = env.leaves.getDataRange().getValues();
      for (var l = 1; l < leaves.length; l++) {
        if (String(leaves[l][lc.STATUS] || '').trim().toLowerCase() === 'cancelled') continue;
        var start = _histYmd_(leaves[l][lc.START_DATE]);
        var end = _histYmd_(leaves[l][lc.END_DATE]);
        if (start && end && start === end) continue; // already covered by attendance
        out.push({
          module: 'staff',
          type: 'Leave range',
          date: start,
          title: String(leaves[l][lc.STAFF_ID] || '').trim() + ' · Leave ' + start + ' → ' + end,
          details: String(leaves[l][lc.REASON] || '').trim(),
          who: String(leaves[l][lc.APPLIED_BY] || '').trim(),
          ref: String(leaves[l][lc.STAFF_ID] || '').trim(),
          sortKey: String(leaves[l][lc.APPLIED_AT] || '')
        });
      }
    }
  } catch (ignore) {}
  return out;
}

function _histFromSessions_(from, to) {
  var out = [];
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (!sheet || sheet.getLastRow() < 2) return out;
    var cols = CONFIG.SCHEDULE_COLS || {};
    var data = sheet.getDataRange().getValues();
    // Prefer date/status columns when mapped; otherwise skip heavy scan
    var dateCol = cols.DATE != null ? cols.DATE : -1;
    var nameCol = cols.RIDER_NAME != null ? cols.RIDER_NAME : (cols.NAME != null ? cols.NAME : -1);
    var statusCol = cols.ATTENDANCE != null ? cols.ATTENDANCE : (cols.STATUS != null ? cols.STATUS : -1);
    if (dateCol < 0) return out;
    var cutoff = from || _histAddDays_(_histYmd_(new Date()), -30);
    for (var i = 1; i < data.length; i++) {
      var date = _histYmd_(data[i][dateCol]);
      if (!date || date < cutoff) continue;
      if (to && date > to) continue;
      var att = statusCol >= 0 ? String(data[i][statusCol] || '').trim() : '';
      out.push({
        module: 'sessions',
        type: att || 'Session',
        date: date,
        title: (nameCol >= 0 ? String(data[i][nameCol] || '').trim() : 'Session') + (att ? ' · ' + att : ''),
        details: '',
        who: '',
        ref: date,
        sortKey: date
      });
    }
  } catch (ignore) {}
  return out;
}

function _histAddDays_(ymd, days) {
  if (!ymd) return '';
  var p = ymd.split('-');
  var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  d.setDate(d.getDate() + days);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Weekly admin digest — staff absences/leaves, horse care & leases,
 * feed/tack movements, session counts for the last 7 days.
 */
function sendWeeklyOpsAdminSummary() {
  try {
    var today = _histYmd_(new Date());
    var from = _histAddDays_(today, -7);
    var to = _histAddDays_(today, -1);
    var summary = _buildWeeklyOpsSummary_(from, to);
    var emails = (typeof getAdminEmails === 'function') ? getAdminEmails() : [];
    if (!emails || !emails.length) {
      Logger.log('Weekly ops: no admin emails configured');
      return { success: false, message: 'No admin emails in Mail Info.' };
    }
    var subject = 'KE Weekly Ops · ' + from + ' → ' + to + ' · ' + (CONFIG.LOCATION_CITY || 'Hyderabad');
    var html = _weeklyOpsHtml_(summary, from, to);
    if (typeof sendMailKE_ === 'function') {
      sendMailKE_(emails[0], subject, html, {
        cc: emails.slice(1),
        textBody: summary.textFallback
      });
    } else {
      MailApp.sendEmail({ to: emails[0], cc: emails.slice(1).join(','), subject: subject, htmlBody: html });
    }
    return { success: true, message: 'Weekly ops summary sent to ' + emails[0], from: from, to: to };
  } catch (e) {
    Logger.log('Weekly ops error: ' + e);
    return { success: false, message: String(e.message || e) };
  }
}

function _buildWeeklyOpsSummary_(from, to) {
  var entries = []
    .concat(_histFromFeed_())
    .concat(_histFromTack_())
    .concat(_histFromHorses_())
    .concat(_histFromStaff_())
    .filter(function (r) {
      return (!from || !r.date || r.date >= from) && (!to || !r.date || r.date <= to);
    });
  var by = { feed: [], tack: [], horses: [], staff: [], sessions: [] };
  entries.forEach(function (e) {
    if (by[e.module]) by[e.module].push(e);
  });

  var feedRestock = by.feed.filter(function (e) { return e.type === 'Restock'; }).length;
  var feedOut = by.feed.filter(function (e) {
    return e.type === 'Consume' || e.type === 'Use' || e.type === 'Expire';
  }).length;
  var tackWear = by.tack.filter(function (e) { return /wear/i.test(e.type); }).length;
  var tackIn = by.tack.filter(function (e) { return e.type === 'Restock'; }).length;
  var leases = by.horses.filter(function (e) { return e.type === 'Lease'; });
  var care = by.horses.filter(function (e) { return e.type === 'Care'; });
  var absent = by.staff.filter(function (e) { return e.type === 'Absent'; }).length;
  var leave = by.staff.filter(function (e) { return /leave/i.test(e.type); }).length;

  var textFallback = [
    'Weekly ops ' + from + ' to ' + to,
    'Staff: ' + absent + ' absent, ' + leave + ' leave entries',
    'Horses: ' + care.length + ' care updates, ' + leases.length + ' lease changes',
    'Feed: ' + feedRestock + ' stock-in, ' + feedOut + ' stock-out',
    'Tack: ' + tackIn + ' stock-in, ' + tackWear + ' worn out'
  ].join('\n');

  return {
    from: from, to: to,
    feedRestock: feedRestock, feedOut: feedOut,
    tackIn: tackIn, tackWear: tackWear,
    careCount: care.length, leaseCount: leases.length,
    absent: absent, leave: leave,
    by: by,
    textFallback: textFallback,
    lowFeed: _weeklyLowStock_('feed'),
    lowTack: _weeklyLowStock_('tack')
  };
}

function _weeklyLowStock_(kind) {
  try {
    if (typeof getStockSummaries_ !== 'function') return [];
    var stock = getStockSummaries_();
    var items = kind === 'tack' ? (stock.tack.items || []) : (stock.feed.items || []);
    return items.filter(function (i) { return i.low || i.projectedLow || i.requiresCount; }).slice(0, 8);
  } catch (e) { return []; }
}

function _weeklyOpsHtml_(s, from, to) {
  function row(label, value) {
    return '<tr><td style="padding:8px 12px;border-bottom:1px solid #e5e7eb;color:#4b5563">' + label
      + '</td><td style="padding:8px 12px;border-bottom:1px solid #e5e7eb;font-weight:700;text-align:right">' + value + '</td></tr>';
  }
  function list(entries, empty) {
    if (!entries || !entries.length) return '<p style="color:#6b7280;font-size:13px">' + empty + '</p>';
    return '<ul style="padding-left:18px;margin:0;font-size:13px;line-height:1.5">'
      + entries.slice(0, 12).map(function (e) {
        return '<li><strong>' + (e.date || '') + '</strong> — ' + (e.title || '')
          + (e.details ? ' <span style="color:#6b7280">(' + e.details + ')</span>' : '') + '</li>';
      }).join('')
      + (entries.length > 12 ? '<li style="color:#6b7280">+' + (entries.length - 12) + ' more…</li>' : '')
      + '</ul>';
  }
  var city = (typeof CONFIG !== 'undefined' && CONFIG.LOCATION_CITY) ? CONFIG.LOCATION_CITY : 'Hyderabad';
  return '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:680px;margin:0 auto;color:#1b2118">'
    + '<div style="background:#14330f;color:#fff;padding:18px 20px;border-radius:12px 12px 0 0">'
    + '<div style="font-size:12px;opacity:.85;letter-spacing:.06em;text-transform:uppercase">Kings Equestrian · ' + city + '</div>'
    + '<div style="font-size:22px;font-weight:700;margin-top:4px">Weekly operations summary</div>'
    + '<div style="opacity:.9;margin-top:4px">' + from + ' → ' + to + '</div></div>'
    + '<div style="border:1px solid #e4ebe0;border-top:0;padding:16px 18px;border-radius:0 0 12px 12px;background:#fff">'
    + '<table style="width:100%;border-collapse:collapse;margin-bottom:16px">'
    + row('Staff absent days', s.absent)
    + row('Staff leave entries', s.leave)
    + row('Horse care updates', s.careCount)
    + row('Horse lease changes', s.leaseCount)
    + row('Feed stock-in', s.feedRestock)
    + row('Feed used / expired', s.feedOut)
    + row('Tack stock-in', s.tackIn)
    + row('Tack worn out', s.tackWear)
    + '</table>'
    + '<h3 style="font-size:15px;margin:18px 0 8px;color:#14330f">Staff highlights</h3>'
    + list(s.by.staff, 'No absences or leave recorded this week.')
    + '<h3 style="font-size:15px;margin:18px 0 8px;color:#14330f">Horse care & leases</h3>'
    + list(s.by.horses, 'No horse activity logged this week.')
    + '<h3 style="font-size:15px;margin:18px 0 8px;color:#14330f">Feed movements</h3>'
    + list(s.by.feed, 'No feed movements this week.')
    + '<h3 style="font-size:15px;margin:18px 0 8px;color:#14330f">Tack movements</h3>'
    + list(s.by.tack, 'No tack movements this week.')
    + (s.lowFeed.length || s.lowTack.length
      ? ('<h3 style="font-size:15px;margin:18px 0 8px;color:#991b1b">Stock needing attention</h3><ul style="padding-left:18px;font-size:13px">'
        + s.lowFeed.map(function (i) {
          return '<li>Feed: ' + i.name + ' (' + i.quantity + ' ' + (i.unit || '') + ')</li>';
        }).join('')
        + s.lowTack.map(function (i) {
          return '<li>Tack: ' + i.name + ' (' + i.quantity + ' at ' + (i.location || '') + ')</li>';
        }).join('')
        + '</ul>')
      : '')
    + '<p style="font-size:11px;color:#6b7280;margin-top:20px">Automated weekly digest from Stable Management.</p>'
    + '</div></div>';
}

/** Trigger-safe wrapper: skip trainer token when ScriptApp runs the job. */
function getOpsHistory(username, token, filters) {
  filters = filters || {};
  try {
    var isTrigger = !username || username === 'system';
    if (!isTrigger && typeof validateTrainerToken === 'function') {
      var v = validateTrainerToken(username, token);
      if (!v || !v.valid) throw new Error('Your trainer session has expired. Sign in again.');
    }
    var module = String(filters.module || filters.tab || 'all').trim().toLowerCase();
    var q = String(filters.q || filters.search || '').trim().toLowerCase();
    var from = String(filters.from || '').trim();
    var to = String(filters.to || '').trim();
    // Default last 30 days when no dates — avoids scanning years of ledger rows
    if (!from && !to) {
      to = _histYmd_(new Date());
      from = _histAddDays_(to, -30);
    }
    var limit = Math.min(300, Math.max(20, Number(filters.limit) || 100));
    var rows = [];
    if (module === 'all' || module === 'feed') rows = rows.concat(_histFromFeed_());
    if (module === 'all' || module === 'tack') rows = rows.concat(_histFromTack_());
    if (module === 'all' || module === 'horses') rows = rows.concat(_histFromHorses_());
    if (module === 'all' || module === 'staff') rows = rows.concat(_histFromStaff_());
    // Sessions only when explicitly requested (schedule sheet is large)
    if (module === 'sessions') rows = rows.concat(_histFromSessions_(from, to));
    rows = rows.filter(function (r) {
      if (from && r.date && r.date < from) return false;
      if (to && r.date && r.date > to) return false;
      if (!q) return true;
      var blob = [r.module, r.type, r.title, r.details, r.who, r.ref].join(' ').toLowerCase();
      return blob.indexOf(q) >= 0;
    });
    rows.sort(function (a, b) {
      var ad = a.date || '';
      var bd = b.date || '';
      if (ad !== bd) return bd.localeCompare(ad);
      return String(b.sortKey || '').localeCompare(String(a.sortKey || ''));
    });
    return {
      success: true,
      entries: rows.slice(0, limit),
      totalMatched: rows.length,
      from: from,
      to: to,
      modules: ['all', 'feed', 'tack', 'horses', 'staff', 'sessions']
    };
  } catch (e) {
    return { success: false, message: String(e.message || e), entries: [] };
  }
}
