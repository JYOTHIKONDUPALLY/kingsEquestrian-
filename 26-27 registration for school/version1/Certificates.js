// ============================================================
// KINGS EQUESTRIAN — LEVEL COMPLETION CERTIFICATES
// File: Certificates.gs
//
// When a rider passes the LAST remaining class of a level, email them a
// congratulations note with the level certificate — attached as a PDF AND
// linked — using the file in the CERTIFICATES tab (Level | Certificate_PDF_Link).
//
// • Level completion is computed from PROGRESS_LOG (distinct passed classes)
//   vs CURRICULUM (distinct classes in that level).
// • Each rider receives a given level's certificate ONCE (deduped via Script
//   Properties, and mirrored to the Email Log for visibility).
// • Wired into saveBulkCurriculumAssessments + saveMakeupCurriculumAssessment,
//   and available as a manual catch-up menu action.
// ============================================================

var CERT_PROP_PREFIX = 'CERT|';

function _certKey_(keNo, level) {
  return CERT_PROP_PREFIX + String(keNo || '').trim() + '|' + String(level || '').trim();
}

/** Distinct number of classes that make up a level (from CURRICULUM). */
function _certLevelTotal_(level) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var c = ss.getSheetByName(TRAINING_CFG.SHEETS.CURRICULUM);
  if (!c || c.getLastRow() < 2) return 0;
  var d = c.getDataRange().getValues();
  var set = {};
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][0] || '').trim() === String(level).trim()) {
      var cn = String(d[i][1] || '').trim();
      if (cn) set[cn] = true;
    }
  }
  return Object.keys(set).length;
}

/** Distinct passed classes for a rider within a level (from PROGRESS_LOG). */
function _certLevelPassedDistinct_(keNo, level, progData) {
  if (!progData || progData.length < 2) return 0;
  var set = {};
  for (var i = 1; i < progData.length; i++) {
    var row = progData[i];
    if (String(row[2] || '').trim() !== String(keNo).trim()) continue;      // Student_ID
    if (String(row[6] || '').trim() !== String(level).trim()) continue;     // Level
    if (String(row[11] || '').trim().toLowerCase() !== 'pass') continue;    // Pass_Fail (ignore Outdated)
    var cn = String(row[7] || '').trim();                                    // Class_Number
    if (cn) set[cn] = true;
  }
  return Object.keys(set).length;
}

/** Certificate file id + shareable URL for a level (from CERTIFICATES tab). */
function _certificateInfo_(level) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var c = ss.getSheetByName(TRAINING_CFG.SHEETS.CERTIFICATES);
  if (!c || c.getLastRow() < 2) return null;
  var d = c.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][0] || '').trim() === String(level).trim()) {
      var raw = String(d[i][1] || '').trim();
      if (!raw) return null;
      var fileId = (typeof _fileId === 'function') ? _fileId(raw) : raw;
      var url = raw.indexOf('http') === 0
        ? raw
        : (fileId ? 'https://drive.google.com/file/d/' + fileId + '/view' : '');
      return { fileId: fileId, url: url };
    }
  }
  return null;
}

/** Best-effort rider email lookup (Riders sheet, then STUDENTS). */
function _certRiderEmail_(keNo) {
  try {
    if (typeof findRiderByKENo === 'function') {
      var r = findRiderByKENo(keNo);
      if (r && r.row) {
        var e = String(r.row[CONFIG.RIDER_COLS.EMAIL] || '').trim();
        if (e) return e;
      }
    }
  } catch (e) {}
  try { if (typeof _studentEmail === 'function') return _studentEmail(keNo); } catch (e2) {}
  return '';
}

/** Branded HTML body for the certificate email. */
function _certEmailHtml_(rider, linkHtml, attached) {
  var logo = (typeof CONFIG !== 'undefined' && CONFIG.LOGO_URL) ? CONFIG.LOGO_URL : '';
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
    + '<body style="font-family:Arial,sans-serif;background:#f4f7f2;margin:0;padding:0;color:#1b2118">'
    + '<div style="max-width:600px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 10px rgba(20,40,15,.12)">'
    + '  <div style="background:linear-gradient(135deg,#14330f,#1f4617);padding:26px 30px;text-align:center;color:#fff;border-bottom:3px solid #a9781a">'
    + (logo ? '    <img src="' + logo + '" style="width:66px;height:66px;border-radius:12px;border:2px solid #e9d59a;margin-bottom:10px;background:#fff;padding:4px">' : '')
    + '    <h1 style="margin:0;font-size:21px">Level Completed!</h1>'
    + '    <p style="margin:4px 0 0;font-size:12px;opacity:.9">Kings Equestrian · ' + schoolLocationShort_() + ' · ' + (CONFIG.ACADEMIC_YEAR_LABEL || '') + '</p>'
    + '  </div>'
    + '  <div style="padding:26px 30px">'
    + '    <p style="font-size:15px;margin:0 0 12px">Dear <strong>' + _certSafe_(rider.name) + '</strong>,</p>'
    + '    <p style="font-size:14px;line-height:1.6;margin:0 0 16px">Congratulations on successfully completing <strong>'
    +        _certSafe_(rider.level) + '</strong> of the Kings Equestrian riding program at ' + schoolLocationShort_() + '! '
    + '        This is a wonderful milestone — your dedication in the arena has paid off.</p>'
    + '    <div style="background:#fbf3db;border-left:4px solid #a9781a;padding:14px 18px;border-radius:4px;margin:0 0 18px">'
    + '      <p style="margin:0;font-size:14px;color:#7a5a12"><strong>&#127942; ' + _certSafe_(rider.level) + ' — Certificate of Completion</strong></p>'
    + (attached ? '      <p style="margin:6px 0 0;font-size:13px">Your certificate is attached to this email as a PDF.</p>' : '')
    +        linkHtml
    + '    </div>'
    + '    <p style="font-size:13px;line-height:1.6;color:#555;margin:0 0 8px">Keep up the great work as you progress to the next level. We look forward to seeing you continue your equestrian journey with us.</p>'
    + '    <div style="text-align:center;margin:20px 0 6px">'
    + '      <a href="' + (CONFIG.MY_RIDES_PORTAL_URL || '#') + '" style="background:#1f4617;color:#fff;padding:12px 26px;text-decoration:none;border-radius:6px;font-weight:bold;font-size:13px;display:inline-block">Open My Rides Portal</a>'
    + '    </div>'
    + '  </div>'
    + '  <div style="background:#14330f;color:#fff;padding:16px 30px;text-align:center;font-size:12px">'
    + '    ' + emailFooterHtml_()
    + '  </div>'
    + '</div></body></html>';
}

function _certSafe_(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Send one rider their level certificate (attach PDF + include link). */
function sendLevelCertificateEmail_(rider, info) {
  var subject = 'Congratulations! ' + rider.level + ' completed — Kings Equestrian · ' + (CONFIG.LOCATION_CITY || 'Hyderabad');
  try {
    var attachments = [];
    var linkHtml = '';
    if (info) {
      if (info.fileId) {
        try { attachments.push(DriveApp.getFileById(info.fileId).getBlob()); }
        catch (blobErr) { Logger.log('Certificate blob error (' + rider.level + '): ' + blobErr); }
      }
      if (info.url) {
        linkHtml = '      <p style="margin:6px 0 0;font-size:13px">View / download: '
          + '<a href="' + info.url + '" style="color:#1f4617;font-weight:600" target="_blank">' + rider.level + ' Certificate</a></p>';
      }
    }
    var htmlBody = _certEmailHtml_(rider, linkHtml, attachments.length > 0);
    var opts = { htmlBody: htmlBody };
    if (attachments.length) opts.attachments = attachments;

    var cc = [];
    try { if (typeof getCCRecipients === 'function') cc = getCCRecipients('certificate') || []; } catch (cce) {}
    if (cc.length) opts.cc = cc.join(',');

    var sendRes = sendMailKE_(rider.email, subject, opts.htmlBody, {
      cc: opts.cc || '',
      attachments: opts.attachments || []
    });
    if (typeof logEmail === 'function') {
      logEmail('Level-Certificate', rider.email, cc.join(','), subject, rider.keNo, 'Sent', 'via ' + sendRes.provider);
    }
    return true;
  } catch (e) {
    Logger.log('sendLevelCertificateEmail_ error: ' + e);
    if (typeof logEmailFailed === 'function') {
      logEmailFailed('Level-Certificate', rider.email || '', '', subject, rider.keNo, String(e));
    }
    return false;
  }
}

/**
 * For each completion candidate, send the level certificate if the rider has
 * now passed every class of that level and hasn't already received it.
 * completions: [{ keNo, name, email, level }]
 * Returns the number of certificates sent.
 */
function processLevelCompletionCertificates_(completions) {
  if (!completions || !completions.length) return 0;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var prog = ss.getSheetByName('PROGRESS_LOG');
  var progData = prog && prog.getLastRow() > 1 ? prog.getDataRange().getValues() : [];
  var props = PropertiesService.getScriptProperties();

  var sent = 0;
  var seen = {};
  for (var i = 0; i < completions.length; i++) {
    var r = completions[i] || {};
    var keNo = String(r.keNo || '').trim();
    var level = String(r.level || '').trim();
    if (!keNo || !level) continue;

    var key = _certKey_(keNo, level);
    if (seen[key]) continue;               // de-dupe within this batch
    seen[key] = true;
    if (props.getProperty(key)) continue;  // already sent in a previous run

    var total = _certLevelTotal_(level);
    if (total <= 0) continue;                                     // level unknown / no curriculum
    if (_certLevelPassedDistinct_(keNo, level, progData) < total) continue; // not finished yet

    var info = _certificateInfo_(level);    // may be null → email still sent, no attachment/link
    var email = String(r.email || '').trim() || _certRiderEmail_(keNo);
    if (!email) {
      if (typeof logEmailFailed === 'function') {
        logEmailFailed('Level-Certificate', '', '', 'Level cert (' + level + ')', keNo, 'No email on file for rider');
      }
      continue;
    }

    var ok = sendLevelCertificateEmail_({ keNo: keNo, name: r.name || keNo, level: level, email: email }, info);
    if (ok) {
      props.setProperty(key, new Date().toISOString());
      sent++;
    }
  }
  return sent;
}

// ────────────────────────────────────────────────────────────
//  MANUAL CATCH-UP — scan all riders for completed levels
// ────────────────────────────────────────────────────────────

/**
 * Menu action: scan PROGRESS_LOG for every rider/level, and email any
 * outstanding level certificates that were never sent (respects the
 * once-per-rider-per-level rule). Useful for back-filling existing data.
 */
function sendPendingLevelCertificates() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var prog = ss.getSheetByName('PROGRESS_LOG');
  if (!prog || prog.getLastRow() < 2) { ui.alert('No PROGRESS_LOG data to scan.'); return; }

  var d = prog.getDataRange().getValues();
  var seen = {};
  var completions = [];
  for (var i = 1; i < d.length; i++) {
    var keNo = String(d[i][2] || '').trim();
    var level = String(d[i][6] || '').trim();
    if (!keNo || !level) continue;
    var k = keNo + '|' + level;
    if (seen[k]) continue;
    seen[k] = true;
    completions.push({ keNo: keNo, name: String(d[i][3] || '').trim() || keNo, level: level, email: '' });
  }

  var sent = processLevelCompletionCertificates_(completions);
  ui.alert('Level certificates scan complete.\n\nSent: ' + sent + ' certificate(s).\n\n'
    + '(Only riders who have passed every class of a level — and who had not already '
    + 'received that level\'s certificate — were emailed.)');
}

/**
 * Clear the "already sent" markers so certificates can be re-sent.
 * Handy for testing or a fresh academic year.
 */
function resetLevelCertificateHistory() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert('Reset certificate history?',
    'This clears the record of which level certificates were already sent, so they '
    + 'can be sent again. Continue?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var removed = 0;
  for (var key in all) {
    if (key.indexOf(CERT_PROP_PREFIX) === 0) { props.deleteProperty(key); removed++; }
  }
  ui.alert('Cleared ' + removed + ' certificate record(s).');
}
