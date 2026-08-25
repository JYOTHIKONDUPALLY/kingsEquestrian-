// ============================================================
// Training Automation (Booking_ID driven)
// ============================================================

var TRAINING_CFG = {
  SHEETS: {
    STUDENTS: 'STUDENTS',
    CURRICULUM: 'CURRICULUM',
    BOOKINGS: 'BOOKINGS',
    ATTENDANCE: 'ATTENDANCE',
    ASSESSMENT: 'ASSESSMENT',
    REFLECTION: 'REFLECTION',
    PROGRESS_LOG: 'PROGRESS_LOG',
    QUOTES: 'QUOTES',
    CERTIFICATES: 'CERTIFICATES'
  },
  PASS_THRESHOLD: 2
};

function setupTrainingSystem_2627() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  _ensureSheet(ss, TRAINING_CFG.SHEETS.STUDENTS, [
    'Student_ID', 'Student_Name', 'Parent_Name', 'Email', 'Phone', 'Program',
    'Last_Passed_Sequence', 'Current_Sequence', 'Classes_Completed', 'Total_Classes',
    'Progress_%', 'Skill_Safety_Avg', 'Skill_Riding_Avg', 'Skill_Knowledge_Avg',
    'Skill_Attitude_Avg', 'Progress_Bar', 'Grade', 'Section'
  ]);
  _ensureSheet(ss, TRAINING_CFG.SHEETS.CURRICULUM, [
    'Level', 'Class_Number', 'Title', 'Objective', 'Arena_Exercise',
    'Assessment_Criteria', 'Primary_Criteria', 'Secondary_Criteria'
  ]);
  _ensureSheet(ss, TRAINING_CFG.SHEETS.BOOKINGS, [
    'Timestamp', 'Booking_ID', 'Student_ID', 'Student_Name', 'Program', 'Booking_Date',
    'Assigned_Level', 'Assigned_Class_Number', 'Assigned_Title', 'Assigned_Sequence',
    'Status', 'Source'
  ]);
  _ensureSheet(ss, TRAINING_CFG.SHEETS.ATTENDANCE, [
    'Timestamp', 'Booking_ID', 'Student_ID', 'Present', 'Trainer_Name', 'Notes'
  ]);
  _ensureSheet(ss, TRAINING_CFG.SHEETS.ASSESSMENT, [
    'Timestamp', 'Booking_ID', 'Student_ID', 'Safety_1_4', 'Riding_1_4',
    'Knowledge_1_4', 'Attitude_1_4', 'Avg_Score', 'Pass_Fail', 'Trainer_Notes'
  ]);
  _ensureSheet(ss, TRAINING_CFG.SHEETS.REFLECTION, [
    'Timestamp', 'Booking_ID', 'Student_ID', 'Photo_Link', 'Video_Link', 'Reflection_Text'
  ]);
  _ensureSheet(ss, TRAINING_CFG.SHEETS.PROGRESS_LOG, [
    'Timestamp', 'Booking_ID', 'Student_ID', 'Student_Name', 'Program', 'Sequence',
    'Level', 'Class_Number', 'Title', 'Present', 'Avg_Score', 'Pass_Fail',
    'Action', 'Next_Sequence', 'Trainer_Notes'
  ]);
  _ensureSheet(ss, TRAINING_CFG.SHEETS.QUOTES, [
    'Quote_ID', 'Quote_Text', 'Author', 'Used_Last_On'
  ]);
  _ensureSheet(ss, TRAINING_CFG.SHEETS.CERTIFICATES, [
    'Level', 'Certificate_PDF_Link'
  ]);
}

function installTrainingTriggers_2627() {
  removeTrainingTriggers_2627();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onTrainingFormSubmit_2627').forSpreadsheet(ss).onFormSubmit().create();
  // ScriptApp.newTrigger('sendWeeklyTrainingReport_2627').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
}

function removeTrainingTriggers_2627() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'onTrainingFormSubmit_2627' || fn === 'sendWeeklyTrainingReport_2627') ScriptApp.deleteTrigger(t);
  });
}

function onTrainingFormSubmit_2627(e) {
  var sh = e.range.getSheet();
  if (sh.getName() === TRAINING_CFG.SHEETS.BOOKINGS) {
    _onBookingSubmit(e);
  } else if (sh.getName() === TRAINING_CFG.SHEETS.ASSESSMENT) {
    var map = _rowMap(sh, e.range.getRow());
    processAssessmentByBooking_2627(String(map.booking_id || '').trim());
  }
}

function _onBookingSubmit(e) {
  var sh = e.range.getSheet();
  var row = e.range.getRow();
  if (row < 2) return;
  var map = _rowMap(sh, row);
  var bookingId = String(map.booking_id || '').trim();
  if (!bookingId) {
    bookingId = _nextBookingId(sh);
    _setByHeader(sh, row, 'Booking_ID', bookingId);
  }
  var studentId = String(map.student_id || '').trim();
  var studentName = String(map.student_name || '').trim();
  var program = String(map.program || '').trim();
  if (!studentId && studentName) {
    studentId = 'STD-' + String(row).padStart(4, '0');
    _setByHeader(sh, row, 'Student_ID', studentId);
  }
  _ensureStudent(studentId, studentName, program);
  var student = _getStudent(studentId);
  if (!student) return;
  var seq = Number(student.currentSequence || 1);
  var cur = _curriculumBySeq(seq);
  if (!cur) return;
  _setByHeader(sh, row, 'Assigned_Level', cur.level);
  _setByHeader(sh, row, 'Assigned_Class_Number', cur.classNumber);
  _setByHeader(sh, row, 'Assigned_Title', cur.title);
  _setByHeader(sh, row, 'Assigned_Sequence', seq);
  _setByHeader(sh, row, 'Status', 'Booked');
}

/**
 * After school registration assigns a KE No, creates STUDENTS + BOOKINGS rows so the
 * trainer app (Booking_ID / curriculum linkage) and STUDENTS tab stay in sync.
 * Uses KE No as Student_ID — matches BOOKINGS column and rider lookup helpers.
 *
 * @param {Object} opts
 * @param {string} opts.keNo
 * @param {string} opts.studentName
 * @param {string} [opts.parentName]
 * @param {string} [opts.email]
 * @param {string} [opts.phone]
 * @param {string} [opts.program]
 * @param {string} [opts.grade]
 * @param {string} [opts.section]
 */
function ensureTrainingRecordsForSchoolRegistration(opts) {
  try {
    var keNo = String(opts && opts.keNo || '').trim();
    var studentName = String(opts && opts.studentName || '').trim();
    var program = String(opts && opts.program || 'school').trim();
    if (!keNo || !studentName) return;

    setupTrainingSystem_2627();
    _ensureStudentGradeSectionColumns();

    var shSt = _sheet(TRAINING_CFG.SHEETS.STUDENTS);
    var shBk = _sheet(TRAINING_CFG.SHEETS.BOOKINGS);
    if (!shSt || !shBk) {
      Logger.log('ensureTrainingRecordsForSchoolRegistration: STUDENTS or BOOKINGS sheet missing');
      return;
    }

    var parentName = String(opts && opts.parentName || '').trim();
    var email = String(opts && opts.email || '').trim();
    var phone = String(opts && opts.phone || '').trim();
    var grade = String(opts && opts.grade || '').trim();
    var section = String(opts && opts.section || '').trim();

    var existing = _getStudent(keNo);
    if (!existing) {
      shSt.appendRow([
        keNo,
        studentName,
        parentName,
        email,
        phone,
        program,
        0,
        1,
        0,
        _curriculumCount(),
        0,
        '',
        '',
        '',
        '',
        '',
        grade,
        section
      ]);
    } else if (grade || section) {
      _setStudentGradeSection(keNo, grade, section);
    }

    var bookingId = _nextBookingId(shBk);
    var student = _getStudent(keNo);
    var seq = student ? Number(student.currentSequence || 1) : 1;
    var cur = _curriculumBySeq(seq);

    shBk.appendRow([
      new Date(),
      bookingId,
      keNo,
      studentName,
      program,
      new Date(),
      cur ? cur.level : '',
      cur ? cur.classNumber : '',
      cur ? cur.title : '',
      seq,
      'Booked',
      'school-registration'
    ]);
    Logger.log('School registration → BOOKINGS ' + bookingId + ' STUDENT ' + keNo);
  } catch (e) {
    Logger.log('ensureTrainingRecordsForSchoolRegistration error: ' + e);
  }
}

function processAssessmentByBooking_2627(bookingId) {
  if (!bookingId) return;
  var as = _sheet(TRAINING_CFG.SHEETS.ASSESSMENT);
  var row = _latestAssessmentRowByBooking(bookingId);
  if (!as || !row) return;
  var map = _rowMap(as, row);
  var studentId = String(map.student_id || '').trim();
  if (!studentId) return;
  var safety = Number(map.safety_1_4 || 0);
  var riding = Number(map.riding_1_4 || 0);
  var knowledge = Number(map.knowledge_1_4 || 0);
  var attitude = Number(map.attitude_1_4 || 0);
  var avg = Math.round(((safety + riding + knowledge + attitude) / 4) * 100) / 100;
  var pass = avg >= TRAINING_CFG.PASS_THRESHOLD ? 'Pass' : 'Repeat';
  _setByHeader(as, row, 'Avg_Score', avg);
  _setByHeader(as, row, 'Pass_Fail', pass);

  var booking = _bookingById(bookingId);
  if (!booking) return;
  var seq = Number(booking.assignedSequence || 1);
  var nextSeq = pass === 'Pass' ? seq + 1 : seq;
  _updateStudentProgress(studentId, pass, seq, nextSeq);
  _appendProgress({
    bookingId: bookingId,
    studentId: studentId,
    studentName: booking.studentName,
    program: booking.program,
    sequence: seq,
    level: booking.assignedLevel,
    classNumber: booking.assignedClassNumber,
    title: booking.assignedTitle,
    present: _attendanceForBooking(bookingId),
    avgScore: avg,
    passFail: pass,
    action: pass === 'Pass' ? 'Advance' : 'Repeat',
    nextSequence: nextSeq,
    notes: String(map.trainer_notes || '')
  });
  _setBookingStatus(bookingId, pass === 'Pass' ? 'Completed' : 'Repeat');
  _refreshStudentMetrics(studentId);
  _maybeSendLevelCertificate(studentId, booking.studentName, booking.program, booking.assignedLevel);
}

function sendWeeklyTrainingReport_2627() {
  var students = _sheet(TRAINING_CFG.SHEETS.STUDENTS);
  var logs = _sheet(TRAINING_CFG.SHEETS.PROGRESS_LOG);
  if (!students || !logs) return;
  var sData = students.getDataRange().getValues();
  var lData = logs.getDataRange().getValues();
  var quote = _nextQuote();
  for (var i = 1; i < sData.length; i++) {
    var studentId = String(sData[i][0] || '').trim();
    var name = String(sData[i][1] || '').trim();
    var email = String(sData[i][3] || '').trim();
    if (!studentId || !email) continue;
    var count = 0, sum = 0, n = 0, passCount = 0, repeatCount = 0;
    var trend = [];
    for (var j = 1; j < lData.length; j++) {
      if (String(lData[j][2] || '').trim() !== studentId) continue;
      var ts = new Date(lData[j][0]);
      if (!_last7(ts)) continue;
      count++;
      var a = Number(lData[j][10] || 0);
      if (a > 0) { sum += a; n++; }
      var pf = String(lData[j][11] || '').trim().toLowerCase();
      if (pf === 'pass') passCount++;
      if (pf === 'repeat') repeatCount++;
      trend.push({ date: Utilities.formatDate(ts, Session.getScriptTimeZone(), 'dd-MMM'), avg: a, pf: pf });
    }
    var avg = n ? Math.round((sum / n) * 100) / 100 : 0;
    var student = _getStudent(studentId);
    var currentSeq = student ? Number(student.currentSequence || 1) : 1;
    var nextCur = _curriculumBySeq(currentSeq) || null;
    var passRate = count ? Math.round((passCount / count) * 100) : 0;
    var trendRows = trend.length
      ? trend.map(function(t){
          var bar = Math.max(0, Math.min(100, Math.round((Number(t.avg || 0) / 4) * 100)));
          return '<tr><td style="padding:6px 8px">' + _safe(t.date) + '</td><td style="padding:6px 8px">' + Number(t.avg || 0).toFixed(2) + '</td><td style="padding:6px 8px">' + _safe((t.pf || '').toUpperCase()) + '</td><td style="padding:6px 8px"><div style="height:8px;background:#edf7f1;border-radius:999px"><div style="height:8px;width:' + bar + '%;background:#2e8a5c;border-radius:999px"></div></div></td></tr>';
        }).join('')
      : '<tr><td colspan="4" style="padding:8px;color:#6b7280">No assessed sessions this week.</td></tr>';
    sendMailKE_(email, 'Weekly Progress Update - Kings Equestrian · ' + (CONFIG.LOCATION_CITY || 'Hyderabad'),
        '<div style="font-family:Arial,sans-serif;max-width:640px;margin:auto;background:#fff;border:1px solid #e5efe8;border-radius:10px;overflow:hidden">'
        + '<div style="background:#1f4e3d;color:#fff;padding:16px 18px"><h2 style="margin:0;font-size:20px">Weekly Progress Summary</h2><div style="font-size:12px;opacity:.9">Kings Equestrian Foundation · ' + schoolLocationShort_() + '</div></div>'
        + '<div style="padding:16px 18px">'
        + '<p style="margin-top:0">Dear <strong>' + _safe(name) + '</strong>,</p>'
        + '<div style="display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px">'
        + '<div style="border:1px solid #e6f2eb;border-radius:8px;padding:8px;text-align:center"><div style="font-size:11px;color:#6b7280">Attended</div><div style="font-size:20px;font-weight:700;color:#1f4e3d">' + count + '</div></div>'
        + '<div style="border:1px solid #e6f2eb;border-radius:8px;padding:8px;text-align:center"><div style="font-size:11px;color:#6b7280">Avg Score</div><div style="font-size:20px;font-weight:700;color:#1f4e3d">' + Number(avg).toFixed(2) + '</div></div>'
        + '<div style="border:1px solid #e6f2eb;border-radius:8px;padding:8px;text-align:center"><div style="font-size:11px;color:#6b7280">Pass</div><div style="font-size:20px;font-weight:700;color:#1f4e3d">' + passCount + '</div></div>'
        + '<div style="border:1px solid #e6f2eb;border-radius:8px;padding:8px;text-align:center"><div style="font-size:11px;color:#6b7280">Pass %</div><div style="font-size:20px;font-weight:700;color:#1f4e3d">' + passRate + '%</div></div>'
        + '</div>'
        + (nextCur ? '<p style="margin:14px 0 8px"><strong>Next class:</strong> ' + _safe(nextCur.level) + ' · Class ' + _safe(nextCur.classNumber) + ' · ' + _safe(nextCur.title) + '</p>' : '')
        + '<div style="margin-top:10px;border:1px solid #e6f2eb;border-radius:8px;overflow:hidden"><table style="width:100%;border-collapse:collapse;font-size:12px"><thead><tr style="background:#f4faf6"><th style="text-align:left;padding:6px 8px">Date</th><th style="text-align:left;padding:6px 8px">Avg</th><th style="text-align:left;padding:6px 8px">Status</th><th style="text-align:left;padding:6px 8px">Trend</th></tr></thead><tbody>' + trendRows + '</tbody></table></div>'
        + '<blockquote style="border-left:3px solid #1f4e3d;padding-left:10px;color:#444;margin:14px 0 0">' + _safe(quote) + '</blockquote>'
        + '<p style="margin:16px 0 0;font-size:11px;color:#6b7280">' + schoolLocationFull_() + '</p>'
        + '</div></div>',
      {});
  }
}

function createTrainingForms_2627() {
  var rider = FormApp.create('Kings Equestrian - Rider Form');
  rider.addTextItem().setTitle('Student_ID').setRequired(true);
  rider.addTextItem().setTitle('Student_Name').setRequired(true);
  rider.addDateItem().setTitle('Booking_Date').setRequired(true);
  rider.addTextItem().setTitle('Program').setRequired(true);
  rider.addSectionHeaderItem().setTitle('Reflection');
  rider.addTextItem().setTitle('Booking_ID').setRequired(true);
  rider.addTextItem().setTitle('Photo_Link');
  rider.addTextItem().setTitle('Video_Link');
  rider.addParagraphTextItem().setTitle('Reflection_Text');

  var trainer = FormApp.create('Kings Equestrian - Trainer Form');
  trainer.addTextItem().setTitle('Booking_ID').setRequired(true);
  trainer.addTextItem().setTitle('Student_ID').setRequired(true);
  trainer.addMultipleChoiceItem().setTitle('Present').setChoiceValues(['Yes', 'No']).setRequired(true);
  trainer.addMultipleChoiceItem().setTitle('Safety_1_4').setChoiceValues(['1', '2', '3', '4']).setRequired(true);
  trainer.addMultipleChoiceItem().setTitle('Riding_1_4').setChoiceValues(['1', '2', '3', '4']).setRequired(true);
  trainer.addMultipleChoiceItem().setTitle('Knowledge_1_4').setChoiceValues(['1', '2', '3', '4']).setRequired(true);
  trainer.addMultipleChoiceItem().setTitle('Attitude_1_4').setChoiceValues(['1', '2', '3', '4']).setRequired(true);
  trainer.addParagraphTextItem().setTitle('Trainer_Notes');

  return {
    riderFormEditUrl: rider.getEditUrl(),
    riderFormPublishedUrl: rider.getPublishedUrl(),
    trainerFormEditUrl: trainer.getEditUrl(),
    trainerFormPublishedUrl: trainer.getPublishedUrl()
  };
}

// ---- internal helpers ----
function _sheet(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
function _ensureSheet(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (!sh.getLastRow()) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
}
function _rowMap(sheet, row) {
  var h = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var v = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  var m = {};
  for (var i = 0; i < h.length; i++) m[String(h[i] || '').toLowerCase().trim().replace(/\s+/g, '_')] = v[i];
  return m;
}
function _setByHeader(sheet, row, header, value) {
  var h = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var key = String(header || '').toLowerCase().trim();
  for (var i = 0; i < h.length; i++) if (String(h[i] || '').toLowerCase().trim() === key) { sheet.getRange(row, i + 1).setValue(value); return; }
}
function _nextBookingId(sh) {
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  var data = sh.getDataRange().getValues(), max = 0;
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][1] || '');
    if (id.indexOf('BKG-' + today + '-') === 0) { var n = parseInt(id.split('-')[2], 10); if (!isNaN(n) && n > max) max = n; }
  }
  return 'BKG-' + today + '-' + String(max + 1).padStart(3, '0');
}
function _curriculumCount() { var sh = _sheet(TRAINING_CFG.SHEETS.CURRICULUM); return sh ? Math.max(0, sh.getLastRow() - 1) : 0; }
function _ensureStudent(studentId, studentName, program) {
  if (!studentId) return;
  var sh = _sheet(TRAINING_CFG.SHEETS.STUDENTS); if (!sh) return;
  _ensureStudentGradeSectionColumns();
  if (_getStudent(studentId)) return;
  sh.appendRow([studentId, studentName || '', '', '', '', program || '', 0, 1, 0, _curriculumCount(), 0, '', '', '', '', '', '', '']);
}
function _ensureStudentGradeSectionColumns() {
  var sh = _sheet(TRAINING_CFG.SHEETS.STUDENTS);
  if (!sh) return;
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var existing = headers.map(function (h) { return String(h || '').trim().toLowerCase(); });
  if (existing.indexOf('grade') < 0) {
    sh.getRange(1, lastCol + 1).setValue('Grade').setFontWeight('bold');
    lastCol++;
  }
  if (existing.indexOf('section') < 0 && existing.indexOf('grade') >= 0) {
    var h2 = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (x) { return String(x || '').trim().toLowerCase(); });
    if (h2.indexOf('section') < 0) sh.getRange(1, sh.getLastColumn() + 1).setValue('Section').setFontWeight('bold');
  } else if (existing.indexOf('section') < 0) {
    sh.getRange(1, lastCol + 1).setValue('Section').setFontWeight('bold');
  }
}
function _studentHeaderIndex_(sheet, headerName) {
  var h = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var key = String(headerName || '').trim().toLowerCase();
  for (var i = 0; i < h.length; i++) if (String(h[i] || '').trim().toLowerCase() === key) return i;
  return -1;
}
function _setStudentGradeSection(studentId, grade, section) {
  var sh = _sheet(TRAINING_CFG.SHEETS.STUDENTS);
  var s = _getStudent(studentId);
  if (!sh || !s) return;
  _ensureStudentGradeSectionColumns();
  if (grade) _setByHeader(sh, s.row, 'Grade', grade);
  if (section) _setByHeader(sh, s.row, 'Section', section);
}
function _getStudent(studentId) {
  var sh = _sheet(TRAINING_CFG.SHEETS.STUDENTS); if (!sh) return null;
  var d = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) if (String(d[i][0] || '').trim() === String(studentId || '').trim()) return { row: i + 1, currentSequence: Number(d[i][7] || 1) };
  return null;
}
function _curriculumBySeq(seq) {
  var sh = _sheet(TRAINING_CFG.SHEETS.CURRICULUM); if (!sh) return null;
  var d = sh.getDataRange().getValues();
  if (seq < 1 || seq >= d.length) return null;
  return { level: String(d[seq][0] || ''), classNumber: String(d[seq][1] || ''), title: String(d[seq][2] || '') };
}
function _bookingById(bookingId) {
  var sh = _sheet(TRAINING_CFG.SHEETS.BOOKINGS); if (!sh) return null;
  var d = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) if (String(d[i][1] || '').trim() === String(bookingId).trim()) return {
    row: i + 1, studentId: d[i][2], studentName: d[i][3], program: d[i][4], assignedLevel: d[i][6], assignedClassNumber: d[i][7], assignedTitle: d[i][8], assignedSequence: d[i][9]
  };
  return null;
}
function _latestAssessmentRowByBooking(bookingId) {
  var sh = _sheet(TRAINING_CFG.SHEETS.ASSESSMENT); if (!sh || sh.getLastRow() < 2) return 0;
  var d = sh.getDataRange().getValues();
  for (var i = d.length - 1; i >= 1; i--) if (String(d[i][1] || '').trim() === String(bookingId).trim()) return i + 1;
  return 0;
}
function _setBookingStatus(bookingId, status) { var b = _bookingById(bookingId); if (b) _setByHeader(_sheet(TRAINING_CFG.SHEETS.BOOKINGS), b.row, 'Status', status); }
function _attendanceForBooking(bookingId) {
  var sh = _sheet(TRAINING_CFG.SHEETS.ATTENDANCE); if (!sh) return '';
  var d = sh.getDataRange().getValues();
  for (var i = d.length - 1; i >= 1; i--) if (String(d[i][1] || '').trim() === String(bookingId).trim()) return String(d[i][3] || '');
  return '';
}
function _updateStudentProgress(studentId, pass, seq, nextSeq) {
  var sh = _sheet(TRAINING_CFG.SHEETS.STUDENTS); var s = _getStudent(studentId); if (!sh || !s) return;
  if (pass === 'Pass') {
    sh.getRange(s.row, 7).setValue(seq);
    sh.getRange(s.row, 9).setValue(Number(sh.getRange(s.row, 9).getValue() || 0) + 1);
  }
  sh.getRange(s.row, 8).setValue(nextSeq);
}
function _appendProgress(p) {
  var sh = _sheet(TRAINING_CFG.SHEETS.PROGRESS_LOG); if (!sh) return;
  sh.appendRow([new Date(), p.bookingId, p.studentId, p.studentName, p.program, p.sequence, p.level, p.classNumber, p.title, p.present, p.avgScore, p.passFail, p.action, p.nextSequence, p.notes]);
}
function _refreshStudentMetrics(studentId) {
  var sh = _sheet(TRAINING_CFG.SHEETS.STUDENTS); var a = _sheet(TRAINING_CFG.SHEETS.ASSESSMENT); var s = _getStudent(studentId); if (!sh || !a || !s) return;
  var total = _curriculumCount(), completed = Number(sh.getRange(s.row, 9).getValue() || 0), progress = total ? Math.round((completed / total) * 10000) / 100 : 0;
  sh.getRange(s.row, 10).setValue(total); sh.getRange(s.row, 11).setValue(progress);
  sh.getRange(s.row, 16).setFormula('=SPARKLINE(' + progress + ',{"charttype","bar";"max",100})');
  var d = a.getDataRange().getValues(), c = 0, s1 = 0, s2 = 0, s3 = 0, s4 = 0;
  for (var i = 1; i < d.length; i++) if (String(d[i][2] || '').trim() === String(studentId).trim()) { c++; s1 += Number(d[i][3] || 0); s2 += Number(d[i][4] || 0); s3 += Number(d[i][5] || 0); s4 += Number(d[i][6] || 0); }
  if (c) { sh.getRange(s.row, 12).setValue(Math.round((s1 / c) * 100) / 100); sh.getRange(s.row, 13).setValue(Math.round((s2 / c) * 100) / 100); sh.getRange(s.row, 14).setValue(Math.round((s3 / c) * 100) / 100); sh.getRange(s.row, 15).setValue(Math.round((s4 / c) * 100) / 100); }
}
function _levelTotal(level) { var c = _sheet(TRAINING_CFG.SHEETS.CURRICULUM); if (!c) return 0; var d = c.getDataRange().getValues(), n = 0; for (var i = 1; i < d.length; i++) if (String(d[i][0] || '').trim() === String(level).trim()) n++; return n; }
function _levelPassed(studentId, level) { var p = _sheet(TRAINING_CFG.SHEETS.PROGRESS_LOG); if (!p) return 0; var d = p.getDataRange().getValues(), n = 0; for (var i = 1; i < d.length; i++) if (String(d[i][2] || '').trim() === String(studentId).trim() && String(d[i][6] || '').trim() === String(level).trim() && String(d[i][11] || '') === 'Pass') n++; return n; }
function _studentEmail(studentId) { var s = _sheet(TRAINING_CFG.SHEETS.STUDENTS); if (!s) return ''; var d = s.getDataRange().getValues(); for (var i = 1; i < d.length; i++) if (String(d[i][0] || '').trim() === String(studentId).trim()) return String(d[i][3] || '').trim(); return ''; }
function _certificateFileId(level) { var c = _sheet(TRAINING_CFG.SHEETS.CERTIFICATES); if (!c) return ''; var d = c.getDataRange().getValues(); for (var i = 1; i < d.length; i++) if (String(d[i][0] || '').trim() === String(level).trim()) return _fileId(String(d[i][1] || '')); return ''; }
function _maybeSendLevelCertificate(studentId, studentName, program, level) {
  if (!level) return; if (_levelPassed(studentId, level) < _levelTotal(level)) return;
  var fileId = _certificateFileId(level), email = _studentEmail(studentId);
  if (!fileId || !email) return;
  sendMailKE_(email, 'Congratulations! ' + level + ' completed - Kings Equestrian',
    '<p>Dear ' + _safe(studentName) + ',</p><p>Congratulations on completing <strong>' + _safe(level) + '</strong> in ' + _safe(program) + '.</p><p>Your certificate is attached.</p>',
    { attachments: [DriveApp.getFileById(fileId).getBlob()] });
}
function _nextQuote() {
  var q = _sheet(TRAINING_CFG.SHEETS.QUOTES); if (!q || q.getLastRow() < 2) return 'Consistency builds confidence.';
  var d = q.getDataRange().getValues(), p = PropertiesService.getScriptProperties(), idx = Number(p.getProperty('QUOTE_IDX_2627') || 1);
  if (idx >= d.length) idx = 1; var text = String(d[idx][1] || '').trim() || 'Consistency builds confidence.';
  q.getRange(idx + 1, 4).setValue(new Date()); p.setProperty('QUOTE_IDX_2627', String(idx + 1)); return text;
}
function _last7(dt) { if (!(dt instanceof Date) || isNaN(dt.getTime())) return false; var ms = new Date().getTime() - dt.getTime(); return ms >= 0 && ms <= 7 * 24 * 60 * 60 * 1000; }
function _fileId(v) { var s = String(v || '').trim(); if (!s) return ''; if (s.indexOf('http') !== 0) return s; var m = s.match(/\/d\/([a-zA-Z0-9_-]+)/); if (m && m[1]) return m[1]; var q = s.match(/[?&]id=([a-zA-Z0-9_-]+)/); return q && q[1] ? q[1] : ''; }
function _safe(s) { return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
