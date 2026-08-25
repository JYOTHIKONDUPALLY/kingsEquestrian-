// ============================================================
// DEMO DATA — Attendance app (today + tomorrow sessions)
// Run from menu: Indus Equestrian → Seed Demo Attendance Data
// Safe to re-run: removes prior demo-seed schedule rows first.
// ============================================================

var DEMO_RIDERS = [
  { keNo: 'KE-DEMO-001', name: 'Aanya Reddy',   phone: '9876500001', email: 'aanya.demo@example.com', grade: '6', section: 'A' },
  { keNo: 'KE-DEMO-002', name: 'Rohan Mehta',   phone: '9876500002', email: 'rohan.demo@example.com', grade: '6', section: 'A' },
  { keNo: 'KE-DEMO-003', name: 'Priya Sharma',  phone: '9876500003', email: 'priya.demo@example.com', grade: '6', section: 'A' },
  { keNo: 'KE-DEMO-004', name: 'Kabir Nair',    phone: '9876500004', email: 'kabir.demo@example.com', grade: '6', section: 'A' },
  { keNo: 'KE-DEMO-005', name: 'Diya Patel',    phone: '9876500005', email: 'diya.demo@example.com', grade: '6', section: 'A' },
  { keNo: 'KE-DEMO-006', name: 'Arjun Iyer',    phone: '9876500006', email: 'arjun.demo@example.com', grade: '6', section: 'A' },
  { keNo: 'KE-DEMO-007', name: 'Sneha Kapoor',  phone: '9876500007', email: 'sneha.demo@example.com', grade: '6', section: 'A' },
  { keNo: 'KE-DEMO-008', name: 'Vivaan Das',    phone: '9876500008', email: 'vivaan.demo@example.com', grade: '6', section: 'A' },
  { keNo: 'KE-DEMO-009', name: 'Isha Menon',    phone: '9876500009', email: 'isha.demo@example.com', grade: '6', section: 'B' },
  { keNo: 'KE-DEMO-010', name: 'Karan Singh',   phone: '9876500010', email: 'karan.demo@example.com', grade: '6', section: 'B' },
  { keNo: 'KE-DEMO-011', name: 'Meera Joshi',   phone: '9876500011', email: 'meera.demo@example.com', grade: '6', section: 'B' },
  { keNo: 'KE-DEMO-012', name: 'Aditya Rao',    phone: '9876500012', email: 'aditya.demo@example.com', grade: '6', section: 'B' },
  { keNo: 'KE-DEMO-013', name: 'Nisha Verma',   phone: '9876500013', email: 'nisha.demo@example.com', grade: '6', section: 'B' },
  { keNo: 'KE-DEMO-014', name: 'Rahul Khanna',  phone: '9876500014', email: 'rahul.demo@example.com', grade: '6', section: 'B' },
  { keNo: 'KE-DEMO-015', name: 'Ananya Gupta',  phone: '9876500015', email: 'ananya.demo@example.com', grade: '7', section: 'A' },
  { keNo: 'KE-DEMO-016', name: 'Dev Malhotra',  phone: '9876500016', email: 'dev.demo@example.com', grade: '7', section: 'A' },
  { keNo: 'KE-DEMO-017', name: 'Sara Thomas',   phone: '9876500017', email: 'sara.demo@example.com', grade: '7', section: 'A' },
  { keNo: 'KE-DEMO-018', name: 'Yash Desai',    phone: '9876500018', email: 'yash.demo@example.com', grade: '7', section: 'A' }
];

/**
 * Menu entry: fills Riders, CURRICULUM, STUDENTS, BOOKINGS, Schedule for today/tomorrow.
 */
function seedDemoAttendanceData() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert(
    'Seed demo attendance data?',
    'This adds sample riders and sessions for TODAY and TOMORROW.\n\n'
      + 'Existing demo schedule rows (source = demo-seed) are removed first.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  if (ok !== ui.Button.YES) return;

  try {
    var stats = _runSeedDemoAttendanceData_();
    ui.alert(
      'Demo data ready!\n\n'
        + 'Today: ' + stats.today + ' sessions\n'
        + 'Tomorrow: ' + stats.tomorrow + ' sessions\n'
        + 'Riders: ' + stats.riders + '\n'
        + 'Curriculum rows: ' + stats.curriculum + '\n\n'
        + 'Open the Stable Management web app → Sessions → Today / Tomorrow.'
    );
  } catch (e) {
    Logger.log('seedDemoAttendanceData error: ' + e);
    ui.alert('Seed failed: ' + (e.message || e));
  }
}

function _runSeedDemoAttendanceData_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = Session.getScriptTimeZone();
  var today = new Date();
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);

  setupTrainingSystem_2627();
  _seedCurriculumIfNeeded_(ss);
  _seedEnsureRidersSheet_(ss);
  _clearDemoScheduleRows_(ss);

  DEMO_RIDERS.forEach(function (r) {
    _seedEnsureRider_(r);
    _seedEnsureStudentAndBooking_(r);
  });

  var todayCount = 0;
  var tomorrowCount = 0;

  // ── TODAY: group scoring demos (same time + level + class) ──
  var todaySlots = [
    { time: '10:00 - 10:30', riderIdx: [0, 1, 2, 3, 4, 5, 6, 7] },
    { time: '11:00 - 11:30', riderIdx: [8, 9, 10, 11, 12, 13] },
    { time: '15:00 - 15:30', riderIdx: [14, 15, 16, 17] }
  ];
  todaySlots.forEach(function (slot) {
    slot.riderIdx.forEach(function (ri) {
      _seedScheduleRow_(DEMO_RIDERS[ri], today, slot.time);
      todayCount++;
    });
  });

  // ── TOMORROW ──
  var tmrwSlots = [
    { time: '10:00 - 10:30', riderIdx: [0, 1, 2, 3, 4, 5, 6, 7, 8] },
    { time: '11:00 - 11:30', riderIdx: [9, 10, 11, 12, 13, 14, 15] }
  ];
  tmrwSlots.forEach(function (slot) {
    slot.riderIdx.forEach(function (ri) {
      _seedScheduleRow_(DEMO_RIDERS[ri], tomorrow, slot.time);
      tomorrowCount++;
    });
  });

  var cur = ss.getSheetByName('CURRICULUM');
  return {
    today: todayCount,
    tomorrow: tomorrowCount,
    riders: DEMO_RIDERS.length,
    curriculum: cur ? Math.max(0, cur.getLastRow() - 1) : 0,
    todayLabel: Utilities.formatDate(today, tz, 'dd-MMM-yyyy'),
    tomorrowLabel: Utilities.formatDate(tomorrow, tz, 'dd-MMM-yyyy')
  };
}

function _seedCurriculumIfNeeded_(ss) {
  var sh = ss.getSheetByName('CURRICULUM');
  if (!sh) {
    setupTrainingSystem_2627();
    sh = ss.getSheetByName('CURRICULUM');
  }
  if (!sh) return;
  if (sh.getLastRow() > 1) return;

  var rows = [
    ['KE Level 1', '1', 'Introduction & Mounting',
      'Mount safely and hold reins correctly.',
      'Mount from block; walk halt transitions.',
      '4-Confident; 3-Minor hesitation; 2-Needs help; 1-Unsafe',
      'Mounting', 'Rein hold'],
    ['KE Level 1', '2', 'Lunge & Balance',
      'Maintain balance at walk and trot on lunge.',
      'Lunge line work; posting trot intro.',
      '4-Controlled; 3-Minor issue; 2-Inconsistent; 1-Unsafe',
      'Balance', 'Rhythm'],
    ['KE Level 1', '3', 'Steering & Circles',
      'Steer through cones and 20m circles.',
      'Cone weave; circle at A.',
      '4-Accurate; 3-Minor errors; 2-Needs prompting; 1-Incorrect',
      'Steering', 'Circle shape'],
    ['KE Level 2', '1', 'Independent Walk-Trot',
      'Trot independently with correct diagonals.',
      'Rise trot on long side; change diagonal.',
      '4-Controlled; 3-Minor issue; 2-Inconsistent; 1-Unsafe',
      'Diagonals', 'Posture']
  ];
  sh.getRange(1, 1, 1, 8).setValues([[
    'Level', 'Class_Number', 'Title', 'Objective', 'Arena_Exercise',
    'Assessment_Criteria', 'Primary_Criteria', 'Secondary_Criteria'
  ]]);
  sh.getRange(2, 1, rows.length + 1, 8).setValues(rows);
  sh.setFrozenRows(1);
}

function _seedEnsureRidersSheet_(ss) {
  var sh = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
  if (sh) return;
  sh = ss.insertSheet(CONFIG.SHEETS.RIDERS);
  sh.appendRow(['KE No', 'Name', 'Email', 'Phone', 'Services', 'Participants', 'Registered', 'Notes']);
  sh.setFrozenRows(1);
}

function _seedEnsureRider_(r) {
  if (findRiderByKENo(r.keNo)) return;
  _createRiderRecord({
    keNo: r.keNo,
    name: r.name,
    email: r.email,
    phone: r.phone,
    services: '2 classes per week'
  });
}

function _seedEnsureStudentAndBooking_(r) {
  var shSt = _sheet(TRAINING_CFG.SHEETS.STUDENTS);
  var shBk = _sheet(TRAINING_CFG.SHEETS.BOOKINGS);
  if (!shSt || !shBk) return;

  if (!_getStudent(r.keNo)) {
    shSt.appendRow([
      r.keNo, r.name, '', r.email, r.phone, 'school',
      0, 1, 0, _curriculumCount(), 0, '', '', '', '', '',
      r.grade || '6', r.section || 'A'
    ]);
  } else if (r.grade || r.section) {
    _setStudentGradeSection(r.keNo, r.grade, r.section);
  }

  var data = shBk.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2] || '').trim() === r.keNo) return;
  }

  var cur = _curriculumBySeq(1);
  var bookingId = _nextBookingId(shBk);
  shBk.appendRow([
    new Date(),
    bookingId,
    r.keNo,
    r.name,
    'school',
    new Date(),
    cur ? cur.level : 'KE Level 1',
    cur ? cur.classNumber : '1',
    cur ? cur.title : 'Introduction & Mounting',
    1,
    'Booked',
    'demo-seed'
  ]);
}

function _seedScheduleRow_(rider, date, timeSlot) {
  addSessionToSchedule({
    keNo: rider.keNo,
    name: rider.name,
    phone: rider.phone,
    email: rider.email,
    service: '2 classes per week',
    date: date,
    timeSlot: timeSlot,
    participants: 1,
    status: 'Scheduled',
    source: 'demo-seed'
  });
}

function _clearDemoScheduleRows_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
  if (!sheet || sheet.getLastRow() < 2) return;
  var srcCol = CONFIG.SCHED_COLS.SOURCE + 1;
  var last = sheet.getLastRow();
  for (var row = last; row >= 2; row--) {
    var src = String(sheet.getRange(row, srcCol).getValue() || '').trim();
    var ke = String(sheet.getRange(row, CONFIG.SCHED_COLS.KE_NO + 1).getValue() || '').trim();
    if (src === 'demo-seed' || ke.indexOf('KE-DEMO-') === 0) {
      sheet.deleteRow(row);
    }
  }
}
