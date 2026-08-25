// ============================================================
// KINGS EQUESTRIAN — ATTENDANCE APP BACKEND HELPERS
// File: AttendanceBackendHelpers.gs
// Extra server functions called by the new AttendanceHTML
// ============================================================

// ────────────────────────────────────────────────────────────
//  REFLECTIONS  (maps to TRAINING_CFG.SHEETS.REFLECTION)
//  Columns: Timestamp | Booking_ID | Student_ID | Photo_Link | Video_Link | Reflection_Text
// ────────────────────────────────────────────────────────────

function getReflections() {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var sh   = ss.getSheetByName('REFLECTION') || ss.getSheetByName('Reflection') || ss.getSheetByName('REFLECTIONS');
    if (!sh || sh.getLastRow() < 2) return [];
    var data = sh.getDataRange().getValues();
    var hdrs = data[0].map(function(h){ return String(h||'').trim().toLowerCase().replace(/\s+/g,'_'); });

    function col(name){ return hdrs.indexOf(name); }
    var out = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var studentId = String(row[col('student_id')] || '').trim();
      if (!studentId) continue;
      var studentName = _getRiderNameForHelper(ss, studentId) || studentId;
      out.push({
        timestamp     : row[col('timestamp')] ? fmtDate(new Date(row[col('timestamp')])) : '',
        bookingId     : String(row[col('booking_id')] || '').trim(),
        keNo          : studentId,
        studentName   : studentName,
        photoLink     : String(row[col('photo_link')] || '').trim(),
        videoLink     : String(row[col('video_link')] || '').trim(),
        reflectionText: String(row[col('reflection_text')] || '').trim()
      });
    }
    return out.reverse(); // newest first
  } catch(e) {
    Logger.log('getReflections error: ' + e);
    return [];
  }
}

function saveReflection(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('REFLECTION') || ss.getSheetByName('Reflection');
    if (!sh) {
      // Auto-create sheet if missing (matches setupTrainingSystem_2627)
      sh = ss.insertSheet('REFLECTION');
      sh.appendRow(['Timestamp','Booking_ID','Student_ID','Photo_Link','Video_Link','Reflection_Text']);
    }
    sh.appendRow([
      new Date(),
      '',  // Booking_ID — optional, can be left blank from trainer UI
      String(data.keNo || '').trim(),
      String(data.photoLink || '').trim(),
      String(data.videoLink || '').trim(),
      String(data.reflectionText || '').trim()
    ]);
    return { success: true };
  } catch(e) {
    Logger.log('saveReflection error: ' + e);
    return { success: false, message: e.message };
  }
}

/**
 * Trainer submits reflection from Stable Management app (optional; same REFLECTION sheet as portal).
 * Booking_ID optional when not yet assigned; if provided, must belong to the rider.
 */
function saveTrainerReflection(data) {
  try {
    data = data || {};
    var keNo = String(data.keNo || '').trim();
    var bookingId = String(data.bookingId || '').trim();
    var photoLink = String(data.photoLink || '').trim();
    var videoLink = String(data.videoLink || '').trim();
    var reflectionText = String(data.reflectionText || '').trim();
    if (!keNo) return { success: false, message: 'Missing rider (KE No).' };
    if (!reflectionText && !photoLink && !videoLink) {
      return { success: false, message: 'Add reflection text or at least one photo/video link.' };
    }
    if (bookingId && typeof _bookingBelongsToKe_ === 'function' && !_bookingBelongsToKe_(bookingId, keNo)) {
      return { success: false, message: 'Booking_ID does not match this rider.' };
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('REFLECTION') || ss.getSheetByName('Reflection');
    if (!sh) {
      sh = ss.insertSheet('REFLECTION');
      sh.appendRow(['Timestamp', 'Booking_ID', 'Student_ID', 'Photo_Link', 'Video_Link', 'Reflection_Text']);
    }
    sh.appendRow([new Date(), bookingId, keNo, photoLink, videoLink, reflectionText]);
    return { success: true, message: 'Reflection saved.' };
  } catch (e) {
    Logger.log('saveTrainerReflection error: ' + e);
    return { success: false, message: String(e.message || e) };
  }
}

// ────────────────────────────────────────────────────────────
//  ENRICHED STUDENT LIST — includes curriculum progress and
//  skill averages so the Students tab can show them
//  This augments (wraps) the existing getAllRidersWithStats_Fast
// ────────────────────────────────────────────────────────────

/** Public entry point used by the PWA — served from cache when warm. */
function getAllRidersWithStats_WithCurriculum() {
  if (typeof getRidersWithStatsCached === 'function') return getRidersWithStatsCached();
  return getAllRidersWithStats_WithCurriculum_UNCACHED();
}

function getAllRidersWithStats_WithCurriculum_UNCACHED() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // Read the (large) Schedule once and reuse for both the base stats and
    // the pending-makeup counts instead of reading it twice.
    var schedSheet = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    var schedData = schedSheet && schedSheet.getLastRow() > 0 ? schedSheet.getDataRange().getValues() : [];

    var base = getAllRidersWithStats_Fast({ schedData: schedData });
    if (!base.length) return base;

    var cache = _loadCurriculumEngineCache_(ss);

    var aSheet = ss.getSheetByName('ASSESSMENT');
    var aData = aSheet && aSheet.getLastRow() > 1 ? aSheet.getDataRange().getValues() : [];
    var skillMap = {};
    for (var i = 1; i < aData.length; i++) {
      var keNo = String(aData[i][2] || '').trim();
      if (!keNo) continue;
      if (!skillMap[keNo]) skillMap[keNo] = { safety: 0, riding: 0, knowledge: 0, attitude: 0, count: 0 };
      var sm = skillMap[keNo];
      sm.safety += Number(aData[i][3] || 0);
      sm.riding += Number(aData[i][4] || 0);
      sm.knowledge += Number(aData[i][5] || 0);
      sm.attitude += Number(aData[i][6] || 0);
      sm.count++;
    }

    var gradeSectionMap = _buildGradeSectionMap_(ss);
    var pendingMakeupMap = _buildPendingMakeupCountMap_(ss, cache, schedData);

    base.forEach(function (r) {
      var gs = gradeSectionMap[r.keNo] || {};
      r.grade = gs.grade || '';
      r.section = gs.section || '';

      try {
        var cur = _curriculumProgressFromCache_(r.keNo, cache);
        r.curriculumProgress = {
          passedCount: cur.passedCount,
          totalCount: cur.totalCount,
          currentLevel: cur.currentLevel,
          currentClassNumber: cur.currentClassNumber,
          currentTitle: cur.currentTitle
        };
      } catch (e) {
        r.curriculumProgress = { passedCount: 0, totalCount: 0, currentLevel: '', currentClassNumber: '', currentTitle: '' };
      }

      var sm = skillMap[r.keNo];
      if (sm && sm.count > 0) {
        r.skillAvgs = {
          safety: Math.round((sm.safety / sm.count) * 100) / 100,
          riding: Math.round((sm.riding / sm.count) * 100) / 100,
          knowledge: Math.round((sm.knowledge / sm.count) * 100) / 100,
          attitude: Math.round((sm.attitude / sm.count) * 100) / 100
        };
      }

      r.pendingMakeupCount = pendingMakeupMap[r.keNo] || 0;
    });

    return base;
  } catch (e) {
    Logger.log('getAllRidersWithStats_WithCurriculum error: ' + e);
    return getAllRidersWithStats_Fast();
  }
}

function _buildGradeSectionMap_(ss) {
  var map = {};
  var sh = ss.getSheetByName('STUDENTS');
  if (!sh || sh.getLastRow() < 2) return map;
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var gCol = -1, sCol = -1;
  for (var h = 0; h < headers.length; h++) {
    var key = String(headers[h] || '').trim().toLowerCase();
    if (key === 'grade') gCol = h;
    if (key === 'section') sCol = h;
  }
  for (var i = 1; i < data.length; i++) {
    var ke = String(data[i][0] || '').trim();
    if (!ke) continue;
    map[ke] = {
      grade: gCol >= 0 ? String(data[i][gCol] || '').trim() : '',
      section: sCol >= 0 ? String(data[i][sCol] || '').trim() : ''
    };
  }
  return map;
}

/**
 * Build the immutable inputs used to derive late-join curriculum gaps.
 *
 * A gap is not an attendance absence. It is a curriculum class that the
 * rider's current grade+section cohort demonstrably completed before that
 * rider registered. Evidence is collected from:
 *   1) Present Schedule / Schedule Archive rows pinned with "Class done:"
 *   2) non-makeup PROGRESS_LOG Pass/Repeat rows (legacy fallback)
 *
 * Grade/section are sourced from STUDENTS. Registration time prefers the
 * matching Registration Response row and falls back to Riders.Registered.
 */
function _buildLateJoinGapContext_(ss, cache) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  cache = cache || _loadCurriculumEngineCache_(ss);
  var gsByKe = _buildGradeSectionMap_(ss);
  var joinAtByKe = {};
  var deliveredByCohort = {};

  function norm(v) { return String(v || '').trim().toLowerCase(); }
  function cohortKey(gs) {
    if (!gs || !gs.grade || !gs.section) return '';
    return norm(gs.grade) + '|' + norm(gs.section);
  }
  function addDelivery(keNo, when, level, classNo, title, source) {
    var gs = gsByKe[String(keNo || '').trim()];
    var ck = cohortKey(gs);
    var ts = when instanceof Date ? when.getTime() : new Date(when || 0).getTime();
    level = String(level || '').trim();
    classNo = String(classNo || '').trim();
    if (!ck || !ts || isNaN(ts) || !level || !classNo) return;
    if (!deliveredByCohort[ck]) deliveredByCohort[ck] = [];
    deliveredByCohort[ck].push({
      time: ts,
      level: level,
      classNumber: classNo,
      title: String(title || '').trim(),
      classKey: level + '|' + classNo,
      source: source || ''
    });
  }

  // Registration Response is the best source for the current school intake.
  try {
    var reg = typeof getRegistrationSheet_ === 'function' ? getRegistrationSheet_(ss) : null;
    if (reg && reg.getLastRow() > 1) {
      var rData = reg.getDataRange().getValues();
      for (var r = 1; r < rData.length; r++) {
        var rowKe = String(
          rData[r][CONFIG.REG_COLS.KE_NO] ||
          rData[r][CONFIG.REG_COLS.REG_REF] || ''
        ).trim();
        if (!rowKe || !gsByKe[rowKe]) continue;
        var rg = String(rData[r][CONFIG.REG_COLS.GRADE] || '').trim();
        var rs = String(rData[r][CONFIG.REG_COLS.SECTION] || '').trim();
        if (norm(rg) !== norm(gsByKe[rowKe].grade) ||
            norm(rs) !== norm(gsByKe[rowKe].section)) continue;
        var rt = rData[r][CONFIG.REG_COLS.TIMESTAMP];
        var rms = rt ? new Date(rt).getTime() : 0;
        // Latest matching registration represents the current cohort assignment.
        if (rms && (!joinAtByKe[rowKe] || rms > joinAtByKe[rowKe])) {
          joinAtByKe[rowKe] = rms;
        }
      }
    }
  } catch (regErr) {
    Logger.log('_buildLateJoinGapContext_ registration read: ' + regErr);
  }

  // Fallback for older records without a usable registration response row.
  try {
    var riders = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
    if (riders && riders.getLastRow() > 1) {
      var riderData = riders.getDataRange().getValues();
      for (var rr = 1; rr < riderData.length; rr++) {
        var rke = String(riderData[rr][CONFIG.RIDER_COLS.KE_NO] || '').trim();
        if (!rke || joinAtByKe[rke]) continue;
        var registered = riderData[rr][CONFIG.RIDER_COLS.REGISTERED];
        var registeredMs = registered ? new Date(registered).getTime() : 0;
        if (registeredMs) joinAtByKe[rke] = registeredMs;
      }
    }
  } catch (riderErr) {
    Logger.log('_buildLateJoinGapContext_ rider read: ' + riderErr);
  }

  // Modern, strongest delivery evidence: Present + pinned "Class done:".
  [CONFIG.SHEETS.SCHEDULE, 'Schedule Archive'].forEach(function (sheetName) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var status = String(data[i][CONFIG.SCHED_COLS.STATUS] || '').trim().toLowerCase();
      if (status === 'cancelled') continue;
      var att = String(data[i][CONFIG.SCHED_COLS.ATTENDANCE] || '').trim().toLowerCase();
      if (att !== 'present') continue;
      var delivered = typeof _parseDeliveredClassFromNotes_ === 'function'
        ? _parseDeliveredClassFromNotes_(data[i][CONFIG.SCHED_COLS.STAFF_NOTES]) : null;
      if (!delivered || delivered.kind !== 'done') continue;
      addDelivery(
        String(data[i][CONFIG.SCHED_COLS.KE_NO] || '').trim(),
        data[i][CONFIG.SCHED_COLS.DATE],
        delivered.level,
        delivered.classNumber,
        delivered.title,
        sheetName
      );
    }
  });

  // Legacy fallback/corroboration. Makeup work must never advance a cohort.
  try {
    var prog = ss.getSheetByName('PROGRESS_LOG');
    if (prog && prog.getLastRow() > 1) {
      var pData = prog.getDataRange().getValues();
      for (var p = 1; p < pData.length; p++) {
        var pf = String(pData[p][11] || '').trim().toLowerCase();
        if (pf !== 'pass' && pf !== 'repeat') continue;
        var action = String(pData[p][12] || '').trim().toLowerCase();
        var notes = String(pData[p][14] || '').trim().toLowerCase();
        if (action.indexOf('makeup') >= 0 || notes.indexOf('makeup') >= 0 ||
            notes.indexOf('late joiner') >= 0) continue;
        addDelivery(
          String(pData[p][2] || '').trim(),
          pData[p][0],
          pData[p][6],
          pData[p][7],
          pData[p][8],
          'PROGRESS_LOG'
        );
      }
    }
  } catch (progErr) {
    Logger.log('_buildLateJoinGapContext_ progress read: ' + progErr);
  }

  Object.keys(deliveredByCohort).forEach(function (ck) {
    deliveredByCohort[ck].sort(function (a, b) { return a.time - b.time; });
  });

  return {
    gradeSectionByKe: gsByKe,
    joinAtByKe: joinAtByKe,
    deliveredByCohort: deliveredByCohort,
    curriculumItems: cache.curriculumItems || []
  };
}

/**
 * Return scoreable synthetic gaps for one late rider.
 * excludeClassKeys prevents duplication with real No-Show makeup rows.
 */
function _getLateJoinCurriculumGaps_(ss, keNo, cache, context, excludeClassKeys) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  cache = cache || _loadCurriculumEngineCache_(ss);
  context = context || _buildLateJoinGapContext_(ss, cache);
  excludeClassKeys = excludeClassKeys || {};
  keNo = String(keNo || '').trim();
  var gs = context.gradeSectionByKe[keNo] || {};
  var joinAt = Number(context.joinAtByKe[keNo] || 0);
  if (!keNo || !joinAt || !gs.grade || !gs.section) return [];

  var ck = String(gs.grade).trim().toLowerCase() + '|' +
    String(gs.section).trim().toLowerCase();
  var events = context.deliveredByCohort[ck] || [];
  var delivered = {};
  var sequenceByKey = {};
  (context.curriculumItems || []).forEach(function (item) {
    sequenceByKey[item.level + '|' + item.classNumber] = Number(item.sequence || 0);
  });
  var furthestSequence = 0;
  var furthestEvent = null;
  events.forEach(function (ev) {
    if (ev.time >= joinAt) return;
    var previous = delivered[ev.classKey];
    if (!previous || ev.time < previous.time) delivered[ev.classKey] = ev;
    var seq = sequenceByKey[ev.classKey] || 0;
    if (seq > furthestSequence) {
      furthestSequence = seq;
      furthestEvent = ev;
    }
  });

  var passedMap = cache.passedByKe[keNo] || {};
  var gaps = [];
  (context.curriculumItems || []).forEach(function (item) {
    var key = item.level + '|' + item.classNumber;
    // Curriculum is linear in sheet row order. Reaching a later regular cohort
    // class means every preceding row was part of the cohort syllabus even if
    // an older Schedule row lacks a Class-done note.
    if (!furthestSequence || Number(item.sequence || 0) > furthestSequence) return;
    var ev = delivered[key] || furthestEvent;
    if (!ev || excludeClassKeys[key] || _isClassPassed_(passedMap, item.level, item.classNumber)) return;
    gaps.push({
      synthetic: true,
      source: 'late-join',
      classKey: key,
      level: item.level,
      classNumber: item.classNumber,
      title: item.title || ev.title || '',
      sequence: item.sequence || 0,
      cohortCompletedAt: ev.time,
      cohortCompletedDate: fmtDate(new Date(ev.time)),
      grade: gs.grade,
      section: gs.section
    });
  });
  return gaps;
}

function _buildPendingMakeupCountMap_(ss, cache, preloadedSchedData) {
  var counts = {};
  cache = cache || _loadCurriculumEngineCache_(ss);
  var sData = preloadedSchedData;
  if (!sData) {
    var sched = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    sData = sched && sched.getLastRow() > 0 ? sched.getDataRange().getValues() : [];
  }
  sData = sData || [];
  var noShowsByKe = {};

  for (var i = 1; i < sData.length; i++) {
    var ke = String(sData[i][CONFIG.SCHED_COLS.KE_NO] || '').trim();
    if (!ke) continue;
    var att = String(sData[i][CONFIG.SCHED_COLS.ATTENDANCE] || '').trim().toLowerCase();
    if (att !== 'no-show' && att !== 'noshow' && att !== 'no show') continue;
    if (!noShowsByKe[ke]) noShowsByKe[ke] = [];
    var dateVal = sData[i][CONFIG.SCHED_COLS.DATE];
    noShowsByKe[ke].push({
      staffNotes: String(sData[i][CONFIG.SCHED_COLS.STAFF_NOTES] || '').trim(),
      sortDate: dateVal ? new Date(dateVal).getTime() : 0
    });
  }

  var lateContext = _buildLateJoinGapContext_(ss, cache);
  Object.keys(lateContext.gradeSectionByKe || {}).forEach(function (keNo) {
    var assigned = {};
    var realCount = _countPendingMakeup_(keNo, noShowsByKe[keNo] || [], cache, assigned);
    var lateGaps = _getLateJoinCurriculumGaps_(ss, keNo, cache, lateContext, assigned);
    counts[keNo] = realCount + lateGaps.length;
  });
  return counts;
}

function _countPendingMakeup_(keNo, noShows, cache, assignedOut) {
  var passedMap = cache.passedByKe[keNo] || {};
  var curriculum = cache.curriculumItems;
  noShows = noShows || [];
  noShows.sort(function (a, b) { return a.sortDate - b.sortDate; });
  var assignedClasses = assignedOut || {};
  var count = 0;
  for (var n = 0; n < noShows.length; n++) {
    var row = noShows[n];
    var makeupDone = row.staffNotes.toLowerCase().indexOf('makeup scored') >= 0;
    var parsed = _parseMissedClassFromNotes_(row.staffNotes);
    var missedLevel = parsed ? parsed.level : '';
    var missedClass = parsed ? parsed.classNumber : '';

    if (!missedLevel || !missedClass) {
      for (var c = 0; c < curriculum.length; c++) {
        var item = curriculum[c];
        var key = item.level + '|' + item.classNumber;
        if (_isClassPassed_(passedMap, item.level, item.classNumber)) continue;
        if (assignedClasses[key]) continue;
        missedLevel = item.level;
        missedClass = item.classNumber;
        assignedClasses[key] = true;
        break;
      }
    } else {
      assignedClasses[missedLevel + '|' + missedClass] = true;
    }

    var classPassed = missedLevel && missedClass && _isClassPassed_(passedMap, missedLevel, missedClass);
    if (!makeupDone && !classPassed && missedLevel && missedClass) count++;
  }
  return count;
}

/**
 * Build the class-key set already represented by real No-Show rows.
 * Used by the synthetic save endpoint to prevent duplicate makeup paths.
 */
function _getRealNoShowAssignedClassKeys_(ss, keNo, cache) {
  var assigned = {};
  var noShows = [];
  var sched = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
  if (!sched || sched.getLastRow() < 2) return assigned;
  var data = sched.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][CONFIG.SCHED_COLS.KE_NO] || '').trim() !== String(keNo || '').trim()) continue;
    var att = String(data[i][CONFIG.SCHED_COLS.ATTENDANCE] || '').trim().toLowerCase();
    if (att !== 'no-show' && att !== 'noshow' && att !== 'no show') continue;
    var dateVal = data[i][CONFIG.SCHED_COLS.DATE];
    noShows.push({
      staffNotes: String(data[i][CONFIG.SCHED_COLS.STAFF_NOTES] || '').trim(),
      sortDate: dateVal ? new Date(dateVal).getTime() : 0
    });
  }
  _countPendingMakeup_(keNo, noShows, cache, assigned);
  return assigned;
}

// ────────────────────────────────────────────────────────────
//  RIDER MAKEUP / MISSED CLASS SCORING (Riders tab)
// ────────────────────────────────────────────────────────────

function getRiderMakeupDetail(keNo) {
  keNo = String(keNo || '').trim();
  if (!keNo) return { success: false, message: 'Missing KE No.' };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = Session.getScriptTimeZone();
    var cache = _loadCurriculumEngineCache_(ss);
    var gs = (_buildGradeSectionMap_(ss)[keNo]) || _getStudentGradeSection_(ss, keNo);
    var name = cache.nameByKe[keNo] || keNo;
    var passedMap = cache.passedByKe[keNo] || {};
    var curriculum = cache.curriculumItems;
    var cur = _curriculumProgressFromCache_(keNo, cache);

    var completed = [];
    var prog = ss.getSheetByName('PROGRESS_LOG');
    if (prog && prog.getLastRow() > 1) {
      var pData = prog.getDataRange().getValues();
      for (var p = 1; p < pData.length; p++) {
        if (String(pData[p][2] || '').trim() !== keNo) continue;
        var pf = String(pData[p][11] || '').trim();
        if (!pf) continue;
        completed.push({
          date: pData[p][0] ? fmtDate(new Date(pData[p][0])) : '',
          level: String(pData[p][6] || '').trim(),
          classNumber: String(pData[p][7] || '').trim(),
          title: String(pData[p][8] || '').trim(),
          avgScore: Number(pData[p][10] || 0),
          passFail: pf,
          notes: String(pData[p][14] || '').trim()
        });
      }
      completed.reverse();
    }

    var absentSessions = [];
    var pendingMakeup = [];
    var assignedClasses = {};
    var sched = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (sched && sched.getLastRow() > 1) {
      var sData = sched.getDataRange().getValues();
      var noShows = [];
      for (var i = 1; i < sData.length; i++) {
        if (String(sData[i][CONFIG.SCHED_COLS.KE_NO] || '').trim() !== keNo) continue;
        var att = String(sData[i][CONFIG.SCHED_COLS.ATTENDANCE] || '').trim().toLowerCase();
        if (att !== 'no-show' && att !== 'noshow' && att !== 'no show') continue;
        var dateVal = sData[i][CONFIG.SCHED_COLS.DATE];
        var dateLabel = dateVal ? fmtDate(new Date(dateVal)) : '';
        var dateYMD = '';
        try { dateYMD = dateVal ? Utilities.formatDate(new Date(dateVal), tz, 'yyyy-MM-dd') : ''; } catch (eD) {}
        noShows.push({
          rowIndex: i + 1,
          date: dateLabel,
          dateYMD: dateYMD,
          timeSlot: String(sData[i][CONFIG.SCHED_COLS.TIME_SLOT] || '').trim(),
          staffNotes: String(sData[i][CONFIG.SCHED_COLS.STAFF_NOTES] || '').trim(),
          sortDate: dateVal ? new Date(dateVal).getTime() : 0
        });
      }
      noShows.sort(function (a, b) { return a.sortDate - b.sortDate; });

      for (var n = 0; n < noShows.length; n++) {
        var row = noShows[n];
        var makeupDone = row.staffNotes.toLowerCase().indexOf('makeup scored') >= 0;
        var parsed = _parseMissedClassFromNotes_(row.staffNotes);
        var missedLevel = parsed ? parsed.level : '';
        var missedClass = parsed ? parsed.classNumber : '';
        var missedTitle = parsed ? parsed.title : '';

        if (!missedLevel || !missedClass) {
          for (var c = 0; c < curriculum.length; c++) {
            var item = curriculum[c];
            var key = item.level + '|' + item.classNumber;
            if (_isClassPassed_(passedMap, item.level, item.classNumber)) continue;
            if (assignedClasses[key]) continue;
            missedLevel = item.level;
            missedClass = item.classNumber;
            missedTitle = item.title;
            assignedClasses[key] = true;
            break;
          }
        } else {
          assignedClasses[missedLevel + '|' + missedClass] = true;
        }

        var classPassed = missedLevel && missedClass && _isClassPassed_(passedMap, missedLevel, missedClass);
        var needsMakeup = !makeupDone && !classPassed && missedLevel && missedClass;

        var entry = {
          rowIndex: row.rowIndex,
          date: row.date,
          dateYMD: row.dateYMD,
          timeSlot: row.timeSlot,
          missedLevel: missedLevel,
          missedClassNumber: missedClass,
          missedTitle: missedTitle,
          makeupDone: makeupDone,
          needsMakeup: needsMakeup
        };
        absentSessions.push(entry);
        if (needsMakeup) pendingMakeup.push(entry);
      }
      absentSessions.reverse();
      pendingMakeup.reverse();
    }

    // Add derived classes delivered by this grade+section before registration.
    // These are scoreable curriculum gaps, not false historical No-Show rows.
    var lateContext = _buildLateJoinGapContext_(ss, cache);
    var lateGaps = _getLateJoinCurriculumGaps_(ss, keNo, cache, lateContext, assignedClasses);
    lateGaps.forEach(function (gap) {
      var entry = {
        rowIndex: 0,
        synthetic: true,
        source: 'late-join',
        classKey: gap.classKey,
        date: gap.cohortCompletedDate || '',
        dateYMD: '',
        timeSlot: '',
        missedLevel: gap.level,
        missedClassNumber: gap.classNumber,
        missedTitle: gap.title,
        cohortCompletedAt: gap.cohortCompletedAt,
        cohortCompletedDate: gap.cohortCompletedDate,
        makeupDone: false,
        needsMakeup: true,
        displayStatus: 'Missed before joining'
      };
      absentSessions.push(entry);
      pendingMakeup.push(entry);
    });

    var bookingId = cache.bookingByKe[keNo] || '';

    return {
      success: true,
      keNo: keNo,
      name: name,
      grade: gs.grade,
      section: gs.section,
      gradeSectionLabel: gs.grade && gs.section ? (gs.grade + ' · Section ' + gs.section)
        : (gs.grade || gs.section || ''),
      bookingId: bookingId,
      currentLevel: cur.currentLevel || '',
      currentClassNumber: cur.currentClassNumber || '',
      currentTitle: cur.currentTitle || '',
      completedClasses: completed,
      absentSessions: absentSessions,
      pendingMakeup: pendingMakeup
    };
  } catch (e) {
    Logger.log('getRiderMakeupDetail error: ' + e);
    return { success: false, message: String(e.message || e) };
  }
}

// ────────────────────────────────────────────────────────────
//  PRIVATE HELPER
// ────────────────────────────────────────────────────────────

function _getRiderNameForHelper(ss, keNo) {
  try {
    var sheet = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
    if (!sheet) return '';
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][CONFIG.RIDER_COLS.KE_NO] || '').trim() === keNo)
        return String(data[i][CONFIG.RIDER_COLS.NAME] || '');
    }
    // Fall back to STUDENTS sheet
    var students = ss.getSheetByName('STUDENTS');
    if (students) {
      var sData = students.getDataRange().getValues();
      for (var j = 1; j < sData.length; j++) {
        if (String(sData[j][0] || '').trim() === keNo)
          return String(sData[j][1] || '');
      }
    }
    return keNo;
  } catch(e) { return keNo; }
}

// ────────────────────────────────────────────────────────────
//  BOOK CLASS BY GRADE + SECTION (trainer attendance app)
// ────────────────────────────────────────────────────────────

function getGradeSectionBookingOptions() {
  try {
    if (typeof _ensureStudentGradeSectionColumns === 'function') _ensureStudentGradeSectionColumns();
    var gradeSet = {}, sectionsByGrade = {};
    _collectGradeSectionFromStudents_(gradeSet, sectionsByGrade);
    _collectGradeSectionFromRegistration_(gradeSet, sectionsByGrade);
    var grades = Object.keys(gradeSet).sort(function (a, b) {
      var na = parseFloat(a), nb = parseFloat(b);
      if (!isNaN(na) && !isNaN(nb) && na !== nb) return na - nb;
      return String(a).localeCompare(String(b));
    });
    return { grades: grades, sectionsByGrade: _finalizeSectionsByGrade_(sectionsByGrade) };
  } catch (e) {
    Logger.log('getGradeSectionBookingOptions error: ' + e);
    return { grades: [], sectionsByGrade: {} };
  }
}

function previewGradeSectionBooking(grade, section) {
  var students = _getStudentsByGradeSection_(grade, section);
  return {
    count: students.length,
    students: students.map(function (s) { return { keNo: s.keNo, name: s.name }; })
  };
}

/**
 * Schedule all students in a grade + section for one or many dates.
 * opts: { grade, section, timeSlot, service?,
 *         mode: 'single'|'recurring',
 *         date (single),
 *         month | months[] | monthFrom+monthTo + pattern + customDays + excludeDates (recurring),
 *         excludeKeNos?: string[]  // riders opted out of this booking }
 */
function bookClassByGradeSection(opts) {
  opts = opts || {};
  var grade = String(opts.grade || '').trim();
  var section = String(opts.section || '').trim();
  var timeSlot = String(opts.timeSlot || '').trim();
  var service = String(opts.service || 'Regular School Classes (2026-27)').trim();
  var bookedBy = String(opts.bookedBy || '').trim();
  if (!grade || !section) return { success: false, message: 'Select grade and section.' };
  if (!timeSlot) return { success: false, message: 'Enter a class start and end time.' };

  var students = _getStudentsByGradeSection_(grade, section);
  students = _filterExcludedBookingRiders_(students, opts.excludeKeNos);
  if (!students.length) {
    return { success: false, message: 'No students left to schedule (all opted out, or none found for Grade ' + grade + ' · Section ' + section + ').' };
  }

  var dateObjs = [];
  var mode = String(opts.mode || 'single').trim();
  if (mode === 'recurring' || opts.month || opts.months || opts.monthFrom) {
    var pattern = String(opts.pattern || '').trim();
    var customDays = opts.customDays || [];
    if (!pattern) {
      return { success: false, message: 'Select a day pattern for recurring booking.' };
    }
    var months = _normalizeBookingMonths_(opts);
    if (!months.length) {
      return { success: false, message: 'Select at least one month for recurring booking.' };
    }
    months.forEach(function (m) {
      dateObjs = dateObjs.concat(_getRecurringDatesInMonth_(m, pattern, customDays));
    });
    dateObjs = _filterExcludedBookingDates_(dateObjs, opts.excludeDates);
    if (!dateObjs.length) {
      return { success: false, message: 'No matching dates left to schedule (all opted out, or none in selected months).' };
    }
  } else {
    var dateObj = _parseBookingDate_(opts.date);
    if (!dateObj) return { success: false, message: 'Invalid date.' };
    dateObjs = [dateObj];
  }

  return _bookGradeSectionSessionsBulk_(students, dateObjs, timeSlot, service, bookedBy);
}

/** Drop opted-out riders (by KE No) from a booking list. */
function _filterExcludedBookingRiders_(students, excludeKeNos) {
  var excl = {};
  (excludeKeNos || []).forEach(function (k) {
    var s = String(k || '').trim().toUpperCase();
    if (s) excl[s] = true;
  });
  if (!Object.keys(excl).length) return students || [];
  return (students || []).filter(function (stu) {
    return !excl[String(stu.keNo || '').trim().toUpperCase()];
  });
}

/**
 * Accept months as:
 *  - opts.months: ['2026-07','2026-08']
 *  - opts.monthFrom + opts.monthTo (inclusive range)
 *  - opts.month: single 'yyyy-MM' (legacy)
 */
function _normalizeBookingMonths_(opts) {
  opts = opts || {};
  var out = [];
  var seen = {};
  function add(m) {
    m = String(m || '').trim();
    if (!/^\d{4}-\d{2}$/.test(m) || seen[m]) return;
    seen[m] = true;
    out.push(m);
  }
  if (opts.months && opts.months.length) {
    opts.months.forEach(add);
  } else if (opts.monthFrom || opts.monthTo) {
    var from = String(opts.monthFrom || opts.monthTo || '').trim();
    var to = String(opts.monthTo || opts.monthFrom || '').trim();
    if (from && to) {
      var a = from <= to ? from : to;
      var b = from <= to ? to : from;
      var cur = a.split('-');
      var y = parseInt(cur[0], 10), mo = parseInt(cur[1], 10);
      var guard = 0;
      while (guard++ < 36) {
        var key = y + '-' + String(mo).padStart(2, '0');
        add(key);
        if (key === b) break;
        mo++;
        if (mo > 12) { mo = 1; y++; }
      }
    }
  } else if (opts.month) {
    add(opts.month);
  }
  out.sort();
  return out;
}

/** Drop opted-out yyyy-MM-dd dates from a recurring date list. */
function _filterExcludedBookingDates_(dateObjs, excludeDates) {
  var excl = {};
  (excludeDates || []).forEach(function (d) {
    var s = String(d || '').trim();
    if (s) excl[s] = true;
  });
  if (!Object.keys(excl).length) return dateObjs || [];
  var tz = Session.getScriptTimeZone();
  return (dateObjs || []).filter(function (dt) {
    try {
      return !excl[Utilities.formatDate(dt, tz, 'yyyy-MM-dd')];
    } catch (e) {
      return true;
    }
  });
}

function _getRecurringDatesInMonth_(monthStr, pattern, customDays) {
  var parts = String(monthStr || '').split('-');
  if (parts.length < 2) return [];
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10) - 1;
  if (isNaN(year) || isNaN(month)) return [];

  var tgt = [];
  var pat = String(pattern || '').trim();
  if (pat === 'mon') tgt = [1];
  else if (pat === 'tue') tgt = [2];
  else if (pat === 'wed') tgt = [3];
  else if (pat === 'thu') tgt = [4];
  else if (pat === 'fri') tgt = [5];
  else if (pat === 'sat') tgt = [6];
  else if (pat === 'sun') tgt = [0];
  else if (pat === 'wknd') tgt = [0, 6];
  else if (pat === 'wkdy') tgt = [1, 2, 3, 4, 5];
  else if (pat === 'cust') tgt = (customDays || []).map(function (d) { return Number(d); });
  else return [];

  var dates = [];
  var d = new Date(year, month, 1, 12, 0, 0);
  while (d.getMonth() === month) {
    if (tgt.indexOf(d.getDay()) >= 0) dates.push(new Date(d.getTime()));
    d.setDate(d.getDate() + 1);
  }
  return dates;
}

function _bookGradeSectionSessionsBulk_(students, dateObjs, timeSlot, service, bookedBy, source) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sched = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
  if (!sched) return { success: false, message: 'Schedule sheet not found.' };

  if (typeof _ensureScheduleAuditColumns_ === 'function') _ensureScheduleAuditColumns_();
  var colCount = CONFIG.SCHED_COLS.SCORED_BY + 1;
  var booker = String(bookedBy || '').trim() || 'Staff';
  var sourceTag = String(source || '').trim() || 'grade-section-booking';

  var tz = Session.getScriptTimeZone();
  var slotNorm = String(timeSlot || '').trim();
  var existing = _buildFullScheduleKeySet_(sched, tz);
  var toAppend = [];
  var skipped = 0;

  dateObjs.forEach(function (dateObj) {
    var dateYMD = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
    students.forEach(function (stu) {
      var key = stu.keNo + '|' + dateYMD + '|' + slotNorm;
      if (existing[key]) { skipped++; return; }
      var row = new Array(colCount).fill('');
      row[CONFIG.SCHED_COLS.KE_NO] = stu.keNo;
      row[CONFIG.SCHED_COLS.NAME] = stu.name;
      row[CONFIG.SCHED_COLS.PHONE] = stu.phone;
      row[CONFIG.SCHED_COLS.EMAIL] = stu.email;
      row[CONFIG.SCHED_COLS.SERVICE] = service;
      row[CONFIG.SCHED_COLS.DATE] = dateObj;
      row[CONFIG.SCHED_COLS.TIME_SLOT] = slotNorm;
      row[CONFIG.SCHED_COLS.PARTICIPANTS] = 1;
      row[CONFIG.SCHED_COLS.STATUS] = 'Scheduled';
      row[CONFIG.SCHED_COLS.ATTENDANCE] = '';
      row[CONFIG.SCHED_COLS.STAFF_NOTES] = '';
      row[CONFIG.SCHED_COLS.CAL_EVENT_ID] = '';
      row[CONFIG.SCHED_COLS.SOURCE] = sourceTag;
      row[CONFIG.SCHED_COLS.BOOKED_BY] = booker;
      row[CONFIG.SCHED_COLS.SCORED_BY] = '';
      toAppend.push(row);
      existing[key] = true;
    });
  });

  if (!toAppend.length) {
    return {
      success: skipped > 0,
      booked: 0,
      skipped: skipped,
      sessionDates: dateObjs.length,
      riders: students.length,
      message: skipped ? (skipped + ' session(s) already booked — nothing new to add') : 'No sessions added'
    };
  }

  var startRow = sched.getLastRow() + 1;
  sched.getRange(startRow, 1, toAppend.length, colCount).setValues(toAppend);
  var dateCol = CONFIG.SCHED_COLS.DATE + 1;
  sched.getRange(startRow, dateCol, toAppend.length, 1).setNumberFormat('dd-MMM-yyyy');

  if (typeof invalidateAttendanceCaches === 'function') invalidateAttendanceCaches();

  var booked = toAppend.length;
  var dateLabel = dateObjs.length > 1
    ? (dateObjs.length + ' dates')
    : Utilities.formatDate(dateObjs[0], tz, 'dd-MMM-yyyy');

  return {
    success: true,
    booked: booked,
    skipped: skipped,
    sessionDates: dateObjs.length,
    riders: students.length,
    errors: [],
    message: 'Scheduled ' + booked + ' session(s) for ' + students.length + ' rider(s)'
      + ' across ' + dateLabel
      + (skipped ? ' (' + skipped + ' skipped — already booked)' : '')
  };
}

// ────────────────────────────────────────────────────────────
//  LATE JOINERS — add unscheduled grade+section riders to a session
// ────────────────────────────────────────────────────────────

/**
 * Riders in the same grade+section who are NOT already on this date+slot.
 * opts: { grade, section, date ('today'|'tomorrow'|'yyyy-MM-dd'), timeSlot }
 */
function getEligibleLateJoinersForSession(opts) {
  try {
    opts = opts || {};
    var grade = String(opts.grade || '').trim();
    var section = String(opts.section || '').trim();
    var timeSlot = String(opts.timeSlot || '').trim();
    if (!grade || !section) {
      return { success: false, students: [], message: 'Grade and section are required.' };
    }
    if (!timeSlot) {
      return { success: false, students: [], message: 'Time slot is required.' };
    }
    var dateObj = _parseBookingDate_(opts.date);
    if (!dateObj) {
      return { success: false, students: [], message: 'Invalid date.' };
    }

    var tz = Session.getScriptTimeZone();
    var dateYMD = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
    var slotNorm = timeSlot;
    var students = _getStudentsByGradeSection_(grade, section);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sched = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    var booked = sched ? _buildScheduleKeySet_(sched, dateYMD, slotNorm) : {};

    var eligible = [];
    students.forEach(function (stu) {
      var key = stu.keNo + '|' + dateYMD + '|' + slotNorm;
      if (booked[key]) return;
      eligible.push({ keNo: stu.keNo, name: stu.name || stu.keNo });
    });

    return {
      success: true,
      students: eligible,
      grade: grade,
      section: section,
      date: dateYMD,
      timeSlot: slotNorm,
      message: eligible.length
        ? (eligible.length + ' rider(s) available to add')
        : 'All riders in this grade/section are already on this session.'
    };
  } catch (e) {
    return { success: false, students: [], message: 'Error: ' + e.message };
  }
}

/**
 * Append Schedule rows for selected late joiners on one date+slot.
 * opts: { grade, section, date, timeSlot, keNos[], service?, bookedBy? }
 */
function addRidersToSessionGroup(opts) {
  try {
    opts = opts || {};
    var grade = String(opts.grade || '').trim();
    var section = String(opts.section || '').trim();
    var timeSlot = String(opts.timeSlot || '').trim();
    var bookedBy = String(opts.bookedBy || '').trim();
    var service = String(opts.service || '').trim() || 'Regular School Classes (2026-27)';
    var keNos = opts.keNos || [];

    if (!grade || !section) return { success: false, booked: 0, message: 'Grade and section are required.' };
    if (!timeSlot) return { success: false, booked: 0, message: 'Time slot is required.' };
    if (!keNos.length) return { success: false, booked: 0, message: 'Select at least one rider.' };

    var dateObj = _parseBookingDate_(opts.date);
    if (!dateObj) return { success: false, booked: 0, message: 'Invalid date.' };

    var cohort = _getStudentsByGradeSection_(grade, section);
    var byKe = {};
    cohort.forEach(function (stu) {
      byKe[String(stu.keNo || '').trim().toUpperCase()] = stu;
    });

    var selected = [];
    var seen = {};
    keNos.forEach(function (k) {
      var u = String(k || '').trim().toUpperCase();
      if (!u || seen[u]) return;
      seen[u] = true;
      if (byKe[u]) selected.push(byKe[u]);
    });

    if (!selected.length) {
      return {
        success: false,
        booked: 0,
        message: 'No matching riders in Grade ' + grade + ' · Section ' + section + '.'
      };
    }

    var result = _bookGradeSectionSessionsBulk_(
      selected, [dateObj], timeSlot, service, bookedBy, 'late-join-session'
    );
    result.keNos = selected.map(function (s) { return s.keNo; });
    if (result.success && result.booked > 0) {
      result.message = 'Added ' + result.booked + ' rider(s) to this session'
        + (result.skipped ? ' (' + result.skipped + ' already scheduled)' : '');
    }
    return result;
  } catch (e) {
    return { success: false, booked: 0, message: 'Error: ' + e.message };
  }
}

/**
 * Mark a set of schedule rows as Cancelled (optionally with a reason).
 * rowIndexes: [sheetRowNumber, ...] (1-based, as used in the app).
 */
function cancelGroupSessions(rowIndexes, reason, actor) {
  try {
    var rows = _normalizeRowIndexes_(rowIndexes);
    if (!rows.length) return { success: false, message: 'No sessions selected.' };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sched = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (!sched) return { success: false, message: 'Schedule sheet not found.' };

    var why = String(reason || '').trim();
    var by = String(actor || '').trim();
    var stamp = 'Class cancelled' + (why ? ': ' + why : '') + (by ? ' (by ' + by + ')' : '');
    var lastRow = sched.getLastRow();
    var n = 0;

    rows.forEach(function (r) {
      if (r < 2 || r > lastRow) return;
      sched.getRange(r, CONFIG.SCHED_COLS.STATUS + 1).setValue('Cancelled');
      var noteCell = sched.getRange(r, CONFIG.SCHED_COLS.STAFF_NOTES + 1);
      var prev = String(noteCell.getValue() || '').trim();
      noteCell.setValue(prev ? prev + ' · ' + stamp : stamp);
      // Best-effort: remove the linked calendar event if present.
      try {
        var calId = String(sched.getRange(r, CONFIG.SCHED_COLS.CAL_EVENT_ID + 1).getValue() || '').trim();
        if (calId && typeof _deleteCalEventById_ === 'function') _deleteCalEventById_(calId);
      } catch (ce) {}
      n++;
    });

    if (typeof invalidateAttendanceCaches === 'function') invalidateAttendanceCaches();
    return { success: true, cancelled: n, message: n + ' session(s) cancelled.' };
  } catch (e) {
    Logger.log('cancelGroupSessions error: ' + e);
    return { success: false, message: 'Error: ' + String(e.message || e) };
  }
}

/**
 * Move a set of schedule rows to a new date (and optionally a new time slot).
 * newDate: 'yyyy-MM-dd'. newTimeSlot: 'HH:MM - HH:MM' (optional — blank keeps current).
 */
function moveGroupSessions(rowIndexes, newDate, reason, actor, newTimeSlot) {
  try {
    var rows = _normalizeRowIndexes_(rowIndexes);
    if (!rows.length) return { success: false, message: 'No sessions selected.' };
    var dateObj = _parseBookingDate_(newDate);
    if (!dateObj) return { success: false, message: 'Pick a valid new date.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sched = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (!sched) return { success: false, message: 'Schedule sheet not found.' };

    var tz = Session.getScriptTimeZone();
    var newLabel = Utilities.formatDate(dateObj, tz, 'dd-MMM-yyyy');
    var slotNorm = String(newTimeSlot || '').trim();
    var why = String(reason || '').trim();
    var by = String(actor || '').trim();
    var lastRow = sched.getLastRow();
    var n = 0;

    rows.forEach(function (r) {
      if (r < 2 || r > lastRow) return;
      var dateCell = sched.getRange(r, CONFIG.SCHED_COLS.DATE + 1);
      var oldVal = dateCell.getValue();
      var oldLabel = '';
      try { oldLabel = oldVal ? Utilities.formatDate(new Date(oldVal), tz, 'dd-MMM-yyyy') : ''; } catch (e) { oldLabel = ''; }
      var oldSlot = String(sched.getRange(r, CONFIG.SCHED_COLS.TIME_SLOT + 1).getValue() || '').trim();
      dateCell.setValue(dateObj).setNumberFormat('dd-MMM-yyyy');
      if (slotNorm) {
        sched.getRange(r, CONFIG.SCHED_COLS.TIME_SLOT + 1).setValue(slotNorm);
      }
      sched.getRange(r, CONFIG.SCHED_COLS.STATUS + 1).setValue('Rescheduled');
      var stamp = 'Class moved' + (oldLabel ? ' from ' + oldLabel : '') + ' to ' + newLabel
        + (slotNorm ? (' · time ' + (oldSlot ? oldSlot + ' → ' : '') + slotNorm) : '')
        + (why ? ' — ' + why : '') + (by ? ' (by ' + by + ')' : '');
      var noteCell = sched.getRange(r, CONFIG.SCHED_COLS.STAFF_NOTES + 1);
      var prev = String(noteCell.getValue() || '').trim();
      noteCell.setValue(prev ? prev + ' · ' + stamp : stamp);
      n++;
    });

    if (typeof invalidateAttendanceCaches === 'function') invalidateAttendanceCaches();
    var msg = n + ' session(s) moved to ' + newLabel + (slotNorm ? (' · ' + slotNorm) : '') + '.';
    return { success: true, moved: n, newDate: newLabel, newTimeSlot: slotNorm || '', message: msg };
  } catch (e) {
    Logger.log('moveGroupSessions error: ' + e);
    return { success: false, message: 'Error: ' + String(e.message || e) };
  }
}

function _normalizeRowIndexes_(rowIndexes) {
  var out = [];
  (rowIndexes || []).forEach(function (r) {
    var n = Number(r);
    if (n && out.indexOf(n) < 0) out.push(n);
  });
  return out;
}

function _buildFullScheduleKeySet_(schedSheet, tz) {
  tz = tz || Session.getScriptTimeZone();
  var data = schedSheet.getDataRange().getValues();
  var keys = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var ke = String(row[CONFIG.SCHED_COLS.KE_NO] || '').trim();
    if (!ke) continue;
    var date = row[CONFIG.SCHED_COLS.DATE];
    if (!date) continue;
    var rowYMD;
    try { rowYMD = Utilities.formatDate(new Date(date), tz, 'yyyy-MM-dd'); } catch (e) { continue; }
    var slot = String(row[CONFIG.SCHED_COLS.TIME_SLOT] || '').trim();
    if (String(row[CONFIG.SCHED_COLS.STATUS] || '').toLowerCase() === 'cancelled') continue;
    keys[ke + '|' + rowYMD + '|' + slot] = true;
  }
  return keys;
}

function _parseBookingDate_(dateStr) {
  if (!dateStr) return null;
  var s = String(dateStr).trim();
  if (s === 'today') return new Date();
  if (s === 'tomorrow') { var t = new Date(); t.setDate(t.getDate() + 1); return t; }
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    var p = s.split('-');
    return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0);
  }
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function _buildScheduleKeySet_(schedSheet, dateYMD, timeSlot) {
  var tz = Session.getScriptTimeZone();
  var data = schedSheet.getDataRange().getValues();
  var keys = {};
  var slotNorm = String(timeSlot || '').trim();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var date = row[CONFIG.SCHED_COLS.DATE];
    if (!date) continue;
    var rowYMD;
    try { rowYMD = Utilities.formatDate(new Date(date), tz, 'yyyy-MM-dd'); } catch (e) { continue; }
    if (rowYMD !== dateYMD) continue;
    if (String(row[CONFIG.SCHED_COLS.TIME_SLOT] || '').trim() !== slotNorm) continue;
    if (String(row[CONFIG.SCHED_COLS.STATUS] || '').toLowerCase() === 'cancelled') continue;
    var ke = String(row[CONFIG.SCHED_COLS.KE_NO] || '').trim();
    if (ke) keys[ke + '|' + dateYMD + '|' + slotNorm] = true;
  }
  return keys;
}

function _collectGradeSectionFromStudents_(gradeSet, sectionsByGrade) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('STUDENTS');
  if (!sh || sh.getLastRow() < 2) return;
  var data = sh.getDataRange().getValues();
  var gCol = _studentHeaderIndexForBooking_(sh, 'Grade');
  var sCol = _studentHeaderIndexForBooking_(sh, 'Section');
  if (gCol < 0) return;
  for (var i = 1; i < data.length; i++) {
    _addGradeSectionOption_(gradeSet, sectionsByGrade, data[i][gCol], sCol >= 0 ? data[i][sCol] : '');
  }
}

function _collectGradeSectionFromRegistration_(gradeSet, sectionsByGrade) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getRegistrationSheet_(ss);
  if (!sh || sh.getLastRow() < 2) return;
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var gCol = _regHeaderCol_(headers, ['grade']);
  var sCol = _regHeaderCol_(headers, ['section']);
  if (gCol < 0) gCol = CONFIG.REG_COLS.GRADE;
  if (sCol < 0 && CONFIG.REG_COLS.SECTION >= 0) sCol = CONFIG.REG_COLS.SECTION;
  for (var i = 1; i < data.length; i++) {
    _addGradeSectionOption_(gradeSet, sectionsByGrade, data[i][gCol], sCol >= 0 ? data[i][sCol] : '');
  }
}

function _regHeaderCol_(headers, names) {
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').trim().toLowerCase();
    for (var n = 0; n < names.length; n++) {
      if (h === names[n] || h.indexOf(names[n]) >= 0) return i;
    }
  }
  return -1;
}

function _addGradeSectionOption_(gradeSet, sectionsByGrade, grade, section) {
  var g = String(grade || '').trim();
  var sec = String(section || '').trim();
  if (!g) return;
  gradeSet[g] = true;
  if (!sectionsByGrade[g]) sectionsByGrade[g] = {};
  if (sec) sectionsByGrade[g][sec] = true;
}

function _finalizeSectionsByGrade_(sectionsByGrade) {
  var out = {};
  Object.keys(sectionsByGrade).forEach(function (g) {
    out[g] = Object.keys(sectionsByGrade[g] || {}).sort(function (a, b) { return String(a).localeCompare(String(b)); });
  });
  return out;
}

function _studentHeaderIndexForBooking_(sheet, headerName) {
  if (typeof _studentHeaderIndex_ === 'function') return _studentHeaderIndex_(sheet, headerName);
  var h = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var key = String(headerName || '').trim().toLowerCase();
  for (var i = 0; i < h.length; i++) if (String(h[i] || '').trim().toLowerCase() === key) return i;
  return -1;
}

function _getStudentsByGradeSection_(grade, section) {
  var targetG = String(grade || '').trim().toLowerCase();
  var targetS = String(section || '').trim().toLowerCase();
  if (!targetG || !targetS) return [];

  if (typeof _ensureStudentGradeSectionColumns === 'function') _ensureStudentGradeSectionColumns();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shSt = ss.getSheetByName('STUDENTS');
  if (!shSt || shSt.getLastRow() < 2) return [];

  var data = shSt.getDataRange().getValues();
  var gCol = _studentHeaderIndexForBooking_(shSt, 'Grade');
  var sCol = _studentHeaderIndexForBooking_(shSt, 'Section');
  if (gCol < 0 || sCol < 0) return [];

  var riderSheet = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
  var riderData = riderSheet ? riderSheet.getDataRange().getValues() : [];
  var riderByKe = {};
  for (var r = 1; r < riderData.length; r++) {
    var ke = String(riderData[r][CONFIG.RIDER_COLS.KE_NO] || '').trim();
    if (ke) riderByKe[ke] = riderData[r];
  }

  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var keNo = String(row[0] || '').trim();
    if (!keNo) continue;
    var g = String(row[gCol] || '').trim().toLowerCase();
    var sec = String(row[sCol] || '').trim().toLowerCase();
    if (g !== targetG || sec !== targetS) continue;
    var name = String(row[1] || '').trim();
    var email = String(row[3] || '').trim();
    var phone = String(row[4] || '').trim();
    var program = String(row[5] || '').trim();
    var rider = riderByKe[keNo];
    if (rider) {
      name = name || String(rider[CONFIG.RIDER_COLS.NAME] || '').trim();
      email = email || String(rider[CONFIG.RIDER_COLS.EMAIL] || '').trim();
      phone = phone || String(rider[CONFIG.RIDER_COLS.PHONE] || '').trim();
      program = program || String(rider[CONFIG.RIDER_COLS.SERVICES] || '').trim();
    }
    out.push({ keNo: keNo, name: name, email: email, phone: phone, service: program || 'school' });
  }
  out.sort(function (a, b) { return String(a.name).localeCompare(String(b.name)); });
  return out;
}

/** Menu: copy Grade + Section from Registration Response into STUDENTS (by KE No). */
function syncGradeSectionFromRegistrationToStudents() {
  var ui = SpreadsheetApp.getUi();
  try {
    if (typeof _ensureStudentGradeSectionColumns === 'function') _ensureStudentGradeSectionColumns();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var reg = getRegistrationSheet_(ss);
    var shSt = ss.getSheetByName('STUDENTS');
    if (!reg || !shSt) {
      ui.alert('Registration or STUDENTS sheet not found.');
      return;
    }
    var data = reg.getDataRange().getValues();
    var headers = data[0];
    var gCol = _regHeaderCol_(headers, ['grade']);
    var sCol = _regHeaderCol_(headers, ['section']);
    var keCol = CONFIG.REG_COLS.KE_NO >= 0 ? CONFIG.REG_COLS.KE_NO : CONFIG.REG_COLS.REG_REF;
    if (gCol < 0) gCol = CONFIG.REG_COLS.GRADE;
    if (sCol < 0 && CONFIG.REG_COLS.SECTION >= 0) sCol = CONFIG.REG_COLS.SECTION;
    var updated = 0;
    for (var i = 1; i < data.length; i++) {
      var keNo = String(data[i][keCol] || '').trim();
      var grade = String(data[i][gCol] || '').trim();
      var section = sCol >= 0 ? String(data[i][sCol] || '').trim() : '';
      if (!keNo || (!grade && !section)) continue;
      if (typeof _setStudentGradeSection === 'function') {
        _setStudentGradeSection(keNo, grade, section);
        updated++;
      }
    }
    ui.alert('Grade & section sync complete.\n\nUpdated ' + updated + ' student row(s) from registration.');
  } catch (e) {
    ui.alert('Sync failed: ' + (e.message || e));
  }
}