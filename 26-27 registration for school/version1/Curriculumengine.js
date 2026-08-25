// ============================================================
// KINGS EQUESTRIAN — CURRICULUM ENGINE
// File: CurriculumEngine.gs
// Handles: curriculum reads, booking_ID generation, trainer
//          assessment, pass/repeat progression, portal data
// ============================================================

// ────────────────────────────────────────────────────────────
//  CURRICULUM READER
//  Returns all rows from the 'circulum' (or 'CURRICULUM') tab
//  as structured objects, enriched with pass status for a rider
// ────────────────────────────────────────────────────────────

/**
 * Resolve Curriculum "Document" column index (0-based).
 * Prefers a header matching Document / Doc / Link; defaults to column I (index 8).
 */
function _curriculumDocColIndex_(headerRow) {
  var headers = headerRow || [];
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').trim().toLowerCase().replace(/[\s_]+/g, '');
    if (h === 'document' || h === 'documentlink' || h === 'doclink' || h === 'classdocument'
        || h === 'link' || h === 'doc' || h === 'googledoc') {
      return i;
    }
  }
  // Fallback: column after Secondary_Criteria (index 8)
  return 8;
}

/** Extract a usable URL from a cell value or HYPERLINK formula. */
function _extractDocUrl_(value, formula) {
  var v = String(value || '').trim();
  if (/^https?:\/\//i.test(v)) return v;
  var f = String(formula || '').trim();
  if (f) {
    var m = f.match(/HYPERLINK\s*\(\s*"([^"]+)"/i) || f.match(/HYPERLINK\s*\(\s*'([^']+)'/i);
    if (m && m[1] && /^https?:\/\//i.test(m[1])) return m[1].trim();
  }
  if (/^(docs\.google\.com|drive\.google\.com)\//i.test(v)) return 'https://' + v;
  return v;
}

/**
 * Read document URLs for all curriculum rows (handles HYPERLINK formulas).
 * Returns a map of 0-based rowIndex -> url string.
 */
function _readCurriculumDocLinks_(cur, cData) {
  var out = {};
  if (!cur || !cData || cData.length < 2) return out;
  var docCol = _curriculumDocColIndex_(cData[0]);
  var lastRow = cData.length;
  var lastCol = Math.max(cData[0].length, docCol + 1);
  var formulas = null;
  try {
    formulas = cur.getRange(1, 1, lastRow, lastCol).getFormulas();
  } catch (e) { formulas = null; }
  for (var r = 1; r < cData.length; r++) {
    var val = cData[r][docCol];
    var formula = formulas && formulas[r] ? formulas[r][docCol] : '';
    var url = _extractDocUrl_(val, formula);
    if (url && /^https?:\/\//i.test(url)) out[r] = url;
  }
  return out;
}

function getCurriculumWithProgress(keNo) {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var cur = ss.getSheetByName('CURRICULUM') || ss.getSheetByName('circulum') || ss.getSheetByName('Curriculum');
    if (!cur) return { items: [], passedCount: 0, totalCount: 0, currentLevel: '', currentClassNumber: '', currentTitle: '' };

    var cData = cur.getDataRange().getValues();
    var docByRow = _readCurriculumDocLinks_(cur, cData);
    var docCol = _curriculumDocColIndex_(cData[0]);
    // Build pass map from PROGRESS_LOG
    var passedMap = _buildPassedMap(ss, keNo);

    var items = [];
    var passedCount = 0;
    for (var r = 1; r < cData.length; r++) {
      var level     = String(cData[r][0] || '').trim();
      var classNo   = String(cData[r][1] || '').trim();
      var title     = String(cData[r][2] || '').trim();
      var objective = String(cData[r][3] || '').trim();
      var exercise  = String(cData[r][4] || '').trim();
      var criteria  = String(cData[r][5] || '').trim();
      var primary   = String(cData[r][6] || '').trim();
      var secondary = String(cData[r][7] || '').trim();
      var docLink   = docByRow[r] || _extractDocUrl_(cData[r][docCol], '') || '';
      if (!level && !classNo && !title) continue;

      var key      = level + '|' + classNo;
      var isPassed = !!passedMap[key];
      var passData = passedMap[key + '_data'] || null;
      if (isPassed) passedCount++;

      items.push({
        level: level, classNumber: classNo, title: title,
        objective: objective, exercise: exercise, criteria: criteria,
        primary: primary, secondary: secondary,
        docLink: docLink,
        passed: isPassed,
        avgScore: passData ? passData.avg : null,
        passDate: passData ? passData.date : null
      });
    }

    // Find current = first unpassed item
    var current = null;
    for (var i = 0; i < items.length; i++) {
      if (!items[i].passed) { current = items[i]; break; }
    }
    if (!current && items.length) current = items[items.length - 1];

    // Group by level for summary
    var levelMap = {};
    items.forEach(function(it) {
      if (!levelMap[it.level]) levelMap[it.level] = { total: 0, passed: 0 };
      levelMap[it.level].total++;
      if (it.passed) levelMap[it.level].passed++;
    });
    var levels = Object.keys(levelMap).map(function(lv) {
      return { level: lv, total: levelMap[lv].total, passed: levelMap[lv].passed,
               complete: levelMap[lv].passed === levelMap[lv].total };
    });

    return {
      items        : items,
      passedCount  : passedCount,
      totalCount   : items.length,
      levels       : levels,
      currentLevel : current ? current.level : '',
      currentClassNumber: current ? current.classNumber : '',
      currentTitle : current ? current.title : ''
    };
  } catch (e) {
    Logger.log('getCurriculumWithProgress error: ' + e);
    return { items: [], passedCount: 0, totalCount: 0, currentLevel: '', currentClassNumber: '', currentTitle: '' };
  }
}

function _buildPassedMap(ss, keNo) {
  var passedMap = {};
  var prog = ss.getSheetByName('PROGRESS_LOG');
  if (!prog || prog.getLastRow() < 2) return passedMap;
  var pData = prog.getDataRange().getValues();
  return _passedMapFromProgData_(pData, keNo);
}

/** One PROGRESS_LOG read → pass maps for all riders. */
function _buildAllPassedMaps_(pData) {
  var byKe = {};
  if (!pData || pData.length < 2) return byKe;
  for (var i = 1; i < pData.length; i++) {
    if (!_isActiveProgressPass_(pData[i][11])) continue;
    var keNo = String(pData[i][2] || '').trim();
    if (!keNo) continue;
    if (!byKe[keNo]) byKe[keNo] = {};
    var lv = String(pData[i][6] || '').trim();
    var cls = String(pData[i][7] || '').trim();
    var key = lv + '|' + cls;
    byKe[keNo][key] = true;
    byKe[keNo][key + '_data'] = {
      avg: pData[i][10] || '',
      date: pData[i][0] ? fmtDate(new Date(pData[i][0])) : ''
    };
  }
  return byKe;
}

function _passedMapFromProgData_(pData, keNo) {
  var target = String(keNo || '').trim().toLowerCase();
  var passedMap = {};
  if (!pData || !target) return passedMap;
  for (var i = 1; i < pData.length; i++) {
    if (!_isActiveProgressPass_(pData[i][11])) continue;
    var rowKe = String(pData[i][2] || '').trim();
    if (!rowKe || rowKe.toLowerCase() !== target) continue;
    var lv = String(pData[i][6] || '').trim();
    var cls = String(pData[i][7] || '').trim();
    var key = lv + '|' + cls;
    passedMap[key] = true;
    passedMap[key + '_data'] = {
      avg: pData[i][10] || '',
      date: pData[i][0] ? fmtDate(new Date(pData[i][0])) : ''
    };
  }
  return passedMap;
}

/** Active Pass only — Outdated / Repeat do not count toward current class. */
function _isActiveProgressPass_(passFail) {
  return String(passFail || '').trim().toLowerCase() === 'pass';
}

/**
 * Load CURRICULUM + PROGRESS_LOG + BOOKINGS + RIDERS once for batch operations.
 */
function _loadCurriculumEngineCache_(ss) {
  // Per-execution memo: many endpoints call this several times within a
  // single request (sessions + grade/section + riders). Reuse the first build.
  if (typeof _EXEC_MEMO === 'object' && _EXEC_MEMO && _EXEC_MEMO.curriculumCache) {
    return _EXEC_MEMO.curriculumCache;
  }
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var cur = ss.getSheetByName('CURRICULUM') || ss.getSheetByName('circulum') || ss.getSheetByName('Curriculum');
  var cData = cur && cur.getLastRow() > 0 ? cur.getDataRange().getValues() : [[]];
  var docByRow = cur ? _readCurriculumDocLinks_(cur, cData) : {};
  var docCol = _curriculumDocColIndex_(cData[0] || []);

  var curriculumItems = [];
  var seqByKey = {};
  var titleByKey = {};
  var docLinkByKey = {};
  var seq = 0;
  for (var r = 1; r < cData.length; r++) {
    var level = String(cData[r][0] || '').trim();
    var classNo = String(cData[r][1] || '').trim();
    var title = String(cData[r][2] || '').trim();
    if (!level && !classNo && !title) continue;
    seq++;
    var key = level + '|' + classNo;
    seqByKey[key] = seq;
    titleByKey[key] = title;
    var docLink = docByRow[r] || _extractDocUrl_(cData[r][docCol], '') || '';
    if (docLink) docLinkByKey[key] = docLink;
    curriculumItems.push({
      level: level, classNumber: classNo, title: title,
      objective: String(cData[r][3] || '').trim(),
      exercise: String(cData[r][4] || '').trim(),
      criteria: String(cData[r][5] || '').trim(),
      primary: String(cData[r][6] || '').trim(),
      secondary: String(cData[r][7] || '').trim(),
      docLink: docLink,
      sequence: seq
    });
  }

  var prog = ss.getSheetByName('PROGRESS_LOG');
  var pData = prog && prog.getLastRow() > 0 ? prog.getDataRange().getValues() : [[]];
  var passedByKe = _buildAllPassedMaps_(pData);

  var bookingByKe = {};
  var bookingOwner = {};
  var bookings = ss.getSheetByName('BOOKINGS');
  if (bookings && bookings.getLastRow() > 1) {
    var bData = bookings.getDataRange().getValues();
    for (var b = bData.length - 1; b >= 1; b--) {
      var studentId = String(bData[b][2] || '').trim();
      var bookingId = String(bData[b][1] || '').trim();
      if (bookingId) bookingOwner[bookingId] = studentId;
      if (studentId && bookingId && !bookingByKe[studentId]) bookingByKe[studentId] = bookingId;
    }
  }

  var nameByKe = {};
  var riders = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
  if (riders && riders.getLastRow() > 1) {
    var rData = riders.getDataRange().getValues();
    for (var i = 1; i < rData.length; i++) {
      var ke = String(rData[i][CONFIG.RIDER_COLS.KE_NO] || '').trim();
      if (ke) nameByKe[ke] = String(rData[i][CONFIG.RIDER_COLS.NAME] || '').trim();
    }
  }
  var students = ss.getSheetByName('STUDENTS');
  if (students && students.getLastRow() > 1) {
    var sData = students.getDataRange().getValues();
    for (var s = 1; s < sData.length; s++) {
      var ke2 = String(sData[s][0] || '').trim();
      if (ke2 && !nameByKe[ke2]) nameByKe[ke2] = String(sData[s][1] || '').trim();
    }
  }

  var built = {
    curriculumItems: curriculumItems,
    totalCount: curriculumItems.length,
    seqByKey: seqByKey,
    titleByKey: titleByKey,
    docLinkByKey: docLinkByKey,
    passedByKe: passedByKe,
    bookingByKe: bookingByKe,
    bookingOwner: bookingOwner,
    nameByKe: nameByKe
  };
  if (typeof _EXEC_MEMO === 'object' && _EXEC_MEMO) _EXEC_MEMO.curriculumCache = built;
  return built;
}

function _curriculumProgressFromCache_(keNo, cache) {
  cache = cache || _loadCurriculumEngineCache_();
  var passedMap = cache.passedByKe[keNo] || {};
  var items = [];
  var passedCount = 0;
  var levelMap = {};

  cache.curriculumItems.forEach(function (it) {
    var key = it.level + '|' + it.classNumber;
    var isPassed = !!passedMap[key];
    var passData = passedMap[key + '_data'] || null;
    if (isPassed) passedCount++;
    if (!levelMap[it.level]) levelMap[it.level] = { total: 0, passed: 0 };
    levelMap[it.level].total++;
    if (isPassed) levelMap[it.level].passed++;
    items.push({
      level: it.level, classNumber: it.classNumber, title: it.title,
      objective: it.objective, exercise: it.exercise, criteria: it.criteria,
      primary: it.primary, secondary: it.secondary,
      docLink: it.docLink || '',
      passed: isPassed,
      avgScore: passData ? passData.avg : null,
      passDate: passData ? passData.date : null
    });
  });

  var current = null;
  for (var i = 0; i < items.length; i++) {
    if (!items[i].passed) { current = items[i]; break; }
  }
  if (!current && items.length) current = items[items.length - 1];

  var levels = Object.keys(levelMap).map(function (lv) {
    return {
      level: lv, total: levelMap[lv].total, passed: levelMap[lv].passed,
      complete: levelMap[lv].passed === levelMap[lv].total
    };
  });

  return {
    items: items,
    passedCount: passedCount,
    totalCount: cache.totalCount,
    levels: levels,
    currentLevel: current ? current.level : '',
    currentClassNumber: current ? current.classNumber : '',
    currentTitle: current ? current.title : ''
  };
}

function _attachCurriculumToSession_(s, cache) {
  s.bookingId = cache.bookingByKe[s.keNo] || '';
  var cur = _curriculumProgressFromCache_(s.keNo, cache);
  s.curriculumLevel = cur.currentLevel;
  s.curriculumClassNo = cur.currentClassNumber;
  s.curriculumTitle = cur.currentTitle;
  s.curriculumPassedCount = cur.passedCount;
  s.curriculumTotalCount = cur.totalCount;
  var curLv = String(cur.currentLevel || '').trim();
  var lvlStats = null;
  (cur.levels || []).forEach(function (L) {
    if (String(L.level || '').trim() === curLv) lvlStats = L;
  });
  s.curriculumLevelPassedCount = lvlStats ? Number(lvlStats.passed) || 0 : 0;
  s.curriculumLevelTotalCount = lvlStats ? Number(lvlStats.total) || 0 : 0;
  var curItem = null;
  (cur.items || []).some(function (it) {
    if (it.level === cur.currentLevel && it.classNumber === cur.currentClassNumber) {
      curItem = it;
      return true;
    }
    return false;
  });
  s.curriculumItem = curItem;
}

// ────────────────────────────────────────────────────────────
//  TRAINER: GET SESSIONS WITH CURRICULUM CONTEXT
//  Used in Attendance App — returns today's sessions with the
//  rider's current curriculum class attached
// ────────────────────────────────────────────────────────────

function getSessionsForDate_Curriculum(dateStr) {
  var sessions = getSessionsForDate_Fast(dateStr || 'today');
  if (!sessions.length) return sessions;
  var cache = _loadCurriculumEngineCache_();

  // Attach Grade/Section for group booking display (from STUDENTS / Registration).
  // _buildGradeSectionMap_ lives in Attendancebackendhelpers.js (same GAS project).
  var gsMap = {};
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (typeof _buildGradeSectionMap_ === 'function') gsMap = _buildGradeSectionMap_(ss) || {};
  } catch (e) { gsMap = {}; }

  sessions.forEach(function (s) {
    try {
      _attachCurriculumToSession_(s, cache);
    } catch (e) {
      s.curriculumLevel = '';
      s.curriculumClassNo = '';
      s.curriculumTitle = '';
    }

    try {
      var gs = gsMap[String(s.keNo || '').trim()] || {};
      s.grade = gs.grade || '';
      s.section = gs.section || '';
      s.gradeSectionLabel = (s.grade && s.section)
        ? (s.grade + ' · Section ' + s.section)
        : (s.grade || s.section || '');
    } catch (e2) {
      s.grade = s.grade || '';
      s.section = s.section || '';
      s.gradeSectionLabel = s.gradeSectionLabel || '';
    }

    // ── Finished session pinning ────────────────────────────────
    // If this session has already been scored (or marked no-show),
    // pin it to the class actually taught that day and flag it as
    // finished, so it does NOT silently advance to the next class.
    try {
      s.sessionFinished = false;
      s.assessed = false;
      var delivered = _parseDeliveredClassFromNotes_(s.staffNotes);
      if (delivered) {
        s.curriculumLevel = delivered.level;
        s.curriculumClassNo = delivered.classNumber;
        s.curriculumTitle = delivered.title
          || cache.titleByKey[delivered.level + '|' + delivered.classNumber] || '';
        var pinnedItem = _findCurriculumItemFromCache_(cache, delivered.level, delivered.classNumber);
        if (pinnedItem) s.curriculumItem = pinnedItem;
        s.sessionFinished = true;
        s.assessed = delivered.kind === 'done';
      } else {
        // Trainer class shift (back/forward) — pin before cohort align.
        var override = _parseCurriculumOverrideFromNotes_(s.staffNotes);
        if (override && override.level && override.classNumber) {
          s.curriculumLevel = override.level;
          s.curriculumClassNo = override.classNumber;
          s.curriculumTitle = override.title
            || cache.titleByKey[override.level + '|' + override.classNumber] || '';
          var overItem = _findCurriculumItemFromCache_(cache, override.level, override.classNumber);
          if (overItem) s.curriculumItem = overItem;
          s.curriculumOverride = true;
        }
      }
    } catch (e3) {}
  });

  // ── Cohort alignment ──────────────────────────────────────────
  // Students are grouped by grade + section and move through the
  // curriculum together. Every rider in the same grade+section does
  // the SAME class on a given day. The cohort's current class is the
  // furthest-along member (max curriculum sequence): riders who were
  // absent for an earlier class are pulled forward with the group and
  // make up the missed class separately from the Riders list.
  // Sessions with an explicit Curriculum class override stay pinned.
  _alignSessionsToCohortClass_(sessions, cache);

  // Expose document link on each session for clickable class labels in the UI.
  sessions.forEach(function (s) {
    var key = String(s.curriculumLevel || '') + '|' + String(s.curriculumClassNo || '');
    var fromItem = s.curriculumItem && s.curriculumItem.docLink ? String(s.curriculumItem.docLink).trim() : '';
    var fromMap = cache.docLinkByKey ? String(cache.docLinkByKey[key] || '').trim() : '';
    s.docLink = fromItem || fromMap || '';
  });

  return sessions;
}

/**
 * Force every session sharing a grade+section to display/score the
 * cohort's current curriculum class (the furthest-along member).
 */
function _alignSessionsToCohortClass_(sessions, cache) {
  cache = cache || _loadCurriculumEngineCache_();
  var cohorts = {};

  sessions.forEach(function (s) {
    if (s.sessionFinished || s.curriculumOverride) return;
    var label = String(s.gradeSectionLabel || '').trim();
    if (!label) return;
    var key = String(s.curriculumLevel || '') + '|' + String(s.curriculumClassNo || '');
    var seq = cache.seqByKey[key] || 0;
    var cur = cohorts[label];
    if (!cur || seq > cur.seq) {
      cohorts[label] = {
        seq: seq,
        level: s.curriculumLevel || '',
        classNo: s.curriculumClassNo || '',
        title: s.curriculumTitle || '',
        item: s.curriculumItem || null
      };
    }
  });

  sessions.forEach(function (s) {
    if (s.sessionFinished || s.curriculumOverride) return;
    var label = String(s.gradeSectionLabel || '').trim();
    if (!label) return;
    var c = cohorts[label];
    if (!c || !c.seq) return;
    s.curriculumLevel = c.level;
    s.curriculumClassNo = c.classNo;
    s.curriculumTitle = c.title;
    s.curriculumItem = c.item;
  });
}

// ────────────────────────────────────────────────────────────
//  TRAINER: SAVE CURRICULUM ASSESSMENT
//  Marks attendance + records assessment + runs progression
// ────────────────────────────────────────────────────────────

/**
 * Batch save for group scorer — loops saveCurriculumAssessment per rider.
 * payloads: [{ rowIndex, keNo, bookingId, level, classNumber, safety, riding, knowledge, attitude, trainerNotes, present }, ...]
 */
function saveBulkCurriculumAssessments(payloads, scoredBy) {
  var out = { success: true, saved: 0, noShows: 0, errors: [] };
  if (!payloads || !payloads.length) {
    return { success: false, saved: 0, noShows: 0, errors: [], message: 'No assessments to save.' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sched = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (!sched) return { success: false, saved: 0, noShows: 0, errors: [], message: 'Schedule sheet not found.' };

    if (typeof _ensureScheduleAuditColumns_ === 'function') _ensureScheduleAuditColumns_();
    var scorer = String(scoredBy || '').trim();

    var cache = _loadCurriculumEngineCache_(ss);
    var aSheet = ss.getSheetByName('ASSESSMENT');
    var progSheet = ss.getSheetByName('PROGRESS_LOG');

    var lastRow = sched.getLastRow();
    var dataRows = Math.max(0, lastRow - 1);
    var attCol = CONFIG.SCHED_COLS.ATTENDANCE + 1;
    var notesCol = CONFIG.SCHED_COLS.STAFF_NOTES + 1;
    var scoredByCol = CONFIG.SCHED_COLS.SCORED_BY + 1;
    var emailCol = CONFIG.SCHED_COLS.EMAIL + 1;
    var nameCol = CONFIG.SCHED_COLS.NAME + 1;
    var dateCol = CONFIG.SCHED_COLS.DATE + 1;
    var timeCol = CONFIG.SCHED_COLS.TIME_SLOT + 1;
    var attVals = dataRows ? sched.getRange(2, attCol, dataRows, 1).getValues() : [];
    var notesVals = dataRows ? sched.getRange(2, notesCol, dataRows, 1).getValues() : [];
    var scoredByVals = dataRows ? sched.getRange(2, scoredByCol, dataRows, 1).getValues() : [];
    var emailVals = dataRows ? sched.getRange(2, emailCol, dataRows, 1).getValues() : [];
    var nameVals = dataRows ? sched.getRange(2, nameCol, dataRows, 1).getValues() : [];
    var dateVals = dataRows ? sched.getRange(2, dateCol, dataRows, 1).getValues() : [];
    var timeVals = dataRows ? sched.getRange(2, timeCol, dataRows, 1).getValues() : [];

    var assessRows = [];
    var progRows = [];
    var absentToNotify = [];
    var levelCompletions = [];
    var now = new Date();

    for (var i = 0; i < payloads.length; i++) {
      var p = payloads[i] || {};
      var ri = Number(p.rowIndex || 0);
      var idx = ri - 2;
      var keNo = String(p.keNo || '').trim();
      var level = String(p.level || '').trim();
      var classNo = String(p.classNumber || '').trim();
      var present = p.present !== false;
      var notes = String(p.trainerNotes || '').trim();

      if (!ri || idx < 0 || idx >= attVals.length) {
        out.errors.push({ rowIndex: ri, keNo: keNo, message: 'Invalid schedule row.' });
        continue;
      }

      if (!present) {
        attVals[idx][0] = 'No-Show';
        if (scorer && scoredByVals[idx]) scoredByVals[idx][0] = scorer;
        var noShowTitle = cache.titleByKey[level + '|' + classNo] || '';
        if (level && classNo) {
          var missedNote = 'Missed class: ' + level + ' · Class ' + classNo + (noShowTitle ? ' — ' + noShowTitle : '');
          // Rebuild cleanly (strip any prior class marker) so re-saves don't duplicate.
          var carry = _stripClassMarkers_(notesVals[idx][0]);
          notesVals[idx][0] = missedNote + (carry ? ' · ' + carry : '');
        } else if (notes) {
          notesVals[idx][0] = notes;
        }
        // Queue an absent-notification email to the rider's registered address.
        var absentEmail = emailVals[idx] ? String(emailVals[idx][0] || '').trim() : '';
        if (absentEmail) {
          absentToNotify.push({
            keNo: keNo,
            name: nameVals[idx] ? String(nameVals[idx][0] || '').trim() : (cache.nameByKe[keNo] || keNo),
            email: absentEmail,
            level: level,
            classNumber: classNo,
            title: noShowTitle,
            date: dateVals[idx] ? dateVals[idx][0] : '',
            timeSlot: timeVals[idx] ? String(timeVals[idx][0] || '').trim() : '',
            scoredBy: scorer
          });
        }
        out.noShows++;
        continue;
      }

      var safety = Number(p.safety || 0);
      var riding = Number(p.riding || 0);
      var knowledge = Number(p.knowledge || 0);
      var attitude = Number(p.attitude || 0);
      if (!safety || !riding || !knowledge || !attitude) {
        out.errors.push({ rowIndex: ri, keNo: keNo, message: 'All four scores required.' });
        continue;
      }

      var avg = Math.round(((safety + riding + knowledge + attitude) / 4) * 100) / 100;
      var passFail = avg >= 2 ? 'Pass' : 'Repeat';
      var bookingId = String(p.bookingId || '').trim() || cache.bookingByKe[keNo] || '';
      if (!bookingId) {
        out.errors.push({ rowIndex: ri, keNo: keNo, message: 'Booking_ID is required.' });
        continue;
      }
      if (String(cache.bookingOwner[bookingId] || '').trim() !== keNo) {
        out.errors.push({ rowIndex: ri, keNo: keNo, message: 'Booking_ID does not belong to this rider.' });
        continue;
      }

      attVals[idx][0] = 'Present';
      if (scorer && scoredByVals[idx]) scoredByVals[idx][0] = scorer;
      // Record the delivered class so the session card can show it as finished
      // and stays pinned to the class actually taught (not the next class).
      var doneTitle = cache.titleByKey[level + '|' + classNo] || '';
      var doneNote = (level && classNo)
        ? ('Class done: ' + level + ' · Class ' + classNo + (doneTitle ? ' — ' + doneTitle : ''))
        : '';
      var carryPresent = _stripClassMarkers_(notesVals[idx][0]);
      var extraNotes = notes || carryPresent;
      notesVals[idx][0] = doneNote
        ? (doneNote + (extraNotes ? ' · ' + extraNotes : ''))
        : extraNotes;

      assessRows.push([now, bookingId, keNo, safety, riding, knowledge, attitude, avg, passFail, notes]);

      var sequence = cache.seqByKey[level + '|' + classNo] || 0;
      var nextSeq = passFail === 'Pass' ? sequence + 1 : sequence;
      var riderName = cache.nameByKe[keNo] || keNo;
      var classTitle = cache.titleByKey[level + '|' + classNo] || '';
      progRows.push([
        now, bookingId, keNo, riderName, '', sequence, level, classNo, classTitle,
        'Yes', avg, passFail,
        passFail === 'Pass' ? 'Move to next class' : 'Repeat class',
        nextSeq, notes
      ]);
      // Candidate for a level-completion certificate (gated later on full pass).
      if (passFail === 'Pass' && level) {
        levelCompletions.push({
          keNo: keNo,
          name: riderName,
          email: emailVals[idx] ? String(emailVals[idx][0] || '').trim() : '',
          level: level
        });
      }
      out.saved++;
    }

    if (attVals.length) sched.getRange(2, attCol, attVals.length, 1).setValues(attVals);
    if (notesVals.length) sched.getRange(2, notesCol, notesVals.length, 1).setValues(notesVals);
    if (scorer && scoredByVals.length) sched.getRange(2, scoredByCol, scoredByVals.length, 1).setValues(scoredByVals);

    if (assessRows.length && aSheet) {
      var aStart = aSheet.getLastRow() + 1;
      aSheet.getRange(aStart, 1, assessRows.length, assessRows[0].length).setValues(assessRows);
    }
    if (progRows.length && progSheet) {
      var pStart = progSheet.getLastRow() + 1;
      progSheet.getRange(pStart, 1, progRows.length, progRows[0].length).setValues(progRows);
    }

    // Notify absent riders (best-effort; never fails the save).
    var notified = 0;
    for (var n = 0; n < absentToNotify.length; n++) {
      try {
        if (typeof sendAbsentNotificationEmail_ === 'function' && sendAbsentNotificationEmail_(absentToNotify[n])) notified++;
      } catch (mailErr) {
        Logger.log('absent email error: ' + mailErr);
      }
    }
    out.notified = notified;

    // Email level-completion certificates to any rider who just finished a level
    // (best-effort; deduped so each rider gets a level's certificate only once).
    var certsSent = 0;
    try {
      if (levelCompletions.length && typeof processLevelCompletionCertificates_ === 'function') {
        certsSent = processLevelCompletionCertificates_(levelCompletions);
      }
    } catch (certErr) {
      Logger.log('level certificate error: ' + certErr);
    }
    out.certificates = certsSent;

    if (typeof invalidateAttendanceCaches === 'function') invalidateAttendanceCaches();

    out.success = out.errors.length === 0;
    out.message = out.success
      ? 'Saved ' + out.saved + ' assessment(s)' + (out.noShows ? ', ' + out.noShows + ' no-show(s)' : '') + (notified ? ' · ' + notified + ' absent email(s) sent' : '') + (certsSent ? ' · ' + certsSent + ' level certificate(s) sent' : '')
      : 'Saved ' + out.saved + ' with ' + out.errors.length + ' error(s)';
    return out;
  } catch (e) {
    Logger.log('saveBulkCurriculumAssessments error: ' + e);
    return { success: false, saved: 0, noShows: 0, errors: [{ message: String(e.message || e) }], message: 'Error: ' + e.message };
  }
}

/**
 * Return the most-recent saved scores for a set of group riders so the
 * scorer can pre-highlight what was scored last time (review / edit).
 * items: [{ rowIndex, keNo, bookingId }]
 * Returns: { <rowIndex>: { safety, riding, knowledge, attitude, notes } }
 * ASSESSMENT columns: 0 ts, 1 bookingId, 2 keNo, 3 safety, 4 riding,
 *                     5 knowledge, 6 attitude, 7 avg, 8 passFail, 9 notes
 */
function getGroupSavedScores(items) {
  var out = {};
  try {
    if (!items || !items.length) return out;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aSheet = ss.getSheetByName('ASSESSMENT');
    if (!aSheet || aSheet.getLastRow() < 2) return out;
    var data = aSheet.getDataRange().getValues();
    items.forEach(function (it) {
      var ke = String(it.keNo || '').trim();
      var bid = String(it.bookingId || '').trim();
      if (!ke) return;
      for (var i = data.length - 1; i >= 1; i--) {
        if (String(data[i][2] || '').trim() !== ke) continue;
        if (bid && String(data[i][1] || '').trim() !== bid) continue;
        if (String(data[i][8] || '').trim().toLowerCase() === 'outdated') continue;
        out[it.rowIndex] = {
          safety: Number(data[i][3] || 0),
          riding: Number(data[i][4] || 0),
          knowledge: Number(data[i][5] || 0),
          attitude: Number(data[i][6] || 0),
          notes: String(data[i][9] || '')
        };
        break;
      }
    });
  } catch (e) {
    Logger.log('getGroupSavedScores error: ' + e);
  }
  return out;
}

function saveCurriculumAssessment(rowIndex, assessmentData) {
  // assessmentData: { safety, riding, knowledge, attitude, trainerNotes, present,
  //                   keNo, level, classNumber }
  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var sched  = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (!sched) throw new Error('Schedule sheet not found');

    var keNo       = String(assessmentData.keNo       || '').trim();
    var level      = String(assessmentData.level      || '').trim();
    var classNo    = String(assessmentData.classNumber|| '').trim();
    var safety     = Number(assessmentData.safety     || 0);
    var riding     = Number(assessmentData.riding     || 0);
    var knowledge  = Number(assessmentData.knowledge  || 0);
    var attitude   = Number(assessmentData.attitude   || 0);
    var present    = assessmentData.present !== false;
    var notes      = String(assessmentData.trainerNotes || '').trim();

    // 1. Mark attendance on Schedule sheet
    var attendanceVal = present ? 'Present' : 'No-Show';
    sched.getRange(rowIndex, CONFIG.SCHED_COLS.ATTENDANCE + 1).setValue(attendanceVal);
    if (notes) sched.getRange(rowIndex, CONFIG.SCHED_COLS.STAFF_NOTES + 1).setValue(notes);

    if (!present) {
      var missedNote = '';
      if (level && classNo) {
        var missedTitle = _getTitleForClass(ss, level, classNo);
        missedNote = 'Missed class: ' + level + ' · Class ' + classNo + (missedTitle ? ' — ' + missedTitle : '');
      }
      if (missedNote) {
        var prevNote = String(sched.getRange(rowIndex, CONFIG.SCHED_COLS.STAFF_NOTES + 1).getValue() || '').trim();
        sched.getRange(rowIndex, CONFIG.SCHED_COLS.STAFF_NOTES + 1).setValue(prevNote ? prevNote + ' · ' + missedNote : missedNote);
      }
      return { success: true, passFail: 'Absent', avg: 0, message: 'Marked as No-Show' };
    }

    // 2. Calculate scores
    var avg      = Math.round(((safety + riding + knowledge + attitude) / 4) * 100) / 100;
    var passFail = avg >= 2 ? 'Pass' : 'Repeat';

    var bookingId = String(assessmentData.bookingId || '').trim();
    if (!bookingId) bookingId = _findBookingIdForSession_(keNo, _getRiderName(ss, keNo));
    if (!bookingId) {
      return { success: false, message: 'Booking_ID is required for assessment linkage.' };
    }
    if (!_isBookingOwnedByRider_(bookingId, keNo)) {
      return { success: false, message: 'Booking_ID does not belong to this rider.' };
    }

    // 3. Write to ASSESSMENT sheet
    var aSheet = ss.getSheetByName('ASSESSMENT');
    if (aSheet) {
      aSheet.appendRow([
        new Date(), bookingId || '', keNo, safety, riding, knowledge, attitude, avg, passFail, notes
      ]);
    }

    // 4. Write to PROGRESS_LOG
    var progSheet = ss.getSheetByName('PROGRESS_LOG');
    if (progSheet) {
      var cur      = getCurriculumWithProgress(keNo);
      var sequence = _getSequenceForClass(ss, level, classNo);
      var nextSeq  = passFail === 'Pass' ? sequence + 1 : sequence;
      var nextItem = _getItemBySequence(ss, nextSeq);

      progSheet.appendRow([
        new Date(),              // Timestamp
        bookingId || '',         // Booking_ID
        keNo,                    // Student_ID / KE
        _getRiderName(ss, keNo), // Student_Name
        '',                      // Program
        sequence,                // Sequence
        level,                   // Level
        classNo,                 // Class_Number
        _getTitleForClass(ss, level, classNo), // Title
        'Yes',                   // Present
        avg,                     // Avg_Score
        passFail,                // Pass_Fail
        passFail === 'Pass' ? 'Move to next class' : 'Repeat class',
        nextSeq,                 // Next_Sequence
        notes                    // Trainer_Notes
      ]);
    }

    return {
      success  : true,
      passFail : passFail,
      avg      : avg,
      safety   : safety,
      riding   : riding,
      knowledge: knowledge,
      attitude : attitude,
      message  : passFail === 'Pass'
        ? '✅ PASS! Avg ' + avg + ' — advancing to next class'
        : '🔁 REPEAT — Avg ' + avg + ' (need ≥ 2 to pass)'
    };
  } catch (e) {
    Logger.log('saveCurriculumAssessment error: ' + e);
    return { success: false, message: 'Error: ' + e.message };
  }
}

/**
 * Score a missed/absent class from the Riders tab (makeup session).
 * payload: { keNo, scheduleRowIndex, level, classNumber, title, bookingId,
 *            safety, riding, knowledge, attitude, trainerNotes }
 */
function saveMakeupCurriculumAssessment(payload) {
  payload = payload || {};
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (payload.synthetic === true || String(payload.source || '') === 'late-join') {
      return _saveLateJoinMakeupAssessment_(ss, payload);
    }
    var sched = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (!sched) throw new Error('Schedule sheet not found');

    var keNo = String(payload.keNo || '').trim();
    var rowIndex = Number(payload.scheduleRowIndex || 0);
    var level = String(payload.level || '').trim();
    var classNo = String(payload.classNumber || '').trim();
    var safety = Number(payload.safety || 0);
    var riding = Number(payload.riding || 0);
    var knowledge = Number(payload.knowledge || 0);
    var attitude = Number(payload.attitude || 0);
    var notes = String(payload.trainerNotes || 'Makeup class').trim();

    if (!keNo || !rowIndex || !level || !classNo) {
      return { success: false, message: 'Missing rider, session, or class info.' };
    }
    if (!safety || !riding || !knowledge || !attitude) {
      return { success: false, message: 'All four skill scores (1–4) are required.' };
    }

    var rowData = sched.getRange(rowIndex, 1, 1, sched.getLastColumn()).getValues()[0];
    var rowKe = String(rowData[CONFIG.SCHED_COLS.KE_NO] || '').trim();
    if (rowKe !== keNo) return { success: false, message: 'Session row does not match this rider.' };

    var att = String(rowData[CONFIG.SCHED_COLS.ATTENDANCE] || '').trim().toLowerCase();
    if (att !== 'no-show' && att !== 'noshow' && att !== 'no show') {
      return { success: false, message: 'This session is not marked as absent.' };
    }

    var staffNote = String(rowData[CONFIG.SCHED_COLS.STAFF_NOTES] || '').trim();
    if (staffNote.toLowerCase().indexOf('makeup scored') >= 0) {
      return { success: false, message: 'Makeup already scored for this absent day.' };
    }

    var avg = Math.round(((safety + riding + knowledge + attitude) / 4) * 100) / 100;
    var passFail = avg >= 2 ? 'Pass' : 'Repeat';

    var bookingId = String(payload.bookingId || '').trim();
    if (!bookingId) bookingId = _findBookingIdForSession_(keNo, _getRiderName(ss, keNo));
    if (!bookingId) return { success: false, message: 'Booking_ID is required for assessment linkage.' };
    if (!_isBookingOwnedByRider_(bookingId, keNo)) {
      return { success: false, message: 'Booking_ID does not belong to this rider.' };
    }

    var aSheet = ss.getSheetByName('ASSESSMENT');
    if (aSheet) {
      aSheet.appendRow([
        new Date(), bookingId, keNo, safety, riding, knowledge, attitude, avg, passFail,
        notes + ' (makeup)'
      ]);
    }

    var progSheet = ss.getSheetByName('PROGRESS_LOG');
    if (progSheet) {
      var sequence = _getSequenceForClass(ss, level, classNo);
      var nextSeq = passFail === 'Pass' ? sequence + 1 : sequence;
      progSheet.appendRow([
        new Date(),
        bookingId,
        keNo,
        _getRiderName(ss, keNo),
        '',
        sequence,
        level,
        classNo,
        payload.title || _getTitleForClass(ss, level, classNo),
        'Yes',
        avg,
        passFail,
        passFail === 'Pass' ? 'Makeup — advance' : 'Makeup — repeat',
        nextSeq,
        notes
      ]);
    }

    if (typeof _updateStudentProgress === 'function') {
      var seq = _getSequenceForClass(ss, level, classNo);
      var nextSeq2 = passFail === 'Pass' ? seq + 1 : seq;
      _updateStudentProgress(keNo, passFail, seq, nextSeq2);
    }
    if (typeof _refreshStudentMetrics === 'function') _refreshStudentMetrics(keNo);

    // A passing makeup may complete the rider's level → email their certificate.
    var makeupCertSent = false;
    if (passFail === 'Pass' && typeof processLevelCompletionCertificates_ === 'function') {
      try {
        var mkEmail = String(rowData[CONFIG.SCHED_COLS.EMAIL] || '').trim();
        makeupCertSent = processLevelCompletionCertificates_([{
          keNo: keNo,
          name: _getRiderName(ss, keNo),
          email: mkEmail,
          level: level
        }]) > 0;
      } catch (certErr) {
        Logger.log('makeup level certificate error: ' + certErr);
      }
    }

    var scorer = String(payload.scoredBy || '').trim();
    var makeupTag = 'Makeup scored ' + fmtDate(new Date()) + (scorer ? ' by ' + scorer : '');
    sched.getRange(rowIndex, CONFIG.SCHED_COLS.STAFF_NOTES + 1).setValue(
      staffNote ? staffNote + ' · ' + makeupTag : makeupTag
    );
    if (scorer) {
      try {
        if (typeof _ensureScheduleAuditColumns_ === 'function') _ensureScheduleAuditColumns_();
        sched.getRange(rowIndex, CONFIG.SCHED_COLS.SCORED_BY + 1).setValue(scorer);
      } catch (se) {}
    }

    if (typeof invalidateAttendanceCaches === 'function') invalidateAttendanceCaches();

    return {
      success: true,
      passFail: passFail,
      avg: avg,
      message: passFail === 'Pass'
        ? 'Makeup PASS — avg ' + avg + ' · class completed' + (makeupCertSent ? ' · level certificate emailed' : '')
        : 'Makeup REPEAT — avg ' + avg
    };
  } catch (e) {
    Logger.log('saveMakeupCurriculumAssessment error: ' + e);
    return { success: false, message: 'Error: ' + e.message };
  }
}

/**
 * Score a curriculum gap that predates the rider's registration.
 * This intentionally does not create or mutate Schedule attendance rows.
 */
function _saveLateJoinMakeupAssessment_(ss, payload) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return { success: false, message: 'Another score is being saved. Please try again.' };
  }
  try {
    var keNo = String(payload.keNo || '').trim();
    var level = String(payload.level || '').trim();
    var classNo = String(payload.classNumber || '').trim();
    var requestedKey = String(payload.classKey || (level + '|' + classNo)).trim();
    var safety = Number(payload.safety || 0);
    var riding = Number(payload.riding || 0);
    var knowledge = Number(payload.knowledge || 0);
    var attitude = Number(payload.attitude || 0);
    if (!keNo || !level || !classNo || requestedKey !== level + '|' + classNo) {
      return { success: false, message: 'Missing or invalid late-join class information.' };
    }
    if (!safety || !riding || !knowledge || !attitude) {
      return { success: false, message: 'All four skill scores (1–4) are required.' };
    }
    if ([safety, riding, knowledge, attitude].some(function (v) { return v < 1 || v > 4; })) {
      return { success: false, message: 'Skill scores must be between 1 and 4.' };
    }

    // Recompute eligibility under the lock; never trust a stale browser item.
    var cache = _loadCurriculumEngineCache_(ss);
    if (_isClassPassed_(cache.passedByKe[keNo] || {}, level, classNo)) {
      return { success: false, message: 'This class is already passed.' };
    }
    if (typeof _getLateJoinCurriculumGaps_ !== 'function') {
      return { success: false, message: 'Late-join validation is unavailable.' };
    }
    var context = _buildLateJoinGapContext_(ss, cache);
    var representedByAbsence = typeof _getRealNoShowAssignedClassKeys_ === 'function'
      ? _getRealNoShowAssignedClassKeys_(ss, keNo, cache) : {};
    var gaps = _getLateJoinCurriculumGaps_(ss, keNo, cache, context, representedByAbsence);
    var gap = null;
    for (var g = 0; g < gaps.length; g++) {
      if (gaps[g].classKey === requestedKey) { gap = gaps[g]; break; }
    }
    if (!gap) {
      return { success: false, message: 'This class is no longer an eligible late-join makeup.' };
    }

    var bookingId = String(payload.bookingId || '').trim();
    if (!bookingId) bookingId = cache.bookingByKe[keNo] || '';
    if (!bookingId) bookingId = _findBookingIdForSession_(keNo, _getRiderName(ss, keNo));
    if (!bookingId) return { success: false, message: 'Booking_ID is required for assessment linkage.' };
    if (!_isBookingOwnedByRider_(bookingId, keNo)) {
      return { success: false, message: 'Booking_ID does not belong to this rider.' };
    }

    var avg = Math.round(((safety + riding + knowledge + attitude) / 4) * 100) / 100;
    var passFail = avg >= 2 ? 'Pass' : 'Repeat';
    var title = gap.title || _getTitleForClass(ss, level, classNo);
    var scorer = String(payload.scoredBy || '').trim();
    var userNotes = String(payload.trainerNotes || '').trim();
    var auditNote = 'Late joiner makeup · Grade ' + gap.grade + ' · Section ' + gap.section;
    if (gap.cohortCompletedDate) auditNote += ' · cohort completed ' + gap.cohortCompletedDate;
    if (scorer) auditNote += ' · scored by ' + scorer;
    if (userNotes) auditNote += ' · ' + userNotes;

    var aSheet = ss.getSheetByName('ASSESSMENT');
    var progSheet = ss.getSheetByName('PROGRESS_LOG');
    if (!aSheet || !progSheet) {
      return { success: false, message: 'ASSESSMENT or PROGRESS_LOG sheet not found.' };
    }
    aSheet.appendRow([
      new Date(), bookingId, keNo, safety, riding, knowledge, attitude, avg, passFail,
      auditNote
    ]);

    var sequence = _getSequenceForClass(ss, level, classNo);
    var nextSeq = passFail === 'Pass' ? sequence + 1 : sequence;
    progSheet.appendRow([
      new Date(),
      bookingId,
      keNo,
      _getRiderName(ss, keNo),
      '',
      sequence,
      level,
      classNo,
      title,
      'Yes',
      avg,
      passFail,
      passFail === 'Pass' ? 'Late joiner makeup — advance' : 'Late joiner makeup — repeat',
      nextSeq,
      auditNote
    ]);

    _updateLateJoinStudentProgress_(ss, keNo, passFail, sequence, nextSeq);
    if (typeof _refreshStudentMetrics === 'function') _refreshStudentMetrics(keNo);

    var certSent = false;
    if (passFail === 'Pass' && typeof processLevelCompletionCertificates_ === 'function') {
      try {
        var email = typeof _studentEmail === 'function' ? _studentEmail(keNo) : '';
        certSent = processLevelCompletionCertificates_([{
          keNo: keNo,
          name: _getRiderName(ss, keNo),
          email: email,
          level: level
        }]) > 0;
      } catch (certErr) {
        Logger.log('late join makeup certificate error: ' + certErr);
      }
    }

    if (typeof invalidateAttendanceCaches === 'function') invalidateAttendanceCaches();
    return {
      success: true,
      passFail: passFail,
      avg: avg,
      synthetic: true,
      message: passFail === 'Pass'
        ? 'Late-join makeup PASS — avg ' + avg + ' · class completed' +
          (certSent ? ' · level certificate emailed' : '')
        : 'Late-join makeup REPEAT — avg ' + avg + ' · remains in missed list'
    };
  } catch (e) {
    Logger.log('_saveLateJoinMakeupAssessment_ error: ' + e);
    return { success: false, message: 'Error: ' + e.message };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/**
 * Update legacy STUDENTS sequence counters without moving a rider backwards
 * when an older late-join class is made up.
 */
function _updateLateJoinStudentProgress_(ss, keNo, passFail, sequence, nextSeq) {
  if (passFail !== 'Pass') return;
  var sh = ss.getSheetByName('STUDENTS');
  if (!sh || sh.getLastRow() < 2) return;
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() !== keNo) continue;
    var row = i + 1;
    var completedSeq = Number(data[i][6] || 0);
    var currentSeq = Number(data[i][7] || 0);
    var completedCount = Number(data[i][8] || 0);
    sh.getRange(row, 7).setValue(Math.max(completedSeq, Number(sequence || 0)));
    sh.getRange(row, 8).setValue(Math.max(currentSeq, Number(nextSeq || 0)));
    sh.getRange(row, 9).setValue(completedCount + 1);
    return;
  }
}

function _parseMissedClassFromNotes_(notes) {
  var s = String(notes || '');
  var m = s.match(/Missed class:\s*([^·]+)\s*·\s*Class\s*([^\s—]+)(?:\s*—\s*(.+))?/i);
  if (!m) return null;
  return {
    level: String(m[1] || '').trim(),
    classNumber: String(m[2] || '').trim(),
    title: String(m[3] || '').trim()
  };
}

/**
 * Parse either a "Class done:" (present/scored) or "Missed class:" (no-show)
 * marker from a schedule row's staff notes. Returns the delivered class so
 * the session can be pinned to what was actually taught that day.
 */
function _parseDeliveredClassFromNotes_(notes) {
  var s = String(notes || '');
  var done = s.match(/Class done:\s*([^·]+)\s*·\s*Class\s*([^\s—·]+)(?:\s*—\s*([^·]+))?/i);
  if (done) {
    return {
      kind: 'done',
      level: String(done[1] || '').trim(),
      classNumber: String(done[2] || '').trim(),
      title: String(done[3] || '').trim()
    };
  }
  var missed = s.match(/Missed class:\s*([^·]+)\s*·\s*Class\s*([^\s—·]+)(?:\s*—\s*([^·]+))?/i);
  if (missed) {
    return {
      kind: 'missed',
      level: String(missed[1] || '').trim(),
      classNumber: String(missed[2] || '').trim(),
      title: String(missed[3] || '').trim()
    };
  }
  return null;
}

/** Remove any "Class done:" / "Missed class:" marker, returning leftover notes. */
function _stripClassMarkers_(notes) {
  var s = String(notes || '');
  s = s.replace(/Class done:\s*[^·]+·\s*Class\s*[^\s—·]+(?:\s*—\s*[^·]+)?/gi, '');
  s = s.replace(/Missed class:\s*[^·]+·\s*Class\s*[^\s—·]+(?:\s*—\s*[^·]+)?/gi, '');
  s = s.replace(/Curriculum class:\s*[^·]+·\s*Class\s*[^\s—·]+(?:\s*—\s*[^·]+)?/gi, '');
  s = s.replace(/^\s*·\s*/, '').replace(/\s*·\s*$/, '').trim();
  return s;
}

/**
 * Trainer override: "Curriculum class: Level · Class N — Title"
 * Pins an unfinished session to a chosen class (before/after cohort align).
 */
function _parseCurriculumOverrideFromNotes_(notes) {
  var s = String(notes || '');
  var m = s.match(/Curriculum class:\s*([^·]+)\s*·\s*Class\s*([^\s—·]+)(?:\s*—\s*([^·]+))?/i);
  if (!m) return null;
  return {
    level: String(m[1] || '').trim(),
    classNumber: String(m[2] || '').trim(),
    title: String(m[3] || '').trim(),
    kind: 'override'
  };
}

/** Find the curriculum item object for a level + class from the cache. */
function _findCurriculumItemFromCache_(cache, level, classNo) {
  if (!cache || !cache.curriculumItems) return null;
  var lv = String(level || '').trim();
  var cn = String(classNo || '').trim();
  for (var i = 0; i < cache.curriculumItems.length; i++) {
    var it = cache.curriculumItems[i];
    if (String(it.level).trim() === lv && String(it.classNumber).trim() === cn) return it;
  }
  return null;
}

function _isClassPassed_(passedMap, level, classNo) {
  return !!passedMap[String(level || '').trim() + '|' + String(classNo || '').trim()];
}

function _getCurriculumItemsOrdered_(ss) {
  var cur = ss.getSheetByName('CURRICULUM') || ss.getSheetByName('circulum') || ss.getSheetByName('Curriculum');
  if (!cur) return [];
  var data = cur.getDataRange().getValues();
  var items = [], seq = 0;
  for (var i = 1; i < data.length; i++) {
    var level = String(data[i][0] || '').trim();
    var classNo = String(data[i][1] || '').trim();
    var title = String(data[i][2] || '').trim();
    if (!level && !classNo && !title) continue;
    seq++;
    items.push({ sequence: seq, level: level, classNumber: classNo, title: title });
  }
  return items;
}

function _buildPassedMapForRider_(ss, keNo) {
  var passed = {};
  var prog = ss.getSheetByName('PROGRESS_LOG');
  if (!prog || prog.getLastRow() < 2) return passed;
  var data = prog.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2] || '').trim() !== keNo) continue;
    if (!_isActiveProgressPass_(data[i][11])) continue;
    var lv = String(data[i][6] || '').trim();
    var cn = String(data[i][7] || '').trim();
    if (lv && cn) passed[lv + '|' + cn] = true;
  }
  return passed;
}

function _getStudentGradeSection_(ss, keNo) {
  var sh = ss.getSheetByName('STUDENTS');
  if (!sh || sh.getLastRow() < 2) return { grade: '', section: '' };
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var gCol = -1, sCol = -1;
  for (var h = 0; h < headers.length; h++) {
    var key = String(headers[h] || '').trim().toLowerCase();
    if (key === 'grade') gCol = h;
    if (key === 'section') sCol = h;
  }
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() !== keNo) continue;
    return {
      grade: gCol >= 0 ? String(data[i][gCol] || '').trim() : '',
      section: sCol >= 0 ? String(data[i][sCol] || '').trim() : ''
    };
  }
  return { grade: '', section: '' };
}

function _findBookingIdForSession_(keNo, riderName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var students = ss.getSheetByName('STUDENTS');
    var bookings = ss.getSheetByName('BOOKINGS');
    if (!students || !bookings) return '';
    var targetKe = String(keNo || '').trim();
    var targetName = String(riderName || '').trim().toLowerCase();
    var studentId = '';
    var sData = students.getDataRange().getValues();
    if (targetKe) {
      for (var k = 1; k < sData.length; k++) {
        if (String(sData[k][0] || '').trim() === targetKe) {
          studentId = targetKe;
          break;
        }
      }
    }
    if (!studentId && targetName) {
      for (var i = 1; i < sData.length; i++) {
        if (String(sData[i][1] || '').trim().toLowerCase() === targetName) {
          studentId = String(sData[i][0] || '').trim();
          break;
        }
      }
    }
    if (!studentId) return '';
    var bData = bookings.getDataRange().getValues();
    for (var r = bData.length - 1; r >= 1; r--) {
      if (String(bData[r][2] || '').trim() === studentId) {
        return String(bData[r][1] || '').trim();
      }
    }
    return '';
  } catch (e) {
    Logger.log('_findBookingIdForSession_ error: ' + e);
    return '';
  }
}

function _isBookingOwnedByRider_(bookingId, keNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var bookings = ss.getSheetByName('BOOKINGS');
    if (!bookings || bookings.getLastRow() < 2) return false;
    var data = bookings.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1] || '').trim() !== String(bookingId || '').trim()) continue;
      var studentId = String(data[i][2] || '').trim();
      if (!studentId) return false;
      return String(_lookupKENoFromStudentId_(studentId) || '').trim() === String(keNo || '').trim();
    }
    return false;
  } catch (e) {
    Logger.log('_isBookingOwnedByRider_ error: ' + e);
    return false;
  }
}

function _dashYmd_(value) {
  if (!value) return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var text = String(value).trim();
  var match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return match[1] + '-' + match[2] + '-' + match[3];
  var date = new Date(value);
  return isNaN(date.getTime()) ? '' : Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function _dashTodayYmd_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getTrainingDashboardData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var bookings = ss.getSheetByName('BOOKINGS');
    var assess = ss.getSheetByName('ASSESSMENT');
    var progress = ss.getSheetByName('PROGRESS_LOG');
    var today = _dashTodayYmd_();
    var out = {
      asOf: today,
      totalBookings: 0,
      assessed: 0,
      passCount: 0,
      repeatCount: 0,
      avgScore: 0,
      recentTrend: [],
      levelProgress: [],
      totalRiders: 0,
      ridersPresent: 0,
      ridersScheduledToday: 0,
      totalHorses: 0,
      horseStatus: { Active: 0, Leased: 0, Rehab: 0, Lame: 0, Retired: 0 },
      totalStaff: 0,
      staffPresent: 0,
      staffOnLeave: 0,
      feedStock: { items: [], itemCount: 0, lowCount: 0 },
      tackStock: { items: [], itemCount: 0, lowCount: 0 },
      alerts: []
    };

    if (bookings && bookings.getLastRow() > 1) out.totalBookings = bookings.getLastRow() - 1;

    if (assess && assess.getLastRow() > 1) {
      var aData = assess.getDataRange().getValues();
      var sum = 0, cnt = 0;
      for (var i = 1; i < aData.length; i++) {
        var avg = Number(aData[i][7] || 0);
        var pf = String(aData[i][8] || '').trim().toLowerCase();
        if (avg > 0) { sum += avg; cnt++; }
        if (pf === 'pass') out.passCount++;
        if (pf === 'repeat') out.repeatCount++;
        out.recentTrend.push({
          date: aData[i][0] ? fmtDate(new Date(aData[i][0])) : '',
          avg: avg
        });
      }
      out.assessed = cnt;
      out.avgScore = cnt ? Math.round((sum / cnt) * 100) / 100 : 0;
      out.recentTrend = out.recentTrend.slice(-8);
    }

    if (progress && progress.getLastRow() > 1) {
      var pData = progress.getDataRange().getValues();
      var map = {};
      for (var r = 1; r < pData.length; r++) {
        var lv = String(pData[r][6] || '').trim() || 'Unknown';
        if (!map[lv]) map[lv] = { level: lv, pass: 0, repeat: 0 };
        var st = String(pData[r][11] || '').trim().toLowerCase();
        if (st === 'pass') map[lv].pass++;
        else if (st === 'repeat') map[lv].repeat++;
      }
      out.levelProgress = Object.keys(map).map(function (k) { return map[k]; });
    }

    // Total riders
    var ridersSheet = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
    if (ridersSheet && ridersSheet.getLastRow() > 1) {
      out.totalRiders = ridersSheet.getLastRow() - 1;
    }

    // Riders present / scheduled today (Schedule sheet)
    var schedule = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (schedule && schedule.getLastRow() > 1) {
      var sData = schedule.getDataRange().getValues();
      var presentKe = {};
      var scheduledKe = {};
      for (var s = 1; s < sData.length; s++) {
        if (_dashYmd_(sData[s][CONFIG.SCHED_COLS.DATE]) !== today) continue;
        var ke = String(sData[s][CONFIG.SCHED_COLS.KE_NO] || '').trim();
        if (!ke) continue;
        scheduledKe[ke] = true;
        var att = String(sData[s][CONFIG.SCHED_COLS.ATTENDANCE] || '').trim().toLowerCase();
        if (att === 'present') presentKe[ke] = true;
      }
      out.ridersScheduledToday = Object.keys(scheduledKe).length;
      out.ridersPresent = Object.keys(presentKe).length;
    }

    // Horses
    try {
      if (typeof _ensureHorsesSheet_ === 'function') {
        var horseSheet = _ensureHorsesSheet_();
        if (horseSheet.getLastRow() > 1) {
          var hData = horseSheet.getDataRange().getValues();
          var hc = CONFIG.HORSE_COLS;
          for (var h = 1; h < hData.length; h++) {
            if (!String(hData[h][hc.KE_HORSE_ID] || '').trim() && !String(hData[h][hc.NAME] || '').trim()) continue;
            out.totalHorses++;
            var hStatus = String(hData[h][hc.STATUS] || 'Active').trim() || 'Active';
            if (out.horseStatus[hStatus] == null) out.horseStatus[hStatus] = 0;
            out.horseStatus[hStatus]++;
            if (hStatus === 'Lame') {
              out.alerts.push({ severity: 'high', message: String(hData[h][hc.NAME] || 'Horse') + ' is marked Lame.' });
            } else if (hStatus === 'Rehab') {
              out.alerts.push({ severity: 'medium', message: String(hData[h][hc.NAME] || 'Horse') + ' is in Rehab.' });
            }
            if (typeof _horseHealthSummary_ === 'function') {
              var careMeta = {
                vaccinationStatus: hData[h][hc.VACCINATION_STATUS],
                vaccinationPostponedTo: _dashYmd_(hData[h][hc.VACCINATION_POSTPONED_TO]),
                dewormingStatus: hData[h][hc.DEWORMING_STATUS],
                dewormingPostponedTo: _dashYmd_(hData[h][hc.DEWORMING_POSTPONED_TO]),
                farrierStatus: hData[h][hc.FARRIER_STATUS],
                farrierPostponedTo: _dashYmd_(hData[h][hc.FARRIER_POSTPONED_TO])
              };
              if (!String(hData[h][hc.VACCINATION_STATUS] || '').trim() && _dashYmd_(hData[h][hc.VACCINATION_DATE])) {
                careMeta.vaccinationStatus = 'Done';
              }
              if (!String(hData[h][hc.DEWORMING_STATUS] || '').trim() && _dashYmd_(hData[h][hc.DEWORMING_DATE])) {
                careMeta.dewormingStatus = 'Done';
              }
              if (!String(hData[h][hc.FARRIER_STATUS] || '').trim() && _dashYmd_(hData[h][hc.FARRIER_DATE])) {
                careMeta.farrierStatus = 'Done';
              }
              var health = _horseHealthSummary_(
                _dashYmd_(hData[h][hc.VACCINATION_DATE]),
                _dashYmd_(hData[h][hc.DEWORMING_DATE]),
                _dashYmd_(hData[h][hc.FARRIER_DATE]),
                careMeta
              );
              if (health.vaccination && health.vaccination.status === 'Not Done') {
                out.alerts.push({
                  severity: 'high',
                  message: String(hData[h][hc.NAME] || 'Horse') + ': Vaccination not done.'
                });
              }
              if (health.deworming && health.deworming.status === 'Not Done') {
                out.alerts.push({
                  severity: 'medium',
                  message: String(hData[h][hc.NAME] || 'Horse') + ': Deworming not done.'
                });
              }
              if (health.farrier && health.farrier.status === 'Not Done') {
                out.alerts.push({
                  severity: 'medium',
                  message: String(hData[h][hc.NAME] || 'Horse') + ': Farrier not done.'
                });
              }
              if (health.nearest && health.nearest.daysUntil != null && health.nearest.daysUntil < 0
                && health.nearest.status !== 'Not Done') {
                out.alerts.push({
                  severity: 'high',
                  message: String(hData[h][hc.NAME] || 'Horse') + ': ' + health.nearest.label
                    + ' — ' + health.nearest.dueLabel + '.'
                });
              } else if (health.nearest && health.nearest.daysUntil != null && health.nearest.daysUntil <= 7
                && health.nearest.status !== 'Not Done') {
                out.alerts.push({
                  severity: 'medium',
                  message: String(hData[h][hc.NAME] || 'Horse') + ': ' + health.nearest.label
                    + ' ' + health.nearest.dueLabel + '.'
                });
              }
            }
          }
        }
      }
    } catch (horseErr) {
      Logger.log('dashboard horses: ' + horseErr);
    }

    // Staff present today
    try {
      if (typeof _ensureGroomerSheets_ === 'function') {
        var staffEnv = _ensureGroomerSheets_();
        var noUniform = [];
        if (staffEnv.groomers.getLastRow() > 1) {
          var gData = staffEnv.groomers.getDataRange().getValues();
          var gc = CONFIG.GROOMER_COLS;
          for (var g = 1; g < gData.length; g++) {
            var gStatus = String(gData[g][gc.STATUS] || gData[g][gc.ACTIVE] || 'Active').trim().toLowerCase();
            if (gStatus !== 'active' && gStatus !== 'yes' && gStatus !== 'true' && gStatus !== '1') continue;
            if (!String(gData[g][gc.STAFF_ID] || '').trim()) continue;
            out.totalStaff++;
            var bal = Number(gData[g][gc.LEAVE_BALANCE] || 0);
            if (bal <= 1) {
              out.alerts.push({
                severity: 'medium',
                message: String(gData[g][gc.NAME] || 'Staff') + ' leave balance is low (' + bal + ').'
              });
            }
            if (gc.UNIFORM_ISSUED_DATE != null && !gData[g][gc.UNIFORM_ISSUED_DATE]) {
              noUniform.push(String(gData[g][gc.NAME] || 'Staff'));
            }
          }
          // One grouped alert — a line per person would push out real warnings
          if (noUniform.length) {
            out.alerts.push({
              severity: 'info',
              message: 'Uniform issue date missing for ' + noUniform.length + ' staff: '
                + noUniform.slice(0, 4).join(', ')
                + (noUniform.length > 4 ? ' and ' + (noUniform.length - 4) + ' more' : '') + '.'
            });
          }
        }
        if (staffEnv.attendance.getLastRow() > 1) {
          var attData = staffEnv.attendance.getDataRange().getValues();
          var ac = CONFIG.GROOMER_ATTENDANCE_COLS;
          for (var a = 1; a < attData.length; a++) {
            if (_dashYmd_(attData[a][ac.DATE]) !== today) continue;
            var aStatus = String(attData[a][ac.STATUS] || '').trim();
            if (aStatus === 'Present') out.staffPresent++;
            if (aStatus === 'Leave') out.staffOnLeave++;
          }
        }
      }
    } catch (staffErr) {
      Logger.log('dashboard staff: ' + staffErr);
    }

    // Feed / tack stock
    try {
      if (typeof getStockSummaries_ === 'function') {
        var stock = getStockSummaries_();
        out.feedStock = stock.feed;
        out.tackStock = stock.tack;
        var unconfirmedFeed = [];
        var shortfallFeed = [];
        var countNeededFeed = [];
        (stock.feed.items || []).forEach(function (item) {
          if (item.requiresCount) countNeededFeed.push(item.name || item.itemId);
          if (item.needsConfirm) {
            unconfirmedFeed.push(item.name + ' (' + Number(item.pendingDays || 0) + 'd)');
            if (item.shortfall) shortfallFeed.push(item.name);
          }
          var checkQty = item.needsConfirm ? Number(item.projectedQuantity) : Number(item.quantity || 0);
          var checkLow = item.needsConfirm ? !!item.projectedLow : !!item.low;
          if (checkLow) {
            out.alerts.push({
              severity: 'high',
              message: 'Feed running out: ' + item.name + ' at ' + (item.location || '—')
                + ' — ' + checkQty + ' ' + (item.unit || '') + ' left against a minimum of ' + item.minLevel
                + (item.needsConfirm ? ' (based on estimated daily use).' : '.')
            });
          } else if (item.daysLeft != null && item.daysLeft <= 3) {
            out.alerts.push({
              severity: 'medium',
              message: 'Feed: ' + item.name + ' has about ' + item.daysLeft + ' day(s) left'
                + (item.needsConfirm ? ', assuming the estimated daily use is right.' : '.')
            });
          }
        });
        if (countNeededFeed.length) {
          out.alerts.push({
            severity: 'high',
            message: 'Feed figures cannot be trusted for ' + countNeededFeed.join(', ')
              + ' — nothing has been entered for weeks. Weigh what is actually left and record it '
              + 'as a stock count in the Feed tab.'
          });
        }
        // Grouped so a week of missed entries does not bury the stock warnings
        if (unconfirmedFeed.length) {
          out.alerts.push({
            severity: 'medium',
            message: 'Daily feed entry missing for ' + unconfirmedFeed.length + ' item(s): '
              + unconfirmedFeed.slice(0, 4).join(', ')
              + (unconfirmedFeed.length > 4 ? ' and ' + (unconfirmedFeed.length - 4) + ' more' : '')
              + '. Figures shown are estimates until confirmed in the Feed tab.'
          });
        }
        if (shortfallFeed.length) {
          out.alerts.push({
            severity: 'high',
            message: 'Feed records do not add up for ' + shortfallFeed.join(', ')
              + ' — the estimated use is more than the stock on record, so a delivery was probably never entered.'
          });
        }
        (stock.tack.items || []).forEach(function (item) {
          if (Number(item.quantity || 0) <= 0) {
            out.alerts.push({
              severity: 'high',
              message: 'Tack out of stock: ' + item.name + ' at ' + (item.location || '—') + '.'
            });
          } else if (item.low) {
            out.alerts.push({
              severity: 'medium',
              message: 'Tack low: ' + item.name + ' at ' + (item.location || '—')
                + ' — ' + item.quantity + ' left against a minimum of ' + item.minLevel + '.'
            });
          }
        });
      }
    } catch (stockErr) {
      Logger.log('dashboard stock: ' + stockErr);
    }

    // Cap alerts so the UI stays readable
    var severityRank = { high: 0, medium: 1, info: 2 };
    out.alerts.sort(function (x, y) {
      return (severityRank[x.severity] || 9) - (severityRank[y.severity] || 9);
    });
    if (out.alerts.length > 20) {
      var hidden = out.alerts.length - 20;
      out.alerts = out.alerts.slice(0, 20);
      out.alerts.push({
        severity: 'info',
        message: hidden + ' more alert(s) not shown — clear the ones above, or open the Feed / Tack / Horses tabs for the full list.'
      });
    }
    if (!out.alerts.length) {
      out.alerts.push({ severity: 'info', message: 'No alerts right now. Stable looks clear.' });
    }

    return out;
  } catch (e) {
    Logger.log('getTrainingDashboardData error: ' + e);
    return {
      asOf: '',
      totalBookings: 0, assessed: 0, passCount: 0, repeatCount: 0, avgScore: 0,
      recentTrend: [], levelProgress: [],
      totalRiders: 0, ridersPresent: 0, ridersScheduledToday: 0,
      totalHorses: 0, horseStatus: { Active: 0, Leased: 0, Rehab: 0, Lame: 0, Retired: 0 },
      totalStaff: 0, staffPresent: 0, staffOnLeave: 0,
      feedStock: { items: [], itemCount: 0, lowCount: 0 },
      tackStock: { items: [], itemCount: 0, lowCount: 0 },
      alerts: [{ severity: 'high', message: 'Dashboard error: ' + String(e.message || e) }]
    };
  }
}

// ────────────────────────────────────────────────────────────
//  HELPER: sequence / title lookups
// ────────────────────────────────────────────────────────────

function _getSequenceForClass(ss, level, classNo) {
  var cur = ss.getSheetByName('CURRICULUM') || ss.getSheetByName('circulum') || ss.getSheetByName('Curriculum');
  if (!cur) return 0;
  var data = cur.getDataRange().getValues();
  var seq  = 0;
  for (var i = 1; i < data.length; i++) {
    var lv = String(data[i][0] || '').trim();
    var cn = String(data[i][1] || '').trim();
    if (!lv && !cn) continue;
    seq++;
    if (lv === level && cn === classNo) return seq;
  }
  return seq;
}

function _getTitleForClass(ss, level, classNo) {
  var cur = ss.getSheetByName('CURRICULUM') || ss.getSheetByName('circulum') || ss.getSheetByName('Curriculum');
  if (!cur) return '';
  var data = cur.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === level && String(data[i][1] || '').trim() === classNo)
      return String(data[i][2] || '').trim();
  }
  return '';
}

function _getItemBySequence(ss, seq) {
  var cur = ss.getSheetByName('CURRICULUM') || ss.getSheetByName('circulum') || ss.getSheetByName('Curriculum');
  if (!cur || seq < 1) return null;
  var data = cur.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    var lv = String(data[i][0] || '').trim();
    var cn = String(data[i][1] || '').trim();
    if (!lv && !cn) continue;
    count++;
    if (count === seq) return { level: lv, classNumber: cn, title: String(data[i][2] || '').trim() };
  }
  return null;
}

function _getRiderName(ss, keNo) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
  if (!sheet) return '';
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][CONFIG.RIDER_COLS.KE_NO] || '').trim() === keNo)
      return String(data[i][CONFIG.RIDER_COLS.NAME] || '');
  }
  return '';
}

// ────────────────────────────────────────────────────────────
//  PORTAL: enriched rider data including full curriculum
// ────────────────────────────────────────────────────────────

function getRiderDataWithCurriculum(identifier) {
  var base = getRiderData(identifier);
  if (!base || !base.found || base.multiProfile) return base;
  try {
    base.curriculum = getCurriculumWithProgress(base.keNo);
  } catch (e) {
    base.curriculum = { items: [], passedCount: 0, totalCount: 0 };
  }
  return base;
}

// ────────────────────────────────────────────────────────────
//  SKILL RADAR DATA — for portal charts
// ────────────────────────────────────────────────────────────

function getSkillRadarData(keNo) {
  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var aSheet = ss.getSheetByName('ASSESSMENT');
    if (!aSheet || aSheet.getLastRow() < 2) return null;
    var data   = aSheet.getDataRange().getValues();
    // Cols: 0=Timestamp, 2=KE, 3=Safety, 4=Riding, 5=Knowledge, 6=Attitude
    var totals = { safety: 0, riding: 0, knowledge: 0, attitude: 0, count: 0 };
    var trend  = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2] || '').trim() !== keNo) continue;
      if (String(data[i][8] || '').trim().toLowerCase() === 'outdated') continue;
      var s = Number(data[i][3] || 0), r = Number(data[i][4] || 0);
      var k = Number(data[i][5] || 0), a = Number(data[i][6] || 0);
      if (!s && !r && !k && !a) continue;
      totals.safety    += s; totals.riding   += r;
      totals.knowledge += k; totals.attitude += a;
      totals.count++;
      trend.push({ date: data[i][0] ? fmtDate(new Date(data[i][0])) : '', avg: Number(data[i][7] || 0) });
    }
    if (!totals.count) return null;
    var n = totals.count;
    return {
      radar: {
        safety   : Math.round((totals.safety    / n) * 100) / 100,
        riding   : Math.round((totals.riding    / n) * 100) / 100,
        knowledge: Math.round((totals.knowledge / n) * 100) / 100,
        attitude : Math.round((totals.attitude  / n) * 100) / 100
      },
      trend: trend.slice(-10) // last 10 sessions
    };
  } catch (e) {
    Logger.log('getSkillRadarData error: ' + e);
    return null;
  }
}

// ============================================================
//  CLASS SHIFT (back / forward on session group)
// ============================================================

/**
 * Ordered curriculum list for the Class shift modal.
 */
function listCurriculumClassesForShift() {
  try {
    var cache = _loadCurriculumEngineCache_();
    var items = (cache.curriculumItems || []).map(function (it) {
      return {
        level: it.level,
        classNumber: it.classNumber,
        title: it.title,
        sequence: it.sequence,
        label: it.level + ' · Class ' + it.classNumber + (it.title ? ' — ' + it.title : '')
      };
    });
    return { success: true, items: items };
  } catch (e) {
    return { success: false, items: [], message: String(e.message || e) };
  }
}

/**
 * Move a session group's curriculum class back or forward.
 *
 * Example: shift to Class 2 → this session teaches Class 2;
 * prior Pass rows for Class 2+ become Outdated so the next class is 3
 * after this session is passed. Shifting forward assumes Pass for
 * skipped earlier classes, then outdates Class N+ so N is current.
 *
 * opts: { rowIndexes, level, classNumber, reason, actor }
 */
function shiftGroupCurriculumClass(opts) {
  opts = opts || {};
  try {
    var rows = (typeof _normalizeRowIndexes_ === 'function')
      ? _normalizeRowIndexes_(opts.rowIndexes)
      : (opts.rowIndexes || []).map(Number).filter(Boolean);
    if (!rows.length) return { success: false, message: 'No sessions selected.' };

    var targetLevel = String(opts.level || '').trim();
    var targetClass = String(opts.classNumber || '').trim();
    if (!targetLevel || !targetClass) {
      return { success: false, message: 'Pick a curriculum class.' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sched = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
    if (!sched) return { success: false, message: 'Schedule sheet not found.' };

    var cache = _loadCurriculumEngineCache_(ss);
    var targetKey = targetLevel + '|' + targetClass;
    var targetSeq = Number(cache.seqByKey[targetKey] || 0);
    if (!targetSeq) {
      return { success: false, message: 'That class was not found in CURRICULUM.' };
    }
    var targetTitle = cache.titleByKey[targetKey] || '';
    var nextItem = null;
    for (var ni = 0; ni < (cache.curriculumItems || []).length; ni++) {
      if (Number(cache.curriculumItems[ni].sequence) === targetSeq + 1) {
        nextItem = cache.curriculumItems[ni];
        break;
      }
    }
    var nextLabel = nextItem
      ? (nextItem.level + ' · Class ' + nextItem.classNumber)
      : '(end of curriculum)';

    var reason = String(opts.reason || '').trim();
    var actor = String(opts.actor || '').trim();
    var lastRow = sched.getLastRow();
    var keNos = [];
    var touched = 0;

    rows.forEach(function (r) {
      if (r < 2 || r > lastRow) return;
      var ke = String(sched.getRange(r, CONFIG.SCHED_COLS.KE_NO + 1).getValue() || '').trim();
      if (ke && keNos.indexOf(ke) < 0) keNos.push(ke);

      var noteCell = sched.getRange(r, CONFIG.SCHED_COLS.STAFF_NOTES + 1);
      var prev = String(noteCell.getValue() || '').trim();
      var carry = _stripClassMarkers_(prev);
      var pin = 'Curriculum class: ' + targetLevel + ' · Class ' + targetClass
        + (targetTitle ? ' — ' + targetTitle : '');
      var stamp = 'Class shifted to ' + targetLevel + ' · Class ' + targetClass
        + (reason ? ' — ' + reason : '')
        + (actor ? ' (by ' + actor + ')' : '');
      noteCell.setValue([pin, stamp, carry].filter(Boolean).join(' · '));
      touched++;
    });

    if (!keNos.length) {
      return { success: false, message: 'No riders found on those sessions.' };
    }

    var progResult = _shiftProgressForRiders_(ss, cache, keNos, targetLevel, targetClass, targetSeq, reason, actor);
    var assessOutdated = _outdateAssessmentsForRiders_(ss, keNos, progResult.matchKeys || []);

    try { _updateBookingsAssignedClass_(ss, keNos, targetLevel, targetClass, targetTitle, targetSeq); } catch (ignoreBk) {}

    if (typeof invalidateAttendanceCaches === 'function') invalidateAttendanceCaches();
    try {
      if (typeof _EXEC_MEMO === 'object' && _EXEC_MEMO) _EXEC_MEMO.curriculumCache = null;
    } catch (ignoreMemo) {}

    return {
      success: true,
      shifted: touched,
      riders: keNos.length,
      outdated: progResult.outdated || 0,
      assumedPasses: progResult.assumed || 0,
      assessmentsOutdated: assessOutdated,
      level: targetLevel,
      classNumber: targetClass,
      nextClassLabel: nextLabel,
      message: 'Moved to ' + targetLevel + ' · Class ' + targetClass
        + '. This session teaches that class; next class will be ' + nextLabel + '.'
        + (progResult.outdated ? (' · ' + progResult.outdated + ' prior mark(s) outdated') : '')
        + (progResult.assumed ? (' · ' + progResult.assumed + ' earlier class(es) marked passed to advance') : '')
    };
  } catch (e) {
    Logger.log('shiftGroupCurriculumClass error: ' + e);
    return { success: false, message: 'Error: ' + String(e.message || e) };
  }
}

/**
 * For each rider: outdate Pass/Repeat for sequence >= target;
 * when jumping forward, add assumed Pass rows for missing earlier classes.
 */
function _shiftProgressForRiders_(ss, cache, keNos, targetLevel, targetClass, targetSeq, reason, actor) {
  var prog = ss.getSheetByName('PROGRESS_LOG');
  if (!prog) return { outdated: 0, assumed: 0, matchKeys: [] };

  var data = prog.getLastRow() > 0 ? prog.getDataRange().getValues() : [['']];
  var outdated = 0;
  var assumed = 0;
  var matchKeys = [];
  var keSet = {};
  keNos.forEach(function (k) { keSet[String(k).trim()] = true; });

  for (var i = 1; i < data.length; i++) {
    var ke = String(data[i][2] || '').trim();
    if (!keSet[ke]) continue;
    var pf = String(data[i][11] || '').trim().toLowerCase();
    if (pf !== 'pass' && pf !== 'repeat') continue;
    var lv = String(data[i][6] || '').trim();
    var cn = String(data[i][7] || '').trim();
    var seq = Number(cache.seqByKey[lv + '|' + cn] || data[i][5] || 0);
    if (!seq || seq < targetSeq) continue;

    prog.getRange(i + 1, 12).setValue('Outdated');
    var prevAction = String(data[i][12] || '').trim();
    var action = 'Outdated by class shift → ' + targetLevel + ' · Class ' + targetClass
      + (reason ? ' (' + reason + ')' : '');
    prog.getRange(i + 1, 13).setValue(action);
    var prevNotes = String(data[i][14] || '').trim();
    var noteAdd = 'Was: ' + (prevAction || pf) + (actor ? ' · by ' + actor : '');
    prog.getRange(i + 1, 15).setValue(prevNotes ? prevNotes + ' · ' + noteAdd : noteAdd);
    outdated++;
    matchKeys.push({
      bookingId: String(data[i][1] || '').trim(),
      keNo: ke,
      avg: Number(data[i][10] || 0)
    });
  }

  var now = new Date();
  var assumedRows = [];
  keNos.forEach(function (ke) {
    var passedMap = {};
    for (var p = 1; p < data.length; p++) {
      if (String(data[p][2] || '').trim() !== ke) continue;
      var pf2 = String(data[p][11] || '').trim().toLowerCase();
      var lv2 = String(data[p][6] || '').trim();
      var cn2 = String(data[p][7] || '').trim();
      var seq2 = Number(cache.seqByKey[lv2 + '|' + cn2] || data[p][5] || 0);
      if (seq2 >= targetSeq) continue;
      if (pf2 === 'pass') passedMap[lv2 + '|' + cn2] = true;
    }
    (cache.curriculumItems || []).forEach(function (it) {
      if (Number(it.sequence) >= targetSeq) return;
      var key = it.level + '|' + it.classNumber;
      if (passedMap[key]) return;
      var bookingId = cache.bookingByKe[ke] || '';
      assumedRows.push([
        now,
        bookingId,
        ke,
        cache.nameByKe[ke] || ke,
        '',
        it.sequence,
        it.level,
        it.classNumber,
        it.title || '',
        'No',
        '',
        'Pass',
        'Advanced by class shift',
        targetSeq,
        'Assumed pass so current class is ' + targetLevel + ' · Class ' + targetClass
          + (reason ? ' — ' + reason : '')
          + (actor ? ' (by ' + actor + ')' : '')
      ]);
      passedMap[key] = true;
      assumed++;
    });
  });

  if (assumedRows.length) {
    var start = prog.getLastRow() + 1;
    prog.getRange(start, 1, assumedRows.length, assumedRows[0].length).setValues(assumedRows);
  }

  return { outdated: outdated, assumed: assumed, matchKeys: matchKeys };
}

/** Mark matching ASSESSMENT rows Outdated so old scores are not reused. */
function _outdateAssessmentsForRiders_(ss, keNos, matchKeys) {
  var aSheet = ss.getSheetByName('ASSESSMENT');
  if (!aSheet || aSheet.getLastRow() < 2) return 0;
  var data = aSheet.getDataRange().getValues();
  var keSet = {};
  keNos.forEach(function (k) { keSet[String(k).trim()] = true; });
  var n = 0;
  var hasMatches = (matchKeys || []).length > 0;
  for (var i = 1; i < data.length; i++) {
    var ke = String(data[i][2] || '').trim();
    if (!keSet[ke]) continue;
    var pf = String(data[i][8] || '').trim().toLowerCase();
    if (pf === 'outdated') continue;
    if (pf !== 'pass' && pf !== 'repeat') continue;
    var match = false;
    if (hasMatches) {
      for (var m = 0; m < matchKeys.length; m++) {
        if (matchKeys[m].keNo === ke) { match = true; break; }
      }
    } else {
      match = true;
    }
    if (!match) continue;
    aSheet.getRange(i + 1, 9).setValue('Outdated');
    var notes = String(data[i][9] || '').trim();
    aSheet.getRange(i + 1, 10).setValue(
      notes ? notes + ' · Outdated by class shift' : 'Outdated by class shift'
    );
    n++;
  }
  return n;
}

function _updateBookingsAssignedClass_(ss, keNos, level, classNo, title, sequence) {
  var bookings = ss.getSheetByName('BOOKINGS');
  if (!bookings || bookings.getLastRow() < 2) return;
  var headers = bookings.getRange(1, 1, 1, bookings.getLastColumn()).getValues()[0];
  function col(name) {
    var want = String(name || '').trim().toLowerCase();
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i] || '').trim().toLowerCase() === want) return i + 1;
    }
    return 0;
  }
  var studentCol = col('Student_ID') || 3;
  var levelCol = col('Assigned_Level');
  var classCol = col('Assigned_Class_Number');
  var titleCol = col('Assigned_Title');
  var seqCol = col('Assigned_Sequence');
  if (!levelCol && !classCol) return;
  var data = bookings.getDataRange().getValues();
  var keSet = {};
  keNos.forEach(function (k) { keSet[String(k).trim()] = true; });
  for (var r = 1; r < data.length; r++) {
    var sid = String(data[r][studentCol - 1] || '').trim();
    if (!keSet[sid]) continue;
    if (levelCol) bookings.getRange(r + 1, levelCol).setValue(level);
    if (classCol) bookings.getRange(r + 1, classCol).setValue(classNo);
    if (titleCol) bookings.getRange(r + 1, titleCol).setValue(title || '');
    if (seqCol) bookings.getRange(r + 1, seqCol).setValue(sequence || '');
  }
}