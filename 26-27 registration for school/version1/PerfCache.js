// ============================================================
// KINGS EQUESTRIAN — PERFORMANCE / CACHING LAYER
// File: PerfCache.gs
// Chunked CacheService helpers + cache warming + invalidation.
// Designed for ~500 students where full-sheet reads are expensive.
//
// Strategy:
//   • Heavy read endpoints (riders list, today/tomorrow sessions) store
//     their computed JSON in the shared script cache.
//   • A time trigger (warmAttendanceCaches) refreshes them off the
//     critical path so interactive calls are near-instant.
//   • Every write path calls invalidateAttendanceCaches() so data is
//     never stale after a save/booking/cancel/move.
//   • All caching is best-effort: on any miss/error we compute live.
// ============================================================

var PERF = {
  TTL_RIDERS: 1800,        // 30 min (also refreshed by trigger + invalidated on write)
  TTL_SESSIONS: 900,       // 15 min
  KEY_RIDERS: 'ke_riders_stats_v1',
  KEY_SESSIONS_PREFIX: 'ke_sessions_v1_',   // + today | tomorrow
  CHUNK: 90000             // < 100KB CacheService per-value limit
};

// ── Per-execution memo (avoids re-reading sheets within one request) ──
var _EXEC_MEMO = { curriculumCache: null };

// ────────────────────────────────────────────────────────────
//  CHUNKED CACHE HELPERS
// ────────────────────────────────────────────────────────────

function _cachePutLarge_(cache, baseKey, str, ttl) {
  try {
    var n = Math.ceil(str.length / PERF.CHUNK) || 1;
    var parts = {};
    parts[baseKey + ':meta'] = String(n);
    for (var i = 0; i < n; i++) parts[baseKey + ':' + i] = str.substr(i * PERF.CHUNK, PERF.CHUNK);
    cache.putAll(parts, ttl);
    return true;
  } catch (e) {
    Logger.log('_cachePutLarge_ error: ' + e);
    return false;
  }
}

function _cacheGetLarge_(cache, baseKey) {
  try {
    var meta = cache.get(baseKey + ':meta');
    if (!meta) return null;
    var n = Number(meta);
    if (!n) return null;
    var keys = [];
    for (var i = 0; i < n; i++) keys.push(baseKey + ':' + i);
    var got = cache.getAll(keys);
    var s = '';
    for (var j = 0; j < n; j++) {
      var p = got[baseKey + ':' + j];
      if (p == null) return null; // a chunk expired — treat as miss
      s += p;
    }
    return s;
  } catch (e) {
    Logger.log('_cacheGetLarge_ error: ' + e);
    return null;
  }
}

function _cacheRemoveLarge_(cache, baseKey) {
  try {
    var meta = cache.get(baseKey + ':meta');
    var keys = [baseKey + ':meta'];
    var n = meta ? Number(meta) : 24; // clear generously even if meta missing
    for (var i = 0; i < n; i++) keys.push(baseKey + ':' + i);
    cache.removeAll(keys);
  } catch (e) {
    Logger.log('_cacheRemoveLarge_ error: ' + e);
  }
}

function _scriptCache_() {
  try { return CacheService.getScriptCache(); } catch (e) { return null; }
}

// ────────────────────────────────────────────────────────────
//  INVALIDATION  (call after any write that affects riders/sessions)
// ────────────────────────────────────────────────────────────

function invalidateAttendanceCaches() {
  var cache = _scriptCache_();
  if (!cache) return;
  _cacheRemoveLarge_(cache, PERF.KEY_RIDERS);
  _cacheRemoveLarge_(cache, PERF.KEY_SESSIONS_PREFIX + 'today');
  _cacheRemoveLarge_(cache, PERF.KEY_SESSIONS_PREFIX + 'tomorrow');
  _EXEC_MEMO.curriculumCache = null;
}

// ────────────────────────────────────────────────────────────
//  RIDERS LIST — cached wrapper
// ────────────────────────────────────────────────────────────

/** Compute the full riders-with-curriculum payload (heavy). */
function _computeRidersWithCurriculum_() {
  return getAllRidersWithStats_WithCurriculum_UNCACHED();
}

/** Cached entry point used by the attendance PWA. */
function getRidersWithStatsCached() {
  var cache = _scriptCache_();
  if (cache) {
    var hit = _cacheGetLarge_(cache, PERF.KEY_RIDERS);
    if (hit) {
      try { return JSON.parse(hit); } catch (e) {}
    }
  }
  var data = _computeRidersWithCurriculum_();
  if (cache) {
    try { _cachePutLarge_(cache, PERF.KEY_RIDERS, JSON.stringify(data), PERF.TTL_RIDERS); } catch (e) {}
  }
  return data;
}

// ────────────────────────────────────────────────────────────
//  SESSIONS — cached wrapper (today / tomorrow only)
// ────────────────────────────────────────────────────────────

function _sessionCacheKeyFor_(dateArg) {
  var a = String(dateArg || 'today').trim().toLowerCase();
  if (a === 'today' || a === 'tomorrow') return PERF.KEY_SESSIONS_PREFIX + a;
  return null; // custom dates are not cached (rare, and hard to invalidate)
}

/**
 * Re-attach curriculum document links onto sessions (cheap).
 * Survives stale Script Cache entries built before docLink existed.
 */
function _enrichSessionsDocLinks_(sessions) {
  if (!sessions || !sessions.length) return sessions || [];
  var missing = false;
  for (var i = 0; i < sessions.length; i++) {
    if (!String(sessions[i].docLink || '').trim()) { missing = true; break; }
  }
  if (!missing) return sessions;
  try {
    var curCache = _loadCurriculumEngineCache_();
    sessions.forEach(function (s) {
      if (String(s.docLink || '').trim()) return;
      var key = String(s.curriculumLevel || '') + '|' + String(s.curriculumClassNo || '');
      var fromItem = s.curriculumItem && s.curriculumItem.docLink
        ? String(s.curriculumItem.docLink).trim() : '';
      var fromMap = curCache.docLinkByKey
        ? String(curCache.docLinkByKey[key] || '').trim() : '';
      s.docLink = fromItem || fromMap || '';
      if (s.docLink && s.curriculumItem && !s.curriculumItem.docLink) {
        s.curriculumItem.docLink = s.docLink;
      }
    });
  } catch (e) {}
  return sessions;
}

function getSessionsForDateCached(dateArg) {
  var key = _sessionCacheKeyFor_(dateArg);
  var cache = key ? _scriptCache_() : null;
  if (cache) {
    var hit = _cacheGetLarge_(cache, key);
    if (hit) {
      try {
        return _enrichSessionsDocLinks_(JSON.parse(hit));
      } catch (e) {}
    }
  }
  var data = getSessionsForDate_Curriculum(dateArg);
  if (cache) {
    try { _cachePutLarge_(cache, key, JSON.stringify(data), PERF.TTL_SESSIONS); } catch (e) {}
  }
  return data;
}

// ────────────────────────────────────────────────────────────
//  CACHE WARMING  (time trigger — keeps hot data ready)
// ────────────────────────────────────────────────────────────

function warmAttendanceCaches() {
  var cache = _scriptCache_();
  if (!cache) return;
  try {
    var riders = _computeRidersWithCurriculum_();
    _cachePutLarge_(cache, PERF.KEY_RIDERS, JSON.stringify(riders), PERF.TTL_RIDERS);
  } catch (e) { Logger.log('warm riders error: ' + e); }
  try {
    var today = getSessionsForDate_Curriculum('today');
    _cachePutLarge_(cache, PERF.KEY_SESSIONS_PREFIX + 'today', JSON.stringify(today), PERF.TTL_SESSIONS);
  } catch (e) { Logger.log('warm today error: ' + e); }
  try {
    var tom = getSessionsForDate_Curriculum('tomorrow');
    _cachePutLarge_(cache, PERF.KEY_SESSIONS_PREFIX + 'tomorrow', JSON.stringify(tom), PERF.TTL_SESSIONS);
  } catch (e) { Logger.log('warm tomorrow error: ' + e); }
}

/** Menu helper: install a 10-minute cache-warming trigger. */
function installCacheWarmTrigger() {
  var exists = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'warmAttendanceCaches';
  });
  if (!exists) {
    ScriptApp.newTrigger('warmAttendanceCaches').timeBased().everyMinutes(10).create();
  }
  warmAttendanceCaches();
  try {
    SpreadsheetApp.getUi().alert('Cache warming enabled (every 10 min) and caches primed.');
  } catch (e) {}
}

// ────────────────────────────────────────────────────────────
//  SCHEDULE ARCHIVING
//  Moves rows dated before a cutoff into a "Schedule Archive" tab so
//  the live Schedule stays small and fast to read.
//
//  ⚠️ History-dependent stats (classes attended, pending make-ups,
//  portal "My Sessions") read the LIVE Schedule. Only archive rows old
//  enough to be outside the current academic year — the default cutoff
//  is 12 months ago, so within-year data is never touched.
// ────────────────────────────────────────────────────────────

function archiveOldScheduleRows(beforeYMD) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sched = ss.getSheetByName(CONFIG.SHEETS.SCHEDULE);
  if (!sched) return { success: false, message: 'Schedule sheet not found.' };

  var cutoff = new Date(String(beforeYMD || '') + 'T00:00:00');
  if (isNaN(cutoff.getTime())) return { success: false, message: 'Invalid cutoff date (use yyyy-MM-dd).' };

  var lastRow = sched.getLastRow();
  var lastCol = sched.getLastColumn();
  if (lastRow < 2) return { success: true, archived: 0, kept: 0, message: 'Nothing to archive.' };

  var values = sched.getRange(1, 1, lastRow, lastCol).getValues();
  var header = values[0];
  var keep = [], archive = [];
  for (var i = 1; i < values.length; i++) {
    var d = values[i][CONFIG.SCHED_COLS.DATE];
    var isOld = false;
    if (d) { try { isOld = new Date(d) < cutoff; } catch (e) { isOld = false; } }
    if (isOld) archive.push(values[i]); else keep.push(values[i]);
  }
  if (!archive.length) return { success: true, archived: 0, kept: keep.length, message: 'No rows older than ' + beforeYMD + '.' };

  var arch = ss.getSheetByName('Schedule Archive');
  if (!arch) {
    arch = ss.insertSheet('Schedule Archive');
    arch.getRange(1, 1, 1, header.length).setValues([header]);
    arch.setFrozenRows(1);
  }
  var aStart = arch.getLastRow() + 1;
  arch.getRange(aStart, 1, archive.length, archive[0].length).setValues(archive);
  try { arch.getRange(aStart, CONFIG.SCHED_COLS.DATE + 1, archive.length, 1).setNumberFormat('dd-MMM-yyyy'); } catch (e) {}

  // Rewrite the live sheet: clear data region, write only the kept rows.
  sched.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  if (keep.length) {
    sched.getRange(2, 1, keep.length, keep[0].length).setValues(keep);
    try { sched.getRange(2, CONFIG.SCHED_COLS.DATE + 1, keep.length, 1).setNumberFormat('dd-MMM-yyyy'); } catch (e) {}
  }

  if (typeof invalidateAttendanceCaches === 'function') invalidateAttendanceCaches();
  return {
    success: true,
    archived: archive.length,
    kept: keep.length,
    message: 'Archived ' + archive.length + ' row(s) older than ' + beforeYMD + '. Live Schedule now has ' + keep.length + ' row(s).'
  };
}

/** Menu wrapper — prompts for a cutoff date (defaults to 12 months ago). */
function archiveOldScheduleMenu() {
  var ui = SpreadsheetApp.getUi();
  var tz = Session.getScriptTimeZone();
  var d = new Date(); d.setFullYear(d.getFullYear() - 1);
  var suggested = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  var resp = ui.prompt(
    'Archive old Schedule rows',
    'Move sessions dated BEFORE this date (yyyy-MM-dd) to "Schedule Archive".\n\n' +
    'Keep this within old academic years — within-year data is needed for stats.\n\n' +
    'Suggested: ' + suggested,
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var val = String(resp.getResponseText() || '').trim() || suggested;
  var res = archiveOldScheduleRows(val);
  ui.alert(res.message || (res.success ? 'Done.' : 'Failed.'));
}

/** Monthly trigger target — archives rows older than 12 months automatically. */
function archiveOldScheduleAuto() {
  var tz = Session.getScriptTimeZone();
  var d = new Date(); d.setFullYear(d.getFullYear() - 1);
  archiveOldScheduleRows(Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
}
