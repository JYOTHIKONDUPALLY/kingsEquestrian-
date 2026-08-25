// ============================================================
// KINGS EQUESTRIAN — FEED & TACK STOCK + MOVEMENTS
// ============================================================

var FEED_HEADERS = [
  'Item ID', 'Item Name', 'Location', 'Quantity', 'Unit', 'Min Level',
  'Consumed Per Day', 'Notes', 'Updated At', 'Updated By', 'Pack Size (kg)', 'Usage Mode'
];
var FEED_MOVEMENT_HEADERS = [
  'Movement ID', 'Item ID', 'Type', 'Date', 'Qty', 'Notes', 'Recorded At', 'Recorded By',
  'Ref Horse ID', 'Ref Horse Name'
];
var TACK_HEADERS = [
  'Item ID', 'Item Name', 'Category', 'Model', 'Vendor', 'Location',
  'Qty', 'Min Level', 'Photo URL', 'Photo File ID', 'Notes', 'Updated At', 'Updated By'
];
var TACK_MOVEMENT_HEADERS = [
  'Movement ID', 'Item ID', 'Type', 'Date', 'Qty', 'Reason',
  'Photo URL', 'Photo File ID', 'From Location', 'To Location', 'Recorded At', 'Recorded By'
];

var FEED_MOVEMENT_TYPES = ['Restock', 'Consume', 'Adjust', 'Use', 'Expire'];
var FEED_USAGE_MODES = ['Regular', 'Occasional'];
var TACK_MOVEMENT_TYPES = ['Wear Out', 'Restock', 'Transfer'];
/**
 * How many past days (ending yesterday) to backfill if EOD consume was skipped.
 * Beyond this the estimate is not trustworthy, so the item is locked until
 * someone enters a physical stock count (see _pendingFeedDefaultsForItem_).
 */
var FEED_DEFAULT_LOOKBACK_DAYS = 14;
/** Movements older than this are moved to the archive sheet to keep reads fast. */
var STOCK_ARCHIVE_AFTER_DAYS = 365;
/** Row count that triggers the daily auto-archive check. */
var STOCK_ARCHIVE_ROW_TRIGGER = 3000;

function _styleStockHeader_(sheet, headers) {
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1f4e3d').setFontColor('#fff').setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function _stockYmd_(value) {
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

function _stockDate_(ymdValue) {
  var ymd = _stockYmd_(ymdValue);
  if (!ymd) return '';
  var parts = ymd.split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 12, 0, 0);
}

function _requireStockTrainer_(username, token) {
  var trainer = validateTrainerToken(username, token);
  if (!trainer || !trainer.valid) throw new Error('Your trainer session has expired. Sign in again.');
  return trainer;
}

function _ensureNamedSheet_(sheetName, headers, migrateFn) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _styleStockHeader_(sheet, headers);
  } else if (typeof migrateFn === 'function') {
    migrateFn(sheet);
  } else {
    var existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    headers.forEach(function (header, i) {
      if (!String(existing[i] || '').trim()) sheet.getRange(1, i + 1).setValue(header);
    });
    _styleStockHeader_(sheet, headers);
  }
  return sheet;
}

/** Migrate old FEED_STOCK (Item ID, Name, Qty, Unit, Min, Notes…) into location-aware layout. */
function _migrateFeedStockSheet_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h || '').trim();
  });
  var hasLocation = headers.indexOf('Location') >= 0;
  var hasConsumed = headers.indexOf('Consumed Per Day') >= 0;
  if (hasLocation && hasConsumed && headers[0] === 'Item ID') {
    if (sheet.getMaxColumns() < FEED_HEADERS.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), FEED_HEADERS.length - sheet.getMaxColumns());
    }
    sheet.getRange(1, 1, 1, FEED_HEADERS.length).setValues([FEED_HEADERS]);
    _styleStockHeader_(sheet, FEED_HEADERS);
    return;
  }
  // Legacy: Item ID, Item Name, Quantity, Unit, Min Level, Notes, Updated At, Updated By
  var data = sheet.getLastRow() > 0 ? sheet.getDataRange().getValues() : [];
  var out = [FEED_HEADERS];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!String(r[0] || '').trim() && !String(r[1] || '').trim()) continue;
    out.push([
      String(r[0] || '').trim(),
      String(r[1] || '').trim(),
      'Hyderabad',
      Number(r[2] || 0),
      String(r[3] || '').trim() || 'bags',
      Number(r[4] || 0),
      0,
      String(r[5] || '').trim(),
      r[6] || '',
      String(r[7] || '').trim(),
      0
    ]);
  }
  sheet.clear();
  if (sheet.getMaxColumns() < FEED_HEADERS.length) {
    sheet.insertColumnsAfter(Math.max(sheet.getMaxColumns(), 1), FEED_HEADERS.length - Math.max(sheet.getMaxColumns(), 1));
  }
  sheet.getRange(1, 1, out.length, FEED_HEADERS.length).setValues(out);
  _styleStockHeader_(sheet, FEED_HEADERS);
}

/** Migrate old TACK_STOCK into category/model/vendor/location layout. */
function _migrateTackStockSheet_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h || '').trim();
  });
  var hasCategory = headers.indexOf('Category') >= 0;
  var hasModel = headers.indexOf('Model') >= 0;
  if (hasCategory && hasModel && headers[0] === 'Item ID') {
    if (sheet.getMaxColumns() < TACK_HEADERS.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), TACK_HEADERS.length - sheet.getMaxColumns());
    }
    sheet.getRange(1, 1, 1, TACK_HEADERS.length).setValues([TACK_HEADERS]);
    _styleStockHeader_(sheet, TACK_HEADERS);
    return;
  }
  var data = sheet.getLastRow() > 0 ? sheet.getDataRange().getValues() : [];
  var out = [TACK_HEADERS];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!String(r[0] || '').trim() && !String(r[1] || '').trim()) continue;
    out.push([
      String(r[0] || '').trim(),
      String(r[1] || '').trim(),
      '',
      '',
      '',
      'Hyderabad',
      Number(r[2] || 0),
      Number(r[4] || 0),
      '',
      '',
      String(r[5] || '').trim(),
      r[6] || '',
      String(r[7] || '').trim()
    ]);
  }
  sheet.clear();
  if (sheet.getMaxColumns() < TACK_HEADERS.length) {
    sheet.insertColumnsAfter(Math.max(sheet.getMaxColumns(), 1), TACK_HEADERS.length - Math.max(sheet.getMaxColumns(), 1));
  }
  sheet.getRange(1, 1, out.length, TACK_HEADERS.length).setValues(out);
  _styleStockHeader_(sheet, TACK_HEADERS);
}

function _feedStockSeedRows_() {
  return [
    ['FEED-HAY', 'Hay', 'Hyderabad', 40, 'bales', 10, 2, '', '', '', 20, 'Regular'],
    ['FEED-PELLET', 'Pellets', 'Hyderabad', 25, 'bags', 8, 1, '', '', '', 50, 'Regular'],
    ['FEED-BRAN', 'Bran', 'Hyderabad', 12, 'bags', 5, 0.5, '', '', '', 50, 'Regular'],
    ['FEED-SUPP', 'Supplements', 'Hyderabad', 8, 'tubs', 3, 0.1, '', '', '', 0, 'Occasional']
  ];
}

function _tackStockSeedRows_() {
  return [
    ['TACK-SADDLE', 'Saddles', 'HEL', 'PRO1', '', 'Hyderabad', 6, 2, '', '', '', '', ''],
    ['TACK-BRIDLE', 'Bridles', 'HEL', '', '', 'Hyderabad', 8, 3, '', '', '', '', ''],
    ['TACK-GIRTH', 'Girths', 'HEL', '', '', 'Hyderabad', 10, 3, '', '', '', '', ''],
    ['TACK-HELMET', 'School Helmets', 'HEL', '', '', 'Hyderabad', 12, 4, '', '', '', '', ''],
    ['TACK-BOOTS', 'Horse Boots', 'HEL', '', '', 'Hyderabad', 10, 3, '', '', '', '', '']
  ];
}

function _ensureStockSheets_() {
  var feed = _ensureNamedSheet_(CONFIG.SHEETS.FEED_STOCK, FEED_HEADERS, _migrateFeedStockSheet_);
  if (feed.getLastRow() === 1) {
    var feedSeed = _feedStockSeedRows_();
    feed.getRange(2, 1, feedSeed.length, FEED_HEADERS.length).setValues(feedSeed);
  }
  var tack = _ensureNamedSheet_(CONFIG.SHEETS.TACK_STOCK, TACK_HEADERS, _migrateTackStockSheet_);
  if (tack.getLastRow() === 1) {
    var tackSeed = _tackStockSeedRows_();
    tack.getRange(2, 1, tackSeed.length, TACK_HEADERS.length).setValues(tackSeed);
  }
  return {
    feed: feed,
    feedMovements: _ensureNamedSheet_(CONFIG.SHEETS.FEED_MOVEMENTS, FEED_MOVEMENT_HEADERS),
    tack: tack,
    tackMovements: _ensureNamedSheet_(CONFIG.SHEETS.TACK_MOVEMENTS, TACK_MOVEMENT_HEADERS)
  };
}

/**
 * Moves movement rows older than STOCK_ARCHIVE_AFTER_DAYS into an archive sheet.
 * Live reads scan the whole movement sheet, so this keeps the tabs fast as the
 * daily entries build up. Archived rows stay readable in the sheet itself.
 */
function _archiveStockMovementSheet_(sheet, headers, archiveName, cutoffYmd) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var dateCol = archiveName === CONFIG.SHEETS.FEED_MOVEMENTS_ARCHIVE
    ? CONFIG.FEED_MOVEMENT_COLS.DATE
    : CONFIG.TACK_MOVEMENT_COLS.DATE;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  var keep = [];
  var move = [];
  data.forEach(function (row) {
    var ymd = _stockYmd_(row[dateCol]);
    // Rows with an unreadable date are kept, never silently filed away
    if (ymd && ymd < cutoffYmd) move.push(row); else keep.push(row);
  });
  if (!move.length) return 0;
  var archive = _ensureNamedSheet_(archiveName, headers);
  _appendStockRows_(archive, move, headers.length);
  sheet.getRange(2, 1, data.length, headers.length).clearContent();
  if (keep.length) sheet.getRange(2, 1, keep.length, headers.length).setValues(keep);
  return move.length;
}

/** Manual/scheduled archive run. olderThanDays defaults to one year. */
function archiveOldStockMovements(username, token, olderThanDays) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, message: 'Please try again.' };
  try {
    _requireStockTrainer_(username, token);
    var days = Number(olderThanDays || STOCK_ARCHIVE_AFTER_DAYS) || STOCK_ARCHIVE_AFTER_DAYS;
    var sheets = _ensureStockSheets_();
    var cutoff = _feedAddDaysYmd_(_stockYmd_(new Date()), -days);
    var feedMoved = _archiveStockMovementSheet_(
      sheets.feedMovements, FEED_MOVEMENT_HEADERS, CONFIG.SHEETS.FEED_MOVEMENTS_ARCHIVE, cutoff);
    var tackMoved = _archiveStockMovementSheet_(
      sheets.tackMovements, TACK_MOVEMENT_HEADERS, CONFIG.SHEETS.TACK_MOVEMENTS_ARCHIVE, cutoff);
    PropertiesService.getScriptProperties()
      .setProperty('STOCK_ARCHIVE_LAST_RUN', _stockYmd_(new Date()));
    return {
      success: true,
      message: (feedMoved + tackMoved)
        ? ('Archived ' + feedMoved + ' feed and ' + tackMoved + ' tack movement(s) older than '
          + cutoff + '.')
        : 'Nothing older than ' + cutoff + ' to archive.',
      feedMoved: feedMoved,
      tackMoved: tackMoved,
      cutoff: cutoff
    };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/**
 * Runs the archive at most once a day, and only once the movement sheets are big
 * enough to matter. Failures are swallowed: this must never block a stock read.
 */
function _maybeAutoArchiveStock_(sheets) {
  try {
    var rows = sheets.feedMovements.getLastRow() + sheets.tackMovements.getLastRow();
    if (rows < STOCK_ARCHIVE_ROW_TRIGGER) return;
    var props = PropertiesService.getScriptProperties();
    var today = _stockYmd_(new Date());
    if (props.getProperty('STOCK_ARCHIVE_LAST_RUN') === today) return;
    props.setProperty('STOCK_ARCHIVE_LAST_RUN', today);
    var cutoff = _feedAddDaysYmd_(today, -STOCK_ARCHIVE_AFTER_DAYS);
    _archiveStockMovementSheet_(sheets.feedMovements, FEED_MOVEMENT_HEADERS,
      CONFIG.SHEETS.FEED_MOVEMENTS_ARCHIVE, cutoff);
    _archiveStockMovementSheet_(sheets.tackMovements, TACK_MOVEMENT_HEADERS,
      CONFIG.SHEETS.TACK_MOVEMENTS_ARCHIVE, cutoff);
  } catch (e) {
    Logger.log('stock auto-archive skipped: ' + e);
  }
}

function ensureStockSetup(username, token) {
  try {
    _requireStockTrainer_(username, token);
    _ensureStockSheets_();
    return { success: true };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  }
}

function _nextStockId_(sheet, colIndex, prefix) {
  var max = 0;
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, colIndex + 1, sheet.getLastRow(), 1).getValues().forEach(function (row) {
      var id = String(row[0] || '').trim().toUpperCase();
      if (id.indexOf(prefix) !== 0) return;
      max = Math.max(max, Number(id.substring(prefix.length).replace(/\D/g, '')) || 0);
    });
  }
  return prefix + String(max + 1).padStart(3, '0');
}

function _findStockRow_(sheet, colIndex, itemId) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var ids = sheet.getRange(2, colIndex + 1, sheet.getLastRow(), 1).getValues();
  var target = String(itemId || '').trim();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === target) return i + 2;
  }
  return 0;
}

function _getStockPhotosFolder_(subName) {
  var rootName = (typeof CONFIG !== 'undefined' && CONFIG.DRIVE_ROOT_FOLDER) ? CONFIG.DRIVE_ROOT_FOLDER : 'Kings Equestrian';
  var rootIter = DriveApp.getFoldersByName(rootName);
  var root = rootIter.hasNext() ? rootIter.next() : DriveApp.createFolder(rootName);
  var subIter = root.getFoldersByName(subName);
  return subIter.hasNext() ? subIter.next() : root.createFolder(subName);
}

function _stockPhotoViewUrl_(fileId, existingUrl) {
  var url = String(existingUrl || '').trim();
  if (url) return url;
  var id = String(fileId || '').trim();
  return id ? ('https://drive.google.com/uc?export=view&id=' + id) : '';
}

function uploadStockPhoto(payload) {
  payload = payload || {};
  try {
    _requireStockTrainer_(payload.username, payload.token);
    var raw = String(payload.base64Data || payload.data || '').trim();
    if (!raw) throw new Error('No photo selected.');
    var mime = String(payload.mimeType || 'image/jpeg').trim();
    var b64 = raw;
    var match = raw.match(/^data:([^;]+);base64,(.+)$/);
    if (match) {
      mime = match[1];
      b64 = match[2];
    }
    var bytes = Utilities.base64Decode(b64);
    if (!bytes || !bytes.length) throw new Error('Could not read photo data.');
    if (bytes.length > 8 * 1024 * 1024) throw new Error('Photo is too large.');
    var kind = String(payload.kind || 'tack').trim().toLowerCase();
    var folder = _getStockPhotosFolder_(kind === 'feed' ? 'Feed Photos' : 'Tack Photos');
    var label = String(payload.itemName || payload.itemId || kind).trim().replace(/[^\w\-]+/g, '_') || kind;
    var ext = mime.indexOf('png') >= 0 ? 'png' : 'jpg';
    var fileName = kind + '_' + label + '_'
      + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.' + ext;
    var file = folder.createFile(Utilities.newBlob(bytes, mime, fileName));
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      Logger.log('uploadStockPhoto sharing: ' + shareErr);
    }
    // The previous file is intentionally left in Drive: transferred tack rows and
    // wear-out movement records can share a photo, so trashing it would break them.
    return {
      success: true,
      fileId: file.getId(),
      photoUrl: 'https://drive.google.com/uc?export=view&id=' + file.getId(),
      message: 'Photo uploaded.'
    };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  }
}

// ── Feed helpers ─────────────────────────────────────────────
function _feedAddDaysYmd_(ymd, days) {
  var date = _stockDate_(ymd);
  if (!date) return '';
  date.setDate(date.getDate() + Number(days || 0));
  return _stockYmd_(date);
}

function _feedYesterdayYmd_() {
  return _feedAddDaysYmd_(_stockYmd_(new Date()), -1);
}

/**
 * Index of Consume movements by item:
 * { ITEM: { dates: {ymd: qty}, lastDate, lastQty } }
 */
function _buildFeedConsumeIndex_(movementsSheet) {
  var index = {};
  if (!movementsSheet || movementsSheet.getLastRow() < 2) return index;
  var mc = CONFIG.FEED_MOVEMENT_COLS;
  var data = movementsSheet.getDataRange().getValues();
  function slot(itemId) {
    if (!index[itemId]) {
      index[itemId] = { dates: {}, lastDate: '', lastQty: 0, lastCountDate: '' };
    }
    return index[itemId];
  }
  for (var i = 1; i < data.length; i++) {
    var type = String(data[i][mc.TYPE] || '').trim();
    var itemId = String(data[i][mc.ITEM_ID] || '').trim();
    var date = _stockYmd_(data[i][mc.DATE]);
    if (!itemId || !date) continue;
    // A physical stock count re-baselines the item: days before it are settled
    if (type === 'Adjust') {
      var adj = slot(itemId);
      if (!adj.lastCountDate || date > adj.lastCountDate) adj.lastCountDate = date;
      continue;
    }
    if (type !== 'Consume') continue;
    var qty = Number(data[i][mc.QTY] || 0);
    if (!(qty > 0)) continue;
    var entry = slot(itemId);
    // Sum same-day consumes so a split entry still counts as that day's total
    entry.dates[date] = Number(entry.dates[date] || 0) + qty;
    if (!entry.lastDate || date >= entry.lastDate) {
      entry.lastDate = date;
      entry.lastQty = entry.dates[date];
    }
  }
  // Rolling average of the most recent logged days (smooths one-off outliers)
  Object.keys(index).forEach(function (itemId) {
    var dates = Object.keys(index[itemId].dates).sort().reverse().slice(0, 7);
    var sum = 0;
    dates.forEach(function (d) { sum += Number(index[itemId].dates[d] || 0); });
    index[itemId].avgPerDay = dates.length ? Math.round((sum / dates.length) * 100) / 100 : 0;
    index[itemId].loggedDays = Object.keys(index[itemId].dates).length;
  });
  return index;
}

/**
 * Missing EOD consume days through yesterday.
 * Default qty = last confirmed consume (yesterday's entry when available),
 * else the item's Consumed Per Day rate.
 */
function _pendingFeedDefaultsForItem_(item, consumeInfo, throughYmd) {
  var result = { pending: [], skippedOlderDays: 0, defaultQty: 0, source: '', requiresCount: false };
  // Occasional items (vitamins, special meds) are logged when used — never auto-estimated
  if (String(item.usageMode || 'Regular').trim() === 'Occasional') return result;
  var end = throughYmd || _feedYesterdayYmd_();
  if (!end) return result;
  var defaultQty = 0;
  var source = '';
  if (consumeInfo && consumeInfo.lastQty > 0) {
    defaultQty = Number(consumeInfo.lastQty);
    source = 'last_consume';
  } else if (Number(item.consumedPerDay || 0) > 0) {
    defaultQty = Number(item.consumedPerDay);
    source = 'rate';
  } else {
    return result;
  }
  result.defaultQty = defaultQty;
  result.source = source;

  // A stock count settles everything up to its date, so start after whichever
  // is later: the last logged consume or the last physical count.
  var baseline = '';
  if (consumeInfo && consumeInfo.lastDate) baseline = consumeInfo.lastDate;
  if (consumeInfo && consumeInfo.lastCountDate && consumeInfo.lastCountDate > baseline) {
    baseline = consumeInfo.lastCountDate;
  }
  var start = baseline
    ? _feedAddDaysYmd_(baseline, 1)
    : end; // No history at all — only nudge for yesterday, don't invent a backlog
  if (!start || start > end) return result;

  // Beyond the lookback window an estimate is guesswork, so we stop estimating
  // and ask for a physical count instead of quietly understating consumption.
  var earliest = _feedAddDaysYmd_(end, -(FEED_DEFAULT_LOOKBACK_DAYS - 1));
  if (start < earliest) {
    var skipCursor = start;
    var skipGuard = 0;
    while (skipCursor && skipCursor < earliest && skipGuard < 400) {
      skipGuard++;
      if (!(consumeInfo && consumeInfo.dates && consumeInfo.dates[skipCursor])) {
        result.skippedOlderDays++;
      }
      skipCursor = _feedAddDaysYmd_(skipCursor, 1);
    }
    if (result.skippedOlderDays > 0) {
      result.requiresCount = true;
      return result;
    }
    start = earliest;
  }

  var cursor = start;
  var guard = 0;
  while (cursor && cursor <= end && guard < FEED_DEFAULT_LOOKBACK_DAYS + 2) {
    guard++;
    if (!(consumeInfo && consumeInfo.dates && consumeInfo.dates[cursor])) {
      result.pending.push({
        date: cursor,
        suggestedQty: defaultQty,
        source: source,
        isDefault: true,
        sourceLabel: source === 'last_consume'
          ? 'Copied from last consume (' + (consumeInfo.lastDate || 'prior day') + ')'
          : 'Copied from consumed/day rate'
      });
    }
    cursor = _feedAddDaysYmd_(cursor, 1);
  }
  return result;
}

function _enrichFeedItemWithDefaults_(item, consumeIndex) {
  var info = (consumeIndex && consumeIndex[item.itemId]) || null;
  var calc = _pendingFeedDefaultsForItem_(item, info, _feedYesterdayYmd_());
  var pending = calc.pending;
  var pendingQty = 0;
  pending.forEach(function (p) { pendingQty += Number(p.suggestedQty || 0); });
  var projected = Math.max(0, Number(item.quantity || 0) - pendingQty);
  // Burn rate preference: rolling average → last consume → planned rate
  var rate = 0;
  if (info && Number(info.avgPerDay || 0) > 0) rate = Number(info.avgPerDay);
  else if (info && Number(info.lastQty || 0) > 0) rate = Number(info.lastQty);
  else rate = Number(item.consumedPerDay || 0);
  var projectedDaysLeft = rate > 0 ? Math.floor(projected / rate) : null;
  item.lastConsumeDate = info ? info.lastDate : '';
  item.lastConsumeQty = info ? Number(info.lastQty || 0) : 0;
  item.avgPerDay = info ? Number(info.avgPerDay || 0) : 0;
  item.loggedDays = info ? Number(info.loggedDays || 0) : 0;
  item.burnRate = rate;
  item.pendingDefaults = pending;
  item.pendingDays = pending.length;
  item.pendingQtyTotal = Math.round(pendingQty * 100) / 100;
  item.needsConfirm = pending.length > 0;
  item.skippedOlderDays = Number(calc.skippedOlderDays || 0);
  item.requiresCount = !!calc.requiresCount;
  item.lastCountDate = info ? (info.lastCountDate || '') : '';
  item.gapSince = info ? ((info.lastCountDate && info.lastCountDate > (info.lastDate || ''))
    ? info.lastCountDate : (info.lastDate || '')) : '';
  item.projectedQuantity = Math.round(projected * 100) / 100;
  item.projectedDaysLeft = projectedDaysLeft;
  item.projectedLow = projected <= Number(item.minLevel || 0);
  item.shortfall = pendingQty > Number(item.quantity || 0);
  // Prefer projected days-left for "about to finish" when defaults pending
  if (item.needsConfirm && projectedDaysLeft != null) {
    item.daysLeft = projectedDaysLeft;
  } else if (rate > 0) {
    item.daysLeft = Math.floor(Number(item.quantity || 0) / rate);
  }
  return item;
}

function _feedFromRow_(row) {
  var c = CONFIG.FEED_COLS;
  var qty = Number(row[c.QUANTITY] || 0);
  var min = Number(row[c.MIN_LEVEL] || 0);
  var consumed = Number(row[c.CONSUMED_PER_DAY] || 0);
  var daysLeft = consumed > 0 ? Math.floor(qty / consumed) : null;
  return {
    itemId: String(row[c.ITEM_ID] || '').trim(),
    name: String(row[c.ITEM_NAME] || '').trim(),
    location: String(row[c.LOCATION] || '').trim() || 'Hyderabad',
    quantity: qty,
    unit: String(row[c.UNIT] || '').trim() || 'bags',
    minLevel: min,
    consumedPerDay: consumed,
    daysLeft: daysLeft,
    low: qty <= min,
    notes: String(row[c.NOTES] || '').trim(),
    pendingDefaults: [],
    pendingDays: 0,
    pendingQtyTotal: 0,
    needsConfirm: false,
    skippedOlderDays: 0,
    projectedQuantity: qty,
    projectedDaysLeft: daysLeft,
    projectedLow: qty <= min,
    shortfall: false,
    requiresCount: false,
    lastCountDate: '',
    gapSince: '',
    lastConsumeDate: '',
    lastConsumeQty: 0,
    avgPerDay: 0,
    loggedDays: 0,
    burnRate: consumed,
    packSizeKg: Number(row[c.PACK_SIZE_KG] || 0) || 0,
    usageMode: String(row[c.USAGE_MODE] || 'Regular').trim() === 'Occasional' ? 'Occasional' : 'Regular'
  };
}

function _readFeedSummary_(sheet, movementsSheet) {
  var items = [];
  var lowCount = 0;
  var totalQty = 0;
  var pendingConfirmCount = 0;
  var pendingDayTotal = 0;
  var staleCount = 0;
  var locations = {};
  if (!sheet || sheet.getLastRow() < 2) {
    return {
      items: items, itemCount: 0, lowCount: 0, totalQty: 0,
      pendingConfirmCount: 0, pendingDayTotal: 0, staleCount: 0, locations: []
    };
  }
  var consumeIndex = _buildFeedConsumeIndex_(
    movementsSheet || (_ensureStockSheets_().feedMovements)
  );
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var item = _enrichFeedItemWithDefaults_(_feedFromRow_(data[i]), consumeIndex);
    if (!item.itemId && !item.name) continue;
    if (item.low || item.projectedLow) lowCount++;
    if (item.needsConfirm) {
      pendingConfirmCount++;
      pendingDayTotal += Number(item.pendingDays || 0);
    }
    if (item.requiresCount) staleCount++;
    if (item.location) locations[item.location] = true;
    totalQty += item.quantity;
    items.push(item);
  }
  items.sort(function (a, b) {
    if (a.requiresCount !== b.requiresCount) return a.requiresCount ? -1 : 1;
    if (a.needsConfirm !== b.needsConfirm) return a.needsConfirm ? -1 : 1;
    if (a.projectedLow !== b.projectedLow) return a.projectedLow ? -1 : 1;
    if (a.low !== b.low) return a.low ? -1 : 1;
    return a.name.localeCompare(b.name);
  });
  return {
    items: items,
    itemCount: items.length,
    lowCount: lowCount,
    totalQty: Math.round(totalQty * 100) / 100,
    pendingConfirmCount: pendingConfirmCount,
    pendingDayTotal: pendingDayTotal,
    staleCount: staleCount,
    locations: Object.keys(locations).sort()
  };
}

function getFeedList(username, token) {
  try {
    _requireStockTrainer_(username, token);
    var sheets = _ensureStockSheets_();
    _maybeAutoArchiveStock_(sheets);
    var summary = _readFeedSummary_(sheets.feed, sheets.feedMovements);
    return {
      success: true,
      items: summary.items,
      stats: summary,
      locations: summary.locations,
      asOf: _stockYmd_(new Date()),
      throughDate: _feedYesterdayYmd_(),
      lookbackDays: FEED_DEFAULT_LOOKBACK_DAYS
    };
  } catch (e) {
    return { success: false, message: String(e.message || e), items: [] };
  }
}

/**
 * Confirm (or edit) missed EOD consume defaults for one feed item.
 * payload.entries: [{ date, qty }] — qty may be edited from the suggested default.
 */
function confirmFeedConsumeDefaults(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { success: false, message: 'Please try again.' };
  try {
    var trainer = _requireStockTrainer_(payload.username, payload.token);
    var itemId = String(payload.itemId || '').trim();
    if (!itemId) throw new Error('Feed item is required.');
    var entries = payload.entries || [];
    if (!entries.length) throw new Error('Nothing to confirm.');

    var sheets = _ensureStockSheets_();
    var c = CONFIG.FEED_COLS;
    var rowIndex = _findStockRow_(sheets.feed, c.ITEM_ID, itemId);
    if (!rowIndex) throw new Error('Feed item not found.');

    var consumeIndex = _buildFeedConsumeIndex_(sheets.feedMovements);
    var info = consumeIndex[itemId] || { dates: {}, lastDate: '', lastQty: 0, lastCountDate: '' };
    var current = Number(sheets.feed.getRange(rowIndex, c.QUANTITY + 1).getValue() || 0);
    var now = new Date();

    var gapCheck = _pendingFeedDefaultsForItem_(_feedFromRow_(
      sheets.feed.getRange(rowIndex, 1, 1, FEED_HEADERS.length).getValues()[0]
    ), info, _feedYesterdayYmd_());
    if (gapCheck.requiresCount) {
      throw new Error('More than ' + FEED_DEFAULT_LOOKBACK_DAYS + ' days have no entry, so an estimate '
        + 'would not be reliable. Weigh or count what is actually left and record it as a stock count.');
    }

    // Validate the whole batch BEFORE writing anything, so the ledger and the
    // stock quantity can never drift apart on a mid-batch failure.
    var plan = _planFeedConsumeEntries_(entries, info, current, {
      allowShortfall: payload.allowShortfall === true
    });
    if (plan.error) throw new Error(plan.error);
    if (!plan.rows.length) {
      return { success: true, message: 'Those days were already logged.', quantity: current };
    }

    // Single block write keeps IDs unique and avoids partial appends
    var seed = _feedMovementSeed_(sheets.feedMovements);
    var appended = plan.rows.map(function (r, idx) {
      return [
        'FM-' + String(seed + idx).padStart(3, '0'),
        itemId,
        'Consume',
        _stockDate_(r.date),
        r.qty,
        r.note,
        now,
        trainer.name || trainer.username
      ];
    });
    _appendStockRows_(sheets.feedMovements, appended, FEED_MOVEMENT_HEADERS.length);

    sheets.feed.getRange(rowIndex, c.QUANTITY + 1).setValue(plan.remaining);
    sheets.feed.getRange(rowIndex, c.UPDATED_AT + 1).setValue(now);
    sheets.feed.getRange(rowIndex, c.UPDATED_BY + 1).setValue(trainer.name || trainer.username);

    return {
      success: true,
      message: plan.rows.length + ' day(s) confirmed. Qty now ' + plan.remaining + '.'
        + (plan.clamped ? ' Stock hit zero — check for a missed restock.' : ''),
      quantity: plan.remaining,
      recorded: plan.rows.length,
      clamped: plan.clamped
    };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/** Appends rows in one write, growing the grid first if the sheet was trimmed. */
function _appendStockRows_(sheet, rows, width) {
  if (!rows || !rows.length) return;
  var startRow = sheet.getLastRow() + 1;
  var needed = startRow + rows.length - 1;
  if (sheet.getMaxRows() < needed) sheet.insertRowsAfter(sheet.getMaxRows(), needed - sheet.getMaxRows());
  sheet.getRange(startRow, 1, rows.length, width).setValues(rows);
}

/** Next numeric suffix for FM- movement ids. */
function _feedMovementSeed_(movementsSheet) {
  var max = 0;
  var mc = CONFIG.FEED_MOVEMENT_COLS;
  if (movementsSheet.getLastRow() > 1) {
    movementsSheet.getRange(2, mc.MOVEMENT_ID + 1, movementsSheet.getLastRow() - 1, 1)
      .getValues().forEach(function (row) {
        var digits = String(row[0] || '').replace(/\D/g, '');
        max = Math.max(max, Number(digits) || 0);
      });
  }
  return max + 1;
}

/**
 * Validate + cost out a batch of confirmed consume days without touching sheets.
 * Returns { rows, remaining, clamped, error }.
 */
function _planFeedConsumeEntries_(entries, info, current, opts) {
  opts = opts || {};
  var out = { rows: [], remaining: Number(current || 0), clamped: false, error: '' };
  var seen = {};
  var sorted = (entries || []).slice().sort(function (a, b) {
    return String(a.date || '').localeCompare(String(b.date || ''));
  });
  var running = Number(current || 0);
  var total = 0;
  for (var i = 0; i < sorted.length; i++) {
    var date = _stockYmd_(sorted[i].date);
    var qty = Number(sorted[i].qty != null ? sorted[i].qty : sorted[i].quantity);
    if (!date) { out.error = 'Each entry needs a date.'; return out; }
    if (!(qty > 0)) { out.error = 'Qty must be more than zero for ' + date + '.'; return out; }
    if (date > _stockYmd_(new Date())) { out.error = 'Cannot log a future date (' + date + ').'; return out; }
    if (seen[date]) { out.error = 'Duplicate rows for ' + date + '.'; return out; }
    seen[date] = true;
    if (info && info.dates && info.dates[date]) continue; // already logged
    total += qty;
    var wasEdited = sorted[i].edited === true || sorted[i].wasEdited === true;
    out.rows.push({
      date: date,
      qty: qty,
      note: wasEdited
        ? 'Confirmed after edit (was default estimate)'
        : 'Confirmed default (copied from prior-day consume)'
    });
  }
  if (total > running) {
    if (!opts.allowShortfall) {
      out.error = 'These days need ' + Math.round(total * 100) / 100 + ' but only '
        + running + ' is on hand. Add the missing restock first, or lower the amounts.';
      return out;
    }
    out.clamped = true;
    out.remaining = 0;
  } else {
    out.remaining = Math.round((running - total) * 100) / 100;
  }
  return out;
}

/**
 * Confirm every pending default across all feed items using suggested qty.
 */
function confirmAllFeedConsumeDefaults(username, token) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, message: 'Please try again.' };
  try {
    var trainer = _requireStockTrainer_(username, token);
    var sheets = _ensureStockSheets_();
    var summary = _readFeedSummary_(sheets.feed, sheets.feedMovements);
    var total = 0;
    var skipped = [];
    var needCount = [];
    summary.items.forEach(function (item) {
      if (item.requiresCount) { needCount.push(item.name || item.itemId); return; }
      if (!item.needsConfirm || !(item.pendingDefaults || []).length) return;
      var res = _confirmFeedDefaultsUnlocked_(item.itemId, item.pendingDefaults.map(function (p) {
        return { date: p.date, qty: p.suggestedQty, edited: false };
      }), trainer, sheets);
      if (res && res.success) total += Number(res.recorded || 0);
      else skipped.push(item.name || item.itemId);
    });
    var notes = '';
    if (skipped.length) notes += ' Skipped (needs a restock entry first): ' + skipped.join(', ') + '.';
    if (needCount.length) notes += ' Needs a physical stock count: ' + needCount.join(', ') + '.';
    if (!total && (skipped.length || needCount.length)) {
      return { success: false, message: 'Nothing could be confirmed.' + notes };
    }
    return {
      success: true,
      message: total
        ? ('Confirmed ' + total + ' estimated day(s).' + notes)
        : 'Nothing was waiting to be confirmed.',
      recorded: total,
      skipped: skipped,
      needCount: needCount
    };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/** Internal confirm used by confirmAll (caller already holds the script lock). */
function _confirmFeedDefaultsUnlocked_(itemId, entries, trainer, sheets) {
  try {
    var c = CONFIG.FEED_COLS;
    var rowIndex = _findStockRow_(sheets.feed, c.ITEM_ID, itemId);
    if (!rowIndex) return { success: false, message: 'Feed item not found.' };
    var consumeIndex = _buildFeedConsumeIndex_(sheets.feedMovements);
    var info = consumeIndex[itemId] || { dates: {}, lastDate: '', lastQty: 0 };
    var current = Number(sheets.feed.getRange(rowIndex, c.QUANTITY + 1).getValue() || 0);
    var plan = _planFeedConsumeEntries_(entries, info, current, { allowShortfall: false });
    if (plan.error) return { success: false, message: plan.error };
    if (!plan.rows.length) return { success: true, recorded: 0, quantity: current };

    var now = new Date();
    var seed = _feedMovementSeed_(sheets.feedMovements);
    var rows = plan.rows.map(function (r, idx) {
      return [
        'FM-' + String(seed + idx).padStart(3, '0'),
        itemId, 'Consume', _stockDate_(r.date), r.qty, r.note,
        now, trainer.name || trainer.username
      ];
    });
    _appendStockRows_(sheets.feedMovements, rows, FEED_MOVEMENT_HEADERS.length);
    sheets.feed.getRange(rowIndex, c.QUANTITY + 1).setValue(plan.remaining);
    sheets.feed.getRange(rowIndex, c.UPDATED_AT + 1).setValue(now);
    sheets.feed.getRange(rowIndex, c.UPDATED_BY + 1).setValue(trainer.name || trainer.username);
    return { success: true, recorded: plan.rows.length, quantity: plan.remaining };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  }
}

function addFeedItem(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Please try again.' };
  try {
    var trainer = _requireStockTrainer_(payload.username, payload.token);
    var name = String(payload.name || payload.itemName || '').trim();
    if (!name) throw new Error('Item name is required.');
    var sheets = _ensureStockSheets_();
    var c = CONFIG.FEED_COLS;
    var itemId = String(payload.itemId || '').trim().toUpperCase();
    if (!itemId) itemId = _nextStockId_(sheets.feed, c.ITEM_ID, 'FEED-');
    if (_findStockRow_(sheets.feed, c.ITEM_ID, itemId)) throw new Error('Item ID already exists.');
    var now = new Date();
    var qty = Number(payload.quantity != null ? payload.quantity : payload.qty) || 0;
    var row = new Array(FEED_HEADERS.length).fill('');
    row[c.ITEM_ID] = itemId;
    row[c.ITEM_NAME] = name;
    row[c.LOCATION] = String(payload.location || 'Hyderabad').trim() || 'Hyderabad';
    row[c.QUANTITY] = qty;
    row[c.UNIT] = String(payload.unit || 'bags').trim() || 'bags';
    row[c.MIN_LEVEL] = Number(payload.minLevel || 0) || 0;
    row[c.CONSUMED_PER_DAY] = Number(payload.consumedPerDay || 0) || 0;
    row[c.NOTES] = String(payload.notes || '').trim();
    row[c.UPDATED_AT] = now;
    row[c.UPDATED_BY] = trainer.name || trainer.username;
    row[c.PACK_SIZE_KG] = Math.max(0, Number(payload.packSizeKg || 0) || 0);
    var usageMode = String(payload.usageMode || 'Regular').trim() === 'Occasional' ? 'Occasional' : 'Regular';
    row[c.USAGE_MODE] = usageMode;
    if (usageMode === 'Occasional') row[c.CONSUMED_PER_DAY] = 0;
    sheets.feed.appendRow(row);
    if (qty > 0) {
      sheets.feedMovements.appendRow([
        _nextStockId_(sheets.feedMovements, CONFIG.FEED_MOVEMENT_COLS.MOVEMENT_ID, 'FM-'),
        itemId, 'Restock', _stockDate_(_stockYmd_(payload.inventoryDate) || _stockYmd_(now)),
        qty, 'Opening stock', now, trainer.name || trainer.username, '', ''
      ]);
    }
    return { success: true, message: name + ' added.', itemId: itemId };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function updateFeedItem(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Please try again.' };
  try {
    var trainer = _requireStockTrainer_(payload.username, payload.token);
    var sheets = _ensureStockSheets_();
    var c = CONFIG.FEED_COLS;
    var rowIndex = _findStockRow_(sheets.feed, c.ITEM_ID, payload.itemId);
    if (!rowIndex) throw new Error('Feed item not found.');
    var name = String(payload.name || payload.itemName || '').trim();
    if (!name) throw new Error('Item name is required.');
    sheets.feed.getRange(rowIndex, c.ITEM_NAME + 1).setValue(name);
    sheets.feed.getRange(rowIndex, c.LOCATION + 1).setValue(String(payload.location || 'Hyderabad').trim() || 'Hyderabad');
    // The unit is deliberately not editable: past movements are stored in it, so
    // changing it would silently reinterpret the whole history.
    sheets.feed.getRange(rowIndex, c.MIN_LEVEL + 1).setValue(Number(payload.minLevel || 0) || 0);
    sheets.feed.getRange(rowIndex, c.CONSUMED_PER_DAY + 1).setValue(Number(payload.consumedPerDay || 0) || 0);
    sheets.feed.getRange(rowIndex, c.NOTES + 1).setValue(String(payload.notes || '').trim());
    sheets.feed.getRange(rowIndex, c.PACK_SIZE_KG + 1).setValue(Math.max(0, Number(payload.packSizeKg || 0) || 0));
    var usageMode = String(payload.usageMode || 'Regular').trim() === 'Occasional' ? 'Occasional' : 'Regular';
    sheets.feed.getRange(rowIndex, c.USAGE_MODE + 1).setValue(usageMode);
    if (usageMode === 'Occasional') {
      sheets.feed.getRange(rowIndex, c.CONSUMED_PER_DAY + 1).setValue(0);
    }
    sheets.feed.getRange(rowIndex, c.UPDATED_AT + 1).setValue(new Date());
    sheets.feed.getRange(rowIndex, c.UPDATED_BY + 1).setValue(trainer.name || trainer.username);
    return { success: true, message: name + ' updated.' };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function recordFeedMovement(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Please try again.' };
  try {
    var trainer = _requireStockTrainer_(payload.username, payload.token);
    var type = String(payload.type || '').trim();
    if (FEED_MOVEMENT_TYPES.indexOf(type) < 0) {
      throw new Error('Choose Restock, Consume, Adjust, Use or Expire.');
    }
    var qty = Number(payload.qty != null ? payload.qty : payload.quantity);
    if (!(qty > 0) && !(type === 'Adjust' && qty === 0)) {
      throw new Error('Quantity must be greater than zero.');
    }
    var sheets = _ensureStockSheets_();
    var c = CONFIG.FEED_COLS;
    var rowIndex = _findStockRow_(sheets.feed, c.ITEM_ID, payload.itemId);
    if (!rowIndex) throw new Error('Feed item not found.');

    // Optional kg entry: convert to the item's stock unit using its pack size
    var stockUnit = String(sheets.feed.getRange(rowIndex, c.UNIT + 1).getValue() || 'bags').trim() || 'bags';
    var packSize = Number(sheets.feed.getRange(rowIndex, c.PACK_SIZE_KG + 1).getValue() || 0) || 0;
    var enteredNote = '';
    if (String(payload.enteredUnit || '').trim().toLowerCase() === 'kg'
      && stockUnit.toLowerCase() !== 'kg') {
      if (!(packSize > 0)) {
        throw new Error('Set the weight of one ' + stockUnit.replace(/s$/, '')
          + ' on this feed item before entering kilograms.');
      }
      enteredNote = 'entered as ' + qty + ' kg @ ' + packSize + ' kg/' + stockUnit.replace(/s$/, '');
      qty = Math.round((qty / packSize) * 100) / 100;
      if (!(qty > 0) && type !== 'Adjust') throw new Error('That works out to zero ' + stockUnit + '.');
    }

    var moveYmd = _stockYmd_(payload.date) || _stockYmd_(new Date());
    if (moveYmd > _stockYmd_(new Date())) throw new Error('Cannot record a future date.');
    if (type === 'Consume' && payload.allowDuplicate !== true) {
      var existing = _buildFeedConsumeIndex_(sheets.feedMovements)[String(payload.itemId || '').trim()];
      if (existing && existing.dates && existing.dates[moveYmd]) {
        throw new Error('Consume for ' + moveYmd + ' is already logged ('
          + existing.dates[moveYmd] + '). Use Adjust to correct it instead.');
      }
    }
    var current = Number(sheets.feed.getRange(rowIndex, c.QUANTITY + 1).getValue() || 0);
    var next = current;
    if (type === 'Restock') next = current + qty;
    else if (type === 'Consume' || type === 'Use' || type === 'Expire') {
      if (qty > current) throw new Error('Not enough stock (' + current + ' available).');
      next = current - qty;
    } else if (type === 'Adjust') next = qty;
    var now = new Date();
    sheets.feed.getRange(rowIndex, c.QUANTITY + 1).setValue(next);
    sheets.feed.getRange(rowIndex, c.UPDATED_AT + 1).setValue(now);
    sheets.feed.getRange(rowIndex, c.UPDATED_BY + 1).setValue(trainer.name || trainer.username);
    var note = String(payload.notes || '').trim();
    var horseId = String(payload.horseId || payload.refHorseId || '').trim();
    var horseName = String(payload.horseName || payload.refHorseName || '').trim();
    if (type === 'Use' && horseName) {
      note = (note ? note + ' · ' : '') + 'for ' + horseName + (horseId ? ' (' + horseId + ')' : '');
    }
    if (type === 'Expire' && !note) note = 'Expired / discarded';
    if (type === 'Adjust') {
      var delta = Math.round((next - current) * 100) / 100;
      note = ('Stock count corrected ' + current + ' → ' + next
        + ' (' + (delta >= 0 ? '+' : '') + delta + ')' + (note ? ' · ' + note : ''));
    }
    if (enteredNote) note = note ? (note + ' · ' + enteredNote) : enteredNote;
    sheets.feedMovements.appendRow([
      _nextStockId_(sheets.feedMovements, CONFIG.FEED_MOVEMENT_COLS.MOVEMENT_ID, 'FM-'),
      String(payload.itemId || '').trim(),
      type,
      _stockDate_(moveYmd),
      type === 'Adjust' ? qty : qty,
      note,
      now,
      trainer.name || trainer.username,
      horseId,
      horseName
    ]);
    return {
      success: true,
      message: type + ' recorded' + (enteredNote ? ' as ' + qty + ' ' + stockUnit : '')
        + '. Stock now ' + next + ' ' + stockUnit + '.',
      quantity: next
    };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function getFeedMovements(username, token, itemId, limit) {
  try {
    _requireStockTrainer_(username, token);
    var sheets = _ensureStockSheets_();
    var target = String(itemId || '').trim();
    var cap = Math.max(1, Number(limit) || 10);
    var movements = [];
    if (sheets.feedMovements.getLastRow() > 1) {
      var mc = CONFIG.FEED_MOVEMENT_COLS;
      var data = sheets.feedMovements.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var id = String(data[i][mc.ITEM_ID] || '').trim();
        if (target && id !== target) continue;
        movements.push({
          movementId: String(data[i][mc.MOVEMENT_ID] || '').trim(),
          itemId: id,
          type: String(data[i][mc.TYPE] || '').trim(),
          date: _stockYmd_(data[i][mc.DATE]),
          qty: Number(data[i][mc.QTY] || 0),
          notes: String(data[i][mc.NOTES] || '').trim(),
          recordedBy: String(data[i][mc.RECORDED_BY] || '').trim(),
          horseId: mc.REF_HORSE_ID != null ? String(data[i][mc.REF_HORSE_ID] || '').trim() : '',
          horseName: mc.REF_HORSE_NAME != null ? String(data[i][mc.REF_HORSE_NAME] || '').trim() : ''
        });
      }
    }
    movements.reverse();
    return { success: true, movements: movements.slice(0, cap) };
  } catch (e) {
    return { success: false, message: String(e.message || e), movements: [] };
  }
}

// ── Tack helpers ─────────────────────────────────────────────
function _tackFromRow_(row) {
  var c = CONFIG.TACK_COLS;
  var qty = Number(row[c.QUANTITY] || 0);
  var min = Number(row[c.MIN_LEVEL] || 0);
  return {
    itemId: String(row[c.ITEM_ID] || '').trim(),
    name: String(row[c.ITEM_NAME] || '').trim(),
    category: String(row[c.CATEGORY] || '').trim(),
    model: String(row[c.MODEL] || '').trim(),
    vendor: String(row[c.VENDOR] || '').trim(),
    location: String(row[c.LOCATION] || '').trim() || 'Hyderabad',
    quantity: qty,
    minLevel: min,
    low: qty <= min,
    photoUrl: _stockPhotoViewUrl_(row[c.PHOTO_FILE_ID], row[c.PHOTO_URL]),
    photoFileId: String(row[c.PHOTO_FILE_ID] || '').trim(),
    notes: String(row[c.NOTES] || '').trim()
  };
}

function _readTackSummary_(sheet) {
  var items = [];
  var lowCount = 0;
  var totalQty = 0;
  var outOfStock = 0;
  var locations = {};
  var categories = {};
  if (!sheet || sheet.getLastRow() < 2) {
    return {
      items: items, itemCount: 0, lowCount: 0, totalQty: 0,
      outOfStock: 0, locations: [], categories: []
    };
  }
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var item = _tackFromRow_(data[i]);
    if (!item.itemId && !item.name) continue;
    if (item.low) lowCount++;
    if (item.quantity <= 0) outOfStock++;
    if (item.location) locations[item.location] = true;
    if (item.category) categories[item.category] = true;
    totalQty += item.quantity;
    items.push(item);
  }
  items.sort(function (a, b) {
    if (a.low !== b.low) return a.low ? -1 : 1;
    return a.name.localeCompare(b.name);
  });
  return {
    items: items,
    itemCount: items.length,
    lowCount: lowCount,
    totalQty: totalQty,
    outOfStock: outOfStock,
    locations: Object.keys(locations).sort(),
    categories: Object.keys(categories).sort()
  };
}

function getTackList(username, token) {
  try {
    _requireStockTrainer_(username, token);
    var sheets = _ensureStockSheets_();
    var summary = _readTackSummary_(sheets.tack);
    return {
      success: true,
      items: summary.items,
      stats: summary,
      locations: summary.locations,
      categories: summary.categories
    };
  } catch (e) {
    return { success: false, message: String(e.message || e), items: [] };
  }
}

function addTackItem(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Please try again.' };
  try {
    var trainer = _requireStockTrainer_(payload.username, payload.token);
    var name = String(payload.name || payload.itemName || '').trim();
    if (!name) throw new Error('Item name is required.');
    var sheets = _ensureStockSheets_();
    var c = CONFIG.TACK_COLS;
    var itemId = String(payload.itemId || '').trim().toUpperCase();
    if (!itemId) itemId = _nextStockId_(sheets.tack, c.ITEM_ID, 'TACK-');
    if (_findStockRow_(sheets.tack, c.ITEM_ID, itemId)) throw new Error('Item ID already exists.');
    var now = new Date();
    var qty = Number(payload.quantity != null ? payload.quantity : payload.qty) || 0;
    var location = String(payload.location || 'Hyderabad').trim() || 'Hyderabad';
    var row = new Array(TACK_HEADERS.length).fill('');
    row[c.ITEM_ID] = itemId;
    row[c.ITEM_NAME] = name;
    row[c.CATEGORY] = String(payload.category || '').trim();
    row[c.MODEL] = String(payload.model || '').trim();
    row[c.VENDOR] = String(payload.vendor || '').trim();
    row[c.LOCATION] = location;
    row[c.QUANTITY] = qty;
    row[c.MIN_LEVEL] = Number(payload.minLevel || 0) || 0;
    row[c.PHOTO_URL] = String(payload.photoUrl || '').trim();
    row[c.PHOTO_FILE_ID] = String(payload.photoFileId || '').trim();
    row[c.NOTES] = String(payload.notes || '').trim();
    row[c.UPDATED_AT] = now;
    row[c.UPDATED_BY] = trainer.name || trainer.username;
    sheets.tack.appendRow(row);
    if (qty > 0) {
      var mc = CONFIG.TACK_MOVEMENT_COLS;
      var mRow = new Array(TACK_MOVEMENT_HEADERS.length).fill('');
      mRow[mc.MOVEMENT_ID] = _nextStockId_(sheets.tackMovements, mc.MOVEMENT_ID, 'TM-');
      mRow[mc.ITEM_ID] = itemId;
      mRow[mc.TYPE] = 'Restock';
      mRow[mc.DATE] = _stockDate_(_stockYmd_(payload.inventoryDate) || _stockYmd_(now));
      mRow[mc.QTY] = qty;
      mRow[mc.REASON] = 'Opening stock';
      mRow[mc.FROM_LOCATION] = '';
      mRow[mc.TO_LOCATION] = location;
      mRow[mc.RECORDED_AT] = now;
      mRow[mc.RECORDED_BY] = trainer.name || trainer.username;
      sheets.tackMovements.appendRow(mRow);
    }
    return { success: true, message: name + ' added.', itemId: itemId };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function updateTackItem(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Please try again.' };
  try {
    var trainer = _requireStockTrainer_(payload.username, payload.token);
    var sheets = _ensureStockSheets_();
    var c = CONFIG.TACK_COLS;
    var rowIndex = _findStockRow_(sheets.tack, c.ITEM_ID, payload.itemId);
    if (!rowIndex) throw new Error('Tack item not found.');
    var name = String(payload.name || payload.itemName || '').trim();
    if (!name) throw new Error('Item name is required.');
    sheets.tack.getRange(rowIndex, c.ITEM_NAME + 1).setValue(name);
    sheets.tack.getRange(rowIndex, c.CATEGORY + 1).setValue(String(payload.category || '').trim());
    sheets.tack.getRange(rowIndex, c.MODEL + 1).setValue(String(payload.model || '').trim());
    sheets.tack.getRange(rowIndex, c.VENDOR + 1).setValue(String(payload.vendor || '').trim());
    sheets.tack.getRange(rowIndex, c.LOCATION + 1).setValue(String(payload.location || 'Hyderabad').trim() || 'Hyderabad');
    sheets.tack.getRange(rowIndex, c.MIN_LEVEL + 1).setValue(Number(payload.minLevel || 0) || 0);
    sheets.tack.getRange(rowIndex, c.PHOTO_URL + 1).setValue(String(payload.photoUrl || '').trim());
    sheets.tack.getRange(rowIndex, c.PHOTO_FILE_ID + 1).setValue(String(payload.photoFileId || '').trim());
    sheets.tack.getRange(rowIndex, c.NOTES + 1).setValue(String(payload.notes || '').trim());
    sheets.tack.getRange(rowIndex, c.UPDATED_AT + 1).setValue(new Date());
    sheets.tack.getRange(rowIndex, c.UPDATED_BY + 1).setValue(trainer.name || trainer.username);
    return { success: true, message: name + ' updated.' };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function recordTackMovement(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Please try again.' };
  try {
    var trainer = _requireStockTrainer_(payload.username, payload.token);
    var type = String(payload.type || '').trim();
    if (TACK_MOVEMENT_TYPES.indexOf(type) < 0) throw new Error('Choose Wear Out, Restock or Transfer.');
    var qty = Number(payload.qty != null ? payload.qty : payload.quantity);
    if (!(qty > 0)) throw new Error('Quantity must be greater than zero.');
    var sheets = _ensureStockSheets_();
    var c = CONFIG.TACK_COLS;
    var rowIndex = _findStockRow_(sheets.tack, c.ITEM_ID, payload.itemId);
    if (!rowIndex) throw new Error('Tack item not found.');
    var rowVals = sheets.tack.getRange(rowIndex, 1, 1, TACK_HEADERS.length).getValues()[0];
    var current = Number(rowVals[c.QUANTITY] || 0);
    var fromLoc = String(rowVals[c.LOCATION] || '').trim() || 'Hyderabad';
    var toLoc = String(payload.toLocation || '').trim();
    var next = current;
    var reason = String(payload.reason || payload.notes || '').trim();

    if (type === 'Restock') {
      next = current + qty;
    } else if (type === 'Wear Out') {
      if (qty > current) throw new Error('Not enough qty to mark worn out (' + current + ' available).');
      if (!reason) throw new Error('Reason is required for wear-out.');
      next = current - qty;
    } else if (type === 'Transfer') {
      if (!toLoc) throw new Error('Destination location is required for transfer.');
      if (toLoc.toLowerCase() === fromLoc.toLowerCase()) {
        throw new Error('Destination is the same as the current location (' + fromLoc + ').');
      }
      if (qty > current) throw new Error('Not enough qty to transfer (' + current + ' available).');
      next = current - qty;
      // Prefer matching same item name + location
      var destRow = 0;
      if (sheets.tack.getLastRow() > 1) {
        var tData = sheets.tack.getDataRange().getValues();
        for (var i = 1; i < tData.length; i++) {
          if (String(tData[i][c.ITEM_NAME] || '').trim() === String(rowVals[c.ITEM_NAME] || '').trim()
            && String(tData[i][c.LOCATION] || '').trim().toLowerCase() === toLoc.toLowerCase()
            && String(tData[i][c.CATEGORY] || '').trim() === String(rowVals[c.CATEGORY] || '').trim()
            && String(tData[i][c.MODEL] || '').trim() === String(rowVals[c.MODEL] || '').trim()) {
            destRow = i + 1;
            break;
          }
        }
      }
      var nowInner = new Date();
      if (destRow) {
        var destQty = Number(sheets.tack.getRange(destRow, c.QUANTITY + 1).getValue() || 0) + qty;
        sheets.tack.getRange(destRow, c.QUANTITY + 1).setValue(destQty);
        sheets.tack.getRange(destRow, c.UPDATED_AT + 1).setValue(nowInner);
        sheets.tack.getRange(destRow, c.UPDATED_BY + 1).setValue(trainer.name || trainer.username);
      } else {
        var newId = _nextStockId_(sheets.tack, c.ITEM_ID, 'TACK-');
        var newRow = rowVals.slice();
        newRow[c.ITEM_ID] = newId;
        newRow[c.LOCATION] = toLoc;
        newRow[c.QUANTITY] = qty;
        // Share the image by URL only — the destination row must not own the Drive
        // file, otherwise replacing one row's photo would affect the other.
        newRow[c.PHOTO_FILE_ID] = '';
        newRow[c.NOTES] = ('Transferred from ' + fromLoc + ' (' + String(rowVals[c.ITEM_ID] || '').trim() + ')');
        newRow[c.UPDATED_AT] = nowInner;
        newRow[c.UPDATED_BY] = trainer.name || trainer.username;
        sheets.tack.appendRow(newRow);
      }
      reason = reason || ('Transferred to ' + toLoc);
    }

    var now = new Date();
    sheets.tack.getRange(rowIndex, c.QUANTITY + 1).setValue(next);
    sheets.tack.getRange(rowIndex, c.UPDATED_AT + 1).setValue(now);
    sheets.tack.getRange(rowIndex, c.UPDATED_BY + 1).setValue(trainer.name || trainer.username);

    var mc = CONFIG.TACK_MOVEMENT_COLS;
    var mRow = new Array(TACK_MOVEMENT_HEADERS.length).fill('');
    mRow[mc.MOVEMENT_ID] = _nextStockId_(sheets.tackMovements, mc.MOVEMENT_ID, 'TM-');
    mRow[mc.ITEM_ID] = String(payload.itemId || '').trim();
    mRow[mc.TYPE] = type;
    mRow[mc.DATE] = _stockDate_(_stockYmd_(payload.date) || _stockYmd_(now));
    mRow[mc.QTY] = qty;
    mRow[mc.REASON] = reason;
    mRow[mc.PHOTO_URL] = String(payload.photoUrl || '').trim();
    mRow[mc.PHOTO_FILE_ID] = String(payload.photoFileId || '').trim();
    mRow[mc.FROM_LOCATION] = fromLoc;
    mRow[mc.TO_LOCATION] = type === 'Transfer' ? toLoc : (type === 'Restock' ? fromLoc : '');
    mRow[mc.RECORDED_AT] = now;
    mRow[mc.RECORDED_BY] = trainer.name || trainer.username;
    sheets.tackMovements.appendRow(mRow);

    return { success: true, message: type + ' recorded. Qty now ' + next + '.', quantity: next };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function getTackMovements(username, token, itemId, limit) {
  try {
    _requireStockTrainer_(username, token);
    var sheets = _ensureStockSheets_();
    var target = String(itemId || '').trim();
    var cap = Math.max(1, Number(limit) || 10);
    var movements = [];
    if (sheets.tackMovements.getLastRow() > 1) {
      var mc = CONFIG.TACK_MOVEMENT_COLS;
      var data = sheets.tackMovements.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var id = String(data[i][mc.ITEM_ID] || '').trim();
        if (target && id !== target) continue;
        movements.push({
          movementId: String(data[i][mc.MOVEMENT_ID] || '').trim(),
          itemId: id,
          type: String(data[i][mc.TYPE] || '').trim(),
          date: _stockYmd_(data[i][mc.DATE]),
          qty: Number(data[i][mc.QTY] || 0),
          reason: String(data[i][mc.REASON] || '').trim(),
          photoUrl: _stockPhotoViewUrl_(data[i][mc.PHOTO_FILE_ID], data[i][mc.PHOTO_URL]),
          fromLocation: String(data[i][mc.FROM_LOCATION] || '').trim(),
          toLocation: String(data[i][mc.TO_LOCATION] || '').trim(),
          recordedBy: String(data[i][mc.RECORDED_BY] || '').trim()
        });
      }
    }
    movements.reverse();
    return { success: true, movements: movements.slice(0, cap) };
  } catch (e) {
    return { success: false, message: String(e.message || e), movements: [] };
  }
}

/** Used by dashboard — returns feed + tack summaries in the shape UI already expects. */
function getStockSummaries_() {
  var sheets = _ensureStockSheets_();
  return {
    feed: _readFeedSummary_(sheets.feed, sheets.feedMovements),
    tack: _readTackSummary_(sheets.tack)
  };
}
