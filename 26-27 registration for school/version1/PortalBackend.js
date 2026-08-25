// ============================================================
// My Rides portal — server functions called from RidersPortalHTML
// (google.script.run.getRiderData, etc.)
// ============================================================

function getRiderData(identifier) {
  try {
    identifier = String(identifier || '').trim();
    if (!identifier) return { found: false, error: 'Please enter your phone number or KE Number.' };

    if (identifier.toUpperCase().startsWith('KE')) {
      const rider = findRiderByKENo(identifier.toUpperCase());
      if (!rider) return { found: false, error: 'No account found. Contact us at +91-9980895533.' };
      return _buildRiderData(rider);
    }

    const riders = findAllRidersByPhone(identifier);
    if (!riders || !riders.length) {
      return { found: false, error: 'No account found. Contact us at +91-9980895533.' };
    }

    if (riders.length === 1) {
      return _buildRiderData(riders[0]);
    }

    return {
      found        : true,
      multiProfile : true,
      profiles     : riders.map(function(r) {
        return {
          keNo    : String(r.row[CONFIG.RIDER_COLS.KE_NO]   || '').trim(),
          name    : r.row[CONFIG.RIDER_COLS.NAME]            || '',
          services: r.row[CONFIG.RIDER_COLS.SERVICES]          || ''
        };
      })
    };
  } catch (err) {
    Logger.log('getRiderData error: ' + err);
    return { found: false, error: 'Something went wrong. Please try again.' };
  }
}

function _buildRiderData(rider) {
  const r    = rider.row;
  const keNo = String(r[CONFIG.RIDER_COLS.KE_NO] || '').trim();
  const ss   = SpreadsheetApp.getActiveSpreadsheet();

  const payments        = _getPaymentsForRider(ss, keNo);
  const sessions        = getSessionsForRider(keNo);
  const classesAttended = _countClassesAttended(ss, keNo);
  const noShowCount     = sessions.filter(function(s) {
    return String(s.attendance || '').toLowerCase() === 'no-show';
  }).length;

  return {
    found            : true,
    keNo             : keNo,
    name             : r[CONFIG.RIDER_COLS.NAME]         || '',
    phone            : String(r[CONFIG.RIDER_COLS.PHONE] || ''),
    email            : r[CONFIG.RIDER_COLS.EMAIL]        || '',
    services         : r[CONFIG.RIDER_COLS.SERVICES]     || '',
    participants     : 1,
    pan              : _lookupPriorPanForRider_(keNo, String(r[CONFIG.RIDER_COLS.PHONE] || '')),
    registeredOn     : r[CONFIG.RIDER_COLS.REGISTERED]
      ? fmtDate(new Date(r[CONFIG.RIDER_COLS.REGISTERED]))
      : '',
    payments         : payments,
    sessions         : sessions,
    classesAttended  : classesAttended,
    noShowCount      : noShowCount,
    totalParticipants: 1,
    shopToken        : _makeShopRiderToken_(keNo),
    curriculum       : getCurriculumProgressForRider(keNo),
    performanceChart : getSkillRadarData(keNo),
    nextSuggestion   : getNextClassSuggestionForRider(keNo)
  };
}

function _makeShopRiderToken_(keNo) {
  var secret = String(ScriptApp.getScriptId() || 'ke-shop');
  if (typeof _hashPassword_ === 'function') {
    return _hashPassword_(String(keNo || '').trim().toUpperCase() + '|shop|' + secret).substring(0, 40);
  }
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(keNo || '').trim().toUpperCase() + '|shop|' + secret,
    Utilities.Charset.UTF_8
  );
  return Utilities.base64EncodeWebSafe(bytes).substring(0, 40);
}

function _requireShopRider_(keNo, token) {
  keNo = String(keNo || '').trim().toUpperCase();
  if (!keNo || String(token || '') !== _makeShopRiderToken_(keNo)) {
    throw new Error('Your My Rides session has expired. Please sign in again.');
  }
  var rider = findRiderByKENo(keNo);
  if (!rider) throw new Error('Rider account not found.');
  return rider;
}

function getNextClassSuggestionForRider(keNo) {
  try {
    var cur = getCurriculumProgressForRider(keNo);
    if (!cur || !cur.items || !cur.items.length) return null;
    var item = cur.items.find(function(x){ return !x.passed; }) || null;
    if (!item) return null;
    return {
      level      : item.level || '',
      classNumber: item.classNumber || '',
      title      : item.title || '',
      objective  : item.objective || '',
      exercise   : item.exercise || '',
      docLink    : item.docLink || ''
    };
  } catch (e) {
    Logger.log('getNextClassSuggestionForRider error: ' + e);
    return null;
  }
}

function submitRiderReflection(payload) {
  try {
    payload = payload || {};
    var bookingId = String(payload.bookingId || '').trim();
    var keNo      = String(payload.keNo || '').trim();
    var text      = String(payload.reflectionText || '').trim();
    var photoLink = String(payload.photoLink || '').trim();
    var videoLink = String(payload.videoLink || '').trim();
    if (!bookingId) return { success: false, error: 'Booking_ID is required.' };
    if (!keNo) return { success: false, error: 'KE No is required.' };
    if (!_bookingBelongsToKe_(bookingId, keNo)) return { success: false, error: 'Booking_ID does not match rider.' };
    if (!text && !photoLink && !videoLink) return { success: false, error: 'Add reflection text or evidence link.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('REFLECTION') || ss.getSheetByName('Reflection');
    if (!sh) {
      sh = ss.insertSheet('REFLECTION');
      sh.appendRow(['Timestamp', 'Booking_ID', 'Student_ID', 'Photo_Link', 'Video_Link', 'Reflection_Text']);
    }
    sh.appendRow([new Date(), bookingId, keNo, photoLink, videoLink, text]);
    return { success: true, message: 'Reflection submitted.' };
  } catch (e) {
    Logger.log('submitRiderReflection error: ' + e);
    return { success: false, error: String(e) };
  }
}

function _bookingBelongsToKe_(bookingId, keNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var b  = ss.getSheetByName('BOOKINGS');
  if (!b || b.getLastRow() < 2) return false;
  var data = b.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1] || '').trim() !== String(bookingId).trim()) continue;
    var studentId = String(data[i][2] || '').trim();
    var mappedKe = _lookupKENoFromStudentId_(studentId);
    return String(mappedKe || '').trim() === String(keNo || '').trim();
  }
  return false;
}

function getCurriculumProgressForRider(keNo) {
  try {
    if (typeof getCurriculumWithProgress === 'function') {
      return getCurriculumWithProgress(keNo);
    }
    // Fallback for older deployments where Curriculumengine.gs is not yet synced.
    return { items: [], passedCount: 0, totalCount: 0, currentLevel: '', currentClassNumber: '', currentTitle: '' };
  } catch (e) {
    Logger.log('getCurriculumProgressForRider error: ' + e);
    return { items: [], passedCount: 0, totalCount: 0, currentLevel: '', currentClassNumber: '', currentTitle: '' };
  }
}

function _lookupKENoFromBookingId_(bookingId) {
  if (!bookingId) return '';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var b = ss.getSheetByName('BOOKINGS');
  if (!b || b.getLastRow() < 2) return '';
  var data = b.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1] || '').trim() === String(bookingId).trim()) {
      var studentId = String(data[i][2] || '').trim();
      return _lookupKENoFromStudentId_(studentId);
    }
  }
  return '';
}

function _lookupKENoFromStudentId_(studentId) {
  if (!studentId) return '';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var students = ss.getSheetByName('STUDENTS');
  var riders = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
  if (!students || !riders) return '';
  var sData = students.getDataRange().getValues();
  var studentName = '';
  for (var i = 1; i < sData.length; i++) {
    if (String(sData[i][0] || '').trim() === String(studentId).trim()) {
      studentName = String(sData[i][1] || '').trim().toLowerCase();
      break;
    }
  }
  if (!studentName) return '';
  var rData = riders.getDataRange().getValues();
  for (var r = 1; r < rData.length; r++) {
    if (String(rData[r][CONFIG.RIDER_COLS.NAME] || '').trim().toLowerCase() === studentName) {
      return String(rData[r][CONFIG.RIDER_COLS.KE_NO] || '').trim();
    }
  }
  return '';
}

/** Service dropdown for portal booking — reads PRICING / service sheet. */
function getServicesList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(CONFIG.SHEETS.PRICING);
    if (!sh) return [];
    const data = sh.getDataRange().getValues();
    const out = [];
    for (let i = 1; i < data.length; i++) {
      const name = String(data[i][CONFIG.PRICING_COLS.NAME] || '').trim();
      if (!name) continue;
      out.push({
        name : name,
        price: Number(data[i][CONFIG.PRICING_COLS.PRICE]) || 0,
        type : String(data[i][CONFIG.PRICING_COLS.TYPE] || 'Regular').trim()
      });
    }
    return out;
  } catch (e) {
    Logger.log('getServicesList error: ' + e);
    return [];
  }
}

function findAllRidersByPhone(phone) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
  if (!sheet) return [];
  const target = normalisePhone(phone);
  if (!target || target.length < 10) return [];
  const data   = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (normalisePhone(data[i][CONFIG.RIDER_COLS.PHONE]) === target) {
      result.push({ rowIndex: i + 1, row: data[i] });
    }
  }
  return result;
}

// ============================================================
//  IN-PORTAL PAYMENT FORM (replaces Google Form for My Rides)
// ============================================================

/** Last known PAN/Aadhaar from payment responses for this KE / phone. */
function _lookupPriorPanForRider_(keNo, phone) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
    if (!sheet || sheet.getLastRow() < 2) return '';
    var data = sheet.getDataRange().getValues();
    var headers = data[0] || [];
    var keNorm = String(keNo || '').trim().toUpperCase();
    var phNorm = normalisePhone(phone);
    for (var i = data.length - 1; i >= 1; i--) {
      var rowKe = String(_paymentVal_(data[i], 'REG_REF', headers) || '').trim().toUpperCase();
      var rowPh = normalisePhone(_paymentVal_(data[i], 'PHONE', headers));
      var pan = String(_paymentVal_(data[i], 'PAN', headers) || '').trim();
      if (!pan) continue;
      if ((keNorm && rowKe === keNorm) || (phNorm && rowPh === phNorm)) return pan;
    }
  } catch (e) {
    Logger.log('_lookupPriorPanForRider_ ' + e);
  }
  return '';
}

function getPortalPaymentPrefill(keNo, token) {
  try {
    var rider = _requireShopRider_(keNo, token);
    var phone = String(rider.row[CONFIG.RIDER_COLS.PHONE] || '').trim();
    var pan = _lookupPriorPanForRider_(keNo, phone);
    return {
      success: true,
      keNo: String(keNo || '').trim().toUpperCase(),
      phone: phone,
      name: String(rider.row[CONFIG.RIDER_COLS.NAME] || '').trim(),
      pan: pan,
      upiId: CONFIG.UPI_ID || ''
    };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  }
}

function getPortalPaymentQr(keNo, token, amount) {
  try {
    _requireShopRider_(keNo, token);
    var amt = Number(amount || 0);
    if (!(amt > 0)) return { success: false, message: 'Enter a valid amount.' };
    var link = createUPILink(amt, String(keNo || '').trim().toUpperCase());
    return {
      success: true,
      upiId: CONFIG.UPI_ID || '',
      upiLink: link,
      qrUrl: createQRCode(link)
    };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  }
}

function _getPaymentScreenshotsFolder_() {
  var rootName = (CONFIG && CONFIG.DRIVE_ROOT_FOLDER) ? CONFIG.DRIVE_ROOT_FOLDER : 'Kings Equestrian';
  var rootIter = DriveApp.getFoldersByName(rootName);
  var root = rootIter.hasNext() ? rootIter.next() : DriveApp.createFolder(rootName);
  var subName = 'Payment Screenshots';
  var subIter = root.getFoldersByName(subName);
  return subIter.hasNext() ? subIter.next() : root.createFolder(subName);
}

function uploadPortalPaymentScreenshot(payload) {
  payload = payload || {};
  try {
    var keNo = String(payload.keNo || '').trim().toUpperCase();
    _requireShopRider_(keNo, payload.token);
    var raw = String(payload.base64Data || payload.data || '').trim();
    if (!raw) throw new Error('No screenshot selected.');
    var mime = String(payload.mimeType || 'image/jpeg').trim();
    var b64 = raw;
    var match = raw.match(/^data:([^;]+);base64,(.+)$/);
    if (match) {
      mime = match[1];
      b64 = match[2];
    }
    var bytes = Utilities.base64Decode(b64);
    if (!bytes || !bytes.length) throw new Error('Could not read screenshot.');
    if (bytes.length > 8 * 1024 * 1024) throw new Error('Screenshot is too large. Please use a smaller image.');
    var ext = mime.indexOf('png') >= 0 ? 'png' : 'jpg';
    var fileName = 'Pay_' + keNo + '_'
      + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.' + ext;
    var file = _getPaymentScreenshotsFolder_().createFile(Utilities.newBlob(bytes, mime, fileName));
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      Logger.log('uploadPortalPaymentScreenshot sharing: ' + shareErr);
    }
    var fileId = file.getId();
    return {
      success: true,
      fileId: fileId,
      screenshotUrl: 'https://drive.google.com/file/d/' + fileId + '/view',
      message: 'Screenshot uploaded.'
    };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  }
}

/**
 * Append one payment response row using header aliases, then run receipt flow.
 * payload: { keNo, token, amount, payDate, txnRef, pan, mode, paymentFor, screenshotUrl|base64Data }
 */
function submitPortalPayment(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    return { success: false, message: 'Please wait a moment and try again.' };
  }
  try {
    var keNo = String(payload.keNo || '').trim().toUpperCase();
    var rider = _requireShopRider_(keNo, payload.token);
    var phone = String(payload.phone || rider.row[CONFIG.RIDER_COLS.PHONE] || '').trim();
    var amount = Number(payload.amount || 0);
    if (!(amount > 0)) throw new Error('Enter a valid amount.');
    var payDateRaw = String(payload.payDate || '').trim();
    if (!payDateRaw) throw new Error('Select the payment date.');
    var mode = String(payload.mode || '').trim();
    if (!mode) throw new Error('Select the mode of payment.');
    var paymentFor = _normalisePaymentType_(payload.paymentFor);
    var txnRef = String(payload.txnRef || '').trim();
    var pan = String(payload.pan || '').trim();
    if (!pan) pan = _lookupPriorPanForRider_(keNo, phone);

    var screenshotUrl = String(payload.screenshotUrl || '').trim();
    if (!screenshotUrl && (payload.base64Data || payload.data)) {
      var up = uploadPortalPaymentScreenshot({
        keNo: keNo,
        token: payload.token,
        base64Data: payload.base64Data || payload.data,
        mimeType: payload.mimeType
      });
      if (!up || !up.success) throw new Error((up && up.message) || 'Screenshot upload failed.');
      screenshotUrl = up.screenshotUrl;
    }
    if (!screenshotUrl) throw new Error('Please upload a payment screenshot.');

    var payDate = payDateRaw;
    if (/^\d{4}-\d{2}-\d{2}$/.test(payDateRaw)) {
      var parts = payDateRaw.split('-');
      payDate = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 12, 0, 0);
    }

    if (_isDuplicatePayment(phone, amount, payDate, new Date(), paymentFor)) {
      return {
        success: false,
        message: 'A matching payment was already recorded for this amount and date.'
      };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
    if (!sheet) throw new Error('Payment Responses sheet not found.');
    _ensurePortalPaymentSheetHeaders_(sheet);

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var row = new Array(headers.length);
    for (var i = 0; i < row.length; i++) row[i] = '';

    function setCol(key, value) {
      var idx = _paymentColIndex_(headers, key);
      if (idx >= 0) row[idx] = value;
    }

    setCol('TIMESTAMP', new Date());
    setCol('REG_REF', keNo);
    setCol('PHONE', phone);
    setCol('AMOUNT', amount);
    setCol('SCREENSHOT', screenshotUrl);
    setCol('PAY_DATE', payDate);
    setCol('TXN_REF', txnRef);
    setCol('PAN', pan);
    setCol('MODE', mode);
    setCol('PAYMENT_FOR', paymentFor);
    setCol('RECEIPT_SENT', 'Verifying');

    sheet.appendRow(row);
    var rowIndex = sheet.getLastRow();

    sendReceiptForRow(rowIndex);

    var paymentSummary = null;
    if (paymentFor === 'Shopping Kit / Equipment' && typeof syncShoppingPaymentsForRider_ === 'function') {
      try {
        paymentSummary = syncShoppingPaymentsForRider_(keNo);
      } catch (syncErr) {
        Logger.log('submitPortalPayment shopping sync: ' + syncErr);
      }
    }

    return {
      success: true,
      message: 'Payment submitted. Receipt will be emailed shortly.',
      keNo: keNo,
      amount: amount,
      paymentFor: paymentFor,
      paymentSummary: paymentSummary
        ? {
            totalOrdered: paymentSummary.totalOrdered,
            totalPaid: paymentSummary.totalPaid,
            balance: paymentSummary.balance,
            extraPaid: paymentSummary.extraPaid
          }
        : null
    };
  } catch (e) {
    Logger.log('submitPortalPayment ERROR: ' + e);
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/** Ensure Payment For + receipt columns exist on the responses sheet. */
function _ensurePortalPaymentSheetHeaders_(sheet) {
  var needed = (typeof _setupPaymentHeaders_ === 'function')
    ? _setupPaymentHeaders_()
    : [
      'Timestamp', 'Registration No', 'Phone number', 'Amount Paid (₹)', 'ScreenShot',
      'Payment Date', 'Transcation Reference Number', 'Pan / AAdhar Number', 'Mode of Payment',
      'Payment For', 'Receipt Sent', 'Receipt Sent At', 'Receipt No', 'Receipt Link'
    ];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, needed.length).setValues([needed]);
    return;
  }
  if (sheet.getMaxColumns() < needed.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needed.length - sheet.getMaxColumns());
  }
  var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), needed.length)).getValues()[0];
  needed.forEach(function (header) {
    var want = _normalisePaymentHeader_(header);
    var found = false;
    for (var i = 0; i < headers.length; i++) {
      if (_normalisePaymentHeader_(headers[i]) === want) { found = true; break; }
    }
    if (!found) {
      var col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(header);
      headers.push(header);
    }
  });
}
