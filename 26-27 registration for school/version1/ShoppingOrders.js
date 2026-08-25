// ============================================================
// KINGS EQUESTRIAN — SHOP CATALOG & ORDERS
// Product configuration and order state live in Google Sheets.
// ============================================================

var SHOP_PRODUCT_HEADERS = [
  'Product ID', 'Product', 'Category', 'Option / Tier', 'Allowed Sizes',
  'Price', 'Image URL', 'Active', 'Sort Order'
];

var SHOP_ORDER_HEADERS = [
  'Created At', 'Order ID', 'KE No', 'Rider Name', 'Email', 'Phone',
  'Items JSON', 'Total', 'Payment Status', 'Payment Verified By',
  'Payment Verified At', 'Order Status', 'Expected Delivery Date',
  'Delivered At', 'Delivery Notes', 'Parent Confirmation',
  'Parent Confirmed At', 'Parent Note', 'Updated At', 'Updated By',
  'Payment Form URL', 'UPI Reference', 'Status Email At', 'Client Request ID',
  'Payment Transaction Ref'
];

function _shopProductSeedRows_() {
  return [
    ['HELMET', 'Helmet', 'Safety', 'Standard', 'Small|Medium', 0, '', 'Yes', 10],
    ['BREECH-BASIC', 'Breeches', 'Clothing', 'Basic', 'TEXT', 0, '', 'Yes', 20],
    ['BREECH-PREMIUM', 'Breeches', 'Clothing', 'Premium', 'TEXT', 0, '', 'Yes', 21],
    ['SHORT-BOOTS', 'Short Boots', 'Footwear', 'Standard', 'UK_SIZE', 0, '', 'Yes', 30],
    ['LONG-BOOTS-STANDARD', 'Long Boots', 'Footwear', 'Standard', 'UK_SIZE', 0, '', 'Yes', 40],
    ['LONG-BOOTS-PREMIUM', 'Long Boots', 'Footwear', 'Premium', 'UK_SIZE', 0, '', 'Yes', 41],
    ['BODY-BASIC', 'Body Protector', 'Safety', 'Basic', 'Small|Medium|Large', 0, '', 'Yes', 50],
    ['BODY-PREMIUM', 'Body Protector', 'Safety', 'Premium', 'Small|Medium|Large', 0, '', 'Yes', 51],
    ['KIT-PREMIUM', 'Premium Kit', 'Kit', 'Premium', 'TEXT', 20000, '', 'Yes', 1],
    ['KIT-BASIC', 'Basic Kit', 'Kit', 'Basic', 'TEXT', 15000, '', 'Yes', 2]
  ];
}

function _ensureShopSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var products = ss.getSheetByName(CONFIG.SHEETS.SHOP_PRODUCTS);
  if (!products) products = ss.insertSheet(CONFIG.SHEETS.SHOP_PRODUCTS);
  if (products.getLastRow() === 0) {
    products.getRange(1, 1, 1, SHOP_PRODUCT_HEADERS.length).setValues([SHOP_PRODUCT_HEADERS]);
    products.getRange(2, 1, _shopProductSeedRows_().length, SHOP_PRODUCT_HEADERS.length)
      .setValues(_shopProductSeedRows_());
    _styleShopHeader_(products, SHOP_PRODUCT_HEADERS.length);
  } else {
    var productIds = {};
    if (products.getLastRow() > 1) {
      products.getRange(2, 1, products.getLastRow() - 1, 1).getValues()
        .forEach(function (r) { productIds[String(r[0] || '').trim()] = true; });
    }
    _shopProductSeedRows_().forEach(function (seed) {
      if (!productIds[String(seed[0])]) products.appendRow(seed);
    });
  }

  var orders = ss.getSheetByName(CONFIG.SHEETS.SHOP_ORDERS);
  if (!orders) orders = ss.insertSheet(CONFIG.SHEETS.SHOP_ORDERS);
  var needCols = SHOP_ORDER_HEADERS.length;
  if (orders.getMaxColumns() < needCols) {
    orders.insertColumnsAfter(orders.getMaxColumns(), needCols - orders.getMaxColumns());
  }
  if (orders.getLastRow() === 0) {
    orders.getRange(1, 1, 1, needCols).setValues([SHOP_ORDER_HEADERS]);
    _styleShopHeader_(orders, needCols);
  } else {
    var existing = orders.getRange(1, 1, 1, needCols).getValues()[0];
    for (var i = 0; i < needCols; i++) {
      if (!String(existing[i] || '').trim()) orders.getRange(1, i + 1).setValue(SHOP_ORDER_HEADERS[i]);
    }
  }
  return { ss: ss, products: products, orders: orders };
}

function _styleShopHeader_(sheet, cols) {
  sheet.getRange(1, 1, 1, cols)
    .setBackground('#1f4e3d').setFontColor('#fff').setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function _shopIsActive_(v) {
  var s = String(v == null ? 'Yes' : v).trim().toLowerCase();
  return s !== 'no' && s !== 'false' && s !== '0' && s !== 'inactive';
}

/**
 * Convert Google Drive share/view links into a URL browsers can show in <img>.
 *
 * Links like https://drive.google.com/file/d/FILE_ID/view?usp=drivesdk are HTML
 * pages — not image files — so <img src> cannot display them.
 *
 * Strategy:
 *  1) Cached data-URI thumbnail fetched with the script owner's Drive access
 *  2) Public thumbnail / uc URLs (works when file is "Anyone with the link")
 * Never return a /file/d/.../view URL to the client.
 */
function _shopDriveFileId_(raw) {
  var url = String(raw || '').trim();
  if (!url) return '';
  if (/^[a-zA-Z0-9_-]{10,}$/.test(url) && url.indexOf('/') < 0 && url.indexOf('http') !== 0) {
    return url;
  }
  var idMatch = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/)
    || url.match(/[?&]id=([a-zA-Z0-9_-]+)/)
    || url.match(/\/d\/([a-zA-Z0-9_-]+)/)
    || url.match(/\/thumbnail\?[^#]*\bid=([a-zA-Z0-9_-]+)/);
  return idMatch ? idMatch[1] : '';
}

function _shopDetectImageMime_(bytes, headerCt) {
  var ct = String(headerCt || '').split(';')[0].trim().toLowerCase();
  if (ct.indexOf('image/') === 0) return ct;
  if (!bytes || bytes.length < 12) return '';
  var b0 = bytes[0] & 0xFF;
  var b1 = bytes[1] & 0xFF;
  var b2 = bytes[2] & 0xFF;
  var b3 = bytes[3] & 0xFF;
  if (b0 === 0xFF && b1 === 0xD8 && b2 === 0xFF) return 'image/jpeg';
  if (b0 === 0x89 && b1 === 0x50 && b2 === 0x4E && b3 === 0x47) return 'image/png';
  if (b0 === 0x47 && b1 === 0x49 && b2 === 0x46) return 'image/gif';
  // RIFF....WEBP
  if (b0 === 0x52 && b1 === 0x49 && b2 === 0x46 && b3 === 0x46
    && (bytes[8] & 0xFF) === 0x57 && (bytes[9] & 0xFF) === 0x45) return 'image/webp';
  if (ct === 'application/octet-stream' || !ct) return 'image/jpeg';
  return '';
}

function _shopBytesToDataUrl_(bytes, headerCt) {
  if (!bytes || !bytes.length || bytes.length > 85000) return '';
  var mime = _shopDetectImageMime_(bytes, headerCt);
  if (!mime) return '';
  var dataUrl = 'data:' + mime + ';base64,' + Utilities.base64Encode(bytes);
  return dataUrl.length < 95000 ? dataUrl : '';
}

function _shopDriveThumbDataUrl_(fileId) {
  fileId = String(fileId || '').trim();
  if (!fileId) return '';
  try {
    var cache = CacheService.getScriptCache();
    var key = 'shopimg_v3_' + fileId;
    var hit = cache.get(key);
    if (hit) return hit;

    var dataUrl = '';
    var token = ScriptApp.getOAuthToken();

    // 1) Authenticated Drive thumbnail (small — fits CacheService + catalog payload)
    try {
      var thumbResp = UrlFetchApp.fetch(
        'https://drive.google.com/thumbnail?id=' + encodeURIComponent(fileId) + '&sz=w600',
        {
          headers: { Authorization: 'Bearer ' + token },
          muteHttpExceptions: true,
          followRedirects: true
        }
      );
      if (thumbResp.getResponseCode() === 200) {
        var th = thumbResp.getHeaders() || {};
        dataUrl = _shopBytesToDataUrl_(
          thumbResp.getContent(),
          th['Content-Type'] || th['content-type'] || ''
        );
      }
    } catch (ignoreThumb) {}

    // 2) DriveApp blob (script owner access) — only if small enough for data-URI
    if (!dataUrl) {
      try {
        var blob = DriveApp.getFileById(fileId).getBlob();
        dataUrl = _shopBytesToDataUrl_(blob.getBytes(), blob.getContentType() || '');
      } catch (ignoreDrive) {}
    }

    // 3) Drive API full media (skip if huge)
    if (!dataUrl) {
      try {
        var mediaResp = UrlFetchApp.fetch(
          'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId)
            + '?alt=media&supportsAllDrives=true',
          {
            headers: { Authorization: 'Bearer ' + token },
            muteHttpExceptions: true,
            followRedirects: true
          }
        );
        if (mediaResp.getResponseCode() === 200) {
          var mh = mediaResp.getHeaders() || {};
          dataUrl = _shopBytesToDataUrl_(
            mediaResp.getContent(),
            mh['Content-Type'] || mh['content-type'] || ''
          );
        }
      } catch (ignoreMedia) {}
    }

    if (dataUrl) {
      try { cache.put(key, dataUrl, 21600); } catch (ignoreCache) {}
    }
    return dataUrl;
  } catch (e) {
    Logger.log('_shopDriveThumbDataUrl_ ' + fileId + ': ' + e);
    return '';
  }
}

function _shopPublicImageUrl_(fileId) {
  // These work as <img src> when the file is shared "Anyone with the link".
  // /file/d/.../view never works in <img>.
  return 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(fileId) + '&sz=w800';
}

function _shopDisplayImageUrl_(raw) {
  var url = String(raw || '').trim();
  if (!url) return String(CONFIG.LOGO_URL || '');

  // Already a data-URI or non-Drive direct image URL
  if (url.indexOf('data:image/') === 0) return url;

  var fileId = _shopDriveFileId_(url);
  if (!fileId) return url;

  // Best: inline thumbnail so every parent sees the image (no Drive login needed)
  var dataUrl = _shopDriveThumbDataUrl_(fileId);
  if (dataUrl) return dataUrl;

  // Fallback: Drive thumbnail URL (not the /view HTML page)
  return _shopPublicImageUrl_(fileId);
}

/** Unique catalog key — Product ID alone is not enough when tiers share an ID. */
function _shopCatalogKey_(productId, option) {
  return String(productId || '').trim() + '||' + String(option || '').trim().toLowerCase();
}

function _readShopCatalog_(sheet, includeInactive) {
  var out = [];
  if (!sheet || sheet.getLastRow() < 2) return out;
  var c = CONFIG.SHOP_PRODUCT_COLS;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][c.PRODUCT_ID] || '').trim();
    if (!id) continue;
    var active = _shopIsActive_(data[i][c.ACTIVE]);
    if (!includeInactive && !active) continue;
    var rawSizes = String(data[i][c.ALLOWED_SIZES] || '').trim();
    var sizeMode = rawSizes.toUpperCase() === 'TEXT' ? 'text'
      : (rawSizes.toUpperCase() === 'UK_SIZE' ? 'uk-size' : 'select');
    var sizes = sizeMode === 'select'
      ? rawSizes.split('|').map(function (x) { return String(x || '').trim(); }).filter(String)
      : [];
    var price = Number(data[i][c.PRICE] || 0);
    var option = String(data[i][c.OPTION] || '').trim();
    out.push({
      productId: id,
      catalogKey: _shopCatalogKey_(id, option),
      product: String(data[i][c.PRODUCT] || '').trim(),
      category: String(data[i][c.CATEGORY] || '').trim(),
      option: option,
      allowedSizes: rawSizes,
      sizeMode: sizeMode,
      sizes: sizes,
      price: price,
      orderable: active && price > 0,
      imageUrl: _shopDisplayImageUrl_(data[i][c.IMAGE_URL]),
      includedItems: id.indexOf('KIT-') === 0
        ? ['Helmet', 'Body Protector', 'Breeches / Riding Pants', 'Short Shoes']
        : [],
      active: active,
      sortOrder: Number(data[i][c.SORT_ORDER] || 999)
    });
  }
  out.sort(function (a, b) {
    return a.sortOrder - b.sortOrder || a.product.localeCompare(b.product)
      || String(a.option || '').localeCompare(String(b.option || ''));
  });
  return out;
}

function getShopCatalog() {
  try {
    var env = _ensureShopSheets_();
    return {
      success: true,
      products: _readShopCatalog_(env.products, false),
      sizingNote: 'Unknown sizes can be coordinated in the WhatsApp group.'
    };
  } catch (e) {
    Logger.log('getShopCatalog error: ' + e);
    return { success: false, products: [], message: 'Could not load shop products.' };
  }
}

function _validateShopSize_(product, supplied) {
  var size = String(supplied || '').trim();
  if (!size) throw new Error('Select or enter a size for ' + product.product + '.');
  if (size.length > (product.category === 'Kit' ? 80 : 40)) throw new Error('Size is too long for ' + product.product + '.');
  if (product.sizeMode === 'select') {
    var match = product.sizes.filter(function (x) {
      return x.toLowerCase() === size.toLowerCase();
    })[0];
    if (!match) throw new Error('Invalid size for ' + product.product + '.');
    return match;
  }
  if (product.sizeMode === 'uk-size' && !/^[0-9]{1,2}(?:\.5)?$/.test(size)) {
    throw new Error('Enter a valid UK shoe size for ' + product.product + '.');
  }
  return size;
}

function _shopOrderId_(sheet) {
  var tz = Session.getScriptTimeZone();
  var day = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  var prefix = 'ORD-' + day + '-';
  var max = 0;
  if (sheet.getLastRow() > 1) {
    var ids = sheet.getRange(2, CONFIG.SHOP_ORDER_COLS.ORDER_ID + 1, sheet.getLastRow() - 1, 1).getValues();
    ids.forEach(function (r) {
      var id = String(r[0] || '');
      if (id.indexOf(prefix) !== 0) return;
      max = Math.max(max, Number(id.substring(prefix.length)) || 0);
    });
  }
  return prefix + String(max + 1).padStart(4, '0');
}

function _findShopOrderRow_(sheet, orderId) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var vals = sheet.getRange(2, CONFIG.SHOP_ORDER_COLS.ORDER_ID + 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim() === String(orderId || '').trim()) return i + 2;
  }
  return 0;
}

function _findShopRequestRow_(sheet, requestId) {
  if (!requestId || !sheet || sheet.getLastRow() < 2) return 0;
  var vals = sheet.getRange(2, CONFIG.SHOP_ORDER_COLS.CLIENT_REQUEST_ID + 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim() === String(requestId).trim()) return i + 2;
  }
  return 0;
}

function placeShopOrder(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { success: false, message: 'Shop is busy. Please try again.' };
  var orderForEmail = null;
  try {
    var keNo = String(payload.keNo || '').trim().toUpperCase();
    var rider = _requireShopRider_(keNo, payload.token);
    var requested = payload.items || [];
    if (!requested.length) return { success: false, message: 'Your cart is empty.' };
    if (requested.length > 20) return { success: false, message: 'Too many cart items.' };

    var env = _ensureShopSheets_();
    var requestId = String(payload.clientRequestId || '').trim();
    var existingRow = _findShopRequestRow_(env.orders, requestId);
    if (existingRow) {
      var existingOrder = _shopOrderFromRow_(
        env.orders.getRange(existingRow, 1, 1, SHOP_ORDER_HEADERS.length).getValues()[0]
      );
      var existingSummary = syncShoppingPaymentsForRider_(keNo);
      var existingDue = Number(existingSummary.balance || 0);
      var existingUpi = existingDue > 0
        ? createUPILink(existingDue, existingOrder.upiReference || existingOrder.orderId) : '';
      return {
        success: true,
        duplicate: true,
        order: existingOrder,
        paymentFormUrl: existingOrder.paymentFormUrl || CONFIG.PAYMENT_FORM_BASE_URL || '',
        upiId: CONFIG.UPI_ID,
        upiLink: existingUpi,
        qrUrl: existingUpi ? createQRCode(existingUpi) : '',
        paymentDue: existingDue
      };
    }

    var catalog = _readShopCatalog_(env.products, false);
    var byKey = {};
    var byIdOnly = {};
    catalog.forEach(function (p) {
      byKey[p.catalogKey] = p;
      // Keep first match only when Product ID is duplicated across tiers
      if (!byIdOnly[p.productId]) byIdOnly[p.productId] = p;
    });
    var cleanItems = [];
    var total = 0;
    requested.forEach(function (req) {
      var productId = String(req.productId || '').trim();
      var option = String(req.option || '').trim();
      var product = byKey[_shopCatalogKey_(productId, option)]
        || (option ? null : byIdOnly[productId]);
      if (!product || !product.active) {
        throw new Error('A selected product is unavailable'
          + (option ? ' (' + option + ')' : '') + '.');
      }
      if (!product.orderable || product.price <= 0) {
        throw new Error(product.product + (product.option ? ' — ' + product.option : '')
          + ' does not have an order price yet.');
      }
      var qty = Math.floor(Number(req.quantity || 1));
      if (qty < 1 || qty > 10) throw new Error('Quantity must be between 1 and 10.');
      var size = _validateShopSize_(product, req.size);
      var lineTotal = product.price * qty;
      cleanItems.push({
        productId: product.productId,
        product: product.product,
        option: product.option,
        size: size,
        quantity: qty,
        unitPrice: product.price,
        lineTotal: lineTotal,
        imageUrl: product.imageUrl,
        includedItems: product.includedItems || [],
        deliveredAt: '',
        deliveredBy: ''
      });
      total += lineTotal;
    });
    total = Math.round(total * 100) / 100;
    if (total <= 0) throw new Error('Order total must be greater than zero.');

    var r = rider.row;
    var orderId = _shopOrderId_(env.orders);
    var now = new Date();
    var paymentFormUrl = CONFIG.PAYMENT_FORM_BASE_URL || '';
    var upiReference = orderId + '-' + keNo;
    var row = new Array(SHOP_ORDER_HEADERS.length).fill('');
    var c = CONFIG.SHOP_ORDER_COLS;
    row[c.CREATED_AT] = now;
    row[c.ORDER_ID] = orderId;
    row[c.KE_NO] = keNo;
    row[c.RIDER_NAME] = String(r[CONFIG.RIDER_COLS.NAME] || '').trim();
    row[c.EMAIL] = String(r[CONFIG.RIDER_COLS.EMAIL] || '').trim();
    row[c.PHONE] = String(r[CONFIG.RIDER_COLS.PHONE] || '').trim();
    row[c.ITEMS_JSON] = JSON.stringify(cleanItems);
    row[c.TOTAL] = total;
    row[c.PAYMENT_STATUS] = 'Pending';
    row[c.ORDER_STATUS] = 'Order Placed';
    row[c.PARENT_CONFIRM_STATUS] = 'Pending';
    row[c.UPDATED_AT] = now;
    row[c.UPDATED_BY] = 'Parent portal';
    row[c.PAYMENT_FORM_URL] = paymentFormUrl;
    row[c.UPI_REFERENCE] = upiReference;
    row[c.CLIENT_REQUEST_ID] = requestId;
    env.orders.appendRow(row);
    var appended = env.orders.getLastRow();
    env.orders.getRange(appended, c.CREATED_AT + 1).setNumberFormat('dd-MMM-yyyy HH:mm');
    env.orders.getRange(appended, c.TOTAL + 1).setNumberFormat('₹#,##0.00');

    var paymentSummary = syncShoppingPaymentsForRider_(keNo);
    var paymentDue = Number(paymentSummary.balance || 0);
    var latestRow = env.orders.getRange(appended, 1, 1, SHOP_ORDER_HEADERS.length).getValues()[0];
    orderForEmail = _shopOrderFromRow_(latestRow);
    orderForEmail.paymentDue = paymentDue;
    orderForEmail.upiLink = paymentDue > 0 ? createUPILink(paymentDue, upiReference) : '';
    orderForEmail.qrUrl = orderForEmail.upiLink ? createQRCode(orderForEmail.upiLink) : '';
    return {
      success: true,
      message: 'Order placed successfully.',
      order: orderForEmail,
      paymentFormUrl: paymentFormUrl,
      upiId: CONFIG.UPI_ID,
      upiLink: orderForEmail.upiLink,
      qrUrl: orderForEmail.qrUrl,
      paymentDue: paymentDue,
      paymentSummary: {
        totalOrdered: paymentSummary.totalOrdered,
        totalPaid: paymentSummary.totalPaid,
        balance: paymentSummary.balance,
        extraPaid: paymentSummary.extraPaid
      }
    };
  } catch (e) {
    Logger.log('placeShopOrder error: ' + e);
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
    if (orderForEmail && typeof sendShopOrderPlacedEmail_ === 'function') {
      try { sendShopOrderPlacedEmail_(orderForEmail); } catch (mailErr) {
        Logger.log('shop placed email error: ' + mailErr);
      }
    }
  }
}

function _shopOrderFromRow_(row) {
  var c = CONFIG.SHOP_ORDER_COLS;
  var items = [];
  try { items = JSON.parse(String(row[c.ITEMS_JSON] || '[]')); } catch (e) {}
  function dateText(v, withTime) {
    if (!v) return '';
    var d = new Date(v);
    if (isNaN(d.getTime())) return String(v);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), withTime ? 'dd-MMM-yyyy HH:mm' : 'dd-MMM-yyyy');
  }
  return {
    createdAt: dateText(row[c.CREATED_AT], true),
    orderId: String(row[c.ORDER_ID] || ''),
    keNo: String(row[c.KE_NO] || ''),
    riderName: String(row[c.RIDER_NAME] || ''),
    email: String(row[c.EMAIL] || ''),
    phone: String(row[c.PHONE] || ''),
    items: items,
    total: Number(row[c.TOTAL] || 0),
    paymentStatus: String(row[c.PAYMENT_STATUS] || 'Pending'),
    paymentVerifiedBy: String(row[c.PAYMENT_VERIFIED_BY] || ''),
    paymentVerifiedAt: dateText(row[c.PAYMENT_VERIFIED_AT], true),
    orderStatus: String(row[c.ORDER_STATUS] || 'Order Placed'),
    expectedDate: dateText(row[c.EXPECTED_DATE], false),
    expectedDateYMD: row[c.EXPECTED_DATE] instanceof Date
      ? Utilities.formatDate(row[c.EXPECTED_DATE], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(row[c.EXPECTED_DATE] || ''),
    deliveredAt: dateText(row[c.DELIVERED_AT], true),
    deliveryNotes: String(row[c.DELIVERY_NOTES] || ''),
    parentConfirmation: String(row[c.PARENT_CONFIRM_STATUS] || 'Pending'),
    parentConfirmedAt: dateText(row[c.PARENT_CONFIRMED_AT], true),
    parentNote: String(row[c.PARENT_NOTE] || ''),
    updatedAt: dateText(row[c.UPDATED_AT], true),
    updatedBy: String(row[c.UPDATED_BY] || ''),
    paymentFormUrl: String(row[c.PAYMENT_FORM_URL] || CONFIG.PAYMENT_FORM_BASE_URL || ''),
    upiReference: String(row[c.UPI_REFERENCE] || ''),
    paymentTxnRef: String(row[c.PAYMENT_TXN_REF] || '')
  };
}

/**
 * Payment form submissions whose Registration No is an Order ID are routed
 * here. This records evidence for manual staff verification and deliberately
 * does not issue a donation/80G receipt.
 */
function recordShopPaymentSubmission_(orderId, amount, txnRef) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var env = _ensureShopSheets_();
    var rowIndex = _findShopOrderRow_(env.orders, orderId);
    if (!rowIndex) return { found: false };
    var c = CONFIG.SHOP_ORDER_COLS;
    var row = env.orders.getRange(rowIndex, 1, 1, SHOP_ORDER_HEADERS.length).getValues()[0];
    var expected = Number(row[c.TOTAL] || 0);
    var paid = Number(amount || 0);
    var matches = expected > 0 && Math.abs(expected - paid) < 0.01;
    row[c.PAYMENT_STATUS] = matches ? 'Submitted' : 'Submitted - Amount Mismatch';
    row[c.PAYMENT_TXN_REF] = String(txnRef || '').trim();
    row[c.UPDATED_AT] = new Date();
    row[c.UPDATED_BY] = 'Payment form';
    env.orders.getRange(rowIndex, 1, 1, SHOP_ORDER_HEADERS.length).setValues([row]);
    return { found: true, amountMatches: matches, expected: expected, paid: paid };
  } finally {
    lock.releaseLock();
  }
}

function _isShoppingPaymentType_(value) {
  var text = String(value || '').trim().toLowerCase();
  return text.indexOf('shop') >= 0 || text.indexOf('kit') >= 0 || text.indexOf('equipment') >= 0;
}

function _calculateShoppingAllocation_(orderTotals, totalPaid) {
  var credit = Math.max(Number(totalPaid || 0), 0);
  var totalOrdered = 0;
  var allocations = (orderTotals || []).map(function (value) {
    var orderTotal = Math.max(Number(value || 0), 0);
    totalOrdered += orderTotal;
    var allocated = Math.min(credit, orderTotal);
    credit -= allocated;
    return allocated;
  });
  return {
    totalOrdered: totalOrdered,
    totalPaid: Math.max(Number(totalPaid || 0), 0),
    balance: Math.max(totalOrdered - Number(totalPaid || 0), 0),
    extraPaid: Math.max(Number(totalPaid || 0) - totalOrdered, 0),
    allocations: allocations
  };
}

/**
 * Allocate all shopping payments for a rider across all orders oldest-first.
 * This supports partial payment, several orders, and overpayment.
 */
function syncShoppingPaymentsForRider_(keNo) {
  keNo = String(keNo || '').trim().toUpperCase();
  var env = _ensureShopSheets_();
  var ledger = env.ss.getSheetByName(CONFIG.SHEETS.PAYMENTS);
  var totalPaid = 0;
  if (ledger && ledger.getLastRow() > 1) {
    var payData = ledger.getDataRange().getValues();
    for (var p = 1; p < payData.length; p++) {
      if (String(payData[p][CONFIG.LEDGER_COLS.KE_NO] || '').trim().toUpperCase() !== keNo) continue;
      if (!_isShoppingPaymentType_(payData[p][CONFIG.LEDGER_COLS.PAYMENT_TYPE])) continue;
      totalPaid += Number(payData[p][CONFIG.LEDGER_COLS.AMOUNT] || 0);
    }
  }

  var orderRows = [];
  if (env.orders.getLastRow() > 1) {
    var orderData = env.orders.getRange(2, 1, env.orders.getLastRow() - 1, SHOP_ORDER_HEADERS.length).getValues();
    orderData.forEach(function (row, i) {
      if (String(row[CONFIG.SHOP_ORDER_COLS.KE_NO] || '').trim().toUpperCase() === keNo) {
        orderRows.push({ rowIndex: i + 2, row: row });
      }
    });
  }
  orderRows.sort(function (a, b) {
    return new Date(a.row[CONFIG.SHOP_ORDER_COLS.CREATED_AT]).getTime()
      - new Date(b.row[CONFIG.SHOP_ORDER_COLS.CREATED_AT]).getTime();
  });
  var allocation = _calculateShoppingAllocation_(
    orderRows.map(function (entry) { return Number(entry.row[CONFIG.SHOP_ORDER_COLS.TOTAL] || 0); }),
    totalPaid
  );
  orderRows.forEach(function (entry, orderIndex) {
    var row = entry.row;
    var c = CONFIG.SHOP_ORDER_COLS;
    var orderTotal = Number(row[c.TOTAL] || 0);
    var allocated = allocation.allocations[orderIndex] || 0;
    row[c.PAYMENT_STATUS] = allocated >= orderTotal && orderTotal > 0
      ? 'Paid' : (allocated > 0 ? 'Partially Paid' : 'Pending');
    if (row[c.PAYMENT_STATUS] === 'Paid') {
      if (!row[c.PAYMENT_VERIFIED_AT]) row[c.PAYMENT_VERIFIED_AT] = new Date();
      if (!row[c.PAYMENT_VERIFIED_BY]) row[c.PAYMENT_VERIFIED_BY] = 'Payment form';
      if (String(row[c.ORDER_STATUS] || 'Order Placed') === 'Order Placed') row[c.ORDER_STATUS] = 'In Progress';
    }
    env.orders.getRange(entry.rowIndex, 1, 1, SHOP_ORDER_HEADERS.length).setValues([row]);
    entry.allocated = allocated;
  });
  return {
    totalOrdered: allocation.totalOrdered,
    totalPaid: allocation.totalPaid,
    balance: allocation.balance,
    extraPaid: allocation.extraPaid,
    allocations: orderRows
  };
}

function getShopOrdersForRider(keNo, token) {
  try {
    keNo = String(keNo || '').trim().toUpperCase();
    _requireShopRider_(keNo, token);
    var summary = syncShoppingPaymentsForRider_(keNo);
    var env = _ensureShopSheets_();
    var out = [];
    if (env.orders.getLastRow() > 1) {
      var data = env.orders.getRange(2, 1, env.orders.getLastRow() - 1, SHOP_ORDER_HEADERS.length).getValues();
      data.forEach(function (row) {
        if (String(row[CONFIG.SHOP_ORDER_COLS.KE_NO] || '').trim().toUpperCase() === keNo) {
          out.push(_shopOrderFromRow_(row));
        }
      });
    }
    out.reverse();
    var allocationById = {};
    (summary.allocations || []).forEach(function (x) {
      allocationById[String(x.row[CONFIG.SHOP_ORDER_COLS.ORDER_ID] || '')] = Number(x.allocated || 0);
    });
    out.forEach(function (o) {
      o.paidAmount = allocationById[o.orderId] || 0;
      o.balance = Math.max(Number(o.total || 0) - o.paidAmount, 0);
    });
    return {
      success: true,
      orders: out,
      paymentSummary: {
        totalOrdered: summary.totalOrdered,
        totalPaid: summary.totalPaid,
        balance: summary.balance,
        extraPaid: summary.extraPaid
      },
      paymentFormUrl: CONFIG.PAYMENT_FORM_BASE_URL || ''
    };
  } catch (e) {
    return { success: false, orders: [], message: String(e.message || e) };
  }
}

function getShopOrdersForStaff(keNo, username, token) {
  try {
    var trainer = validateTrainerToken(username, token);
    if (!trainer || !trainer.valid) throw new Error('Your trainer session has expired.');
    keNo = String(keNo || '').trim().toUpperCase();
    var rider = findRiderByKENo(keNo);
    if (!rider) throw new Error('Rider not found.');
    var summary = syncShoppingPaymentsForRider_(keNo);
    var env = _ensureShopSheets_();
    var orders = [];
    if (env.orders.getLastRow() > 1) {
      var data = env.orders.getRange(2, 1, env.orders.getLastRow() - 1, SHOP_ORDER_HEADERS.length).getValues();
      data.forEach(function (row) {
        if (String(row[CONFIG.SHOP_ORDER_COLS.KE_NO] || '').trim().toUpperCase() === keNo) {
          orders.push(_shopOrderFromRow_(row));
        }
      });
    }
    orders.reverse();
    return {
      success: true,
      orders: orders,
      paymentSummary: {
        totalOrdered: summary.totalOrdered,
        totalPaid: summary.totalPaid,
        balance: summary.balance,
        extraPaid: summary.extraPaid
      }
    };
  } catch (e) {
    return { success: false, orders: [], message: String(e.message || e) };
  }
}

function markShopOrderItemsDelivered(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Please try again.' };
  var completedOrder = null;
  try {
    var trainer = validateTrainerToken(payload.username, payload.token);
    if (!trainer || !trainer.valid) throw new Error('Your trainer session has expired.');
    var env = _ensureShopSheets_();
    var rowIndex = _findShopOrderRow_(env.orders, payload.orderId);
    if (!rowIndex) throw new Error('Order not found.');
    var c = CONFIG.SHOP_ORDER_COLS;
    var row = env.orders.getRange(rowIndex, 1, 1, SHOP_ORDER_HEADERS.length).getValues()[0];
    syncShoppingPaymentsForRider_(String(row[c.KE_NO] || ''));
    row = env.orders.getRange(rowIndex, 1, 1, SHOP_ORDER_HEADERS.length).getValues()[0];
    if (String(row[c.PAYMENT_STATUS] || '') !== 'Paid') {
      throw new Error('Full shopping payment is required before delivery.');
    }
    var items = [];
    try { items = JSON.parse(String(row[c.ITEMS_JSON] || '[]')); } catch (ignore) {}
    var indexes = payload.itemIndexes || [];
    if (!indexes.length) throw new Error('Select at least one item.');
    var now = new Date();
    var deliveredBy = trainer.name || trainer.username;
    indexes.forEach(function (value) {
      var idx = Number(value);
      if (!Number.isInteger(idx) || idx < 0 || idx >= items.length) throw new Error('Invalid order item.');
      if (!items[idx].deliveredAt) {
        items[idx].deliveredAt = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
        items[idx].deliveredBy = deliveredBy;
      }
    });
    var allDelivered = items.length > 0 && items.every(function (item) { return !!item.deliveredAt; });
    var anyDelivered = items.some(function (item) { return !!item.deliveredAt; });
    row[c.ITEMS_JSON] = JSON.stringify(items);
    row[c.ORDER_STATUS] = allDelivered ? 'Delivered' : (anyDelivered ? 'Partially Delivered' : 'In Progress');
    if (allDelivered) row[c.DELIVERED_AT] = now;
    row[c.UPDATED_AT] = now;
    row[c.UPDATED_BY] = deliveredBy;
    env.orders.getRange(rowIndex, 1, 1, SHOP_ORDER_HEADERS.length).setValues([row]);
    completedOrder = allDelivered ? _shopOrderFromRow_(row) : null;
    return {
      success: true,
      message: allDelivered ? 'All items marked delivered.' : 'Selected items marked delivered.',
      order: _shopOrderFromRow_(row)
    };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
    if (completedOrder && typeof sendShopOrderStatusEmail_ === 'function') {
      try { sendShopOrderStatusEmail_(completedOrder, 'delivered'); } catch (mailErr) {}
    }
  }
}

function confirmShopOrderReceipt(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Please try again.' };
  try {
    var keNo = String(payload.keNo || '').trim().toUpperCase();
    _requireShopRider_(keNo, payload.token);
    var choice = String(payload.confirmation || '').trim();
    if (choice !== 'Received' && choice !== 'Issue Reported') {
      return { success: false, message: 'Choose Received or Report Issue.' };
    }
    var env = _ensureShopSheets_();
    var rowIndex = _findShopOrderRow_(env.orders, payload.orderId);
    if (!rowIndex) return { success: false, message: 'Order not found.' };
    var row = env.orders.getRange(rowIndex, 1, 1, SHOP_ORDER_HEADERS.length).getValues()[0];
    var c = CONFIG.SHOP_ORDER_COLS;
    if (String(row[c.KE_NO] || '').trim().toUpperCase() !== keNo) {
      return { success: false, message: 'Order does not belong to this rider.' };
    }
    if (String(row[c.ORDER_STATUS] || '') !== 'Delivered') {
      return { success: false, message: 'Receipt can be confirmed only after delivery.' };
    }
    var now = new Date();
    env.orders.getRange(rowIndex, c.PARENT_CONFIRM_STATUS + 1).setValue(choice);
    env.orders.getRange(rowIndex, c.PARENT_CONFIRMED_AT + 1).setValue(now);
    env.orders.getRange(rowIndex, c.PARENT_NOTE + 1).setValue(String(payload.note || '').trim());
    env.orders.getRange(rowIndex, c.UPDATED_AT + 1).setValue(now);
    env.orders.getRange(rowIndex, c.UPDATED_BY + 1).setValue('Parent portal');
    return { success: true, message: choice === 'Received' ? 'Delivery confirmed. Thank you!' : 'Issue reported. Our team will contact you.' };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function getShopOrdersForAdmin(username, token) {
  try {
    _requireShopAdmin_(username, token);
    var env = _ensureShopSheets_();
    var out = [];
    if (env.orders.getLastRow() > 1) {
      var data = env.orders.getRange(2, 1, env.orders.getLastRow() - 1, SHOP_ORDER_HEADERS.length).getValues();
      data.forEach(function (row) { out.push(_shopOrderFromRow_(row)); });
    }
    out.reverse();
    return { success: true, orders: out };
  } catch (e) {
    return { success: false, orders: [], message: String(e.message || e) };
  }
}

function updateShopOrderByAdmin(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: 'Please try again.' };
  var emailOrder = null;
  var emailType = '';
  try {
    var admin = _requireShopAdmin_(payload.username, payload.token);
    var env = _ensureShopSheets_();
    var rowIndex = _findShopOrderRow_(env.orders, payload.orderId);
    if (!rowIndex) return { success: false, message: 'Order not found.' };
    var c = CONFIG.SHOP_ORDER_COLS;
    var row = env.orders.getRange(rowIndex, 1, 1, SHOP_ORDER_HEADERS.length).getValues()[0];
    var oldStatus = String(row[c.ORDER_STATUS] || 'Order Placed');
    var oldExpected = row[c.EXPECTED_DATE] ? new Date(row[c.EXPECTED_DATE]).getTime() : 0;
    var now = new Date();

    if (payload.verifyPayment === true && String(row[c.PAYMENT_STATUS] || '') !== 'Verified') {
      row[c.PAYMENT_STATUS] = 'Verified';
      row[c.PAYMENT_VERIFIED_BY] = admin.name || admin.username;
      row[c.PAYMENT_VERIFIED_AT] = now;
      if (oldStatus === 'Order Placed') row[c.ORDER_STATUS] = 'In Progress';
      emailType = 'payment';
    }

    var requestedStatus = String(payload.orderStatus || row[c.ORDER_STATUS] || '').trim();
    var rank = { 'Order Placed': 1, 'In Progress': 2, 'Delivered': 3 };
    if (!rank[requestedStatus]) return { success: false, message: 'Invalid order status.' };
    if (rank[requestedStatus] < rank[String(row[c.ORDER_STATUS] || 'Order Placed')]) {
      return { success: false, message: 'Order status cannot move backwards.' };
    }
    if (rank[requestedStatus] > rank[String(row[c.ORDER_STATUS] || 'Order Placed')] + 1) {
      return { success: false, message: 'Move the order through each status in sequence.' };
    }
    if (requestedStatus !== 'Order Placed' && String(row[c.PAYMENT_STATUS] || '') !== 'Verified') {
      return { success: false, message: 'Verify payment before processing this order.' };
    }
    if (requestedStatus === 'Delivered' && String(row[c.ORDER_STATUS] || '') !== 'Delivered') {
      row[c.DELIVERED_AT] = now;
      row[c.PARENT_CONFIRM_STATUS] = 'Pending';
      emailType = 'delivered';
    } else if (requestedStatus !== oldStatus && !emailType) {
      emailType = 'status';
    }
    row[c.ORDER_STATUS] = requestedStatus;

    var expectedYMD = String(payload.expectedDate || '').trim();
    if (expectedYMD) {
      if (String(row[c.PAYMENT_STATUS] || '') !== 'Verified') {
        return { success: false, message: 'Verify payment before setting expected delivery.' };
      }
      var parts = expectedYMD.split('-');
      if (parts.length !== 3) return { success: false, message: 'Invalid expected delivery date.' };
      row[c.EXPECTED_DATE] = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 12, 0, 0);
      if (!emailType && row[c.EXPECTED_DATE].getTime() !== oldExpected) emailType = 'expected';
    }
    row[c.DELIVERY_NOTES] = String(payload.deliveryNotes || row[c.DELIVERY_NOTES] || '').trim();
    row[c.UPDATED_AT] = now;
    row[c.UPDATED_BY] = admin.name || admin.username;
    if (emailType) row[c.STATUS_EMAIL_AT] = now;
    env.orders.getRange(rowIndex, 1, 1, SHOP_ORDER_HEADERS.length).setValues([row]);
    emailOrder = _shopOrderFromRow_(row);
    return { success: true, message: 'Order updated.', order: emailOrder };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
    if (emailOrder && emailType && typeof sendShopOrderStatusEmail_ === 'function') {
      try { sendShopOrderStatusEmail_(emailOrder, emailType); } catch (mailErr) {
        Logger.log('shop status email error: ' + mailErr);
      }
    }
  }
}

// ============================================================
//  DAILY SHOP ORDERS ADMIN REPORT (counts in email + Excel detail)
// ============================================================

function _shopItemsSummaryText_(items) {
  return (items || []).map(function (item) {
    var name = String(item.product || '').trim();
    var opt = String(item.option || '').trim();
    var size = String(item.size || '').trim();
    var qty = Number(item.quantity || 1);
    var del = item.deliveredAt ? ' [Delivered]' : '';
    return name
      + (opt ? ' — ' + opt : '')
      + (size ? ' · Size ' + size : '')
      + ' × ' + qty
      + del;
  }).join('; ');
}

function _shopPaymentBucket_(status) {
  var s = String(status || 'Pending').trim().toLowerCase();
  if (s === 'paid' || s === 'verified') return 'Paid / Verified';
  if (s.indexOf('partial') >= 0) return 'Partially Paid';
  if (s.indexOf('submitted') >= 0) return 'Submitted';
  if (s === 'pending' || !s) return 'Pending';
  return String(status || 'Other').trim() || 'Other';
}

function _shopOrderStatusBucket_(status) {
  var s = String(status || 'Order Placed').trim();
  if (s === 'Delivered') return 'Delivered';
  if (s === 'Order Placed') return 'Order Placed';
  // In Progress, Partially Delivered, and any other active state
  return 'In Progress';
}

function _shopPaidEstimate_(order) {
  var total = Number(order.total || 0);
  if (order.paidAmount != null && order.paidAmount !== '') {
    var paid = Number(order.paidAmount || 0);
    return { paid: paid, balance: Math.max(total - paid, 0) };
  }
  var bucket = _shopPaymentBucket_(order.paymentStatus);
  if (bucket === 'Paid / Verified') return { paid: total, balance: 0 };
  if (bucket === 'Partially Paid') return { paid: '', balance: '' };
  return { paid: 0, balance: total };
}

/** Attach paid/balance from shopping ledger (read-only; does not rewrite order rows). */
function _attachShopPaidAmounts_(orders) {
  var env = _ensureShopSheets_();
  var paidByKe = {};
  var ledger = env.ss.getSheetByName(CONFIG.SHEETS.PAYMENTS);
  if (ledger && ledger.getLastRow() > 1) {
    var payData = ledger.getDataRange().getValues();
    for (var p = 1; p < payData.length; p++) {
      if (!_isShoppingPaymentType_(payData[p][CONFIG.LEDGER_COLS.PAYMENT_TYPE])) continue;
      var ke = String(payData[p][CONFIG.LEDGER_COLS.KE_NO] || '').trim().toUpperCase();
      if (!ke) continue;
      paidByKe[ke] = (paidByKe[ke] || 0) + Number(payData[p][CONFIG.LEDGER_COLS.AMOUNT] || 0);
    }
  }
  var byKe = {};
  orders.forEach(function (o) {
    var ke = String(o.keNo || '').trim().toUpperCase();
    if (!byKe[ke]) byKe[ke] = [];
    byKe[ke].push(o);
  });
  Object.keys(byKe).forEach(function (ke) {
    var list = byKe[ke].slice().sort(function (a, b) {
      return String(a.orderId || '').localeCompare(String(b.orderId || ''));
    });
    var allocation = _calculateShoppingAllocation_(
      list.map(function (o) { return Number(o.total || 0); }),
      paidByKe[ke] || 0
    );
    list.forEach(function (o, i) {
      o.paidAmount = allocation.allocations[i] || 0;
      o.balance = Math.max(Number(o.total || 0) - o.paidAmount, 0);
    });
  });
}

function _loadShopOrdersForDailyReport_() {
  var env = _ensureShopSheets_();
  var orders = [];
  if (!env.orders || env.orders.getLastRow() < 2) return orders;
  var data = env.orders.getRange(2, 1, env.orders.getLastRow() - 1, SHOP_ORDER_HEADERS.length).getValues();
  data.forEach(function (row) {
    if (!String(row[CONFIG.SHOP_ORDER_COLS.ORDER_ID] || '').trim()) return;
    orders.push(_shopOrderFromRow_(row));
  });
  _attachShopPaidAmounts_(orders);
  return orders;
}

function _buildShopOrdersDailyCounts_(orders) {
  var tz = Session.getScriptTimeZone();
  var todayLblShort = Utilities.formatDate(new Date(), tz, 'dd-MMM-yyyy');
  var counts = {
    total: orders.length,
    orderPlaced: 0,
    inProgress: 0,
    delivered: 0,
    newToday: 0,
    payment: {},
    orderValueTotal: 0,
    paidValueEstimate: 0
  };
  orders.forEach(function (o) {
    var statusBucket = _shopOrderStatusBucket_(o.orderStatus);
    if (statusBucket === 'Order Placed') counts.orderPlaced++;
    else if (statusBucket === 'Delivered') counts.delivered++;
    else counts.inProgress++;

    var payBucket = _shopPaymentBucket_(o.paymentStatus);
    counts.payment[payBucket] = (counts.payment[payBucket] || 0) + 1;

    var total = Number(o.total || 0);
    counts.orderValueTotal += total;
    var est = _shopPaidEstimate_(o);
    if (typeof est.paid === 'number') counts.paidValueEstimate += est.paid;

    if (String(o.createdAt || '').indexOf(todayLblShort) === 0) counts.newToday++;
  });
  return counts;
}

/**
 * Export a temporary Google Sheet as .xlsx via Drive API
 * (DriveApp.getAs(MICROSOFT_EXCEL) is not supported for Sheets).
 * Falls back to .csv if export fails — Excel opens CSV fine.
 */
function _exportSheetFileAsExcelBlob_(fileId, baseName) {
  var xlsxMime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  try {
    var url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId)
      + '/export?mimeType=' + encodeURIComponent(xlsxMime);
    var resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
      followRedirects: true
    });
    if (resp.getResponseCode() === 200) {
      var blob = resp.getBlob();
      blob.setName(baseName + '.xlsx');
      blob.setContentType(xlsxMime);
      return blob;
    }
    Logger.log('_exportSheetFileAsExcelBlob_ HTTP ' + resp.getResponseCode() + ': ' + resp.getContentText());
  } catch (e) {
    Logger.log('_exportSheetFileAsExcelBlob_ error: ' + e);
  }
  return null;
}

function _shopCsvEscape_(value) {
  var s = String(value == null ? '' : value);
  if (/[",\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function _buildShopOrdersCsvBlob_(rows, baseName) {
  var lines = rows.map(function (row) {
    return row.map(_shopCsvEscape_).join(',');
  });
  // BOM helps Excel open UTF-8 (₹ / names) correctly
  var csv = '\uFEFF' + lines.join('\r\n');
  return Utilities.newBlob(csv, 'text/csv', baseName + '.csv');
}

/**
 * Build spreadsheet attachment: prefer .xlsx via Drive export; else .csv.
 */
function _buildShopOrdersExcelBlob_(orders, dateLbl) {
  var safeDate = String(dateLbl || '').replace(/[^0-9A-Za-z_-]+/g, '_');
  var baseName = 'Shop_Orders_' + safeDate;
  var headers = [
    'Order Date', 'Order ID', 'KE No', 'Student Name', 'Email', 'Phone',
    'Products Ordered', 'Order Total (₹)', 'Payment Status', 'Amount Paid (₹)',
    'Balance (₹)', 'Order Status', 'Delivered?', 'Delivered At',
    'Expected Delivery', 'Parent Confirmation', 'Payment / UPI Ref', 'Updated At', 'Updated By'
  ];
  var rows = [headers];
  orders.forEach(function (o) {
    var est = _shopPaidEstimate_(o);
    var statusBucket = _shopOrderStatusBucket_(o.orderStatus);
    var deliveredLabel = statusBucket === 'Delivered' ? 'Yes'
      : (String(o.orderStatus || '').toLowerCase().indexOf('partial') >= 0 ? 'Partial' : 'No');
    rows.push([
      o.createdAt || '',
      o.orderId || '',
      o.keNo || '',
      o.riderName || '',
      o.email || '',
      o.phone || '',
      _shopItemsSummaryText_(o.items),
      Number(o.total || 0),
      o.paymentStatus || 'Pending',
      est.paid === '' ? '' : Number(est.paid || 0),
      est.balance === '' ? '' : Number(est.balance || 0),
      o.orderStatus || 'Order Placed',
      deliveredLabel,
      o.deliveredAt || '',
      o.expectedDate || '',
      o.parentConfirmation || '',
      o.paymentTxnRef || o.upiReference || '',
      o.updatedAt || '',
      o.updatedBy || ''
    ]);
  });

  var fileId = '';
  try {
    var temp = SpreadsheetApp.create('Shop_Orders_Daily_' + Date.now());
    fileId = temp.getId();
    var sheet = temp.getActiveSheet();
    sheet.setName('Shop Orders');
    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#1f4e3d').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();

    var xlsx = _exportSheetFileAsExcelBlob_(fileId, baseName);
    if (xlsx) return xlsx;
  } catch (e) {
    Logger.log('_buildShopOrdersExcelBlob_ sheet path: ' + e);
  } finally {
    if (fileId) {
      try { DriveApp.getFileById(fileId).setTrashed(true); } catch (ignoreTrash) {}
    }
  }

  // Reliable fallback — opens in Excel / Google Sheets
  return _buildShopOrdersCsvBlob_(rows, baseName);
}

function _buildShopOrdersDailyEmailHtml_(counts, dateLbl) {
  function row(label, value) {
    return '<tr><td style="padding:8px 12px;border-bottom:1px solid #e5e7eb;color:#374151">'
      + label + '</td><td style="padding:8px 12px;border-bottom:1px solid #e5e7eb;text-align:right;font-weight:700;color:#14330f">'
      + value + '</td></tr>';
  }
  var payRows = Object.keys(counts.payment || {}).sort().map(function (k) {
    return row(k, counts.payment[k]);
  }).join('');
  return '<!doctype html><html><body style="margin:0;background:#f4f7f2;font-family:Arial,sans-serif;color:#1b2118">'
    + '<div style="max-width:560px;margin:24px auto;background:#fff;border-radius:16px;overflow:hidden;border:1px solid #dce7d7">'
    + '<div style="padding:20px;background:#14330f;color:#fff">'
    + '<h2 style="margin:0;font-size:20px">Daily Shop Orders Report</h2>'
    + '<div style="margin-top:6px;color:#d5e8cf;font-size:13px">' + dateLbl + ' · '
    + (CONFIG.LOCATION_CITY || 'Hyderabad') + '</div></div>'
    + '<div style="padding:20px">'
    + '<p style="margin:0 0 14px;font-size:14px;color:#4b5563">Summary counts only. Full student / order details are in the attached Excel file.</p>'
    + '<table style="width:100%;border-collapse:collapse;margin-bottom:18px">'
    + row('Total orders', counts.total)
    + row('New orders today', counts.newToday)
    + row('Order Placed', counts.orderPlaced)
    + row('In Progress', counts.inProgress)
    + row('Delivered', counts.delivered)
    + '</table>'
    + '<div style="font-size:13px;font-weight:700;color:#14330f;margin:0 0 6px">Payment status</div>'
    + '<table style="width:100%;border-collapse:collapse;margin-bottom:18px">'
    + (payRows || row('—', '0'))
    + '</table>'
    + '<table style="width:100%;border-collapse:collapse">'
    + row('Total order value', '₹' + Number(counts.orderValueTotal || 0).toLocaleString('en-IN'))
    + row('Amount paid (ledger)', '₹' + Number(counts.paidValueEstimate || 0).toLocaleString('en-IN'))
    + '</table>'
    + '<p style="margin:18px 0 0;font-size:12px;color:#6b7280">Open the attached spreadsheet for who ordered, products, delivery and payment details.</p>'
    + '</div></div></body></html>';
}

/**
 * Daily admin email: shop order counts in the body + Excel attachment with full details.
 */
function sendDailyShopOrdersReport() {
  try {
    Logger.log('=== Daily Shop Orders Report START ===');
    var tz = Session.getScriptTimeZone();
    var today = new Date();
    var dateLbl = Utilities.formatDate(today, tz, 'EEEE, dd MMM yyyy');
    var ymd = Utilities.formatDate(today, tz, 'yyyy-MM-dd');

    var orders = _loadShopOrdersForDailyReport_();
    var counts = _buildShopOrdersDailyCounts_(orders);
    var excelBlob = _buildShopOrdersExcelBlob_(orders, ymd);
    var html = _buildShopOrdersDailyEmailHtml_(counts, dateLbl);

    var adminEmails = (typeof getAdminEmails === 'function') ? getAdminEmails() : [];
    if (!adminEmails.length) {
      Logger.log('No admin emails — skipping shop orders report');
      return { success: false, message: 'No admin emails in Mail Info.' };
    }
    var primary = adminEmails[0];
    var cc = adminEmails.slice(1);
    var subject = 'KE Shop Orders Daily · ' + (CONFIG.LOCATION_CITY || 'Hyderabad') + ': ' + dateLbl;

    try {
      var sent = sendMailKE_(primary, subject, html, {
        attachments: [excelBlob],
        cc: cc.join(','),
        textBody: 'Shop orders summary: total ' + counts.total
          + ', placed ' + counts.orderPlaced
          + ', in progress ' + counts.inProgress
          + ', delivered ' + counts.delivered
          + '. Details in attached Excel.'
      });
      if (typeof logEmail === 'function') {
        logEmail('ShopDailyReport', primary, cc.join(','), subject, '', 'Sent',
          'via ' + (sent && sent.provider) + ' orders=' + counts.total);
      }
    } catch (mailErr) {
      if (typeof logEmailFailed === 'function') {
        logEmailFailed('ShopDailyReport', primary, cc.join(','), subject, '', String(mailErr));
      }
      throw mailErr;
    }

    Logger.log('Shop orders daily report sent to ' + primary
      + ' (total=' + counts.total + ', delivered=' + counts.delivered + ')');
    return { success: true, counts: counts };
  } catch (err) {
    Logger.log('sendDailyShopOrdersReport ERROR: ' + err + '\n' + (err.stack || ''));
    return { success: false, message: String(err.message || err) };
  }
}

function testSendDailyShopOrdersReportNow() {
  var result = sendDailyShopOrdersReport();
  SpreadsheetApp.getUi().alert(
    result && result.success
      ? '✅ Shop orders daily report sent. Check admin inboxes (Excel attached).'
      : ('Could not send shop report: ' + ((result && result.message) || 'unknown error'))
  );
}
