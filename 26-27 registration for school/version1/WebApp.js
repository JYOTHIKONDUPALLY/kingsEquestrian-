// ============================================================
// Web app entry — Deploy → Manage deployments → Edit → New version
// URL: .../exec              → Stable Management
// URL: .../exec?app=portal   → My Rides portal
// URL: .../exec?app=shopimg&id=FILE_ID → shop product image (owner access)
// ============================================================

/**
 * Live UI version from server code (always fresh via google.script.run).
 * Cached HTML may still show an older page — clients use this to detect that.
 */
function getAppUiVersion() {
  return {
    version: String((typeof CONFIG !== 'undefined' && CONFIG.APP_UI_VERSION) || ''),
    label: String((typeof CONFIG !== 'undefined' && CONFIG.APP_UI_VERSION) || ''),
    serverTime: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
  };
}

function doGet(e) {
  e = e || {};
  var params = e.parameter || {};
  var app = params.app ? String(params.app) : 'attendance';

  // Product images: serve via script identity so parents don't need Drive access
  if (app === 'shopimg') {
    return _serveShopImage_(params.id);
  }

  if (app === 'portal') {
    return HtmlService
      .createHtmlOutput(getRiderPortalHtml())
      .setTitle('My Rides · Kings Equestrian')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1,maximum-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Serve Stable Management from HTML file AttendanceApp.
  // Requires AttendanceHTML.js (buildAttendanceApp_) + HTML file named AttendanceApp.
  if (typeof buildAttendanceApp_ === 'function') {
    return buildAttendanceApp_();
  }
  // Fallback if AttendanceHTML.js was not pasted into the project yet
  try {
    var t = HtmlService.createTemplateFromFile('AttendanceApp');
    t.logoUrl = (typeof CONFIG !== 'undefined' && CONFIG.LOGO_URL) ? CONFIG.LOGO_URL : '';
    t.locShort = (typeof schoolLocationShort_ === 'function') ? schoolLocationShort_() : 'Hyderabad';
    t.locCity = (typeof CONFIG !== 'undefined' && CONFIG.LOCATION_CITY) ? CONFIG.LOCATION_CITY : 'Hyderabad';
    t.buildStamp = String(Date.now());
    t.buildLabel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    t.appUiVersion = String((typeof CONFIG !== 'undefined' && CONFIG.APP_UI_VERSION) || '');
    return t.evaluate()
      .setTitle('Stable Management')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1,maximum-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><body style="font-family:sans-serif;padding:24px;color:#991b1b">'
      + '<h2>Stable Management failed to load</h2>'
      + '<p><code>buildAttendanceApp_</code> is missing. In Apps Script, add:</p>'
      + '<ol><li>Script file <b>AttendanceHTML</b> (paste AttendanceHTML.js)</li>'
      + '<li>HTML file named exactly <b>AttendanceApp</b> (paste AttendanceApp.html)</li></ol>'
      + '<p style="color:#6b7280;font-size:12px">' + String(err.message || err) + '</p>'
      + '</body></html>'
    ).setTitle('Setup needed');
  }
}

/**
 * Serve a Drive file as an HTML page wrapping a data-URI image.
 * Used when <img> cannot load private Drive links for parents who are not
 * logged into the owner's Google account.
 * Prefer catalog data-URLs from getShopCatalog; this is a fallback URL.
 */
function _serveShopImage_(fileId) {
  fileId = String(fileId || '').trim();
  try {
    if (!fileId || !/^[a-zA-Z0-9_-]+$/.test(fileId)) {
      throw new Error('Missing image id');
    }
    var dataUrl = _shopDriveThumbDataUrl_(fileId);
    if (!dataUrl) throw new Error('Could not load image');
    var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">'
      + '<meta name="viewport" content="width=device-width,initial-scale=1">'
      + '<style>html,body{margin:0;height:100%;background:#f3f4f6}'
      + 'img{display:block;width:100%;height:100%;object-fit:contain}</style></head>'
      + '<body><img src="' + dataUrl + '" alt=""></body></html>';
    return HtmlService.createHtmlOutput(html)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><body style="font-family:sans-serif;padding:16px;color:#991b1b">'
      + 'Image unavailable. Ask admin to share the Drive file with the script owner, '
      + 'or set sharing to Anyone with the link (Viewer).'
      + '<div style="margin-top:8px;font-size:12px;color:#6b7280">' + String(err.message || err) + '</div>'
      + '</body></html>'
    );
  }
}
