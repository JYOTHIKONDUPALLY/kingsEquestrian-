// ============================================================
// Stable Management UI — served as Apps Script HTML file
// ============================================================
// IMPORTANT:
//   createHtmlOutput(hugeString) wraps the page in document.write().
//   Once the UI grew past ~250KB (Feed/Tack V2), that bootstrap throws:
//     "Failed to execute 'write' on 'Document': Invalid or unexpected token"
//   Serve the page via createTemplateFromFile('AttendanceApp') instead.
//
// Apps Script setup:
//   1. File → New → HTML file → name it exactly: AttendanceApp
//   2. Paste the contents of AttendanceApp.html into that file
//   3. Keep this AttendanceHTML.gs (or AttendanceHTML.js) as a Script file
// ============================================================

/** Build the Stable Management HtmlOutput (do not wrap again in createHtmlOutput). */
function buildAttendanceApp_() {
  var t = HtmlService.createTemplateFromFile('AttendanceApp');
  t.logoUrl = (typeof CONFIG !== 'undefined' && CONFIG.LOGO_URL) ? CONFIG.LOGO_URL : '';
  t.locShort = (typeof schoolLocationShort_ === 'function') ? schoolLocationShort_() : 'Hyderabad';
  t.locCity = (typeof CONFIG !== 'undefined' && CONFIG.LOCATION_CITY) ? CONFIG.LOCATION_CITY : 'Hyderabad';
  // Changes every request so browsers / Google iframe don't keep an old UI after redeploy
  t.buildStamp = String(Date.now());
  t.buildLabel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  t.appUiVersion = String((typeof CONFIG !== 'undefined' && CONFIG.APP_UI_VERSION) || '');
  return t.evaluate()
    .setTitle('Stable Management')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1,maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Legacy helper — returns HTML text only.
 * Prefer buildAttendanceApp_() / doGet for serving; wrapping this string in
 * createHtmlOutput() is what triggers the document.write crash on large pages.
 */
function getAttendanceAppHtml() {
  return buildAttendanceApp_().getContent();
}
