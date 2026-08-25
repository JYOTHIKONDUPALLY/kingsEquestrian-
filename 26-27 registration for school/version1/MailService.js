// ============================================================
// KINGS EQUESTRIAN — UNIFIED MAIL SENDER
// File: MailService.gs
//
// sendMailKE_() is a drop-in replacement for MailApp.sendEmail that routes
// mail through the Brevo transactional API (via UrlFetchApp) when a
// BREVO_API_KEY is configured, and transparently FALLS BACK to MailApp if
// Brevo is unavailable, unconfigured, or returns an error.
//
// Why: MailApp/Gmail is capped at 100 recipients/day (consumer) or
// 1,500/day (Workspace). Sending through Brevo's API instead counts against
// the URL Fetch quota (20,000/day) and Brevo's own allowance (300/day free),
// which lifts the practical ceiling for a ~500-student school.
//
// SETUP (one time):
//   1. Apps Script editor → Project Settings (gear) → Script Properties
//      → add:  BREVO_API_KEY = xkeysib-....
//   2. In Brevo, verify your sender/domain so CONFIG.MAIL_FROM can send.
//   3. Run testBrevoEmail() (menu) to confirm delivery + which provider ran.
//
// If BREVO_API_KEY is absent, everything keeps working via MailApp exactly
// as before — this module is safe to deploy before the key is set.
// ============================================================

/** Read the Brevo API key from Script Properties (preferred) or CONFIG. */
function _getBrevoKey_() {
  try {
    var p = PropertiesService.getScriptProperties().getProperty('BREVO_API_KEY');
    if (p && String(p).trim()) return String(p).trim();
  } catch (e) {}
  try {
    if (typeof CONFIG !== 'undefined' && CONFIG.BREVO_API_KEY && String(CONFIG.BREVO_API_KEY).trim()) {
      return String(CONFIG.BREVO_API_KEY).trim();
    }
  } catch (e2) {}
  return '';
}

/** Sender identity used for Brevo (must be a verified sender/domain in Brevo). */
function _mailFrom_() {
  var email = (typeof CONFIG !== 'undefined' && CONFIG.MAIL_FROM) ? CONFIG.MAIL_FROM : Session.getEffectiveUser().getEmail();
  var name  = (typeof CONFIG !== 'undefined' && CONFIG.MAIL_FROM_NAME) ? CONFIG.MAIL_FROM_NAME : 'Kings Equestrian';
  return { email: email, name: name };
}

/** Normalise a "a@b.com, c@d.com" string OR array into Brevo [{email}] form. */
function _brevoAddrs_(v) {
  var out = [];
  if (!v) return out;
  var arr = Array.isArray(v) ? v : String(v).split(',');
  for (var i = 0; i < arr.length; i++) {
    var e = String(arr[i] || '').trim();
    if (e) out.push({ email: e });
  }
  return out;
}

/**
 * Unified send. Returns { ok, provider, id?, error? }.
 * @param {string}        to          primary recipient (single address)
 * @param {string}        subject
 * @param {string}        htmlBody
 * @param {Object}        [options]   { cc, bcc, replyTo, attachments:[Blob], textBody }
 */
function sendMailKE_(to, subject, htmlBody, options) {
  options = options || {};
  var cc = options.cc || '';
  var bcc = options.bcc || '';
  var replyTo = options.replyTo || '';
  var attachments = options.attachments || [];
  var textBody = options.textBody || '';

  var key = _getBrevoKey_();
  if (key) {
    try {
      var res = _sendViaBrevo_(key, to, subject, htmlBody, cc, bcc, replyTo, attachments);
      if (res && res.ok) return { ok: true, provider: 'brevo', id: res.id || '' };
      Logger.log('Brevo send failed (' + (res && res.error) + ') — falling back to MailApp.');
    } catch (e) {
      Logger.log('Brevo exception — falling back to MailApp: ' + e);
    }
  }

  // ── Fallback: Google MailApp ──
  var opts = { htmlBody: htmlBody };
  if (cc) opts.cc = Array.isArray(cc) ? cc.join(',') : cc;
  if (bcc) opts.bcc = Array.isArray(bcc) ? bcc.join(',') : bcc;
  if (replyTo) opts.replyTo = replyTo;
  if (attachments && attachments.length) opts.attachments = attachments;
  var from = _mailFrom_();
  if (from && from.name) opts.name = from.name;
  MailApp.sendEmail(to, subject, textBody || '', opts);
  return { ok: true, provider: 'mailapp' };
}

/** POST one message to Brevo's transactional email API. */
function _sendViaBrevo_(key, to, subject, htmlBody, cc, bcc, replyTo, attachments) {
  var payload = {
    sender: _mailFrom_(),
    to: _brevoAddrs_(to),
    subject: String(subject || ''),
    htmlContent: String(htmlBody || ' ')
  };
  if (!payload.to.length) return { ok: false, error: 'No valid recipient' };

  var ccArr = _brevoAddrs_(cc);
  if (ccArr.length) payload.cc = ccArr;
  var bccArr = _brevoAddrs_(bcc);
  if (bccArr.length) payload.bcc = bccArr;
  if (replyTo) payload.replyTo = { email: String(replyTo).trim() };

  if (attachments && attachments.length) {
    var atts = [];
    for (var i = 0; i < attachments.length; i++) {
      var b = attachments[i];
      if (!b || typeof b.getBytes !== 'function') continue;
      atts.push({
        name: (typeof b.getName === 'function' && b.getName()) ? b.getName() : ('attachment-' + (i + 1)),
        content: Utilities.base64Encode(b.getBytes())
      });
    }
    if (atts.length) payload.attachment = atts;
  }

  var resp = UrlFetchApp.fetch('https://api.brevo.com/v3/smtp/email', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'api-key': key, 'accept': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  var body = resp.getContentText();
  if (code >= 200 && code < 300) {
    var id = '';
    try { id = (JSON.parse(body) || {}).messageId || ''; } catch (e) {}
    return { ok: true, id: id };
  }
  return { ok: false, error: 'HTTP ' + code + ': ' + body };
}

// ────────────────────────────────────────────────────────────
//  DIAGNOSTICS (menu)
// ────────────────────────────────────────────────────────────

/** Show remaining MailApp quota + whether Brevo is configured. */
function checkEmailQuota() {
  var remaining = 'unknown';
  try { remaining = MailApp.getRemainingDailyQuota(); } catch (e) {}
  var brevo = _getBrevoKey_() ? 'CONFIGURED (Brevo API in use)' : 'NOT set (using Google MailApp only)';
  SpreadsheetApp.getUi().alert(
    'Email status\n\n'
    + 'Brevo API key: ' + brevo + '\n'
    + 'MailApp recipients remaining today: ' + remaining + '\n\n'
    + (typeof CONFIG !== 'undefined' && CONFIG.MAIL_FROM ? 'Sender (From): ' + CONFIG.MAIL_FROM + '\n' : '')
    + '\nNote: MailApp quota only matters for the fallback path. When Brevo is '
    + 'configured, sends go through Brevo and do not consume the MailApp quota.'
  );
}

/** Send a test email to the current user and report which provider handled it. */
function testBrevoEmail() {
  var ui = SpreadsheetApp.getUi();
  var me = Session.getEffectiveUser().getEmail();
  var resp = ui.prompt('Send test email',
    'Send a test email to (blank = ' + me + '):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var to = String(resp.getResponseText() || '').trim() || me;

  var html = '<div style="font-family:Arial,sans-serif;padding:20px">'
    + '<h2 style="color:#1f4617">Kings Equestrian — mail test</h2>'
    + '<p>If you can read this, sending works. Sent at ' + new Date() + '.</p></div>';

  try {
    var r = sendMailKE_(to, 'Kings Equestrian — mail test', html, {});
    if (typeof logEmail === 'function') logEmail('Test', to, '', 'Kings Equestrian — mail test', '', 'Sent', 'via ' + r.provider + (r.id ? ' id=' + r.id : ''));
    var from = _mailFrom_();
    ui.alert('Test accepted by ' + (r.provider === 'brevo' ? 'Brevo API ✅' : 'Google MailApp (fallback)') + '\n\n'
      + 'To: ' + to + '\n'
      + 'From (sender): ' + from.email + '\n'
      + (r.id ? 'Brevo message id: ' + r.id + '\n' : '')
      + '\n'
      + (r.provider === 'brevo'
          ? 'IMPORTANT: "accepted" is not "delivered". If it did not arrive:\n'
            + '1. In Brevo → Transactional → Logs, search this recipient / message id to see the REAL status (delivered / blocked / bounced).\n'
            + '2. New Brevo accounts are often under review — sending is paused until activated.\n'
            + '3. The sender "' + from.email + '" must be a VERIFIED sender (or its domain authenticated with DKIM) in Brevo.'
          : 'Brevo did NOT run — BREVO_API_KEY missing or Brevo errored (check execution log).'));
  } catch (e) {
    ui.alert('Test FAILED: ' + e);
  }
}
