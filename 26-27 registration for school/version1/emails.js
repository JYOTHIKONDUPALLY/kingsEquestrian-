// ============================================================
// INDUS EQUESTRIAN — SCHOOL SYSTEM
// File: 3_Emails.gs
// All email construction and sending functions
// ============================================================

/** Best-effort fee for UPI QR from service sheet; 0 = open amount in payer app. */
function _welcomeFeeAmountForUpi(serviceProgram) {
  const label = String(serviceProgram || '').trim().toLowerCase();
  if (!label) return 0;
  try {
    const map = getPricingData();
    let best = 0;
    for (const name in map) {
      const key = String(name || '').trim().toLowerCase();
      if (!key || key.length < 3) continue;
      if (label.indexOf(key) !== -1 || key.indexOf(label.substring(0, Math.min(20, label.length))) !== -1) {
        const p = Number(map[name].price);
        if (!isNaN(p) && p > best) best = p;
      }
    }
    return best;
  } catch (e) {
    return 0;
  }
}

// ────────────────────────────────────────────────────────────
//  WELCOME / REGISTRATION EMAIL
//  Sent when a student submits the registration form.
//  Attaches the appropriate consent PDF (school or summer).
//  Includes a prefilled link to the payment form.
// ────────────────────────────────────────────────────────────

/**
 * Build the consent PDF attachment for a registration, with retries.
 * DocumentApp/DriveApp calls can fail transiently under form-submit triggers,
 * so we attempt a few times before giving up. Returns { pdf, error }.
 */
function _buildConsentAttachment_(d, attempts) {
  attempts = attempts || 3;
  var lastErr = '';
  for (var i = 1; i <= attempts; i++) {
    try {
      var consentPDF = generateConsentPDF(d.program, {
        studentName     : d.studentName,
        parentName      : d.parentName,
        email           : d.email,
        phone           : d.phone,
        grade           : d.grade,
        dob             : d.dob,
        address         : d.address,
        motherName      : d.motherName,
        fatherName      : d.fatherName,
        motherContact   : d.motherContact,
        motherWhatsApp  : d.motherWhatsApp,
        fatherContact   : d.fatherContact,
        fatherWhatsApp  : d.fatherWhatsApp,
        emergencyContact: d.emergencyContact,
        relationship    : d.relationship,
        consentDate     : d.consentDate,
        horseLease      : d.horseLease,
        serviceProgram  : d.serviceProgram,
        sessionDates    : ''     // uses default from CONFIG.ACADEMIC_YEAR_LABEL
      });
      if (consentPDF) return { pdf: consentPDF, error: '' };
      lastErr = 'generateConsentPDF returned nothing (attempt ' + i + ')';
    } catch (e) {
      lastErr = 'attempt ' + i + ': ' + e + (e && e.stack ? ('\n' + e.stack) : '');
      Logger.log('Consent PDF error (program=' + String(d.program) + '): ' + lastErr);
    }
    if (i < attempts) Utilities.sleep(1200); // brief pause before retrying transient failures
  }
  return { pdf: null, error: lastErr };
}

/**
 * Sends the school registration welcome mail via MailApp only (scope: script.send_mail).
 * Named distinctly so another script file / library cannot override `sendWelcomeEmail` with GmailApp.
 */
function sendIndusSchoolWelcomeEmail(d) {
  /*  d: {
        studentName, parentName, email, phone, grade,
        program, serviceProgram, horseLease,
        dob, address, motherName, fatherName,
        motherContact, motherWhatsApp, fatherContact, fatherWhatsApp,
        emergencyContact, relationship, consentDate, timestamp,
        keNo, payFormUrl, isFirstTime, sheet, row
      }
  */

  const attachments = [];

  // ── Consent PDF — school vs. summer ──────────────────────
  // Generation uses DocumentApp + DriveApp (+ UrlFetchApp for logos). Under an
  // installable form-submit trigger these can fail transiently or if the
  // trigger was created before those scopes were authorized — which silently
  // dropped the attachment before. We now retry, and if it still fails we log
  // a visible FAILED row to the Email Log instead of swallowing the error.
  const consentResult = _buildConsentAttachment_(d);
  if (consentResult.pdf) {
    attachments.push(consentResult.pdf);
  } else {
    logEmailFailed('Welcome-Consent', d.email, '', 'Consent PDF not attached (' + d.keNo + ')', d.keNo,
      consentResult.error || 'Consent PDF generation returned nothing');
  }

  // ── Optional program info brochure ───────────────────────
  const brochureLink = _isSummerTrackProgram(d.program)
    ? (CONFIG.SUMMER_PROGRAM_INFO_DOC_ID || '')
    : (CONFIG.SCHOOL_PROGRAM_INFO_DOC_ID || '');

  const brochureBlock = brochureLink
    ? '<div style="background:#e8f4ff;border-left:4px solid #2196f3;padding:13px 18px;margin:16px 0;border-radius:4px">'
      + '<p style="margin:0;font-size:13px"><strong>Program information:</strong><br>'
      + '<a href="' + brochureLink + '" style="color:#1565c0;font-weight:600" target="_blank">View Program Guide</a>'
      + '<br><br><strong>Track your sessions:</strong><br>'
      + '<a href="' + CONFIG.MY_RIDES_PORTAL_URL + '" style="color:#1565c0;font-weight:600" target="_blank">My Rides Portal</a>'
      + '</p></div>'
    : '';

  const subject = d.isFirstTime
    ? 'Welcome to Kings Equestrian · ' + (CONFIG.LOCATION_CITY || 'Hyderabad') + '! Registration Ref: ' + d.keNo
    : 'Registration Confirmed – Kings Equestrian ' + (CONFIG.LOCATION_CITY || 'Hyderabad') + ' (' + d.keNo + ')';

  const greeting = d.isFirstTime
    ? '<h2 style="color:#1f4e3d;margin:0 0 8px">Welcome, ' + d.studentName + '!</h2>'
      + '<p>Your registration with <strong>Kings Equestrian at ' + schoolLocationShort_() + '</strong> has been received. '
      + 'Your <strong>Registration Ref is ' + d.keNo + '</strong> — keep this safe for all future correspondence and payments.</p>'
    : '<h2 style="color:#1f4e3d;margin:0 0 8px">Hi ' + d.studentName + '!</h2>'
      + '<p>A new registration has been received for your account <strong>(' + d.keNo + ')</strong>.</p>';

  // ── program badge ─────────────────────────────────────────
  const programLabel = String(d.serviceProgram || d.program || '').trim();
  const isSummer     = _isSummerTrackProgram(d.program);
  const badgeColor   = isSummer ? '#ff7043' : '#1f4e3d';

  const feeForQr = _welcomeFeeAmountForUpi(programLabel);
  const upiLink  = createUPILink(feeForQr, d.keNo);
  const qrCodeUrl = createQRCode(upiLink);

  const htmlBody =
    '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>'
  + '<body style="font-family:Arial,sans-serif;background:#f5f5f5;margin:0;padding:0;color:#333">'
  + '<div style="max-width:640px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,.1)">'

  // header
  + '  <div style="background:linear-gradient(135deg,#1f4e3d,#4f9c7a);padding:28px 30px;text-align:center;color:#fff">'
  + '    <img src="' + CONFIG.LOGO_URL + '" style="width:72px;height:72px;border-radius:50%;border:3px solid #000;margin-bottom:12px">'
  + '    <h1 style="margin:0;font-size:22px">Indus Equestrian Centre of Excellence</h1>'
  + '    <p style="margin:4px 0 0;font-size:13px;opacity:.9">' + schoolLocationShort_() + ' · Academic Year ' + CONFIG.ACADEMIC_YEAR_LABEL + '</p>'
  + '  </div>'

  // body
  + '  <div style="padding:28px 30px">'
  + '    ' + greeting

  // registration summary box
  + '    <div style="background:#f0f8f0;border-left:4px solid #2c5f2d;padding:14px 18px;margin:20px 0;border-radius:4px">'
  + '      <p style="margin:0;font-size:13px;line-height:1.9">'
  + '      <strong>Reg Ref:</strong> <span style="font-size:20px;color:#1f4e3d;font-weight:bold">' + d.keNo + '</span><br>'
  + '      <strong>Student:</strong> ' + d.studentName + '<br>'
  + '      <strong>Parent / Guardian:</strong> ' + (d.parentName || '—') + '<br>'
  + '      <strong>Grade:</strong> ' + (d.grade || '—') + '<br>'
  + '      <strong>Program:</strong> <span style="background:' + badgeColor + ';color:#fff;padding:2px 8px;border-radius:10px;font-size:12px">' + (programLabel || 'School Program') + '</span>'
  + '      </p>'
  + '    </div>'

  // brochure + portal links
  + '    ' + brochureBlock

  // payment section (UPI QR + form — same pattern as Kings Farm welcome)
  + '    <div style="background:#e8f5e9;border:2px solid #4caf50;padding:20px;border-radius:8px;margin:20px 0">'
  + '      <h3 style="color:#2e7d32;margin:0 0 10px">Make Your Fee Payment'
  + (feeForQr > 0 ? ' — Rs.' + Number(feeForQr).toLocaleString('en-IN') : '')
  + '</h3>'
  + '      <p style="font-size:13px;color:#555;margin:0 0 12px">Scan the QR code with any UPI app (reference: <strong>' + d.keNo + '</strong>), then open the payment form to upload your screenshot and details.</p>'
  + (feeForQr <= 0
    ? '      <p style="font-size:12px;color:#555;margin:0 0 12px;text-align:center">Enter the fee amount shown in your program / fee letter if your app asks for an amount.</p>'
    : '')
  + '      <div style="text-align:center;margin:16px 0">'
  + '        <img src="' + qrCodeUrl + '" alt="UPI QR Code" style="width:160px;height:160px;border:2px solid #c8e6c9;border-radius:8px">'
  + '      </div>'
  + '      <p style="font-size:12px;color:#444;margin:0 0 16px;text-align:center">UPI ID: <strong style="color:#1f4e3d">' + CONFIG.UPI_ID + '</strong></p>'
  + '      <div style="text-align:center">'
  + '        <a href="' + (d.payFormUrl || CONFIG.PAYMENT_FORM_BASE_URL) + '" style="background:#1f4e3d;color:#fff;padding:13px 30px;text-decoration:none;border-radius:6px;font-weight:bold;font-size:14px;display:inline-block">Open payment form</a>'
  + '      </div>'
  + '      <p style="font-size:11px;color:#777;margin:12px 0 0;text-align:center">Bank transfer is fine too — use the same registration reference. Upload proof via the form.</p>'
  + '    </div>'

  // what's next
  + '    <div style="background:#f9f9f9;padding:16px;border-radius:8px">'
  + '      <h4 style="color:#1f4e3d;margin:0 0 10px">Next Steps</h4>'
  + '      <ol style="margin:0;padding-left:20px;font-size:13px;color:#555;line-height:2">'
  + '        <li>Sign and return the attached Consent Form (you may email a photo/scan)</li>'
  + '        <li>Pay fees via UPI (scan QR above) or bank transfer, then submit the payment form with screenshot</li>'
  + '        <li>Await your payment receipt and schedule confirmation</li>'
  + '        <li>Arrive 10 minutes before your first class — helmet and appropriate footwear are mandatory</li>'
  + '      </ol>'
  + '    </div>'

  + '  </div>'

  // footer
  + '  <div style="background:#1f4e3d;color:#fff;padding:18px 30px;text-align:center;font-size:12px">'
  + '    ' + emailFooterHtml_()
  + '  </div>'
  + '</div>'
  + '</body></html>';

  const ccEmails = getCCRecipients('welcome');
  try {
    const r = sendMailKE_(d.email, subject, htmlBody, {
      attachments : attachments,
      cc          : ccEmails.join(',')
    });
    logEmail('Welcome', d.email, ccEmails.join(','), subject, d.keNo, 'Sent', 'via ' + r.provider);
  } catch (e) {
    logEmailFailed('Welcome', d.email, ccEmails.join(','), subject, d.keNo, String(e));
    throw e;
  }

  // Mark sheet row
  if (d.sheet && d.row) {
    d.sheet.getRange(d.row, CONFIG.REG_COLS.WELCOME_SENT + 1)
      .setValue('Yes').setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
    d.sheet.getRange(d.row, CONFIG.REG_COLS.WELCOME_AT + 1)
      .setValue(new Date()).setNumberFormat('dd-MMM-yyyy HH:mm:ss');
  }

  Logger.log('Welcome email sent to ' + d.email + ' (KE: ' + d.keNo + ')');
}

/**
 * Run once from the menu after updating appsscript.json — triggers OAuth for Documents + Mail + Drive.
 * Fixes "DocumentApp.create … documents" and incomplete mail authorization (often reported as gmail.*).
 */
function authorizeIndusScopesOnce() {
  const ui = SpreadsheetApp.getUi();
  const addr = Session.getActiveUser().getEmail();
  let tmpId = '';
  try {
    const doc = DocumentApp.create('_Indus_auth_delete_me_');
    doc.getBody().appendParagraph('authorization probe');
    doc.saveAndClose();
    tmpId = doc.getId();
    DriveApp.getFileById(tmpId).setTrashed(true);
  } catch (e) {
    ui.alert(
      'DocumentApp check failed (needs Google Docs scope).\n\n'
        + String(e)
        + '\n\nEnsure appsscript.json is in this project with oauthScopes including '
        + 'https://www.googleapis.com/auth/documents — then Save and run this again.'
    );
    return;
  }
  try {
    MailApp.sendEmail(addr, 'Indus registration script — Docs + Mail OK', 'OAuth probe succeeded.');
  } catch (e) {
    ui.alert('MailApp check failed:\n' + String(e) + '\n\nSave project, run this item again, and approve all prompts.');
    return;
  }
  ui.alert('Success.\n\n• Temporary Doc was created and trashed.\n• Test mail sent to:\n' + addr + '\n\nRetry Resend Welcome Email.');
}

/** @deprecated Use authorizeIndusScopesOnce */
function authorizeIndusMailOnce() {
  authorizeIndusScopesOnce();
}

/**
 * Diagnostic: generate a sample consent PDF and email it to the running user.
 * Confirms DocumentApp/DriveApp/UrlFetchApp are authorized so the welcome mail
 * can attach the consent form. Run this in a new account to force the OAuth
 * prompt and verify attachment generation end-to-end.
 */
function testConsentPdfGeneration() {
  const ui = SpreadsheetApp.getUi();
  const addr = Session.getActiveUser().getEmail();
  const sample = {
    program        : 'school',
    studentName    : 'Test Student',
    parentName     : 'Test Parent',
    email          : addr,
    phone          : '9999999999',
    grade          : '6',
    address        : 'Test Address',
    consentDate    : new Date(),
    serviceProgram : 'Regular School Classes (' + CONFIG.ACADEMIC_YEAR_LABEL + ')'
  };
  const res = _buildConsentAttachment_(sample, 2);
  if (!res.pdf) {
    ui.alert(
      'Consent PDF generation FAILED ❌\n\n' + (res.error || 'Unknown error') +
      '\n\nFix: run "Authorize script — Docs + mail (run once)", approve ALL prompts, ' +
      'then re-install triggers via "Setup All Triggers" (installable triggers keep the ' +
      'authorization they had when created).'
    );
    return;
  }
  try {
    MailApp.sendEmail(addr, 'KE consent PDF test — OK', 'The attached consent PDF was generated successfully.',
      { attachments: [res.pdf] });
    ui.alert('Success ✅\n\nA test consent PDF was generated and emailed to:\n' + addr +
      '\n\nWelcome emails will now attach the consent form. If real registrations still ' +
      'miss it, re-install triggers via "Setup All Triggers".');
  } catch (e) {
    ui.alert('Consent PDF generated, but sending the test mail failed:\n' + String(e));
  }
}

// ────────────────────────────────────────────────────────────
//  RECEIPT EMAIL HTML
// ────────────────────────────────────────────────────────────

function buildReceiptEmailHTML(d) {
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
    + '<body style="font-family:Arial,sans-serif;background:#f4f4f4;margin:0;padding:0;color:#333">'
    + '<div style="max-width:620px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 10px rgba(0,0,0,.1)">'
    + '  <div style="background:linear-gradient(135deg,#1f4e3d,#4f9c7a);padding:28px 30px;text-align:center;color:#fff">'
    + '    <img src="' + CONFIG.LOGO_URL + '" style="width:72px;height:72px;border-radius:50%;border:3px solid #000;margin-bottom:12px">'
    + '    <h1 style="margin:0;font-size:22px">Kings Equestrian Foundation</h1>'
    + '    <p style="margin:4px 0 0;font-size:12px;opacity:.9"><strong>' + (CONFIG.LOCATION_CODE || 'HYD') + '</strong> · ' + schoolLocationShort_() + ' · ' + CONFIG.ACADEMIC_YEAR_LABEL + '</p>'
    + '  </div>'
    + '  <div style="padding:28px 30px">'
    + '    <p style="font-size:16px">Dear <strong>' + d.name + '</strong>,</p>'
    + '    <div style="background:#d4edda;border-left:4px solid #28a745;padding:16px;border-radius:6px;margin:18px 0;text-align:center">'
    + '      <div style="font-size:20px;font-weight:bold;color:#155724">Payment Confirmed!</div>'
    + '      <p style="margin:6px 0 0;font-size:13px">Your 80G receipt is attached to this email.</p>'
    + '    </div>'
    + '    <table style="width:100%;border-collapse:collapse;font-size:13px;margin:16px 0">'
    + '      <tr><td style="padding:7px 0;color:#666;width:40%">Registration Ref</td><td style="padding:7px 0;font-weight:600;color:#1f4e3d">' + d.keNo + '</td></tr>'
    + '      <tr><td style="padding:7px 0;color:#666">Receipt No</td><td style="padding:7px 0;font-weight:600">' + d.receiptNo + '</td></tr>'
    + '      <tr><td style="padding:7px 0;color:#666">Amount Paid</td><td style="padding:7px 0;font-size:18px;font-weight:700;color:#28a745">Rs.' + Number(d.amount).toLocaleString('en-IN') + '</td></tr>'
    + (d.txnRef ? '      <tr><td style="padding:7px 0;color:#666">Transaction ID</td><td style="padding:7px 0;font-weight:600">' + d.txnRef + '</td></tr>' : '')
    + (d.payDate && fmtDate(d.payDate) ? '      <tr><td style="padding:7px 0;color:#666">Payment Date</td><td style="padding:7px 0">' + fmtDate(d.payDate) + '</td></tr>' : '')
    + '    </table>'
    + '    <div style="background:#fff8e1;border-left:4px solid #ffc107;padding:13px;border-radius:4px;font-size:12px;color:#856404">'
    + '      This receipt is eligible for deduction under Section 80G of the Income Tax Act, 1961.'
    + '    </div>'
    + '  </div>'
    + '  <div style="background:#1f4e3d;color:#fff;padding:16px 30px;text-align:center;font-size:12px">'
    + '    ' + emailFooterHtml_()
    + '  </div>'
    + '</div>'
    + '</body></html>';
}

// ────────────────────────────────────────────────────────────
//  80G RECEIPT PDF GENERATOR
// ────────────────────────────────────────────────────────────

function generate80GReceipt(riderName, pan, amount, txnRef, receiptNo, paymentDate) {
  const logoB64  = imgBase64FromUrl(CONFIG.LOGO_URL);
  const stampB64 = imgBase64FromDrive(CONFIG.STAMP_FILE_ID);
  const signB64  = imgBase64FromDrive(CONFIG.SIGN_FILE_ID);
  let receiptDate = paymentDate ? new Date(paymentDate) : new Date();
  if (isNaN(receiptDate.getTime())) receiptDate = new Date();
  const dateStr  = Utilities.formatDate(receiptDate, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const words    = numberToWords(amount);

  const html = '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    + '<style>'
    + '@page{size:A4;margin:0}'
    + 'body{font-family:"Times New Roman",serif;margin:0;padding:24px;background:#fff}'
    + '.box{border:2px solid #000;border-radius:32px;padding:24px 28px;max-width:800px;margin:auto;position:relative}'
    + '.hdr{display:flex;align-items:flex-start;gap:14px}'
    + '.logo{width:100px;text-align:center}.logo img{width:90px}'
    + '.hdr-c{flex:1;text-align:center}'
    + '.org{font-size:26px;font-weight:bold}'
    + '.reg{font-size:12px;margin-top:4px}'
    + '.sub{font-size:12px;margin-top:3px}'
    + '.tagline{margin-top:8px;font-style:italic;font-weight:bold;text-decoration:underline;font-size:12px}'
    + '.rno{position:absolute;right:28px;top:14px;font-size:15px;font-weight:bold;color:red}'
    + '.rbox{border:2px solid #000;border-radius:10px;text-align:center;padding:8px;margin:18px 0 8px}'
    + '.rtitle{font-size:19px;font-weight:bold}'
    + '.rsub{font-size:11px}'
    + '.date-r{text-align:right;font-size:13px;margin-bottom:8px}'
    + '.mcols{display:flex;gap:28px;margin-top:8px}'
    + '.col{flex:1;font-size:13px}'
    + '.cb{display:inline-block;width:12px;height:12px;border:1px solid #000;margin-right:5px;vertical-align:middle}'
    + '.cb.on{background:#000;position:relative}'
    + '.cb.on::after{content:"v";color:#fff;font-size:10px;position:absolute;left:0px;top:-3px}'
    + '.dr{margin:6px 0}.dl{font-weight:bold}'
    + '.amt-box{border:2px solid #000;margin:18px 0;padding:16px;position:relative;text-align:center}'
    + '.rs{position:absolute;left:18px;top:50%;transform:translateY(-50%);font-size:38px;color:goldenrod;font-weight:bold}'
    + '.av{font-size:32px;font-weight:bold}'
    + '.pm{font-size:13px;margin-top:8px}'
    + '.decl{margin-top:14px;font-size:12px;text-align:justify}'
    + '.sig{margin-top:32px;text-align:right}'
    + '.org-lbl{font-weight:bold;margin-bottom:4px}'
    + '.sign-area{position:relative;height:110px}'
    + '.sign-area img.sig-img{width:100px}'
    + '.sign-area img.stp-img{width:110px}'
    + '</style></head>'
    + '<body><div class="box">'
    + '  <div class="rno">' + receiptNo + '</div>'
    + '  <div class="hdr">'
    + '    <div class="logo"><img src="' + logoB64 + '"></div>'
    + '    <div class="hdr-c">'
    + '      <div class="org">Kings Equestrian Foundation</div>'
    + '      <div class="reg">Registered u/s 80G | Reg No: AAJCK7191GE20231 | PAN: AAJCK7191G</div>'
    + '      <div class="sub"><strong>Location: ' + (CONFIG.LOCATION_CODE || 'HYD') + '</strong> · ' + (CONFIG.LOCATION_CITY || 'Hyderabad') + '</div>'
    + '      <div class="sub">' + (CONFIG.BUSINESS_ADDRESS || locationCityState_()) + '</div>'
    + '      <div class="sub">' + (CONFIG.CONTACT_EMAIL || CONFIG.MAIL_FROM || '') + ' | ' + (CONFIG.CONTACT_PHONE || '') + '</div>'
    + '      <div class="tagline">We gratefully acknowledge your generous contribution in support of our programmes.</div>'
    + '    </div>'
    + '  </div>'
    + '  <div class="rbox">'
    + '    <div class="rtitle">Receipt</div>'
    + '    <div class="rsub">Issued in compliance with Rule 18AB and Form 10BD requirements</div>'
    + '  </div>'
    + '  <div class="date-r"><strong>Payment Date:</strong> ' + dateStr + '</div>'
    + '  <div class="mcols">'
    + '    <div class="col">'
    + '      <div style="font-weight:bold;margin-bottom:8px">Donor Category (tick Applicable)</div>'
    + '      <div><span class="cb on"></span> Resident Indian Donor</div>'
    + '      <div style="margin-top:5px"><span class="cb"></span> Non-Resident Indian (NRI)</div>'
    + '    </div>'
    + '    <div class="col">'
    + '      <div style="font-weight:bold;margin-bottom:8px">Donor Details</div>'
    + '      <div class="dr"><span class="dl">Name:</span> ' + riderName + '</div>'
    + '      <div class="dr"><span class="dl">PAN / Aadhaar:</span> ' + (pan || 'N/A') + '</div>'
    + '      <div class="dr"><span class="dl">Amount in Words:</span> ' + words + '</div>'
    + '    </div>'
    + '  </div>'
    + '  <div class="amt-box">'
    + '    <span class="rs">Rs.</span>'
    + '    <div class="av">' + Number(amount).toLocaleString('en-IN') + '</div>'
    + '  </div>'
    + '  <div class="pm">'
    + '    <strong>Mode of Payment:</strong> NEFT / RTGS / UPI (Cash not eligible u/s 80G)<br><br>'
    + (txnRef && txnRef !== 'N/A' ? '    <strong>Transaction Reference:</strong> ' + txnRef + '<br><br>' : '')
    + '    <strong>Amount in Words:</strong> ' + words
    + '  </div>'
    + '  <div class="decl">'
    + '    Certified that the above donation is received by trust for charitable purposes only. '
    + '    This donation is eligible for deduction under Section 80G of the Income Tax Act, 1961. '
    + '    This receipt will be reported in Form 10BD and Form 10BE will be issued to the donor.'
    + '  </div>'
    + '  <div class="sig">'
    + '    <div class="org-lbl">For Kings Equestrian Foundation</div>'
    + '    <div class="sign-area">'
    + '      <img class="sig-img" src="' + signB64 + '">'
    + '      <img class="stp-img" src="' + stampB64 + '">'
    + '    </div>'
    + '  </div>'
    + '</div></body></html>';

  const tmp  = DriveApp.createFile('receipt_temp_' + Date.now() + '.html', html, MimeType.HTML);
  const blob = tmp.getAs('application/pdf');
  blob.setName('80G_Receipt_' + riderName.replace(/\s+/g,'_') + '_' + receiptNo.replace(/\//g,'-') + '.pdf');
  tmp.setTrashed(true);
  return blob;
}

// ────────────────────────────────────────────────────────────
//  ABSENT / NO-SHOW NOTIFICATION
//  Sent to the rider's registered email when a trainer marks the
//  rider No-Show for a group session and saves.
//  info: { keNo, name, email, level, classNumber, title, date, timeSlot, scoredBy }
// ────────────────────────────────────────────────────────────

function sendAbsentNotificationEmail_(info) {
  info = info || {};
  var to = String(info.email || '').trim();
  if (!to) return false;

  var name = String(info.name || '').trim() || 'Rider';
  var keNo = String(info.keNo || '').trim();
  var dateLabel = '';
  try { dateLabel = info.date ? fmtDate(new Date(info.date)) : ''; } catch (e) { dateLabel = String(info.date || ''); }
  var timeSlot = String(info.timeSlot || '').trim();
  var level = String(info.level || '').trim();
  var classNo = String(info.classNumber || '').trim();
  var title = String(info.title || '').trim();

  var classLine = '';
  if (level || classNo) {
    classLine = (level ? level : '') + (classNo ? (level ? ' · ' : '') + 'Class ' + classNo : '') + (title ? ' — ' + title : '');
  }

  var subject = 'Missed Riding Class' + (dateLabel ? ' on ' + dateLabel : '') + (keNo ? ' — ' + keNo : '');

  var htmlBody =
    '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>'
  + '<body style="font-family:Arial,sans-serif;background:#f5f5f5;margin:0;padding:0;color:#333">'
  + '<div style="max-width:600px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,.1)">'
  + '  <div style="background:linear-gradient(135deg,#14330f,#1f4617);padding:24px 30px;text-align:center;color:#fff;border-bottom:3px solid #a9781a">'
  + '    <img src="' + CONFIG.LOGO_URL + '" style="width:64px;height:64px;border-radius:12px;border:2px solid #e9d59a;margin-bottom:10px;background:#fff;padding:4px">'
  + '    <h1 style="margin:0;font-size:20px">Attendance Update</h1>'
  + '    <p style="margin:4px 0 0;font-size:12px;opacity:.9">Kings Equestrian · ' + schoolLocationShort_() + ' · ' + CONFIG.ACADEMIC_YEAR_LABEL + '</p>'
  + '  </div>'
  + '  <div style="padding:26px 30px">'
  + '    <p style="font-size:15px;margin:0 0 12px">Dear <strong>' + name + '</strong>,</p>'
  + '    <p style="font-size:14px;line-height:1.6;margin:0 0 16px">Our records show that you were marked <strong style="color:#991b1b">absent (No-Show)</strong> for the following riding session:</p>'
  + '    <div style="background:#eaf1e7;border-left:4px solid #1f4617;padding:14px 18px;border-radius:4px;margin:0 0 18px">'
  + '      <p style="margin:0;font-size:13px;line-height:1.9">'
  + (keNo ? '        <strong>Reg Ref:</strong> ' + keNo + '<br>' : '')
  + (dateLabel ? '        <strong>Date:</strong> ' + dateLabel + '<br>' : '')
  + (timeSlot ? '        <strong>Time:</strong> ' + timeSlot + '<br>' : '')
  + (classLine ? '        <strong>Class:</strong> ' + classLine : '')
  + '      </p>'
  + '    </div>'
  + '    <p style="font-size:13px;line-height:1.6;color:#555;margin:0 0 8px">Missed classes can be made up in that same week/weekend where possible, and will otherwise lapse. Please reach out to your trainer to arrange a make-up session.</p>'
  + '    <div style="text-align:center;margin:20px 0 6px">'
  + '      <a href="' + CONFIG.MY_RIDES_PORTAL_URL + '" style="background:#1f4617;color:#fff;padding:12px 26px;text-decoration:none;border-radius:6px;font-weight:bold;font-size:13px;display:inline-block">Open My Rides Portal</a>'
  + '    </div>'
  + '  </div>'
  + '  <div style="background:#14330f;color:#fff;padding:16px 30px;text-align:center;font-size:12px">'
  + '    ' + emailFooterHtml_()
  + '  </div>'
  + '</div></body></html>';

  try {
    const r = sendMailKE_(to, subject, htmlBody, {});
    logEmail('Absent', to, '', subject, keNo, 'Sent', 'via ' + r.provider);
    return true;
  } catch (e) {
    logEmailFailed('Absent', to, '', subject, keNo, String(e));
    return false;
  }
}

// ────────────────────────────────────────────────────────────
//  SHOP ORDER EMAILS (merchandise — never an 80G receipt)
// ────────────────────────────────────────────────────────────

function _shopEmailEsc_(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function _shopEmailItems_(items) {
  return (items || []).map(function (item) {
    var label = _shopEmailEsc_(item.product);
    if (item.option && item.option !== 'Standard') label += ' — ' + _shopEmailEsc_(item.option);
    return '<tr><td style="padding:9px;border-bottom:1px solid #e5e7eb">' + label
      + '<div style="font-size:12px;color:#6b7280">Size: ' + _shopEmailEsc_(item.size)
      + ' · Qty: ' + Number(item.quantity || 1) + '</div></td>'
      + '<td style="padding:9px;text-align:right;border-bottom:1px solid #e5e7eb">₹'
      + Number(item.lineTotal || 0).toLocaleString('en-IN') + '</td></tr>';
  }).join('');
}

function sendShopOrderPlacedEmail_(order) {
  var to = String(order && order.email || '').trim();
  if (!to) return false;
  var paymentDue = Number(order.paymentDue == null ? order.total : order.paymentDue);
  var upiLink = paymentDue > 0 ? (order.upiLink || createUPILink(paymentDue, order.upiReference || order.orderId)) : '';
  var qrUrl = upiLink ? (order.qrUrl || createQRCode(upiLink)) : '';
  var subject = 'Order placed ' + order.orderId + (paymentDue > 0
    ? ' — shopping balance ₹' + paymentDue.toLocaleString('en-IN')
    : ' — covered by existing payment credit');
  var paymentBlock = paymentDue > 0
    ? '<div style="font-size:22px;font-weight:bold;text-align:right;color:#14330f">Shopping balance: ₹'
      + paymentDue.toLocaleString('en-IN') + '</div>'
      + '<div style="text-align:center;padding:20px 0"><img src="' + _shopEmailEsc_(qrUrl)
      + '" width="210" height="210" alt="Payment QR" style="max-width:100%"><br>'
      + '<strong>UPI: ' + _shopEmailEsc_(CONFIG.UPI_ID) + '</strong><br>'
      + '<span style="font-size:12px;color:#6b7280">Reference: ' + _shopEmailEsc_(order.upiReference || order.orderId) + '</span></div>'
      + '<p style="text-align:center"><a href="' + _shopEmailEsc_(order.paymentFormUrl || CONFIG.PAYMENT_FORM_BASE_URL)
      + '" style="display:inline-block;background:#1f4617;color:#fff;text-decoration:none;padding:12px 20px;border-radius:9px;font-weight:bold">Complete payment form</a></p>'
      + '<p style="font-size:13px;color:#6b7280">In the payment form, enter KE No <strong>'
      + _shopEmailEsc_(order.keNo) + '</strong> in <strong>Registration No</strong>, use your registered phone, '
      + 'and choose <strong>Shopping Kit / Equipment</strong> for Payment For.</p>'
    : '<div style="padding:14px;background:#f0fdf4;color:#166534;border-radius:10px">This order is covered by your existing shopping payment credit. No additional payment is required.</div>';
  var html = '<!doctype html><html><body style="margin:0;background:#f4f7f2;font-family:Arial,sans-serif;color:#1b2118">'
    + '<div style="max-width:620px;margin:24px auto;background:#fff;border-radius:16px;overflow:hidden;border:1px solid #dce7d7">'
    + '<div style="padding:22px;background:#14330f;color:#fff"><h2 style="margin:0">Order placed</h2>'
    + '<div style="margin-top:5px;color:#d5e8cf">' + _shopEmailEsc_(order.orderId) + '</div></div>'
    + '<div style="padding:22px"><p>Dear ' + _shopEmailEsc_(order.riderName) + ',</p>'
    + '<p>Your riding essentials order has been recorded. Complete payment using the exact amount below.</p>'
    + '<table style="width:100%;border-collapse:collapse;margin:14px 0">' + _shopEmailItems_(order.items) + '</table>'
    + '<div style="font-size:15px;font-weight:bold;text-align:right;margin-bottom:10px">Order total: ₹'
    + Number(order.total || 0).toLocaleString('en-IN') + '</div>'
    + paymentBlock
    + '<p style="font-size:12px;color:#6b7280">Unknown sizes can be coordinated in the WhatsApp group.</p>'
    + '<div style="margin-top:22px;padding-top:16px;border-top:1px solid #e5e7eb;font-size:12px;color:#6b7280">'
    + emailFooterHtml_() + '</div></div></div></body></html>';
  try {
    var sent = sendMailKE_(to, subject, html, { textBody: 'Order ' + order.orderId + ' total ₹' + order.total });
    logEmail('Shop Order', to, '', subject, order.keNo, 'Sent', 'via ' + sent.provider);
    return true;
  } catch (e) {
    logEmailFailed('Shop Order', to, '', subject, order.keNo, String(e));
    return false;
  }
}

function sendShopOrderStatusEmail_(order, changeType) {
  var to = String(order && order.email || '').trim();
  if (!to) return false;
  var title = changeType === 'payment' ? 'Payment verified'
    : (changeType === 'delivered' ? 'Your order has been delivered' : 'Order update');
  var subject = title + ' — ' + order.orderId;
  var detail = '<p><strong>Status:</strong> ' + _shopEmailEsc_(order.orderStatus) + '</p>';
  if (order.deliveryNotes) detail += '<p><strong>Team note:</strong> ' + _shopEmailEsc_(order.deliveryNotes) + '</p>';
  if (changeType === 'delivered') {
    detail += '<div style="padding:14px;background:#eff6ff;border-radius:10px;color:#1e3a8a">'
      + 'The trainer has marked all items in this order as delivered. Delivery details are available in My Rides.</div>';
  }
  var html = '<!doctype html><html><body style="margin:0;background:#f4f7f2;font-family:Arial,sans-serif;color:#1b2118">'
    + '<div style="max-width:620px;margin:24px auto;background:#fff;border-radius:16px;border:1px solid #dce7d7;padding:24px">'
    + '<h2 style="color:#14330f;margin-top:0">' + _shopEmailEsc_(title) + '</h2>'
    + '<p>Dear ' + _shopEmailEsc_(order.riderName) + ',</p><p>Order <strong>'
    + _shopEmailEsc_(order.orderId) + '</strong> has been updated.</p>' + detail
    + '<p><a href="' + _shopEmailEsc_(CONFIG.MY_RIDES_PORTAL_URL)
    + '" style="display:inline-block;background:#1f4617;color:#fff;text-decoration:none;padding:11px 18px;border-radius:9px">Open My Rides</a></p>'
    + '<div style="margin-top:22px;padding-top:16px;border-top:1px solid #e5e7eb;font-size:12px;color:#6b7280">'
    + emailFooterHtml_() + '</div></div></body></html>';
  try {
    var sent = sendMailKE_(to, subject, html, {});
    logEmail('Shop Status', to, '', subject, order.keNo, 'Sent', 'via ' + sent.provider);
    return true;
  } catch (e) {
    logEmailFailed('Shop Status', to, '', subject, order.keNo, String(e));
    return false;
  }
}

// ────────────────────────────────────────────────────────────
//  EMAIL LOG
// ────────────────────────────────────────────────────────────

function _getOrCreateEmailLogSheet() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.EMAIL_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.EMAIL_LOG);
    const headers = ['Timestamp', 'Type', 'To', 'CC', 'Subject', 'KE No', 'Status', 'Error', 'Retry Count', 'Last Retry At'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1f4e3d').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function logEmail(type, to, cc, subject, keNo, status, errorMsg) {
  try {
    _getOrCreateEmailLogSheet().appendRow([
      new Date(), type || '', to || '', cc || '', subject || '',
      keNo || '', status || 'Sent', errorMsg || '', 0, ''
    ]);
  } catch (e) { Logger.log('logEmail error: ' + e); }
}

function logEmailFailed(type, to, cc, subject, keNo, errorMsg) {
  logEmail(type, to, cc, subject, keNo, 'Failed', errorMsg);
}

// ────────────────────────────────────────────────────────────
//  RETRY FAILED EMAILS
// ────────────────────────────────────────────────────────────

function retryFailedEmails() {
  try {
    const sheet      = _getOrCreateEmailLogSheet();
    const data       = sheet.getDataRange().getValues();
    const MAX_RETRIES = 3;
    let retried = 0, skipped = 0;

    for (let i = 1; i < data.length; i++) {
      const status     = data[i][6];
      const retryCount = Number(data[i][8]) || 0;
      const type       = data[i][1];
      const keNo       = data[i][5];
      const row        = i + 1;

      if (status !== 'Failed') continue;
      if (retryCount >= MAX_RETRIES) {
        sheet.getRange(row, 7).setValue('Abandoned').setBackground('#f8d7da').setFontColor('#721c24');
        skipped++;
        continue;
      }

      try {
        if (type === 'Welcome') {
          _retrySendWelcome(keNo);
        } else {
          sendDailyAdminSummary();
        }
        sheet.getRange(row, 7).setValue('Sent (Retried)').setBackground('#d4edda').setFontColor('#155724');
        sheet.getRange(row, 9).setValue(retryCount + 1);
        sheet.getRange(row, 10).setValue(new Date());
        retried++;
      } catch (err) {
        sheet.getRange(row, 8).setValue(String(err));
        sheet.getRange(row, 9).setValue(retryCount + 1);
        sheet.getRange(row, 10).setValue(new Date());
      }
    }

    SpreadsheetApp.getUi().alert('Retry complete.\nSent: ' + retried + '\nAbandoned (3 attempts): ' + skipped);
  } catch (e) {
    Logger.log('retryFailedEmails ERROR: ' + e);
  }
}

function _retrySendWelcome(keNo) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getRegistrationSheet_(ss);
  if (!sheet) throw new Error('Registration sheet not found');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][CONFIG.REG_COLS.KE_NO]) === String(keNo)) {
      // rebuild and resend
      const vals         = data[i];
      const studentName  = String(vals[CONFIG.REG_COLS.STUDENT]  || '').trim();
      const email        = String(vals[CONFIG.REG_COLS.EMAIL]    || '').trim();
      const phone        = String(vals[CONFIG.REG_COLS.PHONE]    || '').trim();
      const grade        = String(vals[CONFIG.REG_COLS.GRADE]    || '').trim();
      const program      = String(vals[CONFIG.REG_COLS.PROGRAM]  || 'school').trim().toLowerCase();
      const parentName   = String(vals[CONFIG.REG_COLS.PARENT]   || '').trim();
      const payFormUrl   = buildPaymentFormUrl({ regRef: keNo, student: studentName, parent: parentName, email, phone, grade });
      sendIndusSchoolWelcomeEmail({
        studentName, parentName, email, phone, grade, program,
        serviceProgram: program, keNo, payFormUrl,
        isFirstTime: false, sheet, row: i + 1
      });
      return;
    }
  }
  throw new Error('KE No not found: ' + keNo);
}

// ────────────────────────────────────────────────────────────
//  EMAIL REPORT
// ────────────────────────────────────────────────────────────

function showEmailSendReport() {
  const sheet = _getOrCreateEmailLogSheet();
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) { SpreadsheetApp.getUi().alert('No emails logged yet.'); return; }

  const counts = {}, byType = {};
  let total = 0, failed = 0, retried = 0;

  for (let i = 1; i < data.length; i++) {
    const type   = data[i][1] || 'Unknown';
    const status = data[i][6] || '';
    total++;
    counts[type]   = (counts[type]  || 0) + 1;
    byType[status] = (byType[status] || 0) + 1;
    if (status === 'Failed')         failed++;
    if (status === 'Sent (Retried)') retried++;
  }

  let msg = 'EMAIL SEND REPORT\n─────────────────────\nTotal: ' + total + '\n\nBy Type:\n';
  Object.keys(counts).forEach(function(t) { msg += '  • ' + t + ': ' + counts[t] + '\n'; });
  msg += '\nBy Status:\n';
  Object.keys(byType).forEach(function(s) { msg += '  • ' + s + ': ' + byType[s] + '\n'; });
  msg += '\nFailed (pending retry): ' + failed + '\nSuccessfully retried: ' + retried;

  SpreadsheetApp.getUi().alert(msg);
}