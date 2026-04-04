// ============================================================
// KINGS EQUESTRIAN — NEW SYSTEM
// File: 1_Config.gs
// Central config, column maps, and shared utilities
// ============================================================

// ⚠️  PAYMENT FORM COLUMN NOTE
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Payment Form Response sheet column order (0-based):
//   0  Timestamp (auto)
//   1  KE No
//   2  Phone
//   3  Amount Paid
//   4  Screenshot (file upload)   ← col 4
//   5  Payment Date               ← col 5
//   6  Transaction Ref            ← col 6
//   7  PAN/Aadhaar
//   8  Verified (script)
//   9  Receipt Sent (script)
//  10  Receipt Sent At (script)
//  11  Receipt No (script)
//  12  Drive Link (script)
//
// Payments Ledger sheet column order (0-based):
//   0  KE No
//   1  Name
//   2  Phone
//   3  Amount
//   4  Payment Date               ← ONLY payment date goes here
//   5  Transaction Ref            ← ONLY txn ref goes here
//   6  Receipt No
//   7  Sent At
//
// Run diagnosPaymentFormColumns() if data still looks wrong.
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const CONFIG = {
  UPI_ID: 'vyapar.176548151976@hdfcbank',
  BUSINESS_NAME: 'KingsEquestrian',
  PAYMENT_FORM_LINK: 'https://docs.google.com/forms/d/e/1FAIpQLSeQEpr82za7CflvDNidtCU93LVHW7NjbjCNIaGm386XGax_Qg/viewform?usp=header',
  TERMS_CONDITIONS_DOC_ID: '1QbJHA5keyTLvgw-5stTY74i92BQ89TYya-NvtJ4YGx4',
  ADVANCE_BOOKING_AMOUNT: 1000,
  WEB_APP_URL: 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec',
  // FIX #7/#9: This ID was a published-web URL, not a Drive file ID.
  // Use the actual Drive file ID from the sharing link:
  // https://drive.google.com/file/d/1CpWYOphlAJzJSHtuS9au35tWdg743rAW/view
  ADDITIONAL_PDF_DRIVE_LINK: 'https://drive.google.com/file/d/1Wnh-GR2G7DE7SO_It1OtsaOq77YDfMxK/view?usp=sharing',
  MYRIDES:'https://script.google.com/macros/s/AKfycbyzCGHcVGHQQFP-VTepIZ4ipsMjaoXAYSvowU-qahWrd45ckslE2kO1XafDNskxma0BFw/exec?app=portal',

  // ── Sheet names ─────────────────────────────────────────
  SHEETS: {
    BOOKING_FORM : 'Booking Form Response',
    PAYMENT_FORM : 'Payment Form Response',
    RIDERS       : 'Riders',
    SCHEDULE     : 'Schedule',
    PAYMENTS     : 'Payments Ledger',
    PRICING      : 'Pricing',
    MAIL_INFO    : 'Mail Info'
  },

  // ── Booking Form Response columns (0-based) ──────────────
  BOOKING_COLS: {
    TIMESTAMP   : 0,
    NAME        : 1,
    EMAIL       : 2,
    PHONE       : 3,
    SERVICES    : 4,
    PARTICIPANTS: 5,
    PREF_DATE   : 6,
    PREF_TIME   : 7,
    CONSENT     : 8,
    KE_NO       : 9,
    WELCOME_SENT: 10,
    WELCOME_AT  : 11
  },

  // ── Payment Form Response columns (0-based) ──────────────
  // FIX #11: Columns were mapped incorrectly before.
  // Screenshot is col 4, Payment Date is col 5, TxnRef is col 6.
  PAYMENT_COLS: {
    TIMESTAMP   : 0,
    KE_NO       : 1,
    PHONE       : 2,
    AMOUNT      : 3,
    SCREENSHOT  : 4,   // file upload — col 4
    PAY_DATE    : 5,   // actual payment date — col 5
    TXN_REF     : 6,   // transaction ref — col 6
    PAN         : 7,
    VERIFIED    : 8,
    RECEIPT_SENT: 9,
    RECEIPT_AT  : 10,
    RECEIPT_NO  : 11,
    DRIVE_LINK  : 12
  },

  // ── Riders sheet columns (0-based) ───────────────────────
  RIDER_COLS: {
    KE_NO       : 0,
    NAME        : 1,
    EMAIL       : 2,
    PHONE       : 3,
    SERVICES    : 4,
    PARTICIPANTS: 5,
    REGISTERED  : 6,
    NOTES       : 7
  },

  // ── Schedule sheet columns (0-based) ─────────────────────
  SCHED_COLS: {
    KE_NO       : 0,
    NAME        : 1,
    PHONE       : 2,
    EMAIL       : 3,
    SERVICE     : 4,
    DATE        : 5,
    TIME_SLOT   : 6,
    PARTICIPANTS: 7,
    STATUS      : 8,
    ATTENDANCE  : 9,
    STAFF_NOTES : 10,
    CAL_EVENT_ID: 11,
    SOURCE      : 12
  },

  // ── Payments Ledger columns (0-based) ────────────────────
  // FIX #11: PAY_DATE = col 4, TXN_REF = col 5 (never overlap)
  LEDGER_COLS: {
    KE_NO      : 0,
    NAME       : 1,
    PHONE      : 2,
    AMOUNT     : 3,
    SCREENSHOT :4,
    PAY_DATE   : 5,
    SCHEDULE_DATE: 6,   // payment date ONLY
    TXN_REF    :7,  // transaction ref ONLY
    RECEIPT_NO : 8,
    SENT_AT    : 9
  },

  // ── Pricing sheet columns (0-based) ──────────────────────
  PRICING_COLS: {
    ROW      : 0,
    NAME     : 1,
    PRICE    : 2,
    DOC_ID   : 3,
    TYPE     : 4
  }
};

// ============================================================
//  KE NUMBER GENERATOR
// ============================================================

function generateKENo() {
   const date = new Date();
    const year = date.getFullYear().toString().substr(-2);
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const random = Math.floor(Math.random() * 9000) + 1000;
    return `KE${year}${month}${day}${random}`;
}

// ============================================================
//  PHONE NORMALISE (last 10 digits)
// ============================================================

function normalisePhone(p) {
  return String(p || '').replace(/\D/g, '').slice(-10);
}

// ============================================================
//  RIDER LOOKUPS
// ============================================================

function findRiderByPhone(phone) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
  if (!sheet) return null;
  const target = normalisePhone(phone);
  if (!target || target.length < 10) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (normalisePhone(data[i][CONFIG.RIDER_COLS.PHONE]) === target) {
      return { rowIndex: i + 1, row: data[i] };
    }
  }
  return null;
}
function findRiderByPhoneAndName(phone, name) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
  if (!sheet) return null;
  const target     = normalisePhone(phone);
  const targetName = String(name || '').trim().toLowerCase();
  if (!target || target.length < 10) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowPhone = normalisePhone(data[i][CONFIG.RIDER_COLS.PHONE]);
    const rowName  = String(data[i][CONFIG.RIDER_COLS.NAME] || '').trim().toLowerCase();
    if (rowPhone === target && rowName === targetName) {
      return { rowIndex: i + 1, row: data[i] };
    }
  }
  return null;
}

function findRiderByKENo(keNo) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
  if (!sheet) return null;
  const target = String(keNo || '').trim().toUpperCase();
  if (!target) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][CONFIG.RIDER_COLS.KE_NO] || '').trim().toUpperCase() === target) {
      return { rowIndex: i + 1, row: data[i] };
    }
  }
  return null;
}

// ============================================================
//  DATE HELPERS
// ============================================================

function fmtDate(d) {
  if (!d) return '';
  try {
    const dt = (d instanceof Date) ? d : new Date(d);
    if (isNaN(dt.getTime())) return '';
    // FIX #4: guard against epoch (1 Jan 1970) which means the field was empty
    if (dt.getFullYear() < 2000) return '';
    return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd MMM yyyy');
  }
  catch (e) { return String(d); }
}

function fmtDateTime(d) {
  if (!d) return '';
  try {
    const dt = (d instanceof Date) ? d : new Date(d);
    if (isNaN(dt.getTime())) return '';
    if (dt.getFullYear() < 2000) return '';
    return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
  }
  catch (e) { return String(d); }
}

function ymd(d) {
  if (!d) return '';
  try {
    const dt = (d instanceof Date) ? d : new Date(d);
    if (isNaN(dt.getTime())) return '';
    if (dt.getFullYear() < 2000) return '';
    return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  catch (e) { return ''; }
}

// ============================================================
//  UPI / QR
// ============================================================

function createUPILink(amount, reference) {
  return 'upi://pay?pa=' + CONFIG.UPI_ID
    + '&pn=' + encodeURIComponent(CONFIG.BUSINESS_NAME)
    + '&am=' + amount
    + '&cu=INR'
    + '&tn=' + encodeURIComponent(reference);
}

function createQRCode(link) {
  return 'https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=' + encodeURIComponent(link);
}

// ============================================================
//  PRICING DATA
// ============================================================

function getPricingData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PRICING);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const map  = {};
  for (let i = 1; i < data.length; i++) {
    const name  = String(data[i][CONFIG.PRICING_COLS.NAME]   || '').trim();
    const price = data[i][CONFIG.PRICING_COLS.PRICE];
    const docId = String(data[i][CONFIG.PRICING_COLS.DOC_ID] || '').trim();
    const type  = String(data[i][CONFIG.PRICING_COLS.TYPE]   || 'Regular').trim();
    if (name) map[name] = { price, docId, type };
  }
  return map;
}

function getServicesList() {
  const pricing = getPricingData();
  return Object.keys(pricing).map(k => ({
    name : k,
    price: pricing[k].price,
    type : pricing[k].type
  }));
}

// ============================================================
//  CC RECIPIENTS
// ============================================================

function getCCRecipients(mailType) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MAIL_INFO);
    if (!sheet) return [];
    const data  = sheet.getDataRange().getValues();
    const cc    = [];
    for (let i = 1; i < data.length; i++) {
      const email = String(data[i][0] || '').trim();
      const type  = String(data[i][1] || '').toLowerCase();
      if (email && type.includes(mailType.toLowerCase())) cc.push(email);
    }
    return cc;
  } catch (e) {
    Logger.log('getCCRecipients error: ' + e);
    return [];
  }
}

function getAdminEmails() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MAIL_INFO);
    if (!sheet) return [];
    const data  = sheet.getDataRange().getValues();
    const out   = [];
    for (let i = 1; i < data.length; i++) {
      const email = String(data[i][0] || '').trim();
      const type  = String(data[i][1] || '').toLowerCase();
      if (email && (type.includes('admin') || type.includes('daily'))) out.push(email);
    }
    return out;
  } catch (e) { return []; }
}

// ============================================================
//  NUMBER → WORDS
// ============================================================

function numberToWords(num) {
  const ones  = ['','One','Two','Three','Four','Five','Six','Seven','Eight','Nine'];
  const teens = ['Ten','Eleven','Twelve','Thirteen','Fourteen','Fifteen','Sixteen','Seventeen','Eighteen','Nineteen'];
  const tens  = ['','','Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety'];
  function convert(n) {
    if (n === 0) return 'Zero';
    if (n < 10)  return ones[n];
    if (n < 20)  return teens[n - 10];
    if (n < 100) return tens[Math.floor(n/10)] + (n%10 ? ' '+ones[n%10] : '');
    if (n < 1000) return ones[Math.floor(n/100)] + ' Hundred' + (n%100 ? ' '+convert(n%100) : '');
    if (n < 100000) return convert(Math.floor(n/1000)) + ' Thousand' + (n%1000 ? ' '+convert(n%1000) : '');
    if (n < 10000000) return convert(Math.floor(n/100000)) + ' Lakh' + (n%100000 ? ' '+convert(n%100000) : '');
    return convert(Math.floor(n/10000000)) + ' Crore' + (n%10000000 ? ' '+convert(n%10000000) : '');
  }
  return convert(num).trim() + ' Rupees Only';
}

// ============================================================
//  DRIVE — RECEIPTS FOLDER
// ============================================================

function getReceiptsFolder() {
  const year = new Date().getFullYear().toString();
  let main   = DriveApp.getFoldersByName('Kings Equestrian Receipts');
  main       = main.hasNext() ? main.next() : DriveApp.createFolder('Kings Equestrian Receipts');
  let yf     = main.getFoldersByName(year);
  return yf.hasNext() ? yf.next() : main.createFolder(year);
}

function storeReceiptInDrive(blob, riderName, receiptNo) {
  try {
    const folder   = getReceiptsFolder();
    const ts       = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const fileName = 'Receipt_' + receiptNo.replace(/\//g,'-') + '_' + riderName.replace(/\s+/g,'_') + '_' + ts + '.pdf';
    const file     = folder.createFile(blob);
    file.setName(fileName);
    return { fileId: file.getId(), fileUrl: file.getUrl() };
  } catch (e) {
    Logger.log('storeReceiptInDrive error: ' + e);
    return null;
  }
}

// ============================================================
//  IMAGE HELPERS
// ============================================================

function imgBase64FromDrive(fileId) {
  try {
    const b = DriveApp.getFileById(fileId).getBlob();
    return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
  } catch (e) { return ''; }
}

function imgBase64FromUrl(url) {
  try {
    const b = UrlFetchApp.fetch(url).getBlob();
    return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
  } catch (e) { return ''; }
}

// ============================================================
//  TERMS PDF
// ============================================================

function getTermsPDF() {
  try {
    const doc  = DocumentApp.openById(CONFIG.TERMS_CONDITIONS_DOC_ID);
    const blob = doc.getAs('application/pdf');
    blob.setName('Terms_and_Conditions.pdf');
    return blob;
  } catch (e) {
    Logger.log('getTermsPDF error: ' + e);
    return null;
  }
}

function getServicePDF(docId, serviceName) {
  try {
    const doc  = DocumentApp.openById(docId);
    const blob = doc.getAs('application/pdf');
    blob.setName(serviceName.replace(/\s+/g,'_') + '_Details.pdf');
    return blob;
  } catch (e) { return null; }
}

// ============================================================
//  CONSENT PDF GENERATOR
// ============================================================

function generateConsentPDF(name, email, phone, bookingDate) {

    const LABEL_FONT = 'Arial';
    const FONT_SIZE = 11;

    const doc = DocumentApp.create('Consent Form - ' + (name || 'Participant'));
    const body = doc.getBody();
    body.clear();

    body.setMarginTop(40);
    body.setMarginBottom(40);
    body.setMarginLeft(50);
    body.setMarginRight(50);

    function paragraph(textStr, size, bold, spacing, align) {
        size    = size    || FONT_SIZE;
        bold    = bold    || false;
        spacing = spacing || 6;
        const p = body.appendParagraph(textStr);
        const t = p.editAsText();
        t.setFontFamily(LABEL_FONT).setFontSize(size).setBold(bold);
        if (align) p.setAlignment(align);
        p.setSpacingAfter(spacing);
        return p;
    }

    function formatValue(textObj, fullText, value) {
        if (!value || value.toString().trim().length === 0) return;
        if (!fullText) return;
        const valStr = value.toString();
        const start = fullText.indexOf(valStr);
        if (start === -1) return;
        const end = start + valStr.length - 1;
        if (end >= start && start >= 0 && end < fullText.length) {
            textObj.setBold(start, end, true);
            textObj.setUnderline(start, end, true);
        }
    }

    function formatDateOnly(dateValue) {
        if (!dateValue) return null;
        if (typeof dateValue === 'string') {
            try {
                const parsed = new Date(dateValue);
                if (!isNaN(parsed.getTime())) { dateValue = parsed; } else { return dateValue; }
            } catch (e) { return dateValue; }
        }
        if (dateValue instanceof Date) {
            const day = String(dateValue.getDate()).padStart(2, '0');
            const month = String(dateValue.getMonth() + 1).padStart(2, '0');
            const year = dateValue.getFullYear();
            return day + '/' + month + '/' + year;
        }
        return dateValue.toString();
    }

    const logoUrl = 'https://drive.google.com/uc?export=view&id=1EAkJ8_EeOVmpX3L1RGLi8b9amX5wuLhb';
    let logoBlob;
    try {
        logoBlob = UrlFetchApp.fetch(logoUrl).getBlob();
    } catch (e) {
        Logger.log('Error fetching logo: ' + e);
    }

    if (logoBlob) {
        const logoPara = body.appendParagraph('');
        logoPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        const logoImg = logoPara.appendInlineImage(logoBlob);
        logoImg.setWidth(120);
        logoImg.setHeight(120);
        logoPara.setSpacingAfter(20);
    }

    paragraph('KINGS EQUESTRIAN FOUNDATION', 16, true, 5, DocumentApp.HorizontalAlignment.CENTER);
    paragraph('Acknowledgement & Consent Form - Horse Riding Participants', 13, true, 3, DocumentApp.HorizontalAlignment.CENTER);
    paragraph('(Applicable for Individual / Group / Family Participants)', 10, false, 25, DocumentApp.HorizontalAlignment.CENTER);

    paragraph('Kings Equestrian Foundation offers horse riding programs and related activities, which may include casual riding, dressage, jumping, workshops, clinics, and equine interaction.', 11, false, 12);
    paragraph('I/we understand and acknowledge that participation in equestrian activities involves inherent risks, including but not limited to falls, bruises, muscle strain, fractures, head injuries, or other serious injuries. I/we further acknowledge that horses are live animals and their behaviour can be unpredictable.', 11, false, 12);
    paragraph('I/we also acknowledge that Kings Equestrian Foundation follows reasonable safety precautions, provides trained supervision, and enforces established safety guidelines. However, despite all precautions, accidents may occasionally occur.', 11, false, 20);

    let sepPara = body.appendParagraph('---');
    sepPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    sepPara.setSpacingAfter(20);

    paragraph('Medical Fitness & Insurance Declaration', 12, true, 12);
    paragraph('I/we hereby declare that I / my child / all participants covered under this consent are medically fit to participate in horse riding and equestrian-related activities. To the best of my/our knowledge, there are no undisclosed medical conditions, injuries, or health concerns that would prevent safe participation, except those disclosed in writing to Kings Equestrian Foundation prior to participation.', 11, false, 12);
    paragraph('I/we further confirm that I / my child / all participants are covered by valid medical and/or personal accident insurance, which will cover any injuries, medical treatment, or emergencies arising from participation.', 11, false, 12);
    paragraph('I/we understand and agree that Kings Equestrian Foundation is not responsible for medical expenses, and all such costs shall be borne by the participant(s) or covered under their insurance.', 11, false, 20);

    sepPara = body.appendParagraph('---');
    sepPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    sepPara.setSpacingAfter(20);

    paragraph('Acknowledgement & Agreement', 12, true, 12);
    paragraph('I/we confirm that:', 11, false, 8);

    const bulletPoints = [
        'I/we have carefully read and fully understood this consent form.',
        'I/we understand the nature of equestrian activities and the associated risks.',
        'I/we voluntarily consent to participation.',
        'For participants under 18 years of age, I/we am/are the parent(s) or legal guardian(s) and authorised to provide consent.',
        'All participants agree to follow safety instructions, rules, and guidelines issued by Kings Equestrian Foundation and its instructors at all times.'
    ];

    bulletPoints.forEach(function(point) {
        const bp = body.appendParagraph('- ' + point);
        bp.editAsText().setFontFamily(LABEL_FONT).setFontSize(11);
        bp.setSpacingAfter(6);
        bp.setIndentStart(20);
    });

    body.appendParagraph('').setSpacingAfter(8);
    paragraph('I/we agree that Kings Equestrian Foundation, its trainers, staff, and associates shall not be held responsible for injuries arising from participation, except in cases of proven negligence.', 11, false, 20);

    sepPara = body.appendParagraph('---');
    sepPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    sepPara.setSpacingAfter(20);

    paragraph('Primary Contact / Parent / Guardian Details', 12, true, 12);

    var p = body.appendParagraph('');
    var t = p.editAsText();
    const nameSpaced = name ? ('  ' + name + '  ') : '___________________________________';
    const nameLine = 'Name: ' + nameSpaced;
    t.setText(nameLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (name) formatValue(t, nameLine, nameSpaced);
    p.setSpacingAfter(12);

    p = body.appendParagraph('');
    t = p.editAsText();
    const phoneSpaced = phone ? ('  ' + phone + '  ') : '___________________________________';
    const phoneLine = 'Contact Number: ' + phoneSpaced;
    t.setText(phoneLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (phone) formatValue(t, phoneLine, phoneSpaced);
    p.setSpacingAfter(12);

    p = body.appendParagraph('');
    t = p.editAsText();
    const emailSpaced = email ? ('  ' + email + '  ') : '___________________________________';
    const emailLine = 'Email ID: ' + emailSpaced;
    t.setText(emailLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (email) formatValue(t, emailLine, emailSpaced);
    p.setSpacingAfter(25);

    sepPara = body.appendParagraph('---');
    sepPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    sepPara.setSpacingAfter(25);

    p = body.appendParagraph('');
    t = p.editAsText();
    const signatureSpaced = name ? ('  ' + name + '  ') : '___________________________________';
    const dateFormatted = formatDateOnly(bookingDate);
    const dateSpaced = dateFormatted ? ('  ' + dateFormatted + '  ') : '_______________';
    const signatureLine = 'Signature of Participant / Parent / Guardian: ' + signatureSpaced + '     Date: ' + dateSpaced;
    t.setText(signatureLine).setFontFamily(LABEL_FONT).setFontSize(11);

    if (name) {
        const sigStart = signatureLine.indexOf(signatureSpaced);
        if (sigStart !== -1) {
            const sigEnd = sigStart + signatureSpaced.length - 1;
            if (sigEnd >= sigStart && sigStart >= 0) {
                t.setFontFamily(sigStart, sigEnd, 'Dancing Script');
                t.setFontSize(sigStart, sigEnd, 16);
                t.setBold(sigStart, sigEnd, false);
                t.setUnderline(sigStart, sigEnd, false);
            }
        }
    }

    if (dateFormatted) formatValue(t, signatureLine, dateSpaced);
    p.setSpacingAfter(30);

    const footerPara = body.appendParagraph('Kings Equestrian Foundation | Karnataka, India | +91-9980895533 | info@kingsequestrian.com');
    footerPara.editAsText().setFontFamily(LABEL_FONT).setFontSize(9).setForegroundColor('#666666');
    footerPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    doc.saveAndClose();

    const pdf = doc.getAs('application/pdf');
    pdf.setName('Consent_Form_' + (name || 'Participant').replace(/\s+/g, '_') + '.pdf');

    DriveApp.getFileById(doc.getId()).setTrashed(true);

    return pdf;
}

// ============================================================
//  DIAGNOSTIC
// ============================================================

function diagnosPaymentFormColumns() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENT_FORM);
  if (!sheet) { Logger.log('Payment Form Response sheet not found'); return; }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  headers.forEach(function(h, i) { Logger.log('Col ' + i + ' (' + String.fromCharCode(65+i) + '): ' + h); });
}

// ============================================================
//  MENU & TRIGGERS
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Kings Equestrian')
    .addItem('Resend Welcome Email',    'resendWelcomeEmail')
    .addItem('Send Payment Receipt',    'sendPaymentReceiptMenu')
    .addSeparator()
    .addItem('Send Daily Summary Now',  'testSendDailySummaryNow')
    .addItem('Dry-Run Daily Summary',   'testDailySummaryDryRun')
    .addSeparator()
    .addItem('Setup All Triggers',      'setupTriggers')
    .addItem('Diagnose Payment Columns','diagnosPaymentFormColumns')
    .addToUi();
}

function setupTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onBookingFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('onPaymentFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('sendDailyAdminSummary').timeBased().everyDays(1).atHour(20).create();
  SpreadsheetApp.getUi().alert('Triggers set!\n\n- Booking form: welcome email + KE No\n- Payment form: receipt email\n- Nightly 9 PM: admin summary email');
}