// ============================================================
// KINGS EQUESTRIAN — NEW SYSTEM
// File: 1_Config.gs
// Central config, column maps, and shared utilities
// ============================================================

// ⚠️  TROUBLESHOOTING: PAYMENT FORM COLUMN MISMATCH
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// If payment data is uploaded incorrectly (e.g., screenshot URLs 
// appearing in Payment Date field), your Google Form field order 
// does NOT match the CONFIG.PAYMENT_COLS mapping below.
//
// FIX: Open your Payment Form and verify fields are in this EXACT order:
//   1. Timestamp (auto)
//   2. KE No (text)
//   3. Phone (text)
//   4. Amount Paid (number)
//   5. Payment Date (date)
//   6. Transaction Ref (text)
//   7. PAN/Aadhaar (text)
//   8. Screenshot (file upload)
//   
// If you need to reorder: Delete the sheet "Payment Form Response", 
// then reshare the form or delete & recreate it with correct order.
// 
// RUN THIS TO DIAGNOSE: diagnosPaymentFormColumns()
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const CONFIG = {
  UPI_ID: 'vyapar.176548151976@hdfcbank',
  BUSINESS_NAME: 'KingsEquestrian',
  PAYMENT_FORM_LINK: 'https://docs.google.com/forms/d/e/1FAIpQLSeQEpr82za7CflvDNidtCU93LVHW7NjbjCNIaGm386XGax_Qg/viewform?usp=header',
  TERMS_CONDITIONS_DOC_ID: '1QbJHA5keyTLvgw-5stTY74i92BQ89TYya-NvtJ4YGx4',
  ADVANCE_BOOKING_AMOUNT: 1000,
  WEB_APP_URL: 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec',
  ADDITIONAL_PDF_DOC_ID:'2PACX-1vReen5Pof84-7XoZdzAmbh6JnSW0vJ_LW_C2wfVfZnUl3PzbglIbJtcBEqEoxlJyw',

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
  // A Timestamp | B Name | C Email ID | D Phone Number
  // E Our Services | F Number of Participants
  // G Preferred Service Date | H Preferred Time Slot | I Consent
  // J KE No (script) | K Welcome Sent (script) | L Welcome Sent At (script)
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
  // A Timestamp | B KE No | C Phone | D Amount Paid
  // E Payment Date | F Transaction Ref | G PAN/Aadhaar
  // H Screenshot | I Verified (script) | J Receipt Sent (script)
  // K Receipt Sent At (script) | L Receipt No (script) | M Drive Link (script)
  PAYMENT_COLS: {
    TIMESTAMP   : 0,
    KE_NO       : 1,
    PHONE       : 2,
    AMOUNT      : 3,
    SCREENSHOT: 4,
    PAY_DATE    : 5,
    TXN_REF     : 6,
    PAN         : 7,
    VERIFIED    : 8,
    RECEIPT_SENT: 9,
    RECEIPT_AT  : 10,
    RECEIPT_NO  : 11,
    DRIVE_LINK  : 12
  },

  // ── Riders sheet columns (0-based) ───────────────────────
  // A KE No | B Name | C Email | D Phone | E Services
  // F Participants | G Registered On | H Notes
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
  // A KE No | B Name | C Phone | D Email | E Service
  // F Session Date | G Time Slot | H Participants
  // I Status | J Attendance | K Staff Notes
  // L Calendar Event ID | M Booking Source
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
  // A KE No | B Name | C Phone | D Amount | E Payment Date
  // F Transaction Ref | G Receipt No | H Sent At
  LEDGER_COLS: {
    KE_NO      : 0,
    NAME       : 1,
    PHONE      : 2,
    AMOUNT     : 3,
    PAY_DATE   : 4,
    TXN_REF    : 5,
    RECEIPT_NO : 6,
    SENT_AT    : 7
  },

  // ── Pricing sheet columns (0-based) ──────────────────────
  // A Row | B Service Name | C Price | D Google Doc ID | E Type
  // Type values: 'One-Time' | 'Regular'
  PRICING_COLS: {
    ROW      : 0,
    NAME     : 1,
    PRICE    : 2,
    DOC_ID   : 3,
    TYPE     : 4   // 'One-Time' or 'Regular'
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

function quickTest() {
  const testId = 'd/1CpWYOphlAJzJSHtuS9au35tWdg743rAW'; // Replace with your ID
  Logger.log('Testing document access...');
  const result = testDocumentAccess(testId);
  if (result) {
    Logger.log('✅ Document is accessible - ready for welcome emails!');
  } else {
    Logger.log('❌ Document access failed - check permissions and ID');
  }
}
function testDocumentAccess(docId) {
  try {
    Logger.log('Testing access to document/presentation ID: ' + docId);

    let doc, blob, fileType;

    // Try as Google Doc first
    try {
      doc = DocumentApp.openById(docId);
      Logger.log('✅ SUCCESS: Google Doc "' + doc.getName() + '" is accessible');
      fileType = 'Google Doc';
      blob = doc.getAs('application/pdf');
    } catch (docError) {
      // Try as Google Slides presentation
      try {
        doc = SlidesApp.openById(docId);
        Logger.log('✅ SUCCESS: Google Slides "' + doc.getName() + '" is accessible');
        fileType = 'Google Slides';
        blob = doc.getAs('application/pdf');
      } catch (slidesError) {
        throw new Error('Neither Google Doc nor Slides accessible: ' + docError.message + ' | ' + slidesError.message);
      }
    }

    Logger.log('Document URL: https://docs.google.com/' + (fileType === 'Google Doc' ? 'document' : 'presentation') + '/d/' + docId + '/edit');
    Logger.log('✅ SUCCESS: PDF conversion works (' + blob.getBytes().length + ' bytes)');

    return true;
  } catch (e) {
    Logger.log('❌ FAILED: ' + e.message);
    Logger.log('💡 Possible solutions:');
    Logger.log('  1. Check if document/presentation ID is correct');
    Logger.log('  2. For Google Docs: Use ID from https://docs.google.com/document/d/YOUR_ID/edit');
    Logger.log('  3. For Google Slides: Use ID from https://docs.google.com/presentation/d/YOUR_ID/edit');
    Logger.log('  4. Ensure sharing is set to "Anyone with the link can view"');
    Logger.log('  5. Try publishing the document (File > Publish to web)');
    return false;
  }
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
  try { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'dd MMM yyyy'); }
  catch (e) { return String(d); }
}

function fmtDateTime(d) {
  if (!d) return '';
  try { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm'); }
  catch (e) { return String(d); }
}

function ymd(d) {
  if (!d) return '';
  try { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
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
//  PRICING DATA  (includes Type column)
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
    const type  = String(data[i][CONFIG.PRICING_COLS.TYPE]   || 'Regular').trim(); // 'One-Time' | 'Regular'
    if (name) map[name] = { price, docId, type };
  }
  return map;
}

// Returns flat array of services for portal dropdown
// [{name, price, type}]
function getServicesList() {
  const pricing = getPricingData();
  return Object.keys(pricing).map(k => ({
    name : k,
    price: pricing[k].price,
    type : pricing[k].type
  }));
}

function testEmailPermission() {
  try {
    GmailApp.sendEmail('jyothikondupally@gmail.com', 'Test', 'Test message');
    Logger.log('✅ Email permission granted!');
  } catch (e) {
    Logger.log('❌ Still no permission: ' + e);
  }
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

    function paragraph(textStr, size = FONT_SIZE, bold = false, spacing = 6, align = null) {
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
            return `${day}/${month}/${year}`;
        }
        return dateValue.toString();
    }

    const logoUrl = 'https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg';
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
    paragraph('Acknowledgement & Consent Form – Horse Riding Participants', 13, true, 3, DocumentApp.HorizontalAlignment.CENTER);
    paragraph('(Applicable for Individual / Group / Family Participants)', 10, false, 25, DocumentApp.HorizontalAlignment.CENTER);

    paragraph('Kings Equestrian Foundation offers horse riding programs and related activities, which may include casual riding, dressage, jumping, workshops, clinics, and equine interaction.', 11, false, 12);
    paragraph('I/we understand and acknowledge that participation in equestrian activities involves inherent risks, including but not limited to falls, bruises, muscle strain, fractures, head injuries, or other serious injuries. I/we further acknowledge that horses are live animals and their behaviour can be unpredictable.', 11, false, 12);
    paragraph('I/we also acknowledge that Kings Equestrian Foundation follows reasonable safety precautions, provides trained supervision, and enforces established safety guidelines. However, despite all precautions, accidents may occasionally occur.', 11, false, 20);

    let sepPara = body.appendParagraph('⸻');
    sepPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    sepPara.setSpacingAfter(20);

    paragraph('Medical Fitness & Insurance Declaration', 12, true, 12);
    paragraph('I/we hereby declare that I / my child / all participants covered under this consent are medically fit to participate in horse riding and equestrian-related activities. To the best of my/our knowledge, there are no undisclosed medical conditions, injuries, or health concerns that would prevent safe participation, except those disclosed in writing to Kings Equestrian Foundation prior to participation.', 11, false, 12);
    paragraph('I/we further confirm that I / my child / all participants are covered by valid medical and/or personal accident insurance, which will cover any injuries, medical treatment, or emergencies arising from participation.', 11, false, 12);
    paragraph('I/we understand and agree that Kings Equestrian Foundation is not responsible for medical expenses, and all such costs shall be borne by the participant(s) or covered under their insurance.', 11, false, 20);

    sepPara = body.appendParagraph('⸻');
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

    bulletPoints.forEach(point => {
        const p = body.appendParagraph('• ' + point);
        p.editAsText().setFontFamily(LABEL_FONT).setFontSize(11);
        p.setSpacingAfter(6);
        p.setIndentStart(20);
        p.setIndentFirstLine(0);
    });

    body.appendParagraph('').setSpacingAfter(8);
    paragraph('I/we agree that Kings Equestrian Foundation, its trainers, staff, and associates shall not be held responsible for injuries arising from participation, except in cases of proven negligence.', 11, false, 20);

    sepPara = body.appendParagraph('⸻');
    sepPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    sepPara.setSpacingAfter(20);

    paragraph('Primary Contact / Parent / Guardian Details', 12, true, 12);

    let p = body.appendParagraph('');
    let t = p.editAsText();
    const nameSpaced = name ? `  ${name}  ` : '___________________________________';
    const nameLine = `Name: ${nameSpaced}`;
    t.setText(nameLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (name) formatValue(t, nameLine, nameSpaced);
    p.setSpacingAfter(12);

    p = body.appendParagraph('');
    t = p.editAsText();
    const phoneSpaced = phone ? `  ${phone}  ` : '___________________________________';
    const phoneLine = `Contact Number: ${phoneSpaced}`;
    t.setText(phoneLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (phone) formatValue(t, phoneLine, phoneSpaced);
    p.setSpacingAfter(12);

    p = body.appendParagraph('');
    t = p.editAsText();
    const emailSpaced = email ? `  ${email}  ` : '___________________________________';
    const emailLine = `Email ID: ${emailSpaced}`;
    t.setText(emailLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
    if (email) formatValue(t, emailLine, emailSpaced);
    p.setSpacingAfter(25);

    sepPara = body.appendParagraph('⸻');
    sepPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    sepPara.setSpacingAfter(25);

    p = body.appendParagraph('');
    t = p.editAsText();
    const signatureSpaced = name ? `  ${name}  ` : '___________________________________';
    const dateFormatted = formatDateOnly(bookingDate);
    const dateSpaced = dateFormatted ? `  ${dateFormatted}  ` : '_______________';
    const signatureLine = `Signature of Participant / Parent / Guardian: ${signatureSpaced}     Date: ${dateSpaced}`;
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

    const footerPara = paragraph('Kings Equestrian Foundation | Karnataka, India | +91-9980895533 | info@kingsequestrian.com', 9, false, 0, DocumentApp.HorizontalAlignment.CENTER);
    footerPara.editAsText().setForegroundColor('#666666');

    doc.saveAndClose();

    const pdf = doc.getAs('application/pdf');
    pdf.setName(`Consent_Form_${(name || 'Participant').replace(/\s+/g, '_')}.pdf`);

    DriveApp.getFileById(doc.getId()).setTrashed(true);

    return pdf;
}

function getAdditionalPDF() {
  try {
    // Skip if no DOC ID configured
    if (!CONFIG.ADDITIONAL_PDF_DOC_ID || CONFIG.ADDITIONAL_PDF_DOC_ID === 'YOUR_ADDITIONAL_PDF_DOC_ID_HERE') {
      Logger.log('Additional PDF not configured - skipping');
      return null;
    }

    Logger.log('Attempting to access document/presentation with ID: ' + CONFIG.ADDITIONAL_PDF_DOC_ID);

    let doc, blob;

    // Try as Google Doc first
    try {
      doc = DocumentApp.openById(CONFIG.ADDITIONAL_PDF_DOC_ID);
      Logger.log('✅ Document accessed successfully: ' + doc.getName());
      blob = doc.getAs('application/pdf');
    } catch (docError) {
      Logger.log('Not a Google Doc, trying as Google Slides...');

      // Try as Google Slides presentation
      try {
        const presentation = SlidesApp.openById(CONFIG.ADDITIONAL_PDF_DOC_ID);
        Logger.log('✅ Presentation accessed successfully: ' + presentation.getName());
        blob = presentation.getAs('application/pdf');
      } catch (slidesError) {
        throw new Error('Neither Google Doc nor Slides: ' + docError.message + ' | ' + slidesError.message);
      }
    }

    blob.setName('Kings_Equestrian_Presentation.pdf');
    Logger.log('✅ PDF generated successfully');
    return blob;
  } catch (e) {
    Logger.log('❌ getAdditionalPDF error: ' + e.message);
    Logger.log('Document ID being used: ' + CONFIG.ADDITIONAL_PDF_DOC_ID);
    Logger.log('💡 Troubleshooting tips:');
    Logger.log('  1. For Google Docs: Use the ID from https://docs.google.com/document/d/YOUR_ID/edit');
    Logger.log('  2. For Google Slides: Use the ID from https://docs.google.com/presentation/d/YOUR_ID/edit');
    Logger.log('  3. Ensure sharing is set to "Anyone with the link can view"');
    Logger.log('  4. Try publishing the document (File > Publish to web)');
    Logger.log('  5. Test URL manually: https://docs.google.com/document/d/' + CONFIG.ADDITIONAL_PDF_DOC_ID + '/edit');
  }

// ────────────────────────────────────────────────────────────
//  DIAGNOSTIC: Test document access
// ────────────────────────────────────────────────────────────

function testDocumentAccess(docId) {
  try {
    Logger.log('Testing access to document/presentation ID: ' + docId);

    let doc, blob, fileType;

    // Try as Google Doc first
    try {
      doc = DocumentApp.openById(docId);
      Logger.log('✅ SUCCESS: Google Doc "' + doc.getName() + '" is accessible');
      fileType = 'Google Doc';
      blob = doc.getAs('application/pdf');
    } catch (docError) {
      // Try as Google Slides presentation
      try {
        doc = SlidesApp.openById(docId);
        Logger.log('✅ SUCCESS: Google Slides "' + doc.getName() + '" is accessible');
        fileType = 'Google Slides';
        blob = doc.getAs('application/pdf');
      } catch (slidesError) {
        throw new Error('Neither Google Doc nor Slides accessible: ' + docError.message + ' | ' + slidesError.message);
      }
    }

    Logger.log('Document URL: https://docs.google.com/' + (fileType === 'Google Doc' ? 'document' : 'presentation') + '/d/' + docId + '/edit');
    Logger.log('✅ SUCCESS: PDF conversion works (' + blob.getBytes().length + ' bytes)');

    return true;
  } catch (e) {
    Logger.log('❌ FAILED: ' + e.message);
    Logger.log('💡 Possible solutions:');
    Logger.log('  1. Check if document/presentation ID is correct');
    Logger.log('  2. For Google Docs: Use ID from https://docs.google.com/document/d/YOUR_ID/edit');
    Logger.log('  3. For Google Slides: Use ID from https://docs.google.com/presentation/d/YOUR_ID/edit');
    Logger.log('  4. Ensure sharing is set to "Anyone with the link can view"');
    Logger.log('  5. Try publishing the document (File > Publish to web)');
    return false;
  }
}

function testMyAdditionalPDF() {
  Logger.log('🧪 Testing your additional PDF configuration...');
  const result = testDocumentAccess(CONFIG.ADDITIONAL_PDF_DOC_ID);
  if (result) {
    Logger.log('🎉 SUCCESS: Your additional PDF is ready for welcome emails!');
    Logger.log('📧 Next: Submit a test booking to see the PDF attachment');
  } else {
    Logger.log('❌ FAILED: Check the troubleshooting tips above');
  }
}
   

  try {
    const logoBlob = UrlFetchApp.fetch('https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg').getBlob();
    const lPara    = body.appendParagraph('');
    lPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    lPara.appendInlineImage(logoBlob).setWidth(100).setHeight(100);
    lPara.setSpacingAfter(16);
  } catch (e) { Logger.log('Consent logo: ' + e); }

  p('KINGS EQUESTRIAN FOUNDATION', 16, true, DocumentApp.HorizontalAlignment.CENTER);
  p('Acknowledgement & Consent Form', 13, true, DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('').setSpacingAfter(10);
  p('I/we acknowledge that horse-riding carries inherent risks including falls and injuries. Kings Equestrian Foundation takes all reasonable precautions but cannot guarantee against accidents.',11,false);
  p('I/we declare all participants are medically fit and hold valid personal accident insurance. Kings Equestrian Foundation is not liable for medical expenses.',11,false);
  body.appendParagraph('').setSpacingAfter(10);
  p('Details', 12, true);
  p('Name: ' + (name  || '____________________'), 11);
  p('Phone: '+ (phone || '____________________'), 11);
  p('Email: '+ (email || '____________________'), 11);
  body.appendParagraph('').setSpacingAfter(20);
  const dateStr = bookingDate ? fmtDate(new Date(bookingDate)) : '___________';
  p('Signature: ___________________________    Date: ' + dateStr, 11);
  body.appendParagraph('').setSpacingAfter(6);
  p('Kings Equestrian Foundation | Karnataka, India | +91-9980895533', 9, false, DocumentApp.HorizontalAlignment.CENTER);

  doc.saveAndClose();
  const pdf = doc.getAs('application/pdf');
  pdf.setName('Consent_' + (name || 'Participant').replace(/\s+/g,'_') + '.pdf');
  try { DriveApp.getFileById(doc.getId()).setTrashed(true); } catch(e){}
  return pdf;
}

// ============================================================
//  MENU & TRIGGERS
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎠 Kings Equestrian')
    .addItem('📧 Resend Welcome Email',    'resendWelcomeEmail')
    .addItem('🧾 Send Payment Receipt',    'sendPaymentReceiptMenu')
    .addSeparator()
    .addItem('📅 Send Daily Summary Now',  'testSendDailySummaryNow')
    .addItem('🧪 Dry-Run Daily Summary',   'testDailySummaryDryRun')
    .addSeparator()
    .addItem('⚙️  Setup All Triggers',     'setupTriggers')
    .addToUi();
}

function setupTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onBookingFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('onPaymentFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('sendDailyAdminSummary').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getUi().alert('✅ Triggers set!\n\n• Booking form → welcome email + KE No\n• Payment form → receipt email\n• Daily 7 AM → admin summary email');
}