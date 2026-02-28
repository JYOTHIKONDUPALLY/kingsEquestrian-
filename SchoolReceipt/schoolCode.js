// ============================================
// KINGS EQUESTRIAN - SCHOOL REGISTRATION RECEIPTS
// ============================================

// --------------- CONFIG ---------------

const SCHOOL_CONFIG = {
    SHEET_NAME: 'SchoolRegistration',

    // Column indices (0-based) — adjust if your sheet differs
    COLS: {
        SR_NO: 0,                  // A - Sr. No (used for sequence)
        STUDENT_NAME: 1,           // B
        GRADE_SECTION: 2,// C  e.g. "Grade 5 - A"
        LOCATION:3,
        PARENT_NAME: 4,            
        CONTACT_NUMBER: 5,
        CONCENT_FORM:6,
        REGISTRATION_FEES: 7, 
        DATE_OF_PAYMENT: 8,        
        MODE_OF_PAYMENT: 9,
        PAN_AADHAAR:10,  
        ADDRESS:11,      
        PAYMENT_STATUS: 12,         // J  "Paid" / "Due"
        RECEIPT_NO: 13,            // L  ← written back after generation
        RECEIPT_DRIVE_LINK: 14     // M  ← written back after generation
    },

    // Google Drive folder name for all school receipts
    ROOT_FOLDER: 'Kings Equestrian - School Receipts',

    // Drive file IDs reused from the main script
    STAMP_FILE_ID: '1ReTtpLb8gNVKMaqdQuUgJVnMA64SJbJQ',
    SIGN_FILE_ID:  '1_zCecmE02RSevbKvnUlY3Y14B1RV9k9o',
    LOGO_URL: 'https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg'
};

// -----------------------------------------------
// MENU  (registers under a dedicated top-level menu)
// -----------------------------------------------

function onOpen_School() {
    SpreadsheetApp.getUi()
        .createMenu('🏫 School Registration')
        .addItem('🧾 Generate Receipt(s) for Selected Rows', 'generateSchoolReceiptsForSelection')
        .addToUi();
}

// Call this from the master onOpen if you have one, or set it as its own
// installable trigger on the spreadsheet.

// -----------------------------------------------
// MAIN ENTRY POINT  (triggered by menu)
// -----------------------------------------------

function generateSchoolReceiptsForSelection() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SCHOOL_CONFIG.SHEET_NAME);

    if (!sheet) {
        ui.alert(`❌ Sheet "${SCHOOL_CONFIG.SHEET_NAME}" not found.\nPlease check the sheet name in SCHOOL_CONFIG.`);
        return;
    }

    // Warn if user is on a different tab — selection must be on SchoolRegistration
    const activeSheet = ss.getActiveSheet();
    if (activeSheet.getName() !== SCHOOL_CONFIG.SHEET_NAME) {
        ui.alert(
            '⚠️ Wrong Tab',
            `Please switch to the "${SCHOOL_CONFIG.SHEET_NAME}" tab first, ` +
            `select the rows you want, then click this menu again.`,
            ui.ButtonSet.OK
        );
        return;
    }

    const selection = sheet.getActiveRange();
    if (!selection || selection.getRow() <= 1) {
        ui.alert('⚠️ Please select one or more data rows first (not the header row).');
        return;
    }

    const startRow = selection.getRow();
    const numRows  = selection.getNumRows();

    const confirm = ui.alert(
        '🧾 Generate School Receipts',
        `Generate receipts for ${numRows} selected row(s)?`,
        ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;

    let successCount = 0;
    let failCount    = 0;
    const errors     = [];

    for (let i = 0; i < numRows; i++) {
        const rowIndex = startRow + i;
        try {
            generateSchoolReceiptForRow(sheet, rowIndex);
            successCount++;
            Utilities.sleep(800);
        } catch (err) {
            failCount++;
            errors.push(`Row ${rowIndex}: ${err.message}`);
            Logger.log(`School receipt failed at row ${rowIndex}: ${err.message}`);
        }
    }

    let msg = `✅ Generated: ${successCount}   ❌ Failed: ${failCount}`;
    if (errors.length > 0) {
        msg += '\n\nErrors:\n' + errors.slice(0, 5).join('\n');
        if (errors.length > 5) msg += `\n…and ${errors.length - 5} more`;
    }
    ui.alert(msg);
}

// -----------------------------------------------
// CORE — generate receipt for one row
// -----------------------------------------------

function generateSchoolReceiptForRow(sheet, rowIndex) {
    const C = SCHOOL_CONFIG.COLS;
    const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

    // ── Read fields ──────────────────────────────────────────────────────────
    const studentName   = String(row[C.STUDENT_NAME]   || '').trim();
    const gradeSection  = String(row[C.GRADE_SECTION]  || '').trim();
    const parentName    = String(row[C.PARENT_NAME]    || '').trim();
    const contactNumber = String(row[C.CONTACT_NUMBER] || '').trim();
    const address       = String(row[C.ADDRESS]        || '').trim();
    const panAadhaar    = String(row[C.PAN_AADHAAR]    || '').trim();
    const modeOfPayment = String(row[C.MODE_OF_PAYMENT]|| '').trim();
    const paymentStatus = String(row[C.PAYMENT_STATUS] || 'Paid').trim();
    const dateOfPayment = row[C.DATE_OF_PAYMENT];
    const amount        = Number(row[C.REGISTRATION_FEES]);

    console.log("amount", amount)

    if (!studentName) throw new Error('Student name is missing');
    if (!parentName)  throw new Error('Parent name is missing');
    if (!amount || isNaN(amount)) throw new Error('Valid registration fee is required');

    // ── Generate receipt number ───────────────────────────────────────────────
    // Format: KE(YY)(MM)(DD)(GradeCode)(4-digit sequence)
    // GradeCode = first alphanumeric chars of grade/section (e.g. "5A" from "Grade 5 - A")
    const receiptNumber = generateSchoolReceiptNumber(gradeSection, rowIndex);

    // ── Build PDF ─────────────────────────────────────────────────────────────
    Logger.log(`Generating school receipt for ${studentName} — ${receiptNumber}`);
    const receiptPDF = generateSchoolReceiptPDF(
        studentName, gradeSection, parentName, contactNumber,
        address, panAadhaar, modeOfPayment, paymentStatus,
        dateOfPayment, amount, receiptNumber
    );

    // ── Store in Drive ────────────────────────────────────────────────────────
    const driveInfo = storeSchoolReceiptInDrive(receiptPDF, studentName, gradeSection, receiptNumber);

    // ── Write back receipt number & Drive link ────────────────────────────────
    if (driveInfo) {
        sheet.getRange(rowIndex, C.RECEIPT_NO + 1).setValue(receiptNumber);
        sheet.getRange(rowIndex, C.RECEIPT_DRIVE_LINK + 1)
            .setValue(driveInfo.fileUrl);
    }

    Logger.log(`School receipt stored: ${driveInfo ? driveInfo.fileName : 'unknown'}`);
    return true;
}

// -----------------------------------------------
// RECEIPT NUMBER GENERATOR
// -----------------------------------------------

function generateSchoolReceiptNumber(gradeSection, rowIndex) {
    const now = new Date();
    const yy  = now.getFullYear().toString().slice(-2);
    const mm  = String(now.getMonth() + 1).padStart(2, '0');
    const dd  = String(now.getDate()).padStart(2, '0');

    // Extract a short grade code from gradeSection
    // "Grade 5 - A"  →  "5A"
    // "10B"          →  "10B"
    // "LKG"          →  "LKG"
    const gradeCode = gradeSection
        .replace(/grade/gi, '')
        .replace(/section/gi, '')
        .replace(/[-–\s]+/g, '')
        .toUpperCase()
        .substring(0, 4) || 'SCH';

    // 4-digit sequence from row index (padded)
    const seq = String(rowIndex).padStart(4, '0');

    return `KE${yy}${mm}${dd}${gradeCode}${seq}`;
}

// -----------------------------------------------
// PDF GENERATOR
// -----------------------------------------------

function generateSchoolReceiptPDF(
    studentName, gradeSection, parentName, contactNumber,
    address, pan, modeOfPayment, paymentStatus,
    dateOfPayment, amount, receiptNumber
) {
    Logger.log('Fetching logo...');
    const logoBase64  = schoolGetImageFromUrl(SCHOOL_CONFIG.LOGO_URL);
    Logger.log('Fetching stamp...');
    const stampBase64 = schoolGetImageFromDrive(SCHOOL_CONFIG.STAMP_FILE_ID);
    Logger.log('Fetching signature...');
    const signBase64  = schoolGetImageFromDrive(SCHOOL_CONFIG.SIGN_FILE_ID);

    const htmlContent = buildSchoolReceiptHTML(
        studentName, gradeSection, parentName, contactNumber,
        address, pan, modeOfPayment, paymentStatus,
        dateOfPayment, amount, receiptNumber,
        logoBase64, stampBase64, signBase64
    );

    const tempFile = DriveApp.createFile(
        `school_receipt_temp_${Date.now()}.html`,
        htmlContent,
        MimeType.HTML
    );
    const pdfBlob = tempFile.getAs('application/pdf');
    const safeName = studentName.replace(/\s+/g, '_');
    const safeGrade = gradeSection.replace(/[\s\/\\:*?"<>|]+/g, '_');
    pdfBlob.setName(`Receipt_${safeName}_${safeGrade}_${receiptNumber}.pdf`);
    tempFile.setTrashed(true);

    return pdfBlob;
}

// -----------------------------------------------
// HTML RECEIPT TEMPLATE
// -----------------------------------------------

function buildSchoolReceiptHTML(
    studentName, gradeSection, parentName, contactNumber,
    address, pan, modeOfPayment, paymentStatus,
    dateOfPayment, amount, receiptNumber,
    logoBase64, stampBase64, signBase64
) {
    const formattedDate = schoolFormatDate(dateOfPayment);
    const amountInWords = schoolNumberToWords(amount);
    const currentDate   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const isPaid        = String(paymentStatus).toLowerCase() === 'paid';

    return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
    @page { size: A4; margin: 0; }

    body {
        font-family: "Times New Roman", serif;
        margin: 0;
        padding: 25px;
        background: #fff;
    }

    .receipt-container {
        border: 2px solid #000;
        border-radius: 35px;
        padding: 25px 30px;
        max-width: 800px;
        margin: auto;
        position: relative;
    }

    /* HEADER */
    .header { display: flex; align-items: flex-start; }

    .logo-section { width: 140px; text-align: center; }
    .logo-img { width: 110px; }

    .header-center { flex: 1; text-align: center; }
    .org-name { font-size: 28px; font-weight: bold; margin-bottom: 5px; }
    .registration-info { font-size: 13px; }
    .registration-subdetails { font-size: 13px; margin-top: 3px; }
    .subtext {
        margin-top: 10px;
        font-style: italic;
        font-weight: bold;
        text-decoration: underline;
        font-size: 12px;
    }

    /* RECEIPT NUMBER */
    .receipt-number {
        position: absolute;
        right: 30px;
        top: 15px;
        font-size: 14px;
        font-weight: bold;
        color: red;
    }

    /* TITLE BOX */
    .receipt-box {
        border: 2px solid #000;
        border-radius: 12px;
        text-align: center;
        padding: 8px 10px;
        margin: 18px 0 10px;
    }
    .receipt-title { font-size: 20px; font-weight: bold; }
    .receipt-subtitle { font-size: 12px; }
     /* TWO COLUMN LAYOUT */
    .main-content {
        display: flex;
        gap: 30px;
        margin-top: 10px;
    }

    .left-column,
    .right-column {
        flex: 1;
        font-size: 14px;
    }

    /* CHECKBOX */
    .checkbox-item {
        margin: 5px 0;
    }

    .checkbox {
        display: inline-block;
        width: 13px;
        height: 13px;
        border: 1px solid #000;
        margin-right: 6px;
        vertical-align: middle;
    }

    .checkbox.checked {
        background: #000;
        position: relative;
    }

    .checkbox.checked::after {
        content: "✓";
        color: #fff;
        font-size: 11px;
        position: absolute;
        left: 1px;
        top: -2px;
    }

    /* DONOR DETAILS */
    .detail-row {
        margin: 8px 0;
    }

    .detail-label {
        font-weight: bold;
    }

    .date-row { text-align: right; font-size: 14px; margin-bottom: 12px; }

    /* DETAILS TABLE */
    .details-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 14px;
        margin: 12px 0;
    }
    .details-table td {
        padding: 6px 8px;
        vertical-align: top;
    }
    .details-table td:first-child {
        font-weight: bold;
        width: 38%;
        white-space: nowrap;
    }
    .details-table tr:nth-child(even) td {
        background: #f9f9f9;
    }

    /* AMOUNT BOX */
    .amount-section {
        border: 2px solid #000;
        margin: 18px 0;
        padding: 16px;
        position: relative;
        text-align: center;
    }
    .rupee-symbol {
        position: absolute;
        left: 20px;
        top: 50%;
        transform: translateY(-50%);
        font-size: 40px;
        color: goldenrod;
        font-weight: bold;
    }
    .amount-value { font-size: 34px; font-weight: bold; }

    .status-badge {
        display: inline-block;
        padding: 3px 14px;
        border-radius: 20px;
        font-size: 13px;
        font-weight: bold;
        margin-left: 10px;
        vertical-align: middle;
    }
    .status-paid { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
    .status-due  { background: #fff3cd; color: #856404; border: 1px solid #ffeeba; }

    .payment-mode { font-size: 14px; margin-top: 10px; }

    .declaration-section {
        margin-top: 14px;
        font-size: 13px;
        text-align: justify;
    }

    /* SIGNATURE */
    .signature-section { margin-top: 36px; text-align: right; }
    .org-label { font-weight: bold; margin-bottom: 5px; }
    .stamp-and-sign { position: relative; height: 120px; }
    .sign-img  { width: 110px; }
    .stamp-img { width: 120px; }
    .authorized-text {
        margin-top: 90px;
        text-decoration: underline;
        font-size: 14px;
    }
</style>
</head>
<body>
<div class="receipt-container">

    <div class="receipt-number">${receiptNumber}</div>

    <!-- HEADER -->
    <div class="header">
        <div class="logo-section">
            <img src="${logoBase64}" class="logo-img" alt="Logo" />
        </div>
        <div class="header-center">
            <div class="org-name">Kings Equestrian Foundation</div>
            <div class="registration-info">
                Registered u/s 80G of Income-tax Act &nbsp;|&nbsp; Rg no: AAJCK7191GE20231 &nbsp;|&nbsp; PAN: AAJCK7191G
            </div>
            <div class="registration-subdetails">
                K202, Tower-6, Jacaranda Block, Devarabisanahalli, Bellandur S.O, Bengaluru – 560103, Karnataka, India<br>
                kingsequestrianfoundation@gmail.com &nbsp;|&nbsp; kingsequestrianfoundation.com
            </div>
            <div class="subtext">
               We gratefully acknowledge your generous contribution in support of our programmes promoting education, well-being, and personal development through sport and experiential learning.
            </div>
        </div>
    </div>

    <!-- TITLE BOX -->
    <div class="receipt-box">
        <div class="receipt-title">Registration Receipt</div>
        <div class="receipt-subtitle">This receipt is issued in compliance with Rule 18AB and Form 10BD requirements</div>
    </div>

    <div class="date-row"><strong>Receipt Date:</strong> ${currentDate}</div>

    <!-- STUDENT & PARENT DETAILS -->
      <div class="main-content">
        <div class="left-column">
            <div class="section-title">Donor Category (✓ Tick Applicable)</div>
            <div class="checkbox-item">
                <span class="checkbox checked"></span> Resident Indian Donor
            </div>
            <div class="checkbox-item">
                <span class="checkbox"></span> Non-Resident Indian (NRI)
            </div>
        </div>

        <div class="right-column">
            <div class="section-title">Donor Details</div>
            <div class="detail-row"><span class="detail-label">Name of Donor:</span> ${parentName}</div>
            <div class="detail-row"><span class="detail-label">PAN / Aadhaar:</span> ${pan}</div>
        </div>
    </div>

    <!-- AMOUNT BOX -->
    <div class="amount-section">
        <span class="rupee-symbol">₹</span>
        <div class="amount-value">
            ${amount.toLocaleString('en-IN')}
        </div>
    </div>

    <!-- PAYMENT MODE -->
    <div class="payment-mode">
        <strong>Mode of Payment:</strong> ${modeOfPayment || '—'}<br><br>
        <strong>Amount in Words:</strong> ${amountInWords}
    </div>

    <!-- DECLARATION -->
    <div class="declaration-section">
        Certified that the above donation is received by trust for charitable purposes only. This donation is eligible for deduction under Section 80G of the Income Tax Act, 1961. This receipt will be reported in Form 10BD and Form 10BE will be issued to the donor.
    </div>

    <!-- SIGNATURE -->
    <div class="signature-section">
        <div class="org-label">For Kings Equestrian Foundation</div>
        <div class="stamp-and-sign">
            <img src="${signBase64}"  class="sign-img"  alt="Signature" />
            <img src="${stampBase64}" class="stamp-img" alt="Stamp" />
           
        </div>
    </div>

</div>
</body>
</html>`;
}

// -----------------------------------------------
// DRIVE STORAGE
// -----------------------------------------------

function storeSchoolReceiptInDrive(receiptBlob, studentName, gradeSection, receiptNumber) {
    try {
        const folder = getSchoolClassFolder(gradeSection);

        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
        const safeName  = studentName.replace(/\s+/g, '_');
        const fileName  = `Receipt_${safeName}_${receiptNumber}_${timestamp}.pdf`;

        const file = folder.createFile(receiptBlob);
        file.setName(fileName);
        file.setDescription(`School Registration Receipt | Student: ${studentName} | Grade: ${gradeSection} | Receipt No: ${receiptNumber}`);

        Logger.log(`School receipt saved: ${fileName}`);

        return {
            fileId:    file.getId(),
            fileUrl:   file.getUrl(),
            fileName:  fileName,
            folderId:  folder.getId(),
            folderUrl: folder.getUrl()
        };
    } catch (err) {
        Logger.log('Error storing school receipt in Drive: ' + err);
        return null;
    }
}

// -----------------------------------------------
// FOLDER HELPER
// Kings Equestrian - School Receipts
//   └── 2025
//       └── Grade 5 - A
//           └── Receipt_...pdf
// -----------------------------------------------

function getSchoolClassFolder(gradeSection) {
    const rootName  = SCHOOL_CONFIG.ROOT_FOLDER;
    const yearName  = String(new Date().getFullYear());
    const gradeName = gradeSection.trim() || 'Unclassified';

    // Root folder
    let rootIter = DriveApp.getFoldersByName(rootName);
    const rootFolder = rootIter.hasNext()
        ? rootIter.next()
        : DriveApp.createFolder(rootName);

    // Year sub-folder
    let yearIter = rootFolder.getFoldersByName(yearName);
    const yearFolder = yearIter.hasNext()
        ? yearIter.next()
        : rootFolder.createFolder(yearName);

    // Grade/Section sub-folder
    let gradeIter = yearFolder.getFoldersByName(gradeName);
    return gradeIter.hasNext()
        ? gradeIter.next()
        : yearFolder.createFolder(gradeName);
}

// -----------------------------------------------
// UTILITY HELPERS  (self-contained, no dependency on main script)
// -----------------------------------------------

function schoolGetImageFromUrl(url) {
    try {
        const resp = UrlFetchApp.fetch(url);
        const b64  = Utilities.base64Encode(resp.getBlob().getBytes());
        return `data:${resp.getBlob().getContentType()};base64,${b64}`;
    } catch (e) {
        Logger.log('schoolGetImageFromUrl error: ' + e);
        return '';
    }
}

function schoolGetImageFromDrive(fileId) {
    try {
        const blob = DriveApp.getFileById(fileId).getBlob();
        const b64  = Utilities.base64Encode(blob.getBytes());
        return `data:${blob.getContentType()};base64,${b64}`;
    } catch (e) {
        Logger.log('schoolGetImageFromDrive error: ' + e);
        return '';
    }
}

function schoolFormatDate(dateValue) {
    if (!dateValue) return '—';
    try {
        return Utilities.formatDate(new Date(dateValue), Session.getScriptTimeZone(), 'dd MMM yyyy');
    } catch (e) {
        return String(dateValue);
    }
}

function schoolNumberToWords(num) {
    const ones  = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine'];
    const teens = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen',
                   'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
    const tens  = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty',
                   'Sixty', 'Seventy', 'Eighty', 'Ninety'];

    function convert(n) {
        if (n === 0)        return 'Zero';
        if (n < 10)         return ones[n];
        if (n < 20)         return teens[n - 10];
        if (n < 100)        return tens[Math.floor(n / 10)] + (n % 10 ? ' ' + ones[n % 10] : '');
        if (n < 1000)       return ones[Math.floor(n / 100)] + ' Hundred' + (n % 100 ? ' ' + convert(n % 100) : '');
        if (n < 100000)     return convert(Math.floor(n / 1000)) + ' Thousand' + (n % 1000 ? ' ' + convert(n % 1000) : '');
        if (n < 10000000)   return convert(Math.floor(n / 100000)) + ' Lakh' + (n % 100000 ? ' ' + convert(n % 100000) : '');
        return convert(Math.floor(n / 10000000)) + ' Crore' + (n % 10000000 ? ' ' + convert(n % 10000000) : '');
    }

    return convert(Math.round(num)).trim() + ' Rupees Only';
}