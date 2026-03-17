// ============================================
// KINGS EQUESTRIAN - CONFIG & SHARED UTILITIES
// File: 1_Config.js
// ============================================

const CONFIG = {
    UPI_ID: "vyapar.176548151976@hdfcbank",
    BUSINESS_NAME: "KingsEquestrian",
    PAYMENT_FORM_LINK: "https://forms.gle/WxskpjCcDQWkA7L57",
    EMAIL_TEMPLATE_DOC_ID: "1bUTpk9QCR4n1uUmMuoSRRcflTShG3jawuhemE28aTio",
    TERMS_CONDITIONS_DOC_ID: "1QbJHA5keyTLvgw-5stTY74i92BQ89TYya-NvtJ4YGx4",
    ADVANCE_BOOKING_AMOUNT: 1000,
    webAppUrl: "https://script.google.com/macros/s/AKfycbxGNi137N_vvd6kFWe0CL2clALwKLp7QKsLgiWUd9fGcvYhTlaeQIy15n2vai_1g-PIig/exec",

    SHEETS: {
        BOOKING_FORM: "Booking Form Response",
        PAYMENT_FORM: "Payment Form Response",
        PRICING: "Pricing",
        MAIL_INFO: "Mail Info"
    },

    BOOKING_COLS: {
        TIMESTAMP: 0,
        NAME: 2,
        EMAIL_ID: 3,
        PHONE_NUMBER: 4,
        OUR_SERVICES: 5,
        NUMBER_OF_PARTICIPANTS: 6,
        PREFERRED_SERVICE_DATE: 7,
        PREFERRED_TIME_SLOT: 8,
        CONSENT: 9,
        REFERENCE: 10,
        WELCOME_EMAIL_SENT: 11,
        WELCOME_EMAIL_TIMESTAMP: 12,
    },

    PAYMENT_COLS: {
        TIMESTAMP: 0,
        // Column index 2: user submits their PHONE NUMBER here (used as primary lookup key)
        // Previously labelled "Registration No" in the form, now accepts phone number for easy access.
        // The actual KE-reference (if provided) is looked up from the Booking sheet by phone.
        PHONE_OR_REG: 2,       // Raw value the user typed — could be phone or KE-ref
        PHONE_NUMBER: 3,       // (kept for any legacy column; may be same as above)
        AMOUNT_PAID: 4,
        SCREENSHOT: 5,
        PAYMENT_DATE: 6,
        TRANSACTION_REFERENCE_NUMBER: 7,
        PAN_AADHAAR: 8,
        TRANSACTION_VERIFIED: 9,
        RECEIPT_SENT: 10,
        RECEIPT_SENT_TIMESTAMP: 11,
        PAYMENT_RECEIPT_NO: 12,
        PAYMENT_RECEIPT_DRIVER_LINK: 13
    }
};

// --------------- SHARED UTILITY FUNCTIONS ---------------

function generateReference() {
    const date = new Date();
    const year = date.getFullYear().toString().substr(-2);
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const random = Math.floor(Math.random() * 9000) + 1000;
    return `KE${year}${month}${day}${random}`;
}

function createUPILink(amount, reference) {
    return `upi://pay?pa=${CONFIG.UPI_ID}&pn=${encodeURIComponent(CONFIG.BUSINESS_NAME)}&am=${amount}&cu=INR&tn=${encodeURIComponent(reference)}`;
}

function createQRCode(link) {
    return `https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=${encodeURIComponent(link)}`;
}

function formatDate(date) {
    if (!date) return 'N/A';
    if (typeof date === 'string') return date;
    try {
        return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'dd MMM yyyy');
    } catch (e) {
        return String(date);
    }
}

function normalizeDate(dateValue) {
    if (!dateValue) return '';
    try {
        const date = new Date(dateValue);
        if (isNaN(date.getTime())) return String(dateValue);
        return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } catch (e) {
        return String(dateValue);
    }
}

function getCCRecipients(mailType) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const mailInfoSheet = ss.getSheetByName(CONFIG.SHEETS.MAIL_INFO);
        if (!mailInfoSheet) {
            Logger.log('Mail Info sheet not found');
            return [];
        }
        const data = mailInfoSheet.getDataRange().getValues();
        const ccEmails = [];
        for (let i = 1; i < data.length; i++) {
            const email = data[i][0];
            const type = data[i][1];
            if (email && type && type.toLowerCase().includes(mailType.toLowerCase())) {
                ccEmails.push(email);
            }
        }
        return ccEmails;
    } catch (error) {
        Logger.log('Error getting CC recipients: ' + error);
        return [];
    }
}

function getPricingData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pricingSheet = ss.getSheetByName(CONFIG.SHEETS.PRICING);
    if (!pricingSheet) throw new Error('Pricing sheet not found');
    const data = pricingSheet.getDataRange().getValues();
    const pricingMap = {};
    for (let i = 1; i < data.length; i++) {
        const service = data[i][1];
        const pricePerHalfHour = data[i][2];
        const docId = data[i][3];
        if (service) {
            pricingMap[service] = { price: pricePerHalfHour, docId: docId };
        }
    }
    return pricingMap;
}

function getServicePDF(docId, serviceName) {
    try {
        const doc = DocumentApp.openById(docId);
        const blob = doc.getAs('application/pdf');
        blob.setName(`${serviceName.replace(/\s+/g, '_')}_Details.pdf`);
        return blob;
    } catch (error) {
        Logger.log('Error creating PDF: ' + error);
        return null;
    }
}

function getTermsAndConditionsPDF() {
    try {
        const doc = DocumentApp.openById(CONFIG.TERMS_CONDITIONS_DOC_ID);
        const blob = doc.getAs('application/pdf');
        blob.setName('Terms_and_Conditions.pdf');
        return blob;
    } catch (error) {
        Logger.log('Error creating T&C PDF: ' + error);
        return null;
    }
}

// --------------- BOOKING LOOKUP HELPERS ---------------

/**
 * Determines if the submitted value looks like a KE-reference number.
 * KE-refs follow the pattern: KE + 2-digit year + 2-digit month + 2-digit day + 4 digits
 * e.g. KE250314xxxx
 */
function isKEReference(value) {
    return /^KE\d{8,}$/i.test(String(value).trim());
}

/**
 * Normalise a phone number to digits only for fuzzy matching.
 */
function normalizePhone(phone) {
    return String(phone || '').replace(/\D/g, '').slice(-10); // last 10 digits
}

/**
 * Finds a booking row by EITHER a KE-reference number OR a phone number.
 * Returns { rowIndex, row } or null.
 *
 * Priority:
 *  1. If value looks like a KE-ref → match by reference column
 *  2. Otherwise → match by phone number (last 10 digits)
 *  3. If value looks like KE-ref but not found → fallback to phone match
 */
function findBookingByPhoneOrRef(submittedValue) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bookingSheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_FORM);
    if (!bookingSheet) {
        Logger.log('Booking sheet not found');
        return null;
    }

    const bookingValues = bookingSheet.getDataRange().getValues();
    const trimmed = String(submittedValue || '').trim();

    // --- Attempt 1: match by KE-reference if it looks like one ---
    if (isKEReference(trimmed)) {
        for (let j = 1; j < bookingValues.length; j++) {
            const ref = String(bookingValues[j][CONFIG.BOOKING_COLS.REFERENCE] || '').trim();
            if (ref.toLowerCase() === trimmed.toLowerCase()) {
                Logger.log(`Booking found by KE-reference: ${trimmed} at row ${j + 1}`);
                return { rowIndex: j + 1, row: bookingValues[j] };
            }
        }
        Logger.log(`KE-reference ${trimmed} not found, falling back to phone match`);
    }

    // --- Attempt 2: match by phone number (last 10 digits) ---
    const normalizedInput = normalizePhone(trimmed);
    if (normalizedInput.length >= 10) {
        for (let j = bookingValues.length-1; j >= 1; j--) {
            const bookingPhone = normalizePhone(bookingValues[j][CONFIG.BOOKING_COLS.PHONE_NUMBER]);
            if (bookingPhone === normalizedInput) {
                Logger.log(`Booking found by phone: ${normalizedInput} at row ${j + 1}`);
                return { rowIndex: j + 1, row: bookingValues[j] };
            }
        }
    }

    Logger.log(`No booking found for submitted value: ${trimmed}`);
    return null;
}

// --------------- MENU SETUP ---------------

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🎠 Kings Equestrian')
        .addItem('📧 Resend Welcome Email', 'ResendWelcomeEmail')
        .addItem('🧾 Send Payment Receipt', 'SendPaymentReceipt')
        .addSeparator()
        .addItem('⚙️ Setup Triggers', 'setupTriggers')
        .addSeparator()
        .addItem('📅 Send Daily Summary Now', 'testSendDailySummaryNow')
        .addItem('🧪 Test Summary (Dry Run)', 'testDailySummaryDryRun')
        .addItem('⚙️ Setup New Feature Triggers', 'setupNewFeaturesTriggers')
        .addToUi();
}

function setupTriggers() {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    ScriptApp.newTrigger('onBookingFormSubmit')
        .forSpreadsheet(ss)
        .onFormSubmit()
        .create();

    ScriptApp.newTrigger('onPaymentFormSubmit')
        .forSpreadsheet(ss)
        .onFormSubmit()
        .create();

    SpreadsheetApp.getUi().alert('✅ Triggers set up successfully!\n\n' +
        'The system will now automatically:\n' +
        '- Generate reference numbers and send welcome emails on booking\n' +
        '- Auto-send receipts when payment form is submitted\n' +
        '- Resend existing receipts for duplicate submissions\n' +
        '- Store receipts in Google Drive');
}