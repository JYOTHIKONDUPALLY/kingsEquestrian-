// ============================================================
// INDUS EQUESTRIAN — SCHOOL SYSTEM
// File: 1_Config.gs
// Central config, column maps, and shared utilities
// ============================================================

// ⚠️  PAYMENT FORM COLUMN NOTE
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Payment Response sheet column order (0-based) — "Payment Responses 26-27":
//   0  Timestamp
//   1  Registration No
//   2  Phone number
//   3  Amount Paid (₹)
//   4  ScreenShot
//   5  Payment Date
//   6  Transcation Reference Number
//   7  Pan / AAdhar Number
//   8  Mode of Payment
//   9  Payment For (new form question; header-based lookup also supports other positions)
//  10  Receipt Sent
//  11  Receipt Sent At
//  12  Receipt No
//  13  Receipt Link
// Student / parent / email / grade are not on this form — use REG_REF + Riders
// (PAYMENT_COLS.STUDENT … ADDRESS = -1).
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const CONFIG = {
  ACADEMIC_YEAR_LABEL         : '2026-27',
  BUSINESS_NAME               : 'Kings Equestrian Foundation',
  /**
   * Bump this EVERY time you Deploy → Edit → New version.
   * Clients compare this to the version baked into their cached HTML.
   * If different, they auto-reload once so stale users pick up the new UI.
   */
  APP_UI_VERSION              : '2026-08-24b',

  // ── Location (Hyderabad campus) ───────────────────────────
  SCHOOL_NAME                 : '',  // leave blank — location branding uses city only
  LOCATION_CODE               : 'HYD',
  LOCATION_CITY               : 'Hyderabad',
  LOCATION_STATE              : 'Telangana',
  LOCATION_COUNTRY            : 'India',
  BUSINESS_ADDRESS            : 'Hyderabad, Telangana, India',
  CONTACT_PHONE               : '+91-9980895533',
  CONTACT_EMAIL               : 'kingsequestrianhyderabad@gmail.com',

  // ── Email sending (Brevo) ─────────────────────────────────
  // The API key itself lives in Script Properties (BREVO_API_KEY), NOT here.
  // MAIL_FROM must be a sender/domain VERIFIED in your Brevo account.
  MAIL_FROM                   : 'kingsequestrianhyderabad@gmail.com',
  MAIL_FROM_NAME              : 'Kings Equestrian Foundation',

  UPI_ID                      : 'vyapar.176548151976@hdfcbank',
  ADVANCE_BOOKING_AMOUNT      : 0,          // set to 0 — school registrations do not collect advance
  BACKUP_FOLDER_ID            : '',         // optional: set a Drive folder ID, else auto-created
  BACKUP_KEEP_DAYS            : 60,

  // ── Payment / registration form ──────────────────────────
  PAYMENT_FORM_BASE_URL       : 'https://forms.gle/8Rk6ojt2PGjJnzn86',
  PREFILL_ENTRY_IDS: {
    regRef   : '',   // fill these in after creating the Google Form
    student  : '',
    parent   : '',
    email    : '',
    phone    : '',
    grade    : ''
  },

  // ── Drive / file IDs ─────────────────────────────────────
  DRIVE_ROOT_FOLDER           : 'IndusSchool Registration 26-27',
  RECEIPTS_SUBFOLDER          : 'Receipts',
  SCHOOL_PROGRAM_INFO_DOC_ID  : 'https://drive.google.com/file/d/1CtMJouMyhbpPQzdBbmBRKYlUNXhCqSFk/view?usp=sharing',
  SUMMER_PROGRAM_INFO_DOC_ID  : '',        // fill if a summer-specific brochure exists
  STAMP_FILE_ID               : '1PL7IulIdSnbQvoDUDyq0Pre6P6xbRK65',
  SIGN_FILE_ID                : '1qiDCyQIr6BbWV8QRJPyQrZ7FjPP0t6Dn',
  LOGO_URL                    : 'https://kingsfarmequestrian.com/wp-content/uploads/2023/08/Logo2.jpg',

  // ── Portal URLs ───────────────────────────────────────────
  MY_RIDES_PORTAL_URL         : 'https://script.google.com/macros/s/AKfycbyDxlHFMJDUIzImJF-RLLTJwCtViRY9XJKWeejiLDiILmJ_SV_KhuQMd1WOu94DnGcU/exec?app=portal',
  ATTENDANCE_APP_URL          : 'https://script.google.com/macros/s/AKfycbyDxlHFMJDUIzImJF-RLLTJwCtViRY9XJKWeejiLDiILmJ_SV_KhuQMd1WOu94DnGcU/exec',

  // ── Sheet names ───────────────────────────────────────────
  SHEETS: {
    REGISTRATION_FORM : 'Registration Response',
    PAYMENT_FORM      : 'Payment Responses 26-27',
    RIDERS            : 'Riders',
    SCHEDULE          : 'Schedule',
    PAYMENTS          : 'Payments Ledger',
    PRICING           : 'service',                 // your SERVICE_SHEET
    MAIL_INFO         : 'Mail Info',
    EMAIL_LOG         : 'Email Log',
    TRAINERS          : 'TRAINERS',                 // attendance app login + audit
    SHOP_PRODUCTS     : 'SHOP_PRODUCTS',
    SHOP_ORDERS       : 'SHOP_ORDERS',
    GROOMERS          : 'GROOMERS',
    GROOMER_ATTENDANCE: 'GROOMER_ATTENDANCE',
    GROOMER_LEAVES    : 'GROOMER_LEAVES',
    HORSES            : 'HORSES',
    FEED_STOCK        : 'FEED_STOCK',
    FEED_MOVEMENTS    : 'FEED_MOVEMENTS',
    TACK_STOCK        : 'TACK_STOCK',
    TACK_MOVEMENTS    : 'TACK_MOVEMENTS',
    FEED_MOVEMENTS_ARCHIVE : 'FEED_MOVEMENTS_ARCHIVE',
    TACK_MOVEMENTS_ARCHIVE : 'TACK_MOVEMENTS_ARCHIVE',
    HORSE_CARE_TYPES       : 'HORSE_CARE_TYPES',
    HORSE_CUSTOM_CARE      : 'HORSE_CUSTOM_CARE',
    HORSE_ACTIVITY_LOG     : 'HORSE_ACTIVITY_LOG'
  },

  // ── Registration Form Response columns (0-based) ──────────
  // Adjust indices to match your actual Google Form column order.
  // Set optional/removed fields to -1 so downstream reads resolve to undefined.
  // EXTRA_REG_HEADERS are appended by the script after the last form column.
  REG_COLS: {
    TIMESTAMP    : 0,
    STUDENT      : 2,   // Student Name
    GRADE        : 3,   // Grade (separate form field)
    SECTION      : 4,   // Section (separate form field — insert after Grade in Google Form)
    PARENT       : 1,   // Parent / Guardian Name
    PHONE        : 5,
    EMAIL        : 6,
    PROGRAM      : 7,   // "school" | "summer" — value from form
    HORSE_LEASE  : -1,  // removed from form
    DOB          : -1,
    ADDRESS      : 8,
    MOTHER_NAME      : -1, // removed from form
    FATHER_NAME      : -1, // removed from form
    MOTHER_CONTACT   : 4,  // fallback to primary phone
    MOTHER_WHATSAPP  : 4,  // fallback to primary phone
    FATHER_CONTACT   : 4,  // fallback to primary phone
    FATHER_WHATSAPP  : 4,  // fallback to primary phone
    EMERGENCY_CONTACT: 4,  // fallback to primary phone
    RELATIONSHIP     : -1, // removed from form
    CONSENT_DATE : 9,  // date of consent (may equal timestamp)
    // ── Script-appended columns (after last form column) ──
    WELCOME_SENT : 10,
    WELCOME_AT   : 11,
    KE_NO        : 12,
    REG_REF      : 13,  // Registration Ref / KE No
    PROGRAM_TRACK: 14   // normalised program track stored by script
  },

  EXTRA_REG_HEADERS: ['Registration Ref', 'Program Track', 'Welcome Email Sent', 'Welcome Sent At', 'KE No'],

  // ── Payment Form Response columns (0-based) ──────────────
  // Matches linked form headers for "Payment Responses 26-27".
  PAYMENT_COLS: {
    TIMESTAMP    : 0,
    REG_REF      : 1,   // Registration No
    PHONE        : 2,
    AMOUNT       : 3,   // Amount Paid (₹)
    SCREENSHOT   : 4,
    PAY_DATE     : 5,
    TXN_REF      : 6,   // Transcation Reference Number (form spelling)
    PAN          : 7,   // Pan / AAdhar Number
    MODE         : 8,
    PAYMENT_FOR  : -1,  // resolved by header: Shopping Kit / Equipment | Riding Classes
    RECEIPT_SENT : 9,
    RECEIPT_AT   : 10,
    RECEIPT_NO   : 11,
    RECEIPT_LINK : 12,
    // Not collected on this payment form — receipt flow uses Riders / registration
    STUDENT      : -1,
    PARENT       : -1,
    EMAIL        : -1,
    GRADE        : -1,
    ADDRESS      : -1
  },

  // ── Riders sheet columns (0-based) ───────────────────────
  // NOTE: PARTICIPANTS column is retained in the schema for
  // cross-compatibility with Schedule / Portal code, but the
  // value is always written as 1 for school registrations.
  RIDER_COLS: {
    KE_NO        : 0,
    NAME         : 1,
    EMAIL        : 2,
    PHONE        : 3,
    SERVICES     : 4,
    PARTICIPANTS : 5,   // always 1 for school; kept for portal compatibility
    REGISTERED   : 6,
    NOTES        : 7
  },

  // ── Schedule sheet columns (0-based) ─────────────────────
  SCHED_COLS: {
    KE_NO        : 0,
    NAME         : 1,
    PHONE        : 2,
    EMAIL        : 3,
    SERVICE      : 4,
    DATE         : 5,
    TIME_SLOT    : 6,
    PARTICIPANTS : 7,   // always 1
    STATUS       : 8,
    ATTENDANCE   : 9,
    STAFF_NOTES  : 10,
    CAL_EVENT_ID : 11,
    SOURCE       : 12,
    BOOKED_BY    : 13,  // trainer who booked (or 'Self' when rider booked)
    SCORED_BY    : 14   // trainer who marked attendance / scored
  },

  // ── TRAINERS sheet columns (0-based) ─────────────────────
  TRAINER_COLS: {
    NAME      : 0,
    EMAIL     : 1,
    PHONE     : 2,
    AADHAR    : 3,
    USERNAME  : 4,
    PASSWORD  : 5,   // plaintext (admin convenience, per request)
    PASS_HASH : 6,   // SHA-256 hash used for verification
    ACTIVE    : 7,
    ROLE      : 8    // Admin can manage shop orders; Trainer is read-only/no access
  },

  SHOP_PRODUCT_COLS: {
    PRODUCT_ID    : 0,
    PRODUCT       : 1,
    CATEGORY      : 2,
    OPTION        : 3,
    ALLOWED_SIZES : 4,  // pipe-separated, or TEXT / UK_SIZE
    PRICE         : 5,
    IMAGE_URL     : 6,
    ACTIVE        : 7,
    SORT_ORDER    : 8
  },

  SHOP_ORDER_COLS: {
    CREATED_AT             : 0,
    ORDER_ID               : 1,
    KE_NO                  : 2,
    RIDER_NAME             : 3,
    EMAIL                  : 4,
    PHONE                  : 5,
    ITEMS_JSON             : 6,
    TOTAL                  : 7,
    PAYMENT_STATUS         : 8,
    PAYMENT_VERIFIED_BY    : 9,
    PAYMENT_VERIFIED_AT    : 10,
    ORDER_STATUS           : 11,
    EXPECTED_DATE          : 12,
    DELIVERED_AT           : 13,
    DELIVERY_NOTES         : 14,
    PARENT_CONFIRM_STATUS  : 15,
    PARENT_CONFIRMED_AT    : 16,
    PARENT_NOTE            : 17,
    UPDATED_AT             : 18,
    UPDATED_BY             : 19,
    PAYMENT_FORM_URL       : 20,
    UPI_REFERENCE          : 21,
    STATUS_EMAIL_AT        : 22,
    CLIENT_REQUEST_ID      : 23,
    PAYMENT_TXN_REF        : 24
  },

  GROOMER_COLS: {
    STAFF_ID           : 0,  // Employee ID
    NAME               : 1,
    PHOTO_URL          : 2,
    AADHAAR            : 3,
    PHONE              : 4,  // Mobile Number
    DESIGNATION        : 5,  // Staff / Groomer
    JOINING_DATE       : 6,
    STATUS             : 7,  // Active / Inactive
    LEAVE_BALANCE      : 8,
    LAST_CREDIT_MONTH  : 9,
    ADDED_AT           : 10,
    UPDATED_AT         : 11,
    UPDATED_BY         : 12,
    NOTES              : 13,
    PHOTO_FILE_ID      : 14,
    PASSPORT_PHOTO_URL : 15,
    PASSPORT_PHOTO_FILE_ID : 16,
    BANK_NAME          : 17,
    BANK_ACCOUNT_NO    : 18,
    IFSC_CODE          : 19,
    PASSBOOK_PHOTO_URL : 20,
    PASSBOOK_PHOTO_FILE_ID : 21,
    UNIFORM_ISSUED_DATE : 22,
    // Legacy aliases (same indexes as above)
    ROLE               : 5,
    ACTIVE             : 7
  },

  GROOMER_ATTENDANCE_COLS: {
    DATE       : 0,
    STAFF_ID   : 1,
    NAME       : 2,
    STATUS     : 3,
    MARKED_AT  : 4,
    MARKED_BY  : 5,
    NOTES      : 6
  },

  GROOMER_LEAVE_COLS: {
    LEAVE_ID      : 0,
    STAFF_ID      : 1,
    START_DATE    : 2,
    END_DATE      : 3,
    REASON        : 4,
    APPLIED_AT    : 5,
    APPLIED_BY    : 6,
    STATUS        : 7,
    DAYS_DEDUCTED : 8
  },

  // ── Horses stable register (0-based) ─────────────────────
  HORSE_COLS: {
    HORSE_ID            : 0,
    NAME                : 1,
    LOCATION            : 2,
    TRAINER             : 3,
    GROOM               : 4,
    STATUS              : 5,
    BREED               : 6,
    AGE                 : 7,
    GENDER              : 8,
    OWNER               : 9,
    WEIGHT_KG           : 10,
    FACILITY_MULTIPLIER : 11,
    LEASE_RIDER         : 12,
    LEASE_DATE          : 13,
    PHOTO_URL           : 14,
    PHOTO_FILE_ID       : 15,
    CHIP_NO             : 16,
    EFI_ID              : 17,
    DOB                 : 18,
    COLOUR              : 19,
    VACCINATION_DATE    : 20,
    DEWORMING_DATE      : 21,
    FARRIER_DATE        : 22,
    VET_NOTES           : 23,
    ADDED_AT            : 24,
    UPDATED_AT          : 25,
    UPDATED_BY          : 26,
    VACCINATION_STATUS  : 27,
    VACCINATION_POSTPONED_TO : 28,
    DEWORMING_STATUS    : 29,
    DEWORMING_POSTPONED_TO : 30,
    FARRIER_STATUS      : 31,
    FARRIER_POSTPONED_TO : 32,
    // Legacy aliases
    KE_HORSE_ID         : 0,
    ASSIGNED_TRAINER    : 3,
    ASSIGNED_GROOMER    : 4
  },

  // Reminder intervals (days after last care date)
  HORSE_HEALTH_INTERVALS: {
    VACCINATION: 365,
    DEWORMING  : 90,
    FARRIER    : 42
  },

  // Extra care activity types beyond Vaccination / Deworming / Farrier (e.g. Shoeing)
  HORSE_CARE_TYPE_COLS: {
    TYPE_ID       : 0,
    LABEL         : 1,
    INTERVAL_DAYS : 2,
    ACTIVE        : 3,
    ADDED_AT      : 4,
    ADDED_BY      : 5
  },

  // Chronological horse activity feed (care, lease, status, notes)
  HORSE_ACTIVITY_COLS: {
    ACTIVITY_ID   : 0,
    HORSE_ID      : 1,
    HORSE_NAME    : 2,
    ACTIVITY_TYPE : 3,
    DATE          : 4,
    TITLE         : 5,
    DETAILS       : 6,
    RECORDED_AT   : 7,
    RECORDED_BY   : 8
  },

  // Per-horse rows for custom care types (one row per horse + type)
  HORSE_CUSTOM_CARE_COLS: {
    ROW_ID        : 0,
    HORSE_ID      : 1,
    TYPE_ID       : 2,
    TYPE_LABEL    : 3,
    LAST_DATE     : 4,
    STATUS        : 5,
    POSTPONED_TO  : 6,
    NOTES         : 7,
    UPDATED_AT    : 8,
    UPDATED_BY    : 9
  },

  // Legacy simple stock layout (kept for migration helpers)
  STOCK_COLS: {
    ITEM_ID    : 0,
    ITEM_NAME  : 1,
    QUANTITY   : 2,
    UNIT       : 3,
    MIN_LEVEL  : 4,
    NOTES      : 5,
    UPDATED_AT : 6,
    UPDATED_BY : 7
  },

  // Feed inventory (location qty + daily consumption)
  // PACK_SIZE_KG = weight of one stock unit (e.g. 50 for a 50 kg bag). 0 = unknown,
  // which disables kg entry for that item.
  FEED_COLS: {
    ITEM_ID           : 0,
    ITEM_NAME         : 1,
    LOCATION          : 2,
    QUANTITY          : 3,
    UNIT              : 4,
    MIN_LEVEL         : 5,
    CONSUMED_PER_DAY  : 6,
    NOTES             : 7,
    UPDATED_AT        : 8,
    UPDATED_BY        : 9,
    PACK_SIZE_KG      : 10,
    // Regular = daily EOD consume + auto-defaults; Occasional = vitamins etc., log use only
    USAGE_MODE        : 11
  },

  FEED_MOVEMENT_COLS: {
    MOVEMENT_ID : 0,
    ITEM_ID     : 1,
    TYPE        : 2,  // Restock | Consume | Adjust | Use | Expire
    DATE        : 3,
    QTY         : 4,
    NOTES       : 5,
    RECORDED_AT : 6,
    RECORDED_BY : 7,
    REF_HORSE_ID: 8,
    REF_HORSE_NAME: 9
  },

  // Tack inventory
  TACK_COLS: {
    ITEM_ID       : 0,
    ITEM_NAME     : 1,
    CATEGORY      : 2,
    MODEL         : 3,
    VENDOR        : 4,
    LOCATION      : 5,
    QUANTITY      : 6,
    MIN_LEVEL     : 7,
    PHOTO_URL     : 8,
    PHOTO_FILE_ID : 9,
    NOTES         : 10,
    UPDATED_AT    : 11,
    UPDATED_BY    : 12
  },

  TACK_MOVEMENT_COLS: {
    MOVEMENT_ID    : 0,
    ITEM_ID        : 1,
    TYPE           : 2,  // Wear Out | Restock | Transfer
    DATE           : 3,
    QTY            : 4,
    REASON         : 5,
    PHOTO_URL      : 6,
    PHOTO_FILE_ID  : 7,
    FROM_LOCATION  : 8,
    TO_LOCATION    : 9,
    RECORDED_AT    : 10,
    RECORDED_BY    : 11
  },

  // ── Payments Ledger columns (0-based) ────────────────────
  LEDGER_COLS: {
    KE_NO        : 0,
    NAME         : 1,
    PHONE        : 2,
    AMOUNT       : 3,
    PAY_DATE     : 4,
    TXN_REF      : 5,
    RECEIPT_NO   : 6,
    SENT_AT      : 7,
    STUDENT      : 8,
    GRADE        : 9,
    SECTION      : 10,
    PAYMENT_TYPE : 11,
    PARENT       : 12
  },

  // ── Service / Pricing sheet columns (0-based) ────────────
  PRICING_COLS: {
    ROW          : 0,
    NAME         : 1,   // SERVICE_COLUMN = 'Service'
    PRICE        : 2,
    DOC_ID       : 3,
    TYPE         : 4
  }
};

// Legacy typo used in some older spreadsheets — tried automatically by getRegistrationSheet_().
var REGISTRATION_SHEET_ALIASES = ['Registration Response', 'Regestration Response'];

/** Find the registration form-response sheet (supports legacy tab names). */
function getRegistrationSheet_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var names = [CONFIG.SHEETS.REGISTRATION_FORM];
  REGISTRATION_SHEET_ALIASES.forEach(function (n) {
    if (names.indexOf(n) === -1) names.push(n);
  });
  for (var i = 0; i < names.length; i++) {
    var sh = ss.getSheetByName(names[i]);
    if (sh) return sh;
  }
  return null;
}

/** True if a tab name is the registration response sheet (any known spelling). */
function isRegistrationSheetName_(name) {
  var n = String(name || '').trim();
  if (n === CONFIG.SHEETS.REGISTRATION_FORM) return true;
  return REGISTRATION_SHEET_ALIASES.indexOf(n) >= 0;
}

// NOTE:
// Time slot booking is intentionally handled via My Rides portal (not at registration).
// Include appsscript.json (oauthScopes: documents, script.send_mail, drive, …) in this project,
// then run menu “Authorize script — Docs + mail (run once)” after any manifest change.

// ============================================================
//  KE / REGISTRATION NUMBER GENERATOR
// ============================================================

function generateKENo() {
  const date   = new Date();
  const year   = date.getFullYear().toString().substr(-2);
  const month  = String(date.getMonth() + 1).padStart(2, '0');
  const day    = String(date.getDate()).padStart(2, '0');
  const random = Math.floor(Math.random() * 9000) + 1000;
  return 'KE' + year + month + day + random;
}

// ============================================================
//  PHONE NORMALISE (last 10 digits)
// ============================================================

function normalisePhone(p) {
  return String(p || '').replace(/\D/g, '').slice(-10);
}

// ============================================================
//  PREFILLED PAYMENT FORM URL
//  Builds a pre-filled Google Form URL for the payment form.
//  Unused entry IDs should be left as '' in PREFILL_ENTRY_IDS.
// ============================================================

function buildPaymentFormUrl(d) {
  // d: { regRef, student, parent, email, phone, grade }
  const base    = CONFIG.PAYMENT_FORM_BASE_URL;
  const entries = CONFIG.PREFILL_ENTRY_IDS;
  const params  = [];
  if (entries.regRef  && d.regRef)  params.push(entries.regRef  + '=' + encodeURIComponent(d.regRef));
  if (entries.student && d.student) params.push(entries.student + '=' + encodeURIComponent(d.student));
  if (entries.parent  && d.parent)  params.push(entries.parent  + '=' + encodeURIComponent(d.parent));
  if (entries.email   && d.email)   params.push(entries.email   + '=' + encodeURIComponent(d.email));
  if (entries.phone   && d.phone)   params.push(entries.phone   + '=' + encodeURIComponent(d.phone));
  if (entries.grade   && d.grade)   params.push(entries.grade   + '=' + encodeURIComponent(d.grade));
  return base + (params.length ? '?' + params.join('&') : '');
}

// ============================================================
//  RIDER LOOKUPS
// ============================================================

function findRiderByPhone(phone) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
  if (!sheet) return null;
  const target = normalisePhone(phone);
  if (!target || target.length < 10) return null;
  const data   = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (normalisePhone(data[i][CONFIG.RIDER_COLS.PHONE]) === target) {
      return { rowIndex: i + 1, row: data[i] };
    }
  }
  return null;
}

function findRiderByPhoneAndName(phone, name) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const sheet      = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
  if (!sheet) return null;
  const target     = normalisePhone(phone);
  const targetName = String(name || '').trim().toLowerCase();
  if (!target || target.length < 10) return null;
  const data       = sheet.getDataRange().getValues();
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
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getSheetByName(CONFIG.SHEETS.RIDERS);
  if (!sheet) return null;
  const target = String(keNo || '').trim().toUpperCase();
  if (!target) return null;
  const data   = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][CONFIG.RIDER_COLS.KE_NO] || '').trim().toUpperCase() === target) {
      return { rowIndex: i + 1, row: data[i] };
    }
  }
  return null;
}

// ============================================================
//  LOCATION / BRANDING HELPERS (Hyderabad campus)
// ============================================================

/** Short location label — "Hyderabad" (or school · city if SCHOOL_NAME is set). */
function schoolLocationShort_() {
  var city = String(CONFIG.LOCATION_CITY || 'Hyderabad').trim();
  var school = String(CONFIG.SCHOOL_NAME || '').trim();
  return school ? (school + ' · ' + city) : city;
}

/** "Hyderabad, Telangana, India" */
function locationCityState_() {
  var city = String(CONFIG.LOCATION_CITY || 'Hyderabad').trim();
  var state = String(CONFIG.LOCATION_STATE || 'Telangana').trim();
  var country = String(CONFIG.LOCATION_COUNTRY || 'India').trim();
  return city + ', ' + state + ', ' + country;
}

/** Full location — city/state/country, optionally prefixed with school name. */
function schoolLocationFull_() {
  var school = String(CONFIG.SCHOOL_NAME || '').trim();
  var loc = locationCityState_();
  return school ? (school + ' · ' + loc) : loc;
}

/** Compact footer used in most HTML emails. */
function emailFooterHtml_() {
  var phone = String(CONFIG.CONTACT_PHONE || '+91-9980895533').trim();
  var email = String(CONFIG.CONTACT_EMAIL || CONFIG.MAIL_FROM || '').trim();
  return '<strong>' + String(CONFIG.BUSINESS_NAME || 'Kings Equestrian Foundation') + '</strong><br>'
    + schoolLocationFull_() + '<br>'
    + phone + (email ? ' | ' + email : '');
}

/** One-line plain-text footer. */
function emailFooterPlain_() {
  var phone = String(CONFIG.CONTACT_PHONE || '+91-9980895533').trim();
  var email = String(CONFIG.CONTACT_EMAIL || CONFIG.MAIL_FROM || '').trim();
  return String(CONFIG.BUSINESS_NAME || 'Kings Equestrian Foundation')
    + ' | ' + schoolLocationFull_()
    + ' | ' + phone
    + (email ? ' | ' + email : '');
}

// ============================================================
//  DATE HELPERS
// ============================================================

function fmtDate(d) {
  if (!d) return '';
  try {
    const dt = (d instanceof Date) ? d : new Date(d);
    if (isNaN(dt.getTime())) return '';
    if (dt.getFullYear() < 2000) return '';
    return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd MMM yyyy');
  } catch (e) { return String(d); }
}

function fmtDateTime(d) {
  if (!d) return '';
  try {
    const dt = (d instanceof Date) ? d : new Date(d);
    if (isNaN(dt.getTime())) return '';
    if (dt.getFullYear() < 2000) return '';
    return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
  } catch (e) { return String(d); }
}

function ymd(d) {
  if (!d) return '';
  try {
    const dt = (d instanceof Date) ? d : new Date(d);
    if (isNaN(dt.getTime())) return '';
    if (dt.getFullYear() < 2000) return '';
    return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) { return ''; }
}

// ============================================================
//  UPI / QR  (retained for payment emails)
// ============================================================

function createUPILink(amount, reference) {
  var exactAmount = Number(amount || 0).toFixed(2);
  return 'upi://pay?pa=' + CONFIG.UPI_ID
    + '&pn=' + encodeURIComponent(CONFIG.BUSINESS_NAME)
    + '&am=' + exactAmount
    + '&cu=INR'
    + '&tn=' + encodeURIComponent(reference);
}

function createQRCode(link) {
  return 'https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=' + encodeURIComponent(link);
}

// ============================================================
//  PRICING / SERVICE DATA
// ============================================================

function getPricingData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PRICING);
  if (!sheet) return {};
  const data  = sheet.getDataRange().getValues();
  const map   = {};
  for (let i = 1; i < data.length; i++) {
    const name  = String(data[i][CONFIG.PRICING_COLS.NAME]  || '').trim();
    const price = data[i][CONFIG.PRICING_COLS.PRICE];
    const docId = String(data[i][CONFIG.PRICING_COLS.DOC_ID]|| '').trim();
    const type  = String(data[i][CONFIG.PRICING_COLS.TYPE]  || 'Regular').trim();
    if (name) map[name] = { price, docId, type };
  }
  return map;
}

// ============================================================
//  CC / ADMIN EMAIL HELPERS
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
  let root = DriveApp.getFoldersByName(CONFIG.DRIVE_ROOT_FOLDER);
  root     = root.hasNext() ? root.next() : DriveApp.createFolder(CONFIG.DRIVE_ROOT_FOLDER);
  let sub  = root.getFoldersByName(CONFIG.RECEIPTS_SUBFOLDER);
  return sub.hasNext() ? sub.next() : root.createFolder(CONFIG.RECEIPTS_SUBFOLDER);
}

function storeReceiptInDrive(blob, studentName, receiptNo) {
  try {
    const folder   = getReceiptsFolder();
    const ts       = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const fileName = 'Receipt_' + receiptNo.replace(/\//g,'-') + '_' + studentName.replace(/\s+/g,'_') + '_' + ts + '.pdf';
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
//  CONSENT PDF ROUTER
//  Summer track (summer / summer program, lowercased in form handler) → farm-style consent PDF
//  Otherwise → Indus school registration consent
// ============================================================

/**
 * PROGRAM is normalised to lowercase in formhandlers — compare lowercase only.
 * Accepts "summer", "summer program", "summer programme".
 */
function _isSummerTrackProgram(program) {
  const p = String(program || '').trim().toLowerCase().replace(/\s+/g, ' ');
  if (p === 'summer') return true;
  if (p === 'summer program' || p === 'summer programme') return true;
  return false;
}

/**
 * Main entry point called from FormHandler / Email.
 *
 * @param {string} program        - 'summer' for farm-style consent; any other value → school consent
 * @param {Object} d              - all registration fields
 * @param {string} d.studentName
 * @param {string} d.parentName
 * @param {string} d.email
 * @param {string} d.phone
 * @param {string} d.grade
 * @param {string} d.dob
 * @param {string} d.address
 * @param {string} d.motherName
 * @param {string} d.fatherName
 * @param {string} d.motherContact
 * @param {string} d.motherWhatsApp
 * @param {string} d.fatherContact
 * @param {string} d.fatherWhatsApp
 * @param {string} d.emergencyContact
 * @param {string} d.relationship
 * @param {Date}   d.consentDate
 * @param {string} d.horseLease   - 'yes' / 'no' / boolean
 * @param {string} d.serviceProgram - e.g. '2 classes per week'
 * @param {string} d.sessionDates
 * @returns {Blob} PDF blob
 */
function generateConsentPDF(program, d) {
  if (_isSummerTrackProgram(program)) {
    // Farm-style consent (name, email, phone, bookingDate) — see _generateSummerConsentPDF
    return _generateSummerConsentPDF(d.studentName || d.parentName, d.email, d.phone, d.consentDate);
  }
  return _generateSchoolConsentPDF(d);
}

// ============================================================
//  SCHOOL CONSENT PDF  (Indus registration form layout)
// ============================================================

function _generateSchoolConsentPDF(d) {
  const LABEL_FONT = 'Comic Sans MS';
  const VALUE_FONT = 'Arial';
  const FONT_SIZE  = 11;

  const doc  = DocumentApp.create('Consent Form - ' + (d.studentName || 'Student'));
  const body = doc.getBody();
  body.clear();
  body.setMarginTop(40);
  body.setMarginBottom(40);
  body.setMarginLeft(40);
  body.setMarginRight(40);

  /* ── helpers ── */
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
    const start  = fullText.indexOf(valStr);
    if (start === -1) return;
    const end = start + valStr.length - 1;
    if (end >= start && start >= 0 && end < fullText.length) {
      textObj.setFontFamily(start, end, VALUE_FONT);
      textObj.setBold(start, end, false);
      textObj.setUnderline(start, end, true);
    }
  }

  function fmtD(dateValue) {
    if (!dateValue) return null;
    if (dateValue instanceof Date) {
      const day   = String(dateValue.getDate()).padStart(2, '0');
      const month = String(dateValue.getMonth() + 1).padStart(2, '0');
      const year  = dateValue.getFullYear();
      return day + '/' + month + '/' + year;
    }
    // Only accept short, date-like strings. This guards against a mismapped
    // form column (e.g. a consent checkbox) landing in the Date field —
    // anything that isn't a real date leaves the line blank for manual signing.
    const s = String(dateValue).trim();
    if (!s || s.length > 24) return null;
    const parsed = new Date(s);
    if (isNaN(parsed.getTime())) return null;
    const day   = String(parsed.getDate()).padStart(2, '0');
    const month = String(parsed.getMonth() + 1).padStart(2, '0');
    const year  = parsed.getFullYear();
    return day + '/' + month + '/' + year;
  }

  function sp(value, blank) {
    blank = blank || '____________________';
    return value ? ('  ' + value + '  ') : blank;
  }

  /* ── fetch logos ── */
  const leftImageUrl  = 'https://iais.in/wp-content/uploads/2025/11/Indus-Altum-International-School-.png';
  const rightImageUrl = CONFIG.LOGO_URL;
  let leftBlob, rightBlob;
  try { leftBlob  = UrlFetchApp.fetch(leftImageUrl).getBlob();  } catch (e) { Logger.log('Left logo error: ' + e); }
  try { rightBlob = UrlFetchApp.fetch(rightImageUrl).getBlob(); } catch (e) { Logger.log('Right logo error: ' + e); }

  /* ── header table ── */
  const headerTable = body.appendTable();
  const headerRow   = headerTable.appendTableRow();

  const leftCell = headerRow.appendTableCell();
  leftCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
  leftCell.setPaddingRight(15);
  leftCell.setWidth(120);
  if (leftBlob) {
    const img = leftCell.appendParagraph('').appendInlineImage(leftBlob);
    img.setWidth(110);
    img.setHeight(110);
  }

  const centerCell = headerRow.appendTableCell();
  centerCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);

  let cp = centerCell.appendParagraph('INDUS EQUESTRIAN CENTRE OF EXCELLENCE');
  cp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  cp.editAsText().setFontFamily(LABEL_FONT).setFontSize(14).setBold(true);
  cp.setSpacingAfter(4);

  cp = centerCell.appendParagraph('KINGS EQUESTRIAN FOUNDATION HORSE RIDING');
  cp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  cp.editAsText().setFontFamily(LABEL_FONT).setFontSize(12).setBold(true);
  cp.setSpacingAfter(2);

  cp = centerCell.appendParagraph(schoolLocationShort_());
  cp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  cp.editAsText().setFontFamily(LABEL_FONT).setFontSize(10).setForegroundColor('#555555');
  cp.setSpacingAfter(4);

  cp = centerCell.appendParagraph('REGISTRATION (' + CONFIG.ACADEMIC_YEAR_LABEL + ')');
  cp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  cp.editAsText().setFontFamily(LABEL_FONT).setFontSize(11).setBold(true);
  cp.setSpacingAfter(0);

  const rightCell = headerRow.appendTableCell();
  rightCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
  rightCell.setPaddingLeft(15);
  rightCell.setWidth(120);
  if (rightBlob) {
    const img = rightCell.appendParagraph('').appendInlineImage(rightBlob);
    img.setWidth(110);
    img.setHeight(110);
  }

  headerTable.setBorderWidth(0);
  body.appendParagraph('').setSpacingAfter(12);

  paragraph('Dear Equestrian Team,', 11, false, 8);

  /* ── enrolment line ── */
  let p = body.appendParagraph('');
  let t = p.editAsText();
  const nameVal    = sp(d.studentName, '____________________________________________');
  const enrollText = 'Please enrol ' + nameVal + ' (Name) in the horse-riding program.';
  t.setText(enrollText).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
  if (d.studentName) formatValue(t, enrollText, nameVal);
  p.setSpacingAfter(8);

  const sessionDatesText = d.sessionDates || '15th July ' + CONFIG.ACADEMIC_YEAR_LABEL.split('-')[0] + ' – April ' + CONFIG.ACADEMIC_YEAR_LABEL.split('-')[1];
  paragraph('Session: As per School Academic Year (' + sessionDatesText + ')', 11, true, 10);


  /* ── rider information ── */
  paragraph("RIDER'S INFORMATION", 11, false, 8);

  p = body.appendParagraph('');
  t = p.editAsText();
  const rNameSp = sp(d.studentName, '____________________');
  const gradeSp = sp(d.grade, '__________________');
  const dobSp   = sp(d.dob ? fmtD(d.dob) : '', '__________________');
  const riderLine = "Rider's Name: " + rNameSp + '     \nDate of Birth: ' + dobSp + '     Grade & Section: ' + gradeSp;
  t.setText(riderLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
  if (d.studentName) formatValue(t, riderLine, rNameSp);
  if (d.dob) formatValue(t, riderLine, dobSp);
  if (d.grade) formatValue(t, riderLine, gradeSp);
  p.setSpacingAfter(8);

  p = body.appendParagraph('');
  t = p.editAsText();
  const addrSp   = sp(d.address, '_______________________________________________________________________________');
  const addrLine = 'Address: ' + addrSp;
  t.setText(addrLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
  if (d.address) formatValue(t, addrLine, addrSp);
  p.setSpacingAfter(8);

  p = body.appendParagraph('');
  t = p.editAsText();
  const parentNameSp = sp(d.parentName, '____________________________________________');
  const parentOnlyLine = "Parent's Name: " + parentNameSp;
  t.setText(parentOnlyLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
  if (d.parentName) formatValue(t, parentOnlyLine, parentNameSp);
  p.setSpacingAfter(8);

  paragraph("Parent's contact:", 11, false, 2);

  p = body.appendParagraph('');
  t = p.editAsText();
  const parentPhSp = sp(d.phone, '____________________');
  const parentPhoneLine = 'Phone / WhatsApp: ' + parentPhSp;
  t.setText(parentPhoneLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
  if (d.phone) formatValue(t, parentPhoneLine, parentPhSp);
  p.setSpacingAfter(8);

  p = body.appendParagraph('');
  t = p.editAsText();
  const emailSp   = sp(d.email, '_______________________________________________________________________________');
  const emailLine = 'Email: ' + emailSp;
  t.setText(emailLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
  if (d.email) formatValue(t, emailLine, emailSp);
  p.setSpacingAfter(8);

  p = body.appendParagraph('');
  t = p.editAsText();
  const emergSp   = sp(d.phone, '____________________');
  const emergLine = 'In Case of Emergency Call: ' + emergSp + ' (Phone Number)';
  t.setText(emergLine).setFontFamily(LABEL_FONT).setFontSize(FONT_SIZE);
  if (d.emergencyContact) formatValue(t, emergLine, emergSp);
  p.setSpacingAfter(17);

  /* ── consent section ── */
  paragraph('ACKNOWLEDGEMENT / CONSENT FORM – HORSE RIDING PARTICIPANTS', 11, true, 8);

  const consentParas = [
    'Kings Equestrian (' + (CONFIG.LOCATION_CITY || 'Hyderabad') + ') offers horseback riding programs for those interested in Casual riding, Dressage, Jumping and related workshops and clinics. Programs of this sort involve risk of personal injury, including, but not limited to, bruises, broken bones, head injuries and death. All normal safety precautions are taken to protect our participants, but occasionally accidents do happen.',
    'Horses may, without warning or apparent cause, buck, rear, stumble, fall, spook or make unanticipated movements, jump obstacles in their path, bite, kick, step on a person\'s foot, or push or shove a person, and saddles or bridles may loosen or break, all of which may result in injury.',
    'We further note that if you are pregnant or immunocompromised you may be at a greater risk of injury and/or of contracting possible zoonotic agents due to your close proximity to animals and should consult your healthcare provider before undertaking equestrian activities.',
    'The school does not provide insurance to program participants. We strongly suggest that you should be covered under your own private insurance plan.',
    'All fees/funds will be used for animal welfare and well-being including feeding the horse, maintenance of the stable, and horse upkeep. Fees will not be refunded after enrolment.',
    'Kings Equestrian will not be responsible for any cancellations due to unpredicted weather, school leaves, school camps, student health, etc.',
    'Missed classes can only be accommodated in that particular week/weekend and will automatically lapse.',
    'Helmet, body protector, and shoes are mandatory for horse riding. Students and parents are responsible for having the required riding gear.',
    'Competition-related travel and expenses are not included and will be paid separately.'
  ];
  consentParas.forEach(function(txt) { paragraph(txt, 11, false, 8); });

  /* ── consent paragraph with filled-in values ── */
  p = body.appendParagraph('');
  t = p.editAsText();
  const cNameSp  = sp(d.studentName, '____________________');
  const relSp    = sp(d.relationship, 'ward');
  const cText    = 'I give my consent for ' + cNameSp + ', my ' + relSp + ' ("RIDER"), to participate in the above-mentioned riding programs and/or workshop. I have read the information provided above and understand the inherent risks involved. I further attest that I am at least eighteen (18) years of age and fully authorized to sign this consent.';
  t.setText(cText).setFontFamily(LABEL_FONT).setFontSize(11);
  if (d.studentName) formatValue(t, cText, cNameSp);
  if (d.relationship) formatValue(t, cText, relSp);
  p.setSpacingAfter(10);

  /* ── signature section ── */
  p = body.appendParagraph('');
  t = p.editAsText();
  const signerName    = d.parentName || d.studentName || '';
  const pNameSp       = sp(signerName, '____________________');
  const pNameLine     = 'Name (Parent): ' + pNameSp;
  t.setText(pNameLine).setFontFamily(LABEL_FONT).setFontSize(11);
  if (signerName) formatValue(t, pNameLine, pNameSp);
  p.setSpacingAfter(10);

  p = body.appendParagraph('');
  t = p.editAsText();
  const sigSp         = sp(signerName, '____________________');
  // Date on the form = the moment this consent is generated / emailed.
  const conDateFmt    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const dateSp        = sp(conDateFmt, '____________________');
  const sigLine       = 'Signature: ' + sigSp + '     Date: ' + dateSp;
  t.setText(sigLine).setFontFamily(LABEL_FONT).setFontSize(18);

  if (signerName) {
    const sigStart = sigLine.indexOf(sigSp);
    if (sigStart !== -1) {
      const sigEnd = sigStart + sigSp.length - 1;
      if (sigEnd >= sigStart && sigStart >= 0) {
        t.setFontFamily(sigStart, sigEnd, 'Dancing Script');
        t.setBold(sigStart, sigEnd, false);
        t.setUnderline(sigStart, sigEnd, false);
      }
    }
  }
  if (conDateFmt) formatValue(t, sigLine, dateSp);
  p.setSpacingAfter(20);

  /* ── save / export ── */
  doc.saveAndClose();
  const pdf = doc.getAs('application/pdf');
  pdf.setName('Consent_Form_' + (d.studentName || 'Student').replace(/\s+/g, '_') + '.pdf');
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  return pdf;
}

// ============================================================
//  SUMMER CONSENT PDF — same body as farm `generateConsentPDF(name,email,phone,bookingDate)`
//  Called from generateConsentPDF(program, d) when _isSummerTrackProgram(program)
// ============================================================

function _generateSummerConsentPDF(name, email, phone, bookingDate) {
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
    size = size || FONT_SIZE;
    bold = bold || false;
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
    if (!(dateValue instanceof Date)) {
      // Reject non-date text (e.g. a mismapped consent checkbox column).
      const s = String(dateValue).trim();
      if (!s || s.length > 24) return null;
      const parsed = new Date(s);
      if (isNaN(parsed.getTime())) return null;
      dateValue = parsed;
    }
    const day = String(dateValue.getDate()).padStart(2, '0');
    const month = String(dateValue.getMonth() + 1).padStart(2, '0');
    const year = dateValue.getFullYear();
    return day + '/' + month + '/' + year;
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
  // Date on the form = the moment this consent is generated / emailed.
  const dateFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
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

  const footerPara = body.appendParagraph(emailFooterPlain_());
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
    .createMenu('Indus Equestrian')
    .addItem('🚀 Setup All Sheets/Tabs (new project)', 'setupAllSheets')
    .addSeparator()
    .addItem('Authorize script — Docs + mail (run once)', 'authorizeIndusScopesOnce')
    .addItem('Test Email (Brevo / fallback)', 'testBrevoEmail')
    .addItem('Check Email Quota / Provider', 'checkEmailQuota')
    .addItem('Test Consent PDF (attach check)', 'testConsentPdfGeneration')
    .addItem('Resend Welcome Email',      'resendWelcomeEmail')
    .addItem('Send Payment Receipt',      'sendPaymentReceiptMenu')
    .addSeparator()
    .addItem('Create Backup Now',         'backupSpreadsheetNow')
    .addItem('Backup / Restore Help',     'showBackupRestoreHelp')
    .addSeparator()
    .addItem('Send Daily Summary Now',    'testSendDailySummaryNow')
    .addItem('Dry-Run Daily Summary',     'testDailySummaryDryRun')
    .addItem('Send Shop Orders Daily Report Now', 'testSendDailyShopOrdersReportNow')
    .addSeparator()
    .addItem('Seed Demo Attendance Data (Today/Tomorrow)', 'seedDemoAttendanceData')
    .addSeparator()
    .addItem('Sync Grade & Section → STUDENTS', 'syncGradeSectionFromRegistrationToStudents')
    .addSeparator()
    .addItem('Setup Trainers Sheet (seed logins)', 'setupTrainersSheet')
    .addSeparator()
    .addItem('Enable Cache Warming (fast loads)', 'installCacheWarmTrigger')
    .addItem('Archive Old Schedule Rows…', 'archiveOldScheduleMenu')
    .addSeparator()
    .addItem('Setup Training System Tabs', 'setupTrainingSystem_2627')
    .addItem('Install Training Triggers',  'installTrainingTriggers_2627')
    .addItem('Remove Training Triggers',   'removeTrainingTriggers_2627')
    .addItem('Send Weekly Training Report Now', 'sendWeeklyTrainingReport_2627')
    .addItem('Send Weekly Ops Summary Now', 'sendWeeklyOpsAdminSummary')
    .addItem('Create Training Forms', 'createTrainingForms_2627')
    .addItem('Send Pending Level Certificates', 'sendPendingLevelCertificates')
    .addItem('Reset Certificate History', 'resetLevelCertificateHistory')
    .addSeparator()
    .addItem('Setup All Triggers',        'setupTriggers')
    .addItem('Diagnose Payment Columns',  'diagnosPaymentFormColumns')
    .addItem('Email Send Report',         'showEmailSendReport')
    .addItem('Retry Failed Emails',       'retryFailedEmails')
    .addToUi();
}

function setupTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onRegistrationFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('onPaymentFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('sendDailyAdminSummary').timeBased().everyDays(1).atHour(20).create();
  ScriptApp.newTrigger('sendDailyShopOrdersReport').timeBased().everyDays(1).atHour(21).create();
  ScriptApp.newTrigger('sendWeeklyOpsAdminSummary').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  ScriptApp.newTrigger('createDailyDriveBackup').timeBased().everyDays(1).atHour(23).create();
  ScriptApp.newTrigger('onTrainingFormSubmit_2627').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('sendWeeklyTrainingReport_2627').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  ScriptApp.newTrigger('warmAttendanceCaches').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('archiveOldScheduleAuto').timeBased().everyDays(30).atHour(1).create();
  SpreadsheetApp.getUi().alert(
    'Triggers set!\n\n'
    + '- Registration form: welcome email + KE No\n'
    + '- Payment form: receipt email\n'
    + '- Daily summary: evening admin report\n'
    + '- Daily shop orders report: 9 PM (counts + Excel)\n'
    + '- Weekly ops summary: Monday 7am (staff/horses/feed/tack)\n'
    + '- Daily backup: Drive copy at 11 PM'
  );
}