// ============================================================
// KINGS EQUESTRIAN — ONE-CLICK PROJECT SETUP
// File: SetupProject.gs
//
// Creates every tab the system needs in a brand-new spreadsheet
// (new Google account), each with the correct header row.
//
// Idempotent & non-destructive:
//   • A sheet that already exists is left in place.
//   • Headers are only written when the sheet is empty.
//   • Existing data is never cleared or overwritten.
//
// Run from menu: Indus Equestrian → "🚀 Setup All Sheets/Tabs".
// ============================================================

var SETUP_HEADER_BG = '#1f4617';   // brand green (horseshoe)
var SETUP_HEADER_FG = '#ffffff';

/**
 * Build the header row for the Schedule sheet directly from SCHED_COLS
 * so it always stays in sync with the column map in config.gs.
 */
function _setupScheduleHeaders_() {
  var c = CONFIG.SCHED_COLS;
  var h = [];
  h[c.KE_NO]        = 'KE No';
  h[c.NAME]         = 'Name';
  h[c.PHONE]        = 'Phone';
  h[c.EMAIL]        = 'Email';
  h[c.SERVICE]      = 'Service';
  h[c.DATE]         = 'Date';
  h[c.TIME_SLOT]    = 'Time Slot';
  h[c.PARTICIPANTS] = 'Participants';
  h[c.STATUS]       = 'Status';
  h[c.ATTENDANCE]   = 'Attendance';
  h[c.STAFF_NOTES]  = 'Staff Notes';
  h[c.CAL_EVENT_ID] = 'Calendar Event ID';
  h[c.SOURCE]       = 'Source';
  h[c.BOOKED_BY]    = 'Booked By';
  h[c.SCORED_BY]    = 'Scored By';
  for (var i = 0; i < h.length; i++) if (!h[i]) h[i] = 'Col ' + i;
  return h;
}

function _setupRegistrationHeaders_() {
  var c = CONFIG.REG_COLS;
  var h = [];
  h[c.TIMESTAMP]     = 'Timestamp';
  h[c.STUDENT]       = 'Student Name';
  h[c.GRADE]         = 'Grade';
  h[c.SECTION]       = 'Section';
  h[c.PARENT]        = 'Parent / Guardian Name';
  h[c.PHONE]         = 'Phone';
  h[c.EMAIL]         = 'Email';
  h[c.PROGRAM]       = 'Program';
  if (c.CONSENT_DATE >= 0) h[c.CONSENT_DATE] = 'Consent Date';
  if (c.ADDRESS >= 0)      h[c.ADDRESS]      = 'Address';
  h[c.WELCOME_SENT]  = 'Welcome Email Sent';
  h[c.WELCOME_AT]    = 'Welcome Sent At';
  h[c.KE_NO]         = 'KE No';
  h[c.REG_REF]       = 'Registration Ref';
  h[c.PROGRAM_TRACK] = 'Program Track';
  for (var i = 0; i < h.length; i++) if (!h[i]) h[i] = 'Col ' + i;
  return h;
}

function _setupPaymentHeaders_() {
  // Matches the linked "Payment Responses 26-27" form column order.
  return [
    'Timestamp', 'Registration No', 'Phone number', 'Amount Paid (₹)', 'ScreenShot',
    'Payment Date', 'Transcation Reference Number', 'Pan / AAdhar Number', 'Mode of Payment',
    'Payment For', 'Receipt Sent', 'Receipt Sent At', 'Receipt No', 'Receipt Link'
  ];
}

/**
 * Ensure a sheet exists and — only if it is empty — write & style a header
 * row (plus optional seed rows). Never touches a sheet that already has data.
 * Returns 'created' | 'existing'.
 */
function _setupEnsureSheet_(ss, name, headers, seedRows) {
  var sheet = ss.getSheetByName(name);
  var created = false;
  if (!sheet) { sheet = ss.insertSheet(name); created = true; }

  if (sheet.getLastRow() === 0 && headers && headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground(SETUP_HEADER_BG).setFontColor(SETUP_HEADER_FG).setFontWeight('bold');
    sheet.setFrozenRows(1);
    if (seedRows && seedRows.length) {
      sheet.getRange(2, 1, seedRows.length, seedRows[0].length).setValues(seedRows);
    }
  }
  return created ? 'created' : 'existing';
}

/**
 * MAIN ENTRY — build all tabs for a fresh project.
 * Wired to the menu; shows a summary alert at the end.
 */
function setupAllSheets() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Set up all tabs?',
    'This creates every sheet the Kings Equestrian system needs (with headers) '
      + 'in this spreadsheet.\n\n'
      + 'It is safe to run on an existing project: sheets that already exist are '
      + 'left untouched — no data is cleared.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var summary = _runSetupAllSheets_();

  ui.alert(
    'Setup complete ✅\n\n'
      + 'Created: ' + summary.created.length + ' tab(s)\n'
      + (summary.created.length ? '  • ' + summary.created.join('\n  • ') + '\n\n' : '\n')
      + 'Already present: ' + summary.existing.length + ' tab(s)\n\n'
      + 'Trainer logins seeded: ' + summary.trainers.added + ' added / '
      + summary.trainers.refreshed + ' refreshed\n'
      + '  (aarti / aarti@123 · rahul / rahul@123 · sana / sana@123)\n\n'
      + 'Next steps:\n'
      + '  1. Link your Google Forms to "' + CONFIG.SHEETS.REGISTRATION_FORM + '" and "'
      + CONFIG.SHEETS.PAYMENT_FORM + '" (or paste responses there).\n'
      + '  2. Add curriculum rows in CURRICULUM (or run "Seed Demo Attendance Data").\n'
      + '  3. Add admin emails in "' + CONFIG.SHEETS.MAIL_INFO + '".\n'
      + '  4. Run "Setup All Triggers".'
  );
}

/**
 * Core setup logic (no UI) — usable from scripts/tests too.
 * Returns { created:[], existing:[], trainers:{added,refreshed} }.
 */
function _runSetupAllSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var created = [], existing = [];

  function track(name, status) {
    if (status === 'created') created.push(name); else existing.push(name);
  }

  // ── Core operational sheets ──────────────────────────────
  track(CONFIG.SHEETS.RIDERS, _setupEnsureSheet_(ss, CONFIG.SHEETS.RIDERS,
    ['KE No', 'Name', 'Email', 'Phone', 'Services', 'Participants', 'Registered', 'Notes']));

  track(CONFIG.SHEETS.SCHEDULE, _setupEnsureSheet_(ss, CONFIG.SHEETS.SCHEDULE,
    _setupScheduleHeaders_()));

  track(CONFIG.SHEETS.PAYMENTS, _setupEnsureSheet_(ss, CONFIG.SHEETS.PAYMENTS,
    ['KE No', 'Name', 'Phone', 'Amount', 'Payment Date', 'Txn Ref', 'Receipt No', 'Sent At',
      'Student Name', 'Grade', 'Section', 'Payment Type', 'Parent Name']));

  track(CONFIG.SHEETS.SHOP_PRODUCTS, _setupEnsureSheet_(ss, CONFIG.SHEETS.SHOP_PRODUCTS,
    ['Product ID', 'Product', 'Category', 'Option / Tier', 'Allowed Sizes', 'Price',
      'Image URL', 'Active', 'Sort Order'],
    [
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
    ]));

  track(CONFIG.SHEETS.SHOP_ORDERS, _setupEnsureSheet_(ss, CONFIG.SHEETS.SHOP_ORDERS,
    ['Created At', 'Order ID', 'KE No', 'Rider Name', 'Email', 'Phone', 'Items JSON',
      'Total', 'Payment Status', 'Payment Verified By', 'Payment Verified At',
      'Order Status', 'Expected Delivery Date', 'Delivered At', 'Delivery Notes',
      'Parent Confirmation', 'Parent Confirmed At', 'Parent Note', 'Updated At',
      'Updated By', 'Payment Form URL', 'UPI Reference', 'Status Email At',
      'Client Request ID', 'Payment Transaction Ref']));

  track(CONFIG.SHEETS.GROOMERS, _setupEnsureSheet_(ss, CONFIG.SHEETS.GROOMERS,
    ['Employee ID', 'Name', 'Photo URL', 'Aadhaar Number', 'Mobile Number',
     'Designation', 'Joining Date', 'Status', 'Leave Balance', 'Last Credit Month',
     'Added At', 'Updated At', 'Updated By', 'Notes',
     'Photo_File_ID', 'Passport_Photo_URL', 'Passport_Photo_File_ID',
     'Bank_Name', 'Bank_Account_No', 'IFSC_Code',
     'Passbook_Photo_URL', 'Passbook_Photo_File_ID']));

  track(CONFIG.SHEETS.GROOMER_ATTENDANCE, _setupEnsureSheet_(ss, CONFIG.SHEETS.GROOMER_ATTENDANCE,
    ['Date', 'Staff ID', 'Name Snapshot', 'Status', 'Marked At', 'Marked By', 'Notes']));

  track(CONFIG.SHEETS.GROOMER_LEAVES, _setupEnsureSheet_(ss, CONFIG.SHEETS.GROOMER_LEAVES,
    ['Leave ID', 'Staff ID', 'Start Date', 'End Date', 'Reason',
     'Applied At', 'Applied By', 'Status', 'Days Deducted']));

  track(CONFIG.SHEETS.HORSES, _setupEnsureSheet_(ss, CONFIG.SHEETS.HORSES,
    ['Horse_ID', 'Horse_Name', 'Location', 'Trainer', 'Groom', 'Status', 'Breed', 'Age',
     'Gender', 'Owner', 'Weight_Kg', 'Facility_Multiplier', 'Lease_Rider', 'Lease_Date',
     'Photo_URL', 'Photo_File_ID', 'Chip_No', 'EFI_ID', 'Date of Birth', 'Colour',
     'Vaccination Date', 'Deworming Date', 'Farrier Date', 'Vet Notes',
     'Added At', 'Updated At', 'Updated By']));

  track(CONFIG.SHEETS.FEED_STOCK, _setupEnsureSheet_(ss, CONFIG.SHEETS.FEED_STOCK,
    ['Item ID', 'Item Name', 'Quantity', 'Unit', 'Min Level', 'Notes', 'Updated At', 'Updated By'],
    [
      ['FEED-HAY', 'Hay', 40, 'bales', 10, '', '', ''],
      ['FEED-PELLET', 'Pellets', 25, 'bags', 8, '', '', ''],
      ['FEED-BRAN', 'Bran', 12, 'bags', 5, '', '', ''],
      ['FEED-SUPP', 'Supplements', 8, 'tubs', 3, '', '', '']
    ]));

  track(CONFIG.SHEETS.TACK_STOCK, _setupEnsureSheet_(ss, CONFIG.SHEETS.TACK_STOCK,
    ['Item ID', 'Item Name', 'Quantity', 'Unit', 'Min Level', 'Notes', 'Updated At', 'Updated By'],
    [
      ['TACK-SADDLE', 'Saddles', 6, 'pcs', 2, '', '', ''],
      ['TACK-BRIDLE', 'Bridles', 8, 'pcs', 3, '', '', ''],
      ['TACK-GIRTH', 'Girths', 10, 'pcs', 3, '', '', ''],
      ['TACK-HELMET', 'School Helmets', 12, 'pcs', 4, '', '', ''],
      ['TACK-BOOTS', 'Horse Boots', 10, 'pairs', 3, '', '', '']
    ]));

  // Service / pricing — seed the standard school service so booking works.
  track(CONFIG.SHEETS.PRICING, _setupEnsureSheet_(ss, CONFIG.SHEETS.PRICING,
    ['Row', 'Service', 'Price', 'Doc ID', 'Type'],
    [[1, 'Regular School Classes (2026-27)', 0, '', 'Regular']]));

  track(CONFIG.SHEETS.MAIL_INFO, _setupEnsureSheet_(ss, CONFIG.SHEETS.MAIL_INFO,
    ['Email', 'Type (admin / daily / welcome / receipt)']));

  // ── Form-response landing sheets (placeholders until forms linked) ──
  track(CONFIG.SHEETS.REGISTRATION_FORM, _setupEnsureSheet_(ss, CONFIG.SHEETS.REGISTRATION_FORM,
    _setupRegistrationHeaders_()));

  track(CONFIG.SHEETS.PAYMENT_FORM, _setupEnsureSheet_(ss, CONFIG.SHEETS.PAYMENT_FORM,
    _setupPaymentHeaders_()));

  // ── Email log (reuse existing creator for identical structure) ──
  var hadEmailLog = !!ss.getSheetByName(CONFIG.SHEETS.EMAIL_LOG);
  try { _getOrCreateEmailLogSheet(); } catch (e) { Logger.log('email log setup: ' + e); }
  track(CONFIG.SHEETS.EMAIL_LOG, hadEmailLog ? 'existing' : 'created');

  // ── Training system sheets (STUDENTS, CURRICULUM, BOOKINGS, …) ──
  var trainingNames = [];
  for (var k in TRAINING_CFG.SHEETS) trainingNames.push(TRAINING_CFG.SHEETS[k]);
  var trainingBefore = {};
  trainingNames.forEach(function (n) { trainingBefore[n] = !!ss.getSheetByName(n); });
  try { setupTrainingSystem_2627(); } catch (e) { Logger.log('training setup: ' + e); }
  trainingNames.forEach(function (n) { track(n, trainingBefore[n] ? 'existing' : 'created'); });

  // ── TRAINERS sheet + seed logins ──
  var hadTrainers = !!ss.getSheetByName(CONFIG.SHEETS.TRAINERS);
  var trainers = { added: 0, refreshed: 0 };
  try { trainers = _seedTrainerLogins_(); } catch (e) { Logger.log('trainer seed: ' + e); }
  track(CONFIG.SHEETS.TRAINERS, hadTrainers ? 'existing' : 'created');
  try { _ensureTrainerRoleColumn_(); } catch (e) { Logger.log('trainer role column: ' + e); }

  // Make sure the Schedule audit columns exist (no-op on a fresh build).
  try { _ensureScheduleAuditColumns_(); } catch (e) { Logger.log('audit cols: ' + e); }

  return { created: created, existing: existing, trainers: trainers };
}
