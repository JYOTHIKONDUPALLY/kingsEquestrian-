// ============================================================
// KINGS EQUESTRIAN - EXPENSE & PAYMENT MANAGEMENT SYSTEM
// Google Apps Script Web App  |  Code.gs
// ============================================================

const SHEET_CONFIG = {
  VENDOR_RESPONSES  : "Vendor Registration",
  EMPLOYEE_RESPONSES: "Employee Registration",
  ACCESS_CONTROL    : "Access Control",
  REQUESTS          : "Requests",
  APPROVAL_LOGS     : "ApprovalLog",
  PAYMENT_LOGS      : "PaymentsLog",
};

// ── Column indexes (0-based) ─────────────────────────────────
const VENDOR_COLS = {
  TIMESTAMP: 0, NAME: 1, ADDRESS: 2, EXPENSE_TYPE: 3,
  LOCATION: 4, PAN_GST: 5, BANK_NAME: 6, BENEFICIARY: 7,
  ACCOUNT_NO: 8, IFSC: 9, REG_NO: 10,
};

const EMP_COLS = {
  TIMESTAMP: 0, NAME: 1, DESIGNATION:2,LOCATION: 3,
  PAN: 4, AADHAR:5,BANK_NAME: 6, BENEFICIARY: 7,
  ACCOUNT_NO: 8, IFSC: 9, REG_NO: 10,
};

const ACCESS_COLS = { NAME: 0, EMAIL: 1, ROLE: 2, ADDED_ON: 3 };

const REQ_COLS = {
  ID: 0, TIMESTAMP: 1, REQUESTOR: 2, CONTACT: 3, EMAIL: 4,
  DEPT: 5, EXPENSE_TYPE: 6, CATEGORY: 7, DESCRIPTION: 8,
  AMOUNT: 9, DATE: 10, PAYEE_NAME: 11, PAYEE_TYPE: 12,
  UPI: 13, REF_ID: 14, INVOICE_URL: 15, ADVANCE_REASON: 16,
  EXPECTED_INVOICE_DATE: 17, STATUS: 18, CREATED_AT: 19,LOCATION:20
};

const APPROVAL_COLS = {
  LOG_ID: 0, REQUEST_ID: 1, ACTION: 2, BY_NAME: 3,
  BY_EMAIL: 4, REMARKS: 5, TIMESTAMP: 6,
};

// UPDATED: added PAYMENT_AMOUNT and MODE columns
const PAYMENT_COLS = {
  LOG_ID: 0, REQUEST_ID: 1, UTR: 2, PAYMENT_AMOUNT: 3,
  PAID_BY_NAME: 4, PAID_BY_EMAIL: 5, PAYMENT_DATE: 6,
  MODE: 7, REMARKS: 8, TIMESTAMP: 9,
};

// ─────────────────────────────────────────────────────────────
// ENTRY POINT
// ─────────────────────────────────────────────────────────────
function doGet(e) {
  initSheets();
  const tmpl = HtmlService.createTemplateFromFile("Index");
  return tmpl.evaluate()
    .setTitle("Kings Equestrian – Expense Management")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─────────────────────────────────────────────────────────────
// SHEET INITIALIZATION
// ─────────────────────────────────────────────────────────────
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheet(ss, SHEET_CONFIG.REQUESTS, [
    "Request ID", "Timestamp", "Requestor Name", "Contact", "Email",
    "Department", "Expense Type", "Expense Category", "Description",
    "Amount (₹)", "Expense Date", "Payee Name", "Payee Type",
    "UPI ID", "Reference Expense ID", "Invoice URL",
    "Advance Reason", "Expected Invoice Date", "Status", "Created At"
  ]);

  ensureSheet(ss, SHEET_CONFIG.APPROVAL_LOGS, [
    "Log ID", "Request ID", "Action", "By Name", "By Email",
    "Remarks", "Timestamp"
  ]);

  // UPDATED: Payment logs now include Payment Amount and Mode
  ensureSheet(ss, SHEET_CONFIG.PAYMENT_LOGS, [
    "Log ID", "Request ID", "UTR Number", "Payment Amount (₹)",
    "Paid By Name", "Paid By Email", "Payment Date",
    "Mode", "Remarks", "Timestamp"
  ]);

  generateVendorRegNos(ss);
  generateEmployeeRegNos(ss);
}

function ensureSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#0D1F0F")
      .setFontColor("#D4AF5A");
  }
  return sh;
}

// ─────────────────────────────────────────────────────────────
// REGISTRATION NUMBER GENERATORS
// ─────────────────────────────────────────────────────────────
function generateVendorRegNos(ss) {
  const sh = ss.getSheetByName(SHEET_CONFIG.VENDOR_RESPONSES);
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (!data[i][VENDOR_COLS.REG_NO]) {
      sh.getRange(i + 1, VENDOR_COLS.REG_NO + 1)
        .setValue("KE-VND-" + String(i).padStart(4, "0"));
    }
  }
}

function generateEmployeeRegNos(ss) {
  const sh = ss.getSheetByName(SHEET_CONFIG.EMPLOYEE_RESPONSES);
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (!data[i][EMP_COLS.REG_NO]) {
      sh.getRange(i + 1, EMP_COLS.REG_NO + 1)
        .setValue("KE-EMP-" + String(i).padStart(4, "0"));
    }
  }
}

function onFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  generateVendorRegNos(ss);
  generateEmployeeRegNos(ss);
}

// ─────────────────────────────────────────────────────────────
// USER / ACCESS
// ─────────────────────────────────────────────────────────────
function getUserInfo() {
  // getActiveUser() returns empty string when deployed as "Execute as: Me"
  // and the user hasn't explicitly granted permission yet, OR when accessed
  // anonymously. We fall back to getEffectiveUser() which always works.
  let email = "";
  try { email = Session.getActiveUser().getEmail() || ""; } catch(e) {}
  if (!email) {
    try { email = Session.getEffectiveUser().getEmail() || ""; } catch(e) {}
  }
  // If still empty, the app is deployed as "Execute as: Me" — that mode
  // never exposes the viewer's email. Redeploy as "Execute as: User accessing".
  // We still return a safe object so the UI loads instead of crashing.
  if (!email) {
    return { email: "", role: "viewer", name: "Guest", _authError: true };
  }
  return {
    email,
    role : getUserRole(email),
    name : getUserName(email),
    _authError: false,
  };
}

function getUserRole(email) {
  if (!email) return "viewer";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_CONFIG.ACCESS_CONTROL);
  if (!sh) return "viewer";
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][ACCESS_COLS.EMAIL] || "").toLowerCase().trim() ===
        email.toLowerCase().trim()) {
      return (data[i][ACCESS_COLS.ROLE] || "viewer").toLowerCase().trim();
    }
  }
  return "viewer";
}

function getUserName(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_CONFIG.ACCESS_CONTROL);
  if (sh) {
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if ((data[i][ACCESS_COLS.EMAIL] || "").toLowerCase().trim() ===
          email.toLowerCase().trim()) {
        return data[i][ACCESS_COLS.NAME] || email;
      }
    }
  }
  return email;
}

// ─────────────────────────────────────────────────────────────
// DATA GETTERS
// ─────────────────────────────────────────────────────────────
function getEmployeeList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_CONFIG.EMPLOYEE_RESPONSES);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const nameVal = (data[i][EMP_COLS.NAME] || "").toString().trim();
    if (nameVal) {
      list.push({
        regNo      : (data[i][EMP_COLS.REG_NO]    || "").toString().trim(),
        name       : nameVal,
        // empId      : (data[i][EMP_COLS.EMP_ID]    || "").toString().trim(),
        location   : (data[i][EMP_COLS.LOCATION]  || "").toString().trim(),
        pan        : (data[i][EMP_COLS.PAN]        || "").toString().trim(),
        bank       : (data[i][EMP_COLS.BANK_NAME]  || "").toString().trim(),
        beneficiary: (data[i][EMP_COLS.BENEFICIARY]|| "").toString().trim(),
        accountNo  : (data[i][EMP_COLS.ACCOUNT_NO] || "").toString().trim(),
        ifsc       : (data[i][EMP_COLS.IFSC]       || "").toString().trim(),
      });
    }
  }
  return list;
}

function getVendorList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_CONFIG.VENDOR_RESPONSES);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const nameVal = (data[i][VENDOR_COLS.NAME] || "").toString().trim();
    if (nameVal) {
      list.push({
        regNo      : (data[i][VENDOR_COLS.REG_NO]     || "").toString().trim(),
        name       : nameVal,
        address    : (data[i][VENDOR_COLS.ADDRESS]    || "").toString().trim(),
        panGst     : (data[i][VENDOR_COLS.PAN_GST]    || "").toString().trim(),
        location   : (data[i][VENDOR_COLS.LOCATION]   || "").toString().trim(),
        bank       : (data[i][VENDOR_COLS.BANK_NAME]  || "").toString().trim(),
        beneficiary: (data[i][VENDOR_COLS.BENEFICIARY]|| "").toString().trim(),
        accountNo  : (data[i][VENDOR_COLS.ACCOUNT_NO] || "").toString().trim(),
        ifsc       : (data[i][VENDOR_COLS.IFSC]       || "").toString().trim(),
      });
    }
  }
  return list;
}

// UPDATED: returns ALL requests sorted by created date desc
function getAllRequests() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName(SHEET_CONFIG.REQUESTS);
  if (!sh)  return [];
  const data = sh.getDataRange().getValues();
  const requests = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][REQ_COLS.ID]) {
      requests.push(rowToRequest(data[i]));
    }
  }
  // Sort newest first — _createdRaw is already a ms timestamp number
  requests.sort((a, b) => (b._createdRaw || 0) - (a._createdRaw || 0));
  return requests;
}

function getRequestById(reqId) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName(SHEET_CONFIG.REQUESTS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][REQ_COLS.ID] === reqId) return rowToRequest(data[i]);
  }
  return null;
}

// ─────────────────────────────────────────────────────────────
// Safe date formatter — works whether the cell value is a
// JS Date object (from Apps Script) OR a string/number.
// new Date(dateObj).toLocaleString() silently returns "Invalid Date"
// in the Apps Script → client RPC bridge because Date objects get
// serialised as strings like "Mon Jan 01 2024 ..."; Utilities.formatDate
// is the only reliable way to format them server-side.
// ─────────────────────────────────────────────────────────────
function fmtDateTime(val) {
  if (!val || val === "") return "";
  try {
    const d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) return String(val);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  } catch(e) { return String(val); }
}
function fmtDate(val) {
  if (!val || val === "") return "";
  try {
    const d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) return String(val);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
  } catch(e) { return String(val); }
}

function rowToRequest(row) {
  return {
    id                  : (row[REQ_COLS.ID]           || "").toString(),
    timestamp           : fmtDateTime(row[REQ_COLS.TIMESTAMP]),
    requestor           : (row[REQ_COLS.REQUESTOR]    || "").toString(),
    contact             : (row[REQ_COLS.CONTACT]      || "").toString(),
    email               : (row[REQ_COLS.EMAIL]        || "").toString(),
    dept                : (row[REQ_COLS.DEPT]         || "").toString(),
    expenseType         : (row[REQ_COLS.EXPENSE_TYPE] || "").toString(),
    category            : (row[REQ_COLS.CATEGORY]     || "").toString(),
    description         : (row[REQ_COLS.DESCRIPTION]  || "").toString(),
    amount              : row[REQ_COLS.AMOUNT] || 0,
    date                : fmtDate(row[REQ_COLS.DATE]),
    payeeName           : (row[REQ_COLS.PAYEE_NAME]   || "").toString().trim(),
    payeeType           : (row[REQ_COLS.PAYEE_TYPE]   || "").toString().trim(),
    upi                 : (row[REQ_COLS.UPI]           || "").toString(),
    refId               : (row[REQ_COLS.REF_ID]        || "").toString(),
    invoiceUrl          : (row[REQ_COLS.INVOICE_URL]   || "").toString(),
    advanceReason       : (row[REQ_COLS.ADVANCE_REASON]|| "").toString(),
    expectedInvoiceDate : fmtDate(row[REQ_COLS.EXPECTED_INVOICE_DATE]),
    status              : (row[REQ_COLS.STATUS]        || "Pending").toString(),
    createdAt           : fmtDateTime(row[REQ_COLS.CREATED_AT] || row[REQ_COLS.TIMESTAMP]),
    // Keep raw value only for sorting — never rendered directly
    _createdRaw         : row[REQ_COLS.CREATED_AT] instanceof Date
                            ? row[REQ_COLS.CREATED_AT].getTime()
                            : (row[REQ_COLS.TIMESTAMP] instanceof Date
                                ? row[REQ_COLS.TIMESTAMP].getTime() : 0)
  };
}

function getApprovalLogs(reqId) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName(SHEET_CONFIG.APPROVAL_LOGS);
  if (!sh)  return [];
  const data = sh.getDataRange().getValues();
  const logs = [];
  for (let i = 1; i < data.length; i++) {
    if ((data[i][APPROVAL_COLS.REQUEST_ID] || "").toString() === reqId) {
      logs.push({
        logId    : (data[i][APPROVAL_COLS.LOG_ID]     || "").toString(),
        requestId: (data[i][APPROVAL_COLS.REQUEST_ID] || "").toString(),
        action   : (data[i][APPROVAL_COLS.ACTION]     || "").toString(),
        byName   : (data[i][APPROVAL_COLS.BY_NAME]    || "").toString(),
        byEmail  : (data[i][APPROVAL_COLS.BY_EMAIL]   || "").toString(),
        remarks  : (data[i][APPROVAL_COLS.REMARKS]    || "").toString(),
        timestamp: fmtDateTime(data[i][APPROVAL_COLS.TIMESTAMP]),
      });
    }
  }
  return logs;
}

function getPaymentLogs(reqId) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName(SHEET_CONFIG.PAYMENT_LOGS);
  if (!sh)  return [];
  const data = sh.getDataRange().getValues();
  const logs = [];
  for (let i = 1; i < data.length; i++) {
    if ((data[i][PAYMENT_COLS.REQUEST_ID] || "").toString() === reqId) {
      logs.push({
        logId        : (data[i][PAYMENT_COLS.LOG_ID]       || "").toString(),
        requestId    : (data[i][PAYMENT_COLS.REQUEST_ID]   || "").toString(),
        utr          : (data[i][PAYMENT_COLS.UTR]          || "").toString(),
        paymentAmount: data[i][PAYMENT_COLS.PAYMENT_AMOUNT] || 0,
        paidByName   : (data[i][PAYMENT_COLS.PAID_BY_NAME] || "").toString(),
        paidByEmail  : (data[i][PAYMENT_COLS.PAID_BY_EMAIL]|| "").toString(),
        paymentDate  : fmtDate(data[i][PAYMENT_COLS.PAYMENT_DATE]),
        mode         : (data[i][PAYMENT_COLS.MODE]         || "").toString(),
        remarks      : (data[i][PAYMENT_COLS.REMARKS]      || "").toString(),
        timestamp    : fmtDateTime(data[i][PAYMENT_COLS.TIMESTAMP]),
      });
    }
  }
  return logs;
}

// ─────────────────────────────────────────────────────────────
// PAYEE BANK DETAILS LOOKUP  (FIX: robust name matching)
// ─────────────────────────────────────────────────────────────
function getPayeeBankDetails(payeeName, payeeType) {
  if (!payeeName) return null;
  const name = payeeName.toString().trim().toLowerCase();
  const type = (payeeType || "").toString().trim().toLowerCase();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (type === "vendor") {
    const sh = ss.getSheetByName(SHEET_CONFIG.VENDOR_RESPONSES);
    if (!sh) return null;
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rowName = (data[i][VENDOR_COLS.NAME] || "").toString().trim().toLowerCase();
      if (rowName === name) {
        return {
          type       : "vendor",
          name       : data[i][VENDOR_COLS.NAME],
          regNo      : data[i][VENDOR_COLS.REG_NO]      || "",
          address    : data[i][VENDOR_COLS.ADDRESS]     || "",
          panGst     : data[i][VENDOR_COLS.PAN_GST]     || "",
          bank       : data[i][VENDOR_COLS.BANK_NAME]   || "",
          beneficiary: data[i][VENDOR_COLS.BENEFICIARY] || "",
          accountNo  : data[i][VENDOR_COLS.ACCOUNT_NO]  || "",
          ifsc       : data[i][VENDOR_COLS.IFSC]        || "",
        };
      }
    }
  } else {
    // employee / staff / organization → try employee sheet
    const sh = ss.getSheetByName(SHEET_CONFIG.EMPLOYEE_RESPONSES);
    if (!sh) return null;
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rowName = (data[i][EMP_COLS.NAME] || "").toString().trim().toLowerCase();
      if (rowName === name) {
        return {
          type       : "employee",
          name       : data[i][EMP_COLS.NAME],
          regNo      : data[i][EMP_COLS.REG_NO]       || "",
          // empId      : data[i][EMP_COLS.EMP_ID]       || "",
          pan        : data[i][EMP_COLS.PAN]           || "",
          bank       : data[i][EMP_COLS.BANK_NAME]     || "",
          beneficiary: data[i][EMP_COLS.BENEFICIARY]   || "",
          accountNo  : data[i][EMP_COLS.ACCOUNT_NO]    || "",
          ifsc       : data[i][EMP_COLS.IFSC]          || "",
        };
      }
    }
  }
  return null;
}

// ─────────────────────────────────────────────────────────────
// SUBMIT REQUEST
// ─────────────────────────────────────────────────────────────
function submitRequest(formData) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const sh  = ss.getSheetByName(SHEET_CONFIG.REQUESTS);
    const now = new Date();

    const lastRow = sh.getLastRow();
    const reqId   = "KE-EXP-" + now.getFullYear() + "-" + String(lastRow).padStart(4, "0");

    let invoiceUrl = "";
    if (formData.invoiceBase64 && formData.invoiceFileName) {
      invoiceUrl = saveInvoiceToDrive(formData.invoiceBase64, formData.invoiceFileName, reqId);
    }

    const row = new Array(20).fill("");
    row[REQ_COLS.ID]                    = reqId;
    row[REQ_COLS.TIMESTAMP]             = now;
    row[REQ_COLS.REQUESTOR]             = formData.requestorName || "";
    row[REQ_COLS.CONTACT]               = formData.contact       || "";
    row[REQ_COLS.EMAIL]                 = formData.email         || "";
    row[REQ_COLS.DEPT]                  = formData.department    || "";
    row[REQ_COLS.EXPENSE_TYPE]          = formData.expenseType   || "";
    row[REQ_COLS.CATEGORY]              = formData.category      || "";
    row[REQ_COLS.DESCRIPTION]           = formData.description   || "";
    row[REQ_COLS.AMOUNT]                = parseFloat(formData.amount) || 0;
    row[REQ_COLS.DATE]                  = formData.expenseDate ? new Date(formData.expenseDate) : "";
    row[REQ_COLS.PAYEE_NAME]            = formData.payeeName     || "";
    row[REQ_COLS.PAYEE_TYPE]            = formData.payeeType     || "";
    row[REQ_COLS.UPI]                   = formData.upiId         || "";
    row[REQ_COLS.REF_ID]                = formData.refId         || "";
    row[REQ_COLS.INVOICE_URL]           = invoiceUrl;
    row[REQ_COLS.ADVANCE_REASON]        = formData.advanceReason || "";
    row[REQ_COLS.EXPECTED_INVOICE_DATE] = formData.expectedInvoiceDate
                                            ? new Date(formData.expectedInvoiceDate) : "";
    row[REQ_COLS.STATUS]                = "Pending";
    row[REQ_COLS.CREATED_AT]            = now;

    sh.appendRow(row);
    return { success: true, reqId };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function saveInvoiceToDrive(base64Data, fileName, reqId) {
  try {
    const folder = getOrCreateFolder("Kings Equestrian Invoices/" + reqId);
    const blob   = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      getMimeType(fileName),
      fileName
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    return "Upload failed: " + e.message;
  }
}

function getMimeType(fileName) {
  const ext = fileName.split(".").pop().toLowerCase();
  const map = { pdf: "application/pdf", jpg: "image/jpeg", jpeg: "image/jpeg", png: "image/png" };
  return map[ext] || "application/octet-stream";
}

function getOrCreateFolder(path) {
  const parts  = path.split("/");
  let   folder = DriveApp.getRootFolder();
  for (const part of parts) {
    const found = folder.getFoldersByName(part);
    folder = found.hasNext() ? found.next() : folder.createFolder(part);
  }
  return folder;
}

// ─────────────────────────────────────────────────────────────
// ADMIN: APPROVE / REJECT
// ─────────────────────────────────────────────────────────────
function approveRequest(reqId, remarks) {
  return changeRequestStatus(reqId, "Approved", remarks);
}

function rejectRequest(reqId, remarks) {
  return changeRequestStatus(reqId, "Rejected", remarks);
}

function changeRequestStatus(reqId, status, remarks) {
  try {
    const user = getUserInfo();
    if (user.role !== "admin") throw new Error("Unauthorized");

    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const sh   = ss.getSheetByName(SHEET_CONFIG.REQUESTS);
    const data = sh.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][REQ_COLS.ID] === reqId) {
        sh.getRange(i + 1, REQ_COLS.STATUS + 1).setValue(status);
        break;
      }
    }

    const logSh = ss.getSheetByName(SHEET_CONFIG.APPROVAL_LOGS);
    logSh.appendRow([
      "AL-" + Date.now(), reqId, status,
      user.name, user.email, remarks || "", new Date()
    ]);

    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ─────────────────────────────────────────────────────────────
// ACCOUNTS: SETTLE PAYMENT  (UPDATED — records paymentAmount + mode)
// ─────────────────────────────────────────────────────────────
function settlePayment(reqId, utr, paymentAmount, paymentDate, mode, remarks) {
  try {
    const user = getUserInfo();
    if (user.role !== "accounts") throw new Error("Unauthorized");

    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const sh   = ss.getSheetByName(SHEET_CONFIG.REQUESTS);
    const data = sh.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][REQ_COLS.ID] === reqId) {
        sh.getRange(i + 1, REQ_COLS.STATUS + 1).setValue("Paid");
        break;
      }
    }

    // Payment log
    const payLog = ss.getSheetByName(SHEET_CONFIG.PAYMENT_LOGS);
    payLog.appendRow([
      "PL-" + Date.now(), reqId, utr,
      parseFloat(paymentAmount) || 0,
      user.name, user.email,
      paymentDate ? new Date(paymentDate) : new Date(),
      mode || "",
      remarks || "",
      new Date()
    ]);

    // Approval trail entry
    const appLog = ss.getSheetByName(SHEET_CONFIG.APPROVAL_LOGS);
    appLog.appendRow([
      "AL-" + Date.now(), reqId, "Payment Settled",
      user.name, user.email,
      "UTR: " + utr + " | Amount: ₹" + paymentAmount + " | Mode: " + (mode || "—"),
      new Date()
    ]);

    // Email the requestor
    sendPaymentEmail(reqId, utr, paymentAmount, mode, paymentDate);

    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function sendPaymentEmail(reqId, utr, paymentAmount, mode, paymentDate) {
  try {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const sh   = ss.getSheetByName(SHEET_CONFIG.REQUESTS);
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][REQ_COLS.ID] === reqId) {
        const to      = data[i][REQ_COLS.EMAIL];
        const name    = data[i][REQ_COLS.REQUESTOR];
        const reqAmt  = data[i][REQ_COLS.AMOUNT];
        if (!to) return;
        MailApp.sendEmail(
          to,
          "[Kings Equestrian] Payment Processed — " + reqId,
          "Dear " + name + ",\n\n" +
          "Your expense request has been processed.\n\n" +
          "Request ID     : " + reqId + "\n" +
          "Request Amount : ₹" + reqAmt + "\n" +
          "Payment Amount : ₹" + paymentAmount + "\n" +
          "UTR / Ref      : " + utr + "\n" +
          "Mode           : " + (mode || "—") + "\n" +
          "Date           : " + (paymentDate || new Date().toLocaleDateString("en-IN")) + "\n\n" +
          "Kings Equestrian Accounts Team"
        );
        break;
      }
    }
  } catch (e) {
    Logger.log("Email error: " + e.message);
  }
}

// ─────────────────────────────────────────────────────────────
// DASHBOARD SUMMARY
// ─────────────────────────────────────────────────────────────
function getDashboardSummary() {
  const requests = getAllRequests();
  const s = {
    total: requests.length,
    pending: 0, approved: 0, rejected: 0, paid: 0,
    totalAmount: 0, pendingAmount: 0, approvedAmount: 0, paidAmount: 0,
  };
  for (const r of requests) {
    s.totalAmount += Number(r.amount) || 0;
    switch ((r.status || "").toLowerCase()) {
      case "pending" : s.pending++;  s.pendingAmount  += Number(r.amount) || 0; break;
      case "approved": s.approved++; s.approvedAmount += Number(r.amount) || 0; break;
      case "rejected": s.rejected++; break;
      case "paid"    : s.paid++;     s.paidAmount     += Number(r.amount) || 0; break;
    }
  }
  return s;
}