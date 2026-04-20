// ============================================================
// KINGS EQUESTRIAN - EXPENSE & PAYMENT MANAGEMENT SYSTEM
// Google Apps Script Web App  |  Code.gs
// ============================================================

const SHEET_CONFIG = {
  USERS             : "Users",
  VENDOR_RESPONSES  : "Vendor Registration",
  EMPLOYEE_RESPONSES: "Employee Registration",
  ACCESS_CONTROL    : "Access Control",
  REQUESTS          : "Requests",
  APPROVAL_LOGS     : "ApprovalLog",
  PAYMENT_LOGS      : "PaymentsLog",
  AUDIT_LOGS        : "AuditLogs",
};

// ── Column indexes (0-based) ─────────────────────────────────
const USERS_COLS = {
  EMAIL: 0, PASSWORD_HASH: 1, ROLE: 2, SECRET: 3, BACKUP_EMAIL: 4, NAME: 5, LAST_LOGIN: 6,
  USER_ID: 7, EMP_REG_NO: 8, IS_ACTIVE: 9
};

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

  ensureSheet(ss, SHEET_CONFIG.USERS, [
    "Email", "PasswordHash", "Role", "Secret", "BackupEmail", "Name", "LastLogin", "UserId", "EmployeeRegNo", "IsActive"
  ]);

  ensureSheet(ss, SHEET_CONFIG.REQUESTS, [
    "Request ID", "Timestamp", "Requestor Name", "Contact", "Email",
    "Department", "Expense Type", "Expense Category", "Description",
    "Amount (₹)", "Expense Date", "Payee Name", "Payee Type",
    "UPI ID", "Reference Expense ID", "Invoice URL",
    "Advance Reason", "Expected Invoice Date", "Status", "Created At", "Location"
  ]);

  ensureSheet(ss, SHEET_CONFIG.APPROVAL_LOGS, [
    "Log ID", "Request ID", "Action", "By Name", "By Email",
    "Remarks", "Timestamp"
  ]);

  ensureSheet(ss, SHEET_CONFIG.PAYMENT_LOGS, [
    "Log ID", "Request ID", "UTR Number", "Payment Amount (₹)",
    "Paid By Name", "Paid By Email", "Payment Date",
    "Mode", "Remarks", "Timestamp"
  ]);

  ensureSheet(ss, SHEET_CONFIG.AUDIT_LOGS, [
    "Timestamp", "Action", "Email", "Details"
  ]);

  ensureUsersSheetStructure(ss);
  generateVendorRegNos(ss);
  generateEmployeeRegNos(ss);
  syncEmployeeUserAccounts(ss);
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

function ensureUsersSheetStructure(ss) {
  const sh = ss.getSheetByName(SHEET_CONFIG.USERS);
  if (!sh) return;
  const headers = [
    "Email", "PasswordHash", "Role", "Secret", "BackupEmail", "Name", "LastLogin", "UserId", "EmployeeRegNo", "IsActive"
  ];
  headers.forEach((header, idx) => {
    const cell = sh.getRange(1, idx + 1);
    if (!cell.getValue()) {
      cell.setValue(header)
        .setFontWeight("bold")
        .setBackground("#0D1F0F")
        .setFontColor("#D4AF5A");
    }
  });
}

function normalizeRole(role) {
  const value = (role || "").toString().toLowerCase().trim();
  if (["admin", "administrator", "approver", "super admin", "superadmin"].includes(value)) return "admin";
  if (["accounts", "account", "finance", "accounts team"].includes(value)) return "accounts";
  if (["user", "employee", "staff", "requestor", "viewer"].includes(value)) return "user";
  return value || "user";
}

function generateUserId(name, regNo, usedIds) {
  const baseName = (name || "user").toString().toLowerCase().replace(/[^a-z0-9]+/g, "").slice(0, 8) || "user";
  const regPart = (regNo || Utilities.getUuid()).toString().replace(/[^0-9a-z]/gi, "").slice(-4).toLowerCase() || "0001";
  let userId = (baseName + regPart).toLowerCase();
  let counter = 1;
  while (usedIds[userId]) {
    userId = (baseName + regPart + counter).toLowerCase();
    counter++;
  }
  usedIds[userId] = true;
  return userId;
}

function createInitialPassword(regNo, aadhar) {
  const regPart = (regNo || "0000").toString().replace(/\D/g, "").slice(-4).padStart(4, "0");
  const aadPart = (aadhar || "1234").toString().replace(/\D/g, "").slice(-4).padStart(4, "0");
  return "KE@" + regPart + aadPart;
}

function syncEmployeeUserAccounts(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const userSh = ss.getSheetByName(SHEET_CONFIG.USERS);
  const empSh = ss.getSheetByName(SHEET_CONFIG.EMPLOYEE_RESPONSES);
  const accessSh = ss.getSheetByName(SHEET_CONFIG.ACCESS_CONTROL);
  if (!userSh || !empSh) return [];

  const users = userSh.getDataRange().getValues();
  const employees = empSh.getDataRange().getValues();
  const accessRows = accessSh ? accessSh.getDataRange().getValues() : [];
  const created = [];
  const usedIds = {};

  for (let i = 1; i < users.length; i++) {
    const existingId = (users[i][USERS_COLS.USER_ID] || "").toString().trim().toLowerCase();
    if (existingId) usedIds[existingId] = true;

    const existingRole = normalizeRole(users[i][USERS_COLS.ROLE] || "user");
    if (users[i][USERS_COLS.ROLE] !== existingRole) {
      userSh.getRange(i + 1, USERS_COLS.ROLE + 1).setValue(existingRole);
    }
    if (!users[i][USERS_COLS.IS_ACTIVE]) {
      userSh.getRange(i + 1, USERS_COLS.IS_ACTIVE + 1).setValue("Active");
    }
  }

  for (let i = 1; i < employees.length; i++) {
    const name = (employees[i][EMP_COLS.NAME] || "").toString().trim();
    const regNo = (employees[i][EMP_COLS.REG_NO] || "").toString().trim();
    if (!name || !regNo) continue;

    let existingRow = -1;
    for (let j = 1; j < users.length; j++) {
      const userRegNo = (users[j][USERS_COLS.EMP_REG_NO] || "").toString().trim();
      const userName = (users[j][USERS_COLS.NAME] || "").toString().trim().toLowerCase();
      if (userRegNo === regNo || userName === name.toLowerCase()) {
        existingRow = j;
        break;
      }
    }

    const accessRow = accessRows.find(row => (row[ACCESS_COLS.NAME] || "").toString().trim().toLowerCase() === name.toLowerCase());
    const mappedEmail = accessRow ? (accessRow[ACCESS_COLS.EMAIL] || "").toString().trim() : "";
    const mappedRole = normalizeRole(accessRow ? accessRow[ACCESS_COLS.ROLE] : "user");

    if (existingRow >= 0) {
      if (!users[existingRow][USERS_COLS.USER_ID]) {
        const generatedId = generateUserId(name, regNo, usedIds);
        userSh.getRange(existingRow + 1, USERS_COLS.USER_ID + 1).setValue(generatedId);
        users[existingRow][USERS_COLS.USER_ID] = generatedId;
      }
      if (!users[existingRow][USERS_COLS.EMP_REG_NO]) {
        userSh.getRange(existingRow + 1, USERS_COLS.EMP_REG_NO + 1).setValue(regNo);
      }
      if (mappedEmail && !users[existingRow][USERS_COLS.EMAIL]) {
        userSh.getRange(existingRow + 1, USERS_COLS.EMAIL + 1).setValue(mappedEmail);
      }
      if (mappedRole && users[existingRow][USERS_COLS.ROLE] !== mappedRole) {
        userSh.getRange(existingRow + 1, USERS_COLS.ROLE + 1).setValue(mappedRole);
      }
      if (!users[existingRow][USERS_COLS.IS_ACTIVE]) {
        userSh.getRange(existingRow + 1, USERS_COLS.IS_ACTIVE + 1).setValue("Active");
      }
      continue;
    }

    const userId = generateUserId(name, regNo, usedIds);
    const initialPassword = createInitialPassword(regNo, employees[i][EMP_COLS.AADHAR]);
    userSh.appendRow([
      mappedEmail,
      hashPassword(initialPassword),
      mappedRole,
      "",
      mappedEmail,
      name,
      "",
      userId,
      regNo,
      "Active"
    ]);

    created.push({ name, regNo, userId, tempPassword: initialPassword, role: mappedRole, email: mappedEmail });
  }

  return created;
}

function provisionEmployeeLogins() {
  const created = syncEmployeeUserAccounts(SpreadsheetApp.getActiveSpreadsheet());
  return {
    success: true,
    count: created.length,
    accounts: created,
    message: created.length ? (created.length + " employee login(s) created/refreshed.") : "No new employee accounts were required."
  };
}

function resetUserPassword(userIdOrEmail) {
  const value = (userIdOrEmail || "").toString().trim().toLowerCase();
  if (!value) return { success: false, message: "User ID or email is required." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSh = ss.getSheetByName(SHEET_CONFIG.USERS);
  const empSh = ss.getSheetByName(SHEET_CONFIG.EMPLOYEE_RESPONSES);
  if (!userSh) return { success: false, message: "Users sheet not found." };

  const users = userSh.getDataRange().getValues();
  const employees = empSh ? empSh.getDataRange().getValues() : [];

  for (let i = 1; i < users.length; i++) {
    const rowEmail = (users[i][USERS_COLS.EMAIL] || "").toString().trim().toLowerCase();
    const rowUserId = (users[i][USERS_COLS.USER_ID] || "").toString().trim().toLowerCase();
    if (rowEmail === value || rowUserId === value) {
      const regNo = (users[i][USERS_COLS.EMP_REG_NO] || "").toString().trim();
      let aadhar = "1234";
      for (let j = 1; j < employees.length; j++) {
        if ((employees[j][EMP_COLS.REG_NO] || "").toString().trim() === regNo) {
          aadhar = employees[j][EMP_COLS.AADHAR] || "1234";
          break;
        }
      }
      const tempPassword = createInitialPassword(regNo, aadhar);
      userSh.getRange(i + 1, USERS_COLS.PASSWORD_HASH + 1).setValue(hashPassword(tempPassword));
      userSh.getRange(i + 1, USERS_COLS.IS_ACTIVE + 1).setValue("Active");
      return {
        success: true,
        userId: users[i][USERS_COLS.USER_ID] || "",
        email: users[i][USERS_COLS.EMAIL] || "",
        tempPassword,
        message: "Password reset successfully."
      };
    }
  }

  return { success: false, message: "User not found." };
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
  syncEmployeeUserAccounts(ss);
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
      return normalizeRole(data[i][ACCESS_COLS.ROLE] || "user");
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
// AUTHENTICATION SYSTEM
// ─────────────────────────────────────────────────────────────
function hashPassword(password) {
  return Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password)
  );
}

function loginUser(userIdOrEmail, password) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG.USERS);
    if (!sheet) return { status: "ERROR", message: "Users sheet not found" };

    const identifier = (userIdOrEmail || "").toString().toLowerCase().trim();
    const data = sheet.getDataRange().getValues();
    const hash = hashPassword(password);

    for (let i = 1; i < data.length; i++) {
      const userEmail = (data[i][USERS_COLS.EMAIL] || "").toString().toLowerCase().trim();
      const userId = (data[i][USERS_COLS.USER_ID] || "").toString().toLowerCase().trim();
      const isActive = ((data[i][USERS_COLS.IS_ACTIVE] || "Active").toString().trim().toLowerCase() !== "inactive");
      if (userEmail === identifier || userId === identifier) {
        if (!isActive) {
          logAction("LOGIN_BLOCKED", userEmail || userIdOrEmail, "Inactive account attempted login");
          return { status: "INVALID", message: "This account is inactive. Please contact admin." };
        }

        if (data[i][USERS_COLS.PASSWORD_HASH] === hash) {
          const email = data[i][USERS_COLS.EMAIL] || userIdOrEmail;
          const role = normalizeRole(data[i][USERS_COLS.ROLE] || "user");
          const userName = data[i][USERS_COLS.NAME] || email;
          const resolvedUserId = data[i][USERS_COLS.USER_ID] || userIdOrEmail;

          if (role === "admin") {
            generateSecret(email);
            logAction("LOGIN_INITIATED", email, "2FA required for " + role);
            return { status: "2FA_REQUIRED", role: role, email: email, name: userName, userId: resolvedUserId };
          }

          const sessionToken = generateSessionToken(email);
          sheet.getRange(i + 1, USERS_COLS.LAST_LOGIN + 1).setValue(new Date());
          logAction("LOGIN_SUCCESS", email, "Secure login via user-specific credentials");

          return {
            status: "SUCCESS",
            role: role,
            email: email,
            name: userName,
            userId: resolvedUserId,
            sessionToken: sessionToken
          };
        }

        logAction("LOGIN_FAILED", userEmail || userIdOrEmail, "Invalid password");
        return { status: "INVALID", message: "Invalid user ID/email or password" };
      }
    }

    logAction("LOGIN_FAILED", userIdOrEmail, "User not found");
    return { status: "INVALID", message: "Invalid user ID/email or password" };
  } catch (e) {
    return { status: "ERROR", message: e.toString() };
  }
}

function generateSecret(email) {
  const secret = Utilities.getUuid().replace(/-/g, '').substring(0, 16);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG.USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if ((data[i][USERS_COLS.EMAIL] || "").toString().toLowerCase().trim() === email.toLowerCase().trim()) {
      sheet.getRange(i + 1, USERS_COLS.SECRET + 1).setValue(secret);
      break;
    }
  }
  return secret;
}

function generateSessionToken(email) {
  const token = Utilities.base64Encode(email + "|" + Date.now() + "|" + Utilities.getUuid());
  PropertiesService.getScriptProperties().setProperty("session_" + email, JSON.stringify({
    token: token,
    createdAt: Date.now(),
    expiresAt: Date.now() + 24 * 60 * 60 * 1000 // 24 hours
  }));
  return token;
}

function verifySessionToken(email, token) {
  const sessionStr = PropertiesService.getScriptProperties().getProperty("session_" + email);
  if (!sessionStr) return false;
  
  const session = JSON.parse(sessionStr);
  if (session.token !== token) return false;
  if (Date.now() > session.expiresAt) return false;
  
  return true;
}

function sendOTP(email) {
  const otp = Math.floor(100000 + Math.random() * 900000).toString();
  const expiryTime = Date.now() + 5 * 60 * 1000; // 5 minutes
  
  PropertiesService.getScriptProperties().setProperty("otp_" + email, JSON.stringify({
    code: otp,
    expiry: expiryTime,
    attempts: 0
  }));
  
  // Send OTP via email
  try {
    const backupEmail = getBackupEmail(email);
    MailApp.sendEmail(
      backupEmail || email,
      "Kings Equestrian - Your OTP Code",
      "Your verification code is: " + otp + "\n\nThis code expires in 5 minutes."
    );
    return { status: "SUCCESS", message: "OTP sent to " + (backupEmail || email) };
  } catch (e) {
    return { status: "ERROR", message: "Failed to send OTP: " + e.toString() };
  }
}

function verifyOTP(email, otp) {
  const otpStr = PropertiesService.getScriptProperties().getProperty("otp_" + email);
  if (!otpStr) {
    return { status: "INVALID", message: "No OTP found. Request a new one." };
  }
  
  const otpData = JSON.parse(otpStr);
  
  if (Date.now() > otpData.expiry) {
    PropertiesService.getScriptProperties().deleteProperty("otp_" + email);
    return { status: "EXPIRED", message: "OTP has expired. Request a new one." };
  }
  
  if (otpData.attempts >= 3) {
    PropertiesService.getScriptProperties().deleteProperty("otp_" + email);
    return { status: "LOCKED", message: "Too many attempts. Request a new OTP." };
  }
  
  if (otpData.code !== otp.toString()) {
    otpData.attempts++;
    PropertiesService.getScriptProperties().setProperty("otp_" + email, JSON.stringify(otpData));
    return { status: "INVALID", message: "Invalid OTP. " + (3 - otpData.attempts) + " attempts remaining." };
  }
  
  // OTP verified successfully
  PropertiesService.getScriptProperties().deleteProperty("otp_" + email);
  const sessionToken = generateSessionToken(email);
  logAction("LOGIN_SUCCESS", email, "2FA verified");
  
  return { 
    status: "SUCCESS", 
    sessionToken: sessionToken,
    message: "Verified successfully" 
  };
}

function getBackupEmail(email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG.USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if ((data[i][USERS_COLS.EMAIL] || "").toString().toLowerCase().trim() === email.toLowerCase().trim()) {
      return data[i][USERS_COLS.BACKUP_EMAIL] || null;
    }
  }
  return null;
}

function getCurrentUser(email, sessionToken) {
  if (!verifySessionToken(email, sessionToken)) {
    return { status: "UNAUTHORIZED" };
  }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG.USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if ((data[i][USERS_COLS.EMAIL] || "").toString().toLowerCase().trim() === email.toLowerCase().trim()) {
      return {
        status: "SUCCESS",
        email: email,
        name: data[i][USERS_COLS.NAME] || email,
        role: normalizeRole(data[i][USERS_COLS.ROLE] || "user"),
        userId: data[i][USERS_COLS.USER_ID] || "",
        employeeRegNo: data[i][USERS_COLS.EMP_REG_NO] || "",
        lastLogin: data[i][USERS_COLS.LAST_LOGIN] || null
      };
    }
  }
  
  return { status: "UNAUTHORIZED" };
}

function authorizeSessionUser(email, sessionToken, allowedRoles) {
  const user = getCurrentUser(email || "", sessionToken || "");
  if (!user || user.status !== "SUCCESS") {
    throw new Error("Unauthorized. Please sign in again.");
  }

  const role = normalizeRole(user.role || "user");
  if (allowedRoles && allowedRoles.length && allowedRoles.indexOf(role) === -1) {
    throw new Error("Unauthorized");
  }

  return user;
}

function logAction(action, email, details) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG.AUDIT_LOGS);
  if (sheet) {
    sheet.appendRow([
      new Date(),
      action,
      email,
      details
    ]);
  }
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
  
  if (!sh) {
    // Logger.log("Sheet not found: " + SHEET_CONFIG.VENDOR_RESPONSES);
    return [];
  }

  const data = sh.getDataRange().getValues();
  // Logger.log("Total rows (including header): " + data.length);

  const list = [];

  for (let i = 1; i < data.length; i++) {
    const nameVal = (data[i][VENDOR_COLS.NAME] || "").toString().trim();

    if (nameVal) {
      const vendorObj = {
        regNo      : (data[i][VENDOR_COLS.REG_NO]     || "").toString().trim(),
        name       : nameVal,
        address    : (data[i][VENDOR_COLS.ADDRESS]    || "").toString().trim(),
        panGst     : (data[i][VENDOR_COLS.PAN_GST]    || "").toString().trim(),
        location   : (data[i][VENDOR_COLS.LOCATION]   || "").toString().trim(),
        bank       : (data[i][VENDOR_COLS.BANK_NAME]  || "").toString().trim(),
        beneficiary: (data[i][VENDOR_COLS.BENEFICIARY]|| "").toString().trim(),
        accountNo  : (data[i][VENDOR_COLS.ACCOUNT_NO] || "").toString().trim(),
        ifsc       : (data[i][VENDOR_COLS.IFSC]       || "").toString().trim(),
      };

       Logger.log("Row " + i + " Vendor: " + JSON.stringify(vendorObj));

      list.push(vendorObj);
    } else {
      // Logger.log("Row " + i + " skipped (no name)");
    }
  }

  // Logger.log("Final Vendor List Count: " + list.length);
  // Logger.log("Final Vendor List: " + JSON.stringify(list));

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
    const authUser = getCurrentUser(formData.email || "", formData.sessionToken || "");
    if (!authUser || authUser.status !== "SUCCESS") {
      throw new Error("Unauthorized. Please sign in again.");
    }

    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const sh  = ss.getSheetByName(SHEET_CONFIG.REQUESTS);
    const now = new Date();

    const lastRow = sh.getLastRow();
    const reqId   = "KE-EXP-" + now.getFullYear() + "-" + String(lastRow).padStart(4, "0");

    let invoiceUrl = "";
    if (formData.invoiceBase64 && formData.invoiceFileName) {
      invoiceUrl = saveInvoiceToDrive(formData.invoiceBase64, formData.invoiceFileName, reqId);
    }

    const row = new Array(21).fill("");
    row[REQ_COLS.ID]                    = reqId;
    row[REQ_COLS.TIMESTAMP]             = now;
    row[REQ_COLS.REQUESTOR]             = authUser.name || formData.requestorName || "";
    row[REQ_COLS.CONTACT]               = formData.contact       || "";
    row[REQ_COLS.EMAIL]                 = authUser.email || formData.email || "";
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
    row[REQ_COLS.LOCATION]              = formData.location      || "";

    sh.appendRow(row);
    logAction("REQUEST_SUBMITTED", authUser.email || authUser.userId || "", "Request " + reqId + " submitted by authenticated user " + (authUser.name || ""));
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
function approveRequest(reqId, remarks, email, sessionToken) {
  return changeRequestStatus(reqId, "Approved", remarks, email, sessionToken);
}

function rejectRequest(reqId, remarks, email, sessionToken) {
  return changeRequestStatus(reqId, "Rejected", remarks, email, sessionToken);
}

function changeRequestStatus(reqId, status, remarks, email, sessionToken) {
  try {
    const user = authorizeSessionUser(email, sessionToken, ["admin"]);

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
function settlePayment(reqId, utr, paymentAmount, paymentDate, mode, remarks, email, sessionToken) {
  try {
    const user = authorizeSessionUser(email, sessionToken, ["accounts", "admin"]);

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



function generateHash() {
  const password = "emp@123KG"; // change this
  const hash = Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password)
  );
  Logger.log(hash); // copy this output
}