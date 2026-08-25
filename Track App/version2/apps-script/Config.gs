/**
 * Kings Equestrian Track App – Version 2 (simplified)
 * Sheets-backed inventory: Vendors, Items, Inventory, Requests, Issues.
 */

var KE = {
  SHEETS: {
    VENDOR: "VENDOR_MASTER",
    ITEM: "ITEM_MASTER",
    INVENTORY: "INVENTORY",
    REQUEST: "REQUEST_REGISTER",
    ISSUE: "ISSUE_REGISTER",
    STUDENT: "STUDENT_MASTER",
    USER: "USER_MASTER",
    ACTIVITY: "ACTIVITY_LOG"
  },
  LOCATIONS: ["Bangalore", "Hyderabad", "Pune", "Farm"],
  LOCATION_ALL: "All",
  DEFAULT_ADMIN_LOCATION: "Farm",
  LIST_DEFAULT_LIMIT: 10,
  ROLES: {
    ADMIN: "Admin",
    MANAGER: "Manager",
    TRAINER: "Trainer"
  },
  REQUEST_STATUS: {
    PENDING: "Pending",
    ORDERED: "Ordered",
    RECEIVED: "Received",
    CANCELLED: "Cancelled"
  },
  ISSUE_STATUS: {
    ISSUED: "Issued"
  },
  WEEKLY_REPORT_RECIPIENTS: [
    "kingsequestrianfoundation@gmail.com",
    "kingsequestrianclub@gmail.com",
    "komalshriwas4539@gmail.com"
  ]
};

var SHEET_HEADERS_ = {};
SHEET_HEADERS_[KE.SHEETS.VENDOR] = [
  "Vendor ID", "Vendor Name", "Code", "Phone", "Email", "Location",
  "Added By", "Date Added"
];
SHEET_HEADERS_[KE.SHEETS.ITEM] = [
  "Item Code", "Item Name", "Category", "Vendor", "Location", "Min Level",
  "Added By", "Date Added", "Image URL"
];
SHEET_HEADERS_[KE.SHEETS.INVENTORY] = [
  "Item Code", "Item Name", "Location", "Qty Added", "Date Added", "Added By"
];
SHEET_HEADERS_[KE.SHEETS.REQUEST] = [
  "Request ID", "Date", "Vendor", "Item", "Qty",
  "Location", "Requested By", "Status", "Notes"
];
SHEET_HEADERS_[KE.SHEETS.ISSUE] = [
  "Issue ID", "Item", "Qty", "Student Name", "KE Number",
  "Location", "Issued By", "Date"
];
SHEET_HEADERS_[KE.SHEETS.STUDENT] = [
  "Student_ID", "Student_Name", "Parent_Name", "Email", "Phone",
  "Program", "Skill_Riding_Avg", "Grade", "Section", "Location",
  "Added By", "Date Added"
];
SHEET_HEADERS_[KE.SHEETS.USER] = [
  "Email", "Name", "Role", "Location", "Password"
];
SHEET_HEADERS_[KE.SHEETS.ACTIVITY] = [
  "Date", "Action", "Entity Type", "Entity ID", "Details", "Location", "User", "Role"
];
