/**
 * Kings Equestrian Inventory System – configuration
 */

var KE = {
  SHEETS: {
    INVENTORY: "INVENTORY_MASTER",
    VENDOR: "VENDOR_MASTER",
    REQUEST: "REQUEST_REGISTER",
    PAYMENT: "PAYMENTS",
    ORDER: "VENDOR_ORDERS",
    RECEIVED: "GOODS_RECEIVED",
    ISSUE: "ISSUE_REGISTER",
    USER: "USER_MASTER",
    TRANSFER: "STOCK_TRANSFER",
    SAMPLE: "SAMPLE_MOVEMENT",
    PRICING: "ITEM_PRICING_MASTER",
    STORAGE: "STORAGE_LOCATIONS"
  },
  DEFAULT_ADMIN_LOCATION: "Farm",
  ROLES: {
    ADMIN: "Admin",
    ACCOUNTS: "Accounts",
    TRAINER: "Trainer",
    MANAGER: "Manager"
  },
  REQUEST_STATUS: {
    PENDING: "Pending",
    PAID: "Payment Received",
    APPROVED: "Approved",
    ORDERED: "Ordered",
    RECEIVED: "Received",
    ISSUED: "Issued",
    COMPLETED: "Completed",
    CANCELLED: "Cancelled"
  },
  PAYMENT_STATUS: {
    RECORDED: "Recorded",
    VERIFIED: "Verified"
  },
  ORDER_STATUS: {
    PLACED: "Placed",
    PARTIAL: "Partially Received",
    RECEIVED: "Received",
    CANCELLED: "Cancelled"
  },
  INVENTORY_TYPES: [
    "Business Inventory",
    "Procurement Inventory",
    "Sample / Demo Inventory",
    "Operational Inventory"
  ],
  DEFAULT_INVENTORY_TYPE: "Business Inventory",
  TRANSFER_STATUS: {
    REQUESTED: "Requested",
    APPROVED: "Approved",
    COMPLETED: "Completed",
    CANCELLED: "Cancelled"
  },
  SAMPLE_STATUS: {
    OUT: "Out",
    RETURNED: "Returned",
    LOST: "Lost"
  },
  DEAD_STOCK_DAYS: 90
};

var SHEET_HEADERS_ = {};
SHEET_HEADERS_[KE.SHEETS.INVENTORY] = [
  "Item Code", "Item Name", "Category", "Model", "Vendor", "Location",
  "Current Qty", "Min Level", "Storage Location", "Last Updated",
  "Inventory Type", "Allocated To", "Reserved Qty", "Cost Price", "Last Movement", "Image URL"
];
SHEET_HEADERS_[KE.SHEETS.VENDOR] = [
  "Vendor ID", "Vendor Name", "Code", "Phone", "Email", "Location"
];
SHEET_HEADERS_[KE.SHEETS.REQUEST] = [
  "Request ID", "Date", "Vendor", "Item", "Qty",
  "Location", "Requested By", "Payment Status", "Status"
];
SHEET_HEADERS_[KE.SHEETS.PAYMENT] = [
  "Payment ID", "Request ID", "Vendor", "Amount", "Mode", "Status", "Date"
];
SHEET_HEADERS_[KE.SHEETS.ORDER] = [
  "Order ID", "Request ID", "Vendor", "Item", "Qty", "Location", "Order Date", "Status"
];
SHEET_HEADERS_[KE.SHEETS.RECEIVED] = [
  "GRN ID", "Order ID", "Item", "Qty", "Location", "Received By", "Storage Location", "Date"
];
SHEET_HEADERS_[KE.SHEETS.ISSUE] = [
  "Issue ID", "Request ID", "Item", "Qty", "Issued To", "Location", "Issued By", "Date"
];
SHEET_HEADERS_[KE.SHEETS.USER] = [
  "Email", "Name", "Role", "Location", "Password"
];
SHEET_HEADERS_[KE.SHEETS.TRANSFER] = [
  "Transfer ID", "Item Code", "Item Name", "From Location", "To Location",
  "Qty", "Requested By", "Approved By", "Date", "Status", "Notes"
];
SHEET_HEADERS_[KE.SHEETS.SAMPLE] = [
  "Sample ID", "Item Code", "Item Name", "Location", "Issued To",
  "Purpose", "Date Out", "Expected Return", "Date Returned", "Status", "Notes"
];
SHEET_HEADERS_[KE.SHEETS.PRICING] = [
  "Item Code", "Item Name", "Cost Price", "Selling Price", "MRP",
  "Profit Margin %", "Preferred Vendor", "Last Purchase Cost", "Last Updated"
];

SHEET_HEADERS_[KE.SHEETS.STORAGE] = [
  "Storage ID", "Storage Name", "Location", "Description", "Added By", "Date Added"
];

var ADMIN_ONLY_SHEETS_ = [KE.SHEETS.PRICING];
