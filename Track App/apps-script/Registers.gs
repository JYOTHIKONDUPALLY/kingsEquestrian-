/**
 * Register log views – sort by date desc, 5 per page, optional date/status filters.
 */

var REGISTER_KEYS_ = {
  payments: { sheet: KE.SHEETS.PAYMENT, dateCol: "Date" },
  orders: { sheet: KE.SHEETS.ORDER, dateCol: "Order Date" },
  received: { sheet: KE.SHEETS.RECEIVED, dateCol: "Date" },
  issues: { sheet: KE.SHEETS.ISSUE, dateCol: "Date" },
  requests: { sheet: KE.SHEETS.REQUEST, dateCol: "Date" }
};

var REGISTER_PAGE_SIZE_ = 5;

function parseDateCell_(value) {
  if (value == null || value === "") {
    return 0;
  }
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return value.getTime();
  }
  var d = new Date(String(value).trim());
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function normalizeDateFilter_(s) {
  s = normalize_(s);
  if (!s) {
    return "";
  }
  var d = new Date(s);
  if (isNaN(d.getTime())) {
    return "";
  }
  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone() || "Asia/Kolkata",
    "yyyy-MM-dd"
  );
}

function rowDateYmd_(row, headers, dateCol) {
  var col = findCol_(headers, dateCol);
  if (col < 0) {
    return "";
  }
  var ts = parseDateCell_(row[col]);
  if (!ts) {
    return "";
  }
  return Utilities.formatDate(
    new Date(ts),
    Session.getScriptTimeZone() || "Asia/Kolkata",
    "yyyy-MM-dd"
  );
}

function filterRowsByDateRange_(rows, headers, dateCol, dateFrom, dateTo) {
  dateFrom = normalizeDateFilter_(dateFrom);
  dateTo = normalizeDateFilter_(dateTo);
  if (!dateFrom && !dateTo) {
    return rows;
  }
  return rows.filter(function (row) {
    var ymd = rowDateYmd_(row, headers, dateCol);
    if (!ymd) {
      return false;
    }
    if (dateFrom && ymd < dateFrom) {
      return false;
    }
    if (dateTo && ymd > dateTo) {
      return false;
    }
    return true;
  });
}

function sortRowsByDateDesc_(rows, headers, dateCol) {
  var col = findCol_(headers, dateCol);
  if (col < 0) {
    return rows;
  }
  return rows.slice().sort(function (a, b) {
    return parseDateCell_(b[col]) - parseDateCell_(a[col]);
  });
}

function filterRowsByStatus_(rows, headers, statusValue) {
  statusValue = normalize_(statusValue);
  if (!statusValue) {
    return rows;
  }
  var cStatus = findCol_(headers, "Status");
  var cPay = findCol_(headers, "Payment Status");
  var wanted = statusValue.toLowerCase();
  return rows.filter(function (row) {
    var a = cStatus >= 0 ? normalize_(row[cStatus]).toLowerCase() : "";
    var b = cPay >= 0 ? normalize_(row[cPay]).toLowerCase() : "";
    return a === wanted || b === wanted;
  });
}

function getRegisterPageData_(sheetName, viewLocation, opts) {
  opts = opts || {};
  var page = Math.max(1, parseInt(opts.page, 10) || 1);
  var pageSize = Math.max(1, parseInt(opts.pageSize, 10) || REGISTER_PAGE_SIZE_);
  var dateCol = opts.dateCol || "Date";
  var dateFrom = opts.dateFrom || "";
  var dateTo = opts.dateTo || "";
  var status = opts.status || "";

  var data = getSheetData_(sheetName);
  data.rows = applySheetLocationFilter_(sheetName, data.rows, data.headers, viewLocation);
  data.rows = filterRowsByDateRange_(data.rows, data.headers, dateCol, dateFrom, dateTo);
  data.rows = filterRowsByStatus_(data.rows, data.headers, status);
  data.rows = sortRowsByDateDesc_(data.rows, data.headers, dateCol);

  var totalCount = data.rows.length;
  var totalPages = totalCount > 0 ? Math.ceil(totalCount / pageSize) : 1;
  if (page > totalPages) {
    page = totalPages;
  }
  var start = (page - 1) * pageSize;
  var pageRows = data.rows.slice(start, start + pageSize);
  var rows = pageRows.map(function (row) {
    return rowToObject_(data.headers, row);
  });

  return {
    headers: data.headers,
    rows: rows,
    sheetName: sheetName,
    count: totalCount,
    page: page,
    pageSize: pageSize,
    totalPages: totalPages,
    dateCol: dateCol,
    dateFrom: normalizeDateFilter_(dateFrom),
    dateTo: normalizeDateFilter_(dateTo),
    status: normalize_(status)
  };
}

function getRegisterPageForKey_(registerKey, viewLocation, opts) {
  var meta = REGISTER_KEYS_[registerKey];
  if (!meta) {
    throw new Error("Unknown register: " + registerKey);
  }
  opts = opts || {};
  opts.dateCol = meta.dateCol;
  opts.pageSize = opts.pageSize || REGISTER_PAGE_SIZE_;
  try {
    return getRegisterPageData_(meta.sheet, viewLocation, opts);
  } catch (e) {
    return {
      headers: [],
      rows: [],
      sheetName: meta.sheet,
      count: 0,
      page: 1,
      pageSize: REGISTER_PAGE_SIZE_,
      totalPages: 1,
      error: e.message || String(e)
    };
  }
}

function getRegisterPreviewForView_(sheetName, limit, viewLocation) {
  var dateCol = "Date";
  if (sheetName === KE.SHEETS.ORDER) {
    dateCol = "Order Date";
  }
  return getRegisterPageData_(sheetName, viewLocation, {
    page: 1,
    pageSize: REGISTER_PAGE_SIZE_,
    dateCol: dateCol
  });
}

function getRegisterViews_(viewLocation) {
  return {
    payments: getRegisterPageForKey_("payments", viewLocation, { page: 1 }),
    orders: getRegisterPageForKey_("orders", viewLocation, { page: 1 }),
    received: getRegisterPageForKey_("received", viewLocation, { page: 1 }),
    issues: getRegisterPageForKey_("issues", viewLocation, { page: 1 }),
    requests: getRegisterPageForKey_("requests", viewLocation, { page: 1 })
  };
}

function getRegisterPage(token, viewLocation, registerKey, page, dateFrom, dateTo, status) {
  validateSessionToken_(token);
  return sanitizeForClient_(getRegisterPageForKey_(registerKey, viewLocation, {
    page: page,
    dateFrom: dateFrom,
    dateTo: dateTo,
    status: status
  }));
}
