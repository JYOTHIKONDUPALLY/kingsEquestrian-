/**
 * Weekly inventory report for Track App v2.
 *
 * Email (admins): inventory by location, items issued (summary),
 * current stock, what to order from vendors.
 * Excel attachment: who received issues this week (student list — not in the mail body).
 *
 * Run installWeeklyReportTrigger() once (Mondays 08:00 IST),
 * or KE Track v2 menu → Send weekly report now.
 */

var KE_V2_REPORT_FN_ = "sendWeeklyInventoryReport";

function installWeeklyReportTrigger() {
  uninstallWeeklyReportTrigger();
  ScriptApp.newTrigger(KE_V2_REPORT_FN_)
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
  return "Weekly report trigger installed (Mondays at 08:00 IST).";
}

function uninstallWeeklyReportTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === KE_V2_REPORT_FN_) {
      ScriptApp.deleteTrigger(tr);
      removed++;
    }
  });
  return removed + " trigger(s) removed.";
}

function sendWeeklyReportNow(token) {
  assertAdmin_(token);
  return sanitizeForClient_(sendWeeklyInventoryReport());
}

function sendWeeklyInventoryReport() {
  var recipients = getAdminEmailRecipients_();
  if (!recipients.length) {
    return { sent: 0, message: "No weekly report recipients configured." };
  }

  var report = buildWeeklyReportPayload_();
  var html = renderWeeklyReportHtml_(report);
  var subject = "Kings Equestrian – Weekly inventory report (" +
    report.period.start + " to " + report.period.end + ")";

  var attach = buildWeeklyIssueExcel_(report);
  var mail = {
    to: recipients.join(","),
    subject: subject,
    htmlBody: html,
    name: "KE Track v2"
  };
  if (attach) {
    mail.attachments = [attach];
  }

  MailApp.sendEmail(mail);

  return {
    sent: recipients.length,
    recipients: recipients,
    attachedIssueRows: report.issueRows.length,
    message: "Weekly report sent to " + recipients.length + " recipient(s)."
  };
}

function getAdminEmailRecipients_() {
  var list = KE.WEEKLY_REPORT_RECIPIENTS || [];
  var out = [];
  list.forEach(function (email) {
    email = normalize_(email).toLowerCase();
    if (email && email.indexOf("@") > 0 && out.indexOf(email) < 0) {
      out.push(email);
    }
  });
  return out;
}

function parseDateMs_(v) {
  if (v == null || v === "") {
    return 0;
  }
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
    return v.getTime();
  }
  var s = String(v).trim();
  var dt = new Date(s);
  if (!isNaN(dt.getTime())) {
    return dt.getTime();
  }
  // yyyy-MM-dd or yyyy-MM-dd HH:mm
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}))?/);
  if (m) {
    return new Date(
      Number(m[1]), Number(m[2]) - 1, Number(m[3]),
      Number(m[4] || 0), Number(m[5] || 0)
    ).getTime();
  }
  return 0;
}

function getItemMasterMap_() {
  var data = getSheetData_(KE.SHEETS.ITEM);
  var map = {};
  data.rows.forEach(function (row) {
    var o = rowToObject_(data.headers, row);
    var code = normalize_(o["Item Code"]);
    if (code) {
      map[code] = o;
    }
    var name = normalize_(o["Item Name"]);
    if (name && !map[name]) {
      map[name] = o;
    }
  });
  return map;
}

function buildWeeklyReportPayload_() {
  var tz = Session.getScriptTimeZone() || "Asia/Kolkata";
  var now = new Date();
  var weekStart = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  var fmt = function (d) {
    return Utilities.formatDate(d, tz, "yyyy-MM-dd");
  };
  var period = {
    start: fmt(weekStart),
    end: fmt(now),
    startMs: weekStart.getTime()
  };

  var itemMap = getItemMasterMap_();
  var inventory = getInventoryRowsForView_("");
  var totalUnits = 0;
  var lowStock = [];
  var byLocation = {};
  KE.LOCATIONS.forEach(function (loc) {
    byLocation[loc] = { location: loc, rows: 0, qty: 0, low: 0, issuedQty: 0, issuedCount: 0 };
  });

  inventory.forEach(function (r) {
    var qty = Number(r["Current Qty"]) || 0;
    var loc = resolveLocationName_(r["Location"]);
    var code = normalize_(r["Item Code"]);
    var master = itemMap[code] || itemMap[normalize_(r["Item Name"])] || {};
    r.vendor = normalize_(master["Vendor"] || r.vendor);
    r.category = normalize_(master["Category"] || "");
    totalUnits += qty;
    if (!byLocation[loc]) {
      byLocation[loc] = { location: loc || "—", rows: 0, qty: 0, low: 0, issuedQty: 0, issuedCount: 0 };
    }
    byLocation[loc].rows += 1;
    byLocation[loc].qty += qty;
    if (r.stockStatus === "Low stock" || r.stockStatus === "Out of stock") {
      byLocation[loc].low += 1;
      lowStock.push(r);
    }
  });

  inventory.sort(function (a, b) {
    var loc = String(a["Location"]).localeCompare(String(b["Location"]));
    if (loc !== 0) {
      return loc;
    }
    return String(a["Item Name"]).localeCompare(String(b["Item Name"]));
  });
  lowStock.sort(function (a, b) {
    return (Number(a["Current Qty"]) || 0) - (Number(b["Current Qty"]) || 0);
  });

  var toOrder = [];
  lowStock.forEach(function (r) {
    var qty = Number(r["Current Qty"]) || 0;
    var min = Number(r.minLevel) || 0;
    var need = min > qty ? (min - qty) : 0;
    if (qty <= 0 && need < 1) {
      need = Math.max(min, 1);
    } else if (qty > 0 && qty <= min && need < 1) {
      need = min;
    }
    toOrder.push({
      itemCode: r["Item Code"],
      itemName: r["Item Name"],
      location: r["Location"],
      vendor: r.vendor || "—",
      currentQty: qty,
      minLevel: min,
      orderQty: need,
      status: r.stockStatus
    });
  });

  var pendingRequests = getRequestRowsForView_("").filter(function (r) {
    var st = normalize_(r["Status"]);
    return st === KE.REQUEST_STATUS.PENDING || st === KE.REQUEST_STATUS.ORDERED;
  });

  var issueData = getSheetData_(KE.SHEETS.ISSUE);
  var issueRows = [];
  var issuedByItem = {};
  for (var i = 0; i < issueData.rows.length; i++) {
    var obj = rowToObject_(issueData.headers, issueData.rows[i]);
    var ms = parseDateMs_(issueData.rows[i][findCol_(issueData.headers, "Date")] || obj["Date"]);
    if (!ms || ms < period.startMs) {
      continue;
    }
    issueRows.push(obj);
    var item = normalize_(obj["Item"]) || "—";
    var loc = resolveLocationName_(obj["Location"]) || "—";
    var key = item + "||" + loc;
    if (!issuedByItem[key]) {
      issuedByItem[key] = { item: item, location: loc, qty: 0, count: 0 };
    }
    issuedByItem[key].qty += Number(obj["Qty"]) || 0;
    issuedByItem[key].count += 1;
    if (byLocation[loc]) {
      byLocation[loc].issuedQty += Number(obj["Qty"]) || 0;
      byLocation[loc].issuedCount += 1;
    }
  }
  issueRows.sort(function (a, b) {
    return String(b["Date"] || "").localeCompare(String(a["Date"] || ""));
  });

  var issuedSummary = Object.keys(issuedByItem).map(function (k) {
    return issuedByItem[k];
  }).sort(function (a, b) {
    return b.qty - a.qty;
  });

  var issuedQtyTotal = 0;
  issuedSummary.forEach(function (x) { issuedQtyTotal += x.qty; });

  return {
    period: period,
    totals: {
      totalUnits: totalUnits,
      inventoryRows: inventory.length,
      lowStockCount: lowStock.length,
      issuedCount: issueRows.length,
      issuedQty: issuedQtyTotal,
      toOrderCount: toOrder.length
    },
    locationInventory: KE.LOCATIONS.map(function (loc) {
      return byLocation[loc];
    }),
    currentStock: inventory,
    stockByItem: buildStockByItem_(inventory),
    lowStock: lowStock,
    toOrder: toOrder,
    pendingRequests: pendingRequests,
    issuedSummary: issuedSummary,
    issueRows: issueRows
  };
}

function buildWeeklyIssueExcel_(report) {
  var rows = report.issueRows || [];
  var ss;
  try {
    var name = "KE_Weekly_Issues_" + report.period.start + "_to_" + report.period.end +
      "_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "HHmmss");
    ss = SpreadsheetApp.create(name);
    var sh = ss.getSheets()[0];
    sh.setName("Issued to students");
    var headers = [
      "Date", "Issue ID", "Item", "Qty", "Student Name", "Student_ID",
      "Location", "Issued By"
    ];
    var values = [headers];
    rows.forEach(function (r) {
      values.push([
        r["Date"] || "",
        r["Issue ID"] || "",
        r["Item"] || "",
        r["Qty"] || 0,
        r["Student Name"] || "",
        r["KE Number"] || "",
        r["Location"] || "",
        r["Issued By"] || ""
      ]);
    });
    if (values.length === 1) {
      values.push(["—", "No issues in this period", "", "", "", "", "", ""]);
    }
    sh.getRange(1, 1, values.length, headers.length).setValues(values);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#1a2235")
      .setFontColor("#C9A84C");
    sh.setFrozenRows(1);
    SpreadsheetApp.flush();

    var blob = DriveApp.getFileById(ss.getId())
      .getAs("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    blob.setName("KE_Weekly_Issue_List_" + report.period.start + "_to_" + report.period.end + ".xlsx");

    try {
      DriveApp.getFileById(ss.getId()).setTrashed(true);
    } catch (e2) { /* ignore */ }
    return blob;
  } catch (e) {
    Logger.log("Issue Excel attach failed: " + e);
    try {
      if (ss) {
        DriveApp.getFileById(ss.getId()).setTrashed(true);
      }
    } catch (e3) { /* ignore */ }
    return buildWeeklyIssueCsv_(report);
  }
}

function buildWeeklyIssueCsv_(report) {
  var lines = ["Date,Issue ID,Item,Qty,Student Name,Student_ID,Location,Issued By"];
  (report.issueRows || []).forEach(function (r) {
    var cells = [
      r["Date"], r["Issue ID"], r["Item"], r["Qty"],
      r["Student Name"], r["KE Number"], r["Location"], r["Issued By"]
    ].map(function (v) {
      var s = String(v == null ? "" : v);
      if (s.indexOf(",") >= 0 || s.indexOf('"') >= 0) {
        return '"' + s.replace(/"/g, '""') + '"';
      }
      return s;
    });
    lines.push(cells.join(","));
  });
  return Utilities.newBlob(lines.join("\n"), "text/csv",
    "KE_Weekly_Issue_List_" + report.period.start + "_to_" + report.period.end + ".csv");
}

function escHtml_(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function renderWeeklyReportHtml_(r) {
  var th = function (txt) {
    return '<th style="padding:8px 10px;border-bottom:2px solid #C9A84C;text-align:left;font-size:12px;color:#1A2235">' + txt + "</th>";
  };
  var td = function (txt, align) {
    return '<td style="padding:7px 10px;border-bottom:1px solid #eee;font-size:13px;' +
      (align ? "text-align:" + align + ";" : "") + '">' + (txt == null || txt === "" ? "—" : txt) + "</td>";
  };

  var html = '<div style="font-family:Arial,Helvetica,sans-serif;max-width:720px;margin:0 auto;color:#222;background:#fff">';
  html += '<div style="background:#1A2235;color:#F0EAD6;padding:18px 20px">';
  html += '<div style="font-size:13px;color:#C9A84C;letter-spacing:0.06em;text-transform:uppercase">Kings Equestrian</div>';
  html += '<h2 style="margin:4px 0 0;color:#fff;font-weight:600">Weekly inventory report</h2>';
  html += '<p style="margin:8px 0 0;color:#9BACC4;font-size:13px">Period: <strong style="color:#F0EAD6">' +
    escHtml_(r.period.start) + "</strong> to <strong style=\"color:#F0EAD6\">" + escHtml_(r.period.end) + "</strong></p>";
  html += "</div>";

  html += '<table style="width:100%;border-collapse:collapse;margin:16px 0"><tr>';
  [
    [r.totals.totalUnits, "Units on hand"],
    [r.totals.lowStockCount, "Low / out of stock"],
    [r.totals.issuedQty, "Units issued (7 days)"],
    [r.totals.toOrderCount, "Lines to order"]
  ].forEach(function (k) {
    html += '<td style="padding:10px;background:#f7f4ea;text-align:center;width:25%">';
    html += '<div style="font-size:22px;font-weight:700;color:#1A2235">' + k[0] + "</div>";
    html += '<div style="font-size:11px;color:#666;text-transform:uppercase">' + k[1] + "</div></td>";
  });
  html += "</tr></table>";

  html += '<h3 style="color:#1A2235;margin:20px 0 8px">Inventory by location</h3>';
  html += '<table style="border-collapse:collapse;width:100%">';
  html += "<tr>" + th("Location") + th("Stock rows") + th("Units") + th("Low / out") + th("Issued (7 days)") + "</tr>";
  (r.locationInventory || []).forEach(function (loc) {
    html += "<tr>" +
      td(escHtml_(loc.location)) +
      td(loc.rows, "right") +
      td(loc.qty, "right") +
      td(loc.low, "right") +
      td((loc.issuedQty || 0) + " units / " + (loc.issuedCount || 0) + " issues", "right") +
      "</tr>";
  });
  html += "</table>";

  html += '<h3 style="color:#1A2235;margin:20px 0 8px">Total qty by item / size</h3>';
  html += '<p style="font-size:12px;color:#666;margin:0 0 8px">Each item (including size variants) with total on-hand and split by location.</p>';
  html += '<table style="border-collapse:collapse;width:100%">';
  html += "<tr>" + th("Item") + th("Code") + th("Total qty");
  KE.LOCATIONS.forEach(function (loc) { html += th(loc); });
  html += "</tr>";
  (r.stockByItem || []).forEach(function (it) {
    html += "<tr>" +
      td(escHtml_(it.itemName)) +
      td(escHtml_(it.itemCode)) +
      td("<strong>" + (it.totalQty || 0) + "</strong>", "right");
    KE.LOCATIONS.forEach(function (loc) {
      html += td((it.locations && it.locations[loc]) || 0, "right");
    });
    html += "</tr>";
  });
  html += "</table>";

  html += '<h3 style="color:#1A2235;margin:20px 0 8px">Items issued (last 7 days)</h3>';
  html += '<p style="font-size:12px;color:#666;margin:0 0 8px">Summary only. Full student list is in the Excel attachment — not listed in this email.</p>';
  if (!(r.issuedSummary || []).length) {
    html += '<p style="color:#666">No issues recorded in this period.</p>';
  } else {
    html += '<table style="border-collapse:collapse;width:100%">';
    html += "<tr>" + th("Item") + th("Location") + th("Qty issued") + th("Issue count") + "</tr>";
    r.issuedSummary.forEach(function (x) {
      html += "<tr>" +
        td(escHtml_(x.item)) +
        td(escHtml_(x.location)) +
        td(x.qty, "right") +
        td(x.count, "right") +
        "</tr>";
    });
    html += "</table>";
  }

  html += '<h3 style="color:#1A2235;margin:20px 0 8px">Current stock levels</h3>';
  html += '<table style="border-collapse:collapse;width:100%">';
  html += "<tr>" + th("Item") + th("Code") + th("Location") + th("Qty") + th("Min") + th("Status") + "</tr>";
  (r.currentStock || []).forEach(function (s) {
    var statusColor = s.stockStatus === "Out of stock" || s.stockStatus === "Low stock"
      ? "color:#b42318;font-weight:600" : "color:#1A2235";
    html += "<tr>" +
      td(escHtml_(s["Item Name"])) +
      td(escHtml_(s["Item Code"])) +
      td(escHtml_(s["Location"])) +
      td(s["Current Qty"], "right") +
      td(s.minLevel, "right") +
      '<td style="padding:7px 10px;border-bottom:1px solid #eee;font-size:13px;' + statusColor + '">' +
        escHtml_(s.stockStatus) + "</td>" +
      "</tr>";
  });
  html += "</table>";

  html += '<h3 style="color:#1A2235;margin:20px 0 8px">What to order from vendors</h3>';
  html += '<p style="font-size:12px;color:#666;margin:0 0 8px">Based on current qty at or below min level (or out of stock).</p>';
  if (!(r.toOrder || []).length) {
    html += '<p style="color:#666">Nothing below min level right now.</p>';
  } else {
    html += '<table style="border-collapse:collapse;width:100%">';
    html += "<tr>" + th("Vendor") + th("Item") + th("Location") + th("On hand") + th("Min") + th("Order qty") + th("Status") + "</tr>";
    r.toOrder.forEach(function (o) {
      html += "<tr>" +
        td(escHtml_(o.vendor)) +
        td(escHtml_(o.itemName)) +
        td(escHtml_(o.location)) +
        td(o.currentQty, "right") +
        td(o.minLevel, "right") +
        td("<strong>" + o.orderQty + "</strong>", "right") +
        td(escHtml_(o.status)) +
        "</tr>";
    });
    html += "</table>";
  }

  if ((r.pendingRequests || []).length) {
    html += '<h3 style="color:#1A2235;margin:20px 0 8px">Open requests (already raised)</h3>';
    html += '<table style="border-collapse:collapse;width:100%">';
    html += "<tr>" + th("Request ID") + th("Vendor") + th("Item") + th("Qty") + th("Location") + th("Status") + "</tr>";
    r.pendingRequests.forEach(function (q) {
      html += "<tr>" +
        td(escHtml_(q["Request ID"])) +
        td(escHtml_(q["Vendor"])) +
        td(escHtml_(q["Item"])) +
        td(q["Qty"], "right") +
        td(escHtml_(q["Location"])) +
        td(escHtml_(q["Status"])) +
        "</tr>";
    });
    html += "</table>";
  }

  html += '<p style="margin:22px 0 8px;font-size:12px;color:#888">Attachment: Excel list of students who received items this week (name, Student_ID, item, qty, location, issued by).</p>';
  html += '<p style="font-size:11px;color:#aaa">Sent automatically every Monday 08:00 IST · KE Track v2</p>';
  html += "</div>";
  return html;
}
