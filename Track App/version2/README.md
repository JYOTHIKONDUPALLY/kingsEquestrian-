# Kings Equestrian Track App — Version 2 (Simplified)

A lighter Google Sheets + Apps Script inventory app based on the main Track App.

## What is included

| Area | Behaviour |
|------|-----------|
| **Vendors** | Add / list. Location can be a site **or All** (vendor for every location). |
| **Items** | Add / list. Location can be a site **or All**. Min level drives alerts. |
| **Inventory** | Qty available **per location**. Set / update count for a site. |
| **Requests** | Written to the **REQUEST_REGISTER** sheet tab (Excel-style register). Statuses: Pending → Ordered → Received / Cancelled. Optional stock add on Received. |
| **Issue** | Issue stock to a student. Dropdown shows **Student_Name - Student_ID** filtered by location. |
| **Students** | Same columns as `students-hyderabad` plus **Location**. Paste all students into `STUDENT_MASTER`. |
| **Low stock** | Overview alerts when `Current Qty ≤ Min Level` (or qty = 0). |
| **Admin / location UI** | Admin can switch location (including All). Managers/Trainers are locked to their home location. |

## What was removed (vs main Track App)

Payments, vendor orders/GRN pipeline, stock transfer, samples, pricing, analytics email, Drive image upload.

## Sheet tabs

- `VENDOR_MASTER`
- `ITEM_MASTER`
- `INVENTORY`
- `REQUEST_REGISTER`
- `ISSUE_REGISTER`
- `STUDENT_MASTER` — columns: `Student_ID`, `Student_Name`, `Parent_Name`, `Email`, `Phone`, `Program`, `Skill_Riding_Avg`, `Grade`, `Section`, `Location`
- `USER_MASTER`

Sites: **Bangalore**, **Hyderabad**, **Pune**, **Farm**. Masters may use **All**.

## Roles

| Role | Access |
|------|--------|
| **Admin** | Everything + location picker (All / each site) |
| **Manager** | Masters, inventory, requests, issue — own location |
| **Trainer** | Create / view requests — own location |

## Setup

1. Create a new Google Spreadsheet (or use a blank one).
2. **Extensions → Apps Script**, delete any default code.
3. Copy every file from `version2/apps-script/` into the Apps Script project (same filenames).
4. Save, then run **`setupInventorySheets`** (authorize when prompted).
5. Optionally run **`seedSampleData`** for demo rows.
6. **Deploy → New deployment → Web app**
   - Execute as: *Me*
   - Who has access: your domain / anyone in org
7. Open the web app URL → login.

### Demo logins (after seed)

| User | Password | Role |
|------|----------|------|
| `KE Admin` / `admin@ke.demo` | `admin@123` | Admin |
| `BLR Manager` | `manager@123` | Manager (Bangalore) |
| `Farm Trainer` | `trainer@123` | Trainer (Farm) |

## Typical flow

1. Add vendors (All or site-specific).
2. Add items (All or site-specific) with a min level.
3. Set inventory counts per location.
4. Create requests → they appear on **REQUEST_REGISTER**.
5. Mark request **Received** (optionally add stock).
6. **Issue** to student (by Order/Request ID or new Name + KE#).

## Note

Keep this project separate from the main Track App Apps Script project. Point each deployment at its own spreadsheet (run `setSpreadsheetId` from that sheet if needed).
