# Kings Equestrian Inventory – Setup Guide

Follow these steps in order (from `guidance/README.txt`).

## 1. Create the spreadsheet

1. Open `guidance/Kings_Equestrian_Inventory_System_Template.xlsx` in Google Sheets (upload to Drive first if needed).
2. Confirm sheet tabs match:  
   `INVENTORY_MASTER`, `VENDOR_MASTER`,  
   `REQUEST_REGISTER`, `PAYMENTS`, `VENDOR_ORDERS`, `GOODS_RECEIVED`, `ISSUE_REGISTER`.
3. Add a tab **`USER_MASTER`** with headers:  
   `Email | Name | Role | Location | Password`  
   Or import **`guidance/Kings_Equestrian_Inventory_System_Template_WITH_SAMPLE.xlsx`** (includes demo rows).

## 2. Install Apps Script

### Option A – Manual copy

1. In the spreadsheet: **Extensions → Apps Script**.
2. Create script files matching `apps-script/*.gs` (same file names).
3. **+** → **HTML** → name it `Dashboard` → paste `apps-script/Dashboard.html`.
4. Replace `appsscript.json` content if prompted (timezone `Asia/Kolkata`, runtime V8).

### Option B – clasp

```bash
cd "Track App/apps-script"
clasp login
clasp create --type sheets --title "Kings Equestrian Inventory"
clasp push
```

Bind the script to your inventory spreadsheet in the Apps Script project settings.

## 3. First-time script setup

Run these functions once from the Apps Script editor (authorize when asked):

| Function | Purpose |
|----------|---------|
| `setupInventorySheets` | Creates missing tabs, headers, dropdown validations |
| `setSpreadsheetId` | Saves spreadsheet ID for web app / triggers |

## 4. Add users (login + RBAC)

Login uses **USER_MASTER** username/password — **not** your Google account.

| Email | Name | Role | Location | Password |
|-------|------|------|----------|----------|
| admin@ke.demo | KE Admin | Admin | Farm | admin@123 |

**Quick demo:** Run **`seedSampleData()`** in Apps Script (or menu **Kings Inventory → Load sample demo data**), then sign in with **KE Admin** / **admin@123**.

Roles: `Admin`, `Accounts`, `Trainer`, `Manager` — see `docs/RBAC.md`.

## 5. Deploy the dashboard

1. **Deploy → New deployment → Web app**
2. Execute as: **Me** (or deploying user)
3. Who has access: **Anyone** (or your organisation) — users sign in with USER_MASTER password
4. Open the `/exec` URL — you should see the **login** page first

From the sheet: **Kings Inventory → Open dashboard** (sidebar login) or **Open login (web app)**.

## 6. Seed master data (recommended)

1. **Vendors** — add vendors with short **Code** (used in item codes, e.g. `DEC`, `GPA`).
2. **Items** — item codes auto-generate: `KE-[VendorCode]-[Category]-[Model]`.

## 7. Test the workflow

1. **Create procurement request** (Trainer/Admin) — vendor, item, qty, location → `Pending`.
2. **Pay vendor** (Accounts) — approves request (`Approved`).
3. **Place order** (Manager/Admin) — vendor PO for location.
4. **Receive goods** — increases inventory qty.
5. **Issue item** (optional) — decreases inventory for internal use.

## 8. Weekly report (optional)

Run `installWeeklyReportTrigger()` once.  
Emails Admins and Accounts on Mondays at 8:00 AM (script timezone).

## Troubleshooting

| Issue | Fix |
|-------|-----|
| Access denied | Add your Google email to USER_MASTER |
| Missing sheet | Run `setupInventorySheets` |
| Web app blank | Redeploy web app; run `setSpreadsheetId` with sheet open |
| NO RECEIPT → NO ISSUE | Complete **Receive goods** for the request’s order first |
