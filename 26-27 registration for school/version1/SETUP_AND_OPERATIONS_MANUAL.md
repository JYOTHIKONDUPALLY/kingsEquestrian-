# Kings Equestrian — School Registration & Stable Management  
## Setup and operations manual (multi-location)

**Version:** 2 · Academic year 2026–27  
**Audience:** Admins setting up Hyderabad, Pune, or Bangalore (or any new campus)  
**Purpose:** What to copy, what to change, and how to run the system without mixing locations or breaking forms.

---

## 1. How the system is built

Each **city / campus** should be a **separate stack**:

| Component | One per city? | Notes |
|-----------|---------------|--------|
| Google account (or Workspace user) | Yes | Keeps data and billing isolated |
| Google Spreadsheet (master workbook) | Yes | All tabs live here |
| Apps Script project (bound to spreadsheet) | Yes | Same code; different `config.js` |
| Registration Google Form | Yes | Linked to that city’s spreadsheet only |
| Payment Google Form | Yes | Linked to that city’s spreadsheet only |
| Web app deployment URL (`…/exec`) | Yes | Stable Management + My Rides portal |
| Brevo / email sender | Can share or split | `MAIL_FROM` must be verified for that sender |

**Rule:** Same **code files** can be pasted everywhere. **Never** share one spreadsheet or one set of form links across cities.

```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│ Registration    │────▶│                  │     │ Apps Script     │
│ Form (city A)   │     │  Spreadsheet     │◀────│ (bound project) │
└─────────────────┘     │  (city A only)   │     │ + config.js A   │
┌─────────────────┐     │                  │     └────────┬────────┘
│ Payment         │────▶│  Tabs: Riders,   │              │
│ Form (city A)   │     │  Schedule, Staff,│              ▼
└─────────────────┘     │  Horses, Feed…   │     ┌─────────────────┐
                        └──────────────────┘     │ Web app URL     │
                                                 │ (city A /exec)  │
                                                 └─────────────────┘
```

---

## 2. What you can copy unchanged vs what you must change

### 2.1 Copy the same in every city (code)

Paste all script files from the project folder, for example:

- `config.js` — **file yes, contents NO** (edit per city; see §3)
- `WebApp.js`, `AttendanceHTML.js`, `formhandlers.js`, `emails.js`, `MailService.js`
- `Schedule.js`, `TrainerAuth.js`, `Curriculumengine.js`, `PortalBackend.js`
- `Horses.js`, `StaffAttendance.js`, `StockInventory.js`, `OpsHistory.js`
- `SetupProject.js`, `DailySummary.js`, `PerfCache.js`, `Backup.js`, etc.
- HTML file in Apps Script named exactly **`AttendanceApp`** (paste from `AttendanceApp.html`)
- `appsscript.json` (keep timezone `Asia/Kolkata`)

### 2.2 Must change for each city (`config.js`)

Open **`config.js`** and set these for **that campus only**:

| Setting | Example Hyderabad | Example Pune | Example Bangalore |
|---------|-------------------|--------------|-------------------|
| `LOCATION_CODE` | `HYD` | `PUN` | `BLR` |
| `LOCATION_CITY` | `Hyderabad` | `Pune` | `Bengaluru` |
| `LOCATION_STATE` | `Telangana` | `Maharashtra` | `Karnataka` |
| `BUSINESS_ADDRESS` | Full address | Full address | Full address |
| `CONTACT_PHONE` | Campus phone | Campus phone | Campus phone |
| `CONTACT_EMAIL` | Campus email | Campus email | Campus email |
| `MAIL_FROM` | Same as verified Brevo sender | … | … |
| `MAIL_FROM_NAME` | e.g. Kings Equestrian Foundation | … | … |
| `UPI_ID` | Campus UPI | Campus UPI | Campus UPI |
| `PAYMENT_FORM_BASE_URL` | **This city’s** payment form short link | … | … |
| `PREFILL_ENTRY_IDS` | Entry IDs from **this** payment form | … | … |
| `ATTENDANCE_APP_URL` | **This** web app `/exec` URL | … | … |
| `MY_RIDES_PORTAL_URL` | Same URL + `?app=portal` | … | … |
| `DRIVE_ROOT_FOLDER` | e.g. `Kings Equestrian Hyderabad 26-27` | … | … |
| `STAMP_FILE_ID` | Receipt stamp in **this** Drive | … | … |
| `SIGN_FILE_ID` | Receipt signature in **this** Drive | … | … |
| `SCHOOL_PROGRAM_INFO_DOC_ID` | Brochure / program PDF link | … | … |
| `BACKUP_FOLDER_ID` | Optional Drive folder for backups | … | … |

If registration or payment form **column order** differs from the template, also update:

- `REG_COLS` — registration form columns (0-based indexes)
- `PAYMENT_COLS` — payment form columns (0-based indexes)

Use menu **Diagnose Payment Columns** after the first test payment.

### 2.3 Set in Script Properties (each project)

Apps Script → **Project settings** (gear) → **Script properties**:

| Property | Required? | Purpose |
|----------|-----------|---------|
| `BREVO_API_KEY` | Recommended | Transactional email via Brevo; falls back to Gmail if missing |

Do **not** put API keys inside `config.js` if you can avoid it.

### 2.4 Data in the spreadsheet (each city)

Fill per campus, not in code:

| Sheet | What to configure |
|-------|-------------------|
| **Mail Info** | Admin emails, daily summary, welcome CC, receipt CC |
| **TRAINERS** | Trainer logins (menu: Setup Trainers Sheet) |
| **service** (pricing) | Program names and prices |
| **GROOMERS** | Staff list for attendance |
| **HORSES**, **FEED_STOCK**, **TACK_STOCK** | Campus operations (or add via app) |

---

## 3. Google Forms — registration and payment

### 3.1 Registration form

1. Create a **new** Google Form in the **city’s** Google account.
2. Set questions to match your campus (student name, parent, grade, section, phone, email, program, address, consent date, etc.).
3. **Responses** → Link to **this city’s spreadsheet** (create or select).
4. Response sheet name should be **`Registration Response`** (or add alias in `REGISTRATION_SHEET_ALIASES` in `config.js`).
5. On first submit, the script will:
   - Generate a **KE number** (e.g. `KE2607271234`)
   - Write it back to the sheet
   - Send **welcome email** with consent PDF (if mail is configured)

### 3.2 Payment form

1. Create a **separate** payment form per city.
2. Link responses to tab **`Payment Responses 26-27`** (name in `CONFIG.SHEETS.PAYMENT_FORM`).
3. Put the form’s short URL in `PAYMENT_FORM_BASE_URL`.
4. On submit, the script sends a **payment receipt** email (when configured).

**Important:** Payment form columns on the current template are roughly:

Timestamp → Registration No → Phone → Amount → Screenshot → Payment Date → Transaction ref → PAN/Aadhaar → Mode of Payment → Payment For → (script columns for receipt)

If your form order differs, run **Diagnose Payment Columns** and adjust `PAYMENT_COLS`.

### 3.3 Prefilled payment links

If welcome emails include a pre-filled payment link, fill `PREFILL_ENTRY_IDS` with each field’s **entry ID** from the Google Form URL when you pre-fill manually once and copy the `entry.xxxxx` parameters.

---

## 4. New campus — step-by-step setup

Use this checklist **in order** for Hyderabad, Pune, or Bangalore.

### Phase A — Spreadsheet and code

1. Create a **new Google Spreadsheet** in the campus account.
2. **Extensions → Apps Script** → paste all `.gs` / `.js` files.
3. Create HTML file **`AttendanceApp`** → paste contents of `AttendanceApp.html`.
4. Edit **`config.js`** for this city (§2.2).
5. Save all files.

### Phase B — Sheets and auth

6. Open the spreadsheet → menu **Indus Equestrian** (or Kings Equestrian) → **Setup All Sheets/Tabs**.
7. Run **Authorize script — Docs + mail (run once)** and approve permissions.
8. Fill **Mail Info** sheet with admin and notification emails.
9. Run **Setup Trainers Sheet** — creates trainer logins for the attendance app.
10. (Optional) **Setup Training System Tabs** if you use curriculum / assessments.

### Phase C — Forms

11. Create and link **registration form** → test one submission.
12. Create and link **payment form** → test one submission.
13. Confirm welcome email and receipt email (or check **Email Log** tab).

### Phase D — Triggers and web app

14. Menu → **Setup All Triggers** (registration, payment, daily summary, weekly ops, backup, cache warm, etc.).
15. **Deploy → New deployment → Web app**
    - Execute as: **Me**
    - Who has access: **Anyone** (or your org policy)
16. Copy the **`/exec`** URL into `config.js`:
    - `ATTENDANCE_APP_URL` = `https://script.google.com/.../exec`
    - `MY_RIDES_PORTAL_URL` = same + `?app=portal`
17. Save `config.js` → **Deploy → Manage deployments → Edit → New version → Deploy** (URL stays the same).

### Phase E — Go live

18. Share registration and payment form links with parents (campus-specific only).
19. Share attendance app URL with trainers (campus-specific only).
20. Do **not** run **Seed Demo Attendance Data** on production.

---

## 5. Updating code after go-live (same URL)

You do **not** need a **new** deployment (new URL) for normal updates.

| Goal | Action | URL changes? |
|------|--------|--------------|
| Update HTML or scripts | Save files → **Manage deployments → Edit → New version → Deploy** | **No** |
| First-time publish | **New deployment → Web app** | Yes (save this URL in `config.js`) |
| Test latest saved code quickly | Use **Test deployments** URL (`/dev`) | Separate test link |

After UI changes, trainers may need a **hard refresh** (Ctrl+F5) or to clear browser cache.

---

## 6. Stable Management app (trainer PWA) — tabs overview

After login, trainers see (in order):

| Tab | Purpose |
|-----|---------|
| **Dashboard** | Overview: riders, horses, staff, stock alerts, care overdue, sessions |
| **Sessions** | Today’s / selected date classes, attendance, booking, group actions |
| **Staff** | Daily attendance (default Present; Absent locks the day), leave, HR fields |
| **Riders** | Rider list, curriculum progress, payments context |
| **Feed** | Feed inventory — **Regular** (daily auto-estimate) vs **Occasional** (vitamins; log Use/Expire only) |
| **Tack** | Tack inventory — restock, wear-out, transfer |
| **Horses** | Stable register, vaccination/deworming/farrier, **+ Care type** (shoeing, dental…), **Activity** timeline |
| **History** | Cross-module log with search and filters (default last 30 days) |

**History on cards:** Feed/Tack/Horses show **last 10** movements in modals; full search is on the History tab.

---

## 7. Automated emails and reports

Configured via **Setup All Triggers** and **Mail Info**:

| Job | When | Content |
|-----|------|---------|
| Welcome email | On registration form submit | KE number, consent PDF, payment link |
| Payment receipt | On payment form submit | Receipt PDF |
| Daily admin summary | Daily ~8 PM | Registrations, payments, sessions |
| Weekly ops summary | Monday 7 AM | Staff absences/leaves, horses, feed/tack, sessions |
| Weekly training report | Monday 8 AM | Training / curriculum (if enabled) |
| Drive backup | Daily 11 PM | Spreadsheet backup copy |

Manual test: menu **Send Weekly Ops Summary Now**, **Send Daily Summary Now**, **Test Email**.

---

## 8. Multi-location — common mistakes to avoid

| Mistake | What goes wrong | Fix |
|---------|-----------------|-----|
| Same payment form URL in all cities | Payments recorded in wrong sheet / wrong receipts | Unique form + URL per city in `config.js` |
| Hyderabad `ATTENDANCE_APP_URL` in Pune config | Pune trainers hit Hyderabad data | Deploy web app in Pune account; update URL |
| Forgot **Setup All Triggers** | No welcome email, no receipts, no backups | Run once per new project |
| Forms linked to wrong spreadsheet | Data in another city’s workbook | Re-link form responses to correct sheet |
| Column order changed on form without updating `REG_COLS` / `PAYMENT_COLS` | Wrong fields read; emails fail silently | Match indexes; use Diagnose Payment Columns |
| Pasted code but not **AttendanceApp** HTML file | App blank or errors on load | HTML file name must be exactly `AttendanceApp` |
| **New deployment** every update | Parents/trainers get new URLs | **Edit existing deployment → New version** |
| Shared Brevo sender not verified | Emails bounce | Verify domain/sender in Brevo for `MAIL_FROM` |
| Demo seed on production | Fake riders/sessions | Never run seed on live sheet |

---

## 9. Troubleshooting

| Symptom | Likely cause | What to do |
|---------|--------------|------------|
| Welcome email not sent | Triggers missing / Brevo / Mail Info | Setup All Triggers; check Email Log; Test Email |
| KE number not written | Form not linked or wrong sheet name | Link form; check `Registration Response` tab name |
| Receipt wrong amount or missing | Payment column mismatch | Diagnose Payment Columns; fix `PAYMENT_COLS` |
| Attendance app “Invalid token” on load | Old deployment or huge HTML wrapped wrong | Use `AttendanceApp` HTML file + `buildAttendanceApp_()`; redeploy version |
| Trainer cannot log in | No row in TRAINERS | Setup Trainers Sheet; check username/password |
| Horse vaccination saved but not in Activity | Old bug fixed in latest `Horses.js` | Update code; open Horses tab to backfill |
| Tab slow to load | Large sheets; first load fetches server | Normal on first open; use Refresh when data changed |
| Weekly email not received | Mail Info empty / trigger missing | Fill Mail Info admin rows; Setup All Triggers |

---

## 10. Per-city setup checklist (printable)

**City:** _______________  
**Google account:** _______________  
**Spreadsheet URL:** _______________  
**Web app URL:** _______________  

- [ ] All script files pasted  
- [ ] `AttendanceApp` HTML file created  
- [ ] `config.js` updated (location, email, UPI, form URLs, web app URLs)  
- [ ] Script property `BREVO_API_KEY` set (if using Brevo)  
- [ ] Setup All Sheets/Tabs run  
- [ ] Authorize script run  
- [ ] Mail Info filled  
- [ ] Registration form created and linked  
- [ ] Payment form created and linked  
- [ ] Test registration → KE no + welcome email  
- [ ] Test payment → receipt email  
- [ ] Setup Trainers Sheet  
- [ ] Setup All Triggers  
- [ ] Web app deployed; URLs saved in config; deployment version updated  
- [ ] Forms and app links shared with campus staff only  

---

## 11. Related guides (share with users)

| Guide | Audience | Files |
|-------|----------|--------|
| **Trainer & Stable Manager Guide** | Trainers / yard staff | `TRAINER_STABLE_MANAGER_GUIDE.md` / `.docx` |
| **My Rides Portal Guide** | Parents & riders | `MY_RIDES_PORTAL_GUIDE.md` / `.docx` |

The My Rides guide explains login, sessions, booking, shop, **order status**, and **payment status** in plain language. Share the `.docx` (or PDF export) with parents.

---

## 12. File reference (quick)

| File | Role |
|------|------|
| `config.js` | **Per-city settings**, sheet names, column maps |
| `WebApp.js` | Web app entry (`doGet`) — attendance + portal |
| `AttendanceHTML.js` | Builds HTML from `AttendanceApp` file |
| `AttendanceApp.html` | Trainer UI (paste into Apps Script HTML file) |
| `formhandlers.js` | Registration + payment submit handlers |
| `emails.js` | Welcome, receipt, consent PDF |
| `SetupProject.js` | Creates all spreadsheet tabs |
| `TrainerAuth.js` | Trainer login |
| `Horses.js`, `StaffAttendance.js`, `StockInventory.js`, `OpsHistory.js` | V2 stable operations |
| `ShoppingOrders.js` | Shop catalog, orders, daily shop report |
| `RidersPortalHTML.js` | My Rides portal UI |
| `appsscript.json` | Timezone, OAuth scopes, web app settings |

---

## 13. Support notes for developers

- **KE numbers** are unique per spreadsheet (`KE` + date + random). For cross-city reporting later, consider adding `LOCATION_CODE` into `generateKENo()` in `config.js`.
- Code uses `SpreadsheetApp.getActiveSpreadsheet()` — the project must stay **bound** to the campus spreadsheet.
- OAuth scopes include Drive, Docs, Calendar, Mail, External requests (Brevo).
- For clasp users: `clasp push` updates the project; you still **update deployment version** in the UI for web app changes to go live.

---

*Document generated for Kings Equestrian Foundation · School Registration & Stable Management V2 · 2026–27*
