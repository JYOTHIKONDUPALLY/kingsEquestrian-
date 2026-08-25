# Kings Equestrian — Version 2 Enhancement Log

**Project:** School Registration & Attendance System (Version 2)  
**Client:** Kings Equestrian Foundation  
**Base:** Version 1 (2026–27 academic year)  
**V2 kickoff:** 21 Jul 2026  
**Document owner:** Development team  
**Last updated:** 26 Jul 2026  

---

## Purpose

This document is the single source of truth for **Version 2** work: what was built, when, and how long it took. Use it to:

- Track progress across modules
- Prepare invoices / bills with dated line items
- Hand off context to future developers

**Rule:** Add a new log entry **the same day** work is done (or when a feature is marked complete).

---

## Version 2 scope

| # | Module | Description | Status |
|---|--------|-------------|--------|
| 1 | **Staff tab** | Staff / groomer directory, HR fields, daily attendance, leave credits & leave ranges, uniform issued date, edit on card | In progress |
| 2 | **Horses tab** | Stable register, built-in + custom care types, activity history, photos, trainer/groom assignment | In progress |
| 3 | **Dashboard enhancement** | Ops overview: riders, horses, staff, stock alerts, care status alerts, feed days-left, training outcomes | In progress |
| 4 | **Feed tab** | Feed inventory: Regular vs Occasional usage, restock / consume / use / expire | Done |
| 5 | **Tack tab** | Tack inventory UI: add with category/model/vendor/location/image; wear-out / restock / transfer | Done |
| 6 | **History tab** | Cross-module activity feed with search + filters; per-tab history capped at last 10 | Done |
| 7 | **Weekly ops email** | Monday digest to admins: staff, horses, feed, tack, sessions | Done |

**Status key:** `Planned` · `In progress` · `Done` · `On hold`

---

## Billing defaults

Fill these in once; they apply to all entries unless overridden per row.

| Field | Value |
|-------|-------|
| Currency | INR (₹) |
| Hourly rate | _TBD_ |
| Billing cycle | _TBD_ (e.g. fortnightly / monthly) |
| Billable by default | Yes |

---

## Module summary (for quick billing)

Roll-up of logged hours by module. Update when adding entries.

| Module | Est. hours | Logged hours | Status | Notes |
|--------|------------|--------------|--------|-------|
| Staff tab | — | 2.0 | In progress | Uniform issued + card edit icon + search |
| Horses tab | — | 3.0 | In progress | Care status Done / Not Done / Postponed |
| Dashboard enhancement | — | 1.5 | In progress | Care + feed days-left + stock location |
| Feed tab | — | 9.0 | Done | Full tab, movements, missed-day defaults, stock-count rule, kg conversion, archiving |
| Tack tab | — | 5.0 | Done | Full tab + wear-out / restock / transfer |
| Cross-cutting / config | — | 1.5 | In progress | Sheet schemas + migrations |
| UI/UX pass (all V2 tabs) | — | 3.5 | Done | Search, filters, history, empty/error states |
| Data-integrity fixes | — | 1.5 | Done | Feed ledger atomicity + tack transfer/photo fixes |
| **Total** | — | **27.0** | | |

---

## Chronological timeline

High-level milestones. Detail lives in the **Work log** below.

| Date | Milestone |
|------|-----------|
| 21 Jul 2026 | V2 enhancement log created; scope defined (Staff, Horses, Dashboard, Feed, Tack) |
| 21 Jul 2026 | Baseline inventory: Staff, Horses, Dashboard, and stock backend already present in codebase from prior V2 work |
| 25 Jul 2026 | Horses care status (Done / Not Done / Postponed); Staff uniform issued + edit icon |
| 25 Jul 2026 | Feed tab + Tack tab with inventory movements; dashboard reflections |
| 25 Jul 2026 | Feed missed-day default consume with confirm/edit workflow |
| 25 Jul 2026 | UI/UX pass (search, filters, movement history, empty/error states) + feed ledger integrity fixes |
| 25 Jul 2026 | Feed long-gap stock count rule, locked units with kg conversion, movement archiving |
| _TBD_ | V2 sign-off / deployment |

---

## Work log

Copy the template block for each session or completed item.

### Entry template

```markdown
### YYYY-MM-DD — [Module] — Short title

| Field | Value |
|-------|-------|
| **Date** | YYYY-MM-DD |
| **Module** | Staff / Horses / Dashboard / Feed / Tack / Cross-cutting |
| **Type** | Feature / Enhancement / Bug fix / Refactor / Docs |
| **Hours** | 0.0 |
| **Billable** | Yes / No |
| **Status** | Done / In progress |
| **Files** | `File1.js`, `File2.js` |

**Description:**  
What was delivered and why.

**Acceptance / test notes:**  
How it was verified.

**Invoice line (optional):**  
One-line text suitable for a bill PDF.
```

---

### 21 Jul 2026 — Cross-cutting — V2 documentation & baseline inventory

| Field | Value |
|-------|-------|
| **Date** | 2026-07-21 |
| **Module** | Cross-cutting |
| **Type** | Docs |
| **Hours** | 0.5 |
| **Billable** | Yes |
| **Status** | Done |
| **Files** | `VERSION_2_ENHANCEMENT_LOG.md` |

**Description:**  
Created Version 2 enhancement log with scope, timeline, billing fields, and work-log template. Documented existing V2-related code already in the repository.

**Acceptance / test notes:**  
N/A — documentation only.

**Invoice line:**  
V2 project documentation — enhancement log and billing tracker setup.

---

### Pre-log — Staff tab — Staff directory & attendance (baseline)

| Field | Value |
|-------|-------|
| **Date** | _Estimate date when work was done_ |
| **Module** | Staff |
| **Type** | Feature |
| **Hours** | _TBD_ |
| **Billable** | Yes |
| **Status** | In progress |
| **Files** | `StaffAttendance.js`, `AttendanceHTML.js`, `config.js` |

**Description:**  
- **Staff tab** in attendance app (`data-pane="staff"`)
- GROOMERS sheet: expanded HR schema (Aadhaar, bank details, profile / passport / passbook photos, leave balance)
- GROOMER_ATTENDANCE + GROOMER_LEAVES sheets
- Daily attendance: Present / Absent / Leave with monthly leave credits (4/month)
- Add / edit / remove staff; apply leave date ranges
- Photo upload (gallery / camera) for profile, passport, passbook
- Legacy sheet migration from older 9-column layout
- Auto sheet setup on trainer login (`ensureGroomerAttendanceSetup`)

**Acceptance / test notes:**  
Verify on trainer login → Staff tab → mark attendance, add employee, apply leave.

**Invoice line:**  
V2 — Staff tab: directory, HR fields, attendance marking, and leave management.

---

### Pre-log — Horses tab — Stable register (baseline)

| Field | Value |
|-------|-------|
| **Date** | _Estimate date when work was done_ |
| **Module** | Horses |
| **Type** | Feature |
| **Hours** | _TBD_ |
| **Billable** | Yes |
| **Status** | In progress |
| **Files** | `Horses.js`, `AttendanceHTML.js`, `config.js` |

**Description:**  
- **Horses tab** in attendance app (`data-pane="horses"`)
- HORSES sheet: full register (ID, breed, DOB/age, trainer, groom, status, lease, weight, care dates, vet notes, photo)
- Status filter: Active, Leased, Rehab, Lame, Retired
- Care tracking: vaccination, deworming, farrier — overdue / due-soon indicators
- Add / edit / retire horse; photo upload to Drive
- Trainer & groomer dropdowns populated from active staff
- Legacy HORSES sheet migration

**Acceptance / test notes:**  
Verify Horses tab → add horse, set care dates, filter by status, edit and retire.

**Invoice line:**  
V2 — Horses tab: stable register, care schedule tracking, and photo management.

---

### Pre-log — Dashboard — Operations overview (baseline)

| Field | Value |
|-------|-------|
| **Date** | _Estimate date when work was done_ |
| **Module** | Dashboard |
| **Type** | Enhancement |
| **Hours** | _TBD_ |
| **Billable** | Yes |
| **Status** | In progress |
| **Files** | `Curriculumengine.js`, `AttendanceHTML.js` |

**Description:**  
Extended dashboard beyond training outcomes:
- Total riders; riders present vs scheduled today
- Total horses + breakdown by status (Active, Leased, Rehab, Lame, Retired)
- Total staff; staff present; staff on leave today
- Feed & tack stock summaries (item count, low-stock count, top items)
- Unified alerts: lame/rehab horses, overdue care, low leave balance, low feed/tack stock
- Training outcomes section retained (bookings, pass/repeat, level progress)

**Acceptance / test notes:**  
Dashboard tab → Refresh; confirm counts match sheets and alerts appear for low stock / overdue care.

**Invoice line:**  
V2 — Dashboard enhancement: unified ops metrics, care alerts, and stock summaries.

---

### Pre-log — Feed & Tack — Backend & dashboard integration (partial)

| Field | Value |
|-------|-------|
| **Date** | _Estimate date when work was done_ |
| **Module** | Feed + Tack |
| **Type** | Feature (partial) |
| **Hours** | _TBD_ |
| **Billable** | Yes |
| **Status** | In progress |
| **Files** | `StockInventory.js`, `config.js`, `Curriculumengine.js`, `AttendanceHTML.js` |

**Description:**  
- FEED_STOCK and TACK_STOCK sheets with seed data (hay, pellets, saddles, helmets, etc.)
- `ensureStockSetup` on login; `getStockSummaries_` for read API
- Dashboard shows feed/tack summaries and low-stock alerts
- **Not yet done:** dedicated **Feed tab** and **Tack tab** in the app for add/edit/adjust stock from the UI

**Acceptance / test notes:**  
Edit FEED_STOCK / TACK_STOCK in spreadsheet → Dashboard reflects counts and low-stock alerts.

**Invoice line:**  
V2 — Feed & tack inventory backend and dashboard integration (UI tabs pending).

---

## Planned work (not started)

Use this checklist; move items to the Work log when started or completed.

- [x] **Feed tab** — list, add, edit, adjust quantity, min level, consumed/day from app UI
- [x] **Tack tab** — add/list with category/model/vendor/location/image; wear-out / restock / transfer
- [x] **Horses care status** — Done / Not Done / Postponed
- [x] **Staff uniform issued** + edit icon on card
- [ ] **Staff tab** — remaining polish (reports, export, search/filter)
- [ ] **Horses tab** — remaining polish (bulk import, care reminders email)
- [ ] **Dashboard** — charts, date-range filters, printable summary
- [ ] **V2 deployment** — clasp push, trainer UAT, production sign-off

---

## Billing period summary

Duplicate this section per invoice period.

### Period: _YYYY-MM-DD to YYYY-MM-DD_

| Date | Module | Description | Hours | Rate (₹) | Amount (₹) |
|------|--------|-------------|-------|----------|------------|
| | | | | | |
| | | **Subtotal** | | | **0** |

**Notes for invoice:**  
_Optional: PO number, payment terms, deliverables summary._

---

## Related files (Version 2)

| Area | Primary files |
|------|----------------|
| Staff | `StaffAttendance.js`, `AttendanceHTML.js` |
| Horses | `Horses.js`, `AttendanceHTML.js` |
| Dashboard | `Curriculumengine.js` (`getTrainingDashboardData`), `AttendanceHTML.js` |
| Feed / Tack | `StockInventory.js`, `AttendanceHTML.js`, `config.js` (FEED_STOCK, FEED_MOVEMENTS, TACK_STOCK, TACK_MOVEMENTS) |
| Config / sheets | `config.js`, `SetupProject.js` |

---

## How to update this document

1. **After each work session:** add a Work log entry (use template above).
2. **Update** Module summary hours and Chronological timeline if a milestone is reached.
3. **Change status** in the scope table when a module moves to Done.
4. **At billing time:** copy rows from Work log into Billing period summary; export or attach this file to the invoice.

---

### 25 Jul 2026 — Feed — Missed EOD consume defaults + confirm/edit

| Field | Value |
|-------|-------|
| **Date** | 2026-07-25 |
| **Module** | Feed |
| **Type** | Enhancement |
| **Hours** | 2.0 |
| **Billable** | Yes |
| **Status** | Done |
| **Files** | `StockInventory.js`, `AttendanceHTML.js`, `Curriculumengine.js` |

**Description:**  
If daily EOD consume was skipped, Feed tab shows DEFAULT estimates copied from the last consume (or consumed/day rate), marked clearly as unconfirmed. Trainer must Confirm or Edit before stock is deducted. Dashboard warns about pending defaults and projected low stock.

**Invoice line:**  
V2 — Feed missed-day default consume with confirm/edit workflow.

---

### 25 Jul 2026 — Horses — Care status Done / Not Done / Postponed

| Field | Value |
|-------|-------|
| **Date** | 2026-07-25 |
| **Module** | Horses |
| **Type** | Enhancement |
| **Hours** | 3.0 |
| **Billable** | Yes |
| **Status** | Done |
| **Files** | `Horses.js`, `AttendanceHTML.js`, `config.js`, `Curriculumengine.js` |

**Description:**  
Vaccination, deworming, and farrier now each have status (Done / Not Done / Postponed) plus postponed-to date. Cards show status badges; dashboard alerts for not-done and overdue care.

**Invoice line:**  
V2 — Horses care status tracking (done / not done / postponed).

---

### 25 Jul 2026 — Staff — Uniform issued + edit icon

| Field | Value |
|-------|-------|
| **Date** | 2026-07-25 |
| **Module** | Staff |
| **Type** | Enhancement |
| **Hours** | 2.0 |
| **Billable** | Yes |
| **Status** | Done |
| **Files** | `StaffAttendance.js`, `AttendanceHTML.js`, `config.js` |

**Description:**  
Added Uniform Issued Date on staff form and card. Edit details available via ✎ icon on the staff card (and Edit details button).

**Invoice line:**  
V2 — Staff uniform issued date and edit-on-card controls.

---

### 25 Jul 2026 — Tack — Full tab with movements

| Field | Value |
|-------|-------|
| **Date** | 2026-07-25 |
| **Module** | Tack |
| **Type** | Feature |
| **Hours** | 5.0 |
| **Billable** | Yes |
| **Status** | Done |
| **Files** | `StockInventory.js`, `AttendanceHTML.js`, `config.js` |

**Description:**  
Tack tab: add items (name, category, model, vendor, location, qty, min level, image). List tracks qty per location. Movements: Restock (date + qty), Wear Out (date, reason, photo), Transfer (to another location). Sheets: TACK_STOCK, TACK_MOVEMENTS.

**Invoice line:**  
V2 — Tack tab with inventory, wear-out, restock, and transfer.

---

### 25 Jul 2026 — Feed — Full tab with consumption tracking

| Field | Value |
|-------|-------|
| **Date** | 2026-07-25 |
| **Module** | Feed |
| **Type** | Feature |
| **Hours** | 4.0 |
| **Billable** | Yes |
| **Status** | Done |
| **Files** | `StockInventory.js`, `AttendanceHTML.js`, `config.js` |

**Description:**  
Feed tab: add/edit feed (location, qty, unit, min level, consumed/day). Shows remaining qty and estimated days left. Movements: Restock and Consume. Sheets: FEED_STOCK, FEED_MOVEMENTS.

**Invoice line:**  
V2 — Feed tab with stock, daily consumption, and inventory entries.

---

### 25 Jul 2026 — Dashboard — Reflect all V2 stock & care updates

| Field | Value |
|-------|-------|
| **Date** | 2026-07-25 |
| **Module** | Dashboard |
| **Type** | Enhancement |
| **Hours** | 1.5 |
| **Billable** | Yes |
| **Status** | Done |
| **Files** | `Curriculumengine.js`, `AttendanceHTML.js` |

**Description:**  
Dashboard now surfaces care not-done/postponed alerts, feed days-left, location on stock rows, and staff uniform-not-issued info alerts.

**Invoice line:**  
V2 — Dashboard updates for care status, feed, and tack.

---

### 25 Jul 2026 — Feed — Missed EOD consume defaults + confirm/edit

| Field | Value |
|-------|-------|
| **Date** | 2026-07-25 |
| **Module** | Feed |
| **Type** | Enhancement |
| **Hours** | 2.0 |
| **Billable** | Yes |
| **Status** | Done |
| **Files** | `StockInventory.js`, `AttendanceHTML.js`, `Curriculumengine.js` |

**Description:**  
If daily EOD consume was skipped, Feed tab shows DEFAULT estimates copied from the last consume (or consumed/day rate), marked as unconfirmed. Trainer must Confirm or Edit before stock is deducted. Dashboard warns about pending defaults and projected low stock.

**Invoice line:**  
V2 — Feed missed-day default consume with confirm/edit workflow.

---

### 25 Jul 2026 — Cross-cutting — UI/UX pass on Feed, Tack, Horses, Staff & Dashboard

| Field | Value |
|-------|-------|
| **Date** | 2026-07-25 |
| **Module** | Feed, Tack, Horses, Staff, Dashboard |
| **Type** | Enhancement |
| **Hours** | 3.5 |
| **Billable** | Yes |
| **Status** | Done |
| **Files** | `AttendanceHTML.js`, `Horses.js`, `StockInventory.js`, `Curriculumengine.js` |

**Description:**  
Usability pass across the V2 tabs. Feed and Tack now have a sticky search bar with filter chips (low stock / needs confirm / out of stock, plus location and category), stock level bars, plain-English status tags, guided empty states, loading skeletons and error panes with a retry button. Added a movement history modal per item (wires up the previously unused `getFeedMovements` / `getTackMovements`). Feed and tack movement forms now explain each option, block future dates and over-issue, and the confirm-defaults modal shows a live running total, a resulting-stock preview and a per-day include/skip tick. Horses gained search plus care filters and an explicit "not confirmed" marker for statuses inferred from a date; horse care and staff uniform dates are now validated for consistency. Dashboard groups repetitive alerts, flags tack that is fully out of stock, reports hidden alerts when capped, and links straight to the relevant tab.

**Invoice line:**  
V2 — UI/UX improvements and validation hardening across Feed, Tack, Horses, Staff and Dashboard.

---

### 25 Jul 2026 — Feed & Tack — Data-integrity fixes

| Field | Value |
|-------|-------|
| **Date** | 2026-07-25 |
| **Module** | Feed, Tack |
| **Type** | Bug fix |
| **Hours** | 1.5 |
| **Billable** | Yes |
| **Status** | Done |
| **Files** | `StockInventory.js` |

**Description:**  
Fixed a case where confirming several missed feed days could write movement rows and then abort before updating the stock quantity, leaving the ledger and the stock figure out of step; the whole batch is now validated and costed before anything is written. Same-day consume entries are now summed instead of overwritten, the burn rate uses a 7-day rolling average instead of a single last value, duplicate and future-dated consume entries are rejected, `Adjust` rows record the change rather than the new total, tack transfers no longer share a Drive photo file that could be trashed, and destination checks are case-insensitive.

**Invoice line:**  
V2 — Feed/tack stock ledger integrity fixes.

---

### 25 Jul 2026 — Feed — Long-gap stock count, unit locking & movement archiving

| Field | Value |
|-------|-------|
| **Date** | 2026-07-25 |
| **Module** | Feed, Dashboard |
| **Type** | Enhancement |
| **Hours** | 3.0 |
| **Billable** | Yes |
| **Status** | Done |
| **Files** | `config.js`, `StockInventory.js`, `AttendanceHTML.js`, `Curriculumengine.js` |

**Description:**  
Three changes agreed after review.

1. **Long gaps now require a physical count.** Past 14 days with no entry the app stops estimating: the item is marked "Stock count needed", the confirm action is replaced with "Record stock count", and both the estimate API and the confirm-all action refuse it with an explanation. A recorded stock count re-baselines the item, so the missing days before it are settled and normal daily entry resumes. The dashboard raises this as a high-severity alert.
2. **Feed units are locked with kg conversion.** New `Pack Size (kg)` column on `FEED_STOCK` holds the weight of one bag or bale. The unit can no longer be edited after the item is saved (past movements are stored in it), and where a pack size is known the movement form offers an "Entering in: bags / kilograms" toggle with a live conversion preview; the converted figure and the original kg entry are both written to the movement note. Stock counts stay in the item's own unit, since an absolute figure in kg would be ambiguous.
3. **Movement archiving.** New `FEED_MOVEMENTS_ARCHIVE` and `TACK_MOVEMENTS_ARCHIVE` sheets. `archiveOldStockMovements()` files anything older than a year, exposed as an "Archive old movements" button on the Feed tab, and an automatic run fires at most once a day once the movement sheets pass 3,000 rows. Rows with an unreadable date are never archived, and the History view notes that older entries live in the archive sheet.

**Invoice line:**  
V2 — Feed stock-count enforcement, unit locking with kg conversion, and movement archiving.

---

### 26 Jul 2026 — Custom horse care, occasional feed, History tab & weekly ops email

| Field | Value |
|-------|-------|
| **Date** | 2026-07-26 |
| **Module** | Horses, Feed, Staff, History, Ops email |
| **Type** | Enhancement |
| **Hours** | 5.0 |
| **Billable** | Yes |
| **Status** | Done |
| **Files** | `AttendanceApp.html`, `Horses.js`, `StockInventory.js`, `StaffAttendance.js`, `OpsHistory.js`, `config.js` |

**Description:**  
1. **Horses — extensible care + activity log.** Trainers can add care types beyond Vaccination / Deworming / Farrier (e.g. Shoeing) and log them per horse with Done / Not Done / Postponed. Each horse card shows recent activity (leases, care, status) newest-first; Activity modal and History tab cover the rest.  
2. **Feed — Regular vs Occasional.** Occasional items (vitamins etc.) skip missed-day auto-estimates; movements support Use / Expire with optional horse, plus Restock / Adjust.  
3. **Staff — default Present.** Unmarked staff today show as Present; marking Absent locks Present/Leave for that day.  
4. **Tab order + History.** Dashboard → Sessions → Staff → Riders → Feed → Tack → Horses → History. History tab aggregates feed/tack/horses/staff/sessions with search and date filters; in-tab history modals show last 10 only.  
5. **Weekly admin email.** `sendWeeklyOpsAdminSummary` (Monday 7am trigger + menu) digests absences, leaves, leases, inventory in/out, and sessions for the prior week.

**Invoice line:**  
V2 — Custom horse care & activity history, occasional feed tracking, History tab, staff Present-default, weekly ops digest.
