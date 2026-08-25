# Kings Equestrian Inventory System

Google Sheets + Apps Script ERP-style inventory for Kings Equestrian (Bangalore, Hyderabad, Pune, Farm).

Built from `guidance/Kings_Equestrian_Final_Production_v3.docx` and `guidance/Kings_Equestrian_Inventory_System_Template.xlsx`.

## What's included

| Path | Purpose |
|------|---------|
| `apps-script/` | All Apps Script source (deploy into your Google Sheet) |
| `guidance/` | Original production doc, Excel template, starter README |
| `docs/SETUP.md` | Step-by-step deployment |
| `docs/RBAC.md` | Roles and permissions |
| `docs/FORMS_SETUP.md` | Optional Google Forms |

## Quick start

1. Upload `guidance/Kings_Equestrian_Inventory_System_Template.xlsx` to Google Drive → **Open with Google Sheets**.
2. **Extensions → Apps Script** — create one `.gs` file per file in `apps-script/` (or use [clasp](https://github.com/google/clasp) to push the folder).
3. Add `Dashboard.html` as an HTML file named `Dashboard`.
4. Run **`setupInventorySheets`** once, then **`setSpreadsheetId`**.
5. Add users to **USER_MASTER** with **Password** column (see `docs/RBAC.md`), or run **`seedSampleData()`**.
6. **Deploy → New deployment → Web app** — open `/exec` URL → login with USER_MASTER credentials (not Google).
7. Demo login: **KE Admin** / **admin@123** after sample data is loaded.

## Workflow

```
Procurement request (vendor + item + qty + location)
  → Pay vendor → Approved
  → Place vendor order
  → Receive goods (stock +)
  → Optional: Issue stock internally
```

Control rules enforced in code:

- No vendor payment → no approval  
- No approval → no order  
- No receipt → no issue  
- Cannot issue below available stock  

## Modules

- Inventory, vendor masters  
- Procurement requests, vendor payments, vendor orders, goods received, issue register  
- RBAC via USER_MASTER  
- Dashboard (summary + all actions)  
- Weekly email report (optional trigger)  

See `docs/SETUP.md` for full instructions.
