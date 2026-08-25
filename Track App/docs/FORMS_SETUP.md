# Google Forms Setup (Optional)

Staff procurement runs through the **dashboard** (vendor request → pay vendor → place order → receive goods).

Optional Google Forms for **goods received** or **internal issue** can call the same Apps Script functions as the dashboard. Use the web app for RBAC instead of raw form triggers when possible.

## Goods Received Form (optional)

**Fields:** Order ID, Item, Qty, Location, Storage location, Received by  

Trigger calls `receiveGoods(payload)` with mapped fields.

## Issue Form (optional)

**Fields:** Request ID, Issued to, Qty, Location  

Trigger calls `issueItem(payload)`.

## Connecting forms

1. Create each form in Google Forms.
2. **Responses → Link to Sheets** (same workbook or linked sheet).
3. In Apps Script: **Triggers → Add trigger** → function above → **From spreadsheet → On form submit**.

## Notes

- Parent / customer order forms are **not** part of this inventory app — orders are captured on other platforms; this app tracks **vendor procurement** only.
- Validations (payment before order, receipt before issue) remain in `InventoryCore.gs`.
