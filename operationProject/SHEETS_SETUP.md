# Kings Equestrian Ops — Sheet Setup (per readme)

Copy updated HTML/JS into Apps Script and deploy. On first routine save, these sheets are created automatically if missing.

## Required existing sheets
- Users, Horses, Trainers, Grooms, Locations
- Raw Data, Daily Inventory, Inventory Items
- ActivityData (legacy), HorseHealthData, MedicalHistory, VaccinationData, ShoeingData

## Auto-created sheets
### HorseProfile (master — one-time defaults per horse)
| Column | Purpose |
|--------|---------|
| Horse_ID | Serial / ID shown on routine form |
| Horse_Name | |
| Location, Trainer, Groom, Status | Assignments (Active / Rehabilitation / Leave) |
| Default_Wet_Grass … Default_Oats | Feed kg per serving (pre-fills 7AM slot) |
| Default_Water_Liters | Water L per serving |

### DailyRoutineLog (daily ops — 90% of work)
One row per horse per day: activities (×2), feed slots (7AM, 1PM, 6PM, Extra) + totals, water slots + total.

### inventory_transaction
Audit log when routine log reduces stock (Type = CONSUMPTION).

## Trainers / Grooms
Add optional **Status** column: `Active` or `Leave`.

## Item names in Raw Data / Inventory
Match feed labels for auto stock reduction:
- Wet Grass, Dry Grass, Feed (Mixed), Barley, Oats

## Main form
**Daily Routine Log** (`form-daily-routine`) — horses in left column, all fields in grid, saves batch + reduces inventory.
