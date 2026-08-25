# Role-Based Access Control (RBAC)

## USER_MASTER columns

| Column | Description |
|--------|-------------|
| Email | Login identifier (e.g. `admin@ke.demo`) — **not** tied to Google sign-in |
| Name | Login identifier alternative (e.g. `KE Admin`) |
| Role | `Admin`, `Accounts`, `Trainer`, or `Manager` |
| Location | Required for Manager; optional for others |
| Password | Plain text or SHA-256 hash (run `hashUserPassword("KE Admin", "yourpass")` to hash) |

## Roles

| Role | Access |
|------|--------|
| **Admin** | Full access — all masters, requests, payments, orders, receive, issue |
| **Accounts** | Payments (and approval after payment); view summary; weekly report recipient |
| **Trainer** | Create requests; view summary |
| **Manager** | Location-scoped orders, receive, issue; master data when no location conflict |

## Enforcement

- `loginUser(identifier, password)` checks USER_MASTER and returns a session **token**.
- `validateSessionToken_(token)` loads the user for each API call.
- Dashboard hides tabs the user cannot use (`permissions` in bootstrap).

## Location scope

- On login, **USER_MASTER → Location** is the user’s home location (shown top-right).
- **Non-admin** users only see and create data for that location; location fields in forms are locked.
- **Admin** users get a **header dropdown** (default **Farm**, optional **All Locations** or any site).
- Payments without a Location column are filtered by linked **Request ID** location.

## Control rules (business logic)

These apply regardless of role (except Admin still must follow data integrity rules):

1. **NO PAYMENT → NO APPROVAL** — `recordPayment` must run before order placement.
2. **NO APPROVAL → NO ORDER** — `placeOrder` requires status `Approved`.
3. **NO RECEIPT → NO ISSUE** — `issueItem` requires a goods-received row for the request’s order.
4. **Stock check** — issue blocked if `Current Qty < issue qty`.

## Adding a user

1. Open **USER_MASTER**.
2. Append: Email, Name, Role, Location, Password.
3. User signs in on the login page with **Email or Name** + **Password**.

## Script functions

- `loginUser(identifier, password)` — returns `{ token, user }`
- `logoutUser(token)`, `hashUserPassword(emailOrName, plainPassword)`
- Session lasts 6 hours (cache + script property backup)
