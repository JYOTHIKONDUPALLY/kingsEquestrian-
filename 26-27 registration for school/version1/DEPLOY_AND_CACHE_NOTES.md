# Deployment, Cache & Drive Images — What to Keep in Mind

## Why only *some* people see your update

This is usually **cache**, not a failed deploy.

| Layer | What happens |
|-------|----------------|
| **Google Apps Script `/exec` CDN** | Serves an old HTML snapshot to some users for minutes–hours |
| **Phone / browser cache** | Keeps the old page even after you published a new version |
| **Home-screen / PWA shortcut** | Often the stickiest — still opens the cached shell |
| **Wrong deployment URL** | You created a **New deployment** (new URL) while staff still open the **old** `/exec` link |

Server code (`google.script.run`) updates immediately. The **HTML shell** (buttons, payment form, layout) is what stays stale. That is why one phone looks updated and another does not.

---

## Correct update steps (URL stays the same)

1. Paste/save code in the Apps Script editor.
2. **Bump `CONFIG.APP_UI_VERSION`** in `config.js` (e.g. `2026-08-19d` → `2026-08-19e`).  
   Required — this is how phones detect “I am on an old page”.
3. **Deploy → Manage deployments → pencil (Edit)** on the **existing** Web app.  
   Do **not** create a New deployment unless you want a new URL.
4. Set **Version** → **New version** → **Deploy**.

### After deploying — what others must do

1. Fully close the tab / app (not just switch away).
2. Open the **same** campus `/exec` URL again.
3. Hard refresh: `Ctrl+Shift+R` (Windows) / `Cmd+Shift+R` (Mac).
4. Phone: Incognito once, or clear site data for `script.google.com`.
5. If a brown banner says **“A newer version is available”** → tap **Load update**.

### How to verify

- Stable Management header shows `v2026-08-19d` (or whatever you set) next to the trainer name.
- If that version is older than `CONFIG.APP_UI_VERSION`, they are still on a cached page.

---

## What we added in code

1. `CONFIG.APP_UI_VERSION` — bump on every deploy.
2. `getAppUiVersion()` — live version from the server.
3. Client auto-reload once when page version ≠ live version.
4. If reload still fails (aggressive CDN), a **Load update** banner appears.

---

## Drive image error (“Sorry, unable to open the file”)

Drive `/view` links are web pages, not image files. Shop photos are fetched with the script owner’s access and inlined when possible. Prefer **Anyone with the link → Viewer**, or host real `.jpg`/`.png` URLs.

---

## Deploy checklist for UI changes

- [ ] `config.js` — bump `APP_UI_VERSION`
- [ ] Paste changed files (`WebApp.js`, `AttendanceHTML.js`, `AttendanceApp.html`, `RidersPortalHTML.js`, …)
- [ ] Manage deployments → **Edit** → **New version** → Deploy
- [ ] Test in Incognito; confirm header shows the new `v…` stamp
