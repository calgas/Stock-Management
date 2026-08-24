# CALGAS Capacitors — Stock Management

A stock management system for Raw Materials (RM) and Finished Goods (FG), built entirely on Google Sheets and Google Apps Script, with an installable PWA frontend. No external hosting, database, or server required for the backend — the spreadsheet *is* the database.

**Current version:** 2.5.0

---

## What it does

- Tracks RM and FG inventory: master item data, a full transaction ledger, and live current-stock balances with automatic Low/Critical status.
- Lets people post stock transactions (Receipt, Issue, Adjustment, Return for RM; Production, Dispatch, Adjustment, Return for FG) against any item, with department, reference, and supplier/invoice details on receipts.
- Gives every item a searchable, date-filterable transaction history, exportable to a styled Excel workbook.
- Bulk-imports new RM or FG items from an Excel template instead of one-by-one entry.
- Sends an optional daily email (per person, per RM/FG) with a fully styled current-stock workbook attached.
- Role-based access (Admin / Manager / Supervisor) with per-department scoping for Supervisors.
- Installs as a PWA on desktop or mobile, with a branded splash screen and offline-friendly app shell.

---

## Architecture

```
┌─────────────────────┐         ┌──────────────────────────┐
│   index.html (PWA)   │  HTTPS  │   Code.gs (Apps Script)  │
│  static, hosted      │ ──────► │   bound to the Sheet,    │
│  anywhere (GitHub    │  POST   │   deployed as a Web App  │
│  Pages, Firebase,    │ ◄────── │                          │
│  Drive, etc.)         │  JSON   │                          │
└─────────────────────┘         └────────────┬─────────────┘
        │                                     │
        │ manifest.json + sw.js               │ reads/writes
        │ (installability, app-shell cache)    ▼
        │                          ┌──────────────────────────┐
        └─────────────────────────►│ Stock_Management.xlsx     │
                                    │ (the actual data store)   │
                                    └──────────────────────────┘
```

- **`index.html`** is a single-file PWA (vanilla JS, no build step, no framework). It's hosted as static files completely separately from the Apps Script project — Apps Script here is *only* the JSON API, never serves HTML.
- **`Code.gs`** is bound to the spreadsheet (deployed from *Extensions → Apps Script* inside the Sheet itself, not a standalone script) and deployed as a Web App. It's the only thing that ever touches the spreadsheet directly.
- All communication is a single `doPost` endpoint accepting a JSON body (posted as `text/plain` to sidestep Apps Script's lack of CORS preflight support) with an `action` field that routes to the right handler.
- **`sw.js`** caches the app shell (HTML/CSS/JS) for fast repeat loads and offline resilience, but every request to the backend is always network-only — stock data is never served stale.

---

## Project structure

| File | Purpose |
|---|---|
| `Code.gs` | The entire backend: auth, sessions, RM/FG CRUD, transactions, ledger, bulk import, daily mail, admin tools. |
| `index.html` | The entire frontend: login/auth screens, dashboard, RM/FG stock pages, item history, Team & Access, all modals. |
| `manifest.json` | PWA manifest (name, icons, theme colors). |
| `sw.js` | Service worker — app-shell caching, network-only for API calls. |
| `Stock_Management.xlsx` | The Google Sheet itself — the data store (see [Spreadsheet structure](#spreadsheet-structure)). |

---

## Spreadsheet structure

| Sheet | Purpose |
|---|---|
| `RM_Master` / `FG_Master` | One row per item: code, name, category/family, unit, reorder level, active flag, ID No. |
| `RM_Stock_Ledger` / `FG_Stock_Ledger` | Every transaction ever posted, append-only. |
| `RM_Current_Stock` / `FG_Current_Stock` | Live computed balances, kept in sync incrementally on every transaction (and fully rebuildable via Recalculate). |
| `Users` | Accounts: name, username, salted password hash, role, department, email, active flag, daily-mail subscriptions. |
| `Sessions` | Active login tokens (auto-created, swept nightly). |
| `PasswordResets` | Pending OTP codes for password resets (auto-created, swept nightly). |
| `AuditLog` | Every add/edit/deactivate/login/admin action, with before/after values where relevant. |
| `Settings` | Key-value config — currently just `APP_SECRET`, used to hash passwords and sign session tokens. |
| `Departments` | The department list used in transaction forms and Supervisor scoping. |

Master data and Current Stock are separate on purpose: Current Stock is a derived, rebuildable view (via **Recalculate**), while Master is the source of truth for what items exist.

---

## Roles & permissions

| | Admin | Manager | Supervisor |
|---|:---:|:---:|:---:|
| Post transactions | ✅ | ✅ | ✅ (own department only) |
| View item history / export | ✅ | ✅ | ✅ |
| Add / edit master items, bulk import | ✅ | ✅ | ❌ |
| Deactivate / reactivate items | ✅ | ❌ | ❌ |
| Recalculate stock | ✅ | ❌ | ❌ |
| Manage people (Team & Access) | ✅ | ❌ | ❌ |

Every permission is enforced **server-side** (`requireRole_` / department-scope checks in `Code.gs`) — the frontend hiding a button is a convenience, not the actual security boundary.

---

## Setup

### 1. Spreadsheet + Apps Script

1. Open `Stock_Management.xlsx` in Google Sheets (or import it as one).
2. **Extensions → Apps Script**, paste in `Code.gs`.
3. Run `ADMIN_generateAppSecret()` once from the Apps Script editor. This sets `Settings!APP_SECRET`, used for password hashing and session signing — without it, login will fail with an explicit error telling you to run this.
4. **Deploy → New deployment → Web app.** Execute as *Me*, accessible to *Anyone* (the app handles its own auth on top of this). Copy the deployment URL.
5. Run `ADMIN_installNightlyCleanupTrigger()` once — sweeps expired sessions and OTP codes nightly so those sheets don't grow unbounded.
6. If you want the daily stock report emails: run `ADMIN_installDailyStockMailTrigger()` once. This needs Google Drive access (to generate `.xlsx` attachments) and Gmail send access, so it'll prompt a fresh authorization the first time — expected, not an error.
7. Create your first Admin user directly in the `Users` sheet (Username, Name, Role=`Admin`, Active=`TRUE`, a real Email), then run `ADMIN_setUserPassword('their-username', 'a-temporary-password')` to set their initial password.

### 2. Frontend

1. In `index.html`, set `GOOGLE_API_URL` to the Web App URL from step 4 above.
2. Host `index.html`, `manifest.json`, `sw.js` and `logo/` together as static files — GitHub Pages off this repo's `main` branch serves them at `https://calgas.github.io/Stock-Management/`. They don't need to be anywhere near the Apps Script project.
3. Open the hosted URL, log in as the Admin user you just created, and change the temporary password via **Forgot password** (this also verifies the OTP email path is working).

### 3. Ongoing accounts

New users are created from **Team & Access** in the app — that flow emails them a password-setup code automatically; there's no need to touch `ADMIN_setUserPassword` again except as a break-glass fallback if someone's email is unreachable.

---

## Notable design decisions

- **Passwords are salted** (`Salt` column per user), with automatic migration: any account created before salting was added keeps working, and gets silently upgraded to a salted hash the next time it logs in successfully — no bulk migration step was needed.
- **History lookups scan backward from the newest ledger rows in chunks**, stopping once enough matches are found, rather than reading the entire ledger on every click — this keeps item history fast regardless of how large the overall ledger grows.
- **Excel output (exports and daily mail) is generated entirely server-side** via native Google Sheets formatting, then exported to real `.xlsx` — not via a client-side JS library. The free tier of the usual browser-side Excel library can't write cell colors/fills/borders at all; doing it server-side means real styling (branded headers, status color-coding, banded rows) with no such limitation, and guarantees exports and mail look identical since they share the same styling code.
- **Bulk import creates each item's Current Stock row directly** at import time, seeded from an Opening Stock column — items are visible in the Stock tables immediately rather than only appearing after their first transaction.
- **The PWA app shell is cached, but stock data never is** — `sw.js` serves `index.html`/static assets cache-first for fast repeat loads, while every request to the Apps Script backend is always network-only.

---

## Known limitations / not yet done

- **PWA icons and the wordmark are served from `logo/` in this repo** via `raw.githubusercontent.com`, which is also what the sibling apps point at — this repo is the single source of truth for the family's brand assets.
- **No incremental recalculation** — `Recalculate` does a full rebuild of Current Stock from the entire ledger. This is intentional (a rare, manual, correctness-first admin tool), and is fine at normal usage; a very large ledger (20,000+ rows) will get a heads-up that it may take a while, since Apps Script executions have a hard 6-minute ceiling.
- **Gmail send quota** applies to OTPs and daily report emails — ~100/day on a consumer account, 1,500/day on Workspace. Irrelevant at a handful of subscribers, worth knowing if that list grows a lot.
- **Sending as `noreply@calgas.in` requires that address to be a verified "Send mail as" alias** on the Google account owning the Apps Script project. `sendMail_` checks `GmailApp.getAliases()` at runtime; if the alias is absent, mail still sends, just under the owner's own address.
- **Bulk import is capped at 500 rows per file**; split larger imports into batches.
