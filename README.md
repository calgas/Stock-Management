# CALGAS Capacitors — Stock Management

A stock management system for Raw Materials (RM) and Finished Goods (FG), built entirely on Google Sheets and Google Apps Script, with an installable PWA frontend. No external hosting, database, or server required for the backend — the spreadsheet *is* the database.

**Current version:** 3.6.2
**Repo:** `calgas/Stock-Management` · **Hosted:** https://calgas.github.io/Stock-Management/

---

## What it does

- Tracks RM and FG inventory: master item data, a full transaction ledger, and live current-stock balances with automatic Low/Critical status.
- Lets people post stock transactions (Receipt, Issue, Adjustment, Return for RM; Production, Dispatch, Adjustment, Return for FG) against any item, with department, reference, and supplier/invoice details on receipts.
- Gives every item a searchable, date-filterable transaction history, exportable to a styled Excel workbook.
- Bulk-imports new RM or FG items from an Excel template instead of one-by-one entry.
- Sends an optional daily email (per person, per RM/FG) with a fully styled current-stock workbook attached.
- Role-based access (Admin / Manager / Supervisor) with per-department scoping for Supervisors.
- Installs as a PWA on desktop or mobile, with a branded splash screen and offline-friendly app shell.
- Detects and reports balance drift nightly, without silently correcting it.

---

## What changed in 3.6.2

**Scroll-to-top works on the grid pages again.** The button already existed, but it watched `.page`, `#appShell` and `window` — and 3.6.0 moved the scrolling out of all three into `.table-scroll`. So on exactly the long lists that need it, it never appeared. It now collects every scrollable view for the visible page, including its table body, and takes the deepest. Scroll listeners are bound once at setup rather than per navigation, so repeated page changes can't stack duplicates on the same element.

It was also sitting a single pixel above the pagination bar on desktop; moved from `bottom: 96px` to `124px` for real clearance.

---

## What changed in 3.6.1

Two fixes to the 3.6.0 grid work.

- **The sticky column header vanished behind a hovered row.** Hovered rows are raised to `z-index: 5` so their lift shadow sits over their neighbours; the new sticky header was only at `z-index: 3` and lost that contest. The header is now at 12.
- **Account controls jumped to the left on Overview.** The top bar uses `space-between`, and Overview hides the title slot — leaving one flex child, which `space-between` parks at the start. `.user-badge` now carries `margin-left: auto`, so it stays right whether or not the title is present.

---

## What changed in 3.6.0

The stock, ledger and team screens are now real data grids rather than long documents.

### The table scrolls, not the page

Previously the whole page scrolled and the table sat in a container as tall as all fifty rows, so its horizontal scrollbar was parked at the bottom of that container — you had to scroll the entire page to the end to reach it, and vertical scrolling moved the page rather than the rows.

Those pages (`.page-grid`: RM, FG, Ledger, Team) no longer scroll themselves. The title and filters are fixed, and the table body is the only scrolling region. The horizontal scrollbar therefore sits permanently at the bottom of the visible table, and the column headers stay in view while the rows move.

Two structural fixes were needed for this:

- **Pagination was inside `.table-scroll`**, so it scrolled away sideways with the rows. Table and footer are now siblings inside a `.grid-panel`; only the body scrolls.
- **`min-width` was set only inside the mobile media query.** In a narrowed desktop window the browser satisfied the width by crushing columns until product names broke one word per line — and because the table then technically fit, there was nothing to scroll. Tables now hold a 1060px floor at every breakpoint, with an explicit 220px minimum on the name column, which carries the longest values.

Also fixed: toolbar selects collapsed to just their chevron in a narrow window; they now hold a 150px floor.

### The tab name lives in the header

The top bar had unused space on the left and the page heading scrolled away with the content. The active tab's title and description now render in the top bar, read from the page's own `<h2>`/`<p>` so the heading is defined in one place and the two can't drift. Grid pages drop their in-page heading, which returns that vertical space to the table; Overview keeps its own and leaves the top bar clear rather than stating the title twice.

### Mobile

Mobile can't use the desktop flex-fill — the address bar resizes the viewport — so the table body gets a bounded `62vh` height instead. Same outcome: scrollbar at the bottom of the visible table, headers fixed. The top bar becomes two rows (brand and controls above, tab title below), and the filter row is sticky so it stays reachable while the rows scroll.

---

## What changed in 3.5.0

Interface release: mobile, notifications, and inline feedback.

### Tables scroll sideways on mobile

`.panel` sets `overflow: hidden` and the stock tables sat in a bare `<div>`, so a table wider than the panel was **clipped, not scrolled** — there was nothing to swipe. Every table now sits in a `.table-scroll` container, and tables keep a real `min-width` on mobile (720px for stock, 620px for the ledger) so they overflow and scroll instead of being compressed into unreadable columns. `overscroll-behavior-x: contain` stops a swipe that reaches the end of the table from triggering browser back-navigation, and scrollbars are given an explicit height so the area visibly announces itself as scrollable.

### Warnings when an issue exceeds available stock

A `.modal-warn` component (amber, triangle icon) now appears **directly under the Quantity field** as you type — not at the foot of the form, where the sticky action bar would cover it. It reads `Only 43.84 on hand — this is 56.16 more than available.` and never disables the submit button: the server owns that decision, and a stale cached balance must not block a valid transaction.

`.modal-error` had no CSS rule at all and rendered as unstyled text; it is now a matching red component. Both collapse when empty via `:empty`, so setting `textContent` is the whole API. Error and warning use different icon *shapes*, not just colours, so they stay distinguishable to colour-blind users and in direct sunlight.

### Toasts moved to top-centre

Bottom-right put them under the thumb on mobile and behind the nav rail. Each toast now has a coloured status rail and icon badge, a dismiss button, and a countdown bar so it's clear it will leave on its own. Errors last 6s rather than 3.8s — long enough to actually read a stock message. `toast(msg, true)` still means error; `toast(msg, 'warn' | 'info' | 'success')` reaches the new variants.

### Mobile interface

- **Safe-area insets** on the nav rail, sheets and toasts, plus `viewport-fit=cover` on the viewport meta — without that meta, `env(safe-area-inset-*)` silently resolves to zero and every such rule does nothing.
- **16px inputs**, which is the threshold below which iOS zooms the page on focus and forces a manual pinch-out after every field.
- **44px minimum tap targets** on icon buttons and nav links.
- **Two-up toolbar filters** instead of one control per row; the toolbar was taller than the phone before any data loaded.
- **Sticky sheet header and actions** — the form scrolls, not the whole sheet, so Cancel/Save stay reachable without scrolling to the end of a long form.

---

## What changed in 3.4.0

### Stock can no longer go negative

Nothing prevented issuing more than was on hand, which is how balances like `-43.84 KG` and `-1,380 NOS` appeared. Those are not display bugs — the ledger really does record more going out than came in.

`addTransaction_` now refuses any movement that would take a balance below zero. The check runs **inside the existing lock, before the ledger append**, because a check outside it can be raced: two Issues posted in the same moment could each read a sufficient balance and both commit.

What is and isn't blocked:

| Case | Result |
|---|---|
| Issue/Dispatch within the balance | allowed |
| Issue/Dispatch of the **entire** balance | allowed |
| Issue/Dispatch beyond the balance, or at zero | **blocked** |
| Receipt / Production / Return | always allowed |
| Positive Adjustment | always allowed — this is how you repair an existing negative |
| Negative Adjustment past zero | **blocked** |

The overdraw comparison uses the same `DRIFT_TOLERANCE` as the drift check, so issuing 56.56 out of a 56.56 balance isn't rejected by float noise.

Bulk import applies the same floor: a negative Opening column is now skipped with a reason rather than seeding a balance nobody could then correct.

`ALLOW_NEGATIVE_STOCK` (default `false`) exists as a deliberate escape hatch if a situation ever genuinely calls for it.

**The existing negative balances are not fixed by this**, and Recalculate won't fix them either — it rebuilds from the ledger, and the ledger is what says the stock left. Correct each one by entering the Receipt that was missed, or by posting a positive Adjustment.

### The transaction form shows what's available

The Quantity hint now reads the item's balance (`Available: 118.32.`, or `Nothing on hand to remove — post a Receipt first.`), and a warning appears as you type past it rather than after you submit. It deliberately does **not** disable the submit button: the server owns the decision, and a stale cached balance must not lock someone out of a transaction that is actually valid.

### History view uses the screen

The ledger modal was on the 760px `.modal-wide` frame, which squeezed six columns until dates, departments and usernames each wrapped onto two lines. It now uses `.modal-table` (1180px), with date/qty/dept/user set to `nowrap` and the table body growing to `min(58vh, 640px)` instead of a fixed 420px. On a 1600px viewport every row is a single line and the inner scrollbar is gone.

---

## What changed in 3.3.1

**Form modals now use the width of a desktop screen.** The Stock Entry, Edit Person and Item modals were locked to a 460px frame regardless of viewport, so on a wide screen they rendered as a tall narrow column that scrolled vertically — and the Edit Person footer, which carries up to five buttons, overflowed that frame and forced a *horizontal* scrollbar inside the modal.

- Those three modals now use `.modal-form` (860px) with a two-column field grid above 780px, collapsing to a single column below it. The mobile bottom-sheet behaviour is unchanged.
- Footer rows (`.modal-actions`, `.modal-actions-split`, `.modal-actions-right`) now wrap, so no future button label can bring the sideways scroll back.
- `.btn-primary` is full-width by design — correct as a thumb target in the mobile sheet, wrong in a desktop footer, where it wrapped onto its own line beneath Cancel. It is now auto-width in modal footers above 780px only.
- The item modal fills `#itemFieldsMain` via `innerHTML`, leaving its fields one level too deep to be grid items. `.grid-passthrough` (`display: contents`) dissolves that wrapper so generated fields line up beside the static Code field.

The frame is capped at 860px rather than filling the viewport deliberately: a field stretched across 1900px puts its label and input at opposite ends of the eye's travel, which reads as harder to use, not roomier. Two columns at ~390px is the width a field is actually comfortable at, and it roughly halves the form's height — which is what removes the vertical scrollbar.

---

## What changed in 3.3.0

This release is correctness, performance and supply-chain work. No new screens.

### Duplicate transactions can no longer be created by a retry

The frontend retries a request that times out. A timeout only means the *response* was lost — the write may well have committed. The retry then posted the same stock movement a second time, which corrupted balances quietly and only on bad Wi-Fi, so it looked random.

Fixed on both sides:

- Each transaction submit generates a `clientTxnId` once, **per submit, not per network attempt**, so a retry carries the same ID.
- `addTransaction_` looks that ID up in the ledger *inside the existing lock* before appending. On a match it returns the original result with `duplicate: true` rather than posting again.
- The retry itself is now gated: only reads, or writes carrying a `clientTxnId`, are ever retried. Any other mutating action fails fast instead of gambling.

The lookup scans the most recent `IDEMPOTENCY_SCAN_ROWS` (300) ledger rows rather than the whole sheet — a duplicate can only ever arrive seconds after the original, and scanning everything would make every post cost O(ledger size).

This adds a `ClientTxnId` column to both ledgers, created automatically on first use.

### Session validation no longer reads the whole Sessions sheet

`validateSession_` runs on **every authenticated request** and was reading the entire Sessions sheet each time, so the cost grew with every login ever issued. Tokens are now cached in `CacheService` (`SESSION_CACHE_TTL_SECONDS`, 6h — the platform maximum), with the sheet consulted only on a miss.

The sheet stays authoritative. The cache is written on login, dropped on logout, and its TTL is capped by the session's own stored expiry, so a cached entry can never extend a session beyond what the sheet recorded. Cache failures are swallowed — an optimisation must never fail a request.

### Nightly balance drift detection

Balances are maintained incrementally by `applyStockDelta_`, which is fast but can diverge from the ledger — a half-completed write, a manual edit in the sheet, a restored row. The divergence is silent; the numbers still look plausible.

`nightlyDriftCheck_` now runs as part of the existing nightly cleanup trigger. It recomputes balances from the ledger in memory, compares against stored `CurrentBalance`, and emails every active Admin a table of any mismatches beyond `DRIFT_TOLERANCE` (0.001, to absorb float noise from fractional Kg/Ltr quantities).

**It deliberately does not auto-correct.** An automatic fix on a nightly trigger could overwrite a legitimate manual correction and would destroy the evidence of whatever caused the drift. A human reviews, then uses **Recalculate stock**. Silence means no drift — it does not email on a clean run.

### PWA icons are real icons now

The manifest previously pointed at JPEGs, which Android cannot build a proper adaptive icon from. Replaced with PNGs, plus a dedicated maskable variant:

| File | Size | Purpose |
|---|---|---|
| `logo/icon-192.png` | 192×192 | `any` |
| `logo/icon-512.png` | 512×512 | `any` |
| `logo/icon-512-maskable.png` | 512×512 | `maskable` |

The maskable icon centres the mark in a **60%** safe zone on a white plate, not the usual 80% — the source artwork already carries its own internal margin, and at 80% the mark touched the launcher's crop edge.

### Assets are served same-origin, and libraries are vendored

- Logo and icon references in `index.html` and `manifest.json` are now **relative** (`./logo/…`). Same origin as the app, served by GitHub Pages' CDN, and cacheable by the service worker. `raw.githubusercontent.com` is not a CDN, is rate-limited, discourages hotlinking, and is blocked on some corporate networks.
- `COMPANY_LOGO_URL` in `Code.gs` stays absolute — it goes into HTML email, which has no origin to be relative to — but points at the Pages URL rather than raw.
- **Chart.js and SheetJS are no longer loaded from public CDNs.** Both live in `vendor/` (see [Vendored libraries](#vendored-libraries)) and are cached in the app shell, so charts and Excel import work offline.

### Mail sending

All outbound mail (password OTPs and the daily report) now funnels through a single `sendMail_()` helper, sending as `noreply@calgas.in` — see [Mail configuration](#mail-configuration) for the alias requirement, which is not optional.

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
- **`sw.js`** caches the app shell (HTML/CSS/JS/icons/vendored libs) for fast repeat loads and offline resilience, but every request to the backend is always network-only — stock data is never served stale.

---

## Project structure

| File | Purpose |
|---|---|
| `Code.gs` | The entire backend: auth, sessions, RM/FG CRUD, transactions, ledger, bulk import, daily mail, drift check, admin tools. |
| `index.html` | The entire frontend: login/auth screens, dashboard, RM/FG stock pages, item history, Team & Access, all modals. |
| `manifest.json` | PWA manifest (name, icons, theme colors, scope). |
| `sw.js` | Service worker — app-shell caching, network-only for API calls. |
| `logo/` | Brand assets: PWA icons (PNG, incl. maskable) and the wordmark. Single source of truth for the whole app family. |
| `vendor/` | Self-hosted third-party libraries — **not committed by default**, see below. |
| `Stock_Management.xlsx` | The Google Sheet itself — the data store (see [Spreadsheet structure](#spreadsheet-structure)). |

---

## Vendored libraries

`vendor/` is intentionally empty in a fresh clone. Download both files once and commit them. Until you do, charts won't render and bulk import won't parse.

### `vendor/xlsx.full.min.js` — SheetJS **0.19.3**

```
https://cdn.sheetjs.com/xlsx-0.19.3/package/dist/xlsx.full.min.js
```

**Do not use the npm or cdnjs copy.** The newest version ever published to npm is 0.18.5, which is vulnerable to **CVE-2023-30533** (prototype pollution when *reading* a crafted file). This app calls `XLSX.read()` on operator-supplied files during bulk import, so that path is directly exposed — this is not a hygiene upgrade. The fix ships only from SheetJS's own CDN.

Save it with the browser's *Save page as*, not copy-paste: the file contains binary codepage tables that a rendered browser view will mangle.

### `vendor/chart.umd.min.js` — Chart.js **4.4.1**

```
https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js
```

Expected size 205,125 bytes. Previously loaded from an **unpinned** `cdn.jsdelivr.net/npm/chart.js` URL, which resolves to whatever is latest — a major release could have broken every chart without a single line of this repo changing.

> The npm tarball's `dist/chart.umd.js` is already minified and byte-identical to what jsDelivr serves as `chart.umd.min.js`; either source is fine.

---

## Spreadsheet structure

| Sheet | Purpose |
|---|---|
| `RM_Master` / `FG_Master` | One row per item: code, name, category/family, unit, reorder level, active flag, ID No. |
| `RM_Stock_Ledger` / `FG_Stock_Ledger` | Every transaction ever posted, append-only. Includes `ClientTxnId` for duplicate suppression. |
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
| Receive drift-check alerts | ✅ | ❌ | ❌ |

Every permission is enforced **server-side** (`requireRole_` / department-scope checks in `Code.gs`) — the frontend hiding a button is a convenience, not the actual security boundary.

---

## Mail configuration

Outbound mail is centralised in `sendMail_()` and sends as `MAIL_FROM` (`noreply@calgas.in`).

Apps Script **cannot send as an arbitrary address**. `MailApp.sendEmail` has no `from` parameter at all; `GmailApp.sendEmail` accepts one, but only when that address is a verified *Send mail as* alias on the Google account owning the script. So:

1. On the account that owns the Apps Script project, go to **Gmail → Settings → Accounts and Import → Send mail as** and add `noreply@calgas.in`, completing verification.
2. Re-authorize the script once. `GmailApp` needs a scope `MailApp` did not.

`sendMail_` checks `GmailApp.getAliases()` at runtime. If the alias is missing, mail **still sends** under the owner's own address rather than failing — a Workspace misconfiguration should not silently swallow password reset codes. `replyTo` is set to `MAIL_FROM` either way.

---

## Setup

### 1. Spreadsheet + Apps Script

1. Open `Stock_Management.xlsx` in Google Sheets (or import it as one).
2. **Extensions → Apps Script**, paste in `Code.gs`.
3. Run `ADMIN_generateAppSecret()` once from the Apps Script editor. This sets `Settings!APP_SECRET`, used for password hashing and session signing — without it, login will fail with an explicit error telling you to run this.
4. Add and verify the `noreply@calgas.in` alias — see [Mail configuration](#mail-configuration).
5. **Deploy → New deployment → Web app.** Execute as *Me*, accessible to *Anyone* (the app handles its own auth on top of this). Copy the deployment URL.
6. Run `ADMIN_installNightlyCleanupTrigger()` once — sweeps expired sessions and OTP codes nightly, and runs the balance drift check.
7. If you want the daily stock report emails: run `ADMIN_installDailyStockMailTrigger()` once. This needs Google Drive access (to generate `.xlsx` attachments) and Gmail send access, so it'll prompt a fresh authorization the first time — expected, not an error.
8. Create your first Admin user directly in the `Users` sheet (Username, Name, Role=`Admin`, Active=`TRUE`, a real Email), then run `ADMIN_setUserPassword('their-username', 'a-temporary-password')` to set their initial password.

> Do **not** switch the deployment to *Execute as: user accessing*. It would hand you `Session.getActiveUser()` for free, but it breaks the CORS-free `text/plain` POST from GitHub Pages.

### 2. Frontend

1. Download the two [vendored libraries](#vendored-libraries) into `vendor/` and commit them.
2. In `index.html`, set `GOOGLE_API_URL` to the Web App URL from step 5 above.
3. Host `index.html`, `manifest.json`, `sw.js`, `logo/` and `vendor/` together as static files — GitHub Pages off this repo's `main` branch serves them at `https://calgas.github.io/Stock-Management/`. They don't need to be anywhere near the Apps Script project.
4. Open the hosted URL, log in as the Admin user you just created, and change the temporary password via **Forgot password** (this also verifies the OTP email path is working).

> **Commit order matters.** Push `logo/` and `vendor/` *before* `index.html`, or the first visitors get a shell that 404s on its own icons and libraries — and the service worker may cache that broken state.

### 3. Ongoing accounts

New users are created from **Team & Access** in the app — that flow emails them a password-setup code automatically; there's no need to touch `ADMIN_setUserPassword` again except as a break-glass fallback if someone's email is unreachable.

### 4. Deploying an update

Bump **all three** version markers together, or installed clients will keep serving the old shell:

| Marker | File |
|---|---|
| `<meta name="app-version">` | `index.html` |
| `APP_VERSION` | `index.html` |
| `CACHE_VERSION` | `sw.js` |

Currently `3.6.2` / `3.6.2` / `calgas-shell-v16`.

---

## Tunable constants (`Code.gs`)

| Constant | Value | Notes |
|---|---|---|
| `SESSION_DURATION_HOURS` | 12 | Covers a full shift plus overrun. |
| `SESSION_CACHE_TTL_SECONDS` | 21600 | 6h — the `CacheService` maximum. Capped further by the session's own expiry. |
| `OTP_VALIDITY_MINUTES` | 10 | |
| `OTP_MAX_ATTEMPTS` | 5 | |
| `LOGIN_MAX_ATTEMPTS` | 5 | Then locked for `LOGIN_LOCKOUT_MINUTES`. |
| `LOGIN_LOCKOUT_MINUTES` | 15 | |
| `IDEMPOTENCY_SCAN_ROWS` | 300 | Ledger tail scanned for a duplicate `ClientTxnId`. |
| `DRIFT_TOLERANCE` | 0.001 | Absorbs float noise on fractional quantities. |
| `MAIL_FROM` | `noreply@calgas.in` | Requires a verified alias. |
| `ALLOW_NEGATIVE_STOCK` | `false` | Escape hatch only. `true` restores the old unguarded behaviour. |

---

## Notable design decisions

- **Passwords are salted** (`Salt` column per user), with automatic migration: any account created before salting was added keeps working, and gets silently upgraded to a salted hash the next time it logs in successfully — no bulk migration step was needed.
- **Retries are gated by idempotency, not by hope.** A request is only retried if it is a read, or a write carrying a `clientTxnId` the backend de-duplicates. That is the difference between a resilient client and one that corrupts data on a flaky connection.
- **The duplicate check lives inside the transaction lock**, not before it — otherwise two concurrent duplicates could both pass the check and both append.
- **Drift is reported, never auto-corrected.** Detection and correction are separate responsibilities; conflating them on an unattended trigger destroys the evidence you need to find the cause.
- **The session cache never becomes authoritative.** Expiry is always checked against the stored `ExpiresAt`, so caching can only save a read — it can't extend a session.
- **History lookups scan backward from the newest ledger rows in chunks**, stopping once enough matches are found, rather than reading the entire ledger on every click — this keeps item history fast regardless of how large the overall ledger grows.
- **Excel output (exports and daily mail) is generated entirely server-side** via native Google Sheets formatting, then exported to real `.xlsx` — not via a client-side JS library. The free tier of the usual browser-side Excel library can't write cell colors/fills/borders at all; doing it server-side means real styling (branded headers, status color-coding, banded rows) with no such limitation, and guarantees exports and mail look identical since they share the same styling code.
- **Bulk import creates each item's Current Stock row directly** at import time, seeded from an Opening Stock column — items are visible in the Stock tables immediately rather than only appearing after their first transaction.
- **The PWA app shell is cached, but stock data never is** — `sw.js` serves `index.html`/static assets cache-first for fast repeat loads, while every request to the Apps Script backend is always network-only.

---

## Known limitations / not yet done

- **Google Fonts is still fetched from the network**, so "works offline" is not yet strictly true — the app loads and functions, but falls back to system fonts with no connection. Self-hosting the font files would close this.
- **`index.html` is a single 3,400-line file.** Splitting out `app.css` and `app.js` needs no build step and would make every future edit cheaper. Deferred deliberately: it touches every line, and doing it alongside the 3.3.0 correctness work would have buried those changes in the diff.
- **Google sign-in is not implemented.** It's worth doing for office staff, but it is an *addition*, not a simplification — shop-floor devices are shared and need the username/password path to stay, so password hashing, OTP, `PasswordResets` and throttling all remain either way. Blocked on creating a Google Cloud OAuth client ID. If added, the ID token must be verified server-side against `https://oauth2.googleapis.com/tokeninfo`, checking `aud`, `exp` and `email_verified` — and `hd` too, if `calgas.in` becomes a Workspace domain.
- **No ledger archival.** Every transaction stays in one sheet forever. The backward-chunked scan keeps this fast for a long time, but past roughly 50k rows, rolling older rows into `RM_Stock_Ledger_Archive` will beat a clever scan.
- **No incremental recalculation** — `Recalculate` does a full rebuild of Current Stock from the entire ledger. This is intentional (a rare, manual, correctness-first admin tool), and is fine at normal usage; a very large ledger (20,000+ rows) will get a heads-up that it may take a while, since Apps Script executions have a hard 6-minute ceiling.
- **No explicit transaction reversal.** Corrections go through free-form Adjustments; a `ReversalOf` column and a "reverse this" action would give a much cleaner audit trail.
- **Gmail send quota** applies to OTPs, drift alerts and daily report emails — ~100/day on a consumer account, 1,500/day on Workspace. Irrelevant at a handful of subscribers, worth knowing if that list grows a lot.
- **Bulk import is capped at 500 rows per file**; split larger imports into batches. Exports are capped at 5,000 rows.
- **Sheet version history does not cover a deleted tab.** A weekly timestamped copy of the spreadsheet to Drive would.

---

## If this ever outgrows Sheets

Sheets is the right call at current scale — free, zero-ops, and "just open the sheet and look" is worth a great deal on a factory floor. The known limits above are localised, not architectural.

If it does outgrow it, **Supabase** (free Postgres) is a better target than Firebase for one reason specific to this design: Current Stock becomes a SQL view over the ledger, which deletes `applyStockDelta_`, `recalcCurrentStock_`, and drift as a concept entirely. It would also let ELE Tracker and Manufacturing Tracker join against one item master instead of calling across to this app. The cost is losing the spreadsheet-as-UI, which is a real loss.

Revisit when the ledger passes ~50k rows, or when write-lock contention becomes noticeable.
