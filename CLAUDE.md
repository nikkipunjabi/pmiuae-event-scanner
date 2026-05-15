# CLAUDE.md — context for the next Claude session

This file is the briefing you need before touching anything in this repo.
Read it end-to-end. It's deliberately written for an LLM, so I've kept the
prose dense and the file paths/column letters explicit.

---

## What this project is

A web-based QR scanner for PMI UAE Chapter event check-in. It scans a QR
code on each attendee's ticket, looks them up in a Google Sheet, shows
their details, lets the front-desk volunteer record the check-in plus an
optional at-door fee (default AED 200, with a Complimentary option for
guest speakers/VIPs), and writes everything back to the sheet. A second
tab on the sheet shows a live event summary driven entirely by formulas.

There is no build step, no framework, no bundler. Three files.

---

## File layout

```
pmiuae-event-scanner/
├── index.html      Single-page scanner UI. HTML + CSS + inline JS.
├── config.js       Window-level config; just the Apps Script /exec URL.
├── Code.gs         Apps Script backend (pasted into the Sheet's
│                   Extensions → Apps Script editor; not run from this repo).
├── README.md       General project overview.
└── CLAUDE.md       This file.
```

The repo is hosted on Netlify (auto-deploy on push to GitHub
`nikkipunjabi/pmiuae-event-scanner`, branch `main`). HTTPS is required
for the camera to start.

---

## Architecture at a glance

```
┌──────────────────┐    HTTPS    ┌──────────────────────┐    SpreadsheetApp
│  index.html      │ ──────────► │  Apps Script Web App │ ────────────► ┌───────────────┐
│  on Netlify      │             │  /exec endpoint      │               │ Google Sheet  │
│  + html5-qrcode  │ ◄────────── │  doGet / doPost      │ ◄──────────── │  Sheet1       │
└──────────────────┘    JSON     └──────────────────────┘               │  ActiveMember…│
                                                                         │  Summary      │
                                                                         └───────────────┘
```

Frontend talks to the Apps Script `/exec` URL via `fetch()`. The Apps
Script reads/writes a single Google Sheet by ID (hardcoded in `Code.gs`
as `SHEET_ID`). The Sheet has three tabs:

- **Sheet1 (`registrants_list`)** — every registrant row, replaced
  wholesale when Nikki imports a fresh list from their event platform.
- **ActiveMembersList** — the PMI member directory used to detect
  whether a registrant is a member.
- **Summary** — a formula-driven dashboard, fully rebuilt by
  `buildSummaryTab()` in Code.gs.

---

## Data model

### Sheet1 columns

| #  | Letter | Header                | Written by    | Notes |
|----|--------|-----------------------|---------------|-------|
| 1  | A      | ID                    | source system | Primary key. QR encodes this in `?id=` query string. |
| 2  | B      | Event ID              | source system | Used to filter by event in Summary. |
| 3  | C      | Event                 | source system | Event name. |
| 4  | D      | Event Date            | source system | |
| 5  | E      | User ID               | source system | |
| 6  | F      | First Name            | source system | |
| 7  | G      | Last Name             | source system | |
| 8  | H      | Email                 | source system | **Used to match against ActiveMembersList!Primaryemail** |
| 9  | I      | #Registrants          | source system | |
| 10 | J      | Gross Amount          | source system | |
| 11 | K      | Discount Amount       | source system | |
| 12 | L      | Late Fee              | source system | |
| 13 | M      | Net amount            | source system | |
| 14 | N      | Coupon                | source system | |
| 15 | O      | Registration Date     | source system | Text dates `dd-MM-yyyy HH:mm:ss`. |
| 16 | P      | Transaction ID        | source system | |
| 17 | Q      | Payment date          | source system | |
| 18 | R      | Payment Status        | source system | `Paid` / `Unpaid` / `Pending` / `Waiting List` |
| 19 | **S**  | **Checked In**        | scanner       | `Yes` / blank |
| 20 | **T**  | **Checked In Time**   | scanner       | `dd-MM-yyyy HH:mm:ss` in spreadsheet timezone |
| 21 | **U**  | **PMI ID**            | ARRAYFORMULA  | Auto-fills from ActiveMembersList via email match |
| 22 | **V**  | **Fee Amount (AED)**  | scanner       | `0` = complimentary, `>0` = paid amount, blank = no fee recorded |
| 23 | **W**  | **Fee Collection Time** | scanner     | Set whenever V is written, even for 0 |

Columns A–T are imported wholesale by the user when they refresh the
list before an event. Columns U/V/W are auxiliary — the user does NOT
paste over them. The PMI ID column's `ARRAYFORMULA` lives in **U2** and
auto-extends over open-ended range `H2:H`.

### ActiveMembersList

Two columns matter; their letters vary, so they're auto-detected by
header name in `setupPmiIdLookup()`:

- Email column header: **`Primaryemail`** (variants accepted: `primary email`, `email`, `email address`, `e-mail`)
- PMI ID column header: **`Personid`** (variants accepted: `person id`, `pmi id`, `pmiid`, `pmi_id`, `member id`, `membership id`)

If the user adds more headers later, extend the `MEMBER_EMAIL_HEADERS`
or `MEMBER_PMI_ID_HEADERS` arrays in `Code.gs`.

### Summary tab

Built by `buildSummaryTab()`. The picker in **B3** drives every formula
on the tab — change B3, and Event Name (E3), Event Date (E4), and every
section recalculates. The dropdown options come from a `UNIQUE` formula
in cell `L1` (column L is hidden).

Sections (rows are dynamic):

1. **HEADCOUNT** — Total Registrants, Checked In, Not Yet Checked In, Check-In Rate
2. **PAYMENT STATUS** — Paid, Unpaid, Pending, Other (each with % of total)
3. **REVENUE** — Gross/Discount/Late Fee/Net Revenue (from source columns J/K/L/M)
4. **FEE COLLECTION (AT-EVENT)** — Total Collected (sum of V), Paid (non-members), Complimentary (guests), PMI Members Checked-In
5. **ATTENDEE LIST** — A live `QUERY` with status conditional formatting; section header includes a live registrant count

---

## Key flows

### Scan → display

1. User taps **Start Scanner** → `html5-qrcode` opens the camera.
2. On scan success, `onScanSuccess(decodedText)` debounces (2 s) and stops the scanner.
3. `extractIdFromScan(text)` parses out the `id` query param. Real QRs look like
   `https://pmiuae.org/index.php?option=com_eventbooking&task=registrant.checkin&id=8289&Itemid=4729`,
   but we also support raw IDs and `#id=…` fragments.
4. `lookupRegistrant(id)` checks `localCache` first. Cache hit → render instantly.
5. Cache miss → GET `/exec?action=lookup&id=<id>` → render.
6. After the first successful lookup, `prefetchEvent(eventId)` fires in
   the background, GET `/exec?action=event&id=<eventId>`, which returns
   every registrant for that event. These all go into `localCache` so
   subsequent scans are instant.

### Check-in (with optional fee)

1. The button label varies by member status:
   - Member: `Check-In`
   - Non-member: `Check-In & Record Fee`
2. On click:
   - `readFeePicker_(detailsCard, isMember, feePreviouslyRecorded)` reads the fee picker.
     - Member or fee already recorded → returns `{ recordFee: false }`.
     - Non-member with fee picker visible → reads `#fee-custom-amount.value`. The custom input is the **single source of truth**; the AED 200 and Complimentary buttons just fill it. Empty input falls back to `200` (deliberately, to avoid accidental zero-fee check-ins).
   - POST `/exec` with body `{ action: 'checkin', id, force, recordFee, feeAmount }`.
3. Backend `checkInRegistrant(id, force, feeOpts)`:
   - Refuses if already checked in unless `force === true` (Re-check in anyway button sets this).
   - Writes `[Yes, timestamp]` to cells S–T in one batched `setValues`.
   - If `feeOpts.recordFee === true`: writes `[amount, timestamp]` to cells V–W.
   - Invalidates the row cache.
4. Frontend updates `localCache` with the returned registrant so a re-scan reflects the new state.

### "Already checked in" / "Already paid" guards

- `Re-check in anyway` button appears with `data-force="1"`. Clicking it sets `force = true` server-side.
- If `r.feeTime` is non-empty on lookup, the fee picker is HIDDEN and replaced with an info banner ("Fee already collected: AED 200 at …" or "Marked Complimentary at …"). The Check-In click then sends `recordFee: false`.

---

## Apps Script gotchas (read carefully — these have all bitten us)

### 1. Editing Code.gs does **not** update the live Web App

This is the most common confusion. Saving in the editor only affects:
- The editor's "Run" button
- Spreadsheet-bound functions (onOpen menu, buildSummaryTab, etc.)
- The `/dev` test URL

To update what `/exec` serves you MUST:
1. Save Code.gs
2. **Deploy → Manage deployments → ✏️ pencil on existing deployment**
3. **Version dropdown → "New version"** ← this is the critical step
4. Click **Deploy**

The URL stays the same. Without "New version", the deployment is unchanged.

**To verify what's deployed**, hit `<url>?action=version` — it returns a string that includes a build identifier (e.g. `v3-fee-support-2026-05-15`). If the build string doesn't match what's in the editor's Code.gs, the deployment is stale.

### 2. Deployment access must be "Anyone", not "Anyone with Google account"

The latter triggers a Google OAuth consent screen which CORS preflight can't traverse — fetch gets a "Load failed" error. "Anyone" makes the script truly public and lets it run with the deployer's permissions ("Execute as Me").

### 3. CORS: avoid the preflight

POST requests use `Content-Type: text/plain;charset=utf-8` and send a JSON-stringified body. Apps Script reads `e.postData.contents` and JSON.parses it. Using `application/json` would trigger a CORS preflight that Apps Script's redirect chain (`/exec` → `googleusercontent.com`) can't handle cleanly.

### 4. Cache

`readAllRows_()` reads all of Sheet1 in one `getRange().getValues()` and memoises the result in `CacheService.getScriptCache()` for 5 min. Date objects are coerced to ISO strings before caching because `CacheService` can't serialize Dates. Cache is invalidated by `invalidateRowCache_()` after every write (check-in, fee record, column setup).

The frontend has its own `localCache` Map; it's wiped on page reload.

### 5. Default test functions

`_testLookup()`, `_testCheckIn()`, `_diagnoseEmail(id)`, and `diagnoseEmail()` are runnable from the editor. The last two are the email-mismatch diagnostic — change `TARGET_ID` inside `diagnoseEmail()` to the registrant you want to inspect, run it, then check View → Logs.

---

## Frontend gotchas

1. **Camera requires HTTPS.** Netlify provides this automatically.
2. **No `localStorage`/`sessionStorage`.** The Cowork artifact rules disallow them; we use a plain JS `Map` for `localCache` instead. It clears on page reload, which is fine — the prefetch fills it again after the first scan of the event.
3. **html5-qrcode fires repeatedly while a code is in frame.** `lastScanAt` debounces by 2 s.
4. **Pre-warm ping.** On page load, the frontend fires a `?action=ping` against the backend. Apps Script Web Apps have a 1–3 s cold-start; this hides it.
5. **The fee custom input is the single source of truth.** Buttons (`AED 200`, `Complimentary`) just write their `data-fee-value` into the input. `readFeePicker_()` reads the input and never inspects which button is selected. This was a deliberate redesign — the picker used to track selected state independently of the input and that led to drift.
6. **Check-in payload logging.** Every check-in logs `[check-in payload]` and `[check-in response]` to `console.log`. Useful when debugging "fee not recorded" — if the payload contains `recordFee: true, feeAmount: 200` but the sheet didn't update, the backend deployment is stale.

---

## Common tasks (how to extend this)

### Change a metric in the Summary tab
Edit `buildSummaryTab()` in Code.gs. Use the `metric_(sheet, row, label, formula, opts)` helper for rows. Use `section_(sheet, row, label)` for section headers. After saving, run `buildSummaryTab` from the editor or via the **Event Summary → Build / Refresh Summary tab** menu — no redeployment needed (Summary code runs from inside the sheet, not via the Web App).

### Add a column to Sheet1
1. Update the `COL` constant in Code.gs (it goes 1-indexed by column letter).
2. Update `rowToObject_()` to include the new field.
3. Update the bulk read range in `readAllRows_()` — it currently reads up to `COL.FEE_TIME` (W = 23).
4. If the scanner should display it, update `renderRegistrant()` in index.html.
5. If the Summary should aggregate it, update `buildSummaryTab()`.
6. If the column needs auto-fill (like PMI ID), add a setup function and call it from `setupSheetColumns()`.

### Add a new endpoint
1. Add an `if (action === 'newaction') { ... }` branch in `doGet` or `doPost` in Code.gs.
2. Deploy a new version (see gotcha #1 above).
3. Call from the frontend with `callBackend(`${BACKEND_URL}?action=newaction&…`, …)`.

### Debug a failing PMI lookup
The most common cause is the registrant having a different email in the event registration system vs. the PMI member directory. Use `diagnoseEmail()`:
1. Open Code.gs in the Apps Script editor.
2. Edit `TARGET_ID = '8426'` inside the `diagnoseEmail()` function.
3. Click Run.
4. View → Logs. You'll see the email in both sheets and any fuzzy matches (same local-part, different domain).

### Refresh after pasting new Sheet1 data
The user's typical workflow before an event is to paste a fresh list of registrants into A2:T of Sheet1. The PMI ID formula in U2 auto-extends, so they don't need to do anything. But if they wipe the whole sheet (including U2), run **Event Summary → Setup / Refresh sheet columns (U, V, W)** to recreate the ARRAYFORMULA and the V/W headers.

---

## URLs & keys

- **Sheet ID** (in Code.gs): `1Y7aRxKemiTLT1Llk-uq7p2X8b3eUxQ7yqJeSK3HGwBg`
- **Deployment URL** (in config.js): the current `/exec` URL, changed every time the user creates a new deployment instead of editing the existing one.
- **GitHub repo**: `nikkipunjabi/pmiuae-event-scanner` (branch `main`)
- **Live app**: hosted on Netlify (URL Nikki sets up; auto-deploys on `git push`)

The Apps Script "Web App" URL is effectively a public API key — anyone with it can read/write the sheet via the defined actions. The user has chosen to live with this for now because the sheet is event-specific and not sensitive long-term. If we ever need to lock it down, add a shared-secret token check in `doGet`/`doPost`.

---

## What's been tested

- ✅ Camera scan of real PMI QR (`...?id=8289` format)
- ✅ Manual ID entry
- ✅ Pre-warm + bulk prefetch (instant subsequent scans)
- ✅ Check-in with fee (AED 200 default)
- ✅ Complimentary (V = 0, W = timestamp)
- ✅ Already-checked-in warning
- ✅ Already-fee-collected guard (picker hidden)
- ✅ Summary tab recalc on cell B3 change
- ✅ PMI ID auto-fill via ARRAYFORMULA
- ✅ `diagnoseEmail()` for mismatch debugging
- ✅ `?action=version` for deployment verification

## Known limitations

- No offline mode. If the network drops mid-event, nothing works.
- No name-based fallback if email doesn't match a PMI member — diagnosis is manual via `diagnoseEmail()`.
- No undo for an accidental check-in (would need a new endpoint).
- The "Already paid online" case for non-members (Payment Status = Paid in column R but pmiId is empty) still shows the fee picker. Staff should look at the Payment Status row in the details and tap Complimentary if appropriate.

---

## Bug history (so we don't re-introduce these)

1. **Circular dependency in Summary metrics.** Used to construct formulas inline with `row++` post-increment, which JS evaluates before the formula string, producing `B9 = B8 - B9`. Fixed by capturing each row into a named variable (`rTotal`, `rCheckedIn`, etc.) before building the formula.

2. **"Cannot merge cells over a filtered row."** Rebuild path now removes the filter, bandings, and existing merges before merging anything new.

3. **"Load failed" on Netlify.** Was caused by deployment access being set to "Anyone with Google account". Switched to "Anyone" and the consent screen disappeared.

4. **Fee not recorded.** Was caused by deploying without selecting "New version" — the `/exec` URL kept serving an older `doPost` that didn't read `recordFee`/`feeAmount`. `?action=version` was added so this can be confirmed in 5 seconds.

5. **Hostinger/cors confusion.** None — we use Netlify. Apps Script Web Apps work fine with static hosts as long as POST uses `text/plain`.

---

If anything in this file no longer matches reality, fix the file at the same time you fix the code — future sessions need it accurate.
