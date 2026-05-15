# PMI UAE Event Scanner

A QR code check-in app for PMI UAE Chapter events. Volunteers at the
front desk scan a ticket QR code, see the attendee's details with their
PMI member status front-and-centre, and check them in with one tap.
Non-members can be marked as Paid (default AED 200), Complimentary
(guest speakers / VIPs), or charged a custom amount — all in the same
flow. A live event Summary tab in the Google Sheet shows headcount,
payment breakdown, revenue, and at-door fee totals.

No build step. No framework. Three files.

---

## How it's wired

```
┌──────────────────┐    HTTPS    ┌──────────────────────┐    SpreadsheetApp
│  index.html      │ ──────────► │  Apps Script Web App │ ────────────► ┌───────────────┐
│  on Netlify      │             │  /exec endpoint      │               │ Google Sheet  │
│                  │ ◄────────── │  Code.gs             │ ◄──────────── │               │
└──────────────────┘    JSON     └──────────────────────┘               └───────────────┘
```

- **Frontend**: `index.html` + `config.js`. Static site, deployed via Netlify.
- **Backend**: `Code.gs` — a Google Apps Script Web App. The script lives inside the Google Sheet (Extensions → Apps Script).
- **Storage**: a single Google Sheet with three tabs (`Sheet1`, `ActiveMembersList`, `Summary`).

---

## Setup

### One-time

1. **Open the Google Sheet** → Extensions → Apps Script. Paste the contents of `Code.gs` into the editor and Save.
2. **Deploy the Web App**: Deploy → New deployment → Type: Web app. Set **Execute as: Me** and **Who has access: Anyone** (not "Anyone with a Google Account" — that one breaks CORS). Click Deploy. Copy the resulting `/exec` URL.
3. **Wire the frontend**: paste that URL into `config.js` as `backendUrl`.
4. **Prep the sheet**: refresh the spreadsheet, then run **Event Summary → Setup / Refresh sheet columns (U, V, W)** from the menu. This sets up the PMI ID auto-lookup and the fee tracking headers.
5. **Build the dashboard**: **Event Summary → Build / Refresh Summary tab**.
6. **Host the frontend**: push to GitHub and connect the repo to Netlify (or any static host that gives you HTTPS — the camera won't work without it).

### Day-of

1. Open the deployed scanner URL on a phone.
2. Tap **Start Scanner**, scan a ticket QR.
3. Card shows a green **PMI Member** banner (with their PMI ID) or an amber **Non-Member** banner.
4. **Member** → tap **Check-In**.
5. **Non-member** → the fee picker appears (default AED 200). Tap **Complimentary** for guests/speakers, edit the amount for one-offs, then tap **Check-In & Record Fee**.
6. Watch the **Summary** tab on the sheet for live totals — Total Collected ticks up in real time.

---

## Features

- QR code scanning via phone camera (works on iOS Safari + Android Chrome on HTTPS)
- Manual ID entry fallback for when the camera isn't available
- Automatic PMI member detection via lookup against an `ActiveMembersList` tab
- Big, unmissable member-status banner on every scan
- One-tap fee recording (default AED 200, complimentary, or custom)
- Live event summary with headcount, payment status, revenue, and at-door fee totals
- Auto-prefetch: after the first scan, every registrant for that event is cached so subsequent scans are instant
- Pre-warm ping on page load to hide the Apps Script cold-start
- "Test backend" link in the footer for quick diagnostics
- `?action=version` endpoint to verify which build of the backend is live

---

## Files

- `index.html` — Scanner UI (HTML/CSS/JS inline).
- `config.js` — Backend URL.
- `Code.gs` — Apps Script backend. Lives inside the Google Sheet's editor; this file is the source of truth.
- `CLAUDE.md` — Detailed context for AI-assisted edits (architecture, data model, gotchas, common tasks).

---

## Updating the backend

When you change `Code.gs`, you **must** deploy a new version for the
live `/exec` URL to pick up the change:

1. Save Code.gs in the Apps Script editor.
2. Deploy → Manage deployments → ✏️ pencil on the existing deployment.
3. **Version dropdown → "New version"** ← this step is the easy one to miss.
4. Click Deploy.

To verify it worked, open `<your-exec-url>?action=version` in a browser — it will return a build string like `v3-fee-support-2026-05-15`.

---

## Updating the frontend

`git push` to `main` → Netlify auto-deploys in ~30 seconds.

---

## Tech

- Vanilla HTML/CSS/JS, no build, no framework
- [html5-qrcode](https://github.com/mebjas/html5-qrcode) for camera scanning (loaded from CDN)
- Google Apps Script Web App for the backend
- Google Sheets as the database
- Hosted on Netlify (any static HTTPS host works)
