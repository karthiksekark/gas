# GAS Trigger — JIRA → Google Sheet Sync

A Chrome extension that pulls JIRA issues by due date and writes them into a Google Sheet, powered by a Google Apps Script web app backend.

---

## How It Works

```
Chrome Extension  →  Google Apps Script Web App  →  Google Sheet
      ↑                        ↑
  JIRA REST API          your Code.gs
```

1. The extension reads due dates from **column A** of your sheet.
2. For each date it queries JIRA and collects matching issues.
3. It takes a snapshot of the sheet (for safe cancellation), then writes the data back via GAS.
4. On success the snapshot is deleted. On cancel the sheet is fully restored.

---

## Prerequisites

- Google account with access to Google Drive and Google Sheets
- Google Apps Script enabled (no Workspace plan required)
- Chrome browser
- JIRA instance (cloud or server) — you must be logged in to JIRA in Chrome during sync
- Node.js ≥ 18 and npm (to build the extension)

---

## Part 1 — Set Up the Google Apps Script

### Step 1 — Open or create your Google Sheet

Open the Google Sheet you want to sync into. Make sure column A contains dates in `M/D/YYYY` format (e.g. `5/1/2026`).

> **Sheet on someone else's Drive?** No problem — open their shared sheet directly in your browser and follow the same steps below. The script you deploy will run as **you** and access that sheet using your permissions.

---

### Step 2 — Create the Apps Script

1. Inside the sheet, go to **Extensions → Apps Script**.
2. Delete any existing code in `Code.gs`.
3. Paste the entire contents of [`gas/Code.gs`](gas/Code.gs) from this repo.
4. Find the line near the top:
   ```javascript
   const SECRET_KEY = 'your-secret-key-here'
   ```
   Replace `your-secret-key-here` with a strong secret of your choice (e.g. a random 32-character string). Keep this value — you will enter it in the extension later.

---

### Step 3 — Deploy as a Web App

1. Click **Deploy → New deployment**.
2. Click the gear icon ⚙ next to **Type** and select **Web app**.
3. Set:
   - **Description**: anything (e.g. `v1`)
   - **Execute as**: **Me**
   - **Who has access**: **Anyone**
4. Click **Deploy**.
5. Copy the **Web app URL** — it looks like:
   ```
   https://script.google.com/macros/s/AKfycb.../exec
   ```
   You will paste this into the extension.

> **Note:** Every time you edit `Code.gs` you must create a **new deployment** (or **manage deployments → edit** the existing one) for changes to take effect. Saving the script alone is not enough.

---

## Part 2 — Build and Install the Chrome Extension

### Step 4 — Build the extension

```bash
git clone <this-repo>
cd gas
npm install
npm run build
```

This produces a `dist/` folder containing the built extension.

---

### Step 5 — Load the extension in Chrome

1. Open Chrome and go to `chrome://extensions`.
2. Enable **Developer mode** (toggle in the top-right corner).
3. Click **Load unpacked**.
4. Select the `dist/` folder generated in the previous step.
5. The **GAS Trigger** extension appears in your toolbar. Pin it for easy access.

---

## Part 3 — Configure the Extension

### Step 6 — Open Preferences

Click the extension icon in the toolbar, then click the **⚙** button.

---

### GAS Connection

| Field | What to enter |
|---|---|
| **Web App URL** | The URL copied in Step 3 |
| **Secret Key** | The value you set for `SECRET_KEY` in `Code.gs` |
| **Sheet ID** *(optional)* | The spreadsheet ID from the sheet's URL — see below |
| **Tab Name** *(optional)* | The exact name of the worksheet tab to sync |

**When to fill in Sheet ID:**
Leave it blank if the Apps Script is bound to the sheet you want to sync (the most common setup). Fill it in when you want one script deployment to target a *different* sheet — for example, a sheet shared with you from someone else's Drive.

The Sheet ID is the long string in the sheet's URL:
```
https://docs.google.com/spreadsheets/d/  ← SHEET_ID_HERE →  /edit
```

**When to fill in Tab Name:**
Leave it blank to always sync the first visible tab. Fill it in to target a specific tab by its exact name (e.g. `Sprint 12`), which is reliable regardless of tab order.

---

### JIRA Integration

| Field | What to enter |
|---|---|
| **JIRA Base URL** | Your JIRA instance root, e.g. `https://yourorg.atlassian.net` |
| **JQL Query** *(optional)* | Custom JQL with `{date}` as a placeholder — replaced per date at sync time |

**Default JQL** (used when the field is left blank):
```
due="{date}" ORDER BY created ASC
```

**Example custom JQL:**
```
due="{date}" AND project = MYPROJ AND assignee = currentUser() ORDER BY priority DESC
```

---

### Step 7 — Log in to JIRA in Chrome

The extension reads your JIRA session cookie (`JSESSIONID`) automatically. Simply make sure you are **logged in to JIRA in the same Chrome profile** before running a sync. The config card in the popup shows `✓ found` next to JSESSIONID once a sync has started and the cookie is detected.

---

## Part 4 — Run a Sync

1. Make sure column A of your sheet contains dates in `M/D/YYYY` format.
2. Click the extension icon and click **⟳ Sync from JIRA**.
3. The popup shows live progress. The extension badge on the icon also shows status.
4. You can safely close the popup — the sync continues in the background.
5. A desktop notification appears when the sync completes or fails.

### Cancelling a sync

Click **✕ Cancel Sync** at any time. If the sheet write had already started, the extension automatically restores the sheet to its pre-sync state using the snapshot taken before the write.

---

## Sheet Structure

The script expects and maintains the following layout in your sheet:

| Column A | Column B | Column C | Column D | Column E |
|---|---|---|---|---|
| Date (date row) | — | Total Tickets | Count | — |
| Ticket Number | Title | Status | Due Date | Comments |
| `PROJ-123` | Issue title | In Progress | 5/1/2026 | your notes |

- **Date rows** are yellow with a ticket count summary.
- **Header rows** are blue.
- **Comments (column E) and any columns beyond E** are never overwritten by the sync — they are yours to manage.

---

## Re-deploying After Code Changes

Whenever you update `Code.gs`:

1. In the Apps Script editor, click **Deploy → Manage deployments**.
2. Click the pencil ✏ icon on your existing deployment.
3. Change **Version** to **New version**.
4. Click **Deploy**.

The Web App URL stays the same — no need to update the extension.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `Secret key mismatch` error | `SECRET_KEY` in Code.gs doesn't match the extension | Re-check both values, redeploy |
| `No valid dates found` | Column A is empty or dates are in the wrong format | Use `M/D/YYYY` format (e.g. `5/1/2026`) |
| Tab not found error | Tab Name in preferences doesn't match exactly | Check capitalisation and spaces |
| JSESSIONID `✕ not found` | Not logged in to JIRA in Chrome | Log in to JIRA and retry |
| JIRA 401 error | JIRA session expired | Log in to JIRA again |
| Badge stays orange after sync | GAS write is still in progress | Wait — the extension recovers automatically within 6 minutes |
| Sync worked but `_snapshot` tab remains | deleteSnapshot timed out | Delete the `_snapshot` tab manually; it will be cleaned up on the next sync |
| Script doesn't see the sheet's data | Sheet ID not set for a shared/external sheet | Add the Sheet ID in ⚙ Preferences |
