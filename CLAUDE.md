# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm run build          # production build → dist/
npm run dev            # Vite dev server (UI preview only — extension APIs unavailable)
npm run generate-icons # regenerate public/icons/ from scripts/generate-icons.js
```

There are no tests or linters configured. After any change to `src/`, run `npm run build` and verify the `dist/` output before committing.

## One-time configuration

Before the system works end-to-end, two constants need real values:

| File | Constant | What to set |
|---|---|---|
| `src/background/worker.js` line 8 | `JIRA_DEPLOY_PATHS_FIELD` | Real Jira custom field ID, e.g. `customfield_10042` |
| `gas/Code.gs` line 7 | `SECRET_KEY` | Any secret string — must match the extension's Preferences → Secret Key |
| `gas/Code.gs` line 30 | `JIRA_BASE_URL` | Your Jira instance base URL, e.g. `https://yourorg.atlassian.net` |
| `gas/CategorizeSheet.gs` line 32 | `CAT_PATH_CATEGORIES` | Replace placeholder entries with real path prefixes |

## Architecture

This repo has two completely separate codebases that communicate over HTTP:

### 1. Chrome Extension (`src/`, built by Vite → `dist/`)

- **`src/background/worker.js`** — Service worker. Owns the entire sync lifecycle. Receives `START_SYNC` from the popup, fetches all dated blocks from GAS, queries Jira per date via JQL, posts `issuesByDate` to GAS, then runs a post-sync stale-ticket reconciliation pass. Persists running state to `chrome.storage.local` so the popup can hydrate after being closed mid-sync. Two keepalive layers (10s `setInterval` + 30s `chrome.alarms`) prevent Chrome from killing the worker during long GAS writes.
- **`src/popup/`** — React 18 UI. `App.jsx` is the shell; `SyncPanel.jsx` drives sync and progress display; `Settings.jsx` handles preferences. All Chrome API calls go through `src/hooks/useStorage.js` which falls back to `localStorage` when running outside the extension.

**Message protocol (popup ↔ worker):**
- `START_SYNC` → worker starts; responds `{ ok: true }` or `{ ok: false, reason }`
- `GET_SYNC_STATE` → returns persisted state for hydration
- `SYNC_PROGRESS` (worker→popup) — progress updates broadcast via `chrome.runtime.sendMessage`
- `SYNC_COMPLETE` (worker→popup) — final result

**Preferences** are stored in `chrome.storage.sync` under key `gas_trigger_preferences`: `url`, `secretKey`, `jiraBaseUrl`, `jiraJqlQuery`, `sheetId`, `sheetName`.

**Jira auth** uses the browser's existing `JSESSIONID` cookie (read via `chrome.cookies.get`) — the user must be logged in to Jira in Chrome. No OAuth.

**Date format** throughout the extension is `M/D/YYYY` (no leading zeros). `toJiraDate` converts to `YYYY-MM-DD` for JQL; `fromJiraDate` converts back.

### 2. Google Apps Script Backend (`gas/`)

Two independent `.gs` files — **never mix them**:

**`gas/Code.gs`** — Standalone Web App deployed on script.google.com. Entry points are `doGet` (read/getDates/takeSnapshot) and `doPost` (create/update/delete/syncJira/applyReconciliation/revertSnapshot/deleteSnapshot). All requests are authenticated by matching against `SECRET_KEY`.

Key operations:
- `syncJira(issuesByDate)` — the main write: builds `globalMap` (all tickets currently in sheet), classifies each incoming issue as update/move/insert, applies writes, runs `finaliseDate` on every affected block, returns `staleKeys` (tickets in sheet not returned by any JQL query this sync).
- `applyReconciliation(cancelled, rescheduled)` — post-sync pass called by the worker after it individually queries Jira for each stale ticket. Cancelled/On Hold → grey full-row style + clear CRP + clear Launches. Rescheduled (due date changed to a date not in sheet) → amber-orange full-row style + update due date cell + note + clear CRP + Launches.
- `finaliseDate(sheet, triggerRow)` — called after every block write; bulk-clears all row backgrounds then re-applies per-status cell colours. **Runs inside `syncJira` before `applyReconciliation` is called**, so reconciliation colours always win.

Sheet structure per dated block:
```
Row N:   Date row (yellow bg, col C = "Total Tickets", col D = count)
Row N+1: Header row (blue bg, cols A–G = COLUMNS)
Row N+2…: Ticket data rows (col A = HYPERLINK formula with ticket key as label)
Row last: Empty separator row
```

`isCancelled(status)` matches `"cancelled"`, `"canceled"`, `"on hold"` (case-insensitive) — all three get identical grey full-row treatment: `#bdbdbd`/`#424242` across A–G, CRP and Launches cleared.

**`gas/CategorizeSheet.gs`** — Bound script pasted directly into each target Google Sheet (not the standalone Web App). Adds a "GAS Trigger" menu → installs an `onEdit` trigger that detects a checkbox in col H of any block header row → opens a modal grouping Content Release Paths by `CAT_PATH_CATEGORIES` prefix matching.

### Deployment model

Every change to `gas/Code.gs` requires a **new deployment version** in Apps Script (Deploy → Manage deployments → pencil → New version). The web app URL stays the same but the old deployed code continues to run until a new version is created. This is the most common cause of features not working after code changes.

The extension is loaded unpacked from `dist/` in Chrome (chrome://extensions → Load unpacked). Rebuild with `npm run build` and click the refresh button in chrome://extensions to pick up changes.

### Stale ticket reconciliation flow

```
syncJira returns staleKeys
  └─ worker queries each stale ticket: GET /rest/api/3/issue/{key}?fields=status,duedate
       ├─ /^cancell?ed$/i or /^on hold$/i  → cancelled[]
       └─ newRawDate ≠ stale.blockDate
          AND newRawDate not in rawDatesSet → rescheduled[]
          (tickets with null duedate or new date already in sheet are silently skipped)
  └─ gasPost applyReconciliation({ cancelled, rescheduled })
       ├─ cancelled    → applyCancelledRow  (grey  + clear CRP/Launches)
       └─ rescheduled  → applyRescheduledRow (orange + update due date + note + clear CRP/Launches)
```

When a rescheduled ticket's target date block is later added to the sheet, the next `syncJira` detects it as a move (`globalMap.blockDate ≠ new date`) and `writeJiraRow`/`updateJiraFields` clear the note on the due date cell automatically.

### Known behavioural invariants

- **CRP/Launches are always empty for cancelled/on-hold rows.** `finaliseDate` calls `applyCancelledRow` (which clears those cells) for every cancelled/on-hold row in the entire sheet on every sync — not only when a row first transitions to that status. Any value manually entered into CRP or Launches on a cancelled/on-hold row will be erased on the next sync.
- **`writeJiraRow` and `updateJiraFields` always clear the note on the due date cell** (`.setNote('')`). Do not set notes on the due date cell from outside these functions — they will be wiped on the next sync that touches the row.
- **Stale ticket reconciliation is skipped entirely when `syncJira` times out** (`gasData.timedOut = true`). The sheet write succeeded but stale styling is deferred to the next full sync.
