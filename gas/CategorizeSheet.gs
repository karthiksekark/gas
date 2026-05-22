// ============================================================
//  GAS Trigger — Categorize Sheet Script
//
//  Paste the ENTIRE contents of this file into the bound Apps
//  Script of your target Google Sheet:
//    Inside the sheet → Extensions → Apps Script → Code.gs
//    (replace or add alongside any existing code)
//
//  This script is separate from Code.gs which lives on
//  script.google.com as a standalone Web App.
//
//  Setup (one time per sheet):
//    1. Open the sheet.
//    2. A "GAS Trigger" menu appears in the toolbar automatically.
//    3. Click GAS Trigger → Set Up Categorize Button.
//    4. Approve the permission screen that appears.
//    5. Done — the checkbox in each block header is now active.
// ============================================================

// ── Constants — keep in sync with Code.gs ──────────────────
var CAT_CRPATHS_IDX    = 4   // col E — Content Release Paths
var CAT_JIRA_COL_COUNT = 7   // A–G (checkbox lives in col H = CAT_JIRA_COL_COUNT + 1)

// ── Predefined path categories ─────────────────────────────
// Each entry: { name, prefix, extras }
//   name   — group heading shown in the modal
//   prefix — a path is assigned to the FIRST entry whose prefix it starts with
//   extras — additional paths appended after the matched paths when the
//            category has at least one match (omit or leave [] if none needed)
// Paths that match no entry go to "Uncategorized".
// Replace the placeholder entries with your real ones.
var CAT_PATH_CATEGORIES = [
  { name: 'My App',      prefix: '/content/releases/my-app',      extras: ['/extra/path/for-my-app'] },
  { name: 'Another App', prefix: '/content/releases/another-app', extras: [] },
  { name: 'Launches',    prefix: '/launches/',                     extras: ['/extra/launch/path-a', '/extra/launch/path-b'] },
]

// ── Custom menu ────────────────────────────────────────────
// Simple trigger — runs automatically on every sheet open.
// No installation needed for this function itself.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('GAS Trigger')
    .addItem('Set Up Categorize Button', 'setupCategorizeTrigger')
    .addToUi()
}

// ── Trigger installation ───────────────────────────────────
// Called from GAS Trigger → Set Up Categorize Button.
// Installs the onEdit trigger that powers the Categorize checkbox.
// Re-running is safe — removes the previous trigger first.
function setupCategorizeTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onCategorizeEdit') ScriptApp.deleteTrigger(t)
  })
  ScriptApp.newTrigger('onCategorizeEdit').forSpreadsheet(ss).onEdit().create()
  SpreadsheetApp.getUi().alert(
    'Categorize trigger installed.\nThe checkbox in each block header is now active.'
  )
}

// ── Installable onEdit trigger ─────────────────────────────
// Detects the col-H checkbox being checked, resets it (button behaviour),
// finds the block's date by scanning upward, then shows the modal.
function onCategorizeEdit(e) {
  if (!e || !e.range) return
  var range = e.range
  if (range.getColumn() !== CAT_JIRA_COL_COUNT + 1) return
  if (String(e.value) !== 'TRUE') return

  // Reset immediately so the checkbox acts as a button, not a toggle
  range.setValue(false)

  var sheet     = range.getSheet()
  var headerRow = range.getRow()
  var tz        = e.source.getSpreadsheetTimeZone()

  // Scan upward from the header row to find the date row above it
  var blockDate       = null
  var blockTriggerRow = null
  for (var r = headerRow - 1; r >= 1; r--) {
    var d = catParseDate(sheet.getRange(r, 1).getValue(), tz)
    if (d) { blockDate = d; blockTriggerRow = r; break }
  }
  if (!blockDate) return

  // Notify user the script is running (visible before the modal opens)
  e.source.toast('Loading paths for ' + blockDate + '…', 'Running script', 10)

  catShowModal(sheet, blockTriggerRow, blockDate, tz)

  // showModalDialog blocks until the user closes it — this toast appears after
  e.source.toast('Script completed', 'Categorize', 4)
}

// ── Modal ──────────────────────────────────────────────────
// Reads all Content Release Paths in the block, matches each against
// CAT_PATH_CATEGORIES (prefix match, first match wins), then shows
// a read-only modal with one section per matched category + Uncategorized.
function catShowModal(sheet, triggerRow, blockDate, tz) {
  var lr = sheet.getLastRow()
  var s  = triggerRow + 2   // first data row (skips date row + header row)
  var e  = lr

  if (s <= lr) {
    var cA = sheet.getRange(s, 1, lr - s + 1, 1).getValues()
    for (var i = 0; i < cA.length; i++) {
      if (catParseDate(cA[i][0], tz) !== null) { e = s + i - 1; break }
    }
  }

  var allPaths = []
  if (s <= e) {
    var pathVals = sheet.getRange(s, CAT_CRPATHS_IDX + 1, e - s + 1, 1).getValues()
    for (var pi = 0; pi < pathVals.length; pi++) {
      var cell = String(pathVals[pi][0] || '').trim()
      if (!cell) continue
      cell.split('\n').forEach(function(p) {
        p = p.trim()
        if (p) allPaths.push(p)
      })
    }
  }

  // Match each path against CAT_PATH_CATEGORIES (first matching prefix wins).
  // Unmatched paths land in __uncategorized__.
  var groups = {}   // key → [path, …]
  allPaths.forEach(function(p) {
    var matched = false
    for (var ci = 0; ci < CAT_PATH_CATEGORIES.length; ci++) {
      if (p.indexOf(CAT_PATH_CATEGORIES[ci].prefix) === 0) {
        var key = CAT_PATH_CATEGORIES[ci].name
        if (!groups[key]) groups[key] = []
        groups[key].push(p)
        matched = true
        break
      }
    }
    if (!matched) {
      if (!groups['__uncategorized__']) groups['__uncategorized__'] = []
      groups['__uncategorized__'].push(p)
    }
  })

  // Append per-category extras after matched paths (only when the group has a match).
  CAT_PATH_CATEGORIES.forEach(function(cat) {
    if (groups[cat.name] && cat.extras && cat.extras.length) {
      cat.extras.forEach(function(ep) {
        var p = ep.trim()
        if (p) groups[cat.name].push(p)
      })
    }
  })

  // Preserve the order defined in CAT_PATH_CATEGORIES; skip empty categories.
  var namedOrder = CAT_PATH_CATEGORIES
    .map(function(c) { return c.name })
    .filter(function(name) { return !!groups[name] })

  var html = catBuildHtml(blockDate, groups, namedOrder)
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(540).setHeight(460),
    'Categorize — ' + blockDate
  )
}

// ── HTML builder ───────────────────────────────────────────
function catBuildHtml(blockDate, groups, namedOrder) {
  var h = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
  h += 'body{font-family:Arial,sans-serif;font-size:13px;margin:0;padding:16px 20px;color:#333;overflow-y:auto;}'
  h += 'h2{font-size:15px;color:#1a1a2e;margin:0 0 16px;padding-bottom:8px;border-bottom:2px solid #4f6ef7;}'
  h += '.group{margin-bottom:16px;}'
  h += '.gh{font-weight:bold;color:#4f6ef7;margin-bottom:5px;font-size:13px;text-transform:uppercase;letter-spacing:.5px;}'
  h += '.gh.unc{color:#bbb;}'
  h += '.path{font-family:monospace;font-size:11px;color:#555;padding:3px 4px 3px 10px;'
  h += 'border-left:3px solid #e0e7ff;margin:2px 0;word-break:break-all;line-height:1.5;}'
  h += '.empty{color:#aaa;font-style:italic;}'
  h += '</style></head><body>'
  h += '<h2>' + catEsc(blockDate) + '</h2>'

  var hasContent = namedOrder.length > 0 || groups['__uncategorized__']
  if (!hasContent) {
    h += '<p class="empty">No Content Release Paths found for this block.</p>'
  } else {
    namedOrder.forEach(function(name) {
      h += '<div class="group"><div class="gh">' + catEsc(name) + '</div>'
      groups[name].forEach(function(p) { h += '<div class="path">' + catEsc(p) + '</div>' })
      h += '</div>'
    })
    if (groups['__uncategorized__']) {
      h += '<div class="group"><div class="gh unc">Uncategorized</div>'
      groups['__uncategorized__'].forEach(function(p) { h += '<div class="path">' + catEsc(p) + '</div>' })
      h += '</div>'
    }
  }
  h += '</body></html>'
  return h
}

// ── Helpers ────────────────────────────────────────────────

// Date parser — accepts Date objects or M/D/YYYY / YYYY-MM-DD strings.
// Uses the spreadsheet's own timezone (passed in as tz) so Date objects
// from getValues() are interpreted correctly regardless of server locale.
function catParseDate(val, tz) {
  if (!val && val !== 0) return null
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null
    var s = Utilities.formatDate(val, tz, 'M/d/yyyy')
    var mp = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/)
    if (!mp || +mp[1]<1||+mp[1]>12||+mp[2]<1||+mp[2]>31||+mp[3]<2000) return null
    return +mp[1]+'/'+mp[2]+'/'+mp[3]
  }
  var s = String(val).trim()
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/)
  if (m && +m[1]>=1&&+m[1]<=12&&+m[2]>=1&&+m[2]<=31&&+m[3]>=2000)
    return +m[1]+'/'+m[2]+'/'+m[3]
  var y = s.match(/^(\d{4})-(\d{2})-(\d{2})$/)
  if (y && +y[2]>=1&&+y[2]<=12&&+y[3]>=1&&+y[3]<=31&&+y[1]>=2000)
    return +y[2]+'/'+y[3]+'/'+y[1]
  return null
}

function catEsc(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}
