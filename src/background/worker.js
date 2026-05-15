// Background service worker — owns the entire sync lifecycle.
// The popup sends START_SYNC / CANCEL_SYNC and listens for progress events.
// State is persisted to chrome.storage.local so the popup can hydrate
// after being closed and reopened mid-sync.
//
// Snapshot flow (sheet revert on cancel):
//   1. After all JIRA data is fetched but BEFORE the GAS write:
//      → GET action=takeSnapshot  (GAS copies Sheet1 → hidden _snapshot tab)
//   2a. Sync succeeds → POST action=deleteSnapshot
//   2b. Cancel triggered → POST action=revertSnapshot
//       (GAS restores _snapshot → Sheet1, then deletes _snapshot tab)

const JIRA_TZ          = 'America/New_York'
const JIRA_MAX         = 50
const STATE_KEY        = 'gas_sync_state'
const NOTIF_ID         = 'gas-trigger-sync'
const WRITE_TIMEOUT_MS  = 5 * 60 * 1000   // 5 min — GAS execution ceiling
const REVERT_TIMEOUT_MS = 2 * 60 * 1000   // 2 min for revert/snapshot ops

// ── Per-sync mutable state ─────────────────────────────────────────────────
let syncAbortController = null
let cancelRequested     = false
let snapshotTaken       = false
let syncPayload         = null  // kept for revert fetch (no abort signal needed)

// ── Storage helpers ────────────────────────────────────────────────────────
function saveState(patch) {
  return chrome.storage.local.set({ [STATE_KEY]: patch })
}
function loadState() {
  return chrome.storage.local.get([STATE_KEY]).then((r) => r[STATE_KEY] ?? null)
}

// ── Badge helpers ──────────────────────────────────────────────────────────
function badgeSet(text, color) {
  chrome.action.setBadgeText({ text })
  chrome.action.setBadgeBackgroundColor({ color })
}
function badgeRunning(step, total) {
  badgeSet(total ? `${step}/${total}` : '⟳', '#f59e0b')
}
function badgeReverting() {
  badgeSet('↩', '#7c5b00')
}
function badgeSuccess() {
  badgeSet('✓', '#0f9e6e')
  setTimeout(() => chrome.action.setBadgeText({ text: '' }), 3000)
}
function badgeError() {
  badgeSet('✗', '#d63b3b')
  setTimeout(() => chrome.action.setBadgeText({ text: '' }), 5000)
}
function badgeClear() {
  chrome.action.setBadgeText({ text: '' })
}

// ── Notification helper ────────────────────────────────────────────────────
function notify(title, message) {
  chrome.notifications.create(NOTIF_ID, { type: 'basic', iconUrl: 'icons/icon48.png', title, message })
}

// ── Popup messaging (best-effort — popup may be closed) ───────────────────
function tellPopup(type, payload) {
  chrome.runtime.sendMessage({ type, payload }).catch(() => {})
}

async function broadcastProgress(progress, status, extra = {}) {
  const state = { running: true, progress, status, result: null, ...extra }
  await saveState(state)
  tellPopup('SYNC_PROGRESS', { progress, status, ...extra })
  if (extra.dateStep != null && extra.dateTotal != null) {
    badgeRunning(extra.dateStep, extra.dateTotal)
  } else {
    badgeRunning()
  }
}

// ── EST date conversion (no moment in the worker) ─────────────────────────
function toJiraDate(rawDate) {
  const m = rawDate.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/)
  if (!m) return null
  const [, mo, dy, yr] = m
  const d = new Date(`${yr}-${mo.padStart(2, '0')}-${dy.padStart(2, '0')}T12:00:00`)
  if (isNaN(d)) return null
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: JIRA_TZ, year: 'numeric', month: '2-digit', day: '2-digit',
  }).formatToParts(d)
  const p = Object.fromEntries(parts.filter((x) => x.type !== 'literal').map((x) => [x.type, x.value]))
  return `${p.year}-${p.month}-${p.day}`
}

// ── HTTP helpers ───────────────────────────────────────────────────────────
function buildGetUrl(url, action, secretKey) {
  const sep = url.includes('?') ? '&' : '?'
  return `${url}${sep}action=${action}${secretKey ? `&key=${encodeURIComponent(secretKey)}` : ''}`
}

// Races a fetch against a timer. Resolves { timedOut: true } if the timer
// fires first. Propagates AbortError immediately so cancel still works.
function fetchWithTimeout(url, options, ms) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => resolve({ timedOut: true }), ms)
    fetch(url, options)
      .then(res => { clearTimeout(timer); resolve(res) })
      .catch(err => { clearTimeout(timer); reject(err) })
  })
}

async function parseResponse(res) {
  const text = await res.text()
  try { return JSON.parse(text) }
  catch { return { success: false, error: text || 'Invalid response', code: res.status } }
}

async function getJiraSessionCookie(jiraBaseUrl) {
  try {
    const full   = jiraBaseUrl.startsWith('http') ? jiraBaseUrl : 'https://' + jiraBaseUrl
    // chrome.cookies.get matches against cookie path scope; using origin (no
    // path) ensures JSESSIONID (always scoped to /) is found regardless of
    // any path the user appended to their Jira base URL.
    const origin = new URL(full).origin
    const cookie = await chrome.cookies.get({ url: origin, name: 'JSESSIONID' })
    return cookie ? cookie.value : null
  } catch { return null }
}

// ── GAS POST helper (always fresh fetch — no abort signal) ────────────────
// Pass timeoutMs > 0 to race the request against a hard deadline.
function gasPost(url, secretKey, action, extra = {}, timeoutMs = 0) {
  const p = fetch(url, {
    method:  'POST',
    headers: { 'Content-Type': 'text/plain' },
    body:    JSON.stringify({ action, key: secretKey || undefined, ...extra }),
  }).then(parseResponse)
  if (!timeoutMs) return p
  return Promise.race([
    p,
    new Promise((_, reject) =>
      setTimeout(
        () => reject(new Error(`${action} timed out after ${Math.round(timeoutMs / 60000)} min`)),
        timeoutMs,
      )
    ),
  ])
}

// ── Revert logic ───────────────────────────────────────────────────────────
async function doRevert(url, secretKey) {
  badgeReverting()
  tellPopup('SYNC_PROGRESS', { progress: 0, status: 'Reverting sheet to pre-sync state…', phase: 'reverting' })

  const revertData = await gasPost(url, secretKey, 'revertSnapshot', {}, REVERT_TIMEOUT_MS)

  if (revertData.success) {
    const msg = revertData.noSnapshot
      ? 'Sync cancelled — no changes were made to the sheet.'
      : 'Sync cancelled — sheet restored to pre-sync state.'
    await saveState({ running: false, progress: 0, status: '', result: { cancelled: true, reverted: true } })
    badgeClear()
    notify('GAS Trigger — Sync Cancelled', msg)
    tellPopup('SYNC_COMPLETE', { cancelled: true, reverted: true })
  } else {
    throw new Error(revertData.error || 'Revert call failed')
  }
}

// ── Main sync function ─────────────────────────────────────────────────────
async function startSync({ url, secretKey, jiraBaseUrl, jiraJqlQuery }) {
  cancelRequested     = false
  snapshotTaken       = false
  syncAbortController = new AbortController()
  syncPayload         = { url, secretKey }
  const { signal }    = syncAbortController

  badgeRunning()

  try {
    // ── Step 1: read dates from GAS ───────────────────────────────────────
    await broadcastProgress(0, 'Reading dates from sheet…')
    const datesRes  = await fetch(buildGetUrl(url, 'getDates', secretKey), { signal })
    const datesData = await parseResponse(datesRes)

    if (datesData?.error === 'Unauthorized' || datesData?.code === 401) {
      throw Object.assign(new Error('Secret key mismatch. Check ⚙ Preferences.'), { code: 401 })
    }
    if (!datesData.success || !Array.isArray(datesData.dates) || !datesData.dates.length) {
      throw Object.assign(new Error('No valid dates found in column A of the sheet.'), { code: 404 })
    }

    const rawDates = datesData.dates
    await broadcastProgress(10, 'Dates loaded — reading JSESSIONID…', { dates: rawDates })

    // ── Step 2: fetch JIRA per date ───────────────────────────────────────
    const jsessionId  = await getJiraSessionCookie(jiraBaseUrl)
    const jiraHeaders = { Accept: 'application/json' }
    if (jsessionId) jiraHeaders['Cookie'] = `JSESSIONID=${jsessionId}`

    const issuesByDate = {}
    const perDate      = []

    for (let di = 0; di < rawDates.length; di++) {
      if (cancelRequested) throw Object.assign(new DOMException('Cancelled', 'AbortError'), { cancelled: true })

      const rawDate  = rawDates[di]
      const jiraDate = toJiraDate(rawDate)
      if (!jiraDate) { perDate.push({ date: rawDate, skipped: true }); continue }

      await broadcastProgress(
        10 + Math.round((di / rawDates.length) * 50),
        `[${di + 1}/${rawDates.length}] Fetching JIRA for ${jiraDate}…`,
        { dates: rawDates, cookieFound: !!jsessionId, dateStep: di + 1, dateTotal: rawDates.length }
      )

      const jqlTemplate = jiraJqlQuery?.trim() || 'due="{date}" ORDER BY created ASC'
      const jql         = encodeURIComponent(jqlTemplate.replace(/\{date\}/g, jiraDate))

      let allIssues = [], startAt = 0, total = null
      while (true) {
        if (cancelRequested) throw Object.assign(new DOMException('Cancelled', 'AbortError'), { cancelled: true })

        const jiraUrl = `${jiraBaseUrl}/rest/api/3/search?jql=${jql}&fields=summary,status,duedate&maxResults=${JIRA_MAX}&startAt=${startAt}`
        // credentials:'include' is not needed — the Cookie header is set
        // manually above. Including it triggers strict CORS credentialed-
        // request mode, which Jira's CORS policy rejects for extension origins.
        const jiraRes = await fetch(jiraUrl, { headers: jiraHeaders, signal })

        if (!jiraRes.ok) {
          throw Object.assign(new Error(`JIRA ${jiraRes.status} for ${jiraDate}`), { code: jiraRes.status })
        }

        const jiraData = await jiraRes.json()
        if (total === null) total = jiraData.total
        allIssues = allIssues.concat(jiraData.issues || [])
        startAt  += (jiraData.issues || []).length
        if (startAt >= total || !(jiraData.issues || []).length) break
      }

      issuesByDate[rawDate] = allIssues.map((issue) => ({
        'Ticket Number': issue.key || '',
        'Title':         issue.fields?.summary || '',
        'Status':        issue.fields?.status?.name || 'unknown',
        'Due Date':      rawDate,
      }))
      perDate.push({ date: rawDate, jiraDate, issues: allIssues.length })
    }

    // ── Step 3: snapshot (before any sheet mutations) ─────────────────────
    await broadcastProgress(62, 'Taking sheet snapshot…')
    const snapData = await fetch(buildGetUrl(url, 'takeSnapshot', secretKey), { signal }).then(parseResponse)
    if (snapData.success) snapshotTaken = true

    // Final cancel gate — after snapshot, before write
    if (cancelRequested) throw Object.assign(new DOMException('Cancelled', 'AbortError'), { cancelled: true })

    // ── Step 4: write to GAS ──────────────────────────────────────────────
    await broadcastProgress(65, 'Writing to Google Sheet…')
    const gasResOrTimeout = await fetchWithTimeout(url, {
      method:  'POST',
      headers: { 'Content-Type': 'text/plain' },
      body:    JSON.stringify({ action: 'syncJira', issuesByDate, key: secretKey || undefined }),
      signal,
    }, WRITE_TIMEOUT_MS)

    let gasData
    if (gasResOrTimeout.timedOut) {
      // GAS finished writing (data visible in sheet) but the HTTP response
      // never arrived within 5 min. Treat as success so the badge clears.
      gasData = { success: true, stats: {}, timedOut: true }
    } else {
      gasData = await parseResponse(gasResOrTimeout)
      // Cancel check after write — handles cancel arriving mid-request
      if (cancelRequested) throw Object.assign(new DOMException('Cancelled', 'AbortError'), { cancelled: true })
      if (!gasData.success) {
        throw Object.assign(new Error(gasData.error || 'Sync failed'), { code: gasData.code })
      }
    }

    // ── Step 5: delete snapshot (sync succeeded) ──────────────────────────
    await broadcastProgress(90, 'Cleaning up…')
    // Fire-and-forget — cleanup is non-critical and GAS may still be slow
    // to respond after a large write. The next sync's takeSnapshot removes
    // any leftover _snapshot tab if this request doesn't reach GAS.
    gasPost(url, secretKey, 'deleteSnapshot', {}, 60_000).catch(() => {})

    // ── Success ───────────────────────────────────────────────────────────
    const st      = gasData.stats || {}
    const summary = gasData.timedOut
      ? 'sheet written — response timed out (data was saved)'
      : `${st.inserted || 0} inserted, ${st.updated || 0} updated, ${st.moved || 0} moved, ${st.skipped || 0} skipped`
    const result  = { success: true, stats: { ...st, perDate }, message: `Sync complete — ${summary}` }

    await saveState({ running: false, progress: 100, status: '', result })
    badgeSuccess()
    notify('GAS Trigger — Sync Complete ✓', summary)
    tellPopup('SYNC_COMPLETE', { success: true, result })

    // Persist lastUsed timestamp
    const stored = await chrome.storage.sync.get(['gas_trigger_preferences'])
    const prefs  = stored.gas_trigger_preferences || {}
    chrome.storage.sync.set({ gas_trigger_preferences: { ...prefs, lastUsed: new Date().toISOString() } })

  } catch (err) {
    const isCancelled = err.cancelled || err.name === 'AbortError'

    if (isCancelled) {
      if (snapshotTaken) {
        // Snapshot exists → sheet may have been partially or fully written → revert
        try {
          await doRevert(url, secretKey)
        } catch (revertErr) {
          await saveState({ running: false, progress: 0, status: '', result: { cancelled: true, revertFailed: true, error: revertErr.message } })
          badgeError()
          notify('GAS Trigger — Revert Failed ✗', 'The _snapshot tab in your sheet was preserved for manual recovery.')
          tellPopup('SYNC_COMPLETE', { cancelled: true, revertFailed: true, error: revertErr.message })
        }
      } else {
        // Cancel happened before the snapshot — no sheet changes were made
        await saveState({ running: false, progress: 0, status: '', result: { cancelled: true, reverted: false } })
        badgeClear()
        notify('GAS Trigger — Sync Cancelled', 'No changes were made to the sheet.')
        tellPopup('SYNC_COMPLETE', { cancelled: true, reverted: false })
      }
    } else {
      // Sync error (not a cancel)
      const result = { success: false, error: err.message || 'Unknown error', code: err.code }
      await saveState({ running: false, progress: 0, status: '', result })
      badgeError()
      notify('GAS Trigger — Sync Failed ✗', err.message || 'Unknown error')
      tellPopup('SYNC_COMPLETE', { success: false, error: err.message, code: err.code })
    }
  } finally {
    syncAbortController = null
    syncPayload         = null
  }
}

// ── Message router ─────────────────────────────────────────────────────────
chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  switch (message.type) {
    case 'START_SYNC': {
      if (syncAbortController) {
        sendResponse({ ok: false, reason: 'already_running' })
        return false
      }
      startSync(message.payload).catch(console.error)
      sendResponse({ ok: true })
      return false
    }

    case 'CANCEL_SYNC': {
      cancelRequested = true
      if (syncAbortController) {
        // Sync is active in this worker — abort the live fetch.
        // The catch block inside startSync will handle doRevert if needed.
        syncAbortController.abort()
        sendResponse({ ok: true })
        return false
      }
      // syncAbortController is null: either no sync is running, or the worker
      // was restarted mid-sync (MV3 can kill the worker between async ops),
      // resetting all in-memory state. Check persisted state and revert directly.
      ;(async () => {
        const state = await loadState()
        if (!state?.running) {
          sendResponse({ ok: true })
          return
        }
        // Worker was restarted mid-sync — retrieve prefs and attempt revert.
        // GAS's revertSnapshot handles the no-snapshot case (returns noSnapshot:true),
        // so it is safe to call regardless of whether a snapshot was taken.
        const stored = await chrome.storage.sync.get(['gas_trigger_preferences'])
        const prefs  = stored.gas_trigger_preferences || {}
        if (!prefs.url) {
          await saveState({ running: false, progress: 0, status: '', result: { cancelled: true, reverted: false } })
          badgeClear()
          tellPopup('SYNC_COMPLETE', { cancelled: true, reverted: false })
          sendResponse({ ok: true })
          return
        }
        try {
          await doRevert(prefs.url, prefs.secretKey || '')
        } catch (revertErr) {
          await saveState({ running: false, progress: 0, status: '', result: { cancelled: true, revertFailed: true, error: revertErr.message } })
          badgeError()
          notify('GAS Trigger — Revert Failed ✗', 'The _snapshot tab in your sheet was preserved for manual recovery.')
          tellPopup('SYNC_COMPLETE', { cancelled: true, revertFailed: true, error: revertErr.message })
        }
        sendResponse({ ok: true })
      })()
      return true  // keep the message channel open for the async sendResponse
    }

    case 'GET_SYNC_STATE': {
      loadState().then(sendResponse)
      return true
    }

    default:
      return false
  }
})
