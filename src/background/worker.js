// Background service worker — owns the entire sync lifecycle.
// The popup sends START_SYNC / CANCEL_SYNC and listens for progress events.
// State is also persisted to chrome.storage.local so the popup can hydrate
// after being closed and reopened mid-sync.

const JIRA_TZ      = 'America/New_York'
const JIRA_MAX     = 50
const STATE_KEY    = 'gas_sync_state'
const NOTIF_ID     = 'gas-trigger-sync'

// ── Shared mutable state ───────────────────────────────────────────────────
let syncAbortController = null
let cancelRequested     = false

// ── Storage helpers ────────────────────────────────────────────────────────
function saveState(patch) {
  return chrome.storage.local.set({ [STATE_KEY]: patch })
}

function loadState() {
  return chrome.storage.local.get([STATE_KEY]).then((r) => r[STATE_KEY] ?? null)
}

// ── Badge helpers ──────────────────────────────────────────────────────────
function badgeRunning(step, total) {
  const text = total ? `${step}/${total}` : '⟳'
  chrome.action.setBadgeText({ text })
  chrome.action.setBadgeBackgroundColor({ color: '#f59e0b' })
}

function badgeSuccess() {
  chrome.action.setBadgeText({ text: '✓' })
  chrome.action.setBadgeBackgroundColor({ color: '#0f9e6e' })
  setTimeout(() => chrome.action.setBadgeText({ text: '' }), 3000)
}

function badgeError() {
  chrome.action.setBadgeText({ text: '✗' })
  chrome.action.setBadgeBackgroundColor({ color: '#d63b3b' })
  setTimeout(() => chrome.action.setBadgeText({ text: '' }), 5000)
}

function badgeClear() {
  chrome.action.setBadgeText({ text: '' })
}

// ── Notification helpers ───────────────────────────────────────────────────
function notify(title, message) {
  chrome.notifications.create(NOTIF_ID, {
    type:    'basic',
    iconUrl: 'icons/icon48.png',
    title,
    message,
  })
}

// ── Popup messaging (best-effort — popup may be closed) ───────────────────
function tellPopup(type, payload) {
  chrome.runtime.sendMessage({ type, payload }).catch(() => {})
}

async function broadcastProgress(progress, status, extra = {}) {
  const state = { running: true, progress, status, result: null, ...extra }
  await saveState(state)
  tellPopup('SYNC_PROGRESS', { progress, status, ...extra })
  // Update badge step counter when processing dates
  if (extra.dateStep != null && extra.dateTotal != null) {
    badgeRunning(extra.dateStep, extra.dateTotal)
  } else {
    badgeRunning()
  }
}

// ── Date helpers (no moment in the worker — pure JS) ──────────────────────
// Converts M/D/YYYY or MM/DD/YYYY → YYYY-MM-DD in EST (America/New_York).
// Chrome extension workers support Intl, so we use that for timezone offset.
function toJiraDate(rawDate) {
  // Parse M/D/YYYY
  const m = rawDate.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/)
  if (!m) return null
  const [, mo, dy, yr] = m
  // Build a date at noon EST so UTC offset doesn't bleed into the next day.
  // EST = UTC-5, EDT = UTC-4. We use toLocaleDateString with timeZone to
  // get the display date, then reformat it as YYYY-MM-DD.
  const isoNoon = `${yr}-${mo.padStart(2, '0')}-${dy.padStart(2, '0')}T12:00:00`
  const d = new Date(isoNoon) // treated as local → good enough for noon
  if (isNaN(d)) return null
  // Format in New York timezone
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: JIRA_TZ,
    year:  'numeric',
    month: '2-digit',
    day:   '2-digit',
  }).formatToParts(d)
  const p = Object.fromEntries(parts.filter((x) => x.type !== 'literal').map((x) => [x.type, x.value]))
  return `${p.year}-${p.month}-${p.day}`
}

// ── HTTP helpers ───────────────────────────────────────────────────────────
function buildGetUrl(url, action, secretKey) {
  const sep = url.includes('?') ? '&' : '?'
  return `${url}${sep}action=${action}${secretKey ? `&key=${encodeURIComponent(secretKey)}` : ''}`
}

async function parseResponse(res) {
  const text = await res.text()
  try { return JSON.parse(text) }
  catch { return { success: false, error: text || 'Invalid response', code: res.status } }
}

async function getJiraSessionCookie(jiraBaseUrl) {
  try {
    const url    = jiraBaseUrl.startsWith('http') ? jiraBaseUrl : 'https://' + jiraBaseUrl
    const cookie = await chrome.cookies.get({ url, name: 'JSESSIONID' })
    return cookie ? cookie.value : null
  } catch {
    return null
  }
}

// ── Main sync function ─────────────────────────────────────────────────────
async function startSync({ url, secretKey, jiraBaseUrl, jiraJqlQuery }) {
  cancelRequested     = false
  syncAbortController = new AbortController()
  const { signal }    = syncAbortController

  badgeRunning()

  try {
    // Step 1 — get dates from column A
    await broadcastProgress(0, 'Reading dates from sheet...')

    const datesRes  = await fetch(buildGetUrl(url, 'getDates', secretKey), { signal })
    const datesData = await parseResponse(datesRes)

    if (datesData?.error === 'Unauthorized' || datesData?.code === 401) {
      throw Object.assign(new Error('Secret key mismatch. Check ⚙ Preferences.'), { code: 401 })
    }
    if (!datesData.success || !Array.isArray(datesData.dates) || !datesData.dates.length) {
      throw Object.assign(new Error('No valid dates found in column A of the sheet.'), { code: 404 })
    }

    const rawDates = datesData.dates
    await broadcastProgress(10, 'Dates loaded — reading JSESSIONID...', { dates: rawDates })

    // Step 2 — cookie + per-date JIRA fetch
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

      const statusText = `[${di + 1}/${rawDates.length}] Fetching JIRA for ${jiraDate}...`
      await broadcastProgress(
        10 + Math.round((di / rawDates.length) * 50),
        statusText,
        { dates: rawDates, cookieFound: !!jsessionId, dateStep: di + 1, dateTotal: rawDates.length }
      )

      const jqlTemplate = jiraJqlQuery?.trim() || 'due="{date}" ORDER BY created ASC'
      const jql         = encodeURIComponent(jqlTemplate.replace(/\{date\}/g, jiraDate))

      let allIssues = [], startAt = 0, total = null
      while (true) {
        if (cancelRequested) throw Object.assign(new DOMException('Cancelled', 'AbortError'), { cancelled: true })

        const jiraUrl = `${jiraBaseUrl}/rest/api/3/search?jql=${jql}&fields=summary,status,duedate&maxResults=${JIRA_MAX}&startAt=${startAt}`
        const jiraRes = await fetch(jiraUrl, { headers: jiraHeaders, credentials: 'include', signal })

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

    // Step 3 — write to GAS
    await broadcastProgress(65, 'Writing to Google Sheet...')

    const gasRes  = await fetch(url, {
      method:  'POST',
      headers: { 'Content-Type': 'text/plain' },
      body:    JSON.stringify({ action: 'syncJira', issuesByDate, key: secretKey || undefined }),
      signal,
    })
    const gasData = await parseResponse(gasRes)

    if (!gasData.success) {
      throw Object.assign(new Error(gasData.error || 'Sync failed'), { code: gasData.code })
    }

    // ── Success ────────────────────────────────────────────────────────────
    const st      = gasData.stats || {}
    const summary = `${st.inserted || 0} inserted, ${st.updated || 0} updated, ${st.moved || 0} moved, ${st.skipped || 0} skipped`
    const result  = { success: true, stats: { ...st, perDate }, message: `Sync complete — ${summary}` }

    await saveState({ running: false, progress: 100, status: '', result })
    badgeSuccess()
    notify('GAS Trigger — Sync Complete ✓', summary)
    tellPopup('SYNC_COMPLETE', { success: true, result })

    // Persist lastUsed
    const stored = await chrome.storage.sync.get(['gas_trigger_preferences'])
    const prefs  = stored.gas_trigger_preferences || {}
    chrome.storage.sync.set({ gas_trigger_preferences: { ...prefs, lastUsed: new Date().toISOString() } })

  } catch (err) {
    const isCancelled = err.cancelled || err.name === 'AbortError'

    if (isCancelled) {
      await saveState({ running: false, progress: 0, status: '', result: { cancelled: true } })
      badgeClear()
      tellPopup('SYNC_COMPLETE', { cancelled: true })
    } else {
      const result = { success: false, error: err.message || 'Unknown error', code: err.code }
      await saveState({ running: false, progress: 0, status: '', result })
      badgeError()
      notify('GAS Trigger — Sync Failed ✗', err.message || 'Unknown error')
      tellPopup('SYNC_COMPLETE', { success: false, error: err.message, code: err.code })
    }
  } finally {
    syncAbortController = null
  }
}

// ── Message router ─────────────────────────────────────────────────────────
chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  switch (message.type) {
    case 'START_SYNC': {
      if (cancelRequested || syncAbortController) {
        sendResponse({ ok: false, reason: 'already_running' })
        return false
      }
      startSync(message.payload).catch(console.error)
      sendResponse({ ok: true })
      return false
    }

    case 'CANCEL_SYNC': {
      cancelRequested = true
      if (syncAbortController) syncAbortController.abort()
      sendResponse({ ok: true })
      return false
    }

    case 'GET_SYNC_STATE': {
      loadState().then(sendResponse)
      return true // async
    }

    default:
      return false
  }
})
