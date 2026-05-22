// Background service worker — owns the entire sync lifecycle.
// The popup sends START_SYNC and listens for progress events.
// State is persisted to chrome.storage.local so the popup can hydrate
// after being closed and reopened mid-sync.

const JIRA_MAX             = 50
// TODO: replace with the real Jira custom field ID (e.g. customfield_10042)
const JIRA_DEPLOY_PATHS_FIELD = 'customfield_DEPLOY_PATHS'
const STATE_KEY            = 'gas_sync_state'
const NOTIF_ID             = 'gas-trigger-sync'
const WRITE_TIMEOUT_MS     = 5 * 60 * 1000   // 5 min — GAS execution ceiling
const WRITE_ALARM_NAME     = 'gas_write_timeout'
const KEEPALIVE_ALARM_NAME = 'gas_sync_keepalive'

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
function badgeSuccess() {
  badgeSet('✓', '#0f9e6e')
  setTimeout(() => chrome.action.setBadgeText({ text: '' }), 3000)
}
function badgeError() {
  badgeSet('✗', '#d63b3b')
  setTimeout(() => chrome.action.setBadgeText({ text: '' }), 5000)
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

// ── Date conversion: M/D/YYYY → YYYY-MM-DD ────────────────────────────────
// Pure string reformat — no Date object or timezone involved.
// The sheet date IS the intended Jira due-date; converting through a local
// Date and re-formatting in America/New_York caused off-by-one errors for
// users in timezones ≥ UTC+9 (noon local < 04:00 UTC = still May 17 in NYC).
function toJiraDate(rawDate) {
  const m = rawDate.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/)
  if (!m) return null
  const [, mo, dy, yr] = m
  if (+mo < 1 || +mo > 12 || +dy < 1 || +dy > 31 || +yr < 2000) return null
  return `${yr}-${mo.padStart(2, '0')}-${dy.padStart(2, '0')}`
}

// ── HTTP helpers ───────────────────────────────────────────────────────────
function buildGetUrl(url, action, secretKey, sheetId, sheetName) {
  const sep = url.includes('?') ? '&' : '?'
  let result = `${url}${sep}action=${action}${secretKey ? `&key=${encodeURIComponent(secretKey)}` : ''}`
  if (sheetId)   result += `&spreadsheetId=${encodeURIComponent(sheetId)}`
  if (sheetName) result += `&sheetName=${encodeURIComponent(sheetName)}`
  return result
}

// Races a fetch against a timer. Resolves { timedOut: true } if the timer
// fires first. Rejects on network errors.
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

// ── GAS POST helper ────────────────────────────────────────────────────────
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

// ── Main sync function ─────────────────────────────────────────────────────
async function startSync({ url, secretKey, jiraBaseUrl, jiraJqlQuery, sheetId, sheetName }) {
  // Two-layer keepalive so Chrome delivers fetch responses to promise
  // callbacks even during long-running GAS operations:
  //
  // Layer 1 — setInterval (10s): continuously tickles the V8 event loop via
  //   a lightweight storage read, preventing Chrome from idling the microtask
  //   queue (the Heisenbug where fetch responses are silently dropped without
  //   DevTools attached).
  //
  // Layer 2 — chrome.alarms repeating (30s): survives worker kills. If Chrome
  //   terminates the worker between events, the alarm fires in a fresh instance
  //   which allows any pending network responses to be requeued and delivered.
  const keepAlive = setInterval(() => chrome.storage.local.get(STATE_KEY), 10_000)
  await chrome.alarms.create(KEEPALIVE_ALARM_NAME, { periodInMinutes: 0.5 })

  badgeRunning()

  try {
    // ── Step 1: read dates from GAS ───────────────────────────────────────
    await broadcastProgress(0, 'Reading dates from sheet…')
    const datesRes  = await fetch(buildGetUrl(url, 'getDates', secretKey, sheetId, sheetName))
    const datesData = await parseResponse(datesRes)

    if (datesData?.error === 'Unauthorized' || datesData?.code === 401) {
      throw Object.assign(new Error('Secret key mismatch. Check ⚙ Preferences.'), { code: 401 })
    }
    if (!datesData.success || !Array.isArray(datesData.dates) || !datesData.dates.length) {
      throw Object.assign(new Error('No valid dates found in column A of the sheet.'), { code: 404 })
    }

    const rawDates = datesData.dates
    await broadcastProgress(10, 'Dates loaded — checking Jira session…', { dates: rawDates })

    // ── Step 2: fetch JIRA per date ───────────────────────────────────────
    const jsessionId = await getJiraSessionCookie(jiraBaseUrl)

    // Pre-check: surface a clear message before touching Jira rather than
    // letting the first fetch fail silently or return a cryptic 401.
    if (!jsessionId) {
      throw Object.assign(
        new Error('No Jira session found. Please log in to Jira in Chrome and sync again.'),
        { code: 401 }
      )
    }

    const jiraHeaders = { Accept: 'application/json', Cookie: `JSESSIONID=${jsessionId}` }

    const issuesByDate = {}
    const perDate      = []

    for (let di = 0; di < rawDates.length; di++) {
      const rawDate  = rawDates[di]
      const jiraDate = toJiraDate(rawDate)
      if (!jiraDate) { perDate.push({ date: rawDate, skipped: true }); continue }

      await broadcastProgress(
        10 + Math.round((di / rawDates.length) * 55),
        `[${di + 1}/${rawDates.length}] Fetching JIRA for ${jiraDate}…`,
        { dates: rawDates, cookieFound: !!jsessionId, dateStep: di + 1, dateTotal: rawDates.length }
      )

      const jqlTemplate = jiraJqlQuery?.trim() || 'due="{date}" ORDER BY created ASC'
      const jql         = encodeURIComponent(jqlTemplate.replace(/\{date\}/g, jiraDate))

      let allIssues = [], startAt = 0, total = null
      while (true) {
        const jiraUrl = `${jiraBaseUrl}/rest/api/3/search?jql=${jql}&fields=summary,status,duedate,${JIRA_DEPLOY_PATHS_FIELD}&maxResults=${JIRA_MAX}&startAt=${startAt}`
        // credentials:'include' is not needed — the Cookie header is set
        // manually above. Including it triggers strict CORS credentialed-
        // request mode, which Jira's CORS policy rejects for extension origins.
        const jiraRes = await fetch(jiraUrl, { headers: jiraHeaders })

        if (!jiraRes.ok) {
          const msg =
            jiraRes.status === 401 ? 'Jira session expired. Please log in to Jira in Chrome and sync again.' :
            jiraRes.status === 403 ? 'Access denied by Jira. Your account may not have permission to view this project.' :
            `Jira returned ${jiraRes.status} for ${jiraDate}.`
          throw Object.assign(new Error(msg), { code: jiraRes.status })
        }

        const jiraData = await jiraRes.json()
        if (total === null) total = jiraData.total
        allIssues = allIssues.concat(jiraData.issues || [])
        startAt  += (jiraData.issues || []).length
        if (startAt >= total || !(jiraData.issues || []).length) break
      }

      issuesByDate[rawDate] = allIssues.map((issue) => {
        const rawPaths = issue.fields?.[JIRA_DEPLOY_PATHS_FIELD]
        const allPaths = Array.isArray(rawPaths)
          ? rawPaths.join('\n')
          : (rawPaths || '')
        return {
          'Ticket Number':        issue.key || '',
          'Title':                issue.fields?.summary || '',
          'Status':               issue.fields?.status?.name || 'unknown',
          'Due Date':             rawDate,
          'Content Release Paths': allPaths,
        }
      })
      perDate.push({ date: rawDate, jiraDate, issues: allIssues.length })
    }

    // ── Step 3: write to GAS ──────────────────────────────────────────────
    await broadcastProgress(65, 'Writing to Google Sheet…')

    // One-shot alarm at 6 min (GAS execution ceiling) as last-resort recovery:
    // if the worker is killed before fetchWithTimeout's setTimeout fires, the
    // alarm fires in a fresh instance which reads the stale running state and
    // fires success signals. Cleared in the finally block below.
    await chrome.alarms.create(WRITE_ALARM_NAME, { delayInMinutes: 6 })

    let gasData
    try {
      const gasResOrTimeout = await fetchWithTimeout(url, {
        method:  'POST',
        headers: { 'Content-Type': 'text/plain' },
        body:    JSON.stringify({
          action:        'syncJira',
          issuesByDate,
          key:           secretKey  || undefined,
          spreadsheetId: sheetId   || undefined,
          sheetName:     sheetName || undefined,
        }),
      }, WRITE_TIMEOUT_MS)

      if (gasResOrTimeout.timedOut) {
        // 5-min timer fired — data is written but HTTP response never arrived.
        gasData = { success: true, stats: {}, timedOut: true }
      } else {
        // Headers arrived; read body with a 30s guard against a stalled stream.
        const parsed = await Promise.race([
          parseResponse(gasResOrTimeout),
          new Promise((resolve) => setTimeout(() => resolve({ timedOut: true }), 30_000)),
        ])
        if (parsed.timedOut) {
          gasData = { success: true, stats: {}, timedOut: true }
        } else {
          if (!parsed.success) {
            throw Object.assign(new Error(parsed.error || 'Sync failed'), { code: parsed.code })
          }
          gasData = parsed
        }
      }
    } finally {
      // Always clear the write alarm when the write resolves normally.
      // If the worker was killed, this finally never runs and the alarm fires.
      chrome.alarms.clear(WRITE_ALARM_NAME)
    }

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
    const result = { success: false, error: err.message || 'Unknown error', code: err.code }
    await saveState({ running: false, progress: 0, status: '', result })
    badgeError()
    notify('GAS Trigger — Sync Failed ✗', err.message || 'Unknown error')
    tellPopup('SYNC_COMPLETE', { success: false, error: err.message, code: err.code })
  } finally {
    clearInterval(keepAlive)
    chrome.alarms.clear(KEEPALIVE_ALARM_NAME)
  }
}

// ── Alarm handlers ─────────────────────────────────────────────────────────
chrome.alarms.onAlarm.addListener(async (alarm) => {

  if (alarm.name === KEEPALIVE_ALARM_NAME) {
    // Fires every 30s during a sync. No action needed — waking the worker
    // is enough to allow Chrome to deliver pending fetch responses.
    return
  }

  if (alarm.name === WRITE_ALARM_NAME) {
    // The worker was killed during the GAS write before the 5-min setTimeout
    // could fire. GAS has had 6 min (its execution ceiling) to finish.
    // Recover the stale running state as a success.
    const state = await loadState()
    if (!state?.running) return  // write already completed via the normal path

    const stored = await chrome.storage.sync.get(['gas_trigger_preferences'])
    const prefs  = stored.gas_trigger_preferences || {}

    const result = { success: true, stats: {}, message: 'Sync complete — sheet written (connection was lost, data was saved)' }
    await saveState({ running: false, progress: 100, status: '', result })
    badgeSuccess()
    notify('GAS Trigger — Sync Complete ✓', 'sheet written (connection was lost, data was saved)')
    tellPopup('SYNC_COMPLETE', { success: true, result })

    if (prefs.url) {
      chrome.storage.sync.set({ gas_trigger_preferences: { ...prefs, lastUsed: new Date().toISOString() } })
    }
  }
})

// ── Message router ─────────────────────────────────────────────────────────
chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  switch (message.type) {
    case 'START_SYNC': {
      // Use persisted state (not an in-memory flag) to guard against double-
      // start after a worker restart — in-memory state resets on every kill.
      loadState().then((state) => {
        if (state?.running) {
          sendResponse({ ok: false, reason: 'already_running' })
          return
        }
        startSync(message.payload).catch(console.error)
        sendResponse({ ok: true })
      })
      return true  // async sendResponse
    }

    case 'GET_SYNC_STATE': {
      loadState().then(sendResponse)
      return true
    }

    default:
      return false
  }
})
