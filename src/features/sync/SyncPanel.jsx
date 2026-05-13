import React, { useState } from 'react'
import moment from 'moment-timezone'
import styles from './SyncPanel.module.scss'

const JIRA_TZ  = 'America/New_York'
const JIRA_MAX = 50

async function getJiraSessionCookie(jiraBaseUrl) {
  if (typeof chrome === 'undefined' || !chrome.cookies) return null
  try {
    const url    = jiraBaseUrl.startsWith('http') ? jiraBaseUrl : 'https://' + jiraBaseUrl
    const cookie = await chrome.cookies.get({ url, name: 'JSESSIONID' })
    return cookie ? cookie.value : null
  } catch (err) {
    console.warn('Could not read JSESSIONID cookie:', err)
    return null
  }
}

function buildGetUrl(url, action, secretKey) {
  const sep = url.includes('?') ? '&' : '?'
  return `${url}${sep}action=${action}${secretKey ? `&key=${encodeURIComponent(secretKey)}` : ''}`
}

async function parseResponse(res) {
  const text = await res.text()
  try { return JSON.parse(text) }
  catch { return { success: false, error: text || 'Invalid response', code: res.status } }
}

export default function SyncPanel({ url, secretKey, jiraBaseUrl, jiraJqlQuery, onSaveLastUsed }) {
  const [loading, setLoading]               = useState(false)
  const [result, setResult]                 = useState(null)
  const [syncStats, setSyncStats]           = useState(null)
  const [syncProgress, setSyncProgress]     = useState(0)
  const [syncStatus, setSyncStatus]         = useState('')
  const [detectedDates, setDetectedDates]   = useState([])
  const [cookieFound, setCookieFound]       = useState(null)

  const isJiraConfigured = !!(jiraBaseUrl && url)

  const postBody = (payload) =>
    JSON.stringify({ ...payload, key: secretKey || undefined })

  const doJiraSync = async () => {
    if (!isJiraConfigured) return
    setLoading(true)
    setResult(null)
    setSyncStats(null)
    setSyncProgress(0)
    setSyncStatus('Reading dates from sheet...')

    try {
      // Step 1 — get dates from column A
      const datesRes  = await fetch(buildGetUrl(url, 'getDates', secretKey))
      const datesData = await parseResponse(datesRes)

      if (datesData?.error === 'Unauthorized' || datesData?.code === 401) {
        setResult({ success: false, error: 'Secret key mismatch. Check ⚙ Preferences.', code: 401 })
        return
      }
      if (!datesData.success || !Array.isArray(datesData.dates) || !datesData.dates.length) {
        setResult({ success: false, error: 'No valid dates found in column A of the sheet.', code: 404 })
        return
      }

      const rawDates = datesData.dates
      setDetectedDates(rawDates)
      setSyncProgress(10)

      // Step 2 — fetch JIRA for each date
      const jsessionId = await getJiraSessionCookie(jiraBaseUrl)
      setCookieFound(!!jsessionId)
      const jiraHeaders = { Accept: 'application/json' }
      if (jsessionId) {
        jiraHeaders['Cookie'] = `JSESSIONID=${jsessionId}`
      }

      const issuesByDate = {}
      const perDate      = []

      for (let di = 0; di < rawDates.length; di++) {
        const rawDate = rawDates[di]
        const m = moment.tz(rawDate, ['M/D/YYYY', 'MM/DD/YYYY'], true, JIRA_TZ)
        if (!m.isValid()) { perDate.push({ date: rawDate, skipped: true }); continue }

        const jiraDate = m.format('YYYY-MM-DD')
        setSyncStatus(`[${di + 1}/${rawDates.length}] Fetching JIRA for ${jiraDate}...`)
        setSyncProgress(10 + Math.round((di / rawDates.length) * 50))

        const jqlTemplate = jiraJqlQuery?.trim() || 'due="{date}" ORDER BY created ASC'
        const resolvedJql = jqlTemplate.replace(/\{date\}/g, jiraDate)
        const jql         = encodeURIComponent(resolvedJql)

        let allIssues = [], startAt = 0, total = null
        while (true) {
          const jiraUrl = `${jiraBaseUrl}/rest/api/3/search?jql=${jql}&fields=summary,status,duedate&maxResults=${JIRA_MAX}&startAt=${startAt}`
          const jiraRes = await fetch(jiraUrl, { headers: jiraHeaders, credentials: 'include' })

          if (!jiraRes.ok) {
            setResult({ success: false, error: `JIRA ${jiraRes.status} for ${jiraDate}`, code: jiraRes.status })
            return
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

      setSyncStatus('Writing to sheet...')
      setSyncProgress(65)

      // Step 3 — single GAS call with all issues
      const gasRes  = await fetch(url, {
        method:  'POST',
        headers: { 'Content-Type': 'text/plain' },
        body:    postBody({ action: 'syncJira', issuesByDate }),
      })
      const gasData = await parseResponse(gasRes)
      setSyncProgress(100)

      if (gasData.success) {
        const st = gasData.stats || {}
        setSyncStats({ ...st, perDate })
        setResult({
          success: true,
          message: `Sync complete — ${st.inserted || 0} inserted, ${st.updated || 0} updated, ${st.moved || 0} moved, ${st.skipped || 0} skipped`,
        })
        onSaveLastUsed()
      } else {
        setResult({ success: false, error: gasData.error || 'Sync failed', code: gasData.code })
      }
    } catch (err) {
      setResult({ success: false, error: err.message || 'Network error' })
    } finally {
      setLoading(false)
      setSyncProgress(0)
      setSyncStatus('')
    }
  }

  const cookieStatusClass =
    cookieFound === true  ? styles.cookieFound :
    cookieFound === false ? styles.cookieMissing :
    styles.cookieUnknown

  const cookieStatusText =
    cookieFound === true  ? '✓ found' :
    cookieFound === false ? '✕ not found' :
    '— checked on sync'

  return (
    <div className={styles.wrap}>

      {/* Config card */}
      {isJiraConfigured ? (
        <div className={styles.configCard}>
          <div className={styles.configRow}>
            <span className={styles.configLabel}>JIRA URL</span>
            <span className={`${styles.configVal} ${styles.configValAccent}`}>{jiraBaseUrl}</span>
          </div>
          <div className={styles.configRow}>
            <span className={styles.configLabel}>Dates from</span>
            <span className={styles.configVal}>column A of sheet</span>
          </div>
          <div className={styles.configRow}>
            <span className={styles.configLabel}>Timezone</span>
            <span className={styles.configVal}>America/New_York (EST)</span>
          </div>
          <div className={styles.configRow}>
            <span className={styles.configLabel}>JSESSIONID</span>
            <span className={`${styles.configVal} ${cookieStatusClass}`}>{cookieStatusText}</span>
          </div>
          <div className={styles.divider} />
          <div className={styles.jqlLabel}>JQL Template</div>
          <div className={styles.jqlBlock}>
            {jiraJqlQuery?.trim()
              ? jiraJqlQuery.trim()
              : <span className={styles.jqlDefault}>{'due="{date}" ORDER BY created ASC (default)'}</span>
            }
          </div>
        </div>
      ) : (
        <div className={styles.warning}>
          ⚠ JIRA Base URL or Web App URL not configured.
          <br />
          Open ⚙ Preferences to set them up.
        </div>
      )}

      {/* Detected dates */}
      {detectedDates.length > 0 && (
        <div className={styles.datesWrap}>
          <div className={styles.datesLabel}>Dates found in column A</div>
          <div className={styles.pills}>
            {detectedDates.map((d) => (
              <span key={d} className={styles.pill}>{d}</span>
            ))}
          </div>
        </div>
      )}

      {/* Progress */}
      {loading && syncStatus && (
        <div className={styles.progressWrap}>
          <div className={styles.progressStatus}>{syncStatus}</div>
          <div className={styles.progressBar}>
            <div className={styles.progressFill} style={{ width: `${syncProgress}%` }} />
          </div>
        </div>
      )}

      {/* Sync button */}
      <button
        className={`${styles.syncBtn} ${(!isJiraConfigured || loading) ? styles.syncBtnDisabled : ''}`}
        onClick={doJiraSync}
        disabled={!isJiraConfigured || loading}
      >
        {loading ? (
          <>
            <div className={styles.btnSpinner} />
            {syncStatus || 'Syncing...'}
          </>
        ) : (
          '⟳ Sync from JIRA'
        )}
      </button>

      {/* Stats */}
      {syncStats && (
        <div className={styles.statsCard}>
          <div className={styles.statsTitle}>Sync results</div>
          <div className={styles.statsRow}>
            <span className={styles.statsLabel}>Inserted</span>
            <span className={`${styles.statsVal} ${styles.statsValSuccess}`}>{syncStats.inserted ?? 0}</span>
          </div>
          <div className={styles.statsRow}>
            <span className={styles.statsLabel}>Updated</span>
            <span className={`${styles.statsVal} ${styles.statsValWarning}`}>{syncStats.updated ?? 0}</span>
          </div>
          <div className={styles.statsRow}>
            <span className={styles.statsLabel}>Moved</span>
            <span className={`${styles.statsVal} ${styles.statsValMoved}`}>{syncStats.moved ?? 0}</span>
          </div>
          <div className={styles.statsRow}>
            <span className={styles.statsLabel}>Skipped</span>
            <span className={styles.statsVal}>{syncStats.skipped ?? 0}</span>
          </div>
          {syncStats.perDate?.length > 0 && (
            <>
              <div className={styles.divider} />
              {syncStats.perDate.map((pd, i) => (
                <div key={i} className={styles.perDateRow}>
                  <span className={styles.perDateKey}>{pd.date}</span>
                  <span>{pd.issues ?? 0} issues</span>
                </div>
              ))}
            </>
          )}
        </div>
      )}

      {/* Result messages */}
      {result?.success && (
        <div className={styles.successMsg}>
          <span>✓ {result.message}</span>
          <button className={styles.dismissBtn} onClick={() => setResult(null)}>✕</button>
        </div>
      )}
      {result && !result.success && (
        <div className={`${styles.errorMsg} ${(result.code === 401 || result.code === 403) ? styles.errorMsgAuth : styles.errorMsgNetwork}`}>
          <span className={styles.errorTitle}>
            {result.code ? `${result.code} Error` : 'Error'}
          </span>
          {result.error}
        </div>
      )}
    </div>
  )
}
