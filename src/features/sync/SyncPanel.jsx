import React, { useState, useEffect, useCallback } from 'react'
import styles from './SyncPanel.module.scss'

const chromeRuntime = () => typeof chrome !== 'undefined' && chrome.runtime

function sendToWorker(type, payload = {}) {
  return new Promise((resolve) => {
    if (!chromeRuntime()) return resolve(null)
    chrome.runtime.sendMessage({ type, payload }, (response) => {
      if (chrome.runtime.lastError) return resolve(null)
      resolve(response)
    })
  })
}

export default function SyncPanel({ url, secretKey, sheetName, jiraBaseUrl, jiraJqlQuery }) {
  const [loading, setLoading]             = useState(false)
  const [phase, setPhase]                 = useState('idle') // 'idle' | 'syncing' | 'reverting'
  const [syncProgress, setSyncProgress]   = useState(0)
  const [syncStatus, setSyncStatus]       = useState('')
  const [result, setResult]               = useState(null)
  const [syncStats, setSyncStats]         = useState(null)
  const [detectedDates, setDetectedDates] = useState([])
  const [cookieFound, setCookieFound]     = useState(null)

  // Hydrate from persisted worker state when popup opens
  useEffect(() => {
    if (!chromeRuntime()) return

    sendToWorker('GET_SYNC_STATE').then((state) => {
      if (!state) return
      if (state.running) {
        setLoading(true)
        setPhase(state.phase === 'reverting' ? 'reverting' : 'syncing')
        setSyncProgress(state.progress || 0)
        setSyncStatus(state.status || 'Syncing…')
        if (state.dates) setDetectedDates(state.dates)
        if (state.cookieFound != null) setCookieFound(state.cookieFound)
      } else if (state.result) {
        applyResult(state.result)
      }
    })

    // Listen for live updates from the worker
    const onMessage = (message) => {
      if (message.type === 'SYNC_PROGRESS') {
        const { progress, status, dates, cookieFound: cf, phase: ph } = message.payload
        setSyncProgress(progress)
        setSyncStatus(status)
        if (ph) setPhase(ph)
        if (dates) setDetectedDates(dates)
        if (cf != null) setCookieFound(cf)
      } else if (message.type === 'SYNC_COMPLETE') {
        setLoading(false)
        setPhase('idle')
        setSyncProgress(0)
        setSyncStatus('')
        if (message.payload.cancelled) {
          if (message.payload.revertFailed) {
            setResult({
              success: false,
              error:   'Revert failed — the _snapshot tab in your sheet was preserved for manual recovery.',
            })
          } else {
            setResult(null)
          }
        } else {
          applyResult(message.payload)
        }
      }
    }

    chrome.runtime.onMessage.addListener(onMessage)
    return () => chrome.runtime.onMessage.removeListener(onMessage)
  }, [])

  function applyResult(payload) {
    if (payload.success) {
      setSyncStats(payload.result?.stats ?? payload.stats ?? null)
      setResult({ success: true, message: payload.result?.message ?? payload.message ?? 'Sync complete' })
    } else if (payload.success === false) {
      setResult({ success: false, error: payload.error, code: payload.code })
    }
  }

  const isJiraConfigured = !!(jiraBaseUrl && url)

  const doJiraSync = useCallback(async () => {
    if (!isJiraConfigured || loading) return
    setLoading(true)
    setPhase('syncing')
    setResult(null)
    setSyncStats(null)
    setSyncProgress(0)
    setSyncStatus('Starting sync…')
    setDetectedDates([])
    setCookieFound(null)

    const response = await sendToWorker('START_SYNC', { url, secretKey, sheetName, jiraBaseUrl, jiraJqlQuery })
    if (!response?.ok) {
      setLoading(false)
      const reason = response?.reason === 'already_running'
        ? 'A sync is already running.'
        : 'Could not reach background worker.'
      setResult({ success: false, error: reason })
    }
  }, [isJiraConfigured, loading, url, secretKey, sheetName, jiraBaseUrl, jiraJqlQuery])

  const doCancelSync = useCallback(() => {
    if (phase === 'reverting') return  // already reverting — ignore
    sendToWorker('CANCEL_SYNC')
    setPhase('reverting')
    setSyncStatus('Cancelling…')
  }, [phase])

  const cookieClass =
    cookieFound === true  ? styles.cookieFound  :
    cookieFound === false ? styles.cookieMissing :
    styles.cookieUnknown

  const cookieText =
    cookieFound === true  ? '✓ found'       :
    cookieFound === false ? '✕ not found'   :
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
            <span className={`${styles.configVal} ${cookieClass}`}>{cookieText}</span>
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

      {/* Background-sync warning banner */}
      {loading && phase !== 'reverting' && (
        <div className={styles.bgWarning}>
          <span className={styles.bgWarningIcon}>⟳</span>
          <span>
            Sync is running in the background — safe to close this popup.
            Watch the extension icon for live status.
          </span>
        </div>
      )}

      {/* Reverting banner */}
      {loading && phase === 'reverting' && (
        <div className={styles.revertWarning}>
          <span className={styles.revertWarningIcon}>↩</span>
          <span>Reverting sheet to pre-sync state — please wait.</span>
        </div>
      )}

      {/* Action button — Sync, Cancel, or locked Reverting */}
      {loading && phase === 'reverting' ? (
        <button className={`${styles.cancelBtn} ${styles.cancelBtnDisabled}`} disabled>
          ↩ Reverting…
        </button>
      ) : loading ? (
        <button className={styles.cancelBtn} onClick={doCancelSync}>
          ✕ Cancel Sync
        </button>
      ) : (
        <button
          className={`${styles.syncBtn} ${!isJiraConfigured ? styles.syncBtnDisabled : ''}`}
          onClick={doJiraSync}
          disabled={!isJiraConfigured}
        >
          ⟳ Sync from JIRA
        </button>
      )}

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
