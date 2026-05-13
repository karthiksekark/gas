import React, { useState } from 'react'
import styles from './Settings.module.scss'

export default function Settings({ prefs, onUpdate, onClose }) {
  const [localUrl, setLocalUrl]         = useState(prefs.url || '')
  const [localSecret, setLocalSecret]   = useState(prefs.secretKey || '')
  const [localMethod, setLocalMethod]   = useState(prefs.requestMethod || 'GET')
  const [localHeaders, setLocalHeaders] = useState(prefs.customHeaders || '')
  const [jiraBaseUrl, setJiraBaseUrl]   = useState(prefs.jiraBaseUrl || '')
  const [jiraJqlQuery, setJiraJqlQuery] = useState(prefs.jiraJqlQuery || '')
  const [showSecret, setShowSecret]     = useState(false)
  const [saved, setSaved]               = useState(false)

  const handleSave = () => {
    onUpdate({
      url:           localUrl.trim(),
      secretKey:     localSecret.trim(),
      requestMethod: localMethod,
      customHeaders: localHeaders.trim(),
      jiraBaseUrl:   jiraBaseUrl.trim().replace(/\/$/, ''),
      jiraJqlQuery:  jiraJqlQuery.trim(),
    })
    setSaved(true)
    setTimeout(() => { setSaved(false); onClose() }, 800)
  }

  const handleOverlayClick = (e) => {
    if (e.target === e.currentTarget) onClose()
  }

  return (
    <div className={styles.overlay} onClick={handleOverlayClick}>
      <div className={styles.panel}>

        <div className={styles.header}>
          <span className={styles.title}>⚙ Preferences</span>
          <button className={styles.closeBtn} onClick={onClose} aria-label="Close">✕</button>
        </div>

        <div className={styles.body}>

          {/* ── GAS Connection ── */}
          <span className={styles.sectionLabel}>GAS Connection</span>

          <div className={styles.fieldGroup}>
            <label className={styles.label}>Web App URL</label>
            <input
              className={styles.input}
              type="url"
              placeholder="https://script.google.com/macros/s/..."
              value={localUrl}
              onChange={(e) => setLocalUrl(e.target.value)}
            />
          </div>

          <div className={styles.fieldGroup}>
            <label className={styles.label}>
              Secret Key{' '}
              <span className={styles.labelNote}>— sent as key field</span>
            </label>
            <div className={styles.secretWrapper}>
              <input
                className={styles.secretInput}
                type={showSecret ? 'text' : 'password'}
                placeholder="your-secret-key-here"
                value={localSecret}
                onChange={(e) => setLocalSecret(e.target.value)}
              />
              <button
                className={styles.eyeBtn}
                onClick={() => setShowSecret((v) => !v)}
                aria-label={showSecret ? 'Hide secret' : 'Show secret'}
              >
                {showSecret ? '🙈' : '👁'}
              </button>
            </div>
            <span className={styles.hint}>
              Must match{' '}
              <code className={styles.code}>SECRET_KEY</code>
              {' '}in Code.gs
            </span>
          </div>

          <div className={styles.divider} />

          {/* ── JIRA Integration ── */}
          <span className={styles.sectionLabel}>JIRA Integration</span>

          <div className={styles.fieldGroup}>
            <label className={styles.label}>JIRA Base URL</label>
            <input
              className={`${styles.input} ${styles.inputJira}`}
              type="url"
              placeholder="https://yourorg.atlassian.net"
              value={jiraBaseUrl}
              onChange={(e) => setJiraBaseUrl(e.target.value)}
            />
          </div>

          <div className={styles.fieldGroup}>
            <label className={styles.label}>JQL Query</label>
            <textarea
              className={`${styles.textarea} ${styles.textareaJira}`}
              placeholder={'due="{date}" AND project = MYPROJ\nOR\ndue="{date}" AND assignee = currentUser()'}
              value={jiraJqlQuery}
              onChange={(e) => setJiraJqlQuery(e.target.value)}
              spellCheck={false}
            />
            <span className={styles.hint}>
              Use{' '}
              <code className={`${styles.code} ${styles.codeAccent}`}>{'{date}'}</code>
              {' '}as a placeholder — replaced with each date from column A
              (converted to{' '}
              <code className={`${styles.code} ${styles.codeAccent}`}>YYYY-MM-DD</code>
              {' '}EST) at sync time.
            </span>
          </div>

          <div className={styles.divider} />

          {/* ── Request ── */}
          <span className={styles.sectionLabel}>Request</span>

          <div className={styles.fieldGroup}>
            <label className={styles.label}>Default HTTP Method</label>
            <div className={styles.methodRow}>
              {['GET', 'POST'].map((m) => (
                <button
                  key={m}
                  className={`${styles.methodBtn} ${localMethod === m ? styles.methodBtnActive : ''}`}
                  onClick={() => setLocalMethod(m)}
                >
                  {m}
                </button>
              ))}
            </div>
          </div>

          <div className={styles.fieldGroup}>
            <label className={styles.label}>
              Custom Headers{' '}
              <span className={styles.labelNote}>(optional)</span>
            </label>
            <textarea
              className={styles.textarea}
              placeholder="Content-Type: application/json"
              value={localHeaders}
              onChange={(e) => setLocalHeaders(e.target.value)}
            />
          </div>
        </div>

        <div className={styles.footer}>
          {saved && <span className={styles.savedBadge}>✓ Saved!</span>}
          <button className={styles.cancelBtn} onClick={onClose}>Cancel</button>
          <button className={styles.saveBtn} onClick={handleSave}>Save Changes</button>
        </div>
      </div>
    </div>
  )
}
