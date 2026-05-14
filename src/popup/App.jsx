import React, { useState, useCallback } from 'react'
import { useStorage } from '../hooks/useStorage'
import SyncPanel from '../features/sync/SyncPanel'
import Settings from '../features/settings/Settings'
import styles from './App.module.scss'

function formatLastUsed(iso) {
  if (!iso) return null
  const d = new Date(iso)
  const diff = Date.now() - d
  if (diff < 60_000) return 'just now'
  if (diff < 3_600_000) return `${Math.floor(diff / 60_000)}m ago`
  if (diff < 86_400_000) return `${Math.floor(diff / 3_600_000)}h ago`
  return d.toLocaleDateString()
}

export default function App() {
  const { prefs, loaded, setPrefs } = useStorage()
  const [showSettings, setShowSettings] = useState(false)

  const handlePrefsUpdate = useCallback((updates) => {
    if (typeof setPrefs === 'function') setPrefs((p) => ({ ...p, ...updates }))
  }, [setPrefs])

  if (!loaded) {
    return (
      <div className={styles.page}>
        <div className={styles.loadingWrap}>
          <div className={styles.spinner} />
        </div>
      </div>
    )
  }

  return (
    <div className={styles.page}>
      <header className={styles.header}>
        <div className={styles.logoRow}>
          <div className={styles.logoMark}>⚡</div>
          <div>
            <div className={styles.title}>GAS Trigger</div>
            <div className={styles.subtitle}>JIRA → Google Sheet Sync</div>
          </div>
        </div>
        <button
          className={styles.settingsBtn}
          onClick={() => setShowSettings(true)}
          aria-label="Open preferences"
        >
          ⚙
        </button>
      </header>

      <main className={styles.body}>
        <SyncPanel
          url={prefs.url}
          secretKey={prefs.secretKey || ''}
          jiraBaseUrl={prefs.jiraBaseUrl || ''}
          jiraJqlQuery={prefs.jiraJqlQuery || ''}
        />
      </main>

      <footer className={styles.footer}>
        <span className={styles.version}>v2.0.0</span>
        {prefs.lastUsed && (
          <span className={styles.lastUsed}>
            Last sync: {formatLastUsed(prefs.lastUsed)}
          </span>
        )}
      </footer>

      {showSettings && (
        <Settings
          prefs={prefs}
          onUpdate={handlePrefsUpdate}
          onClose={() => setShowSettings(false)}
        />
      )}
    </div>
  )
}
