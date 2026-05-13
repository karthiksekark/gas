import { useState, useEffect, useCallback } from 'react'

const STORAGE_KEY = 'gas_trigger_preferences'

const defaultPrefs = {
  url: '',
  lastUsed: null,
  requestMethod: 'GET',
  customHeaders: '',
  requestBody: '',
  secretKey: '',
  jiraBaseUrl: '',
  jiraJqlQuery: '',
}

const storage = {
  get(key) {
    if (typeof chrome !== 'undefined' && chrome.storage) {
      return new Promise((resolve) => {
        chrome.storage.sync.get([key], (result) => resolve(result[key]))
      })
    }
    try {
      const val = localStorage.getItem(key)
      return Promise.resolve(val ? JSON.parse(val) : undefined)
    } catch {
      return Promise.resolve(undefined)
    }
  },

  set(key, value) {
    if (typeof chrome !== 'undefined' && chrome.storage) {
      return new Promise((resolve) => {
        chrome.storage.sync.set({ [key]: value }, resolve)
      })
    }
    try {
      localStorage.setItem(key, JSON.stringify(value))
      return Promise.resolve()
    } catch {
      return Promise.resolve()
    }
  },
}

export function useStorage() {
  const [prefs, setPrefsState] = useState(defaultPrefs)
  const [loaded, setLoaded] = useState(false)

  useEffect(() => {
    storage.get(STORAGE_KEY).then((saved) => {
      if (saved) setPrefsState((prev) => ({ ...prev, ...saved }))
      setLoaded(true)
    })
  }, [])

  const setPrefs = useCallback(async (updates) => {
    const next = typeof updates === 'function' ? updates(prefs) : { ...prefs, ...updates }
    setPrefsState(next)
    await storage.set(STORAGE_KEY, next)
  }, [prefs])

  const saveLastUsed = useCallback(
    () => setPrefs({ lastUsed: new Date().toISOString() }),
    [setPrefs]
  )

  return { prefs, loaded, saveLastUsed, setPrefs }
}
