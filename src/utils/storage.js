import { ROUTES, normalizeRoute } from '../config/routes'

export const STORAGE_PREFIX = 'pptxpro:slides:v1'
export const DB_NAME = 'pptxpro-slides'
export const DB_STORE = 'pairs'
export const DB_VERSION = 1

// Master preset DB keys
export const MASTER_PRESETS_KEY = 'pptxpro:master-presets'
export const MASTER_ACTIVE_KEY = 'pptxpro:master-active-preset-id'
export const MASTER_LEGACY_KEY = 'pptxpro:custom-master-layout'

export const buildStorageKey = (route, variant = '') => {
  const normalizedRoute = normalizeRoute(route)
  const resolvedVariant =
    normalizedRoute === ROUTES.dailyPlot ? variant || 'urban' : 'default'
  return `${STORAGE_PREFIX}:${normalizedRoute}:${resolvedVariant}`
}

export const canUseStorage = () => {
  try {
    return typeof window !== 'undefined' && Boolean(window.localStorage)
  } catch {
    return false
  }
}

export const canUseIndexedDb = () => {
  try {
    return typeof window !== 'undefined' && Boolean(window.indexedDB)
  } catch {
    return false
  }
}

export const openPairsDb = () =>
  new Promise((resolve, reject) => {
    if (!canUseIndexedDb()) {
      resolve(null)
      return
    }
    const request = window.indexedDB.open(DB_NAME, DB_VERSION)
    request.onupgradeneeded = () => {
      const db = request.result
      if (!db.objectStoreNames.contains(DB_STORE)) {
        db.createObjectStore(DB_STORE)
      }
    }
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error)
  })

export const loadPairsFromDb = async (storageKey) => {
  if (!storageKey || !canUseIndexedDb()) {
    return null
  }
  try {
    const db = await openPairsDb()
    if (!db) {
      return null
    }
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, 'readonly')
      const store = tx.objectStore(DB_STORE)
      const request = store.get(storageKey)
      request.onsuccess = () => resolve(request.result ?? null)
      request.onerror = () => reject(request.error)
      tx.oncomplete = () => db.close()
    })
  } catch {
    return null
  }
}

export const savePairsToDb = async (storageKey, pairs) => {
  if (!storageKey || !canUseIndexedDb()) {
    return
  }
  try {
    const db = await openPairsDb()
    if (!db) {
      return
    }
    await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, 'readwrite')
      const store = tx.objectStore(DB_STORE)
      store.put(pairs, storageKey)
      tx.oncomplete = () => {
        db.close()
        resolve()
      }
      tx.onerror = () => {
        db.close()
        reject(tx.error)
      }
    })
  } catch {
    // Ignore storage errors.
  }
}

export const removePairsFromDb = async (storageKey) => {
  if (!storageKey || !canUseIndexedDb()) {
    return
  }
  try {
    const db = await openPairsDb()
    if (!db) {
      return
    }
    await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, 'readwrite')
      const store = tx.objectStore(DB_STORE)
      store.delete(storageKey)
      tx.oncomplete = () => {
        db.close()
        resolve()
      }
      tx.onerror = () => {
        db.close()
        reject(tx.error)
      }
    })
  } catch {
    // Ignore storage errors.
  }
}

export const loadStoredPairsSync = (storageKey) => {
  if (!storageKey || !canUseStorage()) {
    return null
  }
  try {
    const raw = window.localStorage.getItem(storageKey)
    if (!raw) {
      return null
    }
    const parsed = JSON.parse(raw)
    return Array.isArray(parsed) ? parsed : null
  } catch {
    return null
  }
}

export const loadStoredPairs = async (storageKey) => {
  return await loadPairsFromDb(storageKey)
}

export const buildStoredPairs = (pairs, slotKeys) => {
  if (!Array.isArray(pairs)) {
    return []
  }
  return pairs.map((pair) => {
    const nextPair = {
      ...pair,
      slideText: typeof pair?.slideText === 'string' ? pair.slideText : '',
    }
    slotKeys.forEach((key) => {
      nextPair[key] = typeof pair?.[key] === 'string' ? pair[key] : ''
    })
    return nextPair
  })
}

export const savePairsToStorage = (storageKey, pairs, { slotKeys, requiresText, hasPairContent, trimTrailingEmptyPairs }) => {
  if (!storageKey) {
    return
  }
  try {
    const prepared = buildStoredPairs(pairs, slotKeys)
    const trimmed = trimTrailingEmptyPairs(prepared, { slotKeys, requiresText })
    if (trimmed.length === 0) {
      void removePairsFromDb(storageKey)
      return
    }
    void savePairsToDb(storageKey, trimmed)
  } catch {
    // Ignore storage errors.
  }
}

// ── Master Preset persistence ───────────────────────────────────────────────

export const loadMasterPresets = async () => {
  try {
    const data = await loadPairsFromDb(MASTER_PRESETS_KEY)
    return Array.isArray(data) ? data : []
  } catch {
    return []
  }
}

export const saveMasterPresets = async (presets) => {
  try {
    await savePairsToDb(MASTER_PRESETS_KEY, presets)
  } catch {
    // Ignore
  }
}

export const loadActivePresetId = async () => {
  try {
    const id = await loadPairsFromDb(MASTER_ACTIVE_KEY)
    return typeof id === 'string' ? id : null
  } catch {
    return null
  }
}

export const saveActivePresetId = async (id) => {
  try {
    if (id) {
      await savePairsToDb(MASTER_ACTIVE_KEY, id)
    } else {
      await removePairsFromDb(MASTER_ACTIVE_KEY)
    }
  } catch {
    // Ignore
  }
}

export const migrateLegacyLayout = async (presets) => {
  try {
    const legacy = await loadPairsFromDb(MASTER_LEGACY_KEY)
    if (legacy && typeof legacy === 'object') {
      const already = presets.some((p) => p._migrated)
      if (!already) {
        const newPreset = {
          id: `preset_migrated_${Date.now()}`,
          name: 'Default',
          createdAt: Date.now(),
          _migrated: true,
          layout: legacy,
        }
        return [newPreset, ...presets]
      }
    }
  } catch {
    // Ignore
  }
  return null // no migration needed
}

// ── PPTX Editor persistence ──────────────────────────────────────────────────
export const PPTX_EDITOR_KEY = 'pptxpro:pptx-editor:v1'

export const loadPptxEditorState = async () => {
  try {
    const data = await loadPairsFromDb(PPTX_EDITOR_KEY)
    if (!data || !data.parsedData) return null
    let file = null
    if (data.fileBuffer && data.filename) {
      file = new File([data.fileBuffer], data.filename, {
        type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      })
    }
    return {
      file,
      parsedData: data.parsedData,
      activeSlideIndex: data.activeSlideIndex || 0,
    }
  } catch {
    return null
  }
}

export const savePptxEditorState = async ({ fileBuffer, filename, parsedData, activeSlideIndex }) => {
  try {
    await savePairsToDb(PPTX_EDITOR_KEY, {
      filename,
      fileBuffer,
      parsedData,
      activeSlideIndex,
    })
  } catch {
    // Ignore
  }
}

export const clearPptxEditorState = async () => {
  try {
    await removePairsFromDb(PPTX_EDITOR_KEY)
  } catch {
    // Ignore
  }
}

