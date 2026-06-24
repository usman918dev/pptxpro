import { useEffect, useMemo, useRef, useState } from 'react'
import './App.css'
import { generateReport } from './report/generateReport'
import { importPptxSlides } from './report/importPptx'
import { MasterDesigner } from './MasterDesigner'
import { ImageExtractor } from './ImageExtractor'

const EMPTY_PAIR = {
  beforeImage: '',
  middleImage: '',
  afterImage: '',
  slideText: '',
}

const normalizeRoute = (path = '/') => {
  const trimmed = path.replace(/\/+$/, '')
  return trimmed === '' ? '/' : trimmed
}

const ROUTES = {
  clean: '/',
  compliance: '/compliance',
  desilting: '/desilting',
  dailyPlot: '/daily-plot',
  master: '/master',
  extract: '/extract',
}

const STORAGE_PREFIX = 'pptxpro:slides:v1'
const DB_NAME = 'pptxpro-slides'
const DB_STORE = 'pairs'
const DB_VERSION = 1

// Master preset DB keys
const MASTER_PRESETS_KEY = 'pptxpro:master-presets'
const MASTER_ACTIVE_KEY = 'pptxpro:master-active-preset-id'
const MASTER_LEGACY_KEY = 'pptxpro:custom-master-layout'

const buildStorageKey = (route, variant = '') => {
  const normalizedRoute = normalizeRoute(route)
  const resolvedVariant =
    normalizedRoute === ROUTES.dailyPlot ? variant || 'urban' : 'default'
  return `${STORAGE_PREFIX}:${normalizedRoute}:${resolvedVariant}`
}

const canUseStorage = () => {
  try {
    return typeof window !== 'undefined' && Boolean(window.localStorage)
  } catch {
    return false
  }
}

const canUseIndexedDb = () => {
  try {
    return typeof window !== 'undefined' && Boolean(window.indexedDB)
  } catch {
    return false
  }
}

const openPairsDb = () =>
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

const loadPairsFromDb = async (storageKey) => {
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

const savePairsToDb = async (storageKey, pairs) => {
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

const removePairsFromDb = async (storageKey) => {
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

const loadStoredPairsSync = (storageKey) => {
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

const loadStoredPairs = async (storageKey) => {
  return await loadPairsFromDb(storageKey)
}

const buildStoredPairs = (pairs, slotKeys) => {
  if (!Array.isArray(pairs)) {
    return []
  }
  return pairs.map((pair) => {
    const nextPair = {
      slideText: typeof pair?.slideText === 'string' ? pair.slideText : '',
    }
    slotKeys.forEach((key) => {
      nextPair[key] = typeof pair?.[key] === 'string' ? pair[key] : ''
    })
    return nextPair
  })
}

const trimTrailingEmptyPairs = (pairs, { slotKeys, requiresText }) => {
  const trimmed = [...pairs]
  while (trimmed.length > 0) {
    const last = trimmed[trimmed.length - 1]
    const hasImages = slotKeys.some((key) => Boolean(last?.[key]))
    const hasText = requiresText
      ? typeof last?.slideText === 'string' && last.slideText.trim()
      : false
    if (hasImages || hasText) {
      break
    }
    trimmed.pop()
  }
  return trimmed
}

const savePairsToStorage = (storageKey, pairs, { slotKeys, requiresText }) => {
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

const DEFAULT_SLOTS = [
  {
    key: 'beforeImage',
    label: 'Before',
    className: 'slide-slot slide-slot--before',
    urlMode: 'prompt',
  },
  {
    key: 'afterImage',
    label: 'After',
    className: 'slide-slot slide-slot--after',
    urlMode: 'prompt',
  },
]

const DESILTING_PRESET_TEXT = [
  'UC-55',
  'UC-56',
  'UC-57',
  'UC-58',
  'UC-59',
  'UC-60',
  'UC-61',
  'UC-62',
  'UC-63',
  'UC-64',
  'UC-65',
  'UC-66',
  'UC-67',
  'UC-68',
  'UC-69',
  'Sector-1',
  'Sector-2',
  'Sector-3',
  'Sector-4',
  'Sector-5',
  'Sector-6',
  'Sector-7',
  'Sector-8',
  'Sector-9',
  'Sector-10',
]

const getDateOffset = (offsetDays = 0) => {
  const date = new Date()
  date.setDate(date.getDate() + offsetDays)
  return date
}

const formatDailyPlotDate = (date) => {
  const day = String(date.getDate()).padStart(2, '0')
  const monthNames = [
    'Jan',
    'Feb',
    'Mar',
    'Apr',
    'May',
    'Jun',
    'Jul',
    'Aug',
    'Sep',
    'Oct',
    'Nov',
    'Dec',
  ]
  const month = monthNames[date.getMonth()]
  const year = date.getFullYear()
  return `${day}_${month}_${year}`
}

const buildDailyPlotFileName = (variant = 'urban') => {
  const dateLabel = formatDailyPlotDate(getDateOffset(-1))
  const suffix = variant === 'rural' ? 'Rural' : 'Urban'
  return `Daily_Plot's_Clearance_Report_${suffix} ${dateLabel}.pptx`
}

const TEMPLATES = {
  [ROUTES.clean]: {
    eyebrow: 'Clean Punjab',
    title: 'Plots Cleaning-Activity',
    subtext:
      'Drop a Before and After image. A new row appears automatically and the page scrolls to it.',
    masterBgUrl: '/master-slide.png',
    firstSlideUrl: '/first-slide.png',
    secondSlideUrl: '',
    lastSlideUrl: '/last-slide.png',
    fileNamePrefix: 'Plots  Cleaning-Activity_Tehsil Kamoke',
    masterTitle: 'CLEAN_PUNJAB_MASTER',
    slideTitle: 'Plots Cleaning-Activity',
    themeLabel: 'Master slide theme',
    importSkipFirst: 1,
    importSkipLast: 1,
    imageCount: 2,
    slots: DEFAULT_SLOTS,
  },
  [ROUTES.compliance]: {
    eyebrow: 'Suthra Punjab',
    title: 'Compliance Report',
    subtext:
      'Drop Before and After images. A new row appears automatically and the page scrolls to it.',
    masterBgUrl: '/master-slide-2.png',
    firstSlideUrl: '/slide-first-2.png',
    secondSlideUrl: '/slide-second-2.png',
    lastSlideUrl: '/last-slide-2.png',
    fileNamePrefix: 'Compliance Report_Tehsil Kamoke',
    masterTitle: 'COMPLIANCE_MASTER',
    slideTitle: 'Compliance Report',
    themeLabel: 'Compliance master slide theme',
    leftBox: { x: 0.4, y: 1.5, w: 6.0, h: 4.65 },
    rightBox: { x: 6.93, y: 1.5, w: 6.0, h: 4.65 },
    dateSlide: 'second',
    dateBox: { x: 5.07, y: 4.91, w: 3.2, h: 0.4 },
    dateColor: 'FFD54D',
    dateAlign: 'center',
    dateFontSize: 22,
    dateBold: true,
    importSkipFirst: 2,
    importSkipLast: 1,
    imageCount: 2,
    slots: DEFAULT_SLOTS,
  },
  [ROUTES.desilting]: {
    eyebrow: 'Suthra Punjab',
    title: 'Desilting Report',
    subtext:
      'Drop three images for each slide and add the sector text. A new row appears automatically and the page scrolls to it.',
    masterBgUrl: '/master-slide-3.png',
    firstSlideUrl: '/first-slide-3.png',
    secondSlideUrl: '',
    lastSlideUrl: '/last-slide-3.png',
    fileNamePrefix: 'Desilting Report_Tehsil Kamoke',
    masterTitle: 'DESILTING_MASTER',
    slideTitle: 'Desilting Report',
    themeLabel: 'Desilting master slide theme',
    leftBox: { x: 0.16, y: 1.9, w: 4.27, h: 4.975 },
    middleBox: { x: 4.533, y: 1.9, w: 4.27, h: 4.975 },
    rightBox: { x: 8.88, y: 1.9, w: 4.27, h: 4.975 },
    textBox: { x: 4.0, y: 0.88, w: 5.33, h: 0.5 },
    textColor: '111111',
    textAlign: 'center',
    textFontSize: 22,
    textBold: true,
    textDefault: '',
    textLabel: 'Sector text',
    importSkipFirst: 1,
    importSkipLast: 1,
    imageCount: 3,
    slots: [
      {
        key: 'beforeImage',
        label: 'Before',
        className: 'slide-slot slide-slot--before',
        urlMode: 'prompt',
      },
      {
        key: 'middleImage',
        label: 'During',
        className: 'slide-slot slide-slot--middle',
        urlMode: 'prompt',
      },
      {
        key: 'afterImage',
        label: 'After',
        className: 'slide-slot slide-slot--after',
        urlMode: 'prompt',
      },
    ],
  },
  [ROUTES.dailyPlot]: {
    eyebrow: 'Daily Plot',
    title: 'OTC Plot Report',
    subtext:
      'Drop a Before and After image. A new row appears automatically and the page scrolls to it.',
    masterBgUrl: '/master-slide-4.png',
    firstSlideUrl: '/first-slide-4-urban.png',
    firstSlideUrlRural: '/first-slide-4-rural.png',
    secondSlideUrl: '',
    lastSlideUrl: '/last-slide-4.png',
    fileName: buildDailyPlotFileName,
    masterTitle: 'DAILY_PLOT_MASTER',
    slideTitle: 'OTC Plot Report',
    themeLabel: 'Daily plot master slide theme',
    leftBox: { x: 1.0, y: 2.16, w: 5.47, h: 4.05 },
    rightBox: { x: 6.8, y: 2.16, w: 5.47, h: 4.05 },
    dateBox: { x: 5.80, y: 2.25, w: 2.4, h: 0.4 },
    dateAlign: 'center',
    dateColor: '000000',
    dateFontSize: 20,
    dateFontFace: 'Times New Roman',
    dateSlide: 'first',
    dateOffsetDays: -1,
    importSkipFirst: 1,
    importSkipLast: 1,
    imageCount: 2,
    slots: DEFAULT_SLOTS,
  },
}

// ── Master Preset helpers ────────────────────────────────────────────────────

const loadMasterPresets = async () => {
  try {
    const data = await loadPairsFromDb(MASTER_PRESETS_KEY)
    return Array.isArray(data) ? data : []
  } catch {
    return []
  }
}

const saveMasterPresets = async (presets) => {
  try {
    await savePairsToDb(MASTER_PRESETS_KEY, presets)
  } catch {
    // Ignore
  }
}

const loadActivePresetId = async () => {
  try {
    const id = await loadPairsFromDb(MASTER_ACTIVE_KEY)
    return typeof id === 'string' ? id : null
  } catch {
    return null
  }
}

const saveActivePresetId = async (id) => {
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

// Migrate old single layout to preset system (runs once)
const migrateLegacyLayout = async (presets) => {
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

const getTemplateForPath = (path) => {
  const normalized = normalizeRoute(path)
  return TEMPLATES[normalized] || TEMPLATES[ROUTES.clean]
}


const buildEmptyPair = (slotKeys, textDefault = '') => {
  const emptyPair = {
    ...EMPTY_PAIR,
    slideText: textDefault,
  }
  slotKeys.forEach((key) => {
    emptyPair[key] = ''
  })
  return emptyPair
}

const buildPresetPairs = (template, { slotKeys, textDefault }) => {
  if (template?.masterTitle !== TEMPLATES[ROUTES.desilting]?.masterTitle) {
    return null
  }
  const base = buildEmptyPair(slotKeys, textDefault)
  const presets = DESILTING_PRESET_TEXT.map((label) => ({
    ...base,
    slideText: label,
  }))
  return presets
}

const mergeDesiltingPresetPairs = (pairs, { slotKeys, textDefault }) => {
  const base = buildEmptyPair(slotKeys, textDefault)
  const existing = Array.isArray(pairs) ? pairs : []
  const next = DESILTING_PRESET_TEXT.map((label, index) => {
    const pair = existing[index] || base
    const hasText =
      typeof pair?.slideText === 'string' && pair.slideText.trim()
    return {
      ...base,
      ...pair,
      slideText: hasText ? pair.slideText : label,
    }
  })

  if (existing.length > DESILTING_PRESET_TEXT.length) {
    next.push(
      ...existing
        .slice(DESILTING_PRESET_TEXT.length)
        .filter(Boolean)
        .map((pair) => ({ ...base, ...pair })),
    )
  }

  return next
}

const resolvePairsSource = ({
  storedPairs,
  data,
  template,
  slotKeys,
  textDefault,
  requiresText,
}) => {
  const isDesilting =
    template?.masterTitle === TEMPLATES[ROUTES.desilting]?.masterTitle
  const hasAnyContent = (pairs) =>
    Array.isArray(pairs) &&
    pairs.some((pair) => {
      const hasImages = slotKeys.some((key) => Boolean(pair?.[key]))
      const hasText = requiresText
        ? typeof pair?.slideText === 'string' && pair.slideText.trim()
        : false
      return hasImages || hasText
    })

  let source = null
  if (storedPairs !== null) {
    if (!isDesilting || hasAnyContent(storedPairs)) {
      source = storedPairs
    }
  }
  if (!source && Array.isArray(data) && data.length > 0) {
    source = data
  }
  if (!source) {
    const presetPairs = buildPresetPairs(template, { slotKeys, textDefault })
    source = presetPairs || []
  }
  return isDesilting
    ? mergeDesiltingPresetPairs(source, { slotKeys, textDefault })
    : source
}

const isPairComplete = (pair, { slotKeys, requiresText }) => {
  if (!pair) {
    return false
  }
  const hasAllImages = slotKeys.every((key) => Boolean(pair?.[key]))
  if (!hasAllImages) {
    return false
  }
  if (!requiresText) {
    return true
  }
  const textValue = typeof pair?.slideText === 'string' ? pair.slideText.trim() : ''
  return Boolean(textValue)
}

const hasPairContent = (pair, { slotKeys, requiresText }) => {
  if (!pair) {
    return false
  }
  const hasImages = slotKeys.some((key) => Boolean(pair?.[key]))
  const hasText = requiresText
    ? typeof pair?.slideText === 'string' && pair.slideText.trim()
    : false
  return hasImages || hasText
}

const getMovableCount = (pairs = [], { slotKeys, requiresText }) => {
  if (!Array.isArray(pairs) || pairs.length === 0) {
    return 0
  }
  const lastIndex = pairs.length - 1
  const lastIsEmpty =
    lastIndex >= 0 &&
    !hasPairContent(pairs[lastIndex], { slotKeys, requiresText })
  return lastIsEmpty ? lastIndex : pairs.length
}

const normalizePairs = (
  items = [],
  { slotKeys = ['beforeImage', 'afterImage'], requiresText = false, textDefault = '' } = {},
) => {
  const emptyPair = buildEmptyPair(slotKeys, textDefault)
  const basePairs = items
    .filter(Boolean)
    .map((pair) => {
      const migratedPair = { ...pair }
      Object.keys(pair).forEach((key) => {
        if (key.startsWith('image_image_')) {
          const newKey = key.replace('image_image_', 'image_')
          migratedPair[newKey] = pair[key]
        }
        if (key.startsWith('text_text_')) {
          const newKey = key.replace('text_text_', 'text_')
          migratedPair[newKey] = pair[key]
        }
        if (key.startsWith('first_image_image_')) {
          const newKey = key.replace('first_image_image_', 'first_image_')
          migratedPair[newKey] = pair[key]
        }
        if (key.startsWith('first_text_text_')) {
          const newKey = key.replace('first_text_text_', 'first_text_')
          migratedPair[newKey] = pair[key]
        }
        if (key.startsWith('last_image_image_')) {
          const newKey = key.replace('last_image_image_', 'last_image_')
          migratedPair[newKey] = pair[key]
        }
        if (key.startsWith('last_text_text_')) {
          const newKey = key.replace('last_text_text_', 'last_text_')
          migratedPair[newKey] = pair[key]
        }
      })
      return {
        ...emptyPair,
        ...migratedPair,
        slideText:
          typeof pair?.slideText === 'string' ? pair.slideText : textDefault,
      }
    })

  if (basePairs.length === 0) {
    return [{ ...emptyPair }]
  }

  const lastPair = basePairs[basePairs.length - 1]
  if (isPairComplete(lastPair, { slotKeys, requiresText })) {
    basePairs.push({ ...emptyPair })
  }

  return basePairs
}

const readFileAsDataUrl = (file) =>
  new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result)
    reader.onerror = () => reject(reader.error)
    reader.readAsDataURL(file)
  })

const UNDO_DURATION_MS = 5000

function DropSlot({ label, value, onChange, className = '', urlMode = 'inline', style }) {
  const [isDragging, setIsDragging] = useState(false)
  const [urlInput, setUrlInput] = useState('')
  const [pendingDelete, setPendingDelete] = useState(null) // stores the image url pending deletion
  const [undoProgress, setUndoProgress] = useState(100) // 100 → 0 countdown
  const fileInputRef = useRef(null)
  const undoTimerRef = useRef(null)
  const undoRafRef = useRef(null)

  const clearUndoTimer = () => {
    if (undoTimerRef.current) {
      clearTimeout(undoTimerRef.current)
      undoTimerRef.current = null
    }
    if (undoRafRef.current) {
      cancelAnimationFrame(undoRafRef.current)
      undoRafRef.current = null
    }
  }

  const commitDelete = (imageUrl) => {
    clearUndoTimer()
    setPendingDelete(null)
    setUndoProgress(100)
    onChange('')
  }

  const handleDeleteClick = () => {
    clearUndoTimer()
    const imageToDelete = value
    setPendingDelete(imageToDelete)
    setUndoProgress(100)

    // Animate the countdown ring
    const startTime = performance.now()
    const tick = (now) => {
      const elapsed = now - startTime
      const remaining = Math.max(0, 100 - (elapsed / UNDO_DURATION_MS) * 100)
      setUndoProgress(remaining)
      if (remaining > 0) {
        undoRafRef.current = requestAnimationFrame(tick)
      }
    }
    undoRafRef.current = requestAnimationFrame(tick)

    undoTimerRef.current = setTimeout(() => {
      commitDelete(imageToDelete)
    }, UNDO_DURATION_MS)
  }

  const handleUndo = () => {
    clearUndoTimer()
    setPendingDelete(null)
    setUndoProgress(100)
    // value is already set (we never called onChange) so nothing more needed
  }

  const handleFile = async (file) => {
    if (!file) {
      return
    }
    clearUndoTimer()
    setPendingDelete(null)
    const dataUrl = await readFileAsDataUrl(file)
    onChange(dataUrl)
  }

  const handleDrop = (event) => {
    event.preventDefault()
    setIsDragging(false)
    const file = event.dataTransfer.files?.[0]
    if (file) {
      handleFile(file)
    }
  }

  const handleDragOver = (event) => {
    event.preventDefault()
  }

  const handleDragEnter = (event) => {
    event.preventDefault()
    setIsDragging(true)
  }

  const handleDragLeave = () => {
    setIsDragging(false)
  }

  const handleBrowse = (event) => {
    const file = event.target.files?.[0]
    if (file) {
      handleFile(file)
    }
    event.target.value = ''
  }

  const handleUrlAdd = (nextUrl) => {
    const trimmed = (nextUrl ?? urlInput).trim()
    if (!trimmed) {
      return
    }
    clearUndoTimer()
    setPendingDelete(null)
    onChange(trimmed)
    if (nextUrl === undefined) {
      setUrlInput('')
    }
  }

  const handleUrlPrompt = () => {
    const response = window.prompt('Paste image URL')
    if (response) {
      handleUrlAdd(response)
    }
  }

  const handleUrlKeyDown = (event) => {
    if (event.key === 'Enter') {
      event.preventDefault()
      handleUrlAdd()
    }
  }

  // Cleanup on unmount
  useEffect(() => {
    return () => clearUndoTimer()
  }, [])

  // Determine what image to show: if pending delete, show the pending image (greyed)
  const displayValue = pendingDelete !== null ? pendingDelete : value
  const isPendingDelete = pendingDelete !== null
  const slotClassName = `drop-slot${displayValue ? ' has-image' : ''}${
    isDragging ? ' is-dragging' : ''
  }${isPendingDelete ? ' is-pending-delete' : ''} ${className}`

  // SVG ring circumference for the countdown
  const RING_R = 14
  const RING_CIRC = 2 * Math.PI * RING_R
  const ringDash = (undoProgress / 100) * RING_CIRC

  return (
    <div
      className={slotClassName.trim()}
      onDrop={handleDrop}
      onDragOver={handleDragOver}
      onDragEnter={handleDragEnter}
      onDragLeave={handleDragLeave}
      style={style}
    >
      <div className="drop-slot__label">{label}</div>
      {displayValue ? (
        <div className="drop-slot__preview-wrap">
          <img
            className={`drop-slot__preview${isPendingDelete ? ' is-fading' : ''}`}
            src={displayValue}
            alt={`${label} preview`}
          />
          {isPendingDelete ? (
            <button
              type="button"
              className="drop-slot__undo"
              onClick={handleUndo}
              aria-label={`Undo remove ${label} image`}
            >
              <svg className="drop-slot__undo-ring" viewBox="0 0 36 36" aria-hidden="true">
                <circle
                  cx="18" cy="18" r={RING_R}
                  fill="none"
                  stroke="rgba(255,255,255,0.25)"
                  strokeWidth="3"
                />
                <circle
                  cx="18" cy="18" r={RING_R}
                  fill="none"
                  stroke="#fff"
                  strokeWidth="3"
                  strokeDasharray={`${ringDash} ${RING_CIRC}`}
                  strokeLinecap="round"
                  transform="rotate(-90 18 18)"
                />
              </svg>
              <span>↩ Undo</span>
            </button>
          ) : (
            <button
              type="button"
              className="drop-slot__delete"
              onClick={handleDeleteClick}
              aria-label={`Remove ${label} image`}
            >
              Delete
            </button>
          )}
        </div>
      ) : (
        <div className="drop-slot__placeholder">
          <p>Drag & drop an image</p>
          <span>or click Browse</span>
        </div>
      )}
      {!displayValue && (
        <>
          <div className="drop-slot__actions">
            <button
              type="button"
              className="ghost"
              onClick={() => fileInputRef.current?.click()}
            >
              Browse
            </button>
            {urlMode === 'prompt' && (
              <button type="button" className="ghost" onClick={handleUrlPrompt}>
                Paste URL
              </button>
            )}
          </div>
          {urlMode === 'inline' && (
            <div className="drop-slot__url">
              <input
                type="text"
                placeholder="Paste image URL"
                value={urlInput}
                onChange={(event) => setUrlInput(event.target.value)}
                onKeyDown={handleUrlKeyDown}
              />
              <button type="button" className="ghost" onClick={handleUrlAdd}>
                Add URL
              </button>
            </div>
          )}
        </>
      )}
      <input
        ref={fileInputRef}
        className="drop-slot__file"
        type="file"
        accept="image/*"
        onChange={handleBrowse}
      />
    </div>
  )
}

function SlideCanvas({
  pair,
  slots = DEFAULT_SLOTS,
  onChange,
  backgroundUrl,
  showText,
  textValue,
  textPlaceholder,
  textBoxes = [],
  onTextChange,
}) {
  const displayText = textValue?.trim() || textPlaceholder || ''
  return (
    <div
      className="slide-canvas"
      style={{ backgroundImage: `url(${backgroundUrl})` }}
    >
      {showText && displayText && (
        <div className="slide-text" aria-hidden="true">
          {displayText}
        </div>
      )}
      {slots.map((slot) => (
        <DropSlot
          key={slot.key}
          label={slot.label}
          value={pair?.[slot.key]}
          onChange={(value) => onChange(slot.key, value)}
          className={slot.className}
          style={slot.style}
          urlMode={slot.urlMode || 'inline'}
        />
      ))}
      {textBoxes.map((box) => {
        const textVal = pair?.[box.key] ?? ''
        return (
          <input
            key={box.key}
            type="text"
            className="custom-slide-textbox-input"
            value={textVal}
            onChange={(e) => onTextChange?.(box.key, e.target.value)}
            placeholder={box.textDefault || 'Enter text'}
            style={{
              position: 'absolute',
              left: `${(box.x / 13.333) * 100}%`,
              top: `${(box.y / 7.5) * 100}%`,
              width: `${(box.w / 13.333) * 100}%`,
              height: `${(box.h / 7.5) * 100}%`,
              fontSize: `${((box.fontSize || 20) * 0.104).toFixed(3)}cqw`,
              fontFamily: box.fontFace || 'Calibri',
              color: box.fontColor ? `#${box.fontColor}` : '#111111',
              fontWeight: box.bold ? 'bold' : 'normal',
              textAlign: box.align || 'center',
              background: 'rgba(255, 255, 255, 0.75)',
              border: '1px dashed rgba(11, 122, 56, 0.4)',
              borderRadius: '8px',
              padding: '4px 8px',
              outline: 'none',
              boxSizing: 'border-box',
            }}
          />
        )
      })}
    </div>
  )
}

function App({ data }) {
  const currentRoute = normalizeRoute(window.location.pathname)
  // ── Preset management state ─────────────────────────────────────────────
  const [masterPresets, setMasterPresets] = useState([])
  const [activePresetId, setActivePresetId] = useState(null)
  // Derive customLayout from active preset
  const customLayout = masterPresets.find((p) => p.id === activePresetId)?.layout || null

  const [designerMode, setDesignerMode] = useState('design')
  const [firstSlideData, setFirstSlideData] = useState({})
  const [lastSlideData, setLastSlideData] = useState({})
  const [isHydrated, setIsHydrated] = useState(false)

  useEffect(() => {
    const loadCustomData = async () => {
      try {
        // Load presets
        let presets = await loadMasterPresets()

        // Migrate old single layout key if needed
        if (presets.length === 0) {
          const migrated = await migrateLegacyLayout(presets)
          if (migrated) {
            presets = migrated
            await saveMasterPresets(presets)
          }
        } else {
          const migrated = await migrateLegacyLayout(presets)
          if (migrated) {
            presets = migrated
            await saveMasterPresets(presets)
          }
        }

        setMasterPresets(presets)

        // Load active preset id
        let activeId = await loadActivePresetId()
        if (!activeId && presets.length > 0) {
          activeId = presets[0].id
          await saveActivePresetId(activeId)
        }
        setActivePresetId(activeId)

        // Switch to Use mode if a preset has placeholders
        const active = presets.find((p) => p.id === activeId)
        if (active?.layout?.placeholders?.length) {
          setDesignerMode('use')
        }

        // Load first/last slide data (scoped to the active preset)
        if (activeId) {
          const first = await loadPairsFromDb(`pptxpro:custom-first-slide-data:${activeId}`)
          const last = await loadPairsFromDb(`pptxpro:custom-last-slide-data:${activeId}`)
          if (first) setFirstSlideData(first)
          if (last) setLastSlideData(last)
        }
      } catch (err) {
        console.error('Failed to load master preset data from IndexedDB', err)
      }
    }
    loadCustomData()
  }, [])

  // ── Preset action handlers ───────────────────────────────────────────────
  const handleSaveNewPreset = async (name, layout) => {
    const newPreset = {
      id: `preset_${Date.now()}`,
      name,
      createdAt: Date.now(),
      layout,
    }
    const updated = [...masterPresets, newPreset]
    setMasterPresets(updated)
    setActivePresetId(newPreset.id)
    await saveMasterPresets(updated)
    await saveActivePresetId(newPreset.id)
    setDesignerMode('design')
    alert(`Preset "${name}" saved! Switch to "Use Template" to use it.`)
  }

  const handleLoadPreset = async (id) => {
    setActivePresetId(id)
    await saveActivePresetId(id)
    // Load slide data that belongs to this specific preset (don't clear it —
    // each preset stores its own data under a scoped key)
    const first = await loadPairsFromDb(`pptxpro:custom-first-slide-data:${id}`)
    const last = await loadPairsFromDb(`pptxpro:custom-last-slide-data:${id}`)
    setFirstSlideData(first || {})
    setLastSlideData(last || {})
  }

  const handleDeletePreset = async (id) => {
    const updated = masterPresets.filter((p) => p.id !== id)
    setMasterPresets(updated)
    await saveMasterPresets(updated)
    if (activePresetId === id) {
      const newActive = updated[0]?.id || null
      setActivePresetId(newActive)
      await saveActivePresetId(newActive)
    }
  }

  // ── Export / Import handlers ────────────────────────────────────────────
  const handleExportPreset = (id) => {
    const preset = masterPresets.find((p) => p.id === id)
    if (!preset) return
    const blob = new Blob([JSON.stringify([preset], null, 2)], { type: 'application/json' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `${preset.name.replace(/[^a-z0-9_-]/gi, '_')}.pptxpro-template.json`
    a.click()
    URL.revokeObjectURL(url)
  }

  const handleExportAllPresets = () => {
    if (masterPresets.length === 0) {
      alert('No templates to export.')
      return
    }
    const blob = new Blob([JSON.stringify(masterPresets, null, 2)], { type: 'application/json' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = 'all-templates.pptxpro-template.json'
    a.click()
    URL.revokeObjectURL(url)
  }

  const handleImportPresetsFile = (file) => {
    if (!file) return
    const reader = new FileReader()
    reader.onload = async (e) => {
      try {
        const text = e.target.result
        let parsed
        try {
          parsed = JSON.parse(text)
        } catch {
          alert('❌ Invalid file — could not parse JSON. Make sure you are importing a valid .pptxpro-template.json file.')
          return
        }

        // Support both a single preset object and an array of presets
        const incoming = Array.isArray(parsed) ? parsed : (parsed && parsed.id ? [parsed] : [])

        if (incoming.length === 0) {
          alert('❌ No templates found in the file.')
          return
        }

        // Relaxed validation — only require that layout exists
        const valid = incoming.filter(
          (p) => p && typeof p === 'object' && p.layout
        )

        if (valid.length === 0) {
          alert('❌ The file does not contain valid template data. Each template must have a "layout" field.')
          return
        }

        // Ensure every imported preset has a valid id
        const sanitized = valid.map((p) => ({
          ...p,
          id: typeof p.id === 'string' && p.id ? p.id : `preset_${Date.now()}_${Math.random().toString(36).slice(2)}`,
          name: typeof p.name === 'string' && p.name ? p.name : 'Imported Template',
          createdAt: p.createdAt || Date.now(),
        }))

        // Merge: re-id any collisions with existing presets
        const existingIds = new Set(masterPresets.map((p) => p.id))
        const toAdd = sanitized.map((p) => {
          if (existingIds.has(p.id)) {
            return {
              ...p,
              id: `preset_imported_${Date.now()}_${Math.random().toString(36).slice(2)}`,
              name: `${p.name} (imported)`,
            }
          }
          return p
        })

        const updated = [...masterPresets, ...toAdd]
        setMasterPresets(updated)
        await saveMasterPresets(updated)

        // Auto-load the first imported preset if none is active
        if (!activePresetId && toAdd.length > 0) {
          setActivePresetId(toAdd[0].id)
          await saveActivePresetId(toAdd[0].id)
        }

        alert(`✅ Imported ${toAdd.length} template${toAdd.length !== 1 ? 's' : ''} successfully! ${toAdd.length === 1 ? `"${toAdd[0].name}"` : ''} is now in your library.`)
      } catch (err) {
        alert(`❌ Import failed: ${err.message}`)
      }
    }
    reader.onerror = () => {
      alert('❌ Could not read the file. Please try again.')
    }
    reader.readAsText(file)
  }

  const handleRenamePreset = async (id, newName) => {
    const updated = masterPresets.map((p) => p.id === id ? { ...p, name: newName } : p)
    setMasterPresets(updated)
    await saveMasterPresets(updated)
  }

  const template = useMemo(() => {
    if (currentRoute === ROUTES.extract) {
      return {
        eyebrow: 'PPTXPro',
        title: 'Image Extractor',
        subtext: 'Upload any PPTX to extract every image from every slide. Review, reorder, and export them as individual slides.',
        masterBgUrl: '',
        slots: [],
        themeLabel: '',
      }
    }
    if (currentRoute === ROUTES.master) {
      return {
        eyebrow: 'Custom Master',
        title: 'Custom Template Report',
        subtext: 'Upload images to automatically generate slides and compile PPTX reports.',
        masterBgUrl: customLayout?.masterBgUrl || '',
        firstSlideUrl: customLayout?.firstSlideUrl || '',
        secondSlideUrl: '',
        lastSlideUrl: customLayout?.lastSlideUrl || '',
        fileNamePrefix: masterPresets.find((p) => p.id === activePresetId)?.name || 'Custom_Report',
        masterTitle: 'CUSTOM_MASTER',
        slideTitle: 'Custom Report',
        themeLabel: 'Custom template slide background',
        importSkipFirst: 1,
        importSkipLast: 1,
        imageCount: customLayout?.placeholders?.length || 0,
        slots: (customLayout?.placeholders || []).map((p) => ({
          key: p.key,
          label: p.label,
          className: 'slide-slot',
          style: {
            left: `${(p.x / 13.333) * 100}%`,
            top: `${(p.y / 7.5) * 100}%`,
            width: `${(p.w / 13.333) * 100}%`,
            height: `${(p.h / 7.5) * 100}%`,
            position: 'absolute',
          }
        })),
        textBoxes: customLayout?.textboxes || []
      }
    }
    return getTemplateForPath(window.location.pathname)
  }, [currentRoute, customLayout, masterPresets, activePresetId])

  const [dailyVariant, setDailyVariant] = useState('urban')
  const slotKeys = useMemo(() => {
    return template.slots?.length
      ? template.slots.map((slot) => slot.key)
      : ['beforeImage', 'afterImage']
  }, [template])
  const requiresText = Boolean(template.textBox)
  const textDefault = template.textDefault || ''
  // On the master route each preset gets its own storage bucket so images
  // don't bleed between templates.
  const storageKey =
    currentRoute === ROUTES.master && activePresetId
      ? `${STORAGE_PREFIX}:${currentRoute}:${activePresetId}`
      : buildStorageKey(
          currentRoute,
          currentRoute === ROUTES.dailyPlot ? dailyVariant : '',
        )
  const [prevKeyInfo, setPrevKeyInfo] = useState({ slotKeys, storageKey })
  if (slotKeys !== prevKeyInfo.slotKeys || storageKey !== prevKeyInfo.storageKey) {
    setPrevKeyInfo({ slotKeys, storageKey })
    setIsHydrated(false)
  }
  const [pairs, setPairs] = useState(() => {
    const storedPairs = loadStoredPairsSync(storageKey)
    const source = resolvePairsSource({
      storedPairs,
      data,
      template,
      slotKeys,
      textDefault,
      requiresText,
    })
    return normalizePairs(source, { slotKeys, requiresText, textDefault })
  })
  const [isGenerating, setIsGenerating] = useState(false)
  const [isImporting, setIsImporting] = useState(false)
  const [importStatus, setImportStatus] = useState({ type: 'idle', message: '' })
  const [dragIndex, setDragIndex] = useState(null)
  const [dragOverIndex, setDragOverIndex] = useState(null)
  const [moveMenuIndex, setMoveMenuIndex] = useState(null)
  const pairRefs = useRef([])
  const pptxInputRef = useRef(null)
  const pendingScrollIndex = useRef(null)
  const moveMenuRef = useRef(null)
  const hydrationRef = useRef({ key: null, skipSave: false })

  useEffect(() => {
    const currentRoute = normalizeRoute(window.location.pathname)
    if (!TEMPLATES[currentRoute] && currentRoute !== ROUTES.master && currentRoute !== ROUTES.extract) {
      window.history.replaceState(null, '', ROUTES.clean)
    }
  }, [])

  useEffect(() => {
    if (currentRoute === ROUTES.dailyPlot) {
      setDailyVariant('urban')
    }
  }, [currentRoute])

  useEffect(() => {
    setIsHydrated(false)
    let cancelled = false
    const hydrate = async () => {
      const storedPairs = await loadStoredPairs(storageKey)
      const source = resolvePairsSource({
        storedPairs,
        data,
        template,
        slotKeys,
        textDefault,
        requiresText,
      })
      if (cancelled) {
        return
      }
      setPairs(normalizePairs(source, { slotKeys, requiresText, textDefault }))
      hydrationRef.current = { key: storageKey, skipSave: true }
      setIsHydrated(true)
    }
    hydrate()
    return () => {
      cancelled = true
    }
  }, [storageKey, data, slotKeys, requiresText, textDefault])

  useEffect(() => {
    if (!isHydrated) {
      return
    }
    if (
      hydrationRef.current.key === storageKey &&
      hydrationRef.current.skipSave
    ) {
      hydrationRef.current.skipSave = false
      return
    }
    savePairsToStorage(storageKey, pairs, { slotKeys, requiresText })
  }, [pairs, storageKey, slotKeys, requiresText, isHydrated])

  useEffect(() => {
    const index = pendingScrollIndex.current
    if (index === null || index === undefined) {
      return
    }

    const target = pairRefs.current[index]
    if (target) {
      target.scrollIntoView({ behavior: 'smooth', block: 'center' })
    }
    pendingScrollIndex.current = null
  }, [pairs])

  useEffect(() => {
    if (moveMenuIndex === null) {
      return
    }
    const handleClick = (event) => {
      if (!moveMenuRef.current) {
        setMoveMenuIndex(null)
        return
      }
      if (!moveMenuRef.current.contains(event.target)) {
        setMoveMenuIndex(null)
      }
    }
    window.addEventListener('click', handleClick)
    return () => window.removeEventListener('click', handleClick)
  }, [moveMenuIndex])

  const updatePair = (index, key, value) => {
    setPairs((prev) => {
      const next = prev.map((pair, pairIndex) =>
        pairIndex === index ? { ...pair, [key]: value } : pair,
      )

      const isComplete = isPairComplete(next[index], { slotKeys, requiresText })
      if (isComplete) {
        if (index === next.length - 1) {
          next.push(buildEmptyPair(slotKeys, textDefault))
          pendingScrollIndex.current = next.length - 1
        } else {
          pendingScrollIndex.current = index + 1
        }
      }

      return next
    })
  }

  const movePair = (fromIndex, toIndex) => {
    setPairs((prev) => {
      const next = [...prev]
      const lastIndex = next.length - 1
      const lastIsEmpty =
        lastIndex >= 0 &&
        !hasPairContent(next[lastIndex], { slotKeys, requiresText })
      const resolvedToIndex =
        lastIsEmpty && toIndex >= lastIndex ? lastIndex - 1 : toIndex
      if (
        fromIndex === resolvedToIndex ||
        resolvedToIndex < 0 ||
        resolvedToIndex >= next.length
      ) {
        return prev
      }
      const [moved] = next.splice(fromIndex, 1)
      next.splice(resolvedToIndex, 0, moved)
      return next
    })
  }

  const swapPairImages = (index) => {
    setPairs((prev) => {
      const next = prev.map((pair, pairIndex) => {
        if (pairIndex !== index) {
          return pair
        }

        return {
          ...pair,
          beforeImage: pair?.afterImage || '',
          afterImage: pair?.beforeImage || '',
        }
      })

      return next
    })
  }

  const handlePairDragStart = (event, index) => {
    setDragIndex(index)
    setMoveMenuIndex(null)
    event.dataTransfer.effectAllowed = 'move'
    event.dataTransfer.setData('text/plain', String(index))
  }

  const handlePairDragEnd = () => {
    setDragIndex(null)
    setDragOverIndex(null)
  }

  const handlePairDragOver = (event, index) => {
    if (dragIndex === null || dragIndex === index) {
      return
    }
    event.preventDefault()
    event.dataTransfer.dropEffect = 'move'
    setDragOverIndex(index)
  }

  const handlePairDrop = (event, index) => {
    if (dragIndex === null) {
      return
    }
    event.preventDefault()
    movePair(dragIndex, index)
    setDragIndex(null)
    setDragOverIndex(null)
  }

  const completePairs = pairs.filter((pair) =>
    isPairComplete(pair, { slotKeys, requiresText }),
  )
  const slideCount = completePairs.length
  const incompletePairCount = pairs.filter((pair) => {
    if (!hasPairContent(pair, { slotKeys, requiresText })) return false
    return !isPairComplete(pair, { slotKeys, requiresText })
  }).length
  const movableCount = getMovableCount(pairs, { slotKeys, requiresText })
  const canDownload = (slideCount > 0 || incompletePairCount > 0) && !isGenerating && !isImporting

  const handlePptxUpload = async (event) => {
    const file = event.target.files?.[0]
    event.target.value = ''
    if (!file) {
      return
    }

    const hasExisting = pairs.some((pair) => {
      const hasImages = slotKeys.some((key) => pair?.[key])
      const hasText = requiresText
        ? typeof pair?.slideText === 'string' && pair.slideText.trim()
        : false
      return hasImages || hasText
    })
    if (hasExisting) {
      const confirmed = window.confirm(
        'Importing will replace the current slides. Continue?',
      )
      if (!confirmed) {
        return
      }
    }

    try {
      setIsImporting(true)
      setImportStatus({ type: 'working', message: 'Importing PPTX...' })
      const skipFirstSlides = template.importSkipFirst ?? 1
      const skipLastSlides = template.importSkipLast ?? 1
      const { pairs: importedPairs, importedSlides, emptySlides } =
        await importPptxSlides(file, {
          skipFirstSlides,
          skipLastSlides,
          imageCount: template.imageCount || slotKeys.length,
          // On the master route, pass textbox definitions so text is auto-extracted
          textboxDefs: currentRoute === ROUTES.master ? (customLayout?.textboxes || []) : [],
        })

      setPairs(
        normalizePairs(importedPairs, { slotKeys, requiresText, textDefault }),
      )
      const emptyNote = emptySlides
        ? ` ${emptySlides} slide(s) need images.`
        : ''
      const removedParts = []
      if (skipFirstSlides) {
        removedParts.push(`${skipFirstSlides} intro`)
      }
      if (skipLastSlides) {
        removedParts.push(`${skipLastSlides} ending`)
      }
      const removedNote = removedParts.length
        ? ` (${removedParts.join(' + ')} removed).`
        : '.'
      setImportStatus({
        type: 'success',
        message: `Imported ${importedSlides} slides${removedNote}${emptyNote}`,
      })
    } catch (error) {
      setImportStatus({
        type: 'error',
        message: error?.message || 'PPTX import failed.',
      })
    } finally {
      setIsImporting(false)
    }
  }

  const handlePptxButtonClick = () => {
    pptxInputRef.current?.click()
  }

  const handleClearStored = async () => {
    const confirmed = window.confirm(
      'Clear stored slides for this report? This cannot be undone.',
    )
    if (!confirmed) {
      return
    }
    try {
      if (canUseStorage()) {
        window.localStorage.removeItem(storageKey)
      }
      await removePairsFromDb(storageKey)
    } finally {
      setPairs(normalizePairs([], { slotKeys, requiresText, textDefault }))
      setImportStatus({ type: 'idle', message: '' })
      setDragIndex(null)
      setDragOverIndex(null)
    }
  }

  const handleDownload = async () => {
    if (!canDownload) {
      return
    }

    try {
      setIsGenerating(true)
      const resolvedFirstSlideUrl =
        currentRoute === ROUTES.dailyPlot && dailyVariant === 'rural'
          ? template.firstSlideUrlRural
          : template.firstSlideUrl
      const resolvedFileName =
        typeof template.fileName === 'function'
          ? template.fileName(dailyVariant)
          : template.fileName
      const backgroundImage = template.masterBgUrl
        ? new URL(template.masterBgUrl, window.location.href).toString()
        : undefined
      await generateReport(pairs, {
        fileName: resolvedFileName,
        backgroundImage,
        firstSlideImage: resolvedFirstSlideUrl || undefined,
        secondSlideImage: template.secondSlideUrl || undefined,
        lastSlideImage: template.lastSlideUrl || undefined,
        fileNamePrefix: template.fileNamePrefix,
        slideTitle: template.slideTitle,
        masterTitle: template.masterTitle,
        leftBox: template.leftBox,
        middleBox: template.middleBox,
        rightBox: template.rightBox,
        dateSlide: template.dateSlide,
        dateBox: template.dateBox,
        dateColor: template.dateColor,
        dateAlign: template.dateAlign,
        dateFontSize: template.dateFontSize,
        dateBold: template.dateBold,
        dateFontFace: template.dateFontFace,
        dateOffsetDays: template.dateOffsetDays,
        textBox: template.textBox,
        textColor: template.textColor,
        textAlign: template.textAlign,
        textFontSize: template.textFontSize,
        textBold: template.textBold,
        textDefault: template.textDefault,
        placeholders: currentRoute === ROUTES.master ? customLayout?.placeholders : undefined,
        textboxes: currentRoute === ROUTES.master ? customLayout?.textboxes : undefined,
        firstSlidePlaceholders: currentRoute === ROUTES.master ? customLayout?.firstSlidePlaceholders : undefined,
        firstSlideTextboxes: currentRoute === ROUTES.master ? customLayout?.firstSlideTextboxes : undefined,
        firstSlideData: currentRoute === ROUTES.master ? firstSlideData : undefined,
        lastSlidePlaceholders: currentRoute === ROUTES.master ? customLayout?.lastSlidePlaceholders : undefined,
        lastSlideTextboxes: currentRoute === ROUTES.master ? customLayout?.lastSlideTextboxes : undefined,
        lastSlideData: currentRoute === ROUTES.master ? lastSlideData : undefined,
      })
    } finally {
      setIsGenerating(false)
    }
  }

  return (
    <main
      className={`app${currentRoute === ROUTES.compliance ? ' app--compliance' : ''}${
        currentRoute === ROUTES.desilting ? ' app--desilting' : ''
      }${currentRoute === ROUTES.dailyPlot ? ' app--daily-plot' : ''}`}
    >
      <header className="app__header">
        <div>
          <p className="app__eyebrow">{template.eyebrow}</p>
          <h1>{template.title}</h1>
          <p className="app__subtext">{template.subtext}</p>
        </div>
        <div className="app__actions">
          <div className="app__nav">
            <button
              type="button"
              className={`ghost${currentRoute === ROUTES.clean ? ' is-active' : ''}`}
              onClick={() => {
                if (currentRoute !== ROUTES.clean) {
                  window.location.pathname = ROUTES.clean
                }
              }}
            >
              Clean Punjab
            </button>
            <button
              type="button"
              className={`ghost${currentRoute === ROUTES.compliance ? ' is-active' : ''}`}
              onClick={() => {
                if (currentRoute !== ROUTES.compliance) {
                  window.location.pathname = ROUTES.compliance
                }
              }}
            >
              Compliance
            </button>
            <button
              type="button"
              className={`ghost${currentRoute === ROUTES.desilting ? ' is-active' : ''}`}
              onClick={() => {
                if (currentRoute !== ROUTES.desilting) {
                  window.location.pathname = ROUTES.desilting
                }
              }}
            >
              Desilting
            </button>
            <button
              type="button"
              className={`ghost${currentRoute === ROUTES.dailyPlot ? ' is-active' : ''}`}
              onClick={() => {
                if (currentRoute !== ROUTES.dailyPlot) {
                  window.location.pathname = ROUTES.dailyPlot
                }
              }}
            >
              Daily Plot
            </button>
            <button
              type="button"
              className={`ghost${currentRoute === ROUTES.master ? ' is-active' : ''}`}
              onClick={() => {
                if (currentRoute !== ROUTES.master) {
                  window.location.pathname = ROUTES.master
                }
              }}
            >
              Master Creator
            </button>
            <button
              type="button"
              className={`ghost${currentRoute === ROUTES.extract ? ' is-active' : ''}`}
              onClick={() => {
                if (currentRoute !== ROUTES.extract) {
                  window.location.pathname = ROUTES.extract
                }
              }}
            >
              Extract Images
            </button>
          </div>
          {currentRoute === ROUTES.dailyPlot && (
            <div className="app__subnav">
              <button
                type="button"
                className={`ghost${dailyVariant === 'urban' ? ' is-active' : ''}`}
                onClick={() => setDailyVariant('urban')}
              >
                Urban
              </button>
              <button
                type="button"
                className={`ghost${dailyVariant === 'rural' ? ' is-active' : ''}`}
                onClick={() => setDailyVariant('rural')}
              >
                Rural
              </button>
            </div>
          )}
          {currentRoute === ROUTES.master && (
            <div className="app__subnav">
              <button
                type="button"
                className={`ghost${designerMode === 'design' ? ' is-active' : ''}`}
                onClick={() => setDesignerMode('design')}
              >
                Design Layout
              </button>
              <button
                type="button"
                className={`ghost${designerMode === 'use' ? ' is-active' : ''}`}
                onClick={() => {
                  if (customLayout?.placeholders?.length) {
                    setDesignerMode('use')
                  } else {
                    alert('Please add at least one Image Placeholder to your master slide layout before switching!')
                  }
                }}
              >
                Use Template
              </button>
            </div>
          )}
          {!(currentRoute === ROUTES.master && designerMode === 'design') && currentRoute !== ROUTES.extract && (
            <>
              <div className="app__badge">Slides ready: {slideCount}</div>
              <button
                type="button"
                className="button button--secondary"
                onClick={handlePptxButtonClick}
                disabled={isImporting}
              >
                {isImporting ? 'Importing...' : 'Upload PPTX'}
              </button>
              <button type="button" className="ghost" onClick={handleClearStored}>
                Clear Saved
              </button>
              <input
                ref={pptxInputRef}
                className="file-input"
                type="file"
                accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
                onChange={handlePptxUpload}
              />
              <button
                type="button"
                className="button"
                onClick={handleDownload}
                disabled={!canDownload}
              >
                {isGenerating ? 'Building PPTX...' : 'Download Report'}
              </button>
            </>
          )}
        </div>
      </header>

      {importStatus.message && (
        <p className={`app__note app__note--${importStatus.type}`}>
          {importStatus.message}
        </p>
      )}

      {currentRoute === ROUTES.extract ? (
        <ImageExtractor />
      ) : currentRoute === ROUTES.master && designerMode === 'design' ? (
        <MasterDesigner
          customLayout={customLayout}
          onSave={async (layout) => {
            // Overwrite active preset's layout
            if (activePresetId) {
              const updated = masterPresets.map((p) =>
                p.id === activePresetId ? { ...p, layout } : p
              )
              setMasterPresets(updated)
              await saveMasterPresets(updated)
              alert('Active preset updated! Switch to "Use Template" to use it.')
              setDesignerMode('use')
            } else {
              // No active preset — save as new unnamed preset
              await handleSaveNewPreset('Default', layout)
            }
          }}
          presets={masterPresets}
          activePresetId={activePresetId}
          onSavePreset={handleSaveNewPreset}
          onLoadPreset={handleLoadPreset}
          onDeletePreset={handleDeletePreset}
          onRenamePreset={handleRenamePreset}
          onExportPreset={handleExportPreset}
          onExportAll={handleExportAllPresets}
          onImportPresets={handleImportPresetsFile}
        />
      ) : (
        <>
          <section className="card card--theme">
            <div>
              <p className="label">{template.themeLabel}</p>
              <p className="value">
                Using background image: <span>{template.masterBgUrl ? (template.masterBgUrl.startsWith('data:') ? 'Custom background data' : template.masterBgUrl) : 'None'}</span>
              </p>
            </div>
            {currentRoute !== ROUTES.master && (
              <p className="value">
                Export your master slide from the PPTX as PNG and place it in the
                public folder with this name. The image should include the title and
                footer styling.
              </p>
            )}
          </section>

      <section className="pairs">
        {currentRoute === ROUTES.master && ((customLayout?.firstSlidePlaceholders?.length > 0) || (customLayout?.firstSlideTextboxes?.length > 0)) && (
          <article className="pair" style={{ border: '2px solid rgba(11, 122, 56, 0.4)' }}>
            <div className="pair__header">
              <p className="label" style={{ fontWeight: 'bold', color: '#0b7a38' }}>Title Slide (First Slide)</p>
              <span className="pair__status is-complete">Cover Slide</span>
            </div>
            <SlideCanvas
              pair={firstSlideData}
              slots={(customLayout?.firstSlidePlaceholders || []).map((p) => ({
                key: p.key,
                label: p.label,
                className: 'slide-slot',
                style: {
                  left: `${(p.x / 13.333) * 100}%`,
                  top: `${(p.y / 7.5) * 100}%`,
                  width: `${(p.w / 13.333) * 100}%`,
                  height: `${(p.h / 7.5) * 100}%`,
                  position: 'absolute',
                }
              }))}
              onChange={(key, value) => {
                setFirstSlideData((prev) => {
                  const next = { ...prev, [key]: value }
                  if (activePresetId) savePairsToDb(`pptxpro:custom-first-slide-data:${activePresetId}`, next)
                  return next
                })
              }}
              backgroundUrl={customLayout?.firstSlideUrl}
              textBoxes={customLayout?.firstSlideTextboxes || []}
              onTextChange={(key, value) => {
                setFirstSlideData((prev) => {
                  const next = { ...prev, [key]: value }
                  if (activePresetId) savePairsToDb(`pptxpro:custom-first-slide-data:${activePresetId}`, next)
                  return next
                })
              }}
            />
          </article>
        )}
        {pairs.map((pair, index) => {
          const canDrag = hasPairContent(pair, { slotKeys, requiresText })
          const isDragging = dragIndex === index
          const isDragOver = dragOverIndex === index
          const isMenuOpen = moveMenuIndex === index
          const canMove = canDrag && movableCount > 1
          return (
            <article
              key={`pair-${index}`}
              className={`pair${isDragging ? ' is-dragging' : ''}${
                isDragOver ? ' is-drag-over' : ''
              }`}
              ref={(node) => {
                pairRefs.current[index] = node
              }}
              onDragOver={(event) => handlePairDragOver(event, index)}
              onDrop={(event) => handlePairDrop(event, index)}
            >
              <div className="pair__header">
                <p className="label">Pair {index + 1}</p>
                <div className="pair__controls">
                  {(() => {
                    const isComplete = isPairComplete(pair, {
                      slotKeys,
                      requiresText,
                    })
                    const statusLabel = isComplete
                      ? 'Complete'
                      : slotKeys.some((key) => pair?.[key])
                        ? 'In progress'
                        : 'Empty'
                    return (
                      <span
                        className={`pair__status${isComplete ? ' is-complete' : ''}`}
                      >
                        {statusLabel}
                      </span>
                    )
                  })()}
                  <button
                    type="button"
                    className="ghost"
                    onClick={() => swapPairImages(index)}
                    disabled={!slotKeys.some((key) => pair?.[key])}
                    title="Swap before and after images"
                  >
                    Swap
                  </button>
                  <div
                    className="pair__move"
                    ref={isMenuOpen ? moveMenuRef : null}
                  >
                    <button
                      type="button"
                      className="pair__drag"
                      draggable={canDrag}
                      disabled={!canMove}
                      aria-label="Drag to reorder or click to move"
                      aria-expanded={isMenuOpen}
                      aria-haspopup="menu"
                      title={
                        canMove
                          ? 'Drag to reorder or click to move'
                          : 'Add another slide to enable moving'
                      }
                      onDragStart={(event) => handlePairDragStart(event, index)}
                      onDragEnd={handlePairDragEnd}
                      onClick={(event) => {
                        if (!canMove) {
                          return
                        }
                        event.stopPropagation()
                        setMoveMenuIndex((prev) =>
                          prev === index ? null : index,
                        )
                      }}
                    >
                      Drag
                    </button>
                    {isMenuOpen && (
                      <div className="pair__move-menu" role="menu">
                        <p className="pair__move-title">Move to position</p>
                        <div className="pair__move-list">
                          {Array.from({ length: movableCount }, (_, target) => {
                            const position = target + 1
                            const isCurrent = target === index
                            return (
                              <button
                                type="button"
                                key={`move-${index}-to-${position}`}
                                className={`pair__move-option${
                                  isCurrent ? ' is-current' : ''
                                }`}
                                disabled={isCurrent}
                                onClick={() => {
                                  movePair(index, target)
                                  setMoveMenuIndex(null)
                                }}
                              >
                                {position}
                              </button>
                            )
                          })}
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              </div>
              {template.textBox && (
                <div className="pair__text">
                  <label className="label" htmlFor={`pair-text-${index}`}>
                    {template.textLabel || 'Slide text'}
                  </label>
                  <input
                    id={`pair-text-${index}`}
                    type="text"
                    value={pair.slideText || ''}
                    onChange={(event) =>
                      updatePair(index, 'slideText', event.target.value)
                    }
                    placeholder={template.textDefault || 'Enter text'}
                  />
                </div>
              )}
              <SlideCanvas
                pair={pair}
                slots={template.slots}
                onChange={(key, value) => updatePair(index, key, value)}
                backgroundUrl={template.masterBgUrl}
                showText={Boolean(template.textBox)}
                textValue={pair.slideText || ''}
                textPlaceholder={template.textDefault || ''}
                textBoxes={template.textBoxes || []}
                onTextChange={(key, value) => updatePair(index, key, value)}
              />
            </article>
          )
        })}
        {currentRoute === ROUTES.master && ((customLayout?.lastSlidePlaceholders?.length > 0) || (customLayout?.lastSlideTextboxes?.length > 0)) && (
          <article className="pair" style={{ border: '2px solid rgba(11, 122, 56, 0.4)' }}>
            <div className="pair__header">
              <p className="label" style={{ fontWeight: 'bold', color: '#0b7a38' }}>Closing Slide (Last Slide)</p>
              <span className="pair__status is-complete">End Slide</span>
            </div>
            <SlideCanvas
              pair={lastSlideData}
              slots={(customLayout?.lastSlidePlaceholders || []).map((p) => ({
                key: p.key,
                label: p.label,
                className: 'slide-slot',
                style: {
                  left: `${(p.x / 13.333) * 100}%`,
                  top: `${(p.y / 7.5) * 100}%`,
                  width: `${(p.w / 13.333) * 100}%`,
                  height: `${(p.h / 7.5) * 100}%`,
                  position: 'absolute',
                }
              }))}
              onChange={(key, value) => {
                setLastSlideData((prev) => {
                  const next = { ...prev, [key]: value }
                  if (activePresetId) savePairsToDb(`pptxpro:custom-last-slide-data:${activePresetId}`, next)
                  return next
                })
              }}
              backgroundUrl={customLayout?.lastSlideUrl}
              textBoxes={customLayout?.lastSlideTextboxes || []}
              onTextChange={(key, value) => {
                setLastSlideData((prev) => {
                  const next = { ...prev, [key]: value }
                  if (activePresetId) savePairsToDb(`pptxpro:custom-last-slide-data:${activePresetId}`, next)
                  return next
                })
              }}
            />
          </article>
        )}
      </section>
      </>
      )}
    </main>
  )
}

export default App
