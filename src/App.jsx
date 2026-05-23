import { useEffect, useMemo, useRef, useState } from 'react'
import './App.css'
import { generateReport } from './report/generateReport'
import { importPptxSlides } from './report/importPptx'

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
}

const STORAGE_PREFIX = 'pptxpro:slides:v1'
const DB_NAME = 'pptxpro-slides'
const DB_STORE = 'pairs'
const DB_VERSION = 1

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
  const fromLocal = loadStoredPairsSync(storageKey)
  if (fromLocal !== null) {
    return fromLocal
  }
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
      if (canUseStorage()) {
        window.localStorage.removeItem(storageKey)
      }
      void removePairsFromDb(storageKey)
      return
    }
    let savedToLocal = false
    if (canUseStorage()) {
      try {
        window.localStorage.setItem(storageKey, JSON.stringify(trimmed))
        savedToLocal = true
      } catch {
        savedToLocal = false
      }
    }
    if (!savedToLocal || canUseIndexedDb()) {
      void savePairsToDb(storageKey, trimmed)
    }
  } catch {
    // Ignore storage errors (e.g. quota exceeded).
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
  'UC-55 Rana',
  'UC-56 Saroye',
  'UC-57 Lalupur',
  'UC-58 Mari Thakran',
  'UC-59 Bakapur',
  'UC-60 Kotli Mutwalian',
  'UC-61 Rajewala',
  'UC-62 Lalupur',
  'UC-63 Wahndo',
  'UC-64 Tamboli',
  'UC-65 Tolakey',
  'UC-66 Dargapur',
  'UC-67 Sadhoki',
  'UC-68 Harpoki',
  'UC-69 Ghanoki',
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

const SLIDE_HEIGHT = 7.5
const DESILTING_DATE_OFFSET_IN = 84 / 96
const DESILTING_DATE_BOX = {
  x: DESILTING_DATE_OFFSET_IN,
  y: SLIDE_HEIGHT - DESILTING_DATE_OFFSET_IN - 0.4,
  w: 2.4,
  h: 0.4,
}

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
    textBox: { x: 4.0 - 8 / 96, y: 0.88, w: 5.33, h: 0.5 },
    textColor: '111111',
    textAlign: 'center',
    textFontSize: 22,
    textBold: true,
    textDefault: '',
    textLabel: 'Sector text',
    dateSlide: 'first',
    dateBox: DESILTING_DATE_BOX,
    dateColor: '000000',
    dateAlign: 'left',
    dateFontSize: 28,
    dateBold: true,
    dateFontFace: 'Times New Roman',
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
    .map((pair) => ({
      ...emptyPair,
      ...pair,
      slideText:
        typeof pair?.slideText === 'string' ? pair.slideText : textDefault,
    }))

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

function DropSlot({ label, value, onChange, className = '', urlMode = 'inline' }) {
  const [isDragging, setIsDragging] = useState(false)
  const [urlInput, setUrlInput] = useState('')
  const fileInputRef = useRef(null)

  const handleFile = async (file) => {
    if (!file) {
      return
    }
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

  const slotClassName = `drop-slot${value ? ' has-image' : ''}${
    isDragging ? ' is-dragging' : ''
  } ${className}`

  return (
    <div
      className={slotClassName.trim()}
      onDrop={handleDrop}
      onDragOver={handleDragOver}
      onDragEnter={handleDragEnter}
      onDragLeave={handleDragLeave}
    >
      <div className="drop-slot__label">{label}</div>
      {value ? (
        <div className="drop-slot__preview-wrap">
          <img
            className="drop-slot__preview"
            src={value}
            alt={`${label} preview`}
          />
          <button
            type="button"
            className="drop-slot__delete"
            onClick={() => onChange('')}
            aria-label={`Remove ${label} image`}
          >
            Delete
          </button>
        </div>
      ) : (
        <div className="drop-slot__placeholder">
          <p>Drag & drop an image</p>
          <span>or click Browse</span>
        </div>
      )}
      <div className="drop-slot__actions">
        <button
          type="button"
          className="ghost"
          onClick={() => fileInputRef.current?.click()}
        >
          Browse
        </button>
        {value && (
          <button type="button" className="ghost" onClick={() => onChange('')}>
            Clear
          </button>
        )}
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
          urlMode={slot.urlMode || 'inline'}
        />
      ))}
    </div>
  )
}

function App({ data }) {
  const template = getTemplateForPath(window.location.pathname)
  const currentRoute = normalizeRoute(window.location.pathname)
  const [dailyVariant, setDailyVariant] = useState('urban')
  const slotKeys = useMemo(() => {
    return template.slots?.length
      ? template.slots.map((slot) => slot.key)
      : ['beforeImage', 'afterImage']
  }, [template])
  const requiresText = Boolean(template.textBox)
  const textDefault = template.textDefault || ''
  const storageKey = buildStorageKey(
    currentRoute,
    currentRoute === ROUTES.dailyPlot ? dailyVariant : '',
  )
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
    if (!TEMPLATES[currentRoute]) {
      window.history.replaceState(null, '', ROUTES.clean)
    }
  }, [])

  useEffect(() => {
    if (currentRoute === ROUTES.dailyPlot) {
      setDailyVariant('urban')
    }
  }, [currentRoute])

  useEffect(() => {
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
    }
    hydrate()
    return () => {
      cancelled = true
    }
  }, [storageKey, data, slotKeys, requiresText, textDefault])

  useEffect(() => {
    if (
      hydrationRef.current.key === storageKey &&
      hydrationRef.current.skipSave
    ) {
      hydrationRef.current.skipSave = false
      return
    }
    savePairsToStorage(storageKey, pairs, { slotKeys, requiresText })
  }, [pairs, storageKey, slotKeys, requiresText])

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

  const isDesiltingTemplate =
    template?.masterTitle === TEMPLATES[ROUTES.desilting]?.masterTitle
  const applyDesiltingLabels = (nextPairs) =>
    isDesiltingTemplate
      ? mergeDesiltingPresetPairs(nextPairs, { slotKeys, textDefault })
      : nextPairs

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

      return applyDesiltingLabels(next)
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
  const movableCount = getMovableCount(pairs, { slotKeys, requiresText })
  const canDownload = slideCount > 0 && !isGenerating && !isImporting

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
        })
      const preparedPairs = applyDesiltingLabels(importedPairs)
      setPairs(
        normalizePairs(preparedPairs, { slotKeys, requiresText, textDefault }),
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
      const clearedSource = resolvePairsSource({
        storedPairs: null,
        data: [],
        template,
        slotKeys,
        textDefault,
        requiresText,
      })
      setPairs(
        normalizePairs(clearedSource, { slotKeys, requiresText, textDefault }),
      )
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
      await generateReport(completePairs, {
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
        </div>
      </header>

      {importStatus.message && (
        <p className={`app__note app__note--${importStatus.type}`}>
          {importStatus.message}
        </p>
      )}

      <section className="card card--theme">
        <div>
          <p className="label">{template.themeLabel}</p>
          <p className="value">
            Using background image: <span>{template.masterBgUrl}</span>
          </p>
        </div>
        <p className="value">
          Export your master slide from the PPTX as PNG and place it in the
          public folder with this name. The image should include the title and
          footer styling.
        </p>
      </section>

      <section className="pairs">
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
              />
            </article>
          )
        })}
      </section>
    </main>
  )
}

export default App
