import { Suspense, lazy, useEffect, useMemo, useRef, useState } from 'react'
import './App.css'
import Navbar from './Navbar'
import { ImportedSlideEditor } from './ImportedSlideEditor'
import { SlideCanvas } from './components/SlideCanvas'
import { PairCard } from './components/PairCard'
import { ROUTES, normalizeRoute } from './config/routes'
import { TEMPLATES, getTemplateForPath } from './config/templates'
import {
  buildStorageKey,
  canUseStorage,
  loadPairsFromDb,
  savePairsToDb,
  removePairsFromDb,
  loadStoredPairsSync,
  loadStoredPairs,
  savePairsToStorage,
  loadMasterPresets,
  saveMasterPresets,
  loadActivePresetId,
  saveActivePresetId,
  migrateLegacyLayout,
} from './utils/storage'
import {
  buildEmptyPair,
  resolvePairsSource,
  isPairComplete,
  getMovableCount,
  normalizePairs,
  hasPairContent,
  trimTrailingEmptyPairs,
} from './utils/pairUtils'

const MasterDesigner = lazy(() =>
  import('./MasterDesigner').then((module) => ({
    default: module.MasterDesigner,
  })),
)
const ImageExtractor = lazy(() =>
  import('./ImageExtractor').then((module) => ({
    default: module.ImageExtractor,
  })),
)
const PptxMerger = lazy(() =>
  import('./PptxMerger').then((module) => ({
    default: module.PptxMerger,
  })),
)
const PdfToPptx = lazy(() =>
  import('./PdfToPptx').then((module) => ({
    default: module.PdfToPptx,
  })),
)
const PdfMerger = lazy(() =>
  import('./PdfMerger').then((module) => ({
    default: module.PdfMerger,
  })),
)
const CollageMaker = lazy(() =>
  import('./CollageMaker').then((module) => ({
    default: module.CollageMaker,
  })),
)
const PptxToPdfOcr = lazy(() =>
  import('./PptxToPdfOcr').then((module) => ({
    default: module.PptxToPdfOcr,
  })),
)
const PptxToPdf = lazy(() =>
  import('./PptxToPdf').then((module) => ({
    default: module.PptxToPdf,
  })),
)
const PptxEditor = lazy(() =>
  import('./PptxEditor').then((module) => ({
    default: module.PptxEditor,
  })),
)

const loadGenerateReportModule = () => import('./report/generateReport')
const loadImportPptxModule = () => import('./report/importPptx')
const loadJsZip = () => import('jszip').then((module) => module.default)

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

  // ── Imported first/last slide state (master route only) ─────────────────
  const [importedFirstSlide, setImportedFirstSlide] = useState(null)
  const [importedLastSlide, setImportedLastSlide] = useState(null)
  const [useTemplateFirst, setUseTemplateFirst] = useState(false)
  const [useTemplateLast, setUseTemplateLast] = useState(false)
  // Refs hold the latest edited data so canvas render can access it without stale closures
  const importedFirstSlideRef = useRef(null)
  const importedLastSlideRef = useRef(null)

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

        const incoming = Array.isArray(parsed) ? parsed : (parsed && parsed.id ? [parsed] : [])
        if (incoming.length === 0) {
          alert('❌ No templates found in the file.')
          return
        }

        const valid = incoming.filter((p) => p && typeof p === 'object' && p.layout)
        if (valid.length === 0) {
          alert('❌ The file does not contain valid template data. Each template must have a "layout" field.')
          return
        }

        const sanitized = valid.map((p) => ({
          ...p,
          id: typeof p.id === 'string' && p.id ? p.id : `preset_${Date.now()}_${Math.random().toString(36).slice(2)}`,
          name: typeof p.name === 'string' && p.name ? p.name : 'Imported Template',
          createdAt: p.createdAt || Date.now(),
        }))

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
    const updated = masterPresets.map((p) => (p.id === id ? { ...p, name: newName } : p))
    setMasterPresets(updated)
    await saveMasterPresets(updated)
  }

  const template = useMemo(() => {
    if (currentRoute === ROUTES.gpsPdf) {
      return {
        eyebrow: 'PPTXPro',
        title: 'PPTX to PDF (OCR Studio)',
        subtext: 'Upload any PPTX presentation or PDF document. Scans custom slide/page height from bottom for GPS coordinates or embeds a searchable/selectable text layer.',
        masterBgUrl: '',
        slots: [],
        themeLabel: '',
      }
    }
    if (currentRoute === ROUTES.pptxToPdf) {
      return {
        eyebrow: 'PPTXPro',
        title: 'PPTX to PDF Converter',
        subtext: 'Upload a PPTX presentation and convert every slide into a high-quality PDF document. Select slides, reorder pages, and download instantly.',
        masterBgUrl: '',
        slots: [],
        themeLabel: '',
      }
    }
    if (currentRoute === ROUTES.collage) {
      return {
        eyebrow: 'PPTXPro',
        title: 'Collage Maker',
        subtext: 'Upload images in bulk, arrange them in a 2x3 or 3x3 grid, adjust margins, and export to PPTX or ZIP.',
        masterBgUrl: '',
        slots: [],
        themeLabel: '',
      }
    }
    if (currentRoute === ROUTES.mergePdf) {
      return {
        eyebrow: 'PPTXPro',
        title: 'Merge PDF Documents',
        subtext: 'Upload multiple PDF files to combine them into a single PDF document and check total page count.',
        masterBgUrl: '',
        slots: [],
        themeLabel: '',
      }
    }
    if (currentRoute === ROUTES.pdf) {
      return {
        eyebrow: 'PPTXPro',
        title: 'PDF to PPTX Converter',
        subtext: 'Upload a PDF file to convert its pages into PowerPoint slides. Reorder, exclude, and download them.',
        masterBgUrl: '',
        slots: [],
        themeLabel: '',
      }
    }
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
    if (currentRoute === ROUTES.merge) {
      return {
        eyebrow: 'PPTXPro',
        title: 'Merge Presentations',
        subtext: 'Upload multiple PPTX files to merge their slides in sequential order.',
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
            borderRadius: p.borderRadius ? `${p.borderRadius}px` : '0px',
          },
        })),
        textBoxes: customLayout?.textboxes || [],
      }
    }
    if (currentRoute === ROUTES.pptxEditor) {
      return {
        eyebrow: 'PPTXPro',
        title: 'PPTX Editor',
        subtext: 'Edit basic text & image tags preserving slide structure.',
        masterBgUrl: '',
        slots: [],
        themeLabel: '',
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
  const storageKey =
    currentRoute === ROUTES.master && activePresetId
      ? `pptxpro:slides:v1:${currentRoute}:${activePresetId}`
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
  const routeFallback = <p className="app__note">Loading tool...</p>
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
    if (
      !TEMPLATES[currentRoute] &&
      currentRoute !== ROUTES.master &&
      currentRoute !== ROUTES.extract &&
      currentRoute !== ROUTES.merge &&
      currentRoute !== ROUTES.pdf &&
      currentRoute !== ROUTES.mergePdf &&
      currentRoute !== ROUTES.collage &&
      currentRoute !== ROUTES.gpsPdf &&
      currentRoute !== ROUTES.pptxToPdf &&
      currentRoute !== ROUTES.pptxEditor
    ) {
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
    savePairsToStorage(storageKey, pairs, {
      slotKeys,
      requiresText,
      hasPairContent,
      trimTrailingEmptyPairs,
    })
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
      const { importPptxSlides, importFirstLastSlideData } = await loadImportPptxModule()
      const { pairs: importedPairs, importedSlides, emptySlides } =
        await importPptxSlides(file, {
          skipFirstSlides,
          skipLastSlides,
          imageCount: template.imageCount || slotKeys.length,
          textboxDefs: currentRoute === ROUTES.master ? (customLayout?.textboxes || []) : [],
        })

      setPairs(
        normalizePairs(importedPairs, { slotKeys, requiresText, textDefault }),
      )

      if (currentRoute === ROUTES.master) {
        try {
          const { firstSlide, lastSlide } = await importFirstLastSlideData(file)
          if (firstSlide) {
            setImportedFirstSlide(firstSlide)
            importedFirstSlideRef.current = firstSlide
            setUseTemplateFirst(false)
          }
          if (lastSlide) {
            setImportedLastSlide(lastSlide)
            importedLastSlideRef.current = lastSlide
            setUseTemplateLast(false)
          }
        } catch (err) {
          console.warn('Could not extract first/last slide elements:', err)
        }
      }

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
      setImportedFirstSlide(null)
      importedFirstSlideRef.current = null
      setUseTemplateFirst(false)
      setImportedLastSlide(null)
      importedLastSlideRef.current = null
      setUseTemplateLast(false)
      setImportStatus({ type: 'idle', message: '' })
      setDragIndex(null)
      setDragOverIndex(null)
    }
  }

  // "Use Template" / "Use Imported" handlers — toggle slide mode while retaining imported state
  const handleUseTemplateFirst = () => setUseTemplateFirst(true)
  const handleUseImportedFirst = () => setUseTemplateFirst(false)

  const handleUseTemplateLast = () => setUseTemplateLast(true)
  const handleUseImportedLast = () => setUseTemplateLast(false)

  const buildReportOptions = (overrides = {}) => {
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
    return {
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
      ...overrides,
    }
  }

  const triggerBlobDownload = (blob, fileName) => {
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = fileName
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    setTimeout(() => URL.revokeObjectURL(url), 5000)
  }

  const reportFileHandleRef = useRef(null)

  const handleSaveReportDirect = async (forceSaveAs = false) => {
    if (!canDownload) return
    try {
      setIsGenerating(true)
      const { generateReport } = await loadGenerateReportModule()
      const { blob, fileName } = await generateReport(pairs, buildReportOptions())

      let finalBlob = blob
      if (currentRoute === ROUTES.master && (importedFirstSlide || importedLastSlide)) {
        const { postProcessPptxWithImportedSlides } = await loadImportPptxModule()
        const firstOpts = importedFirstSlide && !useTemplateFirst
          ? { importedSlide: importedFirstSlide, editedSlide: importedFirstSlideRef.current }
          : null
        const lastOpts = importedLastSlide && !useTemplateLast
          ? { importedSlide: importedLastSlide, editedSlide: importedLastSlideRef.current }
          : null
        finalBlob = await postProcessPptxWithImportedSlides(blob, firstOpts, lastOpts)
      }

      if (!forceSaveAs && reportFileHandleRef.current && typeof reportFileHandleRef.current.createWritable === 'function') {
        const writable = await reportFileHandleRef.current.createWritable()
        await writable.write(finalBlob)
        await writable.close()
        alert(`✓ Saved directly to "${reportFileHandleRef.current.name}" on disk!`)
      } else if (typeof window.showSaveFilePicker === 'function') {
        const handle = await window.showSaveFilePicker({
          suggestedName: fileName,
          types: [
            {
              description: 'PowerPoint Presentation',
              accept: {
                'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'],
              },
            },
          ],
        })
        const writable = await handle.createWritable()
        await writable.write(finalBlob)
        await writable.close()
        reportFileHandleRef.current = handle
        alert(`✓ Saved to "${handle.name}" on disk!`)
      } else {
        triggerBlobDownload(finalBlob, fileName)
      }
    } catch (err) {
      if (err.name !== 'AbortError') {
        console.error('Save report error:', err)
        alert('Failed to save presentation: ' + err.message)
      }
    } finally {
      setIsGenerating(false)
    }
  }

  const handleDownload = async () => {
    if (!canDownload) return
    try {
      setIsGenerating(true)
      const { generateReport } = await loadGenerateReportModule()
      const { blob, fileName } = await generateReport(pairs, buildReportOptions())

      if (currentRoute === ROUTES.master && (importedFirstSlide || importedLastSlide)) {
        const { postProcessPptxWithImportedSlides } = await loadImportPptxModule()
        const firstOpts = importedFirstSlide && !useTemplateFirst
          ? { importedSlide: importedFirstSlide, editedSlide: importedFirstSlideRef.current }
          : null
        const lastOpts = importedLastSlide && !useTemplateLast
          ? { importedSlide: importedLastSlide, editedSlide: importedLastSlideRef.current }
          : null
        const finalBlob = await postProcessPptxWithImportedSlides(blob, firstOpts, lastOpts)
        triggerBlobDownload(finalBlob, fileName)
      } else {
        triggerBlobDownload(blob, fileName)
      }
    } finally {
      setIsGenerating(false)
    }
  }

  const [isGeneratingCompleted, setIsGeneratingCompleted] = useState(false)
  const canDownloadCompleted = slideCount > 0 && !isGenerating && !isGeneratingCompleted && !isImporting

  const handleDownloadCompleted = async () => {
    if (!canDownloadCompleted) return
    try {
      setIsGeneratingCompleted(true)
      const { generateCompletedSlidesReport } = await loadGenerateReportModule()
      const baseOptions = buildReportOptions()
      const completedFileName = baseOptions.fileName
        ? baseOptions.fileName.replace(/\.pptx$/i, '_Completed.pptx')
        : `${template.fileNamePrefix || 'Report'}_Completed.pptx`
      const { blob, fileName } = await generateCompletedSlidesReport(
        pairs, { ...baseOptions, fileName: completedFileName }
      )
      if (currentRoute === ROUTES.master && (importedFirstSlide || importedLastSlide)) {
        const { postProcessPptxWithImportedSlides } = await loadImportPptxModule()
        const firstOpts = importedFirstSlide && !useTemplateFirst
          ? { importedSlide: importedFirstSlide, editedSlide: importedFirstSlideRef.current }
          : null
        const lastOpts = importedLastSlide && !useTemplateLast
          ? { importedSlide: importedLastSlide, editedSlide: importedLastSlideRef.current }
          : null
        const finalBlob = await postProcessPptxWithImportedSlides(blob, firstOpts, lastOpts)
        triggerBlobDownload(finalBlob, fileName)
      } else {
        triggerBlobDownload(blob, fileName)
      }
    } finally {
      setIsGeneratingCompleted(false)
    }
  }

  const [isExportingZip, setIsExportingZip] = useState(false)

  const incompletePairsForZip = pairs.filter((pair) => {
    if (!hasPairContent(pair, { slotKeys, requiresText })) return false
    return !isPairComplete(pair, { slotKeys, requiresText })
  })
  const canExportZip = incompletePairsForZip.length > 0 && !isExportingZip && !isGenerating && !isImporting

  const handleExportRemainingPics = async () => {
    if (!canExportZip) return
    try {
      setIsExportingZip(true)
      const JSZip = await loadJsZip()
      const zip = new JSZip()
      let fileIndex = 1

      for (const pair of incompletePairsForZip) {
        for (const key of slotKeys) {
          const imgData = pair?.[key]
          if (!imgData || typeof imgData !== 'string') continue

          let blob
          let ext = 'jpg'

          if (imgData.startsWith('data:')) {
            const mimeMatch = imgData.match(/^data:([^;]+);base64,/)
            const mime = mimeMatch?.[1] || 'image/jpeg'
            const b64 = imgData.replace(/^data:[^;]+;base64,/, '')
            const byteChars = atob(b64)
            const byteArr = new Uint8Array(byteChars.length)
            for (let i = 0; i < byteChars.length; i++) {
              byteArr[i] = byteChars.charCodeAt(i)
            }
            blob = new Blob([byteArr], { type: mime })
            if (mime.includes('png')) ext = 'png'
            else if (mime.includes('webp')) ext = 'webp'
            else if (mime.includes('gif')) ext = 'gif'
            else ext = 'jpg'
          } else {
            try {
              const response = await fetch(imgData)
              blob = await response.blob()
              const ct = blob.type || ''
              if (ct.includes('png')) ext = 'png'
              else if (ct.includes('webp')) ext = 'webp'
              else if (ct.includes('gif')) ext = 'gif'
              else ext = 'jpg'
            } catch {
              continue
            }
          }

          const slotLabel = key.replace(/([A-Z])/g, '_$1').toUpperCase()
          zip.file(`slide_${String(fileIndex).padStart(3, '0')}_${slotLabel}.${ext}`, blob)
          fileIndex++
        }
      }

      if (fileIndex === 1) {
        alert('No images found in incomplete slides.')
        return
      }

      const zipBlob = await zip.generateAsync({ type: 'blob' })
      const url = URL.createObjectURL(zipBlob)
      const a = document.createElement('a')
      a.href = url
      const prefix = template.fileNamePrefix || 'Report'
      a.download = `${prefix}_Remaining_Pics.zip`
      a.click()
      URL.revokeObjectURL(url)
    } finally {
      setIsExportingZip(false)
    }
  }

  return (
    <main
      className={`app${currentRoute === ROUTES.compliance ? ' app--compliance' : ''}${
        currentRoute === ROUTES.desilting ? ' app--desilting' : ''
      }${currentRoute === ROUTES.dailyPlot ? ' app--daily-plot' : ''}`}
    >
      <Navbar
        currentRoute={currentRoute}
        dailyVariant={dailyVariant}
        setDailyVariant={setDailyVariant}
        designerMode={designerMode}
        setDesignerMode={setDesignerMode}
        customLayout={customLayout}
        ROUTES={ROUTES}
      />

      {currentRoute !== ROUTES.pptxEditor && (
        <header className="app__header">
          <div>
            <p className="app__eyebrow">{template.eyebrow}</p>
            <h1>{template.title}</h1>
            <p className="app__subtext">{template.subtext}</p>
          </div>
          <div className="app__actions">
            {!(currentRoute === ROUTES.master && designerMode === 'design') &&
              currentRoute !== ROUTES.extract &&
              currentRoute !== ROUTES.collage &&
              currentRoute !== ROUTES.merge &&
              currentRoute !== ROUTES.pdf &&
              currentRoute !== ROUTES.mergePdf &&
              currentRoute !== ROUTES.gpsPdf &&
              currentRoute !== ROUTES.pptxToPdf &&
              currentRoute !== ROUTES.pptxEditor && (
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
                    className="button button--completed"
                    onClick={handleDownloadCompleted}
                    disabled={!canDownloadCompleted}
                    title={`Export only the ${slideCount} complete slide(s) as PPTX`}
                  >
                    {isGeneratingCompleted ? 'Building...' : `✅ Completed (${slideCount})`}
                  </button>
                  <button
                    type="button"
                    className="button button--zip"
                    onClick={handleExportRemainingPics}
                    disabled={!canExportZip}
                    title={`Export images from ${incompletePairsForZip.length} incomplete slide(s) as ZIP`}
                  >
                    {isExportingZip ? 'Zipping...' : `📦 Remaining Pics (${incompletePairsForZip.length})`}
                  </button>
                  <button
                    type="button"
                    className="button"
                    style={{ background: 'linear-gradient(135deg, #10b981, #059669)', borderColor: 'transparent', color: '#fff' }}
                    onClick={() => handleSaveReportDirect(false)}
                    disabled={!canDownload}
                    title="Save report directly to file on disk without downloading separate copies"
                  >
                    {isGenerating ? 'Building...' : '💾 Save File'}
                  </button>
                  <button
                    type="button"
                    className="button"
                    onClick={handleDownload}
                    disabled={!canDownload}
                    title="Download a separate copy file via browser"
                  >
                    {isGenerating ? 'Building PPTX...' : '⬇️ Download Report'}
                  </button>
                </>
              )}
          </div>
        </header>
      )}

      {importStatus.message && (
        <p className={`app__note app__note--${importStatus.type}`}>
          {importStatus.message}
        </p>
      )}

      {currentRoute === ROUTES.gpsPdf ? (
        <Suspense fallback={routeFallback}>
          <PptxToPdfOcr />
        </Suspense>
      ) : currentRoute === ROUTES.pptxToPdf ? (
        <Suspense fallback={routeFallback}>
          <PptxToPdf />
        </Suspense>
      ) : currentRoute === ROUTES.pptxEditor ? (
        <Suspense fallback={routeFallback}>
          <PptxEditor />
        </Suspense>
      ) : currentRoute === ROUTES.extract ? (
        <Suspense fallback={routeFallback}>
          <ImageExtractor />
        </Suspense>
      ) : currentRoute === ROUTES.collage ? (
        <Suspense fallback={routeFallback}>
          <CollageMaker />
        </Suspense>
      ) : currentRoute === ROUTES.merge ? (
        <Suspense fallback={routeFallback}>
          <PptxMerger />
        </Suspense>
      ) : currentRoute === ROUTES.pdf ? (
        <Suspense fallback={routeFallback}>
          <PdfToPptx />
        </Suspense>
      ) : currentRoute === ROUTES.mergePdf ? (
        <Suspense fallback={routeFallback}>
          <PdfMerger />
        </Suspense>
      ) : currentRoute === ROUTES.master && designerMode === 'design' ? (
        <Suspense fallback={routeFallback}>
          <MasterDesigner
            customLayout={customLayout}
            onSave={async (layout) => {
              if (activePresetId) {
                const updated = masterPresets.map((p) =>
                  p.id === activePresetId ? { ...p, layout } : p,
                )
                setMasterPresets(updated)
                await saveMasterPresets(updated)
                alert('Active preset updated! Switch to "Use Template" to use it.')
                setDesignerMode('use')
              } else {
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
        </Suspense>
      ) : (
        <>
          <section className="card card--theme">
            <div>
              <p className="label">{template.themeLabel}</p>
              <p className="value">
                Using background image:{' '}
                <span>
                  {template.masterBgUrl
                    ? template.masterBgUrl.startsWith('data:')
                      ? 'Custom background data'
                      : template.masterBgUrl
                    : 'None'}
                </span>
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
            {/* ── First slide: imported editor or template slide ── */}
            {currentRoute === ROUTES.master && importedFirstSlide && !useTemplateFirst ? (
              <ImportedSlideEditor
                title="First Slide (from imported PPTX)"
                slideData={importedFirstSlide}
                templateUrl={customLayout?.firstSlideUrl}
                onUseTemplate={handleUseTemplateFirst}
                onChange={(data) => {
                  importedFirstSlideRef.current = data
                }}
              />
            ) : (
              currentRoute === ROUTES.master && (
                <>
                  {importedFirstSlide && useTemplateFirst && (
                    <article className="pair" style={{ border: '2px solid rgba(124, 58, 237, 0.4)', background: 'rgba(124, 58, 237, 0.03)' }}>
                      <div className="pair__header">
                        <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                          <p className="label" style={{ margin: 0, color: '#7c3aed', fontWeight: 700 }}>
                            📎 First Slide: Using Default Template Slide
                          </p>
                          <span className="pair__status">Template Mode Active</span>
                        </div>
                        <button
                          type="button"
                          className="ghost"
                          onClick={handleUseImportedFirst}
                          style={{ borderColor: '#7c3aed', color: '#7c3aed', fontWeight: 600 }}
                          title="Switch back to using and editing the first slide imported from the PPTX"
                        >
                          ↩ Use Imported Slide
                        </button>
                      </div>
                    </article>
                  )}

                  {((customLayout?.firstSlidePlaceholders?.length > 0) || (customLayout?.firstSlideTextboxes?.length > 0)) && (
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
                            borderRadius: p.borderRadius ? `${p.borderRadius}px` : '0px',
                          },
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
                </>
              )
            )}

            {pairs.map((pair, index) => (
              <PairCard
                key={`pair-${index}`}
                pair={pair}
                index={index}
                template={template}
                slotKeys={slotKeys}
                requiresText={requiresText}
                movableCount={movableCount}
                dragIndex={dragIndex}
                dragOverIndex={dragOverIndex}
                moveMenuIndex={moveMenuIndex}
                pairRef={(node) => {
                  pairRefs.current[index] = node
                }}
                moveMenuRef={moveMenuRef}
                onUpdatePair={updatePair}
                onSwapImages={swapPairImages}
                onDragStart={handlePairDragStart}
                onDragEnd={handlePairDragEnd}
                onDragOver={handlePairDragOver}
                onDrop={handlePairDrop}
                onMovePair={movePair}
                onSetMoveMenuIndex={setMoveMenuIndex}
              />
            ))}

            {/* ── Last slide: imported editor or template slide ── */}
            {currentRoute === ROUTES.master && importedLastSlide && !useTemplateLast ? (
              <ImportedSlideEditor
                title="Last Slide (from imported PPTX)"
                slideData={importedLastSlide}
                templateUrl={customLayout?.lastSlideUrl}
                onUseTemplate={handleUseTemplateLast}
                onChange={(data) => {
                  importedLastSlideRef.current = data
                }}
              />
            ) : (
              currentRoute === ROUTES.master && (
                <>
                  {importedLastSlide && useTemplateLast && (
                    <article className="pair" style={{ border: '2px solid rgba(124, 58, 237, 0.4)', background: 'rgba(124, 58, 237, 0.03)' }}>
                      <div className="pair__header">
                        <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                          <p className="label" style={{ margin: 0, color: '#7c3aed', fontWeight: 700 }}>
                            📎 Last Slide: Using Default Template Slide
                          </p>
                          <span className="pair__status">Template Mode Active</span>
                        </div>
                        <button
                          type="button"
                          className="ghost"
                          onClick={handleUseImportedLast}
                          style={{ borderColor: '#7c3aed', color: '#7c3aed', fontWeight: 600 }}
                          title="Switch back to using and editing the last slide imported from the PPTX"
                        >
                          ↩ Use Imported Slide
                        </button>
                      </div>
                    </article>
                  )}

                  {((customLayout?.lastSlidePlaceholders?.length > 0) || (customLayout?.lastSlideTextboxes?.length > 0)) && (
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
                            borderRadius: p.borderRadius ? `${p.borderRadius}px` : '0px',
                          },
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
                </>
              )
            )}
          </section>
        </>
      )}
    </main>
  )
}

export default App
