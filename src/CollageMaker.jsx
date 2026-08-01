import { useState, useEffect, useRef, useCallback, useMemo } from 'react'
import PptxGenJS from 'pptxgenjs'
import JSZip from 'jszip'

// ── IndexedDB Scoped Storage Helpers ─────────────────────────────────────────
const DB_NAME = 'pptxpro-collage'
const STORE_IMAGES = 'images'
const STORE_SETTINGS = 'settings'
const DB_VERSION = 1

const openCollageDb = () =>
  new Promise((resolve, reject) => {
    if (typeof window === 'undefined' || !window.indexedDB) {
      resolve(null)
      return
    }
    const request = window.indexedDB.open(DB_NAME, DB_VERSION)
    request.onupgradeneeded = () => {
      const db = request.result
      if (!db.objectStoreNames.contains(STORE_IMAGES)) {
        db.createObjectStore(STORE_IMAGES)
      }
      if (!db.objectStoreNames.contains(STORE_SETTINGS)) {
        db.createObjectStore(STORE_SETTINGS)
      }
    }
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error)
  })

const loadFromStore = async (storeName, key) => {
  try {
    const db = await openCollageDb()
    if (!db) return null
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(storeName, 'readonly')
      const store = tx.objectStore(storeName)
      const request = store.get(key)
      request.onsuccess = () => resolve(request.result ?? null)
      request.onerror = () => reject(request.error)
      tx.oncomplete = () => db.close()
    })
  } catch {
    return null
  }
}

const saveToStore = async (storeName, key, value) => {
  try {
    const db = await openCollageDb()
    if (!db) return
    await new Promise((resolve, reject) => {
      const tx = db.transaction(storeName, 'readwrite')
      const store = tx.objectStore(storeName)
      store.put(value, key)
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
    // Ignore storage error
  }
}

const removeFromStore = async (storeName, key) => {
  try {
    const db = await openCollageDb()
    if (!db) return
    await new Promise((resolve, reject) => {
      const tx = db.transaction(storeName, 'readwrite')
      const store = tx.objectStore(storeName)
      store.delete(key)
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
    // Ignore storage error
  }
}

// ── Image Rendering Utilities ────────────────────────────────────────────────
const loadImage = (src) =>
  new Promise((resolve, reject) => {
    const img = new Image()
    img.crossOrigin = 'anonymous'
    img.onload = () => resolve(img)
    img.onerror = (err) => reject(new Error('Failed to load image: ' + err.message))
    img.src = src
  })

const drawImageProp = (ctx, img, x, y, w, h, mode = 'cover') => {
  if (mode === 'stretch') {
    ctx.drawImage(img, x, y, w, h)
    return
  }
  const imgRatio = img.width / img.height
  const boundsRatio = w / h
  let sourceX = 0, sourceY = 0, sourceWidth = img.width, sourceHeight = img.height
  let destX = x, destY = y, destWidth = w, destHeight = h

  if (mode === 'cover') {
    if (imgRatio > boundsRatio) {
      sourceWidth = img.height * boundsRatio
      sourceX = (img.width - sourceWidth) / 2
    } else {
      sourceHeight = img.width / boundsRatio
      sourceY = (img.height - sourceHeight) / 2
    }
  } else {
    // contain
    if (imgRatio > boundsRatio) {
      destHeight = w / imgRatio
      destY = y + (h - destHeight) / 2
    } else {
      destWidth = h * imgRatio
      destX = x + (w - destWidth) / 2
    }
  }
  ctx.drawImage(img, sourceX, sourceY, sourceWidth, sourceHeight, destX, destY, destWidth, destHeight)
}

const renderCollageToCanvas = async (images, rows, cols, spacing, fitMode, bgColor, orientation = 'portrait') => {
  const canvas = document.createElement('canvas')
  const isPortrait = orientation === 'portrait'
  canvas.width = isPortrait ? 1440 : 1920
  canvas.height = isPortrait ? 1920 : 1080
  const ctx = canvas.getContext('2d')

  // Background
  ctx.fillStyle = bgColor || '#ffffff'
  ctx.fillRect(0, 0, canvas.width, canvas.height)

  // Layout cell calculation
  const availableWidth = canvas.width - spacing * (cols + 1)
  const availableHeight = canvas.height - spacing * (rows + 1)
  const cellWidth = availableWidth / cols
  const cellHeight = availableHeight / rows

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const idx = r * cols + c
      if (idx < images.length) {
        const imgObj = images[idx]
        if (imgObj && imgObj.dataUrl) {
          try {
            const img = await loadImage(imgObj.dataUrl)
            const x = spacing + c * (cellWidth + spacing)
            const y = spacing + r * (cellHeight + spacing)

            ctx.save()
            // Clip to cell boundary
            ctx.beginPath()
            ctx.rect(x, y, cellWidth, cellHeight)
            ctx.clip()

            drawImageProp(ctx, img, x, y, cellWidth, cellHeight, fitMode)
            ctx.restore()
          } catch (e) {
            console.error('Error drawing image on canvas:', e)
          }
        }
      }
    }
  }

  return canvas
}

const readFileAsDataUrl = (file) =>
  new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result)
    reader.onerror = () => reject(reader.error)
    reader.readAsDataURL(file)
  })

// ── Drag & Drop Reordering List Helpers ─────────────────────────────────────
const reorder = (list, from, to) => {
  const next = [...list]
  const [moved] = next.splice(from, 1)
  next.splice(to, 0, moved)
  return next
}

// ── Sub-components ────────────────────────────────────────────────────────────

function UploadZone({ onFiles, isLoading }) {
  const inputRef = useRef(null)
  const [isDragging, setIsDragging] = useState(false)

  const handleFiles = async (files) => {
    if (!files || files.length === 0) return
    const validFiles = Array.from(files).filter((file) =>
      file.type.startsWith('image/')
    )
    if (validFiles.length === 0) {
      alert('Please upload valid image files.')
      return
    }
    onFiles(validFiles)
  }

  return (
    <div
      className={`collage-upload${isDragging ? ' is-dragging' : ''}${isLoading ? ' is-loading' : ''}`}
      onClick={() => !isLoading && inputRef.current?.click()}
      onDragOver={(e) => {
        e.preventDefault()
        setIsDragging(true)
      }}
      onDragEnter={(e) => {
        e.preventDefault()
        setIsDragging(true)
      }}
      onDragLeave={() => setIsDragging(false)}
      onDrop={(e) => {
        e.preventDefault()
        setIsDragging(false)
        if (e.dataTransfer.files) handleFiles(e.dataTransfer.files)
      }}
    >
      <div className="collage-upload__inner">
        <span className="collage-upload__icon" aria-hidden="true">🖼️</span>
        <p className="collage-upload__label">Drop your images here or click to browse</p>
        <p className="collage-upload__hint">Upload multiple images to generate grids</p>
      </div>
      <input
        ref={inputRef}
        type="file"
        multiple
        accept="image/*"
        className="collage-upload__input"
        onChange={(e) => {
          if (e.target.files) handleFiles(e.target.files)
          e.target.value = ''
        }}
      />
    </div>
  )
}

function ImageCard({
  item,
  onDelete,
  onDragStart,
  onDragOver,
  onDrop,
  onDragEnd,
  isDragging,
  isDragOver,
}) {
  return (
    <div
      className={`collage-card${isDragging ? ' is-dragging' : ''}${isDragOver ? ' is-drag-over' : ''}`}
      draggable
      onDragStart={onDragStart}
      onDragOver={(e) => {
        e.preventDefault()
        onDragOver()
      }}
      onDrop={(e) => {
        e.preventDefault()
        onDrop()
      }}
      onDragEnd={onDragEnd}
    >
      <div className="collage-card__img-wrap">
        <img src={item.dataUrl} alt={item.name} className="collage-card__img" draggable={false} />
        <button
          type="button"
          className="collage-card__delete"
          onClick={(e) => {
            e.stopPropagation()
            onDelete(item.id)
          }}
          title="Remove image"
        >
          &times;
        </button>
      </div>
      <div className="collage-card__footer">
        <span className="collage-card__label" title={item.name}>
          {item.name.length > 18 ? item.name.substring(0, 15) + '...' : item.name}
        </span>
        <span className="collage-card__drag-hint" aria-hidden="true">⠿</span>
      </div>
    </div>
  )
}

// ── Main Component ────────────────────────────────────────────────────────────
export function CollageMaker() {
  const [images, setImages] = useState([])
  const [gridRows, setGridRows] = useState(2)
  const [gridCols, setGridCols] = useState(3)
  const [spacing, setSpacing] = useState(14)
  const [fitMode, setFitMode] = useState('stretch') // stretch, cover or contain
  const [orientation, setOrientation] = useState('portrait') // portrait or landscape
  const [bgColor, setBgColor] = useState('#ffffff')
  const [isLoading, setIsLoading] = useState(false)
  const [isExporting, setIsExporting] = useState(false)
  const [status, setStatus] = useState({ type: 'idle', message: '' })

  const [dragIndex, setDragIndex] = useState(null)
  const [dragOverIndex, setDragOverIndex] = useState(null)
  const [activeTab, setActiveTab] = useState('preview') // 'preview' or 'media'

  const fileInputRef = useRef(null)

  // Capacity of current grid layout
  const capacity = useMemo(() => gridRows * gridCols, [gridRows, gridCols])

  // Split images into slide groups
  const collageSlides = useMemo(() => {
    const slides = []
    for (let i = 0; i < images.length; i += capacity) {
      slides.push(images.slice(i, i + capacity))
    }
    return slides
  }, [images, capacity])

  // Load saved state from DB
  useEffect(() => {
    const loadState = async () => {
      try {
        const storedImages = await loadFromStore(STORE_IMAGES, 'collage-images-list')
        const storedSettings = await loadFromStore(STORE_SETTINGS, 'collage-settings')
        if (storedImages) setImages(storedImages)
        if (storedSettings) {
          if (storedSettings.gridRows) setGridRows(storedSettings.gridRows)
          if (storedSettings.gridCols) setGridCols(storedSettings.gridCols)
          if (storedSettings.spacing !== undefined) setSpacing(storedSettings.spacing)
          if (storedSettings.fitMode) setFitMode(storedSettings.fitMode)
          if (storedSettings.orientation) setOrientation(storedSettings.orientation)
          if (storedSettings.bgColor) setBgColor(storedSettings.bgColor)
        }
      } catch (err) {
        console.error('Failed to load state from database', err)
      }
    }
    loadState()
  }, [])

  // Save changes to DB
  const saveState = useCallback(async (updatedImages, settingsUpdate = {}) => {
    if (updatedImages !== null) {
      await saveToStore(STORE_IMAGES, 'collage-images-list', updatedImages)
    }
    if (Object.keys(settingsUpdate).length > 0) {
      const currentSettings = (await loadFromStore(STORE_SETTINGS, 'collage-settings')) || {}
      await saveToStore(STORE_SETTINGS, 'collage-settings', {
        ...currentSettings,
        ...settingsUpdate,
      })
    }
  }, [])

  const handleFilesUploaded = async (files) => {
    try {
      setIsLoading(true)
      setStatus({ type: 'working', message: 'Processing uploaded images...' })

      const newLoadedImages = []
      for (const file of files) {
        const dataUrl = await readFileAsDataUrl(file)
        newLoadedImages.push({
          id: `img_${Date.now()}_${Math.random().toString(36).slice(2)}`,
          name: file.name,
          dataUrl,
        })
      }

      setImages((prev) => {
        const next = [...prev, ...newLoadedImages]
        saveState(next, {})
        return next
      })

      setStatus({
        type: 'success',
        message: `Successfully added ${files.length} image${files.length !== 1 ? 's' : ''} to library.`,
      })
    } catch (err) {
      setStatus({
        type: 'error',
        message: 'Failed to upload images: ' + (err?.message || 'Unknown error'),
      })
    } finally {
      setIsLoading(false)
    }
  }

  const handleDeleteImage = (id) => {
    setImages((prev) => {
      const next = prev.filter((img) => img.id !== id)
      saveState(next, {})
      return next
    })
    setStatus({ type: 'idle', message: '' })
  }

  const handleClearAll = () => {
    if (window.confirm('Are you sure you want to clear all uploaded images?')) {
      setImages([])
      void removeFromStore(STORE_IMAGES, 'collage-images-list')
      setStatus({ type: 'idle', message: '' })
    }
  }

  // Preset Handlers
  const applyPreset = (rows, cols) => {
    setGridRows(rows)
    setGridCols(cols)
    void saveState(null, { gridRows: rows, gridCols: cols })
  }

  const handleSpacingChange = (val) => {
    setSpacing(val)
    void saveState(null, { spacing: val })
  }

  const handleFitModeChange = (mode) => {
    setFitMode(mode)
    void saveState(null, { fitMode: mode })
  }

  const handleOrientationChange = (orient) => {
    setOrientation(orient)
    void saveState(null, { orientation: orient })
  }

  const handleBgColorChange = (color) => {
    setBgColor(color)
    void saveState(null, { bgColor: color })
  }

  // Drag and drop sorting list
  const handleDragStart = (index) => {
    setDragIndex(index)
    setDragOverIndex(null)
  }

  const handleDragOver = (index) => {
    if (dragIndex === null || dragIndex === index) return
    setDragOverIndex(index)
  }

  const handleDrop = (index) => {
    if (dragIndex === null || dragIndex === index) return
    setImages((prev) => {
      const next = reorder(prev, dragIndex, index)
      saveState(next, {})
      return next
    })
    setDragIndex(null)
    setDragOverIndex(null)
  }

  const handleDragEnd = () => {
    setDragIndex(null)
    setDragOverIndex(null)
  }

  // Export collages to ZIP
  const handleExportZip = async () => {
    if (images.length === 0) return
    try {
      setIsExporting(true)
      setStatus({ type: 'working', message: 'Generating collages for download...' })

      const zip = new JSZip()
      const d = new Date()
      const dateStr = `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`

      for (let i = 0; i < collageSlides.length; i++) {
        const slideImages = collageSlides[i]
        setStatus({
          type: 'working',
          message: `Rendering collage slide ${i + 1} of ${collageSlides.length}...`,
        })

        // Render Canvas
        const canvas = await renderCollageToCanvas(
          slideImages,
          gridRows,
          gridCols,
          spacing,
          fitMode,
          bgColor,
          orientation
        )

        const jpegUrl = canvas.toDataURL('image/jpeg', 0.92)
        const base64Data = jpegUrl.replace(/^data:image\/jpeg;base64,/, '')
        zip.file(`collage_slide_${String(i + 1).padStart(2, '0')}.jpg`, base64Data, { base64: true })
      }

      setStatus({ type: 'working', message: 'Packing ZIP file...' })
      const blob = await zip.generateAsync({ type: 'blob' })
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `Collages_${dateStr}.zip`
      a.click()
      URL.revokeObjectURL(url)

      setStatus({
        type: 'success',
        message: `Successfully exported ${collageSlides.length} collage${collageSlides.length !== 1 ? 's' : ''} in a ZIP file!`,
      })
    } catch (err) {
      setStatus({
        type: 'error',
        message: 'Failed to export collages: ' + (err?.message || 'Unknown error'),
      })
    } finally {
      setIsExporting(false)
    }
  }

  // Export collages to PPTX
  const handleExportPptx = async () => {
    if (images.length === 0) return
    try {
      setIsExporting(true)
      setStatus({ type: 'working', message: 'Generating PowerPoint slides...' })

      const pptx = new PptxGenJS()
      const isPortrait = orientation === 'portrait'
      
      if (isPortrait) {
        pptx.defineLayout({ name: 'PORTRAIT_3_4', width: 7.5, height: 10.0 })
        pptx.layout = 'PORTRAIT_3_4'
      } else {
        pptx.layout = 'LAYOUT_WIDE'
      }

      const d = new Date()
      const dateStr = `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`

      for (let i = 0; i < collageSlides.length; i++) {
        const slideImages = collageSlides[i]
        setStatus({
          type: 'working',
          message: `Rendering slide ${i + 1} of ${collageSlides.length}...`,
        })

        // Render Canvas
        const canvas = await renderCollageToCanvas(
          slideImages,
          gridRows,
          gridCols,
          spacing,
          fitMode,
          bgColor,
          orientation
        )

        const jpegUrl = canvas.toDataURL('image/jpeg', 0.90)

        // Add Slide and apply full screen image
        const slide = pptx.addSlide()
        slide.addImage({
          data: jpegUrl,
          x: 0,
          y: 0,
          w: isPortrait ? 7.5 : 13.333,
          h: isPortrait ? 10.0 : 7.5,
        })
      }

      setStatus({ type: 'working', message: 'Saving presentation...' })
      await pptx.writeFile({ fileName: `Collage_Report_${dateStr}.pptx` })

      setStatus({
        type: 'success',
        message: `Successfully generated PPTX presentation with ${collageSlides.length} slide${collageSlides.length !== 1 ? 's' : ''}!`,
      })
    } catch (err) {
      setStatus({
        type: 'error',
        message: 'Failed to build presentation: ' + (err?.message || 'Unknown error'),
      })
    } finally {
      setIsExporting(false)
    }
  }

  return (
    <div className="collage-maker">
      {images.length === 0 ? (
        <UploadZone onFiles={handleFilesUploaded} isLoading={isLoading} />
      ) : (
        <div className="collage-maker__workspace">
          {/* Controls Panel */}
          <aside className="collage-controls card">
            <div className="collage-controls__section">
              <p className="label">Grid Layout</p>
              <div className="collage-controls__presets">
                <button
                  type="button"
                  className={`ghost ${gridRows === 2 && gridCols === 3 ? 'is-active' : ''}`}
                  onClick={() => applyPreset(2, 3)}
                  disabled={isExporting}
                >
                  2 &times; 3 Grid
                </button>
                <button
                  type="button"
                  className={`ghost ${gridRows === 3 && gridCols === 2 ? 'is-active' : ''}`}
                  onClick={() => applyPreset(3, 2)}
                  disabled={isExporting}
                >
                  3 &times; 2 Grid
                </button>
                <button
                  type="button"
                  className={`ghost ${gridRows === 3 && gridCols === 3 ? 'is-active' : ''}`}
                  onClick={() => applyPreset(3, 3)}
                  disabled={isExporting}
                >
                  3 &times; 3 Grid
                </button>
                <button
                  type="button"
                  className={`ghost ${gridRows === 2 && gridCols === 2 ? 'is-active' : ''}`}
                  onClick={() => applyPreset(2, 2)}
                  disabled={isExporting}
                >
                  2 &times; 2 Grid
                </button>
              </div>
            </div>

            <div className="collage-controls__section">
              <div className="collage-controls__slider-header">
                <p className="label">Cell Spacing &amp; Border</p>
                <span className="collage-controls__slider-value">{spacing}px</span>
              </div>
              <input
                type="range"
                min="0"
                max="60"
                value={spacing}
                onChange={(e) => handleSpacingChange(parseInt(e.target.value, 10))}
                className="collage-slider"
                disabled={isExporting}
              />
              <span className="collage-controls__hint">Default is 14px. Spacing creates a white margin grid.</span>
            </div>

            <div className="collage-controls__section">
              <p className="label">Page Orientation</p>
              <div className="collage-controls__presets">
                <button
                  type="button"
                  className={`ghost ${orientation === 'portrait' ? 'is-active' : ''}`}
                  onClick={() => handleOrientationChange('portrait')}
                  disabled={isExporting}
                >
                  Portrait (3:4)
                </button>
                <button
                  type="button"
                  className={`ghost ${orientation === 'landscape' ? 'is-active' : ''}`}
                  onClick={() => handleOrientationChange('landscape')}
                  disabled={isExporting}
                >
                  Landscape (16:9)
                </button>
              </div>
            </div>

            <div className="collage-controls__section">
              <p className="label">Image Fitting Style</p>
              <div className="collage-controls__presets">
                <button
                  type="button"
                  className={`ghost ${fitMode === 'stretch' ? 'is-active' : ''}`}
                  onClick={() => handleFitModeChange('stretch')}
                  disabled={isExporting}
                >
                  Stretch to Fit (Default)
                </button>
                <button
                  type="button"
                  className={`ghost ${fitMode === 'cover' ? 'is-active' : ''}`}
                  onClick={() => handleFitModeChange('cover')}
                  disabled={isExporting}
                >
                  Crop (Cover)
                </button>
                <button
                  type="button"
                  className={`ghost ${fitMode === 'contain' ? 'is-active' : ''}`}
                  onClick={() => handleFitModeChange('contain')}
                  disabled={isExporting}
                >
                  Fit In (Contain)
                </button>
              </div>
            </div>

            <div className="collage-controls__section">
              <p className="label">Background Color</p>
              <div className="collage-controls__color-picker-wrap">
                <input
                  type="color"
                  value={bgColor}
                  onChange={(e) => handleBgColorChange(e.target.value)}
                  className="collage-color-picker"
                  disabled={isExporting}
                />
                <span className="collage-controls__color-hex">{bgColor.toUpperCase()}</span>
              </div>
            </div>

            <div className="collage-controls__section collage-controls__actions">
              <button
                type="button"
                className="button button--secondary"
                onClick={() => fileInputRef.current?.click()}
                disabled={isExporting || isLoading}
              >
                Add Images
              </button>
              <button
                type="button"
                className="ghost"
                onClick={handleClearAll}
                disabled={isExporting || isLoading}
              >
                Clear All
              </button>
              <input
                ref={fileInputRef}
                type="file"
                multiple
                accept="image/*"
                className="file-input"
                onChange={(e) => {
                  if (e.target.files) handleFilesUploaded(e.target.files)
                  e.target.value = ''
                }}
              />
            </div>
          </aside>

          {/* Main workspace */}
          <div className="collage-maker__main">
            {/* Toolbar for views and exports */}
            <div className="collage-toolbar card">
              <div className="collage-toolbar__left">
                <button
                  type="button"
                  className={`ghost ${activeTab === 'preview' ? 'is-active' : ''}`}
                  onClick={() => setActiveTab('preview')}
                >
                  Collage Previews ({collageSlides.length})
                </button>
                <button
                  type="button"
                  className={`ghost ${activeTab === 'media' ? 'is-active' : ''}`}
                  onClick={() => setActiveTab('media')}
                >
                  Media Library ({images.length})
                </button>
              </div>
              <div className="collage-toolbar__right">
                <button
                  type="button"
                  className="button button--jpeg"
                  onClick={handleExportZip}
                  disabled={isExporting || images.length === 0}
                  title="Download collages as JPEGs inside a ZIP archive"
                >
                  Download ZIP
                </button>
                <button
                  type="button"
                  className="button"
                  onClick={handleExportPptx}
                  disabled={isExporting || images.length === 0}
                  title="Download collages as a PowerPoint slides presentation"
                >
                  {isExporting ? 'Building...' : 'Export PPTX'}
                </button>
              </div>
            </div>

            {status.message && (
              <p className={`app__note app__note--${status.type}`}>{status.message}</p>
            )}

            {/* Content Tabs */}
            {activeTab === 'preview' ? (
              <div className="collage-previews">
                <p className="collage-maker__hint">
                  Showing {orientation === 'portrait' ? '3:4 portrait' : '16:9 widescreen'} layout previews. Spacing adjusts the padding around and between grid items.
                </p>
                <div className="collage-slides-grid">
                  {collageSlides.map((slideImages, slideIndex) => (
                    <article key={`slide-${slideIndex}`} className="collage-slide-card">
                      <div className="collage-slide-card__header">
                        <span className="collage-slide-card__title">Slide {slideIndex + 1}</span>
                        <span className="collage-slide-card__badge">
                          {slideImages.length} / {capacity} Cells
                        </span>
                      </div>
                      
                      {/* Responsive slide frame mimicking aspect-ratio 3:4 or 16:9 */}
                      <div 
                        className="collage-slide-canvas" 
                        style={{ 
                          backgroundColor: bgColor,
                          aspectRatio: orientation === 'portrait' ? '3 / 4' : '16 / 9'
                        }}
                      >
                        <div 
                          className="collage-slide-grid" 
                          style={{
                            display: 'grid',
                            gridTemplateColumns: `repeat(${gridCols}, 1fr)`,
                            gridTemplateRows: `repeat(${gridRows}, 1fr)`,
                            gap: `${spacing}px`,
                            padding: `${spacing}px`,
                            width: '100%',
                            height: '100%',
                            boxSizing: 'border-box'
                          }}
                        >
                          {/* Render cells */}
                          {Array.from({ length: capacity }).map((_, cellIndex) => {
                            const img = slideImages[cellIndex]
                            return (
                              <div 
                                key={img?.id || `cell-${cellIndex}`} 
                                className="collage-slide-cell"
                              >
                                {img ? (
                                  <img 
                                    src={img.dataUrl} 
                                    className={`collage-slide-img is-${fitMode}`} 
                                    alt="" 
                                  />
                                ) : (
                                  <div className="collage-slide-cell-empty" />
                                )}
                              </div>
                            )
                          })}
                        </div>
                      </div>
                    </article>
                  ))}
                </div>
              </div>
            ) : (
              <div className="collage-media-library">
                <p className="collage-maker__hint">
                  Drag and drop cards to reorder images · Hover card to delete. The collages will adapt automatically.
                </p>
                <div className="collage-media-grid">
                  {images.map((img, index) => (
                    <ImageCard
                      key={img.id}
                      item={img}
                      onDelete={handleDeleteImage}
                      onDragStart={() => handleDragStart(index)}
                      onDragOver={() => handleDragOver(index)}
                      onDrop={() => handleDrop(index)}
                      onDragEnd={handleDragEnd}
                      isDragging={dragIndex === index}
                      isDragOver={dragOverIndex === index}
                    />
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  )
}
