import { useRef, useState, useCallback } from 'react'
import { extractAllImages, exportSlidesWithContext } from './report/importPptx'
import PptxGenJS from 'pptxgenjs'
import JSZip from 'jszip'

// ── Export helpers ────────────────────────────────────────────────────────────

// Slide dimensions for 3:4 portrait (width × height in inches)
const SLIDE_W = 7.5
const SLIDE_H = 10.0

const exportImagesToPptx = async (images, fileName = 'Extracted_Images.pptx') => {
  const pptx = new PptxGenJS()

  // Define a custom 3:4 layout (portrait)
  pptx.defineLayout({ name: 'PORTRAIT_3_4', width: SLIDE_W, height: SLIDE_H })
  pptx.layout = 'PORTRAIT_3_4'

  for (const img of images) {
    const slide = pptx.addSlide()

    // Convert each image to a 3:4 portrait JPEG with a white background/border.
    // This prevents the image from being stretched or distorted in PowerPoint.
    const jpegBase64 = await dataUrlToJpegBase64(img.dataUrl)

    slide.addImage({
      data: `data:image/jpeg;base64,${jpegBase64}`,
      x: 0,
      y: 0,
      w: SLIDE_W,
      h: SLIDE_H,
    })
  }

  await pptx.writeFile({ fileName })
}

/**
 * Convert a dataUrl to a high-resolution JPEG.
 * The image is upscaled to fit inside TARGET_W × TARGET_H while preserving
 * aspect ratio. The canvas is always TARGET_W × TARGET_H with a white
 * background, so even very small/narrow images become properly-sized JPEGs.
 */
// 3:4 portrait canvas dimensions (width × height)
const TARGET_W = 1440
const TARGET_H = 1920

/**
 * Convert a dataUrl to a high-resolution JPEG at 3:4 (portrait) aspect ratio.
 * The image is stretched to fill the entire TARGET_W x TARGET_H canvas.
 */
const dataUrlToJpegBase64 = (dataUrl, quality = 0.92) =>
  new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => {
      const canvas = document.createElement('canvas')
      canvas.width = TARGET_W
      canvas.height = TARGET_H
      const ctx = canvas.getContext('2d')

      // Draw the image stretched to cover the entire canvas
      ctx.imageSmoothingEnabled = true
      ctx.imageSmoothingQuality = 'high'
      ctx.drawImage(img, 0, 0, TARGET_W, TARGET_H)

      const jpegDataUrl = canvas.toDataURL('image/jpeg', quality)
      resolve(jpegDataUrl.replace(/^data:image\/jpeg;base64,/, ''))
    }
    img.onerror = () => reject(new Error('Failed to load image for JPEG conversion'))
    img.src = dataUrl
  })

const exportImagesAsJpegs = async (images, zipFileName = 'Extracted_Images.zip') => {
  const zip = new JSZip()

  // Track how many images belong to each slide so we can pad image indices correctly
  const slideImageCount = {}
  for (const img of images) {
    slideImageCount[img.slideNumber] = (slideImageCount[img.slideNumber] ?? 0) + 1
  }

  // Per-slide counters so filenames are numbered within each folder
  const slideCounter = {}

  for (const img of images) {
    const slideKey = String(img.slideNumber).padStart(3, '0')
    const folderName = `slide_${slideKey}`
    const slideFolder = zip.folder(folderName)

    slideCounter[img.slideNumber] = (slideCounter[img.slideNumber] ?? 0) + 1
    const imgDigits = String(slideImageCount[img.slideNumber]).length
    const imgNum = String(slideCounter[img.slideNumber]).padStart(Math.max(imgDigits, 2), '0')

    const jpegBase64 = await dataUrlToJpegBase64(img.dataUrl)
    const fileName = `slide_${slideKey}_img${imgNum}.jpg`
    slideFolder.file(fileName, jpegBase64, { base64: true })
  }

  const blob = await zip.generateAsync({ type: 'blob' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = zipFileName
  a.click()
  URL.revokeObjectURL(url)
}

// ── Drag-to-reorder helpers ───────────────────────────────────────────────────

const reorder = (list, from, to) => {
  const next = [...list]
  const [moved] = next.splice(from, 1)
  next.splice(to, 0, moved)
  return next
}

// ── Sub-components ────────────────────────────────────────────────────────────

function UploadZone({ onFile, isLoading }) {
  const inputRef = useRef(null)
  const [isDragging, setIsDragging] = useState(false)

  const handleFile = (file) => {
    if (!file) return
    if (!file.name.toLowerCase().endsWith('.pptx')) {
      alert('Please upload a valid .pptx file.')
      return
    }
    onFile(file)
  }

  return (
    <div
      className={`extractor-upload${isDragging ? ' is-dragging' : ''}${isLoading ? ' is-loading' : ''}`}
      onClick={() => !isLoading && inputRef.current?.click()}
      onDragOver={(e) => { e.preventDefault(); setIsDragging(true) }}
      onDragEnter={(e) => { e.preventDefault(); setIsDragging(true) }}
      onDragLeave={() => setIsDragging(false)}
      onDrop={(e) => {
        e.preventDefault()
        setIsDragging(false)
        const file = e.dataTransfer.files?.[0]
        if (file) handleFile(file)
      }}
    >
      {isLoading ? (
        <div className="extractor-upload__inner">
          <div className="extractor-spinner" aria-label="Extracting images…" />
          <p className="extractor-upload__label">Extracting images…</p>
        </div>
      ) : (
        <div className="extractor-upload__inner">
          <span className="extractor-upload__icon" aria-hidden="true">📂</span>
          <p className="extractor-upload__label">Drop your PPTX here or click to browse</p>
          <p className="extractor-upload__hint">Every image from every slide will be extracted</p>
        </div>
      )}
      <input
        ref={inputRef}
        type="file"
        accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
        className="extractor-upload__input"
        onChange={(e) => {
          const file = e.target.files?.[0]
          if (file) handleFile(file)
          e.target.value = ''
        }}
      />
    </div>
  )
}

function ImageCard({ item, isSelected, onToggle, onDragStart, onDragOver, onDrop, onDragEnd, isDragging, isDragOver }) {
  return (
    <div
      className={`extractor-card${isSelected ? ' is-selected' : ''}${isDragging ? ' is-dragging' : ''}${isDragOver ? ' is-drag-over' : ''}`}
      draggable
      onDragStart={onDragStart}
      onDragOver={(e) => { e.preventDefault(); onDragOver() }}
      onDrop={(e) => { e.preventDefault(); onDrop() }}
      onDragEnd={onDragEnd}
      onClick={onToggle}
      title={`Slide ${item.slideNumber}, image ${item.imageIndex + 1} — click to ${isSelected ? 'deselect' : 'select'}`}
    >
      <div className="extractor-card__img-wrap">
        <img
          src={item.dataUrl}
          alt={`Slide ${item.slideNumber} image ${item.imageIndex + 1}`}
          className="extractor-card__img"
          draggable={false}
        />
        <div className={`extractor-card__check${isSelected ? ' is-checked' : ''}`} aria-hidden="true">
          {isSelected ? '✓' : ''}
        </div>
      </div>
      <div className="extractor-card__footer">
        <span className="extractor-card__label">
          Slide {item.slideNumber}
          {item.imageIndex > 0 ? ` · img ${item.imageIndex + 1}` : ''}
        </span>
        <span className="extractor-card__drag-hint" aria-hidden="true">⠿</span>
      </div>
    </div>
  )
}

// ── Main component ────────────────────────────────────────────────────────────

export function ImageExtractor() {
  const [images, setImages] = useState([])
  const [selected, setSelected] = useState(new Set())
  const [originalFile, setOriginalFile] = useState(null)
  const [isLoading, setIsLoading] = useState(false)
  const [isExportingPptx, setIsExportingPptx] = useState(false)
  const [isExportingJpeg, setIsExportingJpeg] = useState(false)
  const [isExportingContext, setIsExportingContext] = useState(false)
  const [status, setStatus] = useState({ type: 'idle', message: '' })
  const [dragIndex, setDragIndex] = useState(null)
  const [dragOverIndex, setDragOverIndex] = useState(null)

  const isBusy = isLoading || isExportingPptx || isExportingJpeg || isExportingContext

  const handleFile = useCallback(async (file) => {
    try {
      setIsLoading(true)
      setStatus({ type: 'working', message: 'Reading PPTX…' })
      setImages([])
      setSelected(new Set())
      setOriginalFile(file)

      const extracted = await extractAllImages(file)

      if (extracted.length === 0) {
        setStatus({ type: 'error', message: 'No images found in this PPTX.' })
        return
      }

      setImages(extracted)
      setSelected(new Set(extracted.map((img) => img.id)))
      const slideCount = new Set(extracted.map((i) => i.slideNumber)).size
      setStatus({
        type: 'success',
        message: `Extracted ${extracted.length} image${extracted.length !== 1 ? 's' : ''} from ${slideCount} slide${slideCount !== 1 ? 's' : ''}.`,
      })
    } catch (err) {
      setStatus({ type: 'error', message: err?.message || 'Extraction failed.' })
    } finally {
      setIsLoading(false)
    }
  }, [])

  const toggleSelect = useCallback((id) => {
    setSelected((prev) => {
      const next = new Set(prev)
      if (next.has(id)) next.delete(id)
      else next.add(id)
      return next
    })
  }, [])

  const selectAll = () => setSelected(new Set(images.map((i) => i.id)))
  const deselectAll = () => setSelected(new Set())
  const selectedImages = images.filter((img) => selected.has(img.id))
  const canExport = selectedImages.length > 0 && !isBusy

  const getDateStr = () => {
    const d = new Date()
    return `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`
  }

  const handleExportPptx = async () => {
    if (!canExport) return
    try {
      setIsExportingPptx(true)
      await exportImagesToPptx(selectedImages, `Extracted_Images_${getDateStr()}.pptx`)
    } catch (err) {
      alert('PPTX export failed: ' + (err?.message || 'Unknown error'))
    } finally {
      setIsExportingPptx(false)
    }
  }

  const handleExportJpeg = async () => {
    if (!canExport) return
    try {
      setIsExportingJpeg(true)
      setStatus({ type: 'working', message: `Converting ${selectedImages.length} image${selectedImages.length !== 1 ? 's' : ''} to JPEG…` })
      await exportImagesAsJpegs(selectedImages, `Extracted_Images_${getDateStr()}.zip`)
      setStatus({ type: 'success', message: `Downloaded ${selectedImages.length} JPEG${selectedImages.length !== 1 ? 's' : ''} as a ZIP file.` })
    } catch (err) {
      setStatus({ type: 'error', message: 'JPEG export failed: ' + (err?.message || 'Unknown error') })
    } finally {
      setIsExportingJpeg(false)
    }
  }

  const handleExportWithContext = async () => {
    if (!canExport || !originalFile) return
    try {
      setIsExportingContext(true)
      setStatus({
        type: 'working',
        message: `Building ${selectedImages.length} slide${selectedImages.length !== 1 ? 's' : ''} with original context…`,
      })
      await exportSlidesWithContext(
        originalFile,
        selectedImages,
        `Slides_With_Context_${getDateStr()}.pptx`,
      )
      setStatus({
        type: 'success',
        message: `Downloaded ${selectedImages.length} slide${selectedImages.length !== 1 ? 's' : ''} with original layout preserved.`,
      })
    } catch (err) {
      setStatus({ type: 'error', message: 'Context export failed: ' + (err?.message || 'Unknown error') })
    } finally {
      setIsExportingContext(false)
    }
  }

  const handleReset = () => {
    setImages([])
    setSelected(new Set())
    setOriginalFile(null)
    setStatus({ type: 'idle', message: '' })
  }

  // Drag-to-reorder
  const handleDragStart = (index) => { setDragIndex(index); setDragOverIndex(null) }
  const handleDragOver = (index) => {
    if (dragIndex === null || dragIndex === index) return
    setDragOverIndex(index)
  }
  const handleDrop = (index) => {
    if (dragIndex === null || dragIndex === index) return
    setImages((prev) => reorder(prev, dragIndex, index))
    setDragIndex(null)
    setDragOverIndex(null)
  }
  const handleDragEnd = () => { setDragIndex(null); setDragOverIndex(null) }

  return (
    <div className="extractor">
      {images.length === 0 ? (
        <UploadZone onFile={handleFile} isLoading={isLoading} />
      ) : (
        <div className="extractor__toolbar">
          <div className="extractor__toolbar-left">
            <span className="extractor__count">
              {selectedImages.length} / {images.length} selected
            </span>
            <button type="button" className="ghost" onClick={selectAll} disabled={isBusy}>Select All</button>
            <button type="button" className="ghost" onClick={deselectAll} disabled={isBusy}>Deselect All</button>
            <button type="button" className="ghost" onClick={handleReset} disabled={isBusy}>Upload New</button>
          </div>
          <div className="extractor__toolbar-right">
            {/* JPEG export */}
            <button
              type="button"
              className="button button--jpeg"
              onClick={handleExportJpeg}
              disabled={!canExport}
              title="Download selected images as JPEG files inside a ZIP"
            >
              {isExportingJpeg
                ? 'Converting…'
                : `Export JPEG${selectedImages.length !== 1 ? 's' : ''} (${selectedImages.length})`}
            </button>
            {/* PPTX — images only */}
            <button
              type="button"
              className="button"
              onClick={handleExportPptx}
              disabled={!canExport}
              title="Download selected images as a PPTX, one image per slide (images only)"
            >
              {isExportingPptx
                ? 'Building PPTX…'
                : `Export PPTX (${selectedImages.length})`}
            </button>
            {/* PPTX — with original slide context */}
            <button
              type="button"
              className="button button--context"
              onClick={handleExportWithContext}
              disabled={!canExport || !originalFile}
              title="Each image gets its own slide cloned from the original — text boxes, shapes and backgrounds are preserved"
            >
              {isExportingContext
                ? 'Building slides…'
                : `Export PPTX + Context (${selectedImages.length})`}
            </button>
          </div>
        </div>
      )}

      {status.message && (
        <p className={`app__note app__note--${status.type}`}>{status.message}</p>
      )}

      {images.length > 0 && (
        <>
          <p className="extractor__hint">
            Click to select/deselect · Drag to reorder · <strong>Export PPTX + Context</strong> clones the original slide for each image (text, shapes &amp; layout preserved)
          </p>
          <div className="extractor__grid">
            {images.map((img, index) => (
              <ImageCard
                key={img.id}
                item={img}
                isSelected={selected.has(img.id)}
                onToggle={() => toggleSelect(img.id)}
                onDragStart={() => handleDragStart(index)}
                onDragOver={() => handleDragOver(index)}
                onDrop={() => handleDrop(index)}
                onDragEnd={handleDragEnd}
                isDragging={dragIndex === index}
                isDragOver={dragOverIndex === index}
              />
            ))}
          </div>
        </>
      )}
    </div>
  )
}
