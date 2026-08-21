import { useRef, useState, useCallback } from 'react'
import { PDFDocument } from 'pdf-lib'
import { extractAllImages } from './report/importPptx'

// ── Helper: render all slide images into a downloadable PDF ──────────────────

async function slidesToPdf(slides) {
  const pdfDoc = await PDFDocument.create()

  for (const slide of slides) {
    if (!slide.dataUrl) continue

    let embeddedImage
    if (slide.dataUrl.startsWith('data:image/png')) {
      embeddedImage = await pdfDoc.embedPng(slide.dataUrl)
    } else {
      embeddedImage = await pdfDoc.embedJpg(slide.dataUrl)
    }

    const dims = embeddedImage.scale(1)
    // Standard widescreen 16:9 → 960×540 pt (13.33in × 7.5in @72dpi)
    // but use actual slide dims for best fidelity
    const page = pdfDoc.addPage([dims.width, dims.height])
    page.drawImage(embeddedImage, {
      x: 0,
      y: 0,
      width: dims.width,
      height: dims.height,
    })
  }

  const pdfBytes = await pdfDoc.save()
  return new Blob([pdfBytes], { type: 'application/pdf' })
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = fileName
  a.click()
  URL.revokeObjectURL(url)
}

// ── Upload drop zone ─────────────────────────────────────────────────────────

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
          <div className="extractor-spinner" aria-label="Rendering slides…" />
          <p className="extractor-upload__label">Rendering slide images…</p>
        </div>
      ) : (
        <div className="extractor-upload__inner">
          <span className="extractor-upload__icon" aria-hidden="true">📄</span>
          <p className="extractor-upload__label">Drop your PPTX here or click to browse</p>
          <p className="extractor-upload__hint">Every slide will be rendered and packed into a PDF</p>
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

// ── Slide card ───────────────────────────────────────────────────────────────

function SlideCard({ item, isSelected, onToggle, onDragStart, onDragOver, onDrop, onDragEnd, isDragging, isDragOver }) {
  return (
    <div
      className={`extractor-card${isSelected ? ' is-selected' : ''}${isDragging ? ' is-dragging' : ''}${isDragOver ? ' is-drag-over' : ''}`}
      draggable
      onDragStart={onDragStart}
      onDragOver={(e) => { e.preventDefault(); onDragOver() }}
      onDrop={(e) => { e.preventDefault(); onDrop() }}
      onDragEnd={onDragEnd}
      onClick={onToggle}
      title={`Slide ${item.slideNumber} — click to ${isSelected ? 'deselect' : 'select'}`}
    >
      <div className="extractor-card__img-wrap">
        <img
          src={item.dataUrl}
          alt={`Slide ${item.slideNumber}`}
          className="extractor-card__img"
          draggable={false}
        />
        <div className={`extractor-card__check${isSelected ? ' is-checked' : ''}`} aria-hidden="true">
          {isSelected ? '✓' : ''}
        </div>
      </div>
      <div className="extractor-card__footer">
        <span className="extractor-card__label">Slide {item.slideNumber}</span>
        <span className="extractor-card__drag-hint" aria-hidden="true">⠿</span>
      </div>
    </div>
  )
}

// ── Main component ────────────────────────────────────────────────────────────

export function PptxToPdf() {
  const [slides, setSlides] = useState([])
  const [selected, setSelected] = useState(new Set())
  const [isLoading, setIsLoading] = useState(false)
  const [isConverting, setIsConverting] = useState(false)
  const [status, setStatus] = useState({ type: 'idle', message: '' })
  const [pdfFileName, setPdfFileName] = useState('')
  const [dragIndex, setDragIndex] = useState(null)
  const [dragOverIndex, setDragOverIndex] = useState(null)

  const isBusy = isLoading || isConverting

  const handleFile = useCallback(async (file) => {
    setIsLoading(true)
    setStatus({ type: 'working', message: 'Extracting slides from PPTX…' })
    setSlides([])
    setSelected(new Set())
    setPdfFileName(file.name.replace(/\.pptx$/i, ''))

    try {
      const extracted = await extractAllImages(file)

      if (!extracted || extracted.length === 0) {
        setStatus({ type: 'error', message: 'No slides found in this PPTX file.' })
        return
      }

      // De-duplicate: keep only one image per slide (the first / full-slide image)
      const seen = new Set()
      const deduplicated = []
      for (const img of extracted) {
        if (!seen.has(img.slideNumber)) {
          seen.add(img.slideNumber)
          deduplicated.push(img)
        }
      }

      setSlides(deduplicated)
      setSelected(new Set(deduplicated.map((s) => s.id)))
      setStatus({
        type: 'success',
        message: `Loaded ${deduplicated.length} slide${deduplicated.length !== 1 ? 's' : ''} from "${file.name}".`,
      })
    } catch (err) {
      console.error(err)
      setStatus({ type: 'error', message: `Failed to extract slides: ${err?.message || 'Unknown error'}` })
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

  const selectAll = () => setSelected(new Set(slides.map((s) => s.id)))
  const deselectAll = () => setSelected(new Set())
  const handleReset = () => {
    setSlides([])
    setSelected(new Set())
    setPdfFileName('')
    setStatus({ type: 'idle', message: '' })
  }

  const selectedSlides = slides.filter((s) => selected.has(s.id))

  const handleConvert = async () => {
    if (selectedSlides.length === 0) {
      alert('Please select at least one slide to convert.')
      return
    }

    try {
      setIsConverting(true)
      setStatus({ type: 'working', message: `Building PDF from ${selectedSlides.length} slide${selectedSlides.length !== 1 ? 's' : ''}…` })

      const blob = await slidesToPdf(selectedSlides)
      const outName = `${pdfFileName || 'Converted'}.pdf`
      downloadBlob(blob, outName)

      setStatus({
        type: 'success',
        message: `✅ Downloaded "${outName}" (${selectedSlides.length} page${selectedSlides.length !== 1 ? 's' : ''}).`,
      })
    } catch (err) {
      console.error(err)
      setStatus({ type: 'error', message: `PDF generation failed: ${err?.message || 'Unknown error'}` })
    } finally {
      setIsConverting(false)
    }
  }

  // Drag-to-reorder
  const handleDragStart = (index) => { setDragIndex(index); setDragOverIndex(null) }
  const handleDragOver = (index) => {
    if (dragIndex === null || dragIndex === index) return
    setDragOverIndex(index)
  }
  const handleDrop = (index) => {
    if (dragIndex === null || dragIndex === index) return
    setSlides((prev) => {
      const next = [...prev]
      const [moved] = next.splice(dragIndex, 1)
      next.splice(index, 0, moved)
      return next
    })
    setDragIndex(null)
    setDragOverIndex(null)
  }
  const handleDragEnd = () => { setDragIndex(null); setDragOverIndex(null) }

  const canConvert = selectedSlides.length > 0 && !isBusy

  return (
    <div className="extractor">
      {slides.length === 0 ? (
        <UploadZone onFile={handleFile} isLoading={isLoading} />
      ) : (
        <>
          <div className="extractor__toolbar">
            <div className="extractor__toolbar-left">
              <span className="extractor__count">
                {selectedSlides.length} / {slides.length} selected
              </span>
              <button type="button" className="ghost" onClick={selectAll} disabled={isBusy}>Select All</button>
              <button type="button" className="ghost" onClick={deselectAll} disabled={isBusy}>Deselect All</button>
              <button type="button" className="ghost" onClick={handleReset} disabled={isBusy}>Upload New</button>
            </div>
            <div className="extractor__toolbar-right">
              <button
                type="button"
                className="button"
                onClick={handleConvert}
                disabled={!canConvert}
              >
                {isConverting
                  ? '⏳ Building PDF…'
                  : `📥 Convert to PDF (${selectedSlides.length} slide${selectedSlides.length !== 1 ? 's' : ''})`}
              </button>
            </div>
          </div>

          {status.message && (
            <p className={`app__note app__note--${status.type}`}>{status.message}</p>
          )}

          <p className="extractor__hint">
            Click to select/deselect slides · Drag ⠿ to reorder · Only selected slides will be in the PDF
          </p>

          <div className="extractor__grid">
            {slides.map((slide, index) => (
              <SlideCard
                key={slide.id}
                item={slide}
                isSelected={selected.has(slide.id)}
                onToggle={() => toggleSelect(slide.id)}
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

      {slides.length === 0 && status.message && (
        <p className={`app__note app__note--${status.type}`}>{status.message}</p>
      )}
    </div>
  )
}
