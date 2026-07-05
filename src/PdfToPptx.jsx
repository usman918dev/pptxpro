import { useRef, useState, useCallback } from 'react'
import PptxGenJS from 'pptxgenjs'

// Helper to format file size in a human readable way
const formatSize = (bytes) => {
  if (bytes === 0) return '0 B'
  const k = 1024
  const sizes = ['B', 'KB', 'MB', 'GB']
  const i = Math.floor(Math.log(bytes) / Math.log(k))
  return `${parseFloat((bytes / Math.pow(k, i)).toFixed(2))} ${sizes[i]}`
}

function UploadZone({ onFile, isLoading }) {
  const inputRef = useRef(null)
  const [isDragging, setIsDragging] = useState(false)

  const handleFile = (file) => {
    if (!file) return
    if (!file.name.toLowerCase().endsWith('.pdf')) {
      alert('Please upload a valid .pdf file.')
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
          <div className="extractor-spinner" aria-label="Processing PDF..." />
          <p className="extractor-upload__label">Rendering PDF pages...</p>
        </div>
      ) : (
        <div className="extractor-upload__inner">
          <span className="extractor-upload__icon" aria-hidden="true">📄</span>
          <p className="extractor-upload__label">Drop your PDF here or click to browse</p>
          <p className="extractor-upload__hint">Each page of the PDF will be rendered into a PPTX slide</p>
        </div>
      )}
      <input
        ref={inputRef}
        type="file"
        accept=".pdf,application/pdf"
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

function PageCard({ item, isSelected, onToggle, onDragStart, onDragOver, onDrop, onDragEnd, isDragging, isDragOver }) {
  return (
    <div
      className={`extractor-card${isSelected ? ' is-selected' : ''}${isDragging ? ' is-dragging' : ''}${isDragOver ? ' is-drag-over' : ''}`}
      draggable
      onDragStart={onDragStart}
      onDragOver={(e) => { e.preventDefault(); onDragOver() }}
      onDrop={(e) => { e.preventDefault(); onDrop() }}
      onDragEnd={onDragEnd}
      onClick={onToggle}
      title={`Page ${item.pageNumber} — click to ${isSelected ? 'deselect' : 'select'}`}
    >
      <div className="extractor-card__img-wrap">
        <img
          src={item.dataUrl}
          alt={`Page ${item.pageNumber} preview`}
          className="extractor-card__img"
          draggable={false}
        />
        <div className={`extractor-card__check${isSelected ? ' is-checked' : ''}`} aria-hidden="true">
          {isSelected ? '✓' : ''}
        </div>
      </div>
      <div className="extractor-card__footer">
        <span className="extractor-card__label">
          Page {item.pageNumber}
          <span style={{ fontSize: '11px', color: 'var(--muted)', marginLeft: '6px', fontWeight: 'normal' }}>
            ({item.wInches.toFixed(1)}" × {item.hInches.toFixed(1)}")
          </span>
        </span>
        <span className="extractor-card__drag-hint" aria-hidden="true">⠿</span>
      </div>
    </div>
  )
}

export function PdfToPptx() {
  const [pages, setPages] = useState([])
  const [isLoading, setIsLoading] = useState(false)
  const [isExporting, setIsExporting] = useState(false)
  const [status, setStatus] = useState({ type: 'idle', message: '' })
  const [pdfFileName, setPdfFileName] = useState('')
  const [pdfFileSize, setPdfFileSize] = useState(0)
  const [layoutMode, setLayoutMode] = useState('auto') // 'auto' | '16_9' | '3_4'
  const [dragIndex, setDragIndex] = useState(null)
  const [dragOverIndex, setDragOverIndex] = useState(null)

  const isBusy = isLoading || isExporting

  const handleFile = useCallback(async (file) => {
    setIsLoading(true)
    setStatus({ type: 'working', message: 'Opening PDF file...' })
    setPages([])
    setPdfFileName(file.name.replace(/\.pdf$/i, ''))
    setPdfFileSize(file.size)

    try {
      const arrayBuffer = await file.arrayBuffer()
      const typedarray = new Uint8Array(arrayBuffer)
      
      if (!window.pdfjsLib) {
        throw new Error('PDF.js library is not available. Please verify index.html CDN scripts.')
      }
      
      const pdf = await window.pdfjsLib.getDocument({ data: typedarray }).promise
      const parsedPages = []

      for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
        setStatus({
          type: 'working',
          message: `Rendering page ${pageNum} of ${pdf.numPages}...`
        })
        const page = await pdf.getPage(pageNum)
        
        // Render PDF page to canvas at high DPI (2.0 scale)
        const scale = 2.0
        const viewport = page.getViewport({ scale })
        
        const canvas = document.createElement('canvas')
        canvas.width = viewport.width
        canvas.height = viewport.height
        const context = canvas.getContext('2d')
        
        await page.render({
          canvasContext: context,
          viewport: viewport
        }).promise
        
        const dataUrl = canvas.toDataURL('image/jpeg', 0.90)
        
        // Get dimensions in inches at 1.0 scale (72 points = 1 inch)
        const baseViewport = page.getViewport({ scale: 1.0 })
        const wInches = baseViewport.width / 72
        const hInches = baseViewport.height / 72

        parsedPages.push({
          id: `page-${pageNum}-${Date.now()}-${Math.random().toString(36).substring(2, 7)}`,
          pageNumber: pageNum,
          dataUrl,
          width: viewport.width,
          height: viewport.height,
          wInches,
          hInches,
          isSelected: true
        })
      }

      setPages(parsedPages)
      setStatus({
        type: 'success',
        message: `Successfully loaded ${parsedPages.length} pages from "${file.name}".`
      })
    } catch (err) {
      console.error(err)
      setStatus({
        type: 'error',
        message: `PDF rendering failed: ${err?.message || 'Unknown error'}`
      })
    } finally {
      setIsLoading(false)
    }
  }, [])

  const togglePageSelection = useCallback((id) => {
    setPages((prev) =>
      prev.map((page) =>
        page.id === id ? { ...page, isSelected: !page.isSelected } : page
      )
    )
  }, [])

  const selectAll = () => {
    setPages((prev) => prev.map((page) => ({ ...page, isSelected: true })))
  }

  const deselectAll = () => {
    setPages((prev) => prev.map((page) => ({ ...page, isSelected: false })))
  }

  const handleReset = () => {
    setPages([])
    setPdfFileName('')
    setPdfFileSize(0)
    setStatus({ type: 'idle', message: '' })
  }

  const handleExport = async () => {
    const selectedPages = pages.filter((page) => page.isSelected)
    if (selectedPages.length === 0) {
      alert('Please select at least one page to export.')
      return
    }

    try {
      setIsExporting(true)
      setStatus({ type: 'working', message: 'Generating PowerPoint slides...' })

      const pptx = new PptxGenJS()
      const firstPage = selectedPages[0]

      let slideW = 13.333
      let slideH = 7.5

      if (layoutMode === 'auto' && firstPage) {
        slideW = firstPage.wInches
        slideH = firstPage.hInches
        pptx.defineLayout({ name: 'PDF_AUTO', width: slideW, height: slideH })
        pptx.layout = 'PDF_AUTO'
      } else if (layoutMode === '16_9') {
        pptx.layout = 'LAYOUT_WIDE'
      } else if (layoutMode === '3_4') {
        pptx.defineLayout({ name: 'PORTRAIT_3_4', width: 7.5, height: 10.0 })
        pptx.layout = 'PORTRAIT_3_4'
        slideW = 7.5
        slideH = 10.0
      }

      for (const img of selectedPages) {
        const slide = pptx.addSlide()
        slide.background = { fill: 'FFFFFF' }
        
        slide.addImage({
          data: img.dataUrl,
          x: 0,
          y: 0,
          w: slideW,
          h: slideH,
          sizing: { type: 'contain', w: slideW, h: slideH }
        })
      }

      const outputFileName = `${pdfFileName || 'Converted_PDF'}.pptx`
      await pptx.writeFile({ fileName: outputFileName })

      setStatus({
        type: 'success',
        message: `Successfully generated and downloaded "${outputFileName}"!`
      })
    } catch (err) {
      console.error(err)
      setStatus({
        type: 'error',
        message: `PowerPoint generation failed: ${err?.message || 'Unknown error'}`
      })
    } finally {
      setIsExporting(false)
    }
  }

  // Drag to reorder
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
    setPages((prev) => {
      const next = [...prev]
      const [moved] = next.splice(dragIndex, 1)
      next.splice(index, 0, moved)
      return next
    })
    setDragIndex(null)
    setDragOverIndex(null)
  }

  const handleDragEnd = () => {
    setDragIndex(null)
    setDragOverIndex(null)
  }

  const selectedPages = pages.filter((p) => p.isSelected)
  const canExport = selectedPages.length > 0 && !isBusy

  return (
    <div className="extractor">
      {pages.length === 0 ? (
        <UploadZone onFile={handleFile} isLoading={isLoading} />
      ) : (
        <>
          <div className="extractor__toolbar">
            <div className="extractor__toolbar-left">
              <span className="extractor__count">
                {selectedPages.length} / {pages.length} selected
              </span>
              <button type="button" className="ghost" onClick={selectAll} disabled={isBusy}>Select All</button>
              <button type="button" className="ghost" onClick={deselectAll} disabled={isBusy}>Deselect All</button>
              <button type="button" className="ghost" onClick={handleReset} disabled={isBusy}>Upload New</button>
            </div>
            
            <div className="extractor__toolbar-left" style={{ gap: '8px' }}>
              <span className="extractor__count" style={{ background: 'none', border: 'none', paddingLeft: 0 }}>Layout:</span>
              <div className="app__subnav" style={{ margin: 0 }}>
                <button
                  type="button"
                  className={`ghost${layoutMode === 'auto' ? ' is-active' : ''}`}
                  onClick={() => setLayoutMode('auto')}
                  title="Preserve original page proportions"
                >
                  Auto
                </button>
                <button
                  type="button"
                  className={`ghost${layoutMode === '16_9' ? ' is-active' : ''}`}
                  onClick={() => setLayoutMode('16_9')}
                  title="Fit standard landscape widescreen"
                >
                  Widescreen
                </button>
                <button
                  type="button"
                  className={`ghost${layoutMode === '3_4' ? ' is-active' : ''}`}
                  onClick={() => setLayoutMode('3_4')}
                  title="Fit standard portrait layout"
                >
                  Portrait
                </button>
              </div>
            </div>

            <div className="extractor__toolbar-right">
              <button
                type="button"
                className="button"
                onClick={handleExport}
                disabled={!canExport}
              >
                {isExporting ? 'Building Slide Deck...' : 'Convert to PPTX'}
              </button>
            </div>
          </div>

          {status.message && (
            <p className={`app__note app__note--${status.type}`}>{status.message}</p>
          )}

          <p className="extractor__hint">
            Click page cards to select/deselect slides · Drag cards ⠿ to change slide order
          </p>

          <div className="extractor__grid">
            {pages.map((item, index) => (
              <PageCard
                key={item.id}
                item={item}
                isSelected={item.isSelected}
                onToggle={() => togglePageSelection(item.id)}
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
