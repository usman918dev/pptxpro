import React, { useState, useRef } from 'react'
import { PDFDocument, PDFName, PDFString } from 'pdf-lib'
import createWorker from 'tesseract.js/src/createWorker'
import { extractAllImages } from './report/importPptx'

// ── GPS Coordinate Parsing Engine ──────────────────────────────────────────────

export function extractGpsCoordinates(ocrText) {
  if (!ocrText || typeof ocrText !== 'string') return null

  // Standardize common OCR misreads and symbols
  const cleaned = ocrText
    .replace(/[|]/g, ' ')
    .replace(/°/g, ' ')
    .replace(/,/g, ' , ')

  // 1. Lat 32.028175 Long 74.392718 (or Lat: ... Long: ...)
  const latLongMatch = /Lat[a-z]*[:\s\-=]*(-?\d{1,3}\.\d+)[^\d\-]*Long[a-z]*[:\s\-=]*(-?\d{1,3}\.\d+)/i.exec(cleaned)
  if (latLongMatch) {
    const lat = parseFloat(latLongMatch[1])
    const lng = parseFloat(latLongMatch[2])
    if (isValidLatLng(lat, lng)) {
      return formatCoordsResult(lat, lng, latLongMatch[0])
    }
  }

  // 2. Reverse order: Long 74.392718 Lat 32.028175
  const longLatMatch = /Long[a-z]*[:\s\-=]*(-?\d{1,3}\.\d+)[^\d\-]*Lat[a-z]*[:\s\-=]*(-?\d{1,3}\.\d+)/i.exec(cleaned)
  if (longLatMatch) {
    const lng = parseFloat(longLatMatch[1])
    const lat = parseFloat(longLatMatch[2])
    if (isValidLatLng(lat, lng)) {
      return formatCoordsResult(lat, lng, longLatMatch[0])
    }
  }

  // 3. Fallback: Search for separate Lat and Long matches in the text block
  const latOnly = /Lat[a-z]*[:\s\-=]*(-?\d{1,3}\.\d+)/i.exec(cleaned)
  const longOnly = /Long[a-z]*[:\s\-=]*(-?\d{1,3}\.\d+)/i.exec(cleaned)
  if (latOnly && longOnly) {
    const lat = parseFloat(latOnly[1])
    const lng = parseFloat(longOnly[1])
    if (isValidLatLng(lat, lng)) {
      return formatCoordsResult(lat, lng, `${latOnly[0]} ${longOnly[0]}`)
    }
  }

  // 4. Pair match: 32.028175 , 74.392718
  const pairMatch = /(-?\d{1,2}\.\d{4,8})\s*,\s*(-?\d{1,3}\.\d{4,8})/.exec(cleaned)
  if (pairMatch) {
    const lat = parseFloat(pairMatch[1])
    const lng = parseFloat(pairMatch[2])
    if (isValidLatLng(lat, lng)) {
      return formatCoordsResult(lat, lng, pairMatch[0])
    }
  }

  return null
}

function isValidLatLng(lat, lng) {
  return (
    !isNaN(lat) &&
    !isNaN(lng) &&
    lat >= -90 &&
    lat <= 90 &&
    lng >= -180 &&
    lng <= 180 &&
    (lat !== 0 || lng !== 0)
  )
}

function formatCoordsResult(lat, lng, rawMatch) {
  const latFormatted = lat.toFixed(6)
  const lngFormatted = lng.toFixed(6)
  return {
    lat,
    lng,
    latFormatted,
    lngFormatted,
    display: `Lat ${latFormatted}° Long ${lngFormatted}°`,
    mapUrl: `https://www.google.com/maps?q=${latFormatted},${lngFormatted}`,
    rawMatch,
  }
}

// ── Crop Bottom 30% Canvas Helper ──────────────────────────────────────────────

function cropBottom30Canvas(imgElement) {
  const canvas = document.createElement('canvas')
  const width = imgElement.naturalWidth || imgElement.width || 800
  const height = imgElement.naturalHeight || imgElement.height || 600

  const cropHeight = Math.floor(height * 0.30)
  const startY = height - cropHeight

  canvas.width = width
  canvas.height = cropHeight

  const ctx = canvas.getContext('2d')
  ctx.drawImage(
    imgElement,
    0,
    startY,
    width,
    cropHeight,
    0,
    0,
    width,
    cropHeight
  )

  // Contrast enhancement for OCR
  try {
    const imageData = ctx.getImageData(0, 0, width, cropHeight)
    const data = imageData.data
    for (let i = 0; i < data.length; i += 4) {
      const avg = (data[i] + data[i + 1] + data[i + 2]) / 3
      const val = avg < 130 ? Math.max(0, avg - 30) : Math.min(255, avg + 30)
      data[i] = val
      data[i + 1] = val
      data[i + 2] = val
    }
    ctx.putImageData(imageData, 0, 0)
  } catch {
    // Ignore canvas security taint if any
  }

  return canvas
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.crossOrigin = 'anonymous'
    img.onload = () => resolve(img)
    img.onerror = (err) => reject(err)
    img.src = src
  })
}

// ── PDF Generation with Embedded Google Maps Links ────────────────────────────

export async function generatePdfWithGpsLinks(processedSlides) {
  const pdfDoc = await PDFDocument.create()

  for (const slide of processedSlides) {
    if (!slide.dataUrl) continue

    let embeddedImage
    if (slide.dataUrl.startsWith('data:image/png')) {
      embeddedImage = await pdfDoc.embedPng(slide.dataUrl)
    } else {
      // Default to JPG for data:image/jpeg or URL
      embeddedImage = await pdfDoc.embedJpg(slide.dataUrl)
    }

    const dims = embeddedImage.scale(1)
    const page = pdfDoc.addPage([dims.width, dims.height])

    // Draw full slide image
    page.drawImage(embeddedImage, {
      x: 0,
      y: 0,
      width: dims.width,
      height: dims.height,
    })

    // If GPS coordinates found, embed clickable PDF Link Annotation over bottom 30% area
    if (slide.gps && slide.gps.mapUrl) {
      const bottom30Height = dims.height * 0.30
      // PDF rectangle coordinates: [lower-left x, lower-left y, upper-right x, upper-right y]
      const rect = [0, 0, dims.width, bottom30Height]

      const linkAnnotation = pdfDoc.context.obj({
        Type: 'Annot',
        Subtype: 'Link',
        Rect: rect,
        Border: [0, 0, 0],
        A: {
          Type: 'Action',
          S: 'URI',
          URI: PDFString.of(slide.gps.mapUrl),
        },
      })

      const linkRef = pdfDoc.context.register(linkAnnotation)

      let annots = page.node.get(PDFName.of('Annots'))
      if (!annots) {
        annots = pdfDoc.context.obj([])
        page.node.set(PDFName.of('Annots'), annots)
      }
      annots.push(linkRef)
    }
  }

  const pdfBytes = await pdfDoc.save()
  return new Blob([pdfBytes], { type: 'application/pdf' })
}

// ── Main React Component ───────────────────────────────────────────────────────

export function PptxToPdfOcr() {
  const [file, setFile] = useState(null)
  const [isProcessing, setIsProcessing] = useState(false)
  const [progressMsg, setProgressMsg] = useState('')
  const [progressPct, setProgressPct] = useState(0)
  const [processedSlides, setProcessedSlides] = useState([])
  const [pdfBlob, setPdfBlob] = useState(null)
  const [pdfFileName, setPdfFileName] = useState('Converted_GPS_Report.pdf')

  const fileInputRef = useRef(null)

  const handleFileSelect = (selectedFile) => {
    if (!selectedFile) return
    setFile(selectedFile)
    setProcessedSlides([])
    setPdfBlob(null)
    const baseName = selectedFile.name.replace(/\.pptx$/i, '')
    setPdfFileName(`${baseName}_GPS_Report.pdf`)
  }

  const startConversion = async () => {
    if (!file) return

    setIsProcessing(true)
    setProgressMsg('Extracting slide images from PPTX...')
    setProgressPct(5)

    let worker = null

    try {
      // Step 1: Extract all images/slides from PPTX
      const rawImages = await extractAllImages(file)

      if (!rawImages || rawImages.length === 0) {
        alert('❌ No images or slides found in the PPTX file.')
        setIsProcessing(false)
        return
      }

      setProgressMsg(`Initializing OCR engine... (${rawImages.length} images found)`)
      setProgressPct(15)

      // Step 2: Initialize Tesseract.js worker
      worker = await createWorker('eng')

      const results = []

      // Step 3: Process each image (Crop bottom 30% & Run OCR)
      for (let i = 0; i < rawImages.length; i++) {
        const item = rawImages[i]
        const currentNum = i + 1
        const total = rawImages.length

        setProgressMsg(`OCR scanning bottom 30% of slide ${currentNum} of ${total}...`)
        const pct = Math.floor(15 + (i / total) * 65)
        setProgressPct(pct)

        try {
          const imgEl = await loadImage(item.dataUrl)
          const croppedCanvas = cropBottom30Canvas(imgEl)
          const croppedDataUrl = croppedCanvas.toDataURL('image/png')

          // Run OCR on cropped bottom 30% canvas
          const { data: { text } } = await worker.recognize(croppedCanvas)
          const gps = extractGpsCoordinates(text)

          results.push({
            ...item,
            croppedDataUrl,
            ocrText: text,
            gps,
          })
        } catch (err) {
          console.warn(`Failed OCR on slide ${currentNum}`, err)
          results.push({
            ...item,
            croppedDataUrl: '',
            ocrText: '',
            gps: null,
          })
        }
      }

      // Step 4: Build High-Quality PDF with Google Maps Links
      setProgressMsg('Generating high-quality PDF with embedded Google Maps links...')
      setProgressPct(90)

      const generatedBlob = await generatePdfWithGpsLinks(results)

      setProcessedSlides(results)
      setPdfBlob(generatedBlob)
      setProgressPct(100)
      setProgressMsg('✅ Complete! PDF with GPS Google Maps links ready for download.')
    } catch (err) {
      alert(`❌ Error during processing: ${err.message}`)
      console.error(err)
    } finally {
      if (worker) {
        await worker.terminate()
      }
      setIsProcessing(false)
    }
  }

  const handleDownloadPdf = () => {
    if (!pdfBlob) return
    const url = URL.createObjectURL(pdfBlob)
    const a = document.createElement('a')
    a.href = url
    a.download = pdfFileName
    a.click()
    URL.revokeObjectURL(url)
  }

  const handleOpenAllMaps = () => {
    const slidesWithGps = processedSlides.filter((s) => s.gps && s.gps.mapUrl)
    if (slidesWithGps.length === 0) {
      alert('No GPS coordinates detected to open.')
      return
    }
    slidesWithGps.forEach((s) => {
      window.open(s.gps.mapUrl, '_blank', 'noopener,noreferrer')
    })
  }

  const handleCopyAllCoords = () => {
    const lines = processedSlides
      .filter((s) => s.gps)
      .map((s, idx) => `Slide ${idx + 1}: ${s.gps.display} (${s.gps.mapUrl})`)

    if (lines.length === 0) {
      alert('No coordinates to copy.')
      return
    }

    navigator.clipboard.writeText(lines.join('\n'))
    alert(`📋 Copied ${lines.length} coordinate records to clipboard!`)
  }

  const gpsCount = processedSlides.filter((s) => s.gps).length

  return (
    <div className="gps-pdf-tool">
      {/* ── Control Header Panel ── */}
      <div className="card gps-pdf-tool__header">
        <div>
          <h2 style={{ margin: 0, fontSize: '18px', fontWeight: 700 }}>
            PPTX to PDF (Bottom 30% GPS OCR)
          </h2>
          <p style={{ margin: '4px 0 0', fontSize: '13px', color: 'var(--muted)' }}>
            Upload any PPTX presentation. Scans the <strong>bottom 30%</strong> region of each slide for GPS coordinates
            and builds a high-quality PDF with embedded Google Maps links.
          </p>
        </div>

        <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', alignItems: 'center' }}>
          <input
            ref={fileInputRef}
            type="file"
            accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
            style={{ display: 'none' }}
            onChange={(e) => handleFileSelect(e.target.files?.[0])}
          />

          <button
            type="button"
            className="button button--secondary"
            onClick={() => fileInputRef.current?.click()}
            disabled={isProcessing}
          >
            {file ? '📁 Change PPTX' : '📁 Choose PPTX File'}
          </button>

          {file && (
            <button
              type="button"
              className="button"
              onClick={startConversion}
              disabled={isProcessing}
            >
              {isProcessing ? '⚡ Processing...' : '🚀 Convert to PDF with OCR Links'}
            </button>
          )}

          {pdfBlob && (
            <button
              type="button"
              className="button button--completed"
              onClick={handleDownloadPdf}
            >
              📥 Download PDF ({gpsCount} Map Links)
            </button>
          )}
        </div>
      </div>

      {/* ── Active File & Progress Bar ── */}
      {file && (
        <div className="card" style={{ marginTop: '16px', padding: '16px 20px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
            <span style={{ fontWeight: 600, fontSize: '14px' }}>
              📄 Selected: <span style={{ color: 'var(--accent)' }}>{file.name}</span>
            </span>
            <span style={{ fontSize: '12px', color: 'var(--muted)' }}>
              {(file.size / (1024 * 1024)).toFixed(2)} MB
            </span>
          </div>

          {isProcessing && (
            <div>
              <div className="gps-pdf-tool__progress-bg">
                <div
                  className="gps-pdf-tool__progress-fill"
                  style={{ width: `${progressPct}%` }}
                />
              </div>
              <p style={{ margin: '8px 0 0', fontSize: '12px', color: 'var(--muted)', fontWeight: 500 }}>
                {progressMsg} ({progressPct}%)
              </p>
            </div>
          )}
        </div>
      )}

      {/* ── Results Dashboard Header ── */}
      {processedSlides.length > 0 && (
        <div className="gps-pdf-tool__summary card">
          <div className="gps-pdf-tool__badges">
            <span className="gps-pdf-badge gps-pdf-badge--total">
              📊 Total Slides: <strong>{processedSlides.length}</strong>
            </span>
            <span className="gps-pdf-badge gps-pdf-badge--gps">
              📍 GPS Detected: <strong>{gpsCount}</strong>
            </span>
          </div>

          <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
            <button
              type="button"
              className="ghost"
              style={{ fontSize: '12px', padding: '6px 14px' }}
              onClick={handleOpenAllMaps}
              disabled={gpsCount === 0}
            >
              🌐 Open All Maps Links in Tabs
            </button>

            <button
              type="button"
              className="ghost"
              style={{ fontSize: '12px', padding: '6px 14px' }}
              onClick={handleCopyAllCoords}
              disabled={gpsCount === 0}
            >
              📋 Copy Coordinates List
            </button>

            <button
              type="button"
              className="button"
              style={{ fontSize: '12px', padding: '6px 18px' }}
              onClick={handleDownloadPdf}
            >
              📥 Download PDF Document
            </button>
          </div>
        </div>
      )}

      {/* ── Slide Cards Grid ── */}
      {processedSlides.length > 0 && (
        <div className="gps-pdf-grid">
          {processedSlides.map((slide, idx) => (
            <div key={`slide-${idx}`} className={`gps-slide-card${slide.gps ? ' has-gps' : ''}`}>
              <div className="gps-slide-card__header">
                <span>Slide {idx + 1}</span>
                {slide.gps ? (
                  <span className="gps-status-badge is-found">📍 Coordinates Found</span>
                ) : (
                  <span className="gps-status-badge is-none">No GPS Text</span>
                )}
              </div>

              <div className="gps-slide-card__media">
                <img src={slide.dataUrl} alt={`Slide ${idx + 1}`} className="gps-slide-card__img" />
                {slide.croppedDataUrl && (
                  <div className="gps-slide-card__crop-preview" title="Cropped Bottom 30% OCR Region">
                    <img src={slide.croppedDataUrl} alt="Bottom 30% Region" />
                    <span className="gps-slide-card__crop-tag">Scanned Bottom 30%</span>
                  </div>
                )}
              </div>

              <div className="gps-slide-card__body">
                {slide.gps ? (
                  <div>
                    <p className="gps-coords-text">{slide.gps.display}</p>
                    <a
                      href={slide.gps.mapUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="button button--map-link"
                    >
                      🗺️ Open in Google Maps ↗
                    </a>
                  </div>
                ) : (
                  <p style={{ margin: 0, fontSize: '12px', color: 'var(--muted)', fontStyle: 'italic' }}>
                    No coordinate patterns (Lat/Long) detected in the bottom 30% overlay.
                  </p>
                )}

                {slide.ocrText && (
                  <details style={{ marginTop: '10px' }}>
                    <summary style={{ fontSize: '11px', color: 'var(--muted)', cursor: 'pointer' }}>
                      View Raw Bottom 30% OCR Snippet
                    </summary>
                    <pre className="gps-ocr-snippet">{slide.ocrText.trim()}</pre>
                  </details>
                )}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}
