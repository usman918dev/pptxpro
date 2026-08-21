import React, { useState, useRef } from 'react'
import { PDFDocument, PDFName, PDFString, rgb } from 'pdf-lib'
import createWorker from 'tesseract.js/src/createWorker'
import { extractAllImages } from './report/importPptx'

// ── GPS Coordinate Parsing Engine ──────────────────────────────────────────────

export function extractGpsCoordinates(ocrText) {
  if (!ocrText || typeof ocrText !== 'string') return null

  // 1. Regex to find Latitude labels
  const latLabelRegex = /\b(lat[a-z]*|ot|at|lt|1at|tat|l@t|la1|l1t|ldt|lct|la)\b/gi
  // 2. Regex to find Longitude labels
  const lngLabelRegex = /\b(long[a-z]*|lng|l0ng|1ong|tong|lon|10ng|l\.ong|l\s*o\s*n\s*g)\b/gi

  // Find all matches for lat labels
  const latMatches = []
  let match
  while ((match = latLabelRegex.exec(ocrText)) !== null) {
    latMatches.push({
      index: match.index,
      text: match[0],
      endIndex: match.index + match[0].length,
    })
  }

  // Find all matches for lng labels
  const lngMatches = []
  while ((match = lngLabelRegex.exec(ocrText)) !== null) {
    lngMatches.push({
      index: match.index,
      text: match[0],
      endIndex: match.index + match[0].length,
    })
  }

  // Helper to extract the first number-like string following an index
  function extractNumberAfter(index) {
    const sub = ocrText.slice(index)
    // Matches a number pattern: optional minus, followed by digits, spaces, dots, commas
    // We exclude '-' from the leading separators class so it is captured by the negative group
    const numRegex = /^\s*[:\s=]*(-?\s*\d+(?:[\s.,\-/\\]+\d+)*)/i
    const m = numRegex.exec(sub)
    if (m) {
      return {
        raw: m[1],
        fullMatch: m[0],
        endIndex: index + m.index + m[0].length,
      }
    }
    return null
  }

  // Helper to clean coordinate string
  function cleanCoordinateNumber(numStr, isLatitude) {
    if (!numStr) return null
    const trimmed = numStr.trim()
    // Check if it starts with negative sign (possibly with spaces after it, e.g. "- 31.9")
    const isNegative = trimmed.startsWith('-')
    const digitsOnlyStr = trimmed.replace(/^-/, '').trim()

    // Split by any non-digit character (including space, dots, commas, etc.)
    const groups = digitsOnlyStr.split(/[^0-9]+/).filter(Boolean)
    if (groups.length === 0) return null

    let parsedVal = ''
    if (groups.length >= 2) {
      // We have a natural separator (like dot, space, comma)
      const integerPart = groups[0]
      const decimalPart = groups.slice(1).join('')
      parsedVal = integerPart + '.' + decimalPart
    } else {
      // No separator found (e.g. 741857 or 31996471)
      const digits = groups[0]
      if (isLatitude) {
        // Latitude check
        if (digits.startsWith('0')) {
          parsedVal = '0.' + digits.slice(1)
        } else {
          const firstTwo = parseInt(digits.slice(0, 2), 10)
          if (firstTwo <= 90) {
            parsedVal = digits.slice(0, 2) + '.' + digits.slice(2)
          } else {
            parsedVal = digits.slice(0, 1) + '.' + digits.slice(1)
          }
        }
      } else {
        // Longitude check
        if (digits.startsWith('0')) {
          parsedVal = '0.' + digits.slice(1)
        } else {
          const firstThree = parseInt(digits.slice(0, 3), 10)
          const firstTwo = parseInt(digits.slice(0, 2), 10)
          if (firstThree <= 180) {
            parsedVal = digits.slice(0, 3) + '.' + digits.slice(3)
          } else if (firstTwo <= 99) {
            parsedVal = digits.slice(0, 2) + '.' + digits.slice(2)
          } else {
            parsedVal = digits.slice(0, 1) + '.' + digits.slice(1)
          }
        }
      }
    }

    const finalNum = parseFloat(parsedVal) * (isNegative ? -1 : 1)
    return isNaN(finalNum) ? null : finalNum
  }

  // Now, try to find a pair of Lat and Lng
  let bestPair = null
  let minDistance = Infinity

  for (const latM of latMatches) {
    const latNumObj = extractNumberAfter(latM.endIndex)
    if (!latNumObj) continue
    const latVal = cleanCoordinateNumber(latNumObj.raw, true)
    if (latVal === null || !isValidLatLng(latVal, 0)) continue

    for (const lngM of lngMatches) {
      const lngNumObj = extractNumberAfter(lngM.endIndex)
      if (!lngNumObj) continue
      const lngVal = cleanCoordinateNumber(lngNumObj.raw, false)
      if (lngVal === null || !isValidLatLng(0, lngVal)) continue

      // Calculate distance between the labels in the text to pair the closest ones
      const distance = Math.abs(latM.index - lngM.index)
      if (distance < minDistance) {
        minDistance = distance
        bestPair = {
          lat: latVal,
          lng: lngVal,
          rawMatch: `LatMatch: "${latM.text} ${latNumObj.raw}", LngMatch: "${lngM.text} ${lngNumObj.raw}"`,
        }
      }
    }
  }

  // Fallback: If we couldn't find labelled pairs, maybe they are just written as a pair:
  // e.g. "32.028175 , 74.392718" or "31 9064420, 74 185587"
  if (!bestPair) {
    // Regex for pair of numbers separated by comma or slash or semicolon
    const pairRegex = /(-?\s*\d+(?:[\s.,\-/\\]+\d+)*)\s*[,;/]\s*(-?\s*\d+(?:[\s.,\-/\\]+\d+)*)/g
    let pairM
    while ((pairM = pairRegex.exec(ocrText)) !== null) {
      const latVal = cleanCoordinateNumber(pairM[1], true)
      const lngVal = cleanCoordinateNumber(pairM[2], false)
      if (latVal !== null && lngVal !== null && isValidLatLng(latVal, lngVal)) {
        bestPair = {
          lat: latVal,
          lng: lngVal,
          rawMatch: pairM[0],
        }
        break // take first valid pair match
      }
    }
  }

  if (bestPair) {
    return formatCoordsResult(bestPair.lat, bestPair.lng, bestPair.rawMatch)
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

// ── Crop Bottom Percentage Canvas Helper ───────────────────────────────────────

function cropBottomPercentageCanvas(imgElement, percentage) {
  const srcWidth = imgElement.naturalWidth || imgElement.width || 800
  const srcHeight = imgElement.naturalHeight || imgElement.height || 600

  const cropHeight = Math.floor(srcHeight * (percentage / 100))
  const startY = srcHeight - cropHeight

  // Work at 2× resolution for better OCR sharpness
  const scale = 2
  const canvas = document.createElement('canvas')
  canvas.width = srcWidth * scale
  canvas.height = cropHeight * scale

  const ctx = canvas.getContext('2d')
  ctx.imageSmoothingEnabled = true
  ctx.imageSmoothingQuality = 'high'
  ctx.drawImage(
    imgElement,
    0,
    startY,
    srcWidth,
    cropHeight,
    0,
    0,
    srcWidth * scale,
    cropHeight * scale
  )

  // Step 1: Strong grayscale + contrast boost
  try {
    const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height)
    const data = imageData.data
    const factor = 2.0 // contrast multiplier
    for (let i = 0; i < data.length; i += 4) {
      const avg = (data[i] * 0.299 + data[i + 1] * 0.587 + data[i + 2] * 0.114)
      const val = Math.min(255, Math.max(0, factor * (avg - 128) + 128))
      data[i] = val
      data[i + 1] = val
      data[i + 2] = val
    }
    ctx.putImageData(imageData, 0, 0)
  } catch {
    // Ignore canvas security taint if any
  }

  // Step 2: Sharpen via convolution kernel
  try {
    const w = canvas.width
    const h = canvas.height
    const src = ctx.getImageData(0, 0, w, h)
    const dst = ctx.createImageData(w, h)
    const sd = src.data
    const dd = dst.data
    // 3×3 sharpen kernel: centre=5, cardinal=-1
    for (let y = 1; y < h - 1; y++) {
      for (let x = 1; x < w - 1; x++) {
        const idx = (y * w + x) * 4
        for (let c = 0; c < 3; c++) {
          const sharpened =
            5 * sd[idx + c] -
            sd[((y - 1) * w + x) * 4 + c] -
            sd[((y + 1) * w + x) * 4 + c] -
            sd[(y * w + (x - 1)) * 4 + c] -
            sd[(y * w + (x + 1)) * 4 + c]
          dd[idx + c] = Math.min(255, Math.max(0, sharpened))
        }
        dd[idx + 3] = 255
      }
    }
    ctx.putImageData(dst, 0, 0)
  } catch {
    // Ignore sharpen failure
  }

  return canvas
}

// Allow only ASCII English letters, digits, space, and the specified symbols
const ALLOWED_CHARS_RE = /[^a-zA-Z0-9 "":;.,/\\+\n\r]/g

function sanitizePdfText(text) {
  if (!text) return ''
  return text
    .replace(ALLOWED_CHARS_RE, ' ')   // strip disallowed chars
    .replace(/ {2,}/g, ' ')           // collapse multiple spaces
    .trim()
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

async function renderPdfPages(file, onProgress) {
  const arrayBuffer = await file.arrayBuffer()
  const typedarray = new Uint8Array(arrayBuffer)
  
  if (!window.pdfjsLib) {
    throw new Error('PDF.js library is not available. Please verify index.html CDN scripts.')
  }
  
  const pdf = await window.pdfjsLib.getDocument({ data: typedarray }).promise
  const images = []

  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    if (onProgress) {
      onProgress(pageNum, pdf.numPages)
    }
    const page = await pdf.getPage(pageNum)
    // Use scale 3.0 for high-DPI rendering — much better OCR accuracy
    const scale = 3.0
    const viewport = page.getViewport({ scale })
    
    const canvas = document.createElement('canvas')
    canvas.width = viewport.width
    canvas.height = viewport.height
    const ctx = canvas.getContext('2d')
    // White background so text contrasts clearly
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, canvas.width, canvas.height)
    
    await page.render({
      canvasContext: ctx,
      viewport: viewport,
      background: 'white',
    }).promise
    
    // Use lossless PNG so Tesseract gets crisp pixel data
    const dataUrl = canvas.toDataURL('image/png')
    images.push({ dataUrl })
  }
  return images
}

// ── PDF Generation with Embedded Google Maps Links ────────────────────────────

// ── PDF Generation with Embedded Google Maps Links or Selectable OCR Text ────

export async function generatePdf(processedSlides, ocrMode, scanHeight) {
  const pdfDoc = await PDFDocument.create()
  const helveticaFont = await pdfDoc.embedFont('Helvetica')

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

    if (ocrMode === 'gps') {
      // If GPS coordinates found, embed clickable PDF Link Annotation over bottom scanned area
      if (slide.gps && slide.gps.mapUrl) {
        const bottomHeight = dims.height * (scanHeight / 100)
        // PDF rectangle coordinates: [lower-left x, lower-left y, upper-right x, upper-right y]
        const rect = [0, 0, dims.width, bottomHeight]

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
    } else if (ocrMode === 'searchable') {
      // If in searchable mode, overlay invisible text lines matching their bounding boxes
      if (slide.ocrLines && slide.ocrLines.length > 0) {
        const cropHeight = dims.height * (scanHeight / 100)
        for (const line of slide.ocrLines) {
          const cleanedText = sanitizePdfText(line.text.trim())
          if (!cleanedText) continue

          const pdfX = line.bbox.x0
          // Offset text coordinates vertically based on crop area offset from top
          const pdfY = cropHeight - line.bbox.y1
          const fontSize = line.bbox.y1 - line.bbox.y0

          if (fontSize <= 0) continue

          try {
            page.drawText(cleanedText, {
              x: pdfX,
              y: pdfY,
              size: fontSize,
              font: helveticaFont,
              color: rgb(0, 0, 0),
              opacity: 0,
            })
          } catch (e) {
            console.warn('Failed to draw text line in PDF', cleanedText, e)
          }
        }
      }
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
  const [ocrMode, setOcrMode] = useState('gps') // 'gps' or 'searchable'
  const [scanHeight, setScanHeight] = useState(30) // default 30% for gps, 100% for searchable

  const fileInputRef = useRef(null)

  const handleOcrModeChange = (mode) => {
    setOcrMode(mode)
    setScanHeight(mode === 'gps' ? 30 : 100)
    setPdfBlob(null)
    setProcessedSlides([])
    if (file) {
      const baseName = file.name.replace(/\.(pptx|pdf)$/i, '')
      setPdfFileName(mode === 'gps' ? `${baseName}_GPS_Report.pdf` : `${baseName}_Searchable.pdf`)
    }
  }

  const handleFileSelect = (selectedFile) => {
    if (!selectedFile) return
    setFile(selectedFile)
    setProcessedSlides([])
    setPdfBlob(null)
    const baseName = selectedFile.name.replace(/\.(pptx|pdf)$/i, '')
    setPdfFileName(ocrMode === 'gps' ? `${baseName}_GPS_Report.pdf` : `${baseName}_Searchable.pdf`)
  }

  const startConversion = async () => {
    if (!file) return

    setIsProcessing(true)
    const isPdf = file.name.toLowerCase().endsWith('.pdf')
    setProgressMsg(isPdf ? 'Extracting page images from PDF...' : 'Extracting slide images from PPTX...')
    setProgressPct(5)

    let worker = null

    try {
      // Step 1: Extract all images/slides
      let rawImages = []
      if (isPdf) {
        rawImages = await renderPdfPages(file, (curr, total) => {
          setProgressMsg(`Rendering PDF page ${curr} of ${total}...`)
          setProgressPct(Math.floor(5 + (curr / total) * 10))
        })
      } else {
        rawImages = await extractAllImages(file)
      }

      if (!rawImages || rawImages.length === 0) {
        alert(isPdf ? '❌ Could not render any pages from the PDF file.' : '❌ No images or slides found in the PPTX file.')
        setIsProcessing(false)
        return
      }

      setProgressMsg(`Initializing OCR engine... (${rawImages.length} pages found)`)
      setProgressPct(15)

      // Step 2: Spin up a pool of parallel Tesseract workers for maximum throughput.
      // Use hardware concurrency capped at 4 workers (Tesseract WASM is memory-heavy).
      const POOL_SIZE = Math.min(Math.max(navigator.hardwareConcurrency || 2, 2), 4)
      const workers = []
      const workerParams = {
        tessedit_pageseg_mode: '3',
        preserve_interword_spaces: '1',
      }

      // Boot all workers in parallel
      await Promise.all(
        Array.from({ length: POOL_SIZE }, async () => {
          const w = await createWorker('eng')
          await w.setParameters(workerParams)
          workers.push(w)
        })
      )

      setProgressMsg(`OCR workers ready (${workers.length} threads). Processing ${rawImages.length} pages...`)
      setProgressPct(20)

      // Step 3: Parallel queue — each worker picks the next unstarted page.
      const results = new Array(rawImages.length)
      let nextIndex = 0
      let doneCount = 0
      const total = rawImages.length

      async function processWithWorker(w) {
        while (true) {
          const i = nextIndex++
          if (i >= total) break

          const item = rawImages[i]
          const currentNum = i + 1

          try {
            const imgEl = await loadImage(item.dataUrl)
            const croppedCanvas = cropBottomPercentageCanvas(imgEl, scanHeight)
            const croppedDataUrl = croppedCanvas.toDataURL('image/png')

            if (ocrMode === 'gps') {
              const { data: { text } } = await w.recognize(croppedCanvas)
              const gps = extractGpsCoordinates(text)
              results[i] = { ...item, croppedDataUrl, ocrText: text, ocrLines: [], gps }
            } else {
              const { data: { text, lines } } = await w.recognize(croppedCanvas)

              let usableLines = (lines || []).filter(line => {
                const cleaned = sanitizePdfText(line.text || '')
                return cleaned.trim().length > 0 && line.bbox && (line.bbox.y1 - line.bbox.y0) > 1
              })

              if (usableLines.length === 0 && text && text.trim().length > 0) {
                const rawLines = text.split('\n').filter(l => sanitizePdfText(l).trim().length > 0)
                const canvasH = croppedCanvas.height
                const canvasW = croppedCanvas.width
                const lineH = rawLines.length > 0 ? Math.floor(canvasH / rawLines.length) : 20
                usableLines = rawLines.map((lineText, idx) => ({
                  text: lineText,
                  bbox: { x0: 0, y0: idx * lineH, x1: canvasW, y1: (idx + 1) * lineH }
                }))
              }

              results[i] = {
                ...item,
                croppedDataUrl,
                ocrText: text,
                ocrLines: usableLines.map(line => ({ text: sanitizePdfText(line.text), bbox: line.bbox })),
                gps: null,
              }
            }
          } catch (err) {
            console.warn(`Failed OCR on page ${currentNum}`, err)
            results[i] = { ...item, croppedDataUrl: '', ocrText: '', ocrLines: [], gps: null }
          }

          doneCount++
          const pct = Math.floor(20 + (doneCount / total) * 65)
          setProgressPct(pct)
          setProgressMsg(`OCR: ${doneCount}/${total} pages done (${workers.length} parallel workers)...`)
        }
      }

      // Run all workers concurrently — they race to consume the queue
      await Promise.all(workers.map(w => processWithWorker(w)))

      // Terminate all workers
      await Promise.all(workers.map(w => w.terminate()))
      workers.length = 0

      // Step 4: Build High-Quality PDF
      setProgressMsg(ocrMode === 'gps' ? 'Generating PDF with embedded Google Maps links...' : 'Generating searchable PDF with selectable text overlay...')
      setProgressPct(90)

      const generatedBlob = await generatePdf(results, ocrMode, scanHeight)

      setProcessedSlides(results)
      setPdfBlob(generatedBlob)
      setProgressPct(100)
      setProgressMsg(ocrMode === 'gps' ? '✅ Complete! PDF with GPS Google Maps links ready.' : '✅ Complete! Searchable PDF ready for download.')
    } catch (err) {
      alert(`❌ Error during processing: ${err.message}`)
      console.error(err)
    } finally {
      // Worker pool is terminated inside the try block after processing.
      // This finally handles the legacy `worker` variable if init fails early.
      if (worker) {
        try { await worker.terminate() } catch { /* ignore */ }
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
      <div className="card gps-pdf-tool__header" style={{ flexDirection: 'column', alignItems: 'stretch', gap: '20px' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', flexWrap: 'wrap', gap: '16px' }}>
          <div>
            <h2 style={{ margin: 0, fontSize: '18px', fontWeight: 700 }}>
              {ocrMode === 'gps' ? 'PPTX to PDF (GPS OCR)' : 'PPTX to PDF (Searchable OCR)'}
            </h2>
            <p style={{ margin: '4px 0 0', fontSize: '13px', color: 'var(--muted)' }}>
              {ocrMode === 'gps'
                ? 'Upload any PPTX. Scans the bottom region for GPS coordinates and builds a PDF with embedded Google Maps links.'
                : 'Upload any PPTX. Scans the selected slide height from bottom to generate a searchable PDF with selectable text.'}
            </p>
          </div>

          {/* Mode Switcher */}
          <div className="ocr-mode-toggle" style={{ display: 'flex', background: 'var(--card-soft)', padding: '4px', borderRadius: '30px', border: '1px solid var(--border)' }}>
            <button
              type="button"
              className={`ghost ${ocrMode === 'gps' ? 'is-active' : ''}`}
              style={{
                border: 'none',
                borderRadius: '24px',
                padding: '6px 16px',
                fontSize: '12px',
                background: ocrMode === 'gps' ? '#0b7a38' : 'transparent',
                color: ocrMode === 'gps' ? '#fff' : 'var(--ink)',
                boxShadow: ocrMode === 'gps' ? '0 4px 10px rgba(11, 122, 56, 0.2)' : 'none',
                fontWeight: 600,
                textTransform: 'uppercase',
                letterSpacing: '0.05em'
              }}
              onClick={() => handleOcrModeChange('gps')}
              disabled={isProcessing}
            >
              📍 GPS Extract
            </button>
            <button
              type="button"
              className={`ghost ${ocrMode === 'searchable' ? 'is-active' : ''}`}
              style={{
                border: 'none',
                borderRadius: '24px',
                padding: '6px 16px',
                fontSize: '12px',
                background: ocrMode === 'searchable' ? '#0b7a38' : 'transparent',
                color: ocrMode === 'searchable' ? '#fff' : 'var(--ink)',
                boxShadow: ocrMode === 'searchable' ? '0 4px 10px rgba(11, 122, 56, 0.2)' : 'none',
                fontWeight: 600,
                textTransform: 'uppercase',
                letterSpacing: '0.05em'
              }}
              onClick={() => handleOcrModeChange('searchable')}
              disabled={isProcessing}
            >
              🔍 Searchable PDF
            </button>
          </div>
        </div>

        {/* Dynamic Scan Height Slider */}
        <div style={{
          display: 'flex',
          flexDirection: 'column',
          gap: '8px',
          background: 'var(--card-soft)',
          padding: '14px 18px',
          borderRadius: '14px',
          border: '1px solid var(--border)'
        }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '13px', fontWeight: 600 }}>
            <span style={{ color: 'var(--muted)' }}>
              Scan Height from Bottom:
            </span>
            <span style={{ color: '#0b7a38', background: 'rgba(11, 122, 56, 0.1)', padding: '2px 8px', borderRadius: '12px' }}>
              {scanHeight}% of Slide Height
            </span>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '14px' }}>
            <span style={{ fontSize: '11px', color: 'var(--muted)', fontWeight: 500 }}>10% (Footer)</span>
            <input
              type="range"
              min="10"
              max="100"
              step="5"
              value={scanHeight}
              onChange={(e) => {
                setScanHeight(parseInt(e.target.value, 10))
                setPdfBlob(null)
                setProcessedSlides([])
              }}
              disabled={isProcessing}
              style={{
                flex: 1,
                accentColor: '#0b7a38',
                cursor: 'pointer',
                height: '6px',
                borderRadius: '3px'
              }}
            />
            <span style={{ fontSize: '11px', color: 'var(--muted)', fontWeight: 500 }}>100% (Full Slide)</span>
          </div>
        </div>

        <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', alignItems: 'center', justifyContent: 'flex-end', marginTop: '4px' }}>
          <input
            ref={fileInputRef}
            type="file"
            accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation,.pdf,application/pdf"
            style={{ display: 'none' }}
            onChange={(e) => handleFileSelect(e.target.files?.[0])}
          />

          <button
            type="button"
            className="button button--secondary"
            onClick={() => fileInputRef.current?.click()}
            disabled={isProcessing}
          >
            {file ? '📁 Change File' : '📁 Choose PPTX/PDF File'}
          </button>

          {file && (
            <button
              type="button"
              className="button"
              onClick={startConversion}
              disabled={isProcessing}
            >
              {isProcessing ? '⚡ Processing...' : `🚀 Convert to ${ocrMode === 'gps' ? 'GPS PDF' : 'Searchable PDF'}`}
            </button>
          )}

          {pdfBlob && (
            <button
              type="button"
              className="button button--completed"
              onClick={handleDownloadPdf}
            >
              📥 Download {ocrMode === 'gps' ? 'GPS PDF' : 'Searchable PDF'}
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
              📊 Total Pages/Slides: <strong>{processedSlides.length}</strong>
            </span>
            {ocrMode === 'gps' ? (
              <span className="gps-pdf-badge gps-pdf-badge--gps">
                📍 GPS Detected: <strong>{gpsCount}</strong>
              </span>
            ) : (
              <span className="gps-pdf-badge" style={{ border: '1px solid rgba(11, 122, 56, 0.4)', background: 'rgba(11, 122, 56, 0.1)', color: '#0b7a38', borderRadius: '20px', padding: '6px 14px', fontSize: '13px', fontWeight: 600 }}>
                🔍 Text Layer Embedded: <strong>Yes</strong>
              </span>
            )}
          </div>

          <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
            {ocrMode === 'gps' ? (
              <>
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
              </>
            ) : (
              <button
                type="button"
                className="ghost"
                style={{ fontSize: '12px', padding: '6px 14px' }}
                onClick={() => {
                  const fullText = processedSlides
                    .map((s, idx) => `--- Page/Slide ${idx + 1} ---\n${s.ocrText.trim()}`)
                    .join('\n\n')
                  navigator.clipboard.writeText(fullText)
                  alert('📋 Copied all pages/slides raw OCR text to clipboard!')
                }}
              >
                📋 Copy All Extracted Text
              </button>
            )}

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
          {processedSlides.map((slide, idx) => {
            const hasData = (ocrMode === 'gps' && slide.gps) || (ocrMode === 'searchable' && slide.ocrLines && slide.ocrLines.length > 0)
            return (
              <div key={`slide-${idx}`} className={`gps-slide-card${hasData ? ' has-gps' : ''}`}>
                <div className="gps-slide-card__header">
                  <span>Page/Slide {idx + 1}</span>
                  {ocrMode === 'gps' ? (
                    slide.gps ? (
                      <span className="gps-status-badge is-found">📍 Coordinates Found</span>
                    ) : (
                      <span className="gps-status-badge is-none">No GPS Text</span>
                    )
                  ) : (
                    slide.ocrLines && slide.ocrLines.length > 0 ? (
                      <span className="gps-status-badge is-found" style={{ background: 'rgba(11, 122, 56, 0.1)', color: '#0b7a38', borderColor: 'rgba(11, 122, 56, 0.3)' }}>🔍 Text Indexed</span>
                    ) : (
                      <span className="gps-status-badge is-none">No Text Detected</span>
                    )
                  )}
                </div>

                <div className="gps-slide-card__media">
                  <img src={slide.dataUrl} alt={`Page/Slide ${idx + 1}`} className="gps-slide-card__img" />
                  {slide.croppedDataUrl && (
                    <div className="gps-slide-card__crop-preview" title={`Cropped Bottom ${scanHeight}% OCR Region`}>
                      <img src={slide.croppedDataUrl} alt={`Bottom ${scanHeight}% Region`} />
                      <span className="gps-slide-card__crop-tag">Scanned Bottom {scanHeight}%</span>
                    </div>
                  )}
                </div>

                <div className="gps-slide-card__body">
                  {ocrMode === 'gps' ? (
                    slide.gps ? (
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
                        No coordinate patterns (Lat/Long) detected in the bottom {scanHeight}% overlay.
                      </p>
                    )
                  ) : (
                    slide.ocrLines && slide.ocrLines.length > 0 ? (
                      <div>
                        <p style={{ margin: 0, fontSize: '13px', fontWeight: 600, color: 'var(--ink)' }}>
                          📝 {slide.ocrLines.length} lines of text indexed.
                        </p>
                        <p style={{ margin: '4px 0 0', fontSize: '12px', color: 'var(--muted)' }}>
                          Embedded invisibly in the PDF at their original visual coordinates.
                        </p>
                      </div>
                    ) : (
                      <p style={{ margin: 0, fontSize: '12px', color: 'var(--muted)', fontStyle: 'italic' }}>
                        No text detected in the bottom {scanHeight}% region of this slide.
                      </p>
                    )
                  )}

                  {slide.ocrText && (
                    <details style={{ marginTop: '10px' }}>
                      <summary style={{ fontSize: '11px', color: 'var(--muted)', cursor: 'pointer' }}>
                        View Raw OCR Text
                      </summary>
                      <pre className="gps-ocr-snippet">{slide.ocrText.trim()}</pre>
                    </details>
                  )}
                </div>
              </div>
            )
          })}
        </div>
      )}
    </div>
  )
}
