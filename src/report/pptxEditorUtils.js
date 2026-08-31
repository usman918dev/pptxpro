import JSZip from 'jszip'
import { applyEditsToSlideXml } from './importPptx'

const EMU_PER_INCH = 914400
const DEFAULT_SLIDE_SIZE = { cx: 12192000, cy: 6858000 }
const MAX_BACKGROUND_RATIO = 0.9

const SCHEME_COLOR_MAP = {
  dk1: '000000', tx1: '000000',
  lt1: 'FFFFFF', bg1: 'FFFFFF',
  dk2: '1F3864', tx2: '1F3864',
  lt2: 'D9E2F3', bg2: 'D9E2F3',
  accent1: '4472C4',
  accent2: 'ED7D31',
  accent3: 'A9D18E',
  accent4: 'FFC000',
  accent5: '5A96D6',
  accent6: '70AD47',
  hlink: '0563C1',
  folHlink: '954F72',
}

const parseXml = (xmlText) => {
  const parser = new DOMParser()
  const doc = parser.parseFromString(xmlText, 'application/xml')
  if (doc.getElementsByTagName('parsererror').length > 0) {
    throw new Error('Unable to parse PPTX XML data.')
  }
  return doc
}

const getFirstTag = (node, names) => {
  for (const name of names) {
    const list = node.getElementsByTagName(name)
    if (list.length > 0) {
      return list[0]
    }
  }
  return null
}

const resolveSolidFillColor = (solidFillEl) => {
  if (!solidFillEl) return ''
  const srgb = solidFillEl.getElementsByTagName('a:srgbClr')[0]
  if (srgb) return srgb.getAttribute('val') || ''
  const scheme = solidFillEl.getElementsByTagName('a:schemeClr')[0]
  if (scheme) return SCHEME_COLOR_MAP[scheme.getAttribute('val') || ''] || ''
  return ''
}

const getSlideSize = async (zip) => {
  const file = zip.file('ppt/presentation.xml')
  if (!file) return { ...DEFAULT_SLIDE_SIZE }
  try {
    const doc = parseXml(await file.async('text'))
    const sizeNode =
      doc.getElementsByTagName('p:sldSz')[0] || doc.getElementsByTagName('sldSz')[0]
    if (!sizeNode) return { ...DEFAULT_SLIDE_SIZE }
    const cx = Number(sizeNode.getAttribute('cx') || DEFAULT_SLIDE_SIZE.cx)
    const cy = Number(sizeNode.getAttribute('cy') || DEFAULT_SLIDE_SIZE.cy)
    return { cx, cy }
  } catch {
    return { ...DEFAULT_SLIDE_SIZE }
  }
}

const getSlidePaths = (zip) => {
  const slideRegex = /^ppt\/slides\/slide(\d+)\.xml$/i
  const slides = []
  zip.forEach((relativePath) => {
    const match = slideRegex.exec(relativePath)
    if (match) {
      slides.push({ path: relativePath, index: Number(match[1]) })
    }
  })
  slides.sort((a, b) => a.index - b.index)
  return slides.map((item) => item.path)
}

const getSlideRelsPath = (slidePath) =>
  `${slidePath.replace('ppt/slides/slide', 'ppt/slides/_rels/slide')}.rels`

const normalizeTarget = (target) => {
  let cleaned = target.replace(/^\.\//, '')
  if (cleaned.startsWith('../')) {
    cleaned = cleaned.replace(/^\.\.\//, '')
  }
  cleaned = cleaned.replace(/^\/+/, '')
  if (cleaned.startsWith('ppt/')) return cleaned
  return `ppt/${cleaned}`
}

const getSlideRelations = async (zip, relsPath) => {
  const relsFile = zip.file(relsPath)
  if (!relsFile) return new Map()
  try {
    const doc = parseXml(await relsFile.async('text'))
    const nodes = Array.from(doc.getElementsByTagName('Relationship'))
    const map = new Map()
    nodes.forEach((node) => {
      const id = node.getAttribute('Id')
      const target = node.getAttribute('Target')
      const type = node.getAttribute('Type')
      if (!id || !target) return
      if (type && !/image/i.test(type)) return
      map.set(id, normalizeTarget(target))
    })
    return map
  } catch {
    return new Map()
  }
}

const getMimeType = (path) => {
  const lower = path.toLowerCase()
  if (lower.endsWith('.png')) return 'image/png'
  if (lower.endsWith('.gif')) return 'image/gif'
  if (lower.endsWith('.webp')) return 'image/webp'
  if (lower.endsWith('.svg')) return 'image/svg+xml'
  return 'image/jpeg'
}

const getImageDataUrl = async (zip, imagePath, cache) => {
  if (!imagePath) return ''
  if (cache.has(imagePath)) return cache.get(imagePath)
  const file = zip.file(imagePath)
  if (!file) return ''
  try {
    const base64 = await file.async('base64')
    const dataUrl = `data:${getMimeType(imagePath)};base64,${base64}`
    cache.set(imagePath, dataUrl)
    return dataUrl
  } catch {
    return ''
  }
}

/**
 * Extract Tables from a slide XML document
 */
const getTablesFromSlide = (slideDoc, slideSize) => {
  const frameNodes = Array.from(
    slideDoc.getElementsByTagName('p:graphicFrame').length
      ? slideDoc.getElementsByTagName('p:graphicFrame')
      : slideDoc.getElementsByTagName('graphicFrame'),
  )
  const results = []

  for (const frame of frameNodes) {
    const tbl =
      frame.getElementsByTagName('a:tbl')[0] ||
      frame.getElementsByTagName('tbl')[0]
    if (!tbl) continue

    const xfrm = getFirstTag(frame, ['p:xfrm', 'a:xfrm', 'xfrm'])
    const off = xfrm ? getFirstTag(xfrm, ['a:off', 'off']) : null
    const ext = xfrm ? getFirstTag(xfrm, ['a:ext', 'ext']) : null

    const xEmu = Number(off?.getAttribute('x') || 0)
    const yEmu = Number(off?.getAttribute('y') || 0)
    const cxEmu = Number(ext?.getAttribute('cx') || 0)
    const cyEmu = Number(ext?.getAttribute('cy') || 0)

    const slideWidthInch = slideSize.cx / EMU_PER_INCH

    const trNodes = Array.from(
      tbl.getElementsByTagName('a:tr').length
        ? tbl.getElementsByTagName('a:tr')
        : tbl.getElementsByTagName('tr'),
    )

    const rows = trNodes.map((tr) => {
      const tcNodes = Array.from(
        tr.getElementsByTagName('a:tc').length
          ? tr.getElementsByTagName('a:tc')
          : tr.getElementsByTagName('tc'),
      )
      return tcNodes.map((tc) => {
        const tNodes = Array.from(
          tc.getElementsByTagName('a:t').length
            ? tc.getElementsByTagName('a:t')
            : tc.getElementsByTagName('t'),
        )
        const text = tNodes.map((t) => t.textContent || '').join('')

        const rPr =
          tc.getElementsByTagName('a:rPr')[0] ||
          tc.getElementsByTagName('rPr')[0]
        const bold = rPr?.getAttribute('b') === '1'
        const sz = rPr ? Number(rPr.getAttribute('sz') || 0) : 0
        const ptSize = sz > 0 ? sz / 100 : 11
        const fontSizePct = (ptSize / 72 / slideWidthInch) * 100

        let fontFace = 'Calibri'
        if (rPr) {
          const latin = rPr.getElementsByTagName('a:latin')[0]
          if (latin) fontFace = latin.getAttribute('typeface') || 'Calibri'
        }

        const color = rPr ? resolveSolidFillColor(rPr.getElementsByTagName('a:solidFill')[0]) : ''

        const pPr =
          tc.getElementsByTagName('a:pPr')[0] ||
          tc.getElementsByTagName('pPr')[0]
        const algn = pPr?.getAttribute('algn') || 'l'
        const alignMap = { l: 'left', ctr: 'center', r: 'right', just: 'justify' }

        const tcPr =
          tc.getElementsByTagName('a:tcPr')[0] ||
          tc.getElementsByTagName('tcPr')[0]

        const fillColor = resolveSolidFillColor(
          tcPr?.getElementsByTagName('a:solidFill')[0] ?? null
        )

        let borderColor = 'cccccc'
        if (tcPr) {
          const lnEl =
            tcPr.getElementsByTagName('a:lnL')[0] ||
            tcPr.getElementsByTagName('a:lnT')[0]
          const resolved = resolveSolidFillColor(
            lnEl?.getElementsByTagName('a:solidFill')[0] ?? null
          )
          if (resolved) borderColor = resolved
        }

        return {
          text,
          bold,
          fontSizePct,
          color,
          fontFace,
          align: alignMap[algn] || 'left',
          fillColor,
          borderColor,
        }
      })
    })

    results.push({
      id: `tbl_${xEmu}_${yEmu}`,
      type: 'table',
      tagName: '<a:tbl> Table',
      xEmu,
      yEmu,
      cxEmu,
      cyEmu,
      xInch: xEmu / EMU_PER_INCH,
      yInch: yEmu / EMU_PER_INCH,
      wInch: cxEmu / EMU_PER_INCH,
      hInch: cyEmu / EMU_PER_INCH,
      xPct: (xEmu / slideSize.cx) * 100,
      yPct: (yEmu / slideSize.cy) * 100,
      wPct: (cxEmu / slideSize.cx) * 100,
      hPct: (cyEmu / slideSize.cy) * 100,
      fontSizePct: ((11 / 72) / slideWidthInch) * 100,
      rows,
    })
  }

  return results
}

/**
 * Parse entire PPTX presentation into full editable slide representations.
 */
export const parsePptxForEditing = async (file) => {
  if (!file) return null

  const arrayBuffer = await file.arrayBuffer()
  const zip = await JSZip.loadAsync(arrayBuffer)
  const slidePaths = getSlidePaths(zip)
  const slideSize = await getSlideSize(zip)
  const imageCache = new Map()

  const slideAreaEmu = slideSize.cx * slideSize.cy
  const maxBgAreaEmu = slideAreaEmu * MAX_BACKGROUND_RATIO

  const slides = []

  for (let idx = 0; idx < slidePaths.length; idx++) {
    const slidePath = slidePaths[idx]
    const slideFile = zip.file(slidePath)
    if (!slideFile) continue

    const slideXmlText = await slideFile.async('text')
    const slideDoc = parseXml(slideXmlText)
    const relsPath = getSlideRelsPath(slidePath)
    const relMap = await getSlideRelations(zip, relsPath)

    const elements = []

    // 1. Extract Pictures (<p:pic>)
    const picNodes = Array.from(
      slideDoc.getElementsByTagName('p:pic').length
        ? slideDoc.getElementsByTagName('p:pic')
        : slideDoc.getElementsByTagName('pic'),
    )

    let backgroundDataUrl = ''
    const contentPics = []

    // 1a. Check explicit <p:bg> element on slide
    const bgNode = getFirstTag(slideDoc, ['p:bg', 'bg'])
    if (bgNode) {
      const blip = getFirstTag(bgNode, ['a:blip', 'blip'])
      const embed = blip?.getAttribute('r:embed') || blip?.getAttribute('embed') || ''
      const target = embed ? relMap.get(embed) : ''
      if (target) {
        backgroundDataUrl = await getImageDataUrl(zip, target, imageCache)
      }
    }

    // 1b. Extract Pictures (<p:pic>)
    for (let pIdx = 0; pIdx < picNodes.length; pIdx++) {
      const pic = picNodes[pIdx]
      const blip = getFirstTag(pic, ['a:blip', 'blip'])
      const embed = blip?.getAttribute('r:embed') || blip?.getAttribute('embed') || ''
      const target = embed ? relMap.get(embed) : ''

      const xfrm = getFirstTag(pic, ['p:xfrm', 'a:xfrm', 'xfrm'])
      const off = xfrm ? getFirstTag(xfrm, ['a:off', 'off']) : null
      const ext = xfrm ? getFirstTag(xfrm, ['a:ext', 'ext']) : null

      const xEmu = Number(off?.getAttribute('x') || 0)
      const yEmu = Number(off?.getAttribute('y') || 0)
      const cxEmu = Number(ext?.getAttribute('cx') || 0)
      const cyEmu = Number(ext?.getAttribute('cy') || 0)
      const areaEmu = cxEmu * cyEmu

      const dataUrl = target ? await getImageDataUrl(zip, target, imageCache) : ''

      if (areaEmu >= maxBgAreaEmu && !backgroundDataUrl) {
        backgroundDataUrl = dataUrl
      } else {
        contentPics.push({
          pic,
          pIdx,
          embed,
          target,
          dataUrl,
          xEmu,
          yEmu,
          cxEmu,
          cyEmu,
        })
      }
    }

    contentPics.forEach(({ pic, pIdx, embed, target, dataUrl, xEmu, yEmu, cxEmu, cyEmu }) => {
      elements.push({
        id: `slide_${idx + 1}_pic_${pIdx}`,
        picIndex: pIdx,
        type: 'image',
        tagName: '<p:pic> Image',
        embed,
        target,
        dataUrl,
        originalDataUrl: dataUrl,
        xEmu,
        yEmu,
        cxEmu,
        cyEmu,
        xInch: xEmu / EMU_PER_INCH,
        yInch: yEmu / EMU_PER_INCH,
        wInch: cxEmu / EMU_PER_INCH,
        hInch: cyEmu / EMU_PER_INCH,
        xPct: (xEmu / slideSize.cx) * 100,
        yPct: (yEmu / slideSize.cy) * 100,
        wPct: (cxEmu / slideSize.cx) * 100,
        hPct: (cyEmu / slideSize.cy) * 100,
      })
    })

    // 2. Extract Text Shapes (<p:sp>)
    const spNodes = Array.from(
      slideDoc.getElementsByTagName('p:sp').length
        ? slideDoc.getElementsByTagName('p:sp')
        : slideDoc.getElementsByTagName('sp'),
    )

    for (let sIdx = 0; sIdx < spNodes.length; sIdx++) {
      const sp = spNodes[sIdx]

      const xfrm = getFirstTag(sp, ['p:xfrm', 'a:xfrm', 'xfrm'])
      const off = xfrm ? getFirstTag(xfrm, ['a:off', 'off']) : null
      const ext = xfrm ? getFirstTag(xfrm, ['a:ext', 'ext']) : null

      const xEmu = Number(off?.getAttribute('x') || 0)
      const yEmu = Number(off?.getAttribute('y') || 0)
      const cxEmu = Number(ext?.getAttribute('cx') || 0)
      const cyEmu = Number(ext?.getAttribute('cy') || 0)
      const areaEmu = cxEmu * cyEmu

      if (areaEmu >= maxBgAreaEmu) continue

      const tNodes = Array.from(
        sp.getElementsByTagName('a:t').length
          ? sp.getElementsByTagName('a:t')
          : sp.getElementsByTagName('t'),
      )

      const text = tNodes
        .map((t) => t.textContent || '')
        .join('')
        .trim()

      if (!text) continue

      const rPr = sp.getElementsByTagName('a:rPr')[0] || sp.getElementsByTagName('rPr')[0]
      const bold = rPr?.getAttribute('b') === '1'
      const sz = rPr ? Number(rPr.getAttribute('sz') || 0) : 0
      const ptSize = sz > 0 ? sz / 100 : 14
      const fontSizePct = Math.max(ptSize * 0.104, 0.8)

      let fontFace = 'Calibri'
      if (rPr) {
        const latin = rPr.getElementsByTagName('a:latin')[0]
        if (latin) fontFace = latin.getAttribute('typeface') || 'Calibri'
      }

      const color = rPr ? resolveSolidFillColor(rPr.getElementsByTagName('a:solidFill')[0]) : ''

      const pPr = sp.getElementsByTagName('a:pPr')[0] || sp.getElementsByTagName('pPr')[0]
      const algn = pPr?.getAttribute('algn') || 'l'
      const alignMap = { l: 'left', ctr: 'center', r: 'right', just: 'justify' }

      const nvSpPr = getFirstTag(sp, ['p:nvSpPr', 'nvSpPr'])
      const nvPr = nvSpPr ? getFirstTag(nvSpPr, ['p:nvPr', 'nvPr']) : null
      const ph = nvPr ? getFirstTag(nvPr, ['p:ph', 'ph']) : null
      const phType = ph ? ph.getAttribute('type') || 'body' : 'body'
      const isTitle = phType.includes('title') || phType === 'ctrTitle' || sIdx === 0

      elements.push({
        id: `slide_${idx + 1}_text_${sIdx}`,
        spIndex: sIdx,
        type: 'text',
        tagName: isTitle ? '<p:sp> Title' : '<p:sp> Text',
        text,
        originalText: text,
        fontFace,
        fontSizePct,
        bold,
        color,
        align: alignMap[algn] || 'left',
        isTitle,
        phType,
        xEmu,
        yEmu,
        cxEmu,
        cyEmu,
        xInch: xEmu / EMU_PER_INCH,
        yInch: yEmu / EMU_PER_INCH,
        wInch: cxEmu / EMU_PER_INCH,
        hInch: cyEmu / EMU_PER_INCH,
        xPct: (xEmu / slideSize.cx) * 100,
        yPct: (yEmu / slideSize.cy) * 100,
        wPct: (cxEmu / slideSize.cx) * 100,
        hPct: (cyEmu / slideSize.cy) * 100,
      })
    }

    // 3. Extract Tables (<a:tbl>)
    const tables = getTablesFromSlide(slideDoc, slideSize)
    elements.push(...tables)

    // Determine slide title
    const titleElem = elements.find((e) => e.type === 'text' && e.isTitle) || elements.find((e) => e.type === 'text')
    const slideTitle = titleElem ? titleElem.text.slice(0, 40) : `Slide ${idx + 1}`

    slides.push({
      id: `slide_${idx + 1}_${Date.now()}_${idx}`,
      slideNumber: idx + 1,
      title: slideTitle,
      xmlPath: slidePath,
      relsPath,
      elements,
      backgroundDataUrl,
      rawXml: slideXmlText,
    })
  }

  // Propagate default master background image to inner slides if they don't specify their own
  const defaultBg = slides.find((s) => Boolean(s.backgroundDataUrl))?.backgroundDataUrl || ''
  if (defaultBg) {
    slides.forEach((s) => {
      if (!s.backgroundDataUrl) {
        s.backgroundDataUrl = defaultBg
      }
    })
  }

  return {
    filename: file.name,
    slides,
    slideSize: {
      cx: slideSize.cx,
      cy: slideSize.cy,
      wInch: slideSize.cx / EMU_PER_INCH,
      hInch: slideSize.cy / EMU_PER_INCH,
    },
  }
}

/**
 * Convert Data URL (Base64) or Blob to Uint8Array for JSZip file updates.
 */
const dataUrlToUint8Array = (dataUrl) => {
  const arr = dataUrl.split(',')
  const bstr = atob(arr[1])
  let n = bstr.length
  const u8arr = new Uint8Array(n)
  while (n--) {
    u8arr[n] = bstr.charCodeAt(n)
  }
  return u8arr
}

/**
 * Re-compiles edited presentation and triggers browser save/download.
 * Supports Native File System Access API for direct 1-click file overwrite.
 */
export const exportEditedPptx = async (originalFile, slidesData, options = {}) => {
  const { download = true, fileHandle = null, saveAs = false, outputFileName = '' } =
    typeof options === 'string' ? { outputFileName: options } : options

  if (!originalFile || !slidesData || !slidesData.length) return null

  const arrayBuffer = await originalFile.arrayBuffer()
  const zip = await JSZip.loadAsync(arrayBuffer)
  const originalSlidePaths = getSlidePaths(zip)

  const parser = new DOMParser()
  const serializer = new XMLSerializer()

  for (let sIdx = 0; sIdx < slidesData.length; sIdx++) {
    const slide = slidesData[sIdx]
    const origPath = slide.xmlPath || originalSlidePaths[sIdx]
    if (!origPath) continue

    // Resolve exact file key in zip for slide XML
    let zipKey = origPath
    let slideFile = zip.file(origPath)
    if (!slideFile) {
      const lower = origPath.toLowerCase()
      for (const k of Object.keys(zip.files)) {
        if (k.toLowerCase() === lower) {
          zipKey = k
          slideFile = zip.files[k]
          break
        }
      }
    }

    if (!slideFile) continue

    const slideXmlText = await slideFile.async('text')
    const texts = slide.elements ? slide.elements.filter((e) => e.type === 'text') : []
    const tables = slide.elements ? slide.elements.filter((e) => e.type === 'table') : []

    // 1. Apply text and table edits using raw XML DOM transformer engine
    let updatedXml = applyEditsToSlideXml(slideXmlText, { texts, tables })

    // 2. Process image replacements and update slide relationships (.rels)
    const relsPath = getSlideRelsPath(zipKey)
    let relsFile = zip.file(relsPath)
    if (!relsFile) {
      const lowerRels = relsPath.toLowerCase()
      for (const k of Object.keys(zip.files)) {
        if (k.toLowerCase() === lowerRels) {
          relsFile = zip.files[k]
          break
        }
      }
    }

    const imageElements = slide.elements ? slide.elements.filter((e) => e.type === 'image') : []

    if (relsFile && imageElements.length > 0) {
      const relsXmlText = await relsFile.async('text')
      const relsDoc = parser.parseFromString(relsXmlText, 'application/xml')
      const relsRoot =
        relsDoc.getElementsByTagName('Relationships')[0] || relsDoc.documentElement

      let maxRId = 100
      Array.from(relsDoc.getElementsByTagName('Relationship')).forEach((rel) => {
        const m = (rel.getAttribute('Id') || '').match(/rId(\d+)/)
        if (m) maxRId = Math.max(maxRId, Number(m[1]))
      })

      let imgCounter = 0
      imageElements.forEach((img) => {
        if (img.dataUrl && img.dataUrl !== img.originalDataUrl) {
          try {
            const b64 = img.dataUrl.replace(/^data:[^;]+;base64,/, '')
            const binary = atob(b64)
            const bytes = new Uint8Array(binary.length)
            for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)

            const ext = img.dataUrl.includes('image/png')
              ? 'png'
              : img.dataUrl.includes('image/gif')
              ? 'gif'
              : 'jpg'
            const mediaFileName = `imp_s${sIdx + 1}_${imgCounter++}.${ext}`
            const mediaPath = `ppt/media/${mediaFileName}`
            zip.file(mediaPath, bytes.buffer)

            maxRId++
            const newRId = `rId${maxRId}`

            const relNode = relsDoc.createElementNS(
              'http://schemas.openxmlformats.org/package/2006/relationships',
              'Relationship',
            )
            relNode.setAttribute('Id', newRId)
            relNode.setAttribute(
              'Type',
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            )
            relNode.setAttribute('Target', `../media/${mediaFileName}`)
            relsRoot.appendChild(relNode)

            if (img.embed) {
              updatedXml = updatedXml.replace(
                new RegExp(`r:embed="${img.embed}"`, 'g'),
                `r:embed="${newRId}"`,
              )
            }
          } catch (e) {
            console.error('Failed to replace image asset in PPTX rels:', e)
          }
        }
      })

      let newRelsXml = serializer.serializeToString(relsDoc)
      if (!newRelsXml.startsWith('<?xml')) {
        newRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + newRelsXml
      }
      zip.file(relsFile.name || relsPath, newRelsXml)
    }

    zip.file(zipKey, updatedXml)
  }

  // Generate updated Blob
  const blob = await zip.generateAsync({
    type: 'blob',
    mimeType:
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })

  // 1. Direct overwrite to fileHandle on disk (Native File System Access API)
  if (fileHandle && typeof fileHandle.createWritable === 'function' && !saveAs) {
    try {
      const writable = await fileHandle.createWritable()
      await writable.write(blob)
      await writable.close()
      return { blob, savedDirectly: true, fileName: fileHandle.name }
    } catch (err) {
      console.warn('Direct file handle write failed or permission denied, falling back:', err)
    }
  }

  // 2. "Save As" file picker on disk
  if (saveAs && typeof window.showSaveFilePicker === 'function') {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: outputFileName || originalFile.name,
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
      await writable.write(blob)
      await writable.close()
      return { blob, savedDirectly: true, fileName: handle.name, fileHandle: handle }
    } catch (err) {
      if (err.name === 'AbortError') return null
    }
  }

  // 3. Fallback: standard browser download link
  if (download) {
    const downloadName = outputFileName || `Edited_${originalFile.name}`
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = downloadName
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
    URL.revokeObjectURL(url)
  }

  return { blob, savedDirectly: false }
}
