import JSZip from 'jszip'

const EMU_PER_INCH = 914400
const DEFAULT_SLIDE_SIZE = { cx: 12192000, cy: 6858000 }

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
  Parse entire PPTX presentation into editable slide representations.
 */
export const parsePptxForEditing = async (file) => {
  if (!file) return null

  const arrayBuffer = await file.arrayBuffer()
  const zip = await JSZip.loadAsync(arrayBuffer)
  const slidePaths = getSlidePaths(zip)
  const slideSize = await getSlideSize(zip)
  const imageCache = new Map()

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

    // 1. Extract Text Shapes (<p:sp>)
    const spNodes = Array.from(
      slideDoc.getElementsByTagName('p:sp').length
        ? slideDoc.getElementsByTagName('p:sp')
        : slideDoc.getElementsByTagName('sp'),
    )

    for (let sIdx = 0; sIdx < spNodes.length; sIdx++) {
      const sp = spNodes[sIdx]

      // Extract transform
      const xfrm = getFirstTag(sp, ['a:xfrm', 'xfrm'])
      const off = xfrm ? getFirstTag(xfrm, ['a:off', 'off']) : null
      const ext = xfrm ? getFirstTag(xfrm, ['a:ext', 'ext']) : null

      const xEmu = Number(off?.getAttribute('x') || 0)
      const yEmu = Number(off?.getAttribute('y') || 0)
      const cxEmu = Number(ext?.getAttribute('cx') || 0)
      const cyEmu = Number(ext?.getAttribute('cy') || 0)

      // Collect text runs
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

      // Check placeholder type (title / subtitle / body)
      const nvSpPr = getFirstTag(sp, ['p:nvSpPr', 'nvSpPr'])
      const nvPr = nvSpPr ? getFirstTag(nvSpPr, ['p:nvPr', 'nvPr']) : null
      const ph = nvPr ? getFirstTag(nvPr, ['p:ph', 'ph']) : null
      const phType = ph ? ph.getAttribute('type') || 'body' : 'body'
      const isTitle = phType.includes('title') || phType === 'ctrTitle' || sIdx === 0

      elements.push({
        id: `slide_${idx + 1}_text_${sIdx}`,
        spIndex: sIdx,
        type: 'text',
        tagName: isTitle ? '<p:sp> Title Tag' : '<p:sp> Text Tag',
        text,
        originalText: text,
        isTitle,
        phType,
        xInch: xEmu / EMU_PER_INCH,
        yInch: yEmu / EMU_PER_INCH,
        wInch: cxEmu / EMU_PER_INCH,
        hInch: cyEmu / EMU_PER_INCH,
      })
    }

    // 2. Extract Pictures (<p:pic>)
    const picNodes = Array.from(
      slideDoc.getElementsByTagName('p:pic').length
        ? slideDoc.getElementsByTagName('p:pic')
        : slideDoc.getElementsByTagName('pic'),
    )

    for (let pIdx = 0; pIdx < picNodes.length; pIdx++) {
      const pic = picNodes[pIdx]
      const blip = getFirstTag(pic, ['a:blip', 'blip'])
      const embed = blip?.getAttribute('r:embed') || blip?.getAttribute('embed') || ''
      const target = embed ? relMap.get(embed) : ''

      const xfrm = getFirstTag(pic, ['a:xfrm', 'xfrm'])
      const off = xfrm ? getFirstTag(xfrm, ['a:off', 'off']) : null
      const ext = xfrm ? getFirstTag(xfrm, ['a:ext', 'ext']) : null

      const xEmu = Number(off?.getAttribute('x') || 0)
      const yEmu = Number(off?.getAttribute('y') || 0)
      const cxEmu = Number(ext?.getAttribute('cx') || 0)
      const cyEmu = Number(ext?.getAttribute('cy') || 0)

      const dataUrl = target ? await getImageDataUrl(zip, target, imageCache) : ''

      elements.push({
        id: `slide_${idx + 1}_pic_${pIdx}`,
        picIndex: pIdx,
        type: 'image',
        tagName: '<p:pic> Image Tag',
        embed,
        target,
        dataUrl,
        originalDataUrl: dataUrl,
        xInch: xEmu / EMU_PER_INCH,
        yInch: yEmu / EMU_PER_INCH,
        wInch: cxEmu / EMU_PER_INCH,
        hInch: cyEmu / EMU_PER_INCH,
      })
    }

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
      rawXml: slideXmlText,
    })
  }

  return {
    filename: file.name,
    slides,
    slideSize: {
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
 * Re-compiles edited presentation and triggers browser download.
 */
export const exportEditedPptx = async (originalFile, slidesData, outputFileName) => {
  if (!originalFile || !slidesData || !slidesData.length) return

  const arrayBuffer = await originalFile.arrayBuffer()
  const zip = await JSZip.loadAsync(arrayBuffer)
  const originalSlidePaths = getSlidePaths(zip)

  // 1. Update text content & replace images in zip
  for (let sIdx = 0; sIdx < slidesData.length; sIdx++) {
    const slide = slidesData[sIdx]
    const origPath = originalSlidePaths[sIdx] || slide.xmlPath
    const slideFile = zip.file(origPath)
    if (!slideFile) continue

    const slideXmlText = await slideFile.async('text')
    const doc = parseXml(slideXmlText)

    const spNodes = Array.from(
      doc.getElementsByTagName('p:sp').length
        ? doc.getElementsByTagName('p:sp')
        : doc.getElementsByTagName('sp'),
    )

    // Update text elements
    slide.elements.forEach((elem) => {
      if (elem.type === 'text' && elem.text !== elem.originalText) {
        if (spNodes[elem.spIndex]) {
          const sp = spNodes[elem.spIndex]
          const tNodes = Array.from(
            sp.getElementsByTagName('a:t').length
              ? sp.getElementsByTagName('a:t')
              : sp.getElementsByTagName('t'),
          )

          if (tNodes.length > 0) {
            // Set primary text run content
            tNodes[0].textContent = elem.text
            // Clear remaining text runs in this shape so edited text is clean
            for (let k = 1; k < tNodes.length; k++) {
              tNodes[k].textContent = ''
            }
          }
        }
      }

      // Update image elements if image data URL was replaced
      if (elem.type === 'image' && elem.dataUrl && elem.dataUrl !== elem.originalDataUrl) {
        if (elem.target) {
          try {
            const imageBytes = dataUrlToUint8Array(elem.dataUrl)
            zip.file(elem.target, imageBytes)
          } catch (e) {
            console.error('Failed to replace image asset in PPTX:', e)
          }
        }
      }
    })

    // Re-serialize XML
    const serializer = new XMLSerializer()
    const updatedXml = serializer.serializeToString(doc)
    zip.file(origPath, updatedXml)
  }

  // 2. Generate updated Blob and download
  const blob = await zip.generateAsync({
    type: 'blob',
    mimeType:
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })

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
