import JSZip from 'jszip'

const DEFAULT_SLIDE_SIZE = { cx: 12192000, cy: 6858000 }
const MAX_BACKGROUND_RATIO = 0.9
const MIN_IMAGE_RATIO = 0.05

// EMU (English Metric Units) per inch in OOXML
const EMU_PER_INCH = 914400

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
  if (!file) {
    return { ...DEFAULT_SLIDE_SIZE }
  }
  const doc = parseXml(await file.async('text'))
  const sizeNode =
    doc.getElementsByTagName('p:sldSz')[0] || doc.getElementsByTagName('sldSz')[0]
  if (!sizeNode) {
    return { ...DEFAULT_SLIDE_SIZE }
  }
  const cx = Number(sizeNode.getAttribute('cx') || DEFAULT_SLIDE_SIZE.cx)
  const cy = Number(sizeNode.getAttribute('cy') || DEFAULT_SLIDE_SIZE.cy)
  return { cx, cy }
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
  if (cleaned.startsWith('ppt/')) {
    return cleaned
  }
  if (cleaned.startsWith('media/')) {
    return `ppt/${cleaned}`
  }
  return `ppt/${cleaned}`
}

const getSlideRelations = async (zip, relsPath) => {
  const relsFile = zip.file(relsPath)
  if (!relsFile) {
    return new Map()
  }
  const doc = parseXml(await relsFile.async('text'))
  const nodes = Array.from(doc.getElementsByTagName('Relationship'))
  const map = new Map()
  nodes.forEach((node) => {
    const id = node.getAttribute('Id')
    const target = node.getAttribute('Target')
    const type = node.getAttribute('Type')
    if (!id || !target) {
      return
    }
    if (type && !/image/i.test(type)) {
      return
    }
    map.set(id, normalizeTarget(target))
  })
  return map
}

const getMimeType = (path) => {
  const lower = path.toLowerCase()
  if (lower.endsWith('.png')) {
    return 'image/png'
  }
  if (lower.endsWith('.gif')) {
    return 'image/gif'
  }
  if (lower.endsWith('.webp')) {
    return 'image/webp'
  }
  if (lower.endsWith('.svg')) {
    return 'image/svg+xml'
  }
  return 'image/jpeg'
}

const getImageDataUrl = async (zip, imagePath, cache) => {
  if (!imagePath) {
    return ''
  }
  if (cache.has(imagePath)) {
    return cache.get(imagePath)
  }
  const file = zip.file(imagePath)
  if (!file) {
    return ''
  }
  const base64 = await file.async('base64')
  const dataUrl = `data:${getMimeType(imagePath)};base64,${base64}`
  cache.set(imagePath, dataUrl)
  return dataUrl
}

const getPicturesFromSlide = (slideDoc, relMap) => {
  const nodes = Array.from(
    slideDoc.getElementsByTagName('p:pic').length
      ? slideDoc.getElementsByTagName('p:pic')
      : slideDoc.getElementsByTagName('pic'),
  )

  return nodes
    .map((node) => {
      const blip = getFirstTag(node, ['a:blip', 'blip'])
      const embed =
        blip?.getAttribute('r:embed') || blip?.getAttribute('embed') || ''
      const target = embed ? relMap.get(embed) : ''
      if (!target) {
        return null
      }
      const xfrm = getFirstTag(node, ['a:xfrm', 'xfrm'])
      const off = xfrm ? getFirstTag(xfrm, ['a:off', 'off']) : null
      const ext = xfrm ? getFirstTag(xfrm, ['a:ext', 'ext']) : null
      const x = Number(off?.getAttribute('x') || 0)
      const y = Number(off?.getAttribute('y') || 0)
      const cx = Number(ext?.getAttribute('cx') || 0)
      const cy = Number(ext?.getAttribute('cy') || 0)
      const area = cx * cy
      // embed = the r:id (e.g. "rId3") — needed to identify this pic in raw XML
      return { x, y, cx, cy, area, target, embed }
    })
    .filter(Boolean)
}

/**
 * Extract text shapes from a slide XML document.
 * Returns an array of objects: { text, xInch, yInch, wInch, hInch }
 * Only shapes that contain actual non-empty text are returned.
 * Background-covering shapes and shapes without text content are skipped.
 */
const getTextShapesFromSlide = (slideDoc, slideSize) => {
  // p:sp = shape (text box, title, content placeholder, etc.)
  const spNodes = Array.from(
    slideDoc.getElementsByTagName('p:sp').length
      ? slideDoc.getElementsByTagName('p:sp')
      : slideDoc.getElementsByTagName('sp'),
  )

  const slideAreaEmu = slideSize.cx * slideSize.cy
  const maxBackgroundAreaEmu = slideAreaEmu * MAX_BACKGROUND_RATIO

  const results = []

  for (const sp of spNodes) {
    // Get transform to find position/size
    const xfrm = getFirstTag(sp, ['a:xfrm', 'xfrm'])
    const off = xfrm ? getFirstTag(xfrm, ['a:off', 'off']) : null
    const ext = xfrm ? getFirstTag(xfrm, ['a:ext', 'ext']) : null

    const xEmu = Number(off?.getAttribute('x') || 0)
    const yEmu = Number(off?.getAttribute('y') || 0)
    const cxEmu = Number(ext?.getAttribute('cx') || 0)
    const cyEmu = Number(ext?.getAttribute('cy') || 0)
    const areaEmu = cxEmu * cyEmu

    // Skip full-slide background shapes
    if (areaEmu >= maxBackgroundAreaEmu) {
      continue
    }

    // Collect all text runs (a:r > a:t) within this shape
    const tNodes = Array.from(sp.getElementsByTagName('a:t').length
      ? sp.getElementsByTagName('a:t')
      : sp.getElementsByTagName('t'))

    const text = tNodes
      .map((t) => t.textContent || '')
      .join('')
      .trim()

    if (!text) {
      continue
    }

    results.push({
      text,
      xInch: xEmu / EMU_PER_INCH,
      yInch: yEmu / EMU_PER_INCH,
      wInch: cxEmu / EMU_PER_INCH,
      hInch: cyEmu / EMU_PER_INCH,
    })
  }

  return results
}

/**
 * Match extracted text shapes to master textbox placeholders by index order.
 * Text shapes are sorted by reading order (top→bottom, left→right).
 * The N-th text shape maps to the N-th textbox key.
 * Textbox keys beyond the number of extracted text shapes are left empty ('').
 *
 * @param {Array} textShapes  - from getTextShapesFromSlide(): [{ text, xInch, yInch, ... }]
 * @param {Array} textboxDefs - master textbox definitions: [{ key, x, y, w, h }]
 * @returns {Object} - { [key]: text }
 */
const matchTextToBoxes = (textShapes, textboxDefs) => {
  if (!textboxDefs.length) return {}

  // Sort text shapes by reading order: top-to-bottom first, then left-to-right
  // Use a row-tolerance of 0.5 inch so shapes on roughly the same line group together
  const ROW_TOLERANCE = 0.5
  const sorted = [...textShapes].sort((a, b) => {
    const rowDiff = a.yInch - b.yInch
    if (Math.abs(rowDiff) > ROW_TOLERANCE) return rowDiff
    return a.xInch - b.xInch
  })

  const result = {}
  for (let i = 0; i < textboxDefs.length; i++) {
    const key = textboxDefs[i].key
    // If a text shape exists at this index, use it; otherwise leave empty
    result[key] = sorted[i]?.text ?? ''
  }
  return result
}


const pickSlideImages = async (
  zip,
  pictures,
  slideSize,
  cache,
  imageCount = 2,
) => {
  const result = {
    beforeImage: '',
    middleImage: '',
    afterImage: '',
  }

  for (let i = 0; i < imageCount; i++) {
    result[`image_${i}`] = ''
  }

  if (!pictures.length) {
    return result
  }

  const slideArea = slideSize.cx * slideSize.cy
  const maxBackgroundArea = slideArea * MAX_BACKGROUND_RATIO
  const minImageArea = slideArea * MIN_IMAGE_RATIO

  const nonBackground = pictures.filter((pic) => pic.area < maxBackgroundArea)
  let candidates = nonBackground.filter((pic) => pic.area >= minImageArea)
  if (candidates.length < imageCount) {
    candidates = nonBackground.length ? nonBackground : pictures
  }

  candidates.sort((a, b) => b.area - a.area)
  const selected = candidates.slice(0, imageCount).sort((a, b) => a.x - b.x)

  for (let i = 0; i < selected.length; i++) {
    const dataUrl = await getImageDataUrl(zip, selected[i].target, cache)
    if (dataUrl) {
      result[`image_${i}`] = dataUrl
    }
  }

  if (imageCount === 3) {
    result.beforeImage = result.image_0
    result.middleImage = result.image_1
    result.afterImage = result.image_2
  } else {
    result.beforeImage = result.image_0
    result.afterImage = result.image_1
  }

  return result
}

export const importPptxSlides = async (
  file,
  {
    skipFirstSlides = 1,
    skipLastSlides = 1,
    imageCount = 2,
    textboxDefs = [],   // master textbox definitions: [{ key, x, y, w, h }]
  } = {},
) => {
  if (!file) {
    return { pairs: [], totalSlides: 0, importedSlides: 0, emptySlides: 0 }
  }

  const arrayBuffer = await file.arrayBuffer()
  const zip = await JSZip.loadAsync(arrayBuffer)
  const slidePaths = getSlidePaths(zip)

  const minSlides = skipFirstSlides + skipLastSlides + 1
  if (slidePaths.length < minSlides) {
    throw new Error(
      `PPTX must contain at least ${minSlides} slides to import.`,
    )
  }

  const slideSize = await getSlideSize(zip)
  const imageCache = new Map()
  const endIndex = skipLastSlides
    ? slidePaths.length - skipLastSlides
    : slidePaths.length
  const middleSlides = slidePaths.slice(skipFirstSlides, endIndex)

  const pairs = []
  let emptySlides = 0

  for (const slidePath of middleSlides) {
    const slideFile = zip.file(slidePath)
    if (!slideFile) {
      pairs.push({ beforeImage: '', middleImage: '', afterImage: '' })
      emptySlides += 1
      continue
    }

    const slideDoc = parseXml(await slideFile.async('text'))
    const relMap = await getSlideRelations(zip, getSlideRelsPath(slidePath))
    const pictures = getPicturesFromSlide(slideDoc, relMap)
    const pair = await pickSlideImages(
      zip,
      pictures,
      slideSize,
      imageCache,
      imageCount,
    )

    // Extract text shapes and match them to master textbox placeholders
    if (textboxDefs.length > 0) {
      const textShapes = getTextShapesFromSlide(slideDoc, slideSize)
      const extractedTexts = matchTextToBoxes(textShapes, textboxDefs)
      Object.assign(pair, extractedTexts)
    }

    if (!pair.beforeImage && !pair.afterImage && !pair.middleImage) {
      emptySlides += 1
    }

    pairs.push(pair)
  }

  return {
    pairs,
    totalSlides: slidePaths.length,
    importedSlides: middleSlides.length,
    emptySlides,
  }
}

/**
 * Extract EVERY image from every slide in the PPTX.
 * No slide skipping, no per-slide image limit.
 * Returns an array of extracted image objects sorted by slide order,
 * then by position (left-to-right within each slide).
 *
 * Each item: { id, slideIndex, slideNumber, imageIndex, dataUrl }
 */
export const extractAllImages = async (file) => {
  if (!file) return []

  const arrayBuffer = await file.arrayBuffer()
  const zip = await JSZip.loadAsync(arrayBuffer)
  const slidePaths = getSlidePaths(zip)
  const slideSize = await getSlideSize(zip)
  const imageCache = new Map()

  const results = []

  for (let si = 0; si < slidePaths.length; si++) {
    const slidePath = slidePaths[si]
    const slideFile = zip.file(slidePath)
    if (!slideFile) continue

    const slideDoc = parseXml(await slideFile.async('text'))
    const relMap = await getSlideRelations(zip, getSlideRelsPath(slidePath))
    const pictures = getPicturesFromSlide(slideDoc, relMap)

    if (!pictures.length) continue

    const slideArea = slideSize.cx * slideSize.cy
    const maxBackgroundArea = slideArea * MAX_BACKGROUND_RATIO

    // Filter out full-slide background images
    const usable = pictures.filter((pic) => pic.area < maxBackgroundArea)
    const toExtract = usable.length > 0 ? usable : pictures

    // Sort left-to-right, top-to-bottom by position
    toExtract.sort((a, b) => {
      const rowTolerance = slideSize.cy * 0.1
      const rowDiff = a.y - b.y
      if (Math.abs(rowDiff) > rowTolerance) return rowDiff
      return a.x - b.x
    })

    for (let ii = 0; ii < toExtract.length; ii++) {
      const pic = toExtract[ii]
      const dataUrl = await getImageDataUrl(zip, pic.target, imageCache)
      if (!dataUrl) continue
      results.push({
        id: `slide${si + 1}_img${ii + 1}`,
        slideIndex: si,
        slideNumber: si + 1,
        imageIndex: ii,
        dataUrl,
        // Extra metadata needed for context-preserving PPTX export
        embed: pic.embed,   // r:embed rId — identifies this pic in slide XML
        target: pic.target, // media file path inside the zip
        posCx: pic.cx,      // original width  in EMU (used for aspect ratio)
        posCy: pic.cy,      // original height in EMU
      })
    }
  }

  return results
}

// ── Context-preserving PPTX export ───────────────────────────────────────────

const SLIDE_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
const SLIDE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'

/**
 * Build a modified copy of a slide's raw XML where:
 *  - the featured image (matched by its r:embed value) is kept exactly
 *    as-is — same position, same size, same everything.
 *  - every other <p:pic> element is removed entirely.
 *  - all text shapes, backgrounds, and other elements are untouched.
 */
const buildSlideXmlForImage = (slideXmlText, embed) => {
  // Collect all <p:pic>…</p:pic> blocks with their positions in the string
  const picBlocks = []
  const picRe = /<p:pic[\s>][\s\S]*?<\/p:pic>/g
  let m
  while ((m = picRe.exec(slideXmlText)) !== null) {
    picBlocks.push({ full: m[0], start: m.index, end: m.index + m[0].length })
  }

  if (!picBlocks.length) return slideXmlText

  // Process in reverse so earlier indices stay valid after splicing
  let result = slideXmlText
  for (let i = picBlocks.length - 1; i >= 0; i--) {
    const { full, start, end } = picBlocks[i]
    const embedMatch = /r:embed="([^"]+)"/.exec(full) || /\bembed="([^"]+)"/.exec(full)
    const picEmbed = embedMatch ? embedMatch[1] : ''

    if (picEmbed !== embed) {
      // Non-featured image — remove it, leave a zero-width gap
      result = result.slice(0, start) + result.slice(end)
    }
    // Featured image: keep exactly as-is (no positional changes)
  }
  return result
}

/** Remove existing slide <Override> entries and add fresh ones. */
const rebuildContentTypes = (ctXml, numSlides) => {
  const newOverrides = Array.from({ length: numSlides }, (_, i) =>
    `\n  <Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="${SLIDE_CONTENT_TYPE}"/>`,
  ).join('')

  if (!ctXml) {
    return [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '  <Default Extension="xml" ContentType="application/xml"/>',
      newOverrides,
      '</Types>',
    ].join('\n')
  }

  const cleaned = ctXml.replace(
    /[ \t]*<Override[^>]+PartName="[^"]*\/ppt\/slides\/slide\d+\.xml"[^/]*\/>/g,
    '',
  )
  return cleaned.replace('</Types>', `${newOverrides}\n</Types>`)
}

/** Remove existing slide <Relationship> entries and add fresh ones. */
const rebuildPresentationRels = (relsXml, numSlides, slideRIds) => {
  const newRels = slideRIds
    .map(
      (rId, i) =>
        `\n  <Relationship Id="${rId}" Type="${SLIDE_REL_TYPE}" Target="slides/slide${i + 1}.xml"/>`,
    )
    .join('')

  if (!relsXml) {
    return [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      newRels,
      '</Relationships>',
    ].join('\n')
  }

  // Use lookahead so we match the Type attribute regardless of attribute order
  const escapedType = SLIDE_REL_TYPE.replace(/[.+?^${}()|[\]\\]/g, '\\$&')
  const cleaned = relsXml.replace(
    new RegExp(
      `[ \\t]*<Relationship(?=[^>]*\\bType="${escapedType}"[^>]*)[^>]*/>`,
      'g',
    ),
    '',
  )
  return cleaned.replace('</Relationships>', `${newRels}\n</Relationships>`)
}

/** Replace <p:sldIdLst>…</p:sldIdLst> with new entries referencing the output slides. */
const rebuildPresentationXml = (presXml, numSlides, slideRIds) => {
  if (!presXml) return ''
  const newEntries = slideRIds
    .map((rId, i) => `<p:sldId id="${256 + i}" r:id="${rId}"/>`)
    .join('')
  // Matches both self-closing (<p:sldIdLst/>) and full (<p:sldIdLst>…</p:sldIdLst>)
  return presXml.replace(
    /<p:sldIdLst\s*\/>|<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,
    `<p:sldIdLst>${newEntries}</p:sldIdLst>`,
  )
}

/**
 * Export selected images as a PPTX where every image gets its own slide
 * that is a faithful clone of the original slide:
 *   • all text boxes, shapes and backgrounds are preserved
 *   • every other image on that slide is removed
 *   • the featured image is resized to fill the slide (contain — never stretched)
 *
 * @param {File}   file     The original PPTX File object
 * @param {Array}  images   Selected image objects from extractAllImages
 *                          (must include: embed, posCx, posCy, slideIndex)
 * @param {string} fileName Output filename
 */
export const exportSlidesWithContext = async (
  file,
  images,
  fileName = 'Slides_With_Context.pptx',
) => {
  if (!file || !images.length) return

  const arrayBuffer = await file.arrayBuffer()
  const zip = await JSZip.loadAsync(arrayBuffer)
  const slidePaths = getSlidePaths(zip)
  const slideSize = await getSlideSize(zip)

  // Build one output slide per selected image (respects user's drag order)
  const outputSlides = []
  const slideXmlCache = new Map()

  for (const img of images) {
    const slidePath = slidePaths[img.slideIndex]
    if (!slidePath) continue

    if (!slideXmlCache.has(img.slideIndex)) {
      const sf = zip.file(slidePath)
      if (!sf) continue
      slideXmlCache.set(img.slideIndex, await sf.async('text'))
    }

    const slideXmlText = slideXmlCache.get(img.slideIndex)
    if (!slideXmlText) continue

    const modifiedXml = buildSlideXmlForImage(slideXmlText, img.embed)
    outputSlides.push({
      xmlText: modifiedXml,
      relsPath: getSlideRelsPath(slidePath),
    })
  }

  if (!outputSlides.length) return

  // Use high rIds to avoid collisions with existing non-slide relationships
  const slideRIds = outputSlides.map((_, i) => `rId${1000 + i}`)

  // ── Assemble output zip ───────────────────────────────────────────────────
  const outZip = new JSZip()

  const skipSet = new Set([
    'ppt/presentation.xml',
    'ppt/_rels/presentation.xml.rels',
    '[Content_Types].xml',
    ...slidePaths,
    ...slidePaths.map(getSlideRelsPath),
  ])

  // Copy all non-slide files (media, themes, layouts, masters …) unchanged
  const copyTasks = []
  zip.forEach((path, entry) => {
    if (skipSet.has(path)) return
    copyTasks.push(entry.async('arraybuffer').then((buf) => outZip.file(path, buf)))
  })
  await Promise.all(copyTasks)

  // Write modified slides and their rels
  const relsTextCache = new Map()
  for (let i = 0; i < outputSlides.length; i++) {
    outZip.file(`ppt/slides/slide${i + 1}.xml`, outputSlides[i].xmlText)

    const relsPath = outputSlides[i].relsPath
    if (!relsTextCache.has(relsPath)) {
      const rf = zip.file(relsPath)
      relsTextCache.set(relsPath, rf ? await rf.async('text') : null)
    }
    const relsText = relsTextCache.get(relsPath)
    if (relsText) {
      outZip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, relsText)
    }
  }

  // Rebuild the three files that reference slide count / paths
  const readText = (path) => zip.file(path)?.async('text') ?? Promise.resolve(null)
  const [presXml, presRelsXml, ctXml] = await Promise.all([
    readText('ppt/presentation.xml'),
    readText('ppt/_rels/presentation.xml.rels'),
    readText('[Content_Types].xml'),
  ])

  outZip.file(
    'ppt/presentation.xml',
    rebuildPresentationXml(presXml, outputSlides.length, slideRIds),
  )
  outZip.file(
    'ppt/_rels/presentation.xml.rels',
    rebuildPresentationRels(presRelsXml, outputSlides.length, slideRIds),
  )
  outZip.file('[Content_Types].xml', rebuildContentTypes(ctXml, outputSlides.length))

  // Download
  const blob = await outZip.generateAsync({
    type: 'blob',
    mimeType:
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = fileName
  a.click()
  URL.revokeObjectURL(url)
}

/**
 * Merge multiple PPTX files sequentially.
 * 
 * @param {File[]} files Array of PPTX file objects.
 * @param {function} onProgress Callback function for progress updates: (message) => void
 * @returns {Promise<Blob>} The merged PPTX presentation as a Blob.
 */
export const mergePptxFiles = async (files, onProgress) => {
  if (!files || files.length === 0) {
    throw new Error('No files selected for merging.')
  }

  onProgress?.('Loading the base presentation...')
  const baseFile = files[0]
  const baseArrayBuffer = await baseFile.arrayBuffer()
  const mergedZip = await JSZip.loadAsync(baseArrayBuffer)

  // Identify and delete base slide files so we start clean
  onProgress?.('Cleaning up base slides...')
  const baseSlidePaths = getSlidePaths(mergedZip)
  baseSlidePaths.forEach((path) => {
    mergedZip.remove(path)
    mergedZip.remove(getSlideRelsPath(path))
  })

  const mergedSlides = []

  // Iterate through all files and extract slides
  for (let fileIndex = 0; fileIndex < files.length; fileIndex++) {
    const file = files[fileIndex]
    onProgress?.(`Reading file ${fileIndex + 1} of ${files.length}: ${file.name}...`)
    
    const arrayBuffer = await file.arrayBuffer()
    const zip = await JSZip.loadAsync(arrayBuffer)
    const slidePaths = getSlidePaths(zip)

    for (let s = 0; s < slidePaths.length; s++) {
      onProgress?.(`Processing slide ${s + 1} of ${slidePaths.length} from file ${fileIndex + 1}...`)
      const slidePath = slidePaths[s]
      const slideFile = zip.file(slidePath)
      if (!slideFile) continue

      const slideXmlText = await slideFile.async('text')
      const relsPath = getSlideRelsPath(slidePath)
      const relsFile = zip.file(relsPath)
      let relsXmlText = relsFile ? await relsFile.async('text') : null

      if (relsXmlText) {
        const parser = new DOMParser()
        const relsDoc = parser.parseFromString(relsXmlText, 'application/xml')
        const relsList = Array.from(relsDoc.getElementsByTagName('Relationship'))

        for (const rel of relsList) {
          const target = rel.getAttribute('Target')
          const type = rel.getAttribute('Type')
          if (!target || !type) continue

          const isInternalLayoutOrMaster =
            type.endsWith('/slideLayout') ||
            type.endsWith('/slideMaster') ||
            type.endsWith('/notesLayout')

          if (!isInternalLayoutOrMaster) {
            const normalizedPath = normalizeTarget(target)
            const mediaFile = zip.file(normalizedPath)

            if (mediaFile) {
              const parts = target.split('/')
              const targetFilename = parts.pop()
              const dirPath = parts.join('/') // e.g. "../media" or "../charts"

              const newTargetName = `p${fileIndex}_s${s}_${targetFilename}`
              const newTarget = `${dirPath}/${newTargetName}`
              const newZipPath = normalizeTarget(newTarget)

              // Copy file to merged zip
              const mediaData = await mediaFile.async('arraybuffer')
              mergedZip.file(newZipPath, mediaData)

              // Update relationship target
              rel.setAttribute('Target', newTarget)
            }
          }
        }
        relsXmlText = new XMLSerializer().serializeToString(relsDoc)
      }

      mergedSlides.push({
        xmlText: slideXmlText,
        relsText: relsXmlText,
      })
    }
  }

  onProgress?.('Assembling presentation structure...')
  const slideRIds = mergedSlides.map((_, i) => `rId${1000 + i}`)

  for (let i = 0; i < mergedSlides.length; i++) {
    const slideNum = i + 1
    mergedZip.file(`ppt/slides/slide${slideNum}.xml`, mergedSlides[i].xmlText)
    if (mergedSlides[i].relsText) {
      mergedZip.file(`ppt/slides/_rels/slide${slideNum}.xml.rels`, mergedSlides[i].relsText)
    }
  }

  // Rebuild presentation xml, presentation rels, and content types
  const readText = (path) => mergedZip.file(path)?.async('text') ?? Promise.resolve(null)
  const [presXml, presRelsXml, ctXml] = await Promise.all([
    readText('ppt/presentation.xml'),
    readText('ppt/_rels/presentation.xml.rels'),
    readText('[Content_Types].xml'),
  ])

  mergedZip.file(
    'ppt/presentation.xml',
    rebuildPresentationXml(presXml, mergedSlides.length, slideRIds)
  )
  mergedZip.file(
    'ppt/_rels/presentation.xml.rels',
    rebuildPresentationRels(presRelsXml, mergedSlides.length, slideRIds)
  )
  mergedZip.file('[Content_Types].xml', rebuildContentTypes(ctXml, mergedSlides.length))

  onProgress?.('Generating final presentation package...')
  const blob = await mergedZip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })

  return blob
}

