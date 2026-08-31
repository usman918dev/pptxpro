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

const isExternalRelationshipTarget = (target) =>
  /^(https?:|mailto:|ftp:|file:|data:)/i.test(target || '')

const normalizePathSegments = (parts) => {
  const out = []
  for (const part of parts) {
    if (!part || part === '.') continue
    if (part === '..') {
      if (out.length > 0) out.pop()
      continue
    }
    out.push(part)
  }
  return out
}

const resolveRelationshipTargetPath = (sourcePartPath, target) => {
  if (!target || isExternalRelationshipTarget(target)) return ''

  const targetPath = target.split('?')[0].split('#')[0]
  if (!targetPath) return ''

  if (targetPath.startsWith('/')) {
    return normalizePathSegments(targetPath.slice(1).split('/')).join('/')
  }

  const sourceParts = sourcePartPath.split('/')
  sourceParts.pop()
  const merged = normalizePathSegments([...sourceParts, ...targetPath.split('/')])
  return merged.join('/')
}

const toRelationshipTarget = (sourcePartPath, absoluteTargetPath) => {
  const from = normalizePathSegments(sourcePartPath.split('/'))
  from.pop() // source part directory
  const to = normalizePathSegments(absoluteTargetPath.split('/'))

  let common = 0
  while (common < from.length && common < to.length && from[common] === to[common]) {
    common++
  }

  const up = new Array(from.length - common).fill('..')
  const down = to.slice(common)
  const rel = [...up, ...down].join('/')
  return rel || '.'
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

  let baseSlideLayoutTarget = '../slideLayouts/slideLayout1.xml'
  const baseSlidePathsForLayout = getSlidePaths(mergedZip)
  if (baseSlidePathsForLayout.length > 0) {
    const firstBaseSlidePath = baseSlidePathsForLayout[0]
    const firstBaseRelsFile = mergedZip.file(getSlideRelsPath(firstBaseSlidePath))
    if (firstBaseRelsFile) {
      try {
        const baseRelsDoc = parseXml(await firstBaseRelsFile.async('text'))
        const relNodes = Array.from(baseRelsDoc.getElementsByTagName('Relationship'))
        const layoutRel = relNodes.find((rel) =>
          (rel.getAttribute('Type') || '').endsWith('/slideLayout'),
        )
        if (layoutRel?.getAttribute('Target')) {
          baseSlideLayoutTarget = layoutRel.getAttribute('Target')
        }
      } catch {
        // Fall back to default layout target if base rels cannot be parsed.
      }
    }
  }

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
        const relsDoc = parseXml(relsXmlText)
        const relsList = Array.from(relsDoc.getElementsByTagName('Relationship'))
        const relationshipsToRemove = []

        for (const rel of relsList) {
          const target = rel.getAttribute('Target')
          const type = rel.getAttribute('Type')
          const targetMode = rel.getAttribute('TargetMode')
          if (!target || !type) continue

          if (type.endsWith('/slideLayout')) {
            rel.setAttribute('Target', baseSlideLayoutTarget)
            continue
          }

          if (type.endsWith('/notesSlide')) {
            // Notes parts have their own relationship trees and content-type overrides.
            // Dropping them keeps merged output valid for PowerPoint.
            relationshipsToRemove.push(rel)
            continue
          }

          if (targetMode === 'External' || isExternalRelationshipTarget(target)) {
            continue
          }

          if (!type.endsWith('/image')) {
            continue
          }

          const resolvedSourcePath = resolveRelationshipTargetPath(slidePath, target)
          if (!resolvedSourcePath) continue

          const mediaFile = zip.file(resolvedSourcePath)

          if (mediaFile) {
            const sourcePathParts = resolvedSourcePath.split('/')
            const sourceFilename = sourcePathParts.pop()
            const sourceDir = sourcePathParts.join('/')

            const newTargetName = `p${fileIndex}_s${s}_${sourceFilename}`
            const newZipPath = `${sourceDir}/${newTargetName}`

            // Copy image file to merged zip
            const mediaData = await mediaFile.async('arraybuffer')
            mergedZip.file(newZipPath, mediaData)

            // Relationship targets are relative to the slide part path.
            rel.setAttribute('Target', toRelationshipTarget(slidePath, newZipPath))
          }
        }

        relationshipsToRemove.forEach((rel) => rel.parentNode?.removeChild(rel))
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

// ─────────────────────────────────────────────────────────────────────────────
// First / Last Slide Import Editor — public API
// ─────────────────────────────────────────────────────────────────────────────

const IMAGE_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

// ── Scheme colour map ────────────────────────────────────────────────────────
// Maps the 14 standard OOXML theme colour roles to reasonable hex approximations
// used only for the in-browser preview. The raw XML (and therefore the downloaded
// PPTX) always preserves the original colour references exactly.
const SCHEME_COLOR_MAP = {
  dk1:      '000000', tx1:     '000000',
  lt1:      'FFFFFF', bg1:     'FFFFFF',
  dk2:      '1F3864', tx2:     '1F3864',
  lt2:      'D9E2F3', bg2:     'D9E2F3',
  accent1:  '4472C4',
  accent2:  'ED7D31',
  accent3:  'A9D18E',
  accent4:  'FFC000',
  accent5:  '5A96D6',
  accent6:  '70AD47',
  hlink:    '0563C1',
  folHlink: '954F72',
}

/**
 * Resolve a hex colour from an <a:solidFill> element.
 * Handles direct <a:srgbClr> values and <a:schemeClr> theme references.
 * Returns an empty string when no colour can be determined.
 */
const resolveSolidFillColor = (solidFillEl) => {
  if (!solidFillEl) return ''
  const srgb = solidFillEl.getElementsByTagName('a:srgbClr')[0]
  if (srgb) return srgb.getAttribute('val') || ''
  const scheme = solidFillEl.getElementsByTagName('a:schemeClr')[0]
  if (scheme) return SCHEME_COLOR_MAP[scheme.getAttribute('val') || ''] || ''
  return ''
}

// ── Table extraction ──────────────────────────────────────────────────────────

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

    // p:graphicFrame uses p:xfrm (not a:xfrm) for its transform element
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

        let color = ''
        if (rPr) {
          color = resolveSolidFillColor(rPr.getElementsByTagName('a:solidFill')[0])
        }

        const pPr =
          tc.getElementsByTagName('a:pPr')[0] ||
          tc.getElementsByTagName('pPr')[0]
        const algn = pPr?.getAttribute('algn') || 'l'
        const alignMap = { l: 'left', ctr: 'center', r: 'right', just: 'justify' }

        const tcPr =
          tc.getElementsByTagName('a:tcPr')[0] ||
          tc.getElementsByTagName('tcPr')[0]

        // Cell background fill — handles both direct hex and theme colours
        const fillColor = resolveSolidFillColor(
          tcPr?.getElementsByTagName('a:solidFill')[0] ?? null
        )

        // Cell border colour — try left border (a:lnL) first, then top (a:lnT)
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
      xEmu,
      yEmu,
      cxEmu,
      cyEmu,
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

// ── Enhanced text shape extraction (with font info + EMU coords) ──────────────

const getFullTextShapesFromSlide = (slideDoc, slideSize) => {
  const spNodes = Array.from(
    slideDoc.getElementsByTagName('p:sp').length
      ? slideDoc.getElementsByTagName('p:sp')
      : slideDoc.getElementsByTagName('sp'),
  )
  const slideAreaEmu = slideSize.cx * slideSize.cy
  const maxBgAreaEmu = slideAreaEmu * MAX_BACKGROUND_RATIO
  const slideWidthInch = slideSize.cx / EMU_PER_INCH
  const results = []

  for (const sp of spNodes) {
    const xfrm = getFirstTag(sp, ['a:xfrm', 'xfrm'])
    const off = xfrm ? getFirstTag(xfrm, ['a:off', 'off']) : null
    const ext = xfrm ? getFirstTag(xfrm, ['a:ext', 'ext']) : null

    const xEmu = Number(off?.getAttribute('x') || 0)
    const yEmu = Number(off?.getAttribute('y') || 0)
    const cxEmu = Number(ext?.getAttribute('cx') || 0)
    const cyEmu = Number(ext?.getAttribute('cy') || 0)
    if (cxEmu * cyEmu >= maxBgAreaEmu) continue

    const tNodes = Array.from(
      sp.getElementsByTagName('a:t').length
        ? sp.getElementsByTagName('a:t')
        : sp.getElementsByTagName('t'),
    )
    const text = tNodes.map((t) => t.textContent || '').join('').trim()
    if (!text) continue

    const rPr =
      sp.getElementsByTagName('a:rPr')[0] ||
      sp.getElementsByTagName('rPr')[0]
    const bold = rPr?.getAttribute('b') === '1'
    const sz = rPr ? Number(rPr.getAttribute('sz') || 0) : 0
    const ptSize = sz > 0 ? sz / 100 : 14
    // cqw formula matching SlideCanvas: fontPt * 0.104
    const fontSizePct = Math.max(ptSize * 0.104, 0.8)

    let fontFace = 'Calibri'
    if (rPr) {
      const latin = rPr.getElementsByTagName('a:latin')[0]
      if (latin) fontFace = latin.getAttribute('typeface') || 'Calibri'
    }

    let color = ''
    if (rPr) {
      color = resolveSolidFillColor(rPr.getElementsByTagName('a:solidFill')[0])
    }

    const pPr =
      sp.getElementsByTagName('a:pPr')[0] ||
      sp.getElementsByTagName('pPr')[0]
    const algn = pPr?.getAttribute('algn') || 'l'
    const alignMap = { l: 'left', ctr: 'center', r: 'right', just: 'justify' }

    results.push({
      id: `txt_${xEmu}_${yEmu}`,
      xEmu,
      yEmu,
      xPct: (xEmu / slideSize.cx) * 100,
      yPct: (yEmu / slideSize.cy) * 100,
      wPct: (cxEmu / slideSize.cx) * 100,
      hPct: (cyEmu / slideSize.cy) * 100,
      text,
      fontFace,
      fontSizePct,
      bold,
      color,
      align: alignMap[algn] || 'left',
    })
  }

  return results
}

// ── Apply text / table edits to raw slide XML ─────────────────────────────────

/**
 * Applies user-edited text and table-cell content to the raw PPTX slide XML.
 * Uses DOM manipulation then re-serialises, preserving all formatting/structure.
 *
 * @param {string}  rawXml     - Original slide XML string
 * @param {{ texts: Array, tables: Array }} editedData
 * @returns {string} Modified XML string
 */
export const applyEditsToSlideXml = (rawXml, editedData) => {
  const parser = new DOMParser()
  const doc = parser.parseFromString(rawXml, 'application/xml')
  const { texts = [], tables = [] } = editedData

  // ── Update text shapes ────────────────────────────────────────────────────
  const textByKey = new Map()
  texts.forEach((t) => textByKey.set(`${t.xEmu}_${t.yEmu}`, t.text))

  const spNodes = Array.from(
    doc.getElementsByTagName('p:sp').length
      ? doc.getElementsByTagName('p:sp')
      : doc.getElementsByTagName('sp'),
  )

  for (let sIdx = 0; sIdx < spNodes.length; sIdx++) {
    const sp = spNodes[sIdx]
    const xfrm =
      sp.getElementsByTagName('p:xfrm')[0] ||
      sp.getElementsByTagName('a:xfrm')[0] ||
      sp.getElementsByTagName('xfrm')[0]
    const off = xfrm
      ? xfrm.getElementsByTagName('a:off')[0] ||
        xfrm.getElementsByTagName('off')[0]
      : null

    const xEmu = off ? Number(off.getAttribute('x') || 0) : 0
    const yEmu = off ? Number(off.getAttribute('y') || 0) : 0

    // 1. Try spIndex match
    let foundElem = texts.find((t) => t.spIndex === sIdx)

    // 2. Try exact coordinate key match
    if (!foundElem) {
      foundElem = texts.find((t) => t.xEmu === xEmu && t.yEmu === yEmu)
    }

    // 3. Try fuzzy coordinate match
    if (!foundElem) {
      foundElem = texts.find(
        (t) => Math.abs(t.xEmu - xEmu) < 50000 && Math.abs(t.yEmu - yEmu) < 50000,
      )
    }

    // 4. Index order fallback for text shapes
    if (!foundElem && texts[sIdx]) {
      foundElem = texts[sIdx]
    }

    if (!foundElem || foundElem.text === undefined) continue
    foundElem._matched = true

    const tNodes = Array.from(
      sp.getElementsByTagName('a:t').length
        ? sp.getElementsByTagName('a:t')
        : sp.getElementsByTagName('t'),
    )
    if (tNodes.length === 0) continue
    tNodes[0].textContent = foundElem.text || ''
    for (let i = 1; i < tNodes.length; i++) tNodes[i].textContent = ''
  }

  // ── Update table cells ────────────────────────────────────────────────────
  const frameNodes = Array.from(
    doc.getElementsByTagName('p:graphicFrame').length
      ? doc.getElementsByTagName('p:graphicFrame')
      : doc.getElementsByTagName('graphicFrame'),
  )

  let tableIndex = 0
  for (const frame of frameNodes) {
    const tbl =
      frame.getElementsByTagName('a:tbl')[0] ||
      frame.getElementsByTagName('tbl')[0]
    if (!tbl) continue

    const xfrm =
      frame.getElementsByTagName('p:xfrm')[0] ||
      frame.getElementsByTagName('a:xfrm')[0] ||
      frame.getElementsByTagName('xfrm')[0]
    const off = xfrm
      ? xfrm.getElementsByTagName('a:off')[0] ||
        xfrm.getElementsByTagName('off')[0]
      : null

    const xEmu = off ? Number(off.getAttribute('x') || 0) : 0
    const yEmu = off ? Number(off.getAttribute('y') || 0) : 0

    let editedTable = tables.find(
      (t) => Math.abs(t.xEmu - xEmu) < 50000 && Math.abs(t.yEmu - yEmu) < 50000,
    )
    if (!editedTable && tables[tableIndex]) {
      editedTable = tables[tableIndex]
    }
    tableIndex++
    if (!editedTable) continue

    const trNodes = Array.from(
      tbl.getElementsByTagName('a:tr').length
        ? tbl.getElementsByTagName('a:tr')
        : tbl.getElementsByTagName('tr'),
    )
    trNodes.forEach((tr, ri) => {
      const editedRow = editedTable.rows[ri]
      if (!editedRow) return
      const tcNodes = Array.from(
        tr.getElementsByTagName('a:tc').length
          ? tr.getElementsByTagName('a:tc')
          : tr.getElementsByTagName('tc'),
      )
      tcNodes.forEach((tc, ci) => {
        const cell = editedRow[ci]
        if (!cell) return
        const tNodes = Array.from(
          tc.getElementsByTagName('a:t').length
            ? tc.getElementsByTagName('a:t')
            : tc.getElementsByTagName('t'),
        )
        if (tNodes.length > 0) {
          tNodes[0].textContent = cell.text
          for (let i = 1; i < tNodes.length; i++) tNodes[i].textContent = ''
        } else if (cell.text) {
          let txBody =
            tc.getElementsByTagName('a:txBody')[0] ||
            tc.getElementsByTagName('txBody')[0]
          if (!txBody) {
            txBody = doc.createElementNS(
              'http://schemas.openxmlformats.org/drawingml/2006/main',
              'a:txBody',
            )
            tc.appendChild(txBody)
          }
          let p =
            txBody.getElementsByTagName('a:p')[0] ||
            txBody.getElementsByTagName('p')[0]
          if (!p) {
            p = doc.createElementNS(
              'http://schemas.openxmlformats.org/drawingml/2006/main',
              'a:p',
            )
            txBody.appendChild(p)
          }
          const r = doc.createElementNS(
            'http://schemas.openxmlformats.org/drawingml/2006/main',
            'a:r',
          )
          const t = doc.createElementNS(
            'http://schemas.openxmlformats.org/drawingml/2006/main',
            'a:t',
          )
          t.textContent = cell.text
          r.appendChild(t)
          p.appendChild(r)
        }
      })
    })
  }

  const serializer = new XMLSerializer()
  let result = serializer.serializeToString(doc)
  if (!result.startsWith('<?xml')) {
    result = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + result
  }
  return result
}

// ── Import first & last slide data from a PPTX file ──────────────────────────

/**
 * Extracts rich editable data for the first and last slides of a PPTX.
 * Each slide object contains images (with dataUrl), texts (with font info),
 * tables (with cell styling), plus the raw XML and rels needed for PPTX output.
 *
 * Only works on the master route; called manually from App.jsx after upload.
 */
export const importFirstLastSlideData = async (file) => {
  if (!file) return { firstSlide: null, lastSlide: null }

  const arrayBuffer = await file.arrayBuffer()
  const zip = await JSZip.loadAsync(arrayBuffer)
  const slidePaths = getSlidePaths(zip)
  const slideSize = await getSlideSize(zip)
  const imageCache = new Map()

  if (slidePaths.length === 0) return { firstSlide: null, lastSlide: null }

  const extractOne = async (slidePath) => {
    const slideFile = zip.file(slidePath)
    if (!slideFile) return null

    const rawXml = await slideFile.async('text')
    const slideDoc = parseXml(rawXml)

    const relsPath = getSlideRelsPath(slidePath)
    const relsFile = zip.file(relsPath)
    const rawRelsXml = relsFile ? await relsFile.async('text') : null

    // Full rel map (all types, not just images)
    const fullRelMap = new Map()
    if (rawRelsXml) {
      const relsDoc = parseXml(rawRelsXml)
      Array.from(relsDoc.getElementsByTagName('Relationship')).forEach((node) => {
        const id = node.getAttribute('Id')
        const target = node.getAttribute('Target') || ''
        const type = node.getAttribute('Type') || ''
        if (id) fullRelMap.set(id, { target, type })
      })
    }

    // Image-only rel map for getPicturesFromSlide
    const imageRelMap = new Map()
    fullRelMap.forEach((val, id) => {
      if (/image/i.test(val.type)) {
        imageRelMap.set(id, normalizeTarget(val.target))
      }
    })

    const pictures = getPicturesFromSlide(slideDoc, imageRelMap)
    const slideArea = slideSize.cx * slideSize.cy
    const maxBgArea = slideArea * MAX_BACKGROUND_RATIO

    // Background (large image covering most of slide)
    let backgroundDataUrl = ''
    const bgPics = pictures.filter((p) => p.area >= maxBgArea)
    if (bgPics.length > 0) {
      backgroundDataUrl = await getImageDataUrl(zip, bgPics[0].target, imageCache)
    }

    // Content images
    const contentPics = pictures.filter((p) => p.area < maxBgArea)
    const images = []
    for (let i = 0; i < contentPics.length; i++) {
      const pic = contentPics[i]
      const dataUrl = await getImageDataUrl(zip, pic.target, imageCache)
      images.push({
        id: `img_${pic.embed}_${i}`,
        embed: pic.embed,
        target: pic.target,
        xEmu: pic.x,
        yEmu: pic.y,
        cxEmu: pic.cx,
        cyEmu: pic.cy,
        xPct: (pic.x / slideSize.cx) * 100,
        yPct: (pic.y / slideSize.cy) * 100,
        wPct: (pic.cx / slideSize.cx) * 100,
        hPct: (pic.cy / slideSize.cy) * 100,
        dataUrl,
        originalDataUrl: dataUrl,
        replaced: false,
      })
    }

    // Text shapes (enhanced)
    const texts = getFullTextShapesFromSlide(slideDoc, slideSize)

    // Tables
    const tables = getTablesFromSlide(slideDoc, slideSize)

    // Collect media bytes for every image relationship
    const mediaFiles = new Map()
    for (const [, rel] of fullRelMap) {
      if (!/image/i.test(rel.type)) continue
      const norm = normalizeTarget(rel.target)
      if (mediaFiles.has(norm)) continue
      const mf = zip.file(norm)
      if (mf) mediaFiles.set(norm, await mf.async('arraybuffer'))
    }

    return {
      images,
      texts,
      tables,
      backgroundDataUrl,
      rawXml,
      rawRelsXml,
      mediaFiles,
      fullRelMap,
    }
  }

  const firstPath = slidePaths[0]
  const lastPath = slidePaths[slidePaths.length - 1]

  const [firstSlide, lastSlide] = await Promise.all([
    extractOne(firstPath),
    slidePaths.length > 1 ? extractOne(lastPath) : Promise.resolve(null),
  ])

  return { firstSlide, lastSlide }
}

// ── Post-process PptxGenJS blob — inject imported + edited slides ─────────────

/**
 * After PptxGenJS generates the PPTX blob, this function replaces slide 1 and/or
 * the last slide with the imported (and user-edited) XML from the original PPTX.
 *
 * Image embed IDs are remapped to new unique IDs added to the PptxGenJS rels
 * (which contain the correct slide-layout / slide-master references).
 *
 * @param {Blob}  pptxBlob        - Blob from pptx.write({ outputType: 'blob' })
 * @param {{ importedSlide, editedSlide } | null} firstSlideOpts
 * @param {{ importedSlide, editedSlide } | null} lastSlideOpts
 * @returns {Promise<Blob>}
 */
export const postProcessPptxWithImportedSlides = async (
  pptxBlob,
  firstSlideOpts,
  lastSlideOpts,
) => {
  if (!firstSlideOpts && !lastSlideOpts) return pptxBlob

  const zip = await JSZip.loadAsync(await pptxBlob.arrayBuffer())

  // Count slides in the generated PPTX
  let slideCount = 0
  zip.forEach((path) => {
    if (/^ppt\/slides\/slide\d+\.xml$/i.test(path)) slideCount++
  })

  const parser = new DOMParser()
  const serializer = new XMLSerializer()

  const processSlide = async (slideNum, importedSlide, editedSlide) => {
    const slidePath = `ppt/slides/slide${slideNum}.xml`
    const relsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`

    // Must have PptxGenJS rels for valid layout/master references
    const pptxRelsXml = await zip.file(relsPath)?.async('text')
    if (!pptxRelsXml) return

    // Apply user edits (text + table) to the imported raw XML
    const effectiveTexts = editedSlide?.texts ?? importedSlide.texts
    const effectiveTables = editedSlide?.tables ?? importedSlide.tables
    let editedXml = applyEditsToSlideXml(importedSlide.rawXml, {
      texts: effectiveTexts,
      tables: effectiveTables,
    })

    // Parse PptxGenJS rels — keep layout/master refs, remove its image rels
    const pptxRelsDoc = parser.parseFromString(pptxRelsXml, 'application/xml')
    const relsRoot =
      pptxRelsDoc.getElementsByTagName('Relationships')[0] ||
      pptxRelsDoc.documentElement

    Array.from(pptxRelsDoc.getElementsByTagName('Relationship')).forEach((rel) => {
      if (/image/i.test(rel.getAttribute('Type') || '')) {
        relsRoot.removeChild(rel)
      }
    })

    // Find max existing rId to avoid collisions
    let maxRIdNum = 100
    Array.from(pptxRelsDoc.getElementsByTagName('Relationship')).forEach((rel) => {
      const m = (rel.getAttribute('Id') || '').match(/rId(\d+)/)
      if (m) maxRIdNum = Math.max(maxRIdNum, Number(m[1]))
    })

    const effectiveImages = editedSlide?.images ?? importedSlide.images
    const rIdRemap = new Map()
    let mediaCounter = 0

    for (const [oldRId, rel] of (importedSlide.fullRelMap || new Map())) {
      if (!/image/i.test(rel.type)) continue

      maxRIdNum++
      const newRId = `rId${maxRIdNum}`
      rIdRemap.set(oldRId, newRId)

      const normalizedTarget = normalizeTarget(rel.target)
      const editedImg = effectiveImages.find((img) => img.embed === oldRId)
      const isReplaced = editedImg?.replaced && editedImg?.dataUrl &&
        editedImg.dataUrl !== editedImg.originalDataUrl

      let mediaPath, mediaBuf

      if (isReplaced) {
        const dataUrl = editedImg.dataUrl
        const b64 = dataUrl.replace(/^data:[^;]+;base64,/, '')
        const binary = atob(b64)
        const bytes = new Uint8Array(binary.length)
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
        mediaBuf = bytes.buffer
        const ext = dataUrl.includes('image/png') ? 'png'
          : dataUrl.includes('image/gif') ? 'gif' : 'jpg'
        mediaPath = `ppt/media/imp_s${slideNum}_r${mediaCounter++}.${ext}`
      } else if (importedSlide.mediaFiles?.has(normalizedTarget)) {
        mediaBuf = importedSlide.mediaFiles.get(normalizedTarget)
        const origName = normalizedTarget.split('/').pop()
        mediaPath = `ppt/media/imp_s${slideNum}_${mediaCounter++}_${origName}`
      } else {
        continue
      }

      zip.file(mediaPath, mediaBuf)

      // Rel target is relative from the slide part: ../media/<filename>
      const relTarget = `../media/${mediaPath.split('/').pop()}`
      const newRelEl = pptxRelsDoc.createElementNS(
        'http://schemas.openxmlformats.org/package/2006/relationships',
        'Relationship',
      )
      newRelEl.setAttribute('Id', newRId)
      newRelEl.setAttribute('Type', IMAGE_REL_TYPE)
      newRelEl.setAttribute('Target', relTarget)
      relsRoot.appendChild(newRelEl)
    }

    // Remap r:embed values in the slide XML
    rIdRemap.forEach((newRId, oldRId) => {
      editedXml = editedXml
        .replace(new RegExp(`r:embed="${oldRId}"`, 'g'), `r:embed="${newRId}"`)
        .replace(new RegExp(`\\bembed="${oldRId}"`, 'g'), `embed="${newRId}"`)
    })

    zip.file(slidePath, editedXml)

    let newRelsXml = serializer.serializeToString(pptxRelsDoc)
    if (!newRelsXml.startsWith('<?xml')) {
      newRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + newRelsXml
    }
    zip.file(relsPath, newRelsXml)
  }

  if (firstSlideOpts) {
    await processSlide(1, firstSlideOpts.importedSlide, firstSlideOpts.editedSlide)
  }
  if (lastSlideOpts && slideCount > 0) {
    await processSlide(slideCount, lastSlideOpts.importedSlide, lastSlideOpts.editedSlide)
  }

  return zip.generateAsync({
    type: 'blob',
    mimeType:
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
}

