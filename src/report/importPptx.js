import JSZip from 'jszip'

const DEFAULT_SLIDE_SIZE = { cx: 12192000, cy: 6858000 }
const MAX_BACKGROUND_RATIO = 0.9
const MIN_IMAGE_RATIO = 0.05

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
      return { x, y, cx, cy, area, target }
    })
    .filter(Boolean)
}

const pickSlideImages = async (
  zip,
  pictures,
  slideSize,
  cache,
  imageCount = 2,
) => {
  if (!pictures.length) {
    return {
      beforeImage: '',
      middleImage: '',
      afterImage: '',
    }
  }

  const slideArea = slideSize.cx * slideSize.cy
  const maxBackgroundArea = slideArea * MAX_BACKGROUND_RATIO
  const minImageArea = slideArea * MIN_IMAGE_RATIO

  const nonBackground = pictures.filter((pic) => pic.area < maxBackgroundArea)
  let candidates = nonBackground.filter((pic) => pic.area >= minImageArea)
  if (candidates.length < 2) {
    candidates = nonBackground.length ? nonBackground : pictures
  }

  candidates.sort((a, b) => b.area - a.area)
  const useThree = imageCount === 3
  const targetCount = useThree ? 3 : 2
  const selected = candidates.slice(0, targetCount).sort((a, b) => a.x - b.x)

  const beforeImage = selected[0]
    ? await getImageDataUrl(zip, selected[0].target, cache)
    : ''
  const middleImage = useThree && selected[1]
    ? await getImageDataUrl(zip, selected[1].target, cache)
    : ''
  const afterImage = (useThree ? selected[2] : selected[1])
    ? await getImageDataUrl(zip, useThree ? selected[2].target : selected[1].target, cache)
    : ''

  return {
    beforeImage: beforeImage || '',
    middleImage: middleImage || '',
    afterImage: afterImage || '',
  }
}

export const importPptxSlides = async (
  file,
  { skipFirstSlides = 1, skipLastSlides = 1, imageCount = 2 } = {},
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
