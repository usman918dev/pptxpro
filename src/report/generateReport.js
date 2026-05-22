import PptxGenJS from 'pptxgenjs'

const MASTER_TITLE = 'CLEAN_PUNJAB_MASTER'
const SLIDE_TITLE = 'Plots Cleaning-Activity'
const FIRST_SLIDE_URL = '/first-slide.png'
const SECOND_SLIDE_URL = ''
const LAST_SLIDE_URL = '/last-slide.png'
const FILE_NAME_PREFIX = 'Plots  Cleaning-Activity_Tehsil Kamoke'
const SLIDE_SIZE = { w: 13.333, h: 7.5 }

const DATE_BOX = { x: 10.7, y: 6.88, w: 2.4, h: 0.4 }

// Matches the slide preview percentages: 7%, 35.333%, 40%, 53.333% on 13.333x7.5
const LEFT_BOX = { x: 0.93, y: 2.65, w: 5.33, h: 4.0 }
const RIGHT_BOX = { x: 7.1, y: 2.65, w: 5.33, h: 4.0 }

const formatDateForFile = (date) => {
  const day = String(date.getDate()).padStart(2, '0')
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const year = date.getFullYear()
  return `${day}-${month}-${year}`
}

const formatDateForSlide = (date) => {
  const day = String(date.getDate()).padStart(2, '0')
  const year = date.getFullYear()
  const monthNames = [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December',
  ]
  const month = monthNames[date.getMonth()]
  return `${day} ${month} ${year}`
}

const getDateWithOffset = (offsetDays = 0) => {
  const date = new Date()
  if (Number.isFinite(offsetDays) && offsetDays !== 0) {
    date.setDate(date.getDate() + offsetDays)
  }
  return date
}

const normalizeImage = (image) => {
  if (!image || typeof image !== 'string') {
    return null
  }

  const trimmed = image.trim()
  if (!trimmed) {
    return null
  }

  if (trimmed.startsWith('data:')) {
    return { data: trimmed }
  }

  if (/^https?:\/\//i.test(trimmed)) {
    return { path: trimmed }
  }

  if (
    trimmed.startsWith('/') ||
    trimmed.startsWith('./') ||
    trimmed.startsWith('../')
  ) {
    return { path: trimmed }
  }

  return { data: `data:image/jpeg;base64,${trimmed}` }
}

const defineMaster = (
  pptx,
  { backgroundImage, slideTitle = SLIDE_TITLE, masterTitle = MASTER_TITLE } = {},
) => {
  const objects = []

  const backgroundSource = normalizeImage(backgroundImage)
  if (backgroundSource) {
    objects.push({
      image: {
        ...backgroundSource,
        x: 0,
        y: 0,
        ...SLIDE_SIZE,
      },
    })
  } else {
    objects.push(
      {
        shape: pptx.ShapeType.rect,
        options: {
          x: 0,
          y: 0,
          w: 13.333,
          h: 0.7,
          fill: { color: 'FFFFFF' },
          line: { color: 'C7C7C7', width: 0.5 },
        },
      },
      {
        text: slideTitle,
        options: {
          x: 0.5,
          y: 0.12,
          w: 12.333,
          h: 0.5,
          fontFace: 'Calibri',
          fontSize: 28,
          color: '111111',
          bold: true,
          align: 'center',
          valign: 'middle',
        },
      },
      {
        shape: pptx.ShapeType.rect,
        options: {
          x: 0,
          y: 7.02,
          w: 13.333,
          h: 0.12,
          fill: { color: 'D92323' },
          line: { color: 'D92323' },
        },
      },
      {
        shape: pptx.ShapeType.rect,
        options: {
          x: 0,
          y: 7.14,
          w: 13.333,
          h: 0.22,
          fill: { color: '0B7A38' },
          line: { color: '0B7A38' },
        },
      },
    )
  }

  pptx.defineSlideMaster({
    title: masterTitle,
    objects,
  })
}

export const generateReport = async (
  pairs,
  {
    fileName,
    fileNamePrefix,
    backgroundImage,
    firstSlideImage,
    secondSlideImage,
    lastSlideImage,
    slideTitle,
    masterTitle,
    leftBox,
    middleBox,
    rightBox,
    dateSlide,
    dateBox,
    dateColor,
    dateAlign,
    dateFontSize,
    dateBold,
    dateFontFace,
    dateOffsetDays,
    textBox,
    textColor,
    textAlign,
    textFontSize,
    textBold,
    textDefault,
  } = {},
) => {
  if (!Array.isArray(pairs) || pairs.length === 0) {
    return
  }

  const hasMiddleBox = Boolean(middleBox)
  const hasTextBox = Boolean(textBox)
  const resolveSlideText = (pair) => {
    if (!hasTextBox) {
      return ''
    }
    const rawText = typeof pair?.slideText === 'string' ? pair.slideText.trim() : ''
    if (rawText) {
      return rawText
    }
    return typeof textDefault === 'string' ? textDefault.trim() : ''
  }

  const completePairs = pairs.filter((pair) => {
    const hasBefore = Boolean(pair?.beforeImage)
    const hasAfter = Boolean(pair?.afterImage)
    const hasMiddle = hasMiddleBox ? Boolean(pair?.middleImage) : true
    const hasText = hasTextBox ? Boolean(resolveSlideText(pair)) : true
    return hasBefore && hasAfter && hasMiddle && hasText
  })
  if (completePairs.length === 0) {
    return
  }

  const pptx = new PptxGenJS()
  pptx.layout = 'LAYOUT_WIDE'

  const resolvedMasterTitle = masterTitle || MASTER_TITLE
  const resolvedLeftBox = leftBox || LEFT_BOX
  const resolvedMiddleBox = middleBox || null
  const resolvedRightBox = rightBox || RIGHT_BOX
  defineMaster(pptx, {
    backgroundImage,
    slideTitle: slideTitle || SLIDE_TITLE,
    masterTitle: resolvedMasterTitle,
  })
  const showLabels = !backgroundImage
  const baseDate = getDateWithOffset(dateOffsetDays)
  const slideDate = formatDateForSlide(baseDate)
  const fileDate = formatDateForFile(baseDate)
  const resolvedDateBox = dateBox || DATE_BOX
  const resolvedDateSlide = dateSlide || 'first'
  const resolvedDateColor = dateColor || 'FFFFFF'
  const resolvedDateAlign = dateAlign || 'right'
  const resolvedDateFontSize = dateFontSize || 14
  const resolvedDateBold = Boolean(dateBold)
  const resolvedDateFontFace = dateFontFace || 'Calibri'
  const resolvedTextColor = textColor || '111111'
  const resolvedTextAlign = textAlign || 'center'
  const resolvedTextFontSize = textFontSize || 26
  const resolvedTextBold = Boolean(textBold)

  const addDateToSlide = (slide) => {
    slide.addText(slideDate, {
      ...resolvedDateBox,
      fontFace: resolvedDateFontFace,
      fontSize: resolvedDateFontSize,
      color: resolvedDateColor,
      align: resolvedDateAlign,
      valign: 'middle',
      bold: resolvedDateBold,
    })
  }

  const firstSlide = pptx.addSlide()
  const firstSlideSource = normalizeImage(firstSlideImage || FIRST_SLIDE_URL)
  if (firstSlideSource) {
    firstSlide.addImage({
      ...firstSlideSource,
      x: 0,
      y: 0,
      ...SLIDE_SIZE,
    })
  }

  if (resolvedDateSlide === 'first') {
    addDateToSlide(firstSlide)
  }

  const secondSlideSource = normalizeImage(
    secondSlideImage || SECOND_SLIDE_URL,
  )
  if (secondSlideSource) {
    const secondSlide = pptx.addSlide()
    secondSlide.addImage({
      ...secondSlideSource,
      x: 0,
      y: 0,
      ...SLIDE_SIZE,
    })
    if (resolvedDateSlide === 'second') {
      addDateToSlide(secondSlide)
    }
  } else if (resolvedDateSlide === 'second') {
    addDateToSlide(firstSlide)
  }

  completePairs.forEach((pair) => {
    const slide = pptx.addSlide(resolvedMasterTitle)

    if (showLabels) {
      slide.addText('Before', {
        x: resolvedLeftBox.x,
        y: 1.6,
        w: resolvedLeftBox.w,
        h: 0.4,
        fontFace: 'Calibri',
        fontSize: 18,
        color: '111111',
        bold: true,
        align: 'center',
      })

      if (resolvedMiddleBox) {
        slide.addText('Middle', {
          x: resolvedMiddleBox.x,
          y: 1.6,
          w: resolvedMiddleBox.w,
          h: 0.4,
          fontFace: 'Calibri',
          fontSize: 18,
          color: '111111',
          bold: true,
          align: 'center',
        })
      }

      slide.addText('After', {
        x: resolvedRightBox.x,
        y: 1.6,
        w: resolvedRightBox.w,
        h: 0.4,
        fontFace: 'Calibri',
        fontSize: 18,
        color: '111111',
        bold: true,
        align: 'center',
      })
    }

    const beforeSource = normalizeImage(pair?.beforeImage)
    const middleSource = normalizeImage(pair?.middleImage)
    const afterSource = normalizeImage(pair?.afterImage)

    if (beforeSource) {
      slide.addImage({
        ...beforeSource,
        ...resolvedLeftBox,
        sizing: { type: 'contain' },
      })
    }

    if (middleSource && resolvedMiddleBox) {
      slide.addImage({
        ...middleSource,
        ...resolvedMiddleBox,
        sizing: { type: 'contain' },
      })
    }

    if (afterSource) {
      slide.addImage({
        ...afterSource,
        ...resolvedRightBox,
        sizing: { type: 'contain' },
      })
    }

    if (hasTextBox) {
      const slideText = resolveSlideText(pair)
      if (slideText) {
        slide.addText(slideText, {
          ...textBox,
          fontFace: 'Calibri',
          fontSize: resolvedTextFontSize,
          color: resolvedTextColor,
          align: resolvedTextAlign,
          valign: 'middle',
          bold: resolvedTextBold,
        })
      }
    }
  })

  const lastSlide = pptx.addSlide()
  const lastSlideSource = normalizeImage(lastSlideImage || LAST_SLIDE_URL)
  if (lastSlideSource) {
    lastSlide.addImage({
      ...lastSlideSource,
      x: 0,
      y: 0,
      ...SLIDE_SIZE,
    })
  }

  const safeFileName =
    fileName || `${fileNamePrefix || FILE_NAME_PREFIX}-${fileDate}.pptx`

  await pptx.writeFile({ fileName: safeFileName })
}
