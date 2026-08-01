import PptxGenJS from 'pptxgenjs'

const MASTER_TITLE = 'CLEAN_PUNJAB_MASTER'
const SLIDE_TITLE = 'Plots Cleaning-Activity'
const FIRST_SLIDE_URL = '/first-slide.jpeg'
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

const roundImageCorners = (imageSource, borderRadius, slotW, slotH) => {
  return new Promise((resolve) => {
    if (!borderRadius || borderRadius <= 0) {
      resolve(imageSource)
      return
    }

    let src = ''
    if (imageSource && typeof imageSource === 'object') {
      src = imageSource.data || imageSource.path || ''
    }

    if (!src) {
      resolve(imageSource)
      return
    }

    const img = new Image()
    img.crossOrigin = 'anonymous'
    img.onload = () => {
      try {
        const canvas = document.createElement('canvas')
        canvas.width = img.width
        canvas.height = img.height
        const ctx = canvas.getContext('2d')
        if (!ctx) {
          resolve(imageSource)
          return
        }

        const w = slotW || 4.0
        const h = slotH || 3.0
        const renderedW = Math.min(w, h * (img.width / img.height))
        const scale = img.width / (renderedW * 100)
        const radius = Math.min(img.width / 2, img.height / 2, borderRadius * scale)

        ctx.beginPath()
        ctx.moveTo(radius, 0)
        ctx.lineTo(img.width - radius, 0)
        ctx.quadraticCurveTo(img.width, 0, img.width, radius)
        ctx.lineTo(img.width, img.height - radius)
        ctx.quadraticCurveTo(img.width, img.height, img.width - radius, img.height)
        ctx.lineTo(radius, img.height)
        ctx.quadraticCurveTo(0, img.height, 0, img.height - radius)
        ctx.lineTo(0, radius)
        ctx.quadraticCurveTo(0, 0, radius, 0)
        ctx.closePath()
        ctx.clip()

        ctx.drawImage(img, 0, 0)

        const dataUrl = canvas.toDataURL('image/png')
        resolve({ data: dataUrl })
      } catch (err) {
        console.error('Error rounding corners:', err)
        resolve(imageSource)
      }
    }
    img.onerror = () => {
      resolve(imageSource)
    }
    img.src = src
  })
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

/**
 * generateCompletedSlidesReport – identical to generateReport but only
 * renders slides whose every required image slot is filled (complete pairs).
 * Incomplete pairs are excluded entirely.
 */
export const generateCompletedSlidesReport = async (
  pairs,
  options = {},
) => {
  const {
    placeholders,
    textboxes,
  } = options

  const hasMiddleBox = Boolean(options.middleBox)
  const hasTextBox = Boolean(options.textBox)

  const resolveSlideText = (pair) => {
    if (!hasTextBox) return ''
    const rawText = typeof pair?.slideText === 'string' ? pair.slideText.trim() : ''
    if (rawText) return rawText
    return typeof options.textDefault === 'string' ? options.textDefault.trim() : ''
  }

  const completedOnly = pairs.filter((pair) => {
    if (placeholders && placeholders.length > 0) {
      return placeholders.every((slot) => Boolean(pair?.[slot.key]))
    }
    const hasBefore = Boolean(pair?.beforeImage)
    const hasAfter = Boolean(pair?.afterImage)
    const hasMiddle = hasMiddleBox ? Boolean(pair?.middleImage) : true
    const hasText = hasTextBox ? Boolean(resolveSlideText(pair)) : true
    return hasBefore && hasAfter && hasMiddle && hasText
  })

  // Re-use the main function but pass only the complete pairs.
  // We still need incomplete pairs to be empty so generateReport's own filter
  // sees no incomplete pairs to render.
  await generateReport(completedOnly, options)
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
    placeholders,
    textboxes,
    firstSlidePlaceholders,
    firstSlideTextboxes,
    firstSlideData,
    lastSlidePlaceholders,
    lastSlideTextboxes,
    lastSlideData,
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
    if (placeholders && placeholders.length > 0) {
      return placeholders.every((slot) => Boolean(pair?.[slot.key]))
    }
    const hasBefore = Boolean(pair?.beforeImage)
    const hasAfter = Boolean(pair?.afterImage)
    const hasMiddle = hasMiddleBox ? Boolean(pair?.middleImage) : true
    const hasText = hasTextBox ? Boolean(resolveSlideText(pair)) : true
    return hasBefore && hasAfter && hasMiddle && hasText
  })

  // For custom master layout: ALL non-complete pairs that have ANY content (image or text) get a slide.
  // For standard layout: collect pairs that have at least one image but are not complete.
  const incompletePairs = pairs.filter((pair) => {
    if (placeholders && placeholders.length > 0) {
      const isComplete = placeholders.every((slot) => Boolean(pair?.[slot.key]))
      if (isComplete) return false
      // Include pairs with at least one image slot filled
      const hasAnyImage = placeholders.some((slot) => Boolean(pair?.[slot.key]))
      // OR at least one textbox value filled (for textbox-only slides)
      const hasAnyText = textboxes && textboxes.length > 0
        ? textboxes.some((box) => typeof pair?.[box.key] === 'string' && pair[box.key].trim())
        : false
      return hasAnyImage || hasAnyText
    }
    const hasBefore = Boolean(pair?.beforeImage)
    const hasAfter = Boolean(pair?.afterImage)
    const hasMiddle = hasMiddleBox ? Boolean(pair?.middleImage) : true
    const hasText = hasTextBox ? Boolean(resolveSlideText(pair)) : true
    const isComplete = hasBefore && hasAfter && hasMiddle && hasText
    const hasAny = hasBefore || hasAfter || (hasMiddleBox && Boolean(pair?.middleImage))
    return hasAny && !isComplete
  })


  if (completePairs.length === 0 && incompletePairs.length === 0) {
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

  if (firstSlidePlaceholders && firstSlidePlaceholders.length > 0) {
    for (const slot of firstSlidePlaceholders) {
      let imageSource = normalizeImage(firstSlideData?.[slot.key])
      if (imageSource) {
        if (slot.borderRadius && slot.borderRadius > 0) {
          imageSource = await roundImageCorners(imageSource, slot.borderRadius, slot.w, slot.h)
        }
        firstSlide.addImage({
          ...imageSource,
          x: slot.x,
          y: slot.y,
          w: slot.w,
          h: slot.h,
          sizing: { type: 'contain' },
        })
      }
    }
  }
  if (firstSlideTextboxes && firstSlideTextboxes.length > 0) {
    firstSlideTextboxes.forEach((box) => {
      const textVal = typeof firstSlideData?.[box.key] === 'string' ? firstSlideData[box.key].trim() : ''
      const textToUse = textVal || box.textDefault || ''
      if (textToUse) {
        firstSlide.addText(textToUse, {
          x: box.x,
          y: box.y,
          w: box.w,
          h: box.h,
          fontFace: box.fontFace || 'Calibri',
          fontSize: box.fontSize || 20,
          color: box.fontColor || '000000',
          align: box.align || 'left',
          valign: 'middle',
          bold: Boolean(box.bold),
        })
      }
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

  // ── Render all eligible pairs in original order ──────────────────────────
  // Build a Set of all pairs that should generate a slide (complete + incomplete).
  const eligiblePairSet = new Set([...completePairs, ...incompletePairs])

  for (const pair of pairs) {
    if (!eligiblePairSet.has(pair)) continue

    const slide = pptx.addSlide(resolvedMasterTitle)

    if (placeholders && placeholders.length > 0) {
      // Custom layout: render any image slots that are filled
      for (const slot of placeholders) {
        let imageSource = normalizeImage(pair?.[slot.key])
        if (imageSource) {
          if (slot.borderRadius && slot.borderRadius > 0) {
            imageSource = await roundImageCorners(imageSource, slot.borderRadius, slot.w, slot.h)
          }
          slide.addImage({
            ...imageSource,
            x: slot.x,
            y: slot.y,
            w: slot.w,
            h: slot.h,
            sizing: { type: 'contain' },
          })
        }
      }

      if (textboxes && textboxes.length > 0) {
        textboxes.forEach((box) => {
          const textVal = typeof pair?.[box.key] === 'string' ? pair[box.key].trim() : ''
          const textToUse = textVal || box.textDefault || ''
          if (textToUse) {
            slide.addText(textToUse, {
              x: box.x,
              y: box.y,
              w: box.w,
              h: box.h,
              fontFace: box.fontFace || 'Calibri',
              fontSize: box.fontSize || 20,
              color: box.fontColor || '000000',
              align: box.align || 'left',
              valign: 'middle',
              bold: Boolean(box.bold),
            })
          }
        })
      }
    } else {
      // Standard layout
      const isComplete = completePairs.includes(pair)

      if (isComplete) {
        // All images present — show labels for all slots
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
      } else {
        // Partial — render only the filled slots
        const slotMap = [
          { key: 'beforeImage', label: 'Before', box: resolvedLeftBox },
          { key: 'middleImage', label: 'During', box: resolvedMiddleBox },
          { key: 'afterImage', label: 'After', box: resolvedRightBox },
        ]
        for (const slot of slotMap) {
          if (!slot.box) continue
          const imageSource = normalizeImage(pair?.[slot.key])
          if (imageSource) {
            if (showLabels) {
              slide.addText(slot.label, {
                x: slot.box.x,
                y: 1.6,
                w: slot.box.w,
                h: 0.4,
                fontFace: 'Calibri',
                fontSize: 18,
                color: '111111',
                bold: true,
                align: 'center',
              })
            }
            slide.addImage({
              ...imageSource,
              ...slot.box,
              sizing: { type: 'contain' },
            })
          }
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
      }
    }
  }

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

  if (lastSlidePlaceholders && lastSlidePlaceholders.length > 0) {
    for (const slot of lastSlidePlaceholders) {
      let imageSource = normalizeImage(lastSlideData?.[slot.key])
      if (imageSource) {
        if (slot.borderRadius && slot.borderRadius > 0) {
          imageSource = await roundImageCorners(imageSource, slot.borderRadius, slot.w, slot.h)
        }
        lastSlide.addImage({
          ...imageSource,
          x: slot.x,
          y: slot.y,
          w: slot.w,
          h: slot.h,
          sizing: { type: 'contain' },
        })
      }
    }
  }
  if (lastSlideTextboxes && lastSlideTextboxes.length > 0) {
    lastSlideTextboxes.forEach((box) => {
      const textVal = typeof lastSlideData?.[box.key] === 'string' ? lastSlideData[box.key].trim() : ''
      const textToUse = textVal || box.textDefault || ''
      if (textToUse) {
        lastSlide.addText(textToUse, {
          x: box.x,
          y: box.y,
          w: box.w,
          h: box.h,
          fontFace: box.fontFace || 'Calibri',
          fontSize: box.fontSize || 20,
          color: box.fontColor || '000000',
          align: box.align || 'left',
          valign: 'middle',
          bold: Boolean(box.bold),
        })
      }
    })
  }

  const safeFileName =
    fileName || `${fileNamePrefix || FILE_NAME_PREFIX}-${fileDate}.pptx`

  await pptx.writeFile({ fileName: safeFileName })
}
