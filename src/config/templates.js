import { ROUTES, normalizeRoute } from './routes'

export const DEFAULT_SLOTS = [
  {
    key: 'beforeImage',
    label: 'Before',
    className: 'slide-slot slide-slot--before',
    urlMode: 'prompt',
  },
  {
    key: 'afterImage',
    label: 'After',
    className: 'slide-slot slide-slot--after',
    urlMode: 'prompt',
  },
]

export const DESILTING_PRESET_TEXT = [
  'UC-55',
  'UC-56',
  'UC-57',
  'UC-58',
  'UC-59',
  'UC-60',
  'UC-61',
  'UC-62',
  'UC-63',
  'UC-64',
  'UC-65',
  'UC-66',
  'UC-67',
  'UC-68',
  'UC-69',
  'Sector-1',
  'Sector-2',
  'Sector-3',
  'Sector-4',
  'Sector-5',
  'Sector-6',
  'Sector-7',
  'Sector-8',
  'Sector-9',
  'Sector-10',
]

export const getDateOffset = (offsetDays = 0) => {
  const date = new Date()
  date.setDate(date.getDate() + offsetDays)
  return date
}

export const formatDailyPlotDate = (date) => {
  const day = String(date.getDate()).padStart(2, '0')
  const monthNames = [
    'Jan',
    'Feb',
    'Mar',
    'Apr',
    'May',
    'Jun',
    'Jul',
    'Aug',
    'Sep',
    'Oct',
    'Nov',
    'Dec',
  ]
  const month = monthNames[date.getMonth()]
  const year = date.getFullYear()
  return `${day}_${month}_${year}`
}

export const buildDailyPlotFileName = (variant = 'urban') => {
  const dateLabel = formatDailyPlotDate(getDateOffset(-1))
  const suffix = variant === 'rural' ? 'Rural' : 'Urban'
  return `Daily_Plot's_Clearance_Report_${suffix} ${dateLabel}.pptx`
}

export const TEMPLATES = {
  [ROUTES.clean]: {
    eyebrow: 'Clean Punjab',
    title: 'Plots Cleaning-Activity',
    subtext:
      'Drop a Before and After image. A new row appears automatically and the page scrolls to it.',
    masterBgUrl: '/master-slide.png',
    firstSlideUrl: '/first-slide.jpeg',
    secondSlideUrl: '',
    lastSlideUrl: '/last-slide.jpeg',
    fileNamePrefix: 'Plots  Cleaning-Activity_Tehsil Kamoke',
    masterTitle: 'CLEAN_PUNJAB_MASTER',
    slideTitle: 'Plots Cleaning-Activity',
    themeLabel: 'Master slide theme',
    importSkipFirst: 1,
    importSkipLast: 1,
    imageCount: 2,
    slots: DEFAULT_SLOTS,
  },
  [ROUTES.compliance]: {
    eyebrow: 'Suthra Punjab',
    title: 'Compliance Report',
    subtext:
      'Drop Before and After images. A new row appears automatically and the page scrolls to it.',
    masterBgUrl: '/master-slide-2.png',
    firstSlideUrl: '/slide-first-2.png',
    secondSlideUrl: '/slide-second-2.png',
    lastSlideUrl: '/last-slide-2.png',
    fileNamePrefix: 'Compliance Report_Tehsil Kamoke',
    masterTitle: 'COMPLIANCE_MASTER',
    slideTitle: 'Compliance Report',
    themeLabel: 'Compliance master slide theme',
    leftBox: { x: 0.4, y: 1.5, w: 6.0, h: 4.65 },
    rightBox: { x: 6.93, y: 1.5, w: 6.0, h: 4.65 },
    dateSlide: 'second',
    dateBox: { x: 5.07, y: 4.91, w: 3.2, h: 0.4 },
    dateColor: 'FFD54D',
    dateAlign: 'center',
    dateFontSize: 22,
    dateBold: true,
    importSkipFirst: 2,
    importSkipLast: 1,
    imageCount: 2,
    slots: DEFAULT_SLOTS,
  },
  [ROUTES.desilting]: {
    eyebrow: 'Suthra Punjab',
    title: 'Desilting Report',
    subtext:
      'Drop three images for each slide and add the sector text. A new row appears automatically and the page scrolls to it.',
    masterBgUrl: '/master-slide-3.png',
    firstSlideUrl: '/first-slide-3.png',
    secondSlideUrl: '',
    lastSlideUrl: '/last-slide-3.png',
    fileNamePrefix: 'Desilting Report_Tehsil Kamoke',
    masterTitle: 'DESILTING_MASTER',
    slideTitle: 'Desilting Report',
    themeLabel: 'Desilting master slide theme',
    leftBox: { x: 0.16, y: 1.9, w: 4.27, h: 4.975 },
    middleBox: { x: 4.533, y: 1.9, w: 4.27, h: 4.975 },
    rightBox: { x: 8.88, y: 1.9, w: 4.27, h: 4.975 },
    textBox: { x: 4.0, y: 0.88, w: 5.33, h: 0.5 },
    textColor: '111111',
    textAlign: 'center',
    textFontSize: 22,
    textBold: true,
    textDefault: '',
    textLabel: 'Sector text',
    importSkipFirst: 1,
    importSkipLast: 1,
    imageCount: 3,
    slots: [
      {
        key: 'beforeImage',
        label: 'Before',
        className: 'slide-slot slide-slot--before',
        urlMode: 'prompt',
      },
      {
        key: 'middleImage',
        label: 'During',
        className: 'slide-slot slide-slot--middle',
        urlMode: 'prompt',
      },
      {
        key: 'afterImage',
        label: 'After',
        className: 'slide-slot slide-slot--after',
        urlMode: 'prompt',
      },
    ],
  },
  [ROUTES.dailyPlot]: {
    eyebrow: 'Daily Plot',
    title: 'OTC Plot Report',
    subtext:
      'Drop a Before and After image. A new row appears automatically and the page scrolls to it.',
    masterBgUrl: '/master-slide-4.png',
    firstSlideUrl: '/first-slide-4-urban.png',
    firstSlideUrlRural: '/first-slide-4-rural.png',
    secondSlideUrl: '',
    lastSlideUrl: '/last-slide-4.png',
    fileName: buildDailyPlotFileName,
    masterTitle: 'DAILY_PLOT_MASTER',
    slideTitle: 'OTC Plot Report',
    themeLabel: 'Daily plot master slide theme',
    leftBox: { x: 1.0, y: 2.16, w: 5.47, h: 4.05 },
    rightBox: { x: 6.8, y: 2.16, w: 5.47, h: 4.05 },
    dateBox: { x: 5.80, y: 2.25, w: 2.4, h: 0.4 },
    dateAlign: 'center',
    dateColor: '000000',
    dateFontSize: 20,
    dateFontFace: 'Times New Roman',
    dateSlide: 'first',
    dateOffsetDays: -1,
    importSkipFirst: 1,
    importSkipLast: 1,
    imageCount: 2,
    slots: DEFAULT_SLOTS,
  },
}

export const getTemplateForPath = (path) => {
  const normalized = normalizeRoute(path)
  return TEMPLATES[normalized] || TEMPLATES[ROUTES.clean]
}
