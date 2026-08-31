export const normalizeRoute = (path = '/') => {
  const trimmed = path.replace(/\/+$/, '')
  return trimmed === '' ? '/' : trimmed
}

export const ROUTES = {
  clean: '/',
  compliance: '/compliance',
  desilting: '/desilting',
  dailyPlot: '/daily-plot',
  master: '/master',
  extract: '/extract',
  merge: '/merge',
  pdf: '/pdf',
  mergePdf: '/merge-pdf',
  collage: '/collage',
  gpsPdf: '/gps-pdf',
  pptxToPdf: '/pptx-to-pdf',
  pptxEditor: '/pptx-editor',
}
