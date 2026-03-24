export {
  inspectEmbeddedFontFile,
  listEmbeddedFonts,
  main as runRecoverCli,
  parseRecoverArgs,
  recoverEmbeddedFonts,
  recoverEmbeddedFonts as recoverFonts,
  runRecover,
} from './ppt-font-recover.mjs'

export {
  buildMetricsJsonOutput,
  buildMetricsMap,
  collectMetrics,
  collectMetrics as extractFontMetrics,
  formatMetricsCode,
  formatMetricsTable,
  getFontMetrics,
  getSystemFontDirs,
  main as runMetricsCli,
  parseMetricsArgs,
  runMetrics,
  scanFontFiles,
} from './font-metrics.mjs'
