#!/usr/bin/env node

/**
 * PPT-Font-Toolkit / Font Metrics
 * 
 * 从字体文件中提取 usWinAscent、usWinDescent、unitsPerEm，
 * 计算 PPT 单倍行距比率 (usWinAscent + usWinDescent) / unitsPerEm
 * 
 * 用法：
 *   1. 扫描系统字体：  node ppt-font-toolkit.mjs metrics --scan
 *   2. 指定字体文件：  node font-metrics.mjs /path/to/font.ttf
 *   3. 指定目录扫描：  node ppt-font-toolkit.mjs metrics --dir /path/to/fonts
 *   4. 输出 JSON：     node font-metrics.mjs --scan --json
 *   5. JSON Map模式：  node font-metrics.mjs --scan --json --map
 *   6. 保存到文件：    node font-metrics.mjs --scan --json --save output.json
 *   7. 搜索字体名称：  node ppt-font-toolkit.mjs metrics --scan --filter 微软
 *   8. 安装后命令：    ppt-font-metrics --scan
 */

import { readFileSync, writeFileSync, readdirSync, statSync, existsSync } from 'fs'
import { join, extname, resolve } from 'path'
import { platform } from 'os'
import { fileURLToPath } from 'url'

// ============================================================
// 核心：从字体文件二进制中提取 metrics（不依赖第三方库）
// ============================================================

/**
 * 读取字体文件并提取 metrics
 * @param {string} fontPath 字体文件路径
 * @returns {{ familyName: string, usWinAscent: number, usWinDescent: number, unitsPerEm: number, lineRatio: number } | null}
 */
export function getFontMetrics(fontPath) {
  try {
    const buffer = readFileSync(fontPath)
    const view = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength)

    // 检查是否是 TTC (TrueType Collection)
    const tag = String.fromCharCode(view.getUint8(0), view.getUint8(1), view.getUint8(2), view.getUint8(3))
    
    let offsets = []
    if (tag === 'ttcf') {
      // TTC 文件，包含多个字体
      const numFonts = view.getUint32(8)
      for (let i = 0; i < numFonts; i++) {
        offsets.push(view.getUint32(12 + i * 4))
      }
    } else {
      offsets = [0]
    }

    const results = []
    for (const offset of offsets) {
      const result = parseSingleFont(view, offset)
      if (result) {
        result.path = fontPath
        results.push(result)
      }
    }
    return results
  } catch (e) {
    return null
  }
}

/**
 * 解析单个字体
 */
function parseSingleFont(view, startOffset) {
  try {
    const numTables = view.getUint16(startOffset + 4)
    
    let headOffset = -1, headLength = 0
    let os2Offset = -1, os2Length = 0
    let nameOffset = -1, nameLength = 0

    // 查找 head、OS/2、name 表
    for (let i = 0; i < numTables; i++) {
      const tableOffset = startOffset + 12 + i * 16
      const tableTag = String.fromCharCode(
        view.getUint8(tableOffset),
        view.getUint8(tableOffset + 1),
        view.getUint8(tableOffset + 2),
        view.getUint8(tableOffset + 3)
      )
      const offset = view.getUint32(tableOffset + 8)
      const length = view.getUint32(tableOffset + 12)

      if (tableTag === 'head') { headOffset = offset; headLength = length }
      if (tableTag === 'OS/2') { os2Offset = offset; os2Length = length }
      if (tableTag === 'name') { nameOffset = offset; nameLength = length }
    }

    if (headOffset < 0 || os2Offset < 0) return null

    // head 表：unitsPerEm 在偏移 18
    const unitsPerEm = view.getUint16(headOffset + 18)

    // OS/2 表：usWinAscent 在偏移 68，usWinDescent 在偏移 70
    // OS/2 版本
    const os2Version = view.getUint16(os2Offset)
    
    // sTypoAscender (偏移 68), sTypoDescender (偏移 70) - 这些是 signed
    // usWinAscent (偏移 74), usWinDescent (偏移 76) - 这些是 unsigned
    // 但实际偏移取决于 OS/2 表版本，标准位置：
    // usWinAscent = offset + 68 (在 version 0+)
    // usWinDescent = offset + 70
    
    let usWinAscent, usWinDescent

    if (os2Length >= 78) {
      // usWinAscent 在 OS/2 偏移 68, usWinDescent 在 70
      // 注意：在 OS/2 表中，字段顺序是固定的
      // Offset 68: sTypoAscender (int16)
      // Offset 70: sTypoDescender (int16)  
      // Offset 72: sTypoLineGap (int16)
      // Offset 74: usWinAscent (uint16)
      // Offset 76: usWinDescent (uint16)
      usWinAscent = view.getUint16(os2Offset + 74)
      usWinDescent = view.getUint16(os2Offset + 76)
    } else {
      return null
    }

    // 提取字体名称
    let names = { familyName: '', familyNameEn: '' }
    if (nameOffset >= 0) {
      names = extractFontName(view, nameOffset, nameLength)
    }

    const lineRatio = (usWinAscent + usWinDescent) / unitsPerEm

    return {
      familyName: names.familyName,
      familyNameEn: names.familyNameEn,
      usWinAscent,
      usWinDescent,
      unitsPerEm,
      lineRatio: +lineRatio.toFixed(6),
    }
  } catch (e) {
    return null
  }
}

/**
 * 解码 Mac 平台 GB2312 (encodingID=25) 编码的字节为中文字符串
 */
function decodeMacChineseGB(view, start, length) {
  // GB2312 双字节编码：高字节 0xA1-0xF7，低字节 0xA1-0xFE
  // 使用 TextDecoder 解码
  const bytes = new Uint8Array(length)
  for (let i = 0; i < length; i++) {
    bytes[i] = view.getUint8(start + i)
  }
  try {
    return new TextDecoder('gb18030').decode(bytes)
  } catch {
    return ''
  }
}

/**
 * 解码 Mac 平台 Big5 (encodingID=2) 编码的字节为中文字符串
 */
function decodeMacChineseBig5(view, start, length) {
  const bytes = new Uint8Array(length)
  for (let i = 0; i < length; i++) {
    bytes[i] = view.getUint8(start + i)
  }
  try {
    return new TextDecoder('big5').decode(bytes)
  } catch {
    return ''
  }
}

/**
 * 从 name 表提取字体族名
 * 返回 { familyName, familyNameEn }
 * familyName: 优先简体中文 > 繁体中文 > 英文
 * familyNameEn: 英文名
 */
function extractFontName(view, nameOffset, nameLength) {
  try {
    const count = view.getUint16(nameOffset + 2)
    const stringOffset = nameOffset + view.getUint16(nameOffset + 4)
    
    // 收集候选名称
    let nameZhCN = ''     // 简体中文 (Windows platform, langID=2052)
    let nameZhTW = ''     // 繁体中文 (Windows platform, langID=1028)
    let nameEn = ''       // 英文 (Windows platform, langID=1033)
    let nameMacEn = ''    // Mac 英文
    let nameMacZhCN = ''  // Mac 简体中文 (encodingID=25)
    let nameMacZhTW = ''  // Mac 繁体中文 (encodingID=2)
    
    for (let i = 0; i < count; i++) {
      const recordOffset = nameOffset + 6 + i * 12
      const platformID = view.getUint16(recordOffset)
      const encodingID = view.getUint16(recordOffset + 2)
      const languageID = view.getUint16(recordOffset + 4)
      const nameID = view.getUint16(recordOffset + 6)
      const length = view.getUint16(recordOffset + 8)
      const offset = view.getUint16(recordOffset + 10)

      // nameID 1 = Font Family name, nameID 16 = Typographic Family name
      if (nameID !== 1 && nameID !== 16) continue

      const strStart = stringOffset + offset
      let name = ''

      if (platformID === 3 || platformID === 0) {
        // Windows/Unicode: UTF-16BE
        for (let j = 0; j < length; j += 2) {
          name += String.fromCharCode(view.getUint16(strStart + j))
        }
        if (!name) continue
        if (nameID === 1 || nameID === 16) {
          if (languageID === 2052 && !nameZhCN) {
            nameZhCN = name
          } else if (languageID === 1028 && !nameZhTW) {
            nameZhTW = name
          } else if (languageID === 1033 && !nameEn) {
            nameEn = name
          }
        }
      } else if (platformID === 1) {
        // Mac platform
        if (encodingID === 25 && nameID === 1) {
          // Mac 简体中文 (GB2312)
          name = decodeMacChineseGB(view, strStart, length)
          if (name && !nameMacZhCN) nameMacZhCN = name
        } else if (encodingID === 2 && nameID === 1) {
          // Mac 繁体中文 (Big5)
          name = decodeMacChineseBig5(view, strStart, length)
          if (name && !nameMacZhTW) nameMacZhTW = name
        } else if (encodingID === 0 && nameID === 1) {
          // Mac Roman (ASCII/Latin)
          for (let j = 0; j < length; j++) {
            name += String.fromCharCode(view.getUint8(strStart + j))
          }
          if (name && !nameMacEn) nameMacEn = name
        }
      }
    }

    // 优先级：简体中文 > 繁体中文 > Mac简中 > Mac繁中 > 英文 > Mac英文
    const familyName = nameZhCN || nameZhTW || nameMacZhCN || nameMacZhTW || nameEn || nameMacEn || ''
    const familyNameEn = nameEn || nameMacEn || ''
    
    return { familyName, familyNameEn }
  } catch (e) {
    return { familyName: '', familyNameEn: '' }
  }
}

// ============================================================
// 系统字体目录扫描
// ============================================================

export function getSystemFontDirs() {
  const os = platform()
  
  if (os === 'darwin') {
    return [
      '/System/Library/Fonts',
      '/Library/Fonts',
      join(process.env.HOME || '', 'Library/Fonts'),
    ]
  }
  
  if (os === 'win32') {
    return [
      join(process.env.WINDIR || 'C:\\Windows', 'Fonts'),
      join(process.env.LOCALAPPDATA || '', 'Microsoft\\Windows\\Fonts'),
    ]
  }
  
  // Linux
  return [
    '/usr/share/fonts',
    '/usr/local/share/fonts',
    join(process.env.HOME || '', '.fonts'),
    join(process.env.HOME || '', '.local/share/fonts'),
  ]
}

const FONT_EXTENSIONS = new Set(['.ttf', '.otf', '.ttc', '.woff'])

export function scanFontFiles(dir) {
  const fonts = []
  
  if (!existsSync(dir)) return fonts

  try {
    const entries = readdirSync(dir)
    for (const entry of entries) {
      const fullPath = join(dir, entry)
      try {
        const stat = statSync(fullPath)
        if (stat.isDirectory()) {
          fonts.push(...scanFontFiles(fullPath))
        } else if (FONT_EXTENSIONS.has(extname(entry).toLowerCase())) {
          fonts.push(fullPath)
        }
      } catch {
        // 跳过权限不足的文件
      }
    }
  } catch {
    // 跳过权限不足的目录
  }

  return fonts
}

// ============================================================
// 主逻辑
// ============================================================

function printUsage() {
  console.log(`
PPT-Font-Toolkit — Font Metrics

计算 PPT 单倍行距比率

用法:
  node ppt-font-toolkit.mjs metrics <字体文件路径>
  node ppt-font-toolkit.mjs metrics --dir <目录路径>
  node ppt-font-toolkit.mjs metrics --scan
  node font-metrics.mjs <字体文件路径>
  node font-metrics.mjs --dir <目录路径>
  node font-metrics.mjs --scan           (兼容旧脚本名)
  ppt-font-metrics <字体文件路径>        (安装后可用)

参数示例:
  --scan                  扫描系统字体
  --filter <关键字>       按名称过滤
  --json                  以 JSON 格式输出（数组）
  --json --map            以 JSON 格式输出（familyName 为 key 的 Map）
  --json --save <文件路径> 输出 JSON 并保存到文件
  --code                  输出可直接使用的代码

示例:
  node ppt-font-toolkit.mjs metrics /Library/Fonts/msyh.ttf
  node font-metrics.mjs --scan --filter "微软|Arial|Times|宋体|Courier"
  node font-metrics.mjs --scan --code
  ppt-font-metrics --scan --json        (安装后可用)
`)
}

function readNextArg(argv, index, flagName) {
  const value = argv[index]
  if (!value) {
    throw new Error(`${flagName} 缺少参数`)
  }
  return value
}

function emitMetricsEvent(callbacks, type, payload = {}) {
  const event = { type, ...payload }

  if (typeof callbacks.onEvent === 'function') {
    callbacks.onEvent(event)
  }

  if (type === 'scanStart' && typeof callbacks.onScanStart === 'function') {
    callbacks.onScanStart(event)
  } else if (type === 'scanComplete' && typeof callbacks.onScanComplete === 'function') {
    callbacks.onScanComplete(event)
  } else if (type === 'save' && typeof callbacks.onSave === 'function') {
    callbacks.onSave(event)
  } else if (type === 'complete' && typeof callbacks.onComplete === 'function') {
    callbacks.onComplete(event)
  }
}

export function parseMetricsArgs(argv = process.argv.slice(2)) {
  const options = {
    help: false,
    json: false,
    code: false,
    map: false,
    scan: false,
    savePath: null,
    filter: null,
    dir: null,
    inputs: [],
  }

  for (let index = 0; index < argv.length; index++) {
    const arg = argv[index]

    if (arg === '--help' || arg === '-h') {
      options.help = true
    } else if (arg === '--json') {
      options.json = true
    } else if (arg === '--code') {
      options.code = true
    } else if (arg === '--map') {
      options.map = true
    } else if (arg === '--scan') {
      options.scan = true
    } else if (arg === '--save') {
      options.savePath = readNextArg(argv, ++index, '--save')
    } else if (arg === '--filter') {
      options.filter = readNextArg(argv, ++index, '--filter')
    } else if (arg === '--dir') {
      options.dir = readNextArg(argv, ++index, '--dir')
    } else if (!arg.startsWith('-')) {
      options.inputs.push(arg)
    } else {
      throw new Error(`未知参数: ${arg}`)
    }
  }

  return options
}

function normalizeMetricsOptions(options = {}) {
  const rawInputs = []

  const appendValues = (value) => {
    if (Array.isArray(value)) {
      for (const item of value) appendValues(item)
      return
    }
    if (value !== undefined && value !== null && value !== '') {
      rawInputs.push(String(value))
    }
  }

  appendValues(options.inputs)
  appendValues(options.files)
  appendValues(options.fontFiles)
  appendValues(options.paths)
  appendValues(options.input)

  return {
    help: Boolean(options.help),
    json: Boolean(options.json),
    code: Boolean(options.code),
    map: Boolean(options.map),
    scan: Boolean(options.scan),
    savePath: options.savePath || options.save || null,
    filter: options.filter || null,
    dir: options.dir || options.directory || null,
    inputs: rawInputs,
  }
}

function dedupeMetricsResults(results) {
  const seen = new Map()
  for (const result of results) {
    const key = result.familyName || result.path
    if (!seen.has(key)) {
      seen.set(key, result)
    }
  }
  return Array.from(seen.values())
}

export function buildMetricsJsonOutput(results) {
  return results.map((result) => {
    const entry = {
      familyName: result.familyName,
      usWinAscent: result.usWinAscent,
      usWinDescent: result.usWinDescent,
      unitsPerEm: result.unitsPerEm,
      lineRatio: result.lineRatio,
    }
    if (result.familyNameEn && result.familyNameEn !== result.familyName) {
      entry.familyNameEn = result.familyNameEn
    }
    return entry
  })
}

export function buildMetricsMap(results) {
  const output = {}
  for (const result of results) {
    if (!result.familyName) continue
    const entry = {
      usWinAscent: result.usWinAscent,
      usWinDescent: result.usWinDescent,
      unitsPerEm: result.unitsPerEm,
      lineRatio: result.lineRatio,
    }
    if (result.familyNameEn && result.familyNameEn !== result.familyName) {
      entry.familyNameEn = result.familyNameEn
    }
    output[result.familyName] = entry
  }
  return output
}

export function formatMetricsCode(results) {
  const lines = [
    '// PPT-Font-Toolkit 字体单倍行距比率表',
    '// lineRatio = (usWinAscent + usWinDescent) / unitsPerEm',
    '// PPT行高(px) = fontSize(pt) × lineRatio × spcPct × 96/72',
    'const FONT_LINE_RATIO: Record<string, number> = {',
  ]

  for (const result of results) {
    if (!result.familyName) continue
    const enInfo = result.familyNameEn && result.familyNameEn !== result.familyName ? ` [${result.familyNameEn}]` : ''
    lines.push(`  '${result.familyName}': ${result.lineRatio},  // ascent=${result.usWinAscent} descent=${result.usWinDescent} em=${result.unitsPerEm}${enInfo}`)
  }

  lines.push('}')
  lines.push('')
  lines.push('const DEFAULT_LINE_RATIO = 1.2')
  return lines.join('\n')
}

export function formatMetricsTable(results) {
  const displayNames = results.map((result) => {
    if (result.familyNameEn && result.familyNameEn !== result.familyName) {
      return `${result.familyName} (${result.familyNameEn})`
    }
    return result.familyName
  })

  const maxName = Math.max(12, ...displayNames.map((name) => name.length))
  const header = [
    '字体名称'.padEnd(maxName),
    'Ascent'.padStart(8),
    'Descent'.padStart(8),
    'EmSize'.padStart(8),
    'LineRatio'.padStart(10),
  ].join('  ')

  const lines = [header, '-'.repeat(header.length)]
  for (let index = 0; index < results.length; index++) {
    const result = results[index]
    lines.push([
      displayNames[index].padEnd(maxName),
      String(result.usWinAscent).padStart(8),
      String(result.usWinDescent).padStart(8),
      String(result.unitsPerEm).padStart(8),
      String(result.lineRatio).padStart(10),
    ].join('  '))
  }

  lines.push(`\n共 ${results.length} 个字体`)
  return lines.join('\n')
}

export function collectMetrics(options = {}, callbacks = {}) {
  const normalized = normalizeMetricsOptions(options)
  let fontFiles = []

  if (normalized.scan) {
    const dirs = getSystemFontDirs()
    emitMetricsEvent(callbacks, 'scanStart', { mode: 'system', directories: dirs })
    for (const dir of dirs) {
      fontFiles.push(...scanFontFiles(dir))
    }
    emitMetricsEvent(callbacks, 'scanComplete', { mode: 'system', directories: dirs, count: fontFiles.length })
  } else if (normalized.dir) {
    const absDir = resolve(normalized.dir)
    emitMetricsEvent(callbacks, 'scanStart', { mode: 'directory', directory: absDir })
    fontFiles = scanFontFiles(absDir)
    emitMetricsEvent(callbacks, 'scanComplete', { mode: 'directory', directory: absDir, count: fontFiles.length })
  } else {
    fontFiles = normalized.inputs.map((input) => resolve(input))
  }

  const allResults = []
  for (const file of fontFiles) {
    const results = getFontMetrics(file)
    if (results) {
      allResults.push(...results)
    }
  }

  let filtered = allResults
  if (normalized.filter) {
    const regex = new RegExp(normalized.filter, 'i')
    filtered = allResults.filter((result) =>
      regex.test(result.familyName) || regex.test(result.familyNameEn) || regex.test(result.path)
    )
  }

  filtered = dedupeMetricsResults(filtered)
  filtered.sort((left, right) => left.familyName.localeCompare(right.familyName))

  if (filtered.length === 0) {
    throw new Error('未找到匹配的字体')
  }

  return filtered
}

export function runMetrics(options = {}, callbacks = {}) {
  const normalized = normalizeMetricsOptions(options)
  const results = collectMetrics(normalized, callbacks)

  let format = 'array'
  let output = buildMetricsJsonOutput(results)
  let savedTo = null

  if (normalized.code) {
    format = 'code'
    output = formatMetricsCode(results)
  } else if (normalized.map) {
    format = 'map'
    output = buildMetricsMap(results)
  }

  if (!normalized.code && normalized.savePath) {
    savedTo = resolve(normalized.savePath)
    writeFileSync(savedTo, JSON.stringify(output, null, 2) + '\n', 'utf-8')
    emitMetricsEvent(callbacks, 'save', { path: savedTo, format })
  }

  const summary = { format, results, output, savedTo }
  emitMetricsEvent(callbacks, 'complete', summary)
  return summary
}

export function main(argv = process.argv.slice(2)) {
  try {
    const options = parseMetricsArgs(argv)

    if (options.help || argv.length === 0) {
      printUsage()
      return 0
    }

    const summary = runMetrics(options, {
      onScanStart(event) {
        if (event.mode === 'system') {
          console.error(`扫描系统字体目录: ${event.directories.join(', ')}`)
        } else if (event.directory) {
          console.error(`扫描目录: ${event.directory}`)
        }
      },
      onScanComplete(event) {
        console.error(`找到 ${event.count} 个字体文件\n`)
      },
      onSave(event) {
        console.error(`已保存到: ${event.path}`)
      },
    })

    if (summary.format === 'code') {
      console.log(summary.output)
    } else if (options.json) {
      if (!summary.savedTo) {
        console.log(JSON.stringify(summary.output, null, 2))
      }
    } else {
      console.log(formatMetricsTable(summary.results))
    }

    return 0
  } catch (error) {
    console.error(error.message)
    return 1
  }
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  process.exit(main())
}
