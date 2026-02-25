#!/usr/bin/env node

/**
 * 字体 Metrics 提取工具
 * 
 * 从字体文件中提取 usWinAscent、usWinDescent、unitsPerEm，
 * 计算 PPT 单倍行距比率 (usWinAscent + usWinDescent) / unitsPerEm
 * 
 * 用法：
 *   1. 扫描系统字体：  node font-metrics.mjs --scan
 *   2. 指定字体文件：  node font-metrics.mjs /path/to/font.ttf
 *   3. 指定目录扫描：  node font-metrics.mjs --dir /path/to/fonts
 *   4. 输出 JSON：     node font-metrics.mjs --scan --json
 *   5. 搜索字体名称：  node font-metrics.mjs --scan --filter 微软
 */

import { readFileSync, readdirSync, statSync, existsSync } from 'fs'
import { join, extname, resolve } from 'path'
import { platform } from 'os'

// ============================================================
// 核心：从字体文件二进制中提取 metrics（不依赖第三方库）
// ============================================================

/**
 * 读取字体文件并提取 metrics
 * @param {string} fontPath 字体文件路径
 * @returns {{ familyName: string, usWinAscent: number, usWinDescent: number, unitsPerEm: number, lineRatio: number } | null}
 */
function getFontMetrics(fontPath) {
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
    let familyName = ''
    if (nameOffset >= 0) {
      familyName = extractFontName(view, nameOffset, nameLength)
    }

    const lineRatio = (usWinAscent + usWinDescent) / unitsPerEm

    return {
      familyName,
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
 * 从 name 表提取字体族名
 */
function extractFontName(view, nameOffset, nameLength) {
  try {
    const count = view.getUint16(nameOffset + 2)
    const stringOffset = nameOffset + view.getUint16(nameOffset + 4)
    
    let familyName = ''
    
    for (let i = 0; i < count; i++) {
      const recordOffset = nameOffset + 6 + i * 12
      const platformID = view.getUint16(recordOffset)
      const encodingID = view.getUint16(recordOffset + 2)
      const languageID = view.getUint16(recordOffset + 4)
      const nameID = view.getUint16(recordOffset + 6)
      const length = view.getUint16(recordOffset + 8)
      const offset = view.getUint16(recordOffset + 10)

      // nameID 1 = Font Family name, nameID 4 = Full font name
      if (nameID === 1 || nameID === 4) {
        const strStart = stringOffset + offset
        let name = ''

        if (platformID === 3 || platformID === 0) {
          // Windows/Unicode: UTF-16BE
          for (let j = 0; j < length; j += 2) {
            name += String.fromCharCode(view.getUint16(strStart + j))
          }
        } else if (platformID === 1) {
          // Mac: ASCII/Latin
          for (let j = 0; j < length; j++) {
            name += String.fromCharCode(view.getUint8(strStart + j))
          }
        }

        if (name && nameID === 1) {
          // 优先用中文名（platformID=3, languageID=2052 是简体中文）
          if (languageID === 2052 || languageID === 1028) {
            return name
          }
          if (!familyName) {
            familyName = name
          }
        }
      }
    }

    return familyName
  } catch (e) {
    return ''
  }
}

// ============================================================
// 系统字体目录扫描
// ============================================================

function getSystemFontDirs() {
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

function scanFontFiles(dir) {
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
字体 Metrics 提取工具 — 计算 PPT 单倍行距比率

用法:
  node font-metrics.mjs <字体文件路径>          提取单个字体文件的 metrics
  node font-metrics.mjs --dir <目录路径>        扫描指定目录
  node font-metrics.mjs --scan                  扫描系统字体
  node font-metrics.mjs --scan --filter <关键字> 按名称过滤
  node font-metrics.mjs --scan --json           以 JSON 格式输出
  node font-metrics.mjs --scan --code           输出可直接使用的代码

示例:
  node font-metrics.mjs /Library/Fonts/msyh.ttf
  node font-metrics.mjs --scan --filter "微软|Arial|Times|宋体|Courier"
  node font-metrics.mjs --scan --code
`)
}

function main() {
  const args = process.argv.slice(2)

  if (args.length === 0 || args.includes('--help') || args.includes('-h')) {
    printUsage()
    process.exit(0)
  }

  const isJson = args.includes('--json')
  const isCode = args.includes('--code')
  const isScan = args.includes('--scan')
  const filterIdx = args.indexOf('--filter')
  const filterPattern = filterIdx >= 0 ? args[filterIdx + 1] : null
  const dirIdx = args.indexOf('--dir')
  const dirPath = dirIdx >= 0 ? args[dirIdx + 1] : null

  let fontFiles = []

  if (isScan) {
    // 扫描系统字体
    const dirs = getSystemFontDirs()
    console.error(`扫描系统字体目录: ${dirs.join(', ')}`)
    for (const dir of dirs) {
      fontFiles.push(...scanFontFiles(dir))
    }
    console.error(`找到 ${fontFiles.length} 个字体文件\n`)
  } else if (dirPath) {
    // 扫描指定目录
    const absDir = resolve(dirPath)
    console.error(`扫描目录: ${absDir}`)
    fontFiles = scanFontFiles(absDir)
    console.error(`找到 ${fontFiles.length} 个字体文件\n`)
  } else {
    // 处理指定的字体文件
    for (const arg of args) {
      if (!arg.startsWith('-')) {
        fontFiles.push(resolve(arg))
      }
    }
  }

  // 提取 metrics
  const allResults = []
  for (const file of fontFiles) {
    const results = getFontMetrics(file)
    if (results) {
      allResults.push(...results)
    }
  }

  // 按名称过滤
  let filtered = allResults
  if (filterPattern) {
    const regex = new RegExp(filterPattern, 'i')
    filtered = allResults.filter(r => regex.test(r.familyName) || regex.test(r.path))
  }

  // 去重（同名字体保留第一个）
  const seen = new Map()
  for (const r of filtered) {
    const key = r.familyName || r.path
    if (!seen.has(key)) {
      seen.set(key, r)
    }
  }
  filtered = Array.from(seen.values())

  // 按名称排序
  filtered.sort((a, b) => a.familyName.localeCompare(b.familyName))

  if (filtered.length === 0) {
    console.error('未找到匹配的字体')
    process.exit(1)
  }

  // 输出
  if (isCode) {
    console.log('// PPT 字体单倍行距比率表')
    console.log('// lineRatio = (usWinAscent + usWinDescent) / unitsPerEm')
    console.log('// PPT行高(px) = fontSize(pt) × lineRatio × spcPct × 96/72')
    console.log('const FONT_LINE_RATIO: Record<string, number> = {')
    for (const r of filtered) {
      if (r.familyName) {
        console.log(`  '${r.familyName}': ${r.lineRatio},  // ascent=${r.usWinAscent} descent=${r.usWinDescent} em=${r.unitsPerEm}`)
      }
    }
    console.log('}')
    console.log('')
    console.log('const DEFAULT_LINE_RATIO = 1.2')
  } else if (isJson) {
    const output = filtered.map(r => ({
      familyName: r.familyName,
      usWinAscent: r.usWinAscent,
      usWinDescent: r.usWinDescent,
      unitsPerEm: r.unitsPerEm,
      lineRatio: r.lineRatio,
    }))
    console.log(JSON.stringify(output, null, 2))
  } else {
    // 表格输出
    const maxName = Math.max(12, ...filtered.map(r => r.familyName.length))
    const header = [
      '字体名称'.padEnd(maxName),
      'Ascent'.padStart(8),
      'Descent'.padStart(8),
      'EmSize'.padStart(8),
      'LineRatio'.padStart(10),
    ].join('  ')
    
    console.log(header)
    console.log('-'.repeat(header.length))
    
    for (const r of filtered) {
      console.log([
        r.familyName.padEnd(maxName),
        String(r.usWinAscent).padStart(8),
        String(r.usWinDescent).padStart(8),
        String(r.unitsPerEm).padStart(8),
        String(r.lineRatio).padStart(10),
      ].join('  '))
    }
    
    console.log(`\n共 ${filtered.length} 个字体`)
  }
}

main()
