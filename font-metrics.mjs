#!/usr/bin/env node

/**
 * PPT-Font-Toolkit / Font Metrics
 * 
 * 从字体文件中提取 usWinAscent、usWinDescent、unitsPerEm，
 * 计算 PPT 单倍行距比率 (usWinAscent + usWinDescent) / unitsPerEm
 * 
 * 用法：
 *   1. 扫描系统字体：  node ppt-font-toolkit.mjs metrics --scan
 *   2. 指定字体文件：  node ppt-font-metrics /path/to/font.ttf
 *   3. 指定目录扫描：  node ppt-font-toolkit.mjs metrics --dir /path/to/fonts
 *   4. 输出 JSON：     node ppt-font-metrics --scan --json
 *   5. JSON Map模式：  node ppt-font-metrics --scan --json --map
 *   6. 保存到文件：    node ppt-font-metrics --scan --json --save output.json
 *   7. 搜索字体名称：  node ppt-font-toolkit.mjs metrics --scan --filter 微软
 *   8. 兼容旧命令：    node font-metrics.mjs --scan
 */

import { readFileSync, writeFileSync, readdirSync, statSync, existsSync } from 'fs'
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
PPT-Font-Toolkit — Font Metrics

计算 PPT 单倍行距比率

用法:
  node ppt-font-toolkit.mjs metrics <字体文件路径>
  node ppt-font-toolkit.mjs metrics --dir <目录路径>
  node ppt-font-toolkit.mjs metrics --scan
  node ppt-font-metrics <字体文件路径>
  node font-metrics.mjs <字体文件路径>   (兼容旧脚本名)

参数示例:
  --scan                  扫描系统字体
  --filter <关键字>       按名称过滤
  --json                  以 JSON 格式输出（数组）
  --json --map            以 JSON 格式输出（familyName 为 key 的 Map）
  --json --save <文件路径> 输出 JSON 并保存到文件
  --code                  输出可直接使用的代码

示例:
  node ppt-font-toolkit.mjs metrics /Library/Fonts/msyh.ttf
  node ppt-font-metrics --scan --filter "微软|Arial|Times|宋体|Courier"
  node ppt-font-metrics --scan --code
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
  const isMap = args.includes('--map')
  const isScan = args.includes('--scan')
  const saveIdx = args.indexOf('--save')
  const savePath = saveIdx >= 0 ? args[saveIdx + 1] : null
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

  // 按名称过滤（同时匹配中文名、英文名、文件路径）
  let filtered = allResults
  if (filterPattern) {
    const regex = new RegExp(filterPattern, 'i')
    filtered = allResults.filter(r => 
      regex.test(r.familyName) || regex.test(r.familyNameEn) || regex.test(r.path)
    )
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
    console.log('// PPT-Font-Toolkit 字体单倍行距比率表')
    console.log('// lineRatio = (usWinAscent + usWinDescent) / unitsPerEm')
    console.log('// PPT行高(px) = fontSize(pt) × lineRatio × spcPct × 96/72')
    console.log('const FONT_LINE_RATIO: Record<string, number> = {')
    for (const r of filtered) {
      if (r.familyName) {
        const enInfo = r.familyNameEn && r.familyNameEn !== r.familyName ? ` [${r.familyNameEn}]` : ''
        console.log(`  '${r.familyName}': ${r.lineRatio},  // ascent=${r.usWinAscent} descent=${r.usWinDescent} em=${r.unitsPerEm}${enInfo}`)
      }
    }
    console.log('}')
    console.log('')
    console.log('const DEFAULT_LINE_RATIO = 1.2')
  } else if (isJson) {
    let output
    if (isMap) {
      // 以 familyName 为 key 生成 Map 对象
      output = {}
      for (const r of filtered) {
        if (r.familyName) {
          const entry = {
            usWinAscent: r.usWinAscent,
            usWinDescent: r.usWinDescent,
            unitsPerEm: r.unitsPerEm,
            lineRatio: r.lineRatio,
          }
          if (r.familyNameEn && r.familyNameEn !== r.familyName) {
            entry.familyNameEn = r.familyNameEn
          }
          output[r.familyName] = entry
        }
      }
    } else {
      output = filtered.map(r => {
        const entry = {
          familyName: r.familyName,
          usWinAscent: r.usWinAscent,
          usWinDescent: r.usWinDescent,
          unitsPerEm: r.unitsPerEm,
          lineRatio: r.lineRatio,
        }
        if (r.familyNameEn && r.familyNameEn !== r.familyName) {
          entry.familyNameEn = r.familyNameEn
        }
        return entry
      })
    }
    const jsonStr = JSON.stringify(output, null, 2)
    if (savePath) {
      const absPath = resolve(savePath)
      writeFileSync(absPath, jsonStr + '\n', 'utf-8')
      console.error(`已保存到: ${absPath}`)
    } else {
      console.log(jsonStr)
    }
  } else {
    // 表格输出
    // 生成显示名称，包含英文名（如有）
    const displayNames = filtered.map(r => {
      if (r.familyNameEn && r.familyNameEn !== r.familyName) {
        return `${r.familyName} (${r.familyNameEn})`
      }
      return r.familyName
    })
    const maxName = Math.max(12, ...displayNames.map(n => n.length))
    const header = [
      '字体名称'.padEnd(maxName),
      'Ascent'.padStart(8),
      'Descent'.padStart(8),
      'EmSize'.padStart(8),
      'LineRatio'.padStart(10),
    ].join('  ')
    
    console.log(header)
    console.log('-'.repeat(header.length))
    
    for (let i = 0; i < filtered.length; i++) {
      const r = filtered[i]
      console.log([
        displayNames[i].padEnd(maxName),
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
