#!/usr/bin/env node

import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'fs'
import { tmpdir, platform as osPlatform, arch as osArch } from 'os'
import { basename, dirname, extname, join, resolve } from 'path'
import { fileURLToPath } from 'url'
import { spawnSync } from 'child_process'

const SCRIPT_PATH = fileURLToPath(import.meta.url)
const SCRIPT_DIR = dirname(SCRIPT_PATH)
const ARCHIVE_EXTENSIONS = new Set(['.pptx', '.pptm', '.ppsx', '.ppsm', '.potx', '.potm', '.zip'])
const FONT_EXTENSIONS = new Set(['.fntdata'])
const WPS_BINARY_NAME = `wps-eot-extract-${osPlatform()}-${osArch()}${osPlatform() === 'win32' ? '.exe' : ''}`
const WPS_BINARY_PATH = join(SCRIPT_DIR, '.cache', 'bin', WPS_BINARY_NAME)
const WPS_SOURCE_ROOT = join(SCRIPT_DIR, 'vendor', 'libeot')
const WPS_SOURCE_FILES = [
  'vendor/libeot/wps-eot-extract.c',
  'vendor/libeot/src/triplet_encodings.c',
  'vendor/libeot/src/ctf/SFNTContainer.c',
  'vendor/libeot/src/ctf/parseCTF.c',
  'vendor/libeot/src/ctf/parseTTF.c',
  'vendor/libeot/src/lzcomp/liblzcomp.c',
  'vendor/libeot/src/lzcomp/lzcomp.c',
  'vendor/libeot/src/lzcomp/ahuff.c',
  'vendor/libeot/src/lzcomp/bitio.c',
  'vendor/libeot/src/lzcomp/mtxmem.c',
  'vendor/libeot/src/util/stream.c',
]

function printUsage() {
  console.log(`
PPT 内嵌字体恢复工具

功能:
  - 自动识别并恢复 Microsoft Office ODTTF（fontKey + XOR）
  - 自动识别并恢复 WPS / Kingsoft 压缩 EOT（自动编译本地 helper）
  - 支持单文件、解压目录、.pptx/.pptm/.ppsx 归档、目录批量扫描

用法:
  node ppt-font-recover.mjs <source>
  node ppt-font-recover.mjs <source1> <source2> ...
  node ppt-font-recover.mjs --input <source>
  node ppt-font-recover.mjs --ppt-dir <解压后的PPT目录>

source 可以是:
  - .pptx / .pptm / .ppsx / .potx
  - 已解压的 PPT 根目录，或其中的 ppt 目录
  - 单个 .fntdata 文件
  - 包含上述内容的目录（会自动递归扫描并批量处理）

常用参数:
  --input <path>       添加输入源，可重复
  --ppt-dir <path>     兼容旧参数，等价于添加一个 package 输入
  --font <value>       仅处理指定字体；支持 rId / 目标文件名 / 字体名，可重复或逗号分隔
  --key <guid>         单个 Office .fntdata 直解时手动提供 fontKey
  --output <path>      单任务输出文件路径
  --output-dir <path>  批量输出目录
  --list               仅列出嵌入字体和识别结果，不执行恢复
  --json               以 JSON 输出结果
  --keep-temp          保留 .pptx 解压临时目录
  --help, -h           显示帮助

示例:
  node ppt-font-recover.mjs ./demo.pptx
  node ppt-font-recover.mjs ./demo-unzipped --list
  node ppt-font-recover.mjs ./demo-unzipped --font "rId10,汉仪汉黑简"
  node ppt-font-recover.mjs ./ppt/fonts/font1.fntdata
  node ppt-font-recover.mjs ./ppt/fonts/font1.fntdata --key "{A1B2C3D4-E5F6-1234-ABCD-1234567890AB}"
  node ppt-font-recover.mjs ./downloads --output-dir ./recovered-fonts
`)
}

export function parseRecoverArgs(argv) {
  const options = {
    sources: [],
    fontRefs: [],
    output: null,
    outputDir: null,
    key: null,
    list: false,
    json: false,
    keepTemp: false,
    help: false,
  }

  for (let index = 0; index < argv.length; index++) {
    const arg = argv[index]

    if (arg === '--help' || arg === '-h') {
      options.help = true
    } else if (arg === '--list') {
      options.list = true
    } else if (arg === '--json') {
      options.json = true
    } else if (arg === '--keep-temp') {
      options.keepTemp = true
    } else if (arg === '--input') {
      options.sources.push(readNextArg(argv, ++index, '--input'))
    } else if (arg === '--ppt-dir') {
      options.sources.push(readNextArg(argv, ++index, '--ppt-dir'))
    } else if (arg === '--font') {
      options.fontRefs.push(...splitMultiValue(readNextArg(argv, ++index, '--font')))
    } else if (arg === '--output') {
      options.output = readNextArg(argv, ++index, '--output')
    } else if (arg === '--output-dir') {
      options.outputDir = readNextArg(argv, ++index, '--output-dir')
    } else if (arg === '--key') {
      options.key = readNextArg(argv, ++index, '--key')
    } else if (!arg.startsWith('-')) {
      options.sources.push(arg)
    } else {
      throw new Error(`未知参数: ${arg}`)
    }
  }

  return options
}

function readNextArg(argv, index, flagName) {
  const value = argv[index]
  if (!value) {
    throw new Error(`${flagName} 缺少参数`)
  }
  return value
}

function splitMultiValue(value) {
  return String(value)
    .split(/[,\u3001;；]+/)
    .map((item) => item.trim())
    .filter(Boolean)
}

function readText(filePath) {
  return readFileSync(filePath, 'utf8')
}

function parseAttributes(source) {
  const attributes = {}
  const pattern = /([A-Za-z0-9:_-]+)="([^"]*)"/g
  let match = pattern.exec(source)
  while (match) {
    attributes[match[1]] = match[2]
    match = pattern.exec(source)
  }
  return attributes
}

function readUInt16LE(buffer, offset) {
  return buffer.readUInt16LE(offset)
}

function readUInt32LE(buffer, offset) {
  return buffer.readUInt32LE(offset)
}

function normalizeFsPath(filePath) {
  const resolved = resolve(filePath).replace(/\\/g, '/')
  return osPlatform() === 'win32' ? resolved.toLowerCase() : resolved
}

function sanitizeFileName(name) {
  const fallback = 'font'
  const sanitized = String(name || fallback)
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
  return sanitized || fallback
}

function looksLikeSfnt(buffer) {
  if (buffer.length < 4) return false
  const signature = buffer.subarray(0, 4)
  return (
    (signature[0] === 0x00 && signature[1] === 0x01 && signature[2] === 0x00 && signature[3] === 0x00) ||
    (signature[0] === 0x4f && signature[1] === 0x54 && signature[2] === 0x54 && signature[3] === 0x4f) ||
    (signature[0] === 0x74 && signature[1] === 0x74 && signature[2] === 0x63 && signature[3] === 0x66) ||
    (signature[0] === 0x74 && signature[1] === 0x72 && signature[2] === 0x75 && signature[3] === 0x65)
  )
}

function readUtf16String(buffer, offset, byteLength) {
  if (byteLength <= 0 || offset + byteLength > buffer.length || byteLength % 2 !== 0) {
    return ''
  }
  return buffer.subarray(offset, offset + byteLength).toString('utf16le').replace(/\u0000+$/g, '')
}

function parseWpsEotMetadata(buffer) {
  try {
    let offset = 82
    const familySize = readUInt16LE(buffer, offset)
    offset += 2
    const familyName = readUtf16String(buffer, offset, familySize)
    offset += familySize + 2
    const styleSize = readUInt16LE(buffer, offset)
    offset += 2
    const styleName = readUtf16String(buffer, offset, styleSize)
    return {
      familyName,
      styleName,
    }
  } catch {
    return {
      familyName: '',
      styleName: '',
    }
  }
}

export function inspectEmbeddedFontFile(inputPath) {
  const absoluteInput = resolve(inputPath)
  if (!existsSync(absoluteInput)) {
    throw new Error(`找不到输入文件: ${absoluteInput}`)
  }

  const buffer = readFileSync(absoluteInput)
  if (buffer.length >= 36) {
    const totalSize = readUInt32LE(buffer, 0)
    const fontDataSize = readUInt32LE(buffer, 4)
    const version = readUInt32LE(buffer, 8)
    const magic = readUInt16LE(buffer, 34)
    const knownVersions = new Set([0x00010000, 0x00020001, 0x00020002])

    if (
      totalSize === buffer.length &&
      fontDataSize > 0 &&
      totalSize >= fontDataSize &&
      magic === 0x504c &&
      knownVersions.has(version)
    ) {
      return {
        input: absoluteInput,
        buffer,
        kind: 'wps-eot',
        totalSize,
        fontDataSize,
        mtxOffset: totalSize - fontDataSize,
        ...parseWpsEotMetadata(buffer),
      }
    }
  }

  if (looksLikeSfnt(buffer)) {
    return {
      input: absoluteInput,
      buffer,
      kind: 'plain-sfnt',
    }
  }

  return {
    input: absoluteInput,
    buffer,
    kind: 'office-odttf',
  }
}

function detectFontOutputExtension(buffer) {
  const signature = buffer.subarray(0, 4)

  if (signature[0] === 0x4f && signature[1] === 0x54 && signature[2] === 0x54 && signature[3] === 0x4f) {
    return '.otf'
  }

  if (signature[0] === 0x74 && signature[1] === 0x74 && signature[2] === 0x63 && signature[3] === 0x66) {
    return '.ttc'
  }

  return '.ttf'
}

function resolvePptPaths(pptDir) {
  const base = resolve(pptDir)
  const presentationFromRoot = join(base, 'ppt', 'presentation.xml')
  const relsFromRoot = join(base, 'ppt', '_rels', 'presentation.xml.rels')

  if (existsSync(presentationFromRoot) && existsSync(relsFromRoot)) {
    return {
      rootDir: base,
      pptDir: join(base, 'ppt'),
      presentationPath: presentationFromRoot,
      relsPath: relsFromRoot,
    }
  }

  const presentationFromPpt = join(base, 'presentation.xml')
  const relsFromPpt = join(base, '_rels', 'presentation.xml.rels')

  if (existsSync(presentationFromPpt) && existsSync(relsFromPpt)) {
    return {
      rootDir: dirname(base),
      pptDir: base,
      presentationPath: presentationFromPpt,
      relsPath: relsFromPpt,
    }
  }

  throw new Error(`未找到 presentation.xml: ${base}`)
}

function isPptPackageDirectory(dirPath) {
  try {
    resolvePptPaths(dirPath)
    return true
  } catch {
    return false
  }
}

function resolveMethodLabel(fileInfo, fontKey) {
  if (fileInfo.kind === 'wps-eot') return 'wps-eot'
  if (fileInfo.kind === 'plain-sfnt') return 'plain-sfnt'
  return fontKey ? 'office-odttf' : 'office-odttf (need key)'
}

function readPptFontMappings(pptDir) {
  const { pptDir: actualPptDir, presentationPath, relsPath } = resolvePptPaths(pptDir)
  const presentationXml = readText(presentationPath)
  const relsXml = readText(relsPath)

  const relationshipMap = new Map()
  const relationshipPattern = /<Relationship\b([^>]*)\/?>/g
  let relationshipMatch = relationshipPattern.exec(relsXml)
  while (relationshipMatch) {
    const attributes = parseAttributes(relationshipMatch[1])
    if (attributes.Id && attributes.Target) {
      relationshipMap.set(attributes.Id, attributes.Target)
    }
    relationshipMatch = relationshipPattern.exec(relsXml)
  }

  const results = []
  const embeddedFontPattern = /<p:embeddedFont\b([^>]*)>([\s\S]*?)<\/p:embeddedFont>/g
  let embeddedFontMatch = embeddedFontPattern.exec(presentationXml)
  while (embeddedFontMatch) {
    const fontAttributes = parseAttributes(embeddedFontMatch[1])
    const body = embeddedFontMatch[2]
    const fontTagMatch = body.match(/<p:font\b([^>]*)\/?>/)
    const fontTagAttributes = fontTagMatch ? parseAttributes(fontTagMatch[1]) : {}
    const typeface = fontTagAttributes.typeface || fontAttributes.typeface || ''
    const stylePattern = /<p:(regular|italic|bold|boldItalic)\b([^>]*)\/?>/g
    let styleMatch = stylePattern.exec(body)
    while (styleMatch) {
      const style = styleMatch[1]
      const attributes = parseAttributes(styleMatch[2])
      const rId = attributes['r:id'] || ''
      const fontKey = attributes.fontKey || ''
      const target = relationshipMap.get(rId) || ''
      const inputPath = target ? resolve(actualPptDir, target) : null
      const fileInfo = inputPath && existsSync(inputPath) ? inspectEmbeddedFontFile(inputPath) : null

      results.push({
        typeface,
        style,
        rId,
        fontKey,
        target,
        inputPath,
        fileInfo,
        methodLabel: fileInfo ? resolveMethodLabel(fileInfo, fontKey) : (fontKey ? 'office-odttf' : 'unknown'),
      })

      styleMatch = stylePattern.exec(body)
    }
    embeddedFontMatch = embeddedFontPattern.exec(presentationXml)
  }

  return results
}

function matchFontRef(mapping, fontRef) {
  const normalizedRef = String(fontRef).replace(/\\/g, '/').toLowerCase()
  const target = (mapping.target || '').replace(/\\/g, '/').toLowerCase()
  const input = mapping.inputPath ? mapping.inputPath.replace(/\\/g, '/').toLowerCase() : ''
  const inputBase = input ? basename(input) : ''
  const targetBase = target ? basename(target) : ''

  return (
    mapping.rId.toLowerCase() === normalizedRef ||
    mapping.typeface.toLowerCase() === normalizedRef ||
    target === normalizedRef ||
    targetBase === normalizedRef ||
    input === normalizedRef ||
    inputBase === normalizedRef
  )
}

function filterMappings(mappings, fontRefs) {
  if (!fontRefs.length) return mappings
  const filtered = mappings.filter((mapping) => fontRefs.some((fontRef) => matchFontRef(mapping, fontRef)))
  if (!filtered.length) {
    throw new Error(`没有找到与 ${fontRefs.join(', ')} 对应的内嵌字体`)
  }
  return filtered
}

function findNearbyPackageMapping(inputPath) {
  const absoluteInput = resolve(inputPath)
  let current = dirname(absoluteInput)

  for (let depth = 0; depth < 8; depth++) {
    try {
      const mappings = readPptFontMappings(current)
      const normalizedInput = normalizeFsPath(absoluteInput)
      const matched = mappings.filter((mapping) => mapping.inputPath && normalizeFsPath(mapping.inputPath) === normalizedInput)
      if (matched.length === 1) {
        return matched[0]
      }
    } catch {
      // ignore and keep climbing
    }

    const parent = dirname(current)
    if (parent === current) break
    current = parent
  }

  return null
}

function normalizeGuid(fontKey) {
  const cleanGuid = String(fontKey || '').replace(/[^a-fA-F0-9]/g, '')
  if (cleanGuid.length !== 32) {
    throw new Error('无效的 fontKey：提取出的十六进制字符必须正好是 32 位')
  }

  const guidBytes = Buffer.alloc(16)
  for (let index = 0; index < 16; index++) {
    guidBytes[index] = Number.parseInt(cleanGuid.slice(index * 2, index * 2 + 2), 16)
  }
  guidBytes.reverse()
  return guidBytes
}

function recoverOfficeBuffer(inputPath, fontKey) {
  const guidBytes = normalizeGuid(fontKey)
  const fontData = Buffer.from(readFileSync(inputPath))
  const limit = Math.min(32, fontData.length)
  for (let index = 0; index < limit; index++) {
    fontData[index] = fontData[index] ^ guidBytes[index % 16]
  }
  return fontData
}

function findCompiler() {
  for (const compiler of ['cc', 'clang', 'gcc']) {
    const result = spawnSync(compiler, ['--version'], { encoding: 'utf8' })
    if (!result.error) {
      return compiler
    }
  }
  return null
}

function needsWpsHelperRebuild(binaryPath) {
  if (!existsSync(binaryPath)) return true
  const binaryMtime = statSync(binaryPath).mtimeMs
  return WPS_SOURCE_FILES.some((relativePath) => {
    const absolutePath = join(SCRIPT_DIR, relativePath)
    return existsSync(absolutePath) && statSync(absolutePath).mtimeMs > binaryMtime
  })
}

function ensureWpsHelper() {
  if (!existsSync(WPS_SOURCE_ROOT)) {
    throw new Error('缺少 vendor/libeot 源码，无法构建 WPS 解包 helper')
  }

  if (!needsWpsHelperRebuild(WPS_BINARY_PATH)) {
    return WPS_BINARY_PATH
  }

  const compiler = findCompiler()
  if (!compiler) {
    throw new Error('WPS 模式需要本机可用的 C 编译器（cc / clang / gcc）')
  }

  mkdirSync(dirname(WPS_BINARY_PATH), { recursive: true })

  const compileArgs = [
    '-O2',
    '-std=gnu99',
    '-DDECOMPRESS_ON',
    '-Wno-implicit-function-declaration',
    '-I',
    WPS_SOURCE_ROOT,
    '-I',
    join(WPS_SOURCE_ROOT, 'inc'),
    '-o',
    WPS_BINARY_PATH,
    ...WPS_SOURCE_FILES.map((relativePath) => join(SCRIPT_DIR, relativePath)),
  ]

  const result = spawnSync(compiler, compileArgs, {
    encoding: 'utf8',
  })

  if (result.error || result.status !== 0) {
    const detail = [result.stdout, result.stderr].filter(Boolean).join('\n').trim()
    throw new Error(`构建 WPS helper 失败${detail ? `:\n${detail}` : ''}`)
  }

  return WPS_BINARY_PATH
}

function recoverWpsBuffer(inputPath) {
  const helperPath = ensureWpsHelper()
  const tempDir = mkdtempSync(join(tmpdir(), 'ppt-font-toolkit-wps-'))
  const tempOutput = join(tempDir, 'recovered.ttf')

  try {
    const result = spawnSync(helperPath, [resolve(inputPath), tempOutput], {
      encoding: 'utf8',
    })

    if (result.error || result.status !== 0) {
      const detail = [result.stdout, result.stderr].filter(Boolean).join('\n').trim()
      throw new Error(`WPS 字体解包失败${detail ? `:\n${detail}` : ''}`)
    }

    return readFileSync(tempOutput)
  } finally {
    rmSync(tempDir, { recursive: true, force: true })
  }
}

function escapePowerShellValue(value) {
  return String(value).replace(/'/g, "''")
}

function extractArchiveToTemp(archivePath) {
  const tempRoot = mkdtempSync(join(tmpdir(), 'ppt-font-toolkit-'))
  const outputDir = join(tempRoot, 'package')
  mkdirSync(outputDir, { recursive: true })

  let result
  if (osPlatform() === 'win32') {
    const command = `Expand-Archive -LiteralPath '${escapePowerShellValue(resolve(archivePath))}' -DestinationPath '${escapePowerShellValue(outputDir)}' -Force`
    result = spawnSync('powershell', ['-NoProfile', '-Command', command], { encoding: 'utf8' })
  } else {
    result = spawnSync('unzip', ['-qq', '-o', resolve(archivePath), '-d', outputDir], { encoding: 'utf8' })
  }

  if (result.error || result.status !== 0) {
    const detail = [result.stdout, result.stderr].filter(Boolean).join('\n').trim()
    rmSync(tempRoot, { recursive: true, force: true })
    throw new Error(`解压归档失败${detail ? `:\n${detail}` : ''}`)
  }

  return {
    packagePath: outputDir,
    cleanupPath: tempRoot,
  }
}

function sourceLabelForPath(kind, inputPath) {
  if (kind === 'archive') {
    return basename(inputPath, extname(inputPath))
  }
  if (kind === 'font-file') {
    return basename(inputPath, extname(inputPath))
  }

  const absolutePath = resolve(inputPath)
  if (basename(absolutePath).toLowerCase() === 'ppt') {
    return basename(dirname(absolutePath))
  }
  return basename(absolutePath)
}

function collectInputEntries(inputPath) {
  const absolutePath = resolve(inputPath)
  if (!existsSync(absolutePath)) {
    throw new Error(`输入不存在: ${absolutePath}`)
  }

  const stat = statSync(absolutePath)
  if (stat.isFile()) {
    const extension = extname(absolutePath).toLowerCase()
    if (ARCHIVE_EXTENSIONS.has(extension)) {
      return [{ kind: 'archive', originalPath: absolutePath, label: sourceLabelForPath('archive', absolutePath) }]
    }
    if (FONT_EXTENSIONS.has(extension)) {
      return [{ kind: 'font-file', originalPath: absolutePath, label: sourceLabelForPath('font-file', absolutePath) }]
    }
    throw new Error(`暂不支持的文件类型: ${absolutePath}`)
  }

  if (isPptPackageDirectory(absolutePath)) {
    return [{ kind: 'package', originalPath: absolutePath, label: sourceLabelForPath('package', absolutePath) }]
  }

  const results = []
  for (const entry of readdirSync(absolutePath)) {
    const fullPath = join(absolutePath, entry)
    const childStat = statSync(fullPath)
    if (childStat.isDirectory()) {
      results.push(...collectInputEntries(fullPath))
    } else {
      const extension = extname(fullPath).toLowerCase()
      if (ARCHIVE_EXTENSIONS.has(extension)) {
        results.push({ kind: 'archive', originalPath: fullPath, label: sourceLabelForPath('archive', fullPath) })
      } else if (FONT_EXTENSIONS.has(extension)) {
        results.push({ kind: 'font-file', originalPath: fullPath, label: sourceLabelForPath('font-file', fullPath) })
      }
    }
  }

  return results
}

function dedupeSources(entries) {
  const seen = new Set()
  return entries.filter((entry) => {
    const key = `${entry.kind}:${normalizeFsPath(entry.originalPath)}`
    if (seen.has(key)) return false
    seen.add(key)
    return true
  })
}

function buildPackageTasks(source, options) {
  const mappings = readPptFontMappings(source.workingPath)
  const selectedMappings = filterMappings(mappings, options.fontRefs)

  return selectedMappings.map((mapping) => {
    if (!mapping.inputPath) {
      throw new Error(`缺少嵌入字体文件: ${mapping.target || mapping.rId || source.originalPath}`)
    }

    if (mapping.fileInfo?.kind !== 'wps-eot' && mapping.fileInfo?.kind !== 'plain-sfnt' && !mapping.fontKey) {
      throw new Error(`缺少 fontKey: ${mapping.typeface || mapping.rId || mapping.target}`)
    }

    return {
      sourceLabel: source.label,
      sourceKind: source.kind,
      sourcePath: source.originalPath,
      inputPath: mapping.inputPath,
      typeface: mapping.typeface || mapping.fileInfo?.familyName || basename(mapping.inputPath, extname(mapping.inputPath)),
      style: mapping.style || mapping.fileInfo?.styleName || 'regular',
      rId: mapping.rId,
      target: mapping.target,
      fontKey: mapping.fontKey || null,
      fileInfo: mapping.fileInfo || inspectEmbeddedFontFile(mapping.inputPath),
      method: mapping.fileInfo?.kind === 'wps-eot' ? 'wps' : mapping.fileInfo?.kind === 'plain-sfnt' ? 'plain' : 'office',
      methodLabel: mapping.methodLabel,
    }
  })
}

function buildSingleFileTask(source, options, allowMissingKey = false) {
  const fileInfo = inspectEmbeddedFontFile(source.originalPath)
  const nearbyMapping = findNearbyPackageMapping(source.originalPath)
  const fontKey = options.key || nearbyMapping?.fontKey || null
  const method = fileInfo.kind === 'wps-eot' ? 'wps' : fileInfo.kind === 'plain-sfnt' ? 'plain' : 'office'

  if (!allowMissingKey && method === 'office' && !fontKey) {
    throw new Error(`缺少 fontKey：${source.originalPath}。请传 --key，或直接把整个 PPT / 解压目录作为输入`)
  }

  return {
    sourceLabel: source.label,
    sourceKind: source.kind,
    sourcePath: source.originalPath,
    inputPath: source.originalPath,
    typeface: nearbyMapping?.typeface || fileInfo.familyName || basename(source.originalPath, extname(source.originalPath)),
    style: nearbyMapping?.style || fileInfo.styleName || 'regular',
    rId: nearbyMapping?.rId || '',
    target: nearbyMapping?.target || basename(source.originalPath),
    fontKey,
    fileInfo,
    method,
    methodLabel: resolveMethodLabel(fileInfo, fontKey),
  }
}

function defaultStemForTask(task) {
  const baseName = sanitizeFileName(task.typeface || basename(task.inputPath, extname(task.inputPath)))
  const style = String(task.style || '').trim()
  if (style && style.toLowerCase() !== 'regular') {
    return sanitizeFileName(`${baseName}-${style}`)
  }
  return baseName
}

function ensureUniqueFilePath(targetPath) {
  if (!existsSync(targetPath)) return targetPath

  const extension = extname(targetPath)
  const base = targetPath.slice(0, -extension.length)
  let counter = 2
  let candidate = `${base}-${counter}${extension}`
  while (existsSync(candidate)) {
    counter += 1
    candidate = `${base}-${counter}${extension}`
  }
  return candidate
}

function defaultOutputDirForSource(source) {
  if (source.kind === 'font-file') {
    return dirname(source.originalPath)
  }
  return join(dirname(source.originalPath), `${sanitizeFileName(source.label)}-recovered-fonts`)
}

function resolveSourceOutputDir(source, options, totalSources) {
  if (!options.outputDir) {
    return defaultOutputDirForSource(source)
  }

  const absoluteOutputDir = resolve(options.outputDir)
  if (totalSources > 1) {
    return join(absoluteOutputDir, sanitizeFileName(source.label))
  }
  return absoluteOutputDir
}

function recoverTaskBuffer(task) {
  if (task.method === 'wps') {
    return recoverWpsBuffer(task.inputPath)
  }
  if (task.method === 'plain') {
    return Buffer.from(readFileSync(task.inputPath))
  }
  return recoverOfficeBuffer(task.inputPath, task.fontKey)
}

function writeRecoveredTask(task, buffer, outputPath, outputDir) {
  const extension = detectFontOutputExtension(buffer)
  let finalPath

  if (outputPath) {
    const absoluteOutputPath = resolve(outputPath)
    const withoutExt = extname(absoluteOutputPath)
      ? absoluteOutputPath.slice(0, -extname(absoluteOutputPath).length)
      : absoluteOutputPath
    finalPath = `${withoutExt}${extension}`
  } else {
    mkdirSync(outputDir, { recursive: true })
    finalPath = ensureUniqueFilePath(join(outputDir, `${defaultStemForTask(task)}${extension}`))
  }

  mkdirSync(dirname(finalPath), { recursive: true })
  writeFileSync(finalPath, buffer)
  return finalPath
}

function tableString(rows, headers) {
  const widths = headers.map((header, index) => {
    return Math.max(header.length, ...rows.map((row) => String(row[index] ?? '').length))
  })
  const head = headers.map((header, index) => header.padEnd(widths[index])).join('  ')
  const separator = '-'.repeat(head.length)
  const body = rows.map((row) => row.map((value, index) => String(value ?? '').padEnd(widths[index])).join('  '))
  return [head, separator, ...body].join('\n')
}

function printListReport(report, showTitle = false) {
  if (showTitle) {
    console.log(`# ${report.sourceLabel}`)
  }

  const rows = report.rows.map((item) => [
    item.typeface || '-',
    item.style || '-',
    item.rId || '-',
    item.target || '-',
    item.methodLabel || item.method || '-',
    item.fontKey || '-',
  ])

  console.log(tableString(rows, ['Typeface', 'Style', 'rId', 'Target', 'Method', 'fontKey']))
}

function printRecoveryResults(results) {
  for (const result of results) {
    const typeface = result.typeface || basename(result.output, extname(result.output))
    console.log(`✅ [${result.methodLabel || result.method}] ${typeface} -> ${result.output}`)
  }
}

function isDirectRun() {
  return process.argv[1] && resolve(process.argv[1]) === SCRIPT_PATH
}

function emitRecoverEvent(callbacks, type, payload = {}) {
  const event = { type, ...payload }

  if (typeof callbacks.onEvent === 'function') {
    callbacks.onEvent(event)
  }

  if (type === 'tempDir' && typeof callbacks.onTempDir === 'function') {
    callbacks.onTempDir(event)
  } else if (type === 'report' && typeof callbacks.onReport === 'function') {
    callbacks.onReport(event)
  } else if (type === 'recovered' && typeof callbacks.onRecovered === 'function') {
    callbacks.onRecovered(event)
  } else if (type === 'complete' && typeof callbacks.onComplete === 'function') {
    callbacks.onComplete(event)
  }
}

function normalizeRecoverOptions(options = {}) {
  const rawSources = []
  const rawFontRefs = []

  const appendValues = (target, value) => {
    if (Array.isArray(value)) {
      for (const item of value) appendValues(target, item)
      return
    }
    if (value !== undefined && value !== null && value !== '') {
      target.push(String(value))
    }
  }

  appendValues(rawSources, options.sources)
  appendValues(rawSources, options.inputs)
  appendValues(rawSources, options.source)
  appendValues(rawSources, options.input)

  appendValues(rawFontRefs, options.fontRefs)
  appendValues(rawFontRefs, options.fonts)
  appendValues(rawFontRefs, options.font)

  return {
    sources: rawSources,
    fontRefs: rawFontRefs.flatMap((value) => splitMultiValue(value)),
    output: options.output || null,
    outputDir: options.outputDir || null,
    key: options.key || null,
    list: Boolean(options.list),
    json: Boolean(options.json),
    keepTemp: Boolean(options.keepTemp),
    help: Boolean(options.help),
  }
}

function cleanupTemporaryPaths(temporaryPaths) {
  for (const tempPath of temporaryPaths) {
    try {
      rmSync(tempPath, { recursive: true, force: true })
    } catch {
      // ignore cleanup failures
    }
  }
}

function resolveRecoverSources(options, temporaryPaths, callbacks) {
  const rawSources = dedupeSources(options.sources.flatMap((source) => collectInputEntries(source)))
  if (!rawSources.length) {
    throw new Error('没有找到可处理的 .pptx / package / .fntdata 输入')
  }

  return rawSources.map((source) => {
    if (source.kind !== 'archive') {
      return {
        ...source,
        workingPath: source.originalPath,
      }
    }

    const extracted = extractArchiveToTemp(source.originalPath)
    temporaryPaths.push(extracted.cleanupPath)
    emitRecoverEvent(callbacks, 'tempDir', {
      sourceLabel: source.label,
      sourcePath: source.originalPath,
      path: extracted.cleanupPath,
      packagePath: extracted.packagePath,
    })
    return {
      ...source,
      workingPath: extracted.packagePath,
      extractedPath: extracted.packagePath,
    }
  })
}

function createListReports(resolvedSources, options) {
  const reports = []

  for (const source of resolvedSources) {
    if (source.kind === 'package' || source.kind === 'archive') {
      const mappings = readPptFontMappings(source.workingPath)
      reports.push({
        sourceLabel: source.label,
        rows: filterMappings(mappings, options.fontRefs).map((item) => ({
          typeface: item.typeface || '',
          style: item.style || '',
          rId: item.rId || '',
          target: item.target || '',
          method: item.methodLabel || '',
          methodLabel: item.methodLabel || '',
          fontKey: item.fontKey || '',
          inputPath: item.inputPath || '',
        })),
      })
    } else {
      const item = buildSingleFileTask(source, { ...options, key: options.key || null }, true)
      reports.push({
        sourceLabel: source.label,
        rows: [{
          typeface: item.typeface || '',
          style: item.style || '',
          rId: item.rId || '',
          target: item.target || '',
          method: item.methodLabel || '',
          methodLabel: item.methodLabel || '',
          fontKey: item.fontKey || '',
          inputPath: item.inputPath || '',
        }],
      })
    }
  }

  return reports
}

export function listEmbeddedFonts(options = {}, callbacks = {}) {
  const normalized = normalizeRecoverOptions(options)
  const temporaryPaths = []

  try {
    const resolvedSources = resolveRecoverSources(normalized, temporaryPaths, callbacks)
    const reports = createListReports(resolvedSources, normalized)

    for (const report of reports) {
      emitRecoverEvent(callbacks, 'report', report)
    }

    const summary = {
      reports,
      temporaryPaths: [...temporaryPaths],
    }
    emitRecoverEvent(callbacks, 'complete', { mode: 'list', ...summary })
    return summary
  } finally {
    if (!normalized.keepTemp) {
      cleanupTemporaryPaths(temporaryPaths)
    }
  }
}

export function recoverEmbeddedFonts(options = {}, callbacks = {}) {
  const normalized = normalizeRecoverOptions(options)
  const temporaryPaths = []

  try {
    const resolvedSources = resolveRecoverSources(normalized, temporaryPaths, callbacks)

    const tasks = []
    for (const source of resolvedSources) {
      if (source.kind === 'package' || source.kind === 'archive') {
        tasks.push(...buildPackageTasks(source, normalized))
      } else {
        tasks.push(buildSingleFileTask(source, normalized))
      }
    }

    if (!tasks.length) {
      throw new Error('没有匹配到可恢复的内嵌字体')
    }

    if (normalized.output && tasks.length !== 1) {
      throw new Error('--output 只能用于单个恢复任务；批量模式请使用 --output-dir')
    }

    const results = []
    for (const task of tasks) {
      const buffer = recoverTaskBuffer(task)
      const matchingSource = resolvedSources.find((source) => normalizeFsPath(source.originalPath) === normalizeFsPath(task.sourcePath))
      const outputDir = resolveSourceOutputDir(matchingSource, normalized, resolvedSources.length)
      const outputPath = writeRecoveredTask(task, buffer, normalized.output, outputDir)
      const result = {
        sourceLabel: task.sourceLabel,
        inputPath: task.inputPath,
        output: outputPath,
        typeface: task.typeface,
        style: task.style,
        rId: task.rId,
        target: task.target,
        fontKey: task.fontKey || '',
        methodLabel: task.methodLabel,
        format: detectFontOutputExtension(buffer).slice(1),
      }
      results.push(result)
      emitRecoverEvent(callbacks, 'recovered', result)
    }

    const summary = {
      results,
      temporaryPaths: [...temporaryPaths],
    }
    emitRecoverEvent(callbacks, 'complete', { mode: 'recover', ...summary })
    return summary
  } finally {
    if (!normalized.keepTemp) {
      cleanupTemporaryPaths(temporaryPaths)
    }
  }
}

export function runRecover(options = {}, callbacks = {}) {
  const normalized = normalizeRecoverOptions(options)
  if (normalized.list) {
    return { mode: 'list', ...listEmbeddedFonts(normalized, callbacks) }
  }
  return { mode: 'recover', ...recoverEmbeddedFonts(normalized, callbacks) }
}

export function main(argv = process.argv.slice(2)) {
  try {
    const options = parseRecoverArgs(argv)

    if (options.help || options.sources.length === 0) {
      printUsage()
      return 0
    }

    if (options.list) {
      const { reports, temporaryPaths: keptPaths } = listEmbeddedFonts(options)
      if (options.json) {
        console.log(JSON.stringify(reports, null, 2))
      } else {
        for (let index = 0; index < reports.length; index++) {
          if (index > 0) console.log('')
          printListReport(reports[index], reports.length > 1)
        }
      }

      if (options.keepTemp && keptPaths.length) {
        console.error(`已保留临时解压目录: ${keptPaths.join(', ')}`)
      }
      return 0
    }

    const { results, temporaryPaths: keptPaths } = recoverEmbeddedFonts(options)
    if (options.json) {
      console.log(JSON.stringify(results, null, 2))
    } else {
      printRecoveryResults(results)
    }

    if (options.keepTemp && keptPaths.length) {
      console.error(`已保留临时解压目录: ${keptPaths.join(', ')}`)
    }

    return 0
  } catch (error) {
    console.error(`❌ ${error.message}`)
    return 1
  }
}

if (isDirectRun()) {
  process.exit(main())
}
