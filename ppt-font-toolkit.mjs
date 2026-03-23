#!/usr/bin/env node

import { spawnSync } from 'child_process'
import { dirname, join, resolve } from 'path'
import { fileURLToPath } from 'url'

const SCRIPT_PATH = fileURLToPath(import.meta.url)
const SCRIPT_DIR = dirname(SCRIPT_PATH)

function printUsage() {
  console.log(`
PPT-Font-Toolkit

用法:
  node ppt-font-toolkit.mjs recover <source ...>
  node ppt-font-toolkit.mjs metrics <args ...>
  node ppt-font-toolkit.mjs <source ...>

说明:
  - 省略子命令时，默认走 recover
  - recover 对应: ppt-font-recover.mjs
  - metrics 对应: ppt-font-metrics / font-metrics.mjs（兼容）
`)
}

function runScript(scriptName, args) {
  const scriptPath = join(SCRIPT_DIR, scriptName)
  const result = spawnSync(process.execPath, [scriptPath, ...args], {
    stdio: 'inherit',
  })

  if (result.error) {
    throw result.error
  }

  process.exit(result.status ?? 1)
}

function main() {
  const args = process.argv.slice(2)
  if (args.length === 0 || args[0] === '--help' || args[0] === '-h' || args[0] === 'help') {
    printUsage()
    process.exit(0)
  }

  const [firstArg, ...restArgs] = args
  if (firstArg === 'metrics') {
    runScript('font-metrics.mjs', restArgs)
    return
  }

  if (firstArg === 'recover') {
    runScript('ppt-font-recover.mjs', restArgs)
    return
  }

  runScript('ppt-font-recover.mjs', args)
}

main()
