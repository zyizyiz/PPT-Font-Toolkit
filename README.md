# PPT-Font-Toolkit

中文优先的 PPT 字体工具箱（Node.js CLI）。

现在这个项目不只是做字体 `Metrics` 提取，还支持从 `Microsoft Office` 和 `WPS / Kingsoft Presentation` 的演示文稿里恢复内嵌字体。

当前包含两类能力：

- `recover`：恢复 `.pptx / .pptm / .ppsx / .potx` 里的内嵌字体
- `metrics`：提取字体 `usWinAscent / usWinDescent / unitsPerEm / lineRatio`

---

## 中文说明

### 功能特性

- 自动识别 `Microsoft Office ODTTF` 与 `WPS / Kingsoft EOT`
- 支持输入 `.pptx`、已解压目录、单个 `.fntdata`
- 支持目录递归扫描，批量处理多个文件 / 多个 PPT
- 支持 `--list` 先查看嵌入字体清单，再决定是否恢复
- 支持按 `rId`、字体名、目标文件名筛选指定字体
- 自动识别输出格式：`.ttf` / `.otf` / `.ttc`
- 保留原有字体 Metrics 提取能力
- 不依赖第三方 npm 包

### 自动识别逻辑

- `Office`：如果是标准 `ODTTF`，脚本会使用 `fontKey + XOR` 恢复
- `WPS`：如果是压缩 `EOT / MTX`，脚本会自动走 `WPS helper` 解包流程
- `普通字体文件`：如果已经是标准 `sfnt` 字体，会直接按明文字体处理

### 运行环境

- Node.js 16+（推荐 18+）
- 若要处理 `WPS` 内嵌字体，首次运行需要本机可用的 C 编译器：`cc` / `clang` / `gcc`
- 若直接输入 `.pptx`，脚本需要系统解压能力：
  - macOS / Linux：`unzip`
  - Windows：`PowerShell Expand-Archive`

---

## npm 安装

发布到 npm 后，可以这样安装和使用：

### 全局安装

```bash
npm install -g ppt-font-toolkit

ppt-font-toolkit --help
ppt-font-recover --help
ppt-font-metrics --help
```

### 直接用 `npx`

```bash
npx ppt-font-toolkit --help
npx ppt-font-recover ./demo.pptx
npx ppt-font-metrics --scan --json
```

### 本地开发联调

```bash
npm link

ppt-font-toolkit --help
ppt-font-recover --help
ppt-font-metrics --help
```

---

## 快速开始

```bash
git clone https://github.com/zyizyiz/PPT-Font-Toolkit.git
cd PPT-Font-Toolkit

# 查看总入口帮助
node ppt-font-toolkit.mjs --help

# 查看恢复工具帮助
node ppt-font-recover.mjs --help

# 查看 metrics 工具帮助
node font-metrics.mjs --help
```

---

## 后端集成

### Node.js 后端：直接当库调用

从 `1.2.0` 起（代码已支持，发布新版本后即可直接安装使用），包同时提供 CLI 和程序化 API。

```js
import {
  buildMetricsMap,
  extractFontMetrics,
  listEmbeddedFonts,
  recoverFonts,
} from 'ppt-font-toolkit'

try {
  const { reports } = listEmbeddedFonts(
    { sources: ['./demo.pptx'] },
    {
      onReport(report) {
        console.log('list report:', report.sourceLabel)
      },
      onComplete(summary) {
        console.log('list done:', summary.reports.length)
      },
    }
  )

  const { results } = recoverFonts(
    { sources: ['./demo.pptx'], outputDir: './recovered-fonts' },
    {
      onRecovered(result) {
        console.log('recovered:', result.output)
      },
      onComplete(summary) {
        console.log('recover done:', summary.results.length)
      },
    }
  )

  const metrics = extractFontMetrics({
    inputs: ['/Library/Fonts/Arial.ttf'],
  })
  const metricsMap = buildMetricsMap(metrics)

  console.log(reports, results, metricsMap)
} catch (error) {
  console.error('failed:', error.message)
}
```

约定如下：

- 成功：函数直接返回结构化结果
- 失败：函数抛出 `Error`
- 回调：可选 `onEvent / onComplete`
- `recover` 额外支持：`onRecovered / onReport / onTempDir`
- `metrics` 额外支持：`onScanStart / onScanComplete / onSave`

也可以按子路径导入：

```js
import { recoverEmbeddedFonts } from 'ppt-font-toolkit/recover'
import { collectMetrics } from 'ppt-font-toolkit/metrics'
```

### Java / 其他后端：调用 CLI

Java 后端不直接调用 Node SDK，推荐继续起子进程调用 CLI，并统一使用 `--json`：

```bash
ppt-font-recover ./demo.pptx --json
ppt-font-recover ./demo.pptx --list --json
ppt-font-metrics /Library/Fonts/Arial.ttf --json
```

约定如下：

- 成功：退出码 `0`
- 失败：退出码非 `0`
- 结果：`stdout` 输出 JSON
- 错误：`stderr` 输出错误信息

如果你是 Java `ProcessBuilder` / Spring Boot 集成，这一层自己监听进程结束即可把它当“完成回调”。

---

## 一键恢复 PPT 内嵌字体

### 推荐用法

你现在可以直接把 `.pptx`、解压目录、单个 `.fntdata` 或一个大目录丢给脚本，它会自动判断是 `Office` 还是 `WPS`。

```bash
# 1) 直接处理一个 .pptx
node ppt-font-toolkit.mjs ./demo.pptx

# 2) 直接处理一个已解压的 PPT 目录
node ppt-font-toolkit.mjs ./demo-unzipped

# 3) 直接处理单个 .fntdata
node ppt-font-toolkit.mjs ./ppt/fonts/font1.fntdata

# 4) 批量扫描一个目录
node ppt-font-toolkit.mjs ./downloads --output-dir ./recovered-fonts
```

### 恢复命令

```bash
node ppt-font-recover.mjs <source>
node ppt-font-recover.mjs <source1> <source2> ...
node ppt-font-recover.mjs --input <source>
node ppt-font-recover.mjs --ppt-dir <path>
```

### `source` 支持的输入类型

- `.pptx / .pptm / .ppsx / .ppsm / .potx / .potm`
- 已解压的 PPT 根目录
- 已解压目录中的 `ppt` 目录
- 单个 `.fntdata`
- 包含上述内容的目录（会自动递归扫描）

### 常用示例

```bash
# 1) 直接恢复一个 PPT 里的全部嵌入字体
node ppt-font-recover.mjs ./demo.pptx

# 2) 只列出字体，不恢复
node ppt-font-recover.mjs ./demo.pptx --list

# 3) 指定输出目录
node ppt-font-recover.mjs ./demo.pptx --output-dir ./demo-fonts

# 4) 只恢复某几个字体
node ppt-font-recover.mjs ./demo-unzipped --font "rId10,汉仪汉黑简,font1.fntdata"

# 5) 批量处理多个 PPT
node ppt-font-recover.mjs ./a.pptx ./b.pptx --output-dir ./all-fonts

# 6) 单个 Office .fntdata 直解（手动给 key）
node ppt-font-recover.mjs ./font1.fntdata --key "{A1B2C3D4-E5F6-1234-ABCD-1234567890AB}"

# 7) 单个 .fntdata，但它就在已解压 PPT 的 fonts 目录里
# 脚本会自动向上查找 presentation.xml 并尝试补全 fontKey
node ppt-font-recover.mjs ./ppt/fonts/font1.fntdata

# 8) 批量扫描整个目录
node ppt-font-recover.mjs ./downloads --output-dir ./recovered-fonts
```

### 输出规则

- 输入是单个 `.pptx / 解压目录`：
  - 默认输出到同级目录：`<源名>-recovered-fonts/`
- 输入是单个 `.fntdata`：
  - 默认输出到原文件所在目录
- 传 `--output-dir`：
  - 多任务时会按源名自动分子目录
- 传 `--output`：
  - 只能用于单个恢复任务

### 常用参数

- `--input <path>`：添加输入源，可重复
- `--ppt-dir <path>`：兼容旧参数，等价于添加一个解压目录输入
- `--font <value>`：按 `rId` / 字体名 / 文件名筛选，可重复或逗号分隔
- `--key <guid>`：单个 Office `.fntdata` 直解时手动指定 `fontKey`
- `--output <path>`：单任务输出文件
- `--output-dir <path>`：批量输出目录
- `--list`：仅列出内嵌字体
- `--json`：JSON 输出结果
- `--keep-temp`：保留 `.pptx` 自动解压的临时目录

### `--list` 示例

```bash
node ppt-font-recover.mjs ./demo.pptx --list
```

输出会包含：

- `Typeface`
- `Style`
- `rId`
- `Target`
- `Method`
- `fontKey`

其中 `Method` 会自动标出：

- `office-odttf`
- `office-odttf (need key)`
- `wps-eot`

### 关于 WPS

如果脚本识别到 `WPS / Kingsoft` 的压缩 `EOT / MTX`：

- 不需要 `fontKey`
- 首次运行会自动编译本地 helper
- helper 源码已随仓库 vendored 在 `vendor/libeot`

---

## 字体 Metrics 提取

### 功能

从字体文件中提取：

- `usWinAscent`
- `usWinDescent`
- `unitsPerEm`

并计算：

`lineRatio = (usWinAscent + usWinDescent) / unitsPerEm`

这个比率可用于估算 PPT / 排版中的单倍行距基线。

### 支持格式

- `.ttf`
- `.otf`
- `.ttc`
- `.woff`

### 用法

```bash
node ppt-font-toolkit.mjs metrics <字体文件路径>
node ppt-font-toolkit.mjs metrics --dir <目录路径>
node ppt-font-toolkit.mjs metrics --scan
node font-metrics.mjs <字体文件路径>
node font-metrics.mjs --dir <目录路径>   # 兼容旧脚本名
node font-metrics.mjs --scan             # 兼容旧脚本名
node font-metrics.mjs --scan --filter <正则关键字>
node font-metrics.mjs --scan --json
node font-metrics.mjs --scan --json --map
node font-metrics.mjs --scan --json --save <文件路径>
node font-metrics.mjs --scan --code
```

### 常用示例

```bash
# 1) 提取单个字体
node ppt-font-toolkit.mjs metrics /Library/Fonts/Arial.ttf

# 2) 扫描系统字体并筛选
node font-metrics.mjs --scan --filter "微软雅黑|Arial|Times|宋体|Courier"

# 3) 扫描指定目录
node font-metrics.mjs --dir ./fonts

# 4) 输出 JSON
node font-metrics.mjs --scan --json

# 5) 输出 familyName => metrics 的 Map
node font-metrics.mjs --scan --json --map

# 6) 保存到文件
node font-metrics.mjs --scan --json --save fonts.json

# 7) 输出可直接粘贴的代码
node font-metrics.mjs --scan --code
```

### 输出字段说明

- `familyName`
- `familyNameEn`
- `usWinAscent`
- `usWinDescent`
- `unitsPerEm`
- `lineRatio`

### 公式参考

```text
PPT行高(px) = fontSize(pt) × lineRatio × spcPct × 96/72
```

---

## 总入口与命令别名

以下命令别名在执行 `npm link` 或全局安装后可直接使用；在仓库目录内直接运行时，请继续使用 `node *.mjs`。

### 新命令

- `ppt-font-toolkit`：总入口
- `ppt-font-recover`：恢复内嵌字体
- `ppt-font-metrics`：提取字体 metrics（推荐）

### 兼容旧命令

- `font-metrics`：旧别名，仍可用
- `font-odttf-decrypt`

也就是说，这几种方式都可以：

```bash
node ppt-font-toolkit.mjs ./demo.pptx
node ppt-font-toolkit.mjs recover ./demo.pptx
node ppt-font-recover.mjs ./demo.pptx
node odttf-decrypt.mjs ./demo.pptx
```

---

## Vendored 代码说明

仓库内包含 `vendor/libeot`，用于处理 `WPS / EOT / MTX` 恢复流程。

- 上游许可证见：`vendor/libeot/LICENSE`
- 相关专利声明见：`vendor/libeot/PATENTS`

---

## License

项目主许可证：`Apache-2.0`（见根目录 `LICENSE`）

其中 `vendor/libeot` 保持其上游许可证与声明文件。

---

## English

`PPT-Font-Toolkit` is a Chinese-first Node.js CLI for two tasks:

- recovering embedded fonts from `Microsoft Office` and `WPS / Kingsoft Presentation`
- extracting font metrics from normal font files

### Features

- Auto-detect `Office ODTTF` and `WPS / Kingsoft EOT / MTX`
- Accept `.pptx`, unpacked PPT directories, single `.fntdata`, or recursive folders
- Support `--list` to inspect embedded fonts before recovery
- Keep the original font metrics extractor in the same toolkit

### Requirements

- Node.js 16+ (`18+` recommended)
- For `WPS` embedded fonts, the first run needs a local C compiler: `cc`, `clang`, or `gcc`
- For direct `.pptx` input, the system needs an unzip command:
  - macOS / Linux: `unzip`
  - Windows: `PowerShell Expand-Archive`

### Install from npm

After publishing to npm, you can install and run it like this:

```bash
npm install -g ppt-font-toolkit

ppt-font-toolkit --help
ppt-font-recover --help
ppt-font-metrics --help
```

Or use `npx` without a global install:

```bash
npx ppt-font-toolkit --help
npx ppt-font-recover ./demo.pptx
npx ppt-font-metrics --scan --json
```

### Quick start from source

```bash
git clone https://github.com/zyizyiz/PPT-Font-Toolkit.git
cd PPT-Font-Toolkit

node ppt-font-toolkit.mjs --help
node ppt-font-recover.mjs --help
node font-metrics.mjs --help
```

### Recover embedded fonts

```bash
node ppt-font-toolkit.mjs ./demo.pptx
node ppt-font-recover.mjs ./demo.pptx
node ppt-font-recover.mjs ./demo-unzipped --list
node ppt-font-recover.mjs ./downloads --output-dir ./recovered-fonts
```

It auto-detects:

- `Office ODTTF` → recover with `fontKey + XOR`
- `WPS EOT / MTX` → recover through the bundled helper source

### Extract metrics

```bash
node ppt-font-toolkit.mjs metrics /Library/Fonts/Arial.ttf
node font-metrics.mjs --scan --json
```

If you install the package with `npm link` or a global npm install, you can also use these bin names directly:

- `ppt-font-toolkit`
- `ppt-font-recover`
- `ppt-font-metrics`

### License

Main project license: `Apache-2.0` (see `LICENSE`).

`vendor/libeot` keeps its upstream license and patent notices.
