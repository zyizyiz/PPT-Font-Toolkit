# Font-Metrics

中文优先的字体 Metrics 提取工具（Node.js CLI）。

从字体文件中提取 `usWinAscent`、`usWinDescent`、`unitsPerEm`，并计算：

`lineRatio = (usWinAscent + usWinDescent) / unitsPerEm`

这个比率可用于估算 PPT / 排版中的单倍行距基线。

---

## 中文说明

### 功能特性

- 支持单文件提取（`.ttf` / `.otf` / `.ttc` / `.woff`）
- 支持扫描系统字体目录（macOS / Windows / Linux）
- 支持扫描任意目录
- 支持按字体名或路径正则过滤
- 支持三种输出格式：表格、JSON、可直接粘贴的代码片段
- 不依赖第三方字体解析库（直接读取字体二进制表）

### 运行环境

- Node.js 16+（推荐 18+）

### 快速开始

```bash
git clone https://github.com/zyizyiz/Font-Metrics.git
cd Font-Metrics
node font-metrics.mjs --help
```

### 命令用法

```bash
node font-metrics.mjs <字体文件路径>
node font-metrics.mjs --dir <目录路径>
node font-metrics.mjs --scan
node font-metrics.mjs --scan --filter <正则关键字>
node font-metrics.mjs --scan --json
node font-metrics.mjs --scan --code
```

### 常用示例

```bash
# 1) 提取单个字体
node font-metrics.mjs /Library/Fonts/Arial.ttf

# 2) 扫描系统字体并筛选（正则）
node font-metrics.mjs --scan --filter "微软雅黑|Arial|Times|宋体|Courier"

# 3) 扫描指定目录
node font-metrics.mjs --dir ./fonts

# 4) 输出 JSON（方便程序消费）
node font-metrics.mjs --scan --json

# 5) 输出可直接使用的映射代码
node font-metrics.mjs --scan --code
```

### 输出字段说明

- `familyName`：字体族名称
- `usWinAscent`：Windows Ascender
- `usWinDescent`：Windows Descender
- `unitsPerEm`：每 Em 单位数
- `lineRatio`：`(usWinAscent + usWinDescent) / unitsPerEm`

### 计算说明

- 本工具基于字体的 `OS/2` 与 `head` 表提取数据。
- `lineRatio` 用于给出字体“单倍行高”的近似比例，不同渲染引擎仍可能存在轻微差异。
- 在本项目输出代码注释中，也给出了一个常见换算参考：

```text
PPT行高(px) = fontSize(pt) × lineRatio × spcPct × 96/72
```

### 适用场景

- PPT 文字布局还原
- 设计稿到演示稿/导出图片的行高一致性对齐
- 批量建立字体行高比率表

### License

ISC

---

## English

`Font-Metrics` is a Node.js CLI tool to extract font metrics and compute:

`lineRatio = (usWinAscent + usWinDescent) / unitsPerEm`

This ratio is useful for estimating single-line height behavior in PPT/typesetting workflows.

### Features

- Parse `.ttf`, `.otf`, `.ttc`, `.woff`
- Scan system font folders (macOS / Windows / Linux)
- Scan custom directories
- Filter results by regex (`--filter`)
- Output as table, JSON, or ready-to-paste code mapping
- No third-party font parsing dependency

### Requirements

- Node.js 16+ (18+ recommended)

### Usage

```bash
node font-metrics.mjs <font-file-path>
node font-metrics.mjs --dir <directory>
node font-metrics.mjs --scan
node font-metrics.mjs --scan --filter <regex>
node font-metrics.mjs --scan --json
node font-metrics.mjs --scan --code
```

### Examples

```bash
node font-metrics.mjs /Library/Fonts/Arial.ttf
node font-metrics.mjs --scan --filter "Arial|Times|Courier"
node font-metrics.mjs --dir ./fonts
node font-metrics.mjs --scan --json
node font-metrics.mjs --scan --code
```

### Output Fields

- `familyName`
- `usWinAscent`
- `usWinDescent`
- `unitsPerEm`
- `lineRatio`

### Formula Reference

```text
PPT lineHeight(px) = fontSize(pt) × lineRatio × spcPct × 96/72
```

### License

ISC
