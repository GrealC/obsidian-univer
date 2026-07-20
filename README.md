<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="./assets/banner-light.png">
    <img src="./assets/banner-dark.png" alt="univer-plus" width="420">
  </picture>
</p>

<p align="center">
  <a href="https://github.com/GrealC/obsidian-univer/releases"><img src="https://img.shields.io/github/v/release/GrealC/obsidian-univer?display_name=tag" alt="GitHub release"></a>
  <a href="./LICENSE"><img src="https://img.shields.io/github/license/GrealC/obsidian-univer" alt="License"></a>
  <a href="https://github.com/GrealC/obsidian-univer/actions/workflows/test.yml"><img src="https://github.com/GrealC/obsidian-univer/actions/workflows/test.yml/badge.svg" alt="Test status"></a>
</p>

# univer-plus

`univer-plus` 是一款 Obsidian 桌面端插件，用于在仓库内创建、查看和编辑电子表格与文档。它基于 [dream-num/obsidian-univer](https://github.com/dream-num/obsidian-univer) 继续开发，并以 [Univer](https://github.com/dream-num/univer) 作为编辑器内核。

本分支重点解决旧版长期未更新后出现的兼容性、文件安全和交互问题。当前版本为 **1.0.6**，Univer 已由 **0.2.14** 升级至 **0.25.1**。

> [!IMPORTANT]
> `.xlsx` 和 `.docx` 的目标是安全处理常用 Office 文件，而不是宣称完整复刻 Microsoft Excel 或 Word。插件会在写回可能造成数据丢失时自动进入保护视图。

## 本分支更新了什么

### 核心升级

| 项目 | 原实现 | univer-plus |
| --- | --- | --- |
| 插件标识 | `univer` | `univer-plus` |
| 版本线 | 原项目版本 | 从 `1.0.0` 重新开始，当前 `1.0.6` |
| Univer 内核 | `0.2.14` | `0.25.1` |
| Excel 转换 | Univer Pro 私有交换服务及 WASM | 本地 `ExcelJS` + `JSZip` |
| 包来源 | 依赖私有 Verdaccio 源 | 全部使用公共 npm registry |
| Univer 接入方式 | 旧插件与内部接口组合 | 基于新版 presets 和公开 API |
| 文件保存 | 直接覆盖源文件 | 备份、变更检测、串行保存和空文件保护 |

旧版专有交换插件、远程转换依赖、相关 WASM 文件和 worker 已从项目中移除。`.xlsx` 内容不会为了格式转换而上传到外部服务。

### 表格编辑能力

- 公式、常用单元格格式、合并单元格和多工作表编辑。
- 条件格式、数据验证、筛选、排序、查找与替换、超链接。
- 插入行列、冻结首行或首列、取消冻结、显示或隐藏网格线。
- 本机字体发现与搜索，不再依赖写死的字体列表。
- 支持使用中文名称搜索常见中文字体，例如“微软雅黑”“楷体”“华文楷体”和“霞鹜文楷”。
- 表格和文档编辑器共用可搜索的本机字体菜单。
- 修复字体搜索框点击或输入时下拉层意外关闭的问题，并扩大本机字体发现容量。
- 为 `.xlsx` 编辑器增加加载、未保存、保存中、已保存、保护视图和错误状态反馈。

![表格编辑](./assets/sheet.gif)

### 文档编辑能力

`.udoc` 使用 Univer 文档快照保存，支持文本样式、段落、列表及基础页面编辑。文档编辑器与表格编辑器共享新版 Univer 运行时和多语言配置。

从 `1.0.3` 开始，插件可直接注册 `.docx`：Word 文件会显示在 Obsidian 文件列表中，点击后使用 Univer 预览；基础文档可以直接编辑并保存回原文件，不再只能通过 Docxer 预览。

- 本地读取和写回 DOCX，不上传文档内容。
- 支持普通段落、标题、字体、字号、粗体、斜体、下划线、删除线、颜色、上下标和段落对齐。
- 支持缩进、行距、段前段后间距、分页符、制表符、页面尺寸和页边距。
- 支持新建标准 `.docx` 文件。
- 复用 Office 文件备份、外部修改检测、串行保存及空文档保护。

含复杂 Word 资源的文档会以高保真保护视图打开，通过 `docx-preview` 保留分页、表格、图片、页眉和页脚，原文件不会被覆盖。需要修改文字时，可以点击“编辑副本”，在原文件旁创建并打开一个移除复杂对象的可编辑 `.docx` 副本。

![文档编辑](./assets/doc.gif)

## 文件格式与兼容性

### `.usheet`

`.usheet` 是插件的原生表格格式，直接保存 Univer 快照。需要使用条件格式、数据验证、筛选等 Univer 功能时，优先选择该格式，可以避免 Excel 格式转换造成的能力差异。

### `.udoc`

`.udoc` 是插件的原生文档格式，用于在 Obsidian 中创建和编辑 Univer 文档。

### `.docx`

本地 DOCX 转换器当前支持安全往返以下常用内容：

- 普通段落、Title、Subtitle 和一至五级标题。
- 字体、字号、粗体、斜体、下划线、删除线、文字颜色和上下标。
- 左对齐、居中、右对齐、两端对齐、缩进及段落间距。
- 制表符、分页符、行距、页面尺寸和页边距。

检测到以下内容时，插件会进入高保真保护视图并保留原文件：

- 图片、图表、SmartArt、文本框、嵌入对象和宏。
- 表格、编号或项目符号列表、超链接和书签。
- 批注、修订、域、内容控件、脚注、尾注、页眉和页脚。
- 自定义样式、复杂段落格式、多节文档或受保护文档。

旧版 `.doc` 和启用宏的 `.docm` 不会被插件注册。请先使用 Word 或兼容应用将副本转换为 `.docx`。

### `.xlsx`

本地转换器当前支持读写以下常用内容：

- 单元格值、公式和共享公式。
- 字体、字号、粗体、斜体、下划线、删除线、文字颜色和纯色填充。
- 边框、对齐、自动换行和文字旋转。
- 行高、列宽以及隐藏的行列。
- 合并单元格、隐藏工作表、冻结窗格、网格线和从右到左视图。
- 超链接、基础结构化表格及其当前值。
- 全局 A1 命名区域、打印区域和打印标题。

以下资源暂时不能安全地往返转换。检测到它们时，插件会以保护视图打开工作簿，允许选择、滚动、查找和复制，但不会修改源文件：

- 图表、数据透视表及缓存、绘图、图片、批注和嵌入对象。
- VBA 宏、ActiveX、窗体控件、切片器、连接和外部链接。
- 条件格式、数据验证、自动筛选和工作表保护。
- 高级表格公式、汇总行、扩展信息。
- 工作表级、公式型或其他高级命名区域。

插件不会注册旧版 `.xls` 或启用宏的 `.xlsm` 文件。请使用桌面 Excel 或兼容应用打开，或先将副本转换为 `.xlsx`。

## 文件安全

- **保护视图**：写回前检查不受支持的 Excel 和 Word 部件，避免静默删除图表、图片、宏等内容。
- **首次保存备份**：默认在每个 Excel 或 Word 编辑会话第一次保存前备份原文件，位置为 `<vault>/.obsidian/plugins/univer-plus/backups`。
- **自动清理备份**：每个源文件只保留最近 3 份备份，成功保存后自动删除更旧版本，避免长期累积。
- **外部修改检测**：文件在编辑期间被其他程序改动时，阻止无提示覆盖。
- **保存协调**：自动保存与手动保存串行执行，降低竞态覆盖风险。
- **空文件保护**：原本有内容的工作簿或文档意外变为空时拒绝自动保存；确需清空时，连续执行两次 `Ctrl/Cmd+S` 确认。

建议仍然对重要工作簿保留独立备份，并在含复杂 Excel 功能的文件上使用保护视图。

## 本机字体

字体菜单会读取当前桌面系统实际安装的字体，并同时匹配字体的显示名称和 CSS family 名称。Windows 环境下会在浏览器字体 API 不可用时回退到系统字体集合，因此中文字体无需预先写入插件列表。

如果刚安装的字体未出现，请完全重启 Obsidian；字体是否可用仍取决于操作系统及当前 Electron 运行环境。

## 安装

### 从 GitHub Release 安装

1. 打开本仓库的 [Releases](https://github.com/GrealC/obsidian-univer/releases) 页面并下载最新版本中的 `main.js`、`manifest.json` 和 `styles.css`。
2. 在 Obsidian 仓库内创建 `.obsidian/plugins/univer-plus/`。
3. 将三个文件放入该目录。
4. 重启 Obsidian，在“设置 -> 第三方插件”中启用 `univer-plus`。

请保持目录名为 `univer-plus`，不要将旧版 `univer` 与本插件混装在同一目录中。

如已安装 Docxer 等其他 DOCX 预览插件，请停用其 `.docx` 文件接管并重启 Obsidian，确保同一扩展名只由一个插件注册。

### 从源码构建

需要 Node.js 22 和 pnpm 11：

```sh
pnpm install --frozen-lockfile
pnpm build
```

构建产物位于 `dist/`。将其中的 `main.js`、`manifest.json` 和 `styles.css` 放入 `.obsidian/plugins/univer-plus/` 即可。

## 使用

1. 在 Obsidian 左侧功能区打开新建菜单。
2. 选择 Sheet、Doc、Excel 或 Word 创建对应文件。
3. 也可以直接在文件列表中打开已有 `.usheet`、`.udoc`、`.xlsx` 或 `.docx` 文件。
4. 右键文件夹可直接在该文件夹内新建 Excel 工作簿或 Word 文档；右键某个文件时会在其父文件夹内新建。
5. Office 文件包含不受支持的内容时，顶部状态会显示保护视图；点击状态可查看原因。

![使用演示](./assets/use.gif)

## 设置

- **界面语言**：English、简体中文、繁体中文、Русский、Tiếng Việt。修改后会立即刷新设置页、命令面板、右键菜单、文件状态以及已打开的 Univer 编辑器，无需重启。
- **Open `.xlsx` files with univer-plus**：启用或停用 `.xlsx` 文件处理器；修改后需重启 Obsidian。
- **Open Word documents**：启用或停用 `.docx` 文件处理器；修改后需重启 Obsidian。
- **Create backup before first save**：控制每个 Excel 或 Word 编辑会话第一次保存前是否创建备份，默认开启。

![语言设置](./assets/language.gif)

## 开发与验证

```sh
pnpm install --frozen-lockfile
pnpm typecheck
pnpm lint
pnpm test
pnpm build
```

持续集成在 Node.js 22 上执行类型检查、ESLint、Vitest 和生产构建。发布工作流会生成并上传 Obsidian 所需的三个文件：

```text
dist/main.js
dist/manifest.json
dist/styles.css
```

## 已知限制

- 仅支持 Obsidian 桌面端。
- `.xlsx` 转换覆盖常用场景，但不等同于 Excel 的完整文件模型。
- `.usheet` 中可用的部分高级功能不能安全写回 `.xlsx`。
- 图片、图表、数据透视表、宏和批注等内容目前只能通过保护视图保留源文件。
- `.udoc` 是 Univer 原生快照，不是 `.docx` 转换器。
- `.docx` 目前侧重基础文字文档；复杂 Word 文档可高保真预览，但编辑时需要创建简化副本。

## 致谢与许可

本项目由 [dream-num/obsidian-univer](https://github.com/dream-num/obsidian-univer) 衍生，编辑能力由 [Univer](https://github.com/dream-num/univer) 提供。感谢原项目维护者和 Univer 社区。

项目依据 [Apache License 2.0](./LICENSE) 开源。第三方依赖遵循各自的许可证。
