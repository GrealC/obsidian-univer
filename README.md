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

`univer-plus` 是一款 Obsidian 桌面端插件，用于在仓库中创建、打开和编辑电子表格与文档。插件同时支持 Univer 原生格式以及常用的 Excel、Word 文件，并为可能丢失 Office 高级内容的场景提供保护视图、备份和覆盖检查。

项目基于 [dream-num/obsidian-univer](https://github.com/dream-num/obsidian-univer) 继续开发。当前插件版本为 **1.0.7**，编辑器内核已从 Univer `0.2.14` 更新至 `0.25.1`。

> [!IMPORTANT]
> `univer-plus` 可以安全处理常见 `.xlsx` 和基础 `.docx`，但不能完整替代 Microsoft Excel 或 Word。检测到当前转换器无法无损写回的内容时，插件会进入只读保护视图，不会冒险覆盖原文件。

## 格式支持概览

| 格式 | 新建 | 打开 | 直接编辑 | 保存方式 |
| --- | :---: | :---: | :---: | --- |
| `.usheet` | 是 | 是 | 是 | 保存 Univer 表格快照，适合使用完整的 Univer 表格功能 |
| `.udoc` | 是 | 是 | 是 | 保存 Univer 文档快照 |
| `.xlsx` | 是 | 是 | 视内容而定 | 常用内容可本地写回；含不兼容部件时进入保护视图 |
| `.docx` | 是 | 是 | 视内容而定 | 基础文字文档可写回；复杂文档高保真预览并可创建简化编辑副本 |

插件不会接管 `.xls`、`.xlsm`、`.doc` 或 `.docm`。请先在 Office 或兼容软件中另存为 `.xlsx` 或 `.docx` 副本。

## 主要改进

### 从旧版项目替换了什么

| 项目 | 旧版实现 | univer-plus |
| --- | --- | --- |
| 插件标识 | `univer` | `univer-plus`，版本线从 `1.0.0` 重新开始 |
| Univer 内核 | `0.2.14` | `0.25.1` presets 与公开 API |
| Excel 转换 | Univer Pro 私有交换服务、WASM 和 worker | 本地 `ExcelJS` 与 `JSZip` |
| 依赖来源 | 包含私有 Verdaccio 源 | 公共 npm registry |
| Office 文件保存 | 直接覆盖 | 格式检查、备份、外部修改检测、串行保存和空文件保护 |
| DOCX | 依赖其他插件预览 | 文件列表显示、基础编辑、高保真保护预览和编辑副本 |
| 字体 | 固定列表 | 搜索当前系统安装字体，支持本地化名称 |
| 多语言 | 主要影响新建编辑器 | 设置、命令、右键菜单、状态和已打开编辑器即时同步 |

旧版专有交换插件、远程转换代码及相关 WASM/worker 已移除。Excel 和 Word 内容不会因为格式转换而上传到外部服务。

### 使用体验

- 在左侧功能区、命令面板或文件列表右键菜单中创建文件。
- `.xlsx` 和 `.docx` 直接显示在 Obsidian 文件列表中。
- Office 编辑器提供加载、未保存、保存中、已保存、保护视图和错误状态。
- XLSX 与 DOCX 共用可搜索的本机字体菜单。
- 支持 English、简体中文、繁体中文、Русский 和 Tiếng Việt。
- 主题跟随 Obsidian，并针对工具栏、状态按钮和文档预览优化了样式。

![使用演示](./assets/use.gif)

## 安装与升级

### 从 GitHub Release 安装

1. 前往 [Releases](https://github.com/GrealC/obsidian-univer/releases) 下载最新版的 `main.js`、`manifest.json` 和 `styles.css`。
2. 在 Obsidian 仓库中创建 `.obsidian/plugins/univer-plus/`。
3. 将三个文件放入该目录，文件名保持不变。
4. 完全退出并重新启动 Obsidian。
5. 在“设置 -> 第三方插件”中启用 `univer-plus`，并确认显示的版本与 Release 一致。

升级时同样需要同时替换三个文件。仅替换 `main.js` 容易造成清单版本、样式和代码不一致。

> [!WARNING]
> 不要同时保留旧版 `.obsidian/plugins/univer/` 和新版 `.obsidian/plugins/univer-plus/`。如果 Docxer 或其他插件也注册了 `.docx`，请关闭冲突插件的 DOCX 接管功能并重启 Obsidian。

### 从源码构建

推荐使用 Node.js 22 和 pnpm 11：

```sh
pnpm install --frozen-lockfile
pnpm typecheck
pnpm lint
pnpm test
pnpm build
```

生产构建位于 `dist/`：

```text
dist/main.js
dist/manifest.json
dist/styles.css
```

## 快速使用

### 新建文件

可以通过以下入口创建文件：

- 点击 Obsidian 左侧功能区的 Univer 图标。
- 在命令面板中执行“创建 Univer 表格”“创建 Univer 文档”“创建 Excel 工作簿”或“创建 Word 文档”。
- 右键文件夹，在该文件夹内新建 `.xlsx` 或 `.docx`。
- 右键已有文件，在它的父文件夹内新建 `.xlsx` 或 `.docx`。

新文件默认命名为 `Untitled.ext`；名称占用时依次使用 `Untitled-1.ext`、`Untitled-2.ext`。

### 打开与保存

- 点击文件列表中的 `.usheet`、`.udoc`、`.xlsx` 或 `.docx` 即可打开。
- 可编辑的 Office 文件在内容变化约 2 秒后自动保存。
- `Ctrl+S` 或 `Cmd+S` 可以立即保存。
- 顶部状态图标可查看保存状态；未保存或保存失败时也可点击重试。
- 检测到外部程序修改了当前文件时，插件会拒绝覆盖并要求重新打开。

## 表格能力

### Univer 原生表格 `.usheet`

`.usheet` 直接保存 Univer 工作簿快照，适合需要更多 Univer 原生功能、且不要求在 Excel 中打开的场景。

编辑器当前集成：

- 公式、常用单元格格式、边框、填充、对齐和合并单元格。
- 多工作表、行列操作、冻结首行或首列、网格线和缩放。
- 条件格式、数据验证、筛选、排序、查找替换和超链接。
- 可搜索的系统字体。

![Univer 表格编辑](./assets/sheet.gif)

### Excel 工作簿 `.xlsx`

XLSX 转换完全在本地进行。以下常用内容支持读取、编辑和写回：

- 文本、数字、日期、布尔值、公式和共享公式。
- 字体、字号、粗体、斜体、下划线、删除线、颜色和纯色填充。
- 边框、水平/垂直对齐、自动换行和文字旋转。
- 行高、列宽、隐藏行列、合并单元格和多工作表。
- 隐藏工作表、冻结窗格、网格线和从右到左视图。
- 超链接、基础结构化表格及其当前单元格值。
- 全局 A1 命名区域、打印区域和打印标题。

以下内容无法保证无损往返。检测到任一项时，工作簿会以保护视图打开，可选择、滚动、查找和复制，但不会写回原文件：

- 图表、数据透视表及缓存、绘图、图片、批注和嵌入对象。
- VBA、ActiveX、窗体控件、切片器、查询表、连接和外部链接。
- 条件格式、数据验证、自动筛选、工作表保护和手动分页符。
- 高级表格公式、汇总行和表格扩展信息。
- 工作表级、公式型或包含高级属性的命名区域。

`.usheet` 中可用的高级功能不一定能够写回 `.xlsx`。需要保留这些功能时，请优先使用 `.usheet`，或继续在桌面 Excel 中处理原工作簿。

![Excel 工作簿编辑](./assets/excel.gif)

## 文档能力

### Univer 原生文档 `.udoc`

`.udoc` 直接保存 Univer 文档快照，支持基础文字、段落、标题、列表和页面编辑。它不是 DOCX 文件，适合仅在 Obsidian 与 Univer 中使用的文档。

### Word 文档 `.docx`

基础 DOCX 可以直接编辑并保存回原文件。目前支持安全往返的主要内容包括：

- 普通段落、Title、Subtitle 和一至五级标题。
- 字体、字号、粗体、斜体、下划线、删除线、文字颜色和上下标。
- 左对齐、居中、右对齐、两端对齐、缩进和段落间距。
- 行距、制表符、分页符、页面尺寸和页边距。

检测到复杂 Word 内容时，插件使用 `docx-preview` 显示高保真保护视图，保留原文件，不直接写回。触发保护视图的内容包括：

- 表格、图片、图表、SmartArt、绘图、文本框和嵌入对象。
- 编号或项目符号列表、超链接、书签、域和内容控件。
- 批注、修订、脚注、尾注、页眉和页脚。
- 自定义样式、字符样式、高级文字效果和复杂段落格式。
- 多节布局、高级分栏/边框、文档保护和宏。

保护视图顶部提供“编辑副本”。该操作会在原文件旁生成 `<原文件名>-editable.docx`，移除当前转换器无法保留的复杂对象，然后在 Univer 中打开副本。原 DOCX 不会被修改。

![Word 文档编辑与预览](./assets/doc.gif)

## 本机字体搜索

字体列表不是写死的。插件优先调用当前 Electron 环境提供的本机字体 API；Windows 上还会读取系统字体集合，合并、去重后注册到 XLSX、DOCX、USheet 和 UDoc 编辑器。

- 最多读取 2048 个字体家族。
- 可按字体 family 名称搜索，例如 `Microsoft YaHei`、`KaiTi`。
- 可按系统返回的本地化名称搜索，例如“微软雅黑”“楷体”。
- 对常见中文字体提供额外的中英文名称映射。
- 搜索结果先过滤再限制显示数量，因此未显示在初始列表中的字体仍可被搜索到。

刚安装的字体不会自动注入已经运行的 Electron 进程，请完全重启 Obsidian。某些字体仅为应用私有字体、可变字体实例或未向系统注册时，仍可能无法被发现或无法在 Canvas 中正确渲染。

## 多语言

支持以下界面语言：

- English
- 简体中文
- 繁體中文
- Русский
- Tiếng Việt

修改语言后，设置页、命令名称、左侧功能区提示、右键菜单、Office 状态、保护视图以及已经打开的 Univer 编辑器会立即更新。

XLSX/DOCX 文件处理器开关属于 Obsidian 扩展名注册，修改这两个开关后仍需重启 Obsidian。

![语言设置](./assets/language.gif)

## 文件安全与隐私

### 保存保护

- **能力检查**：打开 Office 文件前检查 OOXML 部件和标记，无法安全写回时自动只读。
- **首次保存备份**：每次 Excel 或 Word 编辑会话第一次写回前备份原文件，可在设置中关闭。
- **备份清理**：每个源文件最多保留最近 3 份备份。
- **外部修改检测**：保存前再次读取文件，避免覆盖其他程序刚写入的版本。
- **串行保存**：自动保存和手动保存进入同一队列，避免并发覆盖。
- **空内容保护**：原本有内容的文件意外变空时拒绝自动保存；确需清空时，连续执行两次 `Ctrl/Cmd+S` 确认。
- **导出验证**：写入前验证生成的 XLSX/DOCX 是否包含必要的包结构和正文/工作表。

备份默认位于：

```text
<vault>/.obsidian/plugins/univer-plus/backups/
```

即使启用了插件备份，也建议为重要 Office 文件保留独立的版本历史或仓库级备份。

### 本地处理

- XLSX 由 `ExcelJS` 与 `JSZip` 在本地解析和生成。
- DOCX 由本地 OOXML 转换器解析和生成，复杂预览使用 `docx-preview`。
- 插件不需要远程文档转换服务，不会为格式转换上传文件内容。

## 设置说明

| 设置 | 作用 | 是否需要重启 |
| --- | --- | :---: |
| 界面语言 | 切换 Univer 与插件界面语言 | 否 |
| 打开 Excel 工作簿 | 注册或停用 `.xlsx` 编辑器 | 是 |
| 打开 Word 文档 | 注册或停用 `.docx` 编辑器 | 是 |
| 备份 Office 文件 | 控制每次编辑会话首次保存前是否备份 | 否 |

## 常见问题

### 替换了 `dist` 文件，但功能还是旧的

1. 确认目录是 `.obsidian/plugins/univer-plus/`，不是旧版 `univer/`。
2. 同时替换 `main.js`、`manifest.json` 和 `styles.css`。
3. 完全退出 Obsidian，而不是只关闭当前窗口或重新加载页面。
4. 重启后在第三方插件列表中核对版本号。

### DOCX 没有由 univer-plus 打开

同一扩展名只能由一个插件接管。关闭 Docxer 或其他 DOCX 插件的文件处理功能，然后重启 Obsidian。也请确认“打开 Word 文档”设置已启用。

### 字体搜索不到

- 同时尝试中文显示名和英文字体 family 名称。
- 确认字体已安装到系统，而不是仅供某个应用临时使用。
- 完全重启 Obsidian，让插件重新枚举字体。
- 字体即使被发现，也可能因 Electron/Canvas 不支持该字体格式而显示异常。

### 文件为什么只能查看，不能编辑

这是保护视图，不是权限错误。顶部状态或提示会列出不能安全保留的内容。XLSX 应继续使用 Excel 处理；DOCX 可以保留原文件查看，或使用“编辑副本”创建简化版。

### 保存提示文件已在外部修改

插件检测到磁盘内容与打开时不同，因此停止覆盖。关闭当前标签页并重新打开文件，确认外部修改后再继续编辑。

## 开发与验证

```sh
pnpm install --frozen-lockfile
pnpm typecheck
pnpm lint
pnpm test
pnpm build
```

CI 使用 Node.js 22，依次执行类型检查、ESLint、Vitest 和生产构建。发布 GitHub Release 后，工作流会重新验证项目并上传 Obsidian 所需的三个构建文件。

主要技术组件：

- Obsidian Plugin API
- Univer `0.25.1`
- React 18
- ExcelJS、JSZip
- docx-preview
- TypeScript、Vite、Vitest

## 已知限制

- 仅支持 Obsidian 桌面端，最低 Obsidian 版本为 `1.5.11`。
- XLSX/DOCX 转换覆盖常用内容，不等同于完整的 Excel/Word 文件模型。
- 保护视图中的 Office 文件不会写回原文件。
- 复杂 DOCX 的“编辑副本”会移除当前转换器不能表达的对象和版式。
- `.usheet` 和 `.udoc` 是 Univer 原生快照，不是标准 Office 格式。
- 自动保存不能替代正式的版本控制和异地备份。

完整版本记录见 [CHANGELOG.md](./CHANGELOG.md)。问题反馈请使用仓库的 [Issues](https://github.com/GrealC/obsidian-univer/issues)。

## 项目来源与许可

本项目由 [dream-num/obsidian-univer](https://github.com/dream-num/obsidian-univer) 衍生，编辑能力由 [Univer](https://github.com/dream-num/univer) 提供。感谢原项目维护者和 Univer 社区。

项目依据 [Apache License 2.0](./LICENSE) 开源。第三方依赖遵循各自许可证。
