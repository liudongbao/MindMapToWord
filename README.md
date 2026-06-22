# MindMapToWord —— 思维导图转 Word 桌面程序

一款 Windows 桌面应用（WinForms + .NET 8），可将主流思维导图格式一键转换为排版良好的 `.docx` 文件，**文字、图片、备注、超链接** 等内容均完整保留，并保持与导图一致的目录结构和节点顺序。

## ✨ 功能特色

- **多格式支持**：XMind（`.xmind`）、MindMaster / 亿图脑图（`.emmx`）、FreeMind（`.mm`、`.mmap`）、OPML（`.opml`）、Markdown（`.md`）。
- **Emmx 深度解析**：针对 MindMaster（`.emmx`）特有的 `page/page.xml` 结构，通过 `LevelData/Super V` 建立父子关系、`LevelData/SubLevel V` 按视觉顺序排序子节点，过滤连接线（MMConnector）。
- **完整保留内容**：节点标题、备注/注释、超链接、节点内嵌图片（支持 PNG / JPEG / GIF / BMP / TIFF）。
- **Word 结构化输出**：按导图层级自动映射到 Word 标题 1~4 + 正文，图片居中且自动等比缩放，**目录结构与导图一致**。
- **目录美观度优化**：**最后两层节点 + 大段文字（>50 字符）** 自动渲染为普通正文段落，不进入 Word 目录，提升文档美观度。
- **离线运行**：Word 生成基于 OpenXML SDK，**无需安装 Microsoft Office**。
- **树结构诊断**：内置节点层级诊断窗口，可查看节点 ID、坐标、排序、原始属性，便于排查解析问题。
- **批量导出**：支持一次选择多个思维导图文件，批量转换为 Word，实时显示进度条，单个文件失败不影响其余文件。
- **友好 UI**：文件选择 → 树状预览（显示节点数、图片数、链接等统计）→ 一键导出。

## 🚀 快速开始

### 环境要求

- Windows 7 / 10 / 11
- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) 或 [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0)（运行已发布的可执行文件）

### 构建与运行

```powershell
# 进入解决方案目录
cd d:\trae\demo2

# 还原依赖（首次运行）
dotnet restore

# 构建
dotnet build -c Release

# 直接运行
dotnet run -c Release --project MindMapToWord
```

或发布为单文件可执行程序：

```powershell
dotnet publish MindMapToWord -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true
# 输出路径: MindMapToWord\bin\Release\net8.0-windows\win-x64\publish\MindMapToWord.exe
```

### 使用步骤

**单文件导出**：
1. 启动程序后，点击 **「选择导图文件…」** 选择本地的 `.xmind` / `.emmx` / `.mm` / `.opml` / `.md` 文件。
2. 点击 **「解析预览」** 在左侧树状视图中查看导图结构（节点数、图片数会显示在状态栏）。
3. 勾选「包含备注」「包含超链接」选项以决定是否导出相应内容。
4. 点击 **「导出为 Word」** 选择保存位置，生成 `.docx`。生成后可立即打开查看。

**批量导出**：
1. 启动程序后，点击 **「选择导图文件…」**，按住 `Ctrl` 或 `Shift` 键选择多个文件。
2. 点击 **「批量导出…」** 按钮，选择输出目录。
3. 进度窗口实时显示当前处理的文件名和进度条。
4. 批量转换完成后，弹出汇总结果窗口，显示成功/失败数量及详细列表。

## 📁 项目结构

```
MindMapToWord/
├── Program.cs                  // 程序入口
├── MainForm.cs                 // 主窗口 UI（包含树结构诊断对话框、批量导出）
├── BatchProgressForm.cs        // 批量导出进度对话框
├── MindMapToWord.csproj        // 项目文件（.NET 8 - Windows）
├── Core/
│   ├── IMindMapParser.cs       // 解析器接口
│   ├── MindMapModel.cs         // 统一的导图/节点/图片模型（含调试属性）
│   └── MindMapParserFactory.cs // 按扩展名自动选择解析器
├── Parsers/
│   ├── XMindParser.cs          // .xmind（ZIP + content.json/xml）
│   ├── EmmxParser.cs           // .emmx 亿图脑图 / MindMaster（支持 LevelData 父子关系与 SubLevel 排序）
│   ├── FreeMindParser.cs       // .mm / .mmap（XML）
│   ├── OpmlParser.cs           // .opml
│   └── MarkdownParser.cs       // .md 标题/列表结构
└── Exporters/
    └── WordExporter.cs         // 使用 OpenXML SDK 生成 .docx（支持叶子节点 + 长文本的正文段落优化）
```

## 🔧 技术栈

| 组件 | 实现 |
|------|------|
| UI 框架 | WinForms + .NET 8 |
| Word 生成 | [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml) 2.20 |
| JSON 解析 | Newtonsoft.Json 13.0 |
| ZIP 解包 | `System.IO.Compression`（内置于 .NET） |
| 图片处理 | `System.Drawing`（仅用于读取尺寸信息） |
| 目标框架 | `net8.0-windows` |

## 🔄 版本更新记录

### v1.0.0 — 近期更新

- **框架升级**：目标框架从 `net6.0-windows` 升级至 `net8.0-windows`，修复编译兼容性问题。
- **Emmx 解析增强**：重写 `EmmxParser`，支持 MindMaster `page/page.xml` 结构。通过 `LevelData/Super V` 建立正确的父子关系，通过 `LevelData/SubLevel V` 按视觉顺序排序子节点，过滤 `MMConnector` 连接线。
- **目录结构修复**：确保 Word 目录层级与导图一致，节点顺序保持视觉顺序。
- **目录美观度优化**：深度 ≥ 3 层的叶子节点，若标题 > 50 字符，自动渲染为正文段落（不进入 Word 目录）。
- **树结构诊断**：新增节点层级诊断窗口，显示 `[层级]`、`children`、`order`、`X/Y` 坐标及原始属性，方便排查解析问题。
- **OpenXML 兼容性修复**：改用 `CoreFilePropertiesPart`/`ExtendedFilePropertiesPart` 替代过时 API；`EnumValue<OnOffOnlyValues>` 显式转换。
- **批量导出功能**：支持一次选择多个文件批量转换，使用 `BackgroundWorker` 在后台线程执行，实时进度条显示，输出目录可写性预验证，单个文件失败不影响其余文件，完成后显示汇总结果。

## 📝 输出效果示例

```
[居中大标题：文件名]

■ 中心主题（标题 1）
  ├ 链接：http://...
  ├ [图片，居中]
  │
  ├ ■ 子主题 A（标题 2）
  │   └ ■ 更细的节点（标题 3）
  │
  └ ■ 子主题 B（标题 2）
      └ 备注：这是一段备注文本
```

## ❓ 常见问题

**Q1: 导出的 Word 无法打开？**
A: 请确保使用的是 `.docx`（Office 2007+ 格式）。本程序仅生成 OpenXML 格式，不生成旧版 `.doc`。

**Q2: 图片没有显示？**
A: 某些导图格式的图片存储在外部本地路径。请确保原始导图及其附带的资源文件未被移动/删除；程序会尝试在文件同目录中查找图片。

**Q3: 支持中文文件名吗？**
A: 支持。所有路径、文件名、文本内容均使用 UTF-8 编码处理。

**Q4: 能在 Linux / macOS 运行吗？**
A: 本程序是 Windows 专用的 WinForms 应用。若需跨平台版本，可将 `Core` 和 `Exporters` 两个模块分离为控制台程序或 ASP.NET 服务使用。

## 📄 许可

MIT License
