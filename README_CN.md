# cli-anything-wps

<h3 align="center">让 AI Agent 操控 WPS Office 的命令行工具</h3>

<p align="center">
  <img src="https://img.shields.io/badge/平台-Windows-blue?logo=windows" alt="Platform">
  <img src="https://img.shields.io/badge/Python-3.10+-blue?logo=python" alt="Python">
  <img src="https://img.shields.io/badge/协议-MIT-green" alt="License">
  <img src="https://img.shields.io/badge/WPS-Office%2012.0+-red" alt="WPS">
</p>

---

## 项目简介

**cli-anything-wps** 是一个基于 [CLI-Anything](https://github.com/HKUDS/CLI-Anything) 架构构建的命令行工具。它将 WPS Office 的 COM 自动化接口封装为 47 个 CLI 命令，让 AI Agent 能够像操控开源软件一样操控闭源的 WPS Office。

你可以用它在终端里创建文档、编辑内容、设置格式、导出 PDF——全程不需要打开 WPS 界面，或者让 WPS 在前台可见地执行每一步操作。

## 为什么做这个项目

CLI-Anything 是一个很棒的项目：一行命令就能让 AI 操控 GIMP、Blender、LibreOffice 等开源软件。但它有一个硬伤——只支持开源软件。

可是现实中，大量行业软件是闭源的：
- **WPS Office**——国内最主流的办公套件
- **Adobe 全家桶**——设计行业的事实标准
- 它们的闭源生态让 AI Agent 无法直接接入

我们的答案是：**只要软件有可编程接口，就能接入 AI 的指挥系统。**

WPS 在 Windows 上暴露了完整的 COM 自动化接口（与 Microsoft Office VBA 兼容）。我们按照 CLI-Anything 的 7 阶段 Harness 方法论，手搓了一个 WPS 控制器。

## 功能概览

### WPS 文字（Writer）

- 创建/打开/保存文档
- 添加段落、标题、列表、表格、图片、分页符
- 字体/字号/颜色/粗体/斜体/对齐
- 查找替换、页面设置、页眉页脚
- 导出：DOCX / PDF / TXT / HTML / RTF / ODT

### WPS 表格（Calc）

- 工作表管理（添加/删除/重命名）
- 单元格读写（单个 + 批量区域）
- 公式设置、单元格合并
- 排序、筛选、条件格式
- 导出：XLSX / PDF / CSV / HTML

### WPS 演示（Impress）

- 幻灯片管理（添加/删除/排序/复制）
- 内容编辑（标题、正文、元素）
- 形状绘制（矩形、圆角、圆形、箭头）
- 文本框任意位置/尺寸/颜色/字体
- 导出：PPTX / PDF

### 通用能力

- 所有命令支持 `--json` 输出，AI Agent 可直接解析
- 交互式 REPL 模式（无参数运行时进入）
- 会话持久化 + 撤销/重做（最多 50 步历史）
- 后台模式（`Visible=False`）用于批量自动化

## 系统要求

| 组件 | 要求 |
|------|------|
| 操作系统 | Windows 10 / 11 |
| WPS Office | 2019 及以上（家庭和学生版/专业版均可） |
| Python | 3.10 及以上 |
| pywin32 | `pip install pywin32` |

> COM 接口是 Windows 专有技术，不支持 macOS 或 Linux。

## 安装

### 方式一：pip 安装（推荐）

```bash
pip install git+https://github.com/yb2460/cli-anything-wps.git
```

### 方式二：克隆安装

```bash
git clone https://github.com/yb2460/cli-anything-wps.git
cd cli-anything-wps
pip install -e .
```

### 验证

```bash
cli-anything-wps --help
```

## 快速开始

### 创建并编辑文档

```bash
# 新建 Writer 文档
cli-anything-wps document new --type writer --name "年度报告" -o report.json

# 添加内容
cli-anything-wps --project report.json writer add-heading -t "前言" -l 1
cli-anything-wps --project report.json writer add-paragraph -t "这是AI自动生成的报告。"
cli-anything-wps --project report.json writer add-table -r 3 -c 3

# 导出
cli-anything-wps --project report.json export render report.docx -p docx
cli-anything-wps --project report.json export render report.pdf -p pdf
```

### 操作电子表格

```bash
cli-anything-wps document new --type calc --name "数据" -o data.json
cli-anything-wps --project data.json calc set-cell A1 "产品名"
cli-anything-wps --project data.json calc set-cell A2 "WPS CLI"
cli-anything-wps --project data.json calc set-range A3 -d '[["张三",28],["李四",35]]'
cli-anything-wps --project data.json export render data.xlsx -p xlsx
```

### 制作演示文稿

```bash
cli-anything-wps document new --type impress --name "演示" -o slides.json
cli-anything-wps --project slides.json impress add-slide -t "标题页" -c "正文内容"
cli-anything-wps --project slides.json impress add-element 0 --type text_box --text "Hello!"
cli-anything-wps --project slides.json export render slides.pptx -p pptx
```

### AI Agent 模式

```bash
# 所有命令加 --json 返回结构化数据
cli-anything-wps --json document new --type writer --name "test"
cli-anything-wps --json --project test.json session status
```

## 全部命令

```
cli-anything-wps
├── document new|open|save|info|profiles|json    文档管理
├── writer                                         文字处理
│   ├── add-paragraph|add-heading|add-list
│   ├── add-table|add-image|add-page-break
│   ├── remove|list|set-text|find-replace
├── calc                                           电子表格
│   ├── add-sheet|remove-sheet|rename-sheet
│   ├── set-cell|get-cell|set-range|merge-cells
│   └── list-sheets
├── impress                                        演示文稿
│   ├── add-slide|remove-slide|set-content
│   └── list-slides|add-element
├── style create|modify|list|apply|remove          样式管理
├── export presets|preset-info|render              导出渲染
├── session status|undo|redo|history               会话管理
└── repl                                           交互模式
```

## 工作原理

```
CLI 命令 (Click)
    │
    ▼
Session 层 —— 撤销/重做/持久化
    │
    ▼
Core 模块 —— writer.py / calc.py / impress.py / export.py
    │
    ▼
WPS Backend (wps_backend.py) —— COM 接口封装
    │
    ▼
COM 接口 —— KWPS / KET / KWPP.Application
    │
    ▼
WPS Office —— 执行实际操作
```

### 三个 COM 入口

| ProgID | 对应应用 |
|--------|---------|
| `KWPS.Application` | WPS 文字（类 Word） |
| `KET.Application` | WPS 表格（类 Excel） |
| `KWPP.Application` | WPS 演示（类 PPT） |

## 常见问题

**Q: 为什么需要安装 WPS 才能用？**
本工具通过 COM 接口操控真实的 WPS 程序，不是模拟文件格式。它的渲染效果和手动操作完全一致。

**Q: 支持 Microsoft Office 吗？**
WPS COM 接口与 MS Office VBA 高度兼容。如果你只有 Office，把 `wps_backend.py` 中的 ProgID 换成 `Word.Application`、`Excel.Application`、`PowerPoint.Application` 即可。

**Q: 能做批量自动化吗？**
可以。设置 `app.Visible = False` 即可后台静默运行，适合批量报告生成、格式互转等场景。

**Q: 如何贡献？**
Fork 本仓库，提交 PR。或把 `registry_entry.json` 提交到 CLI-Anything 官方市场。

## 许可证

MIT License

---

<p align="center">
  基于 <a href="https://github.com/HKUDS/CLI-Anything">CLI-Anything</a> Harness 方法论构建
</p>
