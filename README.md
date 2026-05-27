# cli-anything-wps

<h3 align="center">让 AI Agent 操控 WPS Office 的命令行工具</h3>

<p align="center">
  <img src="https://img.shields.io/badge/platform-Windows-blue?logo=windows" alt="Platform">
  <img src="https://img.shields.io/badge/python-3.10+-blue?logo=python" alt="Python">
  <img src="https://img.shields.io/badge/license-MIT-green" alt="License">
  <img src="https://img.shields.io/badge/WPS-Office%2012.0+-red" alt="WPS">
</p>

---

## 这是什么

**cli-anything-wps** 是一个基于 [CLI-Anything](https://github.com/HKUDS/CLI-Anything) 架构的命令行工具，让 AI Agent（如 Claude Code、ChatGPT）能够直接操控 WPS Office 完成文档创建、编辑、格式转换等操作。

WPS Office 虽然是闭源软件，但它在 Windows 上暴露了完整的 **COM 自动化接口**，与 Microsoft Office VBA 高度兼容。本项目将这套接口封装为 47 个 CLI 命令，覆盖 WPS 文字、WPS 表格、WPS 演示三大应用。

## 为什么需要它

CLI-Anything 官方只支持开源软件（通过源码分析生成 CLI）。但现实是，大量行业软件是闭源的：

- **WPS Office** 是中国最主流的办公套件之一
- **Adobe Photoshop** 是图像处理的事实标准
- 它们的闭环生态让 AI Agent 无法直接操控

本项目的答案是：**只要软件有可编程接口（COM/API/SDK），就能接入 AI Agent 的指挥系统。**

## 功能概览

### WPS 文字 (Writer) — `KWPS.Application`

| 功能 | 命令 |
|------|------|
| 创建/打开/保存文档 | `document new`, `open`, `save` |
| 段落、标题、列表 | `writer add-paragraph`, `add-heading`, `add-list` |
| 表格、图片、分页 | `writer add-table`, `add-image`, `add-page-break` |
| 查找替换 | `writer find-replace` |
| 字体/字号/颜色/样式 | `style create`, `modify`, `apply` |

### WPS 表格 (Calc) — `KET.Application`

| 功能 | 命令 |
|------|------|
| 工作表管理 | `calc add-sheet`, `remove-sheet`, `rename-sheet` |
| 单元格读写 | `calc set-cell`, `get-cell`, `set-range` |
| 公式、合并 | `calc set-cell --formula`, `merge-cells` |

### WPS 演示 (Impress) — `KWPP.Application`

| 功能 | 命令 |
|------|------|
| 幻灯片管理 | `impress add-slide`, `remove-slide`, `list-slides` |
| 内容编辑 | `impress set-content`, `add-element` |
| 排版布局 | 文本框任意位置/尺寸/颜色/字体 |

### 通用能力

| 功能 | 说明 |
|------|------|
| 格式导出 | DOCX / XLSX / PPTX / PDF / TXT / HTML / RTF / CSV |
| 撤销/重做 | `session undo`, `redo`，最多 50 步历史 |
| JSON 输出 | 所有命令支持 `--json` 标志，AI 可直接解析 |
| 交互式 REPL | 无参数运行进入对话模式，支持 `help`、`quit` |
| 会话持久化 | 项目状态保存为 `.wps-cli.json`，支持断点续操作 |

## 安装

### 系统要求

- Windows 10 或 Windows 11
- WPS Office 2019 及以上版本（家庭和学生版或专业版均可）
- Python 3.10 或更高版本
- 已安装 pywin32

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

### 方式三：开发模式

```bash
git clone https://github.com/yb2460/cli-anything-wps.git
cd cli-anything-wps
pip install -e .[dev]
```

### 验证安装

```bash
cli-anything-wps --help
```

## 快速开始

### 创建一个文档

```bash
# 新建 Writer 文档
cli-anything-wps document new --type writer --name "年度报告" -o report.json

# 添加内容
cli-anything-wps --project report.json writer add-heading -t "前言" -l 1
cli-anything-wps --project report.json writer add-paragraph -t "这是自动生成的报告。"

# 导出 DOCX
cli-anything-wps --project report.json export render report.docx -p docx

# 导出 PDF
cli-anything-wps --project report.json export render report.pdf -p pdf
```

### 操作电子表格

```bash
cli-anything-wps document new --type calc --name "数据表" -o data.json
cli-anything-wps --project data.json calc set-cell A1 "产品名"
cli-anything-wps --project data.json calc set-cell A2 "WPS CLI"
cli-anything-wps --project data.json export render data.xlsx -p xlsx
```

### 制作演示文稿

```bash
cli-anything-wps document new --type impress --name "演示" -o slides.json
cli-anything-wps --project slides.json impress add-slide -t "标题页"
cli-anything-wps --project slides.json export render slides.pptx -p pptx
```

### AI Agent 使用（JSON 模式）

```bash
# 所有命令加 --json 即返回结构化数据
cli-anything-wps --json document new --type writer --name "test"
cli-anything-wps --json --project test.json session status
cli-anything-wps --json --project test.json writer list
```

### 交互式 REPL

```bash
cli-anything-wps
#> document new --type writer --name demo
#> writer add-paragraph --text "Hello from REPL!"
#> export render demo.pdf --preset pdf
#> quit
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
│   ├── set-cell|get-cell|set-range
│   ├── merge-cells|list-sheets
├── impress                                        演示文稿
│   ├── add-slide|remove-slide|set-content
│   ├── list-slides|add-element
├── style create|modify|list|apply|remove          样式管理
├── export presets|preset-info|render              导出渲染
├── session status|undo|redo|history               会话管理
└── repl                                           交互模式
```

## 架构

```
CLI 命令 (Click)
    │
    ▼
Session 层 (撤销/重做/持久化)
    │
    ▼
Core 模块 (writer.py / calc.py / impress.py / export.py)
    │
    ▼
WPS Backend (wps_backend.py)
    │
    ▼
COM 接口 (KWPS / KET / KWPP.Application)
    │
    ▼
WPS Office (执行实际操作)
```

### 目录结构

```
cli_anything/wps/
├── wps_cli.py          # CLI 入口 (Click + REPL)
├── core/
│   ├── document.py     # 文档 CRUD
│   ├── writer.py       # 文字处理
│   ├── calc.py         # 电子表格
│   ├── impress.py      # 演示文稿
│   ├── styles.py       # 样式管理
│   ├── export.py       # 导出管道
│   └── session.py      # 会话管理
├── utils/
│   ├── wps_backend.py  # COM 后端封装
│   └── repl_skin.py    # 终端 UI
├── tests/
│   └── test_core.py    # 58 个单元测试
└── skills/
    └── SKILL.md        # AI Agent 指令文档
```

## 运行测试

```bash
pip install -e .[dev]
python -m pytest cli_anything/wps/tests/ -v
```

## 常见问题

**Q: 为什么需要 WPS Office 安装？**
A: 本工具通过 COM 接口操控真实的 WPS 程序，所以 WPS 必须安装在本地。它不像 python-docx 那样模拟文件格式，而是使用 WPS 自己的渲染引擎。

**Q: 支持 macOS 或 Linux 吗？**
A: COM 接口是 Windows 专有技术，不支持 macOS/Linux。WPS 的 Mac 版本没有暴露 COM 接口。

**Q: 和 Microsoft Office 兼容吗？**
A: WPS 的 COM 接口与 Microsoft Office VBA 高度兼容，但部分细节不同。如果你只有 MS Office，可以修改 `wps_backend.py` 中的 ProgID（如 `Word.Application`）。

**Q: 如何提交到 CLI-Anything 官方市场？**
A: 将 `registry_entry.json` 提交到 [CLI-Anything](https://github.com/HKUDS/CLI-Anything) 的 `public_registry.json` 文件中即可。

## 许可证

MIT License — 详见 [LICENSE](LICENSE)

---

<p align="center">
  Built with <a href="https://github.com/HKUDS/CLI-Anything">CLI-Anything</a> Harness Methodology
</p>
