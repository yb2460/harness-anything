# cli-anything-wps

<h3 align="center">让 AI Agent 操控 WPS Office 的命令行工具</h3>

<p align="center">
  <img src="https://img.shields.io/badge/平台-Windows-blue?logo=windows" alt="Platform">
  <img src="https://img.shields.io/badge/Python-3.10+-blue?logo=python" alt="Python">
  <img src="https://img.shields.io/badge/协议-MIT-green" alt="License">
  <img src="https://img.shields.io/badge/WPS-Office%2012.0+-red" alt="WPS">
</p>

---

## 这是什么

一个命令行工具，将 WPS Office 的 COM 自动化接口封装为 47 个 CLI 命令，让 AI Agent 直接操控 WPS 文字、WPS 表格、WPS 演示。

## 安装

```bash
pip install git+https://github.com/yb2460/cli-anything-wps.git
```

**前提**: Windows，WPS Office 2019+，Python 3.10+，`pip install pywin32`

## 快速上手

```bash
# 创建文档
cli-anything-wps document new --type writer --name "报告" -o report.json
cli-anything-wps --project report.json writer add-heading -t "前言" -l 1
cli-anything-wps --project report.json export render report.pdf -p pdf
```

## 新增：PPT 设计风格系统

本项目整合了 4 个专业 PPT Skill 的设计精华，内置：

**4 套设计预设**（一键切换风格）

| 预设 | 风格 | 适用场景 |
|------|------|---------|
| `academic` | 学术答辩 | 会议报告、论文答辩、基金申请 |
| `consultant` | 咨询顾问 | 商业计划书、咨询报告 |
| `business` | 商务汇报 | 会议汇报、项目提案、教学课件 |
| `tech` | 科技极简 | 产品发布、AI 演示、数据报告 |

**14 种布局模板**（cover/toc/overview/timeline/grid_cards/quadrant/stats/three_col/pipeline/data_table/content_image/closing）

**4 种演讲类型**（conference/business/defense/school，每种带完整页面序列）

**5 维度质量审查**（visual/pedagogy/proofreading/parity/substance）

```bash
# 查看所有预设
cli-anything-wps preset list

# 应用学术风格
cli-anything-wps preset apply academic --talk-type conference
```

## 全部命令

```
cli-anything-wps
├── document new|open|save|info|profiles|json   文档管理
├── writer add-paragraph|add-heading|add-list
│         add-table|add-image|add-page-break
│         remove|list|set-text|find-replace      文字处理
├── calc add-sheet|remove-sheet|rename-sheet
│        set-cell|get-cell|set-range|merge-cells  电子表格
├── impress add-slide|remove-slide|set-content
│           list-slides|add-element              演示文稿
├── style create|modify|list|apply|remove         样式管理
├── preset list|info|apply                        设计预设 ★新增
├── export presets|preset-info|render             导出渲染
└── session status|undo|redo|history              会话管理
```

## 支持的格式导出

DOCX / XLSX / PPTX / PDF / TXT / HTML / RTF / CSV / ODT

## 系统要求

| 组件 | 要求 |
|------|------|
| 操作系统 | Windows 10/11 |
| WPS Office | 2019 及以上 |
| Python | 3.10+ |
| pywin32 | `pip install pywin32` |

## 运行测试

```bash
pip install -e .[dev]
python -m pytest cli_anything/wps/tests/ -v
```

## 原理

```
CLI 命令 → Session → Core 模块 → WPS COM 接口 → WPS Office
                           KWPS/KET/KWPP.Application
```

## 许可证

MIT
