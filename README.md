# cli-anything-wps

<h3 align="center">AI Agent 工具集 —— 办公 + 设计 + 学术 全栈操控</h3>

<p align="center">
  <img src="https://img.shields.io/badge/平台-Windows-blue?logo=windows" alt="Platform">
  <img src="https://img.shields.io/badge/Python-3.10+-blue?logo=python" alt="Python">
  <img src="https://img.shields.io/badge/协议-MIT-green" alt="License">
</p>

---

## 包含项目

### 1. cli-anything-wps — WPS / Microsoft Office 操控

47 个 CLI 命令，通过 COM 自动化接口操控 WPS Office 或 Microsoft Office（Word/Excel/PowerPoint）。

```bash
pip install git+https://github.com/yb2460/cli-anything-wps.git
```

- **Writer**：段落/标题/列表/表格/图片/查找替换/字体样式
- **Calc**：工作表管理/单元格读写/公式/合并/批量填充
- **Impress**：幻灯片增删改/文本框/形状/背景/导出
- **导出**：DOCX/XLSX/PPTX/PDF/TXT/HTML/CSV/RTF
- **PPT 设计系统**：4 套预设 + 14 种布局 + 5 维度质量审查

---

### 2. cli-anything-zotero — 学术研究智能体

文献管理 + 27 个学术 Skill 集成。写综述、写论文、审稿、做图表一站式。

```bash
cli-anything-zotero skills list                    # 列出所有学术 Skill
cli-anything-zotero skills pipeline original_article  # 论著推荐流程
cli-anything-zotero skills pipeline meta_analysis     # Meta分析流程
cli-anything-zotero skills journal "Nature"           # 期刊图表规范
```

**7 大 Skill 分类：**

| 分类 | Skill 数 | 能力 |
|------|---------|------|
| search | 3 | 快速检索 / 系统评价 / 深度文献搜索 |
| research | 4 | 创意生成 / 头脑风暴 / 假设 / 深度研究 |
| writing | 5 | 写论文 / IMRAD稿件 / 引用 / 大纲 / 修改 |
| review | 5 | 5人审稿 / 同行评审 / 七轮对抗 / 引用验证 |
| visualization | 4 | 幻灯片 / 示意图 / 海报 / 期刊图表 |
| analysis | 3 | 探索分析 / 统计 / 证据评估 |
| pipeline | 2 | 完整学术流水线 / 研究到论文 |

---

### 3. Illustrator Harness — Adobe Illustrator AI Agent

通过 Windows COM 自动化接口操控 **Adobe Illustrator**，让 AI Agent 直接创建和编辑矢量图形。

```
CLI 命令 → Click CLI → Core 模块 → COM Bridge → Illustrator.Application → Illustrator 引擎
```

```bash
cd illustrator-harness/agent-harness
pip install -e .
```

| 命令组 | 功能 |
|--------|------|
| `project` | 新建/打开/保存 AI 文档 |
| `layers` | 图层增删改、可见性、锁定 |
| `shapes` | 矩形、椭圆、线条、多边形绘制 |
| `text` | 文字添加/修改（字体、大小、颜色） |
| `export` | 导出 PNG / JPEG / SVG / PDF / AI |

```bash
# 快速上手
cli-anything-illustrator project new logo.ai -w 500 -h 500
cli-anything-illustrator text add "Brand" --x 100 --y 100 --font "Arial" --size 72
cli-anything-illustrator shapes rect --x 50 --y 50 --w 200 --h 100
cli-anything-illustrator export svg output.svg
```

**前置条件**：Windows 10/11 + Adobe Illustrator 2023+ + Python 3.10+ + pywin32 + click

---

### 5. WPS PPT 自动化生成 — 5校招生数据全景

**JSON 数据驱动 + 元素类型路由 + WPS COM 即时生成**，全自动制作高质量 PPT。

| 项目 | 主题色 | 页数 | 元素 |
|------|--------|------|------|
| 华中科技大学 | `#004098` 蓝 | 9页 | 4图表+5校园图+3表格 |
| 南方科技大学 | `#006B3F` 绿 | 10页 | 4图表+5表格 |
| 北京大学 | `#8B0012` 红 | 14页 | 3图表+5表格 |
| 清华大学 | `#660874` 紫 | 13页 | 3图表+4表格 |
| 浙大城市学院 | `#005A9C` 蓝 | 10页 | 3图表+4表格 |

```bash
# 每个项目含三件套：模板底图 + JSON数据 + Python引擎
cd WPS/南科大
python build_sustech.py   # WPS可见生成 → PPTX + PDF
```

**13 种元素类型**：`text` / `image` / `table` / `cards_2x3` / `cards_1x4_info` / `card_list_wide` / `tagline_bar` / `timeline_horiz` 等

**布局规范**：透明卡片不挡背景 · 标题居中无装饰线 · 正文≥22pt · 每页图表+表格≥2种

> 📖 完整文档见 [WPS/README.md](WPS/README.md)

---

### 4. Photoshop Harness — Adobe Photoshop AI Agent

通过 Windows COM 自动化接口操控 **Adobe Photoshop**，让 AI Agent 直接创建和编辑位图图像。

```
CLI 命令 → Click CLI → Core 模块 → COM Bridge → Photoshop.Application → Photoshop 引擎
```

```bash
cd photoshop-harness/agent-harness
pip install -e .
```

| 命令组 | 功能 |
|--------|------|
| `project` | 新建/打开/保存 PSD 文档 |
| `document` | 文档属性（尺寸、分辨率、色彩模式） |
| `layer` | 图层增删改、可见性、透明度、混合模式 |
| `selection` | 选区操作（全选/羽化/反选/扩展） |
| `image` | 图像调整（裁切、旋转、翻转、画布大小） |
| `text` | 文字图层（字体、大小、颜色） |
| `export` | 导出 PNG / JPEG / WebP / PSD |
| `filter` | 滤镜操作 |

```bash
# 快速上手
cli-anything-photoshop project new poster.psd -w 1920 -h 1080
cli-anything-photoshop text add --content "Hello World" --font "Arial" --size 72
cli-anything-photoshop export png --output result.png
```

**前置条件**：Windows 10/11 + Adobe Photoshop 2023+ + Python 3.10+ + pywin32

---

## 快速上手

```bash
# WPS 办公
cli-anything-wps document new --type impress --name "演示"
cli-anything-wps preset apply academic --talk-type defense
cli-anything-wps export render output.pptx -p pptx

# Zotero 学术
cli-anything-zotero skills pipeline thesis
cli-anything-zotero catalog search "machine learning"

# Illustrator 设计
cli-anything-illustrator project new logo.ai -w 500 -h 500
cli-anything-illustrator shapes rect --w 200 --h 200

# Photoshop 设计
cli-anything-photoshop project new banner.psd -w 1920 -h 1080
cli-anything-photoshop export png banner.png
```

## 系统要求

- Windows 10/11
- WPS Office 2019+ 或 Microsoft Office 2016+ / Zotero 7+
- Adobe Illustrator 2023+（可选）/ Adobe Photoshop 2023+（可选）
- Python 3.10+ + pywin32

> COM 接口与 Microsoft Office VBA 兼容。如需操控 MS Office，将 ProgID 改为 `PowerPoint.Application` / `Word.Application` / `Excel.Application` 即可。

## 许可证

MIT
