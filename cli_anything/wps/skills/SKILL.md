---
name: "cli-anything-wps"
description: "WPS Office CLI — 通过命令行操控 WPS 文字/表格/演示文稿"
---

# cli-anything-wps

通过 COM 自动化接口操控 WPS Office 的命令行工具。支持 WPS 文字（Writer）、WPS 表格（Calc）和 WPS 演示（Impress）。

## 前置条件

- **Windows 操作系统**
- **WPS Office**（已安装，版本 12.0+）
- **Python 3.10+**
- **pywin32**：`pip install pywin32`

## 安装

```bash
pip install -e .
# 验证
cli-anything-wps --help
```

## 命令结构

```
wps
├── document          # 文档管理
│   ├── new           # 创建新文档
│   ├── open          # 打开项目文件
│   ├── save          # 保存文档
│   ├── info          # 文档信息
│   ├── profiles      # 页面配置列表
│   └── json          # 打印原始 JSON
├── writer            # WPS 文字
│   ├── add-paragraph
│   ├── add-heading
│   ├── add-list
│   ├── add-table
│   ├── add-page-break
│   ├── add-image
│   ├── remove
│   ├── list
│   ├── set-text
│   └── find-replace
├── calc              # WPS 表格
│   ├── add-sheet
│   ├── remove-sheet
│   ├── rename-sheet
│   ├── set-cell
│   ├── get-cell
│   ├── set-range
│   ├── merge-cells
│   └── list-sheets
├── impress           # WPS 演示
│   ├── add-slide
│   ├── remove-slide
│   ├── set-content
│   ├── list-slides
│   └── add-element
├── style             # 样式管理
│   ├── create
│   ├── modify
│   ├── list
│   ├── apply
│   └── remove
├── preset            # 设计预设（PPT风格一键切换）
│   ├── list           # 列出所有预设
│   ├── info [name]    # 查看预设详情
│   └── apply [name]   # 应用设计预设
├── export            # 导出
│   ├── presets       # 列出预设
│   ├── preset-info   # 预设详情
│   └── render        # 导出到文件
└── session           # 会话管理
    ├── status
    ├── undo
    ├── redo
    └── history
```

## 使用示例

### 创建并编辑文档

```bash
# 创建新文档
cli-anything-wps document new --type writer --name "报告" -o report.json

# 添加内容
cli-anything-wps --project report.json writer add-heading -t "年度报告" -l 1
cli-anything-wps --project report.json writer add-paragraph -t "这是报告正文内容。"
cli-anything-wps --project report.json writer add-table -r 3 -c 3

# 导出为 DOCX
cli-anything-wps --project report.json export render report.docx -p docx

# 导出为 PDF
cli-anything-wps --project report.json export render report.pdf -p pdf
```

### 使用电子表格

```bash
cli-anything-wps document new --type calc --name "数据" -o data.json
cli-anything-wps --project data.json calc set-cell A1 "姓名"
cli-anything-wps --project data.json calc set-cell B1 "年龄"
cli-anything-wps --project data.json calc set-range A2 -d '[["张三",28],["李四",35]]'
cli-anything-wps --project data.json export render data.xlsx -p xlsx
```

### JSON 输出（Agent 使用）

```bash
# 所有命令都支持 --json 标志用于机器解析
cli-anything-wps --json document new --type writer --name "test"
cli-anything-wps --json --project test.json session status
```

### REPL 模式

```bash
cli-anything-wps
#> document new --type writer --name "演示"
#> writer add-paragraph -t "你好 WPS！"
#> export render demo.docx -p docx
#> quit
```

## Agent 使用指南

1. **所有命令都支持 `--json`**，返回结构化 JSON
2. **使用 `--project` 标志**加载现有项目，然后链式执行命令
3. **会话自动保存**：单次命令模式下，退出时会自动保存修改
4. **REPL 模式**适合交互式探索
5. **导出前确保内容已保存**到项目文件

## 导出预设

| 预设 | 说明 | 适用类型 |
|------|------|---------|
| docx | Word 文档 | writer |
| doc | Word 97-2003 | writer |
| pdf | PDF（从 Writer） | writer |
| txt | 纯文本 | writer |
| html | 网页 | writer |
| xlsx | Excel 工作簿 | calc |
| xls | Excel 97-2003 | calc |
| csv | CSV | calc |
| pdf-calc | PDF（从 Calc） | calc |
| pptx | PowerPoint | impress |
| ppt | PowerPoint 97-2003 | impress |
| pdf-impress | PDF（从 Impress） | impress |

## 原理

此 CLI 通过 Windows COM 接口与 WPS Office 通信：
- `KWPS.Application` → WPS 文字
- `KET.Application` → WPS 表格
- `KWPP.Application` → WPS 演示

WPS COM API 与 Microsoft Office VBA API 高度兼容。

---

## PPT 设计风格系统

本工具内置 4 套设计预设、14 种布局模板、4 种演讲类型预设和 5 维度质量审查。
做 PPT 时自动调取，生成精美幻灯片。

### 一、设计预设（Design Presets）

通过 `preset apply <名称>` 一键切换。AI Agent 应在创建 PPT 前先调用此命令。

#### 1. 学术答辩 (academic)
- **来源**: scientific-slides skill
- **配色**: 深蓝 #1A3C8B / 橙 #E67733 / 绿 #188050
- **字体**: Arial + 微软雅黑，标题 40pt，正文 24pt，标注 16pt
- **规则**: 视觉优先 65% | 色盲友好 | 每页一个主题 | 留白 40% | 最多 6 条要点
- **适用**: 学术会议、论文答辩、基金申请、Journal Club

#### 2. 咨询顾问 (consultant)
- **来源**: pptx-from-layouts skill
- **配色**: 深蓝 #003366 / 亮青 #00A8E8 / 橙 #FF8C00
- **字体**: Arial + 微软雅黑，标题 36pt，正文 18pt，标注 14pt
- **规则**: 59种网格布局 | 中等密度 | 4色限制 | 最多 5 条要点
- **适用**: 商业计划书、咨询报告、年度汇报

#### 3. 商务汇报 (business)
- **来源**: pptx skill
- **配色**: 商务蓝 #005294 / 强调红 #C82828 / 绿 #2DA050
- **字体**: Arial + 微软雅黑，标题 36pt，正文 18pt，标注 14pt
- **规则**: 视觉 45% | 留白 35% | 最多 6 条要点
- **适用**: 会议汇报、项目提案、教学课件

#### 4. 科技极简 (tech)
- **来源**: 现代科技设计
- **配色**: 近黑 #0F1423 / 亮青 #00C8FF / 橙红 #FF643C
- **字体**: Arial + 微软雅黑，标题 44pt，正文 20pt，标注 14pt
- **规则**: 暗色模式 | 3色极简 | 留白 50% | 最多 4 条要点 | 低密度
- **适用**: 科技产品发布、AI/技术演示、数据报告

---

### 二、布局模板（Layout Templates）

14 种标准布局，覆盖 PPT 全场景。Agent 做 PPT 时按顺序选用。

| 模板名称 | 分类 | 用途 | 关键元素 |
|---------|------|------|---------|
| cover | 封面 | 开场标题页 | 72pt大字 + 装饰线 + 副标题 |
| toc | 目录 | 内容导航 | 蓝色竖线 + 圆角方块编号 + 6项 |
| overview | 概览 | 信息总览 | 深色横幅 + 2x4卡片 + 侧边荣誉栏 |
| timeline | 时间轴 | 历史/发展 | 圆点+竖线 + 日期+事件 + 名人语录卡 |
| grid_cards | 卡片网格 | 人物/地标/产品 | 4x2网格 + 头像圆 + 名字+简介 |
| quadrant | 四象限 | 对比/分类 | 2x2布局 + 彩色标题块 + 内容区 |
| stats | 数字统计 | 数据展示 | 6大数字 + 底部荣誉/说明区 |
| three_col | 三列对比 | 三种方案对比 | 深色全屏背景 + 三列并排 |
| pipeline | 流程图 | 管道/流程 | 6个彩色模块 + 连接箭头 + 总结条 |
| data_table | 数据表格 | 排名/指标 | 左表格 + 右解读区 |
| content_image | 图文 | 图文并茂 | 左文右图 / 左图右文 |
| closing | 结语 | 收尾页 | 深色全屏 + 总结 + 校训/联系方式 |

---

### 三、演讲类型预设（Talk Type Presets）

完整的幻灯片序列推荐。Agent 应按类型选用。

#### conference（学术会议，12-20页）
```
cover → toc → overview → timeline → quadrant → grid_cards
      → stats → pipeline → data_table → content_image
      → quadrant → timeline → stats → closing
```
规则: 最多 20 页 | 视觉 65% | 1-2 个核心发现

#### business（商务汇报，8-15页）
```
cover → toc → overview → stats → three_col
      → pipeline → grid_cards → data_table → closing
```
规则: 最多 15 页 | 视觉 50%

#### defense（论文答辩，45-65页）
```
cover → toc → overview → timeline → quadrant → content_image
      → pipeline → data_table → stats → quadrant → timeline
      → stats → closing
```
规则: 最多 65 页 | 视觉 60%

#### school（学校介绍，14页）
```
cover → toc → overview → timeline → three_col
      → grid_cards → quadrant → stats → closing
```
规则: 最多 14 页 | 视觉 55%

---

### 四、质量审查标准

5 维度审查，确保 PPT 出品质量。

| 维度 | 来源 | 检查项 | 通过阈值 |
|------|------|--------|---------|
| visual（视觉） | slide-excellence | 字体层级、颜色对比度、留白比例、布局一致性 | 70分 |
| pedagogy（教学法） | slide-excellence | 叙事弧完整性、预备知识、示例、符号一致性 | 75分 |
| proofreading（校对） | slide-excellence | 拼写、语法、术语、标点、字体溢出 | 80分 |
| parity（格式一致） | slide-excellence | PPTX/PDF一致、字体嵌入、图片、动画兼容 | 85分 |
| substance（内容实质） | slide-excellence | 数据准确性、引用完整、结论支撑 | 90分 |

---

### 五、Agent 做 PPT 标准流程

```
1. 确认主题和类型 → 选择预设和演讲类型
   cli-anything-wps preset apply academic --talk-type conference

2. 按演讲类型序列逐页生成内容
   每页应用对应的 layout template

3. 生成完成后执行质量审查
   确保 visual >= 70, pedagogy >= 75

4. 导出 PPTX + PDF 双格式
   cli-anything-wps export render output.pptx -p pptx
   cli-anything-wps export render output.pdf -p pdf
```
