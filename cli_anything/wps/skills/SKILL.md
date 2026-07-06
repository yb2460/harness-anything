---
name: "cli-anything-wps"
description: "WPS Office CLI — JSON数据驱动PPT自动生成 + 命令行操控WPS文字/表格/演示文稿"
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
cli-anything-wps --help
```

---

## JSON 数据驱动 PPT 自动生成 ⭐ 生产环境核心工作流

**这是最常用的模式**：搜数据 → 生成 matplotlib 图表 → 写 JSON → WPS COM 一键生成。

### 完整流程（6步）

```
1. 提取模板背景: 模版.pptx → template_bg.png (960x540)
2. WebSearch 搜索数据: 招生分数线/学科排名/招生计划/科研数据
3. matplotlib 生成图表: 柱状图/饼图/折线图/横向柱/气泡图
4. 编写 data.json: 12-15页 elements[] 编排（标题≤4字 + 间距≥24pt）
5. python build_xxx.py: WPS COM 逐页构建
6. 输出: PPTX + PDF 双格式
```

### 标准页序

```
S1 封面 → S2 目录 → S3-S10 内容页(图+表+卡片) → S11 总结卡片 → S12 致谢
```

### 元素类型路由

| type | 用途 |
|------|------|
| `text` | 文本框 |
| `image` | 图片(含matplotlib图表) |
| `table` | 表格 |
| `cards_2x3` | 2行x3列彩色卡片 |
| `cards_1x4_info` | 4列数字统计卡 |
| `card_list_wide` | 目录编号列表 |
| `tagline_bar` | 页面底部总结条 |

### 关键约束 ⚠️

- **标题与内容间距**：标题 y=14 h=40(结束于y≈54)，第一个内容元素**必须**起始于 y≥76-78(≥24pt gap)
- **标题最多4字**，居中，SimHei 40-44pt，品牌色，**无装饰线**
- **JSON中文引号**：文本内引用用「」代替 `""`
- **所有元素不出画布**：960×540，y+h≤518
- **WPS COM**：`Fill.ForeColor.RGB`，`SaveAs(path, 32)`导出PDF
- **执行前清理**：`taskkill //F //IM wps.exe //T`

### 参考案例

[WPS/](../WPS/) 目录下16个项目：

| 项目 | 页数 | 主题 |
|------|------|------|
| 清华协和/兰州大学/同济医学院/哈工大/重庆大学/南华大学 | 12页 | 各校招生 |
| 中山大学/中科大/国科大 | 15页 | 名校招生 |
| 复旦大学 | 12页 | 新工科+医学院 |
| 南科大肿瘤医院 | 12页 | 联合培养硕博 |
| 北大/清华/南科大/华中科大/浙大城市学院 | 9-14页 | 各校介绍 |

```bash
# 完整执行示例
python -c "import zipfile; ..."  # 提取模板
python gen_charts.py              # 生成图表
python -c "import json; json.load(open('data.json','r',encoding='utf-8'))"  # 验证JSON
taskkill //F //IM wps.exe //T
python build_xxx.py               # 构建PPTX+PDF
```

---

## CLI 命令结构

```
wps
├── document          # 文档管理: new/open/save/info
├── writer            # 文字: add-paragraph/heading/list/table/image
├── calc              # 表格: set-cell/get-cell/set-range/merge-cells
├── impress           # 演示: add-slide/remove-slide/set-content/add-element
├── style             # 样式: create/modify/list/apply/remove
├── preset            # 设计预设: list/info/apply
├── export            # 导出: render output.pptx -p pptx
└── session           # 会话: status/undo/redo/history
```

### 使用示例

```bash
# 创建文档
cli-anything-wps document new --type writer --name "报告" -o report.json
cli-anything-wps --project report.json writer add-heading -t "年度报告" -l 1
cli-anything-wps --project report.json export render report.docx -p docx

# JSON模式（Agent使用）
cli-anything-wps --json document new --type writer --name "test"
```

### 设计预设

| 预设 | 配色 | 适用场景 |
|------|------|---------|
| academic | 深蓝#1A3C8B | 学术答辩/基金申请 |
| consultant | 深蓝#003366+亮青 | 咨询报告 |
| business | 商务蓝#005294 | 会议汇报/课件 |
| tech | 近黑#0F1423 | 科技/AI/数据报告 |

### 质量审查（5维度）

| 维度 | 检查项 | 阈值 |
|------|--------|------|
| visual | 字体层级/对比度/留白 | 70分 |
| pedagogy | 叙事弧/预备知识 | 75分 |
| proofreading | 拼写/术语/溢出 | 80分 |
| parity | PPTX/PDF一致性 | 85分 |
| substance | 数据准确性/引用 | 90分 |
