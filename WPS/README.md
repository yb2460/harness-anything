# WPS PPT 自动化生成工作流

通过 **WPS COM 自动化** + **JSON 数据驱动** + **元素类型路由** 全自动生成 PPT。

## 核心架构

```
template_bg.png（16:9母版底图，WPS做母版→截图）
        +
ppt_data.json（全部中文内容，每页 elements[]）
        +
build_ppt.py（纯 ASCII 引擎，读 JSON 驱动 WPS COM）
        ↓
  WPS 可见生成 → PPTX + PDF
```

## 项目列表

| 项目 | 主题色 | 校训/关键词 | 页数 |
|------|--------|------|------|
| [清华协和](清华协和/) | `#003D5C` 蓝 | 严谨 博精 创新 奉献 | 12页 |
| [兰州大学](兰州大学/) | `#003D5C` 蓝 | 自强不息 独树一帜 | 12页 |
| [同济医学院](同济医学院/) | `#C41230` 红 | 同舟共济 | 12页 |
| [哈工大](哈工大/) | `#00508F` 蓝 | 规格严格 功夫到家 | 12页 |
| [重庆大学](重庆大学/) | `#004B8D` 蓝 | 耐劳苦 尚俭朴 | 12页 |
| [南华大学](南华大学/) | `#004C97` 蓝 | 核特班+卓越医师班 | 12页 |
| [中山大学](中山大学/) | `#00663D` 绿 | 博学 审问 慎思 明辨 笃行 | 15页 |
| [中科大](中科大/) | `#003B6F` 蓝 | 红专并进 理实交融 | 15页 |
| [国科大](国科大/) | `#003B6F` 蓝 | 博学笃志 格物明德 | 15页 |
| [复旦大学](复旦大学/) | `#00356B` 蓝 | 日月光华 旦复旦兮 | 12页 |
| [南科大肿瘤医院](南科大肿瘤医院/) | `#005A8B` 蓝 | 联合培养 硕博招生 | 12页 |
| [北大](北大/) | `#8B0012` 红 | 爱国 进步 民主 科学 | 14页 |
| [清华](清华/) | `#660874` 紫 | 自强不息 厚德载物 | 13页 |
| [南科大](南科大/) | `#006B3F` 绿 | 明德求是 日新自强 | 10页 |
| [浙大城市学院](浙大城市学院/) | `#005A9C` 蓝 | 求是创新 | 10页 |
| [华中科大](华中科大/) | `#004098` 蓝 | 明德厚学 求是创新 | 9页 |

## 使用方法

### 1. 准备模板背景

```bash
# WPS打开模版.pptx → 另存为PNG(960x540) → template_bg.png
```

### 2. 编写数据JSON

```json
{
  "canvas": {"w": 960, "h": 540},
  "slides": [
    {
      "id": 1, "title": "封面",
      "elements": [
        {"type": "text", "x": 60, "y": 60, "w": 840, "h": 150, "text": "标题", "fs": 48, "color": "#004098", "bold": true, "align": 2, "font": "SimHei"},
        {"type": "cards_1x4_info", "items": [{"num": "数据", "label": "标签"}], "start_y": 340}
      ]
    }
  ]
}
```

### 3. 运行生成

```bash
taskkill //F //IM wps.exe //T 2>/dev/null
taskkill //F //IM wpp.exe //T 2>/dev/null
python build_xxx.py
```

## 已支持元素类型

| type | 用途 | 参数 |
|------|------|------|
| `text` | 文本框 | x,y,w,h,text,fs,color,bold,align,font,line_spacing |
| `image` | 图片 | x,y,w,h,file |
| `table` | 数据表格 | x,y,w,h,rows,cols,data,header_color |
| `card_list_wide` | 目录列表 | items[{num,title,sub}], start_y |
| `cards_1x4_info` | 4列统计卡 | items[{num,label}], start_y |
| `cards_2x3` | 2行x3列网格 | items[{title,desc,color}], start_y |
| `tagline_bar` | 底部总结条 | x,y,w,h,text,color |

## 布局规范（2026-07-06 更新）

| 规则 | 标准 |
|------|------|
| 背景 | 模板底图铺满，绝不遮挡 |
| 标题 | 居中align=2，SimHei 40-44pt，品牌色，**无装饰线**，**最多4字** |
| **标题-内容间距** ⚠️ | **≥24pt（约2行）**。标题y=14,h=40→结束于y≈54，第一个内容元素**必须**起始于y≥76-78。**绝不让内容紧贴标题！** |
| 正文 | Microsoft YaHei 15-20pt，黑色 #333 |
| 卡片 | 仅顶部4-5pt彩色细线，**无灰色填充** |
| 表格 | 品牌色表头白字 + 隔行交替着色 |
| 图表 | matplotlib transparent=True |
| 底部条 | tagline_bar y≈498, h≈28 |
| 内容密度 | 每页 ≤80 中文字 + ≥2 种视觉元素（图表/表格/卡片/编号圆） |
| 元素边界 | 所有元素 y+height ≤518，x+width ≤960 |

## matplotlib 图表规范

- DPI: 250以上，`transparent=True`
- 字体: SimHei（标题）/ Microsoft YaHei（标签）
- 品牌色: 从校徽提取，占视觉60-70%
- 绝不用: 默认蓝、彩虹色、低对比灰
- 优先: 柱状图（分组/横向）、饼图、折线/面积图

## 依赖

- Windows + WPS Office
- Python: `pywin32`, `matplotlib`, `numpy`
