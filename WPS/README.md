# WPS PPT 自动化生成工作流

通过 **WPS COM 自动化** + **JSON 数据驱动** + **元素类型路由** 全自动生成 PPT。

## 核心架构

```
template_bg.png（16:9母版底图，WPS做母版→截图）
        +
ppt_data.json（全部中文内容）
        +
build_ppt.py（纯ASCII引擎，读JSON驱动WPS COM）
        ↓
  WPS 可见生成 → PPTX + PDF
```

## 项目列表

| 项目 | 主题色 | 校训 | 页数 |
|------|--------|------|------|
| [南科大](南科大/) | `#006B3F` 绿 | 明德求是 日新自强 | 10页 |
| [北大](北大/) | `#8B0012` 红 | 爱国 进步 民主 科学 | 14页 |
| [清华](清华/) | `#660874` 紫 | 自强不息 厚德载物 | 13页 |
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
        {"type": "cards_1x4_info", "items": [{"num": "数据", "label": "标签"}, ...], "start_y": 340}
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
| `line` | 装饰线 | x,y,w,h,color |
| `card_list_wide` | 目录列表 | items[{num,title,sub}], start_y |
| `cards_1x4_info` | 4列统计卡 | items[{num,label}], start_y |
| `cards_1x3_big` | 3列大卡片 | items[{title,desc,color}], start_y |
| `cards_2x3` | 2行×3列网格 | items[{title,desc,color}], start_y |
| `cards_2x2_four` | 2行×2列卡片 | items[{title,desc,color}], start_y |
| `card_row_5` | 5列图标卡 | items[{icon,title,desc,color}], start_y |
| `timeline_horiz` | 水平时间轴 | items[{year,label,desc,color}], start_y |
| `tagline_bar` | 底部总结条 | x,y,w,h,text,color |
| `quote_block` | 引用块 | x,y,w,h,text,color |
| `stories_2col` | 2列故事卡 | items[{title,desc}], start_y |

## 布局规范（2026-06-11 南科大版确立）

| 规则 | 标准 |
|------|------|
| 背景 | 模板底图铺满，绝不遮挡 |
| 标题 | 居中align=2，SimHei 40-44pt，品牌色，**无装饰线** |
| 正文 | Microsoft YaHei 22-28pt，黑色 #333 |
| 卡片 | 仅顶部4-5pt彩色细线，**无灰色填充** |
| 表格 | 品牌色表头白字 + 隔行交替着色 |
| 图表 | matplotlib transparent=True |
| 底部条 | tagline_bar y≈498, h≈28 |

## 依赖

- Windows + WPS Office
- Python 3.x: `win32com`, `matplotlib`, `numpy`
