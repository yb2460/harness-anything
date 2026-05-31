# PowerPoint CLI Harness — 软件分析 SOP

## 后端引擎

| 引擎 | 用途 | 接口 |
|------|------|------|
| **python-pptx** | 数据层：创建/编辑 PPTX 文件 | Python API |
| **WPS COM (KWPP.Application)** | 渲染层 Windows：导出 PDF/图片 | COM 自动化 (win32com) |
| **LibreOffice Impress** | 渲染层跨平台：导出 PDF/图片 | `soffice --headless --convert-to` |
| **Markitdown** | 文本提取 | `python -m markitdown` |

## GUI 操作 → API 映射

| GUI 操作 | python-pptx API | CLI 命令 |
|----------|-----------------|----------|
| 新建演示文稿 | `Presentation()` | `project new` |
| 打开演示文稿 | `Presentation(path)` | `project open` |
| 保存 | `presentation.save(path)` | `project save` |
| 添加幻灯片 | `slides.add_slide(layout)` | `slides add` |
| 删除幻灯片 | `del slides[i]` | `slides remove` |
| 添加文本框 | `shapes.add_textbox()` | `content add-text` |
| 添加形状 | `shapes.add_shape()` | `content add-shape` |
| 添加图片 | `shapes.add_picture()` | `content add-image` |
| 添加表格 | `shapes.add_table()` | `content add-table` |
| 导出 PDF | WPS COM / LibreOffice | `export render` |
| 缩略图 | LibreOffice PDF → pdftoppm | `preview thumbnails` |

## 数据模型 (PPTX 文件格式)

PPTX 是 ZIP 压缩的 OOXML 包：
- `ppt/presentation.xml` — 幻灯片顺序和全局属性
- `ppt/slides/slide{N}.xml` — 每张幻灯片内容
- `ppt/slideLayouts/slideLayout{N}.xml` — 布局定义
- `ppt/slideMasters/slideMaster{N}.xml` — 母版定义
- `ppt/theme/theme{N}.xml` — 主题/配色/字体
- `ppt/media/` — 图片/视频资源

## 渲染后端选择

优先级：
1. WPS COM (Windows，用户环境已配置) — 最可靠
2. LibreOffice Impress (跨平台) — 备选
3. python-pptx 原生保存 — 仅 PPTX 格式
