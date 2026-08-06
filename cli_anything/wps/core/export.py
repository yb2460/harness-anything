"""WPS CLI - 导出/渲染管道。

将项目 JSON 转换为真实文档文件，通过 WPS COM 自动化完成。
"""

import os
import tempfile
from typing import Dict, Any, Optional

try:
    from cli_anything.wps.utils.wps_backend import (
        find_wps, create_document, save_as, export_pdf,
        close_document, quit_app, get_version,
        PROGID_MAP,
    )
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

# WPS 文件格式映射
WPS_FORMAT_MAP = {
    "writer": {
        "wps": (".wps", "WPS 文字"),
        "docx": (".docx", "Word 文档"),
        "doc": (".doc", "Word 97-2003 文档"),
        "pdf": (".pdf", "PDF"),
        "txt": (".txt", "纯文本"),
        "rtf": (".rtf", "RTF 格式"),
        "html": (".html", "网页"),
        "xps": (".xps", "XPS"),
    },
    "calc": {
        "et": (".et", "WPS 表格"),
        "xlsx": (".xlsx", "Excel 工作簿"),
        "xls": (".xls", "Excel 97-2003 工作簿"),
        "csv": (".csv", "CSV 逗号分隔"),
        "pdf": (".pdf", "PDF"),
        "html": (".html", "网页"),
    },
    "impress": {
        "dps": (".dps", "WPS 演示"),
        "pptx": (".pptx", "PowerPoint 演示"),
        "ppt": (".ppt", "PowerPoint 97-2003 演示"),
        "pdf": (".pdf", "PDF"),
        "html": (".html", "网页"),
    },
}

EXPORT_PRESETS = {
    # Writer 导出
    "docx": {
        "name": "Word 文档 (.docx)",
        "doc_type": "writer",
        "format": "docx",
        "extension": ".docx",
        "description": "Word 文档格式",
    },
    "doc": {
        "name": "Word 97-2003 (.doc)",
        "doc_type": "writer",
        "format": "doc",
        "extension": ".doc",
        "description": "Word 97-2003 兼容格式",
    },
    "pdf": {
        "name": "PDF",
        "doc_type": "writer",
        "format": "pdf",
        "extension": ".pdf",
        "description": "PDF 便携文档格式",
    },
    "txt": {
        "name": "纯文本 (.txt)",
        "doc_type": "writer",
        "format": "txt",
        "extension": ".txt",
        "description": "纯文本文件",
    },
    "html": {
        "name": "网页 (.html)",
        "doc_type": "writer",
        "format": "html",
        "extension": ".html",
        "description": "HTML 网页格式",
    },
    # Calc 导出
    "xlsx": {
        "name": "Excel 工作簿 (.xlsx)",
        "doc_type": "calc",
        "format": "xlsx",
        "extension": ".xlsx",
        "description": "Excel 工作簿格式",
    },
    "xls": {
        "name": "Excel 97-2003 (.xls)",
        "doc_type": "calc",
        "format": "xls",
        "extension": ".xls",
        "description": "Excel 97-2003 兼容格式",
    },
    "csv": {
        "name": "CSV 逗号分隔 (.csv)",
        "doc_type": "calc",
        "format": "csv",
        "extension": ".csv",
        "description": "CSV 逗号分隔文本",
    },
    "pdf-calc": {
        "name": "PDF（从 Calc）",
        "doc_type": "calc",
        "format": "pdf",
        "extension": ".pdf",
        "description": "PDF 从电子表格导出",
    },
    # Impress 导出
    "pptx": {
        "name": "PowerPoint 演示 (.pptx)",
        "doc_type": "impress",
        "format": "pptx",
        "extension": ".pptx",
        "description": "PowerPoint 演示文稿",
    },
    "ppt": {
        "name": "PowerPoint 97-2003 (.ppt)",
        "doc_type": "impress",
        "format": "ppt",
        "extension": ".ppt",
        "description": "PowerPoint 97-2003 兼容格式",
    },
    "pdf-impress": {
        "name": "PDF（从 Impress）",
        "doc_type": "impress",
        "format": "pdf",
        "extension": ".pdf",
        "description": "PDF 从演示文稿导出",
    },
}


def list_presets() -> Dict[str, Any]:
    """列出所有导出预设。"""
    return {
        name: info
        for name, info in EXPORT_PRESETS.items()
    }


def get_preset_info(name: str) -> Optional[Dict[str, Any]]:
    """获取指定预设的详细信息。"""
    preset = EXPORT_PRESETS.get(name)
    if not preset:
        available = ", ".join(EXPORT_PRESETS.keys())
        raise ValueError(f"不支持的导出预设: {name}。可用: {available}")
    return dict(preset)


def export(
    project: Dict[str, Any],
    output_path: str,
    preset: str = "docx",
    overwrite: bool = False,
) -> Dict[str, Any]:
    """导出文档到指定格式（使用 WPS COM 自动化）。

    Args:
        project: 项目字典
        output_path: 输出文件路径
        preset: 导出预设名称
        overwrite: 是否覆盖已有文件

    Returns:
        包含输出路径、格式、文件大小等信息的字典

    Raises:
        RuntimeError: COM 不可用或导出失败
        FileExistsError: 输出文件已存在且未启用覆盖
    """
    if not COM_AVAILABLE:
        raise RuntimeError(
            "WPS COM 自动化不可用。请在 Windows 上运行并确保 WPS Office 已安装。"
            "此外需安装 pywin32: pip install pywin32"
        )

    preset_info = get_preset_info(preset)
    output_path = os.path.abspath(output_path)

    if os.path.exists(output_path) and not overwrite:
        raise FileExistsError(
            f"输出文件已存在: {output_path}。使用 --overwrite 覆盖。"
        )

    doc_type = preset_info["doc_type"]
    fmt = preset_info["format"]
    extension = preset_info["extension"]

    # 确保输出文件有正确的扩展名
    if not output_path.lower().endswith(extension):
        output_path = output_path + extension

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # 通过 COM 创建文档并填充内容
    app = None
    doc = None
    try:
        app = find_wps(doc_type)
        # WPS 演示（KWPP.Application）的 Visible 属性只读，赋值会抛 com_error；
        # 且 WPS 默认即后台运行，故 impress 跳过该赋值（writer/calc 保持原行为）。
        if doc_type != "impress":
            app.Visible = False  # 后台运行

        doc = create_document(app, doc_type)
        _fill_document(doc, project, doc_type)

        # 如果是 PDF，使用专用导出方法
        if fmt == "pdf":
            export_pdf(doc, output_path, doc_type)
        else:
            save_as(doc, output_path, doc_type, fmt)

        file_size = os.path.getsize(output_path)

        result = {
            "output": os.path.abspath(output_path),
            "format": fmt,
            "extension": extension,
            "preset": preset,
            "file_size": file_size,
            "method": "wps-com-automation",
            "wps_version": get_version(app),
        }

        close_document(doc, save=False)
        quit_app(app)
        return result

    except Exception:
        if doc:
            try:
                close_document(doc, save=False)
            except Exception:
                pass
        if app:
            try:
                quit_app(app)
            except Exception:
                pass
        raise


def _fill_document(doc, project: Dict[str, Any], doc_type: str) -> None:
    """将项目数据填充到 WPS COM 文档对象中。

    这是核心的数据到文档转换逻辑。
    """
    if doc_type == "writer":
        _fill_writer(doc, project)
    elif doc_type == "calc":
        _fill_calc(doc, project)
    elif doc_type == "impress":
        _fill_impress(doc, project)


def _fill_writer(doc, project: Dict[str, Any]) -> None:
    """将内容填充到 WPS Writer 文档。"""
    content_items = project.get("content", [])

    if not content_items:
        # 无内容时设置文档正文
        return

    # 使用 Range API 逐项写入
    for item in content_items:
        item_type = item.get("type", "paragraph")

        try:
            if item_type == "paragraph":
                _add_writer_paragraph(doc, item)
            elif item_type == "heading":
                _add_writer_heading(doc, item)
            elif item_type == "list":
                _add_writer_list(doc, item)
            elif item_type == "table":
                _add_writer_table(doc, item)
            elif item_type == "page_break":
                _add_writer_page_break(doc)
            elif item_type == "image_ref":
                _add_writer_image(doc, item)
        except Exception as e:
            raise RuntimeError(f"写入内容项失败（类型: {item_type}）: {e}")


def _add_writer_paragraph(doc, item: Dict[str, Any]) -> None:
    """在 Writer 中添加段落。"""
    # 获取文档末尾位置
    rng = doc.Range()
    rng.Collapse(0)  # wdCollapseEnd = 0, 折叠到末尾

    text = item.get("text", "")
    style = item.get("style", {})

    # 如果文本不为空，写入后添加换行
    if text:
        rng.InsertAfter(text)
        rng.InsertParagraphAfter()

    # 对段落应用样式（通过最后写入的段落）
    _apply_paragraph_style(doc, style)


def _add_writer_heading(doc, item: Dict[str, Any]) -> None:
    """在 Writer 中添加标题。"""
    text = item.get("text", "")
    level = item.get("level", 1)
    style = item.get("style", {})

    rng = doc.Range()
    rng.Collapse(0)

    rng.InsertAfter(text)

    # WPS COM 不支持 SetStyle 属性，使用 OutlineLevel + 手动格式
    para = doc.Paragraphs.Last
    try:
        para.Range.ParagraphFormat.OutlineLevel = min(level, 9)
    except Exception:
        pass

    # 根据级别设置字体大小和粗体
    font_sizes = {1: 22, 2: 18, 3: 16, 4: 14, 5: 13, 6: 12}
    try:
        para.Range.Font.Size = font_sizes.get(level, 14)
        para.Range.Font.Bold = True
    except Exception:
        pass

    rng.InsertParagraphAfter()


def _apply_paragraph_style(doc, style: Dict[str, Any]) -> None:
    """将样式属性应用到文档最后一段。"""
    try:
        para = doc.Paragraphs.Last
        p_format = para.Range.Font

        if "font_size" in style:
            size = style["font_size"].replace("pt", "").strip()
            try:
                p_format.Size = float(size)
            except ValueError:
                pass

        if "bold" in style and style["bold"]:
            p_format.Bold = True
        if "italic" in style and style["italic"]:
            p_format.Italic = True
        if "underline" in style and style["underline"]:
            p_format.Underline = True
    except Exception:
        pass  # 样式应用失败不中断导出


def _add_writer_list(doc, item: Dict[str, Any]) -> None:
    """在 Writer 中添加列表。"""
    items = item.get("items", [])
    list_style = item.get("list_style", "bullet")
    rng = doc.Range()
    rng.Collapse(0)

    for i, text in enumerate(items):
        prefix = "• " if list_style == "bullet" else f"{i + 1}. "
        rng.InsertAfter(prefix + text)
        rng.InsertParagraphAfter()


def _add_writer_table(doc, item: Dict[str, Any]) -> None:
    """在 Writer 中添加表格。"""
    rows = item.get("rows", 2)
    cols = item.get("cols", 2)
    data = item.get("data", [])

    rng = doc.Range()
    rng.Collapse(0)
    table = doc.Tables.Add(rng, rows, cols)
    table.AutoFitBehavior(2)  # wdAutoFitWindow = 2

    for ri, row in enumerate(data):
        for ci, val in enumerate(row):
            if ri < rows and ci < cols:
                try:
                    table.Cell(ri + 1, ci + 1).Range.Text = str(val)
                except Exception:
                    pass

    doc.Range().InsertParagraphAfter()


def _add_writer_page_break(doc) -> None:
    """在 Writer 中添加分页符。"""
    rng = doc.Range()
    rng.Collapse(0)
    rng.InsertBreak(7)  # wdPageBreak = 7


def _add_writer_image(doc, item: Dict[str, Any]) -> None:
    """在 Writer 中添加图片。"""
    image_path = item.get("path", "")
    if not os.path.exists(image_path):
        return

    rng = doc.Range()
    rng.Collapse(0)
    try:
        doc.InlineShapes.AddPicture(image_path, Range=rng)
    except Exception:
        pass
    doc.Range().InsertParagraphAfter()


def _fill_calc(doc, project: Dict[str, Any]) -> None:
    """将内容填充到 WPS Calc 工作簿。"""
    sheets = project.get("sheets", [])

    # 删除默认工作表，重建
    try:
        while doc.Worksheets.Count > len(sheets):
            doc.Worksheets(1).Delete()
    except Exception:
        pass

    for si, sheet_data in enumerate(sheets):
        if si == 0:
            ws = doc.Worksheets(1)
        else:
            ws = doc.Worksheets.Add()
        ws.Name = sheet_data.get("name", f"Sheet{si + 1}")

        # 填充单元格
        for ref, cell in sheet_data.get("cells", {}).items():
            try:
                cell_value = cell.get("value", "")
                cell_type = cell.get("type", "string")

                # 根据类型设置值
                if cell_type == "float":
                    try:
                        ws.Range(ref).Value = float(cell_value)
                    except (ValueError, TypeError):
                        ws.Range(ref).Value = str(cell_value)
                else:
                    ws.Range(ref).Value = str(cell_value)

                # 设置公式
                if cell.get("formula"):
                    ws.Range(ref).Formula = cell["formula"]
            except Exception:
                pass

        # 合并单元格
        for merge in sheet_data.get("merged_cells", []):
            try:
                ws.Range(f"{merge['start']}:{merge['end']}").Merge()
            except Exception:
                pass


def _fill_impress(doc, project: Dict[str, Any]) -> None:
    """将内容填充到 WPS Impress 演示文稿。"""
    slides = project.get("slides", [])

    # 如果内容为空，至少保留一张幻灯片
    if not slides:
        return

    # WPS 演示 Presentations.Add() 产生空演示文稿（Slides.Count=0），
    # 与 MS PowerPoint 不同。此处兼容：若无幻灯片则用 Slides.Add 占位（ppLayoutText=2）。
    if doc.Slides.Count == 0:
        doc.Slides.Add(1, 2)

    for si, slide_data in enumerate(slides):
        if si == 0:
            slide = doc.Slides(1)
        else:
            slide = doc.Slides.Add(si + 1, 2)  # ppLayoutText = 2

        # 设置标题和内容（通过占位符）
        title = slide_data.get("title", "")
        content = slide_data.get("content", "")

        for shape in slide.Shapes:
            try:
                if shape.Type == 14 and title:  # msoPlaceholder = 14
                    if "Title" in str(shape.PlaceholderFormat.Type):
                        shape.TextFrame.TextRange.Text = title
            except Exception:
                pass
            try:
                if shape.HasTextFrame and content:
                    shape.TextFrame.TextRange.Text = content
            except Exception:
                pass


