# -*- coding: utf-8 -*-
"""导出模块 —— 将 PSD 导出为 PNG、JPEG、WebP 等格式。"""
import os
import win32com.client
from ..utils.com_bridge import get_bridge


def export(path: str, fmt: str = None, quality: int = 90,
           width: int = None, height: int = None) -> dict:
    """导出文档为指定格式。"""
    bridge = get_bridge()
    doc = bridge.require_document()

    abs_path = os.path.abspath(path)
    os.makedirs(os.path.dirname(abs_path), exist_ok=True)

    # 自动检测格式
    ext = fmt or os.path.splitext(path)[1].lstrip(".").lower()

    # 如果需要调整尺寸，先创建副本
    source_doc = doc
    temp_doc = None
    if width or height:
        from win32com.client import constants as ps
        w = width or doc.Width
        h = height or doc.Height
        temp_doc = doc.Duplicate()
        temp_doc.ResizeImage(w, h, doc.Resolution)
        source_doc = temp_doc

    ext_handlers = {
        "png": _export_png, "jpg": _export_jpg, "jpeg": _export_jpg,
        "bmp": _export_bmp, "gif": _export_gif, "tiff": _export_tiff,
        "tif": _export_tiff, "pdf": _export_pdf,
    }

    handler = ext_handlers.get(ext)
    if handler is None:
        raise ValueError(f"不支持的导出格式: {ext}，可选: {list(ext_handlers.keys())}")

    result = handler(source_doc, abs_path, quality)
    file_size = os.path.getsize(abs_path) if os.path.exists(abs_path) else 0

    return {
        "exported": abs_path,
        "format": ext.upper(),
        "size_bytes": file_size,
        "status": "ok",
    }


def save_as_jpeg(path: str, quality: int = 90) -> dict:
    """保存为 JPEG 格式。"""
    return export(path, "jpg", quality)


def save_as_png(path: str) -> dict:
    """保存为 PNG 格式。"""
    return export(path, "png")


def _export_png(doc, path: str, quality: int) -> dict:
    """导出为 PNG。"""
    # 本地补丁: Photoshop CS6 的 Save-for-Web 在默认参数下会输出超大错误位图
    # （实测 18142x11311 的 GIF 写到 .png）。优先 SaveAs + PNGSaveOptions 并校验 PNG 头。
    try:
        options = win32com.client.Dispatch("Photoshop.PNGSaveOptions")
        doc.SaveAs(path, options, True)
        with open(path, "rb") as handle:
            head = handle.read(8)
        if head.startswith(b"\x89PNG\r\n\x1a\n"):
            return {"format": "PNG", "path": path, "method": "SaveAs+PNGSaveOptions"}
    except Exception:
        pass

    options = win32com.client.Dispatch(
        "Photoshop.ExportOptionsSaveForWeb"
    )
    options.Format = 2  # psPNGFormat
    options.PNG8 = False
    options.Transparency = True
    doc.Export(path, 2, options)  # psSaveForWeb
    return {"format": "PNG", "path": path, "method": "SaveForWeb"}


def _export_jpg(doc, path: str, quality: int) -> dict:
    """导出为 JPEG。"""
    options = win32com.client.Dispatch(
        "Photoshop.ExportOptionsSaveForWeb"
    )
    options.Format = 1  # psJPEGFormat
    options.Quality = quality
    doc.Export(path, 2, options)  # psSaveForWeb
    return {"format": "JPEG", "path": path}


def _export_bmp(doc, path: str, quality: int) -> dict:
    doc.SaveAs(path, None, True)  # BMP via SaveAs
    return {"format": "BMP", "path": path}


def _export_gif(doc, path: str, quality: int) -> dict:
    options = win32com.client.Dispatch(
        "Photoshop.ExportOptionsSaveForWeb"
    )
    options.Format = 4  # psGIFFormat
    doc.Export(path, 2, options)
    return {"format": "GIF", "path": path}


def _export_tiff(doc, path: str, quality: int) -> dict:
    options = win32com.client.Dispatch(
        "Photoshop.TiffSaveOptions"
    )
    doc.SaveAs(path, options, True)
    return {"format": "TIFF", "path": path}


def _export_pdf(doc, path: str, quality: int) -> dict:
    options = win32com.client.Dispatch(
        "Photoshop.PDFSaveOptions"
    )
    doc.SaveAs(path, options, True)
    return {"format": "PDF", "path": path}
