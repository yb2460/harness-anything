"""WPS COM 后端 —— 通过 COM 自动化接口操控 WPS Office。

本模块是 CLI 与 WPS Office 之间的桥梁。利用 Windows COM 接口
（KWPS.Application / KET.Application / KWPP.Application）实现：

- 文档的创建、打开和保存
- 内容编辑（文字、表格、演示文稿）
- 格式导出（PDF、DOCX、XLSX、PPTX 等）

要求：Windows 系统 + WPS Office（已安装）+ pywin32

WPS COM ProgID:
  KWPS.Application  → WPS 文字（类 Word）
  KET.Application   → WPS 表格（类 Excel）
  KWPP.Application  → WPS 演示（类 PPT）
"""

import os
import sys
import time
import shutil
import subprocess
from typing import Optional, Tuple

# COM 常量 —— 与 Microsoft Office 兼容
wdFormatDocumentDefault = 16      # .docx（默认）
wdFormatPDF = 17                  # .pdf
wdFormatTemplate = 1              # .dot
wdFormatText = 2                  # .txt
wdFormatTextLineBreaks = 3        # .txt（带换行）
wdFormatDOSText = 4               # .txt（DOS）
wdFormatDOSTextLineBreaks = 5     # .txt（DOS 带换行）
wdFormatRTF = 6                   # .rtf
wdFormatHTML = 8                  # .html（Office 2000+）
wdFormatWebArchive = 9            # .mhtml
wdFormatFilteredHTML = 10         # .html（过滤后）
wdFormatXML = 11                  # .xml
wdFormatDocument97 = 0            # .doc（97-2003）
wdFormatXMLDocument = 12          # .docx（XML）
wdFormatFlatXML = 19              # .xml（Flat OPC）
wdFormatOpenDocumentText = 23     # .odt
wdFormatXPS = 18                  # .xps

# Excel 格式常量
xlOpenXMLWorkbook = 51            # .xlsx
xlWorkbookDefault = 51            # .xlsx
xlCSV = 6                         # .csv
xlCSVUTF8 = 62                    # .csv（UTF-8）
xlCSVMSDOS = 24                   # .csv（MS-DOS）
xlHtml = 44                       # .html
xlTextPrinter = 36                # .prn
xlExcel9795 = 43                  # .xls（95）
xlWorkbookNormal = -4143          # .xls（97-2003）
xlPDF = 0                         # PDF

# PPT 格式常量
ppSaveAsPresentation = 1          # .pptx
ppSaveAsPowerPoint7 = 2           # .ppt（95）
ppSaveAsPowerPoint4 = 3           # .ppt（4.0）
ppSaveAsPowerPoint3 = 4           # .ppt（3.0）
ppSaveAsPDF = 32                  # .pdf
ppSaveAsHTML = 12                 # .html
ppSaveAsOpenXMLPresentation = 1   # .pptx
ppSaveAsOpenXMLShow = 36          # .ppsx

# WPS COM ProgID 映射
PROGID_MAP = {
    "writer": "KWPS.Application",
    "wps": "KWPS.Application",
    "calc": "KET.Application",
    "et": "KET.Application",
    "impress": "KWPP.Application",
    "wpp": "KWPP.Application",
}

# WPS 可执行文件搜索路径
WPS_PATHS = [
    os.path.join(os.environ.get("LOCALAPPDATA", ""),
                 r"Kingsoft\WPS Office"),
    os.path.join(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"),
                 r"Kingsoft\WPS Office"),
    os.path.join(os.environ.get("PROGRAMFILES", r"C:\Program Files"),
                 r"Kingsoft\WPS Office"),
    r"C:\Program Files (x86)\Kingsoft\WPS Office",
    r"C:\Program Files\Kingsoft\WPS Office",
]

# 格式映射表
FORMAT_SAVEAS_MAP = {
    "writer": {
        "docx": wdFormatDocumentDefault,
        "doc": wdFormatDocument97,
        "pdf": wdFormatPDF,
        "txt": wdFormatDOSTextLineBreaks,
        "html": wdFormatFilteredHTML,
        "rtf": wdFormatRTF,
        "xml": wdFormatXMLDocument,
        "odt": wdFormatOpenDocumentText,
        "xps": wdFormatXPS,
    },
    "calc": {
        "xlsx": xlOpenXMLWorkbook,
        "xls": xlWorkbookNormal,
        "csv": xlCSVUTF8,
        "html": xlHtml,
        "pdf": xlPDF,
    },
    "impress": {
        "pptx": ppSaveAsOpenXMLPresentation,
        "ppt": ppSaveAsPowerPoint7,
        "pdf": ppSaveAsPDF,
        "html": ppSaveAsHTML,
    },
}


def find_wps_executable() -> str:
    """查找 WPS 可执行文件路径。

    遍历已知安装目录，返回最新版本 wps.exe 路径。
    """
    candidates = []
    for base in WPS_PATHS:
        if not os.path.isdir(base):
            continue
        for item in os.listdir(base):
            item_path = os.path.join(base, item)
            if not os.path.isdir(item_path):
                continue
            wps_exe = os.path.join(item_path, "office6", "wps.exe")
            if os.path.isfile(wps_exe):
                candidates.append(wps_exe)

    if candidates:
        candidates.sort(reverse=True)
        return candidates[0]

    # 搜索 PATH
    for name in ("wps.exe", "wps"):
        path = shutil.which(name)
        if path:
            return path

    return "wps.exe"


def find_wps(app_type: str = "writer"):
    """获取 WPS COM Application 对象。

    Args:
        app_type: 应用类型 —— "writer" / "calc" / "impress"

    Returns:
        win32com COM Dispatch 对象

    Raises:
        RuntimeError: 无法创建 WPS COM 对象（WPS 未安装或版本不支持）
        ValueError: 不支持的应用类型
    """
    progid = PROGID_MAP.get(app_type.lower())
    if not progid:
        raise ValueError(
            f"不支持的应用类型: {app_type}。有效值: {', '.join(sorted(set(PROGID_MAP.keys())))}"
        )

    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch(progid)
        return app
    except ImportError:
        raise RuntimeError(
            "缺少 pywin32 库。请运行: pip install pywin32"
        )
    except Exception as e:
        raise RuntimeError(
            f"无法创建 {app_type} COM 对象 ({progid})。"
            f"请确认 WPS Office 已正确安装。错误: {e}"
        )


def get_version(app=None) -> str:
    """获取 WPS 版本号。

    Args:
        app: WPS COM Application 对象（可选，如未提供则创建临时对象）

    Returns:
        版本号字符串（如 "12.0"）
    """
    if app is None:
        app = find_wps("writer")
        try:
            return app.Version
        finally:
            quit_app(app)
    return app.Version


def create_document(app, doc_type: str = "writer"):
    """在 WPS 中创建新文档。

    Args:
        app: WPS COM Application 对象
        doc_type: 文档类型 —— "writer" / "calc" / "impress"

    Returns:
        COM Document 对象
    """
    if doc_type in ("writer", "wps"):
        return app.Documents.Add()
    elif doc_type in ("calc", "et"):
        return app.Workbooks.Add()
    elif doc_type in ("impress", "wpp"):
        return app.Presentations.Add()
    else:
        raise ValueError(f"不支持的文档类型: {doc_type}")


def open_document(app, path: str):
    """打开现有文档。

    Args:
        app: WPS COM Application 对象
        path: 文件路径

    Returns:
        COM Document 对象
    """
    abs_path = os.path.abspath(path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"文件不存在: {abs_path}")

    ext = os.path.splitext(abs_path)[1].lower()
    if ext in (".doc", ".docx", ".wps", ".wpt", ".rtf", ".txt", ".dot", ".dotx"):
        return app.Documents.Open(abs_path)
    elif ext in (".xls", ".xlsx", ".et", ".csv", ".xlt", ".xltx"):
        return app.Workbooks.Open(abs_path)
    elif ext in (".ppt", ".pptx", ".dps", ".dpt", ".pot", ".potx"):
        return app.Presentations.Open(abs_path)
    else:
        raise ValueError(f"不支持的文件格式: {ext}")


def save_as(doc, path: str, doc_type: str = "writer", format_name: str = None):
    """另存为指定格式。

    Args:
        doc: COM Document 对象
        path: 输出路径
        doc_type: 文档类型
        format_name: 格式名称（docx/pdf/xlsx/pptx/csv 等），
                     为 None 时根据扩展名自动判断

    Returns:
        保存的绝对路径
    """
    abs_path = os.path.abspath(path)
    os.makedirs(os.path.dirname(abs_path), exist_ok=True)

    if format_name is None:
        ext = os.path.splitext(abs_path)[1].lower().lstrip(".")
        format_name = ext

    formats = FORMAT_SAVEAS_MAP.get(doc_type, {})
    fmt_const = formats.get(format_name)
    if fmt_const is None:
        # 格式不在映射表中，尝试直接保存
        doc.SaveAs2(abs_path)
        return abs_path

    doc.SaveAs2(abs_path, FileFormat=fmt_const)
    return abs_path


def export_pdf(doc, output_path: str, doc_type: str = "writer"):
    """导出为 PDF。

    Args:
        doc: COM Document 对象
        output_path: 输出 PDF 路径
        doc_type: 文档类型

    Returns:
        输出的绝对路径
    """
    abs_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(abs_path), exist_ok=True)

    if doc_type == "impress":
        doc.SaveAs(abs_path, ppSaveAsPDF)
    else:
        doc.ExportAsFixedFormat(abs_path, 17)  # 17 = wdFormatPDF
    return abs_path


def close_document(doc, save: bool = False):
    """关闭文档。

    Args:
        doc: COM Document 对象
        save: 是否在关闭前保存
    """
    try:
        if save:
            doc.Save()
        doc.Close()
    except Exception:
        pass  # 忽略已关闭的文档错误


def quit_app(app, force: bool = False):
    """退出 WPS 应用程序。

    Args:
        app: WPS COM Application 对象
        force: 是否强制退出（不提示保存）
    """
    try:
        if force:
            app.DisplayAlerts = False
        app.Quit()
    except Exception:
        pass  # 忽略已退出的错误


def kill_all_wps_processes():
    """终止所有 WPS 相关后台进程（用于清理残留）。"""
    wps_names = ["wps.exe", "et.exe", "wpp.exe", "wpscloudsvr.exe"]
    for name in wps_names:
        try:
            subprocess.run(
                ["taskkill", "/F", "/IM", name, "/T"],
                capture_output=True, text=True,
            )
        except Exception:
            pass


def is_wps_running() -> bool:
    """检查是否有 WPS 进程在运行。"""
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq wps.exe"],
            capture_output=True, text=True,
        )
        return "wps.exe" in result.stdout.lower()
    except Exception:
        return False


def get_doc_type_from_ext(path: str) -> str:
    """根据文件扩展名判断文档类型。

    Returns:
        "writer" / "calc" / "impress"
    """
    ext = os.path.splitext(path)[1].lower()
    if ext in (".doc", ".docx", ".wps", ".rtf", ".txt", ".odt"):
        return "writer"
    elif ext in (".xls", ".xlsx", ".et", ".csv"):
        return "calc"
    elif ext in (".ppt", ".pptx", ".dps"):
        return "impress"
    return "writer"


def read_document_content(doc, doc_type: str = "writer") -> str:
    """读取文档全部文字内容。

    Args:
        doc: COM Document 对象
        doc_type: 文档类型

    Returns:
        文档的全部文字内容
    """
    if doc_type == "writer":
        return doc.Content.Text
    elif doc_type == "calc":
        # 读取所有工作表的全部已使用范围
        texts = []
        for ws in doc.Worksheets:
            try:
                used = ws.UsedRange
                texts.append(str(used.Value))
            except Exception:
                pass
        return "\n".join(texts)
    elif doc_type == "impress":
        texts = []
        for slide in doc.Slides:
            for shape in slide.Shapes:
                if hasattr(shape, "TextFrame") and shape.TextFrame.HasText:
                    texts.append(shape.TextFrame.TextRange.Text)
        return "\n".join(texts)
    return ""


def get_document_properties(doc) -> dict:
    """提取文档属性。

    Returns:
        包含文档基本属性的字典
    """
    info = {}
    try:
        info["name"] = doc.Name
    except Exception:
        pass
    try:
        info["path"] = doc.FullName
    except Exception:
        pass
    try:
        info["saved"] = doc.Saved
    except Exception:
        pass
    return info
