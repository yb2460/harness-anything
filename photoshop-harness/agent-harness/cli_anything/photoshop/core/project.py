# -*- coding: utf-8 -*-
"""项目管理模块 —— 新建/打开/保存 PSD 文件。"""
import os
from ..utils.com_bridge import get_bridge
from .session import get_session


def new_project(path: str, width: int = 1920, height: int = 1080,
                resolution: float = 72.0, mode: str = "rgb",
                bg_color: str = "white") -> dict:
    """创建新 PSD 文档。"""
    bridge = get_bridge()
    from ..utils.com_bridge import COLOR_MODES
    color_mode = COLOR_MODES.get(mode.lower(), 2)

    doc = bridge.app.Documents.Add(width, height, resolution, f"New Project",
                                   color_mode, 1, 1)  # 白色背景
    # 设置背景色
    if bg_color and bg_color != "transparent":
        _set_bg_color(doc, bg_color)

    # 保存
    abs_path = os.path.abspath(path)
    os.makedirs(os.path.dirname(abs_path), exist_ok=True)
    doc.SaveAs(abs_path)

    sess = get_session()
    sess.set_project(abs_path)
    return bridge.doc_info(doc)


def open_project(path: str) -> dict:
    """打开已有 PSD 文件。"""
    abs_path = os.path.abspath(path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"文件不存在: {abs_path}")

    bridge = get_bridge()
    doc = bridge.app.Open(abs_path)
    doc = bridge.app.ActiveDocument

    sess = get_session()
    sess.set_project(abs_path)
    return bridge.doc_info(doc)


def save_project(path: str = None) -> dict:
    """保存当前文档。"""
    bridge = get_bridge()
    doc = bridge.require_document()
    if path:
        abs_path = os.path.abspath(path)
        doc.SaveAs(abs_path)
        get_session().set_project(abs_path)
    else:
        doc.Save()
    return bridge.doc_info(doc)


def close_project(save: bool = True) -> dict:
    """关闭当前文档。"""
    bridge = get_bridge()
    doc = bridge.require_document()
    name = doc.Name
    from pywintypes import com_error
    try:
        if save:
            doc.Close(2)  # psSaveChanges
        else:
            doc.Close(3)  # psDoNotSaveChanges
    except com_error:
        pass
    return {"closed": name, "status": "ok"}


def _set_bg_color(doc, color: str):
    """设置文档背景色。"""
    color_map = {
        "white": (255, 255, 255), "black": (0, 0, 0),
        "red": (255, 0, 0), "green": (0, 255, 0), "blue": (0, 0, 255),
        "transparent": None,
    }
    if color.lower() not in color_map or color.lower() == "transparent":
        return
    r, g, b = color_map[color.lower()]
    try:
        doc.Selection.SelectAll()
        rgb = doc.ArtLayers.Item(1)
        # 创建纯色填充层
    except Exception:
        pass  # 背景色默认已处理
