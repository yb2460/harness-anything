# -*- coding: utf-8 -*-
"""Photoshop COM 自动化桥接层。

封装所有 Photoshop COM API 调用，提供：
- 单例 Photoshop.Application 连接
- 自动启动/连接 Photoshop
- COM 错误统一处理
- 枚举常量映射
"""
import sys
import time
import pythoncom
import win32com.client
from win32com.client import constants as ps
from pywintypes import com_error
from typing import Optional, Any


# ── Photoshop COM 常量映射 ──────────────────────────────────────────────

# 文档色彩模式
COLOR_MODES = {
    "bitmap": 5, "grayscale": 1, "indexed": 4,
    "rgb": 2, "cmyk": 5, "lab": 7, "multichannel": 6,
}
COLOR_MODE_NAMES = {v: k.upper() for k, v in COLOR_MODES.items()}

# 图层混合模式
BLEND_MODES = [
    "normal", "dissolve", "darken", "multiply", "colorBurn", "linearBurn",
    "darkerColor", "lighten", "screen", "colorDodge", "linearDodge",
    "lighterColor", "overlay", "softLight", "hardLight", "vividLight",
    "linearLight", "pinLight", "hardMix", "difference", "exclusion",
    "subtract", "divide", "hue", "saturation", "color", "luminosity",
]

# 导出格式
EXPORT_FORMATS = {
    "png": 2,    # psPNGFormat
    "jpg": 1,    # psJPEG
    "gif": 4,    # psGIF
    "bmp": 5,    # psBMP
    "tiff": 3,   # psTIFF
    "pdf": 6,    # psPDF
}

# 文件保存格式
SAVE_FORMATS = {
    "psd": "Photoshop",
    "tiff": "TIFF",
    "jpg": "JPEG",
    "png": "PNG",
    "bmp": "BMP",
    "gif": "GIF",
}

# 单位
UNITS = {"pixels": 1, "inches": 2, "cm": 3, "mm": 4, "points": 5, "picas": 6}

# 锚点位置 (用于 resize)
ANCHORS = {
    "topLeft": 1, "topCenter": 2, "topRight": 3,
    "middleLeft": 4, "middleCenter": 5, "middleRight": 6,
    "bottomLeft": 7, "bottomCenter": 8, "bottomRight": 9,
}

# 图层类型
LAYER_TYPES = {"pixel": 1, "text": 2, "adjustment": 3, "smartObject": 4}


class ComBridge:
    """Photoshop COM 单例桥接器。"""

    _instance: Optional["ComBridge"] = None
    _app: Any = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    @property
    def app(self) -> Any:
        """获取或创建 Photoshop.Application 实例。"""
        if self._app is None:
            self._connect()
        return self._app

    def _connect(self):
        """连接或启动 Photoshop。"""
        pythoncom.CoInitialize()
        try:
            self._app = win32com.client.Dispatch("Photoshop.Application")
            self._app.Visible = True  # 确保可见
            # 本地补丁: 强制标尺单位为像素。CS6 在标尺单位为厘米时会把
            # Documents.Add(640, 400, 72) 的宽高当作厘米，生成 640x400 cm 文档，
            # 导出即 18142x11339 px 巨型位图。
            try:
                self._app.Preferences.RulerUnits = 1  # psPixels
            except Exception:
                pass
            time.sleep(0.5)  # 等待 COM 就绪
        except com_error as e:
            raise RuntimeError(
                f"无法连接 Photoshop。请确认 Photoshop 2023+ 已安装。\n"
                f"COM 错误: {e.strerror or e}"
            )

    @property
    def active_document(self):
        """获取当前活跃文档，无文档时返回 None。"""
        try:
            doc = self.app.ActiveDocument
            return doc if doc else None
        except com_error:
            return None

    def get_document(self, name: str = None):
        """按名称获取文档，不指定则返回活跃文档。"""
        if name:
            for i in range(1, self.app.Documents.Count + 1):
                doc = self.app.Documents.Item(i)
                if doc.Name == name:
                    return doc
            return None
        return self.active_document

    def require_document(self):
        """获取活跃文档，无文档时抛出异常。"""
        doc = self.active_document
        if doc is None:
            raise RuntimeError("没有打开的文档。请先用 project open 或 project new 创建项目。")
        return doc

    def get_layer(self, name: str, doc=None):
        """按名称查找图层。"""
        doc = doc or self.require_document()
        for i in range(1, doc.ArtLayers.Count + 1):
            layer = doc.ArtLayers.Item(i)
            if layer.Name == name:
                return layer
        # 也搜索图层组
        for i in range(1, doc.LayerSets.Count + 1):
            group = doc.LayerSets.Item(i)
            if group.Name == name:
                return group
            for j in range(1, group.ArtLayers.Count + 1):
                if group.ArtLayers.Item(j).Name == name:
                    return group.ArtLayers.Item(j)
        return None

    def list_layers(self, doc=None) -> list[dict]:
        """列出所有图层信息。"""
        doc = doc or self.require_document()
        layers = []
        for i in range(1, doc.ArtLayers.Count + 1):
            layer = doc.ArtLayers.Item(i)
            layers.append(self._layer_info(layer))
        for i in range(1, doc.LayerSets.Count + 1):
            group = doc.LayerSets.Item(i)
            info = self._layer_info(group)
            info["type"] = "group"
            layers.append(info)
        return layers

    def _layer_info(self, layer) -> dict:
        """提取图层属性字典。"""
        try:
            kind = "text" if layer.Kind and layer.Kind == 2 else "pixel"
        except (com_error, AttributeError):
            kind = "pixel"
        return {
            "name": layer.Name,
            "type": kind,
            "visible": bool(layer.Visible),
            "opacity": layer.Opacity,
            "blend_mode": str(layer.Mode) if hasattr(layer, "Mode") else "normal",
            "bounds": list(layer.Bounds) if hasattr(layer, "Bounds") else None,
        }

    def doc_info(self, doc=None) -> dict:
        """获取文档信息字典。"""
        doc = doc or self.require_document()
        try:
            mode_name = COLOR_MODE_NAMES.get(doc.Mode, str(doc.Mode))
        except com_error:
            mode_name = "unknown"
        return {
            "name": doc.Name,
            "path": doc.Path if hasattr(doc, "Path") else None,
            "width": int(doc.Width),
            "height": int(doc.Height),
            "resolution": round(doc.Resolution, 1),
            "mode": mode_name,
            "bits_per_channel": doc.BitsPerChannel if hasattr(doc, "BitsPerChannel") else 8,
            "layer_count": doc.ArtLayers.Count,
        }

    def dispose(self):
        """释放 COM 资源（不关闭 Photoshop）。"""
        try:
            self._app = None
            pythoncom.CoUninitialize()
        except Exception:
            pass


# ── 全局单例 ────────────────────────────────────────────────────────────

_bridge: Optional[ComBridge] = None


def get_bridge() -> ComBridge:
    """获取 COM Bridge 全局单例。"""
    global _bridge
    if _bridge is None:
        _bridge = ComBridge()
    return _bridge
