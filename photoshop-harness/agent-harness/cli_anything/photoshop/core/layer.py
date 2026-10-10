# -*- coding: utf-8 -*-
"""图层操作模块 —— 图层的增删改、排序、可见性、透明度、混合模式。"""
from ..utils.com_bridge import get_bridge, BLEND_MODES
from .session import get_session


def add_layer(name: str, layer_type: str = "pixel", opacity: int = 100,
              blend_mode: str = "normal", group: str = None) -> dict:
    """添加新图层。"""
    bridge = get_bridge()
    doc = bridge.require_document()
    sess = get_session()
    sess.snapshot()

    from ..utils.com_bridge import LAYER_TYPES
    lt = LAYER_TYPES.get(layer_type, 1)

    if lt == 1:  # pixel layer
        layer = doc.ArtLayers.Add()
    elif lt == 2:  # text layer (handled by text module)
        layer = doc.ArtLayers.Add()
        layer.Kind = 2  # psTextLayer
    else:
        layer = doc.ArtLayers.Add()

    layer.Name = name
    layer.Opacity = opacity

    # 设置混合模式
    bm = blend_mode.lower()
    if bm in BLEND_MODES:
        try:
            blend_const = getattr(
                __import__("win32com.client", fromlist=["constants"]).constants,
                f"ps{bm[0].upper()}{bm[1:]}"
            )
            layer.Mode = blend_const
        except Exception:
            pass  # 保持默认模式

    # 移动到图层组
    if group:
        target_group = bridge.get_layer(group, doc)
        if target_group:
            layer.Move(target_group, 3)  # psPlaceInside

    return bridge._layer_info(layer)


def delete_layer(name: str) -> dict:
    """删除指定图层。"""
    bridge = get_bridge()
    sess = get_session()
    sess.snapshot()

    layer = bridge.get_layer(name)
    if layer is None:
        existing = [l["name"] for l in bridge.list_layers()]
        raise ValueError(f"图层 '{name}' 不存在。当前图层: {existing}")
    layer.Remove()
    return {"deleted": name, "status": "ok"}


def duplicate_layer(name: str, new_name: str = None) -> dict:
    """复制图层。"""
    bridge = get_bridge()
    sess = get_session()
    sess.snapshot()

    layer = bridge.get_layer(name)
    if layer is None:
        raise ValueError(f"图层 '{name}' 不存在")
    new_layer = layer.Duplicate()
    if new_name:
        new_layer.Name = new_name
    return bridge._layer_info(new_layer)


def rename_layer(old_name: str, new_name: str) -> dict:
    """重命名图层。"""
    bridge = get_bridge()
    layer = bridge.get_layer(old_name)
    if layer is None:
        raise ValueError(f"图层 '{old_name}' 不存在")
    layer.Name = new_name
    return bridge._layer_info(layer)


def set_visibility(name: str, visible: bool) -> dict:
    """设置图层可见性。"""
    bridge = get_bridge()
    layer = bridge.get_layer(name)
    if layer is None:
        raise ValueError(f"图层 '{name}' 不存在")
    layer.Visible = visible
    return bridge._layer_info(layer)


def set_opacity(name: str, opacity: int) -> dict:
    """设置图层透明度 (0-100)。"""
    bridge = get_bridge()
    layer = bridge.get_layer(name)
    if layer is None:
        raise ValueError(f"图层 '{name}' 不存在")
    layer.Opacity = max(0, min(100, opacity))
    return bridge._layer_info(layer)


def set_blend_mode(name: str, blend_mode: str) -> dict:
    """设置图层混合模式。"""
    bridge = get_bridge()
    layer = bridge.get_layer(name)
    if layer is None:
        raise ValueError(f"图层 '{name}' 不存在")
    if blend_mode.lower() not in BLEND_MODES:
        raise ValueError(f"无效混合模式: {blend_mode}，可选: {BLEND_MODES}")
    layer.Mode = blend_mode
    return bridge._layer_info(layer)


def reorder(name: str, position: int) -> list[dict]:
    """移动图层到指定位置（1=最底层）。"""
    bridge = get_bridge()
    sess = get_session()
    sess.snapshot()

    layer = bridge.get_layer(name)
    if layer is None:
        raise ValueError(f"图层 '{name}' 不存在")
    doc = bridge.require_document()
    layer.Move(doc, 2)  # psPlaceAtBeginning
    # 然后移动到目标位置
    if position > 1:
        for _ in range(position - 1):
            try:
                layer.Move(doc.ArtLayers.Item(1), 4)  # psPlaceAfter
            except Exception:
                break
    return bridge.list_layers()


def list_layers() -> list[dict]:
    """列出所有图层。"""
    return get_bridge().list_layers()


def merge_layers(name1: str, name2: str = None) -> dict:
    """合并图层。"""
    bridge = get_bridge()
    sess = get_session()
    sess.snapshot()

    layer1 = bridge.get_layer(name1)
    if layer1 is None:
        raise ValueError(f"图层 '{name1}' 不存在")
    if name2:
        layer2 = bridge.get_layer(name2)
        if layer2 is None:
            raise ValueError(f"图层 '{name2}' 不存在")
        layer1.Merge()
    else:
        # 向下合并
        layer1.Merge()
    return {"merged": name1, "status": "ok"}
