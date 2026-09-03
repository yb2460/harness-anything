# -*- coding: utf-8 -*-
"""Illustrator layer operations."""


def list_layers(app) -> list[dict]:
    """List all layers in the active document."""
    doc = app.ActiveDocument
    layers = []
    for i in range(1, doc.Layers.Count + 1):
        layer = doc.Layers[i]
        layers.append({
            "index": i,
            "name": layer.Name,
            "visible": layer.Visible,
            "locked": layer.Locked,
            "printable": layer.Printable,
            "path_items": layer.PathItems.Count,
            "text_frames": layer.TextFrames.Count,
        })
    return layers


def add_layer(app, name: str, above_index: int = -1) -> dict:
    """Add a new layer."""
    doc = app.ActiveDocument
    layer = doc.Layers.Add()
    layer.Name = name
    return {"name": layer.Name, "index": doc.Layers.Count}


def get_layer(app, index: int) -> dict:
    """Get a specific layer by index."""
    doc = app.ActiveDocument
    if index < 1 or index > doc.Layers.Count:
        raise ValueError(f"Layer index out of range: {index}")
    layer = doc.Layers[index]
    return {
        "index": index,
        "name": layer.Name,
        "visible": layer.Visible,
        "locked": layer.Locked,
    }


def remove_layer(app, index: int) -> dict:
    """Remove a layer by index."""
    doc = app.ActiveDocument
    if index < 1 or index > doc.Layers.Count:
        raise ValueError(f"Layer index out of range: {index}")
    name = doc.Layers[index].Name
    doc.Layers[index].Remove()
    return {"removed": name, "remaining": doc.Layers.Count}


def set_layer_visibility(app, index: int, visible: bool) -> dict:
    """Set layer visibility."""
    doc = app.ActiveDocument
    layer = doc.Layers[index]
    layer.Visible = visible
    return {"index": index, "name": layer.Name, "visible": visible}
