# -*- coding: utf-8 -*-
"""Illustrator document management: create, open, save, info."""

import os
import pythoncom
import win32com.client


def _get_app():
    """Get or create Illustrator COM application instance."""
    pythoncom.CoInitialize()
    try:
        app = win32com.client.Dispatch("Illustrator.Application")
        return app
    except Exception:
        raise RuntimeError(
            "Adobe Illustrator not found. Install Illustrator first."
        )


def create_document(width: float = 612.0, height: float = 792.0,
                    color_mode: str = "RGB") -> dict:
    """Create a new Illustrator document.

    Args:
        width: Artboard width in points (default 612 = 8.5 inches)
        height: Artboard height in points (default 792 = 11 inches)
        color_mode: "RGB" or "CMYK"
    """
    app = _get_app()
    doc = app.Documents.Add()
    # Set artboard size via JavaScript
    js = f"""
    var doc = app.activeDocument;
    doc.artboards[0].artboardRect = [0, {height}, {width}, 0];
    doc.documentColorSpace = DocumentColorSpace.{color_mode};
    """
    try:
        app.DoJavaScript(js)
    except Exception:
        pass

    return {
        "name": doc.Name,
        "width": width,
        "height": height,
        "color_mode": color_mode,
        "path": "",
    }


def open_document(path: str) -> dict:
    """Open an existing Illustrator document."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"File not found: {path}")

    app = _get_app()
    doc = app.Open(os.path.abspath(path))

    return {
        "name": doc.Name,
        "full_path": doc.FullName if hasattr(doc, "FullName") else path,
        "width": doc.Width if hasattr(doc, "Width") else 0,
        "height": doc.Height if hasattr(doc, "Height") else 0,
    }


def save_document(app, path: str) -> dict:
    """Save the active document."""
    doc = app.ActiveDocument
    abs_path = os.path.abspath(path)
    os.makedirs(os.path.dirname(abs_path) if os.path.dirname(abs_path) else ".", exist_ok=True)
    doc.SaveAs(abs_path)
    size = os.path.getsize(abs_path)
    return {
        "file": abs_path,
        "name": doc.Name,
        "size_bytes": size,
    }


def get_document_info(app) -> dict:
    """Get information about the active document."""
    doc = app.ActiveDocument
    return {
        "name": doc.Name,
        "full_path": doc.FullName if hasattr(doc, "FullName") else "",
        "width": float(doc.Width) if hasattr(doc, "Width") else 0,
        "height": float(doc.Height) if hasattr(doc, "Height") else 0,
        "layer_count": doc.Layers.Count,
        "text_frame_count": doc.TextFrames.Count,
        "path_item_count": doc.PathItems.Count,
        "artboard_count": doc.Artboards.Count,
        "swatch_count": doc.Swatches.Count,
        "selection_count": len(doc.Selection) if doc.Selection else 0,
    }
