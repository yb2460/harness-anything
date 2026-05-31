# -*- coding: utf-8 -*-
"""Illustrator export operations."""

import os


def export_as(app, output_path: str, file_type: str = "PNG",
              options: dict = None) -> dict:
    """Export the active document to various formats.

    Args:
        app: Illustrator COM application
        output_path: Full output file path
        file_type: One of "PNG", "JPG", "SVG", "PDF", "TIFF", "EPS"
        options: Additional export options
    """
    doc = app.ActiveDocument
    abs_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(abs_path) if os.path.dirname(abs_path) else ".", exist_ok=True)

    # Illustrator export type constants
    export_types = {
        "PNG": 5,
        "JPG": 1,
        "SVG": 2,
        "PDF": 3,
        "TIFF": 4,
        "EPS": 6,
    }

    export_type = export_types.get(file_type.upper(), 5)

    export_opts = {}
    if options:
        export_opts.update(options)

    try:
        doc.Export(abs_path, export_type, export_opts)
    except Exception:
        # Fallback: JavaScript export
        js = f"""
        var doc = app.activeDocument;
        var file = new File("{abs_path.replace(chr(92), '/')}");
        var opts = new ExportOptions{file_type.upper()}();
        doc.exportFile(file, ExportType.{file_type.upper()}, opts);
        """
        try:
            app.DoJavaScript(js)
        except Exception as e:
            raise RuntimeError(f"Export failed: {e}")

    size = os.path.getsize(abs_path)
    return {
        "output": abs_path,
        "format": file_type.upper(),
        "size_bytes": size,
    }


def save_as_ai(app, output_path: str) -> dict:
    """Save as .ai file."""
    doc = app.ActiveDocument
    abs_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(abs_path) if os.path.dirname(abs_path) else ".", exist_ok=True)
    doc.SaveAs(abs_path)
    return {
        "output": abs_path,
        "format": "AI",
        "size_bytes": os.path.getsize(abs_path),
    }


def export_artboards(app, output_dir: str, file_type: str = "PNG") -> list[dict]:
    """Export each artboard as a separate file."""
    results = []
    doc = app.ActiveDocument
    name_base = os.path.splitext(doc.Name)[0] if doc.Name else "artboard"

    for i in range(1, doc.Artboards.Count + 1):
        doc.Artboards[i].Active = True
        output_path = os.path.join(output_dir, f"{name_base}_artboard_{i}.{file_type.lower()}")
        result = export_as(app, output_path, file_type)
        results.append(result)

    return results
