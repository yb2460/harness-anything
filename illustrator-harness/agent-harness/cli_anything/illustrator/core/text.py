# -*- coding: utf-8 -*-
"""Illustrator text operations."""


def add_text(app, text: str, x: float = 100, y: float = 100,
             font_size: float = 24, font_name: str = "Arial",
             color_rgb: tuple = (0, 0, 0)) -> dict:
    """Add a text frame to the active document."""
    doc = app.ActiveDocument
    js = f"""
    var doc = app.activeDocument;
    var tf = doc.textFrames.add();
    tf.contents = {repr(text)};
    tf.position = [{x}, {y}];
    tf.textRange.characterAttributes.size = {font_size};
    tf.textRange.characterAttributes.textFont = textFonts.getByName("{font_name}");
    var c = new RGBColor();
    c.red = {color_rgb[0]};
    c.green = {color_rgb[1]};
    c.blue = {color_rgb[2]};
    tf.textRange.characterAttributes.fillColor = c;
    """
    try:
        result = app.DoJavaScript(js)
        return {
            "text": text[:80],
            "position": [x, y],
            "font_size": font_size,
            "font_name": font_name,
        }
    except Exception as e:
        # Fallback: try COM method
        try:
            tf = doc.TextFrames.Add()
            tf.Contents = text
            return {"text": text[:80], "method": "COM"}
        except Exception:
            raise RuntimeError(f"Failed to add text: {e}")


def add_point_text(app, text: str, x: float = 100, y: float = 100,
                   font_size: float = 24, font_name: str = "ArialMT") -> dict:
    """Add point text (not area text) at a specific position."""
    js = f"""
    var doc = app.activeDocument;
    var pt = doc.textFrames.pointText([{x}, {y}]);
    pt.contents = {repr(text)};
    pt.textRange.characterAttributes.size = {font_size};
    """
    app.DoJavaScript(js)
    return {"text": text[:80], "position": [x, y], "type": "point_text"}


def add_area_text(app, text: str, x: float = 100, y: float = 100,
                  w: float = 300, h: float = 200, font_size: float = 14) -> dict:
    """Add area text within a rectangle boundary."""
    js = f"""
    var doc = app.activeDocument;
    var rect = doc.pathItems.rectangle({y + h}, {x}, {w}, {h});
    var at = doc.textFrames.areaText(rect);
    at.contents = {repr(text)};
    at.textRange.characterAttributes.size = {font_size};
    """
    app.DoJavaScript(js)
    return {"text": text[:80], "bounds": [x, y, w, h], "type": "area_text"}


def list_text_frames(app) -> list[dict]:
    """List all text frames in the active document."""
    doc = app.ActiveDocument
    frames = []
    for i in range(1, doc.TextFrames.Count + 1):
        tf = doc.TextFrames[i]
        frames.append({
            "index": i,
            "contents": tf.Contents[:100] if tf.Contents else "",
            "char_count": len(tf.Contents) if tf.Contents else 0,
        })
    return frames
