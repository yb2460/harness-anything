# -*- coding: utf-8 -*-
"""Illustrator shape/path operations."""


def _js_color(r, g, b):
    return f"var c = new RGBColor(); c.red = {r}; c.green = {g}; c.blue = {b};"


def add_rectangle(app, x: float = 100, y: float = 100,
                  w: float = 200, h: float = 200,
                  fill_rgb: tuple = (128, 128, 128),
                  stroke_rgb: tuple = (0, 0, 0),
                  stroke_width: float = 1.0) -> dict:
    """Add a rectangle to the active document."""
    js = f"""
    var doc = app.activeDocument;
    var rect = doc.pathItems.rectangle({y + h}, {x}, {w}, {h});
    {_js_color(*fill_rgb)}
    rect.fillColor = c;
    rect.filled = true;
    var sc = new RGBColor();
    sc.red = {stroke_rgb[0]}; sc.green = {stroke_rgb[1]}; sc.blue = {stroke_rgb[2]};
    rect.strokeColor = sc;
    rect.strokeWidth = {stroke_width};
    rect.name = "rect_{x}_{y}";
    """
    app.DoJavaScript(js)
    return {"type": "rectangle", "bounds": [x, y, w, h]}


def add_ellipse(app, cx: float = 200, cy: float = 200,
                rx: float = 100, ry: float = 80,
                fill_rgb: tuple = (128, 128, 128)) -> dict:
    """Add an ellipse to the active document."""
    js = f"""
    var doc = app.activeDocument;
    var el = doc.pathItems.ellipse({cy + ry}, {cx - rx}, {rx * 2}, {ry * 2});
    {_js_color(*fill_rgb)}
    el.fillColor = c;
    el.filled = true;
    """
    app.DoJavaScript(js)
    return {"type": "ellipse", "center": [cx, cy], "radii": [rx, ry]}


def add_circle(app, cx: float = 200, cy: float = 200,
               r: float = 100, fill_rgb: tuple = (128, 128, 128)) -> dict:
    """Add a circle. Convenience wrapper around add_ellipse."""
    return add_ellipse(app, cx, cy, r, r, fill_rgb)


def add_rounded_rectangle(app, x: float = 100, y: float = 100,
                          w: float = 200, h: float = 200,
                          corner_radius: float = 20,
                          fill_rgb: tuple = (128, 128, 128)) -> dict:
    """Add a rounded rectangle."""
    js = f"""
    var doc = app.activeDocument;
    var rect = doc.pathItems.roundedRectangle({y + h}, {x}, {w}, {h}, {corner_radius}, {corner_radius});
    {_js_color(*fill_rgb)}
    rect.fillColor = c;
    rect.filled = true;
    """
    app.DoJavaScript(js)
    return {"type": "rounded_rectangle", "bounds": [x, y, w, h], "radius": corner_radius}


def add_polygon(app, cx: float = 200, cy: float = 200,
                radius: float = 100, sides: int = 6,
                fill_rgb: tuple = (128, 128, 128)) -> dict:
    """Add a regular polygon."""
    js = f"""
    var doc = app.activeDocument;
    var poly = doc.pathItems.polygon({cy}, {cx}, {radius}, {sides});
    {_js_color(*fill_rgb)}
    poly.fillColor = c;
    poly.filled = true;
    """
    app.DoJavaScript(js)
    return {"type": "polygon", "center": [cx, cy], "sides": sides, "radius": radius}


def add_star(app, cx: float = 200, cy: float = 200,
             outer_radius: float = 100, inner_radius: float = 50,
             points: int = 5, fill_rgb: tuple = (255, 215, 0)) -> dict:
    """Add a star shape."""
    js = f"""
    var doc = app.activeDocument;
    var star = doc.pathItems.star({cy}, {cx}, {outer_radius}, {inner_radius}, {points});
    {_js_color(*fill_rgb)}
    star.fillColor = c;
    star.filled = true;
    """
    app.DoJavaScript(js)
    return {"type": "star", "center": [cx, cy], "points": points}


def list_path_items(app) -> list[dict]:
    """List all path items in the active document."""
    doc = app.ActiveDocument
    items = []
    for i in range(1, doc.PathItems.Count + 1):
        pi = doc.PathItems[i]
        items.append({
            "index": i,
            "name": pi.Name if pi.Name else f"Path_{i}",
            "filled": pi.Filled if hasattr(pi, "Filled") else False,
            "stroked": pi.Stroked if hasattr(pi, "Stroked") else False,
        })
    return items
