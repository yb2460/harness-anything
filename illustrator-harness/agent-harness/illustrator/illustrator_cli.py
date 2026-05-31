# -*- coding: utf-8 -*-
"""Illustrator CLI harness -- command-line control of Adobe Illustrator via COM.

Usage:
    cli-anything-illustrator project new
    cli-anything-illustrator text add "Hello" --x 100 --y 100
    cli-anything-illustrator shape rect --w 200 --h 200
    cli-anything-illustrator export png output.png
    cli-anything-illustrator                         # Interactive REPL
"""

import os
import sys
import json
import click
import pythoncom

from cli_anything.illustrator.utils.ai_backend import detect_illustrator, launch_illustrator
from cli_anything.illustrator.core import project, layers, text, shapes, export as export_mod
from cli_anything.illustrator.utils.repl_skin import ReplSkin

_JSON_HELP = "JSON output (machine-readable)"


def _output(ctx, data):
    if ctx.obj.get("json_output"):
        click.echo(json.dumps(data, ensure_ascii=False, indent=2, default=str))
    elif isinstance(data, dict):
        for k, v in data.items():
            click.echo(f"  {k}: {v}")
    elif isinstance(data, list):
        for item in data:
            click.echo(f"  {item}")

def _get_app(ctx):
    if "app" not in ctx.obj:
        ctx.obj["app"] = launch_illustrator(visible=ctx.obj.get("visible", True))
    return ctx.obj["app"]

# ═══════════════════════════════════════════════════════════
@click.group(invoke_without_command=True)
@click.option("--json", "json_output", is_flag=True, help=_JSON_HELP)
@click.option("--visible/--no-visible", default=True, help="Show/hide Illustrator window")
@click.pass_context
def cli(ctx, json_output, visible):
    """cli-anything-illustrator -- CLI harness for Adobe Illustrator."""
    ctx.ensure_object(dict)
    ctx.obj["json_output"] = json_output
    ctx.obj["visible"] = visible

    if ctx.invoked_subcommand is None:
        ctx.invoke(repl)

# ════════════════════════ project ════════════════════════
@cli.group()
def project():
    """Document management: new, open, save, info."""
    pass

@project.command("new")
@click.option("--width", type=float, default=612.0, help="Artboard width (pts)")
@click.option("--height", type=float, default=792.0, help="Artboard height (pts)")
@click.option("--color-mode", default="RGB", type=click.Choice(["RGB", "CMYK"]))
@click.pass_context
def project_new(ctx, width, height, color_mode):
    """Create a new document."""
    result = project.create_document(width, height, color_mode)
    ctx.obj["app"] = launch_illustrator(visible=ctx.obj.get("visible", True))
    _output(ctx, result)

@project.command("open")
@click.argument("path")
@click.pass_context
def project_open(ctx, path):
    """Open an existing document."""
    result = project.open_document(path)
    ctx.obj["app"] = launch_illustrator(visible=ctx.obj.get("visible", True))
    _output(ctx, result)

@project.command("save")
@click.argument("path")
@click.pass_context
def project_save(ctx, path):
    """Save the active document."""
    app = _get_app(ctx)
    result = project.save_document(app, path)
    _output(ctx, result)

@project.command("info")
@click.pass_context
def project_info(ctx):
    """Show active document information."""
    app = _get_app(ctx)
    result = project.get_document_info(app)
    _output(ctx, result)

# ════════════════════════ layer ════════════════════════
@cli.group()
def layer():
    """Layer management: list, add, remove, toggle."""
    pass

@layer.command("list")
@click.pass_context
def layer_list(ctx):
    """List all layers."""
    app = _get_app(ctx)
    result = layers.list_layers(app)
    _output(ctx, result)

@layer.command("add")
@click.argument("name")
@click.pass_context
def layer_add(ctx, name):
    """Add a new layer."""
    app = _get_app(ctx)
    result = layers.add_layer(app, name)
    _output(ctx, result)

@layer.command("remove")
@click.argument("index", type=int)
@click.pass_context
def layer_remove(ctx, index):
    """Remove a layer by index."""
    app = _get_app(ctx)
    result = layers.remove_layer(app, index)
    _output(ctx, result)

@layer.command("hide")
@click.argument("index", type=int)
@click.pass_context
def layer_hide(ctx, index):
    """Hide a layer."""
    app = _get_app(ctx)
    result = layers.set_layer_visibility(app, index, False)
    _output(ctx, result)

@layer.command("show")
@click.argument("index", type=int)
@click.pass_context
def layer_show(ctx, index):
    """Show a layer."""
    app = _get_app(ctx)
    result = layers.set_layer_visibility(app, index, True)
    _output(ctx, result)

# ════════════════════════ text ════════════════════════
@cli.group()
def text():
    """Text operations: add, list."""
    pass

@text.command("add")
@click.argument("content")
@click.option("--x", type=float, default=100.0)
@click.option("--y", type=float, default=100.0)
@click.option("--font-size", type=float, default=24.0)
@click.option("--font", "font_name", default="Arial")
@click.option("--r", type=int, default=0)
@click.option("--g", type=int, default=0)
@click.option("--b", type=int, default=0)
@click.pass_context
def text_add(ctx, content, x, y, font_size, font_name, r, g, b):
    """Add a text frame."""
    app = _get_app(ctx)
    result = text.add_text(app, content, x, y, font_size, font_name, (r, g, b))
    _output(ctx, result)

@text.command("list")
@click.pass_context
def text_list(ctx):
    """List all text frames."""
    app = _get_app(ctx)
    result = text.list_text_frames(app)
    _output(ctx, result)

# ════════════════════════ shape ════════════════════════
@cli.group()
def shape():
    """Shape operations: rect, ellipse, circle, polygon, star, list."""
    pass

@shape.command("rect")
@click.option("--x", type=float, default=100.0)
@click.option("--y", type=float, default=100.0)
@click.option("--w", type=float, default=200.0)
@click.option("--h", type=float, default=200.0)
@click.option("--fill-r", type=int, default=128)
@click.option("--fill-g", type=int, default=128)
@click.option("--fill-b", type=int, default=128)
@click.pass_context
def shape_rect(ctx, x, y, w, h, fill_r, fill_g, fill_b):
    """Add a rectangle."""
    app = _get_app(ctx)
    result = shapes.add_rectangle(app, x, y, w, h, (fill_r, fill_g, fill_b))
    _output(ctx, result)

@shape.command("ellipse")
@click.option("--cx", type=float, default=200.0)
@click.option("--cy", type=float, default=200.0)
@click.option("--rx", type=float, default=100.0)
@click.option("--ry", type=float, default=80.0)
@click.option("--fill-r", type=int, default=128)
@click.option("--fill-g", type=int, default=128)
@click.option("--fill-b", type=int, default=128)
@click.pass_context
def shape_ellipse(ctx, cx, cy, rx, ry, fill_r, fill_g, fill_b):
    """Add an ellipse."""
    app = _get_app(ctx)
    result = shapes.add_ellipse(app, cx, cy, rx, ry, (fill_r, fill_g, fill_b))
    _output(ctx, result)

@shape.command("circle")
@click.option("--cx", type=float, default=200.0)
@click.option("--cy", type=float, default=200.0)
@click.option("--r", "radius", type=float, default=100.0)
@click.option("--fill-r", type=int, default=128)
@click.option("--fill-g", type=int, default=128)
@click.option("--fill-b", type=int, default=128)
@click.pass_context
def shape_circle(ctx, cx, cy, radius, fill_r, fill_g, fill_b):
    """Add a circle."""
    app = _get_app(ctx)
    result = shapes.add_circle(app, cx, cy, radius, (fill_r, fill_g, fill_b))
    _output(ctx, result)

@shape.command("polygon")
@click.option("--cx", type=float, default=200.0)
@click.option("--cy", type=float, default=200.0)
@click.option("--r", "radius", type=float, default=100.0)
@click.option("--sides", type=int, default=6)
@click.option("--fill-r", type=int, default=128)
@click.option("--fill-g", type=int, default=128)
@click.option("--fill-b", type=int, default=128)
@click.pass_context
def shape_polygon(ctx, cx, cy, radius, sides, fill_r, fill_g, fill_b):
    """Add a polygon."""
    app = _get_app(ctx)
    result = shapes.add_polygon(app, cx, cy, radius, sides, (fill_r, fill_g, fill_b))
    _output(ctx, result)

@shape.command("star")
@click.option("--cx", type=float, default=200.0)
@click.option("--cy", type=float, default=200.0)
@click.option("--outer", type=float, default=100.0)
@click.option("--inner", type=float, default=50.0)
@click.option("--points", type=int, default=5)
@click.pass_context
def shape_star(ctx, cx, cy, outer, inner, points):
    """Add a star."""
    app = _get_app(ctx)
    result = shapes.add_star(app, cx, cy, outer, inner, points)
    _output(ctx, result)

@shape.command("list")
@click.pass_context
def shape_list(ctx):
    """List all path items."""
    app = _get_app(ctx)
    result = shapes.list_path_items(app)
    _output(ctx, result)

# ════════════════════════ export ════════════════════════
@cli.group()
def export():
    """Export operations: png, svg, pdf, jpg."""
    pass

@export.command("png")
@click.argument("output_path")
@click.pass_context
def export_png(ctx, output_path):
    """Export as PNG."""
    app = _get_app(ctx)
    result = export_mod.export_as(app, output_path, "PNG")
    _output(ctx, result)

@export.command("svg")
@click.argument("output_path")
@click.pass_context
def export_svg(ctx, output_path):
    """Export as SVG."""
    app = _get_app(ctx)
    result = export_mod.export_as(app, output_path, "SVG")
    _output(ctx, result)

@export.command("pdf")
@click.argument("output_path")
@click.pass_context
def export_pdf(ctx, output_path):
    """Export as PDF."""
    app = _get_app(ctx)
    result = export_mod.export_as(app, output_path, "PDF")
    _output(ctx, result)

@export.command("ai")
@click.argument("output_path")
@click.pass_context
def export_ai(ctx, output_path):
    """Save as .ai file."""
    app = _get_app(ctx)
    result = export_mod.save_as_ai(app, output_path)
    _output(ctx, result)

# ════════════════════════ backend ════════════════════════
@cli.group()
def backend():
    """Backend utilities."""
    pass

@backend.command("detect")
@click.pass_context
def backend_detect(ctx):
    """Detect Illustrator installation."""
    result = detect_illustrator()
    _output(ctx, result)

@backend.command("launch")
@click.pass_context
def backend_launch(ctx):
    """Launch Illustrator."""
    app = launch_illustrator(visible=True)
    ctx.obj["app"] = app
    click.echo(f"Illustrator launched: v{app.Version}")

# ════════════════════════ REPL ════════════════════════
@cli.command("repl")
@click.pass_context
def repl(ctx):
    """Interactive REPL mode."""
    skin = ReplSkin("illustrator")
    skin.print_banner()

    try:
        info = detect_illustrator()
        if info["com_available"]:
            skin.success(f"Illustrator v{info['version']} detected: {info['path']}")
        else:
            skin.error("Illustrator not found. Install Adobe Illustrator 2020+.")
            return
    except Exception:
        pass

    cmds = {
        "project new": "Create new document",
        "project open <path>": "Open document",
        "project save <path>": "Save document",
        "project info": "Show document info",
        "layer list": "List layers",
        "layer add <name>": "Add layer",
        "text add <text>": "Add text frame",
        "shape rect --w 200 --h 200": "Add rectangle",
        "shape circle --r 100": "Add circle",
        "shape star": "Add star",
        "export png <path>": "Export PNG",
        "export svg <path>": "Export SVG",
        "backend detect": "Check installation",
        "help": "Show this help",
        "quit": "Exit REPL",
    }
    skin.help(cmds)

    while True:
        try:
            ui = input(skin.prompt()).strip()
            if not ui: continue
            if ui in ("quit", "exit", "q"):
                click.echo("Goodbye!")
                break
            if ui == "help":
                skin.help(cmds)
                continue
            try:
                cli(ui.split(), standalone_mode=False)
            except SystemExit:
                pass
            except Exception as e:
                skin.error(str(e))
        except KeyboardInterrupt:
            click.echo(); continue
        except EOFError:
            break


if __name__ == "__main__":
    cli()
