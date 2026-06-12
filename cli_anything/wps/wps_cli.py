#!/usr/bin/env python3
"""WPS Office CLI —— 通过命令行操控 WPS Office 文档。

用法:
    # 一次性命令
    cli-anything-wps document new --type writer --name "报告"
    cli-anything-wps writer add-paragraph --text "Hello WPS!"
    cli-anything-wps export render output.docx --preset docx

    # 交互式 REPL
    cli-anything-wps
"""

import sys
import os
import json
import shlex
import click
from typing import Optional

# 添加父级路径以供导入
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cli_anything.wps.core.session import Session
from cli_anything.wps.core import document as doc_mod
from cli_anything.wps.core import writer as writer_mod
from cli_anything.wps.core import calc as calc_mod
from cli_anything.wps.core import impress as impress_mod
from cli_anything.wps.core import styles as styles_mod
from cli_anything.wps.core import export as export_mod

# 全局会话状态
_session: Optional[Session] = None
_json_output = False
_repl_mode = False


def get_session() -> Session:
    global _session
    if _session is None:
        _session = Session()
    return _session


def output(data, message: str = ""):
    if _json_output:
        click.echo(json.dumps(data, indent=2, default=str))
    else:
        if message:
            click.echo(message)
        if isinstance(data, dict):
            _print_dict(data)
        elif isinstance(data, list):
            _print_list(data)
        else:
            click.echo(str(data))


def _print_dict(d: dict, indent: int = 0):
    prefix = "  " * indent
    for k, v in d.items():
        if isinstance(v, dict):
            click.echo(f"{prefix}{k}:")
            _print_dict(v, indent + 1)
        elif isinstance(v, list):
            click.echo(f"{prefix}{k}:")
            _print_list(v, indent + 1)
        else:
            click.echo(f"{prefix}{k}: {v}")


def _print_list(items: list, indent: int = 0):
    prefix = "  " * indent
    for i, item in enumerate(items):
        if isinstance(item, dict):
            click.echo(f"{prefix}[{i}]")
            _print_dict(item, indent + 1)
        else:
            click.echo(f"{prefix}- {item}")


def handle_error(func):
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except FileNotFoundError as e:
            if _json_output:
                click.echo(json.dumps({"error": str(e), "type": "file_not_found"}))
            else:
                click.echo(f"错误: {e}", err=True)
            if not _repl_mode:
                sys.exit(1)
        except FileExistsError as e:
            if _json_output:
                click.echo(json.dumps({"error": str(e), "type": "file_exists"}))
            else:
                click.echo(f"错误: {e}", err=True)
            if not _repl_mode:
                sys.exit(1)
        except (ValueError, IndexError, RuntimeError) as e:
            if _json_output:
                click.echo(json.dumps({"error": str(e), "type": type(e).__name__}))
            else:
                click.echo(f"错误: {e}", err=True)
            if not _repl_mode:
                sys.exit(1)
    wrapper.__name__ = func.__name__
    wrapper.__doc__ = func.__doc__
    return wrapper


# ── 主 CLI 组 ──────────────────────────────────────────────────
@click.group(invoke_without_command=True)
@click.option("--json", "use_json", is_flag=True, help="以 JSON 格式输出")
@click.option("--project", "project_path", type=str, default=None,
              help=".wps-cli.json 项目文件路径")
@click.option("--dry-run", "dry_run", is_flag=True, default=False,
              help="运行命令但不保存到磁盘")
@click.pass_context
def cli(ctx, use_json, project_path, dry_run):
    """WPS Office CLI —— 从命令行编辑文档。

    不带子命令运行时进入交互式 REPL 模式。
    """
    global _json_output
    _json_output = use_json

    if project_path:
        sess = get_session()
        if not sess.has_project():
            proj = doc_mod.open_document(project_path)
            sess.set_project(proj, project_path)

    if ctx.invoked_subcommand is None:
        ctx.invoke(repl, project_path=project_path)


@cli.result_callback()
def auto_save_on_exit(result, use_json, project_path, dry_run, **kwargs):
    """单次命令执行后自动保存项目。"""
    if _repl_mode:
        return
    if dry_run:
        return
    sess = get_session()
    if sess.has_project() and sess._modified and sess.project_path:
        try:
            sess.save_session()
        except Exception as e:
            click.echo(f"警告: 自动保存失败: {e}", err=True)


# ── 文档命令 ──────────────────────────────────────────────────
@cli.group()
def document():
    """文档管理命令。"""
    pass


@document.command("new")
@click.option("--type", "doc_type",
              type=click.Choice(["writer", "calc", "impress"]),
              default="writer", help="文档类型")
@click.option("--name", "-n", default="untitled", help="文档名称")
@click.option("--profile", "-p", type=str, default=None, help="页面配置")
@click.option("--output", "-o", "output_path", type=str, default=None, help="保存路径")
@handle_error
def document_new(doc_type, name, profile, output_path):
    """创建新文档。"""
    proj = doc_mod.create_document(doc_type=doc_type, name=name, profile=profile)
    sess = get_session()
    sess.set_project(proj, output_path)
    if output_path:
        doc_mod.save_document(proj, output_path)
    info = doc_mod.get_document_info(proj)
    output(info, f"创建了 {doc_type} 文档: {name}")


@document.command("open")
@click.argument("path")
@click.option("--output", "-o", "output_path", type=str, default=None,
              help="保存打开的项目到此 JSON 路径")
@handle_error
def document_open(path, output_path):
    """打开项目 JSON 文件。"""
    proj = doc_mod.open_document(path)
    sess = get_session()
    sess.set_project(proj, output_path or path)
    info = doc_mod.get_document_info(proj)
    output(info, f"已打开: {path}")


@document.command("save")
@click.argument("path", required=False)
@handle_error
def document_save(path):
    """保存当前文档。"""
    sess = get_session()
    saved = sess.save_session(path)
    output({"saved": saved}, f"已保存到: {saved}")


@document.command("info")
@handle_error
def document_info():
    """显示文档信息。"""
    sess = get_session()
    info = doc_mod.get_document_info(sess.get_project())
    output(info)


@document.command("profiles")
@handle_error
def document_profiles():
    """列出可用的页面配置。"""
    profiles = doc_mod.list_profiles()
    output(profiles, "可用的页面配置:")


@document.command("json")
@handle_error
def document_json():
    """打印原始项目 JSON。"""
    sess = get_session()
    click.echo(json.dumps(sess.get_project(), indent=2, default=str))


# ── Writer 命令 ────────────────────────────────────────────────
@cli.group()
def writer():
    """Writer（文字处理）命令。"""
    pass


@writer.command("add-paragraph")
@click.option("--text", "-t", default="", help="段落文本")
@click.option("--position", "-p", type=int, default=None, help="插入位置")
@click.option("--font-size", type=str, default=None, help="字体大小（如 12pt）")
@click.option("--bold", is_flag=True, help="粗体")
@click.option("--italic", is_flag=True, help="斜体")
@click.option("--alignment", type=click.Choice(["left", "center", "right", "justify"]),
              default=None)
@handle_error
def writer_add_paragraph(text, position, font_size, bold, italic, alignment):
    """添加段落。"""
    sess = get_session()
    sess.snapshot("添加段落")
    style = {}
    if font_size:
        style["font_size"] = font_size
    if bold:
        style["bold"] = True
    if italic:
        style["italic"] = True
    if alignment:
        style["alignment"] = alignment
    item = writer_mod.add_paragraph(
        sess.get_project(), text=text, style=style or None, position=position,
    )
    output(item, "已添加段落")


@writer.command("add-heading")
@click.option("--text", "-t", default="", help="标题文本")
@click.option("--level", "-l", type=int, default=1, help="标题级别（1-6）")
@click.option("--position", "-p", type=int, default=None, help="插入位置")
@handle_error
def writer_add_heading(text, level, position):
    """添加标题。"""
    sess = get_session()
    sess.snapshot("添加标题")
    item = writer_mod.add_heading(
        sess.get_project(), text=text, level=level, position=position,
    )
    output(item, f"已添加标题（级别 {level}）")


@writer.command("add-list")
@click.option("--items", "-i", multiple=True, help="列表项")
@click.option("--style", "list_style",
              type=click.Choice(["bullet", "number"]),
              default="bullet", help="列表样式")
@click.option("--position", "-p", type=int, default=None, help="插入位置")
@handle_error
def writer_add_list(items, list_style, position):
    """添加列表。"""
    sess = get_session()
    sess.snapshot("添加列表")
    item = writer_mod.add_list(
        sess.get_project(), items=list(items), list_style=list_style,
        position=position,
    )
    output(item, f"已添加 {list_style} 列表")


@writer.command("add-table")
@click.option("--rows", "-r", type=int, default=2, help="行数")
@click.option("--cols", "-c", type=int, default=2, help="列数")
@click.option("--position", "-p", type=int, default=None, help="插入位置")
@handle_error
def writer_add_table(rows, cols, position):
    """添加表格。"""
    sess = get_session()
    sess.snapshot("添加表格")
    item = writer_mod.add_table(
        sess.get_project(), rows=rows, cols=cols, position=position,
    )
    output(item, f"已添加 {rows}x{cols} 表格")


@writer.command("add-page-break")
@click.option("--position", "-p", type=int, default=None, help="插入位置")
@handle_error
def writer_add_page_break(position):
    """添加分页符。"""
    sess = get_session()
    sess.snapshot("添加分页符")
    item = writer_mod.add_page_break(sess.get_project(), position=position)
    output(item, "已添加分页符")


@writer.command("add-image")
@click.argument("image_path")
@click.option("--width", "-w", default="10cm", help="宽度")
@click.option("--height", "-h", default="10cm", help="高度")
@click.option("--position", "-p", type=int, default=None, help="插入位置")
@handle_error
def writer_add_image(image_path, width, height, position):
    """添加图片（引用路径）。"""
    sess = get_session()
    sess.snapshot("添加图片")
    item = writer_mod.add_image(
        sess.get_project(), image_path, width=width, height=height,
        position=position,
    )
    output(item, f"已添加图片: {image_path}")


@writer.command("remove")
@click.argument("index", type=int)
@handle_error
def writer_remove(index):
    """按索引删除内容项。"""
    sess = get_session()
    sess.snapshot(f"删除内容 {index}")
    removed = writer_mod.remove_content(sess.get_project(), index)
    output(removed, f"已删除内容索引 {index}")


@writer.command("list")
@handle_error
def writer_list():
    """列出所有内容项。"""
    sess = get_session()
    items = writer_mod.list_content(sess.get_project())
    output(items, "内容项:")


@writer.command("get")
@click.argument("index", type=int)
@handle_error
def writer_get(index):
    """获取指定索引的内容项详情。"""
    sess = get_session()
    item = writer_mod.get_content(sess.get_project(), index)
    output(item)


@writer.command("set-text")
@click.argument("index", type=int)
@click.argument("text")
@handle_error
def writer_set_text(index, text):
    """设置内容项的文本。"""
    sess = get_session()
    sess.snapshot(f"设置索引 {index} 的文本")
    item = writer_mod.set_content_text(sess.get_project(), index, text)
    output(item, f"已更新索引 {index} 的文本")


@writer.command("find-replace")
@click.argument("find_text")
@click.argument("replace_text")
@handle_error
def writer_find_replace(find_text, replace_text):
    """查找并替换文本。"""
    sess = get_session()
    sess.snapshot(f"查找替换: {find_text} → {replace_text}")
    result = writer_mod.find_replace(sess.get_project(), find_text, replace_text)
    output(result, f"已替换 {result['replaced']} 处")


# ── Calc 命令 ──────────────────────────────────────────────────
@cli.group()
def calc():
    """Calc（电子表格）命令。"""
    pass


@calc.command("add-sheet")
@click.option("--name", "-n", default="Sheet", help="工作表名称")
@click.option("--position", "-p", type=int, default=None, help="插入位置")
@handle_error
def calc_add_sheet(name, position):
    """添加新工作表。"""
    sess = get_session()
    sess.snapshot(f"添加工作表: {name}")
    sheet = calc_mod.add_sheet(sess.get_project(), name=name, position=position)
    output(sheet, f"已添加工作表: {name}")


@calc.command("remove-sheet")
@click.argument("index", type=int)
@handle_error
def calc_remove_sheet(index):
    """按索引删除工作表。"""
    sess = get_session()
    sess.snapshot(f"删除工作表 {index}")
    removed = calc_mod.remove_sheet(sess.get_project(), index)
    output(removed, f"已删除工作表索引 {index}")


@calc.command("rename-sheet")
@click.argument("index", type=int)
@click.argument("name")
@handle_error
def calc_rename_sheet(index, name):
    """重命名工作表。"""
    sess = get_session()
    sess.snapshot(f"重命名工作表 {index}")
    sheet = calc_mod.rename_sheet(sess.get_project(), index, name)
    output(sheet, f"已重命名工作表 {index} 为: {name}")


@calc.command("set-cell")
@click.argument("ref")
@click.argument("value")
@click.option("--type", "cell_type", default="string", help="数据类型: string / float")
@click.option("--sheet", "-s", type=int, default=0, help="工作表索引")
@click.option("--formula", type=str, default=None, help="单元格公式")
@handle_error
def calc_set_cell(ref, value, cell_type, sheet, formula):
    """设置单元格的值。"""
    sess = get_session()
    sess.snapshot(f"设置单元格 {ref}")
    result = calc_mod.set_cell(
        sess.get_project(), ref=ref, value=value, cell_type=cell_type,
        sheet=sheet, formula=formula,
    )
    output(result, f"设置 {ref} = {value}")


@calc.command("get-cell")
@click.argument("ref")
@click.option("--sheet", "-s", type=int, default=0, help="工作表索引")
@handle_error
def calc_get_cell(ref, sheet):
    """获取单元格的值。"""
    sess = get_session()
    result = calc_mod.get_cell(sess.get_project(), ref=ref, sheet=sheet)
    output(result)


@calc.command("set-range")
@click.argument("start_ref")
@click.option("--data", "-d", type=str, required=True,
              help="JSON 格式的二维数据数组")
@click.option("--sheet", "-s", type=int, default=0, help="工作表索引")
@handle_error
def calc_set_range(start_ref, data, sheet):
    """批量写入矩形区域数据。"""
    import json as _json
    sess = get_session()
    sess.snapshot(f"批量写入从 {start_ref}")
    arr = _json.loads(data)
    result = calc_mod.set_range(sess.get_project(), start_ref, arr, sheet=sheet)
    output(result, f"已写入 {result['cells_set']} 个单元格")


@calc.command("merge-cells")
@click.argument("start_ref")
@click.argument("end_ref")
@click.option("--sheet", "-s", type=int, default=0, help="工作表索引")
@handle_error
def calc_merge_cells(start_ref, end_ref, sheet):
    """合并单元格区域。"""
    sess = get_session()
    sess.snapshot(f"合并单元格 {start_ref}:{end_ref}")
    result = calc_mod.merge_cells(sess.get_project(), start_ref, end_ref, sheet=sheet)
    output(result, f"已标记合并: {start_ref}:{end_ref}")


@calc.command("clear-cell")
@click.argument("ref")
@click.option("--sheet", "-s", type=int, default=0, help="工作表索引")
@handle_error
def calc_clear_cell(ref, sheet):
    """清除单元格的内容。"""
    sess = get_session()
    sess.snapshot(f"清除单元格 {ref}")
    result = calc_mod.clear_cell(sess.get_project(), ref=ref, sheet=sheet)
    output(result, f"已清除单元格 {ref}")


@calc.command("list-sheets")
@handle_error
def calc_list_sheets():
    """列出所有工作表。"""
    sess = get_session()
    sheets = calc_mod.list_sheets(sess.get_project())
    output(sheets, "工作表:")


# ── Impress 命令 ───────────────────────────────────────────────
@cli.group()
def impress():
    """Impress（演示文稿）命令。"""
    pass


@impress.command("add-slide")
@click.option("--title", "-t", default="", help="幻灯片标题")
@click.option("--content", "-c", default="", help="幻灯片内容")
@click.option("--position", "-p", type=int, default=None, help="插入位置")
@handle_error
def impress_add_slide(title, content, position):
    """添加幻灯片。"""
    sess = get_session()
    sess.snapshot("添加幻灯片")
    slide = impress_mod.add_slide(
        sess.get_project(), title=title, content=content, position=position,
    )
    output(slide, f"已添加幻灯片: {title}")


@impress.command("remove-slide")
@click.argument("index", type=int)
@handle_error
def impress_remove_slide(index):
    """按索引删除幻灯片。"""
    sess = get_session()
    sess.snapshot(f"删除幻灯片 {index}")
    removed = impress_mod.remove_slide(sess.get_project(), index)
    output(removed, f"已删除幻灯片 {index}")


@impress.command("set-content")
@click.argument("index", type=int)
@click.option("--title", "-t", type=str, default=None, help="新标题")
@click.option("--content", "-c", type=str, default=None, help="新内容")
@handle_error
def impress_set_content(index, title, content):
    """更新幻灯片内容。"""
    sess = get_session()
    sess.snapshot(f"更新幻灯片 {index}")
    slide = impress_mod.set_slide_content(
        sess.get_project(), index, title=title, content=content,
    )
    output(slide, f"已更新幻灯片 {index}")


@impress.command("list-slides")
@handle_error
def impress_list_slides():
    """列出所有幻灯片。"""
    sess = get_session()
    slides = impress_mod.list_slides(sess.get_project())
    output(slides, "幻灯片:")


@impress.command("add-element")
@click.argument("slide_index", type=int)
@click.option("--type", "element_type", default="text_box", help="元素类型")
@click.option("--text", "-t", default="", help="元素文本")
@click.option("--x", default="2cm", help="X 位置")
@click.option("--y", default="2cm", help="Y 位置")
@click.option("--width", "-w", default="10cm", help="宽度")
@click.option("--height", "-h", default="5cm", help="高度")
@handle_error
def impress_add_element(slide_index, element_type, text, x, y, width, height):
    """向幻灯片添加元素。"""
    sess = get_session()
    sess.snapshot(f"向幻灯片 {slide_index} 添加元素")
    elem = impress_mod.add_slide_element(
        sess.get_project(), slide_index,
        element_type=element_type, text=text,
        x=x, y=y, width=width, height=height,
    )
    output(elem, f"已添加 {element_type} 到幻灯片 {slide_index}")


@impress.command("remove-element")
@click.argument("slide_index", type=int)
@click.argument("element_index", type=int)
@handle_error
def impress_remove_element(slide_index, element_index):
    """删除幻灯片中的元素。"""
    sess = get_session()
    sess.snapshot(f"删除幻灯片 {slide_index} 中的元素 {element_index}")
    removed = impress_mod.remove_slide_element(sess.get_project(), slide_index, element_index)
    output(removed, f"已删除幻灯片 {slide_index} 中的元素 {element_index}")


@impress.command("move-slide")
@click.argument("from_index", type=int)
@click.argument("to_index", type=int)
@handle_error
def impress_move_slide(from_index, to_index):
    """移动幻灯片到新位置。"""
    sess = get_session()
    sess.snapshot(f"移动幻灯片 {from_index} → {to_index}")
    result = impress_mod.move_slide(sess.get_project(), from_index, to_index)
    output(result, f"已移动幻灯片 {from_index} 到 {to_index}")


@impress.command("duplicate-slide")
@click.argument("index", type=int)
@handle_error
def impress_duplicate_slide(index):
    """复制幻灯片。"""
    sess = get_session()
    sess.snapshot(f"复制幻灯片 {index}")
    slide = impress_mod.duplicate_slide(sess.get_project(), index)
    output(slide, f"已复制幻灯片 {index}")


# ── 样式命令 ──────────────────────────────────────────────────
@cli.group("style")
def style_group():
    """样式管理命令。"""
    pass


@style_group.command("create")
@click.argument("name")
@click.option("--family", type=click.Choice(["paragraph", "text"]),
              default="paragraph", help="样式系列")
@click.option("--parent", type=str, default=None, help="父样式名称")
@click.option("--prop", "-p", multiple=True, help="属性: key=value")
@handle_error
def style_create(name, family, parent, prop):
    """创建新样式。"""
    props = _parse_props(prop)
    sess = get_session()
    sess.snapshot(f"创建样式: {name}")
    result = styles_mod.create_style(
        sess.get_project(), name=name, family=family,
        parent=parent, properties=props,
    )
    output(result, f"已创建样式: {name}")


@style_group.command("modify")
@click.argument("name")
@click.option("--prop", "-p", multiple=True, help="属性: key=value")
@handle_error
def style_modify(name, prop):
    """修改已有样式。"""
    props = _parse_props(prop)
    sess = get_session()
    sess.snapshot(f"修改样式: {name}")
    result = styles_mod.modify_style(
        sess.get_project(), name=name, properties=props,
    )
    output(result, f"已修改样式: {name}")


@style_group.command("list")
@handle_error
def style_list():
    """列出所有样式。"""
    sess = get_session()
    styles = styles_mod.list_styles(sess.get_project())
    output(styles, "样式:")


@style_group.command("apply")
@click.argument("style_name")
@click.argument("content_index", type=int)
@handle_error
def style_apply(style_name, content_index):
    """将样式应用于 Writer 内容项。"""
    sess = get_session()
    sess.snapshot(f"应用样式 {style_name} 到 {content_index}")
    result = styles_mod.apply_style(
        sess.get_project(), style_name, content_index,
    )
    output(result, f"已将 '{style_name}' 应用到内容 {content_index}")


@style_group.command("remove")
@click.argument("name")
@handle_error
def style_remove(name):
    """删除样式。"""
    sess = get_session()
    sess.snapshot(f"删除样式: {name}")
    result = styles_mod.remove_style(sess.get_project(), name)
    output(result, f"已删除样式: {name}")


# ── 导出命令 ──────────────────────────────────────────────────
@cli.group("export")
def export_group():
    """导出/渲染命令。"""
    pass


@export_group.command("presets")
@handle_error
def export_presets():
    """列出导出预设。"""
    presets = export_mod.list_presets()
    output(presets, "导出预设:")


@export_group.command("preset-info")
@click.argument("name")
@handle_error
def export_preset_info(name):
    """显示预设详情。"""
    info = export_mod.get_preset_info(name)
    output(info)


@export_group.command("render")
@click.argument("output_path")
@click.option("--preset", "-p", default="docx", help="导出预设")
@click.option("--overwrite", is_flag=True, help="覆盖已有文件")
@handle_error
def export_render(output_path, preset, overwrite):
    """导出文档到文件（使用 WPS COM 自动化）。"""
    sess = get_session()
    result = export_mod.export(
        sess.get_project(), output_path,
        preset=preset, overwrite=overwrite,
    )
    output(result, f"已导出到: {result['output']} ({result.get('file_size', 0):,} 字节)")


# ── 设计预设命令 ───────────────────────────────────────────
@cli.group("preset")
def preset_group():
    """设计预设 —— 一键切换 PPT 风格和布局模板。"""
    pass


@preset_group.command("list")
@handle_error
def preset_list():
    """列出所有可用的设计预设和布局模板。"""
    from cli_anything.wps.styles.design_presets import PRESETS
    from cli_anything.wps.styles.layout_templates import LAYOUTS, TALK_PRESETS

    print("\n=== 设计预设 (Design Presets) ===")
    for name, preset in PRESETS.items():
        colors = preset.colors
        print(f"  {name}: {preset.name}")
        print(f"    主色: RGB{colors['primary']}  辅色: RGB{colors['secondary']}")
        print(f"    标题: {preset.fonts['title'][0]} {preset.fonts['title'][1]}pt")

    print("\n=== 布局模板 (Layout Templates) ===")
    for name, layout in LAYOUTS.items():
        print(f"  {name}: {layout.description}")

    print("\n=== 演讲类型预设 ===")
    for name, tp in TALK_PRESETS.items():
        pages = len(tp["slides"])
        print(f"  {name}: {tp['name']} ({pages}页)")

    output({"presets": list(PRESETS.keys()),
            "layouts": list(LAYOUTS.keys()),
            "talk_types": list(TALK_PRESETS.keys())})


@preset_group.command("info")
@click.argument("name")
@handle_error
def preset_info(name):
    """查看指定预设的详细配置。"""
    from cli_anything.wps.styles.design_presets import PRESETS
    from cli_anything.wps.styles.layout_templates import TALK_PRESETS

    if name in PRESETS:
        p = PRESETS[name]
        print(f"=== {p.name} ({name}) ===")
        print(f"颜色: {p.colors}")
        print(f"字体: {p.fonts}")
        print(f"间距: {p.spacing}")
        print(f"规则: {p.rules}")
    elif name in TALK_PRESETS:
        tp = TALK_PRESETS[name]
        print(f"=== {tp['name']} ({name}) ===")
        print(f"推荐页面序列: {' → '.join(tp['slides'])}")
        print(f"规则: {tp['rules']}")
    else:
        from cli_anything.wps.styles.layout_templates import LAYOUTS
        if name in LAYOUTS:
            lt = LAYOUTS[name]
            print(f"=== {lt.name} ({name}) ===")
            print(f"分类: {lt.category}")
            print(f"描述: {lt.description}")
            print(f"元素数: {len(lt.elements)}")


@preset_group.command("apply")
@click.argument("preset_name")
@click.option("--talk-type", "-t", type=str, default=None,
              help="演讲类型: conference/business/defense/school")
@handle_error
def preset_apply(preset_name, talk_type):
    """应用设计预设到当前项目。

    preset_name: academic / consultant / business / tech
    """
    from cli_anything.wps.styles.design_presets import PRESETS
    from cli_anything.wps.styles.layout_templates import TALK_PRESETS

    if preset_name not in PRESETS:
        available = ", ".join(PRESETS.keys())
        raise ValueError(f"未知预设: {preset_name}。可用: {available}")

    sess = get_session()
    proj = sess.get_project()
    proj.setdefault("design", {})["preset"] = preset_name
    if talk_type:
        proj["design"]["talk_type"] = talk_type
    sess._modified = True

    preset = PRESETS[preset_name]
    output({"applied": preset_name, "preset": preset.to_dict()},
           f"已应用设计预设: {preset.name}")


# ── 会话命令 ──────────────────────────────────────────────────
@cli.group()
def session():
    """会话管理命令。"""
    pass


@session.command("status")
@handle_error
def session_status():
    """显示会话状态。"""
    sess = get_session()
    output(sess.status())


@session.command("undo")
@handle_error
def session_undo():
    """撤销上一步操作。"""
    sess = get_session()
    desc = sess.undo()
    output({"undone": desc}, f"已撤销: {desc}")


@session.command("redo")
@handle_error
def session_redo():
    """重做上一步撤销。"""
    sess = get_session()
    desc = sess.redo()
    output({"redone": desc}, f"已重做: {desc}")


@session.command("history")
@handle_error
def session_history():
    """显示撤销历史。"""
    sess = get_session()
    history = sess.list_history()
    output(history, "撤销历史:")


# ── REPL ───────────────────────────────────────────────────────
@cli.command()
@click.option("--project", "project_path", type=str, default=None)
@handle_error
def repl(project_path):
    """启动交互式 REPL 会话。"""
    from cli_anything.wps.utils.repl_skin import ReplSkin

    global _repl_mode
    _repl_mode = True

    skin = ReplSkin("wps", version="1.0.0")

    if project_path:
        sess = get_session()
        proj = doc_mod.open_document(project_path)
        sess.set_project(proj, project_path)

    skin.print_banner()

    pt_session = skin.create_prompt_session()

    def _get_project_name():
        try:
            s = get_session()
            proj = s.get_project()
            if proj and isinstance(proj, dict):
                return proj.get("name", "")
        except Exception:
            pass
        return ""

    def _is_modified():
        try:
            s = get_session()
            return s.is_modified()
        except Exception:
            return False

    while True:
        try:
            line = skin.get_input(
                pt_session,
                project_name=_get_project_name(),
                modified=_is_modified(),
            ).strip()
            if not line:
                continue
            if line.lower() in ("quit", "exit", "q"):
                skin.print_goodbye()
                break
            if line.lower() == "help":
                _repl_help(skin)
                continue

            try:
                args = shlex.split(line)
            except ValueError:
                args = line.split()
            try:
                cli.main(args, standalone_mode=False)
            except SystemExit:
                pass
            except click.exceptions.UsageError as e:
                skin.error(f"用法错误: {e}")
            except Exception as e:
                skin.error(str(e))

        except (EOFError, KeyboardInterrupt):
            skin.print_goodbye()
            break

    _repl_mode = False


def _repl_help(skin=None):
    commands = {
        "document new|open|save|info|profiles|json": "文档管理",
        "writer add-paragraph|add-heading|add-list|add-table|add-page-break|add-image|remove|list|get|set-text|find-replace": "Writer 编辑",
        "calc add-sheet|remove-sheet|rename-sheet|set-cell|get-cell|set-range|merge-cells|clear-cell|list-sheets": "电子表格编辑",
        "impress add-slide|remove-slide|set-content|list-slides|add-element|remove-element|move-slide|duplicate-slide": "演示文稿编辑",
        "style create|modify|list|apply|remove": "样式管理",
        "export presets|preset-info|render": "导出文档",
        "session status|undo|redo|history": "会话管理",
        "help": "显示帮助",
        "quit": "退出 REPL",
    }
    if skin is not None:
        skin.help(commands)
    else:
        click.echo("\n命令:")
        for cmd, desc in commands.items():
            click.echo(f"  {cmd:60s}  {desc}")
        click.echo()


def _parse_props(prop_list):
    """解析 CLI 中的 key=value 属性对。"""
    props = {}
    for p in prop_list:
        if "=" not in p:
            raise ValueError(f"属性格式无效: '{p}'。请使用 key=value。")
        k, v = p.split("=", 1)
        if v.lower() == "true":
            v = True
        elif v.lower() == "false":
            v = False
        else:
            try:
                v = float(v) if "." in v else int(v)
            except ValueError:
                pass
        props[k] = v
    return props


# ── 入口点 ────────────────────────────────────────────────────
def main():
    cli()


if __name__ == "__main__":
    main()
