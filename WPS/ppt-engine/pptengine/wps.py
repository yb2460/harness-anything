# -*- coding: utf-8 -*-
"""WPS COM 会话与 Deck（演示文稿）封装。

依赖：Windows + 已安装 WPS Office + pywin32。
COM ProgID 固定为 KWPP.Application（WPS 演示）。
"""
from __future__ import annotations

import os
import time
from pathlib import Path

from . import paths


class WPSConnectionError(RuntimeError):
    pass


# 演示文稿保存格式常量（WPS 12.x 实测；ppSaveAsOpenXMLPresentation=24 才落成 OOXML/ZIP）
_SAVEAS_FORMAT = {".pptx": 24, ".ppt": 1, ".ppsx": 28, ".pps": 1}


def connect_wps(visible: bool = True, kill_first: bool = False, retries: int = 4):
    """连接 WPS 演示 COM。

    kill_first 默认关闭：`taskkill /F /IM wps.exe` 会连带杀掉用户已打开的文档，
    只有显式传 True 时才清场。

    WPS 12.x 的 KWPP.Application 不接受 Visible=False（抛 E_FAIL 0x80004005），
    因此 visible=False 用「Visible=True + WindowState=1（最小化）」实现后台生成。
    """
    import pythoncom
    import win32com.client

    if kill_first:
        os.system("taskkill /F /IM wps.exe 2>nul")
        time.sleep(3)

    last = None
    for i in range(retries):
        try:
            pythoncom.CoInitialize()
            app = win32com.client.Dispatch("KWPP.Application")
            time.sleep(1.5)
            _ = app.Presentations.Count   # 触发 RPC 验证
            app.Visible = True
            if not visible:
                try:
                    app.WindowState = 1   # ppWindowMinimized
                except Exception:         # noqa: BLE001
                    pass
            return app
        except Exception as e:            # noqa: BLE001
            last = e
            time.sleep(2)
    raise WPSConnectionError(f"无法连接 WPS COM（请确认已安装 WPS）: {last}")


class Deck:
    """一个演示文稿会话。

    deck = Deck()                       # 960x540
    s = deck.add_slide()                # 空白页
    ...  # 用 pptengine.elements 画元素
    deck.save("demo.pptx")              # 同时导出 pdf
    """

    def __init__(self, width: int = 960, height: int = 540,
                 assets_dir=None, visible: bool = True, bg=None,
                 font: str = "微软雅黑"):
        self.width = int(width)
        self.height = int(height)
        self.font = font
        self.assets_dir = Path(assets_dir).expanduser().resolve() if assets_dir else paths.assets_dir()
        paths.ensure_dir(self.assets_dir)
        self.bg = bg
        self.app = connect_wps(visible=visible)
        self.pres = self.app.Presentations.Add(True)   # True = 空白演示
        self.pres.PageSetup.SlideWidth = self.width
        self.pres.PageSetup.SlideHeight = self.height
        # 让 elements 模块知道去哪里找图片
        try:
            from . import elements
            elements.set_assets_dir(self.assets_dir)
        except Exception:                                # noqa: BLE001
            pass

    # ---------- 页面 ----------
    def add_slide(self, bg=None):
        """新增一页（布局 12 = 空白），并铺母版底图。"""
        idx = self.pres.Slides.Count + 1
        s = self.pres.Slides.Add(idx, 12)
        bg = bg or self.bg
        if bg and os.path.exists(bg):
            try:
                s.Background.Fill.UserPicture(str(bg))
            except Exception as e:                       # noqa: BLE001
                print(f"[bg] 铺底图失败: {e}")
        try:
            self.app.ActiveWindow.View.GotoSlide(idx)
        except Exception:                                # noqa: BLE001
            pass
        return s

    @property
    def page_count(self) -> int:
        return self.pres.Slides.Count

    # ---------- 保存 ----------
    def save(self, out_path, pdf: bool = True, pdf_path=None):
        """保存 pptx；pdf=True 时同目录再导出一份 PDF。

        返回 (pptx路径, pdf路径或None)。落盘后做存在性与大小校验，
        避免 WPS SaveAs 静默失败。
        """
        out_path = Path(out_path).expanduser().resolve()
        paths.ensure_dir(out_path.parent)
        fmt = _SAVEAS_FORMAT.get(out_path.suffix.lower())
        if fmt is not None:
            self.pres.SaveAs(str(out_path), fmt)          # 24 = pptx(OXML)，1 = ppt
        else:
            self.pres.SaveAs(str(out_path))
        pdf_out = None
        if pdf:
            pdf_out = Path(pdf_path).expanduser().resolve() if pdf_path else out_path.with_suffix(".pdf")
            self.pres.SaveAs(str(pdf_out), 32)           # 32 = PDF
        for f in filter(None, (out_path, pdf_out)):
            ok = os.path.exists(f) and os.path.getsize(f) > 0
            print(f"[save] {f}  存在={ok} 大小={os.path.getsize(f) if ok else 0}")
        return str(out_path), (str(pdf_out) if pdf_out else None)

    def close(self, quit_app: bool = False):
        """默认不退出 WPS，保留窗口供人工查看/编辑。"""
        if quit_app:
            try:
                self.pres.Close()
            except Exception:                            # noqa: BLE001
                pass
            try:
                self.app.Quit()
            except Exception:                            # noqa: BLE001
                pass
