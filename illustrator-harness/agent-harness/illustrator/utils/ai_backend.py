# -*- coding: utf-8 -*-
"""Illustrator backend detection and management."""

import os
import pythoncom
import win32com.client


def detect_illustrator() -> dict:
    """Detect Adobe Illustrator installation."""
    info = {
        "com_available": False,
        "version": "",
        "path": "",
        "executable": "",
    }

    pythoncom.CoInitialize()
    try:
        app = win32com.client.Dispatch("Illustrator.Application")
        info["com_available"] = True
        info["version"] = str(app.Version)
        info["path"] = str(app.Path)
        # Find executable
        exe_path = os.path.join(app.Path, "Support Files", "Contents", "Windows", "Illustrator.exe")
        if os.path.exists(exe_path):
            info["executable"] = exe_path
        else:
            info["executable"] = "Illustrator.exe (in PATH)"
    except Exception as e:
        info["error"] = str(e)
    finally:
        pythoncom.CoUninitialize()

    return info


def launch_illustrator(visible: bool = True) -> object:
    """Launch or connect to Illustrator COM application.

    Note: Illustrator COM does not support setting Visible property.
    The app inherits visibility from its current window state.
    """
    pythoncom.CoInitialize()
    try:
        app = win32com.client.Dispatch("Illustrator.Application")
        # Note: Visible is read-only for Illustrator COM
        try:
            app.Visible = visible
        except Exception:
            pass  # OK - Illustrator manages its own visibility
        return app
    except Exception as e:
        raise RuntimeError(
            f"Adobe Illustrator not found. Error: {e}"
        )
