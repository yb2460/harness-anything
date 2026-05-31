# -*- coding: utf-8 -*-
"""Quick test of Illustrator harness."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cli_anything.illustrator.utils.ai_backend import launch_illustrator, detect_illustrator
from cli_anything.illustrator.core.shapes import add_rectangle, add_circle, add_star
from cli_anything.illustrator.core.text import add_text
from cli_anything.illustrator.core.export import export_as, save_as_ai
from cli_anything.illustrator.core.project import get_document_info
import pythoncom

pythoncom.CoInitialize()

# Detect
info = detect_illustrator()
print("Illustrator:", info["version"], info["path"])

# Launch
app = launch_illustrator()

# Ensure we have a document
try:
    doc = app.ActiveDocument
    print("Using existing doc:", doc.Name)
except Exception:
    app.Documents.Add()
    print("Created new doc")

# Add shapes
add_rectangle(app, 50, 50, 200, 150, fill_rgb=(66, 133, 244))
print("+ Rectangle (Google Blue)")

add_circle(app, 400, 200, 80, fill_rgb=(234, 67, 53))
print("+ Circle (Google Red)")

add_star(app, 600, 200, 80, 40, 5, fill_rgb=(251, 188, 4))
print("+ Star (Google Yellow)")

# Add text
add_text(app, "Illustrator CLI Harness v1.0", x=50, y=300, font_size=36)
print("+ Text")

# Info
info = get_document_info(app)
print("\nDocument info:")
for k, v in info.items():
    print(f"  {k}: {v}")

# Export
out = r"D:\A-资料\A-claudewenjian\ms-ppt\ai_harness_test.png"
try:
    export_as(app, out, "PNG")
    print(f"\nExported: {out} ({os.path.getsize(out):,} bytes)")
except Exception as e:
    print(f"Export error (non-fatal): {e}")

# Save AI
ai_out = r"D:\A-资料\A-claudewenjian\ms-ppt\ai_harness_test.ai"
save_as_ai(app, ai_out)
print(f"Saved: {ai_out} ({os.path.getsize(ai_out):,} bytes)")

pythoncom.CoUninitialize()
print("\nDONE - Illustrator harness works!")
