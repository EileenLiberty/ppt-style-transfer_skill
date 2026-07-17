# -*- coding: utf-8 -*-
"""Re-convert WMF to high-res PNG by scaling shape before export."""
from pathlib import Path
import win32com.client

SRC_MEDIA = Path(r"g:\ppt generate\src_unpacked\ppt\media")
OUT_DIR = Path(r"g:\ppt generate\media_png")
OUT_DIR.mkdir(parents=True, exist_ok=True)

pp = win32com.client.Dispatch("PowerPoint.Application")
try:
    pp.Visible = 1
except Exception:
    pass

for wmf in sorted(SRC_MEDIA.glob("*.wmf")):
    png = OUT_DIR / (wmf.stem + ".png")
    pres = pp.Presentations.Add()
    try:
        slide = pres.Slides.Add(1, 12)
        shape = slide.Shapes.AddPicture(str(wmf.resolve()), False, True, 10, 10)
        # Scale up to ~2000px wide equivalent (~20" at 96dpi -> use 15")
        target_w = 15 * 72  # points
        if shape.Width > 0:
            ratio = target_w / float(shape.Width)
            shape.Width = target_w
            shape.Height = float(shape.Height) * ratio
        shape.Export(str(png.resolve()), 2)
        print(wmf.name, "->", png.name, png.stat().st_size)
    except Exception as e:
        print("FAIL", wmf.name, e)
    finally:
        try:
            pres.Saved = True
            pres.Close()
        except Exception:
            pass

try:
    pp.Quit()
except Exception:
    pass
print("done")
