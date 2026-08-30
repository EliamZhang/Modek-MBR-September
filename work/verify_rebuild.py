# -*- coding: utf-8 -*-
"""COM 验证重建后的体系规划页:文本、折行、溢出、重叠。"""
import win32com.client

PPTX = r"c:\Users\zhangyuliang02\Desktop\临时文件\9月MBR\Model MBR August.pptx"

app = win32com.client.DispatchEx("PowerPoint.Application")
pres = app.Presentations.Open(PPTX, ReadOnly=True, WithWindow=False)
try:
    n = pres.Slides.Count
    print("total slides:", n)
    for si in range(1, n + 1):
        sl = pres.Slides(si)
        print(f"\n=== Slide {si} shapes: {sl.Shapes.Count}")
        for shp in sl.Shapes:
            t = ""
            if shp.HasTextFrame and shp.TextFrame.HasText:
                tr = shp.TextFrame.TextRange
                t = tr.Text.replace("\r", " | ")[:60]
            print(f"  ({shp.Left/72:.2f},{shp.Top/72:.2f}) {shp.Width/72:.2f}x{shp.Height/72:.2f} :: {t}")
            if shp.HasTextFrame and shp.TextFrame.HasText:
                tr = shp.TextFrame.TextRange
                npara = tr.Paragraphs().Count
                nline = tr.Lines().Count
                bh = tr.BoundHeight
                avail = shp.Height * 72 - shp.TextFrame.MarginTop - shp.TextFrame.MarginBottom
                o = ""
                if nline > npara:
                    o += " <-- WRAP"
                if bh > avail + 3:
                    o += f" **OVERFLOW {bh/72:.2f}>{avail/72:.2f}"
                print(f"      p={npara} l={nline}{o}")

        print("  -- overlap --")
        shapes = list(sl.Shapes)
        for i in range(len(shapes)):
            for j in range(i + 1, len(shapes)):
                a, b = shapes[i], shapes[j]
                ox = min(a.Left + a.Width, b.Left + b.Width) - max(a.Left, b.Left)
                oy = min(a.Top + a.Height, b.Top + b.Height) - max(a.Top, b.Top)
                if ox > 3 and oy > 3:
                    print(f"    OVERLAP x: {ox/72:.2f} y: {oy/72:.2f}")
finally:
    pres.Close()
    app.Quit()
