# -*- coding: utf-8 -*-
"""COM 验证新页:文本完整、折行、边界、重叠。"""
import win32com.client
import pythoncom

PPTX = r"c:\Users\zhangyuliang02\Desktop\临时文件\9月MBR\Model MBR August.pptx"

app = win32com.client.DispatchEx("PowerPoint.Application")
pres = app.Presentations.Open(PPTX, ReadOnly=True, WithWindow=False)
try:
    n = pres.Slides.Count
    print("total slides:", n)
    s = pres.Slides(n)  # 新页应在最后?不,插到第 2 位
    for si in range(1, n + 1):
        sl = pres.Slides(si)
        print(f"\n=== Slide {si} shapes: {sl.Shapes.Count}")
        for shp in sl.Shapes:
            t = ""
            if shp.HasTextFrame and shp.TextFrame.HasText:
                tr = shp.TextFrame.TextRange
                t = tr.Text.replace("\r", " | ")[:60]
            print(f"  [{shp.Name}] ({shp.Left/72:.2f},{shp.Top/72:.2f}) "
                  f"{shp.Width/72:.2f}x{shp.Height/72:.2f} type={shp.Type} :: {t}")
            if shp.HasTextFrame and shp.TextFrame.HasText:
                tr = shp.TextFrame.TextRange
                npara = tr.Paragraphs().Count
                nline = tr.Lines().Count
                flag = "  <-- CHECK" if (npara == 1 and nline > 1) else ""
                print(f"      paras={npara} lines={nline}{flag}")
                # 溢出检查:文本边界高度(点) vs 框高(点)
                bh = tr.BoundHeight
                avail = shp.Height * 72 - shp.TextFrame.MarginTop - shp.TextFrame.MarginBottom
                if bh > avail + 3:
                    print(f"      ** OVERFLOW: BoundHeight={bh/72:.2f}in > avail {avail/72:.2f}in")
finally:
    pres.Close()
    app.Quit()
