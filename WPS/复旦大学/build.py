# -*- coding: utf-8 -*-
"""Fudan Academic PPT - WPS COM automation."""
import os, json, pythoncom, win32com.client

OUT = r"D:\A-资料\A-claudewenjian\PPT制作\复旦大学\work"
JSON_PATH = os.path.join(OUT, "fudan_data.json")
BG_IMAGE = os.path.join(OUT, "template_bg.png")

FONT_TITLE = "SimHei"
FONT_BODY = "Microsoft YaHei"
BLUE = '#003366'

def hex_to_bgr(h):
    h = h.lstrip('#')
    r,g,b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
    return (b<<16)|(g<<8)|r

def run():
    with open(JSON_PATH,'r',encoding='utf-8') as f:
        data = json.load(f)
    W,H = data['canvas']['w'], data['canvas']['h']
    slides_data = data['slides']

    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("KWPP.Application")
    app.Visible = True
    ppt = app.Presentations.Add()
    ppt.PageSetup.SlideWidth = W
    ppt.PageSetup.SlideHeight = H

    idx = [1]
    def slide():
        s = ppt.Slides.Add(idx[0], 12)
        idx[0] += 1
        try: s.FollowMasterBackground = False
        except: pass
        s.Background.Fill.UserPicture(BG_IMAGE)
        return s

    def txt(s, x, y, w, h, text, fs=28, color=0x333333, bold=False, align=1, font=FONT_BODY, spacing=1.3):
        t = s.Shapes.AddTextbox(1, x, y, w, h)
        tr = t.TextFrame.TextRange; tr.Text = text
        tr.Font.Size = fs; tr.Font.Color = color; tr.Font.Name = font
        tr.Font.Bold = bold; tr.ParagraphFormat.Alignment = align
        try: tr.ParagraphFormat.SpaceWithin = spacing
        except: pass
        return t

    def rect(s, x, y, w, h, color):
        r = s.Shapes.AddShape(1, x, y, w, h)
        r.Fill.ForeColor.RGB = color; r.Fill.Visible = True; r.Line.Visible = False
        return r

    def draw_text(s, e):
        c = hex_to_bgr(e.get('color','#333333'))
        return txt(s, e['x'], e['y'], e['w'], e['h'], e['text'],
                   fs=e.get('fs',28), color=c, bold=e.get('bold',False),
                   align=e.get('align',1), font=e.get('font',FONT_BODY), spacing=e.get('line_spacing',1.3))

    def draw_image(s, e):
        path = os.path.join(OUT, e['file'])
        if os.path.exists(path):
            s.Shapes.AddPicture(path, False, True, e['x'], e['y'], e['w'], e['h'])

    def draw_cards_2x3(s, e):
        items = e['items']; sy = e.get('start_y',68); cw,ch=295,215; gx,gy=14,12
        for i, item in enumerate(items):
            r,c = i//3, i%3; x=22+c*(cw+gx); y=sy+r*(ch+gy)
            col = hex_to_bgr(item['color'])
            rect(s, x, y, cw, 4, col)
            txt(s, x+10, y+14, cw-20, 32, item['title'], fs=26, color=col, bold=True, align=1, font=FONT_TITLE)
            txt(s, x+10, y+52, cw-20, ch-62, item['desc'], fs=20, color=hex_to_bgr('#333333'), bold=False, align=1, font=FONT_BODY, spacing=1.4)

    ROUTERS = {'text': draw_text, 'image': draw_image, 'cards_2x3': draw_cards_2x3}

    for sd in slides_data:
        s = slide()
        for elem in sd.get('elements', []):
            etype = elem.get('type','text')
            router = ROUTERS.get(etype)
            if router:
                try: router(s, elem)
                except Exception as ex: print(f"WARN [{sd.get('title','?')}] {etype}: {ex}")

    pptx_path = os.path.join(OUT, "肝癌TIME亚型-学术汇报-复旦模版.pptx")
    ppt.SaveAs(pptx_path)
    print(f"PPTX: {pptx_path} ({os.path.getsize(pptx_path):,} bytes)")
    pdf_path = os.path.join(OUT, "肝癌TIME亚型-学术汇报-复旦模版.pdf")
    try:
        ppt.SaveAs(pdf_path, 32)
        print(f"PDF: {pdf_path} ({os.path.getsize(pdf_path):,} bytes)")
    except Exception as ex: print(f"PDF failed: {ex}")
    ppt.Close()
    try: app.Quit()
    except: pass
    print("Done!")

if __name__ == "__main__":
    run()