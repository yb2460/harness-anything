# -*- coding: utf-8 -*-
"""NJU Admissions PPT - WPS COM automation."""
import os, json, pythoncom, win32com.client

OUT = r"D:\A-资料\A-claudewenjian\PPT制作\南京大学\work"
JSON_PATH = os.path.join(OUT, "nju_data.json")
BG_IMAGE = os.path.join(OUT, "template_bg.png")

FONT_TITLE = "SimHei"
FONT_BODY = "Microsoft YaHei"
PURPLE = '#6A1B9A'

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

    def circle(s, x, y, size, color, text="", fs=24, tc=0xFFFFFF):
        c = s.Shapes.AddShape(9, x, y, size, size)
        c.Fill.ForeColor.RGB = color; c.Fill.Visible = True; c.Line.Visible = False
        if text:
            t2 = c.TextFrame.TextRange; t2.Text = text
            t2.Font.Size = fs; t2.Font.Color = tc; t2.Font.Name = FONT_TITLE
            t2.Font.Bold = True; t2.ParagraphFormat.Alignment = 2
        return c

    def rrect(s, x, y, w, h, color, text="", fs=18, tc=0xFFFFFF, bold=True, align=2, font=FONT_TITLE):
        r = s.Shapes.AddShape(5, x, y, w, h)
        r.Fill.ForeColor.RGB = color; r.Fill.Visible = True; r.Line.Visible = False
        if text:
            t2 = r.TextFrame.TextRange; t2.Text = text
            t2.Font.Size = fs; t2.Font.Color = tc; t2.Font.Name = font
            t2.Font.Bold = bold; t2.ParagraphFormat.Alignment = align
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

    def draw_table(s, e):
        rows, cols = e['rows'], e['cols']
        x,y,w,h = e['x'], e['y'], e['w'], e['h']
        dta = e['data']
        hdr = hex_to_bgr(e.get('header_color',PURPLE))
        row_h = h // rows; col_w = w // cols
        for r in range(rows):
            for c in range(cols):
                cx, cy = x+c*col_w, y+r*row_h
                val = dta[r][c] if r<len(dta) and c<len(dta[r]) else ""
                is_hdr = (r==0)
                bg_c = hdr if is_hdr else (0xFFFFFF if r%2==0 else hex_to_bgr('#F5F7FA'))
                fs = e.get('th_fs',15) if is_hdr else e.get('td_fs',14)
                tc = 0xFFFFFF if is_hdr else 0x333333
                al = 2 if c>0 else 1
                rrect(s, cx, cy, col_w, row_h, bg_c)
                if val:
                    txt(s, cx+4, cy+2, col_w-8, row_h-4, str(val), fs=fs, color=tc, bold=is_hdr, align=al,
                        font=FONT_TITLE if is_hdr else FONT_BODY, spacing=1.0)

    def draw_card_list_wide(s, e):
        items = e['items']; sy = e.get('start_y',92); ih = e.get('item_h',55); gap = e.get('gap',3)
        title_c = hex_to_bgr(e.get('title_color','#1A1A1A'))
        sub_c = hex_to_bgr(e.get('sub_color','#555555'))
        cols = ['#6A1B9A','#C62828','#E65100','#00695C','#8E24AA','#5C6BC0']
        for i, item in enumerate(items):
            y = sy + i*(ih+gap); col = hex_to_bgr(cols[i%len(cols)])
            circle(s, 100, y+8, 34, col, item['num'], fs=14)
            txt(s, 148, y+5, 300, 30, item['title'], fs=24, color=title_c, bold=True, font=FONT_TITLE)
            txt(s, 148, y+34, 680, 18, item['sub'], fs=18, color=sub_c, bold=False, font=FONT_BODY)

    def draw_cards_2x3(s, e):
        items = e['items']; sy = e.get('start_y',68); cw,ch=295,215; gx,gy=14,12
        for i, item in enumerate(items):
            r,c = i//3, i%3; x=22+c*(cw+gx); y=sy+r*(ch+gy)
            col = hex_to_bgr(item['color'])
            rect(s, x, y, cw, 4, col)
            txt(s, x+10, y+14, cw-20, 32, item['title'], fs=26, color=col, bold=True, align=1, font=FONT_TITLE)
            txt(s, x+10, y+52, cw-20, ch-62, item['desc'], fs=20, color=hex_to_bgr('#333333'), bold=False, align=1, font=FONT_BODY, spacing=1.4)

    def draw_cards_1x4_info(s, e):
        items = e['items']; sy = e.get('start_y',340); cw,ch=220,150; gap=8
        for i, item in enumerate(items):
            x=30+i*(cw+gap); col=hex_to_bgr(PURPLE)
            rect(s, x, sy, cw, 4, col)
            txt(s, x, sy+20, cw, 65, item['num'], fs=44, color=col, bold=True, align=2, font=FONT_TITLE)
            txt(s, x, sy+90, cw, 55, item['label'], fs=20, color=hex_to_bgr('#333333'), bold=False, align=2, font=FONT_BODY, spacing=1.3)

    ROUTERS = {
        'text': draw_text, 'image': draw_image, 'table': draw_table,
        'card_list_wide': draw_card_list_wide, 'cards_2x3': draw_cards_2x3,
        'cards_1x4_info': draw_cards_1x4_info,
    }

    for sd in slides_data:
        s = slide()
        for elem in sd.get('elements', []):
            etype = elem.get('type','text')
            router = ROUTERS.get(etype)
            if router:
                try: router(s, elem)
                except Exception as ex: print(f"WARN [{sd.get('title','?')}] {etype}: {ex}")

    pptx_path = os.path.join(OUT, "南京大学2025年本科招生录取数据全景.pptx")
    ppt.SaveAs(pptx_path)
    print(f"PPTX: {pptx_path} ({os.path.getsize(pptx_path):,} bytes)")
    pdf_path = os.path.join(OUT, "南京大学2025年本科招生录取数据全景.pdf")
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