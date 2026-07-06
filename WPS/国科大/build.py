# -*- coding: utf-8 -*-
"""中科院大学2025招生PPT — WPS COM automation."""
import os, json, pythoncom, win32com.client

OUT = r"D:\A-资料\A-claudewenjian\PPT制作\中国科学院大学\work"
JSON_PATH = os.path.join(OUT, "data.json")
BG_IMAGE = os.path.join(OUT, "template_bg.png")
FONT_TITLE = "SimHei"
FONT_BODY = "Microsoft YaHei"

def hex2bgr(h):
    h = h.lstrip('#')
    r,g,b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
    return (b<<16)|(g<<8)|r

def run():
    with open(JSON_PATH,'r',encoding='utf-8') as f:
        data = json.load(f)
    W,H = data['canvas']['w'], data['canvas']['h']

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
        if os.path.exists(BG_IMAGE):
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
        c = hex2bgr(e.get('color','#333333'))
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
        hdr = hex2bgr(e.get('header_color','#003B6F'))
        row_h = h // rows; col_w = w // cols
        for r in range(rows):
            for c in range(cols):
                cx, cy = x+c*col_w, y+r*row_h
                val = dta[r][c] if r<len(dta) and c<len(dta[r]) else ""
                is_hdr = (r==0)
                bg_c = hdr if is_hdr else (0xFFFFFF if r%2==0 else hex2bgr('#F5F7FA'))
                fs = e.get('th_fs',15) if is_hdr else e.get('td_fs',14)
                tc = 0xFFFFFF if is_hdr else 0x333333
                al = 2 if c>0 else 1
                rrect(s, cx, cy, col_w, row_h, bg_c)
                if val:
                    txt(s, cx+4, cy+2, col_w-8, row_h-4, str(val), fs=fs, color=tc, bold=is_hdr, align=al,
                        font=FONT_TITLE if is_hdr else FONT_BODY, spacing=1.0)

    def draw_card_list_wide(s, e):
        items = e['items']; sy = e.get('start_y',90); ih = e.get('item_h',44); gap = e.get('gap',2)
        title_c = hex2bgr(e.get('title_color','#1A1A1A'))
        sub_c = hex2bgr(e.get('sub_color','#555555'))
        cols = ['#003B6F','#C41230','#C8922A','#007A6E','#E0682A','#5C2D82','#003B6F','#C41230','#C8922A','#007A6E','#E0682A','#5C2D82','#003B6F']
        for i, item in enumerate(items):
            y = sy + i*(ih+gap); col = hex2bgr(cols[i%len(cols)])
            circle(s, 100, y+6, 32, col, item['num'], fs=13)
            txt(s, 148, y+4, 300, 28, item['title'], fs=22, color=title_c, bold=True, font=FONT_TITLE)
            txt(s, 148, y+28, 680, 16, item['sub'], fs=14, color=sub_c, bold=False, font=FONT_BODY)

    def draw_tagline_bar(s, e):
        col = hex2bgr(e.get('color','#003B6F'))
        x,y,w,h = e['x'],e['y'],e['w'],e['h']
        rrect(s, x, y, w, h, col)
        txt(s, x+10, y+5, w-20, h-10, e['text'], fs=14, color=0xFFFFFF, bold=False, align=2, font=FONT_BODY)

    def draw_cards_2x3(s, e):
        items = e['items']; sy = e.get('start_y',68); cw,ch=295,225; gx,gy=14,12
        for i, item in enumerate(items):
            r,c = i//3, i%3; x=22+c*(cw+gx); y=sy+r*(ch+gy)
            col = hex2bgr(item['color'])
            rect(s, x, y, cw, 4, col)
            txt(s, x+10, y+16, cw-20, 32, item['title'], fs=26, color=col, bold=True, align=1, font=FONT_TITLE)
            txt(s, x+10, y+56, cw-20, ch-66, item['desc'], fs=17, color=hex2bgr('#333333'), bold=False, align=1, font=FONT_BODY, spacing=1.35)

    ROUTERS = {
        'text': draw_text, 'image': draw_image, 'table': draw_table,
        'card_list_wide': draw_card_list_wide, 'tagline_bar': draw_tagline_bar,
        'cards_2x3': draw_cards_2x3,
    }

    for sd in data['slides']:
        s = slide()
        for elem in sd.get('elements', []):
            etype = elem.get('type','text')
            router = ROUTERS.get(etype)
            if router:
                try: router(s, elem)
                except Exception as ex:
                    print(f"WARN [{sd.get('title','?')}] {etype}: {ex}")

    pptx_path = os.path.join(OUT, "国科大2025本科招生数据全景.pptx")
    ppt.SaveAs(pptx_path)
    print(f"PPTX: {pptx_path} ({os.path.getsize(pptx_path):,} bytes)")
    pdf_path = os.path.join(OUT, "国科大2025本科招生数据全景.pdf")
    try:
        ppt.SaveAs(pdf_path, 32)
        print(f"PDF: {pdf_path} ({os.path.getsize(pdf_path):,} bytes)")
    except: print("PDF failed")
    ppt.Close()
    try: app.Quit()
    except: pass
    print("Done!")

if __name__ == "__main__":
    run()
