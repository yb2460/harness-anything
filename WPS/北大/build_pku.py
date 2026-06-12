# -*- coding: utf-8 -*-
"""PKU Admissions PPT - scientific-slides: tables, images, large fonts."""
import os, json, pythoncom, win32com.client

OUT = os.environ.get("WPS_BUILD_OUT", os.path.join(os.path.dirname(os.path.abspath(__file__)), "work"))
JSON_PATH = os.path.join(OUT, "pku_data.json")
BG_IMAGE = os.path.join(OUT, "template_bg.png")

FONT_TITLE = "SimHei"
FONT_BODY = "Microsoft YaHei"
PKU_RED = '#8B0012'

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

    # ---- element types ----

    def draw_text(s, e):
        c = hex_to_bgr(e.get('color','#333333'))
        fnt = e.get('font', FONT_BODY)
        return txt(s, e['x'], e['y'], e['w'], e['h'], e['text'],
                   fs=e.get('fs',28), color=c, bold=e.get('bold',False),
                   align=e.get('align',1), font=fnt, spacing=e.get('line_spacing',1.3))

    def draw_line(s, e):
        return rect(s, e['x'], e['y'], e['w'], e['h'], hex_to_bgr(e.get('color',PKU_RED)))

    def draw_image(s, e):
        path = os.path.join(OUT, e['file'])
        if os.path.exists(path):
            img = s.Shapes.AddPicture(path, False, True, e['x'], e['y'], e['w'], e['h'])
        else:
            print(f"  WARN: image not found {path}")

    def draw_table(s, e):
        rows, cols = e['rows'], e['cols']
        x,y,w,h = e['x'], e['y'], e['w'], e['h']
        data = e['data']
        header_color = hex_to_bgr(e.get('header_color', PKU_RED))
        row_h = h // rows
        col_w = w // cols
        th_fs = e.get('th_fs', 16)
        td_fs = e.get('td_fs', 15)

        for r in range(rows):
            for c in range(cols):
                cx, cy = x + c*col_w, y + r*row_h
                val = data[r][c] if r < len(data) and c < len(data[r]) else ""
                is_header = (r == 0)
                bg_c = header_color if is_header else (0xFFFFFF if r%2==0 else hex_to_bgr('#F5F0EB'))
                fs = th_fs if is_header else td_fs
                tc = 0xFFFFFF if is_header else 0x333333
                bold = is_header
                align = 2 if c > 0 else 1
                rrect(s, cx, cy, col_w, row_h, bg_c)
                if val:
                    txt(s, cx+4, cy+2, col_w-8, row_h-4, str(val), fs=fs, color=tc, bold=bold, align=align, font=FONT_TITLE if is_header else FONT_BODY, spacing=1.0)

    def draw_card_list_wide(s, e):
        items = e['items']; sy = e.get('start_y',90); ih = e.get('item_h',50); gap = e.get('gap',2)
        title_c = hex_to_bgr(e.get('title_color','#1A1A1A'))
        sub_c = hex_to_bgr(e.get('sub_color','#555555'))
        cols = ['#8B0012','#B22222','#D2691E','#2E5090','#3A6B2C','#8B1A6B','#B8860B','#8B0012','#B22222']
        for i, item in enumerate(items):
            y = sy + i*(ih+gap); col = hex_to_bgr(cols[i%len(cols)])
            circle(s, 100, y+8, 34, col, item['num'], fs=14)
            txt(s, 148, y+5, 300, 30, item['title'], fs=24, color=title_c, bold=True, font=FONT_TITLE)
            txt(s, 148, y+34, 650, 18, item['sub'], fs=18, color=sub_c, bold=False, font=FONT_BODY)

    def draw_tagline_bar(s, e):
        col = hex_to_bgr(e.get('color',PKU_RED))
        x,y,w,h = e['x'],e['y'],e['w'],e['h']
        rrect(s, x, y, w, h, col)
        txt(s, x+10, y+5, w-20, h-10, e['text'], fs=15, color=0xFFFFFF, bold=False, align=2, font=FONT_BODY)

    def draw_cards_2x3(s, e):
        items = e['items']; sy = e.get('start_y',82); cw,ch=295,210; gx,gy=14,12
        for i, item in enumerate(items):
            r,c = i//3, i%3; x=22+c*(cw+gx); y=sy+r*(ch+gy)
            col = hex_to_bgr(item['color'])
            rect(s, x, y, cw, 4, col)
            txt(s, x+10, y+14, cw-20, 34, item['title'], fs=26, color=col, bold=True, align=1, font=FONT_TITLE)
            txt(s, x+10, y+54, cw-20, ch-64, item['desc'], fs=22, color=hex_to_bgr('#333333'), bold=False, align=1, font=FONT_BODY, spacing=1.4)

    def draw_cards_2x2_four(s, e):
        items = e['items']; sy = e.get('start_y',86); cw,ch=440,210; gx,gy=16,12
        for i, item in enumerate(items):
            r,c = i//2, i%2; x=30+c*(cw+gx); y=sy+r*(ch+gy)
            col = hex_to_bgr(item['color'])
            rect(s, x, y, cw, 4, col)
            txt(s, x+14, y+16, cw-28, 34, item['title'], fs=28, color=col, bold=True, align=1, font=FONT_TITLE)
            txt(s, x+14, y+56, cw-28, ch-68, item['desc'], fs=22, color=hex_to_bgr('#333333'), bold=False, align=1, font=FONT_BODY, spacing=1.45)

    def draw_cards_1x4_info(s, e):
        items = e['items']; sy = e.get('start_y',360); cw,ch=220,140; gap=8
        for i, item in enumerate(items):
            x=30+i*(cw+gap); col=hex_to_bgr(PKU_RED)
            rect(s, x, sy, cw, 4, col)
            txt(s, x, sy+18, cw, 60, item['num'], fs=44, color=col, bold=True, align=2, font=FONT_TITLE)
            txt(s, x, sy+85, cw, 48, item['label'], fs=20, color=hex_to_bgr('#333333'), bold=False, align=2, font=FONT_BODY, spacing=1.3)

    def draw_cards_1x3_big(s, e):
        items = e['items']; sy = e.get('start_y',80); cw,ch=290,240; gap=16
        for i, item in enumerate(items):
            x=36+i*(cw+gap); col=hex_to_bgr(item['color'])
            rect(s, x, sy, cw, 5, col)
            txt(s, x+14, sy+20, cw-28, 64, item['title'], fs=28, color=col, bold=True, align=1, font=FONT_TITLE, spacing=1.2)
            rect(s, x+14, sy+92, cw-28, 1, hex_to_bgr('#CCCCCC'))
            txt(s, x+14, sy+104, cw-28, ch-118, item['desc'], fs=22, color=hex_to_bgr('#333333'), bold=False, align=1, font=FONT_BODY, spacing=1.55)

    def draw_card_row_5(s, e):
        items = e['items']; sy = e.get('start_y',80); cw,ch=174,340; gap=5
        for i, item in enumerate(items):
            x=25+i*(cw+gap); col=hex_to_bgr(item['color'])
            rect(s, x, sy, cw, 4, col)
            circle(s, x+52, sy+22, 52, col, item['icon'], fs=22, tc=0xFFFFFF)
            txt(s, x+8, sy+86, cw-16, 44, item['title'], fs=20, color=col, bold=True, align=2, font=FONT_TITLE, spacing=1.2)
            rect(s, x+14, sy+138, cw-28, 1, hex_to_bgr('#CCCCCC'))
            txt(s, x+8, sy+146, cw-16, ch-158, item['desc'], fs=18, color=hex_to_bgr('#444444'), bold=False, align=2, font=FONT_BODY, spacing=1.45)

    ROUTERS = {
        'text': draw_text, 'line': draw_line, 'image': draw_image, 'table': draw_table,
        'card_list_wide': draw_card_list_wide, 'tagline_bar': draw_tagline_bar,
        'cards_2x3': draw_cards_2x3, 'cards_2x2_four': draw_cards_2x2_four,
        'cards_1x4_info': draw_cards_1x4_info, 'cards_1x3_big': draw_cards_1x3_big,
        'card_row_5': draw_card_row_5,
    }

    for sd in slides_data:
        s = slide()
        for elem in sd.get('elements', []):
            etype = elem.get('type','text')
            router = ROUTERS.get(etype)
            if router:
                try: router(s, elem)
                except Exception as ex: print(f"WARN [{sd.get('title','?')}] {etype}: {ex}")
            else: print(f"WARN [{sd.get('title','?')}] unknown: {etype}")

    pptx_path = os.path.join(OUT, "北京大学2025年本科招生录取数据全景.pptx")
    ppt.SaveAs(pptx_path)
    print(f"PPTX: {pptx_path} ({os.path.getsize(pptx_path):,} bytes)")
    pdf_path = os.path.join(OUT, "北京大学2025年本科招生录取数据全景.pdf")
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