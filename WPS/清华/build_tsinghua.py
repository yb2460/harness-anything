# -*- coding: utf-8 -*-
"""Tsinghua Admissions PPT - WPS COM automation."""
import os, json, pythoncom, win32com.client

OUT = os.environ.get("WPS_BUILD_OUT", os.path.join(os.path.dirname(os.path.abspath(__file__)), "work"))
JSON_PATH = os.path.join(OUT, "tsinghua_data.json")
BG_IMAGE = os.path.join(OUT, "template_bg.png")

FONT_TITLE = "SimHei"
FONT_BODY = "Microsoft YaHei"

def hex_to_bgr(h):
    h = h.lstrip('#')
    r, g, b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
    return (b<<16)|(g<<8)|r

def run():
    with open(JSON_PATH,'r',encoding='utf-8') as f:
        data = json.load(f)
    W, H = data['canvas']['w'], data['canvas']['h']
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

    def txt(s, x, y, w, h, text, fs=18, color=0x333333, bold=False, align=1, font=FONT_BODY, spacing=1.4):
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

    def rrect(s, x, y, w, h, color, text="", fs=18, tc=0xFFFFFF, bold=True, align=2, font=FONT_TITLE):
        r = s.Shapes.AddShape(5, x, y, w, h)
        r.Fill.ForeColor.RGB = color; r.Fill.Visible = True; r.Line.Visible = False
        if text:
            t2 = r.TextFrame.TextRange; t2.Text = text
            t2.Font.Size = fs; t2.Font.Color = tc; t2.Font.Name = font
            t2.Font.Bold = bold; t2.ParagraphFormat.Alignment = align
        return r

    def circle(s, x, y, size, color, text="", fs=24, tc=0xFFFFFF):
        c = s.Shapes.AddShape(9, x, y, size, size)
        c.Fill.ForeColor.RGB = color; c.Fill.Visible = True; c.Line.Visible = False
        if text:
            t2 = c.TextFrame.TextRange; t2.Text = text
            t2.Font.Size = fs; t2.Font.Color = tc; t2.Font.Name = FONT_TITLE
            t2.Font.Bold = True; t2.ParagraphFormat.Alignment = 2
        return c

    def draw_text(s, e):
        c = hex_to_bgr(e.get('color','#333333'))
        fnt = e.get('font', FONT_BODY)
        return txt(s, e['x'], e['y'], e['w'], e['h'], e['text'],
                   fs=e.get('fs',18), color=c, bold=e.get('bold',False),
                   align=e.get('align',1), font=fnt, spacing=e.get('line_spacing',1.4))

    def draw_line(s, e):
        return rect(s, e['x'], e['y'], e['w'], e['h'], hex_to_bgr(e.get('color','#660874')))

    def draw_card_list_wide(s, e):
        items = e['items']; sy = e.get('start_y',85); ih = e.get('item_h',48); gap = e.get('gap',2)
        # Support color overrides: for S2 all-black
        title_c = hex_to_bgr(e.get('title_color','#FFFFFF'))
        sub_c = hex_to_bgr(e.get('sub_color','#AABBCC'))
        sep_c = hex_to_bgr(e.get('sep_color','#334455'))
        cols = ['#660874','#8B1A6B','#D2691E','#2E5090','#3A6B2C','#B8860B','#B22222','#660874','#8B1A6B','#D2691E']
        for i, item in enumerate(items):
            y = sy + i*(ih+gap); col = hex_to_bgr(cols[i%len(cols)])
            circle(s, 150, y+8, 30, col, item['num'], fs=11)
            txt(s, 195, y+6, 260, 26, item['title'], fs=19, color=title_c, bold=True, font=FONT_TITLE)
            txt(s, 195, y+34, 600, 15, item['sub'], fs=14, color=sub_c, bold=False, font=FONT_BODY)
            if i<len(items)-1: rect(s, 195, (y+ih)-2, 600, 1, sep_c)

    def draw_tagline_bar(s, e):
        col = hex_to_bgr(e.get('color','#660874'))
        x,y,w,h = e['x'],e['y'],e['w'],e['h']
        rrect(s, x, y, w, h, col)
        txt(s, x+10, y+6, w-20, h-12, e['text'], fs=14, color=0xFFFFFF, bold=False, align=2, font=FONT_BODY)

    def draw_cards_2x3(s, e):
        items = e['items']; sy = e.get('start_y',72); cw,ch=290,208; gx,gy=8,8
        for i, item in enumerate(items):
            r,c = i//3, i%3; x=30+c*(cw+gx); y=sy+r*(ch+gy)
            col = hex_to_bgr(item['color'])
            rrect(s, x, y, cw, ch, hex_to_bgr('#F2F4F7'))
            rect(s, x, y, cw, 4, col)
            txt(s, x+14, y+12, cw-28, 28, item['title'], fs=18, color=col, bold=True, align=1, font=FONT_TITLE)
            txt(s, x+14, y+44, cw-28, ch-54, item['desc'], fs=14, color=hex_to_bgr('#333333'), bold=False, align=1, font=FONT_BODY, spacing=1.3)

    def draw_cards_2x2_four(s, e):
        items = e['items']; sy = e.get('start_y',72); cw,ch=440,220; gx,gy=15,10
        for i, item in enumerate(items):
            r,c = i//2, i%2; x=30+c*(cw+gx); y=sy+r*(ch+gy)
            col = hex_to_bgr(item['color'])
            rrect(s, x, y, cw, ch, hex_to_bgr('#F2F4F7'))
            rect(s, x, y, cw, 4, col)
            txt(s, x+16, y+14, cw-32, 28, item['title'], fs=20, color=col, bold=True, align=1, font=FONT_TITLE)
            txt(s, x+16, y+48, cw-32, ch-60, item['desc'], fs=15, color=hex_to_bgr('#333333'), bold=False, align=1, font=FONT_BODY, spacing=1.35)

    def draw_cards_1x4_info(s, e):
        items = e['items']; sy = e.get('start_y',390); cw,ch=220,120; gap=8
        for i, item in enumerate(items):
            x=30+i*(cw+gap); col=hex_to_bgr('#660874')
            rrect(s, x, sy, cw, ch, hex_to_bgr('#F2F4F7'))
            rect(s, x, sy, cw, 4, col)
            txt(s, x, sy+18, cw, 48, item['num'], fs=30, color=col, bold=True, align=2, font=FONT_TITLE)
            txt(s, x, sy+68, cw, 48, item['label'], fs=15, color=hex_to_bgr('#333333'), bold=False, align=2, font=FONT_BODY, spacing=1.3)

    def draw_cards_1x3_big(s, e):
        items = e['items']; sy = e.get('start_y',72); cw,ch=285,250; gap=14
        for i, item in enumerate(items):
            x=40+i*(cw+gap); col=hex_to_bgr(item['color'])
            rrect(s, x, sy, cw, ch, hex_to_bgr('#F2F4F7'))
            rect(s, x, sy, cw, 5, col)
            txt(s, x+16, sy+20, cw-32, 58, item['title'], fs=20, color=col, bold=True, align=1, font=FONT_TITLE, spacing=1.2)
            rect(s, x+20, sy+86, cw-40, 1, hex_to_bgr('#D0D4D8'))
            txt(s, x+16, sy+98, cw-32, ch-112, item['desc'], fs=16, color=hex_to_bgr('#333333'), bold=False, align=1, font=FONT_BODY, spacing=1.4)

    def draw_card_row_5(s, e):
        items = e['items']; sy = e.get('start_y',72); cw,ch=172,350; gap=5
        for i, item in enumerate(items):
            x=25+i*(cw+gap); col=hex_to_bgr(item['color'])
            rrect(s, x, sy, cw, ch, hex_to_bgr('#F2F4F7'))
            rect(s, x, sy, cw, 4, col)
            circle(s, x+51, sy+25, 50, col, item['icon'], fs=20, tc=0xFFFFFF)
            txt(s, x+10, sy+86, cw-20, 40, item['title'], fs=15, color=col, bold=True, align=2, font=FONT_TITLE, spacing=1.2)
            rect(s, x+16, sy+130, cw-32, 1, hex_to_bgr('#D0D4D8'))
            txt(s, x+10, sy+140, cw-20, ch-152, item['desc'], fs=14, color=hex_to_bgr('#555555'), bold=False, align=2, font=FONT_BODY, spacing=1.35)

    ROUTERS = {
        'text': draw_text, 'line': draw_line,
        'card_list_wide': draw_card_list_wide,
        'tagline_bar': draw_tagline_bar,
        'cards_2x3': draw_cards_2x3,
        'cards_2x2_four': draw_cards_2x2_four,
        'cards_1x4_info': draw_cards_1x4_info,
        'cards_1x3_big': draw_cards_1x3_big,
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

    pptx_path = os.path.join(OUT, "清华大学2025年本科招生录取数据全景.pptx")
    ppt.SaveAs(pptx_path)
    print(f"PPTX: {pptx_path} ({os.path.getsize(pptx_path):,} bytes)")
    pdf_path = os.path.join(OUT, "清华大学2025年本科招生录取数据全景.pdf")
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