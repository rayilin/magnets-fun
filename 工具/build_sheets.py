# -*- coding: utf-8 -*-
"""學習單 2 份 + 測驗卷"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = r'C:\Users\rayi\Documents\自然科學教學專欄\磁鐵好好玩'
IMG = os.path.join(ROOT, '教學網頁', 'media', 'img')
OUT = os.path.join(ROOT, '教案')
INK = RGBColor(0x35, 0x54, 0x4A)
ORANGE = RGBColor(0xC0, 0x6A, 0x2B)
GREY = RGBColor(0x6B, 0x7B, 0x72)
FONT = '微軟正黑體'


def setup(doc, margin=1.6):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    for k in ('top_margin', 'bottom_margin', 'left_margin', 'right_margin'):
        setattr(sec, k, Cm(margin))
    st = doc.styles['Normal']
    st.font.name = FONT
    st.font.size = Pt(11)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    st.paragraph_format.space_after = Pt(3)
    st.paragraph_format.line_spacing = 1.3


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    el = OxmlElement('w:shd'); el.set(qn('w:val'), 'clear'); el.set(qn('w:fill'), hexcolor)
    tcPr.append(el)


def para(doc, text='', size=11, bold=False, color=None, align=None, space_before=0, space_after=3, indent=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    r = p.add_run(text)
    r.font.size = Pt(size); r.font.bold = bold
    r.font.name = FONT
    r._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    if color:
        r.font.color.rgb = color
    return p


def cell_text(cell, text, size=10, bold=False, align='c'):
    cell.text = ''
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(1); p.paragraph_format.space_after = Pt(1)
    if align == 'c':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.font.size = Pt(size); r.font.bold = bold; r.font.name = FONT
    r._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)


def cell_img(cell, code, width=1.7):
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(1); p.paragraph_format.space_after = Pt(1)
    path = os.path.join(IMG, f'm{code}.jpg')
    if os.path.exists(path):
        p.add_run().add_picture(path, width=Cm(width))


def blank_line(doc, n=1):
    for _ in range(n):
        para(doc, '＿' * 42, color=RGBColor(0xC8, 0xC0, 0xB0), space_after=6)


def header_block(doc, title, subtitle):
    para(doc, '翰林版三年級上學期 自然科學｜第 2 單元　磁鐵好好玩', 9.5, color=ORANGE, align=WD_ALIGN_PARAGRAPH.CENTER)
    para(doc, title, 20, bold=True, color=INK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=0)
    para(doc, subtitle, 10.5, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    para(doc, '班級：＿＿＿＿＿　座號：＿＿＿　姓名：＿＿＿＿＿＿＿＿　日期：＿＿＿月＿＿＿日', 10.5, space_after=8)


# ══════════════════ 學習單① ══════════════════
d1 = Document(); setup(d1)
header_block(d1, '學習單①　磁鐵能吸引什麼？', '先猜一猜，再用磁鐵測測看（每樣測三次）')
para(d1, '任務一：預測與實測', 12, bold=True, color=INK, space_before=4)
para(d1, '在「預測」欄打 ✓（會吸）或 ✗（不會吸）；實測後把結果寫在「結果」欄。', 9.5, color=GREY, space_after=6)

ITEMS = [('004', '長尾夾'), ('008', '鐵釘'), ('005', '鐵尺'), ('012', '迴紋針'), ('017', '鐵湯匙'), ('036', '鐵罐'),
         ('002', '橡皮擦'), ('013', '塑膠尺'), ('009', '硬幣'), ('020', '紙張'), ('022', '木塊'), ('074', '玻璃杯')]
t = d1.add_table(rows=len(ITEMS) + 1, cols=6)
t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
heads = ['物品', '名稱', '我的預測', '第 1 次', '第 2 次', '結果']
for j, hh in enumerate(heads):
    cell_text(t.cell(0, j), hh, 10, True); shade(t.cell(0, j), 'F3EEE1')
widths = [2.2, 3.0, 2.6, 2.4, 2.4, 4.4]
for i, (code, name) in enumerate(ITEMS, 1):
    cell_img(t.cell(i, 0), code, 1.25)
    cell_text(t.cell(i, 1), name, 10.5, align='l')
    for j in range(2, 6):
        cell_text(t.cell(i, j), '', 10.5)
for j, w in enumerate(widths):
    for row in t.rows:
        row.cells[j].width = Cm(w)

d1.add_page_break()
para(d1, '任務二：我的發現', 12, bold=True, color=INK, space_before=12)
for q in ['1. 會被吸住的物品，共同點是什麼？', '2. 不會被吸住的物品，共同點是什麼？',
          '3. 我的預測和實測結果一樣嗎？如果不一樣，是哪一樣？為什麼？']:
    para(d1, q, 11, space_before=4, space_after=2)
    blank_line(d1, 1)

para(d1, '挑戰題：硬幣摸起來又硬又冰，為什麼磁鐵吸不住它？', 11, bold=True, color=ORANGE, space_before=8)
blank_line(d1, 2)

# ══════════════════ 學習單② ══════════════════
d2 = Document(); setup(d2)
header_block(d2, '學習單②　磁極與相斥相吸', '找出磁鐵的兩極，再試試看兩個磁鐵靠近會怎樣')

para(d2, '任務一：標出磁極', 12, bold=True, color=INK, space_before=2)
para(d2, '在磁鐵兩端寫上 N 或 S，並把 N 極塗紅色、S 極塗藍色。', 9.5, color=GREY, space_after=6)
t = d2.add_table(rows=2, cols=3)
t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
for j, (code, name) in enumerate([('025', '長條形磁鐵'), ('026', '馬蹄形磁鐵'), ('027', '圓形磁鐵')]):
    cell_img(t.cell(0, j), code, 3.2)
    cell_text(t.cell(1, j), name, 10)
for j in range(3):
    for row in t.rows:
        row.cells[j].width = Cm(5.9)

para(d2, '任務二：兩個磁鐵靠近會怎樣？', 12, bold=True, color=INK, space_before=12)
para(d2, '先預測（相吸／相斥），再做實驗，把結果圈起來。', 9.5, color=GREY, space_after=6)
t = d2.add_table(rows=4, cols=4)
t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
for j, hh in enumerate(['兩個磁鐵靠近的方式', '我的預測', '實驗結果', '畫下來（箭頭）']):
    cell_text(t.cell(0, j), hh, 10, True); shade(t.cell(0, j), 'F3EEE1')
rows = [('N 極 ←→ N 極', '相吸　／　相斥', '相吸　／　相斥'),
        ('S 極 ←→ S 極', '相吸　／　相斥', '相吸　／　相斥'),
        ('N 極 ←→ S 極', '相吸　／　相斥', '相吸　／　相斥')]
for i, r in enumerate(rows, 1):
    cell_text(t.cell(i, 0), r[0], 10.5)
    cell_text(t.cell(i, 1), r[1], 10)
    cell_text(t.cell(i, 2), r[2], 10)
    cell_text(t.cell(i, 3), '', 10)
for j, w in enumerate([5.0, 4.2, 4.2, 4.4]):
    for row in t.rows:
        row.cells[j].width = Cm(w)

para(d2, '任務三：磁鐵會指南北嗎？', 12, bold=True, color=INK, space_before=12)
para(d2, '把磁鐵吊起來（或放在水面上）讓它自由轉動，等它停下來，記錄兩個磁極指向的方向。', 9.5, color=GREY, space_after=4)
para(d2, 'N 極指向：＿＿＿＿＿＿　　S 極指向：＿＿＿＿＿＿', 11, space_after=8)

para(d2, '任務四：寫下結論', 12, bold=True, color=INK, space_before=6)
for q in ['1. 磁鐵的磁力在哪裡最強？', '2. 相同的磁極靠近會＿＿＿＿＿＿，不同的磁極靠近會＿＿＿＿＿＿。',
          '3. 寫出兩個家裡或教室裡「有磁鐵」的東西，並說說它用磁鐵做什麼。']:
    para(d2, q, 11, space_before=4, space_after=2)
    blank_line(d2, 1)

# ══════════════════ 測驗卷 ══════════════════
d3 = Document(); setup(d3)
para(d3, '翰林版三年級上學期 自然科學｜第 2 單元', 9.5, color=ORANGE, align=WD_ALIGN_PARAGRAPH.CENTER)
para(d3, '磁鐵好好玩　測驗卷', 20, bold=True, color=INK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
para(d3, '（10 題，每題 10 分，共 100 分）', 10, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
para(d3, '班級：＿＿＿＿＿　座號：＿＿＿　姓名：＿＿＿＿＿＿＿＿　得分：＿＿＿＿', 11, space_after=10)

QUIZ = [
    ('用磁鐵靠近下列哪一種物品，會被吸住？', ['橡皮擦', '鐵釘', '塑膠尺', '紙張'], 1,
     '鐵釘是鐵做的，磁鐵吸引鐵製品。'),
    ('下列哪一種物品「不會」被磁鐵吸引？', ['迴紋針', '鐵湯匙', '硬幣', '鐵尺'], 2,
     '臺灣硬幣主要是銅、鎳、鋅，不是鐵。'),
    ('磁鐵的哪一個部位磁力最強？', ['正中央', '兩端', '整塊都一樣強', '背面'], 1,
     '磁力集中在兩端，這兩端是磁極。'),
    ('把磁鐵慢慢遠離迴紋針，會發生什麼事？', ['吸力變強', '吸力不變', '吸力變弱，最後吸不到', '迴紋針被推開'], 2,
     '距離越遠，磁力越弱。'),
    ('長條形磁鐵的 N 極和 S 極分別在哪裡？', ['都在中間', '都在同一端', '分別在兩端', '只有一端有磁極'], 2,
     '每一塊磁鐵都有兩個磁極，分別在兩端。'),
    ('把兩個磁鐵的 N 極靠近 N 極，會怎樣？', ['互相吸引', '互相排斥', '沒有任何反應', '合成一個磁鐵'], 1,
     '同極相斥。'),
    ('把磁鐵吊起來讓它自由轉動，停止後 N 極會指向哪裡？', ['東方', '南方', '北方', '正上方'], 2,
     'N 極指北、S 極指南。'),
    ('把一塊長條磁鐵折成兩段，每一段會怎樣？', ['只有一端有磁極', '變成沒有磁性', '兩段各有 N 極和 S 極', '會互相排斥到很遠'], 2,
     '斷成幾段，每段都是完整磁鐵。'),
    ('下列哪一個生活用品「沒有」用到磁鐵？', ['冰箱門的磁條', '鉛筆盒的蓋子', '風車', '黑板上的磁鐵'], 2,
     '風車靠風轉動，和磁鐵無關。'),
    ('用膠帶黏起頭髮，和磁鐵吸引鐵釘，原理一樣嗎？', ['一樣，都是磁力', '不一樣，膠帶靠靜電', '一樣，都是靜電', '不一樣，磁鐵是靜電'], 1,
     '膠帶黏頭髮是靜電，不是磁力。'),
]
for i, (q, o, a, e) in enumerate(QUIZ, 1):
    p = para(d3, f'{i}. {q}', 11.5, bold=True, space_before=6, space_after=2)
    para(d3, '　　' + '　　'.join(f'(　) {v}' for v in o), 11, space_after=2)

d3.add_page_break()
para(d3, '答案與解析（教師用）', 16, bold=True, color=ORANGE, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
t = d3.add_table(rows=len(QUIZ) + 1, cols=3)
t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
for j, hh in enumerate(['題號', '答案', '解析']):
    cell_text(t.cell(0, j), hh, 10.5, True); shade(t.cell(0, j), 'F3EEE1')
for i, (q, o, a, e) in enumerate(QUIZ, 1):
    cell_text(t.cell(i, 0), str(i), 10.5)
    cell_text(t.cell(i, 1), 'ABCD'[a], 10.5, True)
    cell_text(t.cell(i, 2), e, 10, align='l')
for j, w in enumerate([1.6, 1.8, 14.0]):
    for row in t.rows:
        row.cells[j].width = Cm(w)

for name, d in [('學習單1_磁鐵能吸引什麼.docx', d1),
                ('學習單2_磁極與相斥相吸.docx', d2),
                ('測驗卷_磁鐵好好玩.docx', d3)]:
    p = os.path.join(OUT, name)
    d.save(p)
    print('產出：', p)
