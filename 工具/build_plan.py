# -*- coding: utf-8 -*-
"""產生「磁鐵好好玩」教案 docx（4 節）與 2 份學習單、測驗卷"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = r'C:\Users\rayi\Documents\自然科學教學專欄\磁鐵好好玩'
IMG = os.path.join(ROOT, '教學網頁', 'media', 'img')
OUT = os.path.join(ROOT, '教案')
INK = RGBColor(0x35, 0x54, 0x4A)
ORANGE = RGBColor(0xC0, 0x6A, 0x2B)
GREY = RGBColor(0x6B, 0x7B, 0x72)
FONT = '微軟正黑體'


def setup(doc, margin=2.0):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    for k in ('top_margin', 'bottom_margin', 'left_margin', 'right_margin'):
        setattr(sec, k, Cm(margin))
    st = doc.styles['Normal']
    st.font.name = FONT
    st.font.size = Pt(11.5)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    st.paragraph_format.space_after = Pt(4)
    st.paragraph_format.line_spacing = 1.35


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    el = OxmlElement('w:shd')
    el.set(qn('w:val'), 'clear'); el.set(qn('w:fill'), hexcolor)
    tcPr.append(el)


def para(doc, text='', size=11.5, bold=False, color=None, align=None, space_before=0, space_after=4, indent=None):
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


def h(doc, text, size=15, space_before=14):
    p = para(doc, text, size=size, bold=True, color=INK, space_before=space_before, space_after=6)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), '10')
    bottom.set(qn('w:space'), '3'); bottom.set(qn('w:color'), 'C06A2B')
    pbdr.append(bottom); pPr.append(pbdr)
    return p


def table(doc, rows, widths=None, header=True, size=10.5, aligns=None):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = t.cell(i, j)
            cell.text = ''
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.space_before = Pt(2)
            if aligns and aligns[j] == 'c':
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(str(val))
            r.font.size = Pt(size); r.font.name = FONT
            r._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
            if header and i == 0:
                r.font.bold = True
                r.font.color.rgb = INK
                shade(cell, 'F3EEE1')
    if widths:
        for j, w in enumerate(widths):
            for row in t.rows:
                row.cells[j].width = Cm(w)
    return t


# ════════════════════════════════════════════════════════
# 教案
# ════════════════════════════════════════════════════════
doc = Document()
setup(doc)

para(doc, '翰林版 三年級上學期 自然科學', 11, color=ORANGE, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=10)
para(doc, '磁鐵好好玩', 26, bold=True, color=INK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
para(doc, '教學活動設計（4 節課）', 13, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14)

h(doc, '一、基本資料', space_before=4)
table(doc, [
    ['單元名稱', '第 2 單元　磁鐵好好玩', '教學年級', '三年級'],
    ['教材來源', '翰林版自然科學 3 上', '教學節數', '4 節（每節 40 分鐘）'],
    ['教學者', '　　　　　　　', '教學日期', '10/26（一）、10/29（四）、11/2（一）、11/5（四）'],
    ['教學地點', '自然教室', '教學對象', '三年級學生（分組 4 人一組，共 6 組）'],
    ['設計者', '　　　　　　　', '班級', '　　年　　班'],
], widths=[2.6, 7.2, 2.6, 4.6])

h(doc, '二、設計理念')
para(doc, '磁鐵是三年級學生最容易「玩」也最容易「誤解」的單元。孩子常憑外觀判斷：金屬做的硬幣、'
          '鋁罐看起來像鐵，就以為會被磁鐵吸住；也會把「膠帶黏頭髮」的靜電現象誤認成磁力。')
para(doc, '因此本單元以「預測 → 實測 → 記錄 → 歸納」的科學探究流程為主軸：每節課都先讓學生猜，'
          '再用真實物品驗證，最後由學生自己說出規律。四節課的順序是「吸引什麼 → 磁力在哪裡 → '
          '磁極怎麼互動 → 生活中怎麼用」，讓概念一層一層疊上去。')
para(doc, '評量以學習單與實作記錄為主、紙筆測驗為輔，並搭配教學網頁的互動實驗、10 題線上測驗與'
          '兩人搶答對戰，讓學生在課後也能自己複習。')

h(doc, '三、學習重點')
para(doc, '（一）學習表現', 11.5, bold=True, color=ORANGE, space_before=6)
for s in ['ti-Ⅱ-1　能在指導下觀察日常生活現象的規律性，並運用想像力與好奇心，了解及描述自然環境的現象。',
          'po-Ⅱ-1　能從日常經驗、學習活動、自然環境，進行觀察，進而能察覺問題。',
          'pe-Ⅱ-1　能了解一個因素改變可能造成的影響，進而預測活動的大致結果。',
          'pe-Ⅱ-2　能正確安全操作適合學習階段的物品、器材儀器，並能觀察和記錄。',
          'pa-Ⅱ-1　能運用簡單的分類、比較、分析與推論等方法，整理已有的資訊或數據。',
          'pc-Ⅱ-1　能專注聆聽同學報告，提出疑問或意見，並對探究過程或結果進行檢討。']:
    para(doc, s, indent=0.5, space_after=2)

para(doc, '（二）學習內容', 11.5, bold=True, color=ORANGE, space_before=8)
para(doc, 'INe-Ⅱ-7　磁鐵具有兩極，同極相斥、異極相吸；磁鐵會吸引含鐵的物體。磁力強弱可由吸起'
          '含鐵物質數量多寡得知。', indent=0.5)
para(doc, '（本單元另連結生活應用：磁鐵在生活中的用途，如門擋、冰箱磁條、磁性文具與磁鐵玩具。）',
     10.5, color=GREY, indent=0.5)

h(doc, '四、學習目標')
for i, s in enumerate([
    '能預測並實測哪些物品會被磁鐵吸引，並歸納出「磁鐵吸引鐵製品」。',
    '能指出磁鐵兩端磁力最強，說明距離越遠磁力越弱，並找出磁鐵的 N 極與 S 極。',
    '能說出同極相斥、異極相吸，並做出會自由轉動指向南北的磁鐵。',
    '能舉出生活中應用磁鐵的例子，並分辨哪些用品、玩具含有磁鐵。'], 1):
    para(doc, f'{i}. {s}', indent=0.5, space_after=2)

h(doc, '五、教材教具')
table(doc, [
    ['類別', '內容'],
    ['教具', '長條形磁鐵、馬蹄形磁鐵、圓形磁鐵、環形磁鐵（每組各 1–2 個）'],
    ['實驗材料', '迴紋針、釘書針、鐵釘、鐵尺、鐵湯匙、長尾夾、鋁罐、鐵罐、硬幣、橡皮擦、塑膠尺、'
                '紙張、木塊、玻璃杯、寶特瓶、牙籤、燒杯或培養皿'],
    ['其他', '鐵粉或鐵釘屑（裝在夾鏈袋）、棉線、水盆、學習單、平板或電腦（教學網頁）'],
    ['數位', '教學網頁（互動實驗、簡報、10 題測驗、對戰）、教學簡報 24 頁'],
], widths=[2.6, 14.4])

h(doc, '六、教學流程')

FLOW1 = [
    ['時間', '教學活動', '學習目標', '評量'],
    ['5 分', '【引起動機】\n播放教學簡報 1–4 頁：門擋、黑板、冰箱磁條、磁鐵跳棋。\n'
             '提問：「這些東西裡面藏了什麼？如果拿掉它，會怎樣？」', '1', '口頭發表'],
    ['8 分', '【預測】\n發下學習單①。每人先猜 12 樣物品哪些會被吸住，寫下理由。\n'
             '提醒：先猜再驗證，猜錯沒關係。', '1', '學習單①預測欄'],
    ['15 分', '【實測】\n分組操作：用同一個磁鐵依序靠近 12 樣物品，每樣測三次。\n'
              '教師巡視，提醒「輕輕靠近就好、不要壓」。', '1、2', '實作觀察'],
    ['8 分', '【歸納】\n各組把物品分成兩堆貼在海報上，說出共同點。\n'
             '教師追問：「硬幣摸起來像鐵，為什麼沒被吸住？」', '1', '分組報告'],
    ['4 分', '【統整】\n歸納：磁鐵吸引鐵製品，不是所有金屬都會被吸。\n'
             '預告下節：磁鐵的哪裡吸力最強？', '1', '口頭問答'],
]

FLOW2 = [
    ['時間', '教學活動', '學習目標', '評量'],
    ['5 分', '【複習】\n快問快答：鐵釘、硬幣、鋁罐、鐵湯匙，哪些會被吸住？', '1', '口頭問答'],
    ['10 分', '【磁力在哪裡】\n簡報 11–12 頁。用迴紋針掛在長條磁鐵各位置，找出兩端最強。\n'
              '將鐵粉撒在磁鐵上，觀察鐵粉集中在兩端。', '2', '實驗記錄'],
    ['8 分', '【距離與隔物】\n把迴紋針排成一排，量出磁鐵離多遠還吸得到；\n'
             '再用紙、課本隔在中間，試試看還能不能吸住。', '2', '學習單／口頭'],
    ['12 分', '【找出磁極】\n觀察長條、馬蹄、圓形、環形磁鐵上的標示。\n'
              '在學習單②把 N 極塗紅、S 極塗藍，並說出兩極一定同時存在。', '2', '學習單②'],
    ['5 分', '【統整】\n磁力集中在兩端（磁極），離越遠越弱，隔著物品也能吸引。', '2', '口頭問答'],
]

FLOW3 = [
    ['時間', '教學活動', '學習目標', '評量'],
    ['5 分', '【複習】\n指著磁鐵問：這裡是什麼極？磁力哪裡最強？', '2', '口頭問答'],
    ['12 分', '【相斥相吸】\n兩人一組，用兩塊磁鐵依序測 N–N、N–S、S–S，記錄結果並畫下來。\n'
              '（提醒：只用磁鐵相靠，不要用力硬推）', '3', '學習單②記錄表'],
    ['10 分', '【相斥小車】\n把兩台小車的磁鐵同極相對，觀察車子被推開；\n'
              '再換成異極相對，比較差別。', '3', '實作觀察'],
    ['9 分', '【指南北】\n用棉線把磁鐵吊起來，或用保麗龍片讓磁鐵浮在水面，觀察停止時的方向。\n'
             '對照指南針，說明 N 極指北。', '3', '實作記錄'],
    ['4 分', '【統整】\n同極相斥、異極相吸；磁鐵會指南北。', '3', '口頭問答'],
]

FLOW4 = [
    ['時間', '教學活動', '學習目標', '評量'],
    ['10 分', '【尋寶】\n教室尋寶：找出 5 個「有磁鐵」的地方，記下它解決了什麼問題。', '4', '尋寶紀錄'],
    ['10 分', '【玩具分類】\n簡報 22–23 頁：磁鐵釣魚、磁力火車、磁鐵飛鏢、磁鐵跳棋 vs 風車、風箏、陀螺、飛盤。\n'
              '討論：膠帶黏頭髮是磁力嗎？（靜電）', '4', '分類卡'],
    ['12 分', '【線上測驗】\n用平板或電腦開教學網頁，完成 10 題測驗，記錄分數並檢討錯題。', '1–4', '測驗成績'],
    ['8 分', '【兩人對戰】\n兩人一組進入「磁鐵知識王」，10 題搶答。\n'
             '（若無網路，改用實體舉牌搶答）', '1–4', '對戰紀錄'],
]

for n, (title, flow, goal) in enumerate([
        ('第 1 節｜磁鐵能吸引什麼？', FLOW1, '2-1'),
        ('第 2 節｜磁力有多大？磁極在哪裡？', FLOW2, '2-1・2-2'),
        ('第 3 節｜同極相斥、異極相吸', FLOW3, '2-2'),
        ('第 4 節｜生活中的磁鐵＆知識王', FLOW4, '2-3＋統整')], 1):
    para(doc, f'（{["一","二","三","四"][n-1]}）{title}　{goal}', 12, bold=True, color=ORANGE, space_before=10, space_after=4)
    t = table(doc, flow, widths=[1.9, 10.6, 1.6, 2.9], size=10)
    for j, w in enumerate([1.9, 10.6, 1.6, 2.9]):
        for row in t.rows:
            row.cells[j].width = Cm(w)

h(doc, '七、隨堂測驗（10 題，與教學網頁相同）')
QUIZ = [
    ('用磁鐵靠近下列哪一種物品，會被吸住？', ['橡皮擦', '鐵釘', '塑膠尺', '紙張'], 1),
    ('下列哪一種物品「不會」被磁鐵吸引？', ['迴紋針', '鐵湯匙', '硬幣', '鐵尺'], 2),
    ('磁鐵的哪一個部位磁力最強？', ['正中央', '兩端', '整塊都一樣強', '背面'], 1),
    ('把磁鐵慢慢遠離迴紋針，會發生什麼事？', ['吸力變強', '吸力不變', '吸力變弱，最後吸不到', '迴紋針被推開'], 2),
    ('長條形磁鐵的 N 極和 S 極分別在哪裡？', ['都在中間', '都在同一端', '分別在兩端', '只有一端有磁極'], 2),
    ('把兩個磁鐵的 N 極靠近 N 極，會怎樣？', ['互相吸引', '互相排斥', '沒有任何反應', '合成一個磁鐵'], 1),
    ('把磁鐵吊起來自由轉動，停止後 N 極會指向哪裡？', ['東方', '南方', '北方', '正上方'], 2),
    ('把一塊長條磁鐵折成兩段，每一段會怎樣？', ['只有一端有磁極', '變成沒有磁性', '兩段各有 N 極和 S 極', '會互相排斥到很遠'], 2),
    ('下列哪一個生活用品「沒有」用到磁鐵？', ['冰箱門的磁條', '鉛筆盒的蓋子', '風車', '黑板上的磁鐵'], 2),
    ('用膠帶黏起頭髮，和磁鐵吸引鐵釘，原理一樣嗎？', ['一樣，都是磁力', '不一樣，膠帶靠靜電', '一樣，都是靜電', '不一樣，磁鐵是靜電'], 1),
]
for i, (q, o, a) in enumerate(QUIZ, 1):
    para(doc, f'{i}. {q}', bold=True, space_before=2, space_after=0)
    para(doc, '　　' + '　'.join(f'({"ABCD"[k]}) {v}' for k, v in enumerate(o)), size=10.5, space_after=2)
para(doc, '答案：' + '　'.join(f'{i}-{"ABCD"[a]}' for i, (_, _, a) in enumerate(QUIZ, 1)), 10.5, color=ORANGE, bold=True, space_before=6)

h(doc, '八、常見迷思與澄清')
table(doc, [
    ['迷思', '澄清方式'],
    ['「金屬做的東西都會被磁鐵吸住」', '實測鋁罐、硬幣、銅幣，讓學生發現只有鐵（含鐵的鋼）會被吸引。'],
    ['「硬幣是鐵做的」', '查硬幣材質（銅、鎳、鋅），再用磁鐵實際測一次。'],
    ['「磁鐵每一處磁力都一樣」', '用迴紋針掛在磁鐵各位置，或撒鐵粉觀察集中處。'],
    ['「磁鐵斷了就沒有磁性」', '折斷磁鐵後再測兩段，發現各自都有 N、S 兩極。'],
    ['「膠帶黏頭髮是磁力」', '比較膠帶與磁鐵：膠帶黏紙屑、頭髮，磁鐵只吸引鐵製品。'],
    ['「磁鐵越大磁力一定越強」', '用同樣方法吸迴紋針比較數量，讓數據說話。'],
], widths=[6.0, 11.0])

h(doc, '九、評量設計')
table(doc, [
    ['評量方式', '比重', '內容'],
    ['形成性評量', '60%', '學習單①（預測與分類記錄）、學習單②（磁極與相斥相吸記錄）、實驗操作與口頭發表'],
    ['總結性評量', '30%', '教學網頁 10 題線上測驗（或紙筆測驗卷）'],
    ['態度與參與', '10%', '分組合作、器材歸位、對同學報告的回饋'],
], widths=[3.4, 2.0, 11.6])

h(doc, '十、差異化教學')
for s in ['操作較慢的學生：學習單提供「勾選式」表格，只填 ✓／✗，不需書寫長句。',
          '語文表達較弱的學生：允許用畫圖或口頭報告取代文字說明。',
          '學習較快的學生：挑戰「設計一個用磁鐵的小工具」或「比較 4 種磁鐵的磁力排名」。',
          '需要支持的學生：同一組安排一位小老師協助記錄，教師優先巡視該組。']:
    para(doc, '・' + s, indent=0.4, space_after=2)

h(doc, '十一、課後延伸')
for s in ['回家找 3 個家裡用到磁鐵的東西，拍照或畫下來，下次上課分享。',
          '用教學網頁的「吸引大挑戰」再玩一次，挑戰 20 題全對。',
          '和同學約一場「磁鐵知識王」線上對戰，把結果記在聯絡簿。',
          '閱讀：《磁鐵的祕密》等科普書，找出一個課本沒提到的磁鐵應用。']:
    para(doc, '・' + s, indent=0.4, space_after=2)

h(doc, '十二、板書規劃')
table(doc, [
    ['第 1 節', '磁鐵吸引什麼？　→　會被吸：鐵釘、迴紋針、鐵尺…　→　不會被吸：橡皮擦、硬幣、紙…　→　結論：鐵製品'],
    ['第 2 節', '磁力哪裡最強？　→　兩端（磁極）最強　→　距離越遠越弱　→　隔著物品也能吸　→　N 極／S 極'],
    ['第 3 節', '同極相斥（N–N、S–S）　｜　異極相吸（N–S）　→　自由轉動指南北'],
    ['第 4 節', '生活應用：門擋、冰箱磁條、鉛筆盒、磁性白板、磁鐵玩具　｜　靜電 ≠ 磁力'],
], widths=[2.4, 14.6], size=10.5)

h(doc, '十三、教學省思（課後填寫）')
for s in ['學生最感興趣的活動是：', '最常出現的迷思是：', '時間分配需要調整的地方：', '下次要改進的地方：']:
    para(doc, '・' + s, indent=0.4, space_before=6, space_after=2)
    para(doc, '　' + '＿' * 34, color=GREY, indent=0.4, space_after=6)

out_plan = os.path.join(OUT, '教案_磁鐵好好玩.docx')
doc.save(out_plan)
print('教案：', out_plan)
