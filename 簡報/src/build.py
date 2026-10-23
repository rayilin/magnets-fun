# -*- coding: utf-8 -*-
"""Build the grade-3 magnet deck in a fantasy-forest style using officecli batch."""
import json, os, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
BG = os.path.join(HERE, "bg")                      # 背景底圖（由 gen_bg2.py 產生）
OUT = os.path.join(HERE, "ops.json")               # officecli batch 指令檔
DECK = os.path.abspath(os.path.join(HERE, "..", "磁鐵的性質與運用_教學簡報.pptx"))

W, H = 33.867, 19.05
M = 2.7
CW = 28.4

INK, INK2, MUTED = "F2F0EA", "C9BFA8", "8C8271"
GOLD, ORANGE, RED, TEAL, PURPLE, GREEN = "B7A740", "B8602A", "BB331D", "45A69F", "7B4BB7", "449B51"
CARD, CARD2 = "171310", "241B13"
HEAD_FONT, BODY_FONT = "源石黑體月 H", "源石黑體月 M"

ops = []
def op(**kw):
    ops.append(kw)

def slide(bg):
    op(command="add", parent="/", type="slide")
    n = len([o for o in ops if o.get("type") == "slide"])
    op(command="add", parent=f"/slide[{n}]", type="picture",
       props={"src": os.path.join(BG, bg + ".png"), "x": "0cm", "y": "0cm",
              "width": f"{W}cm", "height": f"{H}cm"})
    return n

def panel(idx, x, y, w, h, fill=CARD, opacity=82, line=None, geo="roundRect", lw=0.09, radius=None):
    p = {"x": f"{x}cm", "y": f"{y}cm", "width": f"{w}cm", "height": f"{h}cm",
         "fill": fill, "opacity": opacity, "geometry": geo, "lineWidth": lw}
    if line:
        p["line"] = line
    if radius:
        p["radius"] = f"{radius}cm"
    op(command="add", parent=f"/slide[{idx}]", type="shape", props=p)

def rect(idx, x, y, w, h, fill, opacity=100, geo="rect"):
    op(command="add", parent=f"/slide[{idx}]", type="shape",
       props={"x": f"{x}cm", "y": f"{y}cm", "width": f"{w}cm", "height": f"{h}cm",
              "fill": fill, "opacity": opacity, "geometry": geo, "line": "none", "lineWidth": 0})

def txt(idx, x, y, w, h, s, size=15, color=INK, bold=False, font=BODY_FONT,
        align="left", valign="top", spacing=None, margin=0):
    p = {"text": s, "x": f"{x}cm", "y": f"{y}cm", "width": f"{w}cm", "height": f"{h}cm",
         "size": f"{size}pt", "color": color, "font": font, "align": align,
         "valign": valign, "margin": f"{margin}cm", "wrap": "true"}
    if bold:
        p["bold"] = "true"
    if spacing:
        p["lineSpacing"] = spacing
    op(command="add", parent=f"/slide[{idx}]", type="textbox", props=p)

def header(idx, kicker, title, accent=GOLD, lead=None):
    txt(idx, M, 1.35, CW, 0.9, kicker, 11, accent, True, HEAD_FONT)
    txt(idx, M, 2.15, CW, 2.1, title, 34, INK, True, HEAD_FONT, spacing="1.05x")
    rect(idx, M, 4.55, 2.6, 0.09, accent)
    if lead:
        txt(idx, M, 4.95, CW - 1.0, 1.5, lead, 14.5, INK2, False, BODY_FONT, spacing="1.3x")

def footer(idx, total, label="磁鐵的性質與運用"):
    txt(idx, M, 17.65, 18, 0.7, label, 9.5, MUTED, False, BODY_FONT)
    txt(idx, W - 6.2, 17.65, 3.5, 0.7, f"{idx:02d} / {total:02d}", 9.5, MUTED, False, BODY_FONT, align="right")

def note_strip(idx, s, y=15.55, h=1.35, accent=GOLD):
    panel(idx, M, y, CW, h, accent, 14, accent, lw=0.06, geo="roundRect")
    txt(idx, M + 0.9, y + 0.34, CW - 1.8, h - 0.6, s, 13, "E9DCAF", False, BODY_FONT, valign="middle")

def card3(idx, items, y=6.4, h=6.1, accent=GOLD):
    gap = 0.6
    w = (CW - 2 * gap) / 3
    for i, (t, body) in enumerate(items):
        x = M + i * (w + gap)
        panel(idx, x, y, w, h, CARD, 84, GOLD, lw=0.07)
        rect(idx, x + 0.6, y + 0.55, 1.5, 0.09, accent)
        txt(idx, x + 0.6, y + 1.0, w - 1.2, 1.0, t, 19, INK, True, HEAD_FONT)
        txt(idx, x + 0.6, y + 2.15, w - 1.2, h - 2.6, body, 13.5, INK2, False, BODY_FONT, spacing="1.32x")

def rows(idx, items, y=6.2, h=1.85, gap=0.38, accent=GOLD, num_w=1.45, size=15):
    for i, (t, body) in enumerate(items):
        yy = y + i * (h + gap)
        panel(idx, M, yy, CW, h, CARD, 80, None, lw=0)
        op(command="add", parent=f"/slide[{idx}]", type="shape",
           props={"x": f"{M + 0.55}cm", "y": f"{yy + (h - num_w) / 2}cm",
                  "width": f"{num_w}cm", "height": f"{num_w}cm", "fill": accent,
                  "geometry": "ellipse", "line": "none", "lineWidth": 0})
        txt(idx, M + 0.55, yy + (h - num_w) / 2, num_w, num_w, str(i + 1), 13, "14100C",
            True, HEAD_FONT, align="center", valign="middle")
        txt(idx, M + 2.35, yy + 0.22, 10.6, h - 0.35, t, size, INK, True, HEAD_FONT, valign="middle")
        txt(idx, M + 13.2, yy + 0.22, CW - 13.6, h - 0.35, body, 13.5, INK2, False, BODY_FONT,
            valign="middle", spacing="1.25x")

def magnet(idx, x, y, w=13.0, h=2.3):
    """Simple bar-magnet diagram: N (red) + S (teal)."""
    op(command="add", parent=f"/slide[{idx}]", type="shape",
       props={"x": f"{x}cm", "y": f"{y}cm", "width": f"{w}cm", "height": f"{h}cm",
              "fill": TEAL, "geometry": "roundRect", "line": "none", "lineWidth": 0})
    op(command="add", parent=f"/slide[{idx}]", type="shape",
       props={"x": f"{x}cm", "y": f"{y}cm", "width": f"{w/2}cm", "height": f"{h}cm",
              "fill": RED, "geometry": "roundRect", "line": "none", "lineWidth": 0})
    txt(idx, x, y, w / 2, h, "N", 26, "FFF6EA", True, HEAD_FONT, "center", "middle")
    txt(idx, x + w / 2, y, w / 2, h, "S", 26, "FFF6EA", True, HEAD_FONT, "center", "middle")
    txt(idx, x, y + h + 0.15, w / 2, 0.7, "N 極（指北）", 11, INK2, False, BODY_FONT, "center")
    txt(idx, x + w / 2, y + h + 0.15, w / 2, 0.7, "S 極（指南）", 11, INK2, False, BODY_FONT, "center")
    txt(idx, x, y - 0.75, w / 2, 0.7, "◀ 磁力最強", 10.5, GOLD, True, BODY_FONT, "left")
    txt(idx, x + w / 2 - 0.2, y - 0.75, w / 2, 0.7, "磁力最強 ▶", 10.5, GOLD, True, BODY_FONT, "right")
    txt(idx, x + w / 2 - 2.2, y - 0.78, 4.4, 0.7, "中間最弱", 10.5, MUTED, False, BODY_FONT, "center")

# ---------------------------------------------------------------- 1 封面
i = slide("bg_cover")
panel(i, 0, 0, W, H, "0F0B08", 45, None, lw=0, geo="rect")
rect(i, 0, 0, W, 0.35, GOLD, 90)
txt(i, M, 4.5, 12, 0.9, "國小三年級・自然科學・第二單元", 13, GOLD, True, HEAD_FONT)
txt(i, M, 5.55, CW, 5.6, "磁鐵的\n性質與運用", 68, INK, True, HEAD_FONT, spacing="1.02x")
rect(i, M, 12.0, 4.0, 0.11, GOLD)
txt(i, M, 12.55, 24, 2.2,
    "看得見的吸力，看不見的力場——\n走進森林，把磁力一條一條找出來。", 17, INK2, False, BODY_FONT, spacing="1.35x")
txt(i, M, 16.2, 26, 0.9, "對應課綱：INe-II-7・INa-II-3・INc-II-1　｜　建議 2 節課（80 分鐘）", 11, MUTED, False, BODY_FONT)

# ---------------------------------------------------------------- 2 學習地圖
i = slide("bg_c")
header(i, "單元導覽", "這一單元，我們要解開五道鎖", GOLD, "每一道鎖，都對應一個可以動手做的實驗。先看地圖，再出發。")
rows(i, [
    ("磁鐵能吸什麼？", "鐵、鈷、鎳可以；銅、鋁、金、銀不行"),
    ("磁力哪裡最強？", "兩端最強叫磁極，中間最弱"),
    ("同極相斥、異極相吸", "兩顆磁鐵靠近，會推開或黏住"),
    ("磁鐵會指方向", "自由轉動後，N 指北、S 指南"),
    ("磁鐵怎麼用？", "從冰箱貼到磁浮列車"),
], y=6.55, h=1.55, gap=0.3, accent=GOLD)
note_strip(i, "學習目標：能說出磁鐵吸引的物質、找出磁極、用實驗比較磁力大小，並舉出生活中三種磁鐵應用。", y=16.05)

# ---------------------------------------------------------------- 3 生活中的磁鐵
i = slide("bg_a")
header(i, "任務一・找一找", "磁鐵藏在哪裡？", ORANGE,
       "它常常躲起來：藏在門裡、盒子裡、玩具裡，甚至藏在你耳朵旁邊。")
card3(i, [
    ("冰箱與櫃子", "冰箱門上的磁鐵貼、櫃子門的磁扣，都是靠磁鐵吸住鐵片。"),
    ("文具與白板", "鉛筆盒的磁扣、白板上的磁鐵、磁性收納條。"),
    ("玩具與聲音", "磁力釣魚、磁浮玩具；耳機和喇叭裡面也有磁鐵。"),
], y=6.6, h=5.0, accent=ORANGE)
note_strip(i, "課堂活動：給學生 2 分鐘，在教室裡找出 3 個「有磁鐵」的東西，說說它為什麼需要磁鐵。", y=13.6, accent=ORANGE)

# ---------------------------------------------------------------- 4 能吸 / 不能吸
i = slide("bg_b")
header(i, "任務一・實驗", "磁鐵能吸什麼？", TEAL, "先猜一猜，再一個一個測試——這一頁的答案，很多大人也會答錯。")
gw = (CW - 0.8) / 2
panel(i, M, 6.5, gw, 7.9, CARD, 84, GREEN, lw=0.08)
txt(i, M + 0.8, 6.95, gw - 1.6, 0.9, "✓ 吸得住", 20, GREEN, True, HEAD_FONT)
txt(i, M + 0.8, 8.0, gw - 1.6, 6.7,
    "鐵釘、迴紋針、鐵尺\n鐵製罐頭、鐵湯匙\n鎳、鈷（生活中較少見）\n\n共同點：裡面含有「鐵」", 15, INK2, False, BODY_FONT, spacing="1.35x")
x2 = M + gw + 0.8
panel(i, x2, 6.5, gw, 7.9, CARD, 84, RED, lw=0.08)
txt(i, x2 + 0.8, 6.95, gw - 1.6, 0.9, "✗ 吸不住", 20, RED, True, HEAD_FONT)
txt(i, x2 + 0.8, 8.0, gw - 1.6, 6.7,
    "銅線、鋁罐、金戒指\n銀幣、不鏽鋼部分餐具\n塑膠、木頭、玻璃、紙\n\n迷思：不是「所有金屬」都能吸", 15, INK2, False, BODY_FONT, spacing="1.35x")
note_strip(i, "小提醒：不鏽鋼要看成分，含鐵的會被吸、含鎳鉻的常常不會——所以「不鏽鋼」不是單一答案。", y=15.0, accent=TEAL)

# ---------------------------------------------------------------- 5 磁力大小
i = slide("bg_c")
header(i, "任務二・實驗", "磁力大小，怎麼比才公平？", GOLD, "磁力看不見，但可以用「距離」和「數量」把它量出來。")
card3(i, [
    ("方法一：看距離", "磁鐵慢慢靠近迴紋針，能吸到的距離越遠，代表磁力越大。"),
    ("方法二：看數量", "一次能吸起幾根迴紋針？吸得越多，代表磁力越大。"),
    ("公平測試三件事", "同一種迴紋針、同一張桌面、同一顆磁鐵；每次都從一樣的高度開始。"),
], y=6.6, h=5.4, accent=GOLD)
note_strip(i, "記錄方式：把結果畫成長條圖，直條越高＝磁力越強；再比較條形、圓形、U 形三種磁鐵。", y=13.7)

# ---------------------------------------------------------------- 6 磁極
i = slide("bg_b")
header(i, "任務二・發現", "磁力最強的地方：磁極", TEAL, "同一顆磁鐵，每個位置的磁力並不一樣強。")
card3(i, [
    ("兩端最強", "條形磁鐵的兩個尖端磁力最大，這兩個位置就叫「磁極」。"),
    ("中間最弱", "磁鐵正中間幾乎吸不住迴紋針，磁力最弱。"),
    ("一定有兩個", "不管磁鐵多小、什麼形狀，都一定有 N 極和 S 極兩個磁極。"),
], y=6.6, h=4.5, accent=TEAL)
magnet(i, 10.4, 12.2, w=13.0, h=2.1)
note_strip(i, "形狀實驗：條形、U 形（馬蹄形）、圓形、環形——圓形磁鐵的磁極在「上下兩面」，不在側邊。", y=15.45, accent=TEAL)

# ---------------------------------------------------------------- 7 鐵粉看見磁力
i = slide("bg_dark")
header(i, "任務二・觀察", "用鐵粉，讓磁力現形", PURPLE,
       "磁力看不見，但鐵粉會排隊告訴你它在哪裡。")
panel(i, M, 6.6, 15.6, 8.4, CARD, 86, PURPLE, lw=0.08)
txt(i, M + 0.9, 7.1, 13.8, 1.0, "你會看到", 16, PURPLE, True, HEAD_FONT)
txt(i, M + 0.9, 8.1, 13.8, 6.5,
    "磁極附近：鐵粉又密又尖\n中間：鐵粉稀稀疏疏\n\n鐵粉排出的線，就是「磁力線」\n從 N 極出發，繞一圈回到 S 極。", 14.5, INK2, False, BODY_FONT, spacing="1.4x")
x2 = M + 16.4
panel(i, x2, 6.6, CW - 16.4, 8.4, CARD2, 80, GOLD, lw=0.07)
txt(i, x2 + 0.8, 7.1, CW - 18.0, 1.0, "動手做", 16, GOLD, True, HEAD_FONT)
txt(i, x2 + 0.8, 8.1, CW - 18.0, 6.5,
    "1. 把磁鐵放在桌上\n2. 蓋上一張薄紙\n3. 撒少量鐵粉\n4. 輕敲紙邊，看圖案\n\n（鐵粉用完要回收）", 14.5, INK2, False, BODY_FONT, spacing="1.4x")

# ---------------------------------------------------------------- 8 相斥相吸
i = slide("bg_a")
header(i, "任務三・核心", "同極相斥，異極相吸", RED, "兩顆磁鐵靠近，只有兩種結果：推開，或黏住。")
gw = (CW - 0.8) / 2
panel(i, M, 6.5, gw, 4.6, CARD, 86, RED, lw=0.08)
txt(i, M + 0.8, 6.95, gw - 1.6, 1.0, "相斥（推開）", 20, RED, True, HEAD_FONT)
txt(i, M + 0.8, 7.95, gw - 1.6, 3.2, "N 對 N　→　推開\nS 對 S　→　推開", 16, INK2, False, BODY_FONT, spacing="1.4x")
x2 = M + gw + 0.8
panel(i, x2, 6.5, gw, 4.6, CARD, 86, GREEN, lw=0.08)
txt(i, x2 + 0.8, 6.95, gw - 1.6, 1.0, "相吸（黏住）", 20, GREEN, True, HEAD_FONT)
txt(i, x2 + 0.8, 7.95, gw - 1.6, 3.2, "N 對 S　→　吸住\nS 對 N　→　吸住", 16, INK2, False, BODY_FONT, spacing="1.4x")
panel(i, M, 11.7, CW, 2.0, GOLD, 14, GOLD, lw=0.06)
txt(i, M + 0.9, 11.95, CW - 1.8, 1.5, "口訣：同極相斥，異極相吸。　試試看：把兩顆磁鐵慢慢靠近，手心會感覺到一股推力。",
    15, "E9DCAF", True, BODY_FONT, valign="middle")
note_strip(i, "應用連連看：磁浮玩具、磁浮列車，都是利用「相斥」把車體撐起來，減少摩擦。", y=14.1, accent=RED)

# ---------------------------------------------------------------- 9 指方向
i = slide("bg_c")
header(i, "任務四・發現", "磁鐵會自己找方向", GOLD, "只要讓它自由轉動，它最後一定會停在同一條線上。")
rows(i, [
    ("把磁鐵吊起來", "用細線綁住中間，等它停止轉動。"),
    ("或讓它浮在水面", "把小磁鐵放在保麗龍片上，浮在水盆裡。"),
    ("觀察它停在哪", "N 極指向北方，S 極指向南方——這就是指南針的原理。"),
], y=6.4, h=2.05, gap=0.42, accent=GOLD)
note_strip(i, "注意：旁邊如果有其他磁鐵或一大塊鐵，指針就會被拉偏，不再準確指向南北。", y=14.2)

# ---------------------------------------------------------------- 10 穿透
i = slide("bg_b")
header(i, "任務四・實驗", "磁力能穿過東西嗎？", TEAL, "可以——但要看你用什麼擋在中間。")
card3(i, [
    ("穿得過", "紙、塑膠片、玻璃、水、薄木板——磁力照樣吸得住。"),
    ("穿不過", "太厚的鐵或鋼板會把磁力「接走」，後面的鐵就吸不動了。"),
    ("生活應用", "隔著紙盒玩磁力迷宮、冰箱門內外、磁性門扣。"),
], y=6.5, h=5.0, accent=TEAL)
note_strip(i, "實驗設計：把迴紋針放在杯底，磁鐵在杯外移動——水、玻璃都擋不住磁力。", y=13.6, accent=TEAL)

# ---------------------------------------------------------------- 11 磁化
i = slide("bg_dark")
header(i, "任務五・延伸", "鐵，也能變成磁鐵", PURPLE, "磁鐵不只會吸鐵，還能「教」鐵變成磁鐵。")
rows(i, [
    ("摩擦法", "拿鐵釘在磁鐵上，同一個方向摩擦 20 次。"),
    ("接觸法", "把鐵釘貼在磁鐵上放一會兒，它也會吸住迴紋針。"),
    ("會消失", "拿開磁鐵後，鐵釘的磁性會慢慢消失——這叫「暫時磁化」。"),
], y=6.4, h=2.0, gap=0.4, accent=PURPLE)
note_strip(i, "想一想：為什麼螺絲起子前端常帶磁性？因為可以先把螺絲「吸」在起子上，單手就能鎖。", y=14.0, accent=PURPLE)

# ---------------------------------------------------------------- 12 變弱
i = slide("bg_a")
header(i, "任務五・保養", "磁鐵為什麼會變弱？", ORANGE, "磁力不是永遠不變——有些動作會讓它慢慢失效。")
rows(i, [
    ("加熱", "用火烤或放在很熱的地方，磁力會明顯減弱。"),
    ("撞擊", "重摔、用力敲打，磁鐵內部的排列被打亂。"),
    ("亂放", "長期和其他磁鐵隨便堆疊，也會讓磁性變弱；成對、同極相鄰收納較好。"),
], y=6.4, h=2.0, gap=0.4, accent=ORANGE)
note_strip(i, "保養四招：不摔、不烤、成對收、遠離手錶與卡片（磁條會被消磁）。", y=14.0, accent=ORANGE)

# ---------------------------------------------------------------- 13 生活應用
i = slide("bg_b")
header(i, "任務五・應用（一）", "生活裡的磁鐵小幫手", GREEN, "最好的發明，通常小到你沒發現它是磁鐵。")
card3(i, [
    ("固定東西", "冰箱磁鐵貼、磁性門擋、窗簾磁扣、包包磁扣。"),
    ("收納整理", "鉛筆盒磁扣、磁性收納條、工具掛架、磁性刀架。"),
    ("好玩好學", "磁力釣魚、磁力迷宮、磁浮玩具、磁性七巧板。"),
], y=6.5, h=5.0, accent=GREEN)
note_strip(i, "課堂任務：選一個生活磁鐵用品，畫出它「哪裡是磁鐵、哪裡是鐵片」，並說明為什麼要這樣設計。", y=13.5, accent=GREEN)

# ---------------------------------------------------------------- 14 科技應用
i = slide("bg_c")
header(i, "任務五・應用（二）", "藏在科技裡的磁鐵", PURPLE, "馬達、喇叭、列車、醫院——磁鐵撐起了現代生活。")
card3(i, [
    ("馬達與聲音", "電風扇、玩具車、洗衣機的馬達；喇叭與耳機靠磁鐵推動振膜。"),
    ("磁浮列車", "用同極相斥把車體撐離軌道，幾乎沒有摩擦，速度可以非常快。"),
    ("醫療與能源", "MRI 核磁共振檢查、風力發電機、硬碟讀寫頭。"),
], y=6.5, h=5.0, accent=PURPLE)
note_strip(i, "比較看看：一般磁鐵的磁力是「永久」且固定的；電磁鐵要通電才有磁力，而且可以開關、可以變強。", y=13.5, accent=PURPLE)

# ---------------------------------------------------------------- 15 總結
i = slide("bg_end")
header(i, "單元總結", "三句話，帶走整個單元", GOLD)
card3(i, [
    ("① 只能吸含鐵物", "鐵、鈷、鎳可以；銅、鋁、金、銀、塑膠不行。"),
    ("② 磁極最強", "兩端最強、中間最弱；同極相斥、異極相吸。"),
    ("③ 能指方向、能穿透", "自由轉動指南北；紙、水、玻璃擋不住磁力。"),
], y=5.6, h=4.9, accent=GOLD)
panel(i, M, 11.5, CW, 4.6, CARD2, 84, GOLD, lw=0.08)
txt(i, M + 0.9, 11.85, 12, 0.9, "動手做：自製指南針", 17, GOLD, True, HEAD_FONT)
txt(i, M + 0.9, 12.9, CW - 1.8, 3.0,
    "① 用磁鐵同方向摩擦縫衣針 30 次　② 把針穿過小保麗龍片　③ 輕輕放在水面上　④ 等它停下來，針尖指的就是北方。\n"
    "（提醒：附近不要放其他磁鐵，否則指針會亂轉。）", 13.5, INK2, False, BODY_FONT, spacing="1.35x")

# ---------------------------------------------------------------- footer
for n in range(1, len([o for o in ops if o.get("type") == "slide"]) + 1):
    footer(n, 15)

with open(OUT, "w", encoding="utf-8") as f:
    json.dump(ops, f, ensure_ascii=False)

CLI = r"C:\Users\rayi\AppData\Local\OfficeCli\officecli.exe"

if os.path.exists(DECK):
    subprocess.run([CLI, "close", DECK], capture_output=True, text=True, encoding="utf-8", errors="replace")
    os.remove(DECK)
os.makedirs(os.path.dirname(DECK), exist_ok=True)
subprocess.run([CLI, "create", DECK], check=True, capture_output=True, text=True, encoding="utf-8", errors="replace")
r = subprocess.run([CLI, "batch", DECK, "--input", OUT, "--json", "--best-effort"],
                   capture_output=True, text=True, encoding="utf-8", errors="replace")
print("exit", r.returncode)
print((r.stdout or "")[:3000])
print((r.stderr or "")[:2000])

subprocess.run([CLI, "close", DECK], capture_output=True, text=True, encoding="utf-8", errors="replace")
print("PUBLISHED:", DECK)
