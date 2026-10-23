"""Derive dark painterly background plates from the reference art (style-consistent)."""
import os
import numpy as np
from PIL import Image, ImageFilter, ImageEnhance

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.abspath(os.path.join(HERE, "..", "素材", "風格參考圖.png"))
OUT = os.path.join(HERE, "bg")
os.makedirs(OUT, exist_ok=True)
W, H = 1920, 1080

def vignette(img, k=0.55):
    a = np.asarray(img.convert("RGB"), dtype=np.float32) / 255
    yy, xx = np.mgrid[0:H, 0:W].astype(np.float32)
    nx, ny = (xx - W / 2) / (W / 2), (yy - H / 2) / (H / 2)
    d = np.clip(np.sqrt(nx ** 2 + ny ** 2) / 1.41, 0, 1) ** 1.5
    return Image.fromarray(((a * (1 - k * d)[..., None]) * 255).clip(0, 255).astype(np.uint8))

def plate(name, box, dark=0.30, blur=12, bright=0.92, sat=0.85, vig=0.55):
    im = Image.open(SRC).convert("RGB").crop(box).resize((W, H), Image.LANCZOS)
    im = im.filter(ImageFilter.GaussianBlur(blur))
    im = ImageEnhance.Brightness(im).enhance(bright)
    im = ImageEnhance.Color(im).enhance(sat)
    a = np.asarray(im, dtype=np.float32) / 255
    a = a * (1 - dark)          # darken for text contrast
    im = Image.fromarray((a * 255).clip(0, 255).astype(np.uint8))
    im = vignette(im, vig)
    im.save(f"{OUT}\\{name}.png")
    print("wrote", name)

# 1376 x 768 source
plate("bg_cover",  (330, 0, 1050, 430), dark=0.30, bright=0.90, blur=10, vig=0.50)   # 天光 + 藤蔓
plate("bg_a",      (0, 150, 700, 560), dark=0.42, bright=0.85, blur=16, vig=0.58)    # 番茄暖區
plate("bg_b",      (700, 80, 1376, 470), dark=0.44, bright=0.85, blur=16, vig=0.58)  # 花椰冷區
plate("bg_c",      (380, 300, 1080, 720), dark=0.46, bright=0.82, blur=16, vig=0.60) # 中央草地
plate("bg_dark",   (560, 150, 1130, 470), dark=0.62, bright=0.75, blur=22, vig=0.66) # 深處
plate("bg_end",    (60, 480, 760, 768), dark=0.50, bright=0.85, blur=14, vig=0.58)   # 蘑菇角落
