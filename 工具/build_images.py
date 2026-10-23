import os, json, glob, re
from PIL import Image

SRC = r'D:\##115學年度桃源訓導\新增資料夾\02磁鐵好好玩'
ROOT = r'C:\Users\rayi\Documents\自然科學教學專欄\磁鐵好好玩'
SITE = os.path.join(ROOT, '教學網頁')
IMG = os.path.join(SITE, 'media', 'img')
THUMB = os.path.join(SITE, 'media', 'thumb')
os.makedirs(IMG, exist_ok=True)
os.makedirs(THUMB, exist_ok=True)

files = sorted(glob.glob(os.path.join(SRC, '02-*.jpg')))
rows = []
for i, f in enumerate(files, 1):
    base = os.path.splitext(os.path.basename(f))[0]          # 02-01-001-馬蹄形磁鐵
    m = re.match(r'^(02-\d\d)-(\d{3})-(.+)$', base)
    if not m:
        print('略過', base); continue
    group, num, name = m.group(1), m.group(2), m.group(3)
    code = f'{group}-{num}'
    fn = f'm{i:03d}.jpg'
    im = Image.open(f).convert('RGB')
    im.save(os.path.join(IMG, fn), 'JPEG', quality=88, optimize=True)
    t = im.copy(); t.thumbnail((360, 360))
    t.save(os.path.join(THUMB, fn), 'JPEG', quality=80, optimize=True)
    rows.append({'code': code, 'name': name, 'g': group[-2:], 'f': 'media/img/' + fn, 'h': 'media/thumb/' + fn})

# 分組統計
from collections import Counter
print('總數', len(rows), Counter(r['g'] for r in rows))

js = '/* 素材圖庫清單（自動產生，勿手改） */\nconst IMAGES = ' + json.dumps(rows, ensure_ascii=False, indent=1) + ';\n'
with open(os.path.join(SITE, 'media', 'images.js'), 'w', encoding='utf-8') as fp:
    fp.write(js)
print('images.js 已寫入', len(js), 'bytes')
