"""处理证件照和生活照,生成网页用压缩图"""
from PIL import Image
import os

BASE = "/Users/fan.plain/Documents/工作/应聘简历"
SRC_ID = os.path.join(BASE, "京东/TET/TET面试材料汇总/TET21+综合方向+赵奕帆+证件照.JPG")
SRC_LIFE = os.path.join(BASE, "京东/TET/TET面试材料汇总/TET21+综合方向+赵奕帆+生活照.jpg")
DST = os.path.join(BASE, "作品集网站/assets")

os.makedirs(DST, exist_ok=True)

# 证件照:裁剪为正方形居中头像
img_id = Image.open(SRC_ID)
w, h = img_id.size
s = min(w, h)
left = (w - s) // 2
top = (h - s) // 2
img_id = img_id.crop((left, top, left + s, top + s))
if img_id.mode != "RGB":
    img_id = img_id.convert("RGB")
img_id = img_id.resize((900, 900), Image.LANCZOS)
img_id.save(os.path.join(DST, "avatar.jpg"), quality=88, optimize=True)

# 生活照:压缩到最大宽度 1400
img_life = Image.open(SRC_LIFE)
if img_life.mode != "RGB":
    img_life = img_life.convert("RGB")
wl, hl = img_life.size
ratio = min(1.0, 1400 / wl)
img_life = img_life.resize((int(wl * ratio), int(hl * ratio)), Image.LANCZOS)
img_life.save(os.path.join(DST, "life.jpg"), quality=82, optimize=True)

for f in sorted(os.listdir(DST)):
    p = os.path.join(DST, f)
    print(f, os.path.getsize(p) // 1024, "KB")
print("done")
