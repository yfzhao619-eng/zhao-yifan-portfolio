"""把生活照顺时针旋转 90 度并重新压缩"""
from PIL import Image
import os

BASE = "/Users/fan.plain/Documents/工作/应聘简历"
SRC = os.path.join(BASE, "京东/TET/TET面试材料汇总/TET21+综合方向+赵奕帆+生活照.jpg")
DST = os.path.join(BASE, "作品集网站/assets/life.jpg")

img = Image.open(SRC)
print("original size:", img.size, img.mode)

# 顺时针旋转 90 度 = 逆时针 -90 度,Pillow rotate 角度为逆时针方向
img = img.rotate(-90, expand=True)
print("rotated size:", img.size)

if img.mode != "RGB":
    img = img.convert("RGB")

# 压缩:最大宽度 1400(旋转后宽高对调,按最长边控制)
w, h = img.size
max_side = 1400
ratio = min(1.0, max_side / max(w, h))
img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)
img.save(DST, quality=82, optimize=True)
print("saved:", DST, os.path.getsize(DST) // 1024, "KB")
