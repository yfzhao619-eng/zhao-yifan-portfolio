"""生成作品集网站二维码
用法: python scripts/make_qr.py <URL>
输出: 作品集网站/assets/qrcode.png
"""
import sys
import os

import qrcode
from qrcode.constants import ERROR_CORRECT_M

BASE = "/Users/fan.plain/Documents/工作/应聘简历/作品集网站"
OUT = os.path.join(BASE, "assets/qrcode.png")

if len(sys.argv) > 1:
    url = sys.argv[1]
else:
    url = "https://zhao-yifan-portfolio.vercel.app"

os.makedirs(os.path.dirname(OUT), exist_ok=True)
qr = qrcode.QRCode(
    version=None,
    error_correction=ERROR_CORRECT_M,
    box_size=12,
    border=2,
)
qr.add_data(url)
qr.make(fit=True)
img = qr.make_image(fill_color="#1F1F1F", back_color="white")
img = img.convert("RGB")
img.save(OUT, quality=95)
print("saved:", OUT, img.size, "| url:", url)
