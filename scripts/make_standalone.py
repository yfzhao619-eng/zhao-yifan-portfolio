"""生成单文件版作品集(图片内嵌 base64,可离线/微信发送打开)"""
import base64
import os

BASE = "/Users/fan.plain/Documents/工作/应聘简历/作品集网站"
SRC_HTML = os.path.join(BASE, "index.html")
OUT = os.path.join(BASE, "赵奕帆作品集-单文件版.html")

with open(SRC_HTML, "r", encoding="utf-8") as f:
    html = f.read()

for name, mime in [
    ("assets/avatar.jpg", "image/jpeg"),
    ("assets/life.jpg", "image/jpeg"),
    ("assets/qrcode.png", "image/png"),
]:
    path = os.path.join(BASE, name)
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    data_uri = f"data:{mime};base64,{b64}"
    html = html.replace(name, data_uri)
    html = html.replace('src="' + name + '"', 'src="' + data_uri + '"')

# 标题加上单文件标识
html = html.replace(
    "<title>赵奕帆 · 个人作品集 | 把复杂的事,做成确定的事</title>",
    "<title>赵奕帆 · 个人作品集 | 把复杂的事,做成确定的事</title>",
)

with open(OUT, "w", encoding="utf-8") as f:
    f.write(html)
print("saved:", OUT, round(os.path.getsize(OUT) / 1024 / 1024, 2), "MB")
