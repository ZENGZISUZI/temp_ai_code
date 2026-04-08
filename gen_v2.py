# -*- coding: utf-8 -*-
"""
生成更真实的手写风格感谢信
"""
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import os
import random
import math

def create_handwriting():
    width, height = 1000, 1800
    bg_color = (255, 252, 245)
    text_color = (20, 30, 50)  # 深蓝黑，像钢笔

    img = Image.new('RGB', (width, height), bg_color)
    draw = ImageDraw.Draw(img)

    # 加载字体 - 优先手写风格
    font_paths = [
        r"C:\Windows\Fonts\STXINGKA.TTF",  # 华文行楷
        r"C:\Windows\Fonts\HWFS.TTF",      # 华文仿宋
        r"C:\Windows\Fonts\STKAITI.TTF",   # 楷体
        r"C:\Windows\Fonts\simkai.ttf",
        r"C:\Windows\Fonts\msyh.ttc",
    ]

    font_large = None
    font_normal = None

    for font_path in font_paths:
        if os.path.exists(font_path):
            try:
                font_large = ImageFont.truetype(font_path, 44)
                font_normal = ImageFont.truetype(font_path, 28)
                break
            except:
                continue

    if font_large is None:
        font_large = ImageFont.load_default()
        font_normal = ImageFont.load_default()

    # 纸张纹理
    for _ in range(3000):
        x = random.randint(0, width)
        y = random.randint(0, height)
        gray = random.randint(245, 255)
        draw.point((x, y), fill=(gray, gray, gray))

    # 淡淡的横线
    line_color = (230, 225, 215)
    for y in range(110, height - 50, 52):
        points = []
        for x in range(90, width - 70, 3):
            offset = int(math.sin(x * 0.008 + y * 0.001) * 2)
            points.append((x, y + offset))
        for i in range(len(points) - 1):
            draw.line([points[i], points[i+1]], fill=line_color, width=1)

    # 左边红线
    for y in range(60, height - 40):
        offset = int(math.sin(y * 0.01) * 1.5)
        draw.point((85 + offset, y), fill=(200, 160, 160))

    def draw_char(x, y, char, font, color):
        """绘制单个字符，带随机变换"""
        # 随机旋转
        angle = random.uniform(-12, 12)
        # 随机缩放
        scale = random.uniform(0.85, 1.15)
        # 随机位置偏移
        ox = random.uniform(-3, 3)
        oy = random.uniform(-3, 3)

        # 创建字符图像
        temp = Image.new('RGBA', (80, 80), (0, 0, 0, 0))
        temp_draw = ImageDraw.Draw(temp)
        temp_draw.text((15, 15), char, fill=color, font=font)

        # 旋转
        temp = temp.rotate(angle, expand=True, fillcolor=(0, 0, 0, 0))
        
        # 缩放
        new_w = int(temp.width * scale)
        new_h = int(temp.height * scale)
        temp = temp.resize((new_w, new_h), Image.Resampling.LANCZOS)

        # 粘贴
        px = int(x + ox - temp.width // 2 + 20)
        py = int(y + oy - temp.height // 2 + 20)
        img.paste(temp, (px, py), temp)

    def draw_line(x, y, text, font, color, spacing=30):
        """绘制一行文字"""
        cx = x
        for char in text:
            # 字间距随机变化
            gap = spacing + random.uniform(-4, 4)
            draw_char(cx, y, char, font, color)
            cx += gap

    y = 130

    # 标题
    title = "\u611f\u8c22\u4fe1"
    tw = len(title) * 50
    draw_line((width - tw) // 2, y, title, font_large, text_color, spacing=52)
    y += 85

    # 称呼
    draw_line(130, y, "\u4eb2\u7231\u7684\u90ed\u5955\u9633\uff1a", font_normal, text_color)
    y += 58
    draw_line(130, y, "\u4f60\u597d\uff01", font_normal, text_color)
    y += 65

    # 正文
    paragraphs = [
        "\u63d0\u7b14\u5199\u8fd9\u5c01\u4fe1\uff0c\u662f\u60f3\u90d1\u91cd\u5730\u611f\u8c22\u4f60\u524d\u4e9b\u65e5\u5b50\u7ed9\u6211\u5145\u503cB\u7ad9\u5927\u4f1a\u5458\u7684\u5fc3\u610f\u3002\u867d\u7136\u5f53\u6211\u6536\u5230\u8fd9\u4efd\u793c\u7269\u65f6\uff0c\u53d1\u73b0\u81ea\u5df1\u5df2\u7ecf\u6709\u4f1a\u5458\u4e86\uff0c\u4f46\u8fd9\u4e1d\u6beb\u4e0d\u5f71\u54cd\u6211\u5185\u5fc3\u7684\u611f\u52a8\u4e0e\u611f\u6fc0\u3002",
        "\u4e09\u5341\u5757\u94b1\uff0c\u8bf4\u591a\u4e0d\u591a\uff0c\u8bf4\u5c11\u4e0d\u5c11\u3002\u4f46\u5728\u6211\u770b\u6765\uff0c\u8fd9\u4efd\u5fc3\u610f\u8fdc\u6bd4\u91d1\u94b1\u672c\u8eab\u73cd\u8d35\u5f97\u591a\u3002\u5728\u8fd9\u4e2a\u5feb\u8282\u594f\u7684\u65f6\u4ee3\uff0c\u80fd\u6709\u4eba\u60f3\u8bb0\u7740\u6211\u3001\u613f\u610f\u4e3a\u6211\u82b1\u5fc3\u601d\uff0c\u5df2\u7ecf\u662f\u4e00\u4ef6\u975e\u5e38\u6e29\u6696\u7684\u4e8b\u60c5\u4e86\u3002\u4f60\u6ca1\u6709\u5fd8\u8bb0\u6211\uff0c\u5728\u67d0\u4e2a\u65f6\u523b\u60f3\u8d77\u4e86\u6211\uff0c\u5e76\u4e14\u4ed8\u8bf8\u884c\u52a8\u2014\u2014\u8fd9\u4efd\u60c5\u8c0a\uff0c\u8ba9\u6211\u500d\u611f\u73cd\u60dc\u3002",
        "\u5176\u5b9e\uff0c\u793c\u7269\u7684\u4ef7\u503c\u4ece\u6765\u4e0d\u5728\u4e8e\u5b83\u662f\u5426\u201c\u6709\u7528\u201d\uff0c\u800c\u5728\u4e8e\u5b83\u80cc\u540e\u627f\u8f7d\u7684\u90a3\u4efd\u60c5\u8c0a\u3002\u4f60\u7684\u8fd9\u4efd\u793c\u7269\uff0c\u8ba9\u6211\u611f\u53d7\u5230\u7684\u662f\u88ab\u91cd\u89c6\u3001\u88ab\u5173\u5fc3\u7684\u6e29\u6696\u3002\u5c31\u50cf\u51ac\u65e5\u91cc\u7684\u4e00\u676f\u70ed\u8336\uff0c\u5373\u4f7f\u6211\u4e0d\u6e34\uff0c\u4f46\u90a3\u4efd\u6e29\u5ea6\u5df2\u7ecf\u6696\u5230\u4e86\u5fc3\u91cc\u3002",
        "\u6211\u60f3\u8bf4\u7684\u662f\uff0c\u8c22\u8c22\u4f60\u613f\u610f\u5bf9\u6211\u597d\u3002\u5728\u8fd9\u4e2a\u4e16\u754c\u4e0a\uff0c\u613f\u610f\u771f\u5fc3\u5bf9\u5f85\u670b\u53cb\u7684\u4eba\u5e76\u4e0d\u591a\uff0c\u800c\u4f60\u65e0\u7591\u662f\u5176\u4e2d\u4e4b\u4e00\u3002\u8fd9\u4efd\u4e09\u5341\u5757\u94b1\u7684\u5fc3\u610f\uff0c\u6211\u4f1a\u4e00\u76f4\u8bb0\u5728\u5fc3\u91cc\u3002\u5b83\u63d0\u9192\u6211\uff0c\u53cb\u8c0a\u662f\u9700\u8981\u7528\u5fc3\u7ecf\u8425\u7684\uff0c\u4e5f\u662f\u9700\u8981\u5f7c\u6b64\u73cd\u60dc\u7684\u3002",
        "\u4ee5\u540e\u6709\u673a\u4f1a\uff0c\u6211\u4e00\u5b9a\u4f1a\u56de\u8bf7\u4f60\u4e00\u987f\u597d\u5403\u7684\uff0c\u6216\u8005\u9001\u4f60\u4e00\u4efd\u540c\u6837\u7528\u5fc3\u7684\u793c\u7269\u3002\u4e0d\u662f\u56e0\u4e3a\u8981\u201c\u8fd8\u793c\u201d\uff0c\u800c\u662f\u56e0\u4e3a\u6211\u4e5f\u60f3\u8ba9\u4f60\u611f\u53d7\u5230\u88ab\u91cd\u89c6\u7684\u6e29\u6696\u2014\u2014\u5c31\u50cf\u4f60\u8ba9\u6211\u611f\u53d7\u5230\u7684\u90a3\u6837\u3002",
        "\u6700\u540e\uff0c\u518d\u6b21\u611f\u8c22\u4f60\u7684\u8fd9\u4efd\u5fc3\u610f\u3002\u5e0c\u671b\u6211\u4eec\u7684\u53cb\u8c0a\u80fd\u591f\u957f\u957f\u4e45\u4e45\uff0c\u5e0c\u671b\u672a\u6765\u7684\u65e5\u5b50\u91cc\uff0c\u6211\u4eec\u8fd8\u80fd\u4e92\u76f8\u60f3\u8bb0\u3001\u4e92\u76f8\u6276\u6301\u3002",
        "\u795d\u4f60\u4e00\u5207\u987a\u5229\uff0c\u5929\u5929\u5f00\u5fc3\uff01",
    ]

    for para in paragraphs:
        chars_per_line = 22
        lines = []
        for i in range(0, len(para), chars_per_line):
            lines.append(para[i:i+chars_per_line])

        for line in lines:
            # 行倾斜
            line_off = random.uniform(-6, 6)
            draw_line(150 + line_off, y, "\u3000\u3000" + line, font_normal, text_color)
            y += 52
        y += 18

    # 署名
    y += 55
    sigs = [
        "\u6b64\u81f4",
        "\u656c\u793c\uff01",
        "\u4f60\u7684\u670b\u53cb",
        "\u82cf",
        "2026.4.4"
    ]

    for s in sigs:
        lw = len(s) * 30
        off = random.uniform(-10, 10)
        draw_line(width - lw - 110 + off, y, s, font_normal, text_color)
        y += 52

    # 轻微模糊
    img = img.filter(ImageFilter.SMOOTH_MORE)

    output = r"D:\AI\thank_you_v2.png"
    img.save(output, 'PNG')
    print(f"Saved: {output}")

if __name__ == '__main__':
    random.seed()
    create_handwriting()
