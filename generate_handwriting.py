# -*- coding: utf-8 -*-
"""
生成手写风格感谢信图片
"""
from PIL import Image, ImageDraw, ImageFont
import os

def create_handwritten_letter():
    # 创建画布 - 信纸大小
    width, height = 800, 1400
    # 信纸背景色（米黄色）
    bg_color = (250, 248, 243)
    # 文字颜色（深灰）
    text_color = (44, 44, 44)

    # 创建图片
    img = Image.new('RGB', (width, height), bg_color)
    draw = ImageDraw.Draw(img)

    # 尝试加载楷体字体
    font_paths = [
        r"C:\Windows\Fonts\simkai.ttf",  # 楷体
        r"C:\Windows\Fonts\STKAITI.TTF",  # 华文楷体
        r"C:\Windows\Fonts\simhei.ttf",   # 黑体
        r"C:\Windows\Fonts\msyh.ttc",     # 微软雅黑
    ]

    font_large = None
    font_normal = None

    for font_path in font_paths:
        if os.path.exists(font_path):
            try:
                font_large = ImageFont.truetype(font_path, 36)
                font_normal = ImageFont.truetype(font_path, 22)
                break
            except:
                continue

    if font_large is None:
        # 使用默认字体
        font_large = ImageFont.load_default()
        font_normal = ImageFont.load_default()

    # 绘制信纸左边线
    line_color = (224, 213, 197)
    draw.line([(60, 40), (60, height - 40)], fill=line_color, width=2)

    # 内容
    y_position = 60

    # 标题
    title = "感谢信"
    title_bbox = draw.textbbox((0, 0), title, font=font_large)
    title_width = title_bbox[2] - title_bbox[0]
    draw.text(((width - title_width) // 2, y_position), title, fill=text_color, font=font_large)
    y_position += 60

    # 称呼
    draw.text((80, y_position), "亲爱的郭奕阳：", fill=text_color, font=font_normal)
    y_position += 45

    draw.text((80, y_position), "你好！", fill=text_color, font=font_normal)
    y_position += 50

    # 正文段落
    paragraphs = [
        "提笔写这封信，是想郑重地感谢你前些日子给我充值B站大会员的心意。虽然当我收到这份礼物时，发现自己已经有会员了，但这丝毫不影响我内心的感动与感激。",
        "三十块钱，说多不多，说少不少。但在我看来，这份心意远比金钱本身珍贵得多。在这个快节奏的时代，能有人惦记着我、愿意为我花心思，已经是一件非常温暖的事情了。你没有忘记我，在某个时刻想起了我，并且付诸行动——这份情谊，让我倍感珍惜。",
        "其实，礼物的价值从来不在于它是否"有用"，而在于它背后承载的那份情谊。你的这份礼物，让我感受到的是被重视、被关心的温暖。就像冬日里的一杯热茶，即使我不渴，但那份温度已经暖到了心里。",
        "我想说的是，谢谢你愿意对我好。在这个世界上，愿意真心对待朋友的人并不多，而你无疑是其中之一。这份三十块钱的心意，我会一直记在心里。它提醒我，友谊是需要用心经营的，也是需要彼此珍惜的。",
        "以后有机会，我一定会回请你一顿好吃的，或者送你一份同样用心的礼物。不是因为要"还礼"，而是因为我也想让你感受到被重视的温暖——就像你让我感受到的那样。",
        "最后，再次感谢你的这份心意。希望我们的友谊能够长长久久，希望未来的日子里，我们还能互相惦记、互相扶持。",
        "祝你一切顺利，天天开心！",
    ]

    for para in paragraphs:
        # 自动换行
        chars_per_line = 28
        lines = []
        for i in range(0, len(para), chars_per_line):
            lines.append(para[i:i+chars_per_line])

        for line in lines:
            draw.text((100, y_position), "　　" + line, fill=text_color, font=font_normal)
            y_position += 40
        y_position += 10

    # 署名
    y_position += 30
    signature_lines = [
        "此致",
        "敬礼！",
        "你忠实的朋友",
        "狗腿子 🐕",
        "2026年4月4日"
    ]

    for line in signature_lines:
        line_bbox = draw.textbbox((0, 0), line, font=font_normal)
        line_width = line_bbox[2] - line_bbox[0]
        draw.text((width - line_width - 80, y_position), line, fill=text_color, font=font_normal)
        y_position += 38

    # 保存图片
    output_path = r"D:\AI\感谢信_手写稿.png"
    img.save(output_path, 'PNG', quality=95)
    print(f"图片已保存到: {output_path}")
    return output_path


if __name__ == '__main__':
    create_handwritten_letter()
