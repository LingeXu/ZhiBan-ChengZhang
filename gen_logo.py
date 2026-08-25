# -*- coding: utf-8 -*-
"""生成「智伴成长」商标/Logo

设计概念：
- 蓝色渐变圆角方形为底（科技感）
- 白色对话气泡 + 蓝色"智"（AI 智能问答）
- 橙色对话气泡 + 白色"伴"（朋辈陪伴的温度感）
- 双气泡交叠构成"智伴"，呼应"智慧与温暖"的品牌主张

输出：
- 智伴成长_logo_icon.png   （1024×1024 方形图标）
- 智伴成长_logo_横版.png    （1600×560 透明背景，白/浅底使用）
- 智伴成长_logo_横版.jpg    （1600×560 白底，可直接插入 Word/PPT）
"""
from PIL import Image, ImageDraw, ImageFont

# ---------------- 品牌色 ----------------
BLUE_D = (31, 78, 121)      # #1F4E79 深蓝
BLUE_M = (46, 117, 182)     # #2E75B6 中蓝
ORANGE = (237, 125, 49)     # #ED7D31 橙
DARK = (70, 70, 70)
WHITE = (255, 255, 255)

FONT_PATHS = [
    ('/System/Library/Fonts/Hiragino Sans GB.ttc', {0: 'W3', 1: 'W6'}),   # 0=常规 1=粗体
    ('/System/Library/Fonts/STHeiti Medium.ttc', {0: 'M'}),
    ('/System/Library/Fonts/PingFang.ttc', {3: 'R', 5: 'SB', 0: 'R'}),
]

def get_font(size, bold=True):
    for path, faces in FONT_PATHS:
        try:
            idx = 1 if bold and 1 in faces else 0
            return ImageFont.truetype(path, size, index=idx)
        except Exception:
            continue
    return ImageFont.load_default()

def v_gradient(w, h, top, bottom):
    """垂直渐变（linear_gradient 自上而下为白→黑，故两色交换传入）"""
    grad = Image.linear_gradient('L').resize((w, h))
    return Image.composite(Image.new('RGB', (w, h), bottom),
                           Image.new('RGB', (w, h), top), grad)

def rounded_mask(size, radius):
    m = Image.new('L', (size, size), 0)
    d = ImageDraw.Draw(m)
    d.rounded_rectangle([0, 0, size - 1, size - 1], radius=radius, fill=255)
    return m

def draw_icon(size):
    """返回带透明圆角的图标（RGBA，size×size）"""
    S = size * 2  # 2× 渲染再缩放，抗锯齿
    img = v_gradient(S, S, BLUE_M, BLUE_D).convert('RGBA')
    d = ImageDraw.Draw(img)
    # 装饰圆环（极淡白）
    for (cx, cy, r, a) in [(int(S * 0.83), int(S * 0.15), int(S * 0.13), 40),
                           (int(S * 0.14), int(S * 0.87), int(S * 0.07), 30)]:
        ov = Image.new('RGBA', (S, S), (0, 0, 0, 0))
        od = ImageDraw.Draw(ov)
        od.ellipse([cx - r, cy - r, cx + r, cy + r], outline=(255, 255, 255, a),
                   width=max(2, S // 200))
        img.alpha_composite(ov)
    # 白色大气泡（智）
    bx1, by1, bx2, by2 = int(S * 0.147), int(S * 0.156), int(S * 0.693), int(S * 0.547)
    d.polygon([(int(S * 0.21), by2 - int(S * 0.012)), (int(S * 0.33), by2 - int(S * 0.012)),
               (int(S * 0.218), int(S * 0.63))], fill=WHITE)          # 尾巴
    d.rounded_rectangle([bx1, by1, bx2, by2], radius=int(S * 0.127), fill=WHITE)
    fz = get_font(int(S * 0.29), bold=True)
    d.text(((bx1 + bx2) / 2, (by1 + by2) / 2), '智', font=fz, fill=BLUE_D, anchor='mm')
    # 橙色小气泡（伴），交叠于右下
    sx1, sy1, sx2, sy2 = int(S * 0.557), int(S * 0.459), int(S * 0.85), int(S * 0.708)
    d.polygon([(int(S * 0.684), sy2 - int(S * 0.012)), (int(S * 0.781), sy2 - int(S * 0.012)),
               (int(S * 0.757), int(S * 0.791))], fill=ORANGE)        # 尾巴
    d.rounded_rectangle([sx1, sy1, sx2, sy2], radius=int(S * 0.093), fill=ORANGE)
    fb = get_font(int(S * 0.16), bold=True)
    d.text(((sx1 + sx2) / 2, (sy1 + sy2) / 2), '伴', font=fb, fill=WHITE, anchor='mm')
    # 圆角裁切
    img.putalpha(rounded_mask(S, int(S * 0.225)))
    return img.resize((size, size), Image.LANCZOS)

def draw_lockup(size, transparent=True):
    """横版组合：图标 + 中英文名 + 标语"""
    W, H = size
    S = W * 2, H * 2  # 2× 渲染
    base = Image.new('RGBA', S, (255, 255, 255, 0 if transparent else 255))
    d = ImageDraw.Draw(base)
    # 左侧图标
    icon = draw_icon(int(S[1] * 0.86))
    ix, iy = int(S[0] * 0.031), int(S[1] * 0.071)
    base.alpha_composite(icon, (ix, iy))
    # 右侧文字
    x = ix + icon.size[0] + int(S[0] * 0.062)
    f_title = get_font(int(S[1] * 0.25), bold=True)
    d.text((x, int(S[1] * 0.095)), '智伴成长', font=f_title, fill=BLUE_D)
    # 橙色装饰线
    y_rule = int(S[1] * 0.455)
    d.line([(x + int(S[0] * 0.004), y_rule), (x + int(S[0] * 0.10), y_rule)],
           fill=ORANGE, width=max(6, S[1] // 56))
    f_sub = get_font(int(S[1] * 0.082), bold=False)
    d.text((x + int(S[0] * 0.005), int(S[1] * 0.53)),
           '大学生全周期智能导学系统 · AI 数字学长', font=f_sub, fill=DARK)
    f_en = get_font(int(S[1] * 0.052), bold=False)
    d.text((x + int(S[0] * 0.005), int(S[1] * 0.69)),
           'Z H I B A N  ·  A I   S E N I O R   C O M P A N I O N', font=f_en, fill=BLUE_M)
    return base.resize(S and (W, H), Image.LANCZOS)

if __name__ == '__main__':
    icon = draw_icon(1024)
    icon.save('智伴成长_logo_icon.png')
    print('已生成: 智伴成长_logo_icon.png (1024×1024)')

    png = draw_lockup((1600, 560), transparent=True)
    png.save('智伴成长_logo_横版.png')
    print('已生成: 智伴成长_logo_横版.png (1600×560 透明底)')

    jpg = draw_lockup((1600, 560), transparent=False).convert('RGB')
    jpg.save('智伴成长_logo_横版.jpg', quality=92)
    print('已生成: 智伴成长_logo_横版.jpg (1600×560 白底)')
