# -*- coding: utf-8 -*-
"""生成「智伴成长」大创答辩 PPT（19 页，16:9）

风格：简洁、视觉化、数据驱动；蓝色系（科技感）+ 橙色点缀（温度感）。
所有数字与《智伴成长_商业计划书.docx》终稿一致。
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.oxml.ns import qn

# ---------------- 配色 ----------------
BLUE = RGBColor(0x1F, 0x4E, 0x79)      # 深蓝（主色）
MIDBLUE = RGBColor(0x2E, 0x75, 0xB6)    # 中蓝
LIGHTBLUE = RGBColor(0xDE, 0xEA, 0xF6)  # 浅蓝
ORANGE = RGBColor(0xED, 0x7D, 0x31)     # 橙色（温度）
LIGHTORANGE = RGBColor(0xFB, 0xE5, 0xD6)
GRAY = RGBColor(0x7F, 0x7F, 0x7F)
MGRAY = RGBColor(0xBF, 0xBF, 0xBF)
DARK = RGBColor(0x26, 0x26, 0x26)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BGRAY = RGBColor(0xF2, 0xF2, 0xF2)

FONT = '微软雅黑'
TOTAL = 19
OUT = '/Users/xulingexu/Desktop/大创/智伴成长_答辩PPT.pptx'

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

# ---------------- 基础工具 ----------------
def set_ea(run, name=FONT):
    rPr = run._r.get_or_add_rPr()
    latin = rPr.find(qn('a:latin'))
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        if latin is not None:
            latin.addnext(ea)
        else:
            rPr.append(ea)
    ea.set('typeface', name)

def run_tf(run, text, size=14, bold=False, color=DARK, font=FONT):
    run.text = text
    f = run.font
    f.name = font
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = color
    set_ea(run, font)

def add_text(slide, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP, wrap=True):
    """paras: 段落列表；每段 = (runs, align, space_after_pt, line_spacing)
    runs = [(text, size, bold, color), ...]"""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, (runs, align, sa, ls) in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if sa:
            p.space_after = Pt(sa)
        if ls:
            p.line_spacing = ls
        for (t, s, b, c) in runs:
            r = p.add_run()
            run_tf(r, t, s, b, c)
    return tb

def add_shape(slide, st, x, y, w, h, fill=None, line=None, line_w=0.75, radius=None):
    sp = slide.shapes.add_shape(st, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w)
    if radius is not None and st == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            sp.adjustments[0] = radius
        except Exception:
            pass
    sp.text_frame.margin_left = sp.text_frame.margin_right = 0
    sp.text_frame.margin_top = sp.text_frame.margin_bottom = 0
    return sp

def card(slide, x, y, w, h, fill=WHITE, line=MGRAY, radius=0.06, line_w=1.0):
    return add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h,
                     fill=fill, line=line, line_w=line_w, radius=radius)

def shape_text(sp, paras, anchor=MSO_ANCHOR.MIDDLE):
    tf = sp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for i, (runs, align, sa, ls) in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if sa:
            p.space_after = Pt(sa)
        if ls:
            p.line_spacing = ls
        for (t, s, b, c) in runs:
            r = p.add_run()
            run_tf(r, t, s, b, c)

def new_slide():
    return prs.slides.add_slide(BLANK)

def header(slide, title, idx, accent=ORANGE):
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0.55, 0.52, 0.13, 0.56, fill=BLUE)
    add_text(slide, 0.85, 0.40, 11.0, 0.8,
             [([(title, 25, True, BLUE)], PP_ALIGN.LEFT, 0, 1.0)])
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0.55, 1.18, 2.4, 0.045, fill=accent)
    add_text(slide, 12.35, 7.02, 0.8, 0.35,
             [([(f'{idx:02d} / {TOTAL}', 10, False, GRAY)], PP_ALIGN.RIGHT, 0, 1.0)])

def chip(slide, x, y, w, h, text, fill, tcolor=WHITE, size=13, bold=True):
    sp = add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h,
                   fill=fill, radius=0.5)
    shape_text(sp, [([(text, size, bold, tcolor)], PP_ALIGN.CENTER, 0, 1.0)])
    return sp

# ================= S1 封面 =================
s = new_slide()
add_shape(s, MSO_SHAPE.RECTANGLE, 0, 0, 13.334, 7.5, fill=BLUE)
add_shape(s, MSO_SHAPE.RECTANGLE, 0, 0, 13.334, 0.12, fill=ORANGE)
add_text(s, 0.9, 0.75, 11.5, 0.5,
         [([('大学生创新创业训练计划项目 · 立项答辩', 15, False, RGBColor(0xBF, 0xD3, 0xE6))], PP_ALIGN.LEFT, 0, 1.0)])
add_text(s, 0.9, 2.15, 11.5, 1.5,
         [([('智伴成长', 54, True, WHITE)], PP_ALIGN.LEFT, 0, 1.0)])
add_text(s, 0.95, 3.45, 11.5, 0.6,
         [([('—— 基于 AI 与朋辈导航的大学生全周期智能导学系统', 20, False, RGBColor(0xD9, 0xE2, 0xF3))], PP_ALIGN.LEFT, 0, 1.0)])
add_shape(s, MSO_SHAPE.RECTANGLE, 0.95, 4.25, 2.6, 0.05, fill=ORANGE)
add_text(s, 0.95, 4.5, 11.5, 0.5,
         [([('让每一段大学时光，都有智慧与温暖相伴', 15, False, RGBColor(0xF4, 0xB1, 0x83))], PP_ALIGN.LEFT, 0, 1.0)])
add_text(s, 0.95, 5.9, 11.5, 1.2,
         [([('答辩人：徐麟阁　｜　团队成员：徐麟阁、董芮铭、章宇洲', 13, False, WHITE)], PP_ALIGN.LEFT, 0, 1.2),
          ([('指导教师：陈圣波', 13, False, WHITE)], PP_ALIGN.LEFT, 0, 1.2),
          ([('2026 年 8 月', 13, False, RGBColor(0xBF, 0xD3, 0xE6))], PP_ALIGN.LEFT, 0, 1.2)])

# ================= S2 目录 =================
s = new_slide()
header(s, '目录', 2)
toc = [
    ('01', '核心定位与市场', '数字学长 · 双向痛点 · 市场规模'),
    ('02', '产品与差异化', '四层架构 · 四大模块 · 量化指标'),
    ('03', '商业模式', 'B2B 订阅 · 三档定价 · 单校经济账'),
    ('04', '竞争策略', '错位竞争 · 三重壁垒 · 竞品对比'),
    ('05', '财务与融资', '三年预测 · 种子轮 50 万'),
    ('06', '团队与规划', '团队分工 · 导师支持 · 三年里程碑'),
    ('07', '风险与展望', 'Top3 风险应对 · 使命结语'),
]
positions = [(0.9, 1.9), (6.8, 1.9), (0.9, 3.3), (6.8, 3.3), (0.9, 4.7), (6.8, 4.7), (0.9, 6.1)]
for (num, t, sub), (x, y) in zip(toc, positions):
    c = card(s, x, y, 5.5, 1.05, fill=WHITE, line=MGRAY)
    ov = add_shape(s, MSO_SHAPE.OVAL, x + 0.25, y + 0.27, 0.5, 0.5, fill=BLUE)
    shape_text(ov, [([(num, 15, True, WHITE)], PP_ALIGN.CENTER, 0, 1.0)])
    add_text(s, x + 0.95, y + 0.16, 4.4, 0.8,
             [([(t, 16, True, DARK)], PP_ALIGN.LEFT, 0, 1.1),
              ([(sub, 11.5, False, GRAY)], PP_ALIGN.LEFT, 0, 1.1)])

# ================= S3 核心定位 =================
s = new_slide()
header(s, '核心定位：什么是“数字学长”？', 3)
ban = add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.9, 1.6, 11.5, 1.15, fill=LIGHTBLUE, radius=0.12)
shape_text(ban, [([('一句话定义：', 17, True, BLUE), ('为每一位大学生配备的 7×24 在线智能导学助手——', 17, True, DARK)], PP_ALIGN.CENTER, 0, 1.15),
                 ([('将 AI 的高效与朋辈导师的亲和力结合，覆盖大学全周期成长', 14, False, GRAY)], PP_ALIGN.CENTER, 0, 1.15)])
roles = [
    ('工具', '高效的知识管理工具', '“查什么、怎么办、去哪办”，三步之内得到答案', MIDBLUE),
    ('导师', '可靠的学业与生活导师', '选课建议 · 挂科预警 · 关键节点主动提醒', BLUE),
    ('陪伴', '有温度的陪伴者', '情绪感知与心理支持 · 危机信号自动转介', ORANGE),
]
for i, (tag, t, d, c) in enumerate(roles):
    x = 0.9 + i * 3.95
    cd = card(s, x, 3.25, 3.6, 2.6, fill=WHITE, line=c, line_w=1.5)
    chip(s, x + 1.28, 3.5, 1.04, 0.44, tag, c, size=13)
    add_text(s, x + 0.3, 4.15, 3.0, 1.6,
             [([(t, 16, True, DARK)], PP_ALIGN.CENTER, 4, 1.15),
              ([(d, 12, False, GRAY)], PP_ALIGN.CENTER, 0, 1.25)])
add_text(s, 0.9, 6.35, 11.5, 0.5,
         [([('三重角色 = 工具 + 导师 + 陪伴，覆盖“适应—成长—关怀”完整用户旅程', 13, True, BLUE)], PP_ALIGN.CENTER, 0, 1.0)])

# ================= S4 市场痛点 =================
s = new_slide()
header(s, '市场痛点：学生端与学校端的双向困境', 4)
add_text(s, 0.9, 1.45, 5.6, 0.5, [([('学生端之痛', 17, True, BLUE)], PP_ALIGN.CENTER, 0, 1.0)])
add_text(s, 6.85, 1.45, 5.6, 0.5, [([('学校端之痛', 17, True, ORANGE)], PP_ALIGN.CENTER, 0, 1.0)])
stu = [('信息获取碎片化', '办事指南散落在官网、部门与聊天记录中'),
       ('选课决策迷茫', '学长经验靠口口相传，缺乏系统化渠道'),
       ('心理疏导滞后', '愿意倾诉、不愿当面求助')]
sch = [('辅导员“一对多”', '师生比不低于 1:200，难以兼顾个体需求'),
       ('经验难以沉淀', '学长经验随毕业流失，无法代际传承'),
       ('重复答疑负担重', '事务性咨询占用大量学工工时')]
c1 = card(s, 0.9, 2.0, 5.6, 3.9, fill=LIGHTBLUE, line=None)
c2 = card(s, 6.85, 2.0, 5.6, 3.9, fill=LIGHTORANGE, line=None)
for (cd, items, c) in [(c1, stu, BLUE), (c2, sch, ORANGE)]:
    paras = []
    for t, d in items:
        paras.append(([( '▍', 14, True, c), (t, 14.5, True, DARK)], PP_ALIGN.LEFT, 2, 1.15))
        paras.append(([(d, 12, False, GRAY)], PP_ALIGN.LEFT, 14, 1.2))
    shape_text(cd, paras, anchor=MSO_ANCHOR.MIDDLE)
    cd.text_frame.margin_left = Inches(0.35)
    cd.text_frame.margin_right = Inches(0.25)
ban = add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.9, 6.3, 11.5, 0.75, fill=WHITE, line=ORANGE, line_w=1.5, radius=0.5)
shape_text(ban, [([('双向痛点 → 一个方案：', 14, True, DARK), ('“智伴成长”数字学长', 14, True, ORANGE)], PP_ALIGN.CENTER, 0, 1.0)])

# ================= S5 市场规模 =================
s = new_slide()
header(s, '市场规模：TAM / SAM / SOM 三层漏斗', 5)
funnel = [
    ('TAM　总体市场', '全国 3,000+ 所高校 × 年均 5 万元', '约 15,000 万元/年', 11.4, MIDBLUE, WHITE),
    ('SAM　可服务市场', '具备采购意愿的本科院校约 800 所（约 27%）', '约 4,000 万元/年', 9.4, BLUE, WHITE),
    ('SOM　可获取市场', '第 3 年累计签约 40 所 × 加权年费约 5 万元', '约 200 万元/年', 7.4, ORANGE, WHITE),
]
y = 1.75
for i, (name, basis, val, w, c, tc) in enumerate(funnel):
    x = (13.333 - w) / 2
    bar = add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, 1.15, fill=c, radius=0.12)
    shape_text(bar, [([(name, 15, True, tc), ('　' + val, 15, True, WHITE)], PP_ALIGN.CENTER, 0, 1.1),
                     ([(basis, 11.5, False, RGBColor(0xEA, 0xEF, 0xF6) if tc == WHITE else WHITE)], PP_ALIGN.CENTER, 0, 1.1)])
    if i < 2:
        ar = add_shape(s, MSO_SHAPE.DOWN_ARROW, (13.333 - 0.32) / 2, y + 1.13, 0.32, 0.32, fill=GRAY)
    y += 1.5
note = card(s, 9.8, 6.35, 2.35, 0.8, fill=LIGHTBLUE, line=None)
shape_text(note, [([('约占 SAM 5%', 13, True, BLUE), ('即可达成目标', 11, False, DARK)], PP_ALIGN.CENTER, 0, 1.15)])
add_text(s, 0.9, 6.45, 8.6, 0.7,
         [([('在校生约 3,800 万人 · 每年新增新生 1,000 万+ · 朋辈导学场景供给空白', 12.5, False, DARK)], PP_ALIGN.LEFT, 0, 1.1)])

# ================= S6 产品总览 =================
s = new_slide()
header(s, '产品总览：四层技术架构', 6)
layers = [
    ('01', '数据层', '官网公告 · 办事流程 · 课程资料 · 学长经验 → 清洗去重', LIGHTBLUE, BLUE),
    ('02', '知识层', 'Neo4j 知识图谱 + 向量数据库——“数字学长的大脑”', RGBColor(0xBD, 0xD7, 0xEE), BLUE),
    ('03', '服务层', 'RAG 问答引擎 · 个性化推荐 · 上下文感知（情绪识别）', LIGHTBLUE, BLUE),
    ('04', '交互层', 'Electron 跨平台客户端 · 文字/语音交互 · 主动推送', LIGHTBLUE, BLUE),
]
y = 1.6
for num, name, desc, fill, c in layers:
    cd = add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 1.5, y, 10.3, 0.92, fill=fill, radius=0.14)
    sq = add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 1.75, y + 0.21, 0.52, 0.52, fill=c, radius=0.2)
    shape_text(sq, [([(num, 14, True, WHITE)], PP_ALIGN.CENTER, 0, 1.0)])
    add_text(s, 2.55, y + 0.13, 9.0, 0.7,
             [([(name, 16, True, BLUE), ('　' + desc, 12.5, False, DARK)], PP_ALIGN.LEFT, 0, 1.05)])
    if num != '04':
        ar = add_shape(s, MSO_SHAPE.DOWN_ARROW, 6.52, y + 0.9, 0.3, 0.26, fill=GRAY)
    y += 1.2
add_text(s, 1.5, 6.55, 10.3, 0.5,
         [([('技术底座：RAG 约束生成（答案可溯源） + 知识图谱关联推荐 + TextCNN 情绪识别', 12.5, True, BLUE)], PP_ALIGN.CENTER, 0, 1.0)])

# ================= S7 四大功能模块 =================
s = new_slide()
header(s, '四大功能模块：覆盖大学全周期', 7)
mods = [
    ('生活导航', '校园事务实操指引', '校园卡激活 · 宿舍水电 · 图书馆预约', MIDBLUE),
    ('学业规划', '选课与成长指导', '选课组合建议 · 挂科风险提醒 · 学习转型指导', BLUE),
    ('心理疏导', '情绪感知与支持', '情绪识别 · 心理调节方案 · 危机转介提醒', ORANGE),
    ('校园融入', '归属感建设', '社团推荐 · 校园文化传递 · 归属感建设', MIDBLUE),
]
for i, (t, sub, d, c) in enumerate(mods):
    x = 0.9 + (i % 2) * 5.95
    y = 1.65 + (i // 2) * 2.15
    cd = card(s, x, y, 5.6, 1.85, fill=WHITE, line=c, line_w=1.5)
    add_shape(s, MSO_SHAPE.RECTANGLE, x, y, 0.12, 1.85, fill=c)
    add_text(s, x + 0.4, y + 0.22, 5.0, 1.4,
             [([(t, 17, True, c), ('　' + sub, 13, False, DARK)], PP_ALIGN.LEFT, 4, 1.1),
              ([(d, 12, False, GRAY)], PP_ALIGN.LEFT, 0, 1.2)])
ban = add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.9, 6.15, 11.5, 0.75, fill=LIGHTORANGE, radius=0.5)
shape_text(ban, [([('+ 个性化推荐与主动提醒：', 14, True, ORANGE),
                   ('选课截止 · 考试周 · 新生报到等 20+ 类关键节点主动推送', 14, True, DARK)], PP_ALIGN.CENTER, 0, 1.0)])

# ================= S8 差异化能力 =================
s = new_slide()
header(s, '差异化：错位竞争——“办事效率 vs 成长陪伴”', 8)
add_text(s, 0.9, 1.45, 5.7, 0.5, [([('竞品主场：办事效率', 16, True, GRAY)], PP_ALIGN.CENTER, 0, 1.0)])
add_text(s, 6.75, 1.45, 5.7, 0.5, [([('智伴成长：成长陪伴', 16, True, ORANGE)], PP_ALIGN.CENTER, 0, 1.0)])
c1 = card(s, 0.9, 2.0, 5.7, 3.75, fill=BGRAY, line=None)
shape_text(c1, [([('一卡通服务 · 事务办理 · 官方资讯问答', 14, True, DARK)], PP_ALIGN.CENTER, 6, 1.2),
                ([('金智教育（1,200+ 所高校）', 12.5, False, GRAY)], PP_ALIGN.CENTER, 2, 1.2),
                ([('新开普（一卡通约 45% 份额）', 12.5, False, GRAY)], PP_ALIGN.CENTER, 2, 1.2),
                ([('高校自研 AI 助手（清华/浙大等）', 12.5, False, GRAY)], PP_ALIGN.CENTER, 0, 1.2)], anchor=MSO_ANCHOR.MIDDLE)
c2 = card(s, 6.75, 2.0, 5.7, 3.75, fill=LIGHTORANGE, line=ORANGE, line_w=1.5)
shape_text(c2, [([('选课避雷 · 学习方法 · 心理经验 · 社团融入', 14, True, ORANGE)], PP_ALIGN.CENTER, 6, 1.2),
                ([('朋辈经验众包沉淀，代际传承', 12.5, False, DARK)], PP_ALIGN.CENTER, 2, 1.2),
                ([('“数字学长”人格 + 情绪感知', 12.5, False, DARK)], PP_ALIGN.CENTER, 2, 1.2),
                ([('无人系统化覆盖的空白地带', 12.5, False, DARK)], PP_ALIGN.CENTER, 0, 1.2)], anchor=MSO_ANCHOR.MIDDLE)
ov = add_shape(s, MSO_SHAPE.OVAL, 5.79, 3.3, 1.75, 1.05, fill=WHITE, line=ORANGE, line_w=1.5)
shape_text(ov, [([('错位', 16, True, ORANGE)], PP_ALIGN.CENTER, 0, 1.0)])
add_text(s, 0.9, 6.15, 11.5, 0.7,
         [([('策略：合作大于对抗（直销 + 被集成） · 单点极致 · 窗口期 1—2 年', 13.5, True, BLUE)], PP_ALIGN.CENTER, 0, 1.0)])

# ================= S9 核心量化指标 =================
s = new_slide()
header(s, '核心量化指标：差异化必须可验证', 9)
stats = [
    ('≥95%', '高频问答准确率', 'Top 200 问题双盲抽检', BLUE),
    ('100%', '答案附来源可溯源', '系统强制机制', BLUE),
    ('20+ 类', '主动提醒场景', '选课/考试周/报到等', ORANGE),
    ('≥85%', '情绪识别准确率', '负面情绪 5 分钟内关怀', ORANGE),
    ('≥40%', '试点校月活跃率', '次月留存 ≥60%', BLUE),
    ('≤2 秒', '文字问答首字响应', '客户端埋点统计', BLUE),
]
for i, (num, t, d, c) in enumerate(stats):
    x = 0.9 + (i % 3) * 3.95
    y = 1.7 + (i // 3) * 2.15
    cd = card(s, x, y, 3.6, 1.85, fill=WHITE, line=MGRAY)
    add_text(s, x + 0.25, y + 0.22, 3.1, 1.4,
             [([(num, 30, True, c)], PP_ALIGN.LEFT, 2, 1.0),
              ([(t, 13.5, True, DARK)], PP_ALIGN.LEFT, 2, 1.1),
              ([(d, 11, False, GRAY)], PP_ALIGN.LEFT, 0, 1.1)])
add_text(s, 0.9, 6.3, 11.5, 0.6,
         [([('口径说明：以上为 MVP/V1.0 验收目标值 + 验证方式，试点实测校准；', 12, False, GRAY),
            ('未达标不对外宣传', 12, True, ORANGE)], PP_ALIGN.CENTER, 0, 1.0)])

# ================= S10 商业模式 =================
s = new_slide()
header(s, '商业模式：B2B 订阅制，学校付费、学生免费', 10)
ban = add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.9, 1.5, 11.5, 0.7, fill=BLUE, radius=0.5)
shape_text(ban, [([('B2B2C：', 14, True, WHITE), ('高校按年订阅（学生端完全免费） · 辅导员减负是学校的刚性需求', 13.5, False, WHITE)], PP_ALIGN.CENTER, 0, 1.0)])
prices = [
    ('基础版', '2.98 万元/年', '万人以下院校 / 二级学院\nRAG 问答 + 校园知识库', False),
    ('标准版', '4.98 万元/年', '普通本科院校\n+ 知识图谱 · 学业规划 · 主动提醒', True),
    ('旗舰版', '8.8 万元/年', '重点高校 / 多校区\n+ 个性化推荐 · 私有化部署', False),
]
for i, (t, p, d, hot) in enumerate(prices):
    x = 0.9 + i * 3.95
    cd = card(s, x, 2.55, 3.6, 2.7, fill=WHITE, line=(ORANGE if hot else MGRAY), line_w=(2.0 if hot else 1.0))
    if hot:
        chip(s, x + 2.45, 2.38, 1.0, 0.4, '主打', ORANGE, size=12)
    add_text(s, x + 0.3, 2.85, 3.0, 0.5, [([(t, 17, True, BLUE)], PP_ALIGN.CENTER, 0, 1.0)])
    add_text(s, x + 0.3, 3.45, 3.0, 0.6, [([(p, 22, True, ORANGE)], PP_ALIGN.CENTER, 0, 1.0)])
    add_text(s, x + 0.3, 4.15, 3.0, 1.0, [([(ln, 11.5, False, GRAY)], PP_ALIGN.CENTER, 0, 1.2) for ln in d.split('\n')])
add_text(s, 0.9, 5.65, 11.5, 1.0,
         [([('销售结构：基础版 30% / 标准版 55% / 旗舰版 15% → 加权平均年费 4.95 万元', 13, True, DARK)], PP_ALIGN.CENTER, 2, 1.15),
          ([('首年另收实施与数据初始化费 0.5—2 万元（校均 0.8 万元） · 续费率假设 80%', 12, False, GRAY)], PP_ALIGN.CENTER, 0, 1.15)])

# ================= S11 单校经济账 =================
s = new_slide()
header(s, '单校经济账：首年即回本', 11)
add_text(s, 1.0, 1.45, 5.3, 0.5, [([('首年（含实施费）', 16, True, BLUE)], PP_ALIGN.CENTER, 0, 1.0)])
add_text(s, 7.0, 1.45, 5.3, 0.5, [([('续费年（第 2 年起）', 16, True, ORANGE)], PP_ALIGN.CENTER, 0, 1.0)])
rows = [('合同收入', 5.75, 4.95), ('单校毛利', 3.05, 3.05), ('全成本净贡献', 0.75, 1.55)]
UNIT = 0.62  # 每万元宽度（英寸），max 5.75 → 3.57in
y = 2.15
for name, v1, v2 in rows:
    add_text(s, 1.0, y, 1.7, 0.4, [([(name, 12.5, True, DARK)], PP_ALIGN.RIGHT, 0, 1.0)])
    b1 = add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 2.85, y, max(0.25, v1 * UNIT), 0.42, fill=BLUE, radius=0.5)
    add_text(s, 2.95 + v1 * UNIT, y - 0.02, 1.6, 0.45, [([(f'{v1:g} 万', 12, True, BLUE)], PP_ALIGN.LEFT, 0, 1.0)])
    b2 = add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 7.85, y, max(0.25, v2 * UNIT), 0.42, fill=ORANGE, radius=0.5)
    add_text(s, 7.95 + v2 * UNIT, y - 0.02, 1.6, 0.45, [([(f'{v2:g} 万', 12, True, ORANGE)], PP_ALIGN.LEFT, 0, 1.0)])
    y += 0.78
add_text(s, 1.0, y + 0.05, 11.3, 0.4,
         [([('毛利率：首年 53% → 续费年边际毛利率 62%（客户存续越久，利润越厚）', 12.5, False, GRAY)], PP_ALIGN.CENTER, 0, 1.0)])
ban = add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.9, 6.15, 11.5, 0.85, fill=ORANGE, radius=0.5)
shape_text(ban, [([('核心结论：', 15, True, WHITE),
                   ('单校首年即回本 · LTV/CAC ≈ 3.7（健康线为 3）· 公司盈亏平衡点 15—16 所', 15, True, WHITE)], PP_ALIGN.CENTER, 0, 1.0)])

# ================= S12 竞争壁垒 =================
s = new_slide()
header(s, '竞争壁垒：三重壁垒护城河', 12)
walls = [
    ('01', '数据壁垒', '校园私域知识库 + 朋辈经验众包生态\n——厂商与自研系统均不覆盖的数据形态', BLUE, LIGHTBLUE),
    ('02', '场景壁垒', '嵌入选课、报到、考试周关键节点\n“用得上、离不开”的使用习惯', ORANGE, LIGHTORANGE),
    ('03', '角色壁垒', '“数字学长”人格与情感联结\n工具型竞品难以模仿的软性壁垒', MIDBLUE, LIGHTBLUE),
]
for i, (num, t, d, c, fill) in enumerate(walls):
    x = 0.9 + i * 3.95
    cd = card(s, x, 1.7, 3.6, 3.9, fill=fill, line=None)
    add_shape(s, MSO_SHAPE.OVAL, x + 1.43, 2.0, 0.75, 0.75, fill=c)
    add_text(s, x + 1.43, 2.0, 0.75, 0.75, [([(num, 20, True, WHITE)], PP_ALIGN.CENTER, 0, 1.0)])
    add_text(s, x + 0.25, 2.95, 3.1, 0.6, [([(t, 18, True, c)], PP_ALIGN.CENTER, 0, 1.0)])
    add_text(s, x + 0.25, 3.65, 3.1, 1.6, [([(ln, 12, False, DARK)], PP_ALIGN.CENTER, 0, 1.25) for ln in d.split('\n')])
add_text(s, 0.9, 6.05, 11.5, 0.7,
         [([('打法：错位聚焦 · 单点极致 · 以快打慢 —— 1—2 年窗口期内完成 3 校标杆与 40 所目标', 13.5, True, BLUE)], PP_ALIGN.CENTER, 0, 1.0)])

# ================= S13 竞品对比 =================
s = new_slide()
header(s, '竞品对比：错位优势一目了然', 13)
rows, cols = 6, 5
tbl_shape = s.shapes.add_table(rows, cols, Inches(0.9), Inches(1.55), Inches(11.5), Inches(5.0))
tbl = tbl_shape.table
tbl.columns[0].width = Inches(1.7)
for j in range(1, 5):
    tbl.columns[j].width = Inches(2.45)
tbl.rows[0].height = Inches(0.6)
for r in range(1, rows):
    tbl.rows[r].height = Inches(0.8)
data = [
    ['维度', '智伴成长', '厂商 AI 助手', '高校自研', '通用大模型'],
    ['产品形态', '数字学长 · 全周期导学', '服务 App + 办事助手', '校内问答机器人', '通用对话助手'],
    ['朋辈经验', '众包沉淀 · 代际传承', '基本不涉及', '基本不涉及', '无'],
    ['主动提醒', '20+ 类关键场景', '服务办理类为主', '较少', '无'],
    ['情感角色', '学长人设 + 情绪感知', '数字人客服（工具）', '工具属性', '通用人格'],
    ['收费模式', 'B2B 年订阅', '项目制 / 平台费', '校方投入', '免费 / 订阅'],
]
for r in range(rows):
    for c in range(cols):
        cell = tbl.cell(r, c)
        cell.margin_left = Inches(0.08)
        cell.margin_right = Inches(0.08)
        cell.margin_top = Inches(0.03)
        cell.margin_bottom = Inches(0.03)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        if r == 0:
            cell.fill.solid()
            cell.fill.fore_color.rgb = BLUE
            run_tf(run, data[r][c], 13, True, WHITE)
        elif c == 0:
            cell.fill.solid()
            cell.fill.fore_color.rgb = LIGHTBLUE
            run_tf(run, data[r][c], 12, True, DARK)
        elif c == 1:
            cell.fill.solid()
            cell.fill.fore_color.rgb = LIGHTORANGE
            run_tf(run, data[r][c], 12, True, DARK)
        else:
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE
            run_tf(run, data[r][c], 11.5, False, GRAY)
add_text(s, 0.9, 6.85, 11.5, 0.4,
         [([('注：竞品信息依据公开资料整理（截至 2026 年 8 月）；完整版见商业计划书表 6-1。', 10.5, False, GRAY)], PP_ALIGN.LEFT, 0, 1.0)])

# ================= S14 财务预测 =================
s = new_slide()
header(s, '财务预测：三年收入与利润（万元）', 14)
# 图例
add_shape(s, MSO_SHAPE.RECTANGLE, 9.0, 1.45, 0.22, 0.16, fill=BLUE)
add_text(s, 9.28, 1.38, 1.0, 0.3, [([('营业收入', 11, False, DARK)], PP_ALIGN.LEFT, 0, 1.0)])
add_shape(s, MSO_SHAPE.RECTANGLE, 10.15, 1.45, 0.22, 0.16, fill=MGRAY)
add_text(s, 10.43, 1.38, 1.0, 0.3, [([('成本费用', 11, False, DARK)], PP_ALIGN.LEFT, 0, 1.0)])
add_shape(s, MSO_SHAPE.OVAL, 11.32, 1.43, 0.2, 0.2, fill=ORANGE)
add_text(s, 11.58, 1.38, 1.0, 0.3, [([('净利润', 11, False, ORANGE)], PP_ALIGN.LEFT, 0, 1.0)])
years = ['第 1 年（2027）', '第 2 年（2028）', '第 3 年（2029）']
rev = [23.5, 86.3, 209.4]
cost = [26.0, 84.2, 190.7]
net = [-2.5, 3.0, 19.3]
BASE_Y, TOP_Y, MAXV = 6.05, 2.1, 220.0
SC = (BASE_Y - TOP_Y) / MAXV
gx = [2.0, 6.2, 10.4]
bar_w = 1.05
for i in range(3):
    add_text(s, gx[i] - 0.4, 6.2, 2.6, 0.35, [([(years[i], 12.5, True, DARK)], PP_ALIGN.CENTER, 0, 1.0)])
    h1 = rev[i] * SC
    h2 = cost[i] * SC
    add_shape(s, MSO_SHAPE.RECTANGLE, gx[i], BASE_Y - h1, bar_w, h1, fill=BLUE)
    add_shape(s, MSO_SHAPE.RECTANGLE, gx[i] + bar_w + 0.18, BASE_Y - h2, bar_w, h2, fill=MGRAY)
    add_text(s, gx[i] - 0.05, BASE_Y - h1 - 0.38, 1.15, 0.32,
             [([(f'{rev[i]:g}', 13, True, BLUE)], PP_ALIGN.CENTER, 0, 1.0)])
    add_text(s, gx[i] + 1.18, BASE_Y - h2 - 0.38, 1.15, 0.32,
             [([(f'{cost[i]:g}', 11, False, GRAY)], PP_ALIGN.CENTER, 0, 1.0)])
# 净利润橙色折线
ny = [BASE_Y - v * SC for v in net]
for i in range(3):
    add_shape(s, MSO_SHAPE.OVAL, gx[i] + 1.05, ny[i] - 0.1, 0.2, 0.2, fill=ORANGE)
    add_text(s, gx[i] + 0.85, ny[i] - 0.38, 0.8, 0.32,
             [([(f'{net[i]:+.1f}', 12, True, ORANGE)], PP_ALIGN.CENTER, 0, 1.0)])
for i in range(2):
    ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                Inches(gx[i] + 1.15), Inches(ny[i]),
                                Inches(gx[i + 1] + 1.05), Inches(ny[i + 1]))
    ln.line.color.rgb = ORANGE
    ln.line.width = Pt(2.25)
    ln.shadow.inherit = False
add_text(s, 0.9, 1.45, 7.9, 0.5,
         [([('累计签约：4 所 → 16 所 → 40 所', 13, True, DARK)], PP_ALIGN.LEFT, 0, 1.0)])
add_text(s, 0.9, 6.65, 11.5, 0.5,
         [([('第 2 年盈亏平衡 · 第 3 年净利率 9.2%（成长期 SaaS 合理区间）· 第 4 年起向 15%—25% 演进', 12, False, GRAY)], PP_ALIGN.CENTER, 0, 1.0)])

# ================= S15 融资需求 =================
s = new_slide()
header(s, '融资需求：种子轮 50 万元', 15)
add_text(s, 1.0, 2.0, 4.6, 2.5,
         [([('50', 60, True, BLUE), (' 万元', 24, True, BLUE)], PP_ALIGN.LEFT, 0, 1.0),
          ([('种子轮 · 拟出让股权 10%—15%', 15, True, DARK)], PP_ALIGN.LEFT, 10, 1.1),
          ([('启动时点：2027 年 6 月', 13, False, GRAY)], PP_ALIGN.LEFT, 4, 1.2),
          ([('（大创结题 + 试点成效数据确认后）', 13, False, GRAY)], PP_ALIGN.LEFT, 0, 1.2)])
add_text(s, 1.0, 5.0, 4.6, 1.4,
         [([('投资逻辑：', 13.5, True, BLUE)], PP_ALIGN.LEFT, 4, 1.2),
          ([('结题背书 + 试点数据支撑估值', 12.5, False, DARK)], PP_ALIGN.LEFT, 2, 1.2),
          ([('按 3—5 倍市销率，第 3 年估值约 630—1050 万元', 12.5, False, DARK)], PP_ALIGN.LEFT, 0, 1.2)])
chart_data = CategoryChartData()
chart_data.categories = ['产品研发', '试点部署', '市场拓展', '运营储备']
chart_data.add_series('资金用途', (20, 10, 12, 8))
gf = s.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(6.1), Inches(1.6), Inches(4.6), Inches(4.9), chart_data)
ch = gf.chart
ch.has_legend = True
ch.legend.position = XL_LEGEND_POSITION.RIGHT
ch.legend.include_in_layout = False
ch.legend.font.size = Pt(11)
ch.legend.font.color.rgb = DARK
plot = ch.plots[0]
plot.has_data_labels = True
dl = plot.data_labels
dl.show_percentage = True
dl.show_value = False
dl.number_format = '0%'
dl.number_format_is_linked = False
dl.font.size = Pt(11)
dl.font.color.rgb = WHITE
dl.font.bold = True
for pt, c in zip(ch.series[0].points, [BLUE, MIDBLUE, ORANGE, RGBColor(0xA6, 0xC5, 0xE0)]):
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = c
add_text(s, 10.7, 2.2, 1.9, 3.6,
         [([('20 万（40%）', 12, True, BLUE)], PP_ALIGN.LEFT, 8, 1.1),
          ([('10 万（20%）', 12, True, MIDBLUE)], PP_ALIGN.LEFT, 8, 1.1),
          ([('12 万（24%）', 12, True, ORANGE)], PP_ALIGN.LEFT, 8, 1.1),
          ([('8 万（16%）', 12, True, RGBColor(0xA6, 0xC5, 0xE0))], PP_ALIGN.LEFT, 0, 1.1)])

# ================= S16 团队与导师 =================
s = new_slide()
header(s, '团队与导师：技术 × 商业双轮驱动', 16)
team = [
    ('项目负责人 / 算法工程师', '总体统筹 · 商业模式 · RAG 链路', '徐麟阁'),
    ('知识工程师 / 前端工程师', '知识图谱 · 客户端开发 · 交互设计', '董芮铭'),
    ('运营与市场', '内容运营 · 试点推广 · 合规管理', '章宇洲'),
]
c1 = card(s, 0.9, 1.55, 6.1, 4.6, fill=LIGHTBLUE, line=None)
add_text(s, 1.25, 1.75, 5.4, 0.5, [([('核心团队（3 人，一人多角互补）', 15, True, BLUE)], PP_ALIGN.LEFT, 0, 1.0)])
for i, (role, duty, name) in enumerate(team):
    y = 2.35 + i * 0.72
    add_shape(s, MSO_SHAPE.OVAL, 1.25, y + 0.05, 0.16, 0.16, fill=ORANGE)
    add_text(s, 1.6, y, 5.3, 0.65,
             [([(role, 12.5, True, DARK), ('　' + name, 12, False, GRAY)], PP_ALIGN.LEFT, 1, 1.1),
              ([(duty, 11, False, GRAY)], PP_ALIGN.LEFT, 0, 1.1)])
c2 = card(s, 7.3, 1.55, 5.15, 2.5, fill=WHITE, line=ORANGE, line_w=1.5)
add_text(s, 7.6, 1.8, 4.6, 2.0,
         [([('指导教师：陈圣波', 16, True, ORANGE)], PP_ALIGN.LEFT, 6, 1.1),
          ([('技术路线把关 · 商业模式辅导 · 融资与渠道引荐', 12.5, False, DARK)], PP_ALIGN.LEFT, 2, 1.2),
          ([('2026 年 9 月前完成聘任；12 月前签署首份渠道合作备忘录', 11.5, False, GRAY)], PP_ALIGN.LEFT, 0, 1.2)])
c3 = card(s, 7.3, 4.2, 5.15, 1.05, fill=WHITE, line=MGRAY)
add_text(s, 7.6, 4.35, 4.6, 0.8,
         [([('依托学校创新创业学院', 14, True, BLUE)], PP_ALIGN.LEFT, 3, 1.1),
          ([('创业辅导 · 场地支持 · 政策对接', 11.5, False, GRAY)], PP_ALIGN.LEFT, 0, 1.2)])
add_text(s, 7.3, 5.5, 5.15, 0.6,
         [([('已具备：Word2Vec · TextCNN · 医学知识图谱实训经验', 12, True, DARK)], PP_ALIGN.LEFT, 0, 1.1)])

# ================= S17 里程碑与规划 =================
s = new_slide()
header(s, '里程碑：三年三步走', 17)
line_y = 5.35
add_shape(s, MSO_SHAPE.RECTANGLE, 1.4, line_y - 0.02, 10.6, 0.045, fill=MGRAY)
phases = [
    ('验证期', '2026.09 — 2027.06', '校内跑通闭环\nMVP · 成效报告', '签约 1—4 所', BLUE),
    ('放量期', '2027.07 — 2028.12', '公司化运营\n渠道合作 · 盈亏平衡', '累计 16 所 · 年收入约 86 万', ORANGE),
    ('平台期', '2029.01 — 2029.12', '规模化复制\nSaaS 多租户 · 增值服务', '累计 40 所 · 年收入约 210 万', MIDBLUE),
]
for i, (name, t, body, mile, c) in enumerate(phases):
    x = 1.0 + i * 3.95
    cd = card(s, x, 1.6, 3.6, 2.5, fill=WHITE, line=c, line_w=1.5)
    chip(s, x + 1.05, 1.85, 1.5, 0.45, name, c, size=13)
    add_text(s, x + 0.3, 2.5, 3.0, 0.45, [([(t, 13, True, DARK)], PP_ALIGN.CENTER, 0, 1.0)])
    add_text(s, x + 0.3, 3.0, 3.0, 0.9, [([(ln, 11.5, False, GRAY)], PP_ALIGN.CENTER, 0, 1.2) for ln in body.split('\n')])
    ov = add_shape(s, MSO_SHAPE.OVAL, x + 1.55, line_y - 0.16, 0.34, 0.34, fill=c, line=WHITE, line_w=1.5)
    add_text(s, x, 5.55, 3.6, 0.7, [([(mile, 14, True, c)], PP_ALIGN.CENTER, 0, 1.1)])
add_text(s, 0.9, 6.5, 11.5, 0.6,
         [([('衔接大创周期：2026.09 立项 → 2027.06 结题（校内全量部署 + 首份成效报告）→ 公司化运营', 12.5, False, DARK)], PP_ALIGN.CENTER, 0, 1.0)])

# ================= S18 风险与应对 =================
s = new_slide()
header(s, '风险与应对：Top 3', 18)
risks = [
    ('市场风险', '采购周期长、预算不确定', '学院级轻量入口 · 标杆案例背书 · 6 个月现金安全线', BLUE, LIGHTBLUE),
    ('技术风险', '大模型“幻觉”损害口碑', 'RAG 约束生成 + 来源引用 · 白名单答案库 · 知识审核机制', ORANGE, LIGHTORANGE),
    ('数据合规', '学生隐私与安全审查', '私有化部署 · 最小化采集 · 脱敏训练 · 配合校方审查', MIDBLUE, LIGHTBLUE),
]
for i, (t, r, m, c, fill) in enumerate(risks):
    x = 0.9 + i * 3.95
    cd = card(s, x, 1.8, 3.6, 3.5, fill=fill, line=None)
    add_text(s, x + 0.25, 2.1, 3.1, 0.6, [([(t, 18, True, c)], PP_ALIGN.CENTER, 0, 1.0)])
    add_text(s, x + 0.25, 2.8, 3.1, 0.8, [([(r, 12.5, False, DARK)], PP_ALIGN.CENTER, 0, 1.2)])
    add_shape(s, MSO_SHAPE.RECTANGLE, x + 1.3, 3.7, 1.0, 0.03, fill=c)
    add_text(s, x + 0.25, 3.9, 3.1, 1.2, [([('应对：', 12.5, True, c), (m, 12, False, DARK)], PP_ALIGN.CENTER, 0, 1.25)])
add_text(s, 0.9, 5.75, 11.5, 0.8,
         [([('其余风险（政策/经营）应对详见商业计划书表 10-1；', 12, False, GRAY),
            ('整体原则：弹性成本与签约进度联动，保守扩张', 12, True, DARK)], PP_ALIGN.CENTER, 0, 1.1)])

# ================= S19 结语 =================
s = new_slide()
add_shape(s, MSO_SHAPE.RECTANGLE, 0, 0, 13.334, 7.5, fill=BLUE)
add_shape(s, MSO_SHAPE.RECTANGLE, 0, 7.38, 13.334, 0.12, fill=ORANGE)
add_text(s, 0.9, 2.35, 11.5, 1.2,
         [([('让每一段大学时光，都有智慧与温暖相伴', 34, True, WHITE)], PP_ALIGN.CENTER, 0, 1.0)])
add_shape(s, MSO_SHAPE.RECTANGLE, 5.77, 3.95, 1.8, 0.05, fill=ORANGE)
add_text(s, 0.9, 4.25, 11.5, 0.6,
         [([('从“数字学长”出发，构建可传承的校园知识生态', 17, False, RGBColor(0xD9, 0xE2, 0xF3))], PP_ALIGN.CENTER, 0, 1.0)])
add_text(s, 0.9, 5.6, 11.5, 1.0,
         [([('恳请各位老师批评指正 · 谢谢聆听', 20, True, RGBColor(0xF4, 0xB1, 0x83))], PP_ALIGN.CENTER, 0, 1.0)])

prs.save(OUT)
print('已生成:', OUT, f'（共 {len(prs.slides.__iter__.__self__._sldIdLst)} 页）')
