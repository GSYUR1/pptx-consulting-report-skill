# -*- coding: utf-8 -*-
"""版式族 Builder 库 —— 按 assets/版式规格.md 实现，13 个版式族全覆盖。

用法:
    from builders import new_deck, save_deck
    from builders import build_cover, build_toc, ...   # 见各函数 docstring
    prs, manifest = new_deck('A')
    build_cover(prs, title=..., subtitle=..., meta_lines=[...])
    save_deck(prs, 'out.pptx', manifest)

约定（与版式规格.md 一致）:
- 颜色一律来自 design_system.json，不硬编码业务色
- 字号阶梯降级 FONT_LADDER，容量不足时自动降档
- 每页构建结果记入 manifest（元素清单 JSON），供 QA 比对
"""
import json
import math
import os
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS = os.path.join(SKILL_DIR, "assets")
DS = json.load(open(os.path.join(ASSETS, "design_system.json"), encoding="utf-8"))

TEMPLATES = {
    "A": os.path.join(ASSETS, "template_illum.pptx"),
    "B": os.path.join(ASSETS, "template_iqvia.pptx"),
}

STYLE = {
    "A": {
        "主题色": DS["风格A_ILLUMINERA青"]["主题色"],
        "西文": DS["风格A_ILLUMINERA青"]["字体"]["西文"],
        "中文": DS["风格A_ILLUMINERA青"]["字体"]["中文"],
        "layout_keys": {"cover": "封面", "toc": "目录", "section": "分割页",
                        "content": "单行标题", "content2": "双行标题"},
        "标题y": 1.69, "标题x": 0.94, "标题w": 30.74,
    },
    "B": {
        "主题色": DS["风格B_IQVIA深蓝"]["主题色"],
        "西文": DS["风格B_IQVIA深蓝"]["字体"]["西文"],
        "中文": DS["风格B_IQVIA深蓝"]["字体"]["西文"],
        "layout_keys": {"cover": "Cover 2024", "toc": "Title and Content",
                        "section": "Cover 2024", "content": "Title and Content"},
        "标题y": 0.92, "标题x": 0.93, "标题w": 32.01,
    },
}

# 页面常量（设计坐标系基准 = 33.87 x 19.05 cm）
PAGE_W, PAGE_H = 33.87, 19.05
# 实际模板画布缩放系数：new_deck 时按模板真实尺寸设定（如 illum 模板为 25.4 宽 → 0.75）
SCALE = 1.0


def U(v):
    """设计坐标(cm) -> 按当前模板画布缩放的 Length。所有几何坐标必须经此转换。"""
    return Cm(v * SCALE)


def S(pt_size):
    """字号按画布比例缩放（25.4cm 画布上的 12pt 视觉效果 = 33.87 画布的 16pt）。"""
    return Pt(pt_size * (SCALE ** 0.5))
FONT_LADDER = [14, 12, 10.5, 9]          # 字号阶梯降级
EA_FONT = "微软雅黑"


# ---------------------------------------------------------------- 基础工具

def C(hexstr):
    """'#1EB3C8' -> RGBColor"""
    return RGBColor.from_string(hexstr.lstrip("#"))


def theme(style, key):
    return C(STYLE[style]["主题色"][key])


def tint(color, ratio):
    """向白色混入 ratio(0~1) 的比例，近似 lumMod/lumOff 浅色系。
    接受 hex 字符串或 RGBColor。"""
    if isinstance(color, RGBColor):
        r, g, b = color[0], color[1], color[2]
    else:
        color_hex = str(color)
        r = int(color_hex[1:3], 16)
        g = int(color_hex[3:5], 16)
        b = int(color_hex[5:7], 16)
    mix = lambda v: int(v + (255 - v) * ratio)
    return RGBColor(mix(r), mix(g), mix(b))


def _fit_font(text, base, max_chars_per_line, max_lines):
    """按容量选择字号：超长按 FONT_LADDER 降档。"""
    size = base
    for s in FONT_LADDER:
        size = s
        if s <= base:
            lines = math.ceil(len(text) / (max_chars_per_line * s / base))
            if lines <= max_lines:
                break
    return size


def new_deck(style="A"):
    """打开模板资产，返回 (prs, manifest)。按模板真实画布尺寸设定缩放系数。"""
    global SCALE
    prs = Presentation(TEMPLATES[style])
    SCALE = prs.slide_width / 360000 / PAGE_W
    return prs, []


def save_deck(prs, path, manifest=None):
    prs.save(path)
    if manifest is not None:
        mp = os.path.splitext(path)[0] + "_manifest.json"
        json.dump(manifest, open(mp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return path


def _find_layout(prs, style, kind):
    key = STYLE[style]["layout_keys"].get(kind, "")
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            if key and key in layout.name:
                return layout
    return prs.slide_masters[0].slide_layouts[0]


def _add_slide(prs, style, kind="content"):
    slide = prs.slides.add_slide(_find_layout(prs, style, kind))
    # 删除母版继承的占位符：空占位在 OfficeCLI 渲染时会显示版式提示 ghost text
    for ph in list(slide.placeholders):
        ph._element.getparent().remove(ph._element)
    return slide


def _set_ea(run):
    """设置东亚字体，保证中文回退正确。"""
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        rPr.append(ea)
    ea.set("typeface", EA_FONT)


def textbox(slide, x, y, w, h, text, size=12, bold=False, color=None,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, fill=None,
            line=None, font=None, wrap=True, pt=None):
    """通用文本框/色块。坐标 cm。pt 为绝对字号(pt)，设置后跳过 S() 缩放。
    跨画布复刻原稿标题等大字号元素时必须用 pt，保证视觉大小一致。"""
    style_font = font or "Calibri"
    tb = slide.shapes.add_textbox(U(x), U(y), U(w), U(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = U(0.1)
    tf.margin_top = tf.margin_bottom = U(0.05)
    lines = text.split("\n") if isinstance(text, str) else list(text)
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(pt) if pt else S(size)
        r.font.bold = bold
        r.font.name = style_font
        _set_ea(r)
        if color is not None:
            r.font.color.rgb = color
    if fill is not None:
        tb.fill.solid()
        tb.fill.fore_color.rgb = fill
    else:
        tb.fill.background()
    if line is not None:
        tb.line.color.rgb = line
        tb.line.width = Pt(0.75)
    else:
        tb.line.fill.background()
    return tb


def rect(slide, x, y, w, h, fill=None, line=None, rounded=False, shape=None):
    kind = shape or (MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE)
    sp = slide.shapes.add_shape(kind, U(x), U(y), U(w), U(h))
    if fill is not None:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    else:
        sp.fill.background()
    if line is not None:
        sp.line.color.rgb = line
        sp.line.width = Pt(0.75)
    else:
        sp.line.fill.background()
    sp.shadow.inherit = False
    return sp


def hline(slide, x, y, w, color, weight=1.0):
    ln = slide.shapes.add_connector(1, U(x), U(y), U(x + w), U(y))
    ln.line.color.rgb = color
    ln.line.width = Pt(weight)
    return ln


def oval(slide, x, y, d, fill, text="", size=7, color=None):
    sp = slide.shapes.add_shape(MSO_SHAPE.OVAL, U(x), U(y), U(d), U(d))
    sp.fill.solid()
    sp.fill.fore_color.rgb = fill
    sp.line.fill.background()
    sp.shadow.inherit = False
    if text:
        tf = sp.text_frame
        tf.margin_left = tf.margin_right = U(0)
        tf.margin_top = tf.margin_bottom = U(0)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = text
        r.font.size = S(size)
        r.font.bold = True
        r.font.name = "Calibri"
        r.font.color.rgb = color or C("#FFFFFF")
    return sp


def _title(slide, style, text, size=20):
    """Action title（页首大标题）。
    样式规范（源自原稿版式 ph idx=12）：绝对字号 风格A=18pt / 风格B=24pt、
    常规字重、#465059、底缘锚定（盒 y1.69 h0.93，文字向上溢出，不占下方 Special Note 区）。
    注意：字号必须用 pt 绝对值，不能经 S() 缩放，否则视觉上偏大并与 Note 重叠。"""
    st = STYLE[style]
    abs_pt = {"A": 18, "B": 24}[style]
    tb = textbox(slide, st["标题x"], st["标题y"], st["标题w"], 0.93, text,
                 size=size, pt=abs_pt, bold=False, color=C("#465059"),
                 anchor=MSO_ANCHOR.BOTTOM)
    return tb


def _log(manifest, family, elements):
    manifest.append({"版式族": family, "元素数": len(elements), "元素": elements})


# ---------------------------------------------------------------- 通用元素

def build_chapter_bar(slide, style, sections, current, x=0.11, y=0.03):
    """顶部章节进度条：五边形 + V 形箭头，当前章 accent 色高亮。"""
    main = theme(style, "主青" if style == "A" else "主蓝")
    w = 4.96
    shapes = [MSO_SHAPE.PENTAGON] + [MSO_SHAPE.CHEVRON] * (len(sections) - 1)
    for i, name in enumerate(sections):
        cur = (i == current)
        fill = main if cur else tint("#FFFFFF" if style == "A" else "#F3F3F3", 0)
        fill = main if cur else C("#E9E9E9")
        sp = rect(slide, x + i * (w - 0.35), y, w, 0.71,
                  fill=fill, shape=shapes[i])
        tf = sp.text_frame
        tf.margin_left = U(0.2)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = name
        r.font.size = S(8)
        r.font.bold = True
        r.font.name = STYLE[style]["西文"]
        r.font.color.rgb = C("#FFFFFF") if cur else C("#999999")
    return slide


# ---------------------------------------------------------------- F01 封面

def build_cover(prs, style, title, subtitle="", meta_lines=None, manifest=None):
    slide = _add_slide(prs, style, "cover")
    m = manifest if manifest is not None else []
    st = STYLE[style]
    if style == "A":
        textbox(slide, 2.42, 4.36, 12.27, 2.07, title, size=28, bold=True,
                color=theme(style, "主青"))
        textbox(slide, 2.42, 6.47, 12.05, 1.05, subtitle, size=16, color=C("#465059"))
        for i, ln in enumerate(meta_lines or []):
            textbox(slide, 2.53, 8.10 + i * 0.73, 9.33, 0.5, ln, size=11,
                    color=C("#465059"))
    else:
        textbox(slide, 2.0, 7.0, 24.0, 2.2, title, size=32, bold=True, color=C("#FFFFFF"))
        textbox(slide, 2.0, 9.4, 24.0, 1.2, subtitle, size=16, color=tint_hex("B", 0.4))
        for i, ln in enumerate(meta_lines or []):
            textbox(slide, 2.0, 11.2 + i * 0.8, 20.0, 0.6, ln, size=12,
                    color=tint_hex("B", 0.4))
    _log(m, "F01", [title, subtitle] + list(meta_lines or []))
    return slide


def tint_hex(style, ratio):
    key = "主青" if style == "A" else "主蓝"
    hexv = STYLE[style]["主题色"][key]
    return tint(hexv, ratio)


# ---------------------------------------------------------------- F02 目录

def build_toc(prs, style, chapters, current=0, manifest=None):
    slide = _add_slide(prs, style, "toc")
    m = manifest if manifest is not None else []
    main = theme(style, "主青" if style == "A" else "主蓝")
    deep = theme(style, "深蓝")
    n = len(chapters)
    if style == "A":
        top = 3.0
        rowh = min(1.6, 12.0 / max(n, 1))
        for i, ch in enumerate(chapters):
            y = top + i * rowh
            cur = (i == current)
            rect(slide, 2.41, y, 0.9, rowh - 0.25, fill=deep if cur else tint_hex_to(deep, 0.75))
            textbox(slide, 2.41, y, 0.9, rowh - 0.25, f"{i+1:02d}", size=12, bold=True,
                    color=C("#FFFFFF"), align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            textbox(slide, 3.6, y, 18.0, rowh - 0.25, ch, size=14 if cur else 12,
                    bold=cur, color=C("#404040"), anchor=MSO_ANCHOR.MIDDLE)
    else:
        x0, y0, gap = 18.58, 8.21, 1.6
        for i, ch in enumerate(chapters[:8]):
            y = y0 + i * gap
            cur = (i == current)
            rect(slide, x0, y, 1.03, 1.03, fill=main if cur else C("#C9C9C9"))
            textbox(slide, x0, y, 1.03, 1.03, str(i + 1), size=16, bold=True,
                    color=C("#FFFFFF"), align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            textbox(slide, x0 + 1.03, y, 12.0, 1.03, ch, size=16, bold=True,
                    color=C("#2B3A42") if cur else C("#909090"),
                    anchor=MSO_ANCHOR.MIDDLE)
    _log(m, "F02", list(chapters))
    return slide


def tint_hex_to(color_hexstr, ratio):
    return tint(color_hexstr, ratio)


# ---------------------------------------------------------------- F03 章节分隔

def build_section(prs, style, number, title, manifest=None):
    slide = _add_slide(prs, style, "section")
    m = manifest if manifest is not None else []
    if style == "A":
        textbox(slide, 13.04, 6.38, 12.36, 1.05,
                f"{number}. {title}" if number else title,
                size=30, bold=True, color=C("#404040"))
    else:
        textbox(slide, 6.0, 8.0, 22.0, 1.8, title, size=36, bold=True,
                color=C("#FFFFFF"), align=PP_ALIGN.CENTER)
    _log(m, "F03", [number, title])
    return slide


# ---------------------------------------------------------------- F04 过渡箭头

def build_transition(prs, style, stages, notes=None, manifest=None):
    slide = _add_slide(prs, style, "content")
    m = manifest if manifest is not None else []
    main = theme(style, "主青" if style == "A" else "主蓝")
    n = len(stages)
    y0 = 4.49 if n <= 4 else 3.2
    gap = (12.5 - y0) / max(n - 1, 1) if n > 1 else 0
    for i, s in enumerate(stages):
        y = y0 + i * max(gap, 1.5)
        rect(slide, 11.12, y, 13.38, 0.69, fill=tint(main, 0.75))
        textbox(slide, 11.32, y, 13.0, 0.69, s, size=14, bold=True,
                color=C("#404040"), anchor=MSO_ANCHOR.MIDDLE)
        rect(slide, 14.82, y + 0.79, 0.53, 0.11, fill=main)
        if notes and i < len(notes) and notes[i]:
            fs = _fit_font(notes[i], 12, 60, 2)
            textbox(slide, 15.35, y + 0.6, 17.31, 0.91, notes[i], size=fs,
                    color=C("#465059"))
    _log(m, "F04", list(stages))
    return slide


# ---------------------------------------------------------------- F05 行标签总览

def build_row_overview(prs, title, rows, cols, cells, manifest=None):
    """rows: [行标签...]; cols: [列标题...]; cells: {行号: {列号: 文本}}"""
    slide = _add_slide(prs, "A", "content")
    m = manifest if manifest is not None else []
    _title(slide, "A", title)
    main = theme("A", "主青")
    deep = theme("A", "深蓝")
    x_label, w_label = 1.03, 2.15
    x0, x1 = 3.47, 24.28
    n = len(cols)
    cw = (x1 - x0) / n
    y = 2.69
    for ci, col in enumerate(cols):
        textbox(slide, x0 + ci * cw, y, cw - 0.2, 0.6, col, size=11, bold=True,
                color=C("#404040"), anchor=MSO_ANCHOR.MIDDLE)
    body_rows = rows[1:]
    rh = min(3.0, 11.5 / max(len(body_rows), 1))
    for ri, rl in enumerate(body_rows):
        ry = 3.6 + ri * rh
        rect(slide, x_label, ry, w_label, rh - 0.2, fill=C("#F2F2F2"))
        textbox(slide, x_label, ry, w_label, rh - 0.2, rl, size=9, bold=True,
                color=C("#595959"), align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        for ci in range(n):
            txt = cells.get(ri, {}).get(ci, "")
            if not txt:
                continue
            fs = _fit_font(txt, 9, int((cw - 0.4) * 1.6), 4)
            textbox(slide, x0 + ci * cw, ry, cw - 0.25, rh - 0.2, txt, size=fs,
                    color=deep, anchor=MSO_ANCHOR.TOP)
    _log(m, "F05", {"行": body_rows, "列": cols})
    return slide


# ---------------------------------------------------------------- F06 标准内容页

def build_content(prs, style, title, subhead="", blocks=None, manifest=None):
    """blocks: [ (栏标题, [ (小标题, 说明), ... ]) ]  或 [bullet, ...]"""
    slide = _add_slide(prs, style, "content")
    m = manifest if manifest is not None else []
    _title(slide, style, title)
    main = theme(style, "主青" if style == "A" else "主蓝")
    top = STYLE[style]["标题y"] + 1.9
    if subhead:
        textbox(slide, STYLE[style]["标题x"], top, STYLE[style]["标题w"], 0.9,
                subhead, size=14, bold=True, color=C("#2B3A42"))
        top += 1.2
    bottom = 16.9 if style == "B" else 15.4
    if not blocks:
        _log(m, "F06", [title])
        return slide
    # 简单分栏：单组=通栏，多组=等分栏
    is_tuples = isinstance(blocks[0], (tuple, list))
    if not is_tuples:
        fs = _fit_font("\n".join(blocks), 12, 90, 10)
        textbox(slide, STYLE[style]["标题x"] + 0.5, top + 0.3,
                STYLE[style]["标题w"] - 1.0, bottom - top,
                "\n".join("• " + b for b in blocks), size=fs, color=C("#404040"))
        _log(m, "F06", list(blocks))
        return slide
    n = len(blocks)
    x0 = STYLE[style]["标题x"]
    total = STYLE[style]["标题w"]
    gap = 0.5
    cw = (total - gap * (n - 1)) / n
    for i, (head, items) in enumerate(blocks):
        cx = x0 + i * (cw + gap)
        rect(slide, cx, top, cw, 0.6, fill=main)
        textbox(slide, cx, top, cw, 0.6, head, size=11, bold=True,
                color=C("#FFFFFF"), align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        lines = []
        for it in items:
            if isinstance(it, (tuple, list)):
                lines.append(f"▪ {it[0]}：{it[1]}")
            else:
                lines.append(f"▪ {it}")
        fs = _fit_font("\n".join(lines), 10, int(cw * 1.6), int((bottom - top) / 0.55))
        textbox(slide, cx, top + 0.8, cw, bottom - top - 0.8,
                "\n".join(lines), size=fs, color=C("#404040"))
    _log(m, "F06", [b[0] for b in blocks])
    return slide


# ---------------------------------------------------------------- F07 双栏对比

def build_compare(prs, style, columns, with_charts=False, manifest=None):
    """columns: [ (栏目标题, [bullet...]), ... ]  2~3 栏"""
    slide = _add_slide(prs, style, "content")
    m = manifest if manifest is not None else []
    main = theme(style, "主青" if style == "A" else "主蓝")
    n = len(columns)
    x0, gap = 1.48, 0.5
    total = 17.0 if n <= 2 else STYLE[style]["标题w"] - 1.0
    cw = (total - gap * (n - 1)) / n
    for i, (head, bullets) in enumerate(columns):
        cx = x0 + i * (cw + gap)
        hline(slide, cx, 4.82, cw, main, 1.0)
        rect(slide, cx, 4.39, min(cw, 8.28), 0.85, fill=main)
        textbox(slide, cx, 4.39, min(cw, 8.28), 0.85, head, size=14, bold=True,
                color=C("#FFFFFF"), align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        box_h = 8.52
        rect(slide, cx, 5.65, cw, box_h, fill=tint(main, 0.8))
        body = "\n\n".join("• " + b for b in bullets)
        fs = _fit_font(body, 12, int(cw * 1.6), 12)
        textbox(slide, cx + 0.35, 6.1, cw - 0.7, box_h - 0.9, body, size=fs,
                color=C("#333333"), anchor=MSO_ANCHOR.MIDDLE)
    _log(m, "F07", [c[0] for c in columns])
    return slide


# ---------------------------------------------------------------- F08 表格矩阵

MATRIX_PALETTE = ["#1EB3C8", "#6DDAE9", "#CEF3F8", "#D9D9D9"]
MATRIX_PALETTE_B = ["#00A3E0", "#3FBAEA", "#80D1EF", "#BFE8F8"]


def build_matrix(prs, style, headers, row_labels, cells, manifest=None,
                 cell_labels=None):
    """矩阵热力表。
    headers: 列标题; row_labels: 行标题; cells: {行号: {列号: 级别0~3}};
    cell_labels: {行号: {列号: 单元格文字}}（可选）"""
    slide = _add_slide(prs, style, "content")
    m = manifest if manifest is not None else []
    _title(slide, style, headers[0] if headers else "")
    pal = MATRIX_PALETTE if style == "A" else MATRIX_PALETTE_B
    x0, y0 = 0.74, 3.24
    lw = 1.78 if style == "A" else 2.0
    nr, nc = len(row_labels), len(headers) - 1
    cw = min(2.02, (STYLE[style]["标题w"] - lw - 6.0) / max(nc, 1))
    main = theme(style, "主青" if style == "A" else "主蓝")
    rect(slide, x0 + lw, y0 - 0.5, nc * cw, 0.5, fill=tint(main, 0.6))
    for ci in range(nc):
        textbox(slide, x0 + lw + ci * cw, y0 - 0.5, cw, 0.5, headers[ci + 1],
                size=8, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    rh = min(3.0, 11.5 / max(nr, 1))
    for ri, rl in enumerate(row_labels):
        ry = y0 + ri * rh
        rect(slide, x0, ry, lw, rh - 0.15, fill=C("#F0F6F8") if ri % 2 == 0 else C("#CEF3F8"))
        textbox(slide, x0, ry, lw, rh - 0.15, rl, size=8, bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        for ci in range(nc):
            lv = cells.get(ri, {}).get(ci)
            if lv is None:
                continue
            rect(slide, x0 + lw + ci * cw, ry, cw - 0.1, rh - 0.15, fill=C(pal[lv]))
            lbl = (cell_labels or {}).get(ri, {}).get(ci, "")
            if lbl:
                textbox(slide, x0 + lw + ci * cw, ry, cw - 0.1, rh - 0.15, lbl,
                        size=7, color=C("#FFFFFF") if lv <= 1 else C("#595959"),
                        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    _log(m, "F08", {"行": row_labels, "列": headers[1:]})
    return slide


# ---------------------------------------------------------------- F09 流程/路径

def build_flow(prs, style, stages, nodes, conclusion=None, manifest=None):
    """stages: [阶段名...]（左列）; nodes: {阶段号: [节点文本...]};
    conclusion: 底部结论条文字（可选）"""
    slide = _add_slide(prs, style, "content")
    m = manifest if manifest is not None else []
    _title(slide, style, stages[0] if stages else "")
    main = theme(style, "主青" if style == "A" else "主蓝")
    accent4 = C(pal_aux(style))
    x0 = 2.32
    top, bottom = 3.0, 15.0
    rh = (bottom - top) / max(len(stages), 1)
    for si, st in enumerate(stages):
        sy = top + si * rh
        rect(slide, x0, sy, 2.02, rh - 0.2, fill=tint(accent4, 0.2 + 0.2 * si))
        textbox(slide, x0, sy, 2.02, rh - 0.2, st, size=8, bold=True,
                color=C("#FFFFFF"), align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        ns = nodes.get(si, [])
        nx = 4.49
        nw = (24.3 - nx) / max(len(ns), 1)
        for ni, nd in enumerate(ns):
            cx = nx + ni * nw
            oval(slide, cx + 0.15, sy + (rh - 0.2) / 2 - 0.27, 0.54, main,
                 text=str(ni + 1))
            fs = _fit_font(nd, 8, int(nw * 1.6), 3)
            textbox(slide, cx + 0.85, sy, nw - 0.9, rh - 0.2, nd, size=fs,
                    bold=True, color=main, anchor=MSO_ANCHOR.MIDDLE)
        if si < len(stages) - 1:
            hline(slide, x0, sy + rh - 0.1, 21.5, theme(style, "绿"), 0.75)
    if conclusion:
        rect(slide, 1.2, 15.1, 31.3, 2.3, fill=theme(style, "深蓝") if style == "A" else C("#0070C0"))
        textbox(slide, 1.6, 15.1, 30.5, 2.3, conclusion, size=12, bold=True,
                color=C("#FFFFFF"), anchor=MSO_ANCHOR.MIDDLE)
    _log(m, "F09", {"阶段": stages, "节点数": sum(len(v) for v in nodes.values())})
    return slide


def pal_aux(style):
    return STYLE[style]["主题色"]["浅青" if style == "A" else "深蓝"]


# ---------------------------------------------------------------- F10 图表+洞察

def build_chart_insight(prs, style, title, chart_data, chart_type="bar",
                        insight="", subhead="", manifest=None):
    """chart_data: {类别: 数值} 简单单系列图 + 洞察条。
    复杂图表请用 python-pptx chart API 自行扩展或保留草稿原图表。"""
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    slide = _add_slide(prs, style, "content")
    m = manifest if manifest is not None else []
    _title(slide, style, title)
    main = theme(style, "主青" if style == "A" else "主蓝")
    if subhead:
        textbox(slide, STYLE[style]["标题x"], STYLE[style]["标题y"] + 1.9,
                STYLE[style]["标题w"], 0.9, subhead, size=14, bold=True,
                color=C("#2B3A42"))
    cd = CategoryChartData()
    cd.categories = list(chart_data.keys())
    cd.add_series("数值", list(chart_data.values()))
    ct = XL_CHART_TYPE.COLUMN_CLUSTERED if chart_type == "bar" else XL_CHART_TYPE.PIE
    gx, gy, gw, gh = (1.13, 4.6, 19.0, 10.5) if style == "B" else (1.48, 4.6, 18.0, 10.0)
    gf = slide.shapes.add_chart(ct, U(gx), U(gy), U(gw), U(gh), cd)
    chart = gf.chart
    chart.has_legend = False
    try:
        chart.series[0].format.fill.solid()
        chart.series[0].format.fill.fore_color.rgb = main
    except Exception:
        pass
    if insight:
        if style == "A":
            rect(slide, 21.0, 6.0, 11.5, 6.0, fill=main, rounded=True)
            textbox(slide, 21.3, 6.2, 10.9, 5.6, insight, size=10,
                    color=C("#FFFFFF"), anchor=MSO_ANCHOR.MIDDLE)
        else:
            rect(slide, 20.5, 5.5, 12.0, 3.0, fill=main)
            textbox(slide, 20.7, 5.6, 11.6, 2.8, insight, size=10,
                    color=C("#FFFFFF"), anchor=MSO_ANCHOR.MIDDLE)
    _log(m, "F10", {"title": title, "chart": chart_type})
    return slide


# ---------------------------------------------------------------- F11 图文/引述

def build_quote(prs, style, quotes, images=None, note="", title="", manifest=None):
    """quotes: [ (引述人/来源, 引述文字) ]; images: [图片路径...]（可选）"""
    slide = _add_slide(prs, style, "content")
    m = manifest if manifest is not None else []
    main = theme(style, "主青" if style == "A" else "主蓝")
    if title:
        _title(slide, style, title)
    main2 = theme(style, "主青" if style == "A" else "主蓝")
    if images:
        iw = 6.92
        for i, img in enumerate(images[:3]):
            slide.shapes.add_picture(img, U(2.03 + i * (iw + 0.7)), U(11.87),
                                     U(iw), U(3.19))
    x0 = 19.67 if images else 1.48
    y = 5.4
    for i, (who, q) in enumerate(quotes[:3]):
        bh = 2.0
        rect(slide, x0, y, 12.71 if images else STYLE[style]["标题w"] - 2.0,
             bh, fill=tint(main2, 0.95), rounded=True)
        textbox(slide, x0 + 0.3, y + 0.15,
                (12.1 if images else STYLE[style]["标题w"] - 2.6), bh - 0.3,
                f"“{q}” —— {who}", size=10, color=C("#2B3A42"),
                anchor=MSO_ANCHOR.MIDDLE)
        y += bh + 0.4
    if note:
        textbox(slide, 0.7, 13.68, 19.89, 0.6, note, size=8, color=C("#FFFFFF"))
    _log(m, "F11", [q[0] for q in quotes])
    return slide


# ---------------------------------------------------------------- F12 总结

def build_summary(prs, style, points, manifest=None):
    """points: [ (加粗结论句, 支撑说明), ... ] 3~5 条"""
    slide = _add_slide(prs, style, "content")
    m = manifest if manifest is not None else []
    main = theme(style, "主青" if style == "A" else "主蓝")
    accent2 = theme(style, "橙" if style == "A" else "金黄")
    n = len(points)
    y0, gap = 3.6, min(2.8, 11.5 / max(n, 1))
    for i, (head, body) in enumerate(points):
        y = y0 + i * gap
        oval(slide, 2.61, y + 0.2, 0.51, accent2, text=str(i + 1), size=11)
        rect(slide, 3.3, y, 26.0, min(gap - 0.4, 1.6), fill=tint(main, 0.85))
        textbox(slide, 3.6, y + 0.12, 25.4, min(gap - 0.4, 1.6) - 0.2,
                f"{head}  {body}", size=10, color=C("#2B3A42"),
                anchor=MSO_ANCHOR.MIDDLE)
    _log(m, "F12", [p[0] for p in points])
    return slide


# ---------------------------------------------------------------- F13 附录/结束

def build_appendix_chart(prs, style, sections, manifest=None):
    """sections: [ (栏目标题, 图表或占位说明), ... ] 1~2 栏"""
    slide = _add_slide(prs, style, "content")
    m = manifest if manifest is not None else []
    main = theme(style, "主青" if style == "A" else "主蓝")
    deep = theme(style, "深蓝")
    n = len(sections)
    for i, (head, body) in enumerate(sections[:2]):
        x = 1.19 + i * 16.4
        rect(slide, x, 3.7, 4.6 if len(head) < 12 else 6.0, 1.07,
             fill=deep if i == 0 else main)
        textbox(slide, x, 3.7, 4.6 if len(head) < 12 else 6.0, 1.07, head,
                size=16, bold=True, color=C("#FFFFFF"),
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        textbox(slide, x, 5.2, 15.1, 10.0, body, size=10, color=C("#404040"))
        textbox(slide, x + 0.2, 16.1, 15.1, 0.64, "Source: ", size=9,
                color=C("#000000"))
    _log(m, "F13", [s[0] for s in sections])
    return slide


def build_thanks(prs, style, manifest=None):
    slide = _add_slide(prs, style, "cover")
    m = manifest if manifest is not None else []
    textbox(slide, 6.0, 8.0, 22.0, 2.0, "Thanks", size=44, bold=True,
            color=C("#FFFFFF") if style == "B" else theme(style, "主青"),
            align=PP_ALIGN.CENTER)
    _log(m, "F13", ["Thanks"])
    return slide
