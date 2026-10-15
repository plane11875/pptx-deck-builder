#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_slides.py — 生成 8 页 SVG 幻灯片(设计系统集中定义)。

为什么要先写 SVG:PPT Master(https://github.com/hugohe3/ppt-master)的编译器会把
受支持的 SVG 元素转换成 PowerPoint **原生 DrawingML 形状 / 文本框**。所以这些
SVG 就是本次演示文稿的"源代码":可 diff、可评审、可版本化。

结构约定(来自 ppt-master 的 SVG 契约,质检脚本会强制):
  * 根 <svg> 声明 lang 与 data-pptx-page-role(cover / toc / section / content / ending);
  * 页面框架与页眉页脚放入带 data-pptx-role 的组(decoration / chrome);
  * 其余元素按逻辑单元分组,每个根级 <g> 自动声明 data-pptx-bounds 子画布;
  * 多行文字是一个 <text> + 若干 <tspan dy>,这样在 PowerPoint 里仍然是一个文本框。

输出: ../slides/slide_01.svg ... slide_08.svg
用法: python3 builder/make_slides.py
"""

from html import escape
from pathlib import Path
from xml.etree import ElementTree as ET

# ── 画布与设计系统 ────────────────────────────────────────────────────────────
W, H = 1280, 720                      # 16:9,对应 ppt-master 的 ppt169 画布
PAGES = 8
SVG_NS = "{http://www.w3.org/2000/svg}"

INK_BG      = "#0B1220"               # 背景
INK_SURFACE = "#121C2F"               # 卡片面
INK_RAISED  = "#14243D"               # 高亮卡片面
INK_LINE    = "#22304A"               # 描边 / 分隔线
INK_HAIR    = "#1E2B44"               # 卡片内细分隔线
INK_TEXT    = "#F8FAFC"               # 主文字
INK_BODY    = "#CBD5E1"               # 正文
INK_MUTED   = "#7C8CA6"               # 次要文字
INK_ACCENT  = "#38BDF8"               # 主强调色
INK_INDIGO  = "#818CF8"               # 次强调色
INK_CYAN    = "#22D3EE"               # 渐变起点
INK_DEEP    = "#0B1220"               # 强调底上的深色文字

FONT = "Microsoft YaHei, PingFang SC, Noto Sans SC, sans-serif"
MONO = "Consolas, Menlo, monospace"

MARGIN = 80
CONTENT_W = W - MARGIN * 2            # 1120

OUT_DIR = Path(__file__).resolve().parent.parent / "slides"


# ── 基础图元 ──────────────────────────────────────────────────────────────────
def text(x, y, s, size=20, fill=INK_BODY, weight="normal", anchor="start",
         family=FONT, spacing=None):
    a = [f'x="{x}"', f'y="{y}"', f'font-family="{family}"', f'font-size="{size}"',
         f'fill="{fill}"']
    if weight != "normal":
        a.append(f'font-weight="{weight}"')
    if anchor != "start":
        a.append(f'text-anchor="{anchor}"')
    if spacing:
        a.append(f'letter-spacing="{spacing}"')
    return f'<text {" ".join(a)}>{escape(s)}</text>'


def paragraph(x, y, lines, size=18, fill=INK_BODY, lh=None, **kw):
    """一个逻辑段落 = 一个 <text>,后续行用 <tspan x=.. dy=..> 相对换行。"""
    if isinstance(lines, str):
        lines = [lines]
    lh = lh or int(size * 1.55)
    attrs = dict(size=size, fill=fill, **kw)
    body = "".join(f'<tspan x="{x}" dy="{lh}">{escape(line)}</tspan>'
                   for line in lines[1:])
    return text(x, y, lines[0], **attrs).replace("</text>", f"{body}</text>")


def rect(x, y, w, h, fill, stroke=None, rx=None, sw=1):
    a = [f'x="{x}"', f'y="{y}"', f'width="{w}"', f'height="{h}"', f'fill="{fill}"']
    if stroke:
        a += [f'stroke="{stroke}"', f'stroke-width="{sw}"']
    if rx is not None:
        a.append(f'rx="{rx}"')
    return f'<rect {" ".join(a)}/>'


def circle(cx, cy, r, fill, stroke=None, sw=1):
    a = [f'cx="{cx}"', f'cy="{cy}"', f'r="{r}"', f'fill="{fill}"']
    if stroke:
        a += [f'stroke="{stroke}"', f'stroke-width="{sw}"']
    return f'<circle {" ".join(a)}/>'


def hline(x1, x2, y, stroke=INK_LINE, sw=1):
    return (f'<line x1="{x1}" y1="{y}" x2="{x2}" y2="{y}" '
            f'stroke="{stroke}" stroke-width="{sw}"/>')


def card(x, y, w, h, fill=INK_SURFACE, stroke=INK_LINE, rx=18, sw=1):
    return rect(x, y, w, h, fill, stroke, rx, sw)


def group(gid, role, children):
    """逻辑分组。role 仅用于框架类元素(decoration / chrome / background)。"""
    return ("g", gid, role, list(children))


# ── 根级 <g> 的子画布声明(data-pptx-bounds)自动计算 ───────────────────────────
def _text_lines(el):
    """返回 [(文本, 基线 y, x)],首行为直接文本,其余行来自 <tspan dy>。"""
    size = float(el.get("font-size") or 16)
    base_y = float(el.get("y") or 0)
    base_x = float(el.get("x") or 0)
    out, y = [(el.text or "", base_y, base_x)], base_y
    for ts in el:
        if ts.tag != SVG_NS + "tspan":
            continue
        y += float(ts.get("dy") or 0)
        out.append((ts.text or "", y, float(ts.get("x") or base_x)))
    return out, size


def _line_width(s, size, spacing=0.0):
    w = sum(size * (1.0 if ord(ch) > 0x2E7F else 0.62) for ch in s)
    return w + spacing * max(len(s) - 1, 0)


def _measure(elements):
    """粗略量出若干图元的包围盒(用于声明子画布,不参与定位)。"""
    root = ET.fromstring(f'<svg xmlns="http://www.w3.org/2000/svg">'
                         f'{"".join(elements)}</svg>')
    box = [float("inf"), float("inf"), float("-inf"), float("-inf")]

    def add(x1, y1, x2, y2):
        box[0], box[1] = min(box[0], x1), min(box[1], y1)
        box[2], box[3] = max(box[2], x2), max(box[3], y2)

    def walk(el):
        tag = el.tag
        if tag == SVG_NS + "g":
            for child in el:
                walk(child)
        elif tag == SVG_NS + "rect":
            x, y = float(el.get("x") or 0), float(el.get("y") or 0)
            add(x, y, x + float(el.get("width") or 0), y + float(el.get("height") or 0))
        elif tag == SVG_NS + "circle":
            cx, cy = float(el.get("cx") or 0), float(el.get("cy") or 0)
            r = float(el.get("r") or 0)
            add(cx - r, cy - r, cx + r, cy + r)
        elif tag == SVG_NS + "line":
            x1, y1 = float(el.get("x1") or 0), float(el.get("y1") or 0)
            x2, y2 = float(el.get("x2") or 0), float(el.get("y2") or 0)
            add(min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))
        elif tag == SVG_NS + "text":
            lines, size = _text_lines(el)
            spacing = float(el.get("letter-spacing") or 0)
            anchor = el.get("text-anchor") or "start"
            tops = [y - size for _, y, _ in lines]
            bottoms = [y + size * 0.40 for _, y, _ in lines]
            spans = []
            for content, _, x0 in lines:
                w = _line_width(content, size, spacing) * 1.08 + 3
                if anchor == "middle":
                    spans.append((x0 - w / 2, x0 + w / 2))
                elif anchor == "end":
                    spans.append((x0 - w, x0))
                else:
                    spans.append((x0, x0 + w))
            add(min(s[0] for s in spans), min(tops),
                max(s[1] for s in spans), max(bottoms))

    for child in root:
        walk(child)
    return box


def serialize(node, depth=1):
    pad = "  " * depth
    if isinstance(node, str):
        return [pad + line for line in node.split("\n")]
    _, gid, role, children = node
    rendered = []
    for child in children:
        rendered += serialize(child, depth + 1)
    raw = [line.strip() for line in rendered]
    x1, y1, x2, y2 = _measure(raw)
    attr = (f' id="{gid}"'
            + (f' data-pptx-role="{role}"' if role else "")
            + f' data-pptx-bounds="{x1:.0f} {y1:.0f} {x2 - x1:.0f} {y2 - y1:.0f}"')
    out = [f"{pad}<g{attr}>"] + rendered + [f"{pad}</g>"]
    return out


def svg(page_role, nodes):
    head = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" '
        f'viewBox="0 0 {W} {H}" lang="zh-CN" data-pptx-page-role="{page_role}">',
        "  <defs>",
        '    <linearGradient id="deckBg" x1="0" y1="0" x2="0" y2="1">',
        '      <stop offset="0" stop-color="#101B31"/>',
        '      <stop offset="1" stop-color="#0B1220"/>',
        "    </linearGradient>",
        '    <linearGradient id="accent" x1="0" y1="0" x2="1" y2="0">',
        f'      <stop offset="0" stop-color="{INK_CYAN}"/>',
        f'      <stop offset="0.5" stop-color="{INK_ACCENT}"/>',
        f'      <stop offset="1" stop-color="{INK_INDIGO}"/>',
        "    </linearGradient>",
        '    <linearGradient id="accentV" x1="0" y1="0" x2="0" y2="1">',
        f'      <stop offset="0" stop-color="{INK_ACCENT}"/>',
        f'      <stop offset="1" stop-color="{INK_INDIGO}"/>',
        "    </linearGradient>",
        "  </defs>",
    ]
    body = [
        '  <g id="page-frame" data-pptx-role="decoration" '
        f'data-pptx-bounds="0 0 {W} {H}">',
        serialize(rect(0, 0, W, H, "url(#deckBg)"), 2)[0],
        serialize(rect(0, 0, W, 4, "url(#accent)"), 2)[0],
        "  </g>",
    ]
    for node in nodes:
        body += serialize(node, 1)
    return "\n".join(head + body + ["</svg>", ""])


# ── 页眉页脚(静态 chrome) ─────────────────────────────────────────────────────
def chrome(page_no, eyebrow, title):
    return group("page-chrome", "chrome", [
        rect(MARGIN, 76, 44, 4, "url(#accent)", rx=2),
        text(MARGIN, 106, eyebrow, size=14, fill=INK_ACCENT, spacing=3),
        text(MARGIN, 158, title, size=40, fill=INK_TEXT, weight="700"),
        hline(MARGIN, W - MARGIN, 188),
        hline(MARGIN, W - MARGIN, 660, stroke=INK_HAIR),
        text(MARGIN, 684, "PPT Master v6.7.0 × python-pptx · 在 GitHub Codespace 中自动生成",
             size=13, fill=INK_MUTED),
        text(W - MARGIN, 684, f"{page_no:02d} / {PAGES:02d}", size=13,
             fill=INK_MUTED, anchor="end"),
    ])


# ── 01 封面 ───────────────────────────────────────────────────────────────────
def slide_01():
    return "cover", [
        group("cover-rings", "decoration", [
            circle(1040, 540, 80, "none", "#1B2942"),
            circle(1040, 540, 130, "none", "#16223A"),
            circle(1040, 540, 180, "none", "#131F35"),
        ]),
        group("cover-kicker", None, [
            text(MARGIN, 262, "PPT MASTER × PYTHON-PPTX · 自动化演示文稿流水线",
                 size=15, fill=INK_ACCENT, spacing=4),
            rect(MARGIN, 292, 96, 6, "url(#accent)", rx=3),
        ]),
        group("cover-headline", None, [
            text(MARGIN, 372, "用开源工具链", size=58, fill=INK_TEXT, weight="700"),
            text(MARGIN, 442, "生成原生 PowerPoint", size=58, fill=INK_ACCENT, weight="700"),
        ]),
        group("cover-statement", None, [
            paragraph(MARGIN, 508, [
                "8 页 · 16:9 · 全部元素均为可编辑的原生形状与文本框",
                "没有任何一张图片,也没有一个 SmartArt",
            ], size=20, fill=INK_MUTED, lh=34),
            hline(MARGIN, 560, 596, stroke=INK_LINE),
        ]),
        group("cover-meta", None, [
            text(MARGIN, 628, "github.com/plane11875/pptx-deck-builder", size=16,
                 fill=INK_BODY, family=MONO),
            text(W - MARGIN, 628, "2026-10-15", size=16, fill=INK_MUTED, anchor="end"),
        ]),
    ]


# ── 02 这份 PPT 是怎么来的 ────────────────────────────────────────────────────
def slide_02():
    steps = [
        ("01", "初始化项目", "project_manager.py init",
         ["pptx 项目的目录结构、", "画布与执行日志"]),
        ("02", "生成幻灯片源", "builder/make_slides.py",
         ["8 个 SVG 页面,", "就是这份稿子的源文件"]),
        ("03", "质量门禁", "svg_quality_checker.py",
         ["校验 SVG 契约,", "不合规就不让编译"]),
        ("04", "编译成 PPTX", "svg_to_pptx.py",
         ["SVG → DrawingML,", "输出原生可编辑对象"]),
    ]
    els = []
    for i, (num, name, cmd, body) in enumerate(steps):
        x = MARGIN + i * 285
        els += [
            card(x, 236, 255, 300),
            circle(x + 40, 288, 22, "url(#accentV)"),
            text(x + 40, 296, num, size=18, fill=INK_DEEP, weight="700", anchor="middle"),
            text(x + 32, 352, name, size=22, fill=INK_TEXT, weight="700"),
            text(x + 32, 392, cmd, size=13, fill=INK_ACCENT, family=MONO),
            hline(x + 32, x + 223, 418, stroke=INK_HAIR),
            paragraph(x + 32, 452, body, size=17, fill=INK_MUTED, lh=28),
        ]
    return "content", [
        chrome(2, "OVERVIEW", "这份 PPT 是怎么做出来的"),
        group("s2-steps", None, els),
        group("s2-note", None, [
            card(MARGIN, 566, CONTENT_W, 68, fill="#0F1A2E", stroke="#1E3A52"),
            text(MARGIN + 28, 608,
                 "整条链路都在 Codespace 里跑完:不需要模型 API Key,也不需要图片生成接口。",
                 size=18, fill=INK_BODY),
        ]),
    ]


# ── 03 为什么选它 ─────────────────────────────────────────────────────────────
def slide_03():
    items = [
        ("真·原生", ["输出的是 DrawingML 形状", "与文本框,不是导出的图片,",
                     "双击就能接着改。"]),
        ("可复现", ["同一份 SVG 源 → 同一个", "PPTX。版式可以进 Git,",
                    "可以评审、可以回滚。"]),
        ("零密钥", ["纯本地 Python 流水线,", "不依赖任何模型或图片 API,",
                    "离线也能出稿。"]),
    ]
    els = []
    for i, (name, body) in enumerate(items):
        x = MARGIN + i * 385
        els += [
            card(x, 246, 350, 330),
            rect(x + 32, 286, 56, 6, "url(#accent)", rx=3),
            text(x + 32, 348, name, size=27, fill=INK_TEXT, weight="700"),
            paragraph(x + 32, 404, body, size=18, fill=INK_MUTED, lh=32),
        ]
    return "content", [
        chrome(3, "WHY PPT MASTER", "为什么选 PPT Master"),
        group("s3-cards", None, els),
        group("s3-footnote", None, [
            text(MARGIN, 612,
                 "选型依据:GitHub 上 star 数最高的 AI 演示文稿项目,且自带纯 Python 的 SVG→PPTX 编译器。",
                 size=16, fill=INK_MUTED),
        ]),
    ]


# ── 04 技术栈 ─────────────────────────────────────────────────────────────────
def slide_04():
    rows = [
        ("hugohe3/ppt-master", "62,295 ★ · MIT · Python",
         "工作流 + 质检 + SVG→DrawingML 编译器,把矢量元素物化成 PowerPoint 原生形状"),
        ("scanny/python-pptx", "3,556 ★ · MIT · Python",
         "PPTX 读写底座:负责 OOXML 组装、母版与版式绑定、文件落盘"),
        ("skia-pathops · uharfbuzz", "可选依赖",
         "路径布尔运算与文字整形(合并形状、文字轮廓);本次构建未触发该路径"),
    ]
    els = []
    for i, (name, meta, desc) in enumerate(rows):
        y = 232 + i * 124
        els += [
            card(MARGIN, y, CONTENT_W, 104),
            rect(MARGIN + 24, y + 24, 6, 56, "url(#accentV)", rx=3),
            text(MARGIN + 52, y + 48, name, size=22, fill=INK_TEXT, weight="700",
                 family=MONO),
            text(W - MARGIN - 28, y + 48, meta, size=16, fill=INK_ACCENT, anchor="end"),
            text(MARGIN + 52, y + 80, desc, size=17, fill=INK_MUTED),
        ]
    return "content", [
        chrome(4, "TECH STACK", "三层技术栈"),
        group("s4-rows", None, els),
        group("s4-footnote", None, [
            text(MARGIN, 616, "star 数与版本号截至 2026-10-15;三个项目均为 MIT 许可。",
                 size=15, fill=INK_MUTED),
        ]),
    ]


# ── 05 六步流水线 ─────────────────────────────────────────────────────────────
def slide_05():
    steps = [
        ("01", "init", "project_manager.py init"),
        ("02", "写稿", "builder/make_slides.py → slides/"),
        ("03", "质检", "svg_quality_checker.py"),
        ("04", "编译", "svg_to_pptx.py"),
        ("05", "校验", "python-pptx 读回,核对页数与文字"),
        ("06", "交付", "output/ 与 output/preview/"),
    ]
    els = []
    for i, (num, name, cmd) in enumerate(steps):
        x = MARGIN + (i % 3) * 385
        y = 236 + (i // 3) * 196
        els += [
            card(x, y, 350, 168),
            text(x + 28, y + 58, num, size=34, fill=INK_ACCENT, weight="700"),
            text(x + 96, y + 56, name, size=24, fill=INK_TEXT, weight="700"),
            hline(x + 28, x + 322, y + 86, stroke=INK_HAIR),
            text(x + 28, y + 122, cmd, size=15, fill=INK_MUTED, family=MONO),
        ]
    return "content", [
        chrome(5, "PIPELINE", "六步流水线"),
        group("s5-steps", None, els),
    ]


# ── 06 数据看板 ───────────────────────────────────────────────────────────────
def slide_06():
    kpis = [
        ("62,295", "★", "选用的开源项目", "hugohe3/ppt-master"),
        ("3,556", "★", "底层读写库", "scanny/python-pptx"),
        ("8", "页", "本次幻灯片", "1280 × 720 · 16:9"),
        ("0", "个", "API Key 消耗", "全流程本地完成"),
    ]
    els = []
    for i, (num, unit, label, sub) in enumerate(kpis):
        x = MARGIN + i * 285
        els += [
            card(x, 246, 255, 232),
            text(x + 28, 336, num, size=46, fill=INK_TEXT, weight="700"),
            text(x + 28 + len(num) * 26 + 8, 336, unit, size=20, fill=INK_MUTED),
            hline(x + 28, x + 227, 366, stroke=INK_HAIR),
            text(x + 28, 404, label, size=18, fill=INK_BODY),
            text(x + 28, 436, sub, size=15, fill=INK_MUTED),
        ]
    return "content", [
        chrome(6, "BY THE NUMBERS", "这次构建的数据"),
        group("s6-kpis", None, els),
        group("s6-summary", None, [
            card(MARGIN, 514, CONTENT_W, 120, fill="#0F1A2E", stroke="#1E3A52"),
            text(MARGIN + 32, 562, "成品里共有 8 页、约 200 个原生形状与文本框。",
                 size=21, fill=INK_TEXT, weight="700"),
            text(MARGIN + 32, 602,
                 "换句话说:这份稿子的每一根线、每一个圆角矩形,在 PowerPoint 里都能选中、改色、改字。",
                 size=17, fill=INK_MUTED),
        ]),
    ]


# ── 07 路线对比 ───────────────────────────────────────────────────────────────
def slide_07():
    def panel(x, tag, title, pros, cons, hot=False):
        return [
            card(x, 236, 540, 380, fill=INK_RAISED if hot else INK_SURFACE,
                 stroke=INK_ACCENT if hot else INK_LINE),
            rect(x + 32, 236, 96, 4, "url(#accent)", rx=2),
            text(x + 32, 288, tag, size=14, fill=INK_ACCENT if hot else INK_MUTED,
                 spacing=3),
            text(x + 32, 336, title, size=27, fill=INK_TEXT, weight="700"),
            hline(x + 32, x + 508, 364, stroke=INK_HAIR),
            text(x + 32, 404, "得到什么", size=16, fill=INK_ACCENT),
            paragraph(x + 32, 440, pros, size=17, fill=INK_BODY, lh=30),
            text(x + 32, 524, "代价是什么", size=16, fill=INK_MUTED),
            paragraph(x + 32, 556, cons, size=17, fill=INK_MUTED, lh=30),
        ]

    els = panel(MARGIN, "TEMPLATE FILL", "模板填充", [
        "沿用自家母版与品牌规范",
        "AI 只负责填字,风格稳定",
    ], [
        "版式被母版锁死,溢出难控",
        "元素级改动要写 OOXML 补丁",
    ])
    els += panel(MARGIN + 580, "NATIVE COMPILE", "SVG 原生编译", [
        "版式完全自定义,字号行距可控",
        "矢量元素直接物化成原生形状",
    ], [
        "需要先写 SVG(可由脚本生成)",
        "字体要自己保证可用(如思源黑体)",
    ], hot=True)
    return "content", [
        chrome(7, "TRADE-OFF", "两条路线的差别"),
        group("s7-compare", None, els),
        group("s7-note", None, [
            text(MARGIN, 648, "本页采用右栏路线:版式可控、产物可再编辑。",
                 size=15, fill=INK_MUTED),
        ]),
    ]


# ── 08 收尾 ───────────────────────────────────────────────────────────────────
def slide_08():
    code = [
        "git clone https://github.com/hugohe3/ppt-master",
        "pip install python-pptx XlsxWriter skia-pathops uharfbuzz",
        "python builder/make_slides.py",
        "python scripts/svg_quality_checker.py <project>/svg_output --quick-generate",
        "python scripts/svg_to_pptx.py <project> --quick-generate -o output/deck.pptx",
    ]
    return "ending", [
        chrome(8, "RUN IT YOURSELF", "让版式和内容一样可以版本化"),
        group("s8-code", None, [
            card(MARGIN, 228, CONTENT_W, 268, fill="#0A1120", stroke="#1E3A52"),
            rect(MARGIN + 28, 262, 40, 4, "url(#accent)", rx=2),
            text(MARGIN + 28, 296, "bash", size=14, fill=INK_ACCENT, spacing=3,
                 family=MONO),
            paragraph(MARGIN + 28, 342, code, size=16, fill=INK_BODY, lh=32,
                      family=MONO),
        ]),
        group("s8-links", None, [
            text(MARGIN, 546, "仓库:github.com/plane11875/pptx-deck-builder", size=20,
                 fill=INK_TEXT, weight="700"),
            text(MARGIN, 588, "成品在 output/,预览图在 output/preview/,幻灯片源在 slides/。",
                 size=17, fill=INK_MUTED),
            circle(W - MARGIN - 24, 566, 24, "url(#accentV)"),
            text(W - MARGIN - 24, 574, "→", size=22, fill=INK_DEEP, weight="700",
                 anchor="middle"),
        ]),
    ]


# ── 主流程 ────────────────────────────────────────────────────────────────────
SLIDES = [slide_01, slide_02, slide_03, slide_04, slide_05, slide_06, slide_07, slide_08]


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for i, fn in enumerate(SLIDES, start=1):
        page_role, nodes = fn()
        p = OUT_DIR / f"slide_{i:02d}.svg"
        p.write_text(svg(page_role, nodes), encoding="utf-8")
        print(f"wrote {p}")
    print(f"共 {len(SLIDES)} 页 → {OUT_DIR}")


if __name__ == "__main__":
    main()
