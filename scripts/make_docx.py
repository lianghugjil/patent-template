#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""五册 docx 生成器 —— 只用 Python 标准库，不需要 python-docx / Pillow / node。

它是交付管线的 L0 层：沙箱、纯净容器、没有 pip 的机器上，只要有一个 python3，
就能把一份 Markdown 稿子排成符合《专利审查指南》版式的五册申请文件。

用法：
    python make_docx.py 稿子.md                      # 输出 稿子.docx
    python make_docx.py 稿子.md -o 成品.docx
    python make_docx.py 稿子.md --report             # 只体检版式结构，不写文件

Markdown 约定（只要五个一级标题认册名，其余照写）：
    # 说明书摘要
    摘要正文一段，≤300 字。
    # 摘要附图
    ![](fig1.png)
    # 权利要求书
    1. 一种……，其特征在于……
    2. 根据权利要求1所述的……，其特征在于……
    # 说明书
    一种……

    技术领域
    ……
    背景技术
    ……
    发明内容
    ……
    附图说明
    ……
    具体实施方式
    ……
    # 说明书附图
    ![](fig1.png)
    图1 本发明的整体结构示意图

版式（写死在下面的常量里，与 references/docx_format.md 一致）：
    A4 210×297 mm；上边距 2.5、下 2.0、左 2.5、右 2.0 cm；版心 165×252 mm
    正文 宋体 12 pt / 西文 Times New Roman / 固定行距 23 pt / 首行缩进 2 字符
    页眉 黑体 16 pt 居中写册名 + 0.75 pt 下横线；页脚 宋体 12 pt 居中 PAGE 域，每册从 1 起
    附图铺满版心 16.5 cm；摘要附图 10 cm 居中；说明书附图第 2 幅起一页一幅
"""
from __future__ import annotations

import argparse
import datetime as dt
import math
import os
import re
import struct
import sys
import zipfile
from xml.sax.saxutils import escape

# ---------------------------------------------------------------- 版式常量

A4_W, A4_H = 11906, 16838          # twips
MAR_TOP, MAR_BOT, MAR_LEFT, MAR_RIGHT = 1418, 1134, 1418, 1134
HDR_DIST, FTR_DIST = 851, 992
TEXT_W_TWIP = A4_W - MAR_LEFT - MAR_RIGHT      # 9354 twips = 165 mm
CENTER_COMP = 142                              # 左右边距差的一半（0.25 cm），居中补偿
LINE_EXACT = 460                               # 23 pt
SZ_BODY, SZ_HDR, SZ_FTR = 24, 32, 24           # 半磅：12 pt / 16 pt / 12 pt
EMU_PER_TWIP = 635
EMU_FULL_W = TEXT_W_TWIP * EMU_PER_TWIP        # 5940, 约 594 万 EMU = 16.5 cm
EMU_ABS_FIG_W = 3600000                        # 10 cm
FONT_CN, FONT_EN = "宋体", "Times New Roman"
FONT_HDR = "黑体"

BOOKS = ["说明书摘要", "摘要附图", "权利要求书", "说明书", "说明书附图"]
HDR_TEXT = {
    "说明书摘要": "说　明　书　摘　要",
    "摘要附图": "摘　要　附　图",
    "权利要求书": "权　利　要　求　书",
    "说明书": "说　明　书",
    "说明书附图": "说　明　书　附　图",
}
ALIAS = {
    "摘要": "说明书摘要", "说明书摘要": "说明书摘要",
    "摘要附图": "摘要附图", "附图摘要": "摘要附图",
    "权利要求书": "权利要求书", "权利要求": "权利要求书", "权项": "权利要求书",
    "说明书": "说明书", "专利说明书": "说明书",
    "说明书附图": "说明书附图", "附图": "说明书附图",
}
SPEC_HEADINGS = ("技术领域", "背景技术", "发明内容", "附图说明", "具体实施方式")

NS_W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
NS_R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
NS_WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
NS_A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
NS_PIC = 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
NS_M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


# ---------------------------------------------------------------- 小工具

def rpr(size=SZ_BODY, bold=False, cn=FONT_CN, en=FONT_EN):
    s = (f'<w:rPr><w:rFonts w:ascii="{en}" w:hAnsi="{en}" w:eastAsia="{cn}" w:cs="{en}"/>'
         + ('<w:b/><w:bCs/>' if bold else '')
         + f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr>')
    return s


# ---------------------------------------------------------------- 公式（OMML）
# 说明书里的公式如果是纯文本，Word 里不可编辑、审查员无法复算，也不符合五册正式稿的观感。
# 这里用一个够用的子集把 LaTeX 片段转成原生 OMML：分式、根式、上下标、括号、希腊字母、运算符。
# 稿子里独占一行写 $$...$$ 即排成居中公式段。
MATH_GREEK = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
    "zeta": "ζ", "eta": "η", "theta": "θ", "kappa": "κ", "lambda": "λ",
    "mu": "μ", "nu": "ν", "xi": "ξ", "pi": "π", "rho": "ρ", "sigma": "σ",
    "tau": "τ", "phi": "φ", "varphi": "φ", "chi": "χ", "psi": "ψ", "omega": "ω",
    "Delta": "Δ", "Sigma": "Σ", "Omega": "Ω", "Phi": "Φ", "Gamma": "Γ",
    "Lambda": "Λ", "Theta": "Θ", "Psi": "Ψ", "Pi": "Π",
}
MATH_SYMB = {
    "cdot": "·", "times": "×", "div": "÷", "le": "≤", "leq": "≤", "ge": "≥",
    "geq": "≥", "approx": "≈", "ne": "≠", "neq": "≠", "pm": "±", "mp": "∓",
    "to": "→", "infty": "∞", "sum": "∑", "int": "∫", "partial": "∂",
    "degree": "°", "propto": "∝", "ll": "≪", "gg": "≫",
}
MATH_FUNC = ("cos", "sin", "tan", "cot", "exp", "ln", "log", "lg", "max", "min",
             "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh")


def _m_run(text, plain=False):
    """一个公式 run。plain=True 用于中文/符号（正体），否则按变量斜体。"""
    if not text:
        return ""
    sty = '<m:rPr><m:sty m:val="p"/></m:rPr>' if plain else ""
    return f'<m:r>{sty}<m:t xml:space="preserve">{escape(text)}</m:t></m:r>'


def _m_arg(s, i):
    """取一个参数：{...} 取括号内，否则取单个字符或一条命令。返回 (内容, 新位置)。"""
    while i < len(s) and s[i] in " \t":
        i += 1
    if i >= len(s):
        return "", i
    if s[i] == "{":
        depth, j = 1, i + 1
        while j < len(s) and depth:
            if s[j] == "{":
                depth += 1
            elif s[j] == "}":
                depth -= 1
            j += 1
        return s[i + 1:j - 1], j
    if s[i] == "\\":
        j = i + 1
        while j < len(s) and s[j].isalpha():
            j += 1
        return s[i:j], j
    return s[i], i + 1


def _m_atoms(s):
    """解析成一串原子：每个原子带 xml 与可作上下标底数的纯文本。"""
    atoms, i = [], 0
    while i < len(s):
        c = s[i]
        if c in " \t":
            i += 1
            continue
        if c == "}":
            break
        if c == "{":
            inner, i = _m_arg(s, i)
            atoms.append({"xml": _m_atoms_xml(inner), "plain": None})
            continue
        if c == "\\":
            j = i + 1
            while j < len(s) and s[j].isalpha():
                j += 1
            name = s[i + 1:j]
            if name in ("frac", "dfrac", "tfrac"):
                a, i = _m_arg(s, j)
                b, i = _m_arg(s, i)
                atoms.append({"xml": f"<m:f><m:num>{_m_atoms_xml(a)}</m:num>"
                                     f"<m:den>{_m_atoms_xml(b)}</m:den></m:f>", "plain": None})
                continue
            if name == "sqrt":
                a, i = _m_arg(s, j)
                atoms.append({"xml": '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr>'
                                     f'<m:deg/><m:e>{_m_atoms_xml(a)}</m:e></m:rad>',
                              "plain": None})
                continue
            if name in ("left", "right"):
                i = j
                if i < len(s) and s[i] in "()[]|":
                    atoms.append({"xml": _m_run(s[i], True), "plain": s[i]})
                    i += 1
                continue
            if name == "text":
                a, i = _m_arg(s, j)
                atoms.append({"xml": _m_run(a, True), "plain": a})
                continue
            if name in MATH_GREEK:
                atoms.append({"xml": _m_run(MATH_GREEK[name]), "plain": MATH_GREEK[name]})
                i = j
                continue
            if name in MATH_SYMB:
                atoms.append({"xml": _m_run(MATH_SYMB[name], True), "plain": MATH_SYMB[name]})
                i = j
                continue
            if name in MATH_FUNC:          # 函数名用正体，别当变量斜体
                atoms.append({"xml": _m_run(name, True), "plain": name})
                i = j
                continue
            atoms.append({"xml": _m_run(name), "plain": name})   # 未知命令按文字输出
            i = j
            continue
        if c in "_^":
            is_sub = c == "_"
            a, i = _m_arg(s, i + 1)
            base = atoms.pop() if atoms else {"xml": "", "plain": ""}
            b = _m_atoms_xml(a)
            # 同一底数同时带上下标时合并成一个 sSubSup，否则 Word 会拆成两个对象、基线错乱
            if i < len(s) and s[i] in "_^" and s[i] != c:
                a2, i = _m_arg(s, i + 1)
                sub = b if is_sub else _m_atoms_xml(a2)
                sup = _m_atoms_xml(a2) if is_sub else b
                atoms.append({"xml": f"<m:sSubSup><m:e>{base['xml']}</m:e>"
                                     f"<m:sub>{sub}</m:sub><m:sup>{sup}</m:sup></m:sSubSup>",
                              "plain": None})
            else:
                tag = "sub" if is_sub else "sup"
                atoms.append({"xml": f"<m:s{'Sub' if is_sub else 'Sup'}><m:e>{base['xml']}</m:e>"
                                     f"<m:{tag}>{b}</m:{tag}></m:s{'Sub' if is_sub else 'Sup'}>",
                              "plain": None})
            continue
        atoms.append({"xml": _m_run(c), "plain": c})
        i += 1
    return atoms


def _m_atoms_xml(s):
    return "".join(a["xml"] for a in _m_atoms(s))


def math_para(latex):
    """独占一行的居中公式段（原生 OMML，Word 里可编辑）。"""
    ppr = ('<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
           '<w:ind w:firstLine="0"/><w:jc w:val="center"/></w:pPr>')
    return (f'<w:p>{ppr}<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>'
            f'<m:oMath>{_m_atoms_xml(latex)}</m:oMath></m:oMathPara></w:p>')


def para(text, *, size=SZ_BODY, bold=False, jc="both", first_line=True,
         hanging=False, center_page=False, space_before=0, page_break=False):
    """一个正文段落。center_page=True 时用 w:ind 把居中窗口从版心中心挪到纸张中心。"""
    ppr = ['<w:pPr>']
    ppr.append(f'<w:spacing w:before="{space_before}" w:after="0" w:line="{LINE_EXACT}" w:lineRule="exact"/>')
    ind = []
    if first_line:
        ind.append('w:firstLineChars="200" w:firstLine="480"')
    if hanging:
        ind = ['w:left="480" w:hanging="480"']
    if center_page:
        ind = [f'w:left="-{CENTER_COMP}" w:right="{CENTER_COMP}"']
    if ind:
        ppr.append(f'<w:ind {" ".join(ind)}/>')
    ppr.append(f'<w:jc w:val="{jc}"/>')
    ppr.append(rpr(size, bold))
    ppr.append('</w:pPr>')
    runs = []
    if page_break:
        runs.append('<w:r><w:br w:type="page"/></w:r>')
    if text:
        runs.append(f'<w:r>{rpr(size, bold)}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>')
    return f'<w:p>{"".join(ppr)}{"".join(runs)}</w:p>'


def para_image(rid, cx, cy, name, center_page=True, page_break=False):
    ppr = ['<w:pPr>']
    ppr.append(f'<w:spacing w:before="120" w:after="120" w:line="240" w:lineRule="auto"/>')
    if center_page:
        ppr.append(f'<w:ind w:left="-{CENTER_COMP}" w:right="{CENTER_COMP}"/>')
    ppr.append('<w:jc w:val="center"/>')
    ppr.append(rpr())
    ppr.append('</w:pPr>')
    runs = []
    if page_break:
        runs.append('<w:r><w:br w:type="page"/></w:r>')
    runs.append(
        '<w:r><w:drawing>'
        '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{abs(hash(name)) % 60000 + 1}" name="{escape(name)}"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        f'<a:graphic {NS_A}><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic {NS_PIC}><pic:nvPicPr><pic:cNvPr id="0" name="{escape(name)}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')
    return f'<w:p>{"".join(ppr)}{"".join(runs)}</w:p>'


def image_size(path):
    """返回 (px_w, px_h)；只认 PNG / JPEG / GIF，其余返回 None。"""
    try:
        data = open(path, "rb").read(200000)
    except OSError:
        return None
    if data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) >= 24:
        w, h = struct.unpack(">II", data[16:24])
        return int(w), int(h)
    if data[:6] in (b"GIF87a", b"GIF89a"):
        w, h = struct.unpack("<HH", data[6:10])
        return int(w), int(h)
    if data[:2] == b"\xff\xd8":
        i = 2
        while i + 9 < len(data):
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                h, w = struct.unpack(">HH", data[i + 5:i + 9])
                return int(w), int(h)
            seg = struct.unpack(">H", data[i + 2:i + 4])[0]
            i += 2 + seg
    return None


def sized(name, target_emu):
    """按目标宽度算 (cx, cy)，保持原图宽高比。"""
    px = image_size(name)
    if not px or not px[0]:
        return target_emu, int(target_emu * 0.75)
    w, h = px
    return target_emu, int(target_emu * h / w)


# ---------------------------------------------------------------- 解析 md

IMG_RE = re.compile(r"^!\[[^\]]*\]\(([^)]+)\)\s*(?:\{[^}]*\})?\s*$")
CLAIM_RE = re.compile(r"^(\d{1,3})[.．、]\s*")
CAPTION_RE = re.compile(r"^图\s*\d+")


def norm_book_title(raw):
    """把「一、说明书摘要」「(二) 权利要求书：」这类写法归一到册名。"""
    t = raw.strip().strip("*#").strip()
    t = re.sub(r"^[（(]?[一二三四五六七八九十\d]{1,2}[）)]?\s*[、.．:：]?\s*", "", t)
    t = t.strip().rstrip(":：").strip()
    return ALIAS.get(t)


def split_books(text):
    """按标题里的册名切分，返回 [(册名, [行...]), ...]（按官方装订顺序），以及识别到的全部标题。"""
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    found, cur = {}, None
    for ln in lines:
        m = re.match(r"^[ \t]*#{1,3}\s*(.+?)\s*#*\s*$", ln)
        if m:
            name = norm_book_title(m.group(1))
            if name:
                cur = name
                found.setdefault(cur, [])
                continue
        if cur:
            # 不是册名的 # 小标题（技术领域、背景技术……）去掉井号后按普通段落处理，
            # 否则「## 技术领域」会连井号一起写进 docx。
            found[cur].append(m.group(1) if m else ln)
    return [(b, found[b]) for b in BOOKS if b in found], sorted(found.keys())


def usable_figs(lines, base):
    """这一册里真正能嵌入的图片路径（文件存在且是 png/jpg/gif）。"""
    out = []
    for ln in lines:
        m = IMG_RE.match(ln.strip())
        if not m:
            continue
        p = m.group(1).strip()
        if not p.lower().endswith((".png", ".jpg", ".jpeg", ".gif")):
            continue
        if not os.path.isabs(p):
            p = os.path.join(base, p)
        if os.path.exists(p):
            out.append(p)
    return out


def clean_inline(s):
    s = re.sub(r"\*\*(.+?)\*\*", r"\1", s)
    s = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", r"\1", s)
    s = s.replace("`", "")
    s = re.sub(r"\$(.+?)\$", r"\1", s)
    return s.strip()


RE_FIGNO = re.compile(r"^图\s*\d{1,3}$")


def fig_blocks(lines):
    """说明书附图册：把「图N」+ 图片 排成「图在上、图号在下」。

    规范要求图号在图**下方**；而稿子里两种写法都有人用（图号在前 / 图号在后）。
    若图号写在前面，这里就地换位；否则分页符会落在图片上，图号被甩到上一页，
    成品就成了「一页只有图号、下一页才是图」的错位稿。
    """
    out, i = [], 0
    while i < len(lines):
        line = lines[i]
        s = line.strip()
        nxt = lines[i + 1].strip() if i + 1 < len(lines) else ""
        if RE_FIGNO.match(s) and IMG_RE.match(nxt):
            out.append(nxt)          # 先图
            out.append(s)            # 后图号
            i += 2
            continue
        out.append(line)
        i += 1
    return out


def build_book_xml(name, lines, state):
    """把一册的行转成段落 XML。state 里带图片序号与关系 id 计数器。"""
    if name == "说明书附图":
        lines = fig_blocks(lines)
    out = []
    first_para = [True]

    def add_para(txt, **kw):
        out.append(para(txt, **kw))

    idx = 0
    while idx < len(lines):
        raw = lines[idx]
        line = raw.strip()
        idx += 1
        if not line:
            continue

        # 独占一行的 $$...$$ 排成原生 OMML 公式段（Word 可编辑、可复算）
        if len(line) > 4 and line.startswith("$$") and line.endswith("$$"):
            out.append(math_para(line[2:-2].strip()))
            first_para[0] = False
            continue

        m = IMG_RE.match(line)
        if m:
            path = m.group(1).strip()
            if not path.lower().endswith((".png", ".jpg", ".jpeg", ".gif")):
                out.append(para(f"[图未嵌入：{os.path.basename(path)}（仅支持 png/jpg，请先转位图）]",
                                jc="center", first_line=False))
                state["warn"].append(f"{name}：附图 {os.path.basename(path)} 不是 png/jpg，未嵌入")
                continue
            if not os.path.isabs(path):
                path = os.path.join(state["base"], path)
            if not os.path.exists(path):
                out.append(para(f"[缺图：{os.path.basename(m.group(1))}]", jc="center", first_line=False))
                state["warn"].append(f"{name}：找不到附图 {m.group(1)}")
                continue
            rid = f"rIdImg{state['img_n'] + 1}"
            state["img_n"] += 1
            target = EMU_ABS_FIG_W if name == "摘要附图" else EMU_FULL_W
            cx, cy = sized(path, target)
            ext = os.path.splitext(path)[1].lower()
            media = f"image{state['img_n']}{'.jpeg' if ext == '.jpg' else ext}"
            state["media"].append((path, media, rid))
            pb = (name == "说明书附图" and state["fig_in_book"] > 0)
            out.append(para_image(rid, cx, cy, media, page_break=pb))
            if name == "说明书附图":
                state["fig_in_book"] += 1
            first_para[0] = False
            continue

        if line.startswith("|"):
            cells = [c.strip() for c in line.strip("|").split("|")]
            if all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
                continue
            add_para("　".join(c for c in cells if c), first_line=False)
            state["warn"].append(f"{name}：表格已降级为文字行（本版式不做原生表格）")
            continue

        if line.startswith(("---", "***", "___")):
            continue

        txt = clean_inline(line)

        if name == "说明书":
            # 首段居中当作发明名称；但若首段本身就是"技术领域"这类小标题，
            # 说明稿子里没写发明名称，不能把小标题拿去居中当标题
            if first_para[0] and txt not in SPEC_HEADINGS:
                add_para(txt, jc="center", center_page=True, first_line=False)
                first_para[0] = False
                continue
            first_para[0] = False
            if txt in SPEC_HEADINGS:
                add_para(txt, bold=True, first_line=False, space_before=LINE_EXACT // 2)
                continue

        if name == "权利要求书":
            if CLAIM_RE.match(txt):
                # 序号必须留在正文里（审查要按条引用），只做悬挂缩进：首行顶格、续行缩进 2 字
                add_para(txt, hanging=True, first_line=False)
                continue
            if txt.startswith(("权利要求", "如权利要求")):
                add_para(txt, hanging=True, first_line=False)
                continue

        if name == "说明书附图" and CAPTION_RE.match(txt) and len(txt) <= 60:
            add_para(txt, jc="center", first_line=False)
            continue

        if name in ("说明书摘要",):
            for seg in [txt]:
                add_para(seg, first_line=True)
            first_para[0] = False
            continue

        add_para(txt, first_line=True)
        first_para[0] = False
    return "".join(out)


def build_header(text, rid):
    return (XML_HEAD + f'<w:hdr {NS_W} {NS_R}><w:p><w:pPr>'
            '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/></w:pBdr>'
            f'<w:ind w:left="-{CENTER_COMP}" w:right="{CENTER_COMP}"/>'
            '<w:jc w:val="center"/><w:spacing w:line="240" w:lineRule="auto"/>'
            f'{rpr(SZ_HDR, False, FONT_HDR, FONT_HDR)}</w:pPr>'
            f'<w:r>{rpr(SZ_HDR, False, FONT_HDR, FONT_HDR)}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'
            '</w:p></w:hdr>')


def build_footer():
    return (XML_HEAD + f'<w:ftr {NS_W} {NS_R}><w:p><w:pPr>'
            f'<w:ind w:left="-{CENTER_COMP}" w:right="{CENTER_COMP}"/>'
            '<w:jc w:val="center"/><w:spacing w:line="240" w:lineRule="auto"/>'
            f'{rpr(SZ_FTR)}</w:pPr>'
            f'<w:r>{rpr(SZ_FTR)}<w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r>{rpr(SZ_FTR)}<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
            f'<w:r>{rpr(SZ_FTR)}<w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r>{rpr(SZ_FTR)}<w:t>1</w:t></w:r>'
            f'<w:r>{rpr(SZ_FTR)}<w:fldChar w:fldCharType="end"/></w:r>'
            '</w:p></w:ftr>')


def sect_pr(hdr_rid, ftr_rid, last=False):
    return ('<w:sectPr>'
            f'<w:headerReference w:type="default" r:id="{hdr_rid}"/>'
            f'<w:footerReference w:type="default" r:id="{ftr_rid}"/>'
            '<w:type w:val="nextPage"/>'
            f'<w:pgSz w:w="{A4_W}" w:h="{A4_H}"/>'
            f'<w:pgMar w:top="{MAR_TOP}" w:right="{MAR_RIGHT}" w:bottom="{MAR_BOT}" w:left="{MAR_LEFT}" '
            f'w:header="{HDR_DIST}" w:footer="{FTR_DIST}" w:gutter="0"/>'
            '<w:pgNumType w:start="1"/><w:docGrid w:type="lines" w:linePitch="312"/>'
            '</w:sectPr>')


STYLES = XML_HEAD + f'''<w:styles {NS_W}>
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="{FONT_EN}" w:hAnsi="{FONT_EN}" w:eastAsia="{FONT_CN}" w:cs="{FONT_EN}"/>
<w:sz w:val="{SZ_BODY}"/><w:szCs w:val="{SZ_BODY}"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>
<w:spacing w:after="0" w:line="{LINE_EXACT}" w:lineRule="exact"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>
<w:pPr><w:jc w:val="both"/></w:pPr>
<w:rPr><w:rFonts w:ascii="{FONT_EN}" w:hAnsi="{FONT_EN}" w:eastAsia="{FONT_CN}"/><w:sz w:val="{SZ_BODY}"/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/>
<w:pPr><w:jc w:val="center"/><w:spacing w:line="240" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/>
<w:pPr><w:jc w:val="center"/><w:spacing w:line="240" w:lineRule="auto"/></w:pPr></w:style>
</w:styles>'''

SETTINGS = XML_HEAD + f'''<w:settings {NS_W}>
<w:zoom w:percent="100"/><w:defaultTabStop w:val="420"/>
<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
<w:themeFontLang w:val="en-US" w:eastAsia="zh-CN"/>
</w:settings>'''


def read_text(path):
    """读稿子：UTF-8 / UTF-8-BOM / GBK 都能吃，绝不因编码丢字。"""
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def build(md_path, out_path):
    text = read_text(md_path)
    books, all_marks = split_books(text)
    if not books:
        raise SystemExit("[×] 没找到册名标题。请在稿子里用五个一级标题分隔："
                         + "、".join("# " + b for b in BOOKS))
    missing = [b for b in BOOKS if b not in dict(books)]
    if missing:
        print(f"[!] 缺册：{'、'.join(missing)}（照常出稿，但审查会挑）")

    # 附图兜底：一张图都嵌不进去时，宁可不出这两册，也不留两页空白
    base = os.path.dirname(os.path.abspath(md_path))
    notes = []
    figs_ok = usable_figs(dict(books).get("说明书附图", []), base)
    abs_ok = usable_figs(dict(books).get("摘要附图", []), base)
    if not figs_ok:
        drop = [b for b in ("摘要附图", "说明书附图") if b in dict(books)]
        if drop:
            books = [(b, l) for b, l in books if b not in drop]
            notes.append("说明书附图册里没有可嵌入的 png/jpg，已省略「摘要附图」与「说明书附图」两册"
                         "——有附图时这两册必填，无附图时正文也不要再写「图1」")
    elif not abs_ok:
        notes.append("有说明书附图，但摘要附图册缺图：有附图就必须有摘要附图，请补一幅")

    state = {"base": base, "media": [],
             "img_n": 0, "warn": list(notes), "fig_in_book": 0}

    body, rels, hdrs, ftrs = [], [], [], []
    rid_n = [0]

    def new_rid():
        rid_n[0] += 1
        return rid_n[0]

    body.append('<w:body>')
    book_xmls = []
    for n, (name, lines) in enumerate(books):
        state["fig_in_book"] = 0
        book_xml = build_book_xml(name, lines, state)
        book_xmls.append(book_xml)
        hrid = f"rIdH{n + 1}"
        frid = f"rIdF{n + 1}"
        rels.append((hrid, "header", f"header{n + 1}.xml"))
        rels.append((frid, "footer", f"footer{n + 1}.xml"))
        hdrs.append((f"header{n + 1}.xml", HDR_TEXT[name]))
        ftrs.append(f"footer{n + 1}.xml")
        if n + 1 < len(books):
            # 分节符挂在本书最后一个段落的 pPr 里，不额外加空段（否则册尾多一页空白）
            i = book_xml.rfind("</w:pPr>")
            if i > 0:
                book_xml = book_xml[:i] + sect_pr(hrid, frid) + book_xml[i:]
            else:
                book_xml += f'<w:p><w:pPr>{sect_pr(hrid, frid)}</w:pPr></w:p>'
        body.append(book_xml)
    last_h, last_f = f"rIdH{len(books)}", f"rIdF{len(books)}"
    body.append(sect_pr(last_h, last_f, last=True))
    body.append('</w:body>')

    doc = XML_HEAD + f'<w:document {NS_W} {NS_R} {NS_WP} {NS_A} {NS_PIC} {NS_M}>' + "".join(body) + '</w:document>'

    rels_xml = (XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>')
    for rid, typ, target in rels:
        rels_xml += (f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{typ}" '
                     f'Target="{target}"/>')
    for path, media, rid in state["media"]:
        rels_xml += (f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                     f'Target="media/{media}"/>')
    rels_xml += '</Relationships>'

    exts = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif"}
    used_ext = sorted({os.path.splitext(m)[1].lower() for _, m, _ in state["media"]})
    ctypes = (XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
              '<Default Extension="xml" ContentType="application/xml"/>')
    for e in used_ext or [".png"]:
        ctypes += f'<Default Extension="{e.lstrip(".")}" ContentType="{exts.get(e, "image/png")}"/>'
    ctypes += ('<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
               '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
               '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>')
    for i, _ in enumerate(hdrs, 1):
        ctypes += (f'<Override PartName="/word/header{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>')
    for i, _ in enumerate(ftrs, 1):
        ctypes += (f'<Override PartName="/word/footer{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>')
    ctypes += '</Types>'

    pkg = (XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
           '</Relationships>')

    if os.path.exists(out_path):
        os.remove(out_path)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ctypes)
        z.writestr("_rels/.rels", pkg)
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/_rels/document.xml.rels", rels_xml)
        for fname, htext in hdrs:
            z.writestr(f"word/{fname}", build_header(htext, ""))
        for fname in ftrs:
            z.writestr(f"word/{fname}", build_footer())
        for path, media, _ in state["media"]:
            z.write(path, f"word/media/{media}")
    return books, state, doc, book_xmls


def book_pages(xml):
    """按版心 39 字/行、30.8 行/页粗估一册的页数（每册单独起页，各自向上取整）。

    逐段累计而不是全册字数一除：每段末尾的半行不能并到下一段，否则短段多的稿子会少算好几页。
    """
    chars = 0
    lines = 0
    for pm in re.findall(r"<w:p>.*?</w:p>", xml, re.S):
        txt = "".join(re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", pm, re.S))
        chars += len(txt)
        if "<w:drawing>" in pm:
            lines += 15                     # 铺满版心的图约 11 cm ≈ 14 行，加图注 1 行
        elif txt.strip():
            lines += max(1, math.ceil(len(txt) / 39))
    imgs = xml.count("<w:drawing>")
    return max(1, math.ceil(lines / 30.8)), chars, imgs


def estimate_pages(book_xmls):
    pages = chars = imgs = 0
    for xml in book_xmls:
        p, c, i = book_pages(xml)
        pages += p
        chars += c
        imgs += i
    return pages, chars, imgs


def main():
    ap = argparse.ArgumentParser(description="零依赖五册 docx 生成器")
    ap.add_argument("md")
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("--report", action="store_true", help="只体检结构，不写文件")
    args = ap.parse_args()

    md = args.md
    if not os.path.exists(md):
        raise SystemExit(f"[×] 文件不存在：{md}")
    out = args.out or (os.path.splitext(md)[0] + ".docx")
    _d = os.path.dirname(os.path.abspath(out))
    if _d and not os.path.isdir(_d):
        os.makedirs(_d, exist_ok=True)

    if args.report:
        text = read_text(md)
        books, extra = split_books(text)
        print(f"册数：{len(books)}/{len(BOOKS)}  →  {'、'.join(b for b, _ in books) or '无'}")
        for b in BOOKS:
            if b not in dict(books):
                print(f"  [E] 缺册：{b}")
        for b, lines in books:
            body = "".join(lines)
            print(f"  - {b}: {len(re.sub(chr(92)+'s', '', body))} 字符，"
                  f"{len([l for l in lines if IMG_RE.match(l.strip())])} 图")
        return 0

    books, state, doc, book_xmls = build(md, out)
    pages, chars, imgs = estimate_pages(book_xmls)
    size_kb = os.path.getsize(out) / 1024
    print(f"[✓] {out}")
    print(f"    册数 {len(books)}  （{'、'.join(b for b, _ in books)}）")
    print(f"    正文约 {chars} 字，嵌入图 {imgs} 幅，估约 {pages} 页，{size_kb:.0f} KB")
    print(f"    版式：A4 210×297，边距 2.5/2.0/2.5/2.0，正文宋体 12 pt，固定行距 23 pt，"
          f"页眉册名 + 0.75 pt 横线，每册页码从 1 起")
    for w in dict.fromkeys(state["warn"]):
        print(f"    [!] {w}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
