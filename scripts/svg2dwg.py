# -*- coding: utf-8 -*-
"""svg2dwg.py —— 专利附图 SVG 批量转 DXF + DWG（CAD 原生可编辑）

为什么要它：专利说明书正文里的附图是 SVG（或 PNG 内嵌），审查/答复/改型时用户往往要在
CAD 里改结构、挪标号、加件。SVG 进不了 CAD，所以每张附图必须同时产出 DXF（保底）与
DWG（交付）。

一条命令搞定（默认就出 DXF + DWG，不用加参数）：

    python svg2dwg.py 附图/*.svg --out 说明书附图
    python svg2dwg.py fig1.svg --out out --oda "D:/tools/ODA/ODAFileConverter.exe"

支持的元素：line / rect / polygon / polyline / path(M L H V C Q S T A Z) /
circle / ellipse / text，以及 g 上的 translate / scale / rotate / matrix 变换。
**同时展开 SVG 专有的两种表达**（CAD 不认，不展开就会静默丢失）：
  - `<pattern>` 剖面线填充（`fill="url(#h1)"`）→ 按 pattern 角度/间距用扫描线求交生成
    真实平行线段，写进 HATCH 图层；原图形只留轮廓，**不会变成实心黑块**；
  - `<marker>` 引线箭头（`marker-end="url(#ar)"`）→ 按末段方向生成实心三角（SOLID）。
线宽→图层按 `references/drawing_spec.md` 的约定自动映射，不需要人工分层。

坐标：SVG y 向下、DWG y 向上，自动翻转（画布高从 viewBox / height 读，读不到用 600）。
单位：1 SVG 单位 = 1 mm（`--scale 0.35` 可整体缩到 A4 版心）。

依赖：ezdxf；DWG 需要 ODA File Converter（见 references/dwg_export.md）。
"""
from __future__ import annotations

import argparse
import glob
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from collections import Counter

import ezdxf
from ezdxf.tools.standards import setup_linetypes

NS = "{http://www.w3.org/2000/svg}"

# 线宽 → (图层, 线宽 mm×100)。与 drawing_spec.md §1 的 SVG 约定一一对应。
# 2.2 轮廓 / 1.2 引线 / 0.9 剖面线 / 0.8 中心线（带 dasharray 时强制 CENTER）
LAYER_RULES = [
    (1.80, "OUTLINE", 35),   # 粗实线：轮廓
    (1.05, "LEADER", 18),    # 引线
    (0.85, "HATCH", 13),     # 剖面线
    (0.00, "CENTER", 13),    # 中心线 / 细线
]
LAYER_ORDER = ["OUTLINE", "HATCH", "CENTER", "LEADER", "LABEL", "CURVE", "AXIS", "GRID"]

# 曲线图（数据图）常用别的宽度，兜底按“是否闭合”分到 CURVE / AXIS
CJK = re.compile(r"[\u4e00-\u9fff]")


# --------------------------------------------------------------------------- #
# 变换矩阵（SVG 的 translate/scale/rotate/matrix/skew）
# --------------------------------------------------------------------------- #
def mat_mul(a, b):
    a0, a1, a2, a3, a4, a5 = a
    b0, b1, b2, b3, b4, b5 = b
    return (
        a0 * b0 + a2 * b1, a1 * b0 + a3 * b1,
        a0 * b2 + a2 * b3, a1 * b2 + a3 * b3,
        a0 * b4 + a2 * b5 + a4, a1 * b4 + a3 * b5 + a5,
    )


IDENT = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)


def parse_transform(s):
    m = IDENT
    if not s:
        return m
    for name, arg in re.findall(r"(translate|scale|rotate|matrix|skewX|skewY)\s*\(([^)]*)\)", s):
        v = [float(x) for x in re.findall(r"-?\d*\.?\d+(?:[eE][-+]?\d+)?", arg)]
        if name == "translate":
            t = (1, 0, 0, 1, v[0], v[1] if len(v) > 1 else 0.0)
        elif name == "scale":
            sx = v[0]
            sy = v[1] if len(v) > 1 else sx
            t = (sx, 0, 0, sy, 0, 0)
        elif name == "rotate":
            a = math.radians(v[0])
            ca, sa = math.cos(a), math.sin(a)
            t = (ca, sa, -sa, ca, 0, 0)
            if len(v) >= 3:
                t = mat_mul(mat_mul((1, 0, 0, 1, v[1], v[2]), t), (1, 0, 0, 1, -v[1], -v[2]))
        elif name == "matrix":
            t = tuple(v[:6])
        elif name == "skewX":
            t = (1, 0, math.tan(math.radians(v[0])), 1, 0, 0)
        else:
            t = (1, math.tan(math.radians(v[0])), 0, 1, 0, 0)
        m = mat_mul(m, t)
    return m


def apply(m, x, y):
    return (m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5])


# --------------------------------------------------------------------------- #
# 属性解析（含 style="stroke-width:2.2;..." 的内联写法）
# --------------------------------------------------------------------------- #
def prop(el, key, default=None):
    d = {}
    for part in (el.get("style") or "").split(";"):
        if ":" in part:
            k, v = part.split(":", 1)
            d[k.strip()] = v.strip()
    if key in d:
        return d[key]
    return el.get(key, default)


def fnum(v, default=0.0):
    if v is None:
        return default
    m = re.search(r"-?\d*\.?\d+(?:[eE][-+]?\d+)?", str(v))
    return float(m.group()) if m else default


def pick_layer(width, dashed, closed=False):
    if dashed:
        return "CENTER" if width >= 0.5 else "GRID"
    for lo, name, lw in LAYER_RULES:
        if width >= lo:
            return name
    return "CURVE" if not closed else "OUTLINE"


def new_doc():
    doc = ezdxf.new("R2013", setup=True)
    setup_linetypes(doc)
    for name in LAYER_ORDER:
        if name in doc.layers:
            continue
        lw = {"OUTLINE": 35, "HATCH": 13, "CENTER": 13, "LEADER": 18,
              "LABEL": 18, "CURVE": 30, "AXIS": 25, "GRID": 13}[name]
        lt = {"CENTER": "CENTER", "GRID": "DASHED"}.get(name, "CONTINUOUS")
        doc.layers.add(name, color=7, linetype=lt, lineweight=lw)
    # 国标字体：gbenor.shx（西文/数字）+ gbcbig.shx（中文大字体）
    doc.styles.add("GB", font="gbenor.shx", dxfattribs={"bigfont": "gbcbig.shx"})
    return doc


# --------------------------------------------------------------------------- #
# path d 解析（含贝塞尔展开）
# --------------------------------------------------------------------------- #
TOKEN = re.compile(r"[MmLlHhVvCcSsQqTtAaZz]|-?\d*\.?\d+(?:[eE][-+]?\d+)?")


def hatch_segments(poly, angle_deg=45.0, spacing=9.0):
    """在多边形内生成 angle_deg 方向的平行剖面线（扫描线求交 + 奇偶配对）。

    用于展开 SVG 的 <pattern> 填充（fill="url(#h1)"）——CAD 不认 pattern，不展开就
    只剩轮廓，剖面线全丢（更糟的是某些转换器会把整块画成实心黑块）。
    """
    if len(poly) < 3:
        return []
    a = math.radians(angle_deg)
    ca, sa = math.cos(-a), math.sin(-a)          # 旋到水平，扫完再转回
    pts = [(x * ca - y * sa, x * sa + y * ca) for x, y in poly]
    ys = [p[1] for p in pts]
    y = math.floor(min(ys) / spacing) * spacing
    segs = []
    n = len(pts)
    while y <= max(ys):
        xs = []
        for i in range(n):
            x1, y1 = pts[i]
            x2, y2 = pts[(i + 1) % n]
            if (y1 <= y < y2) or (y2 <= y < y1):
                t = (y - y1) / (y2 - y1)
                xs.append(x1 + t * (x2 - x1))
        xs.sort()
        for i in range(0, len(xs) - 1, 2):
            if xs[i + 1] - xs[i] > 1e-6:
                segs.append(((xs[i], y), (xs[i + 1], y)))
        y += spacing
    cb, sb = math.cos(a), math.sin(a)
    return [((p1[0] * cb - p1[1] * sb, p1[0] * sb + p1[1] * cb),
             (p2[0] * cb - p2[1] * sb, p2[0] * sb + p2[1] * cb)) for p1, p2 in segs]


def arrow_triangle(p_end, p_prev, length=11.0, half=4.5):
    """引线末端实心三角：尖端在 p_end（DWG 坐标）。"""
    dx, dy = p_end[0] - p_prev[0], p_end[1] - p_prev[1]
    d = math.hypot(dx, dy)
    if d < 1e-9:
        return None
    dx, dy = dx / d, dy / d
    bx, by = p_end[0] - dx * length, p_end[1] - dy * length
    nx, ny = -dy, dx
    return [p_end, (bx + nx * half, by + ny * half), (bx - nx * half, by - ny * half)]


def path_polylines(d, steps=18):
    """返回 [(points, closed), ...]，坐标仍是 SVG 原始坐标。"""
    toks = TOKEN.findall(d or "")
    out, pts, closed = [], [], False
    i, cmd = 0, None
    cur = (0.0, 0.0)
    start = (0.0, 0.0)
    prev_ctrl = None

    def flush():
        nonlocal pts
        if len(pts) > 1:
            out.append((pts, closed))
        pts = [pts[-1]] if pts else []

    def num():
        nonlocal i
        v = float(toks[i]); i += 1
        return v

    while i < len(toks):
        if re.match(r"[A-Za-z]", toks[i]):
            cmd = toks[i]; i += 1
            if cmd in "Zz":
                closed = True
                flush()
                cur = start
                continue
        if cmd is None:
            break
        rel = cmd.islower()
        c = cmd.upper()
        try:
            if c == "M":
                x, y = num(), num()
                if rel:
                    x, y = x + cur[0], y + cur[1]
                flush()
                closed = False
                pts = [(x, y)]
                cur = start = (x, y)
                cmd = "l" if rel else "L"          # 后续隐式 lineto
            elif c == "L":
                x, y = num(), num()
                if rel:
                    x, y = x + cur[0], y + cur[1]
                pts.append((x, y)); cur = (x, y)
            elif c == "H":
                x = num()
                x = x + cur[0] if rel else x
                pts.append((x, cur[1])); cur = (x, cur[1])
            elif c == "V":
                y = num()
                y = y + cur[1] if rel else y
                pts.append((cur[0], y)); cur = (cur[0], y)
            elif c in "CQST":
                if c == "C":
                    p1 = (num() + (cur[0] if rel else 0), num() + (cur[1] if rel else 0))
                    p2 = (num() + (cur[0] if rel else 0), num() + (cur[1] if rel else 0))
                    p3 = (num() + (cur[0] if rel else 0), num() + (cur[1] if rel else 0))
                elif c == "S":
                    p1 = (2 * cur[0] - prev_ctrl[0], 2 * cur[1] - prev_ctrl[1]) if prev_ctrl else cur
                    p2 = (num() + (cur[0] if rel else 0), num() + (cur[1] if rel else 0))
                    p3 = (num() + (cur[0] if rel else 0), num() + (cur[1] if rel else 0))
                elif c == "Q":
                    p1 = (num() + (cur[0] if rel else 0), num() + (cur[1] if rel else 0))
                    p3 = (num() + (cur[0] if rel else 0), num() + (cur[1] if rel else 0))
                    p2 = p1
                else:  # T
                    p1 = (2 * cur[0] - prev_ctrl[0], 2 * cur[1] - prev_ctrl[1]) if prev_ctrl else cur
                    p2 = p1
                    p3 = (num() + (cur[0] if rel else 0), num() + (cur[1] if rel else 0))
                for k in range(1, steps + 1):
                    t = k / steps
                    if c in "CST":
                        mt = 1 - t
                        x = (mt ** 3 * cur[0] + 3 * mt * mt * t * p1[0]
                             + 3 * mt * t * t * p2[0] + t ** 3 * p3[0])
                        y = (mt ** 3 * cur[1] + 3 * mt * mt * t * p1[1]
                             + 3 * mt * t * t * p2[1] + t ** 3 * p3[1])
                    else:
                        mt = 1 - t
                        x = mt * mt * cur[0] + 2 * mt * t * p1[0] + t * t * p3[0]
                        y = mt * mt * cur[1] + 2 * mt * t * p1[1] + t * t * p3[1]
                    pts.append((x, y))
                prev_ctrl = p2
                cur = p3
            elif c == "A":
                num(); num(); num(); num(); num()          # rx ry rot large sweep
                x, y = num(), num()
                if rel:
                    x, y = x + cur[0], y + cur[1]
                pts.append((x, y)); cur = (x, y)           # 圆弧按直线近似（附图里极少用）
            else:
                i += 1
        except IndexError:
            break
    flush()
    return out


# --------------------------------------------------------------------------- #
# 单张转换
# --------------------------------------------------------------------------- #
def convert_one(svg_path, dxf_path, scale=1.0, default_h=600.0, hatch_layers=None):
    root = ET.parse(svg_path).getroot()

    H = None
    vb = root.get("viewBox")
    if vb:
        v = [float(x) for x in re.split(r"[,\s]+", vb.strip()) if x]
        if len(v) == 4:
            H = v[3]
    if H is None:
        H = fnum(root.get("height"), default_h)
    H = H or default_h

    def T(p, m):
        """SVG 点 -> DWG 点（缩放 + y 翻转）"""
        x, y = apply(m, p[0], p[1])
        return (x * scale, (H - y) * scale)

    doc = new_doc()
    msp = doc.modelspace()
    stat = {"LINE": 0, "LWPOLYLINE": 0, "TEXT": 0, "SOLID": 0, "CIRCLE": 0, "ELLIPSE": 0,
            "HATCH": 0, "ARROW": 0}

    # ---- <pattern>（剖面线）与 <marker>（箭头）参数：CAD 不认，必须在几何层展开 ----
    patterns, markers = {}, {}
    for d in root.iter(NS + "defs"):
        for p in d.iter(NS + "pattern"):
            ang, sp = 45.0, 9.0
            mt = re.search(r"rotate\(\s*(-?[\d.]+)", p.get("patternTransform") or "")
            if mt:
                ang = float(mt.group(1))
            sp = fnum(p.get("width"), 9.0) or 9.0
            patterns[p.get("id")] = (ang, sp)
        for mk in d.iter(NS + "marker"):
            markers[mk.get("id")] = (fnum(mk.get("markerWidth"), 11.0),
                                     fnum(mk.get("markerHeight"), 9.0))

    def pattern_of(el):
        f = (prop(el, "fill") or "").strip()
        m = re.match(r"url\(#([^)]+)\)", f)
        return m.group(1) if m else None

    def do_hatch(poly_pts, pid, m):
        ang, sp = patterns.get(pid, (45.0, 9.0))
        for a, b in hatch_segments(poly_pts, ang, sp):
            msp.add_line(T(a, m), T(b, m), dxfattribs={"layer": "HATCH"})
            stat["HATCH"] += 1

    def _arrow(el, pts):
        mk = (el.get("marker-end") or prop(el, "marker-end") or "").strip()
        m = re.match(r"url\(#([^)]+)\)", mk)
        if not m or len(pts) < 2:
            return
        mw, mh = markers.get(m.group(1), (11.0, 9.0))
        tri = arrow_triangle(pts[-1], pts[-2], length=mw * scale, half=mh * scale / 2.0)
        if tri:
            msp.add_solid(tri, dxfattribs={"layer": "LEADER"})
            stat["ARROW"] += 1

    def walk(el, m):
        tag = el.tag.replace(NS, "")
        mm = mat_mul(m, parse_transform(el.get("transform")))
        if tag in ("g", "svg", "a"):
            for ch in el:
                walk(ch, mm)
            return

        w = fnum(prop(el, "stroke-width"), 1.0)
        dashed = bool(prop(el, "stroke-dasharray")) and prop(el, "stroke-dasharray") != "none"
        layer = pick_layer(w, dashed)

        if tag == "line":
            p1 = T((fnum(el.get("x1")), fnum(el.get("y1"))), mm)
            p2 = T((fnum(el.get("x2")), fnum(el.get("y2"))), mm)
            msp.add_line(p1, p2, dxfattribs={"layer": layer})
            stat["LINE"] += 1
            _arrow(el, [p1, p2])

        elif tag == "rect":
            # 背景白底（无 x/width 或满幅）跳过，否则在 CAD 里变成一个大框
            if el.get("x") is None or el.get("width") is None:
                return
            x0, y0 = fnum(el.get("x")), fnum(el.get("y"))
            x1, y1 = x0 + fnum(el.get("width")), y0 + fnum(el.get("height"))
            if x0 <= 0.01 and y0 <= 0.01 and (x1 >= H or y1 >= H):
                return
            pts = [T(p, mm) for p in ((x0, y0), (x1, y0), (x1, y1), (x0, y1))]
            msp.add_lwpolyline(pts, close=True, dxfattribs={"layer": "OUTLINE" if w >= 0.5 else layer})
            stat["LWPOLYLINE"] += 1
            pid = pattern_of(el)
            if pid:                                   # 剖面填充 → 展开成 45° 平行线
                do_hatch([(x0, y0), (x1, y0), (x1, y1), (x0, y1)], pid, mm)

        elif tag in ("polygon", "polyline"):
            raw = (el.get("points") or "").split()
            pts0 = []
            for pair in raw:
                a, b = re.split(r"[,\s]+", pair.strip()) if re.search(r"[,\s]", pair) else (None, None)
                if a is None:
                    continue
                pts0.append((float(a), float(b)))
            if len(pts0) < 2:
                return
            pts = [T(p, mm) for p in pts0]
            fill = (prop(el, "fill") or "").lower()
            pid = pattern_of(el)
            if tag == "polygon" and pid:              # 剖面填充：轮廓 + 剖面线，绝不画成实心
                msp.add_lwpolyline(pts, close=True, dxfattribs={"layer": layer})
                stat["LWPOLYLINE"] += 1
                do_hatch(pts0, pid, mm)
            elif tag == "polygon" and len(pts) <= 4 and fill not in ("none", "white", "#fff", "#ffffff"):
                msp.add_solid(pts[:3] if len(pts) == 3 else pts, dxfattribs={"layer": "LEADER"})
                stat["SOLID"] += 1
            else:
                msp.add_lwpolyline(pts, close=(tag == "polygon"),
                                   dxfattribs={"layer": layer if tag == "polyline" else
                                               ("OUTLINE" if w >= 0.5 else layer)})
                stat["LWPOLYLINE"] += 1
            _arrow(el, [T(p, mm) for p in pts0])

        elif tag == "path":
            subs = path_polylines(el.get("d"))
            pid = pattern_of(el)
            for k, (pts0, closed) in enumerate(subs):
                pts = [T(p, mm) for p in pts0]
                if len(pts) < 2:
                    continue
                msp.add_lwpolyline(pts, close=closed, dxfattribs={"layer": layer})
                stat["LWPOLYLINE"] += 1
                if pid and k == 0:
                    do_hatch(pts0, pid, mm)
            if subs:
                _arrow(el, [T(p, mm) for p in subs[-1][0]])

        elif tag == "circle":
            c = T((fnum(el.get("cx")), fnum(el.get("cy"))), mm)
            r = fnum(el.get("r")) * scale
            if r > 0:
                msp.add_circle(c, r, dxfattribs={"layer": layer})
                stat["CIRCLE"] += 1

        elif tag == "ellipse":
            c = T((fnum(el.get("cx")), fnum(el.get("cy"))), mm)
            rx = fnum(el.get("rx")) * scale
            ry = fnum(el.get("ry")) * scale
            if rx > 0 and ry > 0:
                msp.add_ellipse(c, major_axis=(rx, 0), ratio=ry / rx,
                                dxfattribs={"layer": layer})
                stat["ELLIPSE"] += 1

        elif tag == "text":
            s = "".join(el.itertext()).strip()
            if not s:
                return
            fs = fnum(prop(el, "font-size"), 22.0) * scale
            tx, ty = fnum(el.get("x")), fnum(el.get("y"))
            anchor = prop(el, "text-anchor", "start")
            rot = 0.0
            mt = re.search(r"rotate\(\s*(-?[\d.]+)", el.get("transform", "") or "")
            if mt:
                rot = -float(mt.group(1))          # SVG 顺时针 -> CAD 逆时针
            pos = T((tx, ty), mm)
            t = msp.add_text(s, height=fs,
                             dxfattribs={"layer": "LABEL", "style": "GB",
                                         "rotation": rot, "insert": pos})
            if anchor in ("middle", "end"):
                t.dxf.halign = 1 if anchor == "middle" else 2
                t.dxf.valign = 0                    # SVG 的 y 是基线，不是垂直中心
                t.dxf.align_point = (pos[0], pos[1], 0.0)
            stat["TEXT"] += 1

    walk(root, IDENT)
    doc.saveas(dxf_path)
    return stat


# --------------------------------------------------------------------------- #
# DWG：找 ODA 转换器
# --------------------------------------------------------------------------- #
def find_oda(explicit=None):
    """按 显式参数 → 环境变量 → 项目内 _tools/ → 系统安装目录 → PATH 的顺序找。"""
    cands = []
    if explicit:
        cands.append(explicit)
    if os.environ.get("ODA_CONVERTER"):
        cands.append(os.environ["ODA_CONVERTER"])
    # 项目内便携部署（免安装解包出来的那套）
    here = os.path.abspath(os.getcwd())
    for _ in range(3):
        cands.append(os.path.join(here, "_tools", "ODA", "ODAFileConverter.exe"))
        here = os.path.dirname(here)
    cands += [
        r"C:\Program Files\ODA\ODAFileConverter\ODAFileConverter.exe",
        r"C:\Program Files (x86)\ODA\ODAFileConverter\ODAFileConverter.exe",
        r"C:\Program Files\ODA\ODAFileConverter 25.4\ODAFileConverter.exe",
        "/usr/bin/ODAFileConverter",
        "/opt/ODA/ODAFileConverter/ODAFileConverter",
        "/Applications/ODAFileConverter.app/Contents/MacOS/ODAFileConverter",
    ]
    for c in cands:
        if c and os.path.isfile(c):
            return c
    return shutil.which("ODAFileConverter")


def dxf_to_dwg(oda, dxf_paths, out_dir, version="ACAD2018"):
    """ODA 只吃目录，这里用临时目录中转。返回 {dxf_basename: dwg_path}"""
    with tempfile.TemporaryDirectory() as tin, tempfile.TemporaryDirectory() as tout:
        for p in dxf_paths:
            shutil.copy2(p, os.path.join(tin, os.path.basename(p)))
        r = subprocess.run(
            [oda, tin, tout, version, "DWG", "0", "1", "*.dxf"],
            capture_output=True, text=True, timeout=600)
        out = {}
        for p in dxf_paths:
            b = os.path.basename(p)
            got = os.path.join(tout, os.path.splitext(b)[0] + ".dwg")
            if os.path.isfile(got) and os.path.getsize(got) > 0:
                dst = os.path.join(out_dir, os.path.splitext(b)[0] + ".dwg")
                shutil.move(got, dst)
                out[b] = dst
            else:
                sys.stderr.write(f"!! DWG 转换失败: {b}\n{r.stdout[-400:]}{r.stderr[-400:]}\n")
        return out


def verify_dwg(oda, dwg_path, dxf_path):
    """把 DWG 反向转回 DXF，比对 实体分类计数 / 文字多重集 / 图层表。返回 (ok, detail)"""
    with tempfile.TemporaryDirectory() as tin, tempfile.TemporaryDirectory() as tout:
        shutil.copy2(dwg_path, os.path.join(tin, os.path.basename(dwg_path)))
        subprocess.run([oda, tin, tout, "ACAD2018", "DXF", "0", "0", "*.dwg"],
                       capture_output=True, text=True, timeout=600)
        back = os.path.join(tout, os.path.splitext(os.path.basename(dwg_path))[0] + ".dxf")
        if not os.path.isfile(back):
            return False, "回读失败（ODA 未产出 DXF）"
        a = ezdxf.readfile(dxf_path)
        b = ezdxf.readfile(back)
        ca = Counter(e.dxftype() for e in a.modelspace())
        cb = Counter(e.dxftype() for e in b.modelspace())
        ta = sorted(e.dxf.text for e in a.modelspace() if e.dxftype() == "TEXT")
        tb = sorted(e.dxf.text for e in b.modelspace() if e.dxftype() == "TEXT")
        la = sorted(l.dxf.name for l in a.layers)
        lb = sorted(l.dxf.name for l in b.layers)
        ok = (ca == cb) and (ta == tb) and (la == lb)
        detail = (f"实体 {sum(ca.values())}/{sum(cb.values())}  "
                  f"文字 {len(ta)}/{len(tb)}  图层 {len(la)}/{len(lb)}"
                  f"{'' if la == lb else ' 图层不一致'}")
        if ca != cb:
            detail += f"  分类差异 {dict(ca - cb)} / {dict(cb - ca)}"
        if ta != tb:
            detail += f"  漏字 {sorted(set(ta) - set(tb))[:6]}"
        return ok, detail


# --------------------------------------------------------------------------- #
def main(argv=None):
    ap = argparse.ArgumentParser(description="专利附图 SVG → DXF + DWG")
    ap.add_argument("inputs", nargs="+", help="SVG 文件或目录（目录取 *.svg）")
    ap.add_argument("--out", default=None, help="输出目录，默认与 SVG 同目录")
    ap.add_argument("--scale", type=float, default=1.0, help="SVG 单位→mm 比例，默认 1:1")
    ap.add_argument("--height", type=float, default=600.0, help="读不到画布高时的兜底值")
    ap.add_argument("--oda", default=None, help="ODAFileConverter.exe 路径")
    ap.add_argument("--version", default="ACAD2018", help="DWG 版本，默认 ACAD2018")
    ap.add_argument("--dxf-only", action="store_true", help="只出 DXF，不转 DWG")
    ap.add_argument("--no-verify", action="store_true", help="跳过 DWG 回读校验")
    args = ap.parse_args(argv)

    svgs = []
    for p in args.inputs:
        if os.path.isdir(p):
            svgs += sorted(glob.glob(os.path.join(p, "*.svg")))
        elif os.path.isfile(p):
            svgs.append(p)
        else:
            sys.stderr.write(f"跳过（不存在）: {p}\n")
    if not svgs:
        sys.stderr.write("没有找到 SVG\n")
        return 2

    out_dir = os.path.abspath(args.out) if args.out else None
    done, dxfs = [], []
    for s in svgs:
        d = out_dir or os.path.dirname(os.path.abspath(s))
        os.makedirs(d, exist_ok=True)
        dxf = os.path.join(d, os.path.splitext(os.path.basename(s))[0] + ".dxf")
        st = convert_one(s, dxf, scale=args.scale, default_h=args.height)
        tot = sum(st.values())
        print(f"✓ {os.path.basename(s)} → {os.path.basename(dxf)}  {tot} 实体 "
              f"(线{st['LINE']} 多段线{st['LWPOLYLINE']} 文字{st['TEXT']} "
              f"实心{st['SOLID']} 圆{st['CIRCLE']} "
              f"剖面线{st['HATCH']} 箭头{st['ARROW']})")
        done.append(dxf)
        dxfs.append(dxf)

    if args.dxf_only:
        print("（--dxf-only：未生成 DWG）")
        return 0

    oda = find_oda(args.oda)
    if not oda:
        sys.stderr.write(
            "!! 未找到 ODAFileConverter。DWG 这一步不能省（CAD 用户要靠它改图）。\n"
            "   装法见 references/dwg_export.md：\n"
            "     python svg2dwg.py … --oda <安装路径>/ODAFileConverter.exe\n"
            "     或设环境变量 ODA_CONVERTER\n")
        return 3

    # 同名 DXF 可能分散在多个目录，ODA 用临时目录中转，最后按各自 out 目录落盘
    by_dir = {}
    for p in dxfs:
        by_dir.setdefault(os.path.dirname(p), []).append(p)

    rc = 0
    for d, paths in by_dir.items():
        made = dxf_to_dwg(oda, paths, d, args.version)
        for dxf in paths:
            dwg = made.get(os.path.basename(dxf))
            if not dwg:
                rc = 4
                continue
            size = os.path.getsize(dwg)
            line = f"  → {os.path.basename(dwg)}  {size/1024:.1f} KB  {args.version}"
            if args.no_verify:
                print(line)
                continue
            ok, detail = verify_dwg(oda, dwg, dxf)
            print(line + f"  {'✅' if ok else '❌'} {detail}")
            if not ok:
                rc = 5
    if rc == 0:
        print(f"✓ 完成：{len(dxfs)} 张 DXF + DWG，回读校验全部一致")
    return rc


if __name__ == "__main__":
    sys.exit(main())
