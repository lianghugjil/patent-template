#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fig_lite.py —— 零依赖附图生成器（纯标准库）

专利附图是黑白线框图，图元只有直线、矩形、圆、折线、剖面线和部件标号。
本脚本用标准库把它们直接光栅化成 PNG，不依赖 Pillow / cairosvg / 浏览器，
因此在没有第三方库的环境里也能产出「带标号」的合格附图。

用法：
    python fig_lite.py 附图1.fig                 # 按文件里的 SAVE 行输出
    python fig_lite.py 附图1.fig -o 图1.png      # 指定输出名
    python fig_lite.py 附图1.fig --dpi 300       # 按目标宽度反算像素（排版用）

规格文件（每行一条，`#` 开头为注释；坐标原点左上、y 向下、单位像素）：
    SIZE  1200 800          画布宽高
    LW    5                 线宽（后续图元共用，可反复改）
    FONT  4                 字模放大倍数（5×7 点阵 × N）
    RECT  150 150 900 500   矩形描边：x y 宽 高
    FILL  400 300 200 100   矩形填实
    LINE  600 150 600 650   直线：x1 y1 x2 y2
    ARROW 600 120 600 200   带箭头的引线（箭头在第二点端）
    LINE  600 400 700 300   直线标注引线：x1 y1 x2 y2（默认）
    LEAD  600 400 700 300   S 形绕行引线：x1 y1 x2 y2 [弯幅]（直线会交叉时使用）
                            标号旁，两端各留 22% 直线段，中段反向对称弯；
                            缺省弯幅 = min(9, 距离×0.13)，短于 34 退化为直线）
    CIRCLE 600 400 90       圆描边：cx cy r
    DISC  600 400 40        圆填实
    POLY  200,200 400,200 400,300    折线（默认闭合=false，末位加 C 则闭合）
    HATCH 150 150 900 500   剖面线（45°，间距取线宽 ×4）
    TEXT  600 80 1          标号：以 (x,y) 为文字中心
    SAVE  fig1.png          输出文件名

只内置数字 0-9 与 A-Z、`- . / ( ) :` 字模；中文请勿写进专利附图（国知局附图
只用标号，说明文字一律在说明书正文），脚本遇到中文会跳过并在 stderr 提示。
"""

import argparse
import math
import os
import shutil
import struct
import subprocess
import sys
import tempfile
import zlib

# ---------------------------------------------------------------- 5×7 点阵字模
_FONT = {
    "0": ("01110", "10001", "10011", "10101", "11001", "10001", "01110"),
    "1": ("00100", "01100", "00100", "00100", "00100", "00100", "01110"),
    "2": ("01110", "10001", "00001", "00010", "00100", "01000", "11111"),
    "3": ("11111", "00010", "00100", "00010", "00001", "10001", "01110"),
    "4": ("00010", "00110", "01010", "10010", "11111", "00010", "00010"),
    "5": ("11111", "10000", "11110", "00001", "00001", "10001", "01110"),
    "6": ("00110", "01000", "10000", "11110", "10001", "10001", "01110"),
    "7": ("11111", "00001", "00010", "00100", "01000", "01000", "01000"),
    "8": ("01110", "10001", "10001", "01110", "10001", "10001", "01110"),
    "9": ("01110", "10001", "10001", "01111", "00001", "00010", "01100"),
    "A": ("01110", "10001", "10001", "11111", "10001", "10001", "10001"),
    "B": ("11110", "10001", "10001", "11110", "10001", "10001", "11110"),
    "C": ("01110", "10001", "10000", "10000", "10000", "10001", "01110"),
    "D": ("11100", "10010", "10001", "10001", "10001", "10010", "11100"),
    "E": ("11111", "10000", "10000", "11110", "10000", "10000", "11111"),
    "F": ("11111", "10000", "10000", "11110", "10000", "10000", "10000"),
    "G": ("01110", "10001", "10000", "10111", "10001", "10001", "01111"),
    "H": ("10001", "10001", "10001", "11111", "10001", "10001", "10001"),
    "I": ("01110", "00100", "00100", "00100", "00100", "00100", "01110"),
    "J": ("00111", "00010", "00010", "00010", "00010", "10010", "01100"),
    "K": ("10001", "10010", "10100", "11000", "10100", "10010", "10001"),
    "L": ("10000", "10000", "10000", "10000", "10000", "10000", "11111"),
    "M": ("10001", "11011", "10101", "10101", "10001", "10001", "10001"),
    "N": ("10001", "11001", "10101", "10011", "10001", "10001", "10001"),
    "O": ("01110", "10001", "10001", "10001", "10001", "10001", "01110"),
    "P": ("11110", "10001", "10001", "11110", "10000", "10000", "10000"),
    "Q": ("01110", "10001", "10001", "10001", "10101", "10010", "01101"),
    "R": ("11110", "10001", "10001", "11110", "10100", "10010", "10001"),
    "S": ("01111", "10000", "10000", "01110", "00001", "00001", "11110"),
    "T": ("11111", "00100", "00100", "00100", "00100", "00100", "00100"),
    "U": ("10001", "10001", "10001", "10001", "10001", "10001", "01110"),
    "V": ("10001", "10001", "10001", "10001", "10001", "01010", "00100"),
    "W": ("10001", "10001", "10001", "10101", "10101", "11011", "10001"),
    "X": ("10001", "10001", "01010", "00100", "01010", "10001", "10001"),
    "Y": ("10001", "10001", "01010", "00100", "00100", "00100", "00100"),
    "Z": ("11111", "00001", "00010", "00100", "01000", "10000", "11111"),
    "-": ("00000", "00000", "00000", "11111", "00000", "00000", "00000"),
    ".": ("00000", "00000", "00000", "00000", "00000", "01100", "01100"),
    "/": ("00001", "00010", "00010", "00100", "01000", "01000", "10000"),
    "(": ("00010", "00100", "01000", "01000", "01000", "00100", "00010"),
    ")": ("01000", "00100", "00010", "00010", "00010", "00100", "01000"),
    ":": ("00000", "01100", "01100", "00000", "01100", "01100", "00000"),
    " ": ("00000",) * 7,
}

SS = 2  # 超采样倍数（先放大画，再降采样，线条不毛刺）


class Canvas:
    def __init__(self, w, h):
        self.w = w * SS
        self.h = h * SS
        self.buf = bytearray(b"\xff") * (self.w * self.h)
        self.ops = []      # 矢量图元记录（供 DXF 导出，与 PNG 同源）
        self.box = None    # 裁边后的边界（缓冲坐标），DXF 用同一框对齐

    # ---- 基本笔触：以 (x,y) 为中心落一个 lw 方块
    def dot(self, x, y, lw):
        x = int(round(x * SS))
        y = int(round(y * SS))
        n = max(1, int(round(lw * SS)))
        half = n // 2
        x0, x1 = x - half, x - half + n
        y0, y1 = y - half, y - half + n
        if x1 < 0 or y1 < 0 or x0 >= self.w or y0 >= self.h:
            return
        x0, x1 = max(0, x0), min(self.w, x1)
        y0, y1 = max(0, y0), min(self.h, y1)
        W = self.w
        row = b"\x00" * (x1 - x0)
        for yy in range(y0, y1):
            self.buf[yy * W + x0: yy * W + x1] = row

    def line(self, x1, y1, x2, y2, lw):
        self.ops.append(("LINE", x1, y1, x2, y2, lw))
        dx, dy = x2 - x1, y2 - y1
        steps = int(max(abs(dx), abs(dy))) + 1
        if steps <= 1:
            self.dot(x1, y1, lw)
            return
        for i in range(steps + 1):
            t = i / steps
            self.dot(x1 + dx * t, y1 + dy * t, lw)

    def arrow(self, x1, y1, x2, y2, lw):
        """带箭头的引线（箭头在 (x2,y2) 端）。"""
        self.line(x1, y1, x2, y2, lw)
        dx, dy = x2 - x1, y2 - y1
        L = math.hypot(dx, dy) or 1.0
        ux, uy = dx / L, dy / L
        ln = max(12, lw * 5)
        for sgn in (1, -1):
            a = math.radians(28) * sgn
            vx = ux * math.cos(a) - uy * math.sin(a)
            vy = ux * math.sin(a) + uy * math.cos(a)
            self.line(x2, y2, x2 - vx * ln, y2 - vy * ln, lw)

    def rect(self, x, y, w, h, lw):
        self.line(x, y, x + w, y, lw)
        self.line(x + w, y, x + w, y + h, lw)
        self.line(x + w, y + h, x, y + h, lw)
        self.line(x, y + h, x, y, lw)

    def fill(self, x, y, w, h, _rec=True):
        if _rec:
            self.ops.append(("SOLID", x, y, w, h))
        X0, Y0 = int(round(x * SS)), int(round(y * SS))
        X1, Y1 = int(round((x + w) * SS)), int(round((y + h) * SS))
        X0, Y0 = max(0, X0), max(0, Y0)
        X1, Y1 = min(self.w, X1), min(self.h, Y1)
        if X1 <= X0 or Y1 <= Y0:
            return
        row = b"\x00" * (X1 - X0)
        for yy in range(Y0, Y1):
            self.buf[yy * self.w + X0: yy * self.w + X1] = row

    def clear(self, x, y, w, h):
        """把矩形区域刷白（给标号垫底，防止引线穿字）。"""
        X0, Y0 = int(round(x * SS)), int(round(y * SS))
        X1, Y1 = int(round((x + w) * SS)), int(round((y + h) * SS))
        X0, Y0 = max(0, X0), max(0, Y0)
        X1, Y1 = min(self.w, X1), min(self.h, Y1)
        if X1 <= X0 or Y1 <= Y0:
            return
        row = b"\xff" * (X1 - X0)
        for yy in range(Y0, Y1):
            self.buf[yy * self.w + X0: yy * self.w + X1] = row

    def circle(self, cx, cy, r, lw, fill=False):
        if r <= 0:
            return
        self.ops.append(("CIRCLE", cx, cy, r, lw, fill))
        if fill:
            R = int(round(r * SS))
            CX, CY = int(round(cx * SS)), int(round(cy * SS))
            for dy in range(-R, R + 1):
                span = int(math.sqrt(max(0, R * R - dy * dy)))
                y = CY + dy
                if not (0 <= y < self.h):
                    continue
                a, b = max(0, CX - span), min(self.w, CX + span + 1)
                if b > a:
                    self.buf[y * self.w + a: y * self.w + b] = b"\x00" * (b - a)
            return
        n = max(64, int(2 * math.pi * r))
        for i in range(n + 1):
            t = 2 * math.pi * i / n
            self.dot(cx + r * math.cos(t), cy + r * math.sin(t), lw)

    def poly(self, pts, lw, close=False):
        for i in range(len(pts) - 1):
            self.line(pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1], lw)
        if close and len(pts) > 2:
            self.line(pts[-1][0], pts[-1][1], pts[0][0], pts[0][1], lw)

    def lead(self, x1, y1, x2, y2, lw, bulge=None, steps=24):
        """S 形绕行引线（仅在直线会交叉或穿过无关部件时使用）。

        两端各留 22% 直线段——从零件引出、到标号旁都是正的；中段用三次
        贝塞尔在法向做一对反向偏移，形成对称 S 弯，中点仍落在两点连线上。
        DXF 侧按 24 段折线落盘，CAD 里可整体 JOIN 回一条多段线。
        """
        dx, dy = x2 - x1, y2 - y1
        L = math.hypot(dx, dy) or 1.0
        if L < 34:
            self.line(x1, y1, x2, y2, lw)
            return
        ux, uy = dx / L, dy / L
        nx, ny = -uy, ux
        b = min(13.0, 0.20 * L) if bulge is None else bulge
        k = 0.18
        ax, ay = x1 + ux * L * k, y1 + uy * L * k
        bx, by = x2 - ux * L * k, y2 - uy * L * k
        c1 = (ax + nx * b, ay + ny * b)
        c2 = (bx - nx * b, by - ny * b)
        pts = [(x1, y1), (ax, ay)]
        for i in range(1, steps + 1):
            t = i / steps
            mt = 1 - t
            pts.append((mt ** 3 * ax + 3 * mt * mt * t * c1[0] + 3 * mt * t * t * c2[0] + t ** 3 * bx,
                        mt ** 3 * ay + 3 * mt * mt * t * c1[1] + 3 * mt * t * t * c2[1] + t ** 3 * by))
        pts.append((x2, y2))
        self.poly(pts, lw)

    def hatch(self, x, y, w, h, lw):
        """45° 剖面线，间距 4×线宽；按矩形裁剪。"""
        gap = max(4, lw * 4)
        X0, Y0, X1, Y1 = x, y, x + w, y + h
        span = int(w + h)
        for k in range(-span, span, gap):
            # 直线 y = -x + c，c = k + X0 + Y1
            c = k + X0 + Y1
            ax, ay = X0, c - X0
            bx, by = X1, c - X1
            # 裁剪到矩形
            if ay < Y0:
                ax, ay = c - Y0, Y0
            if ay > Y1:
                ax, ay = c - Y1, Y1
            if by < Y0:
                bx, by = c - Y0, Y0
            if by > Y1:
                bx, by = c - Y1, Y1
            if max(ax, bx) < X0 or min(ax, bx) > X1:
                continue
            self.line(ax, ay, bx, by, max(1, lw // 2))

    def text(self, cx, cy, s, scale, bold=True):
        """以 (cx,cy) 为中心写字。中文等无字模字符跳过。"""
        chars = [c for c in s]
        cell_w, cell_h = 5, 7
        track = 2  # 字间距（点阵单位）：两位数标号不会粘成一个方块
        total_w = (len(chars) * (cell_w + track) - track) * scale
        total_h = cell_h * scale
        ox = cx - total_w / 2.0
        oy = cy - total_h / 2.0
        lw = max(1, scale)
        pad = max(6, scale * 2)     # 白底外扩，引线不会贴到字上
        self.clear(ox - pad, oy - pad, total_w + 2 * pad, total_h + 2 * pad)
        skipped = []
        for i, ch in enumerate(chars):
            g = _FONT.get(ch.upper())
            if g is None:
                skipped.append(ch)
                continue
            bx = ox + i * (cell_w + track) * scale
            for ry, row in enumerate(g):
                for rx, bit in enumerate(row):
                    if bit != "1":
                        continue
                    px = bx + rx * scale
                    py = oy + ry * scale
                    if bold and scale >= 2:
                        self.fill(px, py, scale, scale, _rec=False)
                    else:
                        self.dot(px + scale / 2, py + scale / 2, lw)
        shown = "".join(ch for ch in chars if ch.upper() in _FONT and ch != " ")
        if shown:
            self.ops.append(("TEXT", cx, cy, shown, scale))
        if skipped:
            sys.stderr.write(f"[!] 字模里没有这些字符，已跳过：{''.join(skipped)}"
                             f"（附图里请只用数字/字母标号）\n")
        return total_w, total_h

    # ---- 裁掉四周白边（图幅用到满，插进 docx 才不会是"小图"）
    def trim(self, margin):
        W, H = self.w, self.h
        buf = self.buf
        top, bot, left, right = H, -1, W, -1
        for y in range(H):
            row = buf[y * W:(y + 1) * W]
            if row.count(0xFF) == W:          # 整行全白
                continue
            a = row.find(b"\x00")
            b = row.rfind(b"\x00")
            if a < 0:
                continue
            if y < top:
                top = y
            bot = y
            left = min(left, a)
            right = max(right, b)
        if bot < 0 or right < 0:
            return False                       # 空图，别裁
        m = int(margin * SS)
        x0, y0 = max(0, left - m), max(0, top - m)
        x1, y1 = min(W, right + 1 + m), min(H, bot + 1 + m)
        nw, nh = x1 - x0, y1 - y0
        if nw < 40 or nh < 40:
            return False
        nb = bytearray(nw * nh)
        for y in range(y0, y1):
            nb[(y - y0) * nw:(y - y0 + 1) * nw] = buf[y * W + x0:y * W + x1]
        self.box = (x0, y0, x1, y1)
        self.buf, self.w, self.h = nb, nw, nh
        return True

    # ---- 降采样并写 PNG
    def save(self, path, trim=24):
        if trim is not None:
            self.trim(trim)
        W, H = self.w // SS, self.h // SS
        src = self.buf
        sw = self.w
        out = bytearray(W * H)
        for y in range(H):
            ro = y * W
            r0 = (y * SS) * sw
            r1 = r0 + sw
            for x in range(W):
                x0 = r0 + x * SS
                # SS×SS 平均
                v = 0
                for k in range(SS):
                    base = x0 + k * sw
                    for j in range(SS):
                        v += src[base + j]
                out[ro + x] = v // (SS * SS)
        write_png(path, W, H, out)


def write_png(path, w, h, gray):
    """写 8 位灰度 PNG（zlib + 手写块）。"""
    raw = bytearray()
    for y in range(h):
        raw.append(0)                      # filter type 0
        raw += gray[y * w:(y + 1) * w]

    def chunk(tag, data):
        c = struct.pack(">I", len(data)) + tag + data
        c += struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        return c

    png = b"\x89PNG\r\n\x1a\n"
    png += chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 0, 0, 0, 0))
    png += chunk(b"IDAT", zlib.compress(bytes(raw), 9))
    png += chunk(b"IEND", b"")
    d = os.path.dirname(os.path.abspath(path))
    if d and not os.path.isdir(d):
        os.makedirs(d, exist_ok=True)
    open(path, "wb").write(png)


def write_dxf(path, ops, box, page_w_mm=165.0):
    """把同一份矢量图元写成 DXF R12（AC1009）。

    为什么要它：DWG 是二进制专有格式，转出来必须有 ODA File Converter 或
    LibreDWG；DXF 是纯文本，标准库就能写，任何 CAD（AutoCAD / 中望 / 浩辰 /
    CAXA）都能打开编辑。零依赖环境里这就是唯一能给的「可编辑 CAD 文件」。

    坐标按打印尺寸折算：整幅图宽 = page_w_mm（默认 165 mm，即 A4 版心宽），
    所以在 CAD 里量到的尺寸就是印到纸上的尺寸，改图不用换算。
    图层：OUTLINE（轮廓）/ HATCH（剖面线）/ TEXT（标号），颜色 7（随背景反色，打印为黑）。
    """
    if not ops:
        return None
    if not box:                              # 没裁边时用包围盒兜底
        xs, ys = [], []
        for o in ops:
            if o[0] == "LINE":
                xs += [o[1], o[3]]
                ys += [o[2], o[4]]
            elif o[0] == "CIRCLE":
                xs += [o[1] - o[3], o[1] + o[3]]
                ys += [o[2] - o[3], o[2] + o[3]]
            elif o[0] == "SOLID":
                xs += [o[1], o[1] + o[3]]
                ys += [o[2], o[2] + o[4]]
            elif o[0] == "TEXT":
                xs.append(o[1])
                ys.append(o[2])
        if not xs:
            return None
        box = (min(xs) * SS, min(ys) * SS, max(xs) * SS, max(ys) * SS)

    x0, y0, x1, y1 = box
    w_px = max(1.0, (x1 - x0) / SS)
    k = page_w_mm / w_px                     # 1 逻辑像素 = k 毫米

    def tx(x):
        return (x - x0 / SS) * k             # x 是逻辑像素，k 是「逻辑像素 → mm」

    def ty(y):
        return (y1 / SS - y) * k             # 图像 y 向下，DXF y 向上

    out = []

    def g(code, val):
        if isinstance(val, float):
            val = f"{val:.4f}"
        out.append(f"{code}\r\n{val}\r\n")

    # ---- HEADER ----
    g(0, "SECTION"); g(2, "HEADER")
    g(9, "$ACADVER"); g(1, "AC1009")
    g(9, "$INSBASE"); g(10, 0.0); g(20, 0.0); g(30, 0.0)
    g(9, "$EXTMIN"); g(10, 0.0); g(20, 0.0); g(30, 0.0)
    g(9, "$EXTMAX"); g(10, page_w_mm); g(20, page_w_mm * (y1 - y0) / max(1.0, (x1 - x0))); g(30, 0.0)
    g(9, "$LTSCALE"); g(40, 1.0)
    g(0, "ENDSEC")

    # ---- TABLES：线型 + 图层 ----
    g(0, "SECTION"); g(2, "TABLES")
    g(0, "TABLE"); g(2, "LTYPE"); g(70, 1)
    g(0, "LTYPE"); g(2, "CONTINUOUS"); g(70, 64); g(3, "Solid line"); g(72, 65); g(73, 0); g(40, 0.0)
    g(0, "ENDTAB")
    # 全部 color 7（黑/白自适应色）：专利附图要求黑白线图，打印一律出黑；
    # 靠图层名区分轮廓/剖面线/标号，便于在 CAD 里按层选、按层改。
    layers = [("0", 7), ("OUTLINE", 7), ("HATCH", 7), ("TEXT", 7)]
    g(0, "TABLE"); g(2, "LAYER"); g(70, len(layers))
    for name, color in layers:
        g(0, "LAYER"); g(2, name); g(70, 0); g(62, color); g(6, "CONTINUOUS")
    g(0, "ENDTAB")
    g(0, "ENDSEC")

    # ---- BLOCKS（空） ----
    g(0, "SECTION"); g(2, "BLOCKS"); g(0, "ENDSEC")

    # ---- ENTITIES ----
    g(0, "SECTION"); g(2, "ENTITIES")
    n_tri = 0
    for o in ops:
        kind = o[0]
        if kind == "LINE":
            _, ax, ay, bx, by, lw = o
            # 剖面线线宽只有轮廓的一半，据此归层（hatch 用 max(1,lw//2) 画）
            lay = "HATCH" if lw <= 2 else "OUTLINE"
            g(0, "LINE"); g(8, lay)
            g(10, tx(ax)); g(20, ty(ay)); g(30, 0.0)
            g(11, tx(bx)); g(21, ty(by)); g(31, 0.0)
        elif kind == "CIRCLE":
            _, cx, cy, r, lw, fill = o
            if not fill:
                g(0, "CIRCLE"); g(8, "OUTLINE")
                g(10, tx(cx)); g(20, ty(cy)); g(30, 0.0); g(40, r * k)
            else:                            # 填实圆：用三角面片拼（R12 无 HATCH）
                N = 36
                for i in range(N):
                    a1 = 2 * math.pi * i / N
                    a2 = 2 * math.pi * (i + 1) / N
                    p1 = (cx + r * math.cos(a1), cy + r * math.sin(a1))
                    p2 = (cx + r * math.cos(a2), cy + r * math.sin(a2))
                    g(0, "SOLID"); g(8, "OUTLINE")
                    g(10, tx(cx)); g(20, ty(cy)); g(30, 0.0)
                    g(11, tx(p1[0])); g(21, ty(p1[1])); g(31, 0.0)
                    g(12, tx(p2[0])); g(22, ty(p2[1])); g(32, 0.0)
                    g(13, tx(p2[0])); g(23, ty(p2[1])); g(33, 0.0)
                    n_tri += 1
        elif kind == "SOLID":
            _, x, y, w, h = o
            # DXF SOLID 四角按 Z 序：左上→右上→左下→右下
            g(0, "SOLID"); g(8, "OUTLINE")
            g(10, tx(x)); g(20, ty(y)); g(30, 0.0)
            g(11, tx(x + w)); g(21, ty(y)); g(31, 0.0)
            g(12, tx(x)); g(22, ty(y + h)); g(32, 0.0)
            g(13, tx(x + w)); g(23, ty(y + h)); g(33, 0.0)
        elif kind == "TEXT":
            _, cx, cy, txt, scale = o
            hgt = 7 * scale * k
            g(0, "TEXT"); g(8, "TEXT")
            g(10, tx(cx)); g(20, ty(cy)); g(30, 0.0)
            g(40, hgt)
            g(1, txt)
            g(72, 1)                         # 水平居中
            g(11, tx(cx)); g(21, ty(cy)); g(31, 0.0)
            g(73, 2)                         # 垂直居中
    g(0, "ENDSEC")
    g(0, "EOF")

    d = os.path.dirname(os.path.abspath(path))
    if d and not os.path.isdir(d):
        os.makedirs(d, exist_ok=True)
    with open(path, "wb") as f:
        f.write("".join(out).encode("ascii", errors="replace"))
    return sum(1 for o in ops if o[0] != "TEXT") + n_tri


def parse(spec_text):
    """解析规格文本，返回 (canvas, [输出文件名...])。"""
    c = None
    lw = 4
    scale = 4
    trim = 24
    outs = []
    pending = []

    def flush():
        nonlocal pending, c
        if c is not None and pending:
            for item in pending:
                outs.append((c, item, list(c.ops)))
        pending = []

    for ln in spec_text.splitlines():
        ln = ln.split("#")[0].strip()
        if not ln:
            continue
        parts = ln.split(None, 1)
        cmd = parts[0].upper()
        rest = parts[1] if len(parts) > 1 else ""
        if cmd == "SIZE":
            w, h = [int(float(v)) for v in rest.replace(",", " ").split()[:2]]
            c = Canvas(w, h)
            pending = []
        elif c is None:
            raise SystemExit("[×] 规格文件第一句必须是 SIZE 宽 高")
        elif cmd == "TRIM":
            trim = max(0, int(float(rest)))
        elif cmd == "NOTRIM":
            trim = None
        elif cmd == "LW":
            lw = max(1, int(float(rest)))
        elif cmd == "FONT":
            scale = max(2, int(float(rest)))
        elif cmd == "RECT":
            x, y, w, h = [float(v) for v in rest.replace(",", " ").split()[:4]]
            c.rect(x, y, w, h, lw)
        elif cmd == "FILL":
            x, y, w, h = [float(v) for v in rest.replace(",", " ").split()[:4]]
            c.fill(x, y, w, h)
        elif cmd == "LINE":
            x1, y1, x2, y2 = [float(v) for v in rest.replace(",", " ").split()[:4]]
            c.line(x1, y1, x2, y2, lw)
        elif cmd == "ARROW":
            x1, y1, x2, y2 = [float(v) for v in rest.replace(",", " ").split()[:4]]
            c.arrow(x1, y1, x2, y2, lw)
        elif cmd == "LEAD":
            v = [float(t) for t in rest.replace(",", " ").split()[:5]]
            c.lead(v[0], v[1], v[2], v[3], lw,
                   bulge=(v[4] if len(v) > 4 else None))
        elif cmd == "CIRCLE":
            cx, cy, r = [float(v) for v in rest.replace(",", " ").split()[:3]]
            c.circle(cx, cy, r, lw, fill=False)
        elif cmd == "DISC":
            cx, cy, r = [float(v) for v in rest.replace(",", " ").split()[:3]]
            c.circle(cx, cy, r, lw, fill=True)
        elif cmd == "POLY":
            close = rest.strip().upper().endswith(" C")
            body = rest.strip()[:-1] if close else rest
            # 支持 "x,y x,y" 与 "x1 y1 x2 y2" 两种写法
            toks = [t for t in body.replace(",", " ").split() if t]
            vals = [float(t) for t in toks]
            pts = [(vals[i], vals[i + 1]) for i in range(0, len(vals) - 1, 2)]
            if len(pts) >= 2:
                c.poly(pts, lw, close=close)
        elif cmd == "HATCH":
            x, y, w, h = [float(v) for v in rest.replace(",", " ").split()[:4]]
            c.hatch(x, y, w, h, lw)
        elif cmd == "TEXT":
            toks = rest.split(None, 2)
            if len(toks) < 3:
                continue
            c.text(float(toks[0]), float(toks[1]), toks[2].strip(), scale)
        elif cmd == "SAVE":
            pending.append(rest.strip())
        else:
            sys.stderr.write(f"[!] 不认识的指令，已忽略：{ln}\n")
    flush()
    return c, outs, trim


def main():
    ap = argparse.ArgumentParser(
        description="零依赖附图生成器（纯标准库出 PNG，可选同时出 DXF R12）")
    ap.add_argument("spec")
    ap.add_argument("-o", "--out", default=None, help="指定第一张 PNG 的输出名")
    ap.add_argument("--dxf", action="store_true", default=None,
                    help="同时输出同名 .dxf（R12，CAD 可编辑）；默认开启")
    ap.add_argument("--no-dxf", dest="dxf", action="store_false",
                    help="只要 PNG，不出 DXF")
    ap.add_argument("--mm", type=float, default=165.0,
                    help="DXF 整幅宽度按多少毫米折算（默认 165，即 A4 版心宽）")
    ap.add_argument("--dwg", default=None, metavar="ODA",
                    help="给 ODA File Converter 路径时，顺手把 DXF 转成 DWG")
    args = ap.parse_args()
    if args.dxf is None:
        args.dxf = True

    if not os.path.exists(args.spec):
        raise SystemExit(f"[×] 规格文件不存在：{args.spec}")
    raw = open(args.spec, "rb").read()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("gb18030", errors="replace")

    c, outs, trim = parse(text)
    if c is None:
        raise SystemExit("[×] 规格文件里没有 SIZE")
    if not outs:
        outs = [(c, os.path.splitext(os.path.basename(args.spec))[0] + ".png", list(c.ops))]
    base = os.path.dirname(os.path.abspath(args.spec))
    made = []
    for idx, (cv, name, ops) in enumerate(outs):
        p = name if os.path.isabs(name) else os.path.join(base, name)
        if args.out and idx == 0:
            p = args.out
        cv.save(p, trim=trim)
        print(f"[✓] {p}  {os.path.getsize(p) // 1024} KB")
        made.append(p)
        if args.dxf:
            dp = os.path.splitext(p)[0] + ".dxf"
            n = write_dxf(dp, ops, cv.box, page_w_mm=args.mm)
            if n:
                print(f"[✓] {dp}  {os.path.getsize(dp) // 1024} KB  {n} 个图元"
                      f"（R12，图层 OUTLINE/HATCH/TEXT，图幅 {args.mm:g} mm）")

    if args.dwg:
        for p in made:
            dp = os.path.splitext(p)[0] + ".dxf"
            if not os.path.exists(dp):
                continue
            r = _dxf2dwg(dp, args.dwg)
            print(f"[{'✓' if r else '×'}] {os.path.splitext(dp)[0]}.dwg"
                  + ("" if r else "  （ODA 转换失败，DXF 仍可用）"))
    return 0


def _dxf2dwg(dxf_path, oda_exe):
    """借 ODA File Converter 把 DXF 转 DWG；失败返回 False，不抛异常。"""
    oda = oda_exe
    if os.path.isdir(oda):
        for cand in ("ODAFileConverter.exe", "ODAFileConverter"):
            if os.path.exists(os.path.join(oda, cand)):
                oda = os.path.join(oda, cand)
                break
    if not os.path.exists(oda):
        sys.stderr.write(f"[!] 找不到 ODA File Converter：{oda_exe}\n")
        return False
    src = os.path.dirname(os.path.abspath(dxf_path))
    tmp = tempfile.mkdtemp(prefix="dxf2dwg_")
    try:
        shutil.copy(dxf_path, tmp)
        # 参数：输入目录 输出目录 版本 类型 递归 审核 过滤
        cmd = [oda, tmp, tmp, "ACAD2018", "DWG", "0", "1", os.path.basename(dxf_path)]
        subprocess.run(cmd, check=True, timeout=180,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        out = os.path.join(tmp, os.path.splitext(os.path.basename(dxf_path))[0] + ".dwg")
        if not os.path.exists(out):
            return False
        shutil.move(out, os.path.splitext(dxf_path)[0] + ".dwg")
        return True
    except Exception as e:                    # 没装 ODA / 被杀毒拦 / 超时，都不影响 DXF
        sys.stderr.write(f"[!] DWG 转换失败：{e}\n")
        return False
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
