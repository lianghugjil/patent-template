# -*- coding: utf-8 -*-
"""audit_figures.py — 说明书附图的**图面几何与标号**体检（纯标准库）。

audit_patent.py 只能看 docx 里的文字，看不见图上画了什么；本脚本补这一段：
把 SVG 里的引线、标号、文字框解析成几何，逐条查图面是否合规。

用法：
    python audit_figures.py 附图目录/                      # 只查图面
    python audit_figures.py 附图目录/ --doc 专利.docx       # 加查「图面标号 ↔ 标号表」
    python audit_figures.py 附图目录/ --json                # 机器可读

检查项：
  G1 引线相交   两条标注引线互相穿过 —— 制图硬禁忌（SKILL.md §4.3）
  G2 引线压字   引线穿过别的标号数字/图注文字
  G3 引线压点   引线穿过别的引线起点圆点
  G4 标号重复   同一标号在同一幅图上出现两次
  G5 标号越界   标号数字跑出画布
  G6 引线过短   被指点到标号 < 18 px，线缩成一点、看不清指向
  G7 标号表缺口 图上出现了、标号表里没有（ERROR，审查员会问"71 是什么"）
  G8 标号表冗余 标号表里有、所有图上都没出现（WARN）
  G9 直线引线   引线是纯直线（INFO，默认规范）

为什么要有这一套：SKILL.md 早写了"引线不得交叉"，但它是给人看的规则，
没有机器检查就等于没有。本轮就是靠人眼才发现的。

退出码：0 通过；1 有 ERROR；2 只有 WARN。
"""
import json
import math
import os
import re
import sys
import xml.etree.ElementTree as ET

NS = "{http://www.w3.org/2000/svg}"
BLACK = ("#000000", "#000", "black")

ISSUES = []          # (level, code, msg)
FILE = [""]


def err(code, msg):
    ISSUES.append(("ERROR", code, msg))


def warn(code, msg):
    ISSUES.append(("WARN", code, msg))


def info(code, msg):
    ISSUES.append(("INFO", code, msg))


# --------------------------------------------------------------------------- #
# 图元解析
# --------------------------------------------------------------------------- #
def _num(v, d=0.0):
    try:
        return float(re.sub(r"[^0-9.+-eE]", "", str(v)))
    except (TypeError, ValueError):
        return d


def _pts_of_path(d, steps=16):
    """path 的 d → 折线点列（多条子路径时只取首条；C/S/Q/T 离散成折线）。"""
    toks = re.findall(r"[A-Za-z]|-?\d*\.?\d+(?:[eE][-+]?\d+)?", d or "")
    pts, cur, start, cmd, i = [], (0.0, 0.0), (0.0, 0.0), None, 0
    prev_ctrl = None
    has_curve = False

    def num():
        nonlocal i
        v = float(toks[i]); i += 1
        return v

    while i < len(toks):
        if re.match(r"[A-Za-z]", toks[i]):
            cmd = toks[i]; i += 1
            if cmd in "Zz":
                cur = start
                continue
        if cmd is None:
            break
        rel, c = cmd.islower(), cmd.upper()
        try:
            if c == "M":
                x, y = num(), num()
                x, y = (x + cur[0], y + cur[1]) if rel else (x, y)
                if pts:
                    break                      # 多条子路径：只看第一条
                pts.append((x, y)); cur = start = (x, y)
                cmd = "l" if rel else "L"
            elif c == "L":
                x, y = num(), num()
                x, y = (x + cur[0], y + cur[1]) if rel else (x, y)
                pts.append((x, y)); cur = (x, y)
            elif c == "H":
                x = num(); x = x + cur[0] if rel else x
                pts.append((x, cur[1])); cur = (x, cur[1])
            elif c == "V":
                y = num(); y = y + cur[1] if rel else y
                pts.append((cur[0], y)); cur = (cur[0], y)
            elif c in "CQST":
                has_curve = True
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
                else:
                    p1 = (2 * cur[0] - prev_ctrl[0], 2 * cur[1] - prev_ctrl[1]) if prev_ctrl else cur
                    p2 = p1
                    p3 = (num() + (cur[0] if rel else 0), num() + (cur[1] if rel else 0))
                for k in range(1, steps + 1):
                    t = k / steps; mt = 1 - t
                    if c in "CST":
                        x = (mt ** 3 * cur[0] + 3 * mt * mt * t * p1[0]
                             + 3 * mt * t * t * p2[0] + t ** 3 * p3[0])
                        y = (mt ** 3 * cur[1] + 3 * mt * mt * t * p1[1]
                             + 3 * mt * t * t * p2[1] + t ** 3 * p3[1])
                    else:
                        x = mt * mt * cur[0] + 2 * mt * t * p1[0] + t * t * p3[0]
                        y = mt * mt * cur[1] + 2 * mt * t * p1[1] + t * t * p3[1]
                    pts.append((x, y))
                prev_ctrl = p2; cur = p3
            elif c == "A":
                for _ in range(5):
                    num()
                x, y = num(), num()
                x, y = (x + cur[0], y + cur[1]) if rel else (x, y)
                pts.append((x, y)); cur = (x, y)
            else:
                i += 1
        except IndexError:
            break
    return pts, has_curve


def _text_box(el):
    """按 font-size 与 anchor 估文字框。汉字按 1.0 em，数字/拉丁按 0.55 em。"""
    s = "".join(el.itertext()).strip()
    if not s:
        return None
    size = _num(el.get("font-size"), 26)
    x, y = _num(el.get("x")), _num(el.get("y"))
    anchor = el.get("text-anchor", "start")
    w = sum(size * (1.0 if "\u4e00" <= ch <= "\u9fff" else 0.55) for ch in s)
    if anchor == "middle":
        x0 = x - w / 2
    elif anchor == "end":
        x0 = x - w
    else:
        x0 = x
    return dict(s=s, x0=x0, x1=x0 + w, y0=y - 0.76 * size, y1=y + 0.22 * size,
                cx=x0 + w / 2, cy=y - 0.27 * size)


def _in_clip(el, parents):
    """是否处在 <clipPath> 定义或 clip-path 分组内（剖面线不参与引线判定）。"""
    for p in parents:
        if p.tag.endswith("clipPath") or p.get("clip-path"):
            return True
    return el.get("clip-path") is not None


def collect(root):
    lines, dots, texts = [], [], []      # 引线候选、锚点、文字
    shapes = []                          # 零件轮廓（粗线），用于"引线穿零件"提示

    def walk(el, parents):
        tag = el.tag.replace(NS, "")
        sw = _num(el.get("stroke-width"), 1.0)
        stroke = (el.get("stroke") or "").strip().lower()
        fill = (el.get("fill") or "").strip().lower()
        dash = el.get("stroke-dasharray")
        clipped = _in_clip(el, parents)

        if tag == "circle" and fill in BLACK and _num(el.get("r")) <= 7 and not clipped:
            dots.append((_num(el.get("cx")), _num(el.get("cy"))))
        elif tag == "text":
            b = _text_box(el)
            if b:
                texts.append(b)
        elif tag == "line" and not clipped and not dash:
            p = [(_num(el.get("x1")), _num(el.get("y1"))),
                 (_num(el.get("x2")), _num(el.get("y2")))]
            if sw <= 2.0:
                lines.append(dict(pts=p, curve=False, sw=sw))
            else:
                shapes.append(p)
        elif tag == "path" and not clipped and not dash and fill == "none":
            pts, curve = _pts_of_path(el.get("d"))
            if len(pts) >= 2:
                if sw <= 2.0:
                    lines.append(dict(pts=pts, curve=curve, sw=sw))
                else:
                    shapes.append(pts)
        for ch in list(el):
            walk(ch, parents + [el])

    walk(root, [])
    return lines, dots, texts, shapes


# --------------------------------------------------------------------------- #
# 几何
# --------------------------------------------------------------------------- #
def seg_int(p1, p2, p3, p4, eps=1e-9):
    """两线段是否真相交（端点相接不算）。"""
    def cross(o, a, b):
        return (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0])

    def on_seg(o, a, b):
        return (abs(cross(o, a, b)) < 1e-6
                and min(a[0], b[0]) - eps <= o[0] <= max(a[0], b[0]) + eps
                and min(a[1], b[1]) - eps <= o[1] <= max(a[1], b[1]) + eps)

    d1, d2 = cross(p3, p4, p1), cross(p3, p4, p2)
    d3, d4 = cross(p1, p2, p3), cross(p1, p2, p4)
    if ((d1 > 0) != (d2 > 0)) and ((d3 > 0) != (d4 > 0)):
        return True
    for o, a, b in ((p1, p3, p4), (p2, p3, p4), (p3, p1, p2), (p4, p1, p2)):
        if on_seg(o, a, b):
            return True
    return False


def seg_box(p1, p2, b, pad=2.0):
    """线段是否穿过文字框内部（框按 pad 收缩，端点贴边不算压字）。"""
    x0, x1 = b["x0"] + pad, b["x1"] - pad
    y0, y1 = b["y0"] + pad, b["y1"] - pad
    if x1 <= x0 or y1 <= y0:
        return False
    pts = [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]
    if x0 <= p1[0] <= x1 and y0 <= p1[1] <= y1:
        return True
    if x0 <= p2[0] <= x1 and y0 <= p2[1] <= y1:
        return True
    for i in range(4):
        if seg_int(p1, p2, pts[i], pts[(i + 1) % 4]):
            return True
    return False


def dist(a, b):
    return math.hypot(a[0] - b[0], a[1] - b[1])


# --------------------------------------------------------------------------- #
# 单幅检查
# --------------------------------------------------------------------------- #
def check_svg(path):
    FILE[0] = os.path.basename(path)
    root = ET.parse(path).getroot()
    lines, dots, texts, shapes = collect(root)

    # 引线 = 一端落在锚点圆点（≤3 px）、另一端落在某个文字旁（≤22 px）的细线。
    # 两个条件缺一不可：只有锚点会把以 O′ 等几何点为端点的辅助线误当引线，
    # 只有文字会把坐标刻度旁的短划算进来。
    leaders = []
    for ln in lines:
        for end in (0, -1):
            tail = ln["pts"][end]
            if any(dist(tail, d) <= 3.0 for d in dots):
                others = ln["pts"] if end == 0 else ln["pts"][::-1]
                if dist(others[0], others[-1]) < 12:
                    break                    # 铰轴/中心线的短划，不是引线
                h = others[-1]
                if not any(b["x0"] - 22 <= h[0] <= b["x1"] + 22
                           and b["y0"] - 22 <= h[1] <= b["y1"] + 22 for b in texts):
                    break                    # 端头没有标号，是几何辅助线
                leaders.append(dict(pts=others, curve=ln["curve"],
                                    tail=others[0], head=others[-1]))
                break

    # G9 直线引线（仅提示）
    straight = sum(1 for L in leaders if not L["curve"])
    if straight:
        info("G9", f"直线引线 {straight}/{len(leaders)} 条（默认规范）")

    # 标号 = 纯数字、且有一条引线端点落在它边上（坐标刻度数字不算标号）
    # 容差与"引线另一端靠文字"一致取 22 px：字框坐标是估算出来的，
    # 有的制图把标号放在字框外沿，9 px 会把真标号判成裸数字。
    def _linked(b):
        for L in leaders:
            for p in (L["tail"], L["head"]):
                if (b["x0"] - 22 <= p[0] <= b["x1"] + 22
                        and b["y0"] - 22 <= p[1] <= b["y1"] + 22):
                    return True
        return False

    nums = [b for b in texts if re.fullmatch(r"\d{1,3}", b["s"]) and _linked(b)]
    bare = [b for b in texts if re.fullmatch(r"\d{1,3}", b["s"]) and not _linked(b)]
    if bare and len(bare) > 3:
        info("G10", f"无引线的纯数字 {len(bare)} 个，按坐标刻度处理不参与标号核对")

    # G1 引线两两相交
    for i in range(len(leaders)):
        for j in range(i + 1, len(leaders)):
            a, b = leaders[i]["pts"], leaders[j]["pts"]
            hit = None
            for m in range(len(a) - 1):
                for n in range(len(b) - 1):
                    if seg_int(a[m], a[m + 1], b[n], b[n + 1]):
                        hit = (a[m], a[m + 1])
                        break
                if hit:
                    break
            if hit:
                err("G1", f"引线相交：{_fmt(hit[0])}→{_fmt(hit[1])} 一带"
                          f"（第 {j + 1} 条穿过第 {i + 1} 条）")

    # G2 引线压字 / G3 引线压点 / G6 引线过短
    for k, L in enumerate(leaders):
        for b in texts:
            for m in range(len(L["pts"]) - 1):
                if seg_box(L["pts"][m], L["pts"][m + 1], b):
                    err("G2", f"引线压字：第 {k + 1} 条引线穿过「{b['s']}」"
                              f"（文字中心 {b['cx']:.0f},{b['cy']:.0f}）")
                    break
        for d in dots:
            if dist(d, L["tail"]) <= 3.0:
                continue
            for m in range(len(L["pts"]) - 1):
                if seg_int(L["pts"][m], L["pts"][m + 1], d, (d[0] + 0.01, d[1])):
                    err("G3", f"引线压点：第 {k + 1} 条引线穿过别的起点圆点 {_fmt(d)}")
                    break
        if dist(L["tail"], L["head"]) < 18:
            warn("G6", f"引线过短：{_fmt(L['tail'])}→{_fmt(L['head'])}")

    # G4 标号重复 / G5 越界
    seen = {}
    vb = root.get("viewBox", "")
    vx, vy, vw, vh = (list(map(float, vb.split())) + [0, 0, 1e9, 1e9])[:4] if vb else (0, 0, 1e9, 1e9)
    for b in nums:
        seen.setdefault(b["s"], []).append(b)
        if not (vx - 2 <= b["cx"] <= vx + vw + 2 and vy - 2 <= b["cy"] <= vy + vh + 2):
            err("G5", f"标号 {b['s']} 越出画布（{b['cx']:.0f},{b['cy']:.0f}）")
    for s, bs in seen.items():
        if len(bs) > 1:
            where = "；".join(f"{b['cx']:.0f},{b['cy']:.0f}" for b in bs)
            warn("G4", f"标号 {s} 在本图出现 {len(bs)} 次（{where}）——对称件只标一侧")

    return [b["s"] for b in nums], leaders, [b["s"] for b in bare]


def _fmt(p):
    return f"({p[0]:.0f},{p[1]:.0f})"


# --------------------------------------------------------------------------- #
# 标号表比对
# --------------------------------------------------------------------------- #
RE_PAIR = re.compile(r"(\d{1,3})\s*[：:]\s*([\u4e00-\u9fa5A-Za-z]{1,12})")


def label_table(doc_path):
    """从 docx 的附图说明段取「标号→名称」。成品稿册名在页眉里，故按全文扫引导句后一段。"""
    try:
        from docx import Document
    except ImportError:
        warn("G7", "未装 python-docx，跳过标号表比对（图面检查照做）")
        return None
    doc = Document(doc_path)
    paras = [p.text.strip() for p in doc.paragraphs]
    table = {}
    grab = False
    for t in paras:
        if "各标号所代表的部件列表如下" in t:
            grab = True
            t = t.split("如下", 1)[1]
        if grab:
            for num, name in RE_PAIR.findall(t):
                table.setdefault(int(num), name)
            if table and not re.search(r"\d\s*[：:]", t):
                break
            if len(table) >= 200:
                break
    if not table:
        warn("G7", f"未能从 {os.path.basename(doc_path)} 解析到标号表")
        return None
    return table


# --------------------------------------------------------------------------- #
def main(argv):
    if not argv:
        print(__doc__)
        return 0
    target = argv[0]
    doc_path = None
    as_json = "--json" in argv
    if "--doc" in argv:
        doc_path = argv[argv.index("--doc") + 1]
    files = ([os.path.join(target, f) for f in sorted(os.listdir(target)) if f.lower().endswith(".svg")]
             if os.path.isdir(target) else [target])
    if not files:
        print(f"未找到 SVG：{target}")
        return 1

    table = label_table(doc_path) if doc_path else {}
    all_nums, all_bare = set(), set()
    per_file = {}
    for f in files:
        nums, leaders, bare = check_svg(f)
        per_file[os.path.basename(f)] = (len(nums), len(leaders))
        all_nums |= {int(n) for n in nums}
        # 裸数字按**逐图**判定性质：同一幅图里已经有带引线的标号，说明该图用引线标注，
        # 剩下的裸数字就是坐标刻度；整幅图一个引线都没有（系统框图直标式），裸数字才算标号。
        # 早先把各图的裸数字合并后再判，图3 的纵轴刻度"3"会把图2 漏标的"3"顶掉，
        # G8（表里有、图上没有）因此静默通过。
        if not leaders:
            all_bare |= {int(n) for n in bare}
    if table:
        all_nums |= {n for n in all_bare if n in table}

    if doc_path:
        if table:
            missing = sorted(n for n in all_nums if n not in table)
            unused = sorted(n for n in table if n not in all_nums)
            if missing:
                err("G7", "图上出现、标号表里没有：" +
                    "、".join(f"{n}" for n in missing))
            if unused:
                warn("G8", "标号表里有、所有图上都没出现：" +
                     "、".join(f"{n}:{table[n]}" for n in unused))

    if as_json:
        print(json.dumps([dict(level=a, code=b, msg=c) for a, b, c in ISSUES],
                         ensure_ascii=False, indent=1))
    else:
        for f, (n, l) in per_file.items():
            print(f"  · {f}：标号 {n} 个、引线 {l} 条")
        for lv, code, msg in ISSUES:
            print(f"  [{lv}] {code} {msg}")

    nerr = sum(1 for i in ISSUES if i[0] == "ERROR")
    nwarn = sum(1 for i in ISSUES if i[0] == "WARN")
    print(f"\n合计：ERROR {nerr}、WARN {nwarn}")
    return 1 if nerr else (2 if nwarn else 0)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
