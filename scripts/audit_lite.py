#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""轻量体检 —— 只用 Python 标准库，不装 python-docx 也能跑。

两种模式：
    python audit_lite.py 稿子.md        # 内容模式：册序/摘要字数/权项编号与引用/定点尺寸/附图引用/符号/禁用词/对外泄露(L)
    python audit_lite.py 成品.docx      # 版式模式：分节/页眉/页码/行距/页边距/图宽/表格

它不替代 audit_patent.py（A~Z 全量）与 audit_exam_risk.py（J1~J18），
只保证「环境里没有第三方库时，体检这一步也不会落空、也有数字可报」。

退出码：有 ERROR 返回 1，否则 0。
"""
from __future__ import annotations

import math
import os
import re
import sys
import zipfile

BOOKS = ["说明书摘要", "摘要附图", "权利要求书", "说明书", "说明书附图"]
ALIAS = {
    "摘要": "说明书摘要", "说明书摘要": "说明书摘要",
    "摘要附图": "摘要附图",
    "权利要求书": "权利要求书", "权利要求": "权利要求书",
    "说明书": "说明书", "专利说明书": "说明书",
    "说明书附图": "说明书附图", "附图": "说明书附图",
}
SPEC_HEADINGS = ("技术领域", "背景技术", "发明内容", "附图说明", "具体实施方式")
BANNED_WORDS = ["效果良好", "显著提高", "大大提高", "取得了良好的效果", "本申请"]
PROCESS_WORDS = ["参数口径", "模型口径", "本模型", "不并入", "此处不展开", "并不矛盾", "作者对账"]
BAD_FAIL_CLAIMS = ["谐振被破坏", "频率梳消失", "驻波被破坏"]

# ---- L 类：对外稿泄露检查（构思过程 / 方案族 / 半成品痕迹，一律不得进五册）
# 这些词只属于「内部推演」，写进交付稿等于把方法论和备选路线一并送人。
LEAK_METHOD = [
    "TRIZ", "triz", "进化树", "矛盾矩阵", "物场分析", "物-场分析", "ARIZ", "理想最终结果",
    "IFR", "什帕科夫斯基", "发明原理", "科学效应库", "剪裁法", "S曲线", "S 曲线",
    "方案族", "候选方案", "备选方案", "优选方案", "方案卡", "可专利性", "创造性三步法",
    "风险等级", "授权概率", "打分", "评分", "创造性评分", "新颖性评分",
    "检索", "查新", "对比文件", "公开号", "同族", "法则", "标准解", "效应库",
    "条件分离", "时间分离", "空间分离", "分离原理", "物理矛盾", "技术矛盾",
    "向微观级", "向超系统", "理想化", "动态化",
]
LEAK_DRAFT = [
    "本报告", "本文档", "以上分析", "经分析", "分析表明", "推演表明", "建议申请",
    "待确认", "待补充", "待填", "TODO", "占位",
]
# 占位括号：只认「括号里写着待办」的那种。【】本身是实施例标题的正规写法
# （【实施例一】），拿裸「【」当半成品痕迹会把标准格式判成 ERROR。
DRAFT_BRACKET = re.compile(
    r"【[^】]{0,20}(待填|待补|待定|待确认|请补|请填|此处|略|TODO|todo|XX|xx|占位|删除|插入)[^】]{0,20}】"
    r"|【\s*】")
# 取舍痕迹：说明「我比较过几条路」，等于把没选的那几条也交代了
LEAK_TRADEOFF = ["权衡", "取舍", "比较后", "择优", "次优", "降级为", "退守方案", "备选", "推荐", "主推", "首选"]
LEAK_PATTERNS = [
    (r"方案\s*[一二三四五六七八九十ABCDE1-9](?=[\s与和及、，,。：:；;）)]|$)", "方案编号（方案一/方案A）"),
    (r"第\s*[一二三四五六七八九十\d]+\s*个?方案", "方案序号（第 N 个方案）"),
    (r"(检索|查新)\s*(表明|显示|到|结果|式)", "检索过程"),
    (r"对比文件\s*[0-9CNUS]", "对比文件号"),
    (r"公开号\s*[：:]?\s*[A-Z]{2}", "公开号"),
]

E, W = [], []


def err(tag, msg):
    E.append(f"[E] {tag}  {msg}")


def warn(tag, msg):
    W.append(f"[W] {tag}  {msg}")


# ------------------------------------------------------------------ 内容模式

def split_md(text):
    lines = text.replace("\r\n", "\n").split("\n")
    books, cur = {}, None
    for ln in lines:
        m = re.match(r"^#{1,3}\s*(.+?)\s*#*\s*$", ln)
        if m and m.group(1).strip() in ALIAS:
            cur = ALIAS[m.group(1).strip()]
            books.setdefault(cur, [])
            continue
        if cur is not None:
            books[cur].append(ln)
    return {k: "\n".join(v).strip() for k, v in books.items()}


def check_books(books):
    got = [b for b in BOOKS if b in books]
    if len(got) < 5:
        miss = [b for b in BOOKS if b not in books]
        err("L1 册序", f"缺册：{'、'.join(miss)}（现有：{'、'.join(got) or '无'}）")
    order = [b for b in BOOKS if b in books]
    if order != got:
        err("L2 册序", f"册序不对，应为 说明书摘要→摘要附图→权利要求书→说明书→说明书附图，当前 {'、'.join(order)}")


def check_abstract(books):
    ab = books.get("说明书摘要", "")
    n = len(re.sub(r"\s", "", ab))
    if not ab:
        return
    if n > 300:
        err("J6 摘要", f"摘要 {n} 字 > 300 字，超 {n - 300} 字")
    elif n > 285:
        warn("J6 摘要", f"摘要 {n} 字，已接近 300 字上限")
    if re.search(r"^(本发明|本实用新型)?\s*$", ab[:20]) and len(ab) < 10:
        warn("J6 摘要", "摘要内容过短")


def check_claims(books):
    cl = books.get("权利要求书", "")
    if not cl:
        return
    lines = [l.strip() for l in cl.split("\n") if l.strip()]
    nums = []
    for l in lines:
        m = re.match(r"^(\d{1,3})\s*[.．、]", l)
        if m:
            nums.append(int(m.group(1)))
    if not nums:
        err("J1 权项", "未识别到权利要求编号（应写成 1. 2. 3. …）")
        return
    if nums != list(range(1, len(nums) + 1)):
        err("J1 权项", f"权利要求编号不连续或重复：{nums}")
    if len(nums) == 1:
        warn("J1 权项", "只有 1 项权利要求，通常需要 3~8 项从属")

    # 引用关系：不得引用在后权项；多项从属不得作为多项从属的引用基础
    multi = set()
    for l in lines:
        m = re.match(r"^(\d{1,3})\s*[.．、]", l)
        if not m:
            continue
        cur = int(m.group(1))
        body = l[m.end():]
        refs = set()
        for mm in re.finditer(r"权利要求\s*(\d+)\s*(?:至|~|-|～|—|to)\s*(\d+)|权利要求\s*(\d+)", body):
            if mm.group(3):
                refs.add(int(mm.group(3)))
            else:
                a, b = int(mm.group(1)), int(mm.group(2))
                refs.update(range(min(a, b), max(a, b) + 1))
        if "任一项" in body or "任一" in body:
            multi.add(cur)
        for r in sorted(refs):
            if r >= cur:
                err("J2 权项", f"权利要求 {cur} 引用了在后的权利要求 {r}")
            if r in multi and cur not in multi:
                err("J3 权项", f"权利要求 {cur} 以多项从属权利要求 {r} 为引用基础（不允许）")

    if re.search(r"\d+(\.\d+)?\s*(mm|毫米|cm)", cl):
        warn("R1 权项", "权利要求中出现绝对尺寸（mm/cm），应改为相对量或单侧边界")
    if re.search(r"[φΦ]\s*\d+[\s\-]?[A-Z]{1,2}\d{1,2}\b", cl):
        err("R2 权项", "权利要求中出现公差代号（如 φ16 H7），属制造要求而非结构特征")
    if "$" in cl or "\\frac" in cl or re.search(r"[=≤≥]\s*[\wΔ]", cl):
        warn("V 权项", "权利要求中疑似含公式，结构案权项应尽量为 0 个")


def check_figures(books):
    spec = books.get("说明书", "")
    figs_book = books.get("说明书附图", "")
    refs = set(int(x) for x in re.findall(r"图\s*(\d+)", spec))
    in_figs = set(int(x) for x in re.findall(r"图\s*(\d+)", figs_book))
    n_img = len(re.findall(r"!\[[^\]]*\]\(", figs_book)) or (max(in_figs) if in_figs else 0)
    if n_img:
        for i in range(1, n_img + 1):
            if i not in refs:
                err("S 附图", f"图{i} 在说明书正文中未被引用（S 类必须逐幅实质引用）")
        if refs:
            first = [i for i in re.findall(r"图\s*(\d+)", spec)]
            if first and int(first[0]) != 1:
                warn("S2 附图", f"正文首次引用是图{first[0]}，应为图1")
        if "摘要附图" in books and not re.search(r"!\[[^\]]*\]\(", books["摘要附图"]):
            err("L3 摘要附图", "摘要附图册里没有图")
    elif "摘要附图" not in books and n_img == 0:
        warn("S 附图", "无附图（无附图时不写附图说明、正文不引「图1」，两册附图一并省略）")
    else:
        warn("S 附图", "说明书附图册为空")


def check_language(books):
    full = "\n".join(books.values())
    for w in BANNED_WORDS:
        for m in re.finditer(re.escape(w), full):
            line = full[:m.start()].count("\n") + 1
            err("P 禁用词", f"出现「{w}」（第 {line} 行）")
    for w in PROCESS_WORDS:
        if w not in full:
            continue
        # 「本模型口径」会同时命中「本模型」与「模型口径」，合并成一条
        if w == "模型口径" and "本模型口径" in full:
            continue
        shown = "本模型口径" if (w == "本模型" and "本模型口径" in full) else w
        line = full[:full.find(w)].count("\n") + 1
        err("P1 过程性用语", f"出现「{shown}」（第 {line} 行），正文不写对账口径")
    for w in BAD_FAIL_CLAIMS:
        if w in full:
            warn("U 失效断言", f"出现「{w}」，应写「目标频率处无峰、该处放大失效」这类宾语明确的表述")


# 单位不是符号，别逼作者给 Hz 写定义
UNIT_TOKENS = {"Hz", "kHz", "mm", "cm", "m", "s", "ms", "g", "kg", "N", "Pa", "kPa",
               "MPa", "GPa", "V", "mV", "A", "mA", "W", "mW", "dB", "rpm"}
# 「名词＋符号」并列式定义：这是中文专利最常见的写法，不该只认「为/表示/记作」
APPOSITION_NOUNS = ("半径", "直径", "长度", "宽度", "厚度", "距离", "间距", "夹角",
                    "角度", "面积", "体积", "频率", "质量", "刚度", "阻尼", "比值",
                    "系数", "时间", "周期", "速度", "加速度", "力", "力矩", "位移",
                    "应变", "电压", "电流", "功率", "温度", "压力", "偏心距", "行程",
                    "间隙", "阻抗", "容抗", "功率因数")


def check_symbols(books):
    """保守检查：像符号的 token 首现处 30 字内要有定义线索。

    两种定义写法都算数：① 符号前有「为/表示/记作/定义/即」；② 符号紧跟量名词（「半径R」）。
    """
    spec = books.get("说明书", "") + "\n" + books.get("权利要求书", "")
    if not spec:
        return
    seen, bad = set(), []
    for m in re.finditer(r"(?<![A-Za-z0-9])([A-Za-zΔδΩωφλμ][A-Za-z]?)(?![A-Za-z0-9])", spec):
        tok = m.group(1)
        if tok in seen or len(tok) < 1 or tok in UNIT_TOKENS:
            continue
        # 自然常数与序号（e^{-z/d} 的 e、T_{i} 的 i）不是待定义的物理量，
        # 强行要求「e为…」既不合中文专利写法，也会淹没真正漏定义的符号。
        if tok in ("e", "i", "j", "k") and (spec[m.end():m.end() + 2].startswith("^")
                                            or (m.start() and spec[m.start() - 1].isalpha())):
            continue
        seen.add(tok)
        ctx = spec[max(0, m.start() - 6):m.start() + 34]
        if any(k in ctx for k in ("为", "表示", "记作", "定义", "记", "即")):
            continue
        if any(spec[max(0, m.start() - 8):m.start()].endswith(n) for n in APPOSITION_NOUNS):
            continue
        bad.append(tok)
    if len(bad) > 3:
        warn("M1 符号", f"{len(bad)} 个符号首现处未见定义（样本：{'、'.join(bad[:6])}），逐个核对是否已就地说明")


def check_leak(books):
    """L 类：交付稿里不得残留构思过程、方法名、方案族、半成品占位。

    用户拿到的应该是「一份专利」，不是「一份分析报告 + 一份专利」。
    方案族、检索、方法论名一旦留在正文，等于把技术路线连同淘汰项一起公开。
    """
    full = "\n".join(books.values())
    if not full.strip():
        return

    def at(pos):
        return full[:pos].count("\n") + 1

    for w in LEAK_METHOD:
        for m in re.finditer(re.escape(w), full):
            err("L1 泄露", f"出现「{w}」（第 {at(m.start())} 行）——构思方法名不得进交付稿")
    for w in LEAK_DRAFT:
        if w in full:
            err("L2 半成品", f"残留「{w}」（第 {at(full.find(w))} 行）——交付稿不许有待填/报告口吻")
    m = DRAFT_BRACKET.search(full)
    if m:
        err("L2 半成品", f"占位括号「{m.group(0)}」（第 {at(m.start())} 行）——交付稿不许留待填处")
    for w in LEAK_TRADEOFF:
        if w in full:
            warn("L3 取舍痕迹", f"出现「{w}」，除非是技术本身固有含义，否则删掉（会暴露还比较过别的路线）")
    for pat, name in LEAK_PATTERNS:
        m = re.search(pat, full)
        if m:
            err("L1 泄露", f"{name}「{m.group(0).strip()}」（第 {at(m.start())} 行）")

    # 实施例数量：多到像在罗列方案族，就说明把候选方案写进去了
    spec = books.get("说明书", "")
    n_emb = len(re.findall(r"实施例\s*[一二三四五六七八九十\d]", spec))
    if n_emb > 7:
        warn("L4 方案族痕迹", f"说明书出现 {n_emb} 处实施例，核对是否把候选方案逐条写了（保留必要变体即可）")


def check_embodiments(books):
    spec = books.get("说明书", "")
    n = len(re.findall(r"实施例\s*[一二三四五六七八九十\d]", spec))
    if n > 5:
        warn("G 实施例", f"出现 {n} 处实施例，超过 5 个通常没有必要")


# ------------------------------------------------------------------ 版式模式

def check_docx(path):
    try:
        z = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError) as e:
        err("D0 文件", f"打不开：{e}")
        return
    names = z.namelist()
    doc = z.read("word/document.xml").decode("utf-8", "ignore") if "word/document.xml" in names else ""
    if not doc:
        err("D0 文件", "缺 word/document.xml，不是有效 docx")
        return

    n_sect = doc.count("<w:sectPr")
    n_hdr = len([n for n in names if re.match(r"word/header\d*\.xml", n)])
    n_ftr = len([n for n in names if re.match(r"word/footer\d*\.xml", n)])
    if n_sect < 5:
        err("D1 分节", f"只有 {n_sect} 节，五册应各自起页（各节独立页眉与页码）")
    if n_hdr < n_sect:
        err("D1 页眉", f"{n_sect} 节只有 {n_hdr} 个页眉，缺册名的节无法生成页眉")
    if n_ftr < n_sect:
        warn("D1 页脚", f"{n_sect} 节只有 {n_ftr} 个页脚")
    if doc.count("<w:pgNumType") < n_sect:
        err("D2 页码", "有节未设 pgNumType start=1，页码不会每册从 1 起")

    margins = re.search(r'<w:pgMar[^/]*w:top="(\d+)"[^/]*w:right="(\d+)"[^/]*w:bottom="(\d+)"[^/]*w:left="(\d+)"', doc)
    if margins:
        top, right, bottom, left = (int(x) / 567 for x in margins.groups())
        if abs(top - 2.5) > 0.1 or abs(bottom - 2.0) > 0.1 or abs(left - 2.5) > 0.1 or abs(right - 2.0) > 0.1:
            err("D3 页边距", f"实测 上{top:.2f}/下{bottom:.2f}/左{left:.2f}/右{right:.2f} cm，应为 2.5/2.0/2.5/2.0")
    if doc.count('w:line="460"') < 10:
        err("D4 行距", "固定行距 23 pt（w:line=460）的段落过少，正文未按版式排")
    spacing = len(re.findall(r'w:lineRule="exact"', doc))
    total_p = doc.count("<w:p>")
    if total_p and spacing / total_p < 0.7:
        warn("D4 行距", f"只有 {spacing / total_p:.0%} 的段落是固定行距，其余可能是自动行距")

    if doc.count("<w:tbl>") > 1:
        warn("D5 表格", f"全文 {doc.count('<w:tbl>')} 张表，专利说明书建议 0~1 张")

    hdr_names = []
    for n in sorted(n for n in names if re.match(r"word/header\d+\.xml", n)):
        xml = z.read(n).decode("utf-8", "ignore")
        txt = "".join(re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", xml, re.S))
        hdr_names.append(txt.replace("\u3000", "").replace(" ", ""))
    want = ["说明书摘要", "摘要附图", "权利要求书", "说明书", "说明书附图"]
    for i, (got, exp) in enumerate(zip(hdr_names, want), 1):
        if got != exp:
            err("D6 页眉册名", f"第 {i} 节页眉是「{got}」，应为「{exp}」")
    if len(hdr_names) > len(want):
        warn("D6 页眉册名", f"页眉 {len(hdr_names)} 个，多于五册")

    for n in [n for n in names if re.match(r"word/header\d+\.xml", n)]:
        xml = z.read(n).decode("utf-8", "ignore")
        if "w:pBdr" not in xml:
            warn("D6 页眉横线", f"{n} 缺 0.75 pt 下横线")

    # 图片宽度（说明书附图应铺满 16.5 cm ≈ 5940000 EMU）
    cx = [int(x) for x in re.findall(r'<wp:extent cx="(\d+)"', doc)]
    if cx:
        full = [c for c in cx if abs(c - 5940000) < 60000]
        if len(full) < len(cx) - 1:
            warn("D7 图宽", f"{len(cx)} 幅图中只有 {len(full)} 幅铺满版心（16.5 cm）")
    if re.search(r'<w:drawing>', doc) and "word/media/" not in " ".join(names):
        err("D7 图", "文档里有图形引用但没有图片文件")


# ------------------------------------------------------------------ 主流程

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    path = sys.argv[1]
    if not os.path.exists(path):
        print(f"[×] 文件不存在：{path}")
        return 2
    print(f"[体检] {path}")
    if path.lower().endswith((".md", ".markdown", ".txt")):
        print("模式：内容（标准库，无第三方依赖）")
        books = split_md(open(path, encoding="utf-8").read())
        check_books(books)
        check_abstract(books)
        check_claims(books)
        check_figures(books)
        check_language(books)
        check_symbols(books)
        check_embodiments(books)
        check_leak(books)
    else:
        print("模式：版式（标准库，无第三方依赖）")
        check_docx(path)

    for line in E + W:
        print("  " + line)
    print(f"\n结果：ERROR {len(E)} / WARN {len(W)}"
          + ("  →  改完 ERROR 再交付" if E else "  →  没有硬伤"))
    return 1 if E else 0


if __name__ == "__main__":
    sys.exit(main())
