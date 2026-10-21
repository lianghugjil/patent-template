# -*- coding: utf-8 -*-
"""audit_exam_risk.py — 实质审查避坑自检（J 类，与 audit_patent.py 的 A~H 格式体检互补）。

audit_patent.py 管"格式对不对"，本脚本管"会不会被审查员驳"。

用法：
    python audit_exam_risk.py 专利.docx
    python audit_exam_risk.py 专利.md
    python audit_exam_risk.py 专利.docx --quiet      # 只输出问题，不列通过项

检查项（对应《专利法》《专利法实施细则》《专利审查指南》高频驳回理由）：
  J1  独立权利要求缺「其特征在于」              细则第 21 条 / 法第 26 条第 4 款
  J2  从权引用在后权项（引用了编号不小于自身的权项）  法第 26 条第 4 款（不清楚）
  J3  多项从权被多项从权引用                      细则第 22 条第 2 款
  J4  权项使用「约/左右/适当/必要时/厚/薄」等     法第 26 条第 4 款（保护范围不清楚）
  J5  权项带附图标记                              提示项（易被理解为限制保护范围）
  J6  摘要超过 300 字                             法第 26 条第 5 款 / 细则第 23 条
  J7  摘要含宣传性、商业性用语                    细则第 23 条（不得使用商业性宣传用语）
  J8  效果断言无数据支撑                         法第 26 条第 3 款 / 第 22 条第 3 款（答复时无据）
  J9  结构尺寸前后矛盾（按 φD×t 规格反推内径）    法第 26 条第 3 款（无法实现 / 自相矛盾）
  J10 规格自身不成立（壁厚 ≥ 外径/2）            法第 26 条第 3 款
  J11 说明书未点明要解决的「技术问题」            法第 26 条第 3 款 / 第 2 条第 2 款
  J12 权项技术特征在说明书中找不到对应记载        法第 26 条第 4 款（得不到说明书支持）
  J13 说明书未记载有益效果                        法第 26 条第 3 款 / 第 22 条第 3 款
  J14 权项数值范围的端点在说明书中无对应数据      法第 26 条第 4 款（范围得不到支持）
  J15 融合正确率与最强单类的差值撑不住样本量      法第 26 条第 3 款 / 第 22 条第 3 款（组合优越性无据）
  J16 校验类从权挂在分段等非必要限定之下          法第 26 条第 4 款（退守路径被无关限定收窄）
  J17 独权含多处选择结构（至少两类/中的至少一项）  法第 22 条第 3 款（每项特征都降为任选项）
  J18 独权只有数据动作、无作用于物质的步骤        法第 25 条 1 款(2) / 法第 2 条第 2 款（客体不适格）

退出码：0 全部通过；1 有 ERROR；2 只有 WARN。

说明：J 类是**风险扫描**，不是审查结论。ERROR 是硬矛盾（不改必被指出），
WARN 是"审查员很可能发问、且答复时手里没牌"，需人工判断后补材料或改写。
创造性有无对比文件、单一性、客体这类问题无法形式化，见 references/exam_risks.md 人工清单。
"""
import argparse
import collections
import math
import os
import re
import sys

CN = r"\u4e00-\u9fa5"
VOL_NAMES = ["权利要求书", "说明书摘要", "说明书", "说明书附图"]

# ---------------------------------------------------------------- 参数词表
STRUCT = ["外径", "内径", "壁厚", "厚度", "直径", "半径", "孔径", "宽度", "深度", "间距",
          "锥角", "角度", "长度", "高度"]
PERF = ["谐振频率", "截止频率", "中心频率", "声速", "密度", "增益", "带宽", "压力",
        "温度", "容积", "体积", "质量", "重量", "匝数", "圈数"]
PARAMS = STRUCT + PERF
UNIT = r"(mm|cm|m|kHz|MHz|Hz|dB|%|L|mL|g|kg|℃|°|匝|圈|毫米|厘米|米|赫兹|分贝|千克|克|摄氏)"
LEAD = (r"(?:为|取|选用|等于|达到|不小于|不大于|不低于|不超过|大于|小于|"
        r"≥|≤|>|<|=)?")
# φ114×6 / Φ50*3 这类管规格：直接给出外径与壁厚
RE_SPEC = re.compile(r"[φΦ￠]\s*(\d+(?:\.\d+)?)\s*[×xX*]\s*(\d+(?:\.\d+)?)")

# ---------------------------------------------------------------- 权项
RE_CLAIM_NO = re.compile(r"^(\d{1,2})\s*[.．、]\s*\S")
RE_DEP = re.compile(r"根据权利要求\s*(\d{1,2})(?:\s*(?:或|至|~|～|-|—)\s*(\d{1,2}))?")
RE_DEP_MULTI = re.compile(r"根据权利要求\s*\d+\s*(?:或|至|~|～|-|—)\s*\d+")
# 任意「权利要求N…M」引用串（含「执行权利要求1至9任一项所述的方法」，无「根据」前缀）
RE_DEP_ANY = re.compile(r"权利要求\s*\d+(?:\s*(?:或|至|~|～|-|—|、|,|，)\s*\d+)*"
                        r"\s*(?:中)?\s*(?:任一项|任意一项|任一)?")
VAGUE_CLAIM = ["大约", "约为", "约", "左右", "近似", "适当", "适中", "必要时", "酌情",
               "厚一些", "薄一些", "较大", "较小", "较高", "较低", "若干", "例如", "等等",
               "优选为", "最好是", "尽可能"]
PROMO = ["市场前景", "经济效益", "竞争优势", "填补空白", "国内领先", "国际领先",
         "巨大价值", "广阔应用前景", "极具", "必将", "引领", "颠覆"]
RE_EFFECT = re.compile(r"(显著|大幅|明显|极大|有效|充分)(提高|提升|改善|降低|减小|增强|拓展|抑制|缩短)")
RE_MARK = re.compile(r"([" + CN + r"]{2,10})(\d{2,3})(?![0-9.A-Za-z])")
MARK_SKIP = "图第式表步项例为于至达是约不小于多于少于"
NOUN_TAIL = ("腔", "孔", "管", "层", "件", "体", "板", "槽", "环", "栓", "阀", "壁", "面",
             "段", "筒", "柱", "片", "块", "线", "圈", "装置", "单元", "模块", "组件",
             "机构", "部件", "结构", "开口", "通孔", "通道", "锥面", "螺纹")

ERRORS, WARNS, INFOS = [], [], []


def err(m):
    ERRORS.append(m)


def warn(m):
    WARNS.append(m)


def info(m):
    INFOS.append(m)


# ================================================================ 读入与分册
def cut_by_volume_name(paras):
    """按正文册名行（说明书 / 权利要求书 …）切块。适用于 md 与未套版式的 docx。"""
    out, cur, name = [], [], None
    for t in paras:
        if t in VOL_NAMES:
            if name is not None:
                out.append((name, cur))
            name, cur = t, []
        elif name is not None:
            cur.append(t)
    if name is not None:
        out.append((name, cur))
    return out


def cut_by_section(doc):
    """按 docx 的 sectPr 切块，块名取各节页眉文字。适用于已套版式的 docx。"""
    from docx.oxml.ns import qn
    body = doc.element.body
    groups, cur = [], []
    for el in body.iterchildren():
        if el.tag != qn("w:p"):
            continue
        txt = "".join(x.text or "" for x in el.iter(qn("w:t"))).strip()
        cur.append(txt)
        pPr = el.find(qn("w:pPr"))
        if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
            groups.append(cur)
            cur = []
    if cur:
        groups.append(cur)

    names = []
    for sec in doc.sections:
        nm = "".join(p.text for p in sec.header.paragraphs).replace("\u3000", "").strip()
        names.append(nm)
    while len(names) < len(groups):
        names.append(names[-1] if names else "")
    return [(names[i] if i < len(names) else "", g) for i, g in enumerate(groups)]


def load(path):
    """返回 (blocks, n_tab, kind, flat_paras)。blocks = [(名, [段...])]。"""
    if path.lower().endswith((".md", ".markdown", ".txt")):
        raw = open(path, encoding="utf-8", errors="replace").read()
        raw = re.sub(r"```.*?```", "", raw, flags=re.S)          # svg / 代码块
        raw = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", raw)
        raw = re.sub(r"^\s*#+\s*", "", raw, flags=re.M)
        paras = [ln.strip() for ln in raw.splitlines()]
        return cut_by_volume_name(paras), 0, "md", paras

    from docx import Document
    doc = Document(path)
    paras = [p.text.strip() for p in doc.paragraphs]
    blocks = cut_by_volume_name(paras)
    # 册名行已被移入页眉（apply_patent_format.py 之后）→ 改按节切
    if not any(n in ("权利要求书", "说明书摘要") for n, _ in blocks):
        blocks = cut_by_section(doc)
    return blocks, len(doc.tables), "docx", paras


# ================================================================ J1~J5 权项
def split_claims(block):
    items, cur_no, buf = [], None, []
    for t in block:
        m = RE_CLAIM_NO.match(t)
        if m:
            if cur_no is not None:
                items.append((cur_no, "".join(buf)))
            cur_no, buf = int(m.group(1)), [t]
        elif cur_no is not None and t:
            buf.append(t)
    if cur_no is not None:
        items.append((cur_no, "".join(buf)))
    return items


def check_claims(claims):
    if not claims:
        warn("[J1] 未识别到权利要求（缺「权利要求书」册名或分节）")
        return
    indep = [(n, t) for n, t in claims if not RE_DEP.search(t)]
    bad = [n for n, t in indep if "其特征在于" not in t and "其特征是" not in t]
    if bad:
        err(f"[J1] 独立权项 {bad} 缺「其特征在于」（细则 21 条：应分前序与特征两部分）")
    else:
        info(f"[J1] 独立权项 {[n for n, _ in indep]} 均含「其特征在于」")

    later = []
    for n, t in claims:
        for m in RE_DEP.finditer(t):
            lo = int(m.group(1))
            hi = int(m.group(2)) if m.group(2) else lo
            for r in range(min(lo, hi), max(lo, hi) + 1):
                if r >= n:
                    later.append(f"权{n}→权{r}")
    if later:
        err(f"[J2] 从权引用了编号不小于自身的权项 {sorted(set(later))}（26 条 4 款）")
    else:
        info(f"[J2] 从权 {len(claims) - len(indep)} 项，引用指向均在前")

    multi = {n for n, t in claims if RE_DEP.search(t) and RE_DEP_MULTI.search(t)}
    hit = []
    for n, t in claims:
        for m in RE_DEP.finditer(t):
            lo = int(m.group(1))
            hi = int(m.group(2)) if m.group(2) else lo
            for r in range(min(lo, hi), max(lo, hi) + 1):
                if r in multi and r != n:
                    hit.append(f"权{n}→多项从权{r}")
    if hit:
        err(f"[J3] 多项从权被引用 {sorted(set(hit))}（细则 22 条 2 款，多项从权不得作为另一多项从权的基础）")
    else:
        info(f"[J3] 多项从权 {sorted(multi) or '无'}，无被多项从权引用")

    vague = []
    for n, t in claims:
        for w in VAGUE_CLAIM:
            k = t.find(w)
            if k >= 0:
                vague.append(f"权{n}「{w}」…{t[max(0, k - 8):k + 10]}…")
                break
    if vague:
        warn(f"[J4] 权项含不确定用语 {len(vague)} 处（26 条 4 款，保护范围不清楚）：" + "；".join(vague[:5]))
    else:
        info("[J4] 权项无不确定用语")

    marks = []
    for n, t in claims:
        for m in RE_MARK.finditer(t):
            nm = m.group(1)
            if nm[-1] in MARK_SKIP or nm[0] in MARK_SKIP:
                continue
            # 「根据权利要求11」里的数字是权项号，「内径106」里的数字是尺寸，都不是附图标记
            if "权利要求" in nm or nm.endswith(("要求", "据权", "引用")):
                continue
            if any(p in nm for p in PARAMS):
                continue
            marks.append(f"权{n}:{nm}{m.group(2)}")
    if marks:
        warn(f"[J5] 权项带附图标记 {len(marks)} 处：{marks[:6]}（不违法，但易被理解为限制保护范围，惯例删除）")
    else:
        info("[J5] 权项未带附图标记")


# ================================================================ J6~J7 摘要
def check_abstract(block):
    if not block:
        warn("[J6] 未找到「说明书摘要」块")
        return
    text = "".join(t for t in block if t and t not in VOL_NAMES)
    if not text:
        warn("[J6] 摘要正文为空")
        return
    n = len(re.sub(r"\s", "", text))
    if n > 300:
        err(f"[J6] 摘要 {n} 字，超过 300 字上限（细则 23 条）")
    else:
        info(f"[J6] 摘要 {n} 字（≤300）")
    hit = [w for w in PROMO if w in text]
    if hit:
        warn(f"[J7] 摘要含商业性/宣传性用语 {hit}（细则 23 条，应删除）")
    else:
        info("[J7] 摘要无宣传性用语")


# ================================================================ J8 效果无据
def check_effect(full):
    bad = []
    for m in RE_EFFECT.finditer(full):
        seg = full[max(0, m.start() - 45):m.end() + 45]
        if not re.search(r"\d", seg):
            bad.append(f"「{m.group(0)}」…{full[max(0, m.start() - 18):m.end() + 12]}…")
    if bad:
        warn(f"[J8] 效果断言附近无任何数据 {len(bad)} 处（26 条 3 款；答复创造性时无据可依）："
             + "；".join(bad[:4]))
    else:
        info(f"[J8] 效果断言 {len(RE_EFFECT.findall(full))} 处，均在含数据的上下文中")


# ================================================================ J9~J10 数值
def extract_params(full):
    got = collections.defaultdict(list)      # name -> [(value, unit, ctx)]
    for name in PARAMS:
        for m in re.finditer(re.escape(name), full):
            tail = full[m.end():m.end() + 18]
            mm = re.match(r"\s*" + LEAD + r"\s*[φΦ￠]?\s*(\d+(?:\.\d+)?)\s*" + UNIT, tail)
            if not mm:
                continue
            ctx = full[max(0, m.start() - 16):m.end() + mm.end() + 4].replace("\n", " ")
            got[name].append((float(mm.group(1)), mm.group(2), ctx))
    return got


def check_numbers(full):
    got = extract_params(full)

    specs = []          # (外径, 壁厚, 上下文)
    for m in RE_SPEC.finditer(full):
        specs.append((float(m.group(1)), float(m.group(2)),
                      full[max(0, m.start() - 12):m.end() + 6].replace("\n", " ")))
    for D, t, c in specs:
        got["外径"].append((D, "mm", c))
        got["壁厚"].append((t, "mm", c))

    if not got:
        warn("[J9] 未抽取到任何带单位的参数值（尺寸是否都写成了裸数字？）")
        return

    # ---- J10 规格自身成立性：t < D/2
    bad_spec = [f"φ{D:g}×{t:g}（壁厚 ≥ 外径/2，是实心棒而非管）" for D, t, _ in specs if t >= D / 2]
    if bad_spec:
        err("[J10] 管规格不成立：" + "、".join(bad_spec))
    else:
        info(f"[J10] 管规格 {[f'φ{D:g}×{t:g}' for D, t, _ in specs] or '无'} 均成立")

    # ---- J9 用规格反推合法内径，校验文中「内径」值
    legal_d = {round(D - 2 * t, 3) for D, t, _ in specs}
    legal_D = {round(D, 3) for D, t, _ in specs}
    legal_t = {round(t, 3) for D, t, _ in specs}

    bad_inner = [(v, c) for v, u, c in got.get("内径", []) if u == "mm"
                 and legal_d and round(v, 3) not in legal_d]
    if bad_inner:
        err("[J9] 内径与管规格矛盾（26 条 3 款，审查员按 D−2t 一算即穿）："
            + "；".join(f"文中内径 {v:g} mm，但 φ 规格只能推出 {sorted(legal_d)} mm（{c.strip()[:30]}）"
                        for v, c in bad_inner[:3]))
    elif legal_d:
        info(f"[J9] 内径校验通过：文中内径值均来自 φ 规格反推 {sorted(legal_d)} mm")

    # 只有一个规格且外径/壁厚/内径都唯一时，单独报一次自洽
    if len(specs) == 1:
        D, t, _ = specs[0]
        dset = {round(v, 3) for v, u, _ in got.get("内径", []) if u == "mm"}
        if len(dset) == 1:
            d = dset.pop()
            if abs(D - 2 * t - d) <= 0.05:
                info(f"[J9] 尺寸链自洽：{D:g} − 2×{t:g} = {d:g} mm")

    # ---- 结构尺寸多值：除已在规格集合内的，其余列出待人工判断
    noise = []
    for name in STRUCT:
        rows = got.get(name, [])
        vals = {round(v, 3) for v, u, _ in rows}
        if len(vals) <= 1:
            continue
        if name == "内径" and vals <= legal_d:
            continue
        if name == "外径" and vals <= legal_D:
            continue
        if name == "壁厚" and vals <= legal_t:
            continue
        seen, ctxs = set(), []
        for v, u, c in rows:
            if v in seen:
                continue
            seen.add(v)
            ctxs.append(f"{v:g}{u}（{c.strip()[:26]}）")
        noise.append(f"{name}：" + " / ".join(ctxs[:4]))
    if noise:
        warn(f"[J9] 结构尺寸多值 {len(noise)} 组（若确为不同部件可忽略，否则是前后矛盾）：\n        "
             + "\n        ".join(noise[:5]))

    # ---- 性能参数多值：工况不同属正常，只提示
    perf = []
    for name in PERF:
        vals = {round(v, 3) for v, u, _ in got.get(name, [])}
        if len(vals) > 1:
            perf.append(f"{name} {sorted(vals)[:5]}")
    if perf:
        info("[J9] 性能参数多值（不同工况，正常）：" + "；".join(perf))


# ================================================================ J11/J13 说明书
def check_spec(spec_text):
    if not spec_text:
        warn("[J11] 未找到说明书块")
        return
    # 「本发明的目的在于提供…」是审查实践认可的等价写法，与「所要解决的技术问题」同等效力
    if re.search(r"(技术问题|所要解决|亟待解决|亟需解决|针对上述[^。]{0,30}(缺陷|不足|问题))", spec_text):
        info("[J11] 已点明所要解决的技术问题")
    elif re.search(r"(本发明的目的|目的在于|旨在)", spec_text):
        info("[J11] 以「本发明的目的在于…」表述发明目的（审查实践认可，等价于所要解决的技术问题）")
    else:
        warn("[J11] 说明书未点明「所要解决的技术问题」（26 条 3 款 / 2 条 2 款；"
             "审查员会质疑技术方案缺少立项依据）")

    if re.search(r"(有益效果|技术效果|本发明的效果)", spec_text):
        info("[J13] 已记载有益效果")
    else:
        warn("[J13] 说明书未设「有益效果」段（答复创造性三步法时缺『实际解决的技术问题』依据）")


# ================================================================ J12/J14 支持
def check_support(claims, spec_text):
    if not spec_text:
        return
    indep = [(n, t) for n, t in claims if not RE_DEP.search(t)]
    # 虚词与动词性字符：出现任一即说明抓到的是一句话而不是名词短语
    STOP = set("的于在和或与为以从到使将对由其所述该则并等及且但不有据")

    def supported(w):
        """名词是否在说明书有据。整词不中时用末 4 字回退——
        权项常写「获取待选末端元件」，说明书写「获取末端元件的…」，核心名词仍是有据的。"""
        if w in spec_text:
            return True
        core = w[-4:] if len(w) > 4 else w
        return core in spec_text

    miss = []
    for n, t in indep:
        body = t.split("其特征在于", 1)[-1]
        cands = set()
        # 滑动窗口：finditer 是非重叠匹配，长词会吃掉后面的真名词
        for i in range(len(body)):
            for L in (8, 7, 6, 5, 4):
                w = body[i:i + L]
                if len(w) < L:
                    break
                if not w.endswith(NOUN_TAIL):
                    continue
                if re.search(r"[；。，、：;:,（）()\[\]0-9]", w):
                    continue        # 跨句读/含数字的窗口不是名词短语（"Leff；侧壁"这类假名词）
                if STOP & set(w):
                    continue
                if re.match(r"^(其|根据|按照|通过|采用|利用|基于|作为|位于|包括|设置|"
                            r"连接|固定|一种|上述)", w):
                    continue
                # 窗口里带动词说明抓到的是一截句子而不是名词短语（「温度剖面作线」），
                # 这种片段永远查不到支撑，只会制造假阳性。
                if any(v in w for v in ("作", "进行", "得到", "获取", "输出", "拟合", "求解",
                                        "换算", "合成", "确定", "计算", "判断", "获得",
                                        "处理", "形成", "使得", "以及", "并")):
                    continue
                cands.add(w)
                break
        for w in sorted(cands):
            if not supported(w):
                miss.append(f"权{n}「{w}」")
    if miss:
        warn(f"[J12] 权项技术名词在说明书中找不到 {len(miss)} 处（26 条 4 款，得不到说明书支持）："
             + "、".join(miss[:8]))
    else:
        info("[J12] 独立权项的技术名词均能在说明书找到对应记载")

    spec_nums = set(re.findall(r"\d+(?:\.\d+)?", spec_text))
    bad = []
    for n, t in claims:
        # 「根据权利要求1~11任一项」是引用范围，不是数值范围——端点当然不在说明书里。
        # 注意：载体权项常写「执行权利要求1至9任一项所述的方法」，前面没有「根据」，
        # 只拦 RE_DEP 会漏掉这一类，端点 9 被当成数值端点报「未在说明书出现」。
        t2 = RE_DEP_ANY.sub(" ", t)
        for m in re.finditer(r"(\d+(?:\.\d+)?)\s*(mm|cm|Hz|kHz|dB|%|℃|°)?\s*[~～至\-—]\s*"
                             r"(\d+(?:\.\d+)?)\s*(mm|cm|Hz|kHz|dB|%|℃|°)?", t2):
            lo, hi = m.group(1), m.group(3)
            lost = [x for x in (lo, hi) if x not in spec_nums]
            if lost:
                bad.append(f"权{n} 范围 {lo}~{hi}（端点 {'、'.join(lost)} 未在说明书出现）")
    if bad:
        warn(f"[J14] 权项数值范围端点缺支撑 {len(bad)} 处（26 条 4 款）：" + "；".join(bad[:4]))
    else:
        info("[J14] 权项数值范围端点均能在说明书找到")


# ================================================================ J15 配对比较
RE_PCT = re.compile(r"(\d{1,3}(?:\.\d+)?)\s*%")
RE_NSEG = re.compile(r"(\d{2,5})\s*个?\s*判定段")
FUSE_KW = r"同时启用|融合|组合使用|一并启用|合并投票|多类合并|合并使用"
SINGLE_KW = r"单独启用|仅启用|单类|单独使用|逐类启用|各自单独"


def _sents(text):
    return [s for s in re.split(r"[。；\n]", text) if s.strip()]


def _case_blocks(text):
    """按试验工况切块。不同工况样本量不同，跨块配对比率没有意义。

    只取【实施例X】之后的内容：之前是摘要与说明书概述，其中的数值是对实施例的
    复述或泛指，拿来配对只会制造假分组。全文没有实施例分节时才退回整篇。
    """
    parts = re.split(r"(?=【实施例)", text)
    if len(parts) < 2:
        return [text]
    out = []
    for part in parts[1:]:
        for sub in re.split(r"(?=其[一二三四五六七八九]为)", part):
            if sub.strip():
                out.append(sub)
    return out


def _pairs_in_block(block):
    """块内按关键词位置就近分配百分比：(fuse 值列表, single 值列表, 样本量)。

    一句话里同时出现「单独使用 96.7%」和「同时启用 98.3%」是常见写法，
    按句取百分比会把两个数混到一起，必须按关键词位置切开。
    """
    pat = re.compile(r"(?P<s>" + SINGLE_KW + r")|(?P<f>" + FUSE_KW + r")")
    fuse, single, n = [], [], None
    for s in _sents(block):
        # 样本量常写在没带百分比的那一句里（「试验段共120个判定段」），
        # 所以必须在按关键词切段之前先取一次，不能挂在百分比分支下。
        if n is None:
            m0 = RE_NSEG.search(s)
            if m0:
                n = int(m0.group(1))
        marks = [(m.start(), m.end(), "single" if m.group("s") else "fuse")
                 for m in pat.finditer(s)]
        for i, (_a, b, k) in enumerate(marks):
            seg = s[b:marks[i + 1][0] if i + 1 < len(marks) else len(s)]
            vals = [float(x) for x in RE_PCT.findall(seg)]
            if not vals:
                continue
            # 单类侧的多值列举是正常的（几个单类各自的正确率）；
            # 融合侧出现「分别为」说明是段长/阈值扫描序列，不是单点汇总，跳过。
            if k == "fuse" and "分别为" in seg:
                continue
            (fuse if k == "fuse" else single).extend(vals)
    return fuse, single, n


def check_pair_compare(full):
    """融合/多类合并的正确率必须与最强的单类做配对比较，且样本量要撑得住差值。

    常见写法是给出四五个单类正确率再给一个融合正确率，然后宣称「组合后显著提升」。
    但两个独立比例相减不是配对证据：样本 120 段时，独立两比例差的临界约 7 个百分点，
    差 5 个百分点完全可能只是噪声。真正该给的是同一批样本上的 2×2 计数。
    """
    blocks = _case_blocks(full)
    groups, seen = [], set()
    for blk in blocks:
        f, s, n = _pairs_in_block(blk)
        if f and s:
            key = (round(max(f), 1), round(max(s), 1), n)
            if key not in seen:
                seen.add(key)
                groups.append((max(f), max(s), n))
    if not groups:
        if _pct_in_sents_any(full):
            info("[J15] 见到正确率但对不上「融合 / 单类」成对数据，跳过配对比较")
        else:
            info("[J15] 未见融合与单类正确率的成对断言，跳过配对比较")
        return

    for p_fuse, p_max, n in groups:
        d = p_fuse - p_max
        tag = f"n={n}" if n else "样本量未写"
        if not n:
            warn(f"[J15] 融合 {p_fuse}% 比最强单类 {p_max}% 高 {d:.1f} 个百分点，"
                 "但未写样本量（多少判定段），无法判断是否显著")
            continue
        p = p_max / 100.0
        crit = 196 * math.sqrt(2 * p * (1 - p) / n)      # 单位：百分点
        if d < crit:
            warn(f"[J15] 融合 {p_fuse}% 仅比最强单类 {p_max}% 高 {d:.1f} 个百分点（{tag}）；"
                 f"按独立两比例粗估需 ≥{crit:.1f} 个百分点才显著"
                 "（26 条 3 款 / 22 条 3 款：创造性若靠这组差值支撑，差值不足即等于无支撑）")
        else:
            info(f"[J15] 融合 {p_fuse}% 高于最强单类 {p_max}% 达 {d:.1f} 个百分点（{tag}），粗估显著")

    if not re.search(r"配对|McNemar|同批样本|2×2|2x2|两两比较|逐段对比|同一批样本", full):
        warn("[J15] 未给配对计数（同一批样本上「融合对/单类错」与「融合错/单类对」的段数）；"
             "两个独立正确率相减不构成「组合优于单类」的证据")


def _pct_in_sents_any(text):
    return bool(re.search(FUSE_KW, text)) and bool(RE_PCT.search(text))


# ================================================================ J16 退守路径独立性
VERIFY_KW = ["再次通过", "相反方向", "反向通过", "校验", "一致性", "复核", "复测"]
SEG_KW = ["判定段的长度", "划分为", "段长", "按里程划分", "判定段按", "分段长度"]


def check_fallback_path(claims):
    """校验类从权不得挂在含分段等非必要限定的从权之下。

    校验手段（往返重访、一致性复核）与判定段长度分属两件事，绑在一条引用链上
    等于答复期要主张校验特征时，必须一并主张段长，退路凭空变窄。
    """
    if len(claims) < 3:
        info("[J16] 权项不足 3 项，跳过退守路径检查")
        return
    parent, body = {}, {}
    for n, t in claims:
        body[n] = t
        m = RE_DEP.search(t)
        parent[n] = int(m.group(1)) if m else None

    bad, ok = [], 0
    for n, t in claims:
        if not any(k in t for k in VERIFY_KW):
            continue
        ok += 1
        cur, chain = parent.get(n), []
        while cur:
            chain.append(cur)
            cur = parent.get(cur)
        for a in chain:
            if any(k in body.get(a, "") for k in SEG_KW):
                bad.append(f"权{n}（校验类）挂在 权{a}（含分段限定）之下")
                break
    if bad:
        warn("[J16] 校验类从权被非必要限定绑架：" + "；".join(bad)
             + "（校验与分段长度无技术必然联系；建议让校验类从权直接从属独权）")
    else:
        info(f"[J16] 校验类从权 {ok} 项，均未挂在分段限定之下")


# ================================================================ J17 独权特征封闭性
RE_OPTIONAL = re.compile(r"至少(?:两类|两项|两种|一种|一项)|中的至少一(?:项|个|种)|中的任意一(?:项|个|种)")


def check_distinguishing(claims):
    """独权里的「至少两类 / 中的至少一项」是范围摊薄器。

    每处选择结构都把对应特征降为任选项：既然可以任选，就没有哪一项是必要技术特征，
    审查员评述结合显而易见时不必逐一驳倒，创造性论证处处漏风。
    """
    indep = [(n, t) for n, t in claims if not RE_DEP.search(t)]
    if not indep:
        return
    n0, t0 = indep[0]
    hits = RE_OPTIONAL.findall(t0)
    if len(hits) >= 2:
        warn(f"[J17] 独权{n0} 含 {len(hits)} 处选择结构（{'、'.join(sorted(set(hits)))}）："
             "每项特征都写成任选项，等于没有任何一项是必要技术特征，创造性被摊薄；"
             "建议至少把核心区别特征改为必选")
    elif len(hits) == 1:
        info(f"[J17] 独权{n0} 含 1 处选择结构（{hits[0]}），尚在常规写法范围")
    else:
        info(f"[J17] 独权{n0} 特征为封闭式表述，无选择结构")

    # 载体权项不该重复方法独权的选择结构问题
    for n, t in indep[1:]:
        if "计算机程序" in t or "处理器" in t:
            if any(k in t for k in ("权利要求1至", "权利要求1~", "权利要求1—")) and len(hits) >= 2:
                info(f"[J17] 载体权项{n} 引用方法权项范围，继承其选择结构问题，不另计")


# ================================================================ J18 客体适格
# 法第 25 条第 1 款第（2）项：智力活动的规则和方法不授予专利权；
# 法第 2 条第 2 款：技术方案须用技术手段解决技术问题、产生技术效果。
# 落地点：独权若通篇只有「计算 / 比较 / 判断 / 输出」这类作用于数据与人脑的动作，
# 而对物质世界没有任何测量或处理动作，最容易被直接判为智力活动规则或纯信息呈现。
TECH_VERBS = ["测量", "传感", "采集", "采样", "检测", "埋设", "埋入", "布设", "布置", "植入",
              "安装", "加热", "冷却", "降温", "升温", "干燥", "烘干", "蒸煮", "焙烧", "烧结",
              "施加", "加压", "抽真空", "注入", "通入", "灌入", "抽取", "分离", "过滤", "混合",
              "搅拌", "剪切", "拉伸", "弯曲", "冲压", "焊接", "切割", "涂覆", "浸渍", "清洗",
              "照射", "辐照", "发射", "接收", "拍摄", "录影", "扫描", "驱动", "激励", "通电",
              "保持", "调节", "控制", "记录", "调节", "输送", "循环", "冲洗", "开挖", "钻孔"]
MIND_VERBS = ["计算", "比较", "判断", "判定", "确定", "选取", "选择", "排序", "分类", "统计",
              "评估", "预测", "匹配", "输出", "标记", "汇总", "折算", "查表", "检索", "拟合"]


def check_subject_matter(claims):
    """J18 客体适格性：独权不能只有数据动作、没有作用于物质世界的动作。"""
    indep = [(n, t) for n, t in claims if not RE_DEP.search(t)]
    if not indep:
        return
    n0, t0 = indep[0]
    body = t0.split("其特征在于", 1)[-1]
    tech = [v for v in TECH_VERBS if v in body]
    mind = [v for v in MIND_VERBS if v in body]
    if tech:
        info(f"[J18] 独权{n0} 含作用于物质的步骤（{'、'.join(sorted(set(tech))[:5])}），"
             f"整体构成技术方案，客体无虞")
        return
    if mind:
        warn(f"[J18] 独权{n0} 只见数据动作（{'、'.join(sorted(set(mind))[:5])}），"
             "未见测量/施加/处理等作用于物质世界的步骤：易被认定为智力活动规则或"
             "纯信息呈现（法 25 条 1 款(2)、法 2 条 2 款）；"
             "建议把传感器布设、取样式测量、激励或调节中的至少一项写进独权")
    else:
        warn(f"[J18] 独权{n0} 未识别到任何明确的技术动作，客体适格性存疑，请人工确认")


# ================================================================ 主流程
def main():
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument("path")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()

    blocks, n_tab, kind, flat = load(args.path)
    by_name = collections.defaultdict(list)
    for nm, g in blocks:
        if nm:
            by_name[nm].extend(g)
    full = "\n".join(flat)
    spec_text = "\n".join(by_name.get("说明书", []))
    if not spec_text:                       # 没有分册信息时退回全文
        spec_text = full

    claims = split_claims(by_name.get("权利要求书", [])) if by_name.get("权利要求书") else []
    check_claims(claims)
    check_abstract(by_name.get("说明书摘要", []))
    check_effect(full)
    check_pair_compare(full)
    check_numbers(full)
    check_spec(spec_text)
    if claims:
        check_support(claims, spec_text)
        check_fallback_path(claims)
        check_distinguishing(claims)
        check_subject_matter(claims)

    print("=" * 74)
    print(f"审查避坑自检：{os.path.basename(args.path)}   "
          f"[{kind}，{len(flat)} 段，表 {n_tab} 张，分册 {[n for n, _ in blocks]}]")
    print("=" * 74)
    if not args.quiet:
        for m in INFOS:
            print("  ✓", m)
    for m in WARNS:
        print("  ⚠", m)
    for m in ERRORS:
        print("  ✗", m)
    print("-" * 74)
    if ERRORS:
        print(f"结论：✗ {len(ERRORS)} 处 ERROR、{len(WARNS)} 处 WARN —— 提交前必须改")
        print("       ERROR 是审查员一眼可见的硬伤；WARN 会被问到且答复时手里没牌。")
        return 1
    if WARNS:
        print(f"结论：⚠ {len(WARNS)} 处 WARN，无硬伤 —— 逐条判断，能补材料就补，不能补就改写法")
        return 2
    print("结论：✅ 通过（J 类风险扫描）")
    print("       本脚本只查可形式化风险；创造性有无对比文件、单一性、客体仍需人工判断。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
