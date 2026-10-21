# -*- coding: utf-8 -*-
"""audit_patent.py — 对**已有**专利 docx 做规范体检（改稿模式入口）。

用法：
    python audit_patent.py 专利.docx
    python audit_patent.py 专利.docx --quiet     # 只输出问题

检查项（对应 SKILL.md 各章硬约束）：
  A 公式/标记残留   $ \\mathrm \\frac \\left \\right _{ ^{ 出现在正文文字里
  B 禁用词          本申请 / 我 / 本文 / 我们
  C 权项编号        1..N 连续、无跳号重号
  D 标号体系        附表完整性、父号连续、子件继承、正文是否都出现、列表是否都用到
  E 名称与标号绑定  名称后是否漏标号（列表里的号却没跟名称）
  F 版式禁忌        具体实施方式内是否出现小标题
  G 计数约束        表格 ≤1、实施例 ≤4（>5 才告警）、图号与附图说明是否对齐
  H 公式对象        OMML 数量（0 说明公式没转成可编辑对象）
  K 类型适配（方法/算法类专用，结构类自动跳过，见 references/patent_types.md）
  K1 方法权项引导语  独权须以「包括以下步骤」引出
  K2 步骤编号        限一种风格、1..N 连续
  K3 流程图          方法/算法类附图须含流程图或框图
  K4 载体权项        算法类建议补「电子设备 / 存储介质」权项
  K5 步骤展开        权项 1 的每个步骤须在具体实施方式中逐条展开
  L 摘要附图（指南第一部分第一章 4.5.2）
  L1 有无 / 幅数     有附图就必须有「摘要附图」册，且只能是 1 幅
  L2 图注           摘要附图页不该重复「图N」图注
  L3 册序           摘要 → 摘要附图 → 权利要求书（顺序错即 ERROR）
  M 符号体系（每个符号首次出现都须说明）
  M1 从未定义       说明书里出现的数学符号从未写成「X为…」或「中文名X」
  M2 先用后定义     符号首次出现早于它的定义段（读者先看到生符号）
  M3 写法混用       同一符号既有 ∗(U+2217) 又有 *(ASCII)
  M4 摘要裸符号     摘要里出现未就地说明的符号
  N 数值可溯源
  N1 摘要数值       摘要中的数字+单位在说明书里找不到（易被判超范围）
  N2 有益效果数值   有益效果里的数字+单位在实施例里找不到出处
  O1 段落雷同       同一件事在两处各写一遍（相似度 ≥0.82）
  P1 过程性用语     口径 / 本模型 / 不并入 / 不再赘述 这类内部对账用语
  Q1 有益效果结构   被拆成多段却没有「第一/其一」序号，或整段没有分层
  R 权项固定尺寸
  R1 公称值         权项里写死具体尺寸（φ16 H7 / 长度140 mm），保护范围被钉在一点上
  R2 公差代号       权项里写配合公差代号（H7 / h6 / IT7）——属制造要求，不是结构特征
  S 附图引用
  S1 逐幅引用       附图说明里列了、具体实施方式里没引到（或只有编号顺带提及）
  S2 引用顺序       正文首次引用顺序须为 图1→图2→…，倒置即告警
  Y 孔口流量        同一句里 φd/Δp/Q 三件套按水反算，量级差 3 倍以上即介质没写清
  Z 孔-腔共振       给了 Hz 却只给板厚：反算有效颈长 > 1.4×板厚 时须说明端修正或实测

退出码：0 全部通过；1 有 ERROR；2 只有 WARN。
"""
import re
import sys
import collections

from docx import Document
from docx.oxml.ns import qn

M_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"

CN = r"\u4e00-\u9fa5"
RE_LABEL_PAIR = re.compile(r"(\d{1,3})\s*[：:]\s*([" + CN + r"A-Za-z]{1,12})")
RE_LABEL_USE = re.compile(r"([" + CN + r"]{2,12})(\d{2,3})(?![0-9.])")
RE_CAPTION = re.compile(r"^图\s*(\d+)\s*为")
RE_CLAIM = re.compile(r"^(\d{1,2})[.．、]\s*\S")
RE_EX = re.compile(r"实施例\s*(\d+)|【实施例([一二三四五六七八九十]+)】")
# 「结合图1」「第1步」「内径102 mm」这类不是部件标号，比对名称时要排除
SKIP_BEFORE = "图第式表步项例为于至达是约不小于多于少于取用设按理值调选"
QUANTITY_WORDS = ("内径", "外径", "直径", "孔径", "半径", "厚度", "长度", "深度", "宽度",
                  "高度", "间距", "距离", "容积", "声速", "频率", "增益", "带宽", "密度",
                  "重量", "质量", "压力", "温度", "角度", "页码", "型号")
UNIT_TAIL = re.compile(
    r"^\s*(?:kN|N|GPa|MPa|kPa|Pa|kΩ|MΩ|Ω|kV|V|mA|A|kW|W|GHz|kHz|MHz|Hz"
    r"|km|mm|cm|µm|μm|nm|min|ms|mL|kg|mg|g|m|h|s|r/min|rpm"
    r"|dB|%|℃|°|L"
    r"|小时|分钟|秒|年|月|日|天|周|次|倍|个|项|元|目|号|转)")
# 上面的顺序不能打乱：长单位必须排在它的前缀之前（kN 先于 N、kHz 先于 Hz、mm 先于 m、
# 小时 先于 时），否则「2 N / 6 dB / 10 min」这类"数字+单位"会被当成部件标号，整篇误报成 ERROR。
# 标点（。，、；）**不能**放进 UNIT_TAIL：那会把「压力传感器 3。」这种真标号也放行，
# 连带 D 类"正文未出现的标号"误报。句末的取值（"默认取 3。"）改由 SKIP_BEFORE 里的
# 取/用/设/按/调/选 等动词挡住。
BAN_WORDS = ["本申请", "本专利", "本文", "我们", "笔者"]
SUSPICIOUS = ["$", "\\mathrm", "\\frac", "\\left", "\\right", "_{", "^{",
              "\\cdot", "\\Delta", "\\varepsilon", "\\max", "\\sum", "\\pi"]

ERRORS, WARNS, INFOS = [], [], []


def err(msg):
    ERRORS.append(msg)


def warn(msg):
    WARNS.append(msg)


def info(msg):
    INFOS.append(msg)


def main(path):
    doc = Document(path)
    paras = [p.text.strip() for p in doc.paragraphs]
    full = "\n".join(paras)

    # ---------- A 公式残留 ----------
    for i, t in enumerate(paras):
        hits = [s for s in SUSPICIOUS if s in t]
        if hits:
            err(f"[A] 第{i}段残留公式标记 {hits}: {t[:70]}…")

    # ---------- B 禁用词 ----------
    for w in BAN_WORDS:
        n = full.count(w)
        if n:
            warn(f"[B] 禁用词「{w}」出现 {n} 次（应统一为「本发明」）")

    # ---------- 分章定位 ----------
    # 两套定位：草稿册名在正文里；成品稿（套过版式）册名已移入页眉，
    # 必须按 sectPr 分节 + 页眉册名还原，否则 C/K 类会静默跳过、还报"全部通过"。
    def norm(t):
        return re.sub(r"[\s\u3000]", "", t)

    def idx_of(key):
        for i, t in enumerate(paras):
            if norm(t) == key:
                return i
        return -1

    VOLS = ("权利要求书", "说明书摘要", "摘要附图", "说明书", "说明书附图")
    ranges = {}
    try:
        sec_names = [norm("".join(p.text for p in s.header.paragraphs))
                     for s in doc.sections]
        bounds = [i for i, p in enumerate(doc.paragraphs)
                  if p._p.pPr is not None and p._p.pPr.sectPr is not None]
        blocks, prev = [], -1
        for b in bounds:
            blocks.append((prev + 1, b))
            prev = b
        if prev < len(paras) - 1:
            blocks.append((prev + 1, len(paras) - 1))
        if len(sec_names) == len(blocks) and len(blocks) > 1:
            for nm, rg in zip(sec_names, blocks):
                if nm in VOLS:
                    ranges.setdefault(nm, rg)
    except Exception:
        ranges = {}

    i_claim = ranges["权利要求书"][0] if "权利要求书" in ranges else idx_of("权利要求书")
    i_abs = ranges["说明书摘要"][0] if "说明书摘要" in ranges else idx_of("说明书摘要")
    i_spec = ranges["说明书"][0] if "说明书" in ranges else idx_of("说明书")
    i_impl = idx_of("具体实施方式")
    if i_impl < 0:
        i_impl = i_spec
    i_figdesc = idx_of("附图说明")
    if ranges:
        info(f"[分册] 按页眉册名还原 {len(ranges)} 册：{'/'.join(ranges)}")

    i_abstract_body = idx_of("本发明公开")      # 摘要正文
    if i_abstract_body < 0:
        i_abstract_body = idx_of("本发明提供")

    # ---------- C 权项编号 ----------
    if "权利要求书" in ranges:
        end_claim = ranges["权利要求书"][1] + 1
    elif i_abs > i_claim >= 0:
        end_claim = i_abs
    else:
        end_claim = len(paras)
    claims = []
    if i_claim >= 0:
        # 册名行在正文里时从下一段起；按页眉还原时 i_claim 本身就是权项 1
        start_claim = i_claim if RE_CLAIM.match(paras[i_claim]) else i_claim + 1
        for i in range(start_claim, end_claim):
            m = RE_CLAIM.match(paras[i])
            if m:
                claims.append((int(m.group(1)), i))
        nums = [n for n, _ in claims]
        if nums:
            if nums != list(range(1, len(nums) + 1)):
                err(f"[C] 权项编号不连续：{nums}")
            else:
                info(f"[C] 权项 {len(nums)} 项，编号 1..{len(nums)} 连续")
        else:
            warn("[C] 未识别到权项编号（是否缺「权利要求书」册名行？）")

    # ---------- 判型（方法/算法/系统/结构，见 references/patent_types.md ----------
    def claim_text(n):
        """取第 n 项权项的全文（含其后未标号的续行，直到下一项权项）。"""
        for num, i in claims:
            if num != n:
                continue
            j = end_claim
            for num2, i2 in claims:
                if i2 > i:
                    j = i2
                    break
            return "\n".join(paras[i:j])
        return ""

    c1 = claim_text(1)
    is_method = bool(re.search(r"方法", c1[:90]))
    is_algo = is_method and bool(re.search(
        r"算法|神经网络|机器学习|深度学习|模型训练|损失函数|特征向量|特征提取|数据集|"
        r"计算机程序|处理器|分类器|回归|预测模型|编码器|解码器", full))
    is_system = (not is_method) and bool(re.search(r"系统", c1[:90]))
    kind = ("算法类" if is_algo else "方法类" if is_method
            else "系统类" if is_system else "结构类")
    if claims:
        info(f"[K] 判型：{kind}（按权项 1 主题词判定）")

    # ---------- D 标号体系 ----------
    labels = collections.OrderedDict()   # num(int) -> name
    if i_figdesc >= 0 and i_impl > i_figdesc:
        region = paras[i_figdesc:i_impl]
        for t in region:
            for num, name in RE_LABEL_PAIR.findall(t):
                labels.setdefault(int(num), name)
    if not labels:
        if is_method or is_algo:
            info("[D] 方法/算法类未使用附图标记（允许，见 patent_types.md 第二节）")
        else:
            warn("[D] 未解析到附图标记列表（检查「附图中，各标号所代表的部件列表如下：」）")
    else:
        nums = sorted(labels)
        parents = [n for n in nums if n < 100]
        subs = [n for n in nums if n >= 100]
        if parents and parents != list(range(1, len(parents) + 1)):
            err(f"[D] 父标号不连续：{parents}")
        for s in subs:
            p = s // 100
            if p not in labels:
                err(f"[D] 子标号 {s}:{labels[s]} 的父件 {p} 不存在")
        by_parent = collections.defaultdict(list)
        for s in subs:
            by_parent[s // 100].append(s % 100)
        for p, kids in by_parent.items():
            if sorted(kids) != list(range(1, len(kids) + 1)):
                err(f"[D] 父件 {p} 的子标号不连续：{sorted(kids)}")
        info(f"[D] 标号 {len(labels)} 个：父 {len(parents)}、子 {len(subs)}")

    # 正文使用情况
    body = "\n".join(paras[i_impl:]) if i_impl >= 0 else full

    def find_uses(text, num):
        """该标号在正文里被哪些「名称+数字」用法的名称命中（已剔除图号、尺寸等假命中）。"""
        out = []
        for m in re.finditer(r"([" + CN + r"]{2,12})" + str(num) + r"(?![0-9.A-Za-z])", text):
            nm = m.group(1)
            before = text[m.start(1) - 1] if m.start(1) else ""
            if before in SKIP_BEFORE:
                continue
            if any(q in nm for q in QUANTITY_WORDS):
                continue
            if UNIT_TAIL.match(text[m.end():m.end() + 6]):
                continue
            out.append(nm)
        return out

    not_used, name_mismatch = [], []
    for num, name in labels.items():
        uses = find_uses(body, num)
        if not uses:
            not_used.append(f"{num}:{name}")
        elif not any(name in u or u in name for u in uses):
            name_mismatch.append(f"{num} 列表「{name}」/ 正文「{uses[0]}」")
    if not_used:
        err(f"[D] 正文未出现的标号：{not_used}")
    if name_mismatch:
        warn(f"[E] 名称与列表不一致：{name_mismatch[:8]}")

    # ---------- E 名称后漏标号 / 名称与标号之间有空格 ----------
    for name in sorted(set(labels.values())):
        for t in (paras[i_impl:] if i_impl >= 0 else paras):
            idx = t.find(name)
            if idx < 0:
                continue
            tail = t[idx + len(name):idx + len(name) + 4]
            if re.match(r"[ \u3000]+\d{2,3}", tail):
                warn(f"[E]「{name}」与标号之间有空格：…{t[max(0, idx - 6):idx + len(name) + 5]}…")
            break
    for num in labels:
        for m in re.finditer(r"([" + CN + r"])\s+" + str(num) + r"(?![0-9.A-Za-z])", body):
            # 「2 N」「6 dB」「10 min」「3 kPa」是数字带单位，不是部件标号
            if UNIT_TAIL.match(body[m.end():m.end() + 6]):
                continue
            if m.group(1) in SKIP_BEFORE or any(q in body[max(0, m.start() - 6):m.end() + 2]
                                                 for q in QUANTITY_WORDS):
                continue
            err(f"[E] 标号 {num} 前有空格：…{body[max(0, m.start() - 8):m.end() + 3]}…")
            break

    # ---------- F 具体实施方式内小标题 ----------
    if i_impl >= 0:
        heads = []
        for i in range(i_impl + 1, len(paras)):
            t = paras[i]
            if not t or len(t) > 24:
                continue
            if t.startswith(("图", "实施例", "（", "(", "表")):
                continue
            if t[-1] in "。；，、：;,":
                continue
            if re.match(r"^[一二三四五六七八九十]+[、.]", t) or re.match(r"^\d+[.、]", t):
                heads.append(f"第{i}段「{t}」")
        if heads:
            err(f"[F] 具体实施方式内出现小标题：{heads}")
        else:
            info("[F] 具体实施方式内无小标题")

    # ---------- G 计数约束 ----------
    n_tab = len(doc.tables)
    if n_tab > 1:
        warn(f"[G] 表格 {n_tab} 张（规范：0~1 张，优先 0）")
    ex = sorted({int(a) if a else "一二三四五六七八九十".index(b) + 1
                 for a, b in RE_EX.findall(full) if a or b})
    if len(ex) > 5:
        warn(f"[G] 实施例 {len(ex)} 个（建议 ≤4；多条技术路线可分列，但须控制篇幅）")
    else:
        info(f"[G] 实施例 {len(ex)} 个、表格 {n_tab} 张")

    caps = []
    if i_figdesc >= 0 and i_impl > i_figdesc:
        for t in paras[i_figdesc:i_impl]:
            m = RE_CAPTION.match(t)
            if m:
                caps.append(int(m.group(1)))
    if caps and caps != list(range(1, len(caps) + 1)):
        err(f"[G] 附图说明图号不连续：{caps}")
    tail_caps = [int(m.group(1)) for t in paras for m in [re.match(r"^图\s*(\d+)$", t)] if m]
    if caps and tail_caps and sorted(tail_caps) != sorted(caps):
        warn(f"[G] 文末图号段落 {sorted(tail_caps)} 与附图说明 {caps} 不一致")

    # ---------- S 附图逐幅引用（附图说明只是目录，具体实施方式里必须逐幅引到） ----------
    # 只写「图2~图5 的曲线由…生成」这种区间统称不算引用：审查员会认为该图未被说明书说明，
    # 该图承载的技术特征也就没有文字支持（法 26 条 4 款 / 指南第一部分第一章 4.5.2）。
    if caps and len(caps) >= 2:
        if "说明书" in ranges:
            end_spec = ranges["说明书"][1] + 1
        else:
            end_spec = idx_of("说明书附图")
            if end_spec < 0:
                end_spec = len(paras)
        body = "\n".join(paras[max(i_impl, 0):end_spec])
        missing, vague = [], []
        for n in caps:
            if not re.search(r"图\s*%d(?!\d)" % n, body):
                missing.append(n)
            elif not re.search(r"(?:如|结合|参见|见)\s*图\s*%d(?!\d)|图\s*%d\s*[中所]" % (n, n), body):
                vague.append(n)
        if missing:
            warn(f"[S] 图{missing} 在具体实施方式中完全没有出现（附图说明列了，正文没引）")
        if vague:
            warn(f"[S] 图{vague} 只有编号顺带提及、无「如图X所示/图X中…」式引用")
        if not missing and not vague:
            info(f"[S] {len(caps)} 幅附图均在具体实施方式中被逐幅引用")

        # S2 首次引用顺序：必须按图1→图2→…的顺序首次出现。
        # 先引图5 回头再引图3、图4，属于「附图编号与说明书引用次序不对应」，
        # 审查员会怀疑补图或挪动过段落；修改时把前面那处引用改成「后文图N」或直接后移。
        first = []
        for n in caps:
            m = re.search(r"图\s*%d(?!\d)" % n, body)
            if m:
                first.append((m.start(), n))
        seq = [n for _, n in sorted(first)]
        if seq != sorted(seq):
            shown = "→".join(f"图{n}" for n in seq)
            warn(f"[S2] 附图首次引用顺序为 {shown}，与图号顺序不一致——"
                 "正文须按图1、图2…的顺序首次引用；先引用后面的图应改为「后文图N」或把该句后移")
        elif seq:
            info(f"[S2] 附图首次引用顺序 " + "→".join(f"图{n}" for n in seq) + " ✔")

    # ---------- U 「失效」断言的宾语（失效的到底是什么，必须说准） ----------
    # 反例：「V 大于 V* 时谐振被破坏 / 整支频率梳消失」。末端容积是纯声抗、反射系数模恒为 1，
    # 管内驻波不会因末端而停振；真正失效的是"设计目标频率处的声压放大"。
    # 断言里若同句没有限定词（该处/目标频率/频段内/放大），多半是把"功能失效"写成了"现象消失"。
    if i_impl >= 0:
        _u_end = ranges["说明书"][1] + 1 if "说明书" in ranges else idx_of("说明书附图")
        if _u_end < 0:
            _u_end = len(paras)
        PAT_DEAD = re.compile(r"谐振\s*(?:被破坏|被摧毁|失效|消失|不再发生|完全崩溃)"
                              r"|频率梳\s*(?:消失|完全失效|失效|崩溃)")
        PAT_SCOPE = re.compile(r"该处|该频点|该频率|目标频率|设计频段|频段内|处不再|放大(?:失效|丧失|随之)")
        bad = []
        for t in paras[i_impl:_u_end]:
            for sent in re.split(r"[。；\n]", t):
                if PAT_DEAD.search(sent) and not PAT_SCOPE.search(sent):
                    bad.append(sent.strip()[:34])
        if bad:
            for b in bad:
                warn(f"[U] 「{b}…」把失效写成了现象消失：容积/阻抗类判据不会让驻波停振，"
                     f"失效的是目标频率处的放大，请补限定词或改述")
        else:
            info("[U] 「谐振失效/频率梳消失」类断言均限定了失效对象")

    # ---------- H 公式对象 ----------
    n_math = len(doc.element.body.findall(f".//{M_NS}oMath"))
    if n_math == 0:
        warn("[H] 未发现 OMML 公式对象（若正文有公式，说明没转成可编辑对象）")
    else:
        info(f"[H] OMML 可编辑公式 {n_math} 个")

    # ---------- K 类型适配（方法/算法类专用） ----------
    if claims and (is_method or is_algo):
        claim_region = "\n".join(paras[i_claim:end_claim]) if i_claim >= 0 else ""
        fig_region = ("\n".join(paras[i_figdesc:i_impl])
                      if (i_figdesc >= 0 and i_impl > i_figdesc) else "")
        steps = [int(x) for x in re.findall(r"步骤\s*[（(]\s*(\d+)\s*[）)]", c1)]

        # K1 方法权项引导语
        if not re.search(r"包括(以下|如下)步骤|步骤如下|包括步骤", c1):
            err("[K1] 方法类独权未以「包括以下步骤」引出（别把方法写成装置权项）")
        else:
            info("[K1] 方法权项引导语「包括以下步骤」✔")

        # K2 步骤编号风格与连续性
        if steps:
            if re.search(r"(?<![A-Za-z])S\s?\d", c1) or re.search(r"(?:^|\n)\s*\d+\s*[)）]", c1):
                err("[K2] 步骤编号混用：同一篇只允许「步骤（N）」，不要混 S1 或 1）")
            if steps != list(range(1, len(steps) + 1)):
                err(f"[K2] 步骤编号不连续或重号：{steps}")
            else:
                info(f"[K2] 步骤编号「步骤（1）」..「步骤（{len(steps)}）」连续")

        # K3 必须有流程图
        if not re.search(r"流程|框图", fig_region):
            warn("[K3] 方法/算法类附图说明中没有流程图或框图（建议图1为流程图）")

        # K4 算法类载体权项
        if is_algo and not re.search(r"存储介质|电子设备|程序产品|芯片", claim_region):
            warn("[K4] 算法类未见「电子设备 / 计算机可读存储介质」载体权项（建议补）")

        # K5 权项步骤须在实施方式逐条展开
        if steps:
            miss = [n for n in steps
                    if not re.search(r"步骤\s*[（(]\s*%d\s*[）)]" % n, body)]
            if miss:
                err(f"[K5] 权项步骤 {miss} 未在具体实施方式中逐条展开")
            else:
                info(f"[K5] 权项 {len(steps)} 个步骤在实施方式中均已展开")

    # ---------- L 摘要附图（指南第一部分第一章 4.5.2） ----------
    def n_pics(lo, hi):
        """区间内的图片对象数（嵌入式 drawing + 老式 pict）"""
        c = 0
        for p in doc.paragraphs[lo:hi + 1]:
            c += len(p._p.findall(".//" + qn("w:drawing")))
            c += len(p._p.findall(".//" + qn("w:pict")))
        return c

    n_pic_all = (len(doc.element.body.findall(".//" + qn("w:drawing")))
                 + len(doc.element.body.findall(".//" + qn("w:pict"))))
    i_absfig = ranges["摘要附图"][0] if "摘要附图" in ranges else idx_of("摘要附图")

    if i_absfig >= 0:
        if "摘要附图" in ranges:
            lo, hi = ranges["摘要附图"]
        else:
            lo = i_absfig
            hi = (i_claim - 1) if i_claim > i_absfig else (len(paras) - 1)
        pics = n_pics(lo, hi)
        if pics == 0:
            err("[L1] 「摘要附图」册里没有附图（应当复用说明书附图中的一幅）")
        elif pics > 1:
            err(f"[L1] 「摘要附图」册里有 {pics} 个图片对象，只能是一幅")
        else:
            info("[L1] 摘要附图 1 幅 ✔")
        caps = [paras[k] for k in range(lo, min(hi, len(paras) - 1) + 1)
                if re.match(r"^图\s*\d+\s*$", paras[k])]
        if caps:
            warn(f"[L2] 摘要附图页出现图注「{caps[0][:16]}」——摘要附图一般不重复图注")
    elif n_pic_all > 0:
        err("[L1] 全文有附图，但缺「摘要附图」册：有附图须提交一幅最能说明技术方案的附图作为"
            "摘要附图，单独一页，排在说明书摘要之后、权利要求书之前（指南第一部分第一章 4.5.2）")
    else:
        warn("[L1] 全文未见附图，摘要附图可省——确认确实无附图，否则属于漏册")

    # L3 册序
    if i_absfig >= 0 and i_abs >= 0:
        if i_absfig < i_abs:
            err("[L3] 摘要附图排在说明书摘要之前；官方顺序是 摘要 → 摘要附图 → 权利要求书")
        elif i_claim >= 0 and i_absfig > i_claim:
            err("[L3] 摘要附图排在权利要求书之后；应排在摘要之后、权利要求书之前")
        else:
            info("[L3] 册序正确：说明书摘要 → 摘要附图 → 权利要求书")

    # ---------- M 符号体系（首次出现必须定义） ----------
    W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    MATH_WORDS = {"dB", "mm", "cm", "km", "m", "s", "h", "g", "kg", "mg", "L", "mL",
                  "Hz", "kHz", "MHz", "GHz", "min", "max", "sin", "cos", "tan", "cot",
                  "ln", "lg", "log", "exp", "rms", "MPa", "kPa", "GPa", "Pa", "mol",
                  "DN", "H7"}
    SYM_RE = r"(?<![A-Za-z0-9])[A-Za-zΔΣεηπφθλμρσω][A-Za-z0-9*∗]{0,3}(?![A-Za-z0-9])"

    def omml_tokens(el):
        """OMML → 符号碎片：上下标拼成一个 token（Leff / f0 / B3dB），分数按 分子/分母。"""
        def txt(e):
            return "".join(t.text or "" for t in e.iter(M_NS + "t")) if e is not None else ""

        out = []

        def walk(e):
            for ch in e:
                tag = ch.tag
                if tag == M_NS + "sSub":
                    out.append(txt(ch.find(M_NS + "e")) + txt(ch.find(M_NS + "sub")))
                elif tag == M_NS + "sSup":
                    # 上标是幂次（a⁶ / m²），不是新符号 —— 只取底数
                    out.append(txt(ch.find(M_NS + "e")))
                elif tag == M_NS + "f":
                    out.append(txt(ch.find(M_NS + "num")) + "/" + txt(ch.find(M_NS + "den")))
                elif tag == M_NS + "r":
                    out.append(txt(ch))
                else:
                    walk(ch)
        walk(el)
        return out

    def rtext(par):
        """段落全文：正文 run + 公式（按 token 就地展开，保序）"""
        out = []
        for ch in par._p.iterchildren():
            if ch.tag in (M_NS + "oMath", M_NS + "oMathPara"):
                out.extend(omml_tokens(ch))
            elif ch.tag == W_NS + "r":
                out.append("".join(t.text or "" for t in ch.iter(W_NS + "t")))
        return "".join(out)

    texts = [rtext(p) for p in doc.paragraphs]
    spec_lo = i_spec if i_spec >= 0 else 0
    spec_all = "\n".join(texts[spec_lo:])
    i_impl2 = i_impl if i_impl >= 0 else spec_lo

    # 收集全文出现过的数学符号（来自 OMML），归一化 ∗ → *
    sym_hits = collections.defaultdict(set)
    for i, par in enumerate(doc.paragraphs):
        for om in par._p.iter(M_NS + "oMath"):
            for tok in omml_tokens(om):
                for m in re.finditer(SYM_RE, tok):
                    s = m.group(0).replace("\u2217", "*")
                    if s in MATH_WORDS or s in ("π", "e", "i", "Δ"):
                        continue
                    if re.fullmatch(r"(?:c?m|mm|km|s|h|L)\d", s):
                        continue                      # m2 / cm2 / s2：单位的上标，不是符号
                    sym_hits[s].add(i)

    def sym_patterns(s):
        key = s.rstrip("*")
        if not key:
            return None
        k = re.escape(key)
        # X为… / X(f)表示… / 以X表示…（函数式定义）
        p1 = re.compile(r"(?<![A-Za-z0-9])" + k
                        + r"\s*(?:\([^)]{0,6}\))?\s*(?:为|表示|称为|称作|定义为|记为|记作|即|取|等于|是)")
        # 「中文名+符号」也是定义（例：临界容积V*、等效声长Leff、凹陷深度d）
        p2 = re.compile(r"[\u4e00-\u9fa5]{2,10}\s*" + k + r"(?![A-Za-z0-9])")
        # 倒装定义：为X / 记为X / 即X
        p3 = re.compile(r"(?:为|记为|记作|称为|即|取)\s*" + k + r"(?![A-Za-z0-9])")
        # 并列定义：E、t、ν 分别为…
        p4 = re.compile(r"(?<![A-Za-z0-9])" + k
                        + r"(?:\s*[、，,]\s*[A-Za-zΔΣεηπφθλμρσω*][A-Za-z0-9*]{0,3})*"
                        + r"\s*分别?为")
        # 公式紧跟释义：Cd=πa⁶(1−ν²)/(16Et³) 为所述…声顺（符号后必须紧跟等号，避免误放行）
        p5 = re.compile(r"(?<![A-Za-z0-9])" + k + r"\s*[=＝][^。；\n]{0,60}?(?:为|表示|称为)")
        return p1, p2, p3, p4, p5

    def is_defined(txt, pats):
        return any(p.search(txt) for p in pats)

    sym_undef, sym_late = [], []
    for s, idxs in sym_hits.items():
        pats = sym_patterns(s)
        if not pats:
            continue
        used = [i for i in sorted(idxs) if i >= spec_lo]
        if len(used) < 2:
            continue
        if not is_defined(spec_all, pats):
            sym_undef.append(s)
            continue
        def_par = next((i for i in range(spec_lo, len(texts))
                        if is_defined(texts[i], pats)), None)
        if def_par is not None and used[0] < def_par:
            # 定义紧跟在下一段、且以「式中/其中」开头 = 公式后的规范写法，不算先用后定义
            if def_par == used[0] + 1 and texts[def_par].lstrip().startswith(("式中", "其中")):
                continue
            sym_late.append((s, used[0], def_par))

    if sym_undef:
        err("[M1] 符号在说明书中从未定义：" + "、".join(sorted(sym_undef))
            + "（每个符号首次出现处须写明「XX为…」或「中文名+符号」）")
    if sym_late:
        detail = "、".join(f"{s}(第{a}段首现/第{b}段才定义)" for s, a, b in sym_late[:6])
        warn(f"[M2] {len(sym_late)} 个符号先使用后定义：{detail}")
    if not sym_undef and not sym_late and sym_hits:
        info(f"[M1] {len(sym_hits)} 个数学符号均首现即定义")

    # M3 星号写法一致
    if "\u2217" in full and re.search(r"[A-Za-z]\*", full):
        warn("[M3] 星号写法混用：同一符号既有 ∗(U+2217) 又有 *(ASCII)，须统一")

    # M4 摘要里的符号须在摘要内说明
    abs_body = texts[i_abstract_body] if i_abstract_body >= 0 else (
        texts[i_abs] if i_abs >= 0 else "")
    abs_sym = []
    for m in re.finditer(SYM_RE, abs_body):
        s = m.group(0).replace("\u2217", "*")
        if s in MATH_WORDS or s.rstrip("*") == "":
            continue
        pats = sym_patterns(s)
        if not pats:
            continue
        if not is_defined(abs_body, pats):
            abs_sym.append(s)
    if abs_sym:
        warn("[M4] 摘要中出现未说明的符号：" + "、".join(sorted(set(abs_sym)))
             + "（摘要宜改为文字表述，或就地加「X为…」）")

    # ---------- N 数值可溯源 ----------
    def num_unit(t):
        return [m.group(0) for m in re.finditer(
            r"\d+(?:\.\d+)?\s*(?:dB|%|Hz|kHz|MHz|mm|cm|m/s|L|mL|℃|倍)", t)]

    if abs_body:
        lost = [x for x in num_unit(abs_body)
                if not re.search(r"(?<![0-9.])" + re.escape(x.split()[0]) + r"(?![0-9])", spec_all)]
        if lost:
            err(f"[N1] 摘要中的数值在说明书里找不到：{lost}（摘要不得出现正文没有的数，易被判超范围）")
        else:
            info("[N1] 摘要数值均可在说明书溯源")

    # 有益效果里的数必须在实施例（具体实施方式）里出现
    i_ben = next((i for i, t in enumerate(texts) if "有益效果" in t), -1)
    if i_ben >= 0:
        end_ben = min([x for x in (i_figdesc, i_impl2) if x > i_ben] or [len(texts)])
        ben_txt = "\n".join(texts[i_ben:end_ben])
        impl_txt = "\n".join(texts[i_impl2:])
        b1 = [m.group(0) for m in re.finditer(
            r"\d+(?:\.\d+)?\s*(?:dB|%|倍)", ben_txt)]
        # 只追大数：绝对值指标（≥10 dB / ≥10%）；0.3 dB、1.2 dB 这类差值由两个实测值相减得到，不算无据
        b1 = [x for x in b1 if float(re.match(r"[\d.]+", x).group(0)) >= 10]
        lost = [x for x in b1
                if not re.search(r"(?<![0-9.])" + re.escape(x.split()[0]) + r"(?![0-9])", impl_txt)]
        if lost:
            warn(f"[N2] 有益效果中的数值在实施例里找不到：{lost[:6]}（效果数须有实测/计算出处）")

    # ---------- O 高度雷同句 ----------
    import difflib
    sents = []
    for i in range(spec_lo, len(texts)):
        for s in re.split(r"[。；]", texts[i]):
            s = s.strip()
            if len(s) >= 22 and not re.match(r"(实施例|一种|本发明|本实施例|【实施例)", s):
                sents.append((i, s))
    dups, seen_pairs = [], set()
    for a in range(len(sents)):
        for b in range(a + 1, len(sents)):
            ia, sa = sents[a]
            ib, sb = sents[b]
            if ia == ib or abs(len(sa) - len(sb)) > 14:
                continue
            key = (min(ia, ib), max(ia, ib))
            if key in seen_pairs:
                continue
            if difflib.SequenceMatcher(None, sa, sb).ratio() >= 0.82:
                seen_pairs.add(key)
                dups.append(f"第{ia}段与第{ib}段：{sa[:26]}…")
    if dups:
        warn(f"[O1] {len(dups)} 组段落高度重复（同一件事在两处各写一遍）：" + "；".join(dups[:3]))

    # ---------- P 过程性/内部用语 ----------
    # 注意：「口径」本身是管道术语（大口径 / 小口径 / 公称口径），不能一刀切；
    # 只有「参数口径 / 口径一致 / 口径对齐」这类才是作者对账用语。
    META = [r"(?:参数|计算|模型|设计|统计|数值|数据|量纲)口径",
            r"口径(?:一致|统一|相同|不同|对齐|不符|上一致)",
            r"本模型", r"模型内部", r"内部约定", r"不并入",
            r"此处不展开", r"不再赘述"]
    hit_meta = []
    for i in range(spec_lo, len(texts)):
        for pat in META:
            m = re.search(pat, texts[i])
            if m:
                hit_meta.append(f"第{i}段「{m.group(0)}」")
    if hit_meta:
        warn("[P1] 出现过程性/内部用语（读者不需要的自我对账）：" + "、".join(hit_meta[:6]))

    # ---------- Q 有益效果段结构 ----------
    i_head = next((i for i, t in enumerate(texts) if re.fullmatch(r"本发明的有益效果[在于]*[：:]", t.strip())), -1)
    if i_head >= 0:
        seg, seg_idx = [], []
        for k, t in enumerate(texts[i_head + 1:i_head + 10], start=i_head + 1):
            if not t.strip() or t.strip().startswith(("附图说明", "图1", "图 1")):
                break
            seg.append(t.strip())
            seg_idx.append(k)
        # 序号三种形态都算：文字序号「第一/其一」、(1)/(一) 手打编号、Word 自动编号（numPr，文本里没有）
        def _auto_numbered(idx):
            try:
                pPr = doc.paragraphs[idx]._p.find(qn("w:pPr"))
            except Exception:
                return False
            return pPr is not None and pPr.find(qn("w:numPr")) is not None
        has_seq = any(re.match(r"(第[一二三四五六七八九十]+|其[一二三四五六七八九十]+|[（(]?\d+[）).、])", s)
                      for s in seg)
        if not has_seq and seg_idx and all(_auto_numbered(i) for i in seg_idx):
            has_seq = True
        if len(seg) >= 2 and not has_seq:
            warn(f"[Q1] 有益效果被拆成 {len(seg)} 段，且没有「第一/其一」序号——要么合成一段，要么按序编号")
    elif i_ben >= 0 and not re.search(r"第一|其一|首先", "\n".join(texts[i_ben:i_ben + 3])):
        warn("[Q1] 有益效果未用「第一…第N」分层，答复创造性时不易逐条对应")

    # ---------- R 权项固定尺寸 ----------
    # 权项里写死「直径为φ16 H7」「缩径衬筒的长度为140 mm」，保护范围就被钉在这一点上：
    # 别人做 φ20 或 150 mm 都不落入，等于把范围白送。限定应当用相对量（x ≤ 0.08L）
    # 或单侧边界（d 不大于 8 mm）——这两类只划一条线，不锁一个点。
    if claims:
        # 只查「长度/容积」类尺寸：dB、Hz、%、℃、MPa 是性能量纲，写进权项属正常限定，
        # 不是把保护范围钉在一点上（曾把权16 的「−3 dB 带宽」误报成固定尺寸）。
        UNIT = r"(?:mm|cm|μm|um|mL)"
        SIZE = re.compile(
            r"(?:φ\s*\d+(?:\.\d+)?(?:[^\s，。；、]{0,4})?"
            r"|\bDN\s*\d+"
            r"|(?<![\d.])\d+(?:\.\d+)?\s*" + UNIT + r")")
        OK_PRE = ("不大于", "不小于", "不超过", "不低于", "不多于", "不少于", "至少",
                  "小于", "大于", "≤", "≥", "以内", "以上", "以下", "介于", "范围",
                  "至", "～", "~", "－", "±")   # ±1% / ±0.1 mm 是公差带，不是固定值
        TOL = re.compile(r"(?<![A-Za-z])(?:JS|js|IT|H|h)\s*\d{1,2}(?![0-9])")
        fixed, tolh = [], []
        for num, i in claims:
            t = claim_text(num)
            for m in SIZE.finditer(t):
                pre = t[max(0, m.start() - 12):m.start()]
                post = t[m.end():m.end() + 2]
                if any(k in pre for k in OK_PRE) or any(k in post for k in ("至", "～", "~", "－")):
                    continue
                fixed.append(f"权{num}「{t[max(0, m.start() - 8):m.end()]}」")
            for m in TOL.finditer(t):
                tolh.append(f"权{num}「{t[max(0, m.start() - 6):m.end() + 1]}」")
        if fixed:
            warn("[R1] 权项里写死具体尺寸——保护范围被钉在一点上，别人做相近尺寸就不落入；"
                 "应改成相对量（x ≤ 0.08L）或单侧边界（d 不大于 8 mm）：" + "；".join(fixed[:5]))
        if tolh:
            warn("[R2] 权项里出现配合公差代号（属制造要求，不是结构特征，应移到说明书）："
                 + "；".join(tolh[:5]))

    # ---------- V 公式尽量少（口径是「尽量」，不是一律禁止） ----------
    # 结构案的保护范围与创造性来自结构特征和意料之外的效果，不来自推导：式子写进权项＝给范围定点
    # （别人等效实现差 3% 就能争不落入），ΔL=0.82a 这类通用系数写进去更是自认公知常识。
    # 但「确实离不开式子」的情形允许保留：隐式/超越方程、积分与边值条件、算法更新式。
    # 所以分级：结构/系统案 → WARN；方法类 → 仅提示；算法类 → 不查（公式是本体）。
    # 单纯的 V ≤ V*、V/V* 这类符号关系不算公式，不报。
    if claims and not is_algo:
        fx = []
        starts = [i for _, i in claims]
        for k, (num, i0) in enumerate(claims):
            # 权项正文是「起始段 + 若干无编号续行」，公式常在续行里，必须按整项扫
            i1 = starts[k + 1] if k + 1 < len(starts) else end_claim
            for i in range(i0, min(i1, len(doc.paragraphs))):
                pobj = doc.paragraphs[i]
                for m in pobj._p.findall(f".//{M_NS}oMath"):
                    txt = "".join(t.text or "" for t in m.findall(f".//{M_NS}t"))
                    # 分式/根式结构一定是公式（V ≤ cS/(2πf0) 拼出来的纯文本里没有 "/("，
                    # 只有 m:f 分式节点，不能只靠字符判断）
                    structured = (m.find(f".//{M_NS}f") is not None
                                  or m.find(f".//{M_NS}rad") is not None)
                    if structured or any(k in txt for k in ("=", "/(", "^", "√", "arctan", "\\frac")):
                        fx.append(f"权{num}「{txt[:26]}」")
        if fx:
            if is_method:
                info("[V] 权项含公式（方法类常见，若有物理动作支撑可保留，仅提示）："
                     + "；".join(fx[:5]))
            else:
                warn("[V] 权项里出现公式——尽量换成符号关系（如 V ≤ V*），符号在权项里用一句中文就地定义，"
                     "定义式退到说明书；别人用等效实现差一点就不会落入，通用系数还易被指为公知常识。"
                     "确属无法文字化的（隐式方程、积分、算法更新式）可保留："
                     + "；".join(fx[:5]))

    # 结构案正文公式总量：8 个左右为宜，超过 10 个提示压减（不是错误）
    if not (is_method or is_algo) and n_math > 10:
        warn(f"[V] 结构案正文公式 {n_math} 个偏多——正文宜用文字讲结构关系，"
             "只留定义式（判据/性能指标各一条），推导、代入、比值、中间量改成文字或参数表；"
             "判断法：把这条式子换成一句中文，读者对结构关系的理解是否受损？不受损就删")

    # ---------- Y 孔口流量量级（介质不写清就会被反算出 10~30 倍偏差） ----------
    # 同一句里出现「φd mm 孔 / Δp kPa / Q L/min」时，先按水反算一遍：
    #   Q = Cd·(πd²/4)·√(2Δp/ρ)，Cd 取 0.6，ρ=1000 kg/m³。
    # 若写明是空气介质，则跳过（空气可压缩、按孔板公式另算，不能拿水流量比对）。
    # 反例：φ5 mm 孔、10 kPa 压差写成「通气量 90 L/min」，水流量只有约 3.2 L/min，
    # 相差 28 倍——读者一算就知道介质没交代清楚。
    if i_impl >= 0:
        _y_end = ranges["说明书"][1] + 1 if "说明书" in ranges else idx_of("说明书附图")
        if _y_end < 0:
            _y_end = len(paras)
        PAT_DP = re.compile(r"(\d+(?:\.\d+)?)\s*kPa")
        PAT_Q = re.compile(r"(\d+(?:\.\d+)?)\s*L\s*/\s*min")
        PAT_D = re.compile(r"[φΦ]\s*(\d+(?:\.\d+)?)\s*mm")
        q_bad, q_ok = [], 0
        for t in paras[i_impl:_y_end]:
            for sent in re.split(r"[。；\n]", t):
                mq = PAT_Q.search(sent)
                md = PAT_D.search(sent)
                mp = PAT_DP.search(sent)
                if not (mq and md and mp):
                    continue
                if re.search(r"空气|气相|气体", sent):
                    q_ok += 1
                    continue
                d = float(md.group(1)) * 1e-3
                dp = float(mp.group(1)) * 1e3
                q = float(mq.group(1))
                q_calc = 0.6 * 3.1415926 * d * d / 4 * (2 * dp / 1000.0) ** 0.5 * 60000
                if q_calc <= 0:
                    continue
                r = q / q_calc
                if r > 3 or r < 0.33:
                    q_bad.append(f"「{sent.strip()[:30]}…」按水反算约 {q_calc:.1f} L/min，"
                                 f"文中为 {q:g} L/min（差 {r:.0f} 倍）")
                else:
                    q_ok += 1
        if q_bad:
            warn("[Y] 孔口流量与水介质反算值不在同一量级——必须写明介质（空气/水）、"
                 "压差是表压还是绝压、温度与流量系数，否则数字无法复核：" + "；".join(q_bad[:3]))
        elif q_ok:
            info(f"[Y] 孔口流量量级与水介质反算一致或已注明气体介质（{q_ok} 处）")

    # ---------- Z 孔-腔共振必须给有效颈长，不能拿板厚当颈长 ----------
    # f = (c/2π)·√(A/(V·Leff))，A=πd²/4。若只给「板厚 t」和一个 Hz，
    # 由反算的 Leff 若远大于 t（>1.4 倍），说明端修正占了主要部分（φ5 孔约 1.7a≈0.85d 每端），
    # 必须写明「计入孔口端修正」或给实测值，否则读者复算对不上。
    if i_impl >= 0:
        _z_end = ranges["说明书"][1] + 1 if "说明书" in ranges else idx_of("说明书附图")
        if _z_end < 0:
            _z_end = len(paras)
        PAT_T = re.compile(r"板厚\s*(\d+(?:\.\d+)?)\s*mm")
        PAT_V = re.compile(r"(\d+(?:\.\d+)?)\s*L")
        PAT_F = re.compile(r"(\d+(?:\.\d+)?)\s*Hz")
        z_bad = []
        for i in range(max(i_impl, 2), _z_end):
            sent = paras[i]
            if not re.search(r"共振|反共振", sent):
                continue
            mf = PAT_F.search(sent)
            if not mf:
                continue
            # 孔径/板厚/容积常写在前面两三段（改造方式一段、机理一段、频率一段），取前 3 段窗口
            win = "\n".join(paras[max(i_impl, i - 3):i + 1])
            md, mt, mv = PAT_D.search(win), PAT_T.search(win), PAT_V.search(win)
            if not (md and mt and mv):
                continue
            if not re.search(r"连通孔|节流孔|盲板", win):
                continue
            d = float(md.group(1)) * 1e-3
            tb = float(mt.group(1)) * 1e-3
            vol = float(mv.group(1)) * 1e-3
            f = float(mf.group(1))
            if f <= 0 or vol <= 0:
                continue
            A = 3.1415926 * d * d / 4
            leff = A * 1500.0 ** 2 / (vol * (2 * 3.1415926 * f) ** 2)
            if leff > 1.4 * tb and "端修正" not in win and "有效颈长" not in win \
                    and "实测" not in win:
                z_bad.append(f"「{sent.strip()[:26]}…」按 {f:g} Hz 反算需有效颈长 "
                             f"{leff * 1e3:.1f} mm，而板厚仅 {tb * 1e3:g} mm")
        if z_bad:
            warn("[Z] 孔-腔共振频率与板厚对不上——必须给有效颈长（含孔口端修正）"
                 "或注明为实测值，否则读者复算不出该频率：" + "；".join(z_bad[:3]))

    # ---------- 输出 ----------
    print("=" * 70)
    print(f"体检对象：{path}")
    print(f"段落 {len(paras)} 段、表格 {n_tab} 张、公式 {n_math} 个")
    print("=" * 70)
    for m in INFOS:
        print("  ✓", m)
    for m in WARNS:
        print("  ⚠", m)
    for m in ERRORS:
        print("  ✗", m)
    print("-" * 70)
    if ERRORS:
        print(f"结论：✗ {len(ERRORS)} 处 ERROR、{len(WARNS)} 处 WARN —— 必须改")
        return 1
    if WARNS:
        print(f"结论：⚠ {len(WARNS)} 处 WARN，无 ERROR")
        return 2
    print("结论：✅ 全部通过")
    return 0


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        print(__doc__)
        sys.exit(3)
    sys.exit(main(args[0]))
