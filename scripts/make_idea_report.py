#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
内部构思底稿：骨架生成 + 完整性自检 + 对外交底抽取（纯标准库，零依赖）

**这份文件从出生就是内部件**：不外发、不进交付包、不贴进对话正文。
给用户的最终交付只有五册（摘要/摘要附图/权利要求书/说明书/说明书附图）。

用法:
    python make_idea_report.py --problem "一句话技术问题" -o _内部_构思报告.md
    python make_idea_report.py --problem "..." --prior "现有方案" --assume "管径200~600mm" -o 报告.md
    python make_idea_report.py --check _内部_构思报告.md
    python make_idea_report.py --export _内部_构思报告.md --pick 3   # 抽出纯技术交底，供写稿

--check 的判定口径见 references/innovation_design.md §11：
    未检索却写"具有新颖性/具备创造性/保证授权/零侵权" → ERROR
    缺检索状态、缺章节                                   → ERROR
    方案族 <3、缺低风险/高创造性、残留【待填】             → WARN

--export 只抽「技术问题 / 原理 / 结构或步骤 / 预期效果」四段，
剥掉方案类型、实施难度、矛盾表述、检索结论；产出前再做一次泄露扫描，命中即拒绝落盘。

退出码：ERROR 存在为 1，否则为 0。
"""
from __future__ import annotations

import argparse
import datetime
import os
import re
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# ── 十个必含小节（关键词 → 人话名） ──────────────────────────────
SECTIONS = [
    ("输入与假设", "0. 输入与假设"),
    ("技术系统", "1. 技术系统与 S 曲线定位"),
    ("进化树", "2. 进化树"),
    ("矛盾", "3. 关键矛盾"),
    ("TRIZ", "4. TRIZ 工具与推演"),
    ("方案族", "5. 候选方案族"),
    ("检索校正", "6. 检索校正"),
    ("可专利性", "7. 可专利性初判"),
    ("权项梯度", "8. 权项梯度建议"),
    ("结论", "9. 结论与下一步"),
]

# 越界承诺（未检索时一律 ERROR；带否定词的整行豁免）
BANNED = [
    "具有新颖性", "具备创造性", "保证授权", "确保授权", "一定能授权",
    "可以授权", "零侵权", "包授权", "无侵权风险",
]
NEG = ["不", "禁", "无", "未", "非", "勿", "莫"]

TEMPLATE = """# $title

> **检索状态：$status**
> 生成日期：$date
> 依据：references/innovation_design.md（进化树 + TRIZ + 法 22 条三性）

---

## 0. 输入与假设

- **技术问题**：$problem
- **现有方案**：$prior
- **约束与假设**：$assume
- **输出目标**：【待填】发明 / 实用新型；结构 / 方法 / 算法 / 系统

> 缺的信息写成显式假设，不要默默编。假设一旦写下，后面所有推演都受它约束。

## 1. 技术系统与 S 曲线定位

- 功能句（动词 + 对象 + 约束）：【待填】在【约束】条件下，对【对象】实现【功能】。
- 子系统拆分（承担什么功能）：【待填】
- S 曲线定位：【待填】婴儿期 / 成长期 / 成熟期 / 衰退期
- 判定依据：【待填】同族申请量走势 / 关键参数是否逼近上限 / 是否已有替代原理

## 2. 进化树

| 节点 | 方案 | 关键参数 | 解决的矛盾 | 遗留缺陷 | 来源 |
|---|---|---|---|---|---|
| A | 【待填】 | | | | |
| B | 【待填】 | | | | |

- **主干判断**（多数方案沿哪条走）：【待填】
- **断头分支**（走到一半没人继续）：【待填】← 空白区优先在这里找
- 检索不可用时，来源列写"据已知材料（未核实）"。

## 3. 关键矛盾

- **技术矛盾（TC）**：改善【待填】，恶化【待填】。
- **物理矛盾（PC）**：同一参数【待填】既要【待填】又要【待填】。
- 优先处理的矛盾 + 理由：【待填】

## 4. TRIZ 工具与推演

- 主工具：【待填】进化法则 / 矛盾矩阵+40原理 / 物场+76标准解 / 剪裁 / IFR / 效应库 / ARIZ / SIT
- 辅工具：【待填】

| 进化法则（八条过一遍） | 是否适用 | 推出的方向 | 对应矛盾 | 预期效果 |
|---|---|---|---|---|
| 完备性 | | | | |
| 能量传递 | | | | |
| 协调性 | | | | |
| 理想化 | | | | |
| 动态化 | | | | |
| 向微观级 | | | | |
| 向超系统 | | | | |
| 子系统不均衡 | | | | |

> 写不出预期效果的方向，删掉。

## 5. 候选方案族

（≥3 个；**必须含 1 个低风险 + 1 个高创造性**）

### 方案 1：【待填】

- 类型：**低风险**（贴近现有技术，易授权）
- 原理：【待填】
- 结构或步骤：【待填】
- 解决的矛盾：【待填】
- 预期效果：【待填】
- 实施难度：【待填】低 / 中 / 高

### 方案 2：【待填】

- 类型：**高创造性**（有被质疑风险，需备"预料不到的效果"）
- 原理：【待填】
- 结构或步骤：【待填】
- 解决的矛盾：【待填】
- 预期效果：【待填】
- 实施难度：【待填】

### 方案 3：【待填】

- 类型：【待填】
- 原理：【待填】
- 结构或步骤：【待填】
- 解决的矛盾：【待填】
- 预期效果：【待填】
- 实施难度：【待填】

## 6. 检索校正

- **检索式**（关键词 / 分类号 / 引证）：
  - 关键词：【待填】中英双语 + 同义 + 上位 + 下位
  - 分类号：【待填】IPC/CPC 主组
  - 引证滚雪球：【待填】
- **命中情况**：

| 方案 | 是否命中相同/相近方案 | 处理 |
|---|---|---|
| 方案 1 | 【待填】 | 保留 / 降级为从权 / 换方向 |
| 方案 2 | | |

> 命中相同 → 降级；命中相近 → 写清区别特征；未命中 → 标"待验证"，**不等于**"具有新颖性"。

## 7. 可专利性初判

### 7.1 三性

| 项 | 判断 | 依据 |
|---|---|---|
| 新颖性（法22.2） | 【待填】 | 是否有一篇公开全部特征 |
| 创造性（法22.3） | 【待填】 | 三步法结论 |
| 实用性（法22.4） | 【待填】 | 能制造 / 能使用 / 有积极效果 |

### 7.2 创造性三步法

1. 最接近的现有技术：【待填】
2. 区别技术特征 + 实际解决的技术问题：【待填】
3. 是否显而易见：【待填】

**创造性弹药（成稿时写进有益效果）**：
- 预料不到的技术效果：【待填】
- 非线性跃变：【待填】
- 克服技术偏见：【待填】
- 多重效果协同：【待填】

### 7.3 综合打分

| 方案 | 新颖性 | 创造性 | 实用性 | 综合风险 | 依据 |
|---|---|---|---|---|---|
| 方案 1 | 【待填】 | | | | |
| 方案 2 | | | | | |

> 未检索时，新颖性/创造性只能写"待验证"。

## 8. 权项梯度建议

```
独权 1：核心（只写必要技术特征，最宽）
  └ 从权 2：加限定 A（第一级退守）
      └ 从权 3：加限定 B（第二级退守）
          └ 从权 4：加限定 C（第三级退守）
```

- 独权核心特征：【待填】
- 三级退守分别加什么：【待填】
- 尺寸写法：只划线不定点（相对量 / 单侧边界 / 区间+优选），禁公称尺寸+公差代号。

## 9. 结论与下一步

- 推荐方案 + 理由：【待填】
- 风险等级：**高 / 中 / 低**（附依据）
- 还需要补什么：检索 / 实验数据 / 结构细化 / 成本核算
- **本报告不构成授权保证，也不构成零侵权结论**；FTO 需另行按目标市场、有效期、权利要求比对开展。

---
<!-- 生成: make_idea_report.py | 校验: --check -->
"""


def read_text(path: str) -> str:
    """容错读取：utf-8-sig → utf-8 → gb18030。"""
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def gen(a) -> int:
    status = a.status
    text = TEMPLATE
    repl = {
        "$title": a.title,
        "$status": status,
        "$date": datetime.date.today().isoformat(),
        "$problem": a.problem or "【待填】",
        "$prior": a.prior or "【待填】",
        "$assume": a.assume or "【待填】",
    }
    for k, v in repl.items():
        text = text.replace(k, v)

    out = a.out or "_内部_构思报告.md"
    d = os.path.dirname(os.path.abspath(out))
    if d and not os.path.isdir(d):
        os.makedirs(d, exist_ok=True)
    with open(out, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)

    print(f"[构思报告] 已生成 {out}")
    print("  ⚠ 内部件：不外发、不进交付包、不贴进对话正文；给用户的只有五册。")
    print(f"  检索状态：{status}")
    print(f"  待填处：{text.count('【待填】')} 处")
    print("")
    print("下一步：按 references/innovation_design.md 逐节填，填完跑")
    print(f"  python make_idea_report.py --check {out}")
    if "未检索" in status:
        print("")
        print("⚠ 当前为「未检索」：全篇不得写「具有新颖性/具备创造性/保证授权/零侵权」，")
        print("  只能写「未发现相同方案，待检索验证」。")
    return 0


def check(path: str) -> int:
    if not os.path.exists(path):
        print(f"[ERROR] 文件不存在：{path}")
        return 1

    text = read_text(path)
    lines = text.splitlines()
    errors, warns, oks = [], [], []

    # 1) 检索状态
    m = re.search(r"检索状态[:：]\s*\**\s*(已检索|未检索)", text)
    status = None
    if not m:
        errors.append("未填写「检索状态：已检索 / 未检索」（报告首行必须有）")
    else:
        status = m.group(1)
        if status == "未检索":
            warns.append("检索状态：未检索——新颖性/创造性结论未经验证；本稿只能作「待检索稿」，"
                         "递交前必须闭环，且不得对外宣称具备新颖性或创造性")
        else:
            oks.append("检索状态：已检索")
            if not re.search(r"检索式|关键词.*分类号|IPC|CPC", text):
                warns.append("声明已检索但没写检索式（关键词/分类号），补上才算可复核")

    # 2) 越界承诺
    hit = 0
    for i, line in enumerate(lines, 1):
        if any(b in line for b in BANNED):
            if any(n in line for n in NEG):
                continue  # 否定句豁免（如"不承诺零侵权"）
            hit += 1
            errors.append(f"L{i} 越界承诺「{line.strip()[:36]}」——未授权前不得这样写")
    if hit == 0:
        oks.append("无越界承诺（新颖性/创造性/授权/侵权）")

    # 3) 章节
    missing = [name for kw, name in SECTIONS if kw not in text]
    if missing:
        errors.append("缺小节：" + "、".join(missing))
    else:
        oks.append("十节齐全")

    # 4) 方案族
    plans = re.findall(r"^#{2,4}\s*方案\s*\d+", text, re.M)
    n = len(plans)
    if n < 3:
        warns.append(f"方案族只有 {n} 个（应 ≥3：一检索命中就得换方案）")
    else:
        oks.append(f"方案族 {n} 个")
    if "低风险" not in text:
        warns.append("缺「低风险」方案（贴近现有技术、易授权，建议至少 1 个）")
    if "高创造性" not in text:
        warns.append("缺「高创造性」方案（价值大，建议至少 1 个）")

    # 5) 权项梯度
    tiers = len(re.findall(r"从权", text))
    if tiers < 3:
        warns.append(f"权项退守只写了 {tiers} 级（建议 ≥3 级，否则答复期无路可退）")
    else:
        oks.append(f"权项梯度 {tiers} 级")

    # 6) 三性表
    if not all(k in text for k in ("新颖性", "创造性", "实用性")):
        warns.append("三性（新颖性/创造性/实用性）未写全")
    else:
        oks.append("三性判断齐全")
        if (re.search(r"新颖性[^\n]{0,20}待验证", text)
                and re.search(r"创造性[^\n]{0,20}待验证", text)):
            warns.append("新颖性与创造性均为「待验证」：底稿只到「待检索稿」，"
                         "写五册可以，交付时必须向委托人明示创造性结论未验证")

    # 7) 残留占位
    left = text.count("【待填】")
    if left:
        warns.append(f"仍有 {left} 处【待填】未填")

    # ── 输出 ──
    print(f"[构思报告自检] {os.path.basename(path)}")
    print(f"  ERROR {len(errors)} / WARN {len(warns)}")
    for s in oks:
        print(f"  ✅ {s}")
    for s in errors:
        print(f"  ❌ {s}")
    for s in warns:
        print(f"  ⚠ {s}")
    print("")
    if errors:
        print("  结论：ERROR 必须清零后再进成稿（否则构思阶段就埋了硬伤）。")
        return 1
    if warns:
        print("  结论：无 ERROR，但 WARN 未清。逐条判断后复跑；"
              "其中「检索状态：未检索」必须在递交前闭环——未检索的稿子只能算待检索稿。")
        return 2
    print("  结论：✅ 无 ERROR 无 WARN。")
    return 0


def _plain(s: str) -> str:
    """把底稿里的 markdown 强调与内部标注洗掉，只留能进正文的句子。"""
    s = re.sub(r"\*\*(.+?)\*\*", r"\1", s)
    s = re.sub(r"__(.+?)__", r"\1", s)
    s = re.sub(r"`(.+?)`", r"\1", s)
    s = re.sub(r"[（(]\s*(推荐|主推|首选|备选|高创造性|低风险|中高|待验证)\s*[）)]", "", s)
    return re.sub(r"\s{2,}", " ", s).strip()


def _leak_hits(text: str):
    """对「对外稿」做泄露扫描；与 audit_lite 的 L 类同源，缺模块时用内置最简表。"""
    try:
        from audit_lite import LEAK_METHOD, LEAK_DRAFT, LEAK_TRADEOFF, LEAK_PATTERNS
    except Exception:  # 单独拷走脚本的极端情况
        LEAK_METHOD = ["TRIZ", "进化树", "矛盾矩阵", "物场", "ARIZ", "方案族", "可专利性", "检索"]
        LEAK_DRAFT = ["本报告", "待填", "【", "建议申请"]
        LEAK_TRADEOFF = ["权衡", "取舍", "择优", "备选"]
        LEAK_PATTERNS = [
            (r"方案\s*[一二三四五六七八九十ABCDE1-9]\s*[：:）)、，,]", "方案编号"),
            (r"第\s*[一二三四五六七八九十\d]+\s*个?方案", "方案序号"),
        ]
    hits = []
    for w in LEAK_METHOD + LEAK_DRAFT + LEAK_TRADEOFF:
        if w in text:
            hits.append(w)
    for pat, name in LEAK_PATTERNS:
        if re.search(pat, text):
            hits.append(name)
    return hits


def export(path: str, pick: int, out: str) -> int:
    """从内部报告抽出「只讲技术、不讲怎么想出来」的一页交底。

    这个文件仍然只是写稿的输入件，不是交付物；最终给用户的只有五册。
    输出前会做一次泄露扫描，命中即拒绝落盘（宁可不给，也不交带方法论的稿）。
    """
    if not os.path.exists(path):
        print(f"[ERROR] 文件不存在：{path}")
        return 1
    text = read_text(path)

    m = re.search(r"-\s*\*\*技术问题\*\*[：:]\s*(.+)", text)
    problem = (m.group(1).strip() if m else "")

    # 检索状态会一路影响后面的成稿口径：未检索的稿子只是待检索稿，
    # 必须把这一事实钉在交底件上，避免写五册的人以为创造性已经验证过。
    ms = re.search(r"检索状态[:：]\s*\**\s*(已检索|未检索)", text)
    unsearched = bool(ms and ms.group(1) == "未检索")

    blocks = re.split(r"^###\s*方案\s*(\d+)\s*[：:]\s*", text, flags=re.M)
    cards = {}
    for i in range(1, len(blocks) - 1, 2):
        cards[int(blocks[i])] = blocks[i + 1]
    if pick not in cards:
        print(f"[ERROR] 报告里没有「方案 {pick}」（现有：{'、'.join(map(str, sorted(cards))) or '无'}）")
        return 1

    body = cards[pick]
    title = _plain(body.splitlines()[0].strip())

    def field(name):
        # 容忍「- **预期效果**：xxx」这种加粗写法
        mm = re.search(r"-\s*\**\s*" + name + r"\s*\**\s*[：:]\s*(.+)", body)
        return _plain(mm.group(1)) if mm else ""

    principle, structure, effect = field("原理"), field("结构或步骤"), field("预期效果")
    problem = _plain(problem)
    for k, v in (("原理", principle), ("结构或步骤", structure), ("预期效果", effect)):
        if not v or "【待填】" in v:
            print(f"[ERROR] 方案 {pick} 的「{k}」还没填，先把技术内容写实再导出")
            return 1

    doc = []
    doc.append("# 技术方案说明")
    doc.append("")
    doc.append("## 一、要解决的技术问题")
    doc.append("")
    doc.append(problem or "（待补）")
    doc.append("")
    doc.append("## 二、技术方案")
    doc.append("")
    doc.append(f"技术方案为{title}。{principle}")
    doc.append("")
    doc.append(structure)
    doc.append("")
    doc.append("## 三、有益效果")
    doc.append("")
    doc.append(effect)
    doc.append("")
    body_out = "\n".join(doc)

    hits = _leak_hits(body_out)
    if hits:
        print("[拒绝导出] 抽出内容里仍有过程性表述：" + "、".join(sorted(set(hits))))
        print("  两条路：① 把这些措辞挪到「解决的矛盾 / 类型」字段（导出时会整段剥掉）；")
        print("           ② 直接改写成技术描述（「满足条件分离」→「对两种材料的去除量差两个数量级」）。")
        print("  命中的词留在原理/结构或步骤里，就会一路带进说明书。")
        return 1

    if not out:
        out = os.path.splitext(path)[0] + "_交底.md"
    with open(out, "w", encoding="utf-8", newline="\n") as f:
        f.write(body_out)
        if unsearched:
            f.write("\n\n<!-- 底稿的检索状态为：未检索。新颖性与创造性结论均未经验证；"
                    "本页只作写稿输入，不得据此对外宣称具备新颖性或创造性，"
                    "递交前须补齐检索并把结论回填底稿。 -->\n")

    print(f"[对外交底] {out}")
    print(f"  取自「方案 {pick}：{title}」，已剥离类型标注、实施难度、矛盾表述与检索结论。")
    if unsearched:
        print("  ⚠ 底稿声明未检索：本页文件尾已标注，写五册时可写，但交付口径必须写明"
              "「创造性结论未验证」。")
    print("  写稿时还要做两件事：有益效果改「第一…第N」分层（删掉「最硬/最贵/最关键」这类自评），")
    print("  技术方案改陈述句（不写「本方案」「推荐」）。")
    print("  提醒：这份也只是写稿输入件；给用户的最终交付只有五册。")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(
        description="创新构思报告：骨架生成 + 完整性自检（纯标准库）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("--problem", default="", help="一句话技术问题")
    ap.add_argument("--prior", default="", help="现有方案（哪怕一句）")
    ap.add_argument("--assume", default="", help="约束与假设，如「管径200~600mm，不能停水」")
    ap.add_argument("--title", default="创新构思报告", help="报告标题")
    ap.add_argument(
        "--status", default="未检索，仅供讨论",
        choices=["未检索，仅供讨论", "已检索"],
        help="检索状态（默认未检索，宁保守不夸大）",
    )
    ap.add_argument("-o", "--out", default="", help="输出路径（默认 _内部_构思报告.md；这是内部件）")
    ap.add_argument("--check", metavar="FILE", default="", help="自检一份已填写的报告")
    ap.add_argument("--export", metavar="FILE", default="", help="从报告抽出对外交底（只含技术内容）")
    ap.add_argument("--pick", type=int, default=0, help="配合 --export：选定方案编号")
    a = ap.parse_args()

    if a.check:
        return check(a.check)
    if a.export:
        if not a.pick:
            print("[ERROR] --export 需要配合 --pick N 指定选定方案编号")
            return 1
        return export(a.export, a.pick, a.out)
    return gen(a)


if __name__ == "__main__":
    sys.exit(main())
