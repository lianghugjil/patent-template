# 专利附图绘制规范与可复用模板（专利模板 · 附图篇）

## 0. 铁律（先记这五条）

1. **图上只有数字，没有文字**。中英文、尺寸数字都不许出现在图里。
2. **纯黑白线稿**：黑线白底，不许彩色、灰度填充、阴影、渐变。
3. **剖切面必须有 45° 剖面线**（相邻零件方向相反或疏密不同）。
4. **标号必须与附图说明列表一一对应**，跨图同一部件同号。
5. **每幅图都要在「具体实施方式」里被实质引用一次**（`结合图N所示` / `如图N所示`+趋势 / `图N中…`）。
   图只放进附图册、正文不引，等于该图没被说明书说明——详见 SKILL.md §4.5，`audit_patent.py` S 类查这条。
   曲线图尤其要注意：正文写的趋势/量级必须能从图上读出来（坐标轴反算核对），图文不许打架。

真图参照（语料库导出的原始附图，本仓库不含）：
- `r09_image1.png` — 剖视图 + 单列标号 + 斜引线箭头（最简范式）
- `r12_image3.png` — 复杂剖面 + 多级标号（11/131/132/224）+ 长引线不交叉
- `r19_image3.png` — 原理图 + 场线 + 长引线（14/15/13/21）

## 1. 画布与线宽（SVG, 800×600）

| 元素 | 值 |
|---|---|
| 背景 | `<rect width="100%" height="100%" fill="#fff"/>` |
| 轮廓线 | `stroke="black" stroke-width="2.2" fill="none"` |
| 内部细线/剖面线 | `stroke="black" stroke-width="0.9"` |
| 中心线 | `stroke="black" stroke-width="0.8" stroke-dasharray="12,4,2,4"` |
| 标号数字 | `font-family="Times New Roman" font-size="24" fill="black"` |
| 引线 | `stroke="black" stroke-width="1.2"`，末端箭头用 polygon |
| 图号 | 图内不画，由 Word 图注段落输出"图1" |

## 2. 标号与引线规则

- 数字放在图形**外轮廓之外的空白区**，同一侧集中、从上到下顺序排列。
- 引线：从数字外端出发 → 指向部件表面（末端箭头），**不得交叉、不得穿过其它零件**。
- 引线优先 45° 或水平/垂直方向，长度尽量短。
- 箭头：三角形，长 10、宽 6，用 `<polygon>` 手绘（不要用 `<marker>`，部分渲染器不支持）。

箭头计算函数（Python）：

```python
def arrow(x1, y1, x2, y2, L=10, W=6):
    """从 (x1,y1) 到 (x2,y2) 的引线 + 末端箭头"""
    import math
    a = math.atan2(y2 - y1, x2 - x1)
    bx, by = x2 - L * math.cos(a), y2 - L * math.sin(a)
    px, py = -math.sin(a) * W / 2, math.cos(a) * W / 2
    return (f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{bx:.1f}" y2="{by:.1f}" '
            f'stroke="black" stroke-width="1.2"/>'
            f'<polygon points="{x2:.1f},{y2:.1f} {bx+px:.1f},{by+py:.1f} {bx-px:.1f},{by-py:.1f}" fill="black"/>')
```

## 3. 45° 剖面线生成（矩形区域，无依赖）

```python
def hatch(x0, y0, x1, y1, gap=7, reverse=False):
    """填充矩形剖切区域的 45° 剖面线；reverse=True 时方向相反"""
    lines, span = [], int((x1 - x0) + (y1 - y0)) + gap
    for c in range(gap, span, gap):
        # 正向：x + y = c；反向：x - y = c
        if not reverse:
            xa, xb = max(x0, c - y1), min(x1, c - y0)
            if xa < xb:
                lines.append(f'<line x1="{xa:.1f}" y1="{c-xa:.1f}" x2="{xb:.1f}" y2="{c-xb:.1f}" '
                             f'stroke="black" stroke-width="0.9"/>')
        else:
            xa, xb = max(x0, c + y0), min(x1, c + y1)
            if xa < xb:
                lines.append(f'<line x1="{xa:.1f}" y1="{xa-c:.1f}" x2="{xb:.1f}" y2="{xb-c:.1f}" '
                             f'stroke="black" stroke-width="0.9"/>')
    return '\n'.join(lines)
```

多边形剖切区（管壁弯曲处）用 `clipPath` 包裹线群，或在生成时按多边形逐边求交（推荐后者，兼容性最好）。

## 4. 标准件画法约定

| 零件 | 画法 |
|---|---|
| 管壁（剖切） | 内外两条轮廓线 + 中间剖面线带 |
| 法兰 | 矩形加厚 + 螺栓孔不画（简化）或画同心双线 |
| 密封圈 | 细长矩形 + 反向剖面线 |
| 螺纹连接 | 交错细线（3~4 条） |
| 传感器/电路板 | 矩形轮廓 + 内部剖面线（与壳体反向） |
| 声场/包络 | 虚线曲线，不带剖面线 |
| 流动方向 | 实线箭头（细，无标号或标号指向流体） |

## 5. 完整最小模板（一个管段 + 一个支管 + 端板）

```svg
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 800 600">
  <rect width="800" height="600" fill="#fff"/>
  <!-- 主管上下壁（剖切） -->
  <rect x="60" y="430" width="680" height="16" fill="none" stroke="black" stroke-width="2.2"/>
  <rect x="60" y="554" width="680" height="16" fill="none" stroke="black" stroke-width="2.2"/>
  {hatch_above}
  {hatch_below}
  <!-- 支管（剖切） -->
  <rect x="300" y="120" width="60" height="310" fill="none" stroke="black" stroke-width="2.2"/>
  {hatch_tube}
  <!-- 端板 -->
  <rect x="290" y="100" width="80" height="20" fill="none" stroke="black" stroke-width="2.2"/>
  {hatch_plate}
  <!-- 中心线 -->
  <line x1="330" y1="80" x2="330" y2="590" stroke="black" stroke-width="0.8" stroke-dasharray="12,4,2,4"/>
  <!-- 标号 + 引线 -->
  <text x="380" y="300" font-family="Times New Roman" font-size="24" fill="black">1</text>
  {arrow_1}
  ...
</svg>
```

注意：模板中的 `{...}` 由 Python 拼装，不直接手写；生成后用正则扫 `<text>` 是否含中文。

## 6. 自检脚本（每次画完必跑）

```python
import re
def check_svg(svg: str, expected_labels: list[int]):
    errs = []
    texts = re.findall(r'<text[^>]*>([^<]*)</text>', svg)
    for t in texts:
        if re.search(r'[\u4e00-\u9fa5A-Za-z]', t):
            errs.append(f'图上出现文字: {t!r}')
    nums = {int(t) for t in texts if t.strip().isdigit()}
    missing = set(expected_labels) - nums
    extra = nums - set(expected_labels)
    if missing: errs.append(f'缺标号: {sorted(missing)}')
    if extra: errs.append(f'多余标号: {sorted(extra)}')
    return errs
```
