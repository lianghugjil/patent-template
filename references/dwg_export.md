# 附图 DXF / DWG 导出规范（永久管线）

> 目的：SVG 线稿 → DXF → DWG，交付给代理所/审查员或自己在 AutoCAD / 中望 / 浩辰里改图。
> 本文件是踩坑后的定稿参数，**照抄，不要自由发挥**。

## 0-A. 一条命令（**每篇专利的附图都必须走这一步，不是可选项**）

```bash
python scripts/svg2dwg.py 附图/*.svg --out outputs/说明书附图
```

它一次做完：解析 SVG（line/rect/polygon/polyline/path/circle/ellipse/text，含 g 变换）
→ 展开 `<pattern>` 剖面线与 `<marker>` 箭头（扫描线求交生成真实线段/三角，**不需要再跑 svg_cad_prep.py**）
→ 按线宽自动分图层 → DXF（R2013）→ ODA 转 DWG（ACAD2018）→ 反向回读比对
「实体分类计数 / 文字多重集 / 图层表」→ 全一致打印 ✅，否则退出码非 0。

| 参数 | 用途 |
|---|---|
| `--out DIR` | 输出目录（默认与 SVG 同目录），DXF 与 DWG 同名落盘 |
| `--oda PATH` | 指定 ODAFileConverter.exe；不指定则依次查 `ODA_CONVERTER` 环境变量、项目内 `_tools/ODA/`、系统安装目录、PATH |
| `--scale 0.35` | 整体缩放（1 SVG 单位默认 1 mm；800×600 的画布放大到 800×600 mm，要贴 A4 版心就缩到 0.3 左右） |
| `--dxf-only` | 只出 DXF（**只用于调试**，交付必须带 DWG） |
| `--no-verify` | 跳过回读校验（**不建议**） |

**没有 ODA 不算"跳过 DWG"的理由**——按 §1 免安装拆包装一次，一劳永逸。

实测基准（800×600 线稿）：附图1 类剖视图 346 实体 / 11 文字；曲线图 30~48 实体 / 13~23 文字，
均能 100% 往返还原。

## 0. 工具链（**转换必须用 ODA，不是 LibreDWG**）

| 环节 | 工具 | 说明 |
|---|---|---|
| SVG → DXF | `scripts/svg2dwg.py`（ezdxf） | 写 DXF R2013（AC1027），1 单位 = 1 mm |
| **DXF → DWG** | **ODA File Converter**（官方 msi） | 唯一可靠。输出 ACAD2018 (AC1032) |
| DWG 回读校验 | ODA 反向转换 + LibreDWG `dwg2dxf` | 两个独立解析器都对得上才算过 |
| 打开/编辑 | AutoCAD / 中望 ZWCAD / 浩辰 | R2018 全支持 |

### ⚠️ 绝对不要用 LibreDWG 写 DWG

LibreDWG 0.14 的 `dxf2dwg.exe` **能写出文件、也能被它自己读回来**（所以往返校验会"通过"），
但中望 ZWCAD 打开直接报：

> 无法修复此图形。图形包含不正确或已损坏的信息，无法修复。

它在 R2000 / R2004 / R2018 三个版本上都有这个问题（R2004/R2018 更严重：model space 实体全丢，
`dwg2dxf` 回读报 `ERROR: BLOCK_HEADER * first_owned_entity missing`）。**LibreDWG 只当只读校验器用，
不要用它产出交付文件。** 这是自洽但不符合 DWG 规范造成的假阳性——自己读自己能过，别人的 CAD 打不开。

## 1. ODA File Converter 便携部署（无需管理员、不改注册表）

官方下载页：`https://www.opendesign.com/guestfiles/oda_file_converter`
Windows 包：`ODAFileConverter_QT6_vc16_amd64dll.msi`（约 33 MB）

`msiexec /a` 在某些沙箱/权限环境下会失败（返回 103），改用**直接抽 CAB**：

```python
# 1) 从 MSI 的 OLE 流里找出内嵌 cabinet（以 MSCF 开头）
import olefile
o = olefile.OleFileIO("oda.msi")
for e in o.listdir():
    if o.openstream(e).read(4) == b"MSCF":
        open("oda.cab", "wb").write(o.openstream(e).read())
```

```bash
# 2) 展开 CAB（文件名是 MSI 的 File 键，形如 _01606585...）
expand.exe -F:* oda.cab  ODA_raw
```

```python
# 3) 用 MSI 的 File 表把键还原成真实文件名（需要 Python ≤3.12，3.13 已删 msilib）
import msilib, os, glob, shutil
db = msilib.OpenDatabase("oda.msi", msilib.MSIDBOPEN_READONLY)
v = db.OpenView("select `File`, `FileName` from `File`"); v.Execute(None)
m = {}
while True:
    r = v.Fetch()
    if r is None: break
    m[r.GetString(1)] = r.GetString(2).split("|")[-1]   # 短名|长名
for f in glob.glob("ODA_raw/*"):
    b = os.path.basename(f)
    shutil.move(f, os.path.join("ODA", m.get(b) or m.get(b.lstrip("_")) or b))
```

```bash
# 4) Qt 平台插件必须放子目录，否则起不来
mkdir -p ODA/platforms && cp ODA/qwindows.dll ODA/platforms/
```

展开后共 57 个文件，含 `ODAFileConverter.exe`，放任意工具目录即可（免安装）。

## 2. 命令行转换

```bash
ODAFileConverter.exe <输入目录> <输出目录> <目标版本> <输出类型> <递归0/1> <审核0/1> [过滤器]
```

```bash
# DXF → DWG（R2018），开启审核修错
ODAFileConverter.exe "…\_dwgfix\in" "…\_dwgfix\out" "ACAD2018" "DWG" "0" "1" "*.dxf"
# DWG → DXF 回读（校验用）
ODAFileConverter.exe "…\_dwgfix\out" "…\_dwgfix\back" "ACAD2018" "DXF" "0" "0" "*.dwg"
```

- 版本代号：`ACAD12` `ACAD14` `ACAD2000` `ACAD2004` `ACAD2007` `ACAD2010` `ACAD2013` `ACAD2018`。
- **只吃目录，不吃单文件**；输入输出不能同目录。
- 目标版本选 `ACAD2018`：AC1032 内部是 UTF-8，**中文不需要任何转义**，直接写字面量即可。
  （旧路线用 R2000 + `\U+XXXX` 转义，现在只有在必须兼容 AutoCAD 2004 以前的老机器时才考虑。）

## 3. 图层划分（固定五层）

| 图层 | 内容 |
|---|---|
| `OUTLINE` | 零件轮廓线 |
| `HATCH` | 45° 剖面线 |
| `CENTER` | 中心线、点划线 |
| `LEADER` | 引线 + 箭头 |
| `LABEL` | 标号数字 |

全部黑色（ACI 7），不设彩色；不写 `FRAME`/图框/白底 `SOLID` 矩形——专利附图不要图框。

## 4. 文字

- 中文文字样式用 `GB`：`font = gbenor.shx`，`bigfont = gbcbig.shx`（国标长仿宋大字体）。
- 字高：画布宽 800 时 `height = 22`（与正文 12 pt 视觉相当），坐标轴文字同号。
- 图表类附图允许中文坐标轴标签与图例（数据图惯例）；**结构图仍然只允许数字标号**。
- 注意：用 ezdxf 的 matplotlib 后端本地渲染预览时，SHX 大字体解析不了，中文会显示成方块或
  `\U+XXXX` 字面量——**这是渲染器限制，不是图纸错误**，以 CAD 内实际显示为准。

## 5. 交付前必做校验（三重 + 双解析器）

```
DXF --ODA--> DWG --ODA--> DXF'     比对 实体计数 / 文字 / 图层
              DWG --LibreDWG dwg2dxf--> DXF''  再比对一次（独立解析器）
```

验收线：两项全一致、且 LibreDWG 读取无 `ERROR` 行。参考实现 `_dwgfix` 流程（项目根目录留痕）。

## 6. 交付物组织

```
outputs/说明书附图/
├── 附图1_xxx.dwg   ← DWG + 同名 DXF 成对给
├── 附图1_xxx.dxf
└── ...
```
- 同名成对：DWG 供直接改图，DXF 供版本兼容的二次编辑。
- 打包 zip 一起给，文件名带日期。
