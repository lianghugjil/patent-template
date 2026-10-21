#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""环境自检 —— 一条命令看清这台机器能跑到哪一层，别让缺依赖把交付卡死。

    python selfcheck.py            # 看能力等级与可用命令
    python selfcheck.py --quiet    # 只打印等级（脚本里判断用）

能力分级（能到哪层用哪层，**任何一层都不许中止交付**）：

    L0  只用 python3 标准库 —— fig_lite.py 出带标号附图(PNG+DXF)、make_docx.py 出五册 docx、audit_lite.py 出体检数字
    L1  有 python-docx      —— apply_patent_format.py / fix_layout.py / audit_patent.py / audit_exam_risk.py
    L0 另有 audit_figures.py（图面体检，纯标准库）
    L2  有 ODA File Converter 或本机 CAD —— dxf2dwg.py 把 DXF 批量转 DWG（DXF 在 L0 就有，这层只多 DWG）
    L3  有 Office/WPS/soffice —— docx → PDF 目视验收

    注：已经画好的 SVG 想直出 CAD，需要 ezdxf + ODA，走 svg2dwg.py（可选，不是主线）。
"""
from __future__ import annotations

import glob
import os
import shutil
import subprocess
import sys


def has_module(name):
    try:
        __import__(name)
        return True
    except Exception:
        return False


def which(cmd):
    return shutil.which(cmd) is not None


def find_oda():
    cands = []
    for var in ("ODA_CONVERTER", "ODA_FILE_CONVERTER"):
        env = os.environ.get(var)
        if env:
            cands.append(env)
    here = os.path.dirname(os.path.abspath(__file__))
    up = here
    for _ in range(3):
        cands.append(os.path.join(up, "_tools", "ODA", "ODAFileConverter.exe"))
        up = os.path.dirname(up)
    cands += [
        r"C:\Program Files\ODA\ODAFileConverter*\ODAFileConverter.exe",
        r"C:\Program Files (x86)\ODA\ODAFileConverter*\ODAFileConverter.exe",
        r"C:\Program Files\ODA\*\ODAFileConverter.exe",
    ]
    for c in cands:
        for hit in glob.glob(c):
            if os.path.exists(hit):
                return hit
    return which("ODAFileConverter") and "PATH 中的 ODAFileConverter" or ""


def find_cad():
    """探测本机可用的 CAD（COM），返回 ProgID 或空串。只在没有 ODA 时才值得问。"""
    if os.name != "nt":
        return ""
    try:
        import base64
        ids = ("ZWCAD.Application", "GstarCAD.Application",
               "AutoCAD.Application", "BricscadApp.AcadApplication")
        script = ("$ids=@(%s); foreach($i in $ids){ if([Type]::GetTypeFromProgID($i) "
                  "-ne $null){Write-Output ('FOUND '+$i)} }"
                  % ",".join("'%s'" % i for i in ids))
        enc = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
        out = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-EncodedCommand", enc],
            stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, timeout=25).stdout
        for line in (out or b"").decode("utf-8", "replace").splitlines():
            line = line.strip()
            if line.startswith("FOUND "):
                return line[6:].strip()
    except Exception:
        pass
    return ""


def main():
    quiet = "--quiet" in sys.argv
    py = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
    docx_ok = has_module("docx")
    pil_ok = has_module("PIL")
    ezdxf_ok = has_module("ezdxf")
    node_ok = which("node")
    soffice_ok = which("soffice") or which("libreoffice")
    oda = find_oda()
    cad = "" if oda else find_cad()
    dwg_ok = bool(oda or cad)
    office_ok = bool(soffice_ok) or os.name == "nt"  # Windows 上一般有 WPS/Word 可走 COM

    level = "L0"
    if docx_ok and pil_ok:
        level = "L1"
    if dwg_ok:
        level = "L2"
    if level == "L2" and office_ok:
        level = "L3"

    if quiet:
        print(level)
        return 0

    print("[环境自检]")
    print(f"  python      {py}  （{sys.executable}）")
    print(f"  python-docx {'✓' if docx_ok else '✗'}    Pillow {'✓' if pil_ok else '✗'}    "
          f"ezdxf {'✓' if ezdxf_ok else '✗'}    node {'✓' if node_ok else '✗'}")
    print(f"  PDF 导出    {'✓ ' + (soffice_ok or 'WPS/Word COM') if office_ok else '✗'}")
    if oda:
        print(f"  DWG 转换    ✓ ODA  {oda}")
    elif cad:
        print(f"  DWG 转换    ✓ 本机 CAD  {cad}")
    else:
        print("  DWG 转换    ✗ 无转换器（不影响交付：DXF 照常出，CAD 打开后另存即得 DWG）")
    print(f"\n  能力等级：{level}")
    print("  可用命令：")
    print("    python scripts/fig_lite.py 附图.fig         # L0：附图 PNG + DXF（带标号，仅标准库）")
    print("    python scripts/make_idea_report.py --problem \"一句话\"   # L0：内部构思底稿（--check 自检 / --export 抽交底）")
    print("        ⚠ 底稿是内部件，不给用户；交付物只有五册")
    print("    python scripts/make_docx.py 稿子.md         # L0：五册 docx（仅标准库）")
    print("    python scripts/audit_lite.py 稿子.md        # L0：内容体检（仅标准库）")
    print("    python scripts/audit_lite.py 成品.docx      # L0：版式体检（仅标准库）")
    print("    python scripts/audit_figures.py 附图目录 --doc 成品.docx   # L0：图面体检 G1~G10（引线相交/压字/标号表核对）")
    if docx_ok:
        print("    python scripts/apply_patent_format.py 成品.docx   # L1：套版式（按册名切节）")
        print("    python scripts/fix_layout.py 成品.docx            # L1：行距/图宽收口")
        print("    python scripts/audit_patent.py 成品.docx          # L1：A~Z 全量格式体检")
        print("    python scripts/audit_exam_risk.py 成品.docx       # L1：J1~J18 实质体检")
    else:
        print("    （缺 python-docx：L1 系列脚本不可用，改走 L0——照样能出 docx 与体检数字）")
    if dwg_ok:
        print("    python scripts/dxf2dwg.py 说明书附图/             # L2：DXF 批量转 DWG")
    else:
        print("    （无 ODA/CAD：DXF 已在 L0 出，用任何 CAD 打开另存一次就是 DWG，不挡交付）")
    print("\n  原则：到哪层用哪层，缺的项在交付说明里一句话交代，绝不因缺依赖停下不交稿。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
