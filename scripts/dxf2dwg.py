#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dxf2dwg.py —— 把 DXF 批量转成 DWG（只用标准库，找不到转换器也不报错）。

为什么单列一个脚本：DXF 是纯文本，标准库就能写（fig_lite.py 直接出）；
DWG 是二进制专有格式，自己写不了，只能借现成转换器。本脚本自动去找，
找到就转、找不到就说一句「DXF 打开后另存即得 DWG」，退出码仍是 0，
不会把整条交付流程卡死。

三条通道，按顺序自动试：
  1. ODA File Converter —— 免费、离线、一次转一整个目录（推荐）
  2. 本机 CAD 的 COM 接口 —— 中望 ZWCAD / 浩辰 GstarCAD / AutoCAD / BricsCAD
  3. 都没有 —— 打印提示，退出码 0

用法：
    python dxf2dwg.py 说明书附图/
    python dxf2dwg.py 说明书附图/ --out 说明书附图/CAD
    python dxf2dwg.py 附图1.dxf 附图2.dxf --ver ACAD12
    python dxf2dwg.py 附图/ --oda "D:/ODA/ODAFileConverter.exe"

版本名与 AutoCAD 版本对应：ACAD9=R9 … ACAD12=R12、ACAD14=R14、
ACAD2000=2000、ACAD2004=2004、ACAD2007=2007、ACAD2010=2010、
ACAD2013=2013、ACAD2018=2018（默认，兼容性最好）。
"""

from __future__ import print_function

import glob
import os
import shutil
import subprocess
import sys
import tempfile

VER_OK = ("ACAD9", "ACAD10", "ACAD12", "ACAD13", "ACAD14", "ACAD2000",
          "ACAD2004", "ACAD2007", "ACAD2010", "ACAD2013", "ACAD2018")

CAD_PROGIDS = ("ZWCAD.Application",        # 中望 CAD
               "GstarCAD.Application",     # 浩辰 CAD
               "AutoCAD.Application",      # AutoCAD
               "BricscadApp.AcadApplication")


# --------------------------------------------------------------------------
# 找转换器
# --------------------------------------------------------------------------
def find_oda(explicit=None):
    """返回可用的 ODAFileConverter 路径，找不到返回 None。"""
    cands = []
    if explicit:
        cands.append(explicit)
    for var in ("ODA_CONVERTER", "ODA_FILE_CONVERTER"):   # 与 selfcheck.py 保持同一套变量名
        env = os.environ.get(var)
        if env:
            cands.append(env)
    for name in ("ODAFileConverter.exe", "ODAFileConverter"):
        w = shutil.which(name) if hasattr(shutil, "which") else None
        if w:
            cands.append(w)
    here = os.path.dirname(os.path.abspath(__file__))
    up = here
    for _ in range(3):
        cands.append(os.path.join(up, "_tools", "ODA", "ODAFileConverter.exe"))
        cands.append(os.path.join(up, "ODA", "ODAFileConverter.exe"))
        up = os.path.dirname(up)
    if os.name == "nt":
        roots = [os.environ.get("ProgramFiles"),
                 os.environ.get("ProgramFiles(x86)"),
                 os.environ.get("ProgramW6432"),
                 os.environ.get("LOCALAPPDATA")]
        for base in roots:
            if not base or not os.path.isdir(base):
                continue
            for pat in ("ODA/*/ODAFileConverter.exe",
                        "ODA/ODAFileConverter.exe",
                        "*/ODAFileConverter.exe",
                        "*/ODAFileConverter/*/ODAFileConverter.exe"):
                cands.extend(glob.glob(os.path.join(base, pat)))
        cands.append(r"C:\Program Files\ODA\ODAFileConverter 25.4.0\ODAFileConverter.exe")
    for c in cands:
        if c and os.path.isfile(c):
            return os.path.abspath(c)
    return None


def _ps(script, timeout=300):
    """跑一段 PowerShell（EncodedCommand 避免中文路径被编码打坏）。"""
    if os.name != "nt":
        return None
    try:
        import base64
        enc = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
        p = subprocess.Popen(
            ["powershell", "-NoProfile", "-NonInteractive",
             "-ExecutionPolicy", "Bypass", "-EncodedCommand", enc],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        out, _ = p.communicate(timeout=timeout)
        return (out or b"").decode("utf-8", "replace")
    except Exception:
        return None


def find_cad():
    """探测本机装了哪个 CAD（COM 可用的），返回 ProgID 或 None。"""
    if os.name != "nt":
        return None
    script = ("$ids=@(%s);"
              "foreach($i in $ids){$t=[Type]::GetTypeFromProgID($i);"
              "if($t -ne $null){Write-Output ('FOUND '+$i)}}"
              % ",".join("'%s'" % p for p in CAD_PROGIDS))
    out = _ps(script, timeout=60) or ""
    for line in out.splitlines():
        line = line.strip()
        if line.startswith("FOUND "):
            return line[6:].strip()
    return None


# --------------------------------------------------------------------------
# 转换
# --------------------------------------------------------------------------
def _oda_convert(oda_exe, src_dir, dst_dir, ver, timeout=600):
    cmd = [oda_exe, src_dir, dst_dir, ver, "DWG", "0", "1", "*.dxf"]
    try:
        p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                             cwd=os.path.dirname(oda_exe))
        p.communicate(timeout=timeout)
        return True
    except Exception:
        return False


def _cad_convert(progid, pairs, outdir):
    """用 CAD 的 COM 接口逐个打开另存。pairs = [(dxf, dwg), ...]"""
    lines = ["$ErrorActionPreference='Stop'",
             "$app = New-Object -ComObject '%s'" % progid,
             "try {$app.Visible = $false} catch {}"]
    for dxf, dwg in pairs:
        d = dxf.replace("'", "''")
        g = dwg.replace("'", "''")
        lines.append("try { $doc = $app.Documents.Open('%s', $false);"
                     " $doc.SaveAs('%s'); $doc.Close($false);"
                     " Write-Output ('OK %s') } catch { Write-Output ('FAIL %s') }"
                     % (d, g, os.path.basename(g), os.path.basename(g)))
    lines.append("try {$app.Quit()} catch {}")
    out = _ps("\n".join(lines), timeout=60 + 30 * len(pairs)) or ""
    return out.count("OK ") if "OK " in out else 0, out


def collect(paths):
    files = []
    for p in paths:
        if os.path.isdir(p):
            for root, _dirs, names in os.walk(p):
                for n in sorted(names):
                    if n.lower().endswith(".dxf"):
                        files.append(os.path.join(root, n))
        elif os.path.isfile(p):
            files.append(p)
    return files


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv or argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0

    paths, outdir, oda, ver = [], None, None, "ACAD2018"
    i = 0
    while i < len(argv):
        a = argv[i]
        if a == "--out" and i + 1 < len(argv):
            outdir = argv[i + 1]
            i += 2
        elif a == "--oda" and i + 1 < len(argv):
            oda = argv[i + 1]
            i += 2
        elif a == "--ver" and i + 1 < len(argv):
            ver = argv[i + 1].upper()
            i += 2
        else:
            paths.append(a)
            i += 1

    if ver not in VER_OK:
        print("[!] 未知版本 %s，可选：%s" % (ver, "、".join(VER_OK)))
        return 2

    files = collect(paths)
    if not files:
        print("[!] 没找到 .dxf 文件：%s" % " ".join(paths))
        return 2

    # 统一在临时目录里转，避免 ODA「输入输出不能同目录」的限制
    tmp = tempfile.mkdtemp(prefix="dxf2dwg_")
    src = os.path.join(tmp, "in")
    dst = os.path.join(tmp, "out")
    os.makedirs(src)
    os.makedirs(dst)

    name2orig = {}
    for f in files:
        base = os.path.splitext(os.path.basename(f))[0]
        name, k = base, 1
        while name in name2orig:
            k += 1
            name = "%s_%d" % (base, k)
        name2orig[name] = f
        shutil.copy2(f, os.path.join(src, name + ".dxf"))

    oda_exe = find_oda(oda)
    how = ""
    if oda_exe:
        ok = _oda_convert(oda_exe, src, dst, ver)
        how = "ODA File Converter"
        if not ok:
            print("[!] ODA 调用失败（%s）" % oda_exe)
            oda_exe = None

    made = []
    if not oda_exe:
        cad = find_cad()
        if cad:
            pairs = []
            for name, orig in name2orig.items():
                tgt_dir = outdir or os.path.dirname(os.path.abspath(orig))
                if not os.path.isdir(tgt_dir):
                    os.makedirs(tgt_dir, exist_ok=True)
                pairs.append((os.path.join(src, name + ".dxf"),
                              os.path.join(tgt_dir, name + ".dwg")))
            n_ok, raw = _cad_convert(cad, pairs, outdir)
            how = cad
            for _d, g in pairs:
                if os.path.isfile(g) and os.path.getsize(g) > 0:
                    made.append(g)
        else:
            shutil.rmtree(tmp, ignore_errors=True)
            print("[i] 没找到 DWG 转换器（ODA File Converter 或本机 CAD）。")
            print("    DXF 已可用：AutoCAD / 中望 / 浩辰 / CAXA 打开后"
                  "「另存为 → DWG」一步即得，几何与图层不变。")
            print("    装了 ODA 之后重跑本命令即可自动出 DWG：")
            print("      python dxf2dwg.py <目录> --oda \"<ODAFileConverter.exe 路径>\"")
            return 0

    # ODA 通道：把转好的 dwg 拷回目标位置
    if oda_exe:
        for name, orig in name2orig.items():
            p = os.path.join(dst, name + ".dwg")
            if os.path.isfile(p) and os.path.getsize(p) > 0:
                tgt_dir = outdir or os.path.dirname(os.path.abspath(orig))
                os.makedirs(tgt_dir, exist_ok=True)
                tgt = os.path.join(tgt_dir, name + ".dwg")
                shutil.copy2(p, tgt)
                made.append(tgt)

    shutil.rmtree(tmp, ignore_errors=True)

    total = len(files)
    print("[✓] %d/%d 个 DXF → DWG（%s，%s）" % (len(made), total, how, ver))
    for m in made:
        print("    %s  %.1f KB" % (m, os.path.getsize(m) / 1024.0))
    if len(made) < total:
        miss = total - len(made)
        print("[!] %d 个没转成：源文件可能不是合法 DXF，或路径被占用" % miss)
    return 0


if __name__ == "__main__":
    sys.exit(main())
