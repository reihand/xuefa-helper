# -*- coding: utf-8 -*-
"""
「题库采集器.exe」单文件打包器。

和 build.py（速查助手）分开：采集器需要联网与 requests，
速查助手保持纯标准库、纯离线。两个 exe 放同一目录即可闭环：
  双击 题库采集器.exe  → 同目录生成 questions.json（外置题库）
  双击 学法用法速查助手.exe → 自动优先加载同目录 questions.json

用法：python build_fetcher.py
"""
import argparse
import os
import pathlib
import shutil
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parent
NAME = "题库采集器"


def candidates():
    out = [sys.executable]
    for name in ("python", "python3", "py"):
        p = shutil.which(name)
        if p:
            out.append(p)
    la = os.environ.get("LOCALAPPDATA", "")
    for v in ("Python313", "Python312", "Python311", "Python310"):
        if la:
            out.append(str(pathlib.Path(la) / "Programs" / "Python" / v / "python.exe"))
        out.append(str(pathlib.Path(r"C:\Program Files") / v / "python.exe"))
    seen, uniq = set(), []
    for p in out:
        key = str(p).lower()
        if key not in seen:
            seen.add(key)
            uniq.append(str(p))
    return uniq


def can(py, module):
    try:
        return subprocess.run([py, "-c", "import " + module],
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                              timeout=25).returncode == 0
    except (OSError, subprocess.SubprocessError):
        return False


def pick_python():
    for py in candidates():
        if can(py, "requests") and can(py, "PyInstaller"):
            return py
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--name", default=NAME)
    ap.add_argument("--no-clean", action="store_true")
    args = ap.parse_args()

    print("[1/4] 挑选带 requests + PyInstaller 的 Python")
    py = pick_python()
    if not py:
        print("  未找到，先执行：python -m pip install requests pyinstaller")
        for c in candidates():
            print("    candidate:", c, "| requests", can(c, "requests"), "| pyinstaller", can(c, "PyInstaller"))
        return 1
    print("  using:", py)

    print("[2/4] 采集器源码自检")
    r = subprocess.run([py, str(ROOT / "tools" / "fetch_practice_bank.py"), "--help"],
                       stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True)
    if r.returncode != 0:
        print("  源码导入失败：", r.stderr[-800:])
        return 1

    print("[3/4] 构建 onefile（保留控制台以显示采集进度）")
    cmd = [py, "-m", "PyInstaller", "--noconfirm"]
    if not args.no_clean:
        cmd.append("--clean")
    cmd += [
        "--onefile", "--console",
        "--name", args.name,
        "--paths", "src",
        "--add-data", "src" + os.sep + "data" + os.sep + "laws_scope.json" + ";" + "data",
        "--hidden-import", "xlsx_import",
        "--hidden-import", "requests",
        "--hidden-import", "urllib3",
        "--hidden-import", "charset_normalizer",
        "--hidden-import", "idna",
        "--hidden-import", "certifi",
    ]
    icon = ROOT / "assets" / "icon.ico"
    if icon.exists():
        cmd += ["--icon", str(icon)]
    cmd += [str(ROOT / "tools" / "fetch_practice_bank.py")]
    print("  ", " ".join(cmd))
    if subprocess.call(cmd, cwd=str(ROOT)) != 0:
        print("  BUILD FAILED")
        return 1

    print("[4/4] 验证")
    exe = ROOT / "dist" / (args.name + ".exe")
    if not exe.exists():
        print("  EXE NOT FOUND:", exe)
        return 1
    print("  output:", exe)
    print("  size  : %.2f MB" % (exe.stat().st_size / 1024 / 1024))
    print("  OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
