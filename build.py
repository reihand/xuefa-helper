# -*- coding: utf-8 -*-
"""
单文件 exe 打包器。

自动挑一个带 tkinter 的 Python（WorkBuddy 内置 Python 通常不带 tkinter），
检查 PyInstaller，然后构建 --onefile --windowed exe 并验证产物。

用法：python build.py
     python build.py --name 自定义名称
     python build.py --console      # 带控制台窗口，便于排查启动报错
"""
import argparse
import os
import pathlib
import shutil
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parent
DEFAULT_NAME = "学法用法速查助手"


def has_tk(py):
    try:
        r = subprocess.run([py, "-c", "import tkinter"],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                           timeout=20)
        return r.returncode == 0
    except (OSError, subprocess.SubprocessError):
        return False


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


def pick_python():
    for py in candidates():
        if has_tk(py):
            return py
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--name", default=DEFAULT_NAME)
    ap.add_argument("--console", action="store_true", help="保留控制台窗口用于排错")
    ap.add_argument("--no-clean", action="store_true")
    args = ap.parse_args()

    print("[1/4] pick a Python with tkinter")
    py = pick_python()
    if not py:
        print("  FAILED: no Python with tkinter found.")
        print("  Install official Python 3.10+ (check 'tcl/tk and IDLE'), or run:")
        print("  python build.py --help")
        return 1
    ver = subprocess.check_output(
        [py, "-c", "import sys,tkinter;print(sys.version.split()[0]+' Tk'+str(tkinter.TkVersion))"],
        text=True).strip()
    print("  using:", py, "|", ver)

    print("[2/4] check PyInstaller")
    r = subprocess.run([py, "-m", "PyInstaller", "--version"],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    if r.returncode != 0:
        print("  installing pyinstaller ...")
        if subprocess.call([py, "-m", "pip", "install", "pyinstaller"]) != 0:
            print("  pip install failed")
            return 1

    print("[3/4] build onefile")
    cmd = [py, "-m", "PyInstaller", "--noconfirm"]
    if not args.no_clean:
        cmd.append("--clean")
    cmd += ["--onefile"]
    if not args.console:
        cmd.append("--windowed")
    cmd += [
        "--name", args.name,
        # 只收两个数据文件，避免把 src/data 下的 .bak 备份一起打进 exe
        "--add-data", "src" + os.sep + "data" + os.sep + "questions.json" + ";" + "data",
        "--add-data", "src" + os.sep + "data" + os.sep + "laws_scope.json" + ";" + "data",
        "--hidden-import", "tkinter",
        "--hidden-import", "tkinter.ttk",
        "--hidden-import", "tkinter.filedialog",
        "--hidden-import", "tkinter.messagebox",
        "--hidden-import", "xlsx_import",
        "--hidden-import", "clipboard",
        "--hidden-import", "paths",
    ]
    icon = ROOT / "assets" / "icon.ico"
    if icon.exists():
        cmd += ["--icon", str(icon)]
    cmd += [str(ROOT / "src" / "main.py")]
    print("  ", " ".join(cmd))
    if subprocess.call(cmd, cwd=str(ROOT)) != 0:
        print("  BUILD FAILED")
        return 1

    print("[4/4] verify")
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
