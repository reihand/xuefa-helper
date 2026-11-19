# -*- coding: utf-8 -*-
"""
打包 UOS / 统信 arm64 可用离线版。

不依赖目标架构：交付纯 Python + 浏览器界面，目标机只要有 Python 3 和浏览器即可。
产物：
  dist/学法用法速查助手_UOS_arm64.tar.gz   解压即用
  dist/学法用法速查助手_UOS_arm64.run      单文件自解压，sh 文件.run 即可
"""
import base64
import pathlib
import shutil
import tarfile

ROOT = pathlib.Path(__file__).resolve().parent
STAGE = ROOT / "build_uos" / "xuefa-uos-arm64"
DIST = ROOT / "dist"

LAUNCH = """#!/bin/sh
cd "$(dirname "$0")" || exit 1
if ! command -v python3 >/dev/null 2>&1; then
  echo "未找到 python3，请先安装 Python 3。" >&2
  exit 1
fi
exec python3 src/webui.py "$@"
"""

CLI = """#!/bin/sh
cd "$(dirname "$0")" || exit 1
exec python3 src/webui.py --query "$@"
"""

DESKTOP = """#!/bin/sh
HERE="$(cd "$(dirname "$0")" && pwd)" || exit 1
DESK="$HOME/Desktop"
if [ ! -d "$DESK" ]; then DESK="$HOME/桌面"; fi
cat > "$DESK/学法用法速查助手.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=学法用法速查助手
Comment=本地离线题库速查
Exec=sh "$HERE/启动速查.sh"
Terminal=false
Categories=Utility;
EOF
chmod +x "$DESK/学法用法速查助手.desktop"
echo "已生成：$DESK/学法用法速查助手.desktop"
"""

SELFTEST = """#!/bin/sh
cd "$(dirname "$0")" || exit 1
python3 src/webui.py --query "根据《中华人民共和国民法典》，向人民法院请求保护民事权利的诉讼时效期间为（ ）。A.一年 B.二年 C.三年 D.五年" | head -3
"""

README_UOS = """学法用法速查助手 · UOS / 统信 arm64 离线版
==========================================

这是纯 Python 版本，不区分 x86 / arm64。目标机需要：
  · Python 3.7 或更高版本（统信系统一般自带）
  · 任意浏览器（界面在浏览器里打开）
不需要安装第三方库，不联网，不连考试系统。

【开始使用】
  1. 解压本目录。
  2. 双击「启动速查.sh」，或在终端执行：
        sh 启动速查.sh
  3. 程序自动打开浏览器。没自动打开就手动访问终端显示的地址
     （默认 http://127.0.0.1:8765/）。
  4. 关闭：在终端按 Ctrl+C。

【命令行直接查】
  sh 命令行查询.sh "题干或关键词"

【生成桌面图标】
  sh 生成桌面快捷方式.sh
  之后桌面上出现「学法用法速查助手」。

【自检】
  sh 自检.sh
  预期输出第一行包含：1.000 | Q9924 | C

【补库】
  速查页 →「说明 / 补库」→ 选 Excel 导入。
  支持「广西普法考试题库录入」导出的三张工作表，
  也支持标准表头：题干 | 类型 | A | B | … | 答案 | 依据 | 法规 | 来源。
  导入前自动留 questions.json.bak；同题干去重，原有解析保留。

【文件】
  启动速查.sh                  启动浏览器界面
  命令行查询.sh                命令行查题
  生成桌面快捷方式.sh          桌面图标
  自检.sh                      快速验证
  src/webui.py                 本地服务与页面
  src/bank.py                  2-gram 倒排 + difflib 精排 + 选项一致性加权
  src/xlsx_import.py           纯标准库 xlsx 解析
  src/data/questions.json      10196 道题
  src/data/laws_scope.json     附件3 的 382 部法规清单
  questions.json               外置题库，存在时优先于内置

题库和法规清单都在本地，程序不向外发送题目。
"""


def write_lf(path, text, mode=0o644):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as f:
        f.write(text)
    path.chmod(mode)


def copy_sources():
    if STAGE.exists():
        shutil.rmtree(STAGE)
    (STAGE / "src" / "data").mkdir(parents=True)
    for name in ("webui.py", "bank.py", "paths.py", "xlsx_import.py"):
        shutil.copy2(ROOT / "src" / name, STAGE / "src" / name)
    for name in ("questions.json", "laws_scope.json"):
        shutil.copy2(ROOT / "src" / "data" / name, STAGE / "src" / "data" / name)

    write_lf(STAGE / "启动速查.sh", LAUNCH, 0o755)
    write_lf(STAGE / "命令行查询.sh", CLI, 0o755)
    write_lf(STAGE / "生成桌面快捷方式.sh", DESKTOP, 0o755)
    write_lf(STAGE / "自检.sh", SELFTEST, 0o755)
    write_lf(STAGE / "README-UOS.txt", README_UOS)


def tar_filter(info):
    if info.isdir() or info.name.endswith(".sh"):
        info.mode = 0o755
    else:
        info.mode = 0o644
    info.uid = 0
    info.gid = 0
    info.uname = ""
    info.gname = ""
    return info


def build():
    copy_sources()
    DIST.mkdir(exist_ok=True)
    tar_path = DIST / "学法用法速查助手_UOS_arm64.tar.gz"
    with tarfile.open(tar_path, "w:gz", compresslevel=9) as tf:
        tf.add(STAGE, arcname=STAGE.name, filter=tar_filter)

    # 单文件自解压：把 tar.gz 用 base64 附在 shell 脚本尾部
    run_path = DIST / "学法用法速查助手_UOS_arm64.run"
    payload = base64.b64encode(tar_path.read_bytes()).decode("ascii")
    header = """#!/bin/sh
set -e
TARGET="$(pwd)"
awk '/^__PAYLOAD__$/{f=1;next} f' "$0" | base64 -d | tar -xz -C "$TARGET"
echo "已解压：$TARGET/xuefa-uos-arm64"
echo "启动：sh \\"$TARGET/xuefa-uos-arm64/启动速查.sh\\""
exit 0
__PAYLOAD__
"""
    write_lf(run_path, header + payload + "\n", 0o755)
    print("tar.gz: %.2f MB" % (tar_path.stat().st_size / 1024 / 1024))
    print("run:    %.2f MB" % (run_path.stat().st_size / 1024 / 1024))
    return tar_path, run_path


if __name__ == "__main__":
    build()
