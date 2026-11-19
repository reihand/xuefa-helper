# -*- coding: utf-8 -*-
"""
生成免安装单文件版：一个 HTML 文件 = 页面 + 检索逻辑 + 10196 道题。

用法：python tools/build_standalone_html.py
产物：dist/学法用法速查助手_单文件.html

双击文件即可在浏览器打开，Windows / UOS / macOS 通用，不需要 Python。
"""
import json
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "tools" / "standalone_template.html"
SEARCH_JS = ROOT / "tools" / "standalone_search.js"
BANK = ROOT / "src" / "data" / "questions.json"
LAWS = ROOT / "src" / "data" / "laws_scope.json"
OUT = ROOT / "dist" / "学法用法速查助手_单文件.html"


def safe_json(obj):
    """JSON 内嵌进 <script> 时转义 < 和 U+2028/2029，避免提前闭合标签或语法错误。"""
    return (json.dumps(obj, ensure_ascii=False, separators=(",", ":"))
            .replace("<", "\\u003c")
            .replace("\u2028", "\\u2028")
            .replace("\u2029", "\\u2029"))


def build():
    template = TEMPLATE.read_text(encoding="utf-8")
    search_js = SEARCH_JS.read_text(encoding="utf-8")
    bank = json.loads(BANK.read_text(encoding="utf-8"))
    laws = json.loads(LAWS.read_text(encoding="utf-8"))

    html = template.replace("/*__SEARCH_JS__*/", search_js)
    html = html.replace("/*__BANK_JSON__*/", '{"questions":%s}' % safe_json(bank["questions"]))
    html = html.replace("/*__LAWS_JSON__*/", safe_json(laws))

    if "/*__" in html:
        raise RuntimeError("模板占位符未全部替换")

    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(html, encoding="utf-8", newline="\n")
    print("产物：%s" % OUT)
    print("题目：%d 道；法规：%d 部" % (len(bank["questions"]), len(laws["items"])))
    print("大小：%.2f MB" % (OUT.stat().st_size / 1024 / 1024))
    return OUT


if __name__ == "__main__":
    build()
