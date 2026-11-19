# -*- coding: utf-8 -*-
"""
程序入口。

源码运行：python3 src/main.py
打包后自检：学法用法速查助手.exe --selftest
  → 在 exe 同目录写出 selftest.txt
打包后导入：学法用法速查助手.exe --import <xlsx路径>
  → 合并到 exe 同目录 questions.json，写出 import_report.txt
     （--windowed 没有控制台，所有结果落文件）
"""
import sys

from ui import main


def _report(filename, lines):
    import paths
    out = paths.app_dir() / filename
    out.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return out


def selftest():
    import paths
    import bank as bankmod
    import xlsx_import

    lines = []
    data = bankmod.QuestionBank.load()
    covered, total = data.law_coverage()
    lines.append("学法用法速查助手 自检报告")
    lines.append("题库加载：%d 题" % len(data.questions))
    lines.append("其中官方练习题库：%d 题" % sum(
        1 for q in data.questions if "官方题库" in q.get("source", "")))
    lines.append("法规清单：%d 部，题库直接关联 %d 部" % (total, covered))
    lines.append("数据来源：%s" % paths.resource("questions.json"))
    lines.append("frozen：%s" % paths.is_frozen())
    lines.append("Excel 导入模块：%s" % ("OK" if hasattr(xlsx_import, "import_xlsx") else "MISSING"))

    cases = [
        ("整题粘贴", "3. 根据《中华人民共和国民法典》，向人民法院请求保护民事权利的诉讼时效期间为（ ）。A.一年 B.二年 C.三年 D.五年"),
        ("关键词", "撤销权的行使范围"),
        ("判断题", "不满八周岁的未成年人为无民事行为能力人"),
        ("本地Word题库", "根据法律的地位和效力从高到低排列的是"),
    ]
    ok = 0
    for label, q in cases:
        r = data.search(q)
        if r:
            score, hit = r[0]
            lines.append("%s -> %.3f | %s | 答案 %s" % (
                label, score, hit["id"], data.answer_text(hit)))
            if score >= 0.60:
                ok += 1
        else:
            lines.append("%s -> 无命中" % label)
    lines.append("可信命中：%d/%d" % (ok, len(cases)))
    lines.append("结果：%s" % ("PASS" if ok == len(cases) and len(data.questions) >= 300
                              and total == 382 and hasattr(xlsx_import, "import_xlsx") else "FAIL"))

    _report("selftest.txt", lines)
    return 0 if ok == len(cases) else 1


def import_bank_cli(path):
    """命令行导入：供批量补库，也用来验证 frozen 环境下的 xlsx 解析链路。"""
    import paths
    import bank as bankmod
    import xlsx_import as xi

    lines = ["学法用法速查助手 导入报告", "来源：%s" % path]
    try:
        imported, detail = xi.import_xlsx(path)
    except Exception as exc:
        lines.append("读取失败：%s" % exc)
        lines.append("结果：FAIL")
        _report("import_report.txt", lines)
        return 1
    lines.extend(detail)
    lines.append("解析：%d 题" % len(imported))
    if not imported:
        lines.append("结果：FAIL（无有效题目）")
        _report("import_report.txt", lines)
        return 1

    before = bankmod.QuestionBank.load()
    merged = xi.merge_questions(before.questions, imported)
    target = paths.external_path("questions.json")
    try:
        xi.save_bank(target, merged)
        after = bankmod.QuestionBank.load()
    except Exception as exc:
        lines.append("写入失败：%s" % exc)
        lines.append("结果：FAIL")
        _report("import_report.txt", lines)
        return 1

    lines.append("导入前：%d 题" % len(before.questions))
    lines.append("导入后：%d 题（新增 %d）" % (len(merged), len(merged) - len(before.questions)))
    lines.append("写入：%s" % target)
    lines.append("重载校验：%d 题，法规 %d 部" % (
        len(after.questions), len(after.laws.get("items", []))))
    passed = (len(after.questions) == len(merged)
              and len(after.laws.get("items", [])) == 382)
    lines.append("答案样例：" + " / ".join(q["answer"] for q in imported[:6]))
    lines.append("结果：%s" % ("PASS" if passed else "FAIL"))
    _report("import_report.txt", lines)
    return 0 if passed else 1


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        sys.exit(selftest())
    if "--import" in sys.argv:
        idx = sys.argv.index("--import")
        src = sys.argv[idx + 1] if len(sys.argv) > idx + 1 else ""
        sys.exit(import_bank_cli(src) if src else 2)
    sys.exit(main())
