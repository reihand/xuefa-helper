# -*- coding: utf-8 -*-
"""
命令行补库入口。Excel / JSON -> questions.json，与现有题库合并去重。

解析逻辑统一放在 src/xlsx_import.py（纯标准库，无 openpyxl 依赖），
与 exe 里的「导入题库 Excel」按钮共用同一套代码。

原生支持两种 Excel 来源：
  A) Greasy Fork #560523「广西普法考试题库录入」导出的三工作表文件：
       单选题录入 / 多选题录入 / 判断题录入
  B) 本工具标准表头：题干 | 类型 | A | B | ... | H | 答案 | 依据 | 法规 | 来源

用法：
  python3 tools/import_bank.py 题库导出_2026-11-19.xlsx
  python3 tools/import_bank.py new.xlsx --no-merge -o extra.json
  python3 tools/import_bank.py other.json -o merged.json
  python3 tools/import_bank.py new.xlsx --dry-run        # 只报数，不写文件
"""
import argparse
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
DEFAULT_BANK = ROOT / "src" / "data" / "questions.json"
sys.path.insert(0, str(ROOT / "src"))

import xlsx_import as xi  # noqa: E402


def main():
    ap = argparse.ArgumentParser(description="题库补库工具（支持 #560523 导出格式）")
    ap.add_argument("input", help="Excel(.xlsx/.xlsm) 或 JSON 题库文件")
    ap.add_argument("-o", "--output", default=str(DEFAULT_BANK), help="输出路径")
    ap.add_argument("--no-merge", action="store_true", help="不合并现有题库")
    ap.add_argument("--dry-run", action="store_true", help="只统计，不写文件")
    args = ap.parse_args()

    src = pathlib.Path(args.input)
    if not src.exists():
        print("input not found:", src)
        return 1
    suffix = src.suffix.lower()
    if suffix == ".xls":
        print("只支持 .xlsx / .xlsm，请用 Excel 另存为 .xlsx 后重试")
        return 1

    if suffix in (".xlsx", ".xlsm"):
        try:
            imported, detail = xi.import_xlsx(src)
        except Exception as exc:
            print("读取 Excel 失败：", exc)
            return 1
        print("Excel sheets:")
        print("\n".join(detail) if detail else "  (empty)")
    else:
        imported = xi.read_json_questions(src)

    if not imported:
        print("没有解析出有效题目，请检查表头与答案列")
        return 1
    types = {}
    for q in imported:
        types[q["type"]] = types.get(q["type"], 0) + 1
    print("imported:", len(imported), "types:", types)

    existing = []
    if not args.no_merge and DEFAULT_BANK.exists():
        existing = xi.read_json_questions(DEFAULT_BANK)
    merged = xi.merge_questions(existing, imported)
    print("existing:", len(existing), "merged:", len(merged),
          "(new:", len(merged) - len(existing), ")")
    if args.dry_run:
        print("dry-run: 未写文件")
        return 0

    out = pathlib.Path(args.output)
    had_backup = out.exists()
    xi.save_bank(out, merged)
    if had_backup:
        print("backup:", str(out) + ".bak")
    print("written:", out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
