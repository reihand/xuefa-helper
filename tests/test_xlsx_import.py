# -*- coding: utf-8 -*-
"""#560523 导出格式的 Excel 导入回归测试（纯标准库，无需 openpyxl）。"""
import pathlib
import sys
import tempfile
import unittest

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
import xlsx_import as xi  # noqa: E402

FIXTURE = ROOT / "tests" / "fixtures" / "gf560523_sample.xlsx"


class XlsxImportTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.questions, cls.detail = xi.import_xlsx(FIXTURE)

    def test_three_sheets(self):
        sheets = xi.read_xlsx(FIXTURE)
        self.assertEqual([n for n, _ in sheets],
                         ["单选题录入", "多选题录入", "判断题录入"])

    def test_all_six_parsed(self):
        self.assertEqual(len(self.questions), 6)
        self.assertEqual(len(self.detail), 3)

    def test_types_by_sheet_name(self):
        self.assertEqual([q["type"] for q in self.questions],
                         ["single", "single", "multi", "multi", "judge", "judge"])

    def test_answers(self):
        self.assertEqual([q["answer"] for q in self.questions],
                         ["A", "C", "ABC", "ABDE", "A", "B"])

    def test_judge_options_generated(self):
        for q in self.questions:
            if q["type"] == "judge":
                self.assertEqual(q["options"], {"A": "正确", "B": "错误"})

    def test_answer_letters_exist_in_options(self):
        for q in self.questions:
            for ch in q["answer"]:
                self.assertIn(ch, q["options"], q["stem"])

    def test_merge_dedup_keeps_existing(self):
        existing = [{
            "id": "Q0001", "type": "single",
            "stem": "《宪法》规定，中华人民共和国的一切权力属于（）。",
            "options": {"A": "人民"}, "answer": "A",
            "basis": "人工校正过的依据", "law": "", "source": "old", "caution": "",
        }]
        merged = xi.merge_questions(existing, self.questions)
        self.assertEqual(len(merged), 6)  # 6 导入题里有 1 题与 existing 同题干
        self.assertEqual(merged[0]["basis"], "人工校正过的依据")
        self.assertEqual([q["id"] for q in merged],
                         ["Q%04d" % i for i in range(1, 7)])

    def test_save_and_reload(self):
        with tempfile.TemporaryDirectory() as d:
            out = pathlib.Path(d) / "questions.json"
            merged = xi.merge_questions([], self.questions)
            xi.save_bank(out, merged)
            self.assertFalse((pathlib.Path(d) / "questions.json.tmp").exists())
            self.assertEqual(len(xi.read_json_questions(out)), 6)
            xi.save_bank(out, merged)  # 第二次写入生成 .bak
            self.assertTrue((pathlib.Path(d) / "questions.json.bak").exists())

    def test_answer_normalization(self):
        self.assertEqual(xi.norm_answer("A,C,D", "multi"), "ACD")
        self.assertEqual(xi.norm_answer("ABD", "multi"), "ABD")
        self.assertEqual(xi.norm_answer("T", "judge"), "A")
        self.assertEqual(xi.norm_answer("错", "judge"), "B")
        self.assertEqual(xi.norm_answer("√", "judge"), "A")
        self.assertEqual(xi.norm_answer("×", "judge"), "B")


if __name__ == "__main__":
    unittest.main(verbosity=2)
