# -*- coding: utf-8 -*-
"""检索核心自测：python3 -m unittest discover -s tests -v"""
import pathlib
import sys
import unittest

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent / "src"))
import bank as bankmod  # noqa: E402


class BankTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.b = bankmod.QuestionBank.load()

    def test_load(self):
        self.assertGreaterEqual(len(self.b.questions), 300)
        self.assertEqual(self.b.laws.get("total"), 382)
        self.assertEqual(len(self.b.laws.get("items", [])), 382)

    def test_exact_paste(self):
        q = "3. 根据《中华人民共和国民法典》，向人民法院请求保护民事权利的诉讼时效期间为（ ）。A.一年 B.二年 C.三年 D.五年"
        r = self.b.search(q)
        self.assertTrue(r)
        self.assertGreaterEqual(r[0][0], 0.90)
        self.assertEqual(r[0][1]["answer"], "C")
        self.assertIn("诉讼时效", r[0][1]["stem"])

    def test_partial_keywords(self):
        r = self.b.search("撤销权的行使范围")
        self.assertTrue(r)
        self.assertIn("撤销权", r[0][1]["stem"])

    def test_judge_answer_text(self):
        r = self.b.search("不满八周岁的未成年人为无民事行为能力人")
        self.assertTrue(r)
        self.assertEqual(self.b.answer_text(r[0][1]), "正确")

    def test_multi_answer_text(self):
        r = self.b.search("消费者和经营者发生消费者权益争议")
        self.assertTrue(r)
        self.assertEqual(self.b.answer_text(r[0][1]), "A、B、C、D")

    def test_docx_imported_question(self):
        """本地 Word 题库合并进来的题也能被检索到。"""
        r = self.b.search("根据法律的地位和效力从高到低排列的是")
        self.assertTrue(r)
        self.assertGreaterEqual(r[0][0], 0.90)
        self.assertEqual(r[0][1]["answer"], "C")

    def test_noise_tolerance(self):
        a = self.b.search("宪法的修改由全国人大常委会或者五分之一以上代表提议")
        b = self.b.search("宪法 修改 提议 三分之二 多数通过")
        self.assertTrue(a)
        self.assertTrue(b)

    def test_law_scope(self):
        items = self.b.search_laws("民法典")
        self.assertTrue(any("民法典" in x["name"] for x in items))
        covered, total = self.b.law_coverage()
        self.assertEqual(total, 382)
        self.assertGreater(covered, 0)

    def test_empty(self):
        self.assertEqual(self.b.search(""), [])
        self.assertEqual(self.b.search("   "), [])

    def test_every_question_well_formed(self):
        """每题必须有题干、答案、合法题型；选项缺字母或整块缺失时必须标 caution。"""
        for q in self.b.questions:
            self.assertTrue(q.get("stem"), q["id"])
            self.assertTrue(q.get("answer"), q["id"])
            self.assertIn(q["type"], ("single", "multi", "judge"))
            opts = q.get("options") or {}
            missing = [ch for ch in q["answer"] if ch not in opts]
            if missing:
                self.assertTrue(q.get("caution"), q["id"])

    def test_option_aware_ranking(self):
        """同题干、选项顺序不同的两道题，粘贴带选项整题时应命中选项一致的那道。"""
        q = ("3. 根据《中华人民共和国民法典》，向人民法院请求保护民事权利的"
             "诉讼时效期间为（ ）。A.一年 B.二年 C.三年 D.五年")
        r = self.b.search(q)
        self.assertTrue(r)
        self.assertEqual(r[0][1]["answer"], "C")
        self.assertEqual(r[0][1]["options"].get("A"), "一年")

    def test_basis_optional_with_caution(self):
        """公开题库可缺解析（界面有兜底文案），但风险题必须带 caution。"""
        no_basis = [q for q in self.b.questions if not q.get("basis")]
        self.assertGreater(len(no_basis), 0)
        cautioned = [q for q in self.b.questions if q.get("caution")]
        self.assertGreater(len(cautioned), 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
