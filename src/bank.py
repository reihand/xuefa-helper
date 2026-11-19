# -*- coding: utf-8 -*-
"""
题库加载与中文模糊检索。

实现要点：
- 纯标准库，无第三方依赖（difflib / re / unicodedata / json）。
- 题干做规范化后建 2-gram 倒排索引；查询先用倒排召回候选，再用
  SequenceMatcher 精排，兼顾「复制整题」「只粘半句」「复制带题号」三种输入。
- 复杂度：建索引 O(N*L)，N=题数、L=平均题干长度；
  查询 O(G*P + K*L^2)，G=查询 gram 数、P=平均倒排链长、K=候选数。
  瓶颈是 SequenceMatcher 的 L^2，靠倒排把 K 压到几十以内。
  本库 10196 题，建索引约 0.8s，单次查询在毫秒级。
- 官方题库存在「题干几乎相同但选项顺序不同」的题，只比题干会把答案字母
  给错；因此粘贴整题时额外解析选项文本，用选项一致性做排序加权。

可独立运行自测：python3 src/bank.py "撤销权的行使范围以什么为限"
"""
import difflib
import json
import re
import unicodedata

import paths

TYPE_NAME = {"single": "单选题", "multi": "多选题", "judge": "判断题"}

# 复制题目时常见的题号前缀、平台噪声
_NOISE = re.compile(
    r"^\s*(?:第\s*)?\d+\s*[.、．:：)）]\s*|[（(]\s*[)）]|[_＿]{2,}|[ \t\r\n]+"
)
# 选项起始标记，用于把「整题文本」截成纯题干
_OPTION_START = re.compile(r"(?:^|[\s，。；;])[A-Ga-g]\s*[.、．:：)）]\s*")
_KEEP = re.compile(r"[^\u4e00-\u9fffA-Za-z0-9]")
# 选项标记：前缀可为行首/空白/标点/汉字，兼容「A.一年B.二年」无空格粘贴
_OPT_MARK = re.compile(r"(^|[\s，。；;（(\u4e00-\u9fff])([A-Ha-h])\s*[.、．:：)）]\s*")


def normalize(text):
    """全角转半角、去标点空白、去题号与填空占位，只留中英文数字。"""
    if not text:
        return ""
    s = unicodedata.normalize("NFKC", str(text))
    s = _NOISE.sub("", s)
    s = _KEEP.sub("", s)
    return s.lower()


def extract_stem(text):
    """从「题干 + 选项」的整段复制文本里截出题干部分。"""
    s = unicodedata.normalize("NFKC", str(text or ""))
    m = _OPTION_START.search(s)
    if m:
        s = s[: m.start()]
    return normalize(s)


def extract_options(text):
    """从整题复制文本里解析 {字母: 选项文本}；至少两个选项才认为有效。"""
    s = unicodedata.normalize("NFKC", str(text or ""))
    marks = list(_OPT_MARK.finditer(s))
    if len(marks) < 2:
        return {}
    out = {}
    for i, m in enumerate(marks):
        letter = m.group(2).upper()
        # 下一个标记的字母位置为界；前缀若是汉字则属于本选项的最后一个字
        end = marks[i + 1].start(2) if i + 1 < len(marks) else len(s)
        seg = s[m.end():end].strip(" \t\r\n，。；;、")
        if seg and letter not in out:
            out[letter] = seg
    return out


def option_agreement(parsed, q):
    """候选题同字母上的选项文本与粘贴内容一致的比例，用于区分同题干异选项的题。"""
    opts = q.get("options") or {}
    if not parsed or not opts:
        return 0.0
    hit = 0
    for letter, txt in parsed.items():
        ref, val = normalize(opts.get(letter, "")), normalize(txt)
        if not ref or not val:
            continue
        if val == ref or (len(val) >= 2 and (val in ref or ref in val)):
            hit += 1
    return hit / len(parsed)


def bigrams(s):
    """中文友好的 2-gram；长度不足 2 时退化为单字。"""
    if len(s) < 2:
        return {s} if s else set()
    return {s[i:i + 2] for i in range(len(s) - 1)}


class QuestionBank:
    def __init__(self, questions=None, laws=None):
        self.questions = list(questions or [])
        self.laws = laws or {"items": [], "total": 0}
        self._grams = []
        self._index = {}
        self._law_names = []
        self._build()

    @classmethod
    def load(cls):
        """优先读 exe 同目录的外置题库（便于补库），否则读内置题库。"""
        q_path = paths.external_path("questions.json")
        if not q_path.exists():
            q_path = paths.resource("questions.json")
        with q_path.open("r", encoding="utf-8") as f:
            q_payload = json.load(f)

        l_path = paths.external_path("laws_scope.json")
        if not l_path.exists():
            l_path = paths.resource("laws_scope.json")
        with l_path.open("r", encoding="utf-8") as f:
            laws = json.load(f)

        return cls(q_payload.get("questions", []), laws)

    def _build(self):
        self._grams = []
        self._index = {}
        for i, q in enumerate(self.questions):
            # 题干 + 选项文本一并入索引：只粘了选项也能捞到题
            body = normalize(q.get("stem", "")) + "".join(
                normalize(v) for v in (q.get("options") or {}).values()
            )
            grams = bigrams(body)
            self._grams.append((grams, body))
            for g in grams:
                self._index.setdefault(g, set()).add(i)
        self._law_names = [normalize(x.get("name", "")) for x in self.laws.get("items", [])]

    def law_coverage(self):
        """返回 (题库覆盖到的法规数, 法规总数)。"""
        covered = set()
        for q in self.questions:
            law = q.get("law") or ""
            if law:
                covered.add(normalize(law))
        all_names = {n for n in self._law_names if n}
        return len(covered & all_names), len(all_names)

    def search_laws(self, keyword, limit=30):
        """按名称检索附件3 法规清单（只作范围浏览，不含全文）。"""
        key = normalize(keyword)
        if not key:
            return self.laws.get("items", [])[:limit]
        out = []
        for item in self.laws.get("items", []):
            name = normalize(item.get("name", ""))
            if key in name:
                out.append((1.0, item))
            else:
                r = difflib.SequenceMatcher(None, key, name).ratio()
                if r >= 0.5:
                    out.append((r, item))
        out.sort(key=lambda x: -x[0])
        return [it for _, it in out[:limit]]

    def search(self, text, topk=8):
        """
        返回 [(score, question), ...]，按 score 降序。
        score >= 0.62 视为可信命中；0.40~0.62 为疑似；< 0.40 仅供参考。
        """
        stem = extract_stem(text)
        raw = normalize(text)
        parsed_opts = extract_options(text)
        if not stem:
            return []
        qg = bigrams(stem)
        if not qg:
            return []

        # 倒排召回：命中任一 gram 的题进候选
        cand = set()
        for g in qg:
            cand |= self._index.get(g, set())
        if not cand:
            return []

        scored = []
        for i in cand:
            grams, body = self._grams[i]
            inter = len(qg & grams)
            cover = inter / len(qg)
            if cover < 0.22:
                continue
            # SequenceMatcher 对长文本较慢，只在规范化后的题干上比
            ratio = difflib.SequenceMatcher(None, stem, body[: max(len(stem) * 2, 40)]).ratio()
            score = 0.68 * cover + 0.32 * ratio

            # 连续长片段命中：复制整题时显著提升区分度
            if len(stem) >= 8 and stem[:8] in body:
                score += 0.12
            # 查询里直接出现该法规名：加权
            law = normalize(self.questions[i].get("law", ""))
            if law and len(law) >= 6 and law[:6] in raw:
                score += 0.06
            score = min(score, 1.0)

            # 同题干不同选项顺序的题：粘贴了选项时用选项一致性打破平局。
            # 排序键不封顶（否则两个 1.000 又并列），展示分仍封顶到 1.0。
            agree = option_agreement(parsed_opts, self.questions[i]) if len(parsed_opts) >= 2 else 0.0
            final = score + 0.30 * agree
            scored.append((final, min(final, 1.0), self.questions[i]))

        scored.sort(key=lambda x: -x[0])
        return [(disp, q) for _, disp, q in scored[:topk]]

    @staticmethod
    def answer_text(q):
        """把内部答案字母转成展示文本。判断题直接给「正确/错误」。"""
        ans = q.get("answer", "")
        if q.get("type") == "judge":
            return "正确" if ans == "A" else "错误"
        return "、".join(list(ans))

    @staticmethod
    def option_text(q, letter):
        return (q.get("options") or {}).get(letter, "")


def _selftest():
    import sys
    bank = QuestionBank.load()
    covered, total = bank.law_coverage()
    print("题库加载：%d 题；法规清单 %d 部，题库直接关联 %d 部" % (
        len(bank.questions), total, covered))
    query = sys.argv[1] if len(sys.argv) > 1 else "撤销权的行使范围以什么为限"
    print("查询：%s" % query)
    results = bank.search(query)
    if not results:
        print("未召回候选")
        return
    for score, q in results[:5]:
        print("  %.3f  [%s] %s  => %s" % (
            score, TYPE_NAME.get(q["type"], q["type"]),
            q["stem"][:36], QuestionBank.answer_text(q)))


if __name__ == "__main__":
    _selftest()
