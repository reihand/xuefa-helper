# -*- coding: utf-8 -*-
"""
从 Word 题库（.docx）解析题目，合并进 questions.json。

用法：
  python3 tools/import_docx.py "我的题库.docx"
  python3 tools/import_docx.py <docx> --no-merge -o tools/docx_only.json
  python3 tools/import_docx.py <docx> --dump tools/docx_dump.txt

只依赖标准库（zipfile / re / json），不需要 python-docx。
解析规则对应该题库的实际排版：
  - 章节标题「一、单选题（40 道）」决定后续题型，可能粘在上一行行尾
  - 题号行「1、xxx（D）」开头，答案在括号里；少数题另起一行写「正确答案：C」
  - 选项行「A.xxx」，也可能多个选项内联在同一段「A.xxxB.xxxC.xxx」
  - 无前缀行是上一字段的续行（Word 自动换行 / 表格拆段）
  - 页脚诗句「……——苏轼」和「第X部分」标题过滤掉
"""
import argparse
import json
import pathlib
import re
import sys
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
DEFAULT_BANK = ROOT / "src" / "data" / "questions.json"
LAWS_FILE = ROOT / "src" / "data" / "laws_scope.json"

PARA_RE = re.compile(r"<w:p[ >].*?</w:p>", re.S)
TEXT_RE = re.compile(r"<w:t[^>]*>(.*?)</w:t>", re.S)
QNUM_RE = re.compile(r"^(\d+)\s*[、.．]\s*(.*)$")
OPT_RE = re.compile(r"^([A-Ha-h])\s*[.、．:：)）]\s*(.*)$")
SECTION_RE = re.compile(r"^[一二三四五六七八九十]+、\s*(单选题|多选题|判断题)")
SEC_INLINE_RE = re.compile(r"[一二三四五六七八九十]+、\s*(单选题|多选题|判断题)\s*[（(]\s*\d+\s*[题道]\s*[）)]")
PART_RE = re.compile(r"^第[一二三四五六七八九十]+部分\s*(.*)$")
ANSWER_LINE_RE = re.compile(r"^正确答案\s*[:：]\s*([A-Ha-h√×✓✕对错TF]+)")
# 只认「A.」「A、」这类选项前缀；不能包含 ) 或 ），否则会把答案括号「（D）」误判成选项
INLINE_MARK_RE = re.compile(r"([A-Ha-h])\s*[.、．:：]")
ANSWER_RE = re.compile(r"[（(]\s*([A-Ha-h]{1,8}|[√×✓✕]|[对错])\s*[）)]")
LAW_RE = re.compile(r"《([^》]{2,60})》")

TYPE_OF = {"单选题": "single", "多选题": "multi", "判断题": "judge"}
JUDGE_TRUE = ("√", "✓", "对", "A", "T")
JUDGE_FALSE = ("×", "✕", "错", "B", "F", "X")


def extract_paragraphs(path):
    """读 word/document.xml，按 w:p 抽纯文本，丢掉页脚装饰与 XML 属性残留。"""
    xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8")
    out = []
    for para in PARA_RE.findall(xml):
        txt = "".join(TEXT_RE.findall(para))
        txt = (txt.replace("&amp;", "&").replace("&lt;", "<")
                  .replace("&gt;", ">").replace("&quot;", '"')
                  .replace("&apos;", "'").strip())
        if not txt:
            continue
        # 段落属性节点被误并进正文（w:tabs / w:spacing 之类）
        if "<w:" in txt:
            continue
        # 页脚名言：「古之立大事者……——苏轼」
        if "——" in txt and not re.search(r"[（()）]", txt) and not QNUM_RE.match(txt):
            continue
        out.append(txt)
    return out


def split_inline(text):
    """把「A.xxxB.xxxC.xxx」拆成 (前缀, {字母: 内容})；拆不动就原样返回。"""
    marks = []
    for m in INLINE_MARK_RE.finditer(text):
        if m.start() and re.match(r"[A-Za-z0-9]", text[m.start() - 1]):
            continue
        marks.append(m)
    if not marks:
        return text, {}
    letters = [m.group(1).upper() for m in marks]
    # 字母必须递增，否则大概率是正文里的英文缩写而不是选项
    if any(ord(letters[i]) >= ord(letters[i + 1]) for i in range(len(letters) - 1)):
        return text, {}
    opts = {}
    for i, m in enumerate(marks):
        end = marks[i + 1].start() if i + 1 < len(marks) else len(text)
        opts[letters[i]] = text[m.end():end].strip()
    if not any(opts.values()):
        return text, {}
    return text[:marks[0].start()].strip(), opts


def parse_blocks(lines):
    """把段落流切成题目块；返回 [{stem, options, type, part, answer_raw}]。"""
    blocks = []
    cur_type = "single"
    cur_part = ""
    cur = None
    field = "stem"

    def flush():
        nonlocal cur
        if cur is not None:
            blocks.append(cur)
        cur = None

    def feed(text):
        """把一段正文并入当前题目：识别答案行、内联选项、普通续行。"""
        nonlocal cur, field
        if cur is None:
            return
        text = text.strip()
        if not text:
            return
        m_ans = ANSWER_LINE_RE.match(text)
        if m_ans:
            cur["answer_raw"] = m_ans.group(1).strip().upper()
            return

        prefix, opts = split_inline(text)
        if opts:
            if prefix:
                if field == "stem":
                    cur["stem"] += prefix
                elif field in cur["options"]:
                    cur["options"][field] += prefix
            for k, v in opts.items():
                if not v:
                    continue
                # 同一选项被 Word 拆到多段时拼接，而不是覆盖
                cur["options"][k] = cur["options"].get(k, "") + v
            field = list(opts)[-1]
            return

        m_opt = OPT_RE.match(text)
        if m_opt:
            letter = m_opt.group(1).upper()
            cur["options"][letter] = m_opt.group(2).strip()
            field = letter
            return

        # 无标记 = 上一字段的续行
        if field == "stem":
            cur["stem"] += text
        elif field in cur["options"]:
            cur["options"][field] += text

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue

        # 章节标题粘在上一行行尾（「…D.宪法、部门规章、地方性法规二、多选题（5 题）」）
        m_inline = SEC_INLINE_RE.search(line)
        if m_inline and m_inline.start() > 0:
            feed(line[:m_inline.start()])
            flush()
            cur_type = TYPE_OF[m_inline.group(1)]
            continue

        m_sec = SECTION_RE.match(line)
        if m_sec:
            flush()
            cur_type = TYPE_OF[m_sec.group(1)]
            continue
        m_part = PART_RE.match(line)
        if m_part:
            flush()
            cur_part = m_part.group(1).strip() or line.strip()
            continue
        if re.match(r"^[（(]\s*试题\s*[:：]", line):
            continue

        m = QNUM_RE.match(line)
        if m:
            flush()
            cur = {
                "stem": "",
                "options": {},
                "type": cur_type,
                "part": cur_part,
                "answer_raw": "",
            }
            field = "stem"
            feed(m.group(2).strip())
            continue

        feed(line)

    flush()
    return blocks


def load_laws():
    if not LAWS_FILE.exists():
        return []
    data = json.loads(LAWS_FILE.read_text(encoding="utf-8"))
    return data.get("items", [])


def match_law(stem, laws):
    """题干里的《XXX》去附件3清单里做包含匹配，匹配不上就留空。"""
    names = [item["name"] for item in laws]
    for name in LAW_RE.findall(stem):
        key = name.strip()
        for full in names:
            if key in full or full in key:
                return full
    # 「宪法」「民法典」这类简称在清单里是全称，做一次保守兜底
    for short, full in (("宪法", "《中华人民共和国宪法》"),
                        ("民法典", "《中华人民共和国民法典》")):
        if short in stem and full in names:
            return full
    return ""


def to_questions(blocks, source, laws):
    questions = []
    stats = {"no_answer": 0, "bad_answer": 0, "placeholder": 0, "ok": 0}

    for b in blocks:
        stem = re.sub(r"\s+", "", b["stem"])
        if not stem:
            continue
        qtype = b["type"]
        options = {k: re.sub(r"\s+", "", v) for k, v in b["options"].items() if v}

        # 答案优先取「正确答案：C」行，否则取题干里最后一个答案括号
        raw = b.get("answer_raw", "").strip().upper()
        hit = None
        if not raw:
            hits = list(ANSWER_RE.finditer(stem))
            if hits:
                hit = hits[-1]
                raw = hit.group(1).strip().upper()
        if hit is not None:
            # 回填成空格，让入库题干与考试页面上的「（ ）」一致
            stem = stem[:hit.start()] + "（  ）" + stem[hit.end():]
        if not raw:
            stats["no_answer"] += 1
            continue

        caution = ""
        if qtype == "judge":
            options = {"A": "正确", "B": "错误"}
            if raw in JUDGE_TRUE:
                answer = "A"
            elif raw in JUDGE_FALSE:
                answer = "B"
            else:
                stats["bad_answer"] += 1
                continue
        else:
            answer = "".join(ch for ch in raw if ch in "ABCDEFGH")
            if not answer:
                stats["bad_answer"] += 1
                continue
            missing = [ch for ch in answer if ch not in options]
            if missing or not options:
                for ch in sorted(set(list("ABCD") + list(answer))):
                    options.setdefault(ch, "（选项内容未解析：" + ch + "）")
                caution = "选项内容解析不全，请核对原文"
                stats["placeholder"] += 1

        law = match_law(stem, laws)
        if "民法总则" in stem:
            caution = (caution + "；" if caution else "") + "《民法总则》已废止，相关内容见《中华人民共和国民法典》，请核对现行条文"
        elif qtype == "multi" and len(answer) < 2:
            caution = (caution + "；" if caution else "") + "题库标注为多选但只解析出单个答案，请核对"
        if "2025" in source or "2025" in b["part"]:
            caution = (caution + "；" if caution else "") + "2025 年度题库，2026 年度范围可能有修订"

        questions.append({
            "id": "",
            "type": qtype,
            "stem": stem,
            "options": options,
            "answer": answer,
            "basis": "",
            "law": law,
            "source": (source + (" / " + b["part"] if b["part"] else "")).strip(),
            "caution": caution,
        })
        stats["ok"] += 1

    return questions, stats


def norm_key(stem):
    return re.sub(r"[^\u4e00-\u9fffA-Za-z0-9]", "", stem)


def merge(existing, imported):
    seen = set()
    merged = []
    for q in existing + imported:
        key = norm_key(q["stem"])
        if not key or key in seen:
            continue
        seen.add(key)
        merged.append(q)
    for i, q in enumerate(merged, 1):
        q["id"] = "Q%04d" % i
    return merged


def main():
    ap = argparse.ArgumentParser(description="从 Word 题库导入题目")
    ap.add_argument("docx", help=".docx 题库文件")
    ap.add_argument("-o", "--output", default=str(DEFAULT_BANK), help="输出路径")
    ap.add_argument("--no-merge", action="store_true", help="不合并现有题库，只导出本次解析结果")
    ap.add_argument("--dump", help="把抽取的段落文本写到指定文件，便于人工核对")
    args = ap.parse_args()

    src = pathlib.Path(args.docx)
    if not src.exists():
        print("input not found:", src)
        return 1

    lines = extract_paragraphs(src)
    if args.dump:
        pathlib.Path(args.dump).write_text("\n".join(lines), encoding="utf-8")
        print("paragraphs dumped:", len(lines), "->", args.dump)

    blocks = parse_blocks(lines)
    questions, stats = to_questions(blocks, src.name, load_laws())
    if not questions:
        print("没有解析出有效题目，请检查 docx 排版")
        return 1

    existing = []
    bank = pathlib.Path(DEFAULT_BANK)
    if not args.no_merge and bank.exists():
        payload = json.loads(bank.read_text(encoding="utf-8"))
        existing = payload.get("questions", [])
    merged = merge(existing, questions)

    out = pathlib.Path(args.output)
    if out.exists():
        bak = out.with_suffix(out.suffix + ".bak")
        bak.write_bytes(out.read_bytes())
        print("backup:", bak)
    out.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "version": "1.2",
        "updated": "2026-11-19",
        "scope": "桂司通〔2026〕号 附件3（382部）范围内公开真题 + 本地 Word 题库 + 导入补充",
        "total": len(merged),
        "note": "题目与答案取自公开题库、本地 Word 题库及导入文件；basis 为法条依据或题库解析；caution 非空时表示需核对。",
        "questions": merged,
    }
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")

    types = {}
    for q in merged:
        types[q["type"]] = types.get(q["type"], 0) + 1
    print("paragraphs:", len(lines), "blocks:", len(blocks))
    print("parsed:", stats, "existing:", len(existing), "merged:", len(merged))
    print("types:", types)
    print("written:", out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
