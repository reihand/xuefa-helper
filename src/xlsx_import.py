# -*- coding: utf-8 -*-
"""
纯标准库 Excel 题库导入模块（供 GUI 与命令行共用）。

支持两类 Excel：
  A) Greasy Fork #560523「广西普法考试题库录入」导出的三工作表文件：
       单选题录入: 题目(stem) | A选项 | B选项 | C选项 | D选项 | 答案(填A/B/C/D) | 备注(可空)
       多选题录入: 题目(stem) | A选项 ... H选项(可空) | 答案(如ACD或A,C,D) | 备注(可空)
       判断题录入: 题目(stem) | 答案(填对/错 或 T/F) | 备注(可空)
  B) 本工具标准表头：题干 | 类型 | A | B | ... | G | 答案 | 依据 | 法规 | 来源

.xlsx 本质是 zip + XML，这里用 zipfile + ElementTree 直接读，
不引入 openpyxl，保证单文件 exe 仍然零第三方依赖。

可独立自测：python3 src/xlsx_import.py <某个.xlsx>
"""
import json
import pathlib
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
OPTION_COLS = list("ABCDEFGH")

TYPE_MAP = {
    "单选": "single", "单选题": "single", "单选题录入": "single", "单项选择": "single",
    "多选": "multi", "多选题": "multi", "多选题录入": "multi", "多项选择": "multi",
    "判断": "judge", "判断题": "judge", "判断题录入": "judge", "truefalse": "judge",
    "single": "single", "multi": "multi", "judge": "judge",
}

HEADER_ALIASES = {
    "stem": ["题干", "题目(stem)", "题目（stem）", "题目", "试题", "问题", "题目名称", "stem", "question"],
    "type": ["类型", "题型", "题目类型", "type", "qtype"],
    "answer": ["答案(填A/B/C/D)", "答案(如ACD或A,C,D)", "答案(填对/错 或 T/F)",
               "答案", "参考答案", "正确答案", "answer", "correct"],
    "basis": ["依据", "解析", "参考解析", "法条依据", "答案解析", "basis", "explain"],
    "law": ["法规", "关联法规", "所属法律", "law"],
    "source": ["来源", "备注", "备注(可空)", "备注（可空）", "source", "note"],
    "caution": ["提示", "注意", "caution"],
}
for _c in OPTION_COLS:
    HEADER_ALIASES[_c] = [
        _c, _c + "选项", _c + "选项(可空)", _c + "选项（可空）",
        "选项" + _c, _c.lower(), _c.lower() + "选项",
    ]


# ---------- xlsx 读取（zipfile + ElementTree） ----------

def _col_index(ref):
    """A1->0, B1->1, AA1->26；ref 缺失时退化为 0。"""
    m = re.match(r"([A-Za-z]+)", ref or "")
    if not m:
        return 0
    n = 0
    for ch in m.group(1).upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _read_sheet(data, shared):
    root = ET.fromstring(data)
    rows = []
    for row in root.iter(NS + "row"):
        cells = []
        for c in row.findall(NS + "c"):
            idx = _col_index(c.get("r"))
            while len(cells) <= idx:
                cells.append("")
            kind = c.get("t")
            v = c.find(NS + "v")
            inline = c.find(NS + "is")
            if kind == "s" and v is not None and v.text is not None:
                try:
                    val = shared[int(v.text)]
                except (ValueError, IndexError):
                    val = ""
            elif kind == "inlineStr" and inline is not None:
                val = "".join(x.text or "" for x in inline.iter(NS + "t"))
            elif v is not None and v.text is not None:
                val = v.text
            else:
                val = ""
            cells[idx] = str(val).strip()
        rows.append(cells)
    return rows


def read_xlsx(path):
    """返回 [(工作表名, [[单元格, ...], ...]), ...]。"""
    z = zipfile.ZipFile(str(path))
    names = set(z.namelist())
    if "xl/workbook.xml" not in names:
        z.close()
        raise ValueError("不是有效的 .xlsx 文件（缺 xl/workbook.xml）")

    # 共享字符串表：SheetJS 导出的文件用 inlineStr 或 sharedStrings 都可能
    shared = []
    if "xl/sharedStrings.xml" in names:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall(NS + "si"):
            shared.append("".join(t.text or "" for t in si.iter(NS + "t")))

    relmap = {}
    if "xl/_rels/workbook.xml.rels" in names:
        for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels")):
            relmap[r.get("Id")] = r.get("Target", "")

    wb = ET.fromstring(z.read("xl/workbook.xml"))
    node = wb.find(NS + "sheets")
    out = []
    for sh in (node if node is not None else []):
        target = relmap.get(sh.get(RNS + "id"), "")
        if target.startswith("/"):
            target = target.lstrip("/")
        elif not target.startswith("xl/"):
            target = "xl/" + target
        if target not in names:
            continue
        out.append((sh.get("name") or "", _read_sheet(z.read(target), shared)))
    z.close()
    return out


# ---------- 表头与答案解析 ----------

def norm_type(v):
    s = str(v or "").strip()
    return TYPE_MAP.get(s, TYPE_MAP.get(s.lower(), ""))


def sheet_hint(name):
    s = str(name or "")
    low = s.lower()
    if "单选" in s or "single" in low:
        return "single"
    if "多选" in s or "multi" in low:
        return "multi"
    if "判断" in s or "judge" in low or "truefalse" in low:
        return "judge"
    return ""


def norm_answer(v, qtype):
    s = str(v or "").strip().upper()
    s = re.sub(r"[、,，\s/|]+", "", s)
    s = s.replace("参考答案", "").replace("答案", "")
    if qtype == "judge":
        if s in ("正确", "对", "√", "✓", "A", "TRUE", "T", "是", "1"):
            return "A"
        if s in ("错误", "错", "×", "✕", "X", "B", "FALSE", "F", "否", "0"):
            return "B"
        return ""
    return "".join(ch for ch in s if ch in "ABCDEFGH")


def resolve_header(header):
    """表头 -> {字段: 列号}。先精确（忽略空白），再按包含匹配。"""
    cells = [str(x or "").strip() for x in header]
    lower = [re.sub(r"\s+", "", c) for c in cells]
    mapping = {}
    for field, aliases in HEADER_ALIASES.items():
        norms = [re.sub(r"\s+", "", a) for a in aliases]
        for i, cell in enumerate(lower):
            if cell in norms:
                mapping[field] = i
                break
        if field in mapping:
            continue
        for i, cell in enumerate(lower):
            if cell and any(len(a) >= 2 and (a in cell or cell in a) for a in norms):
                mapping[field] = i
                break
    if mapping.get("answer") == mapping.get("stem"):
        for i, cell in enumerate(lower):
            if "答案" in cell:
                mapping["answer"] = i
                break
    return mapping


def _cell(row, mapping, field):
    idx = mapping.get(field)
    if idx is None or idx >= len(row):
        return ""
    return str(row[idx] or "").strip()


def from_row(row, mapping, qtype_hint, source, sheet_name):
    stem = _cell(row, mapping, "stem")
    if not stem or stem.lower() in ("题干", "题目(stem)", "题目（stem）"):
        return None
    qtype = norm_type(_cell(row, mapping, "type")) or qtype_hint or "single"
    answer = norm_answer(_cell(row, mapping, "answer"), qtype)
    if not answer:
        return None

    caution = _cell(row, mapping, "caution")
    if qtype == "judge":
        options = {"A": "正确", "B": "错误"}
    else:
        options = {}
        for c in OPTION_COLS:
            val = re.sub(r"^[A-Ha-h]\s*[.、．:：)）]\s*", "", _cell(row, mapping, c))
            if val:
                options[c] = val
        if not options:
            # 导出文件没给选项内容时，只按答案字母建占位，界面会明确提示
            options = {c: "（选项内容未提供：" + c + "）" for c in sorted(set(answer))}
            caution = (caution + "；" if caution else "") + "导出文件缺选项内容，此处仅答案可用"
    for ch in answer:
        if ch not in options:
            options[ch] = "（选项内容未提供：" + ch + "）"
            caution = (caution + "；" if caution else "") + "导出文件缺选项内容，此处仅答案可用"

    return {
        "id": "",
        "type": qtype,
        "stem": stem,
        "options": options,
        "answer": answer,
        "basis": _cell(row, mapping, "basis"),
        "law": _cell(row, mapping, "law"),
        "source": _cell(row, mapping, "source") or (source + " / " + sheet_name),
        "caution": caution,
    }


def import_xlsx(path):
    """解析 Excel，返回 (题目列表, 逐表诊断文本列表)。"""
    questions, detail = [], []
    for sheet_name, rows in read_xlsx(path):
        if not rows:
            detail.append("%s: 空表" % sheet_name)
            continue
        mapping = resolve_header(rows[0])
        if "stem" not in mapping or "answer" not in mapping:
            detail.append("%s: 表头未识别，跳过（首行：%s）" % (sheet_name, rows[0][:4]))
            continue
        hint = sheet_hint(sheet_name)
        before = len(questions)
        for row in rows[1:]:
            q = from_row(row, mapping, hint, pathlib.Path(path).name, sheet_name)
            if q:
                questions.append(q)
        detail.append("%s: 解析 %d 题（题型=%s）" % (
            sheet_name, len(questions) - before, hint or "按列"))
    return questions, detail


def read_json_questions(path):
    payload = json.loads(pathlib.Path(path).read_text(encoding="utf-8"))
    items = payload.get("questions", payload) if isinstance(payload, dict) else payload
    out = []
    for i, q in enumerate(items, 1):
        if not isinstance(q, dict) or not q.get("stem"):
            continue
        qtype = norm_type(q.get("type")) or "single"
        options = q.get("options") or {}
        if qtype == "judge" and not options:
            options = {"A": "正确", "B": "错误"}
        answer = norm_answer(q.get("answer"), qtype)
        if not answer:
            continue
        out.append({
            "id": q.get("id") or "X%04d" % i,
            "type": qtype,
            "stem": str(q["stem"]).strip(),
            "options": options,
            "answer": answer,
            "basis": str(q.get("basis") or "").strip(),
            "law": str(q.get("law") or "").strip(),
            "source": str(q.get("source") or pathlib.Path(path).name).strip(),
            "caution": str(q.get("caution") or "").strip(),
        })
    return out


# ---------- 合并与落盘 ----------

def norm_key(stem):
    return re.sub(r"[^\u4e00-\u9fffA-Za-z0-9]", "", stem)


def merge_questions(existing, imported):
    """按规范化题干去重；existing 优先，保留已有依据与人工校正。"""
    seen, merged = set(), []
    for q in list(existing) + list(imported):
        key = norm_key(q.get("stem", ""))
        if not key or key in seen:
            continue
        seen.add(key)
        merged.append(dict(q))
    for i, q in enumerate(merged, 1):
        q["id"] = "Q%04d" % i
    return merged


def save_bank(path, questions, backup=True):
    path = pathlib.Path(path)
    if backup and path.exists():
        bak = path.with_suffix(path.suffix + ".bak")
        bak.write_bytes(path.read_bytes())
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "version": "1.2",
        "updated": "2026-11-19",
        "scope": "桂司通〔2026〕号 附件3（382部）范围内公开真题 + 本地 Word 题库 + 导入补充",
        "total": len(questions),
        "note": "题目与答案取自公开题库、本地 Word 题库及导入文件；basis 为法条依据或题库解析；caution 非空时表示需核对。",
        "questions": questions,
    }
    # 原子写：先落临时文件再替换，避免写坏后无法回滚
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(path)
    return path


def _selftest():
    if len(sys.argv) < 2:
        print("用法：python3 src/xlsx_import.py <file.xlsx>")
        return 1
    qs, detail = import_xlsx(sys.argv[1])
    print("\n".join(detail))
    print("共 %d 题" % len(qs))
    for q in qs:
        print("  [%s] %s => %s" % (q["type"], q["stem"][:40], q["answer"]))
    return 0


if __name__ == "__main__":
    sys.exit(_selftest())
