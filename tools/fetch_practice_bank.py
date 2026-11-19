# -*- coding: utf-8 -*-
"""
广西普法网「我的练习」题库采集器。

原理
----
练习页前端调用两个公开接口：
  GET /api/exam/api/v1/practice/getQuestion?quesType=1|2|3&excludeIds=...
      随机返回一道题，返回体里已经带 answer（参考答案）与 answerAnalysis（题目解析）。
      页面上「随便选一个 → 确定」只是把前端已有的这两个字段显示出来，不需要真的提交。
  GET /api/exam/question/question/{id}
      按题目 ID 取同一份完整题目对象。ID 分段连续，可直接扫描，绕开 excludeIds 的 URL 长度限制。

鉴权：Authorization: Bearer <localStorage["Exam-Token"]>，Clientid 为站点固定值。
脚本自动从本机 Edge / Chrome 的 Local Storage 中提取当前已登录会话的 token，
也可以 --token / --token-file 手工传入。token 不写入任何输出文件。

采集流程
--------
1. probe      验证 token 与接口可用
2. seed       随机采样若干题，收集题目对象与 ID 种子
3. cluster    把 ID 按间隔聚成连续段
4. scan       对每段外扩后并发按 ID 扫描，补齐段内全部题目（含采样未命中的）
5. convert    转成 xuefa_helper 的 questions.json 字段
6. merge      官方题库优先 + 现有题库补充，按规范化题干去重，原子写回并留 .bak

用法
----
  python3 tools/fetch_practice_bank.py                       # 全自动，写回 src/data/questions.json
  python3 tools/fetch_practice_bank.py --probe-only          # 只验证会话与接口
  python3 tools/fetch_practice_bank.py --sample 800 --workers 8
  python3 tools/fetch_practice_bank.py --token eyJhbGci...   # 手工给 token
  python3 tools/fetch_practice_bank.py --no-merge --out dist/questions.json
  python3 tools/fetch_practice_bank.py --resume              # 用缓存续跑，跳过已抓 ID

前置：在 Edge / Chrome 里登录过 https://gxpf.sft.gxzf.gov.cn/portal/exam/xfzl/index
依赖：requests（第三方，仅采集脚本需要；exe 运行时仍是零依赖）
"""
from __future__ import annotations

import argparse
import base64
import collections
import glob
import io
import json
import os
import pathlib
import re
import ssl
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.ssl_ import create_urllib3_context

ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
import xlsx_import  # noqa: E402  复用 norm_key / merge_questions 的去重规则

FROZEN = getattr(sys, "frozen", False)


def runtime_dir():
    """打包成 exe 后所有产物落在 exe 同目录，方便同事直接用。"""
    if FROZEN:
        return pathlib.Path(sys.executable).resolve().parent
    return ROOT


def laws_path():
    if FROZEN:
        return pathlib.Path(getattr(sys, "_MEIPASS", runtime_dir())) / "data" / "laws_scope.json"
    return ROOT / "src" / "data" / "laws_scope.json"

DEFAULT_BASE = "https://gxpf.sft.gxzf.gov.cn"
CLIENT_ID = "e5cd7e4891bf95d1d19206ce24a7b32e"
ORIGIN_TAG = b"gxpf.sft.gxzf.gov.cn"
TYPE_MAP = {"1": "single", "2": "multi", "3": "judge"}
JUDGE_TRUE = {"对", "正确", "A", "TRUE", "T", "√", "✓", "1", "是", "YES", "Y"}
JUDGE_FALSE = {"错", "错误", "B", "FALSE", "F", "×", "✕", "X", "0", "否", "NO", "N"}
CLEAN_RE = re.compile(r"[\u00a0\u3000\r\n\t]+")
NON_KEY_RE = re.compile(r"[^\u4e00-\u9fffA-Za-z0-9]")


# ---------------------------------------------------------------- TLS 适配
# 站点网关仍用旧版 SSL 重协商，OpenSSL 3 默认拒绝；打开 OP_LEGACY_SERVER_CONNECT 才能连。
class _LegacyTLSAdapter(HTTPAdapter):
    @staticmethod
    def _ctx(kw):
        ctx = create_urllib3_context()
        ctx.options |= getattr(ssl, "OP_LEGACY_SERVER_CONNECT", 0x40000)
        kw["ssl_context"] = ctx
        return kw

    def init_poolmanager(self, *args, **kw):
        return super().init_poolmanager(*args, **self._ctx(kw))

    def proxy_manager_for(self, *args, **kw):
        return super().proxy_manager_for(*args, **self._ctx(kw))


def _session():
    s = requests.Session()
    s.mount("https://", _LegacyTLSAdapter())
    return s


# ---------------------------------------------------------------- token 发现
def _storage_files():
    la = os.environ.get("LOCALAPPDATA", "")
    if not la:
        return []
    patterns = [
        os.path.join(la, "Microsoft", "Edge", "User Data", "*", "Local Storage", "leveldb", "*"),
        os.path.join(la, "Google", "Chrome", "User Data", "*", "Local Storage", "leveldb", "*"),
    ]
    out = []
    for pat in patterns:
        for f in glob.glob(pat):
            name = os.path.basename(f)
            if name in ("LOCK", "CURRENT") or name.endswith((".log", ".ldb")):
                if os.path.isfile(f):
                    out.append(f)
    # 最近修改的优先，当前会话一般落在 .log 里
    out.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return out


TOKEN_RE = re.compile(rb'"token"\s*:\s*"(eyJ[A-Za-z0-9_.\-]{40,})"')
ANY_JWT_RE = re.compile(rb"(eyJ[A-Za-z0-9_\-]{10,}\.[A-Za-z0-9_\-]{20,}\.[A-Za-z0-9_\-]{20,})")


def find_tokens():
    """从浏览器 Local Storage 中提取考试站的会话 JWT，去重后按新鲜度返回。"""
    found, seen = [], set()
    for path in _storage_files():
        try:
            data = open(path, "rb").read()
        except OSError:
            continue
        if ORIGIN_TAG not in data:
            continue
        cands = [m.group(1) for m in TOKEN_RE.finditer(data)]
        if not cands:
            # 退路：同文件里任意 JWT（有些版本把 token 直接作为 value 存）
            cands = [m.group(1) for m in ANY_JWT_RE.finditer(data)]
        for raw in cands:
            tok = raw.decode("ascii", "ignore")
            if tok and tok not in seen:
                seen.add(tok)
                found.append(tok)
    return found


def jwt_exp_hint(token):
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        obj = json.loads(base64.urlsafe_b64decode(payload))
        exp = obj.get("exp")
        if not exp:
            return "无 exp 字段"
        left = int(exp) - int(time.time())
        return "剩余 %d 秒" % left if left > 0 else "已过期 %d 秒" % (-left)
    except Exception:
        return "非标准 JWT 或无法解析"


# ---------------------------------------------------------------- 接口封装
class Api:
    def __init__(self, base, token, clientid=CLIENT_ID, timeout=20):
        self.base = base.rstrip("/")
        self.count_url = self.base + "/api/exam/api/v1/practice/count"
        self.random_url = self.base + "/api/exam/api/v1/practice/getQuestion"
        self.byid_url = self.base + "/api/exam/question/question/{}"
        self.headers = {
            "Authorization": "Bearer " + token,
            "Clientid": clientid,
            "Accept": "application/json",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0",
        }
        self.timeout = timeout
        self._local = threading.local()
        self.requests = 0
        self.failures = 0
        self._lock = threading.Lock()

    @property
    def session(self):
        s = getattr(self._local, "s", None)
        if s is None:
            s = _session()
            self._local.s = s
        return s

    def _bump(self, ok):
        with self._lock:
            self.requests += 1
            if not ok:
                self.failures += 1

    def get(self, url, params=None, retries=3):
        last = None
        for attempt in range(retries):
            try:
                r = self.session.get(url, params=params, headers=self.headers, timeout=self.timeout)
                if r.status_code != 200:
                    last = "HTTP %d" % r.status_code
                    self._bump(False)
                    time.sleep(0.4 * (attempt + 1))
                    continue
                j = r.json()
                self._bump(True)
                return j
            except Exception as exc:  # 网络抖动 / JSON 解析失败，退避重试
                last = str(exc)[:120]
                self._bump(False)
                time.sleep(0.4 * (attempt + 1))
        if last:
            print("  [warn] 请求失败 %s : %s" % (url.split("/")[-1], last), file=sys.stderr)
        return None

    def probe(self):
        j = self.get(self.count_url)
        if not j or j.get("code") != 200:
            return None
        return j.get("data") or {}

    def random_question(self, qtype, exclude_ids):
        j = self.get(self.random_url, params={"quesType": qtype, "excludeIds": ",".join(exclude_ids)})
        if not j or j.get("code") != 200:
            return None
        d = j.get("data") or {}
        return d if d.get("id") else None

    def by_id(self, qid):
        j = self.get(self.byid_url.format(qid))
        if not j or j.get("code") != 200:
            return None
        d = j.get("data") or {}
        return d if d.get("id") else None


# ---------------------------------------------------------------- 采集
def phase_seed(api, n, delay, cache, cache_path):
    """随机采样发现 ID 种子；顺手把题目对象存进缓存。"""
    recent = []
    t0 = time.time()
    for i in range(n):
        qtype = ("1", "2", "3")[i % 3]
        d = api.random_question(qtype, recent[-250:])  # 窗口 250，避免 URL 超限
        if d:
            qid = str(d["id"])
            cache[qid] = d
            recent.append(qid)
        if delay:
            time.sleep(delay)
        if (i + 1) % 100 == 0:
            print("  seed %d/%d | 已收集 %d 题" % (i + 1, n, len(cache)), flush=True)
    _flush_cache(cache, cache_path)
    print("  采样完成：%d 次请求 -> 缓存 %d 题，耗时 %.1fs" % (n, len(cache), time.time() - t0))
    return cache


def cluster_segments(ids, gap=500):
    """把排序后的 ID 按最大间隔聚成连续段。"""
    segs = []
    for x in sorted(ids):
        if not segs or x - segs[-1][1] > gap:
            segs.append([x, x])
        else:
            segs[-1][1] = x
    return [(lo, hi) for lo, hi in segs]


def build_scan_ids(segs, expand, known, max_scan):
    out, seen_local = [], set()
    for lo, hi in segs:
        a, b = max(1, lo - expand), hi + expand
        for qid in range(a, b + 1):
            if qid in known or qid in seen_local:
                continue
            seen_local.add(qid)
            out.append(qid)
            if len(out) >= max_scan:
                return out, True
    return out, False


def phase_scan(api, scan_ids, cache, cache_path, workers):
    """并发按 ID 扫描，命中的题目写回缓存；每 400 条落一次盘，支持 --resume。"""
    t0 = time.time()
    done = 0
    hits = 0
    lock = threading.Lock()
    pending = {}

    def work(qid):
        return qid, api.by_id(qid)

    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(work, q) for q in scan_ids]
        for fut in as_completed(futures):
            qid, d = fut.result()
            done += 1
            with lock:
                if d:
                    pending[str(d["id"])] = d
                    hits += 1
                if len(pending) >= 400:
                    cache.update(pending)
                    _flush_cache(cache, cache_path)
                    pending.clear()
            if done % 1000 == 0:
                print("  scan %d/%d | 新增命中 %d | %.1fs" % (done, len(scan_ids), hits, time.time() - t0), flush=True)
    with lock:
        if pending:
            cache.update(pending)
    _flush_cache(cache, cache_path)
    print("  扫描完成：%d 个 ID -> 命中 %d 题，缓存合计 %d 题，耗时 %.1fs"
          % (len(scan_ids), hits, len(cache), time.time() - t0))
    return cache


def _flush_cache(cache, path):
    path = pathlib.Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
    tmp.replace(path)


# ---------------------------------------------------------------- 转换
def clean_text(v):
    s = CLEAN_RE.sub(" ", str(v or "")).strip()
    return re.sub(r"\s{2,}", " ", s)


# 少量官方题把 B、C 选项内联在 A 的 label 里，如「…个人财产B、《二生》… C、《三生》…」
_INLINE_OPT_RE = re.compile(r"[\s]*([B-H])\s*[、.．:：]\s*")


def expand_inline_options(options):
    out = {}
    for key, label in options.items():
        parts = _INLINE_OPT_RE.split(str(label))
        if len(parts) == 1:
            out[key] = label
            continue
        head = parts[0].strip()
        if head:
            out[key] = head
        for i in range(1, len(parts) - 1, 2):
            letter, text = parts[i], parts[i + 1].strip(" \t，。；;、")
            if letter not in out and text:
                out[letter] = text
    return out


def build_law_index(path):
    data = json.loads(pathlib.Path(path).read_text(encoding="utf-8"))
    idx = []
    for it in data.get("items", []):
        name = str(it.get("name", ""))
        key = NON_KEY_RE.sub("", name)
        if key:
            idx.append((key, name))
    return idx


def match_law(desc, law_index):
    """questionDesc 形如《广西壮族自治区禁毒条例》，匹配附件3清单里的规范名称。"""
    n = NON_KEY_RE.sub("", clean_text(desc))
    if not n:
        return ""
    exact = [name for key, name in law_index if key == n]
    if exact:
        return exact[0]
    contained = [(key, name) for key, name in law_index if key in n or (len(n) >= 6 and n in key)]
    if contained:
        return max(contained, key=lambda x: len(x[0]))[1]
    return ""


def to_question(raw, law_index):
    qtype = TYPE_MAP.get(str(raw.get("questionType") or "").strip())
    if not qtype:
        return None
    stem = clean_text(raw.get("questionTitle"))
    if not stem:
        return None

    try:
        opts_raw = json.loads(raw.get("questionOptions") or "[]")
    except (ValueError, TypeError):
        opts_raw = []
    options = {}
    if qtype == "judge":
        options = {"A": "正确", "B": "错误"}
    else:
        for o in opts_raw:
            if not isinstance(o, dict):
                continue
            value = clean_text(o.get("value")).upper()[:1]
            label = clean_text(o.get("label"))
            if value in "ABCDEFGH" and label and value not in options:
                options[value] = label

    ans = clean_text(raw.get("answer")).upper()
    ans = re.sub(r"[\s,，、/|;；]+", "", ans)
    if qtype == "judge":
        if ans in JUDGE_TRUE:
            answer = "A"
        elif ans in JUDGE_FALSE:
            answer = "B"
        else:
            return None
    else:
        answer = "".join(ch for ch in ans if ch in "ABCDEFGH")
        if not answer:
            return None
        if qtype == "single" and len(answer) > 1:
            answer = answer[0]

    # 答案字母缺位时尝试拆开内联选项；仍缺位的题保留但标注，避免悄悄给错字母
    if qtype != "judge" and options and any(ch not in options for ch in answer):
        options = expand_inline_options(options)

    desc = clean_text(raw.get("questionDesc"))
    law = match_law(desc, law_index)
    basis = clean_text(raw.get("answerAnalysis"))
    if not basis:
        basis = desc  # 没有解析时至少给出法规出处；用户已确认「找不到就跳过」，留空也不影响检索
    notes = []
    if not law and desc:
        notes.append("法规出处《%s》不在附件3清单中，按题目原文保留" % desc.strip("《》"))
    if qtype != "judge" and not options:
        notes.append("接口未返回选项文本，仅答案与依据可用")
    if qtype != "judge" and options:
        missing = [ch for ch in answer if ch not in options]
        if missing:
            notes.append("答案含接口未提供的选项字母 %s，请核对" % "".join(missing))
    caution = "；".join(notes)

    return {
        "id": "",
        "type": qtype,
        "stem": stem,
        "options": options,
        "answer": answer,
        "basis": basis,
        "law": law or desc,
        "source": "广西普法云平台 我的练习（官方题库 ID %s）" % raw.get("id"),
        "caution": caution,
    }


def write_bank(path, questions, scope_note):
    path = pathlib.Path(path)
    if path.exists():
        path.with_suffix(path.suffix + ".bak").write_bytes(path.read_bytes())
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "version": "1.3",
        "updated": time.strftime("%Y-%m-%d"),
        "scope": scope_note,
        "total": len(questions),
        "note": "官方练习题库（含参考答案与题目解析）优先；旧公开题库与本地 Word 题库作为补充。",
        "questions": questions,
    }
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(path)
    return path


# ---------------------------------------------------------------- 主流程
def main(argv=None):
    p = argparse.ArgumentParser(description="广西普法网「我的练习」题库采集器")
    p.add_argument("--token", default="", help="手工指定 Exam-Token（JWT）")
    p.add_argument("--token-file", default="", help="从文本文件读取 token（第一行）")
    p.add_argument("--base-url", default=DEFAULT_BASE, help="站点根地址")
    p.add_argument("--clientid", default=CLIENT_ID, help="请求头 Clientid")
    p.add_argument("--sample", type=int, default=600, help="随机采样次数，用于发现 ID 段（默认 600）")
    p.add_argument("--expand", type=int, default=400, help="每个 ID 段向两侧外扩的扫描宽度（默认 400）")
    p.add_argument("--gap", type=int, default=500, help="ID 间隔超过该值视为另一段（默认 500）")
    p.add_argument("--max-scan", type=int, default=40000, help="单次最多扫描多少个 ID（安全阀）")
    p.add_argument("--workers", type=int, default=6, help="按 ID 扫描的并发数（默认 6）")
    p.add_argument("--delay", type=float, default=0.02, help="采样阶段每次请求间隔秒数")
    default_out = str(runtime_dir() / "questions.json") if FROZEN else str(ROOT / "src" / "data" / "questions.json")
    p.add_argument("--out", default=default_out, help="输出题库路径")
    p.add_argument("--no-merge", action="store_true", help="不合并现有题库，只输出本次采集结果")
    default_official = (str(runtime_dir() / "official_questions.json") if FROZEN
                        else str(ROOT / "tools" / "official_questions.json"))
    default_cache = (str(runtime_dir() / "_raw_questions_cache.json") if FROZEN
                     else str(ROOT / "tools" / "cache" / "raw_questions.json"))
    p.add_argument("--official-out", default=default_official,
                   help="单独保存一份纯官方题库 JSON")
    p.add_argument("--cache", default=default_cache,
                   help="原始题目缓存（断点续传）")
    p.add_argument("--report", default=str(runtime_dir() / ("题库采集报告.txt" if FROZEN else "tools/fetch_report.txt")),
                   help="采集报告路径")
    p.add_argument("--resume", action="store_true", help="使用已有缓存并跳过已抓过的 ID")
    p.add_argument("--skip-scan", action="store_true", help="只做采样，不做 ID 段扫描")
    p.add_argument("--probe-only", action="store_true", help="只验证 token 与接口")
    args = p.parse_args(argv)

    t0 = time.time()

    # 1. token
    tokens = []
    if args.token:
        tokens.append(args.token.strip())
    if args.token_file:
        tokens.append(pathlib.Path(args.token_file).read_text(encoding="utf-8").splitlines()[0].strip())
    if not tokens:
        tokens = find_tokens()
    if not tokens:
        print("未找到会话 token。请先在 Edge / Chrome 登录考试网站，或用 --token 传入。", file=sys.stderr)
        return 2
    print("候选 token：%d 个" % len(tokens))

    api = None
    for tok in tokens:
        cand = Api(args.base_url, tok, args.clientid)
        info = cand.probe()
        if info is not None:
            api = cand
            print("会话有效（%s），userId=%s" % (jwt_exp_hint(tok), info.get("userId")))
            break
        print("  跳过一个失效 token", file=sys.stderr)
    if api is None:
        print("所有候选 token 均失效，请重新登录考试网站后重试。", file=sys.stderr)
        return 3
    if args.probe_only:
        d = api.random_question("1", [])
        print("接口取题：%s" % ("OK, ID=%s" % d["id"] if d else "FAIL"))
        return 0 if d else 4

    # 2. 缓存 / 采样
    cache_path = pathlib.Path(args.cache)
    cache = {}
    if args.resume and cache_path.exists():
        try:
            cache = json.loads(cache_path.read_text(encoding="utf-8"))
            print("续跑：载入缓存 %d 题" % len(cache))
        except ValueError:
            cache = {}
    print("[1/4] 随机采样 %d 次…" % args.sample)
    cache = phase_seed(api, args.sample, args.delay, cache, cache_path)

    seeds = [int(x) for x in cache.keys()]
    segs = cluster_segments(seeds, args.gap)
    print("[2/4] ID 段：%s" % ", ".join("%d-%d" % s for s in segs))

    # 3. 段扫描
    if not args.skip_scan:
        known = set(seeds)
        scan_ids, truncated = build_scan_ids(segs, args.expand, known, args.max_scan)
        print("[3/4] 扫描 %d 个候选 ID（并发 %d）%s"
              % (len(scan_ids), args.workers, "，已触发 max-scan 截断" if truncated else ""))
        cache = phase_scan(api, scan_ids, cache, cache_path, args.workers)
    else:
        truncated = False
        print("[3/4] 按 --skip-scan 跳过段扫描")

    # 4. 转换 / 合并 / 输出
    law_index = build_law_index(laws_path())
    official, dropped = [], collections.Counter()
    for raw in cache.values():
        q = to_question(raw, law_index)
        if q is None:
            dropped["字段缺失/无法规范化"] += 1
            continue
        official.append(q)
    official = xlsx_import.merge_questions([], official)  # 按规范化题干去重并重编 ID
    print("[4/4] 有效官方题 %d 道（丢弃 %d）" % (len(official), sum(dropped.values())))

    pathlib.Path(args.official_out).write_text(
        json.dumps({"total": len(official), "questions": official}, ensure_ascii=False, indent=1),
        encoding="utf-8")

    out = pathlib.Path(args.out)
    existing = []
    if not args.no_merge and out.exists():
        try:
            existing = json.loads(out.read_text(encoding="utf-8")).get("questions", [])
        except ValueError:
            existing = []
    merged = official if args.no_merge else xlsx_import.merge_questions(official, existing)
    scope = ("桂司通〔2026〕号 附件3（382部）｜广西普法云平台我的练习官方题库 %d 题"
             % len(official))
    if existing and not args.no_merge:
        scope += " + 原有补充题库 %d 题" % len(existing)
    write_bank(out, merged, scope)

    types = collections.Counter(q["type"] for q in official)
    has_analysis = sum(1 for raw in cache.values() if clean_text(raw.get("answerAnalysis")))
    law_names = {name for _, name in law_index}
    lines = [
        "广西普法网「我的练习」题库采集报告",
        "生成时间：%s" % time.strftime("%Y-%m-%d %H:%M:%S"),
        "站点：%s" % args.base_url,
        "会话：有效（token 未写入本报告）",
        "",
        "采样请求：%d 次" % args.sample,
        "ID 段：%s" % ", ".join("%d-%d" % s for s in segs),
        "原始缓存题目：%d 道" % len(cache),
        "有效官方题目：%d 道（丢弃 %d）" % (len(official), sum(dropped.values())),
        "题型：单选 %d / 多选 %d / 判断 %d" % (types["single"], types["multi"], types["judge"]),
        "带题目解析（接口原生 answerAnalysis）：%d 道" % has_analysis,
        "精确关联附件3法规：%d 道" % sum(1 for q in official if q["law"] in law_names),
        "段扫描截断：%s" % ("是（达到 --max-scan）" if truncated else "否"),
        "总请求：%d 次，失败重试累计：%d 次" % (api.requests, api.failures),
        "",
        "原有题库：%d 道" % len(existing),
        "合并去重后：%d 道" % len(merged),
        "输出题库：%s" % out,
        "纯官方题库：%s" % args.official_out,
        "原始缓存：%s" % cache_path,
        "备份：%s.bak" % out,
        "耗时：%.1f 秒" % (time.time() - t0),
    ]
    report = pathlib.Path(args.report)
    report.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("\n".join(lines))
    print("\nDONE -> %d 题写入 %s" % (len(merged), out))
    if FROZEN and sys.stdin and sys.stdin.isatty():
        try:
            input("\n按回车键退出…")
        except (EOFError, OSError):
            pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
