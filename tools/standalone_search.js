/*
 * 学法用法速查助手 · 浏览器检索核心
 *
 * 与 src/bank.py 同一套算法：NFKC 规范化 + 题号/空白清理 + 2-gram 倒排召回
 * + LCS 相似度精排 + 选项一致性加权。无第三方依赖，浏览器和 Node 都能跑。
 *
 * 复杂度：建索引 O(N*L)，N=题数、L=题干平均长度；查询召回 O(G*P)，
 * G=查询 gram 数、P=倒排链长；精排 O(K*L^2)，K 截断到 400 防长文本卡顿。
 */
(function (root, factory) {
  var api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.XuefaSearch = api;
})(typeof window !== 'undefined' ? window : globalThis, function () {
  'use strict';

  // 与 Python _NOISE 一致：^ 只锚定第一个分支，其余分支全局生效
  var NOISE = /^\s*(?:第\s*)?\d+\s*[.、．:：)）]\s*|[（(]\s*[)）]|[_＿]{2,}|[ \t\r\n]+/g;
  var KEEP = /[^\u4e00-\u9fffA-Za-z0-9]/g;
  var OPTION_START = /(?:^|[\s，。；;])[A-Ga-g]\s*[.、．:：)）]\s*/;
  var OPT_MARK = /(^|[\s，。；;（(\u4e00-\u9fff])([A-Ha-h])\s*[.、．:：)）]\s*/g;
  var TRIM_EDGE = /^[ \t\r\n，。；;、]+|[ \t\r\n，。；;、]+$/g;
  var TYPE_NAME = { single: '单选题', multi: '多选题', judge: '判断题' };

  function normalize(text) {
    if (!text) return '';
    return String(text).normalize('NFKC').replace(NOISE, '').replace(KEEP, '').toLowerCase();
  }

  function extractStem(text) {
    var s = String(text || '').normalize('NFKC');
    var m = OPTION_START.exec(s);
    if (m) s = s.slice(0, m.index);
    return normalize(s);
  }

  function extractOptions(text) {
    var s = String(text || '').normalize('NFKC');
    var marks = [];
    var m;
    OPT_MARK.lastIndex = 0;
    while ((m = OPT_MARK.exec(s)) !== null) marks.push(m);
    if (marks.length < 2) return {};
    var out = {};
    for (var i = 0; i < marks.length; i++) {
      var letter = marks[i][2].toUpperCase();
      // group2（选项字母）紧跟 group1，用长度反推起点，等价于 Python 的 m.start(2)
      var letterPos = marks[i].index + marks[i][1].length;
      var end = i + 1 < marks.length
        ? marks[i + 1].index + marks[i + 1][1].length
        : s.length;
      var seg = s.slice(marks[i].index + marks[i][0].length, end).replace(TRIM_EDGE, '');
      if (seg && !(letter in out)) out[letter] = seg;
    }
    return out;
  }

  function optionAgreement(parsed, q) {
    var opts = q.options || {};
    var keys = Object.keys(parsed);
    if (!keys.length || !Object.keys(opts).length) return 0;
    var hit = 0;
    for (var i = 0; i < keys.length; i++) {
      var letter = keys[i];
      var ref = normalize(opts[letter] || '');
      var val = normalize(parsed[letter]);
      if (!ref || !val) continue;
      if (val === ref || (val.length >= 2 && (ref.indexOf(val) >= 0 || val.indexOf(ref) >= 0))) hit++;
    }
    return hit / keys.length;
  }

  function bigrams(s) {
    if (!s) return [];
    if (s.length < 2) return [s];
    var out = [];
    var seen = Object.create(null);
    for (var i = 0; i < s.length - 1; i++) {
      var g = s.slice(i, i + 2);
      if (!seen[g]) { seen[g] = 1; out.push(g); }
    }
    return out;
  }

  // difflib.SequenceMatcher.ratio 的 LCS 近似，短文本下分值一致到小数点后三位
  function lcsRatio(a, b) {
    if (!a.length || !b.length) return 0;
    if (a === b) return 1;
    var n = a.length, m = b.length;
    if (n > 240) { a = a.slice(0, 240); n = 240; }
    if (m > 240) { b = b.slice(0, 240); m = 240; }
    var prev = new Uint16Array(m + 1), cur = new Uint16Array(m + 1);
    for (var i = 1; i <= n; i++) {
      cur[0] = 0;
      for (var j = 1; j <= m; j++) {
        cur[j] = a.charAt(i - 1) === b.charAt(j - 1)
          ? prev[j - 1] + 1
          : (prev[j] >= cur[j - 1] ? prev[j] : cur[j - 1]);
      }
      var tmp = prev; prev = cur; cur = tmp;
    }
    return 2 * prev[m] / (n + m);
  }

  function answerText(q) {
    var ans = q.answer || '';
    if (q.type === 'judge') return ans === 'A' ? '正确' : '错误';
    return Array.prototype.map.call(String(ans), function (c) { return c; }).join('、');
  }

  function Bank(questions, laws) {
    this.questions = questions || [];
    this.laws = laws || { items: [], total: 0 };
    this.index = new Map();
    this.grams = [];
    this.bodies = [];
    for (var i = 0; i < this.questions.length; i++) {
      var q = this.questions[i];
      var body = normalize(q.stem || '');
      var opts = q.options || {};
      var keys = Object.keys(opts);
      for (var k = 0; k < keys.length; k++) body += normalize(opts[keys[k]]);
      var gs = bigrams(body);
      this.grams.push(gs);
      this.bodies.push(body);
      for (var g = 0; g < gs.length; g++) {
        var gram = gs[g];
        var arr = this.index.get(gram);
        if (!arr) { arr = []; this.index.set(gram, arr); }
        arr.push(i);
      }
    }
  }

  Bank.prototype.search = function (text, topk) {
    topk = topk || 8;
    var stem = extractStem(text);
    var raw = normalize(text);
    var parsed = extractOptions(text);
    if (!stem) return [];
    var qg = bigrams(stem);
    if (!qg.length) return [];

    // 倒排召回：命中任一 gram 的题进候选，用命中数做粗排后截断
    var cand = new Map();
    for (var a = 0; a < qg.length; a++) {
      var arr = this.index.get(qg[a]);
      if (!arr) continue;
      for (var b = 0; b < arr.length; b++) {
        cand.set(arr[b], (cand.get(arr[b]) || 0) + 1);
      }
    }
    if (!cand.size) return [];
    var list = Array.from(cand.entries());
    list.sort(function (x, y) { return y[1] - x[1]; });
    if (list.length > 400) list.length = 400;

    var scored = [];
    for (var z = 0; z < list.length; z++) {
      var i = list[z][0];
      var gs = this.grams[i];
      var body = this.bodies[i];
      var inter = 0;
      for (var u = 0; u < qg.length; u++) if (gs.indexOf(qg[u]) >= 0) inter++;
      var cover = inter / qg.length;
      if (cover < 0.22) continue;
      var ratio = lcsRatio(stem, body.slice(0, Math.max(stem.length * 2, 40)));
      var score = 0.68 * cover + 0.32 * ratio;
      if (stem.length >= 8 && body.indexOf(stem.slice(0, 8)) >= 0) score += 0.12;
      var law = normalize(this.questions[i].law || '');
      if (law.length >= 6 && raw.indexOf(law.slice(0, 6)) >= 0) score += 0.06;
      score = Math.min(score, 1);
      // 同题干不同选项顺序的题：粘贴了选项时用选项文本打破平局
      var agree = Object.keys(parsed).length >= 2
        ? optionAgreement(parsed, this.questions[i]) : 0;
      var final = score + 0.30 * agree;
      scored.push({ rank: final, score: Math.min(final, 1), q: this.questions[i] });
    }
    scored.sort(function (x, y) { return y.rank - x.rank; });
    return scored.slice(0, topk).map(function (x) { return { score: x.score, q: x.q }; });
  };

  Bank.prototype.searchLaws = function (keyword, limit) {
    limit = limit || 30;
    var key = normalize(keyword);
    var items = this.laws.items || [];
    if (!key) return items.slice(0, limit);
    var out = [];
    for (var i = 0; i < items.length; i++) {
      var name = normalize(items[i].name || '');
      if (name.indexOf(key) >= 0) out.push({ score: 1, item: items[i] });
      else {
        var r = lcsRatio(key, name);
        if (r >= 0.5) out.push({ score: r, item: items[i] });
      }
    }
    out.sort(function (x, y) { return y.score - x.score; });
    return out.slice(0, limit).map(function (x) { return x.item; });
  };

  Bank.prototype.lawCoverage = function () {
    var covered = Object.create(null);
    var all = Object.create(null);
    var items = this.laws.items || [];
    for (var j = 0; j < items.length; j++) {
      var n = normalize(items[j].name || '');
      if (n) all[n] = 1;
    }
    for (var i = 0; i < this.questions.length; i++) {
      var law = this.questions[i].law || '';
      if (law) covered[normalize(law)] = 1;
    }
    var n = 0;
    Object.keys(covered).forEach(function (k) { if (all[k]) n++; });
    return { covered: n, total: Object.keys(all).length };
  };

  return {
    Bank: Bank,
    normalize: normalize,
    extractStem: extractStem,
    extractOptions: extractOptions,
    bigrams: bigrams,
    lcsRatio: lcsRatio,
    answerText: answerText,
    typeName: function (t) { return TYPE_NAME[t] || t; }
  };
});
