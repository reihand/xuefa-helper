# -*- coding: utf-8 -*-
"""
UOS / 统信 arm64 版界面：本地浏览器页面 + 标准库 HTTP 服务。

只依赖 Python 3.7+ 标准库，不依赖 tkinter，不联网。
启动：python3 src/webui.py [--port 8765] [--no-browser]
"""
import argparse
import json
import os
import sys
import tempfile
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, unquote, urlparse

import bank as bankmod
import paths
import xlsx_import as ximp

BANK = None
BANK_LOCK = threading.RLock()

PAGE = """<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>学法用法速查助手 · 离线</title>
<style>
:root{--bg:#f5f7fa;--card:#fff;--line:#dfe4ea;--text:#1f2933;--muted:#6b7280;--green:#0a7d3f;--orange:#b45309;--blue:#1d4ed8}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);font:15px/1.6 -apple-system,"Noto Sans CJK SC","Microsoft YaHei",sans-serif}
header{background:#fff;border-bottom:1px solid var(--line);padding:14px 20px;position:sticky;top:0;z-index:5}
h1{font-size:19px;margin:0 0 2px}
.sub{color:var(--muted);font-size:13px}
.wrap{max-width:1180px;margin:0 auto;padding:18px}
.tabs{display:flex;gap:8px;margin-bottom:14px}
.tab{border:1px solid var(--line);background:#fff;border-radius:8px;padding:7px 16px;cursor:pointer;color:var(--text)}
.tab.on{background:var(--blue);color:#fff;border-color:var(--blue)}
.panel{display:none}.panel.on{display:block}
.searchbar{display:flex;gap:8px;margin-bottom:12px}
input[type=text],textarea{flex:1;border:1px solid var(--line);border-radius:8px;padding:10px 12px;font:inherit;background:#fff;color:var(--text)}
input[type=file]{font:inherit}
button{border:1px solid var(--line);background:#fff;border-radius:8px;padding:9px 16px;cursor:pointer;color:var(--text)}
button.primary{background:var(--blue);color:#fff;border-color:var(--blue)}
.grid{display:grid;grid-template-columns:minmax(280px,42%) 1fr;gap:14px}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px;min-height:300px}
.list{list-style:none;margin:0;padding:0;max-height:70vh;overflow:auto}
.list li{border-bottom:1px solid var(--line);padding:10px 8px;cursor:pointer}
.list li:hover{background:#f0f5ff}
.list li.sel{background:#e8f0fe;border-left:3px solid var(--blue)}
.meta{color:var(--muted);font-size:12px}
.badge{display:inline-block;border-radius:5px;padding:1px 7px;font-size:12px;margin-right:6px}
.single{background:#e0f2fe;color:#075985}.multi{background:#ede9fe;color:#5b21b6}.judge{background:#dcfce7;color:#166534}
.ans{color:var(--green);font-weight:700;font-size:19px}
.level{font-weight:700}.level.good{color:var(--green)}.level.mid{color:var(--orange)}.level.low{color:var(--muted)}
.kv{margin:4px 0}.k{color:var(--muted);display:inline-block;min-width:74px}
pre{white-space:pre-wrap;word-break:break-word;font:inherit;margin:0}
.note{background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;padding:9px;margin-top:10px;color:#9a3412}
.ok{color:var(--green)}.err{color:#b91c1c}
@media(max-width:860px){.grid{grid-template-columns:1fr}.list{max-height:38vh}}
</style>
</head>
<body>
<header>
  <h1>学法用法速查助手 <span class="meta">离线版 · UOS arm64</span></h1>
  <div class="sub" id="stats">题库加载中…</div>
</header>
<div class="wrap">
  <div class="tabs">
    <button class="tab on" data-p="p1">速查</button>
    <button class="tab" data-p="p2">考试范围</button>
    <button class="tab" data-p="p3">说明 / 补库</button>
  </div>

  <section class="panel on" id="p1">
    <div class="searchbar">
      <input type="text" id="q" placeholder="粘贴整题或输入题干关键词，如：诉讼时效期间为" autofocus>
      <button class="primary" id="btn">查询</button>
    </div>
    <div class="grid">
      <div class="card"><ul class="list" id="results"></ul></div>
      <div class="card" id="detail">左侧点一条题目，这里显示答案、依据与解析。</div>
    </div>
  </section>

  <section class="panel" id="p2">
    <div class="searchbar"><input type="text" id="lq" placeholder="法规名称关键词，如：民法典、宪法、纪律处分"></div>
    <div class="card"><ul class="list" id="laws"></ul></div>
  </section>

  <section class="panel" id="p3">
    <div class="card">
      <h3>怎么用</h3>
      <p>在「速查」页粘贴题干或关键词，点查询。带题号、带选项的整题复制也能识别；粘贴选项后，遇到题干几乎相同但选项顺序不同的题，会按选项文本挑出正确那条。</p>
      <p>结果颜色：绿色「可信」可直接看答案；橙色「疑似」核对题干；灰色「参考」只用关键词命中。</p>
      <h3>题库文件</h3>
      <p>内置题库在 <code>src/data/questions.json</code>。若同目录放了 <code>questions.json</code>，启动时会优先读它，外置文件优先于内置。</p>
      <h3>补库</h3>
      <p>选 Excel 文件导入，支持「广西普法考试题库录入」导出的三张工作表。导入前自动留 <code>questions.json.bak</code>，同题干去重，原有解析保留。</p>
      <div class="searchbar">
        <input type="file" id="xlsx" accept=".xlsx">
        <button id="imp">导入</button>
      </div>
      <div id="importLog" class="meta"></div>
    </div>
  </section>
</div>
<script>
var $=function(s){return document.querySelector(s)};
var current=null;
function esc(s){return String(s==null?'':s).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]})}
function level(s){if(s>=0.62)return '<span class="level good">可信</span>';if(s>=0.40)return '<span class="level mid">疑似</span>';return '<span class="level low">参考</span>'}
function tname(t){return t==='single'?'单选':t==='multi'?'多选':'判断'}
function api(u,cb){fetch(u,{cache:'no-store'}).then(function(r){return r.json()}).then(cb).catch(function(e){console.error(e)})}
function stats(){api('/api/stats',function(d){
  $('#stats').textContent='内置 '+d.total+' 题（官方 '+d.official+'）· 法规 '+d.laws_total+' 部，直接关联 '+d.laws_covered+' 部';
})}
function search(){
  var q=$('#q').value.trim();if(!q)return;
  api('/api/search?q='+encodeURIComponent(q),function(d){
    var ul=$('#results');ul.innerHTML='';current=null;
    if(!d.length){ul.innerHTML='<li class="meta">没有命中，换个关键词试试。</li>';$('#detail').textContent='';return}
    d.forEach(function(x){
      var li=document.createElement('li');
      li.innerHTML='<div class="meta">'+level(x.score)+'<span class="badge '+x.type+'">'+tname(x.type)+'</span>'+x.score.toFixed(3)+' · '+esc(x.id)+'</div><div>'+esc(x.stem)+'</div><div class="meta">答案：'+esc(x.answer)+'</div>';
      li.onclick=function(){
        Array.prototype.forEach.call(ul.children,function(n){n.classList.remove('sel')});
        li.classList.add('sel');detail(x.id);
      };
      ul.appendChild(li);
    });
    detail(d[0].id);ul.children[0].classList.add('sel');
  })
}
function detail(id){api('/api/detail?id='+encodeURIComponent(id),function(q){
  if(!q){$('#detail').textContent='未找到题目';return}
  var h='<div class="meta">'+q.id+' · '+tname(q.type)+'</div><h3>'+esc(q.stem)+'</h3>';
  h+='<p><span class="k">答案</span><span class="ans">'+esc(q.answer_display)+'</span> <span class="meta">('+esc(q.answer)+')</span></p>';
  var ks=Object.keys(q.options||{}).sort();
  if(ks.length){h+='<p class="k">选项</p>';ks.forEach(function(k){h+='<div><span class="k">'+k+'</span><span>'+esc(q.options[k])+'</span></div>'})}
  h+='<p><span class="k">依据</span><pre>'+esc(q.basis||'（题库未提供）')+'</pre></p>';
  h+='<p><span class="k">关联法规</span>'+esc(q.law||'—')+'</p>';
  h+='<p class="meta">来源：'+esc(q.source||'—')+'</p>';
  if(q.caution)h+='<div class="note">提示：'+esc(q.caution)+'</div>';
  $('#detail').innerHTML=h;
})}
function laws(){var k=$('#lq').value.trim();api('/api/laws?q='+encodeURIComponent(k),function(d){
  var ul=$('#laws');ul.innerHTML='';
  d.forEach(function(x,i){var li=document.createElement('li');li.innerHTML='<span class="meta">'+(i+1)+'.</span> '+esc(x.name)+' <span class="meta">['+esc(x.category)+']</span>';ul.appendChild(li)});
  if(!d.length)ul.innerHTML='<li class="meta">没有匹配的法规。</li>';
})}
function imp(){
  var f=$('#xlsx').files[0];if(!f){$('#importLog').textContent='请先选择 Excel 文件';return}
  $('#importLog').textContent='导入中…';
  fetch('/api/import?name='+encodeURIComponent(f.name),{method:'POST',headers:{'Content-Type':'application/octet-stream'},body:f})
  .then(function(r){return r.json()}).then(function(d){
    $('#importLog').innerHTML='<span class="'+(d.ok?'ok':'err')+'">'+esc(d.message)+'</span>';
    stats();
  }).catch(function(e){$('#importLog').textContent='导入失败：'+e});
}
document.querySelectorAll('.tab').forEach(function(b){b.onclick=function(){
  document.querySelectorAll('.tab').forEach(function(x){x.classList.remove('on')});
  document.querySelectorAll('.panel').forEach(function(x){x.classList.remove('on')});
  b.classList.add('on');$('#'+b.dataset.p).classList.add('on');
}});
$('#btn').onclick=search;
$('#q').onkeydown=function(e){if(e.keyCode===13)search()};
var timer=null;$('#q').oninput=function(){clearTimeout(timer);timer=setTimeout(search,350)};
$('#lq').oninput=laws;$('#imp').onclick=imp;
stats();laws();
</script>
</body>
</html>
"""


class Handler(BaseHTTPRequestHandler):
    server_version = "XuefaUOS/1.0"

    def log_message(self, fmt, *args):
        # 本地工具，不往终端刷访问日志
        pass

    def _send(self, code, body, ctype):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj, code=200):
        self._send(code, json.dumps(obj, ensure_ascii=False),
                   "application/json; charset=utf-8")

    def do_GET(self):
        u = urlparse(self.path)
        if u.path in ("/", ""):
            self._send(200, PAGE, "text/html; charset=utf-8")
            return
        q = parse_qs(u.query)
        if u.path == "/api/stats":
            covered, total = BANK.law_coverage()
            self._json({
                "total": len(BANK.questions),
                "official": sum(1 for x in BANK.questions
                                if "官方题库" in x.get("source", "")),
                "laws_total": total,
                "laws_covered": covered,
            })
            return
        if u.path == "/api/search":
            text = q.get("q", [""])[0]
            try:
                top = max(1, min(30, int(q.get("top", ["8"])[0])))
            except ValueError:
                top = 8
            hits = BANK.search(text, top)
            self._json([{
                "id": x.get("id", ""),
                "score": round(s, 3),
                "type": x.get("type", ""),
                "stem": x.get("stem", ""),
                "answer": bankmod.QuestionBank.answer_text(x),
            } for s, x in hits])
            return
        if u.path == "/api/detail":
            wanted = q.get("id", [""])[0]
            hit = None
            for x in BANK.questions:
                if str(x.get("id", "")) == wanted:
                    hit = x
                    break
            if hit is None:
                self._json({"error": "not found"}, 404)
                return
            out = dict(hit)
            out["answer_display"] = bankmod.QuestionBank.answer_text(hit)
            self._json(out)
            return
        if u.path == "/api/laws":
            key = q.get("q", [""])[0]
            try:
                limit = max(1, min(382, int(q.get("limit", ["382"])[0])))
            except ValueError:
                limit = 382
            self._json(BANK.search_laws(key, limit))
            return
        if u.path == "/favicon.ico":
            self._send(204, b"", "image/x-icon")
            return
        self._send(404, "not found", "text/plain; charset=utf-8")

    def do_POST(self):
        u = urlparse(self.path)
        if u.path != "/api/import":
            self._send(404, "not found", "text/plain; charset=utf-8")
            return
        global BANK
        try:
            length = int(self.headers.get("Content-Length", "0"))
            name = unquote(parse_qs(u.query).get("name", ["上传.xlsx"])[0])
            suffix = os.path.splitext(name)[1] or ".xlsx"
            raw = self.rfile.read(length)
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as f:
                f.write(raw)
                tmp = f.name
            try:
                imported, detail = ximp.import_xlsx(tmp)
            finally:
                os.unlink(tmp)
            if not imported:
                self._json({"ok": False, "message": "没有解析到有效题目",
                            "detail": detail})
                return
            with BANK_LOCK:
                merged = ximp.merge_questions(BANK.questions, imported)
                target = paths.external_path("questions.json")
                ximp.save_bank(target, merged)
                BANK = bankmod.QuestionBank.load()
            self._json({"ok": True,
                        "message": "导入完成：%d 题，现共 %d 题" % (
                            len(imported), len(BANK.questions)),
                        "detail": detail})
        except Exception as exc:
            self._json({"ok": False, "message": "导入失败：%s" % exc})


def main():
    global BANK
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument("--query", help="命令行直接查一题，不开服务")
    args = ap.parse_args()

    BANK = bankmod.QuestionBank.load()
    if args.query:
        hits = BANK.search(args.query, 5)
        if not hits:
            print("没有命中")
            return 0
        for score, q in hits:
            print("%.3f | %s | %s | %s" % (
                score, q.get("id", ""),
                bankmod.QuestionBank.answer_text(q),
                q.get("stem", "")[:70]))
        return 0

    server = None
    port = args.port
    for attempt in range(20):
        try:
            server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
            break
        except OSError:
            port += 1
    if server is None:
        print("找不到可用端口", file=sys.stderr)
        return 1
    url = "http://127.0.0.1:%d/" % port
    print("题库：%d 题，法规：%d 部" % (
        len(BANK.questions), len(BANK.laws.get("items", []))))
    print("本地地址：%s" % url)
    print("关闭：Ctrl+C")
    if not args.no_browser:
        threading.Timer(0.4, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
