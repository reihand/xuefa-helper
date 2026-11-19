# -*- coding: utf-8 -*-
"""
学法用法速查助手 —— tkinter 界面。

纯标准库。三个页签：速查 / 考试范围 / 使用说明。
- 速查：粘贴或输入题干 → 命中列表 → 答案 + 法条依据
- 剪贴板监听：复制题干后自动查询（只检测序号变化，不占用剪贴板）
- 导入题库 Excel：直接吃 Greasy Fork #560523 导出的三工作表文件
- 置顶小窗：默认开启，Esc 最小化
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import bank as bankmod
import paths
import xlsx_import as ximp
from clipboard import ClipboardWatcher

APP_TITLE = "学法用法速查助手"
FONT = ("Microsoft YaHei UI", 10)
FONT_SMALL = ("Microsoft YaHei UI", 9)
FONT_BOLD = ("Microsoft YaHei UI", 10, "bold")
FONT_ANSWER = ("Microsoft YaHei UI", 15, "bold")

C_ANSWER = "#0a7d3f"
C_WARN = "#b45309"
C_MUTED = "#6b7280"
C_LINK = "#1a56db"
C_BG = "#ffffff"
C_PANEL = "#f6f7f9"

HELP_TEXT = """【这个工具是什么】
本地离线题库速查。速查助手本身不联网、不连接考试系统、不读取考试页面。
内置 10196 道题：9904 道来自广西普法云平台「我的练习」官方题库（含参考
答案与题目解析），其余为公开真题与本地 Word 题库补充。
附件3 的 382 部法规清单用作范围索引，题库直接关联其中 264 部。

【怎么用】
1. 在「速查」页把题干粘进输入框（可整题粘贴，带题号、带选项都行），
   点「查询」或按 Ctrl+Enter。
2. 左侧出候选，点一条，右侧给完整题干、答案、法条依据。
3. 勾选「监听剪贴板」后，在别处复制题干，本窗口会自动查询，不用切回来粘贴。
4. 「窗口置顶」默认开启；按 Esc 立即最小化。

【命中率怎么看】
绿色「可信」= 题干基本对上。
橙色「疑似」= 题干相近，核对完整题干。
灰色「参考」= 仅关键词相关。
粘贴整题时程序会读取选项文本，用来区分「题干相同但选项顺序不同」的题，
避免答案字母给错。带「提示」字样的题目表示解析缺失、选项不全或存在版本差异。

【更新题库 · 题库采集器.exe】
把「题库采集器.exe」和本程序放同一目录，在 Edge / Chrome 登录过考试网站后
双击采集器：它读取本机已登录会话，从「我的练习」取题接口拉取题干、参考答案、
题目解析和法规出处，跑完在同目录生成 questions.json。
本程序启动时优先加载同目录的 questions.json，重启即生效。
  · 首次采集约 2~4 分钟，控制台显示进度，结束生成「题库采集报告.txt」
  · 重复运行自动合并去重并留 .bak；--resume 可断点续跑
  · --probe-only 只验证会话与接口；--sample / --workers 可调采集强度
  · 采集器需要联网，本速查助手始终纯离线

【导入 Excel 补库】
点速查页右上「导入题库 Excel」，选 Greasy Fork 免费脚本
「广西普法考试题库录入」（#560523）在练习模式攒题后导出的 Excel。
三张工作表（单选题录入 / 多选题录入 / 判断题录入）会自动识别：
  · 题型按工作表名判定，答案支持 A / A,C,D / ACD / 对 / 错 / T / F / √ / ×
  · 同题干自动去重，原有题目的解析与人工校正会保留
  · 结果写到 exe 同目录的 questions.json，下次启动仍然生效
  · 写之前会留一份 questions.json.bak，可随时回滚
也支持本工具的标准表头 Excel：题干 | 类型 | A | B | … | 答案 | 依据 | 法规 | 来源。
命令行补库用 tools/import_bank.py，Word 题库用 tools/import_docx.py。

【考试信息（来自桂司通〔2026〕号通知）】
正式考试：2026-11-20 至 2026-12-31，每人 2 次作答机会，取最高分。
补考：2027-01-01 至 2027-01-31，同样 2 次机会。
题型：单选、多选、判断；题目由系统从题库随机抽取。
考试内容：附件3 列出的 382 部法律法规、党内法规及规范性文件。

【数据说明】
题库按题干规范化去重。官方题库中 6789 道带接口原生解析；少量题目接口未返回
选项文本或答案字母越界，已在结果中标注提示。题目由系统随机抽取，题库不覆盖
全部 382 部法规。
"""


class App:
    def __init__(self, root, data):
        self.root = root
        self.bankobj = data
        self.watcher = ClipboardWatcher()
        self.results = []
        self._poll_ms = 400
        self._build_style()
        self._build_ui()
        self._bind_keys()
        self.root.after(self._poll_ms, self._poll_clipboard)

    # ---------- 样式 ----------
    def _build_style(self):
        try:
            ttk.Style().theme_use("vista")
        except tk.TclError:
            try:
                ttk.Style().theme_use("clam")
            except tk.TclError:
                pass
        st = ttk.Style()
        st.configure(".", font=FONT)
        st.configure("Treeview", rowheight=26, font=FONT_SMALL)
        st.configure("Treeview.Heading", font=FONT_SMALL)
        st.configure("TNotebook.Tab", padding=(14, 6))
        st.configure("Status.TLabel", foreground=C_MUTED, font=FONT_SMALL, background=C_PANEL)

    # ---------- 布局 ----------
    def _build_ui(self):
        self.root.title(APP_TITLE)
        self.root.geometry("1020x700")
        self.root.minsize(780, 540)
        self.root.configure(bg=C_PANEL)
        self.root.attributes("-topmost", True)

        nb = ttk.Notebook(self.root)
        nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self.tab_search = ttk.Frame(nb)
        self.tab_scope = ttk.Frame(nb)
        self.tab_help = ttk.Frame(nb)
        nb.add(self.tab_search, text="  速查  ")
        nb.add(self.tab_scope, text="  考试范围  ")
        nb.add(self.tab_help, text="  使用说明  ")

        self._build_search_tab()
        self._build_scope_tab()
        self._build_help_tab()

        # 状态栏
        bar = ttk.Frame(self.root, style="TFrame")
        bar.pack(fill=tk.X, side=tk.BOTTOM)
        self.status = ttk.Label(bar, style="Status.TLabel", text="")
        self.status.pack(side=tk.LEFT, padx=10, pady=4)
        self._refresh_status()
        ttk.Label(bar, style="Status.TLabel",
                  text="Ctrl+Enter 查询 ｜ Esc 最小化").pack(side=tk.RIGHT, padx=10)

    def _refresh_status(self):
        covered, total = self.bankobj.law_coverage()
        self.status.configure(
            text="题库 %d 题 ｜ 附件3 法规 %d 部（题库直接关联 %d 部）｜ 离线运行" % (
                len(self.bankobj.questions), total, covered))

    def _build_search_tab(self):
        f = self.tab_search

        # 输入区
        top = ttk.Frame(f)
        top.pack(fill=tk.X, padx=10, pady=(10, 6))
        ttk.Label(top, text="题干 / 关键词：", font=FONT_BOLD).pack(side=tk.LEFT)
        self.btn_query = ttk.Button(top, text="查询", command=self.do_query, width=9)
        self.btn_query.pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(top, text="清空", command=self.clear_input, width=7).pack(side=tk.RIGHT)
        ttk.Button(top, text="导入题库 Excel", command=self._import_excel,
                   width=16).pack(side=tk.RIGHT, padx=(0, 8))

        box = ttk.Frame(f)
        box.pack(fill=tk.X, padx=10)
        self.input = tk.Text(box, height=3, wrap=tk.WORD, font=FONT,
                             relief=tk.SOLID, borderwidth=1, undo=True)
        self.input.pack(fill=tk.X, ipady=4, padx=(0, 0))
        self.input.configure(highlightthickness=1, highlightbackground="#d1d5db")

        # 开关区
        opts = ttk.Frame(f)
        opts.pack(fill=tk.X, padx=10, pady=8)
        self.clip_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opts, text="监听剪贴板（复制题干自动查）", variable=self.clip_var,
                        command=self._toggle_clip).pack(side=tk.LEFT)
        self.top_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="窗口置顶", variable=self.top_var,
                        command=self._toggle_top).pack(side=tk.LEFT, padx=(18, 0))
        self.hint = ttk.Label(opts, text="", foreground=C_MUTED, font=FONT_SMALL)
        self.hint.pack(side=tk.RIGHT)

        # 结果区：左右分栏
        pane = ttk.PanedWindow(f, orient=tk.HORIZONTAL)
        pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        left = ttk.Frame(pane)
        right = ttk.Frame(pane)
        pane.add(left, weight=4)
        pane.add(right, weight=5)

        ttk.Label(left, text="命中结果", font=FONT_BOLD).pack(anchor=tk.W, pady=(0, 4))
        cols = ("score", "type", "answer", "stem")
        self.tree = ttk.Treeview(left, columns=cols, show="headings", height=16)
        self.tree.heading("score", text="匹配")
        self.tree.heading("type", text="题型")
        self.tree.heading("answer", text="答案")
        self.tree.heading("stem", text="题干")
        self.tree.column("score", width=68, anchor=tk.CENTER, stretch=False)
        self.tree.column("type", width=66, anchor=tk.CENTER, stretch=False)
        self.tree.column("answer", width=82, anchor=tk.CENTER, stretch=False)
        self.tree.column("stem", width=260, anchor=tk.W)
        self.tree.tag_configure("high", foreground="#065f46")
        self.tree.tag_configure("mid", foreground=C_WARN)
        self.tree.tag_configure("low", foreground=C_MUTED)
        sb = ttk.Scrollbar(left, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        ttk.Label(right, text="答案与依据", font=FONT_BOLD).pack(anchor=tk.W, pady=(0, 4))
        dbox = ttk.Frame(right, relief=tk.SOLID, borderwidth=1)
        dbox.pack(fill=tk.BOTH, expand=True)
        self.detail = tk.Text(dbox, wrap=tk.WORD, font=FONT, bg=C_BG,
                              relief=tk.FLAT, padx=12, pady=10, state=tk.DISABLED)
        dsb = ttk.Scrollbar(dbox, orient=tk.VERTICAL, command=self.detail.yview)
        self.detail.configure(yscrollcommand=dsb.set)
        self.detail.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        dsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.detail.tag_configure("h", font=FONT_BOLD, foreground=C_LINK)
        self.detail.tag_configure("ans", font=FONT_ANSWER, foreground=C_ANSWER)
        self.detail.tag_configure("warn", foreground=C_WARN, font=FONT_BOLD)
        self.detail.tag_configure("muted", foreground=C_MUTED, font=FONT_SMALL)
        self.detail.tag_configure("stem", font=FONT)
        self.detail.tag_configure("sep", foreground="#d1d5db")
        self._detail_empty()

    def _build_scope_tab(self):
        f = self.tab_scope
        top = ttk.Frame(f)
        top.pack(fill=tk.X, padx=10, pady=10)
        ttk.Label(top, text="附件3 考试内容范围（382 部）：", font=FONT_BOLD).pack(side=tk.LEFT)
        self.law_kw = tk.StringVar()
        ent = ttk.Entry(top, textvariable=self.law_kw, width=30, font=FONT)
        ent.pack(side=tk.LEFT, padx=(10, 4))
        ttk.Button(top, text="筛选", command=self._filter_laws, width=8).pack(side=tk.LEFT)
        ttk.Button(top, text="全部", command=self._all_laws, width=7).pack(side=tk.LEFT, padx=4)
        self.law_count = ttk.Label(top, text="", foreground=C_MUTED, font=FONT_SMALL)
        self.law_count.pack(side=tk.RIGHT)

        cols = ("no", "name", "cat", "hits")
        self.law_tree = ttk.Treeview(f, columns=cols, show="headings")
        for c, t, w in [("no", "序号", 62), ("name", "法规 / 文件名称", 520),
                        ("cat", "类别", 150), ("hits", "题库关联题数", 110)]:
            self.law_tree.heading(c, text=t)
            self.law_tree.column(c, width=w, anchor=tk.CENTER if c != "name" else tk.W)
        sb = ttk.Scrollbar(f, orient=tk.VERTICAL, command=self.law_tree.yview)
        self.law_tree.configure(yscrollcommand=sb.set)
        self.law_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0), pady=(0, 10))
        sb.pack(side=tk.RIGHT, fill=tk.Y, pady=(0, 10))

        self._rebuild_law_hits()
        self._all_laws()

    def _rebuild_law_hits(self):
        """题库变化后重算「法规关联题数」，供考试范围页显示。"""
        self._law_hits = {}
        for q in self.bankobj.questions:
            key = bankmod.normalize(q.get("law", ""))
            if key:
                self._law_hits[key] = self._law_hits.get(key, 0) + 1

    def _build_help_tab(self):
        f = self.tab_help
        box = ttk.Frame(f)
        box.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
        t = tk.Text(box, wrap=tk.WORD, font=FONT, bg=C_BG, relief=tk.FLAT,
                    padx=16, pady=14)
        sb = ttk.Scrollbar(box, orient=tk.VERTICAL, command=t.yview)
        t.configure(yscrollcommand=sb.set)
        t.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        t.insert("1.0", HELP_TEXT)
        t.tag_add("h", "1.0", "1.end")
        t.tag_configure("h", font=("Microsoft YaHei UI", 13, "bold"), foreground=C_LINK)
        t.configure(state=tk.DISABLED)

    def _bind_keys(self):
        self.root.bind("<Control-Return>", lambda e: self.do_query())
        self.root.bind("<Escape>", lambda e: self.root.iconify())
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------- 行为 ----------
    def _toggle_clip(self):
        if self.clip_var.get():
            self.watcher.enable()
            self.hint.configure(text="剪贴板监听中：复制题干即自动查询")
        else:
            self.watcher.disable()
            self.hint.configure(text="")

    def _toggle_top(self):
        self.root.attributes("-topmost", bool(self.top_var.get()))

    def _poll_clipboard(self):
        try:
            if self.watcher.changed():
                try:
                    text = self.root.clipboard_get()
                except tk.TclError:
                    text = ""
                accepted = self.watcher.accept(text)
                if accepted:
                    self.set_input(accepted)
                    self.do_query()
        except Exception as exc:
            self.hint.configure(text="剪贴板读取异常：%s" % exc)
        self.root.after(self._poll_ms, self._poll_clipboard)

    def set_input(self, text):
        self.input.delete("1.0", tk.END)
        self.input.insert("1.0", text)

    def clear_input(self):
        self.input.delete("1.0", tk.END)
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.results = []
        self._detail_empty()
        self.hint.configure(text="已清空")

    def do_query(self):
        text = self.input.get("1.0", tk.END).strip()
        for i in self.tree.get_children():
            self.tree.delete(i)
        if not text:
            self.results = []
            self._detail_empty()
            self.hint.configure(text="请输入或粘贴题干")
            return
        self.results = self.bankobj.search(text, topk=12)
        if not self.results:
            self._detail_empty("未命中。试试只输入题干里最关键的 6-15 个字。")
            self.hint.configure(text="无匹配")
            return
        for score, q in self.results:
            tag = "high" if score >= 0.62 else ("mid" if score >= 0.40 else "low")
            level = "可信" if score >= 0.62 else ("疑似" if score >= 0.40 else "参考")
            stem = q["stem"].replace("\n", " ")
            if len(stem) > 46:
                stem = stem[:46] + "…"
            self.tree.insert("", tk.END,
                             values=("%s %.2f" % (level, score),
                                     bankmod.TYPE_NAME.get(q["type"], q["type"]),
                                     self.bankobj.answer_text(q), stem),
                             tags=(tag,))
        self.tree.selection_set(self.tree.get_children()[0])
        self._show_detail(0)
        n = sum(1 for s, _ in self.results if s >= 0.62)
        self.hint.configure(text="召回 %d 条，可信 %d 条" % (len(self.results), n))

    def _import_excel(self):
        """导入 #560523 导出的三工作表 Excel，合并后写到外置题库并即时重载。"""
        path = filedialog.askopenfilename(
            title="选择题库 Excel（支持 #560523 导出的三工作表格式）",
            filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            imported, detail = ximp.import_xlsx(path)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, "读取 Excel 失败：\n%s" % exc)
            return
        if not imported:
            messagebox.showwarning(
                APP_TITLE,
                "没有解析出有效题目，请确认表头与答案列。\n\n" + "\n".join(detail))
            return

        before = len(self.bankobj.questions)
        merged = ximp.merge_questions(self.bankobj.questions, imported)
        target = paths.external_path("questions.json")
        try:
            ximp.save_bank(target, merged)
            self.bankobj = bankmod.QuestionBank.load()
        except Exception as exc:
            messagebox.showerror(APP_TITLE, "写入题库失败：\n%s" % exc)
            return

        self.clear_input()
        self._rebuild_law_hits()
        self._all_laws()
        self._refresh_status()
        added = len(merged) - before
        self.hint.configure(text="导入完成：新增 %d 题，共 %d 题" % (added, len(merged)))
        messagebox.showinfo(
            APP_TITLE,
            "导入完成\n\n解析 %d 题，新增 %d 题，当前共 %d 题。\n"
            "已写入：%s\n（原文件已备份为 .bak）\n\n%s" % (
                len(imported), added, len(merged), target, "\n".join(detail)))

    def _on_select(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        self._show_detail(idx)

    def _show_detail(self, idx):
        if not (0 <= idx < len(self.results)):
            return
        score, q = self.results[idx]
        self.detail.configure(state=tk.NORMAL)
        self.detail.delete("1.0", tk.END)

        def line(text="", tag=None):
            self.detail.insert(tk.END, text + "\n", tag or "stem")

        level = "可信" if score >= 0.62 else ("疑似" if score >= 0.40 else "参考")
        line("【%s】　匹配度 %.2f　（%s）" % (
            bankmod.TYPE_NAME.get(q["type"], q["type"]), score, level), "h")
        line()
        line("答案： " + self.bankobj.answer_text(q), "ans")
        line()
        line("─" * 46, "sep")
        line("题干", "h")
        line(q["stem"])
        if q.get("options"):
            line()
            for k in sorted(q["options"].keys()):
                line("  %s. %s" % (k, q["options"][k]))
        line("─" * 46, "sep")
        line("法条依据", "h")
        line(q.get("basis") or "（本题公开题库未提供解析，请核对法条原文）")
        if q.get("law"):
            line()
            line("关联法规：" + q["law"], "muted")
        if q.get("caution"):
            line()
            line("提示：" + q["caution"], "warn")
        line()
        line("来源：" + (q.get("source") or "公开题库") + "　｜　题号 " + q.get("id", ""), "muted")
        self.detail.configure(state=tk.DISABLED)
        self.detail.yview_moveto(0)

    def _detail_empty(self, msg=None):
        self.detail.configure(state=tk.NORMAL)
        self.detail.delete("1.0", tk.END)
        self.detail.insert("1.0", msg or "左侧选中一条结果，这里显示完整题干、答案与法条依据。", "muted")
        self.detail.configure(state=tk.DISABLED)

    def _filter_laws(self):
        items = self.bankobj.search_laws(self.law_kw.get().strip(), limit=400)
        self._fill_laws(items)

    def _all_laws(self):
        self.law_kw.set("")
        self._fill_laws(self.bankobj.laws.get("items", []))

    def _fill_laws(self, items):
        for i in self.law_tree.get_children():
            self.law_tree.delete(i)
        for it in items:
            key = bankmod.normalize(it.get("name", ""))
            self.law_tree.insert("", tk.END, values=(
                it.get("no", ""), it.get("name", ""), it.get("category", ""),
                self._law_hits.get(key, 0)))
        self.law_count.configure(text="显示 %d 部" % len(items))

    def _on_close(self):
        self.watcher.disable()
        self.root.destroy()


def main():
    try:
        data = bankmod.QuestionBank.load()
    except Exception as exc:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(APP_TITLE, "题库加载失败：\n%s" % exc)
        return 1
    root = tk.Tk()
    App(root, data)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
