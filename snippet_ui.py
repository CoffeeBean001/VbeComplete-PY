# -*- coding: utf-8 -*-
"""代码片段窗（★v115，Ctrl+Q 唤出）。

用户口径（原话）：
  2. 按键CTRL+q调出此功能，ui弹出界面起点在vbe里的光标处。
  3. 有文本框，文本框里展示代码片段。
  4. 有列表框，列表框是所有代码片段标题。
  5. 对于列表框的标题，可以通过模糊匹配来搜索。
  6. 点击列表框的标题，文本框里展示相应标题的代码片段。
  7. 有新增，删除代码片段功能。
  8. 点击按钮，可以将选中代码片段插入到vbe中的光标处。

与补全弹窗（ui.Popup）的取向**刻意不同**：
  * 补全窗是 overrideredirect + WS_EX_NOACTIVATE 的"不抢焦点"浮窗 —— 它只
    显示、不接收输入；
  * 这个窗要**收键盘**（搜索框里要打字、文本框里要改代码），所以是普通
    Toplevel：会抢焦点、有标题栏、可拖动。代价是"打开期间键盘钩子会被
    main.reconcile 卸掉"（钩子只在 VBE 里挂），于是窗口内的 Ctrl+Q / Esc
    必须由本窗自己绑 —— 见 _build 里的两条 bind。

数据在 `snippets.SnippetStore`（SQLite 两列：title / code），本模块只管界面。
插入动作由调用方注入（`on_insert`），本模块**不碰 COM**。
"""

import tkinter as tk

import snippets as S


# ---- 配色（与补全弹窗同一家族，浅色、低饱和）----
BG = "#F7F6EF"
PANEL = "#FFFFFF"
FG = "#202020"
MUTED = "#6B6B6B"
ACCENT = "#2A5CAA"
ERR = "#C00000"

UI_FONT = ("Microsoft YaHei UI", 9)
UI_FONT_B = ("Microsoft YaHei UI", 9, "bold")
CODE_FONT = ("Consolas", 10)

DEFAULT_W = 780
DEFAULT_H = 470
MIN_W = 560
MIN_H = 300
LIST_CHARS = 26          # 列表框宽度（字符）
LIST_ROWS = 18

_SBW_FALLBACK = 17       # Tk 默认竖向滚动条宽度（量不到时兜底）


def _scrollbar_width(widget):
    """竖向滚动条的宽度（像素）。

    ★v117：搜索框 / 标题框 / 按钮的右边缘要和**代码文本框**的右边缘对齐，而代码框
    右边还站着一个滚动条 ⇒ 这几行的右侧留白都得多让出"一个滚动条"的宽度。这个宽度
    随 Tk 主题/DPI 变（17 是默认值，高 DPI 下会更宽），所以现场问控件要，别写死。
    """
    try:
        v = int(widget.winfo_reqwidth())
        if 8 <= v <= 48:
            return v
    except Exception:
        pass
    return _SBW_FALLBACK


class SnippetWindow(object):
    """片段管理窗。一个实例长期复用，show/hide 之间不重建控件。"""

    def __init__(self, root, store=None, on_insert=None):
        self.root = root
        self.store = store if store is not None else S.SnippetStore()
        self.on_insert = on_insert          # (code) -> bool
        self.on_closed = None               # 收起之后（把焦点还给 VBE）
        self.win = None
        self._w, self._h = DEFAULT_W, DEFAULT_H
        self._sbw = _SBW_FALLBACK           # 竖向滚动条宽度（_build 里现场量）
        self._rows = []                     # 库里全部 [(title, code)]
        self._view = []                     # 当前过滤结果 [(title, code, hits)]
        self._shown_title = None
        self._shown_code = None
        self._bulk = False                  # 程序性改列表框选中，别当用户点击

    # ------------------------------------------------------------------
    # 建窗
    # ------------------------------------------------------------------
    def _build(self):
        w = tk.Toplevel(self.root)
        w.title("代码片段")
        w.configure(bg=BG)
        w.withdraw()
        w.minsize(MIN_W, MIN_H)
        w.protocol("WM_DELETE_WINDOW", self.close)
        # 打开期间键盘钩子是卸掉的（钩子只在 VBE 里挂），这两条必须自己绑。
        # 用 bind（不是 bind_all）：别把 Esc 从文本框的编辑语义里抢走。
        w.bind("<Escape>", lambda e: self.close())
        w.bind("<Control-q>", lambda e: self.close())
        w.bind("<Control-Q>", lambda e: self.close())

        w.grid_rowconfigure(1, weight=1)
        w.grid_columnconfigure(0, weight=0)
        w.grid_columnconfigure(1, weight=1)

        # ---- 第一行：搜索 ----
        # ★v117（用户口径）：这一行和底下的按钮行，**右边缘都要跟代码文本框的右边
        # 缘落在同一条竖线上**。代码框右边还有一个竖向滚动条，所以这两行的右侧
        # 留白必须是 `8 + 滚动条宽度`，否则会各差十几像素（看着就是"没对齐"）。
        # 滚动条宽度现场量（reqwidth），量不到按 Tk 的默认 17 兜底。
        bar = tk.Frame(w, bg=BG)
        bar.grid(row=0, column=0, columnspan=2, sticky="we", padx=8, pady=(8, 4))
        tk.Label(bar, text="搜索", bg=BG, fg=MUTED, font=UI_FONT).pack(side="left")
        self.var_search = tk.StringVar()
        self.ent_search = tk.Entry(bar, textvariable=self.var_search,
                                   font=UI_FONT, relief="solid", bd=1,
                                   bg=PANEL, fg=FG,
                                   insertbackground=FG,
                                   highlightthickness=0)
        # 右侧 pad 交给 bar 的 padx 统一控制（这里不再留 6，否则又差 6 像素）
        self.ent_search.pack(side="left", fill="x", expand=True, padx=(6, 0))
        self.var_search.trace_add("write", lambda *a: self._refilter())

        # ---- 第二行左：标题列表 ----
        left = tk.Frame(w, bg=BG)
        left.grid(row=1, column=0, sticky="nsew", padx=(8, 4), pady=0)
        left.grid_rowconfigure(0, weight=1)
        left.grid_columnconfigure(0, weight=1)
        self.lb = tk.Listbox(left, font=UI_FONT, width=LIST_CHARS,
                             height=LIST_ROWS, activestyle="none",
                             exportselection=False, highlightthickness=0,
                             bd=1, relief="solid", bg=PANEL, fg=FG,
                             selectbackground=ACCENT, selectforeground="#FFFFFF")
        self.lb.grid(row=0, column=0, sticky="nsew")
        sb = tk.Scrollbar(left, orient="vertical", command=self.lb.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.lb.configure(yscrollcommand=sb.set)
        self.lb.bind("<<ListboxSelect>>", self._on_pick)
        self.lb.bind("<Key>", self._lb_key)
        self.lb.bind("<Return>", lambda e: self._insert())
        self.lb.bind("<Double-Button-1>", lambda e: "break")

        # ---- 第二行右：标题 + 代码 ----
        right = tk.Frame(w, bg=BG)
        right.grid(row=1, column=1, sticky="nsew", padx=(4, 8), pady=0)
        right.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure(1, weight=1)

        tk.Label(right, text="标题", bg=BG, fg=MUTED,
                 font=UI_FONT).grid(row=0, column=0, sticky="w", pady=(0, 4))
        self.var_title = tk.StringVar()
        self.ent_title = tk.Entry(right, textvariable=self.var_title,
                                  font=UI_FONT, relief="solid", bd=1,
                                  bg=PANEL, fg=FG, insertbackground=FG,
                                  highlightthickness=0)
        # ★v117：**不要 columnspan=2** —— 跨到滚动条那列就会比代码框多出十几像素。
        # 只占 col1（col2 是竖向滚动条），右边缘正好落在代码框右边缘上。
        self.ent_title.grid(row=0, column=1, sticky="we",
                            padx=(6, 0), pady=(0, 4))

        self.txt = tk.Text(right, font=CODE_FONT, wrap="none", undo=True,
                           relief="solid", bd=1, bg=PANEL, fg=FG,
                           insertbackground=FG, highlightthickness=0,
                           height=14)
        self.txt.grid(row=1, column=0, columnspan=2, sticky="nsew")
        sbv = tk.Scrollbar(right, orient="vertical", command=self.txt.yview)
        sbv.grid(row=1, column=2, sticky="ns")
        sbh = tk.Scrollbar(right, orient="horizontal", command=self.txt.xview)
        sbh.grid(row=2, column=0, columnspan=2, sticky="we")
        self.txt.configure(yscrollcommand=sbv.set, xscrollcommand=sbh.set)
        # 文本框里的 Tab 应该是"缩进"，不是"换焦点"
        self.txt.bind("<Tab>", self._txt_tab)
        self.txt.bind("<Control-Return>", lambda e: (self._insert(), "break")[1])

        # ★v117：量出竖向滚动条的实际宽度，再把搜索行 / 按钮行的右侧留白对齐过去
        # （搜索行、按钮行都是 `columnspan=2` 铺满整窗，它们的右边比代码框多出一个
        #   滚动条的宽度 —— 就是这个差让人看着"右边没对齐"）。
        self._sbw = _scrollbar_width(sbv)
        bar.grid_configure(padx=(8, 8 + self._sbw))

        # ---- 第三行：状态 + 按钮 ----
        foot = tk.Frame(w, bg=BG)
        foot.grid(row=2, column=0, columnspan=2, sticky="we",
                  padx=(8, 8 + self._sbw), pady=(6, 8))
        self.lbl_status = tk.Label(foot, text="", bg=BG, fg=MUTED,
                                   font=UI_FONT, anchor="w")
        self.lbl_status.pack(side="left", fill="x", expand=True)

        self.btn_insert = tk.Button(foot, text="插入到光标处", font=UI_FONT_B,
                                    command=self._insert, bg=ACCENT, fg="#FFFFFF",
                                    activebackground="#1E4A8C",
                                    activeforeground="#FFFFFF",
                                    relief="flat", bd=0, padx=14, pady=4)
        self.btn_insert.pack(side="right", padx=(6, 0))
        for text, cmd in (("新增", self._add),
                          ("保存修改", self._save),
                          ("删除", self._delete),
                          ("关闭", self.close)):
            tk.Button(foot, text=text, font=UI_FONT, command=cmd,
                      bg="#FFFFFF", fg=FG, relief="solid", bd=1,
                      padx=10, pady=3).pack(side="right", padx=(6, 0))

        self.win = w

    def _txt_tab(self, event):
        self.txt.insert("insert", "    ")
        return "break"

    # ------------------------------------------------------------------
    # 显示 / 收起
    # ------------------------------------------------------------------
    def is_open(self):
        try:
            return (self.win is not None and self.win.winfo_exists()
                    and self.win.state() != "withdrawn")
        except Exception:
            return False

    def open_at_caret(self, rect=None):
        """弹出并定位。`rect` = 光标屏幕坐标 (x, 上, 下)，取不到传 None。"""
        if self.win is None:
            self._build()
        self._bulk = True
        try:
            self.var_search.set("")
        finally:
            self._bulk = False
        self.reload()
        self._place(rect)
        try:
            self.win.deiconify()
            self.win.attributes("-topmost", True)
            self.win.lift()
            self.win.focus_force()
        except Exception:
            pass
        try:
            self.ent_search.focus_set()
            self.ent_search.selection_range(0, "end")
        except Exception:
            pass

    def close(self):
        if self.win is None:
            return
        was_open = self.is_open()
        try:
            if was_open:
                self._w = max(MIN_W, self.win.winfo_width())
                self._h = max(MIN_H, self.win.winfo_height())
        except Exception:
            pass
        try:
            self.win.withdraw()
        except Exception:
            pass
        if was_open:
            # 焦点必须有个去处：我们的窗是 root(withdrawn) 的子窗，收起后
            # Windows 未必把前台还给 VBE（可能落到一个看不见的窗口上，
            # 那样敲键盘就"没反应"了）—— 所以统一交回 VBE。
            cb = self.on_closed
            if cb:
                try:
                    cb()
                except Exception:
                    pass

    def _place(self, rect):
        """窗口左上角落在光标处（下方放不下就翻到上方），并夹在屏幕内。"""
        w, h = int(self._w), int(self._h)
        x, top = 120, 120
        if rect:
            try:
                x = int(rect[0])
                top = int(rect[1])
                bottom = int(rect[2])
            except Exception:
                bottom = top + 16
        else:
            bottom = top + 16
        try:
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
        except Exception:
            sw, sh = 1920, 1080
        y = bottom + 6
        if y + h > sh:
            y = max(0, top - h - 6)       # 下方放不下 -> 翻到光标上方
        x = max(0, min(x, sw - w))
        y = max(0, min(y, sh - h))
        try:
            self.win.geometry("%dx%d+%d+%d" % (w, h, x, y))
        except Exception:
            pass

    # ------------------------------------------------------------------
    # 列表 / 内容
    # ------------------------------------------------------------------
    def reload(self, keep=None):
        """从库里重读全部片段，再按当前搜索词过滤。"""
        try:
            self._rows = self.store.all()
        except Exception:
            self._rows = []
        self._refilter(keep if keep is not None else self._shown_title)

    def _refilter(self, keep=None):
        try:
            q = self.var_search.get()
        except Exception:
            q = ""
        self._view = S.search(self._rows, q)
        self._bulk = True
        try:
            self.lb.delete(0, "end")
            for (t, _c, _h) in self._view:
                self.lb.insert("end", t)
        finally:
            self._bulk = False

        titles = [t for (t, _c, _h) in self._view]
        dirty = self._is_dirty()
        target = None
        if keep and keep in titles:
            target = titles.index(keep)
        elif titles:
            target = 0

        if target is not None:
            self._select_row(target)
            # 文本框里有未保存的改动时**不要**被搜索刷新覆盖掉
            if not dirty:
                self._show(self._view[target][0])
        self._status("共 %d 条，命中 %d 条" % (len(self._rows), len(self._view)))

    def _select_row(self, idx):
        self._bulk = True
        try:
            self.lb.selection_clear(0, "end")
            self.lb.selection_set(idx)
            self.lb.activate(idx)
            self.lb.see(idx)
        finally:
            self._bulk = False

    def _on_pick(self, event=None):
        if self._bulk:
            return
        sel = self.lb.curselection()
        if not sel:
            return
        title = self.lb.get(sel[0])
        if title == self._shown_title:
            # 标题框被手改过又点回自己：把标题框还原
            self.var_title.set(title or "")
            return
        if self._is_dirty():
            if not self._confirm("当前文本框里的内容改过但没保存，切换会丢掉这些修改。"
                                 "\n\n要继续切换吗？"):
                self._reselect_shown()
                return
        self._show(title)

    def _reselect_shown(self):
        titles = [t for (t, _c, _h) in self._view]
        if self._shown_title in titles:
            self._select_row(titles.index(self._shown_title))

    def _is_dirty(self):
        """文本框内容是否与"当前展示的那条"不一致（纯比较，不靠事件记账）。"""
        if self._shown_title is None:
            return bool(self._txt_code().strip())
        return self._txt_code() != (self._shown_code or "")

    def _txt_code(self):
        try:
            return self.txt.get("1.0", "end-1c")
        except Exception:
            return ""

    def _show(self, title):
        code = self.store.get(title)
        if code is None:
            code = ""
        self._shown_title = title
        self._shown_code = code
        try:
            self.txt.delete("1.0", "end")
            self.txt.insert("1.0", code)
            self.txt.edit_reset()
            self.var_title.set(title or "")
        except Exception:
            pass

    def _lb_key(self, event):
        """在列表框里直接打字 -> 把这一下送进搜索框（别让 Listbox 自己跳字）。"""
        ch = getattr(event, "char", "") or ""
        if ch and ch.isprintable():
            try:
                self.ent_search.focus_set()
                self.ent_search.insert("end", ch)
            except Exception:
                pass
            return "break"
        return None

    # ------------------------------------------------------------------
    # 四个动作
    # ------------------------------------------------------------------
    def _add(self):
        title = (self.var_title.get() or "").strip()
        code = self._txt_code()
        if not title:
            self._status("先填标题，再点「新增」", err=True)
            try:
                self.ent_title.focus_set()
            except Exception:
                pass
            return
        if not code.strip():
            self._status("代码是空的，没法存", err=True)
            return
        if self.store.add(title, code):
            self.reload(keep=title)
            # 刚存进去的内容与文本框里的一模一样 —— 显式 _show 一次，让"当前
            # 展示项"认到这条新记录上（否则 _is_dirty 会一直认为有未保存改动，
            # 之后搜索/切换都不敢动文本框）。
            self._show(title)
            self._status("已新增「%s」" % title)
        else:
            self._status("标题「%s」已存在 —— 想改它请用「保存修改」" % title,
                         err=True)

    def _save(self):
        cur = self._shown_title
        if not cur:
            self._status("先在左边选一条要修改的片段", err=True)
            return
        title = (self.var_title.get() or "").strip()
        code = self._txt_code()
        if not title:
            self._status("标题不能为空", err=True)
            return
        if title != cur:
            if self.store.exists(title):
                self._status("标题「%s」已被别的片段占用" % title, err=True)
                return
            if not self.store.rename(cur, title):
                self._status("改标题失败", err=True)
                return
            self.store.update(title, code)
            self.reload(keep=title)
            self._show(title)          # 认到新标题上（见 _add 里的说明）
            self._status("已保存（标题改为「%s」）" % title)
            return
        if self.store.update(title, code):
            self.reload(keep=title)
            self._show(title)
            self._status("已保存「%s」" % title)
        else:
            self._status("保存失败（这条片段可能已被删掉）", err=True)

    def _delete(self):
        cur = self._shown_title
        if not cur:
            self._status("先在左边选一条要删除的片段", err=True)
            return
        if not self._confirm("确定删除片段「%s」？\n\n此操作不可撤销。" % cur):
            return
        if self.store.delete(cur):
            self._shown_title = None
            self._shown_code = None
            try:
                self.txt.delete("1.0", "end")
                self.var_title.set("")
            except Exception:
                pass
            self.reload(keep=None)
            self._status("已删除「%s」" % cur)
        else:
            self._status("删除失败", err=True)

    def _insert(self):
        code = self._txt_code()
        if not code.strip():
            self._status("没有可插入的代码", err=True)
            return
        ok = False
        try:
            ok = bool(self.on_insert(code)) if self.on_insert else False
        except Exception:
            ok = False
        if not ok:
            self._status("插入失败 —— 光标是不是已经离开代码窗格了？", err=True)
            return
        # 收起窗子（close 里会顺手把焦点交回 VBE，见那里的说明）
        self.close()

    # ------------------------------------------------------------------
    def _confirm(self, text):
        """是/否确认框。Toplevel 可能没建（测试里）——那就当作"否"。"""
        try:
            from tkinter import messagebox
            return bool(messagebox.askyesno("代码片段", text, parent=self.win))
        except Exception:
            return False

    def _status(self, msg, err=False):
        try:
            self.lbl_status.config(text=msg, fg=(ERR if err else MUTED))
        except Exception:
            pass
