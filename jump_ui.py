# -*- coding: utf-8 -*-
"""过程大纲 / 查找引用窗（★v121）。

两个功能**共用同一个窗** —— 它们其实是同一件事：一列可模糊搜索的结果，点一下
（或回车）跳过去。差别只在数据来源：

  * `Ctrl+Shift+O` 过程大纲：当前模块里所有 `Sub` / `Function` / `Property`…
  * `Ctrl+Shift+F` 查找引用：某个名字在本工程里出现的每一处（**文本级**）

★v121b：**单击即跳转**（不必记"双击"），而且**跳完窗不关** —— 连点几处引用
不用反复开窗。关窗只有三条路：`Esc`、窗口右上角的关闭按钮、再按一次同一条
快捷键（钩子侧 `_open_outline` / `_open_usages` 见到窗开着就收起）。

分工与 v115 片段窗完全一致（同一套教训）：
  * 本模块只碰 tkinter，**不碰 COM**。读模块、跳光标都在 `main`（主线程），
    这里通过回调 `on_pick(item)` 把"用户选了哪一项"交出去；
  * 窗是普通 `Toplevel`（要收键盘）⇒ 开窗期间主钩子已经被 `reconcile` 卸掉，
    所以 `Esc` / 再按一次同一条快捷键，都必须在**窗内自绑**；
  * 关窗之后**主动把焦点还给 VBE**（`on_closed`）—— root 是 withdrawn 的，
    Windows 未必把前台还给 VBE，落到一个看不见的窗口上就是"键盘没反应了"。

配色 / 字号 / 滚动条宽度全部从 `snippet_ui` 借（同一份实现，别各写一套）。
"""

import tkinter as tk

import snippets as S
from snippet_ui import (BG, PANEL, FG, MUTED, ACCENT, ERR, UI_FONT, UI_FONT_B,
                        _scrollbar_width)

DEFAULT_W = 620
DEFAULT_H = 430
MIN_W = 440
MIN_H = 220
LIST_ROWS = 16

_HINT = "单击 / 回车跳转　·　Esc 关闭（窗不会自动关）"


class JumpWindow(object):
    """大纲 / 引用窗。一个实例长期复用，show/hide 之间不重建控件。"""

    def __init__(self, root):
        self.root = root
        self.on_pick = None      # (item) -> None；item 是 show_items 里那一项
        self.on_closed = None    # 收起之后（把焦点还给 VBE）
        self.win = None
        self._w, self._h = DEFAULT_W, DEFAULT_H
        self._sbw = 17
        self._items = []         # [{"label","module","line","kind","text"}]
        self._view = []          # [(下标, 显示的 label)]
        self._bulk = False
        self._title = "过程大纲"
        self._subtitle = ""

    # ------------------------------------------------------------------
    # 建窗
    # ------------------------------------------------------------------
    def _build(self):
        w = tk.Toplevel(self.root)
        # ⚠️ 这一行不能漏：`show_items` / `close` / `is_open` / `_place` 全都靠
        # `self.win`（漏掉时窗口根本不会出现，而且所有异常都被外层的 except 吞掉，
        # 表现为"按了快捷键什么都没发生"—— 真机上非常难查）。
        self.win = w
        w.title(self._title)
        w.configure(bg=BG)
        w.withdraw()
        w.minsize(MIN_W, MIN_H)
        w.protocol("WM_DELETE_WINDOW", self.close)
        w.bind("<Escape>", lambda e: self.close())
        # 再按一次同一条快捷键 = 收起（与片段窗的 Ctrl+Q 同口径）。
        # 大写/小写两套都绑：Tk 的 keysym 在按下 Shift 时给的是大写字母。
        for seq in ("<Control-Shift-O>", "<Control-Shift-o>",
                    "<Control-Shift-F>", "<Control-Shift-f>"):
            try:
                w.bind(seq, lambda e: (self.close(), "break")[1])
            except Exception:
                pass

        w.grid_rowconfigure(2, weight=1)
        w.grid_columnconfigure(0, weight=1)

        # ---- 第一行：标题（"过程大纲 · 模块1"）----
        self.lbl_head = tk.Label(w, text="", bg=BG, fg=FG, font=UI_FONT_B,
                                 anchor="w")
        self.lbl_head.grid(row=0, column=0, sticky="we", padx=8, pady=(8, 2))

        # ---- 第二行：搜索 ----
        bar = tk.Frame(w, bg=BG)
        bar.grid(row=1, column=0, sticky="we", padx=8, pady=(0, 4))
        tk.Label(bar, text="搜索", bg=BG, fg=MUTED, font=UI_FONT).pack(side="left")
        self.var_search = tk.StringVar()
        self.ent_search = tk.Entry(bar, textvariable=self.var_search,
                                   font=UI_FONT, relief="solid", bd=1,
                                   bg=PANEL, fg=FG, insertbackground=FG,
                                   highlightthickness=0)
        self.ent_search.pack(side="left", fill="x", expand=True, padx=(6, 0))
        self.var_search.trace_add("write", lambda *a: self._refilter())

        # ---- 第三行：列表 ----
        body = tk.Frame(w, bg=BG)
        body.grid(row=2, column=0, sticky="nsew", padx=8, pady=0)
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(0, weight=1)
        self.lb = tk.Listbox(body, font=UI_FONT, height=LIST_ROWS,
                             activestyle="none", exportselection=False,
                             highlightthickness=0, bd=1, relief="solid",
                             bg=PANEL, fg=FG, selectbackground=ACCENT,
                             selectforeground="#FFFFFF")
        self.lb.grid(row=0, column=0, sticky="nsew")
        sb = tk.Scrollbar(body, orient="vertical", command=self.lb.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.lb.configure(yscrollcommand=sb.set)
        self._sbw = _scrollbar_width(sb)
        self.lb.bind("<<ListboxSelect>>", self._on_select)
        self.lb.bind("<Return>", lambda e: (self._pick(), "break")[1])
        self.lb.bind("<KP_Enter>", lambda e: (self._pick(), "break")[1])
        # ★v121b：**单击**（松手那一下）就跳转 —— 不再要求记住"双击"。
        # 选中项是在**按下**时改的（`<<ListboxSelect>>` 已刷过预览），所以松手
        # 时 `_cur_item()` 就是用户点的那一行。
        self.lb.bind("<ButtonRelease-1>", self._on_click)

        # ---- 第四行：选中那一行的原文（长行折行显示）----
        self.lbl_preview = tk.Label(w, text="", bg=BG, fg=MUTED, font=UI_FONT,
                                    anchor="w", justify="left", wraplength=580)
        self.lbl_preview.grid(row=3, column=0, sticky="we", padx=8, pady=(6, 0))

        # ---- 第五行：状态 ----
        self.lbl_status = tk.Label(w, text="", bg=BG, fg=MUTED, font=UI_FONT,
                                   anchor="w")
        self.lbl_status.grid(row=4, column=0, sticky="we", padx=8, pady=(2, 8))

    # ------------------------------------------------------------------
    # 显示 / 收起
    # ------------------------------------------------------------------
    def is_open(self):
        try:
            return (self.win is not None and self.win.winfo_exists()
                    and self.win.state() != "withdrawn")
        except Exception:
            return False

    def show_items(self, rect, items, title=None, subtitle=""):
        """弹出并填好列表。

        `rect` = 光标屏幕坐标 `(x, 上, 下)`（取不到传 None）；`items` 见
        `main._open_outline` / `_open_usages` 的构造。
        """
        if self.win is None:
            self._build()
        self._items = list(items or ())
        self._title = title or "过程大纲"
        self._subtitle = subtitle or ""
        try:
            self.win.title(self._title)
            self.lbl_head.configure(text=(self._title + ("　·　" + self._subtitle
                                                        if self._subtitle else "")))
        except Exception:
            pass
        self._bulk = True
        try:
            self.var_search.set("")
        finally:
            self._bulk = False
        self._refilter()
        self._place(rect)
        try:
            self.win.deiconify()
            self.win.attributes("-topmost", True)
            self.win.lift()
            self.win.focus_force()
        except Exception:
            pass
        # 焦点给**搜索框**：大纲最常见的第一步就是打几个字母过滤
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
            # 焦点必须有个去处（见模块开头的说明）：统一交回 VBE
            cb = self.on_closed
            if cb:
                try:
                    cb()
                except Exception:
                    pass

    def _place(self, rect):
        """窗口左上角贴着光标（下方放不下就翻到上方），并夹在屏幕内。"""
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
            y = max(0, top - h - 6)
        x = max(0, min(x, sw - w))
        y = max(0, min(y, sh - h))
        try:
            self.win.geometry("%dx%d+%d+%d" % (w, h, x, y))
        except Exception:
            pass

    # ------------------------------------------------------------------
    # 列表
    # ------------------------------------------------------------------
    def _refilter(self):
        try:
            q = self.var_search.get()
        except Exception:
            q = ""
        # 复用片段库那套模糊匹配（整标题 > 前缀 > 连续子串 > 子序列，多词 AND）：
        # 同一份实现 —— 两个窗的搜索手感必须一样，各写一套迟早分叉。
        rows = [(it["label"], str(i)) for i, it in enumerate(self._items)]
        try:
            view = S.search(rows, q)
        except Exception:
            view = [(lab, key, []) for (lab, key) in rows]
        self._view = [(int(key), lab) for (lab, key, _h) in view]
        self._bulk = True
        try:
            self.lb.delete(0, "end")
            for (_i, lab) in self._view:
                self.lb.insert("end", lab)
        finally:
            self._bulk = False
        if self._view:
            self._select_row(0)
        else:
            try:
                self.lbl_preview.configure(text="")
            except Exception:
                pass
        self._status()

    def _select_row(self, idx):
        self._bulk = True
        try:
            self.lb.selection_clear(0, "end")
            self.lb.selection_set(idx)
            self.lb.activate(idx)
            self.lb.see(idx)
        finally:
            self._bulk = False
        self._on_select(None)

    def _cur_item(self):
        try:
            sel = self.lb.curselection()
        except Exception:
            sel = ()
        if not sel:
            return None
        i = int(sel[0])
        if not (0 <= i < len(self._view)):
            return None
        idx = self._view[i][0]
        if 0 <= idx < len(self._items):
            return self._items[idx]
        return None

    def _on_select(self, _evt):
        it = self._cur_item()
        if not it:
            try:
                self.lbl_preview.configure(text="")
            except Exception:
                pass
            return
        text = str(it.get("text") or "").strip()
        try:
            self.lbl_preview.configure(text=(text[:400] if text else ""))
        except Exception:
            pass

    def _on_click(self, _evt):
        """单击列表项 = 跳转。

        空列表 / 点在空白处 ⇒ 没有选中项，`_pick()` 静默返回（不弹任何东西）。
        """
        self._pick()
        return "break"

    def _status(self):
        total = len(self._items)
        shown = len(self._view)
        it = self._cur_item()
        if it:
            extra = "　第 %d 行" % int(it.get("line") or 0)
            if it.get("kind"):
                extra = "　" + str(it["kind"]) + extra
            head = "%s%s　·　" % (it.get("module") or "", extra)
        else:
            head = ""
        try:
            self.lbl_status.configure(
                text="%s共 %d 项，命中 %d 项　·　%s"
                     % (head, total, shown, _HINT), fg=MUTED)
        except Exception:
            pass

    def _pick(self):
        """单击 / 回车：把选中项交给调用方（跳转）。

        ★v121b：**跳完成不成都不收窗**（用户口径：导航窗要手动关）。跳成了在
        状态栏说"已跳到 模块1 · 第 3 行"，跳不成红字说原因 —— 这样连点几处引用
        不用反复开窗。关窗只有三条路：`Esc`、关闭按钮、再按一次同一条快捷键。
        """
        it = self._cur_item()
        if not it:
            return False
        cb = self.on_pick
        if cb is None:
            self._hint("跳转动作没接上（这是配置问题，不是你的操作问题）")
            return False
        try:
            ok = bool(cb(it))
        except Exception:
            ok = False
        if ok:
            self._say("已跳到 %s　·　第 %d 行"
                      % (it.get("module") or "", int(it.get("line") or 0)))
        else:
            self._hint("跳不过去：那个模块还没有打开代码窗"
                       "（在 VBE 里双击它一次，再回来点这一项）")
        return ok

    def _say(self, msg):
        """状态栏报一句正常结果（颜色用前景色，不是错误色）。"""
        try:
            self.lbl_status.configure(text=msg, fg=FG)
        except Exception:
            pass

    def _hint(self, msg):
        """状态栏临时显示一句提示（下次刷新列表时会被 `_status()` 覆盖）。"""
        try:
            self.lbl_status.configure(text=msg, fg=ERR)
        except Exception:
            pass
