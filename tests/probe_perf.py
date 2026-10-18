# -*- coding: utf-8 -*-
"""性能探针：量出"提示弹窗变慢"到底慢在哪个环节。

用法（Excel / WPS 开着，且 VBE 里有一个代码窗）：

    py -3 tests\\probe_perf.py

它只读、不写任何工程文件；弹窗用假 UI 顶掉，不会真的画出来。
输出三段：
  A. COM 侧单次耗时（get_context / get_identifiers 缓存命中 / 强制刷新）
  B. 整条 trigger 的耗时（假 UI，纯计算）
  C. 打分本身随候选池大小的曲线（离线合成池，不碰 COM）
"""
import os
import statistics
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import engine            # noqa: E402
import vbe_bridge        # noqa: E402


class FakeUI:
    """只记账、不画窗 —— 把 UI 开销从测量里剔掉（单独测）。"""

    def __init__(self):
        self.shown = 0
        self.hidden = 0

    def show(self, rows, sel, completer):
        self.shown += 1

    def hide(self):
        self.hidden += 1

    def update_selection(self, sel):
        pass

    def contains_point(self, x, y):
        return False


def _stat(label, fn, n=9, warm=2):
    """跑 n+warm 次，返回 (中位数ms, 最小ms, 最大ms)。"""
    for _ in range(warm):
        fn()
    ts = []
    for _ in range(n):
        t0 = time.perf_counter()
        fn()
        ts.append((time.perf_counter() - t0) * 1000.0)
    ts.sort()
    mid = statistics.median(ts)
    print("  %-42s 中位 %7.2f ms   最小 %7.2f   最大 %7.2f"
          % (label, mid, ts[0], ts[-1]))
    return mid


def main():
    backend = vbe_bridge.VbeBackend()
    vbe = vbe_bridge._get_vbe_cached()
    if vbe is None:
        print("!! 拿不到 VBE：先在 Excel/WPS 里打开 VBE（Alt+F11）再跑本探针")
        return 2
    cp = vbe.ActiveCodePane
    print("VBE ok; ActiveCodePane =", "有" if cp is not None else "无（先点一下代码窗）")

    print("\nA. COM 侧")
    _stat("get_context()", lambda: backend.get_context())
    _stat("get_identifiers() 缓存命中", lambda: backend.get_identifiers())
    # 强制刷新：清掉缓存
    def _refresh():
        backend._cache = None
        return backend.get_identifiers()
    _stat("get_identifiers() 强制全量刷新", _refresh, n=5, warm=1)
    ids = backend._cache or []
    print("  候选池规模: %d 条" % len(ids))
    try:
        mods = list(vbe.VBProject.VBComponents)
        total = 0
        for comp in mods:
            total += comp.CodeModule.CountOfLines
        print("  工程规模: %d 个模块 / %d 行代码" % (len(mods), total))
    except Exception as exc:
        print("  工程规模: 读不到（%s）" % (exc,))

    print("\nB. trigger 全链路（假 UI）")
    ui = FakeUI()
    comp = engine.Completer(backend, ui)
    ctx = backend.get_context()
    if not ctx:
        print("!! get_context 返回 None（光标不在代码窗？）—— 跳过 B")
    else:
        print("  当前行=%r 词=%r" % (ctx.get("line_text"), ctx.get("word")))
        _stat("trigger(True) 轮询路径", lambda: comp.trigger(True))
        _stat("trigger(manual=True) Ctrl+Space 路径",
              lambda: comp.trigger(False, manual=True))
        print("  候选条数: %d（窗里只画 %d 行）"
              % (len(comp.matches), engine.VIEW_ROWS))

    print("\nC. 打分曲线（离线，合成池）")
    # 真实池 + 合成填充，看"候选太多"到底占多少
    base = [str(r[0]) for r in (ids or [("dummy",)])][:200]
    for extra in (0, 1000, 3000, 6000):
        pool = list(base) + ["syntheticName%04d" % i for i in range(extra)]
        def run(pool=pool):
            for name in pool:
                engine.fuzzy_match(name, "syn")
        _stat("fuzzy_match 扫 %d 条" % len(pool), run, n=5, warm=1)
    return 0


if __name__ == "__main__":
    sys.exit(main())
