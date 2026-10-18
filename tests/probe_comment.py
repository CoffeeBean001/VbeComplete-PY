# -*- coding: utf-8 -*-
"""Ctrl+/ 批量注释的真机探针（★v120）。

**只读 + 只在当前模块【末尾】追加/删除测试行**，不碰用户已有的任何一行：
最后会把追加的行删掉，再把全文与开工前的备份逐字比对，不一致就大声报错。

用法（Excel / WPS 开着、VBE 里有一个代码窗）：
    py -3 tests\\probe_comment.py
"""

import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import vbe_bridge as VB


MARK = "ZZProbeComment"


def _read_all(cm):
    n = int(cm.CountOfLines)
    return [cm.Lines(i, 1) or "" for i in range(1, n + 1)] if n else []


def _show(tag, lines, base):
    print("  %-10s %s" % (tag, lines[base:]))


def main():
    vbe = VB._get_vbe_cached()
    if vbe is None:
        print("没有可用的 VBE（宿主没开 / 代码窗没打开）")
        return 2
    cp = vbe.ActiveCodePane
    if cp is None:
        print("没有活动的代码窗")
        return 2
    cm = cp.CodeModule
    print("模块 = %r  行数 = %d" % (cm.Name, cm.CountOfLines))

    before = _read_all(cm)
    n0 = len(before)

    # 追加 3 行（缩进 4 空格的"块"）+ 1 行空行，模拟真实的整段操作
    block = ["Sub %s()" % MARK,
             "    Dim x As Long",
             "    x = 1",
             ""]
    try:
        for d, ln in enumerate(block):
            cm.InsertLines(n0 + 1 + d, ln)

        sl, el = n0 + 1, n0 + 3          # 选中前三行（第四行是空行，不该被动）
        cp.SetSelection(sl, 1, el, 999)   # EndCol 给足，别落成"停在行首"
        ok1 = VB.VbeBackend().toggle_comment()
        after1 = _read_all(cm)
        _show("加注释 ok=%s" % ok1, after1, n0)

        # 选区应当被恢复成同样这 3 行（否则"连按两次 = 取消"就不成立）
        sel = tuple(cp.GetSelection())
        print("  选区恢复 =", sel, "（期望 起=%d 末=%d）" % (sl, el))

        ok2 = VB.VbeBackend().toggle_comment()
        after2 = _read_all(cm)
        _show("取消 ok=%s" % ok2, after2, n0)

        # 单行、无选区
        cp.SetSelection(n0 + 2, 1, n0 + 2, 1)
        ok3 = VB.VbeBackend().toggle_comment()
        _show("单行 ok=%s" % ok3, _read_all(cm), n0)
        ok4 = VB.VbeBackend().toggle_comment()
        _show("单行回来 ok=%s" % ok4, _read_all(cm), n0)

        # 全空白选区：必须返回 False（调用方据此把按键还回去）
        cp.SetSelection(n0 + 4, 1, n0 + 4, 1)
        ok5 = VB.VbeBackend().toggle_comment()
        print("  空行     ok=%s（期望 False = 把按键还回去）" % ok5)
    finally:
        # 收尾：把追加的行删掉，并与备份逐字比对
        try:
            extra = int(cm.CountOfLines) - n0
            if extra > 0:
                cm.DeleteLines(n0 + 1, extra)
        except Exception as e:
            print("⚠️ 删除测试行失败（请手工删掉末尾 %d 行）: %s" % (extra, e))
        after = _read_all(cm)
        same = (after == before)
        print("\n还原核对: %s（前 %d 行逐字%s）"
              % ("OK" if same else "失败", n0, "相同" if same else "不同"))
        if not same:
            for i, (a, b) in enumerate(zip(before, after), 1):
                if a != b:
                    print("  第 %d 行 前=%r 后=%r" % (i, a, b))
            return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
