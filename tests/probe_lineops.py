"""★v122 真机探针：三个"按行"的编辑动作。

在**真 Excel + 真 VBE**里跑，验证 `new_line_above` / `duplicate_lines` /
`delete_lines_here` 真的能改动编辑器，而不是在假对象上"看起来对"。

安全措施（与 `probe_comment.py` 同一套）：
  * **只在活动模块末尾追加**一个临时过程，绝不动你已有的任何一行；
  * 跑完把追加的那几行删掉，再与开工前的备份**逐字比对**；
  * 万一中途抛异常，`finally` 里照样恢复（恢复失败会红字告诉你，并打印备份
    的前几行，方便手动还原）。

用法：

    py -3 tests/probe_lineops.py

需要：Excel（或 WPS）开着、已启用"信任对 VBA 工程对象模型的访问"、VBE 里
**至少有一个打开的代码窗**。
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import vbe_bridge as V

PASS = 0
FAIL = 0


def ck(tag, ok, extra=""):
    global PASS, FAIL
    if ok:
        PASS += 1
        print("  OK    %s" % tag)
    else:
        FAIL += 1
        print("  FAIL  %s   %s" % (tag, extra))


def main():
    vbe = V._get_vbe()
    if vbe is None:
        print("拿不到 VBE —— 先把 Excel 打开、进 VBE（并确认已信任工程对象模型）")
        return 1
    cp = vbe.ActiveCodePane
    if cp is None:
        print("当前没有打开的代码窗 —— 在 VBE 里双击一个模块再跑")
        return 1
    cm = cp.CodeModule
    be = V.VbeBackend()

    base = int(cm.CountOfLines)
    backup = str(cm.Lines(1, base)) if base else ""
    print("活动模块 %s：开工前 %d 行\n" % (cm.Parent.Name
                                          if hasattr(cm, "Parent") else "?", base))

    block = ["Sub VbeCompleteProbe122()",
             "    Dim probeVar As Long",
             "    probeVar = 1",
             "End Sub"]
    added = 0
    try:
        cm.InsertLines(base + 1, "\r\n".join(block))
        added = len(block)
        n0 = int(cm.CountOfLines)
        ck("在模块末尾追加 4 行临时过程", n0 == base + 4, (base, n0))

        target = base + 3                       # `    probeVar = 1` 那一行
        got = str(cm.Lines(target, 1)).rstrip("\r\n")
        ck("临时块形状正确（第 %d 行是赋值语句）" % target,
           got.strip() == "probeVar = 1", repr(got))

        # ---- 1) Ctrl+Shift+Enter：在上方插一行，与赋值语句【同级】 ----
        cp.SetSelection(target, 1, target, 1)
        ok1 = be.new_line_above()
        n1 = int(cm.CountOfLines)
        new_line = str(cm.Lines(target, 1)).rstrip("\r\n")
        shifted = str(cm.Lines(target + 1, 1)).rstrip("\r\n")
        ck("Ctrl+Shift+Enter 在上方插一行：行数 +1、新行缩进 4 格、"
           "原来那行整体下移",
           ok1 and n1 == base + 5 and new_line == "    "
           and shifted.strip() == "probeVar = 1",
           (ok1, n1, repr(new_line), repr(shifted)))

        # ---- 2) Ctrl+D：复制光标这一行到下方 ----
        cp.SetSelection(target, 1, target, 1)
        ok2 = be.duplicate_lines()
        n2 = int(cm.CountOfLines)
        dup = str(cm.Lines(target + 1, 1)).rstrip("\r\n")
        ck("Ctrl+D 复制当前行到下方：行数 +1、复制出来的内容逐字相同",
           ok2 and n2 == base + 6 and dup == new_line,
           (ok2, n2, repr(dup), repr(new_line)))

        # ---- 3) Ctrl+K：删掉一行 ----
        cp.SetSelection(target + 1, 1, target + 1, 1)
        ok3 = be.delete_lines_here()
        n3 = int(cm.CountOfLines)
        ck("Ctrl+K 删掉整行：行数 -1，剩下的行没被牵连",
           ok3 and n3 == base + 5
           and str(cm.Lines(target, 1)).rstrip("\r\n") == "    "
           and str(cm.Lines(target + 1, 1)).rstrip("\r\n").strip()
           == "probeVar = 1",
           (ok3, n3))

        # ★v126b：删完光标要停在【被删行上一行】的行尾。这里上一行是 `    `
        # 那行（4 个空格），行尾显示列 = 5。真机读一次 —— 这条正是"VBE 会不会
        # 把远超行尾的列夹回行尾"的现场验证，假对象上永远测不出来。
        try:
            _sl3, _sc3, _el3, _ec3 = cp.GetSelection()
            sel3 = (int(_sl3), int(_sc3), int(_el3), int(_ec3))
        except Exception:
            sel3 = None
        ck("Ctrl+K 删完光标落到【上一行行尾】（第 %d 行、显示列 5）" % target,
           sel3 is not None and sel3[0] == target and sel3[1] == 5
           and sel3[2] == target and sel3[3] == 5,
           (sel3, target))
    finally:
        # 无论中间发生什么，把追加的内容删干净，并与备份逐字比对。
        try:
            n = int(cm.CountOfLines)
            if n > base:
                cm.DeleteLines(base + 1, n - base)
            now_n = int(cm.CountOfLines)
            now = str(cm.Lines(1, now_n)) if now_n else ""
            if now == backup:
                ck("跑完恢复原样（与开工前备份**逐字**相同）", True)
            else:
                ck("跑完恢复原样（与开工前备份**逐字**相同）", False,
                   "行数 %d -> %d；备份前 3 行：%r"
                   % (base, now_n, backup.split("\r\n")[:3]))
        except Exception as exc:
            ck("跑完恢复原样", False, "恢复时炸了：%r（请手动检查模块末尾）" % (exc,))

    print("\n%d OK / %d FAIL" % (PASS, FAIL))
    return 0 if FAIL == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
