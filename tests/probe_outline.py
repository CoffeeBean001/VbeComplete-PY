# -*- coding: utf-8 -*-
"""手工探针：真机验证过程大纲 / 查找引用（★v121）。

跑法（Excel 开着、VBE 里有工程）：

    py -3 tests\\probe_outline.py

它做四件事，**一个字节都不写用户代码**：

  1. 把 VBE 里第一个已打开的代码窗切到前台（`CodePanes.Item(1).Show()`）；
  2. 读当前模块全文，列出过程（大纲的数据源）；
  3. 用第一个过程名做一次"文本级查找引用"；
  4. 跳回那个过程的第一行，读回 `GetSelection()` 核对；
  5. 再挑一个**没有打开代码窗**的模块，确认 `goto_location` 老老实实返回 False
     （这条是 v121 的真机教训：`VBComponent` 上没有 `CodePane` 属性，
     没打开过的模块在 COM 里根本打不开 —— 假对象测不出来）。

副作用：会让 VBE 停在某个模块的某个位置（不写文本、不改代码）。
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import vbe_bridge as VB


def main():
    fails = []

    def ck(name, cond, detail=""):
        print("%-48s %s%s" % (name, "OK" if cond else "FAIL",
                              ("  " + str(detail)) if detail else ""))
        if not cond:
            fails.append(name)

    vbe = VB._get_vbe_cached()
    if vbe is None:
        print("没有 VBE：先把 Excel 打开（并允许访问 VBA 工程对象模型）。")
        return 2

    panes = VB._code_panes_by_module(vbe)
    print("已打开的代码窗: %s" % (sorted(panes.keys()) or "（无）"))
    if not panes:
        print("一个代码窗都没开 —— 先在 VBE 里双击任意模块，再跑一次。")
        return 2

    first = vbe.CodePanes.Item(1)
    first.Show()                     # 切到前台，后面所有读取都基于它
    b = VB.VbeBackend()

    mod, code = b.current_module_text()
    ck("current_module_text 读到模块", bool(mod) and bool(code), (mod, len(code)))

    mod2, procs = b.list_procedures_here()
    print("过程（前 5 条）:")
    for row in procs[:5]:
        print("   ", row)
    ck("list_procedures_here 有过程", bool(procs), len(procs))
    ck("大纲行号与过程头对得上",
       all(str(code.split("\n")[r[2] - 1]).strip().lower().find(r[0].lower()) >= 0
           for r in procs if 0 < r[2] <= len(code.split("\n"))))

    if procs:
        name = procs[0][0]
        line = procs[0][2]
        mod0, hits, scanned = b.find_references(name)
        print("查找 %r -> %d 处 / %d 个模块（当前模块 %r）"
              % (name, len(hits), scanned, mod0))
        for h in hits[:5]:
            print("   ", h[:5])
        ck("find_references 至少命中 1 处", bool(hits), len(hits))
        ck("当前模块的命中排最前",
           not mod0 or str(hits[0][0]).lower() == str(mod0).lower())

        ok = b.goto_location(mod, line, expect=name)
        ck("goto_location 跳回过程头", bool(ok), (mod, line))
        try:
            sl, sc, el, ec = first.GetSelection()
            ck("落点是目标行、第 1 列", int(sl) == int(line) and int(sc) == 1,
               (sl, sc, el, ec))
        except Exception as e:
            ck("读回 GetSelection", False, e)

        ck("锚点复核：expect 不对就不跳",
           b.goto_location(mod, line, expect="绝不可能出现的名字") is False)

    # 没有打开代码窗的模块：必须老实返回 False
    proj = VB._active_vb_project(vbe)
    closed = None
    for comp in VB._components_of(proj):
        try:
            nm = str(comp.Name)
            if str(nm).lower() not in panes and int(comp.CodeModule.CountOfLines) > 0:
                closed = nm
                break
        except Exception:
            continue
    if closed:
        ck("没打开代码窗的模块不会假装跳成功",
           b.goto_location(closed, 1) is False, closed)
    else:
        print("（本工程每个有代码的模块都开着代码窗，跳过这一条）")

    print("\n结果: %s" % ("全部通过" if not fails else "失败 %d 项: %s"
                        % (len(fails), fails)))
    return 0 if not fails else 1


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    sys.exit(main())
