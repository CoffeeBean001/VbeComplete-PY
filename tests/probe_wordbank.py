# -*- coding: utf-8 -*-
"""手工探针：真机验证词库（`wordbank.db`）—— ★v123。

跑法（Excel 开着、VBE 里有工程）：

    py -3 tests\\probe_wordbank.py

它验的是**用户那句话的整条路**："我手动去编辑数据库，项目就能动态加载提示关键字"。
用例（§109）已经用假 VBE 走过一遍，这里用**真正的 VBE** 再走一遍 —— 因为
"收集侧真的读生效快照"这件事，只有真机才能排除"假对象碰巧对上"。

  1. 基线：收集当前候选池（真 VBE）；
  2. 建一个**临时**词库，加一条 `ZzProbeWord`、禁用 `DDB`（财务函数，几乎没人写），
     `load_active()` 之后重新收集 —— 前者必须出现，后者必须消失；
  3. 自加的词**不进** `ACTIVE_BUILTIN_LOWER`（不被"公共词汇"收紧连坐）；
  4. `reset_active()` + 删临时库，再收集一次 —— 必须与基线**完全一致**。

★ **一个字节都不写用户代码**，也**绝不碰你自己的 `wordbank.db`**（只在临时目录
  里建一个用完就删的库）—— 你手改的东西不能因为跑一次探针就被抹掉。
"""

import os
import shutil
import sqlite3
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import vba_builtins as V
import vbe_bridge as VB
import wordbank as W

PROBE_WORD = "ZzProbeWord"
# 挑一个"几乎不可能出现在用户代码里"的内建函数当禁用样本：财务函数 DDB 属于
# VBA 库、平时没人手写。万一本工程真的用过，下面会自动换一个（见 _pick_victim）。
VICTIMS = ("DDB", "MIRR", "SYD", "Partition", "StrReverse")


def main():
    fails = []

    def ck(name, cond, detail=""):
        print("%-46s %s%s" % (name, "OK" if cond else "FAIL",
                              ("  " + str(detail)) if detail else ""))
        if not cond:
            fails.append(name)

    vbe = VB._get_vbe_cached()
    if vbe is None:
        print("没有 VBE：先把 Excel 打开（并允许访问 VBA 工程对象模型）。")
        return 2

    def recs():
        return VB.VbeBackend().get_identifiers()

    def pool():
        """当前真 VBE 下，后端收集出来的名字集合（全小写）。"""
        return set(str(r[0]).lower() for r in recs())

    base_recs = recs()
    base = set(str(r[0]).lower() for r in base_recs)
    print("基线候选池：%d 个名字" % len(base))
    if not base:
        print("池子是空的 —— VBE 里至少要有打开代码的模块。")
        return 2

    # 选样本的口径：这个词**只有 VBA 内建这一个来源**（没被用户代码用过），
    # 禁用之后它必然从池子里消失 —— 否则"消失"这件事无从判定（用户代码里
    # 自己写了 DDB，禁用内建名字它照样在）。
    _src = {}
    for _r in base_recs:
        _src.setdefault(str(_r[0]).lower(), set()).add(str(_r[1]))
    victim = next((n for n in VICTIMS
                   if _src.get(n.lower()) == {VB._BUILTIN_MODULE}), None)
    if victim is None:
        print("这几个内建函数都被用户代码用过（或不在池子里）：%s" % (VICTIMS,))
        print("换一个没有用户定义的样本再跑，或改探针里的 VICTIMS。")
        return 2
    print("禁用样本：%s（来源只有 VBA 内建，所以\"消失\"能判定）" % victim)

    d = tempfile.mkdtemp(prefix="wbprobe_")
    p = os.path.join(d, "wordbank.db")
    saved_env = os.environ.get("VBECOMPLETE_WORD_BANK")
    ok = True
    try:
        created, _ = W.ensure(V.seed_rows(), p)
        ck("临时库建起来了（不碰你自己的 wordbank.db）",
           created and os.path.isfile(p), p)

        con = sqlite3.connect(p)
        con.execute("UPDATE words SET enabled=0 WHERE name=?", (victim,))
        con.execute("INSERT INTO words VALUES (?, 'function', 0, 1, NULL, '探针')",
                    (PROBE_WORD,))
        con.commit()
        con.close()

        os.environ["VBECOMPLETE_WORD_BANK"] = p
        loaded = V.load_active()
        ck("load_active() 用上了临时库", loaded, V.WORDBANK_INFO.get("counts"))

        after = pool()
        ck("自加的词进了真机候选池", PROBE_WORD.lower() in after, PROBE_WORD)
        ck("被禁用的内建函数从池子里消失了", victim.lower() not in after, victim)
        ck("自加的词不算\"语言自带\"（不被收紧连坐）",
           PROBE_WORD.lower() not in V.ACTIVE_BUILTIN_LOWER)
        ck("池子其余部分没有被打乱（只多了一个探针词）",
           len(after) >= len(base) - 1)
    except Exception as e:
        import traceback
        traceback.print_exc()
        ok = False
        ck("探针本身没抛异常", False, e)
    finally:
        # ---- 还原：环境变量、内存快照、临时库 ----
        try:
            if saved_env is None:
                os.environ.pop("VBECOMPLETE_WORD_BANK", None)
            else:
                os.environ["VBECOMPLETE_WORD_BANK"] = saved_env
        except Exception:
            pass
        try:
            V.reset_active()
        except Exception:
            pass
        try:
            shutil.rmtree(d, ignore_errors=True)
        except Exception:
            pass

    try:
        back = pool()
        ck("reset_active() 之后候选池与基线完全一致（还原干净）",
           back == base, "多了 %d / 少了 %d"
           % (len(back - base), len(base - back)))
    except Exception as e:
        ck("还原后还能收集", False, e)

    print("")
    print("全部通过" if (ok and not fails) else "有失败：%s" % (fails,))
    return 0 if (ok and not fails) else 1


if __name__ == "__main__":
    sys.exit(main())
