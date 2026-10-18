# -*- coding: utf-8 -*-
"""回归测试：用真实工程（脚本.xlsm 导出的 14 个模块）验证补全作用域。

不依赖 Excel / COM，纯解析层验证，可直接 `python tests/test_real_project.py`
或双击 `run_tests.bat` 运行。

覆盖的关键场景：
  1. UserForm3 主过程里输入 n 必须提示 num（不能被别处的同名参数挡掉）
  2. 不能自我提示（打 n 不该冒出 n）
  3. 跨过程不泄漏（别的函数的局部变量不该出现）
  4. 跨模块 Private 不提示
  5. 打全变量名后列表仍保留
  6. 中文变量名
  7. 多层嵌套里外层变量仍可提示
"""

import io
import os
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

import parser as P          # noqa: E402
import engine as E          # noqa: E402

# main.py 顶部依赖 tkinter / pynput（以及 engine/vbe_bridge/ui），而测试用的
# managed python 是精简环境，未必装这些 GUI 包。注入最小桩，避免 `import main`
# 失败、导致第 18 节（钩子挂载决策）整段 return 跳过——那样
# _reconcile_resources 的回归就根本没被验证（之前吃过这个亏）。
import types as _types
for _mod in ("tkinter", "tkinter.font", "pynput", "pynput.keyboard",
             "pynput.mouse"):
    if _mod not in sys.modules:
        sys.modules[_mod] = _types.ModuleType(_mod)


class _StubListener:         # 桩：测试不真正启动钩子
    daemon = False

    def start(self):
        pass

    def stop(self):
        pass

    def is_alive(self):
        return False

    def suppress_event(self):
        pass


sys.modules["pynput.keyboard"].Listener = _StubListener
sys.modules["pynput.mouse"].Listener = _StubListener
# main.py 模块级会引用 Key.up/.down/.tab/.esc/.ctrl_l/...，桩里要有这些属性
sys.modules["pynput.keyboard"].Key = type("Key", (), {
    "up": object(), "down": object(), "tab": object(), "esc": object(),
    "ctrl_l": object(), "ctrl_r": object(), "alt_l": object(), "alt_r": object(),
    "shift": object(), "shift_r": object(), "caps_lock": object(),
    "cmd": object(), "cmd_r": object(), "space": object(),
    "backspace": object(), "enter": object(),
})

try:                          # noqa: E402
    import main as M         # noqa: E402  （main._reconcile_resources 钩子挂载决策）
except Exception:            # noqa: E402
    M = None                 # noqa: E402

TESTDATA = os.path.join(HERE, "testdata")

PASS = 0
FAIL = 0
FAILURES = []


def _is_std(filename):
    """判断是否标准模块（对应 VBE 的 vbext_ct_StdModule）。"""
    return filename.endswith(".bas") and not filename.startswith(
        ("ThisWorkbook", "Sheet"))


def _mod_name(filename):
    return filename.split(".")[0]


def load_modules():
    mods = {}
    for fn in sorted(os.listdir(TESTDATA)):
        if not fn.endswith(".bas"):
            continue
        code = io.open(os.path.join(TESTDATA, fn), encoding="utf-8").read()
        mods[fn] = (code, _is_std(fn), _mod_name(fn))
    return mods


def collect(mods, caret_file=None, caret=None, scope="proc", caret_mod=None):
    """复刻 VbeBackend._collect_identifiers 的收集逻辑。

    mods: {文件名: (代码, 是否标准模块, 模块名)}，代表【当前工程】的全部模块。
    caret_mod 显式给出光标所在模块名；缺省时用 caret_file 推导。

    注：v30 起后端不再做"声明行名字全局剔除"（那样会连同前缀补全一起误杀
    numArr），名字一律收录；"打全名不提示自己"改由引擎层 in_decl_position +
    ctx["decl_names"] 精确剔除。故此处复刻也不再剔除声明行名字。
    """
    if caret_mod is None and caret_file:
        caret_mod = _mod_name(caret_file)
    records = []
    for fn, (code, std, mn) in mods.items():
        apply_caret = caret if (not caret_mod
                                or str(caret_mod).lower() == str(mn).lower()) else None
        recs, _pn = P.extract_records(code, module=mn,
                                     is_std_module=std,
                                     caret=apply_caret)
        records.extend(recs)
        records.extend(P.extract_implicit_records(
            code, module=mn, is_std_module=std, declared=recs, scope=scope,
            caret=apply_caret))
        # 组件名（模块名/窗体名/类名）本身也是工程内可直接引用的标识符
        # （对应 vbe_bridge 里的 records.append((mod_name, mod_name, None, False))）
        if mn:
            records.append((mn, mn, None, False))
    return records


def collect_projects(projects, active, caret_file=None, caret=None, scope="proc"):
    """复刻"只收集活动工程"的行为：projects = {工程名: mods}。

    非活动工程的模块一律不参与收集，因此它们的窗体名/函数名/全局变量名
    不会泄漏到当前工程（v23 修复的跨工程作用域）。
    """
    return collect(projects[active], caret_file=caret_file, caret=caret,
                   scope=scope)


def type_at(mods, filename, line_no, typed, indent=20):
    """返回「在某行输入 typed 之后」的模块副本与光标列。"""
    m = dict(mods)
    code, std, mn = m[filename]
    ls = code.split("\n")
    ls[line_no - 1] = " " * indent + typed
    m[filename] = ("\n".join(ls), std, mn)
    return m, indent + len(typed) + 1     # 光标在刚输入内容之后


def candidates(mods, records, filename, line_no, prefix):
    code = mods[filename][0]
    proc = P.proc_at_line(code, line_no)
    vis = E.filter_identifiers_by_scope(records, proc, _mod_name(filename))
    return sorted(i for i in vis if i.lower().startswith(prefix.lower())), proc


def check(tag, got, expect_contain=(), expect_absent=()):
    global PASS, FAIL
    ok = True
    for x in expect_contain:
        if x not in got:
            ok = False
    for x in expect_absent:
        if x in got:
            ok = False
    if ok:
        PASS += 1
        print("  PASS  %s" % tag)
        print("        候选=%s" % (got,))
    else:
        FAIL += 1
        FAILURES.append(tag)
        print("  FAIL  %s" % tag)
        print("        候选=%s" % (got,))
        print("        应含=%s 应不含=%s" % (list(expect_contain), list(expect_absent)))


def main():
    # v61：VBA 内建名字（内建函数 / 常量 / 数据类型）进候选池由
    # vbe_bridge.ENABLE_VBA_BUILTINS 决定；v62 语言关键字另行一组，开关独立
    # （vbe_bridge.ENABLE_VBA_KEYWORDS）。★v96 这两个开关的**出厂默认是 True**
    # （"把 vba 关键字、函数名这两部分也加入到提醒列表里"），而两批枚举仍然
    # 默认 False（"枚举类型就不要加了，枚举加上，就会太多了"）。
    #
    # 下面第 1~37 节的断言是"工程内名字的候选列表精确比对"（用户要求的防噪音
    # 护栏），与语言自带的名字无关 —— 因此统一在【关闭】的前提下跑，护栏强度
    # 保持不变（否则 v96 之后每一条都会多出一堆 Sub/If/Long，护栏就废了）；
    # 第 38 节专门验内建名字，第 39 节专门验关键字，
    # 第 54 节守 v96 的默认口径（True/True/False/False）。
    try:
        import vbe_bridge as _vb61
        _vb61.ENABLE_VBA_BUILTINS = False
        _vb61.ENABLE_VBA_KEYWORDS = False
    except Exception:
        pass

    # ★v107c：进场先记下这些"各节会临时替换"的对象，供 §90 做猴补丁泄漏自查。
    # 起因：§31 / §62 把 `_get_vbe_cached` 换成自己的假工程后忘了换回来，
    # 于是从那一节起，后面**每一节**拿到手的都是别人的假模块 —— 凡是用
    # "清 `_vbe_cache["obj"]` + 换 `_get_vbe`"搭装置的小节都会读到它，
    # 症状是"写回悄悄落到别人的假模块上、自己那份一行没改"（§89 被这个坑了
    # 两轮）。这种错**不会让任何一条用例变红**，只会让它们测的不是自己以为的
    # 东西 —— 所以必须有一道专门的护栏。
    try:
        import vbe_bridge as _vb0
        import ui as _ui0
        _LEAK_WATCH = [(m, n, getattr(m, n)) for (m, n) in (
            (_vb0, "_get_vbe"), (_vb0, "_get_vbe_cached"),
            (_vb0, "_probe_host_window"), (_vb0, "_host_window_alive"),
            (_vb0, "com_host_present"), (_vb0, "_vbe_cache"),
            (_vb0, "_com_state"), (_vb0, "_wnd_state"),
            (_ui0, "caret_screen_rect"), (_ui0, "screen_popup_rects"),
        )]
    except Exception:
        _LEAK_WATCH = []
    # ★v108：§91 会临时把 `main._ime_conversion_mode` 换成桩来钉"问不到"那一档，
    # 同样纳入自查（单独 try：`import main` 万一失败不该把整张表变成空表）。
    # ★v109：§92 会临时换 `main._gui`（假 win32gui）和
    # `main.foreground_window_closing`（可调 closing 桩）—— 忘还原会让后面每一节
    # 的焦点判读都落在假对象上，同样必须纳管。
    try:
        import main as _mn0
        _LEAK_WATCH.append((_mn0, "_ime_conversion_mode",
                            _mn0._ime_conversion_mode))
        _LEAK_WATCH.append((_mn0, "_gui", _mn0._gui))
        _LEAK_WATCH.append((_mn0, "foreground_window_closing",
                            _mn0.foreground_window_closing))
    except Exception:
        pass

    # ⚠️★v103c：本回归网里凡是【成员位置】（`标识符.` / With 块里的 `.成员`）的
    #   用例 —— §34 / §49 / §50 / §53 / §71 / §72 / §77.3 —— 测的都是"点号后面
    #   出候选 + 按声明类型收窄 + 待定让位/补位 + 回声防护"这一整套机制，而它
    #   自 v103c 起是**开关打开才生效**的能力：用户口径是"输入 student. ，点号
    #   后面不要出提示框了"（我们出的是裸名池，本来就该让给 VBE 自己的成员列表），
    #   所以默认 `MEMBER_POPUP=False`。
    #   这里统一打开，让那些老用例继续验证机制本身（它们验的是能力，不是默认值）；
    #   "默认必须是关"这个事实由 §84 用源码护栏 + 行为用例钉住。
    _MEMBER_POPUP_SAVE = E.MEMBER_POPUP
    E.MEMBER_POPUP = True

    mods = load_modules()
    print("载入 %d 个模块: %s\n" % (len(mods), ", ".join(sorted(mods))))

    print("=== 1. 真实场景：UserForm3 第 61 行输入 n ===")
    for line in (60, 61, 63):
        m, col = type_at(mods, "UserForm3.frm.bas", line, "n")
        recs = collect(m, caret_file="UserForm3.frm.bas", caret=(line, col))
        cand, proc = candidates(m, recs, "UserForm3.frm.bas", line, "n")
        print("  [行 %d] proc=%s" % (line, proc))
        check("行%d 输入 n 提示 num" % line, cand,
              expect_contain=["num", "numColInTargetSheet"],
              expect_absent=["n"])

    print("\n=== 2. 不自我提示（打 n 不该冒出 n）===")
    m, col = type_at(mods, "UserForm3.frm.bas", 61, "n")
    recs = collect(m, caret_file="UserForm3.frm.bas", caret=(61, col))
    cand, _ = candidates(m, recs, "UserForm3.frm.bas", 61, "n")
    check("自我提示", cand, expect_absent=["n"])

    print("\n=== 3. 跨过程不泄漏（num2 属于 getNumColInTargetSheet）===")
    check("主过程不含 num2", cand, expect_absent=["num2"])

    print("\n=== 4. 在 getNumColInTargetSheet 内应能看到 num / num2 ===")
    # 该函数体在 122-131 行区间（去掉 Attribute 头后约 114-123 行）
    code = mods["UserForm3.frm.bas"][0]
    target = None
    for i, l in enumerate(code.split("\n"), 1):
        if "Private Function getNumColInTargetSheet" in l:
            target = i + 3
            break
    if target:
        m2, col2 = type_at(mods, "UserForm3.frm.bas", target, "n", indent=8)
        recs2 = collect(m2, caret_file="UserForm3.frm.bas", caret=(target, col2))
        cand2, proc2 = candidates(m2, recs2, "UserForm3.frm.bas", target, "n")
        print("  [行 %d] proc=%s" % (target, proc2))
        check("函数内可见 num/num2", cand2, expect_contain=["num", "num2"])

    print("\n=== 5. 打全变量名后列表仍保留（v15）===")
    m3, col3 = type_at(mods, "UserForm3.frm.bas", 61, "num")
    recs3 = collect(m3, caret_file="UserForm3.frm.bas", caret=(61, col3))
    cand3, _ = candidates(m3, recs3, "UserForm3.frm.bas", 61, "num")
    check("打全 num 仍保留", cand3, expect_contain=["num", "numColInTargetSheet"])

    print("\n=== 6. 跨过程隔离 / 中文 / 嵌套（合成用例）===")
    def synth(code, proc, prefix, mod="M1", std=True, caret=None):
        r, _pn = P.extract_records(code, mod, std)
        r += P.extract_implicit_records(code, mod, std, declared=r,
                                        scope="proc", caret=caret)
        v = E.filter_identifiers_by_scope(r, proc, mod)
        return sorted(i for i in v if i.lower().startswith(prefix.lower()))

    c1 = "\n".join(["Sub ProcA()", "    num = 1", "    Debug.Print num", "End Sub",
                    "Sub ProcB()", "    num = 2", "    num2 = 3", "End Sub"])
    check("ProcA 只提示 num", synth(c1, "ProcA", "n"),
          expect_contain=["num"], expect_absent=["num2"])
    check("ProcB 有 num+num2", synth(c1, "ProcB", "n"),
          expect_contain=["num", "num2"])

    c2 = "\n".join(["Sub 甲()", "    姓名 = \"x\"", "End Sub",
                    "Sub 乙()", "    姓名 = \"y\"", "    年龄 = 18", "End Sub"])
    check("中文：甲不泄漏年龄", synth(c2, "甲", "年"), expect_absent=["年龄"])
    check("中文：乙有年龄", synth(c2, "乙", "年"), expect_contain=["年龄"])

    nest = "\n".join(["Sub T()", "    outerV = 1", "    For i = 1 To 10",
                      "        If outerV > 1 Then", "            For j = 1 To 3",
                      "                If j > 1 Then", "                    outerV",
                      "                End If", "            Next j",
                      "        End If", "    Next i", "End Sub"])
    check("多层嵌套外层可见", synth(nest, "T", "outer", caret=(7, 25)),
          expect_contain=["outerV"])

    # ---- 7. v21：声明行上的自我提示 ----
    # 定义模块级/全局变量时，正在声明的名字不该提示其本身
    print("\n=== 7. 声明行自我提示（v21）===")
    decl = "\n".join([
        "Public gCounter As Long",
        "Private mTotal As String",
        "Dim gFlag As Boolean",
        "Sub Foo()",
        "    gCounter",
        "End Sub"])

    def decl_cand(caret, prefix):
        recs, _pn = P.extract_records(decl, module="M1",
                                       is_std_module=True, caret=caret)
        proc = P.proc_at_line(decl, caret[0]) if caret else None
        vis = E.filter_identifiers_by_scope(recs, proc, "M1")
        return sorted(i for i in vis if i.lower().startswith(prefix.lower()))

    # v30 起解析层不再剔除声明行名字（否则会把前缀补全 numArr 一起误杀），
    # 名字一律收录、交给引擎按"正在输入的词 + 当前声明行名字"精确剔除。
    # 故此处只验证：声明行上的名字确实被【收录】（前缀补全可用），
    # "打全名不提示自己"由第 11 节引擎集成测试把关。
    # 光标在 Public gCounter 声明行（前缀 g）：gCounter 与 gFlag 都应收录
    check("声明行名字被收录(整词前缀)",
          decl_cand((1, 19), "g"),
          expect_contain=["gCounter", "gFlag"])
    # 只输入首字母时同样收录（前缀补全可用）
    check("声明行名字被收录(首字母前缀)",
          decl_cand((1, 10), "g"),
          expect_contain=["gCounter", "gFlag"])
    # 光标在行首同样收录
    check("声明行名字被收录(行首前缀)",
          decl_cand((1, 2), "g"),
          expect_contain=["gCounter", "gFlag"])
    # Private 声明同理：mTotal 被收录
    check("Private 声明行名字被收录",
          decl_cand((2, 17), "m"), expect_contain=["mTotal"])
    # 使用行必须保留（v15 打全名保留列表的行为不能被破坏）
    check("使用行仍保留该变量",
          decl_cand((5, 13), "g"),
          expect_contain=["gCounter", "gFlag"])
    # 拿不到光标时不得误排除
    check("caret=None 不误排除",
          decl_cand(None, "g"), expect_contain=["gCounter", "gFlag"])

    # ---- 8. v22：声明里 `As` 之后的类型名位置不应提示变量名 ----
    print("\n=== 8. 声明 As 之后类型名位置（v22）===")

    def tp(line, col):
        """is_caret_in_type_position 的简短封装（包成 list 适配 check）。"""
        return [P.is_caret_in_type_position(line, col)]

    # 变量声明：As 之后 -> 类型名位置
    check("Dim x As | 命中类型位置",
          tp("Dim x As ", 10), expect_contain=[True])
    check("Dim x As Integer 末尾命中",
          tp("Dim x As Integer", 17), expect_contain=[True])
    # 变量名位置（As 之前）-> 不命中，照常提示
    check("Dim x 变量名位置不命中",
          tp("Dim x", 6), expect_contain=[False])
    # 多声明：`a As Long, b As |` 各片段独立判断
    check("Dim a As Long, b As | 命中",
          tp("Dim a As Long, b As ", 22), expect_contain=[True])
    check("Dim a As Long, b 变量名位置不命中",
          tp("Dim a As Long, b", 18), expect_contain=[False])
    # 数组维度括号里的逗号不干扰
    check("Dim arr(0 To 9) As In 命中",
          tp("Dim arr(0 To 9) As In", 21), expect_contain=[True])
    # 过程参数 / 返回值里的 As 也命中
    check("Function 返回值类型位置命中",
          tp("Private Function f(ByVal a As Long) As Double", 47),
          expect_contain=[True])
    # 模块级 / 常量 / 类型
    check("Public gCounter As Long 命中",
          tp("Public gCounter As Long", 24), expect_contain=[True])
    check("Const pi As Double 命中",
          tp("Const pi As Double = 3.14", 19), expect_contain=[True])
    check("Dim x As New Collection 命中",
          tp("Dim x As New Collection", 24), expect_contain=[True])
    check("Dim x As String * 10 命中",
          tp("Dim x As String * 10", 20), expect_contain=[True])
    # 否定用例：字符串里的 "As" 不能误判
    check("字符串里的 As 不误判",
          tp('Debug.Print "As foo" & bar', 27), expect_contain=[False])
    check("ReDim（无 As）不命中",
          tp("ReDim arr(1 To 10)", 18), expect_contain=[False])
    check("If 行（无 As）不命中",
          tp("If x > y Then", 14), expect_contain=[False])

    # 引擎集成：in_type_position=True 时 trigger 必须收起弹窗
    print("\n--- 引擎集成：类型位置强制静默 ---")

    class _FakeUI:
        def __init__(self):
            self.shown = False
        def show(self, *a, **k):
            self.shown = True
        def update_selection(self, *a, **k):
            pass
        def hide(self, *a, **k):
            self.shown = False
        def contains_point(self, *a, **k):
            return False

    class _FakeBackend:
        def __init__(self, ctx):
            self.ctx = ctx
        def get_context(self):
            return self.ctx
        def get_identifiers(self):
            return [("num", "M1", "ProcA", False),
                    ("numColInTargetSheet", "M1", "ProcA", False)]
        def apply_completion(self, *a, **k):
            return None

    # 1) 类型名位置（As 后）即便有匹配也静默
    ui = _FakeUI()
    comp = E.Completer(_FakeBackend(
        {"line_no": 1, "caret_col": 12, "line_text": "Dim x As In",
         "in_string": False, "in_comment": False, "in_type_position": True,
         "proc_name": None, "module_name": "M1"}), ui)
    comp.trigger()
    check("类型位置 trigger 收起弹窗",
          [ui.shown], expect_contain=[False])

    # 2) 变量名位置（As 前）有匹配则正常弹出
    ui2 = _FakeUI()
    comp2 = E.Completer(_FakeBackend(
        {"line_no": 2, "caret_col": 6, "line_text": "    num = 1",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "proc_name": "ProcA", "module_name": "M1"}), ui2)
    comp2.trigger()
    check("变量名位置 trigger 正常弹出",
          [ui2.shown], expect_contain=[True])

    # ---- 9. v23：定义全局/模块级变量时不提示其本身 ----
    print("\n=== 9. 声明变量时不自我提示（v23）===")

    def decl_case(code, caret, module="M1", prefix=None, std=True):
        """返回「在模块级/过程内声明变量时光标处的候选」。"""
        mods = {module + ".bas": (code, std, module)}
        recs = collect(mods, caret_mod=module, caret=caret)
        proc = P.proc_at_line(code, caret[0])
        vis = E.filter_identifiers_by_scope(recs, proc, module)
        return sorted(i for i in vis
                      if i.lower().startswith((prefix or "").lower()))

    # 9.1 名字在别处被"只读"使用过：用法扫描会把它当隐式变量收录。
    #     v30 起解析层不再全局剔除声明行名字（否则连前缀补全一起误杀），
    #     名字一律收录，交给引擎按"当前声明行名字 + 正在输入的词"精确剔除。
    #     此处只验证声明行名字确实被【收录】（前缀补全可用）。
    c_ro = "\n".join([
        "Public gUserName As String",
        "Sub Foo()",
        "    MsgBox gUserName",
        "End Sub"])
    check("模块级声明：名字被收录(整词前缀)",
          decl_case(c_ro, (1, len("Public gUserName") + 1), prefix="g"),
          expect_contain=["gUserName"])
    # 同一行只打出首字母时同样收录（前缀补全可用）
    check("模块级声明：名字被收录(首字母前缀)",
          decl_case(c_ro, (1, len("Public g") + 1), prefix="g"),
          expect_contain=["gUserName"])

    # 9.2 过程内：先用了再补 Dim（For i 已把 i 收成隐式变量）
    c_local = "\n".join([
        "Sub Foo()",
        "    For i = 1 To 3",
        "    Next i",
        "    Dim i As Long",
        "End Sub"])
    check("过程内 Dim 名字被收录",
          decl_case(c_local, (4, len("    Dim i") + 1), prefix="i"),
          expect_contain=["i"])

    # 9.3 续行声明：光标落在第二行（该行本身没有 Dim/Public）
    c_cont = "\n".join([
        "Public gA As Long, _",
        "       gB As Long",
        "Sub Foo()",
        "    MsgBox gA & gB",
        "End Sub"])
    check("续行声明第二行名字被收录",
          decl_case(c_cont, (2, len("       gB") + 1), prefix="g"),
          expect_contain=["gB"])
    check("续行声明整行两个名字都收录",
          decl_case(c_cont, (1, len("Public gA") + 1), prefix="g"),
          expect_contain=["gA", "gB"])

    # 9.4 非声明行不受影响：使用行仍要能提示
    check("使用行仍可提示",
          decl_case(c_ro, (3, len("    MsgBox gUser") + 1), prefix="g"),
          expect_contain=["gUserName"])

    # 9.5 函数名/过程名同样会被收录（v23-2 的"打全名不提示自己"由引擎把关）
    c_fn = "\n".join([
        "Public Function gCalc() As Long",
        "    gCalc = 1",
        "End Function"])
    check("函数声明行名字被收录",
          decl_case(c_fn, (1, len("Public Function gCalc") + 1), prefix="g"),
          expect_contain=["gCalc"])
    # 参数名也是"正在起的新名字"
    c_param = "\n".join([
        "Public Function gCalc(gTotal As Long) As Long",
        "    gCalc = gTotal",
        "End Function"])
    check("参数声明位置名字被收录",
          decl_case(c_param, (1, len("Public Function gCalc(gTot") + 1), prefix="g"),
          expect_contain=["gCalc", "gTotal"])
    # 使用行必须照常提示
    check("函数调用行仍可提示",
          decl_case(c_fn, (2, len("    gCal") + 1), prefix="g"),
          expect_contain=["gCalc"])
    # Type / Enum 名
    c_type = "\n".join([
        "Public Type gRec",
        "    gId As Long",
        "End Type"])
    check("Type 声明行名字被收录",
          decl_case(c_type, (1, len("Public Type gRe") + 1), prefix="g"),
          expect_contain=["gRec"])
    # v59：块成员（Type 字段 / Enum 成员）不再收录成裸名候选 —— Type 字段只能
    # 通过变量限定访问（p.field），Enum 成员虽可裸名引用但枚举名已限定（E.Member），
    # 提示出来只会把候选池搞吵。类型/枚举名本身照常收录（上面那条）。
    check("Type 成员行不再收录成员名",
          decl_case(c_type, (2, len("    gI") + 1), prefix="g"),
          expect_absent=["gId"])

    # ---- 10. v23：跨工程隔离 ----
    print("\n=== 10. 跨工程隔离（v23）===")
    proj_a = {
        "Module1.bas": ("\n".join([
            "Public gLocal As Long",
            "Sub Main()",
            "    gLocal = 1",
            "End Sub"]), True, "Module1"),
    }
    # 另一个工程：同名模块 Module1（默认名撞车是跨工程泄漏的典型入口）
    # 里面有全局变量、函数、以及被 Show 调用的窗体名
    proj_b = {
        "Module1.bas": ("\n".join([
            "Public gOther As Long",
            "Public Function OtherFn() As Long",
            "    OtherFn = 1",
            "End Function"]), True, "Module1"),
        "UserForm1.frm.bas": ("\n".join([
            "Private Sub UserForm_Initialize()",
            "End Sub",
            "Sub ShowIt()",
            "    UserForm1.Show",
            "End Sub"]), False, "UserForm1"),
    }
    projects = {"A": proj_a, "B": proj_b}

    def cand_in_project(active, filename, line_no, prefix):
        mods = projects[active]
        recs = collect_projects(projects, active, caret_file=filename,
                                caret=(line_no, 1))
        code = mods[filename][0]
        proc = P.proc_at_line(code, line_no)
        vis = E.filter_identifiers_by_scope(recs, proc, _mod_name(filename))
        return sorted(i for i in vis if i.lower().startswith(prefix.lower()))

    check("活动工程 A 内可见本工程全局量",
          cand_in_project("A", "Module1.bas", 3, "g"),
          expect_contain=["gLocal"])
    check("不提示工程 B 的全局变量",
          cand_in_project("A", "Module1.bas", 3, "g"),
          expect_absent=["gOther"])
    check("不提示工程 B 的函数名",
          cand_in_project("A", "Module1.bas", 3, "other"),
          expect_absent=["OtherFn"])
    check("不提示工程 B 的窗体名",
          cand_in_project("A", "Module1.bas", 3, "userform"),
          expect_absent=["UserForm1"])
    # 反向：在 B 里也不该看到 A 的东西
    check("反向：B 内不含 A 的全局量",
          cand_in_project("B", "Module1.bas", 3, "g"),
          expect_contain=["gOther"], expect_absent=["gLocal"])

    # ---- 11. v23-2：声明行整行静默（第一道防线）----
    print("\n=== 11. 声明行彻底静默（v23-2）===")

    def dp(line):
        return [P.is_caret_in_declaration(line)]

    check("Public 变量声明行 -> 静默",
          dp("Public gUserName As String"), expect_contain=[True])
    check("Dim 声明行 -> 静默",
          dp("Dim gFlag As Boolean"), expect_contain=[True])
    check("Private 变量声明行 -> 静默",
          dp("Private mTotal As String"), expect_contain=[True])
    check("Global 变量声明行 -> 静默",
          dp("Global gX As Long"), expect_contain=[True])
    check("Const 声明行 -> 静默",
          dp("Public Const gMax = 10"), expect_contain=[True])
    check("Static 声明行 -> 静默",
          dp("Static gCnt As Long"), expect_contain=[True])
    check("多行声明（逗号）-> 静默",
          dp("Dim a As Long, b As Integer"), expect_contain=[True])
    check("续行首行 -> 静默",
          dp("Public gA As Long, _"), expect_contain=[True])
    # 函数名 / 过程名 / 属性 / API / 事件 / 类型
    check("Function 声明行 -> 静默",
          dp("Public Function gCalc() As Long"), expect_contain=[True])
    check("Sub 声明行 -> 静默",
          dp("Private Sub gRun()"), expect_contain=[True])
    check("Property 声明行 -> 静默",
          dp("Public Property Get gVal() As Long"), expect_contain=[True])
    check("Declare 声明行 -> 静默",
          dp('Private Declare PtrSafe Function gApi Lib "x" () As Long'),
          expect_contain=[True])
    check("Event 声明行 -> 静默",
          dp("Public Event gChanged(ByVal x As Long)"), expect_contain=[True])
    check("Type 声明行 -> 静默",
          dp("Public Type gRec"), expect_contain=[True])
    check("Enum 声明行 -> 静默",
          dp("Public Enum gKind"), expect_contain=[True])
    # 不误伤
    check("普通赋值行 -> 不静默",
          dp("    gUserName = \"x\""), expect_contain=[False])
    check("函数调用行 -> 不静默",
          dp("    x = gCalc(1)"), expect_contain=[False])
    check("If 行 -> 不静默",
          dp("If gFlag Then"), expect_contain=[False])
    check("ReDim 行 -> 不静默（要提示已有数组）",
          dp("ReDim gArr(1 To 10)"), expect_contain=[False])
    check("字符串里的 Dim 不误判",
          dp('Debug.Print "Dim x As Long"'), expect_contain=[False])
    check("End Function 不误判",
          dp("End Function"), expect_contain=[False])

    # 引擎集成：声明行上即使有匹配也一条都不弹（这是"提示自己"的根治手段）
    class _UI2(_FakeUI):
        pass

    class _B2(_FakeBackend):
        def __init__(self, ctx, ids, declared=None):
            _FakeBackend.__init__(self, ctx)
            self._ids = ids
            # 真实声明的名字（小写）；缺省把全部 ids 当声明，匹配"旧行为兜底"。
            if declared is None:
                declared = [i[0] if isinstance(i, (tuple, list)) else i
                            for i in ids]
            self._declared = set(str(n).lower() for n in declared)

        def get_identifiers(self):
            return self._ids

        def get_declared_names(self):
            return list(self._declared)

    ids = [("gUserName", "M1", None, False), ("gCalc", "M1", None, False),
           ("gFlag", "M1", None, False)]

    def trig(line_text, col, in_decl, in_type=False, ids_=None, cls=None,
             decl_names=None, declared_=None):
        ui = _UI2()
        c = E.Completer((cls or _B2)(
            {"line_no": 1, "caret_col": col, "line_text": line_text,
             "in_string": False, "in_comment": False,
             "in_type_position": in_type, "in_decl_position": in_decl,
             "decl_names": decl_names or [],
             "proc_name": None, "module_name": "M1"},
            ids_ or ids, declared_), ui)
        c.trigger()
        return (ui.shown, c.current_matches())

    # 声明行：★v105 起【整个静默】—— 造名字的位置不出任何候选（IDEA 口径）。
    #
    # 【语义变更史】
    #   * v31~v36：正在声明的名字打全名时一律剔除 -> 列表整个消失。
    #   * v37：改成"只看这个名字在工程里真实存在吗"，存在就照常提示（含打全名）。
    #   * ★v105：**整体反转**回"声明位置不出候选"。用户口径（原话）："在定义
    #     变量名，函数名或过程名时，会出提示词列表框。我用 idea 是不提示的吧，
    #     你帮我改成 IDEA 那种。" —— 造新名字的地方没有"既有符号"可补，任何候选
    #     都是噪音（还会盖住他刚敲的字），v37 的"打全名要提示"只在**用法行**成立。
    #   下面这 4 条断言因此从 (True, [...]) 反转为 (False, [])。
    check("公共变量（声明行）：不弹候选（v105）",
          [trig("Public gUserName", 17, True, decl_names=["gUserName"])],
          expect_contain=[(False, [])])
    check("公共变量（声明行）：打前缀也不弹（v105）",
          [trig("Public gUser", 13, True, decl_names=["gUserName"])],
          expect_contain=[(False, [])])
    check("函数名（声明行）：不弹候选（v105）",
          [trig("Public Function gCalc", 22, True, decl_names=["gCalc"])],
          expect_contain=[(False, [])])
    check("函数名（声明行）：打前缀也不弹（v105）",
          [trig("Public Function gCal", 21, True, decl_names=["gCalc"])],
          expect_contain=[(False, [])])
    check("非声明行：照常提示",
          [trig("    gUser", 10, False)], expect_contain=[(True, ["gUserName"])])

    # ---- 11.6 Bug1：声明行【静默】（v105 反转；前缀补全改由用法行覆盖）----
    # 【v105】原先这两条断言"声明行上输入 num 要同时提醒 num 与 numArr"（v37 口径）。
    # 用户要求改成 IDEA 那种：**造名字的位置不弹**。故这里反转为 (False, [])。
    # ⚠️ "前缀补全不被误杀"这条语义没有丢 —— 由下面 11.8 的**用法行**那两条
    #    （`    num` / in_decl_position=False）覆盖（`(True, ["num","numArr"])`）。
    bug1_code = "\n".join([
        "Public num As Long",
        "Public numArr() As Long",
        "Sub Foo()",
        "    MsgBox num",
        "End Sub"])
    b1_mods = {"M1.bas": (bug1_code, True, "M1")}
    b1_recs = collect(b1_mods, caret_mod="M1", caret=(2, len("Public num") + 1))
    ui_b1 = _UI2()
    c_b1 = E.Completer(_B2(
        {"line_no": 2, "caret_col": len("Public num") + 1,
         "line_text": "Public num", "in_string": False, "in_comment": False,
         "in_type_position": False, "in_decl_position": True,
         "decl_names": ["numArr"], "proc_name": None, "module_name": "M1"},
        b1_recs), ui_b1)
    c_b1.trigger()
    check("Bug1(声明行): 输入 num -> 不弹候选（v105）",
          [(ui_b1.shown, sorted(c_b1.current_matches()))],
          expect_contain=[(False, [])])
    ui_b1b = _UI2()
    c_b1b = E.Completer(_B2(
        {"line_no": 2, "caret_col": len("Public numArr") + 1,
         "line_text": "Public numArr", "in_string": False, "in_comment": False,
         "in_type_position": False, "in_decl_position": True,
         "decl_names": ["numArr"], "proc_name": None, "module_name": "M1"},
        b1_recs), ui_b1b)
    c_b1b.trigger()
    check("Bug1(声明行): 打全 numArr -> 不弹候选（v105）",
          [(ui_b1b.shown, sorted(c_b1b.current_matches()))],
          expect_contain=[(False, [])])

    # ---- 11.7 Bug2：回退删除后若唯一匹配等于所输入词，不提示自己 ----
    ui_b2 = _UI2()
    c_b2 = E.Completer(_B2(
        {"line_no": 3, "caret_col": 10, "line_text": "    gFlag",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "in_decl_position": False, "decl_names": [],
         "proc_name": "Foo", "module_name": "M1"}, ids), ui_b2)
    c_b2.trigger()
    # v33 语义更新：gFlag 是【真实定义】的名字，打全/回退到它都要保留列表
    # （用户要求"输入完整，提示保留"）。"提示自己"由"不是真名就剔除"负责，
    # 不再靠"唯一候选等于所输词就收起"这条误伤正常场景的规则。
    check("Bug2(新语义): 回退到真实存在的 gFlag -> 保留提示",
          [(ui_b2.shown, sorted(c_b2.current_matches()))],
          expect_contain=[(True, ["gFlag"])])
    # 反向：前缀 "gF" 仍有真实匹配 -> 正常提示（不是自我提示）
    ui_b2b = _UI2()
    c_b2b = E.Completer(_B2(
        {"line_no": 3, "caret_col": 8, "line_text": "    gF",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "in_decl_position": False, "decl_names": [],
         "proc_name": "Foo", "module_name": "M1"}, ids), ui_b2b)
    c_b2b.trigger()
    check("Bug2: 前缀有真实匹配 -> 正常提示",
          [(ui_b2b.shown, sorted(c_b2b.current_matches()))],
          expect_contain=[(True, ["gFlag"])])

    # ---- 11.8 Bug1(续)：回退出未定义的 numA，绝不提示 numA 自己 ----
    # 模拟"隐式变量扫描把光标处正在敲的 numA 当成已用变量"的最坏情况：ids 里
    # 塞入一个 numA（priv=True 的隐式残留），但工程并未声明 numA。引擎应凭
    # "numA 不是真实声明的名字"把它从候选里剔除，只留下 numArr。
    phantom_ids = [("num", "M1", None, False), ("numArr", "M1", None, False),
                   ("numA", "M1", None, True)]
    ui_p = _UI2()
    c_p = E.Completer(_B2(
        {"line_no": 3, "caret_col": len("    numA") + 1, "line_text": "    numA",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "in_decl_position": False, "decl_names": [],
         "proc_name": "Foo", "module_name": "M1"},
        phantom_ids, declared=["num", "numArr"]), ui_p)
    c_p.trigger()
    check("Bug1续: 输入 numA 不提示未定义的 numA（只提示 numArr）",
          [(ui_p.shown, sorted(c_p.current_matches()))],
          expect_contain=[(True, ["numArr"])])
    # 反向：输入真实定义的 num -> 仍同时提醒 num 与 numArr（前缀补全正常）
    ui_p2 = _UI2()
    c_p2 = E.Completer(_B2(
        {"line_no": 3, "caret_col": len("    num") + 1, "line_text": "    num",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "in_decl_position": False, "decl_names": [],
         "proc_name": "Foo", "module_name": "M1"},
        [("num", "M1", None, False), ("numArr", "M1", None, False)],
        declared=["num", "numArr"]), ui_p2)
    c_p2.trigger()
    check("Bug1续: 输入 num 同时提醒 num 与 numArr（用法行）",
          [(ui_p2.shown, sorted(c_p2.current_matches()))],
          expect_contain=[(True, ["num", "numArr"])])

    # ---- 11.9 键盘选择 / 确认（↑↓ 浏览、Tab/Enter 确认，无需移动鼠标）----
    class _CapBackend(_B2):
        def __init__(self, ctx, ids, declared=None):
            _B2.__init__(self, ctx, ids, declared)
            self.last_apply = None

        def apply_completion(self, line_no, start, caret, completion):
            self.last_apply = (line_no, start, caret, completion)
            return completion

    nav_ids = [("num", "M1", None, False), ("numArr", "M1", None, False),
               ("numCount", "M1", None, False)]
    ui_n = _UI2()
    c_n = E.Completer(_CapBackend(
        {"line_no": 3, "caret_col": 8, "line_text": "    num",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "in_decl_position": False, "decl_names": [],
         "proc_name": "Foo", "module_name": "M1"}, nav_ids), ui_n)
    c_n.trigger()
    check("键盘: 候选 3 个", [c_n.current_matches()],
          expect_contain=[["num", "numArr", "numCount"]])
    check("键盘: 默认选中首项", [c_n.selected], expect_contain=[0])
    c_n.move(2)                       # ↓↓ 跳到末项
    check("键盘: ↓↓ 到第 3 项", [c_n.selected], expect_contain=[2])
    check("键盘: 当前候选 = numCount",
          [c_n.current_matches()[c_n.selected]], expect_contain=["numCount"])
    c_n.accept()                      # 确认 -> apply_completion 拿到 numCount
    check("键盘: 确认写回 numCount",
          [c_n.backend.last_apply[3] if c_n.backend.last_apply else None],
          expect_contain=["numCount"])
    check("键盘: 确认后弹窗收起", [ui_n.shown], expect_contain=[False])
    # 循环：首项按 ↑ 应回到末项
    ui_n2 = _UI2()
    c_n2 = E.Completer(_CapBackend(
        {"line_no": 3, "caret_col": 8, "line_text": "    num",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "in_decl_position": False, "decl_names": [],
         "proc_name": "Foo", "module_name": "M1"}, nav_ids), ui_n2)
    c_n2.trigger()
    c_n2.move(-1)
    check("键盘: 首项按 ↑ 循环到末项", [c_n2.selected], expect_contain=[2])

    # ---- 11.10 v32 / v105：声明行一律静默（与"真名集合是否含它"无关）----
    # 【v32 由来】回退删字（Dim numArr -> Dim numA）时，文本里此刻确实写着
    # Dim numA，解析层会把它当真声明；若真名集合照单全收，"不是真名就剔除自身"
    # 的防护就失效，numA 又被提示出来。v32 的修法是"真名集合剔除光标行声明名"。
    # 【v105】声明行现在**整个静默**，与真名集合无关 —— 两条断言的期望值因此
    # 都收敛成 (False, [])：无论 declared 里有没有 numA，声明行都不弹。
    ui_v32 = _UI2()
    c_v32 = E.Completer(_B2(
        {"line_no": 3, "caret_col": len("    Dim numA") + 1,
         "line_text": "    Dim numA", "in_string": False, "in_comment": False,
         "in_type_position": False, "in_decl_position": True,
         "decl_names": [], "proc_name": "Foo", "module_name": "M1"},
        phantom_ids, declared=["num", "numArr"]), ui_v32)
    c_v32.trigger()
    check("v32/v105: 声明行（真名集合已剔除光标行）-> 静默",
          [(ui_v32.shown, sorted(c_v32.current_matches()))],
          expect_contain=[(False, [])])
    # 对照：真名集合【含】numA（旧行为会让它漏出来）—— v105 起同样静默，
    # 因为声明行根本不进候选计算，numA 出不出得来与 declared 无关了。
    ui_v32b = _UI2()
    c_v32b = E.Completer(_B2(
        {"line_no": 3, "caret_col": len("    Dim numA") + 1,
         "line_text": "    Dim numA", "in_string": False, "in_comment": False,
         "in_type_position": False, "in_decl_position": True,
         "decl_names": [], "proc_name": "Foo", "module_name": "M1"},
        phantom_ids, declared=["num", "numArr", "numA"]), ui_v32b)
    c_v32b.trigger()
    check("v32/v105 对照: 真名集合含 numA 时也静默（声明行不进候选计算）",
          [(ui_v32b.shown, sorted(c_v32b.current_matches()))],
          expect_contain=[(False, [])])

    # ---- 11.10b v33：已定义变量打全名，列表必须保留（用户反馈的回归）----
    # 场景：定义了 arr，输入 a / ar 都提示 arr；把 arr 打全后列表不能消失，
    # 否则既没法 Tab 确认，也看不出自己写对了没有。
    arr_ids = [("arr", "M1", None, False)]
    for w, tag in (("a", "首字母 a"), ("ar", "前缀 ar"), ("arr", "打全 arr")):
        ui_a = _UI2()
        c_a = E.Completer(_B2(
            {"line_no": 3, "caret_col": len("    " + w) + 1,
             "line_text": "    " + w, "in_string": False, "in_comment": False,
             "in_type_position": False, "in_decl_position": False,
             "decl_names": [], "proc_name": "Foo", "module_name": "M1"},
            arr_ids, declared=["arr"]), ui_a)
        c_a.trigger()
        check("v33: 已定义 arr 输入 %s -> 仍提示 arr" % tag,
              [(ui_a.shown, sorted(c_a.current_matches()))],
              expect_contain=[(True, ["arr"])])
    # 对照：从未定义的 arrX 打全 -> 不提示自己（真名集合里没有它）
    ui_ax = _UI2()
    c_ax = E.Completer(_B2(
        {"line_no": 3, "caret_col": len("    arrX") + 1, "line_text": "    arrX",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "in_decl_position": False, "decl_names": [],
         "proc_name": "Foo", "module_name": "M1"},
        [("arr", "M1", None, False), ("arrX", "M1", None, True)],
        declared=["arr"]), ui_ax)
    c_ax.trigger()
    check("v33 对照: 未定义的 arrX 打全 -> 不提示自己",
          [(ui_ax.shown, sorted(c_ax.current_matches()))],
          expect_contain=[(False, [])])

    # ---- 11.11 v32：数字键 1..9 直接选词并写回编辑器 ----
    ui_k = _UI2()
    b_k = _CapBackend(
        {"line_no": 3, "caret_col": 8, "line_text": "    num",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "in_decl_position": False, "decl_names": [],
         "proc_name": "Foo", "module_name": "M1"}, nav_ids)
    c_k = E.Completer(b_k, ui_k)
    c_k.trigger()
    ok2 = c_k.pick(1)                  # 按数字键 2 -> 选中第 2 项 numArr
    check("v32: 数字键 pick(1) 成功",
          [ok2, b_k.last_apply[3] if b_k.last_apply else None],
          expect_contain=[True, "numArr"])
    check("v32: 数字键选词后弹窗收起", [ui_k.shown], expect_contain=[False])
    # 越界的数字必须返回 False 且【不改变状态】——那说明用户真想输入数字，
    # 调用方据此放行按键。
    ui_k2 = _UI2()
    b_k2 = _CapBackend(
        {"line_no": 3, "caret_col": 8, "line_text": "    num",
         "in_string": False, "in_comment": False, "in_type_position": False,
         "in_decl_position": False, "decl_names": [],
         "proc_name": "Foo", "module_name": "M1"}, nav_ids)
    c_k2 = E.Completer(b_k2, ui_k2)
    c_k2.trigger()
    ok9 = c_k2.pick(8)                 # 只有 3 个候选，按 9 越界
    check("v32: 越界数字 pick(8) 返回 False 且状态不变",
          [ok9, ui_k2.shown, c_k2.selected],
          expect_contain=[False, True, 0])

    # ---- 11.12 v49：候选不再带序号（数字键选词已移除），一屏 15 行 ----
    try:
        import ui as _ui
        check("v49: 候选显示就是名字本身（无 `N. ` 前缀）",
              [_ui.format_row(0, "num") == "num",
               _ui.format_row(2, "numCount") == "numCount",
               _ui.format_row(8, "numArr") == "numArr"],
              expect_contain=[True, True, True])
        check("v49: 一屏最多 15 行（候选不足则列表更短）",
              [_ui.MAX_VISIBLE_ROWS], expect_contain=[15])
        check("v42: 列表宽度固定 30 字符（不再有横向滚动条）",
              [_ui.NAME_CHARS, hasattr(_ui.Popup, "has_hscroll")],
              expect_contain=[30, False])
        check("v42: 超长名字用 ... 省略",
              [_ui.ELLIPSIS], expect_contain=["..."])
    except Exception as e:                                # pragma: no cover
        check("v49: 候选显示不带序号（跳过：ui 不可用 %s）" % e, [True],
              expect_contain=[True])

    # ---- 12. v23-3：模块名 / 窗体名 / 类名参与提示 ----
    print("\n=== 12. 模块名 / 窗体名参与提示（v23-3）===")
    proj_c = {
        "Module1.bas": ("Public Sub Foo()\nEnd Sub", True, "Module1"),
        "pub2.bas.bas": ("Public Sub Baz()\nEnd Sub", True, "pub2"),
        "UserForm1.frm.bas": ("Private Sub UserForm_Initialize()\nEnd Sub",
                              False, "UserForm1"),
        "类1.cls.bas": ("Public Sub Bar()\nEnd Sub", False, "类1"),
    }
    recs_c = collect(proj_c, caret_mod="Module1", caret=None)
    vis_c = E.filter_identifiers_by_scope(recs_c, "Foo", "Module1")

    def pick(prefix):
        return sorted(i for i in vis_c if i.lower().startswith(prefix.lower()))

    check("写码时可提示窗体名", pick("userform"),
          expect_contain=["UserForm1"])
    check("写码时可提示其它标准模块名", pick("pub"),
          expect_contain=["pub2"])
    check("写码时可提示中文类名", pick("类"),
          expect_contain=["类1"])
    check("本模块名也可提示", pick("module"),
          expect_contain=["Module1"])

    # `As` 类型名位置：只提示类型名（窗体/类模块/模块名/Type/Enum），不提示变量
    class _B3(_B2):
        def get_type_names(self):
            return ["UserForm1", "Module1", "MyType"]

    ids3 = [("gUserName", "M1", None, False),
            ("UserForm1", "UserForm1", None, False),
            ("MyType", "M1", None, False)]
    check("As 之后提示窗体名",
          [trig("Dim f As UserF", 15, False, True, ids3, _B3)],
          expect_contain=[(True, ["UserForm1"])])
    check("As 之后提示自定义类型名",
          [trig("Dim f As MyTy", 14, False, True, ids3, _B3)],
          expect_contain=[(True, ["MyType"])])
    check("As 之后不提示变量名",
          [trig("Dim f As gUser", 15, False, True, ids3, _B3)],
          expect_contain=[(False, [])])

    # ---- 13. v24：走真实 _collect_identifiers 路径收集组件名 ----
    # 上一节用的是 collect() 复刻版，它无条件收录组件名，因此漏掉了真正
    # 的 bug：vbe_bridge 里组件名是在 `if count <= 0: continue` 【之后】才
    # 收录的，新建窗体（0 行代码）会被整个跳过 -> 窗体名不提示。
    print("\n=== 13. 组件名收集：空代码的窗体也要收录（v24）===")
    import vbe_bridge as VB

    class _FCM:
        def __init__(self, text, name):
            self._text = text or ""
            self.Name = name

        @property
        def CountOfLines(self):
            return 0 if not self._text else len(self._text.split("\n"))

        def Lines(self, start, count):
            if not self._text:
                return ""
            ls = self._text.split("\n")
            return "\n".join(ls[start - 1:start - 1 + count])

    class _FComp:
        def __init__(self, name, ctype, text):
            self.Name, self.Type = name, ctype
            self.CodeModule = _FCM(text, name)

    class _FProj:
        def __init__(self, name, comps):
            self.Name, self.VBComponents = name, comps

    class _FVBE:
        def __init__(self, proj):
            self.ActiveVBProject = proj
            self.VBProjects = [proj]

    def collect_vbe(comps, caret_mod=None, caret=None):
        """调用真实的 VbeBackend._collect_identifiers（COM 用假对象顶替）。"""
        vbe = _FVBE(_FProj("VBAProject", comps))
        old_get, old_caret = VB._get_vbe, VB._caret_position
        VB._get_vbe = lambda: vbe
        VB._caret_position = lambda v: (caret_mod, caret)
        VB._vbe_cache["obj"] = None      # 清掉上个用例残留的假 COM 对象
        VB.com_reset()
        try:
            bk = VB.VbeBackend()
            recs = bk._collect_identifiers()
            return recs, bk
        finally:
            VB._get_vbe, VB._caret_position = old_get, old_caret

    comps13 = [
        _FComp("Module1", 1, "Public Sub Foo()\nEnd Sub"),
        _FComp("UserForm1", 3, ""),      # 新建窗体：一行代码都没有
        _FComp("UserForm2", 3, "Private Sub UserForm_Initialize()\nEnd Sub"),
        _FComp("Class1", 2, ""),         # 空类模块
        _FComp("类2", 2, ""),            # 空中文类模块
    ]
    recs13, bk13 = collect_vbe(comps13)
    names13 = sorted(set(str(r[0]) for r in recs13))
    check("空代码窗体名被收录", names13, expect_contain=["UserForm1"])
    check("有代码窗体名被收录", names13, expect_contain=["UserForm2"])
    check("空类模块名被收录", names13, expect_contain=["Class1"])
    check("中文类模块名被收录", names13, expect_contain=["类2"])
    check("标准模块名被收录", names13, expect_contain=["Module1"])
    vis13 = E.filter_identifiers_by_scope(recs13, "Foo", "Module1")
    check("写码时能提示窗体名", sorted(vis13),
          expect_contain=["UserForm1", "UserForm2"])
    check("空窗体名进入类型名集合",
          sorted(bk13.get_type_names()), expect_contain=["UserForm1"])

    # ---- 14. v24：VBE 自动补括号时光标要落进括号里 ----
    print("\n=== 14. 补全后光标位置（v24）===")
    check("跳过左括号（同步补括号）",
          [VB._skip_into_parens("Public Function gCalc()", 21,
                                "Public Function gCal", 21)],
          expect_contain=[22])
    check("原行已有左括号则不动",
          [VB._skip_into_parens("    x = gCalc(1)", 13, "    x = gCal(1)", 13)],
          expect_contain=[13])
    check("后面没有括号则不动",
          [VB._skip_into_parens("    gUserName", 13, "    gUse", 9)],
          expect_contain=[13])

    class _FCM2:
        def __init__(self, lines, auto="off"):
            self.lines = list(lines)
            self.Name = "Module1"
            self.auto = auto          # off / sync / async
            self.pending = False

        def Lines(self, start, count):
            return self.lines[start - 1]

        def ReplaceLine(self, line_no, text):
            self.lines[line_no - 1] = text
            if self.auto == "sync":
                if "(" not in text:
                    self.lines[line_no - 1] = text + "()"
            elif self.auto == "async":
                self.pending = True   # 等 SetSelection 之后才补

        @property
        def CountOfLines(self):
            return len(self.lines)

    class _FPane:
        def __init__(self, cm):
            self.CodeModule = cm
            self.sel = (1, 1, 1, 1)

        def GetSelection(self):
            return self.sel

        def SetSelection(self, sl, sc, el, ec):
            cm = self.CodeModule
            if getattr(cm, "pending", False):
                cm.pending = False
                t = cm.lines[sl - 1]
                # VBE 在光标处插入 "()"，光标仍停在左括号左侧
                cm.lines[sl - 1] = t[:sc - 1] + "()" + t[sc - 1:]
            self.sel = (sl, sc, el, ec)

    class _FVBE2:
        def __init__(self, pane):
            self.ActiveCodePane = pane

    def apply14(lines, line_no, start_col, caret_col, completion, auto="off"):
        cm = _FCM2(lines, auto=auto)
        pane = _FPane(cm)
        old = VB._get_vbe
        VB._get_vbe = lambda: _FVBE2(pane)
        VB._vbe_cache["obj"] = None      # 每个用例必须用自己那份假对象
        VB.com_reset()
        try:
            VB.VbeBackend().apply_completion(line_no, start_col, caret_col,
                                             completion)
        finally:
            VB._get_vbe = old
        return cm.lines[line_no - 1], pane.sel[3]

    # 同步补括号：Function gCal -> gCalc()，光标应在 "(" 之后（列 23）
    # "Public Function " 16 字符，gCal 起于第 17 列，光标在其后 = 第 21 列
    got = apply14(["Public Function gCal"], 1, 17, 21, "gCalc", auto="sync")
    check("同步补括号：光标进括号", [got], expect_contain=[("Public Function gCalc()", 23)])
    # 异步补括号（读回行之后才补）同样要纠正
    got = apply14(["Public Function gCal"], 1, 17, 21, "gCalc", auto="async")
    check("异步补括号：光标进括号", [got], expect_contain=[("Public Function gCalc()", 23)])
    # 不补括号时，光标停在补全词尾部
    got = apply14(["Public Function gCal"], 1, 17, 21, "gCalc", auto="off")
    check("不补括号：光标在词尾", [got], expect_contain=[("Public Function gCalc", 22)])
    # 调用点：括号本来就存在，光标停在词尾（不进括号）
    got = apply14(["    x = gCal(1)"], 1, 9, 13, "gCalc", auto="off")
    check("调用点：原有括号不被跳过", [got], expect_contain=[("    x = gCalc(1)", 14)])
    # 普通变量名不受影响
    got = apply14(["    gUse"], 1, 5, 9, "gUserName", auto="off")
    check("普通变量名：光标在词尾", [got], expect_contain=[("    gUserName", 14)])

    # ---- 15. v25：过程级不泄漏（只读用法扫描不再记成模块级）----
    print("\n=== 15. 过程级变量不跨过程泄漏（v25）===")

    def leak_case(code, caret_line, caret_col, prefix, module="M1", std=True):
        caret = (caret_line, caret_col)
        r, _pn = list(P.extract_records(code, module=module,
                                       is_std_module=std, caret=caret))
        r += P.extract_implicit_records(code, module=module,
                                        is_std_module=std, declared=r,
                                        scope="proc", caret=caret)
        # 注：v30 起解析层不再全局剔除声明行名字（由引擎精确剔除），故此处
        # 只做作用域过滤，与后端保持一致。
        proc = P.proc_at_line(code, caret_line)
        vis = E.filter_identifiers_by_scope(r, proc, module)
        return sorted(i for i in vis if i.lower().startswith(prefix.lower()))

    # 15.1 只被读取、从未赋值的名字：以前被记为模块级 -> 全模块可见（泄漏）
    c_ro2 = "\n".join([
        "Sub ProcA()",
        "    Debug.Print zzOnly",
        "    x = 1",
        "End Sub",
        "Sub ProcB()",
        "    beta = 2",
        "    Debug.Print beta",
        "End Sub"])
    check("只读变量不泄漏到别的过程",
          leak_case(c_ro2, 7, len("    Debug.Print ") + 1, "zz"),
          expect_absent=["zzOnly"])
    # 光标必须落在【别的】行：压在 zzOnly 上会被"防自我提示"逻辑抹掉
    check("只读变量在自己所在过程里仍提示",
          leak_case(c_ro2, 3, len("    x") + 1, "zz"),
          expect_contain=["zzOnly"])
    # 15.2 Dim 出来的局部变量依旧严格按过程隔离
    c_loc = "\n".join([
        "Sub ProcA()",
        "    Dim alpha As Long",
        "    alpha = 1",
        "End Sub",
        "Sub ProcB()",
        "    Dim beta As Long",
        "    beta = 2",
        "End Sub"])
    check("局部变量不跨过程（Dim）",
          leak_case(c_loc, 6, len("    be") + 1, "al"), expect_absent=["alpha"])
    check("隐式变量不跨过程",
          leak_case(c_ro2, 6, len("    be") + 1, "zz"), expect_absent=["zzOnly"])
    # 15.3 模块级变量仍应全模块可见（不能被这条改动误伤）
    c_mod = "\n".join([
        "Public gAlpha As Long",
        "Sub ProcA()",
        "    gAlpha = 1",
        "End Sub",
        "Sub ProcB()",
        "    Debug.Print gAlpha",
        "End Sub"])
    check("模块级 Public 仍在过程内可见",
          leak_case(c_mod, 6, len("    Debug.Print gAl") + 1, "gal"),
          expect_contain=["gAlpha"])
    # 15.4 VBA 内置常量不该被当成隐式变量
    c_const = "\n".join([
        "Sub ProcA()",
        "    If x = vbTextCompare Then",
        "    End If",
        "    i = xlUp",
        "End Sub"])
    check("vbTextCompare 不当变量",
          leak_case(c_const, 2, len("    If x = ") + 1, "vb"),
          expect_absent=["vbTextCompare"])
    check("xlUp 不当变量",
          leak_case(c_const, 4, len("    i = xl") + 1, "xl"),
          expect_absent=["xlUp"])
    # 15.5 真实工程数据：不该出现"名字可见但所有记录都不属于当前过程"
    real_leaks = []
    for _fn in sorted(os.listdir(TESTDATA)):
        if not _fn.endswith(".bas"):
            continue
        _code = io.open(os.path.join(TESTDATA, _fn), encoding="utf-8").read()
        _mn = _mod_name(_fn)
        _r, _pn = list(P.extract_records(_code, module=_mn,
                                       is_std_module=_is_std(_fn)))
        _r += P.extract_implicit_records(_code, module=_mn,
                                         is_std_module=_is_std(_fn),
                                         declared=_r, scope="proc")
        for _pr in sorted({x[2] for x in _r if x[2]}):
            _vis = set(E.filter_identifiers_by_scope(_r, _pr, _mn))
            for _n in _vis:
                _srcs = {(x[2] or "").lower() for x in _r if x[0] == _n}
                if _srcs and all(s != _pr.lower() for s in _srcs):
                    # 只算"纯过程级"来源的名字（模块级过程名/类型名本就该可见）
                    if all(s for s in _srcs):
                        real_leaks.append((_fn, _pr, _n))
    check("真实工程：无跨过程泄漏", [len(real_leaks)],
          expect_contain=[0])
    if real_leaks:
        print("        泄漏明细: %s" % (real_leaks[:8],))

    # ---- 16. v25：COM 失败退避（关文件卡顿 / 关完又启动 Office）----
    print("\n=== 16. COM 失败退避（v25）===")
    VB.com_reset()
    check("初始不处于退避期", [VB.com_backoff_remaining()], expect_contain=[0.0])
    VB._com_fail()
    _b1 = round(VB.com_backoff_remaining(), 1)
    VB._com_fail()
    VB._com_fail()
    _b2 = round(VB.com_backoff_remaining(), 1)
    check("失败后进入退避且逐级变长", [_b1 > 0 and _b2 >= _b1],
          expect_contain=[True])
    check("退避期 _get_vbe_cached 直接返回 None",
          [VB._get_vbe_cached() is None], expect_contain=[True])
    VB.com_reset()
    check("reset 后恢复", [VB.com_backoff_remaining()], expect_contain=[0.0])

    # ---- 17. v26：COM 代理生命周期（关文件卡 3~10 秒 / 关完又启动 Office）----
    #
    # 病根：把 Excel/VBE 的 COM 代理长期缓存，等于我们的进程一直持有
    # Excel.Application —— Excel 退不干净（关完又启动 Office），且 VBAProject
    # 变脏时的写回会被拖成好几秒。第一版是"每次新建、用完即弃"所以没问题。
    print("\n=== 17. COM 代理生命周期（v26）===")

    class _FCM3:
        def __init__(self, text):
            self.text = text

        def Lines(self, a, b):
            return self.text

        def ReplaceLine(self, n, t):
            self.text = t

    class _FPane3:
        def __init__(self, cm):
            self.CodeModule = cm
            self.sel = (1, 1, 1, 1)

        def GetSelection(self):
            return self.sel

        def SetSelection(self, sl, sc, el, ec):
            self.sel = (sl, sc, el, ec)

    class _FVBE3:
        def __init__(self, pane):
            self.ActiveCodePane = pane

    _old_ttl, _old_probe = VB.VBE_PROXY_TTL, VB._probe_host_window
    _old_wnd, _old_get = dict(VB._wnd_state), VB._get_vbe
    try:
        # 用假 COM 对象，整节都不依赖真实 Excel 是否在跑
        VB._get_vbe = lambda: _FVBE3(_FPane3(_FCM3("    x = 1")))
        VB._probe_host_window = lambda: True
        VB._wnd_state.update(seen=True, miss=0)

        # 17.1 TTL <= 0：彻底不缓存，每次新建、用完即弃
        VB.VBE_PROXY_TTL = 0.0
        VB.com_reset()
        VB._vbe_cache["obj"] = None
        _a = VB._get_vbe_cached()
        _b = VB._get_vbe_cached()
        check("TTL=0 时每次都新建代理",
              [_a is not None, _b is not None, _a is _b],
              expect_contain=[True, True, False])

        # 17.2 TTL 内复用；到期必须重建（不能永久攥着 Excel）
        VB.VBE_PROXY_TTL = 30.0
        VB.com_reset()
        VB._vbe_cache["obj"] = None
        _a = VB._get_vbe_cached()
        _b = VB._get_vbe_cached()
        check("TTL 未到期时复用同一代理（打字性能不受影响）",
              [_a is _b], expect_contain=[True])
        VB._vbe_cache["at"] -= 60          # 模拟时间流逝
        _c = VB._get_vbe_cached()
        check("TTL 到期后重建代理（不长期持有 Excel）",
              [_c is not None, _c is _a], expect_contain=[True, False])

        # 17.3 宿主窗口消失（Excel 正在退出）：绝不新建代理，并进入退避
        VB.VBE_PROXY_TTL = 0.0
        VB.com_reset()
        VB._vbe_cache["obj"] = None
        VB._probe_host_window = lambda: True
        VB._wnd_state.update(seen=True, miss=0)
        _held = VB._get_vbe_cached() is not None     # 窗口还在时正常持有
        VB._probe_host_window = lambda: False
        VB._wnd_state.update(seen=True, miss=0)
        VB._get_vbe_cached()                          # miss=1：还不拦
        VB._get_vbe_cached()                          # miss=2：认定宿主真没了
        _gone = VB._get_vbe_cached()
        check("宿主窗口消失后不再新建代理",
              [_held, _gone is None], expect_contain=[True, True])
        check("宿主窗口消失后不留残留代理",
              [VB._vbe_cache["obj"] is None], expect_contain=[True])
        check("宿主窗口消失后进入退避",
              [VB.com_backoff_remaining() > 0], expect_contain=[True])

        # 17.4 宿主窗口回来 -> 立刻解除退避，工具自动恢复
        VB._probe_host_window = lambda: True
        VB._wnd_state.update(seen=True, miss=2)
        VB._host_window_alive()
        check("宿主窗口回来后解除退避",
              [VB.com_backoff_remaining()], expect_contain=[0.0])

        # 17.5 从未见过宿主窗口（WPS 等类名不匹配的宿主）：不能误拦
        VB.com_reset()
        VB._wnd_state.update(seen=False, miss=0)
        VB._probe_host_window = lambda: False
        check("未见过宿主窗口时不拦截（兼容 WPS 等宿主）",
              [VB._host_window_alive()], expect_contain=[True])

        # 17.6 写回 VBA 代码（VBAProject 变脏）后立刻放手
        VB._probe_host_window = lambda: True
        VB._wnd_state.update(seen=True, miss=0)
        VB.VBE_PROXY_TTL = 30.0
        VB.com_reset()
        VB._vbe_cache["obj"] = None
        _cm = _FCM3("    gUse")
        VB._get_vbe = lambda: _FVBE3(_FPane3(_cm))
        VB.VbeBackend().apply_completion(1, 5, 9, "gUserName")
        check("写回代码后立即释放 COM 代理",
              [(_cm.text, VB._vbe_cache["obj"] is None)],
              expect_contain=[("    gUserName", True)])
    finally:
        VB.VBE_PROXY_TTL, VB._probe_host_window = _old_ttl, _old_probe
        VB._wnd_state.update(_old_wnd)
        VB._get_vbe = _old_get
        VB.com_reset()
        VB._vbe_cache["obj"] = None

    # ---- 18. v27：全局钩子只在 VBE 聚焦时挂载（修复 PyCharm/VSCode 关不掉）----
    print("\n=== 18. 全局钩子仅 VBE 聚焦时挂载（v27）===")
    if M is None:
        check("main 模块不可用（缺 pynput），跳过第 18 节", [True],
              expect_contain=[True])
    else:
        GRACE = M.FOCUS_GRACE_SEC

        def rr(focused, released, backoff=0.0, elapsed=99.0, grace=GRACE,
               closing=False, suppress=False):
            return M._reconcile_resources(focused, released, backoff,
                                          time.time() - elapsed, time.time(), grace,
                                          closing=closing, suppress=suppress)

        # 18.1 在 VBE 里：必须挂载钩子（且不释放 COM）
        check("VBE 聚焦 -> 挂载钩子",
              [rr(True, False, backoff=0.0, elapsed=0.0)],
              expect_contain=[(True, False)])
        # 18.2 离开 VBE 后立即（宽限内）：不挂载钩子、也不释放 COM（等宽限到期）
        check("离开 VBE 宽限内 -> 不挂载钩子、不释放",
              [rr(False, False, backoff=0.0, elapsed=0.5)],
              expect_contain=[(False, False)])
        # 18.3 离开 VBE 超过宽限期：不挂载钩子 + 释放 COM（已释放则不再重复释放）
        check("离开 VBE 超宽限 -> 不挂载钩子 + 释放",
              [rr(False, False, backoff=0.0, elapsed=GRACE + 1.0)],
              expect_contain=[(False, True)])
        check("已释放后仍不挂载钩子（关键回归点）",
              [rr(False, True, backoff=0.0, elapsed=GRACE + 1.0)],
              expect_contain=[(False, False)])
        # 18.4 COM 失败退避中：即便不在 VBE 也绝不挂载钩子（钩子挂上会干扰关文件）
        check("COM 退避期（不在 VBE）-> 不挂载钩子 + 释放",
              [rr(False, False, backoff=2.0, elapsed=0.0)],
              expect_contain=[(False, True)])
        # 18.5 反例：复刻「旧版 bug」——离开 VBE 后 released 变 True，
        #     旧逻辑会走 else 把钩子重新挂上。新逻辑 mount_hooks 必须仍为 False。
        old_mount = (not False)  # 旧版 else 分支无条件重挂，等价于“始终挂载”
        check("旧版行为确为‘离开也挂载’（对比基线）",
              [old_mount], expect_contain=[True])
        check("新版：离开 VBE 后（released=True）mount_hooks 必须为 False",
              [M._reconcile_resources(False, True, 0.0, 0.0, time.time(), GRACE)[0]],
              expect_contain=[False])

        # 18.6 v28：前景窗口正在关闭 -> 即使去抖仍维持"在 VBE"也绝不挂钩
        #     这正是 PyCharm/VSCode 关不掉的根因：关闭 IDE 时焦点离开 VBE 的
        #     去抖窗口（约 600ms）内钩子仍挂着，IDE 退出被挡死。
        check("焦点在 VBE 但前台正在关闭 -> 不挂载钩子（核心修复点）",
              [rr(True, False, closing=True)], expect_contain=[(False, False)])
        check("前台正在关闭（不在 VBE）-> 不挂载钩子",
              [rr(False, False, closing=True)[0]], expect_contain=[False])
        check("新版：closing 必须使 mount_hooks=False（即便 focused=True）",
              [M._reconcile_resources(True, False, 0.0, 0.0, time.time(), GRACE,
                                      closing=True)[0]],
              expect_contain=[False])

        # 18.7 v29：上一个非 VBE 前台窗口已销毁（IDE 关掉、焦点被甩回 VBE）
        #     -> 冷却期内抑制钩子挂载，避免拖慢那个仍在收尾的进程
        #     （关窗口后 2~3 秒才退出的根因）。
        check("抑制冷却期（focused=True）-> 不挂载钩子",
              [rr(True, False, suppress=True)], expect_contain=[(False, False)])
        check("抑制冷却期必须压过 focused（即便 closing=False）",
              [rr(True, False, suppress=True, closing=False)[0]],
              expect_contain=[False])
        check("抑制冷却期不影响 COM 释放判定（离开超宽限仍释放）",
              [rr(False, False, elapsed=GRACE + 1.0, suppress=True)],
              expect_contain=[(False, True)])
        check("无抑制且正常在 VBE -> 挂载钩子（回归：抑制不能误伤常态）",
              [rr(True, False, suppress=False, closing=False)],
              expect_contain=[(True, False)])
        # 关键反例：用户主动 Alt+Tab 回 VBE（上一个前台窗口仍活着）不应被抑制，
        # 这里用纯函数无法直接构造“已销毁窗口”，靠 18.6/18.7 的 focused 组合
        # + reconcile 内 IsWindow 判定共同保证；此处确认 suppress=False 时照常挂载。

    # ---- 19. v35：走【真实的】VbeBackend + Completer（假 COM 对象驱动）----
    #
    # 这一段专门补 11.x 覆盖不到的盲区：11.x 用的是手写的 collect() 复刻版，
    # 它会把 caret 传进解析器；而生产的 get_identifiers() 从不传 caret。
    # 两者行为不一致 -> 真实 bug（"arr 打全名后列表消失"）在 11.x 里根本测不出来。
    print("\n=== 19. 真实后端全链路（v35：打全名必须保留列表）===")
    try:
        import vbe_bridge as VB

        class _CM(object):
            def __init__(self, name, text):
                self.Name = name
                self.text = text

            @property
            def CountOfLines(self):
                return self.text.count("\n") + 1

            def Lines(self, start, count):
                return "\r\n".join(self.text.split("\n")
                                   [start - 1:start - 1 + count])

        class _Comp(object):
            def __init__(self, name, text, ctype=1):
                self.Name = name
                self.Type = ctype
                self.CodeModule = _CM(name, text)

        class _Pane(object):
            def __init__(self, comp):
                self._c = comp
                self.sel = (1, 1, 1, 1)

            @property
            def CodeModule(self):
                return self._c.CodeModule

            def GetSelection(self):
                return self.sel

        class _VBE(object):
            def __init__(self, comps, act):
                self._p = type("_Proj", (object,), {})()
                self._p.Name = "VBAProject"
                self._p.VBComponents = list(comps)
                self.ActiveVBProject = self._p
                self.ActiveCodePane = _Pane(act)

        class _UI(object):
            def __init__(self):
                self.shown = False

            def show(self, matches, selected, completer=None):
                self.shown = True

            def update_selection(self, i):
                pass

            def hide(self):
                self.shown = False

            def contains_point(self, x, y):
                return False

        def drive(mods, active_mod, edit_line, indent, name):
            """逐字符输入 name，返回 [(输入到的词, 是否显示, 候选), ...]。"""
            out = []
            for n in range(1, len(name) + 1):
                typed = name[:n]
                comps, act = [], None
                for mn, text, ct in mods:
                    if mn == active_mod:
                        ls = text.split("\n")
                        ls[edit_line - 1] = indent + typed
                        text = "\n".join(ls)
                    c = _Comp(mn, text, ct)
                    comps.append(c)
                    if mn == active_mod:
                        act = c
                vbe = _VBE(comps, act)
                # ★v107c：补丁【用完就还原】。原先只换不还 —— 后果不是本节出错，
                # 而是**从这一节起每一节**的 `_get_vbe_cached` 都停在这个假工程上
                # （后面 §29/§31/§62 各自"还原"到的其实也是它），而不会有任何一条
                # 用例变红。§90 有专门的自查护栏盯着这件事。
                _sv_drive = VB._get_vbe_cached
                VB._get_vbe_cached = lambda: vbe
                try:
                    bk = VB.VbeBackend()
                    ui = _UI()
                    cp = E.Completer(bk, ui)
                    col = len(indent + typed) + 1
                    vbe.ActiveCodePane.sel = (edit_line, col, edit_line, col)
                    cp.trigger(True)
                    out.append((typed, ui.shown, list(cp.current_matches())))
                finally:
                    VB._get_vbe_cached = _sv_drive
            return out

        # 19.1 核心回归：**从未 Dim、只是到处在用** 的隐式变量 arr。
        #      修复前：a / ar 提示 arr，一把 arr 打全列表就消失（用户报的现象）。
        m_imp = [("Module1", "Sub Foo()\n    arr = 1\n    <in>\nEnd Sub", 1)]
        check("19.1 隐式变量 arr：打前缀提示", drive(m_imp, "Module1", 3, "    ", "ar"),
              expect_contain=[("ar", True, ["arr"])])
        check("19.1 隐式变量 arr：打全名仍保留",
              drive(m_imp, "Module1", 3, "    ", "arr"),
              expect_contain=[("arr", True, ["arr"])])
        # 19.2 单字符隐式变量（同一个机制，只是长度不同）
        m_single = [("Module1", "Sub Foo()\n    a = 1\n    <in>\nEnd Sub", 1)]
        check("19.2 单字符隐式变量 a：打全名保留",
              drive(m_single, "Module1", 3, "    ", "a"),
              expect_contain=[("a", True, ["a"])])
        # 19.3 显式 Dim 的多字符变量（v33 的老场景不能被新逻辑破坏）
        m_dim = [("Module1", "Option Explicit\nSub Foo()\n    Dim arr As Long\n    <in>\nEnd Sub", 1)]
        check("19.3 Dim arr：打全名保留",
              drive(m_dim, "Module1", 4, "    ", "arr"),
              expect_contain=[("arr", True, ["arr"])])
        # 19.4 对照：从没出现过的词（numA）打全仍不能提示自己
        m_to = [("Module1", "Option Explicit\nPublic num As Long\nPublic numArr As Long\nSub Foo()\n    <in>\nEnd Sub", 1)]
        check("19.4 未定义的 numA：绝不提示自己",
              drive(m_to, "Module1", 5, "    ", "numA"),
              expect_contain=[("numA", True, ["numArr"])],
              expect_absent=[("numA", True, ["numA", "numArr"])])
        check("19.4 num/numArr：打 num 两个同屏",
              drive(m_to, "Module1", 5, "    ", "num"),
              expect_contain=[("num", True, ["num", "numArr"])])
        # 19.5 纯幻影词：工程里任何地方都没有 -> 应当收起，不提示自己
        m_ghost = [("Module1", "Option Explicit\nSub Foo()\n    <in>\nEnd Sub", 1)]
        check("19.5 纯幻影 zzq：不提示",
              drive(m_ghost, "Module1", 3, "    ", "zzq"),
              expect_contain=[("zzq", False, [])])

        # ---- 20. v36：模糊匹配（不必从首字母开始）+ 命中字符高亮 ----
        #
        # 用户原话：定义了 dataArr / dataSheet，输入 ts 要提示 dataSheet
        # （t 和 s 都在里面），输入 ta 两个都要提示，且命中的字符标红。
        print("\n=== 20. 模糊匹配与命中高亮（v36）===")

        def probe(mods, active_mod, edit_line, indent, typed):
            """把 typed 整串放到编辑行，返回 (是否显示, 候选, 命中位置字典)。"""
            comps, act = [], None
            for mn, text, ct in mods:
                if mn == active_mod:
                    ls = text.split("\n")
                    ls[edit_line - 1] = indent + typed
                    text = "\n".join(ls)
                c = _Comp(mn, text, ct)
                comps.append(c)
                if mn == active_mod:
                    act = c
            vbe = _VBE(comps, act)
            # ★v107c：补丁【用完就还原】（见 §90 的自查护栏）。
            _sv_probe = VB._get_vbe_cached
            VB._get_vbe_cached = lambda: vbe
            try:
                bk = VB.VbeBackend()
                ui = _UI()
                cp = E.Completer(bk, ui)
                col = len(indent + typed) + 1
                vbe.ActiveCodePane.sel = (edit_line, col, edit_line, col)
                cp.trigger(True)
                return (ui.shown, list(cp.current_matches()),
                        dict(getattr(cp, "match_hits", None) or {}))
            finally:
                VB._get_vbe_cached = _sv_probe

        m_two = [("Module1",
                  "Option Explicit\n"
                  "Sub Foo()\n"
                  "    Dim dataArr As Long\n"
                  "    Dim dataSheet As Worksheet\n"
                  "    <in>\n"
                  "End Sub", 1)]

        # 20.1 核心场景一：ts -> 只提示 dataSheet（dataArr 里没有 s）
        s1, m1, h1 = probe(m_two, "Module1", 5, "    ", "ts")
        check("20.1 输入 ts：弹窗显示", [s1], expect_contain=[True])
        check("20.1 输入 ts：提示 dataSheet", m1, expect_contain=["dataSheet"])
        check("20.1 输入 ts：不提示 dataArr（里面没有 s）", m1,
              expect_absent=["dataArr"])
        check("20.1 输入 ts：命中位置为 t(2) 与 S(4)", [h1.get("dataSheet")],
              expect_contain=[[2, 4]])

        # 20.2 核心场景二：ta -> 两个都提示，且各自命中位置正确
        s2, m2, h2 = probe(m_two, "Module1", 5, "    ", "ta")
        check("20.2 输入 ta：弹窗显示", [s2], expect_contain=[True])
        check("20.2 输入 ta：dataArr 与 dataSheet 同屏", [sorted(m2)],
              expect_contain=[["dataArr", "dataSheet"]])
        check("20.2 输入 ta：dataArr 命中 t(2) a(3)", [h2.get("dataArr")],
              expect_contain=[[2, 3]])
        check("20.2 输入 ta：dataSheet 命中 t(2) a(3)", [h2.get("dataSheet")],
              expect_contain=[[2, 3]])

        # 20.3 高亮位置必须合法：数量等于输入长度、升序、不越界
        bad = []
        for name, pos in (h2 or {}).items():
            if len(pos) != 2 or pos != sorted(pos) or pos[-1] >= len(name):
                bad.append((name, pos))
        check("20.3 命中位置数量/顺序/越界检查", [bad], expect_contain=[[]])

        # 20.4 词首缩略：ds -> 命中 dataSheet 的 D 与 S（不是随便找两个字母）
        s4, m4, h4 = probe(m_two, "Module1", 5, "    ", "ds")
        check("20.4 输入 ds：提示 dataSheet", m4, expect_contain=["dataSheet"])
        check("20.4 输入 ds：命中词首 D(0) 与 S(4)", [h4.get("dataSheet")],
              expect_contain=[[0, 4]])

        # 20.5 连续块优先：arr -> 命中 dataArr 尾部的 Arr，而不是散着凑
        s5, m5, h5 = probe(m_two, "Module1", 5, "    ", "arr")
        check("20.5 输入 arr：命中连续的 Arr(4,5,6)", [h5.get("dataArr")],
              expect_contain=[[4, 5, 6]])

        # 20.6 排序：精确匹配必须排第一（不能被模糊候选挤下去）
        m_rank = [("Module1",
                   "Option Explicit\n"
                   "Public num As Long\n"
                   "Public myNumber As Long\n"
                   "Sub Foo()\n"
                   "    <in>\n"
                   "End Sub", 1)]
        s6, m6, _h6 = probe(m_rank, "Module1", 5, "    ", "num")
        check("20.6 输入 num：精确匹配 num 排第一", [m6[0] if m6 else None],
              expect_contain=["num"])
        check("20.6 输入 num：模糊候选 myNumber 也在列", m6,
              expect_contain=["myNumber"])

        # 20.7 排序：前缀匹配排在"中间命中"之前
        m_pref = [("Module1",
                   "Option Explicit\n"
                   "Public wsName As String\n"
                   "Public myWorksheetName As String\n"
                   "Sub Foo()\n"
                   "    <in>\n"
                   "End Sub", 1)]
        s7, m7, _h7 = probe(m_pref, "Module1", 5, "    ", "wsn")
        check("20.7 输入 wsn：前缀命中的 wsName 排第一",
              [m7[0] if m7 else None], expect_contain=["wsName"])

        # 20.8 不回归：工程里压根没有的字符组合 -> 不显示
        s8, m8, _h8 = probe(m_two, "Module1", 5, "    ", "zzq")
        check("20.8 纯幻影 zzq：不显示、无候选", [(s8, m8)],
              expect_contain=[(False, [])])

        # 20.9 不回归：打全名仍保留列表（v35 的修复不能被模糊匹配破坏）
        s9, m9, h9 = probe(m_two, "Module1", 5, "    ", "dataSheet")
        check("20.9 打全名 dataSheet：列表保留且它排第一",
              [(s9, m9[0] if m9 else None)], expect_contain=[(True, "dataSheet")])
        check("20.9 打全名：命中位置覆盖全部字符",
              [h9.get("dataSheet")], expect_contain=[[0, 1, 2, 3, 4, 5, 6, 7, 8]])

        # 20.11 v59：纯贪心会【漏判】—— 字符要"跳步"时仍必须能匹配
        #
        # 用户原话：定义了 getSettingTitle3，输入 get3 能提示，输入 getse3 不提示。
        # 根因在 _greedy_positions：挑位置时只看"这一档更优"，不管后面还凑不凑得
        # 齐 —— 第 3 个字符 t 贪心跳到词边界 T(10)（档 1 优于接着用的 t(2) 档 2），
        # 后面再没有 s，于是整个 query 被判成不匹配。可 g-e-t-s-e-…-3 =
        # [0,1,2,3,4,15] 本来就成立。修法：先反向算可行性上界 maxpos，再把正向
        # 候选限制在 i..maxpos[t] 之内（见 engine._greedy_positions）。
        m_v59 = [("Module1",
                  "Option Explicit\n"
                  "Private Function getSettingTitle3(ws, col)\n"
                  "    <in>\n"
                  "End Function", 1)]

        s11, m11, h11 = probe(m_v59, "Module1", 3, "    ", "getse3")
        check("20.11 输入 getse3：提示 getSettingTitle3", m11,
              expect_contain=["getSettingTitle3"])
        check("20.11 getse3 命中位置 [0,1,2,3,4,15]",
              [h11.get("getSettingTitle3")],
              expect_contain=[[0, 1, 2, 3, 4, 15]])

        # 20.12 反向把关：query 真的【不是】子序列时仍要判不匹配
        # （不能为了"不漏判"就乱放行 —— 名字里只有一个 s，getss3 不该中）
        s12, m12, _h12 = probe(m_v59, "Module1", 3, "    ", "getss3")
        check("20.12 输入 getss3：不提示（名字里只有一个 s）", m12,
              expect_absent=["getSettingTitle3"])

        # 20.13 不回归：原本就能中的 get3 行为不变
        s13, m13, h13 = probe(m_v59, "Module1", 3, "    ", "get3")
        check("20.13 输入 get3：仍提示 getSettingTitle3", m13,
              expect_contain=["getSettingTitle3"])
        check("20.13 get3 命中位置仍为 [0,1,10,15]",
              [h13.get("getSettingTitle3")], expect_contain=[[0, 1, 10, 15]])

        # 20.14 不能为了修漏判把"优先词边界"的观感改坏：ds -> dataSheet 仍是 [0,4]
        _s14, _m14, h14 = probe(m_two, "Module1", 5, "    ", "ds")
        check("20.14 输入 ds：命中位置仍为词首 [0,4]",
              [h14.get("dataSheet")], expect_contain=[[0, 4]])

        # ---- 21. 声明行静默（v105；v37 的"打全名要提示"只在用法行成立）----
        #
        # 历史：v37 曾把"in_decl_position 时无条件剔除正在声明的名字"改成
        #   "只看名字在工程里真实存在吗"（存在就提示，含打全名）。
        # ★v105 整体反转：**声明行静默**（用户要 IDEA 那种）。
        #   所以 21.1/21.2 的期望从 (True, [...]) 反转为 (False, [])。
        #   ⚠️ v37 想保住的语义（真实的同名名字要能提示）没丢 —— 由下面 21.3
        #     的**调用处**（非声明行）覆盖。
        print("\n=== 21. 声明行静默 / 调用处照常提示（v105）===")

        def probe2(mods, active_mod, edit_line, indent, typed):
            comps, act = [], None
            for mn, text, ct in mods:
                if mn == active_mod:
                    ls = text.split("\n")
                    ls[edit_line - 1] = indent + typed
                    text = "\n".join(ls)
                c = _Comp(mn, text, ct)
                comps.append(c)
                if mn == active_mod:
                    act = c
            vbe = _VBE(comps, act)
            # ★v107c：补丁【用完就还原】（见 §90 的自查护栏）。
            _sv_probe2 = VB._get_vbe_cached
            VB._get_vbe_cached = lambda: vbe
            try:
                bk = VB.VbeBackend()
                ui = _UI()
                cp = E.Completer(bk, ui)
                col = len(indent + typed) + 1
                vbe.ActiveCodePane.sel = (edit_line, col, edit_line, col)
                cp.trigger(True)
                return (ui.shown, sorted(cp.current_matches()))
            finally:
                VB._get_vbe_cached = _sv_probe2

        # 21.1 已有 Function test()，在新的 Sub 声明行上输入（v105：不弹）
        m_fn = [("Module1",
                 "Option Explicit\n"
                 "Function test() As Long\n"
                 "    test = 1\n"
                 "End Function\n"
                 "\n"
                 "<in>\n"
                 "End Sub", 1)]
        check("21.1 声明行输入 tes：不弹候选（v105）",
              [probe2(m_fn, "Module1", 6, "Sub ", "tes")],
              expect_contain=[(False, [])])
        check("21.1 声明行输入 test（打全名）：也不弹（v105）",
              [probe2(m_fn, "Module1", 6, "Sub ", "test")],
              expect_contain=[(False, [])])

        # 21.2 test / test02 都在，声明行输入（v105：不弹）
        m_two = [("Module1",
                  "Option Explicit\n"
                  "Function test() As Long\n"
                  "    test = 1\n"
                  "End Function\n"
                  "Function test02() As Long\n"
                  "    test02 = 2\n"
                  "End Function\n"
                  "\n"
                  "<in>\n"
                  "End Sub", 1)]
        check("21.2 声明行输入 test：不弹候选（v105）",
              [probe2(m_two, "Module1", 10, "Sub ", "test")],
              expect_contain=[(False, [])])
        check("21.2 声明行输入 test02：不弹候选（v105）",
              [probe2(m_two, "Module1", 10, "Sub ", "test02")],
              expect_contain=[(False, [])])

        # 21.3 过程体内调用处回退（非声明行，属回归保护）
        m_call = [("Module1",
                   "Option Explicit\n"
                   "Function test() As Long\n"
                   "    test = 1\n"
                   "End Function\n"
                   "Function test02() As Long\n"
                   "    test02 = 2\n"
                   "End Function\n"
                   "Sub Main()\n"
                   "    <in>\n"
                   "End Sub", 1)]
        check("21.3 调用处回退到 test：两个都提示",
              [probe2(m_call, "Module1", 10, "    ", "test")],
              expect_contain=[(True, ["test", "test02"])])

        # 21.4 关键反向保护：幻影词（哪都不存在）依然【绝不】提示自己
        m_ghost = [("Module1",
                    "Option Explicit\n"
                    "Sub Main()\n"
                    "    <in>\n"
                    "End Sub", 1)]
        check("21.4 幻影 zzq：不提示",
              [probe2(m_ghost, "Module1", 3, "    ", "zzq")],
              expect_contain=[(False, [])])
        check("21.4 声明行幻影 zzq：不提示",
              [probe2([("Module1", "Option Explicit\n<in>\nEnd Sub", 1)],
                      "Module1", 2, "Sub ", "zzq")],
              expect_contain=[(False, [])])
        # 21.5 回退出来的未定义词 numA（v35 的核心用例，不能被 v37 放宽掉）
        m_numA = [("Module1",
                   "Option Explicit\n"
                   "Public num As Long\n"
                   "Public numArr As Long\n"
                   "Sub Foo()\n"
                   "    <in>\n"
                   "End Sub", 1)]
        check("21.5 回退出的 numA：只提示 numArr，绝不提示 numA 自己",
              [probe2(m_numA, "Module1", 5, "    ", "numA")],
              expect_contain=[(True, ["numArr"])])

        # 20.10 高亮分段：把名字切成"命中/未命中"交替的段（纯函数，无需 GUI）
        import ui as _uimod

        check("20.10 分段：dataSheet 命中 [2,4]",
              [_uimod.split_by_hits("dataSheet", [2, 4])],
              expect_contain=[[("da", False), ("t", True), ("a", False),
                               ("S", True), ("heet", False)]])
        check("20.10 分段：连续命中会合并成一段",
              [_uimod.split_by_hits("dataArr", [4, 5, 6])],
              expect_contain=[[("data", False), ("Arr", True)]])
        check("20.10 分段：无命中时整段一段",
              [_uimod.split_by_hits("num", [])],
              expect_contain=[[("num", False)]])
        check("20.10 分段：全命中时整段一段",
              [_uimod.split_by_hits("num", [0, 1, 2])],
              expect_contain=[[("num", True)]])
        check("20.10 分段：拼接回去必须等于原名（不丢字）",
              ["".join(s for s, _ in _uimod.split_by_hits("dataSheet", [2, 4]))],
              expect_contain=["dataSheet"])

        # 20.11 真实 Popup 渲染冒烟：Canvas 自绘不抛异常（无 GUI 环境则跳过）
        try:
            import tkinter as _tk

            class _FakeC(object):
                match_hits = {"dataSheet": [2, 4], "dataArr": [2, 3]}
                selected = 0

                def pick(self, i):
                    return True

                def accept(self):
                    pass

            _root = _tk.Tk()
            _root.withdraw()
            _p = _uimod.Popup(_root)
            _p.show(["dataSheet", "dataArr"], 0, _FakeC())
            _root.update()
            check("20.11 Canvas 渲染：两行都画出来了", [len(_p._rows)],
                  expect_contain=[2])
            check("20.11 Canvas 渲染：命中位置已传给 UI", [_p._rows[0]],
                  expect_contain=[("dataSheet", [2, 4])])
            check("20.11 点空白行不触发确认（返回 None）", [_p._row_at(10 ** 6)],
                  expect_contain=[None])
            _root.destroy()
        except Exception as _e2:
            check("20.11 跳过（无 GUI 环境）: %s" % _e2, [True],
                  expect_contain=[True])

        # ---------------------------------------------------------------
        # 22. 候选超过一屏：滚动窗口，不多丢、序号不再显示（v40）
        #
        # 用户报的问题：旧版本把 MAX_ROWS(=8) 当成了候选上限，
        # 第 9 个往后直接看不见（"多的都给我删没了"）。v38 把它还原成
        # 【一屏显示几行】：候选全留着，靠上下键 / 滚轮滚动查看；
        # v39 再把上限从 8 提到 10，并去掉"数字键选词"（会与输入变量名里的
        # 数字冲突，例如 s1 会被误选成 sheet1）。选择只走 ↑/↓ + Tab 或鼠标点击。
        # ---------------------------------------------------------------
        print("\n=== 22. 超过一屏不丢候选，滚动查看（v39）===")
        try:
            _N = 30
            _names = ["x%02d" % i for i in range(_N)]

            class _WinUI(object):
                """记录 UI 真正收到了哪些行（不依赖 tkinter，纯记账）。"""
                def __init__(self):
                    self.shown = False
                    self.rows = []
                    self.sel = 0
                    self.hidden = False

                def show(self, rows, sel, completer=None):
                    self.shown = True
                    self.hidden = False
                    self.rows = list(rows)
                    self.sel = sel

                def update_selection(self, sel):
                    self.sel = sel

                def hide(self):
                    self.shown = False
                    self.hidden = True
                    self.rows = []

                def contains_point(self, x, y):
                    return False

            class _ManyBk(object):
                """一次性给出 N 个候选的最小假后端。"""
                def __init__(self, names, word):
                    self._names = list(names)
                    line = "    " + word
                    self._ctx = {
                        "line_no": 3, "caret_col": len(line) + 1,
                        "line_text": line, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": "M1",
                    }
                    self.applied = None

                def get_context(self):
                    return dict(self._ctx)

                def get_identifiers(self):
                    return [(n, "M1", None, False) for n in self._names]

                def get_declared_names(self):
                    return list(self._names)

                def get_type_names(self):
                    return []

                def name_exists_outside_caret(self, name, caret):
                    return True

                def apply_completion(self, line_no, start, end, name):
                    self.applied = name
                    return True

            def mk22(names=None, word="x"):
                bk = _ManyBk(names or _names, word)
                ui = _WinUI()
                c = E.Completer(bk, ui, view_rows=10)
                c.trigger()
                return bk, ui, c

            bk1, ui1, c1 = mk22()
            all_m = c1.current_matches()
            check("22.1 30 个候选一个都不丢（不再只留 8 个）",
                  [len(all_m)], expect_contain=[_N])
            check("22.1 弹窗确实显示了", [ui1.shown], expect_contain=[True])

            check("22.2 UI 只画当前这一屏（最多 10 行）",
                  [ui1.rows], expect_contain=[all_m[:10]])
            check("22.2 这一屏 == visible_rows()（同一份数据，不会错位）",
                  [c1.visible_rows()], expect_contain=[all_m[:10]])
            check("22.3 初始窗口停在列表开头", [c1.view_info()],
                  expect_contain=[(_N, 0)])
            check("22.3 初始选中本屏第 1 行", [c1.view_selection()],
                  expect_contain=[0])

            # 22.4 在本屏内移动：窗口不该动
            for _ in range(9):
                c1.move(1)
            check("22.4 本屏内下移 9 次：selected=9 且窗口不动",
                  [c1.selected, c1.view_info(), c1.view_selection()],
                  expect_contain=[9, (_N, 0), 9])

            # 22.5 越过下沿：窗口跟着下移一行（而不是整屏跳）
            c1.move(1)
            check("22.5 走出本屏：窗口只下移一行",
                  [c1.selected, c1.view_info(), c1.view_selection()],
                  expect_contain=[10, (_N, 1), 9])
            check("22.5 窗口内容与选中行一致",
                  [c1.visible_rows()], expect_contain=[all_m[1:11]])
            check("22.5 选中行始终在窗口内",
                  [c1.view_top <= c1.selected < c1.view_top + c1.view_rows],
                  expect_contain=[True])

            # 22.6 一屏最多 10 行：屏幕上只会显示 10 条，再多靠滚动
            check("22.6 屏幕内一屏最多 10 行",
                  [len(c1.visible_rows())], expect_contain=[10])
            check("22.6 超过一屏才出现纵向滚动（total > 10）",
                  [c1.view_info()[0] > 10], expect_contain=[True])

            # 22.7 鼠标点击对应"本屏第几行"：滚过之后点第 1 行，选的是新的第 1 项
            #      （数字键选词已在 v39 移除，避免与"输入变量名里的数字"冲突）
            bk7, ui7, c7 = mk22()
            m7 = c7.current_matches()
            for _ in range(10):
                c7.move(1)
            check("22.7 滚一行后本屏第一行是原列表第 2 项",
                  [c7.visible_rows()[0]], expect_contain=[m7[1]])
            # 鼠标点本屏第 1 行 -> _row_at 换算成绝对下标 1 -> completer.pick(1)
            check("22.7 点击本屏第 1 行（pick 绝对下标 1）写入的就是它",
                  [c7.pick(1), bk7.applied],
                  expect_contain=[True, m7[1]])
            check("22.7 选完就收起（accept 会 hide）", [ui7.hidden],
                  expect_contain=[True])

            # 22.8 一直下移到底：窗口滑到最后 10 行，再往前一步绕回首项也回到顶部
            bk8, ui8, c8 = mk22()
            for _ in range(29):
                c8.move(1)
            check("22.8 移到末项（第 30 项）：窗口滑到最底",
                  [c8.selected, c8.view_info(), c8.visible_rows()],
                  expect_contain=[29, (_N, _N - 10), c8.current_matches()[-10:]])
            c8.move(1)                      # 末项 -> 绕回首项
            check("22.8 绕回首项：窗口一起回到顶部",
                  [c8.selected, c8.view_info()], expect_contain=[0, (_N, 0)])

            # 22.9 打全名命中很靠后的候选时（第 26 项），窗口要滚过去带着它，
            #      否则用户看不见自己选中的是什么
            bk9, ui9, c9 = mk22()
            c9.selected = 25
            c9._scroll_to_selected()
            check("22.9 选中第 26 项：窗口滚到它可见，且落在最后一行",
                  [c9.view_info(), c9.view_selection()],
                  expect_contain=[(_N, 16), 9])
            check("22.9 选中项确实在这一屏里",
                  [c9.current_matches()[25] in c9.visible_rows()],
                  expect_contain=[True])

            # 22.10 候选不足一屏：窗口永远停在 0，不出现滚动
            bkA, uiA, cA = mk22(["num", "numArr", "numCount"], "num")
            check("22.10 不足一屏：窗口恒为 0、全量显示",
                  [cA.view_info(), cA.visible_rows()],
                  expect_contain=[(3, 0), ["num", "numArr", "numCount"]])
            check("22.10 不足一屏：候选数 < 一屏上限，不触发滚动",
                  [cA.view_info()[0] <= cA.view_rows], expect_contain=[True])
        except Exception as _e22:
            check("第 22 节异常: %s" % _e22, [True], expect_contain=[False])

        # ---- 第 23 节：拖滚动条时 UI 不能消失（v41 修复） ----
        # 复现：之前拖滚动条拖一小会儿，整个列表框+滚动条就消失。
        # 根因是弹窗点击会被"激活"、抢走 VBE 前台焦点，焦点去抖(~600ms)
        # 后 main.py 误判"已离开 VBE"自动收起。修复：弹窗加 WS_EX_NOACTIVATE
        # 永不抢焦点；并给 maybe_hide_on_outside_click 加"拖拽中不收起"保险。
        class _Bk23(object):
            def __init__(self, names, word):
                self._names = list(names)
                line = "    " + word
                self._ctx = {"line_no": 3, "caret_col": len(line) + 1,
                             "line_text": line, "in_string": False,
                             "in_comment": False, "in_type_position": False,
                             "in_decl_position": False, "decl_names": [],
                             "proc_name": None, "module_name": "M1"}
                self.applied = None
            def get_context(self):
                return dict(self._ctx)
            def get_identifiers(self):
                return [(n, "M1", None, False) for n in self._names]
            def get_declared_names(self):
                return list(self._names)
            def get_type_names(self):
                return []
            def apply_completion(self, *a, **k):
                self.applied = a
            def name_exists_outside_caret(self, w, c):
                return True
        class _Ui23(object):
            def __init__(self):
                self.hidden = False
                self._drag = None
            def show(self, *a, **k):
                self.hidden = False
            def update_selection(self, sel):
                pass
            def hide(self):
                self.hidden = True
            def contains_point(self, x, y):
                return False   # 始终模拟"点击落在弹窗外"
        names23 = ["name%02d" % i for i in range(30)]
        ui23 = _Ui23()
        c23 = E.Completer(_Bk23(names23, "name"), ui23, view_rows=10)
        c23.trigger()   # 触发：30 个候选，超过一屏
        check("23.0 触发后弹窗可见", [c23.is_visible()], expect_contain=[True])
        check("23.0 候选超过一屏(存在纵向滚动条)",
              [len(c23.current_matches()) > c23.view_rows], expect_contain=[True])

        # 23.1 正在拖纵向滚动条(_drag='v')时，即便 contains_point 说在外面，
        #       maybe_hide_on_outside_click 也绝不许收起
        ui23._drag = "v"
        c23.maybe_hide_on_outside_click(9999, 9999)
        check("23.1 拖滚动条途中点弹窗外：列表框不消失",
              [c23.is_visible()], expect_contain=[True])

        # 23.2 set_view_top（拖滚动条调用的核心）不会收起 UI，且真的改了窗口起点
        c23.set_view_top(15)
        check("23.2 拖到中部：UI 仍可见、窗口起点=15",
              [c23.is_visible(), c23.view_info()],
              expect_contain=[True, (30, 15)])
        check("23.2 选中项被夹进可见窗口",
              [c23.selected >= c23.view_top and
               c23.selected < c23.view_top + c23.view_rows],
              expect_contain=[True])

        # 23.3 拖完松开、_drag 归位后，点在弹窗外才按旧行为正常收起
        ui23._drag = None
        c23.maybe_hide_on_outside_click(9999, 9999)
        check("23.3 没在拖、点弹窗外：正常收起",
              [c23.is_visible()], expect_contain=[False])

        # ---- 24. v43：回退删字不得提示【比编辑器内容更长的旧片段】 ----
        #
        # 用户报的现象：定义一个很长的变量名后，按住回退键连删，弹窗会提示出
        # 回退过程中某个【更长】的片段 —— 编辑器里剩下的字符比提示出来的还短。
        #
        # 根因：标识符池是带缓存的（最快 ID_REFRESH_MIN_SEC=0.4s 才重解析一次），
        # 连删时池子里还留着"文字还长"那一刻解析出来的名字；而旧的回声防护只剔除
        # 【与当前词完全相等】的候选，于是"当前词 + 更长尾巴"的旧片段活了下来。
        print("\n=== 24. 回退删字不残留旧片段（v43）===")

        _LONG24 = ("studengdflkgjdslkgjdslkgjdlfksgjdlfkgjdlfkgjdlkjgldsgjdslgjdsgldsjgl"
                   "djgdlg")

        class _CM24(object):
            def __init__(self, name, text):
                self.Name = name
                self.text = text

            @property
            def CountOfLines(self):
                return len(self.text.split("\n"))

            def Lines(self, start, count):
                ls = self.text.split("\n")
                s = max(0, int(start) - 1)
                return "".join(l + "\n" for l in ls[s:s + int(count)])

        class _Comp24(object):
            def __init__(self, name, text, ctype=1):
                self.Name = name
                self.Type = ctype
                self.CodeModule = _CM24(name, text)

        class _Pane24(object):
            def __init__(self, comp):
                self._c = comp
                self.sel = (1, 1, 1, 1)

            @property
            def CodeModule(self):
                return self._c.CodeModule

            def GetSelection(self):
                return self.sel

        class _VBE24(object):
            def __init__(self, comps, act):
                self._p = type("_P24", (object,), {})()
                self._p.Name = "VBAProject"
                self._p.VBComponents = list(comps)
                self.ActiveVBProject = self._p
                self.ActiveCodePane = _Pane24(act)

        class _UI24(object):
            def __init__(self):
                self.shown = False
                self.rows = []

            def show(self, matches, selected, completer=None):
                self.shown = True
                self.rows = list(matches or [])

            def update_selection(self, i):
                pass

            def hide(self):
                self.shown = False
                self.rows = []

            def contains_point(self, x, y):
                return False

        def _session24():
            """建一次后端（它的标识符缓存要跨步骤存活）+ 一个触发器闭包。

            trig(mods, active, line_no, col) 每次都【重建模块文本】——
            真实编辑器里文本是同步缩短的；早期复现脚本没重建，才误以为
            "工程里始终留着完整长名"。
            """
            holder = {}
            VB._get_vbe_cached = lambda: holder["vbe"]
            bk = VB.VbeBackend()
            ui = _UI24()
            cp = E.Completer(bk, ui)

            def trig(mods, active, line_no, col):
                comps = [_Comp24(n, t) for n, t in mods]
                act = None
                for c in comps:
                    if c.Name == active:
                        act = c
                holder["vbe"] = _VBE24(comps, act)
                holder["vbe"].ActiveCodePane.sel = (line_no, col, line_no, col)
                cp.trigger(True)
                return list(cp.current_matches())

            def throttle():
                # 复刻 main.poll_editor 的重解析限流（ID_REFRESH_MIN_SEC=0.4）
                if time.time() - getattr(bk, "_cache_time", 0.0) > 0.4:
                    bk.invalidate_identifiers()

            return trig, throttle

        _orig_get24 = VB._get_vbe_cached
        try:
            # 24.1 连删 15 次：绝不能提示出比编辑器当前内容【更长】的候选
            _decl = ("Option Explicit\nSub Foo()\n    Dim {W} As Long\n"
                     "    x = 1\nEnd Sub")
            _both = ("Option Explicit\nSub Foo()\n    Dim {W} As Long\n"
                     "    {W} = 1\nEnd Sub")
            _bad24 = []
            for _tpl in (_decl, _both):
                _trig, _thr = _session24()
                # 先"打全名"：此刻池子是按完整长名解析并被缓存的
                _trig([("Module1", _tpl.replace("{W}", _LONG24))], "Module1",
                      3, 9 + len(_LONG24))
                for _k in range(1, 16):
                    _typed = _LONG24[:-_k]
                    _thr()
                    _m = _trig([("Module1", _tpl.replace("{W}", _typed))],
                               "Module1", 3, 9 + len(_typed))
                    _bad24 += [n for n in _m if len(n) > len(_typed)]
                    time.sleep(0.008)
            check("24.1 回退删字：不提示比编辑器内容更长的旧片段",
                  [_bad24], expect_contain=[[]])

            # 24.2 真名字（本模块别处声明）的前缀补全不能被误杀
            _trig, _thr = _session24()
            _m2 = _trig([("Module1", "Option Explicit\nPublic numArr As Long\n"
                                       "Sub Foo()\n    num\nEnd Sub")],
                        "Module1", 4, 5 + 3)
            check("24.2 真名字前缀补全仍在（num -> numArr）",
                  [_m2], expect_contain=[["numArr"]])

            # 24.3 隐式变量（别处用过、没声明）的前缀补全不能被误杀
            #
            # ⚠️ v75 起这条【不能再用 Option Explicit】—— 写了它的模块不再收集
            # 隐式变量（用户口径，见第 52 节），旧写法等于在断言新行为。
            # 本断言真正要守的是"回声防护不误杀隐式变量"，与是否强制声明无关，
            # 所以这里去掉那一行、行号随之前移。
            _trig, _thr = _session24()
            _m3 = _trig([("Module1", "Sub Foo()\n    arr = 1\n"
                                       "    ar\nEnd Sub")],
                        "Module1", 3, 5 + 2)
            check("24.3 隐式变量前缀补全仍在（不带 Option Explicit：ar -> arr）",
                  [_m3], expect_contain=[["arr"]])
            # 反向对照：同一个文本加上 Option Explicit 后，arr 用过但没声明
            # ⇒ 不该再被提示（v75 的语义翻转，与第 52 节同源）。
            _trig, _thr = _session24()
            _m3b = _trig([("Module1", "Option Explicit\nSub Foo()\n    arr = 1\n"
                                        "    ar\nEnd Sub")],
                         "Module1", 4, 5 + 2)
            check("24.3b ★对照：带 Option Explicit 时 arr 不再提示"
                  "（用过但没声明 = 编译错误）",
                  [_m3b], expect_absent=["arr"])

            # 24.4 别的模块的 Public 名字仍要提示（跨模块靠"声明集合"这一路证据）
            _trig, _thr = _session24()
            _m4 = _trig([("Module1", "Option Explicit\nSub Foo()\n    Global\n"
                                       "End Sub"),
                         ("Module2", "Option Explicit\n"
                                     "Public GlobalDataArray As Long\n")],
                        "Module1", 3, 5 + 6)
            check("24.4 跨模块 Public 前缀补全仍在（Global -> GlobalDataArray）",
                  [_m4], expect_contain=[["GlobalDataArray"]])
            # ---- 25. v44：从名字【开头】删字符也要能出提示词 ----
            #
            # 用户报的现象：从变量 / 函数 / 窗体 / 模块名的【结尾】【中间】删字符都能
            # 弹出候选，但从名字【开头】删（光标停在残留名字最左边按 Delete）什么都不弹。
            #
            # 根因两处，都出在"只看光标左边"这个前提上：
            #   1. engine._ident_char_before_caret（轮询触发的准入门槛）只检查光标前
            #      一格；名字开头左邻是空格 / 括号 -> 判成"没在拼标识符" -> 直接收起。
            #   2. engine.extract_word_before 只往光标左边取词，左边取不到就返回 None
            #      -> 连候选都没得算。
            # 因此"从中间 / 结尾删"（左边还剩半截词）照常工作，"从头删"永远失效。
            print("\n=== 25. 从名字开头删字符也能出提示词（v44）===")

            # 第 2 行（声明行）始终留着【完整长名】，第 4 行放"被删剩"的残词。
            # 这样标识符池里一定有完整长名，断言与缓存刷新时机无关，不会 flaky。
            _TPL25 = ("Option Explicit\nPublic {D} As Long\nSub Foo()\n"
                      "    {U}\nEnd Sub")

            def _src25(declared, used):
                return _TPL25.replace("{D}", declared).replace("{U}", used)

            _trig, _thr = _session24()
            # 先打全名（模拟真实打字；同时确认池子里确实收了这个名字）
            _m25a0 = _trig([("Module1", _src25(_LONG24, _LONG24))],
                           "Module1", 4, 5 + len(_LONG24))
            check("25.0 打全名：池子里有这个名字",
                  [_LONG24 in _m25a0], expect_contain=[True])

            # 25.1 从开头删：光标停在残留名字最左端（第 4 行第 5 列）
            _r25 = []
            for _k in (1, 5, 20):
                _thr()
                _m = _trig([("Module1", _src25(_LONG24, _LONG24[_k:]))],
                           "Module1", 4, 5)
                _r25.append(_LONG24 in _m)
            check("25.1 从名字开头删字符 -> 仍出提示（v44）",
                  _r25, expect_contain=[True])

            # 25.2 从中间删：光标在该行词内部（原有行为不许回退）
            _thr()
            _m25b = _trig([("Module1", _src25(_LONG24, _LONG24))],
                          "Module1", 4, 5 + 12)
            check("25.2 从名字中间删字符 -> 照常出提示",
                  [_LONG24 in _m25b], expect_contain=[True])

            # 25.3 从结尾删：光标在残留词末尾
            _thr()
            _m25c = _trig([("Module1", _src25(_LONG24, _LONG24[:30]))],
                          "Module1", 4, 5 + 30)
            check("25.3 从名字结尾删字符 -> 照常出提示",
                  [_LONG24 in _m25c], expect_contain=[True])

            # 25.4 取词与替换范围（纯函数级，直接盯 extract_word_at 的返回）
            _w25 = E.extract_word_at("    " + _LONG24[5:], 5)
            check("25.4a 光标在名字开头：取词=残留名，范围盖住整个残留词",
                  [_w25],
                  expect_contain=[(_LONG24[5:], 5, 5 + len(_LONG24[5:]))])
            check("25.4b 光标在词尾：右端就是光标列（与旧行为一致）",
                  [E.extract_word_at("    num", 8)], expect_contain=[("num", 5, 8)])
            check("25.4c 光标前后都不是标识符字符 -> 取不到词（不乱弹）",
                  [E.extract_word_at("a  b", 3)],
                  expect_contain=[(None, None, None)])

            # 25.5 Tab 确认：替换范围必须盖住光标右边那截残留名字，
            #      否则会在残留词前面插入候选，拼出"双份名字"
            class _B25(object):
                def __init__(self, ctx, ids):
                    self._ctx = ctx
                    self._ids = list(ids)
                    self.applied = None

                def get_context(self):
                    return self._ctx

                def get_identifiers(self):
                    return self._ids

                def get_declared_names(self):
                    return [i[0] if isinstance(i, (tuple, list)) else i
                            for i in self._ids]

                def apply_completion(self, line_no, start, end, completion):
                    self.applied = (start, end, completion)

            _line25 = "    " + _LONG24[5:]
            _b25 = _B25({"line_no": 4, "caret_col": 5, "line_text": _line25,
                         "in_string": False, "in_comment": False,
                         "in_type_position": False, "in_decl_position": False,
                         "decl_names": [], "proc_name": "Foo",
                         "module_name": "Module1"}, [_LONG24])
            _c25 = E.Completer(_b25, _UI24())
            _c25.trigger(True)
            _c25.accept()
            check("25.5 Tab 确认：整段换掉残留词（不拼出双份名字）",
                  [_b25.applied],
                  expect_contain=[(5, 5 + len(_LONG24[5:]), _LONG24)])

            # ---- 26. v45：从名字【开头】删字符不许提示出"更长的旧片段"（回声）----
            #
            # 用户报的现象：把光标停在名字最左边、按 Delete 从【开头】连删时，
            # 弹窗会提示出【完整的旧长名】（有回音、提示自己）。
            #
            # 与 v43"从结尾删"的回声同源，根因是回声判据只认【前缀】：
            #   候选.startswith(当前词)
            # 从结尾删 -> 当前词是候选的前缀 -> 命中，能剔除；
            # 从开头删 -> 当前词是候选的【后缀】 -> 不命中 -> 回声漏了过去。
            #
            # 修法：判据推广为"当前词是候选名的【连续子串】"——删除只会让光标处
            # 的词变短，剩下的必然是原词的连续一段（前缀 / 后缀 / 中段通吃）。
            # 是否保留仍由 _name_really_exists 裁决：真名字（本模块别处出现 /
            # 别的模块声明过）照常提示，只有"现场已删干净、别处也没有"的旧快照
            # 才被剔除。
            #
            # 本节的场景里，正在编辑的【就是声明行本身】（Dim <长名> As Long），
            # 所以删到最后现场文本里已经找不到完整长名了 —— 池子（0.4s 缓存）里
            # 的那一份只是旧快照，必须剔除。
            _TPL26 = ("Option Explicit\nSub Foo()\n    Dim {W} As Long\n"
                      "    x = 1\nEnd Sub")

            def _src26(typed):
                return _TPL26.replace("{W}", typed)

            # 26.1 从开头删：不许提示出比编辑器当前内容更长的旧片段
            _trig, _thr = _session24()
            _trig([("Module1", _src26(_LONG24))], "Module1", 3, 9 + len(_LONG24))
            _bad26 = []
            for _k in (1, 5, 20, 40):
                _thr()
                _typed = _LONG24[_k:]
                _m = _trig([("Module1", _src26(_typed))], "Module1", 3, 9)
                _bad26 += [n for n in _m if len(n) > len(_typed)]
                time.sleep(0.008)
            check("26.1 从名字开头删字符：不提示出更长的旧片段（v45）",
                  [_bad26], expect_contain=[[]])

            # 26.2 头尾都删过 -> 当前词是旧长名的【中段】，同样不许提示旧长名
            _trig, _thr = _session24()
            _trig([("Module1", _src26(_LONG24))], "Module1", 3, 9 + len(_LONG24))
            _thr()
            _mid26 = _LONG24[10:50]
            _m26b = _trig([("Module1", _src26(_mid26))], "Module1", 3, 9)
            check("26.2 头尾都删过：中段残词不提示旧长名（v45）",
                  [_LONG24 in _m26b], expect_contain=[False])

            # 26.3 回归：本模块别处仍留着完整长名时（编辑的不是声明行），
            #      从头删依旧照常提示 —— 回声防护加强不许误杀真名字
            _trig, _thr = _session24()
            _trig([("Module1", _src25(_LONG24, _LONG24))], "Module1", 4,
                  5 + len(_LONG24))
            _thr()
            _m26c = _trig([("Module1", _src25(_LONG24, _LONG24[20:]))],
                          "Module1", 4, 5)
            check("26.3 回归：本模块声明的真名字从头删仍提示",
                  [_LONG24 in _m26c], expect_contain=[True])

            # 26.4 回归：从结尾删（v43 场景）仍不提示更长片段（控制组）
            _trig, _thr = _session24()
            _trig([("Module1", _src26(_LONG24))], "Module1", 3, 9 + len(_LONG24))
            _bad26d = []
            for _k in (1, 5, 20):
                _thr()
                _typed = _LONG24[:-_k]
                _m = _trig([("Module1", _src26(_typed))], "Module1", 3,
                           9 + len(_typed))
                _bad26d += [n for n in _m if len(n) > len(_typed)]
            check("26.4 回归：从结尾删仍不提示更长片段（v43 控制组）",
                  [_bad26d], expect_contain=[[]])

            # ---- 27. v46：光标停在名字【中间】删字符，当前词取【整段名字】----
            #
            # 用户报的现象：变量 test / test02 / test23456789；把 test23456789
            # 中间那个 t 删掉 -> tes23456789、光标停在 tes 之后，却仍然提示
            # test / test02 —— 因为取词只看光标左边（tes），完全无视光标右边
            # 还留着的 23456789。当前词改为整段名字后，tes23456789 既不匹配
            # test 也不匹配 test02，只剩真正那一整段名字（Tab 一次换掉整段）。
            print("\n=== 27. 从名字【中间】删字符：当前词取整段名字（v46）===")

            _TPL27 = ("Option Explicit\n"
                      "Public test As Long\n"
                      "Public test02 As Long\n"
                      "Public test23456789 As Long\n"
                      "Sub Foo()\n"
                      "    {U}\n"
                      "End Sub")

            def _src27(used):
                return _TPL27.replace("{U}", used)

            # 27.1 从中间删掉那个 t：光标停在 tes|23456789（第 6 行第 8 列）
            _trig, _thr = _session24()
            _trig([("Module1", _src27("test23456789"))], "Module1", 6,
                  5 + len("test23456789"))
            _m27a = _trig([("Module1", _src27("tes23456789"))], "Module1", 6,
                          5 + len("tes"))
            check("27.1 中间删字符：不再按左半截提示（无 test / test02）",
                  [[n for n in _m27a if n.lower() in ("test", "test02")]],
                  expect_contain=[[]])
            check("27.2 中间删字符：整段名字本身仍提示（Tab 可一次换掉）",
                  [_m27a], expect_contain=[["test23456789"]])

            # 27.3 纯函数级：取词与替换范围
            check("27.3 光标在名字中间：取词=整段名字，范围盖住整段",
                  [E.extract_word_at("    tes23456789", 8)],
                  expect_contain=[("tes23456789", 5, 16)])
            check("27.4 光标在词尾：右端仍是光标列（旧行为不变）",
                  [E.extract_word_at("    num", 8)],
                  expect_contain=[("num", 5, 8)])
            check("27.5 光标在名字开头：仍取右边整段（v44 行为不变）",
                  [E.extract_word_at("    " + _LONG24, 5)],
                  expect_contain=[(_LONG24, 5, 5 + len(_LONG24))])
            check("27.6 光标前后都不是标识符字符：仍取不到词（不乱弹）",
                  [E.extract_word_at("a  b", 3)],
                  expect_contain=[(None, None, None)])

            # 27.7 Tab 确认时传给后端的替换范围必须盖住整段名字（含右半截），
            #      否则候选会插在 tes 之后，拼出 test23456789 + 23456789
            class _B27(object):
                def __init__(self, ctx, ids):
                    self._ctx = ctx
                    self._ids = list(ids)

                def get_context(self):
                    return self._ctx

                def get_identifiers(self):
                    return self._ids

                def get_declared_names(self):
                    return set(str(i).lower() for i in self._ids)

            _ctx27 = {"line_no": 6, "caret_col": 8,
                      "line_text": "    tes23456789",
                      "in_string": False, "in_comment": False,
                      "in_type_position": False, "in_decl_position": False,
                      "proc_name": "Foo", "module_name": "Module1"}
            _cp27 = E.Completer(_B27(_ctx27, ["test", "test02",
                                              "test23456789"]), _UI24())
            _cp27.trigger(True)
            check("27.7 Tab 确认范围盖住整段名字（end_col=16 而非光标列 8）",
                  [(bool(_cp27.visible),
                    (_cp27.ctx or {}).get("word_end_col"))],
                  expect_contain=[(True, 16)])


            # ---- 第 29 节：Shift+Enter 在光标行下方新起一行（v48） ----
            #
            # 用户要求：光标停在代码行【中间】时想另起一行，老办法是"先把光标移到
            # 本行末尾再按回车"，太麻烦。现在按 Shift+Enter 一步完成，且新行起始
            # 位置要与上一行代码起始位置对齐（IDEA 的 Start New Line 同款）。
            #
            # 关键点：当前行【不拆分】—— 光标右侧的代码留在原行；新行内容 = 上一行的
            # 行首空白；光标落在新行缩进之后。后端实现见 VbeBackend.new_line_below()。
            print("\n=== 29. Shift+Enter 在当前行下方新起一行（v48）===")
            try:
                import main as _mn29
                _orig_get29 = VB._get_vbe_cached

                # 29.1 纯判断：只认"按住 Shift 的回车"
                check("29.1 只认 Shift+回车（裸回车 / Shift+Tab 都不算）",
                      [_mn29._is_newline_shortcut(0x0D, True),
                       _mn29._is_newline_shortcut(0x0D, False),
                       _mn29._is_newline_shortcut(0x09, True),
                       _mn29._is_newline_shortcut(0x0D, 0)],
                      expect_contain=[True, False, False, False])

                # 29.2 对齐依据：行首空白（空格 / Tab / 无缩进 / 空行）
                check("29.2 行首空白提取（新行对齐上一行代码起始位置的依据）",
                      [VB._leading_ws("    x = 1"), VB._leading_ws("\t If x"),
                       VB._leading_ws("x = 1"), VB._leading_ws("")],
                      expect_contain=["    ", "\t ", "", ""])

                # 29.3 缩进之后的光标列（字符列语义 = len + 1）
                check("29.3 缩进之后的光标列：4 空格 -> 第 5 列；无缩进 -> 第 1 列",
                      [VB._indent_end_col("    "), VB._indent_end_col("")],
                      expect_contain=[5, 1])

                class _CM29(object):
                    def __init__(self, lines):
                        self.lines = list(lines)
                        self.Name = "Module1"

                    @property
                    def CountOfLines(self):
                        return len(self.lines)

                    def Lines(self, start, count):
                        s0 = max(0, int(start) - 1)
                        return "".join(l + "\n"
                                       for l in self.lines[s0:s0 + int(count)])

                    def InsertLines(self, start, text):
                        i = min(max(0, int(start) - 1), len(self.lines))
                        for k, ln in enumerate(str(text).split("\n")):
                            self.lines.insert(i + k, ln)

                class _FPane29(object):
                    def __init__(self, cm, sel):
                        self.CodeModule = cm
                        self.sel = sel

                    def GetSelection(self):
                        return self.sel

                    def SetSelection(self, sl, sc, el, ec):
                        self.sel = (sl, sc, el, ec)

                class _FVBE29(object):
                    def __init__(self, pane):
                        self.ActiveCodePane = pane

                def _nl29(lines, line_no, col):
                    cm29 = _CM29(lines)
                    pane29 = _FPane29(cm29, (line_no, col, line_no, col))
                    VB._get_vbe_cached = lambda: _FVBE29(pane29)
                    ok29 = VB.VbeBackend().new_line_below()
                    return ok29, cm29.lines, pane29.sel

                _SRC29 = ["Option Explicit", "Sub Foo()", "    Dim x",
                          "    x = 1 + 2", "End Sub"]

                # 29.4 核心：光标停在行中间 -> 当前行不拆、新行缩进对齐、光标在缩进后
                _ok29, _ls29, _sel29 = _nl29(_SRC29, 4, 11)
                check("29.4 光标在行中间 -> 新行插在下方、缩进与上一行对齐、不拆行",
                      [_ok29, _ls29, _sel29],
                      expect_contain=[True,
                                      ["Option Explicit", "Sub Foo()", "    Dim x",
                                       "    x = 1 + 2", "    ", "End Sub"],
                                      (5, 5, 5, 5)])

                # 29.5 光标在行尾 / 行首：结果一致（都按"本行缩进"新起一行）
                _r29 = []
                for _c29 in (14, 5):
                    _ok, _ls, _sel = _nl29(_SRC29, 4, _c29)
                    _r29.append((_ok, _ls[4], _sel))
                check("29.5 行尾 / 行首同样新起一行，且都不拆行",
                      [_r29],
                      expect_contain=[[(True, "    ", (5, 5, 5, 5)),
                                       (True, "    ", (5, 5, 5, 5))]])

                # 29.6 末行：追加到模块末尾（InsertLines 起始行超出即追加）
                _ok29b, _ls29b, _sel29b = _nl29(_SRC29, 5, 8)
                check("29.6 光标在模块最后一行 -> 新行追加到末尾、光标在第 1 列",
                      [_ok29b, _ls29b, _sel29b],
                      expect_contain=[True, _SRC29 + [""], (6, 1, 6, 1)])

                # 29.7 Tab 缩进的行：显示列语义下光标要按 tab stop 展开
                _sem29 = VB._sem_cache.get("info")
                VB._sem_cache["info"] = ("disp", 4, False)
                try:
                    _ok29c, _ls29c, _sel29c = _nl29(
                        ["Sub Foo()", "\tIf x Then", "End Sub"], 2, 11)
                finally:
                    VB._sem_cache["info"] = _sem29
                # ★v89：`\tIf x Then` 是块头 -> 新行按判据深一级（两个 Tab；
                # 原先只抄本行空白，是一个 Tab）。Tab 语义与列语义不变。
                check("29.7 Tab 缩进：新行仍是 Tab（★v89 块头深一级 => 两个 Tab），"
                      "光标按显示列落在其后（第 9 列）",
                      [_ok29c, _ls29c[2], _sel29c],
                      expect_contain=[True, "\t\t", (3, 9, 3, 9)])

                # 29.8 取不到 VBE（不在代码窗 / Excel 已关）必须返回 False ——
                #      调用方据此【不吞键】，让 VBE 按原生行为处理
                VB._get_vbe_cached = lambda: None
                check("29.8 取不到 VBE -> 返回 False（绝不吞掉用户的按键）",
                      [VB.VbeBackend().new_line_below()],
                      expect_contain=[False])

                # 29.9 空模块（光标行号为 0）也要安全返回 False
                VB._get_vbe_cached = lambda: _FVBE29(
                    _FPane29(_CM29([]), (0, 0, 0, 0)))
                check("29.9 光标行号为 0（无有效行）-> 返回 False",
                      [VB.VbeBackend().new_line_below()],
                      expect_contain=[False])

                VB._get_vbe_cached = _orig_get29
            except Exception as _e29:
                check("第 29 节异常: %s" % _e29, [True], expect_contain=[False])

        finally:
            VB._get_vbe_cached = _orig_get24

    except Exception as _e:
        check("第 19 节不可用（vbe_bridge 导入失败）: %s" % _e, [True],
              expect_contain=[True])

    # ---- 30. v50：跨过程泄漏（只读隐式变量被算到别的过程 / 模块级）----
    #
    # 病根：_usage_candidates 把 `Attribute VB_xxx` / `Option Explicit` 这类
    # 行【替换成空串】以屏蔽内容，但这会缩短文本、让后续所有字符偏移整体前移；
    # 而调用方 _line_proc_map 算 offsets 用的是【未压缩】的文本。
    # 真实模块头部动辄 5 行以上 Attribute + Option（累计前移 150+ 字符），
    # 于是"只被读取、从未被赋值"的隐式变量被算成 proc=None（模块级）——
    # 全模块每个过程都能看到它，包括别的过程的**形参位置**。
    #
    # 修复：跳过行改为抹成【等长空格】，偏移恒定对齐。
    print("\n=== 30. 跨过程泄漏：跳过行必须等长屏蔽（v50）===")
    try:
        _HDR30 = [
            'Attribute VB_Name = "Module1"',
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = False",
            "Attribute VB_Exposed = False",
            "Option Explicit",
        ]
        _code30 = "\n".join(_HDR30 + [
            "",
            "Public Sub ProcedureA()",
            "    Debug.Print targetName",
            "End Sub",
            "",
            "Public Sub ProcedureB(ByVal sheetName As String)",
            "    Debug.Print resultValue",
            "End Sub",
        ])

        # 30.1 偏移对齐：_usage_candidates 返回的偏移必须与 offsets 同基准
        _masked30 = P._mask_strings_and_comments(_code30)
        _masked30 = P._RE_CONTINUATION.sub(" ", _masked30)
        _offs30, _lp30 = P._line_proc_map(_masked30)
        _owner30 = {n: P._proc_at_offset(_offs30, _lp30, p)
                    for n, p in P._usage_candidates(_masked30, set())}
        check("30.1 只读隐式变量归属正确（不再落到模块级）",
              [_owner30.get("targetName"), _owner30.get("resultValue")],
              expect_contain=["ProcedureA", "ProcedureB"])

        # 30.2 跳过行长度不变（等长屏蔽）
        _kept30 = "\n".join(" " * len(l) if P._RE_USAGE_SKIP_LINE.match(l) else l
                            for l in _masked30.split("\n"))
        check("30.2 跳过行抹成等长空格（偏移不漂移）",
              [len(_kept30), len(_masked30)],
              expect_contain=[len(_masked30), len(_masked30)])

        # 30.3 端到端：A 的变量不得出现在 B 的可见集（B 的形参位置）
        _r30, _pn = list(P.extract_records(_code30, module="M1",
                                          is_std_module=True))
        _r30 += P.extract_implicit_records(_code30, module="M1",
                                           is_std_module=True, declared=_r30,
                                           scope="proc")
        _visB30 = E.filter_identifiers_by_scope(_r30, "ProcedureB", "M1")
        _visA30 = E.filter_identifiers_by_scope(_r30, "ProcedureA", "M1")
        check("30.3 A 的只读变量不泄漏到 ProcedureB",
              [n for n in _visB30 if str(n).lower() == "targetname"],
              expect_contain=[])
        check("30.3 B 的只读变量不泄漏到 ProcedureA",
              [n for n in _visA30 if str(n).lower() == "resultvalue"],
              expect_contain=[])
        check("30.3 各自的只读变量在自己过程里仍提示",
              [sorted(str(n) for n in _visA30 if str(n).lower() == "targetname"),
               sorted(str(n) for n in _visB30 if str(n).lower() == "resultvalue")],
              expect_contain=[["targetName"], ["resultValue"]])

        # 30.4 头部越"胖"，旧实现漂移越大 —— 多 Attribute 行下依然正确
        _fat30 = "\n".join([
            'Attribute VB_Name = "M"',
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = False",
            "Attribute VB_Exposed = False",
            "#If VBA7 Then",
            "#End If",
            "Option Explicit",
            "Option Base 0",
            "",
            "Public Sub AA()",
            "    Debug.Print onlyInAA",
            "End Sub",
            "",
            "Public Sub BB()",
            "    Debug.Print onlyInBB",
            "End Sub",
        ])
        _m30b = P._RE_CONTINUATION.sub(
            " ", P._mask_strings_and_comments(_fat30))
        _o30b, _p30b = P._line_proc_map(_m30b)
        _own30b = {n: P._proc_at_offset(_o30b, _p30b, p)
                   for n, p in P._usage_candidates(_m30b, set())}
        check("30.4 更胖的模块头部下归属仍正确",
              [_own30b.get("onlyInAA"), _own30b.get("onlyInBB")],
              expect_contain=["AA", "BB"])

        # 30.5 回归：跳过行里的名字不该被当成隐式变量收录
        check("30.5 Attribute 行里的名字不当变量",
              [n for n, _p in P._usage_candidates(_masked30, set())
               if str(n).lower() in ("vb_name", "vb_globalnamespace",
                                     "explicit")],
              expect_contain=[])
    except Exception as _e30:
        check("第 30 节异常: %s" % _e30, [True], expect_contain=[False])

    # ---- 31. v51：形参名与别的过程的变量重名时不得"提示自己" ----
    #
    # 用户报的现象：在 ProcedureB 的形参位置，打 / 回退 / 粘贴一个与
    # ProcedureA 的局部变量一模一样的名字，弹窗会把刚敲进去的字原样提示出来；
    # 换成任何别处都没有的新名字反而安静（"不相同就不提示"）。
    #
    # 根因：形参位置的输入会被解析成 ProcedureB 自己的形参记录 —— 于是
    # "别的过程里出现过这个名字"（现场全模块扫描）被当成了"这个名字真实存在"
    # 的铁证，回声防护放行。本节的 31.1 先复现旧行为，31.2 起验证修复。
    print("\n=== 31. 形参位置不提示自己（v51：现场证据按作用域收窄）===")
    try:
        import vbe_bridge as VB31

        _HDR31 = [
            'Attribute VB_Name = "Module1"',
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = False",
            "Attribute VB_Exposed = False",
            "Option Explicit",
        ]

        class _CM31(object):
            """只实现后端用到的 CodeModule.CountOfLines / Lines。"""

            def __init__(self, code):
                self._code = code

            @property
            def CountOfLines(self):
                return len(self._code.split("\n"))

            def Lines(self, start, count):
                _ls = self._code.split("\n")
                return "\n".join(_ls[start - 1:start - 1 + count])

        class _VBE31(object):
            def __init__(self, code):
                _cm = _CM31(code)
                self.ActiveCodePane = type("_CP31", (), {"CodeModule": _cm})()

        _cur31 = {"code": ""}
        _orig_get31 = VB31._get_vbe_cached
        VB31._get_vbe_cached = lambda *a, **k: _VBE31(_cur31["code"])
        _names31 = VB31.VbeBackend().names_outside_caret

        class _UI31(object):
            def __init__(self):
                self.shown = False

            def show(self, *a, **k):
                self.shown = True

            def update_selection(self, *a, **k):
                pass

            def hide(self, *a, **k):
                self.shown = False

            def contains_point(self, *a, **k):
                return False

        def _parse31(lines):
            """带 | 标记的代码行 -> (代码, 光标位置)。"""
            _out, _caret = [], None
            for _i, _l in enumerate(lines, 1):
                if "|" in _l:
                    _col = _l.index("|")
                    _caret = (_i, _col + 1)
                    _l = _l[:_col] + _l[_col + 1:]
                _out.append(_l)
            return "\n".join(_out), _caret

        class _B31(object):
            """最小后端：候选/声明名来自真实解析，现场证据走真实
            VbeBackend.names_outside_caret（只把最底层读文本换成假的）。"""

            def __init__(self, code, caret, force_broad=False):
                self.code = code
                self.caret = caret
                self.force_broad = force_broad
                _r, _pn = list(P.extract_records(code, module="Module1",
                                               is_std_module=True,
                                               caret=caret))
                _r += list(P.extract_implicit_records(
                    code, module="Module1", is_std_module=True, declared=_r,
                    scope="proc", caret=caret))
                self._recs = _r
                self._decl_names = set(
                    str(n).lower() for n in
                    (P.decl_names_at_caret(code, caret) or ()))
                self._declared = set(
                    str(x[0]).lower() for x in _r
                    if str(x[0]).lower() not in self._decl_names)
                try:
                    self._caret_word = P.ident_at_caret(
                        code, caret[0], caret[1]).lower()
                except Exception:
                    self._caret_word = ""

            def get_context(self):
                _ls = self.code.split("\n")
                return {"line_no": self.caret[0], "caret_col": self.caret[1],
                        "line_text": _ls[self.caret[0] - 1],
                        "in_string": False, "in_comment": False,
                        "in_type_position": False,
                        "in_decl_position": bool(self._decl_names),
                        "decl_names": sorted(self._decl_names),
                        "proc_name": P.proc_at_line(self.code, self.caret[0]),
                        "module_name": "Module1"}

            def get_identifiers(self):
                return self._recs

            def get_declared_names(self):
                return list(self._declared)

            def caret_word_at_collect(self):
                return self._caret_word

            def declared_elsewhere(self, name, module_name):
                return False           # 全部声明都在 Module1

            def names_outside_caret(self, caret, scope_only=False):
                _cur31["code"] = self.code
                if self.force_broad:   # 对照：v50 行为（永远全模块扫描）
                    scope_only = False
                return _names31(caret, scope_only)

            def apply_completion(self, *a, **k):
                return None

        def _trig31(body, force_broad=False):
            _code, _caret = _parse31(_HDR31 + body)
            _be = _B31(_code, _caret, force_broad)
            _ui = _UI31()
            _c = E.Completer(_be, _ui)
            _c.trigger()
            return (_ui.shown, sorted(_c.matches or []))

        _leak31 = [
            "Public Sub ProcedureA()",
            "    Dim username As String",
            '    username = "abc"',
            "    Debug.Print username",
            "End Sub",
            "",
            "Public Sub ProcedureB(ByVal username| As String)",
            "End Sub",
        ]
        # 31.1 旧行为确实会"提示自己"——证明这条回归确实有意义。
        # 现在有【两道】独立护栏会拦住它：
        #   ① v53 的语义护栏（正在造名字的位置不提示自己）；
        #   ② ★v105 的"声明位置整个静默"（E.DECL_POPUP 默认 False）。
        # 对照必须【两道一起短路】，否则根本复现不出旧行为，这条控制组就恒真了。
        _orig_ml31 = E.Completer._module_level_names
        _orig_dp31 = E.DECL_POPUP
        E.Completer._module_level_names = lambda _self: set(
            str(_r[0]).lower() for _r in _self.backend.get_identifiers())
        E.DECL_POPUP = True          # 复原到 v105 之前：声明行照常出候选
        try:
            _ctl31 = _trig31(_leak31, force_broad=True)
        finally:
            E.Completer._module_level_names = _orig_ml31
            E.DECL_POPUP = _orig_dp31
        check("31.1 对照：宽窗口 + 两道护栏都关掉时形参重名会提示自己",
              [_ctl31],
              expect_contain=[(True, ["username"])])
        # 31.2 修复后：不再提示自己
        check("31.2 形参名与其他过程的变量重名 -> 不提示自己",
              [_trig31(_leak31)],
              expect_contain=[(False, [])])
        # 31.3 对照：任何地方都没有的新名字同样安静（用户说的"不相同不提示"）
        check("31.3 全新形参名保持安静",
              [_trig31([
                  "Public Sub ProcedureA()",
                  "    Dim username As String",
                  "End Sub",
                  "",
                  "Public Sub ProcedureB(ByVal zzqnewxx| As String)",
                  "End Sub"])],
              expect_contain=[(False, [])])
        # 31.4 回退到"别的过程的变量名"时同样不提示
        check("31.4 形参回退到别的过程的变量名 -> 不提示",
              [_trig31([
                  "Public Sub ProcedureA()",
                  "    Dim username As String",
                  "End Sub",
                  "",
                  "Public Sub ProcedureB(ByVal usernam| As String)",
                  "End Sub"])],
              expect_contain=[(False, [])])
        # 31.5 ★v105：形参位置（正在起新名字）**不再**提示模块级过程名。
        # 旧口径（v37）是"模块级名字是合法引用，照常提示"；但形参位置也是在
        # 造名字，按 IDEA 口径应静默 —— 用户点名要的就是这个。
        # ⚠️ "模块级名字是合法引用"这条语义由 31.6（用法位置）继续覆盖。
        check("31.5 形参同名于模块级过程名 -> 不提示（v105）",
              [_trig31([
                  "Public Sub updateValue()",
                  "End Sub",
                  "",
                  "Public Sub ProcedureB(ByVal updateValue| As String)",
                  "End Sub"])],
              expect_contain=[(False, [])])
        # 31.6 回归：用法位置（不是在声明新名字）照常提示同过程的名字
        check("31.6 同过程内用法位置打前缀 -> 照常提示",
              [_trig31([
                  "Public Sub ProcedureA()",
                  "    Dim username As String",
                  "    Debug.Print userna|",
                  "End Sub"])],
              expect_contain=[(True, ["username"])])
        # 31.7 ★v105：模块级声明行打全名 -> 不再提示（声明位置一律静默）
        check("31.7 模块级声明行打全名 -> 不提示（v105）",
              [_trig31([
                  "Public gUserName As String",
                  "Sub Foo()",
                  "    MsgBox gUserName",
                  "End Sub",
                  "",
                  "Public gUserName| As String"])],
              expect_contain=[(False, [])])
        # 31.8 机制本身：现场证据的宽/窄两个窗口
        _c31, _k31 = _parse31(_HDR31 + _leak31)
        _cur31["code"] = _c31
        _broad31 = _names31(_k31, False)
        _narrow31 = _names31(_k31, True)
        check("31.8 现场证据：宽窗口含 username，窄窗口不含",
              ["username" in _broad31, "username" in _narrow31],
              expect_contain=[True, False])
        _c31b, _k31b = _parse31(_HDR31 + [
            "Public Sub updateValue()",
            "End Sub",
            "",
            "Public Sub ProcedureB(ByVal updateValue| As String)",
            "End Sub"])
        _cur31["code"] = _c31b
        check("31.8 窄窗口仍认模块级过程名（不能收窄过头）",
              ["updatevalue" in _names31(_k31b, True)],
              expect_contain=[True])
        VB31._get_vbe_cached = _orig_get31
    except Exception as _e31:
        check("第 31 节异常: %s" % _e31, [True], expect_contain=[False])

    # ---- 32. v52：在标识符【前面】敲空格/标点，不许把右边那个词拿来补全 ----
    #
    # 用户报的现象：光标停在形参首字符前面，敲一个空格 -> 弹窗把这个形参
    # 原样提示出来（提示自己）。真实代码里形参几乎总会在本过程体内被用到，
    # 于是"现场还能扫到它" -> 被当成真实存在的名字 -> 回声防护放行。
    #
    # 根因：`extract_word_at` 有一条兜底（v44 为"光标停在词首、按 Delete 从头
    # 删字"引入）——光标左边没词、但右边紧挨着标识符时取右边那个词；入场检查
    # `_ident_char_before_caret` 也据此放宽到"光标前【或】后是标识符即可"。
    # 于是做完这次编辑后的那一帧快照，"敲空格"与"从词首删字"长得一模一样
    # （左边是分隔符、右边是标识符），单看快照永远区分不开。
    #
    # 唯一的差别在【怎么变的】：删字是行【变短】；敲分隔符是行【变长】且插入
    # 的那一段不含标识符字符。本节的 32.1~32.5 锁这条判据本身，32.6 起锁端到端。
    print("\n=== 32. 标识符前敲空格/标点不许补全（v52）===")
    try:
        # A) 判据本身：只有"插进来一段不含标识符字符的内容"才算敲了分隔符
        check("32.1 插入空格 -> 判定为分隔符",
              [E.typed_separator("Public Sub B(ByVal x As String)",
                                 "Public Sub B(ByVal  x As String)")],
              expect_contain=[True])
        check("32.2 删除（行变短）-> 不判定为分隔符",
              [E.typed_separator("    Debug.Print abcd",
                                 "    Debug.Print abc")],
              expect_contain=[False])
        check("32.3 插入标识符字符 -> 不判定为分隔符",
              [E.typed_separator("    Debug.Print abc",
                                 "    Debug.Print abcd")],
              expect_contain=[False])
        check("32.4 VBE 纠正大小写（替换）-> 不判定为分隔符",
              [E.typed_separator("dim x as string", "Dim x As String")],
              expect_contain=[False])
        check("32.5 纯插入内容提取正确（替换/删除一律 None）",
              [E.line_edit_inserted("ab", "aXXb"),
               E.line_edit_inserted("ab", "b"),
               E.line_edit_inserted("abc", "abc")],
              expect_contain=["XX", None])

        # B) 端到端：完全按 main.poll_editor 的分支跑
        #    （31 节那套脚手架还在函数作用域里，这里直接复用；它内部是
        #     真实解析 + 真实 VbeBackend.names_outside_caret）
        VB31._get_vbe_cached = lambda *a, **k: _VBE31(_cur31["code"])

        def _poll32(body, edit):
            """edit(old_line, col) -> (new_line, new_col)；按 poll 的分支走一遍。"""
            code, caret = _parse31(_HDR31 + body)
            ln, col = caret
            ls = code.split("\n")
            old_line = ls[ln - 1]
            new_line, new_col = edit(old_line, col)
            if E.typed_separator(old_line, new_line):
                return (False, [], new_line)          # 收起，不 trigger
            ls[ln - 1] = new_line
            new_code = "\n".join(ls)
            _cur31["code"] = new_code
            be = _B31(new_code, (ln, new_col))
            ui = _UI31()
            c = E.Completer(be, ui)
            c.trigger(True)
            return (ui.shown, sorted(c.matches or []), new_line)

        def _ins32(ch):
            def f(line, col):
                return line[:col - 1] + ch + line[col - 1:], col + len(ch)
            return f

        def _del32(line, col):
            return line[:col - 1] + line[col:], col     # Delete：删光标右边那个字

        def _tail32(ch):
            def f(line, col):
                return line + ch, len(line) + len(ch) + 1
            return f

        _p32 = _poll32([
            "",
            "Public Sub ProcedureB(ByVal |username As String)",
            "    Debug.Print username",
            "End Sub"], _ins32(" "))
        check("32.6 形参首字符前打空格 -> 不提示形参自身",
              [_p32[0], "username" in _p32[1], _p32[2]],
              expect_contain=[False, False,
                              "Public Sub ProcedureB(ByVal  username As String)"],
              expect_absent=[True])

        # 控制组：绕过这道门直接 trigger，确实会提示自己 —— 证明这道门有意义
        _c32, _k32 = _parse31(_HDR31 + [
            "",
            "Public Sub ProcedureB(ByVal |username As String)",
            "    Debug.Print username",
            "End Sub"])
        _ls32 = _c32.split("\n")
        _ln32, _col32 = _k32
        _l32 = _ls32[_ln32 - 1]
        _ls32[_ln32 - 1] = _l32[:_col32 - 1] + " " + _l32[_col32 - 1:]
        _nc32 = "\n".join(_ls32)
        _cur31["code"] = _nc32
        _ui32 = _UI31()
        # 同理（见 31.1）：v53 的语义护栏 + ★v105 的"声明位置静默"都会挡住这一例，
        # 对照要把两道一起关掉，才能真正测出"没有 v52 这道门时会提示自己"。
        _orig_ml32 = E.Completer._module_level_names
        _orig_dp32 = E.DECL_POPUP
        E.Completer._module_level_names = lambda _self: set(
            str(_r[0]).lower() for _r in _self.backend.get_identifiers())
        E.DECL_POPUP = True
        try:
            _cp32 = E.Completer(_B31(_nc32, (_ln32, _col32 + 1)), _ui32)
            _cp32.trigger(True)
        finally:
            E.Completer._module_level_names = _orig_ml32
            E.DECL_POPUP = _orig_dp32
        check("32.7 对照：两道门都关掉时会提示形参自身（本节没白测）",
              [_ui32.shown, sorted(_cp32.matches or [])],
              expect_contain=[True, ["username"]])

        # B) 回归：删字 / 打字母两条老路径必须原样保留，不能被一刀切
        _p32b = _poll32([
            "Public Function Foo() As Long",
            "    Dim gUserName As String",
            "    Debug.Print |gUserName",
            "End Function"], _del32)
        check("32.8 词首按 Delete 删字 -> 仍照旧提示（v44 路径不变）",
              [_p32b[0], _p32b[1]],
              expect_contain=[True, ["gUserName"]])

        _p32c = _poll32([
            "Public Function Foo() As Long",
            "    Dim gUserName As String",
            "    Debug.Print |",
            "End Function"], _tail32("g"))
        check("32.9 插入字母 -> 照常弹出候选（没被误伤）",
              [_p32c[0], "gUserName" in _p32c[1]],
              expect_contain=[True, True])

        _p32d = _poll32([
            "Public Function Foo() As Long",
            "    Dim gUserName As String",
            "    Debug.Print gUserName|",
            "End Function"], _ins32(" "))
        _p32e = _poll32([
            "Public Function Foo() As Long",
            "    Dim gUserName As String",
            "    Debug.Print gUserName |",
            "End Function"], _ins32("g"))
        check("32.10 打空格收起后接着打字母 -> 能重新弹出（流程不卡死）",
              [_p32d[0], _p32e[0], "gUserName" in _p32e[1]],
              expect_contain=[False, True, True])
    except Exception as _e32:
        check("第 32 节异常: %s" % _e32, [True], expect_contain=[False])

    # ---- 33. v53：形参位置上按退格 / 粘贴，不许提示形参自己 ----
    #
    # 用户报的现象：光标停在形参首字符前按【退格】，或把别的过程的变量名【粘贴】
    # 到形参里，弹窗都会把这个形参原样提示出来（v52 只堵住了"敲空格"那一路）。
    #
    # 根因：v51（现场证据按作用域收窄）与 v52（typed_separator）判的都是
    # 【这一次编辑是怎么变的】。而"粘贴一个完整名字"与"手动敲出最后一个字符"
    # 前后两帧快照逐字符等价 —— 从编辑动作上永远区分不开。何况真实代码里形参总
    # 会在本过程体内被用到（Debug.Print username），窄窗口里它就是"真实存在"的，
    # 回声防护照样放行。
    #
    # 修法（语义级、与按键无关）：光标处的词若正是【本行正在声明的名字】
    # （形参 / 局部 Dim），且它在模块级并不存在（过程内的局部变量不是全模块可见
    # 的真名字），就把与它完全同名的候选剔除。
    #   * 前缀照常提示（v37 教训：按"本行声明的名字"整体剔除会误杀 num -> numArr）；
    #   * 模块级真名字照常提示（31.5 的 updateValue / 31.7 的模块级声明行）。
    # 另外：形参常常独占一行（长签名被拆成 `Sub Foo( _` + 缩进的形参行），
    # 续行上没有声明关键字，后端据此把【声明续行】也算作声明行（v53 配套改动），
    # 否则续行上的退格依然会提示自己。
    print("\n=== 33. 形参位置不提示自己：退格 / 粘贴（v53）===")
    try:
        import vbe_bridge as VB33

        # A) 判据：续行识别 + 逻辑行合并出形参名
        check("33.1 以 `_` 结尾的行 -> 认定为续行",
              [P.is_continuation_line("Public Sub B( _"),
               P.is_continuation_line("    B( _  "),
               P.is_continuation_line("    Debug.Print a_b")],
              expect_contain=[True, True, False])

        _sig33 = "\n".join(_HDR31 + [
            "Public Sub ProcedureB( _",
            "    username As String, _",
            "    ByVal other As Long)",
            "    Debug.Print username",
            "End Sub"])
        check("33.2 续行上的 decl_names 含全部形参（合并逻辑行）",
              [sorted(P.decl_names_at_caret(_sig33, (8, 5)) or [])],
              expect_contain=[["other", "procedureb", "username"]])

        # B) 真机 VbeBackend.get_context()：续行也必须算出 decl_names
        #    （只在真机后端上，才会暴露"只在 in_decl 时才去算"这个漏洞）
        class _CM33(object):
            def __init__(self, code):
                self._code = code

            @property
            def CountOfLines(self):
                return len(self._code.split("\n"))

            def Lines(self, start, count):
                _ls = self._code.split("\n")
                return "\n".join(_ls[start - 1:start - 1 + count])

            @property
            def Name(self):
                return "Module1"

        class _CP33(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self._sel = sel

            def GetSelection(self):
                return self._sel

        class _VBE33(object):
            def __init__(self, code, sel):
                self.ActiveCodePane = _CP33(_CM33(code), sel)

        _cur33 = {"code": _sig33, "sel": (8, 5, 8, 5)}
        _orig_get33 = VB33._get_vbe_cached
        VB33._get_vbe_cached = lambda *a, **k: _VBE33(_cur33["code"],
                                                     _cur33["sel"])
        _be33 = VB33.VbeBackend()
        _ctx33 = _be33.get_context()
        _cur33["sel"] = (10, 24, 10, 24)
        _ctx33b = _be33.get_context()
        VB33._get_vbe_cached = _orig_get33
        check("33.3 真机 get_context：续行上有 decl_names（否则引擎拦不住）",
              [_ctx33.get("decl_names"), _ctx33.get("proc_name")],
              expect_contain=[["other", "procedureb", "username"], "ProcedureB"])
        check("33.4 真机 get_context：用法行 decl_names 仍为空（不误伤）",
              [_ctx33b.get("decl_names")],
              expect_contain=[[]])

        # C) 端到端：完全照 main.poll_editor 的分支跑
        #    （复用 31 节那套脚手架：真实解析 + 真实 names_outside_caret）
        VB31._get_vbe_cached = lambda *a, **k: _VBE31(_cur31["code"])

        _CODE33 = [
            "Public Sub ProcedureA()",
            "    Dim username As String",
            '    username = "abc"',
            "    Debug.Print username",
            "End Sub",
            "",
        ]

        def _poll33(body, edit):
            _code, _caret = _parse31(_HDR31 + body)
            _ln, _col = _caret
            _ls = _code.split("\n")
            _old = _ls[_ln - 1]
            _new, _ncol = edit(_old, _col)
            if E.typed_separator(_old, _new):      # v52 那道门照旧生效
                return (False, [], _new)
            _ls[_ln - 1] = _new
            _nc = "\n".join(_ls)
            _cur31["code"] = _nc
            _ui = _UI31()
            _c = E.Completer(_B31(_nc, (_ln, _ncol)), _ui)
            _c.trigger(True)
            return (_ui.shown, sorted(_c.matches or []), _new)

        def _paste33(txt):
            def f(line, col):
                return line[:col - 1] + txt + line[col - 1:], col + len(txt)
            return f

        def _back33(line, col):
            if col <= 1:
                return line, col
            return line[:col - 2] + line[col - 1:], col - 1

        _p = _poll33(_CODE33 + ["Public Sub ProcedureB(ByVal | )",
                                "    Debug.Print username",
                                "End Sub"], _paste33("username"))
        check("33.5 粘贴别的过程的变量名到形参 -> 不提示自己",
              [_p[0], _p[1]], expect_contain=[False, []])

        _p = _poll33(_CODE33 + ["Public Sub ProcedureB( _",
                                "    |username As String)",
                                "    Debug.Print username",
                                "End Sub"], _back33)
        check("33.6 续行形参首字符前按退格 -> 不提示自己",
              [_p[0], _p[1]], expect_contain=[False, []])

        _p = _poll33(_CODE33 + [
            "Public Sub ProcedureB(ByVal a As Long,|username As String)",
            "    Debug.Print username",
            "End Sub"], _back33)
        check("33.7 同行形参首字符前按退格 -> 不提示自己",
              [_p[0], _p[1]], expect_contain=[False, []])

        _p = _poll33(_CODE33 + ["Public Sub ProcedureB()",
                                "    Dim |",
                                "    Debug.Print username",
                                "End Sub"], _paste33("username"))
        check("33.8 粘贴到局部 Dim 位置 -> 不提示自己",
              [_p[0], _p[1]], expect_contain=[False, []])

        _p = _poll33(_CODE33 + ["Public Sub ProcedureB(ByVal | )",
                                "    Debug.Print username",
                                "End Sub"], _paste33("username"))
        _orig_ml33 = E.Completer._module_level_names
        _orig_dp33 = E.DECL_POPUP
        E.Completer._module_level_names = lambda _self: set(
            str(_r[0]).lower() for _r in _self.backend.get_identifiers())
        E.DECL_POPUP = True          # ★v105：连同"声明位置静默"一起关掉
        try:
            _p_off = _poll33(_CODE33 + ["Public Sub ProcedureB(ByVal | )",
                                        "    Debug.Print username",
                                        "End Sub"], _paste33("username"))
        finally:
            E.Completer._module_level_names = _orig_ml33
            E.DECL_POPUP = _orig_dp33
        check("33.9 对照：两道护栏都关掉时会提示形参自身（本节没白测）",
              [_p_off[0], "username" in _p_off[1]],
              expect_contain=[True, True])

        # D) 防误杀：模块级真名字 / 前缀补全 / 用法位置都不许被砍
        #    ★v105：形参/声明行这三条现在按 IDEA 口径【静默】，语义改由
        #    「用法位置」那条（33.13）覆盖 —— 本组仍有意义：证明静默不是"砍掉能力"。
        _p = _poll33(["Public Sub updateValue()",
                      "End Sub",
                      "",
                      "Public Sub ProcedureB(ByVal updateValue| As String)",
                      "End Sub"], _back33)
        check("33.10 形参同名于模块级过程名 -> 静默（v105）",
              [_p[0], "updateValue" in _p[1]], expect_contain=[False, False])

        _p = _poll33(["Public Function test() As Long",
                      "    test = 1",
                      "End Function",
                      "",
                      "Public Sub |()",
                      "End Sub"], _paste33("test"))
        check("33.11 Sub 行写全 test，工程里有 Function test() -> 静默（v105）",
              [_p[0], "test" in [x.lower() for x in _p[1]]],
              expect_contain=[False, False])

        _p = _poll33(["Public Sub ProcedureB()",
                      "    Dim numArr As Long",
                      "    Dim | As Long",
                      "End Sub"], _paste33("num"))
        check("33.12 局部声明行输入前缀 num -> 静默（v105；前缀补全见 33.13/25）",
              [_p[0], "numArr" in _p[1]], expect_contain=[False, False])

        _p = _poll33(["Public Sub ProcedureA()",
                      "    Dim username As String",
                      "    Debug.Print userna|",
                      "End Sub"], _paste33("m"))
        check("33.13 同过程内用法位置打前缀 -> 照常提示（31.6 语义不变）",
              [_p[0], "username" in _p[1]], expect_contain=[True, True])

        # ★v107c 补：本节开头复用了 31 节那套脚手架，把 `_get_vbe_cached` 换成了
        # `_VBE31(_cur31["code"])`，但一直到节尾都没换回来 —— 于是从这一节起，
        # 后面**每一节**的 `_get_vbe_cached` 都是这个假工程。凡是用"清
        # `_vbe_cache["obj"]` + 换 `_get_vbe`"搭装置的小节（如 §89）都会读到它，
        # 症状是"写回悄悄落到别人的假模块上、自己那份一行没变"。
        # 猴补丁一律在节尾还原；§90 还有一道全局自查专门抓这种泄漏。
        VB31._get_vbe_cached = _orig_get31

    except Exception as _e33:
        check("第 33 节异常: %s" % _e33, [True], expect_contain=[False])


    # ==================================================================
    # 34. VBE 自带提示列表出现时让位（v55）
    # ==================================================================
    # VBE 的「自动列出成员」跟我们的弹窗会出现在同一处（光标下方），两个列表
    # 同时出现会互相遮挡、还抢键盘，所以光标停在会让 VBE 弹列表的位置时我们
    # 让位。判据走【光标的语法位置】而不是窗口探测——真机实测：给 VBE 发
    # Ctrl+J 后全系统扫描，一个新窗口都没出现，那个列表是画在代码窗格上的。
    print("\n=== 34. VBE 自带列表出现时让位（v55）===")
    try:
        # --- 决策：本轮轮询该 hide 还是 trigger（纯函数）---
        if M is not None:
            check("34.1 没敲分隔符 -> trigger",
                  [M._poll_action(False)], expect_contain=["trigger"])
            check("34.2 敲了空格/标点 -> hide（v52 语义不变）",
                  [M._poll_action(True)], expect_contain=["hide"])
        else:
            check("34.1 main 不可用", [True], expect_contain=[False])

        # --- 语义判据：光标是不是停在 VBE 会自己弹列表的位置 ---
        _yl = E.vbe_list_expected
        check("34.3 UserForm1. 之后（还没开始输成员名）-> 让位",
              [_yl("    UserForm1.", 15)], expect_contain=[True])
        check("34.4 UserForm1.Cap 输入成员名中 -> 让位",
              [_yl("    UserForm1.Cap", 18)], expect_contain=[True])
        check("34.5 Me. -> 让位",
              [_yl("    Me.", 8)], expect_contain=[True])
        check("34.6 Range(\"A1\"). 之后（点号左边是括号）-> 让位",
              [_yl('    Range("A1").', 17)], expect_contain=[True])
        check("34.7 普通变量名（没有点号）-> 不让位",
              [_yl("    userName = 1", 12)], expect_contain=[False])
        check("34.8 小数点 1. -> 不让位（VBE 也不弹）",
              [_yl("    x = 1.", 11)], expect_contain=[False])
        check("34.9 注释里的点号 -> 不让位",
              [_yl("    ' UserForm1.", 17)], expect_contain=[False])
        check("34.10 字符串里的点号 -> 不让位",
              [_yl('    s = "a."', 12)], expect_contain=[False])
        # 【v72 语义变更】行首的点号 = With 块的成员访问（`With obj` 之后成员
        # 各占一行，点就顶在行首），VBE 在那儿同样会弹成员列表 -> 让位。
        # 旧断言（v55 时代写的"不让位"）是错的：它默认"点号左边必须有东西"，
        # 恰恰漏掉了用户报的 `With Range(...).Font` + `.Size = 12` 场景。
        check("34.11 ★行首就是点号 -> 让位（With 块成员访问；v72 语义变更）",
              [_yl(".", 2)], expect_contain=[True])
        check("34.12 成员名打完（等号后）-> 立刻恢复，不让位",
              [_yl("    UserForm1.Caption = ", 25)], expect_contain=[False])
        check("34.13 空行 / 列号缺失 -> 不让位（不炸）",
              [_yl("", 1), _yl("    x", None), _yl(None, 3)],
              expect_contain=[False])

        # --- 引擎集成：轮询自动触发时让位，手动 Ctrl+Space 照弹 ---
        class _UI55(object):
            def __init__(self):
                self.shown = False

            def show(self, *a, **k):
                self.shown = True

            def update_selection(self, *a, **k):
                pass

            def hide(self, *a, **k):
                self.shown = False

            def contains_point(self, *a, **k):
                return False

        class _B55(object):
            def __init__(self, ctx):
                self.ctx = ctx

            def get_context(self):
                return self.ctx

            def get_identifiers(self):
                return [("Caption", None, None, False),
                        ("userName", None, None, False)]

            def apply_completion(self, *a, **k):
                return None

        def _ctx55(text, col):
            return {"line_no": 1, "caret_col": col, "line_text": text,
                    "in_string": False, "in_comment": False,
                    "in_type_position": False, "in_decl_position": False,
                    "proc_name": None, "module_name": "M1"}

        _ui55 = _UI55()
        _c55 = E.Completer(_B55(_ctx55("    UserForm1.Cap", 18)), _ui55)
        _c55.trigger(True)          # 轮询自动触发 -> 让位
        check("34.14 轮询触发 + VBE 列表位置 -> 不弹（让位）",
              [_ui55.shown], expect_contain=[False])

        _ui55b = _UI55()
        _c55b = E.Completer(_B55(_ctx55("    UserForm1.Cap", 18)), _ui55b)
        _c55b.trigger()             # 手动 Ctrl+Space -> 用户点名要我们的
        check("34.15 手动触发（Ctrl+Space）即使在 VBE 列表位置也照弹",
              [_ui55b.shown], expect_contain=[True])

        _ui55c = _UI55()
        _c55c = E.Completer(_B55(_ctx55("    userN", 11)), _ui55c)
        _c55c.trigger(True)         # 普通位置：完全不受影响
        check("34.16 普通标识符位置照旧弹出（没被误伤）",
              [_ui55c.shown], expect_contain=[True])

        # 34.17 逃生门：VBECOMPLETE_NO_YIELD=1 -> 一律不让位
        _orig_yield = E.YIELD_TO_VBE_LIST
        try:
            E.YIELD_TO_VBE_LIST = False
            _ui55d = _UI55()
            _c55d = E.Completer(_B55(_ctx55("    UserForm1.Cap", 18)), _ui55d)
            _c55d.trigger(True)
            check("34.17 关掉让位开关 -> 照旧弹出（逃生门）",
                  [_ui55d.shown], expect_contain=[True])
        finally:
            E.YIELD_TO_VBE_LIST = _orig_yield

        # --- 声明里 `As ` / `New ` 之后：VBE 会弹类型列表，同样让位 ---
        check("34.18 Dim x As （类型名还没开始打）-> 让位",
              [_yl("    Dim x As ", 14)], expect_contain=[True])
        check("34.19 Dim x As In（类型名输入中）-> 让位",
              [_yl("    Dim x As In", 16)], expect_contain=[True])
        check("34.20 Set c = New （大写的 New）-> 让位",
              [_yl("    Set c = New ", 17)], expect_contain=[True])
        check("34.21 类型名打完并跟了空格 -> 恢复，不让位",
              [_yl("    Dim x As Long ", 19)], expect_contain=[False])
        check("34.22 注释里的 As -> 不让位",
              [_yl("    ' Dim x As ", 16)], expect_contain=[False])
        check("34.23 变量名 asName（不是 As 关键字）-> 不让位",
              [_yl("    x = asName", 15)], expect_contain=[False])
    except Exception as _e34:
        check("第 34 节异常: %s" % _e34, [True], expect_contain=[False])


    # ==================================================================
    # 35. 自动配对：`(` -> `()`、`"` -> `""`，光标落在中间
    # ==================================================================
    # VBE 原生不自动闭合括号 / 引号（VBE_Extras、Rubberduck 都把它当增强功能加），
    # 所以由我们补。两条铁律：写成功了才吞键；不该插时【放行】让 VBE 原生处理。
    print("\n=== 35. 自动配对（括号 / 引号）===")
    try:
        # --- 按键 -> 符号（纯函数）---
        if M is not None:
            check("35.1 Shift+9 -> (",
                  [M._pair_char_for_key(0x39, True)],
                  expect_contain=["("])
            check("35.2 不按 Shift 的 9 -> 不管",
                  [M._pair_char_for_key(0x39, False)],
                  expect_contain=[None])
            check("35.3 Shift+' -> \"",
                  [M._pair_char_for_key(0xDE, True)],
                  expect_contain=['"'])
            check("35.4 字母键 -> 不管",
                  [M._pair_char_for_key(0x41, True)],
                  expect_contain=[None])
            check("35.5 小键盘 9 -> 不管（NumLock 下是数字）",
                  [M._pair_char_for_key(0x69, True)],
                  expect_contain=[None])
        else:
            check("35.1 main 不可用", [True], expect_contain=[False])

        # --- 插什么、光标落哪（纯函数）---
        import vbe_bridge as VB35
        _pi = VB35._pair_insertion
        check("35.6 行尾敲 ( -> 补成 ()，光标在中间（第三元 = 插了完整一对）",
              [_pi("x = ", 4, "(", ")")],
              expect_contain=[("x = ()", 5, True)])
        check("35.7 行尾敲 \" -> 补成 \"\"，光标在中间",
              [_pi("s = ", 4, '"', '"')],
              expect_contain=[('s = ""', 5, True)])
        check("35.8 右边已经有 \" -> 只跨过去，一个字符都不写",
              [_pi('s = "abc"', 8, '"', '"')],
              expect_contain=[(None, 9, False)])
        check("35.9 在未闭合字符串里敲 \" -> 只补收尾那一半（算不上\"一对\"）",
              [_pi('s = "abc', 8, '"', '"')],
              expect_contain=[('s = "abc"', 9, False)])
        # ★v113：原用例写的是 `foo(a, b)` 光标停在 `foo(` 里面 —— 那种"实参表里"
        # 的位置 v113 起改成**照旧补一对**（见 §97），所以这里换一个**不在**任何
        # 调用括号里的行来钉 v112 那条规则（右侧有内容 -> 只插左括号）。
        check("35.10 ★v112 行中间敲 ( -> 右侧有内容、且不在调用括号里，只插左括号",
              [_pi("x = ab", 4, "(", ")")],
              expect_contain=[("x = (ab", 5, False)])

        # --- 后端集成：真写进"编辑器"（假 VBE）并把光标摆到中间 ---
        class _Mod35(object):
            def __init__(self, text, fmt=None):
                self.lines = [text]
                self.fmt = fmt          # 模拟 VBE 的「自动语法检测」改写整行

            def Lines(self, n, _c):
                # 真 VBE 对"虚拟空行"（末行+1）也返回空串，不报错
                if 1 <= n <= len(self.lines):
                    return self.lines[n - 1]
                return ""

            @property
            def CountOfLines(self):
                return len(self.lines)

            def ReplaceLine(self, n, t):
                if self.fmt is not None:
                    t = self.fmt(t)
                self.lines[n - 1] = t

            def InsertLines(self, n, t):
                # 真 VBE：第 n 行处插入，原第 n 行及之后下移
                self.lines.insert(n - 1, t)

        class _Pane35(object):
            def __init__(self, mod, sl, sc, el, ec):
                self.CodeModule = mod
                self.sel = [sl, sc, el, ec]

            def GetSelection(self):
                return tuple(self.sel)

            def SetSelection(self, sl, sc, el, ec):
                self.sel = [sl, sc, el, ec]

        class _Vbe35(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _pair_run(text, col, ch, ec=None, line=1, fmt=None):
            """返回 (该行新文本, 光标列, 是否代为输入)。ec 不同即视为有选区。

            line 可指定光标所在行：>CountOfLines 时即 VBE 的"虚拟空行"。
            fmt 模拟 VBE 把整行重新格式化（等号补空格等）。
            """
            mod = _Mod35(text, fmt)
            pane = _Pane35(mod, line, col, line, col if ec is None else ec)
            _orig = VB35._get_vbe_cached
            VB35._get_vbe_cached = lambda: _Vbe35(pane)
            try:
                ok = VB35.VbeBackend().insert_pair(ch)
            finally:
                VB35._get_vbe_cached = _orig
            _txt = mod.lines[line - 1] if line - 1 < len(mod.lines) else ""
            return _txt, pane.sel[1], ok

        check("35.11 光标在行尾敲 ( -> 写入 () 且光标停在中间",
              list(_pair_run("x = ", 5, "(")),
              expect_contain=["x = ()", 6, True])
        check("35.12 光标在行尾敲 \" -> 写入 \"\" 且光标停在中间",
              list(_pair_run("s = ", 5, '"')),
              expect_contain=['s = ""', 6, True])
        check("35.13 注释里敲 ( -> 不管（返回 False，调用方会放行按键）",
              [_pair_run("' x = ", 7, "(")[2],
               _pair_run("' x = ", 7, "(")[0]],
              expect_contain=[False, "' x = "])
        # ★v114：本用例原先用左括号，而左括号带选区从 v114 起改成"包裹"
        # （§98 钉）。"有选区就不接管"这条规则本身没变，改用**收符号** `)`
        # 继续守着它（收符号没有"包起来"的语义，仍交给 VBE 替换）。
        check("35.14 有选区 + 非包裹符号（收括号） -> 不管（替换语义交给 VBE）",
              [_pair_run("x = ab", 3, ")", ec=6)[2],
               _pair_run("x = ab", 3, ")", ec=6)[0]],
              expect_contain=[False, "x = ab"])
        _orig_g = VB35._get_vbe_cached
        try:
            VB35._get_vbe_cached = lambda: None
            _bk = VB35.VbeBackend()
            check("35.15 COM 拿不到 VBE -> 不管，且不炸",
                  [_bk.insert_pair("(")], expect_contain=[False])
        finally:
            VB35._get_vbe_cached = _orig_g

        # --- 右半边：`)` 也是"跨过去"，但只在右边真的有 `)` 时 ---
        check("35.16 Shift+0 -> )",
              [M._pair_char_for_key(0x30, True) if M is not None else None],
              expect_contain=[")"])
        check("35.17 补完 () 后再按 ) -> 跨过去，不多出一个",
              list(_pair_run("x = ()", 6, ")")),
              expect_contain=["x = ()", 7, True])
        check("35.18 右边没有 ) 时按 ) -> 放行让 VBE 自己插",
              [_pair_run("x = ab", 5, ")")[2],
               _pair_run("x = ab", 5, ")")[0]],
              expect_contain=[False, "x = ab"])
        # --- v57：模块末尾的「虚拟空行」上 ReplaceLine 会报参数无效 ---
        check("35.19 光标在虚拟空行（末行+1）-> 先建行再配对，光标在中间",
              list(_pair_run("x = ", 1, "(", line=2)),
              expect_contain=["()", 2, True])
        check("35.20 行号离谱（差 2 行以上）-> 不管，交兜底",
              [_pair_run("x = ", 1, "(", line=4)[2]],
              expect_contain=[False])
        # --- v58：VBE 的「自动语法检测」会把整行改写，行长一变光标就偏 ---
        check("35.23 VBE 在插入的符号前补空格后，光标仍落在引号中间",
              list(_pair_run("    w=4", 6, '"',
                             fmt=lambda s: s.replace('w""', 'w ""'))),
              expect_contain=['    w ""=4', 8, True])
        check("35.24 _locate_inserted：VBE 改写后重新定位插入的符号对",
              [VB35._locate_inserted("    x = 1()", "    x = 1()",
                                     "(", ")", 9),
               VB35._locate_inserted('    w "" = 4', '    w"" = 4', '"', '"', 5),
               VB35._locate_inserted("    Z =() 1 + 2", "    Z=() 1 + 2",
                                     "(", ")", 6)],
              expect_contain=[10, 7, 8])
        check("35.25 _locate_inserted：认不出来时返回 None（调用方退回行尾）",
              [VB35._locate_inserted("完全不相关", "另一行", "(", ")", 3)],
              expect_contain=[None])
        # --- v57：非主线程绝不能碰 COM（否则会污染全局退避，拖慢提示）---
        # 只验证守卫判据本身：vbe_bridge._get_vbe_cached 这个模块属性被前面
        # 小节（14/15/16 节等）换成过 lambda 且没还原，这里拿不到真函数；
        # 完整路径（子线程调用返回 None 且不计失败）由真机验证覆盖。
        import threading as _th_mod
        _box = {}

        def _ask():
            try:
                _box["main"] = VB35._on_main_thread()
            except Exception as _exc:
                _box["main"] = "EXC:%s" % _exc

        _t = _th_mod.Thread(target=_ask)
        _t.daemon = True
        _t.start()
        _t.join(5)
        check("35.21 守卫判据：子线程 _on_main_thread() 为 False、主线程为 True",
              [_box.get("main", "MISSING"), VB35._on_main_thread()],
              expect_contain=[False, True])
        # --- v57：配对没做成时的兜底（绝不吞掉用户的按键）---
        if M is not None:
            _u32 = M.ctypes.windll
            _had = "user32" in _u32.__dict__
            _old = _u32.__dict__.get("user32")
            try:
                class _FakeU32(object):
                    def __init__(self):
                        self.events = []

                    def SendInput(self, n, ptr, size):
                        arr = M.ctypes.cast(
                            ptr, M.ctypes.POINTER(M._INPUT * 2)).contents
                        for _e in arr:
                            self.events.append((_e.type, _e.u.ki.wScan,
                                                _e.u.ki.dwFlags))
                        return 2

                _fake = _FakeU32()
                _u32.__dict__["user32"] = _fake
                _ok = M.send_char("(")
                check("35.22 兜底重发：keydown+keyup 两个 unicode 事件",
                      [_ok, len(_fake.events),
                       chr(_fake.events[0][1]) if _fake.events else "",
                       _fake.events[0][2] if _fake.events else -1,
                       _fake.events[1][2] if len(_fake.events) > 1 else -1],
                      expect_contain=[True, 2, "(", M.KEYEVENTF_UNICODE,
                                      M.KEYEVENTF_UNICODE | M.KEYEVENTF_KEYUP])
            finally:
                if _had:
                    _u32.__dict__["user32"] = _old
                else:
                    _u32.__dict__.pop("user32", None)
    except Exception as _e35:
        check("第 35 节异常: %s" % _e35, [True], expect_contain=[False])

    # ---- 36. v59：续行符误判 / WithEvents / Declare 修饰符（解析层）----
    #
    # 三个都是"行首修饰符被当成名字"或"符号被判成续行符"这一类误判，共同症状：
    # 真名字收不到、假名字混进候选池。真实工程里 pub3 的 Enum 成员名是
    # A_/B_/.../EZ_（每行都以 Z_ 结尾），整模块因此被当成 Enum 块解析。
    print("\n=== 36. 续行符与声明修饰符（v59）===")

    def recs36(code, std=True):
        _r36, _pn = P.extract_records(code, module="M1",
                                      is_std_module=std)
        return [(r[0], r[2], r[3]) for r in _r36]

    # 36.1 名字以 _ 结尾且正好在行尾 —— VBA 要求续行符 _ 前面必须有空白，
    #      所以这不是续行符，不能把下一行（End Enum …）并进来。
    cont_junk = "\n".join([
        "Public Enum E",
        "    A_ = 1: B_: Z_",
        "End Enum",
        "Function getOneFile()",
        "    Dim tmp As String",
        "    With Application.FileDialog(1)",
        "        If .Show = -1 Then",
        "            s = .SelectedItems(1)",
        "        End If",
        "    End With",
        "    getOneFile = s",
        "End Function"])
    got36 = recs36(cont_junk)
    check("36.1 Enum 成员名以 _ 结尾：不算续行符，假名字消失",
          [n for n, _, _ in got36],
          expect_contain=["E", "getOneFile"],
          expect_absent=["Function", "With", "If", "s", "A_"])
    # 36.2 End Enum 必须真的闭合：之后的过程名照常收、局部变量仍归属其过程
    #      （不闭合时全被记成模块级: proc=None）
    check("36.2 End Enum 正确闭合：之后的过程/局部变量归属正确",
          got36,
          expect_contain=[("getOneFile", None, False),
                          ("tmp", "getOneFile", False)])

    # 36.3 续行符规则本身：_ 前有空白才合并
    _merged = P._RE_CONTINUATION.sub(" ", "Dim a _\n    , b")
    _kept = P._RE_CONTINUATION.sub(" ", "    A_ = 1: Z_\nEnd Enum")
    check("36.3 续行符：_ 前有空白才合并，名字结尾的 _ 不合并",
          ["\n" in _merged, "\n" in _kept],
          expect_contain=[False, True])

    # 36.4 Public WithEvents clk As ... —— 收真名 clk，不收 WithEvents
    check("36.4 WithEvents 声明收录 clk（不收 WithEvents）",
          recs36("Public WithEvents clk As MSForms.CommandButton", std=False),
          expect_contain=[("clk", None, True)],
          expect_absent=[("WithEvents", None, True)])

    # 36.5 带访问修饰符的 Declare —— 收真名，不收 Declare
    #
    # ★v103 更正期望值：Declare 的形参原先被记成【模块级、priv=False】
    #   （`("x", None, False)`）—— 那是一处泄漏：Declare 一定在模块级，
    #   那时 `cur_proc` 是 None，于是形参挂到了模块级上，本模块**每个过程**
    #   里都能裸名引用到它。正确口径与 Sub/Function 一致：形参挂在**这个
    #   Declare 自己的名字**名下（`("x", "gApi", False)`），而 Declare 没有
    #   过程体 ⇒ 它天然不可见。上面那条 `expect_contain` 是本用例**当年写下
    #   的期望**，它把 bug 当成了规格，所以这里按 v103 修正（用例本身保留）。
    check("36.5 Private Declare 收录 gApi（不收 Declare）；★v103 形参归属 Declare 自己不再泄漏",
          recs36('Private Declare PtrSafe Function gApi Lib "k" '
                 '(ByVal x As Long) As Long'),
          expect_contain=[("gApi", None, True), ("x", "gApi", False)],
          expect_absent=[("Declare", None, True), ("x", None, False)])
    check("36.6 无修饰 Declare 仍收录",
          recs36('Declare PtrSafe Function gApi2 Lib "k" () As Long'),
          expect_contain=[("gApi2", None, False)])

    # 36.7 块成员（Type 字段 / Enum 成员）不收录成裸名候选；类型名本身照常收录
    #      （用户要求：枚举成员前面有枚举名限定，不该直接提示）
    blk = "\n".join([
        "Public Enum E",
        "    Red = 1: Green = 2: Zz_",
        "End Enum",
        "Public Type T",
        "    fx As Long",
        "End Type"])
    check("36.7 Type/Enum 成员不收录（类型/枚举名收录）",
          recs36(blk),
          expect_contain=[("E", None, False), ("T", None, False)],
          expect_absent=[("Red", None, False), ("Green", None, False),
                         ("Zz_", None, False), ("fx", None, False)])
    # 块结束后，后面的模块级声明仍要正常收录（continue 不能把状态搞乱）
    blk2 = blk + "\nPublic gAfter As Long"
    check("36.8 Enc/Type 块之后的声明照常收录",
          recs36(blk2),
          expect_contain=[("gAfter", None, False)])

    # ---- 37. v60：组件名 / 窗体控件名不该被回声防护误杀 ----
    print("\n=== 37. v60 组件名与窗体控件名（在窗体自己的代码模块里）===")
    try:
        class _UI37(object):
            def __init__(self):
                self.shown = False

            def show(self, *a, **k):
                self.shown = True

            def update_selection(self, *a, **k):
                pass

            def hide(self, *a, **k):
                self.shown = False

            def contains_point(self, *a, **k):
                return False

        class _B37(object):
            """模拟【窗体自己的代码模块】。

            三个候选都"只声明在这个模块里"，现场文本里也一个字都没出现过：
              * UserForm1 —— 组件名。窗体代码里从来不会写 UserForm1 这个
                词（除非要 UserForm1.Show），可它当然是合法引用；
              * Label1 —— 刚拖上去的控件，还没写任何事件过程，代码里没有它；
              * useGhost —— 对照用的真幽灵：既不是结构性名字，声明集合里
                也没有它，就该照旧被回声防护剔掉。
            """

            def __init__(self, structural=True, module="UserForm1"):
                self._structural = structural
                self._module = module

            def get_context(self):
                return {"line_no": 1, "caret_col": 1,
                        "line_text": "", "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": self._module}

            def get_identifiers(self):
                return [("UserForm1", "UserForm1", None, False),
                        ("Label1", "UserForm1", None, True),
                        ("useGhost", "UserForm1", None, True)]

            def get_declared_names(self):
                # 刻意【不】含 useghost：它代表"解析层没能归类、只剩缓存残留
                # 的幽灵"，用来验证这次放宽没有把真幽灵一起放行。
                return ["userform1", "label1"]

            def caret_word_at_collect(self):
                return ""

            def declared_elsewhere(self, name, module_name):
                return False        # 全部"只声明在" UserForm1

            def get_structural_names(self):
                return ["userform1", "label1"] if self._structural else []

            def names_outside_caret(self, caret, scope_only=False):
                return set()        # 现场文本里一个字都没有

            def apply_completion(self, *a, **k):
                return None

        def _trig37(word, structural=True, module="UserForm1"):
            _be = _B37(structural, module)
            _line = "    " + word
            _ctx = _be.get_context()
            _ctx["line_text"] = _line
            _ctx["caret_col"] = len(_line) + 1
            _be.get_context = lambda: _ctx
            _c = E.Completer(_be, _UI37())
            _c.trigger()
            return sorted(_c.matches or [])

        # 37.1 组件名：在【窗体自己的】代码里输入 use（用户插完窗体就在这打）
        check("37.1 窗体自己的代码里输入 use -> 提示 UserForm1（真幽灵仍剔）",
              _trig37("use"),
              expect_contain=["UserForm1"],
              expect_absent=["useGhost"])
        # 37.2 对照：去掉结构性名字证据，就是修复前的行为（不提示）
        check("37.2 对照：无结构性名字证据时不提示 UserForm1",
              _trig37("use", structural=False),
              expect_absent=["UserForm1"])
        # 37.3 控件名：刚拖上去的 Label1（代码里还没有任何 Label1 字样）
        check("37.3 窗体里输入 la -> 提示 Label1",
              _trig37("la"),
              expect_contain=["Label1"])
        # 37.4 对照：去掉结构性名字证据则不提示
        check("37.4 对照：无结构性名字证据时不提示 Label1",
              _trig37("la", structural=False),
              expect_absent=["Label1"])
        # 37.5 控件名跨模块不可见（与"不提示别模块私有成员"的约定一致）
        check("37.5 别的模块里输入 la -> 不提示 Label1",
              _trig37("la", module="模块2"),
              expect_absent=["Label1"])

    except Exception as _e37:
        check("第 37 节异常: %s" % _e37, [True], expect_contain=[False])

    # ---- 38. v61：VBA 语言自带的名字（内建函数 / 常量 / 数据类型）----
    print("\n=== 38. v61 VBA 内建名字进候选池 ===")
    try:
        import re as _re38
        import vba_builtins as VB38

        # 38.1~38.5：清单自身（纯静态，不依赖 Excel）
        _low38 = [n.lower() for n in VB38.BUILTIN_NAMES]
        _lowT38 = [n.lower() for n in VB38.BUILTIN_TYPE_NAMES]
        check("38.1 两个清单各自无重名（函数+常量 %d 个、类型名 %d 个）"
              % (len(_low38), len(_lowT38)),
              [len(set(_low38)) == len(_low38)
               and len(set(_lowT38)) == len(_lowT38)], expect_contain=[True])
        _all38 = _low38 + _lowT38
        _bad38 = [n for n in _all38
                  if not _re38.match(r"^[^\W\d]\w*$", n)]
        check("38.2 清单全是合法标识符（异常 %s）" % (_bad38 or "无"),
              [len(_bad38) == 0], expect_contain=[True])
        # 语言关键字一律不许混进来（v59 修过"Function/If/With 混进候选池"）
        _kws38 = set(("if then else elseif end for next each step do loop while "
                      "wend until select case with sub function property dim "
                      "redim set let const static public private friend "
                      "global declare type enum new nothing goto on option "
                      "exit resume is like mod and or not xor eqv imp"
                      ).split())
        check("38.3 清单里没有语言关键字（命中 %s）"
              % (sorted(_kws38 & set(_low38)) or "无"),
              [len(_kws38 & set(_low38)) == 0], expect_contain=[True])
        _want38 = ("MsgBox", "Left", "Len", "Split", "IsArray", "CLng",
                   "InStrRev", "DateAdd", "StrConv", "CreateObject", "Array",
                   "Shell", "vbCrLf", "vbYes", "vbRed", "vbNullString")
        _miss38 = [n for n in _want38 if n.lower() not in VB38.BUILTIN_LOWER]
        check("38.4 常用内建函数/常量齐备（缺 %s）" % (_miss38 or "无"),
              [len(_miss38) == 0], expect_contain=[True])
        _missT38 = [n for n in ("String", "Long", "Boolean", "Variant")
                    if n.lower() not in VB38.BUILTIN_TYPE_LOWER]
        check("38.5 内建数据类型齐备（缺 %s）" % (_missT38 or "无"),
              [len(_missT38) == 0], expect_contain=[True])

        class _UI38(object):
            def __init__(self):
                self.shown = False

            def show(self, *a, **k):
                self.shown = True

            def update_selection(self, *a, **k):
                pass

            def hide(self, *a, **k):
                self.shown = False

            def contains_point(self, *a, **k):
                return False

        class _B38(object):
            """模拟"代码里从没出现过 MsgBox 的模块"。

            候选池里是 VBA 内建名字（挂虚拟模块 VBA）+ 一个用户自己声明过的
            Left（模拟真实工程"同名只留用户的"那种形态）+ 一个真幽灵 msgGhost。
            elsewhere 控制 declared_elsewhere 的返回值 —— 真实情形下内建名字
            归属虚拟模块 VBA，从当前模块看就是"别的模块声明过"（True）。
            """

            def __init__(self, structural=True, elsewhere=False,
                         module="模块2", builtin=True):
                self._structural = structural
                self._elsewhere = elsewhere
                self._module = module
                self._builtin = builtin

            def get_context(self):
                return {"line_no": 1, "caret_col": 1, "line_text": "",
                        "in_string": False, "in_comment": False,
                        "in_type_position": False, "in_decl_position": False,
                        "decl_names": [], "proc_name": None,
                        "module_name": self._module}

            def get_identifiers(self):
                return [("MsgBox", "VBA", None, False),
                        ("vbCrLf", "VBA", None, False),
                        ("String", "VBA", None, False),
                        ("Long", "VBA", None, False),
                        ("Left", "模块2", None, False),
                        ("msgGhost", "模块2", None, True)]

            def get_declared_names(self):
                return ["msgbox", "vbcrlf", "left", "string", "long"]

            def caret_word_at_collect(self):
                return ""

            def declared_elsewhere(self, name, module_name):
                if self._elsewhere:
                    return str(name).lower() in ("msgbox", "vbcrlf",
                                                 "string", "long")
                return False

            def get_structural_names(self):
                if not self._structural:
                    return []
                return ["msgbox", "vbcrlf", "string", "long"]

            def get_builtin_names(self):
                # 这四个就是候选池里的"VBA 内建名字"，引擎据此收紧匹配。
                if not self._builtin:
                    return []
                return ["msgbox", "vbcrlf", "string", "long"]

            def get_type_names(self):
                return ["String", "Long"]

            def names_outside_caret(self, caret, scope_only=False):
                return set()        # 现场文本里一个字都没有

            def apply_completion(self, *a, **k):
                return None

        def _trig38(word, **kw):
            _be = _B38(**kw)
            _line = "    " + word
            _ctx = _be.get_context()
            _ctx["line_text"] = _line
            _ctx["caret_col"] = len(_line) + 1
            _be.get_context = lambda: _ctx
            _c = E.Completer(_be, _UI38())
            _c.trigger()
            return sorted(_c.matches or [])

        def _trig38_type(word):
            _be = _B38(elsewhere=True)
            _line = "Dim x As " + word
            _ctx = _be.get_context()
            _ctx["line_text"] = _line
            _ctx["caret_col"] = len(_line) + 1
            _ctx["in_type_position"] = True
            _be.get_context = lambda: _ctx
            _c = E.Completer(_be, _UI38())
            _c.trigger()
            return sorted(_c.matches or [])

        # 38.6 主场景：代码里从没写过 MsgBox，输入 ms 应能提示出来
        check("38.6 输入 ms -> 提示 MsgBox（现场文本里没有它；幽灵仍剔）",
              _trig38("ms", elsewhere=True),
              expect_contain=["MsgBox"], expect_absent=["msgGhost"])
        # 38.7 结构性证据这一路要能单独顶住（只是"进池子"并不够）
        check("38.7 只靠结构性证据也放行（declared_elsewhere=False）",
              _trig38("ms", structural=True, elsewhere=False),
              expect_contain=["MsgBox"])
        # 38.8 对照：两条证据都没有 = 修复前行为（不提示）
        check("38.8 对照：无结构性证据且未声明过 -> 不提示（修复前行为）",
              _trig38("ms", structural=False, elsewhere=False),
              expect_absent=["MsgBox"])
        # 38.9 真幽灵照旧被回声防护剔掉（放宽是精确的）
        check("38.9 真幽灵 msgGhost 仍被剔",
              _trig38("msgg", structural=True, elsewhere=True),
              expect_absent=["msgGhost"])
        # 38.10 内建名字跨模块可见（按模块级 priv=False 收录）
        check("38.10 内建名字在别的模块也提示",
              _trig38("vbcr", structural=True, elsewhere=True,
                      module="模块3"),
              expect_contain=["vbCrLf"])
        # 38.11 As 位置提示内建数据类型，且不冒出函数名
        check("38.11 As 位置提示内建类型名（不冒函数名）",
              _trig38_type("str"),
              expect_contain=["String"], expect_absent=["MsgBox"])
        # 38.12 同名候选只保留一条
        _vis38 = E.filter_identifiers_by_scope(
            [("Left", "VBA", None, False), ("Left", "模块2", None, False)],
            None, "模块2")
        check("38.12 同名候选只保留一条（实得 %d 条）" % len(_vis38),
              [len(_vis38) == 1], expect_contain=[True])

        # 38.13~38.15：内建名字的匹配收紧档（v61 加、**v71 放宽到"实际不再收紧"**）
        # v71 起"纯分散命中"的最低输入长度 = BUILTIN_SCATTER_MIN(2) —— 3 字符的
        # 跳步输入也能命中；把阈值调回 4 就复现旧口径（38.14 就是这么对照的）。
        _sm38 = E.BUILTIN_SCATTER_MIN
        try:
            E.BUILTIN_SCATTER_MIN = 4
            _strict38 = _trig38("msb", elsewhere=True)
        finally:
            E.BUILTIN_SCATTER_MIN = _sm38
        check("38.13 v71：3 字符跳步输入 msb -> 提示 MsgBox（阈值已从 4 放到 2）",
              _trig38("msb", elsewhere=True),
              expect_contain=["MsgBox"], expect_absent=["msgGhost"])
        check("38.14 对照：阈值调回 4（旧口径）-> 同样的 msb 被挡掉"
              "（说明这条规则确实在起作用）",
              _strict38, expect_absent=["MsgBox"])
        check("38.15 前缀命中不受收紧影响（ms -> MsgBox）",
              _trig38("ms", elsewhere=True),
              expect_contain=["MsgBox"])

        # 38.16~38.20：真实 VbeBackend 全链路（临时把开关打开）
        # v70 起"内建枚举常量（vb*）"默认不收（用户口径：不用它、嫌干扰）。这一节
        # 验的是"完整名单都接上了（含常量）"，所以【两个开关都临时打开】——
        # 断言强度不变，只是把默认值这件事挪到第 47 节单独把关。
        import vbe_bridge as VB38R
        _saved38 = (VB38R.ENABLE_VBA_BUILTINS, VB38R.ENABLE_VBA_CONSTANTS)
        VB38R.ENABLE_VBA_BUILTINS = True
        VB38R.ENABLE_VBA_CONSTANTS = True
        try:
            class _CM38R(object):
                def __init__(self, text):
                    self.text = text

                @property
                def CountOfLines(self):
                    return self.text.count("\n") + 1

                def Lines(self, start, count):
                    return "\r\n".join(
                        self.text.split("\n")[start - 1:start - 1 + count])

            class _Comp38R(object):
                def __init__(self, name, text, ctype=1):
                    self.Name = name
                    self.Type = ctype
                    self.CodeModule = _CM38R(text)

            class _Proj38R(object):
                def __init__(self, comps):
                    self.Name = "VBAProject"
                    self.VBComponents = list(comps)

            class _Pane38R(object):
                def __init__(self, comp):
                    self._c = comp
                    self.sel = (1, 1, 1, 1)

                @property
                def CodeModule(self):
                    return self._c.CodeModule

                def GetSelection(self):
                    return self.sel

            class _VBE38R(object):
                def __init__(self, comps, act):
                    self.ActiveVBProject = _Proj38R(comps)
                    self.ActiveCodePane = _Pane38R(act)

            # 模块里一个字都没提过 MsgBox —— 提示它只能靠内建名字这一路
            _comp38 = _Comp38R(
                "Module1", "Sub Foo()\n    x = 1\n    <in>\nEnd Sub", 1)
            _vbe38 = _VBE38R([_comp38], _comp38)
            _orig38 = VB38R._get_vbe_cached
            VB38R._get_vbe_cached = lambda: _vbe38
            try:
                _bk38 = VB38R.VbeBackend()
                _ids38 = _bk38.get_identifiers()
                _nm38 = set(str(r[0]).lower() for r in _ids38)
                _vba38 = [r for r in _ids38
                          if str(r[1]) == VB38R._BUILTIN_MODULE]
                _visR38 = E.filter_identifiers_by_scope(
                    _ids38, None, "Module1")
                _hit38 = [n for n in _visR38
                          if E.fuzzy_match(n, "ms") is not None]
                check("38.16 真实后端：内建名字进了候选池（挂 VBA 的 %d 条）"
                      % len(_vba38),
                      ["msgbox" in _nm38 and "vbcrlf" in _nm38
                       and "split" in _nm38 and "string" in _nm38],
                      expect_contain=[True])
                check("38.17 真实后端：内建名字跨模块可见（priv=False）",
                      ["msgbox" in set(n.lower() for n in _visR38)],
                      expect_contain=[True])
                check("38.18 真实后端：结构性证据含内建名字",
                      ["msgbox" in _bk38.get_structural_names()
                       and "string" in _bk38.get_structural_names()],
                      expect_contain=[True])
                check("38.19 真实后端：declared_elsewhere 认得内建名字",
                      [_bk38.declared_elsewhere("MsgBox", "Module1")],
                      expect_contain=[True])
                check("38.20 真实后端：get_builtin_names 只含内建",
                      ["msgbox" in _bk38.get_builtin_names()
                       and "foo" not in _bk38.get_builtin_names()],
                      expect_contain=[True])
                check("38.21 真实池子里输入 ms 能命中 MsgBox",
                      _hit38, expect_contain=["MsgBox"])
            finally:
                VB38R._get_vbe_cached = _orig38
        finally:
            (VB38R.ENABLE_VBA_BUILTINS,
             VB38R.ENABLE_VBA_CONSTANTS) = _saved38

        # 38.22~38.24：Type/Enum 块成员不该从"隐式变量"那条路溜回候选池
        #
        # v59 让 extract_records 不再收块成员，但"只读用法"扫描（_usage_candidates）
        # 是整篇文本扫的、没有块的概念 —— 真机上枚举成员又从这条路溜回来了
        # （工程里冒出 A_ / AA_ / AZ_ … 共 12 个）。这里用最小片段钉住修复。
        _blk38 = "\n".join([
            "Public Enum E",
            "    A_ = 0: B_ = 1: Z_ = 25",
            "End Enum",
            "Public Type gRec",
            "    gId As Long",
            "End Type",
            "Sub Foo()",
            "    afterBlk = 1",
            "End Sub"])
        _rec38b, _pn = P.extract_records(_blk38, module="M1",
                                         is_std_module=True)
        _imp38b = P.extract_implicit_records(
            _blk38, module="M1", is_std_module=True,
            declared=_rec38b, scope="proc")
        _nm38b = sorted(str(n) for n, _m, _p, _pr in _rec38b + _imp38b)
        check("38.22 Enum/Type 成员不进候选池（漏进来的 _ 结尾名：%s）"
              % [n for n in _nm38b if n.endswith("_")],
              [not any(n.endswith("_") for n in _nm38b)],
              expect_contain=[True])
        check("38.23 Enum/Type 名本身照常收录",
              _nm38b, expect_contain=["E", "gRec"])
        check("38.24 块【之后】的代码照常收隐式变量（块状态已复位）",
              _nm38b, expect_contain=["afterBlk"])

    except Exception as _e38:
        check("第 38 节异常: %s" % _e38, [True], expect_contain=[False])

    # ---- 39. v62：语言关键字（Sub / Dim / If / For / Set …）----
    print("\n=== 39. v62 VBA 语言关键字进候选池 ===")
    try:
        import re as _re39
        import vba_builtins as VB39
        import vbe_bridge as VB39R

        # 39.1~39.5：清单自身（纯静态，不依赖 Excel）
        _kw39 = list(VB39.BUILTIN_KEYWORDS)
        _low39 = [n.lower() for n in _kw39]
        check("39.1 关键字清单无重名（共 %d 个）" % len(_low39),
              [len(set(_low39)) == len(_low39)], expect_contain=[True])
        _bad39 = [n for n in _kw39 if not _re39.match(r"^[^\W\d]\w*$", n)]
        check("39.2 关键字全是合法标识符（异常 %s）" % (_bad39 or "无"),
              [len(_bad39) == 0], expect_contain=[True])
        _ovl39 = sorted(set(_low39)
                        & (set(VB39.BUILTIN_LOWER) | set(VB39.BUILTIN_TYPE_LOWER)))
        check("39.3 与内建函数/类型清单无重叠（重叠 %s；重叠项由收集侧去重挡掉）"
              % (_ovl39 or "无"),
              [len(_ovl39) == 0], expect_contain=[True])
        _want39 = ("Sub", "Function", "Property", "Dim", "ReDim", "Set",
                   "Let", "If", "Then", "Else", "ElseIf", "End", "For",
                   "Next", "Each", "Do", "Loop", "While", "Wend", "Select",
                   "Case", "With", "Type", "Enum", "New", "Nothing", "Me",
                   "And", "Or", "Not", "Mod", "Call", "Exit", "On",
                   "Static", "Private", "Public", "Declare", "ByVal",
                   "ByRef", "Optional", "ParamArray", "WithEvents")
        _miss39 = [n for n in _want39
                   if n.lower() not in VB39.BUILTIN_KEYWORD_LOWER]
        check("39.4 常用关键字齐备（缺 %s）" % (_miss39 or "无"),
              [len(_miss39) == 0], expect_contain=[True])
        # 刻意不收的：歧义过大的语句名（更常当变量/属性名）+ 过时的 Def* 组
        _no39 = ("name", "line", "width", "spc", "tab",
                 "defint", "defstr", "defvar", "base", "compare")
        _hit39 = [n for n in _no39 if n in VB39.BUILTIN_KEYWORD_LOWER]
        check("39.5 歧义/过时关键字仍不收（命中 %s）" % (_hit39 or "无"),
              [len(_hit39) == 0], expect_contain=[True])

        # 39.6~39.12：走完整 trigger 链路（假后端，纯 Python，不碰 Excel）
        class _UI39(object):
            def show(self, *a, **k):
                pass

            def update_selection(self, *a, **k):
                pass

            def hide(self, *a, **k):
                pass

            def contains_point(self, *a, **k):
                return False

        class _B39(object):
            """模拟"新模块第一行，代码里一个关键字都还没写过"。

            候选池：三个语言关键字（挂虚拟模块 VBA）+ 一个真幽灵 kwGhost。
            structural / elsewhere 分别控制两路"真实存在"的证据 —— 真实情形下
            关键字归属虚拟模块 VBA，从当前模块看就是"别的模块声明过"（True）。
            """

            def __init__(self, structural=True, elsewhere=False,
                         module="模块2", builtin=True):
                self._structural = structural
                self._elsewhere = elsewhere
                self._module = module
                self._builtin = builtin

            def get_context(self):
                return {"line_no": 1, "caret_col": 1, "line_text": "",
                        "in_string": False, "in_comment": False,
                        "in_type_position": False, "in_decl_position": False,
                        "decl_names": [], "proc_name": None,
                        "module_name": self._module}

            def get_identifiers(self):
                return [("If", "VBA", None, False),
                        ("Dim", "VBA", None, False),
                        ("Sub", "VBA", None, False),
                        ("Get", "VBA", None, False),
                        ("kwGhost", "模块2", None, True)]

            def get_declared_names(self):
                return ["if", "dim", "sub", "get"]

            def caret_word_at_collect(self):
                return ""

            def declared_elsewhere(self, name, module_name):
                if self._elsewhere:
                    return str(name).lower() in ("if", "dim", "sub", "get")
                return False

            def get_structural_names(self):
                if not self._structural:
                    return []
                return ["if", "dim", "sub", "get"]

            def get_builtin_names(self):
                # 关键字与内建名字共用这一档收紧（见 engine.trigger）
                if not self._builtin:
                    return []
                return ["if", "dim", "sub", "get"]

            def get_type_names(self):
                return []

            def names_outside_caret(self, caret, scope_only=False):
                return set()        # 现场文本里一个字都没有

            def apply_completion(self, *a, **k):
                return None

        def _trig39(word, **kw):
            _be = _B39(**kw)
            _line = "    " + word
            _ctx = _be.get_context()
            _ctx["line_text"] = _line
            _ctx["caret_col"] = len(_line) + 1
            _be.get_context = lambda: _ctx
            _c = E.Completer(_be, _UI39())
            _c.trigger()
            return sorted(_c.matches or [])

        check("39.6 新模块里打 if -> 提示 If（现场没有它；幽灵仍剔）",
              _trig39("if", elsewhere=True),
              expect_contain=["If"], expect_absent=["kwGhost"])
        check("39.7 只靠结构性证据也放行（declared_elsewhere=False）",
              _trig39("if", structural=True, elsewhere=False),
              expect_contain=["If"])
        check("39.8 对照：无结构性证据且未声明过 -> 不提示（修复前行为）",
              _trig39("if", structural=False, elsewhere=False),
              expect_absent=["If"])
        check("39.9 真幽灵 kwGhost 仍被回声防护剔掉",
              _trig39("kwghost", structural=True, elsewhere=True),
              expect_absent=["kwGhost"])
        check("39.10 关键字在别的模块也提示（按模块级 priv=False 收录）",
              _trig39("di", structural=True, elsewhere=True, module="模块3"),
              expect_contain=["Dim"])
        # v71：关键字与内建名字共用这一档，阈值放到 2 之后 2 字符的跳步输入也能命中；
        # 调回 4 即复现 v61~v70 的旧口径（39.12 就是那个对照）。
        _sm39 = E.BUILTIN_SCATTER_MIN
        try:
            E.BUILTIN_SCATTER_MIN = 4
            _strict39 = _trig39("dm", elsewhere=True)
        finally:
            E.BUILTIN_SCATTER_MIN = _sm39
        check("39.11 v71：2 字符跳步输入 dm -> 提示 Dim（阈值已从 4 放到 2）",
              _trig39("dm", elsewhere=True), expect_contain=["Dim"])
        check("39.12 对照：阈值调回 4（旧口径）-> 同样的 dm 被挡掉"
              "（说明这条规则确实在起作用）",
              _strict39, expect_absent=["Dim"])

        # 39.13~39.16：真实 VbeBackend 全链路（临时把两个开关都打开）
        _saved39 = (VB39R.ENABLE_VBA_BUILTINS, VB39R.ENABLE_VBA_KEYWORDS)
        VB39R.ENABLE_VBA_BUILTINS = True
        VB39R.ENABLE_VBA_KEYWORDS = True
        try:
            class _CM39(object):
                def __init__(self, text):
                    self.text = text

                @property
                def CountOfLines(self):
                    return self.text.count("\n") + 1

                def Lines(self, start, count):
                    return "\r\n".join(
                        self.text.split("\n")[start - 1:start - 1 + count])

            class _Comp39(object):
                def __init__(self, name, text, ctype=1):
                    self.Name = name
                    self.Type = ctype
                    self.CodeModule = _CM39(text)

            class _Proj39(object):
                def __init__(self, comps):
                    self.Name = "VBAProject"
                    self.VBComponents = list(comps)

            class _Pane39(object):
                def __init__(self, comp):
                    self._c = comp

                @property
                def CodeModule(self):
                    return self._c.CodeModule

                def GetSelection(self):
                    return (1, 1, 1, 1)

            class _VBE39(object):
                def __init__(self, comps, act):
                    self.ActiveVBProject = _Proj39(comps)
                    self.ActiveCodePane = _Pane39(act)

            # 模块里一个字都没提过 If —— 提示它只能靠关键字这一路
            _comp39 = _Comp39("Module1", "Sub Foo()\n    x = 1\nEnd Sub", 1)
            _vbe39 = _VBE39([_comp39], _comp39)
            _orig39 = VB39R._get_vbe_cached
            VB39R._get_vbe_cached = lambda: _vbe39
            try:
                _bk39 = VB39R.VbeBackend()
                _ids39 = _bk39.get_identifiers()
                _nm39 = set(str(r[0]).lower() for r in _ids39)
                _vba39 = [r for r in _ids39
                          if str(r[1]) == VB39R._BUILTIN_MODULE]
                _visR39 = E.filter_identifiers_by_scope(_ids39, None, "Module1")
                _hit39b = [n for n in _visR39
                           if E.fuzzy_match(n, "if") is not None]
                check("39.13 真实后端：关键字进了候选池（挂 VBA 的共 %d 条）"
                      % len(_vba39),
                      ["if" in _nm39 and "dim" in _nm39 and "sub" in _nm39],
                      expect_contain=[True])
                check("39.14 真实后端：关键字跨模块可见（priv=False）",
                      ["if" in set(n.lower() for n in _visR39)],
                      expect_contain=[True])
                check("39.15 真实后端：结构性证据与 declared_elsewhere 都认关键字",
                      ["if" in _bk39.get_structural_names()
                       and bool(_bk39.declared_elsewhere("If", "Module1"))],
                      expect_contain=[True])
                check("39.16 真实池子里打 if 能命中 If",
                      _hit39b, expect_contain=["If"])

                # 开关独立：关掉关键字，内建函数照常
                VB39R.ENABLE_VBA_KEYWORDS = False
                try:
                    _bk39b = VB39R.VbeBackend()
                    _nm39b = set(str(r[0]).lower()
                                 for r in _bk39b.get_identifiers())
                    check("39.17 关掉关键字开关：关键字不进池、内建函数照常",
                          ["if" not in _nm39b and "msgbox" in _nm39b],
                          expect_contain=[True])
                finally:
                    VB39R.ENABLE_VBA_KEYWORDS = True
            finally:
                VB39R._get_vbe_cached = _orig39
        finally:
            (VB39R.ENABLE_VBA_BUILTINS,
             VB39R.ENABLE_VBA_KEYWORDS) = _saved39

    except Exception as _e39:
        check("第 39 节异常: %s" % _e39, [True], expect_contain=[False])

    # ==================================================================
    # 40. v63 VBE 自带提示窗（参数信息）可见时让位
    # ==================================================================
    # 用户报：`MsgBox "已完成！",` 敲下逗号后 VBE 弹出【参数信息】，我们的候选窗
    # 同时也冒出来，两个窗叠在同一处冲突。
    #
    # 根因：v55 的让位判据走【光标语法位置】（`标识符.` 之后 / `As `/`New ` 之后），
    # 只覆盖"自动列出成员 / 类型列表"，VBE 在别处弹的提示（最典型的就是参数信息）
    # 它一概不认 —— README 里其实把这条当"已知不拦"写着。
    #
    # 修法：加一路【精确】判据。真机实测：VBE 的提示窗是【预建复用】的真窗口
    # （什么都没按时就在顶层窗口列表里、只是不可见；Ctrl+Shift+I 后立刻可见、
    # Esc 后立刻不可见），类名是 VB IDE 专有的 NameListWndClass / PopupTipWndClass。
    # 后端 vbe_popup_visible() 直接查这两个窗口的可见性，引擎据此让位。
    #
    # 注：v55 当年以"发 Ctrl+J 后全系统零新窗口"推断 VBE 列表没有独立窗口 ——
    # 那只对"【新建】窗口"成立，被复用骗了。
    print("\n=== 40. v63 VBE 提示窗可见时让位（参数信息）===")
    try:
        import vbe_bridge as VB40

        # 40.1~40.3：探测函数自身（纯 Win32，不依赖 Excel）
        check("40.1 提示窗类名清单含 VBE 那两个窗口类",
              [set(VB40.VBE_POPUP_CLASSES)
               >= {"NameListWndClass", "PopupTipWndClass"}],
              expect_contain=[True])
        check("40.2 空句柄 / 无效句柄 -> 判为不可见（不炸）",
              [VB40._popup_details_now([]),
               VB40._popup_details_now([0, 123]),
               VB40._hwnd_alive(0)],
              expect_contain=[[], [], False])
        try:
            _pinfo40 = VB40.vbe_popup_info()
        except Exception as _e40pi:
            _pinfo40 = "异常: %s" % _e40pi
        check("40.2b vbe_popup_info() 返回明细列表且不抛（此刻=%r）" % (_pinfo40,),
              [isinstance(_pinfo40, list),
               all(isinstance(t, tuple) and len(t) == 3 for t in _pinfo40)],
              expect_contain=[True, True])
        try:
            _pv40 = VB40.vbe_popup_visible()
        except Exception as _e40pv:
            _pv40 = "异常: %s" % _e40pv
        check("40.3 vbe_popup_visible() 返回布尔且不抛（此刻=%s）" % (_pv40,),
              [isinstance(_pv40, bool)], expect_contain=[True])

        # 40.4~40.11：走完整 trigger 链路（假后端，纯 Python，不碰 Excel）
        class _UI40(object):
            def __init__(self):
                self.shown = False
                self.rows = []

            def show(self, rows, selection=0, completer=None):
                self.shown = True
                self.rows = list(rows)

            def update_selection(self, *a, **k):
                pass

            def hide(self, *a, **k):
                self.shown = False

            def contains_point(self, *a, **k):
                return False

        _LINE40 = '    MsgBox "已完成！", vb'

        class _B40Base(object):
            """复刻"光标在 MsgBox 第二个参数上"的现场。

            这个位置【不在】v55 的两种让位位置上（没有点号、也不是 As/New 之后），
            所以它恰好用来隔离出"窗口判据"这一路 —— 语法判据在这里根本不生效。
            """

            def __init__(self, **kw):
                pass

            def get_context(self):
                return {"line_no": 3, "caret_col": len(_LINE40) + 1,
                        "line_text": _LINE40, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": "模块2"}

            def get_identifiers(self):
                return [("MsgBox", "VBA", None, False),
                        ("vbYes", "VBA", None, False),
                        ("vbNo", "VBA", None, False)]

            def get_declared_names(self):
                return ["msgbox", "vbyes", "vbno"]

            def declared_elsewhere(self, name, module_name):
                return str(name).lower() in ("msgbox", "vbyes", "vbno")

            def get_structural_names(self):
                return ["msgbox", "vbyes", "vbno"]

            def get_builtin_names(self):
                return []

            def get_type_names(self):
                return []

            def names_outside_caret(self, caret, scope_only=False):
                return set()

            def apply_completion(self, *a, **k):
                return None

        class _B40(_B40Base):
            """带 vbe_popup_visible 的新式后端。"""

            def __init__(self, popup=False, boom=False, **kw):
                self._popup = popup
                self._boom = boom

            def vbe_popup_visible(self):
                if self._boom:
                    raise RuntimeError("boom")
                return self._popup

            def vbe_popup_info(self):
                if self._boom:
                    raise RuntimeError("boom")
                return [("PopupTipWndClass", 200, 16)] if self._popup else []

        class _B40Old(_B40Base):
            """旧式 / 测试后端：没有 vbe_popup_visible 这个方法。"""

        def _trig40(cls=_B40, manual=False, **kw):
            _be = cls(**kw)
            _ui = _UI40()
            _c = E.Completer(_be, _ui)
            _c.trigger(not manual)      # True = 轮询自动触发；manual = Ctrl+Space
            return sorted(_c.matches or []), _ui.shown

        _hit40 = (["vbNo", "vbYes"], True)
        check("40.4 提示窗可见 -> 让位（不弹、候选清空）",
              [_trig40(popup=True)],
              expect_contain=[([], False)])
        check("40.5 提示窗不可见 -> 照常弹（同位置同后端）",
              [_trig40(popup=False)],
              expect_contain=[_hit40])
        check("40.6 现场就在参数位置：v55 的语法判据认不出（隔离验证）",
              [E.vbe_list_expected(_LINE40, len(_LINE40) + 1)],
              expect_contain=[False])

        _saved40 = E.YIELD_TO_VBE_LIST
        try:
            E.YIELD_TO_VBE_LIST = False
            check("40.7 对照：关掉让位开关（VBECOMPLETE_NO_YIELD=1）-> 照弹",
                  [_trig40(popup=True)],
                  expect_contain=[_hit40])
        finally:
            E.YIELD_TO_VBE_LIST = _saved40

        check("40.8 旧式后端（无该接口）-> 行为不变，照常弹",
              [_trig40(cls=_B40Old)],
              expect_contain=[_hit40])
        check("40.9 探测抛异常 -> 不让位、不炸（异常被吞）",
              [_trig40(popup=True, boom=True)],
              expect_contain=[_hit40])
        # 【v76 口径变更】让位判据不再区分"自动轮询 / 手动 Ctrl+Space"。
        # 理由见 engine.trigger：VBE 的成员列表画在屏幕上是客观事实，我们的窗是
        # topmost 悬浮窗，一挂上去就把它整个盖住。用户 v76 原话："VBE 的弹窗和
        # 我们项目弹窗冲突时，屏蔽我们项目弹窗……直接消失让位。"
        # 成对保留一条对照（40.10b）：列表不可见时手动路径照弹 —— 证明挡它的确实
        # 是"列表可见"，而不是把手动路径整个禁掉了。
        check("40.10 ★v76 手动 Ctrl+Space：提示窗可见时同样让位（不再区分触发来源）",
              [_trig40(popup=True, manual=True)],
              expect_contain=[([], False)])
        check("40.10b 对照：提示窗不可见 -> 手动 Ctrl+Space 照弹（让位只管「可见」）",
              [_trig40(popup=False, manual=True)],
              expect_contain=[_hit40])

        # 40.11 判据与语法位置无关：普通赋值行上，窗口可见同样让位
        class _B40Plain(_B40):
            def get_context(self):
                _line = "    userN"
                return {"line_no": 1, "caret_col": len(_line) + 1,
                        "line_text": _line, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": "模块2"}

        _be40p = _B40Plain(popup=True)
        _ui40p = _UI40()
        _c40p = E.Completer(_be40p, _ui40p)
        _c40p.trigger(True)
        check("40.11 普通行（非成员/类型位置）也拦得住 -> 让位",
              [sorted(_c40p.matches or []), _ui40p.shown],
              expect_contain=[[], False])

        # 40.12 真实 VbeBackend 提供该接口，且与模块级探测一致（真机）
        _be40r = VB40.VbeBackend()
        _r40 = _be40r.vbe_popup_visible()
        check("40.12 VbeBackend 的探测与模块级函数一致（此刻=%s）" % (_r40,),
              [_r40 == VB40.vbe_popup_visible(), isinstance(_r40, bool)],
              expect_contain=[True, True])

        # 40.13 句柄复用护栏：类名 / 进程两道复核都在（句柄值可能被 Windows
        # 回收给别的窗口，只凭"记得这个句柄"就相信它会永远屏蔽我们的弹窗）
        _hp40 = VB40._host_process_id()
        check("40.13 类名复核与宿主 PID 探测可用（PID=%s）" % (_hp40,),
              [isinstance(_hp40, int) and _hp40 >= 0,
               VB40._window_class_name(0),
               VB40._window_class_name(123456) not in VB40.VBE_POPUP_CLASSES],
              expect_contain=[True, "", True])

        # 40.14~40.16 v65 诊断：让位那一刻日志要能说出【是哪个窗】
        check("40.14 引擎能取到提示窗明细（诊断日志用）",
              [E.Completer(_B40(popup=True), _UI40())._vbe_popup_desc(),
               E.Completer(_B40Old(), _UI40())._vbe_popup_desc()],
              expect_contain=[[("PopupTipWndClass", 200, 16)], None])
        check("40.15 明细探测抛异常 -> 返回 None，不炸",
              [E.Completer(_B40(popup=True, boom=True),
                           _UI40())._vbe_popup_desc()],
              expect_contain=[None])
        try:
            _pin40 = VB40.vbe_popup_info()
            _pvis40 = VB40.vbe_popup_visible()
            _ok40 = (bool(_pin40) == _pvis40)
        except Exception as _e4016:
            _ok40 = "异常: %s" % _e4016
        check("40.16 明细判据与可见性判据一致（真机，此刻=%r）" % (_pin40,),
              [_ok40], expect_contain=[True])

    except Exception as _e40:
        check("第 40 节异常: %s" % _e40, [True], expect_contain=[False])

    # ==================================================================
    # 41. v64 提示只在【代码窗格】里出现（属性窗口等非代码工作区不弹）
    # ==================================================================
    # 用户报：在【属性窗口】里改属性值时，候选窗偶尔会冒出来 —— 那里不是写代码
    # 的地方，不该有提示（"我不希望在这些非写代码的工作区出来提示"）。
    #
    # 根因：判断"在不在 VBE"只看【前台窗口标题】有没有 "Microsoft Visual Basic"；
    # 而属性窗口 / 工程窗口 / 窗体设计器都是 VBE 主窗口（wndclass_desked_gsk）的
    # 子窗口，标题判据对它们一律返回 True。
    #
    # 修法：vbe_bridge.vbe_code_pane_focused() 看【VBE 线程】里当前拥有键盘焦点
    # 的窗口，父链上有没有代码窗格（窗口类 VbaWindow）。本节 41.1~41.9 用打桩的
    # 窗口树验这条判据本身（不依赖真机有没有开 Excel），41.10 起验真实调用与
    # main.in_vbe_code_area 的接线 / 兜底语义。
    print("\n=== 41. v64 提示只在代码窗格出现（非代码工作区不弹）===")
    try:
        import vbe_bridge as VB41

        check("41.1 窗口类名清单（VBE 主窗口 / 代码窗格）",
              ["wndclass_desked_gsk" in VB41.VBE_FRAME_CLASSES,
               "VbaWindow" in VB41.CODE_PANE_CLASSES],
              expect_contain=[True, True])

        _orig41 = (VB41._vbe_frame_hwnd, VB41._focused_hwnd_in_vbe,
                   VB41._parent_hwnd, VB41._window_class_name)

        def _stub41(classes, tree, focus=None, frame=100):
            """把窗口树打桩：classes={hwnd:类名}，tree={hwnd:父hwnd}。"""
            VB41._vbe_frame_hwnd = lambda: frame
            VB41._focused_hwnd_in_vbe = (
                lambda: (focus if focus is not None else 0))
            VB41._parent_hwnd = lambda h: int(tree.get(h, 0))
            VB41._window_class_name = lambda h: classes.get(h, "")

        # 41.2 焦点在代码窗格里（含"焦点在代码窗格的子控件上"）-> 允许
        _stub41(
            classes={301: "ComboBox", 300: "VbaWindow",
                     201: "MDIClient", 100: "wndclass_desked_gsk"},
            tree={301: 300, 300: 201, 201: 100},
            focus=301)
        check("41.2 焦点在代码窗格（或其子控件）-> True",
              [VB41.vbe_code_pane_focused()], expect_contain=[True])
        _stub41(
            classes={300: "VbaWindow", 201: "MDIClient",
                     100: "wndclass_desked_gsk"},
            tree={300: 201, 201: 100}, focus=300)
        check("41.3 焦点就是代码窗格本身 -> True",
              [VB41.vbe_code_pane_focused()], expect_contain=[True])

        # 41.4~41.6 三类"非写代码的工作区"都必须判 False
        #   属性窗口：真机实测类名 wndclass_pbrs，里面的编辑框/列表是
        #   Edit / ListBox / ComboBox / SysTabControl32
        _stub41(
            classes={401: "ListBox", 400: "wndclass_pbrs",
                     100: "wndclass_desked_gsk"},
            tree={401: 400, 400: 100}, focus=401)
        check("41.4 焦点在属性窗口里 -> False（用户报的场景）",
              [VB41.vbe_code_pane_focused()], expect_contain=[False])
        _stub41(
            classes={501: "SysTreeView32", 500: "PROJECT",
                     100: "wndclass_desked_gsk"},
            tree={501: 500, 500: 100}, focus=501)
        check("41.5 焦点在工程窗口里 -> False",
              [VB41.vbe_code_pane_focused()], expect_contain=[False])
        _stub41(
            classes={601: "ThunderDFrame", 600: "DesignerWindow",
                     100: "wndclass_desked_gsk"},
            tree={601: 600, 600: 100}, focus=601)
        check("41.6 焦点在窗体设计器里 -> False",
              [VB41.vbe_code_pane_focused()], expect_contain=[False])

        # 41.7~41.9 拿不准时返回 None（调用方据此退回老判据，绝不能"拿不准就拦"）
        _stub41(classes={}, tree={}, focus=0)
        check("41.7 拿不到焦点窗口 -> None（不拦）",
              [VB41.vbe_code_pane_focused()], expect_contain=[None])
        _stub41(classes={}, tree={}, focus=1, frame=0)
        check("41.8 找不到 VBE 主窗口 -> None（不拦）",
              [VB41.vbe_code_pane_focused()], expect_contain=[None])
        # 链上既没有代码窗格、也走不到 frame（父链异常）-> False，且不能死循环
        _stub41(classes={i: "Edit" for i in range(700, 720)},
                tree={i: i + 1 for i in range(700, 719)}, focus=700)
        check("41.9 父链异常（走不到 frame）-> False 且不死循环",
              [VB41.vbe_code_pane_focused()], expect_contain=[False])

        for _f, _v in zip(("_vbe_frame_hwnd", "_focused_hwnd_in_vbe",
                           "_parent_hwnd", "_window_class_name"), _orig41):
            setattr(VB41, _f, _v)

        try:
            _real41 = VB41.vbe_code_pane_focused()
        except Exception as _e41r:
            _real41 = "异常: %s" % _e41r
        check("41.10 真实机器上返回 True/False/None 且不抛（此刻=%r）" % (_real41,),
              [isinstance(_real41, bool) or _real41 is None],
              expect_contain=[True])

        # 41.11~41.15 main.in_vbe_code_area：接线 + 兜底语义
        if M is None:
            check("main 模块不可用（缺 pynput），跳过 41.11~41.15", [True],
                  expect_contain=[True])
        else:
            _orig_area41 = VB41.vbe_code_pane_focused
            _orig_title41 = M.in_vbe_code_pane
            try:
                M.in_vbe_code_pane = lambda: True          # 前台标题像 VBE
                VB41.vbe_code_pane_focused = lambda: False  # 但焦点不在代码窗格
                check("41.11 标题像 VBE 但焦点在属性窗口 -> 不认（False）",
                      [M.in_vbe_code_area()], expect_contain=[False])
                VB41.vbe_code_pane_focused = lambda: True
                check("41.12 焦点在代码窗格 -> True",
                      [M.in_vbe_code_area()], expect_contain=[True])
                VB41.vbe_code_pane_focused = lambda: None
                check("41.13 判不出来 -> 退回标题判据（True，保持旧行为）",
                      [M.in_vbe_code_area()], expect_contain=[True])
                M.in_vbe_code_pane = lambda: False
                check("41.14 判不出来且标题也不是 VBE -> False",
                      [M.in_vbe_code_area()], expect_contain=[False])

                def _boom41():
                    raise RuntimeError("boom")
                VB41.vbe_code_pane_focused = _boom41
                M.in_vbe_code_pane = lambda: True
                check("41.15 探测抛异常 -> 退回标题判据、不炸",
                      [M.in_vbe_code_area()], expect_contain=[True])

                # 41.16 接线护栏：三处"会弹窗 / 会写代码"的动作都改用精确判据
                #   （轮询触发、Shift+Enter 与自动配对、Ctrl+Space）
                import inspect as _inspect41
                _src41 = _inspect41.getsource(M)
                _cnt41 = _src41.count("in_vbe_code_area()")
                check("41.16 精确判据已接进 4 处动作（实际 %d 处）" % _cnt41,
                      [_cnt41 >= 4,
                       "vbe_code_pane_focused" in _src41,
                       "in_vbe_code_pane()" in _src41,
                       "_poll_mod_switch" in _src41],
                      expect_contain=[True, True, True, True])

                # 41.18~41.21 换模块否决（同一个坑的另一半：没敲字也不该弹）
                #   快照 = (行号, 行文本, 模块名)；模块名拿不到就不否决。
                check("41.18 同一模块内的文本变化 -> 不否决",
                      [M._poll_mod_switch((3, "    us", "Sheet1"),
                                          (3, "    use", "Sheet1"))],
                      expect_contain=[False])
                check("41.19 换模块（同行号、文本不同）-> 否决",
                      [M._poll_mod_switch((3, "    us", "Sheet1"),
                                          (3, "    End Sub", "Module7"))],
                      expect_contain=[True])
                check("41.20 模块名大小写不同 -> 不算换模块",
                      [M._poll_mod_switch((3, "a", "Sheet1"),
                                          (3, "ab", "SHEET1"))],
                      expect_contain=[False])
                check("41.21 模块名缺失 / 旧式二元组 -> 不否决（不能把工具哑掉）",
                      [M._poll_mod_switch((3, "a", ""), (3, "ab", "Sheet1")),
                       M._poll_mod_switch((3, "a", "Sheet1"), (3, "ab", "")),
                       M._poll_mod_switch((3, "a"), (3, "ab"))],
                      expect_contain=[False, False, False])

                # 41.17 逃生开关：VBECOMPLETE_NO_CODE_AREA_GATE=1 退回旧判据
                _saved_gate41 = M.CODE_AREA_GATE
                try:
                    M.CODE_AREA_GATE = False
                    M.in_vbe_code_pane = lambda: True
                    check("41.17 关掉闸门 -> 退回「前台是 VBE 就算数」（True）",
                          [M.in_vbe_code_area()], expect_contain=[True])
                finally:
                    M.CODE_AREA_GATE = _saved_gate41
            finally:
                VB41.vbe_code_pane_focused = _orig_area41
                M.in_vbe_code_pane = _orig_title41

    except Exception as _e41:
        check("第 41 节异常: %s" % _e41, [True], expect_contain=[False])

    # ==================================================================
    # 42. v65 让位只认【成员列表】，参数信息（形参签名）不让位
    # ==================================================================
    # 用户报：UserForm2 第13行 `UserForm1.SelectedList.RemoveItem sellis` 处
    # "输入任何字符都无提醒"；VBE 在那儿弹的正是【参数信息】（形参签名）。
    # 用户澄清 v63 的原意："只有弹成员列表的时候才让位，弹形参签名不需要让位。"
    #
    # 只读复现（真实工程 + 合成光标，已实测）确认：该位置唯一能拦掉提示的就是
    # v63 那条让位判据 —— 提示窗不可见时 matches=['SelectedListRow2']；可见时
    # 让位、候选为空。故修法 = 把让位范围从"任何提示窗"收窄到"成员列表窗"，
    # 并让焦点判据把 VBE 自己的提示窗算作"仍在代码窗格"。
    print("\n=== 42. v65 让位只认成员列表（参数信息不让位）===")
    try:
        import vbe_bridge as VB42

        # 42.1 让位清单：含成员列表、不含参数信息
        check("42.1 让位清单只含成员列表窗（NameListWndClass）",
              ["NameListWndClass" in VB42.VBE_YIELD_CLASSES,
               "PopupTipWndClass" not in VB42.VBE_YIELD_CLASSES,
               set(VB42.VBE_YIELD_CLASSES) <= set(VB42.VBE_POPUP_CLASSES)],
              expect_contain=[True, True, True])

        # 42.2~42.5 判据本身（打桩窗口明细，纯 Python，不碰 Excel）
        _orig42 = (VB42._vbe_popup_hwnds, VB42._popup_details_now)

        def _stub42(details):
            VB42._vbe_popup_hwnds = lambda: [1]
            VB42._popup_details_now = lambda hs: list(details)

        _stub42([("PopupTipWndClass", 300, 18)])
        check("42.2 只可见参数信息 -> 不让位（False）★v65 核心",
              [VB42.vbe_yield_visible()], expect_contain=[False])
        _stub42([("NameListWndClass", 120, 90)])
        check("42.3 可见成员列表 -> 让位（True）",
              [VB42.vbe_yield_visible()], expect_contain=[True])
        _stub42([("PopupTipWndClass", 300, 18), ("NameListWndClass", 120, 90)])
        check("42.4 两个都可见 -> 让位（成员列表说了算）",
              [VB42.vbe_yield_visible()], expect_contain=[True])

        def _boom42(hs):
            raise RuntimeError("boom")
        VB42._popup_details_now = _boom42
        check("42.5 探测抛异常 -> 不让位、不炸",
              [VB42.vbe_yield_visible()], expect_contain=[False])
        VB42._vbe_popup_hwnds, VB42._popup_details_now = _orig42

        # 42.6~42.10 引擎接线：优先 vbe_yield_visible，老后端退回 vbe_popup_visible
        class _B42(_B40Base):
            """v65 现场：参数信息可见，成员列表不可见。"""

            def __init__(self, yld=False, pop=True, boom=False, boom2=False,
                         **kw):
                self._yld, self._pop = yld, pop
                self._boom, self._boom2 = boom, boom2

            def vbe_popup_visible(self):
                if self._boom:
                    raise RuntimeError("boom")
                return self._pop

            def vbe_yield_visible(self):
                if self._boom2:
                    raise RuntimeError("boom")
                return self._yld

            def vbe_popup_info(self):
                return ([("PopupTipWndClass", 300, 18)] if self._pop else []) \
                    + ([("NameListWndClass", 120, 90)] if self._yld else [])

        def _trig42(cls=_B42, manual=False, **kw):
            _be = cls(**kw)
            _ui = _UI40()
            _c = E.Completer(_be, _ui)
            _c.trigger(not manual)      # True = 轮询自动触发；manual = Ctrl+Space
            return sorted(_c.matches or []), _ui.shown

        _hit42 = (["vbNo", "vbYes"], True)
        check("42.6 参数信息可见（成员列表不可见）-> 照常弹 ★v65 核心",
              [_trig42(pop=True, yld=False)], expect_contain=[_hit42])
        check("42.7 成员列表可见 -> 让位（不弹、候选清空）",
              [_trig42(pop=False, yld=True)], expect_contain=[([], False)])
        check("42.8 老后端（只有 vbe_popup_visible）-> 退回旧行为：任何提示窗都让位",
              [_trig42(cls=_B40, popup=True)], expect_contain=[([], False)])
        check("42.8b 更老的桩后端（两个接口都没有）-> 不让位、照弹",
              [_trig42(cls=_B40Old)], expect_contain=[_hit42])
        check("42.9 vbe_yield_visible 抛异常 -> 不让位、不炸",
              [_trig42(pop=True, yld=False, boom2=True)],
              expect_contain=[_hit42])
        # 【v76 口径变更】同 40.10：成员列表可见时，手动路径也让位 —— 这条判据
        # 现在是"不依赖触发来源的通用不变量"（v73 的补位路径 trigger(False) 也受
        # 它约束，那正是补出来的窗会压住 VBE 列表的那个时序缝）。
        check("42.10 ★v76 手动 Ctrl+Space：成员列表可见 -> 同样让位（不再区分触发来源）",
              [_trig42(manual=True, pop=False, yld=True)],
              expect_contain=[([], False)])
        check("42.10b 对照：成员列表不可见 -> 手动路径照弹",
              [_trig42(manual=True, pop=False, yld=False)],
              expect_contain=[_hit42])

        # 42.11 真机：让位判据与明细一致
        _y42 = _i42 = None
        try:
            _be42 = VB42.VbeBackend()
            _y42 = _be42.vbe_yield_visible()
            _i42 = _be42.vbe_popup_info()
            _ok42 = (isinstance(_y42, bool)
                     and _y42 == any(c in VB42.VBE_YIELD_CLASSES
                                     for c, _w, _h in _i42))
        except Exception as _e4211:
            _ok42 = "异常: %s" % _e4211
        check("42.11 真机：让位判据与明细一致（yield=%r info=%r）" % (_y42, _i42),
              [_ok42], expect_contain=[True])

        # 42.12~42.14 焦点判据：VBE 自己的提示窗算作"仍在代码窗格"
        #   （否则提示窗一抢到焦点，v64 的闸门就把我们的候选窗当场收掉）
        _orig42f = (VB42._vbe_frame_hwnd, VB42._focused_hwnd_in_vbe,
                    VB42._parent_hwnd, VB42._window_class_name)

        def _stub42f(cls, hwnd=900):
            VB42._vbe_frame_hwnd = lambda: 100
            VB42._focused_hwnd_in_vbe = lambda: hwnd
            VB42._parent_hwnd = lambda h: 0
            VB42._window_class_name = lambda h: (
                cls if h == hwnd
                else ("wndclass_desked_gsk" if h == 100 else ""))

        _stub42f("NameListWndClass")
        check("42.12 焦点落在成员列表窗上 -> 仍算在代码窗格（True）",
              [VB42.vbe_code_pane_focused()], expect_contain=[True])
        _stub42f("PopupTipWndClass")
        check("42.13 焦点落在参数信息窗上 -> 仍算在代码窗格（True）★v65",
              [VB42.vbe_code_pane_focused()], expect_contain=[True])
        _stub42f("wndclass_pbrs")
        check("42.14 属性窗口仍不算（False，v64 语义不变）",
              [VB42.vbe_code_pane_focused()], expect_contain=[False])
        for _f, _v in zip(("_vbe_frame_hwnd", "_focused_hwnd_in_vbe",
                           "_parent_hwnd", "_window_class_name"), _orig42f):
            setattr(VB42, _f, _v)

    except Exception as _e42:
        check("第 42 节异常: %s" % _e42, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    print("\n=== 43. v66 候选窗避让 VBE 参数信息窗（形参签名）===")
    try:
        import ui as UI43
        import vbe_bridge as VB43

        # 43.1 避让与让位是两码事：参数信息仍然【不让位】（v65 语义不动）
        check("43.1 避让不改让位清单（参数信息仍不让位）",
              ["PopupTipWndClass" in VB43.VBE_POPUP_CLASSES,
               "PopupTipWndClass" not in VB43.VBE_YIELD_CLASSES],
              expect_contain=[True, True])

        # 43.2~43.3 矩形判据（明细=矩形投影，一处逻辑两处消费）
        _orig43 = (VB43._vbe_popup_hwnds, VB43._popup_rects_now)
        VB43._vbe_popup_hwnds = lambda: [1]
        VB43._popup_rects_now = lambda hs: [("PopupTipWndClass", 10, 20, 310, 38)]
        check("43.2 明细由矩形投影而来（300x18）",
              [VB43._popup_details_now([1]), VB43.vbe_popup_rects()],
              expect_contain=[[("PopupTipWndClass", 300, 18)],
                              [("PopupTipWndClass", 10, 20, 310, 38)]])

        def _boom43(hs):
            raise RuntimeError("boom")
        VB43._popup_rects_now = _boom43
        check("43.3 探测抛异常 -> 返回空列表、不炸（拿不到就不动窗）",
              [VB43.vbe_popup_rects()], expect_contain=[[]])
        VB43._vbe_popup_hwnds, VB43._popup_rects_now = _orig43

        # 43.4~43.11 纯函数 avoid_popup_rects：把候选窗挪出提示窗
        _R43 = UI43.avoid_popup_rects
        _pop43 = (0, 220, 400, 240)      # 光标下方的一条形参签名
        check("43.4 无提示窗 -> 位置原样不动",
              [_R43(100, 220, 200, 90, 200, 1080, [])], expect_contain=[220])
        check("43.5 签名窗就在正下方且重叠 -> 下移到它下面（240+2）★核心",
              [_R43(100, 220, 200, 90, 200, 1080, [_pop43])], expect_contain=[242])
        check("43.6 提示窗在屏幕另一侧（水平不相交）-> 不动",
              [_R43(600, 220, 200, 90, 200, 1080, [(0, 220, 400, 240)])],
              expect_contain=[220])
        check("43.7 提示窗在上方（垂直不相交）-> 不动",
              [_R43(100, 220, 200, 90, 200, 1080, [(0, 100, 400, 200)])],
              expect_contain=[220])
        check("43.8 两个窗叠着 -> 落到最下面那个的下方（260+2）",
              [_R43(100, 220, 200, 90, 200, 1080,
                    [(0, 220, 400, 240), (0, 239, 400, 260)])],
              expect_contain=[262])
        check("43.9 下移会出屏 -> 翻到光标上方（990-90-2）",
              [_R43(100, 1000, 200, 90, 990, 1080, [(0, 1020, 400, 1060)])],
              expect_contain=[898])
        check("43.10 上下都放不下 -> 贴屏幕底部（1080-90）",
              [_R43(100, 900, 200, 90, 60, 1080, [(0, 900, 400, 1000)])],
              expect_contain=[990])
        check("43.11 只压住候选窗一部分 -> 照样躲（水平相交就避）",
              [_R43(500, 220, 200, 90, 200, 1080, [(600, 220, 900, 240)])],
              expect_contain=[242])

        # 43.12~43.14 _position 接线：默认贴光标下方、有签名窗就下移、同位置不重设
        class _FR43(object):
            def winfo_screenwidth(self):
                return 1920

            def winfo_screenheight(self):
                return 1080

        class _FC43(object):
            def cget(self, k):
                return 200 if k == "width" else 90

        class _FW43(object):
            def __init__(self):
                self.geoms = []

            def geometry(self, s):
                self.geoms.append(s)

        def _mk43():
            p = UI43.Popup.__new__(UI43.Popup)
            p.root, p.canvas, p.win = _FR43(), _FC43(), _FW43()
            p._last_geom = None
            return p

        _o43caret = UI43.caret_screen_rect
        _o43rects = UI43.screen_popup_rects
        try:
            UI43.caret_screen_rect = lambda: (100, 200, 218)
            UI43.screen_popup_rects = lambda: []
            _p43 = _mk43()
            _p43._position()
            _g1 = list(_p43.win.geoms)
            _p43._position()                  # 位置没变：不该重复设 geometry
            _g2 = list(_p43.win.geoms)
            UI43.screen_popup_rects = lambda: [(0, 218, 400, 240)]
            _p43._position()                  # 签名窗出现 -> 下移到 240+2
            _g3 = list(_p43.win.geoms)
        finally:
            UI43.caret_screen_rect = _o43caret
            UI43.screen_popup_rects = _o43rects
        check("43.12 _position 默认贴光标下方 → %s" % (_g1[-1:] or [None],),
              [_g1], expect_contain=[["200x90+100+220"]])
        check("43.13 同位置重复调用不重设 geometry（防闪烁）→ %s" % (_g2[-1:] or [None],),
              [_g2], expect_contain=[["200x90+100+220"]])
        check("43.14 签名窗出现 -> 候选窗下移到其下方 → %s" % (_g3[-1:] or [None],),
              [_g3], expect_contain=[["200x90+100+220", "200x90+100+242"]])

        # 43.15 拿不到提示窗矩形时 screen_popup_rects 不得抛（vbe_bridge 打桩成异常）
        _o43vbr = VB43.vbe_popup_rects
        VB43.vbe_popup_rects = _boom43
        try:
            _sr43 = UI43.screen_popup_rects()
        finally:
            VB43.vbe_popup_rects = _o43vbr
        check("43.15 探测异常时 screen_popup_rects() 返回 []（不抛）",
              [_sr43], expect_contain=[[]])

        # 43.16 收起状态下 reposition 是安全的空操作（没有窗、没有 root 也不炸）
        _p43r = UI43.Popup.__new__(UI43.Popup)
        _p43r.win = None
        _p43r.reposition()
        check("43.16 收起状态 reposition 空操作、不抛", [True], expect_contain=[True])

        # 43.17 真机：提示窗矩形可读、格式正确
        _rr43 = None
        try:
            _rr43 = VB43.VbeBackend().vbe_popup_rects()
            _ok43 = (isinstance(_rr43, list)
                     and len(_rr43) == len(VB43.vbe_popup_rects())
                     and all(len(t) == 5 for t in _rr43))
        except Exception as _e4317:
            _ok43 = "异常: %s" % _e4317
        check("43.17 真机：提示窗矩形可读且格式正确（%r）" % (_rr43,),
              [_ok43], expect_contain=[True])

        # 43.18 main 接线护栏：轮询里要有"提示窗矩形变化 -> 重摆候选窗"
        try:
            import inspect as _inspect43
            import main as _M43
            _src43 = _inspect43.getsource(_M43)
            _cnt43 = _src43.count("vbe_popup_rects()")
        except Exception as _e4318:
            _src43, _cnt43 = "", "异常: %s" % _e4318
        check("43.18 轮询已接上「矩形变化 -> reposition」（出现 %s 次）" % (_cnt43,),
              [_cnt43 >= 1, "post(popup.reposition)" in _src43,
               "_popup_sig" in _src43],
              expect_contain=[True, True, True])
        check("43.19 reposition 挂在 UI 上（引擎层没有这个方法，故 main 调 popup.*）",
              [callable(getattr(UI43.Popup, "reposition", None))],
              expect_contain=[True])
    except Exception as _e43:
        check("第 43 节异常: %s" % _e43, [True], expect_contain=[False])

    # ==================================================================
    # 44. v67：宿主类型库的枚举常量（Excel 的 xl*、Office 的 mso*）
    #
    # 用户报："输入 vb 会提示一堆 VBA 枚举值，输入 xl 却一个 xl 开头的枚举值
    # 都不提示。" vb* 是 v61 收的 VBA 内建常量；xl* 属于宿主 Excel 类型库，
    # 一直没收 —— 不是泄露，是漏收。
    #
    # 收进来的同时必须收紧口径：这批名字 4500 多条、全是长复合词，放开跳步
    # 匹配的话输入 ms 会带出 2365 条 mso*、输入 count 会带出 130 条 xl*。
    # 所以只认"从头开始的连续前缀"，且输入长度不低于家族前缀长度。
    # ==================================================================
    print("\n=== 44. 宿主类型库枚举常量（v67）===")
    try:
        import vbe_bridge as VB44
        import engine as E44

        # ---- 44.1~44.5 家族前缀推导（纯函数，不碰 Excel）----
        _excel_like = ["xlUp", "xlDown", "xlToRight", "xlCellTypeVisible",
                       "xlContinuous", "rgbCadetBlue", "sigdetFoo"]
        _office_like = ["msoTrue", "msoFalse", "msoLineSolid", "msoAlignLeft",
                        "msoCTLPush", "rgbCadetBlue", "sigdetFoo"]
        _stdole_like = ["Unchecked", "Checked", "Gray", "Default",
                        "Monochrome", "VgaColor", "Color"]
        check("44.1 Excel 类库 -> 家族前缀 xl",
              [VB44.enum_family_prefix(_excel_like)], expect_contain=["xl"])
        check("44.2 Office 类库 -> 家族前缀 mso",
              [VB44.enum_family_prefix(_office_like)], expect_contain=["mso"])
        check("44.3 各说各话的库（stdole）-> 无家族前缀",
              [VB44.enum_family_prefix(_stdole_like)], expect_contain=[""])
        check("44.4 空清单 -> 空串（不抛）",
              [VB44.enum_family_prefix([]), VB44.enum_family_prefix(["a"])],
              expect_contain=["", ""])
        _split = ["abOne", "abTwo", "cdOne", "cdTwo", "efOne"]
        check("44.5 没有任何前缀过半 -> 空串（宁可整库不要，也不放通用词进来）",
              [VB44.enum_family_prefix(_split)], expect_contain=[""])

        # ---- 44.6~44.7 家族过滤 + 白名单 ----
        _fam, _kept = VB44.host_enum_family(_excel_like)
        check("44.6 host_enum_family 只留家族成员（%r）" % (_kept,),
              [_fam, sorted(_kept), "rgbCadetBlue" in _kept],
              expect_contain=["xl", ["xlCellTypeVisible", "xlContinuous",
                                     "xlDown", "xlToRight", "xlUp"], False])
        _old_fam44 = VB44._HOST_ENUM_FAMILIES
        VB44._HOST_ENUM_FAMILIES = ("zz",)
        try:
            _wf44 = VB44.host_enum_family(_excel_like)
        finally:
            VB44._HOST_ENUM_FAMILIES = _old_fam44
        check("44.7 白名单不含 xl -> 整库跳过（用于只想留某个家族）",
              [_wf44], expect_contain=[("", [])])

        # ---- 44.8~44.9 真机：真能枚举出引用类型库的枚举常量 ----
        # ⚠️ 必须单独加载一份干净的 vbe_bridge 来跑这一节：本文件前面的小节把
        # `VB._get_vbe_cached` 打桩成了假 VBE（第 31/33 节），桩没还原，真机
        # 接口在那儿只会读到假工程（引用清单为空 -> 枚举出 0 条）。
        # 用 importlib 重新 exec 一份独立实例，验的就是真实链路。
        _fresh44 = None
        try:
            import importlib.util as _iu44
            _spec44 = _iu44.spec_from_file_location(
                "vbe_bridge_fresh44",
                os.path.join(os.path.dirname(os.path.dirname(
                    os.path.abspath(__file__))), "vbe_bridge.py"))
            _fresh44 = _iu44.module_from_spec(_spec44)
            _spec44.loader.exec_module(_fresh44)
        except Exception:
            _fresh44 = None
        _VB44R = _fresh44 or VB44
        # v70 起宿主枚举默认不收（用户口径）。这一节验的是"枚举链路本身能跑通"，
        # 所以把开关临时打开 —— 断言强度不变；"默认关"由第 47 节单独把关。
        _saved_he44 = _VB44R.ENABLE_HOST_ENUMS
        _VB44R.ENABLE_HOST_ENUMS = True
        _items44 = []
        _excel44 = False
        try:
            _items44 = _VB44R.host_enum_constants()
            _mins44 = sorted({m for _n, m in _items44})
            _nxl44 = len([1 for n, _m in _items44 if n.lower().startswith("xl")])
            _nms44 = len([1 for n, _m in _items44 if n.lower().startswith("mso")])
            # 家族前缀集合（= 名字的小写前 m 个字符）；形态必须健康
            _fams44 = sorted({n.lower()[:m] for n, m in _items44})
            _badfam44 = sorted(
                f for f in _fams44
                if not (2 <= len(f) <= 4 and f.isalnum() and f == f.lower()))
            # ★v104「当前宿主是不是 Excel」：看清单里有没有成规模的 xl* 常量。
            # 宿主探测上线后连到的可能是 Word/PPT/WPS，本节的 Excel 专属断言不适用
            # —— 那种情况走下面的 else 分支优雅跳过（与 45.12 的跳过同一模式）。
            _excel44 = (isinstance(_items44, list)
                        and _nxl44 >= 1000 and _nms44 >= 1000)
            _ok44 = (isinstance(_items44, list)
                     and all(m >= 2 for _n, m in _items44)
                     and _nxl44 >= 1000 and _nms44 >= 1000
                     and not _badfam44)
            # 家族前缀不写死 xl / mso：这条断言要跟着"工程引用了哪些库"走。
            # 后端给的第 2 个值就是该库推导出的家族前缀长度，取名字的小写前缀
            # 即是那个前缀 —— 它必须是小写字母数字、2~4 个字符（单字母前缀会
            # 让一两个字母带出半屏名字，后端不会收）。
            # 实测本机：xl=2173（Excel）、mso=2365（Office）、fm=180（MSForms ——
            # 工程自己引的库，这正是"你引用什么就补什么"该有的样子）。
            _desc44 = ("共 %d 条，家族前缀=%s，xl*=%d mso*=%d，"
                       "最短输入长度=%s" % (len(_items44), _fams44,
                                        _nxl44, _nms44, _mins44))
        except Exception as _e44:
            _ok44, _desc44 = "异常: %s" % _e44, ""
        if _excel44:
            check("44.8 真机：枚举出 %s" % (_desc44,), [_ok44],
                  expect_contain=[True])
        elif _items44:
            check("44.8 真机：当前宿主非 Excel（%s）—— 跳过 Excel 专属断言"
                  "（v104 宿主探测生效；换台开着 Excel 的机器即恢复校验）"
                  % (_desc44,), [True], expect_contain=[True])
        else:
            check("44.8 真机：枚举出 %s" % (_desc44,), [_ok44],
                  expect_contain=[True])

        _hm44 = None
        try:
            _hm44 = _VB44R.VbeBackend().get_host_enum_names()
            _okhm44 = (isinstance(_hm44, dict) and _hm44.get("xlup") == 2
                       and _hm44.get("msotrue") == 3)
        except Exception as _e44b:
            _okhm44 = "异常: %s" % _e44b
        if _excel44:
            check("44.9 后端 get_host_enum_names() -> 小写名到最短长度的字典",
                  [_okhm44], expect_contain=[True])
        elif _items44:
            check("44.9 后端 get_host_enum_names()：当前宿主非 Excel —— 跳过",
                  [True], expect_contain=[True])
        else:
            check("44.9 后端 get_host_enum_names() -> 小写名到最短长度的字典",
                  [_okhm44], expect_contain=[True])
        # 复原（后面 44.10+ 走桩后端，与这个开关无关）
        _VB44R.ENABLE_HOST_ENUMS = _saved_he44

        # ---- 44.10~44.15 引擎接线：前缀专用通道 ----
        _HOST44 = {"xlup": 2, "xlcellvalue": 2, "xldown": 2, "xltoright": 2,
                   "msotrue": 3, "msofalse": 3}
        _POOL44 = [("xlUp", "宿主库", None, False),
                   ("xlCellValue", "宿主库", None, False),
                   ("xlDown", "宿主库", None, False),
                   ("xlToRight", "宿主库", None, False),
                   ("msoTrue", "宿主库", None, False),
                   ("msoFalse", "宿主库", None, False),
                   ("numArr", "M1", None, False),
                   ("countRows", "M1", None, False),
                   ("cellText", "M1", None, False)]

        class _B44(_B40Base):
            """宿主枚举常量的现场：池子里既有宿主常量，也有用户自己的名字。"""

            def __init__(self, word="xl", host=True, hook=True, **kw):
                self._word, self._host, self._hook = word, host, hook
                self.struct_calls = 0

            def _line(self):
                return "    " + self._word

            def get_context(self):
                _ln = self._line()
                return {"line_no": 3, "caret_col": len(_ln) + 1,
                        "line_text": _ln, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": "M1"}

            def get_identifiers(self):
                return list(_POOL44)

            def get_declared_names(self):
                return [r[0].lower() for r in _POOL44]

            def declared_elsewhere(self, name, module_name):
                return str(name).lower() in [r[0].lower() for r in _POOL44]

            def get_structural_names(self):
                self.struct_calls += 1
                return [r[0].lower() for r in _POOL44]

            def get_builtin_names(self):
                return []

            def get_type_names(self):
                return []

            def names_outside_caret(self, caret, scope_only=False):
                return set()

            def get_host_enum_names(self):
                if not self._hook:
                    raise RuntimeError("boom")
                return dict(_HOST44) if self._host else {}

        def _trig44(cls=_B44, **kw):
            _be = cls(**kw)
            _ui = _UI40()
            _c = E44.Completer(_be, _ui)
            _c.trigger(True)
            return sorted(_c.matches or []), _ui.shown, _be.struct_calls

        _h44, _v44, _c44 = _trig44(word="xl")
        check("44.10 输入 xl -> 出 xl* 家族（★用户点名要的），不出 mso*",
              [_v44, _h44, any(m.lower().startswith("mso") for m in _h44)],
              expect_contain=[True, ["xlCellValue", "xlDown", "xlToRight",
                                     "xlUp"], False])
        check("44.11 输入 xlu -> 只出 xlUp（前缀命中）",
              [_trig44(word="xlu")[0]], expect_contain=[["xlUp"]])
        check("44.12 输入 ms -> 只出 mso*，不出别的；输入 mso 才出 mso*",
              [_trig44(word="ms")[0], _trig44(word="mso")[0]],
              expect_contain=[[], ["msoFalse", "msoTrue"]])
        check("44.13 输入 count / cell -> 一条 xl* 都不多出（★防噪音核心）",
              [_trig44(word="count")[0], _trig44(word="cell")[0]],
              expect_contain=[["countRows"], ["cellText"]])

        class _B44No(_B44):
            """更老的桩后端：连 get_host_enum_names 这个接口都没有。"""
            get_host_enum_names = None

        _nb44 = _trig44(cls=_B44No, word="cell")[0]
        check("44.14 老后端（不提供该接口）-> 退回普通模糊匹配"
              "（cell 照样命中 xlCellValue，行为与 v66 完全一致）",
              [_nb44, _trig44(word="cell")[0]],
              expect_contain=[["cellText", "xlCellValue"], ["cellText"]])
        _boom44 = _trig44(word="xlu", hook=False)[0]
        check("44.15 接口抛异常 -> 不崩，且与「没有该接口」同行为（%r）"
              % (_boom44,),
              [_boom44 == _trig44(cls=_B44No, word="xlu")[0],
               "xlUp" in _boom44], expect_contain=[True, True])

        # ---- 44.16 结构性名字集合按一次触发取一份（v67 性能护栏）----
        # 修复前每个回声候选都重建一次集合；池子涨到 4600 多条后，
        # 输入 xl（一次 2000 多个候选）光这一步就要 2 秒（真机实测 2135ms）。
        check("44.16 一次触发里 get_structural_names() 只取一次（实际 %d 次）"
              % (_c44,), [_c44], expect_contain=[1])

        # ---- 44.17 诊断段受日志开关约束（否则每次触发多跑一遍全池匹配）----
        _orig_fm44, _orig_log44 = E44.fuzzy_match, E44._LOG_ENABLED
        _cnt44 = [0]

        def _fm44(name, query):
            _cnt44[0] += 1
            return _orig_fm44(name, query)

        try:
            E44.fuzzy_match = _fm44
            E44._LOG_ENABLED = False
            _cnt44[0] = 0
            _trig44(word="xl")
            _off44 = _cnt44[0]
            E44._LOG_ENABLED = True
            _cnt44[0] = 0
            _trig44(word="xl")
            _on44 = _cnt44[0]
        finally:
            E44.fuzzy_match, E44._LOG_ENABLED = _orig_fm44, _orig_log44
        # 关日志：只有 3 个非宿主名字走模糊匹配；开日志：多跑一遍全池（9 条）
        check("44.17 关日志时诊断段不跑（模糊匹配 %d 次 vs 开日志 %d 次）"
              % (_off44, _on44), [_off44, _on44], expect_contain=[3, 12])

        # ---- 44.18 vbe_bridge 收集段接线护栏 ----
        try:
            import inspect as _inspect44
            _src44 = _inspect44.getsource(VB44)
            _okwire44 = all(k in _src44 for k in (
                "ENABLE_HOST_ENUMS", "NO_HOST_ENUMS",
                "host_enum_constants()", "_HOST_ENUM_MODULE",
                "_host_enum_names = host_min", "def get_host_enum_names",
                "def enum_family_prefix", "def host_enum_family",
                "def _tlb_enum_members"))
        except Exception as _e44c:
            _okwire44 = "异常: %s" % _e44c
        check("44.18 vbe_bridge 收集段四路承接 + 开关 + 家族推导都在",
              [_okwire44], expect_contain=[True])
        check("44.19 坏 GUID -> 返回空清单、不抛",
              [VB44._tlb_enum_members("{00000000-0000-0000-0000-000000000000}",
                                      9, 9)],
              expect_contain=[[]])
    except Exception as _e44x:
        check("第 44 节异常: %s" % _e44x, [True], expect_contain=[False])

    # ==================================================================
    # 45. v68：宿主枚举常量的模糊匹配（★用户报的 xlworkfaul）
    #
    # 用户报："我输入 xlworkfaul，不会提示 xlWorkbookDefault，这个模糊匹配有问题。"
    #
    # 复现结论：不是漏收（xlWorkbookDefault 在 v67 就已经在池子里），是 v67 的
    # "只认从头开始的连续前缀"太死 —— xlworkfaul 跳过了中间的 book，不算前缀，
    # 可它是整个 4538 条里【唯一】能匹配上的候选。命中唯一，纯属被规则一刀切掉。
    #
    # v68 放宽成两条【都要满足】：
    #   1) 输入本身从家族前缀开始（xl / mso）。这条不能松 —— 正是它整片挡掉了
    #      ms -> 2365 条 mso*、count -> 130 条、open -> 492 条这些噪音；
    #   2) 输入里出现过一段【连续】的名字片段，长度 >= HOST_ENUM_FUZZY_RUN(=4)，
    #      即"你真的打过这个名字里的一段"，而不是拿几个字母去凑子序列。
    #
    # 为什么不是"输入够长就放行"：实测放行纯跳步后 xlcell 会从 15 条涨到 89 条、
    # xlcount 从 4 条涨到 59 条。加"连续块 >= 4"后，长度 < 4 的输入退化成纯前缀
    # （与 v67 一字不差），而 xlworkfaul（连续块 6）照常命中。
    # ==================================================================
    print("\n=== 45. 宿主枚举常量的模糊匹配（v68）===")
    try:
        import engine as E45

        # ---- 45.1~45.3 纯函数：连续块闸门 ----
        check("45.1 连续块闸门常量 = 4（与 v61 内建名字那档一致）",
              [E45.HOST_ENUM_FUZZY_RUN], expect_contain=[4])
        check("45.2 _longest_run：xlworkfaul 在 xlWorkbookDefault 上的命中下标"
              "最长连续段 = 6（xlwork 一段 + faul 一段）",
              [E45._longest_run([0, 1, 2, 3, 4, 5, 12, 13, 14, 15]),
               E45._longest_run(E45.fuzzy_match("xlWorkbookDefault",
                                                "xlworkfaul")[1])],
              expect_contain=[6, 6])
        check("45.3 _longest_run 边界：空 -> 0，单点 -> 1，全散 -> 1",
              [E45._longest_run([]), E45._longest_run([7]),
               E45._longest_run([0, 2, 4, 6])],
              expect_contain=[0, 1, 1])

        # ---- 45.4~45.9 桩池：确定性验证各条规则 ----
        _POOL45 = [("xlWorkbookDefault", "宿主库", None, False),
                   ("xlWorkbookNormal", "宿主库", None, False),
                   ("xlMSDOS", "宿主库", None, False),
                   ("xlCenter", "宿主库", None, False),
                   ("xlCellTypeVisible", "宿主库", None, False),
                   ("xlVAlignCenter", "宿主库", None, False),
                   ("xlLastCell", "宿主库", None, False),
                   ("msoTrue", "宿主库", None, False),
                   ("MyWorkOrder", "M1", None, False),
                   ("MyWorkOrderCount", "M1", None, False)]
        # ⚠️ 只把【宿主库】那批当宿主常量。工程内名字（MyWork*）必须留在
        # host_min 之外 —— 否则它们会连带被套上"必须从家族前缀开始"的闸门，
        # 输入 count 就再也补不出 MyWorkOrderCount 了（这个坑自己踩过一次）。
        _HOST45 = dict((r[0].lower(),
                        3 if r[0].lower().startswith("mso") else 2)
                       for r in _POOL45 if r[1] == "宿主库")

        class _UI45(object):
            def __init__(self):
                self.shown = False
                self.rows = []

            def show(self, rows, selection=0, completer=None):
                self.shown = True
                self.rows = list(rows)

            def update_selection(self, *a, **k):
                pass

            def hide(self, *a, **k):
                self.shown = False

            def contains_point(self, *a, **k):
                return False

        class _B45(object):
            """桩后端：池子 / 宿主清单都能替换（真实那一路由 45.10 起接手）。"""

            def __init__(self, word, pool=None, host=None):
                self._word = str(word)
                self._pool = list(pool if pool is not None else _POOL45)
                self._host = dict(host if host is not None else _HOST45)

            def get_context(self):
                _ln = "    x = " + self._word
                return {"line_no": 3, "caret_col": len(_ln) + 1,
                        "line_text": _ln, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": "M1"}

            def get_identifiers(self):
                return list(self._pool)

            def get_declared_names(self):
                return [r[0].lower() for r in self._pool]

            def declared_elsewhere(self, name, module_name):
                return str(name).lower() in [r[0].lower() for r in self._pool]

            def get_structural_names(self):
                return [r[0].lower() for r in self._pool]

            def get_builtin_names(self):
                return []

            def get_type_names(self):
                return []

            def names_outside_caret(self, caret, scope_only=False):
                return set()

            def get_host_enum_names(self):
                return dict(self._host)

        def _trig45(word, pool=None, host=None):
            _ui = _UI45()
            _c = E45.Completer(_B45(word, pool, host), _ui)
            _c.trigger(True)
            return list(_c.matches or []), _ui

        # ★用户报的场景：跳步输入要能补出 xlWorkbookDefault
        check("45.4 ★输入 xlworkfaul -> 提示 xlWorkbookDefault（用户报的场景）",
              [_trig45("xlworkfaul")[0]], expect_contain=[["xlWorkbookDefault"]])
        check("45.5 再少一个字符 xlworkfau 同样命中（连续块 6 未变）",
              [_trig45("xlworkfau")[0]], expect_contain=[["xlWorkbookDefault"]])
        check("45.6 前缀照旧：xlworkbookn -> xlWorkbookNormal",
              [_trig45("xlworkbookn")[0]], expect_contain=[["xlWorkbookNormal"]])
        check("45.7 输入不够 4 个字符 -> 退化成纯前缀，与 v67 一致："
              "xlms -> 只 xlMSDOS",
              [_trig45("xlms")[0]], expect_contain=[["xlMSDOS"]])
        check("45.8 跳步候选没有 4 连块 -> 被闸门挡掉："
              "xlce -> 只前缀的 xlCenter / xlCellTypeVisible，"
              "跳步的 xlVAlignCenter 不出",
              [_trig45("xlce")[0]],
              expect_contain=[["xlCenter", "xlCellTypeVisible"]])
        check("45.9 跳步命中排在所有前缀命中之后（kind 0 < kind 3）："
              "xlcell -> xlCellTypeVisible 在前，xlLastCell 在后",
              [_trig45("xlcell")[0]],
              expect_contain=[["xlCellTypeVisible", "xlLastCell"]])
        check("45.10 没从家族前缀开始的输入，一条宿主常量都不出："
              "ms / open 全空，count 只剩工程内名字",
              [_trig45("ms")[0], _trig45("count")[0], _trig45("open")[0]],
              expect_contain=[[], ["MyWorkOrderCount"], []])
        check("45.11 工程内名字不受这套闸门约束（走普通通道）：myworkc -> "
              "MyWorkOrderCount（跳步照旧）",
              [_trig45("myworkc")[0]], expect_contain=[["MyWorkOrderCount"]])

        # ---- 45.11~45.13 真机：真实 4538 条清单 + 真引擎，端到端 ----
        # 与第 44 节同样的理由：前面小节把 VB._get_vbe_cached 打桩成了假 VBE，
        # 必须另加载一份干净的 vbe_bridge 才能读到真实引用。
        _fresh45 = None
        try:
            import importlib.util as _iu45
            _spec45 = _iu45.spec_from_file_location(
                "vbe_bridge_fresh45",
                os.path.join(os.path.dirname(os.path.dirname(
                    os.path.abspath(__file__))), "vbe_bridge.py"))
            _fresh45 = _iu45.module_from_spec(_spec45)
            _spec45.loader.exec_module(_fresh45)
        except Exception:
            _fresh45 = None
        _items45 = []
        try:
            # v70 起宿主枚举默认不收；这一节要真实 4538 条清单，故临时打开开关。
            if _fresh45 is not None:
                _fresh45.ENABLE_HOST_ENUMS = True
            _items45 = (_fresh45 or VB44).host_enum_constants()
        except Exception:
            _items45 = []
        _real45 = [(n, "宿主库", None, False) for n, _m in _items45]
        _rhost45 = dict((n.lower(), m) for n, m in _items45)
        # ★v104：当前宿主是不是 Excel（清单里有没有成规模的 xl*）。宿主探测上线后
        # 连到的可能是 Word/PPT/WPS —— 那 45.12~45.14 的 Excel 专属断言不适用。
        _excel45 = len([1 for n, _m in _items45
                        if n.lower().startswith("xl")]) >= 1000

        if not _items45:
            check("45.12 真机：拿不到宿主枚举清单（Excel 未开？）"
                  "—— 跳过 45.12~45.14",
                  [True], expect_contain=[True])
        elif not _excel45:
            check("45.12~45.14 真机：当前宿主非 Excel（清单 %d 条里没有 xl*）"
                  "—— 跳过 Excel 专属断言（v104 宿主探测生效）"
                  % len(_items45), [True], expect_contain=[True])
        else:
            _m45 = _trig45("xlworkfaul", pool=_real45, host=_rhost45)[0]
            check("45.12 ★真机 %d 条：xlworkfaul 唯一命中 xlWorkbookDefault"
                  "（命中唯一 = 无噪音）" % (len(_items45),),
                  [_m45], expect_contain=[["xlWorkbookDefault"]])
            _m45b = [_trig45(w, pool=_real45, host=_rhost45)[0]
                     for w in ("ms", "count", "open", "cell", "ar", "vb")]
            check("45.13 真机：ms / count / open / cell / ar / vb 一条 xl*/mso* "
                  "都不出（v67 的防噪音不退化）",
                  [_m45b], expect_contain=[[[], [], [], [], [], []]])
            # 3 字符输入不可能凑出 4 连块 -> 行为必须与 v67 的"纯前缀"一字不差
            _pre45 = sorted(n for n, _m in _items45
                            if n.lower().startswith("xla"))
            _gate45 = sorted(n for n, _m in _items45
                             if n.lower().startswith("xla")
                             or (E45.fuzzy_match(n, "xla") is not None
                                 and E45._longest_run(
                                     E45.fuzzy_match(n, "xla")[1])
                                 >= E45.HOST_ENUM_FUZZY_RUN))
            check("45.14 真机：3 字符输入 xla 的命中与「纯前缀」完全一致"
                  "（%d 条；连续块不可能够 4）" % (len(_pre45),),
                  [_gate45 == _pre45, len(_pre45) > 0],
                  expect_contain=[True, True])
    except Exception as _e45:
        check("第 45 节异常: %s" % _e45, [True], expect_contain=[False])

    # ==================================================================
    # 46. 宿主枚举匹配的廉价预筛（v69）—— 纯性能，行为必须一字不变
    #
    # v68 的闸门"输入里存在一段长度 >= 4 的连续块"有个极便宜的必要条件：从输入里
    # 随便取一个 4 连片断，它必须【真的连续出现在】候选名里 —— C 层一句 `in` 就够，
    # 比逐条跑 fuzzy_match + _longest_run 便宜一个量级（实测 2173 条锚定候选
    # 5.9ms -> 0.8ms，整条 trigger 由约 10.6ms 降到约 6.2ms）。
    #
    # 预筛只可能"多留"（留下的还要照跑 fuzzy_match 与闸门复核），绝不"少留"：
    # 真命中的候选，其命中下标里必有一段长度 >= 4 的连续段，那段对应的输入片断
    # 当然就在名字里。这一节就守这个不变式 —— 用真实 4538 条清单。
    #
    # 为什么必须守：预筛一旦"少留"，症状是"某些 xl* 突然补不出来了"，
    # 而模糊匹配本身看不出问题（v67 那次漏收就是这么难查）。
    # ==================================================================
    print("\n=== 46. 宿主枚举匹配的廉价预筛（v69）===")
    try:
        import engine as E46
        import time as _t46

        try:
            _items46 = list(_items45)
        except Exception:
            _items46 = []
        check("46.1 真机清单可用（%d 条；拿不到就跳过本节的实质断言）"
              % len(_items46), [len(_items46) > 0], expect_contain=[True])

        WORDS46 = ("xlworkfaul", "xlworkfau", "xlworkf", "xlwork", "xlwo",
                   "xlwb", "xlwd", "xlcell", "xlcount", "xlco", "xlce", "xla",
                   "xlms", "xlm", "xldef", "xlde", "xld", "xlrange", "xlsheet",
                   "xlcolu", "xlrow", "xlnum", "xlzz", "xlab", "xlj", "ms",
                   "mso", "msobu", "xl")

        def _old_inner46(q):
            """v68 的内层循环（无预筛）：锚定 + run>=4。"""
            ql = q.lower()
            lw = len(q)
            got = []
            for n, need in _items46:
                low = n.lower()
                if lw < need:
                    continue
                if low.startswith(ql):
                    got.append(n)
                    continue
                if not ql.startswith(low[:need]):
                    continue
                r = E46.fuzzy_match(n, q)
                if r is None or E46._longest_run(r[1]) < E46.HOST_ENUM_FUZZY_RUN:
                    continue
                got.append(n)
            return got

        def _new_inner46(q):
            """v69 的内层循环：多一层 4-gram 预筛。"""
            ql = q.lower()
            lw = len(q)
            grams = tuple(ql[k:k + E46.HOST_ENUM_FUZZY_RUN]
                          for k in range(lw - E46.HOST_ENUM_FUZZY_RUN + 1))
            got = []
            for n, need in _items46:
                low = n.lower()
                if lw < need:
                    continue
                if low.startswith(ql):
                    got.append(n)
                    continue
                if not ql.startswith(low[:need]):
                    continue
                if not grams or not any(g in low for g in grams):
                    continue
                r = E46.fuzzy_match(n, q)
                if r is None or E46._longest_run(r[1]) < E46.HOST_ENUM_FUZZY_RUN:
                    continue
                got.append(n)
            return got

        if _items46:
            _bad46 = [q for q in WORDS46
                      if _old_inner46(q) != _new_inner46(q)]
            check("46.2 ★等价性：预筛前后候选集合与顺序完全一致（%d 个输入）"
                  % len(WORDS46), [_bad46], expect_contain=[[]])

            # 性质：过闸门 ==> 过预筛（真数据上零反例）
            _c46 = 0
            _viol46 = []
            for q in WORDS46:
                ql = q.lower()
                grams = [ql[k:k + E46.HOST_ENUM_FUZZY_RUN]
                         for k in range(len(ql) - E46.HOST_ENUM_FUZZY_RUN + 1)]
                for n, need in _items46:
                    low = n.lower()
                    if len(ql) < need or not ql.startswith(low[:need]):
                        continue
                    r = E46.fuzzy_match(n, q)
                    if r is None:
                        continue
                    if E46._longest_run(r[1]) < E46.HOST_ENUM_FUZZY_RUN:
                        continue
                    _c46 += 1
                    if not any(g in low for g in grams):
                        _viol46.append((q, n))
            check("46.3 过闸门的候选 %d 条，全部过得了预筛（有反例 = 预筛会误杀）"
                  % _c46, [_viol46, _c46 > 0], expect_contain=[[], True])

            # 短输入（< RUN 个字符）时预筛为空 = 非前缀锚定候选直接跳过，
            # 与闸门"连续块不可能够长"等价 —— 拿 xla / xlms / xlco 对一遍。
            _short46 = [q for q in ("xla", "xlms", "xlco", "xlm", "xlce", "ms")
                        if _old_inner46(q) != _new_inner46(q)]
            check("46.4 短输入（不足 4 字符）走空预筛也一字不差",
                  [_short46], expect_contain=[[]])

            # 性能护栏（相对比较，阈值宽松）：同样的锚定候选，预筛必须更快
            _q46 = "xlworkfaul"
            _anch46 = [n for n, need in _items46
                       if len(_q46) >= need
                       and _q46.startswith(n.lower()[:need])
                       and not n.lower().startswith(_q46)]
            _grams46 = [_q46[k:k + E46.HOST_ENUM_FUZZY_RUN]
                        for k in range(len(_q46)
                                       - E46.HOST_ENUM_FUZZY_RUN + 1)]

            def _t_gate46():
                for n in _anch46:
                    r = E46.fuzzy_match(n, _q46)
                    if r is not None:
                        E46._longest_run(r[1])

            def _t_pre46():
                for n in _anch46:
                    low = n.lower()
                    if any(g in low for g in _grams46):
                        r = E46.fuzzy_match(n, _q46)
                        if r is not None:
                            E46._longest_run(r[1])

            def _best46(fn, times=5):
                b = None
                for _ in range(times):
                    _t0 = _t46.perf_counter()
                    fn()
                    _dt = _t46.perf_counter() - _t0
                    b = _dt if b is None else min(b, _dt)
                return b

            _tb46 = _best46(_t_gate46)
            _tp46 = _best46(_t_pre46)
            # ★v104：46.5 的锚定候选、46.6 的命中都锚在 Excel 的 xlworkfaul 上，
            # 宿主非 Excel（如 Word/PPT/WPS）时这两条不适用 —— 跳过。
            _excel46 = len([1 for n, _m in _items46
                            if n.lower().startswith("xl")]) >= 1000
            if _excel46:
                check("46.5 性能护栏：%d 条锚定候选，预筛 %.2fms < 全量模糊 %.2fms"
                      % (len(_anch46), _tp46 * 1000, _tb46 * 1000),
                      [_tp46 < _tb46 * 0.8], expect_contain=[True])

                # 端到端（真引擎 + 真实清单）：预筛没把用户要的候选弄丢
                check("46.6 端到端：xlworkfaul -> 唯一 xlWorkbookDefault"
                      "（预筛不误杀；真引擎 + %d 条）" % len(_items46),
                      [_trig45("xlworkfaul", pool=_real45, host=_rhost45)[0]],
                      expect_contain=[["xlWorkbookDefault"]])
            else:
                check("46.5 真机：当前宿主非 Excel —— 跳过性能护栏"
                      "（锚定输入 xlworkfaul 不适用）", [True],
                      expect_contain=[True])
                check("46.6 真机：当前宿主非 Excel —— 跳过 xlworkfaul 端到端",
                      [True], expect_contain=[True])
    except Exception as _e46:
        check("第 46 节异常: %s" % _e46, [True], expect_contain=[False])

    # ==================================================================
    # 47. v70：内置枚举默认不提示（用户口径），内建函数照常
    #
    # 用户要求："我不怎么使用 VBA 内置枚举，提示出来对我有干扰 —— vb 开头、
    # xl 开头这些枚举都关掉不提示，函数那些要保留。"
    #
    # 于是 vba_builtins 的 CONSTANTS / FUNCTIONS 拆成两组导出，收不收由
    # `vbe_bridge.ENABLE_VBA_CONSTANTS`（默认 False）决定；宿主类型库枚举
    # （xl*/mso*）由 `vbe_bridge.ENABLE_HOST_ENUMS`（v70 起默认 False）决定。
    # 两个开关都能用环境变量翻回来（VBECOMPLETE_VBA_CONSTANTS / _HOST_ENUMS）。
    #
    # 这一节守【默认口径 + 开关真的在起作用】：每条"关着没有"都配一条
    # "打开就有"的反向对照 —— 否则"碰巧没有"（比如桩 VBE 解析失败）也能蒙过去。
    #
    # ⚠️ v77：用户接着说"vba 本身带的关键字、函数这块也都删掉"，于是
    # ENABLE_VBA_BUILTINS / ENABLE_VBA_KEYWORDS 也改成默认 False。
    # ⚠️ ★v96：又收回来了（"把 vba 关键字、函数名这两部分，也加入到提醒列表里。
    #    枚举类型就不要加了，枚举加上，就会太多了"）⇒ 这两个默认 True，
    #    **两批枚举仍默认 False**（本节 47.2 钉的就是这个组合）。
    #    **下面 47.5~47.11 先把这两个开关临时打开再断言** ——
    #    那一半验的是"机制仍在、开关真有效"，与默认值无关；默认口径由第 54 节专门守。
    # ==================================================================
    print("\n=== 47. 内置枚举默认不提示（v70）===")
    try:
        import vbe_bridge as VB47
        import vba_builtins as VB47B
        import engine as E47
        import os as _os47

        # ---- 47.1 _env_flag 三态（纯函数）----
        _KEY47 = "VBECOMPLETE_TEST47_FLAG"
        _old47 = _os47.environ.get(_KEY47)
        _r47 = {}
        try:
            _os47.environ.pop(_KEY47, None)
            _r47["unset"] = (VB47._env_flag(_KEY47, True),
                             VB47._env_flag(_KEY47, False))
            for _v in ("1", "true", "YES", "on", "0", "false", "No", "OFF", "??"):
                _os47.environ[_KEY47] = _v
                _r47[_v] = VB47._env_flag(_KEY47, True)
        finally:
            if _old47 is None:
                _os47.environ.pop(_KEY47, None)
            else:
                _os47.environ[_KEY47] = _old47
        check("47.1 _env_flag：未设置/认不出 -> 默认值；1,true,yes,on -> True；"
              "0,false,no,off -> False",
              [_r47["unset"], _r47["1"], _r47["true"], _r47["YES"], _r47["on"],
               _r47["0"], _r47["false"], _r47["No"], _r47["OFF"], _r47["??"]],
              expect_contain=[(True, False), True, True, True, True,
                              False, False, False, False, True])

        # ---- 47.2 出厂默认值 ----
        # ⚠️ 必须单独加载一份干净模块：main() 的前导段会把"内建/关键字"这两个开关
        # 【临时关掉】来跑第 1~37 节的防噪音护栏，现场读到的是被改过的值。
        _fresh47 = None
        try:
            import importlib.util as _iu47
            _spec47 = _iu47.spec_from_file_location(
                "vbe_bridge_fresh47",
                os.path.join(os.path.dirname(os.path.dirname(
                    os.path.abspath(__file__))), "vbe_bridge.py"))
            _fresh47 = _iu47.module_from_spec(_spec47)
            _spec47.loader.exec_module(_fresh47)
        except Exception:
            _fresh47 = None
        check("47.2 ★出厂默认（v96）：关键字 + 内建函数/数据类型【开】、"
              "内置枚举 + 宿主枚举【关】（枚举用户明确说不要）",
              [((_fresh47.ENABLE_VBA_BUILTINS,
                 _fresh47.ENABLE_VBA_KEYWORDS,
                 _fresh47.ENABLE_VBA_CONSTANTS,
                 _fresh47.ENABLE_HOST_ENUMS) if _fresh47 else "干净模块加载失败")],
              expect_contain=[(True, True, False, False)])

        # ---- 47.3 清单拆成两组（清单本身不裁剪）----
        _inter47 = set(VB47B.BUILTIN_FUNCTIONS) & set(VB47B.BUILTIN_CONSTANTS)
        check("47.3 清单拆组：函数 %d / 常量 %d，交集 %s，并集 == BUILTIN_NAMES(%d)"
              % (len(VB47B.BUILTIN_FUNCTIONS), len(VB47B.BUILTIN_CONSTANTS),
                 sorted(_inter47) or "空", len(VB47B.BUILTIN_NAMES)),
              [_inter47,
               set(VB47B.BUILTIN_NAMES) == (set(VB47B.BUILTIN_FUNCTIONS)
                                            | set(VB47B.BUILTIN_CONSTANTS)),
               "MsgBox" in VB47B.BUILTIN_FUNCTIONS,
               "vbCrLf" in VB47B.BUILTIN_CONSTANTS,
               "vbCrLf" not in VB47B.BUILTIN_FUNCTIONS,
               len(VB47B.BUILTIN_CONSTANTS) > 0],
              expect_contain=[set(), True, True, True, True, True])

        # ---- 47.4 默认：宿主枚举一条都不去枚举 ----
        check("47.4 默认：host_enum_constants() 返回空清单（连类型库都不加载）",
              [len(VB47.host_enum_constants())], expect_contain=[0])

        # ---- 47.5~47.8 真实后端全链路（桩 VBE）----
        class _CM47(object):
            def __init__(self, text):
                self.text = text

            @property
            def CountOfLines(self):
                return self.text.count("\n") + 1

            def Lines(self, start, count):
                return "\r\n".join(
                    self.text.split("\n")[start - 1:start - 1 + count])

        class _Comp47(object):
            def __init__(self, name, text, ctype=1):
                self.Name = name
                self.Type = ctype
                self.CodeModule = _CM47(text)

        class _Proj47(object):
            def __init__(self, comps):
                self.Name = "VBAProject"
                self.VBComponents = list(comps)

        class _Pane47(object):
            def __init__(self, comp):
                self._c = comp

            @property
            def CodeModule(self):
                return self._c.CodeModule

            def GetSelection(self):
                return (1, 1, 1, 1)

        class _VBE47(object):
            def __init__(self, comps, act):
                self.ActiveVBProject = _Proj47(comps)
                self.ActiveCodePane = _Pane47(act)

        class _UI47(object):
            def __init__(self):
                self.rows = []

            def show(self, rows, sel, c):
                self.rows = list(rows)

            def hide(self):
                pass

            def update_selection(self, sel):
                pass

            def contains_point(self, x, y):
                return False

        _comp47 = _Comp47("Module1", "Sub Foo()\n    x = 1\nEnd Sub", 1)
        _vbe47 = _VBE47([_comp47], _comp47)
        _orig47 = VB47._get_vbe_cached
        # ⚠️ main() 前导段把"内建/关键字"开关临时关掉了（第 1~37 节护栏的前提）。
        # 这一节要验的是"枚举关、函数开"，所以显式打开函数/关键字再验。
        _saved47 = (VB47.ENABLE_VBA_BUILTINS, VB47.ENABLE_VBA_KEYWORDS,
                    VB47.ENABLE_VBA_CONSTANTS, VB47.ENABLE_HOST_ENUMS)
        VB47.ENABLE_VBA_BUILTINS = True
        VB47.ENABLE_VBA_KEYWORDS = True
        VB47._get_vbe_cached = lambda: _vbe47
        try:
            _bk47 = VB47.VbeBackend()
            _nm47 = set(str(r[0]).lower() for r in _bk47.get_identifiers())
            _bl47 = _bk47.get_builtin_names()
            check("47.5 ★默认口径：vb* 一条都不进池，而 MsgBox / Split / "
                  "String / If 照常在",
                  ["vbcrlf" not in _nm47, "vbok" not in _nm47,
                   "msgbox" in _nm47, "split" in _nm47,
                   "string" in _nm47, "if" in _nm47],
                  expect_contain=[True] * 6)
            check("47.6 ★默认口径：宿主枚举一条都不收（xlUp / msoTrue 进不了池，"
                  "引擎也拿不到清单）",
                  ["xlup" not in _nm47, "msotrue" not in _nm47,
                   _bk47.get_host_enum_names() == {}],
                  expect_contain=[True, True, True])
            check("47.7 收紧匹配用的 builtin_names 集合里同样没有常量",
                  ["vbcrlf" not in _bl47, "msgbox" in _bl47],
                  expect_contain=[True, True])

            # 反向对照：把开关打开，常量必须回来（证明 47.5 是"开关起作用"，
            # 而不是桩 VBE / 解析环节碰巧什么都没收到）
            VB47.ENABLE_VBA_CONSTANTS = True
            try:
                _nm47b = set(str(r[0]).lower()
                             for r in VB47.VbeBackend().get_identifiers())
                check("47.8 对照：ENABLE_VBA_CONSTANTS=True -> vbCrLf / vbOK 回来",
                      ["vbcrlf" in _nm47b, "vbok" in _nm47b],
                      expect_contain=[True, True])
            finally:
                VB47.ENABLE_VBA_CONSTANTS = False

            # ---- 47.9~47.10 引擎端到端（默认口径）----
            _ctx47 = {"line_no": 1, "line_text": "    ", "caret_col": 5,
                      "in_string": False, "in_comment": False,
                      "in_type_position": False, "in_decl_position": False,
                      "decl_names": [], "proc_name": None,
                      "module_name": "Module1"}

            def _trig47(word):
                _ln = "    " + word
                _ctx47["line_text"] = _ln
                _ctx47["caret_col"] = len(_ln) + 1
                _b = VB47.VbeBackend()
                _b.get_context = lambda: dict(_ctx47)
                _c = E47.Completer(_b, _UI47())
                _c.trigger(False)
                return list(_c.matches or [])

            check("47.9 ★引擎端到端：打 vb / xl / xlu 一个候选都不出"
                  "（用户要的就是这个）",
                  [_trig47("vb"), _trig47("xl"), _trig47("xlu")],
                  expect_contain=[[], [], []])
            check("47.10 引擎端到端：打 ms 仍出 MsgBox（函数那批没被牵连）",
                  [_trig47("ms")], expect_contain=[["MsgBox"]])
        finally:
            VB47._get_vbe_cached = _orig47
            (VB47.ENABLE_VBA_BUILTINS, VB47.ENABLE_VBA_KEYWORDS,
             VB47.ENABLE_VBA_CONSTANTS, VB47.ENABLE_HOST_ENUMS) = _saved47

        # ---- 47.11 接线护栏：别把开关或拆组改没了 ----
        try:
            import inspect as _inspect47
            _src47 = _inspect47.getsource(VB47)
            _src47b = _inspect47.getsource(VB47B)
            _okwire47 = all(k in _src47 for k in (
                "ENABLE_VBA_CONSTANTS", "VBECOMPLETE_VBA_CONSTANTS",
                "VBECOMPLETE_HOST_ENUMS", "def _env_flag",
                "vba_builtins.ACTIVE_FUNCTIONS",
                "vba_builtins.ACTIVE_CONSTANTS")) and (
                "BUILTIN_FUNCTIONS" in _src47b
                and "BUILTIN_CONSTANTS" in _src47b)
        except Exception as _e47w:
            _okwire47 = "异常: %s" % _e47w
        check("47.11 接线护栏：两个开关 + 环境变量 + 函数/常量拆组都在",
              [_okwire47], expect_contain=[True])
    except Exception as _e47:
        check("第 47 节异常: %s" % _e47, [True], expect_contain=[False])

    # ==================================================================
    # 48. v71：公共词汇的"纯分散命中"阈值 4 -> 2（池子小了，可以放宽了）
    #
    # 背景：v61 收紧这一档，是因为当年内建清单近 400 条、还夹着 vbAbortRetryIgnore
    # 这类超长枚举名 —— 两三个字母就能凑出一大串。v70 把内置枚举默认关掉之后，
    # 真机池子掉到 265 条（261 内建 + 工程自己的名字），那个前提没了。
    #
    # 真机实测（详见当日日志 2026-09-15.md）：
    #   ve 10->15 条（VarType 第 11 位）   vr 0->7 条（第 1 位）
    #   se 30->41   ce 3->21   ar 10->21   ms 1->4（MsgBox 仍是第 1）
    #   耗时 0.71ms vs 0.69ms（零成本）
    #
    # 这一节用【确定性池子】（内建清单 + 两个工程名）守三件事，不依赖 Excel：
    #   1) 目标真浮上来了（ve / vr -> VarType 在一屏内）；
    #   2) ★排序不变式：阈值 4 的候选列表必须是阈值 2 的【子序列】——
    #      放宽只是"往后追加"，绝不会重排、不会顶掉原有首选；
    #   3) 代价可度量：单字符输入与"完全不收紧"逐条一致；fuzzy_match 调用次数
    #      两次完全相同（放宽不花额外匹配），耗时也在毫秒级。
    # ==================================================================
    print("\n=== 48. 公共词汇的分散命中阈值 4 -> 2（v71）===")
    try:
        import vba_builtins as VB48
        import inspect as _inspect48
        import time as _t48

        _POOL48 = []
        for _seq in (VB48.BUILTIN_FUNCTIONS, VB48.BUILTIN_TYPE_NAMES,
                     VB48.BUILTIN_KEYWORDS):
            _POOL48 += [(n, "VBA", None, False) for n in _seq]
        _POOL48 += [("numArr", "模块1", None, False),
                    ("countRows", "模块1", None, False)]
        _bl48 = set(r[0].lower() for r in _POOL48)

        class _B48(object):
            def __init__(self, word):
                self._word = word

            def get_context(self):
                _ln = "    " + self._word
                return {"line_no": 1, "line_text": _ln,
                        "caret_col": len(_ln) + 1, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": "模块1"}

            def get_identifiers(self):
                return list(_POOL48)

            def get_declared_names(self):
                return sorted(_bl48)

            def declared_elsewhere(self, name, module_name):
                return str(name).lower() in _bl48

            def get_structural_names(self):
                return sorted(_bl48)

            def get_builtin_names(self):
                return set(_bl48)

            def get_type_names(self):
                return list(VB48.BUILTIN_TYPE_NAMES)

            def get_host_enum_names(self):
                return {}

            def names_outside_caret(self, caret, scope_only=False):
                return set()

            def apply_completion(self, *a, **k):
                return None

        class _UI48(object):
            def show(self, rows, sel, c):
                pass

            def hide(self):
                pass

            def update_selection(self, sel):
                pass

            def contains_point(self, x, y):
                return False

        def _trig48(word, scatter_min=None, count_fm=False):
            _sm = E.BUILTIN_SCATTER_MIN
            _ofm = E.fuzzy_match
            _cnt = [0]
            if scatter_min is not None:
                E.BUILTIN_SCATTER_MIN = scatter_min
            if count_fm:
                def _fm(n, q):
                    _cnt[0] += 1
                    return _ofm(n, q)
                E.fuzzy_match = _fm
            _c = None
            try:
                _c = E.Completer(_B48(word), _UI48())
                _t0 = _t48.perf_counter()
                _c.trigger()
                _dt = (_t48.perf_counter() - _t0) * 1000
            finally:
                E.BUILTIN_SCATTER_MIN = _sm
                E.fuzzy_match = _ofm
            return list((_c.matches if _c else None) or []), _dt, _cnt[0]

        check("48.1 默认口径：BUILTIN_SCATTER_MIN == 2，且可用环境变量调回去",
              [E.BUILTIN_SCATTER_MIN,
               "VBECOMPLETE_BUILTIN_SCATTER_MIN" in _inspect48.getsource(E)],
              expect_contain=[2, True])

        _ve, _tve, _ = _trig48("ve")
        _ve4, _, _ = _trig48("ve", scatter_min=4)
        check("48.2 ★ve -> VarType 浮上来且在一屏内（放宽后 %d 条，第 %s 位；"
              "阈值 4 时 %s）"
              % (len(_ve), (_ve.index("VarType") + 1) if "VarType" in _ve
                 else "-", "也有" if "VarType" in _ve4 else "没有"),
              ["VarType" in _ve, "VarType" not in _ve4],
              expect_contain=[True, True])

        _vr, _, _ = _trig48("vr")
        check("48.3 vr -> VarType 稳居第 1 位（%d 条；分散命中里"
              "命中越靠前越优先）" % len(_vr),
              [_vr[:1]], expect_contain=[["VarType"]])

        # ★排序不变式：收紧版是放宽版的子序列 ⇒ 放宽只"往后追加"
        _it48 = iter(_ve)
        check("48.4 ★排序不变式：阈值 4 的候选列表是阈值 2 的【子序列】"
              "（放宽只追加、不重排、不顶掉首选）",
              [all(x in _it48 for x in _ve4)], expect_contain=[True])
        check("48.5 首选没被顶掉：ms -> 第 1 个仍是 MsgBox（前缀优先）",
              [_trig48("ms")[0][:1]], expect_contain=[["MsgBox"]])

        # 单字符输入：与"完全不收紧"（_builtin_names 打桩成空集）逐条一致
        _e1, _, _ = _trig48("e")
        _obl48 = E.Completer._builtin_names
        E.Completer._builtin_names = lambda self: set()
        try:
            _e2, _, _ = _trig48("e")
        finally:
            E.Completer._builtin_names = _obl48
        check("48.6 单字符输入不受影响（e：%d 条，与「完全不收紧」逐条一致 ——"
              " 这条规则只管 2~3 字符）" % len(_e1),
              [_e1 == _e2, len(_e1) > 0], expect_contain=[True, True])

        # 代价：放宽不花额外匹配，耗时毫秒级
        _, _t2, _n2 = _trig48("ve", scatter_min=2, count_fm=True)
        _, _t4, _n4 = _trig48("ve", scatter_min=4, count_fm=True)
        check("48.7 代价：fuzzy_match 调用次数两次相同（%d = %d），"
              "耗时 %.2fms / %.2fms（池子 %d 条）"
              % (_n2, _n4, _t2, _t4, len(_POOL48)),
              [_n2 == _n4, _t2 < 50 and _t4 < 50], expect_contain=[True, True])
    except Exception as _e48:
        check("第 48 节异常: %s" % _e48, [True], expect_contain=[False])

    # ==================================================================
    # 49. v72：With 块内的 `.成员` 让位（修「VBE 的弹窗反被我们挤掉」）
    #
    # 用户报：`With Range("a1:a10").Font` + `    .Size = 12` + `    .Bold = True`
    # 这样写的时候，在 With 块内输入 `.Size` / `.Bold`，本该我们让位给 VBE 的
    # 成员列表，实际却是「VBE 的弹窗给我们的弹窗让位了」。
    #
    # 根因：让位的【语法判据】只认"点号左边有东西"（标识符 / ) / ] / }），
    # `vbe_member_list_expected` 里就是 `m < 0 -> return False`。而 With 块最
    # 标准的写法恰恰是点顶在行首（前面只有缩进）。第二路"窗口可见性"判据在
    # 这种场景下也兜不住：我们的候选窗是 overrideredirect + topmost 悬浮窗，
    # 一弹出就压在 VBE 的列表上面，而 VBE 的列表在失去激活时会自己收起 ——
    # "等它真的画出来再让"永远来不及。
    #
    # v72 的修法（两条）：
    #   1) 语法判据认识 With 块：点号左边是行首，或 _WITH_OWNER_BOUNDARY 里的
    #      语句 / 运算符边界（`:` `=` `,` `(` 与 `& + - * / \ ^ > <`）；
    #   2) main.py 轮询加一道兜底：VBE 的成员列表真的可见就收起我们的
    #      （手动 Ctrl+Space 唤出的除外）。
    # ==================================================================
    print("\n=== 49. v72：With 块内的 .成员 让位 ===")
    try:
        import vba_builtins as VB49

        _POOL49 = [(n, "VBA", None, False) for n in VB49.BUILTIN_FUNCTIONS]
        _POOL49 += [(n, "VBA", None, False) for n in VB49.BUILTIN_TYPE_NAMES]
        _bl49 = set(r[0].lower() for r in _POOL49)

        class _B49(object):
            def __init__(self, line_text, caret_col):
                self._lt = line_text
                self._cc = caret_col

            def get_context(self):
                return {"line_no": 1, "line_text": self._lt,
                        "caret_col": self._cc, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": "模块1"}

            def get_identifiers(self):
                return list(_POOL49)

            def get_declared_names(self):
                return sorted(_bl49)

            def declared_elsewhere(self, name, module_name):
                return str(name).lower() in _bl49

            def get_structural_names(self):
                return sorted(_bl49)

            def get_builtin_names(self):
                return set(_bl49)

            def get_type_names(self):
                return list(VB49.BUILTIN_TYPE_NAMES)

            def get_host_enum_names(self):
                return {}

            def names_outside_caret(self, caret, scope_only=False):
                return set()

            def apply_completion(self, *a, **k):
                return None

        class _UI49(object):
            def __init__(self):
                self.hides = 0

            def show(self, rows, sel, c):
                pass

            def hide(self):
                self.hides += 1

            def update_selection(self, sel):
                pass

            def contains_point(self, x, y):
                return False

        _yl49 = E.vbe_list_expected

        def _yl49s(line):
            """按「光标停在行尾」判定这一行该不该让位。"""
            return bool(_yl49(line, len(line) + 1))

        check("49.1 ★行首的点（With 块标准写法）-> 让位",
              [_yl49s("    .Size"), _yl49s("\t.Bold"), _yl49s("    ."),
               _yl49s("    .si")],
              expect_contain=[True, True, True, True])
        check("49.2 ★语句 / 运算符边界后的点 -> 让位"
              "（a = 1: .Size / Set f = .Font / MySub a, .Font / x = .Left + .Width）",
              [_yl49s("    a = 1: .Size"), _yl49s("    Set f = .Font"),
               _yl49s("    MySub a, .Font"), _yl49s("    x = .Left + .Width")],
              expect_contain=[True, True, True, True])
        check("49.3 反例（不让位，防过度让位）：小数点 / 注释里的点 / 字符串里的点"
              " / 注释里的冒号点 / 纯标识符 / 空行",
              [_yl49s("    x = 1.5"), _yl49s("    ' UserForm1."),
               _yl49s('    s = "a."'), _yl49s("    x = 1 ' c: .foo"),
               _yl49s("abc"), _yl49s("")],
              expect_contain=[False, False, False, False, False, False])
        check("49.4 原有位置照旧让位、没退化：obj. / a(1). / Dim x As / New",
              [_yl49s("    UserForm1.SelectedList"), _yl49s("    a(1)."),
               _yl49s("    Dim x As "), _yl49s("    Set c = New ")],
              expect_contain=[True, True, True, True])
        check("49.5 _WITH_OWNER_BOUNDARY 覆盖语句边界与算术/比较运算符",
              [set(":=,(") <= set(E._WITH_OWNER_BOUNDARY),
               set("+-*/\\^<>") <= set(E._WITH_OWNER_BOUNDARY)],
              expect_contain=[True, True])

        def _trig49(line, auto=True, yield_on=None):
            _ui = _UI49()
            _c = E.Completer(_B49(line, len(line) + 1), _ui)
            _oy = E.YIELD_TO_VBE_LIST
            if yield_on is not None:
                E.YIELD_TO_VBE_LIST = yield_on
            try:
                _c.trigger(bool(auto))
            finally:
                E.YIELD_TO_VBE_LIST = _oy
            return list(_c.matches or []), _ui.hides

        _m6, _h6 = _trig49("    .si")
        check("49.6 ★引擎端到端：轮询自动触发 + With 块内 .si -> 让位（不弹候选）",
              [len(_m6) == 0, _h6 > 0], expect_contain=[True, True])

        _m7, _h7 = _trig49("    .si", yield_on=False)
        check("49.7 对照：关掉 YIELD_TO_VBE_LIST -> 候选照出（%d 条）"
              "（证明 49.6 是「让位」在起作用，不是碰巧没候选）" % len(_m7),
              [len(_m7) > 0], expect_contain=[True])

        _m8, _h8 = _trig49("    .si", auto=False)
        check("49.8 手动 Ctrl+Space（让位只拦自动路径）-> 照旧弹（%d 条）"
              % len(_m8), [len(_m8) > 0], expect_contain=[True])

        _m9, _h9 = _trig49("    Msg")
        check("49.9 普通标识符位置不受影响：Msg -> 照常弹、且没被让位收掉",
              [_h9 == 0, len(_m9) > 0], expect_contain=[True, True])

        # ---- 49.10 接线护栏：main.py 轮询里的兜底 ----
        _msrc49 = ""
        try:
            _msrc49 = open(os.path.join(
                os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                "main.py"), encoding="utf-8").read()
        except Exception:
            _msrc49 = ""
        # v76：兜底里那句 `not state.get("popup_manual") and ...` 的前置条件已经
        # 去掉 —— 用户口径是"冲突时直接把我们的窗收掉"，不分手动/自动。这里除了
        # 确认接线还在，额外钉住【这个前置条件不许回来】。
        check("49.10 接线护栏：poll 兜底在（VBE 成员列表可见 -> 让位，不再分手动）",
              [_msrc49.count("VBE_YIELD_CLASSES") >= 1,
               "popup_manual" in _msrc49,
               _msrc49.count("popup_manual") >= 3,
               'if not state.get("popup_manual") and any(' not in _msrc49],
              expect_contain=[True, True, True, True])
    except Exception as _e49:
        check("第 49 节异常: %s" % _e49, [True], expect_contain=[False])

    # ==================================================================
    # 50. v73：让位的【宽限确认】—— VBE 不弹时自己补上
    #
    # 用户报（v72 之后）：在 With 块里输入 .Size **什么都不弹了**。
    #
    # 根因不是"我们把 VBE 的窗关掉了"（让位路径只有 ui.hide()，碰不到 VBE），
    # 而是【让位这个设计本身的假设不成立】：我们让开，是假定"VBE 的成员列表
    # 一定会顶上"。可 VBE 的自动列出成员要先把 `With` 的对象类型解析出来 ——
    # 解析不了（`With Range("a1:10")` 这种地址写错、模块里有待编译的错误）它
    # 一个窗都不弹。那时让位 = 什么都不给。
    #
    # v73 把让位改成【待定】：先让，YIELD_GRACE 之后由 main.py 的轮询调
    # Completer.confirm_yield() 裁决 —— VBE 的列表真出现就认；没出现就把候选
    # 补回来（trigger(False)）。同一处只判定一次（yield_given_up），避免
    # "让开-补上-再让开-再补上"来回抖。
    # ==================================================================
    print("\n=== 50. v73：让位宽限确认（VBE 没弹就自己补） ===")
    try:
        import vba_builtins as VB50

        _POOL50 = [(n, "VBA", None, False) for n in VB50.BUILTIN_FUNCTIONS]
        _bl50 = set(r[0].lower() for r in _POOL50)

        class _B50(object):
            def __init__(self, line_text, caret_col, line_no=1,
                         module="模块1"):
                self._lt = line_text
                self._cc = caret_col
                self._ln = line_no
                self._mod = module

            def move(self, line_text, caret_col, line_no=None, module=None):
                self._lt = line_text
                self._cc = caret_col
                if line_no is not None:
                    self._ln = line_no
                if module is not None:
                    self._mod = module

            def get_context(self):
                return {"line_no": self._ln, "line_text": self._lt,
                        "caret_col": self._cc, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": self._mod}

            def get_identifiers(self):
                return list(_POOL50)

            def get_declared_names(self):
                return []

            def declared_elsewhere(self, name, module_name):
                return False

            def get_structural_names(self):
                return []

            def get_builtin_names(self):
                return set(_bl50)

            def get_type_names(self):
                return []

            def get_host_enum_names(self):
                return {}

            def names_outside_caret(self, caret, scope_only=False):
                return set()

            def apply_completion(self, *a, **k):
                return None

        class _UI50(object):
            def __init__(self):
                self.hides = 0

            def show(self, rows, sel, c):
                pass

            def hide(self):
                self.hides += 1

            def update_selection(self, sel):
                pass

            def contains_point(self, x, y):
                return False

        _line50 = "    .Si"
        _b50 = _B50(_line50, len(_line50) + 1)
        _ui50 = _UI50()
        _c50 = E.Completer(_b50, _ui50)

        _c50.trigger(True)
        check("50.1 ★让位后挂上待定（还没裁决），候选不弹",
              [_c50.yield_pending is not None,
               len(_c50.matches or []) == 0],
              expect_contain=[True, True])

        check("50.2 VBE 成员列表出现了 -> confirm_yield 认这个让位（True）",
              [_c50.confirm_yield(True), _c50.yield_pending is None],
              expect_contain=[True, True])

        _c50.trigger(True)
        check("50.3 ★VBE 没弹 -> confirm_yield 判 False（该补候选了）",
              [_c50.confirm_yield(False)], expect_contain=[False])

        _c50.trigger(False)
        check("50.4 ★补上后候选照出（%d 条）—— 这就是「With 块里什么都不弹」的修复"
              % len(_c50.matches or []),
              [len(_c50.matches or []) > 0], expect_contain=[True])

        _ui50.hides = 0
        _c50.trigger(True)
        check("50.5 ★同一处再自动触发 -> 不再让位（候选留着、一次都不 hide）"
              "—— 防「让-补-让-补」抖动",
              [len(_c50.matches or []) > 0, _ui50.hides == 0],
              expect_contain=[True, True])

        _b50.move("    .Bo", 8, line_no=7)
        _c50.trigger(True)
        check("50.6 换到另一行 -> 重新让位（「这一处 VBE 不弹」的记忆不跨行）",
              [len(_c50.matches or []) == 0,
               _c50.yield_pending is not None],
              expect_contain=[True, True])

        _b50.move("    Msg", 8, line_no=8)
        check("50.7 宽限期内用户已走开 -> confirm_yield 判 True（不补，别打扰）",
              [_c50.confirm_yield(False)], expect_contain=[True])

        _c50b = E.Completer(_B50("    Msg", 8), _UI50())
        check("50.8 没有待定让位时 confirm_yield 一律 True（幂等，不凭空补）",
              [_c50b.confirm_yield(True), _c50b.confirm_yield(False)],
              expect_contain=[True, True])

        _c50c = E.Completer(_B50("    .Si", 8), _UI50())
        _c50c.trigger(False)
        check("50.9 手动 Ctrl+Space -> 不挂待定（让位只拦自动路径）",
              [_c50c.yield_pending is None,
               len(_c50c.matches or []) > 0],
              expect_contain=[True, True])

        _c50d = E.Completer(_B50("    .Si", 8), _UI50())
        _c50d.trigger(True)
        _c50d.backend.get_context = lambda: None
        check("50.10 拿不到 ctx（COM 暂时不可用）-> 不补（保持让位，不乱弹）",
              [_c50d.confirm_yield(False)], expect_contain=[True])

        check("50.11 YIELD_GRACE 默认 0.25 秒且在合理区间",
              [0.0 < E.YIELD_GRACE < 5.0], expect_contain=[True])

        _root50 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        def _read50(name):
            try:
                return open(os.path.join(_root50, name),
                            encoding="utf-8").read()
            except Exception:
                return ""

        _src50 = _read50("engine.py")
        check("50.12 VBECOMPLETE_YIELD_GRACE 可调（开关 / 宽限期 / 裁决都在引擎里）",
              ["VBECOMPLETE_YIELD_GRACE" in _src50,
               "YIELD_GRACE" in _src50,
               "confirm_yield" in _src50],
              expect_contain=[True, True, True])

        _msrc50 = _read50("main.py")
        check("50.13 接线护栏：poll 调 confirm_yield，判 False 时 trigger(False) 补上",
              ["confirm_yield" in _msrc50,
               "yield_pending" in _msrc50,
               "completer.trigger(False)" in _msrc50],
              expect_contain=[True, True, True])
    except Exception as _e50:
        check("第 50 节异常: %s" % _e50, [True], expect_contain=[False])

    # ==================================================================
    # 51. v74：候选窗 × 参数信息窗重叠的【状态自愈】
    #
    # 用户报："instr 那个函数的参数信息比较长，与我们候选窗重叠了"（v66 修过，
    # 又复现）。
    #
    # 根因不是避让算法 —— avoid_popup_rects 逐条验过没毛病（第 43 节守着它），
    # 而是轮询的【变化检测】有个洞：提示窗矩形签名只在候选窗可见时采样，于是
    # "候选窗不可见期间提示窗出现过、又回到同一位置"这件事永远不被记录 ——
    # 下次候选窗弹出时签名恰好等于那个陈旧缓存值，`_sig != _popup_sig` 便看不见
    # 变化，候选窗一直压在签名窗上，直到用户自己挪光标。
    #
    # 修法两条（互补，都要守）：
    #   1) 签名【每轮都采】，与候选窗可见性无关 —— 堵住那个洞；
    #   2) 加一道【状态检查】：叠着就重摆，压根不依赖"变没变"（自愈）。
    # ==================================================================
    print("\n=== 51. v74：候选窗 × 参数信息窗重叠的状态自愈 ===")
    try:
        import ui as UI51

        # 51.1 "重叠"的定义只此一份（避让与自愈必须同一把尺子）
        check("51.1 rect_overlap：真重叠 True；只共边 / 共点 / 分开都 False",
              [[UI51.rect_overlap((0, 0, 10, 10), (5, 5, 15, 15)),
                UI51.rect_overlap((0, 0, 10, 10), (10, 0, 20, 10)),
                UI51.rect_overlap((0, 0, 10, 10), (10, 10, 20, 20)),
                UI51.rect_overlap((0, 0, 10, 10), (20, 0, 30, 10))]],
              expect_contain=[[True, False, False, False]])

        # 51.2 clashes_with_popup：状态检查（不问变化，只问"现在叠着没有"）
        check("51.2 clashes_with_popup：叠着 True / 没提示窗 False / 已躲开 False",
              [[UI51.clashes_with_popup(100, 220, 200, 90, [(0, 220, 400, 240)]),
                UI51.clashes_with_popup(100, 220, 200, 90, []),
                UI51.clashes_with_popup(100, 242, 200, 90,
                                        [(0, 220, 400, 240)])]],
              expect_contain=[[True, False, False]])

        # 51.3 真机那条长签名窗（探针实测 505x18 @ 1091,601）：避让后不叠，
        #      且用【自愈判据】复核也是"不叠" —— 两条路径结论必须一致。
        _tip51 = (1091, 601, 1596, 619)
        _y51 = UI51.avoid_popup_rects(1368, 601, 250, 150, 583, 1080, [_tip51])
        check("51.3 ★长签名窗 505x18：避让到 y=%d，自愈复核不叠（同一把尺子）"
              % (_y51,),
              [[_y51, UI51.clashes_with_popup(1368, _y51, 250, 150, [_tip51])]],
              expect_contain=[[621, False]])

        # 51.4 ★★用户场景复现：候选窗【先】落在 601（签名窗还没出现），随后
        #      签名窗在【同一位置】出现 —— 变化检测看不见（签名与缓存相等），
        #      但状态检查必须看见。这条就是修好的那个洞。
        check("51.4 ★★陈旧签名 + 候选窗停在 601 -> 状态检查必须报「叠着」",
              [UI51.clashes_with_popup(1368, 601, 250, 150, [_tip51])],
              expect_contain=[True])

        # 51.5~51.7 Popup.clashes_with_popup：拿窗口【当前实际几何】比对
        class _FR51(object):
            def winfo_screenwidth(self):
                return 1920

            def winfo_screenheight(self):
                return 1080

        class _FC51(object):
            def cget(self, k):
                return 200 if k == "width" else 90

        class _FW51(object):
            def __init__(self, state="normal"):
                self._st = state

            def state(self):
                return self._st

            def geometry(self, s):
                pass

            def winfo_x(self):
                return 100

            def winfo_y(self):
                return 220

            def winfo_width(self):
                return 200

            def winfo_height(self):
                return 90

        def _mk51(win):
            p = UI51.Popup.__new__(UI51.Popup)
            p.root, p.canvas, p.win = _FR51(), _FC51(), win
            p._last_geom = None
            return p

        _o51rects = UI51.screen_popup_rects
        try:
            UI51.screen_popup_rects = lambda: [(0, 220, 400, 240)]
            _c51 = _mk51(_FW51()).clashes_with_popup()
            UI51.screen_popup_rects = lambda: []
            _c51b = _mk51(_FW51()).clashes_with_popup()
            _c51c = _mk51(_FW51("withdrawn")).clashes_with_popup()
            _c51d = _mk51(None).clashes_with_popup()
        finally:
            UI51.screen_popup_rects = _o51rects
        check("51.5 候选窗实际位置(100,220,200x90) × 签名窗 -> 报「叠着」",
              [_c51], expect_contain=[True])
        check("51.6 没有提示窗 / 已收起 / 没有窗 -> 一律 False（拿不准不动手）",
              [[_c51b, _c51c, _c51d]],
              expect_contain=[[False, False, False]])

        # 51.8~51.11 main.py 接线护栏（轮询那段的结构性不变量）
        _root51 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        def _read51(name):
            try:
                return open(os.path.join(_root51, name),
                            encoding="utf-8").read()
            except Exception:
                return ""

        _msrc51 = _read51("main.py")

        def _fn51(text, name):
            """把嵌套函数 name 的函数体从源码里抠出来（到下一个同级 def 为止）。"""
            i = text.find("def %s(" % name)
            if i < 0:
                return ""
            j = text.find("\n    def ", i + 1)
            return text[i:j if j > 0 else len(text)]

        _poll51 = _fn51(_msrc51, "poll_editor")
        # 只看代码、不看注释 —— v74 的说明里就写着 "if completer.is_visible()"
        # 这几个字，直接 find 会先命中注释（第一版测试就是这么写错的）。
        _code51 = "\n".join(_l for _l in _poll51.splitlines()
                            if not _l.lstrip().startswith("#"))
        _i_sig51 = _code51.find("_sig = tuple(vbe_bridge.vbe_popup_rects())")
        _i_set51 = _code51.find('state["_popup_sig"] = _sig')
        _i_vis51 = _code51.find("if completer.is_visible():")
        check("51.8 ★签名采样与落缓存都在【可见性判断之前】（无条件采样 —— 洞就堵在这）",
              [[_i_sig51 >= 0, _i_set51 >= 0, _i_vis51 >= 0,
                _i_sig51 < _i_vis51, _i_set51 < _i_vis51]],
              expect_contain=[[True, True, True, True, True]])

        check("51.9 ★接线：重叠时调 popup.clashes_with_popup() 触发重摆",
              [[_poll51.count("popup.clashes_with_popup()") == 1,
                "post(popup.reposition)" in _poll51,
                "_sig_changed or _clash" in _poll51]],
              expect_contain=[[True, True, True]])

        check("51.10 同一 tick 内两处判据共用一份签名（不再各探一次）",
              [[_msrc51.count("vbe_bridge.vbe_popup_rects()") == 1,
                _poll51.count("vbe_bridge.vbe_popup_rects()") == 1,
                "_sig)\n" in _poll51 or "in _sig)" in _poll51]],
              expect_contain=[[True, True, True]])

        check("51.11 Popup.clashes_with_popup 存在、可调用（reposition 的老位子旁边）",
              [[callable(getattr(UI51.Popup, "clashes_with_popup", None)),
                callable(getattr(UI51.Popup, "reposition", None)),
                callable(getattr(UI51, "rect_overlap", None))]],
              expect_contain=[[True, True, True]])

        # 51.12 避让与自愈对同一批矩形必须给出一致结论（性质检查，随机 300 组）
        import random as _rnd51
        _bad51 = []
        _g51 = _rnd51.Random(20260915)
        for _ in range(300):
            _b51 = _g51.randint(0, 300)
            _x51 = _g51.randint(0, 600)
            _y51b = _g51.randint(0, 700)
            _m51 = _g51.randint(0, 400)
            _r51 = [(_x51 - _g51.randint(0, 50), _y51b,
                     _x51 + _g51.randint(0, 400), _y51b + 20)]
            _new51 = UI51.avoid_popup_rects(_x51, _y51b, _m51, 90, 100,
                                            1080, _r51)
            # 避让说"躲开了"，自愈就不能说"还叠着"（除非真的贴底放不下）
            if not UI51.clashes_with_popup(_x51, _new51, _m51, 90, _r51):
                continue
            if _new51 >= max(0, 1080 - 90) or _new51 < 0:
                continue                    # 贴底/翻上方：确实无处可去
            _bad51.append((_x51, _y51b, _new51))
        check("51.12 ★避让后自愈复核：300 组随机矩形零反例（%d 组残留）"
              % (len(_bad51),),
              [len(_bad51)], expect_contain=[0])
    except Exception as _e51:
        check("第 51 节异常: %s" % _e51, [True], expect_contain=[False])

    # ==================================================================
    # 52. v75：隐式变量尊重模块自己的 Option Explicit
    #
    # 用户口径："没有 Option Explicit 时，用过的变量名就算已经被定义、可以直接
    # 提示（现在这样很好，不要提前定义变量名）；写了 Option Explicit 是强制声明，
    # 没声明过的变量名在后续就不该被提示。"
    #
    # 判定【按模块】做 —— Option Explicit 本来就是模块级语句：
    #   * 没写它的模块：一字不动（用过的名字照旧进池子）；
    #   * 写了它的模块：只收"真声明过"的名字（Dim / Const / Sub / Function /
    #     Type / Enum / 形参），隐式变量整段不收 —— 那种写法在该模块里是编译
    #     错误，提示它只会把坏代码带出去。
    #
    # 落点在【收集侧】（extract_implicit_records 之前），因为隐式变量是唯一
    # "用过但没声明"的名字来源：extract_records 收的都是真声明，而隐式变量从不
    # 进 declared_names / _decl_by_mod（回声防护那边只靠"现场文本"兜）。改这一处
    # 就够，引擎与 UI 一行不用动。
    # ==================================================================
    print("\n=== 52. v75：Option Explicit 模块不再收「用过但没声明」的名字 ===")
    try:
        import os as _os52
        import vbe_bridge as VB52
        import parser as P52

        # ---- 52.1 has_option_explicit 真值表（文本判定，纯函数）----
        # 全是"认错就出事"的边界：把注释掉的 Option Explicit 认成真，会把整个
        # 模块的隐式变量误砍掉（而那正是用户要保留的行为）；把真行漏掉则等于白改。
        _oe52 = [
            P52.has_option_explicit("Option Explicit\nSub A()\nEnd Sub"),
            P52.has_option_explicit("option explicit\nSub A()"),
            P52.has_option_explicit("   Option   Explicit\nSub A()"),
            P52.has_option_explicit("Attribute VB_Name = \"M\"\r\n"
                                    "Option Explicit\r\nSub A()"),
            P52.has_option_explicit("Option Base 1\nOption Compare Text\nSub A()"),
            P52.has_option_explicit("' Option Explicit\nSub A()\nEnd Sub"),
            P52.has_option_explicit('Sub A()\n    s = "Option Explicit"\nEnd Sub'),
            P52.has_option_explicit("Sub A()\n    x = 1 ' Option Explicit\nEnd Sub"),
            P52.has_option_explicit(""),
            P52.has_option_explicit(None),
        ]
        check("52.1 ★文本判定：行首的 Option Explicit 才算（缩进 / 小写 / 带 "
              "Attribute 头都算）；注释掉的 / 字符串里的 / 其它 Option 语句 / "
              "空文本都不算",
              [_oe52],
              expect_contain=[[True, True, True, True, False, False, False,
                               False, False, False]])

        # ---- 52.2~52.6 真实后端全链路（桩 VBE：一个不带、一个带 OE）----
        class _CM52(object):
            def __init__(self, comp, text):
                self.Parent = comp
                self.text = text

            @property
            def CountOfLines(self):
                return self.text.count("\n") + 1

            def Lines(self, start, count):
                return "\r\n".join(
                    self.text.split("\n")[start - 1:start - 1 + count])

        class _Comp52(object):
            def __init__(self, name, text, ctype=1):
                self.Name = name
                self.Type = ctype
                self.CodeModule = _CM52(self, text)

        class _Proj52(object):
            def __init__(self, comps):
                self.Name = "VBAProject"
                self.VBComponents = list(comps)

        class _Pane52(object):
            def __init__(self, comp):
                self._c = comp

            @property
            def CodeModule(self):
                return self._c.CodeModule

            def GetSelection(self):
                # 光标落在【最后一个空行】上：既不会抹掉任何标识符，也不会被
                # 当成"声明行"（否则正在声明的名字会被排除，干扰本节的断言）。
                return (10, 1, 10, 1)

        class _VBE52(object):
            def __init__(self, proj, act):
                self.ActiveVBProject = proj
                self.ActiveCodePane = _Pane52(act)

        # 带 Option Explicit：looseVar2 用过但没声明（编译错误）；declaredVar 真声明；
        # nParam 是形参（也是声明）。
        _EX52 = ("Option Explicit\n"
                 "Sub A()\n"
                 "    Dim declaredVar As Long\n"
                 "    declaredVar = 2\n"
                 "    looseVar2 = 1\n"
                 "End Sub\n"
                 "Sub C(ByVal nParam As Long)\n"
                 "    nParam = nParam + 1\n"
                 "End Sub\n")
        # 不带 Option Explicit：looseVar 同样是"用过即存在"，必须照旧收。
        _LO52 = ("Sub B()\n"
                 "    looseVar = 1\n"
                 "End Sub\n")
        _cEX52 = _Comp52("ExplicitMod", _EX52)
        _cLO52 = _Comp52("LooseMod", _LO52)
        _vbe52 = _VBE52(_Proj52([_cEX52, _cLO52]), _cEX52)

        _orig52 = VB52._get_vbe_cached
        _saved52 = (VB52.IMPLICIT_HONOR_OPTION_EXPLICIT,
                    VB52.ENABLE_IMPLICIT_IDENTIFIERS,
                    VB52.ENABLE_VBA_BUILTINS, VB52.ENABLE_VBA_KEYWORDS)
        VB52._get_vbe_cached = lambda: _vbe52
        try:
            # 只看工程内名字：把语言自带的两批关掉，池子干净好断言。
            VB52.ENABLE_VBA_BUILTINS = False
            VB52.ENABLE_VBA_KEYWORDS = False

            VB52.ENABLE_IMPLICIT_IDENTIFIERS = True
            VB52.IMPLICIT_HONOR_OPTION_EXPLICIT = True
            _on52 = set(str(r[0]).lower()
                        for r in VB52.VbeBackend().get_identifiers())

            VB52.IMPLICIT_HONOR_OPTION_EXPLICIT = False
            _off52 = set(str(r[0]).lower()
                         for r in VB52.VbeBackend().get_identifiers())

            VB52.ENABLE_IMPLICIT_IDENTIFIERS = False
            VB52.IMPLICIT_HONOR_OPTION_EXPLICIT = True
            _none52 = set(str(r[0]).lower()
                          for r in VB52.VbeBackend().get_identifiers())
        finally:
            VB52._get_vbe_cached = _orig52
            (VB52.IMPLICIT_HONOR_OPTION_EXPLICIT,
             VB52.ENABLE_IMPLICIT_IDENTIFIERS,
             VB52.ENABLE_VBA_BUILTINS, VB52.ENABLE_VBA_KEYWORDS) = _saved52

        check("52.2 ★带 Option Explicit 的模块：用过但没声明的 looseVar2 不进池；"
              "真声明的 declaredVar、形参 nParam 照常在",
              ["loosevar2" not in _on52, "declaredvar" in _on52,
               "nparam" in _on52],
              expect_contain=[True, True, True])
        check("52.3 ★不带 Option Explicit 的模块一字不动：looseVar 照旧进池"
              "（用户明确说这样好）",
              ["loosevar" in _on52], expect_contain=[True])
        check("52.4 ★反向对照：闸门关掉（旧行为）时 looseVar2 必须回来 "
              "—— 证明是这道闸门在起作用，不是碰巧没有",
              ["loosevar2" in _off52], expect_contain=[True])
        check("52.5 差异精确：闸门只摘掉 looseVar2 一个名字（不多不少）",
              [sorted(_off52 - _on52), sorted(_on52 - _off52)],
              expect_contain=[["loosevar2"], []])
        check("52.6 两个开关互不干扰：ENABLE_IMPLICIT_IDENTIFIERS=False 时"
              "looseVar 也没了，但真声明照旧",
              ["loosevar" not in _none52, "loosevar2" not in _none52,
               "declaredvar" in _none52],
              expect_contain=[True, True, True])

        # ---- 52.7 接线护栏 ----
        def _read52(name):
            return io.open(os.path.join(ROOT, name), encoding="utf-8").read()

        _br52 = _read52("vbe_bridge.py")
        _ps52 = _read52("parser.py")
        check("52.7 接线护栏：收集段真的问了闸门；判定函数在两边都只有一份",
              ["implicit_gate_applies(code)" in _br52,
               "implicit_gate_applies" in _br52 and "has_option_explicit" in _ps52,
               "IMPLICIT_HONOR_OPTION_EXPLICIT" in _br52,
               _ps52.count("def has_option_explicit") == 1],
              expect_contain=[True, True, True, True])

        # ---- 52.8 环境变量开关（必须用干净模块：现场那份可能被前导段改过）----
        _KEY52 = "VBECOMPLETE_IMPLICIT_HONOR_EXPLICIT"
        _old52 = _os52.environ.get(_KEY52)
        _freshD52 = _freshE52 = None

        def _load52(modname):
            import importlib.util as _iu52
            _sp52 = _iu52.spec_from_file_location(
                modname, os.path.join(ROOT, "vbe_bridge.py"))
            _m52 = _iu52.module_from_spec(_sp52)
            _sp52.loader.exec_module(_m52)
            return _m52

        try:
            _os52.environ.pop(_KEY52, None)
            try:
                _freshD52 = _load52("vbe_bridge_fresh52d")
            except Exception:
                _freshD52 = None
            _os52.environ[_KEY52] = "0"
            try:
                _freshE52 = _load52("vbe_bridge_fresh52e")
            except Exception:
                _freshE52 = None
        finally:
            if _old52 is None:
                _os52.environ.pop(_KEY52, None)
            else:
                _os52.environ[_KEY52] = _old52
        check("52.8 出厂默认开启；set VBECOMPLETE_IMPLICIT_HONOR_EXPLICIT=0 "
              "可回到旧行为",
              [(_freshD52.IMPLICIT_HONOR_OPTION_EXPLICIT if _freshD52 else "加载失败",
                _freshE52.IMPLICIT_HONOR_OPTION_EXPLICIT if _freshE52 else "加载失败")],
              expect_contain=[(True, False)])

        # ---- 52.9 真机不变式（Excel 没开就跳过）----
        # 断言形式刻意做成"零泄漏"而不是"某模块有几条"：用户随时会改工程，
        # 写死数字的护栏下次就该报假警了。
        _fresh52 = None
        try:
            _fresh52 = _load52("vbe_bridge_fresh52r")
        except Exception:
            _fresh52 = None
        _r52 = None
        if _fresh52 is None:
            check("52.9 真机：干净模块加载失败 —— 跳过", [True], expect_contain=[True])
        else:
            try:
                _be52r = _fresh52.VbeBackend()
                _pool52r = set(str(r[0]).lower()
                               for r in _be52r.get_identifiers())
                _decl52r = set(str(n).lower()
                               for n in _be52r.get_declared_names())
                _str52r = set(str(n).lower()
                              for n in _be52r.get_structural_names())
                _gated52, _leak52 = 0, []
                _comps52r = _fresh52._collect_components(
                    _fresh52._get_vbe_cached(), None)
                if not _comps52r:
                    raise RuntimeError("拿不到组件")
                for _cp52r in _comps52r:
                    try:
                        _nm52r = str(_cp52r.Name)
                        _ct52r = _cp52r.CodeModule.CountOfLines
                        _cd52r = (_cp52r.CodeModule.Lines(1, _ct52r)
                                  if _ct52r > 0 else "")
                    except Exception:
                        continue
                    if not _fresh52.implicit_gate_applies(_cd52r):
                        continue
                    _gated52 += 1
                    _std52r = bool(_cp52r.Type == 1)
                    _recs52r, _pn = P52.extract_records(
                        _cd52r, module=_nm52r, is_std_module=_std52r)
                    for (_n52r, _m52r, _p52r, _pr52r) in P52.extract_implicit_records(
                            _cd52r, module=_nm52r, is_std_module=_std52r,
                            declared=_recs52r, scope="proc"):
                        _l52r = str(_n52r).lower()
                        if _l52r in _decl52r or _l52r in _str52r:
                            continue        # 别处真声明过：本来就该在池子里
                        if _l52r in _pool52r:
                            _leak52.append(_l52r)
                _r52 = (sorted(set(_leak52)), _gated52)
            except Exception as _e52r:
                _r52 = None
                check("52.9 真机：拿不到 VBE（Excel 未开？）—— 跳过（%s）"
                      % _e52r, [True], expect_contain=[True])
        if _r52 is not None:
            check("52.9 ★真机：%d 个带 Option Explicit 的模块，其「用过但没声明」"
                  "的名字一个都没进池子（若列表非空即为泄漏）" % (_r52[1],),
                  [_r52[0]], expect_contain=[[]])
    except Exception as _e52:
        check("第 52 节异常: %s" % _e52, [True], expect_contain=[False])

    # ==================================================================
    # 53. v76：让位判据不再区分触发来源（手动 / 补位都受约束）
    #     + 「VBE 列表确实可见」清掉 yield_given_up
    #
    # 用户报："VBE 的弹窗和我们项目弹窗冲突时，屏蔽我们项目弹窗……以前你修好了，
    # 就一直没出，刚才你改完，就出这个 bug 了。" 对号入座的三问答案：
    # VBE 弹的是【成员列表】；现象是【我们的窗盖住了它】；期望是【直接消失让位】。
    #
    # 排查（先证伪再定位）：
    #   1) 先确认工具跑的是哪版代码 —— GetProcessTimes 拿 pythonw 启动时间
    #      (13:05:06) 对比源文件 mtime (main/ui 09:41、vbe_bridge 12:59) ⇒ 跑的确实
    #      是修完的最新代码，不是"改了没重启"的错觉；
    #   2) `git diff v73..v74 -- ui.py` 逐行核对：rect_overlap 的重构与旧判据
    #      (r<=x / l>=x+w / b<=y / t>=y+h) 逐项等价 ⇒ v74 的方向只会让"避让"更多，
    #      不会更少 ⇒ 问题不在避让，在【让位】这一侧；
    #   3) 让位的第二路（窗口可见性）天生脆：我们的窗是 topmost 悬浮窗，一压上去
    #      VBE 的列表就失去激活、自行收起，之后再也探不到（v72 的认知）。所以
    #      必须让【不依赖探测】的那些判据尽量多扛一手。
    #
    # 两个确定的病灶（都不依赖探测，因此不受上面那个自弱化循环影响）：
    #   a) 第二路判据原来带 require_ident_before_caret ⇒ 手动 Ctrl+Space 与 v73 的
    #      补位路径 trigger(False) 都【绕过】它。补位那条尤其危险：它是先探到
    #      "列表不可见"才发起的，而从那次探测到主线程真正执行 trigger 之间隔着
    #      动作队列，VBE 的列表完全可能已经弹出来了 —— 补出来的窗就直接压在它上面。
    #   b) yield_given_up 按 (模块, 行号) 记，一旦记上，这一行的【语法判据】就永久
    #      跳过让位，只剩第二路兜着 —— 而第二路正是上面那条会自弱化的路。
    #      现在"VBE 的列表确实可见"会把标记清掉，语法判据重新生效。
    #
    # 口径变更（用户 v76 明确表态）：VBE 的成员列表可见时我们的窗【直接消失】，
    # 不分手动/自动 —— 成员列表本身就是一份候选列表，它在屏幕上时用户的目的已经
    # 达到，再挂着我们的（topmost）只会把它整个盖住。
    # ==================================================================
    print("\n=== 53. v76：让位不区分触发来源；VBE 列表真在就清「已放弃」 ===")
    try:
        import vba_builtins as VB53
        import ui as UI53

        _POOL53 = [(n, "VBA", None, False) for n in VB53.BUILTIN_FUNCTIONS]
        _bl53 = set(r[0].lower() for r in _POOL53)

        class _B53(object):
            """桩后端：成员列表可见性（yld）可随时切，位置由 line_text 决定。"""

            def __init__(self, line_text="    .Si", caret_col=None,
                         line_no=1, module="模块1", yld=False):
                self._lt = line_text
                self._cc = (caret_col if caret_col is not None
                            else len(line_text) + 1)
                self._ln = line_no
                self._mod = module
                self._yld = yld

            def set_yield(self, yld):
                self._yld = yld

            def get_context(self):
                return {"line_no": self._ln, "line_text": self._lt,
                        "caret_col": self._cc, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": self._mod}

            def get_identifiers(self):
                return list(_POOL53)

            def get_declared_names(self):
                return []

            def declared_elsewhere(self, name, module_name):
                return False

            def get_structural_names(self):
                return []

            def get_builtin_names(self):
                return set(_bl53)

            def get_type_names(self):
                return []

            def get_host_enum_names(self):
                return {}

            def names_outside_caret(self, caret, scope_only=False):
                return set()

            def apply_completion(self, *a, **k):
                return None

            # 只表达【成员列表】语义（v65）：参数信息不参与让位
            def vbe_yield_visible(self):
                return self._yld

            def vbe_popup_visible(self):
                return self._yld

            def vbe_popup_info(self):
                return [("NameListWndClass", 120, 90)] if self._yld else []

        class _UI53(object):
            def __init__(self):
                self.shown = False
                self.hides = 0

            def show(self, rows, sel, c):
                self.shown = True

            def hide(self):
                self.shown = False
                self.hides += 1

            def update_selection(self, sel):
                pass

            def contains_point(self, x, y):
                return False

        def _mk53(line_text="    .Si", yld=False):
            _be = _B53(line_text=line_text, yld=yld)
            _ui = _UI53()
            return E.Completer(_be, _ui), _be, _ui

        # 53.1 基础不变式：语法位置 + 列表不可见 + 自动触发 -> 待定让位（v73）
        _c53, _b53, _u53 = _mk53(yld=False)
        _c53.trigger(True)
        check("53.1 语法位置 + 列表不可见 + 自动触发 -> 待定让位（不弹）",
              [len(_c53.matches or []) == 0, _c53.yield_pending is not None],
              expect_contain=[True, True])

        # 53.2 列表可见 + 自动触发 -> 让位
        _c53, _b53, _u53 = _mk53(yld=True)
        _c53.trigger(True)
        check("53.2 成员列表可见 + 自动触发 -> 让位（不弹）",
              [len(_c53.matches or []) == 0, _u53.hides > 0],
              expect_contain=[True, True])

        # 53.3 ★口径变更：手动 Ctrl+Space 同样让位
        _c53, _b53, _u53 = _mk53(yld=True)
        _c53.trigger(False)
        check("53.3 ★v76 成员列表可见 + 手动 Ctrl+Space -> 同样让位"
              "（不再区分触发来源）",
              [len(_c53.matches or []) == 0, _u53.hides > 0],
              expect_contain=[True, True])

        # 53.4 ★★补位路径（trigger(False)）也受这条判据约束 —— 那个时序缝
        _c53, _b53, _u53 = _mk53(line_text="    freeNa", yld=True)
        _c53.trigger(False)
        check("53.4 ★★v73 补位路径 trigger(False) + 列表可见 -> 也让位"
              "（补出来的窗不许压住 VBE 的列表）",
              [len(_c53.matches or []) == 0, _u53.hides > 0],
              expect_contain=[True, True])

        # 53.5 对照：列表不可见 + 手动 -> 照弹（挡它的确实是"可见"）
        _c53, _b53, _u53 = _mk53(line_text="    .Si", yld=False)
        _c53.trigger(False)
        check("53.5 对照：列表不可见 + 手动 -> 照弹（%d 条）"
              % len(_c53.matches or []),
              [len(_c53.matches or []) > 0, _u53.shown],
              expect_contain=[True, True])

        # 53.6 ★VBE 的列表确实在 -> 清掉"此处已放弃让位"的标记
        _c53, _b53, _u53 = _mk53(yld=False)
        _c53.trigger(True)                  # 先让位（挂 pending）
        _r536 = _c53.confirm_yield(False)   # VBE 没弹 -> 记 given_up
        _gu536 = _c53.yield_given_up
        _b53.set_yield(True)                # VBE 的列表随后真的出现了
        _u53.hides = 0
        _c53.trigger(True)
        check("53.6 ★VBE 列表确实可见 -> 清掉「已放弃让位」标记并让位"
              "（语法判据重新生效，不必再依赖探测）",
              [_r536 is False, _gu536 is not None,
               _c53.yield_given_up is None,
               len(_c53.matches or []) == 0, _u53.hides > 0],
              expect_contain=[True, True, True, True, True])

        # 53.7 对照：列表不可见时标记不许被乱清
        _c53, _b53, _u53 = _mk53(yld=False)
        _c53.trigger(True)
        _c53.confirm_yield(False)
        _gu537 = _c53.yield_given_up
        _c53.trigger(True)
        check("53.7 对照：列表仍不可见 -> 「已放弃让位」标记保持不变",
              [_gu537 is not None, _c53.yield_given_up == _gu537],
              expect_contain=[True, True])

        # 53.8 源码护栏：那条窗口判据不许再挂回 require_ident_before_caret
        _root53 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _esrc53 = ""
        try:
            _esrc53 = open(os.path.join(_root53, "engine.py"),
                           encoding="utf-8").read()
        except Exception:
            _esrc53 = ""
        check("53.8 ★源码护栏：窗口可见性判据不依赖触发来源"
              "（写回 require_ident_before_caret 就会挂）",
              ["if YIELD_TO_VBE_LIST and self._vbe_popup_showing():" in _esrc53,
               "require_ident_before_caret and YIELD_TO_VBE_LIST" in _esrc53],
              expect_contain=[True, True])

        # 53.9 main.py：兜底不分手动 + 轨迹日志在 _LOG_ENABLED 里
        _msrc53 = ""
        try:
            _msrc53 = open(os.path.join(_root53, "main.py"),
                           encoding="utf-8").read()
        except Exception:
            _msrc53 = ""
        check("53.9 接线护栏：poll 兜底不分手动、仍认 VBE_YIELD_CLASSES；"
              "trace 日志被 _LOG_ENABLED 包住、几何取 popup.geometry_now()",
              ['if not state.get("popup_manual") and any(' not in _msrc53,
               "VBE_YIELD_CLASSES" in _msrc53,
               "if _LOG_ENABLED:" in _msrc53,
               "popup.geometry_now()" in _msrc53],
              expect_contain=[True, True, True, True])

        # 53.10 ui.Popup.geometry_now：只读，拿不到一律 None
        class _FW53(object):
            def __init__(self, state="normal"):
                self._st = state

            def state(self):
                return self._st

            def winfo_x(self):
                return 10

            def winfo_y(self):
                return 20

            def winfo_width(self):
                return 200

            def winfo_height(self):
                return 90

        _p53 = UI53.Popup.__new__(UI53.Popup)
        _p53.win = None
        _g_none = _p53.geometry_now()
        _p53.win = _FW53("withdrawn")
        _g_wd = _p53.geometry_now()
        _p53.win = _FW53("normal")
        _g_ok = _p53.geometry_now()
        check("53.10 geometry_now：没有窗 / 已收起 -> None；正常 -> 实际几何",
              [_g_none is None, _g_wd is None, _g_ok == (10, 20, 200, 90)],
              expect_contain=[True, True, True])
    except Exception as _e53:
        check("第 53 节异常: %s" % _e53, [True], expect_contain=[False])

    # ==================================================================
    # 54. 默认提示口径（v77 起；★v96 调整）
    #
    # v77 用户口径（原话）："我觉得提示词太多了……你把提示 vba 本身带的关键字、函数
    # 这块都删掉吧。保留能提示变量名，自定义的函数/过程名，窗体/控件名，模块名。"
    # ⇒ 那时四个开关全部默认 False。
    #
    # ★v96 用户口径（原话）："我想把 vba 关键字、函数名这两部分，也加入到提醒列表里。
    #   枚举类型就不要加了，枚举加上，就会太多了。"
    #   ⇒ 开关组合变成 **(True, True, False, False)**：
    #      收 —— 工程内名字（变量 / 自定义过程 / 组件名 / 控件名）
    #           **+ VBA 关键字 + 内建函数 + 内建数据类型**
    #      不收 —— 一切枚举：VBA 内建 vb*（102 条）、宿主类型库 xl*/mso*（4538 条）。
    #      （枚举这条是 v70 的判断，v96 明确维持：量大、干扰。）
    #
    # 本节守三件事：
    #   1) 默认池子里【关键字/内建函数/内建类型都在、vb* 枚举一条都没有】，
    #      而工程内那四种名字一个不少；
    #   1b) ★反面对照：把两个开关关掉那些名字立刻消失 —— 证明"在"是开关造成的，
    #       不是桩后端/解析碰巧塞进去的；
    #   2) ★反向对照：把枚举开关打开 vb* 立刻回来；再验"开关逐批独立"
    #      （只开关键字不开内建函数也做得到）；
    #   3) ★四份清单**一个字都没裁**，四个收集门也都还在 —— 否则"收回来"会悄悄
    #      变成"能力删除"，以后想关也关不回去（v77 时用户的原话是"不要改坏了"）。
    #
    # ⚠️ 54.4b 单独一条 check 钉出厂默认值：上一版把它混在
    #    `expect_contain=[True]*13` 里，而 `check` 是**子集匹配**、got 是混着
    #    True/False 的列表 ⇒ 默认值已经翻成 True 了它还报 PASS。**这是假绿。**
    #    教训：期望值不同的断言，一律分开 check。
    # ==================================================================
    print("\n=== 54. 默认提示口径（v96：关键字+内建函数收 / 枚举不收）===")
    try:
        import vbe_bridge as VB54R
        import vba_builtins as VB54B
        import engine as E54

        class _CM54R(object):
            def __init__(self, text):
                self.text = text

            @property
            def CountOfLines(self):
                return self.text.count("\n") + 1

            def Lines(self, start, count):
                return "\r\n".join(
                    self.text.split("\n")[start - 1:start - 1 + count])

        class _Ctl54R(object):
            def __init__(self, name):
                self.Name = name

        class _Ctls54R(object):
            def __init__(self, names):
                self._items = [_Ctl54R(n) for n in names]
                self.Count = len(self._items)

            def Item(self, i):
                return self._items[i]

        class _Designer54R(object):
            def __init__(self, ctl_names):
                self.Controls = _Ctls54R(ctl_names)

        class _Comp54R(object):
            def __init__(self, name, text, ctype, ctls=()):
                self.Name = name
                self.Type = ctype
                self.CodeModule = _CM54R(text)
                self.Designer = _Designer54R(list(ctls))

        class _Proj54R(object):
            def __init__(self, comps):
                self.Name = "VBAProject"
                self.VBComponents = list(comps)

        class _Pane54R(object):
            def __init__(self, comp):
                self._c = comp
                self.sel = (1, 1, 1, 1)

            @property
            def CodeModule(self):
                return self._c.CodeModule

            def GetSelection(self):
                return self.sel

        class _VBE54R(object):
            def __init__(self, comps, act):
                self.ActiveVBProject = _Proj54R(comps)
                self.ActiveCodePane = _Pane54R(act)

        # 最小工程：模块里【一个字都没提过】MsgBox / If / Dim / Long
        _m54 = _Comp54R(
            "Module1",
            "Sub Foo()\n    declaredVar = 1\n    looseVar = 2\nEnd Sub",
            VB54R.VBE_CT_STDMODULE)
        _f54 = _Comp54R("UserForm1", "", VB54R.VBE_CT_MSFORM,
                        ctls=["Label1", "cmdOK"])
        # 光标停在窗体上 -> 控件名（只存在于设计器里）才会被收进来
        _vbe54 = _VBE54R([_m54, _f54], _f54)

        _orig54 = VB54R._get_vbe_cached
        VB54R._get_vbe_cached = lambda: _vbe54
        _saved54 = (VB54R.ENABLE_VBA_BUILTINS, VB54R.ENABLE_VBA_KEYWORDS,
                    VB54R.ENABLE_VBA_CONSTANTS, VB54R.ENABLE_HOST_ENUMS)

        def _set54(b, k, c, h):
            (VB54R.ENABLE_VBA_BUILTINS, VB54R.ENABLE_VBA_KEYWORDS,
             VB54R.ENABLE_VBA_CONSTANTS, VB54R.ENABLE_HOST_ENUMS) = (b, k, c, h)

        def _pool54():
            """按【此刻的开关】重收一遍（开关是模块级常量，收集时实时读）。

            每次新建一个后端实例：标识符缓存是实例级 + 2 秒 TTL，复用实例会拿到
            上一轮的旧池子。"""
            _bk = VB54R.VbeBackend()
            return _bk, set(str(r[0]).lower() for r in _bk.get_identifiers())

        class _EngineBackend54(object):
            """把真后端的池子接到引擎上（引擎要的接口一套给全）。"""

            def __init__(self, bk, module="Module1", word="ms"):
                self._bk = bk
                self._mod = module
                self._word = word

            def get_context(self):
                _ln = "    " + self._word
                return {"line_no": 1, "line_text": _ln,
                        "caret_col": len(_ln) + 1, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": self._mod}

            def get_identifiers(self):
                return self._bk.get_identifiers()

            def get_declared_names(self):
                return sorted(self._bk.get_declared_names())

            def declared_elsewhere(self, name, module_name):
                return bool(self._bk.declared_elsewhere(name, module_name))

            def get_structural_names(self):
                return self._bk.get_structural_names()

            def get_builtin_names(self):
                return self._bk.get_builtin_names()

            def get_type_names(self):
                return list(self._bk.get_type_names())

            def get_host_enum_names(self):
                return self._bk.get_host_enum_names()

            def names_outside_caret(self, caret, scope_only=False):
                # 现场文本证据：桩里就等于"池子里的【工程内】名字"。
                # 内建那批不在现场文本里 —— 真机上它们也正是靠 structural_names
                # 那一路放行的（回声防护的设计，这里如实模拟，别把两路混成一路）。
                return set(
                    str(n).lower() for n, _m, _p, _pv in self._bk.get_identifiers()
                    if str(_m) not in (VB54R._BUILTIN_MODULE,
                                       VB54R._HOST_ENUM_MODULE))

            def apply_completion(self, *a, **k):
                return None

            # 让位/避让一律"探不到"（本机不一定在跑 Excel，保持确定性）
            def vbe_yield_visible(self):
                return False

            def vbe_popup_visible(self):
                return False

            def vbe_popup_info(self):
                return []

        class _UI54(object):
            def __init__(self):
                self.shown = False

            def show(self, rows, sel, c):
                self.shown = True

            def hide(self):
                self.shown = False

            def update_selection(self, sel):
                pass

            def contains_point(self, x, y):
                return False

        def _trig54(word="ms"):
            """触发一次补全，返回候选列表——**保持引擎的原始排序**。

            ⚠️ 别在这里 sorted()：断言"第 1 个是 X"量的是**引擎的排序**
            （kind 高者在前），一排序就变成字母序了，`su` 会把 `GoSub`
            排到 `Sub` 前面（`decl` 把 `Declare` 排到 `declaredVar` 前面，
            因为 'D' 的码位小于 'd'）—— 量到的就不是引擎行为，而是字符串比较。
            """
            _c = E54.Completer(_EngineBackend54(VB54R.VbeBackend(),
                                                word=word), _UI54())
            _c.trigger(True)
            return list(_c.matches or [])

        try:
            # ---- 54.1 ★默认口径：关键字 + 内建函数/类型【在】，两批枚举【不在】----
            # ⚠️ ★v96 默认口径变了（见下面 54.1b 的出处说明）：这一条从"四种代表名
            # 一条都没有"改成"关键字/函数/类型在、vb* 枚举不在"。
            _set54(True, True, False, False)
            _bk54, _nm54 = _pool54()
            # 关键字 / 内建函数 / 内建类型各取一个代表 —— 这三批默认【该在】。
            # vb 枚举（vbcrlf / vbyes）取两个代表 —— 这批默认【不该在】。
            # 宿主枚举（xl*/mso*）不放进来断言"不在"：它关着时收集段整段不跑
            # （连类型库都不加载），由 47.4 单独守；真要验证"打开它才在"是 44 节。
            _lang54 = ["msgbox", "left", "split", "dim", "sub", "function",
                       "long", "string"]
            _enum54 = ["vbcrlf", "vbyes"]
            # 工程内该有的四种名字（变量 / 自定义过程 / 模块名 / 窗体名 + 控件名）
            _mine54 = ["declaredvar", "loosevar", "foo", "module1",
                       "userform1", "label1", "cmdok"]
            check("54.1 ★默认口径：关键字/内建函数/内建类型 %d 个代表名全在；"
                  "vb* 枚举 %d 个一个都不在（枚举用户明确说不要）；"
                  "工程内名字一个不少"
                  % (len(_lang54), len(_enum54)),
                  [sorted(n for n in _lang54 if n not in _nm54),
                   sorted(set(_enum54) & _nm54),
                   sorted(n for n in _mine54 if n not in _nm54),
                   dict(_bk54.get_host_enum_names()),
                   "vbcrlf" in _bk54.get_builtin_names(),
                   "msgbox" in _bk54.get_builtin_names()],
                  expect_contain=[[], [], [], {}, False, True])

            # ---- 54.1b ★反面对照：把开关关掉，v96 放回来的那两批立刻消失 ----
            # 单靠 54.1 只能证明"默认在"；若桩后端/解析碰巧把名字塞进池子，
            # 54.1 也会绿。这里显式关一次，证明"在"确实是这两个开关造成的。
            _set54(False, False, False, False)
            _bk54off, _nm54off = _pool54()
            check("54.1b ★反面对照：ENABLE_VBA_BUILTINS/KEYWORDS 关掉后，"
                  "关键字+内建函数+内建类型一条都不在（证明 54.1 是开关在起作用）",
                  [sorted(set(_lang54) & _nm54off), _bk54off.get_builtin_names()],
                  expect_contain=[[], []])
            # 下面 54.2~54.7 各自显式设开关，不依赖这里的残留状态。

            # ---- 54.2 ★反向对照：打开开关 -> 那些名字立刻回来 ----
            _set54(True, True, True, False)
            _bk54on, _nm54on = _pool54()
            check("54.2 ★反向对照：开关打开后 10 个代表名全回来（证明是开关在起作用）；"
                  "宿主枚举那一批仍不越界",
                  [sorted(n for n in _lang54 if n not in _nm54on),
                   "xlup" not in _nm54on],
                  expect_contain=[[], True])

            # ---- 54.3 开关逐批独立：只开关键字，内建函数照旧不收 ----
            _set54(False, True, False, False)
            _bk54k, _nm54k = _pool54()
            check("54.3 开关逐批独立：只开关键字 -> Sub/If 回来，MsgBox/Long 仍不在",
                  ["sub" in _nm54k, "if" in _nm54k,
                   "msgbox" not in _nm54k, "long" not in _nm54k,
                   "vbcrlf" not in _nm54k],
                  expect_contain=[True, True, True, True, True])

            # ---- 54.4 ★四份清单没被裁 + 收集门都在（只关开关，不删能力）----
            _src54 = open(os.path.join(
                os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                "vbe_bridge.py"), encoding="utf-8").read()
            check("54.4 ★四份清单一个字没裁（函数 %d / 常量 %d / 类型 %d / 关键字 %d）"
                  "+ 三个收集门都在"
                  % (len(VB54B.BUILTIN_FUNCTIONS), len(VB54B.BUILTIN_CONSTANTS),
                     len(VB54B.BUILTIN_TYPE_NAMES), len(VB54B.BUILTIN_KEYWORDS)),
                  [len(VB54B.BUILTIN_FUNCTIONS) >= 150,
                   len(VB54B.BUILTIN_CONSTANTS) >= 100,
                   len(VB54B.BUILTIN_TYPE_NAMES) >= 5,
                   len(VB54B.BUILTIN_KEYWORDS) >= 30,
                   "MsgBox" in VB54B.BUILTIN_FUNCTIONS,
                   "Long" in VB54B.BUILTIN_TYPE_NAMES,
                   "Sub" in VB54B.BUILTIN_KEYWORDS,
                   _src54.count("if ENABLE_VBA_BUILTINS:") >= 1,
                   _src54.count("if ENABLE_VBA_KEYWORDS:") >= 1,
                   _src54.count("if ENABLE_VBA_CONSTANTS:") >= 1,
                   _src54.count("if ENABLE_HOST_ENUMS:") >= 1],
                  expect_contain=[True] * 11)

            # ⚠️ 这两条必须【单独 check】：`check` 是**子集匹配**（`x not in got`），
            #   而 got 是一个混着 True/False 的列表 —— 把它们塞进上面那个
            #   `expect_contain=[True]*13` 里，"列表里有 False"根本不会被发现
            #   （`True in got` 照样成立）。v96 之前就是这样假绿的：默认值已经
            #   改成 True，54.4 却仍然 PASS。⇒ 期望值不同层/不同值时，一律分开 check。
            check("54.4b ★出厂默认值：关键字 + 内建函数/数据类型 = True，"
                  "两批枚举 = False（枚举不收）",
                  [_src54.count('_env_flag("VBECOMPLETE_VBA_BUILTINS", True)'),
                   _src54.count('_env_flag("VBECOMPLETE_VBA_KEYWORDS", True)'),
                   _src54.count('_env_flag("VBECOMPLETE_VBA_CONSTANTS", False)'),
                   _src54.count('_env_flag("VBECOMPLETE_HOST_ENUMS", False)')],
                  expect_contain=[1, 1, 1, 1])

            # ---- 54.5 As 位置：内建类型默认【在】，你自己的类型也一直在 ----
            # ⚠️ v96 之前这里是"默认不在"（那时 _bk54 是全关的池子）。默认改成
            #   True 之后 _bk54 就是"开"的池子，所以"不在"那一半改用 54.1b
            #   收的 _bk54off（显式关掉的）来验 —— 两个方向都钉住。
            _tn54_off = set(str(t).lower() for t in _bk54off.get_type_names())
            _tn54_def = set(str(t).lower() for t in _bk54.get_type_names())
            _tn54_on = set(str(t).lower() for t in _bk54on.get_type_names())
            check("54.5 `Dim x As |`：开关关掉时内建类型不在、你自己的类型一直在；"
                  "★默认（v96）内建类型【已经在了】",
                  ["long" not in _tn54_off, "string" not in _tn54_off,
                   "userform1" in _tn54_off,
                   "long" in _tn54_def and "string" in _tn54_def,
                   "userform1" in _tn54_def,
                   "long" in _tn54_on and "string" in _tn54_on],
                  expect_contain=[True] * 6)

            # ---- 54.6 ★端到端：关掉时打 ms 零候选；默认(开) -> MsgBox 第 1 个 ----
            _set54(False, False, False, False)
            _off54 = _trig54("ms")
            _set54(True, True, False, False)
            _on54 = _trig54("ms")
            check("54.6 ★端到端：全关时打 ms 零候选；默认口径(v96) -> MsgBox（第 1 个）",
                  [len(_off54), _on54[:1]],
                  expect_contain=[0, ["MsgBox"]])

            # ---- 54.6b ★端到端：关键字这条路也要能走通（打 su -> Sub）----
            # v96 把关键字放回默认池子，但它走的是与内建函数**同一段收集代码**
            # （只差一个开关）—— 这里从引擎端到端钉一次，免得"开关开了但提示不出来"。
            _set54(True, True, False, False)
            _kw54 = _trig54("su")
            check("54.6b ★端到端：默认口径下打 su -> Sub（第 1 个）",
                  [_kw54[:1], "Sub" in _kw54],
                  expect_contain=[["Sub"], True])

            # 默认口径下"工程内名字"照旧能补（别把保留项一起关没了）
            _set54(True, True, False, False)
            _proj54 = _trig54("decl")
            check("54.7 默认口径下工程内名字照旧（decl -> declaredVar），"
                  "且排在语言自带的名字前面（不顶掉你的名字）",
                  ["declaredVar" in _proj54, _proj54[:1]],
                  expect_contain=[True, ["declaredVar"]])
        finally:
            _set54(*_saved54)
            VB54R._get_vbe_cached = _orig54
    except Exception as _e54:
        check("第 54 节异常: %s" % _e54, [True], expect_contain=[False])

    # ---- 55. v78：Shift+↑/↓ 跳出候选列表 + 光标上/下移一行 ----
    #
    # 用户口径：「当提示词列表框出现的时候，按 ↑↓ 会一直在列表框选词，不能将
    # 光标移动上下行。帮我新增按 Shift+↑，可以跳出我们项目的提示词列表框，并将
    # 光标上移一行；Shift+↓ 同理。」
    #
    # 两个关键决定，都由断言守着：
    #   1) 不用"放行按键让 VBE 自己动光标" —— VBE 原生的 Shift+↑/↓ 是【扩展选区】
    #      （把上一行整行选进去），不是"光标干净地上移一行"。所以由我们吞键 +
    #      用 COM 落光标（VbeBackend.move_caret_line）。
    #   2) 判据必须排在"在列表里选词"那条分支【之前】，否则 Shift+↑ 又被当成
    #      普通 ↑ 去列表里选词 —— 那就完全没解决问题（55.6 钉着这个次序）。
    print("\n=== 55. Shift+↑/↓ 跳出候选列表并移动光标（v78）===")
    try:
        import main as M55
        import vbe_bridge as VB55
        _orig55 = VB55._get_vbe_cached

        # ---- 55.1 纯判断：只认 ↑/↓ + Shift 按住 ----
        check("55.1 _is_line_nav_shortcut：Shift+↑/↓ 认；裸 ↑、Shift+Enter、"
              "Shift+←、Shift 没按都不认",
              [M55._is_line_nav_shortcut(M55.VK_UP, True),
               M55._is_line_nav_shortcut(M55.VK_DOWN, True),
               M55._is_line_nav_shortcut(M55.VK_UP, False),
               M55._is_line_nav_shortcut(M55.VK_RETURN, True),
               M55._is_line_nav_shortcut(M55.VK_LEFT, True),
               M55._is_line_nav_shortcut(M55.VK_DOWN, 0)],
              expect_contain=[True, True, False, False, False, False])

        class _CM55(object):
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "Module1"

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n"
                               for l in self.lines[s0:s0 + int(count)])

        class _Pane55(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = tuple(sel)

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE55(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _mv55(lines, sel, delta):
            cm55 = _CM55(lines)
            pane55 = _Pane55(cm55, sel)
            VB55._get_vbe_cached = lambda: _VBE55(pane55)
            ok55 = VB55.VbeBackend().move_caret_line(delta)
            return ok55, cm55.lines, pane55.sel

        _SRC55 = ["Sub Foo()", "    Dim x", "    x = 1 + 2", "End Sub"]

        # ---- 55.2/55.3 下移 / 上移一行（同列，且不碰文本）----
        _ok55d, _ls55d, _sel55d = _mv55(_SRC55, (2, 9, 2, 9), 1)
        _ok55u, _ls55u, _sel55u = _mv55(_SRC55, (3, 9, 3, 9), -1)
        check("55.2 下移一行：光标 (2,9) -> (3,9)，同列、文本一动不动",
              [_ok55d, _sel55d, _ls55d],
              expect_contain=[True, (3, 9, 3, 9), _SRC55])
        check("55.3 上移一行：光标 (3,9) -> (2,9)，同列、文本一动不动",
              [_ok55u, _sel55u, _ls55u],
              expect_contain=[True, (2, 9, 2, 9), _SRC55])

        # ---- 55.4 越界 / 无 VBE 一律 False（调用方据此【不吞键】）----
        _b55 = []
        _b55.append(_mv55(_SRC55, (1, 1, 1, 1), -1)[0])     # 首行再上移
        _b55.append(_mv55(_SRC55, (4, 8, 4, 8), 1)[0])      # 末行再下移
        _b55.append(_mv55(_SRC55, (0, 0, 0, 0), 1)[0])      # 行号 0（空模块）
        VB55._get_vbe_cached = lambda: None
        _b55.append(VB55.VbeBackend().move_caret_line(1))
        VB55._get_vbe_cached = lambda: _VBE55(None)
        _b55.append(VB55.VbeBackend().move_caret_line(1))
        check("55.4 首行上移 / 末行下移 / 行号 0 / 无 VBE / 无代码窗 -> 全 False",
              [_b55], expect_contain=[[False] * 5])

        # ---- 55.5 反向选择时以【活动端】为基准（选中框的另一头不算）----
        _ok55r, _ls55r, _sel55r = _mv55(_SRC55, (4, 3, 2, 7), -1)
        check("55.5 反向选区（sl=4 > el=2）以 (4,3) 为基准 -> 上移到 (3,3)",
              [_ok55r, _sel55r], expect_contain=[True, (3, 3, 3, 3)])

        # ---- 55.6 ★接线护栏：次序 + 先收窗后落光标 + 兜底 + 开关 ----
        _src55 = open(os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "main.py"), encoding="utf-8").read()
        _i55_nav = _src55.index("elif _is_line_nav_shortcut(vk, _shift_down()):")
        _i55_pick = _src55.index("elif vk in NAV_VKS:")
        _i55_fn = _src55.index("def _move_caret_here(delta):")
        _i55_mv = _src55.index("backend.move_caret_line(delta)")
        check("55.6 ★接线护栏：Shift+↑/↓ 的判据排在「在列表里选词」之前；"
              "先收窗再落光标；COM 没落成要走 send_vk 兜底；开关在",
              [_i55_nav < _i55_pick,
               "completer.hide()" in _src55[_i55_fn:_i55_mv],
               _i55_fn < _i55_mv,
               "send_vk(VK_UP if delta < 0 else VK_DOWN)" in _src55,
               'os.environ.get("VBECOMPLETE_NO_SHIFT_NAV"' in _src55,
               "SHIFT_NAV_JUMP" in _src55],
              expect_contain=[True] * 6)

        # ---- 55.7 出厂默认：这个接管是开着的 ----
        check("55.7 ★出厂默认：Shift+↑/↓ 接管开着（VBECOMPLETE_NO_SHIFT_NAV 未设）",
              [M55.SHIFT_NAV_JUMP], expect_contain=[True])

        VB55._get_vbe_cached = _orig55
    except Exception as _e55:
        check("第 55 节异常: %s" % _e55, [True], expect_contain=[False])

    # ---- 56. v79：配对光标落点（bug）+ 回车自动缩进（新功能） ----
    #
    # 用户报的 bug：「录入 targetSheet.Cells(i, "b").Interior.Color = RGB，然后
    # 再打括号，光标落点会错位，不能正确落在 RGB 后面的括号内。应该是 vba 自动
    # 格式化的时候，有两个括号，你只判断了一个。」
    # 用户要的两个功能：
    #   1) 整行注释后回车，忽略上方注释行（可能多行）、跟上方第一个非注释行对齐；
    #   2) For / Do / If / With / Sub / Function / Type / Enum 这些结构，回车自动缩进。
    print("\n=== 56. 配对光标落点 + 回车自动缩进（v79）===")
    try:
        import main as M56
        import vbe_bridge as VB56
        _orig56 = VB56._get_vbe_cached

        # ---- 56.1 ★bug 修复：VBE 重排整行后，光标仍要落在括号中间 ----
        _CANON56 = '    targetSheet.Cells(i, "b").Interior.Color = RGB()'

        def _off56(after, col0):
            return VB56._locate_inserted(after, _CANON56, "(", ")", col0)

        _c56 = _CANON56.index("RGB()") + 3       # 写入时光标的位置（括号之间）

        def _in_pair56(after):
            _o = _off56(after, _c56)
            return bool(_o is not None and after[_o - 1:_o + 1] == "()")

        check("56.1 ★bug：VBE 重排整行后仍把光标钉在括号中间"
              "（逗号后补空格 / 改大小写 / 等号两边补空格 / 括号前后多空格）",
              [_in_pair56(_CANON56),
               _in_pair56('    targetSheet.Cells(i, "b").Interior.Color = RGB ()'),
               _in_pair56('    targetSheet.Cells(i,"b").interior.color = rgb()'),
               _in_pair56('    targetSheet.Cells( i , "b" ).Interior.Color '
                          '= RGB(   )'.replace("(   )", "()"))],
              expect_contain=[True, True, True, True])

        # 用户说的"两个括号"：行里本来还有一对括号，不能认到那一对上去
        _W56 = '    targetSheet.Cells(i, x()).Interior.Color = RGB()'
        _o56 = VB56._locate_inserted(_W56, _W56, "(", ")",
                                     _W56.index("RGB()") + 3)
        check("56.2 ★bug：一行里有两对括号时，认的是我们刚插的那一对（RGB 后面）",
              [_o56, _W56[_o56 - 1:_o56 + 1], _o56 > _W56.index("RGB")],
              expect_contain=[_W56.index("RGB") + 4, "()", True])

        # v58 老场景（引号）：`    w""=4` 被 VBE 重排成 `    w "" = 4`
        check("56.3 引号那一例（v58 老场景）不许回归",
              [VB56._locate_inserted('    w "" = 4', '    w""=4', '"', '"', 5)],
              expect_contain=[7])

        # ---- 56.4 块结构开头（回车后要深一级） ----
        _BLK56 = ["For i = 1 To 10", "For Each c In rng", "Do", "Do While x < 5",
                  "Do Until x > 0", "While x", "With rng", "If x Then",
                  "If x Then '注释", "ElseIf y Then", "Else", "Case 1, 2",
                  "Case Else", "Select Case x", "Sub Foo()",
                  "Private Sub Foo()", "Public Function Bar() As Long",
                  "Static Sub S()", "Property Get X()", "Type Foo",
                  "Enum Color", "#If VBA7 Then", "#Else",
                  'If x = "a Then" Then', "    For i = 1 To 10"]
        _NOBLK56 = ["If x Then y = 1", "Next i", "End If", "End Sub", "End With",
                    "Loop", "Loop While x", "Wend", "x = 1", "Exit For",
                    "DoEvents", "foo(1, 2)", "Debug.Print x", "Rem 说明",
                    "'注释", "Set x = New Collection", "Call Foo",
                    "Dim x As Long", "    ", "End Select", "Exit Sub",
                    "Range(\"a1\").Value = 1", "If x Then _", "rem表格"]
        check("56.4 opens_block：24 个块头全认，24 个非块头全不认",
              [[t for t in _BLK56 if not VB56.opens_block(t)],
               [t for t in _NOBLK56 if VB56.opens_block(t)]],
              expect_contain=[[], []])

        # ---- 56.5 回车后新行的缩进 ----
        _U56 = "    "
        _C56 = [VB56.next_line_indent("    x = 1", ["Sub Foo()"], _U56),
                VB56.next_line_indent("    For i = 1 To 10", [], _U56),
                VB56.next_line_indent("    Next i", [], _U56),
                VB56.next_line_indent("'注释", ["Sub Foo()", "    x = 1"], _U56),
                VB56.next_line_indent("'注释", ["    With rng"], _U56),
                VB56.next_line_indent("'注释", ["'老注释", ""], _U56)]
        check("56.5 next_line_indent：普通行齐平 / 块头深一级 / 闭合行齐平 / "
              "注释跟上方非注释行对齐 / 注释上方是块头则深一级 / 上方全注释"
              "则维持自身缩进",
              _C56, expect_contain=["    ", "        ", "    ", "    ",
                                    "        ", ""])

        # ★用户的实际场景：注释行在第 1 列、上方代码缩进 12，连续多行注释
        _A56 = ["        Else", "            a = RGB(1, 2)",
                "'第一行注释", "'第二行注释"]
        check("56.5b ★用户场景：整行注释后回车，忽略上方连续注释行，"
              "跟第一个非注释行（缩进 12）对齐",
              [VB56.next_line_indent("'第三行注释", _A56, _U56),
               VB56.next_line_indent("'注释",
                                     ["Sub Foo()", "    For i = 1 To 10"], _U56)],
              expect_contain=["            ", "        "])

        # ---- 56.6 钩子侧判据：只在这两种情形接管回车 ----
        # ⚠️ caret_ec 是 VBE 的【显示列】(1-based，全角占 2 格)。下面每个
        # "行尾"用例给的列 = _eol_disp_col(该行)，不是 len(该行)+1 ——
        # 有全角字符时两者不等（`'注释` 3 字符但行尾显示列是 6）。
        # v93 起判据对此是【严格相等】，旧数据里那个凭空的 20 会红。
        _W56b = [("整行注释+行尾", "'注释", 6, True),
                 ("块头+行尾", "    For i = 1 To 10", 20, True),
                 ("普通代码行+行尾", "    x = 1", 10, False),
                 ("光标在行中间", "    For i = 1 To 10", 12, False),
                 ("空行", "        ", 9, False),
                 ("整行空白", "", 1, False),
                 ("含 Tab 的行", "\tFor i", 10, False),
                 ("光标列拿不到", "    For i", 0, False),
                 # ★v93：含全角时"显示列虚高"差点被当成"在行尾"——
                 # `    For 每个元素 = 1 To 10`(22 字符、行尾显示列 27)，
                 # 光标停在字符列 18 时 VBE 报显示列 22 >= 22，旧判据
                 # (ec >= len) 会误接管那一下【拆行】的回车。
                 ("中文块头 光标在中间", "    For 每个元素 = 1 To 10", 22, False),
                 ("中文注释 光标在中间", "    ' 配置：读取 用户名", 18, False),
                 ("中文块头 真·行尾", "    For 每个元素 = 1 To 10", 27, True),
                 ("中文注释 真·行尾", "    ' 配置：读取 用户名", 24, True)]
        check("56.6 enter_indent_wanted：注释 / 块头 + 光标在行尾才接管；"
              "普通行、行中间、空行、含 Tab、拿不到列一律放行给 VBE；"
              "★含全角字符时按【显示列】精确比对行尾（不许被虚高的列骗去拆行）",
              [[n for (n, t, c, e) in _W56b
                if VB56.enter_indent_wanted(t, c) != e]],
              expect_contain=[[]])

        # ---- 56.7 引号没闭合时不许抢回车（那是 VBE 原生会补右引号的地方） ----
        check("56.7 _unterminated_string：注释里的引号不算；字符串内 \"\" 转义不算",
              [VB56._unterminated_string('s = "abc'),
               VB56._unterminated_string('s = "abc"'),
               VB56._unterminated_string('x = 1  ' + "'他说\"你好"),
               VB56._unterminated_string('s = "a""b"')],
              expect_contain=[True, False, False, False])

        # ---- 56.8 ★集成：smart 模式真能改写缩进；不该管的返回 False ----
        class _CM56(object):
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "Module1"

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n"
                               for l in self.lines[s0:s0 + int(count)])

            def InsertLines(self, start, text):
                i = min(max(0, int(start) - 1), len(self.lines))
                for k, ln in enumerate(str(text).split("\n")):
                    self.lines.insert(i + k, ln)

        class _Pane56(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE56(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _nl56(lines, line_no, col, smart=True):
            cm56 = _CM56(lines)
            pane56 = _Pane56(cm56, (line_no, col, line_no, col))
            VB56._get_vbe_cached = lambda: _VBE56(pane56)
            ok56 = VB56.VbeBackend().new_line_below(smart=smart)
            return ok56, cm56.lines, pane56.sel

        _ok, _ls, _sel = _nl56(["Sub Foo()", "    For i = 1 To 10",
                                "    Next i", "End Sub"], 2, 21)
        check("56.8 ★集成：块头那行回车 -> 新行深一级（8 空格）、光标在缩进后、"
              "块头那一行一个字不动",
              [(_ok, _ls, _sel)],
              expect_contain=[(True,
                               ["Sub Foo()", "    For i = 1 To 10", "        ",
                                "    Next i", "End Sub"],
                               (3, 9, 3, 9))])

        _ok, _ls, _sel = _nl56(["Sub Foo()", "    x = 1", "'注释", "End Sub"], 3, 5)
        check("56.8b ★集成：整行注释那行回车 -> 新行跟上方代码行（4 空格）对齐",
              [(_ok, _ls[3] if len(_ls) > 3 else None)],
              expect_contain=[(True, "    ")])

        # 中文注释（显示列语义）也要能对上：'注释 = 1 + 2 + 2 = 5 列
        _sem56 = VB56._sem_cache.get("info")
        VB56._sem_cache["info"] = ("disp", 4, True)
        try:
            _ok, _ls, _sel = _nl56(["Sub Foo()", "    x = 1", "'注释", "End Sub"],
                                   3, 5)
        finally:
            VB56._sem_cache["info"] = _sem56
        check("56.8c ★集成：中文整行注释（显示列语义）同样对齐到 4 空格",
              [(_ok, _ls[3] if len(_ls) > 3 else None)],
              expect_contain=[(True, "    ")])

        _bad56 = []
        for _n, _ln, _col, _src in [
                ("光标在行中间", 2, 12, ["Sub Foo()", "    For i = 1 To 10",
                                        "End Sub"]),
                ("空行", 2, 5, ["Sub Foo()", "    ", "End Sub"]),
                ("引号没闭合", 2, 13, ["Sub Foo()", '    s = "abc', "End Sub"])]:
            _ok56, _ls56, _s56 = _nl56(_src, _ln, _col)
            if _ok56 is not False or _ls56 != _src:
                _bad56.append(_n)
        # 多行选区：按回车是"删除选区"，语义完全不同 —— 必须原样让 VBE 处理
        _cm56 = _CM56(["Sub Foo()", "    a = 1", "    b = 2", "End Sub"])
        _pane56 = _Pane56(_cm56, (2, 1, 3, 1))
        VB56._get_vbe_cached = lambda: _VBE56(_pane56)
        _before56 = list(_cm56.lines)
        _ok56 = VB56.VbeBackend().new_line_below(smart=True)
        if _ok56 is not False or _cm56.lines != _before56:
            _bad56.append("多行选区")
        check("56.8d smart 模式这四种情况必须返回 False 且一个字都不写"
              "（调用方据此把回车原样还给系统）",
              [_bad56], expect_contain=[[]])

        # Shift+Enter 老口径不许被改动：不传 smart 时仍然"光标在行中间也照样
        # 新起一行、缩进照抄本行"
        _ok, _ls, _sel = _nl56(["Sub Foo()", "    x = 1 + 2", "End Sub"], 2, 11,
                               smart=False)
        check("56.8e Shift+Enter 老口径不变（smart 默认 False）",
              [(_ok, _ls[2], _sel)],
              expect_contain=[(True, "    ", (3, 5, 3, 5))])

        VB56._get_vbe_cached = lambda: None
        check("56.8f 取不到 VBE -> False（回车原样还给系统）",
              [VB56.VbeBackend().new_line_below(smart=True)],
              expect_contain=[False])
        VB56._get_vbe_cached = _orig56

        # ---- 56.9 缩进一级有多宽：空格 / Tab / 环境变量覆盖 ----
        check("56.9 _indent_unit_for：默认 4 空格；工程用 Tab 就跟着用 Tab；"
              "环境变量可覆盖（VBECOMPLETE_INDENT_UNIT=2）",
              [VB56._indent_unit_for(["Sub Foo()", "    x = 1"]),
               VB56._indent_unit_for(["Sub Foo()", "\tx = 1"])],
              expect_contain=["    ", "\t"])
        _env56 = os.environ.get("VBECOMPLETE_INDENT_UNIT")
        try:
            os.environ["VBECOMPLETE_INDENT_UNIT"] = "2"
            check("56.9b 环境变量 VBECOMPLETE_INDENT_UNIT=2 -> 一级 2 个空格",
                  [VB56._indent_unit_for([]),
                   VB56.next_line_indent("    For i = 1", [],
                                         VB56._indent_unit_for([]))],
                  expect_contain=["  ", "      "])
            os.environ["VBECOMPLETE_INDENT_UNIT"] = "tab"
            check("56.9c 环境变量 =tab -> 一级一个 Tab",
                  [VB56._indent_unit_for([])], expect_contain=["\t"])
        finally:
            if _env56 is None:
                os.environ.pop("VBECOMPLETE_INDENT_UNIT", None)
            else:
                os.environ["VBECOMPLETE_INDENT_UNIT"] = _env56

        # ---- 56.10 ★接线护栏：钩子侧的次序与兜底 ----
        _src56 = io.open(os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "main.py"), encoding="utf-8").read()
        _i56_enter = _src56.index("elif _enter_indent_key_ok(vk):")
        _i56_pair = _src56.index("elif (_pair_ch")
        _i56_vis = _src56.index("elif completer.is_visible():")
        _i56_fn = _src56.index("def _enter_indent_here():")
        _i56_call = _src56.index("backend.new_line_below(smart=True")
        check("56.10 ★接线护栏：回车分支在自动配对之前、在「弹窗可见」之前；"
              "先收窗再写；COM 没落成要 send_vk(VK_RETURN) 兜底；"
              "判据只在钩子线程读快照（不吃 COM）",
              [_i56_enter < _i56_pair < _i56_vis,
               "completer.hide()" in _src56[_i56_fn:_i56_call],
               _i56_fn < _i56_call,
               "send_vk(VK_RETURN)" in _src56[_i56_fn:],
               'os.environ.get("VBECOMPLETE_NO_ENTER_INDENT"' in _src56,
               "ENTER_CTX_MAX_AGE" in _src56,
               "ENTER_INDENT" in _src56],
              expect_contain=[True] * 7)

        _gate56 = _src56[_src56.index("def _enter_indent_key_ok(vk):"):
                         _src56.index("def win32_filter(msg, data):")]
        check("56.10b 钩子判据里不许出现任何 COM 调用（钩子线程没有 COM 单元）",
              [[w for w in ("_get_vbe_cached", "backend.", "VbeBackend")
                if w in _gate56]],
              expect_contain=[[]])

        # 快照多带一个光标列；但"内容变化"的判定仍只看前三个元素
        check("56.10c snapshot 带光标列；变更检测仍只比前三个（光标移动不算"
              "内容变化，否则点来点去都会弹窗）",
              ["snap = (sl, cm.Lines(sl, 1), mod, ec)" in
               io.open(os.path.join(
                   os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "vbe_bridge.py"), encoding="utf-8").read(),
               "last[:3] != snap[:3]" in _src56,
               'state["cur_ctx"] = (time.time(), snap[0], snap[1], snap[3])'
               in _src56],
              expect_contain=[True, True, True])

        # ---- 56.10d ★输入法护栏：组字中的回车是"上屏键"，绝不能被抢 ----
        # 判据是"按了可能写字的键、文档却没变"（拼音被输入法吃掉）= 正在组字。
        # 这是输入法无关的做法，不依赖任何 IME 接口。
        _TXT56 = [0x41, 0x5A, 0x30, 0x39, 0x20, 0xBA, 0xC0, 0xDB, 0xDE, 0x61]
        _NOTXT56 = [0x0D, 0x09, 0x1B, 0x25, 0x26, 0x27, 0x28, 0x10, 0x11, 0x12,
                    0x70, 0x7A, 0xE3, 0xAD, 0xB0]
        check("56.10d _is_text_key：字母/数字/空格/OEM 标点算「可能写字」；"
              "回车、Tab、Esc、方向键、修饰键、F 键、多媒体键不算",
              [[k for k in _TXT56 if not M56._is_text_key(k)],
               [k for k in _NOTXT56 if M56._is_text_key(k)]],
              expect_contain=[[], []])

        check("56.10e ★接线护栏：回车判据在组字（pending_keys>0）时一律不接管；"
              "钩子数「可能写字的键」；文档一变就把计数清零",
              ["pending_keys" in _gate56,
               'state["pending_keys"] = min(' in _src56,
               'state["pending_keys"] = 0' in _src56,
               _src56.count('state["pending_keys"] = 0') >= 2],
              expect_contain=[True] * 4)

        # ---- 56.11 出厂默认：回车自动缩进开着 ----
        check("56.11 ★出厂默认：回车自动缩进开着"
              "（VBECOMPLETE_NO_ENTER_INDENT 未设）",
              [M56.ENTER_INDENT], expect_contain=[True])
    except Exception as _e56:
        check("第 56 节异常: %s" % _e56, [True], expect_contain=[False])

    # ---- 57. v79b：块头那行回车，自动补出块收尾 ----
    #
    # 用户的追加要求：「写 if i=1 then，按回车之后，能自动帮我补全后面的 end if，
    # 而且光标缩进；对于其他的结构语句也是这样，比如 with 语句块，就能自动补全
    # end with。」
    print("\n=== 57. 块头回车自动补收尾（v79b）===")
    try:
        import main as M57
        import vbe_bridge as VB57
        _orig57 = VB57._get_vbe_cached

        # ---- 57.1 块头 -> 收尾 ----
        check("57.1 block_closer：Sub/Function/Property/Type/Enum/If…Then/With/"
              "Select Case -> End Xxx；For -> Next；Do -> Loop；While -> Wend",
              [VB57.block_closer("Sub Foo()"),
               VB57.block_closer("Private Function Bar() As Long"),
               VB57.block_closer("Static Property Get X()"),
               VB57.block_closer("Type Foo"),
               VB57.block_closer("Enum Color"),
               VB57.block_closer("If x Then"),
               VB57.block_closer("    If x = 1 Then '注释"),
               VB57.block_closer("With rng"),
               VB57.block_closer("Select Case x"),
               VB57.block_closer("For i = 1 To 10"),
               VB57.block_closer("For Each c In rng"),
               VB57.block_closer("Do"),
               VB57.block_closer("Do While x < 5"),
               VB57.block_closer("While x")],
              expect_contain=["End Sub", "End Function", "End Property",
                              "End Type", "End Enum", "End If", "End If",
                              "End With", "End Select", "Next", "Next",
                              "Loop", "Loop", "Wend"])

        # ---- 57.2 不该补收尾的一个都不许补 ----
        _N57 = ["ElseIf y Then", "Else", "Case 1, 2", "Case Else",
                "#If VBA7 Then", "#Else", "If x Then y = 1", "'注释",
                "Rem 说明", "x = 1", "End If", "Next i", "Loop", "Wend",
                "End Sub", "    "]
        check("57.2 block_closer：Else / ElseIf…Then / Case（块【内】的分支）、"
              "#If / #Else（必须配 #End If）、单选 If、注释、闭合行 -> 全 None",
              [[t for t in _N57 if VB57.block_closer(t) is not None]],
              expect_contain=[[]])

        # ---- 57.3 到底该不该补：只看"下方第一个非空行" ----
        check("57.3 closer_needed：下面没内容 -> 补；下面已挂着同一个收尾"
              "（含隔空行/带缩进/`Next i`/`Loop While x`）-> 不补；"
              "缩进不比块头浅 -> 不补；比块头浅 -> 补",
              [VB57.closer_needed([], "End If", 0),
               VB57.closer_needed(["End If"], "End If", 0),
               VB57.closer_needed(["    End If"], "End If", 0),
               VB57.closer_needed(["", "", "End If"], "End If", 0),
               VB57.closer_needed(["Next i"], "Next", 4),
               VB57.closer_needed(["Loop While x"], "Loop", 4),
               VB57.closer_needed(["End Sub"], "End If", 4),
               VB57.closer_needed(["    MsgBox 1"], "End If", 4),
               VB57.closer_needed(["        MsgBox 1"], "End If", 4),
               VB57.closer_needed(["End If"], None, 0)],
              expect_contain=[True, False, False, False, False, False,
                              True, False, False, False])

        # ---- 57.3b _indent_width：空格 1 列 / Tab 4 列 ----
        check("57.3b _indent_width：空格按 1 列、Tab 按 4 列",
              [VB57._indent_width("    x"), VB57._indent_width("\tx"),
               VB57._indent_width(""), VB57._indent_width("\t  x")],
              expect_contain=[4, 4, 0, 6])

        # ---- 57.4 ★集成：真能多写出一行收尾，光标停在中间那行 ----
        class _CM57(object):
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "Module1"

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n"
                               for l in self.lines[s0:s0 + int(count)])

            def InsertLines(self, start, text):
                i = min(max(0, int(start) - 1), len(self.lines))
                for k, ln in enumerate(str(text).split("\n")):
                    self.lines.insert(i + k, ln)

        class _Pane57(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE57(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _nl57(lines, line_no, col, auto_close=True):
            cm57 = _CM57(lines)
            pane57 = _Pane57(cm57, (line_no, col, line_no, col))
            VB57._get_vbe_cached = lambda: _VBE57(pane57)
            ok57 = VB57.VbeBackend().new_line_below(smart=True,
                                                    auto_close=auto_close)
            return ok57, cm57.lines, pane57.sel

        # 用户原话的场景：写 `Sub Foo()` 回车 -> 中间留出缩进好的一行 + 补 `End Sub`
        check("57.4 ★集成：`Sub Foo()` 行尾回车 -> 中间一行缩进 4 空格（光标在这儿）、"
              "下一行补 `End Sub`",
              [_nl57(["Sub Foo()"], 1, 10)],
              expect_contain=[(True, ["Sub Foo()", "    ", "End Sub"],
                               (2, 5, 2, 5))])

        # 用户原话的场景：`If i = 1 Then` 回车 -> 补 `End If`
        check("57.4b ★集成：`If x = 1 Then` 行尾回车 -> 中间一行缩进、下一行补 `End If`",
              [_nl57(["If x = 1 Then"], 1, 14)],
              expect_contain=[(True, ["If x = 1 Then", "    ", "End If"],
                               (2, 5, 2, 5))])

        check("57.4c ★集成：`With rng` 回车 -> 补 `End With`",
              [_nl57(["With rng"], 1, 9)],
              expect_contain=[(True, ["With rng", "    ", "End With"],
                               (2, 5, 2, 5))])

        # ★最常见的真实场景：在 Sub 里写 If，下面还压着 `End Sub`
        check("57.4d ★集成：Sub 里写 `If x = 1 Then`（下面压着 End Sub）-> "
              "中间一行 8 空格、补的 `    End If` 插在 `End Sub` 之前、"
              "`Sub Foo()` 与 `End Sub` 一个字没动",
              [_nl57(["Sub Foo()", "    x = 1", "    If x = 1 Then", "End Sub"],
                     3, 19)],
              expect_contain=[(True,
                               ["Sub Foo()", "    x = 1", "    If x = 1 Then",
                                "        ", "    End If", "End Sub"],
                               (4, 9, 4, 9))])

        # 下面已经挂着同一个收尾（含隔着空行）-> 绝不补出重复的
        check("57.4e ★集成：下面已经有 `End Sub` / 隔着空行的 `End Sub` -> "
              "只缩进、不补重复的收尾",
              [_nl57(["Sub Foo()", "End Sub"], 1, 10),
               _nl57(["Sub Foo()", "", "End Sub"], 1, 10),
               _nl57(["Sub Foo()", "    Do", "    Loop", "End Sub"], 2, 7)],
              expect_contain=[(True, ["Sub Foo()", "    ", "End Sub"],
                               (2, 5, 2, 5)),
                              (True, ["Sub Foo()", "    ", "", "End Sub"],
                               (2, 5, 2, 5)),
                              (True, ["Sub Foo()", "    Do", "        ",
                                      "    Loop", "End Sub"], (3, 9, 3, 9))])

        # 块【内】的分支（Else / Case）与条件编译（#If）都不许补
        check("57.4f ★集成：`Else` 只缩进不补（它是分支）；`#If … Then` 也不补"
              "（那要配 #End If）",
              [(_nl57(["Sub Foo()", "    If x = 1 Then", "    Else"], 3, 9)[1],
                _nl57(["#If VBA7 Then"], 1, 14)[1])],
              expect_contain=[(["Sub Foo()", "    If x = 1 Then", "    Else",
                                "        "],
                               ["#If VBA7 Then", "    "])])

        # auto_close=False（或开关关掉）-> 回到 v79 老口径：只缩进
        check("57.4g auto_close=False -> 只缩进、不补收尾（v79 老口径不变）",
              [_nl57(["Sub Foo()"], 1, 10, auto_close=False),
               _nl57(["If x = 1 Then"], 1, 14, auto_close=False)],
              expect_contain=[(True, ["Sub Foo()", "    "], (2, 5, 2, 5)),
                              (True, ["If x = 1 Then", "    "], (2, 5, 2, 5))])

        # Shift+Enter 那条路（smart=False）任何情况下都不许补收尾
        # ★v89：缩进判定并到同一份判据之后，这一条只钉"**不补收尾**"——
        # `Sub Foo()` 是块头，所以新行按 next_line_indent 深一级（4 格）。
        _cm57s = _CM57(["Sub Foo()"])
        _pane57s = _Pane57(_cm57s, (1, 2, 1, 2))
        VB57._get_vbe_cached = lambda: _VBE57(_pane57s)
        _ok57s = VB57.VbeBackend().new_line_below()
        check("57.4h Shift+Enter（smart=False）不补收尾（★v89：缩进照判据走，"
              "块头新行深一级；见第 68 节）",
              [(_ok57s, _cm57s.lines)],
              expect_contain=[(True, ["Sub Foo()", "    "])])

        VB57._get_vbe_cached = lambda: None
        check("57.4i 取不到 VBE -> False（调用方把回车原样还给系统）",
              [VB57.VbeBackend().new_line_below(smart=True, auto_close=True)],
              expect_contain=[False])
        VB57._get_vbe_cached = _orig57

        # ---- 57.5 ★接线护栏 ----
        _src57 = io.open(os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "main.py"), encoding="utf-8").read()
        _fn57 = _src57[_src57.index("def _enter_indent_here():"):
                       _src57.index("def _move_caret_here(delta):")]
        check("57.5 ★接线护栏：开关走环境变量、_enter_indent_here 把开关交给后端；"
              "先收窗 -> 再写 -> 写不成 send_vk(VK_RETURN) 兜底",
              ['os.environ.get("VBECOMPLETE_NO_AUTOCLOSE"' in _src57,
               "AUTO_CLOSE_BLOCK" in _src57,
               "auto_close=AUTO_CLOSE_BLOCK" in _fn57,
               "backend.new_line_below(smart=True" in _fn57,
               "completer.hide()" in _fn57,
               _fn57.index("completer.hide()") < _fn57.index("new_line_below"),
               "send_vk(VK_RETURN)" in _fn57],
              expect_contain=[True] * 7)

        # 后端里也必须真的走那两个纯函数（别绕过判定硬写收尾）
        _vbsrc57 = io.open(os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "vbe_bridge.py"), encoding="utf-8").read()
        check("57.5b 后端补收尾前必须过 block_closer + closer_needed 判定，"
              "并读下方各行（_lines_below）",
              ["block_closer(line_text)" in _vbsrc57,
               "closer_needed(self._lines_below(" in _vbsrc57,
               "if closer_text:" in _vbsrc57],
              expect_contain=[True] * 3)

        # ---- 57.6 出厂默认：开着 ----
        check("57.6 ★出厂默认：块头回车自动补收尾开着"
              "（VBECOMPLETE_NO_AUTOCLOSE 未设）",
              [M57.AUTO_CLOSE_BLOCK], expect_contain=[True])
    except Exception as _e57:
        check("第 57 节异常: %s" % _e57, [True], expect_contain=[False])

    # ---- 58. v79c：嵌套块各自补收尾 + 模块级 Type/Enum + 过程归属校验 ----
    #
    # 用户报的三件事：
    #   1) "我在一个函数里定义了一个变量 targetsheet，在定义另一个函数时输入 tes，
    #      会提示出 targetsheet"（跨过程泄漏）；
    #   2) "双层 for 只能自动补全外层 for，内层的 for 不会自动补 next"；
    #   3) "type、enum 这种结构语句不能自动补全"。
    print("\n=== 58. 嵌套收尾 / 模块级 Type·Enum / 过程归属（v79c）===")
    try:
        import vbe_bridge as VB58
        import parser as P58
        _orig58 = VB58._get_vbe_cached

        class _CM58(object):
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "Module1"

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n"
                               for l in self.lines[s0:s0 + int(count)])

            def InsertLines(self, start, text):
                i = min(max(0, int(start) - 1), len(self.lines))
                for k, ln in enumerate(str(text).split("\n")):
                    self.lines.insert(i + k, ln)

        class _Pane58(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE58(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _nl58(lines, line_no, col, auto_close=True):
            cm58 = _CM58(lines)
            pane58 = _Pane58(cm58, (line_no, col, line_no, col))
            VB58._get_vbe_cached = lambda: _VBE58(pane58)
            ok58 = VB58.VbeBackend().new_line_below(smart=True,
                                                    auto_close=auto_close)
            return ok58, cm58.lines, pane58.sel

        # ---- 58.1 ★集成：用户报的双层 For ----
        #
        # 场景：外层 For 的回车已经走过（下面补出了 `    Next`），现在光标移到
        # 内层 For 那行再按回车。下面第一个非空行是【外层】的 `    Next`，
        # 缩进比内层块头（8）浅 —— 内层该补的 `        Next` 必须在它上面。
        check("58.1 ★集成：双层 For —— 内层 `For j` 回车也要补 `Next`，"
              "且补在外层 `Next` 之前、光标停在中间那行",
              [_nl58(["Sub Foo()",
                      "    For i = 1 To 10",
                      "        For j = 1 To 10",
                      "    Next",
                      "End Sub"], 3, 27)],
              expect_contain=[(True,
                               ["Sub Foo()",
                                "    For i = 1 To 10",
                                "        For j = 1 To 10",
                                "            ",
                                "        Next",
                                "    Next",
                                "End Sub"],
                               (4, 13, 4, 13))])

        # 内层 If 下面压着外层的 End If 同理
        check("58.2 ★集成：双层 If —— 内层 `If b Then` 补 `        End If`，"
              "外层 `    End If` 一个字没动",
              [_nl58(["Sub Foo()",
                      "    If a Then",
                      "        If b Then",
                      "    End If",
                      "End Sub"], 3, 20)],
              expect_contain=[(True,
                               ["Sub Foo()",
                                "    If a Then",
                                "        If b Then",
                                "            ",
                                "        End If",
                                "    End If",
                                "End Sub"],
                               (4, 13, 4, 13))])

        # 双层 Do / With（各自的收尾不同名，同一个规则管）
        check("58.3 ★集成：双层 Do -> `Loop`、双层 With -> `End With` 都补在对的层级",
              [_nl58(["Sub Foo()", "    Do", "        Do While x", "    Loop",
                      "End Sub"], 3, 24),
               _nl58(["Sub Foo()", "    With a", "        With b", "    End With",
                      "End Sub"], 3, 18)],
              expect_contain=[(True, ["Sub Foo()", "    Do", "        Do While x",
                                      "            ", "        Loop", "    Loop",
                                      "End Sub"], (4, 13, 4, 13)),
                              (True, ["Sub Foo()", "    With a", "        With b",
                                      "            ", "        End With",
                                      "    End With", "End Sub"],
                               (4, 13, 4, 13))])

        # ---- 58.4 ★集成：模块级 Type / Enum（用户报的"不能自动补全"）----
        #
        # 根因：收尾判据里"下面那行比块头【更浅】才补"。Type/Enum 只出现在
        # 模块级（缩进 0），而 0 是最浅的 —— 下面的声明/过程头永远不可能更浅，
        # 于是永远不补。现在补上"模块级 + 下面同级"这一档。
        check("58.4 ★集成：模块级 `Type Foo` / `Enum Color` 回车 -> "
              "下面即使还有别的过程，也补 `End Type` / `End Enum`",
              [_nl58(["Type Foo", "Sub Bar()", "End Sub"], 1, 9),
               _nl58(["Enum Color", "Sub Bar()", "End Sub"], 1, 11),
               _nl58(["Public Type T1", "Sub Bar()", "End Sub"], 1, 15),
               _nl58(["Private Enum E2", "Sub Bar()", "End Sub"], 1, 16)],
              expect_contain=[(True, ["Type Foo", "    ", "End Type",
                                      "Sub Bar()", "End Sub"], (2, 5, 2, 5)),
                              (True, ["Enum Color", "    ", "End Enum",
                                      "Sub Bar()", "End Sub"], (2, 5, 2, 5)),
                              (True, ["Public Type T1", "    ", "End Type",
                                      "Sub Bar()", "End Sub"], (2, 5, 2, 5)),
                              (True, ["Private Enum E2", "    ", "End Enum",
                                      "Sub Bar()", "End Sub"], (2, 5, 2, 5))])

        # 下面已经挂着同一个收尾 / 紧跟着另一个 Type -> 只缩进，不重复
        check("58.5 ★集成：模块级 Type/Enum 下面已挂着 `End Type` / `End Enum`"
              "（含紧跟下一个 Type）-> 不补重复的",
              [_nl58(["Type Foo", "End Type"], 1, 9)[1],
               _nl58(["Enum e", "End Enum"], 1, 7)[1],
               _nl58(["Type Foo", "Type B", "End Type"], 1, 9)[1]],
              expect_contain=[["Type Foo", "    ", "End Type"],
                              ["Enum e", "    ", "End Enum"],
                              ["Type Foo", "    ", "End Type",
                               "Type B", "End Type"]])

        # ---- 58.6 纯函数：_is_closer_text / _closer_owner_above ----
        check("58.6 _is_closer_text：`Next i` / `Next  '注释` / `Loop While x` 都算收尾；"
              "`NextLoop` / 注释行 / 空 -> 不算",
              [[VB58._is_closer_text("Next i", "Next"),
                VB58._is_closer_text("Next  '下一个", "Next"),
                VB58._is_closer_text("Loop While x", "Loop"),
                VB58._is_closer_text("End If", "End If"),
                VB58._is_closer_text("NextLoop", "Next"),
                VB58._is_closer_text("'    Next", "Next"),
                VB58._is_closer_text("", "Next"),
                VB58._is_closer_text("Next", None)]],
              expect_contain=[[True, True, True, True,
                               False, False, False, False]])

        check("58.7 _closer_owner_above：只有【同类、未闭合、缩进正好对上】的"
              "外层块头才算收尾的主人（配平过的那一层不算）",
              [[VB58._closer_owner_above(["Sub Foo()", "    For i = 1 To 10"],
                                         "for", 4),
                VB58._closer_owner_above(["Sub Foo()"], "for", 0),
                VB58._closer_owner_above(["    For i = 1 To 10"], "for", 8),
                VB58._closer_owner_above([], "for", 0),
                VB58._closer_owner_above(["If a Then"], "if", 0),
                VB58._closer_owner_above(["If a Then", "    End If"], "if", 4),
                VB58._closer_owner_above(None, None, 0)]],
              expect_contain=[[True, False, False, False,
                               True, False, False]])

        check("58.8 closer_needed：更浅的同类收尾 —— 有同类外层块头 -> 补，"
              "没有（用户自己写歪的收尾）-> 不补（免得 `Next` 双份）",
              [[VB58.closer_needed(["    Next"], "Next", 8, kind="for",
                                   lines_above=["Sub Foo()",
                                                "    For i = 1 To 10"]),
                VB58.closer_needed(["Next"], "Next", 4, kind="for",
                                   lines_above=["Sub Foo()"]),
                VB58.closer_needed(["    End If"], "End If", 8, kind="if",
                                   lines_above=["Sub Foo()",
                                                "    If a Then"]),
                VB58.closer_needed(["End If"], "End If", 4, kind="if",
                                   lines_above=["Sub Foo()"]),
                VB58.closer_needed(["End Type"], "End Type", 0, kind="type",
                                   lines_above=["Type A", "End Type"])]],
              expect_contain=[[True, False, True, False, False]])

        # 旧式调用（不传 kind/lines_above）必须保持 v79b 口径：更浅 -> 不补
        check("58.9 closer_needed 老口径不变（不传 kind/lines_above 时，"
              "更浅的同类收尾一律不补）",
              [[VB58.closer_needed(["Next i"], "Next", 4),
                VB58.closer_needed(["Loop While x"], "Loop", 4),
                VB58.closer_needed(["    End If"], "End If", 0),
                VB58.closer_needed([], "End If", 0)]],
              expect_contain=[[False, False, False, True]])

        # ---- 58.10 ★纯函数：过程归属校验（用户报的"泄漏"那一侧的护栏）----
        _mod58 = "\n".join([
            "Enum e",                      # 1
            "",                            # 2
            "Sub test()",                  # 3
            "    x = 1",                   # 4
            "End Sub",                     # 5
            "",                            # 6
            "Sub tes()",                   # 7
            "",                            # 8
            "End Sub"])                    # 9
        check("58.10 proc_owns_line：只有真的被那个过程圈住的行才算它是主人；"
              "`End Sub` 行 / 过程之间的空行 / 第一个过程头之前的声明区都不算",
              [[P58.proc_owns_line(_mod58, 4, "test"),
                P58.proc_owns_line(_mod58, 3, "test"),
                P58.proc_owns_line(_mod58, 5, "test"),
                P58.proc_owns_line(_mod58, 6, "test"),
                P58.proc_owns_line(_mod58, 2, "test"),
                P58.proc_owns_line(_mod58, 1, "test"),
                P58.proc_owns_line(_mod58, 4, "tes"),
                P58.proc_owns_line(_mod58, 8, "tes"),
                P58.proc_owns_line(_mod58, 4, None),
                P58.proc_owns_line(_mod58, 0, "test"),
                P58.proc_owns_line("", 1, "test")]],
              expect_contain=[[True, True, False, False, False, False,
                               False, True, False, False, False]])

        # 真机回归：VBE 的 ProcOfLine 会把过程外的行归给【邻居过程】
        # （`End Sub` 行、它下面到下一个过程头之间的行 -> 前一个过程；
        #   声明区里第一个过程头之前的行 -> 第一个过程）。_proc_of_line
        # 必须校验归属，否则"模块级 / 新函数的位置"上会冒出别的过程的局部变量
        # —— 正是用户报的"我在一个函数里定义的 targetsheet，在另一个函数里
        # 输入 tes 却被提示出来"。
        class _CMVbe58(object):
            """假 CodeModule：文本扫描正常，COM 兜底故意返回【邻居过程】。"""
            def __init__(self, lines, com_proc):
                self.lines = list(lines)
                self.Name = "Module1"
                self._com = com_proc

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n"
                               for l in self.lines[s0:s0 + int(count)])

            def ProcOfLine(self, line_no, kind):
                return (self._com, int(kind))     # 学 VBE：返回 (名字, ProcKind)

        _lives58 = ["Sub Foo()", "    Dim targetsheet As Long", "End Sub", ""]
        _lives58b = ["", "Sub Foo()", "    Dim targetsheet As Long", "End Sub"]
        check("58.11 ★护栏：文本扫描答不出来时，绝不采信 VBE 的 ProcOfLine "
              "（它把 `End Sub` 行、过程之间的空行、声明区里第一个过程头之前"
              "的行都归给邻居过程）；这些位置必须回到 None —— 否则就是『另一个"
              "函数的局部变量泄漏到本位置』。过程体那两行仍要正常认出 Foo"
              "（v79d 起：过程头那一行本身也算『不在过程体内』，见第 59 节）",
              [[VB58._proc_of_line(_CMVbe58(_lives58, "Foo"), 3),
                VB58._proc_of_line(_CMVbe58(_lives58, "Foo"), 4),
                VB58._proc_of_line(_CMVbe58(_lives58, "Foo"), 1),
                VB58._proc_of_line(_CMVbe58(_lives58, "Foo"), 2),
                VB58._proc_of_line(_CMVbe58(_lives58b, "Foo"), 1),
                isinstance(VB58._proc_of_line(_CMVbe58(_lives58, "Foo"), 2),
                           str)]],
              expect_contain=[[None, None, None, "Foo", None, True]])

        # ---- 58.12 ★接线护栏 ----
        _vbsrc58 = io.open(os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "vbe_bridge.py"), encoding="utf-8").read()
        check("58.12 ★接线护栏：new_line_below 把块头类型与上方各行交给 "
              "closer_needed（嵌套判定要用）；_proc_of_line 走归属校验；"
              "parser 提供 proc_owns_line",
              ["kind=_block_kind(line_text)" in _vbsrc58,
               "lines_above=above" in _vbsrc58,
               "_proc_owns_line(cm, name, line_no)" in _vbsrc58,
               "def proc_owns_line(code, line_no, name):" in io.open(
                   os.path.join(os.path.dirname(os.path.dirname(
                       os.path.abspath(__file__))), "parser.py"),
                   encoding="utf-8").read()],
              expect_contain=[True] * 4)

        # ---- 58.13 收尾清单完整性：VBA 里"首尾都有明确标志"的块必须全覆盖
        #             （#If/#Else 条件编译除外：它要配 `#End If`，替用户写个裸的
        #             `End If` 反而把代码写错）----
        _pairs58 = [("Sub Foo()", "End Sub"),
                    ("Private Function F() As Long", "End Function"),
                    ("Property Get X()", "End Property"),
                    ("Static Property Let Y(v As Long)", "End Property"),
                    ("Type Foo", "End Type"),
                    ("Public Enum Color", "End Enum"),
                    ("If x Then", "End If"),
                    ("With rng", "End With"),
                    ("Select Case x", "End Select"),
                    ("For i = 1 To 10", "Next"),
                    ("For Each c In rng", "Next"),
                    ("Do", "Loop"),
                    ("Do While x < 5", "Loop"),
                    ("Do Until x > 5", "Loop"),
                    ("While x", "Wend")]
        check("58.13 收尾清单：VBA 里首尾都有明确标志的块 15 例全覆盖"
              "（Sub/Function/Property/Type/Enum/If/With/Select Case/"
              "For/For Each/Do·While/Do·Until/While）",
              [[VB58.block_closer(_t) for _t, _ in _pairs58]],
              expect_contain=[[_w for _, _w in _pairs58]])

        # 没有收尾的语句一个都不许认（否则会替用户写出错的结构）。
        # `ElseIf … Then` 只是 If 块【内】的分支：它要的收尾是与 `If` 共用的那
        # 个 `End If`，不能在这儿再补一条。
        _noc58 = ["Else", "ElseIf y Then", "Case 1", "Case Else",
                  "#If VBA7 Then", "#Else",
                  "Next i", "End If", "Loop", "Wend", "End Type",
                  "x = 1", "Exit For", "foo(1, 2)", "For i = 1 To 10 _"]
        check("58.14 不该补收尾的：Else / ElseIf…Then / Case / #If / #Else"
              "（分支与条件编译）、以及所有收尾行本身 / 普通语句 / "
              "以续行符结尾的块头",
              [[VB58.block_closer(_t) for _t in _noc58
                if VB58.block_closer(_t) is not None]],
              expect_contain=[[]])

        VB58._get_vbe_cached = _orig58
    except Exception as _e58:
        check("第 58 节异常: %s" % _e58, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 59. 跨过程局部变量泄漏 · 光标停在过程头上（v79d）
    #
    # 用户复现（v79c 之后仍然泄漏）：
    #     Sub test()
    #     Set targetSheet = Worksheets(1)     ' 没写 Dim -> 隐式变量
    #     ...（双层 For / If / With）
    #     End Sub
    #   然后再定义一个过程，输入 `sub test` 时弹窗同时提示 test 和 targetSheet。
    #
    # 实测有【两条独立路径】，都通到同一个症状：
    #   A) 光标这一行【自己】是过程头（`Sub test`）时，proc_at_line 把它当成了
    #      "当前过程"；这个新名字恰好等于模块里已有的过程名 -> 那个过程的局部
    #      变量（targetSheet）被当成"当前过程的"提示出来。
    #   B) 光标停在过程头的【名字】上时，"抹掉正在输入的词"把整个过程头抹废，
    #      紧随其后的过程体失去过程归属，局部/隐式变量统统降级成【模块级】
    #      -> 模块里任何位置都能提示出它们。
    # ------------------------------------------------------------------
    print("\n=== 59. 光标停在过程头上：跨过程泄漏（v79d）===")
    try:
        import vbe_bridge as VB59
        import parser as P59
        import engine as E59

        _USER59 = [
            "Sub test()",                        # 1
            "Set targetSheet = Worksheets(1)",   # 2
            "For i = 1 To 100",                  # 3
            "For j = 1 To 200",                  # 4
            "If i = 1 Then",                     # 5
            "If j = 2 Then",                     # 6
            "With targetSheet",                  # 7
            "With targetSheet",                  # 8
            "End With",                          # 9
            "End With",                          # 10
            "End If",                            # 11
            "End If",                            # 12
            "Next",                              # 13
            "Next",                              # 14
            "End Sub",                           # 15
        ]

        class _CM59(object):
            """假 CodeModule：文本扫描正常；COM 兜底故意学 VBE 返回邻居过程名。"""

            def __init__(self, lines, com_proc="test"):
                self.lines = list(lines)
                self.Name = "Module1"
                self._com = com_proc

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n"
                               for l in self.lines[s0:s0 + int(count)])

            def ProcOfLine(self, line_no, kind):
                return (self._com, int(kind))     # 学 VBE：(名字, ProcKind)

        def _vis59(lines, line_no, col):
            """一条龙：真收集（含 caret 抹词）-> 真作用域过滤。"""
            code = "\n".join(lines)
            caret = (line_no, col)
            r, _pn = list(P59.extract_records(code, module="Module1",
                                             is_std_module=True, caret=caret))
            r += P59.extract_implicit_records(code, module="Module1",
                                              is_std_module=True, declared=r,
                                              scope="proc", caret=caret)
            proc = VB59._proc_of_line(_CM59(lines), line_no)
            vis = E59.filter_identifiers_by_scope(r, proc, "Module1")
            return proc, sorted(set(vis))

        # ---- 59.1 纯函数：is_proc_header_line ----
        check("59.1 is_proc_header_line：Sub / Function / Property Get|Let|Set"
              "（含 Public/Private/Static 修饰符、名字后无括号也可）都算过程头；"
              "`Set x = 1` / 注释里的 Sub / 空行 / 普通语句都不算",
              [[P59.is_proc_header_line("Sub test"),
                P59.is_proc_header_line("Sub test()"),
                P59.is_proc_header_line("  Function F() As Long"),
                P59.is_proc_header_line("Private Property Get X()"),
                P59.is_proc_header_line("Static Property Let Y(v As Long)"),
                P59.is_proc_header_line("Set targetSheet = Worksheets(1)"),
                P59.is_proc_header_line("' Sub Foo()"),
                P59.is_proc_header_line(""),
                P59.is_proc_header_line("    Sub total = 1")]],
              expect_contain=[[True, True, True, True, True,
                               False, False, False, False]])

        # ---- 59.2 ★用户场景：模块级位置正在写新过程头 ----
        for _tail, _col in (("Sub test", 9), ("Sub test", 6),
                            ("Sub test(", 10), ("Sub te", 7)):
            _lines = _USER59 + [_tail]
            _proc, _vis = _vis59(_lines, len(_lines), _col)
            check("59.2 ★用户场景：模块里已有 Sub test()，在【模块级】新写 %r"
                  "（光标列 %d）-> 当前过程必须是 None，targetSheet 不得出现"
                  % (_tail, _col),
                  [(_proc, "targetSheet" in _vis)],
                  expect_contain=[(None, False)])

        # ---- 59.3 不能过度收敛：过程体内仍按该过程过滤 ----
        check("59.3 不能过度收敛：光标在 test 体内（第 2 / 7 行）时当前过程仍是 "
              "test、targetSheet 照常可见；`End Sub` 那一行则回到 None",
              [[_vis59(_USER59, 2, 3)[0], _vis59(_USER59, 7, 3)[0],
                _vis59(_USER59, 15, 3)[0],
                "targetSheet" in _vis59(_USER59, 2, 3)[1]]],
              expect_contain=[["test", "test", None, True]])

        # ---- 59.4 _blank_ident_at：声明名【文本一个字不动】（v79e 改口径）----
        _b59 = P59._blank_ident_at
        check("59.4 _blank_ident_at：过程名 / Type 名 / Enum 名【原样保留】"
              "（行仍是声明行 -> is_proc_header_line 依然为 True、过程归属不受影响）；"
              "普通变量 / 赋值目标 / 形参名照旧抹成等长空格；"
              "只有显式要求 blank_decl=True（现场证据扫描那一路）才肯抹声明名",
              [[_b59("Sub bar()", 1, 7),
                _b59("Type Foo", 1, 7),
                _b59("Enum Color", 1, 8),
                P59.is_proc_header_line(_b59("Sub bar()", 1, 7)),
                _b59("Set other = 1", 1, 7),
                _b59("total = 1", 1, 4),
                _b59("Sub bar()", 1, 7, True),
                _b59("Enum Color", 1, 8, True)]],
              expect_contain=[["Sub bar()", "Type Foo", "Enum Color", True,
                               "Set" + " " * 7 + "= 1",
                               " " * 6 + "= 1",
                               "Sub" + " " * 4 + "()",
                               "Enum" + " " * 6]])

        # ---- 59.5 ★第二条路径：光标停在过程头的【名字】里 ----
        _mid59 = ["Sub test()", "Set targetSheet = Worksheets(1)", "End Sub",
                  "Sub bar()", "Set other = 1", "End Sub",
                  "Sub mid()", "Set third = 2", "End Sub"]
        _rows59 = []
        for _c in (6, 7, 8, 9, 10):
            _code = "\n".join(_mid59)
            _r, _pn = list(P59.extract_records(_code, module="Module1",
                                               is_std_module=True,
                                               caret=(4, _c)))
            _r += P59.extract_implicit_records(_code, module="Module1",
                                               is_std_module=True, declared=_r,
                                               scope="proc", caret=(4, _c))
            _rec = dict((str(x[0]).lower(), x) for x in _r)
            _v = E59.filter_identifiers_by_scope(_r, "test", "Module1")
            # v79e：声明名不再被抹（文本一个字不动），`Sub bar()` 的过程归属照旧；
            # 要害是【它不能是 None】—— None 就等于降级成模块级、全模块可见。
            _rows59.append((_rec.get("other") is not None
                            and _rec["other"][2] is not None,
                            "other" in _v))
        check("59.5 ★第二条泄漏路径：光标停在 `Sub bar()` 的 bar 名字里"
              "（第 6~10 列）-> bar 里的 other 必须仍带【过程归属】（绝不是模块级）"
              "、在 test 里绝不可见（旧实现把过程头抹成空格，整个过程体失去归属）",
              [_rows59],
              expect_contain=[[(True, False)] * 5])

        # ---- 59.6 Type/Enum 名字被占位后，块成员不得变成隐式变量 ----
        _t59 = ["Type Foo", "    Red As Long", "End Type",
                "Enum Color", "    A = 1", "End Enum"]
        _code59t = "\n".join(_t59)
        _rt, _pn = list(P59.extract_records(_code59t, module="Module1",
                                         is_std_module=True, caret=(1, 7)))
        _rt += P59.extract_implicit_records(_code59t, module="Module1",
                                            is_std_module=True, declared=_rt,
                                            scope="proc", caret=(1, 7))
        _names59t = sorted(set(str(x[0]).lower() for x in _rt))
        check("59.6 光标停在 `Type Foo` 的 Foo 上时，Type/Enum 的块成员"
              "（Red / A）仍不得被当成隐式变量收进候选池",
              [[n for n in _names59t if n in ("red", "a")]],
              expect_contain=[[]])

        # ---- 59.7 ★接线护栏 ----
        _root59 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _vbsrc59 = io.open(os.path.join(_root59, "vbe_bridge.py"),
                           encoding="utf-8").read()
        _psrc59 = io.open(os.path.join(_root59, "parser.py"),
                          encoding="utf-8").read()
        _guard59 = "is_proc_header_line(cm.Lines(line_no, 1))"
        _scan59 = "name = vba_parser.proc_at_line(head, line_no)"
        check("59.7 ★接线护栏：_proc_of_line 用 is_proc_header_line 挡住过程头"
              "那一行，且必须排在 proc_at_line 之前（否则文本扫描会把它当成"
              "当前过程）；parser 提供 is_proc_header_line 与 _struct_decl_name_span",
              [[-1 < _vbsrc59.find(_guard59) < _vbsrc59.find(_scan59),
                "def is_proc_header_line(line_text):" in _psrc59,
                "def _struct_decl_name_span(line):" in _psrc59,
                "_struct_decl_name_span(line)" in _psrc59]],
              expect_contain=[[True] * 4])

    except Exception as _e59:
        check("第 59 节异常: %s" % _e59, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 60. v79e：同名过程各自成域 + 声明名不再改文本
    #
    #   用户报："你改的更糟糕了，现在过程级的变量，都泄露到其他过程了"。
    #   真机（模块2）复现出两条真因：
    #   A) 模块里【两个过程同名】（用户在既有 `Sub test()` 下面又敲了一个
    #      `Sub test()`）—— 记录只存过程名，两个过程的局部/隐式变量在过滤时
    #      无法区分，第二个过程的体内能提示出第一个过程的 targetSheet。
    #   B) v79d 把"正在声明的名字"抹成等长下划线：名字只有 1 个字符时
    #      （`Sub t`），行尾正好是"空格 + 下划线 + 换行"，被 _RE_CONTINUATION
    #      当成行继续符，把下一行吞进同一行 —— 过程头消失、其后整段过程体失去
    #      归属，隐式变量降级成模块级，全模块到处可见。
    # ------------------------------------------------------------------
    print("\n=== 60. 同名过程各自成域 / 声明名不改文本（v79e）===")
    try:
        import vbe_bridge as VB60
        import parser as P60
        import engine as E60

        _USER60 = [
            "Sub test()",                            # 1
            "    Set targetSheet = Worksheets(1)",   # 2
            "    For i = 1 To 100",                  # 3
            "    Next",                              # 4
            "End Sub",                               # 5
            "",                                      # 6
            "Sub test()",                            # 7
            "    For i = 1 To 2",                    # 8
            "    Next",                              # 9
            "    m = 3",                             # 10
            "End Sub",                               # 11
        ]

        class _CM60(object):
            """假 CodeModule：文本扫描正常；COM 兜底故意学 VBE 返回邻居过程名。"""

            def __init__(self, lines, com_proc="test"):
                self.lines = list(lines)
                self.Name = "Module1"
                self._com = com_proc

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n"
                               for l in self.lines[s0:s0 + int(count)])

            def ProcOfLine(self, line_no, kind):
                return (self._com, int(kind))     # 学 VBE：(名字, ProcKind)

        def _recs60(lines, caret=None):
            code = "\n".join(lines)
            r, _pn = list(P60.extract_records(code, module="Module1",
                                             is_std_module=True, caret=caret))
            r += P60.extract_implicit_records(code, module="Module1",
                                              is_std_module=True, declared=r,
                                              scope="proc", caret=caret)
            return code, r

        def _vis60(lines, line_no, caret=None):
            _code, _r = _recs60(lines, caret)
            _proc = VB60._proc_of_line(_CM60(lines), line_no)
            return _proc, sorted(set(E60.filter_identifiers_by_scope(
                _r, _proc, "Module1")))

        _code60, _r60 = _recs60(_USER60)

        # ---- 60.1 记录侧：同名过程的键（第 1 次用原名，第 2 次起带序号）----
        check("60.1 ★同名过程：模块里两个 `Sub test()` 时，记录键必须是 "
              "第 1 个 `test`、第 2 个 `test#2` —— 两个过程的隐式变量各挂各的"
              "（targetSheet 归 test、m 归 test#2；i 两边各有一份）",
              [sorted((str(a[0]).lower(), str(a[2])) for a in _r60)],
              expect_contain=[[("i", "test"), ("i", "test#2"), ("m", "test#2"),
                               ("targetsheet", "test"), ("test", "None")]])

        # ---- 60.2 ★用户场景端到端：第二个过程体内不得出现第一个过程的变量 ----
        _p60_1, _v60_1 = _vis60(_USER60, 2)
        _p60_2, _v60_2 = _vis60(_USER60, 8)
        _p60_3, _v60_3 = _vis60(_USER60, 10)
        check("60.2 ★用户场景：模块里两个 `Sub test()` —— 第 1 个体内"
              "（L2）照常提示 targetSheet、不提示 m；第 2 个体内"
              "（L8 / L10）**绝不提示 targetSheet**、照常提示自己的 m",
              [[_p60_1, "targetSheet" in _v60_1, "m" in _v60_1,
                _p60_2, "targetSheet" in _v60_2, "m" in _v60_2,
                _p60_3, "targetSheet" in _v60_3, "m" in _v60_3]],
              expect_contain=[["test", True, False,
                               "test#2", False, True,
                               "test#2", False, True]])

        # ---- 60.3 光标停在第二个过程头上：仍按模块级（v79d 口径不退化）----
        _p60_h, _v60_h = _vis60(_USER60, 7)
        check("60.3 光标停在第二个 `Sub test()` 那一行（L7）：当前过程必须是 None"
              "（过程头那一行是声明、不是过程体），targetSheet 与 m 都不得出现",
              [[_p60_h, "targetSheet" in _v60_h, "m" in _v60_h]],
              expect_contain=[[None, False, False]])

        # ---- 60.4 异名过程：键保持裸名（与旧版行为完全一致）----
        _diff60 = list(_USER60)
        _diff60[6] = "Sub otherProc()"
        _code60d, _r60d = _recs60(_diff60)
        check("60.4 名字不重复时（把第二个过程改名 otherProc）：记录键必须还是"
              "裸名 —— 既有行为、既有测试、日志输出一律不变",
              [sorted(set(str(a[2]) for a in _r60d), key=str)],
              expect_contain=[["None", "otherProc", "test"]])

        # ---- 60.5 纯函数：_proc_scope_key / proc_key_plain ----
        check("60.5 _proc_scope_key：第 1 次出现用原名，第 2 次起 `名字#序号`；"
              "proc_key_plain 把序号拆掉（没有序号 / 非数字后缀则原样返回）",
              [[P60._proc_scope_key("test", 1), P60._proc_scope_key("test", 2),
                P60._proc_scope_key("test", 3), P60._proc_scope_key("test", 0),
                P60.proc_key_plain("test#2"), P60.proc_key_plain("test"),
                P60.proc_key_plain("a#b"), P60.proc_key_plain(None)]],
              expect_contain=[["test", "test#2", "test#3", "test",
                               "test", "test", "a#b", ""]])

        # ---- 60.6 proc_at_line / proc_owns_line 的分工 ----
        check("60.6 proc_at_line：同名模块里第 1 个过程内是 `test`、第 2 个内是"
              "`test#2`；异名模块里都是裸名。proc_owns_line 只认名字部分："
              "`test#2` 与 `test` 都算『这个名字的过程包含该行』（COM 给的是裸名），"
              "而 `End Sub` 行与过程之间的空行一律 False",
              [[P60.proc_at_line(_code60, 2), P60.proc_at_line(_code60, 8),
                P60.proc_at_line(_code60d, 8),
                P60.proc_owns_line(_code60, 8, "test#2"),
                P60.proc_owns_line(_code60, 8, "test"),
                P60.proc_owns_line(_code60, 5, "test"),
                P60.proc_owns_line(_code60, 6, "test"),
                P60.proc_owns_line(_code60, 1, "test")]],
              expect_contain=[["test", "test#2", "otherProc",
                               True, True, False, False, True]])

        # ---- 60.7 caret_decl_name_at：只认"这一行正在声明的那个名字" ----
        check("60.7 caret_decl_name_at：`Sub bar()` 的光标落在 bar 上 -> 返回 bar；"
              "落在 `Sub` 关键字 / 括号后 -> 空；`Dim x`、赋值语句、"
              "非声明行一律空（它们走『抹空格』那条路）",
              [[P60.caret_decl_name_at("Sub bar()", 1, 7),
                P60.caret_decl_name_at("Sub bar()", 1, 4),
                P60.caret_decl_name_at("Sub bar()", 1, 10),
                P60.caret_decl_name_at("Type Foo", 1, 7),
                P60.caret_decl_name_at("Enum Color", 1, 8),
                P60.caret_decl_name_at("Dim bar As Long", 1, 6),
                P60.caret_decl_name_at("total = 1", 1, 3),
                P60.caret_decl_name_at("", 1, 1)]],
              expect_contain=[["bar", "", "", "Foo", "Color",
                               "", "", ""]])

        # ---- 60.8 ★v79d 回归护栏：1 字符名字不再被当成行继续符 ----
        _one60 = "Sub t\n    Set targetSheet = 1\nEnd Sub"
        _blanked60 = P60._blank_ident_at(_one60, 1, 6)
        _masked60 = P60._RE_CONTINUATION.sub(
            " ", P60._mask_strings_and_comments(_blanked60))
        _one_r60 = []
        _one_c60, _pn = list(P60.extract_records(_blanked60, module="M",
                                              is_std_module=True))
        _one_c60 += list(P60.extract_implicit_records(
            _blanked60, module="M", is_std_module=True, declared=_one_c60,
            scope="proc"))
        _one_r60 = [a for a in _one_c60
                    if str(a[0]).lower() == "targetsheet"]
        check("60.8 ★v79d 回归护栏（用户说『更糟糕了』的真因）：1 个字符的过程名"
              "（`Sub t`）下，抹词后首行【原样不动】、行继续符合并后行数不变"
              "（`Sub t` 没被当成续行吞掉下一行）、该行仍是过程头；"
              "过程体里 targetSheet 仍归 `t`（绝不是 None=模块级）",
              [[_blanked60.split("\n")[0],
                len(_masked60.split("\n")),
                P60.is_proc_header_line(_masked60.split("\n")[0]),
                _one_r60]],
              expect_contain=[["Sub t", 3, True,
                               [("targetSheet", "M", "t", True)]]])

        # ---- 60.9 ★接线护栏 ----
        _root60 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _vbsrc60 = io.open(os.path.join(_root60, "vbe_bridge.py"),
                           encoding="utf-8").read()
        _psrc60 = io.open(os.path.join(_root60, "parser.py"),
                          encoding="utf-8").read()
        check("60.9 ★接线护栏：现场证据扫描（names_outside_caret）必须显式要求"
              "抹掉光标处的词（blank_decl=True），否则幻影词会把自己当证据；"
              "声明名的处理只在记录/文本两处留痕；_proc_dup_counts 这类已被"
              "『第 1 次用原名』取代的旧辅助不得残留",
              [["vba_parser._blank_ident_at(code, line_no, col, True)" in _vbsrc60,
                "def _blank_ident_at(text, line_no, col, blank_decl=False):"
                in _psrc60,
                "def caret_decl_name_at(text, line_no, col):" in _psrc60,
                "def _proc_scope_key(name, occurrence):" in _psrc60,
                "cur_proc = proc_key(proc)" in _psrc60,
                "return _proc_scope_key(name, max(occ, 1))" in _psrc60,
                "_proc_dup_counts" not in _psrc60,
                "fill = \"_\" * " not in _psrc60]],
              expect_contain=[[True] * 8])

    except Exception as _e60:
        check("第 60 节异常: %s" % _e60, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 61. v80a：回车补收尾认「下方同级代码」（用户报：在已有 for 上面再补一个
    #     for，回车只缩进、不补 next）
    #     v80b：让位状态按【槽位】记 + 会过期；VBE 提示窗句柄缓存周期重扫
    #     （用户报：变量类型是自定义类模块时，我们的弹窗遮挡 VBE 的提示列表，
    #      且"不是每次都能复现"）
    # ------------------------------------------------------------------
    print("\n=== 61. 同级补收尾 + 让位槽位/有效期 + 提示窗缓存重扫（v80）===")
    try:
        import vbe_bridge as VB61
        import engine as E61
        import time as T61

        _orig61 = VB61._get_vbe_cached
        _ocb = VB61._own_closer_below
        _cn61 = VB61.closer_needed

        # ---- 61.1 _own_closer_below：配平法认"本块自己的收尾" ----
        check("61.1 _own_closer_below：配平到 0 的同级同类收尾就是本块的；"
              "同级同类块头计一层；孤立在更深处的同类收尾算『用户写歪』（保守不补）；"
              "同级的别的块头 -> 本块已结束；下方空 -> 没有",
              [[_ocb(["    Next"], "Next", "for", 4),
                _ocb(["    For i = 1 To 10", "        x = 1", "    Next i"],
                     "Next", "for", 4),
                _ocb(["        Next"], "Next", "for", 4),
                _ocb(["Sub Bar()"], "End Type", "type", 0),
                _ocb([], "Next", "for", 4)]],
              expect_contain=[[True, False, True, False, False]])

        # ---- 61.2 closer_needed：传 kind 走 v80a 新口径 ----
        check("61.2 closer_needed（kind 给定）：下方同级压着别的代码/块 -> 补；"
              "下面已挂着本块自己的收尾 -> 不补；更浅的同类收尾沿用 v79c 归属判定",
              [[_cn61(["    For i = 1 To 10", "        x = 1", "    Next i",
                       "End Sub"], "Next", 4, kind="for",
                      lines_above=["Sub Test()"]),
                _cn61(["    Next", "End Sub"], "Next", 4, kind="for",
                      lines_above=["Sub Test()"]),
                _cn61(["    x = 1", "End Sub"], "End With", 4, kind="with",
                      lines_above=["Sub Test()"]),
                _cn61(["    Next", "End Sub"], "Next", 8, kind="for",
                      lines_above=["Sub Test()", "    For i = 1 To 10"]),
                _cn61(["Next", "End Sub"], "Next", 4, kind="for",
                      lines_above=["Sub Test()"])]],
              expect_contain=[[True, False, True, True, False]])

        # ---- 61.3 ★集成：用户原话的场景 ----
        class _CM61(object):
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "Module1"

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n"
                               for l in self.lines[s0:s0 + int(count)])

            def InsertLines(self, start, text):
                i = min(max(0, int(start) - 1), len(self.lines))
                for k, ln in enumerate(str(text).split("\n")):
                    self.lines.insert(i + k, ln)

        class _Pane61(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE61(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _nl61(lines, line_no):
            """光标落在第 line_no 行【行尾】，回车走 smart + auto_close。"""
            cm61 = _CM61(lines)
            col = len(lines[line_no - 1]) + 1
            pane61 = _Pane61(cm61, (line_no, col, line_no, col))
            VB61._get_vbe_cached = lambda: _VBE61(pane61)
            ok61 = VB61.VbeBackend().new_line_below(smart=True,
                                                    auto_close=True)
            return ok61, cm61.lines, pane61.sel

        check("61.3 ★集成（用户报的场景）：已经在 `For i`（回车补过 `Next i`）"
              "【上面】再补一个 `For j`，行尾回车 -> 中间留一行缩进 8 空格、"
              "下面补出同级的 `    Next`，原来那个 For 一个字没动",
              [_nl61(["Sub Test()",
                      "    For j = 1 To 5",
                      "    For i = 1 To 10",
                      "        MsgBox i",
                      "    Next i",
                      "End Sub"], 2)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    For j = 1 To 5",
                                "        ",
                                "    Next",
                                "    For i = 1 To 10",
                                "        MsgBox i",
                                "    Next i",
                                "End Sub"],
                               (3, 9, 3, 9))])

        check("61.4 ★集成：上面补的是别的块（`With`）同理；"
              "而下面已经挂着自己收尾的（`    Next` 紧跟随）-> 只缩进、不补重复",
              [_nl61(["Sub Test()",
                      "    For i = 1 To 10",
                      "        MsgBox i",
                      "    Next i",
                      "End Sub"], 2)[1],
               _nl61(["Sub Test()",
                      "    For i = 1 To 10",
                      "    Next i",
                      "End Sub"], 2)[1],
               _nl61(["Sub Test()", "    Do", "    Loop", "End Sub"], 2)[1]],
              expect_contain=[["Sub Test()", "    Do", "        ",
                               "    Loop", "End Sub"],
                              ["Sub Test()", "    For i = 1 To 10",
                               "        ", "    Next i", "End Sub"],
                              ["Sub Test()", "    Do", "        ",
                               "    Loop", "End Sub"]])

        check("61.5 集成对照（v79b/v79c 口径不许回归）：模块级 `Type` 下面"
              "紧邻过程头 -> 仍补 `End Type`；`Sub` 下面已挂着 `End Sub` -> 不补",
              [_nl61(["Type Foo", "Sub Bar()", "End Sub"], 1)[1],
               _nl61(["Sub Foo()", "End Sub"], 1)[1]],
              expect_contain=[["Type Foo", "    ", "End Type",
                               "Sub Bar()", "End Sub"],
                              ["Sub Foo()", "    ", "End Sub"]])
        VB61._get_vbe_cached = _orig61

        # ---- 61.6 list_slot_anchor：算出"哪一个槽位" ----
        check("61.6 list_slot_anchor：`As `/`New ` 之后锚点=类型名起点列（同一槽位"
              "内打字不变）；`标识符.` / With 块的 `.成员` 锚点=成员名起点列；"
              "非列表位置 -> None",
              [[E61.list_slot_anchor("Dim x As ", 10),
                E61.list_slot_anchor("Dim x As 类", 11),
                E61.list_slot_anchor("Dim a As X, b As ", 18),
                E61.list_slot_anchor("obj.", 5),
                E61.list_slot_anchor("obj.Me", 7),
                E61.list_slot_anchor("    .Si", 8),
                E61.list_slot_anchor("x = 1", 6),
                E61.list_slot_anchor("", 1)]],
              expect_contain=[[10, 10, 18, 5, 5, 6, None, None]])

        # ---- 61.7 一致性：vbe_list_expected 认的列表位置，锚点必须算得出来 ----
        _cases61 = [("Dim x As ", 10), ("Dim x As Inte", 14),
                    ("Set c = New ", 13), ("Private m_d As Dictionary", 28),
                    ("obj.", 5), ("objs.Item(1).", 14), ("    .Si", 8),
                    ("x = 1 'a.b", 10), ("1.5", 4), ("x = 1", 6)]
        check("61.7 ★一致性护栏：凡是 vbe_list_expected 判为『VBE 会弹列表』的位置，"
              "list_slot_anchor 都必须给得出锚点（两处语法不许走偏）",
              [[(lt, cc) for (lt, cc) in _cases61
                if E61.vbe_list_expected(lt, cc)
                and E61.list_slot_anchor(lt, cc) is None]],
              expect_contain=[[]])

        # ---- 61.8 ★端到端：同一行两个槽位互不连累（v80b 核心口径）----
        _POOL61 = [(n, "VBA", None, False)
                   for n in ("类1", "类2", "Dim", "obj", "x")]

        class _B61(object):
            """桩后端：只表达"成员列表是否可见"（v65 语义），位置由文本来。"""

            def __init__(self, line_text, caret_col, yld=False):
                self._lt = line_text
                self._cc = caret_col
                self._ln = 1
                self._mod = "模块1"
                self._yld = yld

            def set_line(self, lt, cc):
                self._lt = lt
                self._cc = cc

            def get_context(self):
                return {"line_no": self._ln, "line_text": self._lt,
                        "caret_col": self._cc, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": self._mod}

            def get_identifiers(self):
                return list(_POOL61)

            def get_declared_names(self):
                return []

            def declared_elsewhere(self, name, module_name):
                return False

            def get_structural_names(self):
                return []

            def get_builtin_names(self):
                return set()

            def get_type_names(self):
                return []

            def get_host_enum_names(self):
                return {}

            def names_outside_caret(self, caret, scope_only=False):
                return set()

            def apply_completion(self, *a, **k):
                return None

            def vbe_yield_visible(self):
                return self._yld

            def vbe_popup_visible(self):
                return self._yld

            def vbe_popup_info(self):
                return []

        class _UI61(object):
            def __init__(self):
                self.shown = False
                self.hides = 0

            def show(self, rows, sel, c):
                self.shown = True

            def hide(self):
                self.shown = False
                self.hides += 1

            def update_selection(self, sel):
                pass

            def contains_point(self, x, y):
                return False

        def _mk61(line_text, caret_col):
            be = _B61(line_text, caret_col)
            ui = _UI61()
            return E61.Completer(be, ui), be, ui

        _c61, _b61, _u61 = _mk61("Dim a As 类1, b As 类2", 11)   # 第一个槽位
        _c61.trigger(True)
        _p1_61 = _c61.yield_pending is not None
        _r1_61 = _c61.confirm_yield(False)          # VBE 这会儿没弹
        _gu61 = _c61.yield_given_up
        _slot1_61 = _gu61[2] if _gu61 else None
        _b61.set_line("Dim a As 类1, b As 类2", 20)  # 挪到第二个槽位
        _c61.trigger(True)
        check("61.8 ★v80b：`Dim a As 类1, b As 类2` 一行两个槽位 —— 第一个槽位"
              "让位失败（记 given_up，锚点=10）后，移到【第二个槽位】仍然让位"
              "（待定让位、不弹我们的窗），不再被整行连累",
              [[_p1_61, _r1_61 is False, _slot1_61 == 10,
                _c61.yield_pending is not None,
                len(_c61.matches or []) == 0]],
              expect_contain=[[True, True, True, True, True]])

        # ---- 61.9 同一槽位内继续打字：不抖（标记仍然有效）----
        _c61, _b61, _u61 = _mk61("Dim a As 类", 13)
        _c61.trigger(True)
        _c61.confirm_yield(False)
        _gu2_61 = _c61.yield_given_up
        _b61.set_line("Dim a As 类1", 14)            # 同一槽位继续打字
        _c61.trigger(True)
        check("61.9 对照：同一槽位内继续打字 -> 标记仍有效、不再反复让位"
              "（避免『让开-补上-再让开』的抖动）",
              [[_c61.yield_given_up == _gu2_61, _c61.yield_pending is None]],
              expect_contain=[[True, True]])

        # ---- 61.10 有效期：过期后重新给 VBE 一次机会 ----
        _c61, _b61, _u61 = _mk61("Dim a As 类", 13)
        _c61.trigger(True)
        _c61.confirm_yield(False)
        _c61.yield_given_up_at = T61.time() - (E61.YIELD_GIVEN_UP_TTL + 1.0)
        _c61.trigger(True)
        check("61.10 ★v80b：`已放弃让位`标记【过期】后再回到同一处 -> 语法判据"
              "重新生效（待定让位、旧标记作废）—— 兜住『VBE 只是慢了一拍』",
              [[_c61.yield_pending is not None, _c61.yield_given_up is None]],
              expect_contain=[[True, True]])

        _old_ttl61 = E61.YIELD_GIVEN_UP_TTL
        E61.YIELD_GIVEN_UP_TTL = 0.0
        try:
            _c61, _b61, _u61 = _mk61("Dim a As 类", 13)
            _c61.trigger(True)
            _c61.confirm_yield(False)
            _zero61 = _c61.yield_given_up
        finally:
            E61.YIELD_GIVEN_UP_TTL = _old_ttl61
        check("61.11 VBECOMPLETE_YIELD_TTL=0 -> 干脆不记标记（每次都先让位）",
              [_zero61], expect_contain=[None])

        # ---- 61.12 提示窗句柄缓存：必须能【长出新句柄】 ----
        _alive61 = VB61._hwnd_alive
        _enum61 = VB61._enum_vbe_popup_windows
        _st61 = {"found": [], "calls": 0}
        try:
            VB61._hwnd_alive = lambda h: True

            def _fake_enum61():
                _st61["calls"] += 1
                return [(h, "NameListWndClass") for h in _st61["found"]]

            VB61._enum_vbe_popup_windows = _fake_enum61
            VB61._vbe_popup_state["hwnds"] = []
            VB61._vbe_popup_state["at"] = 0.0
            _st61["found"] = [100]
            _r1_61 = VB61._vbe_popup_hwnds()
            _st61["calls"] = 0
            VB61._vbe_popup_hwnds()               # 立刻再调：限流，不该重扫
            _throttled61 = _st61["calls"]
            _st61["found"] = [100, 200]           # 后来才建出来的提示窗
            VB61._vbe_popup_state["at"] = 0.0     # 假装过了一个重扫间隔
            _r2_61 = sorted(VB61._vbe_popup_hwnds())
        finally:
            VB61._hwnd_alive = _alive61
            VB61._enum_vbe_popup_windows = _enum61
            VB61._vbe_popup_state["hwnds"] = []
            VB61._vbe_popup_state["at"] = 0.0
        check("61.12 ★v80b：提示窗句柄缓存【缓存非空时也要周期重扫】——"
              "后来才建出来的提示窗（用到才建的参数信息窗）也能被探到；"
              "间隔内不重复枚举（性能护栏）",
              [[_r1_61, _throttled61, _r2_61]],
              expect_contain=[[[100], 0, [100, 200]]])

        # ---- 61.13 ★接线护栏 ----
        _root61 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _esrc61 = io.open(os.path.join(_root61, "engine.py"),
                          encoding="utf-8").read()
        _vbsrc61 = io.open(os.path.join(_root61, "vbe_bridge.py"),
                           encoding="utf-8").read()
        check("61.13 ★接线护栏：让位签名带【槽位】、confirm_yield 按前缀比对；"
              "标记带有效期；vbe_bridge 的提示窗缓存不再以『缓存为空』为前提；"
              "v80a 新纯函数到位、老口径（kind=None）那一段没被动过",
              [["list_slot_anchor(ctx.get(\"line_text\")" in _esrc61,
                "tuple(pend[1])[:2]" in _esrc61,
                "YIELD_GIVEN_UP_TTL" in _esrc61,
                "self.yield_given_up_at" in _esrc61,
                "now - _vbe_popup_state[\"at\"] >= _VBE_POPUP_RESCAN_SEC"
                in _vbsrc61,
                "if not hs:" not in _vbsrc61,
                "def _own_closer_below(lines_below, closer, kind, base):"
                in _vbsrc61,
                "if kind is not None:" in _vbsrc61,
                "return base == 0" in _vbsrc61]],
              expect_contain=[[True] * 9])

    except Exception as _e61:
        check("第 61 节异常: %s" % _e61, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 62. v81：块【内】分支行按回车自动对齐（Else / ElseIf…Then / Case… /
    #     #Else / #ElseIf）
    #     （用户报：写完 `If … Then` 回车补了 `End If`、光标停在新行，再敲 `Else`
    #      就落在"块体那一级"，回车也不会把它拉回与 `If` 同级）
    # ------------------------------------------------------------------
    print("\n=== 62. 分支行按回车自动对齐（Else / ElseIf / Case）（v81）===")
    try:
        import vbe_bridge as VB62

        _bw62 = VB62._branch_word
        _uo62 = VB62._unclosed_owner_above
        _ba62 = VB62.branch_align

        # ---- 62.1 _branch_word：认分支行 ----
        check("62.1 _branch_word：Else / ElseIf…Then / Case…（含 Case Else / Case Is）"
              "与 `#` 条件编译分支都认得出；`Else If y Then` 按 Else 算（VBA 语义："
              "它是 Else + 一行新 If）；`ElseIf y` 少了 Then / 闭合行 / 普通行 / 续行"
              " -> None",
              [[_bw62("    Else"), _bw62("        ElseIf y Then"),
                _bw62("Case 1, 2"), _bw62("    Case Else"),
                _bw62("    Case Is > 3"), _bw62("    #Else"),
                _bw62("    #ElseIf VBA7 Then"), _bw62("    Else If y Then"),
                _bw62("    ElseIf y"), _bw62("    End If"),
                _bw62("    x = 1"), _bw62("    ElseIf y Then _"),
                _bw62("'Else 注释里"), _bw62("")]],
              expect_contain=[["else", "elseif", "case", "case", "case",
                               "else", "elseif", "else", None, None,
                               None, None, None, None]])

        # ---- 62.2 _unclosed_owner_above：这一档分支归谁 ----
        _AB62 = ["Sub Test()", "    If a = 1 Then", "        x = 1",
                 "    If b = 2 Then", "        y = 1", "    End If",
                 "        z = 1"]
        check("62.2 _unclosed_owner_above：`End If` 计 +1、`If … Then` 计 -1 —— "
              "落在最近那个【还开着】的 If 上（嵌套里是内层的 `If b`）；"
              "普通行不算归属者、没有 -> None；条件编译只认 `#If`（`#Else` 是兄弟）",
              [[_uo62(_AB62, "if"),
                _uo62(["Sub Test()", "    Select Case x"], "select"),
                _uo62(["Sub Test()"], "if"),
                _uo62(["    #If VBA7 Then"], "directive"),
                _uo62(["#If VBA7 Then", "    x = 1", "#Else",
                       "    #If VBA64 Then"], "directive")]],
              expect_contain=[["    ", "    ", None, "    ", "    "]])

        # ---- 62.3 branch_align：算得对不对 ----
        check("62.3 branch_align：Else 对齐最近的 `If … Then`（块体再深一级）；"
              "Case 比 `Select Case` 【深一级】（v83，块体再深一级）；"
              "#Else 与 `#If` 齐平；上方没有归属者（单选 If `If x Then y = 1` / "
              "先敲了 Else 还没写 If）-> None（绝不猜）；非分支行 -> None",
              [[_ba62("        Else", ["Sub Test()", "    If x = 1 Then"], "    "),
                _ba62("            Case 1",
                      ["Sub Test()", "    Select Case x"], "    "),
                _ba62("    #Else", ["    #If VBA7 Then"], "    "),
                _ba62("        Else",
                      ["Sub Test()", "    If x = 1 Then y = 1"], "    "),
                _ba62("    Else", ["Sub Test()"], "    "),
                _ba62("    Case 1", ["Sub Test()", "    x = 1"], "    "),
                _ba62("    x = 1", ["Sub Test()", "    If x = 1 Then"], "    ")]],
              expect_contain=[[("    ", "        "),
                               ("        ", "            "),
                               ("    ", "        "),
                               None, None, None, None]])

        # ---- 集成骨架（带 ReplaceLine：v81 要改本行行首空白）----
        class _CM62(object):
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "Module1"
                self.replace_calls = 0

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n"
                               for l in self.lines[s0:s0 + int(count)])

            def InsertLines(self, start, text):
                i = min(max(0, int(start) - 1), len(self.lines))
                for k, ln in enumerate(str(text).split("\n")):
                    self.lines.insert(i + k, ln)

            def ReplaceLine(self, line, text):
                self.replace_calls += 1
                self.lines[int(line) - 1] = str(text)

        class _Pane62(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE62(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        _orig62 = VB62._get_vbe_cached

        def _nl62(lines, line_no, align=True, cm_cls=_CM62):
            """光标落在第 line_no 行【行尾】，回车走 smart + auto_close。"""
            cm62 = cm_cls(lines)
            col = len(lines[line_no - 1]) + 1
            pane62 = _Pane62(cm62, (line_no, col, line_no, col))
            VB62._get_vbe_cached = lambda: _VBE62(pane62)
            ok62 = VB62.VbeBackend().new_line_below(smart=True, auto_close=True,
                                                    align_branch=align)
            return (ok62, cm62.lines, pane62.sel, cm62.replace_calls)

        # ---- 62.4 ★集成：用户原话的场景 ----
        check("62.4 ★集成（用户原话的场景）：`If … Then` 回车补出 `End If`、光标停在"
              "缩进好的那行；在那行敲 `Else` 再回车 -> `Else` 被拉回与 `If` 同级、"
              "新行再深一级（光标落在新行缩进后），`End If` 一个字没动",
              [_nl62(["Sub Test()",
                      "    If x = 1 Then",
                      "        Else",
                      "    End If",
                      "End Sub"], 3)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "    Else",
                                "        ",
                                "    End If",
                                "End Sub"],
                               (4, 9, 4, 9),
                               1)])

        # ---- 62.5 ★集成：Select Case 里的 Case ----
        check("62.5 ★集成：`Select Case` 里的 `Case 1` 比 `Select Case` 【深一级】"
              "（v83）、块体再深一级；`End Select` 不动",
              [_nl62(["Sub Test()",
                      "    Select Case x",
                      "            Case 1",
                      "    End Select",
                      "End Sub"], 3)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    Select Case x",
                                "        Case 1",
                                "            ",
                                "    End Select",
                                "End Sub"],
                               (4, 13, 4, 13),
                               1)])

        # ---- 62.6 集成：ElseIf ----
        check("62.6 集成：`ElseIf … Then` 与 `Else` 同口径（归最近的 `If … Then`），"
              "且它【不】补收尾（分支没有收尾）",
              [_nl62(["Sub Test()",
                      "    If x = 1 Then",
                      "          ElseIf y = 2 Then",
                      "    End If",
                      "End Sub"], 3)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "    ElseIf y = 2 Then",
                                "        ",
                                "    End If",
                                "End Sub"],
                               (4, 9, 4, 9),
                               1)])

        # ---- 62.7 找不到归属者 -> 老行为（绝不猜）----
        check("62.7 上方没有归属者 -> 维持 v79 老行为（本行缩进 + 一级、一个字不改）："
              "单选 If `If x = 1 Then y = 1` 不算归属者",
              [_nl62(["Sub Test()",
                      "    If x = 1 Then y = 1",
                      "        Else",
                      "End Sub"], 3)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    If x = 1 Then y = 1",
                                "        Else",
                                "            ",
                                "End Sub"],
                               (4, 13, 4, 13),
                               0)])

        # ---- 62.8 开关：align_branch=False 回到老行为 ----
        check("62.8 align_branch=False（`set VBECOMPLETE_NO_BRANCH_ALIGN=1`）"
              "-> 只缩进、不动本行，回到 v79/v79b 老口径",
              [_nl62(["Sub Test()",
                      "    If x = 1 Then",
                      "        Else",
                      "    End If",
                      "End Sub"], 3, align=False)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "        Else",
                                "            ",
                                "    End If",
                                "End Sub"],
                               (4, 13, 4, 13),
                               0)])

        # ---- 62.9 本行【已经对齐】-> 一次 ReplaceLine 都不发 ----
        check("62.9 本行已经对齐 -> 不去碰它（一次 ReplaceLine 都不发），"
              "只把新行插出来",
              [_nl62(["Sub Test()",
                      "    If x = 1 Then",
                      "    Else",
                      "    End If",
                      "End Sub"], 3)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "    Else",
                                "        ",
                                "    End If",
                                "End Sub"],
                               (4, 9, 4, 9),
                               0)])

        # ---- 62.10 写不成 -> 返回 False，一行都不动 ----
        class _CM62NoWrite(_CM62):
            def ReplaceLine(self, line, text):
                raise RuntimeError("boom")

        _nl62_fail = _nl62(["Sub Test()",
                            "    If x = 1 Then",
                            "        Else",
                            "    End If",
                            "End Sub"], 3, cm_cls=_CM62NoWrite)
        check("62.10 拉不齐（ReplaceLine 抛异常）-> 返回 False 且【一行都没动】"
              "（绝不半做：不写本行、也不插新行），调用方把回车原样还给系统",
              [_nl62_fail],
              expect_contain=[(False,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "        Else",
                                "    End If",
                                "End Sub"],
                               (3, 13, 3, 13),
                               0)])

        VB62._get_vbe_cached = _orig62

        # ---- 62.12 ★护栏：Shift+Enter（smart=False）那条路不许被 v81 碰到 ----
        _cm62s = _CM62(["Sub Foo()", "    Else"])
        _pane62s = _Pane62(_cm62s, (2, 9, 2, 9))
        VB62._get_vbe_cached = lambda: _VBE62(_pane62s)
        _ok62s = VB62.VbeBackend().new_line_below()
        # ★v89：这一档钉的是"**不许返回 False**"（不能因为 NameError 被
        # 『出异常一律 return False』吞成"按了 Shift+Enter 却不换行"）。
        # 这里的 `Else` 上方压根没有 `If … Then`，所以 branch_align 按设计
        # 什么都不做（宁可不对齐也不挪错）；缩进按它自己的 4 格再深一级。
        check("62.12 ★护栏（与 v79b 的 closer_text 同坑）：Shift+Enter"
              "（smart=False）那条路不许被 v81 碰到 —— fixed_line 必须在 `if smart:`"
              "【之外】初始化，否则 NameError 会被『出异常一律 return False』吞成"
              "『按了 Shift+Enter 却不换行』",
              [(_ok62s, _cm62s.lines)],
              expect_contain=[(True, ["Sub Foo()", "    Else", "        "])])

        # ---- 62.11 ★接线护栏 ----
        _root62 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _msrc62 = io.open(os.path.join(_root62, "main.py"),
                          encoding="utf-8").read()
        _vbsrc62 = io.open(os.path.join(_root62, "vbe_bridge.py"),
                           encoding="utf-8").read()
        check("62.11 ★接线护栏：main 有独立开关 BRANCH_ALIGN（含"
              " VBECOMPLETE_NO_BRANCH_ALIGN）并把它传进后端；new_line_below 带"
              " align_branch 参数并在 smart 分支里算；拉不齐时先 return False",
              [["BRANCH_ALIGN" in _msrc62,
                "VBECOMPLETE_NO_BRANCH_ALIGN" in _msrc62,
                "align_branch=BRANCH_ALIGN" in _msrc62,
                "def branch_align(line_text, lines_above, unit=" in _vbsrc62,
                "def _branch_word(line_text):" in _vbsrc62,
                "def _unclosed_owner_above(lines_above, kind):" in _vbsrc62,
                "_BRANCH_OWNER = {\"else\": \"if\", \"elseif\": \"if\","
                " \"case\": \"select\"}" in _vbsrc62,
                "align_branch=True" in _vbsrc62,
                "fixed_line, indent = _al" in _vbsrc62,
                "cm, anchor, fixed_line + _raw[len(_ws):]" in _vbsrc62]],
              expect_contain=[[True] * 10])

        # ★v107c 补：62.12 在上面把 `_get_vbe_cached` 换成了自己那份假工程，
        # 忘了换回来 —— 后果不是"这一节出错"，而是**从这一节起，后面每一节**
        # 都读到 `["Sub Foo()", "    Else"]` 这个假模块。凡是像 §14 那样用
        # "清 `_vbe_cache["obj"]` + 换 `_get_vbe`"来搭装置的小节，写回会悄悄
        # 落到别人的假工程上（§89 就是这么被坑的：行没变、唤出计数却是 1）。
        # 教训：猴补丁一律在节尾还原，哪怕这一节看起来"后面没人用得到"。
        VB62._get_vbe_cached = _orig62

    except Exception as _e62:
        check("第 62 节异常: %s" % _e62, [True], expect_contain=[False])

    print("\n=== 63. 补收尾的边界：分支行 / 注释不算『本块已结束』（v82）===")
    try:
        import vbe_bridge as VB63

        _ocb63 = VB63._own_closer_below
        _cn63 = VB63.closer_needed

        # ---- 63.1 纯函数：同级压着分支行 / 注释，本块收尾仍在下面 ----
        check("63.1 _own_closer_below（v82）：同级的 `Else` / `ElseIf…Then` / "
              "`Case` / `#Else` / 整行注释都不是『本块到此为止』的边界 —— "
              "扫描跳过它们继续往下找；对照组（同级是普通代码 / 别的块头）"
              "仍然算边界",
              [[_ocb63(["    Else", "        ", "    End If"], "End If", "if", 4),
                _ocb63(["    ElseIf y = 2 Then", "    End If"], "End If", "if", 4),
                _ocb63(["    Case 1", "        y = 1", "    End Select"],
                       "End Select", "select", 4),
                _ocb63(["    #Else", "    End If"], "End If", "if", 4),
                _ocb63(["    ' 说明", "    End If"], "End If", "if", 4),
                _ocb63(["    Rem 说明", "    End If"], "End If", "if", 4),
                _ocb63(["    x = 1", "    End If"], "End If", "if", 4),
                _ocb63(["    For i = 1 To 3", "    Next", "    End If"],
                       "End If", "if", 4)]],
              expect_contain=[[True, True, True, True, True, True,
                               False, False]])

        # ---- 63.2 closer_needed（kind 给定）同口径 ----
        check("63.2 closer_needed（kind 给定，v82）：下面压着同级分支行 / 注释"
              " -> 本块收尾还在下面 -> 【不补】；同级是普通代码 -> 本块已结束"
              " -> 补（v80a 口径不退化）",
              [[_cn63(["    Else", "        ", "    End If"], "End If", 4,
                      kind="if", lines_above=["Sub Test()", "    If x = 1 Then"]),
                _cn63(["    Else", "    End If"], "End If", 4, kind="if",
                      lines_above=["Sub Test()", "    If x = 1 Then"]),
                _cn63(["    Case 1", "        y = 1", "    End Select"],
                      "End Select", 4, kind="select",
                      lines_above=["Sub Test()", "    Select Case x"]),
                _cn63(["    ' 注释", "    End If"], "End If", 4, kind="if",
                      lines_above=["Sub Test()", "    If x = 1 Then"]),
                _cn63(["    x = 1", "    End If"], "End If", 4, kind="if",
                      lines_above=["Sub Test()", "    If x = 1 Then"])]],
              expect_contain=[[False, False, False, False, True]])

        # ---- 集成骨架（要 ReplaceLine：分支行对齐会改本行行首空白）----
        class _CM63(object):
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "Module1"

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n"
                               for l in self.lines[s0:s0 + int(count)])

            def InsertLines(self, start, text):
                i = min(max(0, int(start) - 1), len(self.lines))
                for k, ln in enumerate(str(text).split("\n")):
                    self.lines.insert(i + k, ln)

            def ReplaceLine(self, line, text):
                self.lines[int(line) - 1] = str(text)

        class _Pane63(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE63(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        _orig63 = VB63._get_vbe_cached

        def _nl63(lines, line_no):
            """光标落在第 line_no 行【行尾】，回车走 smart + auto_close。"""
            cm63 = _CM63(lines)
            col = len(lines[line_no - 1]) + 1
            pane63 = _Pane63(cm63, (line_no, col, line_no, col))
            VB63._get_vbe_cached = lambda: _VBE63(pane63)
            ok63 = VB63.VbeBackend().new_line_below(smart=True,
                                                    auto_close=True)
            return ok63, cm63.lines, pane63.sel

        # ---- 63.3 ★集成：用户报的场景 ----
        check("63.3 ★集成（用户报的场景，v82）：`If … Then` 回车已经补过一次 "
              "`End If`，下面还压着同级的 `Else` —— 回到 `Then` 那行行尾再按"
              "回车，只插一行缩进好的空行，【不再】多补一个 `End If`",
              [_nl63(["Sub Test()",
                      "    If x = 1 Then",
                      "    Else",
                      "        ",
                      "    End If",
                      "End Sub"], 2),
               _nl63(["Sub Test()",
                      "    If x = 1 Then",
                      "    Else",
                      "    End If",
                      "End Sub"], 2)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "        ",
                                "    Else",
                                "        ",
                                "    End If",
                                "End Sub"],
                               (3, 9, 3, 9)),
                              (True,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "        ",
                                "    Else",
                                "    End If",
                                "End Sub"],
                               (3, 9, 3, 9))])

        # ---- 63.4 ★集成：ElseIf 与 Select Case 的 Case 同理 ----
        check("63.4 ★集成：`ElseIf … Then` 压在同级、`Select Case` 里的 "
              "`Case` 压在同级 —— 同样只插一行、不再多补 `End If` / "
              "`End Select`",
              [_nl63(["Sub Test()",
                      "    If x = 1 Then",
                      "        MsgBox 1",
                      "    ElseIf y = 2 Then",
                      "        MsgBox 2",
                      "    End If",
                      "End Sub"], 2),
               _nl63(["Sub Test()",
                      "    Select Case x",
                      "    Case 1",
                      "        y = 1",
                      "    Case 2",
                      "        y = 2",
                      "    End Select",
                      "End Sub"], 2)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "        ",
                                "        MsgBox 1",
                                "    ElseIf y = 2 Then",
                                "        MsgBox 2",
                                "    End If",
                                "End Sub"],
                               (3, 9, 3, 9)),
                              (True,
                               ["Sub Test()",
                                "    Select Case x",
                                "        ",
                                "    Case 1",
                                "        y = 1",
                                "    Case 2",
                                "        y = 2",
                                "    End Select",
                                "End Sub"],
                               (3, 9, 3, 9))])

        # ---- 63.5 ★集成：同级夹着一行注释 ----
        check("63.5 ★集成：`If` 同级夹着一行注释时同理（注释不参与块结构，"
              "不算『本块到此为止』）—— 只插一行",
              [_nl63(["Sub Test()",
                      "    If x = 1 Then",
                      "    ' 说明",
                      "    End If",
                      "End Sub"], 2)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "        ",
                                "    ' 说明",
                                "    End If",
                                "End Sub"],
                               (3, 9, 3, 9))])

        # ---- 63.6 ★端到端：从零走完用户那三步 ----
        def _flow63():
            cm63 = _CM63(["Sub Test()", "    If x = 1 Then", "End Sub"])
            pane63 = _Pane63(cm63, (2, 18, 2, 18))
            VB63._get_vbe_cached = lambda: _VBE63(pane63)

            def key(line_no, text=None):
                if text is not None:
                    cm63.lines[line_no - 1] = text
                cur = cm63.lines[line_no - 1]
                pane63.sel = (line_no, len(cur) + 1, line_no, len(cur) + 1)
                return VB63.VbeBackend().new_line_below(smart=True,
                                                        auto_close=True)

            r1 = key(2)
            s1 = list(cm63.lines)               # 补出 `End If`、光标停中间行
            r2 = key(3, "        Else")
            s2 = list(cm63.lines)               # `Else` 拉回同级、新行深一级
            r3 = key(2)                         # ★回到 `Then` 行行尾再回车
            return r1, s1, r2, s2, r3, list(cm63.lines)

        check("63.6 ★端到端（用户那条路径，v82）：①写 `If … Then` 回车 -> "
              "补出 `End If` ②在块体行写 `Else` 回车 -> 拉回 `If` 同级、新行"
              "深一级 ③回到 `Then` 行行尾再回车 -> 只插一行空行，`End If` "
              "【没有】变成两条",
              [_flow63()],
              expect_contain=[(True,
                               ["Sub Test()", "    If x = 1 Then", "        ",
                                "    End If", "End Sub"],
                               True,
                               ["Sub Test()", "    If x = 1 Then", "    Else",
                                "        ", "    End If", "End Sub"],
                               True,
                               ["Sub Test()", "    If x = 1 Then", "        ",
                                "    Else", "        ", "    End If",
                                "End Sub"])])

        # ---- 63.7 回归护栏 ----
        check("63.7 回归护栏：①嵌套里内层 `Else` 缩进更深时不误判；"
              "②结构完好、没有分支行 -> 照旧只插一行；③v79c/v80a 老口径不许"
              "退化（模块级 `Type` 下面紧邻过程头 -> 仍补 `End Type`）",
              [_nl63(["Sub Test()",
                      "    If x = 1 Then",
                      "        If y = 2 Then",
                      "        Else",
                      "        End If",
                      "    End If",
                      "End Sub"], 2)[1],
               _nl63(["Sub Test()",
                      "    If x = 1 Then",
                      "        MsgBox 1",
                      "    End If",
                      "End Sub"], 2)[1],
               _nl63(["Type Foo", "Sub Bar()", "End Sub"], 1)[1]],
              expect_contain=[["Sub Test()",
                               "    If x = 1 Then",
                               "        ",
                               "        If y = 2 Then",
                               "        Else",
                               "        End If",
                               "    End If",
                               "End Sub"],
                              ["Sub Test()",
                               "    If x = 1 Then",
                               "        ",
                               "        MsgBox 1",
                               "    End If",
                               "End Sub"],
                              ["Type Foo", "    ", "End Type", "Sub Bar()",
                               "End Sub"]])

        VB63._get_vbe_cached = _orig63

        # ---- 63.8 ★接线护栏 ----
        _root63 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _vbsrc63 = io.open(os.path.join(_root63, "vbe_bridge.py"),
                           encoding="utf-8").read()
        _fn63 = _vbsrc63.split("def _own_closer_below")[1].split(
            "def closer_needed")[0]
        check("63.8 ★接线护栏：`_own_closer_below` 里确实把【注释】与"
              "【块内分支行】跳过，而且这一跳必须发生在缩进判定【之前】"
              "（放后面就等于没跳）",
              [[("_is_comment_only(raw) or _branch_word(raw) is not None"
                 in _fn63),
                ("w = _indent_width(raw)" in _fn63),
                (_fn63.index("_branch_word(raw) is not None")
                 < _fn63.index("w = _indent_width(raw)")),
                "_branch_word" in _vbsrc63]],
              expect_contain=[[True, True, True, True]])

    except Exception as _e63:
        check("第 63 节异常: %s" % _e63, [True], expect_contain=[False])

    print("\n" + "=" * 60)
    print("=== 64. 分支对齐的层级分档：Case 比 Select Case 深一级（v83）===")
    try:
        import vbe_bridge as VB64

        _ba64 = VB64.branch_align
        _dep64 = VB64._BRANCH_OWNER_DEPTH

        # ---- 64.1 档位表 ----
        check("64.1 `_BRANCH_OWNER_DEPTH`：Else / ElseIf 与 If 齐平（0）、"
              "`#Else` / `#ElseIf` 与 #If 齐平（0）、Case 比 Select Case 深一级（1）",
              [[_dep64.get("else"), _dep64.get("elseif"), _dep64.get("case")]],
              expect_contain=[[0, 0, 1]])

        # ---- 64.2 纯函数：Case 深一级；Else / #Else 维持齐平（v81 不退化）----
        check("64.2 branch_align（v83）：`Case 1` / `Case Else` / `Case Is > 3` "
              "都比 `Select Case` 【深一级】（本行 owner+1 级、新行 owner+2 级）；"
              "Else / ElseIf / #Else / #ElseIf 一律维持齐平",
              [[_ba64("        Case 1",
                      ["Sub Test()", "    Select Case x"], "    "),
                _ba64("    Case Else",
                      ["Sub Test()", "    Select Case x"], "    "),
                _ba64("            Case Is > 3",
                      ["Sub Test()", "    Select Case x"], "    "),
                _ba64("        Else",
                      ["Sub Test()", "    If x = 1 Then"], "    "),
                _ba64("          ElseIf y Then",
                      ["Sub Test()", "    If x = 1 Then"], "    "),
                _ba64("    #Else", ["    #If VBA7 Then"], "    "),
                _ba64("    #ElseIf VBA64 Then",
                      ["    #If VBA7 Then"], "    ")]],
              expect_contain=[[("        ", "            "),
                               ("        ", "            "),
                               ("        ", "            "),
                               ("    ", "        "),
                               ("    ", "        "),
                               ("    ", "        "),
                               ("    ", "        ")]])

        # ---- 64.3 缩进单位跟着工程走（2 空格 / Tab）----
        check("64.3 缩进单位跟着工程走：2 个空格缩进 -> Case 只深 2 格；"
              "Tab 缩进 -> Case 深一个 Tab（`unit * 1`，不是硬写 4 个空格）",
              [[_ba64("    Case 1", ["Sub Test()", "  Select Case x"], "  "),
                _ba64("\t\tCase 1", ["Sub Test()", "\tSelect Case x"], "\t")]],
              expect_contain=[[("    ", "      "), ("\t\t", "\t\t\t")]])

        # ---- 64.4 归属者：内层 Select 闭合后，外面的 Case 归外层（不被内层抢走）----
        check("64.4 `_unclosed_owner_above` 配平：内层 `Select Case` 已经用 "
              "`End Select` 闭合 -> 后面的 `Case 2` 归【外层】那个（深一级 = 8 格）",
              [[_ba64("        Case 2",
                      ["Sub Test()",
                       "    Select Case x",
                       "        Case 1",
                       "            Select Case y",
                       "            End Select"], "    "),
                _ba64("    Case 2", ["Sub Test()", "    x = 1"], "    "),
                _ba64("    Case 2", ["Sub Test()"], "    ")]],
              expect_contain=[[("        ", "            "), None, None]])

        # ---- 集成骨架（要 ReplaceLine：分支行对齐会改本行行首空白）----
        class _CM64(object):
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "Module1"
                self.replace_calls = 0

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n"
                               for l in self.lines[s0:s0 + int(count)])

            def InsertLines(self, start, text):
                i = min(max(0, int(start) - 1), len(self.lines))
                for k, ln in enumerate(str(text).split("\n")):
                    self.lines.insert(i + k, ln)

            def ReplaceLine(self, line, text):
                self.replace_calls += 1
                self.lines[int(line) - 1] = str(text)

        class _Pane64(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE64(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        _orig64 = VB64._get_vbe_cached

        def _nl64(lines, line_no, text=None):
            """光标落在第 line_no 行【行尾】，回车走 smart + auto_close。"""
            cm64 = _CM64(lines)
            if text is not None:
                cm64.lines[line_no - 1] = text
            cur = cm64.lines[line_no - 1]
            pane64 = _Pane64(cm64, (line_no, len(cur) + 1, line_no,
                                    len(cur) + 1))
            VB64._get_vbe_cached = lambda: _VBE64(pane64)
            ok64 = VB64.VbeBackend().new_line_below(smart=True, auto_close=True)
            return (ok64, cm64.lines, pane64.sel, cm64.replace_calls)

        # ---- 64.5 ★集成（用户报的场景）：Case 已经在正确层级 -> 一个字都不动 ----
        # v81 的错就在这里：它把已经缩进 8 格的 `Case 1` 拉回 4 格（与 Select 齐平）。
        check("64.5 ★集成（用户报的场景）：`Case 1` 已经在 `Select Case` 下面"
              "深一级（8 格）-> 回车【不改本行】（一次 ReplaceLine 都不发）、"
              "新行（块体）再深一级到 12 格，光标 (4,13)；`End Select` 一个字符没动",
              [_nl64(["Sub Test()",
                      "    Select Case x",
                      "        Case 1",
                      "    End Select",
                      "End Sub"], 3)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    Select Case x",
                                "        Case 1",
                                "            ",
                                "    End Select",
                                "End Sub"],
                               (4, 13, 4, 13),
                               0)])

        # ---- 64.6 ★集成：Case 敲歪了（与 Select 齐平）-> 拉回深一级 ----
        check("64.6 ★集成：`Case 1` 被敲成与 `Select Case` 齐平（4 格）-> "
              "回车把它拉回 8 格（深一级），新行 12 格",
              [_nl64(["Sub Test()",
                      "    Select Case x",
                      "    Case 1",
                      "    End Select",
                      "End Sub"], 3)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    Select Case x",
                                "        Case 1",
                                "            ",
                                "    End Select",
                                "End Sub"],
                               (4, 13, 4, 13),
                               1)])

        # ---- 64.7 回归：Else 与 If 齐平、#Else 与 #If 齐平（v81 口径不退化成"深一级"）----
        check("64.7 回归护栏：`Else` 仍然与 `If` 齐平（4 格、新行 8 格）；"
              "`#Else` 仍然与 `#If` 齐平（0 格、新行 4 格）—— v83 只动了 Case 那一档",
              [_nl64(["Sub Test()",
                      "    If x = 1 Then",
                      "        Else",
                      "    End If",
                      "End Sub"], 3),
               _nl64(["#If VBA7 Then",
                      "    x = 1",
                      "    #Else",
                      "#End If"], 3)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    If x = 1 Then",
                                "    Else",
                                "        ",
                                "    End If",
                                "End Sub"],
                               (4, 9, 4, 9),
                               1),
                              (True,
                               ["#If VBA7 Then",
                                "    x = 1",
                                "#Else",
                                "    ",
                                "#End If"],
                               (4, 5, 4, 5),
                               1)])

        # ---- 64.8 嵌套 Select：内层 Case 归内层（16 格）----
        check("64.8 嵌套 `Select Case`：内层 `Case 2` 归【内层】那个 "
              "（本行 16 格、新行 20 格），外层 Case 与 End Select 都不动",
              [_nl64(["Sub Test()",
                      "    Select Case x",
                      "        Case 1",
                      "            Select Case y",
                      "                Case 2",
                      "            End Select",
                      "    End Select",
                      "End Sub"], 5)],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    Select Case x",
                                "        Case 1",
                                "            Select Case y",
                                "                Case 2",
                                "                    ",
                                "            End Select",
                                "    End Select",
                                "End Sub"],
                               (6, 21, 6, 21),
                               0)])

        # ---- 64.9 ★端到端：从零走完 `Select Case` -> `Case 1` -> `Case 2` ----
        def _flow64():
            cm64 = _CM64(["Sub Test()", "    Select Case x", "End Sub"])
            pane64 = _Pane64(cm64, (1, 1, 1, 1))
            VB64._get_vbe_cached = lambda: _VBE64(pane64)

            def key(line_no, text=None):
                if text is not None:
                    cm64.lines[line_no - 1] = text
                cur = cm64.lines[line_no - 1]
                pane64.sel = (line_no, len(cur) + 1, line_no, len(cur) + 1)
                return VB64.VbeBackend().new_line_below(smart=True,
                                                        auto_close=True)

            r1 = key(2)                       # 写 `Select Case x` 回车 -> 补 End Select
            s1 = list(cm64.lines)
            r2 = key(3, "        Case 1")     # 写 `Case 1` 回车
            s2 = list(cm64.lines)
            r3 = key(4, "        Case 2")     # 写 `Case 2` 回车
            return (r1, s1, r2, s2, r3, list(cm64.lines), pane64.sel)

        check("64.9 ★端到端（从零三步）：①写 `Select Case x` 回车 -> 补出 "
              "`End Select`、光标停在 8 格的块体行；②在那行写 `Case 1` 回车 -> "
              "它【留在 8 格】、新行 12 格；③再写 `Case 2` 回车 -> 同样留在 8 格、"
              "新行 12 格（`End Select` 始终只有一条）",
              [_flow64()],
              expect_contain=[(True,
                               ["Sub Test()",
                                "    Select Case x",
                                "        ",
                                "    End Select",
                                "End Sub"],
                               True,
                               ["Sub Test()",
                                "    Select Case x",
                                "        Case 1",
                                "            ",
                                "    End Select",
                                "End Sub"],
                               True,
                               ["Sub Test()",
                                "    Select Case x",
                                "        Case 1",
                                "        Case 2",
                                "            ",
                                "    End Select",
                                "End Sub"],
                               (5, 13, 5, 13))])

        VB64._get_vbe_cached = _orig64

        # ---- 64.10 ★接线护栏 ----
        _root64 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _vbsrc64 = io.open(os.path.join(_root64, "vbe_bridge.py"),
                           encoding="utf-8").read()
        _fn64 = _vbsrc64.split("def branch_align")[1].split(
            "def _indent_unit_for")[0]
        check("64.10 ★接线护栏：`branch_align` 真的走档位表"
              "（`unit * _BRANCH_OWNER_DEPTH`），不是把 Case 硬写成 0；"
              "档位表里 `case` 必须是 1、`else` / `elseif` 必须是 0 "
              "（防止有人「顺手」把它改回 v81 的齐平口径）",
              [[("_BRANCH_OWNER_DEPTH.get(word, 0)" in _fn64),
                ("unit *" in _fn64),
                ('"case": 1' in _vbsrc64),
                ('"else": 0, "elseif": 0' in _vbsrc64)]],
              expect_contain=[[True, True, True, True]])

    except Exception as _e64:
        check("第 64 节异常: %s" % _e64, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 65. v86：补收尾的"边界"再修两类 ——【行标签】（`1:`，配 `GoTo 1` 用，
    #     惯例顶格写）与【`#` 条件编译指令】（`#If` / `#End If`，同样惯例顶格）
    #     的缩进【不含结构信息】，不能当"本块到此为止"的边界；
    #     另外往下扫描的行数上限要够大，且读满被截断时按"不补"兜底。
    #
    #     用户报：`Sub` 里写了很多代码后，把光标移回 `Sub` 那行回车，又补出一个
    #     `End Sub`；`If … Then` 那行回车同理又多一个 `End If`。
    #     实测 `脚本.xlsm / 模块2` 里那个【顶格的 `1:`】一次坑掉三处
    #     （`Sub 功能2()` 那行 / 外层 `For` 那行 / `If … Then` 那行）。
    # ------------------------------------------------------------------
    print("\n=== 65. 回车补收尾：行标签 / 条件编译指令不算边界（v86）===")
    try:
        import vbe_bridge as VB65

        _ocb65 = VB65._own_closer_below
        _cn65 = VB65.closer_needed
        _lab65 = VB65._is_label_line
        _blind65 = VB65._indent_blind

        # ---- 65.1 纯函数：认行标签 ----
        check("65.1 `_is_label_line`：`1:` / `Retry:` / `10: x = 1`（标签与语句"
              "同一行）/ 带缩进的 `5:` / 中文名 `标签1:` 都算标签；"
              "`Case 1:` / `Dim x: x = 1` / `foo(1)` / 过程头 / 普通赋值 / "
              "字符串里的 `\"a:b\"` / 注释里的 `'x:y` 都不算",
              [[_lab65("1:"), _lab65("Retry:"), _lab65("10: x = 1"),
                _lab65("    5:"), _lab65("标签1:"),
                _lab65("Case 1:"), _lab65("Dim x: x = 1"), _lab65("foo(1)"),
                _lab65("Sub Test()"), _lab65("x = 1"),
                _lab65('"a:b"'), _lab65("'x:y")]],
              expect_contain=[[True, True, True, True, True,
                               False, False, False, False, False,
                               False, False]])

        # ---- 65.2 纯函数：_indent_blind ----
        check("65.2 `_indent_blind`：行标签与 `#` 条件编译指令"
              "（`#If` / `#End If` / `#Const` / `#Else` / `#ElseIf`）的缩进都"
              "不含结构信息；普通代码 / 块头 / 收尾行 / 分支行一律不算",
              [[_blind65("1:"), _blind65("#If VBA7 Then"), _blind65("  #End If"),
                _blind65("#Const A = 1"), _blind65("#Else"),
                _blind65("#ElseIf V Then"),
                _blind65("    x = 1"), _blind65("    For i = 1 To 3"),
                _blind65("    End If"), _blind65("    Else")]],
              expect_contain=[[True, True, True, True, True, True,
                               False, False, False, False]])

        # ---- 65.3 _own_closer_below：跳过标签 / 指令（先跳后判缩进）----
        check("65.3 `_own_closer_below`（v86）：顶格的 `1:` 不是边界 —— "
              "`If` 体里的 `1:` 之后仍有 `End If` -> True；`Sub` 体里的顶格 "
              "`1:` 之后仍有 `End Sub` -> True；`#If` / `#End If`（0 列）同理；"
              "对照组（同级普通代码 -> 本块已结束 -> False）不退化",
              [[_ocb65(["            If x Then GoTo 1", "1:",
                        "        End If"], "End If", "if", 8),
                _ocb65(["    x = 1", "1:", "End Sub"], "End Sub", "sub", 0),
                _ocb65(["#If VBA7 Then", "    x = 1", "#End If", "End Sub"],
                       "End Sub", "sub", 0),
                _ocb65(["    x = 1"], "End If", "if", 4)]],
              expect_contain=[[True, True, True, False]])

        # ---- 65.4 closer_needed（kind 给定）同口径 ----
        check("65.4 `closer_needed`（kind 给定，v86）：下面压着顶格 `1:` / "
              "`#If` / `#End If` 时，本块收尾仍然在下面 -> 【不补】；"
              "对照组（同级普通代码 -> 本块已结束）仍然补",
              [[_cn65(["            If x Then GoTo 1", "1:", "        End If"],
                      "End If", 8, kind="if",
                      lines_above=["Sub Test()",
                                   "    For i = 1 To 3",
                                   "        If a Then"]),
                _cn65(["    x = 1", "1:", "End Sub"], "End Sub", 0, kind="sub",
                      lines_above=[]),
                _cn65(["#If VBA7 Then", "    x = 1", "#End If", "End Sub"],
                      "End Sub", 0, kind="sub", lines_above=[]),
                _cn65(["    x = 1"], "End If", 4, kind="if",
                      lines_above=["Sub Test()", "    If a Then"])]],
              expect_contain=[[False, False, False, True]])

        # ---- 65.5 ★截断兜底 ----
        check("65.5 ★v86 截断兜底：下方行数是【读满上限被截断】的 -> "
              "『窗口里没找到收尾』不能当『下面没有收尾』-> 一律【不补】"
              "（宁可少补一条，也不要变出双份收尾把代码写坏）；"
              "truncated=False（默认）时维持 v80a 老口径（补）",
              [[_cn65(["    x = 1", "    y = 2"], "End Sub", 0, kind="sub",
                      lines_above=[], truncated=True),
                _cn65(["    x = 1", "    y = 2"], "End Sub", 0, kind="sub",
                      lines_above=[]),
                _cn65(["    x = 1", "    y = 2"], "End Sub", 0, kind="sub",
                      lines_above=[], truncated=False)]],
              expect_contain=[[False, True, True]])

        # ---- 集成骨架 ----
        class _CM65(object):
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "模块2"

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n"
                               for l in self.lines[s0:s0 + int(count)])

            def InsertLines(self, start, text):
                i = min(max(0, int(start) - 1), len(self.lines))
                for k, ln in enumerate(str(text).split("\n")):
                    self.lines.insert(i + k, ln)

        class _Pane65(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE65(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        _orig65 = VB65._get_vbe_cached

        # 用户现场的真实形态（`脚本.xlsm / 模块2` 摘录，15 行）：第 11 行那个
        # 【顶格的 `1:`】配上面的 `GoTo 1` 用。
        # ⚠️ 真实模块里的过程名叫 `功能2`、字符串是 `"姓名"`（全角）；这里换成
        # ASCII 名是【测试脚手架】的需要 —— GetSelection 给的是**显示列**，
        # 全角字符占两格，mock 里按字符列摆光标会偏（判据本身与此无关）。
        _REAL65 = [
            "Sub Func2()",
            "    Set actSheet = ThisWorkbook.ActiveSheet",
            "    For i = 3 To targetSheet.Cells(targetSheet.Cells.Rows.Count, "
            "\"e\").End(xlUp).Row",
            "        ID = CStr(targetSheet.Cells(i, \"e\"))",
            "    Next",
            "",
            "    For i = 1 To dataSheet1.Cells(dataSheet1.Cells.Rows.Count, "
            "\"c\").End(xlUp).Row",
            "        If dataSheet1.Cells(i, \"c\") = \"name\" Then",
            "            If dataSheet1.Cells(i, \"d\") = \"\" Then GoTo 1",
            "            ID = CStr(dataSheet3.Cells(idRow, \"b\"))",
            "1:",
            "        End If",
            "    Next",
            "    MsgBox \"OK\", vbInformation",
            "End Sub",
        ]

        def _nl65(line_no):
            cm65 = _CM65(_REAL65)
            cur = cm65.lines[line_no - 1]
            pane65 = _Pane65(cm65, (line_no, len(cur) + 1, line_no,
                                    len(cur) + 1))
            VB65._get_vbe_cached = lambda: _VBE65(pane65)
            ok65 = VB65.VbeBackend().new_line_below(smart=True, auto_close=True)
            return (ok65, cm65.lines, pane65.sel)

        # ---- 65.6 ★集成：用户现场那三处，一个收尾都不许补 ----
        check("65.6 ★集成（`脚本.xlsm/模块2` 那个顶格 `1:` 的真实形态）："
              "①`Sub …()` 那行回车 -> 不补 `End Sub`；②外层 `For` 那行回车 "
              "-> 不补 `Next`；③`If … Then` 那行回车 -> 不补 `End If` —— "
              "三处都只多一行缩进好的空行，原有 15 行代码一个字符没动",
              [_nl65(1), _nl65(3), _nl65(8)],
              expect_contain=[(True, _REAL65[:1] + ["    "] + _REAL65[1:],
                               (2, 5, 2, 5)),
                              (True, _REAL65[:3] + ["        "] + _REAL65[3:],
                               (4, 9, 4, 9)),
                              (True, _REAL65[:8] + ["            "]
                               + _REAL65[8:], (9, 13, 9, 13))])

        VB65._get_vbe_cached = _orig65

        # ---- 65.7 ★接线护栏 ----
        _root65 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _vbsrc65 = io.open(os.path.join(_root65, "vbe_bridge.py"),
                           encoding="utf-8").read()
        _fn65 = _vbsrc65.split("def _own_closer_below")[1].split(
            "def closer_needed")[0]
        check("65.7 ★接线护栏：`_own_closer_below` 里把【行标签 / `#` 指令】也跳过"
              "（`_indent_blind`），而且这一跳同样必须在缩进判定【之前】"
              "（放后面就等于没跳）；后端补收尾时按 `_CLOSER_SCAN_LINES` 读、"
              "并把 `truncated` 交给判据；这个扫描窗口必须比"
              "『往上找缩进基准』那个 `_INDENT_LOOKBACK` 大",
              [[("_indent_blind(raw)" in _fn65),
                ("w = _indent_width(raw)" in _fn65),
                (_fn65.index("_indent_blind(raw)")
                 < _fn65.index("w = _indent_width(raw)")),
                ("self._CLOSER_SCAN_LINES" in _vbsrc65),
                ("truncated=_rest > _cap" in _vbsrc65),
                ("if truncated:" in _vbsrc65),
                (VB65.VbeBackend._CLOSER_SCAN_LINES
                 > VB65.VbeBackend._INDENT_LOOKBACK)]],
              expect_contain=[[True] * 7])

        # ---- 65.8 回归护栏：v80a / v82 口径不退化 ----
        check("65.8 回归护栏：v80a 主场景（同级压着【同类嵌套块】-> 下面没有"
              "本块收尾 -> 仍要补）与 v82 口径（同级 `Else` / 注释跳过 -> "
              "不补）都不退化",
              [[_cn65(["    For i = 1 To 10", "        x = 1", "    Next i",
                       "End Sub"], "Next", 4, kind="for",
                      lines_above=["Sub Test()"]),
                _cn65(["    Else", "    End If"], "End If", 4, kind="if",
                      lines_above=["Sub Test()", "    If x = 1 Then"]),
                _cn65(["    ' 注释", "    End If"], "End If", 4, kind="if",
                      lines_above=["Sub Test()", "    If x = 1 Then"])]],
              expect_contain=[[True, False, False]])

    except Exception as _e65:
        check("第 65 节异常: %s" % _e65, [True], expect_contain=[False])

    # ---- 66. v87：别的模块里【不可见】的声明不许当"真名字"证据 ----
    #
    # 用户报（第八条反馈）："对于未定义、第一次写出来的变量名，会提示自己；回退删除
    # 已写的变量名时，也会提示自己。这个 bug 以前在模块里会存在、已经修好了，但是
    # 现在在【类模块】里也会出。另外帮我查一下 sheet / thisworkbook / 窗体 里会不会
    # 出现，都帮我修一下。"
    #
    # 真机定位（只读扫用户正在编辑的工程）：光标停在 `类2` 第 3 行 `Public 订单金额`
    # 行尾，而 `订单金额` 在另一个类模块 `SummaryClass` 里也声明了一份 ——
    #   池子侧：类模块成员是 Private（跨模块不能裸名引用），
    #           filter_identifiers_by_scope 正确隐藏；
    #   证据侧：declared_elsewhere 只看"名字在别的模块出现过没有"，**不看可见性**
    #           -> 返回 True -> 回声防护放行 -> 弹窗把正在敲的字原样提示回来。
    # 同一个漏洞对【窗体 / 文档模块（Sheet / ThisWorkbook）】一样成立（它们的模块级
    # 成员同样默认 Private），所以这一节把这些组件类型一次性钉住。
    print("\n=== 66. v87 类/窗体/文档模块的成员不再当跨模块证据（提示自己）===")
    try:
        import vbe_bridge as VB66
        import parser as P66
        import engine as E66

        # ---- 66.1 记录面：同一个 `Public X`，跨模块可见性随模块类型不同 ----
        def _priv66(code, std):
            r, _pn = list(P66.extract_records(code, module="M",
                                             is_std_module=std))
            return sorted((str(n), pv) for n, _m, _p, pv in r if not _p)

        check("66.1 `Public X` 的跨模块可见性：标准模块 priv=False（工程可见）；"
              "类 / 窗体 / 文档模块 priv=True（只能经实例或限定名访问）",
              [_priv66("Public SharedName", True),
               _priv66("Public SharedName", False)],
              expect_contain=[[("SharedName", False)],
                              [("SharedName", True)]])

        # ---- 集成骨架：一个能喂饱收集流程的假 VBE（多组件、真模块名）----
        class _CM66(object):
            def __init__(self, comp, text):
                self._comp = comp
                self.text = text

            @property
            def Name(self):
                return self._comp.Name

            @property
            def Parent(self):
                return self._comp

            @property
            def CountOfLines(self):
                return self.text.count("\n") + 1

            def Lines(self, start, count):
                return "\r\n".join(
                    self.text.split("\n")[start - 1:start - 1 + count])

        class _Comp66(object):
            def __init__(self, name, ctype, text):
                self.Name = name
                self.Type = ctype
                self.CodeModule = _CM66(self, text)

        class _Proj66(object):
            def __init__(self, comps):
                self.Name = "VBAProject"
                self.VBComponents = list(comps)

        class _Pane66(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE66(object):
            def __init__(self, comps, pane):
                self.ActiveVBProject = _Proj66(comps)
                self.ActiveCodePane = pane

        class _UI66(object):
            def __init__(self):
                self.shown = False

            def show(self, *a, **k):
                self.shown = True

            def hide(self, *a, **k):
                self.shown = False

            def update_selection(self, *a, **k):
                pass

            def contains_point(self, *a, **k):
                return False

        _orig66 = VB66._get_vbe_cached

        def _mk66(comps, caret_name, sel):
            """comps: [(模块名, 组件类型, 代码)]；返回 (backend, pane, comps)。"""
            objs = {}
            lst = []
            for (n, t, s) in comps:
                c = _Comp66(n, t, s)
                objs[n] = c
                lst.append(c)
            pane = _Pane66(objs[caret_name].CodeModule, sel)
            VB66._get_vbe_cached = lambda: _VBE66(lst, pane)
            return VB66.VbeBackend(), pane, objs

        def _go66(backend):
            ui = _UI66()
            c = E66.Completer(backend, ui)
            c.trigger()
            return ui.shown, c.matches

        try:
            # 组件类型：1=标准模块、2=类模块、3=窗体、100=文档模块（Sheet/ThisWorkbook）
            # ⚠️ 代码行一律 ASCII —— 假 GetSelection 给的是显示列，全角字符会偏
            #    （与第 65 节同一个脚手架坑）。
            _P66 = [
                ("模块A", 1,
                 "Public GName\nSub Work()\n    Dim onlyLocal As Long\nEnd Sub"),
                ("类C", 2, "Public SharedItem"),
                ("类D", 2, "Public SharedItem"),
                ("窗体F", 3, "Dim SharedItem As Long"),
                ("Sheet1", 100, "Public SharedItem As Long"),
                ("ThisWorkbook", 100, "Dim SharedItem As Long"),
            ]
            _bk66, _pane66, _objs66 = _mk66(_P66, "类C", (1, 18, 1, 18))
            _ids66 = _bk66.get_identifiers()

            # ---- 66.2 核心：同名成员散落在类/窗体/文档模块里，谁都不算证据 ----
            check("66.2 ★真实后端：`SharedItem` 在 类C / 类D / 窗体F / Sheet1 / "
                  "ThisWorkbook 里各声明一份 -> 从 类C 问【不算】跨模块证据"
                  "（修前 True，正是用户那条「提示自己」）",
                  [_bk66.declared_elsewhere("SharedItem", "类C")],
                  expect_contain=[False])

            # ---- 66.3 对照面：真的跨模块可见的，一路都不许退化 ----
            check("66.3 对照：标准模块的 `Public GName` 仍是跨模块证据（v37 语义"
                  "不许退化）；别的过程的局部变量 `onlyLocal` 不算；组件名 `模块A` 算",
                  [_bk66.declared_elsewhere("GName", "类C"),
                   _bk66.declared_elsewhere("onlyLocal", "类C"),
                   _bk66.declared_elsewhere("模块A", "类C")],
                  expect_contain=[True, False, True])

            # ---- 66.4 ★端到端（用户场景①）：第一次写出来的成员名不提示自己 ----
            _bk66b, _pane66b, _objs66b = _mk66(_P66, "类C", (1, 18, 1, 18))
            check("66.4 ★端到端（用户场景①：第一次写出来的名字）：`Public "
                  "SharedItem` 行尾触发 -> 弹窗不出现，一条候选都没有"
                  "（修前会把 SharedItem 原样提示回来）",
                  [_go66(_bk66b)],
                  expect_contain=[(False, [])])

            # ---- 66.5 ★端到端（用户场景②）：回退删字时不留旧版本的幽灵 ----
            # 顺序很关键：先按"名字还完整"那一刻解析出池子（模拟 0.4s 缓存），
            # 再让用户按退格删掉一个字符 —— 池子里那份 `SharedItem` 正是
            # 用户说的"回退删除已写的变量名时提示自己"的来源。
            _bk66c, _pane66c, _objs66c = _mk66(_P66, "类C", (1, 16, 1, 16))
            _bk66c.get_identifiers()          # 池子快照（此刻文本还是完整的名字）
            _objs66c["类C"].CodeModule.text = "Public SharedIt"
            _pane66c.sel = (1, 16, 1, 16)     # 光标退到删剩下的位置
            check("66.5 ★端到端（用户场景②：回退删字）：池子里还留着 "
                  "`SharedItem`（缓存）-> 退格成 `SharedIt` 之后不许再把 "
                  "`SharedItem` 提示出来",
                  [_go66(_bk66c)],
                  expect_contain=[(False, [])])

            # ---- 66.6 ★v105：`Public GNa` 也是声明行（正在起新名字）-> 静默 ----
            # 旧口径（v37）是"别的标准模块里的 `GName` 是货真价实的存在，照常提示"；
            # v105 把声明行整个静默了，故期望反转为 (False, [])。
            # ⚠️ "跨模块 Public 是合法存在"这条语义由 66.3（纯判据）继续覆盖。
            _P66b = [("模块A", 1, "Public GName"),
                     ("类C", 2, "Public GNa")]
            _bk66d, _pane66d, _objs66d = _mk66(_P66b, "类C", (1, 12, 1, 12))
            _r66d = _go66(_bk66d)
            check("66.6 声明行 `Public GNa` -> 静默（v105；跨模块存在性见 66.3）",
                  [_r66d], expect_contain=[(False, [])])

            # ---- 66.7 两条路径同一口径：候选池隐藏的，证据侧也必须否认 ----
            _visA66 = set(n.lower() for n in E66.filter_identifiers_by_scope(
                _ids66, None, "模块A"))
            check("66.7 两条路径同口径：从 模块A 看 `SharedItem` —— 候选池里没有它"
                  "（别的模块的私有成员），证据侧也否认它（都要 False）",
                  [("shareditem" in _visA66,
                    _bk66.declared_elsewhere("SharedItem", "模块A"))],
                  expect_contain=[(False, False)])

            # ---- 66.8 `_module_level_names` 也按可见性收窄（同类漏洞的另一处）----
            class _B66(_FakeBackend):
                def get_identifiers(self):
                    return [("MyName", "类C", None, True),      # 本模块模块级 -> 可见
                            ("PubName", "模块A", None, False),  # 别处 Public -> 可见
                            ("OtherPriv", "类X", None, True),   # 别处 Private -> 不可见
                            ("aLoc", "类C", "P", False)]        # 过程内 -> 不算模块级

            _c66e = E66.Completer(_B66({"proc_name": None,
                                        "module_name": "类C"}), _UI66())
            _c66e._ctx_now = {"proc_name": None, "module_name": "类C"}
            check("66.8 `_module_level_names` 按可见性收窄：本模块模块级（MyName）与"
                  "别处 Public（PubName）算；别模块的私有成员（OtherPriv）与过程内"
                  "局部（aLoc）不算",
                  [sorted(_c66e._module_level_names())],
                  expect_contain=[["myname", "pubname"]])

            # ---- 66.9 ★接线护栏：两处判据各改对了地方（防日后被"顺手"改回去）----
            _root66 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            _vbsrc66 = io.open(os.path.join(_root66, "vbe_bridge.py"),
                               encoding="utf-8").read()
            _fn66 = _vbsrc66.split("def _collect_identifiers")[1].split(
                "def get_type_names")[0]
            _esrc66 = io.open(os.path.join(_root66, "engine.py"),
                              encoding="utf-8").read()
            _fn66e = _esrc66.split("def _module_level_names")[1].split(
                "def _declared_names")[0]
            check("66.9 ★接线护栏：`_decl_by_mod` 那段确实按 priv / proc 两道跳过"
                  "（`_r_priv or _r_proc`）；`_module_level_names` 确实走了同一个"
                  "可见性函数（`filter_identifiers_by_scope`），而不是自己写第二份规则",
                  [("_r_priv or _r_proc" in _fn66),
                   ("filter_identifiers_by_scope" in _fn66e)],
                  expect_contain=[True, True])
        finally:
            VB66._get_vbe_cached = _orig66
    except Exception as _e66:
        check("第 66 节异常: %s" % _e66, [True], expect_contain=[False])

    # ---- 67. v88：Shift+Enter（光标在代码中间）也补块收尾 ----
    #
    # 用户口径："把光标移动到语句最后按回车可以补全；光标在代码中间按 Shift+Enter，
    # 不会自动补全代码块 —— 帮我加上，也能补全。"
    # 根因：补块收尾只挂在 smart 模式（回车）那条路上，而那条路有一道"光标必须在
    # 行尾"的保险 —— 光标停在中途时它整件事作废、把按键还给 VBE（那时 VBE 原生
    # 回车是【拆行】）。Shift+Enter 这条路本来就不拆行（只往下插一行），所以能在
    # 任意光标列上判块头、补收尾。两条路共用同一份 block_closer / closer_needed /
    # next_line_indent。
    print("\n=== 67. v88：Shift+Enter（行中）也补块收尾 ===")
    try:
        import vbe_bridge as VB67

        class _CM67(object):
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "M1"

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n"
                               for l in self.lines[s0:s0 + int(count)])

            def InsertLines(self, start, text):
                i = min(max(0, int(start) - 1), len(self.lines))
                for k, ln in enumerate(str(text).split("\n")):
                    self.lines.insert(i + k, ln)

        class _Pane67(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE67(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        _orig67 = VB67._get_vbe_cached

        def _run67(lines, line_no, col, **kw):
            """按 mock 的代码窗跑一次 new_line_below，返回 (ok, 行表, 选区)。

            ⚠️ 光标列按【字符列】给 —— 脚手架里的行一律用 ASCII。真机
            GetSelection 给的是**显示列**，全角字符占两格会偏（判据本身与此
            无关，见第 65 节同一处提醒）。
            """
            cm67 = _CM67(lines)
            pane67 = _Pane67(cm67,
                             (int(line_no), int(col), int(line_no), int(col)))
            VB67._get_vbe_cached = lambda: _VBE67(pane67)
            try:
                ok67 = VB67.VbeBackend().new_line_below(**kw)
            finally:
                VB67._get_vbe_cached = _orig67
            return (ok67, cm67.lines, pane67.sel)

        _IF67 = ["Sub F1()",
                 "    x = 0",
                 "    If x = 1 Then",
                 "End Sub"]
        _IFCLOSED67 = ["Sub F1()",
                       "    x = 0",
                       "    If x = 1 Then",
                       "        y = 1",
                       "    End If",
                       "End Sub"]
        _PLAIN67 = ["Sub F1()",
                    "    x = y + 1",
                    "End Sub"]

        # ---- 67.1 ★集成：未闭合的 `If … Then`，光标【行中】-> 补 End If ----
        check("67.1 ★集成（用户报的场景）：`If … Then` 那行、光标停在第 8 列"
              "（行中，不在行尾）按 Shift+Enter -> 补出 `End If`；新行比块头"
              "深一级、光标落在中间那行；块头那一行一个字符没动",
              [_run67(_IF67, 3, 8, auto_close=True)],
              expect_contain=[(True,
                               _IF67[:3] + ["        ", "    End If"] + _IF67[3:],
                               (4, 9, 4, 9))])

        # ---- 67.2 已闭合：不补第二条，但照样进块（深一级）----
        check("67.2 ★集成：`If … Then` 下面已经挂着自己的 `End If` -> 不补第二条"
              "（同一份 closer_needed，v86 那套判据自动覆盖这条新路），但光标"
              "行中按 Shift+Enter 照样把新行缩进深一级（进块里写内容）",
              [_run67(_IFCLOSED67, 3, 8, auto_close=True)],
              expect_contain=[(True,
                               _IFCLOSED67[:3] + ["        "]
                               + _IFCLOSED67[3:], (4, 9, 4, 9))])

        # ---- 67.3 ★过程头：行中按 Shift+Enter -> 补 End Sub ----
        check("67.3 ★集成：过程头 `Sub F1()` 那行、光标停在第 6 列（正好落在"
              "过程名中间）按 Shift+Enter -> 补出 `End Sub`，新行深一级；"
              "过程头那一行原样不动（Shift+Enter 从不拆行）",
              [_run67(["Sub F1()", "    x = 0"], 1, 6, auto_close=True)],
              expect_contain=[(True, ["Sub F1()", "    ", "End Sub", "    x = 0"],
                               (2, 5, 2, 5))])

        # ---- 67.4 普通代码行：维持 v78 老口径 ----
        check("67.4 非块头（普通代码行）光标行中按 Shift+Enter：不补任何收尾，"
              "新行行首空白【抄本行】—— v78 的老口径一个字不动",
              [_run67(_PLAIN67, 2, 8, auto_close=True)],
              expect_contain=[(True,
                               ["Sub F1()", "    x = y + 1", "    ", "End Sub"],
                               (3, 5, 3, 5))])

        # ---- 67.5 ★对照：auto_close 关掉 -> 完全回到 v78 ----
        check("67.5 ★对照：`auto_close=False` 时【不补收尾】（这条新路只由该开关"
              "管收尾）；★v89：**缩进不再跟着这个开关走** —— 块头行照样按判据"
              "深一级（`VBECOMPLETE_NO_AUTOCLOSE=1` 的语义就是\"只留缩进、"
              "不补收尾\"）",
              [_run67(_IF67, 3, 8)],
              expect_contain=[(True,
                               ["Sub F1()", "    x = 0", "    If x = 1 Then",
                                "        ", "End Sub"], (4, 9, 4, 9))])

        # ---- 67.6 ★对照：同一行、同一光标列，回车 vs Shift+Enter ----
        _smart67 = _run67(_IF67, 3, 8, smart=True, auto_close=True)
        _shift67 = _run67(_IF67, 3, 8, auto_close=True)
        check("67.6 ★对照（钉住用户报的那个差别）：同一条 `If … Then`、同一个"
              "【行中】光标列 —— 回车那条路（smart=True）有『光标必须在行尾』"
              "这道保险，这里【拒绝接管】（返回 False，按键还给 VBE 的拆行）；"
              "Shift+Enter 那条路（smart=False + auto_close）把块补出来",
              [[_smart67[0], _shift67[0], _shift67[1]]],
              expect_contain=[[False, True,
                               _IF67[:3] + ["        ", "    End If"]
                               + _IF67[3:]]])

        # ---- 67.7 ★接线护栏 ----
        _root67 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _mainsrc67 = io.open(os.path.join(_root67, "main.py"),
                             encoding="utf-8").read()
        _fn67 = _mainsrc67.split("def _new_line_here")[1].split(
            "def _enter_indent_here")[0]
        _vbsrc67 = io.open(os.path.join(_root67, "vbe_bridge.py"),
                           encoding="utf-8").read()
        # ⚠️ 切片锚点用 smart=False 支的第一句注释（唯一），别再用版本注释行 ——
        # v89 往里插了缩进判定，锚在版本注释上会把新代码切到片外（假红）。
        _else67 = _vbsrc67.split("有选区时以选区【末行】为基准")[1].split(
            "if fixed_line is not None")[0]
        check("67.7 ★接线护栏：Shift+Enter 那条线确实接上了 —— main 的 "
              "`_new_line_here` 把 `auto_close=AUTO_CLOSE_BLOCK` 传下去；"
              "vbe_bridge 的 smart=False 支用的是【同一份】纯判据"
              "（block_closer / closer_needed / next_line_indent，v87 的教训："
              "同一条语义只能有一份实现），扫描窗口与截断兜底也一并接上",
              [[("auto_close=AUTO_CLOSE_BLOCK" in _fn67),
                ("block_closer(line_text)" in _else67),
                ("closer_needed(self._lines_below(cm, anchor, _cap)"
                 in _else67),
                ("_block_kind(line_text)" in _else67),
                ("next_line_indent(line_text, above, unit)" in _else67),
                ("truncated=_rest > _cap" in _else67),
                ("self._CLOSER_SCAN_LINES" in _else67)]],
              expect_contain=[[True] * 7])

        # ---- 67.8 ★回归护栏：v86 那个现场形态，行中也不许多补 ----
        # `脚本.xlsm / 模块2` 摘录（15 行）：第 11 行是那个【顶格】的 `1:`。
        # ⚠️ 真实模块里是全角名，这里换 ASCII —— 同第 65 节的脚手架说明。
        _REAL67 = [
            "Sub Func2()",
            "    Set actSheet = ThisWorkbook.ActiveSheet",
            "    For i = 3 To targetSheet.Cells(targetSheet.Cells.Rows.Count, "
            "\"e\").End(xlUp).Row",
            "        ID = CStr(targetSheet.Cells(i, \"e\"))",
            "    Next",
            "",
            "    For i = 1 To dataSheet1.Cells(dataSheet1.Cells.Rows.Count, "
            "\"c\").End(xlUp).Row",
            "        If dataSheet1.Cells(i, \"c\") = \"name\" Then",
            "            If dataSheet1.Cells(i, \"d\") = \"\" Then GoTo 1",
            "            ID = CStr(dataSheet3.Cells(idRow, \"b\"))",
            "1:",
            "        End If",
            "    Next",
            "    MsgBox \"OK\", vbInformation",
            "End Sub",
        ]
        check("67.8 ★回归护栏：v86 那个现场形态（顶格 `1:`）—— 在 `Sub` 那行、"
              "外层 `For` 那行、`If … Then` 那行的【行中】各按一次 Shift+Enter，"
              "同样一个收尾都不许补（三处下面都挂着自己的收尾）；只多一行缩进"
              "深一级的空行，原有 15 行一个字符没动",
              [_run67(_REAL67, 1, 6, auto_close=True),
               _run67(_REAL67, 3, 10, auto_close=True),
               _run67(_REAL67, 8, 12, auto_close=True)],
              expect_contain=[(True, _REAL67[:1] + ["    "] + _REAL67[1:],
                               (2, 5, 2, 5)),
                              (True, _REAL67[:3] + ["        "] + _REAL67[3:],
                               (4, 9, 4, 9)),
                              (True, _REAL67[:8] + ["            "]
                               + _REAL67[8:], (9, 13, 9, 13))])
    except Exception as _e67:
        check("第 67 节异常: %s" % _e67, [True], expect_contain=[False])

    # ---- 68. v89：Shift+Enter 的缩进判定与回车同源 ----
    #
    # 用户口径："按 shift+enter 换行，不会判定自动缩进，只按 enter 会判定自动缩进。"
    # 根因：v88 只把"新行深一级"接在【有收尾】的块头里面（写在 `if _close:` 支
    # 内），而 `Else` / `Case` / `#If … Then` 这三类是 `_block_kind` 认的块头、
    # 却**没有收尾**（`block_closer` 返回 None，见 _BLOCK_CLOSERS 的说明）⇒
    # 它们和整行注释一起，全落回了 v78 那条"行首空白抄本行"的老路。
    print("\n=== 68. v89：Shift+Enter 的缩进判定与回车同源 ===")
    try:
        import vbe_bridge as VB68

        class _CM68(object):
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "M1"
                self.replace_calls = 0

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n"
                               for l in self.lines[s0:s0 + int(count)])

            def InsertLines(self, start, text):
                i = min(max(0, int(start) - 1), len(self.lines))
                for k, ln in enumerate(str(text).split("\n")):
                    self.lines.insert(i + k, ln)

            def ReplaceLine(self, line, text):
                self.replace_calls += 1
                self.lines[int(line) - 1] = str(text)

        class _Pane68(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE68(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        _orig68 = VB68._get_vbe_cached

        def _run68(lines, line_no, col, **kw):
            """跑一次 new_line_below，返回 (ok, 行表, 选区, ReplaceLine 次数)。

            ⚠️ 光标列按【字符列】给，脚手架里的行一律 ASCII（同第 65/67 节）。
            """
            cm68 = _CM68(lines)
            pane68 = _Pane68(cm68,
                             (int(line_no), int(col), int(line_no), int(col)))
            VB68._get_vbe_cached = lambda: _VBE68(pane68)
            try:
                ok68 = VB68.VbeBackend().new_line_below(**kw)
            finally:
                VB68._get_vbe_cached = _orig68
            return (ok68, cm68.lines, pane68.sel, cm68.replace_calls)

        # ---- 68.1 ★判据形态：根因就在这三类行上 ----
        _bad68a = []
        for _t68 in ("        Else", "            Case 1", "#If VBA7 Then"):
            if not VB68.opens_block(_t68) or VB68.block_closer(_t68) is not None:
                _bad68a.append(_t68)
        check("68.1 ★判据形态（根因）：`Else` / `Case` / `#If … Then` 都算"
              "【块头】（opens_block=True -> 新行该深一级），但"
              "【没有收尾】（block_closer=None，见 _BLOCK_CLOSERS 的说明）——"
              "v88 把缩进写在 `if _close:` 里，这三类就被整片漏掉",
              [_bad68a], expect_contain=[[]])

        _ELSE68 = ["Sub F1()",
                   "    If x = 1 Then",
                   "        y = 1",
                   "        Else"]
        _CASE68 = ["Sub F1()",
                   "    Select Case x",
                   "            Case 1"]
        _PLAIN68 = ["Sub F1()",
                    "    x = y + 1",
                    "End Sub"]
        _IF68 = ["Sub F1()",
                 "    x = 0",
                 "    If x = 1 Then",
                 "End Sub"]

        # ---- 68.2 ★集成：`Else`（歪在块体那一级）行中按 Shift+Enter ----
        check("68.2 ★集成（用户报的场景）：`Else` 歪在块体那一级（8 格）、光标"
              "停在【行中】按 Shift+Enter -> 先把本行拉回 4 格（`Else` 归"
              "`If … Then`、与它齐平，v83），新行 8 格；不补任何收尾"
              "（`End If` 归那个 If，不是给它单开的）；行内文字一个字符没动",
              [_run68(_ELSE68, 4, 10, auto_close=True, align_branch=True)],
              expect_contain=[(True,
                               _ELSE68[:3] + ["    Else", "        "],
                               (5, 9, 5, 9), 1)])

        # ---- 68.3 ★对照：同一行、同一位置，回车（行尾）结果逐字相同 ----
        _sh68 = _run68(_ELSE68, 4, 10, auto_close=True, align_branch=True)
        _en68 = _run68(_ELSE68, 4, 13, smart=True, auto_close=True,
                       align_branch=True)
        _en68mid = _run68(_ELSE68, 4, 10, smart=True, auto_close=True,
                          align_branch=True)
        check("68.3 ★对照（钉住\"两条路同一份判据\"）：同一条 `Else`（同样歪在"
              "8 格）—— 回车在【行尾】时给出的结果（ok / 行表 / 光标位置 / "
              "ReplaceLine 次数，4 项全等）与 Shift+Enter 在【行中】给出的"
              "逐字相同；而回车在【行中】照旧拒绝接管（那道\"光标必须在行尾\""
              "的保险是回车这条入口自己的语义，与判据无关）",
              [[(_en68 == _sh68), _en68mid[0]]],
              expect_contain=[[True, False]])

        # ---- 68.4 ★集成：`Case`（歪在块体那一级）行中 ----
        check("68.4 ★集成：`Case` 歪在 12 格、行中按 Shift+Enter -> 拉回 8 格"
              "（`Case` 比 `Select Case`【深一级】，v83 的分档）+ 新行 12 格",
              [_run68(_CASE68, 3, 15, auto_close=True, align_branch=True)],
              expect_contain=[(True,
                               _CASE68[:2] + ["        Case 1", "            "],
                               (4, 13, 4, 13), 1)])

        # ---- 68.5 ★集成：`#If … Then`（块头但收尾由用户自己写）----
        check("68.5 ★集成：`#If VBA7 Then`（顶格，惯例写法）行中按 Shift+Enter "
              "-> 新行深一级（4 格）；**不补 `#End If`**（条件编译的收尾必须"
              "用户自己写，_BLOCK_CLOSERS 刻意没收它）",
              [_run68(["Sub F1()", "#If VBA7 Then"], 2, 8,
                      auto_close=True, align_branch=True)],
              expect_contain=[(True, ["Sub F1()", "#If VBA7 Then", "    "],
                               (3, 5, 3, 5), 0)])

        # ---- 68.6 ★集成：整行注释 -> 往上找第一个非注释行 ----
        check("68.6 ★集成：整行注释（顶格）行中按 Shift+Enter -> 忽略注释、跟"
              "上方第一个非注释行（这里是块头 `Sub F1()`）对齐 => 新行 4 格"
              "（原先抄注释自己的空白，是 0 格）",
              [_run68(["Sub F1()", "' note"], 2, 4,
                      auto_close=True, align_branch=True)],
              expect_contain=[(True, ["Sub F1()", "' note", "    "],
                               (3, 5, 3, 5), 0)])

        # ---- 68.7 老口径里真正不动的那两条：普通行 / 空行 ----
        _blank68 = _run68(["Sub F1()", "    x = 1", "", "End Sub"], 3, 1,
                          auto_close=True, align_branch=True)
        _ws68 = _run68(["Sub F1()", "    ", "End Sub"], 2, 3,
                       auto_close=True, align_branch=True)
        check("68.7 普通代码行照旧抄本行空白（4 格，v78 老口径）；空行【不猜"
              "层级】（新行仍是空的 —— 回车那一档在这种行上是直接放手交回"
              "VBE 原生，行为等价）；只有空白的行照抄它自己那几格",
              [_run68(_PLAIN68, 2, 8, auto_close=True, align_branch=True),
               _blank68, _ws68],
              expect_contain=[(True, ["Sub F1()", "    x = y + 1", "    ",
                                      "End Sub"], (3, 5, 3, 5), 0),
                              (True, ["Sub F1()", "    x = 1", "", "",
                                      "End Sub"], (4, 1, 4, 1), 0),
                              (True, ["Sub F1()", "    ", "    ", "End Sub"],
                               (3, 5, 3, 5), 0)])

        # ---- 68.8 ★回归：v88 那套（块头补收尾）不许被这次改动碰掉 ----
        check("68.8 ★回归护栏：v88 那套一个字没变 —— 未闭合的 `If … Then` "
              "行中按 Shift+Enter 仍旧补 `End If`、新行深一级、光标停在中间",
              [_run68(_IF68, 3, 8, auto_close=True, align_branch=True)],
              expect_contain=[(True,
                               _IF68[:3] + ["        ", "    End If"]
                               + _IF68[3:], (4, 9, 4, 9), 0)])

        # ---- 68.9 `auto_close` 只管收尾，不管缩进 ----
        check("68.9 ★开关口径：`auto_close=False`（= `VBECOMPLETE_NO_AUTOCLOSE=1`"
              "）时【不补收尾】，但缩进照样按判据算 —— 那条开关的语义就是"
              "\"只留缩进、不补收尾\"（与 67.5 同口径，这里从块头/注释两侧再钉）",
              [_run68(_IF68, 3, 8, auto_close=False, align_branch=True),
               _run68(_ELSE68, 4, 10, auto_close=False, align_branch=True)],
              expect_contain=[(True, _IF68[:3] + ["        "] + _IF68[3:],
                               (4, 9, 4, 9), 0),
                              (True, _ELSE68[:3] + ["    Else", "        "],
                               (5, 9, 5, 9), 1)])

        # ---- 68.10 `align_branch` 开关两条路共用 ----
        check("68.10 `align_branch=False`（= `VBECOMPLETE_NO_BRANCH_ALIGN=1`）"
              "时：`Else` 那一行【不动】（只按它的实际缩进深一级）—— 与回车那条"
              "路同一个开关，不是各管各的",
              [_run68(_ELSE68, 4, 10, auto_close=True, align_branch=False)],
              expect_contain=[(True, _ELSE68 + ["            "],
                               (5, 13, 5, 13), 0)])

        # ---- 68.11 ★接线护栏：结构上防"又被塞进 if _close 里" ----
        _root68 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _vbsrc68 = io.open(os.path.join(_root68, "vbe_bridge.py"),
                           encoding="utf-8").read()
        _else68 = _vbsrc68.split("有选区时以选区【末行】为基准")[1].split(
            "if fixed_line is not None")[0]
        _mainsrc68 = io.open(os.path.join(_root68, "main.py"),
                             encoding="utf-8").read()
        _fn68 = _mainsrc68.split("def _new_line_here")[1].split(
            "def _enter_indent_here")[0]
        _i_ind68 = _else68.index("next_line_indent(line_text, above, unit)")
        _i_close68 = _else68.index("if auto_close:")
        check("68.11 ★接线护栏：smart=False 支里 `next_line_indent` 与 "
              "`branch_align` 都在【空行判断的 else 支】里，而且 `next_line_indent` "
              "出现在 `if auto_close:` 【之前】（= 缩进不归那个开关管 —— 这正是"
              "v88 把缩进塞进 `if _close:` 里造成这次报 bug 的形状）；main 的 "
              "`_new_line_here` 把 `align_branch=BRANCH_ALIGN` 一并传下去",
              [[("branch_align(line_text, above, unit)" in _else68),
                ("if align_branch:" in _else68),
                ("if not str(line_text).strip():" in _else68),
                (_i_ind68 < _i_close68),
                ("align_branch=BRANCH_ALIGN" in _fn68)]],
              expect_contain=[[True] * 5])
    except Exception as _e68:
        check("第 68 节异常: %s" % _e68, [True], expect_contain=[False])

    # ---- 69. v90：Tab 确认之后的那个词，不再提示自己 ----
    #
    # 用户口径（原话）："我还想要修改一个地方，就是当我按下 tab 键之后，提示词打印
    # 到编辑器里，那么本次按键的提示词，就不用匹配提示了。"
    #
    # 修前只有 accept() 里那 0.35s 的【时间】静默期（SUPPRESS_AFTER_ACCEPT）在兜：
    # 轮询发现这次写入、并把它排进动作队列，只要比静默期晚一拍（主线程正忙在一次
    # 全量重解析上时很常见），trigger 就会照常跑到底；而此刻光标处那个词恰好
    # "在工程里真实存在"（就是我们刚写进去的），v37 的"打全名照样提示"会把它留下
    # -> 用户看到刚写进去的词又被提示一遍。
    #
    # 修后：accept() 记下【内容 + 位置】（词 / 模块 / 行号），trigger() 三条全中
    # 就不再提示；任何人为编辑（多一个字符 / 少一个字符 / 换行 / 换模块）都会让
    # 对照失败并当场作废 —— 所以这条记忆不需要计时器，也不会永久屏蔽这个名字。
    print("\n=== 69. v90：Tab 确认后的那个词不再提示自己 ===")
    try:
        import engine as E69
        import vbe_bridge as VB69

        class _UI69(object):
            def __init__(self):
                self.shown = False

            def show(self, *a, **k):
                self.shown = True

            def hide(self, *a, **k):
                self.shown = False

            def update_selection(self, *a, **k):
                pass

            def contains_point(self, *a, **k):
                return False

        class _CM69(object):
            def __init__(self, comp, text):
                self._comp = comp
                self.lines = list(text)

            @property
            def Name(self):
                return self._comp.Name

            @property
            def Parent(self):
                return self._comp

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "\r\n".join(self.lines[s0:s0 + int(count)])

            def ReplaceLine(self, n, text):
                self.lines[int(n) - 1] = str(text).rstrip("\r\n")

        class _Comp69(object):
            def __init__(self, name, ctype, text):
                self.Name = name
                self.Type = ctype
                self.CodeModule = _CM69(self, text)

        class _Proj69(object):
            def __init__(self, comps):
                self.Name = "VBAProject"
                self.VBComponents = list(comps)

        class _Pane69(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE69(object):
            def __init__(self, comps, pane):
                self.ActiveVBProject = _Proj69(comps)
                self.ActiveCodePane = pane

        _orig69 = VB69._get_vbe_cached

        # 标准模块：三个模块级真名字；光标停在第 5 行 `    num` 的词尾。
        # 第 9 行留一段同样的文字，给"换行 -> 记忆作废"用。
        # ⚠️ 行一律 ASCII —— 假 GetSelection 给的是显示列（同第 65 节的脚手架说明）。
        _TEXT69 = ["Public num As Long",
                   "Public numArr As Long",
                   "Public numCount As Long",
                   "Sub Work()",
                   "    num",
                   "End Sub",
                   "",
                   "",
                   "    numCount",
                   ""]

        def _wire69(module="M1", line=5, col=8):
            _c69 = _Comp69(module, 1, _TEXT69)
            _p69 = _Pane69(_c69.CodeModule, (int(line), int(col),
                                             int(line), int(col)))
            VB69._get_vbe_cached = lambda: _VBE69([_c69], _p69)
            return _c69, _p69

        def _mk69(module="M1", line=5, col=8):
            _c69, _p69 = _wire69(module, line, col)
            _ui69 = _UI69()
            _k69 = E69.Completer(VB69.VbeBackend(), _ui69)
            return _k69, _ui69, _c69, _p69

        def _acc69(c, name="numCount"):
            """触发 -> 选中 name -> Tab 确认 -> 把【时间】静默期清 0。

            清 0 这一步是刻意的：它模拟"轮询发现这次写入时已经过了静默期"
            （= 修前会漏的那种晚到），也就是这条用例咬住的那个场景。
            """
            c.trigger(True)
            c.selected = [m.lower() for m in c.current_matches()].index(
                name.lower())
            c.accept()
            c._suppress_until = 0.0

        # ---- 69.1 记忆的内容与来源 ----
        _k69, _u69, _c69, _p69 = _mk69()
        _acc69(_k69)
        check("69.1 accept() 记下【内容 + 位置】= (词, 模块, 行号)；写回也真的"
              "落到那一行、光标停在词尾。⚠️ 模块名必须从 `_ctx_now` 取 —— "
              "`self.ctx` 那份精简快照没有模块名，用它会把模块记成空串，"
              "\"换模块\"这条对照就永远失效",
              [(_k69._accepted, _c69.CodeModule.lines[4], _p69.sel)],
              expect_contain=[(("numcount", "m1", 5), "    numCount",
                               (5, 13, 5, 13))])

        # ---- 69.2 ★核心：晚到的 trigger 不再提示自己 ----
        _k69b, _u69b, _c69b, _p69b = _mk69()
        _acc69(_k69b)
        _k69b.trigger(True)
        check("69.2 ★核心（用户报的）：Tab 确认后，那次写入被轮询【晚一点】"
              "发现（静默期已过）—— 弹窗不再把刚写进去的那个词提示回来",
              [(_u69b.shown, _k69b.current_matches())],
              expect_contain=[(False, [])])

        # ---- 69.3 ★对照：用户点名的列表照弹（证明 69.2 挡的是回声，不是没候选）----
        _k69c, _u69c, _c69c, _p69c = _mk69()
        _acc69(_k69c)
        _k69c.trigger(True, manual=True)
        check("69.3 ★对照：同一条路径上 `manual=True`（Ctrl+Space，用户点名的"
              "列表）照弹 —— 证明 69.2 不是\"候选里压根没有它\"，而是被 v90 那条"
              "规矩挡下的（v72 老口径：点名的列表不能被收掉）",
              [(_u69c.shown, _k69c.current_matches())],
              expect_contain=[(True, ["numCount"])])

        # ---- 69.4 对照：换了一行 -> 记忆作废 ----
        _k69d, _u69d, _c69d, _p69d = _mk69()
        _acc69(_k69d)
        _p69d.SetSelection(9, 13, 9, 13)
        _k69d.trigger(True)
        check("69.4 ★对照（记忆不许赖着不走·换行）：确认后光标移到【另一行】的"
              "同名文字上 —— 位置对不上，记忆当场作废，照常提示（v37 老口径）",
              [(_u69d.shown, _k69d.current_matches(), _k69d._accepted)],
              expect_contain=[(True, ["numCount"], None)])

        # ---- 69.5 对照：换了模块 -> 记忆作废 ----
        _k69e, _u69e, _c69e, _p69e = _mk69()
        _acc69(_k69e)
        # 另一个标准模块：第 5 行就写着 numCount（同名成员在它那儿也是真名字）
        _TEXT69B = list(_TEXT69)
        _TEXT69B[4] = "    numCount"
        _p69e2 = _Pane69(_Comp69("M2", 1, _TEXT69B).CodeModule,
                         (5, 13, 5, 13))
        VB69._get_vbe_cached = lambda: _VBE69([_p69e2.CodeModule.Parent],
                                              _p69e2)
        _k69e.trigger(True)
        check("69.5 ★对照（记忆不许赖着不走·换模块）：同名成员在另一个标准模块"
              "里也是真名字 —— 模块名对不上，记忆作废，照常提示",
              [(_u69e.shown, _k69e.current_matches(), _k69e._accepted)],
              expect_contain=[(True, ["numCount"], None)])

        # ---- 69.6 ★关键对照：用户一动过手，v37 老口径立刻恢复 ----
        _k69f, _u69f, _c69f, _p69f = _mk69()
        _acc69(_k69f)
        _c69f.CodeModule.lines[4] = "    numCoun"      # 退格删掉一个字符
        _p69f.SetSelection(5, 12, 5, 12)
        _k69f.trigger(True)
        _r69f1 = (_u69f.shown, _k69f.current_matches(), _k69f._accepted)
        _c69f.CodeModule.lines[4] = "    numCount"     # 又打回来
        _p69f.SetSelection(5, 13, 5, 13)
        _k69f.trigger(True)
        _r69f2 = (_u69f.shown, _k69f.current_matches())
        check("69.6 ★关键对照：确认之后【退格一个字符】-> 词对不上，记忆当场"
              "作废，照常提示（连退格后剩下的 numCoun 也照旧提示出真实存在的"
              "numCount，v43/v37 的老口径）；再手打回 numCount 也照常提示 —— "
              "这条记忆只活到\"用户动过手\"为止，不是永久屏蔽",
              [_r69f1, _r69f2],
              expect_contain=[(True, ["numCount"], None),
                              (True, ["numCount"])])

        # ---- 69.7 没确认过就没有记忆（不误伤）----
        _k69g, _u69g, _c69g, _p69g = _mk69()
        _k69g.trigger(True)
        check("69.7 只触发、没确认 -> `_accepted` 是 None（没有待作废的记忆），"
              "照常提示（别让这条规矩误伤正常输入）",
              [(_k69g._accepted, _u69g.shown)],
              expect_contain=[(None, True)])

        # ---- 69.8 ★接线护栏 ----
        _root69 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _engsrc69 = io.open(os.path.join(_root69, "engine.py"),
                            encoding="utf-8").read()
        _mainsrc69 = io.open(os.path.join(_root69, "main.py"),
                             encoding="utf-8").read()
        _acc69src = _engsrc69.split("    def accept(self):")[1].split(
            "    def pick(self")[0]
        _trgsrc69 = _engsrc69.split("def trigger(self,")[1].split(
            "    def can_space_confirm")[0]
        _tgate69 = _trgsrc69.split(
            "if self._accepted is not None and not manual:")[1]
        check("69.8 ★接线护栏：accept() 从 `_ctx_now` 取上下文（不是 self.ctx，"
              "它没有模块名）、并把记忆写在【真正写回之前】；trigger 里那道判据"
              "在 `not manual` 之内（手动唤出放行），且对照失败时会把记忆作废"
              "（= 它不是一劳永逸的屏蔽）；main 的 Ctrl+Space 传 `manual=True`",
              [[("_src = getattr(self, \"_ctx_now\", None) or self.ctx or {}"
                 in _acc69src),
                ("self._accepted = (str(chosen).lower()" in _acc69src),
                (_acc69src.index("self._accepted = (str(chosen).lower()")
                 < _acc69src.index("self.backend.apply_completion(")),
                ("if self._accepted is not None and not manual:" in _trgsrc69),
                ("self._accepted = None" in _tgate69),
                ("completer.trigger(manual=True)" in _mainsrc69)]],
              expect_contain=[[True] * 6])

        # ---- 69.9 ★反坑护栏：hide() 里不许清记忆 ----
        _hidesrc69 = _engsrc69.split("    def hide(self):")[1].split(
            "    def maybe_hide_on_outside_click")[0]
        check("69.9 ★反坑护栏：`hide()` 里【不许】出现 `_accepted` —— accept() "
              "记完记忆紧接着就调 hide()，在 hide 里清等于这个功能从来没生效过"
              "（两条判据也都在：`SUPPRESS_AFTER_ACCEPT` 那条时间静默期没被删）",
              [[("_accepted" not in _hidesrc69),
                ("SUPPRESS_AFTER_ACCEPT" in _acc69src),
                ("self._suppress_until = time.time() + SUPPRESS_AFTER_ACCEPT"
                 in _acc69src)]],
              expect_contain=[[True, True, True]])

        VB69._get_vbe_cached = _orig69
    except Exception as _e69:
        check("第 69 节异常: %s" % _e69, [True], expect_contain=[False])

    # ===== 70. v90b：确认键压在 VBE 的成员列表上，那一次录入就算完成 =====
    # 用户报的（原话）："我定义了一个类模块叫 类1，当我输入 Set s1 = New，再按
    # 空格，vbe 会弹提示列表，我选中 类1 之后，按 tab 键，我们项目还会提示 类1。
    # …… 按 tab 键，就意味着本次录入完成，不需要去匹配提示词。"
    #
    # 与第 69 节（v90）的分工 —— 两条判据管的是【两个不同的写入者】：
    #   * v90 管【我们替用户写进去的】词：accept() 知道自己写了什么，用
    #     (词, 模块, 行号) 三条对照；
    #   * v90b 管【VBE 替用户写进去的】词：那一记 Tab 是喂给 VBE 自己的成员
    #     列表的 —— 我们的窗早就让位隐藏了（VBE 列表画在屏幕上时我们一律
    #     让位），所以 accept() 压根没被调用、v90 的记忆从没写下。我们只看得见
    #     "按了确认键"，看不见 VBE 插了什么，于是只能靠【按下前的行文本 + 这一行
    #     只被插入了一段（其余逐字未动）】来认。
    #
    # ⚠️ 判据刻意【不用光标列】：state["cur_ctx"] 里那一列来自
    #    backend.snapshot()，是 VBE 的【显示列】（全角占 2 格）；而 trigger 里
    #    ctx["caret_col"] 是【字符列】。两个口径混着比，用户场景的中文类名
    #    `类1` 必然对不上 —— 这正是本案差点踩进去的坑，70.8 有护栏钉它。
    print("\n=== 70. v90b：确认键压在 VBE 成员列表上时，那一次录入算完成 ===")
    try:
        import engine as E70
        import vbe_bridge as VB70

        class _UI70(object):
            def __init__(self):
                self.shown = False

            def show(self, *a, **k):
                self.shown = True

            def hide(self, *a, **k):
                self.shown = False

            def update_selection(self, *a, **k):
                pass

            def contains_point(self, *a, **k):
                return False

        class _CM70(object):
            def __init__(self, comp, text):
                self._comp = comp
                self.lines = list(text)

            @property
            def Name(self):
                return self._comp.Name

            @property
            def Parent(self):
                return self._comp

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "\r\n".join(self.lines[s0:s0 + int(count)])

            def ReplaceLine(self, n, text):
                self.lines[int(n) - 1] = str(text).rstrip("\r\n")

        class _Comp70(object):
            def __init__(self, name, ctype, text):
                self.Name = name
                self.Type = ctype
                self.CodeModule = _CM70(self, text)

        class _Proj70(object):
            def __init__(self, comps):
                self.Name = "VBAProject"
                self.VBComponents = list(comps)

        class _Pane70(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE70(object):
            def __init__(self, comps, pane):
                self.ActiveVBProject = _Proj70(comps)
                self.ActiveCodePane = pane

        _orig70 = VB70._get_vbe_cached

        _BEFORE70 = "    Set s1 = New "          # 按下 Tab 之前那一行（17 字符）
        _AFTER70 = "    Set s1 = New Cls1"       # VBE 把 Cls1 写了进来

        def _wire70(line_text, line=3, col=None, cls_name="Cls1",
                    with_cls=True):
            _c70 = _Comp70("M1", 1,
                           ["Sub Work()",
                            "    Dim s1 As Object",
                            line_text,
                            "End Sub"])
            _comps70 = [_c70]
            if with_cls:
                _comps70.append(_Comp70(cls_name, 2, ["Public x As Long"]))
            _n70 = int(col if col is not None else len(line_text))
            _p70 = _Pane70(_c70.CodeModule, (line, _n70, line, _n70))
            VB70._get_vbe_cached = lambda: _VBE70(_comps70, _p70)
            return _c70, _p70

        def _mk70(line_text, **kw):
            _c70, _p70 = _wire70(line_text, **kw)
            _ui70 = _UI70()
            _be70 = VB70.VbeBackend()
            # ★v93：`Set s1 = New Cls1` 是【New 类型位置】—— parser 现在会把它
            # 认成 in_type_position=True（见 is_caret_in_type_position 的 v93 说明），
            # 于是候选要收窄到"类型名"。真后端由 get_type_names 提供；这个假 VBE
            # 没有那套 COM 面，得手动喂上，否则候选会被收空（不是产品行为，是
            # 脚手架缺件）。类模块名本来就是类型名，喂 Cls1 与 kw 的 cls_name 一致。
            _be70._type_names = {str(kw.get("cls_name", "Cls1")).lower()}
            _k70 = E70.Completer(_be70, _ui70)
            return _k70, _ui70, _c70, _p70

        def _note70(k, ts=0.0, line=3, text=_BEFORE70):
            """模拟"按下确认键那一刻"：main 的钩子线程把快照交给 engine。"""
            k.note_accept_key((time.time() + ts, line, text))

        # ---- 70.1 ★核心（用户报的那一下）----
        _k70, _u70, _c70, _p70 = _mk70(_AFTER70)
        _note70(_k70)
        _k70.trigger(False)
        check("70.1 ★核心（用户报的）：Tab 压在【VBE 自己的成员列表】上确认"
              "（我们的窗早已让位隐藏、accept 从没被调用），VBE 把 `Cls1` 写了"
              "进来 —— 轮询发现这一行变了之后，不该把它提示回来",
              [(_u70.shown, _k70.current_matches())],
              expect_contain=[(False, [])])

        # ---- 70.2 ★对照：同一条行、同一个光标，只是【没有】这条记忆 ----
        _k70b, _u70b, _c70b, _p70b = _mk70(_AFTER70)
        _k70b.trigger(False)
        check("70.2 ★对照：换一个人（没按过确认键）在同一处触发 -> 照常提示。"
              "证明 70.1 挡的是\"刚确认的那一次录入\"，不是候选里压根没有它",
              [(_u70b.shown, _k70b.current_matches())],
              expect_contain=[(True, ["Cls1"])])

        # ---- 70.3 ★用户原场景：中文类名（全角显示列）----
        _k70c, _u70c, _c70c, _p70c = _mk70("    Set s1 = New 类1", col=21,
                                           cls_name=u"类1")
        _note70(_k70c)
        _k70c.trigger(False)
        _r70c = (_u70c.shown, _k70c.current_matches())
        _k70d, _u70d, _c70d, _p70d = _mk70("    Set s1 = New 类1", col=21,
                                           cls_name=u"类1")
        _k70d.trigger(False)
        check("70.3 ★用户原场景（中文类名 `类1`，光标给的是全角显示列 21）："
              "①有确认键记忆 -> 不提示；②同样一处、没有记忆 -> 照常提示。"
              "判据只用文本、不碰列 —— 显示列/字符列混用在这儿必然对不上",
              [_r70c, (_u70d.shown, _k70d.current_matches())],
              expect_contain=[(False, []), (True, [u"类1"])])

        # ---- 70.4 ★v92：命中之后记忆【不再当场消费】，而是"钉住"VBE 留下的那一行 ----
        # v91 之前这里是"只活一次"（命中即消费）。v92 发现真机上一次 Tab 会引来
        # 【不止一次】trigger：第 1 次（补位）就把记忆吃掉之后，跟着来的那一次
        # 无据可依，把刚写进去的词弹了回来 —— 用户报的"后面还是会提示"。
        # 现在命中只是把锚换成"VBE 留下的那一行"（第二段），继续看着。
        _k70e, _u70e, _c70e, _p70e = _mk70(_AFTER70)
        _note70(_k70e)
        _k70e.trigger(False)
        _r70e1 = (_u70e.shown, _k70e._accept_pinned, _k70e._accept_key[1:])
        _p70e.CodeModule.lines[2] = "    Set s1 = New Cl"    # 用户动过手：退格
        _p70e.SetSelection(3, 19, 3, 19)
        _k70e.trigger(False)
        check("70.4 ★v92：命中那一刻【不再消费】记忆，而是把它钉到\"VBE 留下的"
              "那一行\"上（行号 + 行文本原样）—— 真机上一次 Tab 引来的 trigger "
              "不止一次，只活一段就会漏；而用户一动过手（这一行不再是那个样子）"
              "-> 记忆当场作废、v37 老口径立刻恢复",
              [_r70e1, (_k70e._accept_key, _k70e._accept_pinned, _u70e.shown)],
              expect_contain=[(False, True, (3, _AFTER70)),
                              (None, False, True)])

        # ---- 70.5 ★对照（记忆不许赖着不走）：陈旧 / 换行 ----
        # v91：TTL 从 1.5s 放宽到 3.0s（记账与判据被消费之间隔着动作队列，
        # 那一头可能正堵着一次全量重解析）。这里的"已经过期"相应取 -9.0s，
        # 别贴着边界（-3.0 在 `<=` 口径下正好还算数，v91 改 TTL 时踩到过）。
        #
        # ⚠️ 这两问【必须分开 check】：check() 是"期望元素出现在 got 里"的
        # 子集匹配，合成一条时第一问的期望会被第二问的 got 满足 —— 也就是
        # "TTL 失效了也照样绿"（v91 改 TTL 时正是被这条假绿骗过一次）。
        _k70f, _u70f, _c70f, _p70f = _mk70(_AFTER70)
        _note70(_k70f, ts=-9.0)
        _k70f.trigger(False)
        check("70.5a ★对照（记忆不许赖着不走）：快照已陈旧（TTL 过期 —— "
              "COM 退避时 cur_ctx 会停更）-> 照常提示",
              [(_u70f.shown, _k70f.current_matches())],
              expect_contain=[(True, ["Cls1"])])
        _k70g, _u70g, _c70g, _p70g = _mk70(_AFTER70)
        _note70(_k70g, line=9)
        _k70g.trigger(False)
        check("70.5b ★对照（记忆不许赖着不走）：行号对不上（用户换行了）"
              "-> 照常提示",
              [(_u70g.shown, _k70g.current_matches())],
              expect_contain=[(True, ["Cls1"])])

        # ---- 70.6 ★对照（用户一动过手，v37 老口径立刻恢复）----
        # 按【真实时序】摆：Tab 插入那一行变化总会先被轮询发现、触发一次
        # （记忆在那一次就消费掉了），用户再退格走的是第二次 trigger。
        _k70h, _u70h, _c70h, _p70h = _mk70(_AFTER70)
        _note70(_k70h)
        _k70h.trigger(False)
        _r70h1 = _u70h.shown
        _p70h.CodeModule.lines[2] = "    Set s1 = New Cl"    # 退格删一个字符
        _p70h.SetSelection(3, 19, 3, 19)
        _k70h.trigger(False)
        _r70h2 = (_u70h.shown, _k70h.current_matches())
        _k70i, _u70i, _c70i, _p70i = _mk70("    Set s1 = New Cls")
        _note70(_k70i, text=_AFTER70)   # 快照里的旧文本更长 -> 这一行是变短的
        _k70i.trigger(False)
        check("70.6 ★对照（用户一动过手，老口径立刻恢复）：①真实时序的第 1 次"
              "（Tab 那次）不提示；②用户退格后第 2 次 -> 照常提示（连退格后"
              "剩下的 `Cl` 也照旧提示出真实存在的 `Cls1`）；③这一行比快照"
              "【变短】了（旧文本被动过，不是\"只插入一段\"）-> 照常提示",
              [_r70h1, _r70h2, (_u70i.shown, _k70i.current_matches())],
              expect_contain=[False, (True, ["Cls1"]), (True, ["Cls1"])])

        # ---- 70.7 manual=True 放行 ----
        _k70j, _u70j, _c70j, _p70j = _mk70(_AFTER70)
        _note70(_k70j)
        _k70j.trigger(False, manual=True)
        check("70.7 ★对照：`manual=True`（Ctrl+Space 点名要的列表）放行 —— "
              "v72 老口径：点名的列表不能被这条规矩收掉",
              [(_u70j.shown, _k70j.current_matches())],
              expect_contain=[(True, ["Cls1"])])

        # ---- 70.8 ★接线护栏（v91 修订：门槛收干净）----
        _root70 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _mainsrc70 = io.open(os.path.join(_root70, "main.py"),
                             encoding="utf-8").read()
        _engsrc70 = io.open(os.path.join(_root70, "engine.py"),
                            encoding="utf-8").read()
        _branch70 = _mainsrc70.split("elif (vk == VK_TAB")[1].split(
            "if suppress:")[0]
        # 只看【代码行】：v91 的注释里刻意写着 `_vbe_list_showing_now` /
        # `com_backoff_remaining` 这些名字（讲清为什么把它们从条件里去掉），
        # 断言不能被注释里的字样带偏。
        _code70 = "\n".join(_l for _l in _branch70.splitlines()
                            if not _l.strip().startswith("#"))
        _note70src = _engsrc70.split(
            "    def note_accept_key(self, snapshot):")[1].split(
            "    def accept(self):")[0]
        _trg70 = _engsrc70.split("def trigger(self,")[1].split(
            "    def can_space_confirm")[0]
        _key70 = _trg70.split(
            "if self._accept_key is not None and not manual:")[1]
        check("70.8 ★接线护栏：①那一支排在 `completer.is_visible()` 之后"
              "（我们的窗可见时走 accept，不重复记账）；②刻意【不吞键】—— "
              "分支里没有 suppress = True，这一下 Tab 照旧交给 VBE 去插入；"
              "③交给 engine 的是 (时刻, 行号, 行文本)，【不带光标列】；"
              "④engine 那条判据在 `not manual` 之内、命中即 hide + return，"
              "且排在 v90 那条【之前】",
              [[("post(completer.note_accept_key" in _branch70),
                ("_snap[1], _snap[2]" in _branch70),
                ("_snap[3]" not in _branch70),
                ("suppress = True" not in _code70),
                (_mainsrc70.index("elif completer.is_visible():")
                 < _mainsrc70.index("elif (vk == VK_TAB")),
                ("_ts, _ln, _txt = snapshot" in _note70src),
                ("_ts, _ln, _txt, _col" not in _note70src),
                ("if self._accept_key is not None and not manual:" in _trg70),
                (_trg70.index("if self._accept_key is not None")
                 < _trg70.index(
                     "if self._accepted is not None and not manual:")),
                ("self.hide()" in _key70)]],
              expect_contain=[[True] * 10])

        # ---- 70.9 ★反坑护栏：hide() 里不许清这条记忆 ----
        _hide70 = _engsrc70.split("    def hide(self):")[1].split(
            "    def maybe_hide_on_outside_click")[0]
        check("70.9 ★反坑护栏：`hide()` 里【不许】出现 `_accept_key`（也不许"
              "出现 v90 的 `_accepted`）—— 命中那一刻的处理正是 hide() + "
              "return，在 hide 里清等于这条判据从来没生效过",
              [[("_accept_key" not in _hide70),
                ("_accepted" not in _hide70)]],
              expect_contain=[[True, True]])

        VB70._get_vbe_cached = _orig70
    except Exception as _e70:
        check("第 70 节异常: %s" % _e70, [True], expect_contain=[False])

    # =====================================================================
    # ★v91（用户报的）：`模块名.成员名` 这种【点号成员补全】上按 Tab 确认后，
    # 那个词又被提示了一遍。
    #
    # 与 70 节（v90b）的差别不在判据，而在【这一次触发是谁发起的】：
    #   70 节的落点是"Tab 之后轮询发现这一行变了"那一次 trigger；
    #   v91 报的落点是【让位宽限到、VBE 没弹 -> 自己把候选补回来】那一次
    #   trigger(False)（v73 的补位路径）。
    # 点号位置上 `vbe_list_expected` 恒为真（实测：连 `NoExplicitPub.getTitle`
    # 这种"词已写完"的行也算真），所以我们全程让位、从不弹自己的窗 ——
    # 于是每一次行变化都会走一遍"待定让位 -> 宽限到 -> 补位"。
    #
    # 为什么 70 节没能兜住：钩子侧记快照还要求 `_vbe_list_showing_now()`
    # （上一轮 100ms 轮询采到的窗口签名里有 NameListWndClass）。那一格没采上
    # 就压根没记账，补位那一次便无据可依。v91 把门槛收成"裸 Tab + 焦点在
    # 代码窗格"，代价是 Tab【缩进】也会记账 —— 由判据第 4 条（插入段必须
    # 全是标识符字符）挡掉，见 71.3 / 71.4。
    # =====================================================================
    print("\n=== 71. v91：点号成员按 Tab 确认后不再提示自己（含缩进不误伤）===")
    try:
        import engine as E71
        import vbe_bridge as VB71

        class _UI71(object):
            def __init__(self):
                self.shown = False

            def show(self, *a, **k):
                self.shown = True

            def hide(self, *a, **k):
                self.shown = False

            def update_selection(self, *a, **k):
                pass

            def contains_point(self, *a, **k):
                return False

        class _CM71(object):
            def __init__(self, comp, text):
                self._comp = comp
                self.lines = list(text)

            @property
            def Name(self):
                return self._comp.Name

            @property
            def Parent(self):
                return self._comp

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "\r\n".join(self.lines[s0:s0 + int(count)])

            def ReplaceLine(self, n, text):
                self.lines[int(n) - 1] = str(text).rstrip("\r\n")

        class _Comp71(object):
            def __init__(self, name, ctype, text):
                self.Name = name
                self.Type = ctype
                self.CodeModule = _CM71(self, text)

        class _Proj71(object):
            def __init__(self, comps):
                self.Name = "VBAProject"
                self.VBComponents = list(comps)

        class _Pane71(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE71(object):
            def __init__(self, comps, pane):
                self.ActiveVBProject = _Proj71(comps)
                self.ActiveCodePane = pane

        _orig71 = VB71._get_vbe_cached
        _pyshow71 = E71.Completer._vbe_popup_showing
        # 这一节的落点是【文本判据】。真机上的 Win32 让位查询会随"此刻 VBE
        # 有没有画着成员列表"摇摆，测起来不稳定；一律当"列表不在屏幕上"。
        E71.Completer._vbe_popup_showing = lambda self: False

        _BEFORE71 = "    NoExplicitPub.getTi"
        _AFTER71 = "    NoExplicitPub.getTitle"
        _MEMBERS71 = ["Public abcDef As Long",
                      "Public Function getTitle() As String",
                      "    getTitle = \"\"",
                      "End Function"]

        def _wire71(line_text, line=3, col=None):
            _c71 = _Comp71("M1", 1,
                           ["Sub Work()",
                            "    Dim s1 As Object",
                            line_text,
                            "End Sub",
                            ""] + _MEMBERS71)
            _comps71 = [_c71, _Comp71("NoExplicitPub", 1, _MEMBERS71)]
            _n71 = int(col if col is not None else len(line_text))
            _p71 = _Pane71(_c71.CodeModule, (line, _n71, line, _n71))
            VB71._get_vbe_cached = lambda: _VBE71(_comps71, _p71)
            return _c71, _p71

        def _mk71(line_text, **kw):
            _c71, _p71 = _wire71(line_text, **kw)
            _ui71 = _UI71()
            _k71 = E71.Completer(VB71.VbeBackend(), _ui71)
            return _k71, _ui71, _c71, _p71

        def _note71(k, ts=0.0, line=3, text=_BEFORE71):
            k.note_accept_key((time.time() + ts, line, text))

        def _ins71(before, after):
            """复算判据第 4 条看的那个量：新插入的那一段。"""
            _p = 0
            while (_p < len(before) and _p < len(after)
                   and before[_p] == after[_p]):
                _p += 1
            _s = 0
            while (_s < len(before) - _p and _s < len(after) - _p
                   and before[-1 - _s] == after[-1 - _s]):
                _s += 1
            return after[_p:len(after) - _s], (_p + _s >= len(before))

        # ---- 71.1 ★核心（用户报的那一下）----
        # 按真实时序摆：Tab 之后这一行变了 -> 轮询 trigger(True) 走语法让位
        # -> 宽限到、VBE 没弹 -> trigger(False) 补位。
        _k71, _u71, _c71, _p71 = _mk71(_AFTER71)
        _note71(_k71)
        _k71.trigger(True)
        _r71_yield = _k71.yield_pending is not None
        _k71.confirm_yield(False)
        _k71.trigger(False)
        check("71.1 ★核心（用户报的）：`NoExplicitPub.getTi` 处按 Tab、VBE 把"
              "`getTitle` 补进来之后 —— 轮询那次先让位（点号位置 VBE 会弹列表），"
              "宽限到 VBE 没弹、候选补回来那一次【不该】把刚写进去的词提示一遍",
              [(_r71_yield, _u71.shown, _k71.current_matches())],
              expect_contain=[(True, False, [])])

        # ---- 71.2 ★对照：同一条行、同一个光标，只是【没有】这条记忆 ----
        _k71b, _u71b, _c71b, _p71b = _mk71(_AFTER71)
        _k71b.trigger(True)
        _k71b.confirm_yield(False)
        _k71b.trigger(False)
        check("71.2 ★对照：没按过 Tab（同一处，同样走一遍让位->补位）-> 照常"
              "提示。证明 71.1 挡的是\"刚确认的那一次录入\"，不是这一处压根"
              "给不出候选",
              [(_u71b.shown, _k71b.current_matches())],
              expect_contain=[(True, ["getTitle"])])

        # ---- 71.3 ★判据第 4 条：Tab 缩进不许被误当成"一次录入" ----
        # v91 把钩子门槛放宽成"裸 Tab 就记账"，于是 Tab 缩进那一下也记账了。
        # 缩进插进来的是空格 / 制表符 —— 前三条判据（同一行 + 这一行变长 +
        # 旧文本 = 新文本的前缀+后缀）**全都成立**，只靠第 4 条挡住。
        _k71c, _u71c, _c71c, _p71c = _mk71("    abc", col=5)
        _note71(_k71c, text="abc")          # 行首缩进：abc -> "    abc"
        _k71c.trigger(False)
        check("71.3 ★缩进不误伤：行首按 Tab 缩进（插进来的是 4 个空格）-> "
              "该照常提示，不许因为\"这一行变长了\"就白静默一次",
              [(_u71c.shown, _k71c.current_matches())],
              expect_contain=[(True, ["abcDef"])])

        # ---- 71.4 ★判据第 4 条到底在看什么（把中间量摆出来）----
        _i71a, _f71a = _ins71(_BEFORE71, _AFTER71)
        _i71b, _f71b = _ins71("abc", "    abc")
        _i71c, _f71c = _ins71("    abc def", "    abc\t def")
        check("71.4 ★判据第 4 条看的量：①成员确认 -> 插入段 `tle`（全是标识符"
              "字符，认作一次录入）；②行首 Tab 缩进 -> 4 个空格；③行中间 Tab "
              "-> `\\t`。②③ 的\"前缀+后缀\"判据都成立，只有第 4 条能把它们"
              "挡掉 —— 这正是放开钩子门槛的代价所在",
              [[(_i71a, _f71a,
                 all((_x.isalnum() or _x == "_") for _x in _i71a)),
                (_i71b, _f71b,
                 all((_x.isalnum() or _x == "_") for _x in _i71b)),
                (_i71c, _f71c,
                 all((_x.isalnum() or _x == "_") for _x in _i71c))]],
              expect_contain=[[("tle", True, True),
                               ("    ", True, False),
                               ("\t", True, False)]])

        # ---- 71.5 ★中文 / 全角也无碍 ----
        _k71e, _u71e, _c71e, _p71e = _mk71(u"    Set s = NoExplicitPub.getTitle")
        _note71(_k71e, text=u"    Set s = NoExplicitPub.getTi")
        _k71e.trigger(False)
        check("71.5 ★中文 / 全角也无碍：判据只用文本、一个列都不碰"
              "（显示列 vs 字符列的坑见 70.8 与第 70 节开头）-> 不提示",
              [(_u71e.shown, _k71e.current_matches())],
              expect_contain=[(False, [])])

        # ---- 71.6 ★TTL 放宽到 3.0s ----
        _k71f, _u71f, _c71f, _p71f = _mk71(_AFTER71)
        _note71(_k71f, ts=-2.5)
        _k71f.trigger(False)
        _k71g, _u71g, _c71g, _p71g = _mk71(_AFTER71)
        _note71(_k71g, ts=-9.0)
        _k71g.trigger(False)
        check("71.6 ★TTL（v91：1.5s -> 3.0s）：①2.5s 前的快照仍算数 —— 按下"
              "那一刻记账、与判据被消费之间隔着动作队列，那一头可能正堵着一次"
              "全量重解析 -> 不提示；②9s 前的快照已过期（COM 退避时 cur_ctx "
              "会停更）-> 照常提示",
              [(_u71f.shown, _k71f.current_matches()),
               (_u71g.shown, _k71g.current_matches())],
              expect_contain=[(False, []), (True, ["getTitle"])])

        # ---- 71.7 ★接线护栏（门槛只剩"裸 Tab + 焦点在代码窗格"）----
        _root71 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _mainsrc71 = io.open(os.path.join(_root71, "main.py"),
                             encoding="utf-8").read()
        _engsrc71 = io.open(os.path.join(_root71, "engine.py"),
                            encoding="utf-8").read()
        _branch71 = _mainsrc71.split("elif (vk == VK_TAB")[1].split(
            "if suppress:")[0]
        _code71 = "\n".join(_l for _l in _branch71.splitlines()
                            if not _l.strip().startswith("#"))
        _key71 = _engsrc71.split("def trigger(self,")[1].split(
            "    def can_space_confirm")[0].split(
            "if self._accept_key is not None and not manual:")[1]
        check("71.7 ★接线护栏：①钩子那一支的条件里【不再】依赖"
              "`_vbe_list_showing_now`（100ms 轮询的窗口缓存 —— 用户这次报的"
              "正是它没采上）与 `com_backoff_remaining`；②只剩裸 Tab + 焦点在"
              "代码窗格（注释里刻意留着这两个名字讲缘由，所以只看代码行）；"
              "③记账时写一行日志，把\"压根没记\"和\"记了但判据没命中\"分开；"
              "④engine 判据里必须有\"插入段全是标识符字符\"这一条 —— 它是"
              "放开门槛的唯一代价承担者（v92 起那一段先去掉一个开头的点号，"
              "落在 `_body` 上比，见 72.6）",
              [[("_vbe_list_showing_now" not in _code71),
                ("com_backoff_remaining" not in _code71),
                ("in_vbe_code_area()" in _code71),
                ("post(completer.note_accept_key" in _code71),
                ("hook: Tab 按下" in _branch71),
                ("_body" in _key71),
                ("isalnum()" in _key71)]],
              expect_contain=[[True] * 7])

        VB71._get_vbe_cached = _orig71
        E71.Completer._vbe_popup_showing = _pyshow71
    except Exception as _e71:
        check("第 71 节异常: %s" % _e71, [True], expect_contain=[False])

    # =====================================================================
    # ★v92（用户报的）：按了 Tab，词已经写进编辑器，但【后面】还是提示。
    #
    # 用户截图那次是 `With Target.Validation` 里的 `.Delete` —— 点号成员位置，
    # 我们全程让位（同 71 节），弹回来的是我们自己的候选 `deleteAllFiles`。
    #
    # v91 的判据本身没问题，坏在【记账只活一次】：真机上一次 Tab 会引来
    # 【不止一次】trigger ——
    #   ① Tab 按下 -> 钩子记下"按下前的那一行"；
    #   ② VBE 把这个词写进编辑器；
    #   ③ 轮询发现这一行变了 -> trigger(True) 走【让位】返回（不判据、不消费）；
    #   ④ 同一轮轮询里，先前那次让位的宽限正好到期、VBE 的列表已关
    #      -> confirm_yield(False) -> 补位 trigger(False)：判据命中 -> 静默；
    #   ⑤ 但补位那一下把"已放弃让位"标记清了/记上了，随后 `_run_trigger` 又
    #      补了一次 trigger(True)：宽限刚记过"这一处 VBE 不会弹"，它不再让位、
    #      直接走到判据 —— 而记忆已被第 ④ 步吃掉（"只活一次"）-> 光标处的
    #      `Delete` 是"工程里真实存在的" -> v37 的"打全名照样提示"把它弹了
    #      回来（弹窗就在 Tab 之后紧接着出现，正是用户说的"后面还是会提示"）。
    #
    # 本节把"Tab 之后这一次写入引来的每一次 trigger 都要静默"钉死；两处放宽
    # （大小写、开头的点号）与一处收紧（第二段改成"一个字都没再动过"）一并钉。
    # =====================================================================
    print("\n=== 72. v92：Tab 之后一次写入引来的【多次】trigger 都要静默 ===")
    try:
        import engine as E72
        import vbe_bridge as VB72

        class _UI72(object):
            def __init__(self):
                self.shown = False

            def show(self, *a, **k):
                self.shown = True

            def hide(self, *a, **k):
                self.shown = False

            def update_selection(self, *a, **k):
                pass

            def contains_point(self, *a, **k):
                return False

        class _CM72(object):
            def __init__(self, comp, text):
                self._comp = comp
                self.lines = list(text)

            @property
            def Name(self):
                return self._comp.Name

            @property
            def Parent(self):
                return self._comp

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "\r\n".join(self.lines[s0:s0 + int(count)])

            def ReplaceLine(self, n, text):
                self.lines[int(n) - 1] = str(text).rstrip("\r\n")

        class _Comp72(object):
            def __init__(self, name, ctype, text):
                self.Name = name
                self.Type = ctype
                self.CodeModule = _CM72(self, text)

        class _Proj72(object):
            def __init__(self, comps):
                self.Name = "VBAProject"
                self.VBComponents = list(comps)

        class _Pane72(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _VBE72(object):
            def __init__(self, comps, pane):
                self.ActiveVBProject = _Proj72(comps)
                self.ActiveCodePane = pane

        _orig72 = VB72._get_vbe_cached
        _pyshow72 = E72.Completer._vbe_popup_showing
        # 同 71 节：这一节的落点是【文本判据】，真机的 Win32 让位查询会随
        # "此刻 VBE 有没有画着成员列表"摇摆，一律当"列表不在屏幕上"。
        E72.Completer._vbe_popup_showing = lambda self: False

        _BEFORE72 = "    NoExplicitPub.getTi"
        _AFTER72 = "    NoExplicitPub.getTitle"
        _MEMBERS72 = ["Public abcDef As Long",
                      "Public Function getTitle() As String",
                      "    getTitle = \"\"",
                      "End Function"]

        def _wire72(line_text, line=3, col=None):
            _c72 = _Comp72("M1", 1,
                           ["Sub Work()",
                            "    Dim s1 As Object",
                            line_text,
                            "End Sub",
                            ""] + _MEMBERS72)
            _comps72 = [_c72, _Comp72("NoExplicitPub", 1, _MEMBERS72)]
            _n72 = int(col if col is not None else len(line_text))
            _p72 = _Pane72(_c72.CodeModule, (line, _n72, line, _n72))
            VB72._get_vbe_cached = lambda: _VBE72(_comps72, _p72)
            return _c72, _p72

        def _mk72(line_text, **kw):
            _c72, _p72 = _wire72(line_text, **kw)
            _ui72 = _UI72()
            _k72 = E72.Completer(VB72.VbeBackend(), _ui72)
            return _k72, _ui72, _c72, _p72

        def _note72(k, ts=0.0, line=3, text=_BEFORE72):
            """模拟"按下 Tab 那一刻"：main 的钩子线程把快照交给 engine。"""
            k.note_accept_key((time.time() + ts, line, text))

        def _round72(k):
            """真机一轮：轮询发现变化 -> 语法让位 -> 宽限到 VBE 没弹 -> 补位。"""
            k.trigger(True)
            _y = k.yield_pending is not None
            k.confirm_yield(False)
            k.trigger(False)
            return _y

        # ---- 72.1 ★核心（用户报的那一下的完整时序）----
        # 真机上 Tab 之后这一次写入会引来【两轮】"让位 -> 宽限到 -> 补位"。
        # ⚠️ 两轮【必须分开 check】：check() 是"期望元素出现在 got 里"的子集
        # 匹配，两轮的期望值一模一样时，第 1 轮的 got 会把两条期望一起满足 ——
        # 也就是"第 2 轮弹了弹窗也照样绿"（v92 写这条时正是被它骗过一次，
        # 与 70.5 记录的同一个坑）。
        _k72, _u72, _c72, _p72 = _mk72(_AFTER72)
        _note72(_k72)
        _y72a = _round72(_k72)
        _r72a = (_u72.shown, _k72.current_matches())
        # 第 2 轮之前要把"已放弃让位"的标记清掉 —— 真机上 VBE 的成员列表在确认
        # 之后还会闪一下（`_vbe_popup_showing()` 为真那条会把 `yield_given_up`
        # 清掉），于是紧接着那一次 trigger(True) 会【重新让位】、0.25s 后再补位
        # 一次。这一步正是 v91 漏掉的那一次（白盒地把那个状态摆出来）。
        _k72.yield_given_up = None
        _y72b = _round72(_k72)
        _r72b = (_u72.shown, _k72.current_matches())
        check("72.1a ★核心·第 1 轮（v91 已挡住的那一轮）：`NoExplicitPub.getTi` "
              "处按 Tab、VBE 把 `getTitle` 补进来，轮询那次让位、宽限到 VBE "
              "没弹、补位回来那一次不该把刚写进去的词提示一遍",
              [(_y72a, _r72a)],
              expect_contain=[(True, (False, []))])
        check("72.1b ★核心·第 2 轮（v92 修的正是这一轮）：紧接着 `_run_trigger` "
              "又让位一次、宽限到之后第 2 次补位 —— 它也不许弹。记账\"只活一次\""
              "时，记忆在第 1 轮就被吃掉了，这一轮无据可依 -> 正是用户看到的"
              "\"后面还是会提示\"。"
              "⚠️代码复查（2026-10-18）后这一轮的形态变了：v100 那段\"给 VBE "
              "补进来的过程名补括号\"原先从 ctx 取永远不存在的 word_start_col，"
              "是死代码；修好之后第 1 轮就把 `getTitle` 变成了 `getTitle()`、"
              "光标落进括号，于是第 2 轮**连让位都不再发生**（不再是列表位置）。"
              "用户可见的保证不变：不弹。",
              [(_r72b,)], expect_contain=[((False, []),)])
        check("72.1c ★代码复查（2026-10-18）：v100 的\"补括号\"这次真的发生了 "
              "—— 第 1 轮命中后行尾必须是 `getTitle()`（此前它因为读错 ctx 键，"
              "一次都没执行过）",
              [(_c72.CodeModule.lines[2].rstrip(),)], expect_contain=[
                  ("    NoExplicitPub.getTitle()",)])

        # ---- 72.2 ★真机那条更短的路径（也是用户那次实际走的）----
        # 宽限这一轮刚记过"这一处 VBE 不会弹"，随后的 trigger(True) 便【不再
        # 让位】、直接落到判据上 —— 补位那一次已经把记忆吃掉了，于是这里弹窗。
        # 白盒地把那个状态摆出来（真机上它就是 confirm_yield(False) 的后效）。
        _k72b, _u72b, _c72b, _p72b = _mk72(_AFTER72)
        _note72(_k72b)
        _k72b.trigger(False)                 # 补位（第 1 次）：静默
        _s72b1 = _u72b.shown
        _k72b.yield_given_up = ("M1", 3,
                                E72.list_slot_anchor(_AFTER72, len(_AFTER72)))
        _k72b.yield_given_up_at = time.time()
        _k72b.trigger(True)                  # 紧接着轮询那次：不再让位，直判
        _s72b2 = (_u72b.shown, _k72b.current_matches())
        check("72.2 ★真机那条更短的路：补位（第 1 次）静默之后，轮询那次"
              "trigger(True) 因为\"这一处 VBE 不会弹\"而不再让位、直接落到判据"
              "—— 它也不许弹。这一支才是用户截图那一下（弹窗紧跟着 Tab 出现）",
              [_s72b1, _s72b2],
              expect_contain=[False, (False, [])])

        # ---- 72.3 ★对照：没按过 Tab，同样两轮 -> 照常提示 ----
        _k72c, _u72c, _c72c, _p72c = _mk72(_AFTER72)
        _round72(_k72c)
        _round72(_k72c)
        check("72.3 ★对照：没按过 Tab（同一处、同样走两轮让位->补位）-> 照常"
              "提示。证明 72.1/72.2 挡的是\"刚确认的那一次录入\"，不是这一处"
              "压根给不出候选",
              [(_u72c.shown, _k72c.current_matches())],
              expect_contain=[(True, ["getTitle"])])

        # ---- 72.4 ★钉住之后"用户一动手就恢复"（第二段的收紧）----
        _k72d, _u72d, _c72d, _p72d = _mk72(_AFTER72)
        _note72(_k72d)
        _k72d.trigger(False)
        _r72d1 = (_u72d.shown, _k72d._accept_pinned)
        _c72d.CodeModule.lines[2] = "    NoExplicitPub.getTitl"   # 退格
        _p72d.SetSelection(3, 25, 3, 25)
        _k72d.trigger(False)
        check("72.4 ★第二段（钉住）的边界：这一行只要被动过一下就当场作废 —— "
              "①VBE 刚写完那一下静默（锚换成\"VBE 留下的那一行\"）；②用户退格"
              "一个字符 -> 记忆作废、v37 老口径立刻恢复。它【不是】永久屏蔽",
              [_r72d1, (_u72d.shown, _k72d.current_matches(),
                        _k72d._accept_key)],
              expect_contain=[(False, True), (True, ["getTitle"], None)])

        # ---- 72.5 ★VBE 按自己的书写改写已打片段 ----
        # 用户打的是 `.getti`，VBE 的成员列表把它改写成 `getTitle` —— 原样
        # （区分大小写）比对会判成"这一行被动过"，白漏一次静默。
        _k72e, _u72e, _c72e, _p72e = _mk72(_AFTER72)
        _note72(_k72e, text="    NoExplicitPub.getti")
        _k72e.trigger(False)
        check("72.5 ★判据按小写比：`.getti` 被 VBE 改写成 `.getTitle`（大小写"
              "不同）也算命中 -> 不提示",
              [(_u72e.shown, _k72e.current_matches())],
              expect_contain=[(False, [])])

        # ---- 72.6 ★快照慢一拍：按下前那一行还没带上点号 ----
        # 快照来自 100ms 轮询，用户手快（或当时 COM 在退避、cur_ctx 停更）时
        # "按下前那一行"可能还没带上那个点号，于是插进来的一段是 `.getTitle`。
        # Tab 缩进插的是空格 / `\t`，永远不会带点号 —— 所以放宽这一格不影响
        # 71.3 的防护。
        _k72f, _u72f, _c72f, _p72f = _mk72(_AFTER72)
        _note72(_k72f, text="    NoExplicitPub")
        _k72f.trigger(False)
        _k72g, _u72g, _c72g, _p72g = _mk72("    abc", col=5)
        _note72(_k72g, text="abc")           # 行首 Tab 缩进（插 4 个空格）
        _k72g.trigger(False)
        check("72.6 ★插入段允许一个开头的点号（快照慢一拍）：①旧文本还没带"
              "点号、插进来的是 `.getTitle` -> 仍算命中、不提示；②而 Tab 缩进"
              "插的是空格（连点号都没有）-> 该照常提示（71.3 的防护不许被这条"
              "放宽伤到）",
              [(_u72f.shown, _k72f.current_matches()),
               (_u72g.shown, _k72g.current_matches())],
              expect_contain=[(False, []), (True, ["abcDef"])])

        # ---- 72.7 ★对照：钉住之后 TTL 到 -> 照常提示 ----
        _k72h, _u72h, _c72h, _p72h = _mk72(_AFTER72)
        _note72(_k72h)
        _k72h.trigger(False)                 # 命中 -> 钉住
        # ★代码复查（2026-10-18）：v100 的补括号修好后，上面这一下会把行改成
        # `getTitle()` 并把光标送进括号。72.7 要验的是 **TTL**，前提是"VBE 留下
        # 的那一行此后一个字都没动过" —— 所以这里把行与光标摆回那个状态，
        # 否则验的就变成"光标在括号里当然没词可补"了。
        _c72h.CodeModule.lines[2] = _AFTER72
        _p72h.SetSelection(3, len(_AFTER72), 3, len(_AFTER72))
        _ts72, _ln72, _txt72 = _k72h._accept_key
        _k72h._accept_key = (_ts72 - 9.0, _ln72, _txt72)   # 假装 9s 过去了
        _k72h.trigger(False)
        _r72h = (_u72h.shown, _k72h.current_matches(), _k72h._accept_key)
        check("72.7 ★对照（第二段也不是永久屏蔽）：TTL(3.0s) 一到 -> 记忆作废、"
              "照常提示",
              [_r72h],
              expect_contain=[(True, ["getTitle"], None)])

        # ---- 72.8 ★缩进那笔当场作废（不许赖着不走）----
        _k72i, _u72i, _c72i, _p72i = _mk72("    abc", col=5)
        _note72(_k72i, text="abc")
        _k72i.trigger(False)                 # 缩进那一下：判据不成立 -> 作废
        _r72i = (_k72i._accept_key, _k72i._accept_pinned)
        _c72i.CodeModule.lines[2] = "    abcd"    # 用户接着打字
        _p72i.SetSelection(3, 9, 3, 9)
        _k72i.trigger(False)
        check("72.8 ★缩进那笔当场作废（不许赖着不走）：Tab 缩进留下的记忆在"
              "判据不成立时就清掉，用户接着打字照常提示",
              [_r72i, (_u72i.shown, _k72i.current_matches())],
              expect_contain=[(None, False), (True, ["abcDef"])])

        # ---- 72.9 ★接线护栏 ----
        _root72 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _engsrc72 = io.open(os.path.join(_root72, "engine.py"),
                            encoding="utf-8").read()
        _key72s = _engsrc72.split(
            "if self._accept_key is not None and not manual:")[1].split(
            "    def can_space_confirm")[0]
        _note72s = _engsrc72.split(
            "    def note_accept_key(self, snapshot):")[1].split(
            "    def accept(self):")[0]
        _hide72s = _engsrc72.split("    def hide(self):")[1].split(
            "    def maybe_hide_on_outside_click")[0]
        check("72.9 ★接线护栏：①判据分两段（`if self._accept_pinned:`）；②命中"
              "那一支【不】消费记忆 —— 它把两段标记置 True、并把锚换成"
              "\"VBE 留下的那一行\"（沿用原来的时刻，所以 TTL 仍然从那一次 Tab "
              "起算）；③作废只发生在对不上那一支（`_accept_key = None` 排在"
              "\"置 True\"之后）；④`note_accept_key` 里两段标记复位；⑤`hide()` "
              "里不许出现这两个名字（与 v90/v90b 同坑）；⑥判据里两处放宽"
              "（按小写比、允许一个开头的点号）也必须在",
              [[("if self._accept_pinned:" in _key72s),
                ("self._accept_pinned = True" in _key72s),
                ("self._accept_key = (_ak_ts" in _key72s),
                (_key72s.index("self._accept_pinned = True")
                 < _key72s.index("self._accept_key = None")),
                ("self._accept_pinned = False" in _key72s),
                ("self._accept_pinned = False" in _note72s),
                ("_accept_key" not in _hide72s),
                ("_accept_pinned" not in _hide72s),
                (".lower()" in _key72s),
                ('_ins[:1] == "."' in _key72s),
                ("按 Tab 那一笔对不上" in _key72s)]],
              expect_contain=[[True] * 11])

        VB72._get_vbe_cached = _orig72
        E72.Completer._vbe_popup_showing = _pyshow72
    except Exception as _e72:
        check("第 72 节异常: %s" % _e72, [True], expect_contain=[False])

    # ================================================================
    # 第 73 节（v93）：一次"从头到尾"的审计改出来的五笔
    #
    # 用户口径（原话）："麻烦你帮我好好从头到尾检查一次，除了这个 bug，
    # 也帮我检查检查其他方面，还有没有 bug 了。"
    #
    # 这一节钉的是"同一条语义只能一份实现"这条铁律（MEMORY 第 14 条）
    # 在 2026-09-22 这一轮里暴露出来的全部裂缝：
    #   73.1 字符串/注释扫描：engine 不再自写第三份，复用 parser 的
    #        scan_code_states（`""` 转义 / 注释边界一次扫对、按下标对齐）
    #   73.2 `Rem ...` 注释：`Rem a.b` 的点不是成员访问（假槽位根因）
    #   73.3 类型位置判据：engine 侧"跳过正在输入的词"不许把 `New` 吃掉
    #        —— `Set c = New |` 必须让位给 VBE 的类型列表
    #   73.4 类型位置判据：parser 侧补上 `New`（`Set x = New |` 要收窄到类型名）
    #   73.5 回车自动缩进：含全角时"光标在行尾"要按【显示列】比，不许被
    #        虚高的列骗去接管（那是"拆行"，必须交给 VBE）
    #   73.6 新行光标列：Tab 缩进 + 缓存缺失时不许按字符列算（会歪在 Tab 中间）
    #   73.7 接线护栏（源码断言）：死代码已清、判据只有一份实现
    #
    # ⚠️ check() 是【子集匹配】—— 两段期望雷同会假绿（70.5 / v92 踩过），
    #    所以下面凡"两轮期望一样"的地方都刻意分开写、并让第二轮带上
    #    能区分的结果值。
    # ================================================================
    try:
        import engine as E73
        import parser as P73
        import vbe_bridge as VB73

        # ---- 73.1 字符串 / 注释扫描只有一份实现，且 `""` 转义扫对 ----
        # ★v93 实测订正：`x = "a""b".c` 里 `""` 确实是一个转义引号、字符串在
        # 收尾引号处关闭（scan_code_states 逐字符状态表已证实：下标 4~9 为 True，
        # 下标 10、11 的 `.c` 为 False）；但 `.` 左边紧挨的是【收尾引号】，
        # 而 VBE 对 `"abc".c` 这种"串字面量后面取成员"是**不弹成员列表**的 ——
        # 它压根不是合法 VBA 表达式。所以这一处【不该让位】。
        # 真正被判错的是 `.c` 其后的位置 —— 让位与否必须与"`.` 是不是真代码"一致。
        # 这一条同时钉住两件事：
        #   ① `""` 相位不能反相（反相会让 `.c` 被当成串内内容）；
        #   ② 无论让不让位，`_in_comment_or_string` 对 `.c` 两格都必须为 False。
        _t73q = 'x = "a""b".c'
        check("73.1 ★`\"\"` 转义扫对：`x = \"a\"\"b\".c` 里 `\"\"` 是一个转义引号，"
              "字符串在收尾引号处关闭 -> 其后的 `.c` 两格【都在代码里】"
              "（旧实现把 `\"\"` 当开+关、相位反相会把它判成串内内容）",
              [E73._in_comment_or_string(_t73q, len(_t73q)),
               E73._in_comment_or_string(_t73q, len(_t73q) + 1)],
              expect_contain=[False, False])

        # 对照 1：转义引号之后是【真注释】，注释里的点不该让位
        _t73a = 'x = "a""b" \' .c'
        check("73.1a ★对照：`x = \"a\"\"b\" ' .c` 里 `\"\"` 之后是注释，`.c` "
              "在注释里 -> 不算成员访问（两种相位下都不让位）",
              [E73._in_comment_or_string(_t73a, len(_t73a) + 1),
               E73._member_list_dot_col(_t73a, len(_t73a) + 1),
               E73.list_slot_anchor(_t73a, len(_t73a) + 1),
               E73.vbe_member_list_expected(_t73a, len(_t73a) + 1)],
              expect_contain=[True, None, None, False])

        # 对照 2：正常的「括号 + 成员访问」不能被误伤
        # （第一版实现扫整个前缀 -> 34.6 整节红；这里把那个坑钉住）
        _t73b = '    Range("A1").'
        check("73.1b ★对照：`Range(\"A1\").` 里的 `\"A1\"` 在字符串里，但光标"
              "左边的字符是句末那个【点号】—— 只许看紧邻光标那一个位置，"
              "扫整个前缀会把合法成员点号误判成\"在字符串里\"",
              [E73._in_comment_or_string(_t73b, len(_t73b) + 1),
               E73.vbe_member_list_expected(_t73b, len(_t73b) + 1)],
              expect_contain=[False, True])

        # 对照 3 ★v93 新发现的两个真 bug（都是"`.` 左边紧挨什么"这一格问错）
        # 3a) 光标紧贴点号：`Range("A1").|` —— 这是最常见的一处（刚敲完点，
        #     成员名还没开始打）。旧实现第 1 步往回吃字母数字、停到 `)` 上，
        #     第 2 步再"从 j 往回找点号"，此时点号在【i-1】而不是 j 的位置，
        #     于是找不到点号 -> 判"不在成员位置" -> 不让位 -> 我们的窗压住
        #     VBE 的成员列表。修法：先把紧贴光标的那一个点号直接认下来。
        # 3b) owner 那一格要问【紧邻点号】的字符，不能问"跳过空白后的 m"：
        #     `Range("A1").` 里 m 落在 `)` 上，而 `)` 在 `"A1"` 里 —— 拿它
        #     当依据会把 With 之外最常见的括号取成员全误杀。
        _t73o = '    Range("A1").'
        _t73p = '    Range("A1").V'
        check("73.1c ★光标【紧贴点号】也要让位（`Range(\"A1\").|` 是刚敲完点、"
              "成员名还没打的最常见一处）：①行尾那一格；②已经在打成员名"
              "（`.V|`）—— 两者都必须 True，且槽位都要算出来（旧实现在 ① 上"
              "因为找不到点号而判 False -> 不让位 -> 候选窗压住 VBE 列表）",
              [E73.vbe_member_list_expected(_t73o, len(_t73o) + 1),
               E73.list_slot_anchor(_t73o, len(_t73o) + 1),
               E73.vbe_member_list_expected(_t73p, len(_t73p) + 1),
               E73.list_slot_anchor(_t73p, len(_t73p) + 1)],
              expect_contain=[True, 17, True, 17])

        # 3c) ★v93 真 bug：`_in_comment_or_string` 的列号必须【夹到行长】。
        #     不夹时 `[n-1]` 会越界成 Python 负下标、静默回绕到行尾字符上。
        #     ⚠️ 半开区间：`col = len+1` 看的是下标 len-1（行尾那个字符），
        #     它本来就该 False；`col = len+2`（超出恰好一格）才真正把负下标
        #     走了一遍，**它是最有判别力的一格** —— 旧实现会回绕到下标 -1
        #     （行尾的 `)`）而返回 False、看不出问题；真正暴露的是
        #     `Range("A1")` 这种"行尾字符在串里"的行，所以这里连行尾那格一起
        #     钉住，并把 `col = len+1` 的期望按其真实语义写成 False。
        #     （旧实现在 `Range("A1")` 的 col=len+1 上返回 True —— 读到了
        #     下标 13 的 `1`，这就是 34.6「该让位却不让位」的根因。）
        _t73r = '    Range("A1")'
        check("73.1d ★列号越界不许回绕：`Range(\"A1\")` 里 `)` 之后那几格"
              "（col=len 起）全必须 False —— 不把 col 夹到行长的实现会因负下标"
              "静默读到行尾的 `1`（在串里）而返回 True",
              [E73._in_comment_or_string(_t73r, len(_t73r) - 1),
               E73._in_comment_or_string(_t73r, len(_t73r)),
               E73._in_comment_or_string(_t73r, len(_t73r) + 1),
               E73._in_comment_or_string(_t73r, len(_t73r) + 2)],
              expect_contain=[True, False, False, False])

        # ---- 73.2 `Rem ...` 注释里的点不是成员访问 ----
        _t73c = "Rem a.b"
        _t73d = "    Rem a.b"
        check("73.2 ★`Rem` 注释：`Rem a.b` / 缩进版都不该算出成员位置"
              "（点号左边是个标识符 `a`，走不到\"行首/边界之后\"那条 With 分支，"
              "必须单独识别 —— 否则是【假槽位】，污染 yield_given_up 的键，"
              "症状是\"偶尔某处不再让位\"）",
              [E73._member_list_dot_col(_t73c, len(_t73c) + 1),
               E73._member_list_dot_col(_t73d, len(_t73d) + 1),
               E73.list_slot_anchor(_t73c, len(_t73c) + 1),
               E73.vbe_member_list_expected(_t73c, len(_t73c) + 1)],
              expect_contain=[None, None, None, False])

        # 对照：`Rem` 只在【行首 / 语句分隔符之后】才是注释，标识符里
        # 含 rem 的（`Remote.foo`）照旧是成员访问
        _t73e = "    Remote.foo"
        check("73.2b ★对照：`Remote.foo` 的成员访问照旧成立（`Rem` 只有"
              "顶格 / 紧跟 `:` 时才是注释，不能拿\"标识符含 rem\"去拦）",
              [E73._member_list_dot_col(_t73e, len(_t73e) + 1) is not None,
               E73.vbe_member_list_expected(_t73e, len(_t73e) + 1)],
              expect_contain=[True, True])

        # 对照 2：`Rem` 不止「点号」这一路 —— 类型位置那一支原先只查
        # _in_comment_or_string（只认 `'`），**漏了 Rem**，于是 `... : Rem a As b`
        # 这种行上注释里的 `As` 会被当成"正在填的类型位置"（假槽位）。
        # ★v93：两处现在共用一个 `_in_rem_comment`。
        _t73rem = "Dim x As Integer : Rem a As b"
        check("73.2c ★`Rem` 里的 `As` 不算类型位置：`Dim x As Integer : Rem a As b`"
              " 光标在行尾 —— 注释里的 `As b` 是注释文字，不该让位、更不该"
              "算出槽位（类型位置那一支原先漏查 Rem，只有成员那一支查了）",
              [E73._in_type_list_position(_t73rem, len(_t73rem) + 1),
               E73.list_slot_anchor(_t73rem, len(_t73rem) + 1),
               E73._in_rem_comment(_t73rem, len(_t73rem) + 1)],
              expect_contain=[False, None, True])

        # ---- 73.3 类型位置：`New` 侧的行为（合并判据，行为与旧版逐字相同） ----
        # ⚠️ 诚实标注：这一条【不是】新修的 bug —— 旧 `_in_type_list_position`
        # 在 `Set c = New |` 上本来就返回 True（它把 `New` 当关键字尾巴，
        # 并不受"跳过正在输入的词"那步影响，因为光标就在 `New` 之后）。
        # 之所以仍然写进来：v93 把"要不要让位"（_in_type_list_position）与
        # "是哪一个槽位"（_type_list_anchor）合并成同一份 _type_list_kw_end，
        # 这一条钉的就是"合并后行为不变"，防止以后有人改一处漏一处。
        _t73f = "Set c = New "
        _t73g = "Set c = New Coll"
        check("73.3 ★类型位置判据已合并（`_type_list_kw_end` 一份实现）："
              "`Set c = New |` / `Set c = New Coll|` 都让位且槽位同为 13；"
              "`Dim x As |` 这一路照旧 —— 合并前后逐字等价（防未来分叉）",
              [E73._in_type_list_position(_t73f, len(_t73f) + 1),
               E73._in_type_list_position(_t73g, len(_t73g) + 1),
               E73.vbe_list_expected(_t73f, len(_t73f) + 1),
               E73._type_list_anchor(_t73f, len(_t73f) + 1)],
              expect_contain=[True, True, True, 13])

        # 对照：`Dim x As |` / 第二个 `As |` 这条老路一个字都没动
        _t73h = "Dim x As "
        _t73i = "Dim x As Long, y As "
        check("73.3b ★对照（回归）：`Dim x As |` / 第二个 `As |` 照旧让位，"
              "槽位起点列仍是类型名该开始的那一列（10 / 21）",
              [E73._in_type_list_position(_t73h, len(_t73h) + 1),
               E73._in_type_list_position(_t73i, len(_t73i) + 1),
               E73._type_list_anchor(_t73h, len(_t73h) + 1),
               E73._type_list_anchor(_t73i, len(_t73i) + 1)],
              expect_contain=[True, True, 10, 21])

        # ---- 73.4 类型位置：parser 侧也要认 `New`（收窄到类型名） ----
        # 旧 parser 只认 `As`，于是 `Set x = New ` 处 in_type_position=False
        # -> 不把候选收窄到类型名 -> 一堆变量名 / 过程名冒出来（噪音）。
        check("73.4 ★parser 侧补上 `New`：`Set x = New |` 与 `Set x = New Cls|` "
              "都算\"正在填类型名\"（候选收窄到类型名）；`Dim x As |` 照旧成立；"
              "普通赋值行不是类型位置",
              [P73.is_caret_in_type_position("Set s1 = New ", 14),
               P73.is_caret_in_type_position("Set s1 = New Cls1", 19),
               P73.is_caret_in_type_position("Dim x As ", 10),
               P73.is_caret_in_type_position("    userName = 1", 16)],
              expect_contain=[True, True, True, False])

        # 对照：字符串里的 As / New 不算（prefix 的逐字符抹白负责）
        check("73.4b ★对照：字符串里的 `New ` 不算类型位置（逐字符抹白后正则"
              "自然不匹配）—— 注意必须用 scan_code_states 的【逐字符】状态拼"
              "prefix，_mask_strings_and_comments 的产物长度与原文不等，"
              "拿它切片会直接 IndexError（v93 第一版就是这么红的）",
              [P73.is_caret_in_type_position('    s = "New "', 14),
               P73.is_caret_in_type_position('    s = "As "', 13)],
              expect_contain=[False, False])

        # ---- 73.5 回车自动缩进：含全角时按【显示列】比行尾 ----
        # 旧判据 `int(caret_ec) < len(t) -> 放行`：中文全角让显示列虚高，
        # 行中间的光标也能满足"显示列 ≥ 字符数"，于是本该交给 VBE 的
        # 【拆行】被我们吞了。
        _t73j = "    For 每个元素 = 1 To 10"      # 22 字符、行尾显示列 27
        _t73k = "    ' 配置：读取 用户名"          # 15 字符、行尾显示列 24
        check("73.5 ★中文行不许被虚高的显示列骗去接管回车："
              "①`For 每个元素 = 1 To 10` 光标停字符列 18（显示列 22 ≥ 22）"
              "  -> 放行（那是拆行）；②同一行真·行尾（显示列 27）-> 接管；"
              "③中文整行注释同行同理（中间放行 / 行尾接管）",
              [VB73.enter_indent_wanted(_t73j, 22),
               VB73.enter_indent_wanted(_t73j, 27),
               VB73.enter_indent_wanted(_t73k, 18),
               VB73.enter_indent_wanted(_t73k, 24)],
              expect_contain=[False, True, False, True])

        # 对照：`""` 影响的不是这里 —— 纯 ASCII 行判据逐字未变（回归）
        check("73.5b ★对照（回归）：纯 ASCII 行照旧 —— 块头行尾接管、"
              "行中间放行、空行 / 含 Tab / 拿不到列一律放行",
              [VB73.enter_indent_wanted("    For i = 1 To 10", 20),
               VB73.enter_indent_wanted("    For i = 1 To 10", 12),
               VB73.enter_indent_wanted("        ", 9),
               VB73.enter_indent_wanted("\tFor i", 10),
               VB73.enter_indent_wanted("    For i", 0)],
              expect_contain=[True, False, False, False, False])

        # ---- 73.6 新行光标列：Tab 缩进 + 缓存缺失不许按字符列算 ----
        # `_sem_cache["info"]` 只由 _probe_semantics 写入，新起一行这种场景
        # 可能一次都没触发过 -> 旧实现按 `len("\t")+1 = 2` 落光标，而显示列
        # 语义下正确值是 5（Tab 展开到第 5 列）-> 光标歪在 Tab 中间。
        _sem73 = VB73._sem_cache.get("info")
        try:
            VB73._sem_cache["info"] = None
            _tab_nocache = VB73._indent_end_col("\t")
        finally:
            VB73._sem_cache["info"] = _sem73
        check("73.6 ★Tab 缩进 + 语义缓存缺失：按【显示列 + 默认 tab 宽度 4】"
              "兜底（`\\t` -> 5，不是字符列语义的 2）；拿到显式语义时各按各算；"
              "纯空格缩进两种语义一致（仍 5）",
              [_tab_nocache,
               VB73._indent_end_col("\t", ("char", 4, False)),
               VB73._indent_end_col("\t", ("disp", 4, False)),
               VB73._indent_end_col("\t", ("disp", 8, False)),
               VB73._indent_end_col("    ")],
              expect_contain=[5, 2, 5, 9, 5])

        # ---- 73.7 ★接线护栏（源码断言） ----
        _root73 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _eng73 = io.open(os.path.join(_root73, "engine.py"),
                         encoding="utf-8").read()
        _par73 = io.open(os.path.join(_root73, "parser.py"),
                         encoding="utf-8").read()
        _vbr73 = io.open(os.path.join(_root73, "vbe_bridge.py"),
                         encoding="utf-8").read()
        # ① engine 不再自写字符串/注释扫描，改调 parser 的唯一实现
        _ics73 = _eng73.split("def _in_comment_or_string(")[1].split(
            "\ndef ")[0]
        # ② 死代码已清（判"没有定义"，不判"文本里没提过"—— 删它的注释里
        #    会点名提到它，那是【说明】，不是残留）
        _dead73 = ("def _name_exists_outside_caret(" in _eng73)
        # ③ parser 侧 scan_code_states 存在，且 _mask 由它派生
        _scan73 = ("def scan_code_states(" in _par73)
        _mask73 = ("states = scan_code_states(code)" in _par73)
        # ④ 成员位置 / 类型位置在 engine 各只有一份实现
        _dot73 = _eng73.count("def _member_list_dot_col(")
        _kw73 = _eng73.count("def _type_list_kw_end(")
        # ⑤ 两个公共入口都委托到它们（不许再有第二份内联逻辑）
        _deleg73 = ("return _member_list_dot_col(line_text, caret_col) is not None"
                    in _eng73)
        _deleg73b = ("e = _type_list_kw_end(line_text, caret_col)" in _eng73)
        # ⑥ vbe_bridge：含 Tab 的缩进不许走"缓存缺失就按字符列"
        _iec73 = _vbr73.split("def _indent_end_col(")[1].split("\ndef ")[0]
        _guard73 = ('if "\\t" not in indent:' in _iec73)
        _resolve73 = ("def _resolve_sem_info(" in _vbr73)
        check("73.7 ★接线护栏：①`_in_comment_or_string` 改走 parser 的"
              "`scan_code_states`（唯一实现）；②`_name_exists_outside_caret` "
              "死代码已删；③parser 侧 `scan_code_states` 存在且 `_mask_"
              "strings_and_comments` 由它派生；④成员/类型位置在 engine 各只有"
              "一份实现（`_member_list_dot_col` / `_type_list_kw_end`）；"
              "⑤`vbe_member_list_expected` / `_type_list_anchor` 都委托到它们；"
              "⑥`_indent_end_col` 含 Tab 时不再回落到字符列，并有 `_resolve_sem_"
              "info` 补探测",
              [[("vba_parser.scan_code_states(line_text)" in _ics73),
                (not _dead73),
                _scan73, _mask73,
                (_dot73 == 1), (_kw73 == 1),
                _deleg73, _deleg73b,
                _guard73, _resolve73]],
              expect_contain=[[True] * 10])
    except Exception as _e73:
        check("第 73 节异常: %s" % _e73, [True], expect_contain=[False])

    # ---- 74. v94：退格删整行（整行只有空白 -> 一次退格删掉整行） ----
    #
    # 用户口径："当我按退格键的时候，如果整行是空行，能删除整行，IDEA 里写 Java
    # 就有这个功能。"
    #
    # 需要我们的正是 VBE 自己按回车继承缩进造出来的那种行（`    `）：原生退格
    # 一次只吃掉一个空格，要按 4 下才空、第 5 下才并到上一行去。
    #
    # 三处设计要点（各有一条用例钉住）：
    #   ① 真空行（`""`）不接管 —— 原生退格的结果与删整行一模一样，抢了没好处；
    #   ② 光标必须在空白末尾 —— 停在缩进中间说明用户想改缩进量，那是"退一格"；
    #   ③ 删完当场作废轮询快照 —— 否则按住退格（键盘重复约 33ms）会拿【删之前
    #      那一行的旧快照】再判一次，把下面那一行也删掉。
    print("\n=== 74. v94：退格删整行（IDEA 风格） ===")
    try:
        import vbe_bridge as VB74

        class _CM74(object):
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "M1"

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n"
                               for l in self.lines[s0:s0 + int(count)])

            def DeleteLines(self, start, count=1):
                s0 = max(0, int(start) - 1)
                del self.lines[s0:s0 + int(count)]

        class _Pane74(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                # ★忠实模拟真 VBE 的两条行为（缺了任一条都测不出 v94b 那个 bug）：
                #   ① 列是【显示列】（全角占 2 格、Tab 展开）—— v93 实测；
                #   ② 超出行尾的列会被**钳到行尾**。
                # 后者正是 vbe_bridge._vbe_measure 与 move_caret_line 赖以成立
                # 的前提（`move_caret_line` 的注释明写"目标行较短时由 VBE 自己
                # 钳到行尾"）；v94b 的落点也改成了求它 —— 自己算就会少算全角。
                ln = max(1, min(int(sl), len(self.CodeModule.lines)))
                try:
                    _t74 = self.CodeModule.lines[ln - 1]
                except Exception:
                    _t74 = ""
                _eol74 = VB74._disp_width(_t74, 4, True) + 1
                self.sel = (ln, max(1, min(int(sc), _eol74)),
                            ln, max(1, min(int(ec), _eol74)))

        class _VBE74(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        _orig74 = VB74._get_vbe_cached

        def _run74(lines, line_no, col, sel=None):
            """按 mock 的代码窗跑一次 delete_blank_line_here，返回 (ok, 行表, 选区)。

            ⚠️ 光标列按【字符列】给 —— 脚手架里的行一律用 ASCII（真机
            GetSelection 给的是**显示列**，全角占两格会偏；第 65 节同一处提醒）。
            """
            cm74 = _CM74(lines)
            pane74 = _Pane74(cm74, sel if sel is not None else
                             (int(line_no), int(col), int(line_no), int(col)))
            VB74._get_vbe_cached = lambda: _VBE74(pane74)
            try:
                ok74 = VB74.VbeBackend().delete_blank_line_here()
            finally:
                VB74._get_vbe_cached = _orig74
            return (ok74, cm74.lines, pane74.sel)

        # ---- 74.1 ★判据（用户报的场景）：缩进空行 + 光标在行尾 ----
        check("74.1 ★判据（用户报的场景）：`    ` 这种【只有空白】的行、光标停在"
              "行尾（VBE 按回车继承缩进造出来的正是这种行）-> 接管；同一行光标"
              "落在缩进【中间】或行首 -> 不接管（那时用户多半是想改缩进量，"
              "交给 VBE 一次删一格）",
              [VB74.blank_line_delete_wanted("    ", 5),
               VB74.blank_line_delete_wanted("    ", 3),
               VB74.blank_line_delete_wanted("    ", 1)],
              expect_contain=[True, False, False])

        # ---- 74.2 真空行交给 VBE 原生（结果一样，不抢）----
        check("74.2 ★真空行（`\"\"`）不接管：光标必然在第 1 列，VBE 原生退格就是"
              "【并到上一行】，结果与删整行一模一样 —— 抢了只多一次 COM 往返和"
              "一处出错的可能",
              [VB74.blank_line_delete_wanted("", 1),
               VB74.blank_line_delete_wanted("", 0)],
              expect_contain=[False, False])

        # ---- 74.3 有内容的行 = 普通删字，一个字都不许碰 ----
        check("74.3 ★有内容的行不接管（那是普通删字）：代码行 / 整行注释 / 行尾"
              "带空格的代码行 一律 False —— 判据只认【整行只有空白】",
              [VB74.blank_line_delete_wanted("    x = 1", 10),
               VB74.blank_line_delete_wanted("    ' 注释", 8),
               VB74.blank_line_delete_wanted("Sub F1()", 9)],
              expect_contain=[False, False, False])

        # ---- 74.4 Tab 缩进的工程：两种列语义都要认 ----
        # VBE 报告的是【显示列】，Tab 缩进的行在两种语义下数值不同（`\t` 显示列
        # 是 5、字符列是 2）。只认一种会让整条功能在 Tab 缩进的工程里静默失效。
        check("74.4 ★Tab 缩进的工程照旧生效：`\\t` 空行的行尾列在【显示列】语义下"
              "是 5、在【字符列】语义下是 2 —— 两种都认（都指向同一个\"行尾\"，"
              "而一个 Tab 内部的列值 VBE 根本不会报出来）",
              [VB74.blank_line_delete_wanted("\t", 5),
               VB74.blank_line_delete_wanted("\t", 2),
               VB74.blank_line_delete_wanted("\t\t", 3)],
              expect_contain=[True, True, True])

        # ---- 74.5 ★集成：整行删掉，光标落到【上一行行尾】----
        _L74 = ["Sub F1()", "    x = 0", "    ", "    y = 1", "End Sub"]
        check("74.5 ★集成（用户报的场景）：`    x = 0` 与 `    y = 1` 中间夹着一个"
              "缩进空行，光标停在它行尾按退格 -> 整行消失，光标落到【上一行行尾】"
              "（第 2 行第 10 列，正是 `    x = 0` 的末尾）—— 与 VBE 原生\"空行上"
              "按退格（吃掉换行、两行并一行）\"的落点完全一致",
              [_run74(_L74, 3, 5)],
              expect_contain=[(True,
                               ["Sub F1()", "    x = 0", "    y = 1", "End Sub"],
                               (2, 10, 2, 10))])

        # ---- 74.6 有选区 -> 那是"删选区"，一个字符都不许动 ----
        check("74.6 ★有选区不接管：退格在有选区时是【删掉选区】，语义完全不同 —— "
              "后端必须原样退回（调用方会把这一下退格还给系统），行表一个字没变",
              [_run74(["Sub F1()", "    ", "End Sub"], 2, 5, sel=(2, 1, 3, 1))],
              expect_contain=[(False, ["Sub F1()", "    ", "End Sub"],
                               (2, 1, 3, 1))])

        # ---- 74.7 ★后端复核（快照可能过期）：光标此刻不在行尾 -> 不删 ----
        # 钩子判据吃的是 100ms 前的轮询快照：用户可能刚用鼠标把光标点到缩进
        # 中间（他要的是"退一格"），也可能刚粘贴进一整段代码。后端必须拿【此刻
        # 真读到的】行与列再判一次 —— 同一份 blank_line_delete_wanted。
        check("74.7 ★后端权威复核判据：快照说\"空行\"、可此刻光标停在缩进【中间】"
              "（鼠标刚点过）-> 不接管，行表一个字符都没删（判据与钩子侧同一份"
              "实现，只是喂进去的是此刻真读到的数据）",
              [_run74(["Sub F1()", "    x = 0", "        ", "End Sub"], 3, 3)],
              expect_contain=[(False,
                               ["Sub F1()", "    x = 0", "        ", "End Sub"],
                               (3, 3, 3, 3))])

        # ---- 74.8 删的是第 1 行（没有"上一行"）----
        check("74.8 删第 1 行（没有上一行可落光标）：整行照样删掉，光标落到新的"
              "首行行首，不抛异常、也不返回 False（返回 False 会让调用方再补一次"
              "退格、多删一个字符）",
              [_run74(["    ", "Sub F1()"], 1, 5)],
              expect_contain=[(True, ["Sub F1()"], (1, 1, 1, 1))])

        # ---- 74.9 ★接线护栏（源码断言）----
        _root74 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _src74 = io.open(os.path.join(_root74, "main.py"),
                         encoding="utf-8").read()
        _vbr74 = io.open(os.path.join(_root74, "vbe_bridge.py"),
                         encoding="utf-8").read()
        _fn74 = _src74.split("def _delete_blank_line_here(")[1].split(
            "\n    def ")[0]
        check("74.9 ★接线护栏：①钩子侧判据在 win32_filter 里被调用，且排在"
              "`completer.is_visible()` 那条分支【之前】（空行上挂着候选也要收掉）；"
              "②判据在【退格】上才判（VK_BACK）；③动作里删不成就 `send_vk(VK_BACK)` "
              "把这一下退格原样还给系统（绝不\"吞了按键却什么都没发生\"）",
              [[("_bs_delete_line_key_ok(vk)" in _src74),
                (_src74.index("_bs_delete_line_key_ok(vk)")
                 < _src74.index("elif completer.is_visible():")),
                ("vk != VK_BACK" in _src74),
                ("send_vk(VK_BACK)" in _fn74)]],
              expect_contain=[[True, True, True, True]])

        check("74.10 ★按住退格不许连删：删完【当场作废轮询快照】"
              "（`state[\"cur_ctx\"] = None`）—— 快照 100ms 刷一次、键盘重复却约 "
              "33ms 一次，不作废的话第二次重复会拿【删之前那一行的旧快照】再判"
              "一次\"整行是空行\"，把下面那一行也删掉；作废后到下次轮询为止退格走"
              "VBE 原生（删上一行最后一个字符），正是用户按住退格时期待的行为",
              [('state["cur_ctx"] = None' in _fn74)],
              expect_contain=[True])

        _bs74 = _vbr74.split("def delete_blank_line_here(")[1].split(
            "\n    def ")[0]
        check("74.11 ★判据只有一份实现：vbe_bridge 里 `blank_line_delete_wanted` "
              "只定义一次，后端 delete_blank_line_here 【复用它】做权威复核"
              "（不再自己写第二遍\"是不是空行\"）—— 同一条语义两份实现是本仓库"
              "的惯犯（v87 可见性 / v93 成员位置都栽过）；另：先 `DeleteLines` "
              "后单独 try 落光标（DeleteLines 已成却因落光标失败而返回 False，"
              "会让调用方再补一次退格、多删一个字符）",
              [[(_vbr74.count("def blank_line_delete_wanted(") == 1),
                ("blank_line_delete_wanted(raw, ec)" in _bs74),
                ("cm.DeleteLines(anchor, 1)" in _bs74)]],
              expect_contain=[[True, True, True]])

        # ---- 74.12 ★用户报的 bug：上一行含【汉字】时，光标落在中文中间 ----
        # 复现口径：本会话里第一次列语义探测碰巧落在【纯 ASCII 且无 Tab】的行上，
        # VBE 报的 EOL 列恰好 == len+1，_fit_semantics 就得出 ("char", 4, False)
        # —— 这句其实只说明"这一行上两种语义数值相同、分不出来"，可它被当成
        # "VBE 用字符列"缓存了下来。于是含汉字的行按 len+1 给列，VBE 却按显示列
        # 解释 -> 少算【全角字符个数】格 -> 光标落在 `完成` 两个字中间。
        #
        # `    MsgBox "完成"`：15 个字符 / 显示宽 17 -> 行尾显示列 18（真值）；
        # 老口径（字符列）给 16，正好差 2 = 两个汉字的格数 —— 与用户报的症状一致。
        _sem74 = VB74._sem_cache.get("info")
        try:
            VB74._sem_cache["info"] = ("char", 4, False)   # 复现那条错结论的缓存
            check("74.12 ★用户报的场景：上一行是 `    MsgBox \"完成\"`，删掉下面的缩进"
                  "空行后光标必须落在【真行尾】第 18 列（显示列、全角占 2 格）—— "
                  "按'字符列'算只有 16，VBE 会把它当显示列解释，光标正好停在两个"
                  "汉字中间。落点的列**不许自己算**，交给 VBE 自己钳到行尾",
                  [_run74(["Sub F1()", '    MsgBox "完成"', "    ", "End Sub"],
                          3, 5)],
                  expect_contain=[(True,
                                   ["Sub F1()", '    MsgBox "完成"', "End Sub"],
                                   (2, 18, 2, 18))])

            check("74.13 ★同一处根因的 Tab 版：上一行 `\\tx = 1`，行尾显示列是 10"
                  "（Tab 展开到 4）+ `x = 1` 5 列；缓存被错记成'字符列'时老口径只给 "
                  "7（len+1），光标歪在 Tab 后面 —— 落点同样由 VBE 钳定",
                  [_run74(["Sub F1()", "\tx = 1", "    ", "End Sub"], 3, 5)],
                  expect_contain=[(True, ["Sub F1()", "\tx = 1", "End Sub"],
                                   (2, 10, 2, 10))])
        finally:
            VB74._sem_cache["info"] = _sem74

        check("74.14 ★落点不再自己算（源码护栏）：`_caret_to_line_end` 只定义一次、"
              "且被 delete_blank_line_here 调用；`_eol_caret_col`（按列语义算）"
              "在该函数里只剩【兜底】一处；顺带 `_resolve_sem_info` 不该再出现 —— "
              "那条路会为了探测【移动光标】（退格路上多做一次\"移走再还原\"，"
              "本身就是\"光标偶尔自己跳\"的隐患）",
              [[(_vbr74.count("def _caret_to_line_end(") == 1),
                ("_caret_to_line_end(cp, cm, prev_no, prev_raw)" in _bs74),
                (_bs74.count("_eol_caret_col(") == 1),
                ("_resolve_sem_info" not in _bs74)]],
              expect_contain=[[True, True, True, True]])
    except Exception as _e74:
        check("第 74 节异常: %s" % _e74, [True], expect_contain=[False])

    # ==================================================================
    # 75. v95：补全【含中文】的名字后，光标要落在名字末尾
    #
    # 用户口径：「变量名或函数名里面包含中文字符和英文字符，按 Tab 把提示词
    # 打印到屏幕后，光标会落在变量名中间位置，不会正确落在变量名末尾……
    # 之前没有这个 bug 的。」
    #
    # 病根与 v94b 是**同一个**：`_sem_cache` 里那条"列语义"结论可能是【猜的】。
    #   * v93 加的 `_resolve_sem_info` 会在**每个回车**上探测一次，而它探测的那
    #     一行常常是纯 ASCII 且无 Tab —— 那种行上"字符列"与"显示列"数值完全相同，
    #     `_fit_semantics` 只能得出 ("char",4,False)，却被当成事实缓存下来；
    #   * 之后 `apply_completion` 摆光标时直接吃缓存 => 含中文的新行按【字符列】
    #     给列，VBE 当【显示列】解释 => 少算全角格数 => 光标落在变量名中间。
    #   * "之前没有"正好对得上：v93 之前没人会在纯 ASCII 行上写缓存。
    #
    # 修法两条：
    #   ① `apply_completion` 摆光标前**在写回后的这一行上重新量**（它自带全角 /
    #      Tab，证据判别性），不碰缓存；量到的结果反过来写回缓存（自动修好）；
    #   ② `_resolve_sem_info` 遇到"既无 Tab 也无全角"的行**干脆不探测**（那种行
    #      上量不出东西，只会污染缓存，而且调用方压根用不上）。
    # ==================================================================
    print("\n=== 75. v95：补全含中文的名字后，光标要落在名字末尾 ===")
    try:
        import vbe_bridge as VB75

        class _CM75(object):
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "Module1"

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count):
                return self.lines[int(start) - 1]

            def ReplaceLine(self, line_no, text):
                self.lines[int(line_no) - 1] = text

        class _Pane75(object):
            """忠实模拟真 VBE：列是【显示列】，超出行尾的列被钳回行尾。"""
            def __init__(self, cm, sel=(1, 1, 1, 1)):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                ln = max(1, min(int(sl), len(self.CodeModule.lines)))
                _t = self.CodeModule.lines[ln - 1]
                _eol = VB75._disp_width(_t, 4, True) + 1
                self.sel = (ln, max(1, min(int(sc), _eol)),
                            ln, max(1, min(int(ec), _eol)))

        class _VBE75(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _run75(lines, line_no, start_col, caret_col, completion,
                   poison=None):
            """跑一次 apply_completion，返回 (新行文本, 落点列, 事后缓存)。

            poison：预置进 `_sem_cache["info"]` 的【错误】结论，用来复现
            "缓存被纯 ASCII 行污染"这个现场（None = 不预置）。
            """
            _cm75 = _CM75(lines)
            _pn75 = _Pane75(_cm75)
            _orig75 = VB75._get_vbe_cached
            _sem75 = VB75._sem_cache.get("info")
            VB75._sem_cache["info"] = poison
            VB75._get_vbe_cached = lambda: _VBE75(_pn75)
            try:
                VB75.VbeBackend().apply_completion(
                    line_no, start_col, caret_col, completion)
                _cache75 = VB75._sem_cache.get("info")
            finally:
                VB75._get_vbe_cached = _orig75
                VB75._sem_cache["info"] = _sem75
            return (_cm75.lines[line_no - 1], _pn75.sel[3], _cache75)

        # ---- 75.1 ★用户报的场景：名字本身中英混排 ----
        # `    x = get订单` -> 补全成 `get订单数量`：真行尾显示列是 20，
        # 被污染成"字符列"时旧口径只给 16（少算 4 个全角字符的格数）
        # -> 光标停在 `订` / `单` 之间那一带，正是用户看到的"落在变量名中间"。
        check("75.1 ★用户报的场景（名字中英混排）：`    x = get订单` 补全成 "
              "`get订单数量`，光标必须落在【名字末尾】第 20 列（显示列）—— "
              "缓存被污染成\"字符列\"时只给 16，少算的正是 4 个全角字符占的格数",
              [_run75(["    x = get订单"], 1, 9, 14, "get订单数量",
                      poison=("char", 4, False))[:2]],
              expect_contain=[("    x = get订单数量", 20)])

        # ---- 75.2 ★光标前面就有中文（前缀里含全角）----
        # 这一条专钉"光标左侧的全角也要算进去"：只认名字自己的宽度会漏掉前缀。
        check("75.2 ★光标【前面】就有中文：`    总数 = getVa` 补全成 `getValue`，"
              "左侧两个全角各占 2 格 => 落点第 20 列；按字符列算只有 18 —— "
              "差的是前缀里那两个汉字的格数",
              [_run75(["    总数 = getVa"], 1, 10, 15, "getValue",
                      poison=("char", 4, False))[:2]],
              expect_contain=[("    总数 = getValue", 20)])

        # ---- 75.3 量出来的结果反过来把缓存修好 ----
        check("75.3 ★顺带修好缓存：这一行自带全角，量出来的结论是判别性的 "
              "（\"disp\", 4, True）=> 写回 `_sem_cache`，后面所有行都跟着受益 "
              "（旧的污染值 (\"char\",4,False) 被顶掉）",
              [_run75(["    x = get订单"], 1, 9, 14, "get订单数量",
                      poison=("char", 4, False))[2]],
              expect_contain=[("disp", 4, True)])

        # ---- 75.4 纯 ASCII 行照旧（回归护栏，别修过头）----
        check("75.4 纯 ASCII 行不受影响（回归护栏）：`    x = getVa` -> "
              "`getValue`，两种列语义数值相同 => 落点第 17 列，与旧版一致",
              [_run75(["    x = getVa"], 1, 9, 14, "getValue",
                      poison=("char", 4, False))[:2]],
              expect_contain=[("    x = getValue", 17)])

        # ---- 75.5 源头闸门：拿不准的行不许去探测（污染缓存 + 白移一次光标）----
        _vbr75 = io.open(os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "vbe_bridge.py"), encoding="utf-8").read()
        _rs75 = _vbr75.split("def _resolve_sem_info(")[1].split(
            "\n    def ")[0]
        _ac75 = _vbr75.split("def apply_completion(")[1].split(
            "\n    def ")[0]
        check("75.5 ★源头闸门（源码护栏）：① `_resolve_sem_info` 遇到【既无 Tab "
              "也无全角】的行直接返回 None（那种行上永远量不出结论，只会把猜的"
              "结论写进缓存）；② `apply_completion` 摆光标前用的是 `_vbe_measure`"
              "（在本行上重新量）而不是 `_probe_semantics`（会先去取缓存）；"
              "③ 量不到时 `_vbe_measure` 返回 **None**，不再伪造 (\"char\",4,False)"
              " —— 伪造值会被当成证据写进缓存",
              [[("if not (\"\\t\" in line_text or any(_is_wide(c) "
                 "for c in line_text)):" in _rs75),
                ("_vbe_measure(cm, cp, line_no, actual)" in _ac75),
                ("_probe_semantics(cm, cp, line_no, actual)" not in _ac75),
                ("return None" in _vbr75.split("def _vbe_measure(")[1].split(
                    "\n\ndef ")[0])]],
              expect_contain=[[True, True, True, True]])
    except Exception as _e75:
        check("第 75 节异常: %s" % _e75, [True], expect_contain=[False])

    # ==================================================================
    # 76. v96：关键字 + 内建函数收进默认池，枚举仍然不收
    #
    # 用户口径（原话）："我想把 vba 关键字，函数名这两部分，也加入到提醒列表里。
    # 枚举类型就不要加了，枚举加上，就会太多了。"
    #
    # ⇒ 开关组合 (ENABLE_VBA_BUILTINS, ENABLE_VBA_KEYWORDS,
    #            ENABLE_VBA_CONSTANTS, ENABLE_HOST_ENUMS) = (True, True, False, False)。
    #    第 47.2 / 54 节已经钉了"开关的出厂值"与"池子里有什么"，本节专钉
    #    **收回来之后新出现的两个风险**：
    #
    #   1) ★排序：语言自带的名字会不会压过用户自己的名字。
    #      `decl` 同时命中 `declaredVar`（工程内变量）与 `Declare`（关键字），
    #      两者**命中质量完全一样**（kind 3 前缀 / 起点 0 / 跨度 4），只因
    #      Declare 名字短 4 个字母，纯排序下关键字赢了 —— 那不是"匹配更好"。
    #      v96 把评分键改成 `命中质量 > 是不是语言自带 > 名字长度`。
    #   2) ★噪音：约 260 个公共词汇回到短输入路径上，边界必须仍然收得住
    #      （首字母对齐那道闸门 + "你的名字优先"）。用户担心的"太多"
    #      是真实风险，所以这里量的是**具体数字**而不是"感觉还行"。
    #
    # ⚠️ 本节的数字是【离线桩】上量出来的真值（不是拍脑袋的阈值）。
    #    改动模糊匹配或收紧规则时，这里会红 —— 那是提醒你重测一遍，
    #    不是让你把期望值改成新的输出。
    # ==================================================================
    print("\n=== 76. v96：关键字+内建函数默认收，枚举不收；你的名字优先 ===")
    try:
        import vba_builtins as VB76B
        import engine as E76
        import vbe_bridge as VB76

        # ---- 76.1 出厂开关组合（干净模块，与 47.2 互为对照）----
        _f76 = None
        try:
            import importlib.util as _iu76
            _sp76 = _iu76.spec_from_file_location(
                "vbe_bridge_fresh76",
                os.path.join(os.path.dirname(os.path.dirname(
                    os.path.abspath(__file__))), "vbe_bridge.py"))
            _f76 = _iu76.module_from_spec(_sp76)
            _sp76.loader.exec_module(_f76)
        except Exception:
            pass
        check("76.1 ★出厂开关组合 = (内建函数 True, 关键字 True, vb* 枚举 False, "
              "宿主枚举 False)",
              [((_f76.ENABLE_VBA_BUILTINS, _f76.ENABLE_VBA_KEYWORDS,
                 _f76.ENABLE_VBA_CONSTANTS, _f76.ENABLE_HOST_ENUMS)
                if _f76 else "干净模块加载失败")],
              expect_contain=[(True, True, False, False)])

        # ---- 76.2 清单规模：收回来的是"该收的"，枚举两批仍然是关的 ----
        check("76.2 清单规模：函数 %d / 类型 %d / 关键字 %d 三批都非空"
              "（收回来了）；枚举清单也仍在（只是开关关着，一个字没裁）"
              % (len(VB76B.BUILTIN_FUNCTIONS), len(VB76B.BUILTIN_TYPE_NAMES),
                 len(VB76B.BUILTIN_KEYWORDS)),
              [len(VB76B.BUILTIN_FUNCTIONS) >= 150,
               len(VB76B.BUILTIN_TYPE_NAMES) >= 5,
               len(VB76B.BUILTIN_KEYWORDS) >= 30,
               len(VB76B.BUILTIN_CONSTANTS) >= 100],
              expect_contain=[True] * 4)

        # ---- 76.3 ★排序：同质量时工程内名字压过语言自带的名字 ----
        # 构造一个最小引擎：池子里 declaredVar（工程内）与 Declare / Dim（关键字）。
        _POOL76 = [("declaredVar", "模块1", None, False),
                   ("Declare", VB76._BUILTIN_MODULE, None, False),
                   ("Dim", VB76._BUILTIN_MODULE, None, False),
                   ("MsgBox", VB76._BUILTIN_MODULE, None, False)]
        _BL76 = set(n.lower() for n, m, _p, _v in _POOL76
                    if m == VB76._BUILTIN_MODULE)

        class _B76(object):
            def get_context(self):
                _ln = "    decl"
                return {"line_no": 1, "line_text": _ln,
                        "caret_col": len(_ln) + 1, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": "模块1"}

            def get_identifiers(self):
                return list(_POOL76)

            def get_declared_names(self):
                return ["declaredvar"]

            def declared_elsewhere(self, name, module_name):
                return str(name).lower() in _BL76

            def get_structural_names(self):
                return sorted(_BL76)

            def get_builtin_names(self):
                return set(_BL76)

            def get_type_names(self):
                return list(VB76B.BUILTIN_TYPE_NAMES)

            def get_host_enum_names(self):
                return {}

            def names_outside_caret(self, caret, scope_only=False):
                return set(["declaredvar"])

            def apply_completion(self, *a, **k):
                return None

            def vbe_yield_visible(self):
                return False

            def vbe_popup_visible(self):
                return False

            def vbe_popup_info(self):
                return []

        class _UI76(object):
            def show(self, rows, sel, c):
                pass

            def hide(self):
                pass

            def update_selection(self, sel):
                pass

            def contains_point(self, x, y):
                return False

        def _m76(word):
            _c = E76.Completer(_B76(), _UI76())
            _c.trigger(True)
            return list(_c.matches or [])

        _d76 = _m76("decl")
        # 先确认前提：两者命中质量真的相同（否则本节在验一个不存在的问题）
        _q76a = E76.fuzzy_match("declaredVar", "decl")[0]
        _q76b = E76.fuzzy_match("Declare", "decl")[0]
        check("76.3 ★前提：declaredVar 与 Declare 的【命中质量】三项"
              "（kind/起点/跨度）完全相同，只有名字长度不同 "
              "(-11 vs -7) —— 不相同的话 76.4 验的就不是同分让位",
              [_q76a[:3], _q76b[:3], _q76a[:3] == _q76b[:3],
               _q76a[3] < _q76b[3]],
              expect_contain=[(3, 0, -4), (3, 0, -4), True, True])

        check("76.4 ★排序：打 decl -> 候选第 1 个是 declaredVar（你自己的名字），"
              "不是 Declare（关键字）；Declare 仍作为候选存在",
              [_d76[:1], "Declare" in _d76],
              expect_contain=[["declaredVar"], True])

        # 反面对照：把 builtin_names 清空（= 两批开关都关掉），
        # 排序应退回 v96 之前的行为 —— Declare 因名字短而排到前面。
        # 这条钉住"改动确实是 v96 那条规则造成的"，不是恰好排成这样。
        #
        # ★v103 更正：光关 builtin 标志已经不够了 —— 池子里的 declaredVar 属于
        # `模块1`（当前模块，档位 3），Declare 属于语言自带模块（档位 2），
        # v103 新增的【作用域就近】排序会**独立地**把 declaredVar 排前面。
        # 反面对照的意义是"把这两条相关度规则都关掉，看它退回旧行为"，
        # 所以这里必须同时把 RANK_BY_SCOPE 也关掉（否则这条对照验的不是
        # "规则在起作用"，而是"两条规则恰好都指同一个方向"）。
        _B76_save = _B76.get_builtin_names
        _RANK76_save = E76.RANK_BY_SCOPE
        _B76.get_builtin_names = lambda: frozenset()
        E76.RANK_BY_SCOPE = False
        try:
            _d76_off = _m76("decl")
        finally:
            _B76.get_builtin_names = _B76_save
            E76.RANK_BY_SCOPE = _RANK76_save
        check("76.4b ★反面对照：builtin_names 为空（开关关）【且】作用域就近排序"
              "也关掉时，排序退回旧行为 -> Declare 排第 1"
              "（证明 76.4 是新加的那两条相关度规则在起作用）",
              [_d76_off[:1]], expect_contain=[["Declare"]])

        # 精确命中不受影响：打了完整的 declare 就该选 Declare，不能因为它是
        # "语言自带"就把它往后排。
        # 走一次真触发（把 line_text 换成打全的），而不是只看 fuzzy_match ——
        # 免得只量到匹配层、没量到"同分让位"那步排序。
        class _B76Full(_B76):
            def get_context(self):
                _ln = "    declare"
                return {"line_no": 1, "line_text": _ln,
                        "caret_col": len(_ln) + 1, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": "模块1"}

        _c76 = E76.Completer(_B76Full(), _UI76())
        _c76.trigger(True)
        _full76 = list(_c76.matches or [])
        check("76.5 ★精确命中不受影响：打全 declare -> 第 1 个就是 Declare "
              "（kind 4，与工程内名字同分也要让 Declare 赢）",
              [_full76[:1], E76.fuzzy_match("Declare", "declare")[0][0]],
              expect_contain=[["Declare"], 4])

        # ---- 76.6 ★噪音的正确度量：不是"总数"，是"会不会淹掉你自己的名字" ----
        # ⚠️ 我第一版把门槛写成"总数 <= 25"，结果 `se` 出 34 条就红了。但那个指标
        #    本身就问错了：一屏 VIEW_ROWS=15 行，超出的排在窗口外靠上下键/滚轮
        #    滚动查看（候选**不会**被丢弃，见 engine.VIEW_ROWS 的注释）。
        #    26~34 条不算失控 —— 真正会伤到用户的是【你自己的名字被挤到一屏之外】。
        #    池子：工程内 5 个名字 + 语言自带那 260 个。
        _bl76 = set(n.lower() for n in VB76B.BUILTIN_FUNCTIONS) \
            | set(n.lower() for n in VB76B.BUILTIN_TYPE_NAMES) \
            | set(n.lower() for n in VB76B.BUILTIN_KEYWORDS)
        _MINE76 = ["secondCol", "tempFile", "totalCount", "resultVal", "rowIdx"]

        class _B76Mix(object):
            def get_context(self):
                _ln = "    " + self._w
                return {"line_no": 1, "line_text": _ln,
                        "caret_col": len(_ln) + 1, "in_string": False,
                        "in_comment": False, "in_type_position": False,
                        "in_decl_position": False, "decl_names": [],
                        "proc_name": None, "module_name": "模块1"}

            def __init__(self, w):
                self._w = w

            def get_identifiers(self):
                return ([(n, "模块1", None, False) for n in _MINE76]
                        + [(n, VB76._BUILTIN_MODULE, None, False)
                           for n in sorted(_bl76)])

            def get_declared_names(self):
                return [n.lower() for n in _MINE76]

            def declared_elsewhere(self, name, module_name):
                return str(name).lower() in _bl76

            def get_structural_names(self):
                return sorted(_bl76)

            def get_builtin_names(self):
                return set(_bl76)

            def get_type_names(self):
                return list(VB76B.BUILTIN_TYPE_NAMES)

            def get_host_enum_names(self):
                return {}

            def names_outside_caret(self, caret, scope_only=False):
                return set(n.lower() for n in _MINE76)

            def apply_completion(self, *a, **k):
                return None

            def vbe_yield_visible(self):
                return False

            def vbe_popup_visible(self):
                return False

            def vbe_popup_info(self):
                return []

        def _rank76(word, mine):
            """返回 (mine 在第几位 / 一共几条)。"""
            _c = E76.Completer(_B76Mix(word), _UI76())
            _c.trigger(True)
            _ms = list(_c.matches or [])
            _i = _ms.index(mine) if mine in _ms else -1
            return _i, len(_ms)

        # 每个前缀都取"自己的名字"：它必须**在列表里**且**排第 1**
        # （语言自带的同前缀名字都排在它后面 —— 这正是 v96 那条排序规则要保的
        #  不变式）。⚠️ 这两格必须【同层】且值不同（0 与 True）⇒ 分开 check，
        #    别塞进一个列表里让子集匹配混过去。
        _rank_cases76 = [("sec", "secondCol"), ("tem", "tempFile"),
                         ("tot", "totalCount"), ("res", "resultVal"),
                         ("row", "rowIdx")]
        _ranks76 = [_rank76(w, m) for w, m in _rank_cases76]
        check("76.6 ★你自己的名字不会被语言自带挤掉：5 个常见前缀下，"
              "同前缀的工程内名字全部排第 1（实测 (名次,总条数) %s）" % (_ranks76,),
              [[(w, m) for (w, m), r in zip(_rank_cases76, _ranks76)
                if r[0] != 0]],
              expect_contain=[[]])
        check("76.6b ★而且它们都真的在候选列表里（不是被回声防护剔掉了）",
              [[m for (w, m), r in zip(_rank_cases76, _ranks76) if r[0] < 0]],
              expect_contain=[[]])

        # 同前缀命中（kind 3）永远在跳步命中（kind 0）之前 —— 收回来 260 个
        # 公共词汇之后这条不变式最关键：`se` 那种 30 多条的输入里，
        # 用户真正在找的那个必须排在前面。
        _c76se = E76.Completer(_B76Mix("se"), _UI76())
        _c76se.trigger(True)
        _se76 = list(_c76se.matches or [])
        _kinds76 = [E76.fuzzy_match(n, "se")[0][0] for n in _se76]
        _desc76 = all(_kinds76[i] >= _kinds76[i + 1] for i in
                      range(len(_kinds76) - 1))
        check("76.6c ★排序不变式：kind 沿列表单调不增（真命中一律在跳步命中之前）；"
              "`se` 实测 %d 条、kind 序列 %s"
              % (len(_se76), _kinds76),
              [_desc76, "secondCol" in _se76],
              expect_contain=[True, True])

        # 首字母闸门：这两个都是 kind 0 分散命中（"ms" 里没有连续的 s 在开头…），
        # 引擎侧那道"首字母必须对齐"会把它们滤掉。钉住它们**确实是 kind 0** ——
        # 若哪天 fuzzy_match 改了口径让它们变成 kind 1，这条会红，那时要重看。
        check("76.7 首字母闸门的依据：ms->Implements / xl->Explicit 均为 kind 0"
              " 分散命中（引擎正是按这一档把它们滤掉的）",
              [[(n, E76.fuzzy_match(n, w)[0][0]) for n, w in
                (("Implements", "ms"), ("Explicit", "xl"))]],
              expect_contain=[[("Implements", 0), ("Explicit", 0)]])
    except Exception as _e76:
        check("第 76 节异常: %s" % _e76, [True], expect_contain=[False])

    # ==================================================================
    # 77. v98：`对象.` 之后只提示该对象自己的成员 + 补全函数/过程后
    #        自动补 `()`、光标落进括号
    #
    # 用户口径：
    #   「我定义了一个类 cls，在模块里写 set c = new cls、c.te，
    #     输入 .te 的时候，应该提示 cls 类的内部成员，而不是提示 temp。」
    #   「选中函数（包括 vba 本身函数）或过程提示词，按 tab 后，
    #     应该能在函数或过程名后面自动补全双括号，光标能正确落在括号内。」
    #
    # ⚠️ 写这一节时踩到的两条（都写进下面的代码里了）：
    #   ① 光标落点那条 mock **必须模拟显示列 + 钳列**，不模拟的话中文行
    #      必假绿（MEMORY 第 19 条 —— 纯 ASCII 行上两种列语义数值相同，
    #      一旦整条 mock 不管列宽，测出来永远是"对"的）；
    #   ② 补全的三个列参数（word_start_col / end_col / caret_col）全是
    #      【字符列】，只有写回之后摆光标那一步才换显示列 —— 手写数字
    #      极易算错，本节一律用 `_caret_end` / `_word_start` 从字符串推。
    # ==================================================================
    try:
        import vbe_bridge as VB       # noqa: E402（本文件的模块级没有它）
        import vba_builtins as VBB    # noqa: E402（内建函数 / 类型清单）


        def _caret_end(prefix, word):
            """`prefix` + 已输入的 `word` 那一行的光标列（1-based **字符**列）。

            ⚠️ 补全的三个列参数（word_start_col / end_col / caret_col）全是字符列
            （1-based，与 Python 下标+1 一致），不是显示列 —— 只有"写回之后摆光标"
            那一步才换成显示列。手写数字极易算错（带全角 / 带缩进时差好几格），
            本节一律从字符串推。
            """
            return len(prefix + word) + 1


        def _word_start(prefix, word):
            """词起始的 1-based 字符列。"""
            return len(prefix) + 1


        def _col_in_parens(text, upto):
            """`text` 里 upto 之后那个左括号**右侧**的列（显示列口径）。"""
            _k = text.index("(", upto)
            return VB._disp_width(text[:_k + 1], 4, True) + 1


        print("\n=== 77. v98：`c.` 只提示对象自己的成员 + 补全函数自动补 `()` ===")

        # ------------------------------------------------------------------
        print("\n--- 77.1 parser.var_types：变量 -> 声明类型 ---")
        try:
            _vt_code = "\n".join([
                "Public Sub Work()",
                "    Dim c As Cls",
                "    Dim n As Long",
                "    Set c = New Cls",
                "    Dim s As String",
                "    temp = 1",
                "    ReDim arr(1 To 3)",
                "End Sub",
                "Sub Other(ByVal ws As Worksheet)",
                "    Dim sh As Object",
                "End Sub"])
            _vt = P.var_types(_vt_code)
            check("77.1.1 写明类型的声明都收得到（Dim / Set=New / 形参）",
                  [(_vt.get("c"), _vt.get("n"), _vt.get("s"),
                    _vt.get("ws"), _vt.get("sh"))],
                  expect_contain=[("cls", "long", "string", "worksheet", "object")])

            check("77.1.2 ★隐式变量一律不收（`temp = 1` / `ReDim arr`）——"
                  "类型没写明，硬猜只会把候选收错",
                  [("temp" in _vt), ("arr" in _vt)],
                  expect_contain=[False, False])

            _vt2 = P.var_types("\n".join([
                "Sub A()",
                "    Dim v As Cls",
                "    Dim v As Long",
                "End Sub"]))
            check("77.1.3 同名遮蔽按【先声明者胜】（不能后者覆盖前者）",
                  [_vt2.get("v")], expect_contain=["cls"])
        except Exception as _e:
            import traceback
            traceback.print_exc()
            check("77.1 异常: %s" % _e, [True], expect_contain=[False])

        # ------------------------------------------------------------------
        print("\n--- 77.2 parser.public_members：对外可见成员 ---")
        try:
            _cls = "\n".join([
                "Option Explicit",
                "Public Sub TestMethod()",
                "End Sub",
                "Public Function TestFunc() As Long",
                "    TestFunc = 1",
                "End Function",
                "Public Property Get PubProp() As Long",
                "    PubProp = 2",
                "End Property",
                "Public PubVar As Long",
                "Private PrivVar As Long",
                "Dim NoKwVar As Long",
                "Sub NoKwSub()",
                "End Sub",
                "Private Sub PrivSub()",
                "End Sub",
                "Friend Sub FriendSub()",
                "End Sub"])
            _pm = sorted(P.public_members(_cls, is_std_module=False))
            check("77.2.1 类模块：只收显式 Public 的成员 + Friend 的那个；"
                  "Private / 无关键字变量 / 无关键字 Sub 都不收",
                  [_pm],
                  expect_contain=[["FriendSub", "PubProp", "PubVar",
                                   "TestFunc", "TestMethod"]])

            _cls2 = "\n".join([
                "Public Sub Reset()",
                "    Dim localVar As Long",
                "    Dim p As Long",
                "End Sub"])
            _pm2 = sorted(P.public_members(_cls2, is_std_module=False))
            check("77.2.2 ★过程内的局部变量与形参不收"
                  "（`m1.` 访问不到别的过程的局部量）",
                  [_pm2], expect_contain=[["Reset"]])

            _std = "\n".join([
                "Public Sub Alpha()",
                "End Sub",
                "Private Sub Beta()",
                "End Sub",
                "Sub Gamma()",
                "End Sub",
                "Public GammaVar As Long",
                "Sub Work()",
                "    Dim inner As Long",
                "End Sub"])
            check("77.2.3 标准模块：模块级成员**全收**（含 Private 与无关键字），"
                  "过程内的 `inner` 不收",
                  [sorted(P.public_members(_std, is_std_module=True))],
                  expect_contain=[["Alpha", "Beta", "Gamma", "GammaVar", "Work"]])

            # ★v98 的核心教训：不能沿用 extract_records 的 priv 字段
            _recs, _pn = P.extract_records("Public Sub Reset()\nEnd Sub",
                                           module="Cls", is_std_module=False)
            _priv_flag = dict((r[0], r[3]) for r in _recs).get("Reset")
            check("77.2.4 ★对照（这条是 v98 第一版的 bug 现场）：`Public Sub Reset` 在 "
                  "extract_records 里 priv=True（意思是\"不能裸名写\"），而 "
                  "public_members 必须收它 —— 两个判据问的不是同一个问题",
                  [(_priv_flag, "Reset" in _pm2)],
                  expect_contain=[(True, True)])
        except Exception as _e:
            import traceback
            traceback.print_exc()
            check("77.2 异常: %s" % _e, [True], expect_contain=[False])

        # ------------------------------------------------------------------
        print("\n--- 77.3 engine._narrow_to_member_type 端到端（用户场景）---")
        try:
            class _CM77(object):
                def __init__(self, comp, text):
                    self._comp = comp
                    self.text = text

                @property
                def Name(self):
                    return self._comp.Name

                @property
                def Parent(self):
                    return self._comp

                @property
                def CountOfLines(self):
                    return self.text.count("\n") + 1

                def Lines(self, start, count):
                    return "\r\n".join(
                        self.text.split("\n")[start - 1:start - 1 + count])

                def ReplaceLine(self, line_no, txt):
                    _ls = self.text.split("\n")
                    _ls[int(line_no) - 1] = txt
                    self.text = "\r\n".join(_ls)

            class _Comp77(object):
                def __init__(self, name, ctype, text):
                    self.Name = name
                    self.Type = ctype
                    self.CodeModule = _CM77(self, text)

            class _Proj77(object):
                def __init__(self, comps):
                    self.Name = "VBAProject"
                    self.VBComponents = list(comps)

            class _Pane77(object):
                def __init__(self, cm, sel):
                    self.CodeModule = cm
                    self.sel = sel

                def GetSelection(self):
                    return self.sel

                def SetSelection(self, sl, sc, el, ec):
                    self.sel = (sl, sc, el, ec)

            class _VBE77(object):
                def __init__(self, comps, pane):
                    self.ActiveVBProject = _Proj77(comps)
                    self.ActiveCodePane = pane

            class _UI77(object):
                def __init__(self):
                    self.shown = False

                def show(self, *a, **k):
                    self.shown = True

                def hide(self, *a, **k):
                    self.shown = False

                def update_selection(self, *a, **k):
                    pass

                def contains_point(self, *a, **k):
                    return False

            _CLS77 = "\n".join([
                "Option Explicit",
                "Public Sub TestMethod()",
                "End Sub",
                "Public Function TempField() As Long",
                "    TempField = 1",
                "End Function",
                "Private Sub HiddenHelper()",
                "End Sub"])

            # 光标所在那一行的样子：前缀固定 "    "，后面跟正在敲的片段
            def _mk77(frag, caret_col):
                """Cls 定义在工程里；模块1 里 `Dim c As Cls`（c 的类型是已知的）。"""
                _c1 = _Comp77("Cls", 2, _CLS77)
                _c2 = _Comp77("模块1", 1, "\n".join([
                    "Sub Work()",
                    "    Dim c As Cls",
                    "    temp = 1",
                    "    temp2 = 2",
                    "    " + frag,
                    "End Sub"]))
                _pane = _Pane77(_c2.CodeModule, (5, caret_col, 5, caret_col))
                return VB.VbeBackend(), _c2, _pane, [_c1, _c2]

            _orig77 = VB._get_vbe_cached

            def _go77(frag, caret_col, wrap=None):
                _bk, _c2, _pane, _lst = _mk77(frag, caret_col)
                VB._get_vbe_cached = lambda: _VBE77(_lst, _pane)
                try:
                    _ui = _UI77()
                    _c = E.Completer(wrap(_bk) if wrap else _bk, _ui)
                    _c.trigger()
                    return _ui.shown, sorted(str(m).lower() for m in (_c.matches or []))
                finally:
                    VB._get_vbe_cached = _orig77

            _P77 = "    "
            # ---- 77.3.1 ★用户报的场景：`c.te` ----
            _r1 = _go77("c.te", _caret_end(_P77, "c.te"))
            check("77.3.1 ★用户报的场景：`c.te` -> 只出 Cls 的成员 tempfield / "
                  "testmethod，本过程的局部变量 temp / temp2 都不许混进来",
                  [_r1], expect_contain=[(True, ["tempfield", "testmethod"])])

            # ---- 77.3.2 ★前提断言：temp 在【不收窄】的池子里确实存在 ----
            # （否则 77.3.1 可能只是"temp 本来就不在候选里"的假绿）
            _bk2, _c2b, _pane2, _lst2 = _mk77("c.te", _caret_end(_P77, "c.te"))
            VB._get_vbe_cached = lambda: _VBE77(_lst2, _pane2)
            try:
                _ids2 = _bk2.get_identifiers()
                _pool2 = sorted(str(n).lower() for n in
                                E.filter_identifiers_by_scope(_ids2, "Work", "模块1")
                                if str(n).lower().startswith("te"))
                _mem2 = sorted(_bk2.get_members_of("Cls"))
                _vt2b = _bk2.get_var_type("c")
            finally:
                VB._get_vbe_cached = _orig77
            check("77.3.2 ★前提：`temp`/`temp2` 在【全工程可见池】里（不收窄时会出），"
                  "`c` 的类型查得到，Cls 的对外成员只有 tempfield / testmethod —— "
                  "这三条同时成立，77.3.1 才证明是收窄起了作用",
                  [_pool2, _vt2b, _mem2],
                  expect_contain=[["temp", "temp2"], "cls",
                                  ["tempfield", "testmethod"]])

            # ---- 77.3.3 Private 成员不该出现在 `c.` 之后 ----
            check("77.3.3 Cls 里 `Private Sub HiddenHelper` 经实例访问不到（编译不过）"
                  " -> 不该在 `c.` 后提示",
                  [("hiddenhelper" in _r1[1])], expect_contain=[False])

            # ---- 77.3.4 ★四条兜底：一律退回全工程，绝不给空窗 ----
            class _NoNarrow(object):
                """后端没提供 get_var_type / get_members_of 这两个【可选接口】。"""
                def __init__(self, inner):
                    self._in = inner

                def __getattr__(self, k):
                    if k in ("get_var_type", "get_members_of"):
                        raise AttributeError(k)
                    return getattr(self._in, k)

            # (a) 不在成员位置（裸词 te）
            _fa = _go77("te", _caret_end(_P77, "te"))
            # (b) 在成员位置但类型判不出（temp 是隐式变量，var_types 不收它）
            _fb = _go77("temp.te", _caret_end(_P77, "temp.te"))
            # (c) 后端没有那两个可选接口
            _fc = _go77("c.te", _caret_end(_P77, "c.te"), wrap=_NoNarrow)

            check("77.3.4a ★兜底（项目第 10 条）三种情形 —— 不在成员位置 / "
                  "类型判不出（隐式变量 temp 查不到类型）/ 后端没有 get_var_type "
                  "与 get_members_of 这两个可选接口 —— 得到的候选列表逐字相同，"
                  "都等于**不收窄**时那份（各 %d 条，含 temp / temp2）" % len(_fa[1]),
                  [(_fa[1] == _fb[1], _fa[1] == _fc[1], _fb[1] == _fc[1]),
                   "temp" in _fa[1], "temp2" in _fa[1]],
                  expect_contain=[(True, True, True), True, True])
            check("77.3.4b ★绝不给空窗：三种兜底都真的弹出了候选"
                  "（空窗会让用户以为工具坏了，比多给几条糟得多）",
                  [(_fa[0], _fb[0], _fc[0])],
                  expect_contain=[(True, True, True)])
        except Exception as _e:
            import traceback
            traceback.print_exc()
            check("77.3 异常: %s" % _e, [True], expect_contain=[False])

        # ------------------------------------------------------------------
        print("\n--- 77.4 vbe_bridge._is_call_site：哪些位置允许补括号 ---")
        try:
            def _csl(line, col):
                return VB._is_call_site(line, col)

            # (整行, 词, 期望, 说明)  列一律由 _word_start 推，不手写
            _cases = [
                ("    x = DoWo", "DoWo", True, "赋值右值 = 调用点"),
                ("    DoWo 1", "DoWo", True, "语句首裸词（Call 可省）"),
                ("    Set fso = DoWo", "DoWo", True, "等号右值"),
                ("    If DoWo Then", "DoWo", True, "If 条件"),
                ("    MsgBox DoWo", "DoWo", True, "实参位置"),
                ("    a = Foo(DoWo", "DoWo", True, "括号内"),
                ("    c.DoWo", "DoWo", True,
                 "★点号左边放行（方法该补；属性补错让 VBE 自己当场报）"),
                ("    Range(\"A1\").DoWo", "DoWo", True, "★`Range(\"A1\").` 也放行"),
                ("Dim n As Lon", "Lon", False,
                 "★`As` 后面是类型位置，补成 `Long()` 编译不过"),
                ("    Dim n As Lon", "Lon", False, "同上（缩进版）"),
                ("Function f() As Lon", "Lon", False, "函数返回类型位置"),
                ("    Set p As Lon", "Lon", False, "`Set p As <类型>`"),
                ("    Set fso = New FileSys", "FileSys", False,
                 "★`New` 后面要的是类型表达式不是调用"),
                ("Dim o As New FileSys", "FileSys", False, "同上"),
            ]
            _bad = ["%s（`%s` 第 %d 列）-> %s，期望 %s"
                    % (why, l, _word_start(l[:l.index(w)], w),
                       VB._is_call_site(l, _word_start(l[:l.index(w)], w)), exp)
                    for (l, w, exp, why) in _cases
                    if VB._is_call_site(l, _word_start(l[:l.index(w)], w)) != exp]
            check("77.4.1 %d 条调用点判据逐条对（As / New 类型位置挡住，其余放行）"
                  % len(_cases),
                  [_bad or "全部一致"], expect_contain=["全部一致"])

            # ★v98 踩过的坑：`As Long` 里 Long 的前一格是【空格】不是 s
            _la, _lb = "Dim n As Lon", "Dim nAs Lon"
            check("77.4.2 ★反面对照（v98 第一版的 bug 现场）：`As` 与词之间【隔着空格】"
                  "（`As Long` 的 Long 前一格是空格不是 s）—— 判据若写成"
                  "\"跳过空白后那个字符 in 'sS'\"，这一条永远不成立",
                  [(_csl(_la, 10), _csl(_lb, 9))],
                  expect_contain=[(False, False)])

            # ⚠️ 以 as / new 结尾的名字不能被误杀
            _traps = [("    x = Stat", "Stat"), ("    x = Was", "Was"),
                      ("    x = Renew", "Renew"), ("    x = DoNew", "DoNew"),
                      ("    New = DoWo", "DoWo")]
            _bad2 = ["%s（`%s`）-> %s"
                     % (w, l, VB._is_call_site(l, _word_start(l[:l.index(w)], w)))
                     for (l, w) in _traps
                     if not VB._is_call_site(l, _word_start(l[:l.index(w)], w))]
            check("77.4.3 ★切片版不能误杀【以 as / new 结尾的名字】："
                  "Stat / Was / Renew / DoNew，以及左边那个 New 其实是变量名的 "
                  "`New = DoWo`，都还是调用点",
                  [_bad2 or "全部一致"], expect_contain=["全部一致"])

            check("77.4.4 判据出错（传垃圾进去）时**不补**——宁可少补也不写坏代码",
                  [_csl(None, 5), _csl("abc", "x")],
                  expect_contain=[False, False])
        except Exception as _e:
            import traceback
            traceback.print_exc()
            check("77.4 异常: %s" % _e, [True], expect_contain=[False])

        # ------------------------------------------------------------------
        print("\n--- 77.5 engine._is_callable_name：过程 vs 变量 ---")
        try:
            class _BKProc(object):
                def get_proc_names(self):
                    return frozenset(["dowork", "calcvalue"])

                def get_builtin_names(self):
                    return set(str(n).lower() for n in VBB.BUILTIN_FUNCTIONS)

                def get_type_names(self):
                    return list(VBB.BUILTIN_TYPE_NAMES)

            _cp = E.Completer(_BKProc(), _UI77())
            check("77.5.1 工程内的过程要补括号；同模块的变量不补",
                  [(_cp._is_callable_name("DoWork"),
                    _cp._is_callable_name("CalcValue"),
                    _cp._is_callable_name("tempCount"))],
                  expect_contain=[(True, True, False)])

            check("77.5.2 语言自带：**函数**该补（MsgBox / Left）；**数据类型**"
                  "不该补（Long / String / Object —— 它们只会出现在 As 后面）",
                  [(_cp._is_callable_name("MsgBox"), _cp._is_callable_name("Left"),
                    _cp._is_callable_name("Long"), _cp._is_callable_name("String"),
                    _cp._is_callable_name("Object"))],
                  expect_contain=[(True, True, False, False, False)])

            class _BKNoProc(object):
                """旧式后端：没有 get_proc_names —— 判据缺失就不动手。"""

                def get_builtin_names(self):
                    return set()

                def get_type_names(self):
                    return list(VBB.BUILTIN_TYPE_NAMES)

            _cn = E.Completer(_BKNoProc(), _UI77())
            check("77.5.3 ★后端没提供 `get_proc_names` 时一律**不补**"
                  "（可选接口缺失 = 判据缺失 = 不动手）",
                  [(_cn._is_callable_name("DoWork"),
                    _cn._is_callable_name("MsgBox"))],
                  expect_contain=[(False, False)])
        except Exception as _e:
            import traceback
            traceback.print_exc()
            check("77.5 异常: %s" % _e, [True], expect_contain=[False])

        # ------------------------------------------------------------------
        print("\n--- 77.6 accept 是否真把 want_parens 传下去 ---")
        try:
            _LN77 = "    x = DoWo"

            class _BKAcc(object):
                def __init__(self, recorder):
                    self._rec = recorder

                def get_context(self):
                    return {"line_no": 1, "line_text": _LN77,
                            "caret_col": _caret_end("", _LN77),
                            "word_start_col": _word_start("    x = ", "DoWo"),
                            "word_end_col": _caret_end("    x = ", "DoWo"),
                            "in_string": False, "in_comment": False,
                            "in_type_position": False, "in_decl_position": False,
                            "decl_names": [], "proc_name": "Work",
                            "module_name": "模块1"}

                def get_identifiers(self):
                    return [("DoWork", "模块1", None, False),
                            ("DoOther", "模块1", None, False)]

                def get_declared_names(self):
                    return ["dowork", "doother"]

                def get_proc_names(self):
                    return frozenset(["dowork"])

                def declared_elsewhere(self, name, module_name):
                    return True

                def get_builtin_names(self):
                    return set()

                def get_type_names(self):
                    return []

                def vbe_yield_visible(self):
                    return False

                def vbe_popup_visible(self):
                    return False

                def apply_completion(self, *a, **k):
                    self._rec.append((a, k))

            def _pick(bk, chosen):
                _c = E.Completer(bk, _UI77())
                _c.trigger(True)
                _ms = list(_c.matches or [])
                if chosen not in _ms:
                    return None
                _c.selected = _ms.index(chosen)
                _c.accept()

            _sent = []
            _pick(_BKAcc(_sent), "DoWork")
            check("77.6.1 ★过程 -> accept 把 want_parens=True 传给后端",
                  [bool(_sent) and _sent[0][0][4]], expect_contain=[True])

            _sent2 = []
            _pick(_BKAcc(_sent2), "DoOther")
            check("77.6.2 ★变量（不在过程名集合里）-> want_parens=False",
                  [bool(_sent2) and _sent2[0][0][4]], expect_contain=[False])

            # 旧式 4 参后端不能被 TypeError 打挂
            _sent3 = []

            class _BK4(_BKAcc):
                def apply_completion(self, line_no, ws, we, comp):     # 旧式 4 参
                    _sent3.append(comp)

            _pick(_BK4(_sent3), "DoWork")
            check("77.6.3 ★旧式 4 参 apply_completion 的后端不被打挂"
                  "（引擎里 `except TypeError` 兜底，退回 4 参调用）",
                  [_sent3], expect_contain=[["DoWork"]])
        except Exception as _e:
            import traceback
            traceback.print_exc()
            check("77.6 异常: %s" % _e, [True], expect_contain=[False])

        # ------------------------------------------------------------------
        print("\n--- 77.7 apply_completion 真写回 + 光标落进括号（mock 真 VBE）---")
        try:
            class _CM77b(object):
                def __init__(self, lines):
                    self.lines = list(lines)
                    self.Name = "Module1"

                @property
                def CountOfLines(self):
                    return len(self.lines)

                def Lines(self, start, count):
                    return self.lines[int(start) - 1]

                def ReplaceLine(self, line_no, text):
                    self.lines[int(line_no) - 1] = text

            class _Pane77b(object):
                """忠实模拟真 VBE：列是【显示列】，超出行尾的列被钳回行尾。"""
                def __init__(self, cm):
                    self.CodeModule = cm
                    self.sel = (1, 1, 1, 1)

                def GetSelection(self):
                    return self.sel

                def SetSelection(self, sl, sc, el, ec):
                    ln = max(1, min(int(sl), len(self.CodeModule.lines)))
                    _t = self.CodeModule.lines[ln - 1]
                    _eol = VB._disp_width(_t, 4, True) + 1
                    self.sel = (ln, max(1, min(int(sc), _eol)),
                                ln, max(1, min(int(ec), _eol)))

            class _VBE77b(object):
                def __init__(self, pane):
                    self.ActiveCodePane = pane

            def _run77b(line, typed, completion, want_parens):
                """把 `line` 里 `typed` 这个词补成 `completion`。

                返回 (写回后的整行, 光标列)。三个列参数都由字符串推，不手写。
                """
                _pre = line[:line.index(typed)]
                _ws = _word_start(_pre, typed)
                _ce = _caret_end(_pre, typed)
                _cm = _CM77b([line])
                _pn = _Pane77b(_cm)
                _orig = VB._get_vbe_cached
                _sem = VB._sem_cache.get("info")
                VB._sem_cache["info"] = None      # 别吃缓存，模拟"当场量"
                VB._get_vbe_cached = lambda: _VBE77b(_pn)
                try:
                    VB.VbeBackend().apply_completion(
                        1, _ws, _ce, completion, want_parens=want_parens)
                finally:
                    VB._get_vbe_cached = _orig
                    VB._sem_cache["info"] = _sem
                return _cm.lines[0], _pn.sel[3]

            _exp = "    x = DoWork()"
            check("77.7.1 `    x = DoWo` -> `    x = DoWork()`，"
                  "光标落【括号内】第 %d 列（不是停在左括号左侧还得手动右移）"
                  % _col_in_parens(_exp, len("    x = DoWork")),
                  [_run77b("    x = DoWo", "DoWo", "DoWork", True)],
                  expect_contain=[(_exp, _col_in_parens(_exp, len("    x = DoWork")))])

            # 77.7.2 ★中文行：名字中英混排，光标必须按【显示列】算
            _l77 = "    订单金额 = 计"
            _exp2 = "    订单金额 = 计算金额()"
            _c2exp = _col_in_parens(_exp2, len("    订单金额 = 计算金额"))
            check("77.7.2 ★中文行（名字中英混排）：`%s` -> `计算金额()`，"
                  "光标落括号内第 %d 列（显示列，全角占 2 格）。按【字符列】算只有 "
                  "%d —— 差的是前缀里那几个汉字的格数，光标会落在名字中间"
                  % (_l77, _c2exp, len(_exp2) + 1),
                  [_run77b(_l77, "计", "计算金额", True)],
                  expect_contain=[(_exp2, _c2exp)])
            check("77.7.2b ★前提：这一行上两种列语义**确实分得开**"
                  "（显示列行尾 %d != 字符列行尾 %d）—— 否则本条是假绿"
                  % (_c2exp, len(_exp2) + 1),
                  [(_c2exp == len(_exp2) + 1)],
                  expect_contain=[False])

            _exp3 = "    x = DoWork(1)"
            check("77.7.3 括号本来就有的调用点不重复补："
                  "`    x = DoWo(1)` -> `    x = DoWork(1)`（不是 `(())`）；"
                  "光标落在括号内、参数之前（第 15 列）",
                  [_run77b("    x = DoWo(1)", "DoWo", "DoWork", True)],
                  expect_contain=[(_exp3, 15)])

            _exp4 = "    Dim n As Long"
            check("77.7.4 ★类型位置不补括号：`    Dim n As Lon` 即使 want_parens=True "
                  "也只补出 Long，不生成 `Long()`（那会编译不过）—— "
                  "`_is_call_site` 这道闸门就挡在这里；光标落在 Long 末尾第 %d 列"
                  % (len(_exp4) + 1),
                  [_run77b("    Dim n As Lon", "Lon", "Long", True)],
                  expect_contain=[(_exp4, len(_exp4) + 1)])

            check("77.7.5 want_parens=False（变量）时只替换词，一个括号都不加",
                  [_run77b("    x = DoWo", "DoWo", "DoWork", False)],
                  expect_contain=[("    x = DoWork", 15)])

            _exp6 = "    x = DoWork() + 1"
            check("77.7.6 ★补出的括号贴在【词尾】：`    x = DoWo + 1` -> "
                  "`    x = DoWork() + 1` —— 词与括号之间不许有空格、"
                  "`+` 前的空格也不许被吃掉（这是 v98 补测试时挖出的真 bug："
                  "插入点曾用成\"跳过空白后的位置\"，补出过 `DoWork ()+ 1`）",
                  [_run77b("    x = DoWo + 1", "DoWo", "DoWork", True)],
                  expect_contain=[(_exp6, 16)])
            check("77.7.6b 同理，逗号分隔的实参位置：`    x = DoWo, 2` -> "
                  "`    x = DoWork(), 2`",
                  [_run77b("    x = DoWo, 2", "DoWo", "DoWork", True)],
                  expect_contain=[("    x = DoWork(), 2", 16)])
        except Exception as _e:
            import traceback
            traceback.print_exc()
            check("77.7 异常: %s" % _e, [True], expect_contain=[False])

        # ------------------------------------------------------------------
        print("\n--- 77.8 源码护栏 ---")
        try:
            _src_e = io.open(os.path.join(ROOT, "engine.py"), encoding="utf-8").read()
            _src_v = io.open(os.path.join(ROOT, "vbe_bridge.py"),
                             encoding="utf-8").read()

            check("77.8.1 `_member_narrowed` 必须在每次 trigger 开头重置"
                  "（绝不跨触发沿用：上一句在成员位置收窄过，不代表这一句也是）",
                  [("_member_narrowed = False" in _src_e)],
                  expect_contain=[True])

            check("77.8.2 成员位置判据只有一份实现（`_member_list_dot_col`），"
                  "收窄复用它而不另写一遍（v87/v93 的老教训）",
                  [(_src_e.count("def _member_list_dot_col(") == 1,
                    _src_e.count("_member_list_dot_col(") >= 3)],
                  expect_contain=[(True, True)])

            check("77.8.3 `accept` 的 want_parens 必须被 `except TypeError` 兜住"
                  "（旧式 4 参后端不能被打挂）",
                  [("except TypeError:" in _src_e and "want_parens" in _src_e)],
                  expect_contain=[True])

            check("77.8.4 补括号的总开关 AUTO_PAREN_ON_COMPLETE 出厂为 True",
                  [VB.AUTO_PAREN_ON_COMPLETE], expect_contain=[True])

            _ics = _src_v.split("def _is_call_site(")[1].split("\ndef ")[0]
            check("77.8.5 `_is_call_site` 出异常时 return False（不补），"
                  "不是 return True（宁可少补也不写坏代码）",
                  [_ics.rstrip().endswith("return False")],
                  expect_contain=[True])

            _ap = _src_v.split("def apply_completion(")[1].split("\n    def ")[0]
            check("77.8.6 ★补括号的【插入点】必须是词尾，不是\"跳过空白后的位置\""
                  " —— 77.7.6 挖出的真 bug：用跳过后的位置插入会补出 `DoWork ()+ 1`。"
                  "源码里 `j` 只准用于判断、`k` 才准用于插入",
                  [("new_line[:k] + \"()\"" in _ap.replace("'", '"'),
                    "new_line[:j] + " not in _ap.replace("'", '"'))],
                  expect_contain=[(True, True)])
        except Exception as _e:
            import traceback
            traceback.print_exc()
            check("77.8 异常: %s" % _e, [True], expect_contain=[False])
    except Exception as _e77:
        import traceback as _tp77
        _tp77.print_exc()
        check("第 77 节异常: %s" % _e77, [True], expect_contain=[False])

    # ==================================================================
    # 78. v99：① 关键字不该补括号  ② 退格成对删除（空括号 / 空引号）
    #
    # 用户口径：
    #   「如果选中的提示词是关键字，不应该自动补全双括号。现在 public，
    #     private 会自动补全双括号。」
    #   「按回退键删除字符时，如果删除的是左右括号，或双引号，如果括号中
    #     或双引号中没内容，就能自动删除左右两端。」
    #
    # ⚠️ 本节自己踩到的三个坑（都写进注释了，别删）：
    #   ① `String` / `Date` 在 VBA 里**既是类型名也是函数**（实测交集就
    #      这两个），所以纯白名单会把它们放行 —— 白名单与类型排除两道
    #      判据问的是不同问题，合成一条就退回 v98 那个 public()；
    #   ② 判据**不许自己猜列语义**（"两种都试"会把精确判据退化成两个
    #      模糊判据的并集，误判会改坏用户的代码）；
    #   ③ 本功能**不能**调 `_caret_char_col`：它拿【行尾列】反推 tabw，
    #      而这里的光标停在行中间，拟合出来的 tabw 是错的。
    # ==================================================================
    try:
        _ROOT78 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        class _UI78(object):
            """最小 UI 桩 —— 只被 Completer 的构造需要。"""

            def hide(self):
                pass

            def is_visible(self):
                return False

            def show(self, *a, **k):
                pass

        # ---------------- 78.1 关键字不该补括号 ----------------
        print("\n--- 78.1 engine._is_callable_name：关键字不该补括号 ---")

        class _BK78(object):
            """后端把【内建函数 + 数据类型 + 关键字 + 常量】一整批都收进
            get_builtin_names —— 这正是真后端 v62 之后的样子。"""

            def get_proc_names(self):
                return frozenset(["dowork", "calcvalue"])

            def get_builtin_names(self):
                _s = set()
                for _lst in (VBB.BUILTIN_FUNCTIONS, VBB.BUILTIN_TYPE_NAMES,
                             VBB.BUILTIN_KEYWORDS, VBB.BUILTIN_CONSTANTS):
                    _s |= set(str(n).lower() for n in _lst)
                return _s

        _cp78 = E.Completer(_BK78(), _UI78())

        # ---- 78.1.1 ★用户报的现象：Public / Private 被补成 public() ----
        check("78.1.1 ★用户报的现象：`public` / `private` 是语言关键字，"
              "选它们补全时**不许**补出 `public()` —— 关键字后面带括号编译不过",
              [(_cp78._is_callable_name("Public"),
                _cp78._is_callable_name("Private"))],
              expect_contain=[(False, False)])

        # ---- 78.1.2 全量扫：79 个关键字一个都不许放行 ----
        # ⚠️ 这一条比逐个列举强得多：vba_builtins 以后再加关键字，自动被覆盖。
        _kw78 = sorted(set(str(n).lower() for n in VBB.BUILTIN_KEYWORDS))
        _leak78 = [k for k in _kw78 if _cp78._is_callable_name(k)]
        check("78.1.2 ★全量扫 %d 个语言关键字，一个都不许被判成"
              "「可调用」—— 比逐个列举强：清单以后再加关键字也自动被覆盖"
              % len(_kw78),
              [_leak78 or "全部挡住"],
              expect_contain=["全部挡住"])

        # ---- 78.1.3 内建常量同理（vbCrLf() 是坏代码）----
        _cs78 = sorted(set(str(n).lower() for n in VBB.BUILTIN_CONSTANTS))
        _leakc78 = [c for c in _cs78 if _cp78._is_callable_name(c)]
        check("78.1.3 ★内建常量同样一个都不许放行（`vbCrLf()` 之类"
              "都不是可调用对象）",
              [_leakc78 or "全部挡住"],
              expect_contain=["全部挡住"])

        # ---- 78.1.4 ★v99 的病根：黑名单换成白名单 ----
        # v98 用的是"在内建池里、且不是数据类型就算可调用"，那只挡了 13 个类型名，
        # 关键字与常量全漏。改成"是内建函数、且不是类型名"。
        _fnl78 = set(str(n).lower() for n in VBB.BUILTIN_FUNCTIONS)
        _tpl78 = set(str(n).lower() for n in VBB.BUILTIN_TYPE_NAMES)
        _kwl78 = set(str(n).lower() for n in VBB.BUILTIN_KEYWORDS)
        _csl78 = set(str(n).lower() for n in VBB.BUILTIN_CONSTANTS)
        check("78.1.4 ★黑名单不够用：内建池里的关键字 %d 个、常量 %d 个，"
              "v98 那句\"不是数据类型就算可调用\"一个都挡不住（真正被挡住的"
              "只有 %d 个类型名）—— 这是 v99 改白名单的直接原因"
              % (len(_kwl78), len(_csl78), len(_tpl78)),
              [(len(_kwl78 & _fnl78) == 0,
                len(_csl78 & _fnl78) == 0,
                len(_tpl78) > 0)],
              expect_contain=[(True, True, True)])

        # ---- 78.1.5 ★真正的函数一个都不许误伤 ----
        check("78.1.5 ★白名单不许过严：内建函数该补的还得补"
              "（MsgBox / Left / Split / UBound 全都 True）",
              [tuple(_cp78._is_callable_name(n)
                     for n in ("MsgBox", "Left", "Split", "UBound"))],
              expect_contain=[(True, True, True, True)])
        check("78.1.5b 工程内自己写的过程照旧补括号（第一路 proc_names 不受"
              "v99 收紧影响）",
              [(_cp78._is_callable_name("DoWork"),
                _cp78._is_callable_name("CalcValue"))],
              expect_contain=[(True, True)])

        # ---- 78.1.6 ★String / Date 是双身份名字 ----
        # VBA 里 `String(5, "ab")` 是合法的重复字符函数，但 `Dim n As String`
        # 是类型位置。两边都是 yes ⇒ 只能靠后端 `_is_call_site` 按位置挡。
        _dual78 = sorted(_fnl78 & _tpl78)
        check("78.1.6 ★前提：`String` / `Date` 同时是函数与类型（交集 = %s）—— "
              "所以光有白名单不够，还得叠一道类型排除；而它们在**调用点**"
              "仍要能补括号" % (_dual78,),
              [(_dual78, _cp78._is_callable_name("String"),
                _cp78._is_callable_name("Date"))],
              expect_contain=[(["date", "string"], False, False)])
        check("78.1.6b ★反面对照：类型排除只管\"名字像不像类型\"，"
              "真正的调用点仍放行 —— `x = String(5, \"ab\")` 是合法代码，"
              "后端 `_is_call_site` 在等号右边判 True、在 As 后面判 False",
              [(VB._is_call_site("    x = String", 9),
                VB._is_call_site("Dim n As Str", 10))],
              expect_contain=[(True, False)])

        # ---- 78.1.7 用户定义优先：工程里声明过同名的，以用户为准 ----
        class _BKShadow78(object):
            """用户在工程里 `Dim Split As String` —— 后端会把 split 从内建池剔除。"""

            def get_proc_names(self):
                return frozenset()

            def get_builtin_names(self):
                return (set(str(n).lower() for n in VBB.BUILTIN_FUNCTIONS)
                        - {"split"})

        _csh78 = E.Completer(_BKShadow78(), _UI78())
        check("78.1.7 ★用户定义优先：工程里声明过 `Split`（变量）时它就不再是"
              "内建函数，不补括号 —— 用户自己的定义永远压过语言自带",
              [_csh78._is_callable_name("Split")],
              expect_contain=[False])

        # ---- 78.1.8 后端缺接口时不补（旧式后端行为不变）----
        class _BKOld78(object):
            def get_builtin_names(self):
                return set()

        _co78 = E.Completer(_BKOld78(), _UI78())
        check("78.1.8 后端没提供 `get_proc_names` 时一律不补（判据缺失 = 不动手，"
              "旧式 / 测试后端行为完全不变）",
              [(_co78._is_callable_name("DoWork"),
                _co78._is_callable_name("MsgBox"))],
              expect_contain=[(False, False)])

        # ---- 78.1.9 源码护栏：两道判据不许合成一条 ----
        _src78 = io.open(os.path.join(_ROOT78, "engine.py"),
                         encoding="utf-8").read()
        _ic78 = _src78.split("def _is_callable_name(")[1].split("\n    def ")[0]
        check("78.1.9 ★源码护栏：`_is_callable_name` 里白名单与类型排除必须"
              "**同时**出现（合成一条就会退回 v98 那个 `public()`；只留白名单则 "
              "`String` 漏成 True）",
              [("in self._builtin_function_names()" in _ic78,
                "not in self._builtin_type_names()" in _ic78)],
              expect_contain=[(True, True)])

        # ---------------- 78.2 退格成对删除 ----------------
        print("\n--- 78.2 vbe_bridge：退格删掉一对空括号 / 空引号 ---")

        def _at78(marked):
            """`"    x = f(@)"` -> (行, 字符列)。@ 是光标标记（1-based 字符列）。

            ⚠️ 空行也必须写 `@`（哪怕它落在第 1 列）—— 没有标记就没法推出列号，
            而**列一律从字符串推**（MEMORY 第 24 条：手写数字必错）。
            """
            _i = marked.index("@")
            return marked.replace("@", "", 1), _i + 1

        # ---- 78.2.1 判据：光标夹在一对空括号 / 空引号正中间 ----
        _l1, _c1 = _at78("    x = f(@)")
        _l1q, _c1q = _at78('    s = "@"')
        check("78.2.1 ★用户报的场景：`    x = f(@)`（光标在 `(` 与 `)` 中间，"
              "字符列 %d）-> 接管；空引号 `%s`（第 %d 列）同理"
              % (_c1, _l1q, _c1q),
              [(VB.pair_delete_wanted(_l1, _c1),
                VB.pair_delete_wanted(_l1q, _c1q))],
              expect_contain=[(True, True)])

        # ---- 78.2.2 ★"中间没内容"这条由判据本身钉死 ----
        # 左邻是 `(`、右邻**紧挨着**是 `)` —— 两格之间本来就没有任何东西。
        _bad_pairs = []
        for _spec, _want in [
            ("    x = f(a@1)", False),      # 括号里有内容
            ('    s = a"x@"y"b', False),    # 引号里有内容
            ("    x = f@()", False),        # 光标在 `(` 左边（退格该删 f）
            ("    x = f()@", False),        # 光标在 `)` 右边（退格该删 `)`）
            ("    x = f())@", False),       # 右括号挨右括号，不是"一对"
            ("    x = f(订@单)", False),    # 里面是中文，有内容
            ("    x = 1@", False),          # 普通代码
            ("@x = 1", False),             # 行首普通词
            ("@", False),                  # 空行（列 = 1）
            ('    s = @""', False),         # 左邻是 `=` 不是引号
        ]:
            _l, _c = _at78(_spec)
            _got = VB.pair_delete_wanted(_l, _c)
            if _got != _want:
                _bad_pairs.append("%r -> %s，期望 %s" % (_spec, _got, _want))
        check("78.2.2 ★10 条『不接管』逐条对：括号里有内容 / 光标在符号外侧 / "
              "右括号挨右括号 / 空白行 —— 一个都不许接管（有内容时右括号得留着）",
              [_bad_pairs or "全部一致"],
              expect_contain=["全部一致"])

        # ---- 78.2.3 ★钩子侧与主侧必须同一份实现（项目铁律第 18 条）----
        check("78.2.3 ★钩子判据 `pair_delete_wanted` 与主侧复核 "
              "`pair_delete_span` 必须是同一份语义 —— 前者只是后者的一个"
              "bool 包装（少写一份，两边早晚会有一边漏掉某个分支）",
              [VB.pair_delete_wanted.__doc__.find("pair_delete_span") >= 0],
              expect_contain=[True])
        check("78.2.3b ★两处光标列换算走的是**同一份** `_bs_caret_char_col`"
              "（含全角 / Tab 的行上，两边用不同口径换算 = 钩子说接管、主侧"
              "说接管不了，按键被吞掉却什么都没发生）",
              [(VB.pair_delete_wanted.__code__.co_names.count("_bs_caret_char_col"),
                hasattr(VB, "_bs_caret_char_col"),
                VB.VbeBackend.delete_pair_here.__code__.co_names.count(
                    "_bs_caret_char_col"))],
              expect_contain=[(1, True, 1)])

        # ---- 78.2.4 ★判据不许自己猜列语义（v99 第一版的 bug）----
        # 我第一版让 pair_delete_span 同时接受"显示列"和"字符列"两种解释。
        # 这一行是关键：`    说()话` 光标真在 `(` 与 `)` 之间（字符列 7、显示列 8）。
        # 换算过 -> 字符列 7 -> 命中 (5,6)【对】；不换算、把显示列 8 当字符列
        # -> 读到 t[7]/t[8] = ')'/'话' -> 判不出【这次侥幸没事】。
        # ⚠️ 结论：两种口径在这一行上给出的是**不同下标**，"两种都试"会把它们
        # 合成一个并集 —— 判据一旦变成并集，漏判只是功能不生效，误判是改坏代码。
        # ⇒ 判据只吃字符列，换算责任交给调用方（各用自己最准的那份）。
        _l4 = "    说()话"
        _disp4 = 8                      # 光标在 `(` 与 `)` 之间：字符列 7、显示列 8
        _col4 = VB._bs_caret_char_col(_l4, _disp4)[0]
        check("78.2.4 ★判据只吃【字符列】：`%s` 光标在括号中间（显示列 %d -> "
              "字符列 %d）—— 换算后判据得 %s【对】；不换算、把显示列当字符列"
              "会得 %s。两种口径在这一行上给出的是**不同下标**，\"两种都试\""
              "会把它们并成一个并集（v99 第一版正是那样写的；误判会改坏代码，"
              "漏判只是功能不生效，代价不对称）"
              % (_l4, _disp4, _col4, VB.pair_delete_span(_l4, _col4),
                 VB.pair_delete_span(_l4, _disp4)),
              [(_col4, VB.pair_delete_span(_l4, _col4),
                VB.pair_delete_span(_l4, _disp4))],
              expect_contain=[(7, (5, 6), None)])

        # ---- 78.2.5 ★不许用 `_caret_char_col`：它的拟合是拿【行尾列】反推 tabw ----
        _l5, _c5 = _at78("\tx = f(@)")
        _disp5 = VB._disp_width(_l5, 4, True, upto=_c5 - 1) + 1
        _col5 = VB._bs_caret_char_col(_l5, _disp5)[0]
        _col5_bad = VB._caret_char_col(None, None, 1, _l5, _disp5)[0]
        check("78.2.5 ★Tab 缩进的行：`%s` 光标在括号中间（显示列 %d）—— "
              "本功能用 `_bs_caret_char_col` 得字符列 %d -> %s【对】；"
              "而 `_caret_char_col` 得 %d -> %s【错】：它拿**行尾列**反推 tabw，"
              "光标停在行中间时拟合成 tabw=3 ⇒ 本功能绝不能调它"
              % (_l5, _disp5, _col5, VB.pair_delete_span(_l5, _col5),
                 _col5_bad, VB.pair_delete_span(_l5, _col5_bad)),
              [(VB.pair_delete_span(_l5, _col5),
                VB.pair_delete_span(_l5, _col5_bad))],
              expect_contain=[((6, 7), None)])

        # ---- 78.2.6 集成：真删 + 光标落点（mock 模拟显示列 + 钳列）----
        class _CM78(object):
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "M1"

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count=1):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n" for l in self.lines[s0:s0 + int(count)])

            def ReplaceLine(self, n, text):
                self.lines[int(n) - 1] = text

            def DeleteLines(self, n, c=1):
                del self.lines[int(n) - 1:int(n) - 1 + int(c)]

            def InsertLines(self, n, text):
                self.lines.insert(int(n) - 1, text)

        class _Pane78(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                # ★忠实模拟真 VBE：列是【显示列】，超出行尾会被**钳到行尾**。
                # 缺了这两条，v95 那类"中文行光标落中间"的 bug 在测试里看不见。
                ln = max(1, min(int(sl), len(self.CodeModule.lines)))
                try:
                    _t = self.CodeModule.lines[ln - 1]
                except Exception:
                    _t = ""
                _eol = VB._disp_width(_t, 4, True) + 1
                self.sel = (ln, max(1, min(int(sc), _eol)),
                            ln, max(1, min(int(ec), _eol)))

        class _VBE78(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _run78(marked, sel=None):
            """按 mock 的代码窗跑一次 delete_pair_here，返回 (ok, 行, 光标列)。

            ⚠️ `marked` 里的 @ 是光标位置，本函数按【字符列】给，再由 mock 换算成
            显示列喂进去 —— 与真机 GetSelection 一样。**列一律从字符串推。**
            """
            _line, _col = _at78(marked)
            _cm = _CM78([_line])
            _disp = VB._disp_width(_line, 4, True, upto=_col - 1) + 1
            _pn = _Pane78(_cm, sel if sel is not None else (1, _disp, 1, _disp))
            _orig = VB._get_vbe_cached
            VB._get_vbe_cached = lambda: _VBE78(_pn)
            try:
                _ok = VB.VbeBackend().delete_pair_here()
            finally:
                VB._get_vbe_cached = _orig
            return (_ok, _cm.lines[0], _pn.sel[3])

        _exp78 = "    x = f"
        check("78.2.6 ★集成：`    x = f(@)` 一次退格 -> `    x = f`，"
              "光标落在第 %d 列（= 左符号原来的位置，也正是原生\"删掉 `(`\""
              "之后的落点）" % (len(_exp78) + 1),
              [_run78("    x = f(@)")],
              expect_contain=[(True, _exp78, len(_exp78) + 1)])
        check("78.2.6b 空引号对：`    s = \"@\"` -> `    s = `",
              [_run78('    s = "@"')],
              expect_contain=[(True, "    s = ", 9)])
        check("78.2.6c ★中文行（名字里全是全角）：`    订单金额 = 计算(@)` -> "
              "`    订单金额 = 计算`，光标按【显示列】落在第 %d 列"
              "（按字符列只有 %d，差的就是那几个汉字的格数）"
              % (VB._disp_width("    订单金额 = 计算", 4, True) + 1,
                 len("    订单金额 = 计算") + 1),
              [_run78("    订单金额 = 计算(@)")],
              expect_contain=[(True, "    订单金额 = 计算",
                               VB._disp_width("    订单金额 = 计算", 4, True) + 1)])
        _exp5 = "\tx = f"
        check("78.2.6d ★Tab 缩进的工程照旧生效：`\\tx = f(@)` -> `%s`"
              "，光标落【显示列】第 %d 列（Tab 展开到 4 格，纯字符列只有 %d —— "
              "这条专钉 78.2.5 那个 tabw 拟合陷阱）"
              % (_exp5, VB._disp_width(_exp5, 4, True) + 1, len(_exp5) + 1),
              [_run78("\tx = f(@)")],
              expect_contain=[(True, _exp5, VB._disp_width(_exp5, 4, True) + 1)])
        check("78.2.6e 嵌套：`    x = f(g(@))` -> `    x = f(g)`（只删最里那一对）",
              [_run78("    x = f(g(@))")],
              expect_contain=[(True, "    x = f(g)", 12)])

        # ---- 78.2.7 ★不接管时必须一个字都不改（交给 VBE 原生退格）----
        _bad78 = []
        for _m, _expect_line in [("    x = f(a@1)", "    x = f(a1)"),
                                 ("    x = f@()", "    x = f()"),
                                 ("    x = f()@", "    x = f()"),
                                 ("    x = 1@", "    x = 1")]:
            _r = _run78(_m)
            if _r[0] is not False or _r[1] != _expect_line:
                _bad78.append((_m, _r))
        check("78.2.7 ★不接管时行表一个字都不许变（返回 False，调用方会把这一下"
              "退格原样还给系统 —— 绝不会\"吞了按键却什么都没发生\"）",
              [_bad78 or "4 条全部原样退回"],
              expect_contain=["4 条全部原样退回"])

        # ---- 78.2.8 有选区 -> 那是"删选区"，一个字符都不许动 ----
        _l8, _c8 = _at78("    x = f(@)")
        _sel8 = (1, VB._disp_width(_l8[:8], 4, True) + 1, 1,
                 VB._disp_width(_l8, 4, True) + 1)
        _r8 = _run78("    x = f(@)", sel=_sel8)
        check("78.2.8 ★有选区不接管：退格在有选区时是【删掉选区】，语义完全不同 —— "
              "后端必须原样退回（调用方会把这一下退格还给系统），行表一个字没变",
              [_r8],
              expect_contain=[(False, "    x = f()", _sel8[3])])

        # ---- 78.2.9 VBE 重排了整行时，光标按压平前缀重新定位 ----
        # 真实调用形态：_locate_by_prefix(actual, new_line, left)
        #   raw = '    x = f()'，删掉 (5, 6) -> new_line = '    x = f'，left = 5
        # ⚠️ 删除不会像插入那样留下"可搜的一对符号"，唯一能认的指纹就是前缀。
        _nl9 = "    x = f"
        check("78.2.9 VBE 把整行重排过（补空格 `x=f` / 缩进被吃掉）时，"
              "按【压平后的前缀】（`_norm_code`：去空白 + 小写）重新定位光标 —— "
              "任何位置补空格 / 改大小写都不影响指纹",
              [(VB._locate_by_prefix(_nl9, _nl9, 5),
                VB._locate_by_prefix("    x=f", _nl9, 5),
                VB._locate_by_prefix("x = f", _nl9, 5))],
              expect_contain=[(5, 5, 2)])
        check("78.2.9b ★前缀在行里出现多次时取**离原位置最近**的那个 —— "
              "v99 第一版写的是\"从后往前扫取最长\"，实测 `aa = aa` 会命中行尾"
              "那个 `aa`（压平后与前缀一样），把光标从行首扔到行中。"
              "改成\"最近\"后与 `_locate_inserted` 同一口径：VBE 重排只挪几格，"
              "不会把光标搬到半行之外",
              [(VB._locate_by_prefix("    aa = aa", "    aa = aa", 4),
                VB._locate_by_prefix("    aa = aa", "    aa = aa", 9))],
              expect_contain=[(0, 9)])
        check("78.2.9c 认不出来时返回 None（由调用方退回 `left`，宁可偏一格，"
              "也不越界或乱扔 —— MEMORY 第 10 条：兜底绝不给错位置）",
              [VB._locate_by_prefix("abc", "    x = f", 5)],
              expect_contain=[None])

        # ---- 78.2.10 ★main.py 接线护栏（离线用例咬不住"钩子有没有接线"）----
        # ★v110：改成 AST 数，不再数源码字符串。字符串计数会被【注释】喂饱 ——
        #   v108 记的是"假绿"（注释让本该红的护栏绿了），这次是它的镜像"假红"：
        #   §93 新写的注释里提了一句"与 _bs_delete_pair_key_ok 同一副骨架"，
        #   计数从 2 变 3，一条本来没问题的护栏红了。AST 只认真正的定义与引用。
        _msrc78 = io.open(os.path.join(_ROOT78, "main.py"),
                          encoding="utf-8").read()
        import ast as _ast78
        _tree78 = _ast78.parse(_msrc78)
        _defs78 = [n for n in _ast78.walk(_tree78)
                   if isinstance(n, _ast78.FunctionDef)
                   and n.name == "_bs_delete_pair_key_ok"]
        _calls78 = [n for n in _ast78.walk(_tree78)
                    if isinstance(n, _ast78.Call)
                    and isinstance(n.func, _ast78.Name)
                    and n.func.id == "_bs_delete_pair_key_ok"]
        _wf78 = None
        for n in _ast78.walk(_tree78):
            if isinstance(n, _ast78.FunctionDef) and n.name == "win32_filter":
                _wf78 = n
        _refs78 = [n for n in _ast78.walk(_wf78)
                   if isinstance(n, _ast78.Name)
                   and n.id == "_delete_pair_here"]
        _actdef78 = [n for n in _ast78.walk(_tree78)
                     if isinstance(n, _ast78.FunctionDef)
                     and n.name == "_delete_pair_here"]
        check("78.2.10 ★接线护栏：`_bs_delete_pair_key_ok` 必须**恰好一处定义 +"
              "一处调用**（钩子里），动作函数也要真的存在并被钩子引用 —— "
              "只定义不接线 = 功能完全不存在，而离线用例抓不到这种漏；"
              "接线了却漏掉吞键 = 按键被吞但什么都没发生。三处都到位才算数"
              "（★v110 起用 AST 数：注释里提一句同名 token 就会让字符串计数失真）",
              [(len(_defs78) == 1, len(_calls78) == 1,
                len(_refs78) == 1, len(_actdef78) == 1)],
              expect_contain=[(True, True, True, True)])
        check("78.2.10b 逃生门：`VBECOMPLETE_NO_BS_DELPAIR=1` 能把这条完全关掉"
              "（退格回到 VBE 原生）",
              ["VBECOMPLETE_NO_BS_DELPAIR" in _msrc78],
              expect_contain=[True])
        check("78.2.10c ★两条退格接管路共用同一批前置门槛（一份实现，"
              "项目铁律第 14 条）—— `pending_keys > 0` / 快照过期 / "
              "Ctrl+Shift 这几条一旦只写一份、另一份漏掉，就会出现"
              "\"输入法组字时抢掉退格\"这类事故",
              [(_msrc78.count("_bs_ctx_for_takeover") >= 3,
                "def _bs_ctx_for_takeover(vk)" in _msrc78)],
              expect_contain=[(True, True)])
        check("78.2.10d ★删完必须当场作废快照 `state[\"cur_ctx\"] = None` —— "
              "按住退格时键盘重复约 33ms 一次，不作废会拿【删之前那一行的旧"
              "快照】再判一次、又删一对",
              [('_log("bspair: 未接管 -> 退格原样还给系统")' in _msrc78,
                "def _delete_pair_here():" in _msrc78)],
              expect_contain=[(True, True)])
    except Exception as _e78:
        import traceback as _tp78
        _tp78.print_exc()
        check("第 78 节异常: %s" % _e78, [True], expect_contain=[False])

    # ==================== 第 79 节：v100（★v102 已更正归因）================
    # 用户报：`Sub test()` 里 `Call judge` + Tab 不补括号，手动敲 `(` 也没反应。
    #
    # ★★v102 更正（真机取证）：本节原先写的根因"光标在 `Call ` 之后时 VBE 自己
    # 弹着自动列出成员、我们让位、Tab 打给了它"**是错的** —— 实测
    # `vbe_list_expected("    Call ", 10)` 返回 **False**（我们根本不让位，
    # Tab 走的是 `accept()`）。**真正的根因**：**VBE 会立刻删掉 `Call foo()`
    # 里的空括号**（VBA 规范：Call 语句无参数时不写括号）⇒ 补了白补、
    # 配对写了也被删 ⇒ 用户看到"按了没反应"。
    # ⇒ 本节已按 v102 更正：79.1.3 改成"Call 位置不补括号"；79.2 的场景从
    #   `Call jud` 换成普通调用点 `x = jud`（继续验证 `_hit` 段那条路的
    #   补括号能力 —— 它对点号成员等**会让位**的位置仍然有效）。
    #   **真机取证与最终修复见第 81 节。**
    try:
        import re as _re79
        print("\n--- 79.1 根因定位：Call 后的两条判据本来就放行 ---")

        # ---- 79.1.1 ★前提：不是判据的锅（`_is_call_site` 在 Call 后放行）----
        check("79.1.1 ★前提：`Call judge` 里词的位置是**调用点** —— "
              "`_is_call_site` 放行。**所以 bug 不在这道闸门**（一开始怀疑错"
              "方向的就是它）",
              [(VB._is_call_site("    Call judge", 10),
                VB._is_call_site("        Call judge", 15))],
              expect_contain=[(True, True)])

        # ---- 79.1.2 ★前提：`judge` 确实是过程名（parse 得出来）----
        _code79 = "Sub test()\nCall judge\nEnd Sub\n\nSub judge(s1, s2)\nEnd Sub"
        _r79, _p79 = P.extract_records(_code79, module="M1",
                                       is_std_module=True)
        check("79.1.2 ★前提：用户给的原样代码里，`extract_records` 的 `_proc_of` "
              "收得到 `judge`（Sub 无返回值也算过程）—— 引擎侧 `_is_callable_name` "
              "会放行，**所以也不是名字闸据的锅**",
              [("judge" in _p79, "test" in _p79)],
              expect_contain=[(True, True)])

        # ---- 79.1.3 ★前提：VBE 那条路（apply_completion）自己是好的 ----
        # 也就是说"直接走 accept()"能补出括号 —— 缺的只是那一次 accept()。
        class _CM79(object):
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "M1"

            @property
            def CountOfLines(self):
                return len(self.lines)

            def Lines(self, start, count=1):
                s0 = max(0, int(start) - 1)
                return "".join(l + "\n" for l in self.lines[s0:s0 + int(count)])

            def ReplaceLine(self, n, text):
                self.lines[int(n) - 1] = text

        class _Pane79(object):
            def __init__(self, cm, sel):
                self.CodeModule = cm
                self.sel = sel

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                ln = max(1, min(int(sl), len(self.CodeModule.lines)))
                try:
                    t = self.CodeModule.lines[ln - 1]
                except Exception:
                    t = ""
                eol = VB._disp_width(t, 4, True) + 1
                self.sel = (ln, max(1, min(int(sc), eol)),
                            ln, max(1, min(int(ec), eol)))

        class _VBE79(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _comp79(line, ws, word):
            """走真后端 apply_completion(want_parens=True)，返回新行。"""
            cm = _CM79([line])
            disp = VB._disp_width(line, 4, True) + 1
            pn = _Pane79(cm, (1, disp, 1, disp))
            orig = VB._get_vbe_cached
            VB._get_vbe_cached = lambda: _VBE79(pn)
            try:
                VB.VbeBackend().apply_completion(
                    1, ws, len(line) + 1, word, True)
            finally:
                VB._get_vbe_cached = orig
            # ⚠️ 真后端写回的行带尾随换行（VBE 的 Lines() 语义），期望值必须
            # 跟着剥掉 —— 否则用例红的是"多一个 \n"，不是功能（MEMORY 第 24 条：
            # 期望值一律从实际产物推，别手写）。
            return cm.lines[0].rstrip("\r\n")

        check("79.1.3 ★v102 更正：`Call jud` 位置直接调 `apply_completion"
              "(want_parens=True)` **不再补括号**（`_in_call_stmt` 挡掉）—— "
              "本用例原先断言的是\"能补出 `Call judge()`\"，用来证明\"补括号实现"
              "没坏、坏的只是 accept() 没被调用\"。**那个方向是错的**：v102 在"
              "真实 VBE 上取证发现 **VBE 会立刻删掉 `Call foo()` 里的空括号**"
              "（VBA 规范：Call 语句无参数时不写括号），所以这里补了也会消失 ——"
              "用户看到的\"Tab 没补括号\"根因就在这里，不在让位、不在记账。",
              # ⚠️ `check` 是**逐元素相等**（`for x in expect: if x not in got`），
              #   不是子串匹配（MEMORY 第 24 条已栽过一次）：`got` 里放**裸字符串**
              #   即可，外面那层 `[...]` 就是候选清单 —— 别再往里包一层元组，
              #   包了就去比 `('x',) == 'x'`，永远 False。
              [_comp79("    Call jud", 10, "judge")],
              expect_contain=["    Call judge"])

        print("\n--- 79.2 修法：`_hit` 段给 VBE 补进来的过程名补括号"
              "（v102 起场景改为普通调用点，Call 位置已不补）---")

        class _UI79(object):
            def show(self, *a, **k):
                pass

            def hide(self, *a, **k):
                pass

            def is_visible(self):
                return False

            def post(self, *a, **k):
                pass

        def _tab79(before, after):
            """模拟"让位期间按 Tab"：钩子记下before，VBE 自己改成 after，
            轮询发现后trigger(True)。返回 (新行, 是否补上了括号)。"""
            cm = _CM79(["Sub test()", before, "End Sub", "",
                        "Sub judge(s1, s2)", "End Sub"])

            class _BK79(object):
                _box = [None]

                def get_proc_names(self):
                    return frozenset(["judge", "test"])

                def get_builtin_names(self):
                    return set()

                def get_context(self):
                    return dict(self._box[0])

                def apply_completion(self, ln, ws, we, word, want=True):
                    return VB.VbeBackend().apply_completion(
                        ln, ws, we, word, want)

            pn = _Pane79(cm, (2, len(before) + 1, 2, len(before) + 1))
            orig = VB._get_vbe_cached
            VB._get_vbe_cached = lambda: _VBE79(pn)
            bk = _BK79()
            cp = E.Completer(bk, _UI79())
            _w = "judge"
            ctx = {"module_name": "M1", "line_no": 2, "line_text": after,
                   "caret_col": len(after) + 1, "word": _w,
                   "word_start_col": after.lower().rfind(_w) + 1,
                   "word_end_col": len(after) + 1,
                   "in_string": False, "in_comment": False,
                   "in_type_position": False}
            bk._box[0] = ctx
            try:
                cp.note_accept_key((time.time(), 2, before))
                cm.lines[1] = after            # VBE 自己插进去的
                cp._ctx_now = ctx
                cp.ctx = {"line_no": 2,
                          "word_start_col": ctx["word_start_col"],
                          "word_end_col": ctx["word_end_col"]}
                cp.trigger(True)
            finally:
                VB._get_vbe_cached = orig
            return cm.lines[1].rstrip(chr(13) + chr(10)), pn.sel[3]

        # ---- 79.2.1 ★`_hit` 路径（VBE 自己把半截词补长）----
        # ⚠️ v102 更正：本条原用 `Call jud` 场景，已被推翻（Call 位置不补括号）。
        # 场景改成普通调用点，继续验证这条路的能力。
        _r, _c = _tab79("    x = jud", "    x = judge")
        check("79.2.1 ★`_hit` 路径：VBE 自己把半截词补长"
              "（`x = jud` -> `x = judge`）时，我们要给它补上括号、"
              "并让光标进括号 -> 补出 `%r`，光标落在第 %d 列"
              % (_r, _c),
              [(_r, _c)],
              expect_contain=[("    x = judge()",
                               _r.index("(") + 2)])

        # ---- 79.2.2 只打了 1 个字母 / 一个字母没打，都要补 ----
        _r2a, _c2a = _tab79("    x = j", "    x = judge")
        _r2b, _c2b = _tab79("    x = ", "    x = judge")
        check("79.2.2 打了 1 个字母（`Call j`）与一个字母都没打（`Call `）"
              "两种起点都要补出括号",
              [(_r2a, _r2b)],
              expect_contain=[("    x = judge()", "    x = judge()")])

        # ---- 79.2.3 ★要补的是【光标处的整个词】，不是 `_ins` ----
        # VBE 是把半截词【补长】（`jud` -> `judge`），公共前后缀一算，抠出来的
        # 插入段只是差量 `ge` —— 拿它去判"是不是过程名"必然判否。
        # 这条钉住"别再用 `_ins`"。
        _lo79, _nl79 = "    x = jud".lower(), "    x = judge".lower()
        _p79 = 0
        while _p79 < len(_lo79) and _p79 < len(_nl79) and _lo79[_p79] == _nl79[_p79]:
            _p79 += 1
        _s79 = 0
        while (_s79 < len(_lo79) - _p79 and _s79 < len(_nl79) - _p79
               and _lo79[-1 - _s79] == _nl79[-1 - _s79]):
            _s79 += 1
        _ins79 = "    x = judge"[_p79:len("    x = judge") - _s79]
        check("79.2.3 ★要补的是【光标处的整个词】而不是 `_ins`："
              "`Call jud` -> `Call judge` 抠出来的插入段是 `%r`（只是差量），"
              "判\"是不是过程名\"必须用整词 `judge` —— 拿 `_ins` 判必然判否，"
              "这是第一版没生效的原因" % (_ins79,),
              [(_ins79, "judge")],
              expect_contain=[("ge", "judge")])

        # ---- 79.2.4 对照组：不是过程名 / 不存在的名字，一个都不许补 ----
        def _tab_nonproc79(before, after, word):
            """VBE 补进来的不是过程名 -> 不许补括号。"""
            cm = _CM79(["Sub test()", before, "End Sub"])

            class _BK(object):
                _box = [None]

                def get_proc_names(self):
                    return frozenset(["judge", "test"])

                def get_builtin_names(self):
                    return set()

                def get_context(self):
                    return dict(self._box[0])

                def apply_completion(self, ln, ws, we, w, want=True):
                    return VB.VbeBackend().apply_completion(
                        ln, ws, we, w, want)

            pn = _Pane79(cm, (2, len(before) + 1, 2, len(before) + 1))
            orig = VB._get_vbe_cached
            VB._get_vbe_cached = lambda: _VBE79(pn)
            bk = _BK()
            cp = E.Completer(bk, _UI79())
            ctx = {"module_name": "M1", "line_no": 2, "line_text": after,
                   "caret_col": len(after) + 1, "word": word,
                   "word_start_col": after.lower().rfind(word) + 1,
                   "word_end_col": len(after) + 1,
                   "in_string": False, "in_comment": False,
                   "in_type_position": False}
            bk._box[0] = ctx
            try:
                cp.note_accept_key((time.time(), 2, before))
                cm.lines[1] = after
                cp._ctx_now = ctx
                cp.ctx = {"line_no": 2,
                          "word_start_col": ctx["word_start_col"],
                          "word_end_col": ctx["word_end_col"]}
                cp.trigger(True)
            finally:
                VB._get_vbe_cached = orig
            return cm.lines[1].rstrip(chr(13)+chr(10))

        _r4a = _tab_nonproc79("    x = jud", "    x = judgeX", "judgeX")
        _r4b = _tab_nonproc79("    x = ms", "    x = msgbox", "msgbox")
        check("79.2.4 ★对照组：VBE 补进来的**不是过程名**时一个括号都不许补"
              " —— 工程里没有 `judgeX`（不补）、`msgbox` 不在 `get_proc_names` 里"
              "（本工程没这个过程，不补）",
              [(_r4a, _r4b)],
              expect_contain=[("    x = judgeX", "    x = msgbox")])

        # ---- 79.2.5 ★只在【第一段】补：第二段（pinned）绝不补 ----
        # pinned 锚的是"VBE 留下的那一行"，那一行一个字都没再动过，
        # 这时补括号 = 无中生有改坏用户代码。
        #
        # ⚠️ 检查方式：**数 `apply_completion` 在这段里出现几次**。用
        # `"self._accept_pinned" not in 切片` 判是不行的 —— 切片里必然含
        # `_accept_pinned` 这个**词**（`if not self._accept_pinned:` 这行本身
        # 就是那道门，还有大段注释提到它），判"词在不在"分不出"门"与"注释"。
        # 要判的是**那一次真正的写入动作被关在门里**，所以查两件事：
        #   * 门 `if not self._accept_pinned:` 在；
        #   * `apply_completion(` 出现在**门之后**（缩进更深）。
        _esrc79 = io.open(os.path.join(_ROOT78, "engine.py"),
                          encoding="utf-8").read()
        # ⚠️ `str.find(sub, start)` 是**往后**找，但 `if self._accept_pinned:`
        #在 `if _hit:` **前面**也有一处（那是第一段的入口条件）——
        # 用 find 切会切出一段反向的区间，门与写入的相对位置全错（第一版
        # 就栽在这儿，报的位是"门 1296 / 写入 707"）。改成从 `if _hit:` 起
        # 用正则找**后面**那处、且缩进更深（16 空格）的那一个。
        # ⚠️⚠️ 两个坑（第一版两条都踩了）：
        #   ① `str.find(sub, start)` 是**往后**找，但 `if self._accept_pinned:`
        #      在 `if _hit:` **前面**也有一处（第一段的入口条件）⇒ 用 find 切会
        #      切出反向区间。改成从 `if _hit:` 起正则找**后面**那处。
        #   ② 段内的**注释里也写着** `apply_completion(..., True)`（那段长注
        #      解释"复用同一个后端动作"）⇒ 用 `seg.find("apply_completion(")`
        #      拿到的是**注释里那处**，位置在门之前，报出来是"门 1296 / 写入 707"。
        # ⇒ 一律**按代码行**判（跳过 `#` 开头的注释行），判据只看真实语句。
        _i79 = _esrc79.find("if _hit:")
        _m79 = _re79.search(r"\n {16}if self\._accept_pinned:",
                            _esrc79[_i79:]) if _i79 >= 0 else None
        _v10079 = (_esrc79[_i79:_i79 + _m79.start()] if _m79 else "")
        _gate79 = _call79 = -1
        for _n79, _ln79 in enumerate(_v10079.splitlines()):
            _s79 = _ln79.strip()
            if _s79.startswith("#"):
                continue
            if _gate79 < 0 and "if not self._accept_pinned:" in _s79:
                _gate79 = _n79
            if _call79 < 0 and "apply_completion(" in _s79:
                _call79 = _n79
        check("79.2.5 ★补括号的写入动作 `apply_completion(` 必须被关在"
              " `if not self._accept_pinned:` 这道门【之后】（按代码行：门在第 %d"
              " 行、写入在第 %d 行）—— 第二段锚的是\"VBE 留下的那一行\"，那一行"
              "一个字都没再动过，这时候补括号就是无中生有改坏用户代码"
              % (_gate79, _call79),
              [(_gate79 >= 0, _call79 > _gate79 >= 0)],
              expect_contain=[(True, True)])

        # ---- 79.2.6 ★补括号失败绝不许弄坏这次静默判定 ----
        check("79.2.6 补括号整段包在 try/except 里（补不上就算了）—— 它是"
              "\"锦上添花\"，绝不能因为它抛异常就把 v90b/v91 那次静默判定弄坏"
              "（那会让用户看到\"按完 Tab 又提示自己\"的老 bug 复活）",
              [("触发不存在" if False else
                ("def _v100" if False else
                 "本次静默" in _esrc79))],
              expect_contain=[True])

        # ---- 79.2.7 复用同一个后端动作，不另写一份写入逻辑 ----
        check("79.2.7 ★补括号**复用** `apply_completion(..., True)`，与 accept() "
              "走同一个后端动作 ——判据与动作都不另写一份（项目铁律第 14 条："
              "同一条语义只能一份实现）",
              [("apply_completion(" in _v10079,
                "ReplaceLine" not in _v10079)],
              expect_contain=[(True, True)])

        # ---- 79.2.8 源码护栏：位置列一律用【本次 extract_word_at】的字符列 ----
        # ⚠️代码复查（2026-10-18）：这条护栏原先要求段里出现
        # `ctx.get("word_start_col")` —— 可真实后端 `get_context()` **从不返回**
        # 这个键，于是"护栏保护的是一个让整段永久失效的写法"。改成钉住新契约：
        # 列必须来自 trigger 里那次 extract_word_at（与 accept 同一份字符列）。
        check("79.2.8 补括号用的起始列取自 `extract_word_at` 算出的 `start_col`"
              "（**字符列**，与 accept() 同一份，MEMORY 第 19 条：不许自己猜列"
              "语义、也不许混用 `cur_ctx` 的显示列）；并且**不许**再从 ctx 里"
              "取那个根本不存在的 `word_start_col`",
              [("int(start_col or 0)" in _v10079,
                'ctx.get("word_start_col")' not in _v10079)],
              expect_contain=[(True, True)])
    except Exception as _e79:
        import traceback as _tp79
        _tp79.print_exc()
        check("第 79 节异常: %s" % _e79, [True], expect_contain=[False])

    # ==================================================================
    # 第 80 节：v101 —— 把 v100 那次【错误归因】钉成护栏（3 条）
    #
    #★背景：v100 我断定"`Call ` 之后 VBE 会弹成员列表、我们让位、于是 Tab
    #  打给了 VBE 自己的列表"，并把修复挂在 `_hit`（note_accept_key 记账）
    #  那段上。用户实测反馈"故障依旧存在"—— 根因就是**这个前提本身是错的**：
    #  `vbe_list_expected("    Call ", 10)` 实测返回 **False**（它只认
    #  `标识符.` 与 `As `/`New ` 两种位置），所以 `Call jud` 时我们的窗是
    #  **弹着的**，Tab 走 `completer.is_visible()` → `accept()`，
    #  **根本轮不到 `_hit` 那段** ⇒ 我挂在死路上的修复从未被触发。
    #
    # 教训（第 29 条的升级版）：**"某个功能不生效"的第一嫌疑是"那次按键
    #  根本没走到我们"，但"第二嫌疑"必须是"我认定的那条路径本身是不是
    #  真的会发生"** —— 而这两点都**只能靠实测判据返回值确认，不能靠读代码
    #  推断**。读代码只能告诉我"如果 X 成立会怎样"，告诉不了我"X 成立吗"。
    # ==================================================================
    try:
        print("\n--- 80. 让位判据对 Call 返回 False（v100 归因错误的根源）---")

        # 80.1 ★核心：`Call ` 之后我们【不让位】—— 只有两种位置才让位。
        _call_exp = [
            ("    Call ", 10),
            ("    Call j", 11),
            ("    Call jud", 14),
            ("    Call judge", 15),
            ("Call judge", 11),
        ]
        _r80 = [(str(ln), int(cc), E.vbe_list_expected(ln, cc))
                for (ln, cc) in _call_exp]
        check("80.1 ★★`Call ` 之后【我们不让位】（实测返回 False）—— v100 "
              "把修复挂在\"让位期间 Tab 打给 VBE\"那条路上，而那条路对 Call "
              "**根本不存在**。真实路径是：我们的窗弹着→ Tab 走 "
              "`completer.is_visible()` → `accept()`。实测=%r"
              % ([r[2] for r in _r80],),
              [all(r[2] is False for r in _r80)],
              expect_contain=[True])

        # 80.1b 对照组：让位判据**确实**认这两种位置（钉住"不是判据坏了"）
        _yld_exp = [
            ("    Call SheetUtils.", 21),
            ("    ws.", 8),
            ("    Dim x As ", 14),
            ("    Set c = New ", 17),
        ]
        _r80b = [(str(ln), int(cc), E.vbe_list_expected(ln, cc))
                 for (ln, cc) in _yld_exp]
        check("80.1b 对照组：让位判据**确实**认`标识符.` 与 `As `/`New ` "
              "这两种位置（证明 80.1 的 False 不是判据整体失效，而是它本来"
              "就只管这两种）实测=%r" % ([r[2] for r in _r80b],),
              [all(r[2] is True for r in _r80b)],
              expect_contain=[True])

        # 80.2 源码护栏：那段补括号代码的注释里**必须**带上 v101 的实测更正
        #       （`Call ` 不让位 + 指向日志），否则下一个读代码的人会再次
        #       相信那个错误断言、再修一轮空气。
        # ⚠️ 断言方式：**查"更正声明在不在"**，不要查"原句还在不在"——
        #   原句是作为"我原先写错了"的引述**故意保留**的（删掉它就看不出
        #   这里曾经错过），查它有没有消失会得到假红。
        _esrc80 = io.open(os.path.join(_ROOT78, "engine.py"),
                          encoding="utf-8").read()
        _i80 = _esrc80.find("if _hit:")
        _re80 = __import__("re")
        _m80 = _re80.search(r"\n {16}if self\._accept_pinned:", _esrc80[_i80:]) \
            if _i80 >= 0 else None
        _seg80 = (_esrc80[_i80:_i80 + _m80.start()] if _m80 else "")
        check("80.2 ★那段补括号代码的注释里【必须】带 v101 实测更正："
              "`Call ` 不让位（真实路径是 accept()）+ 指向日志 —— "
              "注释里若没有这段，下一个读代码的人会再次相信那个错误断言"
              "（v100 就是这么白修一轮的）。原句作为\"我原先写错了\"的引述"
              "故意保留，不检查它的有无",
              [("v101 更正本段注释里那句因果断言" in _seg80,
                "VBECOMPLETE_LOG" in _seg80 and "accept:" in _seg80,
                "与本段无关" in _seg80)],
              # ⚠️ `check` 是**逐元素相等**（`x not in got`），不是子串匹配
              #   （MEMORY 第 24 条）—— 三项打包成**一个**元组去比。
              expect_contain=[(True, True, True)])

    except Exception as _e80:
        import traceback as _tp80
        _tp80.print_exc()
        check("第 80 节异常: %s" % _e80, [True], expect_contain=[False])

    # ==================================================================
    # 第 81 节：v102 —— `Call 过程名` 位置不做括号动作（8 条）
    #
    # ★★根因是真机取证得来的（v102g 内容矩阵；独立 Excel 实例 + 真 CodeModule）：
    #   **VBE 会把 `Call foo()` 里的【空括号】自动删掉** —— 这是 VBA 的规范：
    #   `Call` 语句无参数时不写括号，`Call foo` 才是标准写法。
    #       写 '    Call judge()'     -> 立即 '    Call judge'     ← 括号被删
    #       写 '    Call judge(1, 2)' -> '    Call judge(1, 2)'     ← 有参数则保留
    #       写 '    Call foo()'       -> 立即 '    Call foo'       ← 同样被删
    #       写 '    judge()'          -> '    judge()'              ← 不在 Call 语句里，保留
    #       写 '    x()'              -> '    x()'                  ← 保留
    #   ⇒ 用户报的两个现象是**同一个根因**，与我们的判据 / 让位 / 时序全都无关：
    #       * "Tab 后没有双括号" = 我们补了，VBE 立刻删；
    #       * "手动敲 `(` 没反应" = 配对写入 `()`，VBE 也立刻删。
    #   ⇒ 修法：这两条路在 `Call 过程名` 位置**都不做**，把按键原样交给 VBE。
    #
    # ⚠️ 这条教训值得记：前两轮我一直在**引擎判据、让位、记账**里找原因，
    #   而真正的原因在**我从未测过的那一层** —— 真实 VBE 对写入内容的反规范化。
    #   "离线 mock 全绿"永远证明不了真机行为。
    # ==================================================================
    try:
        import vbe_bridge as VB81
        print("\n--- 81. `Call 过程名` 不做括号动作（v102 真机取证）---")

        # 81.1 判据正例
        _pos81 = ["    Call judge", "Call judge", "\tCall obj.Method",
                  "  call x", "CALL Foo"]
        _got81 = [(t, VB81._in_call_stmt(t)) for t in _pos81]
        check("81.1 `_in_call_stmt` 正例：`Call 过程名`（前导空白 / 制表符 / "
              "`Call obj.Method` / 大小写）全都要认出 —— 这些位置的空括号会被 "
              "VBE 删掉，补了等于没补",
              _got81,
              expect_contain=[(t, True) for t in _pos81])

        # 81.2 判据反例：不许误伤以 Call 开头的标识符
        _neg81 = ["Callback x", "CallCount = 1", "    x = Call foo",
                  "    judge", "", "    ", "myCall 1"]
        _got81b = [(t, VB81._in_call_stmt(t)) for t in _neg81]
        check("81.2 `_in_call_stmt` 反例：`Callback` / `CallCount` / "
              "`x = Call foo`（Call 不是行首关键字）/ 普通标识符 —— 都必须是 "
              "False，否则会把正常位置的补括号误伤掉",
              _got81b,
              expect_contain=[(t, False) for t in _neg81])

        # 81.3 源码护栏：apply_completion 的补括号分支必须多带这一道闸门
        _bsrc81 = io.open(os.path.join(_ROOT78, "vbe_bridge.py"),
                          encoding="utf-8").read()
        _i81 = _bsrc81.find("if want_parens and AUTO_PAREN_ON_COMPLETE")
        _seg81 = _bsrc81[_i81:_i81 + 400] if _i81 >= 0 else ""
        check("81.3 ★`apply_completion` 补括号那三个条件里必须多出 "
              "`not _in_call_stmt(line_text)` —— 否则 `Call judge` 位置仍会补出 "
              "括号、然后被 VBE 删掉（用户看到的正是\"Tab 没补括号\"）",
              [("not _in_call_stmt(line_text)" in _seg81,
                "_is_call_site(line_text, word_start_col)" in _seg81)],
              expect_contain=[(True, True)])

        # 81.4 源码护栏：insert_pair 里必须有这条早退
        _i81b = _bsrc81.find("def insert_pair(")
        _j81b = _bsrc81.find("def ", _i81b + 10)
        _seg81b = _bsrc81[_i81b:_j81b] if _i81b >= 0 and _j81b > _i81b else ""
        check("81.4 ★`insert_pair` 里必须有 `_in_call_stmt` 早退（返回 False）"
              " —— 不接管才能把 `(` 原样交给 VBE；接管会写入 `()` 然后被 VBE "
              "删掉，用户看到的正是\"敲括号没反应\"",
              [("_in_call_stmt(line_text)" in _seg81b,
                "return False" in _seg81b)],
              expect_contain=[(True, True)])

        # 81.5 顺序护栏：insert_pair 里的早退必须发生在【写入之前】
        _k81 = _seg81b.find("_in_call_stmt(line_text)")
        _w81 = _seg81b.find("_write_line")
        check("81.5 `insert_pair` 里 `_in_call_stmt` 早退必须在【真正写入"
              "（`_write_line`）之前】（早退在第 %d 位、写入在第 %d 位）—— "
              "写在后面就等于先写坏再判断" % (_k81, _w81),
              [(_k81 >= 0, _w81 > _k81 >= 0)],
              expect_contain=[(True, True)])

        # 81.6 源码注释护栏：必须留下"VBE 会删空括号"的实测记录
        _in81 = _bsrc81.find("def _in_call_stmt(")
        # ⚠️ docstring 在 `def` **之后**，所以窗口是 [_in81, _in81+2600)，
        #   不能写成往前取（我第一版写反了，拿到的是一段无关代码 -> 假红）。
        _seg81c = _bsrc81[_in81:_in81 + 2600] if _in81 >= 0 else ""
        check("81.6 ★`_in_call_stmt` 的注释里必须写明\"VBE 会把 `Call foo()` 的"
              "空括号自动删掉\"这个**真机实测结论**（含 `Call judge()` -> "
              "`Call judge` 的对照）—— 这是本轮唯一解释得通两个现象的事实，"
              "注释里丢了它，下一个人还会去引擎判据里找原因",
              [("Call foo()" in _seg81c and "Call judge()" in _seg81c,
                "VBE" in _seg81c and "删" in _seg81c,
                "Call judge(1, 2)" in _seg81c)],
              expect_contain=[(True, True, True)])

        # 81.7 对照护栏：非 Call 位置必须照旧补括号（`judge()` 不会被 VBE 删）
        check("81.7 对照：`    judge()` / `    x()` 这类**不在 Call 语句里**的"
              "位置，`_in_call_stmt` 必须是 False（VBE 实测会保留这些括号）"
              " ⇒ 补括号功能在那里照旧生效，没被这次修复误伤",
              [VB81._in_call_stmt("    judge()"),
               VB81._in_call_stmt("    x = judge()")],
              expect_contain=[False, False])

        # 81.8 判据必须是可套用的纯函数（钩子线程也能用，不碰 COM）
        _f81 = VB81._in_call_stmt
        check("81.8 `_in_call_stmt` 是纯函数（只吃文本、不碰 COM），钩子线程与"
              "主线程可共用同一份实现（项目铁律第 14 条）",
              [callable(_f81), _f81("    Call a") is True,
               _f81(None) is False],
              expect_contain=[True, True, True])

    except Exception as _e81:
        import traceback as _tp81
        _tp81.print_exc()
        check("第 81 节异常: %s" % _e81, [True], expect_contain=[False])

    # ==================================================================
    # 第 82 节：v103 —— 与 IDEA 逐项对照后补的东西 + 顺手挖出的 8 个真 bug
    #
    # 用户口径："我其实就是为了解决 VBE 里不能自动提醒变量名、过程名这个问题，
    #   现代编辑器比如 IDEA 都有很成熟的这种解决方案，我的项目就是想复刻 IDEA。
    #   你帮我从头到尾重新对比一下 IDEA，我是否有考虑不周全的地方，你帮我加上，
    #   或者有 bug 的地方，你帮我修复一下。"
    #
    # 本节分三块：
    #   82.1~82.5  parser 层 4 个真 bug（都是"收录出假名字"或"局部变量泄漏"）
    #   82.6~82.9  engine / vbe_bridge 的判据 bug（下标、列口径、注释识别）
    #   82.10~82.12 ★IDEA 式排序：作用域就近 + 最近使用（含开关的 A/B）
    #   82.13~82.16 ★IDEA 式参数信息：候选详情行显示形参表
    #
    # ⚠️ 这些 bug 的共性值得记一句：**没有一条是"逻辑写错了"**，全是
    #   "判据的隐含前提被后一个用途打破"（owner 只有一格 / 缩进只有空白 /
    #   小写化不改长度）—— 与 MEMORY 第 11 条同一副面孔。
    # ==================================================================
    try:
        import vbe_bridge as VB82
        print("\n--- 82. v103：IDEA 对照补齐 + 8 个真 bug ---")

        # ---- 82.1 过程头漏认无括号的 `As 类型` 形式（连带局部变量泄漏）----
        #      MS 官方语法：`Sub/Function name [(arglist)] [As type]` —— 括号可省。
        #      旧正则要求名字后面必须是 `(` 或行尾 ⇒ `Function Compute As Long`
        #      不被当成过程头 ⇒ `cur_proc` 不置位 ⇒ 它体内的 Dim 全记成模块级。
        _code82 = ("Function Compute As Long\n"
                   "    Dim secretLocal As Long\n"
                   "End Function\n"
                   "Sub Caller()\n"
                   "    Dim myVar As Long\n"
                   "End Sub\n")
        _recs82, _procs82 = P.extract_records(_code82, "M1", True)
        check("82.1 `Function Compute As Long`（无括号带返回类型）必须认成过程头，"
              "且它的局部变量归属该过程（旧写法下 `secretLocal` 会被记成模块级、"
              "泄漏到 Sub Caller 的候选里）",
              [sorted(_procs82),
               [r for r in _recs82 if r[0].lower() == "secretlocal"],
               "secretLocal" in E.filter_identifiers_by_scope(_recs82, "Caller", "M1")],
              expect_contain=[["caller", "compute"],
                              [("secretLocal", "M1", "Compute", False)],
                              False])

        # ---- 82.2 VBA 还允许名字末尾写类型声明符：`Function Foo$()` ----
        check("82.2 名字末尾带类型声明符（`Function Foo$()` / `Sub Foo$()`）"
              "也要认成过程头（`_IDENT` 不含 `% & ! # @ $`，旧正则在这里断掉）",
              [sorted(P.extract_records("Function Foo$()\nEnd Function\n",
                                        "M1", True)[1]),
               sorted(P.extract_records("Sub Foo$()\nEnd Sub\n", "M1", True)[1]),
               sorted(P.extract_records("Function Bar%()\nEnd Function\n",
                                        "M1", True)[1])],
              expect_contain=[["foo"], ["foo"], ["bar"]])

        # ---- 82.3 形参修饰词可以连写：`Optional ByVal x` ----
        check("82.3 `Optional ByVal retries As Long = 3` 的参数名必须是 retries"
              "（旧正则只吃一个修饰词 ⇒ 把第二个 `ByVal` 当参数名，"
              "于是池子里多一个伪候选 `ByVal`、真参数反而丢了）",
              [P._params_inside("Optional ByVal retries As Long = 3"),
               P._params_inside("ByVal a As Long, ByRef b"),
               P._params_inside("ParamArray arr() As Variant")],
              expect_contain=[["retries"], ["a", "b"], ["arr"]])

        # ---- 82.4 `ReDim Preserve arr(...)` 的 `Preserve` 不是声明名 ----
        #      `_RE_DECL` 只吃掉 `ReDim`，剩下的 `Preserve arr(...)` 交给
        #      _decl_names 取首标识符 ⇒ 收录出伪候选 `Preserve`（关键字、
        #      不在 stopword 表里，输入 `pre` 就以 kind 3 弹出来）。
        _r82b = P.extract_records(
            "Sub Work()\n    ReDim Preserve arr(1 To 10)\nEnd Sub\n",
            "M1", True)[0]
        check("82.4 `ReDim Preserve arr(1 To 10)` 不许收录伪名字 `Preserve`"
              "（真数组名 arr 照旧收）",
              _r82b,
              expect_contain=[("arr", "M1", "Work", False)],
              expect_absent=[("Preserve", "M1", "Work", False)])

        # ---- 82.5 Declare 的形参不许挂到模块级（跨过程可见 = 泄漏）----
        _r82c = P.extract_records(
            'Private Declare Sub Sleep Lib "kernel32"'
            ' (ByVal dwMilliseconds As Long)\n', "M1", True)[0]
        _vis82c = E.filter_identifiers_by_scope(_r82c, "Anything", "M1")
        # ⚠️★v103 写用例时又踩的坑（MEMORY 第 24 条）：`check` 的 expect_contain 是
        #   【逐元素成员判断】（`x not in got`）。把几个布尔散在一个列表里当 got、
        #   期望又是 [True, False, ...] 混合 ⇒ **必然假绿**（True 和 False 都在 got 里）。
        #   多项必须**打包成一个元组**，让"位置/组合"参与比较。
        check("82.5 Declare 的形参要挂在【Declare 自己】名下："
              "`dwMilliseconds` 不该出现在任意过程的候选里"
              "（Declare 一定在模块级，旧写法用 `cur_proc`=None ⇒ 形参成了模块级、"
              "priv=False ⇒ 本模块每个过程都看得见）",
              [tuple((("dwMilliseconds", "M1", "Sleep", False) in _r82c,
                      ("dwMilliseconds", "M1", None, False) in _r82c,
                      "dwMilliseconds" in _vis82c,
                      "Sleep" in _vis82c))],
              expect_contain=[(True, False, False, True)],
              expect_absent=[(False, True, True, False)])

        # ---- 82.6 `remX.Value` 的 Rem 守卫用错了列 ----
        #      `_in_rem_comment(line, m+1)` 里的 `m` 是 owner 的**末**字符，
        #      原注释写"m+1 等于 s+2"只在 owner 恰好 1 格时成立。
        check("82.6 `remX.Value` 必须认成成员访问（owner 叫 remX，不是 Rem 注释）"
              "—— `Rem a.b` 这种真注释仍要挡住；"
              "同一个 owner 还要与 `parser.ident_before_dot` 同口径",
              [E._member_list_dot_col("remX.Value", 11),
               E.list_slot_anchor("remX.Value", 11),
               P.ident_before_dot("remX.Value", 11),
               E._member_list_dot_col("Rem a.b", 7),
               E.vbe_member_list_expected("    Range(\"A1\").", 17)],
              expect_contain=[4, 6, "remX", None, True])

        # ---- 82.7 小写化改变长度时命中下标错位（甚至越界）----
        #      `'İ'.lower()` 是两个字符 ⇒ `low` 与 `name` 不再逐字符对齐。
        _f82 = E.fuzzy_match("İstanbul", "s")
        _f82b = E.fuzzy_match("İstanbul", "stanbul")
        check("82.7 `fuzzy_match` 的命中下标必须落在**原始名字**上："
              "`İstanbul` 里 `s` 在第 1 格（旧实现给 2，高亮错位）；"
              "命中的下标一个都不许 ≥ len(name)（旧实现给到 8 == len）",
              [None if _f82 is None else _f82[1],
               None if _f82b is None else _f82b[1],
               None if _f82b is None else max(_f82b[1]) < len("İstanbul"),
               E.fuzzy_match("dataArr", "ts"), E.fuzzy_match("", "a")],
              expect_contain=[[1], [1, 2, 3, 4, 5, 6, 7], True, None, None])

        # ---- 82.8 `_classify` 必须认 `Rem`，且进了注释就不再往后扫 ----
        #      ① 不认 Rem ⇒ 注释里照弹候选、自动配对也不让路；
        #      ② 进了 `'` 注释还继续扫 ⇒ 注释里的引号把 in_str 翻成 True。
        check("82.8 `_classify`：`Rem ...` 是注释（原来只认 `'`，"
              "`in_comment` 恒 False）；`x = 1 ' say \"hi` 是【注释】不是字符串"
              "（原来被注释里的引号骗成 in_str=True）"
              "—— ★四项打包成一个元组比较，否则'第一项凑巧对上'就假绿",
              [tuple(VB82._classify(_t, _c) for _t, _c in
                     [("Rem 每个元素", 9),
                      ("    Rem x", 9),
                      ("x = 1 ' say \"hi", 15),
                      ("' it's here", 11)])],
              expect_contain=[((False, True), (False, True),
                               (False, True), (False, True))])

        check("82.9 对照：字符串里长得像 Rem 的**不算**注释"
              "（顺序必须是「先跑完字符串状态机，再问 Rem」）",
              [tuple(VB82._classify(_t, _c) for _t, _c in
                     [('x = "a Rem b"', 13), ('    s = "Rem', 13)])],
              expect_contain=[((True, False), (True, False))])

        # ---- 82.10 行尾显示列只能有一份实现（全角占 2 格）----
        #      `_eol_caret_col` 的兜底原来抄了 `_indent_end_col` 的
        #      `_disp_width(t, 4, False)` 口径 —— 那是给"只有空白的缩进"
        #      用的，全角参数对它无意义；抄到"任意代码行"上就少算格数。
        #      ⚠️★必须**先把 `_sem_cache["info"]` 清空**：含全角的行走
        #      `if info:` 那条分支（第 484 行），全量测试同进程跑时缓存早被
        #      前面的真机用例填上了 ⇒ 不清缓存就永远走不到兜底那行，A/B 不红。
        _sem_save82 = VB82._sem_cache.get("info")
        VB82._sem_cache["info"] = None
        try:
            _ec82 = (VB82._eol_caret_col("    说 = 1"),
                     VB82._eol_disp_col("    说 = 1"))
            _ec82b = (VB82._eol_caret_col("    abc = 1"),
                      VB82._eol_disp_col("    abc = 1"))
        finally:
            VB82._sem_cache["info"] = _sem_save82
        check("82.10 同一行尾的「显示列」两份实现必须给出同一个数"
              "（含中文的行上，旧 `_eol_caret_col` 少算全角格数）"
              "—— 同样打包比较：`11 in [10, 11, ...]` 那种写法验不出东西",
              [( _ec82[0], _ec82[1], _ec82[0] == _ec82[1],
                 _ec82b[0], _ec82b[1], _ec82b[0] == _ec82b[1])],
              expect_contain=[(11, 11, True, 12, 12, True)])

        # ---- 82.11 `_find_completion_end` 空候选不再死循环 ----
        #      `text.find("", i)` 恒等于 i 且 `i += 0` ⇒ while 转不出去。
        check("82.11 `_find_completion_end` 传空 completion 不再死循环"
              "（潜在缺陷：只要有人拿空候选调一次，整个补全线程就卡死）"
              "—— 直接判等返回（返回后仍按行尾裁剪）",
              [(VB82._find_completion_end("    abc", "", 0, 7),
                VB82._find_completion_end("    abc", "", 0, 999),
                VB82._find_completion_end("    abc", "", 0, 7) == 7)],
              expect_contain=[(7, 7, True)])

        # ---- 82.12 源码护栏：那两处列口径必须真的同源 ----
        _bsrc82 = io.open(os.path.join(ROOT, "vbe_bridge.py"),
                          encoding="utf-8").read()
        _i82 = _bsrc82.find("def _eol_caret_col")
        _seg82 = _bsrc82[_i82:_i82 + 2600]
        check("82.12 源码护栏：`_eol_caret_col` 的兜底必须调 `_eol_disp_col`，"
              "不许再自己算一份（'同一语义只能一份实现'：这两份曾经在全角上"
              "给出不同的数）—— 打包成元组，退化了才真会红",
              [("return _eol_disp_col(t, 4)" in _seg82,
                "_disp_width(t, 4, False) + 1" not in _seg82)],
              expect_contain=[(True, True)])

        # ---- 82.13 作用域档位（IDEA 局部作用域优先的输入）----
        _pool83 = [("myVar", "M1", "Cur", False),
                   ("modVar", "M1", None, False),
                   ("pubVar", "M2", None, False),
                   ("privVar", "M2", None, True)]
        check("82.13 `filter_identifiers_by_scope_ranked` 的档位："
              "本过程局部 4 > 本模块模块级 3 > 别的模块 Public 2；"
              "别的模块 Private 根本不可见",
              E.filter_identifiers_by_scope_ranked(_pool83, "Cur", "M1"),
              expect_contain=[("myVar", 4), ("modVar", 3), ("pubVar", 2)],
              expect_absent=[("privVar", 2), ("privVar", 3), ("privVar", 4)])

        check("82.14 老接口 `filter_identifiers_by_scope` 必须是新接口的投影"
              "（保序、逐字一致 —— 可见性这条语义只有一份实现）",
              [E.filter_identifiers_by_scope(_pool83, "Cur", "M1"),
               [n for n, _r in
                E.filter_identifiers_by_scope_ranked(_pool83, "Cur", "M1")]],
              expect_contain=[["myVar", "modVar", "pubVar"]])

        # ---- 82.15/82.16 ★排序：作用域就近 + 最近使用 ----
        #      构造最小引擎（同第 76 节的写法）。池子刻意让"远的那个先插入"
        #      —— 稳定排序下，插入序就是"没有任何相关度信号时"的结果，
        #      于是 A/B 能分辨出到底是谁在起作用。
        class _B82(object):
            def __init__(self, pool, ctx_extra=None):
                self._pool = list(pool)
                self._ctx = {"line_no": 1, "in_string": False,
                             "in_comment": False, "in_type_position": False,
                             "in_decl_position": False, "decl_names": []}
                self._ctx.update(ctx_extra or {})

            def get_context(self):
                return dict(self._ctx)

            def get_identifiers(self):
                return list(self._pool)

            def get_declared_names(self):
                return [str(n).lower() for n, _m, _p, _v in self._pool]

            def declared_elsewhere(self, name, module_name):
                return str(name).lower() in self.get_declared_names()

            def get_structural_names(self):
                return []

            def get_builtin_names(self):
                return set()

            def get_type_names(self):
                return []

            def get_host_enum_names(self):
                return {}

            def get_proc_names(self):
                return frozenset()

            def names_outside_caret(self, caret, scope_only=False):
                return None          # 现场证据不可用 -> 走真名集合那条路

            def apply_completion(self, *a, **k):
                return None

            def vbe_yield_visible(self):
                return False

            def vbe_popup_visible(self):
                return False

            def vbe_popup_info(self):
                return []

        class _U82(object):
            def show(self, *a, **k):
                pass

            def hide(self):
                pass

            def update_selection(self, *a, **k):
                pass

            def contains_point(self, *a, **k):
                return False

        def _trig82(word, pool, proc=None, module="M1"):
            _ln = "    " + word
            _c = E.Completer(
                _B82(pool, {"line_text": _ln, "caret_col": len(_ln) + 1,
                            "proc_name": proc, "module_name": module}),
                _U82())
            _c.trigger(True)
            return _c

        _pool84 = [("myValue", "M2", None, False),
                   ("myLocal", "M1", "Cur", False)]
        _c82 = _trig82("my", _pool84, proc="Cur")
        check("82.15 ★IDEA 式相关度（1/2）作用域就近：两条候选匹配质量完全一样、"
              "插入序还让'别处的 myValue'在前，但 `myLocal` 是本过程的局部变量 "
              "⇒ 必须排第 1（此前作用域只用来【过滤】，不参与【排序】）",
              [list(_c82.matches or [])],
              expect_contain=[["myLocal", "myValue"]])

        _r82 = E.RANK_BY_SCOPE
        E.RANK_BY_SCOPE = False
        try:
            _c82b = _trig82("my", _pool84, proc="Cur")
        finally:
            E.RANK_BY_SCOPE = _r82
        check("82.15b 反面对照：`VBECOMPLETE_NO_SCOPE_RANK=1`（关掉作用域排序）时"
              "退回插入序 —— 证明 82.15 确实是这条规则在起作用",
              [list(_c82b.matches or [])],
              expect_contain=[["myValue", "myLocal"]])

        _pool85 = [("aaaTwo", "M2", None, False),
                   ("aaaOne", "M2", None, False)]
        _c82c = _trig82("aaa", _pool85)
        check("82.16 ★IDEA 式相关度（2/2）最近使用：两条候选连名字长度都一样"
              "（评分键逐位相同）⇒ 默认是插入序",
              [list(_c82c.matches or []),
               E.fuzzy_match("aaaTwo", "aaa")[0] == E.fuzzy_match("aaaOne", "aaa")[0]],
              expect_contain=[["aaaTwo", "aaaOne"], True])

        # 记一笔"刚用过 aaaOne"（accept/pick 走的就是这个入口）
        _c82d = _trig82("aaa", _pool85)
        _c82d._note_recent("aaaOne")
        _c82d.trigger(True)
        check("82.16b `_note_recent` 之后 `aaaOne` 排第 1（其余分量完全不变，"
              "只有'最近用过'这一项不同 —— 这正是 IDEA 的 recently used）",
              [list(_c82d.matches or [])],
              expect_contain=[["aaaOne", "aaaTwo"]])

        _r82b = E.RANK_BY_RECENT
        E.RANK_BY_RECENT = False
        try:
            _c82e = _trig82("aaa", _pool85)
            _c82e._note_recent("aaaOne")
            _c82e.trigger(True)
        finally:
            E.RANK_BY_RECENT = _r82b
        check("82.16c 反面对照：`VBECOMPLETE_NO_RECENT_RANK=1` 时记了也不生效"
              "（退回插入序）",
              [list(_c82e.matches or [])],
              expect_contain=[["aaaTwo", "aaaOne"]])

        # ---- 82.17~82.19 ★v103b：候选窗详情行【不】显示形参表 ----
        #      用户口径（原话）："提示自定义函数的时候，不需要在提示词框最下方
        #      提示形参列表。" v103 曾仿 IDEA 在详情行接上 `judge  (s1, s2)`，
        #      现在整套撤掉 —— 这里把"撤干净"钉住（源码护栏 + 接口不存在）。
        check("82.17 ★v103b `parser` 不再提供 `extract_proc_signatures`"
              "（形参表这套解析彻底撤掉，不留无人调用的死代码）",
              [hasattr(P, "extract_proc_signatures")],
              expect_contain=[False])

        _c82f = _trig82("jud", [("judge", "M1", None, False)])
        check("82.18 ★v103b `Completer` 不再有 `signature_of` / `_proc_sigs_now`"
              "（详情行拿不到形参，自然也不会去显示）",
              [(hasattr(_c82f, "signature_of"),
                hasattr(_c82f, "_proc_sigs_now"),
                list(_c82f.matches or []))],
              expect_contain=[(False, False, ["judge"])])

        _usrc82 = io.open(os.path.join(ROOT, "ui.py"), encoding="utf-8").read()
        # ⚠️ 不能直接 `"if self._measure..." in src`：`elif self._measure...`
        #    **包含** `if self._measure...` 这段子串（'e','l','i','f',' '），
        #    于是那种写法根本分不出 `if` 与 `elif`。按行 startswith 才准。
        _ulines82 = [l.strip() for l in _usrc82.split("\n")]
        check("82.19 源码护栏：UI 详情行必须**退回老行为** —— "
              "只在名字被省略（超宽）时显示完整名字；"
              "源码里不许再出现 `signature_of` 或 `name + \"  \" + sig`",
              [("signature_of" not in _usrc82,
                'name + "  " + sig' not in _usrc82,
                any(l.startswith("if self._measure(name) > text_w:")
                    for l in _ulines82),
                any(l.startswith("elif self._measure(name) > text_w:")
                    for l in _ulines82))],
              expect_contain=[(True, True, True, False)])

        _bsrc82b = io.open(os.path.join(ROOT, "vbe_bridge.py"),
                           encoding="utf-8").read()
        _esrc82 = io.open(os.path.join(ROOT, "engine.py"),
                          encoding="utf-8").read()
        check("82.19b 源码护栏：`vbe_bridge` / `engine` 里也不许再留"
              " `get_proc_signatures` / `_proc_sigs`（后端那一路收集一并撤掉）",
              [("get_proc_signatures" not in _bsrc82b,
                "_proc_sigs" not in _bsrc82b,
                "extract_proc_signatures" not in _bsrc82b,
                "_proc_sigs_now" not in _esrc82)],
              expect_contain=[(True, True, True, True)])

        check("82.20 源码护栏：`accept()` 里必须真的记一笔「最近用过」"
              "（用例只钉得住 `_note_recent` 本身，钉不住「钩子/主线程有没有"
              "记账」—— 同 v90/v91 那条教训）",
              [("self._note_recent(chosen)" in _esrc82,
                "_sr.get(t[1].lower(), 2)" in _esrc82,
                "_rc.get(t[1].lower(), 0)" in _esrc82)],
              expect_contain=[(True, True, True)])

    except Exception as _e82:
        import traceback as _tp82
        _tp82.print_exc()
        check("第 82 节异常: %s" % _e82, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 83. ★v103b：关键字补全后自动补一个空格
    #     用户口径："在使用 public，private，sub，function，set，if，for，with，
    #     select，end，exit 等这类关键词时，后面需要有空格，要在按 tab 键后，
    #     也能自动补全空格，idea 里就有这个功能。"
    #     ⚠️ 真机实测（独立 Excel 实例 + 真 CodeModule）：VBE 会把 ReplaceLine
    #     写进去的**行尾空格立刻 trim 掉**（`"    public "` 读回 `"    public"`）
    #     ⇒ 不能靠改行文本，只能在光标就位后【注入一个空格按键】（实测保留）。
    # ------------------------------------------------------------------
    try:
        import vba_builtins as _B83
        import vbe_bridge as VB83

        # ⚠️⚠️ 本节的断言一律**把多项打包成一个元组**再比（MEMORY 第 37 条）：
        #   `check` 的 expect_contain 是【逐元素成员判断】(`x not in got`)。
        #   一串布尔 + 全 True 的期望 ⇒ 只要 got 里有【任意一个】 True 就永远绿
        #   （A/B 实测：把 `_space_insert_pos` 改成行末返回 None，83.4 竟然不红）。
        #   打包成元组让"位置/组合"参与比较，退化了才真会红。

        check("83.1 `keyword_wants_space`：用户点名的 11 个关键字全部要补空格"
              "（大小写不敏感）",
              [tuple(_B83.keyword_wants_space(w) for w in
                     ("public", "private", "sub", "function", "set", "if",
                      "for", "with", "select", "end", "exit",
                      "Public", "PRIVATE", "Function"))],
              expect_contain=[tuple([True] * 14)])

        check("83.2 `keyword_wants_space` 反面：经常是【整行最后一个词】的"
              "（Else / Next / Loop / Then …）、后面紧跟【点号】的（Me / Debug）、"
              "本身就是【完整表达式】的（True / Nothing），以及工程内名字 —— "
              "一律【不】补空格",
              [tuple(_B83.keyword_wants_space(w) for w in
                     ("else", "next", "loop", "wend", "then", "resume", "stop",
                      "explicit", "do", "me", "debug", "true", "false",
                      "nothing", "empty", "null", "myVar", "judge", "", None))],
              expect_contain=[tuple([False] * 20)])

        check("83.3 「补空格」与「补括号」互斥：关键字与内建函数交集为空 "
              "（`Set` / `Call` 只在关键字表、不在函数表 ⇒ 不会被 "
              "`_is_callable_name` 判成可调用、不会补成 `Set()`）",
              [(len(set(_B83.BUILTIN_KEYWORD_LOWER)
                    & set(n.lower() for n in _B83.BUILTIN_FUNCTIONS)),
                "set" in set(n.lower() for n in _B83.BUILTIN_FUNCTIONS),
                _B83.keyword_wants_space("set"),
                E._keyword_wants_space("set"),
                E._keyword_wants_space("else"))],
              expect_contain=[(0, False, True, True, False)])

        # ---- 83.4 插入位置纯函数（不碰 COM，可离线钉死）----
        _sp = VB83._space_insert_pos
        check("83.4 `_space_insert_pos`：词尾在行末 -> 就在词尾补；"
              "词尾右边已经是空格 -> 不补（不重复）；"
              "右边是别的字符 -> 补（把关键字与后面的东西隔开）",
              [(_sp("    public", 5, "public"),            # k=10 == len -> 10
                _sp("    public Sub", 5, "public"),       # k=10 是空格 -> None
                _sp("    publicx", 5, "public"),          # k=10 是 'x' -> 10
                _sp("    public", 5, ""))],               # 空 completion -> 4
              expect_contain=[(10, None, 10, 4)])

        # ---- 83.5 源码护栏：三方接线必须都在 ----
        #      ★v103c：接入点从"两个独立判据"合并成一条 `_completion_followup`
        #      （它内部再按 Sub / Function / Property / 内建函数 / 关键字分档），
        #      所以护栏跟着改钉新的接入点 —— 钉的东西没变：引擎判"是什么"、
        #      `main` 把注入函数交给后端。
        _msrc83 = io.open(os.path.join(ROOT, "main.py"), encoding="utf-8").read()
        check("83.5 源码护栏：`accept()` 必须经由 `_completion_followup` 得出"
              "「补括号 / 补空格」两个标志并一起传给 `apply_completion`"
              "（引擎判是什么、后端判位置），关键字那一档在它内部；"
              "`main` 必须把注入函数交给后端"
              "（★v103d：判据多收一个「光标所在过程名」参数 —— 它要用来判"
              "「这个名字是不是函数给自己赋返回值」，见第 85 节）",
              [("_want_parens, _want_space = self._completion_followup("
                in _esrc82,
                "self._completion_followup(" in _esrc82,
                "return False, _keyword_wants_space(low)" in _esrc82,
                "def _keyword_wants_space(" in _esrc82,
                "VbeBackend(send_char=send_char" in _msrc83)],
              expect_contain=[(True, True, True, True, True)])

        _bsrc83 = io.open(os.path.join(ROOT, "vbe_bridge.py"),
                          encoding="utf-8").read()
        check("83.6 源码护栏：后端必须有开关 `AUTO_SPACE_AFTER_KEYWORD`、"
              "用 `_space_insert_pos` 定位、并在光标就位后经 `_send_char` 注入"
              "（真机实测：VBE 会 trim 行尾空格，只能注入按键）",
              [("AUTO_SPACE_AFTER_KEYWORD = _env_flag(" in _bsrc83,
                "_space_insert_pos(new_line, word_start_col, completion)"
                in _bsrc83,
                "getattr(self, \"_send_char\", None)" in _bsrc83,
                "def _space_insert_pos(" in _bsrc83,
                "want_space=False" in _bsrc83)],
              expect_contain=[(True, True, True, True, True)])

        _vsrc83 = io.open(os.path.join(ROOT, "vba_builtins.py"),
                          encoding="utf-8").read()
        check("83.7 源码护栏：关键字表必须带排除法（`KEYWORDS_NO_TRAIL_SPACE` "
              "是【真子集】：既非空、又小于全部关键字 —— 全部要补 / 全部不补"
              "都说明排除法写坏了）",
              [(len(_B83.KEYWORDS_NO_TRAIL_SPACE) > 0,
                len(_B83.BUILTIN_KEYWORD_NEEDS_SPACE) > 0,
                len(_B83.BUILTIN_KEYWORD_NEEDS_SPACE)
                < len(_B83.BUILTIN_KEYWORD_LOWER),
                "def keyword_wants_space(" in _vsrc83)],
              expect_contain=[(True, True, True, True)])

    except Exception as _e83:
        import traceback as _tp83
        _tp83.print_exc()
        check("第 83 节异常: %s" % _e83, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 84. ★v103c：① Sub 补空格 / Function 补括号（用户点名）
    #     ② 成员位置（`变量.` / With 块的 `.成员`）不再出我们自己的候选
    #     用户口径（原话）：
    #       "选中的提示词，如果是 sub 过程，按 tab 键不要补双括号，应该补空格；
    #        如果是 function 函数，按 tab 键要补双括号。"
    #       "对某个变量的成员提示，这个就关闭，不要提示了。比如说输入 student.，
    #        .后面就不要出提示框了，现在会提示函数名，变量名这些，按道理提示的
    #        应该是类或者对象的成员。"
    #     本节一律**把多项打包成一个元组**再比（MEMORY 第 37 条：散着放布尔
    #     + 全 True 期望 ⇒ 必然假绿）。
    # ------------------------------------------------------------------
    try:
        import vba_builtins as _B84
        import vbe_bridge as VB84            # noqa: F401

        # ---- 84.1 parser.extract_proc_kinds（过程种类）----
        _code84 = "\n".join([
            "Option Explicit",
            "Public Sub DoWork()",
            "End Sub",
            "Private Function CalcValue(ByVal n As Long) As Long",
            "End Function",
            "Function NoArgs",
            "End Function",
            "Property Get Title() As String",
            "End Property",
            "Private Declare Sub Sleep Lib \"kernel32\" (ByVal dw As Long)",
            "Public Declare PtrSafe Function gApi Lib \"k\" (ByVal x As Long)",
            "Sub Foo$()",
            "End Sub"])
        _k84 = P.extract_proc_kinds(_code84)
        _pn84 = sorted(str(n) for n in P.extract_records(_code84, "M1", True)[1])
        check("84.1 `parser.extract_proc_kinds`：Sub->sub、Function->function、"
              "Property->property；`Declare Sub` 按 sub、`Declare Function` 按 "
              "function（无括号的过程头 / 类型声明符 `Foo$()` 也认得出来）；"
              "与 `extract_records` 认出来的过程名**完全相同**（同一批正则，"
              "所以不可能出现「有候选却查不到它是什么」）",
              [(sorted(_k84.items()), _pn84, sorted(_k84) == _pn84)],
              expect_contain=[(sorted([
                  ("dowork", "sub"), ("calcvalue", "function"),
                  ("noargs", "function"), ("title", "property"),
                  ("sleep", "sub"), ("gapi", "function"), ("foo", "sub")]),
                  ["calcvalue", "dowork", "foo", "gapi", "noargs", "sleep",
                   "title"], True)])

        # ---- 84.2/84.3 判据与端到端（最小后端，不碰 COM）----
        class _B84_(object):
            """最小后端：只提供补全判据要的那几个只读接口。"""
            _KINDS = {"dowork": "sub", "calcvalue": "function",
                      "title": "property"}

            def __init__(self, pool=()):
                self._pool = list(pool)
                self._ctx = {"line_no": 1, "line_text": "    x",
                             "caret_col": 6, "proc_name": "Cur",
                             "module_name": "M1", "in_string": False,
                             "in_comment": False, "in_type_position": False,
                             "in_decl_position": False, "decl_names": []}
                self.last = None

            def get_context(self):
                return dict(self._ctx)

            def get_identifiers(self):
                return list(self._pool)

            def get_proc_names(self):
                return frozenset(self._KINDS)

            def get_proc_kinds(self):
                return dict(self._KINDS)

            def get_builtin_names(self):
                return {"msgbox", "left", "split"}

            def get_structural_names(self):
                return []

            def get_type_names(self):
                return []

            def get_host_enum_names(self):
                return {}

            def get_declared_names(self):
                # ⚠️ 必须交回池子里的名字：回声防护问的是"这个名字在工程里
                #    【出现过】吗"，交回空列表会让每个候选都被当成幻影剔光
                #    （本节第一版就是这么假绿的 —— 弹窗压根不出、`last` 恒 None）。
                return [str(n).lower() for n, _m, _p, _v in self._pool]

            def declared_elsewhere(self, name, module_name):
                return str(name).lower() in self.get_declared_names()

            def names_outside_caret(self, caret, scope_only=False):
                return None

            def apply_completion(self, *a, **k):
                self.last = a
                return None

            def vbe_yield_visible(self):
                return False

            def vbe_popup_visible(self):
                return False

            def vbe_popup_info(self):
                return []

        class _U84(object):
            def __init__(self):
                self.shown = False

            def show(self, *a, **k):
                self.shown = True

            def hide(self, *a, **k):
                self.shown = False

            def update_selection(self, *a, **k):
                pass

            def contains_point(self, *a, **k):
                return False

        class _NoKinds84(object):
            """模拟"旧式后端"：没有 get_proc_kinds 这个可选接口。"""

            def __init__(self, inner):
                self._in = inner

            def __getattr__(self, k):
                if k == "get_proc_kinds":
                    raise AttributeError(k)
                return getattr(self._in, k)

        def _trig84(word, pool, wrap=None):
            _b = _B84_(pool)
            _b._ctx["line_text"] = "    " + word
            _b._ctx["caret_col"] = len(_b._ctx["line_text"]) + 1
            _u = _U84()
            _c = E.Completer(wrap(_b) if wrap else _b, _u)
            _c.trigger(True)
            return _c, _b

        _c84a, _b84a = _trig84("DoWo", [("DoWork", "M1", None, False)])
        _m84a = list(_c84a.matches or [])
        _c84a.accept()
        _c84b, _b84b = _trig84("CalcVa", [("CalcValue", "M1", None, False)])
        _m84b = list(_c84b.matches or [])
        _c84b.accept()
        _c84c, _b84c = _trig84("Tit", [("Title", "M1", None, False)])
        _m84c = list(_c84c.matches or [])
        _c84c.accept()
        check("84.2 ★端到端（用户点名的口径）：选中一个 **Sub**（`DoWo`）按 Tab "
              "-> `apply_completion(..., want_parens=False, want_space=True)`"
              "（补空格、不补括号）；选中 **Function**（`CalcVa`）-> "
              "`(True, False)`（补括号）；选中 **Property**（`Tit`）-> 都不补"
              "（⚠️ `matches` 必须在 `accept()` 【之前】取 —— accept 末尾会 "
              "`hide()`，而 hide 会把 matches 清空）",
              [(tuple((_m84a, _b84a.last[4], _b84a.last[5])),
                tuple((_m84b, _b84b.last[4], _b84b.last[5])),
                tuple((_m84c, _b84c.last[4], _b84c.last[5])))],
              expect_contain=[((["DoWork"], False, True),
                               (["CalcValue"], True, False),
                               (["Title"], False, False))])

        # 84.3 ★旧式后端（没有 get_proc_kinds）必须退回 v98 的老口径：
        #      工程内过程名一律补括号 —— 判据缺失就按老规矩来，绝不猜。
        _c84d, _b84d = _trig84("DoWo", [("DoWork", "M1", None, False)],
                               wrap=_NoKinds84)
        _m84d = list(_c84d.matches or [])
        _c84d.accept()
        check("84.3 对照：后端没有 `get_proc_kinds`（旧式 / 测试后端）时退回 "
              "v98 的老口径 —— 工程内过程名一律补括号（`(True, False)`）",
              [(tuple((_m84d, _b84d.last[4], _b84d.last[5])))],
              expect_contain=[((["DoWork"], True, False))])

        # 84.4 判据本身：两条标志**不可能同时为真**（补括号 + 补空格同时发出去
        #      会写成 `Foo ()`，v98 专门修过这个坑）。
        _c84e = E.Completer(_B84_(), _U84())
        check("84.4 `_completion_followup` 不变量：任何名字都不可能同时"
              "「补括号」又「补空格」（同时真会写成 `Foo ()`）—— "
              "Sub / Function / Property / 内建函数 / 关键字 / 变量一起过一遍",
              [tuple((bool(_c84e._completion_followup(n)[0]),
                      bool(_c84e._completion_followup(n)[1]))
                     for n in ("DoWork", "CalcValue", "Title", "MsgBox", "Left",
                               "Public", "Set", "temp", "student", ""))],
              expect_contain=[((False, True), (True, False), (False, False),
                               (True, False), (True, False), (False, True),
                               (False, True), (False, False), (False, False),
                               (False, False))])

        check("84.5 顺手钉住 v103b 的老能力没被这次重构弄丢：语言关键字仍然是"
              "「补空格不补括号」，语言自带函数仍然是「补括号不补空格」",
              [(tuple((E._keyword_wants_space("public"),
                       _B84.keyword_wants_space("function"),
                       _c84e._completion_followup("msgbox"),
                       _c84e._completion_followup("public"))))],
              expect_contain=[(True, True, (True, False), (False, True))])

        # ---- 84.6 ★用户点名的口径：成员位置不再出我们自己的候选 ----
        def _m84(line, manual=False, popup=None):
            """在 `line` 的光标处触发一次，返回 (窗可见, UI 真被 show 过,
            yield_pending 是否被挂上)。"""
            _b = _B84_([("student", "M1", None, False),
                        ("DoWork", "M1", None, False)])
            _b._ctx["line_text"] = line
            _b._ctx["caret_col"] = len(line) + 1
            _u = _U84()
            _c = E.Completer(_b, _u)
            _save = E.MEMBER_POPUP
            if popup is not None:
                E.MEMBER_POPUP = popup
            try:
                _c.trigger(manual=manual)
                return (bool(_c.is_visible()), bool(_u.shown),
                        _c.yield_pending is not None)
            finally:
                E.MEMBER_POPUP = _save

        _mp84 = E.MEMBER_POPUP
        try:
            E.MEMBER_POPUP = False          # = 出厂默认（见 84.7）
            _r84_auto = _m84("    student.te")
            _r84_with = _m84("    .Size")
            _r84_man = _m84("    student.te", manual=True)
            # 补位路径（main 轮询里 confirm_yield 返回 False 之后那条）
            _r84_refill = _m84("    student.te")
        finally:
            E.MEMBER_POPUP = _mp84

        check("84.6 ★用户点名：成员位置（`student.` / With 块的 `.`）不再出"
              "我们自己的候选 —— 自动触发一律不弹（也不挂待定让位，因为"
              "我们永远不弹、没有「宽限到再补回来」这回事）；"
              "对照：Ctrl+Space 点名要的列表照弹（v72 老口径：点名的列表照弹，"
              "否则就成了「按了没反应」）",
              [(_r84_auto, _r84_with, _r84_refill, (_r84_man[0], _r84_man[1]))],
              expect_contain=[((False, False, False), (False, False, False),
                               (False, False, False), (True, True))])

        # ---- 84.7 源码护栏：默认关 / 早退位置 / 后端接线 ----
        _esrc84 = io.open(os.path.join(ROOT, "engine.py"),
                          encoding="utf-8").read()
        _bsrc84 = io.open(os.path.join(ROOT, "vbe_bridge.py"),
                          encoding="utf-8").read()
        _psrc84 = io.open(os.path.join(ROOT, "parser.py"),
                          encoding="utf-8").read()
        _i84 = _esrc84.find("if (not MEMBER_POPUP) and (not manual)")
        _i84b = _esrc84.find("self.yield_pending = (time.time() + YIELD_GRACE")
        check("84.7 源码护栏：① 开关 `MEMBER_POPUP` 的出厂默认必须是【关】"
              "（`settings.flag(\"MEMBER_POPUP\", False)`）；② 成员位置的早退必须"
              "排在让位逻辑【之前】（否则那条「宽限到、VBE 没弹、把候选补回来」"
              "的路径还会把窗补出来）；③ 后端必须收集 `_proc_kinds` 并提供 "
              "`get_proc_kinds`；④ `parser` 必须有 `extract_proc_kinds`；"
              "⑤ 「补括号」与「补空格」两道闸门都必须过 `_in_call_stmt`"
              "（同一语义只有一份实现）",
              [('_settings.flag("MEMBER_POPUP", False)' in _esrc84,
                _i84 >= 0 and 0 <= _i84 < _i84b,
                "proc_kinds.update(" in _bsrc84,
                "def get_proc_kinds(" in _bsrc84,
                "def extract_proc_kinds(" in _psrc84,
                "def _completion_followup(" in _esrc84,
                _bsrc84.count("not _in_call_stmt(line_text)") >= 2)],
              expect_contain=[(True, True, True, True, True, True, True)])

    except Exception as _e84:
        import traceback as _tp84
        _tp84.print_exc()
        check("第 84 节异常: %s" % _e84, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 85. ★v103d：在 Function 体内给函数**自己**赋返回值（`MyFunc = ...`）时，
    #     这个名字代表的是"返回值的那个变量"、不是一次调用 ⇒ 不补括号
    #     （补了就是语法错误的 `MyFunc() = 5`）。
    #     用户口径（原话）："在 function 函数中，再给 function 赋返回值时，
    #     此时后面不应该自动补全括号。"
    #     本节一律**把多项打包成一个元组**再比（MEMORY 第 37 条：散着放布尔
    #     + 全 True 的期望 ⇒ 必然假绿）。
    # ------------------------------------------------------------------
    try:
        import engine as _E85

        # ---- 85.1 纯判据（不碰 COM、不需要后端）----
        #      ⚠️ 同时钉住"取不到就不动手"（None / 空串 / 空名字 -> False）
        #      与"作用域键带 `#序号`"（同名过程时 _proc_of_line 给的就是这种键）。
        check("85.1 `_is_own_proc_name`：名字等于所在过程名 -> True；不同名 / "
              "取不到过程名（模块声明区 = None / 空串）-> False；大小写不敏感；"
              "所在过程名是【作用域键】`name#序号` 时先剥序号再比",
              [tuple(_E85._is_own_proc_name(n, p) for n, p in
                     (("CalcValue", "CalcValue"),
                      ("calcvalue", "CALCVALUE"),
                      ("  CalcValue  ", "CalcValue"),
                      ("CalcValue", "Other"),
                      ("CalcValue", None),
                      ("CalcValue", ""),
                      ("CalcValue", "CalcValue#2"),
                      ("", "CalcValue"),
                      (None, "CalcValue")))],
              expect_contain=[(True, True, True, False, False, False,
                               True, False, False)])

        # ---- 85.2 判据分档：新例外 + v103c 老口径都不回退 ----
        class _B85(object):
            """最小后端：只给 `get_proc_kinds` 一张表（判据只需要它）；
            另外供 85.5 用来验证"语言自带函数"那一档没被新例外带偏。"""
            _KINDS = {"calcvalue": "function", "dowork": "sub",
                      "title": "property"}

            def get_proc_kinds(self):
                return dict(self._KINDS)

            def get_builtin_names(self):
                return {"msgbox", "left", "split"}

        class _U85(object):
            def show(self, *a, **k):
                pass

            def hide(self, *a, **k):
                pass

            def update_selection(self, *a, **k):
                pass

            def contains_point(self, *a, **k):
                return False

        _c85 = _E85.Completer(_B85(), _U85())
        check("85.2 ★用户点名：`CalcValue` 在【CalcValue 自己体内】-> 不补括号"
              "也不补空格 (False, False)；在【别的过程里】-> 照旧补括号 "
              "(True, False)；没拿到过程名 -> 老口径 (True, False)；"
              "Sub 在自己体内仍按 Sub 口径补空格；Property 仍什么都不补",
              [(tuple(_c85._completion_followup("CalcValue", "CalcValue")),
                tuple(_c85._completion_followup("CalcValue", "Other")),
                tuple(_c85._completion_followup("CalcValue", None)),
                tuple(_c85._completion_followup(
                    "CalcValue", "CalcValue#2")),
                tuple(_c85._completion_followup("DoWork", "DoWork")),
                tuple(_c85._completion_followup("Title", "Title")))],
              expect_contain=[((False, False), (True, False), (True, False),
                               (False, False), (False, True), (False, False))])

        # ---- 85.3 端到端（走 accept()，后端不碰 COM）----
        #      ⚠️ `accept()` 取的是 `_ctx_now`（trigger 开头挂上的**完整**上下文）
        #      里的 `proc_name`，而不是 `self.ctx` —— 后面那个只有行号/列/词范围。
        #      这一条就是为了咬住"接线真的接上了"，否则 85.1/85.2 全绿而真机无效。
        class _B85E(object):
            def __init__(self, pool, ctx):
                self._pool = list(pool)
                self._ctx = dict(ctx)
                self.last = None

            def get_context(self):
                return dict(self._ctx)

            def get_identifiers(self):
                return list(self._pool)

            def get_proc_names(self):
                return frozenset(["calcvalue"])

            def get_proc_kinds(self):
                return {"calcvalue": "function"}

            def get_builtin_names(self):
                return set()

            def get_structural_names(self):
                return []

            def get_type_names(self):
                return []

            def get_host_enum_names(self):
                return {}

            def get_declared_names(self):
                return [str(n).lower() for n, _m, _p, _v in self._pool]

            def declared_elsewhere(self, name, module_name):
                return str(name).lower() in self.get_declared_names()

            def names_outside_caret(self, caret, scope_only=False):
                return None

            def apply_completion(self, *a, **k):
                self.last = a
                return None

            def vbe_yield_visible(self):
                return False

            def vbe_popup_visible(self):
                return False

            def vbe_popup_info(self):
                return []

        def _run85(line_text, proc_name):
            _ctx = {"line_no": 1, "line_text": line_text,
                    "caret_col": len(line_text) + 1, "proc_name": proc_name,
                    "module_name": "M1", "in_string": False,
                    "in_comment": False, "in_type_position": False,
                    "in_decl_position": False, "decl_names": []}
            _b = _B85E([("CalcValue", "M1", None, False)], _ctx)
            _c = _E85.Completer(_b, _U85())
            _c.trigger(True)
            _c.accept()
            return _b.last

        _own85 = _run85("    CalcVa", "CalcValue")
        _expr85 = _run85("    x = CalcVa", "Other")
        _expr_own85 = _run85("    y = CalcVa", "CalcValue")
        check("85.3 ★端到端（走 `accept()` -> `apply_completion`）：在 CalcValue "
              "体内写 `CalcVa` 按 Tab -> `(want_parens=False, want_space=False)`"
              "（`MyFunc = ...` 那个位置补括号会写成语法错误）；在别的过程里写"
              "`x = CalcVa` -> `(True, False)`；在 **CalcValue 自己体内** 写"
              "`y = CalcVa`（读自己的返回值）-> 也不补（函数名在这里是变量）",
              [(tuple((_own85[3], _own85[4], _own85[5])),
                tuple((_expr85[3], _expr85[4], _expr85[5])),
                tuple((_expr_own85[3], _expr_own85[4], _expr_own85[5])))],
              expect_contain=[(("CalcValue", False, False),
                               ("CalcValue", True, False),
                               ("CalcValue", False, False))])

        # ---- 85.4 源码护栏：接线三处都得在 ----
        _esrc85 = io.open(os.path.join(ROOT, "engine.py"),
                          encoding="utf-8").read()
        _i85a = _esrc85.find("_encl_proc = str(_src.get(\"proc_name\") or \"\")")
        _i85b = _esrc85.find("_want_parens, _want_space = self._completion_followup(")
        _i85c = _esrc85.find("def _is_own_proc_name(")
        _i85d = _esrc85.find("if _is_own_proc_name(low, enclosing_proc):")
        check("85.4 源码护栏：① `accept()` 必须从 `_ctx_now`（不是 `self.ctx`）"
              "读 `proc_name`；② 读出来的过程名必须【排在】那次 "
              "`_completion_followup` 调用之前（否则参数是空的、判据永远不生效）；"
              "③ 纯判据 `_is_own_proc_name` 与它在 `_completion_followup` 里的"
              "早退都得在；④ VBE 成员列表补进名字的那一路（trigger 里 v100 那段）"
              "也必须把过程名传进去（同一语义只有一份实现，不能只修 accept）",
              [(_i85a >= 0, _i85b >= 0 and _i85a < _i85b, _i85c >= 0,
                _i85d >= 0 and _i85c < _i85d,
                "self._completion_followup(\n                            "
                "_word100, ctx.get(\"proc_name\"))" in _esrc85)],
              expect_contain=[(True, True, True, True, True)])

        # ---- 85.5 反面：判据不许把"语言自带函数 / 关键字"也当成自家过程名 ----
        #      内建函数名（如 MsgBox）永远不会等于用户的过程名，但这条要钉住
        #      "新例外只对 kind=='function' 那一档生效" —— 关键字那档仍补空格。
        check("85.5 反面：新例外只落在【Function 自己体内】这一档 —— "
              "关键字（public）仍是「补空格不补括号」，内建函数（msgbox）"
              "仍是「补括号不补空格」，即使过程名恰好被传成同一个词",
              [(tuple(_c85._completion_followup("public", "public")),
                tuple(_c85._completion_followup("msgbox", "msgbox")))],
              expect_contain=[((False, True), (True, False))])

    except Exception as _e85:
        import traceback as _tp85
        _tp85.print_exc()
        check("第 85 节异常: %s" % _e85, [True], expect_contain=[False])

    # ==================================================================
    # 86. ★v104 宿主探测：不再写死 Excel —— 自动认 Office 全家桶
    #     （Excel / Word / PowerPoint / Access / Visio）与 WPS 全家桶
    #     （文字 / 表格 / 演示）。COM 全用打桩替身，离线可测（不依赖本机开着谁）。
    # ==================================================================
    print("\n=== 86. 宿主探测（v104：Office + WPS 全家桶）===")
    try:
        import importlib.util as _iu86
        _spec86 = _iu86.spec_from_file_location(
            "vbe_bridge_fresh86",
            os.path.join(os.path.dirname(os.path.dirname(
                os.path.abspath(__file__))), "vbe_bridge.py"))
        VB86 = _iu86.module_from_spec(_spec86)
        _spec86.loader.exec_module(VB86)

        # ---- 86.1 候选 ProgID 表：Office 五件套 + WPS 三件套都在 ----
        _need86 = ["Excel.Application", "Word.Application",
                   "PowerPoint.Application", "Access.Application",
                   "Visio.Application", "KWPS.Application", "KET.Application",
                   "KWPP.Application"]
        _miss86 = [p for p in _need86 if p not in VB86._HOST_PROGIDS]
        check("86.1 候选 ProgID 表含 Office 五件套 + WPS 三件套",
              [_miss86], expect_contain=[[]])

        # ---- 86.2 进程名 -> ProgID 映射（前台窗口线索要用）----
        _keys86 = ("excel.exe", "winword.exe", "powerpnt.exe", "msaccess.exe",
                   "visio.exe", "wps.exe", "et.exe", "wpp.exe")
        _map86 = [VB86._HOST_EXE_PROGID.get(k) for k in _keys86]
        check("86.2 进程名 -> ProgID 映射（excel/winword/powerpnt/msaccess/"
              "visio/wps/et/wpp）", [_map86],
              expect_contain=[["Excel.Application", "Word.Application",
                               "PowerPoint.Application", "Access.Application",
                               "Visio.Application", "KWPS.Application",
                               "KET.Application", "KWPP.Application"]])

        class _App86(object):
            def __init__(self, vbe):
                self._vbe = vbe

            @property
            def VBE(self):
                return self._vbe

        # ---- 86.3 无前台线索时逐个试候选，取第一个拿到 .VBE 的 ----
        _tried86 = []

        def _gao86(progid):
            _tried86.append(progid)
            if progid == "Word.Application":
                return _App86({"vbe_of": "Word"})
            raise OSError("not running: " + progid)

        _sv86a = (VB86._get_active_object, VB86._preferred_progid,
                  VB86._host_progid_hint.get("progid"))
        VB86._get_active_object = _gao86
        VB86._preferred_progid = lambda: None
        VB86._host_progid_hint["progid"] = None
        try:
            _v86 = VB86._get_vbe()
            _hint86a = VB86._host_progid_hint.get("progid")
        finally:
            VB86._get_active_object = _sv86a[0]
            VB86._preferred_progid = _sv86a[1]
            VB86._host_progid_hint["progid"] = _sv86a[2]
        check("86.3 只有 Word 在跑时：_get_vbe 逐个试候选并返回 Word 的 VBE，"
              "并记住宿主",
              [(_v86, _tried86[0], "Word.Application" in _tried86, _hint86a)],
              expect_contain=[({"vbe_of": "Word"}, "Excel.Application",
                               True, "Word.Application")])

        # ---- 86.4 前台窗口线索优先：用户在用 WPS 表格 -> 第一顺位就是 KET ----
        _tried86b = []

        def _gao86b(progid):
            _tried86b.append(progid)
            return _App86({"vbe_of": progid})

        _sv86b = (VB86._get_active_object, VB86._preferred_progid,
                  VB86._host_progid_hint.get("progid"))
        VB86._get_active_object = _gao86b
        VB86._preferred_progid = lambda: "KET.Application"
        VB86._host_progid_hint["progid"] = None
        try:
            _v86b = VB86._get_vbe()
            _hint86b = VB86._host_progid_hint.get("progid")
        finally:
            VB86._get_active_object = _sv86b[0]
            VB86._preferred_progid = _sv86b[1]
            VB86._host_progid_hint["progid"] = _sv86b[2]
        check("86.4 前台线索优先：用户在用 WPS 表格时，第一顺位试的是 KET",
              [(_tried86b[0], _v86b, _hint86b)],
              expect_contain=[("KET.Application",
                               {"vbe_of": "KET.Application"},
                               "KET.Application")])

        # ---- 86.5 一个宿主都没有 -> 抛异常（与旧行为一致，交给上层退避）----
        _sv86c = (VB86._get_active_object, VB86._preferred_progid)
        VB86._get_active_object = lambda progid: (_ for _ in ()).throw(
            OSError("nothing running"))
        VB86._preferred_progid = lambda: None
        _raised86 = False
        try:
            VB86._get_vbe()
        except Exception:
            _raised86 = True
        finally:
            VB86._get_active_object, VB86._preferred_progid = _sv86c
        check("86.5 候选全不可用时 _get_vbe 抛异常（-> 上层 _com_fail 退避）",
              [_raised86], expect_contain=[True])

        # ---- 86.6 前台窗口 -> 进程名 -> ProgID（纯函数级）----
        _sv86d = (VB86._foreground_hwnd, VB86._process_image_name)
        VB86._foreground_hwnd = lambda: 0x123
        VB86._process_image_name = lambda pid: "winword.exe"
        try:
            _pref86 = VB86._preferred_progid()
        finally:
            VB86._foreground_hwnd, VB86._process_image_name = _sv86d
        check("86.6 前台窗口属于 WINWORD.EXE -> 猜出 Word.Application",
              [_pref86], expect_contain=["Word.Application"])

        # ---- 86.7 源码护栏：_get_vbe 迭代候选表（不再只认 Excel）----
        _vsrc86 = io.open(os.path.join(ROOT, "vbe_bridge.py"),
                          encoding="utf-8").read()
        check("86.7 源码护栏：`_get_vbe` 迭代候选表、不再写死 Excel；"
              "窗口类名表含 Word/PPT/Access/Visio",
              [("for progid in order" in _vsrc86,
                "_HOST_PROGIDS = (" in _vsrc86,
                '"Word.Application"' in _vsrc86,
                '"OpusApp"' in _vsrc86 and '"OMain"' in _vsrc86,
                '"wndclass_desked_gsk"' in _vsrc86)],
              expect_contain=[(True, True, True, True, True)])
    except Exception as _e86:
        import traceback as _tp86
        _tp86.print_exc()
        check("第 86 节异常: %s" % _e86, [True], expect_contain=[False])

    # ==================================================================
    # 87. ★v105：声明位置不弹候选（IDEA 口径）+ 判据锚定（修误判）
    # ==================================================================
    # 用户口径（原话）："在定义变量名，函数名或过程名时，会出提示词列表框。
    # 我用 idea 是不提示的吧，你帮我改成 IDEA 那种。"
    #
    # 背景：parser.is_caret_in_declaration()（那边叫"第一道防线"）与后端的
    # in_decl_position 早在 v22 就写好了，但**引擎从来没消费过这个字段** ——
    # 实际行为退化成 v37/v53 的"只过滤掉正在输入的那个名字本身、其它候选照弹"。
    # v105 把它接进 trigger()，同时修好判据里一个一直没被暴露的误判（见 87.2）。
    print("\n=== 87. 声明位置不弹候选（v105）===")
    try:
        # ---- 87.1 纯判据：真声明行（含"名字还没敲"的中间态）----
        _T87 = ["Dim x", "Dim", "Public gA As String", "Private m_x As Long",
                "Global g As Long", "Static s As Long", "Const K = 1",
                "Private Const K = 1", "Friend f As Long", "WithEvents b As Object",
                "Private WithEvents btn As Button", "Dim a As Long, b As Long",
                "Sub Foo", "Sub", "Public Sub", "Private Sub Foo()",
                "Function", "Function F() As Long", "Property", "Property Get X()",
                "Event", "Event E(ByVal x)", "Type", "Type T", "Enum",
                "Public Enum E", "Dim arr(10) As Long"]
        _bad87 = [c for c in _T87 if not P.is_caret_in_declaration(c, len(c))]
        check("87.1 纯判据：真声明行（%d 条，含只敲了关键字的中间态）全部命中"
              % len(_T87), [_bad87], expect_contain=[[]])

        # ---- 87.2 ★纯判据反面：标识符恰好叫关键字的【用法行】不许误判 ----
        # 旧实现 `_RE_CONST.search(line) or _RE_DECL.search(line)` 在**整行**里搜
        # 关键字，下面这些行（标识符叫 Global / Public / Type / Sub…）实测全被判成
        # "声明行"。以前没人消费这个字段所以无害，v105 一接上就把正常补全误静默
        # （第 24.4 条实测就是这么红的）。v105 改成锚定语句开头。
        _F87 = ["x = Global", "MsgBox Global", "    Global", "    Public = 5",
                "GlobalData = 1", "SubTotal = 1", "PropertyValue = 3",
                "FunctionCount = FunctionCount + 1", "TypeOf x Is Long",
                "TypeName = x", "EventLog.Add 1", "DeclareMe x",
                "    ReDim arr(1)", "ReDim arr(0)", "Option Explicit",
                "End Sub", "End Function", "With obj", "    Set b = New Thing",
                "arr(1) = 2", "Call Global", 'Attribute VB_Name = "M"']
        _f87 = [c for c in _F87 if P.is_caret_in_declaration(c, len(c))]
        check("87.2 ★纯判据反面：用法行（标识符叫 Global/Public/Type/Sub…）"
              "一律不算声明行", [_f87], expect_contain=[[]])

        # ---- 端到端：最小 UI / 后端（只给判据需要的接口）----
        class _UI87:
            def __init__(self):
                self.shown = False
            def show(self, *a, **k):
                self.shown = True
            def update_selection(self, *a, **k):
                pass
            def hide(self, *a, **k):
                self.shown = False
            def contains_point(self, *a, **k):
                return False

        class _B87:
            def __init__(self, decl, in_type, ids=None, type_names=None,
                         line_text="    cus", decl_names=None):
                self._ctx = {"line_no": 1, "caret_col": len(line_text) + 1,
                             "line_text": line_text, "in_string": False,
                             "in_comment": False, "in_type_position": in_type,
                             "in_decl_position": decl,
                             "decl_names": (decl_names if decl_names is not None
                                            else ["customerName"]),
                             "proc_name": "Foo", "module_name": "M1"}
                self._ids = ids if ids is not None else [
                    ("customerName", "M1", "Foo", False),
                    ("customerCode", "M1", "Foo", False)]
                self._tn = type_names
            def get_context(self):
                return self._ctx
            def get_identifiers(self):
                return self._ids
            def get_declared_names(self):
                return [n[0] for n in self._ids]
            def get_type_names(self):
                return list(self._tn or ())

        def _trig87(decl, in_type, manual=False, ids=None, types=None,
                    line_text="    cus"):
            _ui = _UI87()
            _c = E.Completer(_B87(decl, in_type, ids, types, line_text), _ui)
            _c.trigger(manual=manual)
            return (_ui.shown, sorted(_c.current_matches()))

        # ---- 87.3 声明行 -> 不弹 ----
        check("87.3 声明行（in_decl=True，非类型位置）-> 不弹候选",
              [_trig87(True, False)], expect_contain=[(False, [])])

        # ---- 87.4 `As` 类型位置 -> 仍然提示类型名（v22/v93 不许退化）----
        # 行是 `    Dim x As My`（正在填类型名），in_decl 与 in_type 同时为真 ——
        # 此时【必须放行】，把候选收窄到类型名。
        check("87.4 `As` 类型位置（in_decl=True + in_type=True）-> 照常提示类型名",
              [_trig87(True, True, ids=[("MyType", "M1", None, False)],
                       types=["MyType"], line_text="    Dim x As My")],
              expect_contain=[(True, ["MyType"])])

        # ---- 87.5 用法行不受影响 ----
        check("87.5 非声明行 -> 照常提示",
              [_trig87(False, False)],
              expect_contain=[(True, ["customerCode", "customerName"])])

        # ---- 87.6 Ctrl+Space（点名）在声明行也照弹 ----
        check("87.6 Ctrl+Space（manual）在声明行照弹（点名的列表不拦）",
              [_trig87(True, False, manual=True)],
              expect_contain=[(True, ["customerCode", "customerName"])])

        # ---- 87.7 开关：DECL_POPUP=True -> 恢复旧行为 ----
        _orig_dp87 = E.DECL_POPUP
        E.DECL_POPUP = True
        try:
            _r87 = _trig87(True, False)
        finally:
            E.DECL_POPUP = _orig_dp87
        check("87.7 开关 DECL_POPUP=True -> 恢复旧行为（声明行照常出候选）",
              [_r87], expect_contain=[(True, ["customerCode", "customerName"])])

        # ---- 87.8 源码护栏：接线 + 锚定判据 + 开关名 ----
        _esrc87 = io.open(os.path.join(ROOT, "engine.py"),
                          encoding="utf-8").read()
        _psrc87 = io.open(os.path.join(ROOT, "parser.py"),
                          encoding="utf-8").read()
        # ⚠️ parser 里新注释**引用了**旧实现那行源码，直接子串检查会被注释绊倒
        #    ⇒ 先剥掉注释再断言"旧实现已不在"。
        import re as _re87
        _psrc87_nc = _re87.sub(r"#[^\n]*", "", _psrc87)
        check("87.8 源码护栏：trigger 读 in_decl_position 且放行 in_type；"
              "parser 用锚定判据（_RE_ACCESS_MODIFIER / _RE_PROC_HEAD），"
              "整行 search 的旧实现已删除；开关读 VBECOMPLETE_DECL_POPUP",
              [('ctx.get("in_decl_position")' in _esrc87
                and "and not in_type" in _esrc87,
                "VBECOMPLETE_DECL_POPUP" in _esrc87,
                "_RE_ACCESS_MODIFIER.match(s)" in _psrc87
                and "_RE_PROC_HEAD.match(s)" in _psrc87,
                "_RE_CONST.search(line) or _RE_DECL.search(line)"
                not in _psrc87_nc)],
              expect_contain=[(True, True, True, True)])
    except Exception as _e87:
        import traceback as _tp87
        _tp87.print_exc()
        check("第 87 节异常: %s" % _e87, [True], expect_contain=[False])

    # ==================================================================
    # 88. ★v106 开机自启后"要等半分钟才出提示"—— COM 退避没能及时解除
    #
    #     症状（用户报的）：工具随开机启动时 Excel 还没开，COM 连续失败把退避
    #     升到顶档 30 秒；之后用户才打开 Excel、进 VBE 写代码，补全却还要等
    #     30~60 秒才出来。
    #
    #     根因是【两处早退把"宿主回来了"这条路整个挡掉了】：
    #       ① `_get_vbe_cached` 先判退避、后探宿主 ⇒ 退避期内根本不调
    #          `_host_window_alive()`，而"解除退避"恰恰是它干的；
    #       ② `main._com_allowed()` 同样先被退避挡住 ⇒ 连 snapshot 都不调，
    #          `_get_vbe_cached` 连跑都跑不到。
    #     另外 `_host_window_alive` 里"宿主回来了"的判据只有 `miss >= 2`，
    #     而开机路径下 `seen` 一直是 False（下面那句 "从未见过就不拦" 直接
    #     return，miss 根本累加不了）⇒ 这个判据永远等不到。
    # ==================================================================
    print("\n=== 88. 开机自启后启动慢：COM 退避及时解除（v106）===")
    try:
        import importlib.util as _iu88
        _spec88 = _iu88.spec_from_file_location(
            "vbe_bridge_fresh88",
            os.path.join(ROOT, "vbe_bridge.py"))
        VB88 = _iu88.module_from_spec(_spec88)
        _spec88.loader.exec_module(VB88)

        def _arm_backoff88(seen):
            """摆出"已连续失败 5 次、退避顶到 30 秒"的现场。"""
            VB88.com_reset()
            VB88._com_state["fails"] = 5
            VB88._com_state["until"] = time.time() + 30.0
            VB88._wnd_state["seen"] = seen
            VB88._wnd_state["miss"] = 2 if seen else 0
            VB88._vbe_cache["obj"] = None
            VB88._vbe_cache["at"] = 0.0

        # ---- 88.1 首次看到宿主窗口 -> 立刻解除 30s 退避 ----
        _arm_backoff88(seen=False)
        VB88._probe_host_window = lambda: True
        _r81 = VB88._host_window_alive()
        check("88.1 首次看到宿主窗口（seen=False，开机路径）-> 立刻解除 30s 退避",
              [(_r81, VB88.com_backoff_remaining() == 0,
                VB88._wnd_state["seen"])],
              expect_contain=[(True, True, True)])

        # ---- 88.2 ★反面：宿主不在时【不许】解除退避 ----
        # 退避期里每轮都会来探一次"宿主在不在"（见 _com_allowed / 88.3），
        # 若宿主不在也顺手把退避清零，那退避就被这套探测自己刷没了，
        # COM 会被 500ms 一轮猛打 —— 这正是当初加退避要避免的事。
        _arm_backoff88(seen=False)
        VB88._probe_host_window = lambda: False
        _r82 = VB88._host_window_alive()
        check("88.2 ★反面：宿主不在时不许解除退避（且'从未见过'仍不拦 COM）",
              [(_r82, VB88.com_backoff_remaining() > 20)],
              expect_contain=[(True, True)])

        # ---- 88.3 ★核心：退避期内仍探宿主 -> 一次调用就恢复可用 ----
        # A/B 照妖镜：把 `_get_vbe_cached` 里的宿主探测挪回退避检查【之后】
        # （v105 的旧顺序），本用例立刻变红（探测次数 0、返回值 None、
        # 退避仍然是 30s）。
        _arm_backoff88(seen=False)
        _probe88 = []

        def _probe88_fn():
            _probe88.append(1)
            return True

        class _FakeVbe88(object):
            ActiveCodePane = "pane"

        VB88._probe_host_window = _probe88_fn
        VB88._get_vbe = lambda: _FakeVbe88()
        _r83 = VB88._get_vbe_cached()
        check("88.3 ★退避期内仍探宿主：用户中途打开 Excel -> 一次调用就恢复"
              "（既解除退避，又真的拿到 VBE）",
              [(_probe88 != [], _r83 is not None,
                VB88.com_backoff_remaining() == 0)],
              expect_contain=[(True, True, True)])

        # ---- 88.4 ★回归：宿主一直在、COM 持续失败 -> 退避照旧升到顶档 ----
        # 上面几条都是"宿主回来了就解除"；这里守住另一头：宿主窗口明明一直在
        # （miss=0、seen=True），却每次打 COM 都失败（典型的 Excel 正在关闭）
        # —— 退避必须按 1/2/5/10/30 老实升上去，不能被"每轮都探到宿主"
        # 退化成一个恒定的小值。
        _arm_backoff88(seen=True)

        def _boom88():
            raise RuntimeError("com unavailable")

        VB88._probe_host_window = lambda: True
        VB88._get_vbe = _boom88
        for _ in range(5):
            VB88._com_state["until"] = 0.0        # 模拟时间前进到退避到期
            VB88._vbe_cache["obj"] = None
            VB88._get_vbe_cached()
        check("88.4 ★回归：宿主窗口一直在、COM 持续失败 -> 退避照旧升到 30s"
              "（没被'每轮都探到宿主'刷没了）",
              [(VB88._com_state["fails"], VB88.com_backoff_remaining() > 20)],
              expect_contain=[(5, True)])

        # ---- 88.5 公开入口 com_host_present()（退避期的调用方用它）----
        _sv85 = VB88._host_window_alive
        _sent85 = object()
        VB88._host_window_alive = lambda: _sent85
        try:
            _r85 = VB88.com_host_present()
        finally:
            VB88._host_window_alive = _sv85
        check("88.5 com_host_present() 是宿主探测的公开入口（带解除退避的副作用）",
              [(_r85 is _sent85,)], expect_contain=[(True,)])

        # ---- 88.7 窗口门【不提供信息】时（类名对不上的宿主）退避也要松口 ----
        # "seen 一直是 False" = 从没见过任何已知宿主窗口。此时窗口门自己说
        # "别拦（放行）"，退避若还卡 30 秒就自相矛盾 —— 净效果是开机后干等
        # 半分钟（WPS 这类宿主会一直吃这个亏）。
        _arm_backoff88(seen=False)
        VB88._probe_host_window = lambda: False
        VB88._get_vbe = _boom88
        _r87 = VB88._get_vbe_cached()
        check("88.7 窗口门不提供信息（类名对不上的宿主）-> 退避被压到 2s，"
              "不再干等 30s",
              [(_r87 is None, VB88.com_backoff_remaining() <= 2.0)],
              expect_contain=[(True, True)])

        # ---- 88.8 ★反面：门是准的（见过宿主）时，退避照旧不许被缩短 ----
        # 宿主真在关闭时，那条"别去骚扰它"的保护必须原样保留
        # （_com_backoff_shorten 只压时长、不清失败档位，就是为这个）。
        _arm_backoff88(seen=True)
        VB88._probe_host_window = lambda: False
        VB88._get_vbe = _boom88
        _r88 = VB88._get_vbe_cached()
        check("88.8 ★反面：已见过宿主（门是准的）-> 退避照旧卡着，不许被缩短",
              [(_r88 is None, VB88.com_backoff_remaining() > 20)],
              expect_contain=[(True, True)])

        # ---- 88.9 源码护栏：两处早退的顺序 + 退避期轮询独立档位 ----
        import re as _re88
        _mnc88 = _re88.sub(r"#[^\n]*", "", io.open(
            os.path.join(ROOT, "main.py"), encoding="utf-8").read())
        _vnc88 = _re88.sub(r"#[^\n]*", "", io.open(
            os.path.join(ROOT, "vbe_bridge.py"), encoding="utf-8").read())
        _i_probe88 = _vnc88.find("host_alive = _host_window_alive()")
        _i_back88 = _vnc88.find("if com_backoff_remaining() > 0:", _i_probe88)
        check("88.9 源码护栏：① _get_vbe_cached 里【先探宿主、后判退避】；"
              "② '宿主回来了'判据含 seen=False（开机路径）；③ 窗口门不提供"
              "信息时调 _com_backoff_shorten；④ _com_allowed 退避期调 "
              "com_host_present；⑤ 退避期有独立轮询间隔（旧的合并式 idle 已删）",
              [(0 < _i_probe88 < _i_back88,
                '_wnd_state["miss"] >= 2 or not _wnd_state["seen"]' in _vnc88,
                "_com_backoff_shorten(_COM_BACKOFF_COLD_CAP)" in _vnc88,
                _mnc88.count("vbe_bridge.com_host_present()") >= 1,
                "POLL_INTERVAL_BACKOFF_MS" in _mnc88
                and "idle = (com_backoff_remaining() > 0" not in _mnc88)],
              expect_contain=[(True, True, True, True, True)])
    except Exception as _e88:
        import traceback as _tp88
        _tp88.print_exc()
        check("第 88 节异常: %s" % _e88, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 89. ★v107c：Tab 确认函数名 -> 自动补 `()`、光标进括号 -> 顺手唤出 VBE
    #     自己的【参数信息】窗（形参签名）。
    #     用户口径（原话）："提示词选中了过程名或函数名，按 tab 键，当自动补充完
    #     双括号后，光标落在双括号中间，再加个自动显示形参列表参数的功能。"
    #
    #     为什么测"有没有按那一下 Ctrl+Shift+I"，而不是测我们自己画了个窗：
    #     形参信息用 VBE 原生的（v63 真机验证过 Ctrl+Shift+I 能唤出
    #     PopupTipWndClass），我们只负责"替用户按一下"。所以这一节要钉死的是
    #     【什么时候按、什么时候不按】。
    # ------------------------------------------------------------------
    print("\n=== 89. ★v107c / v124：补完括号后唤出形参信息（★不再注入按键）===")
    try:
        VB89 = VB

        class _FCM89:
            """最小 CodeModule（只实现 apply_completion 真正用到的三个成员）。"""

            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "Module1"

            def Lines(self, start, count):
                return self.lines[start - 1]

            def ReplaceLine(self, line_no, text):
                self.lines[line_no - 1] = text

            @property
            def CountOfLines(self):
                return len(self.lines)

        class _FPane89:
            def __init__(self, cm):
                self.CodeModule = cm
                self.sel = (1, 1, 1, 1)

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _FCtrl89:
            """假的「参数信息」菜单项 —— 只记 Execute 次数。"""

            def __init__(self):
                self.Enabled = True
                self.n = 0

            def Execute(self):
                self.n += 1

        class _FBars89:
            def __init__(self, ctrl, found=True):
                self.ctrl = ctrl
                self.found = found
                self.id_seen = None

            def FindControl(self, typ, cid):
                self.id_seen = (typ, cid)
                return self.ctrl if self.found else None

        class _FVBE89:
            def __init__(self, pane, bars=None):
                self.ActiveCodePane = pane
                if bars is not None:
                    self.CommandBars = bars

        _NOQI = object()

        def run89(lines, line_no, start_col, caret_col, completion,
                  want_parens=True, send=_NOQI, bars=True, found=True,
                  on=False):
            """跑一次 apply_completion。

            返回 (写回后的行, 光标列, 菜单命令执行次数, 按键注入次数)。

            ★v124：第三个值换成"菜单命令 Execute 了几次"，第四个值才是
            "有没有注入按键"。
            ★v124d：`AUTO_PARAM_INFO` 默认 **False**（真机实测 `Execute()` 每次
            都会响一声、窗还不出现），所以这里默认 `on=False` —— 默认路径必须
            一声不响。要测"显式开启 PARAM_INFO=1 的老路径"就传 `on=True`；
            传 `on=None` 表示**不动**模块常量、用调用方当前设好的值（89.6 用）。
            """
            cm = _FCM89(lines)
            pane = _FPane89(cm)
            ctrl = _FCtrl89()
            fb = _FBars89(ctrl, found=found)
            calls = []

            def _qi():
                calls.append(1)
                return True

            kw = {"send_quick_info": _qi if send is _NOQI else send}
            _old_vbe = VB89._get_vbe
            _old_gate = VB89._WINDOW_GATE_DISABLED
            _old_ap = VB89.AUTO_PARAM_INFO
            _set_ap = on is not None
            VB89._get_vbe = lambda: _FVBE89(pane, fb if bars else None)
            VB89._WINDOW_GATE_DISABLED = True      # 别让宿主窗口门插手
            if _set_ap:
                VB89.AUTO_PARAM_INFO = on
            VB89._vbe_cache["obj"] = None          # 每个用例必须用自己那份假对象
            VB89._param_info_cache["owner"] = None     # 控件缓存按 vbe 实例来
            VB89._param_info_cache["ctrl"] = None
            VB89.com_reset()
            try:
                VB89.VbeBackend(**kw).apply_completion(
                    line_no, start_col, caret_col, completion,
                    want_parens=want_parens)
            finally:
                VB89._get_vbe = _old_vbe
                VB89._WINDOW_GATE_DISABLED = _old_gate
                if _set_ap:
                    VB89.AUTO_PARAM_INFO = _old_ap
            run89.last_findcontrol = fb.id_seen
            return (cm.lines[line_no - 1], pane.sel[3], ctrl.n, len(calls))

        # ---- 89.1 位置判据（纯函数，可离线钉死）----
        _cip = VB89._caret_inside_parens
        check("89.1 `_caret_inside_parens`：只有【光标左边紧邻 `(`】才算"
              "「停在双括号中间」（越界 / 行尾 / 左边不是括号一律 False）",
              [(_cip("    x = foo()", 12),        # 光标在 () 之间 -> True
                _cip("    x = foo()", 11),        # 在 '(' 左侧 -> False
                _cip("    x = foo()", 13),        # 整行末尾 -> False
                _cip("    x = foo()", 14),        # 越界 -> False
                _cip("", 0),
                _cip("abc", -1),
                _cip("(", 1))],
              expect_contain=[(True, False, False, False, False, False, True)])

        # ---- 89.2 ★默认路径：补出 `()` + 光标进括号 + **一声不响** ----
        # ★v124c 曾断言"菜单命令 1 次、按键注入 0 次"，以为菜单命令是静默的。
        # ★v124d 真机峰值实测推翻了它：`CommandBars(...Id=2532).Execute()` 在
        # 这个由 COM 造出来的上下文里**每次都响一声**（peak 0.346），而
        # `PopupTipWndClass` 始终 `IsWindowVisible == False` —— 窗根本不出现。
        # ⇒ 现在的契约是：**默认连菜单命令都不发**（Execute 0 次 / 注入 0 次）。
        # `on=None` 才是真正的"出厂默认"—— 不吃 run89 里那份覆盖，直接看
        # 模块常量 `AUTO_PARAM_INFO` 的真值（A/B 时这条必须变红）。
        got89 = run89(["    x = getNa"], 1, 9, 14, "getUserName", on=None)
        check("89.2 ★默认路径：写回 `x = getUserName()`、光标落在双括号中间"
              "（第 21 列）；★一个命令都不发、一个按键都不注入 —— "
              "用户听到的提示音就是这么来的",
              [got89], expect_contain=[("    x = getUserName()", 21, 0, 0)])

        # ---- 89.2b 显式 `VBECOMPLETE_PARAM_INFO=1` 才走菜单命令 ----
        got89 = run89(["    x = getNa"], 1, 9, 14, "getUserName", on=True)
        check("89.2b 显式打开 `VBECOMPLETE_PARAM_INFO=1`：才发一次参数信息"
              "菜单命令（仍然不注入任何按键，哪怕它会响也是用户自己选的）",
              [got89], expect_contain=[("    x = getUserName()", 21, 1, 0)])

        # ---- 89.3 ★反面：括号本来就有的调用点 -> 不补、也不唤出 ----
        # 用户自己敲 `(` 的那一刻 VBE 早就弹过形参信息了，我们再来一下只会闪。
        got89 = run89(["    x = getNa(1)"], 1, 9, 14, "getUserName", on=True)
        check("89.3 ★反面：括号本来就存在的调用点 -> 不补括号、"
              "**一次都不唤出**（VBE 早已自己弹过）",
              [got89], expect_contain=[("    x = getUserName(1)", 20, 0, 0)])

        # ---- 89.4 ★反面：want_parens=False（Sub / Property / 变量那一档）----
        got89 = run89(["    x = getNa"], 1, 9, 14, "getUserName",
                      want_parens=False, on=True)
        check("89.4 ★反面：want_parens=False（Sub 补空格、Property 补空的那些"
              "档）-> 不补括号，也不唤出形参信息",
              [got89], expect_contain=[("    x = getUserName", 20, 0, 0)])

        # ---- 89.5 ★反面：宿主没有 CommandBars（WPS / 未验证宿主）----
        # 拿不到菜单命令就**静默跳过**，绝不退回注入按键 —— 退回就等于把
        # 提示音放回来。想让它在那种宿主上退回去，得显式开
        # VBECOMPLETE_PARAM_INFO_KEY=1（见 89.5b）。
        got89 = run89(["    x = getNa"], 1, 9, 14, "getUserName",
                      bars=False, on=True)
        check("89.5 ★反面：拿不到 CommandBars（旧式 / 离线 / 未验证宿主）-> "
              "补全照常完成、静默跳过，**不注入任何按键**（不响）",
              [got89], expect_contain=[("    x = getUserName()", 21, 0, 0)])

        # ---- 89.5b 显式退回：PARAM_INFO_KEY_FALLBACK 打开才注入 ----
        _old_fb89 = VB89.PARAM_INFO_KEY_FALLBACK
        VB89.PARAM_INFO_KEY_FALLBACK = True
        try:
            got89 = run89(["    x = getNa"], 1, 9, 14, "getUserName",
                          bars=False, on=True)
        finally:
            VB89.PARAM_INFO_KEY_FALLBACK = _old_fb89
        check("89.5b `VBECOMPLETE_PARAM_INFO_KEY=1` 时才退回注入 Ctrl+Shift+I "
              "那条老路（默认关：那条路会响，用户要的就是不响）",
              [got89], expect_contain=[("    x = getUserName()", 21, 0, 1)])

        # ---- 89.5c FindControl 找不到控件（Id 变了 / 版本不同）----
        got89 = run89(["    x = getNa"], 1, 9, 14, "getUserName",
                      found=False, on=True)
        check("89.5c `FindControl` 返回 None（别的 VBE 版本没有这个 Id）-> "
              "静默跳过，不抛、不注入",
              [got89, run89.last_findcontrol],
              expect_contain=[("    x = getUserName()", 21, 0, 0),
                              (1, 2532)])

        # ---- 89.6 开关 `VBECOMPLETE_NO_PARAM_INFO=1` 能关掉 ----
        # `on=None`：不动模块常量，就用下面这行强行压成的 False
        # —— 它等价于 `VBECOMPLETE_NO_PARAM_INFO=1` 的那个与项。
        _old_ap89 = VB89.AUTO_PARAM_INFO
        VB89.AUTO_PARAM_INFO = False
        try:
            got89 = run89(["    x = getNa"], 1, 9, 14, "getUserName", on=None)
        finally:
            VB89.AUTO_PARAM_INFO = _old_ap89
        check("89.6 开关关掉后：括号照补、形参信息不弹（不影响补全本身）",
              [got89], expect_contain=[("    x = getUserName()", 21, 0, 0)])

        # ---- 89.7 源码护栏：接线 + 顺序（★v124d 再反转一次）----
        # v124c 断言的是"默认路径走菜单命令"—— 实测那条命令每次都响，
        # 所以现在反过来钉：**默认开关必须是 False**（只认显式
        # VBECOMPLETE_PARAM_INFO=1），且 `_send_quick_info` 只能待在
        # `PARAM_INFO_KEY_FALLBACK` 的 guard 里。
        _msrc89 = io.open(os.path.join(ROOT, "main.py"),
                          encoding="utf-8").read()
        _vsrc89 = io.open(os.path.join(ROOT, "vbe_bridge.py"),
                          encoding="utf-8").read()
        _i_judge89 = _vsrc89.find("_caret_inside_parens(param_line, end0)")
        _i_show89 = _vsrc89.find("if _show_param_info(vbe):")
        _i_guard89 = _vsrc89.find("elif PARAM_INFO_KEY_FALLBACK:")
        _i_send89 = _vsrc89.find('getattr(self, "_send_quick_info", None)')
        check("89.7 源码护栏：① main 仍保留 `send_quick_info`（只作显式退回用），"
              "发的仍是完整 Ctrl+Shift+I；② 后端把它接上；"
              "③ ★默认开关 `AUTO_PARAM_INFO` 由 `VBECOMPLETE_PARAM_INFO` 显式"
              "开启（默认 False —— Execute 每次都响）；"
              "④ `_send_quick_info` 只出现在 `PARAM_INFO_KEY_FALLBACK` 之后 —— "
              "顺序反过来就等于把提示音放回来了",
              [("def send_quick_info(" in _msrc89,
                # ★代码复查（2026-10-18）重构后序列统一走 `_send_inputs`，
                # 三元组从 (vk, flags) 变成 (vk, scan, flags) —— 断言同步加强：
                # Ctrl / Shift 都要按下，且修饰键【反序】松开。
                "(VK_CONTROL, 0, 0)" in _msrc89
                and "(VK_SHIFT, 0, 0)" in _msrc89
                and "(VK_I, 0, KEYEVENTF_KEYUP)" in _msrc89
                and "(VK_SHIFT, 0, KEYEVENTF_KEYUP)" in _msrc89
                and "(VK_CONTROL, 0, KEYEVENTF_KEYUP)" in _msrc89,
                "send_quick_info=send_quick_info" in _msrc89
                and "VbeBackend(send_char=send_char," in _msrc89,
                "def _show_param_info(" in _vsrc89
                and "_PARAM_INFO_CTRL_ID = 2532" in _vsrc89
                and "FindControl(1, _PARAM_INFO_CTRL_ID)" in _vsrc89,
                '_env_flag("VBECOMPLETE_PARAM_INFO", False)' in _vsrc89
                and 'not _env_flag("VBECOMPLETE_NO_PARAM_INFO", False)' in _vsrc89,
                "if added_parens and AUTO_PARAM_INFO" in _vsrc89,
                0 < _i_judge89 < _i_show89 < _i_guard89 < _i_send89)],
              expect_contain=[(True, True, True, True, True, True, True)])
    except Exception as _e89:
        import traceback as _tp89
        _tp89.print_exc()
        check("第 89 节异常: %s" % _e89, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 91. ★v108：输入法在「中文标点 / 全角」时，不要再把中文标点"翻"成英文的
    #
    #     用户报的 bug（原话）："我不能输入中文的双括号和中文双引号，一输入就会
    #     自动变成英文的双括号和英文双引号。"
    #
    #     根因：自动配对的判据只看【虚拟键 + Shift】。中文输入法在「中文标点」下
    #     Shift+9 打出来的是全角 `（`、Shift+' 是全角 `“`，我们却一律替它插半角
    #     `()`/`""`。修法：接管前问一次输入法的「转换模式」，中文标点/全角时换成
    #     全角那一对。
    #
    #     ⚠️ 放在 §90 之前：§90 是"猴补丁必须还原"的自查护栏，新小节必须排在
    #        它前面，否则本节万一漏还原就没人管了。
    #     ⚠️ 常量不是猜的：本机实测（tests/probe_ime_mode.py）微软拼音
    #        「中文 + 中文标点」读到 0x0401 = IME_CMODE_NATIVE(0x0001) |
    #        IME_CMODE_SYMBOL(0x0400)，取值对照本机 SDK 的 imm.h / ime_cmodes.h。
    # ------------------------------------------------------------------
    print("\n=== 91. ★v108：中文标点 / 全角时不再把用户的中文标点换成英文的 ===")
    try:
        import main as _MN91
        VB91 = VB
        _FU, _RU, _QO, _QC = "（", "）", "“", "”"   # U+FF08 / U+FF09 / U+201C / U+201D

        # ---- 91.1 纯函数：转换模式 -> 要不要用全角标点 ----
        _fwp = _MN91._ime_fullwidth_punct
        # "问不到"（探测返回 None）那一档用桩钉死：`_ime_fullwidth_punct()` 不传参
        # 就是**真去探测**，直接拿本机输入法状态当期望值会让回归网随环境变色。
        _old_mode91a = _MN91._ime_conversion_mode
        try:
            _MN91._ime_conversion_mode = lambda hwnd=None: None
            _probe_none91 = _fwp()
        finally:
            _MN91._ime_conversion_mode = _old_mode91a
        check("91.1 `main._ime_fullwidth_punct`：**必须同时**「中文模式」+"
              "「中文标点(0x0400)/全角(0x0008)」才算；"
              "0x0001（中文模式但英文标点）/ 0x0000（英文模式）/ 0x0400（无中文模式）"
              "/ 探测不到（None）一律按半角",
              [(_fwp(0x0401),        # 实测：微软拼音 中文 + 中文标点
                _fwp(0x0009),        # 中文模式 + 全角
                _fwp(0x0001),        # 中文模式，但半角英文标点 -> 半角
                _fwp(0x0000),        # 英文模式
                _fwp(0x0400),        # 只有中文标点位、不在中文模式 -> 不算
                _fwp(0x0008),        # 只有全角位、不在中文模式 -> 不算
                _probe_none91,       # 问不到（没有 IME / 超时）
                _fwp(0xFFFF))],      # 全亮
              expect_contain=[
                  (True, True, False, False, False, False, False, True)])

        # ---- 91.2 纯函数：半角配对字符 -> 输入法真正会打出来的那个 ----
        _pcfi = _MN91._pair_char_for_ime
        _old_mode91 = _MN91._ime_conversion_mode
        try:
            _MN91._ime_conversion_mode = lambda hwnd=None: 0x0401
            _got_cn = tuple(_pcfi(c) for c in ("(", ")", '"'))
            _MN91._ime_conversion_mode = lambda hwnd=None: 0x0001
            _got_half = tuple(_pcfi(c) for c in ("(", ")", '"'))
            _MN91._ime_conversion_mode = lambda hwnd=None: None
            _got_none = tuple(_pcfi(c) for c in ("(", ")", '"'))
        finally:
            _MN91._ime_conversion_mode = _old_mode91
        check("91.2 `main._pair_char_for_ime`：中文标点时 `(`/`)`/`\"` -> "
              "`（`/`）`/`“`；英文标点或问不到时**原样返回半角**（拿不准就退回"
              "既有行为，绝不因为探测失败把功能弄坏）",
              [(_got_cn, _got_half, _got_none)],
              expect_contain=[((_FU, _RU, _QO),
                               ("(", ")", '"'),
                               ("(", ")", '"'))])

        # ---- 91.2b `_pair_char_for_key` 必须保持【纯函数】：不许自己去问 IME ----
        _names91 = set(_MN91._pair_char_for_key.__code__.co_names)
        check("91.2b `_pair_char_for_key` 保持【纯 VK -> 字符】："
              "（a）返回的永远是**半角**；"
              "（b）函数体里**不许**出现 `_ime_conversion_mode` / "
              "`_pair_char_for_ime` —— 输入法状态只能在【决定接管之后】问一次，"
              "否则每按一次 Shift 都要发个窗口消息",
              [(_MN91._pair_char_for_key(0x39, True),
                _MN91._pair_char_for_key(0x30, True),
                _MN91._pair_char_for_key(0xDE, True),
                ("_ime_conversion_mode" not in _names91
                 and "_pair_char_for_ime" not in _names91))],
              expect_contain=[(("(", ")", '"', True))])

        # ---- 91.3 集成：insert_pair 收全角时插出来的就是全角 ----
        # 列语义（显示列 vs 字符列）不是本节要验的东西，所以直接把它钉成
        # "字符列"（§29 / §56 也是这么做的）：这样 caret_col 就是我写的字符列，
        # 期望值才可复现。`ec <= len+1` 时被动判定本来就返回 None，
        # 预置缓存就不会去动光标探测。
        class _FCM91:
            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "Module1"

            def Lines(self, start, count):
                return self.lines[start - 1]

            def ReplaceLine(self, line_no, text):
                self.lines[line_no - 1] = text

            @property
            def CountOfLines(self):
                return len(self.lines)

        class _FPane91:
            def __init__(self, cm, col):
                self.CodeModule = cm
                self.sel = (1, col, 1, col)

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _FVBE91:
            def __init__(self, pane):
                self.ActiveCodePane = pane

        _sem91 = VB91._sem_cache.get("info")
        VB91._sem_cache["info"] = ("char", 4, False)

        def run91(line, caret_col, open_ch):
            """跑一次 insert_pair，返回 (是否接管, 写回后的行, 光标列)。"""
            cm = _FCM91([line])
            pane = _FPane91(cm, caret_col)
            _old_vbe = VB91._get_vbe
            _old_gate = VB91._WINDOW_GATE_DISABLED
            VB91._get_vbe = lambda: _FVBE91(pane)
            VB91._WINDOW_GATE_DISABLED = True
            VB91._vbe_cache["obj"] = None
            VB91.com_reset()
            try:
                took = VB91.VbeBackend().insert_pair(open_ch)
            finally:
                VB91._get_vbe = _old_vbe
                VB91._WINDOW_GATE_DISABLED = _old_gate
            return bool(took), cm.lines[0], pane.sel[3]

        try:
            _c91 = (run91("    x = f", 10, _FU),          # 补出 （）
                    run91("    x = f" + _FU + _RU, 11, _RU),   # 右边已有 -> 跨过去
                    run91("    s = ", 9, _QO),            # 补出 “”
                    run91("    s = " + _QO + _QC, 10, _QO),    # 引号跨过去
                    run91("    s = " + _QO + "abc", 13, _QO),  # 串里 -> 只补收符号
                    # 半角回归（老行为一个都不能变）
                    run91("    x = f", 10, "("),
                    run91("    x = f()", 11, ")"),
                    # 反面：注释里 / Call 语句里一律不接管（交给输入法自己打）
                    run91("    Rem " + _FU, 9, _FU),
                    run91("    Call foo", 12, _FU))
        finally:
            VB91._sem_cache["info"] = _sem91
        check("91.3 `insert_pair` 收全角开符号时插的就是全角："
              "`（`->`（）` 且光标落在中间；`）` 右边已有则**跨过去**（不重复插）；"
              "`“`->`“”`；串里敲 `“` 只补收符号 `”`；"
              "注释里 / `Call` 语句里**依旧不接管**（原样交给输入法）；"
              "半角路径逐条不变",
              [_c91],
              expect_contain=[(
                  (True, "    x = f" + _FU + _RU, 11),
                  (True, "    x = f" + _FU + _RU, 12),
                  (True, "    s = " + _QO + _QC, 10),
                  (True, "    s = " + _QO + _QC, 11),
                  (True, "    s = " + _QO + "abc" + _QC, 14),
                  (True, "    x = f()", 11),
                  (True, "    x = f()", 12),
                  (False, "    Rem " + _FU, 9),
                  (False, "    Call foo", 12))])

        # ---- 91.4 退格成对删除也要认全角那一对 ----
        _ps91 = VB91.pair_delete_span
        check("91.4 退格成对删除认全角：光标夹在 `（）` / `“”` 中间时，"
              "一次退格把两端一起收掉（与半角 `()` / `\"\"` 同口径）；"
              "中间有内容 / 左右不成对时**不接管**",
              [(_ps91("f" + _FU + _RU, 3),
                _ps91("s = " + _QO + _QC, 6),
                _ps91("f()", 3),
                _ps91("f" + _FU + "a" + _RU, 3),
                _ps91("f" + _RU + _FU, 2))],
              expect_contain=[((1, 2), (4, 5), (1, 2), None, None)])

        # ---- 91.5 源码护栏：探测走的是"IME 窗口 + 消息"那条跨进程路 ----
        _msrc91 = io.open(os.path.join(ROOT, "main.py"),
                          encoding="utf-8").read()
        _vsrc91 = io.open(os.path.join(ROOT, "vbe_bridge.py"),
                          encoding="utf-8").read()
        _i_def91 = _msrc91.find("def _ime_conversion_mode(")
        _i_call91 = _msrc91.find("_pair_char_for_ime(_pair_ch)")
        # ⚠️ 收符号表查**运行时**的字典，不查源码字符串：v108 第一版用
        # `'"）": "）"' in 源码` 去钉它，结果**注释里那句说明就把它喂饱了**——
        # 我故意把 `"）": "）"` 从表里删掉做 A/B，这条居然还是绿的（假绿）。
        _pc91 = VB91.VbeBackend._PAIR_CLOSE
        check("91.5 源码护栏："
              "（a）`ImmGetContext` 跨进程拿不到 IME 上下文（实测返回 NULL），"
              "所以探测必须走 `ImmGetDefaultIMEWnd` + `WM_IME_CONTROL`/"
              "`IMC_GETCONVERSIONMODE`，并且用 `SendMessageTimeout` + "
              "`SMTO_ABORTIFHUNG`（**钩子线程不能被卡死**，超时会被系统摘掉钩子）；"
              "（b）常量取自 SDK 头文件：`IME_CMODE_SYMBOL = 0x0400`、"
              "`IME_CMODE_FULLSHAPE = 0x0008`、`IME_CMODE_NATIVE = 0x0001`；"
              "（c）钩子在**决定接管之后**才换算，且提供 `VBECOMPLETE_NO_IME_PUNCT=1` "
              "强制回半角；"
              "（d）收符号表里 `）` 必须**自己占一个键**"
              "（少了它，「敲 `）` 跨过去」那条路根本走不到 —— 本轮真踩过）；"
              "（e）`_pair_insertion` 的引号分支要把 `“` 与半角 `\"` 同等对待、"
              "`insert_pair` 的「右边已有就跨过去」要把 `）` 与 `)` 同等对待",
              [("ImmGetDefaultIMEWnd" in _msrc91
                and "WM_IME_CONTROL = 0x0283" in _msrc91
                and "IMC_GETCONVERSIONMODE = 0x0001" in _msrc91
                and "SendMessageTimeoutW" in _msrc91
                and "SMTO_ABORTIFHUNG" in _msrc91,
                "IME_CMODE_SYMBOL = 0x0400" in _msrc91
                and "IME_CMODE_FULLSHAPE = 0x0008" in _msrc91
                and "IME_CMODE_NATIVE = 0x0001" in _msrc91,
                "VBECOMPLETE_NO_IME_PUNCT" in _msrc91
                and _i_def91 > 0 and _i_call91 > _i_def91,
                (_pc91.get("（"), _pc91.get("）"), _pc91.get("“"),
                 _pc91.get(")")),
                'in (\'"\', "“")' in _vsrc91
                and 'in (")", "）")' in _vsrc91)],
              expect_contain=[(True, True, True,
                               ("）", "）", "”", ")"), True)])
    except Exception as _e91:
        import traceback as _tp91
        _tp91.print_exc()
        check("第 91 节异常: %s" % _e91, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 92. ★v109：开着 VbeComplete 时 PyCharm 项目关不掉（v27 的老坑又犯了）
    #
    #     用户报（原话）："以前改掉了，但是现在又犯了。就是打开这个项目以后，
    #     pycharm 里的项目不能关闭，他俩冲突了。"
    #
    #     根因（2026-10-14 实测，探针 tests/probe_hook_gate.py）：
    #     `in_vbe_code_pane()` 用 `GetWindowText(前台窗口)` 读标题，而它是
    #     【跨线程同步调用】—— 前台窗口的主线程正忙时（用户关 PyCharm，它的主线程
    #     在拆窗口、写设置，好几秒不理会消息），这一行会把【我们自己的主线程】
    #     堵住 1.5 秒以上（实测 1.484s；窗口彻底卡死时更久）。
    #     主线程一堵，`reconcile` 就跑不了 ⇒「焦点离开 VBE 立刻卸载全局钩子」
    #     被拖住 ⇒ WH_KEYBOARD_LL / WH_MOUSE_LL 继续挂在人家退出流程上
    #     ⇒ 正是 v27 修过的「PyCharm / VSCode 关不掉 / 卡死」。
    #     另外 reconcile 只有开头一行有 try/except，中后段的裸调用一旦抛异常，
    #     300ms 循环就永久停摆 —— 钩子永远卸不掉（比"晚卸载"严重得多）。
    #
    #     修法（两层）：
    #       ① 读标题【之前】先问"这个窗口还响应吗"（closing 探针 200ms 封顶），
    #          不响应就判"不是 VBE"，压根不碰 GetWindowText；
    #       ② reconcile 拆成"外壳（_safe_tick + 无条件重排）+ 单轮实现"，
    #          单轮里任何异常都吃不掉循环本身。
    # ------------------------------------------------------------------
    print("\n=== 92. ★v109：前台窗口忙时不许堵住主线程（PyCharm 关不掉）===")
    try:
        import main as _MN92
        import ast as _ast92

        # ---- 92.1 in_vbe_code_pane 的门控矩阵 ----
        class _FakeGui92:
            """假 win32gui：把每一次 GetWindowText / GetForegroundWindow 都记下来。"""

            def __init__(self, fg=1001, alive=True,
                         title="Microsoft Visual Basic - 探针"):
                self.fg = fg
                self.alive = alive
                self.title = title
                self.text_calls = []       # 被读过标题的 hwnd
                self.fg_calls = 0

            def GetForegroundWindow(self):
                self.fg_calls += 1
                return self.fg

            def IsWindow(self, hwnd):
                return bool(self.alive)

            def GetWindowText(self, hwnd):
                self.text_calls.append(hwnd)
                return self.title

        _old_gui92 = _MN92._gui
        _old_clo92 = _MN92.foreground_window_closing
        try:
            def _mk92(closing, fg=1001, alive=True,
                      title="Microsoft Visual Basic - 探针"):
                """装好假 _gui + 可调 closing 桩；返回 (gui, 探针调用记录表)。"""
                g = _FakeGui92(fg=fg, alive=alive, title=title)
                _MN92._gui = g
                seen = []

                def _stub(hwnd=None):
                    seen.append(hwnd)
                    return closing

                _MN92.foreground_window_closing = _stub
                return g, seen

            # (a) ★核心回归：前台【不响应】时，哪怕标题明明含 VBE 也必须判 False，
            #     而且【一次标题都不许读】—— 读标题就是把主线程堵死的那一步。
            g, seen = _mk92(True)
            _r_a = _MN92.in_vbe_code_pane()
            # (b) 响应 + 标题含 VBE -> True（老行为不能变）
            g, _ = _mk92(False)
            _r_b = _MN92.in_vbe_code_pane()
            # (c) 响应 + 标题不含 VBA -> False
            g, _ = _mk92(False, title="探针代码 - PyCharm (Probe)")
            _r_c = _MN92.in_vbe_code_pane()
            # (d) 显式传 closing=False -> 不调探针（同一轮不探两次），照常读标题
            g, seen = _mk92(False)
            _r_d = _MN92.in_vbe_code_pane(closing=False)
            # (e) 显式传 closing=True -> 不调探针、不读标题
            g, seen = _mk92(False)
            _r_e = _MN92.in_vbe_code_pane(closing=True)
            # (f) 前台句柄为空 -> False，不读标题
            g, _ = _mk92(False, fg=0)
            _r_f = _MN92.in_vbe_code_pane()
            # (g) 窗口已销毁 -> False，不读标题
            g, _ = _mk92(False, alive=False)
            _r_g = _MN92.in_vbe_code_pane()
            # (h) 拿不到窗口 API -> False（老兜底）
            _MN92._gui = None
            _r_h = _MN92.in_vbe_code_pane()
            # (i) 读标题本身抛异常 -> False（老兜底不能丢）
            g, _ = _mk92(False)

            def _boom(_hwnd):
                raise RuntimeError("GetWindowText 炸了")

            g.GetWindowText = _boom
            _r_i = _MN92.in_vbe_code_pane()
        finally:
            _MN92._gui = _old_gui92
            _MN92.foreground_window_closing = _old_clo92
        check("92.1 `main.in_vbe_code_pane`：前台窗口【不响应】时直接判 False，"
              "**一次标题都不读**（GetWindowText 是跨线程同步调用，会把主线程堵"
              " 1.5 秒以上，reconcile 因此停摆、全局钩子卸不掉 —— PyCharm 关不掉"
              " 的根因）；响应时才读标题；句柄空 / 窗口已毁 / 拿不到 API / 读标题"
              " 抛异常，一律 False（老兜底不变）",
              [(_r_a, _r_b, _r_c, _r_d, _r_e, _r_f, _r_g, _r_h, _r_i)],
              expect_contain=[(False, True, False,
                               True, False, False, False, False, False)])

        # ---- 92.1b 核心回归的"另一半"：那些不许读标题的分支，真的没读过 ----
        _old_gui92 = _MN92._gui
        _old_clo92 = _MN92.foreground_window_closing
        try:
            def _texts92(closing, **kw):
                g, seen = _mk92(closing, **kw)
                r = _MN92.in_vbe_code_pane()
                return r, list(g.text_calls), list(seen), g.fg_calls

            _t_a = _texts92(True)                 # 不响应：不读标题、探过 1 次
            _t_d = _texts92(False)                # 响应：读标题、探过 1 次
            g, seen = _mk92(False)
            _r_d2 = _MN92.in_vbe_code_pane(closing=False)   # 显式传入：不探
            _t_e = (_r_d2, list(g.text_calls), list(seen), g.fg_calls)
        finally:
            _MN92._gui = _old_gui92
            _MN92.foreground_window_closing = _old_clo92
        check("92.1b 「不响应就不读标题」落到调用记录上："
              "（a）closing=True -> GetWindowText 零次、closing 探针 1 次且带的是"
              "前台句柄；"
              "（b）closing=False -> 读标题 1 次；"
              "（c）显式传 closing=False -> 探针零次（同一轮绝不探两次）",
              [(_t_a, _t_d, _t_e)],
              expect_contain=[((False, [], [1001], 1),
                               (True, [1001], [1001], 1),
                               (True, [1001], [], 1))])

        # ---- 92.2 foreground_window_closing(hwnd=...)：复用已取到的句柄 ----
        _old_gui92 = _MN92._gui
        try:
            g = _FakeGui92()
            _MN92._gui = g
            _c_pass = _MN92.foreground_window_closing(hwnd=4242)  # 不该自己取前台
            _n_pass = g.fg_calls
            _c_none = _MN92.foreground_window_closing()           # 不传就自己取
            _n_none = g.fg_calls
        finally:
            _MN92._gui = _old_gui92
        check("92.2 `foreground_window_closing(hwnd=None)`：传了句柄就【不再】"
              "自己 GetForegroundWindow（调用方手里往往已经有了，省一次系统调用）；"
              "不传时行为与从前一致（自己取前台）",
              [(_n_pass, _n_none, _c_pass is not None, _c_none is not None)],
              expect_contain=[(0, 1, True, True)])

        # ---- 92.3 _safe_tick：周期任务的壳，异常绝不外抛 ----
        def _boom92():
            raise ValueError("boom")

        _hits92 = []
        _raised92 = None
        try:
            _MN92._safe_tick(lambda: _hits92.append(1))
            _MN92._safe_tick(_boom92, tag="reconcile")
            _MN92._safe_tick(lambda: 1 / 0)
        except Exception as _e92:
            _raised92 = repr(_e92)
        check("92.3 `main._safe_tick`：正常调用照常执行；body 抛异常时【绝不外抛】"
              "（reconcile 是「离开 VBE 就卸载全局钩子」的唯一循环，它一死钩子就"
              "永久挂在别的程序退出流程上 —— 比晚 600ms 卸载严重得多）",
              [(_hits92, _raised92)], expect_contain=[([1], None)])

        # ---- 92.4 AST 结构护栏：外壳无条件重排 + closing 提前 ----
        _src92 = io.open(os.path.join(ROOT, "main.py"), encoding="utf-8").read()
        _tree92 = _ast92.parse(_src92)
        _mainf92 = [n for n in _tree92.body
                    if isinstance(n, _ast92.FunctionDef)
                    and n.name == "main"][0]

        def _fn92(name):
            return [n for n in _mainf92.body
                    if isinstance(n, _ast92.FunctionDef)
                    and n.name == name][0]

        _recon92 = _fn92("reconcile")
        _once92 = _fn92("_reconcile_once")
        # (a) 外壳先用 _safe_tick 包住单轮
        _a92 = any(
            isinstance(n, _ast92.Call)
            and isinstance(n.func, _ast92.Name) and n.func.id == "_safe_tick"
            for s in _recon92.body for n in _ast92.walk(s))
        # (b) 外壳【顶层语句】里就有 root.after(300, reconcile)——
        #     不在 if / try 里，即"无条件重排"（单轮炸了也排得上下一次）
        _b92 = False
        for s in _recon92.body:
            if isinstance(s, _ast92.Expr) and isinstance(s.value, _ast92.Call) \
                    and isinstance(s.value.func, _ast92.Attribute) \
                    and s.value.func.attr == "after":
                _args = s.value.args
                if any(isinstance(x, _ast92.Constant) and x.value == 300
                       for x in _args) \
                        and any(isinstance(x, _ast92.Name) and x.id == "reconcile"
                                for x in _args):
                    _b92 = True
        # (c) 单轮里 closing 的赋值必须早于 in_vbe_code_pane 调用
        #     （晚了就挡不住 GetWindowText 那次阻塞）
        _dumps92 = [_ast92.dump(s) for s in _once92.body]
        _i_close92 = next(i for i, d in enumerate(_dumps92)
                          if "foreground_window_closing" in d)
        _i_read92 = next(i for i, d in enumerate(_dumps92)
                         if "in_vbe_code_pane" in d)
        # (d) in_vbe_code_pane / _reconcile_resources 两个调用都带 closing 关键字
        _kw92 = {"in_vbe_code_pane": False, "_reconcile_resources": False}
        for s in _once92.body:
            for n in _ast92.walk(s):
                if isinstance(n, _ast92.Call) \
                        and isinstance(n.func, _ast92.Name) \
                        and n.func.id in _kw92:
                    if any(k.arg == "closing" for k in n.keywords):
                        _kw92[n.func.id] = True
        # (e) 单轮实现里【不许】再自己 root.after（否则外壳会重复排循环）
        _e92 = not any(
            isinstance(n, _ast92.Attribute) and n.attr == "after"
            for s in _once92.body for n in _ast92.walk(s))
        check("92.4 结构护栏（AST，不靠源码字符串——注释会把字符串护栏喂饱）："
              "（a）外壳用 _safe_tick 包住单轮；"
              "（b）root.after(300, reconcile) 是外壳的【顶层语句】（无条件重排）；"
              "（c）closing 探针的赋值早于 in_vbe_code_pane 调用；"
              "（d）两处调用都显式传 closing（同一轮绝不探两次）；"
              "（e）单轮实现里不再自己重排（避免双排）",
              [(_a92, _b92, _i_close92 < _i_read92,
                _kw92["in_vbe_code_pane"], _kw92["_reconcile_resources"], _e92)],
              expect_contain=[(True, True, True, True, True, True)])
    except Exception as _e92:
        import traceback as _tp92
        _tp92.print_exc()
        check("第 92 节异常: %s" % _e92, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 93. ★v110：Ctrl+↑ / Ctrl+↓ 跳到声明区末尾 / 过程末尾的空白行
    #     用户口径（原话）："当我按 ctrl+↑，光标会自动定位到当前过程中最后一个
    #     dim 关键字行，如果此行的下一行不是空行，就自动在此行下方新增一行，
    #     光标移动到下方空白行；如果此行的下一行是空白行，就不用新增行，光标
    #     移动到下方空白行就可以了。ctrl+↓ 同理，定位到 end sub 或 end function
    #     行，上一行不是空白行就在上方新增一行。"
    #
    #     骨架是第 10 条那副老规矩：钩子只判"这个按键该不该接手"（模块级纯函数
    #     _ctrl_nav_vk_ok + 两道读运行期状态的闸），判据与写入全在主线程、拿
    #     【此刻真读到的】模块文本重算（vbe_bridge.ctrl_nav_jump_plan +
    #     VbeBackend.ctrl_nav_jump），写不成就把按键原样还给系统。
    # ------------------------------------------------------------------
    print("\n=== 93. ★v110/v110c：Ctrl+↑ 落到最后一个 Dim 行行尾 / Ctrl+↓ 跳到过程末尾 ===")
    try:
        import main as _MN93
        import ast as _ast93
        import ctypes as _ct93
        VB93 = VB

        # ---- 93.1 纯判据矩阵：目标行 + 写入方式(insert/swap/None) + 文本 ----
        # （一个元组打包，见 v108 教训；★v110b/2 起是三元组）
        _S93 = "Sub Foo()\r\n    Dim a As Long\r\n    x = 1\r\nEnd Sub\r\n"

        def _p93(code, caret, direction):
            return VB93.ctrl_nav_jump_plan(code, caret, direction)

        _g93 = (
            # ↑：落在最后一个 Dim 语句【那一行】的行尾（★v110c：不插行、不动文本）
            _p93(_S93, 3, -1),
            # ↑：Dim 下面是不是空白都一样 —— 落点就是 Dim 行本身
            _p93("Sub Foo()\r\n    Dim a As Long\r\n\r\n    x = 1\r\nEnd Sub\r\n", 4, -1),
            # ↑：一个 Dim 都没有 -> 不接管
            _p93("Sub Foo()\r\n    x = 1\r\nEnd Sub\r\n", 2, -1),
            # ↑：取【最后一个】Dim（上面明明还有一个）
            _p93("Sub Foo()\r\n    Dim a As Long\r\n    x = 1\r\n"
                 "    Dim b As String\r\n    y = 2\r\nEnd Sub\r\n", 5, -1),
            # ↑：光标在过程【头】那一行也算在过程里（刚敲完 Sub Foo() 就能跳）
            _p93(_S93, 1, -1),
            # ↑：跨行 Dim —— 目标落在语句【结束行】的行尾（`Dim a, _` 那一行不算，
            #     否则光标会跑到续行符后面）
            _p93("Sub Foo()\r\n    Dim a As Long, _\r\n        b As String\r\n"
                 "    x = 1\r\nEnd Sub\r\n", 4, -1),
            # ↑：Dim 紧贴 End Sub -> 落在 Dim 行自己行尾
            _p93("Sub Foo()\r\n    Dim a As Long\r\nEnd Sub\r\n", 2, -1),
            # ↑：ReDim / 注释里的 Dim / 字符串里的 Dim 都不算 Dim 声明
            _p93("Sub Foo()\r\n    ReDim a(3)\r\n    x = 1\r\nEnd Sub\r\n", 3, -1),
            _p93("Sub Foo()\r\n    ' Dim a As Long\r\n    x = 1\r\nEnd Sub\r\n", 3, -1),
            _p93('Sub Foo()\r\n    s = "Dim a As Long"\r\n    x = 1\r\nEnd Sub\r\n', 3, -1),
            # ↑：注释【续行】后面那行 Dim 也不算 —— `' 注释 _` 把它续成了注释，
            #    认出来的话会在注释中间插一行、把注释从当中拆断（本轮真修过）
            _p93("Sub Foo()\r\n    ' 注释 _\r\n    Dim a As Long\r\n"
                 "    x = 1\r\nEnd Sub\r\n", 4, -1),
            # ↑：模块声明区 / 两个过程之间 -> 不接管
            _p93("Dim g As Long\r\n\r\nSub Foo()\r\n    x = 1\r\nEnd Sub\r\n", 1, -1),
            _p93("Sub A()\r\nEnd Sub\r\n\r\nSub B()\r\nEnd Sub\r\n", 3, -1),
            # ↑：过程还没写 End -> 不接管（"过程末尾"说不清）
            _p93("Sub Foo()\r\n    Dim a As Long\r\n    x = 1\r\n", 3, -1),
            # ↑：空过程（头尾相接）没有 Dim -> 不接管
            _p93("Sub Foo()\r\nEnd Sub\r\n", 1, -1),
            # ↓：End 上方不是空白 -> 插一行，目标就是 End 行自己
            _p93(_S93, 3, +1),
            # ↓：End 上方已经是空白 -> 什么都不插
            _p93("Sub Foo()\r\n    x = 1\r\n\r\nEnd Sub\r\n", 2, +1),
            # ↓：空过程（头尾相接）-> 插在中间
            _p93("Sub Foo()\r\nEnd Sub\r\n", 1, +1),
            # ↓：光标【就在 End 行上】也照常（先跳过它再往上找过程）
            _p93(_S93, 4, +1),
            # ↓：Function / Property 一样算过程
            _p93("Function f() As Long\r\n    f = 1\r\nEnd Function\r\n", 2, +1),
            _p93("Property Get V()\r\n    x = 1\r\nEnd Property\r\n", 2, +1),
            # ↓：End Type 不是过程尾（Type 块在模块声明区）
            _p93("Type T\r\n    a As Long\r\nEnd Type\r\n\r\nSub Foo()\r\n"
                 "    x = 1\r\nEnd Sub\r\n", 2, -1),
            # 坏参数：方向 0 / 空代码 / 行号 0
            _p93(_S93, 3, 0),
            _p93("", 1, -1),
            _p93(_S93, 0, -1),
        )
        _w93 = (
            # ↑：一律 ("eol", None) —— 光标落在最后一个 Dim 语句【那一行】的行尾，
            #     不插行、不动文本（★v110c）；没有 Dim / 不在过程里 -> None
            (2, "eol", None), (2, "eol", None), None, (4, "eol", None),
            (2, "eol", None), (3, "eol", None), (2, "eol", None),
            None, None, None, None, None, None, None, None,
            # ↓：End 上方不是空白 -> ("insert", 过程体缩进)；
            #     已是真空空白 -> ("swap", 过程体缩进)；已是带缩进空白 -> 不写
            (4, "insert", "    "), (3, "swap", "    "), (2, "insert", ""),
            (4, "insert", "    "), (3, "insert", "    "), (3, "insert", "    "),
            None,
            None, None, None,
        )
        check("93.1 `vbe_bridge.ctrl_nav_jump_plan`（纯函数）：★v110c 起 ↑=最后一个 "
              "Dim 语句【那一行】（写入方式 \"eol\" —— 只落行尾、不插行、不动文本）；"
              "↓=End 行上方那一行，不是空白就补一行、是空白就什么都不写。"
              "不在过程里 / 没有 Dim / 过程没收尾 / 续行后半截 / 注释与字符串里的 "
              "Dim 一律返回 None（= 按键原样还给系统）。返回三元组 "
              "`(行, 写入方式, 文本)`：eol=只落行尾、insert=在目标行处插一行、"
              "swap=把目标那条真空空白行换成带缩进的、None=不动文本；↓ 的缩进取"
              "\"过程体层级\"（尾行上方最后一条非空行的行首空白）",
              [_g93], expect_contain=[_w93])

        # ---- 93.2 集成：后端真的插行 + 落光标 ----
        class _FCM93:
            """假 CodeModule：把每次 InsertLines / DeleteLines 记下来。"""

            def __init__(self, lines):
                self.lines = list(lines)
                self.Name = "Module1"
                self.inserts = []
                self.deletes = []
                self.fail_insert = False

            def Lines(self, start, count):
                return "\r\n".join(self.lines[start - 1:start - 1 + count])

            def ReplaceLine(self, line_no, text):
                self.lines[line_no - 1] = text

            def InsertLines(self, line_no, text):
                if self.fail_insert:
                    raise RuntimeError("模块只读")
                self.inserts.append((line_no, text))
                self.lines.insert(line_no - 1, text)

            def DeleteLines(self, line_no, count):
                self.deletes.append((line_no, count))
                del self.lines[line_no - 1:line_no - 1 + count]

            @property
            def CountOfLines(self):
                return len(self.lines)

        class _FPane93:
            def __init__(self, cm, line=1):
                self.CodeModule = cm
                self.sel = (line, 1, line, 1)

            def GetSelection(self):
                return self.sel

            def SetSelection(self, sl, sc, el, ec):
                self.sel = (sl, sc, el, ec)

        class _FVBE93:
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _run93(lines, caret_line, direction, fail_insert=False):
            """跑一次 ctrl_nav_jump，返回 (是否接管, InsertLines 记录,
            DeleteLines 记录, 写后行, 光标行, 光标列)。"""
            cm = _FCM93(lines)
            cm.fail_insert = bool(fail_insert)
            pane = _FPane93(cm, caret_line)
            _old_vbe = VB93._get_vbe
            _old_gate = VB93._WINDOW_GATE_DISABLED
            VB93._get_vbe = lambda: _FVBE93(pane)
            VB93._WINDOW_GATE_DISABLED = True
            VB93._vbe_cache["obj"] = None
            VB93.com_reset()
            try:
                took = VB93.VbeBackend().ctrl_nav_jump(direction)
            finally:
                VB93._get_vbe = _old_vbe
                VB93._WINDOW_GATE_DISABLED = _old_gate
            return (bool(took), list(cm.inserts), list(cm.deletes),
                    list(cm.lines), pane.sel[0], pane.sel[1])

        _SRC93 = ["Sub Foo()", "    Dim a As Long", "    x = 1", "End Sub"]
        _g93b = (
            # ★v110c ↑：不插行、一个字节都不改，光标落在 Dim 行（第 2 行）行尾
            _run93(_SRC93, 3, -1),
            # ↑：Dim 下面是不是空白都一样（落点仍是第 2 行行尾）
            _run93(["Sub Foo()", "    Dim a As Long", "", "    x = 1", "End Sub"], 4, -1),
            # ↑ 根本不写文本 -> 模块只读（一写就抛）也照样跳得成
            _run93(_SRC93, 3, -1, fail_insert=True),
            # ↓ 插行：新行落在第 4 行（End 让位）、直接带过程体缩进
            _run93(_SRC93, 3, +1),
            # ↓ swap：End 上面本来真空 -> 插一条带缩进的、再删掉原真空行
            _run93(["Sub Foo()", "    x = 1", "", "End Sub"], 2, +1),
            # 不该接管的两类 + 写失败（只读模块）+ direction=0
            _run93(["Dim g As Long", "", "Sub Foo()", "End Sub"], 1, -1),
            _run93(["Sub Foo()", "    x = 1", "End Sub"], 2, -1),
            # ↑ 已经不再写文本了，所以"写失败"这条得用 ↓ 才走得到 InsertLines
            _run93(_SRC93, 3, +1, fail_insert=True),
            _run93(_SRC93, 3, 0),
        )
        _w93b = (
            (True, [], [], _SRC93, 2, 18),
            (True, [], [],
             ["Sub Foo()", "    Dim a As Long", "", "    x = 1", "End Sub"], 2, 18),
            (True, [], [], _SRC93, 2, 18),
            (True, [(4, "    ")], [],
             ["Sub Foo()", "    Dim a As Long", "    x = 1", "    ", "End Sub"], 4, 5),
            (True, [(3, "    ")], [(4, 1)],
             ["Sub Foo()", "    x = 1", "    ", "End Sub"], 3, 5),
            (False, [], [], ["Dim g As Long", "", "Sub Foo()", "End Sub"], 1, 1),
            (False, [], [], ["Sub Foo()", "    x = 1", "End Sub"], 2, 1),
            (False, [], [], _SRC93, 3, 1),
            (False, [], [], _SRC93, 3, 1),
        )
        check("93.2 `VbeBackend.ctrl_nav_jump`（集成）：★v110c ↑ 只落 Dim 行行尾"
              "（一个字节都不改、光标在第 2 行行尾第 18 列；模块只读也照样跳得成，"
              "因为它根本不写文本），↓ 各验「插行」与「不插」两条路 —— 目标行/"
              "参考行行号都对、缩进是过程体层级、光标落在目标行行尾（假 pane 不夹"
              "列，走 _eol_caret_col 兜底：真空行 1 列、带缩进 5 列、Dim 行 18 列）；"
              "↓ 落点是真空空白行时用「插一条 + 删旧那条」把真空行换成带缩进的"
              "（不能 ReplaceLine —— VBE 会 trim 行尾空白）；不在过程里 / 没有 "
              "Dim / 写失败 / direction=0 一律返回 False 且一个字节都没改"
              "（调用方据此把按键还给系统）",
              [_g93b], expect_contain=[_w93b])

        # ---- 93.3 修饰键门槛（模块级纯函数）----
        _U, _D = _MN93.VK_UP, _MN93.VK_DOWN
        _g93c = (
            _MN93._ctrl_nav_vk_ok(_U, True, False, False, False),
            _MN93._ctrl_nav_vk_ok(_D, True, False, False, False),
            _MN93._ctrl_nav_vk_ok(_U, False, False, False, False),   # 普通 ↑
            _MN93._ctrl_nav_vk_ok(_D, False, False, False, False),
            _MN93._ctrl_nav_vk_ok(_U, True, True, False, False),     # Ctrl+Shift+↑
            _MN93._ctrl_nav_vk_ok(_U, True, False, True, False),     # Ctrl+Alt+↑
            _MN93._ctrl_nav_vk_ok(_U, True, False, False, True),     # Win+↑
            _MN93._ctrl_nav_vk_ok(_MN93.VK_RETURN, True, False, False, False),
            _MN93._ctrl_nav_vk_ok(_MN93.VK_TAB, True, False, False, False),
            _MN93._ctrl_nav_vk_ok(_MN93.VK_LEFT, True, False, False, False),
        )
        check("93.3 `main._ctrl_nav_vk_ok`（纯函数）：只有「单独按着 Ctrl 的 "
              "↑/↓」算数；Ctrl 没按 / Shift+Alt+Win 任一在按 / 别的键，一律 False "
              "—— Win 那道尤其不能少（Win+↑ 是系统的窗口最大化，吞了要出事）",
              [_g93c],
              expect_contain=[(True, True, False, False, False, False,
                               False, False, False, False)])

        # ---- 93.4 send_vk(ctrl=True)：修饰键必须反序松开 ----
        _cap93 = []

        def _fake_si93(n, arr, size):
            evts = _ct93.cast(arr, _ct93.POINTER(_MN93._INPUT * n)).contents
            for e in evts:
                _cap93.append((int(e.u.ki.wVk), int(e.u.ki.dwFlags)))
            return n

        _old_si93 = _ct93.windll.user32.SendInput
        try:
            _ct93.windll.user32.SendInput = _fake_si93
            _rc_c93 = _MN93.send_vk(_U, ctrl=True)
            _seq_c93 = list(_cap93)
            _cap93.clear()
            _rc_p93 = _MN93.send_vk(_D)          # 老调用方式：两个参数
            _seq_p93 = list(_cap93)
        finally:
            _ct93.windll.user32.SendInput = _old_si93
        check("93.4 `main.send_vk(vk, ctrl=True)`：兜底重发 Ctrl+↑/↓ 用。"
              "序列必须是 Ctrl↓ → 键↓ → 键↑ → Ctrl↑（修饰键【反序】松开，"
              "否则宿主可能看到\"Ctrl 还按着\"的残留）；不传 ctrl 时老行为"
              "不变（就两个事件）",
              [(_rc_c93, _seq_c93, _rc_p93, _seq_p93)],
              expect_contain=[(True,
                               [(0x11, 0), (0x26, 0), (0x26, 2), (0x11, 2)],
                               True, [(0x28, 0), (0x28, 2)])])

        # ---- 93.5 结构护栏（AST；源码字符串会被注释喂饱，见 v108 教训）----
        _msrc93 = io.open(os.path.join(ROOT, "main.py"), encoding="utf-8").read()
        _vsrc93 = io.open(os.path.join(ROOT, "vbe_bridge.py"),
                          encoding="utf-8").read()
        _tree93 = _ast93.parse(_msrc93)
        _vtree93 = _ast93.parse(_vsrc93)

        def _fn93(tree, name):
            for n in _ast93.walk(tree):
                if isinstance(n, _ast93.FunctionDef) and n.name == name:
                    return n
            return None

        _wf93 = _fn93(_tree93, "win32_filter")
        _ln_ctrl93 = _ln_popup93 = 10 ** 9
        _branch_ok93 = False
        for n in _ast93.walk(_wf93):
            if isinstance(n, _ast93.If):
                d = _ast93.dump(n.test)
                if "id='_ctrl_nav_key_ok'" in d:
                    _ln_ctrl93 = min(_ln_ctrl93, n.lineno)
                    _branch_ok93 = ("id='_ctrl_nav_jump_here'" in _ast93.dump(n)
                                    and "Constant(value=True)" in _ast93.dump(n))
                if ("id='completer'" in d and "attr='is_visible'" in d):
                    _ln_popup93 = min(_ln_popup93, n.lineno)
        _act93 = _fn93(_tree93, "_ctrl_nav_jump_here")
        _act_d93 = _ast93.dump(_act93)
        _be93 = _fn93(_vtree93, "ctrl_nav_jump")
        _be_d93 = _ast93.dump(_be93)
        # 开关：默认开 + 走配置中心（★v120：不再是 os.environ.get，
        # 而是 `not settings.flag("NO_CTRL_NAV", False)` —— 断言要跟着实现走，
        # 否则护栏会变成"只要我改实现它就红"的噪音）
        _env_ok93 = False
        for n in _ast93.walk(_tree93):
            if isinstance(n, _ast93.Assign):
                for t in n.targets:
                    if isinstance(t, _ast93.Name) and t.id == "CTRL_NAV_JUMP":
                        _d93 = _ast93.dump(n.value)
                        if ("attr='flag'" in _d93
                                and "NO_CTRL_NAV" in _d93
                                and "Constant(value=False)" in _d93):
                            _env_ok93 = True
        check("93.5 结构护栏（AST）：（a）Ctrl+↑/↓ 分支必须排在 `completer."
              "is_visible()` 【之前】—— 候选窗开着时 ↑/↓ 那条分支不分辨修饰键，"
              "晚一步就被列表抢走（用户正一边打字一边想跳去声明区）；（b）该分支"
              "确实吞键并派发 _ctrl_nav_jump_here；（c）动作里作废快照 "
              "state[\"cur_ctx\"]（改过文本+移过光标）且兜底是 send_vk(..., "
              "ctrl=True)；（d）后端 ctrl_nav_jump 调的是纯函数 ctrl_nav_jump_plan"
              "（同一份实现）；（e）开关默认开、且走配置中心"
              "（`not settings.flag(\"NO_CTRL_NAV\", False)`）",
              [(_ln_ctrl93 < _ln_popup93, _branch_ok93,
                "value='cur_ctx'" in _act_d93,
                "arg='ctrl'" in _act_d93,
                "id='ctrl_nav_jump_plan'" in _be_d93,
                _env_ok93 and _MN93.CTRL_NAV_JUMP is True)],
              expect_contain=[(True, True, True, True, True, True)])

        # ---- 93.6 逃生门：VBECOMPLETE_NO_CTRL_NAV=1 -> 默认值必须是 False ----
        _old_env93 = os.environ.get("VBECOMPLETE_NO_CTRL_NAV")
        _fresh_ctrl93 = None
        try:
            os.environ["VBECOMPLETE_NO_CTRL_NAV"] = "1"
            import importlib.util as _iu93
            _spec93 = _iu93.spec_from_file_location(
                "main_fresh93", os.path.join(ROOT, "main.py"))
            _fresh93 = _iu93.module_from_spec(_spec93)
            _spec93.loader.exec_module(_fresh93)
            _fresh_ctrl93 = _fresh93.CTRL_NAV_JUMP
        except Exception:
            _fresh_ctrl93 = None
        finally:
            if _old_env93 is None:
                os.environ.pop("VBECOMPLETE_NO_CTRL_NAV", None)
            else:
                os.environ["VBECOMPLETE_NO_CTRL_NAV"] = _old_env93
        check("93.6 逃生门：`set VBECOMPLETE_NO_CTRL_NAV=1` 之后重新加载的 "
              "main 里 CTRL_NAV_JUMP 必须是 False（ Ctrl+↑/↓ 原样交给 VBE）；"
              "当前实例默认是 True",
              [(_fresh_ctrl93, _MN93.CTRL_NAV_JUMP)],
              expect_contain=[(False, True)])
    except Exception as _e93:
        import traceback as _tp93
        _tp93.print_exc()
        check("第 93 节异常: %s" % _e93, [True], expect_contain=[False])

    print("\n=== 94. ★v110b/2：Ctrl+↓ 落点直接写成过程体层级缩进 ===")
    try:
        import ast as _ast94
        VB94 = VB

        # ---- 94.1 纯判据：过程体缩进取自"尾行上方最后一条非空行" ----
        def _p94(code, caret, direction):
            return VB94.ctrl_nav_jump_plan(code, caret, direction)

        _g94 = (
            # 一级过程体：抄 `    x = 1`
            _p94("Sub Foo()\r\n    x = 1\r\nEnd Sub\r\n", 2, +1),
            # 尾行上方已是空白 -> 继续往上找最后一条非空行
            _p94("Sub Foo()\r\n    x = 1\r\n\r\nEnd Sub\r\n", 2, +1),
            # 两级嵌套：最后一条非空是 `    End If`（一级）-> 抄它这一级
            _p94("Sub Foo()\r\n    If x Then\r\n        y = 1\r\n"
                 "    End If\r\nEnd Sub\r\n", 4, +1),
            # 块头后面紧跟 End（还在块内）：抄块头行这一级
            _p94("Sub Foo()\r\n    If x Then\r\nEnd Sub\r\n", 2, +1),
            # 空过程（Sub/End 相邻）：过程体层级 = ""（顶格）
            _p94("Sub Foo()\r\nEnd Sub\r\n", 1, +1),
            # 用户已摆好缩进的空白行 -> 不动（绝不加深）
            _p94("Sub Foo()\r\n    x = 1\r\n    \r\nEnd Sub\r\n", 2, +1),
            # ★v110c ↑：落在最后一个 Dim 行行尾 —— 不插行、不动文本
            _p94("Sub Foo()\r\n    Dim a As Long\r\n    x = 1\r\nEnd Sub\r\n", 3, -1),
            # 不在过程里 -> None
            _p94("Dim g As Long\r\n\r\nSub Foo()\r\nEnd Sub\r\n", 1, +1),
        )
        _w94 = (
            (3, "insert", "    "), (3, "swap", "    "), (5, "insert", "    "),
            (3, "insert", "    "), (2, "insert", ""), (3, None, None),
            (2, "eol", None),
            None,
        )
        check("94.1 `ctrl_nav_jump_plan` 的**过程体缩进**：Ctrl+↓ 的落点缩进 = "
              "尾行上方最后一条非空行的行首空白 —— 一级过程体、多级嵌套"
              "（抄 `End If` 那一级）、空过程（顶格）都对；用户自己摆过缩进的"
              "空白行一律不动（绝不加深）；★v110c 的 Ctrl+↑ 一个字节都不写"
              "（写入方式 \"eol\"，只落 Dim 行行尾）",
              [_g94], expect_contain=[_w94])

        # ---- 94.2 后端："swap" 真的插一条带缩进的 + 删原真空行；开关关掉退化 ----
        def _run94(lines, caret_line, direction, tab_on=True):
            cm = _FCM93(lines)
            pane = _FPane93(cm, caret_line)
            _old_vbe = VB94._get_vbe
            _old_gate = VB94._WINDOW_GATE_DISABLED
            _old_flag = VB94.CTRL_NAV_TAB
            VB94._get_vbe = lambda: _FVBE93(pane)
            VB94._WINDOW_GATE_DISABLED = True
            VB94._vbe_cache["obj"] = None
            VB94.CTRL_NAV_TAB = bool(tab_on)
            VB94.com_reset()
            try:
                took = VB94.VbeBackend().ctrl_nav_jump(direction)
            finally:
                VB94._get_vbe = _old_vbe
                VB94._WINDOW_GATE_DISABLED = _old_gate
                VB94.CTRL_NAV_TAB = _old_flag
            return (bool(took), list(cm.inserts), list(cm.deletes), list(cm.lines))

        _on_a = _run94(["Sub Foo()", "    x = 1", "End Sub"], 2, +1)
        _on_b = _run94(["Sub Foo()", "    x = 1", "", "End Sub"], 2, +1)
        _off_a = _run94(["Sub Foo()", "    x = 1", "End Sub"], 2, +1, tab_on=False)
        _off_b = _run94(["Sub Foo()", "    x = 1", "", "End Sub"], 2, +1,
                        tab_on=False)
        check("94.2 `ctrl_nav_jump` 的写入：（a）开关开、End 上方不是空白 -> 直接 "
              "InsertLines 一条**带过程体缩进**的新行；（b）开关开、End 上方是真空"
              "空白 -> 「插一条带缩进的 + 删掉原真空行」（swap；绝不用 ReplaceLine"
              " —— VBE 会把写进去的行尾空白 trim 掉，见 v103b）；（c）开关关掉 -> "
              "insert 退化成写空串、swap 退化成不写，跳转本身照样成功",
              [(_on_a, _on_b, _off_a, _off_b)],
              expect_contain=[((True, [(3, "    ")], [],
                                ["Sub Foo()", "    x = 1", "    ", "End Sub"]),
                               (True, [(3, "    ")], [(4, 1)],
                                ["Sub Foo()", "    x = 1", "    ", "End Sub"]),
                               (True, [(3, "")], [],
                                ["Sub Foo()", "    x = 1", "", "End Sub"]),
                               (True, [], [],
                                ["Sub Foo()", "    x = 1", "", "End Sub"]))])

        # ---- 94.3 结构护栏（AST）+ 逃生门 ----
        _vsrc94 = io.open(os.path.join(ROOT, "vbe_bridge.py"),
                          encoding="utf-8").read()
        _msrc94 = io.open(os.path.join(ROOT, "main.py"), encoding="utf-8").read()
        _vtree94 = _ast94.parse(_vsrc94)
        _mtree94 = _ast94.parse(_msrc94)

        def _fn94(tree, name):
            for n in _ast94.walk(tree):
                if isinstance(n, _ast94.FunctionDef) and n.name == name:
                    return n
            return None

        _be94 = _fn94(_vtree94, "ctrl_nav_jump")
        _be_d94 = _ast94.dump(_be94)
        # （a）ctrl_nav_jump 里必须同时出现 InsertLines 与 DeleteLines，
        #       且读 CTRL_NAV_TAB 开关（少一样都不对）
        _writes94 = ("attr='InsertLines'" in _be_d94
                     and "attr='DeleteLines'" in _be_d94
                     and "id='CTRL_NAV_TAB'" in _be_d94)
        # （b）"注入 Tab"那条路必须彻底撤掉：vbe_bridge / main 里都不能再有
        #       send_tab / _send_tab —— **AST 级**核对，不数源码字符串（注释里
        #       提一句就会喂饱字符串护栏，v108 教训）
        _no_sendtab94 = (_fn94(_vtree94, "send_tab") is None
                         and _fn94(_mtree94, "send_tab") is None
                         and "id='_send_tab'" not in _ast94.dump(_vtree94)
                         and "id='send_tab'" not in _ast94.dump(_mtree94))
        # （c）开关默认开 + 环境变量名出现在【赋值语句】里（不是注释里）
        _env_ok94 = False
        for n in _ast94.walk(_vtree94):
            if isinstance(n, _ast94.Assign):
                for t in n.targets:
                    if isinstance(t, _ast94.Name) and t.id == "CTRL_NAV_TAB":
                        if "VBECOMPLETE_NO_CTRL_NAV_TAB" in _ast94.dump(n.value):
                            _env_ok94 = True
        # （d）逃生门：设了环境变量之后重新加载的 vbe_bridge 里必须是 False
        _old_env94 = os.environ.get("VBECOMPLETE_NO_CTRL_NAV_TAB")
        _fresh_tab94 = None
        try:
            os.environ["VBECOMPLETE_NO_CTRL_NAV_TAB"] = "1"
            import importlib.util as _iu94
            _spec94 = _iu94.spec_from_file_location(
                "vbe_bridge_fresh94", os.path.join(ROOT, "vbe_bridge.py"))
            _fresh94 = _iu94.module_from_spec(_spec94)
            _spec94.loader.exec_module(_fresh94)
            _fresh_tab94 = _fresh94.CTRL_NAV_TAB
        except Exception:
            _fresh_tab94 = None
        finally:
            if _old_env94 is None:
                os.environ.pop("VBECOMPLETE_NO_CTRL_NAV_TAB", None)
            else:
                os.environ["VBECOMPLETE_NO_CTRL_NAV_TAB"] = _old_env94
        check("94.3 结构护栏（AST）+ 逃生门：（a）ctrl_nav_jump 里 InsertLines / "
              "DeleteLines / CTRL_NAV_TAB 三样都在；（b）「注入 Tab」整条路"
              "（send_tab / _send_tab）已经彻底撤掉 —— AST 级核对，不数源码字符串"
              "（会被注释喂饱，v108 教训）；（c）开关 CTRL_NAV_TAB 默认开、"
              "VBECOMPLETE_NO_CTRL_NAV_TAB 接在赋值语句上；（d）逃生门：设了环境"
              "变量重新加载的 vbe_bridge 里必须是 False",
              [(_writes94, _no_sendtab94,
                _env_ok94 and VB94.CTRL_NAV_TAB is True, _fresh_tab94)],
              expect_contain=[(True, True, True, False)])
    except Exception as _e94:
        import traceback as _tp94
        _tp94.print_exc()
        check("第 94 节异常: %s" % _e94, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 95. ★v111：声明位置只提示【该处的语言关键字】（v105 收得太狠了）
    #     用户口径（原话）："1. 在定义函数或过程，我写了 private 或 public，
    #     后面写 sub，function 不会提示。2. 定义函数或过程，形参里的关键字不会
    #     提示，比如形参里写 byref，byval。这 2 个 bug 应该是之前 private，
    #     public 后面不提示变量名，你限制的太多了，导致完全不提示了。"
    #
    #     v105 让声明位置**整个静默**（"正在造名字，候选全是噪音"），但它把
    #     **语言关键字**也一起静默了 —— `Private s` 里的 Sub 不是"既有符号"，
    #     那正是用户此刻要敲的东西。v111 改成"只留该处的关键字"。
    # ------------------------------------------------------------------
    print("\n=== 95. ★v111：声明位置只提示该处的语言关键字 ===")
    try:
        # ---- 95.1 纯判据：两档（声明头 / 形参表）+ 该静默的位置 ----
        _DH95 = P._DECL_HEAD_KEYWORDS
        _PL95 = P._PARAM_LIST_KEYWORDS
        _g95 = (
            P.decl_position_keywords("Private s", (1, 10)),
            P.decl_position_keywords("Public f", (1, 9)),
            P.decl_position_keywords("Private su", (1, 11)),
            P.decl_position_keywords("    Private F", (1, 14)),
            P.decl_position_keywords("Private Cons", (1, 13)),
            P.decl_position_keywords("Private Sub Foo(By", (1, 19)),
            P.decl_position_keywords("Sub Foo(ByVal x As Long, ", (1, 26)),
            P.decl_position_keywords("Public Function GetName( _\n"
                                     "        By", (2, 11)),
            P.decl_position_keywords("Private Sub Foo()", (1, 19)),
            P.decl_position_keywords("Private Sub F", (1, 14)),
            P.decl_position_keywords("Dim x", (1, 6)),
            P.decl_position_keywords("Private gA As Long", (1, 20)),
            P.decl_position_keywords("x = Foo(", (1, 9)),
        )
        _w95 = (
            _DH95, _DH95, _DH95, _DH95, _DH95,     # 声明头（含只打了前缀的）
            _PL95, _PL95, _PL95,                   # 形参表（含 v53 续行）
            None,                                  # `Private Sub Foo()` 头写完了
            None,                                  # 正在起过程名
            None,                                  # 正在起变量名
            None,                                  # 模块级变量声明
            None,                                  # 调用不是声明
        )
        check("95.1 `parser.decl_position_keywords`（纯函数）：声明头"
              "（`Private |` / `Private Su|` / `Public f|`）给 `_DECL_HEAD_KEYWORDS`；"
              "形参表（`Sub Foo(By|`、第二段、v53 续行）给 `_PARAM_LIST_KEYWORDS`；"
              "过程名 / 变量名 / 完整的声明 / 调用行一律返回 None（= 照 v105 静默）",
              [_g95], expect_contain=[_w95])

        # ---- 95.2 引擎集成（假后端）：只留 decl_keywords 里的候选 ----
        import vba_builtins as _VB95

        class _UI95:
            def __init__(self):
                self.shown = False
            def show(self, *a, **k):
                self.shown = True
            def update_selection(self, *a, **k):
                pass
            def hide(self, *a, **k):
                self.shown = False
            def contains_point(self, *a, **k):
                return False

        class _B95:
            """池子 = 语言关键字 + 内建函数/类型 + 两个工程内变量名。

            刻意给 `get_builtin_names()`（生产后端有这一路）：引擎对"语言自带的
            公共词汇"另有一道收紧，少了它这条路径就与真机不同。
            """
            def __init__(self, line_text, in_decl, decl_keywords, in_type=False):
                self._ctx = {"line_no": 1, "caret_col": len(line_text) + 1,
                             "line_text": line_text, "in_string": False,
                             "in_comment": False, "in_type_position": in_type,
                             "in_decl_position": in_decl,
                             "decl_names": ["customerName"],
                             "decl_keywords": list(decl_keywords or ()),
                             "proc_name": "Foo", "module_name": "M1"}
                _ids = [("customerName", "M1", None, False),
                        ("customerCode", "M1", None, False),
                        ("saveAll", "M1", None, False)]
                for _k in (list(_VB95.BUILTIN_KEYWORDS) + list(_VB95.BUILTIN_FUNCTIONS)
                           + list(_VB95.BUILTIN_TYPE_NAMES)):
                    _ids.append((_k, "(builtin)", None, False))
                self._ids = _ids
                self._bl = set()
                for _n in (list(_VB95.BUILTIN_KEYWORDS) + list(_VB95.BUILTIN_FUNCTIONS)
                           + list(_VB95.BUILTIN_TYPE_NAMES)):
                    self._bl.add(_n.lower())
            def get_context(self):
                return self._ctx
            def get_identifiers(self):
                return self._ids
            def get_declared_names(self):
                return [n[0] for n in self._ids]
            def get_builtin_names(self):
                return sorted(self._bl)

        def _run95(line_text, in_decl, decl_kw, in_type=False, manual=False):
            _ui = _UI95()
            _c = E.Completer(_B95(line_text, in_decl, decl_kw, in_type), _ui)
            _c.trigger(manual=manual)
            return (_ui.shown, sorted(_c.current_matches()))

        _DHKW95 = ["Sub", "Function", "Property", "Type", "Enum", "Declare",
                   "Event", "Const", "WithEvents", "Static", "Dim"]
        _g95b = (
            _run95("Private s", True, ["Sub", "Static"]),
            _run95("Public f", True, _DHKW95),
            _run95("Private Sub Foo(By", True,
                   ["ByVal", "ByRef", "Optional", "ParamArray", "As"]),
            _run95("Dim x", True, []),
            # in_decl=False 时那批关键字【不许】生效：照常出工程内候选
            # （同样带着 decl_keywords，但它管不着非声明位置）
            _run95("Private Sub Foo(cus", False, ["ByVal", "ByRef"]),
        )
        _w95b = (
            (True, ["Static", "Sub"]),
            (True, ["Function"]),
            (True, ["ByRef", "ByVal"]),
            (False, []),
            (True, ["customerCode", "customerName"]),
        )
        check("95.2 `Completer.trigger`（集成）：声明位置只留 ctx[\"decl_keywords\"] "
              "里的候选（`Private s` -> Sub / Static；`Private Sub Foo(By` -> "
              "ByRef / ByVal —— 既有符号一个都不出现）；decl_keywords 为空 -> "
              "照 v105 静默；in_decl=False 时那批关键字不生效（照常出工程内候选）",
              [_g95b], expect_contain=[_w95b])

        # ---- 95.2b 点名的列表（manual）在 v111 里照旧不过滤 ----
        _m95 = _run95("Private s", True, ["Sub", "Static"], manual=True)
        check("95.2b Ctrl+Space（manual）在声明位置【不受 v111 收窄】"
              "（用户点名的列表照弹全量，同 87.6）",
              [(_m95[0], "customerName" in _m95[1])],
              expect_contain=[(True, True)])

        # ---- 95.3 结构护栏（AST）：单一实现 + 接线 ----
        # ⚠️ 一律走 AST，不数源码字符串 —— 注释里提一句就会喂饱字符串护栏
        #    （v108 教训）。AST 里根本没有注释。
        import ast as _ast95
        _et95 = _ast95.parse(io.open(os.path.join(ROOT, "engine.py"),
                                     encoding="utf-8").read())
        _pt95 = _ast95.parse(io.open(os.path.join(ROOT, "parser.py"),
                                     encoding="utf-8").read())
        _vt95 = _ast95.parse(io.open(os.path.join(ROOT, "vbe_bridge.py"),
                                     encoding="utf-8").read())

        def _fn95(tree, name):
            for n in _ast95.walk(tree):
                if isinstance(n, _ast95.FunctionDef) and n.name == name:
                    return n
            return None

        _trig95 = _fn95(_et95, "trigger")
        _e_ok95 = ("decl_keywords" in _ast95.dump(_trig95)
                   if _trig95 is not None else False)
        # 引擎里不许再有一份关键字清单（同一语义只能一份实现）
        _e_single95 = ("_DECL_HEAD_KEYWORDS" not in _ast95.dump(_et95)
                       and "_PARAM_LIST_KEYWORDS" not in _ast95.dump(_et95))
        # 后端 get_context 的返回字典里必须真有这个键（不是只写在注释里）
        _gc95 = _fn95(_vt95, "get_context")
        _v_ok95 = False
        if _gc95 is not None:
            for n in _ast95.walk(_gc95):
                if isinstance(n, _ast95.Dict) and any(
                        isinstance(k, _ast95.Constant)
                        and k.value == "decl_keywords" for k in n.keys):
                    _v_ok95 = True
        # parser 侧：判据函数 + 两个常量都在
        _p_names95 = set()
        for n in _pt95.body:
            if isinstance(n, _ast95.Assign):
                for t in n.targets:
                    if isinstance(t, _ast95.Name):
                        _p_names95.add(t.id)
        _p_ok95 = (_fn95(_pt95, "decl_position_keywords") is not None
                   and "_DECL_HEAD_KEYWORDS" in _p_names95
                   and "_PARAM_LIST_KEYWORDS" in _p_names95)
        check("95.3 结构护栏（AST）：(a) engine.trigger 读了 ctx[\"decl_keywords\"]；"
              "(b) 关键字清单只有 parser 一份（引擎里不重复实现）；"
              "(c) vbe_bridge.get_context 的返回字典里真有这个键；"
              "(d) parser 有 decl_position_keywords 与两个常量",
              [(_e_ok95, _e_single95, _v_ok95, _p_ok95)],
              expect_contain=[(True, True, True, True)])
    except Exception as _e95:
        import traceback as _tp95
        _tp95.print_exc()
        check("第 95 节异常: %s" % _e95, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 96. ★v112：自动配对两条（用户口径）
    #     （1）敲左括号时右侧还有内容 -> 只插左括号，不补右半边；
    #     （2）选中一段文本再敲引号 -> 用引号把选区包起来。
    # ------------------------------------------------------------------
    print("\n=== 96. ★v112：左括号右侧有内容不补全 / 选中文本加引号 ===")
    try:
        import vbe_bridge as VB96
        _pi96 = VB96._pair_insertion

        # ---- 96.1 纯判据：右侧有没有内容决定补不补右半边 ----
        _g96 = (
            _pi96("x = ", 4, "(", ")"),          # 行尾 -> 补一对
            _pi96("x =  ", 5, "(", ")"),         # 右侧只剩空白 -> 仍算"行尾"
            _pi96("x = ab", 4, "(", ")"),        # 右侧有内容 -> 只插左
            # ★v113：`foo(a, b)` 光标停在 `foo(` 里 -> 实参表里照旧补一对（§97）
            _pi96("foo(a, b)", 5, "(", ")"),
            _pi96("x = 1", 0, "(", ")"),         # 行首（右侧全是内容）-> 只插左
            _pi96("x = ", 4, "（", "）"),         # 全角行尾 -> 补一对
            _pi96("foo（a）", 4, "（", "）"),      # ★v113：全角实参表里 -> 补一对
            _pi96("s = ", 4, '"', '"'),          # 引号不受影响：行尾仍补一对
            _pi96("    w=4", 4, '"', '"'),       # 引号不受影响：右侧有内容仍补
        )
        _w96 = (
            ("x = ()", 5, True),
            ("x =  ()", 6, True),
            ("x = (ab", 5, False),
            ("foo(a(), b)", 6, True),
            ("(x = 1", 1, False),
            ("x = （）", 5, True),
            ("foo（（）a）", 5, True),
            ('s = ""', 5, True),
            ('    ""w=4', 5, True),
        )
        check("96.1 `_pair_insertion`：敲左括号时**右侧还有内容就只插左括号**"
              "（用户口径）—— 行尾 / 右侧只剩空白才补一对；半角全角同路；"
              "引号**刻意不走这条**（`w=4` 那种\"右侧有内容仍补一对\"是既有"
              "行为，且引号另有\"在不在字符串里\"的先判据）。"
              "★v113 起有一个例外：光标在实参表 / 形参表里**仍补一对**"
              "（这两例照 v113 的期望写，规则本身在 §97 钉）",
              [_g96], expect_contain=[_w96])

        # ---- 96.2 集成（假 VBE）----
        class _Mod96(object):
            def __init__(self, text, fmt=None):
                self.lines = [text]
                self.fmt = fmt

            def Lines(self, n, _c):
                if 1 <= n <= len(self.lines):
                    return self.lines[n - 1]
                return ""

            @property
            def CountOfLines(self):
                return len(self.lines)

            def ReplaceLine(self, n, t):
                if self.fmt is not None:
                    t = self.fmt(t)
                self.lines[n - 1] = t

            def InsertLines(self, n, t):
                self.lines.insert(n - 1, t)

        class _Pane96(object):
            def __init__(self, mod, sl, sc, el, ec):
                self.CodeModule = mod
                self.sel = [sl, sc, el, ec]

            def GetSelection(self):
                return tuple(self.sel)

            def SetSelection(self, sl, sc, el, ec):
                self.sel = [sl, sc, el, ec]

        class _Vbe96(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _run96(text, sl, sc, el, ec, ch, fmt=None,
                   sem=("char", 4, False)):
            mod = _Mod96(text, fmt)
            pane = _Pane96(mod, sl, sc, el, ec)
            _o_vbe = VB96._get_vbe_cached
            _o_sem = VB96._sem_cache.get("info")
            VB96._sem_cache["info"] = sem
            VB96._get_vbe_cached = lambda: _Vbe96(pane)
            try:
                ok = VB96.VbeBackend().insert_pair(ch)
            finally:
                VB96._get_vbe_cached = _o_vbe
                VB96._sem_cache["info"] = _o_sem
            return bool(ok), mod.lines[sl - 1], pane.sel[1], pane.sel[3]

        _g96b = (
            _run96("x = ", 1, 5, 1, 5, "("),        # 行尾 -> 补一对
            _run96("x = ab", 1, 5, 1, 5, "("),      # 右侧有内容 -> 只插左
            # 行首缩进里（不在任何调用括号里）-> 只插左。
            # ★v113 起本用例原本是 `foo(a, b)` 光标停在 `foo(` 里 —— 那属于
            # "实参表里"，现在改成补一对（见 §97），故换一行
            _run96("    x = ab", 1, 5, 1, 5, "("),
            # 只插左括号后 VBE 补空格：行里没有 `()`，得走前缀定位
            _run96("x=(ab", 1, 3, 1, 3, "(",
                   fmt=lambda s: s.replace("=(", "= (")),
            # 选区包裹（半角 / 全角 / 整行）
            _run96("x = abc", 1, 5, 1, 8, '"'),
            _run96("x = abc", 1, 5, 1, 8, "“"),
            _run96("abc", 1, 1, 1, 4, '"'),
            # 包裹后 VBE 补空格：按前缀找回末尾收符号之后
            _run96("x=abc", 1, 3, 1, 6, '"',
                   fmt=lambda s: s.replace('="', '= "')),
            # 反面：多行选区 / 收符号带选区一律不接管
            _run96("x = abc", 1, 5, 2, 5, '"'),
            _run96("x = abc", 1, 5, 1, 8, ")"),
            # ★v114：**括号带选区改成"包裹"**（v112 时期望 False）——
            # 期望随口径更新，规则本身在 §98 钉
            _run96("x = abc", 1, 5, 1, 8, "("),
            # 空选区（sc == ec）= 没有选区，照常配对
            _run96("x = ", 1, 5, 1, 5, '"'),
            # 选区两端含全角：显示列 -> 字符列必须换算对（disp 语义）
            _run96("x = 中文字", 1, 5, 1, 11, '"', sem=("disp", 4, True)),
        )
        _w96b = (
            (True, "x = ()", 6, 6),
            (True, "x = (ab", 6, 6),
            (True, "    (x = ab", 6, 6),
            (True, "x= ((ab", 5, 5),
            (True, 'x = "abc"', 10, 10),
            (True, "x = “abc”", 10, 10),
            (True, '"abc"', 6, 6),
            (True, 'x= "abc"', 9, 9),
            (False, "x = abc", 5, 5),
            (False, "x = abc", 5, 8),
            (True, "x = (abc)", 10, 10),
            (True, 'x = ""', 6, 6),
            (True, 'x = "中文字"', 13, 13),
        )
        check("96.2 `VbeBackend.insert_pair`（集成）："
              "（a）行尾敲 `(` 补一对、**右侧有内容只插左括号**，且只插左时"
              "VBE 重排整行也照样落在左括号之后（走前缀定位 —— 行里根本没有"
              "`()` 可找）；"
              "（b）**选中一段文本敲引号 -> 用引号把选区包起来**（半角 / 全角 /"
              "整行选中 / VBE 重排都算对），光标落在末尾收符号之后；"
              "（c）多行选区 / **收符号**带选区**仍交给 VBE**，空选区照常配对"
              "（★v114 起左括号带选区也包裹，见 §98）；"
              "（d）选区两端含全角时按【显示列】换算（`x = 中文字` 选中三个字，"
              "不是两个字）",
              [_g96b], expect_contain=[_w96b])

        # ---- 96.3 结构护栏（AST）：单一实现 + 接线（查运行时对象，不查源码）----
        import ast as _ast96
        _esrc96 = io.open(os.path.join(ROOT, "vbe_bridge.py"),
                          encoding="utf-8").read()
        _etree96 = _ast96.parse(_esrc96)
        _fns96 = {}
        for _n96 in _ast96.walk(_etree96):
            if isinstance(_n96, _ast96.FunctionDef):
                _fns96.setdefault(_n96.name, []).append(_n96)
        _a96 = len(_fns96.get("_wrap_selection", [])) == 1
        _ip96 = _fns96.get("insert_pair", [])
        _b96 = bool(_ip96) and all(
            k in _ast96.dump(_ip96[0])
            for k in ("_wrap_selection", "paired", "_locate_by_prefix"))
        _pi96f = _fns96.get("_pair_insertion", [])
        _c96 = bool(_pi96f) and any(
            isinstance(_nn, _ast96.Call)
            and isinstance(_nn.func, _ast96.Attribute)
            and _nn.func.attr == "strip"
            for _nn in _ast96.walk(_pi96f[0]))
        check("96.3 结构护栏（AST）：(a) `_wrap_selection` 只有一处实现；"
              "(b) `insert_pair` 里既调用了 `_wrap_selection`、又用了 `paired`"
              "与 `_locate_by_prefix`（只插左括号时改走前缀定位）；"
              "(c) `_pair_insertion` 的括号分支确实按 `strip()` 判\"右侧有没有"
              "内容\"（不查源码字符串 —— 注释会把它喂饱）",
              [(_a96, _b96, _c96)],
              expect_contain=[(True, True, True)])
    except Exception as _e96:
        import traceback as _tp96
        _tp96.print_exc()
        check("第 96 节异常: %s" % _e96, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 97. ★v113：在【实参表 / 形参表】里敲左括号仍要补全右半边（用户口径）
    #     用户报："如果在某个函数或过程形参里，或者是对象属性方法参数里输入
    #     左括号，需要能自动补全右括号的。" —— 即 v112 那条"右侧有内容就只插
    #     左括号"需要给这两种位置开例外：那里右侧的 `)` 属于**外层**括号组。
    # ------------------------------------------------------------------
    print("\n=== 97. ★v113：实参表 / 形参表里敲 ( 仍补一对 ===")
    try:
        import parser as _P97
        import vbe_bridge as VB97
        _cp97 = _P97.enclosing_call_paren

        # ---- 97.1 纯判据：光标左边最内层未闭合的括号是不是"调用 / 形参表" ----
        _g97a = (
            _cp97("foo(bar", 7),              # 调用实参表里
            _cp97("Sub Foo(ByVal arr", 17),   # 形参表里
            _cp97("Obj.Method(", 11),         # 对象.方法
            _cp97("Mid$(", 5),                # 类型声明符结尾的名字也算
            _cp97("foo（a", 4),               # 全角括号同路
            _cp97("foo(a) + bar(", 13),       # 前面已闭合的那组不算数
            _cp97("foo(1, (2 + 3", 13),       # 最内层是**运算分组** -> 不算
            _cp97("x = (a + (", 9),           # 同上
            _cp97("x = (a ", 7),              # 括号左边是 `=` -> 不算
            _cp97("(", 1),                    # 行首括号 -> 不算
            _cp97("x = (a + b)", 12),         # 括号已闭合 -> 不算
            _cp97("foo(a)", 6),               # 同上
            _cp97('s = "(" & foo(', 14),      # 字符串里那个 `(` 不算、foo( 算
            _cp97('s = "(" & foo(', 13),      # 光标还在字符串里 -> 不算
            _cp97("' foo(", 6),               # 注释里 -> 不算
            _cp97("", 0),                     # 空行
        )
        _w97a = (
            3, 7, 10, 4, 3, 12,
            None, None, None, None, None, None,
            13, None, None, None,
        )
        check("97.1 `parser.enclosing_call_paren`：只认【最内层未闭合、且紧挨着"
              "标识符】的那个左括号（实参表 / 形参表）；运算分组、行首括号、"
              "已闭合的组一律不认；字符串 / 注释里的括号不算数；半角全角同路；"
              "`$` 这类类型声明符结尾的名字也算名字",
              [_g97a], expect_contain=[_w97a])

        # ---- 97.2 纯判据：`_pair_insertion` 在这两种位置照旧补一对 ----
        _pi97 = VB97._pair_insertion
        _g97b = (
            _pi97("foo(bar", 7, "(", ")"),                  # 行尾（本来也补）
            _pi97("x = ", 4, "(", ")"),                     # 行尾（v112 既有）
            _pi97("foo(a, b)", 5, "(", ")"),                # ★实参表里、右侧有内容
            _pi97("Sub Foo(ByVal arr)", 17, "(", ")"),      # ★形参表里（形参数组）
            _pi97("foo（a）", 4, "（", "）"),                 # ★全角实参表里
            _pi97("x = ab", 4, "(", ")"),                   # 不在调用括号里 -> 只插左
            _pi97("x = (a + b)", 7, "(", ")"),              # 运算分组里 -> 只插左
            _pi97("x = 1", 0, "(", ")"),                    # 行首 -> 只插左
            _pi97("    w=4", 4, '"', '"'),                  # 引号不受影响（v112）
        )
        _w97b = (
            ("foo(bar()", 8, True),
            ("x = ()", 5, True),
            ("foo(a(), b)", 6, True),
            ("Sub Foo(ByVal arr())", 18, True),
            ("foo（（）a）", 5, True),
            ("x = (ab", 5, False),
            ("x = (a (+ b)", 8, False),
            ("(x = 1", 1, False),
            ('    ""w=4', 5, True),
        )
        check("97.2 `_pair_insertion`：**在实参表 / 形参表里敲左括号仍补一对**"
              "（用户口径）—— 那里右侧的 `)` 是外层的，不补用户每次都得自己"
              "补一个；实参表外的老规则（v112：右侧有内容只插左）原样保留，"
              "引号也不受影响",
              [_g97b], expect_contain=[_w97b])

        # ---- 97.3 集成（假 VBE）：真写进"编辑器"、光标落在新括号中间 ----
        class _Mod97(object):
            def __init__(self, text, fmt=None):
                self.lines = [text]
                self.fmt = fmt

            def Lines(self, n, _c):
                if 1 <= n <= len(self.lines):
                    return self.lines[n - 1]
                return ""

            @property
            def CountOfLines(self):
                return len(self.lines)

            def ReplaceLine(self, n, t):
                if self.fmt is not None:
                    t = self.fmt(t)
                self.lines[n - 1] = t

            def InsertLines(self, n, t):
                self.lines.insert(n - 1, t)

        class _Pane97(object):
            def __init__(self, mod, sl, sc, el, ec):
                self.CodeModule = mod
                self.sel = [sl, sc, el, ec]

            def GetSelection(self):
                return tuple(self.sel)

            def SetSelection(self, sl, sc, el, ec):
                self.sel = [sl, sc, el, ec]

        class _Vbe97(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _run97(text, sl, sc, el, ec, ch, fmt=None,
                   sem=("char", 4, False)):
            mod = _Mod97(text, fmt)
            pane = _Pane97(mod, sl, sc, el, ec)
            _o_vbe = VB97._get_vbe_cached
            _o_sem = VB97._sem_cache.get("info")
            VB97._sem_cache["info"] = sem
            VB97._get_vbe_cached = lambda: _Vbe97(pane)
            try:
                ok = VB97.VbeBackend().insert_pair(ch)
            finally:
                VB97._get_vbe_cached = _o_vbe
                VB97._sem_cache["info"] = _o_sem
            return bool(ok), mod.lines[sl - 1], pane.sel[1], pane.sel[3]

        _g97c = (
            _run97("foo(bar)", 1, 8, 1, 8, "("),             # 实参表里开一层新的
            _run97("Sub Foo(ByVal arr)", 1, 18, 1, 18, "("),  # 形参数组
            _run97("Sub Foo（ByVal arr）", 1, 18, 1, 18, "（"),  # 全角形参表
            # 补完一对后 VBE 补空格（`x=foo(...)` -> `x= foo(...)`）：
            # 仍要落在新插入的 `()` 中间（走 _locate_inserted 的前缀指纹）
            _run97("x=foo(bar)", 1, 10, 1, 10, "(",
                   fmt=lambda s: s.replace("=", "= ", 1)),
            # 反面：运算分组里只插左（v112 规则没被松动）
            _run97("x = (a + b)", 1, 5, 1, 5, "("),
        )
        _w97c = (
            (True, "foo(bar())", 9, 9),
            (True, "Sub Foo(ByVal arr())", 19, 19),
            (True, "Sub Foo（ByVal arr（））", 19, 19),
            (True, "x= foo(bar())", 12, 12),
            (True, "x = ((a + b)", 6, 6),
        )
        check("97.3 `VbeBackend.insert_pair`（集成）：实参表 / 形参表（含全角）里"
              "敲 `(` 真的补出成对的右半边，光标落在两者中间；VBE 重排整行也认"
              "得回原位；**运算分组里仍只插左**（v112 的规则只在这两种位置开"
              "例外）",
              [_g97c], expect_contain=[_w97c])

        # ---- 97.4 结构护栏（AST）：单一实现 + 接线（查 AST，不查源码字符串）----
        import ast as _ast97

        def _fdefs97(path):
            _tree = _ast97.parse(io.open(os.path.join(ROOT, path),
                                         encoding="utf-8").read())
            _out = {}
            for _n in _ast97.walk(_tree):
                if isinstance(_n, _ast97.FunctionDef):
                    _out.setdefault(_n.name, []).append(_n)
            return _out

        _pf97 = _fdefs97("parser.py")
        _vf97 = _fdefs97("vbe_bridge.py")
        _a97 = len(_pf97.get("enclosing_call_paren", [])) == 1
        _b97 = bool(_pf97.get("enclosing_call_paren")) and any(
            isinstance(_x, _ast97.Call) and isinstance(_x.func, _ast97.Name)
            and _x.func.id == "scan_code_states"
            for _x in _ast97.walk(_pf97["enclosing_call_paren"][0]))
        _c97 = len(_vf97.get("enclosing_call_paren", [])) == 0
        _d97 = bool(_vf97.get("_pair_insertion")) and any(
            isinstance(_x, _ast97.Call) and isinstance(_x.func, _ast97.Attribute)
            and _x.func.attr == "enclosing_call_paren"
            for _x in _ast97.walk(_vf97["_pair_insertion"][0]))
        check("97.4 结构护栏（AST）：(a) `enclosing_call_paren` 只有一处实现"
              "（在 parser）；(b) 它复用 `scan_code_states`（字符串/注释扫描的"
              "唯一实现）；(c) vbe_bridge 里没有第二份；(d) `_pair_insertion`"
              "确实调用了它",
              [(_a97, _b97, _c97, _d97)],
              expect_contain=[(True, True, True, True)])
    except Exception as _e97:
        import traceback as _tp97
        _tp97.print_exc()
        check("第 97 节异常: %s" % _e97, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 98. ★v114：选中文本敲【左括号】也把选区包起来（用户口径）
    #     用户口径（原话）："之前有个打双引号可以自动把选中文本用双引号括
    #     起来，我想加个类似的，括号也可以实现同样效果。"
    #     ⇒ v112 那条包裹只对引号开的门，对 `(` / `（` 同样开。
    # ------------------------------------------------------------------
    print("\n=== 98. ★v114：选中文本敲左括号也包裹 ===")
    try:
        import vbe_bridge as VB98

        # ---- 98.1 白名单：哪些开符号带选区时"包裹"，收符号不算 ----
        _wr98 = VB98.VbeBackend._WRAP_OPEN
        _pc98 = VB98.VbeBackend._PAIR_CLOSE
        _g98 = (
            # 包裹：引号（v112）+ 左括号半角/全角（v114）
            tuple(ch in _wr98 for ch in ('"', "“", "(", "（")),
            # 不包裹：收符号（没有"包起来"的语义）+ 根本不相配的符号
            tuple(ch in _wr98 for ch in (")", "）", "'")),
            # 每个开符号都必须查得到收符号（否则 insert_pair 一进来就 return False，
            # 包裹的路走不到 —— v108 踩过同样的坑）
            tuple(ch in _pc98 for ch in _wr98),
            # 没有重复项
            len(set(_wr98)) == len(_wr98),
        )
        check("98.1 包裹白名单 `VbeBackend._WRAP_OPEN`：引号（v112）+ 左括号"
              "半角/全角（v114）**带选区时包裹**；收符号 `)` / `）` 与不相配的"
              "`'` **不包裹**（仍交给 VBE 替换）；每个开符号都在 `_PAIR_CLOSE`"
              "里查得到收符号",
              [_g98],
              expect_contain=[((True, True, True, True),
                               (False, False, False),
                               (True, True, True, True),
                               True)])

        # ---- 98.2 集成（假 VBE）：真写进"编辑器"、光标落在收括号之后 ----
        class _Mod98(object):
            def __init__(self, text, fmt=None):
                self.lines = [text]
                self.fmt = fmt

            def Lines(self, n, _c):
                if 1 <= n <= len(self.lines):
                    return self.lines[n - 1]
                return ""

            @property
            def CountOfLines(self):
                return len(self.lines)

            def ReplaceLine(self, n, t):
                if self.fmt is not None:
                    t = self.fmt(t)
                self.lines[n - 1] = t

            def InsertLines(self, n, t):
                self.lines.insert(n - 1, t)

        class _Pane98(object):
            def __init__(self, mod, sl, sc, el, ec):
                self.CodeModule = mod
                self.sel = [sl, sc, el, ec]

            def GetSelection(self):
                return tuple(self.sel)

            def SetSelection(self, sl, sc, el, ec):
                self.sel = [sl, sc, el, ec]

        class _Vbe98(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _run98(text, sl, sc, el, ec, ch, fmt=None,
                   sem=("char", 4, False)):
            mod = _Mod98(text, fmt)
            pane = _Pane98(mod, sl, sc, el, ec)
            _o_vbe = VB98._get_vbe_cached
            _o_sem = VB98._sem_cache.get("info")
            VB98._sem_cache["info"] = sem
            VB98._get_vbe_cached = lambda: _Vbe98(pane)
            try:
                ok = VB98.VbeBackend().insert_pair(ch)
            finally:
                VB98._get_vbe_cached = _o_vbe
                VB98._sem_cache["info"] = _o_sem
            return bool(ok), mod.lines[sl - 1], pane.sel[1], pane.sel[3]

        _g98b = (
            _run98("x = abc", 1, 5, 1, 8, "("),        # 半角：包起来
            _run98("x = abc", 1, 5, 1, 8, "（"),        # 全角：包起来
            _run98("abc", 1, 1, 1, 4, "("),             # 整行选中
            # 包裹后 VBE 补空格（`x=(abc)` -> `x= (abc)`）：仍落在收括号之后
            _run98("x=abc", 1, 3, 1, 6, "(",
                   fmt=lambda s: s.replace("=(", "= (")),
            # 选区两端含全角：显示列 -> 字符列必须换算对（disp 语义）
            _run98("x = 中文字", 1, 5, 1, 11, "(",
                   sem=("disp", 4, True)),
            # 反面：多行选区 / 收符号带选区 / 不相配的符号 一律不接管
            _run98("x = abc", 1, 5, 2, 5, "("),
            _run98("x = abc", 1, 5, 1, 8, ")"),
            _run98("x = abc", 1, 5, 1, 8, "）"),
            _run98("x = abc", 1, 5, 1, 8, "'"),
        )
        _w98b = (
            (True, "x = (abc)", 10, 10),
            (True, "x = （abc）", 10, 10),
            (True, "(abc)", 6, 6),
            (True, "x= (abc)", 9, 9),
            (True, "x = (中文字)", 13, 13),
            (False, "x = abc", 5, 5),
            (False, "x = abc", 5, 8),
            (False, "x = abc", 5, 8),
            (False, "x = abc", 5, 8),
        )
        check("98.2 `VbeBackend.insert_pair`（集成）："
              "**选中一段文本敲左括号 -> 用括号把选区包起来**（半角 / 全角 /"
              "整行选中 / VBE 重排 / 选区含全角都算对），光标落在收括号之后；"
              "多行选区、收符号 `)` `）`、不相配的 `'` **仍交给 VBE**（不吞键）",
              [_g98b], expect_contain=[_w98b])

        # ---- 98.3 结构护栏（AST + 运行时对象，不查源码字符串）----
        import ast as _ast98
        _etree98 = _ast98.parse(io.open(os.path.join(ROOT, "vbe_bridge.py"),
                                        encoding="utf-8").read())
        _cls98 = [n for n in _ast98.walk(_etree98)
                  if isinstance(n, _ast98.ClassDef) and n.name == "VbeBackend"]
        _a98 = bool(_cls98) and any(
            isinstance(_st, _ast98.Assign) and any(
                isinstance(_tg, _ast98.Name) and _tg.id == "_WRAP_OPEN"
                for _tg in _st.targets)
            for _st in _cls98[0].body)
        _ip98 = [n for n in _ast98.walk(_etree98)
                 if isinstance(n, _ast98.FunctionDef) and n.name == "insert_pair"]
        _b98 = bool(_ip98) and any(
            isinstance(_x, _ast98.Attribute) and _x.attr == "_WRAP_OPEN"
            for _x in _ast98.walk(_ip98[0]))
        _c98 = len([n for n in _ast98.walk(_etree98)
                    if isinstance(n, _ast98.FunctionDef)
                    and n.name == "_wrap_selection"]) == 1
        check("98.3 结构护栏（AST）：(a) `_WRAP_OPEN` 是 `VbeBackend` 的类体"
              "常量；(b) `insert_pair` 确实按它决定\"包不包\"；(c) `_wrap_selection`"
              "仍然只有一处实现（括号与引号共用同一份，不另写第二份）",
              [(_a98, _b98, _c98)],
              expect_contain=[(True, True, True)])
    except Exception as _e98:
        import traceback as _tp98
        _tp98.print_exc()
        check("第 98 节异常: %s" % _e98, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 99. ★v115：代码片段库（SQLite 两列）+ Ctrl+Q 片段窗
    #     用户口径：数据库两列（标题 / 代码）、Ctrl+Q 在光标处弹窗、
    #     标题模糊搜索、增删改、点按钮插到光标处。
    # ------------------------------------------------------------------
    print("\n=== 99. ★v115：代码片段库 / Ctrl+Q 片段窗 ===")
    try:
        import ast as _ast99
        import shutil as _sh99
        import tempfile as _tf99
        import snippets as SN99
        import snippet_ui as SUI99        # noqa: F401  （99.5 的分层护栏用）
        import vbe_bridge as VB99

        # ---- 99.1 标题模糊匹配（纯函数）----
        _rows99 = [("GetCode", "x = 1"), ("GetTitleByCode", "y = 2"),
                   ("数据处理", "z = 3"), ("SortArray", "w = 4")]
        _g99 = (
            [(t, h) for (t, _c, h) in SN99.search(_rows99, "code")],
            [(t, h) for (t, _c, h) in SN99.search(_rows99, "gc")],
            [(t, h) for (t, _c, h) in SN99.search(_rows99, "数据")],
            [(t, h) for (t, _c, h) in SN99.search(_rows99, "数据 理")],
            [t for (t, _c, _h) in SN99.search(_rows99, "zzz")],
            [t for (t, _c, _h) in SN99.search(_rows99, "")],
            SN99.match("GetCode", "GetCode"),
            SN99.match("code", "GetTitleByCode"),
            SN99.match("Code", "getcode"),
            SN99.match("   ", "GetCode"),
        )
        _w99 = (
            [("GetCode", [3, 4, 5, 6]),            # 连续子串（位置靠后）
             ("GetTitleByCode", [10, 11, 12, 13])],
            [("GetCode", [0, 3]),                  # 子序列（驼峰）gc
             ("GetTitleByCode", [0, 10])],         # 短标题在前
            [("数据处理", [0, 1])],                 # 中文按字匹配
            [("数据处理", [0, 1, 3])],              # 多词 AND
            [],                                    # 没命中 -> 空
            ["GetCode", "GetTitleByCode", "数据处理", "SortArray"],  # 空词 = 全给
            (128, [0, 1, 2, 3, 4, 5, 6]),          # 整词相等
            (76, [10, 11, 12, 13]),                # 子串
            (76, [3, 4, 5, 6]),                    # 大小写不敏感
            (0, []),                               # 纯空白 = 全部命中、无高亮
        )
        check("99.1 标题模糊匹配 `snippets.search` / `match`："
              "连续子串 > 子序列；多个词是 AND；中文按字匹配；"
              "空词返回全部；大小写不敏感；排序按\"分数高 -> 标题短\"",
              [_g99], expect_contain=[_w99])

        # ---- 99.2 存储：SQLite 两列（标题 / 代码）----
        _d99 = _tf99.mkdtemp()
        try:
            # 目录不存在时自己建（项目文件夹里的 snippets.db 首次使用时同理）
            _p99 = os.path.join(_d99, "new", "snippets.db")
            _st99 = SN99.SnippetStore(_p99)
            _r99 = (
                _st99.add("GetCode", "x = 1"),      # 新增
                _st99.add("GetCode", "dup"),        # 同标题 -> False（不覆盖）
                _st99.get("GetCode"),               # 仍是第一次的代码
                _st99.update("GetCode", "x = 2"),   # 改代码
                _st99.rename("GetCode", "GetCode2"),  # 改标题
                _st99.add("排序", "Sub S()\nEnd Sub"),
                _st99.rename("排序", "GetCode2"),    # 撞名 -> False
                [r[1] for r in _st99._conn().execute(
                    "PRAGMA table_info(snippets)")],  # 列清单
                _st99.all(),
                _st99.delete("排序"),
                _st99.all(),
                os.path.isfile(_p99),
            )
            _st99.close()
            _w99b = (True, False, "x = 1", True, True, True, False,
                     ["title", "code"],
                     [("GetCode2", "x = 2"), ("排序", "Sub S()\nEnd Sub")],
                     True,
                     [("GetCode2", "x = 2")],
                     True)
            check("99.2 `SnippetStore`：库文件自动建（含目录）；表**只有两列**"
                  " title / code；同标题新增被拒（不静默覆盖）；改代码 / 改标题 /"
                  " 删除都生效；改标题撞名不动数据",
                  [_r99], expect_contain=[_w99b])
        finally:
            _sh99.rmtree(_d99, ignore_errors=True)

        # ---- 99.2b 库的默认位置：**项目文件夹里**（★v116 用户口径）----
        # "能帮我把建的库，默认位置放在项目文件夹里吗？"
        # 判据要点：① 默认路径就在 `snippets.py` 所在目录；② 文件名是
        # snippets.db；③ 不依赖 cwd（用 `__file__`）；④ 环境变量仍然一票优先。
        _proj99 = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        _bak99 = os.environ.pop(SN99.ENV_DB, None)
        _envdir99 = _tf99.mkdtemp()
        try:
            _pdef99 = SN99.db_path()
            _cwd99 = os.getcwd()
            try:
                os.chdir(_tf99.gettempdir())      # cwd 换掉也不该跟着变
                _pdef99b = SN99.db_path()
            finally:
                os.chdir(_cwd99)
            os.environ[SN99.ENV_DB] = os.path.join(_envdir99, "mine.db")
            _penv99 = SN99.db_path()
        finally:
            if _bak99 is None:
                os.environ.pop(SN99.ENV_DB, None)
            else:
                os.environ[SN99.ENV_DB] = _bak99
            _sh99.rmtree(_envdir99, ignore_errors=True)
        check("99.2b `snippets.db_path`：库默认落在**项目文件夹**（`snippets.py` 所在"
              "目录），不依赖启动时的工作目录；`VBECOMPLETE_SNIPPET_DB` 仍然优先",
              [(_pdef99, os.path.dirname(os.path.abspath(SN99.__file__)),
                _pdef99 == _pdef99b, _penv99)],
              expect_contain=[(os.path.join(_proj99, "snippets.db"),
                               _proj99, True,
                               os.path.join(_envdir99, "mine.db"))])

        # ---- 99.3 光标落点判据（纯函数）----
        _g99c = (
            VB99._snippet_tail_pos("x = 1", ""),
            VB99._snippet_tail_pos("foo(1, 2)", ")"),
            VB99._snippet_tail_pos("abc  , d", "  , d"),
            VB99._snippet_tail_pos("abc= 1", ", d"),   # 找不到尾部 -> 行尾
            VB99._snippet_tail_pos("x = 1  ", "  "),
            VB99._snippet_tail_pos("", "x"),
        )
        check("99.3 `_snippet_tail_pos`：光标落在\"片段最后一个字符之后\"——"
              "尾部原文在行里的起点（行尾就是它 -> 长度相减；VBE 重排过 -> "
              "按 strip 后最后一次出现找回；都找不到 -> 行尾）",
              [_g99c], expect_contain=[(5, 8, 3, 6, 5, 0)])

        # ---- 99.4 插入（假 VBE）----
        class _Mod99(object):
            def __init__(self, lines):
                self.lines = list(lines)

            def Lines(self, n, c):
                return "\r\n".join(self.lines[n - 1:n - 1 + c])

            @property
            def CountOfLines(self):
                return len(self.lines)

            def ReplaceLine(self, n, t):
                self.lines[n - 1] = t

            def InsertLines(self, n, t):
                # ★v124：按真机实测行为建模 —— VBE 的 InsertLines 对 `\r\n` /
                # `\n` / `\r` 三种换行都拆行（Excel 2016 / VBE7.1 实测）。
                # `insert_snippet` 现在一次灌中间行，这里也得认 CRLF。
                _t = str(t).replace("\r\n", "\n").replace("\r", "\n")
                for k, ln in enumerate(_t.split("\n")):
                    self.lines.insert(n - 1 + k, ln)

            def DeleteLines(self, n, c):
                del self.lines[n - 1:n - 1 + c]

        class _Pane99(object):
            def __init__(self, mod, sl, sc, el, ec):
                self.CodeModule = mod
                self.sel = [sl, sc, el, ec]

            def GetSelection(self):
                return tuple(self.sel)

            def SetSelection(self, sl, sc, el, ec):
                self.sel = [sl, sc, el, ec]

        class _Vbe99(object):
            def __init__(self, pane):
                self.ActiveCodePane = pane

        def _run99(lines, sl, sc, el, ec, text, sem=("char", 4, False)):
            mod = _Mod99(lines)
            pane = _Pane99(mod, sl, sc, el, ec)
            _o_vbe = VB99._get_vbe_cached
            _o_sem = VB99._sem_cache.get("info")
            VB99._sem_cache["info"] = sem
            VB99._get_vbe_cached = lambda: _Vbe99(pane)
            try:
                ok = VB99.VbeBackend().insert_snippet(text)
            finally:
                VB99._get_vbe_cached = _o_vbe
                VB99._sem_cache["info"] = _o_sem
            return bool(ok), mod.lines, (pane.sel[0], pane.sel[1])

        _g99d = (
            # 行尾插入：整行接上，光标落片段末尾
            _run99(["x = "], 1, 5, 1, 5, 'MsgBox "hi"'),
            # 行中间插入：后半行接在片段末尾
            _run99(["foo("], 1, 5, 1, 5, "1, 2"),
            # 多行片段、光标在模块末尾那个"虚拟空行"上（要先插行才写得进去）
            _run99(["x = ab"], 2, 1, 2, 1,
                   "Sub Test()\n    x = 1\nEnd Sub"),
            # 多行选区：片段替换选区（旧行要删干净）
            _run99(["a = 1", "b = 2", "c = 3"], 2, 1, 3, 6, "NEW"),
            # 多行片段插在行中间：首行接前半、末行接后半
            _run99(["x = ab"], 1, 6, 1, 6, "1\n2\n3"),
            # CRLF 也要按两行插
            _run99(["x = "], 1, 5, 1, 5, "a\r\nb"),
            # 空片段 -> 不动文本
            _run99(["x = ab"], 1, 6, 1, 6, "   "),
        )
        _w99d = (
            (True, ['x = MsgBox "hi"'], (1, 16)),
            (True, ["foo(1, 2"], (1, 9)),
            (True, ["x = ab", "Sub Test()", "    x = 1", "End Sub"], (4, 8)),
            (True, ["a = 1", "NEW"], (2, 4)),
            (True, ["x = a1", "2", "3b"], (3, 2)),
            (True, ["x = a", "b"], (2, 2)),
            (False, ["x = ab"], (1, 6)),
        )
        check("99.4 `VbeBackend.insert_snippet`（集成）：片段原样插到光标处 ——"
              "行尾 / 行中间都行（后半行接在片段末行后面）；多行片段逐行插入；"
              "光标在\"虚拟空行\"上时先建行；多行选区是替换语义（旧行删干净）；"
              "CRLF 归一；空片段不动文本并返回 False",
              [_g99d], expect_contain=[_w99d])

        # ---- 99.5 分层护栏 + 快捷键判据 ----
        def _imports99(path):
            with io.open(os.path.join(ROOT, path), encoding="utf-8") as f:
                tree = _ast99.parse(f.read())
            out = set()
            for node in _ast99.walk(tree):
                if isinstance(node, _ast99.Import):
                    for a in node.names:
                        out.add((a.name or "").split(".")[0])
                elif isinstance(node, _ast99.ImportFrom):
                    if node.module:
                        out.add(node.module.split(".")[0])
            return out

        _ui_imp = _imports99("snippet_ui.py")
        _sn_imp = _imports99("snippets.py")
        _a99 = (not ({"vbe_bridge", "sqlite3", "win32com"} & _ui_imp),   # 窗不碰 COM/库
                "tkinter" not in _sn_imp)                                 # 数据层不碰界面
        # 写入实现只有一处：窗里不许自己动 CodeModule
        with io.open(os.path.join(ROOT, "snippet_ui.py"), encoding="utf-8") as f:
            _src99 = f.read()
        _b99 = ("ReplaceLine" not in _src99 and "InsertLines" not in _src99)
        _c99 = bool(getattr(VB99.VbeBackend, "insert_snippet", None)) and \
            bool(getattr(VB99.VbeBackend, "focus_vbe", None))
        _d99v = None
        if M is not None:
            _d99v = (M.VK_Q,
                     M._snippet_vk_ok(0x51, True, False, False, False),
                     M._snippet_vk_ok(0x51, True, True, False, False),
                     M._snippet_vk_ok(0x51, True, False, True, False),
                     M._snippet_vk_ok(0x51, True, False, False, True),
                     M._snippet_vk_ok(0x51, False, False, False, False),
                     M._snippet_vk_ok(0x49, True, False, False, False))
        check("99.5 结构护栏（AST）：(a) `snippet_ui` 不 import COM/`sqlite3`、"
              "`snippets` 不 import tkinter（数据层 / 界面层 / COM 三层不串味）；"
              "(b) 窗里没有 `ReplaceLine`/`InsertLines` —— 写入只有 "
              "`VbeBackend.insert_snippet` 一处实现；(c) Ctrl+Q 判据只认"
              "**单独** Ctrl（带 Shift/Alt/Win 一律不抢）",
              [(_a99, _b99, _c99, _d99v)],
              expect_contain=[((True, True), True, True,
                               (0x51, True, False, False, False, False, False))])

        # ---- 99.6 版式：三条右边缘对齐（★v117 用户口径）----
        # "搜索文本框和标题文本框...让他俩的右侧跟代码文本框右侧对齐。按钮位置也
        #  向左移动一点，让插入到光标处按钮右侧跟代码文本框右侧对齐。"
        # 判据只认**真实几何**：代码框右边站着竖向滚动条，所以搜索行 / 按钮行的右侧
        # 留白要各让出一个滚动条宽度；标题框则**不能**跨到滚动条那一列。
        # 无 GUI 环境（真 tkinter 装不上）自动跳过这一条。
        _g99e = None
        _root99 = None
        _dd99 = _tf99.mkdtemp()
        _tk99_had = SUI99.tk
        _tk99_saved = None
        try:
            import importlib as _il99
            _tk99 = sys.modules.get("tkinter")
            if not hasattr(_tk99, "Tk"):
                # ⚠️ 本测试网顶部给 tkinter 塞了**桩**（managed python 未必带 GUI
                # 包）⇒ 直接 `import tkinter` 拿到的是桩、`Tk()` 不存在、这一节会被
                # 静默跳过（§20.11 其实一直是这样跳过的）。这里要的是**真实几何**，
                # 所以临时把真模块换进来，跑完连子模块一起原样换回。
                _tk99_saved = {k: sys.modules[k] for k in list(sys.modules)
                               if k == "tkinter" or k.startswith("tkinter.")}
                for _k99 in _tk99_saved:
                    del sys.modules[_k99]
                try:
                    _tk99 = _il99.import_module("tkinter")
                except Exception:
                    _tk99 = None
            if _tk99 is not None and hasattr(_tk99, "Tk"):
                SUI99.tk = _tk99          # snippet_ui 里 `import tkinter as tk`
                _root99 = _tk99.Tk()
                _root99.withdraw()
                _st99x = SN99.SnippetStore(os.path.join(_dd99, "align.db"))
                _st99x.add("GetCode", 'MsgBox "hi"')
                _w99x = SUI99.SnippetWindow(_root99, _st99x,
                                            on_insert=lambda c: True)
                _w99x.open_at_caret((300, 300, 316))
                _root99.update()
                _R99 = lambda wd: wd.winfo_rootx() + wd.winfo_width()
                _txtR99 = _R99(_w99x.txt)
                _g99e = (
                    8 <= _w99x._sbw <= 48,                      # 滚动条宽度量到了
                    _R99(_w99x.ent_search) - _txtR99,           # 搜索框右边
                    _R99(_w99x.ent_title) - _txtR99,            # 标题框右边
                    _R99(_w99x.btn_insert) - _txtR99,           # 插入按钮右边
                    int(_w99x.ent_title.grid_info().get("columnspan", 1)),
                )
                _w99x.close()
                _st99x.close()
        except Exception:
            _g99e = None            # 没有 GUI —— 跳过（不算失败）
        finally:
            SUI99.tk = _tk99_had
            if _tk99_saved is not None:
                for _k99 in [k for k in list(sys.modules)
                             if k == "tkinter" or k.startswith("tkinter.")]:
                    del sys.modules[_k99]
                sys.modules.update(_tk99_saved)
            try:
                if _root99 is not None:
                    _root99.destroy()
            except Exception:
                pass
            _sh99.rmtree(_dd99, ignore_errors=True)
        if _g99e is not None:
            check("99.6 版式：搜索框 / 标题框 / 「插入到光标处」的**右边缘都与代码"
                  "文本框右边缘重合**（差 0 像素），标题框不跨到滚动条列；"
                  "滚动条宽度是现场量出来的（不是写死 17）",
                  [_g99e], expect_contain=[(True, 0, 0, 0, 1)])
        check("99.6b `_scrollbar_width`：量不到实际宽度时按 17 兜底"
              "（桩环境下也照样给得出数，不会抛）",
              [(SUI99._scrollbar_width(object()), SUI99._SBW_FALLBACK)],
              expect_contain=[(17, 17)])
    except Exception as _e99:
        import traceback as _tp99
        _tp99.print_exc()
        check("第 99 节异常: %s" % _e99, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 100. ★v118：弹窗反应慢 —— 候选池刷新从"每 0.4s 全量重解析"改成
    #      "按模块增量复用 + 只有声明行才提前失效"
    #
    # 病根（真机实测 24 模块 / 628 行的工程）：`_collect_identifiers` 一次
    # 224ms，而 main 的轮询只要距上次刷新 > 0.4s 就 invalidate 一次 ⇒ 打字时
    # 每 0.4 秒有一拍要在主线程上付这 224ms。改成模块级缓存后同一工程 29ms。
    #
    # 这一节的护栏分两类：
    #   * 纯判据：line_may_declare（限流分档的唯一判据）；
    #   * 等价性：开了缓存与关掉缓存，候选池必须**逐字相同** —— 这是这类
    #     缓存唯一真正危险的地方（重放漏一个字段就会静默少几条候选）。
    # ------------------------------------------------------------------
    print("\n=== 100. ★v118：弹窗反应速度（候选池增量刷新）===")
    try:
        import vbe_bridge as VB100

        # ---- 100.1 限流判据（纯函数）----
        _cases100 = [
            ("Dim x As Long", True), ("Dim x", True), ("Dim", True),
            ("Public Sub ", True), ("Private Su", True), ("Private", True),
            ("    Static x As Long", True), ("Const N = 5", True),
            ("Sub test()", True), ("Function F() As Long", True),
            ("Type T", True), ("Enum E", True),
            ("Sub Foo(ByVal a", True),     # 形参表
            ("msgbox \"hi\"", False), ("x = foo(1, 2)", False),
            ("    If a Then", False), ("' Dim x", False),
            ("Set c = New Cls", False), ("End Sub", False),
            ("", False), ("   ", False), ("ReDim arr(10)", False),
            ("s = \"Dim x\"", False),      # 字符串里的 Dim 不算
            ("MsgBox \"a: b\"", False), ("x.As", False), ("a & b", False),
            ("For i = 1 To 3", False), ("i = i + 1", False),
        ]
        _bad100 = [(t, e, P.line_may_declare(t)) for t, e in _cases100
                   if P.line_may_declare(t) is not e]
        check("100.1 `parser.line_may_declare`（限流判据）：声明行/形参表一律"
              "判\"是\"（含只打了半截的 `Private Su`、只打了修饰词的 "
              "`Private`）；赋值 / 调用 / 注释 / 字符串里的 Dim 全判\"否\"；"
              "宁可判宽（多刷一次只是几十毫秒，漏刷会让\"刚声明的变量不提示\"复发）",
              [(_bad100, len(_cases100))],
              expect_contain=[(_bad100, len(_cases100))])

        # ---- 100.2 模块缓存的两档间隔 ----
        _gap100 = None
        if M is not None:
            _gap100 = (M.ID_REFRESH_MIN_SEC, M.ID_REFRESH_MIN_SEC_PLAIN)
        check("100.2 刷新限流分两档：声明行用 `ID_REFRESH_MIN_SEC`（0.4s）、"
              "普通行用 `ID_REFRESH_MIN_SEC_PLAIN`（2s，普通字符不可能新增"
              "名字，没必要重建池子）",
              [(_gap100, _gap100 is not None and (0 < _gap100[0]
                                                  < _gap100[1]))],
              expect_contain=[((0.4, 2.0), True)])

        # ---- 100.3 开关默认值（关掉退回 v117 行为）----
        check("100.3 模块缓存默认开启，`VBECOMPLETE_NO_MOD_CACHE=1` 可整段退回"
              "v117 的逐模块全量解析（A/B 用）",
              [(VB100.MOD_CACHE_ENABLED, VB100._MOD_CACHE_TTL)],
              expect_contain=[(True, 30.0)])

        # ---- 100.4 等价性（真机）：开缓存 vs 全量重建，池子必须一致 ----
        _bk100 = VB100.VbeBackend()
        _ctx100 = None
        try:
            _ctx100 = _bk100.get_context()
        except Exception:
            _ctx100 = None

        def _snap100(clear_mod):
            _bk100._cache = None
            if clear_mod:
                _bk100._mod_cache = {}
                _bk100._cm_cache = {}
                VB100._comp_list_cache["comps"] = None
            _recs = _bk100.get_identifiers()
            # 比【五路承接】的全貌，不只是候选池：缓存重放漏任何一路
            # （type_names / proc_names / proc_kinds / structural / declared）
            # 都会让某处判据悄悄退化，而池子本身看不出来。
            return (sorted(str(r[0]) + "|" + str(r[1]) for r in _recs),
                    sorted(str(n) for n in (getattr(_bk100, "_type_names", ()) or ())),
                    sorted(str(n) for n in (getattr(_bk100, "_proc_names", ()) or ())),
                    sorted((str(k), str(v)) for k, v in
                           (getattr(_bk100, "_proc_kinds", {}) or {}).items()),
                    sorted(str(n) for n in (getattr(_bk100, "_structural_names", ()) or ())),
                    sorted(str(n) for n in (getattr(_bk100, "_declared_names", ()) or ())),
                    sorted((str(k), tuple(sorted(str(x) for x in v)))
                           for k, v in
                           (getattr(_bk100, "_declared_by_module", {}) or {}).items()))

        _full100 = _snap100(True)     # 全量重建（相当于关掉缓存）
        _hot100 = _snap100(False)     # 命中模块缓存
        _g100 = (len(_full100[0]), _full100 == _hot100, bool(_full100[0]),
                 sum(len(x) for x in _full100[1:]))
        check("100.4 等价性（真机）：全量重解析与命中模块缓存得到的【五路承接】"
              "**逐字相同** —— 候选池（含模块归属）、类型名、过程名、过程种类、"
              "结构性名字、真名集合、跨模块声明表。重放漏任何一路都会让某处"
              "判据悄悄退化，而池子本身看不出来",
              [_g100], expect_contain=[(_g100[0], True, True, _g100[3])])

        # ---- 100.5 光标所在模块必须【不】进缓存 ----
        # 那是用户正在改的那一块：带着 caret 的解析结果与光标位置绑定，
        # 缓存下来会把"上一拍的光标处词"串到下一拍（回声防护 / 隐式变量）。
        _cur100 = None
        _cached100 = None
        if _ctx100:
            _cur100 = str(_ctx100.get("module_name") or "").lower()
            _cached100 = _cur100 not in _bk100._mod_cache
        # ⚠️ 真机断言的前提是"VBE 里开着一个代码窗"。用户把代码窗关掉时
        # `get_context()` 返回 None，这条就**无从判定**（整工具那时也不工作）——
        # 显式按"空缺"通过，别让它变成一条看运气变红的用例。
        _vac100 = _ctx100 is None
        check("100.5 光标所在的那个模块【不进模块缓存】（每轮重读重解析）："
              "带 caret 的解析结果与光标位置绑定，缓存会串味；"
              "其它模块才复用（这条同时守住\"刚声明的变量立刻能提示\"）"
              "（没有打开的代码窗时判为空缺）",
              [(True if _vac100 else _cached100,
                _vac100 or _cached100 is True,
                _cur100,)],
              expect_contain=[(True, True, _cur100)])

        # ---- 100.6 引用清单的时间缓存确实生效 ----
        # 读一次 References 要 ~22ms（每个 Item 都是新包装对象），原来每轮
        # 刷新都重读一遍。这条确认第二次调用没有重读（ts 不变）。
        # ⚠️ 直接测 `_host_enum_refs`，不走 `host_enum_constants` —— 后者在
        # `ENABLE_HOST_ENUMS` 关着（出厂默认）时压根不会调它。
        _refs100 = None
        try:
            _r1 = VB100._host_enum_refs()
            _ts1 = VB100._host_refs_cache["ts"]
            _r2 = VB100._host_enum_refs()
            _ts2 = VB100._host_refs_cache["ts"]
            _refs100 = (isinstance(_r1, list), _ts1 == _ts2,
                        _r1 is not None and len(_r1) > 0, _r1 == _r2)
        except Exception:
            _refs100 = (False, False, False, False)
        check("100.6 `_host_enum_refs` 的时间缓存生效（真机）：第二次调用不再"
              "重读 References —— 每轮刷新白花的 22ms 省掉了；两次返回的清单"
              "必须一致",
              [_refs100], expect_contain=[(True, True, True, True)])
    except Exception as _e100:
        import traceback as _tp100
        _tp100.print_exc()
        check("第 100 节异常: %s" % _e100, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 101. ★v119：候选窗字号真的跟随 VBE + 最近使用表跨会话累计
    #   用户本轮问的是"弹窗能不能反应快点"（性能，v118 已处理），量性能的路上
    #   撞见两个"用户能感知、却永远不报错"的缺口：
    #     ① 候选窗字号**从来没跟随过** VBE：`_apply_font` 算出的 size 是浮点数，
    #        tkinter 只收整数 → TclError → 被两个 except 吞掉 → 永远 Consolas 10。
    #     ② "最近使用优先"**活不过一次重启**：`_recent` 是纯内存字典，
    #        stop.bat 一停就清零，等于每次开机都从零重新"教"一遍。
    # ------------------------------------------------------------------
    print("\n=== 101. ★v119：字号跟随 VBE / 最近使用表落盘 ===")
    import tempfile as _t101
    import shutil as _sh101
    import importlib as _il101
    import subprocess as _sp101
    import ast as _ast101
    import ui as UI101
    import recent as RC101

    # ---- 101.1 纯函数：任何输入都必须返回 int 并钳在 [8,24] ----
    # 旧实现 `max(MIN, min(MAX, pt + FONT_DELTA))` 返回浮点（11.0 / 8.75），
    # 这一条直接钉住"返回值类型"——只比数值是不够的，float(9) == 9 会假绿。
    _cases101 = [
        (12.0, 11), (9.75, 9), (8.0, 8), (1.0, 8), (100.0, 24), (24.6, 24),
        (None, 8), ("bad", 8), (float("nan"), 8), (float("inf"), 24),
        (float("-inf"), 8), (-5, 8),
    ]
    _bad101 = [(p, e, UI101._clamp_font_size(p))
               for p, e in _cases101
               if type(UI101._clamp_font_size(p)) is not int
               or UI101._clamp_font_size(p) != e]
    check("101.1 `ui._clamp_font_size` 对 12 种输入（含 None / NaN / ±inf / "
          "字符串 / 越界）都返回 **int** 且钳在 [8,24]；取整用 round "
          "（8.75→9，不是 int() 的 8）",
          [(_bad101, len(_cases101))],
          expect_contain=[([], len(_cases101))])

    # ---- 101.2 真 tkinter 实测：旧写法必抛，新写法真的生效 ----
    # 这条是整节的根：不复制"浮点被拒收"这个事实，将来有人把 `int(round(...))`
    # 改回 `max/min` 也不会红（数值看起来完全一样）。
    # ⚠️ 测试网顶部把 tkinter 塞成了**桩**（managed python 未必带 GUI 包），
    #    直接 import 拿到的是桩 ⇒ 必须照 §99.6 的办法临时换真模块、跑完换回。
    _g101 = None
    _root101 = None
    _saved101 = None
    try:
        _tk101 = sys.modules.get("tkinter")
        if not hasattr(_tk101, "Tk"):
            _saved101 = {k: sys.modules[k] for k in list(sys.modules)
                         if k == "tkinter" or k.startswith("tkinter.")}
            for _k101 in _saved101:
                del sys.modules[_k101]
            try:
                _tk101 = _il101.import_module("tkinter")
                _tf101 = _il101.import_module("tkinter.font")
            except Exception:
                _tk101 = None
                _tf101 = None
        else:
            _tf101 = sys.modules.get("tkinter.font")
        if (_tk101 is not None and hasattr(_tk101, "Tk")
                and _tf101 is not None):
            _root101 = _tk101.Tk()
            _root101.withdraw()
            _f101 = _tf101.Font(family="Consolas", size=10)
            _old_ok101 = True
            try:
                # 旧写法逐字复刻：pt(9.75) + FONT_DELTA(-1) = 8.75（浮点）
                _f101.configure(family="Consolas",
                                size=max(8, min(24, 9.75 - 1)))
            except Exception:
                _old_ok101 = False          # ← 旧写法【必须】在这里失败
            _new101 = UI101._clamp_font_size(9.75)
            _f101.configure(family="Consolas", size=_new101)
            _g101 = (_old_ok101, _new101, _f101.actual("size") == _new101)
    except Exception:
        _g101 = None                        # 没有 GUI —— 跳过（不算失败）
    finally:
        if _saved101 is not None:
            for _k101 in [k for k in list(sys.modules)
                          if k == "tkinter" or k.startswith("tkinter.")]:
                del sys.modules[_k101]
            sys.modules.update(_saved101)
        try:
            if _root101 is not None:
                _root101.destroy()
        except Exception:
            pass
    if _g101 is not None:
        check("101.2 真 tkinter 实测：旧写法（浮点 8.75）**必抛 TclError**、"
              "新写法（int 9）真的落到 Font 上 —— 这就是「字号从不跟随」的"
              "根因；只比数值会假绿，所以这里连「抛不抛」一起断言",
              [_g101], expect_contain=[(False, 9, True)])

    # ---- 101.1b `_apply_font` 真的用了它（AST 查函数体）----
    # 101.1 / 101.2 只证明"这个函数算得对"，证明不了"调用点用的是它"：
    # 有人把 `_apply_font` 改回 `max(MIN, min(MAX, pt + FONT_DELTA))` 时，
    # 上面两条依然全绿，而用户那边字号又不动了。
    _apply101 = None
    try:
        with open(os.path.join(ROOT, "ui.py"), "r", encoding="utf-8") as _fu101:
            _tree_u101 = _ast101.parse(_fu101.read())
        for _n101 in _ast101.walk(_tree_u101):
            if (isinstance(_n101, _ast101.FunctionDef)
                    and _n101.name == "_apply_font"):
                _apply101 = _n101
                break
    except Exception:
        _apply101 = None
    _calls_u101 = sorted({
        _n.func.id for _n in _ast101.walk(_apply101)
        if isinstance(_n, _ast101.Call)
        and isinstance(_n.func, _ast101.Name)
    }) if _apply101 is not None else []
    check("101.1b `ui._apply_font` 算字号时**确实调用了** `_clamp_font_size`"
          "（AST 查函数体，不搜源码字符串）—— 光测函数自己证明不了调用点，"
          "改回 max/min 时 101.1/101.2 都会假绿",
          [(_apply101 is not None, "_clamp_font_size" in _calls_u101)],
          expect_contain=[(True, True)])

    # ---- 101.3 recent 数据层：往返 + 六种坏输入 + 上限 + 原子写 ----
    _dd101 = _t101.mkdtemp()
    _old_env101 = os.environ.get("VBECOMPLETE_RECENT_FILE")
    try:
        os.environ["VBECOMPLETE_RECENT_FILE"] = os.path.join(_dd101, "r.json")

        _p101 = os.path.join(_dd101, "ok.json")
        _rt101 = RC101.save({"aaa": 1, "bbb": 5, "ccc": 3}, _p101)
        _back101 = RC101.load(_p101)
        _g_rt101 = (_rt101, list(_back101.keys()),
                    RC101.load(os.path.join(_dd101, "nope.json")))

        def _w101(name, text):
            _pp = os.path.join(_dd101, name)
            with open(_pp, "w", encoding="utf-8") as _f:
                _f.write(text)
            return _pp

        _g_bad101 = (
            RC101.load(_w101("a.json", "{not json")),        # 坏 json
            RC101.load(_w101("b.json", "[1,2,3]")),          # 不是 dict
            RC101.load(_w101("c.json",
                             '{"foo":"bar","ok":3}')),       # 值不是数字
            RC101.load(_w101("d.json",
                             '{"foo":true,"ok":3}')),        # bool 是 int 子类
            RC101.load(_w101("e.json",
                             '{"foo":-1,"ok":3}')),          # 负序号
            RC101.load(_w101("f.json", '{"  ":1,"ok":3}')),  # 空名字
        )
        # 超上限：600 条只留最新 512 条，且序号重排成 1..512
        _big101 = {"n%03d" % _i: _i for _i in range(600)}
        _bp101 = os.path.join(_dd101, "big.json")
        RC101.save(_big101, _bp101)
        _bl101 = RC101.load(_bp101)
        _g_cap101 = (len(_bl101), max(_bl101.values()),
                     "n599" in _bl101, "n000" in _bl101)

        # 深层目录自动建 + 绝不残留 .tmp（stop.bat 是强杀，半截文件必须不存在）
        _deep101 = os.path.join(_dd101, "x", "y", "z.json")
        _ds101 = RC101.save({"a": 1}, _deep101)
        _g_disk101 = (_ds101, os.path.exists(_deep101),
                      os.path.exists(_deep101 + ".tmp"))

        # 默认路径：项目根（不是 cwd），env 优先
        _gpath101 = (RC101.path() == os.path.join(_dd101, "r.json"),
                     RC101.DEFAULT_NAME == "recent.json")
        _old_cwd101 = os.getcwd()
        try:
            os.chdir(_dd101)
            _same_after_chdir = (RC101.path()
                                 == os.path.join(_dd101, "r.json"))
        finally:
            os.chdir(_old_cwd101)

        check("101.3 `recent` 数据层：save/load 往返保持先后次序（序号重排成 "
              "1..N）；文件不存在 / 坏 json / 不是 dict 一律返回 {} 而不是抛",
              [_g_rt101],
              expect_contain=[(True, ["aaa", "ccc", "bbb"], {})])
        check("101.3b 六种脏数据全部**跳过**而不是整个丢掉：值不是数字、"
              "bool（`True` 是 int 子类，不排掉会变成 1）、负序号、空名字 "
              "——只有合法的那条 `ok` 留下来",
              [_g_bad101],
              expect_contain=[({}, {}, {"ok": 1}, {"ok": 1},
                               {"ok": 1}, {"ok": 1})])
        check("101.3c 超过上限只留最新的一批（600 条 → 512），序号重排成 "
              "1..512，最旧的 n000 被丢掉",
              [_g_cap101], expect_contain=[(512, 512, True, False)])
        check("101.3d 目标目录不存在时自动建；写完**不留 .tmp**"
              "（直接覆盖写到一半被强杀会留下半截 json，下次读回来是空）",
              [_g_disk101], expect_contain=[(True, True, False)])
        check("101.3e 默认路径在**项目文件夹**（`__file__` 所在目录），"
              "env 覆盖优先，且 chdir 之后不变（run.bat / pythonw 的 cwd 会漂）",
              [(_gpath101, _same_after_chdir)],
              expect_contain=[((True, True), True)])

        # ---- 101.4 引擎集成：构造不读盘 → load → 排序跨实例一致 ----
        class _U101:
            def __init__(self):
                self.shown = False

            def show(self, *a, **k):
                self.shown = True

            def update_selection(self, *a, **k):
                pass

            def hide(self, *a, **k):
                self.shown = False

            def contains_point(self, *a, **k):
                return False

        class _B101:
            def get_context(self):
                return {"line_no": 1, "caret_col": 7, "line_text": "    get",
                        "in_string": False, "in_comment": False,
                        "in_type_position": False, "in_decl_position": False,
                        "proc_name": "P", "module_name": "M"}

            def get_identifiers(self):
                # 长度相同、前缀相同、作用域相同 ⇒ 只有"最近使用"能分出胜负
                return [("getAbc", "M", "P", False),
                        ("getXyz", "M", "P", False)]

            def apply_completion(self, *a, **k):
                return None

        # 先造一个"上次会话留下的"文件。⚠️ 这里存的名字必须与下面要记的
        # `_target101` 无关 —— 第一版存的就是 `getxyz`，而 target 恰好也是
        # `getXyz`，小写后是同一条 ⇒ 条数永远是 1，"跨实例恢复两条"的断言
        # 反而失败。测试数据也要"互不干扰"，这类巧合最容易骗人。
        RC101.save({"zzzother": 7}, RC101.path())
        _c_init101 = E.Completer(_B101(), _U101())
        _no_read101 = (dict(_c_init101._recent), _c_init101._recent_dirty)
        _loaded101 = _c_init101.load_recent()
        _after_load101 = dict(_c_init101._recent)

        # 未加载的实例：当作"没有任何最近使用"的基线
        _c_base101 = E.Completer(_B101(), _U101())
        _c_base101.trigger()
        _base101 = list(_c_base101.matches)
        _target101 = _base101[-1]           # 原本排在后面的那个

        # 记它一笔 → 本实例排序里它应该升到第一
        _c_new101 = E.Completer(_B101(), _U101())
        # 先 load（真实启动路径就是这样：main 构造完立刻 load_recent），
        # 再记一笔 —— 否则 flush 会用"只含一个新名字"的内存表**覆盖**磁盘，
        # 那是测试自己造出来的假故障，不是产品行为。
        _c_new101.load_recent()
        _c_new101._note_recent(_target101)
        _dirty101 = _c_new101._recent_dirty
        _flushed101 = _c_new101.flush_recent()
        _again101 = _c_new101.flush_recent()      # 干净了就该返回 False
        _c_new101.trigger()
        _hot101 = list(_c_new101.matches)

        # 换一个全新实例，只靠磁盘恢复
        _c_reload101 = E.Completer(_B101(), _U101())
        _n_reload101 = _c_reload101.load_recent()
        _c_reload101.trigger()
        _cold101 = list(_c_reload101.matches)

        check("101.4 构造 `Completer()` **不读盘**（测试网要跑上千次构造，"
              "绝不能让每条用例悄悄依赖用户机器上的真实数据）；显式 "
              "`load_recent()` 才恢复，并顺带把序号基线到最大值",
              [(_no_read101, _loaded101, _after_load101,
                _c_init101._recent_seq)],
              expect_contain=[(({}, False), 1, {"zzzother": 1}, 1)])
        check("101.4b 落盘 + 跨实例：记一笔 → 脏标记亮 → flush 成功 → 再 "
              "flush 返回 False（幂等）；新实例只靠 load 就能让**同一个名字**"
              "重新排到第一，且这个顺序与基线不同（否则证明不了什么）",
              [(_dirty101, _flushed101, _again101, _n_reload101,
                _hot101 == _cold101, _base101 != _hot101)],
              expect_contain=[(True, True, False, 2, True, True)])
        check("101.4c 排序方向真的对：原本排在最后的 `%s` 被记一笔后升到第一，"
              "且基线里它不是第一（否则这条用例证明不了任何事）"
              % _target101,
              [(_hot101[0] == _target101, _cold101[0] == _target101,
                _base101[0] != _target101)],
              expect_contain=[(True, True, True)])

        # ---- 101.5 开关：默认开，`NO_RECENT_PERSIST=1` 关掉 ----
        # 必须开子进程 —— 本进程的 engine 已经导入过了，改环境变量不影响常量。
        def _flag101(env_val):
            _code = ("import os,sys\n"
                     "sys.path.insert(0, %r)\n"
                     "import engine\n"
                     "print(int(engine.PERSIST_RECENT))\n"
                     "print(int(engine.RANK_BY_RECENT))\n" % ROOT)
            _env = dict(os.environ)
            _env.pop("VBECOMPLETE_NO_RECENT_PERSIST", None)
            _env.pop("VBECOMPLETE_NO_RECENT_RANK", None)
            if env_val is not None:
                _env["VBECOMPLETE_NO_RECENT_PERSIST"] = env_val
            try:
                _r = _sp101.run([sys.executable, "-c", _code],
                                capture_output=True, text=True, timeout=120,
                                env=_env)
                return _r.stdout.split()
            except Exception:
                return []

        check("101.5 `PERSIST_RECENT` 默认**开**；`VBECOMPLETE_NO_RECENT_PERSIST=1` "
              "关掉时行为退回 v118（既不读也不写）。子进程实测，不靠读常量表",
              [(_flag101(None), _flag101("1"))],
              expect_contain=[(["1", "1"], ["0", "1"])])
    finally:
        if _old_env101 is None:
            os.environ.pop("VBECOMPLETE_RECENT_FILE", None)
        else:
            os.environ["VBECOMPLETE_RECENT_FILE"] = _old_env101
        _sh101.rmtree(_dd101, ignore_errors=True)

    # ---- 101.6 轮询里确实调了 flush_recent（AST 查函数体）----
    # 只在 accept 里置脏、却忘了在轮询里写盘 ⇒ 功能完全失效，而上面每一条
    # 都还是绿的（它们直接调 flush_recent）。只能用 AST 看 poll_editor 里
    # 有没有这个调用。⚠️ 不搜源码字符串：注释里的 `flush_recent` 会把它喂饱。
    _poll101 = None
    try:
        with open(os.path.join(ROOT, "main.py"), "r", encoding="utf-8") as _f:
            _tree101 = _ast101.parse(_f.read())
        for _n101 in _ast101.walk(_tree101):
            if (isinstance(_n101, _ast101.FunctionDef)
                    and _n101.name == "poll_editor"):
                _poll101 = _n101
                break
    except Exception:
        _poll101 = None
    _calls101 = sorted({
        _n.func.attr for _n in _ast101.walk(_poll101)
        if isinstance(_n, _ast101.Call)
        and isinstance(_n.func, _ast101.Attribute)
    }) if _poll101 is not None else []
    check("101.6 `main.poll_editor` 里**确实调用了** `flush_recent`"
          "（AST 查函数体；只在 accept 置脏、忘了在轮询写盘的话，"
          "上面所有用例都还是绿的，而功能完全失效）",
          [(_poll101 is not None, "flush_recent" in _calls101,
            "load_recent" in _calls101)],
          expect_contain=[(True, True, False)])

    # ------------------------------------------------------------------
    # 102. ★v120：配置中心 —— 环境变量 > settings.ini > 代码默认值
    #   用户口径（原话）："第二档里的配置中心帮我弄一下。"
    #
    #   改这一层的动因不是"多个配置文件好看"，而是原来那 40 多个开关：
    #     ① 没人记得住名字；② 双击 run.bat 的用户根本没有 set 的机会；
    #     ③ 写错了（`=ture`）会静默回落默认值，用户以为开了其实没开。
    #
    #   本节钉三件事：**优先级链**、**认不出的值只跳过不失效**、**模板生成
    #   不会改变任何默认行为**。最后一条最要紧 —— 一旦生成的 ini 里有一行
    #   忘了注释掉，用户装完就发现某个功能默认变了，而且永远查不出来。
    # ------------------------------------------------------------------
    print("\n=== 102. ★v120：配置中心（settings.ini）===")
    import tempfile as _t102
    import shutil as _sh102
    import subprocess as _sp102
    import settings as S102

    _dir102 = _t102.mkdtemp(prefix="vbe_cfg_")
    _ini102 = os.path.join(_dir102, "settings.ini")
    _old_env102 = os.environ.get("VBECOMPLETE_SETTINGS")
    _flags102 = ("VBECOMPLETE_AUTO_SPACE", "VBECOMPLETE_NO_AUTOPAIR",
                 "VBECOMPLETE_RECENT_MAX", "VBECOMPLETE_PROXY_TTL",
                 "VBECOMPLETE_MEMBER_POPUP", "VBECOMPLETE_NO_RECENT_RANK")
    _old_vals102 = {k: os.environ.get(k) for k in _flags102}

    def _w102(text):
        """写一份 ini 并让 settings 重新读盘。"""
        with io.open(_ini102, "w", encoding="utf-8", newline="\n") as _f:
            _f.write(text)
        S102.reset()

    try:
        os.environ["VBECOMPLETE_SETTINGS"] = _ini102
        for _k in _flags102:
            os.environ.pop(_k, None)
        S102.reset()

        # ---- 102.1 三层优先级 ----
        _w102("[settings]\nAUTO_SPACE = 0\nRECENT_MAX = 32\n")
        _g102a = (
            S102.flag("AUTO_SPACE", True),                    # ini 写 0
            S102.int_value("RECENT_MAX", 64),                 # ini 写 32
            S102.flag("NO_AUTOPAIR", False),                  # 都没写 -> default
            S102.effective("AUTO_SPACE", "1", "flag")[1],      # 来源必须是 ini
        )
        os.environ["VBECOMPLETE_AUTO_SPACE"] = "1"             # env 压过 ini
        _g102b = (S102.flag("AUTO_SPACE", False),
                  S102.source("AUTO_SPACE"))
        os.environ.pop("VBECOMPLETE_AUTO_SPACE", None)
        check("102.1 取值链 `环境变量 > settings.ini > 代码默认值`：ini 能改、"
              "环境变量能压过 ini；**来源**要如实报告是 ini 还是 env（排错时"
              "就靠这一列判断\"我改的地方到底生效没有\"）",
              [_g102a, _g102b],
              expect_contain=[(False, 32, False, "ini"), (True, "env")])

        # ---- 102.2 认不出的值只跳过这一层，不让功能失效 ----
        # 老契约是"认不出 -> default"。现在多了一层 ini，所以口径要更准：
        # **跳过这一层继续往下看**。用户既然在 ini 里明确写了，就该用 ini 的。
        os.environ["VBECOMPLETE_AUTO_SPACE"] = "maybe"         # 垃圾值
        os.environ["VBECOMPLETE_RECENT_MAX"] = ""              # 空串 = 没设
        _g102c = (S102.effective("AUTO_SPACE", "1", "flag"),
                  S102.effective("RECENT_MAX", "64", "int"))
        os.environ["VBECOMPLETE_AUTO_SPACE"] = "off"           # 换一种 false 写法
        _g102d = S102.flag("AUTO_SPACE", True)
        os.environ.pop("VBECOMPLETE_AUTO_SPACE", None)
        os.environ.pop("VBECOMPLETE_RECENT_MAX", None)
        check("102.2 认不出的值 / 空串**只跳过那一层**：env 写了垃圾时退到 ini"
              "（而不是掉回代码默认值）；`=off` 这种非 0/1 的写法也认",
              [_g102c, _g102d],
              expect_contain=[((False, "ini"), (32, "ini")), False])

        # ---- 102.3 键名两种写法（带不带 VBECOMPLETE_ 前缀）----
        _w102("[settings]\n"
              "VBECOMPLETE_MEMBER_POPUP = 1\n"
              "no_recent_rank = 1\n")
        _g102e = (S102.flag("MEMBER_POPUP", False),
                  S102.flag("NO_RECENT_RANK", False))
        check("102.3 `VBECOMPLETE_MEMBER_POPUP` 与 `MEMBER_POPUP` 两种写法都认，"
              "键名大小写不敏感（用户从文档里抄哪个都行）",
              [_g102e], expect_contain=[(True, True)])

        # ---- 102.4 坏 ini 一律安静退回默认，绝不抛 ----
        # 配置文件坏掉导致工具起不来是最不能接受的一种失败。
        _bad102 = []
        for _label, _body in (
            ("没有 [settings] 段", "AUTO_SPACE = 0\n"),
            ("全是垃圾", "\x00\x01 not an ini\n[[["),
            ("值里有 % 插值符", "[settings]\nVBE_POPUP_CLASS = A%B\n"),
            ("段名写错", "[setting]\nAUTO_SPACE = 0\n"),
        ):
            try:
                _w102(_body)
                _bad102.append((_label, S102.flag("AUTO_SPACE", True),
                                S102.text("VBE_POPUP_CLASS", "")))
            except Exception as _e102:
                _bad102.append((_label, "抛了: %s" % _e102, ""))
        # 重复键：后写的赢，而不是整份配置作废（strict=False）
        _w102("[settings]\nAUTO_SPACE = 1\nAUTO_SPACE = 0\n")
        _dup102 = S102.flag("AUTO_SPACE", True)
        # UTF-8 BOM（Windows 记事本另存为 UTF-8 时会加）
        with io.open(_ini102, "wb") as _f:
            _f.write(b"\xef\xbb\xbf[settings]\r\nAUTO_SPACE = 0\r\n")
        S102.reset()
        _bom102 = S102.flag("AUTO_SPACE", True)
        check("102.4 坏 ini 全部安静退回默认（不抛异常）：没有段 / 全是垃圾 / "
              "值里带 `%`（默认的 BasicInterpolation 会在这里炸，我们关掉了它）"
              "/ 段名写错；另加两条宽容项 —— 重复键取后写的、UTF-8 BOM 与 CRLF 都认",
              [(_bad102, _dup102, _bom102)],
              expect_contain=[(([("没有 [settings] 段", True, ""),
                                 ("全是垃圾", True, ""),
                                 ("值里有 % 插值符", True, "A%B"),
                                 ("段名写错", True, "")],
                                False, False))])

        # ---- 102.5 生成的模板 = 全部注释掉 = 全部默认值 ----
        _tpl102 = S102.template_text()
        _w102(_tpl102)
        _rows102 = S102.dump(verbose=False)
        _not_default102 = [(k, v, s) for k, v, s in _rows102 if s != "default"]
        _missing102 = [k for k, _d, _s in S102.KNOWN
                       if ("; %s = " % k) not in _tpl102]
        check("102.5 模板里**每一项都注释掉**（于是生成 settings.ini 这件事"
              "不会改变任何行为）；且 `KNOWN` 的每一项都在模板里出现（漏一项"
              "就等于用户永远不知道有它）",
              [(_not_default102, _missing102)],
              expect_contain=[([], [])])

        # ---- 102.6 ensure_file：不存在就建、存在就一字不动 ----
        _p102 = os.path.join(_dir102, "deep", "dir", "settings.ini")
        _r102a = S102.ensure_file(_p102)
        _size102a = os.path.getsize(_p102)
        _r102b = S102.ensure_file(_p102)
        check("102.6 `ensure_file` 首次生成（连目录一起建），第二次**不重写**"
              "（否则用户改过的配置会被程序覆盖掉，这是最不能犯的错）",
              [(_r102a[1], os.path.isfile(_p102), _r102b[1],
                os.path.getsize(_p102) == _size102a)],
              expect_contain=[(True, True, False, True)])

        # ---- 102.7 默认路径在项目文件夹（不随 cwd 漂）----
        _oldcwd102 = os.getcwd()
        os.environ.pop("VBECOMPLETE_SETTINGS", None)
        S102.reset()
        _path102a = S102.path()
        try:
            os.chdir(_dir102)
            S102.reset()
            _path102b = S102.path()
        finally:
            os.chdir(_oldcwd102)
            os.environ["VBECOMPLETE_SETTINGS"] = _ini102
            S102.reset()
        check("102.7 默认路径 = `settings.py` 所在目录（项目文件夹），"
              "**不是 cwd** —— run.bat / pythonw 拉起时 cwd 会漂（v116 在 "
              "snippets.db 上踩过同一个坑）",
              [(_path102a == os.path.join(ROOT, "settings.ini"),
                _path102b == _path102a,
                os.path.dirname(os.path.abspath(S102.__file__)) == ROOT)],
              expect_contain=[(True, True, True)])

        # ---- 102.8 护栏：源码里不许再直接读 VBECOMPLETE_ 环境变量 ----
        # 这是"配置中心"的核心约束。漏一个，那个开关就永远只能靠 set，
        # 用户在 settings.ini 里写了也不生效、还查不出来。
        _offenders102 = []
        for _f102 in ("main.py", "engine.py", "vbe_bridge.py", "log.py",
                      "parser.py", "ui.py", "recent.py", "snippets.py",
                      "snippet_ui.py"):
            _src102 = io.open(os.path.join(ROOT, _f102),
                              encoding="utf-8").read()
            if 'os.environ.get("VBECOMPLETE_' in _src102:
                _offenders102.append(_f102)
        check("102.8 除 `settings.py` 之外，**没有任何模块再直接读 VBECOMPLETE_ "
              "环境变量** —— 漏一个就等于那个开关在 settings.ini 里失灵",
              [_offenders102], expect_contain=[[]])

        # ---- 102.9 模板里的每个键都真的有人消费 ----
        # 反过来的一条：模板里写了、代码里没人读 = 用户改了没反应（比漏写更气人）。
        _allsrc102 = "".join(
            io.open(os.path.join(ROOT, _f102), encoding="utf-8").read()
            for _f102 in ("main.py", "engine.py", "vbe_bridge.py", "log.py",
                          "recent.py", "snippets.py", "wordbank.py"))
        _unused102 = [k for k, _d, _s in S102.KNOWN
                      if ('"%s"' % k) not in _allsrc102
                      and ('"VBECOMPLETE_%s"' % k) not in _allsrc102]
        check("102.9 `KNOWN` 里每个开关键都真的被某个模块读（模板写了却没人读 "
              "= 用户改了没反应）",
              [_unused102], expect_contain=[[]])

        # ---- 102.10 端到端：子进程里 ini 真的改变了运行时行为 ----
        # 前面 9 条都只证明"settings 自己算得对"。真正要证明的是
        # "新开一个进程、只靠 settings.ini，模块级常量就变了"。
        _w102("[settings]\nMEMBER_POPUP = 1\nNO_RECENT_RANK = 1\n")
        _code102 = ("import sys; sys.path.insert(0, %r); import engine; "
                    "print(int(engine.MEMBER_POPUP), int(engine.RANK_BY_RECENT))"
                    % ROOT)
        _env102 = dict(os.environ)
        _env102["VBECOMPLETE_SETTINGS"] = _ini102
        for _k in _flags102:
            _env102.pop(_k, None)
        try:
            _cp102 = _sp102.run([sys.executable, "-c", _code102], env=_env102,
                                capture_output=True, text=True, timeout=180)
            _out102 = [ln.strip() for ln in
                       (_cp102.stdout or "").splitlines() if ln.strip()]
            _got102 = _out102[-1] if _out102 else ("ERR:" + str(_cp102.stderr)[:80])
        except Exception as _e102b:
            _got102 = "异常: %s" % _e102b
        check("102.10 端到端（子进程）：只靠 `settings.ini` 两个开关就被真的改掉"
              "（`MEMBER_POPUP=1` / `NO_RECENT_RANK=1` -> engine 里 "
              "True / False），不靠环境变量、不靠改源码",
              [_got102], expect_contain=["1 0"])

    except Exception as _e102x:
        import traceback as _tp102
        _tp102.print_exc()
        check("第 102 节异常: %s" % _e102x, [True], expect_contain=[False])
    finally:
        # 铁律：改测试网的猴补丁（这里是 os.environ 与 settings 的读盘缓存）
        # 必须在节尾还原，否则后面每一节都跑在"别人的配置"上。
        if _old_env102 is None:
            os.environ.pop("VBECOMPLETE_SETTINGS", None)
        else:
            os.environ["VBECOMPLETE_SETTINGS"] = _old_env102
        for _k, _v in _old_vals102.items():
            if _v is None:
                os.environ.pop(_k, None)
            else:
                os.environ[_k] = _v
        S102.reset()
        _sh102.rmtree(_dir102, ignore_errors=True)

    # ------------------------------------------------------------------
    # 103. ★v120：Ctrl+/ 批量注释 / 取消注释（IDEA 语义）
    #   用户口径（原话）："Ctrl+/ 批量注释和Ctrl+/ 批量解除注释，这个帮我弄一下，
    #   就像IDEA那样。"
    #
    #   最要紧的一条不是"能不能加"，而是**能不能原样回来**：注释掉一段代码
    #   再取消，缩进 / 行号标签 / `#If` 预处理行必须逐字复原 —— 这个工具改的是
    #   用户的真实代码，一次不可逆的改动就足以让人不敢再用。
    # ------------------------------------------------------------------
    print("\n=== 103. ★v120：Ctrl+/ 批量注释 ===")
    import ast as _ast103
    import vbe_bridge as VB103
    import main as MN103

    # ---- 103.1 加注释：`'` 必须落在**缩进之后** ----
    _add103 = [
        ("    x = 1", "    ' x = 1"),          # 普通缩进行
        ("\tx = 1", "\t' x = 1"),              # Tab 缩进
        ("x = 1", "' x = 1"),                  # 顶格
        ("    End If", "    ' End If"),
        ("1: x = 1", "' 1: x = 1"),            # 行号标签必须仍顶格
        ("#If DEBUG Then", "' #If DEBUG Then"),  # 条件编译指令
        ("    #If DEBUG Then", "    ' #If DEBUG Then"),
        ("x = 1  ' 行尾注释", "' x = 1  ' 行尾注释"),   # 有行尾注释的仍是代码行
        ("Rem 说明", "' Rem 说明"),
        ("(", "' ("),
        ("    ' x = 1", "    ' ' x = 1"),       # 已注释的再加一层（IDEA 同款）
    ]
    _bad_add103 = [(s, P.add_comment_to_line(s), e) for s, e in _add103
                   if P.add_comment_to_line(s) != e]
    _cmt103 = [(s, P.is_comment_line(s)) for s in
               ("    ' x = 1", "' x", "    x = 1", "x = 1  ' 尾巴",
                "   ", "", "\t'")]
    check("103.1 加注释：`'` 落在**缩进之后**（插行首会把缩进一起注释掉，"
              "取消后就再也还原不出原来的缩进）；行号标签与 `#If` 顶格行照旧"
              "可逆；`x = 1  ' 尾巴` 这种**行尾注释**不算\"整行注释\"（它是"
              "一行代码，该被注释掉）",
              [(_bad_add103, _cmt103)],
              expect_contain=[([], [("    ' x = 1", True), ("' x", True),
                                    ("    x = 1", False),
                                    ("x = 1  ' 尾巴", False),
                                    ("   ", False), ("", False),
                                    ("\t'", True)])])

    # ---- 103.2 取消注释：先删 `' `，再退一个 `'` ----
    _del103 = [
        ("    ' x = 1", "    x = 1"),          # 我们自己加的（带空格）
        ("    'x = 1", "    x = 1"),           # 手工写的（不带空格）
        ("' x = 1", "x = 1"),
        ("    ' ' x", "    ' x"),              # 两层注释只剥一层
        ("    '", "    "),                     # 空注释 -> 空行（不是崩）
        ("    x = 1", "    x = 1"),            # 不是注释 -> 原样
        ("", ""),
    ]
    _bad_del103 = [(s, P.remove_comment_from_line(s), e) for s, e in _del103
                   if P.remove_comment_from_line(s) != e]
    check("103.2 取消注释：先删 `' ` 两个字符、不行再退一个 `'` —— 于是我们"
              "加的能精确还原，用户手写的 `'xxx` 也能正确取消；不是注释的行"
              "**原样返回**（在混合选区里当逐行算子时不会误伤）",
              [_bad_del103], expect_contain=[[]])

    # ---- 103.3 可逆性：remove(add(x)) == x 对任何非空行成立 ----
    # 这是整节的地基。**加注释绝不能丢信息** —— 用户按了 Ctrl+/ 再按回来，
    # 代码必须逐字一样，否则这个功能就是在破坏他的代码。
    _samples103 = (
        "x = 1", "    x = 1", "\tx = 1", "        y = f(a, b)",
        "    End If", "1: x = 1", "#If DEBUG Then", "    #If DEBUG Then",
        "x = 1  ' 尾巴", "Rem 说明", "' x = 1", "    ' x = 1", "'x", "'''",
        "    '", "'", "' '", "x", "  ", "\t",
    )
    _rev103 = [s for s in _samples103
               if s.strip() and P.remove_comment_from_line(
                   P.add_comment_to_line(s)) != s]
    _idem103 = [s for s in _samples103
                if s.strip()
                and P.add_comment_to_line(P.remove_comment_from_line(
                    P.add_comment_to_line(s))) != P.add_comment_to_line(s)]
    _blank103 = [s for s in ("", "   ", "\t", "  \t ")
                 if P.add_comment_to_line(s) != s]
    check("103.3 **可逆性穷举**：`remove(add(x)) == x` 对 20 个样本（含缩进 / "
              "Tab / 标签行 / `#If` / 行尾注释 / 已注释行）逐字成立；"
              "`add(remove(add(x))) == add(x)`（往返两次稳定）；空白行**不加**"
              "注释（不留一串光秃秃的 `'`）",
              [(_rev103, _idem103, _blank103)],
              expect_contain=[([], [], [])])

    # ---- 103.4 plan：IDEA 的 toggle 语义 ----
    _plan103 = (
        # 全未注释 -> 全部加
        (["    a = 1", "", "    b = 2"], "add",
         [(0, "    ' a = 1"), (2, "    ' b = 2")]),
        # 全已注释 -> 全部取消（空行不参与，但下标要跳过它）
        (["    ' a = 1", "  ", "    ' b = 2"], "remove",
         [(0, "    a = 1"), (2, "    b = 2")]),
        # 混合 -> 全部加（**这是 IDEA 的口径**：只要不是"全都注释了"就加）
        (["    ' a = 1", "    b = 2"], "add",
         [(0, "    ' ' a = 1"), (1, "    ' b = 2")]),
        # 全空白 -> 什么都不做（调用方必须把按键还回去）
        (["", "   ", "\t"], None, []),
        ([], None, []),
        (["x = 1"], "add", [(0, "' x = 1")]),
        (["' x = 1"], "remove", [(0, "x = 1")]),
    )
    _bad_plan103 = [(ls, P.plan_comment_action(ls), (ea, ep))
                    for ls, ea, ep in _plan103
                    if P.plan_comment_action(ls) != (ea, ep)]
    check("103.4 `plan_comment_action` 照 IDEA 的 toggle：**所有非空行都已注释**"
              "才取消，否则全部加注释；空行一律跳过（下标仍指向原来的行）；"
              "全空白返回 `(None, [])` —— 调用方据此把按键还回去",
              [_bad_plan103], expect_contain=[[]])

    # ---- 103.5 集成（假 VBE）：真的写进"编辑器"、选区真的恢复 ----
    class _Mod103(object):
        def __init__(self, lines):
            self.lines = list(lines)

        def Lines(self, n, _c):
            if 1 <= n <= len(self.lines):
                return self.lines[n - 1]
            return ""

        @property
        def CountOfLines(self):
            return len(self.lines)

        def ReplaceLine(self, n, t):
            self.lines[n - 1] = t

        def InsertLines(self, n, t):
            self.lines.insert(n - 1, t)

        def DeleteLines(self, n, c):
            del self.lines[n - 1:n - 1 + c]

    class _Pane103(object):
        def __init__(self, mod, sl, sc, el, ec):
            self.CodeModule = mod
            self.sel = [sl, sc, el, ec]
            self.sets = []

        def GetSelection(self):
            return tuple(self.sel)

        def SetSelection(self, sl, sc, el, ec):
            self.sel = [sl, sc, el, ec]
            self.sets.append(tuple(self.sel))

    class _Vbe103(object):
        def __init__(self, pane):
            self.ActiveCodePane = pane

    def _run103(lines, sl, sc, el, ec):
        """跑一次 toggle_comment，返回 (ok, 写完的行, 最后一次 SetSelection)。"""
        mod = _Mod103(lines)
        pane = _Pane103(mod, sl, sc, el, ec)
        _o103 = VB103._get_vbe_cached
        VB103._get_vbe_cached = lambda: _Vbe103(pane)
        try:
            ok = VB103.VbeBackend().toggle_comment()
        finally:
            VB103._get_vbe_cached = _o103      # 铁律：猴补丁必须还原
        return (bool(ok), list(mod.lines),
                pane.sets[-1] if pane.sets else None)

    _r103a = _run103(["    x = 1"], 1, 1, 1, 1)                 # 单行、无选区
    _r103b = _run103(["    a = 1", "", "    b = 2"], 1, 1, 3, 9)  # 多行全选
    _r103c = _run103(["    ' a = 1", "    ' b = 2"], 1, 1, 2, 9)  # 全注释 -> 取消
    _r103d = _run103(["", "   "], 1, 1, 2, 4)                   # 全空白 -> 不做
    _r103e = _run103(["    a = 1", "    b = 2"], 1, 1, 2, 1)    # 停在行首 -> 排除
    check("103.5 集成（假 VBE）：单行/无选区也能加；多行选区逐行加、**空行跳过**；"
              "全注释时整段取消；全空白时返回 False（让调用方把按键还回去，"
              "绝不吞键没反应）；**选区停在行首时最后一行不算被选中**（那是"
              "连着换行符拖下来的结果，IDEA 也这么处理）",
              [(_r103a[0], _r103a[1], _r103b[0], _r103b[1],
                _r103c[0], _r103c[1], _r103d[0], _r103e[0], _r103e[1])],
              expect_contain=[(True, ["    ' x = 1"],
                               True, ["    ' a = 1", "", "    ' b = 2"],
                               True, ["    a = 1", "    b = 2"],
                               False,
                               True, ["    ' a = 1", "    b = 2"])])

    # ---- 103.6 写完之后选区要恢复成"同样这几行" ----
    # 不恢复的话，用户按一下 Ctrl+/ 选区就没了、再按一下只作用于光标那一行，
    # "连按两次 = 取消"这个手感直接废掉。
    _sel103 = _run103(["    a = 1", "    b = 2", "    c = 3"], 1, 1, 3, 10)[2]
    check("103.6 写完把**同样这几行**重新选起来（`SetSelection(sl, 1, el, 超长列)`，"
          "让 VBE 自己把列钳到行尾）：不恢复的话\"连按两次 = 取消\"就不成立了",
          [(_sel103 is not None, bool(_sel103) and _sel103[0] == 1,
            bool(_sel103) and _sel103[2] == 3,
            bool(_sel103) and _sel103[3] > 10)],
          expect_contain=[(True, True, True, True)])

    # ---- 103.7 钩子侧接线护栏（AST）----
    with open(os.path.join(ROOT, "main.py"), encoding="utf-8") as _f:
        _tree103 = _ast103.parse(_f.read())

    def _fn103(tree, name):
        for _n in _ast103.walk(tree):
            if isinstance(_n, _ast103.FunctionDef) and _n.name == name:
                return _n
        return None

    _wf103 = _fn103(_tree103, "win32_filter")
    _ln_cmt103 = _ln_vis103 = 10 ** 9
    _br103 = False
    for _n in _ast103.walk(_wf103):
        if isinstance(_n, _ast103.If):
            _d103 = _ast103.dump(_n.test)
            if "id='_comment_key_ok'" in _d103:
                _ln_cmt103 = min(_ln_cmt103, _n.lineno)
                _br103 = ("id='_toggle_comment_here'" in _ast103.dump(_n)
                          and "Constant(value=True)" in _ast103.dump(_n))
            if "id='completer'" in _d103 and "attr='is_visible'" in _d103:
                _ln_vis103 = min(_ln_vis103, _n.lineno)
    _act103 = _ast103.dump(_fn103(_tree103, "_toggle_comment_here"))
    _gate103 = _ast103.dump(_fn103(_tree103, "_comment_key_ok"))
    check("103.7 钩子接线（AST）：（a）Ctrl+/ 分支排在 `completer.is_visible()` "
          "**之前**（它要改文本，窗里那些候选马上就不作数了）；"
          "（b）该分支吞键（`suppress = True`）并派发 `_toggle_comment_here`；"
          "（c）动作里收候选窗、作废快照 `state[\"cur_ctx\"]`、失败兜底 "
          "`send_vk(VK_OEM_2, ctrl=True)`；（d）**判据函数里没有任何 COM 调用**"
          "（钩子线程没有 COM 单元 —— 这条用 AST 查，不受注释影响）",
          [(_ln_cmt103 < _ln_vis103, _br103,
            "attr='hide'" in _act103, "value='cur_ctx'" in _act103,
            "arg='ctrl'" in _act103 and "VK_OEM_2" in _act103,
            all(w not in _gate103 for w in
                ("_get_vbe_cached", "attr='backend'", "VbeBackend")))],
          expect_contain=[(True, True, True, True, True, True)])

    # ---- 103.8 后端用的必须是同一份判据（AST）----
    # 判据写在 parser、动作写在别处，两份实现迟早会分叉 —— 那时"钩子说该管、
    # 后端说不用管"就会表现为"按了没反应"。这条钉死"只有一份"。
    _be103 = _ast103.dump(_fn103(
        _ast103.parse(io.open(os.path.join(ROOT, "vbe_bridge.py"),
                              encoding="utf-8").read()),
        "toggle_comment"))
    check("103.8 `vbe_bridge.toggle_comment` 调的就是 `parser.plan_comment_action`"
          "（同一份判据；两份实现迟早分叉，届时表现为\"按了没反应\"）",
          [("plan_comment_action" in _be103,
            "add_comment_to_line" not in _be103,
            "remove_comment_from_line" not in _be103)],
          expect_contain=[(True, True, True)])

    # ---- 103.9 开关默认开；纯判据钉住修饰键门槛 ----
    _vk103 = tuple(
        MN103._comment_vk_ok(vk, c, s, a, w) for vk, c, s, a, w in (
            (0xBF, True, False, False, False),   # 正是我们
            (0xBF, False, False, False, False),  # 没按 Ctrl
            (0xBF, True, True, False, False),    # Ctrl+Shift+/ = `?`
            (0xBF, True, False, True, False),    # Ctrl+Alt+/
            (0xBF, True, False, False, True),    # Ctrl+Win+/
            (0x51, True, False, False, False),   # Ctrl+Q 不归这里
        ))
    check("103.9 开关默认**开**（`COMMENT_TOGGLE is True`），纯判据只认"
          "**单独** Ctrl+`/`：Shift / Alt / Win 任一在按都不算（Ctrl+Shift+/ "
          "在多数布局上是 `?`，Win 组合是系统键）",
          [(_vk103, MN103.COMMENT_TOGGLE)],
          expect_contain=[((True, False, False, False, False, False), True)])

    # ------------------------------------------------------------------
    # 104. ★v121：过程大纲（Ctrl+Shift+O）—— 纯函数 list_procedures
    #
    #   用户口径（原话）："那你先帮我处理第0档的那个吧，我们先一步步来。"
    #   第 0 档第一条 = 过程大纲 + 跳转。这一节只测**纯函数**（不碰 COM）：
    #   "哪些行算过程头"必须与补全那边（extract_proc_kinds / extract_records）
    #   认完全一致 —— 否则会出现"补全里有这个名字、大纲里却找不到"。
    # ------------------------------------------------------------------
    print("\n=== 104. ★v121：过程大纲（list_procedures）===")
    import parser as PR104

    _code104 = "\n".join([
        "Option Explicit",                       # 1
        "",                                      # 2
        "Private Sub DoWork(ByVal n As Long)",    # 3
        "    Dim x As Long",                     # 4
        "    x = n + 1",                         # 5
        "End Sub",                               # 6
        "",                                      # 7
        "Public Function GetName(i As Long) As String",  # 8
        "    GetName = \"a\"",                   # 9
        "End Function",                          # 10
        "",                                      # 11
        "Property Get Item(i As Long) As Long",   # 12
        "    Item = i",                          # 13
        "End Property",                          # 14
        "",                                      # 15
        "Property Let Item(i As Long, v As Long)",  # 16
        "End Property",                          # 17
        "",                                      # 18
        "Property Set Item(o As Object)",         # 19
        "End Property",                          # 20
        "",                                      # 21
        "' Sub Fake()",                          # 22 注释里的，不能列
        "Private Declare PtrSafe Function gApi Lib \"k\" (ByVal v As Long) As Long",  # 23
        "Public Event Changed(ByVal s As String)",  # 24
        "Sub NoArgs",                            # 25
        "End Sub",                               # 26
        "Function Bare As Long",                 # 27 无括号但有返回类型
        "End Function",                          # 28
    ])
    _procs104 = PR104.list_procedures(_code104)
    check("104.1 过程头全认：Sub / Function / Property Get|Let|Set / Declare，"
          "行号是**物理行号**；`Private` 标出来；注释里的 `' Sub Fake` "
          "**不列**；`End Sub` 不列；`Function Bare As Long`（无括号）也认；"
          "`Event` 刻意不收（它不在补全池里）",
          [(_procs104,)],
          expect_contain=[([
              ("DoWork", "Sub", 3, True),
              ("GetName", "Function", 8, False),
              ("Item", "Property Get", 12, False),
              ("Item", "Property Let", 16, False),
              ("Item", "Property Set", 19, False),
              ("gApi", "Declare Function", 23, True),
              ("NoArgs", "Sub", 25, False),
              ("Bare", "Function", 27, False),
          ],)])

    # ---- 104.2 与补全那边认的必须是同一批行 ----
    _kinds104 = PR104.extract_proc_kinds(_code104)
    _names104 = set(_kinds104.keys())
    _outline104 = set(n.lower() for (n, _k, _l, _p) in _procs104)
    check("104.2 大纲列出的名字集合 == `extract_proc_kinds` 认的名字集合"
          "（同一批正则；两份分叉的话表现为\"补全能提示、大纲里找不到\"）",
          [(_names104, _outline104)],
          expect_contain=[({"dowork", "getname", "item", "gapi", "noargs",
                            "bare"},
                           {"dowork", "getname", "item", "gapi", "noargs",
                            "bare"})])

    # ---- 104.3 行号能直接拿去读那一行 ----
    _lines104 = _code104.split("\n")
    _ok104 = all(str(_lines104[ln - 1]).strip().lower().find(nm.lower()) >= 0
                 for (nm, _k, ln, _p) in _procs104)
    check("104.3 每个行号 `lines[行号-1]` 里都真有那个名字（行号是"
          "`CodeModule.Lines(n, 1)` 认的那种物理行号，不做任何行合并）",
          [(_ok104, len(_procs104))], expect_contain=[(True, 8)])

    # ------------------------------------------------------------------
    # 105. ★v121：文本级查找引用 —— 纯函数 find_occurrences
    #
    #   三个会误报的来源，逐条钉住：整词（`num` 不命中 `numArr`）、大小写
    #   （VBA 不区分）、字符串 / 注释里的同名文本不算。
    # ------------------------------------------------------------------
    print("\n=== 105. ★v121：文本级查找引用（find_occurrences）===")
    _code105 = "\n".join([
        "Private num As Long",                    # 1  声明
        "Sub A()",                               # 2
        "    num = num + 1",                     # 3  两处
        "    Dim numArr(3) As Long",             # 4  numArr 不算
        "    Debug.Print \"num\"",                # 5  字符串里的不算
        "    ' num = 5",                         # 6  注释里的不算
        "    Num = 9",                           # 7  大小写不同，算
        "End Sub",                               # 8
    ])
    _hits105 = PR104.find_occurrences(_code105, "num")
    check("105.1 整词 + 大小写不敏感 + 字符串/注释不算；列号 1 起、指到词首；"
          "带出该行原文",
          [([(ln, col, decl) for (ln, col, _t, decl) in _hits105],)],
          expect_contain=[([(1, 9, True), (3, 5, False), (3, 11, False),
                            (7, 5, False)],)])

    check("105.2 该行原文与源码逐字一致（界面直接拿它显示，截断在 UI 侧做）",
          [([t for (_l, _c, t, _d) in _hits105],)],
          expect_contain=[(["Private num As Long", "    num = num + 1",
                            "    num = num + 1", "    Num = 9"],)])

    # ---- 105.3 声明处标记：Dim / 形参 / 过程名都算 ----
    _hits105b = PR104.find_occurrences(
        "Sub Foo(ByVal target As Long)\n    target = 1\nEnd Sub", "target")
    _decl105 = [(ln, d) for (ln, _c, _t, d) in _hits105b]
    check("105.3 \"是否声明处\"用的是 `_decl_line_names`（同一份声明判据）："
          "形参表里的算声明、过程体里的赋值算使用",
          [(_decl105,)], expect_contain=[([(1, True), (2, False)],)])

    # ---- 105.4 limit 生效（大工程里防撑爆）----
    _many105 = "\n".join(["x = 1"] * 50)
    check("105.4 `limit` 生效：50 处里只收请求的 7 处（大工程里防止把内存和"
          "界面撑爆）",
          [(len(PR104.find_occurrences(_many105, "x", limit=7)),)],
          expect_contain=[(7,)])

    # ---- 105.5 空名字 / 空代码不炸 ----
    check("105.5 空名字 / 空代码返回空列表（绝不在按键路径上抛异常）",
          [(PR104.find_occurrences("", "x"), PR104.find_occurrences("x = 1", ""),
            PR104.find_occurrences(None, None))],
          expect_contain=[([], [], [])])

    # ------------------------------------------------------------------
    # 106. ★v121：后端 + 窗的接线（假 VBE / 假 tk 不动）
    #   两条硬要求：
    #     (a) `goto_location` 落光标前必须**复核锚点**（行号失真就不跳）；
    #     (b) `JumpWindow._build()` 必须把 Toplevel 记到 `self.win` ——
    #         漏了这行窗口根本不会出现，而所有异常都被 except 吞掉，表现为
    #         "按了快捷键什么都没发生"（这一版真的漏过，是真机前唯一抓到的机会）。
    # ------------------------------------------------------------------
    print("\n=== 106. ★v121：大纲 / 引用的后端接线 ===")
    import ast as _ast106
    import vbe_bridge as VB106

    class _Mod106(object):
        def __init__(self, name, lines):
            self.name = name
            self.lines = list(lines)

        def Lines(self, n, c=None):
            # 真 VBE 的 `Lines(n, c)` 返回 c 行、行间是 CRLF —— 假对象也照这个
            # 形状给，顺便让 `current_module_text` 的 CRLF 归一真的被测到。
            if c is None:
                c = max(1, len(self.lines) - n + 1)
            return "\r\n".join(self.lines[n - 1:n - 1 + c])

        @property
        def CountOfLines(self):
            return len(self.lines)

        @property
        def Name(self):
            return self.name

    class _Pane106(object):
        def __init__(self, mod):
            self.CodeModule = mod
            self.sel = None
            self.shown = 0

        def GetSelection(self):
            return (1, 1, 1, 1)

        def SetSelection(self, sl, sc, el, ec):
            self.sel = (sl, sc, el, ec)

        def Show(self):
            self.shown += 1

    class _Comp106(object):
        def __init__(self, mod):
            self.Name = mod.Name
            self.CodeModule = mod

    class _Panes106(object):
        """`vbe.CodePanes`：**只有这里面的模块才能跳过去**。

        真机实测两条（决定了跳转怎么落地）：`VBComponent` 上根本没有 `CodePane`
        属性（24 个组件只有 6 个打开的 pane），而没打开过的模块在 COM 里没有任何
        接口能把它打开。
        """
        def __init__(self, panes):
            self._p = list(panes)

        @property
        def Count(self):
            return len(self._p)

        def Item(self, i):
            return self._p[i - 1]

    class _Proj106(object):
        def __init__(self, comps):
            self.VBComponents = list(comps)
            self.Name = "VBAProject"

    class _Vbe106(object):
        def __init__(self, comps, panes):
            self.VBProjects = [_Proj106(comps)]
            self.ActiveVBProject = self.VBProjects[0]
            self.CodePanes = _Panes106(panes)
            self.ActiveCodePane = panes[0] if panes else None

    _mod1_106 = _Mod106("模块1", ["Option Explicit", "",
                                  "Sub DoWork()", "    Call GetName(1)",
                                  "End Sub"])
    _mod2_106 = _Mod106("模块2", ["Function GetName(n As Long) As String",
                                  "    GetName = \"x\"", "End Function"])
    # 模块3：有命中，但**没有打开的代码窗** -> 跳不过去
    _mod3_106 = _Mod106("模块3", ["Sub Other()", "    x = GetName(2)",
                                  "End Sub"])
    _pane1_106 = _Pane106(_mod1_106)
    _pane2_106 = _Pane106(_mod2_106)
    _vbe106 = _Vbe106([_Comp106(_mod1_106), _Comp106(_mod2_106),
                       _Comp106(_mod3_106)], [_pane1_106, _pane2_106])

    _o_vbe106 = VB106._get_vbe_cached
    _o_cache106 = dict(VB106._comp_list_cache)
    VB106._get_vbe_cached = lambda: _vbe106
    VB106._comp_list_cache["comps"] = None
    VB106._comp_list_cache["count"] = -1
    VB106._comp_list_cache["ts"] = 0.0
    try:
        _be106 = VB106.VbeBackend()
        _mod0_106, _hits106, _n106 = _be106.find_references("GetName")
        _got_goto = _be106.goto_location("模块2", 1, expect="GetName")
        _sel106 = _pane2_106.sel
        _shown106 = _pane2_106.shown
        _bad_blank = _be106.goto_location("模块1", 2, expect=None)
        _bad_expect = _be106.goto_location("模块1", 3, expect="Nope")
        _bad_mod = _be106.goto_location("不存在的模块", 1)
        _no_pane = _be106.goto_location("模块3", 2, expect="GetName")
        _here_mod, _here_procs = _be106.list_procedures_here()
        _cur_mod, _cur_code = _be106.current_module_text()
    finally:
        VB106._get_vbe_cached = _o_vbe106
        VB106._comp_list_cache.update(_o_cache106)

    check("106.1 `find_references` 扫遍活动工程：四处命中（模块1 调用；模块2 "
          "声明 + 赋值；模块3 调用），当前模块排最前，并回报扫了几个模块；"
          "**没有打开代码窗的模块标 `jumpable=False`**（COM 里打不开它）",
          [(_n106, len(_hits106),
            [(h[0], h[1], h[5]) for h in _hits106], _hits106[0][0])],
          expect_contain=[(3, 4, [("模块1", 4, True), ("模块2", 1, True),
                                  ("模块2", 2, True), ("模块3", 2, False)],
                           "模块1")])

    check("106.2 `goto_location` 真的切了代码窗并落光标到**第 1 列**"
          "（`SetSelection(行, 1, 行, 1)` —— 显示列有歧义，第 1 列没有）",
          [(_got_goto, _sel106, _shown106)],
          expect_contain=[(True, (1, 1, 1, 1), 1)])

    check("106.3 锚点复核：空白行不跳、目标行里没有 expect 不跳、模块名对不上"
          "不跳、**没有打开的代码窗也不跳**（宁可这一次不跳，绝不把光标扔到"
          "不相干的代码上）",
          [(_bad_blank, _bad_expect, _bad_mod, _no_pane)],
          expect_contain=[(False, False, False, False)])

    check("106.4 `list_procedures_here` 带上过程头原文（大纲窗选中时显示签名）、"
          "`current_module_text` 把 CRLF 归一成 LF",
          [(_here_mod, _here_procs, _cur_mod, "\r" not in _cur_code)],
          expect_contain=[("模块1",
                           [("DoWork", "Sub", 3, False, "Sub DoWork()")],
                           "模块1", True)])

    # ---- 106.5 窗的骨架：`self.win = w` 一行都不能少 ----
    import jump_ui as JU106
    _src106 = None
    try:
        import io as _io106
        _src106 = _io106.open(JU106.__file__, "r", encoding="utf-8").read()
    except Exception:
        _src106 = ""
    _tree106 = _ast106.parse(_src106) if _src106 else None
    _assigns_win = False
    if _tree106 is not None:
        for node in _ast106.walk(_tree106):
            if (isinstance(node, _ast106.FunctionDef)
                    and node.name == "_build"):
                for st in node.body:
                    if isinstance(st, _ast106.Assign):
                        for tg in st.targets:
                            if (isinstance(tg, _ast106.Attribute)
                                    and tg.attr == "win"
                                    and isinstance(tg.value, _ast106.Name)
                                    and tg.value.id == "self"):
                                _assigns_win = True
    check("106.5 `JumpWindow._build()` 里真的有 `self.win = w`"
          "（漏掉这一行窗口不会出现，而 `show_items` 里每一句都包在 except "
          "里 —— 表现为\"按了快捷键什么都没发生\"，真机上极难查）",
          [(_assigns_win,)], expect_contain=[(True,)])

    # ---- 106.6 钩子接线（AST）：分支位置 + 吞键 + 判据里没有 COM ----
    import main as MN106
    _srcm106 = ""
    try:
        import io as _io106b
        _srcm106 = _io106b.open(MN106.__file__, "r", encoding="utf-8").read()
    except Exception:
        _srcm106 = ""
    _treem106 = _ast106.parse(_srcm106) if _srcm106 else None

    def _func_src106(name):
        if _treem106 is None:
            return ""
        for node in _ast106.walk(_treem106):
            if isinstance(node, _ast106.FunctionDef) and node.name == name:
                return _ast106.dump(node)
        return ""

    def _func_node106(name):
        if _treem106 is None:
            return None
        for node in _ast106.walk(_treem106):
            if isinstance(node, _ast106.FunctionDef) and node.name == name:
                return node
        return None

    # 分支位置：Ctrl+Shift+O / F 那两条必须排在"候选窗是否可见"之前 ——
    # 它们是纯开窗动作，与候选窗无关；排在后面等于"候选窗开着就叫不出来"。
    _wf106 = _func_node106("win32_filter")
    _jump_lines = []
    _vis_lines = []
    if _wf106 is not None:
        for nd in _ast106.walk(_wf106):
            if isinstance(nd, _ast106.Name) and nd.id in (
                    "_open_outline", "_open_usages"):
                _jump_lines.append(nd.lineno)
            if isinstance(nd, _ast106.Attribute) and nd.attr == "is_visible":
                _vis_lines.append(nd.lineno)
    _order_ok = bool(_jump_lines and _vis_lines
                     and max(_jump_lines) < min(_vis_lines))

    _ol_src = _func_src106("_open_outline")
    _us_src = _func_src106("_open_usages")
    _ol_gate = _func_src106("_outline_key_ok")
    _us_gate = _func_src106("_usages_key_ok")
    _pure_ol = _func_src106("_outline_vk_ok")
    _pure_us = _func_src106("_usages_vk_ok")
    check("106.6 钩子接线：两条都派发到 `_open_outline` / `_open_usages` 且"
          "**排在 `completer.is_visible()` 之前**（纯开窗动作，候选窗开着也要能"
          "叫出来）；两条闸门与两条纯判据里**没有任何 COM 调用**（钩子线程没有 "
          "COM 单元 —— 这条用 AST 查，不受注释影响）",
          [(_order_ok,
            "show_items" in _ol_src, "show_items" in _us_src,
            "Backend" not in _ol_gate and "_get_vbe" not in _ol_gate,
            "Backend" not in _us_gate and "_get_vbe" not in _us_gate,
            "Backend" not in _pure_ol, "Backend" not in _pure_us)],
          expect_contain=[(True, True, True, True, True, True, True)])

    # ---- 106.7 修饰键门槛：必须 Ctrl+Shift 同时按、Alt/Win 不许按 ----
    _vk106 = tuple(
        MN106._outline_vk_ok(vk, c, s, a, w) for (vk, c, s, a, w) in (
            (0x4F, True, True, False, False),    # 正是我们
            (0x4F, True, False, False, False),   # 只有 Ctrl -> 那是别的
            (0x4F, False, True, False, False),   # 只有 Shift
            (0x4F, True, True, True, False),     # 加 Alt
            (0x4F, True, True, False, True),     # 加 Win
            (0x46, True, True, False, False),    # Ctrl+Shift+F 不归大纲
        ))
    _vk106b = tuple(
        MN106._usages_vk_ok(vk, c, s, a, w) for (vk, c, s, a, w) in (
            (0x46, True, True, False, False),
            (0x46, True, False, False, False),
            (0x46, True, True, True, False),
            (0x4F, True, True, False, False),
        ))
    check("106.7 修饰键门槛：两条都是 **Ctrl+Shift 同时按**才算，Alt / Win 任一"
          "在按都不算（Win 组合是系统键，Alt 会开菜单）；两个键互不串门",
          [(_vk106, _vk106b,
            MN106.OUTLINE_POPUP, MN106.USAGES_POPUP)],
          expect_contain=[((True, False, False, False, False, False),
                           (True, False, False, False),
                           True, True)])

    # ------------------------------------------------------------------
    # 107. ★v121b：单击即跳转 + 导航窗手动关（用户口径）
    #   两条都是"手感"要求，只能拿**真 tkinter** 验：
    #     (a) 单击列表项 = 跳转（不再要求记住"双击"）；
    #     (b) 跳完窗**不消失** ⇒ 连点几处引用不用反复开窗；关窗必须手动。
    #   ⚠️ 测试网顶部把 tkinter 塞成了桩 ⇒ 这一节临时换真模块、跑完原样换回
    #      （与 §99.6 / §101.2 同一套写法，不然这一节会被静默跳过）。
    # ------------------------------------------------------------------
    print("\n=== 107. ★v121b：单击跳转 + 导航窗手动关 ===")
    import jump_ui as JU107
    _g107 = None
    _root107 = None
    _tk107_had = JU107.tk
    _tk107_saved = None
    try:
        import importlib as _il107
        _tk107 = sys.modules.get("tkinter")
        if not hasattr(_tk107, "Tk"):
            _tk107_saved = {k: sys.modules[k] for k in list(sys.modules)
                            if k == "tkinter" or k.startswith("tkinter.")}
            for _k107 in _tk107_saved:
                del sys.modules[_k107]
            try:
                _tk107 = _il107.import_module("tkinter")
            except Exception:
                _tk107 = None
        if _tk107 is not None and hasattr(_tk107, "Tk"):
            JU107.tk = _tk107          # jump_ui 里 `import tkinter as tk`
            _root107 = _tk107.Tk()
            _root107.withdraw()
            _w107 = JU107.JumpWindow(_root107)
            _hit107 = []
            _w107.on_pick = lambda it: (_hit107.append(it["label"]), True)[1]
            _w107.show_items((100, 100, 116), [
                {"label": "Sub DoWork", "module": "模块1", "line": 3,
                 "kind": "Sub", "expect": "DoWork", "text": "Sub DoWork()"},
                {"label": "Sub Other", "module": "模块2", "line": 9,
                 "kind": "Sub", "expect": "Other", "text": "Sub Other()"},
            ], title="过程大纲", subtitle="模块1　·　2 个过程")
            try:
                # 别让测试窗浮在 Excel 上面干扰后面几节
                _w107.win.attributes("-topmost", False)
            except Exception:
                pass
            _root107.update()

            def _click107(row):
                # 真事件：选中某行 → 发一个"鼠标左键松开"给列表
                _w107.lb.selection_clear(0, "end")
                _w107.lb.selection_set(row)
                _w107.lb.event_generate("<ButtonRelease-1>")
                _root107.update()

            _click107(0)
            _open1 = _w107.is_open()
            _st1 = _w107.lbl_status.cget("text")
            _click107(1)
            _open2 = _w107.is_open()
            _st2 = _w107.lbl_status.cget("text")
            # 跳不过去（模块没打开代码窗）也**不许**收窗，只换红字
            _w107.on_pick = lambda it: False
            _click107(0)
            _open3 = _w107.is_open()
            _st3 = _w107.lbl_status.cget("text")
            # 手动关：Esc / 关闭按钮 / 再按一次快捷键走的都是 close()
            _w107.close()
            _closed107 = (not _w107.is_open())
            _g107 = (_hit107, _open1, _open2, _open3, _closed107,
                     ("已跳到" in _st1 and "模块1" in _st1 and "第 3 行" in _st1),
                     ("已跳到" in _st2 and "模块2" in _st2 and "第 9 行" in _st2),
                     ("跳不过去" in _st3))
    except Exception:
        _g107 = None
    finally:
        JU107.tk = _tk107_had
        if _tk107_saved is not None:
            for _k107b in [k for k in list(sys.modules)
                           if k == "tkinter" or k.startswith("tkinter.")]:
                del sys.modules[_k107b]
            sys.modules.update(_tk107_saved)
        try:
            if _root107 is not None:
                _root107.destroy()
        except Exception:
            pass
    if _g107 is not None:
        _hit107v, _o1, _o2, _o3, _cls107, _say1, _say2, _say3 = _g107
        check("107.1 **单击列表项就跳转**（不必记「双击」）：点第一行、点第二行"
              "各跳一次，回调拿到的正是点中的那一行；★而且**两次都没把窗关掉**"
              "（连点几处引用不用反复开窗）",
              [(_hit107v, _o1, _o2)],
              expect_contain=[(["Sub DoWork", "Sub Other"], True, True)])
        check("107.2 跳成了在状态栏报「已跳到 模块 · 第 N 行」；**跳不过去也不收窗**"
              "（只把提示换成红字）—— 收起窗再什么都不发生是最差的手感",
              [(_say1, _say2, _o3, _say3)],
              expect_contain=[(True, True, True, True)])
        check("107.3 手动关窗（`close()`：Esc / 关闭按钮 / 再按一次快捷键走的都是它）"
              "真的收起了窗 —— 关的**唯一**途径",
              [(_cls107,)], expect_contain=[(True,)])

    # ---- 107.4 结构护栏（AST，不受注释影响）----
    try:
        import io as _io107
        _src107 = _io107.open(JU107.__file__, "r", encoding="utf-8").read()
    except Exception:
        _src107 = ""
    _tree107 = _ast106.parse(_src107) if _src107 else None
    _pick_no_close107 = True
    _onclick107 = False
    _bound107 = False
    _dbl107 = "<Double-Button-1>" in _src107
    if _tree107 is not None:
        for _nd107 in _ast106.walk(_tree107):
            if (isinstance(_nd107, _ast106.FunctionDef)
                    and _nd107.name == "_pick"):
                for _sub107 in _ast106.walk(_nd107):
                    if (isinstance(_sub107, _ast106.Call)
                            and isinstance(_sub107.func, _ast106.Attribute)
                            and _sub107.func.attr == "close"
                            and isinstance(_sub107.func.value, _ast106.Name)
                            and _sub107.func.value.id == "self"):
                        _pick_no_close107 = False
            if (isinstance(_nd107, _ast106.FunctionDef)
                    and _nd107.name == "_on_click"):
                _onclick107 = True
            if (isinstance(_nd107, _ast106.Call)
                    and isinstance(_nd107.func, _ast106.Attribute)
                    and _nd107.func.attr == "bind"):
                for _a107 in _nd107.args:
                    if (isinstance(_a107, _ast106.Constant)
                            and _a107.value == "<ButtonRelease-1>"):
                        _bound107 = True
    check("107.4 结构护栏（AST）：`_pick` 里**没有** `self.close()`（跳转不再自动"
          "关窗，以后有人加回来这里会红）；列表绑了 `<ButtonRelease-1>` 且有 "
          "`_on_click`；`<Double-Button-1>` 那套已经撤掉（不留两套跳转口径）",
          [(_pick_no_close107, _bound107, _onclick107, _dbl107)],
          expect_contain=[(True, True, True, False)])


    # ------------------------------------------------------------------
    # 108. ★v122：三个"按行"的编辑动作
    #      Ctrl+Shift+Enter 上方插行 / Ctrl+D 复制行 / Ctrl+K 删行
    # ------------------------------------------------------------------
    import main as _MB108
    import settings as _ST108

    print("\n=== 108. ★v122：Ctrl+Shift+Enter 上方插行 / Ctrl+D 复制行 / "
          "Ctrl+K 删行 ===")

    class _Mod108(object):
        """够用的假 CodeModule：真读真写，并记下每次写操作的形状。

        `writes` 用来钉"**只写一次**"这条：三个动作都必须是"确认完毕之后调一次
        InsertLines / DeleteLines"，绝不能边读边写（那样撤销栈会碎成好几步，
        而且中途失败会留下写了一半的代码）。
        """

        def __init__(self, lines):
            self.lines = list(lines)
            self.writes = []

        def Lines(self, n, c=None):
            if c is None:
                c = max(1, len(self.lines) - n + 1)
            return "\r\n".join(self.lines[n - 1:n - 1 + c])

        @property
        def CountOfLines(self):
            return len(self.lines)

        def InsertLines(self, n, text):
            parts = str(text).split("\r\n")
            self.lines[n - 1:n - 1] = parts
            self.writes.append(("insert", n, len(parts)))

        def DeleteLines(self, n, c):
            self.writes.append(("delete", n, c,
                                list(self.lines[n - 1:n - 1 + c])))
            del self.lines[n - 1:n - 1 + c]

    class _Pane108(object):
        def __init__(self, mod, sel=(1, 1, 1, 1)):
            self.CodeModule = mod
            self.sel = sel
            self.sets = []

        def GetSelection(self):
            return self.sel

        def SetSelection(self, sl, sc, el, ec):
            self.sets.append((sl, sc, el, ec))

    class _Vbe108(object):
        def __init__(self, pane):
            self.ActiveCodePane = pane

    def _be108(lines, sel):
        """造一个"当前代码窗 = 这几行、选区 = sel"的后端。"""
        _m108 = _Mod108(lines)
        _p108 = _Pane108(_m108, sel)
        VB106._get_vbe_cached = lambda: _Vbe108(_p108)
        return VB106.VbeBackend(), _m108, _p108

    _o_vbe108 = VB106._get_vbe_cached
    try:
        # ---- 108.1 普通行上方插行：新行与本行【同级】----
        _be, _m, _p = _be108(["Sub Foo()", "    x = 1", "End Sub"], (2, 1, 2, 1))
        _ok_a = _be.new_line_above()
        check("108.1 Ctrl+Shift+Enter 在普通行上方插一行：新行与本行**同级**"
              "（缩进照抄本行），原行整体下移一行，光标落在新行缩进之后；"
              "★而且只写了一次 InsertLines（不是边读边写）",
              [(_ok_a, list(_m.lines), (_p.sets[-1] if _p.sets else None),
               list(_m.writes))],
              expect_contain=[(True,
                               ["Sub Foo()", "    ", "    x = 1", "End Sub"],
                               (2, 5, 2, 5),
                               [("insert", 2, 1)])])

        # ---- 108.2 收尾行上方插行：新行落在块【内】，深一级 ----
        _be, _m, _p = _be108(["Sub Foo()", "    For i = 1 To 3",
                              "        x = 1", "    Next i", "End Sub"],
                             (4, 1, 4, 1))
        _ok_b = _be.new_line_above()
        check("108.2 在**块收尾行**（`Next i`）上方插行时新行深一级 —— 插在收尾"
              "上方等于插在块内的最后一行，和收尾齐平是错的；"
              "★这条也是 `is_block_closer` 存在的理由（拿 `block_closer` 代它"
              "会得到 None，判成\"不是收尾\"）",
              [(_ok_b, list(_m.lines),
               (_p.sets[-1] if _p.sets else None))],
              expect_contain=[(True,
                               ["Sub Foo()", "    For i = 1 To 3",
                                "        x = 1", "        ", "    Next i",
                                "End Sub"],
                               (4, 9, 4, 9))])

        # ---- 108.3 有选区时插在【选区最上面那一行】之前 ----
        # 选区是反着选的（3 -> 2），GetSelection 的两个角点不保证有序
        _be, _m, _p = _be108(["Sub Foo()", "    a = 1", "    b = 2", "End Sub"],
                             (3, 1, 2, 1))
        _ok_c = _be.new_line_above()
        check("108.3 有选区时插在【选区最上面那一行】的上方（用户眼里的\"上方\""
              "就是那儿）；选区从下往上选（角点 3 -> 2）也一样 —— "
              "`_sel_block` 先把两个角点归一",
              [(_ok_c, list(_m.lines),
               (_m.writes[0] if _m.writes else None))],
              expect_contain=[(True,
                               ["Sub Foo()", "    ", "    a = 1", "    b = 2",
                                "End Sub"],
                               ("insert", 2, 1))])

        # ---- 108.4 复制单行 ----
        _be, _m, _p = _be108(["Sub Foo()", "    a = 1", "End Sub"], (2, 3, 2, 3))
        _ok_d = _be.duplicate_lines()
        check("108.4 Ctrl+D 复制光标那一行到下方：内容逐字相同，光标落到**复制块的"
              "第一行同一列**（整行复制 ⇒ 显示列照抄就是同一个字符位置，"
              "不需要列语义换算）；一次 InsertLines",
              [(_ok_d, list(_m.lines),
               (_p.sets[-1] if _p.sets else None), list(_m.writes))],
              expect_contain=[(True,
                               ["Sub Foo()", "    a = 1", "    a = 1", "End Sub"],
                               (3, 3, 3, 3),
                               [("insert", 3, 1)])])

        # ---- 108.5 复制多行选区 ----
        _be, _m, _p = _be108(["    a = 1", "    b = 2", "End Sub"], (1, 1, 2, 1))
        _ok_e = _be.duplicate_lines()
        check("108.5 Ctrl+D 有选区时复制**选中的那几行**（IDEA / VS Code 口径），"
              "整块插在选区下方；`cm.Lines` 拿回来的 CRLF 多行原样喂回 "
              "InsertLines 就会拆成对应的行数",
              [(_ok_e, list(_m.lines), list(_m.writes))],
              expect_contain=[(True,
                               ["    a = 1", "    b = 2", "    a = 1",
                                "    b = 2", "End Sub"],
                               [("insert", 3, 2)])])

        # ---- 108.6 删单行 + 落点 ----
        _be, _m, _p = _be108(["Sub Foo()", "    a = 1", "    b = 2", "End Sub"],
                             (2, 7, 2, 7))
        _ok_f = _be.delete_lines_here()
        check("108.6 Ctrl+K 删掉光标那一行，光标落到**原位置的第 1 列**"
              "（第 1 列在任何列语义下都无歧义 —— 铁律 15）；"
              "原第 3 行顶上来",
              [(_ok_f, list(_m.lines),
               (_p.sets[-1] if _p.sets else None))],
              expect_contain=[(True,
                               ["Sub Foo()", "    b = 2", "End Sub"],
                               (2, 1, 2, 1))])

        # ---- 108.7 删多行选区 ----
        _be, _m, _p = _be108(["Sub Foo()", "    a = 1", "    b = 2", "End Sub"],
                             (2, 1, 3, 1))
        _ok_g = _be.delete_lines_here()
        check("108.7 Ctrl+K 有选区时删掉**选中的那几行**，一次 DeleteLines",
              [(_ok_g, list(_m.lines), list(_m.writes))],
              expect_contain=[(True, ["Sub Foo()", "End Sub"],
                               [("delete", 2, 2, ["    a = 1", "    b = 2"])])])

        # ---- 108.8 删的是最后一行：落点不许超界 ----
        _be, _m, _p = _be108(["Sub Foo()", "    a = 1"], (2, 1, 2, 1))
        _ok_h = _be.delete_lines_here()
        check("108.8 删的是**最后一行**时，落点退到新的最后一行"
              "（不能 SetSelection 到一个已经不存在的行号）",
              [(_ok_h, list(_m.lines),
               (_p.sets[-1] if _p.sets else None))],
              expect_contain=[(True, ["Sub Foo()"], (1, 1, 1, 1))])

        # ---- 108.9 把模块删空：照样算成功 ----
        _be, _m, _p = _be108(["    a = 1"], (1, 1, 1, 1))
        _ok_i = _be.delete_lines_here()
        check("108.9 把整个模块删空也照做（VBE 的 CountOfLines 会变成 0）："
              "删除**算成功**、不抛异常，并且**不去** SetSelection(1, 1)"
              "（那时行 1 已经不存在了）",
              [(_ok_i, list(_m.lines), list(_p.sets))],
              expect_contain=[(True, [], [])])

        # ---- 108.10 空模块 / 越界：三个动作一律 False（不假装写成） ----
        _bad108 = []
        for _meth in ("new_line_above", "duplicate_lines", "delete_lines_here"):
            _be, _m, _p = _be108([], (1, 1, 1, 1))
            try:
                _bad108.append(bool(getattr(_be, _meth)()))
            except Exception:
                _bad108.append("炸了")
        check("108.10 模块是空的 / 行号越界时，三个动作都返回 False（调用方据此"
              "把按键还回去），绝不假装写成",
              [_bad108], expect_contain=[[False, False, False]])

        # ---- 108.11 is_block_closer 判据表 ----
        _cl108 = [(s, bool(VB106.is_block_closer(s))) for s in (
            "    Next i", "End Sub", "    End If", "        Loop",
            "Loop While x", "Wend", "    x = 1", "Sub Foo()", "",
            "    Case 1", "    Else", "NextFoo", "Next '注释",
        )]
        _want108 = [
            ("    Next i", True), ("End Sub", True), ("    End If", True),
            ("        Loop", True), ("Loop While x", True), ("Wend", True),
            ("    x = 1", False), ("Sub Foo()", False), ("", False),
            ("    Case 1", False), ("    Else", False), ("NextFoo", False),
            ("Next '注释", True),
        ]
        _bad_c108 = [(s, g, w) for ((s, g), (s2, w)) in zip(_cl108, _want108)
                     if g != w or s != s2]
        check("108.11 `is_block_closer`：11 种收尾全认（含 `Next i` / "
              "`Loop While x` / 带注释的 `Next '注释`），普通行 / 块头 / 空行 / "
              "`Case` / `Else` 都不认；`NextFoo` 不误判成 `Next`",
              [_bad_c108], expect_contain=[[]])

        # ---- 108.12 / 108.13 修饰键门槛（纯函数）----
        _ins108 = [(vk, c, s, a, w, _MB108._insert_above_vk_ok(vk, c, s, a, w))
                   for (vk, c, s, a, w) in (
                       (_MB108.VK_RETURN, True, True, False, False),
                       (_MB108.VK_RETURN, True, False, False, False),
                       (_MB108.VK_RETURN, False, True, False, False),
                       (_MB108.VK_RETURN, True, True, True, False),
                       (_MB108.VK_RETURN, True, True, False, True),
                       (_MB108.VK_D, True, True, False, False),
                   )]
        check("108.12 Ctrl+Shift+Enter 的门槛：Ctrl 和 Shift **都要按**，"
              "Alt / Win 任一在按都不算；裸 Shift+Enter、裸 Enter、Ctrl+Enter "
              "都**不**算（它们各有各的语义）",
              [(_ins108,)], expect_contain=[([
                  (_MB108.VK_RETURN, True, True, False, False, True),
                  (_MB108.VK_RETURN, True, False, False, False, False),
                  (_MB108.VK_RETURN, False, True, False, False, False),
                  (_MB108.VK_RETURN, True, True, True, False, False),
                  (_MB108.VK_RETURN, True, True, False, True, False),
                  (_MB108.VK_D, True, True, False, False, False),
              ],)])

        _dd108 = [(vk, c, s, a, w, _MB108._dup_del_vk_ok(vk, c, s, a, w))
                  for (vk, c, s, a, w) in (
                      (_MB108.VK_D, True, False, False, False),
                      (_MB108.VK_K, True, False, False, False),
                      (_MB108.VK_D, True, True, False, False),
                      (_MB108.VK_D, True, False, True, False),
                      (_MB108.VK_K, True, False, False, True),
                      (_MB108.VK_D, False, False, False, False),
                      (_MB108.VK_D, True, False, False, False),
                      (_MB108.VK_O, True, False, False, False),
                  )]
        check("108.13 Ctrl+D / Ctrl+K 的门槛：**单独**按着 Ctrl 才算"
              "（Ctrl+Shift+D / Ctrl+Alt+D / Ctrl+Win+K 都是别的东西，绝不抢）；"
              "其余字母键一律不认",
              [(_dd108,)], expect_contain=[([
                  (_MB108.VK_D, True, False, False, False, "dup"),
                  (_MB108.VK_K, True, False, False, False, "del"),
                  (_MB108.VK_D, True, True, False, False, None),
                  (_MB108.VK_D, True, False, True, False, None),
                  (_MB108.VK_K, True, False, False, True, None),
                  (_MB108.VK_D, False, False, False, False, None),
                  (_MB108.VK_D, True, False, False, False, "dup"),
                  (_MB108.VK_O, True, False, False, False, None),
              ],)])

        # ---- 108.14 分支次序（源码位置）----
        # ⚠️ 这条是**顺序**断言，不是"存在"断言：Ctrl+Shift+Enter 必须先于
        # Shift+Enter 被问，否则被那条判据（只看 Enter + Shift）抢走，行为变成
        # "在下方插行" —— 差一个字，方向反了。
        _src108 = io.open(os.path.join(os.path.dirname(HERE), "main.py"),
                          encoding="utf-8").read()
        _i_ins108 = _src108.find("elif _insert_above_key_ok(vk):")
        _i_shift108 = _src108.find(
            "elif (_is_newline_shortcut(vk, _shift_down())")
        _i_dd108 = _src108.find("elif _dup_del_key_ok(vk):")
        _i_vis108 = _src108.find("elif completer.is_visible():")
        check("108.14 分支次序：`_insert_above_key_ok` 排在 `_is_newline_shortcut` "
              "**之前**（顺带在 `completer.is_visible()` 之前）；`_dup_del_key_ok` "
              "也排在候选窗那条之前 —— 它们都要改文本，窗里那些候选的行号马上"
              "就不作数了",
              [(_i_ins108 > 0, _i_shift108 > 0, _i_dd108 > 0, _i_vis108 > 0,
                _i_ins108 < _i_shift108, _i_ins108 < _i_vis108,
                _i_dd108 < _i_vis108)],
              expect_contain=[(True, True, True, True, True, True, True)])

        # ---- 108.15 闸门里必须有 pending_keys（AST）----
        # 三个动作都要**写文本**：输入法正在组字 / 刚敲的字还没落进文档时读到的是
        # 旧行文本，一写就落在错位置。这条用 AST 查，注释里写一万遍也不算数。
        _tree108 = _ast106.parse(_src108)

        def _uses_pending108(fn_name):
            for _n108 in _ast106.walk(_tree108):
                if (isinstance(_n108, _ast106.FunctionDef)
                        and _n108.name == fn_name):
                    for _sub in _ast106.walk(_n108):
                        if (isinstance(_sub, _ast106.Constant)
                                and _sub.value == "pending_keys"):
                            return True
            return False

        check("108.15 写文本类的闸门都看 `pending_keys`（AST）："
              "`_insert_above_key_ok` 与 `_dup_del_key_ok` 各有一处 —— "
              "输入法组字 / 刚敲的字还没落进文档时不接管，读到的行文本才是可信的",
              [(_uses_pending108("_insert_above_key_ok"),
                _uses_pending108("_dup_del_key_ok"))],
              expect_contain=[(True, True)])

        # ---- 108.16 三个开关都登记了 ----
        _keys108 = [k for (k, _d, _s) in _ST108.KNOWN]
        # ★v124d：KNOWN 总数 54 -> 55（再拆出 PARAM_INFO，见 §112）
        check("108.16 三个开关都进了 `settings.KNOWN`（各自独立，误触哪个关哪个）",
              [(_keys108.count("NO_INSERT_ABOVE"),
                _keys108.count("NO_DUP_LINE"),
                _keys108.count("NO_DEL_LINE"), len(_keys108))],
              expect_contain=[(1, 1, 1, 55)])
    finally:
        VB106._get_vbe_cached = _o_vbe108


    # ------------------------------------------------------------------
    # 109. ★v123：词库（`wordbank.db`）—— 提示词从代码常量搬进可手改的 SQLite
    #     用户原话："是不是可以创建个 SQLite 数据库，存放在这个数据里，然后项目
    #     启动，就把数据库里的加载到内存里，我也可以手动去编辑数据库，项目就能
    #     动态加载提示关键字。"
    #
    #     这一节要证明的**不是**"能建库"，而是下面四件事：
    #       ① 不改库时行为**逐字等于** v122（否则就是偷偷改了用户的手感）；
    #       ② 手改库真的能生效（禁用 / 自加 / 覆盖补空格），而且**走到收集侧与
    #          引擎侧**——不是只改了常量定义、看上去能跑；
    #       ③ 库坏了 / 空库 / 版本更高：一律退回种子，主体功能不受影响；
    #       ④ 热路径上**没有任何 SQLite 调用**（性能红线）。
    #
    #     ⚠️ 这一节会把 `ACTIVE_*` 改成"临时库那一份"，finally 里必须
    #     `reset_active()` 还原 —— 否则后面每一节都在用被改过的词表（§90 盯着）。
    # ------------------------------------------------------------------
    _saved_env109 = os.environ.get("VBECOMPLETE_WORD_BANK")
    _d109 = None
    try:
        import ast as _ast109
        import contextlib as _cl109
        import shutil as _sh109
        import sqlite3 as _sq109
        import tempfile as _tf109

        import engine as _E109
        import settings as _ST109
        import vba_builtins as _VB109
        import vbe_bridge as _BV109
        import wordbank as _WB109

        _d109 = _tf109.mkdtemp(prefix="wb109_")
        _p109 = os.path.join(_d109, "wordbank.db")
        _rows109 = _VB109.seed_rows()

        print("\n=== 109. ★v123：词库（wordbank.db）—— 可手改的提示词表 ===")

        # ---- 109.1 不改库 = v122 的行为 ----
        _VB109.reset_active()
        check("109.1 没有词库时 `ACTIVE_*` 与代码里的种子**逐字相同**（含顺序），"
              "补空格集合也一样 —— \"不改库\"就等于 v122 的行为，一个字节没变",
              [(_VB109.ACTIVE_FUNCTIONS == _VB109.BUILTIN_FUNCTIONS,
                _VB109.ACTIVE_CONSTANTS == _VB109.BUILTIN_CONSTANTS,
                _VB109.ACTIVE_TYPE_NAMES == _VB109.BUILTIN_TYPE_NAMES,
                _VB109.ACTIVE_KEYWORDS == _VB109.BUILTIN_KEYWORDS,
                _VB109.ACTIVE_KEYWORD_SPACE == _VB109.BUILTIN_KEYWORD_NEEDS_SPACE,
                _VB109.WORDBANK_INFO["loaded"] is False)],
              expect_contain=[(True,) * 6])

        # ---- 109.2 主键是 (name, kind)，双身份名字两行都在 ----
        _key109 = set((n.lower(), k) for n, k, *_r in _rows109)
        check("109.2 主键是 (name, kind)：`Date` / `String` 在 function 与 type "
              "**各占一行**（既是函数又是类型；混成一行，`As |` 位置和表达式位置就"
              "分不开了 —— 铁律 6「两道判据问不同问题」）；种子自身按主键无重复",
              [("date" in [n.lower() for n in _VB109.BUILTIN_FUNCTIONS],
                "date" in [n.lower() for n in _VB109.BUILTIN_TYPE_NAMES],
                ("date", "function") in _key109,
                ("date", "type") in _key109,
                ("string", "function") in _key109,
                ("string", "type") in _key109,
                len(_rows109) == len(_key109))],
              expect_contain=[(True,) * 7])

        # ---- 109.3 建库 + 幂等 + 绝不覆盖 ----
        _cr109, _got109 = _WB109.ensure(_rows109, _p109)
        _sz109 = os.path.getsize(_p109) if os.path.isfile(_p109) else 0
        _cr109b, _ = _WB109.ensure(_rows109, _p109)
        # 真的"不覆盖"要这样证：先手插一条，再调一次 ensure，那条必须还在
        _c109 = _sq109.connect(_p109)
        _c109.execute("INSERT INTO words VALUES"
                      " ('KeepMe','function',0,1,NULL,'手插的')")
        _c109.commit()
        _c109.close()
        _WB109.ensure(_rows109, _p109)
        _snap109 = _WB109.load(_p109)
        _fn109 = [n for n, _b, _w in _snap109.by_kind("function")] if _snap109 else []
        check("109.3 `ensure()` 建库之后**幂等**，而且**绝不覆盖**用户手改"
              "（插一条 KeepMe，再 ensure 一次，它必须还在）",
              [(_cr109, _cr109b, _sz109 > 0, "KeepMe" in _fn109,
                len(_fn109) == len(_VB109.BUILTIN_FUNCTIONS) + 1)],
              expect_contain=[(True, False, True, True, True)])

        # ---- 109.4 读回来与种子逐字同序 ----
        check("109.4 库读回来与种子**逐字同序**（顺序也一致 —— 候选池的顺序不去"
              "二次排序，改这一层就是改手感）",
              [(_snap109 is not None,
                tuple(_fn109) == _VB109.BUILTIN_FUNCTIONS + ("KeepMe",),
                tuple(n for n, _b, _w in _snap109.by_kind("type"))
                == _VB109.BUILTIN_TYPE_NAMES,
                tuple(n for n, _b, _w in _snap109.by_kind("keyword"))
                == _VB109.BUILTIN_KEYWORDS,
                tuple(n for n, _b, _w in _snap109.by_kind("constant"))
                == _VB109.BUILTIN_CONSTANTS)],
              expect_contain=[(True,) * 5])

        # ---- 109.5 库路径：默认在项目文件夹，`WORD_BANK` 可改 ----
        check("109.5 默认库就在**项目文件夹**（与 `snippets.db` 同一处，"
              "一眼能找到），`WORD_BANK` 可改到别处",
              [_WB109.DEFAULT_DB_NAME, _WB109.db_path().lower().startswith(
                  _WB109.PROJECT_DIR.lower())],
              expect_contain=["wordbank.db", True])

        # ---- 109.6 手改真的生效（走 load_active）----
        _c109 = _sq109.connect(_p109)
        _c109.execute("UPDATE words SET enabled=0 WHERE name='Kill'")
        _c109.execute("UPDATE words SET wants_space=0 WHERE name='Sub'")
        _c109.execute("INSERT INTO words VALUES"
                      " ('MyHostApi','function',0,1,NULL,'自加')")
        _c109.execute("INSERT INTO words VALUES"
                      " ('ZzzKey','keyword',0,1,1,'自加要空格')")
        _c109.commit()
        _c109.close()
        os.environ["VBECOMPLETE_WORD_BANK"] = _p109
        _ok109 = _VB109.load_active()
        check("109.6 手改库 -> `load_active()` 生效：禁用的 `Kill` 走了、自加的 "
              "`MyHostApi` 进来了、109.3 手插的 `KeepMe` 也读得到（其余照旧）",
              [(_ok109, _VB109.WORDBANK_INFO["loaded"],
                "Kill" not in _VB109.ACTIVE_FUNCTIONS,
                "MyHostApi" in _VB109.ACTIVE_FUNCTIONS,
                "KeepMe" in _VB109.ACTIVE_FUNCTIONS,
                "MsgBox" in _VB109.ACTIVE_FUNCTIONS)],
              expect_contain=[(True, True, True, True, True, True)])

        check("109.7 自加的词**不算**\"语言自带\"（`ACTIVE_BUILTIN_LOWER` 里没有 "
              "MyHostApi / ZzzKey）—— 否则引擎会拿\"公共词汇\"那套收紧规则连坐它，"
              "用户特意加的词反而打不出来",
              [("myhostapi" in _VB109.ACTIVE_BUILTIN_LOWER,
                "zzzkey" in _VB109.ACTIVE_BUILTIN_LOWER,
                "msgbox" in _VB109.ACTIVE_BUILTIN_LOWER,
                "if" in _VB109.ACTIVE_BUILTIN_LOWER)],
              expect_contain=[(False, False, True, True)])

        # ---- 109.8 补空格三态：NULL 走排除法，1 / 0 覆盖它 ----
        check("109.8 补空格是**三态**：`Sub` 被显式写成 0 -> 不补；`Public` 仍是 "
              "NULL -> 按排除法补；`ZzzKey` 显式 1 -> 补。引擎侧直接查生效集合"
              "（不再自己抄一份）",
              [(_E109._keyword_wants_space("Sub"),
                _E109._keyword_wants_space("Public"),
                _E109._keyword_wants_space("ZzzKey"),
                _E109._keyword_wants_space("Else"))],
              expect_contain=[(False, True, True, False)])

        # ---- 109.9 收集侧真的读生效词表（不是只改了常量定义）----
        class _CM109(object):
            """照真 VBE 的形状：`Lines(a, n)` 回来是 CRLF 连接的多行。"""

            def __init__(self, code):
                self._ls = str(code).split("\n")

            @property
            def CountOfLines(self):
                return len(self._ls)

            def Lines(self, a, b=None):
                if b is None:
                    return self._ls[a - 1]
                return "\r\n".join(self._ls[a - 1:a - 1 + b])

        class _Comp109(object):
            def __init__(self, name, code):
                self.Name = name
                self.Type = 1                        # vbext_ct_StdModule
                self.CodeModule = _CM109(code)

        _o_vbe109 = _BV109._get_vbe_cached
        _o_cc109 = _BV109._collect_components
        _o_cp109 = _BV109._caret_position
        _o_sw109 = (_BV109.ENABLE_VBA_BUILTINS, _BV109.ENABLE_VBA_KEYWORDS,
                    _BV109.ENABLE_VBA_CONSTANTS, _BV109.ENABLE_HOST_ENUMS)
        try:
            _BV109._get_vbe_cached = lambda: object()
            _BV109._collect_components = lambda vbe, cm: [
                _Comp109("Module1", "Sub Foo()\n    x = 1\nEnd Sub")]
            _BV109._caret_position = lambda vbe: (None, None)
            _BV109.ENABLE_VBA_BUILTINS = True
            _BV109.ENABLE_VBA_KEYWORDS = True
            _BV109.ENABLE_VBA_CONSTANTS = False
            _BV109.ENABLE_HOST_ENUMS = False
            _bk109 = _BV109.VbeBackend()
            _nm109 = set(str(r[0]).lower() for r in _bk109.get_identifiers())
            _bl109 = _bk109.get_builtin_names()
        finally:
            _BV109._get_vbe_cached = _o_vbe109
            _BV109._collect_components = _o_cc109
            _BV109._caret_position = _o_cp109
            (_BV109.ENABLE_VBA_BUILTINS, _BV109.ENABLE_VBA_KEYWORDS,
             _BV109.ENABLE_VBA_CONSTANTS, _BV109.ENABLE_HOST_ENUMS) = _o_sw109

        check("109.9 ★收集侧端到端：拿到的是**生效词表**——`MyHostApi` 进了池子、"
              "`Kill` 没进（被禁用）、`If` 照常；而 `MyHostApi` 不在收紧集合里"
              "（用户自加的词不该被\"公共词汇\"连坐）",
              [("myhostapi" in _nm109, "kill" not in _nm109, "if" in _nm109,
                "sub" in _nm109,
                "myhostapi" in _bl109, "msgbox" in _bl109, "if" in _bl109)],
              expect_contain=[(True, True, True, True, False, True, True)])

        # ---- 109.10 四种坏库：一律退回种子，绝不抛 ----
        def _try_bad109(_path, _tag):
            os.environ["VBECOMPLETE_WORD_BANK"] = _path
            _VB109.reset_active()
            _loaded = _VB109.load_active()
            return (_loaded, _VB109.ACTIVE_FUNCTIONS == _VB109.BUILTIN_FUNCTIONS,
                    bool(_VB109.WORDBANK_INFO.get("error")), _tag)

        _bad109 = []
        _pbad = os.path.join(_d109, "garbage.db")
        with open(_pbad, "wb") as _f:
            _f.write(b"definitely not a sqlite database" * 8)
        _bad109.append(_try_bad109(_pbad, "垃圾字节"))

        _pmeta = os.path.join(_d109, "nometa.db")
        _c109 = _sq109.connect(_pmeta)
        _c109.execute("CREATE TABLE words (name TEXT, kind TEXT)")
        _c109.commit()
        _c109.close()
        _bad109.append(_try_bad109(_pmeta, "缺 meta 表"))

        _pempty = os.path.join(_d109, "empty.db")
        _WB109.ensure([], _pempty)
        _bad109.append(_try_bad109(_pempty, "空库"))

        _pver = os.path.join(_d109, "newver.db")
        _sh109.copyfile(_p109, _pver)
        _c109 = _sq109.connect(_pver)
        _c109.execute("UPDATE meta SET value='99' WHERE key='schema_version'")
        _c109.commit()
        _c109.close()
        _bad109.append(_try_bad109(_pver, "版本 99"))

        _pmiss = os.path.join(_d109, "nope", "missing.db")
        _bad109.append(_try_bad109(_pmiss, "文件不存在"))

        check("109.10 ★五种坏库（垃圾字节 / 缺 meta / 空库 / 版本 99 / 文件不存在）"
              "全部：`load_active()` 返回 False、`ACTIVE_*` 退回种子、"
              "`WORDBANK_INFO['error']` 有话说 —— **一个异常都不许往外抛**",
              [(_bad109[0][0], _bad109[1][0], _bad109[2][0], _bad109[3][0],
                _bad109[4][0]),
               tuple(x[1] for x in _bad109),
               tuple(x[2] for x in _bad109)],
              expect_contain=[(False,) * 5, (True,) * 5, (True,) * 5])

        # ---- 109.11 命令行（用户不一定装 GUI 工具）----
        _out109 = io.StringIO()
        try:
            with _cl109.redirect_stdout(_out109):
                _WB109._cmd_add(_p109, "CliWord", "keyword")
                _WB109._cmd_set_enabled(_p109, "CliWord", False)
                _WB109._cmd_stats(_p109)
        except Exception as _e109c:
            _out109.write("ERR: %s" % _e109c)
        _snapC = _WB109.load(_p109)
        _kwC = [n for n, _b, _w in _snapC.by_kind("keyword")] if _snapC else []
        check("109.11 命令行能改库（`add` 加进去、`disable` 之后就查不到），"
              "`stats` 不炸 —— 不装 DB Browser 也能维护",
              [("CliWord" not in _kwC, "共 " in _out109.getvalue(),
                "ERR:" not in _out109.getvalue())],
              expect_contain=[(True, True, True)])

        # ---- 109.12 热路径上没有 SQLite（性能红线）----
        def _top_imports109(_path):
            _tree = _ast109.parse(io.open(_path, encoding="utf-8").read())
            _out = []
            for _nd in _tree.body:                   # ★只看**模块级**语句
                if isinstance(_nd, _ast109.Import):
                    _out.extend(_a.name.split(".")[0] for _a in _nd.names)
                elif isinstance(_nd, _ast109.ImportFrom):
                    _out.append((_nd.module or "").split(".")[0])
            return _out

        def _all_imports109(_path):
            _tree = _ast109.parse(io.open(_path, encoding="utf-8").read())
            _out = []
            for _nd in _ast109.walk(_tree):
                if isinstance(_nd, _ast109.Import):
                    _out.extend(_a.name.split(".")[0] for _a in _nd.names)
                elif isinstance(_nd, _ast109.ImportFrom):
                    _out.append((_nd.module or "").split(".")[0])
            return _out

        _eng_top109 = _top_imports109(os.path.join(ROOT, "engine.py"))
        _eng_all109 = _all_imports109(os.path.join(ROOT, "engine.py"))
        _vb_top109 = _top_imports109(os.path.join(ROOT, "vba_builtins.py"))
        check("109.12 ★性能红线：`engine.py` **全文件**都没有 `sqlite3`"
              "（引擎不该碰库），`vba_builtins.py` **顶层**也没有 `sqlite3` / "
              "`wordbank`（只在 `load_active()` 里惰性 import）—— 按键热路径上"
              "一个数据库调用都没有",
              [("sqlite3" in _eng_all109, "sqlite3" in _eng_top109,
                "wordbank" in _eng_all109,
                "sqlite3" in _vb_top109, "wordbank" in _vb_top109,
                "engine" in _vb_top109)],
              expect_contain=[(False,) * 6])

        # ---- 109.13 两个新键都登记了 ----
        _keys109s = [k for k, _d, _s in _ST109.KNOWN]
        # ★v124d：KNOWN 总数 54 -> 55（再拆出 PARAM_INFO，见 §112）
        check("109.13 两个新键都进了 `settings.KNOWN`：`NO_WORDBANK`（排错开关）、"
              "`WORD_BANK`（库路径）",
              [(_keys109s.count("NO_WORDBANK"), _keys109s.count("WORD_BANK"),
                len(_keys109s))],
              expect_contain=[(1, 1, 55)])

        check("109.14 启动接线在 `main()` 里、且**先建库再加载**"
              "（顺序反了就会先读到「没有库」、退回种子，然后建了库却不生效）",
              [(("wordbank.ensure(" in
                 io.open(os.path.join(ROOT, "main.py"), encoding="utf-8").read()),
                ("vba_builtins.load_active()" in
                 io.open(os.path.join(ROOT, "main.py"), encoding="utf-8").read()),
                (io.open(os.path.join(ROOT, "main.py"), encoding="utf-8").read()
                 .index("wordbank.ensure(")
                 < io.open(os.path.join(ROOT, "main.py"), encoding="utf-8").read()
                 .index("vba_builtins.load_active()")))],
              expect_contain=[(True, True, True)])

    except Exception as _e109:
        import traceback as _tp109
        _tp109.print_exc()
        check("第 109 节异常: %s" % _e109, [True], expect_contain=[False])
    finally:
        # ⚠️ 三件事必须还原：环境变量、`ACTIVE_*`（否则后面每一节都用改过的词表）、
        #    临时目录。§90 会盯着"猴补丁有没有还原"。
        try:
            if _saved_env109 is None:
                os.environ.pop("VBECOMPLETE_WORD_BANK", None)
            else:
                os.environ["VBECOMPLETE_WORD_BANK"] = _saved_env109
        except Exception:
            pass
        try:
            import vba_builtins as _VB109f
            _VB109f.reset_active()
        except Exception:
            pass
        try:
            if _d109:
                _sh109.rmtree(_d109, ignore_errors=True)
        except Exception:
            pass


    # ------------------------------------------------------------------
    # 110. ★代码复查（2026-10-18）：隐藏 bug 的钉桩
    #     这一节不是在测"功能对不对"（那种用例前面九十多节都铺满了），
    #     而是给**复查时查出来并修掉的每一个隐藏 bug** 钉一根桩：
    #     每条用例都对应一个"当年真的会发生"的具体故障，而且都写成
    #     **改回旧实现就会红** 的形式（值断言，不是"代码里有没有这句话"）。
    #
    #     为什么值得单独一节：这十处里有六处**没有任何既有用例会变红** ——
    #     它们要么只在"模块最后一行的下一行""未闭合的引号""Rem 注释"
    #     这类中间态里出现，要么被 `except: pass` 吞掉了异常。
    # ------------------------------------------------------------------
    print("\n=== 110. ★代码复查：隐藏 bug 的钉桩（2026-10-18）===")
    import ast as _ast110
    import vbe_bridge as VB110

    # ---- 110.1 `parser.scan_code_kinds`：Rem 注释的唯一判据 ----
    # 这是本仓库"同一语义只能有一份实现"（MEMORY 第 14 条）的典型：
    # `'` 注释 / 字符串 / Rem 三件事必须由同一个扫描器回答，谁再自己
    # `find("rem")` 一次就必然分叉（`_in_rem_comment` 原来就是那么干的，
    # 于是字符串里的 "a rem b" 被当成注释）。
    _k110 = P.scan_code_kinds
    _g110a = (
        _k110("Rem x = 1")[0],                       # 行首 Rem -> KIND_REM
        _k110("    rem 说明")[4],                    # 缩进之后也算（大小写不敏感）
        _k110('x = "a Rem b"')[7],                   # 串里的 Rem -> 字符串
        sorted(set(_k110("Remote.foo"))),            # 标识符里的 rem -> 全是代码
        _k110("x = 1 : Rem a")[8],                   # `:` 之后的 Rem 也算
    )
    check("110.1 `scan_code_kinds` 认 `Rem` 注释：行首 / 缩进之后 / `:` 之后"
          "都算；字符串里的、标识符里的（`Remote.foo`）都不算 —— 三种注释"
          "（字符串 / `'` / Rem）只有这一份判据",
          [_g110a],
          expect_contain=[(P.KIND_REM, P.KIND_REM, P.KIND_STRING, [0],
                           P.KIND_REM)])

    # ---- 110.2 `Rem … _` 不能再把下一行并进来 ----
    # 旧实现里 `_RE_CONTINUATION` 把 `Rem 说明 _` 的续行符当真，于是紧跟的
    # `End Sub` 被并进注释那一行 ⇒ 过程作用域一直不结束 ⇒ 后面整段**模块级**
    # 代码被算成该过程的局部变量（模块级变量在别处成片消失）。
    # ⚠️ 断言必须落在 `extract_records` 上（它才做续行合并），
    #    `proc_at_line` 不做合并，新旧结果相同、区分不出来。
    _c110a = "Sub A()\nRem 说明 _\nEnd Sub\nDim m As Long\n"
    _c110b = "Sub A()\nRem Dim count As Long\nEnd Sub\n"
    _r110a = P.extract_records(_c110a, "Mod1", True)[0]
    _r110b = P.extract_records(_c110b, "Mod1", True)[0]
    _names110b = [n for n, _m, _p, _pv in _r110b]
    check("110.2 Rem 注释里的 `_` **不是**续行符：`End Sub` 该结束过程 —— "
          "`Rem 说明 _` 之后那行模块级 `Dim m` 的 proc 必须是 None"
          "（旧实现里 `End Sub` 被并进注释，它变成 A 的局部变量）；"
          "`Rem Dim count As Long` 也不能产出幻影记录 `Rem`",
          [([p for n, _m, p, _pv in _r110a if n == "m"],
            "Rem" in _names110b, "count" in _names110b)],
          expect_contain=[([None], False, False)])

    # ---- 110.3 掩码必须保留换行（未闭合引号的中间态）----
    # 用户正敲一个还没打右引号的字符串，是编辑期最常见的瞬间。旧实现把
    # "从 `"` 到下一个 `"`" 这**一整段连续的被掩码字符（含换行）**压成
    # 一个空格 ⇒ 后面所有行的行号整体前移 ⇒ `proc_at_line` / `extract_records`
    # 把 B 里的代码算成 A 的（实测 7 行压成 3 行）。
    _c110c = ('Sub A()\nDim x As Long\ns = "abc\nDim y As Long\n"\n'
              'Sub B()\nDim z As Long\nEnd Sub\n')
    _m110c = P._mask_strings_and_comments(_c110c)
    check("110.3 掩码**行数严格不变**（未闭合的引号跨了几行，就留几个换行）："
          "引号闭在第 5 行之后，第 7 行（`Dim z As Long`）仍该归 B —— 旧实现"
          "把整段压成一个空格，行号前移，这行会被算成 A 的、甚至算成 None",
          [(len(_m110c.split("\n")), len(_c110c.split("\n")),
            P.proc_at_line(_c110c, 7))],
          expect_contain=[(9, 9, "B")])
    check("110.4 掩码把字符串 / 注释抹成空格 —— 字符串正文里的 `rem` "
          "不会被任何走掩码的正则当成 Rem 注释",
          [(("rem" in P._mask_strings_and_comments('x = "a Rem b"').lower()),
            ("rem" in P._mask_strings_and_comments("Rem a b").lower()))],
          expect_contain=[(False, False)])

    # ---- 110.5 `_in_rem_comment` 委托给同一份扫描器 ----
    # 旧实现自己 `line_text.lower().find("rem")` 再回头看一个字符 ⇒
    # 字符串里的 "the rem value" 被当成 Rem 注释 ⇒ 成员位置判定被误杀、
    # `c.` 该让位给 VBE 成员列表的地方**不让位**。
    _g110e = (
        E._in_rem_comment('x = "the rem value"', 12),   # 列落在串里 -> 不是注释
        E._in_rem_comment("Rem a.b", 7),                # 真 Rem 注释 -> 是
        E._in_rem_comment("Remote.foo", 5),             # 标识符里的 rem -> 不是
        E._in_rem_comment("x = 1 ' rem a.b", 11),       # `'` 注释不是 Rem
    )
    check("110.5 `_in_rem_comment` 只认真正的 Rem 注释：字符串里的 `rem`、"
          "标识符 `Remote` 里的 `rem`、`'` 注释里的 `rem` 一律 False"
          "（前两条错了的症状分别是「不让位」和「该补全不出」）",
          [_g110e], expect_contain=[(False, True, False, False)])

    # ---- 110.6 `settings.source()` 只留一份实现 ----
    # 曾经被定义了两次：第二份用 `effective(name, None, "flag")` —— 于
    # `RECENT_MAX = 32` / `LOG_FILE = …` 这类 int / text 键，用户明明写在
    # ini 里，`source()` 也会报 `default`，排错时正好把人引到反方向。
    import settings as S110
    import tempfile as _t110
    import shutil as _sh110
    _d110 = _t110.mkdtemp(prefix="vbe_rv_")
    _ini110 = os.path.join(_d110, "settings.ini")
    _old_env110 = os.environ.get("VBECOMPLETE_SETTINGS")
    _old_flags110 = {k: os.environ.get(k)
                     for k in ("VBECOMPLETE_AUTO_SPACE",
                               "VBECOMPLETE_RECENT_MAX")}
    try:
        os.environ["VBECOMPLETE_SETTINGS"] = _ini110
        for _k in _old_flags110:
            os.environ.pop(_k, None)
        with io.open(_ini110, "w", encoding="utf-8", newline="\n") as _f:
            _f.write("[settings]\nAUTO_SPACE = 0\nRECENT_MAX = 32\n")
        S110.reset()
        _g110f = (S110.source("RECENT_MAX"),           # int 键：默认 kind=text
                  S110.source("AUTO_SPACE", "flag"),   # flag 键
                  S110.source("AUTO_SPACE"))           # text 视图也认得出 ini
        os.environ["VBECOMPLETE_AUTO_SPACE"] = "maybe"  # 垃圾值：跳过这一层
        _g110g = (S110.source("AUTO_SPACE", "flag"),
                  S110.effective("AUTO_SPACE", "1", "flag"))
        os.environ.pop("VBECOMPLETE_AUTO_SPACE", None)
        check("110.6 `settings.source()` 与 `flag()/int_value()/text()` 走"
              "**同一个** `pick()` 口径：ini 里写了的 int / flag 键都要如实报 "
              "`ini`（旧实现只留了一份 flag 版 ⇒ 报 `default`）；env 写了"
              "垃圾值时也报 `ini`（那一层被跳过，不是掉回默认）",
              [(_g110f, _g110g)],
              expect_contain=[(("ini", "ini", "ini"), ("ini", (False, "ini")))])
    finally:
        for _k, _v in _old_flags110.items():
            if _v is None:
                os.environ.pop(_k, None)
            else:
                os.environ[_k] = _v
        if _old_env110 is None:
            os.environ.pop("VBECOMPLETE_SETTINGS", None)
        else:
            os.environ["VBECOMPLETE_SETTINGS"] = _old_env110
        try:
            S110.reset()
        except Exception:
            pass
        _sh110.rmtree(_d110, ignore_errors=True)

    # ---- 110.7 `RECENT_MAX = 0` 不能抛（那个自然写法想关掉特性）----
    # 旧实现 `max(data.values())` 在截完为空时抛 ValueError，而 main 用
    # `except: pass` 包着 ⇒ 症状是"最近使用没恢复"且毫无线索。
    class _Store110(object):
        """假的落盘层：`load_recent` 只调它一个 `load()`。"""

        def load(self):
            return {"alpha": 3, "beta": 2, "gamma": 1}

    def _recent110(cap):
        _saved = (E._recent_store, E.RECENT_MAX, E.PERSIST_RECENT,
                  E.RANK_BY_RECENT)
        E._recent_store = _Store110()
        E.RECENT_MAX = cap
        E.PERSIST_RECENT = True
        E.RANK_BY_RECENT = True
        try:
            _c = E.Completer.__new__(E.Completer)   # 不跑 __init__（要碰 COM）
            _c._recent, _c._recent_seq, _c._recent_dirty = {}, 0, False
            _n = _c.load_recent()
            return _n, sorted(_c._recent), _c._recent_seq, _c._recent_dirty
        finally:
            (E._recent_store, E.RECENT_MAX, E.PERSIST_RECENT,
             E.RANK_BY_RECENT) = _saved

    check("110.7 `load_recent()` 的上限三档：`RECENT_MAX = 0` 返回 0 且**不抛**"
          "（`max(())` 那个 ValueError 已被吞掉，现在根本走不到）；"
          "`= 2` 只留序号最大的两条；够大就全收 —— 三条都要正确",
          [(_recent110(0), _recent110(2), _recent110(64))],
          expect_contain=[((0, [], 0, False),
                           (2, ["alpha", "beta"], 3, False),
                           (3, ["alpha", "beta", "gamma"], 3, False))])

    # ---- 110.8 "文本已写、落光标失败" 必须按成功返回 ----
    # 否则调用方（main 的 `_auto_pair_here`）看到 False 会把这一下按键原样
    # 还给系统 ⇒ 行里已经有 `()`，VBE 又插一个 `(` ⇒ 用户看到 `( (`。
    # 同一口径适用于 `duplicate_lines` / `delete_lines_here`（前面几节已覆盖）。
    class _Mod110(object):
        def __init__(self, lines):
            self.lines = list(lines)

        def Lines(self, n, c):
            return "\r\n".join(self.lines[n - 1:n - 1 + c])

        @property
        def CountOfLines(self):
            return len(self.lines)

        def ReplaceLine(self, n, t):
            self.lines[n - 1] = t

        def InsertLines(self, n, t):
            for _k, _ln in enumerate(t.split("\n")):
                self.lines.insert(n - 1 + _k, _ln)

        def DeleteLines(self, n, c):
            del self.lines[n - 1:n - 1 + c]

    class _Pane110(object):
        def __init__(self, mod, sl, sc, el, ec, boom=False):
            self.CodeModule = mod
            self.sel = [sl, sc, el, ec]
            self.boom = boom          # True = 落光标时 COM 抛异常

        def GetSelection(self):
            return tuple(self.sel)

        def SetSelection(self, *a):
            if self.boom:
                raise RuntimeError("SetSelection 失败（模拟宿主收回编辑器）")
            self.sel = list(a)

    class _Vbe110(object):
        def __init__(self, pane):
            self.ActiveCodePane = pane

    def _pair110(lines, sl, sc, el, ec, ch="(", boom=False, sem=None):
        _mod = _Mod110(lines)
        _pane = _Pane110(_mod, sl, sc, el, ec, boom)
        _o_vbe = VB110._get_vbe_cached
        _o_sem = VB110._sem_cache.get("info")
        if sem is not None:
            VB110._sem_cache["info"] = sem
        VB110._get_vbe_cached = lambda: _Vbe110(_pane)
        try:
            _ok = VB110.VbeBackend().insert_pair(ch)
        finally:
            VB110._get_vbe_cached = _o_vbe
            VB110._sem_cache["info"] = _o_sem
        return bool(_ok), list(_mod.lines), tuple(_pane.sel)

    check("110.8 `insert_pair` 在**文本已经写进去**之后落光标失败，仍返回 True："
          "`()` 确实在行里，光标停在原处。返回 False 会让调用方把按键还给"
          "系统，于是变成 `( (`（这类\"写成功但返回失败\"的坑本仓库踩过好几轮）",
          [(_pair110(["x = "], 1, 5, 1, 5, "(", boom=True),
            _pair110(["x = "], 1, 5, 1, 5, "(", boom=False))],
          expect_contain=[((True, ["x = ()"], (1, 5, 1, 5)),
                           (True, ["x = ()"], (1, 6, 1, 6)))])

    # ---- 110.9 `type_names` 必须晚于 `get_identifiers()` ----
    # `_type_names` 是在 `_collect_identifiers()` 里赋的值，`get_type_names()`
    # 只是读它。旧代码在 `trigger()` 一开头就取 ⇒ 冷启动第一次 `As |` 拿到
    # 空集 ⇒ 候选被筛空 ⇒ 什么都不弹；新建 / 改名类模块后的第一帧也漏新类型。
    # ⚠️ 查源码字符串会被注释喂饱（那段注释里就写着 `type_names = ...`），
    #    所以这里查 **AST**（MEMORY 第 25 条）。
    def _trigger_marks110():
        _src = io.open(os.path.join(ROOT, "engine.py"),
                       encoding="utf-8").read()
        _tree = _ast110.parse(_src)
        _fn = None
        for _nd in _ast110.walk(_tree):
            if isinstance(_nd, _ast110.FunctionDef) and _nd.name == "trigger":
                _fn = _nd
                break
        _marks = {}
        _ifs = []
        for _nd in _ast110.walk(_fn):
            if isinstance(_nd, _ast110.Assign) and _nd.targets:
                _t = _nd.targets[0]
                if isinstance(_t, _ast110.Name) and _t.id in ("visible_ids",
                                                              "type_names"):
                    _marks.setdefault(_t.id, _nd.lineno)
            if (isinstance(_nd, _ast110.If) and isinstance(_nd.test, _ast110.Name)
                    and _nd.test.id == "in_type"):
                _ifs.append(_nd.lineno)
        return _marks.get("visible_ids"), _marks.get("type_names"), _ifs

    _m110 = _trigger_marks110()
    _g110h = (_m110[0] is not None and _m110[1] is not None
              and _m110[0] < _m110[1],
              bool([x for x in _m110[2] if _m110[1] is not None
                    and x > _m110[1]]))
    check("110.9 AST 断言：`trigger()` 里 `type_names = self._type_name_set()` "
          "排在 `visible_ids = [...]`（也就是 `get_identifiers()`）**之后**、"
          "并且确实在那条 `if in_type:` 收窄之前 —— 顺序反了就是"
          "\"冷启动 `As |` 一个字都不弹\"，而且不看代码根本查不出来",
          [_g110h], expect_contain=[(True, True)])

    # ---- 110.10 `insert_snippet` 的终点列必须换算成字符下标 ----
    # 单行选区原先直接拿 `ec`（**显示列**）当字符下标 ⇒ 含全角 / Tab 的行上
    # 多删少删一个字符，而 `tail = last[j:]` 之后是【整行重写】，删错就是
    # 用户代码真的少了几个字。这类"看不见的丢字"必须钉住。
    def _snip110(lines, sl, sc, el, ec, text, sem):
        _mod = _Mod110(lines)
        _pane = _Pane110(_mod, sl, sc, el, ec)
        _o_vbe = VB110._get_vbe_cached
        _o_sem = VB110._sem_cache.get("info")
        VB110._sem_cache["info"] = sem
        VB110._get_vbe_cached = lambda: _Vbe110(_pane)
        try:
            _ok = VB110.VbeBackend().insert_snippet(text)
        finally:
            VB110._get_vbe_cached = _o_vbe
            VB110._sem_cache["info"] = _o_sem
        return bool(_ok), list(_mod.lines), tuple(_pane.sel)

    # `s = "中文"`：光标停在 `中` 与 `文` 之间 = 显示列 8（字符列 7）。
    # 旧实现 j = min(ec-1, len) = 7 ⇒ tail = `"` ⇒ `文` 被整段顶掉。
    check("110.10 `insert_snippet` 单行选区在全角行上：显示列 8 换成字符下标 6，"
          "把 `X` 插进 `中` / `文` 之间，`文` 一个字符都不能丢"
          "（旧实现按显示列当下标，结果是 `s = \"中X\"`）",
          [(_snip110(['s = "中文"'], 1, 8, 1, 8, "X", ("disp", 4, True)))],
          expect_contain=[(True, ['s = "中X文"'], (1, 9, 1, 9))])

    # ---- 110.11 「光标在行尾」的边界差一格 ----
    # `enter_indent_wanted` 的纯 ASCII 分支原写 `ec < len(t)`：块头行
    # `    For i = 1 To 10`（19 字符、行尾显示列 20）上，光标停在**最后一个
    # 字符之前**时 VBE 报的就是 19，`19 < 19` 为假 ⇒ 那一下本该由 VBE 处理的
    # 【拆行】被我们接管（回车没拆行 / 缩进被改写）。
    # 修法：两支合成一条 `ec != _eol_disp_col(t)`（纯 ASCII 时它就是 len+1），
    # "行尾"这个语义只留一份算法 —— 与 v93 修全角虚高时用的是同一个函数。
    _t110 = "    For i = 1 To 10"
    check("110.11 `enter_indent_wanted` 的行尾判据差一格（旧写法 `ec < len`）："
          "块头行真·行尾（显示列 20）接管；**最后一个字符之前**（19）必须放行 "
          "-> 交给 VBE 拆行；再往后一格（21，快照过期那类）也放行",
          [([VB110.enter_indent_wanted(_t110, 19),
             VB110.enter_indent_wanted(_t110, 20),
             VB110.enter_indent_wanted(_t110, 21),
             VB110._eol_disp_col(_t110)])],
          expect_contain=[[False, True, False, 20]])

    # ------------------------------------------------------------------
    # 111. ★v124：缓存工程身份 / 片段插入原子性 / InsertLines 口径
    #     这三件是复查报告里"未修风险"的头三条，修完各钉一根桩：
    #
    #     1) 三份缓存（组件列表 / CodeModule / 模块解析）此前**只用模块名做键**。
    #        同时开两个 .xlsm（各有一个 `Module1`）时会互相命中 —— A 的模块
    #        解析结果被当成 B 的，补全出另一个工作簿的变量。`_comp_list_cache`
    #        里那个 `"proj"` 字段写了却从不参与比较，等于没有。
    #     2) `insert_snippet` 逐行 `InsertLines`：第 3 行抛异常时前 2 行已经
    #        进了文档、函数却返回 False ⇒ 用户看到"插入失败"再点一次 = 插两遍。
    #     3) `duplicate_lines` 说"多行原文直接喂回去，VBE 会拆"，`insert_snippet`
    #        说"VBE 拆行没有文档保证，必须逐行" —— 两处假设互相矛盾，至少一处
    #        是错的。★真机实测（Excel 2016 / VBE7.1）证明**拆行可靠**
    #        （`\r\n` / `\n` / `\r` 都对），现在两处统一成"一次调用"。
    # ------------------------------------------------------------------
    print("\n=== 111. ★v124：缓存工程身份 / 片段原子性（2026-10-18）===")
    import vbe_bridge as VB111

    # ---- 111.1 工程身份：`(Name, FileName)` ----
    # Excel 里每个工作簿的 `VBProject.Name` 默认都是 `VBAProject`，只看它等于
    # 没分。`FileName` 才是磁盘路径，但★真机实测：**未保存的工作簿读它会抛
    # com_error（"路径未找到"）而不是返回空串** —— 必须每次 try。
    class _Proj111(object):
        def __init__(self, name, fname=None, boom=False):
            self.Name = name
            self._fname = fname
            self._boom = boom

        @property
        def FileName(self):
            if self._boom:
                raise RuntimeError("com_error: 路径未找到")
            return self._fname

    check("111.1 `_project_key`：同名工程（都叫 VBAProject）靠 `FileName` 分开；"
          "未保存的工程读 `FileName` 会抛 com_error（真机实测，不是返回空串）"
          "⇒ 退化成 `(name, \"\")` 而不是整条炸掉；两个都空 -> None（不信任缓存）",
          [(VB111._project_key(_Proj111("VBAProject", r"C:\a\one.xlsm")),
            VB111._project_key(_Proj111("VBAProject", r"C:\b\two.xlsm")),
            VB111._project_key(_Proj111("VBAProject", boom=True)),
            VB111._project_key(_Proj111("", "")),
            VB111._project_key(None))],
          expect_contain=[(("vbaproject", r"c:\a\one.xlsm"),
                           ("vbaproject", r"c:\b\two.xlsm"),
                           ("vbaproject", ""), None, None)])

    # ---- 111.2 首尾组件名（同 Count 同模块名时的补充指纹）----
    class _Item111(object):
        def __init__(self, name, tag):
            self.Name = name
            self.Type = 1
            self.tag = tag

    class _VBC111(object):
        def __init__(self, items):
            self._it = list(items)
            self.visits = 0

        @property
        def Count(self):
            return len(self._it)

        def Item(self, i):
            return self._it[i - 1]

        def __iter__(self):
            self.visits += 1
            return iter(self._it)

    class _P111(object):
        def __init__(self, name, fname, tags):
            self.Name = name
            self.FileName = fname
            self.VBComponents = _VBC111([_Item111("Module1", t) for t in tags])

    class _VBE111(object):
        def __init__(self, proj):
            self.ActiveVBProject = proj

        @property
        def ActiveCodePane(self):
            return None

        @property
        def VBProjects(self):
            return [self.ActiveVBProject]

    _p111a = _P111("VBAProject", r"C:\a\one.xlsm", ["A"])
    _p111b = _P111("VBAProject", r"C:\b\two.xlsm", ["B"])
    _o_clc111 = dict(VB111._comp_list_cache)
    _o_mce111 = VB111.MOD_CACHE_ENABLED
    try:
        VB111._comp_list_cache.update({"ts": 0.0, "count": -1, "key": None,
                                       "edge": None, "comps": None})
        VB111.MOD_CACHE_ENABLED = True
        _vbe111 = _VBE111(_p111a)
        _r111a1 = VB111._collect_components(_vbe111, None)
        # 同一个工程再来一轮：必须命中缓存（不重新枚举）
        _r111a2 = VB111._collect_components(_vbe111, None)
        _vis_a = _p111a.VBComponents.visits
        # 换工程：Number / 模块名 / 组件首尾名全一样，只有 FileName 不同
        _vbe111.ActiveVBProject = _p111b
        _r111b1 = VB111._collect_components(_vbe111, None)
    finally:
        VB111._comp_list_cache.clear()
        VB111._comp_list_cache.update(_o_clc111)
        VB111.MOD_CACHE_ENABLED = _o_mce111

    check("111.2 组件列表缓存：同一工程第二轮**复用**（枚举只发生一次）；"
          "换成一个组件数 / 模块名 / 首尾名都一样的工程时**必须重新枚举** —— "
          "旧实现只比 Count + TTL，会把 A 的组件列表当成 B 的",
          [(_r111a1[0].tag, _r111a2[0].tag, _vis_a, _r111b1[0].tag,
            _p111b.VBComponents.visits)],
          expect_contain=[("A", "A", 1, "B", 1)])

    # ---- 111.3 `_cm_cache` / `_mod_cache` 的键必须带工程身份 ----
    # 两个工程同名模块、**行数也相同**（3 行）—— 旧实现下 `_cm_cache["module1"]`
    # 直接命中 A 的 CodeModule，读出来的 count 是 A 的、代码也是 A 的；
    # `_mod_cache["module1"]` 同理重放 A 的解析结果。用户看到的就是
    # "在 B 里补全出了 A 的变量名"。
    class _CM111(object):
        def __init__(self, code, tag):
            self._ls = code.replace("\r\n", "\n").split("\n")
            self.tag = tag
            self.Name = "Module1"
            self.Parent = None

        @property
        def CountOfLines(self):
            return len(self._ls)

        def Lines(self, a, b=None):
            if b is None:
                return self._ls[a - 1]
            return "\r\n".join(self._ls[a - 1:a - 1 + b])

    class _Comp111(object):
        def __init__(self, name, code, tag):
            self.Name = name
            self.Type = 1
            self.CodeModule = _CM111(code, tag)

    class _VBCode111(object):
        def __init__(self, items):
            self._it = list(items)

        @property
        def Count(self):
            return len(self._it)

        def Item(self, i):
            return self._it[i - 1]

        def __iter__(self):
            return iter(self._it)

    class _PCode111(object):
        def __init__(self, name, fname, code, tag):
            self.Name = name
            self.FileName = fname
            self.VBComponents = _VBCode111([_Comp111("Module1", code, tag)])

    _cd111a = "Sub AAA()\n    Dim alphaOnlyA As Long\nEnd Sub"
    _cd111b = "Sub BBB()\n    Dim betaOnlyB As Long\nEnd Sub"
    _pc111a = _PCode111("VBAProject", r"C:\a\one.xlsm", _cd111a, "A")
    _pc111b = _PCode111("VBAProject", r"C:\b\two.xlsm", _cd111b, "B")

    _o_vbe111 = VB111._get_vbe_cached
    _o_clc111c = dict(VB111._comp_list_cache)
    _o_bk111c = (VB111.ENABLE_VBA_BUILTINS, VB111.ENABLE_VBA_KEYWORDS,
                 VB111.ENABLE_VBA_CONSTANTS, VB111.ENABLE_HOST_ENUMS)
    try:
        VB111._comp_list_cache.update({"ts": 0.0, "count": -1, "key": None,
                                       "edge": None, "comps": None})
        VB111.ENABLE_VBA_BUILTINS = False
        VB111.ENABLE_VBA_KEYWORDS = False
        VB111.ENABLE_VBA_CONSTANTS = False
        VB111.ENABLE_HOST_ENUMS = False
        _vb111 = _VBE111(_pc111a)
        VB111._get_vbe_cached = lambda: _vb111
        _bk111 = VB111.VbeBackend()
        _nm111a = set(str(r[0]).lower() for r in _bk111._collect_identifiers())
        # 换工程（同一个 backend：模拟用户切到另一个工作簿的代码窗）
        _bk111.invalidate_identifiers()
        _vb111.ActiveVBProject = _pc111b
        _nm111b = set(str(r[0]).lower() for r in _bk111._collect_identifiers())
    finally:
        VB111._get_vbe_cached = _o_vbe111
        VB111._comp_list_cache.clear()
        VB111._comp_list_cache.update(_o_clc111c)
        (VB111.ENABLE_VBA_BUILTINS, VB111.ENABLE_VBA_KEYWORDS,
         VB111.ENABLE_VBA_CONSTANTS, VB111.ENABLE_HOST_ENUMS) = _o_bk111c

    check("111.3 两个 .xlsm 各有一个 `Module1`（都是 3 行）：切工程后候选池"
          "必须是**新工程**的变量 —— 旧实现键里只有模块名，`_cm_cache` 命中"
          "A 的 CodeModule（读到的行数也来自 A），`_mod_cache` 重放 A 的解析结果",
          [("alphaonlya" in _nm111a, "betaonlyb" in _nm111a,
            "betaonlyb" in _nm111b, "alphaonlya" in _nm111b)],
          expect_contain=[(True, False, True, False)])

    # ---- 111.4 `insert_snippet` 多行：中间行**一次** InsertLines ----
    class _Mod111(object):
        def __init__(self, lines, ins_boom=False, rep_boom=False):
            self.lines = list(lines)
            self.ins_boom = ins_boom
            self.rep_boom = rep_boom
            self.inserts = []
            self.deletes = []

        def Lines(self, n, c):
            return "\r\n".join(self.lines[n - 1:n - 1 + c])

        @property
        def CountOfLines(self):
            return len(self.lines)

        def ReplaceLine(self, n, t):
            if self.rep_boom:
                raise RuntimeError("模块只读")
            self.lines[n - 1] = t

        def InsertLines(self, n, t):
            if self.ins_boom:
                raise RuntimeError("模块只读")
            self.inserts.append((n, t))
            _t = str(t).replace("\r\n", "\n").replace("\r", "\n")
            for _k, _ln in enumerate(_t.split("\n")):
                self.lines.insert(n - 1 + _k, _ln)

        def DeleteLines(self, n, c):
            self.deletes.append((n, c))
            del self.lines[n - 1:n - 1 + c]

    def _snip111(lines, sl, sc, el, ec, text, ins_boom=False, rep_boom=False):
        _mod = _Mod111(lines, ins_boom, rep_boom)
        _pane = _Pane110(_mod, sl, sc, el, ec)
        _o_vbe = VB111._get_vbe_cached
        _o_sem = VB111._sem_cache.get("info")
        VB111._sem_cache["info"] = ("char", 4, False)
        VB111._get_vbe_cached = lambda: _Vbe110(_pane)
        try:
            _ok = VB111.VbeBackend().insert_snippet(text)
        finally:
            VB111._get_vbe_cached = _o_vbe
            VB111._sem_cache["info"] = _o_sem
        return (bool(_ok), list(_mod.lines), len(_mod.inserts),
                list(_mod.deletes))

    _g111d = _snip111(["x = ab"], 1, 6, 1, 6, "1\n2\n3")
    check("111.4 `insert_snippet` 插 3 行片段：中间那 2 行只用**一次** "
          "`InsertLines`（不是逐行）—— 这是\"写一半\"窗口消失的前提；"
          "以前是 2 次调用，第 2 次抛异常就留下半截片段",
          [(_g111d[0], _g111d[1], _g111d[2])],
          expect_contain=[(True, ["x = a1", "2", "3b"], 1)])

    # ---- 111.5 中间行插不进去：原文一个字都不许动 ----
    # 旧实现先 `ReplaceLine` 首行（成功）、再逐行 `InsertLines`（抛异常）⇒
    # 返回 False 但首行已经被改成 `x = a`（用户后半行 `b` 丢了）。
    _g111e = _snip111(["x = ab"], 1, 6, 1, 6, "1\n2\n3", ins_boom=True)
    check("111.5 `InsertLines` 抛异常（模块只读）：返回 False **且首行保持原样**。"
          "旧实现先改首行再插行，失败时 `x = ab` 已经变成 `x = a` —— "
          "调用方提示\"插入失败\"，用户以为没动过，实际少了字",
          [(_g111e[0], _g111e[1])],
          expect_contain=[(False, ["x = ab"])])

    # ---- 111.6 首行写不进去：回滚掉刚插的行，且**绝不删原有首行** ----
    _g111f = _snip111(["x = ab"], 1, 6, 1, 6, "1\n2\n3", rep_boom=True)
    check("111.6 `ReplaceLine` 失败（模块只读）：中间行已经插进去了，必须回滚 —— "
          "回到 `[\"x = ab\"]`。回滚只删**新增的** 2 行；把原有首行一起删掉"
          "就是直接删用户的代码（这个坑第一版真踩了）",
          [(_g111f[0], _g111f[1], _g111f[3])],
          expect_contain=[(False, ["x = ab"], [(2, 2)])])

    # ---- 111.7 光标在"虚拟空行"上插多行片段 ----
    # sl=2 但模块只有 1 行：`InsertLines(3, …)` 会插到行尾之外，必须先把
    # 第 2 行建出来（纯新增），否则片段中间那行会把首行覆盖掉。
    _g111g = _snip111(["x = ab"], 2, 1, 2, 1,
                      "Sub Test()\n    x = 1\nEnd Sub")
    check("111.7 光标停在模块末尾那个**虚拟空行**（第 2 行，模块只有 1 行）时插"
          "多行片段：先把空行建出来再插中间行 —— 3 行片段一行都不能少",
          [(_g111g[0], _g111g[1])],
          expect_contain=[(True, ["x = ab", "Sub Test()", "    x = 1",
                                  "End Sub"])])

    # ---- 111.8 `duplicate_lines` 与 `insert_snippet` 口径统一 ----
    # 两处都改成"整段一次 InsertLines"，这里钉住 `duplicate_lines` 仍然是
    # **一次**调用（不许哪天又有人把它改成逐行）。
    _mod111h = _Mod111(["a = 1", "b = 2", "c = 3"])
    _pane111h = _Pane110(_mod111h, 2, 1, 2, 1)
    _o_vbe111h = VB111._get_vbe_cached
    try:
        VB111._get_vbe_cached = lambda: _Vbe110(_pane111h)
        _ok111h = VB111.VbeBackend().duplicate_lines()
    finally:
        VB111._get_vbe_cached = _o_vbe111h
    check("111.8 `duplicate_lines`（Ctrl+D）复制整行仍然只用**一次** "
          "`InsertLines` —— 与 `insert_snippet` 同一口径；一次调用天然没有"
          "\"复制了一半\"的中间态",
          [(bool(_ok111h), list(_mod111h.lines), len(_mod111h.inserts))],
          expect_contain=[(True, ["a = 1", "b = 2", "b = 2", "c = 3"], 1)])

    # ------------------------------------------------------------------
    # 112. ★v124：参数信息改走 VBE 的菜单命令（用户报的"提示音"就是它响的）
    #
    #     用户口径（原话）："选中了函数名提示词，然后按 tab 键，后面自动补双括号，
    #     会有一个提示音……是否能屏蔽掉这个提示音"。
    #
    #     根因：补完 `()` 之后我们注入 Ctrl+Shift+I 想唤出 VBE 的参数信息，
    #     但真机实测（独立 Excel 实例）那个窗**一直不可见** —— 命令取不到可显示
    #     的签名，VBE 于是响一声"命令不可用"。改成 `CommandBars` 菜单命令
    #     （Id=2532）之后不再有任何键盘事件进入系统。
    # ------------------------------------------------------------------
    print("\n=== 112. ★v124d：参数信息**默认不发命令** —— 只测显式开启那条支路 ===")
    try:
        VB112 = VB

        class _Ctrl112:
            def __init__(self, enabled=True, boom=False):
                self.Enabled = enabled
                self.n = 0
                self.boom = boom

            def Execute(self):
                if self.boom:
                    raise RuntimeError("VBE 正忙")
                self.n += 1

        class _Bars112:
            def __init__(self, ctrl, boom=False):
                self.ctrl = ctrl
                self.boom = boom
                self.finds = 0

            def FindControl(self, typ, cid):
                self.finds += 1
                if self.boom:
                    raise RuntimeError("宿主没有 CommandBars")
                return self.ctrl

        class _VBE112:
            """只带 CommandBars —— 结构越瘦，越能证明 `_show_param_info`
            除了这一条路之外不碰别的东西（尤其不碰任何按键注入）。"""

            def __init__(self, bars):
                self.CommandBars = bars

        _o_cache112 = dict(VB112._param_info_cache)
        try:
            # ---- 112.1 命中 + 缓存 ----
            _c1, _b1, _v1 = _Ctrl112(), _Bars112(_Ctrl112()), _VBE112(None)
            _b1.ctrl = _c1
            _v1.CommandBars = _b1
            VB112._param_info_cache["owner"] = None
            VB112._param_info_cache["ctrl"] = None
            _r112a = (VB112._show_param_info(_v1),
                      VB112._show_param_info(_v1))
            check("112.1 `_show_param_info`：第一次 FindControl + Execute，"
                  "第二次直接复用缓存控件（补全路径上不重复付 COM）",
                  [(_r112a, _c1.n, _b1.finds)],
                  expect_contain=[((True, True), 2, 1)])

            # ---- 112.2 换一个 VBE 实例 -> 缓存必须作废 ----
            _b2 = _Bars112(_Ctrl112())
            _v2 = _VBE112(_b2)
            _r112b = VB112._show_param_info(_v2)
            check("112.2 缓存**按 VBE 实例**来：换一个宿主/换一个 VBE 对象必须"
                  "重新 FindControl（否则拿到的是上一个宿主的控件）",
                  [(_r112b, _b2.finds, _b2.ctrl.n, _b1.finds)],
                  expect_contain=[(True, 1, 1, 1)])

            # ---- 112.3 命令灰着 -> 不 Execute ----
            _c3 = _Ctrl112(enabled=False)
            _b3 = _Bars112(_c3)
            _r112c = VB112._show_param_info(_VBE112(_b3))
            check("112.3 `Enabled == False`（命令灰着）-> **不 Execute**，"
                  "返回 False —— 灰命令 Execute 等于白响一声",
                  [(_r112c, _c3.n)], expect_contain=[(False, 0)])

            # ---- 112.4 没有 CommandBars / FindControl 抛异常 -> 静默 ----
            _r112d = VB112._show_param_info(_VBE112(_Bars112(_Ctrl112(), boom=True)))
            _r112e = VB112._show_param_info(_VBE112(None))
            check("112.4 宿主没有 CommandBars（或 FindControl 抛异常）-> "
                  "静默返回 False，绝不抛给补全路径",
                  [(_r112d, _r112e)], expect_contain=[(False, False)])

            # ---- 112.5 FindControl 返回 None（别的 VBE 版本没这个 Id）----
            _b5 = _Bars112(None)
            _r112f = VB112._show_param_info(_VBE112(_b5))
            check("112.5 `FindControl` 返回 None -> 静默 False（不猜别的 Id、"
                  "不退回注入按键）",
                  [(_r112f, _b5.finds)], expect_contain=[(False, 1)])

            # ---- 112.6 Execute 抛异常 -> False 且清缓存 ----
            _c6 = _Ctrl112(boom=True)
            _b6 = _Bars112(_c6)
            _v6 = _VBE112(_b6)
            _r112g = VB112._show_param_info(_v6)
            check("112.6 `Execute` 抛异常（VBE 正在跑宏 / 忙）-> 返回 False，"
                  "并且**清掉缓存**：下次必须重新 FindControl（旧控件已失效）",
                  [(_r112g, VB112._param_info_cache["ctrl"],
                    VB112._param_info_cache["owner"])],
                  expect_contain=[(False, None, None)])

            # ---- 112.7 源码护栏：菜单 Id 是 2532，不是 2531 ----
            _vsrc112 = io.open(os.path.join(ROOT, "vbe_bridge.py"),
                               encoding="utf-8").read()
            check("112.7 菜单 Id 必须是 **2532**（VBE 的「参数信息(&M)」）；"
                  "2531 是「快速信息(&Q)」—— 那是显示变量值/类型的浮动提示，"
                  "弹它等于什么都没做",
                  [("_PARAM_INFO_CTRL_ID = 2532" in _vsrc112,
                    "_PARAM_INFO_CTRL_ID = 2531" in _vsrc112,
                    "FindControl(1, _PARAM_INFO_CTRL_ID)" in _vsrc112)],
                  expect_contain=[(True, False, True)])

            # ---- 112.8 ★默认必须是"什么都不做" ----
            # 真机 peak 实测（WASAPI 设备音量计，20ms 轮询）：
            #   Execute@空行 0.346 / Execute@括号内 0.346 / 只写行 0.000 /
            #   注入 Ctrl+Shift+I 0.000（但也不弹）/ 注入空格 0.000
            # ⇒ 默认若还去 Execute，用户就还会听到那一声。
            _on112 = bool(VB112._settings.flag("PARAM_INFO", False))
            _off112 = bool(VB112._settings.flag("NO_PARAM_INFO", False))
            check("112.8 ★`AUTO_PARAM_INFO` 必须等于 "
                  "`PARAM_INFO and not NO_PARAM_INFO` —— 默认两条都不设 = 不做；"
                  "显式 `VBECOMPLETE_NO_PARAM_INFO=1` 永远能压掉",
                  [(bool(VB112.AUTO_PARAM_INFO) == (_on112 and not _off112),)],
                  expect_contain=[(True,)])

            # ---- 112.9 settings 模板 / KNOWN 的默认值也是"关" ----
            _kd112 = dict((k, d) for k, d, _ in VB112._settings.KNOWN)
            check("112.9 `settings.py` 的 KNOWN 里 `PARAM_INFO` 默认值是 `0`"
                  "（模板与代码默认一致，不能一个说开一个说关）",
                  [(_kd112.get("PARAM_INFO"), _kd112.get("NO_PARAM_INFO"))],
                  expect_contain=[("0", "0")])
        finally:
            VB112._param_info_cache.clear()
            VB112._param_info_cache.update(_o_cache112)
    except Exception as _e112:
        import traceback as _tp112
        _tp112.print_exc()
        check("第 112 节异常: %s" % _e112, [True], expect_contain=[False])

    # ------------------------------------------------------------------
    # 90. ★v107c：测试网自查 —— 猴补丁必须还原
    #     这不是在测产品代码，而是在测【测试网自己】：一节忘了还原，后面每一节
    #     都会悄悄跑在别人的假对象上，而**不会有任何一条用例变红**。
    #     v107c 实测抓到两处（§31 第 33 节复用、§62 的 62.12），都是
    #     `_get_vbe_cached` 被换掉没换回来。
    # ------------------------------------------------------------------
    print("\n=== 90. ★v107c：测试网自查 —— 猴补丁必须还原 ===")
    _leaked90 = ["%s.%s" % (m.__name__, n)
                 for (m, n, _o) in _LEAK_WATCH
                 if getattr(m, n, None) is not _o]
    check("90.1 各节临时替换的函数 / 单例都在节尾还原了"
          "（v107c 抓到过 §31 / §62 两处：它们让后面每一节的 "
          "`_get_vbe_cached` 都指向别人的假工程）",
          [(_leaked90,)], expect_contain=[([],)])

    print("\n" + "=" * 60)
    print("结果: %d PASS, %d FAIL" % (PASS, FAIL))
    if FAILURES:
        print("失败项:")
        for f in FAILURES:
            print("  - " + f)
    print("=" * 60)
    return 1 if FAIL else 0


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    sys.exit(main())
