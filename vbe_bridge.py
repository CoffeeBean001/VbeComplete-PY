"""
VBE 后端：通过 pywin32 与正在运行的 VBA 宿主的 VBE 交互。

只在这一层依赖 pywin32 / 宿主，引擎层不依赖，方便单测。

支持 Office 全家桶（Excel / Word / PowerPoint / Access / Visio）与 WPS 全家桶
（文字 / 表格 / 演示）—— 它们共用同一个 VBE 与同一套 VBIDE 对象模型，差异只在
外层 Application 的 ProgID，由 `_get_vbe()` 自动探测（详见那一节的说明）。

需要：对应宿主已开启「信任对 VBA 工程对象模型的访问」（各宿主菜单路径一致：
信任中心 → 宏设置 → 信任对 VBA 工程对象模型的访问）。
"""

import difflib
import re
import threading
import time
import unicodedata

from log import log as _log

import parser as vba_parser
import settings as _settings
import vba_builtins
# ★v103：`_in_rem_comment` 是 `Rem` 判据的【唯一实现】（在 engine 里）。
# 本文件的 `_classify` 原先漏认 Rem（只认 `'`），导致注释里照弹候选、
# 自动配对也不让路 —— 复用同一份，不许在这里再写第二份 Rem 判据。
from engine import replace_word, _in_rem_comment

_CACHE_TTL = 2.0  # 标识符缓存刷新间隔（秒）

# ★v118：按模块缓存解析结果 —— 单次全量刷新从 ~220ms 降到 ~45ms。
#
# 病根（真机实测，24 个模块 / 628 行的小工程）：`_collect_identifiers` 一次
# 要 220ms，而 main 的轮询在【每次文本变化】后只要距上次刷新 > 0.4s 就
# invalidate 一次 —— 也就是打着字时每 0.4 秒就有一拍要在主线程上付这 220ms，
# 用户看到的就是"弹窗明显滞后、打字一顿一顿"。
#
# 220ms 里真正做 Python 解析的只有 ~40ms，其余全是 COM：每轮重新枚举
# VBComponents、每个模块重新拿 CodeModule（每次都是新的 pywin32 包装对象，
# 要重新 Bind 类型信息）、再把每个模块的全文读一遍。
#
# 而"其它模块"的解析结果其实【只由它的代码文本决定】，与光标无关：
#   * 光标所在的模块（caret_mod）每轮都重读重解析 —— 用户正在改的就是它；
#   * 其它模块记下 (行数, 解析产物)，行数没变就整段复用；
#   * 再加 _MOD_CACHE_TTL 兜底：超过这个时间无论如何重扫一次，防"行数不变
#     但内容改了"（跨模块查找替换那种改法）永远不生效。
#
# 开关 `NO_MOD_CACHE = 1`（或环境变量 VBECOMPLETE_NO_MOD_CACHE=1）退回 v117 的
# 逐模块全量解析（A/B 用）。
# 模块缓存的最长有效期（秒）。超时强制重扫该模块一次。
_MOD_CACHE_TTL = 30.0
MOD_CACHE_ENABLED = not _settings.flag("NO_MOD_CACHE", False)


def _env_flag(name, default):
    """三态开关（★v120 起转交配置中心）：`1/true/yes/on` -> True，
    `0/false/no/off` -> False，未设置 / 认不出来 -> default。

    用意：下面那些"收不收某类名字"的开关，默认值写死在代码里（跟着用户口径走），
    但要能不改源码就翻过来 —— 在 `settings.ini` 里写一行、或双击 run.bat 之前
    先 `set XXX=1` 即可。

    取值链：**环境变量 > settings.ini > 这里的 default**；"认不出的值"
    （`=ture`）会被跳过而不是让功能失效 —— 细节见 settings 模块顶部。
    """
    return _settings.flag(name, default)


# 未声明即使用的"隐式变量"（不写 Option Explicit 时的用法）是否纳入提示。
# 关闭后行为回到旧版：只提示 Dim/Const/Sub 等显式声明出来的名字。
ENABLE_IMPLICIT_IDENTIFIERS = True
# 隐式变量的作用域粒度：
#   "module" —— 本模块任意位置都能提示（最宽松，但会跨过程泄漏同名变量）
#   "proc"   —— 按"首次出现所在过程"限定作用域（更贴近 VBA 语义，避免跨过程泄漏）
IMPLICIT_SCOPE = "proc"

# ★v75：隐式变量要不要【尊重模块自己写的 Option Explicit】。
#
# 用户口径："没有 Option Explicit 时，用过的变量名就算已经被定义、可以直接提示；
# 写了 Option Explicit 是强制声明，没声明过的变量名在后续就不该被提示。"
#
# 判定是【按模块】做的 —— Option Explicit 本来就是模块级语句：
#   * 没写它的模块：完全保持原样（用过的名字照旧进池子，用户明确说这样好）；
#   * 写了它的模块：只收"真声明过"的名字（Dim / Const / Sub / Function / Type /
#     Enum / 形参）；"只用过、没声明"的名字不再进池子 —— 那种写法在该模块里
#     是编译错误，提示它只会带出坏代码。
#
# 想回到旧行为（不分模块，隐式变量一律收）：
#   set VBECOMPLETE_IMPLICIT_HONOR_EXPLICIT=0
IMPLICIT_HONOR_OPTION_EXPLICIT = _env_flag(
    "VBECOMPLETE_IMPLICIT_HONOR_EXPLICIT", True)

# VBA 语言自带的名字（内建函数 / 内建数据类型）是否纳入提示。
#
# 历史（v61 → v77 → ★v96）：
#   v61  默认 True  —— "VBA 本身自带的函数也加入提示列表，现在不能提示"；
#   v77  默认 False —— "提示词太多了……vba 本身带的关键字、函数这块都删掉吧。
#                      保留变量名、自定义函数/过程名、窗体/控件名、模块名"；
#   v96  默认 True  —— ★用户原话："我想把 vba 关键字、函数名这两部分，也加入到
#                      提醒列表里。枚举类型就不要加了，枚举加上，就会太多了。"
#                      ⇒ 收【关键字 + 内建函数 + 内建数据类型】，仍然不收枚举。
#
# ⚠️ **枚举永远是另一批、且默认不收**（见下面 ENABLE_VBA_CONSTANTS /
#    ENABLE_HOST_ENUMS）—— 用户这次也明确说了"枚举加上就会太多"。
#    v70 的原话是"内置枚举我基本不用，提示出来只是干扰"，v96 没有改这个判断。
#
# 现在默认口径：
#   * 收 —— 你这个工程里真实存在的东西（变量 / 自定义过程 / 组件名 / 控件名），
#           **加上** VBA 关键字（Sub / If / Dim / Set …）与内建函数
#           （MsgBox / Left / Split …）、内建数据类型（`Dim x As Long` 用得上）；
#   * 不收 —— 一切枚举：VBA 内建 vb*（102 条）与宿主类型库的 xl* / mso*（4538 条）。
#
# ⚠️ 只改【收集侧开关】，四份清单本身一个字都没裁（见 vba_builtins.py）——
# 清单保持完整，测试/诊断仍能看到全貌，需要时一行环境变量立刻关掉。
# 想关回去：`set VBECOMPLETE_VBA_BUILTINS=0`。
ENABLE_VBA_BUILTINS = _env_flag("VBECOMPLETE_VBA_BUILTINS", True)

# ★v70：VBA 内建的 **枚举常量**（`vbOK` / `vbCrLf` / `vbYes` … 共 102 条）是否纳入提示。
#
# 用户口径（v70）："我不怎么使用 VBA 内置枚举，提示出来对我有干扰 —— 把枚举关掉，
# 函数那些要保留。" 所以它与"内建函数"【分家】：函数（MsgBox / Left / Split …）
# 与内建数据类型（Long / String …）照常提示，只把这一批枚举常量摘出去。
# 想开回来：`set VBECOMPLETE_VBA_CONSTANTS=1`。
ENABLE_VBA_CONSTANTS = _env_flag("VBECOMPLETE_VBA_CONSTANTS", False)

# 语言关键字 / 保留字（Sub / Dim / If / For / Set / And …）是否纳入提示（v62）。
#
# 独立成一个开关（而不是并进 ENABLE_VBA_BUILTINS）：关键字与内建函数是两拨
# 东西 —— 前者是语法骨架、后者是可调用的库成员。你想单收一批时用得上。
#
# ★v96：**默认打开**（用户口径："我想把 vba 关键字、函数名这两部分，也加入到提醒
# 列表里。枚举类型就不要加了。"）—— 与 ENABLE_VBA_BUILTINS 一起回来了，
# 而两批【枚举】（vb* 内建常量、宿主 xl*/mso*）仍然默认不收。
# 想关回去：`set VBECOMPLETE_VBA_KEYWORDS=0`。
# 与 ENABLE_VBA_BUILTINS 是两个独立开关，所以"只收关键字、不收内建函数"也做得到。
ENABLE_VBA_KEYWORDS = _env_flag("VBECOMPLETE_VBA_KEYWORDS", True)
# 内建名字挂靠的"模块名"。
#
# 这是一个【虚拟模块】——VBA 运行时库。之所以要给它一个名字而不是留空：
#   1) 这些名字在工程里任何模块都能裸名引用，按模块级 priv=False 收录即可；
#   2) 引擎判"回声候选是不是真名字"时会问 declared_elsewhere（名字是不是
#      在【别的模块】声明过）。把内建名字的归属记成 VBA，那一问天然为真，
#      内建名字就不会被"只在本模块声明过 + 现场文本扫不到 -> 当幽灵剔掉"
#      这条规则误杀（详见 engine._name_really_exists）。
_BUILTIN_MODULE = "VBA"

# 宿主类型库的枚举常量是否纳入提示（v67）。
#
# 用户报："输入 vb 会提示一堆 VBA 枚举值，输入 xl 却不提示任何 xl 开头的枚举值。"
# 前者是 v61 收的 VBA 内建常量（vb*），后者是【宿主 Excel 类型库】的枚举常量
# （xl*）—— 那是一整套库，从来没收过，于是同一个动作在 xl 上一个提示都没有。
#
# 这些常量与内建名字同源：由【工程引用的类型库】决定存在，与代码文本无关
# （代码里从没写过 xlUp，打 xl 也该补得出来）。所以走完全相同的四路承接。
#
# 与内建名字的关键区别是【匹配口径】—— 它们数量极大（Excel + Office 两个库
# 实测 4538 条），若像工程内名字那样接受跳步匹配，输入 3~4 个字母就会带出
# 几百条无关项。所以：
#   * 只收"有统一家族前缀"的库（Excel -> xl，Office -> mso），
#     没有家族前缀的库（stdole 的 Checked / Gray / Color 这种通用词）一概不收；
#   * 只认【从头开始的连续前缀】命中，且输入长度不低于家族前缀长度 ——
#     于是 xl -> xl*、xlu -> xlUp；而 ms 只出 MsgBox，绝不带出 mso*。
#
# ★v70：**默认关闭**（用户口径："xl 开头这些枚举也不要提示，我不怎么用"）。
# 想开回来：`set VBECOMPLETE_HOST_ENUMS=1`。关着的时候收集段整段不跑
# （连类型库都不会去加载，零开销），引擎那边也拿不到宿主清单。
ENABLE_HOST_ENUMS = _env_flag("VBECOMPLETE_HOST_ENUMS", False)
# 家族前缀的覆盖率门槛：某前缀能盖住该库这么多比例的枚举成员，才算"家族前缀"。
_HOST_ENUM_FAMILY_PCT = 0.5
# 允许的家族前缀白名单（小写，逗号分隔）。留空 = 自动推导（推荐）。
#   例：想只留 Excel 的 xl*，设 VBECOMPLETE_HOST_ENUM_FAMILIES=xl
_HOST_ENUM_FAMILIES = tuple(
    p.strip().lower() for p in
    _settings.items("HOST_ENUM_FAMILIES"))
# 宿主常量挂靠的虚拟模块名（同 _BUILTIN_MODULE 的道理，见那边的说明）。
_HOST_ENUM_MODULE = "宿主库"

# 合法标识符（含中文）的模块名/窗体名/类名，用于把组件名纳入候选
_RE_PLAIN_IDENT = re.compile(r"^[^\W\d]\w*$")
# 行首空白（空格 / Tab）：新起一行时用它对齐上一行的代码起始位置。
_RE_LEADING_WS = re.compile(r"[ \t]*")

# 跨工程作用域：
#   False（默认）—— 只收集【当前活动工程】里的标识符。
#     VBE 里常常同时开着好几个工程（多个工作簿、加载宏 .xlam、个人宏工作簿
#     Personal.xlsb）。除了显式"工具-引用"之外，一个工程里的窗体名、函数名、
#     全局变量名在另一个工程里根本不可见，提示它们纯属噪音（典型症状：在当前
#     工程里敲一个名字，冒出来的却是另一个工作簿里的 UserForm1 / 全局变量）。
#   True —— 恢复旧行为：收集 VBE 里所有工程的标识符。
ENABLE_CROSS_PROJECT = False

# ★v98 补全时自动补括号 —— 只在【调用点】补，不在声明处补。
#
# 用户口径："选中函数（包括 vba 本身函数）或过程提示词，按 tab 后，应该能在
# 函数或过程名后面自动补全双括号，并且光标能正确落在括号内。"
#
# 为什么必须限定"调用点"：VBE 的**语法**在声明处与调用处不同 —— 声明处
# `Function f` 后面本来就要跟 `()`（VBE 自己也会补），调用处 `f` 后面**不跟**
# 括号才是合法的裸引用。真正会编译不过的是那些"两边都不该带括号"的位置：
#   Dim n As Long   ->  Long()   ✘
#   If x = Long Then ->  Long()   ✘
# 所以判据是"【这一行的这个位置】是不是一个可以直接调用过程的位置"。
AUTO_PAREN_ON_COMPLETE = _env_flag("VBECOMPLETE_AUTO_PAREN", True)

# ★v103b：补全一个"后面要跟空格"的 VBA 关键字时，顺手补一个空格
# （用户口径："在使用 public，private，sub，function，set，if，for，with，
# select，end，exit 等这类关键词时，后面需要有空格，要在按 tab 键后，也能
# 自动补全空格，idea 里就有这个功能"）。
# 哪些关键字要补由 `vba_builtins.keyword_wants_space` 判（纯函数）；
# 这里只管"这个位置允不允许补"以及"怎么补"。
#
# ⚠️⚠️ 真机实测（独立 Excel 实例 + 真 CodeModule）：**VBE 会把 ReplaceLine
# 写进去的行尾空格立刻 trim 掉** ——
#     写 `"    public "` 读回 `"    public"`（尾字符 'c'，空格没了）；
#     两次 ReplaceLine、光标移走再读、读别的行之后再读，一律还是 `"    public"`。
# ⇒ 靠改写行文本（`new_line[:k] + " " + new_line[k:]`）**根本补不上**。
# 唯一可行的办法：写完这一行、光标到位之后，**注入一个空格按键**。
# 实测注入的空格会被保留（读回 `"    public "`），后面再注入 'S' 得到
# `"    public S"` —— 正是用户要的效果。
AUTO_SPACE_AFTER_KEYWORD = _env_flag("VBECOMPLETE_AUTO_SPACE", True)

# ★v107c：按 Tab 确认函数名、自动补出 `()` 并把光标送进括号之后，顺手唤出
# VBE 自己的【参数信息】窗（形参签名）。
#
# 用户口径（原话）："提示词选中了过程名或函数名，按 tab 键，当自动补充完双括号后，
# 光标落在双括号中间，再加个自动显示形参列表参数的功能。"
#
# 为什么是"让 VBE 自己弹"、而不是我们画一个窗：形参信息本来就是 VBE 自带的能力
# —— 它会**高亮当前正在填的那个参数**、跟着逗号往下走、按 Esc 收起、跟着你继续
# 改签名。我们手上只有 CodeModule 的纯文本，重画一遍既费事又做不到位（v103b 曾
# 在详情行里显示过形参表，用户明确说不要那种）。
#
# ★★v124c（2026-10-18，用户第一次报"按 Tab 补括号会有一个提示音"）：v107c~v123 用
# `SendInput` 模拟 Ctrl+Shift+I。真机实测（独立 Excel 实例，Excel 2016 / VBE 7.1）：
#   * VBE 的「参数信息」是 CommandBars 里 **Id=2532** 的菜单项（"参数信息(&M)"），
#     Ctrl+Shift+I 只是它的加速键；Id=2531 是另一个东西（"快速信息(&Q)"）。
#   * `ReplaceLine` 写完 `()` 之后再注入 Ctrl+Shift+I，`PopupTipWndClass`
#     **一直不可见** —— 命令在这个上下文里取不到可显示的签名，而 VBE 对"命令不可用"
#     会响一声系统提示音。**用户听到的就是这一声**（也解释了为什么从来没人见过那个窗）。
#
# ★★v124d（2026-10-18，用户第二次报"还是会响"，实测推翻 v124c 的结论）：
# v124c 把注入改成了 `CommandBars.FindControl(1, 2532).Execute()`，以为"不可用就静默"。
# **这是错的** —— 用设备级音量计（WASAPI `IAudioMeterInformation.GetPeakValue`，
# 20ms 轮询）逐动作测量，结果：
#     A `Execute()`（光标在空行，Enabled=True）  peak = 0.3457  ← 响
#     B `Execute()`（光标在 `()` 内，Enabled=True） peak = 0.3457  ← 响
#     C 只 `ReplaceLine` + `SetSelection`          peak = 0.0000  ← 静
#     D 注入 Ctrl+Shift+I                          peak = 0.0000  ← 静（但也不弹）
#     E 注入一个空格（关键字补空格那条路）            peak = 0.0000  ← 静
# 并且 A/B 之后 `PopupTipWndClass` 仍然 `IsWindowVisible == False`（rect 有值，
# 说明那个提示窗对象存在、只是不显示）。
# ⇒ **`Execute()` 在这个由 COM 造出来的上下文里永远不可用，每次都会响一声**；
#   `Enabled` 返回 True 是假的（CommandBars 的 Enabled 只反映"用户能否自定义"）。
# ⇒ **参数信息无法从外部可靠唤出**（按键不行、菜单命令不行），主动去做只有代价。
#
# 所以 v124d 起：`AUTO_PARAM_INFO` **默认关闭**，只有显式
# `set VBECOMPLETE_PARAM_INFO=1` 才尝试（那时用户自己承担一声提示音）。
# VBE 自己的"工具-选项-编辑器-自动显示参数信息"仍然有效 —— 用户在括号里键入
# 第一个实参时它会自己弹，不需要我们插手。
AUTO_PARAM_INFO = (_env_flag("VBECOMPLETE_PARAM_INFO", False)
                   and not _env_flag("VBECOMPLETE_NO_PARAM_INFO", False))

# ★v124：没有 CommandBars 时要不要退回"注入 Ctrl+Shift+I"。
#   默认 **0（不退回）** —— 那条路会响，正是用户要摆脱的东西。
#   开成 1 只在"确认这个宿主的参数信息只能靠按键唤出、并且不在乎那一声"时用。
PARAM_INFO_KEY_FALLBACK = _env_flag("VBECOMPLETE_PARAM_INFO_KEY", False)

# VBE 的「参数信息」菜单项 Id（VBE 7.1 实测；菜单条上写作"参数信息(&M)"）。
# 「快速信息」是 2531，别混 —— 那是显示变量值/类型的浮动提示，不是形参签名。
_PARAM_INFO_CTRL_ID = 2532

# 控件对象按 vbe 实例缓存：它绑在那个 VBE 上，宿主重启后旧对象会失效，
# 失效时调用方清空这里、下次重新 FindControl。零缓存的话每次都要一次
# COM 调用（FindControl 实测 ~1ms，补全路径上不值得反复付）。
_param_info_cache = {"owner": None, "ctrl": None}


def _param_info_control(vbe):
    """拿到 VBE 的「参数信息」菜单项（Id=2532），拿不到返回 None。★v124"""
    ent = _param_info_cache
    try:
        if ent["owner"] is vbe and ent["ctrl"] is not None:
            return ent["ctrl"]
    except Exception:
        pass
    ctrl = None
    try:
        # FindControl(Type=1, Id=2532)：1 = msoControlButton。
        # 不用遍历那三十来个 CommandBars —— 实测要逐个 bar 再逐个控件，
        # 而 FindControl 一次就命中（覆盖所有菜单条上的同一个命令）。
        ctrl = vbe.CommandBars.FindControl(1, _PARAM_INFO_CTRL_ID)
    except Exception:
        ctrl = None
    if ctrl is not None:
        ent["owner"] = vbe
        ent["ctrl"] = ctrl
    return ctrl


def _show_param_info(vbe):
    """让 VBE 弹出它自己的【参数信息】窗，返回是否真的发了一条命令。★v124

    ⚠️ **默认根本不会被调用**（`AUTO_PARAM_INFO` 默认 False）—— 真机实测这个
    命令在我们造出来的上下文里取不到签名，`Execute()` 只会响一声而窗不出现
    （v124d 的 peak 实测见 `AUTO_PARAM_INFO` 上面那段）。留着它只给
    `VBECOMPLETE_PARAM_INFO=1` 的用户显式尝试。

    任何一步不成立都静默返回 False：形参信息与"补全本身"无关，绝不能因为它
    没出来就影响补全（与 v103b 注入空格同一条规矩）。
    """
    ctrl = _param_info_control(vbe)
    if ctrl is None:
        return False
    try:
        # 命令灰着时 Execute 反而更容易响 —— 先问一句（虽然 Enabled 常是假的）。
        if not bool(ctrl.Enabled):
            return False
    except Exception:
        pass          # 有的宿主不给 Enabled，那就直接试
    try:
        ctrl.Execute()
        return True
    except Exception:
        # 控件随 VBE 重启/换工程失效：清缓存，下次重新 FindControl。
        _param_info_cache["owner"] = None
        _param_info_cache["ctrl"] = None
        return False

# ★v110：Ctrl+↑ / Ctrl+↓ 在【当前过程】内跳转（用户 2026-10-15 提的两条）：
#   Ctrl+↑：光标落到本过程最后一个 `Dim` 语句【那一行】的行尾；
#   Ctrl+↓：跳到本过程 `End Sub` / `End Function` / `End Property` 行【上面】
#           那一行 —— 不是空白就补一行空白，缩进照"过程体层级"。
# 用户口径（原话）："当我按 ctrl+↑，光标会自动定位到当前过程中最后一个 dim
# 关键字行，如果此行的下一行不是空行，就自动在此行下方新增一行，光标移动到
# 下方空白行；如果此行的下一行是空白行，就不用新增行，光标移动到下方空白行
# 就可以了。ctrl+↓ 同理，定位到 end sub 或 end function 行。"
# ★v110c（用户改口径）："光标落在本过程或函数的最后一个 dim 行的行尾，不用
# 落在最后 dim 下一行并插入新行了。" ⇒ Ctrl+↑ 不再插行、不再补缩进，只把光标
# 落到 Dim 语句那一行的行尾（跨行 Dim 落到语句【结束行】的行尾）。
#
# 为什么值得做：VBA 的变量声明区固定压在过程开头，"往声明区加一个变量"和
# "往过程末尾追加几行"是两个高频动作，而 VBE 原生 Ctrl+↑/↓ 什么都不做。
#
# 骨架是第 10 条那副老规矩（v94 / v99 / v78 同一副）：判据全在纯函数里
# （本文件下面 v110_* 一族），既能被单测直接钉住、也能在键盘钩子线程里安全
# 调用；真正的读模块 + 插行 + 落光标在主线程（VbeBackend.ctrl_nav_jump），
# 写不成就把这一下按键原样还给系统。想关掉：`set VBECOMPLETE_NO_CTRL_NAV=1`
# （开关本身在 main.py，与 SHIFT_NAV_JUMP 放在一起）。
#
# ★v110b：Ctrl+↓ 落到收尾区空行时，**把那一行直接写成过程体层级的缩进**。
# 用户口径（原话）："还有个地方需要修改，ctrl+↓ 跳到结尾空行，能不能再加个
# tab 补位？" —— 落点那一行常常是顶格的真空行（收尾行 `End Sub` 惯例顶格），
# 光标落在行首，用户接着写过程体代码还得自己按一次 Tab。
# 缩进抄"尾行上方最后一条非空行"的行首空白（与 Ctrl+↑ 抄 Dim 行缩进同一条理），
# 那正是用户接着要写语句的层级。
# ⚠️ 只在落点是【顶格 / 真空行】时补：带着缩进的空白行说明用户自己摆过层级，
#    尊重，绝不加深（用户没要求的层级我们不加）。
#
# ★v110b/2（修真机 bug）：原先这一处是**注入一下 Tab 键**、让 VBE 自己的
# smart indent 去补。真机实测**不生效**（跳转正常、缩进没有）：SetSelection 是
# COM 调用，返回时 VBE 未必已把光标画到新行，注入的 Tab 打在旧行上；补了
# 轮询 GetSelection + 延时也仍不出缩进。改成**直接把新行写过程体缩进**之后
# 就完全确定（`InsertLines` 写空白行不会被 VBE trim —— 这一点由现成的"回车
# 缩进" `new_line_below` 作旁证，它一直就是直接写 `indent`，真机好用）。
# 想关掉（只落光标 + 只保证有一行空白，不补缩进）：`set VBECOMPLETE_NO_CTRL_NAV_TAB=1`。
CTRL_NAV_TAB = not _env_flag("VBECOMPLETE_NO_CTRL_NAV_TAB", False)


def _co_free():
    """释放 pywin32 缓存的 COM 代理（仅在显式开启时才会真正调用）。

    说明：CoFreeUnusedLibraries 只释放"当前无人引用"的 COM 库，**并不能**释放
    我们持有的 Excel/VBE 引用，对"让 Excel 退出"其实没有帮助；反而在 Excel
    正在关闭时反复调用会触发 COM 库的卸载/重载，是"关文件卡几秒、关完又被
    拉起"的嫌疑点之一。因此默认不再自动调用，需要排障时用环境变量
    VBECOMPLETE_COFREE=1（或 settings.ini 的 `COFREE = 1`）打开。
    pythoncom 延迟导入，避免无 pywin32 环境报错。
    """
    if not _settings.flag("COFREE", False):
        return
    try:
        import pythoncom
        pythoncom.CoFreeUnusedLibraries()
    except Exception:
        pass


# ---------------------------------------------------------------------------
# VBE 行列坐标处理
#
# 现象：按 Tab 确认后，光标偶尔落在"词的中间"而不是词尾。根因是行内含
# 制表符（Tab 缩进）时，字符位置 与 VBE 的列 可能不一致：
#   - 若 VBE 的列=字符位置（Tab 计 1 列）：我们按 Python 字符数定位应无偏差；
#   - 若 VBE 的列=显示位置（Tab 展开为多格）：按字符数 SetSelection 会少算
#     若干格，光标就落在词中间。
#
# 处理（双保险，无需预知 VBE 语义）：
#   1. _vbe_measure：向"远超行尾"的列 SetSelection，读回停靠列 E 来探测
#      该行列语义与 Tab 宽度（字符语义会被钳制到 len+1，显示语义会钳制到
#      显示宽度+1，二者数值不同，可区分）。
#   2. 写回时若行内含 Tab，先把整行的 Tab 展开为空格再 ReplaceLine ——
#      该行此后不含 Tab，字符位置与列恒等，光标定位在任意语义下都精确。
#      唯一的副作用：被补全的那一行的缩进从 Tab 变成等宽空格（纯视觉，
#      VBA 语义不变，视觉宽度也不变）。
# ---------------------------------------------------------------------------

def _is_ident_tail(ch):
    """标识符尾字符判定（Unicode 感知：中文等字母/数字/下划线都算）。"""
    return ch.isalnum() or ch == "_"


def _is_wide(ch):
    """是否会被等宽编辑器当成 2 格宽的全角字符（CJK/全角）。"""
    return unicodedata.east_asian_width(ch) in ("W", "F")


def _disp_width(line_text, tabw, wide2, upto=None):
    """把 line_text[:upto] 按显示宽度展开（Tab 到下一个 tab stop）。"""
    w = 0
    end = len(line_text) if upto is None else min(upto, len(line_text))
    for ch in line_text[:end]:
        if ch == "\t":
            w += tabw - (w % tabw)
        else:
            w += 2 if (wide2 and _is_wide(ch)) else 1
    return w


def _col_to_char_index(line_text, disp_col, tabw, wide2):
    """显示列(1-based) -> 字符下标(0-based，光标停在该字符前)。

    若 disp_col 落在行尾之后则返回 len(line_text)。
    """
    off = 0
    for i, ch in enumerate(line_text):
        if disp_col == off + 1:
            return i
        if ch == "\t":
            off += tabw - (off % tabw)
        else:
            off += 2 if (wide2 and _is_wide(ch)) else 1
    return len(line_text)


def _fit_semantics(line_text, end_col):
    """用"某个已知的行尾列 end_col"反推 (sem, tabw, wide2)。

    end_col 必须来自 VBE 自己报告的列（光标停靠列或行尾列）。
      end_col <= len+1 -> 列没有超过字符数，只能是"字符列"语义；
      否则            -> 必然是"显示列"语义，用候选 (tabw, wide2) 拟合。
    """
    n = len(line_text)
    if end_col <= n + 1:
        return "char", 4, False
    # 候选顺序：优先最常见的 tabw=4。行内无 Tab 时 tabw 不影响结果，
    # 若不把 4 放前面会先命中 2，导致缓存里的 tabw 看起来"很怪"。
    for wide2 in (False, True):
        for tabw in (4, 8, 2, 3):
            try:
                if _disp_width(line_text, tabw, wide2) + 1 == end_col:
                    return "disp", tabw, wide2
            except Exception:
                continue
    # 兜底：假设只有 Tab 造成展开
    ntabs = line_text.count("\t")
    if ntabs > 0:
        extra = (end_col - 1) - n
        if extra % ntabs == 0:
            t = extra // ntabs + 1
            if 1 <= t <= 32:
                return "disp", t, False
    # 兜底：行内含全角字符却无法用候选参数复现时，优先假定"全角占 2 列"——
    # 比当成 1 列更贴近 VBE 实际行为，避免光标落在中文变量名中间。
    if any(_is_wide(c) for c in line_text):
        return "disp", 4, True
    return "disp", 4, False


def _detect_semantics_passive(line_text, ec):
    """【完全不移动光标】地判定列语义；判不出来返回 None。

    原理：字符列语义下 VBE 报告的列不可能超过 len+1。若 ec 已经 > len+1，
    则它必然是显示列，直接用它拟合 (tabw, wide2) 即可。
    打字时光标几乎总在行尾，这条路径能覆盖绝大多数真实输入，
    因此可以避免为了探测而移动光标（某些编辑器会把光标还原错/还原延迟，
    表现为"光标自己跳到代码中间"）。
    """
    if not line_text:
        return None
    if ec <= len(line_text) + 1:
        return None
    return _fit_semantics(line_text, ec)


# 列语义是"编辑器级"属性（不随行变化），因此探测一次即可长期复用。
_sem_cache = {"info": None, "probe_broken": False}


def _leading_ws(line_text):
    """行首空白（空格 / Tab）。

    新起一行时用它对齐上一行的代码起始位置 —— VBE 自己按回车也遵循这个规则
    （新行继承上一行缩进），我们在光标位于行中间时代劳，必须手工复刻。
    """
    m = _RE_LEADING_WS.match(line_text or "")
    return m.group(0) if m else ""


def _resolve_sem_info(cm, cp, line_no):
    """取当前编辑器的列语义三元组 (sem, tabw, wide2)；拿不到返回 None。

    ★v93：`_indent_end_col` 在缩进含 Tab 时【必须】有语义信息才敢算显示列。
    缓存（_sem_cache）可能还是空的 —— 它只由 _probe_semantics 写入，而那条路
    要"行内含 Tab / 全角且 ec 超过 len+1"才走到，新建一行这种场景可能一次都
    没触发过。这里按 _caret_char_col 的同一套顺序补一次：缓存 -> 被动探测 ->
    主动探测（_probe_semantics 自己会还原光标并校验）。任何异常都吞成 None，
    调用方退回"按显示列 + 默认 tab 宽度"的兜底 —— 绝不因为探测失败就写错列。
    """
    info = _sem_cache.get("info")
    if info:
        return info
    try:
        line_text = cm.Lines(int(line_no), 1)
    except Exception:
        return None
    try:
        _sl, _sc, _el, ec = cp.GetSelection()
        ec = int(ec)
    except Exception:
        ec = 0
    # ★v95：这一行既无 Tab 也无全角 => 两种列语义在它上面【数值完全相同】，
    # 探测只能得到一个"碰巧成立"的结论（几乎必然是 ("char",4,False)），
    # 而它会被写进 _sem_cache 留给后面**含中文的行**用 —— 那些行按字符列算就
    # 少算全角格数，光标落在变量名中间（v94b / v95 两个 bug 同一个根因）。
    # 这里提前收工：调用方（`_indent_end_col`）遇到不含 Tab 的缩进本来就
    # 用不上语义信息，返回 None 不损失任何东西，还省掉一次"移动光标再还原"。
    if not ("\t" in line_text or any(_is_wide(c) for c in line_text)):
        return None
    try:
        info = _detect_semantics_passive(line_text, ec)
        if info is None:
            info = _probe_semantics(cm, cp, int(line_no), line_text)
    except Exception:
        return None
    return info


def _indent_end_col(indent, sem_info=None):
    """缩进【之后】那一列（1-based）——新行光标要停在这里。

    缩进只由空格与 Tab 组成、不含全角字符，因此：
      * 编辑器按【字符列】计时 col = len(indent) + 1；
      * 编辑器按【显示列】计时，Tab 要展开到下一个 tab stop。
    列语义是编辑器级属性（与 apply_completion 共用 _sem_cache）。

    ★v93：缩进里【含 Tab】时必须用显示列算。原先的实现是"缓存里没有语义
    信息就一律按字符列算"，而缩进全是【Tab】的工程恰恰最容易踩：

      * `_sem_cache["info"]` 只在 `_probe_semantics` 里写入，而那条路要
        "行内含 Tab / 全角且 ec 超过 len+1"才会走到 —— 新行是**编辑器自己
        刚插出来的**，本次会话可能一次都没触发过探测，缓存还是 None；
      * 于是 `\t` 缩进被算成 `len("\t")+1 = 2`，而显示列语义下正确值是
        5（Tab 展开到第 5 列）—— `SetSelection` 把光标落在 Tab 中间，
        用户接着打字就"歪了一格"。只在 Tab 缩进的工程、且本次会话还没
        碰过任何含 Tab / 中文的行时才复现，正是那种"偶尔"的口径。

    修法：含 Tab 却拿不到缓存时，**就地补一次探测**（调用方把 cm / cp /
    行位置通过 sem_info 传进来，或允许本函数直接读进程级缓存）。仍然拿不到
    就按显示列 + 默认 tab 宽度 4 算 —— 这是"最可能正确"的兜底，比按字符列
    算更贴近 VBE 实际行为（Tab 缩进 + 字符列语义的组合在真实 VBE 上不存在）。
    """
    if "\t" not in indent:
        # 全部是空格：两种语义结果完全相同，无需语义信息（最常见的一档）。
        return len(indent) + 1
    info = sem_info if sem_info else _sem_cache.get("info")
    if info:
        try:
            sem, tabw, wide2 = info
            if sem == "disp":
                return _disp_width(indent, tabw, wide2) + 1
            return len(indent) + 1
        except Exception:
            pass
    # 含 Tab 又没有语义信息：按显示列 + 默认 tab 宽度兜底（见 docstring）。
    return _disp_width(indent, 4, False) + 1


def _caret_to_line_end(cp, cm, line_no, line_text=None):
    """把光标移到 line_no 行的【行尾】—— 列号交给 VBE 自己夹，一行都不算。

    ★v94b 修的是用户报的"上一行含汉字时，删完空行光标落在上一行中间"：
    落点原本靠 `_eol_caret_col` 按【列语义】算，而列语义这套东西在这件事上
    根本不可靠 ——

      * 真 VBE 报的是**显示列**（全角占 2 格，v93 实测：`    For 每个元素 = 1 To 10`
        光标在字符列 18 时报的显示列是 22）；
      * 可列语义是**一次性缓存**的（_sem_cache），命中与否取决于"会话里第一次
        走到探测时碰巧是哪一行"。若那行是**纯 ASCII 且无 Tab**，VBE 报的 EOL 列
        恰好 == len+1，`_fit_semantics` 的结论就是 `("char", 4, False)` ——
        这句翻译过来是"VBE 用字符列"，可真相是"这一行上两种语义数值相同，
        **分不出来**"。缓存于是带着错结论留到下一次；
      * 等到上一行含汉字时，`_eol_caret_col` 按"字符列"给 len+1，VBE 却把它当
        **显示列**解释 => 少算【全角字符个数】格 => 光标落在中文中间；
      * 缓存为空 / 探测被判 broken 时兜底也是 `_disp_width(t, 4, False)`，
        同样按"全角占 1 格"算，症状一模一样。

    与其把这套语义继续猜下去，不如**问 VBE 自己**：`SetSelection` 给一个远超
    行尾的列，VBE 会把它夹回行尾（本文件 `_vbe_measure` 就是靠这个特性反推
    列语义的）。无论内部是字符列还是显示列、全角算 1 还是 2、Tab 多宽，落点
    都是 VBE 自己认的行尾 —— 不需要任何换算。

    返回 True 表示光标已经落在目标行行尾；False 表示这条路没走通（拿不到
    COM / 该编辑器不夹列），调用方按老口径兜底。
    """
    try:
        ln = int(line_no)
    except Exception:
        return False
    if line_text is None:
        try:
            line_text = str(cm.Lines(ln, 1)).rstrip("\r\n")
        except Exception:
            line_text = ""
    # 远超行尾：不管 VBE 内部怎么算，这个值都一定越过它的行长。
    far = len(line_text or "") + 512
    try:
        cp.SetSelection(ln, far, ln, far)
        sl, sc, el, ec = cp.GetSelection()
    except Exception:
        return False
    try:
        if int(sl) != ln or int(el) != ln:
            return False
        if int(sc) != int(ec) or int(sc) < 1 or int(sc) >= far:
            # sc >= far：该编辑器根本不夹列（WPS 的兼容编辑器有这种表现），
            # 光标会停在远超行尾的地方 -> 交给调用方按算出来的列兜底。
            return False
    except Exception:
        return False
    return True


def _eol_caret_col(line_text, sem_info=None):
    """行尾光标的【列】（1-based），按编辑器自己的列语义给（v94）。

    ⚠️★v94b：这只是**兜底**。首选路径是 `_caret_to_line_end`（让 VBE 自己夹列），
    因为列语义可能压根没探测出来、或探测时用的行恰好是纯 ASCII 而得出错误结论
    （详见 `_caret_to_line_end` 的说明）。只有那条路走不通时才退到这里。

    用途：删掉一个空行之后，光标要落到【上一行行尾】—— 与 VBE 原生"空行上按
    退格"的落点一致（原生是把换行符吃掉、两行并成一行，光标停在上一行末尾）。

    两条语义：
      * 字符列 -> len + 1；
      * 显示列 -> Tab 展开到 tab stop、全角按 wide2 计 2 格。
    行内既无 Tab 也无全角时两者相同 —— 绝大多数行走这一档，不需要语义信息。

    ⚠️ 只在行内真的有 Tab / 全角时才去取 sem_info：那条路会碰 COM（探测列语义
    时要移动光标再还原），能不碰就不碰。拿不到语义信息就按显示列 +
    (tabw=4, wide2=False) 兜底 —— 与 _indent_end_col 同一个口径（比按字符列算
    更贴近 VBE 实际行为）。
    """
    t = line_text or ""
    if not ("\t" in t or any(_is_wide(c) for c in t)):
        return len(t) + 1
    info = sem_info if sem_info else _sem_cache.get("info")
    if info:
        try:
            sem, tabw, wide2 = info
            if sem == "disp":
                return _disp_width(t, tabw, wide2) + 1
            return len(t) + 1
        except Exception:
            pass
    # ★v103：兜底口径必须与 `_eol_disp_col` 一致（全角占 2 格）。
    # 原来这里写的是 `_disp_width(t, 4, False)` —— 那是从 `_indent_end_col`
    # 抄来的口径，而 `_indent_end_col` 的输入**只有空白**（缩进），全角参数
    # 对它根本无意义；抄到一个"输入是任意代码行"的函数里，就变成"含中文的行
    # 少算格数"（`    说 = 1` -> 10，而真值是 11）⇒ 删空行后光标落在上一行中文
    # 中间偏左。同一语义（"行尾显示列"）在本文件里只能有一份实现，所以直接调它。
    return _eol_disp_col(t, 4)


# ---- 回车后的自动缩进（v79） ----
#
# 用户口径两条：
#   1) "我在一整行写入了注释，然后按回车换行后，光标会默认在行首、不会缩进；
#      帮我加个能忽略上方注释行（可能有多行注释）、跟上方第一个非注释行对齐的"；
#   2) "写 for / do / if / with / sub / function / type / enum 这些结构时，
#      回车能自动缩进"（块的开头行之后，新行要深一级）。
#
# 下面几个都是【纯函数】：不碰 COM，因此既能被单测直接钉住，也能在键盘钩子
# 线程里安全调用 —— 钩子线程不许碰 COM（v56 的教训），而"这一下回车要不要由
# 我们接管"必须在钩子线程里就定下来（否则要么吞了键才后悔，要么每次都吞）。

# 块结构【开头】：这一行往下的内容都属于新的一级，回车换行后新行要缩进一级。
#
# 判定只看两件事，不做语法分析：
#   * "以 Then 结尾" —— 覆盖 `If x Then` / `ElseIf x Then` / `#If VBA7 Then`；
#     单选 If（`If x Then y = 1`）以别的词结尾，天然被排除在外；
#   * "行首关键字在白名单里" —— Sub / Function / Property / Type / Enum /
#     For / Do / While / With / Else / Case / Select Case。
# 闭合行（Next / Loop / Wend / End If / End Sub …）刻意不在表里：它们本来就
# 缩进在正确的层级上，新行跟它们齐平就是对的结果。
_BLOCK_FIRST_WORDS = ("sub", "function", "property", "type", "enum",
                      "for", "do", "while", "with", "else", "case")
# 可以出现在声明关键字前面的修饰词，判断"行首关键字"时要跳过它们
# （`Private Sub Foo()` / `Public Function Bar()` / `Static Sub`）。
_MOD_FIRST_WORDS = ("public", "private", "friend", "static")
# 行尾就是 Then（前面得有分隔，别把 `Something` 认成 Then）
_RE_TAIL_THEN = re.compile(r"(?:^|\s)then\s*$")

# 块结构【收尾】（v79b 用户新要求：写 `If i = 1 Then` 按回车，后面自动补 `End If`）。
# 键是 _block_kind 认出来的规范化类型名，值是补在下一行（缩进回到块头同级）的收尾。
#   * 只有"真需要收尾"的块才在这儿 —— Else / ElseIf…Then / Case 是块【内】的分支，
#     往下一行只是同级的分支体，不该替用户写收尾；
#   * For 用 Next（VBA 允许不带变量名）、Do 用 Loop、While 用 Wend；其余统一 End Xxx；
#   * 条件编译（#If / #Else）刻意不在表里：它必须配 `#End If`，替用户写个裸的
#     `End If` 反而把代码写错，宁可让用户自己敲。
_BLOCK_CLOSERS = {
    "sub": "End Sub",
    "function": "End Function",
    "property": "End Property",
    "type": "End Type",
    "enum": "End Enum",
    "if": "End If",
    "with": "End With",
    "select": "End Select",
    "for": "Next",
    "do": "Loop",
    "while": "Wend",
}


def _code_part(line_text):
    """一行代码里"去掉字符串与注释之后"的部分（行尾空白也去掉）。纯函数。

    用 parser 里那份掩码（字符串字面量 + `'` 注释换成等长空格），因此
    `If x = "a Then" Then` 里字符串中的 Then 不会被误当成块头，
    `If x Then  '注释` 尾部的注释也不会破坏"以 Then 结尾"的判定。
    整行 `Rem ...` 注释按空串处理。
    """
    t = line_text or ""
    if not t.strip():
        return ""
    _st = t.lstrip().lower()
    if _st == "rem" or _st.startswith("rem ") or _st.startswith("rem\t"):
        return ""
    try:
        return vba_parser._mask_strings_and_comments(t).rstrip()
    except Exception:
        return t.rstrip()


def _is_comment_only(line_text):
    """整行是不是注释（`'...` 或 `Rem ...`）—— 一行里没有任何代码。纯函数。"""
    t = (line_text or "").strip()
    if not t:
        return False
    if t.startswith("'"):
        return True
    low = t.lower()
    return low == "rem" or low.startswith("rem ") or low.startswith("rem\t")


# 行标签：行首是"一个词"（标识符或数字）紧跟冒号 —— `1:` / `Retry:` /
# `10: x = 1`（标签与语句同一行）。冒号前不是"一个词"就不算：
# `Case 1:` / `Dim x: x = 1` / `foo(1)` 都不匹配。
_RE_LABEL = re.compile(r"^\s*(?:[A-Za-z_\u4e00-\u9fff]\w*|\d+)\s*:")


def _is_label_line(line_text):
    """这一行是不是【行标签】（`1:` / `Retry:` / `10: x = 1`）？纯函数。

    ⚠️ 为什么单独认它（v86，用户报的）：标签是用 `GoTo 1` / `On Error GoTo 1`
    跳过来的落点，VBA 里的**惯例写法就是【顶格】**（`脚本.xlsm` 的 `Sub 功能2()`
    里就有一个顶格的 `1:`）—— 它的缩进跟所在嵌套层级【毫无关系】。拿缩进推结构
    的扫描看到它就会以为"已经出了这一层"，把下面真正的 `End If` / `End Sub`
    当成不存在 ⇒ 回车再补一个收尾（用户报的正是这个）。

    字符串与注释先被 `_code_part` 掩掉，所以 `"a:b"` / `'x:y` 不会误判。
    """
    code = _code_part(line_text)
    if not code:
        return False
    return _RE_LABEL.match(code) is not None


def _indent_blind(line_text):
    """这一行的缩进【不代表结构层级】（向下扫描时整行跳过）。纯函数。

    用户的 VBA 里这两类行很常见，而惯例写法与层级无关：
      * 行标签 `1:` / `Retry:`（见 `_is_label_line`）；
      * `#` 条件编译指令 `#If` / `#End If` / `#Const`… —— 同样惯例顶格，
        而且 `#If` 块压根不参与"本块结束没结束"（v82 只跳了 `#Else`/`#ElseIf`）。

    整行注释与块内分支行由调用处一并判 —— 那两条是 v80a/v82 的老口径。
    """
    return _is_label_line(line_text) or _directive_word(line_text) is not None


def _block_kind(line_text):
    """这一行是【哪一种】块结构的开头？返回规范化的类型名；不是块头返回 None。

    返回值的三种含义：
      * `"branch"`    —— Else / ElseIf…Then / Case：要缩进一级，但**没有收尾**；
      * `"directive"` —— #If / #Else 条件编译：要缩进一级，收尾必须由用户写
                        `#End If`，我们不代劳（见 _BLOCK_CLOSERS 的说明）；
      * 其余（sub / function / property / type / enum / if / with / select /
        for / do / while）—— 都在 _BLOCK_CLOSERS 里有对应的收尾行。
    """
    code = _code_part(line_text)
    if not code:
        return None
    if code.endswith("_"):
        # 续行符：这条语句还没写完，下一行是它的延续，不是"块里的第一行"。
        return None
    low = " ".join(code.lower().split())
    hashed = low.startswith("#")
    words = low.split(" ")
    i = 0
    while i < len(words) and words[i] in _MOD_FIRST_WORDS:
        i += 1
    first = words[i].lstrip("#") if i < len(words) else ""
    if first in ("else", "elseif", "case"):
        if first == "elseif" and not _RE_TAIL_THEN.search(low):
            return None                    # 只写 `ElseIf x` 没写 Then —— 不算块头
        return "directive" if hashed else "branch"
    if _RE_TAIL_THEN.search(low):
        # `If x Then` / `ElseIf x Then` / `#If VBA7 Then`：行尾就是 Then。
        # 单选 If（`If x Then y = 1`）以别的词结尾，天然被排除在外。
        return "directive" if hashed else "if"
    if first == "select":
        ok = len(words) > i + 1 and words[i + 1] == "case"
        return ("directive" if hashed else "select") if ok else None
    if first in _BLOCK_FIRST_WORDS:
        return "directive" if hashed else first
    return None


def opens_block(line_text):
    """这一行是不是"块结构的开头"——回车换行后新行该缩进一级。纯函数。

    例：`For i = 1 To 10` / `For Each c In rng` / `Do While x` / `If x Then` /
    `ElseIf y Then` / `Else` / `Case 1` / `Select Case x` / `With rng` /
    `Sub Foo()` / `Private Function Bar() As Long` / `Property Get X` /
    `Type Foo` / `Enum Color` / `#If VBA7 Then` / `#Else` -> True；
    `If x Then y = 1`（单选 If）、`Next i`、`End If`、`Loop`、`Wend`、
    `x = 1`、`Exit For`、`DoEvents`、`foo(1, 2)`、`Debug.Print x`、
    以续行符 `_` 结尾的行 -> False。
    """
    return _block_kind(line_text) is not None


def block_closer(line_text):
    """块头那行【该补的收尾】；不需要收尾时返回 None。纯函数。

    `Sub Foo()` -> "End Sub"、`If x Then` -> "End If"、`With rng` -> "End With"、
    `Select Case x` -> "End Select"、`For i = 1 To 10` -> "Next"、`Do` -> "Loop"、
    `While x` -> "Wend"、`Private Function F()` -> "End Function"；
    `ElseIf y Then` / `Else` / `Case 1` / `#If VBA7 Then` / 整行注释 /
    普通行 -> None（它们要么是块内的分支、要么该由用户自己写收尾）。
    """
    return _BLOCK_CLOSERS.get(_block_kind(line_text))


def is_block_closer(line_text):
    """这一行本身是不是某个块的【收尾行】（`End If` / `Next i` / `Loop`…）。纯函数。

    ★v122 新增。与 `block_closer` 是一对镜像、问的是**相反**的问题：
      * `block_closer("For i = 1 To 3")` -> `"Next"`（块头 -> 它欠的收尾）；
      * `is_block_closer("Next i")` -> True（这一行是不是收尾）。

    只用来判一件事：在**收尾行上方**插一行时，新行落在块**内**（要比收尾深
    一级）。`End Sub` 上方插行同理 —— 那是过程体的最后一行，不是模块级。

    ⚠️ 别拿 `block_closer` 代替它：`block_closer("    Next i")` 返回 None
    （`Next` 不是块头），看着像"不是收尾"，正好反了。v122 第一版就这么写的，
    被 §108.2 当场照出来。
    """
    for _c in _BLOCK_CLOSERS.values():
        if _is_closer_text(line_text, _c):
            return True
    return False


def _indent_width(raw):
    """行首空白占几个字符位（Tab 按 4 列算）。纯函数。"""
    n = 0
    for ch in (raw or ""):
        if ch == " ":
            n += 1
        elif ch == "\t":
            n += 4
        else:
            break
    return n


def _is_closer_text(line_text, closer):
    """这一行是不是 closer 那一行（`End If` / `Next i` / `Loop While x`…）。纯函数。

    先用 _code_part 去掉注释：`Next  '下一个` 也算 `Next`。
    """
    if not closer:
        return False
    s = _code_part(line_text).strip().lower()
    if not s:
        return False
    want = closer.strip().lower()
    return s == want or s.startswith(want + " ")


def _closer_owner_above(lines_above, kind, indent):
    """上方有没有一个【同类、仍未闭合】的块头，正好缩进在 indent 列？（纯函数）

    lines_above: 块头行【上方】各行的文本（越靠近它的越靠后）
    kind:        _block_kind 认出来的类型名（for / if / with / do / while /
                 select / type / enum / sub / function / property）
    indent:      待判定的收尾行缩进宽度

    只干一件事：回答"下面那个更浅的同类收尾，是外层块（嵌套的 For/If/Do/With…）
    的还是本块的"。从近到远扫，用"同类收尾 +1 / 同类块头 -1"的括号配对法，
    遇到一个"还开着"的同类块头时，看它缩进是不是正好等于那个收尾的缩进
    —— 是，就说明那个收尾属于它。

    为什么需要（v79c，用户报的双层 For）：内层 `For` 下面压着的是【外层】的
    `Next`，缩进比内层浅。单看缩进"比块头浅就补"是对的，但用户自己把收尾写歪
    一级的情形也长这样 —— 再补一条就成 `Next` 双份（编译报错）。加上这一问，
    两种情形就分开了：外层真有一个同类块头在那一列 -> 补；没有 -> 不补。
    """
    closer = _BLOCK_CLOSERS.get(kind or "")
    if not closer:
        return False
    depth = 0
    for raw in reversed(list(lines_above or ())):
        if not (raw or "").strip():
            continue
        if _is_closer_text(raw, closer):
            depth += 1                         # 更近处已经闭合了一层
            continue
        if _block_kind(raw) != kind:
            continue
        if depth == 0:
            return _indent_width(raw) == int(indent)
        depth -= 1
    return False


def _own_closer_below(lines_below, closer, kind, base):
    """下面是不是已经挂着【本块自己的】收尾？（v80a，纯函数）

    lines_below: 光标行下方各行（越靠近它的越靠前）
    closer:      本块的收尾文本（`Next` / `End If` / `End Type`…）
    kind:        `_block_kind` 认出来的块头类型
    base:        块头那行的缩进【宽度】

    与 `_closer_owner_above` 正好对称的一问：那一问管"下方第一个非空行是
    【外层块】的收尾"，这一问管"下方压着的那些同级代码里，到底有没有本块的
    收尾"。只看【和块头同一缩进】的行，用"同类块头 +1 / 同类收尾 -1"配平：

      * 同缩进的同类块头（`For` 里嵌 `For`、用户还没缩进）-> 计一层，它的
        `Next` 只闭合它自己；
      * 配平到 0 时遇到的同类收尾 -> 就是【本块】的，下面已经挂着了；
      * 同缩进的其它东西（`If` / `With` / `x = 1` / 别的收尾）-> 本块早就
        结束了（那一行就是它的边界），收尾不在下面；
      * 缩进比块头浅的行 -> 已经出了这一层，收尾不在下面；
      * 缩进更深的行 -> 属于嵌套块，与"本块的收尾在哪"无关，跳过；更深的
        同类块头 / 收尾另用 deeper 配平（免得把内层的 `Next` 当成"用户写歪
        的收尾"而少补一条）；
      * 空行 / 整行注释 / 块内分支行（`Else` / `ElseIf…Then` / `Case…`，
        含 `#Else` / `#ElseIf…Then`）/ 行标签（`1:`）/ `#` 条件编译指令
        -> 都不是"本块到此为止"的边界，跳过（它们的缩进不含结构信息，
        见 `_indent_blind`）。

    为什么要这一问（v80a，用户报的）：先写了一个 `For`（回车自动补了 `Next`），
    再回到它【上面】补一个新的 `For` 并回车 —— 新 `For` 下面压着的正是原来那个
    `For`（同级）。旧口径"同级一律不补（只有模块级缩进 0 才补）"，于是新 `For`
    的 `Next` 永远补不出来，用户看到的就是"只换行缩进、不补 `Next`"。

    ★v82（用户报的"`If…Then` 补过一次 `End If`，写完 `Else` 再回到 `Then` 那行
    回车又补一个"）：`Else` / `Case` 这类**分支行**本来就在本块【里面】，可 v81
    起分支行会被拉回块头那一带（`Else` 与 `If` 齐平；`Case` 从 v83 起比 `Select
    Case` 深一级）—— 齐平的那种长得跟"同级的别的代码"一模一样，旧判据一眼认定
    "本块已结束了"，收尾就当没看见，回车再补一条。注释同理（注释根本不参与块
    结构）。这两类行现在一律跳过，扫描继续往下走 —— 跳过是**先跳后判缩进**，
    所以分支行放在哪一级都不影响这一问（第 63.8 节钉着）。
    ★v86（用户报的"过程里写了很多代码之后，回到 `Sub` 那行、或回到 `Then` 那行
    按回车，又补出一个 `End Sub` / `End If`"）：v80a/v82 那两跳还是不够 ——
    **行标签**（`1:`，配 `GoTo 1` 用，惯例顶格写）和 **`#` 条件编译指令**
    （`#If` / `#End If`，同样惯例顶格）的缩进不代表层级，却照样被当成了
    "同级的别的代码 / 更浅的行"⇒ 下面明明挂着的 `End If` / `End Sub` 全被
    无视 ⇒ 再补一条。现这两类行也整行跳过（判据见 `_indent_blind`）。
    `脚本.xlsm / 模块2` 里就是活样本：`Sub 功能2()` 里那个顶格的 `1:` 一次
    坑掉三处（`Sub` 行、`For` 行、`If … Then` 行各一次）。
    ⚠️ 这一跳必须在缩进判定【之前】（第 63.8 节钉着）。
    """
    depth = 0        # 同缩进的同类块头（本块下面嵌着的同类块）
    deeper = 0       # 真正嵌套（缩进更深）的同类块头
    for raw in (lines_below or ()):
        if not (raw or "").strip():
            continue
        if _is_comment_only(raw) or _branch_word(raw) is not None \
                or _indent_blind(raw):
            # ★v82：注释 / 块内分支行都不是"本块到此为止"的边界（详见 docstring）。
            # ★v86：行标签（`1:`）与 `#` 条件编译指令同理 —— 缩进不含结构信息。
            continue
        w = _indent_width(raw)
        if w < base:
            return False                       # 出了这一层：本块收尾不在下面
        if _is_closer_text(raw, closer):
            if w == base:
                if depth == 0:
                    return True                # 这个收尾就是本块的
                depth -= 1
            elif deeper > 0:
                deeper -= 1                    # 闭合的是更深的那个同类块
            else:
                # 孤零零一个更深的同类收尾：位置不对，但它是用户自己写的 ——
                # 保守起见不再补（宁可少补，也不要 `Loop` / `Next` 双份把代码
                # 写坏）。这与 v79b 的老口径一致。
                return True
            continue
        if _block_kind(raw) == kind:
            if w == base:
                depth += 1                     # 同缩进的同类嵌套块头
            else:
                deeper += 1
            continue
        if w == base:
            return False                       # 同缩进的其它东西 -> 本块已结束
    return False


def closer_needed(lines_below, closer, base_indent, kind=None, lines_above=(),
                  truncated=False):
    """下面到底要不要补 closer（v79b；嵌套判定 v79c；同级判定 v80a，纯函数）。

    lines_below: 光标行【下方】各行的文本（越靠近它的越靠前）
    closer:      block_closer 算出来的收尾文本（空 / None -> 不补）
    base_indent: 块头那行的行首缩进【宽度】
    kind:        块头类型（可选；给了才能判"更浅的同类收尾是不是外层块的"）
    lines_above: 块头行上方各行（可选，同上）
    truncated:   `lines_below` 是不是【被读的行数上限截断】的（v86，可选）。
                 真时"窗口内没找到收尾"不能当"没有" —— 见下面的 ★v86。

    传了 kind（生产路径上一定会传）时的口径：
      * 下方【没有】本块自己的收尾 -> 补（怎么算"没有"，见 _own_closer_below：
        它按缩进分类扫下方各行，同级同类块头配平、同级别的东西算边界、
        空行 / 注释 / 分支行跳过、更深的行跳过）；
      * 下方第一个非空行就是本块的同类收尾、且缩进比块头【更浅】-> 再问一句
        `_closer_owner_above`："那是外层块的收尾（本块该补）还是用户写歪的
        （不补）"；
      * 其余 -> 不补。

    不传 kind 时保持 v79b 老口径（只看"下方第一个非空行"）：
      * 它比块头缩进得【更浅】-> 补（但若它就是本块的收尾且更浅，一律不补）；
      * 它跟块头齐平或更深 -> 不补；唯一的例外是【块头在模块级（缩进 0）】，
        那里"齐平"的只可能是下一个声明/过程头，收尾该补；
      * 下方根本没有非空行（文件到底了）-> 补。
    空行一律跳过（下方可能隔着一堆空行才挂着真正的收尾）。

    不传 kind/lines_above（旧式调用）时，"更浅的同类收尾"一律按【不补】处理 ——
    这是最保守的一档，与 v79b 完全一致。

    ★v80a 变更：传了 kind 时，"同级"一档不再是"只有模块级（缩进 0）才补"，
    改成问一句 `_own_closer_below` —— "下面到底有没有本块自己的收尾"。
    别的分档一律不动。

    ★v82 变更（不再多补一条收尾）：`_own_closer_below` 现在把【整行注释】和
    【块内分支行】也跳过 —— `Else` / `ElseIf…Then` / `Case…` 是块内的分支，
    不是"本块已结束"的边界。旧判据把它们当边界，于是 `If…Then` 补过一次
    `End If`、写完 `Else` 后再回到 `Then` 那行回车，会【再补一个】`End If`
    （用户报的正是这个；`Select Case` 里的 `Case` 完全同理）。这条只让"补"
    更保守，别的分档一字未动。
    """
    if not closer:
        return False
    base = int(base_indent)
    first = None
    for raw in (lines_below or ()):
        if (raw or "").strip():
            first = raw
            break
    if kind is not None:
        # 【v80a 口径】下面挂着本块自己的收尾 -> 不补；没有 -> 补。
        # 这一问把"同级代码"分成三类：同类嵌套块头（配平）、本块的收尾、
        # 以及"本块早就结束了"的边界 —— 详见 _own_closer_below。
        if _own_closer_below(lines_below, closer, kind, base):
            return False
        if truncated:
            # ★v86：行数是【读满上限被截断】的 ⇒ 本块的收尾可能就在窗口之外，
            # "在窗口里没找到"不等于"下面没有"。这一档按"别再补"处理：宁可少补
            # 一条（用户自己敲一下就好），也不要凭"看不到"变出【双份收尾】把
            # 代码写坏 —— `End Sub` / `End If` 双份是直接编译不过的
            # （用户报的"过程里写了很多代码，回到 Sub 行回车又补一个 End Sub"
            #  就有一半是这么来的：窗口只有 200 行，收尾在 200 行之外）。
            return False
        if first is not None:
            w = _indent_width(first)
            if _is_closer_text(first, closer) and w < base:
                # 更浅的同类收尾：可能是【外层块】的（双层 For 的 `Next`），
                # 也可能是用户自己写歪的 —— 分这一下的活交给它（v79c）。
                return _closer_owner_above(lines_above, kind, w)
        return True
    # 【老口径】kind 没传（v79b 那套调用方式）：一个字不动，第 58.9 节钉着。
    if first is None:
        return True
    w = _indent_width(first)
    if _is_closer_text(first, closer):
        if w >= base:
            return False                       # 本块的收尾已经挂在下面了
        # 比块头浅：要么是外层块的收尾（本块该补），要么是用户写歪的
        return _closer_owner_above(lines_above, kind, w)
    if w != base:
        return w < base
    # 同级、且不是本块的收尾：模块级 -> 本块还没内容，补；过程内 -> 维持
    # v79b 老口径（不补），免得把用户正在写的块体顶开。
    return base == 0


def next_line_indent(cur_line, lines_above=(), unit="    "):
    """回车换行后，新行该有的行首缩进（纯函数）。

    cur_line:    光标所在行的文本
    lines_above: 它上方各行的文本（越靠近它的越靠后）
    unit:        一级缩进（默认 4 个空格）

    规则（用户 v79 口径）：
      * 参考行 = cur_line 自己；但当它是【整行注释】或空行时，往上找第一个
        既不是整行注释也不是空行的行 —— 这就是"忽略上方注释行（可能有多行），
        跟上方第一个非注释行对齐"；
      * 新行缩进 = 参考行的行首空白；参考行若是【块结构开头】，再深一级；
      * 连参考行都找不到（首行 / 上方全是注释与空行）→ 就用 cur_line 自己的
        行首空白，也就是 VBE 的原生行为 —— 绝不比原来更差。
    """
    ref = None
    if (cur_line or "").strip() and not _is_comment_only(cur_line):
        ref = cur_line
    else:
        for t in reversed(list(lines_above or ())):
            if (t or "").strip() and not _is_comment_only(t):
                ref = t
                break
    if ref is None:
        return _leading_ws(cur_line)
    base = _leading_ws(ref)
    return base + unit if opens_block(ref) else base


# ── v81：块【内】分支行自动对齐（Else / ElseIf…Then / Case… / #Else / #ElseIf）
# 用户口径："我在写了一行 if 之后，回车会自动缩进和补充 end if，然后我再写 else，
# 此时的 else 是缩进状态，不能自动对齐上面的 if。我想在我打完 else，或 elseif
# （if 里的多分支）再按回车后，else / elseif 能够自动对齐上方的 if。其他类似的结构，
# 比如 select case，里面可以多分支的，都帮我做成这样能自动对齐的。"
#
# 为什么会歪：`Else` 是 _block_kind 认的块头（要缩进一级），所以这一下回车被我们
# 接管了 —— 而 VBE 原生回车【会】把刚敲完的那一行拉回它该在的层级，我们一接管，
# 那个自动对齐就没机会跑。于是 `Else` 一直停在"块体那一级"（上面自动补完 End If 后
# 光标停的那一行）。这里补的正是这一下：回车时先把分支行拉回它该在的层级，
# 再让新行深一级 —— 就是 VBE 原生会做的那件事。
#
# 归属关系在 VBA 里没有歧义：Else / ElseIf…Then 属于最近的还开着的 `If … Then`，
# Case… 属于最近的还开着的 `Select Case`，#Else / #ElseIf 属于最近的还开着的 `#If … Then`。
_BRANCH_OWNER = {"else": "if", "elseif": "if", "case": "select"}

# ★v83：分支行对齐到【块头往下第几级】。归属者一样，但层级不一样 ——
#     If x Then
#         ...
#     Else                      <- 与 `If` 【齐平】（0）
#     End If
#     Select Case x
#         Case 1                <- 比 `Select Case` 【深一级】（1）
#             ...
#     End Select
# 这是 VBA 的惯用缩进，也正是 VBE 原生回车会整理成的样子。v81 把 `Case` 也拉到
# `Select Case` 同级了（用户报的："case 应该是相对于 select 缩进的"）；`#Else` /
# `#ElseIf` 与 `#If … Then` 齐平、不分深浅，所以条件编译这一档同样是 0。
_BRANCH_OWNER_DEPTH = {"else": 0, "elseif": 0, "case": 1}
# 条件编译的收尾。_BLOCK_CLOSERS 里刻意没收 `#If`（那个收尾必须由用户自己写），
# 但"往上找还开着的 #If"这一步照样要用到它，所以单独放在这里。
_DIRECTIVE_CLOSER = "#End If"


def _directive_word(line_text):
    """`#` 开头的条件编译指令的第一段（不含 `#`）：`#If VBA7 Then` -> 'if'。纯函数。"""
    code = _code_part(line_text)
    if not code:
        return None
    low = " ".join(code.lower().split())
    if not low.startswith("#"):
        return None
    rest = low[1:].strip()
    return rest.split(" ")[0] if rest else None


def _branch_word(line_text):
    """这一行是哪种【块内分支】？返回 'else' / 'elseif' / 'case'；不是 -> None。纯函数。

    与 _block_kind 的 'branch' 一档严格对应（`#` 开头的条件编译分支返回同一个词，
    调用方再按"有没有 `#`"决定去找哪类块头）：
      * `Else` / `Case 1` / `Case Else` / `Case Is > 3` -> 'else' / 'case'；
      * `ElseIf y Then`          -> 'elseif'（少了 Then 不算，与 _block_kind 同口径）；
      * `Else If y Then`         -> 'else'   （VBA 里这是"Else + 一行新 If"，不是 ElseIf）；
      * `End If` / `Next` / 普通行 / 续行（行尾 `_`）-> None。
    """
    code = _code_part(line_text)
    if not code:
        return None
    if code.endswith("_"):
        return None
    low = " ".join(code.lower().split())
    words = low.split(" ")
    i = 0
    while i < len(words) and words[i] in _MOD_FIRST_WORDS:
        i += 1
    first = words[i].lstrip("#") if i < len(words) else ""
    if first == "elseif":
        return "elseif" if _RE_TAIL_THEN.search(low) else None
    if first in ("else", "case"):
        return first
    return None


def _unclosed_owner_above(lines_above, kind):
    """上方最近一个【还开着】的 kind 类块头在哪？返回它的行首空白；没有 -> None。纯函数。

    lines_above: 当前行【上方】各行（越靠近它的越靠后）
    kind:        'if' / 'select' / 'directive'（_block_kind 的口径）

    括号配对法（与 _closer_owner_above 同一套思路，只是方向相反）：从近到远扫，
    收尾（`End If` / `End Select` / `#End If`）计 +1，同类的块头计 -1；扫到
    "count 已经是 0 的同类块头"，它就是这一档分支的归属者。

    ⚠️ 为什么 `ElseIf` / `Else` 不算块头：它们和 `If` 共用同一个 `End If`，要是也
    计一层，配平永远对不上（`If` / `Else` / `End If` 会配成 -2）。`_block_kind` 对
    它们返回的是 'branch' 而不是 'if'，天然被跳过 —— 这正是这里必须用 `_block_kind`
    而不是"行首是不是 If"来判的原因。
    """
    closer = _DIRECTIVE_CLOSER if kind == "directive" else _BLOCK_CLOSERS.get(kind)
    if not closer:
        return None
    depth = 0
    for raw in reversed(list(lines_above or ())):
        if not (raw or "").strip():
            continue
        if _is_closer_text(raw, closer):
            depth += 1
            continue
        if kind == "directive":
            # 条件编译：只有 `#If … Then` 开一层；`#Else` / `#ElseIf` 是兄弟分支，
            # 别的指令（`#Const` / `Option …` 之类）一律不看。
            if _directive_word(raw) != "if":
                continue
        elif _block_kind(raw) != kind:
            continue
        if depth == 0:
            return _leading_ws(raw)
        depth -= 1
    return None


def branch_align(line_text, lines_above, unit="    "):
    """块内分支行该对齐到哪里（v81；层级按分支分档 v83，纯函数）。

    line_text:   光标所在行的文本（正在写的那一行，光标在行尾）
    lines_above: 它上方各行的文本（越靠近它的越靠后）
    unit:        一级缩进

    返回 `(本行应有的行首空白, 新行应有的行首空白)`；不是分支行、或上方找不到它
    所属的块头（`If x Then y = 1` 这种单选 If、`Else` 敲在 `If` 前面、`Case` 敲在
    `Select Case` 前面…）-> 返回 None，调用方维持原行为（新行 = 本行缩进 + 一级）。
    **绝不猜**：找不到归属者就当没这回事，宁可不对齐也不要挪错。

    层级由 `_BRANCH_OWNER_DEPTH` 定（v83）：
      * `Else` / `ElseIf … Then` -> 与 `If … Then` 齐平，新行深一级；
      * `Case …` / `Case Else` / `Case Is …` -> 比 `Select Case` **深一级**，
        新行（块体）再深一级；
      * `#Else` / `#ElseIf … Then` -> 与 `#If … Then` 齐平，新行深一级。
    缩进单位是 `"    "` 时就是 0 / 4 / 8 格与 4 / 8 / 12 格的区别；用 Tab 或
    2 空格缩进的工程跟着 `unit` 走（`unit * 1`）。
    """
    word = _branch_word(line_text)
    if not word:
        return None
    hashed = _code_part(line_text).lstrip().startswith("#")
    kind = "directive" if hashed else _BRANCH_OWNER.get(word)
    if not kind:
        return None
    owner = _unclosed_owner_above(lines_above, kind)
    if owner is None:
        return None
    mine = owner + unit * _BRANCH_OWNER_DEPTH.get(word, 0)
    return mine, mine + unit


def _indent_unit_for(lines):
    """一级缩进用什么字符串：默认 4 个空格；这个工程里用 Tab 缩进就跟着用 Tab。

    可用 `set VBECOMPLETE_INDENT_UNIT=2`（2 个空格）或 `=tab` 覆盖，
    也可以在 settings.ini 里写 `INDENT_UNIT = tab`。
    """
    try:
        raw = _settings.text("INDENT_UNIT", "")
    except Exception:
        raw = ""
    if raw:
        if raw.lower() in ("tab", "\\t", "\t"):
            return "\t"
        if raw.isdigit():
            return " " * max(1, min(16, int(raw)))
    for t in lines or ():
        if "\t" in _leading_ws(t):
            return "\t"
    return "    "


def _eol_disp_col(line_text, tabw=4):
    """行尾光标应该报告的【显示列】（1-based）。

    与 VBE 的列语义一致：全角字符占 2 格，Tab 展开到下一个 tab stop。
    纯函数，不碰 COM —— 供 enter_indent_wanted 在钩子线程里用。
    """
    col = 1
    for ch in line_text:
        if ch == "\t":
            col += tabw - (col - 1) % tabw
        elif _is_wide(ch):
            col += 2
        else:
            col += 1
    return col


def enter_indent_wanted(line_text, caret_ec):
    """钩子线程用的纯判据：这一下回车值不值得由我们接管（v79）。

    只有两种情形才接管，其余一律放行给 VBE 原生回车：
      * 整行注释（要忽略它、跟上方第一个非注释行对齐）；
      * 块结构开头（新行要深一级）。
    另有五条保守门槛（拿不准就不抢回车）：
      * 空行 / 只有空白：交给 VBE（它本来就继承上一行的缩进）；
      * 行内有 Tab：列语义算不准，不拿"光标是否在行尾"去赌；
      * 光标不在行尾：那是"拆行"，必须让 VBE 原生处理；
      * caret_ec 拿不到（0）：不接管；
      * ★v93 行内含【全角字符】：caret_ec 是 VBE 的【显示列】，全角占 2 格，
        直接拿它和 len(t)（字符数）比会算错 —— 见下面的详细说明。

    ★v93 修的是"中文行上误吞回车"：原判据是 `int(caret_ec) < len(t) -> 放行`，
    也就是"显示列 ≥ 字符数"就当光标在行尾。纯 ASCII 行两者相等、没问题；
    可中文行里全角让显示列【虚高】，行中间的光标也能满足这个条件，于是本该
    由 VBE 原生处理的【拆行】被我们接管，用户看到"回车没拆行 / 多出一行"。

    实测（_probe_v93n2）：`    For 每个元素 = 1 To 10`（22 字符、行尾列 23）
    光标停在字符列 18 时显示列是 22 ≥ 22 → 旧判据误判"在行尾"并接管。

    修法刻意**不用**"显示列 -> 字符列"的完整换算（那需要 tabw / wide2 的
    列语义探测，而钩子线程不许碰 COM）：改成拿 eol 的【显示列】当靶子，
    要求两边完全相等。行内含全角时"恰好等于行尾显示列"的位置是唯一的
    （每个字符的显示宽度都 ≥ 1，列值单调递增），不会误判；
    而"行尾"这个语义正是我们唯一要认的位置。判据仍然全是纯文本，
    在钩子线程里可安全执行。
    """
    t = line_text or ""
    if not t.strip():
        return False
    if "\t" in t:
        return False
    try:
        ec = int(caret_ec or 0)
    except Exception:
        return False
    if ec <= 0:
        return False
    # ★代码复查修（2026-10-18）：纯 ASCII 分支原来写的是 `ec < len(t)`，
    # 边界上差一格 —— 光标停在**最后一个字符之前**时 VBE 报的就是 `len(t)`
    # （字符列 = 显示列 = 1-based），这个式子判不出来 ⇒ 那一下本该由 VBE
    # 处理的【拆行】被我们接管，表现为"回车没拆行 / 缩进被改写"。
    # 现在两支合成一条：行内已排除 Tab，`_eol_disp_col` 对纯 ASCII 就返回
    # `len+1`、对全角行返回真实的行尾显示列 —— "行尾"这个语义只有一份算法
    # （与 v93 修全角虚高时用的是同一个函数，MEMORY 第 21 条）。
    if ec != _eol_disp_col(t):
        return False
    return bool(_is_comment_only(t) or opens_block(t))


def blank_line_delete_wanted(line_text, caret_ec=None, tabw=4):
    """钩子线程用的纯判据：这一下退格要不要由我们接管 —— 整行【只有空白】就
    删掉整行（v94）。

    用户口径：「按退格的时候，如果整行是空行，能删除整行，IDEA 里写 Java 就有
    这个功能」。

    需要我们的正是 VBE 自己按回车继承缩进造出来的那种行（`    `）：原生退格
    一次只吃掉**一个**空格，要按 4 下才空、第 5 下才并到上一行去。

    三道门槛（拿不准就一律 False，放行给 VBE，与 enter_indent_wanted 同口径）：
      * 【真空行】（`""`）不接管 —— 光标必然在第 1 列，VBE 原生退格就是"并到
        上一行"，结果和我们删整行**一模一样**，没必要抢（多一次 COM 往返、
        也多一处出错的可能）；
      * 【有内容的行】不接管 —— 那是普通删字，一个字都不许碰；
      * 光标必须在【这行空白的末尾】。停在缩进中间说明用户多半是想改缩进量，
        那一档照旧交给 VBE 一次删一格。

    ⚠️ 光标列两种语义都认（`ec == len+1` 或 `ec == 行尾显示列`）：VBE 报告的是
    **显示列**，Tab 缩进的行在两种语义下数值不同，只认一种会让整条功能在 Tab
    缩进的工程里静默失效。两种口径都指向"行尾"这同一个位置，不会误判（一个
    Tab 内部的列值 VBE 根本不会报出来）。
    """
    t = line_text or ""
    if not t or t.strip():
        return False
    try:
        ec = int(caret_ec or 0)
    except Exception:
        return False
    if ec <= 0:
        return False
    return ec == len(t) + 1 or ec == _eol_disp_col(t, tabw)


# ---- v110：Ctrl+↑ / Ctrl+↓ 跳到声明区末尾 / 过程末尾的空行 ----
#
# 用户口径（原话，2026-10-15）：
#   "当我按 ctrl+↑，光标会自动定位到当前过程中最后一个 dim 关键字行，如果
#    此行的下一行不是空行，就自动在此行下方新增一行，光标移动到下方空白行；
#    如果此行的下一行是空白行，就不用新增行，光标移动到下方空白行就可以了。"
#   "当我按 ctrl+↓，光标会自动定位到当前过程中的 end sub 或 end function 行，
#    如果此行的上一行不是空白行，就自动在此行上方新增一行，光标移动到上方
#    空白行；如果此行的上一行是空白行，就不用新增行……"
#
# 下面一族都是【纯函数】：不碰 COM，因此既能被单测直接钉住，也能在键盘钩子
# 线程里安全调用（钩子线程不许碰 COM —— v56 的教训）。后端
# `VbeBackend.ctrl_nav_jump` 与这里是【同一份实现】：后端只是把"此刻真读到
# 的模块文本"喂进同样的纯函数（v87 的规矩：同一条语义只能有一份实现）。
#
# ⚠️ 每个判据都收【两份行表】：`lines`（抹白后的，跑正则）与 `raw`（原始的，
# 判续行）。为什么必须分两份：`_mask_strings_and_comments` 把一整段注释压成
# 【一个空格】，`' 注释 _` 行尾那个续行符 `_` 跟着注释一起被抹掉了 —— 只拿
# 抹白行判续行，就认不出"下一行是注释的续行"，于是把注释里那行 `Dim x` 当成
# 真声明，跳过去之后还会在注释中间插一行、把注释从中间拆断。

# 过程收尾行：只认 Sub / Function / Property。
# ⚠️ 不能直接用 parser._RE_ANY_PROC_END（那个还含 Type|Enum）：Type / Enum 块
# 在【模块声明区】、不在过程里，把 `End Type` 当过程尾会在声明区里乱跳。
_RE_V110_PROC_END = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b", re.I)
# Dim 语句行：行首（忽略缩进）就是 Dim。`ReDim` 天然被排掉（它以 Re 开头，
# 而且 ReDim 本来就不是"声明一个新名字"）。
_RE_V110_DIM = re.compile(r"^\s*Dim\b", re.I)


def _v110_after_continuation(raw, i):
    """第 i 行（0-based）是不是某个【续行的后半截】（上一物理行以 ` _` 结尾）。

    这种行不可能是语句开头：VBA 的行继续符写作 ` _`（空格 + 下划线），而
    `Dim a As Long, _` 的续行只能是 `       b As Long` 这种变量表，不会重复
    出现 Dim / Sub / End 关键字。所以过程头 / 过程尾 / Dim 三类判据都要先排掉
    它 —— 否则 `' 注释 _` 后面跟一个 `Dim x`（注释被续下去了）会被当成真的
    Dim 行：跳过去之后还会在注释中间插一行，把注释从中间拆断。

    只吃原始行（见上面关于两份行表的说明）。
    """
    return i > 0 and bool(vba_parser._RE_CONT_AT_END.search(
        (raw[i - 1] or "").rstrip("\r")))


def v110_proc_span(lines, raw, idx):
    """光标（idx，0-based）所在过程的 (头行, 尾行)（0-based，闭区间）。

    不在任何过程里（模块声明区 / 两个过程之间）/ 过程还没有收尾行时返回
    None —— 调用方据此把按键原样还给系统。

    两条特殊口径：
      * 光标【就在过程头那一行】时算"在过程里"（刚敲完 `Sub Foo()` 按
        Ctrl+↑ 就该能跳进它的声明区）。这与 parser.is_proc_header_line 的
        v79d 口径正好相反，但那边问的是"当前作用域"、这边问的是"归属"——
        两个问题，不能套同一句结论。
      * 光标【就在过程尾行】时先跳过它再往上找：否则一上来就撞见 End 行、
        永远判"不在过程里"，而"光标停在 End Sub 上按 Ctrl+↓"恰恰是最常用
        的姿势之一。
    """
    n = len(lines)
    if not (0 <= idx < n):
        return None
    s = (lines[idx] or "").strip()
    start = None
    if s and (vba_parser._RE_SUB.match(s) or vba_parser._RE_PROP.match(s)):
        # 光标在过程头上：它就是当前过程的头
        if _v110_after_continuation(raw, idx):
            return None                      # 续行后半截：模块形态有问题
        start = idx
    else:
        i = idx - 1
        while i >= 0:
            t = (lines[i] or "").strip()
            if t:
                if (vba_parser._RE_SUB.match(t) or vba_parser._RE_PROP.match(t)):
                    if _v110_after_continuation(raw, i):
                        return None          # 续行后半截：不猜，直接放行
                    start = i
                    break
                if vba_parser._RE_ANY_PROC_END.match(t):
                    return None              # 先撞见收尾行 => 在声明区里
            i -= 1
        if start is None:
            return None
    j = start + 1
    while j < n:
        t = (lines[j] or "").strip()
        if t and _RE_V110_PROC_END.match(t):
            if _v110_after_continuation(raw, j):
                return None                  # 续行后半截：同上，不猜
            return (start, j)
        j += 1
    return None          # 没收尾行（正在写的新过程）：不接管


def v110_stmt_end_line(raw, i, limit):
    """从第 i 行（0-based）出发跨过续行，返回该语句【最后一个物理行】的下标。

    续行一直拖到 limit（过程尾行）还不结束 => 返回 None：那种模块形态下
    "Dim 语句的下一行"是谁根本说不清，不接管。只吃原始行。
    """
    j = i
    while vba_parser._RE_CONT_AT_END.search((raw[j] or "").rstrip("\r")):
        j += 1
        if j >= limit:
            return None
    return j


def v110_dim_jump_plan(lines, raw, span):
    """Ctrl+↑ 的落点计划：`(目标行, 写入方式, 文本)`（均 0-based）。
    不该接管返回 None。

    ★v110c（用户改口径）：目标行 = 本过程最后一个 Dim 语句的【结束行】本身，
    **光标落在该行行尾**，不插新行、也不动任何文本（写入方式恒为 `"eol"`）。
    跨行的 Dim 语句要连它的续行一起跨过去 —— "行尾"取语句【结束行】的行尾
    （`Dim a, _` / `b As String` 那种），绝不能落在 `Dim a, _` 那一行。

    旧口径（v110~v110b/2，已废弃）：目标 = Dim 语句的【下一行】，不是空白就
    照 Dim 行缩进补一行。

    跨行判定与 `v110_end_jump_plan` 共用 `v110_stmt_end_line` / `_RE_V110_DIM`。
    """
    start, end = span
    d = None
    for i in range(end - 1, start, -1):          # 只扫过程体，头/尾两行不算
        t = lines[i]
        if not (t or "").strip():
            continue
        if not _RE_V110_DIM.match(t):
            continue
        if _v110_after_continuation(raw, i):
            continue          # 续行后半截不是 Dim 语句开头，继续往上找
        d = i
        break
    if d is None:
        return None
    de = v110_stmt_end_line(raw, d, end)
    if de is None:
        return None
    return (de, "eol", None)


def v110_end_jump_plan(lines, span):
    """Ctrl+↓ 的落点计划：`(目标行, 写入方式, 文本)`（0-based）。
    不该接管返回 None。

    目标行 = 过程尾行【上面那一行】。缩进取**过程体层级** = 尾行上方最后一条
    非空行的行首空白 —— 那正是用户接着要写语句的层级。

      * 尾行上方不是空白：在尾行处**插一行**，内容就是过程体缩进
        （`("insert", body)`；`End Sub` 顶格时 body == ""，就是一条真空行）；
      * 尾行上方已是空白、且是**真空行**：把它**换成带过程体缩进的一行**
        （`("swap", body)`；复用下面 `ctrl_nav_jump` 的"插入+删旧行"写法，
        因为 InsertLines 写空白行不会被 VBE trim，ReplaceLine 会 —— 见 v103b）；
      * 尾行上方已是空白、且**带着缩进**：什么都不写（`(None, None)`）——
        用户摆过层级，绝不加深。

    ★v110b/2（修真机 bug）。原先这一支是"插入真空行 + 注入一下 Tab 键让
    VBE 的 smart indent 去补缩进"。真机实测**注入的 Tab 不生效**（跳转正常、
    缩进没有）：`SetSelection` 是 COM 调用，返回时 VBE 未必已把光标画到新行，
    注入的 Tab 打在旧行上；补了轮询+延时也仍不出缩进。改成**直接把新行写成
    过程体缩进**——不依赖任何注入时序。这条写法有现成旁证："回车缩进"
    `new_line_below` 一直就是 `InsertLines(n, indent)` 直接写带缩进的（可能
    全空白）新行，真机好用（用户从没报过回车不缩进）。

    ⚠️ 判"空白"前必须 `rstrip("\r")`：模块文本是 \r\n 分隔的，`split("\n")`
    之后每行行尾都带 \r，而 `"\r".lstrip() == ""`（\r 算空白）—— 不 rstrip
    就会把真空行误判成"带着缩进"（本轮真踩到，见 MEMORY 第 42 条）。
    """
    start, end = span
    prv = end - 1
    if prv < start:
        return None
    # 过程体层级：从尾行上方往上找第一条非空行，抄它的行首空白
    body = ""
    for i in range(prv, start - 1, -1):
        t = (lines[i] or "").rstrip("\r")
        if t.strip():
            body = _leading_ws(t)
            break
    if (lines[prv] or "").strip():
        return (end, "insert", body)
    cur = (lines[prv] or "").rstrip("\r")
    if cur == "" and body != "":
        return (prv, "swap", body)
    return (prv, None, None)


def ctrl_nav_jump_plan(code, caret_line, direction):
    """Ctrl+↑（direction<0）/ Ctrl+↓（direction>0）这一跳的完整计划（纯函数）。

    code = 模块全文；caret_line = 光标行号（1-based）；direction = -1 / +1。
    返回 `(目标行[1-based], 写入方式, 文本)`。写入方式：
      * `"eol"`：只把光标落到目标行行尾、不动文本（★v110c 的 Ctrl+↑）；
      * `"insert"`：在目标行处插一行，文本就是新行内容（Ctrl+↓）；
      * `"swap"`：把目标那条真空空白行换成带缩进的（Ctrl+↓）；
      * `None`：不动文本，只落光标。
    不该接管（不在过程里 / 没有 Dim / 过程没收尾 / 模块形态有问题）返回 None。

    ⚠️ caret_line 允许是"模块最后一行 + 1"：VBE 允许光标停在一个尚不存在的
    虚拟空行上（见 `VbeBackend._write_line` 的说明）。那时按最后一行算 ——
    它必然落在过程尾行之下，判出来是"不在过程里"，原样放行。
    """
    try:
        caret_line = int(caret_line)
        direction = int(direction)
    except Exception:
        return None
    if not code or direction == 0 or caret_line < 1:
        return None
    raw = str(code).split("\n")
    masked = vba_parser._mask_strings_and_comments(code).split("\n")
    # 两份行表行数必须一致（抹白只把串/注释压成一个空格、不动换行）。不一致
    # 说明文本形态超出了这套判据的假设 —— 整体不接管，绝不拿错位的行号去写
    # 用户的代码。
    if len(masked) != len(raw):
        return None
    n = len(raw)
    if n <= 0:
        return None
    span = v110_proc_span(masked, raw, min(caret_line, n) - 1)
    if span is None:
        return None
    plan = (v110_dim_jump_plan(masked, raw, span) if direction < 0
            else v110_end_jump_plan(masked, span))
    if plan is None:
        return None
    tgt, mode, text = plan
    return (tgt + 1, mode, text)              # 转 1-based 交给 COM 用


# ---- v99：退格成对删除（空括号 / 空引号）----
#
# 用户口径：「按退格键删除字符时，如果删除的是左右括号，或双引号，如果括号中
# 或双引号中没内容，就能自动删除左右两端。」
#
# 即光标停在 `()` 或 `""` **中间**时按一次退格，两个字符一起消失
# （VSCode / IDEA / VBE 之外的现代编辑器都是这个手感）。有内容时
# （`(a|b)`）不接管 —— 那是一次普通删字，右括号得留着。
#
# 与 v94 退格删整行是同一副骨架：判据是纯函数、跑在键盘钩子线程里（只读
# 轮询快照，绝不碰 COM）；真正的删除在主线程 `delete_pair_here`；删不成就把
# 这一下退格【原样还给系统】。
#
# ⚠️ 本判据与 `_pair_insertion`（自动配对）刻意成对存在：自动配对插出来的
# `()`，退格就该一次收掉，否则用户敲完 `f(` 想反悔，得按两下退格。
_BS_PAIR_PAIRS = {"(": ")", '"': '"',
                  # ★v108：全角那一对也要能"一次退格收掉" —— 自动配对插出来的
                  # `（）` / `“”` 与半角 `()` / `""` 在这一点上等价。
                  "（": "）", "“": "”"}


def pair_delete_span(line_text, char_col=None):
    """光标停在【一对空括号 / 空引号的正中间】时，返回那一对的下标
    `(left, right)`（0-based）；不该接管时返回 None。★v99。

    这是本功能的**唯一一份判据** —— 钩子线程的 `pair_delete_wanted` 与主线程
    的 `delete_pair_here` 都调它（项目铁律：钩子判据与后端复核必须同一份实现，
    否则两边各写一遍，早晚有一边漏掉某个分支）。

    ⚠️★v99 **`char_col` 必须是【字符列】，不是 VBE 报的显示列**。
    我第一版让本函数自己去猜"VBE 这次报的是显示列还是字符列"（两种都试），
    结果在 `    说()话` 这种含全角的行上**假阳性**：光标真在 `(` 与 `)` 之间，
    显示列 7 被当成字符列 ⇒ left=5 ⇒ 命中了右邻的 `()` ⇒ 把不该删的两个字符
    删了。**判据一旦"两种都试"，就从精确判据退化成两个模糊判据的并集** ——
    漏判只是功能不生效，误判是改坏用户的代码，方向完全不可逆（MEMORY 第 19
    条、第 10 条）。⇒ 换算的责任交给调用方，各用自己最准的那份：
      * 主线程 `delete_pair_here` 用 `_bs_caret_char_col`（★v103 更正：这里原先
        写的是 `_caret_char_col`，是 v99 初稿的说法；**本功能的两个调用方**
        （钩子侧 `pair_delete_wanted` 与主线程 `delete_pair_here`）后来统一成
        了 `_bs_caret_char_col`，因为 `_caret_char_col` 的第三档是拿"行尾列"
        反推 tabw 的，而删括号这个功能的典型形态是**光标停在行中间**，反推必错
        （见第 27 条）。文档与代码不一致比没有文档更危险，故更正）；
      * 钩子线程 `pair_delete_wanted` 同样用 `_bs_caret_char_col` —— 两侧
        **同一份**（项目铁律：钩子判据与后端复核必须同口径）。

    判据本身极简，只问一件事：**光标左边那个字符和右边那个字符，是不是一对
    配对符**（`(` + `)`、`"` + `"`，或 ★v108 的全角 `（` + `）`、`“` + `”`）。
    这一条就把"中间没内容"钉死了 ——
    左边是 `(`、右边紧跟着 `)`，本来就意味着两格之间什么都没有。

    刻意**不做**的事：
      * 不问光标位置（行首 / 行尾都行）。`f(|)` 与 `f()| |` 的差别只在于
        用户接下来想干什么，不是"该不该删"；
      * 不问是不是在字符串/注释里。`"(|)"` 是字符串内容，按了也没坏处
        （用户自己敲出来的，不是我们插的），而"只删左边一个"反而更让人
        意外 —— 宁可跟编辑器主流行为一致；
      * **只认左开右闭**（`(` 在左、`)` 在右）。`f()|)` 这种"右括号挨着右
        括号"不接管 —— 退格删的是左边那个 `)`，右边的 `)` 与它无关。
    """
    t = line_text or ""
    if not t:
        return None
    try:
        cc = int(char_col or 0)
    except Exception:
        return None
    if cc <= 0:
        return None
    # 字符列 cc（1-based）⇒ 光标停在第 cc-1 个字符之前 ⇒ 左邻字符下标 cc-2。
    left = cc - 2
    right = left + 1
    if left < 0 or right >= len(t):
        return None
    close = _BS_PAIR_PAIRS.get(t[left])
    if close and t[right] == close:
        return (left, right)
    return None


def pair_delete_wanted(line_text, caret_ec=None, tabw=4):
    """钩子线程用的纯判据：这一下退格要不要由我们接管 —— 删掉一对空括号/
    空引号（★v99）。语义见 `pair_delete_span`。

    `caret_ec` 是**轮询快照里的 VBE 原始列（显示列）**，本函数负责换算成
    字符列再问判据 —— 钩子线程没有 COM，够不到 `_caret_char_col`，只能用
    `_bs_caret_char_col` 那套纯函数换算（与主线程 `delete_pair_here` 走的
    **同一份**，项目铁律第 18 条：两侧必须同口径）。

    ⚠️ 与 `blank_line_delete_wanted` 同处、同为纯函数：跑在键盘钩子线程里，
    只读轮询留下的快照，绝不碰 COM。
    """
    col, _sem, _tabw, _wide2 = _bs_caret_char_col(line_text, caret_ec)
    return pair_delete_span(line_text, col) is not None


def _bs_caret_char_col(line_text, ec):
    """v99 专用：VBE 的光标列 -> (字符列 1-based, sem, tabw, wide2)。

    ★为什么**不能**直接用 `_caret_char_col`：那条路的第三档
    `_fit_semantics` 是拿 **行尾列** 反推 (tabw, wide2) 的（`_detect_semantics_
    passive` 只在 `ec > len+1` 时才走到那里）。而本功能的典型形态是光标停在
    **行中间**的一对空括号上 —— `\tx = f()` 光标在 `()` 中间时 ec=11，
    被当行尾列去拟合，**反推出 tabw=3**（真值是 4）⇒ 字符列偏 1 ⇒
    `pair_delete_span` 读到 `f(` 上，一对都没匹配上，功能静默失效。

    ⇒ 这里只走**两条不需要拟合的路**（宁可判不出、交给 VBE 原生退格）：
      1) 行内既无 Tab 也无全角 ⇒ 两种列语义数值相同，直接当字符列用；
      2) 否则用**最常见的** (tabw=4, 全角占 2) 换算 —— 与
         `_eol_disp_col` / `blank_line_delete_wanted` 用的是同一组默认假设。

    ⚠️ 假设错了顶多让这一次判不出（或判错而删掉一对本该留着的括号）。前者
    无害，后者由调用方的兜底与下方的"读回复核"挡住：删完立刻比对写回结果，
    对不上就还按键。
    """
    t = line_text or ""
    try:
        ec = int(ec or 0)
    except Exception:
        return 1, "char", 4, False
    if ec <= 0:
        return 1, "char", 4, False
    if not ("\t" in t or any(_is_wide(c) for c in t)):
        return min(ec, len(t) + 1), "char", 4, False
    return _col_to_char_index(t, ec, 4, True) + 1, "disp", 4, True


def _unterminated_string(line_text):
    """这一行的【代码部分】是否"引号没闭合"（VBE 原生回车会替用户补上右引号）。

    注释里出现的引号不算（`x = 1  '他说"你好` 是好好的），字符串里的 `""`
    按 VBA 的转义写法跳过去。判定为 True 时调用方【不接管】回车 —— 让 VBE
    自己把右引号补上，那是它原生就有的、很有用的行为。
    """
    t = line_text or ""
    in_str = False
    i, n = 0, len(t)
    while i < n:
        c = t[i]
        if c == "'" and not in_str:
            break                       # 注释开始，后面的引号都不是代码
        if c == '"':
            if in_str and i + 1 < n and t[i + 1] == '"':
                i += 2                  # 字符串里的转义引号 ""
                continue
            in_str = not in_str
        i += 1
    return in_str


def _probe_semantics(cm, cp, line_no, line_text):
    """安全地探测列语义：移动光标探测后必定还原，并校验还原结果。

    某些编辑器（如 WPS 的 VBA 兼容编辑器）对 SetSelection 的处理与 Office
    不同：越界列可能不被钳制到行尾，或还原是异步的——探测列若没被正确还原，
    就会出现"光标自己跳到代码中间"。因此这里：
      1) 有缓存就直接返回，整个会话最多探测一次；
      2) try/finally 保证一定执行还原，且还原后读回校验（不符再还原一次）；
      3) 校验始终失败 -> 标记 probe_broken，此后永久禁用探测，退回被动判定。

    ⚠️★v95 **缓存里装的是"某一次探测的结论"，不是"编辑器的真理"**：若那次探测
    的行恰好【没有判别力】（既无 Tab 也无全角，两种语义数值相同），结论就是猜的
    —— 这种结论会被喂给后面**含中文的行**并算错（v94b / v95 两个 bug 同一个根因）。
    因此：
      * 调用方**别在拿不准的行上探测**（见 `_resolve_sem_info` 的行类别闸门）；
      * 凡"要把光标摆到某个位置"的地方，别信缓存，在本行上重新量
        （见 `apply_completion` 与 `_caret_to_line_end`）。
    """
    if _sem_cache["probe_broken"]:
        return _sem_cache["info"]
    cached = _sem_cache["info"]
    if cached is not None:
        return cached
    try:
        sl, sc, el, ec = cp.GetSelection()
    except Exception:
        return None
    info = None
    try:
        info = _vbe_measure(cm, cp, line_no, line_text)
    except Exception:
        info = None
    finally:
        # 无论探测成功与否，都必须把光标放回原处
        restored = False
        for _ in range(3):
            try:
                cp.SetSelection(sl, sc, el, ec)
                cur = tuple(cp.GetSelection())
                if cur == (sl, sc, el, ec):
                    restored = True
                    break
            except Exception:
                break
        if not restored:
            # 该编辑器无法可靠还原探测列：以后不再探测，避免光标乱跳
            _sem_cache["probe_broken"] = True
    if info is not None:
        _sem_cache["info"] = info
    return info


def _vbe_measure(cm, cp, line_no, line_text):
    """探测 VBE 对当前行的"列"语义与 Tab 宽度（不恢复光标，调用方负责）。

    返回 (sem, tabw, wide2)：
      sem == 'char' -> VBE 列 == 字符位置（Tab 计 1），直接按字符数用列；
      sem == 'disp' -> VBE 列 == 显示列（Tab 展开），需按 (tabw, wide2) 换算。
    探测法：把光标设到第 len+100 列（远超行尾），读回实际停靠列 E，再拟合。

    注意：本函数会【移动光标】。除 apply_completion（随后本来就要设置光标）
    之外，请一律改用 _probe_semantics，它会自动还原并校验。

    ★v95：拿不到就返回 **None**（旧实现返回 ("char", 4, False)）。原因是"量不到"
    与"量出来是用字符列"是两回事 —— 前者是**没有证据**，而旧返回值会被调用方
    当成证据**写进缓存**，于是"哪儿都没量出来"变成了"全会话都按字符列算"
    —— 正是 v94b 那个坑的另一种写法。None 让调用方各走各的兜底，不污染缓存。
    """
    n = len(line_text)
    try:
        cp.SetSelection(line_no, n + 100, line_no, n + 100)
        _sl, _sc, _el, ec = cp.GetSelection()
    except Exception:
        return None
    return _fit_semantics(line_text, ec)


def _find_completion_end(text, completion, near0, fallback0):
    """在(可能被 VBE 规范化过的)行文本里定位补全词 completion 的结尾字符下标。

    要求找到的 occurrence 前后都不是标识符字符（整词匹配）；
    若行内多处出现同名，取起点最接近 near0 的那个。
    找不到则回退 fallback0（裁剪到行尾）。
    """
    L = len(completion)
    if not L:
        # ★v103：`text.find("", i)` 恒等于 i、而 `i += 0` 不动 ⇒ **死循环**
        # （整个补全线程卡死）。当前所有调用方传的都是非空候选，所以这是个
        # 潜伏缺陷；但只要将来有人拿"候选为空"调一次就会挂住，先挡住。
        return min(fallback0, len(text))
    best = None
    i = 0
    while True:
        i = text.find(completion, i)
        if i < 0:
            break
        ok_l = (i == 0) or (not _is_ident_tail(text[i - 1]))
        ok_r = (i + L >= len(text)) or (not _is_ident_tail(text[i + L]))
        if ok_l and ok_r:
            if best is None or abs(i - near0) < abs(best - near0):
                best = i
        i += L
    if best is not None:
        return best + L
    return min(fallback0, len(text))


def _in_call_stmt(line_text):
    """这一行是不是 VBA 的 `Call` 语句（`Call 过程名`）。★v102。纯函数。

    ★★为什么必须加这条（用户报"Call judge 按 Tab 不补括号、手动敲括号也没反应"）：

    **VBE 会把 `Call foo()` 里的【空括号】自动删掉** —— 这是 VBA 的规范：
    `Call` 语句在无参数时不写括号，`Call foo` 才是标准写法。
    真机实测（独立 Excel 实例 + 真 CodeModule，v102g 内容矩阵）：

        写 '    Call judge()'     -> 立即 '    Call judge'      ← 括号被 VBE 删掉
        写 '    Call judge(1, 2)' -> '    Call judge(1, 2)'      ← 有参数则保留
        写 '    Call foo()'       -> 立即 '    Call foo'        ← 同样被删
        写 '    judge()'          -> '    judge()'               ← 不在 Call 语句里，保留
        写 '    x()'              -> '    x()'                   ← 保留

    ⇒ 这个位置上做括号动作**全是白做**，而且是最坏的那种"白做"——
      用户看到的是"按了没反应"：
        * 补全补出 `Call judge()` → VBE 立刻删 → 用户以为"Tab 没补括号"；
        * 自动配对写入 `Call judge()` → VBE 立刻删 → 用户以为"敲括号没反应"。
      两个现象同一个根因，与我们的判据、让位、时序**全都无关**。

    ⇒ 修法就是【什么都不做】：补全不补括号、自动配对不接管，让 VBE 全权处理
      （它自己插入 `(`，用户接着输入参数时会得到 `Call judge(1, 2)` 并保留）。

    ⚠️ 判据只看【行首是不是 Call 关键字】：`Call obj.Method` 同样是 Call 语句、
      同样会被删括号，所以不能用"标识符紧跟在 Call 后面"这种更窄的判据。
    """
    try:
        return bool(_CALL_STMT_RE.match(str(line_text or "")))
    except Exception:
        return False


# `^\s*Call` 且后面不是标识符字符（挡住 `Callback` / `CallCount` 这类名字）。
_CALL_STMT_RE = re.compile(r"^[ \t]*call(?![A-Za-z0-9_])", re.I)


def _space_insert_pos(new_line, word_start_col, completion):
    """要在 `new_line` 的哪个【字符下标】注入一个空格；不需要补就返回 None。

    ★v103b 纯函数（不碰 COM），供 `apply_completion` 判"关键字后要不要补空格"，
    也让这条判据能被离线用例直接钉住。

    规则：补全词的**词尾**右边如果不是空白，就在那里补一个空格。
      * 词尾已在行末            -> 要补（返回词尾下标）；
      * 词尾右边是空格 / Tab    -> 不补（用户自己打了，或上一轮已经补过）；
      * 词尾右边是别的字符      -> 要补（把关键字与后面的东西隔开）。
    注意返回的是**词尾下标**（0-based，插入点），不是"跳过空白之后"的位置 ——
    同 v98 括号那条铁律：为判断而算的位置不能拿去当插入点。
    """
    try:
        k = int(word_start_col) - 1 + len(str(completion or ""))
    except Exception:
        return None
    if k < 0:
        return None
    line = str(new_line or "")
    if k >= len(line):
        return k
    if line[k] in " \t":
        return None
    return k


def _is_call_site(line_text, word_start_col):
    """这一行的这个词位置，是不是一个**可以直接调用过程**的位置。纯函数。

    ★v98：决定"补全时要不要自动补 `()`"的那道闸门 —— 补错了就是坏代码，
    而最容易被补坏的就是**类型名位置**（`Dim n As <这里>` / `Dim o As <这里>` /
    `Property Set o = <这里>` / `Function f() As <这里>`）：VBA 里数据类型后面
    不能带括号，补成 `Dim n As Long()` 编译就过不去。
    所以判据的核心一句话：**这个词的前一个字，是【可以接收一个调用表达式】的
    那个位置**，还是【紧跟在 `As` 后面当类型用】的那个位置。

    line_text     : 该行文本
    word_start_col: 补全词起始的 1-based **字符**列（与 Python 下标一致 +1）
    """
    try:
        i = int(word_start_col) - 1
        if i < 0:
            return True                       # 行首裸词（Call Foo / Foo 1）→ 可以
        # 词左边紧邻的字符（跳过空白）。没有左边 ⇒ 行首 ⇒ 可以。
        j = i - 1
        while j >= 0 and line_text[j] in " \t":
            j -= 1
        if j < 0:
            return True
        # ---- ⚠️ 这两条判据必须一起用切片看，**不能**拿"跳过空白后的那个字符"
        #    去比字面：`As Long` 里 `Long` 的前一格是【空格】，不是 `s`。
        #    （我第一版写成 `prev in "sS"` ⇒ As 分支永远不成立，实测 3 条红；
        #     New 那条因为用了切片才侥幸对。）所以：向左扫掉空白后，
        #    取最后 2~3 个字符做小写比较。
        tail = line_text[:j + 1].lower().rstrip(" \t")
        # `As` 类型位置：`Dim n As Long` / `Function f() As Long` / `Set p As Cls`
        # 补成 `Long()` 编译不过 —— 数据类型后面不能带括号。
        if tail.endswith("as"):
            return False
        # `New` 类型位置：`Set fso = New FileSystemObject` —— New 要的是
        # 类型表达式，不是调用，补成 `FileSystemObject()` 是语法错误。
        if tail.endswith("new"):
            return False
        # 点号左边：这是【成员访问】。`obj.method` 可以补括号，
        # `obj.prop` 补了就错了 —— 但我们分不清成员是方法还是属性，
        # 而用户明确要求"选中函数或过程提示词后补括号" ⇒ 这里放行，
        # 让 VBE 自己裁决（真写成属性时 VBE 会当场报"不能带参数"）。
        # ⚠️ 这条也覆盖 `Range("A1").method`。
        # 其余情况（= + 运算符 / , / ( / 空白 / 行首 …）都可以是调用点。
        return True
    except Exception:
        # 判据出错就【不补】—— 宁可少补，也不要把用户的代码写坏。
        return False


def _skip_into_parens(actual, end0, orig_line, orig_caret_col):
    """补全词之后若紧跟一个「新出现」的左括号，把光标下标移到它右侧。

    VBE 会自动为过程声明补括号：把 `Function gCalc` 写回后，行里会冒出一对
    `()`。光标停在词尾就等于停在左括号【左侧】，用户还得手动右移一格才能写
    参数。这里检测"原行光标后本来不是 `(`、写回后却出现 `(`"，就把光标送进
    括号里；像 `x = gCalc(1)` 这种括号本来就有的调用点则保持不变。

    actual  : 写回后读回的实际行文本
    end0    : 补全词结尾的字符下标（0-based，即光标应停的位置）
    orig_line / orig_caret_col : 写回前的行文本与光标列（1-based 字符列）
    """
    try:
        i = int(orig_caret_col) - 1
        orig_after = orig_line[i:i + 1] if 0 <= i < len(orig_line) else ""
        if orig_after == "(":
            return end0
        j = end0
        n = len(actual)
        while j < n and actual[j] == " ":
            j += 1
        if j < n and actual[j] == "(":
            return j + 1
    except Exception:
        pass
    return end0


def _caret_inside_parens(line_text, caret0):
    """★v107c：光标此刻是否正停在【一对括号的中间】（纯函数）。

    判据只有一条、且只此一份实现：光标左边紧邻的那个字符是 `(`。
    对"我们刚补出来的那对空括号"来说，它恰好等价于"光标落在双括号中间" ——
    要不要唤出参数信息只看这一个事实，不掺任何猜测。

    为什么不能拿 `added_parens` 当同一件事：`_skip_into_parens` 有可能没把
    光标送进去（VBE 把整行改写成了别的样子）。那时再按一下 Ctrl+Shift+I 是
    没头没脑的一下，有了这道位置判据就宁可什么都不做。

    line_text : 写回后的实际行文本
    caret0    : 光标的 0-based 字符下标（= 光标左边内容的长度）
    """
    try:
        i = int(caret0)
        if i < 1 or i > len(line_text):
            return False
        return line_text[i - 1] == "("
    except Exception:
        return False


# ---------------------------------------------------------------------------
# v104：宿主探测 —— 不再写死 Excel
# ---------------------------------------------------------------------------
# VBE（VBA 集成开发环境）是微软交给【每个 Office 宿主】共用的组件：
# Excel / Word / PowerPoint / Access / Visio 用的都是同一个 VBE、同一套
# VBIDE 对象模型（VBProjects -> VBComponents -> CodeModule / ActiveCodePane）。
# 所以本项目对宿主的依赖，实质只有"外面那层 Application 的 ProgID"一处；
# 只要按当前宿主换一个 ProgID，其余全部复用。WPS 的 VBA 兼容环境同理。
#
# 候选宿主 ProgID 表。顺序 = 拿不到"前台线索"时的兜底优先级
# （Excel 放第一，保持与旧版本一致的行为）。
_HOST_PROGIDS = (
    # ---- Microsoft Office ----
    "Excel.Application",
    "Word.Application",
    "PowerPoint.Application",
    "Access.Application",
    "Visio.Application",
    # ---- WPS Office（金山）—— ProgID 视版本有新旧两套写法，都试 ----
    "KWPS.Application",   # WPS 文字（对应 Word）
    "KET.Application",    # WPS 表格（对应 Excel）
    "KWPP.Application",   # WPS 演示（对应 PowerPoint）
    "WPS.Application",    # 旧版 WPS 文字
    "ET.Application",     # 旧版 WPS 表格
    "WPP.Application",    # 旧版 WPS 演示
)

# 进程可执行文件名 -> ProgID。用于"前台窗口属于哪个宿主进程"这条优先线索，
# 让工具跟着【用户当前正在用的】那个宿主走（而不是候选表顺序）。
_HOST_EXE_PROGID = {
    "excel.exe": "Excel.Application",
    "winword.exe": "Word.Application",
    "powerpnt.exe": "PowerPoint.Application",
    "msaccess.exe": "Access.Application",
    "visio.exe": "Visio.Application",
    "wps.exe": "KWPS.Application",
    "et.exe": "KET.Application",
    "wpp.exe": "KWPP.Application",
}

# 上次成功连上的宿主 ProgID —— 下次优先试它，省掉一轮无谓的 GetActiveObject。
_host_progid_hint = {"progid": None}


def _process_image_name(pid):
    """PID -> 进程可执行文件名（小写，含 .exe）；拿不到返回 ""。纯 Win32。"""
    import ctypes
    import ctypes.wintypes
    if not pid:
        return ""
    try:
        k32 = ctypes.windll.kernel32
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        h = k32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
        if not h:
            return ""
        try:
            buf = ctypes.create_unicode_buffer(1024)
            size = ctypes.wintypes.DWORD(len(buf))
            if k32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
                path = buf.value or ""
                return path.replace("/", "\\").rsplit("\\", 1)[-1].lower()
        finally:
            try:
                k32.CloseHandle(h)
            except Exception:
                pass
    except Exception:
        pass
    return ""


def _foreground_hwnd():
    """前台窗口句柄（取不到返回 0）。单独抽出来便于测试打桩。"""
    try:
        import ctypes
        return int(ctypes.windll.user32.GetForegroundWindow() or 0)
    except Exception:
        return 0


def _preferred_progid():
    """用【前台窗口】所属进程猜用户正在用的宿主 -> ProgID；猜不到返回 None。

    打字时光标在 VBE 里，前台窗口就是宿主进程（VBE 是进程内的）的窗口，所以
    这条线索很准。纯 Win32，不碰 COM，可在钩子线程安全调用。
    """
    import ctypes
    import ctypes.wintypes
    try:
        u32 = ctypes.windll.user32
        hwnd = _foreground_hwnd()
        if not hwnd:
            return None
        try:
            root = u32.GetAncestor(hwnd, 2)   # GA_ROOT
            if root:
                hwnd = root
        except Exception:
            pass
        pid = ctypes.wintypes.DWORD()
        u32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        exe = _process_image_name(pid.value)
        if exe:
            return _HOST_EXE_PROGID.get(exe)
    except Exception:
        pass
    return None


def _get_active_object(progid):
    """取一个正在运行的 COM 服务器对象（拿不到则抛异常）。

    单独抽出来是为了让宿主探测能离线测试（用例打桩这一层即可，无需真 COM）。
    """
    import win32com.client
    return win32com.client.GetActiveObject(progid)


def _get_vbe():
    """连到【当前宿主】的 VBE —— 不再写死 Excel（v104）。

    顺序：① 前台窗口所属进程（用户正在用的宿主）→ ② 上次成功的宿主 →
    ③ 候选表兜底。取第一个能拿到 `.VBE` 的。

    全都不行时，把最后一个异常抛出去 —— 与旧行为一致（宿主没开 / 没开信任时
    GetActiveObject 会抛 com_error），交给 `_get_vbe_cached` 走 `_com_fail()`。
    """
    order = []
    pref = None
    try:
        pref = _preferred_progid()
    except Exception:
        pref = None
    for p in (pref, _host_progid_hint.get("progid")):
        if p and p not in order:
            order.append(p)
    for p in _HOST_PROGIDS:
        if p not in order:
            order.append(p)

    last = None
    for progid in order:
        try:
            app = _get_active_object(progid)
        except Exception as e:
            last = e
            continue
        try:
            vbe = app.VBE
        except Exception as e:
            last = e
            continue
        if vbe is not None:
            if _host_progid_hint.get("progid") != progid:
                _log("host: 已连到宿主 %s 的 VBE" % progid)
            _host_progid_hint["progid"] = progid
            return vbe
    if last is not None:
        raise last
    return None


def _caret_position(vbe):
    """只读地取当前光标位置，返回 (module_name, (line_no, char_col)) 或 (None, None)。

    只调用 GetSelection、绝不 SetSelection，因此不会移动光标
    （WPS 等编辑器对 SetSelection 还原不可靠，移动光标会导致光标乱跳）。

    列语义优先用被动判定，其次用会话缓存；两者都拿不到时，仅当该行是
    "纯 ASCII 且无 Tab"（此时 VBE 列必然等于字符列）才直接换算，否则放弃——
    宁可不排除，也不要因为列换算不准而抹掉错误的词。
    """
    try:
        cp = vbe.ActiveCodePane
        if cp is None:
            return None, None
        cm = cp.CodeModule
        _sl, _sc, el, ec = cp.GetSelection()
        if el <= 0:
            return None, None
        line_text = cm.Lines(el, 1)
        info = _detect_semantics_passive(line_text, ec)
        if info is None:
            info = _sem_cache.get("info")
        if info is not None:
            sem, tabw, wide2 = info
            if sem == "disp":
                col = _col_to_char_index(line_text, ec, tabw, wide2) + 1
            else:
                col = min(ec, len(line_text) + 1)
        elif "\t" not in line_text and not any(_is_wide(c) for c in line_text):
            col = min(ec, len(line_text) + 1)
        else:
            return None, None

        # 模块名必须与收集时的 comp.Name 同源，否则比对失败会导致 caret 被丢弃、
        # 光标处的词被当成隐式变量收录（自我提示"打什么提示什么"）。
        # CodeModule.Name 在部分宿主里与 VBComponent.Name 不一致，故优先取后者。
        mod_name = ""
        try:
            mod_name = str(cp.CodeModule.Parent.Name)
        except Exception:
            pass
        if not mod_name:
            try:
                mod_name = str(cm.Name)
            except Exception:
                mod_name = ""
        return mod_name, (el, col)
    except Exception:
        return None, None


def _active_vb_project(vbe):
    """返回当前活动工程（VBProject 对象）；取不到返回 None。

    VBE.ActiveVBProject 直接对应"当前代码窗所属工程"；个别宿主/WPS 上该属性
    可能不可用，则退到 CodeModule.Parent(VBComponent).Collection.Parent 这条
    链路。两条都拿不到时返回 None（调用方退回"全部工程"的旧行为）。
    """
    try:
        p = vbe.ActiveVBProject
        if p is not None:
            return p
    except Exception:
        pass
    try:
        cp = vbe.ActiveCodePane
        if cp is None:
            return None
        return cp.CodeModule.Parent.Collection.Parent
    except Exception:
        return None


def _project_name(proj):
    try:
        return str(proj.Name)
    except Exception:
        return ""


def _project_key(proj):
    """工程身份 `(Name, FileName)`（小写）；取不到返回 None。★v124

    **为什么不能只比 `Name`**：Excel 里每个工作簿的工程名默认都叫
    `VBAProject` —— 同时开两个 .xlsm 时 `Name` 一模一样，缓存（组件列表 /
    CodeModule / 模块解析）都只用模块名做键，A 的 `Module1` 就会命中 B 的
    缓存，补全出**另一个工作簿**的变量。`FileName` 才是工程的磁盘路径。

    ★真机实测（Excel 2016 / VBE7.1）：已保存的工作簿返回完整路径；
    **未保存的工作簿读 `FileName` 会抛 com_error（"路径未找到"）而不是返回
    空串** —— 所以两个属性各自 try，谁抛了就当空。

    两个都没取到（工程被锁 / COM 抖了）→ None，调用方按"不信任缓存"处理。
    """
    if proj is None:
        return None
    try:
        name = str(proj.Name or "")
    except Exception:
        name = ""
    try:
        fname = str(proj.FileName or "")
    except Exception:
        fname = ""
    name = name.strip().lower()
    fname = fname.strip().lower()
    if not name and not fname:
        return None
    return (name, fname)


def _component_edge(proj, count):
    """组件列表的**首尾组件名** —— 工程身份的补充指纹。★v124

    只读 `Item(1)` / `Item(count)` 两次属性（比枚举全部便宜一个数量级）。

    ⚠️ 别指望它包打天下：★真机实测两个都用默认模板新建的工作簿，组件名
    集合**完全相同**（`ThisWorkbook` / `Sheet1` / `模块1`）——这时它分不开，
    真正起作用的是 `FileName`。它只是多一道几乎免费的信号，用来兜住
    "`FileName` 读不到（未保存）+ 名字恰好不同"的那一小撮情况；两个都
    未保存且结构相同时仍会共享缓存，由 `_COMP_LIST_TTL` / `_MOD_CACHE_TTL`
    到期自愈。

    读不到就给空串（不参与区分），绝不抛。
    """
    if not count or count < 1:
        return ("", "")
    out = []
    for i in (1, count):
        try:
            out.append(str(proj.VBComponents.Item(i).Name or "").lower())
        except Exception:
            out.append("")
    return tuple(out)


def _components_of(proj):
    try:
        out = []
        for comp in proj.VBComponents:
            out.append(comp)
        return out
    except Exception:
        return []


def _code_panes_by_module(vbe):
    """`{模块名(小写): CodePane}` —— 只含**已经打开过代码窗**的模块。

    ★v121 真机实测的两件事（这两条决定了跳转怎么落地）：

      1. **`VBComponent` 没有 `CodePane` 属性** —— `comp.CodePane` 直接抛
         `AttributeError: <unknown>.CodePane`。代码窗只能从 `vbe.CodePanes` 里找
         （每个 pane 自带 `CodeModule` / `SetSelection` / `Show`）。
      2. **`CodePanes` 只含打开过的代码窗**：本机 24 个组件、只有 6 个 pane。
         没有 pane 的模块，COM 里没有任何"打开它"的接口（`Windows` 里也没有对应
         条目）—— 所以"跳到没打开过的模块"做不到，只能如实告诉用户去双击一下。

    拿不到（宿主不支持 / 还没进 VBE）时返回空 dict，调用方按"一个都跳不了"处理。
    """
    out = {}
    try:
        n = int(vbe.CodePanes.Count)
    except Exception:
        return out
    for i in range(1, n + 1):
        try:
            pane = vbe.CodePanes.Item(i)
            out[str(pane.CodeModule.Name or "").lower()] = pane
        except Exception:
            continue
    return out


# ★v118：组件列表缓存。`for comp in proj.VBComponents` 这一圈是纯 COM 枚举，
# 而【真正的贵】不在于枚举本身，而在于每次枚举出来的都是**新的 pywin32 包装
# 对象** —— 之后每读一次 `comp.Name` / `comp.Type` / `comp.CodeModule` 都要
# 重新 Bind 一遍类型信息（实测 24 个模块约 40ms/轮）。
#
# 复用上一次那批对象，后续属性访问就几乎免费。失效条件三道：
#   * 组件【数量】变了（增删模块）—— 一次 Count 就够，很便宜；
#   * 工程【身份】变了（`_project_key` + 首尾组件名，★v124）—— 同时开两个
#     工作簿时 `Count` 和模块名都可能撞车，没有这一道就会把 A 工程的组件
#     列表当成 B 的（`"proj"` 字段以前写了却从不参与比较，等于没有）。
#     ★真机实测：两个 .xlsm 的 `Name`（都是 `VBAProject`）**和组件名集合**
#     （ThisWorkbook / Sheet1 / 模块1）可以完全一样，唯一不同的是 `FileName`
#     —— 所以身份必须包含它。
#   * _COMP_LIST_TTL 到点 —— 兜住"改了模块名但数量没变"这种。
_COMP_LIST_TTL = 15.0
_comp_list_cache = {"ts": 0.0, "count": -1, "key": None, "edge": None,
                    "comps": None}


def _collect_components(vbe, caret_mod=None):
    """返回本次应参与补全的组件列表（默认只含【活动工程】的组件）。

    两道保险，宁可退回"全部工程"也不能让补全整个失效：
      1) 活动工程读不到组件（工程被密码保护/锁定）；
      2) 已知光标所在模块名，却不在活动工程的组件里——说明取到的"活动工程"
         不是真正正在编辑的那个，退回全工程（旧行为）。
    """
    if not ENABLE_CROSS_PROJECT:
        proj = _active_vb_project(vbe)
        if proj is not None:
            comps = None
            # 工程身份只算一次（两个属性读），命中判定与写入都要用。
            _key = _project_key(proj)
            # ★v118：先看缓存 —— 同一批组件对象复用，属性访问免 Bind。
            if MOD_CACHE_ENABLED and _comp_list_cache["comps"] is not None:
                try:
                    _cnt = int(proj.VBComponents.Count)
                except Exception:
                    _cnt = -1
                if (_cnt == _comp_list_cache["count"]
                        and _key is not None
                        and _key == _comp_list_cache["key"]
                        and (time.time() - _comp_list_cache["ts"])
                        < _COMP_LIST_TTL
                        # ★v124：再核一道首尾组件名 —— 极便宜的补充信号
                        # （真正把两个 .xlsm 分开的是上面的 `_key`；
                        #  这一道只兜"未保存 + 名字恰好不同"的情况）。
                        and _component_edge(proj, _cnt)
                        == _comp_list_cache["edge"]):
                    comps = _comp_list_cache["comps"]
                    _log("collect: 复用组件列表（%d 个）" % len(comps))
            if comps is None:
                comps = _components_of(proj)
                if comps:
                    _comp_list_cache["ts"] = time.time()
                    _comp_list_cache["count"] = len(comps)
                    _comp_list_cache["key"] = _key
                    _comp_list_cache["edge"] = _component_edge(proj, len(comps))
                    _comp_list_cache["comps"] = comps
            if comps:
                if caret_mod:
                    names = []
                    for c in comps:
                        try:
                            names.append(str(c.Name).lower())
                        except Exception:
                            names.append("")
                    if str(caret_mod).lower() not in names:
                        _log("collect: 活动工程 %r 不含模块 %r，退回全工程"
                             % (_project_name(proj), caret_mod))
                        proj = None
                        comps = []
                        # 缓存与现场对不上（换工程 / 改了模块名）—— 作废
                        _comp_list_cache["comps"] = None
                        _comp_list_cache["key"] = None
                        _comp_list_cache["edge"] = None
            if comps:
                _log("collect: 仅收集活动工程 %r（%d 个组件）"
                     % (_project_name(proj), len(comps)))
                return comps
    comps = []
    try:
        for p in vbe.VBProjects:
            comps.extend(_components_of(p))
    except Exception:
        return []
    _log("collect: 收集全部工程（%d 个组件）" % len(comps))
    return comps


def implicit_gate_applies(code):
    """这个模块的"隐式变量"要不要被 Option Explicit 闸掉（v75）。

    True  = 该模块写了 `Option Explicit` ⇒ 只收真声明过的名字，
            "用过但没声明"的名字不进候选池（那种写法在该模块里是编译错误）。
    False = 照旧收（没写 Option Explicit，或者开关被关掉了）。

    纯函数（只看文本 + 两个开关，不碰 COM），便于单测 —— 这里最容易出错的
    是开关极性，而不是文本判定本身。
    """
    if not ENABLE_IMPLICIT_IDENTIFIERS or not IMPLICIT_HONOR_OPTION_EXPLICIT:
        return False
    try:
        return vba_parser.has_option_explicit(code)
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Excel / VBE 的 COM 代理生命周期  ——  「关文件卡 3~10 秒 + 关完又启动 Office」
# 的病根就在这里
#
# 症状的关键线索：**改过 VBA 代码才卡，只改工作表内容不卡**。
#
# 机理：只要我们的进程持有 Excel.Application（哪怕只是间接经由 VBE 对象），
# Excel 就退不干净 —— 主窗口关了但 EXCEL.EXE 仍驻留后台，之后再打一次 COM
# 就会把它「唤醒」重新弹窗，看起来就是"关完又启动 Office"。
# 而「改过 VBA 才慢」正是因为 VBAProject 变脏后，Excel 退出时要把 VBA 工程
# 写回文件，这一步要与 VBA 环境交互，被外部 COM 引用钉住就会拖上好几秒；
# 没改 VBA 就不必写回，直接退出所以完全无感。
#
# 第一轮之所以没问题，是因为当时每次都 `_get_vbe()` 新建、用完即弃 —— 局部
# 变量在函数返回时引用计数归零，COM 引用立刻释放。后来改成模块级长缓存，
# 引用从此常驻，病就来了。
#
# 对策（三道，缺一不可）：
#   1) 代理设 TTL，到期主动丢弃 —— 回应用完即弃的语义，但不牺牲打字性能；
#   2) 宿主窗口（Excel 主窗口 / VBE 窗口）消失后，绝不再碰 COM；
#   3) 连续失败按档退避，不在 Excel 关闭窗口期猛打 COM。
# ---------------------------------------------------------------------------

# 缓存的 VBE/Excel COM 代理最长存活时间（秒）。0 或负数 = 彻底不缓存。
# 想回到"每次新建、用完即弃"（最保守）就设成 0。
VBE_PROXY_TTL = _settings.float_value("PROXY_TTL", 1.0)

# 宿主窗口探测门。类名对不上的宿主（如 WPS）可能误判，可以关掉：
#   VBECOMPLETE_NO_WINDOW_GATE=1  或  settings.ini 里 `NO_WINDOW_GATE = 1`
_WINDOW_GATE_DISABLED = _settings.flag("NO_WINDOW_GATE", False)

_vbe_cache = {"obj": None, "at": 0.0}
_wnd_state = {"seen": False, "miss": 0}

# 宿主主窗口 / VBE 主窗口的类名（用于"宿主还在不在"的门控与按进程过滤提示窗）。
# ★v104：VBE 主窗类名 wndclass_desked_gsk 是各 Office 宿主【共用】的，放第一位；
# 其后是各宿主自己的主窗类名。类名对不上的宿主（如 WPS）不会因此失效 ——
# _host_window_alive() 有"从没见过就不拦"的兜底。
_HOST_WINDOW_CLASSES = (
    "wndclass_desked_gsk",   # VBE 主窗（Office 各宿主共用）
    "XLMAIN",                # Excel
    "OpusApp",               # Word
    "PP12FrameClass",        # PowerPoint（2007+）
    "screenClass",           # PowerPoint（更早版本）
    "OMain",                 # Access
    "VisioGuy",              # Visio
)


def _probe_host_window():
    """纯 Win32 探测宿主窗口是否还在（完全不碰 COM）。"""
    import ctypes
    u32 = ctypes.windll.user32
    fn = u32.FindWindowW
    fn.restype = ctypes.c_void_p
    fn.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p]
    for cls in _HOST_WINDOW_CLASSES:
        if fn(cls, None):
            return True
    return False


def _host_window_alive():
    """宿主窗口是否还在（可 monkeypatch，测试会替换 _probe_host_window）。

    刻意带"从未见过就别拦"的兜底：若这台机器上从来没探测到过 Excel/VBE
    窗口，很可能是类名不匹配的宿主（WPS 等），此时盲目拦截会让工具整个
    失效。只有"曾经见过、现在连续两次都看不到"才认定宿主真的退出了。

    ★v106：本函数额外承担"宿主回来了就立刻解除 COM 退避"的职责 —— 所以
    **调用方在退避期内也必须调它**（见 `_get_vbe_cached`）。这是"开机自启
    后要等半分钟才有提示"的修法，详见下面 com_reset 那两行注释。
    """
    if _WINDOW_GATE_DISABLED:
        return True
    try:
        alive = _probe_host_window()
    except Exception:
        return True
    if alive:
        # 两种"宿主回来了"的证据，命中就【立刻】解除退避：
        #   * miss >= 2 —— 曾经见过、又连续两次看不到，现在又出现了
        #     （关掉 Excel 再打开）；
        #   * 从没见过（seen=False）—— ★v106 新增，开机自启的典型路径：
        #     工具随开机启动时 Excel 还没开，探测一直看不到 ⇒ 上面那条永远
        #     等不到（miss 根本没机会累加，见下面 `if not seen: return True`），
        #     而 COM 那一侧早已失败 5 次、退避升到 30 秒。用户这时才打开
        #     Excel —— 不在这里解除退避就得白等满 30 秒。
        if _wnd_state["miss"] >= 2 or not _wnd_state["seen"]:
            com_reset()
        _wnd_state["seen"] = True
        _wnd_state["miss"] = 0
        return True
    if not _wnd_state["seen"]:
        return True                   # 从未见过：不拦，避免误伤其它宿主

    _wnd_state["miss"] += 1
    return _wnd_state["miss"] < 2     # 连续 2 次才认定真的没了


# ---------------------------------------------------------------------------
# v63：VBE 自己弹的提示窗（参数信息 / 列出成员）此刻是否可见
# ---------------------------------------------------------------------------
# 让位机制的第二路判据。v55 那套走的是【光标语法位置】（`标识符.` 之后、`As ` /
# `New ` 之后），只能覆盖"自动列出成员 / 类型列表"；VBE 在别处弹的提示它一概不
# 认 —— 最典型的就是【参数信息】：`MsgBox "已完成！",` 敲下逗号后 VBE 弹出参数
# 签名，我们的弹窗同时也冒出来，两个窗叠在同一处（用户 v63 报的 bug）。
#
# 真机实测（本机 Office VBE，2026-09-14）：
#   * VBE 的提示窗是【预建复用】的真窗口 —— 什么都没按的时候它们就在顶层窗口
#     列表里，只是 IsWindowVisible=False；发 Ctrl+Shift+I（参数信息）后立刻
#     变 True、Esc 后立刻回 False。
#   * 类名是 VB IDE 专有的两个：
#       NameListWndClass —— 列出成员 / 列出常量
#       PopupTipWndClass —— 参数信息（Quick Info）
#   * 开销：全量枚举（EnumWindows + GetClassName）1.75ms/次、按进程过滤
#     0.81ms/次；而句柄缓存之后只查 IsWindowVisible，0.0005ms/次。所以这里缓存
#     句柄，只在"缓存为空"或"句柄失效（Excel/VBE 重启过）"时重新枚举，且限流。
#
# ⚠️ v55 曾以"给 VBE 发 Ctrl+J 后全系统零新窗口"推断【VBE 的列表没有独立窗口】。
# 那个结论只对"【新建】窗口"成立 —— 实际上窗口早就建好了，只是被显示/隐藏，
# 所以枚举"新窗口"永远看不到它。这次改成直接查已知类名的可见性，才拿到真相。
#
# 想追加别的类名（不同 VBE 版本类名可能不一样）：
#   set VBECOMPLETE_VBE_POPUP_CLASS=ClassA,ClassB
VBE_POPUP_CLASSES = ("NameListWndClass", "PopupTipWndClass")
try:
    _extra_popup_cls = _settings.items("VBE_POPUP_CLASS")
    if _extra_popup_cls:
        VBE_POPUP_CLASSES = tuple(VBE_POPUP_CLASSES) + _extra_popup_cls
except Exception:
    pass

# 缓存：命中的窗口句柄列表 + 上次枚举时刻。
# v80b 起这是【周期性重扫】的间隔（不再是"只缓存为空时才扫"）：VBE 的提示窗
# 可能是用到才建的，一旦第一次枚举时它还没出现，旧口径会让它永远进不了缓存。
# 想更省（或排障时想更灵敏）：set VBECOMPLETE_VBE_POPUP_RESCAN=2.0
_vbe_popup_state = {"hwnds": [], "at": 0.0}
try:
    _VBE_POPUP_RESCAN_SEC = _settings.float_value("VBE_POPUP_RESCAN", 1.0)
except Exception:
    _VBE_POPUP_RESCAN_SEC = 1.0
if not (0.05 < _VBE_POPUP_RESCAN_SEC < 60.0):
    _VBE_POPUP_RESCAN_SEC = 1.0

# ---------------------------------------------------------------------------
# v65：让位只针对【成员列表】，不针对【参数信息】
# ---------------------------------------------------------------------------
# 用户口径（v65 澄清 v63 的原意）："只有弹成员列表的时候才让位，弹形参签名
# 不需要让位。"
#
# 为什么该这么分：
#   * NameListWndClass（列出成员 / 列出常量）本身就是一份【候选列表】—— 与我们
#     的候选窗完全同质：同样画在光标下方、同样吃 Tab / ↑↓ / Enter 选词。两个
#     列表叠在一起既遮挡又抢键盘，让位是对的（v55 的原意）。
#   * PopupTipWndClass（参数信息 = 形参签名）只是一行只读提示：不给候选、不吃
#     键盘。而它出现的位置恰恰是"正在填实参"——**最需要候选的时候**。这时候
#     让位，就变成了"这里输入任何字符都无提醒"（v65 报的 bug）。
#
# 想调整（比如某版本的成员列表类名不同）：
#   set VBECOMPLETE_VBE_YIELD_CLASS=ClassA,ClassB
VBE_YIELD_CLASSES = ("NameListWndClass",)
try:
    _extra_yield_cls = _settings.items("VBE_YIELD_CLASS")
    if _extra_yield_cls:
        VBE_YIELD_CLASSES = tuple(VBE_YIELD_CLASSES) + _extra_yield_cls
except Exception:
    pass


def _enum_vbe_popup_windows():
    """枚举顶层窗口，返回类名命中 VBE_POPUP_CLASSES 的 [(hwnd, cls)]。"""
    import ctypes
    u32 = ctypes.windll.user32
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool,
                                     ctypes.c_void_p, ctypes.c_void_p)
    buf = ctypes.create_unicode_buffer(256)
    out = []

    def _cb(hwnd, _lp):
        try:
            if u32.GetClassNameW(hwnd, buf, 256):
                cls = buf.value
                if cls in VBE_POPUP_CLASSES:
                    out.append((hwnd, cls))
        except Exception:
            pass
        return True

    u32.EnumWindows(WNDENUMPROC(_cb), 0)
    return out


def _hwnd_alive(hwnd):
    import ctypes
    try:
        return bool(ctypes.windll.user32.IsWindow(hwnd))
    except Exception:
        return False


def _host_process_id():
    """宿主（Office/WPS 主窗或 VBE 主窗）进程 PID；找不到返回 0。纯 Win32，不碰 COM。"""
    import ctypes
    import ctypes.wintypes
    u32 = ctypes.windll.user32
    fn = u32.FindWindowW
    fn.restype = ctypes.c_void_p
    fn.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p]
    for cls in _HOST_WINDOW_CLASSES:
        try:
            hwnd = fn(cls, None)
        except Exception:
            hwnd = 0
        if hwnd:
            pid = ctypes.wintypes.DWORD()
            u32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            return int(pid.value or 0)
    return 0


def _window_class_name(hwnd):
    import ctypes
    try:
        buf = ctypes.create_unicode_buffer(256)
        if ctypes.windll.user32.GetClassNameW(hwnd, buf, 256):
            return buf.value
    except Exception:
        pass
    return ""


def _popup_rects_now(hwnds):
    """这些句柄里正在显示的 VBE 提示窗，返回 [(类名, 左, 上, 右, 下), ...]。

    逐个复核三件事（句柄值会被 Windows 回收复用给别的窗口，只凭"还记得这个
    句柄"就相信它，一旦复用就会出现"永远以为 VBE 在弹提示、我们的窗再也不弹"
    的灾难）：
      1. 类名仍是 VBE 的提示窗类；
      2. 属于宿主 Excel/VBE 进程；
      3. 可见且尺寸正常。

    返回【位置】而不只是 bool/尺寸：v66 起 UI 要拿它做避让（把候选窗挪到
    参数信息窗下方），所以必须知道屏幕坐标。冷路径开销可忽略：缓存里最多
    两三个句柄，逐句柄 GetWindowRect。
    """
    import ctypes
    import ctypes.wintypes
    u32 = ctypes.windll.user32
    rect = ctypes.wintypes.RECT()
    host_pid = _host_process_id()
    out = []
    for hwnd in hwnds:
        try:
            cls = _window_class_name(hwnd)
            if cls not in VBE_POPUP_CLASSES:
                continue
            if host_pid:
                pid = ctypes.wintypes.DWORD()
                u32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                if int(pid.value or 0) != host_pid:
                    continue
            if not u32.IsWindowVisible(hwnd):
                continue
            if not u32.GetWindowRect(hwnd, ctypes.byref(rect)):
                continue
            l, t = int(rect.left), int(rect.top)
            r, b = int(rect.right), int(rect.bottom)
            if r - l <= 0 or b - t <= 0:
                continue
            out.append((cls, l, t, r, b))
        except Exception:
            continue
    return out


def _popup_details_now(hwnds):
    """这些句柄里正在显示的 VBE 提示窗，返回 [(类名, 宽, 高), ...]。

    与 _popup_rects_now 是同一套判据（这里只是把矩形投影成"宽高"），供诊断
    日志与让位判据使用 —— 明细能区分【成员列表】(NameListWndClass) 与
    【参数信息】(PopupTipWndClass)，两者的结论不同（v65）。
    """
    return [(c, r - l, b - t) for (c, l, t, r, b) in _popup_rects_now(hwnds)]


def _vbe_popup_hwnds():
    """取（并维护）VBE 提示窗的句柄缓存；返回当前可用的句柄列表。

    枚举全量顶层窗口约 1.75ms —— 每轮轮询都枚举不可取（轮询约 20 次/秒），
    所以缓存句柄、平时只查 IsWindowVisible（快三个数量级）。

    ★v80b：缓存必须能【长出】新句柄。原来只在"缓存为空"时重扫，于是只要第一次
    枚举时某个提示窗还没建出来（VBE 的部分提示窗是【用到才建】的），它就永远进
    不了缓存 —— 表现是"我们探不到那个窗"，于是既不让位的兜底、避让也拿不到矩形，
    候选窗就一直压着它（用户报的"我们的弹窗遮挡 VBE 的形参提示列表"，还"不是每次
    都能复现"—— 取决于工具启动那一刻那个窗建出来了没有）。
    现在按 _VBE_POPUP_RESCAN_SEC（默认 1s）周期性重扫，句柄集合变了就换掉；
    代价是每秒一次 EnumWindows（约 0.2% CPU）。
    """
    cached = _vbe_popup_state["hwnds"]
    hs = [h for h in cached if _hwnd_alive(h)]
    if len(hs) != len(cached):
        _vbe_popup_state["hwnds"] = hs         # 有句柄失效 -> 作废重找
    now = time.time()
    if now - _vbe_popup_state["at"] >= _VBE_POPUP_RESCAN_SEC:
        _vbe_popup_state["at"] = now
        found = [h for h, _cls in _enum_vbe_popup_windows()]
        if set(found) != set(hs):
            # 句柄集合变化只在"提示窗被建出来 / 被销毁"时发生，很少 —— 值得记一行，
            # 它一眼能分清"探不到 VBE 提示窗"到底是【窗没建出来】还是【别的判据】。
            _log("vbe-popup: 提示窗句柄集合 %d -> %d 个 %r"
                 % (len(hs), len(found), [_window_class_name(h)
                                          for h in found]))
            _vbe_popup_state["hwnds"] = found
        hs = found
    return hs


def vbe_popup_info():
    """当前可见的 VBE 提示窗明细 [(类名, 宽, 高), ...]；探不到返回 []。

    只读 Win32 查询，不碰 COM。专供诊断日志（区分成员列表 / 参数信息）。
    """
    try:
        return _popup_details_now(_vbe_popup_hwnds())
    except Exception:
        return []


def vbe_popup_visible():
    """VBE 自带的提示窗此刻是否可见（v63）。探不到一律 False（不拦）。

    只读 Win32 查询，不碰 COM。Excel / VBE 没开时自然返回 False。
    ⚠️ 语义是"**任何**提示窗"（含参数信息），只用于诊断与兼容；
    引擎真正据以让位的是 vbe_yield_visible()（v65 起只认成员列表）。
    """
    try:
        return bool(_popup_details_now(_vbe_popup_hwnds()))
    except Exception:
        return False


def vbe_yield_visible():
    """VBE 的【成员列表】窗此刻是否可见 —— 引擎据此让位（v65）。

    只认 VBE_YIELD_CLASSES（默认只有 NameListWndClass）。参数信息
    （PopupTipWndClass）**不算**：它不给候选、不吃键盘，而它出现时用户正在填
    实参，正是最需要候选的时候 —— 在那里让位就成了"输入什么都不弹"（v65）。

    只读 Win32，探不到一律 False（不让位）。
    """
    try:
        for _cls, _w, _h in _popup_details_now(_vbe_popup_hwnds()):
            if _cls in VBE_YIELD_CLASSES:
                return True
        return False
    except Exception:
        return False


def vbe_popup_rects():
    """当前可见的 VBE 提示窗的屏幕矩形 [(类名, 左, 上, 右, 下), ...]（v66）。

    UI 拿它做【避让】：VBE 的参数信息窗（形参签名）就画在光标正下方，而我们的
    候选窗默认也在那儿（光标底边 +2px）—— 两窗一重叠就互相遮挡。知道它画在哪，
    就能把候选窗下移到它下面。

    用户口径（v66）："把所有这种我们候选窗和参数信息重叠的情况，都改成把候选窗
    下移到签名提示的下方"；并明确【自定义函数】的参数信息也要一并处理。
    VBE 对自己工程里的过程同样会弹这个窗（它认识过程签名），所以这里**不区分
    "内建还是自定义"** —— 一律按"可见的提示窗"避让即可，天然覆盖自定义函数。

    只读 Win32、不碰 COM；探不到返回 []（UI 就按原位置显示，绝不让"拿不准"
    影响显示 —— v54 的教训）。
    """
    try:
        return _popup_rects_now(_vbe_popup_hwnds())
    except Exception:
        return []


# ---------------------------------------------------------------------------
# v64：焦点是不是真的在【代码窗格】上
# ---------------------------------------------------------------------------
# 症状（用户报）：在【属性窗口】里改属性值时，候选窗偶尔会冒出来 —— 那里根本
# 不是写代码的地方，用户不希望在这些非代码工作区出现提示。
#
# 根因：判断"是否在 VBE 里"一直是【前台窗口标题里有没有 "Microsoft Visual
# Basic"】（main.in_vbe_code_pane）。而属性窗口 / 工程窗口 / 窗体设计器全都是
# VBE 主窗口（wndclass_desked_gsk）里的子窗口，标题判据对它们一律返回 True。
# 于是两件事都会发生：
#   * 轮询发现代码窗格那行的文本变了（切代码窗格、属性改动导致 VBE 改写代码…）
#     就照样弹候选窗；
#   * 自动配对（敲 `(` / `"`）与 Shift+Enter 也会把内容写进【代码窗格】，
#     而用户此刻其实在属性窗口里打字。
#
# 精确判据（真机实测 2026-09-14，本机 Office VBE）：
#   * 代码窗格的窗口类是 **VbaWindow**，标题形如
#     `下单填写模板.xlsm - UserForm1 (代码)`；它挂在 MDIClient 之下：
#       VbaWindow < MDIClient < wndclass_desked_gsk
#   * 属性窗口的窗口类是 **wndclass_pbrs**（标题 `属性 - CommandButton2`），
#     里面的编辑框 / 列表是 Edit / ListBox / ComboBox / SysTabControl32；
#   * 工程窗口是 **PROJECT**（标题 `工程 - VBAProject`）；
#   * 窗体设计器是 **DesignerWindow**，里面装着被设计的窗体（ThunderDFrame）。
#   所以"焦点窗口的父链上有没有 VbaWindow"是一个干净、与语言无关的判据。
#
# 为什么用 GetGUIThreadInfo【VBE 线程】而不是 GetForegroundWindow：
#   中文输入法的候选/组字窗口会短暂抢走【前台窗口】，但 VBE 线程内部的焦点
#   窗口不受影响 —— 用它就不会在中文输入时把代码窗格误判成"没焦点"。
#
# 探不到（VBE 主窗口找不到 / API 失败 / 焦点窗口为 0）一律返回 None（拿不准），
# 由调用方退回老判据 —— 绝不"拿不准就拦"（v54 就是这么把工具搞成什么都不弹的）。
VBE_FRAME_CLASSES = ("wndclass_desked_gsk",)
try:
    _extra_frame_cls = _settings.items("VBE_FRAME_CLASS")
    if _extra_frame_cls:
        VBE_FRAME_CLASSES = tuple(VBE_FRAME_CLASSES) + _extra_frame_cls
except Exception:
    pass

CODE_PANE_CLASSES = ("VbaWindow",)

_focus_state = {"frame": 0}
_GUI_CACHE = {"cls": None, "ready": False}


def _focus_protos():
    """给这几个 Win32 调用设好原型并缓存（句柄是 64 位，restype 不设会被截断）。"""
    import ctypes
    import ctypes.wintypes as wt
    if _GUI_CACHE["ready"]:
        return ctypes.windll.user32
    u32 = ctypes.windll.user32
    u32.FindWindowW.restype = ctypes.c_void_p
    u32.FindWindowW.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p]
    u32.GetParent.restype = ctypes.c_void_p
    u32.GetParent.argtypes = [ctypes.c_void_p]
    u32.GetWindowThreadProcessId.restype = wt.DWORD
    u32.GetWindowThreadProcessId.argtypes = [ctypes.c_void_p,
                                             ctypes.POINTER(wt.DWORD)]
    u32.GetGUIThreadInfo.restype = wt.BOOL
    u32.GetGUIThreadInfo.argtypes = [wt.DWORD, ctypes.c_void_p]
    u32.IsWindow.argtypes = [ctypes.c_void_p]

    class _GUITHREADINFO(ctypes.Structure):
        _fields_ = [("cbSize", wt.DWORD),
                    ("flags", wt.DWORD),
                    ("hwndActive", wt.HWND),
                    ("hwndFocus", wt.HWND),
                    ("hwndCapture", wt.HWND),
                    ("hwndMenuOwner", wt.HWND),
                    ("hwndMoveSize", wt.HWND),
                    ("hwndCaret", wt.HWND),
                    ("rcCaret", wt.RECT)]

    _GUI_CACHE["cls"] = _GUITHREADINFO
    _GUI_CACHE["ready"] = True
    return u32


def _vbe_frame_hwnd():
    """VBE 主窗口句柄（缓存 + 失效重找）；找不到返回 0。纯 Win32。"""
    u32 = _focus_protos()
    h = int(_focus_state.get("frame") or 0)
    if h and u32.IsWindow(h):
        return h
    h = 0
    for cls in VBE_FRAME_CLASSES:
        try:
            got = u32.FindWindowW(cls, None)
        except Exception:
            got = 0
        if got:
            h = int(got)
            break
    _focus_state["frame"] = h
    return h


def _focused_hwnd_in_vbe():
    """VBE 线程里"当前拥有键盘焦点的窗口"；取不到返回 0。"""
    u32 = _focus_protos()
    frame = _vbe_frame_hwnd()
    if not frame:
        return 0
    tid = int(u32.GetWindowThreadProcessId(frame, None) or 0)
    if not tid:
        return 0
    import ctypes
    gti = _GUI_CACHE["cls"]()
    gti.cbSize = ctypes.sizeof(gti)
    if not u32.GetGUIThreadInfo(tid, ctypes.byref(gti)):
        return 0
    return int(gti.hwndFocus or 0)


def _parent_hwnd(h):
    """父窗口句柄（取不到 / 没有父窗口返回 0）。单独一层，便于单测打桩。"""
    try:
        u32 = _focus_protos()
        return int(u32.GetParent(h) or 0)
    except Exception:
        return 0


def vbe_code_pane_focused():
    """焦点此刻是否落在 VBE 的【代码窗格】上。

    返回：
      True  —— 焦点窗口的父链上有代码窗格（类名 VbaWindow）；
      False —— 焦点在 VBE 里，但不在代码窗格（属性窗口 / 工程窗口 / 设计器…）；
      None  —— 拿不准（VBE 主窗口找不到、API 失败、焦点窗口为空）。

    调用方约定：None 必须退回老判据（前台窗口标题），绝不当成 False 去拦 ——
    否则在类名不符的宿主上会把工具整个哑掉。

    只读 Win32 查询，不碰 COM；每次约 0.02ms（FindWindow 有缓存）。
    """
    try:
        frame = _vbe_frame_hwnd()
        if not frame:
            return None
        h = _focused_hwnd_in_vbe()
        if not h:
            return None
        # v65：VBE 自己的提示窗（成员列表 NameListWndClass / 参数信息
        # PopupTipWndClass）会短暂拿到键盘焦点。它们只可能在"正在编辑代码"时
        # 出现，属于代码窗格的临时浮层 —— 必须算作"还在代码窗格"。否则它们一
        # 冒出来，本判据立刻变 False，把我们的候选窗当场收掉：表现就是"在填实参
        # 的位置输入任何字符都不提示"（v65 报的 bug，第二处根因）。
        if _window_class_name(h) in VBE_POPUP_CLASSES:
            return True
        for _ in range(12):
            if not h:
                break
            if _window_class_name(h) in CODE_PANE_CLASSES:
                return True
            if h == frame:
                break
            p = _parent_hwnd(h)
            if not p or p == h:
                break
            h = p
        return False
    except Exception:
        return None


def _release_vbe_proxy(collect=False):
    """丢弃缓存的 VBE/宿主 COM 代理 —— 这是让宿主能真正退出的关键。

    collect=True 时额外跑一次 gc.collect()：刚写过代码、或确认宿主已退出时
    用，确保 COM 包装对象被立刻回收而不是等 GC 想起来。
    """
    _vbe_cache["obj"] = None
    _vbe_cache["at"] = 0.0
    if collect:
        try:
            import gc
            gc.collect()
        except Exception:
            pass


# COM 调用失败退避（秒）。
#
# Excel 正在关闭 / 正在重启时，每次 GetActiveObject 都可能被 COM 拖住数秒
# （等服务器响应、甚至触发重新激活）。若仍按 100ms 猛冲，表现为：
#   - 关文件要卡 3~10 秒；
#   - 文件关完之后 Excel/Office 又被 COM 唤醒启动一次。
# 因此连续失败时按档退避（比早期版本更狠），成功后立刻恢复。
_COM_BACKOFF = (1.0, 2.0, 5.0, 10.0, 30.0)

# ★v106：宿主窗口门【不提供信息】时的退避上限（秒）。
#
# 什么时候算"不提供信息"：这台机器上从来没探到过任何已知宿主窗口
# （`_wnd_state["seen"]` 一直是 False）—— 也就是 `_host_window_alive()` 判定
# "别拦，可能是类名对不上的宿主（某些 WPS 版本）"的那种情况。
# 那时【窗口门说"放行"、退避却说"不许"】，前后自相矛盾：净效果就是开机后
# 干等最长 30 秒（工具随开机启动、宿主还没开 ⇒ COM 连续失败把退避顶到顶档）。
# 所以门一旦不提供信息，退避也必须跟着松口。
#
# 松到 2 秒安全吗：实测一次失败的 `_get_vbe()`（逐个试 11 个 ProgID）约 1.5ms，
# 2 秒一次 ≈ 0.08% CPU。而"宿主窗口门失效"这个前提本身就已经意味着我们放弃
# 了"靠窗口判断宿主在不在"这条保护（见 _host_window_alive 的兜底说明）。
_COM_BACKOFF_COLD_CAP = 2.0

_com_state = {"fails": 0, "until": 0.0}


def com_backoff_remaining():
    """还需退避多少秒（0 表示可以正常访问 COM）。"""
    return max(0.0, _com_state["until"] - time.time())


def com_reset():
    """重置退避状态（仅供测试 / 手动恢复使用）。"""
    _com_state["fails"] = 0
    _com_state["until"] = 0.0


def _com_backoff_shorten(cap):
    """把当前退避【缩短】到不超过 cap 秒（只缩短，绝不延长）。

    与 com_reset() 的区别：com_reset 连"失败次数"一起清零（等于宣告
    "这台机器没问题了"），而本函数只压这一次的等待时长，失败档位保留 ——
    宿主真的在关闭时仍然能一路升到 30 秒（`_wnd_state["seen"]` 为真时
    根本不会走到这里）。
    """
    now = time.time()
    if _com_state["until"] > now + cap:
        _com_state["until"] = now + cap


def com_host_present():
    """宿主（Excel/Word/PPT/WPS…）的窗口现在在不在 —— 纯 Win32，几微秒。

    ★v106：公开入口，专门给**退避期内**的调用方用。它不只是"问一下"：
    `_host_window_alive()` 里带**副作用** —— 宿主刚从"没有"变回"有"时会
    立刻 `com_reset()` 解除退避。退避最长 30 秒，而用户完全可能是在退避
    窗口【中途】才打开 Excel 的；那时候如果没人去探一下，我们就只能把
    30 秒白等完（开机自启后首次使用就是这个症状）。
    """
    return _host_window_alive()


def _com_ok():
    _com_state["fails"] = 0
    _com_state["until"] = 0.0


def _com_fail():
    _com_state["fails"] += 1
    i = min(_com_state["fails"] - 1, len(_COM_BACKOFF) - 1)
    _com_state["until"] = time.time() + _COM_BACKOFF[i]
    _log("com: 第 %d 次失败，退避 %.1fs"
         % (_com_state["fails"], _COM_BACKOFF[i]))


def _on_main_thread():
    """当前是否在【主线程】上 —— 只有主线程允许碰 COM（见 _get_vbe_cached）。"""
    try:
        return threading.current_thread() is threading.main_thread()
    except Exception:
        return True


_warned_thread = {"n": 0}


def _get_vbe_cached():
    """取 VBE 对象（**短命**代理，TTL 到期主动放手）。

    v57 重大约束：**只允许主线程调用**。
    键盘低层钩子运行在系统创建的【钩子线程】上，那个线程没有 COM 单元、
    也不该被阻塞（钩子回调超时会被 Windows 直接摘掉）。在那里调 COM 必然
    失败，更糟的是失败会走 _com_fail() 把全局退避打成 1s/2s/5s —— 连主线程
    的正常轮询一起挡住，表现为"敲一下 `(` 之后提示要等 2~3 秒才出来"。
    所以这里直接拒绝，并【不记失败】，避免污染退避状态。

    注意与"长缓存"的取舍：复用代理确实省一次 GetActiveObject，但代价是我们
    的进程会一直持有宿主的 Application —— 而只要持有它，宿主就退不干净
    （关完又启动 Office），且 VBAProject 变脏时的写回会被拖成 3~10 秒。

    所以这里只做**短缓存**：TTL（默认 1 秒）一到就丢弃，让宿主随时有
    机会干净退出；重建成本只是一次 ROT 查询，对打字毫无影响。

    取到后先访问一次 ActiveCodePane 验证仍可用；宿主/VBE 已关闭会抛异常，
    此时丢弃缓存并返回 None。
    """
    if not _on_main_thread():
        # 钩子线程（或其它后台线程）来要 COM：直接拒绝，且【不】计失败 ——
        # 责任在调用方（它本就不该在这里调），不该让主线程的正常访问连坐。
        if _warned_thread["n"] < 3:
            _warned_thread["n"] += 1
            _log("com: 非主线程请求 COM -> 拒绝（不记失败，请改在主线程调用）")
        return None
    # ★v106：宿主探测【必须排在退避检查之前】。
    #
    # `_host_window_alive()` 自带"宿主回来了就立刻解除退避"的副作用；排在
    # 退避检查后面的话，退避期内根本调不到它，"用户打开了 Excel"这件事就
    # 永远发现不了 —— 只能把最长 30 秒的退避熬完。"开机自启后要等半分钟
    # 才出提示"就是这么来的：开机时 Excel 没开，COM 连续失败把退避顶到
    # 30 秒，之后每个 30 秒窗口里我们连一次都不去看宿主在不在。
    #
    # 代价：退避期每轮多 11 次 FindWindowW（几微秒），可忽略。
    # 宿主窗口门【每次都查】：FindWindow 是纯 Win32 调用，开销可忽略。
    # 刻意放在缓存检查【之前】——只在新建立时才查的话，宿主已经退出、我们却
    # 还在用旧代理继续打 COM，照样会把正在退出的 Excel 拽住。
    host_alive = _host_window_alive()
    if com_backoff_remaining() > 0:
        # ★v106：宿主窗口门【不提供信息】时退避也要松口。
        # "seen 一直是 False" = 这台机器上从没见过任何已知宿主窗口（很可能是
        # 类名对不上的宿主）。那种情况下 `_host_window_alive` 的兜底已经判定
        # "别拦"了，退避若还照旧卡 30 秒，等于把"打开宿主后要等半分钟"原样
        # 留在 WPS 这类宿主上。
        if not _wnd_state["seen"]:
            _com_backoff_shorten(_COM_BACKOFF_COLD_CAP)
        return None
    if not host_alive:
        if _wnd_state["miss"] >= 2:
            _release_vbe_proxy(collect=True)
            _log("com: 宿主窗口已消失 -> 释放 COM 代理并暂停访问")
        _com_fail()
        return None
    v = _vbe_cache.get("obj")
    if v is not None:
        # TTL 从"创建时刻"起算，命中不续期 —— 否则只要一直在打字就永不超时，
        # 又变回长期持有。TTL <= 0 表示彻底不缓存（每次新建、用完即弃）。
        expired = (VBE_PROXY_TTL <= 0
                   or time.time() - _vbe_cache.get("at", 0.0) > VBE_PROXY_TTL)
        if expired:
            _release_vbe_proxy()
            v = None
        else:
            try:
                v.ActiveCodePane
                _com_ok()
                return v
            except Exception:
                _release_vbe_proxy()
                _com_fail()
                v = None
    try:
        v = _get_vbe()
    except Exception:
        _com_fail()
        return None
    _vbe_cache["obj"] = v
    _vbe_cache["at"] = time.time()
    _com_ok()
    return v


def _classify(line_text, caret_col):
    """判断光标前是否处于字符串/注释内。

    caret_col 必须是「字符列」(1-based)。若传入 VBE 原始的「显示列」，
    行内含全角字符时它会大于 len(line_text)，下面的 line_text[i] 就会
    IndexError——所以调用方必须先完成显示列->字符列换算。
    这里再做一次钳制兜底，保证本函数本身永不因越界抛异常。

    ★v103 修两个真 bug（都出在"注释"这一路）：

    ① **不认 `Rem`**：原实现只把 `'` 当注释，于是 `Rem 每个元素` 这样的行
       被判成"代码"（`in_comment=False`）⇒ `engine.trigger` 不再静默、
       **在注释里照弹候选**；`insert_pair` 的"注释里不自动配对"早退也失效。
       同一个文件里的 `_is_comment_only`、parser 的扫描都认 Rem，唯独这里
       不认 —— 又一处"同一语义两/三份实现"。现在复用 engine 的
       `_in_rem_comment`（它是 Rem 判据的唯一实现，见其 docstring）。

    ② **遇到 `'` 后不停下**：原实现只记 `comment_pos`、继续往后扫，于是注释
       里的引号会把 `in_str` 翻成 True —— `x = 1 ' say "hi` 会返回
       `(True, False)`（"在字符串里"），而正确答案是 `(False, True)`（注释）。
       现在一进注释立刻定案返回（注释【到行尾】，后面不可能再有代码）。

    ⚠️ 顺序很关键：先按字符串状态机走一遍（含 `''` 转义），**走出字符串之后**
       才问 Rem —— 否则 `x = "a Rem b"` 这种"字符串里长得像 Rem"的会被误判成
       注释（实测：字符串状态机会给出 in_str=True，直接跳过 Rem 那一问）。
    """
    n = caret_col - 1 if caret_col > 1 else 0
    if n > len(line_text):
        n = len(line_text)
    in_str = False
    i = 0
    while i < n:
        c = line_text[i]
        if c == '"':
            if in_str:
                if i + 1 < n and line_text[i + 1] == '"':  # 转义引号
                    i += 2
                    continue
                in_str = False
            else:
                in_str = True
        elif c == "'" and not in_str:
            # ★v103：注释到行尾 —— 后面不可能再有代码/字符串，直接定案。
            return False, True
        i += 1
    if not in_str and _in_rem_comment(line_text, caret_col):
        return False, True
    return in_str, False


def _proc_owns_line(cm, name, line_no):
    """过程 name 是否真的把第 line_no 行圈在里面（归属校验）。"""
    try:
        return vba_parser.proc_owns_line(cm.Lines(1, line_no), line_no, name)
    except Exception:
        return False


def _proc_of_line(cm, line_no):
    """返回该行所在的过程名；不在任何过程内（模块声明区）则返回 None。

    优先用「代码文本往上找最近的 Sub/Function/Property」（纯 Python，
    行为确定、可单测）；COM 的 ProcOfLine 只作兜底——它在 pywin32 下
    的 byref 参数（ProcKind）常常抛异常而拿不到值。

    ⚠️ v79d：光标【自己这一行】就是过程头时，一律按"没有当前过程"（None）算。
    判据是 `parser.is_proc_header_line`，必须放在最前面 —— 它要同时压住下面
    两条支路（文本扫描会把这一行自己当成过程头；COM 兜底也会这么认，且
    `proc_owns_line` 对过程头那一行恰好返回 True，拦不住）。不这么做的话：
    正在写新过程头 `Sub test`、而模块里已有 `Sub test()` 时，那一行的
    "当前过程"就是 test，test 的局部变量（targetsheet）会被当成当前过程的
    提示出来 —— 用户报的泄漏。过程是"从过程头【之后】开始"的，头那一行是声明。

    ⚠️ v79c 修两处（都在兜底那一支，但都直接影响"作用域"判断）：
      1. pywin32 下 `ProcOfLine` 返回的是【元组】(名字, ProcKind)（ProcKind 是
         byref 出参），旧代码直接 `str(name)` 得到的是 `"('插入工分', 0)"` 这种
         垃圾字符串。它不是 None，于是 ctx["proc_name"] 在【模块级/End Sub 行】
         上变成一个既非空、又不等于任何真实过程名的值 —— 作用域过滤里
         `scope_only` 会被它误判成 True（本该按 v37 口径"模块级声明不受限"），
         过滤阈值也跟着错位。
      2. 更要命的是【不能盲信 VBE】：实测它会把手伸到过程外 —— `End Sub`
         那一行、以及它下面直到下一个过程头之间的所有行，都归给【前一个】
         过程；模块声明区里第一个过程头【之前】的行，归给【第一个】过程。
         照单全收就等于：在模块级 / 新起一个函数的位置上，前一个函数的局部
         变量全部变成"当前过程的"——正是用户报的"别的函数的变量泄漏到本
         位置"。所以拿到名字后必须做一次归属校验（proc_owns_line）：该过程头
         在光标上方、且两者之间没有它的收尾行，才算真的在里面。

    ⚠️ v79e：返回的是【作用域键】，正常情况下就是过程名；同一模块里存在同名
    过程时带 `#序号`（见 parser._proc_scope_key）。那种状态下"过程名"不足以把
    两个过程分开 —— 记录侧也是这么写键的，两边必须一致，否则第二个过程里会
    提示出第一个过程的局部变量（用户报的"过程级的变量都泄露到其他过程"）。
    """
    # 0) v79d：光标这一行自己就是过程头 -> 不在任何 procedure body 里
    try:
        if vba_parser.is_proc_header_line(cm.Lines(line_no, 1)):
            return None
    except Exception:
        pass

    # 1) 文本扫描：取 1..line_no 行，从光标行往上找最近的过程头
    try:
        head = cm.Lines(1, line_no)
        name = vba_parser.proc_at_line(head, line_no)
        if name:
            return str(name)
    except Exception:
        pass

    # 2) 兜底：COM ProcOfLine（vbext_pk_Proc=0 / Let=1 / Set=2 / Get=3）
    for kind in (0, 1, 2, 3):
        try:
            name = cm.ProcOfLine(line_no, kind)
        except Exception:
            continue
        if isinstance(name, (tuple, list)):
            name = name[0] if name else ""      # (名字, ProcKind)
        name = str(name or "").strip()
        if not name:
            continue
        if _proc_owns_line(cm, name, line_no):  # 归属校验：别信"邻居过程"
            # v79e：同名过程（模块里两个 `Sub test`）的键带 `#序号`，与记录侧
            # 保持一致；文本扫描取不到键就退回裸名（宁可少提示，也不乱提示）。
            try:
                key = vba_parser.proc_at_line(cm.Lines(1, line_no), line_no)
                if key:
                    return str(key)
            except Exception:
                pass
            return name
    return None


# VBIDE 组件类型常量（只关心"标准模块"与"其他"两类）
VBE_CT_STDMODULE = 1      # vbext_ct_StdModule
VBE_CT_MSFORM = 3         # vbext_ct_MSForm（UserForm 窗体）

# 取窗体控件名时最多列出的控件数。正常窗体都在这个量级以内；
# 设上限只是为了让"设计器接口异常返回天量条目"时不会把候选池撑爆。
_FORM_CTL_LIMIT = 300


def _is_std_module(comp_type):
    """标准模块（.bas）之外，类模块/窗体/文档模块的成员都只能
    实例或限定名访问，跨模块不做裸名提示。"""
    try:
        return int(comp_type) == VBE_CT_STDMODULE
    except Exception:
        return True   # 取不到类型时按标准模块处理（最宽松，不误伤）


def _is_msform(comp_type):
    """是否 UserForm 窗体组件。"""
    try:
        return int(comp_type) == VBE_CT_MSFORM
    except Exception:
        return False


def _form_control_names(comp):
    """窗体（UserForm）设计器里的控件名列表。

    【为什么必须走 Designer】控件名只存在于【设计器数据】里，窗体代码模块
    （CodeModule）里通常一个字都看不到 —— 除非那个控件挂过事件过程
    （那种情况下才会以 `Label1_Click` 的形式出现在代码里）。所以用户"刚拖
    上去一个 Label1，希望输入 la 就能提示"只能从这里拿。

    【为什么不递归】实测 `Designer.Controls` 已经是【扁平全集】：容器控件
    （Frame / Page）自身也在列表里，它们的子控件同样被列出。真机验证：一个
    含 Frame1/Frame2 的窗体，顶层 9 条已经覆盖了 Frame 里的全部控件；若再
    对 Frame 递归一次，Label1/TextBox1/BOMList 会被重复收集一遍（这也是
    VBA 里控件名必须全窗体唯一的语义所决定的）。所以只取顶层，一次到位。

    任何一步失败都只返回已收到的部分：窗体处于运行态、宿主不支持 Designer
    接口时都可能取不到，绝不能因此让整个候选收集流程崩掉。
    """
    out = []
    try:
        cs = comp.Designer.Controls
        n = int(cs.Count)
    except Exception:
        return out
    if n > _FORM_CTL_LIMIT:
        n = _FORM_CTL_LIMIT
    for i in range(n):
        try:
            nm = str(cs.Item(i).Name)
        except Exception:
            continue
        if nm:
            out.append(nm)
    return out


def _caret_char_col(cm, cp, line_no, line_text, ec):
    """VBE 给的光标列 ec（显示列）-> (字符列 1-based, sem, tabw, wide2)。

    VBE 的 ec 是显示列（中文等全角占 2 格），而 line_text 是字符序列；行内含
    Tab / 全角时必须换算，否则索引会越过行尾或落在中文中间。换算结果连同本次
    的列语义一起返回，供"写回后摆光标"复用同一套算法。

    判定顺序（越靠前越安全，与 get_context 一致）：
      1) 被动判定：完全不移动光标（打字时光标多在行尾，这条最常用）；
      2) 会话缓存：同一编辑器探测一次后长期复用；
      3) 才移动光标探测（_probe_semantics 会自动还原+校验，还原不可靠的编辑器
         会永久禁用探测，避免光标乱跳）。
    """
    if ec <= 0:
        return 1, "char", 4, False
    if not ("\t" in line_text or any(_is_wide(c) for c in line_text)):
        return min(ec, len(line_text) + 1), "char", 4, False
    info = _detect_semantics_passive(line_text, ec)
    if info is None:
        info = _sem_cache.get("info")
    if info is None:
        info = _probe_semantics(cm, cp, line_no, line_text)
    if info is None:
        info = ("char", 4, False)
    # ⚠️★v95 已知残留（别再凭直觉加"重新探测"去治它）：走到这里时必然
    # `ec <= len+1`（否则第 1 步的被动判定就出结果了），而"字符列"在这个范围里
    # **永远不可能被推翻** —— 也就是说，缓存若被污染成 ("char",4,False)，本函数
    # 在 ec 偏小时看不出来，会把【显示列】当成字符列返回（少算全角格数）。
    # 想治得靠"缓存只在同一类行之间复用"那一层（见 `apply_completion` 的做法：
    # 不信任缓存，在本行上重新量），而不是在这里加判据。
    sem, tabw, wide2 = info
    if sem == "disp":
        return (_col_to_char_index(line_text, ec, tabw, wide2) + 1,
                sem, tabw, wide2)
    return min(ec, len(line_text) + 1), sem, tabw, wide2


def _locate_by_prefix(after, written, col0):
    """VBE 把整行重写过一次之后，在 `after` 里找回 `written[:col0]` 这个
    **前缀**的位置，返回对应的 0-based 字符下标；认不出来返回 None。★v99。

    为什么要有它：VBE 的自动语法检测会重排整行（`x=1` -> `x = 1`、
    `if a then` -> `If a Then`），行一变长，写入前算好的光标位置就偏了。

    与 `_locate_inserted` 的差别：那一条找的是"我们**插入**的那对符号"，指纹
    是 `written[:col0]`；这一条找的是"我们**删除**之后前缀结束的位置" ——
    删除不会留下任何可搜的标记，唯一能认的指纹就是前缀本身。

    按压平后的前缀找（`_norm_code`：去空白 + 小写）：VBE 在任何位置补空格 /
    改大小写都不影响指纹。

    ⚠️★v99 前缀在行里可能**出现多次**（`aa = aa`、`    x = x`），此时取
    **离 `col0` 最近**的那一个，而不是"最长"或"第一个"。我第一版写了
    "从后往前扫取最长" —— 实测 `aa = aa` 会命中行**尾**那个 `aa`（压平后与
    前缀一样），把光标从行首扔到行中。改成"最近"就与 `_locate_inserted`
    同一口径：VBE 重排只会挪动几个字符，不会把光标搬到半行之外。
    """
    try:
        pre = _norm_code(written[:col0])
        if pre == "":
            # 前缀是空的（要落到行首）：行首就是答案。
            return 0
        hits = []
        for k in range(len(after) + 1):
            if _norm_code(after[:k]) == pre:
                hits.append(k)
        if hits:
            near = min(col0, len(after))
            return min(hits, key=lambda x: abs(x - near))
        return None
    except Exception:
        return None


def _snippet_tail_pos(actual, tail):
    """★v115：片段插进整行之后，光标该停的位置（0-based）= "尾部原文"的起点。

    `insert_snippet` 的落点模型是**整行拼接**：`前半 + 片段首行` … `片段末行 + 后半`。
    写完之后 VBE 可能重排过最后那一行（补空格 / 改大小写），所以不能拿
    `len(片段末行)` 当答案，只能反过来找"后半原文"在行里的位置：

      1) 行尾就是 `tail`  -> 长度相减（最常见：后半为空 => 落行尾）；
      2) 否则找 `tail.strip()` 最后一次出现（VBE 可能吞了/加了两侧空白）；
      3) 都找不到 -> 退到行尾。

    纯函数，便于单测（§99.4）。
    """
    try:
        actual = actual or ""
        if not tail:
            return len(actual)
        if actual.endswith(tail):
            return len(actual) - len(tail)
        t = tail.strip()
        if t:
            p = actual.rfind(t)
            if p >= 0:
                return p
        return len(actual)
    except Exception:
        return 0


def _pair_insertion(line_text, col0, open_ch, close_ch):
    """算出自动配对后的 (新行文本, 光标字符下标 0-based, 是否插了完整一对)。

    新行文本为 None 表示"一个字符都不用写，只把光标挪到 caret_off"—— 右半边
    已经在光标右边时（打完 `s = "abc` 再按 `"`），再插一对就会变成 `""`，
    此时应当像现代编辑器那样"跨过去"。

    第三元 `paired` = "这次插入的文本里含一对完整符号"（`()` / `""`），供
    调用方在 VBE 重排整行后**选对定位方式**：v112 起括号可能只插一个左括号，
    那种行里根本没有 `()` 可找，得改按"压平前缀"定位（见 insert_pair）。

    引号是特殊的一个键（既开又闭），所以额外看【光标左边引号数的奇偶】：
    奇数说明正处在未闭合的字符串里，这一下是收尾的那一半，只补一个。

    ★v108：全角引号 `“`/`”` 是**两个不同的字符**（半角 `"` 开收同字），
    所以"在不在字符串里"改成比 `“` 与 `”` 的个数（`“` 多出来就是没闭合）。
    其余逻辑（`（`/`）`）与半角完全同路。

    ★v112（用户口径）：**敲左括号时先看右侧还有没有内容，有就只插左括号**、
    不替用户补右半边。光标在 `foo(|bar)` 这种"行中间 / 标识符之前"时，补出
    `()` 会把后面的代码顶走；现代编辑器的通用规则也是"右侧不是空白/行尾就
    不自动闭合"。右侧只剩空白（或已在行尾）才补一对。
    ★v113：**但光标在实参表 / 形参表里时例外，照旧补一对** ——
    `foo(bar(`（在 `foo` 的实参里开一层新的调用）、`Sub Foo(ByVal arr(`（形参
    数组）右侧那个 `)` 是**外层**的，v112 一律"只插左"会让用户每次都得自己
    补一个右括号。判据是 `parser.enclosing_call_paren`（纯函数，只吃这一行）。
    ⚠️ 引号**刻意不走这条**：`w""` 那种"右侧有内容仍补一对引号"是既有行为
    （见 §35.23 的 `    w=4` -> `    w ""=4`），且引号另有"在不在字符串里"
    的先判据，不是同一个问题。
    """
    if open_ch in ('"', "“"):
        if col0 < len(line_text) and line_text[col0] == close_ch:
            return None, col0 + 1, False                # 跨过去，不写文本
        if open_ch == "“":
            _pre = line_text[:col0]
            inside = _pre.count("“") > _pre.count("”")
        else:
            inside = line_text[:col0].count('"') % 2 == 1
        if inside:
            # 补收尾那一半：只有一个字符，算不上"一对"
            return line_text[:col0] + close_ch + line_text[col0:], col0 + 1, False
        return (line_text[:col0] + open_ch + close_ch + line_text[col0:],
                col0 + 1, True)
    if open_ch in ("(", "（"):
        if (line_text[col0:].strip()
                and vba_parser.enclosing_call_paren(line_text, col0) is None):
            # ★v112：右侧还有内容 -> 只插左括号（右半边用户自己敲；那时若
            # 右边恰好是 `)` 还会走"跨过去"，不会多出一个）。
            # ★v113：除非光标正待在实参表 / 形参表里 —— 那是"开一层新的"，
            # 右侧内容是**外层**的，照旧补一对（判据见 enclosing_call_paren）。
            return (line_text[:col0] + open_ch + line_text[col0:],
                    col0 + 1, False)
        return (line_text[:col0] + open_ch + close_ch + line_text[col0:],
                col0 + 1, True)
    # 兜底：将来若给 _PAIR_CLOSE 添了新符号，默认仍按"补一对"处理
    return (line_text[:col0] + open_ch + close_ch + line_text[col0:],
            col0 + 1, True)


def _norm_code(s):
    """把一行代码压成"只看内容"的形式：去掉所有空白 + 统一小写。

    只用于【在 VBE 重排过的行里找回我们的插入点】：VBE 的自动语法检测会补空格
    （`x=1` -> `x = 1`、`foo(1,2)` -> `foo(1, 2)`）、改大小写（`if a then` ->
    `If a Then`、`rgb` -> `RGB`）。这些都是同一条代码的另一种写法，压平之后两边
    的前缀才可比。
    """
    return "".join(ch for ch in s if not ch.isspace()).lower()


def _locate_inserted(after, written, open_ch, close_ch, col0):
    """在【VBE 改写过的行】里重新找出我们插入的那对符号，返回光标该落在的
    位置（0-based："开符号之后"）；找不到返回 None。

    为什么需要：VBE 有「自动语法检测」，会把整行重新格式化 —— 等号两边补空格
    （`x=1` -> `x = 1`）、关键字改大小写（`if a then` -> `If a Then`）、参数逗号
    后补空格（`foo(1,2)` -> `foo(1, 2)`）、标识符与引号之间补空格
    （`w""` -> `w ""`）。行一变长，写入前算好的光标位置就偏了，表现为「光标
    没落在括号/引号中间」。真机实测（v58）：`    w=4` 里 `w` 之后敲 `"`，得到
    `    w "" = 4`，光标停在 7（第一个引号之前）而不是 8（两引号之间）。

    ★v79 换成本法：**按"压平后的前缀"直接定位**（见 _norm_code）。我们插入的
    位置左边是用户已经打好的代码（written[:col0]），把它压平（去空白 + 小写）
    当指纹，再去 after 里找"左边压平后和这枚指纹一模一样"的那一对符号。好处：
      * VBE 在【任何位置】补空格 / 改大小写都不影响指纹；
      * 一行里有好几对括号时也不会认错 —— 用户报的正是这种行：
        `targetSheet.Cells(i, "b").Interior.Color = RGB()`，`Cells(...)` 是
        第一对、`RGB()` 是第二对，把"第几对"或纯差分对齐一挪就会指到别处。
    差分对齐（difflib，v58 的老办法）留作兜底：指纹也对不上（VBE 真改了内容）
    时再退回它。
    """
    want = open_ch + close_ch
    try:
        pre = _norm_code(written[:col0])
    except Exception:
        pre = None
    # ---- 首选：压平前缀指纹 ----
    if pre is not None:
        try:
            hits = []
            k = after.find(want)
            while k >= 0:
                if _norm_code(after[:k]) == pre:
                    hits.append(k)
                k = after.find(want, k + 1)
            if hits:
                # 理论上只该命中一处（指纹含整行前缀）；真撞上重复时取离原位
                # 最近的 —— VBE 重排只挪几个字符，不会把这一对搬到半行之外。
                near = min(col0, len(after))
                return min(hits, key=lambda x: abs(x - near)) + 1
        except Exception:
            pass
    # ---- 兜底一：差分对齐（v58 的老办法） ----
    try:
        sm = difflib.SequenceMatcher(None, written, after, autojunk=False)
        p = None
        for tag, i1, i2, j1, _j2 in sm.get_opcodes():
            if tag == "equal":
                if i1 <= col0 < i2:
                    p = j1 + (col0 - i1)
                    break
            elif i1 <= col0 <= i2:      # replace / insert / delete
                p = j1
                break
        if p is not None:
            if after[p:p + len(want)] == want:
                return p + 1
            if after[p:p + 1] == open_ch:
                return p + 1
            # 对齐点附近再找一下（VBE 把一对符号拆开或挪过位置）
            k = after.find(want, max(0, p - 2), p + len(want) + 3)
            if k >= 0:
                return k + 1
    except Exception:
        pass
    # ---- 兜底二：连那一对符号都没了（VBE 把它删/改没了），至少把光标放到
    #      "前缀压平后对齐"的位置上 —— 也就是用户刚打完那段代码之后，绝不乱扔。
    if pre is not None:
        try:
            for k in range(len(after), -1, -1):
                if _norm_code(after[:k]) == pre:
                    return k
        except Exception:
            pass
    return None


def enum_family_prefix(names):
    """一个类型库的枚举成员里，覆盖率达标的最长【小写家族前缀】（长度 >= 2）。

    用来回答"这个库的枚举常量有没有统一的写法"：
      * Excel 库 -> "xl"（97% 的成员都以 xl 打头）—— 用户本来就得敲 xl 才找得到；
      * Office 库 -> "mso"（83%）；
      * stdole  -> ""（Checked / Gray / Color 各说各话，没有共同前缀）。

    没有家族前缀的库【整库丢弃】：那种库的成员都是 Checked / Default / Color
    这类通用词，收进来只会污染用户自己的变量候选（用户口径：宁可少提示，也不
    要噪音）。

    为什么用"覆盖率"而不是"最长公共前缀"：Excel 库里还夹着 rgb* / sigdet* 等
    零散成员，最长公共前缀只有 1 个字符，判不出任何东西。取"能盖住一半以上成员
    的最长前缀"才能稳稳落在 xl / mso 上。

    前缀越长覆盖率越低（子集关系），所以一旦某长度不达标就可以直接停。
    """
    lows = [str(n).lower() for n in names if n]
    if not lows:
        return ""
    total = len(lows)
    best = ""
    for k in range(2, 25):
        counts = {}
        for x in lows:
            if len(x) >= k:
                p = x[:k]
                counts[p] = counts.get(p, 0) + 1
        if not counts:
            break
        p, n = max(counts.items(), key=lambda kv: kv[1])
        if n < total * _HOST_ENUM_FAMILY_PCT:
            break
        best = p
    return best


def host_enum_family(names):
    """(家族前缀, 该前缀下的成员名列表)。没有家族前缀则 ("", [])。"""
    p = enum_family_prefix(names)
    if not p:
        return "", []
    if _HOST_ENUM_FAMILIES and p not in _HOST_ENUM_FAMILIES:
        return "", []      # 白名单模式：不在清单里的家族整库跳过
    return p, [str(n) for n in names
               if str(n).lower().startswith(p)]


def _tlb_member_name(ti, memid):
    """ITypeInfo.GetNames 的兼容封装（各版本 pywin32 的参数个数不同）。"""
    try:
        return ti.GetNames(memid)
    except TypeError:
        return ti.GetNames(memid, 256)


def _tlb_enum_members(guid, major, minor):
    """按 GUID + 版本加载已注册的类型库，返回其【常量枚举】的全部成员名。

    只读、不碰 VBE 内容；任何一步失败都返回 []（绝不抛 —— 这条路上出异常会
    让整个候选池收集失败，得不偿失）。

    注意 VBA 自己的库加载不了：它没有注册成类型库（在 VBE7.DLL 里），
    LoadRegTypeLib 会报"库没有注册"。VBA 自带的名字本来就有 vba_builtins.py
    那份静态清单兜着，不依赖这里。
    """
    try:
        import pythoncom
        import pywintypes
    except Exception:
        return []
    try:
        tlb = pythoncom.LoadRegTypeLib(
            pywintypes.IID(str(guid)), int(major), int(minor), 0)
    except Exception:
        return []
    out = []
    try:
        count = tlb.GetTypeInfoCount()
    except Exception:
        return []
    for i in range(count):
        try:
            ti = tlb.GetTypeInfo(i)
            attr = ti.GetTypeAttr()
            if int(attr.typekind) != int(pythoncom.TKIND_ENUM):
                continue
            for j in range(attr.cVars):
                try:
                    vd = ti.GetVarDesc(j)
                    nm = _tlb_member_name(ti, vd.memid)[0]
                except Exception:
                    continue
                if nm and _RE_PLAIN_IDENT.match(nm):
                    out.append(nm)
        except Exception:
            continue
    return list(dict.fromkeys(out))


# ★v118：引用清单的时间缓存。读一次 References 要 ~22ms（每个 Item 都是新的
# COM 包装对象，pywin32 要重新 Bind），而"引用了哪些类型库"基本不变 ——
# 原来是每轮刷新都重读一遍，纯属白花。30 秒兜底一次，用户增删引用最多滞后
# 这么点时间才看到新的 xl* 枚举常量。
_HOST_REFS_TTL = 30.0
_host_refs_cache = {"ts": 0.0, "refs": None}


def _host_enum_refs():
    """当前 VBE 里各工程引用的类型库 -> [(名称, GUID, 主版本, 次版本), ...]。

    读不到（Excel 没开 / COM 失败）返回 None —— 与"读完发现没有引用"区分开，
    调用方据此决定要不要缓存结果。
    """
    if (MOD_CACHE_ENABLED and _host_refs_cache["refs"] is not None
            and (time.time() - _host_refs_cache["ts"]) < _HOST_REFS_TTL):
        return _host_refs_cache["refs"]
    try:
        vbe = _get_vbe_cached()
        if vbe is None:
            return None
        projects = []
        try:
            active = vbe.ActiveVBProject
        except Exception:
            active = None
        if active is not None:
            projects.append(active)
        else:
            try:
                col = vbe.VBProjects
                for i in range(1, col.Count + 1):
                    try:
                        projects.append(col.Item(i))
                    except Exception:
                        continue
            except Exception:
                projects = []
        refs = []
        seen = set()
        for pj in projects:
            try:
                col = pj.References
                n = col.Count
            except Exception:
                continue
            for i in range(1, n + 1):
                try:
                    r = col.Item(i)
                    key = (str(r.GUID).upper(), int(r.Major), int(r.Minor))
                    if key in seen:
                        continue
                    seen.add(key)
                    refs.append((str(getattr(r, "Name", "") or ""),) + key)
                except Exception:
                    continue
        _host_refs_cache["ts"] = time.time()
        _host_refs_cache["refs"] = refs
        return refs
    except Exception:
        return None


# 宿主枚举常量的进程级缓存：加载三个类型库约 15ms，而引用清单几乎不变，
# 所以按"引用签名"缓存 —— 用户新增/删除引用后签名一变就重新枚举一次。
_host_enum_cache = {"key": None, "items": []}


def host_enum_constants():
    """工程引用的类型库里那些枚举常量 -> [(规范名, 最短输入长度), ...]。

    最短输入长度 = 该库的家族前缀长度（xl / mso 各 2、3）。引擎据此要求
    "输入至少这么长、且候选从头开始以它开头"才提示这类名字 ——
    于是 xl -> xl*、xlu -> xlUp，而 ms 只出 MsgBox、绝不带出 mso*。

    顺序稳定、已去重；取不到一律返回 []（引擎行为与旧版完全一致）。
    """
    if not ENABLE_HOST_ENUMS:
        return []
    if _settings.flag("NO_HOST_ENUMS", False):
        return []
    try:
        refs = _host_enum_refs()
    except Exception:
        refs = None
    if not refs:
        return []
    key = tuple(sorted(refs))
    if _host_enum_cache["key"] == key:
        return _host_enum_cache["items"]
    items = []
    fams = []
    for name, guid, major, minor in refs:
        members = _tlb_enum_members(guid, major, minor)
        if not members:
            continue
        fam, kept = host_enum_family(members)
        if not fam:
            _log("host-enums: %-12s 无家族前缀 -> 整库跳过（%d 个枚举成员）"
                 % (name, len(members)))
            continue
        fams.append("%s->%s(%d)" % (name, fam, len(kept)))
        for m in kept:
            items.append((m, len(fam)))
    # 去重：同名的以先出现的为准（Excel 里也有 xl* 同名项）
    uniq = {}
    for m, ml in items:
        uniq.setdefault(m.lower(), (m, ml))
    out = list(uniq.values())
    _host_enum_cache["key"] = key
    _host_enum_cache["items"] = out
    _log("host-enums: %d 条（%s）" % (len(out), " ".join(fams)))
    return out


class VbeBackend:
    def __init__(self, send_char=None, send_quick_info=None):
        # ★v103b：把字符"还给系统"的注入函数（main.send_char）。只在
        # "关键字补空格"这一处用到 —— 真机实测 VBE 会把 ReplaceLine 写进去的
        # 行尾空格 trim 掉，唯一可行的补法是注入一个空格按键。
        # 传 None（旧式构造 / 离线测试）时该功能静默不生效，绝不影响其它路径。
        self._send_char = send_char
        # ★v107c：唤出 VBE 自己【参数信息】窗的注入函数（main.send_quick_info，
        # 就是替用户按一下 Ctrl+Shift+I）。只在"我们刚补出一对 `()`、光标落在
        # 括号中间"之后用一次 —— 参数信息是 VBE 原生能力，我们不重画。
        # 同样地：传 None 就不做，绝不半途出错（见 AUTO_PARAM_INFO 处的说明）。
        self._send_quick_info = send_quick_info
        # ★v110b/2：这里原先还有一个 `send_tab`（替用户按一下 Tab 补缩进），
        # 真机实测注入的 Tab 不生效（见文件顶部 v110b/2 注释）⇒ 改成直接把
        # 缩进写进新行，注入函数随之撤掉。
        self._cache = None
        self._cache_time = 0
        self._declared_names = set()
        # ★v98：全工程的过程名（小写）—— Sub / Function / Property / Declare。
        # 与 _declared_names 的差别：后者是"Dim/Const/Sub/Function/Type/Enum 都算"
        # 的合集，**分不出变量与函数**；补全补括号只对后者里的过程/函数该做
        # （`Dim n As Long` 补成 `Long()` 直接编译不过）。收集时由
        # parser.extract_records 一并给出，不额外读代码、不额外打 COM。
        self._proc_names = set()
        # ★v103c：{过程名(小写): "sub" / "function" / "property"}。
        # 与 _proc_names 同源同刷新（同一次 extract_*、同一个 TTL）—— 补全时
        # 用来决定"按 Tab 确认之后后面跟空格还是跟括号"：Sub 是无返回值的语句、
        # 标准写法不带括号（补了 VBE 也会删），Function 要括号。
        self._proc_kinds = {}
        # ★v98：{类型名(小写): (收集时间戳, 成员名小写集合)}。
        # 与 _cache_time 同一个 TTL（get_members_of 里比对）。
        self._members_of_cache = {}
        # 收集标识符那一刻，光标处那个词的原文（小写）。用来识别【回声】：
        # 用户正在输入 / 正在回退删除的词，不该被当成工程里的真名字。
        self._caret_word = ""
        # 名字(小写) -> 声明它的模块名集合(小写)。让引擎能区分
        # 【声明在别的模块】（跨模块 Public，可当证据）与【只在我正在编辑的
        # 这个模块里声明过】（一律以现场文本为准，见 declared_elsewhere）。
        self._declared_by_module = {}
        # 由【工程结构 / 语言运行时】而非代码文本决定其存在的名字（小写）：
        #   * 组件名（标准模块 / 窗体 / 类模块 / ThisWorkbook、Sheet1 等文档模块名）
        #     与窗体控件名 —— 窗体代码里没写过 UserForm1，UserForm1 照样是合法
        #     引用；刚拖上去的 Label1 一行代码都还没有，它的名字照样存在于设计器
        #     里（v60）；
        #   * VBA 语言自带的名字（内建函数 / 内建常量 / 内建数据类型 v61，语言
        #     关键字 v62）—— 它们由语言运行时提供，用户代码里没调用过 MsgBox
        #     也照样能写 MsgBox、新模块第一行没写过 If 也照样能补 If。
        # 引擎靠它避免把这类名字当成"回声幻影"剔掉（见 engine._name_really_exists）。
        self._structural_names = set()
        # 真正作为"语言自带词汇"收录的那些（小写，v61 内建函数/常量/类型 +
        # v62 语言关键字）。引擎对这批名字的模糊匹配收紧一档（纯分散命中要求
        # 输入至少 4 个字符，详见 engine.trigger 里的说明）。
        self._builtin_names = set()
        # 宿主类型库的枚举常量（v67）：{小写名: 最短输入长度}。
        # 只收"有家族前缀"的库（Excel -> xl*、Office -> mso*），引擎对这批名字
        # 只认"从头开始的连续前缀"命中，且输入长度不低于家族前缀长度 ——
        # 它们数量太大（实测 4538 条），放开跳步匹配会淹没正常候选。
        self._host_enum_names = {}
        # ★v118：见 MOD_CACHE_ENABLED 的说明。
        # ★v124：键是 `(工程身份, 模块名小写)` —— 工程身份见 `_project_key`。
        #   只按模块名做键时，两个 .xlsm 各自的 `Module1` 会互相命中。
        #   _mod_cache: (pkey, 模块名小写) -> {"count", "ts", "recs",
        #               "proc_names", "proc_kinds", "type_names", "declared",
        #               "decl_by_mod"}
        #   _cm_cache : (pkey, 模块名小写) -> 复用中的 CodeModule 包装对象
        self._mod_cache = {}
        self._cm_cache = {}

    def release(self):
        """Excel/VBE 关闭或离开 VBE 时调用：清空标识符缓存并释放 COM 资源。

        这样切到其它程序（如 PyCharm）或关掉 Excel 后，占用的内存会及时回收，
        也不会有悬空的 COM 引用拖住 Excel 进程退出。
        """
        self._cache = None
        self._cache_time = 0
        self._caret_word = ""
        self._declared_by_module = {}
        self._structural_names = set()
        self._builtin_names = set()
        self._host_enum_names = {}
        # ★v118：缓存里握着 CodeModule 包装对象 —— 必须一起丢掉，
        # 否则 Excel 关闭时这些悬空引用会把进程拖住（release 的本意就在这儿）。
        self._mod_cache = {}
        self._cm_cache = {}
        # ★v118：组件列表缓存同样握着 COM 包装 —— 一并作废（模块级全局）。
        _comp_list_cache["comps"] = None
        _comp_list_cache["count"] = -1
        _comp_list_cache["key"] = None
        _comp_list_cache["edge"] = None
        _comp_list_cache["ts"] = 0.0
        # 丢掉缓存的 VBE 对象 —— 这是 Excel 能否真正退出的关键一步。
        # 刻意【不】调 CoFreeUnusedLibraries —— 它释放不了我们持有的引用，
        # 在 Excel 关闭期间调用反而会拖慢/惊扰 COM（详见 _co_free 的说明）。
        # 这里用 collect=True 主动 gc，确保 COM 包装对象立刻回收。
        _release_vbe_proxy(collect=True)

    def snapshot(self):
        """轻量读取「模块名 + 当前行号 + 行文本 + 光标列」，用于轮询检测内容变化。

        刻意不做列语义探测（那会临时移动光标），因此足够廉价，可高频调用。
        返回 (line_no, line_text, module_name, caret_col)；不在代码窗/多行选择/
        取不到时返回 None。

        v64：多带一个【模块名】。主线程靠它区分"用户真在敲字"和"换了个代码
        窗格/模块"—— 后者会让"同一行 + 文本不同"看起来像一次编辑（Ctrl+Tab
        切代码窗、点工程树换模块都会），其实一个字都没敲，不该弹候选。
        模块名取不到就带空串，调用方据此不做否决（绝不让工具整个哑掉）。

        v79：再带一个【光标列 ec】。只给"回车自动缩进"那道判据用（见
        enter_indent_wanted）—— 它必须在【键盘钩子线程】里就决定要不要接管
        这一下回车，而钩子线程不许碰 COM，只能读轮询留下的这份快照。
        ⚠️ 变更检测仍然只比前三个元素（行号/行文本/模块名）：光标的移动不算
        "内容变化"，否则点来点去都会弹候选窗（见 main._poll_mod_switch 附近）。
        """
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return None
            cp = vbe.ActiveCodePane
            if cp is None:
                return None
            sl, _sc, el, ec = cp.GetSelection()
            if sl != el:      # 多行选择：不参与补全
                return None
            cm = cp.CodeModule
            try:
                mod = str(cm.Name or "")
            except Exception:
                mod = ""
            try:
                ec = int(ec)
            except Exception:
                ec = 0
            snap = (sl, cm.Lines(sl, 1), mod, ec)
            _com_ok()
            return snap
        except Exception:
            _release_vbe_proxy()
            _com_fail()
            return None

    # ---- 标识符缓存 ----
    def invalidate_identifiers(self, modules=False):
        """文本内容变化后调用：丢弃标识符缓存，下次取用时会重新解析。

        避免「刚声明的变量不提示」——缓存里可能还是若干秒前解析的旧内容，
        新敲出来的名字还不在里面。只清 _cache、不动 _cache_time，
        以便调用方用 _cache_time 做重解析的限流。

        ★v118：modules=True 时连【模块级解析缓存】一起扔掉。给"换模块"用 ——
        用户可能在 A 模块里改过声明（行数没变，比如把 `Foo` 改成 `Fob`），
        然后切到 B 模块里调用它；那一改不会让 A 的行数发生变化，模块缓存
        察觉不到。切模块时整体作废一次最省心（这一次重建 ~90ms，而换模块
        本来就不是打字节奏里的动作）。
        """
        self._cache = None
        self._caret_word = ""
        if modules:
            self._mod_cache = {}

    def get_identifiers(self):
        now = time.time()
        if self._cache is not None and now - self._cache_time < _CACHE_TTL:
            return self._cache
        ids = self._collect_identifiers()
        self._cache = ids
        self._cache_time = now
        # ★v98：成员表跟着标识符池一起作废（同一个 TTL、同一个刷新点）——
        # 留着上一轮的结果会在用户新写了一个 Public 成员之后还看不到它。
        self._members_of_cache = {}
        return ids

    def get_proc_names(self):
        """返回工程里所有【过程名】（小写集合）—— Sub / Function / Property /
        Declare。★v98：用来判断"这次补全该不该自动补括号"。

        与 get_declared_names 的分工：后者是"Dim / Const / Sub / Function /
        Type / Enum 加组件名"的合集，**分不出变量与函数**。补括号只对过程该做 ——
        `Dim n As Long` 补成 `Long()` 会被 VBE 报编译错误。

        ⚠️ 直接交回内部集合（**只读，调用方不得修改**），理由同
        get_structural_names：这一路是【按一次补全】查的，不必为它复制一份。

        ⚠️ 集合里是**小写**名（与 get_proc_names / get_builtin_names 同口径）。
        代价是候选窗在 `c.` 位置可能显示成小写 —— 但那一处 VBE 自己弹得出
        成员列表（带原大小写），我们的名单只是"VBE 还没弹出来"时的兜底，
        优先级远低于保真。
        """
        if not getattr(self, "_proc_names", None):
            # 空就现收一次（顺带把标识符池建起来）—— 只判 None/空会把
            # "还没收集过"当成"工程里一个过程都没有"。
            try:
                self.get_identifiers()
            except Exception:
                return frozenset()
        return frozenset(getattr(self, "_proc_names", None) or ())

    def get_proc_kinds(self):
        """返回 {小写过程名: "sub" / "function" / "property"}。★v103c。

        与 `get_proc_names` **同源**（同一次 `extract_records` +
        `extract_proc_kinds`、同一个 TTL、同一个刷新点）："这个是不是过程"与
        "它是哪一种过程"永远出自同一份代码文本，不会一边新一边旧。

        用途：补全按 Tab 确认时，Sub 补【空格】、Function 补【括号】——
        见 engine._completion_followup。

        ⚠️ 交回内部 dict 本身（**只读，调用方不得修改**），理由同
        get_proc_names：这一路是按一次补全查的。
        ⚠️ 查不到（工程里没这个过程 / 后端只是没收集到）时调用方按"不知道"
        处理 —— 退回 v98 以来的老口径，绝不猜。
        """
        if not getattr(self, "_proc_kinds", None):
            try:
                self.get_identifiers()
            except Exception:
                return {}
        return getattr(self, "_proc_kinds", None) or {}

    def get_declared_names(self):
        """返回工程里真实声明过的名字（小写集合），供引擎剔除"提示自己"的幻影。

        只有被 Dim/Const/Sub/Function/Type/Enum 及组件名声明过的才算；隐式变量
        （用到即存在）不算。缓存随 get_identifiers 一并刷新。
        """
        if getattr(self, "_declared_names", None) is None:
            self.get_identifiers()
        return getattr(self, "_declared_names", set())

    def names_outside_caret(self, caret, scope_only=False):
        """把光标处的词抹掉后，【当前模块】里还出现过的名字（小写集合）。

        `name_exists_outside_caret` 的批量版：一次读取模块文本，供引擎在剔除
        "回声候选"时对多个候选复用 —— 否则每个候选都要打一次 COM。

        刻意只读【当前模块】：回声只可能产生于正在编辑的这个模块；全工程扫描
        既没必要，又会在多模块大工程上明显变慢。

        scope_only=True（v51）：只收集【从光标处看得见】的出现 —— 模块级行
        （含过程/Type 声明头行、模块声明区）与光标所属过程的行；*别的过程体内
        的出现一律不算*。

        为什么需要这个窄窗口：用户正在【过程内声明一个名字】（形参 / 局部 Dim）
        时，这个名字是过程局部的。若现场证据仍是"全模块出现过"，那么"别的过程
        里的同名局部变量"就会被当成"这个名字真实存在"的铁证，回声防护随即放行，
        于是用户把别的过程的变量名原样敲/回退/粘贴到形参位置时，弹窗提示出自己
        ——正是用户报的"只有一模一样才会提示自己"。收窄到可见窗口后，那种出现
        不再是证据，候选照常按回声剔除。
        默认 False = 全模块扫描（模块级声明的场景仍走这一路：模块级名字本就
        全模块可见，任何出现都是它的合法引用，见 v37 语义）。
        """
        # 读不到就返回 None（= 现场证据不可用），与"真的没有名字"（空集）
        # 区分开 —— 调用方据此决定是保守放过还是照常判定。
        try:
            line_no, col = int(caret[0]), int(caret[1])
        except Exception:
            return None
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return None
            cp = vbe.ActiveCodePane
            if cp is None:
                return None
            cm = cp.CodeModule
            if cm.CountOfLines <= 0:
                return set()
            code = cm.Lines(1, cm.CountOfLines)
        except Exception:
            return None
        try:
            # blank_decl=True：这里**必须**把光标处的词抹掉 —— 哪怕它是这一行正在
            # 声明的名字（v79e）。本函数的语义是"抹掉光标处的词之后，模块里还剩
            # 哪些名字"，那个词自己不能当证据；而 _blank_ident_at 默认会放过
            # "正在声明的名字"（为了不改动文本、不破坏行结构），所以在这一路要
            # 显式要求抹掉。不加这个参数会怎样：光标停在 `Sub zzq` 上（zzq 别处
            # 都没有）时，zzq 把自己算成"别处出现过"的铁证，回声防护放行，
            # 弹窗把用户正在敲的字原样提示回来 —— 第 21.4 节钉的正是这条。
            src = vba_parser._blank_ident_at(code, line_no, col, True)
            src = vba_parser._mask_strings_and_comments(src)
        except Exception:
            src = code
        try:
            if not scope_only:
                return set(m.group(0).lower()
                           for m in re.finditer(r"[^\W\d]\w*", src))
            rows = src.split("\n")
            if not (1 <= line_no <= len(rows)):
                # 行号对不上（理论上不该发生）：保守退回全模块扫描，
                # 宁可少剔除候选，也不要凭错的窗口把真名字误杀。
                return set(m.group(0).lower()
                           for m in re.finditer(r"[^\W\d]\w*", src))
            # 逐行标注归属：声明头行归属【它自己声明的那个过程】（与
            # parser._line_proc_map 一致），但它在下面按模块级收集 —— 过程的
            # 声明本身就是模块级语句，`Public Sub update()` 这种模块级过程名
            # 必须在这里找得到，否则在别的过程里就再也搜不到它了。
            owners = []
            cur = None
            for raw in rows:
                s = raw.strip()
                if vba_parser._RE_ANY_PROC_END.match(s):
                    cur = None
                m_head = (vba_parser._RE_SUB.match(s)
                          or vba_parser._RE_PROP.match(s))
                if m_head:
                    cur = m_head.group(1)
                owners.append((cur, bool(m_head)))
            keep = owners[line_no - 1][0] or ""
            if not keep:
                # 光标不在任何过程里（模块声明区）：没有可收窄的窗口。
                return set(m.group(0).lower()
                           for m in re.finditer(r"[^\W\d]\w*", src))
            kl = str(keep).lower()
            names = set()
            for (owner, is_header), raw in zip(owners, rows):
                if owner is None or is_header or str(owner).lower() == kl:
                    names.update(m.group(0).lower()
                                 for m in re.finditer(r"[^\W\d]\w*", raw))
            return names
        except Exception:
            return None

    def name_exists_outside_caret(self, name, caret):
        """把光标处的词抹掉后，【当前模块】里 name 是否还在别处出现。

        这是判断"正在输入的词是不是自己的回声"的最可靠依据：

        * 打 Name到一半/打全时，光标处那几个字符本身会被当成"用到过的隐式
          变量"收录进候选池。若这个名字在工程里再没别处出现，把它提示出来
          就是纯粹的自我提示（回退出来的 numA 就是典型）。
        * 反过来，只要它在别处真实存在（哪怕解析层没能识别出它的声明），
          那它就是个合法候选 —— 用户把它打全时**必须继续提示**，否则会出现
          "输入 a / ar 都提示 arr，一把 arr 打全列表就消失"。

        刻意只在【当前模块】里查：回声只可能产生于正在编辑的这个模块；
        全工程扫描既没必要，又会在多模块大工程上明显变慢。

        刻意只读文本做一次线性扫描（不触发标识符重解析），因此很廉价，
        可以每次按键都调用。
        """
        if not name:
            return False
        return str(name).lower() in (self.names_outside_caret(caret)
                                      or ())

    def caret_word_at_collect(self):
        """返回"收集标识符那一刻光标处的词"（小写，可能为空串）。

        标记符池是带缓存的：调用方拿到的候选可能来自 0.4 秒前的解析。那时
        光标处的词，就是用户此刻可能正在回退删除的那个词 —— 引擎用它把
        "回退途中残留的更长的旧片段"从候选里剔掉（v43）。
        """
        return getattr(self, "_caret_word", "") or ""

    def declared_elsewhere(self, name, module_name):
        """name 是否在【其它模块】里被真实声明过。

        True  = 别的模块声明过（跨模块 Public 证据）；
        False = 只在当前模块声明过，或压根没声明过 -> 请以现场文本为准；
        None  = 后端不提供这项信息 -> 调用方退回旧证据（旧式 / 测试后端）。

        分工：跨模块的 Public 名字只有声明集合认得它（现场扫不到），靠 True；
        而只在【正在编辑的这个模块】里声明过的名字，一律以现场文本为准 ——
        用户完全可能刚把它删掉，声明集合却还是 0.4 秒前解析的快照，信它就会把
        "回退途中残留的更长的旧片段"保下来（v43 修的就是这个）。
        """
        by_mod = getattr(self, "_declared_by_module", None)
        if not isinstance(by_mod, dict):
            return None
        m = str(module_name or "").lower()
        for owner in (by_mod.get(str(name).lower()) or ()):
            if owner != m:
                return True
        return False

    def _collect_identifiers(self):
        # 用缓存的 VBE（会顺带做失败退避），避免每次收集都新建
        # Excel.Application 代理 —— 那会让 Excel 退出时要清理的引用越堆越多。
        vbe = _get_vbe_cached()
        if vbe is None:
            return []
        modules = []
        records = []
        # 光标位置（用于排除"正在输入的词"被当成隐式变量收录）
        # 模块名比对失败就等于没拿到光标，会出现"打什么就提示什么"的自我提示，
        # 因此取不到模块名时宁可对所有模块应用（误伤极小），也不要丢弃 caret。
        caret_mod, caret = _caret_position(vbe)
        _log("collect: caret_mod=%r caret=%r" % (caret_mod, caret))
        # ★v124：模块级缓存（`_cm_cache` / `_mod_cache`）的键必须带【工程身份】。
        # 原来只用模块名小写做键 —— 同时开两个 .xlsm（各有一个 `Module1`）时，
        # A 的 `Module1` 会命中 B 的缓存：补全出另一个工作簿里的变量名，
        # 而且因为"行数恰好相同"还会一直命中（TTL 到期才自愈）。
        #
        # 跨工程模式（ENABLE_CROSS_PROJECT）下一次遍历会混进多个工程的同名
        # 模块，一个 `_pkey` 根本描述不了 —— 那种模式干脆不吃这两份缓存：
        # 慢一点，但绝不会串（该模式默认关闭）。
        _pkey = None if ENABLE_CROSS_PROJECT else _project_key(
            _active_vb_project(vbe))
        # 注意：早先这里会把"声明行正在声明的名字"从【全部候选】里按名字整体剔除
        # （连前缀一起），用于杜绝"定义变量时提示变量自己"。但这会把合法的前缀
        # 补全也误杀——在 numArr 声明行输入 num，numArr 就再也提示不出来（bug：
        # 输入 num 只能提醒 num、不提醒 numArr）。正确的语义是"精确匹配自己才隐藏、
        # 前缀照常提示"，这由引擎层 in_decl_position + ctx["decl_names"] 的精确自我
        # 剔除负责（输入完整 numArr 才隐藏，输入前缀 num 仍正常提示 numArr）。
        # 因此解析/收集层不再做任何声明行全局剔除，名字一律收录，交给引擎按"正在
        # 输入的词"与"当前声明行的名字"决定是否剔除。
        # "类型名"集合：`Dim f As <这里>` 用得上——窗体名、类模块名、标准模块名、
        # 以及代码里定义的 Type / Enum 名。
        # 每次收集都重置：没有光标信息时它就是空串（不误伤任何候选）
        self._caret_word = ""
        _decl_by_mod = {}
        type_names = set()
        # 真实"声明"过的名字（Dim/Const/Sub/Function/Type/Enum 及组件名）。
        # 供引擎层区分"真名字"与"幻影"：回退删字回退出来的未定义词（如 numA）
        # 只可能是隐式残留 / 正在敲的词本身，引擎据此剔除"提示自己"。
        declared_names = set()
        # ★v98：过程名（Sub / Function / Property / Declare），随 extract_records
        # 一并返回 —— 与 declared_names 的差别见 __init__ 里的说明。
        proc_names = set()
        # ★v103c：过程名(小写) -> "sub" / "function" / "property"。
        # 补全时用它决定"后面跟空格还是跟括号"（见 engine._completion_followup）。
        proc_kinds = {}
        # 结构性名字：组件名 + 窗体控件名（v60）+ 语言自带名字（v61/v62）。
        # 详见 __init__ 里的说明。
        structural_names = set()
        # 真正作为"VBA 内建名字"收录的那些（v61，小写）。与清单的差别：工程里
        # 已经自己声明过同名时以用户的为准，那些名字【不】算内建 —— 于是引擎
        # 对它们收紧匹配时不会连用户自己的定义一起收紧。
        builtin_names = set()
        # 宿主类型库枚举常量 -> 最短输入长度（v67），见下面的收集段。
        host_min = {}
        try:
            for comp in _collect_components(vbe, caret_mod):
                try:
                    try:
                        mod_name = str(comp.Name)
                    except Exception:
                        continue
                    # 组件名本身（模块名 / 窗体名 / 类名 / ThisWorkbook、Sheet1 等
                    # 文档模块名）也是工程内可以直接引用的标识符：
                    #   Module1.Foo / UserForm1.Show / Dim c As New Class1
                    # 记为模块级、非私有 -> 本工程任意位置都能提示。
                    #
                    # 必须在读 CodeModule【之前】收录：新建的窗体 / 类模块通常
                    # 一行代码都没有（CountOfLines == 0）。放在读代码之后会被
                    # `if count <= 0: continue` 整个跳过 —— 症状正是"模块名能
                    # 提示、窗体名死活不提示"。工程被密码保护、CodeModule 取不到
                    # 时同理，名字照样该提示。
                    if _RE_PLAIN_IDENT.match(mod_name):
                        records.append((mod_name, mod_name, None, False))
                        type_names.add(mod_name)
                        declared_names.add(mod_name.lower())
                        structural_names.add(mod_name.lower())
                        _decl_by_mod.setdefault(
                            mod_name.lower(), set()).add(mod_name.lower())
                    modules.append(mod_name)
                    # 窗体控件名（v60）：只收集【光标所在的那个窗体】的。
                    #   * 控件名只在它自己的窗体代码模块里能裸名引用（跨模块
                    #     得写 UserForm1.Label1），所以别的窗体的控件名收进来
                    #     也永远不会可见，白白占池子；
                    #   * 取一次控件要过 Designer 接口（真机实测约 23ms/窗体），
                    #     只对正在编辑的窗体取，开销小到可以忽略。
                    # 必须在读 CodeModule【之前】收集：新建窗体一行代码都没有，
                    # 落在 `count <= 0: continue` 之后就永远收不到了。
                    if (_is_msform(comp.Type) and caret_mod
                            and str(caret_mod).lower() == str(mod_name).lower()):
                        for _cn in _form_control_names(comp):
                            if not _RE_PLAIN_IDENT.match(_cn):
                                continue
                            records.append((_cn, mod_name, None, True))
                            declared_names.add(_cn.lower())
                            structural_names.add(_cn.lower())
                            _decl_by_mod.setdefault(
                                _cn.lower(), set()).add(str(mod_name).lower())
                    # ★v118：CodeModule 按模块名复用 —— `comp.CodeModule` 每次都
                    # 返回一个新的 COM 包装对象，pywin32 要为它重新 Bind 一遍类型
                    # 信息（真机实测 24 个模块合计约 19ms/轮）。CodeModule 本身是
                    # 稳定引用，按模块名缓存即可；万一模块被删了重建，旧包装会在
                    # 读数时抛异常 —— 那就丢掉缓存重新拿一次。
                    _mk = str(mod_name).lower()
                    # `_mk` 保持"模块名小写"—— 它在 `_decl_by_mod` 里当模块名用
                    # （引擎按模块名比对）。缓存键另起一个 `_ck`。
                    _ck = (_pkey, _mk) if _pkey is not None else None
                    cm = self._cm_cache.get(_ck) if _ck is not None else None
                    count = None
                    if cm is not None:
                        try:
                            count = int(cm.CountOfLines)
                        except Exception:
                            cm = None
                            count = None
                    if cm is None:
                        cm = comp.CodeModule
                        count = int(cm.CountOfLines)
                        if _ck is not None:
                            self._cm_cache[_ck] = cm
                    if count <= 0:
                        continue
                    # 模块名对不上就等于没拿到光标，会出现"打什么就提示什么"
                    # 的自我提示；取不到模块名时宁可对所有模块应用（误伤极小），
                    # 也不要丢弃 caret。
                    apply_caret = caret if (
                        not caret_mod
                        or str(caret_mod).lower() == _mk
                    ) else None
                    # ★v118：非当前模块且行数没变 -> 直接重放上次的解析产物。
                    # 为什么可以这么做、TTL 兜的是什么，见 MOD_CACHE_ENABLED 的
                    # 说明。重放的是"该模块对四路承接的增量"，与原来逐字等价：
                    # records 的顺序（recs 在前、imp 在后）、proc_kinds 的覆盖
                    # 顺序都按同一个模块遍历顺序重演。
                    _ent = (self._mod_cache.get(_ck)
                            if (MOD_CACHE_ENABLED and _ck is not None
                                and apply_caret is None)
                            else None)
                    if (_ent is not None and _ent["count"] == count
                            and (time.time() - _ent["ts"]) < _MOD_CACHE_TTL):
                        type_names |= _ent["type_names"]
                        records.extend(_ent["recs"])
                        proc_names |= _ent["proc_names"]
                        proc_kinds.update(_ent["proc_kinds"])
                        declared_names |= _ent["declared"]
                        for _dn, _dms in _ent["decl_by_mod"].items():
                            _decl_by_mod.setdefault(_dn, set()).update(_dms)
                        _log("  mod=%-16s 复用模块缓存（%d 行）"
                             % (mod_name, count))
                        continue
                    code = cm.Lines(1, count)
                    is_std = _is_std_module(comp.Type)
                    # 代码里自定义的 Type / Enum 名同样是类型名
                    # ★v118：顺手收成局部集合 —— 缓存要存"本模块的增量"。
                    _tns = set(vba_parser.extract_type_enum_names(code))
                    type_names |= _tns
                    recs, _proc_of = vba_parser.extract_records(
                        code, module=mod_name, is_std_module=is_std,
                        caret=apply_caret)
                    proc_names |= _proc_of
                    # ★v103c：顺手把"每个过程是 Sub / Function / Property"记下来
                    # （同一行、同一批正则，不额外打 COM）。补全时需要它来定
                    # "后面跟空格还是跟括号"，见 engine._completion_followup。
                    try:
                        _pks = dict(vba_parser.extract_proc_kinds(code))
                    except Exception:
                        _pks = {}
                    proc_kinds.update(_pks)
                    records.extend(recs)
                    # 记下【这一刻光标处的词】：它多半是用户正在输入 / 正在
                    # 回退删除的词。回退删字时，缓存的池子里还留着它更长的
                    # 旧版本，引擎靠这个词把那个幽灵剔掉（v43）。
                    # 只在【确知光标在哪个模块】时才记：caret_mod 为空时
                    # 同一个 caret 会被套用到所有模块，这里最后赋值的那个
                    # 词可能来自任意模块，拿它当"光标旧词"就会误伤候选。
                    if apply_caret and caret_mod:
                        try:
                            self._caret_word = vba_parser.ident_at_caret(
                                code, apply_caret[0],
                                apply_caret[1]).lower()
                        except Exception:
                            self._caret_word = ""
                    # "真名集合"必须剔除【光标所在声明行正在声明的那些名字】：
                    # 那是用户正在敲 / 正在删的词，不是"工程里已存在的真名字"。
                    #
                    # 不剔除会怎样：回退删字时（Dim numArr 退成 Dim numA），
                    # 文本里此刻确实写着 Dim numA，解析层会把它当成一个真实的
                    # 声明收进真名集合；引擎"不是真名就剔除自身"的防护随即失效，
                    # numA 又被当作一条补全建议提示出来 —— 正是用户反馈的
                    # "回退快的时候还是提示 numA"。
                    caret_decl = set()
                    if apply_caret:
                        try:
                            caret_decl = set(
                                str(n).lower() for n in
                                (vba_parser.decl_names_at_caret(
                                    code, apply_caret) or ()))
                        except Exception:
                            caret_decl = set()
                    # ★v118：同样收成局部集合 —— 缓存存"本模块的增量"。
                    _decl_add = set(
                        str(r[0]).lower() for r in recs
                        if str(r[0]).lower() not in caret_decl)
                    declared_names |= _decl_add
                    # ★v87：`_decl_by_mod` 只收【别的模块能裸名引用】的声明 ——
                    # 它的唯一消费者是 declared_elsewhere（引擎问"这个名字在
                    # 【别的模块】声明过吗"，用来证明一个回声候选是真名字）。
                    # 口径必须与候选池的可见性（engine.filter_identifiers_by_scope）
                    # 一致，否则会把不可见的声明当成"真名字"的证据：
                    #   * priv=True —— 模块私有。**类模块 / 窗体 / 文档模块的模块级
                    #     成员默认就是 Private**（跨模块只能经实例或限定名访问，见
                    #     parser._is_private、_is_std_module），别的模块压根写不成裸名；
                    #   * proc 非空 —— 别的过程的局部变量 / 形参，同理。
                    # 不排除会怎样（用户报的"在类模块里提示自己"）：同一个成员名
                    # 在两个类模块里各声明一次（比如 类2 的 `Public 订单金额` 与
                    # SummaryClass 的 `订单金额`），在 类2 那一行打 / 退格这个词时，
                    # declared_elsewhere 因【另一个类模块的那份私有声明】返回 True
                    # ⇒ 回声防护放行 ⇒ 弹窗把用户正在敲的字原样提示回来。
                    # ★v118：先收到局部表里，最后统一并进 _decl_by_mod ——
                    # 缓存存的就是这张"本模块的增量表"。
                    _dbm_add = {}
                    for _r in recs:
                        _n = str(_r[0]).lower()
                        if _n in caret_decl:
                            continue
                        _r_priv = bool(_r[3]) if len(_r) > 3 else False
                        _r_proc = _r[2] if len(_r) > 2 else None
                        if _r_priv or _r_proc:
                            continue
                        _dbm_add.setdefault(_n, set()).add(_mk)
                    for _dn, _dms in _dbm_add.items():
                        _decl_by_mod.setdefault(_dn, set()).update(_dms)
                    # 隐式变量（没写 Option Explicit 时"用到即存在"）：
                    # 只补 extract_records 没声明过的名字，避免重复与作用域冲突。
                    #
                    # v75：写了 Option Explicit 的模块整段跳过（详见
                    # IMPLICIT_HONOR_OPTION_EXPLICIT 的说明）—— "用过但没声明"
                    # 在那里是编译错误，收进来只会提示出坏代码。
                    imp = []
                    if ENABLE_IMPLICIT_IDENTIFIERS:
                        if implicit_gate_applies(code):
                            _log("  mod=%-16s Option Explicit -> 不收集隐式变量"
                                 % mod_name)
                        else:
                            imp = vba_parser.extract_implicit_records(
                                code, module=mod_name, is_std_module=is_std,
                                declared=recs,
                                scope=IMPLICIT_SCOPE,
                                caret=apply_caret)
                            records.extend(imp)
                    _log("  mod=%-16s type=%-3s std=%-5s recs=%-4d imp=%-4d caret_applied=%s"
                         % (mod_name, comp.Type, is_std, len(recs), len(imp),
                            bool(apply_caret)))
                    # ★v118：把"本模块对四路承接的增量"整套存下来，下一轮行数
                    # 没变就直接重放（见上面命中分支）。只在【光标不在这块】时
                    # 存 —— 带 caret 的解析结果与光标位置绑定，存了会串。
                    if (MOD_CACHE_ENABLED and _ck is not None
                            and apply_caret is None):
                        self._mod_cache[_ck] = {
                            "count": count,
                            "ts": time.time(),
                            "recs": list(recs) + list(imp),
                            "proc_names": frozenset(_proc_of or ()),
                            "proc_kinds": dict(_pks),
                            "type_names": frozenset(_tns),
                            "declared": frozenset(_decl_add),
                            "decl_by_mod": {k: frozenset(v)
                                            for k, v in _dbm_add.items()},
                        }
                except Exception:
                    continue
        except Exception:
            return []

        # ---- VBA 语言自带的名字（v61）----
        # 用户报："VBA 本身自带的函数也加入提示列表，现在不能提示"。
        #
        # 这些名字由【语言运行时】决定存在，与工程代码文本无关，所以三件事都要做：
        #   1) 进候选池 —— 否则模糊匹配无从命中；
        #   2) 进 structural_names —— 引擎判"回声候选是不是真名字"时最先看它。
        #      打 ms 想补 MsgBox 时，ms 是 msgbox 的连续子串 -> 命中回声候选；
        #      而代码里若从没出现过 MsgBox，现场文本扫不到、工程里也没声明过，
        #      少了这一路证据候选必被踢掉 —— 症状正是"打 ms 死活不提示 MsgBox"；
        #   3) 进 _decl_by_mod（挂到 VBA 这个"别的模块"下），让 declared_elsewhere
        #      也认它们 —— 现场证据不可用（COM 读不到）时的兜底路径同样放行。
        #
        # 已声明过的同名不重复收录：工程里真有 Function Left(...) 时以用户的定义
        # 为准，免得列表里出现两条同名。
        #
        # 【v70】这一批里 **枚举常量（vb*）默认不收** —— 用户口径："我不怎么使用
        # VBA 内置枚举，提示出来对我有干扰；函数那些要保留。" 于是函数（MsgBox /
        # Left / Split …）与内建数据类型照常收，只有 vb* 那一组由
        # ENABLE_VBA_CONSTANTS 裁掉（默认 False；`set VBECOMPLETE_VBA_CONSTANTS=1`
        # 开回来）。注意裁剪放在【收集侧】而不是改清单本身：清单保持完整，
        # 测试与诊断仍能看到全貌，开关一开立刻生效。
        #
        # 【v77】曾把**内建函数与内建数据类型**也默认关掉（用户嫌提示太多）。
        # 【★v96】又收回来了 —— 用户原话："我想把 vba 关键字、函数名这两部分，
        # 也加入到提醒列表里。枚举类型就不要加了，枚举加上，就会太多了。"
        # ⇒ ENABLE_VBA_BUILTINS 默认 True，**枚举仍然默认 False**（vb* 那批归
        # ENABLE_VBA_CONSTANTS、宿主 xl*/mso* 归 ENABLE_HOST_ENUMS，都没动）。
        #
        # ⚠️ 引擎侧对这一批有一道收紧（2~3 字符的跳步命中要求首字母对齐，
        # 见 engine.BUILTIN_SCATTER_MIN）—— 它一直是活的，只是 v77 之后
        # `_builtin_names()` 默认空集、整段不生效。v96 把这一批放回池子，
        # 那道收紧也就自动重新起作用了（这正是当初为它们准备的）。
        # ★v123：读【生效快照】`ACTIVE_*`（= 词库里的那一份），不再是代码里的
        # 常量清单。属性访问发生在每次收集时，所以 `vba_builtins.load_active()`
        # 之后立刻生效，不需要重启进程内任何东西。
        #
        # 语言自带的那批（供引擎收紧 2~3 字符的跳步命中，见 BUILTIN_SCATTER_MIN）。
        # ★只为 is_builtin=1 的词记名：用户自己加进词库的词不该被"公共词汇"那套
        # 收紧连坐 —— 他是特意加的，当然希望打得出来。集合取一次放局部，两个分支
        # 共用（且必须提到分支之外：内建函数关掉、只开关键字时也要能用）。
        _bn_builtin = vba_builtins.ACTIVE_BUILTIN_LOWER
        if ENABLE_VBA_BUILTINS:
            _bn_src = vba_builtins.ACTIVE_FUNCTIONS
            if ENABLE_VBA_CONSTANTS:
                _bn_src = _bn_src + vba_builtins.ACTIVE_CONSTANTS
            for _bn in _bn_src:
                _bl = _bn.lower()
                if _bl in declared_names:
                    continue
                records.append((_bn, _BUILTIN_MODULE, None, False))
                declared_names.add(_bl)
                structural_names.add(_bl)
                if _bl in _bn_builtin:
                    builtin_names.add(_bl)
                _decl_by_mod.setdefault(_bl, set()).add(_BUILTIN_MODULE)
            # 内建数据类型名：既进候选池（`As |` 位置的候选是从可见标识符里筛出来
            # 的，不在池子里就永远显示不出来），也进 type_names（那一位置只该冒
            # 类型名）。与函数清单重合的（Date / String 等）由上面的 declared_names
            # 去重挡掉，不会再收一遍。
            for _tn in vba_builtins.ACTIVE_TYPE_NAMES:
                _tl = _tn.lower()
                type_names.add(_tn)
                if _tl in declared_names:
                    continue
                records.append((_tn, _BUILTIN_MODULE, None, False))
                declared_names.add(_tl)
                structural_names.add(_tl)
                if _tl in _bn_builtin:
                    builtin_names.add(_tl)
                _decl_by_mod.setdefault(_tl, set()).add(_BUILTIN_MODULE)

        # ---- 语言关键字 / 保留字（v62）----
        # 用户要求："把 vba 里面的所有关键字都纳入到提示词里"。
        #
        # 走的是与内建名字**完全相同**的四路承接，原因也一样：关键字同样由语言
        # 本身决定存在，与代码文本无关。比如输入 if 想补 If —— 而此刻代码里恰好
        # 一个 If 都还没写过（新模块的第一行），现场文本扫不到、工程里也没声明过，
        # 少任何一路都会被回声防护当成幻影剔掉（v60 窗体名、v61 内建函数的同一个坑）。
        #
        # 唯一区别是开关独立（ENABLE_VBA_KEYWORDS）：关掉关键字不影响内建函数。
        # 关键字同样进 builtin_names —— 它们是语言自带的公共词汇（数量大、什么字母
        # 组合都凑得出子序列），要跟内建函数一样收紧模糊匹配，见 engine.trigger。
        #
        # 【v77】曾**默认关闭**（用户口径："vba 本身带的关键字也删掉"）。
        # 【★v96】又收回来了（用户原话："我想把 vba 关键字、函数名这两部分，
        # 也加入到提醒列表里。枚举类型就不要加了。"）⇒ ENABLE_VBA_KEYWORDS
        # 默认 True；想关回去 `set VBECOMPLETE_VBA_KEYWORDS=0`。
        #
        # ⚠️ 关键字这一批是"公共词汇"里最松的一批：Set / For / In / To / Is /
        #    Not / On 这些短词几乎能从任何两三个字母凑出子序列。引擎侧那道
        #    "2~3 字符的跳步命中要求首字母对齐"的收紧就是为它们准备的
        #    （见 engine.trigger 与 BUILTIN_SCATTER_MIN）—— 别把它关掉。
        if ENABLE_VBA_KEYWORDS:
            for _kn in vba_builtins.ACTIVE_KEYWORDS:
                _kl = _kn.lower()
                if _kl in declared_names:
                    continue
                records.append((_kn, _BUILTIN_MODULE, None, False))
                declared_names.add(_kl)
                structural_names.add(_kl)
                if _kl in _bn_builtin:
                    builtin_names.add(_kl)
                _decl_by_mod.setdefault(_kl, set()).add(_BUILTIN_MODULE)

        # ---- 宿主类型库的枚举常量（v67）----
        # 用户报："输入 vb 会提示一堆 VBA 枚举值，输入 xl 却一个 xl 开头的枚举值
        # 都不提示。" vb* 是 v61 收的 VBA 内建常量；xl* 属于【宿主 Excel 类型库】，
        # 一直没收 —— 于是同一个动作用在 xl 上什么都出不来。
        #
        # 与内建名字同源：由工程【引用的类型库】决定存在，代码文本里从没写过
        # xlUp 也照样该补得出来。所以四路承接完全一样（池子 / declared_names /
        # structural_names / _decl_by_mod）—— 少任何一路，回声防护都会把
        # "代码里从未出现过的 xlUp" 当幽灵剔掉（v60 窗体名、v61 内建函数、
        # v62 关键字都栽在这一件事上，这是第四次）。
        #
        # 唯一的差别在【匹配口径】：后端把 {小写名: 最短输入长度} 交给引擎，
        # 引擎只认"从头开始的连续前缀"且输入长度不低于该值。理由见 engine.trigger。
        if ENABLE_HOST_ENUMS:
            for _hname, _hmin in host_enum_constants():
                _hl2 = _hname.lower()
                if _hl2 in declared_names:
                    continue
                records.append((_hname, _HOST_ENUM_MODULE, None, False))
                declared_names.add(_hl2)
                structural_names.add(_hl2)
                _decl_by_mod.setdefault(_hl2, set()).add(_HOST_ENUM_MODULE)
                host_min[_hl2] = _hmin

        self._type_names = type_names
        self._declared_names = declared_names
        self._proc_names = proc_names
        self._proc_kinds = proc_kinds          # ★v103c：过程 -> sub/function/property
        self._declared_by_module = _decl_by_mod
        self._structural_names = structural_names
        self._builtin_names = builtin_names
        self._host_enum_names = host_min
        # 只在数量变化时记日志，避免每 2 秒刷一行
        if len(records) != getattr(self, "_last_id_count", -1):
            self._last_id_count = len(records)
        return records

    def get_var_type(self, var_name):
        """★v98：查一个变量的**声明类型名**（小写）；查不到返回 None。

        只读【当前模块】的代码，用 `parser.var_types` 建表 —— 变量可能声明在
        当前模块的任何一个过程里，所以要读整个模块（`Lines(1, CountOfLines)`
        一次，与 `names_outside_caret` 同量级的一次读）。

        返回 None 的情形都交给上层走"不收窄"的兜底：类型没写明、隐式变量、
        宿主对象类型（`Range` 这类不在本工程里，自然也查不到成员）。
        """
        if not var_name:
            return None
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return None
            cp = vbe.ActiveCodePane
            if cp is None:
                return None
            cm = cp.CodeModule
            cnt = int(cm.CountOfLines)
            if cnt <= 0:
                return None
            vt = vba_parser.var_types(cm.Lines(1, cnt))
            return vt.get(str(var_name).lower())
        except Exception:
            return None

    def get_members_of(self, type_name):
        """★v98：返回某个类型（类模块 / 窗体 / 标准模块）**对外可见**的成员名。

        供引擎在 `对象.` 之后收窄候选：用户口径是"c.te 应该提示 cls 类的内部
        成员，而不是提示 temp"。可见性口径见 `parser.public_members`。

        type_name : 类型名（可带模块限定，如 `Sheet1.Parent`）。大小写不敏感。

        返回小写集合；**查不到就返回空集合** —— 调用方把"空"理解成
        "不知道这个类型的成员"，于是走"不收窄"的兜底（宁可多提示，也绝不
        把该提示的成员挡掉）。

        ⚠️ 只走**当前活动工程**，且逐个组件现读现算、结果缓存到收集周期结束 ——
        与 get_identifiers 同一个 TTL，不额外打 COM。
        """
        if not type_name:
            return frozenset()
        key = str(type_name).split(".")[-1].strip().lower()
        if not key:
            return frozenset()
        cache = getattr(self, "_members_of_cache", None)
        if cache is None:
            cache = self._members_of_cache = {}
        stamp = getattr(self, "_cache_time", 0)
        hit = cache.get(key)
        if hit is not None and hit[0] == stamp:
            return hit[1]
        out = frozenset()
        try:
            vbe = _get_vbe_cached()
            if vbe is not None:
                caret_mod, _caret = _caret_position(vbe)
                for comp in _collect_components(vbe, caret_mod):
                    try:
                        if str(comp.Name).lower() != key:
                            continue
                        cm = comp.CodeModule
                        cnt = int(cm.CountOfLines)
                        if cnt <= 0:
                            break
                        code = cm.Lines(1, cnt)
                        out = frozenset(
                            str(n).lower() for n in vba_parser.public_members(
                                code, is_std_module=_is_std_module(comp.Type)))
                        break
                    except Exception:
                        continue
        except Exception:
            out = frozenset()
        cache[key] = (stamp, out)
        return out

    def get_type_names(self):
        """返回可作为「数据类型名」使用的名字（窗体/类模块/标准模块名、Type/Enum 名）。

        供引擎在 `Dim x As |` 这类位置过滤候选：那里只该冒出类型名，
        不该冒出变量与过程名。
        """
        return getattr(self, "_type_names", set())

    def get_structural_names(self):
        """返回由【工程结构 / 语言运行时】决定其存在的名字（小写集合）。

        三类：组件名（模块 / 窗体 / 类 / 文档模块名）+ 窗体控件名 + VBA 语言自带
        的名字（内建函数 / 内建常量 / 内建数据类型 / 关键字）。
        ⚠️ v77：第三类默认是空的（那几个开关都默认关）⇒ 默认口径下只有
        **组件名 + 窗体控件名**。这一路证据本身一个字没删，开关打开就回来。

        与 get_declared_names 的区别：后者是"代码文本里真声明过什么"，前者是
        "本来就有、与代码文本无关"。用户刚拖上去的 Label1、一行代码都没有的新
        窗体 UserForm1、代码里从没调用过的 MsgBox、新模块第一行就想补的 If，
        都属于前者而不属于后者。

        引擎用它来判断"回声候选是不是真名字"——这类名字不能走"只在本模块声明
        过、现场文本又扫不到 -> 当幽灵剔掉"的规则（v60 组件名与控件名、v61 VBA
        内建名字：打 ms 要能补出从没写过的 MsgBox；v62 关键字：打 if 要能补出
        新模块里还没写过的 If）。

        ⚠️ 直接交回内部集合（**只读，调用方不得修改**）：v67 收进宿主枚举常量后
        这个集合有 4600 多条，而引擎的回声防护会对每个候选查一次 —— 每次调用都
        复制一份的话，输入 xl（一次带出 2000 多个候选）光复制就要 2 秒。
        """
        return getattr(self, "_structural_names", None) or set()

    def get_builtin_names(self):
        """返回真正作为"语言自带词汇"收录的那些（小写集合，v61 / v62）。

        含：VBA 内建函数 / 内建常量 / 内建数据类型（v61）+ 语言关键字（v62）。

        与 vba_builtins 的清单不同：工程里已经自己声明过同名时以用户的定义为准，
        那些名字不会出现在这里 —— 引擎据此对"语言自带的公共词汇"收紧模糊匹配
        （纯分散命中要求输入至少 4 个字符），而不会连用户自己的同名定义一起收紧。

        后端不提供该接口时引擎不做任何收紧（旧式 / 测试后端行为不变）。
        """
        return getattr(self, "_builtin_names", None) or set()

    def get_host_enum_names(self):
        """宿主类型库的枚举常量 -> {小写名: 最短输入长度}（v67）。

        内容：工程【引用的类型库】里的枚举常量，且只收有家族前缀的库
        （Excel 的 xl*、Office 的 mso*）。值是"该家族前缀的长度"，引擎据此
        要求输入至少这么长、且候选从头开始以它开头才提示 ——
        于是 xl -> xl*、xlu -> xlUp，而 ms 只出 MsgBox、绝不带出 mso*。

        为什么必须单独给一份而不是并进 get_builtin_names：内建那批（400 来个）
        是"收紧一档跳步匹配"，这批（4500 多条）得直接【只认前缀】—— 同一条
        规则套两批名字，轻则噪音爆炸（实测输入 ms 会带出 2365 条 mso*），
        重则把内建名字该有的跳步补全也一起砍掉。

        工程里已经自己声明过同名的以用户的定义为准，那些名字不会出现在这里。

        后端不提供该接口时引擎不做任何收紧（旧式 / 测试后端行为不变）。
        约定：键已经是小写；直接交回内部字典（**只读，调用方不得修改**）——
        它有 4500 多条，每次调用复制一份纯属浪费。
        """
        if not getattr(self, "_host_enum_names", None):
            # 空就现收一次（顺带把标识符池建起来）—— __init__ 里它初始化为 {}，
            # 光判 None 会让"还没收集过"被当成"没有宿主常量"。
            try:
                self.get_identifiers()
            except Exception:
                return {}
        return getattr(self, "_host_enum_names", None) or {}

    def vbe_popup_visible(self):
        """VBE 自带的提示窗（参数信息 / 列出成员）此刻是否可见（v63）。

        ⚠️ v65 起语义澄清：这是"**任何**提示窗"，含参数信息（形参签名），
        只用于诊断；引擎真正据以让位的是 vbe_yield_visible()。保留本方法是为了
        与旧式 / 测试后端和诊断日志兼容。纯 Win32 查询、不碰 COM。
        """
        return vbe_popup_visible()

    def vbe_yield_visible(self):
        """VBE 的【成员列表】窗此刻是否可见 —— 引擎据此让位（v65）。

        用户口径（v65 澄清 v63）："只有弹成员列表的时候才让位，弹形参签名不需要
        让位。" 成员列表（NameListWndClass）与我们同质（也是候选列表、也吃键盘），
        让位；参数信息（PopupTipWndClass）只是只读签名，不该让我们消失。

        见模块级 vbe_yield_visible()。纯 Win32、不碰 COM，探不到一律 False
        （不让位）—— 后端不提供本方法时引擎退回 vbe_popup_visible() 的老行为。
        """
        return vbe_yield_visible()

    def vbe_popup_info(self):
        """当前可见的 VBE 提示窗明细 [(类名, 宽, 高), ...]（v65 诊断用）。

        与 vbe_popup_visible() 同一套判据，只是把命中的窗描述出来，让日志能
        区分【成员列表】(NameListWndClass) 与【参数信息】(PopupTipWndClass)。
        纯 Win32 只读，探不到返回 []。引擎只在真正要"让位"那一刻记一行日志，
        正常路径不调用，无开销。
        """
        return vbe_popup_info()

    def vbe_popup_rects(self):
        """当前可见的 VBE 提示窗矩形 [(类名, 左, 上, 右, 下), ...]（v66，UI 避让用）。

        UI 用它把候选窗挪到参数信息窗（形参签名）下方，两个窗不再重叠。
        自定义过程的参数信息用的是同一个窗，一并覆盖。见模块级同名函数。
        """
        return vbe_popup_rects()

    # ---- 自动配对：输入 ( / " 自动补右半边 ----
    #
    # VBE 原生【不会】自动闭合括号和引号（VBE_Extras / Rubberduck 都把它当增强
    # 功能往外加），所以这里补上不会出现"两边都补"的双份。唯一要注意的是 VBE
    # 会在【行尾按回车时】补齐缺失的右引号，那与逐字符输入无关，不受影响。
    _PAIR_CLOSE = {"(": ")", '"': '"', ")": ")",
                   # ★v108：输入法在「中文标点 / 全角」时，钩子会送来全角开符号
                   # （main._pair_char_for_ime），这里给出对应的全角收符号。
                   # ⚠️ `）` **必须自己占一个键**（`"）": "）"`），就像半角的
                   #    `")": ")"` 一样：insert_pair 一进来先查这张表，查不到
                   #    就直接 return False —— 少了这一行，"敲 `）` 时跨过去"那
                   #    条路**根本走不到**（v108 实测：写 `（` 会补出 `（）`，但
                   #    再敲 `）` 却什么都没接管，因为表里只有 `（` 没有 `）`）。
                   # ⚠️ `“` 的收符号是 `”`（**两个不同的字符**），与半角引号
                   #    "开收同字"不一样 —— `_pair_insertion` 里那句"引号特殊处理"
                   #    因此按 open_ch 分了两支。
                   "）": "）", "（": "）", "“": "”"}

    # ★v114：这些开符号在【有选区】时改成"把选区包起来"，而不是让 VBE 用
    # 新字符把选区替换掉。引号（v112）与括号（v114）走同一条路。
    # ⚠️ 收符号（`)` / `）`）刻意**不在**表里 —— 它没有"包起来"的语义，
    #    带选区时仍按老规矩交给 VBE。
    # ⚠️ 单行选区才包：字符串不能跨行；括号跨行要续行符 `_`，包上去就是语法错。
    #    多行选区仍交给 VBE（那正是本表不存在时的老行为）。
    _WRAP_OPEN = ('"', "“", "(", "（")

    @staticmethod
    def _write_line(cm, line_no, text):
        """把 text 写回第 line_no 行。返回是否成功。

        VBE 允许光标停在【模块最后一行的下一行】（一个尚不存在的"虚拟空行"），
        此时 `Lines()` 读出来是空串、看起来一切正常，但 `ReplaceLine()` 会抛
        "无效的过程调用或参数"。真机实测（模块 9 行、光标在第 10 行）确实如此。
        这一行得先用 `InsertLines()` 真的建出来，再写内容。
        """
        try:
            n = cm.CountOfLines
        except Exception:
            n = None                  # 拿不到行数（桩对象等）-> 按普通行处理
        try:
            if n is not None and line_no > n:
                if line_no != n + 1:
                    return False      # 差得离谱，不猜，交给调用方兜底
                cm.InsertLines(line_no, text)
            else:
                cm.ReplaceLine(line_no, text)
            return True
        except Exception:
            return False

    def insert_pair(self, open_ch):
        """在光标处插入一对符号（`(` -> `()`、`"` -> `""`），光标落在中间。

        ★v108：`open_ch` 也可能是**全角**（`（` -> `（）`、`“` -> `“”`）——
        输入法在「中文标点 / 全角」时由 main._pair_char_for_ime 换算后送进来，
        收符号查 _PAIR_CLOSE。**绝不在这一层自己"猜"输入法状态**：钩子线程
        已经问过一次了，后端再问一次就是两份判据（项目铁律第 23 条）。

        ★v112 两条（用户口径）：
          * 敲**左括号**时**右侧还有内容就不补右半边** —— 判据在 `_pair_insertion`；
          * **选中一段文本再敲引号** -> 用引号把选区包起来（`abc` -> `"abc"`），
            由 `_wrap_selection` 完成；多行选区仍照旧交给 VBE。

        ★v114（用户口径）：上面第二条**对括号同样生效** —— 选中 `abc` 敲 `(`
        得到 `(abc)`（全角 `（` 得 `（abc）`）。开符号白名单是 `_WRAP_OPEN`；
        收符号（`)` / `）`）不在其中，带选区时仍交给 VBE 替换。

        ★v113（用户口径）：上面第一条有个例外 —— **在某个调用的实参表里、或某个
        过程的形参表里敲 `(`，仍要补全右半边**（`foo(bar(` -> `foo(bar()`、
        `Sub Foo(ByVal arr(` -> `Sub Foo(ByVal arr())`），因为那里的右侧内容属于
        **外层**括号。判据是 `parser.enclosing_call_paren`（纯函数）。

        返回 True = "这一下我代为输入了"（调用方据此吞掉原按键）；
        返回 False = "我没管"（在注释里 / 有选区 / COM 出问题），调用方必须
        **放行按键**让 VBE 按原生行为处理 —— 绝不吞掉用户的按键却什么都没发生。
        """
        close_ch = self._PAIR_CLOSE.get(open_ch)
        if not close_ch:
            return False
        new_line = None
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return False
            cp = vbe.ActiveCodePane
            if cp is None:
                return False
            cm = cp.CodeModule
            sl, sc, el, ec = cp.GetSelection()
            if sl != el or sc != ec:
                # ★v112：单行选区 + 引号 -> 用引号把选中的内容包起来。
                # ★v114：括号同一路（`abc` 选中敲 `(` -> `(abc)`）。
                # 其余选区（多行 / 收符号 / 右括号）仍是"替换语义交给 VBE，别插"。
                if sl == el and sc < ec and open_ch in self._WRAP_OPEN:
                    return self._wrap_selection(cm, cp, sl, sc, ec,
                                                open_ch, close_ch)
                return False          # 有选区：替换语义交给 VBE，别插
            line_text = cm.Lines(sl, 1)
            col, sem, tabw, wide2 = _caret_char_col(
                cm, cp, sl, line_text, ec)
            col0 = max(0, min(col - 1, len(line_text)))
            if _classify(line_text, col0 + 1)[1]:
                return False          # 注释里不自动配对（VBE 也不）
            if _in_call_stmt(line_text):
                # ★v102：`Call 过程名` 里的空括号会被 VBE 立刻删掉（VBA 规范），
                # 我们写入 `()` 等于白写 —— 用户看到的是"敲 `(` 完全没反应"。
                # 返回 False 让按键【原样交给 VBE】，它自己插入 `(`；
                # 用户接着输参数时得到 `Call judge(1, 2)`，那对括号会保留。
                return False
            paired = False
            if open_ch in (")", "）"):
                # 右半边已经在光标右边 -> 只跨过去（补完 `()` 后再按 `)` 不该
                # 多出一个）；否则【放行】，让用户真的插入一个右括号。
                # ★v108：全角 `）` 与半角 `)` 同路（`open_ch` 本身就是配置里的
                # 那个字符，收符号映射到它自己，所以直接比它即可）。
                if col0 >= len(line_text) or line_text[col0] != open_ch:
                    return False
                new_line, caret_off = None, col0 + 1
            else:
                new_line, caret_off, paired = _pair_insertion(
                    line_text, col0, open_ch, close_ch)
            if new_line is not None:
                if not self._write_line(cm, sl, new_line):
                    return False
                try:
                    actual = cm.Lines(sl, 1)
                except Exception:
                    actual = new_line
                if actual != new_line:
                    # VBE 把整行重新格式化过了（见 _locate_inserted 的说明）：
                    # 在改写后的行里重新定位我们插入的那对符号，把光标钉到
                    # 它中间；实在认不出来才退到行尾附近（绝不越界）。
                    if paired:
                        off = _locate_inserted(actual, new_line, open_ch,
                                               close_ch, col0)
                    else:
                        # ★v112：这次只插了左括号（右侧有内容），行里根本没有
                        # `()` 可找 —— 改按"压平前缀"定位（与 _locate_inserted
                        # 首选同一指纹口径，只是指纹里没有收符号）。
                        off = _locate_by_prefix(actual, new_line, col0 + 1)
                    # v79 起把这一路记进日志：VBE 到底怎么重排的、我们又把它
                    # 认到了哪儿 —— 以后再出"光标没落在括号里"，一眼就能看出是
                    # 定位错了还是 VBE 根本没重排（仅 run_debug.bat 下有开销）。
                    _log("autopair: VBE 重排了整行 -> 重新定位"
                         " (写入=%r 实际=%r col0=%d 定位=%r)"
                         % (new_line, actual, col0, off))
                    if off is not None:
                        caret_off = off
                    else:
                        caret_off = min(caret_off, len(actual))
            else:
                # 右半边已经在光标右边：只把光标挪过去，一个字符都不写
                actual = line_text
                caret_off = min(caret_off, len(actual))
            if sem == "disp":
                col = _disp_width(actual, tabw, wide2, upto=caret_off) + 1
            else:
                col = caret_off + 1
            # ★代码复查修（2026-10-18）：落光标单独 try。文本已经写进去了
            # （`_write_line` 返回 True / 右半边已存在时压根不用写），
            # 这里再抛异常**不能**把结果拉成 False —— 调用方（main 的
            # `_auto_pair_here`）看到 False 会把这一下按键原样还给系统，
            # 于是行里已经有 `()` 又多出一个 `(`。与 duplicate_lines /
            # delete_lines_here 同一口径。
            try:
                cp.SetSelection(sl, col, sl, col)
            except Exception:
                _log("autopair: 文本已写、落光标失败（按成功返回）")
            return True
        except Exception:
            return False
        finally:
            # 与 apply_completion 同理：写过代码（VBAProject 变脏），立刻放手。
            if new_line is not None:
                _release_vbe_proxy()

    def _wrap_selection(self, cm, cp, sl, sc, ec, open_ch, close_ch):
        """★v112/v114：把【单行选区】用一对符号包起来（`abc` -> `"abc"` / `(abc)`）。

        用户口径（原话，v112）："我可以选中一段文本，然后输入双引号后，能自动给
        选中的内容括上双引号。"；v114："括号也可以实现同样效果。"

        只做单行 —— 调用方已经判过 `sl == el`：VBA 的字符串不能跨行；括号跨行
        要续行符 `_`，包上去就是语法错。多行选区仍交给 VBE 的替换语义（那也是
        本函数不存在时的老行为）。

        列口径：`sc` / `ec` 都是 **VBE 的显示列**（全角占 2 格），而
        `line_text` 是字符序列 —— 必须换回字符列。这里复用 `_caret_char_col`
        + `_col_to_char_index`（`get_context` / `insert_pair` 同一套），
        **绝不在这里自己另写一份换算**（项目铁律第 16/23 条）。
        两个列用同一组 (sem, tabw, wide2)：一次探测同时管住两端，避免
        "起点按这套、终点按那套"的错位。

        光标落在末尾收符号**之后**。VBE 若重排了整行（补空格 / 改大小写），
        用 `_locate_by_prefix` 按压平后的前缀找回位置 —— 与 insert_pair
        只插左括号时同一条兜底。写不成就返回 False，调用方把按键还给系统。
        """
        try:
            line_text = cm.Lines(sl, 1)
            ec_char, sem, tabw, wide2 = _caret_char_col(
                cm, cp, sl, line_text, ec)
            if sem == "disp":
                sc_char = _col_to_char_index(line_text, sc, tabw, wide2) + 1
            else:
                sc_char = min(sc, len(line_text) + 1)
            a = max(0, min(sc_char - 1, len(line_text)))
            b = max(a, min(ec_char - 1, len(line_text)))
            if b <= a:
                return False          # 实际没选中字符：交给 VBE，别乱包
            new_line = (line_text[:a] + open_ch + line_text[a:b]
                        + close_ch + line_text[b:])
            if not self._write_line(cm, sl, new_line):
                return False
            try:
                actual = cm.Lines(sl, 1)
            except Exception:
                actual = new_line
            tgt = a + len(open_ch) + (b - a) + len(close_ch)
            if actual != new_line:
                off = _locate_by_prefix(actual, new_line, tgt)
                if off is None:
                    off = min(tgt, len(actual))
                _log("autopair: 包裹选区后 VBE 重排了整行 -> 重新定位"
                     " (写入=%r 实际=%r 定位=%r)" % (new_line, actual, off))
            else:
                off = tgt
            if sem == "disp":
                col = _disp_width(actual, tabw, wide2, upto=off) + 1
            else:
                col = off + 1
            # ★代码复查修（2026-10-18）：同上 —— 包裹已经写进去了，
            # 落光标失败不能让调用方再注入一个开符号。
            try:
                cp.SetSelection(sl, col, sl, col)
            except Exception:
                _log("autopair: 包裹已写、落光标失败（按成功返回）")
            return True
        except Exception:
            return False
        finally:
            # 与 insert_pair 同理：写过代码（VBAProject 变脏），立刻放手。
            # （没写时也走一下，release 只是丢缓存引用，无副作用。）
            _release_vbe_proxy()

    # ---- v115：代码片段 ----
    def insert_snippet(self, text):
        """★v115：把一段（可能多行的）代码片段插到**光标处**。返回是否写成。

        用户口径（原话）："点击按钮，可以将选中代码片段插入到vbe中的光标处。"

        落点模型 —— 整行拼接，**片段原样插入、不替用户重排缩进**（片段就是
        用户自己存的排版）：
            前半行 + 片段第 1 行
            片段第 2..n-1 行
            片段最后 1 行 + 后半行
        于是"光标在行中间"也能插（后半行会被接在片段末尾那一行后面），
        "有选区"则是替换语义（单行 / 多行选区都支持）。

        ⚠️ **写入是"尽量原子"的**（★v124 代码复查）：中间的 n-1 行用**一次**
        `InsertLines` 灌进去，然后才覆盖首行 —— 原先是逐行插，第 3 行抛异常时
        前 2 行已经进了文档、函数却返回 False，调用方提示"插入失败"，用户再点
        一次就是**片段插两遍**。现在任何一步失败的落点都是：
          * 中间行插入失败 → 一次调用失败 = 行数没变，原文档一个字没动；
          * 首行覆盖失败 → 把刚插的中间行删掉，回到"什么都没发生"。
        `InsertLines` 支持多行是**真机实测**过的（Excel 2016 / VBE7.1：`\\r\\n`、
        `\\n`、`\\r` 三种都正确拆行），与 `duplicate_lines` 同一口径。

        ⚠️ `ReplaceLine` 会把**行尾空格 trim 掉**（v103b 实测），而 `InsertLines`
        不会。所以片段末行带尾随空格时，只有"片段是单行且要替换当前行"这一
        条路会被吃掉尾部空白 —— 可忽略（VBE 存的就是 trim 过的代码行）。

        失败一律返回 False，由调用方提示（绝不静默）。
        """
        wrote = False
        try:
            text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
            if not text.strip():
                return False
            vbe = _get_vbe_cached()
            if vbe is None:
                return False
            cp = vbe.ActiveCodePane
            if cp is None:
                return False
            cm = cp.CodeModule
            sl, sc, el, ec = cp.GetSelection()
            if sl < 1 or el < sl:
                return False
            try:
                n_lines = int(cm.CountOfLines)
            except Exception:
                n_lines = None

            def _line(no):
                if n_lines is not None and no > n_lines:
                    return ""          # 模块末尾那个"虚拟空行"
                return cm.Lines(no, 1) or ""

            first = _line(sl)
            last = first if el == sl else _line(el)
            # 起点：VBE 的列是显示列（全角 2 格 / Tab 展开），必须换回字符下标
            col_s, sem, tabw, wide2 = _caret_char_col(cm, cp, sl, first, sc)
            i = max(0, min(col_s - 1, len(first)))
            # ★代码复查修（2026-10-18）：终点也必须换算。原先单行选区
            # （`el == sl`）直接拿 `ec`（**显示列**）当字符下标用 —— 含全角
            # 字符或 Tab 的行上会多删/漏删一个字符；而 `tail = last[j:]` 之后
            # 是【整行重写】，删错就是用户代码真的少了几个字。
            # 多行选区那一支本来就换算，两端口径必须一致（同一组 sem/tabw/wide2）。
            col_e, _s2, _t2, _w2 = _caret_char_col(cm, cp, el, last, ec)
            # ★代码复查修（2026-10-18）：`max(i, …)` 只对单行选区有意义 ——
            # 多行时 `i` 是【首行】的字符下标，跟末行 `last` 毫无关系，
            # 从某行中部往下选到下一行行首（很常见的选法）会把 `j` 顶到 `i`，
            # 末行没被选中的那一段被整段丢弃。
            if el == sl:
                j = max(i, min(col_e - 1, len(last)))
            else:
                j = max(0, min(col_e - 1, len(last)))
            head, tail = first[:i], last[j:]

            block = text.split("\n")
            if len(block) == 1:
                out = [head + block[0] + tail]
            else:
                out = [head + block[0]]
                out.extend(block[1:-1])
                out.append(block[-1] + tail)

            # ★v124（代码复查）：写入顺序改成"先插中间行（一次调用）、再覆盖
            # 首行"，让"写了一半"的窗口消失。原来的写法是逐行 `InsertLines`，
            # 第 3 行抛异常时前 2 行已经进去了、函数却返回 False —— 调用方提示
            # "插入失败"，用户再点一次 = 片段插两遍。
            #
            # 为什么现在敢一次灌多行：★真机实测（Excel 2016 / VBE7.1）
            # `InsertLines` 对 `\r\n`、`\n`、`\r` 三种换行**都正确拆行**
            # （3 行文本插进去就是 3 行，含尾随换行时多一个空行）。原先那段
            # "没有文档保证所以逐行插"的注释是多余的谨慎，`duplicate_lines`
            # 一直在用整段插 —— 两处口径现在已经统一。
            k = len(out) - 1
            # 光标可能停在【模块末尾的虚拟空行】（sl == CountOfLines+1）：
            # 那一行还不存在，`InsertLines(sl+1, …)` 会插到行尾之外。先把它
            # 建成一个真实的空行 —— 这是纯新增，失败等于什么都没动。
            # （单行片段不用这一步：`_write_line` 自己会建，只有一次写。）
            _virtual = n_lines is not None and sl > n_lines and k > 0
            if _virtual:
                try:
                    cm.InsertLines(sl, "")
                except Exception:
                    _log("snippet: 建首行失败（原文未改动）")
                    return False
                wrote = True
            # 1) 中间行：一次调用，要么全成要么全不成。
            if k:
                try:
                    cm.InsertLines(sl + 1, "\r\n".join(out[1:]))
                except Exception:
                    if _virtual:
                        # 只多了刚才那个空行 —— 收掉就回到原样。
                        try:
                            cm.DeleteLines(sl, 1)
                            wrote = False
                        except Exception:
                            _log("snippet: 中间行插入失败且回滚失败")
                    _log("snippet: 中间 %d 行插入失败（原文未改动）" % k)
                    return False
            # 2) 首行：走到这里它一定已经存在（虚拟空行上面刚建过）。
            if not self._write_line(cm, sl, out[0]):
                # 半截片段比"失败"更糟：把刚插进去的行全删掉。
                # ⚠️ 只删【新增的】那些行 —— 第 sl 行原本就有内容（`_write_line`
                # 失败说明它没被改），连它一起删就是把用户代码删没了；只有
                # `_virtual`（那一行本来不存在、是我们刚建的）才连 sl 一起删。
                try:
                    if _virtual:
                        cm.DeleteLines(sl, k + 1)
                    elif k:
                        cm.DeleteLines(sl + 1, k)
                    wrote = False
                    _log("snippet: 首行写入失败，已回滚 %d 行" % (k + 1 if _virtual else k))
                except Exception:
                    _log("snippet: 首行写入失败且回滚失败（多出 %d 行）"
                         % (k + 1 if _virtual else k))
                return False
            wrote = True
            # 多行选区：把被替换掉的旧行收掉（插进来的 k 行把它们挤下去了）
            # 片段已经完整写进去了 —— 失败只记日志、不返回 False：内容在文档里，
            # 报失败会让用户再插一遍。
            drop = el - sl
            if drop > 0:
                try:
                    cm.DeleteLines(sl + 1 + k, drop)
                except Exception:
                    _log("snippet: 片段已插入，旧行删除失败（多出 %d 行）" % drop)

            # 摆光标：落在片段最后一个字符之后（= 尾部原文的起点）
            line_no = sl + k
            try:
                actual = cm.Lines(line_no, 1) or ""
            except Exception:
                actual = out[-1]
            pos = _snippet_tail_pos(actual, tail)
            if sem == "disp":
                col = _disp_width(actual, tabw, wide2, upto=pos) + 1
            else:
                col = pos + 1
            # ★代码复查修（2026-10-18）：片段已经插进去了，落光标失败不能让
            # 调用方看到 False —— 那会提示"插入失败"，用户再点一次就重复插入。
            try:
                cp.SetSelection(line_no, col, line_no, col)
            except Exception:
                _log("snippet: 已插入、落光标失败（按成功返回）")
            _log("snippet: 插入 %d 行 -> 行=%d 光标列=%d (前=%r 后=%r)"
                 % (len(block), line_no, col,
                    (head or "")[-20:], (tail or "")[:20]))
            return True
        except Exception:
            return False
        finally:
            # 与 apply_completion / insert_pair 同理：写过代码（VBAProject
            # 变脏）就立刻放手，别让宿主退出被我们拽住。
            if wrote:
                _release_vbe_proxy()

    def toggle_comment(self):
        """★v120：`Ctrl+/` 批量注释 / 取消注释（IDEA 语义）。返回是否处理了。

        用户口径（原话）："Ctrl+/ 批量注释和Ctrl+/ 批量解除注释，这个帮我弄一下，
        就像IDEA那样。"

        判据（哪些行要动、加还是取消）全在纯函数 `parser.plan_comment_action`；
        本方法只负责**读 → 写 → 摆光标**：

          1. `GetSelection()` 取范围。选区跨行且停在**行首**（`EndCol <= 1`）时，
             最后那一行其实没被选中 —— 那是"连着换行符一起拖下来"的结果，
             IDEA 也一样排除它；
          2. 逐行 `Lines(n, 1)` 读原文（**只读这一小段**，不碰整个模块）；
          3. 纯函数给出 `(action, [(相对行, 新文本)])`；什么都不能动（全空白）
             → 返回 False，调用方把这一下按键**原样还给系统**；
          4. **写前锚点复核**：把要动的行再读一遍、与第 2 步逐字相同才动手
             （项目铁律：改用户工程前先断言锚点）。不同就整笔放弃；
          5. 逐行 `_write_line`（即 `ReplaceLine`）。**中途失败也继续写完** ——
             写一半就停会让文档卡在"部分注释"的状态，比写完更糟；返回 True
             （确实改过了，再还键没有意义），失败明细进日志；
          6. 原来有选区 → 把同样这几行重新选起来（用户连按两次就能取消，
             这是 IDEA 的手感）；原来只是光标 → 把光标摆回原行原列。

        ⚠️ **撤销栈**：一次动 N 行就是 N 步撤销。VBE 的 COM 没有"命令组"，
        做不到一次全退 —— 文档里如实写明了（`docs/troubleshooting.zh-CN.md`）。

        ⚠️ `ReplaceLine` 会 trim **行尾**空格（v103b 实测）。注释动作只改行首，
        但只要这一行被写过，行尾空白就会丢 —— VBE 存盘本来就 trim，影响趋零。
        """
        wrote = False
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return False
            cp = vbe.ActiveCodePane
            if cp is None:
                return False
            cm = cp.CodeModule
            sl, sc, el, ec = cp.GetSelection()
            if sl < 1 or el < sl:
                return False
            try:
                n_lines = int(cm.CountOfLines)
            except Exception:
                n_lines = None

            def _line(no):
                if n_lines is not None and no > n_lines:
                    return ""          # 模块末尾那个"虚拟空行"
                return cm.Lines(no, 1) or ""

            # "跨行且停在行首"= 最后那一行没被真正选中
            if el > sl and ec <= 1:
                el -= 1
            if n_lines is not None:
                el = min(el, n_lines)
            if el < sl:
                return False

            old = [_line(n) for n in range(sl, el + 1)]
            action, plan = vba_parser.plan_comment_action(old)
            if not action or not plan:
                return False

            # 写前锚点复核（读与写之间不该有别人动过这几行）
            for i, _new in plan:
                if _line(sl + i) != old[i]:
                    _log("comment: 锚点变了，整笔放弃（行=%d）" % (sl + i))
                    return False

            had_sel = ((sl, sc) != (el, ec))
            ok_all = True
            for i, new in plan:
                if not self._write_line(cm, sl + i, new):
                    ok_all = False
                    _log("comment: 写入失败 行=%d（已写的行不回退）" % (sl + i))
            wrote = True
            _log("comment: %s %d 行 -> 行 %d..%d 成功=%s"
                 % ("取消注释" if action == "remove" else "加注释",
                    len(plan), sl, el, ok_all))

            # 摆光标：行数没变，所以原来的行范围仍然有效
            new_last = dict(plan).get(len(old) - 1, old[-1])
            try:
                if had_sel:
                    # 让 VBE 自己把过长的列钳到行尾（铁律：摆光标的列一律现场量
                    # 或让 VBE 钳，绝不自己算）
                    cp.SetSelection(sl, 1, el, len(new_last) + 100)
                else:
                    cp.SetSelection(sl, max(1, int(sc)), sl, max(1, int(sc)))
            except Exception:
                pass
            return True
        except Exception:
            _log("comment: toggle_comment 异常")
            return False
        finally:
            # 与 apply_completion / insert_pair / insert_snippet 同理：写过代码
            # （VBAProject 变脏）就立刻放手，别让宿主退出被我们拽住。
            if wrote:
                _release_vbe_proxy()

    # ---- ★v121：过程大纲 / 查找引用 ----

    def current_module_text(self):
        """当前代码窗格的 `(模块名, 全文)`；取不到返回 `(None, "")`。

        与 `snapshot()` 的分工：那个只读**当前一行**、每次轮询都跑；这个要**全文**
        （大纲要的是准确行号），只在用户显式按键时跑一次，贵一点没关系。

        刻意不做列语义探测、也不动光标（那会临时移动光标，是 v93 起就避开的手法）。
        ⚠️ VBE 的 `Lines()` 回来的是 CRLF，这里统一成 `\\n` —— 否则 `split("\\n")`
        每行尾巴都留个 `\\r`，行号虽然还对，但拿去做正则/界面显示都会带上脏字符。
        """
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return (None, "")
            cp = vbe.ActiveCodePane
            if cp is None:
                return (None, "")
            cm = cp.CodeModule
            try:
                mod = str(cm.Name or "")
            except Exception:
                mod = ""
            try:
                n = int(cm.CountOfLines)
            except Exception:
                n = 0
            if n <= 0:
                return (mod, "")
            text = str(cm.Lines(1, n))
            return (mod, text.replace("\r\n", "\n").replace("\r", "\n"))
        except Exception:
            _log("outline: 读模块全文失败")
            return (None, "")

    def list_procedures_here(self):
        """`(模块名, [(名字, 种类, 行号, 是否 Private, 过程头原文), ...])`。

        判据全在纯函数 `parser.list_procedures`（与补全那边认的是同一批正则），
        这里只负责把当前模块的文本拿过来，并顺手带上每一行**原文** —— 大纲窗
        选中某一项时把它显示在下面，用户不用跳过去就知道签名长什么样。
        """
        mod, code = self.current_module_text()
        if not code:
            return (mod, [])
        try:
            procs = vba_parser.list_procedures(code)
        except Exception:
            return (mod, [])
        lines = code.split("\n")
        out = []
        for row in procs:
            name, kind, line, priv = row
            text = ""
            if 1 <= line <= len(lines):
                text = lines[line - 1].strip()
            out.append((name, kind, line, priv, text))
        return (mod, out)

    def word_at_caret(self):
        """`(模块名, 光标处（或紧邻左边）的那个名字)`；不在名字上返回 `(模块名, "")`。

        用 `get_context()` 换算列（它已经把**显示列**换成**字符列**了 —— 行内含
        全角字符时这一步自己做必然偏，铁律 15/16），再交给 `parser.ident_at_caret`
        取整词。光标停在词尾（`DoWork|`）也算命中，那是查找引用最常见的姿势。
        """
        try:
            ctx = self.get_context()
        except Exception:
            ctx = None
        if not ctx:
            return (None, "")
        mod = ctx.get("module_name")
        word = ""
        try:
            word = vba_parser.ident_at_caret(
                ctx.get("line_text") or "", 1, ctx.get("caret_col") or 0)
        except Exception:
            word = ""
        return (mod, word or "")

    def find_references(self, name, max_hits=400):
        """在【当前工程】的所有模块里做**文本级**查找。

        返回 `(当前模块名, [(模块名, 行号, 列号, 该行原文, 是否声明处,
        能不能跳过去), ...], 被扫描的模块数)`；当前模块的命中排在最前。

        最后一列的"能不能跳过去"= 那个模块**有没有打开代码窗**（`vbe.CodePanes`
        里有它）。没有的话 COM 里没有任何接口能打开它（见
        `_code_panes_by_module` 的说明），界面据此提前标出来，别让用户回车了才发现
        跳不动。

        刻意说"文本级"：不解析类型、不区分同名变量在不同过程里的作用域 —— 那需要
        一张真实的符号表，VBE 只给我们纯文本。界面与文档里都照这个口径写死，
        绝不假装是"查找引用（语义）"。

        一次性操作直接读 `comp.CodeModule`（不走 `_cm_cache`）：那是补全热路径的
        缓存，键是模块名，塞进这里的临时读取没有收益，反而多一份要维护的不一致。
        """
        name = str(name or "").strip()
        if not name:
            return (None, [], 0)
        mod0 = None
        hits = []
        scanned = 0
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return (None, [], 0)
            panes = _code_panes_by_module(vbe)
            try:
                cp = vbe.ActiveCodePane
                if cp is not None:
                    mod0 = str(cp.CodeModule.Name or "")
            except Exception:
                mod0 = None
            comps = _collect_components(vbe, caret_mod=mod0)
            for comp in comps:
                try:
                    cm = comp.CodeModule
                    n = int(cm.CountOfLines)
                except Exception:
                    continue
                if n <= 0:
                    continue
                try:
                    mname = str(comp.Name or "")
                except Exception:
                    mname = ""
                try:
                    text = str(cm.Lines(1, n))
                except Exception:
                    continue
                scanned += 1
                text = text.replace("\r\n", "\n").replace("\r", "\n")
                try:
                    found = vba_parser.find_occurrences(
                        text, name, limit=(max_hits - len(hits)) or 1)
                except Exception:
                    found = []
                jumpable = str(mname).lower() in panes
                for (ln, col, line, decl) in found:
                    hits.append((mname, ln, col, line, decl, jumpable))
                if len(hits) >= max_hits:
                    break
            # 当前模块排最前（用户八成是从这里出发的），其余保持扫描顺序
            if mod0:
                low = str(mod0).lower()
                hits.sort(key=lambda h: 0 if str(h[0]).lower() == low else 1)
            _log("refs: 查 %r -> %d 处 / %d 个模块（%d 个开着代码窗）"
                 % (name, len(hits), scanned, len(panes)))
            return (mod0, hits, scanned)
        except Exception:
            _log("refs: 查找引用失败")
            return (mod0, hits, scanned)

    def goto_location(self, module_name, line_no, expect=None):
        """把光标跳到 `(模块, 行)`，成功返回 True。

        三处刻意：

          1. **落点一律第 1 列**。`SetSelection` 收的是**显示列**，行内含全角字符
             时自己去换算必然偏（铁律 15/16）；而第 1 列没有任何歧义。
          2. **锚点复核**：目标行必须非空白；给了 `expect`（过程名 / 查找的那个
             名字）时还必须在那一行里出现。扫描到大纲列出、到用户按下回车之间，
             文件完全可能被改过 —— 行号一旦失真，宁可这一次不跳，也绝不把光标
             扔到一段不相干的代码上（v21 的老规矩）。
          3. 不写任何字节，也不调 `_release_vbe_proxy()`（没动过代码，VBAProject
             还是干净的）。

        ⚠️ 只能跳**已经打开过代码窗**的模块：`VBComponent` 上没有 `CodePane` 属性，
        而没打开过的模块在 COM 里没有打开它的办法（见 `_code_panes_by_module`）。
        """
        want = str(module_name or "").strip()
        if not want:
            return False
        try:
            line = int(line_no)
        except Exception:
            return False
        if line <= 0:
            return False
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return False
            low = want.lower()
            cp = _code_panes_by_module(vbe).get(low)
            if cp is None:
                _log("goto: 模块 %r 没有打开的代码窗（CodePanes 里没有）" % want)
                return False
            cm = cp.CodeModule
            try:
                n = int(cm.CountOfLines)
            except Exception:
                n = 0
            if n <= 0:
                return False
            if line > n:
                line = n
            try:
                text = str(cm.Lines(line, 1)).rstrip("\r\n")
            except Exception:
                return False
            if not text.strip():
                _log("goto: 目标行是空白 line=%d（行号可能已失真）" % line)
                return False
            if expect:
                if str(expect).lower() not in text.lower():
                    _log("goto: 锚点复核不过 line=%d text=%r expect=%r"
                         % (line, text, expect))
                    return False
            cp.Show()
            cp.SetSelection(line, 1, line, 1)
            _log("goto: %s:%d ok" % (want, line))
            return True
        except Exception:
            _log("goto: 跳转异常 %r:%r" % (want, line_no))
            return False

    def focus_vbe(self):
        """把焦点还给 VBE 主窗（片段窗关掉之后用）。成功与否只影响手感。"""
        try:
            vbe = _get_vbe_cached()
            if vbe is not None:
                try:
                    vbe.MainWindow.SetFocus()
                    return True
                except Exception:
                    pass
        except Exception:
            pass
        try:
            import ctypes
            u32 = ctypes.windll.user32
            u32.FindWindowW.restype = ctypes.c_void_p
            u32.FindWindowW.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p]
            hwnd = u32.FindWindowW("wndclass_desked_gsk", None)
            if hwnd:
                u32.SetForegroundWindow(ctypes.c_void_p(hwnd))
                return True
        except Exception:
            pass
        return False

    # ---- 当前光标上下文 ----
    def get_context(self):
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return None
            cp = vbe.ActiveCodePane
            if cp is None:
                return None
            sl, sc, el, ec = cp.GetSelection()
            if sl != el:  # 存在多行选择，不触发
                return None
            cm = cp.CodeModule
            line_text = cm.Lines(sl, 1)

            # 顺序是关键：必须「先换算列，再判定字符串/注释」。
            # VBE 的 ec 是显示列（中文等全角占 2 格），而 line_text 是字符序列。
            # 若直接拿 ec 去 _classify 索引 line_text，行内含全角时 ec-1 会
            # 超过 len(line_text)-1 -> IndexError -> get_context 返回 None
            # -> 中文输入永远不弹列表（纯英文行显示列==字符列，所以一直正常）。
            ec_char, sem, tabw, wide2 = _caret_char_col(
                cm, cp, sl, line_text, ec)
            in_str, in_comment = _classify(line_text, ec_char)
            # 声明里 `As` 之后的数据类型名位置：正在填类型名，不该提示变量名。
            # 即便后面的变量名前缀匹配上已有标识符，也应保持静默。
            in_type = vba_parser.is_caret_in_type_position(line_text, ec_char)
            # 正在"起新名字"的声明行（Dim/Const/Sub/Function/Public 变量...）：
            # 此刻任何候选都是噪音，且正在输入的词本身常已存在于标识符池
            # （它在别处被用过、或别的模块有同名声明）-> 表现为提示出自己。
            # 这里直接整行静默，不弹窗。
            in_decl = vba_parser.is_caret_in_declaration(line_text, ec_char)

            # 当前声明行正在声明的名字集合（小写）。供引擎在 in_decl_position
            # 时精确剔除"正在输入的那个词本身"（打全名才隐藏，前缀照常提示），
            # 从而修好"输入 num 只提醒 num、不提醒 numArr"的前缀补全误杀。
            # 只在声明行（含 v53 的续行）上算，平时为空、开销极小。
            decl_names = []
            # v53：声明【续行】也算声明行。形参常常独占一行（长签名被拆成
            # `Sub Foo( _` + 缩进的形参行），续行本身没有 Sub/Dim 关键字，
            # is_caret_in_declaration() 看不出来；但用户在那里打 / 退格 /
            # 粘贴形参名，与在首行上是同一件事——不把这条逻辑行的声明名算出来，
            # 引擎就没法拦住"形参提示形参自己"（用户报的正是续行上的退格）。
            # 判定很便宜：只额外读上一行，看它是否以续行符 `_` 结尾。
            cont_decl = False
            if not in_decl and sl > 1:
                try:
                    cont_decl = vba_parser.is_continuation_line(
                        cm.Lines(sl - 1, 1))
                except Exception:
                    cont_decl = False
            # ★v111：声明位置上【该提示哪些语言关键字】（形参表 / 声明头）。
            # 与 decl_names 同源同路：都只在声明行上算，平时为空、开销极小。
            # 空 = 这个声明位置没有关键字可提示（变量名 / 过程名 / Type 成员行），
            # 引擎那边照 v105 静默。
            decl_keywords = []
            if in_decl or cont_decl:
                try:
                    full = cm.Lines(1, cm.CountOfLines)
                    dn = vba_parser.decl_names_at_caret(full, (sl, ec_char))
                    if dn:
                        decl_names = sorted(dn)
                    if in_decl:
                        _kw = vba_parser.decl_position_keywords(
                            full, (sl, ec_char))
                        if _kw:
                            decl_keywords = sorted(_kw)
                except Exception:
                    decl_names = []
                    decl_keywords = []

            try:
                module_name = str(cm.Name)
            except Exception:
                module_name = None

            # ★v98：`标识符.` 左边那个标识符的名字（如 `c.` 里的 `c`）。
            #
            # ⚠️ 这里**只取名字、不取类型**，而且"是不是成员位置"那道判据
            # 也**不在这里判** —— 它归 engine 的 `_member_list_dot_col`
            # 一家（v93 把它合并成唯一实现，理由见那里的注释：同一判据写两份
            # 必然悄悄分叉）。engine 判完"是成员位置"之后，自己拿这个名字来
            # 问类型、问成员。
            #
            # 所以本函数只做纯取值：光标左边正在输入的成员名不算，往回找到
            # 点号，再取点号左边紧邻的标识符；取不到就返回 None。
            member_root_name = None
            try:
                member_root_name = vba_parser.ident_before_dot(
                    line_text, ec_char)
            except Exception:
                member_root_name = None

            return {
                "line_no": sl,
                "caret_col": ec_char,          # 字符列（1-based）
                "line_text": line_text,
                "in_string": in_str,
                "in_comment": in_comment,
                "in_type_position": in_type,
                "in_decl_position": in_decl,
                "decl_names": decl_names,
                # ★v111：声明位置上该提示的语言关键字（空 = 照旧静默）。
                "decl_keywords": decl_keywords,
                "proc_name": _proc_of_line(cm, sl),
                "module_name": module_name,
                # ★v98：点号左边那个标识符的名字；None = 那里没有标识符。
                # "是不是成员位置"由 engine 判（见上面那段说明）。
                "member_root_name": member_root_name,
            }
        except Exception:
            return None

    # ---- 应用补全 ----
    def apply_completion(self, line_no, word_start_col, end_col, completion,
                         want_parens=False, want_space=False):
        """把 [word_start_col, end_col) 区间的单词替换为 completion 并写回 VBE。

        word_start_col / end_col 均为 1-based 字符列（与 Python 行文本一致）。
        end_col 通常是光标列（光标左边那截已输入的前缀）；光标停在标识符【开头】
        （按 Delete 从名字开头删字符，v44）或【中间】（从名字中间删字符，v46）时
        它才大于光标列——此时要把光标右边那段残留名字一起换掉，否则会在残留词
        前面插入候选、拼出双份名字。

        want_parens（★v98）：True 时，若该位置是【调用点】（`_is_call_site`）
        且词右边本来没有左括号，就在补全词后面补一对 `()` 并把光标送进去。
        **调用方必须先确认 completion 确实是过程/函数**（引擎用
        `get_proc_names()` 判）—— 这里只管"位置允不允许"，不管"是什么"。

        want_space（★v103b）：True 时（且 `AUTO_SPACE_AFTER_KEYWORD` 开着），
        若补全词的词尾右边不是空白，就在词尾补一个【空格】。用于 VBA 关键字
        （`pub` -> `public `）。**调用方必须先确认 completion 是这种关键字**
        （引擎用 `vba_builtins.keyword_wants_space` 判）—— 这里只管位置。
        ⚠️ 空格**不是**改写行文本补的：真机实测 VBE 会把 ReplaceLine 写进去的
        行尾空格立刻 trim 掉，所以只能在光标就位之后【注入一个空格按键】
        （`self._send_char`，由 main 注入；没有它就不补，绝不半途改坏代码）。

        写回后把光标放到补全词尾部：
          - 若该行(含 Tab/全角)在 VBE 里按"显示列"计，就把字符列换算成显示列，
            否则光标会因 Tab 展开的格差落在词的中间；
          - 若 VBE 按"字符列"计（探测结果 'char'），直接用字符列。
        """
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return None
            cp = vbe.ActiveCodePane
            if cp is None:
                return None
            cm = cp.CodeModule
            line_text = cm.Lines(line_no, 1)
            new_line = replace_word(line_text, word_start_col, end_col, completion)

            # ★v98：补括号。判据全在这三行，且**任何一步不确定就不补**
            # —— 补坏代码的代价远大于少补一对括号。
            added_parens = False
            # ★v101 诊断：后端这三道输入（want_parens / 开关 / 位置闸门）
            # 才是"括号到底补没补"的决定者。v100 那次修不好就是因为只验了
            # 引擎侧（_is_callable_name）与离线 mock，没看后端这三项的真值。
            _log("apply_completion: want_parens=%s AUTO_PAREN=%s "
                 "call_site=%s | line=%r ws=%s completion=%r"
                 % (want_parens, AUTO_PAREN_ON_COMPLETE,
                    _is_call_site(line_text, word_start_col),
                    str(line_text).rstrip("\r\n"), word_start_col, completion))
            if want_parens and AUTO_PAREN_ON_COMPLETE \
                    and _is_call_site(line_text, word_start_col) \
                    and not _in_call_stmt(line_text):
                # 补全词的**词尾**（0-based 下标）。括号要贴在这个词的后面。
                k = word_start_col - 1 + len(completion)
                # 跳过空白**只是为了判断**"右边是不是已经紧跟左括号"
                # （`x = getVa|(1)` 这种括号本来就有的调用点）。
                #
                # ⚠️ ★v98 修：插入点必须是【词尾 k】，不能是被跳过空白后的 j。
                # 用 j 插入会把括号甩到空白后面 —— `x = DoWo + 1` 曾补成
                # `x = DoWork ()+ 1`（括号与函数名之间多一个空格，而且
                # `+` 前的空格被吃掉，读起来像 `DoWork ()+ 1`）。
                j = k
                while j < len(new_line) and new_line[j] in " \t":
                    j += 1
                if not (j < len(new_line) and new_line[j] == "("):
                    new_line = new_line[:k] + "()" + new_line[k:]
                    added_parens = True

            # ★v103b：关键字后面补一个空格。
            # ⚠️ 这里**不**改写 new_line —— 真机实测 VBE 会把行尾空格立刻 trim 掉，
            # 改写等于白写。只算好"要往哪个下标注入一个空格"（纯函数，
            # 便于离线用例钉住），真正的注入在光标摆好之后做。
            #
            # ★v103c：`Call 语句` 里不补 —— `Call judge` 的实参位置语法不同
            # （带实参要写成 `Call judge(1, 2)`），补一个尾随空格纯属无意义字符，
            # 而且那个位置的空括号本来就会被 VBE 删掉（v102 实测）。判据与上面
            # 补括号那道用的是**同一个** `_in_call_stmt`（同一语义只有一份实现）。
            space_at = None
            if want_space and AUTO_SPACE_AFTER_KEYWORD \
                    and not _in_call_stmt(line_text):
                space_at = _space_insert_pos(new_line, word_start_col, completion)
            _log("apply_completion: want_space=%s AUTO_SPACE=%s space_at=%s"
                 % (want_space, AUTO_SPACE_AFTER_KEYWORD, space_at))

            # 词尾字符下标（0-based）
            start0 = max(0, word_start_col - 1)
            end0 = start0 + len(completion)

            cm.ReplaceLine(line_no, new_line)
            # ★v102 诊断：ReplaceLine 到底有没有写进 VBE。
            # 真机实测（独立 Excel 实例）出现过"返回算好的新行、VBE 里却没变"
            # —— 必须看清是 ReplaceLine 静默失败，还是读回被缓存。
            try:
                _log("apply_completion: ReplaceLine 后 line_no=%r n=%r "
                     "写入=%r 读回=%r"
                     % (line_no, cm.CountOfLines, new_line,
                        cm.Lines(line_no, 1)))
            except Exception as _e102:
                _log("apply_completion: ReplaceLine 后读回异常 %r" % (_e102,))
            # 读回实际行（VBE 可能做规范化，内容略有变化）
            actual = cm.Lines(line_no, 1)
            if actual != new_line:
                end0 = _find_completion_end(actual, completion,
                                            end0 - len(completion), end0)
            end0 = min(end0, len(actual))
            # ★v107c：留给下面"要不要唤出参数信息"用 —— 光标最终落在哪一行文本上。
            # 若"二次确认"发现 VBE 又改写了这一行，下面会把它换成实际那一行。
            param_line = actual
            # VBE 自动补的括号（`Function gCalc` -> `Function gCalc()`）：
            # 光标要落进括号里，否则停在左括号左侧还得手动右移一格。
            #
            # ★v98：这一条同时负责**我们自己在 v98 补的那对括号** —— 上面的
            # 插入位置恰好是词尾，所以 `actual[end0]` 就是那个新 `(`，这里
            # 自然把光标送进括号里，逻辑完全复用、不必另写一段。
            #
            # ★v103b：补空格时【不】走这一段 —— 关键字不补括号，光标必须稳稳
            # 停在词尾（再注入空格），不能被 `_skip_into_parens` 带进某个 `(`。
            if space_at is None:
                end0 = _skip_into_parens(actual, end0, line_text, end_col)
                end0 = min(end0, len(actual))
            else:
                end0 = min(space_at, len(actual))

            # 关键：列语义必须在【写回之后】的实际行上量，而且**不能拿缓存**。
            #
            # 老故障一（v24 那轮）：用户只打了首字母（纯 ASCII，如 `abc中文` 只打了
            # `a`），旧行不含全角 -> 探测被跳过、默认 'char'；可补全后的新行含中文，
            # VBE 按显示列计 => 按字符列定位会少算格差，光标落在变量名中间。
            #
            # ★v95 老故障二（用户这次报的）：探测结果被 `_sem_cache` 缓存，而那条
            # 结论可能是【在纯 ASCII 行上】探出来的 —— 那种行上两种列语义数值相同，
            # 根本分不出来，却被记成了 'char' 并留给后面所有行用（v94b 同一个坑）。
            # 于是即使这一行含中文，拿缓存也会算成字符列 —— 症状与上面一模一样，
            # 但"探测没被跳过"，光修"基于写回后的行"治不了。
            #
            # 所以这里**在 actual 这一行上重新量**：它自带 Tab / 全角，证据是
            # 判别性的，量出来的一定对。用 `_vbe_measure` 而不是 `_probe_semantics`：
            # 紧接着就要摆光标，不必还原（少两次 COM、也不受 probe_broken 影响）。
            # 量到的结果反过来写回缓存 —— 等于第一次中文补全就把缓存修好了。
            sem, tabw, wide2 = "char", 4, False
            if "\t" in actual or any(_is_wide(c) for c in actual):
                info = _vbe_measure(cm, cp, line_no, actual)
                if info is not None:
                    sem, tabw, wide2 = info
                    _sem_cache["info"] = info

            def _set_caret(text, off):
                """按当前列语义把光标放到字符下标 off 之后。"""
                if sem == "disp":
                    col = _disp_width(text, tabw, wide2, upto=off) + 1
                else:
                    col = off + 1
                cp.SetSelection(line_no, col, line_no, col)

            _set_caret(actual, end0)
            if space_at is not None:
                # ★v103b：光标已在关键字**词尾**，注入一个空格。
                # 这一下是【注入进来的字符】，VBE 不会像行尾空格那样把它 trim
                # 掉 —— 真机实测：注入空格后读回 `"    public "`，再注入 'S'
                # 得到 `"    public S"`。
                # ⚠️ 注入函数由 main 传进来（`VbeBackend(send_char=...)`）；
                # 没有它（旧式构造 / 离线测试）就什么都不做 —— 宁可"没补空格"，
                # 也绝不半途改坏代码。
                try:
                    _send = getattr(self, "_send_char", None)
                    if callable(_send):
                        if not _send(" "):
                            _log("apply_completion: 注入空格失败（SendInput 未命中）")
                except Exception as _e103b:
                    _log("apply_completion: 注入空格异常 %r" % (_e103b,))
                return new_line
            # 二次确认：个别情况下 VBE 是在我们读完行【之后】才补括号的，
            # 那一对 `()` 会把光标"挤"到左括号左侧。再读一次行，若它又变了、
            # 且光标处正是新出现的 `(`，就右移一格把光标放进括号。
            try:
                actual2 = cm.Lines(line_no, 1)
            except Exception:
                actual2 = actual
            if actual2 != actual:
                param_line = actual2
                end0b = _skip_into_parens(actual2, end0, line_text, end_col)
                if end0b != end0:
                    end0 = min(end0b, len(actual2))
                    try:
                        _set_caret(actual2, end0)
                    except Exception:
                        pass
            # ★v107c：光标已经在括号中间了 -> 可选地唤出 VBE 的【参数信息】窗。
            # ★v124c：从"注入 Ctrl+Shift+I"改成"调 VBE 的菜单命令"。
            # ★v124d：真机实测两条路都唤不出窗（`PopupTipWndClass` 始终不可见），
            #         而且 `Execute()` 每次都响一声 ⇒ **默认整段不做**，只有
            #         `VBECOMPLETE_PARAM_INFO=1` 才进来。详见 AUTO_PARAM_INFO
            #         上面那段 peak 实测记录。
            # 三道条件缺一不可，且全部是"宁可不做"的方向：
            #   ① `added_parens` —— 只有【我们补出来的】那对括号才管；
            #      用户自己敲的 `(`，VBE 在敲下去的那一刻就弹过了；
            #   ② 开关 `AUTO_PARAM_INFO`（默认关）；
            #   ③ 光标此刻**真的**停在 `(` 右边（位置判据只此一份实现）。
            # 命令拿不到 / 宿主不支持 CommandBars -> 静默跳过。
            if added_parens and AUTO_PARAM_INFO \
                    and _caret_inside_parens(param_line, end0):
                if _show_param_info(vbe):
                    _log("apply_completion: 参数信息命令已发出（菜单命令）")
                elif PARAM_INFO_KEY_FALLBACK:
                    # 只有用户显式要求时才退回"模拟按键"那条会响的老路。
                    try:
                        _qi = getattr(self, "_send_quick_info", None)
                        if callable(_qi):
                            _log("apply_completion: 退回注入 Ctrl+Shift+I -> %s"
                                 % (_qi(),))
                        else:
                            _log("apply_completion: 无 CommandBars 也无注入函数")
                    except Exception as _e107c:
                        _log("apply_completion: 唤出参数信息异常 %r" % (_e107c,))
                else:
                    _log("apply_completion: 拿不到参数信息命令 -> 静默跳过")
            return new_line
        except Exception:
            return None
        finally:
            # 刚用 ReplaceLine 改过 VBA 代码（VBAProject 变脏），此刻手上还攥着
            # CodeModule / CodePane / VBE 的引用，而 Excel 退出时正是要把这个
            # 工程写回文件 —— 立刻放手，别把它的退出流程钉住。
            # （"只改工作表不卡、改过 VBA 才卡"的针对性处理。）
            _release_vbe_proxy()

    # ---- 新起一行（Shift+Enter） ----
    # 往上读多少行来"跳过注释行"（够用即可：注释再长也不会超过这个数）。
    _INDENT_LOOKBACK = 200
    # ★v86：补收尾前"往下找本块自己的收尾"的扫描行数上限。刻意和
    # _INDENT_LOOKBACK 分开：上面那个管"往上找缩进基准"（只看最近几行就够），
    # 这里要真的看到本块的收尾 —— 原先两者共用 200，于是**长过程 / 长 If 的
    # 收尾会落到窗口之外** ⇒ 判成"下面没有收尾" ⇒ 回车再补一条（双份收尾）。
    # 定得越大越准，代价只是每次"块头行尾回车"多读几行（只在这条少见路径上，
    # 读满也就几十 KB）；真读满还判不出来时由 closer_needed 的 truncated
    # 那一档按"不补"兜底。
    _CLOSER_SCAN_LINES = 600

    def _lines_above(self, cm, line_no, limit=None):
        """取第 line_no 行【上方】的各行文本（最近的在最后）。

        只读、且只在"回车由我们接管"这一条少见路径上调用，行数封顶
        _INDENT_LOOKBACK —— 免得在几千行的模块里把半个模块都读进来。
        """
        n = self._INDENT_LOOKBACK if limit is None else int(limit)
        start = max(1, int(line_no) - n)
        cnt = int(line_no) - start
        if cnt <= 0:
            return []
        try:
            text = cm.Lines(start, cnt)
        except Exception:
            return []
        lines = str(text).split("\n")
        if lines and lines[-1] == "":
            lines.pop()                    # Lines() 末尾带回车，会多一个空元素
        return [l.rstrip("\r") for l in lines]

    def _lines_below(self, cm, line_no, limit=None):
        """取第 line_no 行【下方】的各行文本（最近的在最前）。

        v79b 补收尾时用：得知道下面是不是已经挂着 `End If` / `Next` 了，
        免得替用户补出一个重复的收尾。同样只读；行数默认封顶 _INDENT_LOOKBACK，
        调用方（补收尾那条路）会显式传 `_CLOSER_SCAN_LINES` —— 见那里的说明，
        以及 closer_needed 的 `truncated` 参数。
        """
        n = self._INDENT_LOOKBACK if limit is None else int(limit)
        try:
            total = int(cm.CountOfLines)
        except Exception:
            return []
        start = int(line_no) + 1
        cnt = min(n, total - start + 1)
        if cnt <= 0:
            return []
        try:
            text = cm.Lines(start, cnt)
        except Exception:
            return []
        lines = str(text).split("\n")
        if lines and lines[-1] == "":
            lines.pop()                    # Lines() 末尾带回车，会多一个空元素
        return [l.rstrip("\r") for l in lines]

    def new_line_below(self, smart=False, auto_close=False, align_branch=True):
        """在光标所在行的【下方】新起一行，光标落在缩进之后。

        等价于"先把光标移到本行末尾，再按回车"，但一步到位 —— 而且当前行
        【不拆分】（光标右侧的代码留在原行）。VBE 只在"光标已在行尾"时按回车
        才继承上一行缩进；光标停在行中间时按回车会把行拆开，所以必须由我们代劳。

        smart=False（Shift+Enter，v78 起；★v89 起缩进判定与回车【同源】）：
        不依赖光标列、也不改动本行内容 —— 直接读本行文本算新行缩进：
          * 整行注释：忽略它（以及上方连续的注释行），跟上方第一个非注释行对齐；
          * 块头（含 `Else` / `Case` / `#If … Then` 这些【没有收尾】的块头）：
            新行深一级，而且 `Else` / `Case` 会先把【本行】拉回它该在的层级
            （branch_align，与回车那条路同一份判据）；
          * 空行：沿用本行自己的空白 —— 不替用户猜层级（回车那一档在这种行上
            是直接放手、交回 VBE 原生，行为等价）；
          * 其余普通行：抄本行行首空白（= v78 老口径，与回车那条路同结果）。

        smart=True（回车自动缩进，v79）—— 由用户那两条要求而来：
          * 本行是【整行注释】：忽略它（以及上方连续的注释行），跟上方第一个
            非注释行对齐；
          * 本行是【块结构开头】（For / Do / If…Then / With / Sub / Function /
            Type / Enum / Select Case …）：新行缩进一级。
        缩进由纯函数 next_line_indent 算；这一模式下多了四道保险，任一不满足
        就返回 False 让调用方把回车【原样还给系统】：
          * 有选区：多行选择按回车是删除/覆盖，语义完全不同；
          * 光标不在本行行尾：那是"拆行"，必须交给 VBE 原生回车；
          * 本行是空行（只有空白）：让 VBE 自己继承上一行缩进，别跟用户手动
            排好的缩进较劲；
          * 本行引号没闭合：VBE 原生回车会替用户补上右引号，别抢这个活。

        ★v88（Shift+Enter 也补收尾）：auto_close 不再只对 smart 模式生效 ——
        smart=False（Shift+Enter）那条路现在也判块头、补收尾，**且不要求光标在
        行尾**（用户报的："光标在代码中间按 Shift+Enter，不会自动补全代码块"）。
        两条路共用同一份 block_closer + closer_needed + next_line_indent，
        差别只在"怎么起步"：
          * smart=True（回车）：光标不在行尾 -> 整件事作废，把回车还给 VBE
            （那一档 VBE 原生是【拆行】，我们不能替它决定）；
          * smart=False（Shift+Enter）：这条路本来就不拆行（只往下插一行），
            所以光标停在哪儿都无所谓 —— 照判、照补。
        ★v89（Shift+Enter 的缩进判定并入同一份判据）：v88 只把"新行深一级"
        接在了【有收尾】的块头里面（写在 `if _close:` 支里），于是 `Else` /
        `Case` / `#If … Then` 这些"是块头但没有收尾"的行、以及整行注释，都仍
        走 v78 那条"行首空白抄本行"的老路 —— 用户报的正是这个："按 Shift+Enter
        换行，不会判定自动缩进，只按回车会判定自动缩进。"
        现在这一支的缩进统一由 next_line_indent + branch_align 算（与 smart
        分支逐字同源）。真正保留的老口径只剩两条：**空行不猜层级**、
        **`auto_close` 关掉时不补收尾**（收尾才是那个开关管的语义）。

        auto_close=True（回车自动补收尾，v79b；嵌套判定 v79c）—— smart 模式下：
        块头是 `If x Then` / `With rng` / `Sub Foo()` / `Type Foo` / `Enum E` 这类
        【真需要收尾】的，再在其后补一行同级的 `End If` / `End With` / `Next` /
        `End Type`…，光标停在中间那行（缩进正好），用户直接在块里写内容。要不要补
        由纯函数 block_closer + closer_needed 决定：下面已经挂着本块的收尾、或者
        这块的下文已经在写了 —— 都不补，免得顶开已有代码。

        align_branch=True（分支行自动对齐，v81；层级分档 v83）—— 两个模式都生效
        （★v89 起；此前只管 smart 模式）：
        本行是块【内】的分支（`Else` / `ElseIf … Then` / `Case …`，含条件编译的
        `#Else` / `#ElseIf`）时，先把它【自己】拉回该在的层级（`Else` / `ElseIf` 与
        `If … Then` 齐平、`#Else` / `#ElseIf` 与 `#If … Then` 齐平、**`Case` 比
        `Select Case` 深一级**），新行再深一级。
        要不要对齐、对齐到哪，由纯函数 branch_align 决定：上方找不到那个"还开着的块头"
        就什么都不做（宁可不对齐，也不要挪错）。**只改行首空白**，行内一个字符都不碰。

        ⚠️ v79c：内层块下面压着的往往是【外层】的收尾（双层 `For` 里的 `Next`），
        它的缩进正好比内层块头浅一格 —— 单看缩进会误判成"已经有收尾了"。所以
        closer_needed 还会查"上方有没有一个同类、未闭合、缩进正好对上它的块头"，
        有就说明那个收尾属于外层，本块照样得补。

        成功返回 True；不在代码窗 / 取不到 COM / 出任何异常都返回 False ——
        调用方据此把这一下按键原样还给系统（回车绝不能吞掉）。
        """
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return False
            cp = vbe.ActiveCodePane
            if cp is None:
                return False
            cm = cp.CodeModule
            sl, sc, el, ec = cp.GetSelection()
            line_text = None
            closer_text = None             # 要补的块收尾（smart + auto_close 才可能有）
            # ⚠️ v81：这个必须在 if smart 之外初始化 —— 放在 smart 分支里的话，
            # smart=False（Shift+Enter）走到下面那句 `if fixed_line is not None`
            # 会抛 NameError，被 `except -> return False` 吞成"Shift+Enter 不换行"
            # （第 29.4~29.7 / 56.8e / 57.4h 当场全红，和 v79b 的 closer_text 同坑）。
            fixed_line = None              # 分支行要拉回它所属块头的缩进
            if smart:
                if int(sl) != int(el) or int(sc) != int(ec):
                    return False           # 有选区：不接管
                anchor = int(sl)
                if anchor <= 0:
                    return False
                line_text = cm.Lines(anchor, 1)
                if not str(line_text).strip():
                    return False           # 空行：交给 VBE
                if _unterminated_string(line_text):
                    return False           # 让 VBE 去补右引号
                col, _sem, _tw, _w2 = _caret_char_col(
                    cm, cp, anchor, line_text, ec)
                if line_text[col - 1:].strip():
                    return False           # 光标后面还有内容 -> 是"拆行"
                above = self._lines_above(cm, anchor)
                unit = _indent_unit_for(above + [line_text])
                indent = next_line_indent(line_text, above, unit)
                if align_branch:
                    _al = branch_align(line_text, above, unit)
                    if _al is not None:
                        fixed_line, indent = _al
                if auto_close:
                    # 块头才谈得上收尾；"要不要真的补"交给纯函数判（下面可能
                    # 已经挂着同一个收尾、或者这块的下文已经在写了）。
                    # v79c：把块头类型与上方各行一并交给它 —— 判"更浅的同类收尾
                    # 是不是外层块的"（双层 For 的 `Next` 就长这样）要用。
                    _close = block_closer(line_text)
                    if _close:
                        base = _leading_ws(line_text)
                        # ★v86：往下要找的是"本块自己的收尾"，所以要按
                        # _CLOSER_SCAN_LINES 读（原先跟着 _INDENT_LOOKBACK 用
                        # 200，长过程的收尾会落到窗口外）。真读满被截断时把
                        # truncated 交给判据 —— 那时"窗口内没找到"不能当
                        # "下面没有"，宁可少补一条也不要变出双份收尾。
                        try:
                            _rest = max(0, int(cm.CountOfLines) - anchor)
                        except Exception:
                            _rest = 0
                        _cap = min(_rest, self._CLOSER_SCAN_LINES)
                        if closer_needed(self._lines_below(cm, anchor, _cap),
                                         _close,
                                         _indent_width(base),
                                         kind=_block_kind(line_text),
                                         lines_above=above,
                                         truncated=_rest > _cap):
                            closer_text = base + _close
            else:
                # 有选区时以选区【末行】为基准（正常情况下 sl == el）。
                anchor = max(int(sl), int(el))
                if anchor <= 0:
                    return False
                line_text = cm.Lines(anchor, 1)
                above = self._lines_above(cm, anchor)
                unit = _indent_unit_for(above + [line_text])
                if not str(line_text).strip():
                    # 空行：不替用户猜层级。回车那一档在这种行上是直接放手、把回车
                    # 交回 VBE 原生（= 沿用本行自己的空白），所以这里照抄本行空白
                    # 就是同一个结果。
                    indent = _leading_ws(line_text)
                else:
                    # ★v89（用户报的）：Shift+Enter 的【缩进判定】与回车那条路逐字
                    # 同源 —— 同一份 next_line_indent（整行注释往上找第一个非注释
                    # 行、块头深一级）+ branch_align（`Else` / `Case` / `#Else` 先把
                    # 【本行】拉回它该在的层级）。
                    #
                    # 为什么 v88 之后仍然"不判定"：v88 那段"新行深一级"写在下面
                    # `if _close:` 里面，而 `_close = block_closer(line_text)` 对
                    # `Else` / `Case` / `#If … Then`（`_block_kind` 认它们是块头、
                    # 但**没有收尾**，见 _BLOCK_CLOSERS 的说明）返回 None ⇒
                    # 这几类行、以及整行注释，全都落回了 v78 那条"行首空白抄本行"
                    # 的老路，看起来就是"只有回车才判定自动缩进"。
                    indent = next_line_indent(line_text, above, unit)
                    if align_branch:
                        _al = branch_align(line_text, above, unit)
                        if _al is not None:
                            fixed_line, indent = _al
                # ★v88（用户报的）：Shift+Enter 也要能把块补起来，而且【光标停在
                # 哪一列都行】。回车那条路（smart=True）有一条"光标必须在行尾"的
                # 保险：光标在代码中间时它会把这一下原样还给 VBE，而 VBE 原生回车
                # 在那种位置是【拆行】—— 于是用户拿不到收尾。Shift+Enter 这条路
                # 本来就不拆行（只往下插一行、本行一个字符都不动），所以可以放心地
                # 在任意光标列上判块头、补收尾。
                #
                # 判据走【同一份实现】：block_closer / closer_needed（v87 的教训：
                # 同一条语义规则只能有一份实现），与 smart 分支逐字同源 —— 只是
                # 这里不必先问"光标在不在行尾"。**收尾**才是 `auto_close` 管的事，
                # 缩进不归它管（`VBECOMPLETE_NO_AUTOCLOSE=1` 的语义就是"只留缩进、
                # 不补收尾"）。
                if auto_close:
                    _close = block_closer(line_text)
                    if _close:
                        base = _leading_ws(line_text)
                        # 往下找本块收尾的窗口与"读满被截断"的兜底同 smart 分支
                        # （_CLOSER_SCAN_LINES / closer_needed 的 truncated）。
                        try:
                            _rest = max(0, int(cm.CountOfLines) - anchor)
                        except Exception:
                            _rest = 0
                        _cap = min(_rest, self._CLOSER_SCAN_LINES)
                        if closer_needed(self._lines_below(cm, anchor, _cap),
                                         _close,
                                         _indent_width(base),
                                         kind=_block_kind(line_text),
                                         lines_above=above,
                                         truncated=_rest > _cap):
                            closer_text = base + _close
            if fixed_line is not None:
                # v81：分支行（Else / ElseIf…Then / Case…）先拉回它该在的层级
                # （v83：Case 比 Select Case 深一级，其余与块头齐平）。
                # 只动【行首空白】，行内一个字符都不碰；写不成就不动、也不插新行，
                # 让调用方把这一下回车原样还给系统。
                _raw = str(line_text).rstrip("\r\n")
                _ws = _leading_ws(_raw)
                if _ws != fixed_line and not self._write_line(
                        cm, anchor, fixed_line + _raw[len(_ws):]):
                    return False
            # InsertLines 在 anchor+1 处插入（anchor == CountOfLines 时即追加到末尾），
            # 已有行自动下移 —— 正是"在本行下方新起一行"。
            cm.InsertLines(anchor + 1, indent)
            if closer_text:
                # 收尾插在"新起的空行"下面：块头下面留出缩进好了的一行给用户写
                # 内容，收尾再往下、缩进回到块头同级。原有行整体再下移一行。
                cm.InsertLines(anchor + 2, closer_text)
            # ★代码复查修（2026-10-18）：新行**已经插进去了** —— 下面这一段
            # （重新探测行语义 + 摆光标）再出任何异常都不能返回 False。
            # 原先它整段在同一个 try 里，抛异常就 return False ⇒ 调用方
            # （main `_enter_indent_here` / `_new_line_here`）会再注入一个回车
            # ⇒ 用户按一下回车多出一行。与 delete_lines_here 同一口径。
            try:
                col = _indent_end_col(
                    indent, _resolve_sem_info(cm, cp, anchor + 1))
                cp.SetSelection(anchor + 1, col, anchor + 1, col)
            except Exception:
                _log("newline: 行已插入、落光标失败（按成功返回）")
            return True
        except Exception:
            return False
        finally:
            # 与 apply_completion 同理：写过代码（VBAProject 变脏），立刻放开
            # 手里的 COM 代理，别拖住 Excel 的退出/写回。
            _release_vbe_proxy()

    # ---- ★v122：三个"按行"的编辑动作（插行 / 复制行 / 删行） ----
    #
    # 用户口径（原话）："编辑手感可以帮我添加一下：1. Ctrl+Shift+Enter 在上方插入
    # 新行；2. Shift+d 复制当前行到下方新行；3. Shift+k 删除整行。"
    # 键位最后定成 Ctrl+D / Ctrl+K —— Shift+D / Shift+K 在代码窗里就是"输入大写
    # D / 大写 K"，接管了用户在字符串和注释里就再也敲不出大写字母（VBE 的自动
    # 大写只管关键字）。
    #
    # 三个动作的共同口径：
    #   * **有选区就作用于选中的那几行**（与 IDEA / VS Code 一致），没选区就是
    #     光标那一行；
    #   * 每次只写一次（InsertLines / DeleteLines 各一次），绝不边读边写；
    #   * 写不成返回 False，调用方把按键原样还回去；
    #   * 落点列只在"第 1 列"和"照抄原列"两种没有歧义的地方选（铁律 15/16：
    #     显示列 vs 字符列，中间换算必错）。

    def _sel_block(self, cp, cm):
        """取"这次要动的那几行"：有选区 = 整个选区，没选区 = 光标那一行。

        返回 `(首行, 尾行, 起始显示列, 总行数)`，取不到返回 None。
        `GetSelection` 的两个角点**不保证 sl <= el**（从下往上选就是反的），
        所以这里一律先归一。
        """
        try:
            sl, sc, el, ec = cp.GetSelection()
            sl, el = int(sl), int(el)
            if sl > el:
                sl, el = el, sl
            if sl <= 0:
                return None
            try:
                n_lines = int(cm.CountOfLines)
            except Exception:
                n_lines = 0
            if sl > n_lines:
                return None
            if el > n_lines:
                el = n_lines
            return sl, el, max(1, int(sc)), n_lines
        except Exception:
            return None

    def new_line_above(self):
        """在光标所在行的【上方】新起一行，光标落在新行的缩进之后（★v122）。

        对应 IDEA 的 "Start New Line Before Current"（`Ctrl+Alt+Enter`）。

        与 `new_line_below` 的分工：那条问的是"本行是不是块头"（是的话新行
        深一级、还要补收尾）；这条问的是"本行是不是**块收尾**"——插在
        `End If` / `Next` / `End Sub` 上方，等于插在**块内的最后一行**，所以
        要深一级；其余情况新行与本行【同级】，缩进就是本行的行首空白。

        刻意只做这一条智能：插在上方时用户看到的是"我上面多了一行"，只要
        缩进和本行齐平就不会误判；再猜 `Else` / `#If` 该对齐到哪，反而容易
        挪错（branch_align 那套是给"新起一行"用的，不是给"插到上方"用的）。

        有选区时插在**选区最上面那一行**的上方 —— 用户眼里"上方"就是那个位置。
        `InsertLines(anchor, …)` 正是"插在第 anchor 行之前"，原行整体下移。
        """
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return False
            cp = vbe.ActiveCodePane
            if cp is None:
                return False
            cm = cp.CodeModule
            blk = self._sel_block(cp, cm)
            if blk is None:
                return False
            anchor = blk[0]
            line_text = str(cm.Lines(anchor, 1))
            raw = line_text.rstrip("\r\n")
            indent = _leading_ws(raw)
            if str(raw).strip() and is_block_closer(raw):
                # 本行是块收尾 -> 新行落在块内，深一级。同样要求"本行非空"：
                # 空行不是任何收尾（is_block_closer 对空行本来就返回 False，这里
                # 只是把意图写明，省得以后有人以为漏判了）。
                above = self._lines_above(cm, anchor)
                unit = _indent_unit_for(above + [line_text])
                indent = indent + unit
            cm.InsertLines(anchor, indent)
            # 插入后新行就是 anchor 行；缩进末尾的列由 _indent_end_col 给
            # （它按"缩进只有空白"的语义直接算，不需要问光标列）。
            # ★代码复查修（2026-10-18）：行已经插进去了，落光标失败不能返回
            # False（否则调用方补一个回车 = 多出一行）。
            try:
                col = _indent_end_col(
                    indent, _resolve_sem_info(cm, cp, anchor))
                cp.SetSelection(anchor, col, anchor, col)
            except Exception:
                _log("insert-above: 行已插入、落光标失败（按成功返回）")
            return True
        except Exception:
            return False
        finally:
            _release_vbe_proxy()

    def duplicate_lines(self):
        """把光标所在行（有选区 = 选中的那几行）**复制一份**插到下方（★v122）。

        对应 IDEA 的 Duplicate Line（`Ctrl+D`）。

        写入只有一次 `InsertLines`：`cm.Lines(sl, n)` 拿回来的原文本身就是
        CRLF 连接的多行，直接喂回去即可 —— ★v124 真机实测（Excel 2016 /
        VBE7.1）：`InsertLines` 对 `\r\n` / `\n` / `\r` 三种换行**都按行拆开**，
        3 行文本插进去就是 3 行（`Lines()` 的 join 不带尾随换行，所以原样喂回
        不会多出空行；只有取到"最后一个空行"时才会带一个 `\r\n`，那也正是
        要复制的空行）。`insert_snippet` 已改成同一口径，两处不再互相矛盾。

        一次 COM 调用 = 要么全成、要么全不成，天然没有"复制了一半"的中间态。
        光标落在**复制块的第一行、同一列** —— 整行复制意味着那一行的内容
        逐字相同，所以"显示列"照抄就是同一个字符位置，没有列语义换算的坑。
        """
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return False
            cp = vbe.ActiveCodePane
            if cp is None:
                return False
            cm = cp.CodeModule
            blk = self._sel_block(cp, cm)
            if blk is None:
                return False
            sl, el, sc, _n = blk
            count = el - sl + 1
            txt = str(cm.Lines(sl, count))
            cm.InsertLines(el + 1, txt)
            new_line = el + 1
            try:
                cp.SetSelection(new_line, sc, new_line, sc)
            except Exception:
                pass        # 写已经成了，落光标失败不能再报 False（否则会多复制一份）
            return True
        except Exception:
            return False
        finally:
            _release_vbe_proxy()

    def delete_lines_here(self):
        """删掉光标所在行（有选区 = 选中的那几行），光标落到**上一行的行尾**（★v122/v126b）。

        对应 VS Code 的 Delete Line（`Ctrl+Shift+K`）/ IDEA 的 `Ctrl+Y`。

        ★v126b 落点改口径（用户报「删完光标应该停在被删行上一行的行尾」）：
        原来是落到"下面顶上来的那一行"的第 1 列，现在落到**被删块上方那一行
        （`sl - 1`）的末尾** —— 删完继续打字时，接的就是上一行原来那句话，
        不用先按 End。上一行是空行时行尾就是第 1 列，语义自洽。

        列的算法不自己猜：`_caret_to_line_end` 给一个远超行尾的列让 VBE 自己
        夹回行尾（含全角 / Tab 时按字符列算必错，见那里的说明）；它没走通才
        退回 `_eol_caret_col`。两条路与 `delete_blank_line_here` 完全同一份实现。

        删的是第 1 行（没有"上一行"）时落到新首行第 1 列；删到一行不剩
        （整个模块被删空）时**不落光标** —— 那时行 1 已经不存在，`SetSelection`
        会抛异常。

        ⚠️ 把整个模块删空是允许的：VBE 的 `CountOfLines` 会变成 0，此时
        `SetSelection(1, 1)` 会抛异常 —— 但删除本身已经成功了，所以落光标
        单独 try 包住，绝不让它把返回值拉成 False（那会导致调用方把按键还回去、
        用户又白删一次）。
        """
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return False
            cp = vbe.ActiveCodePane
            if cp is None:
                return False
            cm = cp.CodeModule
            blk = self._sel_block(cp, cm)
            if blk is None:
                return False
            sl, el = blk[0], blk[1]
            # ★v126b：落点要落到**上一行**行尾，所以先把上一行的原文取下来 ——
            # 删除不动它，但删完之后再读就要看 CountOfLines 的脸；删除前读最稳。
            prev_no = sl - 1
            prev_raw = ""
            if prev_no >= 1:
                try:
                    prev_raw = str(cm.Lines(prev_no, 1)).rstrip("\r\n")
                except Exception:
                    prev_raw = ""
            cm.DeleteLines(sl, el - sl + 1)
            try:
                left = int(cm.CountOfLines)
            except Exception:
                left = 0
            if left > 0:
                try:
                    if prev_no < 1:
                        # 删的就是第 1 行：没有"上一行"，落在新首行行首
                        cp.SetSelection(1, 1, 1, 1)
                    elif not _caret_to_line_end(cp, cm, prev_no, prev_raw):
                        # 兜底：该编辑器不夹列 -> 才回去按列语义算（同 v94b）
                        _c = _eol_caret_col(prev_raw, _sem_cache.get("info"))
                        cp.SetSelection(prev_no, _c, prev_no, _c)
                except Exception:
                    pass
            return True
        except Exception:
            return False
        finally:
            _release_vbe_proxy()

    # ---- 退格删整行（v94：整行只有空白时，一次退格删掉整行） ----
    def delete_blank_line_here(self):
        """光标所在的【空行】（只有空白）整行删掉，光标落到【上一行行尾】。

        用户口径：「按退格的时候，如果整行是空行，能删除整行，IDEA 里写 Java
        就有这个功能」。钩子侧的纯判据见 blank_line_delete_wanted —— 那里已经
        按轮询快照筛过一遍，这里是**真正的权威**：再读一次真实行文本确认，
        因为快照最多是 100ms 前的（用户可能刚粘贴完就按退格）。

        落点刻意选【上一行行尾】：与 VBE 原生"空行上按退格"完全一致（原生是
        吃掉换行符、两行并成一行，光标停在上一行末尾）。删完之后用户接着按
        退格删的就是上一行的字 —— 正是他按住退格时期待的行为。

        ★v94b 落点的【列】不再由我们算：交给 VBE 自己钳（`_caret_to_line_end`
        给一个远超行尾的列，VBE 夹回行尾）。原先按列语义算，上一行含汉字时会
        少算全角的格数，光标落在中文中间 —— 详见那里的说明。

        ⚠️ 这些情形一律返回 False（交给 VBE 原生退格，调用方会把这一下按键
        原样还给系统，绝不会出现"吞了按键却什么都没发生"）：
          * 有选区：退格是"删选区"，语义完全不同；
          * 真空行（`""`）：原生退格的结果与删整行一模一样，没必要抢；
          * 行里有内容：那是普通删字；
          * 取不到 COM / 模块只读 / 出任何异常。

        单行一次写：只在确认完毕之后调一次 DeleteLines（绝不边读边写、也绝不
        在 finally 里补写）。落光标单独 try —— DeleteLines 已经成了却因为落
        光标失败而返回 False，会让调用方再补一次退格、多删一个字符。
        """
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return False
            cp = vbe.ActiveCodePane
            if cp is None:
                return False
            cm = cp.CodeModule
            sl, sc, el, ec = cp.GetSelection()
            if int(sl) != int(el) or int(sc) != int(ec):
                return False                    # 有选区：那是"删选区"
            anchor = int(sl)
            if anchor <= 0:
                return False
            try:
                n_lines = int(cm.CountOfLines)
            except Exception:
                n_lines = 0
            if anchor > n_lines:
                return False
            raw = str(cm.Lines(anchor, 1)).rstrip("\r\n")
            # 权威复核：判据与钩子侧【同一份实现】（blank_line_delete_wanted），
            # 只是喂进去的是【此刻真读到的】行与列 —— 快照最多是 100ms 前的，
            # 中间用户可能用鼠标把光标点到缩进中间（那时他要的是"退一格"，
            # 不是删整行），也可能刚粘贴进一整段代码。这里一票否决，调用方会把
            # 这一下退格原样还给系统，绝不会误删。
            if not blank_line_delete_wanted(raw, ec):
                return False
            prev_no = anchor - 1
            prev_raw = ""
            if prev_no >= 1:
                prev_raw = str(cm.Lines(prev_no, 1)).rstrip("\r\n")
            cm.DeleteLines(anchor, 1)
            try:
                if prev_no < 1:
                    # 删的是第 1 行：没有"上一行"，光标落在新首行行首
                    cp.SetSelection(1, 1, 1, 1)
                elif not _caret_to_line_end(cp, cm, prev_no, prev_raw):
                    # 兜底：该编辑器不夹列 -> 才回去按列语义算（见 docstring）
                    _c = _eol_caret_col(prev_raw, _sem_cache.get("info"))
                    cp.SetSelection(prev_no, _c, prev_no, _c)
            except Exception:
                pass
            return True
        except Exception:
            return False
        finally:
            _release_vbe_proxy()

    # ---- v99：退格成对删除（空括号 / 空引号） ----
    def delete_pair_here(self):
        """删掉光标左右紧邻的那一对**空**括号 / 空引号，光标停在它们原来的
        位置。★v99。

        用户口径：「按退格键删除字符时，如果删除的是左右括号，或双引号，如果
        括号中或双引号中没内容，就能自动删除左右两端。」钩子侧的纯判据见
        `pair_delete_wanted` —— 那里已经按轮询快照筛过一遍，这里是**真正的
        权威**：再读一次真实行文本确认，因为快照最多是 100ms 前的（用户可能
        刚用鼠标把光标挪到了别处，也可能刚粘贴进一整段代码）。

        落点选【左符号原来的位置】：`f(|)` 删完是 `f`，光标就该在 `f` 与下
        一段之间 —— 也就是原生"删掉 `(`"之后的落点，两种做法观感一致，但
        一次按键就把 `()` 收干净。

        ⚠️ 这些情形一律返回 False（交给 VBE 原生退格，调用方会把这一下按键
        原样还给系统，绝不会出现"吞了按键却什么都没发生"）：
          * 有选区：退格是"删选区"，语义完全不同；
          * 左右不是一对、或对里有内容（`(a|)`）：那是普通删字；
          * 取不到 COM / 出任何异常。

        单行一次写：只在确认完毕之后调一次 ReplaceLine（绝不边读边写、也绝不
        在 finally 里补写）。落光标单独 try —— 写已经成了却因为落光标失败而
        返回 False，会让调用方再补一次退格、把 `)` 也删掉。

        ★列的处理走【显示列 -> 字符列】换算（★v103 更正：实际用的是
        `_bs_caret_char_col`，**不是** `_caret_char_col` —— 后者第三档是拿行尾
        列反推 tabw，而这里光标一定停在行中间，反推必错。钩子侧
        `pair_delete_wanted` 用的是**同一份**，两侧同口径），再把**字符
        列**喂给判据 —— 判据本身不猜列语义（见 `pair_delete_span` 里 v99 的
        教训）。摆回去时再按同一份语义换回显示列，与 `insert_pair` 的落点口径一致。
        之所以不像 v94b 那样交给 VBE 自己钳：那次是"落到行尾"（给一个超大的
        列让 VBE 夹），这次是"落到某个字符中间"，给不出那么干净的表达。
        """
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return False
            cp = vbe.ActiveCodePane
            if cp is None:
                return False
            cm = cp.CodeModule
            sl, sc, el, ec = cp.GetSelection()
            if int(sl) != int(el) or int(sc) != int(ec):
                return False                    # 有选区：那是"删选区"
            anchor = int(sl)
            if anchor <= 0:
                return False
            raw = str(cm.Lines(anchor, 1)).rstrip("\r\n")
            # ★权威复核：判据与钩子侧【同一份实现】（pair_delete_span），喂
            # 进去的是【此刻真读到的】行与【换算好的字符列】—— 快照最多是
            # 100ms 前的，中间用户可能已把光标点到别处，也可能刚粘贴进一整段
            # 代码。这里一票否决，调用方会把这一下退格原样还给系统，绝不误删。
            col, sem, tabw, wide2 = _bs_caret_char_col(raw, ec)
            span = pair_delete_span(raw, col)
            if not span:
                return False
            left, right = int(span[0]), int(span[1])
            new_line = raw[:left] + raw[right + 1:]
            if not self._write_line(cm, anchor, new_line):
                return False
            try:
                # 光标落在 left（0-based 字符下标）处。写回后读回真实行：
                # VBE 的自动语法检测可能把整行重排过（见 _locate_inserted 的
                # 说明），届时按压平前缀重新定位，认不出来才退到 left
                # （宁可偏一格，也不要越界或乱扔）。
                actual = cm.Lines(anchor, 1)
            except Exception:
                actual = new_line
            actual = ("" if actual is None else str(actual)).rstrip("\r\n")
            caret_off = left
            if actual != new_line:

                try:
                    off = _locate_by_prefix(actual, new_line, left)
                    if off is not None:
                        caret_off = off
                except Exception:
                    pass
            caret_off = max(0, min(int(caret_off), len(actual)))
            # ★摆回去用【本次换算出来的】tabw / wide2，不用写死的 (4, True)
            #   —— 与 insert_pair 同一套算法（MEMORY 第 19 条：列语义要在
            #   当下这一行现场量，不能吃缓存、也不能想当然）。
            if sem == "disp":
                out_col = _disp_width(actual, tabw, wide2, upto=caret_off) + 1
            else:
                out_col = caret_off + 1
            # ★代码复查修（2026-10-18）：本函数的 docstring 早就写着"落光标
            # 单独 try —— 写已经成了却因为落光标失败而返回 False，会让调用方
            # 再补一次退格、把 `)` 也删掉"，但代码里 SetSelection 仍在**外层**
            # try 内。补上这一道。
            try:
                cp.SetSelection(anchor, out_col, anchor, out_col)
            except Exception:
                _log("bspair: 已删一对、落光标失败（按成功返回）")
            return True
        except Exception:
            return False
        finally:
            _release_vbe_proxy()

    # ---- 按行移动光标（Shift+↑ / Shift+↓） ----
    def move_caret_line(self, delta):
        """把代码窗光标上（delta<0）/ 下（delta>0）移 delta 行，尽量保持同一列。

        【为什么不用"放行按键让 VBE 自己动光标"】VBE 里 Shift+↑ / Shift+↓ 的
        原生语义是**扩展选区**——按住 Shift 按一下会把上一行整行选进去，光标
        并没"干净地"上移一行。v78 用户口径是"跳出候选列表，并把光标上移 /
        下移一行"，所以这里由我们用 COM 直接落光标，那一对按键由调用方吞掉。

        列刻意沿用 GetSelection 给的列值原样透传：VBE 的列与 SetSelection 的列
        是同一套语义（显示列，Tab 展开），因此同一列值就是"同一视觉列"；目标
        行较短时由 VBE 自己钳到行尾（不在这里做字符数换算，避免 Tab 行算错）。

        成功返回 True；取不到 COM / 光标行无效 / 目标行越界 / 出任何异常都返回
        False —— 调用方据此把这一下按键原样还给系统，绝不吞掉用户的按键。
        """
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return False
            cp = vbe.ActiveCodePane
            if cp is None:
                return False
            cm = cp.CodeModule
            sl, sc, el, ec = cp.GetSelection()
            # 光标在选区的【活动端】：正向选择在 (el, ec)，反向选择在 (sl, sc)。
            if int(el) >= int(sl):
                line, col = int(el), int(ec)
            else:
                line, col = int(sl), int(sc)
            target = line + int(delta)
            if line <= 0 or target <= 0 or target > int(cm.CountOfLines):
                return False
            cp.SetSelection(target, col, target, col)
            return True
        except Exception:
            return False
        finally:
            # 与 new_line_below 同理：立刻放开手里的 COM 代理，别拖住 Excel 退出。
            _release_vbe_proxy()

    # ---- v110：Ctrl+↑ / Ctrl+↓ 跳到声明区末尾 / 过程末尾的空行 ----
    def ctrl_nav_jump(self, direction):
        """Ctrl+↑（direction<0）/ Ctrl+↓（direction>0）的真正执行者。

        用户口径（原话）："当我按 ctrl+↑，光标会自动定位到当前过程中最后一个
        dim 关键字行，如果此行的下一行不是空行，就自动在此行下方新增一行，
        光标移动到下方空白行；如果此行的下一行是空白行，就不用新增行……
        ctrl+↓ 同理，定位到 end sub 或 end function 行。"
        ★v110b 追加："ctrl+↓ 跳到结尾空行，能不能再加个 tab 补位？"
        ★v110c 改口径（Ctrl+↑）："光标落在本过程或函数的最后一个 dim 行的行尾，
        不用落在最后 dim 下一行并插入新行了。" ⇒ ↑ 不再插行（写入方式 `"eol"`）。

        判据与钩子侧是【同一份实现】：这里把此刻真读到的模块文本喂给纯函数
        `ctrl_nav_jump_plan`（v87 的规矩）。模块是整份读的，所以"当前过程"
        的边界以【真实代码】为准，与光标在过程里哪个位置无关。

        落点一律是目标行的【行尾】：带缩进的空白行上正好停在缩进之后，用户接着
        敲就是对的层级；真空行时列被 VBE 夹到 1；★v110c 的 Ctrl+↑ 目标行是 Dim
        语句那一行本身（`"eol"`，不动文本），光标就落在它的行尾。写完必须重新读
        一遍目标行再去落光标（`_caret_to_line_end` 自己会再读一次并让 VBE 夹列）。

        ★v110b/2（修真机 bug）：Ctrl+↓ 的缩进**直接写进新行**，缩进抄"尾行上方
        最后一条非空行"的过程体缩进。不再注入
        Tab 键（那条路真机不生效，见 `v110_end_jump_plan`）。计划给出**写入方式**：
          * `("insert", text)`：`InsertLines(target, text)`（原有行整体下移）；
          * `("swap", text)`：目标行是一条真空空白行，要换成带缩进的 —— 先
            `InsertLines(target, text)` 再 `DeleteLines(target + 1, 1)` 删掉原来
            那条真空行。**不能用 ReplaceLine**：VBE 会把写进去的行尾空白 trim 掉
            （v103b 教训），InsertLines 不会（"回车缩进" `new_line_below` 就是
            靠它写带缩进的新行，真机好用）；
          * `("eol", None)`：不动文本，只把光标落到目标行行尾（★v110c 的 Ctrl+↑）；
          * `(None, None)`：不动文本，只落光标。
        开关 `CTRL_NAV_TAB` 关掉时，`insert` / `swap` 退化成"只落光标 + 只保证
        有一行空白"（旧行为）；`"eol"` / `None` 本来就不写、不受它影响。

        ⚠️ 这些情形一律返回 False（调用方把这一下按键原样还给系统，绝不会
        出现"吞了按键却什么都没发生"）：
          * 取不到 COM / 没有活动代码窗格 / 模块行数为 0；
          * 纯判据说"不该接管"（不在过程里、没有 Dim、过程没收尾……）；
          * 锚点复核不过：`"insert"` / `"eol"` 要求目标行**非空白**（前者是参考
            行 / 尾行，后者是 Dim 语句行），其余（`"swap"` / None）要求目标行
            **空白**。不一致说明行号换算错了 —— 宁可这一次不跳，也绝不能把空白
            行插到用户代码中间（v21 的老规矩）。

        写绝不在 finally 里补写（v21）；`"swap"` 那两条写紧挨着发生、中间不碰
        光标也不判别的。
        """
        try:
            vbe = _get_vbe_cached()
            if vbe is None:
                return False
            cp = vbe.ActiveCodePane
            if cp is None:
                return False
            cm = cp.CodeModule
            sl, sc, el, ec = cp.GetSelection()
            # 光标在选区的【活动端】（与 move_caret_line 同一口径）：正向选择
            # 在 (el, ec)，反向在 (sl, sc)。
            if int(el) >= int(sl):
                line = int(el)
            else:
                line = int(sl)
            if line <= 0:
                return False
            try:
                n = int(cm.CountOfLines)
            except Exception:
                n = 0
            if n <= 0:
                return False
            code = cm.Lines(1, n)
            plan = ctrl_nav_jump_plan(code, line, direction)
            _log("ctrlnav: dir=%s caret=%s plan=%r" % (direction, line, plan))
            if plan is None:
                return False
            target, mode, text = plan
            # 开关关掉（VBECOMPLETE_NO_CTRL_NAV_TAB=1）：只保证"有一行空白"，
            # 不替用户补过程体缩进 —— insert 退化成写空串、swap 退化成不写。
            # `"eol"` / None 本来就不写文本，不受这个开关影响（★v110c 的 Ctrl+↑
            # 恒为 "eol"，所以 ↑ 在开关两侧行为完全一致）。
            if not CTRL_NAV_TAB and mode in ("insert", "swap"):
                if mode == "insert":
                    text = ""
                else:
                    mode = None
            # 锚点复核（见上面的说明）："insert" / "eol" 要求目标行非空白
            try:
                cur = str(cm.Lines(target, 1)).rstrip("\r\n")
            except Exception:
                return False
            if bool(cur.strip()) != (mode in ("insert", "eol")):
                _log("ctrlnav: 锚点复核不过 target=%s cur=%r mode=%s"
                     % (target, cur, mode))
                return False
            if mode == "insert":
                # InsertLines 在 target 处插入，原有行整体下移 —— 正是"在目标行
                # 处新起一行"。
                cm.InsertLines(target, text)
            elif mode == "swap":
                # 把一条真空空白行换成带缩进的一行：InsertLines 写空白行不会被
                # VBE trim、ReplaceLine 会（v103b），所以"插一条 + 删旧那条"。
                cm.InsertLines(target, text)
                cm.DeleteLines(target + 1, 1)
            # "eol" / None：一个字节都不写，下面直接落光标。
            try:
                actual = str(cm.Lines(target, 1)).rstrip("\r\n")
            except Exception:
                actual = ""
            _log("ctrlnav: mode=%s target=%s actual=%r" % (mode, target, actual))
            if not _caret_to_line_end(cp, cm, target, actual):
                # 兜底：该编辑器不夹列 -> 才回去按列语义算（见 _caret_to_line_end）
                _c = _eol_caret_col(actual, _sem_cache.get("info"))
                cp.SetSelection(target, _c, target, _c)
            return True
        except Exception:
            return False
        finally:
            _release_vbe_proxy()
