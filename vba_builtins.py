"""VBA 语言自带的名字：内建函数、内建常量、内建数据类型、语言关键字。

这些名字【不在用户的 VBE 工程代码里】—— 它们由 VBA 语言运行时提供，代码文本里
一个字都找不到（除非用户恰好调用过 MsgBox / Left 之类）。所以必须像组件名、窗体
控件名那样，作为一路独立的来源喂给候选池（见 vbe_bridge._collect_identifiers）。

收录范围：
  * FUNCTIONS  —— VBA 库的内建过程与函数（类型转换 / 字符串 / 日期 / 数组 / 数学 /
                  文件 IO / 交互 / 注册表 / 财务），含 Open、Close、Print 这类以
                  语句形式出现的 VBA 内建过程；
  * CONSTANTS  —— 常用的 vb* 内建常量（换行符、MsgBox 参数与返回值、StrConv、
                  日期常量、VarType、文件属性、颜色等）；
  * TYPE_NAMES —— VBA 内建数据类型（`Dim x As <这里>` 用得上）；
  * KEYWORDS   —— 语言关键字 / 保留字（v62，按用户要求："把 vba 里面的所有
                  关键字都纳入到提示词里"）：过程与声明（Sub / Function / Dim /
                  Set …）、流程控制（If / For / Select …）、字面量（True /
                  Nothing …）、运算符（And / Mod …）。

【v61 与 v62 的口径变化】v61 时关键字是刻意不收的（用户当时讨厌 If/For/Sub 这类
名字混进候选池）；v62 用户明确要求收进来，于是单列 KEYWORDS 一组，并配独立的开关
`vbe_bridge.ENABLE_VBA_KEYWORDS` —— 想关掉只关这一组，不影响内建函数/常量。

【v70 的口径变化】用户要求"**内置枚举（vb*）不要提示，内建函数保留**"——
理由是内置枚举他基本不用，提示出来只是干扰。于是 CONSTANTS 与 FUNCTIONS 拆成两组
独立导出（BUILTIN_CONSTANTS / BUILTIN_FUNCTIONS），收不收由
`vbe_bridge.ENABLE_VBA_CONSTANTS`（**默认 False**）决定；同一口径下宿主类型库的
枚举常量（xl* / mso*）也默认不收（`vbe_bridge.ENABLE_HOST_ENUMS` 默认 False）。
⚠️ 本模块的清单本身【不裁剪】，开关只在 vbe_bridge 的收集段生效 ——
这样测试 / 诊断仍能看到完整名单，也方便用户随时把开关打开。

【v77 的口径变化】用户接着说"**我觉得提示词太多了……vba 本身带的关键字、函数这块
都删掉吧。保留能提示变量名、自定义的函数/过程名、窗体/控件名、模块名**"——
于是 FUNCTIONS / TYPE_NAMES / KEYWORDS 也改成**默认不收**
（`vbe_bridge.ENABLE_VBA_BUILTINS` / `ENABLE_VBA_KEYWORDS` 均默认 False）。
默认口径下候选池只剩"这个工程里真实存在的东西"。
⚠️ 同样**一个字都不裁**（清单保持完整、收集门都还在），四个开关彼此独立，
想收哪一批就 `set VBECOMPLETE_xxx=1`。见回归测试第 54 节。

刻意【不收】的：
  * 过时的 Def* 类型声明关键字（DefBool / DefInt / … / DefVar，11 个）—— 早在
    VB6 时代就已废弃，实际代码里几乎不会手写；
  * 歧义过大的语句名（Name / Width / Line / Spc / Tab）—— 它们在代码里更常作为
    变量名或属性名出现，提示出来容易误导（v61 起就没有收录，v62 也没放回来）；
  * Option Base / Option Compare 的参数名（Base / Compare / Binary / Text /
    Database）—— 不是保留字，且都是极常见的普通变量名；
  * vbKey* 键盘常量（约 70 个）—— 前缀高度集中、实际极少手写；
  * Excel / Office 对象模型（Application / Worksheets / Range …）—— 那是宿主库、
    不是 VBA 自带，且 VBE 自己会提示点号后的成员。

清单按 VBA 文档 + VBE 对象浏览器（VBA 库）整理；名字用 VBA 惯例的大小写（插入
到代码里就是这个形式）。全部成员都是合法标识符，无需再做正则校验。
"""

# --------------------------------------------------------------------------
# 内建函数 / 过程（VBA 库成员）
# --------------------------------------------------------------------------
FUNCTIONS = (
    # ---- 类型转换与判定 ----
    "CBool", "CByte", "CCur", "CDate", "CDbl", "CDec", "CInt", "CLng",
    "CLngLng", "CLngPtr", "CSng", "CStr", "CVar", "CVErr",
    "IsArray", "IsDate", "IsEmpty", "IsError", "IsMissing", "IsNull",
    "IsNumeric", "IsObject",
    "TypeName", "VarType",

    # ---- 字符串 ----
    "Asc", "AscB", "AscW", "Chr", "ChrB", "ChrW",
    "Format", "FormatCurrency", "FormatDateTime", "FormatNumber",
    "FormatPercent",
    "Hex", "Oct", "InStr", "InStrRev", "Join", "LCase", "Left", "LeftB",
    "Len", "LenB", "LTrim", "Mid", "MidB", "Replace", "Right", "RightB",
    "RTrim", "Space", "Split", "Str", "StrComp", "StrConv", "String",
    "StrReverse", "Trim", "UCase", "Val", "Filter",

    # ---- 数学 ----
    "Abs", "Atn", "Cos", "Exp", "Fix", "Int", "Log", "Rnd", "Round", "Sgn",
    "Sin", "Sqr", "Tan",

    # ---- 日期与时间 ----
    "Date", "DateAdd", "DateDiff", "DatePart", "DateSerial", "DateValue",
    "Day", "Hour", "Minute", "Month", "MonthName", "Now", "Second", "Time",
    "Timer", "TimeSerial", "TimeValue", "Weekday", "WeekdayName", "Year",

    # ---- 数组 ----
    "Array", "LBound", "UBound",

    # ---- 交互 / 宿主 ----
    "MsgBox", "InputBox", "DoEvents", "Shell", "AppActivate", "SendKeys",
    "Command", "Environ", "CurDir",

    # ---- 对象 ----
    "CallByName", "Choose", "CreateObject", "GetObject", "IIf", "Partition",
    "QBColor", "RGB", "Switch",

    # ---- 错误处理 ----
    "Error",

    # ---- 文件与输入输出 ----
    "Dir", "EOF", "FileAttr", "FileDateTime", "FileLen", "FreeFile",
    "GetAttr", "SetAttr", "FileCopy", "Kill", "MkDir", "RmDir", "ChDir",
    "ChDrive",
    "Open", "Close", "Get", "Put", "Input", "Seek", "Lock", "Unlock",
    "Loc", "LOF", "Print", "Write", "Reset", "LSet", "RSet",
    "LoadPicture", "SavePicture",

    # ---- 注册表设置 ----
    "SaveSetting", "GetSetting", "GetAllSettings", "DeleteSetting",

    # ---- 财务 ----
    "DDB", "FV", "IPmt", "IRR", "MIRR", "NPer", "NPV", "Pmt", "PPmt", "PV",
    "Rate", "SLN", "SYD",

    # ---- 窗体（VBA 库内建过程）----
    "Load", "Unload",

    # ---- 其它内建过程（语句形式）----
    "Beep", "Erase", "Randomize",
)

# --------------------------------------------------------------------------
# 常用内建常量（VBA 库常量）
# --------------------------------------------------------------------------
_CONST_STRINGS = (
    "vbCr", "vbCrLf", "vbLf", "vbNewLine", "vbNullChar", "vbNullString",
    "vbTab", "vbBack", "vbFormFeed", "vbVerticalTab",
)
_CONST_MSGBOX = (
    # 按钮组合
    "vbOKOnly", "vbOKCancel", "vbAbortRetryIgnore", "vbYesNoCancel",
    "vbYesNo", "vbRetryCancel",
    # 图标
    "vbCritical", "vbQuestion", "vbExclamation", "vbInformation",
    # 默认按钮
    "vbDefaultButton1", "vbDefaultButton2", "vbDefaultButton3",
    "vbDefaultButton4",
    # 模态 / 外观
    "vbApplicationModal", "vbSystemModal", "vbMsgBoxHelpButton",
    "vbMsgBoxSetForeground", "vbMsgBoxRight", "vbMsgBoxRtlReading",
    # 返回值
    "vbOK", "vbCancel", "vbAbort", "vbRetry", "vbIgnore", "vbYes", "vbNo",
)
_CONST_COMPARE = (
    "vbBinaryCompare", "vbTextCompare", "vbDatabaseCompare",
    "vbUpperCase", "vbLowerCase", "vbProperCase",
    "vbWide", "vbNarrow", "vbKatakana", "vbHiragana",
    "vbUnicode", "vbFromUnicode",
)
_CONST_DATE = (
    "vbGeneralDate", "vbLongDate", "vbShortDate", "vbLongTime", "vbShortTime",
    "vbSunday", "vbMonday", "vbTuesday", "vbWednesday", "vbThursday",
    "vbFriday", "vbSaturday", "vbUseSystemDayOfWeek",
    "vbUseSystem", "vbFirstJan1", "vbFirstFourDays", "vbFirstFullWeek",
)
_CONST_VARTYPE = (
    "vbEmpty", "vbNull", "vbInteger", "vbLong", "vbSingle", "vbDouble",
    "vbCurrency", "vbDate", "vbString", "vbObject", "vbError", "vbBoolean",
    "vbVariant", "vbDataObject", "vbDecimal", "vbByte", "vbLongLong",
    "vbUserDefinedType", "vbArray",
)
_CONST_FILE = (
    "vbNormal", "vbReadOnly", "vbHidden", "vbSystem", "vbVolume",
    "vbDirectory", "vbArchive", "vbAlias",
)
_CONST_COLOR = (
    "vbBlack", "vbRed", "vbGreen", "vbYellow", "vbBlue", "vbMagenta",
    "vbCyan", "vbWhite",
)
_CONST_MISC = (
    "vbObjectError",
)

CONSTANTS = (_CONST_STRINGS + _CONST_MSGBOX + _CONST_COMPARE + _CONST_DATE
             + _CONST_VARTYPE + _CONST_FILE + _CONST_COLOR + _CONST_MISC)

# --------------------------------------------------------------------------
# 内建数据类型（`Dim x As <这里>`）
# --------------------------------------------------------------------------
TYPE_NAMES = (
    "Boolean", "Byte", "Currency", "Date", "Double", "Integer", "Long",
    "LongLong", "LongPtr", "Object", "Single", "String", "Variant",
)

# --------------------------------------------------------------------------
# 语言关键字 / 保留字（v62）
# --------------------------------------------------------------------------
# 按 VBA 官方关键字表整理，另补三个常写的（PtrSafe / Alias / Explicit）：
#   * PtrSafe / Alias —— 64 位 Declare 的修饰关键字，写 API 声明时要用；
#   * Explicit       —— Option Explicit（Excel 新模块默认带上这一行）。
# 与上面几组重合的（String / Date / Error / Erase / Get / Open …）由
# BUILTIN_KEYWORDS 的去重挡掉，不会重复进池。
KEYWORDS = (
    # ---- 过程、声明与可见性 ----
    "Sub", "Function", "Property", "Declare", "PtrSafe", "Alias", "Lib",
    "Public", "Private", "Friend", "Global", "Static", "Const", "Dim",
    "ReDim", "Preserve", "Implements", "Event", "WithEvents", "RaiseEvent",
    "Optional", "ParamArray", "ByVal", "ByRef", "Attribute", "Option",
    "Explicit", "Enum", "Type", "As", "New", "Me", "Is", "Like",
    "Set", "Let", "Call",

    # ---- 流程控制 ----
    "If", "Then", "Else", "ElseIf", "End", "Select", "Case", "For", "Each",
    "In", "To", "Step", "Next", "Do", "Loop", "While", "Wend", "Until",
    "Exit", "GoTo", "GoSub", "Return", "On", "Resume", "Stop", "With",
    "TypeOf", "Rem", "Debug",

    # ---- 字面量 / 空值 ----
    "True", "False", "Nothing", "Empty", "Null",

    # ---- 运算符 ----
    "And", "Or", "Not", "Xor", "Eqv", "Imp", "Mod", "AddressOf",
)

# --------------------------------------------------------------------------
# 对外导出
# --------------------------------------------------------------------------
# v70：函数与常量【分开】导出 —— 用户口径是"内置枚举常量不收，内建函数照常收"。
# 只拆导出、不动清单本身（见文件头 v70 说明）。
BUILTIN_FUNCTIONS = tuple(dict.fromkeys(FUNCTIONS))
BUILTIN_CONSTANTS = tuple(dict.fromkeys(CONSTANTS))

# 两组并集（静态全量清单）：保留给"想知道完整名单"的调用方（测试 / 诊断）。
# ⚠️ 收集侧**不要**直接用它 —— 走 vbe_bridge 的 ENABLE_VBA_CONSTANTS 裁剪。
BUILTIN_NAMES = tuple(dict.fromkeys(FUNCTIONS + CONSTANTS))

# 内建数据类型名（同样进候选池，供 `As |` 位置使用）。
BUILTIN_TYPE_NAMES = tuple(dict.fromkeys(TYPE_NAMES))

# 语言关键字（单独一组：开关独立、也便于测试里单独取舍）。
# 与 BUILTIN_NAMES 的同名项（Error / Erase / Get / Open …）由收集侧的统一去重挡掉。
BUILTIN_KEYWORDS = tuple(dict.fromkeys(KEYWORDS))

# 去重后的小写集合，供需要快速查询的调用方使用。
BUILTIN_LOWER = frozenset(n.lower() for n in BUILTIN_NAMES)
BUILTIN_TYPE_LOWER = frozenset(n.lower() for n in BUILTIN_TYPE_NAMES)
BUILTIN_KEYWORD_LOWER = frozenset(n.lower() for n in BUILTIN_KEYWORDS)

# --------------------------------------------------------------------------
# ★v103b：哪些关键字在补全后要跟一个【空格】
# --------------------------------------------------------------------------
# 用户口径（原话）："在使用 public，private，sub，function，set，if，for，
# with，select，end，exit 等这类关键词时，后面需要有空格，要在按 tab 键后，
# 也能自动补全空格，idea 里就有这个功能。"
#
# 实现用【排除法】：语言关键字里绝大多数后面必然还要跟一个词（过程名 / 变量名 /
# 条件表达式 / 操作数），补一个空格总是对的；只有少数几种情况不该补，逐个排除：
#   * 经常是【整行的最后一个词】：Else / Next / Loop / Wend / Then / Resume /
#     Stop / Return / Explicit / Do —— 补了会在行尾留一个空格（VBE 还不 trim）；
#   * 后面紧跟的是【点号】而不是空格：Me / Debug（`Me.Value` / `Debug.Print`）；
#   * 本身就是【完整表达式】：True / False / Nothing / Empty / Null。
# 排除法（而不是白名单）的好处：关键字表将来加词时，新词自动获得补空格能力，
# 只有真正不适合的才需要显式排除。
KEYWORDS_NO_TRAIL_SPACE = frozenset((
    "explicit", "me", "debug", "then", "else", "next", "do", "loop", "wend",
    "return", "resume", "stop", "true", "false", "nothing", "empty", "null",
))

# 会被补空格的关键字（小写），预先算好，查询 O(1)。
BUILTIN_KEYWORD_NEEDS_SPACE = frozenset(
    BUILTIN_KEYWORD_LOWER - KEYWORDS_NO_TRAIL_SPACE)


def keyword_wants_space(word):
    """这个补全词是不是"后面要跟一个空格"的 VBA 关键字（纯函数）。★v103b。

    只认【语言关键字】这一组；工程里的过程名 / 变量名 / 组件名一律 False
    —— 变量后面多一个空格纯属噪音。用户自建的 Sub / Function 也不在这里
    （它们走"补括号"那条路，由 engine._is_callable_name 判）。

    大小写不敏感（候选窗里给的是关键字原文，如 `Public`）。

    ★v123：查的是【生效】集合 `ACTIVE_KEYWORD_SPACE`（词库里的 `wants_space`
    列可以逐条覆盖），不再是写死的 `BUILTIN_KEYWORD_NEEDS_SPACE`。没有词库 /
    词库不可用时两者**逐字相同**，行为不变。
    """
    return str(word or "").strip().lower() in ACTIVE_KEYWORD_SPACE


# ---------------------------------------------------------------------------
# ★v123：生效快照 —— 词库（SQLite）里加载出来的那一份
# ---------------------------------------------------------------------------
# 为什么要多一层 `ACTIVE_*` 而不是直接让 `BUILTIN_*` 变成"可变的"：
#
#   `BUILTIN_*` 现在是**冻结的种子清单**，被二十多处测试当"代码里带的那份名单"
#   来体检（`len(...) >= 150`、`MsgBox in ...`、`FUNCTIONS | CONSTANTS == NAMES`）。
#   如果它随用户手改的词库变化，这些用例就会"因为用户配置而红/绿"——那是测试网
#   最不该有的性质。
#
#   所以职责拆开：
#     `BUILTIN_*`  代码里带的那份（永不变化，测试和文档引用它）
#     `ACTIVE_*`   此刻真正生效的（词库优先，失败退回种子）
#   收集侧与引擎侧一律读 `ACTIVE_*`；`load_active()` 之前它们**就是**种子。
#
# ⚠️ 本模块顶层**仍然零 import**（sqlite3 只在 `load_active()` 内部惰性引入），
#    这样 import vba_builtins 的成本与 v122 完全一样，热路径不受任何影响。

ACTIVE_FUNCTIONS = BUILTIN_FUNCTIONS
ACTIVE_CONSTANTS = BUILTIN_CONSTANTS
ACTIVE_TYPE_NAMES = BUILTIN_TYPE_NAMES
ACTIVE_KEYWORDS = BUILTIN_KEYWORDS
ACTIVE_FUNCTION_LOWER = frozenset(n.lower() for n in ACTIVE_FUNCTIONS)
ACTIVE_TYPE_LOWER = frozenset(n.lower() for n in ACTIVE_TYPE_NAMES)
ACTIVE_KEYWORD_LOWER = frozenset(n.lower() for n in ACTIVE_KEYWORDS)
ACTIVE_KEYWORD_SPACE = frozenset(BUILTIN_KEYWORD_NEEDS_SPACE)

# "语言自带"那一批（小写）。收集侧用它决定要不要把某个名字算进
# `builtin_names` —— 那个集合专供引擎收紧 2~3 字符的跳步命中（BUILTIN_SCATTER_MIN）。
# 用户自己加进词库的词 (is_builtin=0) 刻意**不**进这里：他是特意加的，
# 不该被"公共词汇"那套收紧规则连坐。
ACTIVE_BUILTIN_LOWER = frozenset(
    n.lower() for n in (BUILTIN_FUNCTIONS + BUILTIN_CONSTANTS
                        + BUILTIN_TYPE_NAMES + BUILTIN_KEYWORDS))

# 词库状态（`load_active()` 之后才有意义，给自检 / 排错 / 测试看）
WORDBANK_INFO = {"loaded": False, "path": "", "counts": {}, "error": "未尝试加载"}

# ★v125：内建函数的【形参签名】—— {小写函数名: 括号内原文}。
#   与 ACTIVE_FUNCTIONS 同源（同一个库、同一次加载），但**不作为候选池使用**：
#   它只在"补完括号那一瞬间"被查一次（冷路径）。
#   查不到 = 没有签名；空串 = 真·无参函数（`Now()`）——两者必须能分开。
#   来源优先级：词库 signatures 表 > `vba_signatures.py` 种子。
ACTIVE_SIGNATURES = {}


def _seed_signatures():
    """`vba_signatures.py` 里的种子（惰性导入，顶层仍然零 import）。"""
    try:
        import vba_signatures
        return {str(n).lower(): str(s or "") for n, s in vba_signatures.SIGNATURES}
    except Exception:
        return {}


def seed_rows():
    """种子行：首次建库时写进 SQLite 的那批。-> [(name, kind, is_b, en, ws, note)]

    `wants_space` 一律 None —— "NULL 是什么意思"（不在排除表里就补空格）是
    VBA 语义，由 `load_active()` 解释，不让词库层去猜。
    """
    rows = []
    seen = set()
    for names, kind in ((BUILTIN_FUNCTIONS, "function"),
                        (BUILTIN_CONSTANTS, "constant"),
                        (BUILTIN_TYPE_NAMES, "type"),
                        (BUILTIN_KEYWORDS, "keyword")):
        for n in names:
            k = (n.lower(), kind)
            if k in seen:
                continue
            seen.add(k)
            rows.append((n, kind, 1, 1, None, ""))
    return rows


def load_active(force=False):
    """从词库加载生效快照。-> True = 用上了库；False = 退回种子。

    只应在**启动时调用一次**（`main.py`）。任何失败都不抛异常：`ACTIVE_*` 保持
    种子、`WORDBANK_INFO["error"]` 记原因，程序照常跑 —— 这是本模块对外的契约。
    """
    global ACTIVE_FUNCTIONS, ACTIVE_CONSTANTS, ACTIVE_TYPE_NAMES, ACTIVE_KEYWORDS
    global ACTIVE_FUNCTION_LOWER, ACTIVE_TYPE_LOWER, ACTIVE_KEYWORD_LOWER
    global ACTIVE_KEYWORD_SPACE, ACTIVE_BUILTIN_LOWER, ACTIVE_SIGNATURES
    # 先把签名按种子铺好 —— 后面无论走哪条分支，它至少是可用的（v125）
    ACTIVE_SIGNATURES = _seed_signatures()
    try:
        import wordbank
    except Exception as e:                           # pragma: no cover
        WORDBANK_INFO.update({"loaded": False, "error": "导入 wordbank 失败: %s" % e})
        return False
    try:
        snap = wordbank.load()
    except Exception as e:                           # 契约要求：不抛
        WORDBANK_INFO.update({"loaded": False, "error": "读库异常: %s" % e})
        return False
    if snap is None:
        WORDBANK_INFO.update({"loaded": False, "path": wordbank.db_path(),
                              "error": wordbank.last_error() or "未知原因"})
        return False

    fn = [r[0] for r in snap.by_kind("function")]
    co = [r[0] for r in snap.by_kind("constant")]
    ty = [r[0] for r in snap.by_kind("type")]
    kw = snap.by_kind("keyword")
    if not (fn or co or ty or kw):
        WORDBANK_INFO.update({"loaded": False, "path": snap.path,
                              "error": "库里一条启用的词都没有"})
        return False

    # 空格三态：NULL -> 默认排除法；1 / 0 -> 用户显式指定（覆盖排除法）
    space = set()
    for name, _is_b, ws in kw:
        low = name.lower()
        if ws is None:
            if low not in KEYWORDS_NO_TRAIL_SPACE:
                space.add(low)
        elif ws:
            space.add(low)

    ACTIVE_FUNCTIONS = tuple(fn)
    ACTIVE_CONSTANTS = tuple(co)
    ACTIVE_TYPE_NAMES = tuple(ty)
    ACTIVE_KEYWORDS = tuple(r[0] for r in kw)
    ACTIVE_FUNCTION_LOWER = frozenset(n.lower() for n in ACTIVE_FUNCTIONS)
    ACTIVE_TYPE_LOWER = frozenset(n.lower() for n in ACTIVE_TYPE_NAMES)
    ACTIVE_KEYWORD_LOWER = frozenset(n.lower() for n in ACTIVE_KEYWORDS)
    ACTIVE_KEYWORD_SPACE = frozenset(space)
    ACTIVE_BUILTIN_LOWER = frozenset(
        r[0].lower() for r in snap.rows if r[2])
    # ★v125：签名以**库为准**，库里没有的才用种子兜底。
    # 为什么方向是"库覆盖种子"而不是反过来：用户改签名（比如把 `MsgBox` 的
    # 五个参数缩成两个）就是想让自己的版本生效；而种子只是"没得选时的默认"。
    _sigs = _seed_signatures()
    try:
        _sigs.update(dict(getattr(snap, "sigs", None) or {}))
    except Exception:
        pass
    ACTIVE_SIGNATURES = _sigs
    WORDBANK_INFO.update({"loaded": True, "path": snap.path,
                          "counts": dict(snap.counts), "error": "",
                          "sigs": len(ACTIVE_SIGNATURES)})
    return True


def reset_active():
    """退回种子（`NO_WORDBANK=1` 或测试用）。"""
    global ACTIVE_FUNCTIONS, ACTIVE_CONSTANTS, ACTIVE_TYPE_NAMES, ACTIVE_KEYWORDS
    global ACTIVE_FUNCTION_LOWER, ACTIVE_TYPE_LOWER, ACTIVE_KEYWORD_LOWER
    global ACTIVE_KEYWORD_SPACE, ACTIVE_BUILTIN_LOWER, ACTIVE_SIGNATURES
    ACTIVE_SIGNATURES = _seed_signatures()
    ACTIVE_FUNCTIONS = BUILTIN_FUNCTIONS
    ACTIVE_CONSTANTS = BUILTIN_CONSTANTS
    ACTIVE_TYPE_NAMES = BUILTIN_TYPE_NAMES
    ACTIVE_KEYWORDS = BUILTIN_KEYWORDS
    ACTIVE_FUNCTION_LOWER = frozenset(n.lower() for n in ACTIVE_FUNCTIONS)
    ACTIVE_TYPE_LOWER = frozenset(n.lower() for n in ACTIVE_TYPE_NAMES)
    ACTIVE_KEYWORD_LOWER = frozenset(n.lower() for n in ACTIVE_KEYWORDS)
    ACTIVE_KEYWORD_SPACE = frozenset(BUILTIN_KEYWORD_NEEDS_SPACE)
    ACTIVE_BUILTIN_LOWER = frozenset(
        n.lower() for n in (BUILTIN_FUNCTIONS + BUILTIN_CONSTANTS
                            + BUILTIN_TYPE_NAMES + BUILTIN_KEYWORDS))
    WORDBANK_INFO.update({"loaded": False, "counts": {}, "error": "已退回种子"})
