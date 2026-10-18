"""VBA 内建函数的【形参签名】种子（★v125）。

## 为什么要单独一个文件

`vba_builtins.py` 只管"有哪些名字"（关键字 / 函数 / 类型 / 常量），本文件只管
"这些函数的参数长什么样"。两者用途完全不同：

  * 名字进候选池，**每个按键都要用**（热路径，必须在内存里是 frozenset）；
  * 签名只在"补完括号的那一瞬间"用一次（冷路径），而且**用户可能改**
    —— 所以它落在 `wordbank.db` 的 `signatures` 表里，本文件只是**首次建库的种子**。

这跟 v123 把名字搬进 SQLite 是同一条思路：语言自带的东西要能被用户手改
（`MsgBox` 的五个参数太长，很多人只想看前两个）。改库之后重启即可生效，
`signatures` 表里没有的名字会退回本文件。

## 数据格式

    (函数名, "括号内原文")

  * **只存括号里面那段**，不存名字 —— 显示时拼 `名字(参数)`，名字就是主键，
    存两遍只会分叉（改名时忘了改另一半）。
  * 可选参数用 `[ ]` 包起来，跟 VBA 文档的写法一致（`Mid(string, start,
    [length])`）。
  * 没有参数的函数（`Now` / `Timer` / `DoEvents`）第二项就是空串。
  * 参数**名**优先照 VBA 官方文档写（`string` / `length` / `expression`），
    这样用户在编辑器里看到一个不认识的形参名时能对上文档。

## 维护口径

  * **只收"能当表达式用"的函数**（与 `vba_builtins.FUNCTIONS` 同一批范围），
    不收语句（`Erase` / `Randomize` / `Kill`）、不收对象成员
    （`Range.Cells` 那种由 VBE 自己管）。
  * 名字与 `vba_builtins.FUNCTIONS` 里的写法保持一致（大小写不影响匹配，
    匹配一律 `.lower()`）。
  * **不追求覆盖全部 300+ 个**：优先"写代码天天见"的那批。宁可少一条
    （最坏是那一瞬间不弹签名），也不写一条错的（错误签名比没有更误导人）。
"""

# ---------------------------------------------------------------------------
# 字符串 / 文本
# ---------------------------------------------------------------------------
_STRING = (
    ("Asc", "string"),
    ("AscW", "string"),
    ("Chr", "charcode"),
    ("ChrW", "charcode"),
    ("Format", "expression, [format], [firstdayofweek], [firstweekofyear]"),
    ("FormatCurrency", "expression, [numdigitsafterdecimal], [includeleadingdigit],"
                       " [useparensfornegativenumbers], [groupdigits]"),
    ("FormatDateTime", "date, [namedformat]"),
    ("FormatNumber", "expression, [numdigitsafterdecimal], [includeleadingdigit],"
                     " [useparensfornegativenumbers], [groupdigits]"),
    ("FormatPercent", "expression, [numdigitsafterdecimal], [includeleadingdigit],"
                      " [useparensfornegativenumbers], [groupdigits]"),
    ("InStr", "[start], string1, string2, [compare]"),
    ("InStrRev", "stringcheck, stringmatch, [start], [compare]"),
    ("Join", "sourcearray, [delimiter]"),
    ("LCase", "string"),
    ("UCase", "string"),
    ("Left", "string, length"),
    ("Right", "string, length"),
    ("Len", "expression"),
    ("LTrim", "string"),
    ("RTrim", "string"),
    ("Trim", "string"),
    ("Mid", "string, start, [length]"),
    ("MidB", "string, start, [length]"),
    ("Replace", "expression, find, replace, [start], [count], [compare]"),
    ("Space", "number"),
    ("String", "number, character"),
    ("Split", "expression, [delimiter], [limit], [compare]"),
    ("StrComp", "string1, string2, [compare]"),
    ("StrConv", "string, conversion, [lcid]"),
    ("StrReverse", "expression"),
    ("Str", "number"),
    ("Val", "string"),
)

# ---------------------------------------------------------------------------
# 数学 / 数值
# ---------------------------------------------------------------------------
_MATH = (
    ("Abs", "number"),
    ("Atn", "number"),
    ("Cos", "number"),
    ("Exp", "number"),
    ("Fix", "number"),
    ("Int", "number"),
    ("Log", "number"),
    ("Round", "expression, [numdecimalplaces]"),
    ("Rnd", "[number]"),
    ("Sgn", "number"),
    ("Sin", "number"),
    ("Sqr", "number"),
    ("Tan", "number"),
)

# ---------------------------------------------------------------------------
# 类型转换
# ---------------------------------------------------------------------------
_CONVERT = (
    ("CBool", "expression"),
    ("CByte", "expression"),
    ("CCur", "expression"),
    ("CDate", "expression"),
    ("CDbl", "expression"),
    ("CDec", "expression"),
    ("CInt", "expression"),
    ("CLng", "expression"),
    ("CLngLng", "expression"),
    ("CLngPtr", "expression"),
    ("CSng", "expression"),
    ("CStr", "expression"),
    ("CVar", "expression"),
    ("CVErr", "errornumber"),
    ("Hex", "number"),
    ("Oct", "number"),
)

# ---------------------------------------------------------------------------
# 日期 / 时间
# ---------------------------------------------------------------------------
_DATE = (
    ("Date", ""),
    ("DateAdd", "interval, number, date"),
    ("DateDiff", "interval, date1, date2, [firstdayofweek], [firstweekofyear]"),
    ("DatePart", "interval, date, [firstdayofweek], [firstweekofyear]"),
    ("DateSerial", "year, month, day"),
    ("DateValue", "date"),
    ("Day", "date"),
    ("Hour", "time"),
    ("Minute", "time"),
    ("Month", "date"),
    ("MonthName", "month, [abbreviate]"),
    ("Now", ""),
    ("Second", "time"),
    ("Time", ""),
    ("TimeSerial", "hour, minute, second"),
    ("TimeValue", "time"),
    ("Timer", ""),
    ("Weekday", "date, [firstdayofweek]"),
    ("WeekdayName", "weekday, [abbreviate], [firstdayofweek]"),
    ("Year", "date"),
)

# ---------------------------------------------------------------------------
# 数组
# ---------------------------------------------------------------------------
_ARRAY = (
    ("Array", "[arglist]"),
    ("Filter", "sourcearray, match, [include], [compare]"),
    ("IsArray", "varname"),
    ("LBound", "arrayname, [dimension]"),
    ("UBound", "arrayname, [dimension]"),
)

# ---------------------------------------------------------------------------
# 判断 / 取值
# ---------------------------------------------------------------------------
_INSPECT = (
    ("IsDate", "expression"),
    ("IsEmpty", "expression"),
    ("IsError", "expression"),
    ("IsMissing", "argname"),
    ("IsNull", "expression"),
    ("IsNumeric", "expression"),
    ("IsObject", "identifier"),
    ("TypeName", "varname"),
    ("VarType", "varname"),
    ("IIf", "expression, truepart, falsepart"),
    ("Choose", "index, choice1, [choice2], ..."),
    ("Switch", "expr1, value1, [expr2, value2], ..."),
    ("Partition", "number, start, stop, interval"),
)

# ---------------------------------------------------------------------------
# 交互 / 宿主
# ---------------------------------------------------------------------------
_INTERACT = (
    ("MsgBox", "prompt, [buttons], [title], [helpfile], [context]"),
    ("InputBox", "prompt, [title], [default], [xpos], [ypos], [helpfile], [context]"),
    ("Shell", "pathname, [windowstyle]"),
    ("DoEvents", ""),
    ("CreateObject", "class, [servername]"),
    ("GetObject", "[pathname], [class]"),
    ("CallByName", "object, procname, calltype, [args]"),
    ("Environ", "[envstring]"),
    ("Command", ""),
    ("RGB", "red, green, blue"),
    ("QBColor", "color"),
    ("LoadPicture", "[filename]"),
    ("Error", "[errornumber]"),
)

# ---------------------------------------------------------------------------
# 文件 / 目录（这些在 VBA 里既是函数也是语句，签名按"表达式用法"给）
# ---------------------------------------------------------------------------
_FILE = (
    ("Dir", "[pathname], [attributes]"),
    ("FileLen", "pathname"),
    ("FileDateTime", "pathname"),
    ("GetAttr", "pathname"),
    ("FreeFile", "[rangenumber]"),
    ("EOF", "filenumber"),
    ("LOF", "filenumber"),
    ("Loc", "filenumber"),
    ("Seek", "filenumber"),
    ("Input", "number, filenumber"),
    ("CurDir", "[drive]"),
)

# ---------------------------------------------------------------------------
# 财务（做报表的人偶尔要用；参数名照文档缩写）
# ---------------------------------------------------------------------------
_FINANCE = (
    ("FV", "rate, nper, pmt, [pv], [type]"),
    ("IPmt", "rate, per, nper, pv, [fv], [type]"),
    ("NPer", "rate, pmt, pv, [fv], [type]"),
    ("NPV", "rate, values()"),
    ("Pmt", "rate, nper, pv, [fv], [type]"),
    ("PPmt", "rate, per, nper, pv, [fv], [type]"),
    ("PV", "rate, nper, pmt, [fv], [type]"),
    ("Rate", "nper, pmt, pv, [fv], [type], [guess]"),
    ("SLN", "cost, salvage, life"),
    ("SYD", "cost, salvage, life, period"),
)

#: 全部内建签名种子。顺序 = 上面的分组顺序（建库时按这个顺序插，`list` 看着顺眼）。
SIGNATURES = tuple(
    _STRING + _MATH + _CONVERT + _DATE + _ARRAY + _INSPECT
    + _INTERACT + _FILE + _FINANCE
)


def seed_rows():
    """给 `wordbank.ensure()` 用：[(name, sig), ...]（已去重、保序）。"""
    out = []
    seen = set()
    for name, sig in SIGNATURES:
        key = str(name).lower()
        if not key or key in seen:
            continue
        seen.add(key)
        out.append((str(name), str(sig or "")))
    return out


def lookup(name):
    """按名字在**种子里**查一份签名（用户改了库之后应以 `ACTIVE_SIGNATURES` 为准）。

    查不到返回 None —— 调用方据此判断"这个名字没有签名"，而不是"签名是空串"
    （后者是 `Now()` 这种真·无参函数）。
    """
    key = str(name or "").lower()
    if not key:
        return None
    for n, s in SIGNATURES:
        if n.lower() == key:
            return s
    return None
