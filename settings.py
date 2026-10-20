"""配置中心（★v120）：把散在几十个环境变量里的开关收进 `settings.ini`。

## 为什么要有这一层

v119 之前，全部开关都是"`set VBECOMPLETE_XXX=1` 然后重启"。问题有三个：

  1. **没人记得住** —— 40 多个名字，光靠文档翻；
  2. **改一次要开命令行** —— 双击 `run.bat` 启动的用户根本没有 `set` 的机会；
  3. **改了不知道有没有生效** —— 环境变量写错了（`=ture`）会静默回落到默认值，
     用户以为开了，其实没开。

## 优先级（第一个"能认出来的"赢）

    环境变量  >  settings.ini  >  代码里的默认值

★ 注意是"第一个**能认出来**的"，不是"第一个**存在**的"：
环境变量写了个认不出的值（`AUTO_SPACE=maybe`）时，会**继续去看 ini**，
而不是直接掉到代码默认值。理由 —— 用户既然在 ini 里明确写了 `AUTO_SPACE = 0`，
那显然比"代码里的默认值"更贴近他的意图；而"写错一个词就让功能整个失效"
正是这条链要避免的事（沿用 `vbe_bridge._env_flag` 的老契约：认不出的值
绝不改变行为，只是被跳过）。

## 文件长什么样

`settings.ini` **不是必须存在的**。第一次启动时 `ensure_file()` 会照着
`settings.example.ini`（仓库里那份，带全部说明）生成一份；没有 example 就
生成内置的简版模板。**里面每一行都是注释掉的** —— 于是"什么都不改"就等于
"全部用默认值"，不会因为生成了一份文件就把行为改掉。

    [settings]
    ; 2026-10-18 改名：这个开关管的是……
    ; AUTO_SPACE = 0
    NO_CTRL_NAV = 1

## 分层

本模块是**最底层**：只依赖标准库，**不 import 项目里任何模块**（`log.py` 依赖它，
它不能反过来依赖 `log`，否则循环 import）。出错一律静默 + 记在 `last_error()`。
"""

import os

# 项目根 = 本文件所在目录。★绝不用 os.getcwd()：工具由 run.bat / pythonw 拉起，
# cwd 会漂（v116 在 snippets.db 上踩过一次）。
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))

INI_NAME = "settings.ini"
EXAMPLE_NAME = "settings.example.ini"
# 用另一个 ini 文件（便携版 / 测试）：set VBECOMPLETE_SETTINGS=D:\path\my.ini
ENV_FILE = "VBECOMPLETE_SETTINGS"
SECTION = "settings"
PREFIX = "VBECOMPLETE_"

_TRUE = ("1", "true", "yes", "on")
_FALSE = ("0", "false", "no", "off")


# ---------------------------------------------------------------------------
# 开关总表：键名（不带 VBECOMPLETE_ 前缀） -> (默认值, 说明)
#
# 只用于两件事：① 生成 settings.example.ini / settings.ini 模板；② `python
# settings.py` 打印"当前实际生效的值"。**真正生效的默认值仍然写在各自的
# 调用点上**（`_env_flag("XXX", True)` 的那个 True）—— 这里若与代码不一致，
# 以代码为准，模板只是给人看的。
#
# 顺序即模板里的顺序，按"用户最可能想改"排前面。
# ---------------------------------------------------------------------------
KNOWN = (
    # ---- 日常最常改的 ----
    ("AUTO_SPACE", "1",
     "关键字后自动补空格（`End If` 敲完补一个空格）。0 = 关"),
    ("AUTO_PAREN", "1",
     "补全过程名后自动插 `()`。0 = 关"),
    ("PARAM_INFO", "0",
     "补完 `()` 后让 VBE 弹形参信息。**默认 0** —— 真机实测这个命令在补全"
     "场景下唤不出窗、只会响一声提示音；1 = 试一下（会响）"),
    ("NO_PARAM_INFO", "0",
     "强制关掉形参信息（优先级高于 PARAM_INFO）。1 = 关"),
    ("PARAM_INFO_KEY", "0",
     "拿不到 VBE 菜单命令时，退回注入 Ctrl+Shift+I。1 = 退回"
     "（★那条路也会让 VBE 响一声，默认不退回）"),
    ("NO_AUTOPAIR", "0",
     "自动配对括号 / 引号。1 = 关"),
    ("NO_IME_PUNCT", "0",
     "中文标点 / 全角时也按半角配对。1 = 强制半角"),
    ("NO_ENTER_INDENT", "0",
     "回车自动缩进。1 = 关（回车完全回到 VBE 原生）"),
    ("NO_AUTOCLOSE", "0",
     "自动补 `End Sub` / `Next` 等收尾。1 = 只做缩进、不补收尾"),
    ("NO_BRANCH_ALIGN", "0",
     "`Else` / `Case` 等分支的缩进对齐。1 = 关"),
    ("INDENT_UNIT", "",
     "一级缩进用什么：`2` / `4` / `tab`；留空 = 照工程里的既有缩进自动判断"),
    ("NO_BS_DELLINE", "0",
     "退格一次删掉整行空白。1 = 关"),
    ("NO_BS_DELPAIR", "0",
     "退格一次删掉成对的空括号 / 空引号。1 = 关"),
    ("HIDE_ON_COMPLETE", "1",
     "名字**打全**后自动收起候选窗（精确命中、且没有更长的候选时才收，像 IDEA）。"
     "0 = 一直挂在屏幕上"),
    ("REQUIRE_KEY_TRIGGER", "1",
     "自动弹窗**只认按键**：鼠标点击 / 右键菜单粘贴造成的文本变化只收起、不弹。"
     "0 = 旧行为（内容变就弹）"),
    # ---- 快捷键 ----
    ("NO_CTRL_NAV", "0",
     "`Ctrl+Shift+↑` / `Ctrl+Shift+↓` 跳转（最后一个 Dim / 收尾区）。1 = 关，按键交还 VBE"),
    ("NO_CTRL_NAV_TAB", "0",
     "`Ctrl+Shift+↓` 保留跳转但不补缩进。1 = 不补"),
    ("NO_SHIFT_NAV", "0",
     "`Shift+↑/↓` 跳出候选列表并上下移一行。1 = 关"),
    ("NO_SNIPPETS", "0",
     "`Ctrl+Q` 代码片段窗。1 = 关，按键交还 VBE"),
    ("NO_COMMENT", "0",
     "`Ctrl+/` 批量注释 / 取消注释。1 = 关，按键交还 VBE"),
    ("NO_INSERT_ABOVE", "0",
     "`Ctrl+Shift+Enter` 在当前行上方插入新行。1 = 关，按键交还 VBE"),
    ("NO_DUP_LINE", "0",
     "`Ctrl+D` 复制当前行（有选区 = 选中的几行）到下方。1 = 关"),
    ("NO_DEL_LINE", "0",
     "`Ctrl+K` 删除当前行（有选区 = 选中的几行）。1 = 关"),
    # ---- 候选内容 ----
    ("VBA_BUILTINS", "1",
     "候选里含 VBA 内建函数与内建数据类型（`MsgBox` / `String` …）。0 = 不含"),
    ("VBA_KEYWORDS", "1",
     "候选里含 VBA 语言关键字（`If` / `For` / `Dim` …）。0 = 不含"),
    ("VBA_CONSTANTS", "0",
     "候选里含 VBA 内建枚举常量（`vbCrLf` / `vbYes` …）。1 = 含"),
    ("HOST_ENUMS", "0",
     "候选里含宿主类型库枚举常量（`xlUp` / `msoTrue` …）。1 = 含"),
    ("HOST_ENUM_FAMILIES", "",
     "只留这些前缀的宿主枚举（逗号分隔，如 `xl`）。留空 = 自动推导"),
    ("IMPLICIT_HONOR_EXPLICIT", "1",
     "模块写了 `Option Explicit` 时不再收\"未声明的隐式变量\"。0 = 照收"),
    ("BUILTIN_SCATTER_MIN", "2",
     "2~3 字符输入时允许跳步命中的门槛。2 = 放宽（默认），4 = 旧口径"),
    ("NO_WORDBANK", "0",
     "1 = 不用词库（`wordbank.db`），一律用代码里的种子清单。词库改坏时用它排错"),
    # ---- 排序 ----
    ("NO_SCOPE_RANK", "0",
     "\"当前过程 / 本模块优先\"的作用域排序。1 = 关"),
    ("NO_RECENT_RANK", "0",
     "\"最近用过优先\"的排序。1 = 关"),
    ("RECENT_MAX", "64",
     "最近使用表最多记多少个名字"),
    ("NO_RECENT_PERSIST", "0",
     "最近使用表只留内存、不读也不写 `recent.json`。1 = 不落盘"),
    ("RECENT_FILE", "",
     "最近使用表路径。留空 = 项目文件夹下的 `recent.json`"),
    # ---- 让位 / 成员位置 ----
    ("NO_YIELD", "0",
     "让位给 VBE 自己的成员列表。1 = 永不让位"),
    ("YIELD_GRACE", "0.25",
     "让位宽限期（秒）：等这么久 VBE 没弹，就补上我们自己的候选"),
    ("YIELD_TTL", "1.5",
     "\"此处 VBE 不会弹列表\"标记的有效期（秒）；0 = 不记这个标记"),
    ("MEMBER_POPUP", "0",
     "`对象.` 之后也出我们的候选。1 = 出（默认整个让给 VBE）"),
    ("DECL_POPUP", "0",
     "声明位置（正在起新名字）也出候选。1 = 出"),
    # ---- 片段库 ----
    ("SNIPPET_DB", "",
     "片段库路径。留空 = 项目文件夹下的 `snippets.db`"),
    # ---- 词库 ----
    ("WORD_BANK", "",
     "词库（提示词：关键字 / 内建函数 / 类型 / 常量）路径。留空 = 项目文件夹下的 `wordbank.db`"),
    # ---- 性能 ----
    ("NO_MOD_CACHE", "0",
     "每个模块每轮全量重解析（v117 行为，慢约 8 倍）。1 = 关掉按模块缓存，只在 A/B 时用"),
    ("PROXY_TTL", "1.0",
     "VBE / 宿主的 COM 代理最长缓存时间（秒）；0 = 不缓存"),
    ("NO_SNAP_CACHE", "0",
     "★v130：关掉「光标没动就不读文本」的两级快照读取（回到每 100ms 全读一次，"
     "空闲时约多占 8% 的一个核）。1 = 关，只在排查「补全看不到刚改的代码」时用"),
    ("SNAP_MAX_AGE", "1.0",
     "★v130：快照最长复用多久（秒）。文本被**别的程序**改掉时光标不动，全靠这个"
     "上限发现；调小 = 更灵敏但更费 CPU，0 = 等于关掉缓存"),
    # ---- 窗口探测（只有类名对不上的宿主才要动）----
    ("NO_WINDOW_GATE", "0",
     "宿主窗口探测门。1 = 关（WPS 等类名不符的宿主用）"),
    ("NO_CODE_AREA_GATE", "0",
     "退回\"前台是 VBE 就算\"（不再要求焦点在代码窗格）。1 = 退回"),
    ("VBE_POPUP_CLASS", "",
     "追加 VBE 提示窗类名（逗号分隔）"),
    ("VBE_POPUP_RESCAN", "1.0",
     "重扫 VBE 提示窗的间隔（秒）"),
    ("VBE_YIELD_CLASS", "",
     "追加\"要让位\"的窗口类名（逗号分隔）"),
    ("VBE_FRAME_CLASS", "",
     "追加 VBE 主框架窗口类名（逗号分隔）"),
    ("NO_HOST_ENUMS", "0",
     "1 = 完全不加载宿主类型库（`HOST_ENUMS` 之外的另一道闸）"),
    # ---- 稳定性（★v128：自动重启）----
    ("AUTO_RESTART", "1",
     "自动重启：主线程被 COM 卡死、或进程被杀之后，守护进程 `watchdog.py` 负责"
     "重新拉起。0 = 关（回到 v127：卡死只能自己重启电脑）"),
    ("HANG_SEC", "30",
     "主循环多久没动就判定卡死（秒，最小 10）。判定后本进程硬退出、由守护进程重启"),
    # ---- 排障 ----
    ("LOG", "0",
     "写诊断日志。1 = 落盘（正常使用保持 0，关着时零开销）"),
    ("LOG_FILE", "",
     "日志文件路径。留空 = 项目文件夹下的 `VbeComplete.log`"),
    ("COFREE", "0",
     "每次用完调用 `CoFreeUnusedLibraries`（排查 COM 问题时才开）"),
)


# ---------------------------------------------------------------------------
# 读 ini（懒加载，全局只读一次）
# ---------------------------------------------------------------------------
_state = {"loaded": False, "data": {}, "path": "", "error": ""}


def path():
    """当前生效的 ini 路径（`VBECOMPLETE_SETTINGS` 可覆盖）。"""
    try:
        p = (os.environ.get(ENV_FILE) or "").strip()
    except Exception:
        p = ""
    return p or os.path.join(PROJECT_DIR, INI_NAME)


def reset():
    """丢掉缓存，下次取用时重新读盘（测试与"我刚改了 ini"时用）。

    ⚠️ 只能靠它让**已经求过值**的模块级常量重新读 —— 那些常量在 import 时就
    定型了，改 ini 之后仍然要重启。这个函数的作用是给测试和 `--check` 用。
    """
    _state.update({"loaded": False, "data": {}, "path": "", "error": ""})


def last_error():
    """上次读 ini 失败的原因（成功时是空串）。"""
    return _state.get("error") or ""


def load_ini():
    """读 ini，返回 {大写键名: 字符串值}。永不抛异常。"""
    if _state["loaded"]:
        return _state["data"]
    _state["loaded"] = True
    p = path()
    _state["path"] = p
    data = {}
    _state["error"] = ""
    try:
        if not os.path.isfile(p):
            _state["data"] = data
            return data
        import configparser
        # interpolation=None 是必须的：ini 里写 `VBE_POPUP_CLASS=a%b` 时，
        # 默认的 BasicInterpolation 会把它当占位符、直接抛异常 ⇒ 整份配置读不出来。
        # strict=False：同一个键写两遍时**后写的赢**，而不是整份配置作废
        #   （⚠️ 必须走构造参数 —— `cp.strict = False` 赋值是无效的：
        #    configparser 在 __init__ 里把它存成 `_strict`，属性赋值不影响解析）。
        cp = configparser.ConfigParser(interpolation=None, strict=False)
        # utf-8-sig：Windows 记事本存 UTF-8 时爱加 BOM，不加这个会把第一个键
        # 读成 "\ufeffAUTO_SPACE"。
        with open(p, encoding="utf-8-sig", errors="replace") as f:
            cp.read_file(f)
        for k, v in cp.defaults().items():
            data[str(k).strip().upper()] = str(v or "").strip()
        if cp.has_section(SECTION):
            for k, v in cp.items(SECTION):
                data[str(k).strip().upper()] = str(v or "").strip()
    except Exception as e:                      # 坏 ini 绝不能让程序起不来
        _state["error"] = "%s: %s" % (type(e).__name__, e)
        data = {}
    _state["data"] = data
    return data


def _norm(name):
    n = str(name or "").strip().upper()
    return n if n.startswith(PREFIX) else PREFIX + n


def candidates(name):
    """按优先级给出 [(值, 来源), ...]，来源是 `"env"` / `"ini"`。

    空值（空串）不算数 —— 当作"这一层没设置"。这是 Windows 上很常见的坑：
    `set VBECOMPLETE_NO_AUTOPAIR=` 会留下一个空串，它应该等同于没设。
    """
    out = []
    n = _norm(name)
    v = os.environ.get(n)
    if v is not None and v.strip():
        out.append((v, "env"))
    data = load_ini()
    short = n[len(PREFIX):]
    for key in (n, short):
        if key in data and data[key]:
            out.append((data[key], "ini"))
            break
    return out


def raw(name):
    """第一个非空字符串值；没有就返回空串。"""
    c = candidates(name)
    return c[0][0] if c else ""


def _parse_flag(t):
    t = t.lower()
    if t in _TRUE:
        return True
    if t in _FALSE:
        return False
    raise ValueError("不是开关值: %r" % t)


def _parse_int(t):
    return int(t)


def _parse_float(t):
    return float(t)


def _parse_text(t):
    if not t:
        raise ValueError("空值")
    return t


def pick(name, parser, default):
    """按优先级找第一个**解析得出来**的值，返回 `(值, 来源)`。

    来源是**真正给出这个值**的那一层（不是"第一个存在的"）—— 环境变量写了个
    认不出的值时来源会显示 `ini` 或 `default`，这正是排错时想看到的。
    """
    for v, src in candidates(name):
        try:
            return (parser(v.strip()), src)
        except Exception:
            continue
    return (default, "default")


def flag(name, default):
    """三态开关。`1/true/yes/on` -> True，`0/false/no/off` -> False。

    认不出 → **跳过这一层继续往下看**（见模块文档），都不认 → default。
    """
    return pick(name, _parse_flag, default)[0]


def int_value(name, default):
    return pick(name, _parse_int, default)[0]


def float_value(name, default):
    return pick(name, _parse_float, default)[0]


def text(name, default=""):
    return pick(name, _parse_text, default)[0]


def effective(name, default, kind="flag"):
    """返回 `(值, 来源)`，给自检与测试用。`kind` 取 `flag` / `int` / `float` / `text`。"""
    parser = {"flag": _parse_flag, "int": _parse_int,
              "float": _parse_float, "text": _parse_text}.get(kind, _parse_text)
    return pick(name, parser, default)


def source(name, kind="text"):
    """当前值来自哪一层：`"env"` / `"ini"` / `"default"`（按开关语义解析）。

    ★代码复查修（2026-10-18）：这个函数原先被定义了**两次** —— 后一个用
    `effective(name, None, "flag")`，于是 `RECENT_MAX = 32` / `LOG_FILE = …`
    这类 int / text 键在 ini 里明明写了，`source()` 也会因为"当开关解析不出来"
    而报 `default`（排错时正好把用户引到错方向）。现在只留一份，并且 `kind`
    与 `flag()/int_value()/text()` 走**同一个** `pick()` 口径：

      * `kind="text"`（默认）：任何非空值都认 —— 与 `text()` 一致；
      * `kind="flag"`：认不出的值会被跳过，与 `flag()` 一致。

    `dump()` 直接调 `effective`（那里本来就知道 kind），所以不受影响；
    本函数只服务于"我刚改的这一项到底生效没有"这类排错与测试。
    """
    return effective(name, None, kind)[1]


def items(name, default=()):
    """逗号分隔的清单（窗口类名、家族前缀那些）。空 -> default。"""
    t = text(name, "")
    if not t:
        return tuple(default)
    return tuple(x.strip() for x in t.split(",") if x.strip())


# ---------------------------------------------------------------------------
# 模板
# ---------------------------------------------------------------------------
_BARE_HEADER = """\
; VbeComplete 配置文件
; ─────────────────────────────────────────────────────────────────────────
; 用法：把想改的那一行的 `;` 删掉，改数值，存盘，重启 VbeComplete 生效。
;
;   优先级：环境变量 > 这个文件 > 程序里的默认值
;   改错了不会把功能弄坏 —— 认不出的值会被跳过（那一项就按下一个来源算）。
;
; 说明：这个文件由程序首次启动时**照 settings.example.ini 复制**生成，里面
;   每一项都是注释掉的，所以"什么都不改"就等于"全部用默认值"。settings.ini
;   是你自己的配置（已在 .gitignore 里），settings.example.ini 是仓库里那份
;   模板 —— 想看每一项干什么，改 example 或直接看它。
;
; 想看"现在到底生效的是什么值、来自哪一层"：在项目文件夹里执行
;     py -3 settings.py
; ─────────────────────────────────────────────────────────────────────────

[settings]
"""


def template_text():
    """生成一份"全部注释掉"的 ini 模板 = 全部用默认值。"""
    out = [_BARE_HEADER]
    last_group = None
    for key, default, desc in KNOWN:
        group = _group_of(key)
        if group and group != last_group:
            out.append("\n; ==== %s ====" % group)
            last_group = group
        if default:
            out.append("; %s = %s" % (key, default))
        else:
            out.append("; %s = " % key)
        out.append(";   %s" % desc)
    out.append("")
    return "\n".join(out)


_GROUP_ORDER = (
    ("AUTO_SPACE", "日常最常改的"),
    ("NO_CTRL_NAV", "快捷键"),
    ("VBA_BUILTINS", "候选内容"),
    ("NO_SCOPE_RANK", "排序"),
    ("NO_YIELD", "让位 / 成员位置"),
    ("SNIPPET_DB", "片段库"),
    ("WORD_BANK", "词库（提示词）"),
    ("NO_MOD_CACHE", "性能"),
    ("NO_WINDOW_GATE", "窗口探测（只有类名对不上的宿主才要动）"),
    ("LOG", "排障"),
)


def _group_of(key):
    """这个键属于哪一组（模板里的小标题）= KNOWN 里排在它之前、最近的那个锚点。"""
    idx = _index_of(key)
    cur = None
    for k, title in _GROUP_ORDER:
        if _index_of(k) <= idx:
            cur = title
        else:
            break
    return cur


def _index_of(key):
    for i, (k, _d, _s) in enumerate(KNOWN):
        if k == key:
            return i
    return -1


def ensure_file(target=None):
    """`settings.ini` 不存在就生成一份（全部注释掉 = 全部默认）。

    返回 (路径, 是否新建)。**任何失败都静默**：配置文件是便利设施，
    绝不能因为它写不出来就让补全器起不来（只读安装目录、权限受限都会失败）。
    """
    p = target or path()
    try:
        if os.path.isfile(p):
            return (p, False)
        body = ""
        ex = os.path.join(PROJECT_DIR, EXAMPLE_NAME)
        if os.path.isfile(ex) and os.path.abspath(ex) != os.path.abspath(p):
            try:
                with open(ex, encoding="utf-8-sig", errors="replace") as f:
                    body = f.read()
            except Exception:
                body = ""
        if not body.strip():
            body = template_text()
        d = os.path.dirname(p)
        if d and not os.path.isdir(d):
            os.makedirs(d, exist_ok=True)
        with open(p, "w", encoding="utf-8", newline="\n") as f:
            f.write(body)
        return (p, True)
    except Exception as e:
        _state["error"] = "ensure_file: %s: %s" % (type(e).__name__, e)
        return (p, False)


# ---------------------------------------------------------------------------
# 自检：`py -3 settings.py`
# ---------------------------------------------------------------------------
def _kind_of(default):
    """按默认值的字面猜这个键是哪一类 —— 只服务于自检输出。"""
    d = str(default).strip().lower()
    if d in _TRUE or d in _FALSE:
        return "flag"
    for kind, parser in (("int", _parse_int), ("float", _parse_float)):
        try:
            parser(d)
            return kind
        except Exception:
            pass
    return "text"


def dump(verbose=True):
    """打印每个开关当前生效的值与来源。返回 [(键, 值, 来源)]。"""
    rows = []
    for key, default, desc in KNOWN:
        val, src = effective(PREFIX + key, default, _kind_of(default))
        rows.append((key, val, src))
        if verbose:
            mark = "  " if src == "default" else "* "
            print("%s%-24s = %-18s [%s]  %s"
                  % (mark, key, val, src, desc))
    return rows


def main(argv=None):
    import sys
    argv = list(sys.argv[1:] if argv is None else argv)
    if "--write" in argv or "-w" in argv:
        p, made = ensure_file()
        print(("已生成 " if made else "已存在 ") + p)
        return 0
    print("配置文件: %s" % path())
    if not os.path.isfile(path()):
        print("（还不存在 —— 执行 `py -3 settings.py --write` 生成一份）")
    err = last_error()
    if err:
        print("⚠️ 读配置失败（已按默认值运行）: %s" % err)
    print("带 * 的是被配置文件 / 环境变量改过的项；没有 * 就是程序默认值。\n")
    dump()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
