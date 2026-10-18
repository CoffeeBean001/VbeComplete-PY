"""词库（★v123）：把"语言自带的提示词"从代码常量搬进一个可手改的 SQLite 库。

## 为什么要有这一层

v122 之前，VBA 的关键字 / 内建函数 / 内建类型 / 常用常量全是 `vba_builtins.py` 里的
Python 元组。想加一个自己常用的宿主 API、想把 `vbKey*` 那批全关掉、觉得 `Kill` 这种
名字不该出现在提示里 —— 每改一处都得动源码、还要保证语法没写错。

用户的口径（原话）："**是不是可以创建个 SQLite 数据库，存放在这个数据里，然后项目
启动，就把数据库里的加载到内存里，我也可以手动去编辑数据库，项目就能动态加载提示
关键字。**"

于是分层变成：

    vba_builtins.py   —— 种子清单（首次建库用）+ 生效快照（ACTIVE_*）
    wordbank.py       —— 本文件：只管"SQLite <-> 行列表"，**不认识 VBA 语义**
    收集侧 / 引擎侧   —— 照旧读 `vba_builtins.ACTIVE_*`，一个字节都不改

## 三条硬约束（决定了这个文件为什么长这样）

1. **热路径零成本**。加载发生在**启动时一次**，之后内存里就是一堆 tuple/frozenset，
   与改之前完全等价。本模块**不允许**被 engine / 补全路径 import（§109.7 用 AST 钉住）。
2. **库坏了绝不影响主功能**。`load()` 出任何问题（不是 SQLite、缺表、版本更高、
   被别的进程锁住）一律返回 `None` + 记在 `last_error()`，调用方退回种子清单。
   一个"可选的增强功能"没有资格让补全整体失效。
3. **绝不覆盖用户的手改**。`ensure()` 只在**文件不存在**时建库；文件已经在那儿了
   就一个字节都不碰（哪怕表结构是旧的 —— 那由 schema_version 判断，见下）。

## 为什么是"只读打开"

用户会拿 DB Browser 之类的工具**一边开着我们一边改库**。所以：

  * `load()` 用 `mode=ro` 的 URI 打开，读完立刻关，**不持有任何锁**；
  * 我们从不写库（除了首次 `ensure()` 建库）；
  * 于是"程序在跑"和"你在改库"互不干扰 —— 改完重启程序即可生效。

## 表结构

    words(name, kind, is_builtin, enabled, wants_space, note)
        PRIMARY KEY (name, kind)

★ 主键是 **(name, kind)** 而不是 name：`Date` 与 `String` 既是内建函数又是内建类型，
  两行都要在（引擎在 `As |` 位置只认 type、在表达式位置只认 function，混成一行就
  分不开了 —— 这是铁律 6 那条"两道判据问不同问题"的直接后果）。

  kind         keyword / function / constant / type（四组，与四个开关一一对应）
  is_builtin   1 = 语言自带（参与"公共词汇"模糊匹配收紧）；0 = 用户自加（不收紧）
  enabled      0 = 保留但不提示（比删掉好：随时能开回来，也不丢 note）
  wants_space  仅对 keyword 有意义。NULL = 按默认排除法；1 = 一定补空格；0 = 一定不补
  note         纯备注，程序不读

    signatures(name, sig)          ← ★v125
        PRIMARY KEY (name)

`sig` 是**括号里面那段原文**（`string, length`），不存函数名 —— 名字就是主键，
  存两遍只会分叉。可选参数按 VBA 文档写成 `[length]`。空串 = 真·无参函数
  （`Now()`），与"这个名字没有签名"（查不到行）是两件事。

★ **为什么签名跟词库同一个文件、而不是"拆成两个库"**（用户提议过）：
  用户要的是"名字与签名分开放、各自能改"—— 那用**两张表**就实现了；
  拆成两个 .db 的话，启动要开两个连接、两套 ensure/load/CLI、两处排错，
  而收益只是"文件能单独拷走"。代价与收益不成比例。表级拆分已经满足
  "改签名不碰名字、改名字不碰签名"。

  ⚠️ 老用户的库是 v123 建的、**没有 signatures 表**。`ensure()` 在库已存在时
  会**只补建这张缺失的表**、并仅在**表为空**时灌种子 —— `words` 表一个字节不碰
  （"绝不覆盖用户手改"那条铁律守的是数据，不是表结构）。建表失败（别的进程
  锁着库）就静默跳过，签名功能退回"只有工程内解析出来的签名"，程序照常跑。

## 命令行

不用装任何 GUI 工具也能改（改完重启 VbeComplete 生效）：

    py -3 wordbank.py                    统计：库在哪、每组多少条、有几条被禁用
    py -3 wordbank.py list [kind]        逐条列出
    py -3 wordbank.py add Name [kind]    加一条（默认 function；is_builtin=0）
    py -3 wordbank.py disable Name       保留但不提示（与 kind 无关，所有 kind 一起关）
    py -3 wordbank.py enable Name
    py -3 wordbank.py remove Name        真删
    py -3 wordbank.py reset --yes        重建（先自动备份成 wordbank.db.bak-<时间>）
    py -3 wordbank.py export 词库.csv
    py -3 wordbank.py import 词库.csv    合并（同名同 kind 覆盖）
    py -3 wordbank.py signatures [过滤]  列签名（★v125）
    py -3 wordbank.py set-sig Name "a, [b]"
                                         改一条签名；传 "" 表示无参函数（★v125）
"""

import csv
import os
import sqlite3
import sys
import time

try:
    import settings as _settings
except Exception:                                    # pragma: no cover
    _settings = None

DEFAULT_DB_NAME = "wordbank.db"
SCHEMA_VERSION = 1
KINDS = ("keyword", "function", "constant", "type")

# 项目文件夹 = 本文件所在目录。★必须用 `__file__` 而不是 `os.getcwd()`：
# 本工具常由 `run.bat` / pythonw 拉起，工作目录未必是项目根（v116 在 snippets.db
# 上踩过一次，这里照同一条规矩）。
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))

_last_error = ""


def last_error():
    """上一次失败的原因（排错用；成功时是空串）。"""
    return _last_error


def _say(msg):
    global _last_error
    _last_error = str(msg)


def db_path():
    """库文件路径：配置中心（`WORD_BANK`）优先，否则**项目文件夹**里的 `wordbank.db`。"""
    try:
        if _settings is not None:
            p = _settings.text("WORD_BANK", "")
            if p:
                return p
    except Exception:
        pass
    return os.path.join(PROJECT_DIR, DEFAULT_DB_NAME)


# ---------------------------------------------------------------------------
# 建库
# ---------------------------------------------------------------------------
_SCHEMA = """
CREATE TABLE IF NOT EXISTS words (
    name        TEXT    NOT NULL,
    kind        TEXT    NOT NULL,
    is_builtin  INTEGER NOT NULL DEFAULT 1,
    enabled     INTEGER NOT NULL DEFAULT 1,
    wants_space INTEGER,
    note        TEXT,
    PRIMARY KEY (name, kind)
);
CREATE TABLE IF NOT EXISTS signatures (
    name TEXT NOT NULL,
    sig  TEXT NOT NULL DEFAULT '',
    PRIMARY KEY (name)
);
CREATE TABLE IF NOT EXISTS meta (
    key   TEXT PRIMARY KEY,
    value TEXT
);
"""


def _connect_rw(path):
    """可写连接。

    ★**刻意不用 WAL**（`snippets.db` 也没有）。原因在只读那一侧：`load()` 用
    `mode=ro` 打开，而 WAL 库的只读连接还需要能写 `-shm` 文件，否则 SQLite 直接
    报 `unable to open database file`；用户再拿 DB Browser 一边开着改，
    三方的锁关系就更难看懂了。我们一次启动只读一遍、写完就关，DELETE 日志
    （SQLite 默认）足够，而且**用户用任何工具编辑都不会有额外要求**。
    """
    return sqlite3.connect(path, timeout=5)


def _fill_signatures(con, sigs):
    """在一个**已打开**的连接上补建 `signatures` 表、并在表为空时灌种子。

    返回灌进去的条数（0 = 本来就有数据 / 没给种子）。

    ⚠️ 只在**表为空**时灌：用户改过签名就把这张表当他的地盘，我们一个字节不碰
    （与 `words` 表同一条规矩）。`INSERT OR IGNORE` 只是兜底 —— 空表里不会有
    冲突，但万一并发写进来一条，也不会炸。
    """
    con.execute("CREATE TABLE IF NOT EXISTS signatures ("
                " name TEXT NOT NULL, sig TEXT NOT NULL DEFAULT '',"
                " PRIMARY KEY (name))")
    n = con.execute("SELECT COUNT(*) FROM signatures").fetchone()[0]
    if n or not sigs:
        return 0
    con.executemany(
        "INSERT OR IGNORE INTO signatures (name, sig) VALUES (?, ?)", sigs)
    con.commit()
    return len(sigs)


def ensure(rows, path=None, sigs=None):
    """库不存在就按 `rows` 建；存在则**只补建缺失的 `signatures` 表**。

    `rows` 是 `[(name, kind, is_builtin, enabled, wants_space, note), ...]`。
    `sigs` 是 `[(name, sig), ...]`（★v125，可省略 —— 省略就不管签名）。
    ★ 本模块**不知道**这些行从哪来（不 import `vba_builtins` / `vba_signatures`）——
    调用方传进来，这样 wordbank 可以独立测试，也不会形成循环 import。
    """
    p = path or db_path()
    if os.path.exists(p):
        # ★v125：老库（v123 建的）没有 signatures 表 —— 补建。任何失败都静默：
        # 签名是增强功能，没有资格让启动失败（尤其"用户正用 DB Browser 开着库"）。
        if sigs:
            try:
                con = _connect_rw(p)
                try:
                    _fill_signatures(con, sigs)
                finally:
                    con.close()
            except Exception as e:
                _say("补建 signatures 表失败（不影响补全）: %s" % e)
        return (False, p)
    d = os.path.dirname(os.path.abspath(p))
    try:
        if d and not os.path.isdir(d):
            os.makedirs(d)
    except Exception as e:
        _say("建目录失败: %s" % e)
        return (False, p)
    try:
        con = _connect_rw(p)
        try:
            con.executescript(_SCHEMA)
            con.executemany(
                "INSERT OR REPLACE INTO words"
                " (name, kind, is_builtin, enabled, wants_space, note)"
                " VALUES (?, ?, ?, ?, ?, ?)", rows)
            if sigs:
                _fill_signatures(con, sigs)
            con.execute("INSERT OR REPLACE INTO meta VALUES ('schema_version', ?)",
                        (str(SCHEMA_VERSION),))
            con.execute("INSERT OR REPLACE INTO meta VALUES ('created_at', ?)",
                        (time.strftime("%Y-%m-%d %H:%M:%S"),))
            con.commit()
        finally:
            con.close()
        return (True, p)
    except Exception as e:
        _say("建库失败: %s" % e)
        try:
            os.remove(p)                             # 别留个半成品挡下次启动
        except Exception:
            pass
        return (False, p)


# ---------------------------------------------------------------------------
# 读库
# ---------------------------------------------------------------------------
class Snapshot(object):
    """一次加载的结果。字段都是**现成的**元组 / 集合，下游零加工。

    `rows` 保留原始行（含 wants_space 的 None 三态），因为"NULL 是什么意思"
    是 VBA 语义，得由 `vba_builtins` 决定，不是本模块的事。

    ★v125：`sigs` 是 `{小写函数名: 括号内原文}`。查不到 = 没有签名；
    空串 = 真·无参函数（`Now()`）—— 这两种必须能分开，所以用 dict 而不是
    "空串即没有"。
    """

    __slots__ = ("rows", "path", "counts", "sigs")

    def __init__(self, rows, path, sigs=None):
        self.rows = rows
        self.path = path
        self.sigs = dict(sigs or {})
        self.counts = {}
        for r in rows:
            self.counts[r[1]] = self.counts.get(r[1], 0) + 1

    def by_kind(self, kind):
        """某个 kind 的 [(name, is_builtin, wants_space)]（只含 enabled=1）。"""
        return [(r[0], r[2], r[4]) for r in self.rows
                if r[1] == kind and r[3]]

    def signature(self, name):
        """按名字查形参表原文；**没有**这条签名返回 None（不是空串）。"""
        try:
            return self.sigs.get(str(name).lower())
        except Exception:
            return None

    def __repr__(self):
        return "<Snapshot %s %r sigs=%d>" % (self.path, self.counts,
                                             len(self.sigs))


def _ro_uri(path):
    """`file:///D:/x/wordbank.db?mode=ro` —— 只读打开，**不存在也不会被创建**。

    ⚠️ 不用 `urllib.request.pathname2url`：它在 Windows 上返回的是 `///C:/…`
    （已经带三斜杠），再补一个 `file://` 就变成五斜杠，SQLite 直接
    `unable to open database file`。自己拼最稳，顺手把空格 / 中文转义掉。
    """
    from urllib.parse import quote
    p = os.path.abspath(path).replace("\\", "/")
    if not p.startswith("/"):
        p = "/" + p                                  # C:/x -> /C:/x
    return "file://" + quote(p, safe="/:") + "?mode=ro"


def load(path=None):
    """把库读成 `Snapshot`；任何问题都返回 `None`（调用方退回种子清单）。

    失败原因记在 `last_error()`，**不抛异常** —— 这是本模块对外的契约。
    """
    p = path or db_path()
    if not os.path.isfile(p):
        _say("库不存在: %s" % p)
        return None
    try:
        con = sqlite3.connect(_ro_uri(p), uri=True, timeout=3)
    except Exception as e:
        _say("打不开: %s" % e)
        return None
    try:
        try:
            ver = con.execute(
                "SELECT value FROM meta WHERE key='schema_version'").fetchone()
        except Exception as e:
            _say("缺 meta 表（不是本程序的词库？）: %s" % e)
            return None
        try:
            iv = int(str(ver[0]).strip()) if ver else 0
        except Exception:
            iv = 0
        if iv > SCHEMA_VERSION:
            # 比本程序新 —— 不认识新列/新语义，宁可整库不读，也不要猜着解析
            _say("库版本 %d 比程序支持的 %d 新，已跳过" % (iv, SCHEMA_VERSION))
            return None
        try:
            cur = con.execute(
                "SELECT name, kind, is_builtin, enabled, wants_space, note"
                " FROM words")
            raw = cur.fetchall()
        except Exception as e:
            _say("words 表读不出来: %s" % e)
            return None
        # ★v125：签名表是可选的 —— v123 建的库没有它，读不到就当"没有签名"，
        # 绝不能让整库退回种子（那会丢掉用户对 words 的手改）。
        sigs = {}
        try:
            for sname, ssig in con.execute("SELECT name, sig FROM signatures"):
                k = str(sname or "").strip().lower()
                if k:
                    sigs[k] = "" if ssig is None else str(ssig)
        except Exception:
            sigs = {}
    finally:
        try:
            con.close()
        except Exception:
            pass

    rows = []
    seen = set()
    for name, kind, is_b, en, ws, note in raw:
        name = str(name or "").strip()
        kind = str(kind or "").strip().lower()
        if not name or kind not in KINDS:
            continue                                 # 用户手改出来的脏行，跳过即可
        key = (name.lower(), kind)
        if key in seen:
            continue                                 # 大小写不同的重复行：第一条赢
        seen.add(key)
        rows.append((name, kind, 0 if not is_b else 1, 0 if not en else 1,
                     None if ws is None else (1 if ws else 0),
                     "" if note is None else str(note)))
    if not rows:
        _say("库是空的（一条可用的词都没有）")
        return None
    _say("")
    return Snapshot(rows, p, sigs)


# ---------------------------------------------------------------------------
# 命令行
# ---------------------------------------------------------------------------
def _open_rw(path):
    """给 CLI 用：打开可写的库（不存在就报错，**不悄悄建空库**）。"""
    if not os.path.isfile(path):
        raise SystemExit("词库不存在：%s\n先用 `py -3 main.py` 启动一次，或 "
                         "`py -3 wordbank.py reset --yes` 重建。" % path)
    return _connect_rw(path)


def _cmd_stats(path):
    snap = load(path)
    if snap is None:
        print("词库：%s" % path)
        print("  ⚠️ 读不出来：%s" % (last_error() or "未知原因"))
        print("  程序此刻会退回代码里的种子清单（功能照常）。")
        return 1
    print("词库：%s" % snap.path)
    con = _open_rw(path)
    try:
        dis = dict(con.execute(
            "SELECT kind, COUNT(*) FROM words WHERE enabled=0 GROUP BY kind"
        ).fetchall())
        cus = dict(con.execute(
            "SELECT kind, COUNT(*) FROM words WHERE is_builtin=0 GROUP BY kind"
        ).fetchall())
    finally:
        con.close()
    tot = sum(snap.counts.values())
    print("  共 %d 条（已禁用的不计）" % tot)
    for k in KINDS:
        n = snap.counts.get(k, 0)
        extra = []
        if dis.get(k):
            extra.append("禁用 %d" % dis[k])
        if cus.get(k):
            extra.append("自加 %d" % cus[k])
        print("    %-9s %4d%s" % (k, n, ("　·　" + "　·　".join(extra)) if extra else ""))
    print("\n改完重启 VbeComplete 生效。逐条看：py -3 wordbank.py list")
    return 0


def _cmd_list(path, kind=None):
    con = _open_rw(path)
    try:
        sql = ("SELECT name, kind, is_builtin, enabled, wants_space, note"
               " FROM words")
        args = ()
        if kind:
            sql += " WHERE kind=?"
            args = (kind,)
        sql += " ORDER BY kind, name COLLATE NOCASE"
        rows = con.execute(sql, args).fetchall()
    finally:
        con.close()
    if not rows:
        print("（没有匹配的词）")
        return 0
    cur = None
    for name, k, is_b, en, ws, note in rows:
        if k != cur:
            print("\n[%s]" % k)
            cur = k
        flags = []
        if not is_b:
            flags.append("自加")
        if not en:
            flags.append("禁用")
        if ws is not None:
            flags.append("空格=%d" % ws)
        tail = ("　# " + note) if note else ""
        print("  %-28s %s%s" % (name, " ".join(flags), tail))
    print("\n共 %d 条" % len(rows))
    return 0


def _cmd_add(path, name, kind="function"):
    name = str(name or "").strip()
    if not name:
        raise SystemExit("名字不能为空")
    kind = (kind or "function").strip().lower()
    if kind not in KINDS:
        raise SystemExit("kind 只能是 %s" % " / ".join(KINDS))
    con = _open_rw(path)
    try:
        con.execute(
            "INSERT OR REPLACE INTO words"
            " (name, kind, is_builtin, enabled, wants_space, note)"
            " VALUES (?, ?, 0, 1, NULL, ?)",
            (name, kind, "手加于 " + time.strftime("%Y-%m-%d")))
        con.commit()
    finally:
        con.close()
    print("已加：%s（%s，自加）" % (name, kind))
    return 0


def _cmd_set_enabled(path, name, enabled):
    con = _open_rw(path)
    try:
        cur = con.execute("UPDATE words SET enabled=? WHERE lower(name)=lower(?)",
                          (1 if enabled else 0, name))
        n = cur.rowcount
        con.commit()
    finally:
        con.close()
    if not n:
        raise SystemExit("没找到：%s" % name)
    print("%s 了 %d 条：%s" % ("启用" if enabled else "禁用", n, name))
    return 0


def _cmd_remove(path, name):
    con = _open_rw(path)
    try:
        cur = con.execute("DELETE FROM words WHERE lower(name)=lower(?)", (name,))
        n = cur.rowcount
        con.commit()
    finally:
        con.close()
    if not n:
        raise SystemExit("没找到：%s" % name)
    print("已删除 %d 条：%s" % (n, name))
    return 0


def _cmd_reset(path, yes):
    if not yes:
        raise SystemExit("这会丢掉你对词库的所有手改。确认请加 --yes")
    if os.path.isfile(path):
        bak = path + ".bak-" + time.strftime("%Y%m%d-%H%M%S")
        try:
            with open(path, "rb") as a, open(bak, "wb") as b:
                b.write(a.read())
            print("已备份：%s" % bak)
        except Exception as e:
            raise SystemExit("备份失败，为了不改掉你的数据，已中止：%s" % e)
        for suf in ("-wal", "-shm"):
            try:
                os.remove(path + suf)
            except Exception:
                pass
        os.remove(path)
    import vba_builtins
    import vba_signatures
    created, p = ensure(vba_builtins.seed_rows(), path,
                        sigs=vba_signatures.seed_rows())
    print("已重建：%s（%s）" % (p, "新建" if created else "失败"))
    return 0 if created else 1


def _cmd_signatures(path, flt=None):
    """列出签名（可按名字子串过滤）。★v125"""
    con = _open_rw(path)
    try:
        try:
            rows = con.execute(
                "SELECT name, sig FROM signatures"
                " ORDER BY name COLLATE NOCASE").fetchall()
        except Exception:
            raise SystemExit("这个库还没有 signatures 表（v123 建的老库）。"
                             "先启动一次 VbeComplete 会自动补建，"
                             "或 `py -3 wordbank.py reset --yes` 重建。")
    finally:
        con.close()
    if flt:
        f = flt.lower()
        rows = [r for r in rows if f in str(r[0]).lower()]
    if not rows:
        print("（没有匹配的签名）")
        return 0
    for name, sig in rows:
        print("%-14s(%s)" % (name, "" if sig is None else sig))
    print("---- 共 %d 条 ----" % len(rows))
    return 0


def _cmd_set_sig(path, name, sig):
    """改一条签名；`sig` 传 "" = 无参函数。★v125"""
    con = _open_rw(path)
    try:
        con.execute("CREATE TABLE IF NOT EXISTS signatures ("
                    " name TEXT NOT NULL, sig TEXT NOT NULL DEFAULT '',"
                    " PRIMARY KEY (name))")
        con.execute("INSERT OR REPLACE INTO signatures (name, sig) VALUES (?, ?)",
                    (str(name), "" if sig is None else str(sig)))
        con.commit()
    finally:
        con.close()
    print("%s(%s)  已写入" % (name, "" if sig is None else sig))
    return 0


_CSV_HEADER = ("name", "kind", "is_builtin", "enabled", "wants_space", "note")


def _cmd_export(path, out):
    con = _open_rw(path)
    try:
        rows = con.execute(
            "SELECT name, kind, is_builtin, enabled, wants_space, note"
            " FROM words ORDER BY kind, name COLLATE NOCASE").fetchall()
    finally:
        con.close()
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(_CSV_HEADER)
        for r in rows:
            w.writerow(["" if v is None else v for v in r])
    print("已导出 %d 条 -> %s" % (len(rows), out))
    return 0


def _cmd_import(path, src):
    with open(src, newline="", encoding="utf-8-sig") as f:
        rd = csv.DictReader(f)
        rows = []
        for d in rd:
            nm = str(d.get("name") or "").strip()
            kd = str(d.get("kind") or "").strip().lower()
            if not nm or kd not in KINDS:
                continue
            ws = str(d.get("wants_space") or "").strip()
            rows.append((nm, kd,
                         int(str(d.get("is_builtin") or "0").strip() or 0),
                         1 if str(d.get("enabled") or "1").strip() not in
                         ("0", "", "False", "false") else 0,
                         None if ws == "" else (1 if ws not in ("0", "False",
                                                               "false") else 0),
                         str(d.get("note") or "")))
    if not rows:
        raise SystemExit("没有可导入的行")
    con = _open_rw(path)
    try:
        con.executemany(
            "INSERT OR REPLACE INTO words"
            " (name, kind, is_builtin, enabled, wants_space, note)"
            " VALUES (?, ?, ?, ?, ?, ?)", rows)
        con.commit()
    finally:
        con.close()
    print("已导入/合并 %d 条 <- %s" % (len(rows), src))
    return 0


def main(argv):
    args = [a for a in argv[1:]]
    cmd = args[0].lower() if args else "stats"
    rest = args[1:]
    path = db_path()
    # ★代码复查修（2026-10-18）：缺参数时原先直接 `rest[0]` -> IndexError
    # 堆栈（`py -3 wordbank.py disable` 就会看到），这里给一句能照抄的用法。
    _need = {"add": "add <名字> [kind]", "disable": "disable <名字>",
             "enable": "enable <名字>", "remove": "remove <名字>"}
    if cmd in _need and not rest:
        raise SystemExit("用法：%s" % _need[cmd])
    if cmd in ("stats", "status"):
        return _cmd_stats(path)
    if cmd == "list":
        return _cmd_list(path, rest[0].lower() if rest else None)
    if cmd == "add":
        if not rest:
            raise SystemExit("用法：add <名字> [kind]")
        return _cmd_add(path, rest[0], rest[1] if len(rest) > 1 else "function")
    if cmd == "disable":
        return _cmd_set_enabled(path, rest[0], False)
    if cmd == "enable":
        return _cmd_set_enabled(path, rest[0], True)
    if cmd == "remove":
        return _cmd_remove(path, rest[0])
    if cmd == "reset":
        return _cmd_reset(path, "--yes" in rest)
    if cmd == "export":
        return _cmd_export(path, rest[0] if rest else "wordbank.csv")
    if cmd == "import":
        return _cmd_import(path, rest[0] if rest else "wordbank.csv")
    # ★v125：签名（形参提示）
    if cmd in ("signatures", "sigs", "sig"):
        return _cmd_signatures(path, rest[0] if rest else None)
    if cmd in ("set-sig", "setsig"):
        if len(rest) < 1:
            raise SystemExit('用法：set-sig <名字> "<形参表>"  例如：'
                             'set-sig Left "string, length"')
        return _cmd_set_sig(path, rest[0],
                            rest[1] if len(rest) > 1 else "")
    raise SystemExit(__doc__.split("## 命令行")[1])


if __name__ == "__main__":
    sys.exit(main(sys.argv))
