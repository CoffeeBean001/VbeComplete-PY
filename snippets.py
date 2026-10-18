# -*- coding: utf-8 -*-
"""代码片段库（SQLite）+ 标题模糊匹配。

用户口径（v115，原话）：
  "我想加个代码片段存储的功能。帮我创建个SQLite数据库，里面用来存储我的代码，
   具体功能如下：
   1. 数据库分两列，第1列是代码片段的标题，第2列是代码片段的具体代码。"

⇒ 表就**只有两列**：`title` + `code`。刻意不加 id：标题天然唯一、足以当主键，
   重命名走 `UPDATE snippets SET title=? WHERE title=?`（见 `rename`）。
   用户打开库看到的列数与要求一致。

文件位置：★v116（用户口径："能帮我把建的库，默认位置放在项目文件夹里吗？"）
    ⇒ 默认就在**本模块所在目录**（项目文件夹）的 `snippets.db`，一眼能找到、打包
    搬走就跟着走。`VBECOMPLETE_SNIPPET_DB` 可覆盖（测试用临时库、想放别处都靠它）。

本模块**不碰 tkinter、不碰 COM**：全是纯数据操作，可单独单测（§99）。
"""

import os
import sqlite3


SCHEMA = """
CREATE TABLE IF NOT EXISTS snippets (
    title TEXT PRIMARY KEY,
    code  TEXT NOT NULL
)
"""

DEFAULT_DB_NAME = "snippets.db"
ENV_DB = "VBECOMPLETE_SNIPPET_DB"

# 项目文件夹 = 本文件所在目录。★必须用 `__file__` 而不是 `os.getcwd()`：
# 本工具常由 `run.bat` / pythonw 拉起，工作目录未必是项目根（双击快捷方式、
# 从别的目录调用都会变），一旦错位就会在别处又建一个空库。
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))


def db_path():
    """库文件路径：环境变量优先，否则**项目文件夹**里的 `snippets.db`。"""
    p = (os.environ.get(ENV_DB) or "").strip()
    if p:
        return p
    return os.path.join(PROJECT_DIR, DEFAULT_DB_NAME)


# ---------------------------------------------------------------------------
# 标题模糊匹配（纯函数，供 UI 过滤列表用）
# ---------------------------------------------------------------------------
# 口径（v115）：用户在搜索框里敲的东西要能"模糊"命中标题。三档命中，强度递减：
#   ① 整标题相等    ② 前缀    ③ 任意位置连续子串    ④ 子序列（`gc` -> `GetCode`）
# 多个词用空格分开，**每个词都要命中**（AND），与 IDEA 的搜索一致。
#
# ⚠️ 大小写归一必须**保持长度**：`"İ".lower()` 长度会变成 2，一旦用它当下标去
# 切原串就整体错位（项目铁律第 22 条）。所以逐字符 lower，长度不变才采用。

def _lower_keep_len(text):
    """逐字符转小写，且保证长度与原文一致（异常字符原样保留）。"""
    out = []
    for ch in text or "":
        low = ch.lower()
        out.append(low if len(low) == 1 else ch)
    return "".join(out)


_BOUNDARY_CHARS = " _-.[](){}+*/\\,;:=<>\"'"


def _is_boundary(title, idx):
    """title[idx] 是否落在"词的起点"上（驼峰分界 / 分隔符之后）。

    只用于给子序列命中**加分**（`gc` 命中 `GetCode` 的两个大写字母时更像用户
    的本意），不影响"中不中"。
    """
    if idx <= 0:
        return True
    prev, cur = title[idx - 1], title[idx]
    if prev in _BOUNDARY_CHARS:
        return True
    if cur.isupper() and prev.islower():
        return True
    if cur.isdigit() and not prev.isdigit():
        return True
    return False


def _term_score(term_low, title, title_low):
    """单个词在标题上的命中情况 -> (分数, 命中下标列表) 或 None（没命中）。"""
    n = len(term_low)
    if n <= 0:
        return None
    pos = title_low.find(term_low)
    if pos >= 0:
        hits = list(range(pos, pos + n))
        if title_low == term_low:
            return 100 + n * 4, hits          # 整标题相等
        if pos == 0:
            return 80 + n * 4, hits           # 前缀
        return 60 + n * 4, hits               # 任意位置连续子串
    # 子序列：贪心从左往右扫（存在性判定用贪心最左是充分的）
    hits = []
    i = 0
    for idx, ch in enumerate(title_low):
        if i < n and ch == term_low[i]:
            hits.append(idx)
            i += 1
            if i == n:
                break
    if i < n:
        return None
    score = 20 + n * 2
    score += sum(3 for h in hits if _is_boundary(title, h))
    score += sum(1 for a, b in zip(hits, hits[1:]) if b == a + 1)
    return score, hits


def match(query, title):
    """query 是否能模糊命中 title。命中返回 (分数, 命中下标列表)，否则 None。"""
    q = (query or "").strip()
    if not q:
        return (0, [])
    title = title or ""
    title_low = _lower_keep_len(title)
    total = 0
    hits = set()
    for term in q.split():
        got = _term_score(_lower_keep_len(term), title, title_low)
        if got is None:
            return None
        total += got[0]
        hits.update(got[1])
    return total, sorted(hits)


def search(rows, query):
    """在 [(title, code), ...] 里按标题模糊过滤。

    返回 [(title, code, hits)]，按"分数高 -> 标题短 -> 字母序"排。
    空 query 原样返回全部（保持库里的顺序）。
    """
    rows = list(rows or ())
    if not (query or "").strip():
        return [(t, c, []) for (t, c) in rows]
    out = []
    for (t, c) in rows:
        got = match(query, t)
        if got is not None:
            out.append((got[0], len(t or ""), t, c, got[1]))
    out.sort(key=lambda r: (-r[0], r[1], _lower_keep_len(r[2])))
    return [(t, c, hits) for (_s, _n, t, c, hits) in out]


# ---------------------------------------------------------------------------
# 存储
# ---------------------------------------------------------------------------

class SnippetStore(object):
    """两列表的增删改查。所有方法都在调用线程（tk 主线程）里跑，不跨线程。"""

    def __init__(self, path=None):
        self.path = path or db_path()
        self._con = None

    # ---- 连接 ----
    def _conn(self):
        if self._con is not None:
            return self._con
        d = os.path.dirname(os.path.abspath(self.path))
        if d and not os.path.isdir(d):
            os.makedirs(d, exist_ok=True)
        con = sqlite3.connect(self.path, timeout=5)
        con.execute(SCHEMA)
        con.commit()
        self._con = con
        return con

    def close(self):
        try:
            if self._con is not None:
                self._con.close()
        except Exception:
            pass
        self._con = None

    # ---- 读 ----
    def all(self):
        """全部片段 [(title, code)]，按标题排序（不区分大小写）。"""
        try:
            cur = self._conn().execute(
                "SELECT title, code FROM snippets ORDER BY title COLLATE NOCASE")
            return [(str(t), str(c)) for (t, c) in cur.fetchall()]
        except Exception:
            return []

    def get(self, title):
        """取一条的代码；不存在返回 None。"""
        try:
            cur = self._conn().execute(
                "SELECT code FROM snippets WHERE title=?", (title,))
            row = cur.fetchone()
            return None if row is None else str(row[0])
        except Exception:
            return None

    def exists(self, title):
        return self.get(title) is not None

    def count(self):
        try:
            cur = self._conn().execute("SELECT COUNT(*) FROM snippets")
            row = cur.fetchone()
            return int(row[0]) if row else 0
        except Exception:
            return 0

    # ---- 写 ----
    def add(self, title, code):
        """新增。标题已存在时**不覆盖**（返回 False，由 UI 提示改走"保存修改"）。"""
        title = (title or "").strip()
        if not title:
            return False
        try:
            con = self._conn()
            con.execute("INSERT INTO snippets (title, code) VALUES (?, ?)",
                        (title, code or ""))
            con.commit()
            return True
        except sqlite3.IntegrityError:
            return False
        except Exception:
            return False

    def update(self, title, code):
        """改代码（标题不变）。返回是否真的改到了行。"""
        title = (title or "").strip()
        if not title:
            return False
        try:
            con = self._conn()
            cur = con.execute("UPDATE snippets SET code=? WHERE title=?",
                              (code or "", title))
            con.commit()
            return cur.rowcount > 0
        except Exception:
            return False

    def rename(self, old, new):
        """改标题。新标题已被占用时不动任何数据，返回 False。"""
        old = (old or "").strip()
        new = (new or "").strip()
        if not old or not new:
            return False
        if old == new:
            return True
        try:
            con = self._conn()
            if self.exists(new):
                return False
            cur = con.execute("UPDATE snippets SET title=? WHERE title=?",
                              (new, old))
            con.commit()
            return cur.rowcount > 0
        except Exception:
            return False

    def delete(self, title):
        """删除一条。返回是否真的删掉了。"""
        title = (title or "").strip()
        if not title:
            return False
        try:
            con = self._conn()
            cur = con.execute("DELETE FROM snippets WHERE title=?", (title,))
            con.commit()
            return cur.rowcount > 0
        except Exception:
            return False
