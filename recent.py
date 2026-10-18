# -*- coding: utf-8 -*-
"""持久化"最近用过的名字"（IDEA 的 recently used 相关度信号）。

★v119：`engine._recent` 原来是**纯内存**字典 —— `stop.bat` 一停就清零，
等于每次开机都要重新"教"一遍。IDEA 那个"最近用过的排前面"是跨会话累积的
（你今天调过的函数，明天打开还在前面），我们只在单次进程内有效，这就把
一个本该越用越顺手的特性变成了"每次重启都从头开始"。

存的东西**极其有限**：名字（小写）+ 一个递增序号。不存代码、不存工程名、
不存文件路径 —— 一个几十条的小 json，几 KB。

位置策略与 `snippets.db` 一致（用户口径："放在项目文件夹里"）：
**本模块所在目录**的 `recent.json`，打包搬走就跟着走。
`VBECOMPLETE_RECENT_FILE` 可覆盖（测试用临时文件、想放别处都靠它）。

⚠️ 必须用 `__file__` 而**不是** `os.getcwd()`：本工具常由 `run.bat` /
pythonw 拉起，工作目录未必是项目根（从别的目录调用、双击快捷方式都会变），
一旦错位就会在别处又建一个空文件（v116 踩过同一个坑）。

本模块**不碰 tkinter、不碰 COM**：纯数据读写，可单独单测（§101）。
"""

import json
import os


DEFAULT_NAME = "recent.json"
ENV_FILE = "VBECOMPLETE_RECENT_FILE"

# 项目文件夹 = 本文件所在目录（同 snippets.PROJECT_DIR 的理由）。
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))

# 数据层自己的硬上限：engine 那边有 RECENT_MAX（默认 64），这里是"就算文件
# 被人手工改出一万条也不至于把内存/排序拖慢"的保险。取 8 倍的余量。
MAX_ENTRIES = 512


def path():
    """文件路径：环境变量优先，否则**项目文件夹**里的 `recent.json`。"""
    p = (os.environ.get(ENV_FILE) or "").strip()
    if p:
        return p
    return os.path.join(PROJECT_DIR, DEFAULT_NAME)


def _rebase(mapping):
    """把序号重排成 1..N（按原有先后顺序），保持相对次序完全不变。

    为什么重排：只有"谁比谁更近"有意义，绝对数值没有。重排之后
    engine 的淘汰逻辑（砍最旧的一批）永远从干净的小整数起步，
    不会因为跨了几百次会话而攒出很大的数字。
    """
    if not mapping:
        return {}
    ordered = sorted(mapping.items(), key=lambda kv: kv[1])
    return {k: i + 1 for i, (k, _v) in enumerate(ordered)}


def load(file_path=None):
    """读回 {名字小写: 序号}；文件不存在 / 坏 / 类型不对一律返回 {}。

    **绝不抛异常**：这个文件是"锦上添花"的数据，读不到就当没有 ——
    绝不能因为它导致工具起不来（那是最不可接受的失败模式）。
    """
    p = file_path or path()
    try:
        with open(p, "r", encoding="utf-8") as f:
            raw = json.load(f)
    except Exception:
        return {}
    if not isinstance(raw, dict):
        return {}

    clean = {}
    for k, v in raw.items():
        if not isinstance(k, str):
            continue
        name = k.strip()
        if not name:
            continue
        # bool 是 int 的子类，`True` 会被 int() 收成 1 —— 显式排掉。
        if isinstance(v, bool):
            continue
        try:
            n = int(v)
        except (TypeError, ValueError):
            continue
        if n < 0:
            continue
        clean[name] = n

    if len(clean) > MAX_ENTRIES:
        clean = dict(sorted(clean.items(),
                            key=lambda kv: kv[1], reverse=True)[:MAX_ENTRIES])
    return _rebase(clean)


def save(mapping, file_path=None):
    """把 {名字: 序号} 写回磁盘。成功返回 True，任何失败返回 False。

    用"临时文件 + os.replace"而不是直接覆盖：直接覆盖写到一半被强杀
    （`stop.bat` 就是 `Stop-Process -Force`）会留下半截 json，下次
    读回来是空的 —— 用户的"最近使用"就整个丢了。replace 在同盘上是原子的。
    """
    if not isinstance(mapping, dict):
        return False
    data = {}
    for k, v in mapping.items():
        try:
            name = str(k).strip()
            if not name or isinstance(v, bool):
                continue
            n = int(v)
        except (TypeError, ValueError):
            continue
        if n < 0:
            continue
        data[name] = n

    p = file_path or path()
    tmp = p + ".tmp"
    try:
        d = os.path.dirname(p)
        if d and not os.path.isdir(d):
            os.makedirs(d, exist_ok=True)
        with open(tmp, "w", encoding="utf-8", newline="\n") as f:
            json.dump(data, f, ensure_ascii=False, sort_keys=True)
        os.replace(tmp, p)
        return True
    except Exception:
        # 失败要清掉临时文件，别在项目文件夹里留垃圾
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass
        return False
