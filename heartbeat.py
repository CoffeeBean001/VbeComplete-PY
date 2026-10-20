# -*- coding: utf-8 -*-
"""运行期心跳（★v128）：让**另一个进程**能判断"我到底还活不活"。

## 为什么需要这个东西

用户报的现象（原话）：

    我用 workbuddy 做其他的 vba 代码，不知道为什么，会搞崩这个项目，而且无法
    启动，只有重启电脑，项目才会生效。

拆开是两件事，都得治：

  1. **"搞崩了"** —— 我们的主线程在 100ms 一轮的轮询里要打 COM（`GetSelection`
     / `Lines`）。COM 调用是**跨进程的同步调用**，没有超时参数：宿主 Excel 正
     被别的程序（WorkBuddy 的 VBA 自动化、一个弹出模态框的宏）占着主线程时，
     这一行会**无限期阻塞**。主线程一堵，tkinter 的消息循环就不再转，全局键盘
     钩子也没人服务 —— 表现就是"全系统还能打字，但这个工具彻底不响应"，而且
     窗也关不掉，只能在任务管理器里结束进程。
  2. **"无法启动"** —— 进程其实**没死透**（还挂着单实例互斥体）。用户再双击
     `run.bat`，`main()` 里 `acquire_single_instance()` 发现互斥体已存在，直接
     `return` —— 于是"双击了、什么也没发生"。Windows 只在进程**终止**时释放
     互斥体，所以只有重启电脑（或手动杀掉那个僵尸进程）才能恢复。

本模块只负责"心跳"这一半：主线程每转一轮就往 `runtime.json` 写一次时间戳。
内容里带 `pid` 与**进程创建时间**（`ctime`，FILETIME 100ns）—— 后者是给守护
进程防"pid 复用误杀"用的：Windows 的 pid 会被复用，只凭 pid 杀进程，理论上有
可能杀掉一个毫不相干的新进程。

写盘用"临时文件 + `os.replace`"，读的一方永远看不到半截 JSON。写入频率由
`tick()` 自己限流（默认 1 秒一次），**不是**每 100ms 写一次盘。

本模块**只依赖标准库**（+ settings 只为拿项目根目录）：它要在"主程序已经半死"
的场景下仍然可用，不能有任何额外依赖。
"""

import json
import os
import time

import settings

# 心跳文件位置：默认就在项目文件夹（与 `recent.json` / `snippets.db` 同级）。
# 测试用 `VBECOMPLETE_RUNTIME_FILE` 指到临时目录 —— 刻意**不做成 settings 的
# 公开开关**：这是给守护进程看的一个内部文件，用户没有理由改它。
ENV_FILE = "VBECOMPLETE_RUNTIME_FILE"
FILE_NAME = "runtime.json"

# 写入间隔。1 秒足够快：判"卡死"的门槛是 20 秒级（见 HANG_SEC），差这几百
# 毫秒毫无意义；而 100ms 写一次盘 = 每秒 10 次 fopen/replace，纯属浪费。
MIN_WRITE_INTERVAL = 1.0

# 心跳里的版本号：将来改字段时，守护进程可以据此判断"这份文件是不是我认识的
# 格式"，认不出就当成过期（宁可多重启一次，也别当成新鲜的然后永远不救）。
FORMAT = 128

_last_write = 0.0
_boot_id = ""


def state_path():
    """心跳文件的绝对路径（测试可用环境变量覆盖）。"""
    try:
        p = (os.environ.get(ENV_FILE) or "").strip()
    except Exception:
        p = ""
    if p:
        return p
    return os.path.join(settings.PROJECT_DIR, FILE_NAME)


def process_ctime(pid):
    """进程创建时间（FILETIME，100ns 整数）。拿不到返回 None。

    为什么不直接用 `psutil`：守护进程要能在最坏情况下工作，多一个三方依赖
    就多一处"起不来"。Win32 的 `GetProcessTimes` 是标准库 `ctypes` 就够的。
    """
    try:
        pid = int(pid)
    except Exception:
        return None
    if pid <= 0:
        return None
    try:
        import ctypes
        from ctypes import wintypes

        class FILETIME(ctypes.Structure):
            _fields_ = [("dwLowDateTime", wintypes.DWORD),
                        ("dwHighDateTime", wintypes.DWORD)]

        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        k32.OpenProcess.restype = wintypes.HANDLE
        k32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL,
                                    wintypes.DWORD]
        # PROCESS_QUERY_LIMITED_INFORMATION：Vista 起可用，权限要求最低
        h = k32.OpenProcess(0x1000, False, pid)
        if not h:
            return None
        try:
            ft_create = FILETIME()
            ft_exit = FILETIME()
            kt = FILETIME()
            ut = FILETIME()
            ok = k32.GetProcessTimes(h, ctypes.byref(ft_create),
                                     ctypes.byref(ft_exit),
                                     ctypes.byref(kt), ctypes.byref(ut))
            if not ok:
                return None
            return (int(ft_create.dwHighDateTime) << 32) | int(
                ft_create.dwLowDateTime)
        finally:
            k32.CloseHandle(h)
    except Exception:
        return None


def process_alive(pid):
    """这个 pid 现在是"还在跑"吗？

    True / False / **None**（问不出来 —— 权限不足等）。
    None 与 False 必须分开：只有 False 才允许当作"进程没了、需要拉起"，
    None 时宁可什么都不做（拿不准就少做，是本项目一贯的口径）。
    """
    try:
        pid = int(pid)
    except Exception:
        return False
    if pid <= 0:
        return False
    try:
        import ctypes
        from ctypes import wintypes

        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        k32.OpenProcess.restype = wintypes.HANDLE
        k32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL,
                                    wintypes.DWORD]
        # SYNCHRONIZE 就够了；再要 QUERY_LIMITED_INFORMATION 是为了
        # GetExitCodeProcess（其实 SYNCHRONIZE 已足够，但两者都要最稳）
        h = k32.OpenProcess(0x00100000 | 0x1000, False, pid)
        if not h:
            # 5 = ERROR_ACCESS_DENIED：进程存在，只是不让我们开（提升权限的
            # 进程）。87 = ERROR_INVALID_PARAMETER：pid 不存在。
            err = ctypes.get_last_error()
            if err == 5:
                return None
            return False
        try:
            code = wintypes.DWORD(0)
            if not k32.GetExitCodeProcess(h, ctypes.byref(code)):
                return None
            return int(code.value) == 259      # STILL_ACTIVE
        finally:
            k32.CloseHandle(h)
    except Exception:
        return None


def pid_of(hb):
    """从心跳内容里取 pid（取不出 -> 0）。"""
    try:
        return int((hb or {}).get("pid") or 0)
    except Exception:
        return 0


def age(hb, now=None):
    """心跳距今多少秒；内容不合法返回 None。

    ★时钟被往回拨时返回 **0**（当成刚写过）—— 否则一次对时就会把还在好好
    干活的主程序当成卡死、拉去重杀一遍。
    """
    try:
        ts = float((hb or {}).get("ts") or 0.0)
    except Exception:
        return None
    if ts <= 0:
        return None
    if now is None:
        now = time.time()
    return max(0.0, float(now) - ts)


def is_fresh(hb, now=None, max_age=20.0):
    """心跳是否够新（纯函数，测试直接调它）。

    `hb` 为 None / 字段缺失 / 格式版本认不出 -> False（当作需要救）。
    """
    if not isinstance(hb, dict):
        return False
    try:
        if int(hb.get("ver") or 0) != FORMAT:
            return False
    except Exception:
        return False
    a = age(hb, now)
    if a is None:
        return False
    return a <= float(max_age)


def read(path=None):
    """读心跳；读不到 / 内容坏了返回 None（绝不抛异常）。"""
    p = path or state_path()
    try:
        with open(p, "r", encoding="utf-8") as f:
            obj = json.load(f)
        return obj if isinstance(obj, dict) else None
    except Exception:
        return None


def write(note="", path=None, pid=None, now=None):
    """写一次心跳（原子替换）。成功 True，失败 False —— **绝不抛异常**。

    写失败（目标正被守护进程读、磁盘满、只读目录…）完全不影响主功能：下一次
    `tick()` 1 秒后还会再写。所以这里 swallow 掉所有异常。
    """
    global _last_write, _boot_id
    p = path or state_path()
    if not _boot_id:
        _boot_id = time.strftime("%Y-%m-%dT%H:%M:%S")
    now = time.time() if now is None else float(now)
    try:
        pid = int(pid) if pid is not None else os.getpid()
    except Exception:
        return False
    data = {
        "ver": FORMAT,
        "pid": pid,
        "ctime": process_ctime(pid),
        "ts": now,
        "boot": _boot_id,
    }
    if note:
        data["note"] = str(note)[:120]
    tmp = p + ".tmp"
    try:
        d = os.path.dirname(p)
        if d and not os.path.isdir(d):
            os.makedirs(d, exist_ok=True)
        with open(tmp, "w", encoding="utf-8", newline="\n") as f:
            json.dump(data, f, ensure_ascii=False)
            f.flush()
            try:
                os.fsync(f.fileno())
            except Exception:
                pass
        # os.replace 在 Windows 上是原子的；但目标若正被别人开着（没带
        # FILE_SHARE_DELETE）会抛 PermissionError —— 丢掉这一次即可。
        os.replace(tmp, p)
        _last_write = now
        return True
    except Exception:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass
        return False


def tick(now=None, note="", force=False, path=None):
    """主线程每轮调一次；内部按 MIN_WRITE_INTERVAL 限流。

    返回"这一轮是否真的写了盘"，方便测试与日志（正常情况下每 10 轮才写 1 次）。
    """
    now = time.time() if now is None else float(now)
    if not force and (now - _last_write) < MIN_WRITE_INTERVAL:
        return False
    return write(note=note, path=path, now=now)


def clear(path=None):
    """删掉心跳文件（测试与"我要干净地停掉"时用）。"""
    p = path or state_path()
    try:
        if os.path.exists(p):
            os.remove(p)
            return True
    except Exception:
        pass
    return False
