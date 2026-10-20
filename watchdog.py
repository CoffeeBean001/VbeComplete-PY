# -*- coding: utf-8 -*-
"""VbeComplete 守护进程（★v128）：主程序死了 / 卡死了，就把它重新拉起来。

用户口径（原话）：

    我用 workbuddy 做其他的 vba 代码，不知道为什么，会搞崩这个项目，而且无法
    启动，只有重启电脑，项目才会生效。能不能帮我加个功能，在项目被搞崩之后，
    会自动重启，并保证项目有效。

## "无法启动"到底是怎么来的

`main()` 里有单实例互斥体（`VbeComplete_SingleInstance_Mutex`）。主线程一旦被
一次**没有超时参数的 COM 调用**卡住（宿主 Excel 被别的自动化占着），进程没死，
互斥体就还挂着。这时再双击 `run.bat`，新实例拿到 `ERROR_ALREADY_EXISTS` 就
`return` 了 —— **双击什么都不会发生**。Windows 只在进程终止时释放互斥体，所以
只能重启电脑。这就是用户那句"只有重启电脑"的全部来历。

所以本进程干两件事：

  1. **它自己不碰 COM、不装钩子、不起窗口** —— 只是一遍遍读 `runtime.json`，
     所以主程序怎么崩都带不走它（连依赖都只用标准库）；
  2. 发现"没有实例" / "实例心跳过期（卡死）" → **先杀干净，再拉一个新的**。
     杀掉僵尸的那一瞬间互斥体就释放了，新实例才起得来。

## 判据只有一个：心跳

`heartbeat.py` 由主线程每轮写一次（限流 1 秒）。它证明的是"**主线程还在转**"，
而不是"进程还在" —— 恰恰是卡死那种情况，进程活着但主线程不动，只查进程会
永远看不见。心跳内容里的 `ctime`（进程创建时间）用来防 pid 复用误杀。

## 三条克制

  * **拿不准就不做事**：`process_alive` 返回 None（问不出来）时一律当"还活着"，
    宁可少救一次，也不乱杀用户的进程。
  * **不反复重启**：连续起不来时按 10s/20s/40s… 退避（上限 5 分钟），并把原因
    写进 `watchdog.log` —— 一个起不来的程序被每 3 秒拉一次，只会把机器拖垮。
  * **默认跟着 `AUTO_RESTART` 走**：设 0 就整个功能关掉（本进程直接退出）。

## 谁拉起我

`main.py` 启动时会确保本进程在跑；`run.bat` / 开机自启也直接指向本文件。
两条路都行：本进程自己也有单实例互斥体（`VbeComplete_Watchdog_Mutex`），
多起几个只会立刻退出，不会变成"每起一个主程序就多一个守护进程"。
"""

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import heartbeat                      # noqa: E402
import settings                       # noqa: E402

PROJECT_DIR = settings.PROJECT_DIR
MAIN_PATH = os.path.join(PROJECT_DIR, "main.py")
LOG_PATH = os.path.join(PROJECT_DIR, "watchdog.log")
LOG_MAX_BYTES = 256 * 1024

TICK_SEC = 3.0            # 多久看一次
START_GRACE_SEC = 15.0    # 刚拉起的这段时间内不判"没起来"（冷启动几秒很正常）
HANG_SLACK_SEC = 10.0     # 心跳过期多久算卡死（比主程序自守的 HANG_SEC 更宽）
START_BACKOFF_SEC = 10.0  # 连续起不来时的起步退避
MAX_BACKOFF_SEC = 300.0   # 退避上限

DETACHED_PROCESS = 0x00000008
CREATE_NO_WINDOW = 0x08000000

# 单实例互斥体句柄（必须保持引用，否则会被 GC 释放导致锁失效）
_mutex_handle = None


# ---------------------------------------------------------------------------
# 纯函数：这一轮该做什么（不碰系统、不读写文件 —— 测试直接调它）
# ---------------------------------------------------------------------------
def decide(hb, now, alive, child_running=False, last_start_ts=0.0,
           grace=START_GRACE_SEC, hang_sec=30.0, slack=HANG_SLACK_SEC):
    """返回 `"ok"` / `"wait"` / `"kill"` / `"start"`。

    hb            心跳内容（None = 从没有过）
    alive         心跳里的那个 pid 现在是否**确实还是同一只**进程在跑
    child_running 我们自己拉起的那只还活着吗（Popen.poll() is None）
    last_start_ts 上一次拉起的时间（0 = 从来没起过）

    ★"wait" 只服务于**启动窗口**：刚拉起的进程还没写出自己的心跳之前，看到
    的仍是上一只留下的旧心跳，这时候既不该杀（新进程还在 import），也不该再
    拉一个（会撞互斥体、白烧 CPU）。窗口一过就按心跳实况判。

    ★判"卡死"的门槛是 `hang_sec + slack`，**故意比主程序自守宽 10 秒**：主程序
    自己的 daemon 线程到点会先硬退出（那条路更快、也更干净），本进程晚一步只
    负责收拾"连自守线程都没了"的残局。两边同时动手没有意义，还会抢着杀。
    """
    starting = bool(child_running) and (now - float(last_start_ts or 0.0)
                                       < float(grace))
    if alive:
        if heartbeat.is_fresh(hb, now, float(hang_sec) + float(slack)):
            return "ok"
        return "wait" if starting else "kill"
    if starting:
        return "wait"
    return "start"


def instance_alive(hb):
    """心跳里那只进程是否还活着**并且确实是同一只**。

    `ctime` 对不上 = pid 被系统复用了（我们那只早死了，现在这个 pid 属于别人）
    ⇒ False。这一条是"绝不误杀无关进程"的最后一道闸。
    """
    pid = heartbeat.pid_of(hb)
    if not pid:
        return False
    alive = heartbeat.process_alive(pid)
    if alive is None:
        return True          # 问不出来：当作还活着（拿不准就不动手）
    if not alive:
        return False
    try:
        want = (hb or {}).get("ctime")
    except Exception:
        want = None
    if want is None:
        return True
    try:
        got = heartbeat.process_ctime(pid)
    except Exception:
        got = None
    if got is None:
        return True          # 读不出来时信 pid
    return int(got) == int(want)


# ---------------------------------------------------------------------------
# 系统动作（只在"确实要杀 / 确实要拉"时才走）
# ---------------------------------------------------------------------------
def kill_pid(pid, ctime=None):
    """强杀一只（自己的）进程。成功 True。

    杀之前**再核一次 ctime** —— 从读到心跳到真正 Terminate 之间也隔了几十
    毫秒，pid 理论上有被复用的窗口。核对不过就放弃（返回 False）。
    """
    try:
        pid = int(pid)
    except Exception:
        return False
    if pid <= 0 or pid == os.getpid():
        return False
    actual = heartbeat.process_ctime(pid)
    if actual is None:
        return False
    if ctime is not None:
        try:
            if int(actual) != int(ctime):
                return False
        except Exception:
            return False
    try:
        import ctypes
        from ctypes import wintypes

        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        k32.OpenProcess.restype = wintypes.HANDLE
        k32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL,
                                    wintypes.DWORD]
        h = k32.OpenProcess(0x0001 | 0x00100000, False, pid)   # TERMINATE|SYNC
        if not h:
            return False
        try:
            ok = k32.TerminateProcess(h, 1)
            if ok:
                k32.WaitForSingleObject(h, 3000)
            return bool(ok)
        finally:
            k32.CloseHandle(h)
    except Exception:
        return False


def pythonw_path():
    """该用哪个解释器拉起主程序：优先同目录的 `pythonw.exe`（不带黑窗）。"""
    exe = sys.executable or ""
    low = exe.lower()
    if low.endswith("python.exe"):
        cand = os.path.join(os.path.dirname(exe), "pythonw.exe")
        if os.path.isfile(cand):
            return cand
    return exe


def spawn_main():
    """拉起 `main.py`（完全脱离本进程）。成功返回 Popen 对象，失败返回 None。"""
    return _spawn(MAIN_PATH)


def _spawn(script):
    """用 pythonw 拉起一个本项目的脚本，**完全脱离**本进程（不留控制台）。"""
    try:
        import subprocess
    except Exception:
        return None
    exe = pythonw_path()
    if not exe or not os.path.isfile(exe) or not os.path.isfile(script):
        return None
    try:
        return subprocess.Popen(
            [exe, script],
            cwd=PROJECT_DIR,
            close_fds=True,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=DETACHED_PROCESS | CREATE_NO_WINDOW)
    except Exception:
        return None


def mutex_exists(name):
    """这个名字的内核互斥体现在是否存在（= 有进程正持有它）。

    True / False / **None**（问不出来）。★v128 修：这里原来写的是
    `ctypes.DWORD` —— 那个名字在 ctypes 里根本不存在，于是每次调用都抛
    AttributeError、被 `except` 吞成 True，`ensure_running()` 便永远认为
    "守护进程已经在跑"、一个都不起（真机实测才发现的：`watchdog.log` 停在旧
    时间戳上，而主程序明明在跑）。所以这一层现在抽成独立函数、由 §119.9 直接调，
    不再靠"看起来没问题"。
    """
    try:
        import ctypes
        from ctypes import wintypes

        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        k32.OpenMutexW.restype = ctypes.c_void_p
        k32.OpenMutexW.argtypes = [wintypes.DWORD, wintypes.BOOL,
                                   wintypes.LPCWSTR]
        h = k32.OpenMutexW(0x00100000, False, name)     # SYNCHRONIZE
        if h:
            k32.CloseHandle(ctypes.c_void_p(h))
            return True
        # 2 = ERROR_FILE_NOT_FOUND：这个名字的内核对象不存在
        return False
    except Exception:
        return None


def watchdog_running():
    """守护进程是否已经在跑（靠它那个单实例互斥体判断）。

    判不出来（None）一律当"在跑" —— 宁可不起第二个，也不要变成"每启动一次主
    程序就多一个守护进程"（多个守护会互相抢着杀/拉起，反而更乱）。
    """
    return mutex_exists("VbeComplete_Watchdog_Mutex") is not False


def ensure_running():
    """（给 main.py 调用）确保守护进程在跑。返回新起的 pid，没起返回 None。"""
    try:
        if not settings.flag("AUTO_RESTART", True):
            return None
    except Exception:
        pass
    if watchdog_running():
        return None
    p = _spawn(os.path.abspath(__file__))
    return getattr(p, "pid", None)


def acquire_single_instance():
    """本进程也只允许一个。多起的那个立刻退出（返回 False）。"""
    global _mutex_handle
    try:
        import ctypes

        ERROR_ALREADY_EXISTS = 183
        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        k32.CreateMutexW.restype = ctypes.c_void_p
        k32.CreateMutexW.argtypes = [ctypes.c_void_p, ctypes.c_bool,
                                     ctypes.c_wchar_p]
        h = k32.CreateMutexW(None, False, "VbeComplete_Watchdog_Mutex")
        if not h:
            return True
        _mutex_handle = h
        return ctypes.get_last_error() != ERROR_ALREADY_EXISTS
    except Exception:
        return True


# ---------------------------------------------------------------------------
# 日志（只有这个进程写，简单即可：单行、带时间、超过 256KB 轮转一份）
# ---------------------------------------------------------------------------
def log(msg):
    """追加一行到 `watchdog.log`。写不成就算了 —— 日志绝不能影响看护。"""
    try:
        size = os.path.getsize(LOG_PATH)
    except Exception:
        size = 0
    try:
        if size > LOG_MAX_BYTES:
            bak = LOG_PATH + ".1"
            try:
                if os.path.exists(bak):
                    os.remove(bak)
                os.replace(LOG_PATH, bak)
            except Exception:
                pass
    except Exception:
        pass
    try:
        with open(LOG_PATH, "a", encoding="utf-8", errors="replace") as f:
            f.write("%s %s\n" % (time.strftime("%Y-%m-%d %H:%M:%S"),
                                 str(msg)[:800]))
    except Exception:
        pass


# ---------------------------------------------------------------------------
def main():
    if not acquire_single_instance():
        return 0
    if not settings.flag("AUTO_RESTART", True):
        log("watchdog: AUTO_RESTART=0，不启动看护")
        return 0
    try:
        hang_sec = max(10, int(settings.int_value("HANG_SEC", 30)))
    except Exception:
        hang_sec = 30

    log("watchdog: 启动 pid=%d 主程序=%s 卡死门槛=%ds"
        % (os.getpid(), MAIN_PATH, hang_sec))

    child = None
    last_start = 0.0
    fail_streak = 0
    told_hint = False

    while True:
        try:
            now = time.time()
            hb = heartbeat.read()
            alive = instance_alive(hb)
            running = bool(child is not None and child.poll() is None)
            act = decide(hb, now, alive, child_running=running,
                         last_start_ts=last_start, hang_sec=hang_sec)

            if act == "ok":
                if fail_streak:
                    log("watchdog: 主程序已恢复正常（pid=%s）"
                        % heartbeat.pid_of(hb))
                fail_streak = 0
                told_hint = False
            elif act == "wait":
                pass
            elif act == "kill":
                pid = heartbeat.pid_of(hb)
                log("watchdog: 心跳已 %.0fs 没更新 -> 判定卡死，强杀 pid=%s"
                    % (heartbeat.age(hb, now) or -1, pid))
                if kill_pid(pid, (hb or {}).get("ctime")):
                    log("watchdog: pid=%s 已终止，互斥体随之释放" % pid)
                    child = None
                    act = "start"
                else:
                    log("watchdog: pid=%s 杀不掉（权限 / 已退出），下轮再看" % pid)
                    act = None
            if act == "start":
                if fail_streak >= 2:
                    back = min(MAX_BACKOFF_SEC,
                               START_BACKOFF_SEC * (2 ** (fail_streak - 1)))
                    log("watchdog: 已连续 %d 次没起来，等 %.0fs 再试"
                        % (fail_streak, back))
                    time.sleep(back)
                    now = time.time()
                child = spawn_main()
                last_start = time.time()
                fail_streak += 1
                if child is None:
                    log("watchdog: 拉起失败（找不到解释器或 main.py）")
                elif child.poll() is not None:
                    log("watchdog: 拉起后立刻退出（退出码 %s）—— 多半是单实例"
                        "互斥体还被占着" % child.returncode)
                    if fail_streak >= 2 and not told_hint:
                        told_hint = True
                        log("watchdog: 提示 —— 若一直起不来，运行 stop.bat "
                            "把残留进程清一次")
                else:
                    log("watchdog: 已拉起 pid=%s" % child.pid)
        except Exception as _e:
            log("watchdog: 循环异常 %s: %s" % (type(_e).__name__, _e))
        try:
            time.sleep(TICK_SEC)
        except Exception:
            pass


if __name__ == "__main__":
    sys.exit(main())
