# -*- coding: utf-8 -*-
"""钩子门控实证探针（v109 排查「开着 VbeComplete 时 PyCharm 项目关不掉」用）。

随仓库发布的手工探针（与 tests/probe_ime_mode.py 同一个性质）：
只读、不改产品代码、不碰任何真实窗口，跑完自动结束。

它确认两件事：
  A. 微实验：对"wndproc 里睡 N 秒"的窗口发 SendMessageTimeout(WM_NULL, 200ms)，
     到底阻塞调用线程多久？（同线程直调时 uTimeout 形同虚设，会等满对方处理完）
  B. 插桩跑【真实的】main.py + 【真实的】pynput 钩子，用假窗口真实切换前台，
     其中一个假窗口可进入"卡死"模式 —— 看主线程卡在哪个调用里、卡多久，
     以及钩子挂载/卸载的时间线。

v109 的结论（本探针跑出来的）：
  * `in_vbe_code_pane()` 里的 `GetWindowText(前台窗口)` 是跨线程同步调用，
    前台窗口正忙时会把工具主线程堵 1.484 秒以上 ⇒ reconcile 停摆 ⇒
    全局钩子卸不掉 ⇒ 别的程序退出被拖住（「PyCharm 关不掉」的根因）。
  * 修法：读标题前先做 200ms 封顶的 closing 预检；reconcile 外壳异常不死。

跑法：py -3 tests/probe_hook_gate.py
"""
import ctypes
import os
import sys
import threading
import time

PROBE_LOG = r"D:\codes\VbeComplete-V587\.workbuddy\hook_probe2.log"
T0 = time.time()


def plog(msg):
    with open(PROBE_LOG, "a", encoding="utf-8") as f:
        f.write("[%7.3f] %s\n" % (time.time() - T0, msg))


# ---------- A. 微实验 ----------
import win32api                                        # noqa: E402
import win32con                                        # noqa: E402
import win32gui                                        # noqa: E402
import win32process                                    # noqa: E402

_SLEEP = [0.0]


def _wp(hwnd, msg, wparam, lparam):
    if msg == win32con.WM_NULL and _SLEEP[0] > 0:
        time.sleep(_SLEEP[0])
        return 0
    if msg == win32con.WM_DESTROY:
        win32gui.PostQuitMessage(0)
        return 0
    return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)


_hinst = win32api.GetModuleHandle(None)
_wc = win32gui.WNDCLASS()
_wc.lpfnWndProc = _wp
_wc.hInstance = _hinst
_wc.lpszClassName = "ProbeMicro"
win32gui.RegisterClass(_wc)
_micro = win32gui.CreateWindowEx(
    0, "ProbeMicro", "micro", win32con.WS_OVERLAPPEDWINDOW,
    100, 100, 300, 200, 0, 0, _hinst, None)
win32gui.ShowWindow(_micro, win32con.SW_SHOW)

plog("=== A. SendMessageTimeout 微实验 ===")
for sl in (0.0, 0.3, 2.0):
    _SLEEP[0] = sl
    win32gui.PumpWaitingMessages()
    res = ctypes.c_ulong(0)
    t = time.time()
    ok = ctypes.windll.user32.SendMessageTimeoutW(
        ctypes.c_void_p(_micro), 0x0000, 0, 0,
        0x0002 | 0x0001, 200, ctypes.byref(res))
    plog("  wndproc睡%.1fs -> SendMessageTimeout(200ms) 返回=%r 耗时=%.3fs"
         % (sl, ok, time.time() - t))
    win32gui.PumpWaitingMessages()
_SLEEP[0] = 0.0

# ---------- B. 插桩跑真实 main ----------
import pynput.keyboard as _pk                            # noqa: E402
import pynput.mouse as _pm                               # noqa: E402

_RealK, _RealM = _pk.Listener, _pm.Listener
_hs = {"kbd": False, "mouse": False}


class _K(_RealK):
    def start(self):
        _hs["kbd"] = True
        plog(">>> kbd.start")
        return _RealK.start(self)

    def stop(self):
        _hs["kbd"] = False
        plog(">>> kbd.stop")
        return _RealK.stop(self)


class _M(_RealM):
    def start(self):
        _hs["mouse"] = True
        plog(">>> mouse.start")
        return _RealM.start(self)

    def stop(self):
        _hs["mouse"] = False
        plog(">>> mouse.stop")
        return _RealM.stop(self)


_pk.Listener, _pm.Listener = _K, _M

os.environ["VBECOMPLETE_LOG"] = "1"
os.environ.setdefault("VBECOMPLETE_LOG_FILE",
                      r"D:\codes\VbeComplete-V587\.workbuddy\probe_main2.log")
sys.path.insert(0, r"D:\codes\VbeComplete-V587")
import main as MN                                        # noqa: E402

MN.acquire_single_instance = lambda: True

_orig_closing = MN.foreground_window_closing
_orig_recon = MN._reconcile_resources
_orig_invbe = MN.in_vbe_code_pane


def _timed_closing():
    t = time.time()
    r = _orig_closing()
    plog("  [主线程] foreground_window_closing() -> %r 耗时=%.3fs"
         % (r, time.time() - t))
    return r


def _timed_invbe(closing=None):
    t = time.time()
    r = _orig_invbe(closing=closing)
    plog("  [主线程] in_vbe_code_pane(closing=%r) -> %r 耗时=%.3fs"
         % (closing, r, time.time() - t))
    return r


_orig_gwt = win32gui.GetWindowText


def _timed_gwt(hwnd):
    t = time.time()
    r = _orig_gwt(hwnd)
    plog("  [主线程] GetWindowText(hwnd=%r) -> %r 耗时=%.3fs"
         % (hwnd, r, time.time() - t))
    return r


win32gui.GetWindowText = _timed_gwt      # 工具里 _gui 就是 win32gui 模块本体


def _timed_recon(*a, **kw):
    t = time.time()
    r = _orig_recon(*a, **kw)
    plog("  [主线程] _reconcile_resources(focused=%r closing=%r suppress=%r)"
         " -> mount=%r 耗时=%.3fs"
         % (a[0] if a else kw.get("focused"),
            kw.get("closing", a[6] if len(a) > 6 else None),
            kw.get("suppress", a[7] if len(a) > 7 else None),
            r[0], time.time() - t))
    return r


MN.foreground_window_closing = _timed_closing
MN.in_vbe_code_pane = _timed_invbe
MN._reconcile_resources = _timed_recon

plog("=== B. 后台跑真实 main() ===")
threading.Thread(target=MN.main, daemon=True).start()
time.sleep(2.0)


def force_foreground(hwnd):
    fg = win32gui.GetForegroundWindow()
    my = ctypes.windll.kernel32.GetCurrentThreadId()
    tid = win32process.GetWindowThreadProcessId(fg)[0] if fg else 0
    if tid and tid != my:
        ctypes.windll.user32.AttachThreadInput(my, tid, True)
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    finally:
        if tid and tid != my:
            ctypes.windll.user32.AttachThreadInput(my, tid, False)
    plog("前台 -> %r" % win32gui.GetWindowText(win32gui.GetForegroundWindow()))


def pump(t):
    end = time.time() + t
    while time.time() < end:
        win32gui.PumpWaitingMessages()
        time.sleep(0.02)


# 卡死窗口（独立的第二个类，避免和微实验窗口互相干扰）
_SLEEP2 = [0.0]


def _wp2(hwnd, msg, wparam, lparam):
    if msg == win32con.WM_NULL and _SLEEP2[0] > 0:
        time.sleep(_SLEEP2[0])
        return 0
    if msg == win32con.WM_DESTROY:
        win32gui.PostQuitMessage(0)
        return 0
    return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)


_wc2 = win32gui.WNDCLASS()
_wc2.lpfnWndProc = _wp2
_wc2.hInstance = _hinst
_wc2.lpszClassName = "ProbeHung"
win32gui.RegisterClass(_wc2)
vbe = win32gui.CreateWindowEx(0, "ProbeHung", "Microsoft Visual Basic - 探针VBE",
                              win32con.WS_OVERLAPPEDWINDOW, 200, 200, 500, 380,
                              0, 0, _hinst, None)
ide = win32gui.CreateWindowEx(0, "ProbeHung", "探针 - PyCharm",
                              win32con.WS_OVERLAPPEDWINDOW, 260, 260, 500, 380,
                              0, 0, _hinst, None)
win32gui.ShowWindow(vbe, win32con.SW_SHOW)
win32gui.ShowWindow(ide, win32con.SW_SHOW)

plog("--- 场景1：正常切换（基线）---")
force_foreground(vbe)
pump(1.5)
force_foreground(ide)
pump(1.5)

plog("--- 场景2：VBE 中 -> 切到【忙 2 秒】的 IDE（未卡死，只是慢）---")
force_foreground(vbe)
pump(1.5)
_SLEEP2[0] = 2.0          # WM_NULL 里睡 2 秒：窗口"忙但没死"
force_foreground(ide)
plog("  已切到忙窗口；开始计时等钩子卸载（探测主线程阻塞）")
t = time.time()
while _hs["kbd"] and time.time() - t < 15:
    # 用 WaitMessage 式空转，避免自己的线程被 WM_NULL 拖住
    time.sleep(0.02)
plog("  钩子卸载耗时=%.2fs" % (time.time() - t))
pump(1.0)

plog("=== 结束 ===")
_SLEEP2[0] = 0.0
for h in (vbe, ide, _micro):
    try:
        win32gui.DestroyWindow(h)
    except Exception:
        pass
