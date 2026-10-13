# -*- coding: utf-8 -*-
"""v108 探针：看看本机输入法（当前装的是 00000804 = 微软拼音）在
`ImmGetConversionStatus` / `WM_IME_CONTROL` 里到底报了什么「转换模式」位。

目的只有一个：确认「中文标点」这个状态对应的是哪一位
（IME_CMODE_SYMBOL 0x0400 / IME_CMODE_FULLSHAPE 0x0008 / ...），
好让 main._ime_chinese_punct() 不做猜测。

跑法：py -3 tests/probe_ime_mode.py
只读、不改任何输入法状态。（顺带验证"跨进程问 IME 窗口"这条路通不通。）
"""
import ctypes
import ctypes.wintypes as wt

u32 = ctypes.windll.user32
imm = ctypes.windll.imm32

imm.ImmGetContext.restype = ctypes.c_void_p
imm.ImmGetContext.argtypes = [ctypes.c_void_p]
imm.ImmReleaseContext.restype = wt.BOOL
imm.ImmReleaseContext.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
imm.ImmGetConversionStatus.restype = wt.BOOL
imm.ImmGetConversionStatus.argtypes = [ctypes.c_void_p,
                                       ctypes.POINTER(wt.DWORD),
                                       ctypes.POINTER(wt.DWORD)]
imm.ImmGetDefaultIMEWnd.restype = ctypes.c_void_p
imm.ImmGetDefaultIMEWnd.argtypes = [ctypes.c_void_p]
imm.ImmIsIME.restype = wt.BOOL
imm.ImmIsIME.argtypes = [ctypes.c_void_p]

u32.GetForegroundWindow.restype = ctypes.c_void_p
u32.SendMessageTimeoutW.restype = ctypes.c_void_p
u32.SendMessageTimeoutW.argtypes = [ctypes.c_void_p, ctypes.c_uint,
                                    ctypes.c_void_p, ctypes.c_void_p,
                                    ctypes.c_uint, ctypes.c_uint,
                                    ctypes.POINTER(ctypes.c_size_t)]

WM_IME_CONTROL = 0x0283
IMC_GETCONVERSIONMODE = 0x0001
SMTO_ABORTIFHUNG = 0x0002

FLAGS = [(0x0001, "IME_CMODE_NATIVE(中文)"),
         (0x0002, "IME_CMODE_KATAKANA"),
         (0x0008, "IME_CMODE_FULLSHAPE(全角)"),
         (0x0010, "IME_CMODE_ROMAN"),
         (0x0020, "IME_CMODE_CHARCODE"),
         (0x0040, "IME_CMODE_HANJACONVERT"),
         (0x0080, "IME_CMODE_SOFTKBD/NATIVESYMBOL(0x80)"),
         (0x0100, "IME_CMODE_NOCONVERSION"),
         (0x0200, "IME_CMODE_EUDC"),
         (0x0400, "IME_CMODE_SYMBOL(中文标点)"),
         (0x0800, "IME_CMODE_FIXED")]


def decode(m):
    on = [nm for bit, nm in FLAGS if m & bit]
    return "0x%04X  %s" % (m, " | ".join(on) if on else "(全 0 / 半角英文)")


def main():
    import win32gui

    hkl = u32.GetKeyboardLayout(0)
    print("当前线程键盘布局 HKL = 0x%08X  ImmIsIME=%s"
          % (int(hkl), bool(imm.ImmIsIME(ctypes.c_void_p(int(hkl))))))

    # --- A) 自己建一个窗口，拿【本线程】的 IME 上下文读转换模式 ---
    hwnd = win32gui.CreateWindow("STATIC", "vbe_ime_probe", 0,
                                 0, 0, 8, 8, 0, 0, 0, None)
    print("探针窗口 hwnd = %s" % (hwnd,))
    himc = imm.ImmGetContext(ctypes.c_void_p(hwnd))
    if himc:
        conv = wt.DWORD(0)
        sent = wt.DWORD(0)
        ok = imm.ImmGetConversionStatus(himc, ctypes.byref(conv),
                                        ctypes.byref(sent))
        print("[A] 本线程 ImmGetContext+ImmGetConversionStatus ok=%s -> %s"
              % (bool(ok), decode(int(conv.value))))
        print("    sentence=0x%04X" % (int(sent.value),))
        imm.ImmReleaseContext(ctypes.c_void_p(hwnd), himc)
    else:
        print("[A] ImmGetContext 返回 NULL")

    # --- B) 跨进程：前台窗口 -> ImmGetDefaultIMEWnd -> WM_IME_CONTROL ---
    fg = int(u32.GetForegroundWindow() or 0)
    print("[B] 前台窗口 hwnd = 0x%X" % fg)
    if fg:
        iw = imm.ImmGetDefaultIMEWnd(ctypes.c_void_p(fg))
        print("    ImmGetDefaultIMEWnd = %s" % (iw,))
        if iw:
            res = ctypes.c_size_t(0)
            got = u32.SendMessageTimeoutW(
                iw, WM_IME_CONTROL, IMC_GETCONVERSIONMODE, 0,
                SMTO_ABORTIFHUNG, 100, ctypes.byref(res))
            print("    WM_IME_CONTROL/IMC_GETCONVERSIONMODE got=%s -> %s"
                  % (bool(got), decode(int(res.value) & 0xFFFF) if got
                     else "无返回"))

    # --- C) 跨进程直接 ImmGetContext(前台窗口) 能不能读到？ ---
    print("[C] 跨进程 ImmGetContext(前台):")
    if fg:
        himc2 = imm.ImmGetContext(ctypes.c_void_p(fg))
        print("    ImmGetContext = %s" % (himc2,))
        if himc2:
            c2 = wt.DWORD(0)
            s2 = wt.DWORD(0)
            ok2 = imm.ImmGetConversionStatus(himc2, ctypes.byref(c2),
                                             ctypes.byref(s2))
            print("    ok=%s -> %s" % (bool(ok2), decode(int(c2.value))))
            imm.ImmReleaseContext(ctypes.c_void_p(fg), himc2)
    win32gui.DestroyWindow(hwnd)


if __name__ == "__main__":
    main()
