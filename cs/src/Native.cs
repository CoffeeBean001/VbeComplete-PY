// Win32 层：全局键盘钩子、按键注入、窗口/进程探测。
//
// ★这是"钩子线程"能碰的全部东西 —— 只有纯 Win32 和本进程内存。
//   钩子回调里【绝不允许】出现任何 COM 调用：VBE 可能正忙，一次阻塞就会
//   让整个系统的键盘卡住（低级钩子在输入队列里串行执行）。

using System;
using System.Runtime.InteropServices;
using System.Text;

namespace VbeComplete
{
    internal static class Native
    {
        // ---- 虚拟键码（只列本项目用到的） ----
        public const int VK_BACK = 0x08;
        public const int VK_TAB = 0x09;
        public const int VK_RETURN = 0x0D;
        public const int VK_ESCAPE = 0x1B;
        public const int VK_SPACE = 0x20;
        public const int VK_PRIOR = 0x21;
        public const int VK_NEXT = 0x22;
        public const int VK_END = 0x23;
        public const int VK_HOME = 0x24;
        public const int VK_LEFT = 0x25;
        public const int VK_UP = 0x26;
        public const int VK_RIGHT = 0x27;
        public const int VK_DOWN = 0x28;
        public const int VK_DELETE = 0x2E;
        public const int VK_OEM_2 = 0xBF;   // /
        public const int VK_OEM_5 = 0xDC;   // \
        public const int VK_SHIFT = 0x10;
        public const int VK_CONTROL = 0x11;
        public const int VK_MENU = 0x12;    // Alt
        public const int VK_LWIN = 0x5B;
        public const int VK_RWIN = 0x5C;

        // ---- 钩子 ----
        public const int WH_KEYBOARD_LL = 13;
        public const int HC_ACTION = 0;
        public const int WM_KEYDOWN = 0x0100;
        public const int WM_KEYUP = 0x0101;
        public const int WM_SYSKEYDOWN = 0x0104;
        public const int WM_SYSKEYUP = 0x0105;

        // KBDLLHOOKSTRUCT.flags
        public const uint LLKHF_EXTENDED = 0x01;
        public const uint LLKHF_INJECTED = 0x10;
        public const uint LLKHF_UP = 0x80;

        [StructLayout(LayoutKind.Sequential)]
        public struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        public delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn,
            IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr GetModuleHandle(string lpModuleName);

        // ---- 消息循环 ----
        [StructLayout(LayoutKind.Sequential)]
        public struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public int ptX;
            public int ptY;
        }

        [DllImport("user32.dll")]
        public static extern int GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool TranslateMessage(ref MSG lpMsg);

        [DllImport("user32.dll")]
        public static extern IntPtr DispatchMessage(ref MSG lpMsg);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool PostThreadMessage(uint idThread, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        public const uint WM_QUIT = 0x0012;
        public const uint WM_APP_QUIT = 0x8000 + 17;

        // ---- 按键状态 ----
        [DllImport("user32.dll")]
        public static extern short GetAsyncKeyState(int vKey);

        public static bool Down(int vk) => (GetAsyncKeyState(vk) & 0x8000) != 0;
        public static bool Toggled(int vk) => (GetAsyncKeyState(vk) & 0x0001) != 0;

        // ---- 窗口 / 进程 ----
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left, Top, Right, Bottom;
            public int Width => Right - Left;
            public int Height => Bottom - Top;
        }

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        public const uint SWP_NOSIZE = 0x0001;
        public const uint SWP_NOMOVE = 0x0002;
        public const uint SWP_NOZORDER = 0x0004;
        public const uint SWP_NOACTIVATE = 0x0010;
        public const uint SWP_SHOWWINDOW = 0x0040;

        public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);

        [DllImport("user32.dll")]
        public static extern int GetSystemMetrics(int nIndex);

        public const int SM_XVIRTUALSCREEN = 76;
        public const int SM_YVIRTUALSCREEN = 77;
        public const int SM_CXVIRTUALSCREEN = 78;
        public const int SM_CYVIRTUALSCREEN = 79;

        [DllImport("user32.dll")]
        public static extern uint GetDpiForWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(POINT p);

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT { public int X, Y; }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetCursorPos(out POINT lpPoint);

        // ---- 按键注入 ----
        [StructLayout(LayoutKind.Sequential)]
        public struct MOUSEINPUT
        {
            public int dx, dy;
            public uint mouseData, dwFlags, time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct KEYBDINPUT
        {
            public ushort wVk, wScan;
            public uint dwFlags, time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL, wParamH;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct INPUTUNION
        {
            [FieldOffset(0)] public MOUSEINPUT mi;
            [FieldOffset(0)] public KEYBDINPUT ki;
            [FieldOffset(0)] public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct INPUT
        {
            public uint type;
            public INPUTUNION u;
        }

        public const uint INPUT_KEYBOARD = 1;
        public const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
        public const uint KEYEVENTF_KEYUP = 0x0002;
        public const uint KEYEVENTF_UNICODE = 0x0004;

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        /// <summary>注入一个 Unicode 字符（绕过输入法，直接进 WM_CHAR）。</summary>
        public static bool SendUnicodeChar(char ch)
        {
            INPUT[] seq = new INPUT[2];
            seq[0].type = INPUT_KEYBOARD;
            seq[0].u.ki.wVk = 0;
            seq[0].u.ki.wScan = ch;
            seq[0].u.ki.dwFlags = KEYEVENTF_UNICODE;
            seq[1] = seq[0];
            seq[1].u.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
            return SendInput(2, seq, Marshal.SizeOf(typeof(INPUT))) == 2;
        }

        /// <summary>注入一个虚拟键（不做修饰键合成，调用方自己决定按不按修饰键）。</summary>
        public static bool SendVk(int vk, bool ctrl = false, bool shift = false, bool alt = false)
        {
            int n = 0;
            INPUT[] seq = new INPUT[8];

            if (ctrl) { seq[n++] = VkInput(VK_CONTROL, false); }
            if (shift) { seq[n++] = VkInput(VK_SHIFT, false); }
            if (alt) { seq[n++] = VkInput(VK_MENU, false); }
            seq[n++] = VkInput(vk, false);
            seq[n++] = VkInput(vk, true);
            if (alt) { seq[n++] = VkInput(VK_MENU, true); }
            if (shift) { seq[n++] = VkInput(VK_SHIFT, true); }
            if (ctrl) { seq[n++] = VkInput(VK_CONTROL, true); }

            INPUT[] exact = new INPUT[n];
            Array.Copy(seq, exact, n);
            return SendInput((uint)n, exact, Marshal.SizeOf(typeof(INPUT))) == (uint)n;
        }

        private static INPUT VkInput(int vk, bool up)
        {
            INPUT it = new INPUT();
            it.type = INPUT_KEYBOARD;
            it.u.ki.wVk = (ushort)vk;
            it.u.ki.wScan = 0;
            it.u.ki.dwFlags = up ? KEYEVENTF_KEYUP : 0;
            // 方向键 / Delete 等需要 EXTENDED 位，否则某些应用收不到
            if (vk == VK_UP || vk == VK_DOWN || vk == VK_LEFT || vk == VK_RIGHT ||
                vk == VK_DELETE || vk == VK_HOME || vk == VK_END ||
                vk == VK_PRIOR || vk == VK_NEXT)
            {
                it.u.ki.dwFlags |= KEYEVENTF_EXTENDEDKEY;
            }
            return it;
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        public static extern IntPtr GetParent(IntPtr hWnd);

        // ---- 线程焦点 / 插入符矩形 ----
        [StructLayout(LayoutKind.Sequential)]
        public struct GUITHREADINFO
        {
            public int cbSize;
            public int flags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public RECT rcCaret;
        }

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO pgui);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

        /// <summary>线程里当前拥有键盘焦点的那个窗口（拿不到返回 Zero）。</summary>
        public static IntPtr FocusedHwndOf(uint tid)
        {
            if (tid == 0) return IntPtr.Zero;
            var gti = new GUITHREADINFO();
            gti.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));
            if (!GetGUIThreadInfo(tid, ref gti)) return IntPtr.Zero;
            return gti.hwndFocus;
        }

        /// <summary>当前插入符的屏幕矩形：(x, top, bottom)。VBE 是自绘的，但它仍然
        /// 维护一个真的 caret（GetGUIThreadInfo.rcCaret），这是唯一可靠的位置来源。</summary>
        public static bool CaretRect(out int x, out int top, out int bottom)
        {
            x = top = bottom = 0;
            IntPtr fg = GetForegroundWindow();
            if (fg == IntPtr.Zero) return false;
            uint pid;
            uint tid = GetWindowThreadProcessId(fg, out pid);
            if (tid == 0) return false;

            var gti = new GUITHREADINFO();
            gti.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));
            if (!GetGUIThreadInfo(tid, ref gti)) return false;
            if (gti.hwndCaret == IntPtr.Zero) return false;

            POINT p1 = new POINT { X = gti.rcCaret.Left, Y = gti.rcCaret.Top };
            POINT p2 = new POINT { X = gti.rcCaret.Left, Y = gti.rcCaret.Bottom };
            if (!ClientToScreen(gti.hwndCaret, ref p1)) return false;
            if (!ClientToScreen(gti.hwndCaret, ref p2)) return false;
            x = p1.X; top = p1.Y; bottom = p2.Y;
            return true;
        }

        public static string ClassNameOf(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return "";
            StringBuilder sb = new StringBuilder(256);
            int n = GetClassName(hwnd, sb, sb.Capacity);
            return n > 0 ? sb.ToString() : "";
        }

        public static uint PidOf(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return 0;
            uint pid;
            GetWindowThreadProcessId(hwnd, out pid);
            return pid;
        }

        public static string ProcessNameOfPid(uint pid)
        {
            if (pid == 0) return "";
            try
            {
                using (var p = System.Diagnostics.Process.GetProcessById((int)pid))
                    return p.ProcessName;
            }
            catch { return ""; }
        }
    }
}
