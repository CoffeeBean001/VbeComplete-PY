// COM 接入层：拿到正在运行的 VBE 对象，并把它缓存起来。
//
// 为什么不用 Microsoft.Office.Interop.*：那套程序集绑定具体 Excel 版本，
// 而且会让"装的是 WPS / 只有 Word / 换了 Office 版本"统统编译期就绑死。
// 这里全部走 IDispatch 后期绑定（C# 的 dynamic），与 Python 版 comtypes 同一路子。
//
// ★三条硬规矩（与 Python 版一致）：
//   1. 只有 UI 线程（STA）能调这些方法，钩子线程一次都不许碰；
//   2. 宿主不在时立刻退让并进入退避，绝不反复试 —— 每次失败的 COM 调用
//      在宿主忙的时候会阻塞到超时，那是"卡死"的来源；
//   3. 缓存代理对象，但每次使用前先探活；探活失败立刻丢弃缓存。

using System;
using System.Runtime.InteropServices;

namespace VbeComplete
{
    internal static class ComCtx
    {
        // ---- oleaut32!GetActiveObject（.NET 5+ 把这个 API 从 Marshal 上删了） ----
        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        private static extern int CLSIDFromProgID(string progId, out Guid pclsid);

        /// <summary>按 ProgID 取正在运行的实例；没有实例返回 null（绝不启动新实例）。</summary>
        public static object GetActive(string progId)
        {
            Guid clsid;
            if (CLSIDFromProgID(progId, out clsid) != 0) return null;
            try
            {
                object o;
                GetActiveObject(ref clsid, IntPtr.Zero, out o);
                return o;
            }
            catch (Exception)
            {
                // MK_E_UNAVAILABLE(0x800401E3) = 宿主没在跑，这是正常情况，不是错误
                return null;
            }
        }

        // ---- 宿主探测 ----
        // 先按前台窗口的进程名猜 ProgID，再退回上次成功的那个，最后按固定顺序试。
        // ★顺序很重要：宿主探测必须排在【退避检查之前】—— 否则"宿主回来了"
        //   会被上一次的退避挡死，表现为启动慢/永远连不上。
        private static readonly (string proc, string progid)[] Hosts =
        {
            ("EXCEL",   "Excel.Application"),
            ("WINWORD", "Word.Application"),
            ("POWERPNT","PowerPoint.Application"),
            ("wpp",     "KWPS.Application"),      // WPS 表格/文字
            ("et",      "KET.Application"),
            ("wps",     "KWPS.Application"),
        };

        private static string _lastOk = "";

        public static string GuessProgId()
        {
            IntPtr fg = Native.GetForegroundWindow();
            uint pid = Native.PidOf(fg);
            string proc = Native.ProcessNameOfPid(pid);
            if (!string.IsNullOrEmpty(proc))
            {
                foreach (var h in Hosts)
                    if (string.Equals(h.proc, proc, StringComparison.OrdinalIgnoreCase))
                        return h.progid;
            }
            if (!string.IsNullOrEmpty(_lastOk)) return _lastOk;
            return "Excel.Application";
        }

        public static void NoteHost(string progid) => _lastOk = progid;

        private static object _app;
        private static object _vbe;
        private static DateTime _lastVerify = DateTime.MinValue;

        /// <summary>丢弃缓存（VBE 关掉了 / 报错之后必须调）。</summary>
        public static void Reset()
        {
            _vbe = null;
            _app = null;
            _lastVerify = DateTime.MinValue;
        }

        /// <summary>
        /// 拿 VBE 对象。verify=true 时顺带探一次 ActiveCodePane（更慢但更准）。
        /// 返回 null 表示"现在拿不到"，调用方必须安静退出。
        /// </summary>
        public static dynamic Vbe(bool verify = true)
        {
            if (_vbe != null)
            {
                if (!verify) return _vbe;
                // 200ms 内的重复探活没有意义，直接给
                if ((DateTime.UtcNow - _lastVerify).TotalMilliseconds < 200) return _vbe;
                _lastVerify = DateTime.UtcNow;
                try
                {
                    object cp = ((dynamic)_vbe).ActiveCodePane;
                    if (cp != null) return _vbe;
                }
                catch { }
                Reset();
            }

            string progid = GuessProgId();
            object app = GetActive(progid);
            if (app == null && progid != "Excel.Application")
                app = GetActive("Excel.Application");
            if (app == null) return null;

            try
            {
                dynamic dapp = app;
                dynamic vbe = dapp.VBE;
                if (ReferenceEquals(vbe, null)) return null;
                _app = app;
                _vbe = vbe;
                _lastVerify = DateTime.UtcNow;
                NoteHost(progid);
                Log.Write("COM 已连接：{0}", progid);
                return _vbe;
            }
            catch (Exception e)
            {
                Log.Write("取 VBE 失败：{0}", e.Message);
                Reset();
                return null;
            }
        }

        public static bool Available => _vbe != null;
    }
}
