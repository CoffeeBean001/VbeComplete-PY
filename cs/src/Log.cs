// 日志：只在 VBECOMPLETE_LOG=1 时落盘（与 Python 版同一个日志文件，方便对照）。
// 任何写日志失败都必须被吞掉 —— 日志永远不允许影响功能。

using System;
using System.IO;
using System.Text;

namespace VbeComplete
{
    internal static class Log
    {
        public static readonly bool Enabled = Settings.Flag("LOG", false);
        private static readonly object _lock = new object();
        private static string _path;

        public static string Path_
        {
            get
            {
                if (_path == null)
                    _path = System.IO.Path.Combine(Settings.Home, "VbeComplete.log");
                return _path;
            }
        }

        public static void Write(string msg)
        {
            if (!Enabled) return;
            try
            {
                string line = string.Format("{0:HH:mm:ss.fff} [cs] {1}\r\n",
                    DateTime.Now, msg);
                lock (_lock) File.AppendAllText(Path_, line, Encoding.UTF8);
            }
            catch { }
        }

        public static void Write(string fmt, params object[] args)
        {
            if (!Enabled) return;
            try { Write(string.Format(fmt, args)); } catch { }
        }

        /// <summary>
        /// 崩溃记录：**不受 LOG 开关控制**，永远落盘。
        /// 原因很直接 —— WinExe 没有控制台，异常被 CLR 吞掉时用户只看到
        /// "双击没反应"，而我们手里一点线索都没有。崩溃恰恰是没开日志时最需要留下的东西。
        /// </summary>
        public static void Crash(string where, Exception e)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendFormat("==== {0:yyyy-MM-dd HH:mm:ss.fff} {1} ====\r\n", DateTime.Now, where);
                for (var cur = e; cur != null; cur = cur.InnerException)
                {
                    sb.AppendFormat("{0}: {1}\r\n{2}\r\n", cur.GetType().FullName, cur.Message, cur.StackTrace);
                }
                lock (_lock)
                {
                    File.AppendAllText(System.IO.Path.Combine(Settings.Home, "crash.log"),
                        sb.ToString(), Encoding.UTF8);
                }
            }
            catch { }
        }
    }
}
