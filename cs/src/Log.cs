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
    }
}
