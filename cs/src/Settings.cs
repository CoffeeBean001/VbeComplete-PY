// 配置中心：与 Python 版【共用同一个 settings.ini】。
//
// 优先级（第一个"能认出来的"赢）：环境变量 > settings.ini > 代码默认值。
// 认不出的值（`AUTO_SPACE = maybe`）继续往下找，绝不因为写错一个词就让功能失效。

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace VbeComplete
{
    internal static class Settings
    {
        public const string PREFIX = "VBECOMPLETE_";

        private static readonly Dictionary<string, string> _ini =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static string _home;
        private static string _lastError = "";

        /// <summary>项目根目录（含 settings.ini / wordbank 的那个目录）。</summary>
        public static string Home
        {
            get
            {
                if (_home != null) return _home;
                string env = Environment.GetEnvironmentVariable(PREFIX + "HOME");
                if (!string.IsNullOrEmpty(env) && Directory.Exists(env)) { _home = env; return _home; }
                // 从 exe 所在目录向上找 4 层，找含 settings.example.ini 的目录
                var dir = new DirectoryInfo(AppContext.BaseDirectory);
                for (int i = 0; i < 8 && dir != null; i++, dir = dir.Parent)
                {
                    if (File.Exists(Path.Combine(dir.FullName, "settings.example.ini")) ||
                        File.Exists(Path.Combine(dir.FullName, "main.py")))
                    {
                        _home = dir.FullName;
                        return _home;
                    }
                }
                _home = AppContext.BaseDirectory;
                return _home;
            }
        }

        public static void Load()
        {
            string path = Environment.GetEnvironmentVariable(PREFIX + "SETTINGS");
            if (string.IsNullOrEmpty(path)) path = Path.Combine(Home, "settings.ini");
            _ini.Clear();
            if (!File.Exists(path)) return;
            try
            {
                string section = "";
                foreach (var raw in File.ReadAllLines(path, Encoding.UTF8))
                {
                    string line = raw.Trim();
                    if (line.Length == 0 || line[0] == ';' || line[0] == '#') continue;
                    if (line[0] == '[' && line[line.Length - 1] == ']')
                    {
                        section = line.Substring(1, line.Length - 2).Trim();
                        continue;
                    }
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    if (!string.Equals(section, "settings", StringComparison.OrdinalIgnoreCase) &&
                        section.Length != 0)
                        continue;
                    string k = line.Substring(0, eq).Trim();
                    string v = line.Substring(eq + 1).Trim();
                    // 行尾注释：`AUTO_SPACE = 1   ; 注释`
                    int sc = v.IndexOf(';');
                    if (sc >= 0 && (sc == 0 || v[sc - 1] == ' ' || v[sc - 1] == '\t'))
                        v = v.Substring(0, sc).Trim();
                    if (k.Length > 0) _ini[k] = v;
                }
            }
            catch (Exception e) { _lastError = e.Message; }
        }

        public static string LastError() => _lastError;

        /// <summary>取原始字符串；都没有就返回 null。</summary>
        public static string Raw(string key)
        {
            string v;
            if (_ini.TryGetValue(key, out v) && v.Length > 0) return v;
            v = Environment.GetEnvironmentVariable(PREFIX + key);
            if (!string.IsNullOrEmpty(v)) return v;
            // 也认不带前缀的环境变量（老用户的 set AUTO_SPACE=1）
            v = Environment.GetEnvironmentVariable(key);
            return string.IsNullOrEmpty(v) ? null : v;
        }

        private static bool? _boolOf(string s)
        {
            if (s == null) return null;
            switch (s.Trim().ToLowerInvariant())
            {
                case "1": case "true": case "yes": case "on": return true;
                case "0": case "false": case "no": case "off": return false;
                default: return null;   // 认不出 -> 继续往下找
            }
        }

        public static bool Flag(string key, bool def)
        {
            var b = _boolOf(Raw(key));
            return b ?? def;
        }

        public static int Int(string key, int def, int min = int.MinValue, int max = int.MaxValue)
        {
            string s = Raw(key);
            int v;
            if (s != null && int.TryParse(s.Trim(), NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out v))
                return Math.Max(min, Math.Min(max, v));
            return def;
        }

        public static double Double(string key, double def)
        {
            string s = Raw(key);
            double v;
            if (s != null && double.TryParse(s.Trim(), NumberStyles.Float,
                    CultureInfo.InvariantCulture, out v))
                return v;
            return def;
        }
    }
}
