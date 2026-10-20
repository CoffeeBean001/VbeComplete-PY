// VBA 源码解析：从模块全文里收"工程自己定义的名字"。
//
// ★这一层【完全不碰 COM】，只吃一段字符串 —— 于是它能被单独测试，
//   也是"候选池里到底有什么"的唯一事实来源。
//
// 收什么（用户一条条要求出来的口径）：
//     过程名（Sub/Function/Property）、变量与常量、类型与枚举名、
//     窗体/控件名、参数名、以及其他一切出现过标识符形状的词。
// 不收：注释里和字符串里的词（那是在说人话，不是代码）。

using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace VbeComplete
{
    internal sealed class ProcInfo
    {
        public string Name;
        public string Kind;        // sub / function / property
        public int StartLine;      // 1-based，等于 VBE 的行号
        public int EndLine;
        public string ScopeKey;    // 同名过程时带 #序号

        public override string ToString() => $"{Kind} {Name} [{StartLine}-{EndLine}]";
    }

    internal static class Parser
    {
        // 标识符：字母/下划线开头，后面可以跟字母数字下划线，可选类型后缀
        private const string IDENT = @"[A-Za-z_][A-Za-z0-9_]*[\%\&\#\@\$]?";

        private static readonly Regex ReIdent = new Regex(IDENT, RegexOptions.Compiled);

        private static readonly Regex ReSub = new Regex(
            @"^[ \t]*(?:(?:Public|Private|Friend|Global)[ \t]+)?(?:Static[ \t]+)?Sub[ \t]+(" + IDENT + @")",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReFunc = new Regex(
            @"^[ \t]*(?:(?:Public|Private|Friend|Global)[ \t]+)?(?:Static[ \t]+)?Function[ \t]+(" + IDENT + @")",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReProp = new Regex(
            @"^[ \t]*(?:(?:Public|Private|Friend|Global)[ \t]+)?(?:Static[ \t]+)?Property[ \t]+(?:Get|Let|Set)[ \t]+(" + IDENT + @")",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReEndProc = new Regex(
            @"^[ \t]*End[ \t]+(?:Sub|Function|Property)\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReDeclHead = new Regex(
            @"^[ \t]*(?:(?:Public|Private|Friend|Global|Static)[ \t]+)*(?:Dim|Const|ReDim|WithEvents)[ \t]+",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReEnumHead = new Regex(
            @"^[ \t]*(?:(?:Public|Private)[ \t]+)?Enum[ \t]+(" + IDENT + ")",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReTypeHead = new Regex(
            @"^[ \t]*(?:(?:Public|Private)[ \t]+)?Type[ \t]+(" + IDENT + ")",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReDeclare = new Regex(
            @"^[ \t]*(?:(?:Public|Private)[ \t]+)?Declare[ \t]+(?:PtrSafe[ \t]+)?(?:Sub|Function)[ \t]+(" + IDENT + ")",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReOptionExplicit = new Regex(
            @"^[ \t]*Option[ \t]+Explicit\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>按物理行切分（VBE 的行号就是物理行号，不能把续行合并掉）。</summary>
        public static string[] SplitLines(string text)
        {
            if (text == null) return Array.Empty<string>();
            string t = text.Replace("\r\n", "\n").Replace('\r', '\n');
            // 末尾的换行会切出一个空元素，VBE 的 CountOfLines 也算这一行，保留
            return t.Split('\n');
        }

        /// <summary>去掉注释与字符串内容，只留"代码形状"的部分。</summary>
        public static string CodeOnly(string line)
        {
            if (string.IsNullOrEmpty(line)) return "";
            var sb = new StringBuilder(line.Length);
            bool inStr = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (inStr)
                {
                    if (c == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"') { i++; continue; }
                        inStr = false;
                    }
                    sb.Append(' ');
                    continue;
                }
                if (c == '"') { inStr = true; sb.Append(' '); continue; }
                if (c == '\'') break;                      // 行内注释
                sb.Append(c);
            }
            string code = sb.ToString();
            // 行首 Rem ... 也是注释（Rem 必须顶格或前面只有空白）
            string trim = code.TrimStart();
            if (trim.Length >= 3 &&
                (trim[0] == 'R' || trim[0] == 'r') &&
                (trim[1] == 'E' || trim[1] == 'e') &&
                (trim[2] == 'M' || trim[2] == 'm') &&
                (trim.Length == 3 || !char.IsLetterOrDigit(trim[3])))
                return "";
            return code;
        }

        /// <summary>模块里所有过程（含同名去重用的 ScopeKey）。</summary>
        public static List<ProcInfo> ExtractProcedures(string text)
        {
            var list = new List<ProcInfo>();
            var lines = SplitLines(text);
            var open = new List<ProcInfo>();
            var used = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < lines.Length; i++)
            {
                string code = CodeOnly(lines[i]);
                if (code.Length == 0) continue;

                Match m = ReSub.Match(code);
                string kind = "sub";
                if (!m.Success) { m = ReFunc.Match(code); kind = "function"; }
                if (!m.Success) { m = ReProp.Match(code); kind = "property"; }

                if (m.Success)
                {
                    string name = m.Groups[1].Value.TrimEnd('%', '&', '!', '#', '@', '$');
                    int n;
                    used.TryGetValue(name, out n);
                    used[name] = ++n;
                    var p = new ProcInfo
                    {
                        Name = name,
                        Kind = kind,
                        StartLine = i + 1,
                        EndLine = lines.Length,
                        ScopeKey = n > 1 ? name + "#" + n : name,
                    };
                    open.Add(p);
                    list.Add(p);
                    continue;
                }

                if (ReEndProc.IsMatch(code) && open.Count > 0)
                {
                    open[open.Count - 1].EndLine = i + 1;
                    open.RemoveAt(open.Count - 1);
                }
            }
            return list;
        }

        /// <summary>第 line 行（1-based）落在哪个过程里；不在任何过程里返回 null。</summary>
        public static ProcInfo ProcOfLine(List<ProcInfo> procs, int line)
        {
            if (procs == null) return null;
            foreach (var p in procs)
                if (line >= p.StartLine && line <= p.EndLine) return p;
            return null;
        }

        /// <summary>收"声明出来的名字"：Dim/Const/Static/参数/Enum 成员/Type 成员。</summary>
        public static HashSet<string> ExtractDeclared(string text)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var lines = SplitLines(text);
            bool inEnumOrType = false;

            foreach (var raw in lines)
            {
                string code = CodeOnly(raw);
                if (code.Length == 0) continue;

                if (ReEnumHead.IsMatch(code) || ReTypeHead.IsMatch(code))
                {
                    var h = ReEnumHead.Match(code);
                    if (!h.Success) h = ReTypeHead.Match(code);
                    if (h.Success) set.Add(h.Groups[1].Value);
                    inEnumOrType = true;
                    continue;
                }
                if (Regex.IsMatch(code, @"^[ \t]*End[ \t]+(?:Enum|Type)\b", RegexOptions.IgnoreCase))
                {
                    inEnumOrType = false;
                    continue;
                }
                if (inEnumOrType)
                {
                    // Enum 成员：`Name = 1`；Type 成员：`Name As Long`
                    var mm = Regex.Match(code, @"^[ \t]*(" + IDENT + @")");
                    if (mm.Success) set.Add(Clean(mm.Groups[1].Value));
                    continue;
                }

                if (ReDeclHead.IsMatch(code))
                {
                    foreach (var n in NamesInDecl(code)) set.Add(n);
                    continue;
                }

                // 过程头里的参数：Sub Foo(a As Long, Optional ByVal b = 1)
                var pm = ReSub.Match(code);
                if (!pm.Success) pm = ReFunc.Match(code);
                if (!pm.Success) pm = ReProp.Match(code);
                if (pm.Success)
                {
                    int lp = code.IndexOf('(');
                    int rp = code.LastIndexOf(')');
                    if (lp >= 0 && rp > lp)
                        foreach (var n in NamesInParamList(code.Substring(lp + 1, rp - lp - 1)))
                            set.Add(n);
                }
            }
            return set;
        }

        private static IEnumerable<string> NamesInDecl(string code)
        {
            var m = ReDeclHead.Match(code);
            if (!m.Success) yield break;
            string rest = code.Substring(m.Index + m.Length);
            // 去掉 `As Type` 之后的部分不影响：我们按逗号切成声明符，每个取首个标识符
            foreach (var part in SplitTopLevel(rest))
            {
                string p = part.Trim();
                if (p.Length == 0) continue;
                // `x(1 To 10) As Long` -> x
                var mm = Regex.Match(p, "^(" + IDENT + @")");
                if (!mm.Success) continue;
                string name = Clean(mm.Groups[1].Value);
                if (name.Length == 0) continue;
                if (VbaBuiltins.IsKeyword(name)) continue;
                yield return name;
            }
        }

        private static IEnumerable<string> NamesInParamList(string inner)
        {
            foreach (var part in SplitTopLevel(inner))
            {
                string p = part.Trim();
                if (p.Length == 0) continue;
                // 去掉 Optional / ByVal / ByRef / ParamArray / 修饰
                p = Regex.Replace(p, @"^\s*(?:Optional|ByVal|ByRef|ParamArray)\b", "",
                        RegexOptions.IgnoreCase).Trim();
                p = Regex.Replace(p, @"^\s*(?:Optional|ByVal|ByRef|ParamArray)\b", "",
                        RegexOptions.IgnoreCase).Trim();
                p = p.TrimStart('(', ')');
                var mm = Regex.Match(p, "^(" + IDENT + @")");
                if (!mm.Success) continue;
                string name = Clean(mm.Groups[1].Value);
                if (name.Length == 0 || VbaBuiltins.IsKeyword(name)) continue;
                yield return name;
            }
        }

        /// <summary>按顶层逗号切分（不切括号里面的逗号）。</summary>
        public static List<string> SplitTopLevel(string s)
        {
            var outp = new List<string>();
            if (s == null) return outp;
            int depth = 0;
            var cur = new StringBuilder();
            foreach (char c in s)
            {
                if (c == '(') depth++;
                else if (c == ')') depth = Math.Max(0, depth - 1);
                if (c == ',' && depth == 0) { outp.Add(cur.ToString()); cur.Clear(); continue; }
                cur.Append(c);
            }
            outp.Add(cur.ToString());
            return outp;
        }

        private static string Clean(string name)
        {
            if (string.IsNullOrEmpty(name)) return "";
            return name.TrimEnd('%', '&', '!', '#', '@', '$').Trim();
        }

        /// <summary>
        /// 模块里出现过的所有标识符形状的词（不含注释/字符串里的）。
        /// 这是候选池的主体 —— 变量、过程、控件名、用户自定义类型都在里面。
        /// </summary>
        public static HashSet<string> ExtractIdentifiers(string text, ICollection<string> alsoAdd = null)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (text == null) return set;
            foreach (var raw in SplitLines(text))
            {
                string code = CodeOnly(raw);
                if (code.Length == 0) continue;
                foreach (Match m in ReIdent.Matches(code))
                {
                    string w = Clean(m.Value);
                    if (w.Length == 0) continue;
                    if (VbaBuiltins.IsKeyword(w)) continue;     // 关键字单独一组
                    set.Add(w);
                }
            }
            if (alsoAdd != null)
                foreach (var n in alsoAdd) if (!string.IsNullOrEmpty(n)) set.Add(n);
            return set;
        }

        /// <summary>模块里声明的类型名（Enum X / Type X）。用于图标档位。</summary>
        public static HashSet<string> ExtractTypeNames(string text)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (text == null) return set;
            foreach (var raw in SplitLines(text))
            {
                string code = CodeOnly(raw);
                if (code.Length == 0) continue;
                var m = ReEnumHead.Match(code);
                if (!m.Success) m = ReTypeHead.Match(code);
                if (m.Success) set.Add(m.Groups[1].Value);
            }
            return set;
        }

        public static bool OptionExplicit(string text)
        {
            if (text == null) return false;
            foreach (var raw in SplitLines(text))
                if (ReOptionExplicit.IsMatch(raw)) return true;
            return false;
        }

        /// <summary>过程头的种类（sub/function/property），键为小写名。</summary>
        public static Dictionary<string, string> ProcKinds(List<ProcInfo> procs)
        {
            var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (procs == null) return d;
            foreach (var p in procs) d[p.Name] = p.Kind;
            return d;
        }
    }
}
