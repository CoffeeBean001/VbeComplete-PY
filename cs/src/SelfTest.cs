// 自检：把"补全准不准"的全部判据跑一遍，结果写进 selftest.log。
//
// 用文件而不是 Console：程序是 WinExe（没有控制台），而自检必须在没有桌面的
// 环境下也能跑。

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace VbeComplete
{
    internal static class SelfTest
    {
        private static int _pass, _fail;
        private static readonly StringBuilder _sb = new StringBuilder();

        private static void Ok(string name, bool cond, string detail = "")
        {
            if (cond) { _pass++; }
            else
            {
                _fail++;
                _sb.AppendLine("FAIL  " + name + (detail.Length > 0 ? "  <- " + detail : ""));
            }
        }

        private static void Eq<T>(string name, T got, T want)
        {
            bool ok = EqualityComparer<T>.Default.Equals(got, want);
            Ok(name, ok, ok ? "" : "got=" + got + " want=" + want);
        }

        public static int Run()
        {
            _sb.AppendLine("=== VbeComplete C# selftest ===");
            _sb.AppendLine("kernel " + Environment.Version + "  64bit=" + Environment.Is64BitProcess);
            _sb.AppendLine("home   " + Settings.Home);
            _sb.AppendLine();

            try
            {
                TestRules();
                TestParser();
                TestEngine();
            }
            catch (Exception e)
            {
                _fail++;
                _sb.AppendLine("EXCEPTION " + e);
            }

            _sb.AppendLine();
            _sb.AppendLine($"PASS {_pass}   FAIL {_fail}");
            string path = Path.Combine(Settings.Home, "selftest.log");
            try { File.WriteAllText(path, _sb.ToString(), new UTF8Encoding(false)); } catch { }
            return _fail == 0 ? 0 : 1;
        }

        // ------------------------------------------------------------------
        private static void TestRules()
        {
            int ws;
            Eq("word-before/末尾", Rules.WordBefore("Dim abc", 8, out ws), "abc");
            Eq("word-before/起始列", ws, 5);
            Eq("word-before/词中间", Rules.WordBefore("abcdef", 4, out ws), "abc");
            Eq("word-before/无词", Rules.WordBefore("x = ", 5, out ws), "");
            Eq("word-before/空行", Rules.WordBefore("", 1, out ws), "");

            Eq("replace-word", Rules.ReplaceWord("Dim abc", 5, 8, "abcd"), "Dim abcd");
            Eq("replace-word/带后缀", Rules.ReplaceWord("MsgBox \"x\",vbExclamation:Exit Su", 31, 33, "Sub"),
                "MsgBox \"x\",vbExclamation:Exit Sub");

            Eq("dispcol/纯ascii", Rules.DispCol("    abc", 7), 8);
            Eq("dispcol/tab", Rules.DispCol("\tabc", 3), 7);
            Eq("dispcol/宽字符", Rules.DispCol("中文a", 3), 6);

            Eq("find-end/同位置", Rules.FindCompletionEnd("    abc = 1", "abc", 7, 7), 7);
            Eq("find-end/被重排后", Rules.FindCompletionEnd("MsgBox \"已\", vbEx: Exit Sub", "Sub", 24, 22), 26);

            int pos, span;
            Ok("match/前缀", Rules.Match("MsgBox", "msg", out pos, out span) && pos == 0);
            Ok("match/边界跳步", Rules.Match("selBegin", "sb", out pos, out span) && pos == 0 && span == 4);
            Ok("match/非边界不跳步", !Rules.Match("abcDef", "b", out pos, out span) || pos == 0);
            Ok("match/不命中", !Rules.Match("abc", "xyz", out pos, out span));

            Ok("resolved/精确且无更长",
                Rules.PrefixIsResolved("sub", new List<string> { "sub", "function" }));
            Ok("resolved/有更长候选不算打全",
                !Rules.PrefixIsResolved("sub", new List<string> { "sub", "subset" }));
            Ok("resolved/没精确命中",
                !Rules.PrefixIsResolved("su", new List<string> { "sub" }));

            Ok("member-position/点号后", Rules.InMemberPosition("student.", 9));
            Ok("member-position/词中间不是", !Rules.InMemberPosition("student.na", 10));
            Ok("member-position/空格后点号", Rules.InMemberPosition("a = b. ", 8));

            Ok("decl-position/Dim 名字", Rules.InDeclarationPosition("Dim cus", 8));
            Ok("decl-position/As 后是类型位置", !Rules.InDeclarationPosition("Dim x As Inte", 13));
            Ok("decl-position/Sub 名字", Rules.InDeclarationPosition("Sub Do", 7));
            Ok("decl-position/普通语句不是", !Rules.InDeclarationPosition("x = cus", 8));
            Ok("decl-position/Const 值位置", !Rules.InDeclarationPosition("Const PI = 3", 13));

            Ok("comment/注释里", Rules.InCommentOrString("' abc", 5));
            Ok("comment/字符串里", Rules.InCommentOrString("x = \"ab\"", 8));
            Ok("comment/普通代码不是", !Rules.InCommentOrString("x = 1", 5));

            char ch;
            Ok("line-edit/插入一个字符", Rules.LineEditInserted("Dim ab", "Dim abc", out ch) && ch == 'c');
            Ok("line-edit/删除不算", !Rules.LineEditInserted("Dim abc", "Dim ab", out ch));
            Ok("line-edit/替换不算", !Rules.LineEditInserted("Dim ab", "Dim xb", out ch));

            Ok("call-site/赋值右侧", Rules.IsCallSite("x = getVa", 9));
            Ok("call-site/点号后不是", !Rules.IsCallSite("a.b", 3));
            Ok("call-site/声明位置不是", !Rules.IsCallSite("Sub foo", 5));
            Ok("call-site/As 之后不是", !Rules.IsCallSite("Dim x As Long", 10));
        }

        // ------------------------------------------------------------------
        private const string Sample = @"Option Explicit
Private mCount As Long

Public Sub Main()
    Dim a As Long, b As String
    a = 1
    Call Helper(a)
    Set obj = New Collection
End Sub

Private Function GetName(ByVal idx As Long, Optional ByVal tag As String = """") As String
    GetName = ""x""
End Function

Public Property Get Size() As Long
    Size = 1
End Property

Public Sub Main()
    ' 同名过程的第二个
End Sub

Public Enum Color2
    clRed = 1
    clBlue
End Enum

Public Type Point2
    X As Long
    Y As Long
End Type
";

        private static void TestParser()
        {
            var procs = Parser.ExtractProcedures(Sample);
            Eq("procs/数量", procs.Count, 4);
            Eq("procs/第一个名字", procs[0].Name, "Main");
            Eq("procs/第一个种类", procs[0].Kind, "sub");
            Eq("procs/函数种类", procs[1].Kind, "function");
            Eq("procs/属性种类", procs[2].Kind, "property");
            Ok("procs/同名第二个带序号", procs[3].ScopeKey == "Main#2", procs[3].ScopeKey);

            var p = Parser.ProcOfLine(procs, 6);
            Ok("procs/第 6 行在 Main 里", p != null && p.Name == "Main", p?.ToString());

            var decl = Parser.ExtractDeclared(Sample);
            Ok("decl/多声明 a", decl.Contains("a"));
            Ok("decl/多声明 b", decl.Contains("b"));
            Ok("decl/参数 idx", decl.Contains("idx"));
            Ok("decl/参数 tag", decl.Contains("tag"));
            Ok("decl/枚举成员", decl.Contains("clRed"));
            Ok("decl/类型成员", decl.Contains("X"));
            Ok("decl/不能有 As", !decl.Contains("As"));
            Ok("decl/不能有 Optional", !decl.Contains("Optional"));

            var ids = Parser.ExtractIdentifiers(Sample);
            Ok("ids/收过程名", ids.Contains("Helper"));
            Ok("ids/收变量", ids.Contains("mCount"));
            Ok("ids/不收关键字", !ids.Contains("Dim"));
            Ok("ids/不收字符串内容", !ids.Contains("同名过程的第二个"));
            Ok("ids/不收注释词", !ids.Contains("同名过程"));

            var types = Parser.ExtractTypeNames(Sample);
            Ok("types/Enum 名", types.Contains("Color2"));
            Ok("types/Type 名", types.Contains("Point2"));

            Ok("option-explicit/有", Parser.OptionExplicit(Sample));
            Ok("option-explicit/没有", !Parser.OptionExplicit("Sub a()\r\nEnd Sub"));

            Eq("code-only/去注释", Parser.CodeOnly("x = 1 ' abc").Trim(), "x = 1");
            Eq("code-only/去字符串", Parser.CodeOnly("x = \"abc\"").Trim(), "x =");
            Eq("code-only/Rem", Parser.CodeOnly("Rem abc"), "");
            Eq("code-only/不误伤 Remark", Parser.CodeOnly("Remark = 1").Trim(), "Remark = 1");
        }

        // ------------------------------------------------------------------
        private static void TestEngine()
        {
            var eng = new Engine
            {
                Recent = new Recent(),
                Bank = new WordBank(),
            };
            eng.PrimePool(Sample, "Module1");

            bool wp, ws;
            eng.Followup("Main", "Other", out wp, out ws);
            Ok("followup/sub 补空格", !wp && ws, $"parens={wp} space={ws}");

            eng.Followup("GetName", "Other", out wp, out ws);
            Ok("followup/function 补括号", wp && !ws, $"parens={wp} space={ws}");

            eng.Followup("GetName", "GetName", out wp, out ws);
            Ok("followup/函数体内给自己赋值不补", !wp && !ws, $"parens={wp} space={ws}");

            eng.Followup("GetName", "getname#2", out wp, out ws);
            Ok("followup/作用域键带序号也认得出自己", !wp && !ws, $"parens={wp} space={ws}");

            eng.Followup("Size", "Other", out wp, out ws);
            Ok("followup/property 都不补", !wp && !ws, $"parens={wp} space={ws}");

            eng.Followup("MsgBox", "Other", out wp, out ws);
            Ok("followup/内建函数补括号", wp && !ws, $"parens={wp} space={ws}");

            eng.Followup("String", "Other", out wp, out ws);
            Ok("followup/双身份名字不补括号", !wp, $"parens={wp}");

            eng.Followup("Public", "Other", out wp, out ws);
            Ok("followup/关键字补空格", !wp && ws, $"parens={wp} space={ws}");

            eng.Followup("Next", "Other", out wp, out ws);
            Ok("followup/Next 不补空格", !wp && !ws, $"parens={wp} space={ws}");

            eng.Followup("Me", "Other", out wp, out ws);
            Ok("followup/Me 不补空格", !wp && !ws, $"parens={wp} space={ws}");

            Eq("classify/工程过程", eng.Classify("Main"), "func");
            Eq("classify/工程属性", eng.Classify("Size"), "prop");
            Eq("classify/工程类型", eng.Classify("Color2"), "type");
            Eq("classify/内建函数", eng.Classify("MsgBox"), "func");
            Eq("classify/内建类型", eng.Classify("Long"), "type");
            Eq("classify/关键字", eng.Classify("Public"), "kw");
            Eq("classify/变量", eng.Classify("mCount"), "var");

            // 排序：工程内的名字必须排在内建前面
            var a = new Candidate { Name = "Main", ScopeKind = 3, Pos = 0, Span = 2, Kind = "func" };
            var b = new Candidate { Name = "MsgBox", ScopeKind = 2, Pos = 0, Span = 2, Kind = "func" };
            Ok("sort/工程名优先于内建", Engine.Cmp(a, b) < 0);

            var c1 = new Candidate { Name = "abc", ScopeKind = 3, Pos = 0, Span = 1 };
            var c2 = new Candidate { Name = "abc", ScopeKind = 3, Pos = 1, Span = 1 };
            Ok("sort/命中位置靠前优先", Engine.Cmp(c1, c2) < 0);

            var r1 = new Candidate { Name = "a", ScopeKind = 3, Pos = 0, Span = 1, RecentRank = 5 };
            var r2 = new Candidate { Name = "b", ScopeKind = 3, Pos = 0, Span = 1, RecentRank = 1 };
            Ok("sort/最近使用优先", Engine.Cmp(r1, r2) < 0);
        }
    }
}
