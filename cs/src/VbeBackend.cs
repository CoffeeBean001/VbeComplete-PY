// VBE 后端：读快照、读模块全文、把补全写回去、以及那些"改一行"的操作。
//
// ★所有方法都只在 UI 线程调用，且【任何 COM 异常都被吞掉并返回失败】——
//   补全工具永远不允许把异常抛到用户脸上，更不允许因为 VBE 忙就把自己卡死。

using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace VbeComplete
{
    internal sealed class Snapshot
    {
        public int Line;
        public string Text;
        public string Module;
        public int CaretCol;
        public override string ToString() => $"L{Line} C{CaretCol} {Module}";
    }

    internal sealed class VbeBackend
    {
        // 快照缓存：行列都没动就复用上一轮那份（读数从 8.7ms 降到 0.13ms 的关键）
        private Snapshot _snapCache;
        private DateTime _snapAt = DateTime.MinValue;
        private readonly bool _snapEnabled = !Settings.Flag("NO_SNAP_CACHE", false);
        private readonly double _snapMaxAge = Settings.Double("SNAP_MAX_AGE", 1.0);

        // 模块全文缓存
        private string _modText;
        private string _modName;
        private int _modCount = -1;
        private DateTime _modAt = DateTime.MinValue;

        // 工程组件缓存
        private List<string> _comps;
        private DateTime _compsAt = DateTime.MinValue;

        public static readonly bool AutoParen = !Settings.Flag("NO_AUTOPAREN", false);
        public static readonly bool AutoSpace = !Settings.Flag("NO_AUTOSPACE", false);

        public Action<char> SendChar;     // 由 main 注入（补空格用）
        public Action<int, bool, bool> SendVk;

        public void InvalidateSnapshot()
        {
            _snapCache = null;
            _modAt = DateTime.MinValue;
        }

        private void Fail(Exception e)
        {
            Log.Write("COM 失败：{0}", e.Message);
            ComCtx.Reset();
            _snapCache = null;
            _modName = null;
            _modCount = -1;
            _modAt = DateTime.MinValue;
        }

        // ------------------------------------------------------------------
        // 读
        // ------------------------------------------------------------------

        /// <summary>
        /// 轻量快照：行号 / 行文本 / 模块名 / 光标列。多行选择或不在代码窗时返回 null。
        /// 两级读取：先只 GetSelection（0.13ms），行列没动就直接返回上一轮。
        /// </summary>
        public Snapshot Snapshot(bool verify = true)
        {
            try
            {
                dynamic vbe = ComCtx.Vbe(verify);
                if (ReferenceEquals(vbe, null)) return null;
                dynamic cp = vbe.ActiveCodePane;
                if (ReferenceEquals(cp, null)) { _snapCache = null; return null; }

                int sl, sc, el, ec;
                if (!ComCall.CallFourOut(cp, "GetSelection", out sl, out sc, out el, out ec))
                {
                    _snapCache = null;
                    return null;
                }
                if (sl != el) { _snapCache = null; return null; }

                if (_snapEnabled && _snapCache != null &&
                    _snapCache.Line == sl && _snapCache.CaretCol == ec &&
                    (DateTime.UtcNow - _snapAt).TotalSeconds <= _snapMaxAge)
                    return _snapCache;

                dynamic cm = cp.CodeModule;
                string mod = "";
                try { mod = Convert.ToString(cm.Name) ?? ""; } catch { }
                string text = Convert.ToString(cm.Lines(sl, 1)) ?? "";

                var s = new Snapshot { Line = sl, Text = text, Module = mod, CaretCol = ec };
                if (_snapEnabled) { _snapCache = s; _snapAt = DateTime.UtcNow; }

                // 当前行文本变了 ⇒ 模块全文缓存作废（不只是光标动）
                if (mod != _modName) { _modName = mod; _modCount = -1; _modAt = DateTime.MinValue; }
                return s;
            }
            catch (Exception e) { Fail(e); return null; }
        }

        /// <summary>当前模块全文。带缓存：模块名 + 行数都没变就不重读。</summary>
        public string ModuleText()
        {
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return null;
                dynamic cp = vbe.ActiveCodePane;
                if (ReferenceEquals(cp, null)) return null;
                dynamic cm = cp.CodeModule;
                string name = Convert.ToString(cm.Name) ?? "";
                int count = Convert.ToInt32(cm.CountOfLines);
                if (_modText != null && _modName == name && _modCount == count &&
                    (DateTime.UtcNow - _modAt).TotalSeconds < 1.5)
                    return _modText;

                string all = count > 0 ? Convert.ToString(cm.Lines(1, count)) : "";
                _modText = all ?? "";
                _modName = name;
                _modCount = count;
                _modAt = DateTime.UtcNow;
                return _modText;
            }
            catch (Exception e) { Fail(e); return null; }
        }

        /// <summary>我在"文本被我们改过"之后必须调它，否则下一轮读到的是旧代码。</summary>
        public void OnWeWrote()
        {
            _modCount = -1;
            _modAt = DateTime.MinValue;
            InvalidateSnapshot();
        }

        /// <summary>"纯文本里失效"的标记：确认模块缓存已经跟不上时强制重读一次。</summary>
        public void ForceReloadModule()
        {
            _modCount = -1;
            _modAt = DateTime.MinValue;
        }

        public int CurrentLine()
        {
            var s = Snapshot();
            return s == null ? -1 : s.Line;
        }

        /// <summary>本工程所有组件的名字（模块 / 类模块 / 窗体）。缓存 30 秒。</summary>
        public List<string> ComponentNames()
        {
            if (_comps != null && (DateTime.UtcNow - _compsAt).TotalSeconds < 30) return _comps;
            var list = new List<string>();
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return list;
                dynamic proj = vbe.ActiveVBProject;
                if (ReferenceEquals(proj, null)) return list;
                dynamic comps = proj.VBComponents;
                int n = Convert.ToInt32(comps.Count);
                for (int i = 1; i <= n; i++)
                {
                    try
                    {
                        dynamic c = comps.Item(i);
                        string nm = Convert.ToString(c.Name);
                        if (!string.IsNullOrEmpty(nm)) list.Add(nm);
                    }
                    catch { }
                }
                _comps = list;
                _compsAt = DateTime.UtcNow;
            }
            catch (Exception e) { Fail(e); }
            return list;
        }

        // ------------------------------------------------------------------
        // 写：补全
        // ------------------------------------------------------------------

        private static readonly Regex ReCallStmt =
            new Regex(@"^[ \t]*Call[ \t]+", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static bool InCallStatement(string line) =>
            line != null && ReCallStmt.IsMatch(line);

        /// <summary>
        /// 把 [wordStartCol, endCol) 换成 completion（都是 1-based 字符列）。
        /// 返回 true 表示真的写进去了。
        /// </summary>
        public bool ApplyCompletion(int lineNo, int wordStartCol, int endCol,
            string completion, bool wantParens, bool wantSpace, string enclosingProc)
        {
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return false;
                dynamic cp = vbe.ActiveCodePane;
                if (ReferenceEquals(cp, null)) return false;
                dynamic cm = cp.CodeModule;

                string lineText = Convert.ToString(cm.Lines(lineNo, 1)) ?? "";
                string newLine = Rules.ReplaceWord(lineText, wordStartCol, endCol, completion);

                if (wantParens && AutoParen && Rules.IsCallSite(lineText, wordStartCol)
                    && !InCallStatement(lineText))
                {
                    int k = wordStartCol - 1 + completion.Length;
                    if (k > newLine.Length) k = newLine.Length;
                    // 跳过空白【只用于判断】右边是不是已经紧跟左括号
                    int j = k;
                    while (j < newLine.Length && (newLine[j] == ' ' || newLine[j] == '\t')) j++;
                    if (!(j < newLine.Length && newLine[j] == '('))
                        // 插在词尾 ⇒ 下面的 end0 正好指向这个新 `(`，
                        // 光标按它算就自然落进括号中间，不必另写分支。
                        newLine = newLine.Substring(0, k) + "()" + newLine.Substring(k);
                }

                int start0 = Math.Max(0, wordStartCol - 1);
                int end0 = start0 + completion.Length;

                cm.ReplaceLine(lineNo, newLine);
                OnWeWrote();

                // ★v126e：ReplaceLine 会重排整行（`,` 后、`:` 后、`=` 两边补空格），
                //   所以落光标的位置必须按【读回来的真行】算，不能拿写入前的下标。
                string actual = Convert.ToString(cm.Lines(lineNo, 1)) ?? "";
                if (actual != newLine)
                    end0 = Rules.FindCompletionEnd(actual, completion, end0 - completion.Length, end0);
                if (end0 > actual.Length) end0 = actual.Length;

                int col = Rules.DispCol(actual, end0, 4);
                // ★v131b：光标就停在 col 这一列 —— 补括号时 end0 指向那个 `(`，
                //   而 DispCol 给的已经是"这个字符之前"的 1-based 列，所以
                //   **不能再减 1**。减了就会跑到词尾字符之前（`Spli|t`），
                //   用户报的"落在 Split 中间"就是这么来的。
                int caretCol = col;
                if (caretCol < 0) caretCol = 0;

                cp.SetSelection(lineNo, 1, lineNo, 1);
                cp.SetSelection(lineNo, caretCol, lineNo, caretCol);

                bool wantSp = wantSpace && AutoSpace && !InCallStatement(lineText);
                if (wantSp && SendChar != null)
                {
                    // 空格不能靠改行文本补：VBE 会把行尾空格立刻 trim 掉
                    if (end0 >= actual.Length || actual[end0] != ' ')
                        SendChar(' ');
                }
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        // ------------------------------------------------------------------
        // 写：行操作
        // ------------------------------------------------------------------

        /// <summary>在第 lineNo 行后面插入一行，并把光标放过去（缩进由调用方给）。</summary>
        public bool InsertLineBelow(int lineNo, string text, int caretCol)
        {
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return false;
                dynamic cp = vbe.ActiveCodePane;
                dynamic cm = cp.CodeModule;
                cm.InsertLines(lineNo + 1, text);
                OnWeWrote();
                cp.SetSelection(lineNo + 1, caretCol, lineNo + 1, caretCol);
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        public bool ReplaceLineRaw(int lineNo, string text)
        {
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return false;
                dynamic cm = vbe.ActiveCodePane.CodeModule;
                cm.ReplaceLine(lineNo, text);
                OnWeWrote();
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        public bool DeleteLines(int from, int count)
        {
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return false;
                dynamic cm = vbe.ActiveCodePane.CodeModule;
                cm.DeleteLines(from, count);
                OnWeWrote();
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        public bool InsertLines(int at, string text)
        {
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return false;
                dynamic cm = vbe.ActiveCodePane.CodeModule;
                cm.InsertLines(at, text);
                OnWeWrote();
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        public int LineCount()
        {
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return -1;
                return Convert.ToInt32(vbe.ActiveCodePane.CodeModule.CountOfLines);
            }
            catch (Exception e) { Fail(e); return -1; }
        }

        public string LineAt(int lineNo)
        {
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return null;
                return Convert.ToString(vbe.ActiveCodePane.CodeModule.Lines(lineNo, 1)) ?? "";
            }
            catch (Exception e) { Fail(e); return null; }
        }

        /// <summary>把光标放到某行的行尾（列交给 VBE 自己钳）。</summary>
        public bool CaretToLineEnd(int lineNo)
        {
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return false;
                dynamic cp = vbe.ActiveCodePane;
                dynamic cm = cp.CodeModule;
                int count = Convert.ToInt32(cm.CountOfLines);
                if (lineNo < 1 || lineNo > count) return false;
                string text = Convert.ToString(cm.Lines(lineNo, 1)) ?? "";
                int col = Rules.DispCol(text.TrimEnd('\r', '\n'), text.TrimEnd('\r', '\n').Length, 4);
                cp.SetSelection(lineNo, 1, lineNo, 1);
                cp.SetSelection(lineNo, col, lineNo, col);
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        public bool SetCaret(int lineNo, int charCol)
        {
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return false;
                dynamic cp = vbe.ActiveCodePane;
                string text = Convert.ToString(cp.CodeModule.Lines(lineNo, 1)) ?? "";
                int col = Rules.DispCol(text, charCol, 4);
                cp.SetSelection(lineNo, 1, lineNo, 1);
                cp.SetSelection(lineNo, col, lineNo, col);
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        /// <summary>当前选区（四个 Long）。四个出参都是 ByRef，只能走 IDispatch 直调。</summary>
        public bool TryGetSelection(out int sl, out int sc, out int el, out int ec)
        {
            sl = sc = el = ec = 0;
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return false;
                object cp = vbe.ActiveCodePane;
                if (ReferenceEquals(cp, null)) return false;
                return ComCall.CallFourOut(cp, "GetSelection", out sl, out sc, out el, out ec);
            }
            catch (Exception e) { Fail(e); return false; }
        }

        public string SelectedText()
        {
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return "";
                dynamic cp = vbe.ActiveCodePane;
                int sl, sc, el, ec;
                if (!ComCall.CallFourOut(cp, "GetSelection", out sl, out sc, out el, out ec)) return "";
                dynamic cm = cp.CodeModule;
                if (el < sl) return "";
                if (el == sl && ec <= sc) return "";
                return Convert.ToString(cm.Lines(sl, el - sl + 1)) ?? "";
            }
            catch (Exception e) { Fail(e); return ""; }
        }

        public bool FocusVbe()
        {
            try
            {
                dynamic vbe = ComCtx.Vbe();
                if (ReferenceEquals(vbe, null)) return false;
                vbe.MainWindow.Visible = true;
                vbe.ActiveCodePane.Show();
                return true;
            }
            catch (Exception e) { Fail(e); return false; }
        }

        public string EnclosingProcName()
        {
            var snap = Snapshot();
            if (snap == null) return null;
            string text = ModuleText();
            if (text == null) return null;
            var procs = Parser.ExtractProcedures(text);
            var p = Parser.ProcOfLine(procs, snap.Line);
            return p?.Name;
        }
    }
}
