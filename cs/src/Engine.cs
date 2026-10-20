// 候选引擎：攒池子、匹配、排序、决定弹不弹、以及"选完之后该补什么"。
//
// ★这个类不 import WinForms 的任何东西 —— 窗口通过回调接进来。
//   于是引擎的全部行为（弹不弹、排第几、补括号还是补空格）都能在没有界面的
//   情况下测。

using System;
using System.Collections.Generic;
using System.Linq;

namespace VbeComplete
{
    internal sealed class Candidate
    {
        public string Name;
        public string Kind;        // func / prop / var / type / kw —— 图标档位
        public int Pos;            // 命中位置（越小越好）
        public int Span;           // 命中跨度（越小越好）
        public int ScopeKind;      // 3 工程内 / 2 内建函数 / 1 内建类型 / 0 关键字
        public int RecentRank;
        public bool IsBuiltin;
    }

    internal sealed class Engine
    {
        public VbeBackend Backend;
        public Recent Recent;
        public WordBank Bank;

        public readonly int ViewRows = Settings.Int("VIEW_ROWS", 8, 3, 30);
        public readonly bool HideOnComplete = Settings.Flag("HIDE_ON_COMPLETE", true);
        public readonly bool MemberPopup = Settings.Flag("MEMBER_POPUP", false);
        public readonly bool DeclPopup = Settings.Flag("DECL_POPUP", false);

        // UI 回调
        public Action<IReadOnlyList<Candidate>, int, int> OnShow;   // items, selected, top
        public Action OnHide;
        public Action OnMove;                                        // 选中项变了，重画

        private List<Candidate> _items = new List<Candidate>();
        private int _sel;
        private int _top;
        public bool Visible { get; private set; }

        // 池缓存
        private List<string> _pool;
        private Dictionary<string, string> _procKinds =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private HashSet<string> _typeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private string _poolKey;
        private int _lastLine = -1;
        private string _lastLineText;

        public IReadOnlyList<Candidate> Items => _items;
        public int SelectedIndex => _sel;
        public int TopIndex => _top;

        // ------------------------------------------------------------------
        // 池子
        // ------------------------------------------------------------------

        private void EnsurePool(string moduleText, string moduleName)
        {
            string key = moduleName + "|" + (moduleText == null ? -1 : moduleText.Length);
            if (_pool != null && _poolKey == key) return;

            var names = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (moduleText != null)
            {
                var procs = Parser.ExtractProcedures(moduleText);
                _procKinds = Parser.ProcKinds(procs);
                _typeNames = Parser.ExtractTypeNames(moduleText);

                foreach (var p in procs) if (seen.Add(p.Name)) names.Add(p.Name);
                foreach (var n in Parser.ExtractDeclared(moduleText)) if (seen.Add(n)) names.Add(n);
                foreach (var n in Parser.ExtractIdentifiers(moduleText)) if (seen.Add(n)) names.Add(n);
            }

            // 组件名（模块 / 类模块 / 窗体）
            if (Backend != null)
                foreach (var n in Backend.ComponentNames())
                    if (seen.Add(n)) names.Add(n);

            // 词库里的额外词
            if (Bank != null)
                foreach (var n in Bank.ExtraNames())
                    if (seen.Add(n)) names.Add(n);

            // 内建（关键字 / 函数 / 类型）—— 排在最后，只在匹配时按档位压后
            foreach (var n in VbaBuiltins.Pool())
                if (seen.Add(n)) names.Add(n);

            _pool = names;
            _poolKey = key;
        }

        /// <summary>把池子按给定源码算出来（自检用；运行期由 Trigger 调）。</summary>
        public void PrimePool(string moduleText, string moduleName)
        {
            EnsurePool(moduleText, moduleName ?? "");
        }

        public void InvalidatePool()
        {
            _poolKey = null;
            _pool = null;
        }

        // ------------------------------------------------------------------
        // 触发
        // ------------------------------------------------------------------

        /// <summary>
        /// 判断现在该不该弹。manual=true 是 Ctrl+Space（豁免"打全即收"等收束规则）。
        /// 返回 true 表示窗口已显示。
        /// </summary>
        public bool Trigger(bool manual, Snapshot snap = null)
        {
            snap = snap ?? Backend?.Snapshot();
            if (snap == null) { Hide(); return false; }

            string line = snap.Text ?? "";
            // ★v131b：GetSelection 给的是**显示列**，后面所有取词/判据都按字符列算
            //   —— 含中文/全角的行不换算就会取到空词，连弹窗都不出现。
            int caret = Rules.CharColFromDisp(line, snap.CaretCol);

            if (!manual)
            {
                if (Rules.InCommentOrString(line, caret)) { Hide(); return false; }
                if (!MemberPopup && Rules.InMemberPosition(line, caret)) { Hide(); return false; }
                if (!DeclPopup && Rules.InDeclarationPosition(line, caret)) { Hide(); return false; }
            }

            int wordStart;
            string word = Rules.WordBefore(line, caret, out wordStart);
            if (!manual && word.Length == 0) { Hide(); return false; }
            if (word.Length == 0)
            {
                // Ctrl+Space 在空位置：列出全部（相当于手动要一份清单）
                wordStart = caret;
            }

            string text = Backend?.ModuleText();
            EnsurePool(text, snap.Module);

            var hits = new List<Candidate>(64);
            foreach (var name in _pool)
            {
                if (name.Length == 0) continue;
                if (!Rules.Match(name, word, out int pos, out int span)) continue;

                // 单个字母的模糊跳步匹配噪音太大：只认前缀
                if (word.Length == 1 && pos != 0) continue;

                var c = new Candidate
                {
                    Name = name,
                    Pos = pos,
                    Span = span,
                    ScopeKind = ScopeOf(name),
                    IsBuiltin = IsBuiltinName(name),
                    RecentRank = Recent != null ? Recent.Rank(name) : 0,
                };
                c.Kind = Classify(c.Name);
                hits.Add(c);
            }

            if (hits.Count == 0) { Hide(); return false; }

            // 完全相同的名字只留一条（工程名与内建表可能重名）
            hits.Sort(Cmp);
            var dedup = new List<Candidate>(hits.Count);
            var seenName = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var h in hits) if (seenName.Add(h.Name)) dedup.Add(h);
            hits = dedup;

            if (!manual && HideOnComplete &&
                Rules.PrefixIsResolved(word, hits.Select(x => x.Name).ToList()))
            {
                // 打全了：候选一并不留 —— 否则按 Tab 会把已经打完的名字再补一遍
                Hide();
                return false;
            }

            _items = hits;
            _lastLine = snap.Line;
            _lastLineText = line;
            _sel = 0;
            _top = 0;
            Visible = true;
            OnShow?.Invoke(_items, _sel, _top);
            return true;
        }

        public static int Cmp(Candidate a, Candidate b)
        {
            int c = b.ScopeKind.CompareTo(a.ScopeKind); if (c != 0) return c;
            c = a.Pos.CompareTo(b.Pos); if (c != 0) return c;
            c = a.Span.CompareTo(b.Span); if (c != 0) return c;
            c = b.RecentRank.CompareTo(a.RecentRank); if (c != 0) return c;
            c = a.IsBuiltin.CompareTo(b.IsBuiltin); if (c != 0) return c;
            c = b.Name.Length.CompareTo(a.Name.Length); if (c != 0) return c;
            return string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);
        }

        private int ScopeOf(string name)
        {
            if (VbaBuiltins.IsKeyword(name)) return 0;
            if (VbaBuiltins.IsBuiltinFunction(name)) return 2;
            if (VbaBuiltins.IsBuiltinType(name)) return 1;
            return 3;
        }

        private bool IsBuiltinName(string name) =>
            VbaBuiltins.IsBuiltinFunction(name) || VbaBuiltins.IsBuiltinType(name) ||
            VbaBuiltins.IsKeyword(name);

        /// <summary>
        /// 图标档位。判据只有这一处：工程过程 > 工程类型 > 内建函数 > 内建类型 >
        /// 关键字 > 其余。拿不到档位就返回 null（不画图标），绝不默认画成"变量"。
        /// </summary>
        public string Classify(string name)
        {
            if (string.IsNullOrEmpty(name)) return null;
            string kind;
            if (_procKinds.TryGetValue(name, out kind))
                return kind == "sub" ? "func" : (kind == "property" ? "prop" : "func");
            if (_typeNames.Contains(name)) return "type";

            if (Bank != null)
            {
                string bk = Bank.KindOf(name);
                if (bk == "sub" || bk == "function") return "func";
                if (bk == "property") return "prop";
                if (bk == "type") return "type";
            }

            if (VbaBuiltins.IsBuiltinFunction(name)) return "func";
            if (VbaBuiltins.IsBuiltinType(name)) return "type";
            if (VbaBuiltins.IsKeyword(name)) return "kw";
            return "var";
        }

        /// <summary>
        /// 按 Tab 确认之后，后面该跟什么：(补括号, 补空格)。
        /// 用户口径：Sub 补空格 / Function 补括号 / Property 都不补；
        /// 例外：在 Function 体内给自己赋值返回值时不补括号。
        /// </summary>
        public void Followup(string name, string enclosingProc, out bool wantParens, out bool wantSpace)
        {
            wantParens = false;
            wantSpace = false;
            if (string.IsNullOrEmpty(name)) return;
            string low = name.ToLowerInvariant();

            string kind;
            if (_procKinds.TryGetValue(name, out kind))
            {
                if (kind == "sub") { wantSpace = true; return; }
                if (kind == "function")
                {
                    if (IsOwnProcName(low, enclosingProc)) return;   // 给返回值赋值，不补
                    wantParens = true;
                    return;
                }
                return;                                              // property：都不补
            }

            if (VbaBuiltins.IsCallableBuiltin(name)) { wantParens = true; return; }

            bool? ov = Bank?.WantsSpace(name);
            wantSpace = ov ?? VbaBuiltins.KeywordWantsSpace(name);
        }

        public static bool IsOwnProcName(string name, string enclosingProc)
        {
            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(enclosingProc)) return false;
            string b = enclosingProc.ToLowerInvariant();
            int hash = b.IndexOf('#');
            if (hash >= 0) b = b.Substring(0, hash);
            return name.ToLowerInvariant() == b;
        }

        // ------------------------------------------------------------------
        // 选择
        // ------------------------------------------------------------------

        public void Move(int delta)
        {
            if (!Visible || _items.Count == 0) return;
            _sel = (_sel + delta) % _items.Count;
            if (_sel < 0) _sel += _items.Count;
            ScrollToSelected();
            OnMove?.Invoke();
        }

        public void SetSelected(int index)
        {
            if (_items.Count == 0) return;
            _sel = Math.Max(0, Math.Min(_items.Count - 1, index));
            ScrollToSelected();
            OnMove?.Invoke();
        }

        public int PageSize => Math.Max(1, ViewRows - 1);

        public void ScrollToSelected()
        {
            if (_sel < _top) _top = _sel;
            else if (_sel >= _top + ViewRows) _top = _sel - ViewRows + 1;
            int maxTop = Math.Max(0, _items.Count - ViewRows);
            if (_top > maxTop) _top = maxTop;
            if (_top < 0) _top = 0;
        }

        public void ScrollBy(int lines)
        {
            if (!Visible) return;
            int maxTop = Math.Max(0, _items.Count - ViewRows);
            _top = Math.Max(0, Math.Min(maxTop, _top + lines));
            OnMove?.Invoke();
        }

        public void Hide()
        {
            if (!Visible) return;
            Visible = false;
            _items = new List<Candidate>();
            _sel = 0; _top = 0;
            OnHide?.Invoke();
        }

        /// <summary>接受选中的候选：写回 VBE，然后收起窗口。</summary>
        public bool Accept()
        {
            if (!Visible || _items.Count == 0) return false;
            var cand = _items[_sel];
            return AcceptName(cand.Name);
        }

        public bool AcceptName(string name)
        {
            if (string.IsNullOrEmpty(name) || Backend == null) { Hide(); return false; }

            // ★写完之前重新取一次快照：候选项算出来之后用户可能又敲了字，
            //   拿旧的 word_start_col 去替换会把行改坏。
            var snap = Backend.Snapshot();
            if (snap == null) { Hide(); return false; }

            int wordStart;
            string acceptLine = snap.Text ?? "";
            int acceptCaret = Rules.CharColFromDisp(acceptLine, snap.CaretCol);
            Rules.WordBefore(acceptLine, acceptCaret, out wordStart);
            int endCol = acceptCaret;
            if (wordStart > endCol) wordStart = endCol;

            string enclosing = null;
            string text = Backend.ModuleText();
            if (text != null)
            {
                var procs = Parser.ExtractProcedures(text);
                enclosing = Parser.ProcOfLine(procs, snap.Line)?.Name;
            }

            Followup(name, enclosing, out bool wantParens, out bool wantSpace);
            bool ok = Backend.ApplyCompletion(snap.Line, wordStart, endCol, name,
                wantParens, wantSpace, enclosing);
            Hide();
            if (ok) { Recent?.Note(name); Recent?.Flush(); InvalidatePool(); }
            return ok;
        }

        /// <summary>Tab 被按住时"这个按键该不该由我们吃掉"。</summary>
        public bool ShouldTakeAccept() => Visible && _items.Count > 0;

        /// <summary>空格确认：选中项是关键字且"空格本来就是它后面该跟的东西"时才接管。</summary>
        public bool CanSpaceConfirm()
        {
            if (!Visible || _items.Count == 0) return false;
            var c = _items[_sel];
            if (!VbaBuiltins.IsKeyword(c.Name)) return false;
            // 已经打全了就没必要再吃空格（吃掉了反而少一个空格）
            return true;
        }
    }
}
