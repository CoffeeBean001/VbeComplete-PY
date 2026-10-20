// 主流程：全局钩子 + 自适应轮询 + 按键路由 + 各种"改一行"的操作。
//
// ★这是整个工具唯一允许"决定吃不吃一个按键"的地方，所以三条铁律在这里落地：
//   1. 钩子线程只读内存里的状态（_codePaneFocused / _popupVisible / _lastKeyTs），
//      一次 COM 都不许碰 —— 碰一次就可能让全系统的键盘卡住；
//   2. 真动作一律投递到 UI 线程执行；
//   3. 吃掉的键如果最后没做成事，必须把键还回去（绝不"吞了但什么都没发生"）。
//
// ★CPU 上花的功夫也在这里：
//   * 前台不是 VBE 代码窗时【完全不调 COM】—— 用户在 Excel 表格里、
//     在浏览器里的时候，我们一个 COM 都不发；
//   * 在代码窗里也按"刚敲过键没有"来调频率：打字时 50ms，停手 200ms。

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace VbeComplete
{
    internal sealed class App : ApplicationContext
    {
        private const string CodePaneClass = "VbaWindow";
        private const string VbeFrameClass = "wndclass_desked_gsk";
        private static readonly string[] VbePopupClasses = { "NameListWndClass", "PopupTipWndClass" };

        private readonly VbeBackend _be = new VbeBackend();
        private readonly Recent _recent = new Recent();
        private readonly WordBank _bank = new WordBank();
        private readonly Engine _eng;
        private readonly PopupForm _popup = new PopupForm();
        private readonly Control _disp = new Control();

        private System.Windows.Forms.Timer _timer;
        private NotifyIcon _tray;
        private EventWaitHandle _showEvent;
        private Thread _showThread;

        // ---- 跨线程共享的状态（钩子线程只读，UI 线程写） ----
        private volatile bool _codePaneFocused;
        private volatile bool _popupVisible;
        private volatile bool _running = true;
        private long _lastKeyTs;

        // ---- 钩子 ----
        private Thread _hookThread;
        private IntPtr _hookHandle = IntPtr.Zero;
        private uint _hookThreadId;
        private Native.LowLevelKeyboardProc _hookProc;   // 必须保持引用，否则被 GC

        // ---- 轮询状态 ----
        private Snapshot _lastSnap;
        private string _lastModName = "";
        private DateTime _lastPollAt = DateTime.MinValue;

        private readonly bool NoCtrlNav = Settings.Flag("NO_CTRL_NAV", false);
        private readonly bool NoSnippets = Settings.Flag("NO_SNIPPETS", false);

        public App()
        {
            // 强制建句柄：跨线程 BeginInvoke 需要一个真窗口
            var _ = _disp.Handle;

            _be.SendChar = ch => Native.SendUnicodeChar(ch);
            _be.SendVk = (vk, ctrl, shift) => Native.SendVk(vk, ctrl, shift);

            _eng = new Engine
            {
                Backend = _be,
                Recent = _recent,
                Bank = _bank,
            };
            _eng.OnShow = (items, sel, top) =>
            {
                _popup.SetData(items, sel, top, _eng.ViewRows);
                ShowPopupAtCaret();
                _popupVisible = true;
            };
            _eng.OnHide = () =>
            {
                _popupVisible = false;
                if (_popup.Visible) _popup.Hide();
            };
            _eng.OnMove = () =>
            {
                _popup.UpdateView(_eng.SelectedIndex, _eng.TopIndex);
            };

            _popup.RowClicked += i => { _eng.SetSelected(i); _eng.Accept(); };
            _popup.Wheel += lines => _eng.ScrollBy(lines);
        }

        public void Start()
        {
            SetupTray();
            StartHook();
            StartShowListener();
            _timer = new System.Windows.Forms.Timer { Interval = 200 };
            _timer.Tick += (s, e) => PollOnce();
            _timer.Start();
            Log.Write("C# 版已启动（Home={0}）", Settings.Home);

            // 启动了就必须让用户看得见。这个程序没有主窗口，图标又默认折叠进
            // Win11 的托盘溢出区 —— 不给提示的话，"双击了没反应"和"已经正常运行"
            // 在屏幕上长得一模一样（用户就是这么报的 bug）。
            ShowBalloon("VbeComplete 已启动",
                "它没有主窗口，图标在右下角托盘的溢出区里。右键那个图标可以退出。");
        }

        private void ShowBalloon(string title, string text)
        {
            try
            {
                if (_tray == null) return;
                _tray.BalloonTipTitle = title;
                _tray.BalloonTipText = text;
                _tray.BalloonTipIcon = ToolTipIcon.Info;
                _tray.ShowBalloonTip(5000);
            }
            catch (Exception e) { Log.Write("气泡提示失败：{0}", e.Message); }
        }

        /// <summary>
        /// 监听"又有人启动了第二个实例"。Program 那边发现互斥体已存在时会 Set
        /// 这个事件，我们收到就把托盘图标的存在感刷一遍 —— 否则用户双击后的
        /// 全部观感就是"什么都没发生"。
        /// </summary>
        private void StartShowListener()
        {
            try
            {
                _showEvent = new EventWaitHandle(false, EventResetMode.AutoReset, Program.ShowEventName);
                _showThread = new Thread(() =>
                {
                    while (_running)
                    {
                        try
                        {
                            if (!_showEvent.WaitOne(500)) continue;
                            if (!_running) return;
                            // 必须回 UI 线程碰 NotifyIcon
                            if (_disp.InvokeRequired)
                                _disp.BeginInvoke((Action)(() => ShowBalloon(
                                    "VbeComplete 已经在运行",
                                    "不用再启动一次。图标在右下角托盘的溢出区里，右键可以退出。")));
                            else
                                ShowBalloon("VbeComplete 已经在运行",
                                    "不用再启动一次。图标在右下角托盘的溢出区里，右键可以退出。");
                        }
                        catch { return; }
                    }
                })
                { IsBackground = true, Name = "show-listener" };
                _showThread.Start();
            }
            catch (Exception e) { Log.Write("重复启动监听未装上：{0}", e.Message); }
        }

        /// <summary>
        /// 托盘图标：没有它，用户只能去任务管理器杀进程。
        /// ★图标必须用【系统自带】的 —— 不引入任何资源文件，保持单文件干净。
        /// </summary>
        private void SetupTray()
        {
            try
            {
                var menu = new ContextMenuStrip();
                var title = new ToolStripMenuItem("VbeComplete · C# 版") { Enabled = false };
                menu.Items.Add(title);
                menu.Items.Add(new ToolStripSeparator());
                menu.Items.Add("打开日志", null, (s, e) =>
                {
                    try { System.Diagnostics.Process.Start("notepad.exe", Log.Path_); } catch { }
                });
                menu.Items.Add("退出", null, (s, e) => Shutdown());

                _tray = new NotifyIcon
                {
                    Icon = System.Drawing.SystemIcons.Application,
                    Text = "VbeComplete（C# 版）",
                    Visible = true,
                    ContextMenuStrip = menu,
                };
                _tray.DoubleClick += (s, e) =>
                {
                    var snap = _be.Snapshot();
                    Log.Write("托盘双击：snapshot={0}", snap == null ? "null" : snap.ToString());
                };
            }
            catch (Exception e) { Log.Write("托盘创建失败：{0}", e.Message); }
        }

        // ------------------------------------------------------------------
        // 钩子线程
        // ------------------------------------------------------------------

        private void StartHook()
        {
            _hookThread = new Thread(HookLoop)
            {
                IsBackground = true,
                Name = "kbd-hook",
            };
            _hookThread.Start();
        }

        private void HookLoop()
        {
            _hookThreadId = Native.GetCurrentThreadId();
            _hookProc = HookProc;
            IntPtr mod = Native.GetModuleHandle(null);
            _hookHandle = Native.SetWindowsHookEx(Native.WH_KEYBOARD_LL, _hookProc, mod, 0);
            if (_hookHandle == IntPtr.Zero)
            {
                Log.Write("装钩子失败，错误码 {0}", Marshal.GetLastWin32Error());
                return;
            }
            Log.Write("钩子已装上（tid={0}）", _hookThreadId);

            Native.MSG msg;
            while (Native.GetMessage(out msg, IntPtr.Zero, 0, 0) > 0)
            {
                Native.TranslateMessage(ref msg);
                Native.DispatchMessage(ref msg);
            }
            Native.UnhookWindowsHookEx(_hookHandle);
            _hookHandle = IntPtr.Zero;
        }

        private IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0) return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);

            int msg = wParam.ToInt32();
            if (msg != Native.WM_KEYDOWN && msg != Native.WM_SYSKEYDOWN)
                return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);

            var kb = (Native.KBDLLHOOKSTRUCT)Marshal.PtrToStructure(
                lParam, typeof(Native.KBDLLHOOKSTRUCT));

            // 注入的键（我们自己发的空格/括号）不算"用户按键"：否则会自己触发自己
            if ((kb.flags & Native.LLKHF_INJECTED) != 0)
                return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);

            // ★时标打在【这里】：放到别的回调里就晚了（那时钩子已经把事件
            //   post 给监听器，文本可能已经变了）
            _lastKeyTs = Environment.TickCount64;

            int vk = (int)kb.vkCode;

            // 不在 VBE 代码窗：一个键都不碰
            if (!_codePaneFocused)
                return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);

            bool ctrl = Native.Down(Native.VK_CONTROL);
            bool shift = Native.Down(Native.VK_SHIFT);
            bool alt = Native.Down(Native.VK_MENU);
            bool win = Native.Down(Native.VK_LWIN) || Native.Down(Native.VK_RWIN);

            if (win) return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);

            // ---- 1. 候选窗可见时的导航键 ----
            if (_popupVisible && !ctrl && !alt)
            {
                if (vk == Native.VK_DOWN && !shift)
                {
                    Post(() => _eng.Move(1));
                    return (IntPtr)1;
                }
                if (vk == Native.VK_UP && !shift)
                {
                    Post(() => _eng.Move(-1));
                    return (IntPtr)1;
                }
                if (vk == Native.VK_NEXT)
                {
                    Post(() => _eng.Move(_eng.PageSize));
                    return (IntPtr)1;
                }
                if (vk == Native.VK_PRIOR)
                {
                    Post(() => _eng.Move(-_eng.PageSize));
                    return (IntPtr)1;
                }
                if (vk == Native.VK_TAB && !shift)
                {
                    Post(() => { if (!_eng.Accept()) Native.SendVk(Native.VK_TAB); });
                    return (IntPtr)1;
                }
                if (vk == Native.VK_RETURN)
                {
                    Post(() => { if (!_eng.Accept()) Native.SendVk(Native.VK_RETURN); });
                    return (IntPtr)1;
                }
                if (vk == Native.VK_ESCAPE)
                {
                    Post(() => _eng.Hide());
                    return (IntPtr)1;
                }
            }

            // ---- 2. 手动唤出 ----
            if (ctrl && vk == Native.VK_SPACE)
            {
                Post(() =>
                {
                    var snap = _be.Snapshot();
                    if (snap != null) _eng.Trigger(true, snap);
                });
                return (IntPtr)1;
            }

            // ---- 3. 行操作 ----
            if (ctrl && shift && vk == Native.VK_RETURN)
            {
                Post(InsertLineAboveHere);
                return (IntPtr)1;
            }
            if (ctrl && !shift && vk == 'D')
            {
                Post(DuplicateLineHere);
                return (IntPtr)1;
            }
            if (ctrl && !shift && vk == 'K')
            {
                Post(DeleteLineHere);
                return (IntPtr)1;
            }
            // ⚠️ Shift+D / Shift+K 绝不能接管 —— 那是输入大写字母
            if (ctrl && !shift && vk == Native.VK_OEM_2)
            {
                Post(ToggleCommentHere);
                return (IntPtr)1;
            }
            if (ctrl && shift && (vk == Native.VK_UP || vk == Native.VK_DOWN) && !NoCtrlNav)
            {
                int dir = vk == Native.VK_UP ? -1 : 1;
                Post(() => CtrlNavJumpHere(dir));
                return (IntPtr)1;
            }

            return Native.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
        }

        private void Post(Action a)
        {
            try { _disp.BeginInvoke(a); }
            catch (Exception e) { Log.Write("投递动作失败：{0}", e.Message); }
        }

        // ------------------------------------------------------------------
        // 轮询
        // ------------------------------------------------------------------

        /// <summary>纯 Win32：焦点是不是在 VBE 的代码窗格里（0.02ms，不碰 COM）。</summary>
        private static bool InVbeCodePane()
        {
            IntPtr frame = Native.FindWindow(VbeFrameClass, null);
            if (frame == IntPtr.Zero) return false;
            uint pid;
            uint tid = Native.GetWindowThreadProcessId(frame, out pid);
            IntPtr focus = Native.FocusedHwndOf(tid);
            if (focus == IntPtr.Zero) return false;

            string cls = Native.ClassNameOf(focus);
            foreach (var p in VbePopupClasses)
                if (string.Equals(cls, p, StringComparison.OrdinalIgnoreCase)) return true;

            IntPtr h = focus;
            for (int i = 0; i < 12 && h != IntPtr.Zero; i++)
            {
                if (string.Equals(Native.ClassNameOf(h), CodePaneClass, StringComparison.OrdinalIgnoreCase))
                    return true;
                if (h == frame) break;
                IntPtr p = Native.GetParent(h);
                if (p == h) break;
                h = p;
            }
            // 拿不准（类名不认识）：退回"前台窗口属于 VBE 进程"
            string fgCls = Native.ClassNameOf(Native.GetForegroundWindow());
            return string.Equals(fgCls, VbeFrameClass, StringComparison.OrdinalIgnoreCase);
        }

        private void PollOnce()
        {
            if (!_running) return;
            try
            {
                bool focused = InVbeCodePane();
                _codePaneFocused = focused;

                if (!focused)
                {
                    if (_popupVisible) _eng.Hide();
                    _lastSnap = null;
                    _lastModName = "";
                    SetInterval(250);
                    return;
                }

                Snapshot snap = _be.Snapshot();
                if (snap == null)
                {
                    if (_popupVisible) _eng.Hide();
                    SetInterval(150);
                    return;
                }

                bool typing = Environment.TickCount64 - _lastKeyTs < 900;
                SetInterval(typing ? 50 : 150);

                bool changed = _lastSnap == null
                    || snap.Line != _lastSnap.Line
                    || !string.Equals(Rules.Rstrip(snap.Text), Rules.Rstrip(_lastSnap.Text), StringComparison.Ordinal)
                    || !string.Equals(snap.Module, _lastModName, StringComparison.OrdinalIgnoreCase);

                if (changed)
                {
                    // 换模块 = 不是打字，只是切了代码窗，不该弹
                    bool modSwitched = !string.Equals(snap.Module, _lastModName, StringComparison.OrdinalIgnoreCase)
                                       && _lastSnap != null;
                    _lastSnap = snap;
                    _lastModName = snap.Module;

                    if (!modSwitched)
                    {
                        _eng.InvalidatePool();
                        if (_popupVisible) _eng.Hide();
                        _eng.Trigger(false, snap);
                    }
                }
                else if (_popupVisible)
                {
                    // 光标可能在同一行里移动了：跟着走
                    ShowPopupAtCaret();
                }
            }
            catch (Exception e)
            {
                Log.Write("轮询异常：{0}", e.Message);
            }
        }

        private void SetInterval(int ms)
        {
            if (_timer == null) return;
            if (_timer.Interval != ms) _timer.Interval = ms;
        }

        private void ShowPopupAtCaret()
        {
            int x, top, bottom;
            if (!Native.CaretRect(out x, out top, out bottom))
            {
                // 拿不到光标就收起 —— 绝不猜一个位置糊上去
                _eng.Hide();
                return;
            }
            if (!_popup.Visible) _popup.Show();
            _popup.PlaceAt(x, top, bottom);
        }

        // ------------------------------------------------------------------
        // 行操作（全部在 UI 线程执行）
        // ------------------------------------------------------------------

        private void InsertLineAboveHere()
        {
            var snap = _be.Snapshot();
            if (snap == null) return;
            string line = snap.Text ?? "";
            string indent = LeadingWs(line);
            if (_be.InsertLines(snap.Line, indent))
                _be.SetCaret(snap.Line, indent.Length + 1);
        }

        private void DuplicateLineHere()
        {
            var snap = _be.Snapshot();
            if (snap == null) return;
            string line = Rules.Rstrip(snap.Text ?? "");
            if (_be.InsertLines(snap.Line + 1, line))
                _be.SetCaret(snap.Line + 1, charColOf(snap, line));
        }

        private static int charColOf(Snapshot s, string line)
        {
            // 光标列是"显示列"，这里只需要"同一行靠左的合理位置"
            int n = Math.Min(line.Length, Math.Max(0, s.CaretCol - 1));
            return n + 1;
        }

        /// <summary>
        /// Ctrl+K：删掉当前整行，光标落到【上一行】的行尾。
        /// 删第 1 行 -> 新首行首列；模块删空 -> 不落光标，但照样算成功。
        /// </summary>
        private void DeleteLineHere()
        {
            var snap = _be.Snapshot();
            if (snap == null) return;
            int count = _be.LineCount();
            if (count <= 0) return;

            if (count == 1)
            {
                _be.ReplaceLineRaw(1, "");
                return;
            }

            if (!_be.DeleteLines(snap.Line, 1)) return;

            if (snap.Line <= 1) _be.SetCaret(1, 1);
            else _be.CaretToLineEnd(snap.Line - 1);
        }

        private void ToggleCommentHere()
        {
            var snap = _be.Snapshot();
            if (snap == null) return;

            // 有没有选区？有就对每行加/去
            int from = snap.Line, to = snap.Line;
            int sl, sc, el, ec;
            if (_be.TryGetSelection(out sl, out sc, out el, out ec))
            {
                from = Math.Min(sl, el);
                to = Math.Max(sl, el);
                if (ec == 1 && to > from) to--;
            }

            var lines = new List<string>();
            for (int i = from; i <= to; i++)
            {
                string t = _be.LineAt(i);
                if (t == null) break;
                lines.Add(Rules.Rstrip(t));
            }
            if (lines.Count == 0) return;

            bool allCommented = true;
            foreach (var l in lines)
            {
                string tr = l.TrimStart();
                if (tr.Length == 0) continue;
                if (!tr.StartsWith("'", StringComparison.Ordinal)) { allCommented = false; break; }
            }

            for (int i = 0; i < lines.Count; i++)
            {
                string l = lines[i];
                string nl;
                if (allCommented)
                {
                    int idx = l.IndexOf('\'');
                    nl = idx >= 0 ? l.Remove(idx, 1) : l;
                }
                else
                {
                    int indent = LeadingWs(l).Length;
                    nl = l.Substring(0, indent) + "'" + l.Substring(indent);
                }
                _be.ReplaceLineRaw(from + i, nl);
            }
        }

        /// <summary>
        /// Ctrl+Shift+↑：落到本过程里【最后一个声明行】的行尾；
        /// Ctrl+Shift+↓：跳到过程收尾的空白行。
        /// </summary>
        private void CtrlNavJumpHere(int dir)
        {
            var snap = _be.Snapshot();
            if (snap == null) return;
            string text = _be.ModuleText();
            if (text == null) return;

            var lines = Parser.SplitLines(text);
            var procs = Parser.ExtractProcedures(text);
            var proc = Parser.ProcOfLine(procs, snap.Line);
            int start = proc?.StartLine ?? 1;
            int end = proc?.EndLine ?? lines.Length;

            if (dir < 0)
            {
                int target = -1;
                for (int i = snap.Line - 1; i >= start; i--)
                {
                    string code = Parser.CodeOnly(lines[i - 1]).Trim();
                    if (code.Length == 0) continue;
                    if (Regex.IsMatch(code, @"^(?:Dim|Const|Static|Private|Public|Global|ReDim)\b",
                            RegexOptions.IgnoreCase))
                    {
                        target = i;
                        break;
                    }
                    if (Regex.IsMatch(code, @"^(?:Sub|Function|Property)\b", RegexOptions.IgnoreCase))
                        break;
                }
                if (target > 0) _be.CaretToLineEnd(target);
            }
            else
            {
                for (int i = snap.Line + 1; i <= end; i++)
                {
                    string code = Parser.CodeOnly(lines[i - 1]).Trim();
                    if (code.Length == 0)
                    {
                        _be.SetCaret(i, 1);
                        return;
                    }
                }
                // 没有空白行：落到过程末尾
                if (end >= 1) _be.CaretToLineEnd(Math.Max(1, end - 1));
            }
        }

        private static string LeadingWs(string s)
        {
            int i = 0;
            while (i < s.Length && (s[i] == ' ' || s[i] == '\t')) i++;
            return s.Substring(0, i);
        }

        // ------------------------------------------------------------------
        // 退出
        // ------------------------------------------------------------------

        public void Shutdown()
        {
            _running = false;
            try { _timer?.Stop(); } catch { }
            try { _recent.Flush(); } catch { }
            try
            {
                if (_hookThreadId != 0)
                    Native.PostThreadMessage(_hookThreadId, Native.WM_QUIT, IntPtr.Zero, IntPtr.Zero);
            }
            catch { }
            try { if (_hookHandle != IntPtr.Zero) Native.UnhookWindowsHookEx(_hookHandle); } catch { }
            try { _showEvent?.Dispose(); } catch { }
            try { _popup.Hide(); _popup.Dispose(); } catch { }
            try
            {
                if (_tray != null) { _tray.Visible = false; _tray.Dispose(); _tray = null; }
            }
            catch { }
            Log.Write("已退出");
            ExitThread();
        }
    }
}
