// 入口：单实例 + 启动；带上 --selftest 时只跑自检、不装钩子（这样能在没有
// 桌面的环境里验证全部纯逻辑）。

using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace VbeComplete
{
    internal static class Program
    {
        private static Mutex _mutex;

        /// <summary>已经有一份在跑时，用它叫那一份把托盘图标提示出来。</summary>
        public const string ShowEventName = @"Local\VbeComplete_CSharp_v1_show";

        [STAThread]
        private static int Main(string[] args)
        {
            // 崩溃必须留下痕迹。WinExe 没有控制台，CLR 把异常吞掉之后用户只看到
            // "窗口闪退、项目没启动"，我们手里连一行线索都没有。
            AppDomain.CurrentDomain.UnhandledException +=
                (s, e) => Log.Crash("AppDomain.UnhandledException", e.ExceptionObject as Exception);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException +=
                (s, e) => Log.Crash("Application.ThreadException", e.Exception);

            try { return RunMain(args); }
            catch (Exception e)
            {
                Log.Crash("Main", e);
                try
                {
                    MessageBox.Show(
                        "启动失败：" + e.Message + "\r\n\r\n详情见程序目录下的 crash.log",
                        "VbeComplete", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch { }
                return 1;
            }
        }

        private static int RunMain(string[] args)
        {
            Settings.Load();

            if (args.Length > 0 && args[0] == "--selftest")
            {
                return SelfTest.Run();
            }
            if (args.Length > 0 && args[0] == "--probe")
            {
                return Probe();
            }
            if (args.Length > 0 && args[0] == "--writetest")
            {
                return WriteTest();
            }
            if (args.Length > 0 && args[0] == "--uitest")
            {
                return UiTest();
            }
            if (args.Length > 0 && args[0] == "--carettest")
            {
                return CaretTest();
            }

            // 单实例：同名互斥体已经存在就说明有一份在跑。
            bool created;
            _mutex = new Mutex(true, "VbeComplete_CSharp_v1", out created);
            if (!created)
            {
                // 只写日志是不够的 —— 用户看到的就是"双击了，什么都没发生"。
                // 叫已经在跑的那一份弹个气泡，把托盘图标指出来。
                SignalExisting();
                Log.Write("已有一份在运行，本次退出（已提示已有实例）");
                return 0;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var app = new App();
            app.Start();

            // Ctrl+C / 关控制台 时收干净
            Console.CancelKeyPress += (s, e) => { e.Cancel = true; app.Shutdown(); };

            Application.Run(app);
            app.Shutdown();
            try { _mutex.ReleaseMutex(); } catch { }
            return 0;
        }

        private static void SignalExisting()
        {
            EventWaitHandle h;
            if (EventWaitHandle.TryOpenExisting(ShowEventName, out h))
            {
                try { h.Set(); } catch { }
                h.Dispose();
            }
        }

        /// <summary>
        /// 写入链路验证：在模块【末尾】追加一行测试代码 -> 用补全改它 -> 删掉。
        /// 全程可逆，最后逐字节比对原文；比对失败会把原文整段写进日志。
        /// ★只由用户显式运行 `--writetest` 触发，正常运行一次都不会动用户代码。
        /// </summary>
        private static int WriteTest()
        {
            var sb = new StringBuilder();
            var be = new VbeBackend();
            string orig = be.ModuleText();
            int count = be.LineCount();
            bool ok = true;

            sb.AppendLine("=== write test ===");
            sb.AppendLine("module lines = " + count + "  text len = " + (orig == null ? -1 : orig.Length));
            if (count <= 0)
            {
                sb.AppendLine("no code module, skip");
                File.WriteAllText(Path.Combine(Settings.Home, "writetest.log"), sb.ToString(), Encoding.UTF8);
                return 1;
            }

            try
            {
                const string probe = "    Dim abc";
                if (!be.InsertLines(count + 1, probe))
                {
                    sb.AppendLine("InsertLines FAILED");
                    ok = false;
                }
                else
                {
                    string back = be.LineAt(count + 1);
                    sb.AppendLine("insert readback = [" + back + "]");
                    bool same = back == probe;
                    if (!same) ok = false;
                    sb.AppendLine("insert roundtrip = " + same);

                    bool w = be.ApplyCompletion(count + 1, 9, 12, "abcd", false, false, null);
                    string after = be.LineAt(count + 1);
                    sb.AppendLine("apply ok=" + w + "  result=[" + after + "]");
                    sb.AppendLine("apply result correct = " + (after == "    Dim abcd"));
                    if (after != "    Dim abcd") ok = false;

                    be.SetCaret(count + 1, 12);
                    var s2 = be.Snapshot();
                    string got = s2 == null ? "null" : s2.Line + ":" + s2.CaretCol;
                    sb.AppendLine("caret after SetCaret(12) = " + got + " (want 13)");
                    if (s2 == null || s2.CaretCol != 13) ok = false;
                }
            }
            finally
            {
                // 无论成败都要把测试行删干净（最多重试 3 次）
                for (int i = 0; i < 3; i++)
                {
                    int c2 = be.LineCount();
                    if (c2 <= count) break;
                    be.DeleteLines(count + 1, c2 - count);
                }
            }

            string now = be.ModuleText();
            bool restored = now == orig;
            sb.AppendLine("restored = " + restored);
            if (!restored)
            {
                ok = false;
                sb.AppendLine("!!! 未能逐字节还原，原文如下（请手工检查模块末尾） !!!");
                sb.AppendLine(orig);
                sb.AppendLine("!!! 现在的全文 !!!");
                sb.AppendLine(now);
            }
            sb.AppendLine(ok && restored ? "RESULT: PASS" : "RESULT: FAIL");

            File.WriteAllText(Path.Combine(Settings.Home, "writetest.log"), sb.ToString(), Encoding.UTF8);
            return ok && restored ? 0 : 1;
        }

        /// <summary>
        /// 光标落点验证（真机）：模拟"打到一半 + Tab 补全"，读回 VBE 真实光标列，
        /// 跟"按读回来的真行算出来的列"对数，并把光标左右两截文本打出来。
        /// ★只在 `--carettest` 时执行；动的是模块末尾的临时行，finally 里删干净
        /// 并逐字节比对原文。
        /// </summary>
        private static int CaretTest()
        {
            var sb = new StringBuilder();
            var be = new VbeBackend();
            string orig = be.ModuleText();
            int count = be.LineCount();
            bool allOk = true;

            sb.AppendLine("=== caret test ===");
            sb.AppendLine("module lines = " + count + "  text len = " + (orig == null ? -1 : orig.Length));
            if (count <= 0)
            {
                sb.AppendLine("no code module, skip");
                File.WriteAllText(Path.Combine(Settings.Home, "carettest.log"), sb.ToString(), Encoding.UTF8);
                return 1;
            }

            // (标签, 插入的行, 已输入的词, 候选)
            // ⚠️ InsertLines 自己也会触发 VBE 的行格式化（`student=spl` 读回是
            //   `student = spl`），所以列**必须按读回来的真行**算 —— 这正是
            //   真机与离线用例的差别，也是这个 bug 之前没被发现的原因。
            var cases = new[]
            {
                ("顶格无缩进",        "student=spl",  "spl", "Split"),
                ("4 空格缩进",        "    student=spl", "spl", "Split"),
                ("Tab 缩进",          "\tstudent=spl", "spl", "Split"),
                ("Tab 缩进 + 全角名",  "\t学生=spl",   "spl", "Split"),
                ("右边还有同名子串",   "    abc=spl+Splitter", "spl", "Split"),
            };

            try
            {
                foreach (var c in cases)
                {
                    int line = count + 1;
                    if (!be.InsertLines(line, c.Item2))
                    {
                        sb.AppendLine(c.Item1 + ": InsertLines FAILED");
                        allOk = false;
                        continue;
                    }
                    string inserted = be.LineAt(line) ?? "";
                    int at0 = WholeWordIndex(inserted, c.Item3);
                    if (at0 < 0)
                    {
                        sb.AppendLine(c.Item1 + ": 读回行里找不到已输入的词 [" + inserted + "]");
                        allOk = false;
                        be.DeleteLines(line, 1);
                        continue;
                    }
                    int wordStart = at0 + 1;                  // 1-based 字符列
                    int endCol = at0 + 1 + c.Item3.Length;

                    bool w = be.ApplyCompletion(line, wordStart, endCol, c.Item4, true, false, null);
                    string actual = be.LineAt(line) ?? "";
                    var snap = be.Snapshot();
                    int got = snap == null ? -1 : snap.CaretCol;

                    // 期望列：补出来的括号在，光标就该停在 `(` 之前
                    int idx = actual.IndexOf('(');
                    int want = idx >= 0
                        ? Rules.DispCol(actual, idx, 4)
                        : Rules.DispCol(actual, at0 + c.Item4.Length, 4);

                    int at = CharIndexAtCol(actual, got, 4);
                    string left = at < 0 ? "?" : actual.Substring(0, at);
                    string right = at < 0 ? "?" : actual.Substring(at);

                    // 光标落在补全词后 ⇒ 按"显示列 -> 字符列"换算后取词，应当就是候选词。
                    // 含全角/缩进的行不换算会取到空词（= 这一行根本不弹候选）。
                    int ws;
                    string wordAtCaret = Rules.WordBefore(actual, Rules.CharColFromDisp(actual, got), out ws);
                    bool okWord = string.Equals(wordAtCaret, c.Item4, StringComparison.OrdinalIgnoreCase);

                    bool ok = w && got == want && okWord;
                    if (!ok) allOk = false;

                    sb.AppendLine("--- " + c.Item1 + " ---");
                    sb.AppendLine("  inserted  = [" + inserted.Replace("\t", "\\t") + "]");
                    sb.AppendLine("  actual    = [" + actual.Replace("\t", "\\t") + "]");
                    sb.AppendLine("  apply ok  = " + w + "   (wordStart=" + wordStart + " endCol=" + endCol + ")");
                    sb.AppendLine("  caret col = " + got + "   want = " + want + "   " + (got == want ? "OK" : "MISMATCH"));
                    sb.AppendLine("  word@caret= [" + wordAtCaret + "]  " + (okWord ? "OK" : "MISMATCH"));
                    sb.AppendLine("  left      = [" + left.Replace("\t", "\\t") + "]");
                    sb.AppendLine("  right     = [" + right.Replace("\t", "\\t") + "]");

                    be.DeleteLines(line, 1);
                }
            }
            finally
            {
                for (int i = 0; i < 3; i++)
                {
                    int c2 = be.LineCount();
                    if (c2 <= count) break;
                    be.DeleteLines(count + 1, c2 - count);
                }
            }

            string now = be.ModuleText();
            bool restored = now == orig;
            sb.AppendLine("restored = " + restored);
            if (!restored) allOk = false;
            sb.AppendLine(allOk ? "RESULT: PASS" : "RESULT: FAIL");
            File.WriteAllText(Path.Combine(Settings.Home, "carettest.log"), sb.ToString(), Encoding.UTF8);
            return allOk ? 0 : 1;
        }

        /// <summary>把"显示列"换回字符下标（找不到返回 -1）。测试用，只在 ASCII/全角混合行上够用。</summary>
        private static int CharIndexAtCol(string text, int col, int tabw)
        {
            for (int i = 0; i <= text.Length; i++)
                if (Rules.DispCol(text, i, tabw) == col) return i;
            return -1;
        }

        /// <summary>整词查找（前后都不是标识符字符），找不到返回 -1。测试用。</summary>
        private static int WholeWordIndex(string text, string word)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(word)) return -1;
            int i = 0;
            while (i <= text.Length - word.Length)
            {
                int hit = text.IndexOf(word, i, StringComparison.OrdinalIgnoreCase);
                if (hit < 0) return -1;
                bool okL = hit == 0 || !Rules.IsIdentChar(text[hit - 1]);
                bool okR = hit + word.Length >= text.Length || !Rules.IsIdentChar(text[hit + word.Length]);
                if (okL && okR) return hit;
                i = hit + 1;
            }
            return -1;
        }

        /// <summary>
        /// 候选窗离屏渲染：把窗口画进一张 PNG，用来一眼核对配色 / 图标 / 布局。
        /// 不需要桌面，也不会出现在屏幕上。
        /// </summary>
        private static int UiTest()
        {
            var sb = new StringBuilder();
            try
            {
                var f = new PopupForm();
                var items = new System.Collections.Generic.List<Candidate>
                {
                    new Candidate { Name = "MsgBox",   Kind = "func", Pos = 0, Span = 2 },
                    new Candidate { Name = "Main",     Kind = "func", Pos = 0, Span = 2 },
                    new Candidate { Name = "mCount",   Kind = "var",  Pos = 1, Span = 1 },
                    new Candidate { Name = "Size",     Kind = "prop", Pos = 0, Span = 2 },
                    new Candidate { Name = "String",   Kind = "type", Pos = 0, Span = 2 },
                    new Candidate { Name = "Select",   Kind = "kw",   Pos = 0, Span = 2 },
                    new Candidate { Name = "selBegin", Kind = "var",  Pos = 0, Span = 2 },
                };
                f.SetData(items, 1, 0, 8);
                int w = 260;
                int h = items.Count * f.RowHeight + f.DetailHeight + 2;
                f.StartPosition = FormStartPosition.Manual;
                f.Location = new System.Drawing.Point(-4000, -4000);
                f.Show();
                f.Width = w; f.Height = h;
                Application.DoEvents();

                using (var bmp = new System.Drawing.Bitmap(w, h))
                {
                    f.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, w, h));
                    string path = Path.Combine(Settings.Home, "popup-preview.png");
                    bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                    sb.AppendLine("saved " + path + "  " + w + "x" + h);
                }
                f.Hide();
                f.Dispose();
                sb.AppendLine("RESULT: PASS");
                File.WriteAllText(Path.Combine(Settings.Home, "uitest.log"), sb.ToString(), Encoding.UTF8);
                return 0;
            }
            catch (Exception e)
            {
                sb.AppendLine("EXCEPTION " + e);
                File.WriteAllText(Path.Combine(Settings.Home, "uitest.log"), sb.ToString(), Encoding.UTF8);
                return 1;
            }
        }

        /// <summary>诊断用：把"能不能连上 VBE"写进 probe.log。</summary>
        private static int Probe()
        {
            var sb = new StringBuilder();
            sb.AppendLine("Home = " + Settings.Home);
            sb.AppendLine("settings.ini = " + Path.Combine(Settings.Home, "settings.ini"));
            sb.AppendLine("settings error = " + Settings.LastError());
            sb.AppendLine("kernel = " + Environment.Version);
            sb.AppendLine("64bit = " + Environment.Is64BitProcess);

            var be = new VbeBackend();
            sb.AppendLine("--- COM ---");
            foreach (var progid in new[] { "Excel.Application", "Word.Application" })
            {
                object app = ComCtx.GetActive(progid);
                sb.AppendLine(progid + " active = " + (app == null ? "null" : "ok"));
                if (app == null) continue;
                try
                {
                    dynamic dapp = app;
                    dynamic vbe = dapp.VBE;
                    sb.AppendLine("  VBE = " + (vbe == null ? "null" : "ok"));
                    if (vbe == null) continue;
                    try
                    {
                        dynamic projs = vbe.VBProjects;
                        int pc = Convert.ToInt32(projs.Count);
                        sb.AppendLine("  projects = " + pc);
                        if (pc > 0)
                        {
                            dynamic p0 = projs.Item(1);
                            sb.AppendLine("  project[1].Name = " + Convert.ToString(p0.Name));
                            dynamic comps = p0.VBComponents;
                            sb.AppendLine("  components = " + Convert.ToInt32(comps.Count));
                        }
                    }
                    catch (Exception e) { sb.AppendLine("  projects ERR " + e.Message); }
                    try
                    {
                        dynamic cp = vbe.ActiveCodePane;
                        sb.AppendLine("  ActiveCodePane = " + (cp == null ? "null" : "ok"));
                        if (cp != null)
                        {
                            dynamic cm = cp.CodeModule;
                            sb.AppendLine("  module = " + Convert.ToString(cm.Name) +
                                          " lines = " + Convert.ToInt32(cm.CountOfLines));
                            try
                            {
                                dynamic sel = cp.Selection;
                                sb.AppendLine("  Selection = " +
                                    Convert.ToInt32(sel.StartLine) + "," + Convert.ToInt32(sel.StartColumn) +
                                    " .. " + Convert.ToInt32(sel.EndLine) + "," + Convert.ToInt32(sel.EndColumn));
                            }
                            catch (Exception e) { sb.AppendLine("  Selection ERR " + e.Message); }
                            try
                            {
                                object g = cp.GetSelection();
                                sb.AppendLine("  GetSelection = " + (g == null ? "null" : g.GetType().FullName));
                            }
                            catch (Exception e) { sb.AppendLine("  GetSelection ERR " + e.Message); }
                        }
                    }
                    catch (Exception e) { sb.AppendLine("  ActiveCodePane ERR " + e.Message); }
                }
                catch (Exception e) { sb.AppendLine("  VBE ERR " + e.Message); }
            }

            sb.AppendLine("--- backend ---");
            var snap = be.Snapshot();
            sb.AppendLine("snapshot = " + (snap == null ? "null" : snap.ToString()));
            if (snap != null)
            {
                sb.AppendLine("line text = " + (snap.Text ?? ""));
                sb.AppendLine("module text len = " + (be.ModuleText()?.Length ?? -1));
                sb.AppendLine("components = " + string.Join(", ", be.ComponentNames()));
            }
            sb.AppendLine("--- settings ---");
            sb.AppendLine("AUTO_SPACE = " + Settings.Flag("NO_AUTOSPACE", false));
            sb.AppendLine("HIDE_ON_COMPLETE = " + Settings.Flag("HIDE_ON_COMPLETE", true));
            sb.AppendLine("VIEW_ROWS = " + Settings.Int("VIEW_ROWS", 8));
            sb.AppendLine("--- win32 ---");
            sb.AppendLine("inVbeCodePane(win32) = " + App_invbe());
            File.WriteAllText(Path.Combine(Settings.Home, "probe-cs.log"), sb.ToString(), Encoding.UTF8);
            return 0;
        }

        private static string App_invbe()
        {
            IntPtr frame = Native.FindWindow("wndclass_desked_gsk", null);
            if (frame == IntPtr.Zero) return "no vbe frame";
            uint pid;
            uint tid = Native.GetWindowThreadProcessId(frame, out pid);
            IntPtr f = Native.FocusedHwndOf(tid);
            return "frame=" + frame + " pid=" + pid + " focusClass=" + Native.ClassNameOf(f);
        }
    }
}
