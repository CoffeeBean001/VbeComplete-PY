// 候选窗：无边框、置顶、【永不抢焦点】、全部自绘。
//
// ★三条不能动的规矩：
//   1. 绝不抢焦点（WS_EX_NOACTIVATE + ShowWithoutActivation）—— 抢一下焦点，
//      用户在 VBE 里的输入就断了，整个工具等于废掉；
//   2. 全部自绘（一个 OnPaint 画完）—— 每行一个 Control 的话，滚一下要重排
//      几十个句柄，那才是"卡"的来源；
//   3. 拿不到光标矩形就【不显示】，绝不猜一个位置糊上去。

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

namespace VbeComplete
{
    internal sealed class PopupForm : Form
    {
        // ---- 配色（与 python 版 ui.py 顶部常量逐项对应） ----
        private static readonly Color CBg = ColorTranslator.FromHtml("#FFFFFF");
        private static readonly Color CFg = ColorTranslator.FromHtml("#202020");
        private static readonly Color CHit = ColorTranslator.FromHtml("#C00000");
        private static readonly Color CSelBg = ColorTranslator.FromHtml("#2A5CAA");
        private static readonly Color CSelFg = ColorTranslator.FromHtml("#FFFFFF");
        private static readonly Color CSelHit = ColorTranslator.FromHtml("#FFD24A");
        private static readonly Color CDetailBg = ColorTranslator.FromHtml("#F7F7F7");
        private static readonly Color CDetailFg = ColorTranslator.FromHtml("#333333");
        private static readonly Color CDetailSep = ColorTranslator.FromHtml("#E0E0E0");
        private static readonly Color CTrack = ColorTranslator.FromHtml("#F0F0F0");
        private static readonly Color CThumb = ColorTranslator.FromHtml("#C8C8C8");
        private static readonly Color CBorder = ColorTranslator.FromHtml("#B8B8B8");

        private static readonly Dictionary<string, Color> IconColor =
            new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase)
            {
                { "func", ColorTranslator.FromHtml("#1F8A3B") },
                { "prop", ColorTranslator.FromHtml("#7A3FA8") },
                { "var",  ColorTranslator.FromHtml("#2A5CAA") },
                { "type", ColorTranslator.FromHtml("#B26A00") },
                { "kw",   ColorTranslator.FromHtml("#8A8A8A") },
            };

        private static readonly Dictionary<string, Color> IconColorOn =
            new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase)
            {
                { "func", ColorTranslator.FromHtml("#6FE08C") },
                { "prop", ColorTranslator.FromHtml("#CBA6F7") },
                { "var",  ColorTranslator.FromHtml("#9CC2F5") },
                { "type", ColorTranslator.FromHtml("#FFC46B") },
                { "kw",   ColorTranslator.FromHtml("#D8D8D8") },
            };

        private const int PadX = 6;
        private const int IconGap = 4;
        private const int ScrollW = 8;
        private const int MaxWidth = 520;
        private const int MinWidth = 150;

        private IReadOnlyList<Candidate> _items = Array.Empty<Candidate>();
        private int _sel;
        private int _top;
        private int _rows = 8;
        private int _rowH = 20;
        private int _detailH = 18;
        private int _iconW = 12;

        private readonly Font _font;

        public event Action<int> RowClicked;      // 双击/单击某行
        public event Action<int> Wheel;           // 滚轮行数

        public PopupForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            TopMost = true;
            BackColor = CBg;
            DoubleBuffered = true;
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);

            _font = BuildFont(out float pt);
            _rowH = Math.Max(16, _font.Height + 6);
            _detailH = Math.Max(14, _font.Height + 4);
            _iconW = Math.Max(9, Math.Min(16, (int)Math.Round(_rowH * 0.62)));

            // 窗口本身不接受键盘 —— 所有按键都由全局钩子处理
            TabStop = false;
        }

        /// <summary>
        /// 跟 VBE 编辑器用同一个字体、同一档字号 —— 候选窗的字看起来才像"编辑器
        /// 自己弹出来的"。注册表读不到就退回 Consolas 10。
        /// </summary>
        private static Font BuildFont(out float pt)
        {
            double want = Settings.Double("POPUP_FONT_SIZE", 0);
            string face = Settings.Raw("POPUP_FONT");
            try
            {
                if (string.IsNullOrEmpty(face) || want <= 0)
                {
                    using (var k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\VBA\6.0\Common"))
                    {
                        if (k != null)
                        {
                            if (string.IsNullOrEmpty(face))
                                face = k.GetValue("FontFace") as string;
                            if (want <= 0)
                            {
                                object h = k.GetValue("FontHeight");
                                if (h is int i && i > 0) want = i;
                            }
                        }
                    }
                }
            }
            catch { }

            if (want <= 0) want = 10;
            if (want < 6) want = 6;
            if (want > 40) want = 40;
            pt = (float)want;
            if (string.IsNullOrEmpty(face)) face = "Consolas";

            try
            {
                var f = new Font(face, pt, FontStyle.Regular);
                if (string.Equals(f.Name, face, StringComparison.OrdinalIgnoreCase)) return f;
                f.Dispose();
            }
            catch { }
            return new Font("Consolas", pt, FontStyle.Regular);
        }

        protected override bool ShowWithoutActivation => true;

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x08000000;   // WS_EX_NOACTIVATE：点了也不抢焦点
                cp.ExStyle |= 0x00000080;   // WS_EX_TOOLWINDOW：不出现在 Alt-Tab
                return cp;
            }
        }

        public IReadOnlyList<Candidate> Items => _items;
        public int RowHeight => _rowH;
        public int DetailHeight => _detailH;

        public void SetData(IReadOnlyList<Candidate> items, int sel, int top, int rows)
        {
            _items = items ?? Array.Empty<Candidate>();
            _sel = sel;
            _top = top;
            _rows = Math.Max(1, rows);
            Invalidate();
        }

        public void UpdateView(int sel, int top)
        {
            _sel = sel;
            _top = top;
            Invalidate();
        }

        /// <summary>
        /// 贴着光标显示。(x, top, bottom) 是插入符的屏幕坐标。
        /// 下方放不下就翻到光标上方 —— 但绝不离开屏幕工作区。
        /// </summary>
        public void PlaceAt(int x, int top, int bottom)
        {
            int visibleRows = Math.Min(_rows, Math.Max(1, _items.Count));
            int w = ComputeWidth();
            int h = visibleRows * _rowH + _detailH + 2;

            var scr = Screen.FromPoint(new Point(x, Math.Max(0, bottom)));
            Rectangle wa = scr.WorkingArea;

            int px = x;
            int py = bottom + 2;
            if (py + h > wa.Bottom) py = top - h - 2;
            if (py < wa.Top) py = wa.Top;
            if (px + w > wa.Right) px = wa.Right - w;
            if (px < wa.Left) px = wa.Left;

            Native.SetWindowPos(Handle, Native.HWND_TOPMOST, px, py, w, h,
                Native.SWP_NOACTIVATE);
            Width = w; Height = h;
        }

        private int ComputeWidth()
        {
            int max = 0;
            using (var g = CreateGraphics())
            {
                int n = Math.Min(_items.Count, 200);
                for (int i = 0; i < n; i++)
                {
                    var s = _items[i];
                    if (s == null || string.IsNullOrEmpty(s.Name)) continue;
                    int tw = TextRenderer.MeasureText(g, s.Name, _font, new Size(int.MaxValue, _rowH),
                        TextFormatFlags.NoPadding).Width;
                    if (tw > max) max = tw;
                }
            }
            int w = max + PadX * 2 + _iconW + IconGap + ScrollW + 4;
            return Math.Max(MinWidth, Math.Min(MaxWidth, w));
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            // 自绘，避免闪烁
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            int visible = Math.Min(_rows, Math.Max(1, _items.Count - _top));
            int listH = visible * _rowH;

            using (var b = new SolidBrush(CBg)) g.FillRectangle(b, 0, 0, Width, listH);
            using (var b = new SolidBrush(CDetailBg)) g.FillRectangle(b, 0, listH, Width, Height - listH);

            for (int i = 0; i < visible; i++)
            {
                int idx = _top + i;
                if (idx >= _items.Count) break;
                var c = _items[idx];
                int y = i * _rowH;
                bool on = idx == _sel;

                if (on)
                {
                    using (var b = new SolidBrush(CSelBg)) g.FillRectangle(b, 0, y, Width, _rowH);
                }
                else if (i % 2 == 1)
                {
                    // 极淡的隔行底色，长列表更好扫
                    using (var b = new SolidBrush(Color.FromArgb(250, 250, 250)))
                        g.FillRectangle(b, 0, y, Width, _rowH);
                }

                Color fg = on ? CSelFg : CFg;
                Color hit = on ? CSelHit : CHit;

                int ix = PadX;
                int iy = y + (_rowH - _iconW) / 2;
                DrawIcon(g, c.Kind, ix, iy, _iconW, on);

                int tx = ix + _iconW + IconGap;
                DrawName(g, c, tx, y, fg, hit);
            }

            // 详情行：选中项的完整名字（长名字被列表宽度截断时看这里）
            if (_sel >= 0 && _sel < _items.Count)
            {
                using (var pen = new Pen(CDetailSep)) g.DrawLine(pen, 0, listH, Width, listH);
                using (var b = new SolidBrush(CDetailFg))
                {
                    var rf = new RectangleF(PadX, listH + (Height - listH - _font.Height) / 2f,
                        Width - PadX * 2 - ScrollW, _font.Height + 2);
                    using (var sf = new StringFormat { Trimming = StringTrimming.EllipsisCharacter })
                        g.DrawString(_items[_sel].Name, _font, b, rf, sf);
                }
            }

            // 滚动条（只有真需要时才画）
            if (_items.Count > _rows)
            {
                int trackX = Width - ScrollW;
                using (var b = new SolidBrush(CTrack)) g.FillRectangle(b, trackX, 0, ScrollW, listH);
                int total = _items.Count;
                int thumbH = Math.Max(12, (int)((long)listH * Math.Min(_rows, total) / total));
                int maxTop = Math.Max(1, total - _rows);
                int thumbY = (int)((long)(listH - thumbH) * Math.Min(_top, maxTop) / maxTop);
                using (var b = new SolidBrush(CThumb))
                    g.FillRectangle(b, trackX + 1, thumbY, ScrollW - 2, thumbH);
            }

            using (var pen = new Pen(CBorder)) g.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
        }

        private void DrawName(Graphics g, Candidate c, int x, int y, Color fg, Color hit)
        {
            string name = c.Name ?? "";
            int baselineY = y + ( _rowH - _font.Height) / 2;

            // 命中字符单独上色：前缀命中的话就是开头那几个字
            if (c.Pos >= 0 && c.Span > 0 && c.Pos + c.Span <= name.Length && c.Span < name.Length)
            {
                string pre = name.Substring(0, c.Pos);
                string mid = name.Substring(c.Pos, c.Span);
                string post = name.Substring(c.Pos + c.Span);

                int cx = x;
                cx = DrawChunk(g, pre, cx, baselineY, fg);
                cx = DrawChunk(g, mid, cx, baselineY, hit);
                DrawChunk(g, post, cx, baselineY, fg);
            }
            else
            {
                TextRenderer.DrawText(g, name, _font, new Point(x, baselineY), fg,
                    TextFormatFlags.NoPadding);
            }
        }

        private int DrawChunk(Graphics g, string s, int x, int y, Color color)
        {
            if (string.IsNullOrEmpty(s)) return x;
            TextRenderer.DrawText(g, s, _font, new Point(x, y), color, TextFormatFlags.NoPadding);
            return x + TextRenderer.MeasureText(g, s, _font, new Size(int.MaxValue, _rowH),
                TextFormatFlags.NoPadding).Width;
        }

        /// <summary>
        /// 五档图标全部现画（不发位图）：函数=绿箭头、属性=紫方块、变量=蓝竖条、
        /// 类型=琥珀菱形、关键字=灰横线。**拿不到档位就不画**。
        /// </summary>
        private static void DrawIcon(Graphics g, string kind, int x, int y, int w, bool on)
        {
            if (string.IsNullOrEmpty(kind)) return;
            Color color;
            if (!(on ? IconColorOn : IconColor).TryGetValue(kind, out color)) return;

            using (var b = new SolidBrush(color))
            using (var pen = new Pen(color))
            {
                switch (kind)
                {
                    case "func":   // 绿箭头（VBE 的方法图标就是一支箭头）
                        {
                            var pts = new[]
                            {
                                new Point(x, y + w / 4),
                                new Point(x + w / 2, y + w / 4),
                                new Point(x + w / 2, y),
                                new Point(x + w, y + w / 2),
                                new Point(x + w / 2, y + w),
                                new Point(x + w / 2, y + w * 3 / 4),
                                new Point(x, y + w * 3 / 4),
                            };
                            g.FillPolygon(b, pts);
                            break;
                        }
                    case "prop":   // 紫色实心方块
                        g.FillRectangle(b, x + 1, y + w / 4, w - 2, w / 2);
                        break;
                    case "type":   // 琥珀菱形
                        {
                            var pts = new[]
                            {
                                new Point(x + w / 2, y),
                                new Point(x + w, y + w / 2),
                                new Point(x + w / 2, y + w),
                                new Point(x, y + w / 2),
                            };
                            g.FillPolygon(b, pts);
                            break;
                        }
                    case "kw":     // 灰色横线
                        g.FillRectangle(b, x + 1, y + w / 2 - 1, w - 2, 2);
                        break;
                    default:       // var：蓝色小竖条
                        g.FillRectangle(b, x + w / 3, y + 1, Math.Max(2, w / 4), w - 2);
                        break;
                }
            }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            int i = e.Y / Math.Max(1, _rowH);
            if (i >= 0 && _top + i < _items.Count) RowClicked?.Invoke(_top + i);
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            Wheel?.Invoke(e.Delta > 0 ? -3 : 3);
        }

        protected override void WndProc(ref System.Windows.Forms.Message m)
        {
            const int WM_MOUSEACTIVATE = 0x0021;
            const int MA_NOACTIVATE = 3;
            if (m.Msg == WM_MOUSEACTIVATE)
            {
                m.Result = (IntPtr)MA_NOACTIVATE;
                return;
            }
            base.WndProc(ref m);
        }
    }
}
