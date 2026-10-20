// 纯判据：不碰 COM、不碰 UI、不碰全局状态 —— 全部可以单独测。
//
// ★为什么这些要独立成一个文件：它们是"补全准不准"的全部决定者。
//   混进 COM 或窗口代码里的话，就只能靠真机试，一条 bug 要试十次。

using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace VbeComplete
{
    internal static class Rules
    {
        public static bool IsIdentStart(char c) => char.IsLetter(c) || c == '_';
        public static bool IsIdentChar(char c) => char.IsLetterOrDigit(c) || c == '_';

        /// <summary>
        /// 取光标左边那个词。caretCol 是 1-based 字符列（VBE 的 GetSelection 给的）。
        /// 返回 (词, 词的起始列 1-based)；取不到返回 ("", caretCol)。
        /// </summary>
        public static string WordBefore(string line, int caretCol, out int wordStartCol)
        {
            wordStartCol = caretCol;
            if (string.IsNullOrEmpty(line)) return "";
            int end = Math.Min(caretCol - 1, line.Length);   // 0-based 排他
            if (end <= 0) return "";
            int i = end - 1;
            if (!IsIdentChar(line[i])) return "";
            int j = i;
            while (j >= 0 && IsIdentChar(line[j])) j--;
            wordStartCol = j + 2;                            // 1-based
            return line.Substring(j + 1, i - j);
        }

        /// <summary>把 [wordStartCol, endCol) 换成 completion（列都是 1-based）。</summary>
        public static string ReplaceWord(string line, int wordStartCol, int endCol, string completion)
        {
            line = line ?? "";
            int start = Math.Max(0, wordStartCol - 1);
            int end = Math.Max(start, Math.Min(line.Length, endCol - 1));
            if (start > line.Length) start = line.Length;
            return line.Substring(0, start) + completion + line.Substring(end);
        }

        /// <summary>
        /// VBE 重排过整行之后，找回补全词真正的结束下标（0-based 排他）。
        /// 只认【整词】出现（前后都不是标识符字符），多处出现取起点最接近 near0 的；
        /// 找不到就用 fallback0。
        /// ★v131b：原来用 LastIndexOf 且不校验词边界 —— 行里出现两次同名
        ///   （`Split = Split(x)`）时会取到最后一个，光标就跑到别的词后面去。
        ///   判据与 Python 版 `_find_completion_end` 同一份语义。
        /// </summary>
        public static int FindCompletionEnd(string actual, string completion, int near0, int fallback0)
        {
            int clamped = Math.Max(0, Math.Min(actual == null ? 0 : actual.Length, fallback0));
            if (string.IsNullOrEmpty(actual) || string.IsNullOrEmpty(completion)) return clamped;
            int L = completion.Length;
            int best = -1;
            int i = 0;
            while (i <= actual.Length - L)
            {
                int hit = actual.IndexOf(completion, i, StringComparison.OrdinalIgnoreCase);
                if (hit < 0) break;
                bool okL = hit == 0 || !IsIdentChar(actual[hit - 1]);
                bool okR = hit + L >= actual.Length || !IsIdentChar(actual[hit + L]);
                if (okL && okR && (best < 0 || Math.Abs(hit - near0) < Math.Abs(best - near0))) best = hit;
                i = hit + L;
            }
            if (best >= 0) return best + L;
            return clamped;
        }

        /// <summary>
        /// 字符下标 -> VBE 认的"显示列"（1-based）。
        /// VBE 的 SetSelection 列是按 Tab 展开后的格数算的：不做换算，光标
        /// 在有 Tab 或缺省缩进（VBE 用 Tab）的行上会落进词的中间。
        /// </summary>
        public static int DispCol(string text, int charIndex, int tabw = 4)
        {
            if (string.IsNullOrEmpty(text)) return 1;
            if (charIndex > text.Length) charIndex = text.Length;
            if (charIndex < 0) charIndex = 0;
            int col = 0;
            for (int i = 0; i < charIndex; i++)
            {
                char c = text[i];
                if (c == '\t') col += tabw - (col % tabw);
                else if (IsWide(c)) col += 2;
                else col += 1;
            }
            return col + 1;
        }

        public static bool IsWide(char c) =>
            (c >= 0x1100 && c <= 0x115F) || (c >= 0x2E80 && c <= 0xA4CF) ||
            (c >= 0xAC00 && c <= 0xD7A3) || (c >= 0xF900 && c <= 0xFAFF) ||
            (c >= 0xFE30 && c <= 0xFE6F) || (c >= 0xFF00 && c <= 0xFF60) ||
            (c >= 0xFFE0 && c <= 0xFFE6);

        /// <summary>
        /// 反向换算：VBE 给的"显示列"（1-based）-> 1-based 字符列。
        /// ★v131b：`CodePane.GetSelection` 返回的列是**显示列**（真机实测：
        ///   `    学生 = Split()` 里 `(` 前是 17 列，按字符数是 15）。把它直接
        ///   当字符列去取词，含中文/全角的行会取错位置甚至取到空词 ——
        ///   症状是"这一行根本不弹候选"。凡是吃 `snap.CaretCol` 的判据，
        ///   都必须先过这一道。
        /// </summary>
        public static int CharColFromDisp(string text, int dispCol, int tabw = 4)
        {
            if (string.IsNullOrEmpty(text) || dispCol <= 1) return 1;
            int col = 0;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                int w = c == '\t' ? tabw - (col % tabw) : (IsWide(c) ? 2 : 1);
                if (col + w >= dispCol) return i + 1;
                col += w;
            }
            return text.Length + 1;
        }

        // ------------------------------------------------------------------
        // 位置判据：这个位置该不该弹
        // ------------------------------------------------------------------

        private static readonly Regex ReDeclHead = new Regex(
            @"^[ \t]*(?:(?:Public|Private|Friend|Global|Static)[ \t]+)*" +
            @"(?:Dim|Const|ReDim|WithEvents|Sub|Function|Property|Declare|Event|Enum|Type)\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ReAsTail = new Regex(
            @"\bAs[ \t]+(?:New[ \t]+)?[\w\.]*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>光标是不是停在"造一个新名字"的位置（`Dim cus|` / `Sub Do|`）。</summary>
        public static bool InDeclarationPosition(string line, int caretCol)
        {
            if (string.IsNullOrEmpty(line)) return false;
            int end = Math.Max(0, Math.Min(caretCol - 1, line.Length));
            string before = line.Substring(0, end);
            if (!ReDeclHead.IsMatch(line)) return false;

            // `Dim x As Inte|` —— 类型位置，要提示类型名，不算"造名字"
            if (ReAsTail.IsMatch(before)) return false;

            // `Dim a As Long, b As Str|` —— 最后一个逗号之后含 As 就是类型位置
            int comma = before.LastIndexOf(',');
            if (comma >= 0)
            {
                string tail = before.Substring(comma + 1);
                if (Regex.IsMatch(tail, @"\bAs\b", RegexOptions.IgnoreCase)) return false;
            }
            // `Const PI = 3.14` 的值位置不是名字位置
            if (Regex.IsMatch(before, @"\bConst\b.*=", RegexOptions.IgnoreCase)) return false;
            return true;
        }

        /// <summary>光标是不是紧跟在点号后面（`student.|`）—— 这个位置整个让给 VBE。</summary>
        public static bool InMemberPosition(string line, int caretCol)
        {
            if (string.IsNullOrEmpty(line)) return false;
            int end = Math.Max(0, Math.Min(caretCol - 1, line.Length));
            string before = line.Substring(0, end).TrimEnd();
            return before.EndsWith(".", StringComparison.Ordinal);
        }

        /// <summary>光标是不是在注释里（`'` 之后）或字符串里。</summary>
        public static bool InCommentOrString(string line, int caretCol)
        {
            if (string.IsNullOrEmpty(line)) return false;
            int end = Math.Max(0, Math.Min(caretCol - 1, line.Length));
            bool inStr = false;
            for (int i = 0; i < end; i++)
            {
                char c = line[i];
                if (inStr)
                {
                    if (c == '"')
                    {
                        if (i + 1 < end && line[i + 1] == '"') { i++; continue; }
                        inStr = false;
                    }
                    continue;
                }
                if (c == '"') { inStr = true; continue; }
                if (c == '\'') return true;
            }
            return inStr;
        }

        /// <summary>
        /// 这个词的位置允不允许补一对括号。
        /// 只管"位置"，不管"是什么名字"（那是引擎的活）。
        /// </summary>
        public static bool IsCallSite(string line, int wordStartCol)
        {
            if (string.IsNullOrEmpty(line)) return false;
            int start = Math.Max(0, Math.Min(wordStartCol - 1, line.Length));
            string before = line.Substring(0, start).TrimEnd();
            if (before.Length == 0) return false;
            if (before.EndsWith(".", StringComparison.Ordinal)) return false;
            if (Regex.IsMatch(before, @"\b(?:Sub|Function|Property|Declare|Enum|Type|Event|As|New|On|GoTo|GoSub|Resume|Call)\s*$",
                    RegexOptions.IgnoreCase))
                return false;
            return true;
        }

        // ------------------------------------------------------------------
        // 匹配
        // ------------------------------------------------------------------

        private static bool IsBoundary(string name, int i)
        {
            if (i <= 0) return true;
            char prev = name[i - 1], cur = name[i];
            if (prev == '_') return true;
            return char.IsLower(prev) && char.IsUpper(cur);
        }

        /// <summary>
        /// 模糊匹配：先认前缀，再认"按词边界跳步"（`sb` 能匹配 `selBegin`）。
        /// 返回 (命中位置, 跨度)；不命中返回 false。
        /// ★跳步命中有个硬性要求：query 的第一个字符必须落在词边界上，
        ///   否则 `e` 会匹配到任意单词内部，噪音把真正的前缀命中挤掉。
        /// </summary>
        public static bool Match(string name, string query, out int pos, out int span)
        {
            pos = 0; span = 0;
            if (string.IsNullOrEmpty(name) || query == null) return false;
            if (query.Length == 0) { pos = 0; span = 0; return true; }

            if (name.StartsWith(query, StringComparison.OrdinalIgnoreCase))
            {
                pos = 0; span = query.Length;
                return true;
            }
            if (query.Length > name.Length) return false;

            // 跳步：在边界上贪心找
            int first = -1, last = -1;
            int qi = 0;
            for (int i = 0; i < name.Length && qi < query.Length; i++)
            {
                if (char.ToLowerInvariant(name[i]) != char.ToLowerInvariant(query[qi])) continue;
                if (qi == 0 && !IsBoundary(name, i)) continue;
                if (first < 0) first = i;
                last = i;
                qi++;
            }
            if (qi < query.Length || first < 0) return false;
            pos = first;
            span = last - first + 1;
            return true;
        }

        /// <summary>
        /// "打全了没有"：精确命中【并且】没有更长的候选，才算打全。
        /// 打全就该收起列表 —— 否则按 Tab 会把已经打完的名字再补一遍。
        /// </summary>
        public static bool PrefixIsResolved(string word, IList<string> matches)
        {
            if (string.IsNullOrEmpty(word) || matches == null || matches.Count == 0) return false;
            bool exact = false;
            foreach (var m in matches)
            {
                if (string.Equals(m, word, StringComparison.OrdinalIgnoreCase)) { exact = true; continue; }
                if (m.Length > word.Length &&
                    m.StartsWith(word, StringComparison.OrdinalIgnoreCase))
                    return false;    // 还有更长的候选，没打全
            }
            return exact;
        }

        /// <summary>
        /// 一次文本变化是不是"在光标处插进了一个字符"（用来判断用户真敲了字）。
        /// 命中时 inserted 是插入的那个字符；删除/替换返回 false。
        /// </summary>
        public static bool LineEditInserted(string oldLine, string newLine, out char inserted)
        {
            inserted = '\0';
            oldLine = oldLine ?? "";
            newLine = newLine ?? "";
            if (newLine.Length != oldLine.Length + 1) return false;
            int i = 0;
            while (i < oldLine.Length && oldLine[i] == newLine[i]) i++;
            // 剩下的部分必须逐字相同（中间插了一个字符）
            if (i + 1 > newLine.Length) return false;
            if (!string.Equals(oldLine.Substring(i), newLine.Substring(i + 1), StringComparison.Ordinal))
                return false;
            inserted = newLine[i];
            return true;
        }

        /// <summary>把行尾的空白去掉（比较"有没有实质变化"时用）。</summary>
        public static string Rstrip(string s) => (s ?? "").TrimEnd(' ', '\t', '\r', '\n');
    }
}
