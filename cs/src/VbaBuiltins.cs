// 内建名字：语言关键字 / 内建函数 / 内建类型。
//
// ★默认口径（与 python 版一致，都是用户在真机上一条条要求出来的）：
//     * 语言关键字      —— 收（`ENABLE_VBA_KEYWORDS`，默认 True）
//     * 内建函数 / 类型  —— 收
//     * 一切枚举常量     —— 【不收】（`VBA_CONSTANTS` / `HOST_ENUMS` 默认 False）
//       用户原话："内置枚举（vb*）不要提示，提示出来只是干扰。"

using System;
using System.Collections.Generic;

namespace VbeComplete
{
    internal static class VbaBuiltins
    {
        public static readonly bool EnableKeywords = Settings.Flag("VBA_KEYWORDS", true);
        public static readonly bool EnableConstants = Settings.Flag("VBA_CONSTANTS", false);
        public static readonly bool EnableHostEnums = Settings.Flag("HOST_ENUMS", false);

        private static readonly HashSet<string> _funcs = Mk(BuiltinTables.Functions);
        private static readonly HashSet<string> _types = Mk(BuiltinTables.TypeNames);
        private static readonly HashSet<string> _kw = Mk(BuiltinTables.Keywords);
        private static readonly HashSet<string> _kwSpace = Mk(BuiltinTables.KeywordsWantSpace);

        private static HashSet<string> Mk(string[] src)
        {
            var s = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var n in src) s.Add(n);
            return s;
        }

        public static bool IsBuiltinFunction(string name) =>
            !string.IsNullOrEmpty(name) && _funcs.Contains(name);

        public static bool IsBuiltinType(string name) =>
            !string.IsNullOrEmpty(name) && _types.Contains(name);

        public static bool IsKeyword(string name) =>
            !string.IsNullOrEmpty(name) && _kw.Contains(name);

        /// <summary>补全后要不要跟一个空格（只有语言关键字，且排除掉那些不该补的）。</summary>
        public static bool KeywordWantsSpace(string name) =>
            !string.IsNullOrEmpty(name) && _kwSpace.Contains(name);

        /// <summary>
        /// 是不是"该补一对括号"的可调用内建名。
        /// 两道判据问的是不同问题，不能合成一条：
        ///   ① 白名单（内建函数）—— 挡住关键字的 `If` / 常量；
        ///   ② 类型排除 —— 挡住 `String` / `Date` 这类"既是类型又是函数"的双身份名字
        ///      （在 `Dim x As String` 的位置补括号就是语法错误）。
        /// </summary>
        public static bool IsCallableBuiltin(string name) =>
            IsBuiltinFunction(name) && !IsBuiltinType(name);

        /// <summary>给引擎用的内建名总表（关键字是否收由开关决定）。</summary>
        public static IEnumerable<string> Pool()
        {
            foreach (var n in BuiltinTables.Functions) yield return n;
            foreach (var n in BuiltinTables.TypeNames) yield return n;
            if (EnableKeywords)
                foreach (var n in BuiltinTables.Keywords) yield return n;
        }
    }
}
