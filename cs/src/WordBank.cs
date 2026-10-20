// 词库：C# 版不再内嵌 SQLite（那会给一个纯工具塞进一个本地数据库驱动）。
//
// 数据来源两处，都只是 JSON：
//   * wordbank.json —— 由 `py -3 wordbank.py export --out wordbank.json` 导出；
//   * 内置种子表 —— BuiltinTables（171 个函数 + 13 个类型 + 79 个关键字）。
// 两份都读不到时，行为与"只有内置种子"完全一致，绝不因此让工具变哑。

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace VbeComplete
{
    internal sealed class WordBank
    {
        private sealed class Entry
        {
            public string Name;
            public string Kind;
            public bool Enabled = true;
            public bool? WantsSpace;
        }

        private readonly Dictionary<string, Entry> _map =
            new Dictionary<string, Entry>(StringComparer.OrdinalIgnoreCase);
        private readonly bool _disabled;
        public string Path_ { get; }

        public WordBank()
        {
            _disabled = Settings.Flag("NO_WORDBANK", false);
            Path_ = System.IO.Path.Combine(Settings.Home, "wordbank.json");
            if (!_disabled) Load();
        }

        public int Count => _map.Count;

        private void Load()
        {
            try
            {
                if (!File.Exists(Path_)) return;
                using (var doc = JsonDocument.Parse(File.ReadAllText(Path_)))
                {
                    JsonElement arr;
                    if (doc.RootElement.ValueKind == JsonValueKind.Object)
                    {
                        if (!doc.RootElement.TryGetProperty("words", out arr)) return;
                    }
                    else arr = doc.RootElement;
                    if (arr.ValueKind != JsonValueKind.Array) return;

                    foreach (var it in arr.EnumerateArray())
                    {
                        if (it.ValueKind != JsonValueKind.Object) continue;
                        var e = new Entry();
                        JsonElement v;
                        if (it.TryGetProperty("name", out v)) e.Name = v.GetString();
                        if (string.IsNullOrEmpty(e.Name)) continue;
                        if (it.TryGetProperty("kind", out v) && v.ValueKind == JsonValueKind.String)
                            e.Kind = v.GetString();
                        if (it.TryGetProperty("enabled", out v))
                            e.Enabled = v.ValueKind != JsonValueKind.False;
                        if (it.TryGetProperty("wants_space", out v) &&
                            (v.ValueKind == JsonValueKind.True || v.ValueKind == JsonValueKind.False))
                            e.WantsSpace = v.GetBoolean();
                        _map[e.Name] = e;
                    }
                }
            }
            catch (Exception ex) { Log.Write("wordbank.json 读取失败：{0}", ex.Message); }
        }

        /// <summary>词库里额外收的名字（内置表里没有的那些）。</summary>
        public IEnumerable<string> ExtraNames()
        {
            foreach (var kv in _map)
            {
                var e = kv.Value;
                if (!e.Enabled) continue;
                if (VbaBuiltins.IsKeyword(e.Name) || VbaBuiltins.IsBuiltinFunction(e.Name) ||
                    VbaBuiltins.IsBuiltinType(e.Name)) continue;
                yield return e.Name;
            }
        }

        /// <summary>词库逐条覆盖"要不要补空格"；没有覆盖就返回 null（用内置判据）。</summary>
        public bool? WantsSpace(string name)
        {
            Entry e;
            if (_map.TryGetValue(name ?? "", out e) && e.Enabled) return e.WantsSpace;
            return null;
        }

        /// <summary>词库说这个是过程/函数/类型吗（用于图标档位）。</summary>
        public string KindOf(string name)
        {
            Entry e;
            if (_map.TryGetValue(name ?? "", out e) && e.Enabled) return e.Kind;
            return null;
        }

        public bool Known(string name) => _map.ContainsKey(name ?? "");
    }
}
