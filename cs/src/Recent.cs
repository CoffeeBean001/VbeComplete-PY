// 最近使用：与 python 版共用同一份 recent.json（扁平字典 {名字: 递增序号}）。
// 序号越大表示越近；读坏了就当空，绝不抛异常。

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace VbeComplete
{
    internal sealed class Recent
    {
        private readonly Dictionary<string, int> _map =
            new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly string _path;
        private int _seq;
        private bool _dirty;
        private readonly bool _persist;

        public Recent()
        {
            _path = Path.Combine(Settings.Home, "recent.json");
            _persist = !Settings.Flag("NO_RECENT_PERSIST", false);
            Load();
        }

        public IReadOnlyDictionary<string, int> All => _map;

        private void Load()
        {
            try
            {
                if (!File.Exists(_path)) return;
                using (var doc = JsonDocument.Parse(File.ReadAllText(_path)))
                {
                    if (doc.RootElement.ValueKind != JsonValueKind.Object) return;
                    foreach (var p in doc.RootElement.EnumerateObject())
                    {
                        int v;
                        if (p.Value.ValueKind == JsonValueKind.Number && p.Value.TryGetInt32(out v))
                        {
                            _map[p.Name] = v;
                            if (v > _seq) _seq = v;
                        }
                    }
                }
            }
            catch (Exception e) { Log.Write("recent.json 读取失败：{0}", e.Message); }
        }

        public int Rank(string name)
        {
            int v;
            return _map.TryGetValue(name ?? "", out v) ? v : 0;
        }

        public void Note(string name)
        {
            if (string.IsNullOrEmpty(name)) return;
            _map[name] = ++_seq;
            _dirty = true;
        }

        public void Flush()
        {
            if (!_dirty || !_persist) return;
            try
            {
                string tmp = _path + ".tmp";
                var opts = new JsonSerializerOptions { WriteIndented = false };
                File.WriteAllText(tmp, JsonSerializer.Serialize(_map, opts));
                File.Move(tmp, _path, true);
                _dirty = false;
            }
            catch (Exception e) { Log.Write("recent.json 写入失败：{0}", e.Message); }
        }
    }
}
