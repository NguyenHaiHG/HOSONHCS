using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace HOSONHCS
{
    /// <summary>
    /// Helper tinh - doc tu toanquoc.json (mang nhieu tinh).
    /// Ten tinh giu nguyen nhu trong file, vi du "Tinh Tuyen Quang".
    /// </summary>
    internal static class TinhHelper
    {
        private const string TOAN_QUOC_FILE = "toanquoc.json";

        private static readonly HashSet<string> _excluded = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "holiday_notice.json",
            "Form2State.json",
            "toanquoc.json",
            "tuyenquang.json",
            "quangngai.json"
        };

        private static Dictionary<string, TinhModel> _cache = null;
        private static readonly object _lock = new object();

        public static void RefreshCache()
        {
            lock (_lock) { _cache = null; }
        }

        private static bool IsNan(string s)
            => !string.IsNullOrWhiteSpace(s) && s.Equals("nan", StringComparison.OrdinalIgnoreCase);

        private static List<Commune> FilterNan(List<Commune> communes)
        {
            if (communes == null) return new List<Commune>();
            foreach (var c in communes)
            {
                // Chỉ xóa groups = "nan", GIỮ NGUYÊN assoc "nan" để code UI còn lấy villages
                if (c.associations != null)
                    foreach (var a in c.associations)
                        if (a.villages != null)
                            foreach (var v in a.villages)
                                v.groups?.RemoveAll(g => IsNan(g));
                if (c.villages != null)
                    foreach (var v in c.villages)
                        v.groups?.RemoveAll(g => IsNan(g));
            }
            return communes;
        }

        private static Dictionary<string, TinhModel> BuildCache()
        {
            var result = new Dictionary<string, TinhModel>(StringComparer.OrdinalIgnoreCase);
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;

            // 1. Doc toanquoc.json
            var jsonPath = Path.Combine(baseDir, TOAN_QUOC_FILE);
            var json = File.Exists(jsonPath) ? File.ReadAllText(jsonPath, Encoding.UTF8) : null;
            if (!string.IsNullOrEmpty(json))
            {
                try
                {
                    var arr = JArray.Parse(json);
                    foreach (JObject item in arr)
                    {
                        try
                        {
                            var tinhName = item["tinh"]?.ToString()?.Trim() ?? "";
                            if (string.IsNullOrWhiteSpace(tinhName)) continue;

                            var pgdsToken = item["pgds"] as JArray;
                            if (pgdsToken == null) continue;

                            var pgds = new List<TinhPgdEntry>();
                            foreach (JObject pgdObj in pgdsToken)
                            {
                                var pgdName = pgdObj["pgd"]?.ToString() ?? "";
                                if (string.IsNullOrWhiteSpace(pgdName) || pgdName.Equals("nan", StringComparison.OrdinalIgnoreCase)) continue;
                                var communesToken = pgdObj["communes"] as JArray;
                                var communes = communesToken != null
                                    ? FilterNan(JsonConvert.DeserializeObject<List<Commune>>(communesToken.ToString()))
                                    : new List<Commune>();
                                pgds.Add(new TinhPgdEntry { pgd = pgdName, communes = communes ?? new List<Commune>() });
                            }

                            result[tinhName] = new TinhModel { tinh = tinhName, pgds = pgds };
                        }
                        catch { }
                    }
                }
                catch { }
            }

            // 2. Doc them file don tinh (neu co, khong ghi de)
            foreach (var file in Directory.GetFiles(baseDir, "*.json", SearchOption.TopDirectoryOnly))
            {
                if (_excluded.Contains(Path.GetFileName(file))) continue;
                try
                {
                    var fileJson = File.ReadAllText(file, Encoding.UTF8);
                    var model = JsonConvert.DeserializeObject<TinhModel>(fileJson);
                    if (model == null || string.IsNullOrWhiteSpace(model.tinh) || model.pgds == null) continue;
                    var tinhName = model.tinh.Trim();
                    if (!result.ContainsKey(tinhName))
                        result[tinhName] = new TinhModel { tinh = tinhName, pgds = model.pgds };
                }
                catch { }
            }

            return result;
        }

        private static Dictionary<string, TinhModel> GetCache()
        {
            lock (_lock)
            {
                if (_cache == null) _cache = BuildCache();
                return _cache;
            }
        }

        public static List<string> GetProvinceNames()
            => GetCache().Keys.OrderBy(k => k).ToList();

        public static TinhModel LoadTinhModel(string tinhName)
        {
            if (string.IsNullOrWhiteSpace(tinhName)) return null;
            var cache = GetCache();
            return cache.TryGetValue(tinhName.Trim(), out var model) ? model : null;
        }

        public static string GetFileName(string tinhName)
        {
            return LoadTinhModel(tinhName) != null ? TOAN_QUOC_FILE : null;
        }

        public static void PopulateComboBox(System.Windows.Forms.ComboBox cb)
        {
            if (cb == null) return;
            var current = cb.Text;
            cb.Items.Clear();
            foreach (var name in GetProvinceNames())
                cb.Items.Add(name);
            if (!string.IsNullOrEmpty(current) && cb.Items.Contains(current))
                cb.Text = current;
        }
    }
}