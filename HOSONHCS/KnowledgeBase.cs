using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace HOSONHCS
{
    // ============================================================
    // MODEL: Một cặp Q&A trong cơ sở kiến thức chatbot
    // ============================================================
    public class KnowledgeItem
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N").Substring(0, 8);
        public string Category { get; set; }        // Danh mục: "Hộ nghèo", "SXKD", ...
        public string Question { get; set; }        // Câu hỏi gốc
        public string Keywords { get; set; }        // Từ khóa cách nhau bởi dấu phẩy: "lãi suất, hộ nghèo"
        public string Answer { get; set; }          // Câu trả lời (có thể nhiều dòng)
        public int Priority { get; set; } = 5;     // Độ ưu tiên 1-10 (10 = cao nhất)
        public bool IsActive { get; set; } = true;  // Bật/tắt không cần xóa
        public string CreatedBy { get; set; } = "Admin";
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        public DateTime UpdatedDate { get; set; } = DateTime.Now;

        // Lấy mảng từ khóa để matching
        public string[] GetKeywordArray()
        {
            if (string.IsNullOrWhiteSpace(Keywords))
                return new string[0];

            return Keywords
                .Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(k => k.Trim().ToLowerInvariant())
                .Where(k => !string.IsNullOrEmpty(k))
                .ToArray();
        }
    }

    // ============================================================
    // LOGIC: Lưu / Tải / Tìm kiếm Knowledge Base
    // ============================================================
    public static class KnowledgeBaseManager
    {
        // Thư mục lưu file JSON kiến thức
        private static string FolderPath =>
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ChatBot", "knowledge");

        // Tên file mặc định gộp tất cả Q&A
        private static string DefaultFile =>
            Path.Combine(FolderPath, "knowledge_base.json");

        // -------------------------------------------------------
        // LƯU: Ghi toàn bộ danh sách vào file JSON
        // -------------------------------------------------------
        public static void SaveAll(IList<KnowledgeItem> items)
        {
            EnsureFolder();
            var json = JsonConvert.SerializeObject(items, Formatting.Indented);
            File.WriteAllText(DefaultFile, json, Encoding.UTF8);
        }

        // -------------------------------------------------------
        // TẢI: Đọc toàn bộ danh sách từ file JSON
        // -------------------------------------------------------
        public static List<KnowledgeItem> LoadAll()
        {
            try
            {
                if (!File.Exists(DefaultFile))
                    return new List<KnowledgeItem>();

                var json = File.ReadAllText(DefaultFile, Encoding.UTF8);
                return JsonConvert.DeserializeObject<List<KnowledgeItem>>(json)
                       ?? new List<KnowledgeItem>();
            }
            catch
            {
                return new List<KnowledgeItem>();
            }
        }

        // -------------------------------------------------------
        // TÌM KIẾM: Trả về danh sách kết quả theo điểm khớp
        // -------------------------------------------------------
        public static List<SearchResult> Search(string userInput, IList<KnowledgeItem> allItems)
        {
            if (string.IsNullOrWhiteSpace(userInput) || allItems == null)
                return new List<SearchResult>();

            var input = userInput.Trim().ToLowerInvariant();
            var results = new List<SearchResult>();

            foreach (var item in allItems.Where(i => i.IsActive))
            {
                int score = 0;

                // 1. Khớp chính xác câu hỏi → điểm cao nhất
                if (string.Equals(item.Question?.Trim(), userInput.Trim(),
                    StringComparison.OrdinalIgnoreCase))
                    score += 100;

                // 2. Câu hỏi chứa input
                else if (item.Question != null &&
                         item.Question.ToLowerInvariant().Contains(input))
                    score += 70;

                // 3. Input chứa câu hỏi (người dùng hỏi dài hơn)
                else if (item.Question != null &&
                         input.Contains(item.Question.ToLowerInvariant()))
                    score += 60;

                // 4. Đếm số từ khóa khớp
                var keywords = item.GetKeywordArray();
                foreach (var kw in keywords)
                {
                    if (input.Contains(kw))
                        score += 15;
                }

                // 5. Áp dụng độ ưu tiên
                score += item.Priority;

                if (score > 0)
                    results.Add(new SearchResult { Item = item, Score = score });
            }

            // Sắp xếp theo điểm cao nhất
            return results.OrderByDescending(r => r.Score).ToList();
        }

        // -------------------------------------------------------
        // THỐNG KÊ
        // -------------------------------------------------------
        public static string GetStats(IList<KnowledgeItem> items)
        {
            if (items == null || items.Count == 0)
                return "Chưa có dữ liệu kiến thức.";

            var active = items.Count(i => i.IsActive);
            var categories = items.Select(i => i.Category).Distinct().Count();

            return $"Tổng Q&A: {items.Count} | Đang bật: {active} | Danh mục: {categories}";
        }

        private static void EnsureFolder()
        {
            if (!Directory.Exists(FolderPath))
                Directory.CreateDirectory(FolderPath);
        }
    }

    // ============================================================
    // KẾT QUẢ TÌM KIẾM
    // ============================================================
    public class SearchResult
    {
        public KnowledgeItem Item { get; set; }
        public int Score { get; set; }

        // Phần trăm khớp (max score tham chiếu = 115)
        public int MatchPercent => Math.Min(100, Score * 100 / 115);
    }
}
