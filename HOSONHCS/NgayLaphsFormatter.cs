using System;
using System.IO;

namespace HOSONHCS
{
    /// <summary>
    /// Xử lý logic format và giá trị placeholder {{ngaylaphs}} cho từng loại mẫu Word.
    /// Quy tắc khi dateLaphs bỏ tick (ngaylaphs == DateTime.MinValue):
    ///   - 01 HN, 01 SXKD, 01 GQVL, 03 DS, 03 DS GROUP, 01TGTV : "Ngày.....tháng.....năm....."
    ///   - BIA (không có người thừa kế) : "Ngày ..... Tháng ..... Năm ....."
    ///   - GUQ                          : "Ngày ...../...../........"
    ///   - Các mẫu còn lại              : "...../...../........"
    /// </summary>
    internal static class NgayLaphsFormatter
    {
        // ── Hằng số chuỗi placeholder ──────────────────────────────────────────
        public const string PlaceholderNgayThangNam = "Ngày.....tháng.....năm.....";
        public const string PlaceholderDots         = "...../...../........";
        public const string PlaceholderBia          = "Ngày ..... Tháng ..... Năm .....";
        public const string PlaceholderGUQ          = "Ngày ...../...../........";

        // ── Nhận dạng loại mẫu ─────────────────────────────────────────────────
        public static bool Is01HN(string docPath)
        {
            if (string.IsNullOrWhiteSpace(docPath)) return false;
            var name = Path.GetFileName(docPath) ?? "";
            return name.IndexOf("01", StringComparison.OrdinalIgnoreCase) >= 0
                && name.IndexOf("HN", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static bool Is01SXKD(string docPath)
        {
            if (string.IsNullOrWhiteSpace(docPath)) return false;
            var name = Path.GetFileName(docPath) ?? "";
            return name.IndexOf("01", StringComparison.OrdinalIgnoreCase) >= 0
                && name.IndexOf("SXKD", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static bool Is01GQVL(string docPath)
        {
            if (string.IsNullOrWhiteSpace(docPath)) return false;
            var name = Path.GetFileName(docPath) ?? "";
            return name.IndexOf("GQVL", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>
        /// Trả về true cho cả "03 DS.docx" và "03 DS GROUP.docx".
        /// </summary>
        public static bool Is03DS(string docPath)
        {
            if (string.IsNullOrWhiteSpace(docPath)) return false;
            var name = Path.GetFileName(docPath) ?? "";
            return (name.IndexOf("03", StringComparison.OrdinalIgnoreCase) >= 0
                    && name.IndexOf("DS", StringComparison.OrdinalIgnoreCase) >= 0)
                   || name.IndexOf("03 DS", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static bool Is01TGTV(string docPath)
        {
            if (string.IsNullOrWhiteSpace(docPath)) return false;
            var name = Path.GetFileName(docPath) ?? "";
            return name.IndexOf("TGTV", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static bool IsGUQ(string docPath)
        {
            if (string.IsNullOrWhiteSpace(docPath)) return false;
            var name = Path.GetFileName(docPath) ?? "";
            return name.IndexOf("GUQ", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static bool IsBia(string docPath)
        {
            if (string.IsNullOrWhiteSpace(docPath)) return false;
            var name = Path.GetFileName(docPath) ?? "";
            return name.IndexOf("BIA", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        // ── Logic chính ────────────────────────────────────────────────────────

        /// <summary>
        /// Trả về giá trị thay thế cho {{ngaylaphs}} dựa trên loại mẫu và giá trị ngày.
        /// </summary>
        /// <param name="docPath">Đường dẫn file Word đang xử lý.</param>
        /// <param name="ngaylaphs">Ngày lập hồ sơ; DateTime.MinValue khi bỏ tick.</param>
        /// <param name="nguoiThuaKe">Số người thừa kế, dùng cho mẫu BIA.</param>
        public static string GetNgaylaphsValue(string docPath, DateTime ngaylaphs, int nguoiThuaKe = 0)
        {
            // BIA không có người thừa kế → bỏ trống dù có ngày hay không
            if (IsBia(docPath) && nguoiThuaKe == 0)
                return PlaceholderBia;

            // Có ngày hợp lệ → luôn format "Ngày DD tháng MM năm YYYY"
            if (ngaylaphs != DateTime.MinValue)
                return FormatNgayThangNam(ngaylaphs);

            // Bỏ tick: áp dụng theo từng loại mẫu
            if (Is01HN(docPath) || Is01SXKD(docPath) || Is01GQVL(docPath)
                || Is03DS(docPath) || Is01TGTV(docPath) || IsBia(docPath))
                return PlaceholderNgayThangNam;

            if (IsGUQ(docPath))
                return PlaceholderGUQ;

            return PlaceholderDots;
        }

        /// <summary>
        /// Format DateTime thành "Ngày DD tháng MM năm YYYY".
        /// Ví dụ: 15/03/2024 → "Ngày 15 tháng 03 năm 2024"
        /// </summary>
        public static string FormatNgayThangNam(DateTime date)
        {
            if (date == DateTime.MinValue) return "";
            return $"Ngày {date.Day:D2} tháng {date.Month:D2} năm {date.Year}";
        }
    }
}
