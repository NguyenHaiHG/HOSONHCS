using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HOSONHCS
{
    internal static class KiemTraTrungGqvl
    {
        /// <summary>
        /// Chỉ cho tạo mới khi CCCD và SĐT đều khác toàn bộ hồ sơ hiện có.
        /// </summary>
        public static bool KiemTraKhachMoi(
            KhachHangGqvl item,
            IEnumerable<KhachHangGqvl> existingCustomers,
            int skipIndex,
            out string message)
        {
            message = null;
            if (item == null) return true;

            string cccd = NormalizeCccd(item.Cccd);
            string sdt = NormalizePhone(item.SdtKh);

            if (string.IsNullOrWhiteSpace(cccd) && string.IsNullOrWhiteSpace(sdt))
                return true;

            var list = existingCustomers?.ToList() ?? new List<KhachHangGqvl>();
            KhachHangGqvl trungCccd = null;
            KhachHangGqvl trungSdt = null;

            for (int i = 0; i < list.Count; i++)
            {
                if (i == skipIndex) continue;

                var other = list[i];
                if (other == null) continue;

                if (trungCccd == null
                    && !string.IsNullOrWhiteSpace(cccd)
                    && string.Equals(NormalizeCccd(other.Cccd), cccd, StringComparison.OrdinalIgnoreCase))
                {
                    trungCccd = other;
                }

                if (trungSdt == null
                    && !string.IsNullOrWhiteSpace(sdt)
                    && string.Equals(NormalizePhone(other.SdtKh), sdt, StringComparison.Ordinal))
                {
                    trungSdt = other;
                }

                if (trungCccd != null && trungSdt != null)
                    break;
            }

            if (trungCccd == null && trungSdt == null)
                return true;

            var sb = new StringBuilder();
            sb.AppendLine("Không thể tạo hợp đồng mới. CCCD và SĐT phải khác toàn bộ hồ sơ đã có:");

            if (trungCccd != null)
            {
                sb.AppendLine();
                sb.AppendLine("• CCCD \"" + cccd + "\" đã tồn tại:");
                sb.AppendLine("  - Tên: " + (trungCccd.Tenkh ?? ""));
                sb.AppendLine("  - Số HĐ: " + (trungCccd.Sohd ?? ""));
                sb.AppendLine("  - SĐT: " + (trungCccd.SdtKh ?? ""));
            }

            if (trungSdt != null)
            {
                sb.AppendLine();
                sb.AppendLine("• SĐT \"" + sdt + "\" đã tồn tại:");
                sb.AppendLine("  - Tên: " + (trungSdt.Tenkh ?? ""));
                sb.AppendLine("  - Số HĐ: " + (trungSdt.Sohd ?? ""));
                sb.AppendLine("  - CCCD: " + (trungSdt.Cccd ?? ""));
            }

            sb.AppendLine();
            sb.AppendLine("Vui lòng dùng Cập nhật khách hàng hoặc nhập CCCD/SĐT khác.");
            message = sb.ToString();
            return false;
        }

        private static string NormalizeCccd(string value)
        {
            return (value ?? string.Empty).Trim();
        }

        private static string NormalizePhone(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;

            var digits = new string(value.Where(char.IsDigit).ToArray());
            return string.IsNullOrWhiteSpace(digits) ? value.Trim() : digits;
        }
    }
}
