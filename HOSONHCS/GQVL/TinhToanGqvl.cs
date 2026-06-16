using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace HOSONHCS
{
    internal static class TinhToanGqvl
    {
        public static long ParseMoney(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return 0;
            string digits = new string(value.Where(char.IsDigit).ToArray());
            long result;
            return long.TryParse(digits, out result) ? result : 0;
        }

        public static int ParseFirstInt(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return 0;
            var match = Regex.Match(value, @"\d+");
            int result;
            return match.Success && int.TryParse(match.Value, out result) ? result : 0;
        }

        public static string FormatMoney(long value)
        {
            return string.Format(CultureInfo.InvariantCulture, "{0:N0}", value).Replace(",", ".");
        }

        public static string FormatDate(DateTime value)
        {
            return value == DateTime.MinValue ? string.Empty : value.ToString("dd/MM/yyyy");
        }

        public static string TinhLaiQuaHan(string laiSuat)
        {
            decimal value;
            if (!TryParseDecimal(laiSuat, out value)) return string.Empty;
            decimal result = value * 1.3m;
            return result.ToString("0.#####", CultureInfo.InvariantCulture).Replace(".", ",");
        }

        public static DateTime TinhHanTraNo(DateTime ngayVay, string thoiHanVay)
        {
            int months = ParseFirstInt(thoiHanVay);
            return months > 0 ? ngayVay.AddMonths(months) : DateTime.MinValue;
        }

        public static string TaoPhanKyBlock(DateTime ngayVay, string soTienText, string thoiHanText, string phanKyText)
        {
            long soTien = ParseMoney(soTienText);
            int thoiHan = ParseFirstInt(thoiHanText);
            int phanKy = ParseFirstInt(phanKyText);
            if (soTien <= 0 || thoiHan <= 0 || phanKy <= 0) return string.Empty;

            int soKy = (int)Math.Ceiling((double)thoiHan / phanKy);
            if (soKy <= 0) return string.Empty;

            long soTienMoiKy = RoundDownToHundredThousand(soTien / soKy);
            var amounts = new List<long>();
            long totalBeforeLast = 0;

            for (int i = 1; i <= soKy; i++)
            {
                if (i == soKy)
                {
                    amounts.Add(soTien - totalBeforeLast);
                }
                else
                {
                    amounts.Add(soTienMoiKy);
                    totalBeforeLast += soTienMoiKy;
                }
            }

            var sb = new StringBuilder();
            for (int i = 0; i < amounts.Count; i++)
            {
                int monthsToAdd = Math.Min((i + 1) * phanKy, thoiHan);
                DateTime date = ngayVay.AddMonths(monthsToAdd);
                string suffix = i == amounts.Count - 1 ? "." : ";";
                sb.Append("\t - Ngày ")
                  .Append(date.ToString("dd/MM/yyyy"))
                  .Append(", số tiền: ")
                  .Append(FormatMoney(amounts[i]))
                  .Append(" đồng")
                  .Append(suffix);

                if (i < amounts.Count - 1) sb.AppendLine();
            }

            return sb.ToString();
        }

        private static long RoundDownToHundredThousand(long value)
        {
            const long unit = 100000;
            if (value <= 0) return 0;
            return (value / unit) * unit;
        }

        private static bool TryParseDecimal(string value, out decimal result)
        {
            result = 0;
            if (string.IsNullOrWhiteSpace(value)) return false;
            string normalized = value.Trim().Replace("%", "").Replace(",", ".");
            return decimal.TryParse(normalized, NumberStyles.Any, CultureInfo.InvariantCulture, out result);
        }
    }
}
