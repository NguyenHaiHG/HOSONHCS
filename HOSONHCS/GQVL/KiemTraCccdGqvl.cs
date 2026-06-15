using System;
using System.Linq;

namespace HOSONHCS
{
    internal static class KiemTraCccdGqvl
    {
        public static bool KiemTraCccd(string cccd, DateTime ngaySinh, out string error)
        {
            error = null;
            string digits = new string((cccd ?? string.Empty).Where(char.IsDigit).ToArray());

            if (digits.Length != 12)
            {
                error = "Số CCCD GQVL phải nhập đủ 12 số.";
                return false;
            }

            string namSinh2So = (ngaySinh.Year % 100).ToString("00");
            string cccdNamSinh = digits.Substring(4, 2);
            if (!string.Equals(cccdNamSinh, namSinh2So, StringComparison.Ordinal))
            {
                error = "Số thứ tự 5 và 6 của CCCD không trùng với 2 số cuối năm sinh. Yêu cầu xem lại.";
                return false;
            }

            return true;
        }

        public static string TinhNoiCap(DateTime ngayCap)
        {
            var mocCanCuoc = new DateTime(2024, 7, 1);
            return ngayCap.Date >= mocCanCuoc
                ? "Bộ Công an"
                : "Cục Cảnh sát quản lý hành chính về trật tự xã hội";
        }

        public static DateTime TinhNgayHetHan(DateTime ngaySinh, DateTime ngayCap, out string hienThi)
        {
            hienThi = string.Empty;
            int tuoiNgayCap = TinhTuoi(ngaySinh, ngayCap);

            if (tuoiNgayCap >= 60)
            {
                hienThi = "Không thời hạn";
                return DateTime.MaxValue.Date;
            }

            int mocTuoi = tuoiNgayCap < 23 ? 25 : (tuoiNgayCap < 38 ? 40 : 60);
            DateTime han = ngaySinh.AddYears(mocTuoi);

            // Theo thực tế CCCD/căn cước: nếu cấp trong vòng 2 năm trước mốc tuổi thì dùng mốc kế tiếp.
            if ((han - ngayCap.Date).TotalDays <= 730 && mocTuoi < 60)
            {
                mocTuoi = mocTuoi == 25 ? 40 : 60;
                han = ngaySinh.AddYears(mocTuoi);
            }

            hienThi = han.ToString("dd/MM/yyyy");
            return han;
        }

        private static int TinhTuoi(DateTime ngaySinh, DateTime ngay)
        {
            int age = ngay.Year - ngaySinh.Year;
            if (ngay.Date < ngaySinh.Date.AddYears(age)) age--;
            return age;
        }
    }
}
