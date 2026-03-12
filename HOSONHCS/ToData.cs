using System;

namespace HOSONHCS
{
    /// <summary>
    /// Class lưu trữ thông tin tổ (tương tự Customer cho Form1)
    /// Dùng để serialize/deserialize vào file JSON
    /// </summary>
    public class ToData
    {
        // Thông tin chung
        public string Pgd { get; set; }
        public string Xa { get; set; }
        public string Thon { get; set; }
        public string Totruong { get; set; }
        public string Chuongtrinh { get; set; }
        public DateTime NgayXuat { get; set; }
        public int SoThanhVien { get; set; }

        // Thông tin 5 tổ viên - Họ tên
        public string Kh1 { get; set; }
        public string Kh2 { get; set; }
        public string Kh3 { get; set; }
        public string Kh4 { get; set; }
        public string Kh5 { get; set; }

        // Số tiền
        public string Tien1 { get; set; }
        public string Tien2 { get; set; }
        public string Tien3 { get; set; }
        public string Tien4 { get; set; }
        public string Tien5 { get; set; }

        // Phương án
        public string Md1 { get; set; }
        public string Md2 { get; set; }
        public string Md3 { get; set; }
        public string Md4 { get; set; }
        public string Md5 { get; set; }

        // Thời hạn vay
        public string Time1 { get; set; }
        public string Time2 { get; set; }
        public string Time3 { get; set; }
        public string Time4 { get; set; }
        public string Time5 { get; set; }

        // Đối tượng
        public string Dt1 { get; set; }
        public string Dt2 { get; set; }
        public string Dt3 { get; set; }
        public string Dt4 { get; set; }
        public string Dt5 { get; set; }

        // Tên file để lưu trữ
        public string _fileName { get; set; }

        public ToData()
        {
            NgayXuat = DateTime.Now;
        }
    }
}
