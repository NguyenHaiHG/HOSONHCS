using System;
using System.Collections.Generic;
using System.Linq;

namespace HOSONHCS
{
    /// <summary>
    /// Class lưu trữ dữ liệu bảng kê tiền theo tổ trưởng
    /// </summary>
    public class BangKeData
    {
        /// <summary>
        /// Tên tổ trưởng
        /// </summary>
        public string Totruong { get; set; }

        /// <summary>
        /// Chi tiết mệnh giá và số lượng
        /// Key: Mệnh giá, Value: Số lượng
        /// </summary>
        public Dictionary<long, long> ChiTiet { get; set; }

        /// <summary>
        /// Tổng số tiền mặt
        /// </summary>
        public long TongTien { get; set; }

        /// <summary>
        /// Số tiền sổ sách
        /// </summary>
        public long SoTienSoSach { get; set; }

        /// <summary>
        /// Chênh lệch (Tiền mặt - Sổ sách)
        /// </summary>
        public long ChenhLech { get; set; }

        /// <summary>
        /// Ngày tạo bảng kê
        /// </summary>
        public DateTime NgayTao { get; set; }

        /// <summary>
        /// Tên file JSON để lưu trữ
        /// </summary>
        public string _fileName { get; set; }

        public BangKeData()
        {
            ChiTiet = new Dictionary<long, long>();
            NgayTao = DateTime.Now;
        }

        /// <summary>
        /// Tạo chuỗi mô tả chi tiết bảng kê
        /// </summary>
        public string GetChiTietText()
        {
            if (ChiTiet == null || ChiTiet.Count == 0)
                return "Chưa có dữ liệu";

            var lines = new List<string>();
            foreach (var item in ChiTiet.OrderByDescending(x => x.Key))
            {
                if (item.Value > 0)
                {
                    lines.Add($"{item.Key:N0} VNĐ × {item.Value:N0} = {item.Key * item.Value:N0} VNĐ");
                }
            }

            return string.Join("; ", lines);
        }
    }
}
