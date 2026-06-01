using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace HOSONHCS
{
    // ============================================================
    // FILE NÀY XỬ LÝ LOGIC TỰ ĐỘNG KHI CHỌN CHƯƠNG TRÌNH (cbctr) Ở FORM 2:
    //   1. Tự động cập nhật danh sách đối tượng (cbdt1~5) theo chương trình
    //   2. Khi chọn "Cấp nước sạch...": giới hạn cbmd1~5 chỉ còn 2 lựa chọn
    //      và khóa không cho nhập tự do
    //
    // ĐỂ THÊM/SỬA CHƯƠNG TRÌNH:
    //   - Đối tượng → sửa hàm LayDanhSachDoiTuong()
    //   - Mục đích nước sạch → sửa hàm ApDungCapNuocSach_MucDich()
    //   - Để thêm rule mới → thêm else if vào LayDanhSachDoiTuong()
    // ============================================================

    public partial class Form2 : Form
    {
        // ── Danh sách phương án mục đích dùng cho cbmd khi chọn Cấp nước sạch ──
        private static readonly string[] MucDichCapNuocSach = new string[]
        {
            "Nâng cấp, sửa chữa CTNS, CTVS",
            "Xây mới CTNS, CTVS"
        };

        // ── Danh sách phương án mục đích thông thường (không bao gồm 2 item nước sạch) ──
        private static readonly string[] MucDichThongThuong = new string[]
        {
            "Mua trâu sinh sản",
            "Nuôi trâu sinh sản",
            "Mua bò sinh sản",
            "Nuôi bò sinh sản",
            "Mua dê sinh sản",
            "Nuôi dê sinh sản",
            "Nuôi lợn sinh sản",
            "Nuôi lợn",
            "Trồng cây quế",
            "Trồng cây keo",
            "Trồng cây mỡ",
            "Trồng cây cam",
            "Mở rộng cửa hàng tạp hoá",
            "Mở rộng cửa hàng ăn uống",
            "Mở rộng cửa hàng bán quần áo",
            "Trồng và chăm sóc cây cà phê",
            "Trồng và chăm sóc cây cao su",
            "Trồng cây ăn quả",
            "Trồng cây bời lời",
            "Trồng cây tiêu"
        };

        // ── Hàm chính: gọi khi cbctr.SelectedIndexChanged ──
        private void XuLyChonChuongTrinh_DoiTuongVaMucDich()
        {
            try
            {
                string ct = cbctr?.Text?.Trim() ?? "";

                if (string.IsNullOrWhiteSpace(ct))
                {
                    ResetDoiTuongComboBoxes();
                    ResetMucDichComboBoxes();
                    return;
                }

                // 1. Cập nhật items cbdt1~5 và điền giá trị theo txtkh
                var doiTuongList = LayDanhSachDoiTuong(ct);
                UpdateDoiTuongComboBoxes(doiTuongList);
                CapNhatTatCaDoiTuong();

                // 2. Xử lý cbmd1~5 nếu là Cấp nước sạch
                if (LaCapNuocSach(ct))
                    ApDungCapNuocSach_MucDich();
                else
                    ResetMucDichComboBoxes();
            }
            catch { }
        }

        // ── Xác định danh sách đối tượng theo tên chương trình ──
        private List<string> LayDanhSachDoiTuong(string chuongTrinh)
        {
            var list = new List<string>();

            if (chuongTrinh.IndexOf("Hộ nghèo", StringComparison.OrdinalIgnoreCase) >= 0 &&
                chuongTrinh.IndexOf("cận", StringComparison.OrdinalIgnoreCase) < 0 &&
                chuongTrinh.IndexOf("thoát", StringComparison.OrdinalIgnoreCase) < 0)
            {
                list.Add("Hộ nghèo");
            }
            else if (chuongTrinh.IndexOf("cận nghèo", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                list.Add("Hộ cận nghèo");
            }
            else if (chuongTrinh.IndexOf("thoát nghèo", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                list.Add("Hộ mới thoát nghèo");
            }
            else if (chuongTrinh.IndexOf("Sản xuất kinh doanh", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     chuongTrinh.IndexOf("SXKD", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                list.Add("Hộ GĐ SXKD VKK");
            }
            else if (chuongTrinh.IndexOf("việc làm", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                list.Add("Người lao động");
                list.Add("NLĐ là người DTTS");
            }
            else if (LaCapNuocSach(chuongTrinh))
            {
                list.Add("HGĐ cư trú tại VNT");
            }

            return list;
        }

        // ── Kiểm tra có phải chương trình Cấp nước sạch không ──
        private bool LaCapNuocSach(string chuongTrinh)
        {
            return chuongTrinh.IndexOf("nước sạch", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   chuongTrinh.IndexOf("vệ sinh môi trường", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        // ── Khi chọn Cấp nước sạch: giới hạn cbmd1~5 chỉ còn 2 lựa chọn, khóa nhập tự do ──
        private void ApDungCapNuocSach_MucDich()
        {
            try
            {
                var cbmdList = new[] { cbmd1, cbmd2, cbmd3, cbmd4, cbmd5 };
                foreach (var cb in cbmdList)
                {
                    if (cb == null) continue;
                    cb.Items.Clear();
                    foreach (var item in MucDichCapNuocSach)
                        cb.Items.Add(item);
                    cb.DropDownStyle = ComboBoxStyle.DropDownList;
                    // Giữ giá trị cũ nếu còn hợp lệ, không thì xóa
                    if (cb.Items.Contains(cb.Text))
                        cb.Text = cb.Text;
                    else
                        cb.Text = "";
                }
            }
            catch { }
        }

        // ── Reset cbmd1~5 về danh sách thông thường, mở DropDown tự do ──
        private void ResetMucDichComboBoxes()
        {
            try
            {
                var cbmdList = new[] { cbmd1, cbmd2, cbmd3, cbmd4, cbmd5 };
                foreach (var cb in cbmdList)
                {
                    if (cb == null) continue;
                    var oldText = cb.Text;
                    cb.Items.Clear();
                    foreach (var item in MucDichThongThuong)
                        cb.Items.Add(item);
                    cb.DropDownStyle = ComboBoxStyle.DropDown;
                    // Giữ giá trị cũ nếu còn hợp lệ, nếu là item nước sạch thì xóa
                    cb.Text = Array.IndexOf(MucDichCapNuocSach, oldText) >= 0 ? "" : oldText;
                }
            }
            catch { }
        }
    }
}
