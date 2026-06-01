using System;
using System.Windows.Forms;

namespace HOSONHCS
{
    // ============================================================
    // FILE XỬ LÝ LOGIC ĐỐI TƯỢNG (cbdt1~5) THEO TÊN KHÁCH VÀ CHƯƠNG TRÌNH
    //
    // QUY TẮC:
    //   • cbdt1~5 luôn bị KHOÁ (DropDownList, Enabled = false)
    //   • Chỉ điền giá trị vào cbdtN khi txttkhN có nội dung
    //   • txttkhN trống → cbdtN trống
    //   • Giá trị điền vào = đối tượng tương ứng với cbctr hiện tại
    //     (do Form2.LogicChuongTrinh.cs cung cấp qua LayDanhSachDoiTuong)
    //
    // ĐỂ SỬA:
    //   • Thêm/sửa đối tượng theo chương trình → Form2.LogicChuongTrinh.cs
    //   • Thay đổi quy tắc khoá/mở cbdt → sửa file này
    // ============================================================

    public partial class Form2 : Form
    {
        // ── Đăng ký events txtkh1~5 TextChanged (gọi từ Form2 constructor) ──
        private void DangKySuKien_DoiTuong()
        {
            try { txtkh1.TextChanged += Txtkh_TextChanged; } catch { }
            try { txtkh2.TextChanged += Txtkh_TextChanged; } catch { }
            try { txtkh3.TextChanged += Txtkh_TextChanged; } catch { }
            try { txtkh4.TextChanged += Txtkh_TextChanged; } catch { }
            try { txtkh5.TextChanged += Txtkh_TextChanged; } catch { }
        }

        // ── Khi bất kỳ txtkh nào thay đổi → cập nhật toàn bộ cbdt ──
        private void Txtkh_TextChanged(object sender, EventArgs e)
        {
            CapNhatTatCaDoiTuong();
        }

        // ── Cập nhật cbdt1~5 dựa theo txtkh1~5 và chương trình đang chọn ──
        internal void CapNhatTatCaDoiTuong()
        {
            try
            {
                string ct = cbctr?.Text?.Trim() ?? "";
                var doiTuongList = LayDanhSachDoiTuong(ct);

                // Lấy giá trị đối tượng đầu tiên trong danh sách (nếu có)
                string giaTriDoiTuong = doiTuongList.Count > 0 ? doiTuongList[0] : "";

                CapNhatMotDoiTuong(cbdt1, txtkh1, giaTriDoiTuong);
                CapNhatMotDoiTuong(cbdt2, txtkh2, giaTriDoiTuong);
                CapNhatMotDoiTuong(cbdt3, txtkh3, giaTriDoiTuong);
                CapNhatMotDoiTuong(cbdt4, txtkh4, giaTriDoiTuong);
                CapNhatMotDoiTuong(cbdt5, txtkh5, giaTriDoiTuong);
            }
            catch { }
        }

        // ── Xử lý 1 cặp (cbdtN, txttkhN) ──
        private void CapNhatMotDoiTuong(ComboBox cbdt, TextBox txtkh, string giaTriDoiTuong)
        {
            if (cbdt == null || txtkh == null) return;

            // Luôn khoá, không cho nhập tự do
            cbdt.DropDownStyle = ComboBoxStyle.DropDownList;
            cbdt.Enabled = false;

            bool coTen = !string.IsNullOrWhiteSpace(txtkh.Text);

            if (coTen && !string.IsNullOrWhiteSpace(giaTriDoiTuong))
            {
                // Đảm bảo item tồn tại trong danh sách trước khi set Text
                if (!cbdt.Items.Contains(giaTriDoiTuong))
                    cbdt.Items.Add(giaTriDoiTuong);
                cbdt.Text = giaTriDoiTuong;
            }
            else
            {
                // DropDownList không nhận Text = "", phải dùng SelectedIndex = -1
                cbdt.SelectedIndex = -1;
            }
        }

        // ── Khoá toàn bộ cbdt1~5 (gọi khi khởi tạo form) ──
        private void KhoaTatCaDoiTuong()
        {
            try
            {
                var cbdtArr = new[] { cbdt1, cbdt2, cbdt3, cbdt4, cbdt5 };
                foreach (var cb in cbdtArr)
                {
                    if (cb == null) continue;
                    cb.DropDownStyle = ComboBoxStyle.DropDownList;
                    cb.Enabled = false;
                }
            }
            catch { }
        }
    }
}
