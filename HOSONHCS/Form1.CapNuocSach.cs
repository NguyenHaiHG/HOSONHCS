using System;
using System.Globalization;
using System.Windows.Forms;

namespace HOSONHCS
{
    // ============================================================
    // MODULE: Cấp nước sạch và vệ sinh môi trường nông thôn
    // Xử lý toàn bộ logic tự động cho chương trình này:
    //   • cbDoituong1/2  → mở khoá, cho phép nhập tự do
    //   • cbSotien       → khi rời/chọn: chia đôi vào cbSotien1/2, khoá
    //   • cbPhuongan     → tự điền cbmucdich1/2 theo phương án, mở khoá cho nhập
    // ============================================================
    partial class Form1
    {
        private const string CTR_CAP_NUOC_SACH =
            "Cấp nước sạch và vệ sinh môi trường nông thôn";

        // ─── Kiểm tra chương trình hiện tại ────────────────────
        private bool IsCapNuocSach()
        {
            try
            {
                return string.Equals(
                    (cbChuongtrinh?.Text ?? "").Trim(),
                    CTR_CAP_NUOC_SACH,
                    StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        // ─── 1. Mở khoá cbDoituong1/2, giới hạn cbPhuongan, khoá cbThoihanvay/cbPhanky ───
        private void ApplyCapNuocSach_Doituong()
        {
            try
            {
                if (cbDoituong1 != null)
                {
                    cbDoituong1.Enabled = true;
                    cbDoituong1.DropDownStyle = ComboBoxStyle.DropDown;
                    cbDoituong1.Text = "01 công trình";
                }
                if (cbDoituong2 != null)
                {
                    cbDoituong2.Enabled = true;
                    cbDoituong2.DropDownStyle = ComboBoxStyle.DropDown;
                    cbDoituong2.Text = "01 công trình";
                }
                if (cbThoihanvay != null)
                {
                    cbThoihanvay.Enabled = false;
                    cbThoihanvay.DropDownStyle = ComboBoxStyle.DropDown;
                    cbThoihanvay.Text = "60 tháng";
                }
                if (cbPhanky != null)
                {
                    cbPhanky.Enabled = false;
                    cbPhanky.DropDownStyle = ComboBoxStyle.DropDown;
                    cbPhanky.Text = "12 tháng";
                }
                ApplyCapNuocSach_Phuongan_Restrict();
            }
            catch { }
        }

        // ─── Giới hạn cbPhuongan chỉ 2 phương án Cấp nước sạch ────────────
        private void ApplyCapNuocSach_Phuongan_Restrict()
        {
            try
            {
                if (cbPhuongan == null) return;
                cbPhuongan.Items.Clear();
                cbPhuongan.Items.Add("Nâng cấp, sửa chữa CTNS, CTVS          ");
                cbPhuongan.Items.Add("Xây mới CTNS, CTVS");
                cbPhuongan.DropDownStyle = ComboBoxStyle.DropDownList;
                cbPhuongan.Enabled = true;
                cbPhuongan.Text = "";
            }
            catch { }
        }

        // ─── 2. Chia đôi cbSotien → cbSotien1/2, khoá ──────────
        private void ApplyCapNuocSach_SplitSotien()
        {
            try
            {
                if (!IsCapNuocSach()) return;
                if (cbSotien == null || cbSotien1 == null || cbSotien2 == null) return;

                var total = ParseMoneyStringToLong(cbSotien.Text);

                // Luôn khoá cbSotien1/2 khi đang ở chương trình này
                cbSotien1.Enabled = false;
                cbSotien2.Enabled = false;

                if (total <= 0) return;

                var half = total / 2;
                var formatted = string.Format(
                    CultureInfo.InvariantCulture, "{0:N0}", half).Replace(",", ".");

                suppressMoneyChange = true;
                try
                {
                    cbSotien1.Text = formatted;
                    cbSotien2.Text = formatted;
                }
                finally { suppressMoneyChange = false; }
            }
            catch { suppressMoneyChange = false; }
        }

        // ─── Event handler: gắn vào cbSotien.Leave + SelectedIndexChanged
        private void CbSotien_CapNuocSach_Leave(object sender, EventArgs e)
        {
            ApplyCapNuocSach_SplitSotien();
        }

        private void CbSotien_CapNuocSach_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyCapNuocSach_SplitSotien();
        }

        // ─── 3. cbPhuongan → cbmucdich1/2 ──────────────────────
        private void ApplyCapNuocSach_Phuongan(string phuongan)
        {
            try
            {
                if (cbmucdich1 == null || cbmucdich2 == null) return;

                var pa = (phuongan ?? "").Trim();

                if (string.Equals(pa, CTR_CAP_NUOC_SACH,
                    StringComparison.OrdinalIgnoreCase))
                {
                    cbmucdich1.DropDownStyle = ComboBoxStyle.DropDown;
                    cbmucdich1.Enabled = true;
                    cbmucdich1.Text = "01 công trình";

                    cbmucdich2.DropDownStyle = ComboBoxStyle.DropDown;
                    cbmucdich2.Enabled = true;
                    cbmucdich2.Text = "01 công trình";
                }
                else if (string.Equals(pa, "Nâng cấp, sửa chữa CTNS, CTVS",
                    StringComparison.OrdinalIgnoreCase))
                {
                    cbmucdich1.DropDownStyle = ComboBoxStyle.DropDown;
                    cbmucdich1.Enabled = true;
                    cbmucdich1.Text = "Nâng cấp, sửa chữa CTNS";

                    cbmucdich2.DropDownStyle = ComboBoxStyle.DropDown;
                    cbmucdich2.Enabled = true;
                    cbmucdich2.Text = "Nâng cấp, sửa chữa CTVS";
                }
                else if (string.Equals(pa, "Xây mới CTNS, CTVS",
                    StringComparison.OrdinalIgnoreCase))
                {
                    cbmucdich1.DropDownStyle = ComboBoxStyle.DropDown;
                    cbmucdich1.Enabled = true;
                    cbmucdich1.Text = "Xây mới CTNS";

                    cbmucdich2.DropDownStyle = ComboBoxStyle.DropDown;
                    cbmucdich2.Enabled = true;
                    cbmucdich2.Text = "Xây mới CTVS";
                }
            }
            catch { }
        }

        // ─── Reset: trả lại trạng thái bình thường khi đổi CT ──
        private void ResetCapNuocSach()
        {
            try
            {
                if (cbDoituong1 != null)
                {
                    cbDoituong1.Enabled = true;
                    cbDoituong1.Text = "";
                }
                if (cbDoituong2 != null)
                {
                    cbDoituong2.Enabled = true;
                    cbDoituong2.Text = "";
                }
                if (cbSotien1 != null) cbSotien1.Enabled = true;
                if (cbSotien2 != null) cbSotien2.Enabled = true;
                if (cbThoihanvay != null)
                {
                    cbThoihanvay.Enabled = true;
                    cbThoihanvay.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                if (cbPhanky != null)
                {
                    cbPhanky.Enabled = true;
                    cbPhanky.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                if (cbPhuongan != null)
                {
                    cbPhuongan.Items.Clear();
                    cbPhuongan.Items.Add("Mua trâu sinh sản");
                    cbPhuongan.Items.Add("Nuôi trâu sinh sản");
                    cbPhuongan.Items.Add("Mua bò sinh sản");
                    cbPhuongan.Items.Add("Nuôi bò sinh sản");
                    cbPhuongan.Items.Add("Mua dê sinh sản");
                    cbPhuongan.Items.Add("Nuôi dê sinh sản");
                    cbPhuongan.Items.Add("Nuôi lợn sinh sản");
                    cbPhuongan.Items.Add("Nuôi lợn");
                    cbPhuongan.Items.Add("Trồng cây quế");
                    cbPhuongan.Items.Add("Trồng cây keo");
                    cbPhuongan.Items.Add("Trồng cây mỡ");
                    cbPhuongan.Items.Add("Trồng cây cam");
                    cbPhuongan.Items.Add("Mở rộng cửa hàng tạp hoá");
                    cbPhuongan.Items.Add("Mở rộng cửa hàng ăn uống");
                    cbPhuongan.Items.Add("Mở rộng cửa hàng bán quần áo");
                    cbPhuongan.Items.Add("Nâng cấp, sửa chữa CTNS, CTVS          ");
                    cbPhuongan.Items.Add("Xây mới CTNS, CTVS");
                    cbPhuongan.Items.Add("Trồng và chăm sóc cây cà phê   ");
                    cbPhuongan.Items.Add("Trồng và chăm sóc cây cao su");
                    cbPhuongan.Items.Add("Trồng cây ăn quả");
                    cbPhuongan.Items.Add("Trồng cây bời lời");
                    cbPhuongan.Items.Add("Trồng cây tiêu                              ");
                    cbPhuongan.DropDownStyle = ComboBoxStyle.DropDown;
                    cbPhuongan.Enabled = true;
                }
            }
            catch { }
        }
    }
}
