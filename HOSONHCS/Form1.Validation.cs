using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace HOSONHCS
{
    public partial class Form1 : Form
    {
        // Kiểm tra các trường bắt buộc trước khi tạo hồ sơ
        private bool ValidateRequiredFields()
        {
            var missingFields = new List<string>();

            // Kiểm tra từng trường và thêm vào danh sách nếu thiếu
            if (string.IsNullOrWhiteSpace(cbPGD?.Text)) missingFields.Add("PGD");
            if (string.IsNullOrWhiteSpace(cbXa?.Text)) missingFields.Add("Xã");
            if (string.IsNullOrWhiteSpace(cbThon?.Text)) missingFields.Add("Thôn");
            if (string.IsNullOrWhiteSpace(cbTo?.Text)) missingFields.Add("Tổ");
            if (string.IsNullOrWhiteSpace(cbHoi?.Text)) missingFields.Add("Hội quản lý");
            if (string.IsNullOrWhiteSpace(cbChuongtrinh?.Text)) missingFields.Add("Chương trình vay");
            if (string.IsNullOrWhiteSpace(cbThoihanvay?.Text)) missingFields.Add("Thời hạn vay");
            if (string.IsNullOrWhiteSpace(cbPhanky?.Text)) missingFields.Add("Phân kỳ");
            if (string.IsNullOrWhiteSpace(cbPhuongan?.Text)) missingFields.Add("Phương án vay");
            if (string.IsNullOrWhiteSpace(cbSotien?.Text)) missingFields.Add("Số tiền vay");
            if (string.IsNullOrWhiteSpace(cbVtc?.Text)) missingFields.Add("Vốn tự có");
            if (string.IsNullOrWhiteSpace(cbmucdich1?.Text)) missingFields.Add("Mục đích vay vốn 1");
            if (string.IsNullOrWhiteSpace(cbDoituong1?.Text)) missingFields.Add("Đối tượng vay vốn 1");
            if (string.IsNullOrWhiteSpace(cbSotien1?.Text)) missingFields.Add("Số tiền vay vốn 1");
            if (string.IsNullOrWhiteSpace(txtHoten?.Text)) missingFields.Add("Họ và tên");
            if (string.IsNullOrWhiteSpace(cbNhandang?.Text)) missingFields.Add("Nhân dạng");
            if (string.IsNullOrWhiteSpace(txtSocccd?.Text)) missingFields.Add("Số CCCD");
            if (string.IsNullOrWhiteSpace(cbDantoc?.Text)) missingFields.Add("Dân tộc");
            if (string.IsNullOrWhiteSpace(cbGioitinh?.Text)) missingFields.Add("Giới tính");
            if (string.IsNullOrWhiteSpace(cbDoituong?.Text)) missingFields.Add("Đối tượng");
            if (string.IsNullOrWhiteSpace(txtSdt?.Text)) missingFields.Add("Số điện thoại");
            if (string.IsNullOrWhiteSpace(txtNhankhau?.Text)) missingFields.Add("Nhân khẩu");

            // Kiểm tra các DateTimePicker bắt buộc
            // (datentk1/2/3, dateGn, dateDH, dateLaphs được miễn kiểm tra)
            if (dateNgaysinh == null || dateNgaysinh.CustomFormat == " ") missingFields.Add("Ngày sinh");
            if (dateNgaycapCCCD == null || dateNgaycapCCCD.CustomFormat == " ") missingFields.Add("Ngày cấp CCCD");
            if (datendhcccd == null || datendhcccd.CustomFormat == " ") missingFields.Add("Ngày hết hạn CCCD");

            // Nếu có trường thiếu, hiển thị thông báo chi tiết
            if (missingFields.Count > 0)
            {
                string message = "⚠️ Vui lòng nhập đầy đủ các trường sau:\n\n";
                for (int i = 0; i < missingFields.Count; i++)
                {
                    message += $"{i + 1}. {missingFields[i]}\n";
                }
                message += $"\n📝 Tổng cộng: {missingFields.Count} trường bị thiếu";

                MessageBox.Show(message, "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            // Kiểm tra tuổi tối thiểu 18 (tính chính xác theo ngày sinh)
            try
            {
                if (dateNgaysinh != null && dateNgaysinh.CustomFormat != " ")
                {
                    var ngaysinh = dateNgaysinh.Value.Date;
                    int age = DateTime.Today.Year - ngaysinh.Year;
                    if (DateTime.Today < ngaysinh.AddYears(age)) age--;

                    if (age < 18)
                    {
                        MessageBox.Show(
                            $"🚫 Khách hàng chưa đủ 18 tuổi!\n\n" +
                            $"  Ngày sinh:     {ngaysinh:dd/MM/yyyy}\n" +
                            $"  Tuổi hiện tại: {age} tuổi\n\n" +
                            $"Không thể tạo hồ sơ cho khách hàng dưới 18 tuổi.",
                            "Không đủ tuổi", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return false;
                    }
                }
            }
            catch { }

            // Kiểm tra tổng cbSotien1 + cbSotien2 phải bằng cbSotien
            try
            {
                long sotienTotal = ParseMoneyStringToLong(cbSotien?.Text ?? "");
                long sotien1Val  = ParseMoneyStringToLong(cbSotien1?.Text ?? "");
                long sotien2Val  = ParseMoneyStringToLong(cbSotien2?.Text ?? "");

                if (sotienTotal > 0 && (sotien1Val + sotien2Val) != sotienTotal)
                {
                    string fmt(long v) => string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:N0}", v).Replace(",", ".");
                    MessageBox.Show(
                        $"⚠️ Tổng số tiền theo mục đích không khớp với số tiền vay!\n\n" +
                        $"  Số tiền vay:      {fmt(sotienTotal)} đ\n" +
                        $"  Thành tiền 1:       {fmt(sotien1Val)} đ\n" +
                        $"  Thành tiền 2:       {fmt(sotien2Val)} đ\n" +
                        $"  Tổng mục đích:    {fmt(sotien1Val + sotien2Val)} đ\n\n" +
                        $"Yêu cầu: Số tiền vay = Thành tiền 1 + Thành tiền 2",
                        "⚠️ Lỗi số tiền", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }
            catch { }

            return true;
        }

        // Lấy chỉ các chữ số từ chuỗi (dùng để chuẩn hoá SDT trước khi so sánh)
        private static string DigitsOnly(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            var sb = new System.Text.StringBuilder();
            foreach (char ch in s)
                if (char.IsDigit(ch)) sb.Append(ch);
            return sb.ToString();
        }

        // Kiểm tra trùng CCCD hoặc SDT với khách hàng đã có trong hệ thống
        private bool ValidateDuplicateCccdSdt()
        {
            string inputCccd = "";
            string inputSdt = "";
            string inputCccd1 = "", inputCccd2 = "", inputCccd3 = "";

            try { inputCccd  = txtSocccd?.Text.Trim() ?? ""; } catch { }
            try { inputSdt   = txtSdt?.Text.Trim()    ?? ""; } catch { }
            try { inputCccd1 = txtcccd1?.Text.Trim()  ?? ""; } catch { }
            try { inputCccd2 = txtcccd2?.Text.Trim()  ?? ""; } catch { }
            try { inputCccd3 = txtcccd3?.Text.Trim()  ?? ""; } catch { }

            // --- Kiểm tra CCCD chính không được trùng với CCCD người thừa kế trong cùng hồ sơ ---
            if (!string.IsNullOrWhiteSpace(inputCccd))
            {
                if (!string.IsNullOrWhiteSpace(inputCccd1) &&
                    string.Equals(inputCccd, inputCccd1, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show(
                        $"⚠️ Số CCCD khách hàng trùng với CCCD người thừa kế thứ 1!\n\n" +
                        $"Số CCCD \"{inputCccd}\" không được giống nhau trong cùng một hồ sơ.",
                        "Trùng CCCD nội bộ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                if (!string.IsNullOrWhiteSpace(inputCccd2) &&
                    string.Equals(inputCccd, inputCccd2, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show(
                        $"⚠️ Số CCCD khách hàng trùng với CCCD người thừa kế thứ 2!\n\n" +
                        $"Số CCCD \"{inputCccd}\" không được giống nhau trong cùng một hồ sơ.",
                        "Trùng CCCD nội bộ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                if (!string.IsNullOrWhiteSpace(inputCccd3) &&
                    string.Equals(inputCccd, inputCccd3, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show(
                        $"⚠️ Số CCCD khách hàng trùng với CCCD người thừa kế thứ 3!\n\n" +
                        $"Số CCCD \"{inputCccd}\" không được giống nhau trong cùng một hồ sơ.",
                        "Trùng CCCD nội bộ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            // --- Kiểm tra các CCCD người thừa kế không được trùng nhau trong cùng hồ sơ ---
            if (!string.IsNullOrWhiteSpace(inputCccd1) && !string.IsNullOrWhiteSpace(inputCccd2) &&
                string.Equals(inputCccd1, inputCccd2, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(
                    $"⚠️ CCCD người thừa kế thứ 1 và thứ 2 trùng nhau!\n\n" +
                    $"Số CCCD \"{inputCccd1}\" không được giống nhau trong cùng một hồ sơ.",
                    "Trùng CCCD nội bộ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (!string.IsNullOrWhiteSpace(inputCccd1) && !string.IsNullOrWhiteSpace(inputCccd3) &&
                string.Equals(inputCccd1, inputCccd3, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(
                    $"⚠️ CCCD người thừa kế thứ 1 và thứ 3 trùng nhau!\n\n" +
                    $"Số CCCD \"{inputCccd1}\" không được giống nhau trong cùng một hồ sơ.",
                    "Trùng CCCD nội bộ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (!string.IsNullOrWhiteSpace(inputCccd2) && !string.IsNullOrWhiteSpace(inputCccd3) &&
                string.Equals(inputCccd2, inputCccd3, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(
                    $"⚠️ CCCD người thừa kế thứ 2 và thứ 3 trùng nhau!\n\n" +
                    $"Số CCCD \"{inputCccd2}\" không được giống nhau trong cùng một hồ sơ.",
                    "Trùng CCCD nội bộ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (customers == null || customers.Count == 0) return true;

            // Chuẩn hoá SDT: chỉ giữ chữ số để so sánh bất kể format (có/không có dấu chấm)
            string inputSdtDigits = DigitsOnly(inputSdt);

            for (int i = 0; i < customers.Count; i++)
            {
                // Bỏ qua khách hàng đang sửa
                if (editingIndex >= 0 && i == editingIndex) continue;

                var c = customers[i];

                // Tập hợp tất cả CCCD của khách hàng đã có (chính + người thừa kế)
                var existingCccds = new List<string>();
                if (!string.IsNullOrWhiteSpace(c.Socccd))    existingCccds.Add(c.Socccd.Trim());
                if (!string.IsNullOrWhiteSpace(c.CccdNtk1))  existingCccds.Add(c.CccdNtk1.Trim());
                if (!string.IsNullOrWhiteSpace(c.CccdNtk2))  existingCccds.Add(c.CccdNtk2.Trim());
                if (!string.IsNullOrWhiteSpace(c.CccdNtk3))  existingCccds.Add(c.CccdNtk3.Trim());

                // --- Kiểm tra CCCD chính ---
                if (!string.IsNullOrWhiteSpace(inputCccd) &&
                    existingCccds.Exists(ec => string.Equals(inputCccd, ec, StringComparison.OrdinalIgnoreCase)))
                {
                    MessageBox.Show(
                        $"⚠️ Trùng số căn cước công dân!\n\n" +
                        $"Số CCCD \"{inputCccd}\" đã tồn tại trong hệ thống:\n" +
                        $"👤 Khách hàng: {c.Hoten}\n\n" +
                        $"Không thể tạo hồ sơ trùng CCCD.",
                        "Trùng CCCD",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return false;
                }

                // --- Kiểm tra CCCD người thừa kế thứ 1 ---
                if (!string.IsNullOrWhiteSpace(inputCccd1) &&
                    existingCccds.Exists(ec => string.Equals(inputCccd1, ec, StringComparison.OrdinalIgnoreCase)))
                {
                    MessageBox.Show(
                        $"⚠️ Trùng căn cước công dân người thừa kế thứ 1!\n\n" +
                        $"Số CCCD \"{inputCccd1}\" đã tồn tại trong hệ thống:\n" +
                        $"👤 Khách hàng: {c.Hoten}\n\n" +
                        $"Không thể tạo hồ sơ trùng CCCD người thừa kế.",
                        "Trùng CCCD người thừa kế thứ 1",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return false;
                }

                // --- Kiểm tra CCCD người thừa kế thứ 2 ---
                if (!string.IsNullOrWhiteSpace(inputCccd2) &&
                    existingCccds.Exists(ec => string.Equals(inputCccd2, ec, StringComparison.OrdinalIgnoreCase)))
                {
                    MessageBox.Show(
                        $"⚠️ Trùng căn cước công dân người thừa kế thứ 2!\n\n" +
                        $"Số CCCD \"{inputCccd2}\" đã tồn tại trong hệ thống:\n" +
                        $"👤 Khách hàng: {c.Hoten}\n\n" +
                        $"Không thể tạo hồ sơ trùng CCCD người thừa kế.",
                        "Trùng CCCD người thừa kế thứ 2",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return false;
                }

                // --- Kiểm tra CCCD người thừa kế thứ 3 ---
                if (!string.IsNullOrWhiteSpace(inputCccd3) &&
                    existingCccds.Exists(ec => string.Equals(inputCccd3, ec, StringComparison.OrdinalIgnoreCase)))
                {
                    MessageBox.Show(
                        $"⚠️ Trùng căn cước công dân người thừa kế thứ 3!\n\n" +
                        $"Số CCCD \"{inputCccd3}\" đã tồn tại trong hệ thống:\n" +
                        $"👤 Khách hàng: {c.Hoten}\n\n" +
                        $"Không thể tạo hồ sơ trùng CCCD người thừa kế.",
                        "Trùng CCCD người thừa kế thứ 3",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return false;
                }

                // --- Kiểm tra trùng số điện thoại ---
                // Chuẩn hoá cả hai phía: chỉ lấy chữ số (tránh lệch format dấu chấm)
                string existingSdtDigits = DigitsOnly(c.Sdt);
                if (!string.IsNullOrWhiteSpace(inputSdtDigits) &&
                    !string.IsNullOrWhiteSpace(existingSdtDigits) &&
                    string.Equals(inputSdtDigits, existingSdtDigits, StringComparison.Ordinal))
                {
                    MessageBox.Show(
                        $"⚠️ Trùng số điện thoại!\n\n" +
                        $"Số điện thoại \"{inputSdt}\" đã tồn tại trong hệ thống:\n" +
                        $"👤 Khách hàng: {c.Hoten}\n\n" +
                        $"Không thể tạo hồ sơ trùng số điện thoại.",
                        "Trùng số điện thoại",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return false;
                }
            }

            return true;
        }
    }
}
