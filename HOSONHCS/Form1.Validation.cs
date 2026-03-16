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

            // Kiểm tra các DateTimePicker
            if (dateNgaysinh == null || !dateNgaysinh.Checked) missingFields.Add("Ngày sinh");
            if (dateNgaycapCCCD == null || !dateNgaycapCCCD.Checked) missingFields.Add("Ngày cấp CCCD");

            // Kiểm tra đặc biệt cho datendhcccd - không kiểm tra Checked vì ShowCheckBox = false
            if (datendhcccd == null)
            {
                missingFields.Add("Ngày hết hạn CCCD (control không tồn tại)");
            }
            else
            {
                try
                {
                    // Kiểm tra xem value có hợp lệ không
                    var _ = datendhcccd.Value;
                }
                catch
                {
                    missingFields.Add("Ngày hết hạn CCCD (giá trị không hợp lệ)");
                }
            }

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

            return true;
        }
    }
}
