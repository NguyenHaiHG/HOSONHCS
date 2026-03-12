using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HOSONHCS
{
    /// <summary>
    /// Form2 - Form xuất hồ sơ nhóm
    /// Form này dùng để nhập thông tin tổ viên và xuất file Word theo mẫu
    /// </summary>
    public partial class Form2 : Form
    {
        // ============================================
        // BIẾN TOÀN CỤC
        // ============================================

        /// <summary>
        /// Danh sách khách hàng đã chọn khi mở form từ Form1
        /// Dùng để lưu trữ và hiển thị trong DataGridView
        /// </summary>
        private List<Customer> selectedCustomers = new List<Customer>();

        /// <summary>
        /// Danh sách lịch sử xuất hồ sơ (Xã - Thôn - Tổ)
        /// Hiển thị trong dgv2
        /// </summary>
        private System.ComponentModel.BindingList<ExportHistory> exportHistories = new System.ComponentModel.BindingList<ExportHistory>();

        /// <summary>
        /// Đường dẫn file JSON để lưu trạng thái form
        /// File này lưu tất cả thông tin đã nhập để khôi phục khi mở lại form
        /// </summary>
        private string Form2StatePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Form2State.json");

        /// <summary>
        /// Cờ x ác định có nên load trạng thái đã lưu hay không
        /// = false khi mở form từ Form1 với khách hàng đã chọn
        /// = true khi mở form độc lập
        /// </summary>
        private bool shouldLoadState = true;

        /// <summary>
        /// Đường dẫn folder lưu trữ file JSON của các tổ
        /// Tương tự Customers folder trong Form1
        /// </summary>
        private string GetToFolderPath() => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "To");

        /// <summary>
        /// Danh sách các tổ đã lưu (load từ folder To)
        /// </summary>
        private System.ComponentModel.BindingList<ToData> savedToList = new System.ComponentModel.BindingList<ToData>();

        /// <summary>
        /// Biến lưu văn bản chạy (marquee) trong richTextBox1
        /// </summary>
        private string richTextMarqueeText = "";

        /// <summary>
        /// Vị trí hiện tại của văn bản chạy (marquee)
        /// </summary>
        private int richTextMarqueePosition = 0;

        /// <summary>
        /// Dữ liệu xinman (PGD/Xã/Thôn/Hội/Tổ) được load từ xinman.json
        /// </summary>
        private XinManModel xinmanModel;

        /// <summary>
        /// Dictionary lưu trữ các XinManModel theo tên PGD
        /// Key: Tên PGD, Value: XinManModel tương ứng
        /// </summary>
        private Dictionary<string, XinManModel> xinmanModels = new Dictionary<string, XinManModel>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Cờ xác định chế độ Edit hay Add
        /// = true: Đang chỉnh sửa tổ đã chọn từ dgv2
        /// = false: Đang tạo tổ mới
        /// </summary>
        private bool isEditMode = false;

        /// <summary>
        /// Lưu trữ ExportHistory đang được chỉnh sửa
        /// Null nếu đang ở chế độ Add (tạo mới)
        /// </summary>
        private ExportHistory currentEditingHistory = null;

        /// <summary>
        /// Cờ để tránh vòng lặp vô hạn khi cập nhật cbThon2 và cbTotruong
        /// = true: Đang cập nhật programmatically, không trigger event
        /// </summary>
        private bool isUpdatingThonTotruong = false;

        /// <summary>
        /// Constructor mặc định của Form2
        /// Khởi tạo các components và đăng ký event handlers
        /// </summary>
        public Form2()
        {
            InitializeComponent();

            // Đăng ký sự kiện click cho nút xuất Word và nút xóa
            btn03to.Click += Btn03to_Click;
            btnxoa.Click += BtnXoa_Click;
            btntaoto.Click += BtnTaoTo_Click;

            // DataGridView (dgv2) dùng để hiển thị lịch sử xuất
            // Đăng ký sự kiện click vào ô để load dữ liệu lên form
            try { dgv2.CellClick += Dgv2_CellClick; } catch { }

            // Đăng ký xử lý nhập liệu cho các ô tiền (cbtien1-5)
            // Chỉ cho phép nhập số và tự động định dạng với dấu chấm phân cách hàng nghìn
            try { cbtien1.KeyPress += CbMoney_KeyPress; cbtien1.TextChanged += CbMoney_TextChanged; } catch { }
            try { cbtien2.KeyPress += CbMoney_KeyPress; cbtien2.TextChanged += CbMoney_TextChanged; } catch { }
            try { cbtien3.KeyPress += CbMoney_KeyPress; cbtien3.TextChanged += CbMoney_TextChanged; } catch { }
            try { cbtien4.KeyPress += CbMoney_KeyPress; cbtien4.TextChanged += CbMoney_TextChanged; } catch { }
            try { cbtien5.KeyPress += CbMoney_KeyPress; cbtien5.TextChanged += CbMoney_TextChanged; } catch { }

            // Đăng ký xử lý nhập liệu cho các ô tên
            // Chỉ cho phép nhập chữ cái, không cho phép nhập số
            try { txtkh1.KeyPress += TextLettersOnly_KeyPress; txtkh2.KeyPress += TextLettersOnly_KeyPress; txtkh3.KeyPress += TextLettersOnly_KeyPress; txtkh4.KeyPress += TextLettersOnly_KeyPress; txtkh5.KeyPress += TextLettersOnly_KeyPress; } catch { }

            // Đăng ký sự kiện cho cbpgd2 - Load dữ liệu khi chọn PGD
            try { cbpgd2.SelectedIndexChanged += CbPgd2_SelectedIndexChanged; } catch { }

            // Đăng ký sự kiện cho cbctr - Cập nhật danh sách đối tượng khi chọn chương trình
            try { cbctr.SelectedIndexChanged += Cbctr_SelectedIndexChanged; } catch { }

            // Đăng ký sự kiện cho cbTotruong - Tự động chọn thôn khi chọn tổ trưởng
            try { cbTotruong.SelectedIndexChanged += CbTotruong_SelectedIndexChanged; } catch { }

            // Khởi tạo hiệu ứng chữ chạy cho richTextBox1
            try { InitializeRichTextMarquee(); } catch { }

            // Bind dgv2 với exportHistories để hiển thị lịch sử xuất
            try 
            { 
                if (dgv2 != null)
                {
                    dgv2.DataSource = exportHistories;
                    dgv2.AutoGenerateColumns = true;
                    dgv2.ReadOnly = true;
                }
            } 
            catch { }

            // Đăng ký sự kiện Load form để khôi phục trạng thái đã lưu (nếu mở độc lập)
            this.Load += Form2_Load;
        }

        /// <summary>
        /// Constructor khi mở Form2 từ Form1 với danh sách khách hàng đã chọn
        /// Dùng để hiển thị nhóm khách hàng và xuất hồ sơ nhóm
        /// </summary>
        /// <param name="selected">Danh sách khách hàng đã chọn từ Form1</param>
        public Form2(List<Customer> selected) : this()
        {
            // Không load trạng thái đã lưu khi mở từ Form1 với khách hàng đã chọn
            // Vì dữ liệu sẽ được điền từ danh sách selected
            shouldLoadState = false;

            try
            {
                if (selected != null)
                {
                    // Nhóm khách hàng theo Tổ trưởng (Totruong)
                    // Chỉ hiển thị một dòng cho mỗi tổ trưởng trong DataGridView
                    // Tránh hiển thị nhiều dòng thành viên từ Form1
                    try
                    {
                        var groups = selected
                            .Where(c => c != null)
                            .GroupBy(c => (c.Totruong ?? "").Trim(), StringComparer.OrdinalIgnoreCase)
                            .ToList();

                        selectedCustomers.Clear();

                        foreach (var g in groups)
                        {
                            var leaderName = g.Key;
                            // Bỏ qua các mục không có tên tổ trưởng
                            if (string.IsNullOrWhiteSpace(leaderName))
                                continue;

                            var first = g.FirstOrDefault();
                            // Tạo đối tượng khách hàng tổng hợp cho nhóm
                            var summary = new Customer
                            {
                                // Hiển thị tên tổ trưởng trong cột tên chính của lưới
                                Hoten = leaderName,
                                Totruong = first?.Totruong ?? leaderName,
                                Xa = first?.Xa ?? "",
                                Chuongtrinh = first?.Chuongtrinh ?? ""
                            };
                            selectedCustomers.Add(summary);
                        }

                        // Cập nhật danh sách khách hàng đã nhóm
                        // (dgv2 dùng để hiển thị lịch sử xuất, không hiển thị selectedCustomers)
                    }
                    catch { }

                    // KHÔNG điền txtkh1..txtkh5 từ danh sách thành viên được truyền từ Form1
                    // Form chỉ hiển thị thông tin người tổ chức theo mặc định
                    // Chỉ điền các trường chung cấp cao nhất từ khách hàng đầu tiên (nếu có)
                    var firstCustomer = selected.FirstOrDefault();
                    if (firstCustomer != null)
                    {
                        try { cbTotruong.Text = firstCustomer.Totruong ?? ""; } catch { }
                        try { cbXa2.Text = firstCustomer.Xa ?? ""; } catch { }
                        try { cbctr.Text = firstCustomer.Chuongtrinh ?? ""; } catch { }
                        try { cbpgd2.Text = firstCustomer.PGD ?? ""; } catch { }
                    }
                }
            }
            catch { }
        }

        // ============================================
        // XỬ LÝ SỰ KIỆN CLICK NÚT XUẤT WORD
        // ============================================

        /// <summary>
        /// Xử lý sự kiện click nút "Xuất Word" (btn03to)
        /// Thu thập dữ liệu từ form và xuất file Word theo mẫu
        /// </summary>
        private async void Btn03to_Click(object sender, EventArgs e)
        {
            try
            {
                // -------- THÔNG TIN CHUNG --------
                // Lấy thông tin tổ trưởng, xã, thôn và chương trình
                string totruong = Clean(cbTotruong.Text);
                string xa = Clean(cbXa2.Text);
                string thon = Clean(cbThon2.Text);
                string chuongtrinh = Clean(cbctr.Text);

                // -------- THÔNG TIN TỔ VIÊN --------
                // Mảng lưu họ tên 5 tổ viên (index 0 để trống, dùng index 1-5)
                string[] hoten = {
                    "",
                    Clean(txtkh1.Text),
                    Clean(txtkh2.Text),
                    Clean(txtkh3.Text),
                    Clean(txtkh4.Text),
                    Clean(txtkh5.Text)
                };

                // Mảng lưu số tiền của 5 tổ viên (đã format với dấu chấm phân cách)
                string[] sotien = {
                    "",
                    Money(cbtien1.Text),
                    Money(cbtien2.Text),
                    Money(cbtien3.Text),
                    Money(cbtien4.Text),
                    Money(cbtien5.Text)
                };

                // Mảng lưu phương án của 5 tổ viên
                string[] phuongan = {
                    "",
                    Clean(cbmd1.Text),
                    Clean(cbmd2.Text),
                    Clean(cbmd3.Text),
                    Clean(cbmd4.Text),
                    Clean(cbmd5.Text)
                };

                // Mảng lưu thời hạn vay của 5 tổ viên
                string[] thoihan = {
                    "",
                    Clean(cbtime1.Text),
                    Clean(cbtime2.Text),
                    Clean(cbtime3.Text),
                    Clean(cbtime4.Text),
                    Clean(cbtime5.Text)
                };

                // Mảng lưu đối tượng của 5 tổ viên
                string[] doituong = {
                    "",
                    Clean(cbdt1.Text),
                    Clean(cbdt2.Text),
                    Clean(cbdt3.Text),
                    Clean(cbdt4.Text),
                    Clean(cbdt5.Text)
                };

                // Đếm số người đã nhập (phải có ít nhất 2 người)
                int soNguoi = Enumerable.Range(1, 5).Count(i => !string.IsNullOrWhiteSpace(hoten[i]));
                if (soNguoi < 2)
                {
                    MessageBox.Show("Phải có tối thiểu 2 người mới được tạo mẫu.", "Cảnh báo");
                    return;
                }

                // -------- TẠO DICTIONARY ÁNH XẠ PLACEHOLDER --------
                // Dictionary này lưu ánh xạ giữa placeholder trong file Word và giá trị thực tế
                // Ví dụ: {{totruong}} -> "Nguyễn Văn A"
                var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["{{totruong}}"] = totruong,
                    ["{{xa}}"] = xa,
                    ["{{thon1}}"] = thon,
                    ["{{chuongtrinh}}"] = chuongtrinh
                };

                // Thêm placeholder cho từng tổ viên (1-5)
                for (int i = 1; i <= 5; i++)
                {
                    map[$"{{{{hoten{i}}}}}"] = hoten[i];
                    map[$"{{{{sotien{i}}}}}"] = sotien[i];
                    map[$"{{{{phuongan{i}}}}}"] = phuongan[i];
                    map[$"{{{{thoihanvay{i}}}}}"] = thoihan[i];
                    map[$"{{{{doituong{i}}}}}"] = doituong[i];
                }

                // -------- THÊM NGƯỜI ĐÃ NHẬP VÀO DATAGRIDVIEW --------
                // Lặp qua 5 tổ viên và thêm vào danh sách selectedCustomers
                try
                {
                    for (int i = 1; i <=5; i++)
                    {
                        if (!string.IsNullOrWhiteSpace(hoten[i]))
                        {
                            // Tạo đối tượng Customer mới cho mỗi tổ viên
                            var c = new Customer
                            {
                                Hoten = hoten[i],
                                Totruong = totruong,
                                Xa = xa,
                                Chuongtrinh = chuongtrinh,
                                Sotien = sotien[i],
                                Phuongan = phuongan[i],
                                Thoihanvay = thoihan[i],
                                Doituong1 = doituong[i]
                            };
                            selectedCustomers.Add(c);
                        }
                    }
                    // Đã thêm vào selectedCustomers
                    // (dgv2 dùng để hiển thị lịch sử xuất)
                }
                catch { }

                // -------- XUẤT FILE WORD --------
                // Gọi phương thức ExportWord để tạo file Word từ mẫu
                await Task.Run(() => ExportWord(map, hoten, sotien, totruong, xa, chuongtrinh));

                // -------- LƯU LỊCH SỬ XUẤT VÀO DGV2 --------
                try
                {
                    if (isEditMode && currentEditingHistory != null)
                    {
                        // Chế độ Edit: Cập nhật thông tin tổ đã chọn (GHI ĐÈ)
                        currentEditingHistory.Xa = xa;
                        currentEditingHistory.Thon = thon;
                        currentEditingHistory.To = totruong;
                        currentEditingHistory.Chuongtrinh = chuongtrinh;
                        currentEditingHistory.Pgd = Clean(cbpgd2.Text);
                        currentEditingHistory.NgayXuat = DateTime.Now;
                        currentEditingHistory.SoThanhVien = soNguoi;

                        // Lưu thông tin 5 tổ viên
                        currentEditingHistory.Kh1 = hoten[1];
                        currentEditingHistory.Kh2 = hoten[2];
                        currentEditingHistory.Kh3 = hoten[3];
                        currentEditingHistory.Kh4 = hoten[4];
                        currentEditingHistory.Kh5 = hoten[5];

                        currentEditingHistory.Tien1 = sotien[1];
                        currentEditingHistory.Tien2 = sotien[2];
                        currentEditingHistory.Tien3 = sotien[3];
                        currentEditingHistory.Tien4 = sotien[4];
                        currentEditingHistory.Tien5 = sotien[5];

                        currentEditingHistory.Md1 = phuongan[1];
                        currentEditingHistory.Md2 = phuongan[2];
                        currentEditingHistory.Md3 = phuongan[3];
                        currentEditingHistory.Md4 = phuongan[4];
                        currentEditingHistory.Md5 = phuongan[5];

                        currentEditingHistory.Time1 = thoihan[1];
                        currentEditingHistory.Time2 = thoihan[2];
                        currentEditingHistory.Time3 = thoihan[3];
                        currentEditingHistory.Time4 = thoihan[4];
                        currentEditingHistory.Time5 = thoihan[5];

                        currentEditingHistory.Dt1 = doituong[1];
                        currentEditingHistory.Dt2 = doituong[2];
                        currentEditingHistory.Dt3 = doituong[3];
                        currentEditingHistory.Dt4 = doituong[4];
                        currentEditingHistory.Dt5 = doituong[5];

                        // Refresh DataGridView để hiển thị thông tin đã cập nhật
                        if (dgv2 != null)
                        {
                            dgv2.Refresh();
                        }

                        // Lưu vào file JSON
                        var toData = ConvertExportHistoryToToData(currentEditingHistory);
                        SaveToDataToFile(toData);

                        MessageBox.Show("Đã cập nhật thông tin tổ.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        // Chế độ Add: Tạo tổ mới
                        var history = new ExportHistory
                        {
                            Xa = xa,
                            Thon = thon,
                            To = totruong,
                            Chuongtrinh = chuongtrinh,
                            Pgd = Clean(cbpgd2.Text),
                            NgayXuat = DateTime.Now,
                            SoThanhVien = soNguoi,

                            // Lưu thông tin 5 tổ viên
                            Kh1 = hoten[1],
                            Kh2 = hoten[2],
                            Kh3 = hoten[3],
                            Kh4 = hoten[4],
                            Kh5 = hoten[5],

                            Tien1 = sotien[1],
                            Tien2 = sotien[2],
                            Tien3 = sotien[3],
                            Tien4 = sotien[4],
                            Tien5 = sotien[5],

                            Md1 = phuongan[1],
                            Md2 = phuongan[2],
                            Md3 = phuongan[3],
                            Md4 = phuongan[4],
                            Md5 = phuongan[5],

                            Time1 = thoihan[1],
                            Time2 = thoihan[2],
                            Time3 = thoihan[3],
                            Time4 = thoihan[4],
                            Time5 = thoihan[5],

                            Dt1 = doituong[1],
                            Dt2 = doituong[2],
                            Dt3 = doituong[3],
                            Dt4 = doituong[4],
                            Dt5 = doituong[5]
                        };
                        exportHistories.Add(history);

                        // Lưu vào file JSON
                        var toData = ConvertExportHistoryToToData(history);
                        SaveToDataToFile(toData);

                        MessageBox.Show("Đã tạo tổ mới và xuất file Word.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch { }

                // -------- LƯU TRẠNG THÁI FORM SAU KHI XUẤT THÀNH CÔNG --------
                // Lưu tất cả dữ liệu đã nhập vào file JSON
                SaveFormState();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xuất Word: " + ex.Message);
            }
        }

        // ============================================
        // LƯU & LOAD TRẠNG THÁI FORM
        // ============================================

        /// <summary>
        /// Lưu trạng thái hiện tại của form vào file JSON
        /// Tất cả dữ liệu đã nhập sẽ được lưu để khôi phục khi mở lại form
        /// File JSON sẽ được lưu tại thư mục gốc của ứng dụng
        /// </summary>
        private void SaveFormState()
        {
            try
            {
                // Tạo đối tượng Form2State chứa tất cả dữ liệu từ form
                var state = new Form2State
                {
                    // Thông tin chung
                    Totruong = cbTotruong.Text,
                    Xa = cbXa2.Text,
                    Thon = cbThon2.Text,
                    Pgd = cbpgd2.Text,
                    Chuongtrinh = cbctr.Text,

                    // Họ tên 5 tổ viên
                    Kh1 = txtkh1.Text,
                    Kh2 = txtkh2.Text,
                    Kh3 = txtkh3.Text,
                    Kh4 = txtkh4.Text,
                    Kh5 = txtkh5.Text,

                    // Số tiền 5 tổ viên
                    Tien1 = cbtien1.Text,
                    Tien2 = cbtien2.Text,
                    Tien3 = cbtien3.Text,
                    Tien4 = cbtien4.Text,
                    Tien5 = cbtien5.Text,

                    // Phương án 5 tổ viên
                    Md1 = cbmd1.Text,
                    Md2 = cbmd2.Text,
                    Md3 = cbmd3.Text,
                    Md4 = cbmd4.Text,
                    Md5 = cbmd5.Text,

                    // Thời hạn 5 tổ viên
                    Time1 = cbtime1.Text,
                    Time2 = cbtime2.Text,
                    Time3 = cbtime3.Text,
                    Time4 = cbtime4.Text,
                    Time5 = cbtime5.Text,

                    // Đối tượng 5 tổ viên
                    Dt1 = cbdt1.Text,
                    Dt2 = cbdt2.Text,
                    Dt3 = cbdt3.Text,
                    Dt4 = cbdt4.Text,
                    Dt5 = cbdt5.Text
                };

                // Chuyển đổi đối tượng state thành chuỗi JSON với format đẹp (Indented)
                var json = Newtonsoft.Json.JsonConvert.SerializeObject(state, Newtonsoft.Json.Formatting.Indented);

                // Ghi chuỗi JSON vào file với encoding UTF8
                File.WriteAllText(Form2StatePath, json, System.Text.Encoding.UTF8);

                // Hiển thị thông báo lưu thành công (không hiển thị đường dẫn)
                MessageBox.Show(
                    "Đã lưu thành công",
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                // Hiển thị thông báo lỗi nếu không lưu được (không hiển thị chi tiết kỹ thuật)
                MessageBox.Show(
                    "Lỗi khi lưu dữ liệu",
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        /// <summary>
        /// Khôi phục trạng thái form từ file JSON đã lưu
        /// Chỉ load khi form được mở độc lập (không phải từ Form1)
        /// </summary>
        private void LoadFormState()
        {
            try
            {
                // Chỉ load trạng thái nếu cờ shouldLoadState = true
                // (được mở độc lập, không phải từ Form1 với khách hàng đã chọn)
                if (!shouldLoadState) return;

                // Kiểm tra file JSON có tồn tại không
                if (!File.Exists(Form2StatePath)) return;

                // Đọc nội dung file JSON
                var json = File.ReadAllText(Form2StatePath, System.Text.Encoding.UTF8);

                // Chuyển đổi JSON thành đối tượng Form2State
                var state = Newtonsoft.Json.JsonConvert.DeserializeObject<Form2State>(json);
                if (state == null) return;

                // -------- KHÔI PHỤC THÔNG TIN CHUNG --------
                try { cbpgd2.Text = state.Pgd ?? ""; } catch { }
                try { cbTotruong.Text = state.Totruong ?? ""; } catch { }
                try { cbXa2.Text = state.Xa ?? ""; } catch { }
                try { cbThon2.Text = state.Thon ?? ""; } catch { }
                try { cbctr.Text = state.Chuongtrinh ?? ""; } catch { }

                // -------- KHÔI PHỤC HỌ TÊN 5 TỔ VIÊN --------
                try { txtkh1.Text = state.Kh1 ?? ""; } catch { }
                try { txtkh2.Text = state.Kh2 ?? ""; } catch { }
                try { txtkh3.Text = state.Kh3 ?? ""; } catch { }
                try { txtkh4.Text = state.Kh4 ?? ""; } catch { }
                try { txtkh5.Text = state.Kh5 ?? ""; } catch { }

                // -------- KHÔI PHỤC SỐ TIỀN 5 TỔ VIÊN --------
                try { cbtien1.Text = state.Tien1 ?? ""; } catch { }
                try { cbtien2.Text = state.Tien2 ?? ""; } catch { }
                try { cbtien3.Text = state.Tien3 ?? ""; } catch { }
                try { cbtien4.Text = state.Tien4 ?? ""; } catch { }
                try { cbtien5.Text = state.Tien5 ?? ""; } catch { }

                // -------- KHÔI PHỤC PHƯƠNG ÁN 5 TỔ VIÊN --------
                try { cbmd1.Text = state.Md1 ?? ""; } catch { }
                try { cbmd2.Text = state.Md2 ?? ""; } catch { }
                try { cbmd3.Text = state.Md3 ?? ""; } catch { }
                try { cbmd4.Text = state.Md4 ?? ""; } catch { }
                try { cbmd5.Text = state.Md5 ?? ""; } catch { }

                // -------- KHÔI PHỤC THỜI HẠN 5 TỔ VIÊN --------
                try { cbtime1.Text = state.Time1 ?? ""; } catch { }
                try { cbtime2.Text = state.Time2 ?? ""; } catch { }
                try { cbtime3.Text = state.Time3 ?? ""; } catch { }
                try { cbtime4.Text = state.Time4 ?? ""; } catch { }
                try { cbtime5.Text = state.Time5 ?? ""; } catch { }

                // -------- KHÔI PHỤC ĐỐI TƯỢNG 5 TỔ VIÊN --------
                try { cbdt1.Text = state.Dt1 ?? ""; } catch { }
                try { cbdt2.Text = state.Dt2 ?? ""; } catch { }
                try { cbdt3.Text = state.Dt3 ?? ""; } catch { }
                try { cbdt4.Text = state.Dt4 ?? ""; } catch { }
                try { cbdt5.Text = state.Dt5 ?? ""; } catch { }
            }
            catch (Exception ex)
            {
                // Nếu có lỗi khi load, im lặng bỏ qua (không hiển thị thông báo)
                // để không làm phiền người dùng khi mở form
            }
        }

        // ============================================
        // XUẤT FILE WORD TỪ MẪU
        // ============================================

        /// <summary>
        /// Xuất file Word từ mẫu (template)
        /// Tự động tìm file mẫu dựa trên chương trình được chọn
        /// </summary>
        /// <param name="map">Dictionary chứa ánh xạ placeholder -> giá trị thực tế</param>
        /// <param name="hoten">Mảng chứa họ tên 5 tổ viên</param>
        /// <param name="sotien">Mảng chứa số tiền 5 tổ viên</param>
        /// <param name="totruong">Tên tổ trưởng</param>
        /// <param name="xa">Tên xã</param>
        /// <param name="chuongtrinh">Tên chương trình</param>
        private void ExportWord(Dictionary<string, string> map, string[] hoten, string[] sotien, string totruong, string xa, string chuongtrinh)
        {
            // -------- XÁC ĐỊNH TÊN FILE MẪU --------
            // Form2 LUÔN dùng mẫu "03 DS GROUP.docx" cho xuất nhóm
            string templateName = "03 DS GROUP.docx";

            // -------- TÌM FILE MẪU TRONG THƯ MỤC TEMPLATES --------
            // Thử tìm trong thư mục Templates bên dưới thư mục chạy của ứng dụng
            string template = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Templates",
                templateName
            );

            // Nếu không tìm thấy trong Templates, thử tìm đệ quy trong toàn bộ thư mục gốc
            if (!File.Exists(template))
            {
                try
                {
                    var asm = Assembly.GetExecutingAssembly();
                    // Tìm đệ quy tất cả các thư mục con
                    var found = Directory.EnumerateFiles(AppDomain.CurrentDomain.BaseDirectory, templateName, SearchOption.AllDirectories).FirstOrDefault();
                    if (!string.IsNullOrEmpty(found)) template = found;

                    // Nếu vẫn không tìm thấy, thử tìm lên các thư mục cha (tối đa 6 cấp)
                    // Vì thư mục bin có thể nằm sâu dưới thư mục gốc dự án
                    if (!File.Exists(template))
                    {
                        var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
                        for (int up = 0; up < 6 && dir.Parent != null; up++)
                        {
                            dir = dir.Parent;
                            var upCandidate = Path.Combine(dir.FullName, templateName);
                            if (File.Exists(upCandidate)) { template = upCandidate; break; }
                        }
                    }

                    // Nếu vẫn không tìm thấy, thử tìm trong embedded resources (resources nhúng trong file .exe)
                    if (!File.Exists(template))
                    {
                        var res = asm.GetManifestResourceNames().FirstOrDefault(n => n.EndsWith(templateName, StringComparison.OrdinalIgnoreCase));
                        if (res != null)
                        {
                            // Tạo file tạm thời từ embedded resource
                            var temp = Path.Combine(Path.GetTempPath(), "template_" + Path.GetFileNameWithoutExtension(templateName) + "_" + Guid.NewGuid().ToString("N") + ".docx");
                            using (var s = asm.GetManifestResourceStream(res))
                            {
                                if (s != null) using (var fs = File.OpenWrite(temp)) s.CopyTo(fs);
                            }
                            if (File.Exists(temp)) template = temp;
                        }
                    }
                }
                catch { }
            }

            // -------- NẾU KHÔNG TÌM THẤY, YÊU CẦU NGƯỜI DÙNG CHỎN FILE --------
            // Hiển thị hộp thoại chọn file (phải chạy trên UI thread)
            if (!File.Exists(template))
            {
                try
                {
                    string userPick = null;
                    this.Invoke((Action)(() =>
                    {
                        using (var dlg = new OpenFileDialog())
                        {
                            dlg.Filter = "Word documents (*.docx)|*.docx|All files (*.*)|*.*";
                            dlg.Title = "Chọn file mẫu " + templateName;
                            if (dlg.ShowDialog(this) == DialogResult.OK) userPick = dlg.FileName;
                        }
                    }));

                    if (!string.IsNullOrEmpty(userPick) && File.Exists(userPick)) template = userPick;
                }
                catch { }
            }

            // Nếu vẫn không tìm thấy file mẫu, thông báo lỗi
            if (!File.Exists(template))
                throw new FileNotFoundException("Không tìm thấy file mẫu Word.");

            // -------- TẠO THƯ MỤC VÀ COPY FILE MẪU --------
            // Tạo thư mục trên Desktop: "Hồ sơ NHCS\[Tên tổ trưởng]"
            string folder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "Hồ sơ NHCS",
                Safe(totruong)
            );
            Directory.CreateDirectory(folder);

            // Tạo tên file xuất: "[Tổ trưởng]-[Xã]-[Tháng-Năm].docx"
            string output = Path.Combine(folder, Safe(totruong) + "-" + Safe(xa) + "-" + DateTime.Now.ToString("MM-yyyy") + ".docx");

            // Copy file mẫu đến đường dẫn xuất (ghi đè nếu đã tồn tại)
            File.Copy(template, output, true);

            // -------- XỬ LÝ FILE WORD --------
            // Mở file Word vừa tạo để thay thế placeholder
            using (var doc = WordprocessingDocument.Open(output, true))
            {
                // Bước 1: Xóa các dòng liên quan đến tổ viên trống trước
                // Để tránh để lại dòng trống trong bảng
                RemoveUnusedRows(doc, hoten);

                // Bước 2: Thay thế placeholder bị chia cắt qua nhiều run (text nodes)
                // Ví dụ: {{ho|ten1}} nếu bị chia thành 2 run
                ReplacePlaceholdersAcrossRuns(doc, map);

                // Bước 3: Thay thế placeholder đơn giản (nằm trọn trong 1 Text node)
                // Đây là phương án dự phòng cho những placeholder chưa được thay thế
                ReplacePlaceholdersPreserveFormatting(doc, map);

                // Bước 4: Tính và điền tổng số tiền từ các placeholder {{sotien}}
                FillTongTien(doc, sotien);

                // Lưu các thay đổi vào file
                doc.MainDocumentPart.Document.Save();
            }

            // -------- MỞ FILE WORD VỮA TẠO --------
            // Mở file Word bằng ứng dụng mặc định
            System.Diagnostics.Process.Start(output);
        }

        // ============================================
        // THAY THẾ PLACEHOLDER TRONG WORD
        // ============================================

        /// <summary>
        /// Thay thế các placeholder giữ nguyên định dạng
        /// Lặp qua các Text elements trong main và các phần liên quan
        /// Dùng cho placeholder nằm trọn trong 1 Text node
        /// </summary>
        /// <param name="doc">Document Word đang xử lý</param>
        /// <param name="map">Dictionary chứa ánh xạ placeholder -> giá trị</param>
        private void ReplacePlaceholdersPreserveFormatting(WordprocessingDocument doc, Dictionary<string, string> map)
        {
            if (doc?.MainDocumentPart == null) return;
            var mainPart = doc.MainDocumentPart;

            // Danh sách các phần cần thay thế placeholder
            var parts = new List<OpenXmlPart> { mainPart };
            parts.AddRange(mainPart.HeaderParts);        // Phần header
            parts.AddRange(mainPart.FooterParts);        // Phần footer
            if (mainPart.FootnotesPart != null) parts.Add(mainPart.FootnotesPart);    // Chú thích cuối trang
            if (mainPart.EndnotesPart != null) parts.Add(mainPart.EndnotesPart);      // Chú thích cuối tài liệu
            if (mainPart.WordprocessingCommentsPart != null) parts.Add(mainPart.WordprocessingCommentsPart); // Comment

            // Lặp qua từng phần của document
            foreach (var part in parts.Distinct())
            {
                try
                {
                    // Lấy tất cả các Text node trong phần này
                    var texts = part.RootElement.Descendants<Text>();
                    foreach (var t in texts)
                    {
                        if (string.IsNullOrEmpty(t.Text)) continue;
                        string newText = t.Text;

                        // Thay thế tất cả các placeholder trong text này
                        foreach (var kv in map)
                        {
                            if (string.IsNullOrEmpty(kv.Key)) continue;
                            if (newText.IndexOf(kv.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                newText = newText.Replace(kv.Key, kv.Value ?? "");
                            }
                        }

                        // Cập nhật text nếu có thay đổi
                        if (newText != t.Text) t.Text = newText;
                    }
                    // Lưu các thay đổi
                    part.RootElement.Save();
                }
                catch { }
            }
        }

        /// <summary>
        /// Thay thế các placeholder có thể bị chia cắt qua nhiều Text nodes (runs)
        /// Ví dụ: {{hoten1}} có thể bị chia thành "{{ho" và "ten1}}" trong 2 run khác nhau
        /// Phương thức này xử lý trường hợp đó
        /// </summary>
        /// <param name="doc">Document Word đang xử lý</param>
        /// <param name="map">Dictionary chứa ánh xạ placeholder -> giá trị</param>
        private void ReplacePlaceholdersAcrossRuns(WordprocessingDocument doc, Dictionary<string, string> map)
        {
            if (doc?.MainDocumentPart == null) return;

            // Danh sách các phần cần xử lý
            var parts = new List<OpenXmlPart> { doc.MainDocumentPart };
            parts.AddRange(doc.MainDocumentPart.HeaderParts);
            parts.AddRange(doc.MainDocumentPart.FooterParts);
            if (doc.MainDocumentPart.FootnotesPart != null) parts.Add(doc.MainDocumentPart.FootnotesPart);
            if (doc.MainDocumentPart.EndnotesPart != null) parts.Add(doc.MainDocumentPart.EndnotesPart);
            if (doc.MainDocumentPart.WordprocessingCommentsPart != null) parts.Add(doc.MainDocumentPart.WordprocessingCommentsPart);

            // Sắp xếp các placeholder theo độ dài giảm dần để tránh khớp một phần
            // Ví dụ: thay {{hoten1}} trước {{hoten}} để tránh nhầm lẫn
            var placeholders = map.Keys.Where(k => !string.IsNullOrEmpty(k)).OrderByDescending(k => k.Length).ToList();

            // Lặp qua từng phần của document
            foreach (var part in parts.Distinct())
            {
                try
                {
                    // Nhóm các Text nodes theo paragraph chứa chúng
                    // Để giữ thay thế trong phạm vi local (không lẫn giữa các đoạn)
                    var paragraphs = part.RootElement.Descendants<Paragraph>().ToList();
                    bool partChanged = false;

                    foreach (var p in paragraphs)
                    {
                        var texts = p.Descendants<Text>().ToList();
                        if (texts.Count == 0) continue;

                        // Nối tất cả text trong paragraph thành 1 chuỗi
                        string full = string.Concat(texts.Select(t => t.Text ?? ""));
                        if (string.IsNullOrEmpty(full)) continue;

                        bool changed = false;

                        foreach (var ph in placeholders)
                        {
                            var replacement = map.ContainsKey(ph) ? map[ph] ?? "" : "";
                            int searchStart = 0;
                            while (true)
                            {
                                int idx = full.IndexOf(ph, searchStart, StringComparison.OrdinalIgnoreCase);
                                if (idx < 0) break;

                                // locate start text node
                                int cum = 0; int startNode = -1; int startOffset = 0;
                                for (int i = 0; i < texts.Count; i++)
                                {
                                    var len = (texts[i].Text ?? "").Length;
                                    if (idx < cum + len)
                                    {
                                        startNode = i;
                                        startOffset = idx - cum;
                                        break;
                                    }
                                    cum += len;
                                }
                                if (startNode < 0) break;

                                // Tính vị trí kết thúc của placeholder (exclusive)
                                int matchEndPos = idx + ph.Length;

                                // Tìm Text node chứa vị trí kết thúc của placeholder
                                cum = 0; int endNode = -1; int endOffsetExclusive = 0;
                                for (int i = 0; i < texts.Count; i++)
                                {
                                    var len = (texts[i].Text ?? "").Length;
                                    if (matchEndPos <= cum + len)
                                    {
                                        endNode = i;
                                        endOffsetExclusive = matchEndPos - cum;
                                        break;
                                    }
                                    cum += len;
                                }
                                if (endNode < 0) break;

                                // Tính đoạn trái (phía trước placeholder) và đoạn phải (phía sau placeholder)
                                var left = (texts[startNode].Text ?? "").Substring(0, startOffset);
                                var right = (texts[endNode].Text ?? "").Substring(endOffsetExclusive);

                                // Gán giá trị mới cho node bắt đầu: trái + thay thế + phải
                                texts[startNode].Text = left + (replacement ?? "") + right;

                                // Xóa các node trung gian (nếu placeholder không nằm trong một node duy nhất)
                                for (int k = startNode + 1; k <= endNode; k++)
                                {
                                    if (k == startNode + 1)
                                    {
                                        // Nếu endNode == startNode+1 và phần phải đã chuyển vào node bắt đầu, xóa node này
                                        texts[k].Text = "";
                                    }
                                    else
                                    {
                                        texts[k].Text = "";
                                    }
                                }

                                changed = true;

                                // Xây dựng lại chuỗi full và tiếp tục tìm sau vị trí thay thế
                                full = string.Concat(texts.Select(t => t.Text ?? ""));
                                searchStart = (left + replacement).Length + texts.Take(startNode).Sum(t => (t.Text ?? "").Length);
                            }
                        }

                        if (changed) partChanged = true;
                    }
                    // Nếu có thay đổi, lưu lại phần này
                    if (partChanged)
                    {
                        try { part.RootElement.Save(); } catch { }
                    }
                }
                catch { }
            }
        }

        // ============================================
        // XÓA DÒNG TRỐNG TRONG BẢNG WORD
        // ============================================

        /// <summary>
        /// Xóa các dòng trong bảng Word chứa placeholder của tổ viên không có tên
        /// Nếu tổ viên thứ i không có tên, dòng chứa {{hoten{i}}} sẽ bị xóa
        /// </summary>
        /// <param name="doc">Document Word đang xử lý</param>
        /// <param name="hoten">Mảng chứa họ tên 5 tổ viên</param>
        private void RemoveUnusedRows(WordprocessingDocument doc, string[] hoten)
        {
            if (doc?.MainDocumentPart == null) return;

            // Lấy tất cả các dòng trong bảng
            var rows = doc.MainDocumentPart.Document
                .Descendants<TableRow>()
                .ToList();

            // Lặp qua từng dòng
            foreach (var row in rows)
            {
                // Nối tất cả text trong dòng thành 1 chuỗi
                var rowText = string.Concat(row.Descendants<Text>().Select(t => t.Text ?? ""));

                // Kiểm tra từng tổ viên (1-5)
                for (int i = 1; i <= 5; i++)
                {
                    // Nếu tổ viên i không có tên và dòng chứa placeholder {{hoten{i}}}
                    if (string.IsNullOrWhiteSpace(hoten[i]) && rowText.IndexOf($"{{{{hoten{i}}}}}", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Xóa dòng này và tiếp tục dòng tiếp theo
                        row.Remove();
                        break;
                    }
                }
            }
        }

        // ============================================
        // TÍNH VÀ ĐIỀN TỔNG SỐ TIỀN
        // ============================================

        /// <summary>
        /// Tính tổng số tiền từ các tổ viên và điền vào placeholder {{cong}}
        /// </summary>
        /// <param name="doc">Document Word đang xử lý</param>
        /// <param name="sotien">Mảng chứa số tiền của 5 tổ viên</param>
        private void FillTongTien(WordprocessingDocument doc, string[] sotien)
        {
            long tong = 0;

            // Cộng dồn số tiền của 5 tổ viên
            for (int i = 1; i <= 5; i++)
            {
                // Lọc chỉ lấy số (loại bỏ dấu chấm phân cách)
                var digits = new string((sotien[i] ?? "").Where(char.IsDigit).ToArray());
                if (long.TryParse(digits, out long v)) tong += v;
            }

            // Thay thế placeholder {{cong}} với tổng số tiền (đã format)
            ReplacePlaceholdersPreserveFormatting(doc, new Dictionary<string, string>
            {
                ["{{cong}}"] = tong > 0 ? tong.ToString("N0", new CultureInfo("vi-VN")) : ""
            });
        }

        // ============================================
        // CÁC PHƯƠNG THỨC TIỆN ÍCH
        // ============================================

        /// <summary>
        /// Làm sạch chuỗi: loại bỏ khoảng trắng thừa và ký tự đặc biệt { }
        /// </summary>
        /// <param name="s">Chuỗi cần làm sạch</param>
        /// <returns>Chuỗi đã làm sạch</returns>
        private string Clean(string s)
        {
            return (s ?? "").Trim().Replace("{", "").Replace("}", "");
        }

        /// <summary>
        /// Định dạng số tiền với dấu chấm phân cách hàng nghìn
        /// Ví dụ: "1000000" -> "1.000.000"
        /// </summary>
        /// <param name="s">Chuỗi số tiền cần định dạng</param>
        /// <returns>Chuỗi đã định dạng</returns>
        private string Money(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            // Lọc chỉ lấy số
            var d = new string(s.Where(char.IsDigit).ToArray());
            if (long.TryParse(d, out long v))
                return v.ToString("N0", new CultureInfo("vi-VN"));
            return s;
        }

        // ============================================
        // XỬ LÝ NHẬP LIỆU (INPUT HANDLERS)
        // ============================================

        /// <summary>
        /// Xử lý sự kiện KeyPress cho các ô tiền (ComboBox)
        /// Chỉ cho phép nhập số và các phím điều khiển (Backspace, Delete, ...)
        /// </summary>
        private void CbMoney_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Cho phép các phím điều khiển (Ctrl+C, Backspace, Delete, ...)
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                // Chặn ký tự không phải là số
                e.Handled = true;
            }
        }

        /// <summary>
        /// Xử lý sự kiện TextChanged cho các ô tiền (ComboBox)
        /// Tự động định dạng số tiền với dấu chấm phân cách hàng nghìn
        /// </summary>
        private void CbMoney_TextChanged(object sender, EventArgs e)
        {
            try
            {
                var cb = sender as ComboBox;
                if (cb == null) return;
                var txt = cb.Text ?? "";

                // Lọc chỉ lấy số
                var digits = new string(txt.Where(char.IsDigit).ToArray());
                if (string.IsNullOrEmpty(digits))
                {
                    if (!string.IsNullOrEmpty(txt)) cb.Text = "";
                    return;
                }

                // Định dạng số với dấu chấm phân cách hàng nghìn
                if (long.TryParse(digits, out var v))
                {
                    var formatted = v.ToString("N0", new CultureInfo("vi-VN"));
                    if (cb.Text != formatted)
                    {
                        cb.Text = formatted;
                        // Giữ con trỏ ở cuối chuỗi
                        try { cb.SelectionStart = cb.Text.Length; cb.SelectionLength = 0; } catch { }
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Xử lý sự kiện KeyPress cho các ô tên và địa điểm
        /// Chỉ cho phép nhập chữ cái, khoảng trắng và một số ký tự đặc biệt (-,',.)
        /// Không cho phép nhập số
        /// </summary>
        private void TextLettersOnly_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Cho phép các phím điều khiển (Backspace, Delete, ...)
            if (char.IsControl(e.KeyChar)) return;
            // Cho phép khoảng trắng
            if (char.IsWhiteSpace(e.KeyChar)) return;
            // Cho phép chữ cái
            if (char.IsLetter(e.KeyChar)) return;
            // Cho phép một số ký tự đặc biệt trong tên
            if (e.KeyChar == '-' || e.KeyChar == '\'' || e.KeyChar == '.' ) return;
            // Chặn tất cả các ký tự khác (bao gồm số)
            e.Handled = true;
        }

        // ============================================
        // STUB HANDLERS (CHO DESIGNER)
        // Các phương thức trống để thỏa mãn designer, không cần thực hiện gì
        // ============================================

        /// <summary>
        /// Stub handler cho sự kiện TextChanged của textBox12 (txtxa)
        /// Để trống để thỏa mãn designer
        /// </summary>
        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            try { /* để trống cố ý để thỏa mãn designer */ } catch { }
        }

        /// <summary>
        /// Stub handler cho sự kiện Click của label18
        /// Để trống để thỏa mãn designer
        /// </summary>
        private void label18_Click(object sender, EventArgs e)
        {
            try { /* để trống cố ý để thỏa mãn designer */ } catch { }
        }

        // ============================================
        // XỬ LÝ NÚT XÓA
        // ============================================

        /// <summary>
        /// Xử lý sự kiện click nút "Xóa" (btnxoa)
        /// Xóa các dòng đã chọn trong DataGridView
        /// </summary>
        private void BtnXoa_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgv2 == null) return;

                // Xóa các dòng lịch sử xuất đã chọn
                var toRemove = new List<ExportHistory>();
                foreach (DataGridViewRow row in dgv2.SelectedRows)
                {
                    try { var item = row.DataBoundItem as ExportHistory; if (item != null) toRemove.Add(item); } catch { }
                }

                // Xóa từng lịch sử khỏi danh sách và file JSON
                foreach (var r in toRemove)
                {
                    // Chuyển đổi thành ToData để xóa file
                    var toData = ConvertExportHistoryToToData(r);

                    // Xóa file JSON
                    DeleteToDataFile(toData);

                    // Xóa khỏi danh sách hiển thị
                    exportHistories.Remove(r);
                }

                // Reset chế độ Edit nếu đang edit tổ bị xóa
                if (isEditMode && currentEditingHistory != null && toRemove.Contains(currentEditingHistory))
                {
                    isEditMode = false;
                    currentEditingHistory = null;
                    try { btn03to.Text = "Xuất tổ"; } catch { }
                }
            }
            catch { }
        }

        // ============================================
        // XỬ LÝ CLICK VÀO DATAGRIDVIEW
        // ============================================

        /// <summary>
        /// Xử lý sự kiện click vào ô trong DataGridView
        /// Load dữ liệu khách hàng từ dòng được chọn lên form để chỉnh sửa
        /// </summary>
        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                // DataGridView1 không còn tồn tại
                // dgv2 chỉ hiển thị lịch sử xuất, không cần load dữ liệu lên form
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi load dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Chuyển đổi chuỗi thành tên file an toàn
        /// Thay thế các ký tự không hợp lệ trong tên file bằng dấu gạch dưới '_'
        /// </summary>
        /// <param name="s">Chuỗi cần chuyển đổi</param>
        /// <returns>Chuỗi an toàn cho tên file</returns>
        private string Safe(string s)
        {
            // Thay thế tất cả ký tự không hợp lệ trong tên file
            foreach (var c in Path.GetInvalidFileNameChars())
                s = s.Replace(c, '_');
            return s.Trim();
        }

        /// <summary>
        /// Xử lý sự kiện Load của Form2
        /// Load trạng thái đã lưu (nếu mở độc lập, không phải từ Form1)
        /// </summary>
        private void Form2_Load(object sender, EventArgs e)
        {
            // Load dữ liệu xinman.json
            LoadXinManData();

            // Load danh sách tổ đã lưu từ file JSON
            LoadToFromFiles();

            // Load trạng thái đã lưu khi form load
            LoadFormState();
        }

        // ============================================
        // LOAD DỮ LIỆU XINMAN.JSON
        // ============================================

        /// <summary>
        /// Load dữ liệu từ tất cả các file JSON có sẵn (xinman.json, dongvan.json, meovac.json, vixuyen.json, v.v.)
        /// </summary>
        private void LoadXinManData()
        {
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;

                // Tìm tất cả các file *.json trong thư mục
                var allJsonFiles = Directory.GetFiles(baseDir, "*.json", SearchOption.TopDirectoryOnly);

                // Lọc bỏ các file không phải dữ liệu xinman (như Form2State.json, Customers\*.json)
                var xinmanFiles = allJsonFiles
                    .Where(f => !Path.GetFileName(f).Equals("Form2State.json", StringComparison.OrdinalIgnoreCase))
                    .Where(f => !f.Contains("\\Customers\\"))
                    .ToList();

                if (xinmanFiles.Count == 0) return;

                xinmanModels.Clear();
                if (cbpgd2 != null) cbpgd2.Items.Clear();

                // Load từng file JSON
                foreach (var filePath in xinmanFiles)
                {
                    try
                    {
                        string json = File.ReadAllText(filePath, System.Text.Encoding.UTF8);
                        var model = Newtonsoft.Json.JsonConvert.DeserializeObject<XinManModel>(json);

                        if (model != null && !string.IsNullOrWhiteSpace(model.pgd))
                        {
                            // Lưu model vào dictionary
                            xinmanModels[model.pgd] = model;

                            // Thêm PGD vào cbpgd2
                            if (cbpgd2 != null && !cbpgd2.Items.Contains(model.pgd))
                            {
                                cbpgd2.Items.Add(model.pgd);
                            }
                        }
                    }
                    catch { }
                }

                // Chọn PGD đầu tiên nếu có
                if (cbpgd2 != null && cbpgd2.Items.Count > 0)
                {
                    cbpgd2.SelectedIndex = 0;
                }
            }
            catch { }
        }

        /// <summary>
        /// Xử lý sự kiện SelectedIndexChanged của cbpgd2
        /// Load dữ liệu Xã, Thôn, Tổ tương ứng với PGD được chọn
        /// </summary>
        private void CbPgd2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cbpgd2 == null || string.IsNullOrWhiteSpace(cbpgd2.Text)) return;

                string selectedPgd = cbpgd2.Text.Trim();

                // Tìm model tương ứng với PGD được chọn
                if (!xinmanModels.TryGetValue(selectedPgd, out xinmanModel))
                {
                    return;
                }

                if (xinmanModel == null || xinmanModel.communes == null) return;

                // Xóa dữ liệu cũ
                if (cbXa2 != null) { cbXa2.Items.Clear(); cbXa2.Text = ""; }
                if (cbThon2 != null) { cbThon2.Items.Clear(); cbThon2.Text = ""; }
                if (cbTotruong != null) { cbTotruong.Items.Clear(); cbTotruong.Text = ""; }

                // Load danh sách Xã
                foreach (var commune in xinmanModel.communes)
                {
                    if (!string.IsNullOrWhiteSpace(commune.name) && cbXa2 != null)
                    {
                        cbXa2.Items.Add(commune.name);
                    }
                }

                // Đăng ký sự kiện cho cbXa2 (chỉ đăng ký một lần)
                if (cbXa2 != null)
                {
                    cbXa2.SelectedIndexChanged -= CbXa2_SelectedIndexChanged;
                    cbXa2.SelectedIndexChanged += CbXa2_SelectedIndexChanged;
                }

                // Đăng ký sự kiện cho cbThon2 (chỉ đăng ký một lần)
                if (cbThon2 != null)
                {
                    cbThon2.SelectedIndexChanged -= CbThon2_SelectedIndexChanged;
                    cbThon2.SelectedIndexChanged += CbThon2_SelectedIndexChanged;
                }
            }
            catch { }
        }

        /// <summary>
        /// Xử lý sự kiện SelectedIndexChanged của cbXa2
        /// Load danh sách Thôn và Tổ trưởng tương ứng với Xã được chọn
        /// </summary>
        private void CbXa2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (xinmanModel == null || xinmanModel.communes == null) return;
                if (cbXa2 == null || cbThon2 == null) return;

                string selectedXa = cbXa2.Text?.Trim() ?? "";
                if (string.IsNullOrWhiteSpace(selectedXa)) return;

                // Tìm xã được chọn
                var commune = xinmanModel.communes.FirstOrDefault(c =>
                    string.Equals(c.name?.Trim(), selectedXa, StringComparison.OrdinalIgnoreCase));

                if (commune == null) return;

                // Xóa dữ liệu cũ
                cbThon2.Items.Clear();
                cbThon2.Text = "";
                if (cbTotruong != null) { cbTotruong.Items.Clear(); cbTotruong.Text = ""; }

                // Thu thập tất cả Thôn và Tổ từ xã được chọn
                var allVillages = new List<Village>();
                var allGroups = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // 1. Lấy Thôn từ commune.villages (nếu có)
                if (commune.villages != null)
                {
                    foreach (var village in commune.villages)
                    {
                        if (!string.IsNullOrWhiteSpace(village.name))
                        {
                            allVillages.Add(village);

                            // Lấy tất cả groups từ village này
                            if (village.groups != null)
                            {
                                foreach (var group in village.groups)
                                {
                                    if (!string.IsNullOrWhiteSpace(group))
                                        allGroups.Add(group);
                                }
                            }
                        }
                    }
                }

                // 2. Lấy Thôn từ commune.associations[].villages (nếu có)
                if (commune.associations != null)
                {
                    foreach (var association in commune.associations)
                    {
                        if (association.villages != null)
                        {
                            foreach (var village in association.villages)
                            {
                                if (!string.IsNullOrWhiteSpace(village.name))
                                {
                                    // Kiểm tra xem village đã có chưa (theo tên)
                                    if (!allVillages.Any(v => string.Equals(v.name?.Trim(), village.name?.Trim(), StringComparison.OrdinalIgnoreCase)))
                                    {
                                        allVillages.Add(village);
                                    }

                                    // Lấy tất cả groups từ village này
                                    if (village.groups != null)
                                    {
                                        foreach (var group in village.groups)
                                        {
                                            if (!string.IsNullOrWhiteSpace(group))
                                                allGroups.Add(group);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // Load danh sách Thôn vào cbThon2
                foreach (var village in allVillages.OrderBy(v => v.name))
                {
                    cbThon2.Items.Add(village.name);
                }

                // Load danh sách Tổ trưởng vào cbTotruong
                if (cbTotruong != null)
                {
                    foreach (var group in allGroups.OrderBy(g => g))
                    {
                        cbTotruong.Items.Add(group);
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Xử lý sự kiện SelectedIndexChanged của cbThon2
        /// Load danh sách Tổ tương ứng với Thôn được chọn
        /// </summary>
        private void CbThon2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                // Bỏ qua nếu đang cập nhật programmatically
                if (isUpdatingThonTotruong) return;

                if (xinmanModel == null || xinmanModel.communes == null) return;
                if (cbXa2 == null || cbThon2 == null || cbTotruong == null) return;

                string selectedXa = cbXa2.Text?.Trim() ?? "";
                string selectedThon = cbThon2.Text?.Trim() ?? "";

                if (string.IsNullOrWhiteSpace(selectedXa) || string.IsNullOrWhiteSpace(selectedThon)) return;

                // Tìm xã được chọn
                var commune = xinmanModel.communes.FirstOrDefault(c =>
                    string.Equals(c.name?.Trim(), selectedXa, StringComparison.OrdinalIgnoreCase));

                if (commune == null) return;

                // Tìm thôn được chọn - Tìm trong tất cả villages (bao gồm cả trong associations)
                Village village = null;

                // 1. Tìm trong commune.villages
                if (commune.villages != null)
                {
                    village = commune.villages.FirstOrDefault(v =>
                        string.Equals(v.name?.Trim(), selectedThon, StringComparison.OrdinalIgnoreCase));
                }

                // 2. Nếu không tìm thấy, tìm trong associations
                if (village == null && commune.associations != null)
                {
                    foreach (var association in commune.associations)
                    {
                        if (association.villages != null)
                        {
                            village = association.villages.FirstOrDefault(v =>
                                string.Equals(v.name?.Trim(), selectedThon, StringComparison.OrdinalIgnoreCase));
                            if (village != null) break;
                        }
                    }
                }

                if (village == null) return;

                // Set flag để tránh trigger event của cbTotruong
                isUpdatingThonTotruong = true;

                try
                {
                    // Xóa dữ liệu cũ
                    cbTotruong.Items.Clear();
                    cbTotruong.Text = "";

                    // Load danh sách Tổ trưởng CHỈ của thôn được chọn
                    if (village.groups != null)
                    {
                        foreach (var group in village.groups)
                        {
                            if (!string.IsNullOrWhiteSpace(group))
                            {
                                cbTotruong.Items.Add(group);
                            }
                        }
                    }
                }
                finally
                {
                    isUpdatingThonTotruong = false;
                }
            }
            catch { }
        }

        /// <summary>
        /// Xử lý sự kiện SelectedIndexChanged của cbTotruong
        /// Tự động chọn thôn tương ứng với tổ trưởng được chọn
        /// </summary>
        private void CbTotruong_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                // Bỏ qua nếu đang cập nhật programmatically
                if (isUpdatingThonTotruong) return;

                if (xinmanModel == null || xinmanModel.communes == null) return;
                if (cbXa2 == null || cbThon2 == null || cbTotruong == null) return;

                string selectedXa = cbXa2.Text?.Trim() ?? "";
                string selectedTotruong = cbTotruong.Text?.Trim() ?? "";

                if (string.IsNullOrWhiteSpace(selectedXa) || string.IsNullOrWhiteSpace(selectedTotruong)) return;

                // Tìm xã được chọn
                var commune = xinmanModel.communes.FirstOrDefault(c =>
                    string.Equals(c.name?.Trim(), selectedXa, StringComparison.OrdinalIgnoreCase));

                if (commune == null) return;

                // Tìm thôn chứa tổ trưởng này
                Village foundVillage = null;

                // 1. Tìm trong commune.villages
                if (commune.villages != null)
                {
                    foreach (var village in commune.villages)
                    {
                        if (village.groups != null && village.groups.Contains(selectedTotruong, StringComparer.OrdinalIgnoreCase))
                        {
                            foundVillage = village;
                            break;
                        }
                    }
                }

                // 2. Nếu không tìm thấy, tìm trong associations
                if (foundVillage == null && commune.associations != null)
                {
                    foreach (var association in commune.associations)
                    {
                        if (association.villages != null)
                        {
                            foreach (var village in association.villages)
                            {
                                if (village.groups != null && village.groups.Contains(selectedTotruong, StringComparer.OrdinalIgnoreCase))
                                {
                                    foundVillage = village;
                                    break;
                                }
                            }
                            if (foundVillage != null) break;
                        }
                    }
                }

                // Nếu tìm thấy thôn, tự động chọn thôn đó trong cbThon2
                if (foundVillage != null)
                {
                    // Set flag để tránh trigger event của cbThon2
                    isUpdatingThonTotruong = true;

                    try
                    {
                        // Tự động chọn thôn tương ứng
                        cbThon2.Text = foundVillage.name;
                    }
                    finally
                    {
                        isUpdatingThonTotruong = false;
                    }
                }
            }
            catch { }
        }

        // ============================================
        // XỬ LÝ CHỌN CHƯƠNG TRÌNH - CẬP NHẬT ĐỐI TƯỢNG
        // ============================================

        /// <summary>
        /// Xử lý sự kiện SelectedIndexChanged của cbctr
        /// Cập nhật danh sách đối tượng (cbdt1-5) dựa trên chương trình được chọn
        /// </summary>
        private void Cbctr_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cbctr == null) return;

                string selectedChuongTrinh = cbctr.Text?.Trim() ?? "";
                if (string.IsNullOrWhiteSpace(selectedChuongTrinh))
                {
                    // Nếu không chọn chương trình, reset cbdt1-5 về rỗng
                    ResetDoiTuongComboBoxes();
                    return;
                }

                // Xác định danh sách đối tượng dựa trên chương trình
                List<string> doiTuongList = new List<string>();

                // So sánh không phân biệt hoa thường và bỏ qua khoảng trắng
                if (selectedChuongTrinh.IndexOf("Hộ nghèo", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    selectedChuongTrinh.IndexOf("cận", StringComparison.OrdinalIgnoreCase) < 0 &&
                    selectedChuongTrinh.IndexOf("thoát", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    // 1. Hộ nghèo
                    doiTuongList.Add("Hộ nghèo");
                }
                else if (selectedChuongTrinh.IndexOf("cận nghèo", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // 2. Hộ cận nghèo
                    doiTuongList.Add("Hộ cận nghèo");
                }
                else if (selectedChuongTrinh.IndexOf("thoát nghèo", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // 3. Hộ mới thoát nghèo
                    doiTuongList.Add("Hộ mới thoát nghèo");
                }
                else if (selectedChuongTrinh.IndexOf("Sản xuất kinh doanh", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         selectedChuongTrinh.IndexOf("SXKD", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // 4. Hộ gia đình Sản xuất kinh doanh tại vùng khó khăn
                    doiTuongList.Add("Hộ GĐ SXKD VKK");
                }
                else if (selectedChuongTrinh.IndexOf("việc làm", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // 5. Hỗ trợ tạo việc làm duy trì và mở rộng việc làm
                    doiTuongList.Add("Người lao động");
                    doiTuongList.Add("NLĐ là người DTTS");
                }
                else if (selectedChuongTrinh.IndexOf("nước sạch", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         selectedChuongTrinh.IndexOf("vệ sinh môi trường", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // 6. Cấp nước sạch và vệ sinh môi trường nông thôn
                    doiTuongList.Add("HGĐ cư trú tại VNT");
                }

                // Cập nhật cbdt1-5 với danh sách mới
                UpdateDoiTuongComboBoxes(doiTuongList);
            }
            catch { }
        }

        /// <summary>
        /// Cập nhật danh sách Items của cbdt1-5 với danh sách đối tượng mới
        /// Chỉ giới hạn lựa chọn, không tự động chọn giá trị
        /// </summary>
        /// <param name="doiTuongList">Danh sách đối tượng cần hiển thị</param>
        private void UpdateDoiTuongComboBoxes(List<string> doiTuongList)
        {
            try
            {
                // Lưu giá trị hiện tại của cbdt1-5
                string[] currentValues = {
                    cbdt1?.Text ?? "",
                    cbdt2?.Text ?? "",
                    cbdt3?.Text ?? "",
                    cbdt4?.Text ?? "",
                    cbdt5?.Text ?? ""
                };

                // Cập nhật cbdt1
                if (cbdt1 != null)
                {
                    cbdt1.Items.Clear();
                    foreach (var item in doiTuongList)
                        cbdt1.Items.Add(item);

                    // Chỉ giữ giá trị cũ nếu nó còn trong danh sách mới
                    if (doiTuongList.Contains(currentValues[0]))
                        cbdt1.Text = currentValues[0];
                    else
                        cbdt1.Text = ""; // Không tự động chọn
                }

                // Cập nhật cbdt2
                if (cbdt2 != null)
                {
                    cbdt2.Items.Clear();
                    foreach (var item in doiTuongList)
                        cbdt2.Items.Add(item);

                    if (doiTuongList.Contains(currentValues[1]))
                        cbdt2.Text = currentValues[1];
                    else
                        cbdt2.Text = "";
                }

                // Cập nhật cbdt3
                if (cbdt3 != null)
                {
                    cbdt3.Items.Clear();
                    foreach (var item in doiTuongList)
                        cbdt3.Items.Add(item);

                    if (doiTuongList.Contains(currentValues[2]))
                        cbdt3.Text = currentValues[2];
                    else
                        cbdt3.Text = "";
                }

                // Cập nhật cbdt4
                if (cbdt4 != null)
                {
                    cbdt4.Items.Clear();
                    foreach (var item in doiTuongList)
                        cbdt4.Items.Add(item);

                    if (doiTuongList.Contains(currentValues[3]))
                        cbdt4.Text = currentValues[3];
                    else
                        cbdt4.Text = "";
                }

                // Cập nhật cbdt5
                if (cbdt5 != null)
                {
                    cbdt5.Items.Clear();
                    foreach (var item in doiTuongList)
                        cbdt5.Items.Add(item);

                    if (doiTuongList.Contains(currentValues[4]))
                        cbdt5.Text = currentValues[4];
                    else
                        cbdt5.Text = "";
                }
            }
            catch { }
        }

        /// <summary>
        /// Reset cbdt1-5 về trạng thái rỗng
        /// </summary>
        private void ResetDoiTuongComboBoxes()
        {
            try
            {
                if (cbdt1 != null) { cbdt1.Items.Clear(); cbdt1.Text = ""; }
                if (cbdt2 != null) { cbdt2.Items.Clear(); cbdt2.Text = ""; }
                if (cbdt3 != null) { cbdt3.Items.Clear(); cbdt3.Text = ""; }
                if (cbdt4 != null) { cbdt4.Items.Clear(); cbdt4.Text = ""; }
                if (cbdt5 != null) { cbdt5.Items.Clear(); cbdt5.Text = ""; }
            }
            catch { }
        }

        // ============================================
        // HIỆU ỨNG CHỮ CHẠY (MARQUEE) CHO RICHTEXTBOX
        // ============================================

        /// <summary>
        /// Khởi tạo hiệu ứng chữ chạy cho richTextBox1
        /// Tạo hiệu ứng chuỗi chữ di chuyển từ phải sang trái
        /// </summary>
        private void InitializeRichTextMarquee()
        {
            try
            {
                if (richTextBox1 == null) return;

                // Lấy văn bản ban đầu hoặc đặt văn bản mặc định
                string initialText = richTextBox1.Text;
                if (string.IsNullOrWhiteSpace(initialText))
                {
                    initialText = "DANH SÁCH KHÁCH HÀNG XUẤT HỒ SƠ NHÓM - NHẬP THÔNG TIN VÀ BẤM EXPORT";
                    richTextBox1.Text = initialText;
                }

                // Thêm khoảng trắng để tạo hiệu ứng chạy mượt mà
                richTextMarqueeText = initialText + "     ";
                richTextMarqueePosition = 0;

                // Tạo và khởi động Timer
                var marqueeTimer = new Timer();
                marqueeTimer.Interval = 150;  // Chậm hơn label (150ms so với 100ms)
                marqueeTimer.Tick += RichTextMarqueeTimer_Tick;
                marqueeTimer.Start();
            }
            catch { }
        }

        /// <summary>
        /// Xử lý sự kiện Tick của Timer cho hiệu ứng chữ chạy
        /// Cập nhật vị trí và hiển thị chuỗi chữ di chuyển
        /// </summary>
        private void RichTextMarqueeTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (richTextBox1 == null || string.IsNullOrEmpty(richTextMarqueeText)) return;

                // Tăng vị trí hiện tại
                richTextMarqueePosition++;
                if (richTextMarqueePosition >= richTextMarqueeText.Length)
                {
                    // Quay về đầu khi hết chuỗi
                    richTextMarqueePosition = 0;
                }

                // Tạo chuỗi hiển thị bằng cách xoay vòng văn bản
                // Ví dụ: "ABCDE" -> "BCDEA" -> "CDEAB" -> ...
                string displayText = richTextMarqueeText.Substring(richTextMarqueePosition) + 
                                   richTextMarqueeText.Substring(0, richTextMarqueePosition);

                // Cập nhật text mà không kích hoạt các sự kiện
                int selectionStart = richTextBox1.SelectionStart;
                richTextBox1.Text = displayText;

                // Khôi phục vị trí con trỏ nếu người dùng không đang chỉnh sửa
                if (!richTextBox1.Focused)
                {
                    richTextBox1.SelectionStart = 0;
                    richTextBox1.SelectionLength = 0;
                }
            }
            catch { }
        }

        // ============================================
        // XỬ LÝ SỰ KIỆN CLICK VÀO DGV2 (CHỌN TỔ ĐỂ CHỈNH SỬA)
        // ============================================

        /// <summary>
        /// Xử lý sự kiện click vào ô trong dgv2
        /// Load thông tin tổ đã chọn lên form để chỉnh sửa
        /// </summary>
        private void Dgv2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (dgv2 == null || e.RowIndex < 0) return;

                // Lấy dòng được chọn
                var row = dgv2.Rows[e.RowIndex];
                var history = row.DataBoundItem as ExportHistory;
                if (history == null) return;

                // Chuyển sang chế độ Edit
                isEditMode = true;
                currentEditingHistory = history;

                // Load thông tin chung lên form
                try { cbpgd2.Text = history.Pgd ?? ""; } catch { }
                try { cbXa2.Text = history.Xa ?? ""; } catch { }
                try { cbThon2.Text = history.Thon ?? ""; } catch { }
                try { cbTotruong.Text = history.To ?? ""; } catch { }
                try { cbctr.Text = history.Chuongtrinh ?? ""; } catch { }

                // Load thông tin 5 tổ viên - Họ tên
                try { txtkh1.Text = history.Kh1 ?? ""; } catch { }
                try { txtkh2.Text = history.Kh2 ?? ""; } catch { }
                try { txtkh3.Text = history.Kh3 ?? ""; } catch { }
                try { txtkh4.Text = history.Kh4 ?? ""; } catch { }
                try { txtkh5.Text = history.Kh5 ?? ""; } catch { }

                // Load số tiền
                try { cbtien1.Text = history.Tien1 ?? ""; } catch { }
                try { cbtien2.Text = history.Tien2 ?? ""; } catch { }
                try { cbtien3.Text = history.Tien3 ?? ""; } catch { }
                try { cbtien4.Text = history.Tien4 ?? ""; } catch { }
                try { cbtien5.Text = history.Tien5 ?? ""; } catch { }

                // Load phương án
                try { cbmd1.Text = history.Md1 ?? ""; } catch { }
                try { cbmd2.Text = history.Md2 ?? ""; } catch { }
                try { cbmd3.Text = history.Md3 ?? ""; } catch { }
                try { cbmd4.Text = history.Md4 ?? ""; } catch { }
                try { cbmd5.Text = history.Md5 ?? ""; } catch { }

                // Load thời hạn vay
                try { cbtime1.Text = history.Time1 ?? ""; } catch { }
                try { cbtime2.Text = history.Time2 ?? ""; } catch { }
                try { cbtime3.Text = history.Time3 ?? ""; } catch { }
                try { cbtime4.Text = history.Time4 ?? ""; } catch { }
                try { cbtime5.Text = history.Time5 ?? ""; } catch { }

                // Load đối tượng
                try { cbdt1.Text = history.Dt1 ?? ""; } catch { }
                try { cbdt2.Text = history.Dt2 ?? ""; } catch { }
                try { cbdt3.Text = history.Dt3 ?? ""; } catch { }
                try { cbdt4.Text = history.Dt4 ?? ""; } catch { }
                try { cbdt5.Text = history.Dt5 ?? ""; } catch { }

                // Cập nhật text của button
                try { btn03to.Text = "Cập nhật tổ"; } catch { }

                MessageBox.Show("Đã load thông tin tổ. Bấm 'Cập nhật tổ' để ghi đè hoặc 'Tạo tổ mới' để lưu thành tổ mới.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi load dữ liệu tổ: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ============================================
        // XỬ LÝ SỰ KIỆN CLICK NÚT TẠO TỔ MỚI
        // ============================================

        /// <summary>
        /// Xử lý sự kiện click nút "Tạo tổ mới" (btntaoto)
        /// Tạo tổ mới từ dữ liệu hiện tại trên form
        /// </summary>
        private void BtnTaoTo_Click(object sender, EventArgs e)
        {
            try
            {
                // Kiểm tra xem có dữ liệu trên form không
                string totruong = Clean(cbTotruong.Text);
                string xa = Clean(cbXa2.Text);
                string thon = Clean(cbThon2.Text);
                string chuongtrinh = Clean(cbctr.Text);

                // Đếm số người đã nhập
                string[] hoten = {
                    "",
                    Clean(txtkh1.Text),
                    Clean(txtkh2.Text),
                    Clean(txtkh3.Text),
                    Clean(txtkh4.Text),
                    Clean(txtkh5.Text)
                };

                int soNguoi = Enumerable.Range(1, 5).Count(i => !string.IsNullOrWhiteSpace(hoten[i]));

                if (soNguoi < 2 || string.IsNullOrWhiteSpace(totruong))
                {
                    MessageBox.Show("Phải có tối thiểu 2 người và tên tổ trưởng mới được tạo tổ mới.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Hỏi xác nhận
                var result = MessageBox.Show(
                    $"Tạo tổ mới với thông tin:\n" +
                    $"- Tổ trưởng: {totruong}\n" +
                    $"- Xã: {xa}\n" +
                    $"- Thôn: {thon}\n" +
                    $"- Số thành viên: {soNguoi}\n\n" +
                    $"Bạn có chắc chắn muốn tạo tổ mới?",
                    "Xác nhận tạo tổ mới",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result != DialogResult.Yes) return;

                // Tạo ExportHistory mới
                var newHistory = new ExportHistory
                {
                    Xa = xa,
                    Thon = thon,
                    To = totruong,
                    Chuongtrinh = chuongtrinh,
                    Pgd = Clean(cbpgd2.Text),
                    NgayXuat = DateTime.Now,
                    SoThanhVien = soNguoi,

                    // Lưu thông tin 5 tổ viên
                    Kh1 = hoten[1],
                    Kh2 = hoten[2],
                    Kh3 = hoten[3],
                    Kh4 = hoten[4],
                    Kh5 = hoten[5],

                    Tien1 = cbtien1.Text,
                    Tien2 = cbtien2.Text,
                    Tien3 = cbtien3.Text,
                    Tien4 = cbtien4.Text,
                    Tien5 = cbtien5.Text,

                    Md1 = cbmd1.Text,
                    Md2 = cbmd2.Text,
                    Md3 = cbmd3.Text,
                    Md4 = cbmd4.Text,
                    Md5 = cbmd5.Text,

                    Time1 = cbtime1.Text,
                    Time2 = cbtime2.Text,
                    Time3 = cbtime3.Text,
                    Time4 = cbtime4.Text,
                    Time5 = cbtime5.Text,

                    Dt1 = cbdt1.Text,
                    Dt2 = cbdt2.Text,
                    Dt3 = cbdt3.Text,
                    Dt4 = cbdt4.Text,
                    Dt5 = cbdt5.Text
                };

                // Thêm tổ mới vào danh sách
                exportHistories.Add(newHistory);

                // Lưu vào file JSON
                var toData = ConvertExportHistoryToToData(newHistory);
                SaveToDataToFile(toData);

                // Chuyển sang chế độ Add (không edit nữa)
                isEditMode = false;
                currentEditingHistory = null;

                // Cập nhật text của button
                try { btn03to.Text = "Xuất tổ"; } catch { }

                // Bỏ chọn dòng trong dgv2
                if (dgv2 != null)
                {
                    dgv2.ClearSelection();
                }

                MessageBox.Show("Đã tạo tổ mới thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tạo tổ mới: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Xóa tất cả dữ liệu trên form
        /// </summary>
        private void ClearAllFields()
        {
            try
            {
                // Xóa thông tin chung
                try { cbTotruong.Text = ""; } catch { }
                try { cbXa2.Text = ""; } catch { }
                try { cbThon2.Text = ""; } catch { }
                try { cbctr.Text = ""; } catch { }
                try { cbpgd2.Text = ""; } catch { }

                // Xóa thông tin tổ viên
                ClearToVienFields();
            }
            catch { }
        }

        /// <summary>
        /// Xóa thông tin 5 tổ viên
        /// </summary>
        private void ClearToVienFields()
        {
            try
            {
                // Xóa họ tên
                try { txtkh1.Text = ""; } catch { }
                try { txtkh2.Text = ""; } catch { }
                try { txtkh3.Text = ""; } catch { }
                try { txtkh4.Text = ""; } catch { }
                try { txtkh5.Text = ""; } catch { }

                // Xóa số tiền
                try { cbtien1.Text = ""; } catch { }
                try { cbtien2.Text = ""; } catch { }
                try { cbtien3.Text = ""; } catch { }
                try { cbtien4.Text = ""; } catch { }
                try { cbtien5.Text = ""; } catch { }

                // Xóa phương án
                try { cbmd1.Text = ""; } catch { }
                try { cbmd2.Text = ""; } catch { }
                try { cbmd3.Text = ""; } catch { }
                try { cbmd4.Text = ""; } catch { }
                try { cbmd5.Text = ""; } catch { }

                // Xóa thời hạn
                try { cbtime1.Text = ""; } catch { }
                try { cbtime2.Text = ""; } catch { }
                try { cbtime3.Text = ""; } catch { }
                try { cbtime4.Text = ""; } catch { }
                try { cbtime5.Text = ""; } catch { }

                // Xóa đối tượng
                try { cbdt1.Text = ""; } catch { }
                try { cbdt2.Text = ""; } catch { }
                try { cbdt3.Text = ""; } catch { }
                try { cbdt4.Text = ""; } catch { }
                try { cbdt5.Text = ""; } catch { }
            }
            catch { }
        }

        // ============================================
        // QUẢN LÝ FILE TỔ (LƯU/TẢI TỪ JSON)
        // ============================================
        // Tương tự Form1 - Mỗi tổ được lưu thành 1 file JSON riêng biệt trong folder "To"
        // Tên file: Totruong_Xa_Thon.json (ví dụ: NguyenVanA_XinMan_Lung.json)
        // ============================================

        /// <summary>
        /// Đảm bảo folder "To" tồn tại
        /// </summary>
        private void EnsureToFolder()
        {
            var folder = GetToFolderPath();
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
        }

        /// <summary>
        /// Load tất cả tổ từ file JSON trong folder "To"
        /// </summary>
        private void LoadToFromFiles()
        {
            savedToList = new System.ComponentModel.BindingList<ToData>();
            exportHistories.Clear();

            try
            {
                EnsureToFolder();
                var folder = GetToFolderPath();

                foreach (var file in Directory.EnumerateFiles(folder, "*.json"))
                {
                    try
                    {
                        var json = File.ReadAllText(file, System.Text.Encoding.UTF8);
                        var toData = Newtonsoft.Json.JsonConvert.DeserializeObject<ToData>(json);
                        if (toData != null)
                        {
                            toData._fileName = Path.GetFileName(file);
                            savedToList.Add(toData);

                            // Chuyển đổi ToData thành ExportHistory để hiển thị trong dgv2
                            var history = ConvertToDataToExportHistory(toData);
                            exportHistories.Add(history);
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

        /// <summary>
        /// Chuyển đổi ToData thành ExportHistory
        /// </summary>
        private ExportHistory ConvertToDataToExportHistory(ToData toData)
        {
            return new ExportHistory
            {
                Pgd = toData.Pgd,
                Xa = toData.Xa,
                Thon = toData.Thon,
                To = toData.Totruong,
                Chuongtrinh = toData.Chuongtrinh,
                NgayXuat = toData.NgayXuat,
                SoThanhVien = toData.SoThanhVien,

                Kh1 = toData.Kh1,
                Kh2 = toData.Kh2,
                Kh3 = toData.Kh3,
                Kh4 = toData.Kh4,
                Kh5 = toData.Kh5,

                Tien1 = toData.Tien1,
                Tien2 = toData.Tien2,
                Tien3 = toData.Tien3,
                Tien4 = toData.Tien4,
                Tien5 = toData.Tien5,

                Md1 = toData.Md1,
                Md2 = toData.Md2,
                Md3 = toData.Md3,
                Md4 = toData.Md4,
                Md5 = toData.Md5,

                Time1 = toData.Time1,
                Time2 = toData.Time2,
                Time3 = toData.Time3,
                Time4 = toData.Time4,
                Time5 = toData.Time5,

                Dt1 = toData.Dt1,
                Dt2 = toData.Dt2,
                Dt3 = toData.Dt3,
                Dt4 = toData.Dt4,
                Dt5 = toData.Dt5
            };
        }

        /// <summary>
        /// Chuyển đổi ExportHistory thành ToData
        /// </summary>
        private ToData ConvertExportHistoryToToData(ExportHistory history)
        {
            return new ToData
            {
                Pgd = history.Pgd,
                Xa = history.Xa,
                Thon = history.Thon,
                Totruong = history.To,
                Chuongtrinh = history.Chuongtrinh,
                NgayXuat = history.NgayXuat,
                SoThanhVien = history.SoThanhVien,

                Kh1 = history.Kh1,
                Kh2 = history.Kh2,
                Kh3 = history.Kh3,
                Kh4 = history.Kh4,
                Kh5 = history.Kh5,

                Tien1 = history.Tien1,
                Tien2 = history.Tien2,
                Tien3 = history.Tien3,
                Tien4 = history.Tien4,
                Tien5 = history.Tien5,

                Md1 = history.Md1,
                Md2 = history.Md2,
                Md3 = history.Md3,
                Md4 = history.Md4,
                Md5 = history.Md5,

                Time1 = history.Time1,
                Time2 = history.Time2,
                Time3 = history.Time3,
                Time4 = history.Time4,
                Time5 = history.Time5,

                Dt1 = history.Dt1,
                Dt2 = history.Dt2,
                Dt3 = history.Dt3,
                Dt4 = history.Dt4,
                Dt5 = history.Dt5
            };
        }

        /// <summary>
        /// Lưu ToData vào file JSON
        /// </summary>
        private void SaveToDataToFile(ToData toData)
        {
            try
            {
                EnsureToFolder();
                string path;

                if (!string.IsNullOrEmpty(toData._fileName))
                {
                    // Cập nhật file hiện có
                    path = Path.Combine(GetToFolderPath(), toData._fileName);
                }
                else
                {
                    // Tạo file mới
                    var baseName = MakeFileSystemSafeForTo(toData.Totruong ?? "Unknown");
                    var xaName = MakeFileSystemSafeForTo(toData.Xa ?? "");
                    var thonName = MakeFileSystemSafeForTo(toData.Thon ?? "");

                    var fileName = $"{baseName}_{xaName}_{thonName}.json";
                    path = Path.Combine(GetToFolderPath(), fileName);

                    // Tránh trùng tên file
                    int i = 1;
                    while (File.Exists(path))
                    {
                        fileName = $"{baseName}_{xaName}_{thonName}_{i}.json";
                        path = Path.Combine(GetToFolderPath(), fileName);
                        i++;
                    }

                    toData._fileName = fileName;
                }

                var json = Newtonsoft.Json.JsonConvert.SerializeObject(toData, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(path, json, System.Text.Encoding.UTF8);
            }
            catch { }
        }

        /// <summary>
        /// Chuyển chuỗi thành tên file an toàn (loại bỏ ký tự đặc biệt)
        /// </summary>
        private string MakeFileSystemSafeForTo(string input)
        {
            if (string.IsNullOrEmpty(input)) input = "unknown";
            var invalid = Path.GetInvalidFileNameChars();
            foreach (var ch in invalid)
                input = input.Replace(ch.ToString(), "_");
            return input.Trim();
        }

        /// <summary>
        /// Xóa file JSON của tổ
        /// </summary>
        private void DeleteToDataFile(ToData toData)
        {
            try
            {
                if (!string.IsNullOrEmpty(toData._fileName))
                {
                    var path = Path.Combine(GetToFolderPath(), toData._fileName);
                    if (File.Exists(path))
                        File.Delete(path);
                }
            }
            catch { }
        }
    }

    // ============================================
    // LỚP LƯU TRẠNG THÁI FORM2
    // ============================================

    /// <summary>
    /// Lớp lưu trữ trạng thái của Form2
    /// Sử dụng để serialize/deserialize dữ liệu form vào file JSON
    /// Cho phép khôi phục dữ liệu đã nhập khi mở lại form
    /// </summary>
    public class Form2State
    {
        // -------- THÔNG TIN CHUNG --------
        /// <summary>Phòng giao dịch</summary>
        public string Pgd { get; set; }

        /// <summary>Tên tổ trưởng</summary>
        public string Totruong { get; set; }

        /// <summary>Tên xã</summary>
        public string Xa { get; set; }

        /// <summary>Tên thôn</summary>
        public string Thon { get; set; }

        /// <summary>Chương trình</summary>
        public string Chuongtrinh { get; set; }

        // -------- HỌ TÊN 5 TỔ VIÊN --------
        /// <summary>Họ tên tổ viên 1</summary>
        public string Kh1 { get; set; }

        /// <summary>Họ tên tổ viên 2</summary>
        public string Kh2 { get; set; }

        /// <summary>Họ tên tổ viên 3</summary>
        public string Kh3 { get; set; }

        /// <summary>Họ tên tổ viên 4</summary>
        public string Kh4 { get; set; }

        /// <summary>Họ tên tổ viên 5</summary>
        public string Kh5 { get; set; }

        // -------- SỐ TIỀN 5 TỔ VIÊN --------
        /// <summary>Số tiền tổ viên 1</summary>
        public string Tien1 { get; set; }

        /// <summary>Số tiền tổ viên 2</summary>
        public string Tien2 { get; set; }

        /// <summary>Số tiền tổ viên 3</summary>
        public string Tien3 { get; set; }

        /// <summary>Số tiền tổ viên 4</summary>
        public string Tien4 { get; set; }

        /// <summary>Số tiền tổ viên 5</summary>
        public string Tien5 { get; set; }

        // -------- PHƯƠNG ÁN 5 TỔ VIÊN --------
        /// <summary>Phương án (mục đích) tổ viên 1</summary>
        public string Md1 { get; set; }

        /// <summary>Phương án (mục đích) tổ viên 2</summary>
        public string Md2 { get; set; }

        /// <summary>Phương án (mục đích) tổ viên 3</summary>
        public string Md3 { get; set; }

        /// <summary>Phương án (mục đích) tổ viên 4</summary>
        public string Md4 { get; set; }

        /// <summary>Phương án (mục đích) tổ viên 5</summary>
        public string Md5 { get; set; }

        // -------- THỜI HẠN 5 TỔ VIÊN --------
        /// <summary>Thời hạn vay tổ viên 1</summary>
        public string Time1 { get; set; }

        /// <summary>Thời hạn vay tổ viên 2</summary>
        public string Time2 { get; set; }

        /// <summary>Thời hạn vay tổ viên 3</summary>
        public string Time3 { get; set; }

        /// <summary>Thời hạn vay tổ viên 4</summary>
        public string Time4 { get; set; }

        /// <summary>Thời hạn vay tổ viên 5</summary>
        public string Time5 { get; set; }

        // -------- ĐỐI TƯỢNG 5 TỔ VIÊN --------
        /// <summary>Đối tượng tổ viên 1</summary>
        public string Dt1 { get; set; }

        /// <summary>Đối tượng tổ viên 2</summary>
        public string Dt2 { get; set; }

        /// <summary>Đối tượng tổ viên 3</summary>
        public string Dt3 { get; set; }

        /// <summary>Đối tượng tổ viên 4</summary>
        public string Dt4 { get; set; }

        /// <summary>Đối tượng tổ viên 5</summary>
        public string Dt5 { get; set; }
    }

    /// <summary>
    /// Class lưu lịch sử xuất hồ sơ
    /// Hiển thị thông tin: Xã - Thôn - Tổ
    /// </summary>
    public class ExportHistory
    {
        /// <summary>Tên Xã</summary>
        public string Xa { get; set; }

        /// <summary>Tên Thôn</summary>
        public string Thon { get; set; }

        /// <summary>Tên Tổ trưởng</summary>
        public string To { get; set; }

        /// <summary>Chương trình</summary>
        public string Chuongtrinh { get; set; }

        /// <summary>Phòng giao dịch</summary>
        public string Pgd { get; set; }

        /// <summary>Thời gian xuất</summary>
        public DateTime NgayXuat { get; set; }

        /// <summary>Số thành viên</summary>
        public int SoThanhVien { get; set; }

        // -------- THÔNG TIN 5 TỔ VIÊN --------
        /// <summary>Họ tên tổ viên 1</summary>
        public string Kh1 { get; set; }
        /// <summary>Họ tên tổ viên 2</summary>
        public string Kh2 { get; set; }
        /// <summary>Họ tên tổ viên 3</summary>
        public string Kh3 { get; set; }
        /// <summary>Họ tên tổ viên 4</summary>
        public string Kh4 { get; set; }
        /// <summary>Họ tên tổ viên 5</summary>
        public string Kh5 { get; set; }

        /// <summary>Số tiền tổ viên 1</summary>
        public string Tien1 { get; set; }
        /// <summary>Số tiền tổ viên 2</summary>
        public string Tien2 { get; set; }
        /// <summary>Số tiền tổ viên 3</summary>
        public string Tien3 { get; set; }
        /// <summary>Số tiền tổ viên 4</summary>
        public string Tien4 { get; set; }
        /// <summary>Số tiền tổ viên 5</summary>
        public string Tien5 { get; set; }

        /// <summary>Phương án tổ viên 1</summary>
        public string Md1 { get; set; }
        /// <summary>Phương án tổ viên 2</summary>
        public string Md2 { get; set; }
        /// <summary>Phương án tổ viên 3</summary>
        public string Md3 { get; set; }
        /// <summary>Phương án tổ viên 4</summary>
        public string Md4 { get; set; }
        /// <summary>Phương án tổ viên 5</summary>
        public string Md5 { get; set; }

        /// <summary>Thời hạn vay tổ viên 1</summary>
        public string Time1 { get; set; }
        /// <summary>Thời hạn vay tổ viên 2</summary>
        public string Time2 { get; set; }
        /// <summary>Thời hạn vay tổ viên 3</summary>
        public string Time3 { get; set; }
        /// <summary>Thời hạn vay tổ viên 4</summary>
        public string Time4 { get; set; }
        /// <summary>Thời hạn vay tổ viên 5</summary>
        public string Time5 { get; set; }

        /// <summary>Đối tượng tổ viên 1</summary>
        public string Dt1 { get; set; }
        /// <summary>Đối tượng tổ viên 2</summary>
        public string Dt2 { get; set; }
        /// <summary>Đối tượng tổ viên 3</summary>
        public string Dt3 { get; set; }
        /// <summary>Đối tượng tổ viên 4</summary>
        public string Dt4 { get; set; }
        /// <summary>Đối tượng tổ viên 5</summary>
        public string Dt5 { get; set; }

        /// <summary>Constructor</summary>
        public ExportHistory()
        {
            NgayXuat = DateTime.Now;
        }
    }
}
