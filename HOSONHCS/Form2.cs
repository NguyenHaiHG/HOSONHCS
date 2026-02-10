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
        /// Đường dẫn file JSON để lưu trạng thái form
        /// File này lưu tất cả thông tin đã nhập để khôi phục khi mở lại form
        /// </summary>
        private string Form2StatePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Form2State.json");

        /// <summary>
        /// Cờ xác định có nên load trạng thái đã lưu hay không
        /// = false khi mở form từ Form1 với khách hàng đã chọn
        /// = true khi mở form độc lập
        /// </summary>
        private bool shouldLoadState = true;

        /// <summary>
        /// Biến lưu văn bản chạy (marquee) trong richTextBox1
        /// </summary>
        private string richTextMarqueeText = "";

        /// <summary>
        /// Vị trí hiện tại của văn bản chạy (marquee)
        /// </summary>
        private int richTextMarqueePosition = 0;

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

            // Đăng ký sự kiện click vào DataGridView để load dữ liệu lên form
            try { dataGridView1.CellClick += DataGridView1_CellClick; } catch { }

            // Đăng ký xử lý nhập liệu cho các ô tiền (cbtien1-5)
            // Chỉ cho phép nhập số và tự động định dạng với dấu chấm phân cách hàng nghìn
            try { cbtien1.KeyPress += CbMoney_KeyPress; cbtien1.TextChanged += CbMoney_TextChanged; } catch { }
            try { cbtien2.KeyPress += CbMoney_KeyPress; cbtien2.TextChanged += CbMoney_TextChanged; } catch { }
            try { cbtien3.KeyPress += CbMoney_KeyPress; cbtien3.TextChanged += CbMoney_TextChanged; } catch { }
            try { cbtien4.KeyPress += CbMoney_KeyPress; cbtien4.TextChanged += CbMoney_TextChanged; } catch { }
            try { cbtien5.KeyPress += CbMoney_KeyPress; cbtien5.TextChanged += CbMoney_TextChanged; } catch { }

            // Đăng ký xử lý nhập liệu cho các ô tên và địa điểm
            // Chỉ cho phép nhập chữ cái, không cho phép nhập số
            try { txtxa.KeyPress += TextLettersOnly_KeyPress; } catch { }
            try { txttotruong.KeyPress += TextLettersOnly_KeyPress; } catch { }
            try { txtkh1.KeyPress += TextLettersOnly_KeyPress; txtkh2.KeyPress += TextLettersOnly_KeyPress; txtkh3.KeyPress += TextLettersOnly_KeyPress; txtkh4.KeyPress += TextLettersOnly_KeyPress; txtkh5.KeyPress += TextLettersOnly_KeyPress; } catch { }

            // Áp dụng giao diện MacBook theme
            try { ApplyMacBookTheme(); } catch { }

            // Khởi tạo hiệu ứng chữ chạy cho richTextBox1
            try { InitializeRichTextMarquee(); } catch { }

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

                        // Cập nhật DataGridView với danh sách đã nhóm
                        try { dataGridView1.DataSource = null; dataGridView1.DataSource = selectedCustomers; } catch { }
                    }
                    catch { }

                    // KHÔNG điền txtkh1..txtkh5 từ danh sách thành viên được truyền từ Form1
                    // Form chỉ hiển thị thông tin người tổ chức theo mặc định
                    // Chỉ điền các trường chung cấp cao nhất từ khách hàng đầu tiên (nếu có)
                    var firstCustomer = selected.FirstOrDefault();
                    if (firstCustomer != null)
                    {
                        try { txttotruong.Text = firstCustomer.Totruong ?? ""; } catch { }
                        try { txtxa.Text = firstCustomer.Xa ?? ""; } catch { }
                        try { cbctr.Text = firstCustomer.Chuongtrinh ?? ""; } catch { }
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
                string totruong = Clean(txttotruong.Text);
                string xa = Clean(txtxa.Text);
                string thon = Clean(txtthon.Text);
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
                    Clean(txtmd1.Text),
                    Clean(txtmd2.Text),
                    Clean(txtmd3.Text),
                    Clean(txtmd4.Text),
                    Clean(txtmd5.Text)
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
                    // Cập nhật lại DataGridView
                    try { dataGridView1.DataSource = null; dataGridView1.DataSource = selectedCustomers; } catch { }
                }
                catch { }

                // -------- XUẤT FILE WORD --------
                // Gọi phương thức ExportWord để tạo file Word từ mẫu
                await Task.Run(() => ExportWord(map, hoten, sotien, totruong, xa, chuongtrinh));

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
                    Totruong = txttotruong.Text,
                    Xa = txtxa.Text,
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
                    Md1 = txtmd1.Text,
                    Md2 = txtmd2.Text,
                    Md3 = txtmd3.Text,
                    Md4 = txtmd4.Text,
                    Md5 = txtmd5.Text,

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
                try { txttotruong.Text = state.Totruong ?? ""; } catch { }
                try { txtxa.Text = state.Xa ?? ""; } catch { }
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
                try { txtmd1.Text = state.Md1 ?? ""; } catch { }
                try { txtmd2.Text = state.Md2 ?? ""; } catch { }
                try { txtmd3.Text = state.Md3 ?? ""; } catch { }
                try { txtmd4.Text = state.Md4 ?? ""; } catch { }
                try { txtmd5.Text = state.Md5 ?? ""; } catch { }

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
                if (dataGridView1 == null) return;

                // Tạo danh sách các khách hàng cần xóa
                var toRemove = new List<Customer>();
                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    try { var item = row.DataBoundItem as Customer; if (item != null) toRemove.Add(item); } catch { }
                }

                // Xóa từng khách hàng khỏi danh sách
                foreach (var r in toRemove)
                {
                    selectedCustomers.Remove(r);
                }

                // Cập nhật lại DataGridView
                try { dataGridView1.DataSource = null; dataGridView1.DataSource = selectedCustomers; } catch { }
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
                // Bỏ qua click vào header
                if (e.RowIndex < 0) return;

                // Lấy dòng được chọn
                var row = dataGridView1.Rows[e.RowIndex];
                var customer = row.DataBoundItem as Customer;
                if (customer == null) return;

                // -------- LOAD THÔNG TIN CHUNG LÊN FORM --------
                try { txttotruong.Text = customer.Totruong ?? ""; } catch { }
                try { txtxa.Text = customer.Xa ?? ""; } catch { }
                try { cbctr.Text = customer.Chuongtrinh ?? ""; } catch { }

                // -------- LOAD THÔNG TIN TỔ VIÊN THỨ 1 --------
                try { txtkh1.Text = customer.Hoten ?? ""; } catch { }
                try { cbtien1.Text = customer.Sotien ?? ""; } catch { }
                try { txtmd1.Text = customer.Phuongan ?? ""; } catch { }
                try { cbtime1.Text = customer.Thoihanvay ?? ""; } catch { }
                try { cbdt1.Text = customer.Doituong1 ?? ""; } catch { }

                // -------- XÓA CÁC Ô KHÁC --------
                try { txtkh2.Text = ""; txtkh3.Text = ""; txtkh4.Text = ""; txtkh5.Text = ""; } catch { }
                try { cbtien2.Text = ""; cbtien3.Text = ""; cbtien4.Text = ""; cbtien5.Text = ""; } catch { }
                try { txtmd2.Text = ""; txtmd3.Text = ""; txtmd4.Text = ""; txtmd5.Text = ""; } catch { }
                try { cbtime2.Text = ""; cbtime3.Text = ""; cbtime4.Text = ""; cbtime5.Text = ""; } catch { }
                try { cbdt2.Text = ""; cbdt3.Text = ""; cbdt4.Text = ""; cbdt5.Text = ""; } catch { }
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
            // Load trạng thái đã lưu khi form load
            LoadFormState();
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
        // ÁP DỤNG GIAO DIỆN MACBOOK THEME
        // ============================================

        /// <summary>
        /// Áp dụng giao diện MacBook Theme cho toàn bộ form
        /// Thay đổi màu sắc, font chữ và hiệu ứng cho tất cả các controls
        /// </summary>
        private void ApplyMacBookTheme()
        {
            try
            {
                // Đặt màu nền cho form
                this.BackColor = AppTheme.MacBackground;

                // Định dạng label tiêu đề (label1)
                if (label1 != null)
                {
                    label1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold);
                    label1.ForeColor = AppTheme.MacBlue;
                }

                // Định dạng các nút bấm
                StyleMacButton(btn03to, AppTheme.MacGreen);   // Nút Xuất - Màu xanh lá
                StyleMacButton(btnxoa, AppTheme.MacRed);       // Nút Xóa - Màu đỏ

                // Định dạng DataGridView
                StyleMacDataGridView();

                // Áp dụng font chữ cho tất cả các labels
                ApplyMacFontsToLabels();

                // Định dạng textboxes và comboboxes
                ApplyMacStyleToTextBoxes();
                ApplyMacStyleToComboBoxes();
            }
            catch { }
        }

        /// <summary>
        /// Định dạng nút bấm theo phong cách MacBook
        /// Thay đổi màu, font, và thêm hiệu ứng hover
        /// </summary>
        /// <param name="btn">Nút bấm cần định dạng</param>
        /// <param name="color">Màu sắc của nút</param>
        private void StyleMacButton(System.Windows.Forms.Button btn, System.Drawing.Color color)
        {
            if (btn == null) return;

            try
            {
                // Cài đặt kiểu hiển thị phẳng (flat style)
                btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderSize = 0;  // Không viền
                btn.BackColor = color;
                btn.ForeColor = System.Drawing.Color.White;
                btn.Font = new System.Drawing.Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Regular);
                btn.Cursor = Cursors.Hand;  // Con trỏ chuột dạng bàn tay
                btn.Height = 36;

                // Xác định màu hover dựa trên màu gốc
                System.Drawing.Color originalColor = color;
                System.Drawing.Color hoverColor = color.ToArgb() == AppTheme.MacGreen.ToArgb() ? AppTheme.MacGreenHover :
                                  color.ToArgb() == AppTheme.MacRed.ToArgb() ? AppTheme.MacRedHover :
                                  AppTheme.MacBlueHover;

                // Thêm sự kiện hover (chuột vào/ra)
                btn.MouseEnter += (s, e) => { btn.BackColor = hoverColor; };
                btn.MouseLeave += (s, e) => { btn.BackColor = originalColor; };
            }
            catch { }
        }

        /// <summary>
        /// Định dạng DataGridView theo phong cách MacBook
        /// Thay đổi màu sắc, font chữ và kiểu hiển thị của bảng
        /// </summary>
        private void StyleMacDataGridView()
        {
            try
            {
                if (dataGridView1 == null) return;

                // Cài đặt kiểu hiển thị cơ bản
                dataGridView1.BorderStyle = BorderStyle.None;  // Không viền
                dataGridView1.BackgroundColor = AppTheme.MacCardBackground;
                dataGridView1.GridColor = AppTheme.MacBorderLight;
                dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;  // Viền ngang

                // Kiểu hiển thị của các ô
                dataGridView1.DefaultCellStyle.BackColor = System.Drawing.Color.White;
                dataGridView1.DefaultCellStyle.ForeColor = AppTheme.MacTextPrimary;
                dataGridView1.DefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 8.5F);
                dataGridView1.DefaultCellStyle.SelectionBackColor = AppTheme.MacBlue;  // Màu khi chọn
                dataGridView1.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.White;
                dataGridView1.DefaultCellStyle.Padding = new Padding(6, 3, 6, 3);  // Khoảng cách bên trong

                // Màu sắc luân phiên cho các dòng
                dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(249, 249, 251);

                // Kiểu hiển thị của header (tiêu đề cột)
                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
                dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = AppTheme.MacHeaderGradient1;
                dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = AppTheme.MacTextPrimary;
                dataGridView1.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
                dataGridView1.ColumnHeadersDefaultCellStyle.Padding = new Padding(8, 6, 8, 6);
                dataGridView1.ColumnHeadersHeight = 36;

                // Chiều cao của dòng
                dataGridView1.RowTemplate.Height = 32;
            }
            catch { }
        }

        /// <summary>
        /// Áp dụng font chữ MacBook cho tất cả các labels
        /// Thay đổi font và màu chữ thành màu đen
        /// </summary>
        private void ApplyMacFontsToLabels()
        {
            try
            {
                System.Drawing.Font labelFont = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular);

                // Lặp qua tất cả các controls trong form
                foreach (System.Windows.Forms.Control ctrl in GetAllControlsForTheme(this))
                {
                    if (ctrl is Label lbl)
                    {
                        // Áp dụng màu ĐEN cho tất cả labels (không có ngoại lệ trong Form2)
                        lbl.Font = labelFont;
                        lbl.ForeColor = System.Drawing.Color.Black; // Đặt màu đen cho tất cả
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Áp dụng phong cách MacBook cho tất cả các TextBoxes - Không viền (borderless)
        /// Thay đổi font, màu và thêm hiệu ứng focus
        /// Chỉ dùng màu background để phân biệt ô nhập liệu
        /// </summary>
        private void ApplyMacStyleToTextBoxes()
        {
            try
            {
                System.Drawing.Font textFont = new System.Drawing.Font("Segoe UI", 8.5F, System.Drawing.FontStyle.Regular);

                foreach (System.Windows.Forms.Control ctrl in GetAllControlsForTheme(this))
                {
                    if (ctrl is TextBox txt)
                    {
                        txt.Font = textFont;
                        txt.BackColor = AppTheme.MacInputBackground;
                        txt.ForeColor = AppTheme.MacTextPrimary;
                        txt.BorderStyle = BorderStyle.None;  // BỎ VIỀN ĐEN - style hiện đại

                        // Thêm hiệu ứng khi focus (đang nhập liệu) - đổi màu background
                        txt.Enter += (s, e) => { ((TextBox)s).BackColor = AppTheme.MacInputBackgroundFocus; };
                        txt.Leave += (s, e) => { ((TextBox)s).BackColor = AppTheme.MacInputBackground; };
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Áp dụng phong cách MacBook cho tất cả các ComboBoxes - Flat style không viền
        /// Thay đổi font, màu và kiểu hiển thị
        /// </summary>
        private void ApplyMacStyleToComboBoxes()
        {
            try
            {
                System.Drawing.Font comboFont = new System.Drawing.Font("Segoe UI", 8.5F, System.Drawing.FontStyle.Regular);

                foreach (System.Windows.Forms.Control ctrl in GetAllControlsForTheme(this))
                {
                    if (ctrl is ComboBox cb)
                    {
                        cb.Font = comboFont;
                        cb.BackColor = AppTheme.MacInputBackground;
                        cb.ForeColor = AppTheme.MacTextPrimary;
                        cb.FlatStyle = FlatStyle.Flat;  // Flat style - không viền 3D

                        // ComboBox trong FlatStyle.Flat tự động không có viền đen
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Lấy tất cả các controls trong container (bao gồm các control con)
        /// Sử dụng đệ quy để lấy hết tất cả controls lồng nhau
        /// </summary>
        /// <param name="container">Container chứa các controls</param>
        /// <returns>IEnumerable chứa tất cả các controls</returns>
        private IEnumerable<System.Windows.Forms.Control> GetAllControlsForTheme(System.Windows.Forms.Control container)
        {
            foreach (System.Windows.Forms.Control ctrl in container.Controls)
            {
                yield return ctrl;
                // Đệ quy để lấy các control con
                foreach (System.Windows.Forms.Control child in GetAllControlsForTheme(ctrl))
                {
                    yield return child;
                }
            }
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
        /// <summary>Tên tổ trưởng</summary>
        public string Totruong { get; set; }

        /// <summary>Tên xã</summary>
        public string Xa { get; set; }

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
}
