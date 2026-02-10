using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HOSONHCS
{
    public partial class Form1 : Form
    {
        // ========== HẰNG SỐ ==========
        // Thư mục chứa các file template Word (.docx)
        private const string TemplatesFolder = "Templates";
        // Tên file template 01 HN.docx được nhúng sẵn trong project (Embedded Resource)
        private const string EmbeddedTemplateFileName = "01 HN.docx";

        // ========== DỮ LIỆU KHÁCH HÀNG ==========
        // Danh sách khách hàng (BindingList để tự động cập nhật DataGridView)
        private BindingList<Customer> customers;
        // Chỉ số khách hàng đang được sửa (-1 = không sửa ai, đang tạo mới)
        private int editingIndex = -1;

        // ========== XINMAN DATA (PGD/XÃ/THÔN/HỘI/TỔ) ==========
        // Model chứa toàn bộ dữ liệu cơ cấu tổ chức từ xinman.json
        private XinManModel xinmanModel;
        // Editor để chỉnh sửa xinman.json trên tab3 (cần login)
        private XinManEditor xinManEditor;

        // ========== CÁC CỞ SUPPRESS (NGĂN SỰ KIỆN ĐỆ QUY) ==========
        // Cờ để ngăn combobox gọi sự kiện SelectedIndexChanged khi đang load dữ liệu
        private bool suppressComboChanged = false;
        // Cờ để ngăn textbox ngày sinh gọi TextChanged khi đang format
        private bool suppressNamsinhChanged = false;
        // Cờ để ngăn textbox tên gọi TextChanged khi đang tự động viết hoa
        private bool suppressNameChanged = false;
        // Cờ để ngăn combobox số tiền gọi TextChanged khi đang format
        private bool suppressMoneyChange = false;

        // ========== MARQUEE TEXT FOR FORM TITLE (THANH TIÊU ĐỀ) ==========
        // Biến để lưu text chạy ở thanh tiêu đề Form (title bar) và vị trí hiện tại
        private string marqueeText = "";
        private int marqueePosition = 0;

        // ========== MARQUEE TEXT FOR LABEL14 (BÊN TRONG FORM) ==========
        // Biến để lưu text chạy cho label14 bên trong Form
        private string label14MarqueeText = "";
        private int label14MarqueePosition = 0;

        // Timer được tạo trong Designer file (Form1.Designer.cs)

        // ========== CACHE ==========
        // Cache đường dẫn template để tránh extract nhiều lần từ Embedded Resource
        private static readonly Dictionary<string, string> templatePathCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public Form1()
        {
            InitializeComponent();
            InitializeApp();
        }

        private void InitializeApp()
        {
            // ========== GẮN CÁC SỰ KIỆN BUTTON ==========
            // Gắn sự kiện click cho các nút chính
            try { btn01.Click += BtnSave_Click; } catch { }
            try { btnDelete.Click += BtnDelete_Click; } catch { }

            // Các nút mới có thể chưa có trong designer cũ → dùng try-catch để tránh lỗi
            try { btn03.Click += Btn03_Click; } catch { /* bỏ qua nếu control không tồn tại */ }
            try { btn03Group.Click += Btn03Group_Click; } catch { /* bỏ qua nếu control không tồn tại */ }
            try { btnGUQ.Click += BtnGUQ_Click; } catch { /* bỏ qua nếu control không tồn tại */ }

            // Các nút tạo khách hàng mới và đăng xuất
            try { btntaokh.Click += BtnTaokh_Click; } catch { /* bỏ qua nếu control không tồn tại */ }
            try { btnexit.Click += BtnExit_Click; } catch { /* bỏ qua nếu control không tồn tại */ }

            // ========== TỰ ĐỘNG VIẾT HOA CHỮ CÁI ĐẦU CHO Ô NHẬP TÊN ==========
            // Các ô nhập tên (txtHoten, txtntk1/2/3): tự động viết hoa chữ cái đầu mỗi từ (Title Case)
            try { txtHoten.TextChanged += TxtName_TextChanged; txtHoten.Leave += TxtName_Leave; } catch { }
            try { txtntk1.TextChanged += TxtName_TextChanged; txtntk1.Leave += TxtName_Leave; } catch { }
            try { txtntk2.TextChanged += TxtName_TextChanged; txtntk2.Leave += TxtName_Leave; } catch { }
            try { txtntk3.TextChanged += TxtName_TextChanged; txtntk3.Leave += TxtName_Leave; } catch { }

            // ========== DATEPICKER NGÀY SINH NGƯỜI THỪA KẾ (NTK) ==========
            // datentk1/2/3 là DateTimePicker, KHÔNG CẦN KeyPress/TextChanged events
            // DateTimePicker tự động xử lý nhập liệu, chỉ cần set MaxDate và ShowCheckBox
            // ========== VALIDATION SỐ CCCD ==========
            // Các ô nhập CCCD (txtSocccd, txtcccd1/2/3): chỉ cho phép nhập số, đúng 12 chữ số
            try { txtSocccd.KeyPress += TxtDigitsOnly_KeyPress; txtSocccd.TextChanged += TxtCccd_TextChanged; txtSocccd.Leave += TxtCccd_Leave; txtSocccd.MaxLength = 12; } catch { }
            try { txtcccd1.KeyPress += TxtDigitsOnly_KeyPress; txtcccd1.TextChanged += TxtCccd_TextChanged; txtcccd1.Leave += TxtCccd_Leave; txtcccd1.MaxLength = 12; } catch { }
            try { txtcccd2.KeyPress += TxtDigitsOnly_KeyPress; txtcccd2.TextChanged += TxtCccd_TextChanged; txtcccd2.Leave += TxtCccd_Leave; txtcccd2.MaxLength = 12; } catch { }
            try { txtcccd3.KeyPress += TxtDigitsOnly_KeyPress; txtcccd3.TextChanged += TxtCccd_TextChanged; txtcccd3.Leave += TxtCccd_Leave; txtcccd3.MaxLength = 12; } catch { }

            // ========== SỐ ĐIỆN THOẠI ==========
            // txtSdt: chỉ cho phép nhập số, phải đúng 10 chữ số (không ít hơn, không nhiều hơn)
            try { txtSdt.KeyPress += TxtDigitsOnly_KeyPress; txtSdt.TextChanged += TxtSdt_TextChanged; txtSdt.Leave += TxtSdt_Leave; txtSdt.MaxLength = 10; } catch { }

            // ========== TỰ ĐỘNG CHỌN NỠI CẤP CCCD ==========
            // dateNgaycapCCCD: tự động chọn cbNoicap dựa trên ngày cấp (trước/sau 01/07/2024)
            try { dateNgaycapCCCD.ValueChanged += DateNgaycapCCCD_ValueChanged; } catch { }

            // ========== NGĂN CHỌN NGÀY TRONG TƯƠNG LAI ==========
            // Set MaxDate = hôm nay cho tất cả DateTimePicker để không cho chọn ngày tương lai
            try 
            {
                if (dateLaphs != null) 
                { 
                    dateLaphs.MaxDate = DateTime.Today;
                    dateLaphs.ValueChanged += DatePicker_ValueChanged;
                }
                if (dateNgaycapCCCD != null) 
                { 
                    dateNgaycapCCCD.MaxDate = DateTime.Today;
                    dateNgaycapCCCD.ValueChanged += DatePicker_ValueChanged;
                }
                if (dateNgaysinh != null) 
                { 
                    dateNgaysinh.MaxDate = DateTime.Today;
                    dateNgaysinh.ValueChanged += DatePicker_ValueChanged;
                }
                if (dateDH != null) 
                { 
                    dateDH.MaxDate = DateTime.Today;
                    dateDH.ValueChanged += DatePicker_ValueChanged;
                }
                if (datendhcccd != null) 
                { 
                    datendhcccd.MaxDate = DateTime.Today;
                    datendhcccd.ValueChanged += DatePicker_ValueChanged;
                }
                // ========== DATEPICKER NGÀY SINH NTK: CẦN SET MAXDATE ==========
                // datentk1/2/3 là DateTimePicker, cũng cần set MaxDate để tránh chọn ngày tương lai
                if (datentk1 != null)
                {
                    datentk1.MaxDate = DateTime.Today;
                    datentk1.ValueChanged += DatePicker_ValueChanged;
                }
                if (datentk2 != null)
                {
                    datentk2.MaxDate = DateTime.Today;
                    datentk2.ValueChanged += DatePicker_ValueChanged;
                }
                if (datentk3 != null)
                {
                    datentk3.MaxDate = DateTime.Today;
                    datentk3.ValueChanged += DatePicker_ValueChanged;
                }
            } catch { }

            // ========== COMBOBOX CHỌN ĐỊA ĐIỂM (PGD, XÃ, THÔN, HỘI) ==========
            // Gắn sự kiện để tự động load dữ liệu cascading từ xinman.json
            try { cbPGD.SelectedIndexChanged += CbPGD_SelectedIndexChanged; } catch { }
            try { cbXa.SelectedIndexChanged += CbXa_SelectedIndexChanged; } catch { }
            try { cbThon.SelectedIndexChanged += CbThon_SelectedIndexChanged; } catch { }
            try { cbHoi.SelectedIndexChanged += CbHoi_SelectedIndexChanged; } catch { }

            // ========== FORMAT SỐ TIỀN TỰ ĐỘNG ==========
            // cbSotien: chỉ cho nhập số, tự động format với dấu '.' ngăn cách hàng nghìn
            try { cbSotien.KeyPress += CbMoney_KeyPress; cbSotien.TextChanged += CbMoney_TextChanged; } catch { }

            // ========== HIỂN THỊ CHECKBOX CHO CÁC TRƯỜNG NGÀY OPTIONAL ==========
            // Các DateTimePicker có ShowCheckBox = true để user có thể bỏ chọn (không bắt buộc nhập)
            // Mặc định unchecked = không có giá trị
            try 
            {
                if (dateDH != null) 
                { 
                    dateDH.ShowCheckBox = true; 
                    dateDH.Checked = false;  // Ngày đến hạn: optional
                }
                if (datendhcccd != null) 
                { 
                    datendhcccd.ShowCheckBox = true; 
                    datendhcccd.Checked = false;  // Thời hạn CCCD: optional
                }
                if (datentk1 != null) 
                { 
                    datentk1.ShowCheckBox = true; 
                    datentk1.Checked = false;  // Ngày sinh NTK 1: optional
                }
                if (datentk2 != null) 
                { 
                    datentk2.ShowCheckBox = true; 
                    datentk2.Checked = false;  // Ngày sinh NTK 2: optional
                }
                if (datentk3 != null) 
                { 
                    datentk3.ShowCheckBox = true; 
                    datentk3.Checked = false;  // Ngày sinh NTK 3: optional
                }
            } catch { }

            // ========== CẤU HÌNH DATAGRIDVIEW DANH SÁCH KHÁCH HÀNG ==========
            try
            {
                // Chọn cả dòng, cho phép chọn nhiều dòng để export nhóm
                dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                dgv.MultiSelect = true;
                dgv.AutoGenerateColumns = true;  // Tự động tạo cột từ Customer properties
                dgv.ReadOnly = true;  // Không cho sửa trực tiếp trong grid
                dgv.AllowUserToAddRows = false;
                dgv.AllowUserToDeleteRows = false;
                dgv.CellDoubleClick += Dgv_CellDoubleClick;  // Double-click để edit
                dgv.CellClick += Dgv_CellClick;  // Click để chọn
                dgv.EditMode = DataGridViewEditMode.EditOnEnter;
            }
            catch { }

            // ========== LOAD DỮ LIỆU BAN ĐẦU ==========
            // Load xinman.json (dữ liệu PGD/Xã/Thôn/Hội/Tổ)
            LoadXinManData();
            // Load danh sách khách hàng từ folder Customers/*.json
            LoadCustomersFromFiles();
            // Bind dữ liệu vào DataGridView
            BindGrid();

            // ========== GẮN XINMAN EDITOR (TAB QUẢN LÝ XINMAN.JSON) ==========
            // Attach XinManEditor vào tab3 để chỉnh sửa xinman.json với login/password
            try
            {
                xinManEditor = new XinManEditor();
                xinManEditor.AttachControls(dgv1, txtUsername, txtPassword, btnLogin, txtSearch, btnSave);
            }
            catch { }

            // ========== KHỞI TẠO CHẠY CHỮ CHO THANH TIÊU ĐỀ FORM ==========
            // Khởi tạo hiệu ứng chạy chữ cho thanh tiêu đề Form (title bar)
            // Text "PHẦN MỀM TẠO HỒ SƠ VAY VỐN" sẽ chạy từ phải sang trái
            try { InitializeMarquee(); } catch { }

            // ========== ÁP DỤNG THEME MACBOOK ==========
            // Áp dụng theme hiện đại theo phong cách MacOS
            try { ApplyMacBookTheme(); } catch { }
        }

        // ========================================================================
        // PHẦN QUẢN LÝ FILE KHÁCH HÀNG (LƯU/TẢI TỪ JSON) VÀ HIỂN THỊ DATAGRID
        // ========================================================================
        // Mỗi khách hàng được lưu thành 1 file JSON riêng biệt trong folder "Customers"
        // Tên file: Hoten.json (ví dụ: NguyenVanA.json)
        // ========================================================================

        private string GetCustomersFolderPath()
        {
            // Lấy đường dẫn folder "Customers" bên cạnh file .exe
            // Ví dụ: D:\C#\HOSONHCS\bin\Debug\Customers\
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Customers");
        }

        private void EnsureCustomersFolder()
        {
            // Tạo folder "Customers" nếu chưa tồn tại
            var folder = GetCustomersFolderPath();
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
        }

        private void LoadCustomersFromFiles()
        {
            // Load tất cả file JSON trong folder "Customers" vào danh sách customers
            customers = new BindingList<Customer>();
            try
            {
                EnsureCustomersFolder();
                var folder = GetCustomersFolderPath();
                // Duyệt qua tất cả file .json trong folder Customers
                foreach (var file in Directory.EnumerateFiles(folder, "*.json"))
                {
                    try
                    {
                        // Đọc nội dung file JSON
                        var json = File.ReadAllText(file, Encoding.UTF8);
                        // Deserialize thành object Customer
                        var c = JsonConvert.DeserializeObject<Customer>(json);
                        if (c != null)
                        {
                            // Lưu tên file để biết sửa ai sau này
                            c._fileName = Path.GetFileName(file);
                            customers.Add(c);
                        }
                    }
                    catch
                    {
                        // Nếu file bị lỗi (JSON sai format) thì bỏ qua, không crash app
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load customer files: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // Gắn sự kiện ListChanged (hiện tại không dùng, nhưng giữ để tương lai mở rộng)
            customers.ListChanged += Customers_ListChanged;
        }

        private void Customers_ListChanged(object sender, ListChangedEventArgs e)
        {
            // KHÔNG tự động lưu khi có thay đổi
            // Lưu khi người dùng bấm nút "Lưu" (btn01) thì mới lưu
        }

        private void BindGrid()
        {
            try
            {
                dgv.DataSource = null;
                dgv.DataSource = customers;

                // Đặt tất cả cột tự động tạo thành readonly; chọn theo dòng
                foreach (DataGridViewColumn col in dgv.Columns)
                {
                    col.ReadOnly = true;
                }

                // Tùy chọn chỉ hiển thị các cột bạn muốn. Đảm bảo Họ tên hiển thị đầu tiên
                if (dgv.Columns["Hoten"] != null)
                {
                    dgv.Columns["Hoten"].DisplayIndex = 1;
                    dgv.Columns["Hoten"].HeaderText = "Họ và tên";
                }
                if (dgv.Columns["_fileName"] != null) dgv.Columns["_fileName"].Visible = false;
            }
            catch { }
        }

        private string MakeFileSystemSafe(string input)
        {
            if (string.IsNullOrEmpty(input)) input = "unknown";
            var invalid = Path.GetInvalidFileNameChars();
            foreach (var ch in invalid) input = input.Replace(ch.ToString(), "_");
            return input.Trim();
        }

        private string GetCustomerJsonPath(Customer c)
        {
            if (!string.IsNullOrEmpty(c._fileName))
                return Path.Combine(GetCustomersFolderPath(), c._fileName);

            var baseName = MakeFileSystemSafe(c.Hoten);
            var file = baseName + ".json";
            var folder = GetCustomersFolderPath();
            var path = Path.Combine(folder, file);
            int i = 1;
            while (File.Exists(path))
            {
                file = $"{baseName}_{i}.json";
                path = Path.Combine(folder, file);
                i++;
            }
            return path;
        }

        private string GetCustomerJsonPathByName(string hoten)
        {
            var baseName = MakeFileSystemSafe(hoten);
            var folder = GetCustomersFolderPath();
            var candidate = Path.Combine(folder, baseName + ".json");
            if (File.Exists(candidate)) return candidate;
            var found = Directory.EnumerateFiles(folder, baseName + "*.json").FirstOrDefault();
            return found;
        }

        private void SaveCustomerToFile(Customer c)
        {
            EnsureCustomersFolder();
            string path;
            if (!string.IsNullOrEmpty(c._fileName))
                path = Path.Combine(GetCustomersFolderPath(), c._fileName);
            else
            {
                var baseName = MakeFileSystemSafe(c.Hoten);
                var file = baseName + ".json";
                var folder = GetCustomersFolderPath();
                path = Path.Combine(folder, file);
                int i = 1;
                while (File.Exists(path))
                {
                    var existing = File.ReadAllText(path, Encoding.UTF8);
                    try
                    {
                        var ec = JsonConvert.DeserializeObject<Customer>(existing);
                        if (ec != null && string.Equals(MakeFileSystemSafe(ec.Hoten), MakeFileSystemSafe(c.Hoten), StringComparison.OrdinalIgnoreCase))
                            break;
                    }
                    catch { }
                    file = $"{baseName}_{i}.json";
                    path = Path.Combine(folder, file);
                    i++;
                }
                c._fileName = Path.GetFileName(path);
            }

            UpdateComputedFields(c);

            var json = JsonConvert.SerializeObject(c, Formatting.Indented);
            File.WriteAllText(path, json, Encoding.UTF8);
        }

        private void DeleteCustomerFiles(Customer c)
        {
            try
            {
                if (!string.IsNullOrEmpty(c._fileName))
                {
                    var jsonPath = Path.Combine(GetCustomersFolderPath(), c._fileName);
                    if (File.Exists(jsonPath)) File.Delete(jsonPath);
                }
                else
                {
                    var found = Directory.EnumerateFiles(GetCustomersFolderPath(), MakeFileSystemSafe(c.Hoten) + "*.json").FirstOrDefault();
                    if (!string.IsNullOrEmpty(found) && File.Exists(found)) File.Delete(found);
                }

                var folder = GetProfileFolderPath(c);
                if (Directory.Exists(folder)) Directory.Delete(folder, true);
            }
            catch { }
        }

        // -------------------------
        // Word template extraction and profile creation
        // -------------------------

        private string ExtractEmbeddedTemplateToTemp()
        {
            var asm = Assembly.GetExecutingAssembly();
            var resName = asm.GetManifestResourceNames()
                .FirstOrDefault(n => n.EndsWith(EmbeddedTemplateFileName, StringComparison.OrdinalIgnoreCase)
                                  || n.EndsWith("01HN.docx", StringComparison.OrdinalIgnoreCase)
                                  || (n.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) && n.IndexOf("01", StringComparison.OrdinalIgnoreCase) >= 0));

            if (resName == null)
                throw new FileNotFoundException("Embedded Word template not found. Make sure 01 HN.docx Build Action = Embedded Resource.");

            var temp = Path.Combine(Path.GetTempPath(), "template_" + Guid.NewGuid().ToString("N") + ".docx");
            using (var stream = asm.GetManifestResourceStream(resName))
            {
                if (stream == null) throw new FileNotFoundException("Failed to open embedded template stream.");
                using (var fs = File.OpenWrite(temp))
                {
                    stream.CopyTo(fs);
                }
            }
            return temp;
        }

        private string GetProfileFolderPath(Customer c)
        {
            var desktopRoot = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var root = Path.Combine(desktopRoot, "Hồ sơ NHCS");
            if (!Directory.Exists(root)) Directory.CreateDirectory(root);

            var safeName = MakeFileSystemSafe(c.Hoten);
            var dateSuffix = DateTime.Now.ToString("MM-yyyy");
            var folder = Path.Combine(root, safeName + "_" + dateSuffix);

            // Nếu folder đã tồn tại, sử dụng nó (cho cùng khách hàng trong cùng tháng)
            // Chỉ thêm hậu tố số nếu có NHIỀU khách hàng cùng tên trong cùng tháng
            if (!Directory.Exists(folder))
            {
                // Folder chưa tồn tại, tạo mới
                return folder;
            }
            else
            {
                // Folder tồn tại - kiểm tra xem nó có thuộc về khách hàng này không bằng cách so sánh _fileName
                // Nếu khách hàng có _fileName và folder khớp tồn tại, dùng folder đó
                if (!string.IsNullOrEmpty(c._fileName))
                {
                    // Khách hàng đã tồn tại trong database - dùng folder hiện có
                    return folder;
                }
                else
                {
                    // Khách hàng mới cùng tên - cần thêm hậu tố
                    var baseFolder = folder;
                    int i = 1;
                    while (Directory.Exists(folder))
                    {
                        folder = baseFolder + "_" + i;
                        i++;
                    }
                    return folder;
                }
            }
        }

        private void CreateProfileFromTemplate(Customer c, bool include03)
        {
            var destFolder = GetProfileFolderPath(c);
            Directory.CreateDirectory(destFolder);

            var tempFilesToDelete = new List<string>();

            try
            {
                foreach (var templateName in GetTemplateNamesForCustomer(c, include03))
                {
                    string templatePath;
                    try { templatePath = ResolveTemplatePath(templateName); }
                    catch (FileNotFoundException ex)
                    {
                        MessageBox.Show($"Template \"{templateName}\" không tìm thấy: {ex.Message}", "Template missing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        continue;
                    }

                    if (!IsDocxFile(templatePath))
                    {
                        MessageBox.Show($"Template \"{templateName}\" không hợp lệ hoặc bị hỏng:\n{templatePath}", "Lỗi định dạng template", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        continue;
                    }

                    if (!string.IsNullOrEmpty(templatePath) && templatePath.StartsWith(Path.GetTempPath(), StringComparison.OrdinalIgnoreCase))
                        tempFilesToDelete.Add(templatePath);

                    var shortName = Path.GetFileNameWithoutExtension(templateName).Replace(" ", "_");
                    var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    var destDoc = Path.Combine(destFolder, MakeFileSystemSafe(c.Hoten) + "_" + shortName + "_" + ts + ".docx");

                    File.Copy(templatePath, destDoc, false);

                    if (!IsDocxFile(destDoc))
                    {
                        MessageBox.Show(
                            $"Template \"{templateName}\" sản xuất file không hợp lệ: {destDoc}\n" +
                            "Template có thể bị hỏng hoặc không phải là .docx thật. Vui lòng kiểm tra lại template trong Resources nhúng.",
                            "Lỗi định dạng template",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);

                        try { if (File.Exists(destDoc)) File.Delete(destDoc); } catch { }
                        continue;
                    }

                    // Thay thế chỉ bằng OpenXML
                    ReplacePlaceholdersInWord(destDoc, c);
                }
            }
            finally
            {
                foreach (var f in tempFilesToDelete) { try { if (File.Exists(f)) File.Delete(f); } catch { } }
            }
        }

        private void ReplacePlaceholdersInWord(string docPath, Customer c)
        {

            // Đảm bảo Sotienchu được điền từ giá trị số cho mẫu 01
            try { EnsureSotienchuFromNumeric(c, docPath); } catch { }
             // Sử dụng helper thay thế OpenXML
             var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
             {
                { "{{hoten}}", c.Hoten },
                { "{{socccd}}", c.Socccd },
                { "{{cccd}}", c.Socccd },
                { "{{gioitinh}}", c.GioiTinh },
                { "{{dantoc}}", c.Dantoc },
                { "{{sdt}}", c.Sdt ?? "" },
                { "{我的}", c.Sdt },
                { "{{nhandang}}", c.Nhandang },
                { "{{ngaycap}}", c.Ngaycap == DateTime.MinValue ? "" : c.Ngaycap.ToString("dd/MM/yyyy") },
                { "{{noicap}}", c.Noicap },
                { "{{xa}}", c.Xa },
                { "{{thon}}", c.Thon },
                { "{{hoi}}", c.Hoi },
                { "{{totruong}}", c.Totruong },
                { "{{to}}", c.Totruong },
                { "{{chuongtrinh}}", c.Chuongtrinh },
                { "{{vtc}}", c.Vtc },
                { "{{phuongan}}", c.Phuongan },
                { "{{ngaydenhan}}", c.Ngaydenhan == DateTime.MinValue ? "" : c.Ngaydenhan.ToString("dd/MM/yyyy") },
                { "{{thoihanvay}}", c.Thoihanvay },
                { "{{sotien}}", c.Sotien },
                { "{{sotien1}}", c.Sotien1 },
                { "{{sotien2}}", c.Sotien2 },
                { "{{sotientong}}", c.Sotientong },
                { "{{sotienchu}}", c.Sotienchu },
                { "{{soluong1}}", c.Soluong1 ?? "" },
                { "{{soluong2}}", string.IsNullOrWhiteSpace(c.Soluong2) ? "" : c.Soluong2 },
                // Ưu tiên trường Mucdich nhập tự do; nếu trống, dùng Doituong (combo) làm dự phòng
                { "{{mucdich1}}", !string.IsNullOrWhiteSpace(c.Mucdich1) ? c.Mucdich1 : (c.Doituong1 ?? "") },
                { "{{mucdich2}}", !string.IsNullOrWhiteSpace(c.Mucdich2) ? c.Mucdich2 : (c.Doituong2 ?? "") },
                { "{{doituong1}}", c.Doituong1 ?? "" },
                { "{{doituong2}}", c.Doituong2 ?? "" },
                { "{{doituong}}", !string.IsNullOrWhiteSpace(c.Doituong1) ? c.Doituong1 : (c.Doituong2 ?? "") },
                { "{{ngaylaphs}}", c.Ngaylaphs == DateTime.MinValue ? "" : c.Ngaylaphs.ToString("dd/MM/yyyy") },
                { "{{ngaysinh}}", c.Ngaysinh == DateTime.MinValue ? "" : c.Ngaysinh.ToString("dd/MM/yyyy") },
                { "{{phanky}}", c.Phanky },
                { "{{pgd}}", c.PGD },
                { "{{thoihancccd}}", c.Thoihancccd == DateTime.MinValue ? "" : c.Thoihancccd.ToString("dd/MM/yyyy") },

                // GUQ
                { "{{ntk1}}", c.Ntk1 ?? "" },
                { "{{ntk2}}", c.Ntk2 ?? "" },
                { "{{ntk3}}", c.Ntk3 ?? "" },
                { "{{cccdntk1}}", c.CccdNtk1 ?? "" },
                { "{{cccdntk2}}", c.CccdNtk2 ?? "" },
                { "{{cccdntk3}}", c.CccdNtk3 ?? "" },
                { "{{qh1}}", c.Qh1 ?? "" },
                { "{{qh2}}", c.Qh2 ?? "" },
                { "{{qh3}}", c.Qh3 ?? "" },
                { "{{namsinh1}}", FormatNamsinhStringForDoc(c.Namsinh1, docPath) },
                { "{{namsinh2}}", FormatNamsinhStringForDoc(c.Namsinh2, docPath) },
                { "{{namsinh3}}", FormatNamsinhStringForDoc(c.Namsinh3, docPath) },
                { "{{namsinh1_year}}", ExtractYearString(c.Namsinh1) ?? "" },
                { "{{namsinh2_year}}", ExtractYearString(c.Namsinh2) ?? "" },
                { "{{namsinh3_year}}", ExtractYearString(c.Namsinh3) ?? "" },
                { "{{namsinh}}", ShouldShowFullNamsinh(docPath) ? (c.Ngaysinh == DateTime.MinValue ? "" : c.Ngaysinh.ToString("dd/MM/yyyy")) : (c.Ngaysinh == DateTime.MinValue ? "" : c.Ngaysinh.Year.ToString()) },
             };

            // Đối với mẫu 03 DS đảm bảo các placeholder cụ thể sử dụng giá trị từ khách hàng hiện tại
            try
            {
                if (Is03DS(docPath))
                {
                    // Ánh xạ chính xác như đã chỉ định: txtHoten -> {{hoten}}, cbDoituong -> {{doituong}}, txtPhuongan -> {{phuongan}}, cbThoihanvay -> {{thoihanvay}}
                    replacements["{{hoten}}"] = c.Hoten ?? "";
                    replacements["{{doituong}}"] = c.Doituong1 ?? "";   // từ cbDoituong
                    replacements["{{phuongan}}"] = c.Phuongan ?? "";     // từ txtPhuongan
                    replacements["{{sotien}}"] = c.Sotien ?? "";
                    replacements["{{thoihanvay}}"] = c.Thoihanvay ?? "";

                    // Biến thể bổ sung cho các placeholder 03 DS
                    replacements["{{mucdich}}"] = c.Phuongan ?? "";  // Mục đích có thể dùng tên này
                    replacements["{{mucdichsudungvon}}"] = c.Phuongan ?? "";  // Mục đích sử dụng vốn
                    replacements["{{mucdich1}}"] = c.Phuongan ?? "";
                    replacements["{{doituongthuhuong}}"] = c.Doituong1 ?? "";  // Đối tượng thụ hưởng
                    replacements["{{thuhuong}}"] = c.Doituong1 ?? "";
                    replacements["{{thoihan}}"] = c.Thoihanvay ?? "";  // Thời hạn (không có "vay")

                    // Cũng ánh xạ các placeholder có số (cho các mẫu dùng {{hoten1}}, {{sotien1}}, v.v.)
                    replacements["{{hoten1}}"] = c.Hoten ?? "";
                    replacements["{{doituong1}}"] = c.Doituong1 ?? "";
                    replacements["{{phuongan1}}"] = c.Phuongan ?? "";
                    replacements["{{sotien1}}"] = c.Sotien ?? "";
                    replacements["{{thoihanvay1}}"] = c.Thoihanvay ?? "";
                    replacements["{{mucdich1}}"] = c.Phuongan ?? "";

                    // Cũng ánh xạ placeholder index 0 (cho các mẫu bắt đầu từ 0)
                    replacements["{{hoten0}}"] = c.Hoten ?? "";
                    replacements["{{doituong0}}"] = c.Doituong1 ?? "";
                    replacements["{{phuongan0}}"] = c.Phuongan ?? "";
                    replacements["{{sotien0}}"] = c.Sotien ?? "";
                    replacements["{{thoihanvay0}}"] = c.Thoihanvay ?? "";
                    replacements["{{mucdich0}}"] = c.Phuongan ?? "";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR in 03 DS mapping: {ex.Message}");
            }

            try { ReplacePlaceholdersUsingOpenXml(docPath, replacements, c); }
            catch (Exception ex) { MessageBox.Show("Error replacing placeholders (OpenXML): " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
         }

        // Thử tính toán `Sotienchu` từ giá trị số `Sotien` (hoặc `Sotientong`) khi tạo mẫu
        private void EnsureSotienchuFromNumeric(Customer c, string docPath)
        {
            try
            {
                if (c == null) return;
                // Chỉ áp dụng cho mẫu 01 (embedded hoặc tên file chứa '01')
                var name = Path.GetFileName(docPath) ?? "";
                if (name.IndexOf("01", StringComparison.OrdinalIgnoreCase) < 0 && !string.Equals(name, EmbeddedTemplateFileName, StringComparison.OrdinalIgnoreCase))
                    return;

                var source = !string.IsNullOrWhiteSpace(c.Sotien) ? c.Sotien : (!string.IsNullOrWhiteSpace(c.Sotientong) ? c.Sotientong : "");
                if (string.IsNullOrWhiteSpace(source)) return;
                var digits = new string((source ?? "").Where(char.IsDigit).ToArray());
                if (string.IsNullOrWhiteSpace(digits)) return;
                if (!long.TryParse(digits, out long value)) return;
                if (value <= 0) return;
                var words = NumberToVietnameseWords(value);
                if (!string.IsNullOrWhiteSpace(words)) c.Sotienchu = words + " đồng";
            }
            catch { }
        }

        // Phân tích chuỗi số thuần túy với dấu chấm/phẩy đã bị xóa trong caller; wrapper giữ lại cho rõ ràng
        private long ParseMoneyStringToLong(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0;
            try
            {
                var digits = new string(s.Where(char.IsDigit).ToArray());
                if (long.TryParse(digits, out long v)) return v;
            }
            catch { }
            return 0;
        }

        // Chuyển đổi số nguyên thành chữ tiếng Việt (hỗ trợ đến tập tỷ một cách hợp lý)
        private string NumberToVietnameseWords(long number)
        {
            if (number == 0) return "không";

            string[] unitNames = { "", "nghìn", "triệu", "tỷ" };
            string[] digits = { "không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín" };

            var groups = new List<int>();
            long n = number;
            while (n > 0)
            {
                groups.Add((int)(n % 1000));
                n /= 1000;
            }

            var parts = new List<string>();
            for (int i = groups.Count - 1; i >= 0; i--)
            {
                var g = groups[i];
                if (g == 0)
                {
                    // Vẫn cần thêm đơn vị khi có các nhóm bên trong (giữ vị trí) chỉ nếu có nhóm thấp hơn khác 0
                    bool lowerNonZero = groups.Take(i).Any(x => x != 0);
                    if (lowerNonZero && i > 0) parts.Add(unitNames[i]);
                    continue;
                }

                int hundreds = g / 100;
                int tens = (g / 10) % 10;
                int units = g % 10;

                var seg = new List<string>();
                if (hundreds > 0)
                {
                    seg.Add(digits[hundreds] + " trăm");
                }
                else
                {
                    // Nếu hàng trăm == 0 và có chục/đơn vị, khi nhóm này không phải cao nhất, có thể cần 'không trăm'
                    if ((tens > 0 || units > 0) && i != groups.Count - 1) seg.Add("không trăm");
                }

                if (tens > 0)
                {
                    if (tens == 1) seg.Add("mười");
                    else seg.Add(digits[tens] + " mươi");
                }
                else
                {
                    if (units > 0) seg.Add("lẻ");
                }

                if (units > 0)
                {
                    string unitWord = digits[units];
                    if (tens >= 1)
                    {
                        if (units == 1) unitWord = "mốt";
                        else if (units == 5) unitWord = "lăm";
                    }
                    seg.Add(unitWord);
                }

                var segText = string.Join(" ", seg.Where(x => !string.IsNullOrWhiteSpace(x)));
                if (!string.IsNullOrWhiteSpace(segText))
                {
                    if (i < unitNames.Length && !string.IsNullOrWhiteSpace(unitNames[i]))
                        parts.Add(segText + " " + unitNames[i]);
                    else parts.Add(segText);
                }
            }

            var result = string.Join(" ", parts).Trim();
            // Dọc dẹp nhiều khoảng trắng
            result = Regex.Replace(result, @"\s+", " ").Trim();
            // Viết thường
            return result;
        }

        private void ReplacePlaceholdersInWordForGroup(string docPath, Customer c, string entriesText)
        {
            var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "{{hoten}}", c.Hoten },
                { "{{socccd}}", c.Socccd },
                { "{{cccd}}", c.Socccd },
                { "{{gioitinh}}", c.GioiTinh },
                { "{{dantoc}}", c.Dantoc },
                { "{ νέα }", c.Sdt },
                { "{{nhandang}}", c.Nhandang },
                { "{{ngaycap}}", c.Ngaycap == DateTime.MinValue ? "" : c.Ngaycap.ToString("dd/MM/yyyy") },
                { "{{noicap}}", c.Noicap },
                { "{{xa}}", c.Xa },
                { "{{thon}}", c.Thon },
                { "{{hoi}}", c.Hoi },
                { "{{totruong}}", c.Totruong },
                { "{{to}}", c.Totruong },
                { "{{chuongtrinh}}", c.Chuongtrinh },
                { "{{03_ENTRIES}}", entriesText ?? "" }
            };

            try { ReplacePlaceholdersUsingOpenXml(docPath, replacements, c, isGroup: true); }
            catch (Exception ex) { MessageBox.Show("Error replacing placeholders for group (OpenXML): " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        // Gộp các text run kề nhau trong cùng paragraph để ngăn phân tách placeholder
        private void MergeAdjacentTextRuns(OpenXmlPart part)
        {
            try
            {
                foreach (var paragraph in part.RootElement.Descendants<Paragraph>())
                {
                    var runs = paragraph.Elements<Run>().ToList();
                    for (int i = 0; i < runs.Count - 1; i++)
                    {
                        var currentRun = runs[i];
                        var nextRun = runs[i + 1];

                        var currentText = currentRun.GetFirstChild<Text>();
                        var nextText = nextRun.GetFirstChild<Text>();

                        if (currentText != null && nextText != null)
                        {
                            // Gộp text tiếp theo vào text hiện tại
                            currentText.Text += nextText.Text;

                            // Xóa run tiếp theo
                            nextRun.Remove();
                            runs.RemoveAt(i + 1);
                            i--; // Kiểm tra lại vị trí hiện tại
                        }
                    }
                }
                part.RootElement.Save();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error merging text runs: {ex.Message}");
            }
        }

        // Thay thế đơn giản cho các ô trong bảng - thay thế text trong từng text node riêng lẻ không gộp
        private void ReplaceInTableCells(OpenXmlPart part, Dictionary<string, string> replacements)
        {
            try
            {
                var tables = part.RootElement.Descendants<Table>().ToList();
                System.Diagnostics.Debug.WriteLine($"Found {tables.Count} tables");

                foreach (var table in tables)
                {
                    int rowIndex = 0;
                    foreach (var row in table.Descendants<TableRow>())
                    {
                        int cellIndex = 0;
                        foreach (var cell in row.Descendants<TableCell>())
                        {
                            // Lấy tất cả text trong ô này chỉ để ghi log
                            var textNodes = cell.Descendants<Text>().ToList();
                            if (textNodes.Count == 0) { cellIndex++; continue; }

                            var cellText = string.Concat(textNodes.Select(t => t.Text ?? ""));

                            // Ghi log mọi ô không rỗng để xem có gì trong mẫu
                            if (!string.IsNullOrWhiteSpace(cellText))
                            {
                                System.Diagnostics.Debug.WriteLine($"[Row {rowIndex}, Cell {cellIndex}]: '{cellText}'");
                            }

                            // Thay thế trong từng text node riêng lẻ (KHÔNG gộp hoặc xóa)
                            foreach (var textNode in textNodes)
                            {
                                if (string.IsNullOrEmpty(textNode.Text)) continue;

                                string originalText = textNode.Text;
                                string newText = originalText;

                                // Thay thế tất cả các placeholder
                                foreach (var kv in replacements)
                                {
                                    if (newText.IndexOf(kv.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"  ✅ Found '{kv.Key}' in cell [{rowIndex},{cellIndex}]");
                                        newText = ReplaceIgnoreCase(newText, kv.Key, kv.Value ?? "");
                                    }
                                }

                                // Chỉ cập nhật nếu có thay đổi
                                if (newText != originalText)
                                {
                                    System.Diagnostics.Debug.WriteLine($"  ➡ Text node changed: '{originalText}' → '{newText}'");
                                    textNode.Text = newText;
                                }
                            }
                            cellIndex++;
                        }
                        rowIndex++;
                    }
                }
                part.RootElement.Save();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ReplaceInTableCells: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"Lỗi trong ReplaceInTableCells:\n\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Thay thế đơn giản cho các paragraph - thay thế trong từng text node riêng lẻ
        private void ReplaceInParagraphs(OpenXmlPart part, Dictionary<string, string> replacements)
        {
            try
            {
                foreach (var para in part.RootElement.Descendants<Paragraph>())
                {
                    var textNodes = para.Descendants<Text>().ToList();
                    if (textNodes.Count == 0) continue;

                    // Thay thế trong từng text node riêng lẻ (KHÔNG gộp)
                    foreach (var textNode in textNodes)
                    {
                        if (string.IsNullOrEmpty(textNode.Text)) continue;

                        string originalText = textNode.Text;
                        string newText = originalText;

                        // Replace all placeholders
                        foreach (var kv in replacements)
                        {
                            if (newText.IndexOf(kv.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                System.Diagnostics.Debug.WriteLine($"Paragraph: Found '{kv.Key}' in '{originalText}'");
                                newText = ReplaceIgnoreCase(newText, kv.Key, kv.Value ?? "");
                            }
                        }

                        // Only update if changed
                        if (newText != originalText)
                        {
                            System.Diagnostics.Debug.WriteLine($"Paragraph: Changed '{originalText}' → '{newText}'");
                            textNode.Text = newText;
                        }
                    }
                }
                part.RootElement.Save();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ReplaceInParagraphs: {ex.Message}\n{ex.StackTrace}");
            }
        }

        private void ReplacePlaceholdersUsingOpenXml(string docPath, Dictionary<string, string> replacements, Customer c = null, bool isGroup = false)
        {
            if (string.IsNullOrWhiteSpace(docPath) || !File.Exists(docPath)) throw new FileNotFoundException("Document not found: " + docPath);

            using (var wordDoc = WordprocessingDocument.Open(docPath, true))
            {
                var mainPart = wordDoc.MainDocumentPart;
                if (mainPart == null) return;

                // Đối với mẫu 03 DS, KHÔNG xóa bất kỳ dòng nào, chỉ thay thế placeholder
                // Mẫu sẽ được điền nguyên trạng mà không thay đổi cấu trúc bảng

                try
                {
                    if (ShouldShowFullNamsinh(docPath) && c != null)
                    {
                        // Đối với mẫu GUQ: Điền các placeholder địa chỉ ({{thon}}, {{xa}}, {{hoi}}) Ở MỌI NƠI
                        // Giữ chúng trong dictionary replacements để chúng được điền toàn cục

                        // Xử lý các dòng với placeholder NTK - chỉ điền địa chỉ nếu NTK có dữ liệu
                        var addressPlaceholders = new[] { "{{thon}}", "{{xa}}", "{{hoi}}" };

                        foreach (var table in mainPart.Document.Descendants<Table>())
                        {
                            try
                            {
                                foreach (var row in table.Elements<TableRow>())
                                {
                                    var rowText = string.Concat(row.Descendants<Text>().Select(t => t.Text ?? ""));

                                    // Kiểm tra xem dòng này có chứa placeholder NTK nào không
                                    bool hasNtk1 = rowText.IndexOf("{{ntk1}}", StringComparison.OrdinalIgnoreCase) >= 0;
                                    bool hasNtk2 = rowText.IndexOf("{{ntk2}}", StringComparison.OrdinalIgnoreCase) >= 0;
                                    bool hasNtk3 = rowText.IndexOf("{{ntk3}}", StringComparison.OrdinalIgnoreCase) >= 0;

                                    // Nếu dòng có placeholder NTK nhưng không có dữ liệu, xóa các placeholder địa chỉ chỉ TRONG DÒNG NÀY
                                    if ((hasNtk1 && string.IsNullOrWhiteSpace(c.Ntk1)) ||
                                        (hasNtk2 && string.IsNullOrWhiteSpace(c.Ntk2)) ||
                                        (hasNtk3 && string.IsNullOrWhiteSpace(c.Ntk3)))
                                    {
                                        // Xóa địa chỉ trong dòng NTK rỗng này
                                        foreach (var text in row.Descendants<Text>())
                                        {
                                            if (text.Text == null) continue;
                                            foreach (var ph in addressPlaceholders)
                                            {
                                                if (text.Text.IndexOf(ph, StringComparison.OrdinalIgnoreCase) >= 0)
                                                {
                                                    text.Text = ReplaceIgnoreCase(text.Text, ph, "");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            catch { }
                        }
                        mainPart.Document.Save();
                    }
                }
                catch { }

                // Đối với mẫu 01 SXKD: tách gộp các ô chứa placeholder mucdich/soluong/sotien
                try
                {
                    if (Is01SXKD(docPath))
                    {
                        foreach (var table in mainPart.Document.Descendants<Table>())
                        {
                            try
                            {
                                var rows = table.Elements<TableRow>().ToList();
                                bool foundHeaderRow = false;

                                // Tìm dòng tiêu đề chứa "Đối tượng" và "Thành tiền"
                                for (int i = 0; i < rows.Count; i++)
                                {
                                    var rowText = string.Concat(rows[i].Descendants<Text>().Select(t => t.Text ?? ""));

                                    if ((rowText.IndexOf("Đối tượng", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                         rowText.IndexOf("Doi tuong", StringComparison.OrdinalIgnoreCase) >= 0) &&
                                        (rowText.IndexOf("Thành tiền", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                         rowText.IndexOf("Thanh tien", StringComparison.OrdinalIgnoreCase) >= 0))
                                    {
                                        foundHeaderRow = true;

                                        // Tách gộp tất cả các ô từ dòng này và 5 dòng tiếp theo (để bao gồm tất cả các dòng dữ liệu)
                                        for (int j = i; j < Math.Min(i + 6, rows.Count); j++)
                                        {
                                            foreach (var cell in rows[j].Elements<TableCell>())
                                            {
                                                var tcPr = cell.GetFirstChild<TableCellProperties>();
                                                if (tcPr != null)
                                                {
                                                    // Xóa gộp dọc
                                                    var vMerge = tcPr.GetFirstChild<VerticalMerge>();
                                                    if (vMerge != null)
                                                    {
                                                        vMerge.Remove();
                                                    }

                                                    // Xóa grid span
                                                    var gridSpan = tcPr.GetFirstChild<GridSpan>();
                                                    if (gridSpan != null)
                                                    {
                                                        gridSpan.Remove();
                                                    }
                                                }
                                            }
                                        }
                                        break;
                                    }
                                }

                                // Nếu không tìm thấy header, thử tìm các dòng có placeholder và tách gộp chúng
                                if (!foundHeaderRow)
                                {
                                    foreach (var row in rows)
                                    {
                                        var rowText = string.Concat(row.Descendants<Text>().Select(t => t.Text ?? ""));
                                        if (rowText.IndexOf("{{mucdich", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                            rowText.IndexOf("{{soluong", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                            rowText.IndexOf("{{sotien", StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            foreach (var cell in row.Elements<TableCell>())
                                            {
                                                var tcPr = cell.GetFirstChild<TableCellProperties>();
                                                if (tcPr != null)
                                                {
                                                    var vMerge = tcPr.GetFirstChild<VerticalMerge>();
                                                    if (vMerge != null) vMerge.Remove();

                                                    var gridSpan = tcPr.GetFirstChild<GridSpan>();
                                                    if (gridSpan != null) gridSpan.Remove();
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            catch { }
                        }
                        mainPart.Document.Save();
                    }
                }
                catch { }

                // Các phần cần xử lý: main, headers, footers, footnotes, endnotes, comments
                var parts = new List<OpenXmlPart>();
                parts.Add(mainPart);
                parts.AddRange(mainPart.HeaderParts);
                parts.AddRange(mainPart.FooterParts);
                if (mainPart.FootnotesPart != null) parts.Add(mainPart.FootnotesPart);
                if (mainPart.EndnotesPart != null) parts.Add(mainPart.EndnotesPart);
                if (mainPart.WordprocessingCommentsPart != null) parts.Add(mainPart.WordprocessingCommentsPart);

                foreach (var part in parts.Distinct())
                {
                    try
                    {
                        System.Diagnostics.Debug.WriteLine($"Processing part: {part.GetType().Name}");

                        // Cách tiếp cận thứ nhất: thay thế ô bảng đơn giản (không gộp)
                        ReplaceInTableCells(part, replacements);

                        // Cách tiếp cận thứ hai: thay thế paragraph đơn giản (không gộp)
                        ReplaceInParagraphs(part, replacements);

                        // Cách tiếp cận thứ ba: thử thay thế across-runs (cho các placeholder bị tách)
                        // Lưu ý: Vẫn có thể gây vấn đề khoảng trắng, nên làm cuối cùng
                        TryReplacePlaceholdersAcrossRuns(part, replacements);

                        var texts = part.RootElement.Descendants<Text>();
                        foreach (var t in texts)
                        {
                            if (string.IsNullOrEmpty(t.Text)) continue;
                            string newText = t.Text;
                            foreach (var kv in replacements)
                            {
                                if (newText.IndexOf(kv.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Found placeholder '{kv.Key}' in text: '{t.Text}'");
                                    newText = ReplaceIgnoreCase(newText, kv.Key, kv.Value ?? "");
                                    System.Diagnostics.Debug.WriteLine($"Replaced with: '{newText}'");
                                }
                            }
                            if (newText != t.Text) t.Text = newText;
                        }
                        part.RootElement.Save();
                    }
                    catch { }
                }

                // Cho mẫu 03 DS: Sửa lỗi nối chuỗi giữa ô STT (số) và ô tiếp theo bắt đầu bằng chữ
                try
                {
                    if (Is03DS(docPath))
                    {
                        foreach (var table in mainPart.Document.Descendants<Table>())
                        {
                            try
                            {
                                foreach (var row in table.Elements<TableRow>())
                                {
                                    var cells = row.Elements<TableCell>().ToList();
                                    for (int ci = 0; ci + 1 < cells.Count; ci++)
                                    {
                                        try
                                        {
                                            var left = string.Concat(cells[ci].Descendants<Text>().Select(t => t.Text ?? ""));
                                            var rightFirstNode = cells[ci + 1].Descendants<Text>().FirstOrDefault();
                                            var right = rightFirstNode != null ? (rightFirstNode.Text ?? "") : "";
                                            if (!string.IsNullOrEmpty(left) && !string.IsNullOrEmpty(right))
                                            {
                                                var lastChar = left[left.Length - 1];
                                                var firstChar = right[0];
                                                if (char.IsDigit(lastChar) && char.IsLetter(firstChar) && !char.IsWhiteSpace(firstChar))
                                                {
                                                    // Thêm khoảng trắng vào đầu text node đầu tiên của ô bên phải
                                                    rightFirstNode.Text = " " + rightFirstNode.Text;
                                                }
                                            }
                                        }
                                        catch { }
                                    }
                                }
                            }
                            catch { }
                        }
                        mainPart.Document.Save();
                    }
                }
                catch { }

                // Lưu tài liệu chính
                mainPart.Document.Save();
            }
        }

        // Thử tìm và thay thế các placeholder bị tách qua nhiều Text node (runs)
        private void TryReplacePlaceholdersAcrossRuns(OpenXmlPart part, Dictionary<string, string> replacements)
        {
            if (part == null || replacements == null || replacements.Count == 0) return;
            try
            {
                var texts = part.RootElement.Descendants<Text>().ToList();
                if (texts.Count == 0) return;

                foreach (var kv in replacements)
                {
                    var rawKey = kv.Key ?? string.Empty;
                    var replacement = kv.Value ?? string.Empty;
                    if (string.IsNullOrEmpty(rawKey)) continue;

                    // Chuẩn hóa token key: nếu key giống như "{{name}}" thì trích xuất token bên trong "name"
                    string token = rawKey;
                    var m = Regex.Match(rawKey, "^\\s*\\{\\{\\s*(.*?)\\s*\\}\\}\\s*$");
                    if (m.Success && m.Groups.Count > 1) token = m.Groups[1].Value;
                    if (string.IsNullOrEmpty(token)) continue;

                    // Xây dựng regex để tìm placeholder cho phép khoảng trắng tùy chọn bên trong dấu ngoặc nhọn
                    var pattern = "\\{\\{\\s*" + Regex.Escape(token) + "\\s*\\}\\}";

                    int i = 0;
                    while (i < texts.Count)
                    {
                        var sb = new StringBuilder();
                        int j = i;
                        // Xây dựng lên đến một cửa sổ hợp lý
                        while (j < texts.Count && sb.Length < token.Length + 1024 && (j - i) < 200)
                        {
                            sb.Append(texts[j].Text ?? string.Empty);
                            j++;
                        }

                        if (sb.Length == 0) { i++; continue; }
                        var combined = sb.ToString();

                        var match = Regex.Match(combined, pattern, RegexOptions.IgnoreCase);
                        if (!match.Success)
                        {
                            i++;
                            continue;
                        }

                        int pos = match.Index;
                        int matchLen = match.Length;

                        int remaining = pos;
                        int startNode = i; int startOffset = 0;
                        for (int k = i; k < j; k++)
                        {
                            var tlen = (texts[k].Text ?? string.Empty).Length;
                            if (remaining <= tlen)
                            {
                                startNode = k;
                                startOffset = remaining;
                                break;
                            }
                            remaining -= tlen;
                        }

                        int endPos = pos + matchLen;
                        remaining = endPos;
                        int endNode = i; int endOffset = 0;
                        for (int k = i; k < j; k++)
                        {
                            var tlen = (texts[k].Text ?? string.Empty).Length;
                            if (remaining <= tlen)
                            {
                                endNode = k;
                                endOffset = remaining;
                                break;
                            }
                            remaining -= tlen;
                        }

                        try
                        {
                            var startText = texts[startNode].Text ?? string.Empty;
                            var endText = texts[endNode].Text ?? string.Empty;

                            var prefix = startText.Substring(0, startOffset);
                            var suffix = endText.Substring(endOffset);

                            // Đảm bảo không vô tình nối nội dung lân cận mà không có khoảng trắng.
                            // Nếu prefix kết thúc bằng không-khoảng-trắng và replacement bắt đầu bằng không-khoảng-trắng, thêm khoảng trắng.
                            var finalReplacement = replacement ?? string.Empty;
                            if (!string.IsNullOrEmpty(prefix) && !char.IsWhiteSpace(prefix[prefix.Length - 1]) && finalReplacement.Length > 0 && !char.IsWhiteSpace(finalReplacement[0]))
                                finalReplacement = " " + finalReplacement;

                            // Nếu replacement kết thúc bằng không-khoảng-trắng và suffix bắt đầu bằng không-khoảng-trắng, chèn khoảng trắng sau replacement
                            if (finalReplacement.Length > 0 && !char.IsWhiteSpace(finalReplacement[finalReplacement.Length - 1]) && !string.IsNullOrEmpty(suffix) && !char.IsWhiteSpace(suffix[0]))
                                finalReplacement = finalReplacement + " ";

                            texts[startNode].Text = prefix + finalReplacement + suffix;

                            for (int k = startNode + 1; k <= endNode; k++)
                            {
                                texts[k].Text = string.Empty;
                            }

                            // Làm mới collection texts và tiếp tục sau node đã thay thế
                            texts = part.RootElement.Descendants<Text>().ToList();
                            i = Math.Min(texts.Count, startNode + 1);
                        }
                        catch
                        {
                            i++;
                        }
                    }
                }

                part.RootElement.Save();
            }
            catch { }
        }

        private int IndexOfIgnoreCase(string source, string value)
        {
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(value)) return -1;
            return source.IndexOf(value, StringComparison.OrdinalIgnoreCase);
        }

        private string ReplaceIgnoreCase(string input, string oldValue, string newValue)
        {
            if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(oldValue)) return input;
            try { 
                // Nếu oldValue là placeholder ngoặc nhọn kép như {{name}}, cho phép khoảng trắng tùy chọn bên trong ngoặc nhọn trong tài liệu
                var m = Regex.Match(oldValue, "^\\s*\\{\\{\\s*(.*?)\\s*\\}\\}\\s*$");
                if (m.Success && m.Groups.Count > 1)
                {
                    var token = m.Groups[1].Value;
                    var pattern = "\\{\\{\\s*" + Regex.Escape(token) + "\\s*\\}\\}";
                    return Regex.Replace(input, pattern, newValue ?? "", RegexOptions.IgnoreCase);
                }

                return Regex.Replace(input, Regex.Escape(oldValue), newValue ?? "", RegexOptions.IgnoreCase);
            }
            catch { return input.Replace(oldValue, newValue ?? ""); }
        }

        // Designer mong đợi tên handler viết hoa chính xác này ở nhiều nơi
        private void cbPGD_SelectedIndexChanged(object sender, EventArgs e)
        {
            CbPGD_SelectedIndexChanged(sender, e);
        }

        // Các stub của Designer được tham chiếu trong Designer.cs
        private void richTextBox1_TextChanged(object sender, EventArgs e) { }
        private void Form1_Load(object sender, EventArgs e) { }

         // -------------------------
         // Xử lý UI / CRUD
         // -------------------------

         private void Dgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
         {
             // Không làm gì: cột checkbox đã bị xóa; giữ handler để tránh lỗi designer
         }

         private bool IsSelectionCompatibleWithChecked(int rowIndex)
         {
             try
             {
                 var candidate = dgv.Rows[rowIndex].DataBoundItem as Customer;
                 if (candidate == null) return true;

                 var checkedCustomers = new List<Customer>();
                 foreach (DataGridViewRow row in dgv.Rows)
                 {
                     if (row.Index == rowIndex) continue;
                     try
                     {
                         var ccell = row.Cells["colSelect"];
                         bool checkedVal = false;
                         if (ccell != null && ccell.Value != null)
                         {
                             if (ccell.Value is bool bb) checkedVal = bb;
                             else bool.TryParse(ccell.Value.ToString(), out checkedVal);
                         }
                         if (checkedVal) { var item = row.DataBoundItem as Customer; if (item != null) checkedCustomers.Add(item); }
                     }
                     catch { }
                 }

                 if (checkedCustomers.Count == 0) return true;

                 foreach (var other in checkedCustomers)
                 {
                     if (!string.Equals((candidate.PGD ?? "").Trim(), (other.PGD ?? "").Trim(), StringComparison.OrdinalIgnoreCase))
                     {
                         MessageBox.Show("Selected customer does not match PGD of already selected ones.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                         return false;
                     }

                     if (!string.Equals((candidate.Chuongtrinh ?? "").Trim(), (other.Chuongtrinh ?? "").Trim(), StringComparison.OrdinalIgnoreCase))
                     {
                         MessageBox.Show("Selected customer does not match Chương trình of already selected ones.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                         return false;
                     }

                     var candTo = (candidate.Totruong ?? candidate.To ?? "").Trim();
                     var otherTo = (other.Totruong ?? other.To ?? "").Trim();
                     if (!string.Equals(candTo, otherTo, StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(candTo))
                     {
                         MessageBox.Show("Selected customer must have the same Tổ / Tổ trưởng as other selected ones.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                         return false;
                     }

                     if (!string.Equals((candidate.Xa ?? "").Trim(), (other.Xa ?? "").Trim(), StringComparison.OrdinalIgnoreCase))
                     {
                         MessageBox.Show("Selected customer must be in the same Xã as other selected ones.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                         return false;
                     }

                     if (!string.Equals((candidate.Thon ?? "").Trim(), (other.Thon ?? "").Trim(), StringComparison.OrdinalIgnoreCase))
                     {
                         MessageBox.Show("Selected customer must be in the same Thôn as other selected ones.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                         return false;
                     }
                 }

                 return true;
             }
             catch { return true; }
         }

         private void Dgv_CellClick(object sender, DataGridViewCellEventArgs e)
         {
             if (e.RowIndex >= 0 && e.RowIndex < customers.Count)
             {
                 if (e.ColumnIndex >= 0 && dgv.Columns[e.ColumnIndex].Name == "colSelect") return;
                 editingIndex = e.RowIndex;
                 PopulateForm(customers[e.RowIndex]);
             }
         }

         private void Dgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
         {
             if (e.RowIndex >= 0 && e.RowIndex < customers.Count)
             {
                 if (e.ColumnIndex >= 0 && dgv.Columns[e.ColumnIndex].Name == "colSelect") return;
                 editingIndex = e.RowIndex;
                 PopulateForm(customers[e.RowIndex]);
             }
         }

         private async void BtnSave_Click(object sender, EventArgs e)
         {
             var customer = ReadForm();
             if (string.IsNullOrWhiteSpace(customer.Hoten))
             {
                 MessageBox.Show("Vui lòng nhập Họ và tên.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                 return;
             }

             try
             {
                 string existingFile = null;
                 if (customers != null && editingIndex >= 0 && editingIndex < customers.Count)
                 {
                     try { existingFile = customers[editingIndex]._fileName; } catch { existingFile = null; }
                 }

                 if (!string.IsNullOrEmpty(existingFile)) customer._fileName = existingFile;
                 else customer._fileName = customer._fileName; // leave as-is (null or set)

                 SaveCustomerToFile(customer);
                 // btn01 should export only template 01 (no 03, no GUQ)
                 await Task.Run(() => CreateProfileFromTemplate(customer, include03: false));

                 UpsertCustomerInList(customer);

                 BindGrid();
                 ClearForm();

                 // Hiển thị thông báo xuất mẫu 01/TD thành công
                 MessageBox.Show(
                     "✅ Đã xuất mẫu 01/TD thành công!\n\n" +
                     $"📄 Khách hàng: {customer.Hoten}\n" +
                     $"💼 Chương trình: {customer.Chuongtrinh}\n" +
                     $"💵 Số tiền: {customer.Sotien}",
                     "✅ Xuất mẫu 01/TD",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Information
                 );
             }
             catch (Exception ex)
             {
                 // Hiển thị thông báo lỗi với icon
                 MessageBox.Show(
                     $"❌ Lỗi khi xuất mẫu 01/TD:\n\n{ex.Message}",
                     "❌ Lỗi",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Error
                 );
             }
         }

         private async void Btn03_Click(object sender, EventArgs e)
         {
             var customer = ReadForm();
             if (string.IsNullOrWhiteSpace(customer.Hoten))
             {
                 MessageBox.Show("Vui lòng nhập Họ và tên.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                 return;
             }

             try
             {
                 string existingFile = null;
                 if (customers != null && editingIndex >= 0 && editingIndex < customers.Count)
                 {
                     try { existingFile = customers[editingIndex]._fileName; } catch { existingFile = null; }
                 }

                 if (!string.IsNullOrEmpty(existingFile)) customer._fileName = existingFile;

                 SaveCustomerToFile(customer);
                 UpsertCustomerInList(customer);

                 await Task.Run(() =>
                 {
                     // Btn03 should export only 03 DS template
                     ExportSpecificTemplate(customer, "03 DS.docx");
                 });

                 BindGrid();
                 ClearForm();
                 MessageBox.Show("Export (03 DS) thành công.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
             }
             catch (Exception ex)
             {
                 MessageBox.Show("Lỗi khi export 03: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         private async void BtnGUQ_Click(object sender, EventArgs e)
         {
             var customer = ReadForm();
             if (string.IsNullOrWhiteSpace(customer.Hoten))
             {
                 MessageBox.Show("Vui lòng nhập Họ và tên.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                 return;
             }

             try
             {
                 string existingFile = null;
                 if (customers != null && editingIndex >= 0 && editingIndex < customers.Count)
                 {
                     try { existingFile = customers[editingIndex]._fileName; } catch { existingFile = null; }
                 }

                 if (!string.IsNullOrEmpty(existingFile)) customer._fileName = existingFile;

                 SaveCustomerToFile(customer);
                 UpsertCustomerInList(customer);

                 await Task.Run(() =>
                 {
                     // BtnGUQ should export only GUQ template
                     ExportSpecificTemplate(customer, "GUQ.docx");
                 });

                 BindGrid();
                 ClearForm();
                 MessageBox.Show("Export GUQ thành công.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
             }
             catch (Exception ex)
             {
                 MessageBox.Show("Lỗi khi export GUQ: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         private void BtnDelete_Click(object sender, EventArgs e)
         {
             if (dgv == null || dgv.SelectedRows.Count == 0)
             {
                 MessageBox.Show("Chọn khách để xoá.", "Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                 return;
             }

             var idx = dgv.SelectedRows[0].Index;
             if (idx < 0 || customers == null || idx >= customers.Count) return;

             var c = customers[idx];
             var r = MessageBox.Show($"Xoá khách \"{c.Hoten}\"? (file JSON và folder hồ sơ sẽ bị xoá)", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
             if (r != DialogResult.Yes) return;

             try
             {
                 DeleteCustomerFiles(c);
                 customers.RemoveAt(idx);
                 BindGrid();
                 ClearForm();
             }
             catch (Exception ex)
             {
                 MessageBox.Show("Lỗi khi xóa: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         // Upsert khách hàng vào danh sách: cập nhật nếu đang sửa, thêm nếu mới
         private void UpsertCustomerInList(Customer customer)
         {
             if (customer == null) return;

             try
             {
                 if (editingIndex >= 0 && editingIndex < customers.Count)
                 {
                     // Cập nhật khách hàng hiện có
                     customer._fileName = customers[editingIndex]._fileName; // Giữ lại tên file
                     customers[editingIndex] = customer;
                 }
                 else
                 {
                     // Thêm khách hàng mới
                     customers.Add(customer);
                 }

                 // Reset chỉ số editing sau khi upsert
                 editingIndex = -1;
             }
             catch (Exception ex)
             {
                 MessageBox.Show($"Lỗi khi cập nhật danh sách: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         private Customer ReadForm()
         {
             string ntk1 = "", ntk2 = "", ntk3 = "";
             string cccdntk1 = "", cccdntk2 = "", cccdntk3 = "";
             string namsinh1 = "", namsinh2 = "", namsinh3 = "";
             string qh1 = "", qh2 = "", qh3 = "";

             try { if (txtntk1 != null) ntk1 = txtntk1.Text.Trim(); } catch { }
             try { if (txtntk2 != null) ntk2 = txtntk2.Text.Trim(); } catch { }
             try { if (txtntk3 != null) ntk3 = txtntk3.Text.Trim(); } catch { }

             try { if (txtcccd1 != null) cccdntk1 = txtcccd1.Text.Trim(); } catch { }
             try { if (txtcccd2 != null) cccdntk2 = txtcccd2.Text.Trim(); } catch { }
             try { if (txtcccd3 != null) cccdntk3 = txtcccd3.Text.Trim(); } catch { }

             // Đọc ngày sinh NTK từ DateTimePicker (nếu được checked)
             // Lưu thành string dd/MM/yyyy để dễ lưu JSON và sửa sau này
             try 
             { 
                 if (datentk1 != null) 
                 {
                     if (datentk1.Checked)  // Nếu checkbox được tích = có giá trị
                         namsinh1 = datentk1.Value.ToString("dd/MM/yyyy");  // Format ngày thành string
                     else
                         namsinh1 = "";  // Không tích = không nhập
                 }
             } catch { }
             try 
             { 
                 if (datentk2 != null) 
                 {
                     if (datentk2.Checked)
                         namsinh2 = datentk2.Value.ToString("dd/MM/yyyy");
                     else
                         namsinh2 = "";
                 }
             } catch { }
             try 
             { 
                 if (datentk3 != null) 
                 {
                     if (datentk3.Checked)
                         namsinh3 = datentk3.Value.ToString("dd/MM/yyyy");
                     else
                         namsinh3 = "";
                 }
             } catch { }

             try { if (cbqh1 != null) qh1 = cbqh1.Text.Trim(); } catch { }
             try { if (cbqh2 != null) qh2 = cbqh2.Text.Trim(); } catch { }
             try { if (cbqh3 != null) qh3 = cbqh3.Text.Trim(); } catch { }

             // Xác thực các ngày không được trong tương lai
             DateTime ngaycap = dateNgaycapCCCD.Value.Date;
             if (ngaycap > DateTime.Today) ngaycap = DateTime.Today;

             DateTime ngaysinh = DateTime.MinValue;
             if (dateNgaysinh != null)
             {
                 ngaysinh = dateNgaysinh.Value.Date;
                 if (ngaysinh > DateTime.Today) ngaysinh = DateTime.Today;
             }

             DateTime ngaylaphs = dateLaphs.Value.Date;
             if (ngaylaphs > DateTime.Today) ngaylaphs = DateTime.Today;

             DateTime ngaydenhan = DateTime.MinValue;
             if (dateDH != null && dateDH.Format != DateTimePickerFormat.Custom)
             {
                 ngaydenhan = dateDH.Value.Date;
                 if (ngaydenhan > DateTime.Today) ngaydenhan = DateTime.Today;
             }

             DateTime thoihancccd = DateTime.MinValue;
             if (datendhcccd != null && datendhcccd.Format != DateTimePickerFormat.Custom)
             {
                 thoihancccd = datendhcccd.Value.Date;
                 if (thoihancccd > DateTime.Today) thoihancccd = DateTime.Today;
             }

             return new Customer
             {
                 Hoten = ToTitleCase(txtHoten.Text.Trim()),
                 Socccd = txtSocccd.Text,
                 GioiTinh = (cbGioitinh != null ? cbGioitinh.Text : ""),
                 Nhandang = cbNhandang.Text,
                 Ngaycap = ngaycap,
                 Ngaysinh = ngaysinh,
                 Noicap = cbNoicap.Text,
                 Xa = cbXa.Text,
                 Thon = cbThon.Text,
                 Hoi = cbHoi.Text,
                 Totruong = cbTo.Text,
                 To = "",
                 PGD = cbPGD.Text,
                 Chuongtrinh = cbChuongtrinh.Text,
                 Vtc = (cbVtc != null ? cbVtc.Text : ""),
                 Phuongan = ToTitleCase(txtPhuongan != null ? txtPhuongan.Text : ""),
                 Thoihanvay = cbThoihanvay.Text,
                 Phanky = (cbPhanky != null ? cbPhanky.Text : ""),
                 Sotien = cbSotien.Text,
                 Sotien1 = cbSotien1.Text,
                 Sotien2 = cbSotien2.Text,
                 Soluong1 = txtDoituong1.Text,
                 Soluong2 = txtDoituong2.Text,
                 Sotientong = "",
                 Sotienchu = "",
                 Mucdich1 = ToTitleCase(txtMucdich1 != null ? txtMucdich1.Text : ""),
                 Mucdich2 = ToTitleCase(txtMucdich2 != null ? txtMucdich2.Text : ""),
                 Doituong1 = (cbDoituong != null ? cbDoituong.Text : ""),
                 Doituong2 = "",
                 Ngaylaphs = ngaylaphs,
                 Ngaydenhan = ngaydenhan,
                 Thoihancccd = thoihancccd,
                 Dantoc = (cbDantoc != null ? cbDantoc.Text : ""),
                 Sdt = (txtSdt != null ? txtSdt.Text : ""),
                 Ntk1 = ToTitleCase(ntk1), Ntk2 = ToTitleCase(ntk2), Ntk3 = ToTitleCase(ntk3),
                 CccdNtk1 = cccdntk1, CccdNtk2 = cccdntk2, CccdNtk3 = cccdntk3,
                 Namsinh1 = namsinh1, Namsinh2 = namsinh2, Namsinh3 = namsinh3,
                 Qh1 = qh1, Qh2 = qh2, Qh3 = qh3
             };
         }

         private void PopulateForm(Customer c)
         {
             if (c == null) return;

             // Thông tin cơ bản
             txtHoten.Text = c.Hoten ?? "";
             txtSocccd.Text = c.Socccd ?? "";
             cbNhandang.Text = c.Nhandang ?? "";
             var ngaycap = c.Ngaycap == DateTime.MinValue ? DateTime.Today : c.Ngaycap;
             if (ngaycap > DateTime.Today) ngaycap = DateTime.Today;
             dateNgaycapCCCD.Value = ngaycap;

             try 
             { 
                 if (dateNgaysinh != null) 
                 {
                     var ngaysinh = c.Ngaysinh == DateTime.MinValue ? DateTime.Today : c.Ngaysinh;
                     if (ngaysinh > DateTime.Today) ngaysinh = DateTime.Today;
                     dateNgaysinh.Value = ngaysinh;
                 }
             } catch { }
             cbNoicap.Text = c.Noicap ?? "";

             // Thông tin cá nhân bổ sung
             try { if (cbGioitinh != null) cbGioitinh.Text = c.GioiTinh ?? ""; } catch { }
             try { if (cbDantoc != null) cbDantoc.Text = c.Dantoc ?? ""; } catch { }
             try { if (txtSdt != null) txtSdt.Text = c.Sdt ?? ""; } catch { }

             // Thông tin vị trí (đã bật suppress để tránh các sự kiện cascading)
             suppressComboChanged = true;
             try
             {
                 cbPGD.Text = c.PGD ?? ""; 
                 LoadXinManData();
                 if (cbXa.Items.Count == 0 && xinmanModel != null) 
                     foreach (var com in xinmanModel.communes) 
                         if (!string.IsNullOrWhiteSpace(com.name) && !cbXa.Items.Contains(com.name)) 
                             cbXa.Items.Add(com.name);

                 try { if (!string.IsNullOrEmpty(c.Xa)) { if (!cbXa.Items.Cast<object>().Any(x => string.Equals((x ?? "").ToString(), c.Xa, StringComparison.OrdinalIgnoreCase))) cbXa.Items.Add(c.Xa); cbXa.Text = c.Xa; } else cbXa.Text = ""; } catch { }
                 try { if (!string.IsNullOrEmpty(c.Thon)) { if (!cbThon.Items.Cast<object>().Any(x => string.Equals((x ?? "").ToString(), c.Thon, StringComparison.OrdinalIgnoreCase))) cbThon.Items.Add(c.Thon); cbThon.Text = c.Thon; } else cbThon.Text = ""; } catch { }
                 try { if (!string.IsNullOrEmpty(c.Hoi)) { if (!cbHoi.Items.Cast<object>().Any(x => string.Equals((x ?? "").ToString(), c.Hoi, StringComparison.OrdinalIgnoreCase))) cbHoi.Items.Add(c.Hoi); cbHoi.Text = c.Hoi; } else cbHoi.Text = ""; } catch { }
                 try { var toVal = !string.IsNullOrEmpty(c.Totruong) ? c.Totruong : (!string.IsNullOrEmpty(c.To) ? c.To : ""); if (!string.IsNullOrEmpty(toVal)) { if (!cbTo.Items.Cast<object>().Any(x => string.Equals((x ?? "").ToString(), toVal, StringComparison.OrdinalIgnoreCase))) cbTo.Items.Add(toVal); cbTo.Text = toVal; } else cbTo.Text = ""; } catch { }
             }
             finally { suppressComboChanged = false; }

             // Thông tin chương trình và khoản vay
             cbChuongtrinh.Text = c.Chuongtrinh ?? "";
             try { if (cbVtc != null) cbVtc.Text = c.Vtc ?? ""; } catch { }
             try { if (txtPhuongan != null) txtPhuongan.Text = c.Phuongan ?? ""; } catch { }
             cbThoihanvay.Text = c.Thoihanvay ?? "";
             try { if (cbPhanky != null) cbPhanky.Text = c.Phanky ?? ""; } catch { }

             // Thông tin số tiền
             cbSotien.Text = c.Sotien ?? "";
             cbSotien1.Text = c.Sotien1 ?? "";
             cbSotien2.Text = c.Sotien2 ?? "";

             // Mục đích và Đối tượng
             try { if (txtMucdich1 != null) txtMucdich1.Text = c.Mucdich1 ?? ""; } catch { }
             try { if (txtMucdich2 != null) txtMucdich2.Text = c.Mucdich2 ?? ""; } catch { }
             try { if (cbDoituong != null) cbDoituong.Text = c.Doituong1 ?? ""; } catch { }
             try { if (txtDoituong1 != null) txtDoituong1.Text = c.Soluong1 ?? ""; } catch { }
             try { if (txtDoituong2 != null) txtDoituong2.Text = c.Soluong2 ?? ""; } catch { }

             // Các ngày - đảm bảo không có ngày tương lai
             var ngaylaphs = c.Ngaylaphs == DateTime.MinValue ? DateTime.Today : c.Ngaylaphs;
             if (ngaylaphs > DateTime.Today) ngaylaphs = DateTime.Today;
             dateLaphs.Value = ngaylaphs;
             try 
             { 
                 if (dateDH != null) 
                 {
                     dateDH.ShowCheckBox = true;
                     if (c.Ngaydenhan == DateTime.MinValue)
                     {
                         dateDH.Checked = false;
                     }
                     else
                     {
                         var ngaydenhan = c.Ngaydenhan;
                         if (ngaydenhan > DateTime.Today) ngaydenhan = DateTime.Today;
                         dateDH.Format = DateTimePickerFormat.Custom;
                         dateDH.CustomFormat = "dd/MM/yyyy";
                         dateDH.Checked = true;
                         dateDH.Value = ngaydenhan;
                     }
                 }
             } catch { }
             try 
             { 
                 if (datendhcccd != null) 
                 {
                     datendhcccd.ShowCheckBox = true;
                     if (c.Thoihancccd == DateTime.MinValue)
                     {
                         datendhcccd.Checked = false;
                     }
                     else
                     {
                         var thoihancccd = c.Thoihancccd;
                         if (thoihancccd > DateTime.Today) thoihancccd = DateTime.Today;
                         datendhcccd.Format = DateTimePickerFormat.Custom;
                         datendhcccd.CustomFormat = "dd/MM/yyyy";
                         datendhcccd.Checked = true;
                         datendhcccd.Value = thoihancccd;
                     }
                 }
             } catch { }

             // NTK (Người thừa kế) info
             try { if (txtntk1 != null) txtntk1.Text = c.Ntk1 ?? ""; } catch { }
             try { if (txtntk2 != null) txtntk2.Text = c.Ntk2 ?? ""; } catch { }
             try { if (txtntk3 != null) txtntk3.Text = c.Ntk3 ?? ""; } catch { }

             try { if (txtcccd1 != null) txtcccd1.Text = c.CccdNtk1 ?? ""; } catch { }
             try { if (txtcccd2 != null) txtcccd2.Text = c.CccdNtk2 ?? ""; } catch { }
             try { if (txtcccd3 != null) txtcccd3.Text = c.CccdNtk3 ?? ""; } catch { }

             // ========== NGÀY SINH NGƯỜI THỪA KẾ (NTK) - DATEPICKER ==========
             // datentk1/2/3 là DateTimePicker: cần parse string Namsinh1/2/3 thành DateTime
             try 
             { 
                 if (datentk1 != null) 
                 {
                     datentk1.ShowCheckBox = true;  // Luôn hiển checkbox
                     if (string.IsNullOrWhiteSpace(c.Namsinh1))  // Nếu không có dữ liệu
                     {
                         datentk1.Checked = false;  // Bỏ tích checkbox = không có giá trị
                     }
                     else
                     {
                         // Parse string Namsinh1 (dd/MM/yyyy) thành DateTime
                         DateTime parsedDate = ParseDateTextOrFallback(c.Namsinh1);
                         if (parsedDate != DateTime.MinValue)  // Nếu parse thành công
                         {
                             datentk1.Checked = true;  // Tích checkbox
                             datentk1.Value = parsedDate;  // Set giá trị DateTime
                         }
                         else
                         {
                             datentk1.Checked = false;  // Parse lỗi thì bỏ tích
                         }
                     }
                 }
             } catch { }
             try 
             { 
                 if (datentk2 != null) 
                 {
                     datentk2.ShowCheckBox = true;
                     if (string.IsNullOrWhiteSpace(c.Namsinh2))
                     {
                         datentk2.Checked = false;
                     }
                     else
                     {
                         DateTime parsedDate = ParseDateTextOrFallback(c.Namsinh2);
                         if (parsedDate != DateTime.MinValue)
                         {
                             datentk2.Checked = true;
                             datentk2.Value = parsedDate;
                         }
                         else
                         {
                             datentk2.Checked = false;
                         }
                     }
                 }
             } catch { }
             try 
             { 
                 if (datentk3 != null) 
                 {
                     datentk3.ShowCheckBox = true;
                     if (string.IsNullOrWhiteSpace(c.Namsinh3))
                     {
                         datentk3.Checked = false;
                     }
                     else
                     {
                         DateTime parsedDate = ParseDateTextOrFallback(c.Namsinh3);
                         if (parsedDate != DateTime.MinValue)
                         {
                             datentk3.Checked = true;
                             datentk3.Value = parsedDate;
                         }
                         else
                         {
                             datentk3.Checked = false;
                         }
                     }
                 }
             } catch { }

             try { if (cbqh1 != null) cbqh1.Text = c.Qh1 ?? ""; } catch { }
             try { if (cbqh2 != null) cbqh2.Text = c.Qh2 ?? ""; } catch { }
             try { if (cbqh3 != null) cbqh3.Text = c.Qh3 ?? ""; } catch { }
         }

         private void ClearForm()
         {
             try { txtHoten.Clear(); } catch { } try { txtSocccd.Text = ""; } catch { } try { cbNhandang.Text = ""; } catch { }
             try { dateNgaycapCCCD.Value = DateTime.Today; } catch { }
             try { if (dateNgaysinh != null) dateNgaysinh.Value = DateTime.Today; } catch { }
             try { cbNoicap.Text = ""; } catch { } try { cbXa.Items.Clear(); cbThon.Items.Clear(); cbHoi.Items.Clear(); cbTo.Items.Clear(); } catch { }
             try { cbXa.Text = ""; cbThon.Text = ""; cbHoi.Text = ""; cbTo.Text = ""; } catch { }
             try { cbChuongtrinh.Text = ""; cbThoihanvay.Text = ""; cbSotien.Text = ""; cbSotien1.Text = ""; cbSotien2.Text = ""; } catch { }
             try { txtMucdich1.Clear(); txtMucdich2.Clear(); } catch { }
             try { dateLaphs.Value = DateTime.Today; cbPGD.Text = ""; editingIndex = -1; ResetVisibilityToDefault(); } catch { }

             // Xóa các trường ngày bằng cách bỏ tích chọn (các control vẫn hiển thị)
             try { if (dateDH != null) dateDH.Checked = false; } catch { }
             try { if (datendhcccd != null) datendhcccd.Checked = false; } catch { }
             try { if (datentk1 != null) datentk1.Checked = false; } catch { }
             try { if (datentk2 != null) datentk2.Checked = false; } catch { }
             try { if (datentk3 != null) datentk3.Checked = false; } catch { }

             // Xóa các trường NTK
             try { if (txtntk1 != null) txtntk1.Text = ""; } catch { }
             try { if (txtntk2 != null) txtntk2.Text = ""; } catch { }
             try { if (txtntk3 != null) txtntk3.Text = ""; } catch { }
             try { if (txtcccd1 != null) txtcccd1.Text = ""; } catch { }
             try { if (txtcccd2 != null) txtcccd2.Text = ""; } catch { }
             try { if (txtcccd3 != null) txtcccd3.Text = ""; } catch { }
             try { if (cbqh1 != null) cbqh1.Text = ""; } catch { }
             try { if (cbqh2 != null) cbqh2.Text = ""; } catch { }
             try { if (cbqh3 != null) cbqh3.Text = ""; } catch { }
         }

         private void LoadXinManData()
         {
             xinmanModel = null;

             var candidate = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "xinman.json");
             if (File.Exists(candidate))
                 try { var json = File.ReadAllText(candidate, Encoding.UTF8); xinmanModel = TryDeserializeXinman(json); }
                 catch { xinmanModel = null; }

             if (xinmanModel == null)
             {
                 try
                 {
                     var asm = Assembly.GetExecutingAssembly();
                     var resName = asm.GetManifestResourceNames().FirstOrDefault(n => n.EndsWith("xinman.json", StringComparison.OrdinalIgnoreCase));
                     if (resName != null) using (var s = asm.GetManifestResourceStream(resName)) using (var sr = new StreamReader(s, Encoding.UTF8)) { var json = sr.ReadToEnd(); xinmanModel = TryDeserializeXinman(json); }
                 }
                 catch { xinmanModel = null; }
             }
         }

         private XinManModel TryDeserializeXinman(string json)
         {
             if (string.IsNullOrWhiteSpace(json)) return null;
             try { var model = JsonConvert.DeserializeObject<XinManModel>(json); if (model != null && !string.IsNullOrWhiteSpace(model.pgd)) return model; } catch { }
             try
             {
                 var generic = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                 if (generic == null) return null;
                 var model = new XinManModel();
                 if (generic.TryGetValue("pgd", out var pgdObj)) model.pgd = pgdObj?.ToString();
                 if (generic.TryGetValue("communes", out var communesObj) && communesObj != null)
                 {
                     var communesJson = JsonConvert.SerializeObject(communesObj);
                     model.communes = JsonConvert.DeserializeObject<List<Commune>>(communesJson) ?? new List<Commune>();
                 }
                 if (!string.IsNullOrWhiteSpace(model.pgd)) return model;
             }
             catch { }
             return null;
         }

        private void CbPGD_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (suppressComboChanged) return; suppressComboChanged = true;
            try
            {
                var selected = (cbPGD.SelectedItem ?? cbPGD.Text)?.ToString() ?? "";
                if (string.IsNullOrWhiteSpace(selected)) { ClearPgdDependentCombos(); return; }
                LoadXinManData(); if (xinmanModel == null) { ClearPgdDependentCombos(); return; }
                var modelPgd = xinmanModel.pgd ?? "";
                bool matches = string.Equals(selected, modelPgd, StringComparison.OrdinalIgnoreCase)
                                || selected.IndexOf(modelPgd, StringComparison.OrdinalIgnoreCase) >= 0
                                || modelPgd.IndexOf(selected, StringComparison.OrdinalIgnoreCase) >= 0;
                if (!matches) { ClearPgdDependentCombos(); return; }

                cbXa.Items.Clear(); cbThon.Items.Clear(); cbHoi.Items.Clear(); cbTo.Items.Clear();
                foreach (var com in xinmanModel.communes) if (!string.IsNullOrWhiteSpace(com.name) && !cbXa.Items.Contains(com.name)) cbXa.Items.Add(com.name);
                var associations = xinmanModel.communes.Where(c => c.associations != null).SelectMany(c => c.associations).Where(a => !string.IsNullOrWhiteSpace(a.name)).Select(a => a.name).Distinct(StringComparer.OrdinalIgnoreCase);
                foreach (var a in associations) cbHoi.Items.Add(a);
                var communeLevelVillages = xinmanModel.communes.Where(c => c.villages != null).SelectMany(c => c.villages).Where(v => !string.IsNullOrWhiteSpace(v.name)).Select(v => v.name).Distinct(StringComparer.OrdinalIgnoreCase);
                foreach (var v in communeLevelVillages) if (!cbThon.Items.Contains(v)) cbThon.Items.Add(v);
                ResetVisibilityToDefault();
            }
            finally { suppressComboChanged = false; }
        }

        private void CbXa_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (suppressComboChanged) return; suppressComboChanged = true;
            try { cbThon.Items.Clear(); cbTo.Items.Clear(); if (cbXa.SelectedIndex < 0 || xinmanModel == null) { ResetVisibilityToDefault(); return; } var xaName = cbXa.SelectedItem.ToString(); var commune = xinmanModel.communes.FirstOrDefault(c => string.Equals(c.name, xaName, StringComparison.OrdinalIgnoreCase)); if (commune != null) { if (commune.associations != null) foreach (var assoc in commune.associations) if (assoc.villages != null) foreach (var v in assoc.villages) if (!string.IsNullOrWhiteSpace(v.name) && !cbThon.Items.Contains(v.name)) cbThon.Items.Add(v.name); if (commune.villages != null) foreach (var v in commune.villages) if (!string.IsNullOrWhiteSpace(v.name) && !cbThon.Items.Contains(v.name)) cbThon.Items.Add(v.name); cbHoi.Items.Clear(); if (commune.associations != null) foreach (var a in commune.associations) if (!string.IsNullOrWhiteSpace(a.name) && !cbHoi.Items.Contains(a.name)) cbHoi.Items.Add(a.name); } ResetVisibilityToDefault(); }
            finally { suppressComboChanged = false; }
        }

        private void CbThon_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (suppressComboChanged) return; suppressComboChanged = true;
            try { cbTo.Items.Clear(); if (cbThon.SelectedIndex < 0 || xinmanModel == null) { ResetVisibilityToDefault(); return; } var thonName = cbThon.SelectedItem.ToString(); Association foundAssoc = null; Commune foundCommune = null; foreach (var com in xinmanModel.communes) { if (com.associations != null) foreach (var assoc in com.associations) if (assoc.villages != null && assoc.villages.Any(v => string.Equals(v.name, thonName, StringComparison.OrdinalIgnoreCase))) { foundAssoc = assoc; foundCommune = com; break; } if (foundAssoc != null) break; if (com.villages != null && com.villages.Any(v => string.Equals(v.name, thonName, StringComparison.OrdinalIgnoreCase))) { foundCommune = com; break; } } if (foundCommune != null) cbXa.SelectedItem = foundCommune.name; if (foundAssoc != null) { cbHoi.SelectedItem = foundAssoc.name; var village = foundAssoc.villages.FirstOrDefault(v => string.Equals(v.name, thonName, StringComparison.OrdinalIgnoreCase)); if (village != null && village.groups != null) foreach (var g in village.groups) if (!string.IsNullOrWhiteSpace(g)) cbTo.Items.Add(g); } else if (foundCommune != null) { var village = foundCommune.villages.FirstOrDefault(v => string.Equals(v.name, thonName, StringComparison.OrdinalIgnoreCase)); if (village != null && village.groups != null) foreach (var g in village.groups) if (!string.IsNullOrWhiteSpace(g)) cbTo.Items.Add(g); } ResetVisibilityToDefault(); }
            finally { suppressComboChanged = false; }
        }

        private void CbHoi_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (suppressComboChanged) return; suppressComboChanged = true;
            try { cbThon.Items.Clear(); cbTo.Items.Clear(); if (cbHoi.SelectedIndex < 0 || xinmanModel == null) { ResetVisibilityToDefault(); return; } var hoiName = cbHoi.SelectedItem.ToString(); var assocList = xinmanModel.communes.Where(c => c.associations != null).SelectMany(c => c.associations.Select(a => new { Commune = c, Assoc = a })).Where(x => string.Equals(x.Assoc.name, hoiName, StringComparison.OrdinalIgnoreCase)).ToList(); var villageNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase); foreach (var pair in assocList) { var assoc = pair.Assoc; if (assoc.villages != null) foreach (var v in assoc.villages) villageNames.Add(v.name); if (assoc.managedVillages != null) foreach (var name in assoc.managedVillages) villageNames.Add(name); } foreach (var vn in villageNames) cbThon.Items.Add(vn); var communesForAssoc = assocList.Select(x => x.Commune.name).Distinct(StringComparer.OrdinalIgnoreCase).ToList(); if (communesForAssoc.Count == 1) cbXa.SelectedItem = communesForAssoc[0]; else cbXa.SelectedIndex = -1; var groupSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase); foreach (var pair in assocList) foreach (var v in pair.Assoc.villages ?? Enumerable.Empty<Village>()) foreach (var g in v.groups ?? Enumerable.Empty<string>()) groupSet.Add(g); foreach (var g in groupSet) cbTo.Items.Add(g); ResetVisibilityToDefault(); }
            finally { suppressComboChanged = false; }
        }

        private void ClearPgdDependentCombos()
        {
            cbXa.Items.Clear(); cbThon.Items.Clear(); cbHoi.Items.Clear(); cbTo.Items.Clear();
            cbXa.SelectedIndex = -1; cbThon.SelectedIndex = -1; cbHoi.SelectedIndex = -1; cbTo.SelectedIndex = -1; ResetVisibilityToDefault();
        }

        private void ResetVisibilityToDefault()
        {
            cbXa.Visible = true; cbThon.Visible = true; cbHoi.Visible = true; cbTo.Visible = true;
        }

        private void ExportSpecificTemplate(Customer c, string templateFileName)
        {
            var destFolder = GetProfileFolderPath(c); Directory.CreateDirectory(destFolder);
            string templatePath; try { templatePath = ResolveTemplatePath(templateFileName); } catch (FileNotFoundException ex) { MessageBox.Show($"Template {templateFileName} not found: {ex.Message}"); return; }
            if (!IsDocxFile(templatePath)) { MessageBox.Show($"{templateFileName} source is not a valid .docx."); return; }
            var shortName = Path.GetFileNameWithoutExtension(templateFileName).Replace(" ", "_"); var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss"); var destDoc = Path.Combine(destFolder, MakeFileSystemSafe(c.Hoten) + "_" + shortName + "_" + ts + ".docx");
            try { File.Copy(templatePath, destDoc, false); } catch (IOException ioex) { MessageBox.Show($"Failed to create destination file: {ioex.Message}"); return; }
            if (!IsDocxFile(destDoc)) { try { if (File.Exists(destDoc)) File.Delete(destDoc); } catch { } MessageBox.Show("Produced file is not a valid .docx. Aborting."); return; }

            ReplacePlaceholdersInWord(destDoc, c);
        }

        private List<Customer> GetSelectedCustomers() { var list = new List<Customer>(); try { foreach (DataGridViewRow row in dgv.SelectedRows) try { var item = row.DataBoundItem as Customer; if (item != null) list.Add(item); } catch { } } catch { } return list; }

        private void Btn03Group_Click(object sender, EventArgs e) { try { var selected = GetSelectedCustomers(); var f2 = new Form2(selected); f2.ShowDialog(); } catch (Exception ex) { MessageBox.Show("Lỗi khi mở Form nhóm: " + ex.Message); } }

        // Tạo mới khách hàng - tạo khách hàng mới từ thông tin trong form
        private void BtnTaokh_Click(object sender, EventArgs e)
        {
            try
            {
                // Đọc thông tin từ form
                var customer = ReadForm();

                // Validate tên khách hàng
                if (string.IsNullOrWhiteSpace(customer.Hoten))
                {
                    MessageBox.Show("Vui lòng nhập Họ và tên để tạo khách hàng mới.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Tạo khách hàng mới (không dùng _fileName cũ, để hệ thống tự tạo file mới)
                customer._fileName = null; // Force tạo file mới

                try
                {
                    // Lưu khách hàng mới
                    SaveCustomerToFile(customer);

                    // Thêm vào danh sách (không update editingIndex)
                    if (customers != null)
                    {
                        customers.Add(customer);
                    }

                    // Refresh grid
                    BindGrid();

                    MessageBox.Show($"Đã tạo khách hàng mới '{customer.Hoten}' thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Reset editing state
                    editingIndex = -1;

                    // Clear all form fields để tiếp tục tạo khách mới nếu muốn
                    ClearForm();

                    // Deselect any selected rows in grid
                    try
                    {
                        if (dgv != null && dgv.SelectedRows.Count > 0)
                        {
                            dgv.ClearSelection();
                        }
                    }
                    catch { }

                    // Set focus to name field to start entering new customer
                    try { txtHoten.Focus(); } catch { }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi tạo khách hàng mới: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tạo mới: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Đăng xuất khỏi phần chỉnh sửa xinman.json
        private void BtnExit_Click(object sender, EventArgs e)
        {
            try
            {
                // Clear login fields if they exist
                try { if (txtUsername != null) txtUsername.Text = ""; } catch { }
                try { if (txtPassword != null) txtPassword.Text = ""; } catch { }

                // Call XinManEditor to logout and disable editing
                try
                {
                    if (xinManEditor != null)
                    {
                        xinManEditor.Logout();
                    }
                }
                catch { }

                MessageBox.Show("Đã đăng xuất thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi đăng xuất: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Tự động viết hoa chữ cái đầu (Title Case) cho các textbox tên (viết hoa chữ cái đầu của mỗi từ)
        private void TxtName_TextChanged(object sender, EventArgs e)
        {
            if (suppressNameChanged) return;

            try
            {
                var tb = sender as TextBox;
                if (tb == null) return;

                var originalText = tb.Text ?? string.Empty;
                if (string.IsNullOrEmpty(originalText)) return;

                var originalSelection = tb.SelectionStart;
                var capitalizedText = CapitalizeWords(originalText);

                if (!string.Equals(capitalizedText, originalText, StringComparison.Ordinal))
                {
                    suppressNameChanged = true;
                    tb.Text = capitalizedText;
                    // Khôi phục vị trí con trỏ
                    tb.SelectionStart = Math.Min(originalSelection, tb.Text.Length);
                    suppressNameChanged = false;
                }
            }
            catch { }
        }

        // Áp dụng Title Case khi rời khỏi textbox tên (dọc dẹp cuối cùng)
        private void TxtName_Leave(object sender, EventArgs e)
        {
            try
            {
                var tb = sender as TextBox;
                if (tb == null) return;

                var originalText = tb.Text ?? string.Empty;
                if (string.IsNullOrWhiteSpace(originalText)) return;

                var titleCased = ToTitleCase(originalText);
                if (!string.Equals(titleCased, originalText, StringComparison.Ordinal))
                {
                    suppressNameChanged = true;
                    tb.Text = titleCased;
                    tb.SelectionStart = tb.Text.Length;
                    suppressNameChanged = false;
                }
            }
            catch { }
        }

        // Viết hoa chữ cái đầu của mỗi từ trong khi gõ
        private string CapitalizeWords(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;

            try
            {
                var result = new StringBuilder(input.Length);
                bool capitalizeNext = true;

                foreach (char c in input)
                {
                    if (char.IsWhiteSpace(c))
                    {
                        result.Append(c);
                        capitalizeNext = true;
                    }
                    else if (capitalizeNext && char.IsLetter(c))
                    {
                        result.Append(char.ToUpper(c));
                        capitalizeNext = false;
                    }
                    else
                    {
                        result.Append(char.ToLower(c));
                    }
                }

                return result.ToString();
            }
            catch
            {
                return input;
            }
        }

        private void TxtNamsinh_KeyPress(object sender, KeyPressEventArgs e) { if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '/' && e.KeyChar != '-') e.Handled = true; }
        private void TxtNamsinh_TextChanged(object sender, EventArgs e) { if (suppressNamsinhChanged) return; try { var tb = sender as TextBox; if (tb == null) return; var originalText = tb.Text ?? string.Empty; var txt = originalText.Trim(); if (string.IsNullOrEmpty(txt)) return; var origSel = tb.SelectionStart; string digits = new string(txt.Where(char.IsDigit).ToArray()); DateTime dt; if (digits.Length == 6) { var d = int.Parse(digits.Substring(0, 2)); var m = int.Parse(digits.Substring(2, 2)); var yy = int.Parse(digits.Substring(4, 2)); int year = (yy >= 50) ? 1900 + yy : 2000 + yy; if (d >= 1 && d <= 31 && m >= 1 && m <= 12) { dt = new DateTime(year, m, d); if (dt.Date > DateTime.Today) { MessageBox.Show("Ngày sinh không được lớn hơn ngày hiện tại.", "Lỗi nhập liệu", MessageBoxButtons.OK, MessageBoxIcon.Warning); suppressNamsinhChanged = true; tb.Text = ""; suppressNamsinhChanged = false; return; } var formatted = dt.ToString("dd/MM/yyyy"); if (!string.Equals(formatted, originalText, StringComparison.Ordinal)) { suppressNamsinhChanged = true; tb.Text = formatted; tb.SelectionStart = formatted.Length; suppressNamsinhChanged = false; } return; } } if (digits.Length == 8) { var d = int.Parse(digits.Substring(0, 2)); var m = int.Parse(digits.Substring(2, 2)); var yyyy = int.Parse(digits.Substring(4, 4)); if (d >= 1 && d <= 31 && m >= 1 && m <= 12) { dt = new DateTime(yyyy, m, d); if (dt.Date > DateTime.Today) { MessageBox.Show("Ngày sinh không được lớn hơn ngày hiện tại.", "Lỗi nhập liệu", MessageBoxButtons.OK, MessageBoxIcon.Warning); suppressNamsinhChanged = true; tb.Text = ""; suppressNamsinhChanged = false; return; } var formatted = dt.ToString("dd/MM/yyyy"); if (!string.Equals(formatted, originalText, StringComparison.Ordinal)) { suppressNamsinhChanged = true; tb.Text = formatted; tb.SelectionStart = formatted.Length; suppressNamsinhChanged = false; } return; } } var formats = new[] { "d/M/yyyy", "dd/MM/yyyy", "d-M-yyyy", "dd-MM-yyyy" }; if (DateTime.TryParseExact(txt, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt)) { if (dt.Date > DateTime.Today) { MessageBox.Show("Ngày sinh không được lớn hơn ngày hiện tại.", "Lỗi nhập liệu", MessageBoxButtons.OK, MessageBoxIcon.Warning); suppressNamsinhChanged = true; tb.Text = ""; suppressNamsinhChanged = false; return; } var formatted = dt.ToString("dd/MM/yyyy"); if (!string.Equals(formatted, originalText, StringComparison.Ordinal)) { suppressNamsinhChanged = true; tb.Text = formatted; int delta = formatted.Length - originalText.Length; int newSel = origSel + delta; if (newSel < 0) newSel = 0; if (newSel > formatted.Length) newSel = formatted.Length; tb.SelectionStart = newSel; suppressNamsinhChanged = false; } return; } } catch { } }

        private DateTime ParseDateTextOrFallback(string text) { if (string.IsNullOrWhiteSpace(text)) return DateTime.MinValue; try { var digits = new string((text ?? "").Where(char.IsDigit).ToArray()); if (digits.Length == 6) { var d = int.Parse(digits.Substring(0, 2)); var m = int.Parse(digits.Substring(2, 2)); var yy = int.Parse(digits.Substring(4, 2)); int year = (yy >= 50) ? 1900 + yy : 2000 + yy; return new DateTime(year, m, d); } if (digits.Length == 8) { var d = int.Parse(digits.Substring(0, 2)); var m = int.Parse(digits.Substring(2, 2)); var yyyy = int.Parse(digits.Substring(4, 4)); return new DateTime(yyyy, m, d); } if (digits.Length == 4) return DateTime.MinValue; DateTime dt; var formats = new[] { "d/M/yyyy", "dd/MM/yyyy", "yyyy", "yyyyMMdd", "d-M-yyyy", "dd-MM-yyyy" }; if (DateTime.TryParseExact(text, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt)) return dt; if (DateTime.TryParse(text, out dt)) return dt; } catch { } return DateTime.MinValue; }

        private string ToTitleCase(string input) { if (string.IsNullOrWhiteSpace(input)) return input; try { var culture = new CultureInfo("vi-VN"); var cleaned = Regex.Replace(input.Trim(), "\\s+", " ").ToLower(culture); return culture.TextInfo.ToTitleCase(cleaned); } catch { return input; } }

        private string FormatNamsinhStringForDoc(string namsinh, string docPath) { if (string.IsNullOrWhiteSpace(namsinh)) return ""; try { DateTime dt; var formats = new[] { "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "d-M-yyyy", "yyyyMMdd", "yyyy" }; if (DateTime.TryParseExact(namsinh.Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt)) return ShouldShowFullNamsinh(docPath) ? dt.ToString("dd/MM/yyyy") : dt.Year.ToString(); if (namsinh.Length == 8) { var d = int.Parse(namsinh.Substring(0, 2)); var m = int.Parse(namsinh.Substring(2, 2)); var yyyy = int.Parse(namsinh.Substring(4, 4)); dt = new DateTime(yyyy, m, d); return ShouldShowFullNamsinh(docPath) ? dt.ToString("dd/MM/yyyy") : dt.Year.ToString(); } if (namsinh.Length == 6) { var d = int.Parse(namsinh.Substring(0, 2)); var m = int.Parse(namsinh.Substring(2, 2)); var yy = int.Parse(namsinh.Substring(4, 2)); int year = (yy >= 50) ? 1900 + yy : 2000 + yy; dt = new DateTime(year, m, d); return ShouldShowFullNamsinh(docPath) ? dt.ToString("dd/MM/yyyy") : dt.Year.ToString(); } var ds = new string(namsinh.Where(char.IsDigit).ToArray()); if (ds.Length >= 4) { var ypart = ds.Substring(ds.Length - 4); if (int.TryParse(ypart, out int yv)) return yv.ToString(); } } catch { } return namsinh; }

        private string ExtractYearString(string namsinh) { if (string.IsNullOrWhiteSpace(namsinh)) return null; try { var digits = new string(namsinh.Where(char.IsDigit).ToArray()); if (digits.Length >= 4) { var yearPart = digits.Substring(digits.Length - 4); if (int.TryParse(yearPart, out int year)) return year.ToString(); } } catch { } return null; }

        private bool ShouldShowFullNamsinh(string docPath) { if (string.IsNullOrWhiteSpace(docPath)) return false; var name = Path.GetFileName(docPath) ?? ""; return name.IndexOf("GUQ", StringComparison.OrdinalIgnoreCase) >= 0; }
        private bool Is03DS(string docPath) { if (string.IsNullOrWhiteSpace(docPath)) return false; var name = Path.GetFileName(docPath) ?? ""; return (name.IndexOf("03", StringComparison.OrdinalIgnoreCase) >= 0 && name.IndexOf("DS", StringComparison.OrdinalIgnoreCase) >= 0) || name.IndexOf("03 DS", StringComparison.OrdinalIgnoreCase) >= 0; }
        private bool Is01SXKD(string docPath) { if (string.IsNullOrWhiteSpace(docPath)) return false; var name = Path.GetFileName(docPath) ?? ""; return name.IndexOf("01", StringComparison.OrdinalIgnoreCase) >= 0 && name.IndexOf("SXKD", StringComparison.OrdinalIgnoreCase) >= 0; }

        // Giải quyết đường dẫn mẫu bằng cách kiểm tra nhiều vị trí và embedded resources; cache kết quả
        private string ResolveTemplatePath(string templateFileName)
        {
            if (string.IsNullOrWhiteSpace(templateFileName)) throw new FileNotFoundException("Template name is empty.");
            lock (templatePathCache)
            {
                if (templatePathCache.TryGetValue(templateFileName, out var cached) && !string.IsNullOrWhiteSpace(cached) && File.Exists(cached))
                    return cached;
            }

            // 1) Kiểm tra thư mục Templates bên cạnh exe
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var candidate = Path.Combine(baseDir, TemplatesFolder, templateFileName);
            if (File.Exists(candidate))
            {
                lock (templatePathCache) { templatePathCache[templateFileName] = candidate; }
                return candidate;
            }

            // 2) Kiểm tra thư mục gốc baseDir
            candidate = Path.Combine(baseDir, templateFileName);
            if (File.Exists(candidate)) { lock (templatePathCache) { templatePathCache[templateFileName] = candidate; } return candidate; }

            // 3) Tìm kiếm đệ quy dưới baseDir
            try
            {
                var found = Directory.EnumerateFiles(baseDir, templateFileName, SearchOption.AllDirectories).FirstOrDefault();
                if (!string.IsNullOrEmpty(found)) { lock (templatePathCache) { templatePathCache[templateFileName] = found; } return found; }
            }
            catch { }

            // 4) Trích xuất embedded resource
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                var resName = asm.GetManifestResourceNames().FirstOrDefault(n => n.EndsWith(templateFileName, StringComparison.OrdinalIgnoreCase) || n.IndexOf(templateFileName, StringComparison.OrdinalIgnoreCase) >= 0);
                if (!string.IsNullOrEmpty(resName))
                {
                    var temp = Path.Combine(Path.GetTempPath(), "template_" + Guid.NewGuid().ToString("N") + Path.GetExtension(templateFileName));
                    using (var s = asm.GetManifestResourceStream(resName))
                    {
                        if (s != null)
                        {
                            using (var fs = File.OpenWrite(temp)) s.CopyTo(fs);
                        }
                    }
                    if (File.Exists(temp))
                    {
                        lock (templatePathCache) { templatePathCache[templateFileName] = temp; }
                        return temp;
                    }
                }
            }
            catch { }

            throw new FileNotFoundException("Template not found: " + templateFileName);
        }

        // Xác thực rất nhỏ: kiểm tra file tồn tại và là .docx; thử mở bằng OpenXML nếu có thể
        private bool IsDocxFile(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return false;
                if (!string.Equals(Path.GetExtension(path), ".docx", StringComparison.OrdinalIgnoreCase)) return false;
                // Thử mở như WordprocessingDocument để đảm bảo nó là package hợp lệ
                try
                {
                    using (var w = WordprocessingDocument.Open(path, false)) { /* thành công */ }
                    return true;
                }
                catch { return false; }
            }
            catch { return false; }
        }

        // --- Đã thêm các handler và helper thiếu ---

        // Các handler rỗng được tham chiếu bởi Designer
        private void tabPage1_Click(object sender, EventArgs e) { }
        private void groupBox2_Enter(object sender, EventArgs e) { }
        private void datentk2_TextChanged(object sender, EventArgs e) { TxtNamsinh_TextChanged(sender, e); }
        private void textBox1_TextChanged(object sender, EventArgs e) { }
        private void label14_Click(object sender, EventArgs e) { }
        private void cbXa_SelectedIndexChanged_1(object sender, EventArgs e) { CbXa_SelectedIndexChanged(sender, e); }
        private void label2_Click(object sender, EventArgs e) { }

        // ================= MARQUEE TEXT - KHỚI TẠO =================
        /// <summary>
        /// Khởi tạo hiệu ứng chữ chạy (marquee) cho:
        /// 1. Thanh tiêu đề Form (title bar) - Text: "PHẦN MỀM TẠO HỒ SƠ VAY VỐN"
        /// 2. Label14 bên trong Form - Text: "BẢNG NHẬP THÔNG TIN KHÁCH HÀNG VÀ NGƯỜI THỪA KẾ"
        /// Cả 2 text sẽ chạy từ phải sang trái với tốc độ khác nhau
        /// </summary>
        private void InitializeMarquee()
        {
            try
            {
                // ========== 1. MARQUEE CHO TITLE BAR (THANH TIÊU ĐỀ) ==========
                // Text hiển thị ở thanh tiêu đề Form (title bar)
                string titleText = "PHẦN MỀM TẠO HỒ SƠ VAY VỐN";

                // Thêm khoảng trắng để tạo khoảng cách giữa các lần lặp
                marqueeText = "     " + titleText + "     ";
                marqueePosition = 0;

                // ========== 2. MARQUEE CHO LABEL14 (BÊN TRONG FORM) ==========
                if (label14 != null)
                {
                    // Lấy text từ label14 (Designer)
                    string label14Text = label14.Text;

                    // Thêm khoảng trắng
                    label14MarqueeText = "     " + label14Text + "     ";
                    label14MarqueePosition = 0;
                }

                // ========== 3. KHỚI ĐỘNG TIMER ==========
                // Cấu hình Timer (đã tạo sẵn trong Designer)
                if (marqueeTimer != null)
                {
                    marqueeTimer.Interval = 100;  // 100ms = 10 FPS (mượt, không quá nhanh)
                    marqueeTimer.Start();
                }
            }
            catch { }
        }

        /// <summary>
        /// Xử lý sự kiện Tick của Timer - Di chuyển text theo từng ký tự
        /// Cập nhật ĐỒNG THỜI cả 2 marquee:
        /// 1. Form.Text (title bar)
        /// 2. label14.Text (bên trong Form)
        /// </summary>
        private void MarqueeTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                // ========== 1. CẬP NHẬT TITLE BAR MARQUEE ==========
                if (!string.IsNullOrEmpty(marqueeText))
                {
                    // Di chuyển vị trí 1 ký tự
                    marqueePosition++;
                    if (marqueePosition >= marqueeText.Length)
                    {
                        marqueePosition = 0;  // Quay lại đầu khi hết chuỗi
                    }

                    // Tạo text hiển thị bằng cách xoay vòng chuỗi
                    string displayText = marqueeText.Substring(marqueePosition) + 
                                         marqueeText.Substring(0, marqueePosition);

                    // Cập nhật text ở THANH TIÊU ĐỀ FORM (title bar)
                    this.Text = displayText;
                }

                // ========== 2. CẬP NHẬT LABEL14 MARQUEE ==========
                if (label14 != null && !string.IsNullOrEmpty(label14MarqueeText))
                {
                    // Di chuyển vị trí 1 ký tự (chạy nhanh hơn title bar một chút)
                    label14MarqueePosition++;
                    if (label14MarqueePosition >= label14MarqueeText.Length)
                    {
                        label14MarqueePosition = 0;  // Quay lại đầu khi hết chuỗi
                    }

                    // Tạo text hiển thị bằng cách xoay vòng chuỗi
                    string displayText = label14MarqueeText.Substring(label14MarqueePosition) + 
                                         label14MarqueeText.Substring(0, label14MarqueePosition);

                    // Cập nhật text cho label14
                    label14.Text = displayText;
                }
            }
            catch { }
        }

        // ================= TẠO STYLE THEME MACBOOK =================
        private void ApplyMacBookTheme()
        {
            try
            {
                // Nền form
                this.BackColor = AppTheme.MacBackground;

                // Tất cả các tab page
                if (tabPage1 != null)
                {
                    tabPage1.BackColor = AppTheme.MacBackground;
                }
                if (tabPage2 != null)
                {
                    tabPage2.BackColor = AppTheme.MacBackground;
                }
                if (tabPage3 != null)
                {
                    tabPage3.BackColor = AppTheme.MacBackground;
                }

                // Tạo style cho TabControl
                if (tabControl1 != null)
                {
                    tabControl1.Font = new System.Drawing.Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Regular);
                }

                // Các GroupBox - kiểu thẻ (off-white thay vì trắng tinh khiết)
                if (groupBox1 != null)
                {
                    groupBox1.BackColor = AppTheme.MacCardBackground;
                    groupBox1.ForeColor = AppTheme.MacTextPrimary;
                    groupBox1.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
                }

                if (groupBox2 != null)
                {
                    groupBox2.BackColor = AppTheme.MacCardBackground;
                    groupBox2.ForeColor = AppTheme.MacTextPrimary;
                    groupBox2.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
                }

                if (groupBox3 != null)
                {
                    groupBox3.BackColor = AppTheme.MacCardBackground;
                    groupBox3.ForeColor = AppTheme.MacTextPrimary;
                    groupBox3.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
                }

                if (groupBox4 != null)
                {
                    groupBox4.BackColor = AppTheme.MacCardBackground;
                    groupBox4.ForeColor = AppTheme.MacTextPrimary;
                    groupBox4.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
                }

                // Header label14 - Màu sắc hiện đại và dễ nhìn
                if (label14 != null)
                {
                    // Font lớn hơn, hiện đại hơn
                    label14.Font = new System.Drawing.Font("Segoe UI", 13F, System.Drawing.FontStyle.Bold);

                    // Màu cyan vibrant - sáng, dễ nhìn trên nền xanh
                    label14.ForeColor = AppTheme.MarqueeCyan;  // RGB(0, 150, 255)

                    // Nếu muốn màu trắng (dễ nhìn nhất), dùng dòng này:
                    // label14.ForeColor = System.Drawing.Color.White;

                    // BackColor trong suốt để nhìn thấy nền form
                    label14.BackColor = System.Drawing.Color.Transparent;
                }

                // Tạo style các nút với màu Mac
                StyleMacButton(btn01, AppTheme.MacGreen);      // Lưu - Green
                StyleMacButton(btn03, AppTheme.MacBlue);       // Export 03 - Blue
                StyleMacButton(btnGUQ, AppTheme.MacBlue);      // Export GUQ - Blue
                StyleMacButton(btnDelete, AppTheme.MacRed);    // Xóa - Red
                StyleMacButton(btntaokh, AppTheme.MacOrange);  // Tạo mới - Orange
                StyleMacButton(btn03Group, AppTheme.MacBlue);  // Nhóm - Blue

                // Các nút Tab2 (nếu tồn tại)
                StyleMacButton(btnPre, AppTheme.MacBlue);      // Previous
                StyleMacButton(btnNext, AppTheme.MacBlue);     // Next

                // Các nút Tab3
                StyleMacButton(btnLogin, AppTheme.MacTeal);    // Login - Teal
                StyleMacButton(btnSave, AppTheme.MacGreen);    // Save - Green
                StyleMacButton(btnexit, AppTheme.MacRed);      // Exit - Red

                // Tạo style cho tất cả các DataGridView
                StyleMacDataGridView();
                StyleAllDataGridViews();

                // Áp dụng font cho tất cả các label
                ApplyMacFontsToLabels();

                // Tạo style cho textbox và combobox
                ApplyMacStyleToTextBoxes();
                ApplyMacStyleToComboBoxes();

                // Tạo style cho RichTextBox
                ApplyMacStyleToRichTextBoxes();
            }
            catch { }
        }

        private void StyleMacButton(Button btn, System.Drawing.Color color)
        {
            if (btn == null) return;

            try
            {
                btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderSize = 0;
                btn.BackColor = color;
                btn.ForeColor = System.Drawing.Color.White;
                btn.Font = new System.Drawing.Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Regular);
                btn.Cursor = Cursors.Hand;
                btn.Height = 36;

                System.Drawing.Color originalColor = color;
                System.Drawing.Color hoverColor = color.ToArgb() == AppTheme.MacGreen.ToArgb() ? AppTheme.MacGreenHover :
                                  color.ToArgb() == AppTheme.MacRed.ToArgb() ? AppTheme.MacRedHover :
                                  color.ToArgb() == AppTheme.MacOrange.ToArgb() ? AppTheme.MacOrangeHover :
                                  AppTheme.MacBlueHover;

                btn.MouseEnter += (s, e) => { btn.BackColor = hoverColor; };
                btn.MouseLeave += (s, e) => { btn.BackColor = originalColor; };
            }
            catch { }
        }

        private void StyleMacDataGridView()
        {
            try
            {
                if (dgv == null) return;

                dgv.BorderStyle = BorderStyle.None;
                dgv.BackgroundColor = AppTheme.MacCardBackground;
                dgv.GridColor = AppTheme.MacBorderLight;
                dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;

                dgv.DefaultCellStyle.BackColor = System.Drawing.Color.White;
                dgv.DefaultCellStyle.ForeColor = AppTheme.MacTextPrimary;
                dgv.DefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 8.5F);
                dgv.DefaultCellStyle.SelectionBackColor = AppTheme.MacBlue;
                dgv.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.White;
                dgv.DefaultCellStyle.Padding = new Padding(6, 3, 6, 3);

                dgv.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(249, 249, 251);

                dgv.EnableHeadersVisualStyles = false;
                dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
                dgv.ColumnHeadersDefaultCellStyle.BackColor = AppTheme.MacHeaderGradient1;
                dgv.ColumnHeadersDefaultCellStyle.ForeColor = AppTheme.MacTextPrimary;
                dgv.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
                dgv.ColumnHeadersDefaultCellStyle.Padding = new Padding(8, 6, 8, 6);
                dgv.ColumnHeadersHeight = 36;

                dgv.RowTemplate.Height = 32;
            }
            catch { }
        }

        private void StyleAllDataGridViews()
        {
            try
            {
                foreach (System.Windows.Forms.Control ctrl in GetAllControlsForTheme(this))
                {
                    if (ctrl is DataGridView grid && grid != dgv)
                    {
                        grid.BorderStyle = BorderStyle.None;
                        grid.BackgroundColor = AppTheme.MacCardBackground;
                        grid.GridColor = AppTheme.MacBorderLight;
                        grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;

                        grid.DefaultCellStyle.BackColor = System.Drawing.Color.White;
                        grid.DefaultCellStyle.ForeColor = AppTheme.MacTextPrimary;
                        grid.DefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 8.5F);
                        grid.DefaultCellStyle.SelectionBackColor = AppTheme.MacBlue;
                        grid.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.White;
                        grid.DefaultCellStyle.Padding = new Padding(6, 3, 6, 3);

                        grid.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(249, 249, 251);

                        grid.EnableHeadersVisualStyles = false;
                        grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
                        grid.ColumnHeadersDefaultCellStyle.BackColor = AppTheme.MacHeaderGradient1;
                        grid.ColumnHeadersDefaultCellStyle.ForeColor = AppTheme.MacTextPrimary;
                        grid.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
                        grid.ColumnHeadersDefaultCellStyle.Padding = new Padding(8, 6, 8, 6);
                        grid.ColumnHeadersHeight = 36;

                        grid.RowTemplate.Height = 32;
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Áp dụng style MacOS cho tất cả RichTextBox - Không viền (borderless)
        /// Chỉ dùng màu background để phân biệt ô nhập liệu
        /// </summary>
        private void ApplyMacStyleToRichTextBoxes()
        {
            try
            {
                System.Drawing.Font textFont = new System.Drawing.Font("Segoe UI", 8.5F, System.Drawing.FontStyle.Regular);

                foreach (System.Windows.Forms.Control ctrl in GetAllControlsForTheme(this))
                {
                    if (ctrl is RichTextBox rtb)
                    {
                        rtb.Font = textFont;
                        rtb.BackColor = AppTheme.MacInputBackground;
                        rtb.ForeColor = AppTheme.MacTextPrimary;
                        rtb.BorderStyle = BorderStyle.None;  // BỎ VIỀN ĐEN - style hiện đại

                        // Add focus effect - đổi màu khi focus
                        rtb.Enter += (s, e) => { ((RichTextBox)s).BackColor = AppTheme.MacInputBackgroundFocus; };
                        rtb.Leave += (s, e) => { ((RichTextBox)s).BackColor = AppTheme.MacInputBackground; };
                    }
                }
            }
            catch { }
        }

        private void ApplyMacFontsToLabels()
        {
            try
            {
                System.Drawing.Font labelFont = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular);

                foreach (System.Windows.Forms.Control ctrl in GetAllControlsForTheme(this))
                {
                    if (ctrl is Label lbl && lbl != label14)
                    {
                        // Apply BLACK color to ALL labels except label14 (header)
                        lbl.Font = labelFont;
                        lbl.ForeColor = System.Drawing.Color.Black; // Force BLACK
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Áp dụng style MacOS cho tất cả TextBox - Không viền (borderless)
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

                        // Thêm padding bằng cách set Multiline (workaround cho BorderStyle.None)
                        // Không cần padding vì background color đã phân biệt rõ

                        // Add focus effect - đổi màu khi focus
                        txt.Enter += (s, e) => { ((TextBox)s).BackColor = AppTheme.MacInputBackgroundFocus; };
                        txt.Leave += (s, e) => { ((TextBox)s).BackColor = AppTheme.MacInputBackground; };
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Áp dụng style MacOS cho tất cả ComboBox - Flat style không viền
        /// Chỉ dùng màu background để phân biệt ô nhập liệu
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

        private IEnumerable<System.Windows.Forms.Control> GetAllControlsForTheme(System.Windows.Forms.Control container)
        {
            foreach (System.Windows.Forms.Control ctrl in container.Controls)
            {
                yield return ctrl;
                foreach (System.Windows.Forms.Control child in GetAllControlsForTheme(ctrl))
                {
                    yield return child;
                }
            }
        }

        // Các helper cho CCCD
        private void TxtDigitsOnly_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar)) e.Handled = true;
        }
        private void TxtCccd_TextChanged(object sender, EventArgs e)
        {
            // Tùy chọn: áp dụng giới hạn độ dài hoặc format; giữ đơn giản
            try
            {
                var tb = sender as TextBox;
                if (tb == null) return;
                var digits = new string((tb.Text ?? "").Where(char.IsDigit).ToArray());
                if (tb.Text != digits) { var sel = tb.SelectionStart; tb.Text = digits; tb.SelectionStart = Math.Min(sel, tb.Text.Length); }
            }
            catch { }
        }
        private void TxtCccd_Leave(object sender, EventArgs e)
        {
            // Xác thực CCCD phải đúng 12 chữ số
            try
            {
                var tb = sender as TextBox;
                if (tb == null) return;

                var text = tb.Text ?? "";
                if (string.IsNullOrWhiteSpace(text)) return; // Cho phép rỗng (trường tùy chọn)

                var digits = new string(text.Where(char.IsDigit).ToArray());
                if (digits.Length > 0 && digits.Length != 12)
                {
                    MessageBox.Show($"Số CCCD phải đúng 12 chữ số (hiện tại: {digits.Length} chữ số)", 
                                    "Lỗi nhập liệu", 
                                    MessageBoxButtons.OK, 
                                    MessageBoxIcon.Warning);
                    tb.Focus();
                    tb.SelectAll();
                }
            }
            catch { }
        }

        // ========== VALIDATION SỐ ĐIỆN THOẠI ==========
        // Tự động xóa các ký tự không phải số
        private void TxtSdt_TextChanged(object sender, EventArgs e)
        {
            try
            {
                var tb = sender as TextBox;
                if (tb == null) return;

                // Chỉ giữ lại các chữ số
                var digits = new string((tb.Text ?? "").Where(char.IsDigit).ToArray());
                if (tb.Text != digits) 
                { 
                    var sel = tb.SelectionStart; 
                    tb.Text = digits; 
                    tb.SelectionStart = Math.Min(sel, tb.Text.Length); 
                }
            }
            catch { }
        }

        // Validate số điện thoại phải đúng 10 số khi rời khỏi textbox
        private void TxtSdt_Leave(object sender, EventArgs e)
        {
            try
            {
                var tb = sender as TextBox;
                if (tb == null) return;

                var text = tb.Text ?? "";
                if (string.IsNullOrWhiteSpace(text)) return; // Cho phép rỗng (trường tùy chọn)

                var digits = new string(text.Where(char.IsDigit).ToArray());

                // Kiểm tra phải đúng 10 số - không ít hơn, không nhiều hơn
                if (digits.Length > 0 && digits.Length != 10)
                {
                    MessageBox.Show(
                        $"Số điện thoại phải có đúng 10 chữ số (hiện tại: {digits.Length} chữ số)\n" +
                        "Ví dụ hợp lệ: 0987654321", 
                        "Lỗi nhập liệu", 
                        MessageBoxButtons.OK, 
                        MessageBoxIcon.Warning
                    );
                    tb.Focus();
                    tb.SelectAll();
                }
            }
            catch { }
        }

        private void DateNgaycapCCCD_ValueChanged(object sender, EventArgs e)
        {
            // Hành vi tùy chọn: điều chỉnh Noicap dựa trên ngày cắt (01/07/2024). Giữ tối thiểu: không làm gì
        }

        // Xác thực tất cả các DateTimePicker để ngăn ngày tương lai
        private void DatePicker_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                var picker = sender as DateTimePicker;
                if (picker == null) return;

                // Nếu ngày trong tương lai, reset về hôm nay
                if (picker.Value.Date > DateTime.Today)
                {
                    picker.Value = DateTime.Today;
                    MessageBox.Show("Ngày không được lớn hơn ngày hiện tại.", "Lỗi nhập liệu", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch { }
        }

        // Format số tiền
        private void CbMoney_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.' && e.KeyChar != ',') e.Handled = true;
        }

        private void CbMoney_TextChanged(object sender, EventArgs e)
        {
            if (suppressMoneyChange) return;
            try
            {
                var cb = sender as ComboBox;
                if (cb == null) return;
                var txt = cb.Text ?? string.Empty;
                var digits = new string(txt.Where(char.IsDigit).ToArray());
                if (string.IsNullOrEmpty(digits)) return;
                var value = ParseMoneyStringToLong(digits);
                var formatted = string.Format(CultureInfo.InvariantCulture, "{0:N0}", value).Replace(",", ".");
                if (formatted != txt)
                {
                    suppressMoneyChange = true;
                    cb.Text = formatted;
                    cb.SelectionStart = cb.Text.Length;
                    suppressMoneyChange = false;
                }
            }
            catch { }
        }

        // Tính toán Sotientong và Sotienchu nếu có thể
        private void UpdateComputedFields(Customer c)
        {
            if (c == null) return;
            try
            {
                // Tính tổng là Vốn tự có (Vtc) + Vốn vay (Sotien)
                if (string.IsNullOrWhiteSpace(c.Sotientong))
                {
                    long loan = ParseMoneyStringToLong(c.Sotien);
                    long own = ParseMoneyStringToLong(c.Vtc);
                    long total = loan + own;
                    if (total > 0)
                    {
                        // format với dấu phân cách hàng nghìn dùng '.' làm dấu phân cách nghìn
                        var formatted = string.Format(CultureInfo.InvariantCulture, "{0:N0}", total).Replace(",", ".");
                        c.Sotientong = formatted;
                    }
                }

                // Tạo Sotienchu nếu thiếu và có giá trị số trong Sotientong
                if (string.IsNullOrWhiteSpace(c.Sotienchu) && !string.IsNullOrWhiteSpace(c.Sotientong))
                {
                    var digits = new string((c.Sotientong ?? "").Where(char.IsDigit).ToArray());
                    if (long.TryParse(digits, out long v) && v > 0)
                    {
                        var words = NumberToVietnameseWords(v);
                        if (!string.IsNullOrWhiteSpace(words)) c.Sotienchu = words + " đồng";
                    }
                }
            }
            catch { }
        }

        private IEnumerable<string> GetTemplateNamesForCustomer(Customer c, bool include03)
        {
            var list = new List<string>();
            // Nếu chương trình được chọn cho thấy "sản xuất kinh doanh" (SXKD) thì dùng mẫu 01 SXKD cụ thể
            try
            {
                var ct = (c?.Chuongtrinh ?? "").Trim();
                if (!string.IsNullOrEmpty(ct) && IsSxkdChuongtrinh(ct))
                {
                    list.Add("01 SXKD.docx");
                }
                else if (!string.IsNullOrEmpty(ct) && IsGqvlChuongtrinh(ct))
                {
                    // Dùng biến thể GQVL khi chương trình chỉ ra GQVL
                    list.Add("01 GQVL.docx");
                }
                else
                {
                    list.Add("01 HN.docx");
                }
            }
            catch
            {
                list.Add("01 HN.docx");
            }
            if (include03) list.Add("03 DS.docx");
            // Lưu ý: GUQ chỉ nên được xuất khi người dùng bấm btnGUQ rõ ràng
            return list;
         }

        // Phát hiện các biến thể phổ biến chỉ ra chương trình GQVL
        private bool IsGqvlChuongtrinh(string chuongtrinh)
        {
            if (string.IsNullOrWhiteSpace(chuongtrinh)) return false;
            try
            {
                string Normalize(string s)
                {
                    var formD = s.Normalize(System.Text.NormalizationForm.FormD);
                    var sb = new System.Text.StringBuilder();
                    foreach (var ch in formD)
                    {
                        var uc = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(ch);
                        if (uc != System.Globalization.UnicodeCategory.NonSpacingMark)
                            sb.Append(ch);
                    }
                    return sb.ToString().Normalize(System.Text.NormalizationForm.FormC).ToLowerInvariant();
                }

                var n = Normalize(chuongtrinh);
                // Kiểm tra các từ viết tắt/biến thể phổ biến cho GQVL
                if (n.Contains("gqvl") || n.Contains("gq vl") || n.Contains("gq-vl") || n.Contains("gq_vl"))
                    return true;

                // Kiểm tra cụm từ đầy đủ "Giải quyết việc làm duy trì và mở rộng việc làm"
                if (n.Contains("giai quyet viec lam duy tri") ||
                    n.Contains("giai quyet viec lam") && n.Contains("duy tri") && n.Contains("mo rong"))
                    return true;
            }
            catch { }
            return false;
        }

        // Phát hiện các biến thể phổ biến chỉ ra "Sản xuất kinh doanh" (SXKD)
        private bool IsSxkdChuongtrinh(string chuongtrinh)
        {
            if (string.IsNullOrWhiteSpace(chuongtrinh)) return false;
            try
            {
                string Normalize(string s)
                {
                    var formD = s.Normalize(System.Text.NormalizationForm.FormD);
                    var sb = new System.Text.StringBuilder();
                    foreach (var ch in formD)
                    {
                        var uc = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(ch);
                        if (uc != System.Globalization.UnicodeCategory.NonSpacingMark)
                            sb.Append(ch);
                    }
                    return sb.ToString().Normalize(System.Text.NormalizationForm.FormC).ToLowerInvariant();
                }

                var n = Normalize(chuongtrinh);

                // Kiểm tra cụm từ chính xác "Hộ gia đình Sản xuất kinh doanh tại vùng khó khăn"
                if (n.Contains("ho gia dinh") && n.Contains("san xuat kinh doanh") && n.Contains("vung kho khan"))
                    return true;

                // Kiểm tra các từ viết tắt/biến thể phổ biến
                if (n.Contains("sxkd"))
                    return true;

                // Kiểm tra từ khóa chung SXKD
                if (n.Contains("san xuat kinh doanh"))
                    return true;
            }
            catch { }
            return false;
        }

        private void txtSocccd_TextChanged(object sender, EventArgs e)
        {

        }
    }

    #region Models

    public class Customer
{
    public string Hoten { get; set; }
    public string Socccd { get; set; }
    public string GioiTinh { get; set; }
    public string Nhandang { get; set; }
    public DateTime Ngaycap { get; set; }
    public DateTime Ngaysinh { get; set; }
    public string Noicap { get; set; }
    public string Xa { get; set; }
    public string Thon { get; set; }
    public string Hoi { get; set; }
    public string To { get; set; }
    public string Totruong { get; set; }
    public string Chuongtrinh { get; set; }
    public string Vtc { get; set; }
    public string Phuongan { get; set; }
    public string Thoihanvay { get; set; }
    public string Phanky { get; set; }
    public string Sotien { get; set; }
    public string Sotien1 { get; set; }
    public string Sotien2 { get; set; }
    public string Soluong1 { get; set; }
    public string Soluong2 { get; set; }
    public string Sotientong { get; set; }
    public string Sotienchu { get; set; }
    public string Mucdich1 { get; set; }
    public string Mucdich2 { get; set; }
    public string Doituong1 { get; set; }
    public string Doituong2 { get; set; }
    public DateTime Ngaylaphs { get; set; }
    public DateTime Ngaydenhan { get; set; }
    public DateTime Thoihancccd { get; set; }
    public string PGD { get; set; }

    public string Dantoc { get; set; }
    public string Sdt = "";

    public string Ntk1 = "";
    public string Ntk2 = "";
    public string Ntk3 = "";
    public string CccdNtk1 = "";
    public string CccdNtk2 = "";
    public string CccdNtk3 = "";
    public string Namsinh1 = "";
    public string Namsinh2 = "";
    public string Namsinh3 = "";

    public string Qh1 = "";
    public string Qh2 = "";
    public string Qh3 = "";

    [JsonIgnore]
    public string _fileName { get; set; }
}

internal class XinManModel { public string pgd { get; set; } public List<Commune> communes { get; set; } = new List<Commune>(); }
internal class Commune { public string name { get; set; } public List<Association> associations { get; set; } = new List<Association>(); public List<Village> villages { get; set; } = new List<Village>(); }
internal class Association { public string name { get; set; } public string code { get; set; } public List<Village> villages { get; set; } = new List<Village>(); public List<string> managedVillages { get; set; } = new List<string>(); }
internal class Village { public string name { get; set; } public List<string> groups { get; set; } = new List<string>(); }

#endregion
}
