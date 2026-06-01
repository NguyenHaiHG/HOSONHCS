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
using Word = Microsoft.Office.Interop.Word;

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
        // Model chứa dữ liệu tỉnh (nhiều PGD trong một file) - ví dụ: tuyenquang.json
        private TinhModel currentTinhModel;
        // Editor để chỉnh sửa xinman.json trên tab3 (cần login)
        private XinManEditor xinManEditor;
        // TinhModel đang được chọn trong cbTinhfix (dùng để populate cbpgdfix → dgv1)
        private TinhModel _editTinhModel = null;

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

        // ========== BẢNG KÊ TIỀN ==========
        // Danh sách bảng kê tiền theo tổ trưởng
        private System.ComponentModel.BindingList<BangKeData> bangKeList = new System.ComponentModel.BindingList<BangKeData>();
        // Danh sách phương án hợp lệ cho cbPhuongan và cbmucdich1
        private static readonly string[] PhuonganList = new string[] {
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
            try { btn01TGTV.Click += Btn01tgtv_Click; } catch { /* bỏ qua nếu control không tồn tại */ }
            try { btnBia.Click += BtnBia_Click; } catch { /* bỏ qua nếu control không tồn tại */ }
            try { btnDn.Click += BtnDn_Click; } catch { /* bỏ qua nếu control không tồn tại */ }
            try { btnall.Click += Btnall_Click; } catch { /* bỏ qua nếu control không tồn tại */ }

            // Các nút tạo khách hàng mới và đăng xuất
            try { btntaokh.Click += BtnTaokh_Click; } catch { /* bỏ qua nếu control không tồn tại */ }
            try { btnexit.Click += BtnExit_Click; } catch { /* bỏ qua nếu control không tồn tại */ }

            // ========== CÁC NÚT BẢNG KÊ TIỀN ==========
            try { btnLuubangke.Click += BtnLuubangke_Click; } catch { /* bỏ qua nếu control không tồn tại */ }
            try { btnXoabangke.Click += BtnXoabangke_Click; } catch { /* bỏ qua nếu control không tồn tại */ }
            try { btnTaobangke.Click += BtnTaobangke_Click; } catch { /* bỏ qua nếu control không tồn tại */ }

            // (Đã xoá btnUpdate)

            // ========== NÚT CẬP NHẬT VÀ XOÁ FORM ==========
            try { btnUpdate.Click += BtnUpdate_Click; } catch { /* bỏ qua nếu control không tồn tại */ }
            try { btnClear.Click += BtnClear_Click; } catch { /* bỏ qua nếu control không tồn tại */ }
            try { btnLoaddata.Click += btnLoaddata_Click; } catch { /* bỏ qua nếu control không tồn tại */ }

            // ComboBox Fix PGD - cho phép chọn nhanh và reload dữ liệu PGD
            try { cbpgdfix.SelectedIndexChanged += CbPgdFix_SelectedIndexChanged; } catch { /* bỏ qua nếu control không tồn tại */ }

            // ========== TỰ ĐỘNG VIẾT HOA CHỮ CÁI ĐẦU CHO Ô NHẬP TÊN ==========
            // Các ô nhập tên (txtHoten, txtntk1/2/3/4): tự động viết hoa chữ cái đầu mỗi từ (Title Case)
            try { txtHoten.TextChanged += TxtName_TextChanged; txtHoten.Leave += TxtName_Leave; } catch { }
            try { txtntk1.TextChanged += TxtName_TextChanged; txtntk1.Leave += TxtName_Leave; } catch { }
            try { txtntk2.TextChanged += TxtName_TextChanged; txtntk2.Leave += TxtName_Leave; } catch { }
            try { txtntk3.TextChanged += TxtName_TextChanged; txtntk3.Leave += TxtName_Leave; } catch { }
            try { txtntk4.TextChanged += TxtName_TextChanged; txtntk4.Leave += TxtName_Leave; } catch { }

            // ========== DATEPICKER NGÀY SINH NGƯỜI THỪA KẾ (NTK) ==========
            // datentk1/2/3/4 là DateTimePicker, KHÔNG CẦN KeyPress/TextChanged events
            // DateTimePicker tự động xử lý nhập liệu, chỉ cần set MaxDate và ShowCheckBox
            // ========== VALIDATION SỐ CCCD ==========
            // Các ô nhập CCCD (txtSocccd, txtcccd1/2/3/4): chỉ cho phép nhập số, đúng 12 chữ số
            try { txtSocccd.KeyPress += TxtDigitsOnly_KeyPress; txtSocccd.TextChanged += TxtCccd_TextChanged; txtSocccd.Leave += TxtCccd_Leave; txtSocccd.MaxLength = 12; } catch { }
            try { txtcccd1.KeyPress += TxtDigitsOnly_KeyPress; txtcccd1.TextChanged += TxtCccd_TextChanged; txtcccd1.Leave += TxtCccd_Leave; txtcccd1.MaxLength = 12; } catch { }
            try { txtcccd2.KeyPress += TxtDigitsOnly_KeyPress; txtcccd2.TextChanged += TxtCccd_TextChanged; txtcccd2.Leave += TxtCccd_Leave; txtcccd2.MaxLength = 12; } catch { }
            try { txtcccd3.KeyPress += TxtDigitsOnly_KeyPress; txtcccd3.TextChanged += TxtCccd_TextChanged; txtcccd3.Leave += TxtCccd_Leave; txtcccd3.MaxLength = 12; } catch { }
            try { txtcccd4.KeyPress += TxtDigitsOnly_KeyPress; txtcccd4.TextChanged += TxtCccd_TextChanged; txtcccd4.Leave += TxtCccd_Leave; txtcccd4.MaxLength = 12; } catch { }

            // ========== SỐ ĐIỆN THOẠI ==========
            // txtSdt: chỉ cho phép nhập số, tự động format với dấu chấm (0812.801.886)
            try { txtSdt.KeyPress += TxtSdt_KeyPress; txtSdt.TextChanged += TxtSdt_TextChanged; txtSdt.Leave += TxtSdt_Leave; txtSdt.MaxLength = 12; } catch { }  // MaxLength=12 để chứa dấu chấm

            // ========== NHÂN KHẨU ==========
            // txtNhankhau: chỉ cho phép nhập số, tối đa 2 chữ số
            try { txtNhankhau.KeyPress += TxtDigitsOnly_KeyPress; txtNhankhau.MaxLength = 2; } catch { }

            // ========== TỰ ĐỘNG CHỌN NỠI CẤP CCCD ==========
            // dateNgaycapCCCD: tự động điền cbNoicap dựa trên ngày cấp (trước/sau 01/07/2024)
            // cbNoicap: khóa không cho chọn, chỉ tự động điền
            try 
            { 
                if (cbNoicap != null) 
                {
                    cbNoicap.DropDownStyle = ComboBoxStyle.DropDownList;  // Khóa không cho gõ
                    cbNoicap.Enabled = false;  // Disable hoàn toàn
                }
                dateNgaycapCCCD.ValueChanged += DateNgaycapCCCD_ValueChanged; 
            } catch { }

            // ========== VALIDATE NGÀY TRONG TƯƠNG LAI: DÙNG LEAVE THAY VALUECHANGED ==========
            // Bỏ MaxDate để user nhập tự do, chỉ validate khi rời khỏi ô (Leave)
            try 
            {
                if (dateLaphs != null) 
                { 
                    dateLaphs.MaxDate = DateTime.MaxValue;
                    dateLaphs.Leave += DatePickerChecked_Leave;
                }
                if (dateNgaycapCCCD != null) 
                { 
                    dateNgaycapCCCD.MaxDate = DateTime.MaxValue;
                    dateNgaycapCCCD.Format = DateTimePickerFormat.Custom;
                    dateNgaycapCCCD.CustomFormat = " ";
                    dateNgaycapCCCD.Leave += DatePickerChecked_Leave;
                    dateNgaycapCCCD.Enter += DatePicker_Enter;
                }
                if (dateNgaysinh != null) 
                { 
                    dateNgaysinh.MaxDate = DateTime.MaxValue;
                    dateNgaysinh.Format = DateTimePickerFormat.Custom;
                    dateNgaysinh.CustomFormat = " ";
                    dateNgaysinh.Leave += DatePickerChecked_Leave;
                    dateNgaysinh.Enter += DatePicker_Enter;
                    dateNgaysinh.ValueChanged += DateNgaysinh_ValueChanged;
                }
                if (dateDH != null) 
                { 
                    // Ngày đến hạn: cho phép nhập ngày tương lai, không validate
                    dateDH.MaxDate = DateTime.MaxValue;
                }
                if (dateGn != null)
                {
                    dateGn.MaxDate = DateTime.MaxValue;
                }
                if (datendhcccd != null) 
                { 
                    // Thời hạn CCCD: cho phép tương lai, chỉ validate định dạng khi rời ô
                    datendhcccd.MaxDate = DateTime.MaxValue;
                    datendhcccd.Leave += DateThoihanCCCD_ValueChanged;
                }
                // ========== DATEPICKER NGÀY SINH NTK: VALIDATE KHI RỜI Ô ==========
                // datentk1/2/3/4 có ShowCheckBox=true, chỉ validate sau khi nhập xong (Leave)
                // Không dùng MaxDate để tránh reset ngay khi đang gõ
                if (datentk1 != null)
                {
                    datentk1.MaxDate = DateTime.MaxValue;
                    datentk1.Leave += DatePickerChecked_Leave;
                    datentk1.ValueChanged += DateNtk_ValueChanged;
                }
                if (datentk2 != null)
                {
                    datentk2.MaxDate = DateTime.MaxValue;
                    datentk2.Leave += DatePickerChecked_Leave;
                    datentk2.ValueChanged += DateNtk_ValueChanged;
                }
                if (datentk3 != null)
                {
                    datentk3.MaxDate = DateTime.MaxValue;
                    datentk3.Leave += DatePickerChecked_Leave;
                    datentk3.ValueChanged += DateNtk_ValueChanged;
                }
                if (datentk4 != null)
                {
                    datentk4.MaxDate = DateTime.MaxValue;
                    datentk4.Leave += DatePickerChecked_Leave;
                    datentk4.ValueChanged += DateNtk_ValueChanged;
                }
            } catch { }

            // ========== COMBOBOX CHỌN ĐỊA ĐIỂM (PGD, XÃ, THÔN, HỘI) ==========
            // Gắn sự kiện để tự động load dữ liệu cascading từ xinman.json
            try { cbPGD.SelectedIndexChanged += CbPGD_SelectedIndexChanged; } catch { }
            try { cbXa.SelectedIndexChanged += CbXa_SelectedIndexChanged; } catch { }
            try { cbThon.SelectedIndexChanged += CbThon_SelectedIndexChanged; } catch { }
            try { cbHoi.SelectedIndexChanged += CbHoi_SelectedIndexChanged; } catch { }
            try { cbHoi.TextChanged += CbHoi_TextChanged; } catch { }

            // ========== TỰ ĐỘNG ĐIỀN CBDOITUONG DỰA TRÊN CBCHUONGTRINH ==========
            try { cbChuongtrinh.SelectedIndexChanged += CbChuongtrinh_SelectedIndexChanged; } catch { }

            // ========== TỰ ĐỘNG TÍNH NGÀY ĐẾN HẠN ==========
            try { cbThoihanvay.SelectedIndexChanged += CbThoihanvay_SelectedIndexChanged; } catch { }

            // ========== TỰ ĐỘNG ĐIỀU KHIỂN CBMUCDICH1/2 DỰA TRÊN CBPHUONGAN ==========
            try { cbPhuongan.SelectedIndexChanged += CbPhuongan_SelectedIndexChanged; } catch { }
            try { cbPhuongan.TextChanged += CbPhuongan_TextChanged; } catch { }
            // cbmucdich2: luôn mở, cho phép nhập tay kèm danh sách
            try { if (cbmucdich2 != null) { cbmucdich2.Enabled = true; cbmucdich2.DropDownStyle = ComboBoxStyle.DropDown; } } catch { }

            // ========== FORMAT SỐ TIỀN TỰ ĐỘNG ==========
            // cbSotien/cbSotien1/cbSotien2/cbSotien3: chỉ cho nhập số, tự động format với dấu '.' ngăn cách hàng nghìn
            try { cbSotien.KeyPress += CbMoney_KeyPress; cbSotien.TextChanged += CbMoney_TextChanged; } catch { }
            try { cbSotien1.KeyPress += CbMoney_KeyPress; cbSotien1.TextChanged += CbMoney_TextChanged; } catch { }
            try { cbSotien2.KeyPress += CbMoney_KeyPress; cbSotien2.TextChanged += CbMoney_TextChanged; } catch { }
            try { cbSotien3.KeyPress += CbMoney_KeyPress; cbSotien3.TextChanged += CbMoney_TextChanged; } catch { }
            // Kiểm tra tổng cbSotien1+2+3 = cbSotien
            try { cbSotien.TextChanged  += CbSotienSum_Changed; } catch { }
            try { cbSotien1.TextChanged += CbSotienSum_Changed; } catch { }
            try { cbSotien2.TextChanged += CbSotienSum_Changed; } catch { }
            try { cbSotien3.TextChanged += CbSotienSum_Changed; } catch { }
            // Module Cấp nước sạch: chia đôi số tiền khi rời/chọn cbSotien
            try { cbSotien.Leave += CbSotien_CapNuocSach_Leave; } catch { }
            try { cbSotien.SelectedIndexChanged += CbSotien_CapNuocSach_SelectedIndexChanged; } catch { }

            // ========== HIỂN THỊ CHECKBOX CHO CÁC TRƯỜNG NGÀY OPTIONAL ==========
            // Các DateTimePicker có ShowCheckBox = true để user có thể bỏ chọn (không bắt buộc nhập)
            // Mặc định unchecked = không có giá trị
            try 
            {
                // dateLaphs: có checkbox - tick để chọn ngày lập hồ sơ
                if (dateLaphs != null)
                {
                    dateLaphs.ShowCheckBox = true;
                    dateLaphs.Checked = false;
                    dateLaphs.Format = DateTimePickerFormat.Custom;
                    dateLaphs.CustomFormat = " ";
                    dateLaphs.ValueChanged += DateLaphs_CheckedChanged;
                }

                if (dateDH != null) 
                { 
                    // Ngày đến hạn: KHOÁ hoàn toàn - chỉ hiển thị, không thao tác
                    dateDH.ShowCheckBox = false;
                    dateDH.Enabled = false;
                    dateDH.CustomFormat = "          "; // mặc định ẩn, chỉ hiện khi tính ra
                }
                if (dateGn != null)
                {
                    // Ngày giải ngân: có checkbox - tick để chọn
                    dateGn.ShowCheckBox = true;
                    dateGn.Checked = false;
                    dateGn.MaxDate = DateTime.MaxValue;
                }
                // datendhcccd: BỎ checkbox, bắt đầu trống
                if (datendhcccd != null) 
                { 
                    datendhcccd.ShowCheckBox = false;
                    datendhcccd.Format = DateTimePickerFormat.Custom;
                    datendhcccd.CustomFormat = " ";
                    datendhcccd.Enter += DatePicker_Enter;
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
                if (datentk4 != null) 
                { 
                    datentk4.ShowCheckBox = true; 
                    datentk4.Checked = false;  // Ngày sinh NTK 4: optional
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
            // Load xinman.json (dữ liệu PGD/Xã/Thôn/Hội/Tổ) - cbPGD items đã được cấu hình trong Designer
            LoadXinManData();
            // Load danh sách khách hàng từ folder Customers/*.json
            LoadCustomersFromFiles();

            // ========== KHỞI TẠO BẢNG KÊ TIỀN (TABPAGE4) ==========
            try
            {
                if (dgvbangke1 != null)
                {
                    BangKeTien.InitializeDataGridView(dgvbangke1);
                }

                // Khởi tạo dgvTotruong để hiển thị danh sách tổ trưởng đã lưu bảng kê
                if (dgvTotruong != null)
                {
                    InitializeDgvTotruong();

                    // Load danh sách bảng kê từ file JSON
                    LoadBangKeFromFiles();
                }
            }
            catch { }

            // ========== KHỞI TẠO TAB GHI CHÚ (TABPAGE5) ==========
            try { InitializeGhiChuTab(); } catch { }

            // ========== KHỞI TẠO CHATBOT (TABPAGE2) ==========
            try { InitializeChatbotTab(); } catch { }

            // ========== KHỞI TẠO TAB GIỚI THIỆU (TABPAGE6) ==========
            try { InitializeTab6(); } catch { }

            // ========== KHỞI TẠO TAB XEM VĂN BẢN WORD (TABPAGE7) ==========
            try { InitializeTab7(); } catch { }

            // Bind dữ liệu vào DataGridView
            BindGrid();

            // ========== NÂNG CẤP GIAO DIỆN TOÀN APP ==========
            try { ApplyModernStyle(); } catch { }

            // ========== KIỂM TRA CẬP NHẬT TỰ ĐỘNG KHI KHỞI ĐỘNG ==========
            try
            {
                CheckForUpdateOnStartup();
            }
            catch { }

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

            // ========== THÔNG BÁO NGÀY LỄ / SỰ KIỆN ĐẶC BIỆT ==========
            // Fetch thông báo từ GitHub và hiện popup nếu hôm nay có sự kiện
            // Gọi sau khi form đã hiển thị hoàn toàn (Shown event) để Invoke không bị lỗi
            this.Shown += async (s, ev) => await HolidayNoticeChecker.CheckAndShowAsync(this);
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

                // Ẩn tất cả cột, chỉ hiện 5 cột cần thiết
                foreach (DataGridViewColumn col in dgv.Columns)
                {
                    col.ReadOnly = true;
                    col.Visible = false;
                }

                // Cấu hình 5 cột hiển thị: tên, xã, thôn, chương trình, CCCD
                var visibleCols = new[]
                {
                    new { Name = "Hoten",        Header = "Họ và tên",        Width = 180, Index = 0 },
                    new { Name = "Xa",           Header = "Xã",               Width = 110, Index = 1 },
                    new { Name = "Thon",         Header = "Thôn",             Width = 100, Index = 2 },
                    new { Name = "Chuongtrinh",  Header = "Chương trình vay", Width = 130, Index = 3 },
                    new { Name = "Socccd",       Header = "Căn cước",         Width = 120, Index = 4 },
                };

                foreach (var c in visibleCols)
                {
                    if (dgv.Columns[c.Name] == null) continue;
                    dgv.Columns[c.Name].Visible      = true;
                    dgv.Columns[c.Name].HeaderText   = c.Header;
                    dgv.Columns[c.Name].Width        = c.Width;
                    dgv.Columns[c.Name].DisplayIndex = c.Index;
                }
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

            // Folder đặt tên theo họ tên tổ trưởng + ngày tháng năm
            var totruong = (!string.IsNullOrWhiteSpace(c.Totruong) ? c.Totruong : c.Hoten).Trim();
            var safeTotruong = MakeFileSystemSafe(totruong);
            var dateSuffix = DateTime.Now.ToString("dd-MM-yyyy");
            var folder = Path.Combine(root, safeTotruong + "_" + dateSuffix);

            // Tất cả khách hàng cùng tổ trưởng trong cùng ngày dùng chung folder
            return folder;
        }

        private List<string> CreateProfileFromTemplate(Customer c, bool include03)
        {
            var createdFiles = new List<string>();
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
                    var dateName = DateTime.Now.ToString("dd-MM-yyyy");
                    var baseName = MakeFileSystemSafe(c.Hoten) + "_" + dateName + "_" + shortName + ".docx";

                    bool is01Template = templateName == "01 GQVL.docx" || templateName == "01 SXKD.docx" || templateName == "01 HN.docx";
                    string destDoc;
                    if (is01Template)
                    {
                        // Chỉ cho phép tối đa 2 hồ sơ cùng tên cho mẫu 01 trong cùng 1 thư mục
                        var existingCount = Directory.GetFiles(destFolder, "*_" + shortName + "*.docx").Length;
                        if (existingCount >= 2) continue;
                        destDoc = UniqueFilePath(Path.Combine(destFolder, baseName));
                    }
                    else
                    {
                        // Ghi đè: xóa file cũ cùng loại trước khi tạo mới
                        foreach (var oldFile in Directory.GetFiles(destFolder, "*_" + shortName + "*.docx"))
                        {
                            try { File.Delete(oldFile); } catch { }
                        }
                        destDoc = Path.Combine(destFolder, baseName);
                    }

                    File.Copy(templatePath, destDoc, true);

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

                    // Thêm vào danh sách file đã tạo
                    createdFiles.Add(destDoc);
                }
            }
            finally
            {
                foreach (var f in tempFilesToDelete) { try { if (File.Exists(f)) File.Delete(f); } catch { } }
            }

            return createdFiles;
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
                { "{{cccd12}}", SplitCCCDInto12Boxes(c.Socccd) },
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
                { "{{ngaydenhan}}", c.Ngaydenhan == DateTime.MinValue ? "...../...../....." : c.Ngaydenhan.ToString("dd/MM/yyyy") },
                { "{{thoihanvay}}", c.Thoihanvay },
                { "{{sotien}}", c.Sotien },
                { "{{sotien1}}", c.Sotien1 },
                { "{{sotien2}}", c.Sotien2 },
                { "{{sotien3}}", c.Sotien3 ?? "" },
                { "{{sotientong}}", c.Sotientong },
                { "{{sotienchu}}", string.IsNullOrEmpty(c.Sotienchu) ? "" : char.ToUpper(c.Sotienchu[0]) + c.Sotienchu.Substring(1) },
                { "{{soluong1}}", c.Soluong1 ?? "" },
                { "{{soluong2}}", string.IsNullOrWhiteSpace(c.Soluong2) ? "" : c.Soluong2 },
                { "{{soluong3}}", c.Soluong3 ?? "" },
                // Ưu tiên trường Mucdich nhập tự do; nếu trống, dùng Doituong (combo) làm dự phòng
                { "{{mucdich1}}", !string.IsNullOrWhiteSpace(c.Mucdich1) ? c.Mucdich1 : (c.Doituong1 ?? "") },
                { "{{mucdich2}}", !string.IsNullOrWhiteSpace(c.Mucdich2) ? c.Mucdich2 : (c.Doituong2 ?? "") },
                { "{{mucdich3}}", !string.IsNullOrWhiteSpace(c.Mucdich3) ? c.Mucdich3 : "" },
                { "{{doituong1}}", c.Doituong1 ?? "" },
                { "{{doituong2}}", c.Doituong2 ?? "" },
                { "{{doituong}}", !string.IsNullOrWhiteSpace(c.Doituong1) ? c.Doituong1 : (c.Doituong2 ?? "") },
                { "{{ngaylaphs}}", NgayLaphsFormatter.GetNgaylaphsValue(docPath, c.Ngaylaphs, CountNguoiThuaKe(c)) },
                { "{{ngaysinh}}", c.Ngaysinh == DateTime.MinValue ? "" : c.Ngaysinh.ToString("dd/MM/yyyy") },
                { "{{ngay}}", c.Ngaysinh == DateTime.MinValue ? "" : c.Ngaysinh.Day.ToString() },
                { "{{thang}}", c.Ngaysinh == DateTime.MinValue ? "" : c.Ngaysinh.Month.ToString() },
                { "{{nam}}", c.Ngaysinh == DateTime.MinValue ? "" : c.Ngaysinh.Year.ToString() },
                { "{{phanky}}", c.Phanky },
                { "{{pgd}}", c.PGD },
                { "{{tinh}}", c.Tinh ?? "" },
                { "{{thoihancccd}}", !string.IsNullOrEmpty(c.ThoihancccdText) ? c.ThoihancccdText : (c.Thoihancccd == DateTime.MinValue ? "" : c.Thoihancccd.ToString("dd/MM/yyyy")) },
                { "{{khau}}", c.Nhankhau ?? "" },
                { "{{ld}}", (CountNguoiThuaKe(c) + 1).ToString() },

                // GUQ
                { "{{ntk1}}", c.Ntk1 ?? "" },
                { "{{ntk2}}", c.Ntk2 ?? "" },
                { "{{ntk3}}", c.Ntk3 ?? "" },
                { "{{ntk4}}", c.Ntk4 ?? "" },
                { "{{cccdntk1}}", c.CccdNtk1 ?? "" },
                { "{{cccdntk2}}", c.CccdNtk2 ?? "" },
                { "{{cccdntk3}}", c.CccdNtk3 ?? "" },
                { "{{cccdntk4}}", c.CccdNtk4 ?? "" },
                { "{{qh1}}", c.Qh1 ?? "" },
                { "{{qh2}}", c.Qh2 ?? "" },
                { "{{qh3}}", c.Qh3 ?? "" },
                { "{{qh4}}", c.Qh4 ?? "" },
                { "{{namsinh1}}", FormatNamsinhStringForDoc(c.Namsinh1, docPath) },
                { "{{namsinh2}}", FormatNamsinhStringForDoc(c.Namsinh2, docPath) },
                { "{{namsinh3}}", FormatNamsinhStringForDoc(c.Namsinh3, docPath) },
                { "{{namsinh4}}", FormatNamsinhStringForDoc(c.Namsinh4, docPath) },
                { "{{namsinh1_year}}", ExtractYearString(c.Namsinh1) ?? "" },
                { "{{namsinh2_year}}", ExtractYearString(c.Namsinh2) ?? "" },
                { "{{namsinh3_year}}", ExtractYearString(c.Namsinh3) ?? "" },
                { "{{namsinh4_year}}", ExtractYearString(c.Namsinh4) ?? "" },
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

            // Đối với mẫu 01 GQVL: {{chiphi}} = sotien + 40%, {{thunhap}} = chiphi + sotien
            try
            {
                if (Is01GQVL(docPath))
                {
                    long sotienVal = ParseMoneyStringToLong(c.Sotien);
                    long chiPhiVal = (long)Math.Round(sotienVal * 1.4);
                    long thuNhapVal = chiPhiVal + sotienVal;
                    replacements["{{chiphi}}"] = string.Format(CultureInfo.InvariantCulture, "{0:N0}", chiPhiVal).Replace(",", ".");
                    replacements["{{thunhap}}"] = string.Format(CultureInfo.InvariantCulture, "{0:N0}", thuNhapVal).Replace(",", ".");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR in 01 GQVL chiphi/thunhap: {ex.Message}");
            }

            try { ReplacePlaceholdersUsingOpenXml(docPath, replacements, c); }
            catch (Exception ex) { MessageBox.Show("Error replacing placeholders (OpenXML): " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            // Mẫu 01 HN: căn giữa paragraph chứa ngaylaphs khi bỏ tick
            try
            {
                if (Is01HN(docPath) && c.Ngaylaphs == DateTime.MinValue)
                    CenterParagraphContaining(docPath, NgayLaphsFormatter.PlaceholderNgayThangNam);
            }
            catch { }
         }

        // Thử tính toán `Sotienchu` từ giá trị số `Sotien` (hoặc `Sotientong`) khi tạo mẫu
        private void EnsureSotienchuFromNumeric(Customer c, string docPath)
        {
            try
            {
                if (c == null) return;
                // Chỉ áp dụng cho mẫu 01 (embedded hoặc tên file chứa '01')
                // Áp dụng cho mọi mẫu có {{sotienchu}} — không giới hạn theo tên file

                var source = !string.IsNullOrWhiteSpace(c.Sotien) ? c.Sotien : (!string.IsNullOrWhiteSpace(c.Sotientong) ? c.Sotientong : "");
                if (string.IsNullOrWhiteSpace(source)) return;
                var digits = new string((source ?? "").Where(char.IsDigit).ToArray());
                if (string.IsNullOrWhiteSpace(digits)) return;
                if (!long.TryParse(digits, out long value)) return;
                if (value <= 0) return;
                var words = NumberToVietnameseWords(value);
                if (!string.IsNullOrWhiteSpace(words)) c.Sotienchu = char.ToUpper(words[0]) + words.Substring(1) + " đồng";
            }
            catch { }
        }

        private string SplitCCCDInto12Boxes(string cccd)
        {
            if (string.IsNullOrWhiteSpace(cccd)) return "";

            var digits = new string(cccd.Where(char.IsDigit).ToArray());

            if (digits.Length < 12)
            {
                digits = digits.PadRight(12, ' ');
            }
            else if (digits.Length > 12)
            {
                digits = digits.Substring(0, 12);
            }

            var separated = string.Join(" ", digits.ToCharArray());
            return separated;
        }

        private int CountNguoiThuaKe(Customer c)
        {
            if (c == null) return 0;

            int count = 0;
            if (!string.IsNullOrWhiteSpace(c.Ntk1)) count++;
            if (!string.IsNullOrWhiteSpace(c.Ntk2)) count++;
            if (!string.IsNullOrWhiteSpace(c.Ntk3)) count++;
            if (!string.IsNullOrWhiteSpace(c.Ntk4)) count++;

            return count;
        }

        private void ProcessCCCD12Placeholder(MainDocumentPart mainPart, string cccd)
        {
            if (mainPart == null || string.IsNullOrWhiteSpace(cccd)) return;

            try
            {
                var digits = new string(cccd.Where(char.IsDigit).ToArray());
                if (digits.Length < 12)
                {
                    digits = digits.PadRight(12, '0');
                }
                else if (digits.Length > 12)
                {
                    digits = digits.Substring(0, 12);
                }

                bool replacedInTable = false;

                foreach (var table in mainPart.Document.Descendants<Table>())
                {
                    foreach (var row in table.Elements<TableRow>())
                    {
                        var cells = row.Elements<TableCell>().ToList();

                        for (int cellIdx = 0; cellIdx < cells.Count; cellIdx++)
                        {
                            var cell = cells[cellIdx];
                            var cellText = string.Concat(cell.Descendants<Text>().Select(t => t.Text ?? ""));

                            if (cellText.IndexOf("{{cccd12}}", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Ô chứa {{cccd12}} sẽ là ô đầu tiên chứa chữ số đầu tiên
                                // 11 ô tiếp theo (bên phải) sẽ chứa 11 chữ số còn lại
                                if (cells.Count - cellIdx >= 12)  // Cần đủ 12 ô
                                {
                                    // Điền 12 chữ số vào 12 ô (bắt đầu từ ô hiện tại)
                                    for (int i = 0; i < 12; i++)
                                    {
                                        var targetCell = cells[cellIdx + i];

                                        // Xóa hết nội dung cũ (bao gồm "Số:", {{cccd12}}, v.v.) và điền số mới
                                        foreach (var para in targetCell.Elements<Paragraph>())
                                        {
                                            para.RemoveAllChildren();
                                            var run = new Run();
                                            var text = new Text(digits[i].ToString());
                                            run.Append(text);
                                            para.Append(run);
                                        }

                                        if (!targetCell.Elements<Paragraph>().Any())
                                        {
                                            var paragraph = new Paragraph();
                                            var run = new Run();
                                            var text = new Text(digits[i].ToString());
                                            run.Append(text);
                                            paragraph.Append(run);
                                            targetCell.Append(paragraph);
                                        }
                                    }
                                    replacedInTable = true;
                                    mainPart.Document.Save();
                                    return;
                                }
                            }
                        }
                    }
                }

                if (!replacedInTable)
                {
                    foreach (var para in mainPart.Document.Descendants<Paragraph>())
                    {
                        var paraText = string.Concat(para.Descendants<Text>().Select(t => t.Text ?? ""));

                        if (paraText.IndexOf("{{cccd12}}", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Giữ lại "Số: " và thay {{cccd12}} bằng table 12 ô
                            // Xóa paragraph cũ và tạo table 13 cột (cột đầu: "Số:", 12 cột sau: số)
                            var table = CreateCCCD12Table(digits);
                            para.Parent.InsertAfter(table, para);
                            para.Remove();

                            mainPart.Document.Save();
                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ProcessCCCD12Placeholder: {ex.Message}");
            }
        }

        private Table CreateCCCD12Table(string digits)
        {
            var table = new Table();

            var tblPr = new TableProperties();
            var tblStyle = new TableStyle() { Val = "TableGrid" };
            // 13 cột: cột đầu "Số:" (800 DXA) + 12 cột số (300 DXA mỗi cột) = 4400 DXA
            var tblWidth = new TableWidth() { Width = "4400", Type = TableWidthUnitValues.Dxa };
            var tblJc = new TableJustification() { Val = TableRowAlignmentValues.Left };

            // Table borders - tắt hết viền table level
            var tblBorders = new TableBorders(
                new TopBorder() { Val = BorderValues.None },
                new BottomBorder() { Val = BorderValues.None },
                new LeftBorder() { Val = BorderValues.None },
                new RightBorder() { Val = BorderValues.None },
                new InsideHorizontalBorder() { Val = BorderValues.None },
                new InsideVerticalBorder() { Val = BorderValues.None }
            );

            tblPr.Append(tblStyle);
            tblPr.Append(tblWidth);
            tblPr.Append(tblJc);
            tblPr.Append(tblBorders);
            table.AppendChild(tblPr);

            var tableGrid = new TableGrid();
            // Cột đầu tiên: "Số:" - tăng lên 800 để THẤY RÕ CHỮ "Số:"
            tableGrid.Append(new GridColumn() { Width = "800" });
            // 12 cột chứa số CCCD - width 300 DXA để ô LỚN, RÕ RÀNG
            for (int i = 0; i < 12; i++)
            {
                tableGrid.Append(new GridColumn() { Width = "300" });
            }
            table.Append(tableGrid);

            var tr = new TableRow();
            var trPr = new TableRowProperties();
            var trHeight = new TableRowHeight() 
            { 
                Val = 300,  // Height 300 DXA để ô vuông lớn (300x300)
                HeightType = HeightRuleValues.Exact
            };
            trPr.Append(trHeight);
            tr.Append(trPr);

            // 13 ô: ô đầu "Số:" (không viền) + 12 ô số (có viền)
            for (int i = 0; i < 13; i++)
            {
                var tc = new TableCell();
                var tcPr = new TableCellProperties();

                if (i == 0)
                {
                    // Ô đầu tiên: "Số:" - KHÔNG VIỀN, WIDTH = 800 DXA ĐỂ NHÌN RÕ
                    var tcWidth = new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "800" };
                    var tcBorders = new TableCellBorders(
                        new TopBorder() { Val = BorderValues.None, Size = 0 },
                        new BottomBorder() { Val = BorderValues.None, Size = 0 },
                        new LeftBorder() { Val = BorderValues.None, Size = 0 },
                        new RightBorder() { Val = BorderValues.None, Size = 0 }
                    );
                    var shading = new Shading() { Fill = "FFFFFF" };
                    var tcVAlign = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
                    tcPr.Append(tcWidth);
                    tcPr.Append(tcBorders);
                    tcPr.Append(shading);
                    tcPr.Append(tcVAlign);
                }
                else
                {
                    // 12 ô chứa số CCCD - CÓ VIỀN ĐEN ĐẬM, 300x300 DXA = ô vuông lớn
                    var tcWidth = new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "300" };
                    var tcBorders = new TableCellBorders(
                        new TopBorder() { Val = BorderValues.Single, Size = 8, Color = "000000" },  // Viền đậm hơn: size 8
                        new BottomBorder() { Val = BorderValues.Single, Size = 8, Color = "000000" },
                        new LeftBorder() { Val = BorderValues.Single, Size = 8, Color = "000000" },
                        new RightBorder() { Val = BorderValues.Single, Size = 8, Color = "000000" }
                    );
                    var shading = new Shading() { Fill = "FFFFFF" };
                    var tcVAlign = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
                    tcPr.Append(tcWidth);
                    tcPr.Append(tcBorders);
                    tcPr.Append(shading);
                    tcPr.Append(tcVAlign);
                }

                tc.Append(tcPr);

                var paragraph = new Paragraph();
                var pPr = new ParagraphProperties();
                var justification = new Justification() { Val = (i == 0) ? JustificationValues.Right : JustificationValues.Center };
                pPr.Append(justification);
                paragraph.Append(pPr);

                var run = new Run();
                var runProperties = new RunProperties();
                var fontSize = new FontSize() { Val = "24" };  // Font 24pt - vừa đủ, KHÔNG BÔI ĐẬM
                runProperties.Append(fontSize);
                // KHÔNG thêm Bold để chữ dễ nhìn hơn
                run.Append(runProperties);

                // Ô đầu: "Số:", các ô sau: số CCCD
                var textContent = (i == 0) ? "Số:" : digits[i - 1].ToString();
                var text = new Text(textContent);
                run.Append(text);
                paragraph.Append(run);

                tc.Append(paragraph);
                tr.Append(tc);
            }

            table.Append(tr);

            return table;
        }

        private string ExportSpecificTemplate(Customer c, string templateFileName)
        {
            try
            {
                var destFolder = GetProfileFolderPath(c);
                Directory.CreateDirectory(destFolder);

                string templatePath = ResolveTemplatePath(templateFileName);

                if (!IsDocxFile(templatePath))
                {
                    throw new Exception($"Template \"{templateFileName}\" không hợp lệ hoặc bị hỏng.");
                }

                var shortName = Path.GetFileNameWithoutExtension(templateFileName).Replace(" ", "_");
                var dateName = DateTime.Now.ToString("dd-MM-yyyy");
                var basePath = Path.Combine(destFolder, MakeFileSystemSafe(c.Hoten) + "_" + dateName + "_" + shortName + ".docx");
                // GUQ, 01TGTV, BIA: ghi đè (1 khách 1 file duy nhất)
                // Các mẫu còn lại (03 DS...): thêm _2, _3... không xóa file cũ
                bool isOverwriteTemplate =
                    templateFileName.Equals("GUQ.docx",    StringComparison.OrdinalIgnoreCase) ||
                    templateFileName.Equals("01TGTV.docx", StringComparison.OrdinalIgnoreCase) ||
                    templateFileName.Equals("BIA.docx",    StringComparison.OrdinalIgnoreCase);
                var destDoc = isOverwriteTemplate ? basePath : UniqueFilePath(basePath);

                File.Copy(templatePath, destDoc, true);

                if (!IsDocxFile(destDoc))
                {
                    throw new Exception($"Không thể tạo file từ template \"{templateFileName}\".");
                }

                ReplacePlaceholdersInWord(destDoc, c);

                return destDoc;
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi khi xuất template {templateFileName}: {ex.Message}", ex);
            }
        }

        // Phân tích chuỗi số thuần túy với dấu chấm/phẩy đã bị xóa trong caller; wrapper giữ lại cho rõ ràng
        private static string UniqueFilePath(string path)
        {
            if (!File.Exists(path)) return path;
            var dir = Path.GetDirectoryName(path);
            var nameNoExt = Path.GetFileNameWithoutExtension(path);
            var ext = Path.GetExtension(path);
            int i = 2;
            string candidate;
            do { candidate = Path.Combine(dir, $"{nameNoExt}_{i}{ext}"); i++; }
            while (File.Exists(candidate));
            return candidate;
        }

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

                if (c != null && !string.IsNullOrEmpty(c.Socccd))
                {
                    ProcessCCCD12Placeholder(mainPart, c.Socccd);
                }

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
                                    bool hasNtk4 = rowText.IndexOf("{{ntk4}}", StringComparison.OrdinalIgnoreCase) >= 0;

                                    // Nếu dòng có placeholder NTK nhưng không có dữ liệu, xóa các placeholder địa chỉ chỉ TRONG DÒNG NÀY
                                    if ((hasNtk1 && string.IsNullOrWhiteSpace(c.Ntk1)) ||
                                        (hasNtk2 && string.IsNullOrWhiteSpace(c.Ntk2)) ||
                                        (hasNtk3 && string.IsNullOrWhiteSpace(c.Ntk3)) ||
                                        (hasNtk4 && string.IsNullOrWhiteSpace(c.Ntk4)))
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
        // Làm việc theo từng paragraph để tránh vượt ranh giới ô bảng
        private void TryReplacePlaceholdersAcrossRuns(OpenXmlPart part, Dictionary<string, string> replacements)
        {
            if (part == null || replacements == null || replacements.Count == 0) return;
            try
            {
                // Lấy tất cả paragraph trong part (bao gồm cả trong bảng)
                var paragraphs = part.RootElement.Descendants<Paragraph>().ToList();

                foreach (var para in paragraphs)
                {
                    foreach (var kv in replacements)
                    {
                        var rawKey = kv.Key ?? string.Empty;
                        var replacement = kv.Value ?? string.Empty;
                        if (string.IsNullOrEmpty(rawKey)) continue;

                        string token = rawKey;
                        var m = Regex.Match(rawKey, "^\\s*\\{\\{\\s*(.*?)\\s*\\}\\}\\s*$");
                        if (m.Success && m.Groups.Count > 1) token = m.Groups[1].Value;
                        if (string.IsNullOrEmpty(token)) continue;

                        var pattern = "\\{\\{\\s*" + Regex.Escape(token) + "\\s*\\}\\}";

                        // Lặp đến khi không còn match trong paragraph này
                        bool found = true;
                        while (found)
                        {
                            found = false;
                            var texts = para.Descendants<Text>().ToList();
                            if (texts.Count == 0) break;

                            var combined = string.Concat(texts.Select(t => t.Text ?? string.Empty));
                            if (string.IsNullOrEmpty(combined)) break;

                            var match = Regex.Match(combined, pattern, RegexOptions.IgnoreCase);
                            if (!match.Success) break;

                            int pos = match.Index;
                            int matchLen = match.Length;

                            // Tìm startNode và startOffset
                            int remaining = pos;
                            int startNode = 0, startOffset = 0;
                            for (int k = 0; k < texts.Count; k++)
                            {
                                var tlen = (texts[k].Text ?? string.Empty).Length;
                                if (remaining <= tlen) { startNode = k; startOffset = remaining; break; }
                                remaining -= tlen;
                            }

                            // Tìm endNode và endOffset
                            remaining = pos + matchLen;
                            int endNode = 0, endOffset = 0;
                            for (int k = 0; k < texts.Count; k++)
                            {
                                var tlen = (texts[k].Text ?? string.Empty).Length;
                                if (remaining <= tlen) { endNode = k; endOffset = remaining; break; }
                                remaining -= tlen;
                            }

                            try
                            {
                                var prefix = (texts[startNode].Text ?? string.Empty).Substring(0, startOffset);
                                var suffix = (texts[endNode].Text ?? string.Empty).Substring(endOffset);

                                // Gán replacement vào startNode, xóa các node còn lại
                                texts[startNode].Text = prefix + replacement + suffix;
                                for (int k = startNode + 1; k <= endNode; k++)
                                    texts[k].Text = string.Empty;

                                found = true;
                            }
                            catch { break; }
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
        private void Form1_Load(object sender, EventArgs e)
        {
            // Auto-populate cbTinh và cbTinhfix từ file JSON trong thư mục
            try { TinhHelper.PopulateComboBox(cbTinh); } catch { }
            try { TinhHelper.PopulateComboBox(cbTinhfix); } catch { }

            // Đăng ký sự kiện cbTinhfix → populate cbpgdfix
            try { cbTinhfix.SelectedIndexChanged += CbTinhfix_SelectedIndexChanged; } catch { }
            // Đăng ký sự kiện cbpgdfix → load dữ liệu PGD lên dgv1
            try { cbpgdfix.SelectedIndexChanged += CbpgdfixEditor_SelectedIndexChanged; } catch { }
            // Đăng ký sự kiện cbpgdfix → populate cbXafix
            try { cbpgdfix.SelectedIndexChanged += CbXafix_PopulateFromPgd; } catch { }
            // Đăng ký sự kiện cbXafix → lọc dgv1 theo xã được chọn
            try { cbXafix.SelectedIndexChanged += CbXafix_SelectedIndexChanged; } catch { }
        }

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
            if (!ValidateRequiredFields()) return;
            if (!ValidateDuplicateCccdSdt()) return;
            try
            {
                var customer = ReadForm();
                 string existingFile = null;
                 if (customers != null && editingIndex >= 0 && editingIndex < customers.Count)
                 {
                     try { existingFile = customers[editingIndex]._fileName; } catch { existingFile = null; }
                 }

                 if (!string.IsNullOrEmpty(existingFile)) customer._fileName = existingFile;
                 else customer._fileName = customer._fileName; // leave as-is (null or set)

                 SaveCustomerToFile(customer);
                  // Thêm vào list ngay sau khi lưu file, trước await, để kiểm tra trùng vẫn hoạt động nếu template lỗi
                  UpsertCustomerInList(customer);

                  // btn01 should export only template 01 (no 03, no GUQ)
                  var createdFiles = await Task.Run(() => CreateProfileFromTemplate(customer, include03: false));

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

                 // Mở file vừa tạo
                 OpenCreatedFiles(createdFiles);
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
            if (!ValidateRequiredFields()) return;
            if (!ValidateDuplicateCccdSdt()) return;
            try
            {
                var customer = ReadForm();
                 string existingFile = null;
                 if (customers != null && editingIndex >= 0 && editingIndex < customers.Count)
                 {
                     try { existingFile = customers[editingIndex]._fileName; } catch { existingFile = null; }
                 }

                 if (!string.IsNullOrEmpty(existingFile)) customer._fileName = existingFile;

                 SaveCustomerToFile(customer);
                 UpsertCustomerInList(customer);

                 var createdFile = await Task.Run(() =>
                 {
                     // Btn03 should export only 03 DS template
                     return ExportSpecificTemplate(customer, "03 DS.docx");
                 });

                 BindGrid();
                 ClearForm();
                 MessageBox.Show("Export (03 DS) thành công.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);

                 // Mở file vừa tạo
                 if (!string.IsNullOrEmpty(createdFile))
                     OpenFile(createdFile);
             }
             catch (Exception ex)
             {
                 MessageBox.Show("Lỗi khi export 03: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         private async void BtnGUQ_Click(object sender, EventArgs e)
         {
            if (!ValidateRequiredFields()) return;
            if (!ValidateDuplicateCccdSdt()) return;
            try
            {
                var customer = ReadForm();
                 string existingFile = null;
                 if (customers != null && editingIndex >= 0 && editingIndex < customers.Count)
                 {
                     try { existingFile = customers[editingIndex]._fileName; } catch { existingFile = null; }
                 }

                 if (!string.IsNullOrEmpty(existingFile)) customer._fileName = existingFile;

                 SaveCustomerToFile(customer);
                 UpsertCustomerInList(customer);

                 var createdFile = await Task.Run(() =>
                 {
                     // BtnGUQ should export only GUQ template
                     return ExportSpecificTemplate(customer, "GUQ.docx");
                 });

                 BindGrid();
                 ClearForm();
                 MessageBox.Show("Export GUQ thành công.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);

                 // Mở file vừa tạo
                 if (!string.IsNullOrEmpty(createdFile))
                     OpenFile(createdFile);
             }
             catch (Exception ex)
             {
                 MessageBox.Show("Lỗi khi export GUQ: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         private async void Btn01tgtv_Click(object sender, EventArgs e)
         {
            if (!ValidateRequiredFields()) return;
            if (!ValidateDuplicateCccdSdt()) return;
            try
            {
                var customer = ReadForm();
                 string existingFile = null;
                 if (customers != null && editingIndex >= 0 && editingIndex < customers.Count)
                 {
                     try { existingFile = customers[editingIndex]._fileName; } catch { existingFile = null; }
                 }

                 if (!string.IsNullOrEmpty(existingFile)) customer._fileName = existingFile;

                 SaveCustomerToFile(customer);
                 UpsertCustomerInList(customer);

                 var createdFile = await Task.Run(() =>
                 {
                     // Btn01tgtv exports 01TGTV template
                     return ExportSpecificTemplate(customer, "01TGTV.docx");
                 });

                 BindGrid();
                 ClearForm();
                 MessageBox.Show("✅ Xuất mẫu 01TGTV thành công.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

                 // Mở file vừa tạo
                 if (!string.IsNullOrEmpty(createdFile))
                     OpenFile(createdFile);
             }
             catch (Exception ex)
             {
                 MessageBox.Show("❌ Lỗi khi xuất mẫu 01TGTV: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         private async void BtnBia_Click(object sender, EventArgs e)
         {
            if (!ValidateRequiredFields()) return;
            if (!ValidateDuplicateCccdSdt()) return;
            try
            {
                var customer = ReadForm();
                 string existingFile = null;
                 if (customers != null && editingIndex >= 0 && editingIndex < customers.Count)
                 {
                     try { existingFile = customers[editingIndex]._fileName; } catch { existingFile = null; }
                 }

                 if (!string.IsNullOrEmpty(existingFile)) customer._fileName = existingFile;

                 SaveCustomerToFile(customer);
                 UpsertCustomerInList(customer);

                 var createdFile = await Task.Run(() =>
                 {
                     return ExportSpecificTemplate(customer, "BIA.docx");
                 });

                 BindGrid();
                 ClearForm();
                 MessageBox.Show("✅ Xuất bìa hồ sơ thành công.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

                 if (!string.IsNullOrEmpty(createdFile))
                     OpenFile(createdFile);
             }
             catch (Exception ex)
             {
                 MessageBox.Show("❌ Lỗi khi xuất bìa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }

         private async void BtnDn_Click(object sender, EventArgs e)
         {
            if (!ValidateRequiredFields()) return;
            if (!ValidateDuplicateCccdSdt()) return;
            try
            {
                var customer = ReadForm();
                string existingFile = null;
                if (customers != null && editingIndex >= 0 && editingIndex < customers.Count)
                {
                    try { existingFile = customers[editingIndex]._fileName; } catch { existingFile = null; }
                }

                if (!string.IsNullOrEmpty(existingFile)) customer._fileName = existingFile;

                SaveCustomerToFile(customer);
                UpsertCustomerInList(customer);

                var createdFile = await Task.Run(() =>
                {
                    return ExportSpecificTemplate(customer, "GNCK.docx");
                });

                BindGrid();
                ClearForm();
                MessageBox.Show("✅ Xuất giấy đề nghị giải ngân thành công.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (!string.IsNullOrEmpty(createdFile))
                    OpenFile(createdFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show("❌ Lỗi khi xuất giấy nhận cam kết: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
             var r = MessageBox.Show($"Bạn có muốn xóa khách hàng \"{c.Hoten}\" không?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
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
             string ntk1 = "", ntk2 = "", ntk3 = "", ntk4 = "";
             string cccdntk1 = "", cccdntk2 = "", cccdntk3 = "", cccdntk4 = "";
             string namsinh1 = "", namsinh2 = "", namsinh3 = "", namsinh4 = "";
             string qh1 = "", qh2 = "", qh3 = "", qh4 = "";

             try { if (txtntk1 != null) ntk1 = txtntk1.Text.Trim(); } catch { }
             try { if (txtntk2 != null) ntk2 = txtntk2.Text.Trim(); } catch { }
             try { if (txtntk3 != null) ntk3 = txtntk3.Text.Trim(); } catch { }
             try { if (txtntk4 != null) ntk4 = txtntk4.Text.Trim(); } catch { }

             try { if (txtcccd1 != null) cccdntk1 = txtcccd1.Text.Trim(); } catch { }
             try { if (txtcccd2 != null) cccdntk2 = txtcccd2.Text.Trim(); } catch { }
             try { if (txtcccd3 != null) cccdntk3 = txtcccd3.Text.Trim(); } catch { }
             try { if (txtcccd4 != null) cccdntk4 = txtcccd4.Text.Trim(); } catch { }

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
             try 
             { 
                 if (datentk4 != null) 
                 {
                     if (datentk4.Checked)
                         namsinh4 = datentk4.Value.ToString("dd/MM/yyyy");
                     else
                         namsinh4 = "";
                 }
             } catch { }

             try { if (cbqh1 != null) qh1 = cbqh1.Text.Trim(); } catch { }
             try { if (cbqh2 != null) qh2 = cbqh2.Text.Trim(); } catch { }
             try { if (cbqh3 != null) qh3 = cbqh3.Text.Trim(); } catch { }
             try { if (cbqh4 != null) qh4 = cbqh4.Text.Trim(); } catch { }

             // Xác thực các ngày không được trong tương lai
             DateTime ngaycap = dateNgaycapCCCD.Value.Date;
             if (ngaycap > DateTime.Today) ngaycap = DateTime.Today;

             DateTime ngaysinh = DateTime.MinValue;
             if (dateNgaysinh != null)
             {
                 ngaysinh = dateNgaysinh.Value.Date;
                 if (ngaysinh > DateTime.Today) ngaysinh = DateTime.Today;
             }

             DateTime ngaylaphs = (dateLaphs != null && dateLaphs.Checked)
                 ? dateLaphs.Value.Date
                 : DateTime.MinValue;

             DateTime ngaydenhan = DateTime.MinValue;
             DateTime ngaygiaingaan = DateTime.MinValue;
             if (dateGn != null && dateGn.Checked)
             {
                 ngaygiaingaan = dateGn.Value.Date;
                 // Ngày đến hạn lấy từ dateDH (đã được tính tự động)
                 if (dateDH != null)
                     ngaydenhan = dateDH.Value.Date;
             }

             DateTime thoihancccd = DateTime.MinValue;
             string thoihancccdText = "";
             if (ngaysinh != DateTime.MinValue)
             {
                 thoihancccd = CalcThoiHanCCCD(ngaysinh, ngaycap);
                 thoihancccdText = thoihancccd == DateTime.MinValue ? "không thời hạn" : thoihancccd.ToString("dd/MM/yyyy");
             }
             else if (datendhcccd != null && datendhcccd.CustomFormat != " ")
             {
                 thoihancccd = datendhcccd.Value.Date;
                 thoihancccdText = thoihancccd.ToString("dd/MM/yyyy");
             }
             // VALIDATION: Không cho phép tạo hồ sơ nếu CCCD đã hết hạn (chỉ khi có ngày, không áp dụng khi không thời hạn)
             if (thoihancccd != DateTime.MinValue && thoihancccd < DateTime.Today)
             {
                 throw new Exception($"CCCD đã hết hạn ngày {thoihancccd:dd/MM/yyyy}.\n\nKhông thể tạo hồ sơ với CCCD hết hạn.\n\nVui lòng cập nhật CCCD mới.");
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
                 Tinh = (cbTinh != null ? cbTinh.Text : ""),
                 PGD = cbPGD.Text,
                 Chuongtrinh = cbChuongtrinh.Text,
                 Vtc = (cbVtc != null ? cbVtc.Text : ""),
                 Phuongan = (cbPhuongan != null ? cbPhuongan.Text : ""),
                 Thoihanvay = cbThoihanvay.Text,
                 Phanky = (cbPhanky != null ? cbPhanky.Text : ""),
                 Sotien = cbSotien.Text,
                 Sotien1 = cbSotien1.Text,
                 Sotien2 = cbSotien2.Text,
                 Sotien3 = (cbSotien3 != null ? cbSotien3.Text : ""),
                 Soluong1 = (cbDoituong1 != null ? cbDoituong1.Text : ""),
                 Soluong2 = (cbDoituong2 != null ? cbDoituong2.Text : ""),
                 Soluong3 = (cbDoituong3 != null ? cbDoituong3.Text : ""),
                 Sotientong = "",
                 Sotienchu = "",
                 Mucdich1 = (cbmucdich1 != null ? cbmucdich1.Text : ""),
                 Mucdich2 = (cbmucdich2 != null ? cbmucdich2.Text : ""),
                 Mucdich3 = (cbmucdich3 != null ? cbmucdich3.Text : ""),
                 Doituong1 = (cbDoituong != null ? cbDoituong.Text : ""),
                 Doituong2 = "",
                 Ngaylaphs = ngaylaphs,
                 Ngaydenhan = ngaydenhan,
                 Ngaygiaingaan = ngaygiaingaan,
                 Thoihancccd = thoihancccd,
                 ThoihancccdText = thoihancccdText,
                 Dantoc = (cbDantoc != null ? cbDantoc.Text : ""),
                 Sdt = (txtSdt != null ? txtSdt.Text : ""),  // Lưu với format có dấu chấm (0812.801.886)
                 Nhankhau = (txtNhankhau != null ? txtNhankhau.Text.Trim() : ""),
                 Ntk1 = ToTitleCase(ntk1), Ntk2 = ToTitleCase(ntk2), Ntk3 = ToTitleCase(ntk3), Ntk4 = ToTitleCase(ntk4),
                 CccdNtk1 = cccdntk1, CccdNtk2 = cccdntk2, CccdNtk3 = cccdntk3, CccdNtk4 = cccdntk4,
                 Namsinh1 = namsinh1, Namsinh2 = namsinh2, Namsinh3 = namsinh3, Namsinh4 = namsinh4,
                 Qh1 = qh1, Qh2 = qh2, Qh3 = qh3, Qh4 = qh4
             };
         }

         private void PopulateForm(Customer c)
         {
             if (c == null) return;

             // Thông tin cơ bản
             txtHoten.Text = c.Hoten ?? "";
             txtSocccd.Text = c.Socccd ?? "";
             cbNhandang.Text = c.Nhandang ?? "";

             // dateNgaycapCCCD
             try
             {
                 var ngaycap = c.Ngaycap == DateTime.MinValue ? DateTime.Today : c.Ngaycap;
                 if (ngaycap > DateTime.Today) ngaycap = DateTime.Today;
                 dateNgaycapCCCD.Format = DateTimePickerFormat.Custom;
                 dateNgaycapCCCD.CustomFormat = "dd/MM/yyyy";
                 dateNgaycapCCCD.Value = ngaycap;
             } catch { }

             // dateNgaysinh
             try 
             { 
                 if (dateNgaysinh != null) 
                 {
                     var ngaysinh = c.Ngaysinh == DateTime.MinValue ? DateTime.Today : c.Ngaysinh;
                     if (ngaysinh > DateTime.Today) ngaysinh = DateTime.Today;
                     dateNgaysinh.Format = DateTimePickerFormat.Custom;
                     dateNgaysinh.CustomFormat = "dd/MM/yyyy";
                     dateNgaysinh.Value = ngaysinh;
                 }
             } catch { }
             cbNoicap.Text = c.Noicap ?? "";

             // Thông tin cá nhân bổ sung
             try { if (cbGioitinh != null) cbGioitinh.Text = c.GioiTinh ?? ""; } catch { }
             try { if (cbDantoc != null) cbDantoc.Text = c.Dantoc ?? ""; } catch { }
             try { if (txtSdt != null) txtSdt.Text = c.Sdt ?? ""; } catch { }
             try { if (txtNhankhau != null) txtNhankhau.Text = c.Nhankhau ?? ""; } catch { }

             // Thông tin vị trí (đã bật suppress để tránh các sự kiện cascading)
             suppressComboChanged = true;
             try
             {
                 // Bước 1: set cbTinh, populate cbPGD theo tỉnh
                 var tinh = !string.IsNullOrEmpty(c.Tinh) ? c.Tinh : GetTinhFromPGD(c.PGD ?? "");
                 if (cbTinh != null)
                 {
                     cbTinh.Text = tinh;
                     cbPGD.Items.Clear();
                     currentTinhModel = null;
                     if (!string.IsNullOrWhiteSpace(tinh))
                     {
                         currentTinhModel = TinhHelper.LoadTinhModel(tinh);
                         if (currentTinhModel?.pgds != null)
                             foreach (var p in currentTinhModel.pgds)
                                 if (!string.IsNullOrWhiteSpace(p.pgd))
                                     cbPGD.Items.Add(p.pgd);
                     }
                 }

                 // Bước 2: set cbPGD
                 cbPGD.Text = c.PGD ?? "";

                 // Bước 3: load data cho cbXa/cbHoi/cbThon/cbTo
                 if (currentTinhModel?.pgds != null)
                 {
                     var pgdEntry = currentTinhModel.pgds.FirstOrDefault(p =>
                         string.Equals(p.pgd, c.PGD, StringComparison.OrdinalIgnoreCase));
                     xinmanModel = pgdEntry != null
                         ? new XinManModel { pgd = pgdEntry.pgd, communes = pgdEntry.communes ?? new List<Commune>() }
                         : null;
                 }
                 else
                 {
                     string jsonFileName = GetJsonFileNameFromPGD(c.PGD ?? "");
                     LoadXinManData(jsonFileName);
                 }

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
             try { if (cbPhuongan != null) cbPhuongan.Text = c.Phuongan ?? ""; } catch { }
             try { ApplyPhuonganState(c.Phuongan ?? ""); } catch { }
             try { ApplyCapNuocSach_Phuongan(c.Phuongan ?? ""); } catch { }
             cbThoihanvay.Text = c.Thoihanvay ?? "";
             try { if (cbPhanky != null) cbPhanky.Text = c.Phanky ?? ""; } catch { }

             // Thông tin số tiền
             cbSotien.Text = c.Sotien ?? "";
             cbSotien1.Text = c.Sotien1 ?? "";
             cbSotien2.Text = c.Sotien2 ?? "";

             // Mục đích và Đối tượng
             // Ghi đè sau ApplyPhuonganState/ApplyCapNuocSach_Phuongan để khôi phục giá trị đã lưu
             try { if (cbmucdich1 != null) { cbmucdich1.DropDownStyle = ComboBoxStyle.DropDown; cbmucdich1.Text = c.Mucdich1 ?? ""; } } catch { }
             try { if (cbmucdich2 != null) { cbmucdich2.DropDownStyle = ComboBoxStyle.DropDown; cbmucdich2.Text = c.Mucdich2 ?? ""; } } catch { }
             try { if (cbDoituong != null) cbDoituong.Text = c.Doituong1 ?? ""; } catch { }
             try { if (cbDoituong1 != null) cbDoituong1.Text = c.Soluong1 ?? ""; } catch { }
             try { if (cbDoituong2 != null) cbDoituong2.Text = c.Soluong2 ?? ""; } catch { }

             // dateLaphs
             try
             {
                 if (c.Ngaylaphs != DateTime.MinValue)
                 {
                     dateLaphs.Checked = true;
                     dateLaphs.Format = DateTimePickerFormat.Custom;
                     dateLaphs.CustomFormat = "dd/MM/yyyy";
                     dateLaphs.Value = c.Ngaylaphs;
                 }
                 else
                 {
                     dateLaphs.Checked = false;
                     dateLaphs.Format = DateTimePickerFormat.Custom;
                     dateLaphs.CustomFormat = " ";
                 }
             } catch { }
             try 
             { 
                 // dateDH: khoá hoàn toàn, chỉ hiển thị giá trị tính toán
                 if (dateDH != null)
                 {
                     dateDH.ShowCheckBox = false;
                     dateDH.Enabled = false;
                     dateDH.Format = DateTimePickerFormat.Custom;
                 }
                 // dateGn: checkbox ở đây - load trạng thái từ dữ liệu khách
                 if (dateGn != null)
                 {
                     if (c.Ngaygiaingaan == DateTime.MinValue)
                     {
                         dateGn.Checked = false;
                         dateDH.CustomFormat = " "; // ẩn ngày đến hạn
                     }
                     else
                     {
                         dateGn.Checked = true;
                         dateGn.Format = DateTimePickerFormat.Custom;
                         dateGn.CustomFormat = "dd/MM/yyyy";
                         dateGn.Value = c.Ngaygiaingaan;
                         if (c.Ngaydenhan != DateTime.MinValue)
                         {
                             dateDH.CustomFormat = "dd/MM/yyyy";
                             dateDH.Value = c.Ngaydenhan;
                         }
                     }
                 }
             } catch { }
             try 
             { 
                 if (datendhcccd != null) 
                 {
                     datendhcccd.Format = DateTimePickerFormat.Custom;
                     if (c.Thoihancccd == DateTime.MinValue && string.IsNullOrEmpty(c.ThoihancccdText))
                     {
                         datendhcccd.CustomFormat = " "; // trống nếu chưa có dữ liệu
                         datendhcccd.Enabled = false;
                     }
                     else if (c.ThoihancccdText == "không thời hạn")
                     {
                         // Trên 60 tuổi: hiển thị text không thời hạn
                         datendhcccd.CustomFormat = "'không thời hạn'";
                         if (c.Thoihancccd != DateTime.MinValue) datendhcccd.Value = c.Thoihancccd;
                         datendhcccd.Enabled = false;
                     }
                     else
                     {
                         datendhcccd.CustomFormat = "dd/MM/yyyy";
                         if (c.Thoihancccd != DateTime.MinValue) datendhcccd.Value = c.Thoihancccd;
                         datendhcccd.Enabled = false;
                     }
                 }
             } catch { }

             // NTK (Người thừa kế) info
             try { if (txtntk1 != null) txtntk1.Text = c.Ntk1 ?? ""; } catch { }
             try { if (txtntk2 != null) txtntk2.Text = c.Ntk2 ?? ""; } catch { }
             try { if (txtntk3 != null) txtntk3.Text = c.Ntk3 ?? ""; } catch { }
             try { if (txtntk4 != null) txtntk4.Text = c.Ntk4 ?? ""; } catch { }

             try { if (txtcccd1 != null) txtcccd1.Text = c.CccdNtk1 ?? ""; } catch { }
             try { if (txtcccd2 != null) txtcccd2.Text = c.CccdNtk2 ?? ""; } catch { }
             try { if (txtcccd3 != null) txtcccd3.Text = c.CccdNtk3 ?? ""; } catch { }
             try { if (txtcccd4 != null) txtcccd4.Text = c.CccdNtk4 ?? ""; } catch { }

             // ========== NGÀY SINH NGƯỜI THỪA KẾ (NTK) - DATEPICKER ==========
             try 
             { 
                 if (datentk1 != null) 
                 {
                     datentk1.Format = DateTimePickerFormat.Custom;
                     datentk1.ShowCheckBox = true;
                     if (string.IsNullOrWhiteSpace(c.Namsinh1))
                     {
                         datentk1.Checked = false;
                         datentk1.CustomFormat = " ";
                     }
                     else
                     {
                         DateTime parsedDate = ParseDateTextOrFallback(c.Namsinh1);
                         if (parsedDate != DateTime.MinValue)
                         {
                             datentk1.CustomFormat = "dd/MM/yyyy";
                             datentk1.Checked = true;
                             datentk1.Value = parsedDate;
                         }
                         else
                         {
                             datentk1.Checked = false;
                             datentk1.CustomFormat = " ";
                         }
                     }
                 }
             } catch { }
             try 
             { 
                 if (datentk2 != null) 
                 {
                     datentk2.Format = DateTimePickerFormat.Custom;
                     datentk2.ShowCheckBox = true;
                     if (string.IsNullOrWhiteSpace(c.Namsinh2))
                     {
                         datentk2.Checked = false;
                         datentk2.CustomFormat = " ";
                     }
                     else
                     {
                         DateTime parsedDate = ParseDateTextOrFallback(c.Namsinh2);
                         if (parsedDate != DateTime.MinValue)
                         {
                             datentk2.CustomFormat = "dd/MM/yyyy";
                             datentk2.Checked = true;
                             datentk2.Value = parsedDate;
                         }
                         else
                         {
                             datentk2.Checked = false;
                             datentk2.CustomFormat = " ";
                         }
                     }
                 }
             } catch { }
             try 
             { 
                 if (datentk3 != null) 
                 {
                     datentk3.Format = DateTimePickerFormat.Custom;
                     datentk3.ShowCheckBox = true;
                     if (string.IsNullOrWhiteSpace(c.Namsinh3))
                     {
                         datentk3.Checked = false;
                         datentk3.CustomFormat = " ";
                     }
                     else
                     {
                         DateTime parsedDate = ParseDateTextOrFallback(c.Namsinh3);
                         if (parsedDate != DateTime.MinValue)
                         {
                             datentk3.CustomFormat = "dd/MM/yyyy";
                             datentk3.Checked = true;
                             datentk3.Value = parsedDate;
                         }
                         else
                         {
                             datentk3.Checked = false;
                             datentk3.CustomFormat = " ";
                         }
                     }
                 }
             } catch { }
             try 
             { 
                 if (datentk4 != null) 
                 {
                     datentk4.Format = DateTimePickerFormat.Custom;
                     datentk4.ShowCheckBox = true;
                     if (string.IsNullOrWhiteSpace(c.Namsinh4))
                     {
                         datentk4.Checked = false;
                         datentk4.CustomFormat = " ";
                     }
                     else
                     {
                         DateTime parsedDate = ParseDateTextOrFallback(c.Namsinh4);
                         if (parsedDate != DateTime.MinValue)
                         {
                             datentk4.CustomFormat = "dd/MM/yyyy";
                             datentk4.Checked = true;
                             datentk4.Value = parsedDate;
                         }
                         else
                         {
                             datentk4.Checked = false;
                             datentk4.CustomFormat = " ";
                         }
                     }
                 }
             } catch { }

             try { if (cbqh1 != null) cbqh1.Text = c.Qh1 ?? ""; } catch { }
             try { if (cbqh2 != null) cbqh2.Text = c.Qh2 ?? ""; } catch { }
             try { if (cbqh3 != null) cbqh3.Text = c.Qh3 ?? ""; } catch { }
             try { if (cbqh4 != null) cbqh4.Text = c.Qh4 ?? ""; } catch { }
         }

         private void ClearForm()
         {
             // ── Text / ComboBox ──
             try { txtHoten.Clear(); } catch { }
             try { txtSocccd.Text = ""; } catch { }
             try { cbNhandang.SelectedIndex = -1; } catch { }
             try { cbNoicap.SelectedIndex = -1; } catch { }
             try { cbPGD.SelectedIndex = -1; } catch { }
             try { cbXa.Items.Clear(); cbXa.Text = ""; } catch { }
             try { cbThon.Items.Clear(); cbThon.Text = ""; } catch { }
             try { cbHoi.Items.Clear(); cbHoi.Text = ""; } catch { }
             try { cbTo.Items.Clear(); cbTo.Text = ""; } catch { }
             try { cbChuongtrinh.Text = ""; } catch { }
             try { cbThoihanvay.SelectedIndex = -1; } catch { }
             try { cbSotien.Text = ""; cbSotien1.Text = ""; cbSotien2.Text = ""; } catch { }
             try { if (cbmucdich1 != null) cbmucdich1.Text = ""; } catch { }
             try { if (cbmucdich2 != null) cbmucdich2.Text = ""; } catch { }
             try { if (cbDoituong != null) { cbDoituong.Enabled = true; cbDoituong.SelectedIndex = -1; } } catch { }
             try { if (cbVtc != null) cbVtc.SelectedIndex = -1; } catch { }
             try { if (cbPhuongan != null) cbPhuongan.Text = ""; } catch { }
             try { if (cbPhanky != null) cbPhanky.SelectedIndex = -1; } catch { }
             try { if (cbGioitinh != null) cbGioitinh.SelectedIndex = -1; } catch { }
             try { if (cbDantoc != null) cbDantoc.SelectedIndex = -1; } catch { }
             try { if (cbDoituong1 != null) cbDoituong1.Text = ""; } catch { }
             try { if (cbDoituong2 != null) cbDoituong2.Text = ""; } catch { }
             ResetCapNuocSach();

             // ── TextBox phụ ──
             try { if (txtSdt != null) txtSdt.Text = ""; } catch { }
             try { if (txtNhankhau != null) txtNhankhau.Text = ""; } catch { }
             try { if (txtntk1 != null) txtntk1.Text = ""; } catch { }
             try { if (txtntk2 != null) txtntk2.Text = ""; } catch { }
             try { if (txtntk3 != null) txtntk3.Text = ""; } catch { }
             try { if (txtntk4 != null) txtntk4.Text = ""; } catch { }
             try { if (txtcccd1 != null) txtcccd1.Text = ""; } catch { }
             try { if (txtcccd2 != null) txtcccd2.Text = ""; } catch { }
             try { if (txtcccd3 != null) txtcccd3.Text = ""; } catch { }
             try { if (txtcccd4 != null) txtcccd4.Text = ""; } catch { }
             try { if (cbqh1 != null) cbqh1.Text = ""; } catch { }
             try { if (cbqh2 != null) cbqh2.Text = ""; } catch { }
             try { if (cbqh3 != null) cbqh3.Text = ""; } catch { }
             try { if (cbqh4 != null) cbqh4.Text = ""; } catch { }

             // ── DateTimePicker không checkbox: ẩn ngày bằng CustomFormat trống ──
             try { dateNgaycapCCCD.Format = DateTimePickerFormat.Custom; dateNgaycapCCCD.CustomFormat = " "; } catch { }
             try { if (dateNgaysinh != null) { dateNgaysinh.Format = DateTimePickerFormat.Custom; dateNgaysinh.CustomFormat = " "; } } catch { }
             try { dateLaphs.Checked = false; dateLaphs.Format = DateTimePickerFormat.Custom; dateLaphs.CustomFormat = " "; } catch { }
             try { if (datendhcccd != null) { datendhcccd.Format = DateTimePickerFormat.Custom; datendhcccd.CustomFormat = " "; datendhcccd.Enabled = false; } } catch { }

             // ── DateTimePicker có checkbox: bỏ tick và ẩn ngày ──
             try { if (dateGn != null) dateGn.Checked = false; } catch { }
             try { if (datentk1 != null) { datentk1.Checked = false; datentk1.CustomFormat = " "; } } catch { }
             try { if (datentk2 != null) { datentk2.Checked = false; datentk2.CustomFormat = " "; } } catch { }
             try { if (datentk3 != null) { datentk3.Checked = false; datentk3.CustomFormat = " "; } } catch { }
             try { if (datentk4 != null) { datentk4.Checked = false; datentk4.CustomFormat = " "; } } catch { }

             // ── dateDH khoá: ẩn ngày ──
             try { if (dateDH != null) dateDH.CustomFormat = " "; } catch { }

             // ── Reset trạng thái khác ──
             editingIndex = -1;
             try { ResetVisibilityToDefault(); } catch { }
         }

         private void LoadXinManData(string jsonFileName = "xinman.json")
         {
             xinmanModel = null;

             var candidate = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, jsonFileName);
             if (File.Exists(candidate))
                 try { var json = File.ReadAllText(candidate, Encoding.UTF8); xinmanModel = TryDeserializeXinman(json); }
                 catch { xinmanModel = null; }

             if (xinmanModel == null)
             {
                 try
                 {
                     var asm = Assembly.GetExecutingAssembly();
                     var resName = asm.GetManifestResourceNames().FirstOrDefault(n => n.EndsWith(jsonFileName, StringComparison.OrdinalIgnoreCase));
                     if (resName != null) using (var s = asm.GetManifestResourceStream(resName)) using (var sr = new StreamReader(s, Encoding.UTF8)) { var json = sr.ReadToEnd(); xinmanModel = TryDeserializeXinman(json); }
                 }
                 catch { xinmanModel = null; }
             }
         }

         // Lấy tên file JSON từ tên PGD
         private string GetJsonFileNameFromPGD(string pgdName)
         {
             if (string.IsNullOrWhiteSpace(pgdName)) return "xinman.json";

             // Normalize chuỗi để loại bỏ dấu (so sánh dễ hơn)
             string normalized = RemoveVietnameseTones(pgdName.Trim().ToLowerInvariant());

             if (normalized.Contains("meo vac") || normalized.Contains("meovac"))
                 return "meovac.json";
             else if (normalized.Contains("vi xuyen") || normalized.Contains("vixuyen"))
                 return "vixuyen.json";
             else if (normalized.Contains("dong van") || normalized.Contains("dongvan"))
                 return "dongvan.json";
             else if (normalized.Contains("hoang su phi") || normalized.Contains("hoangSuphi") || normalized.Contains("hsp"))
                 return "hsp.json";
             else if (normalized.Contains("bac quang") || normalized.Contains("bacquang"))
                 return "bacquang.json";
             else if (normalized.Contains("xin man") || normalized.Contains("xinman"))
                 return "xinman.json";

             // Fallback: thử khớp với tên có dấu
             if (pgdName.IndexOf("Mèo Vạc", StringComparison.OrdinalIgnoreCase) >= 0)
                 return "meovac.json";
             else if (pgdName.IndexOf("Vị Xuyên", StringComparison.OrdinalIgnoreCase) >= 0)
                 return "vixuyen.json";
             else if (pgdName.IndexOf("Đồng Văn", StringComparison.OrdinalIgnoreCase) >= 0)
                 return "dongvan.json";
             else if (pgdName.IndexOf("Hoàng Su Phì", StringComparison.OrdinalIgnoreCase) >= 0)
                 return "hsp.json";
             else if (pgdName.IndexOf("Bắc Quang", StringComparison.OrdinalIgnoreCase) >= 0)
                 return "bacquang.json";
             else if (pgdName.IndexOf("Xín Mần", StringComparison.OrdinalIgnoreCase) >= 0)
                 return "xinman.json";

             return "xinman.json"; // default
         }

         // Loại bỏ dấu tiếng Việt để dễ so sánh
         private string RemoveVietnameseTones(string text)
         {
             if (string.IsNullOrWhiteSpace(text)) return text;

             try
             {
                 string[] vietnameseSigns = new string[]
                 {
                     "aAeEoOuUiIdDyY",
                     "áàạảãâấầậẩẫăắằặẳẵ",
                     "ÁÀẠẢÃÂẤẦẬẨẪĂẮẰẶẲẴ",
                     "éèẹẻẽêếềệểễ",
                     "ÉÈẸẺẼÊẾỀỆỂỄ",
                     "óòọỏõôốồộổỗơớờợởỡ",
                     "ÓÒỌỎÕÔỐỒỘỔỖƠỚỜỢỞỠ",
                     "úùụủũưứừựửữ",
                     "ÚÙỤỦŨƯỨỪỰỬỮ",
                     "íìịỉĩ",
                     "ÍÌỊỈĨ",
                     "đ",
                     "Đ",
                     "ýỳỵỷỹ",
                     "ÝỲỴỶỸ"
                 };

                 for (int i = 1; i < vietnameseSigns.Length; i++)
                 {
                     for (int j = 0; j < vietnameseSigns[i].Length; j++)
                     {
                         text = text.Replace(vietnameseSigns[i][j], vietnameseSigns[0][i - 1]);
                     }
                 }

                 return text;
             }
             catch
             {
                 return text;
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

        private void CbTinh_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (suppressComboChanged) return;
            suppressComboChanged = true;
            try
            {
                cbPGD.Items.Clear();
                cbPGD.Text = "";
                ClearPgdDependentCombos();
                xinmanModel = null;
                currentTinhModel = null;

                var tinh = cbTinh?.Text ?? "";
                if (string.IsNullOrWhiteSpace(tinh)) return;

                currentTinhModel = TinhHelper.LoadTinhModel(tinh);
                if (currentTinhModel?.pgds != null)
                    foreach (var pgd in currentTinhModel.pgds)
                        if (!string.IsNullOrWhiteSpace(pgd.pgd))
                            cbPGD.Items.Add(pgd.pgd);
            }
            finally { suppressComboChanged = false; }
        }

        private string GetTinhJsonFileName(string tinhName)
        {
            return TinhHelper.GetFileName(tinhName);
        }

        private string[] GetPgdItemsForTinh(string tinhName)
        {
            return new string[0];
        }

        private string GetTinhFromPGD(string pgdName)
        {
            if (string.IsNullOrWhiteSpace(pgdName)) return "";
            // Tìm ngược từ TinhHelper cache — không hardcode tên tỉnh
            foreach (var tinhName in TinhHelper.GetProvinceNames())
            {
                var model = TinhHelper.LoadTinhModel(tinhName);
                if (model?.pgds == null) continue;
                if (model.pgds.Any(p => string.Equals(p.pgd, pgdName.Trim(), StringComparison.OrdinalIgnoreCase)))
                    return tinhName;
            }
            return "";
        }

        private void CbPGD_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (suppressComboChanged) return; suppressComboChanged = true;
            try
            {
                var selected = (cbPGD.SelectedItem ?? cbPGD.Text)?.ToString() ?? "";
                if (string.IsNullOrWhiteSpace(selected)) { ClearPgdDependentCombos(); return; }

                // Nếu đang dùng file JSON cấp tỉnh (ví dụ: tuyenquang.json), lấy communes từ đó
                if (currentTinhModel?.pgds != null)
                {
                    var pgdEntry = currentTinhModel.pgds.FirstOrDefault(p =>
                        string.Equals(p.pgd, selected, StringComparison.OrdinalIgnoreCase));
                    if (pgdEntry != null)
                    {
                        xinmanModel = new XinManModel { pgd = pgdEntry.pgd, communes = pgdEntry.communes ?? new List<Commune>() };
                    }
                    else
                    {
                        xinmanModel = null;
                        ClearPgdDependentCombos();
                        return;
                    }
                }
                else
                {
                    // Tỉnh Hà Giang: dùng file JSON riêng cho từng PGD
                    string jsonFileName = GetJsonFileNameFromPGD(selected);
                    System.Diagnostics.Debug.WriteLine($"PGD selected: '{selected}' -> Loading: '{jsonFileName}'");
                    LoadXinManData(jsonFileName);
                    if (xinmanModel == null)
                    {
                        ClearPgdDependentCombos();
                        return;
                    }
                }

                cbXa.Items.Clear(); cbThon.Items.Clear(); cbHoi.Items.Clear(); cbTo.Items.Clear();

                // Thêm các xã
                foreach (var com in xinmanModel.communes)
                    if (!string.IsNullOrWhiteSpace(com.name) && !cbXa.Items.Contains(com.name))
                        cbXa.Items.Add(com.name);

                // Thêm các hội — bỏ qua nan
                var associations = xinmanModel.communes.Where(c => c.associations != null).SelectMany(c => c.associations).Where(a => !string.IsNullOrWhiteSpace(a.name) && !a.name.Equals("nan", StringComparison.OrdinalIgnoreCase)).Select(a => a.name).Distinct(StringComparer.OrdinalIgnoreCase);
                foreach (var a in associations) cbHoi.Items.Add(a);

                // Thêm tất cả thôn (thuộc hội và trực thuộc xã)
                foreach (var com in xinmanModel.communes) { if (com.associations != null) foreach (var assoc in com.associations) if (assoc.villages != null) foreach (var v in assoc.villages) if (!string.IsNullOrWhiteSpace(v.name)) cbThon.Items.Add(v.name); if (com.villages != null) foreach (var v in com.villages) if (!string.IsNullOrWhiteSpace(v.name)) cbThon.Items.Add(v.name); }

                ResetVisibilityToDefault();
            }
            finally { suppressComboChanged = false; }
        }

        private void CbXa_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (suppressComboChanged) return; suppressComboChanged = true;
            try
            {
                cbHoi.Items.Clear(); cbThon.Items.Clear(); cbTo.Items.Clear();
                if (cbXa.SelectedIndex < 0 || xinmanModel == null) { ResetVisibilityToDefault(); return; }
                var xaName = cbXa.SelectedItem.ToString();
                var commune = xinmanModel.communes.FirstOrDefault(c => string.Equals(c.name, xaName, StringComparison.OrdinalIgnoreCase));
                if (commune != null)
                {
                    bool hasRealAssoc = false;
                    if (commune.associations != null)
                        foreach (var a in commune.associations)
                            if (!string.IsNullOrWhiteSpace(a.name) && !a.name.Equals("nan", StringComparison.OrdinalIgnoreCase))
                            { cbHoi.Items.Add(a.name); hasRealAssoc = true; }

                    if (!hasRealAssoc)
                    {
                        if (commune.associations != null)
                            foreach (var a in commune.associations)
                                if (a.villages != null)
                                    foreach (var v in a.villages)
                                        if (!string.IsNullOrWhiteSpace(v.name)) cbThon.Items.Add(v.name);
                        if (commune.villages != null)
                            foreach (var v in commune.villages)
                                if (!string.IsNullOrWhiteSpace(v.name)) cbThon.Items.Add(v.name);
                    }
                }
                ResetVisibilityToDefault();
            }
            finally { suppressComboChanged = false; }
        }

        private void CbThon_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (suppressComboChanged) return; suppressComboChanged = true;
            try { cbTo.Items.Clear(); if (cbThon.SelectedIndex < 0 || xinmanModel == null) { ResetVisibilityToDefault(); return; } var thonName = cbThon.SelectedItem.ToString(); var currentHoiName = cbHoi.Text ?? ""; var currentXaName = cbXa.Text ?? ""; Village foundVillage = null; foreach (var com in xinmanModel.communes) { if (!string.IsNullOrWhiteSpace(currentXaName) && !string.Equals(com.name, currentXaName, StringComparison.OrdinalIgnoreCase)) continue; if (com.associations != null) foreach (var assoc in com.associations) { if (!string.IsNullOrWhiteSpace(currentHoiName) && !string.Equals(assoc.name, currentHoiName, StringComparison.OrdinalIgnoreCase)) continue; if (assoc.villages != null) { foundVillage = assoc.villages.FirstOrDefault(v => string.Equals(v.name, thonName, StringComparison.OrdinalIgnoreCase)); if (foundVillage != null) break; } } if (foundVillage != null) break; if (string.IsNullOrWhiteSpace(currentHoiName) && com.villages != null) { foundVillage = com.villages.FirstOrDefault(v => string.Equals(v.name, thonName, StringComparison.OrdinalIgnoreCase)); if (foundVillage != null) break; } } if (foundVillage != null && foundVillage.groups != null) foreach (var g in foundVillage.groups) if (!string.IsNullOrWhiteSpace(g)) cbTo.Items.Add(g); ResetVisibilityToDefault(); }
            finally { suppressComboChanged = false; }
        }

        private void CbHoi_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (suppressComboChanged) return; suppressComboChanged = true;
            try
            {
                cbThon.Items.Clear(); cbTo.Items.Clear();
                if (cbHoi.SelectedIndex < 0 || xinmanModel == null) { ResetVisibilityToDefault(); return; }
                var hoiName = cbHoi.SelectedItem.ToString();
                // Nếu chọn nan thì bỏ qua, populate cbThon từ tất cả assoc nan của xã hiện tại
                if (hoiName.Equals("nan", StringComparison.OrdinalIgnoreCase))
                {
                    PopulateThonFromNanAssoc();
                    ResetVisibilityToDefault();
                    return;
                }
                var currentXaName = cbXa.Text ?? "";
                var assocList = xinmanModel.communes.Where(c => c.associations != null).Where(c => string.IsNullOrWhiteSpace(currentXaName) || string.Equals(c.name, currentXaName, StringComparison.OrdinalIgnoreCase)).SelectMany(c => c.associations.Select(a => new { Commune = c, Assoc = a })).Where(x => string.Equals(x.Assoc.name, hoiName, StringComparison.OrdinalIgnoreCase)).ToList();
                foreach (var pair in assocList) { var assoc = pair.Assoc; if (assoc.villages != null) foreach (var v in assoc.villages) { if (!string.IsNullOrWhiteSpace(v.name)) cbThon.Items.Add(v.name); } if (assoc.managedVillages != null) foreach (var name in assoc.managedVillages) { if (!string.IsNullOrWhiteSpace(name)) cbThon.Items.Add(name); } }
                if (string.IsNullOrWhiteSpace(currentXaName)) { var communesForAssoc = assocList.Select(x => x.Commune.name).Distinct(StringComparer.OrdinalIgnoreCase).ToList(); if (communesForAssoc.Count == 1) cbXa.SelectedItem = communesForAssoc[0]; }
                ResetVisibilityToDefault();
            }
            finally { suppressComboChanged = false; }
        }

        private void CbHoi_TextChanged(object sender, EventArgs e)
        {
            if (suppressComboChanged) return;
            if (xinmanModel == null) return;
            // cbHoi không có hội thật, user gõ ký tự bất kỳ → populate cbThon ngay
            if (cbHoi.Items.Count == 0 && !string.IsNullOrEmpty(cbHoi.Text))
            {
                suppressComboChanged = true;
                try { cbThon.Items.Clear(); cbTo.Items.Clear(); PopulateThonFromNanAssoc(); }
                finally { suppressComboChanged = false; }
            }
        }

        private void PopulateThonFromNanAssoc()
        {
            if (xinmanModel == null) return;
            cbThon.Items.Clear(); cbTo.Items.Clear();
            var currentXaName = cbXa.Text ?? "";
            foreach (var com in xinmanModel.communes)
            {
                if (!string.IsNullOrWhiteSpace(currentXaName) && !string.Equals(com.name, currentXaName, StringComparison.OrdinalIgnoreCase)) continue;
                if (com.associations != null)
                    foreach (var a in com.associations)
                        if (a.name != null && a.name.Equals("nan", StringComparison.OrdinalIgnoreCase) && a.villages != null)
                            foreach (var v in a.villages)
                                if (!string.IsNullOrWhiteSpace(v.name)) cbThon.Items.Add(v.name);
                if (com.villages != null)
                    foreach (var v in com.villages)
                        if (!string.IsNullOrWhiteSpace(v.name)) cbThon.Items.Add(v.name);
            }
        }

        // Tự động điền cbDoituong dựa trên cbChuongtrinh được chọn
        private void CbChuongtrinh_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cbDoituong == null || cbChuongtrinh == null) return;

                var chuongtrinh = (cbChuongtrinh.Text ?? "").Trim();
                if (string.IsNullOrWhiteSpace(chuongtrinh)) return;

                // Reset module Cấp nước sạch trước khi áp rule mới
                ResetCapNuocSach();

                // Chuẩn hóa để so sánh (loại bỏ dấu)
                string Normalize(string s)
                {
                    if (string.IsNullOrWhiteSpace(s)) return "";
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

                var normalized = Normalize(chuongtrinh);

                // ========== QUY TẮC TỰ ĐỘNG ĐIỀN ==========

                // 1. Hộ nghèo → Hộ nghèo
                if (normalized.Contains("ho nghe") && !normalized.Contains("can") && !normalized.Contains("moi thoat"))
                {
                    cbDoituong.DropDownStyle = ComboBoxStyle.DropDown;  // Reset về mặc định
                    cbDoituong.Enabled = false;
                    cbDoituong.Text = "Hộ nghèo";
                    return;
                }

                // 2. Hộ cận nghèo → Hộ cận nghèo
                if (normalized.Contains("ho can nghe"))
                {
                    cbDoituong.DropDownStyle = ComboBoxStyle.DropDown;  // Reset về mặc định
                    cbDoituong.Enabled = false;
                    cbDoituong.Text = "Hộ cận nghèo";
                    return;
                }

                // 3. Hộ mới thoát nghèo → Hộ mới thoát nghèo
                if (normalized.Contains("ho moi thoat nghe"))
                {
                    cbDoituong.DropDownStyle = ComboBoxStyle.DropDown;  // Reset về mặc định
                    cbDoituong.Enabled = false;
                    cbDoituong.Text = "Hộ mới thoát nghèo";
                    return;
                }

                // 4. Hộ gia đình Sản xuất kinh doanh tại vùng khó khăn → Hộ GĐ SXKD VKK
                if (normalized.Contains("sxkd") ||
                    (normalized.Contains("san xuat kinh doanh") && normalized.Contains("vung kho khan")) ||
                    (normalized.Contains("ho gia dinh") && normalized.Contains("san xuat kinh doanh") && normalized.Contains("vung kho khan")))
                {
                    cbDoituong.DropDownStyle = ComboBoxStyle.DropDown;  // Reset về mặc định
                    cbDoituong.Enabled = false;
                    cbDoituong.Text = "Hộ GĐ SXKD VKK";
                    return;
                }

                // 5. Cấp nước sạch và vệ sinh môi trường nông thôn → HGĐ cư trú tại VNT
                if (normalized.Contains("cap nuoc sach") ||
                    normalized.Contains("ve sinh moi truong") ||
                    (normalized.Contains("nuoc sach") && normalized.Contains("nong thon")))
                {
                    cbDoituong.DropDownStyle = ComboBoxStyle.DropDown;
                    cbDoituong.Enabled = false;
                    cbDoituong.Text = "HGĐ cư trú tại VNT";
                    ApplyCapNuocSach_Doituong();
                    return;
                }

                // 6. Hỗ trợ tạo việc làm duy trì và mở rộng việc làm → Người lao động hoặc NLĐ là người DTTS
                if (normalized.Contains("ho tro tao viec lam") ||
                    normalized.Contains("gqvl") ||
                    (normalized.Contains("duy tri") && normalized.Contains("mo rong") && normalized.Contains("viec lam")))
                {
                    cbDoituong.Enabled = true;  // MỞ để cho phép chọn
                    cbDoituong.DropDownStyle = ComboBoxStyle.DropDownList;  // KHOÁ không cho nhập text, chỉ cho chọn từ dropdown

                    // Clear list và chỉ thêm 2 option được phép
                    cbDoituong.Items.Clear();
                    cbDoituong.Items.Add("Người lao động");
                    cbDoituong.Items.Add("NLĐ là người DTTS");

                    // Nếu chưa chọn hoặc giá trị không hợp lệ, mặc định chọn "Người lao động"
                    if (string.IsNullOrWhiteSpace(cbDoituong.Text) || 
                        (!cbDoituong.Text.Equals("Người lao động") && !cbDoituong.Text.Equals("NLĐ là người DTTS")))
                        cbDoituong.Text = "Người lao động";
                    return;
                }

                // Mặc định: MỞ KHÓA nếu không match bất kỳ rule nào
                cbDoituong.DropDownStyle = ComboBoxStyle.DropDown;  // Reset về mặc định
                cbDoituong.Enabled = true;
            }
            catch { }
        }

        // Áp dụng trạng thái khoá/mở khoá cho cbmucdich1 và cbmucdich2 dựa theo cbPhuongan
        private void ApplyPhuonganState(string phuongan)
        {
            if (cbmucdich1 == null || cbmucdich2 == null || cbPhuongan == null) return;
            var pa = (phuongan ?? "").TrimEnd();
            bool isNangCap = string.Equals(pa, "Nâng cấp, sửa chữa CTNS, CTVS", StringComparison.OrdinalIgnoreCase);
            bool isXayMoi  = string.Equals(pa, "Xây mới CTNS, CTVS", StringComparison.OrdinalIgnoreCase);

            if (isNangCap || isXayMoi)
            {
                // cbmucdich1: tự động điền, khoá không cho người dùng tác động
                cbmucdich1.Items.Clear();
                cbmucdich1.Items.Add("Nâng cấp, sửa chữa CTNS          ");
                cbmucdich1.Items.Add("Xây mới CTNS                              ");
                cbmucdich1.DropDownStyle = ComboBoxStyle.DropDownList;
                cbmucdich1.Enabled = false;

                // cbmucdich2: tự động điền, cho phép nhập tay kèm danh sách
                cbmucdich2.Items.Clear();
                cbmucdich2.Items.Add("Nâng cấp, sửa chữa CTVS          ");
                cbmucdich2.Items.Add("Xây mới CTVS                                    ");
                cbmucdich2.DropDownStyle = ComboBoxStyle.DropDown;
                cbmucdich2.Enabled = true;
            }
            else if (!string.IsNullOrWhiteSpace(pa))
            {
                // Kiểm tra có phải item hợp lệ trong cbPhuongan không
                bool isKnownItem = cbPhuongan.Items.Cast<object>()
                    .Any(i => string.Equals((i ?? "").ToString().TrimEnd(), pa, StringComparison.OrdinalIgnoreCase));

                if (isKnownItem)
                {
                    // Item đã biết: khoá cbmucdich1, hiển thị giống cbPhuongan
                    cbmucdich1.DropDownStyle = ComboBoxStyle.DropDown;
                    cbmucdich1.Enabled = false;
                    cbmucdich1.Text = pa;

                    // cbmucdich2: cho phép nhập tự do
                    cbmucdich2.DropDownStyle = ComboBoxStyle.DropDown;
                    cbmucdich2.Enabled = true;
                    cbmucdich2.Text = "";
                }
                else
                {
                    // Nhập tay tự do: xoá trắng cbmucdich1/2, cho phép nhập tự do
                    cbmucdich1.Text = "";
                    cbmucdich2.Text = "";
                    cbmucdich1.DropDownStyle = ComboBoxStyle.DropDown;
                    cbmucdich1.Enabled = true;
                    cbmucdich2.DropDownStyle = ComboBoxStyle.DropDown;
                    cbmucdich2.Enabled = true;
                }
            }
            else
            {
                // cbPhuongan chưa chọn: khôi phục danh sách gốc, khoá dropdown (không cho nhập)
                cbmucdich1.Items.Clear();
                cbmucdich1.Items.Add("Mua trâu sinh sản");
                cbmucdich1.Items.Add("Nuôi trâu sinh sản");
                cbmucdich1.Items.Add("Mua bò sinh sản");
                cbmucdich1.Items.Add("Nuôi bò sinh sản");
                cbmucdich1.Items.Add("Mua dê sinh sản");
                cbmucdich1.Items.Add("Nuôi dê sinh sản");
                cbmucdich1.Items.Add("Nuôi lợn sinh sản");
                cbmucdich1.Items.Add("Nuôi lợn");
                cbmucdich1.Items.Add("Trồng cây quế");
                cbmucdich1.Items.Add("Trồng cây keo");
                cbmucdich1.Items.Add("Trồng cây mỡ");
                cbmucdich1.Items.Add("Trồng cây cam");
                cbmucdich1.Items.Add("Mở rộng cửa hàng tạp hoá");
                cbmucdich1.Items.Add("Mở rộng cửa hàng ăn uống");
                cbmucdich1.Items.Add("Mở rộng cửa hàng bán quần áo");
                cbmucdich1.Items.Add("Trồng và chăm sóc cây cà phê");
                cbmucdich1.Items.Add("Trồng và chăm sóc cây cao su");
                cbmucdich1.Items.Add("Trồng cây ăn quả");
                cbmucdich1.Items.Add("Trồng cây bời lời");
                cbmucdich1.Items.Add("Trồng cây tiêu");
                cbmucdich1.Items.Add("Nâng cấp, sửa chữa CTNS          ");
                cbmucdich1.Items.Add("Xây mới CTNS                              ");
                cbmucdich1.DropDownStyle = ComboBoxStyle.DropDownList;
                cbmucdich1.Enabled = true;
                cbmucdich1.Text = "";

                cbmucdich2.Items.Clear();
                cbmucdich2.Items.Add("Nâng cấp, sửa chữa CTVS          ");
                cbmucdich2.Items.Add("Xây mới CTVS                                    ");
                cbmucdich2.DropDownStyle = ComboBoxStyle.DropDown;
                cbmucdich2.Enabled = true;
                cbmucdich2.Text = "";
            }
        }

        private void CbPhuongan_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cbPhuongan == null) return;
                var phuongan = cbPhuongan.Text ?? "";
                ApplyPhuonganState(phuongan);

                // Tự động điền giá trị mặc định cho cbmucdich1 và cbmucdich2
                var pa = phuongan.TrimEnd();
                if (string.Equals(pa, "Nâng cấp, sửa chữa CTNS, CTVS", StringComparison.OrdinalIgnoreCase))
                {
                    cbmucdich1.Text = "Nâng cấp, sửa chữa CTNS          ";
                    cbmucdich2.Text = "Nâng cấp, sửa chữa CTVS          ";
                }
                else if (string.Equals(pa, "Xây mới CTNS, CTVS", StringComparison.OrdinalIgnoreCase))
                {
                    cbmucdich1.Text = "Xây mới CTNS                              ";
                    cbmucdich2.Text = "Xây mới CTVS                                    ";
                }
                // Module Cấp nước sạch: ghi đè cbmucdich1/2 với giá trị chính xác (không trailing spaces)
                ApplyCapNuocSach_Phuongan(phuongan);
            }
            catch { }
        }

        private void CbPhuongan_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (cbPhuongan == null || cbmucdich1 == null || cbmucdich2 == null) return;
                // Chỉ xử lý khi đang nhập tay (không phải chọn từ danh sách)
                if (cbPhuongan.SelectedIndex >= 0) return;
                ApplyPhuonganState(cbPhuongan.Text ?? "");
            }
            catch { }
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

                // Luôn kiểm tra tất cả khách hàng (không bỏ qua ai) khi tạo MỚI
                editingIndex = -1;
                if (!ValidateDuplicateCccdSdt()) return;

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

        // ComboBox Fix PGD - Chọn PGD và tự động reload dữ liệu từ file JSON tương ứng
        private void CbPgdFix_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cbpgdfix == null || cbpgdfix.SelectedItem == null) return;

                var selected = cbpgdfix.SelectedItem.ToString();
                if (string.IsNullOrWhiteSpace(selected)) return;

                // Nếu đang có _editTinhModel (chọn từ cbTinhfix), dùng trực tiếp không cần file
                if (_editTinhModel?.pgds != null)
                {
                    var pgdEntry = _editTinhModel.pgds.FirstOrDefault(p =>
                        string.Equals(p.pgd, selected, StringComparison.OrdinalIgnoreCase));

                    if (pgdEntry != null)
                    {
                        xinManEditor?.LoadFromPgdEntry(pgdEntry, _editTinhModel.tinh);
                    }
                    return;
                }

                // Fallback: thử load từ TinhHelper
                var tinhName = cbTinhfix?.Text?.Trim() ?? "";
                if (!string.IsNullOrWhiteSpace(tinhName))
                {
                    _editTinhModel = TinhHelper.LoadTinhModel(tinhName);
                    var pgdEntry = _editTinhModel?.pgds?.FirstOrDefault(p =>
                        string.Equals(p.pgd, selected, StringComparison.OrdinalIgnoreCase));
                    if (pgdEntry != null)
                    {
                        xinManEditor?.LoadFromPgdEntry(pgdEntry, tinhName);
                    }
                }
            }
            catch { }
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
        private bool IsBia(string docPath) { if (string.IsNullOrWhiteSpace(docPath)) return false; var name = Path.GetFileName(docPath) ?? ""; return name.IndexOf("BIA", StringComparison.OrdinalIgnoreCase) >= 0; }
        private bool Is01SXKD(string docPath) { if (string.IsNullOrWhiteSpace(docPath)) return false; var name = Path.GetFileName(docPath) ?? ""; return name.IndexOf("01", StringComparison.OrdinalIgnoreCase) >= 0 && name.IndexOf("SXKD", StringComparison.OrdinalIgnoreCase) >= 0; }
        private bool Is01GQVL(string docPath) { if (string.IsNullOrWhiteSpace(docPath)) return false; var name = Path.GetFileName(docPath) ?? ""; return name.IndexOf("GQVL", StringComparison.OrdinalIgnoreCase) >= 0; }
        private bool Is01HN(string docPath) { if (string.IsNullOrWhiteSpace(docPath)) return false; var name = Path.GetFileName(docPath) ?? ""; return name.IndexOf("01", StringComparison.OrdinalIgnoreCase) >= 0 && name.IndexOf("HN", StringComparison.OrdinalIgnoreCase) >= 0; }
        private bool IsGUQ(string docPath) { if (string.IsNullOrWhiteSpace(docPath)) return false; var name = Path.GetFileName(docPath) ?? ""; return name.IndexOf("GUQ", StringComparison.OrdinalIgnoreCase) >= 0; }
        private bool HasNguoiUyQuyen(Customer customer) { return !string.IsNullOrWhiteSpace(customer?.Ntk1) || !string.IsNullOrWhiteSpace(customer?.Ntk2) || !string.IsNullOrWhiteSpace(customer?.Ntk3) || !string.IsNullOrWhiteSpace(customer?.Ntk4); }

        /// <summary>
        /// Căn giữa paragraph chứa chuỗi <paramref name="marker"/> trong file docx.
        /// </summary>
        private void CenterParagraphContaining(string docPath, string marker)
        {
            if (string.IsNullOrWhiteSpace(docPath) || !File.Exists(docPath)) return;
            using (var wordDoc = WordprocessingDocument.Open(docPath, true))
            {
                var mainPart = wordDoc.MainDocumentPart;
                if (mainPart == null) return;

                var allParts = new List<OpenXmlPart> { mainPart };
                allParts.AddRange(mainPart.HeaderParts);
                allParts.AddRange(mainPart.FooterParts);

                foreach (var part in allParts)
                {
                    bool changed = false;
                    foreach (var para in part.RootElement.Descendants<Paragraph>())
                    {
                        var paraText = string.Concat(para.Descendants<Text>().Select(t => t.Text ?? ""));
                        if (paraText.IndexOf(marker, StringComparison.OrdinalIgnoreCase) < 0) continue;

                        var pPr = para.GetFirstChild<ParagraphProperties>();
                        if (pPr == null) { pPr = new ParagraphProperties(); para.InsertAt(pPr, 0); }
                        var jc = pPr.GetFirstChild<Justification>();
                        if (jc == null) { jc = new Justification(); pPr.Append(jc); }
                        jc.Val = JustificationValues.Center;
                        changed = true;
                    }
                    if (changed) part.RootElement.Save();
                }
                mainPart.Document.Save();
            }
        }

        // Giải quyết đường dẫn mẫu
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
        // KeyPress handler cho số điện thoại: chỉ cho phép số và dấu chấm
        private void TxtSdt_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Cho phép: số, backspace, dấu chấm
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            {
                e.Handled = true;
            }
        }

        // Tự động format số điện thoại với dấu chấm (0812.801.886)
        private void TxtSdt_TextChanged(object sender, EventArgs e)
        {
            try
            {
                var tb = sender as TextBox;
                if (tb == null) return;

                // Lưu vị trí con trỏ
                var cursorPosition = tb.SelectionStart;
                var originalText = tb.Text ?? "";

                // Chỉ giữ lại các chữ số
                var digits = new string(originalText.Where(char.IsDigit).ToArray());

                // Nếu không có thay đổi gì, thoát
                if (digits.Length == 0)
                {
                    if (tb.Text != "") tb.Text = "";
                    return;
                }

                // Giới hạn 10 chữ số
                if (digits.Length > 10)
                {
                    digits = digits.Substring(0, 10);
                }

                // Format: 0812.801.886 (4-3-3)
                string formatted = "";
                if (digits.Length <= 4)
                {
                    formatted = digits;
                }
                else if (digits.Length <= 7)
                {
                    formatted = digits.Substring(0, 4) + "." + digits.Substring(4);
                }
                else
                {
                    formatted = digits.Substring(0, 4) + "." + digits.Substring(4, 3) + "." + digits.Substring(7);
                }

                // Chỉ cập nhật nếu có thay đổi
                if (formatted != originalText)
                {
                    tb.Text = formatted;

                    // Điều chỉnh vị trí con trỏ
                    // Nếu đang gõ, giữ con trỏ ở cuối
                    if (cursorPosition >= originalText.Length)
                    {
                        tb.SelectionStart = formatted.Length;
                    }
                    else
                    {
                        // Nếu đang sửa giữa chuỗi, cố giữ vị trí tương đối
                        tb.SelectionStart = Math.Min(cursorPosition + (formatted.Length - originalText.Length), formatted.Length);
                    }
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

                // Lấy chỉ các chữ số (bỏ dấu chấm)
                var digits = new string(text.Where(char.IsDigit).ToArray());

                // Kiểm tra phải đúng 10 số - không ít hơn, không nhiều hơn
                if (digits.Length > 0 && digits.Length != 10)
                {
                    MessageBox.Show(
                        $"Số điện thoại phải có đúng 10 chữ số (hiện tại: {digits.Length} chữ số)\n" +
                        "Ví dụ hợp lệ: 0812.801.886", 
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
            // Tự động điền cbNoicap dựa trên ngày cấp CCCD
            try
            {
                if (cbNoicap == null || dateNgaycapCCCD == null) return;

                // Ngày cắt: 01/07/2024
                var ngayCat = new DateTime(2024, 7, 1);
                var ngayCap = dateNgaycapCCCD.Value.Date;

                // Từ 01/07/2024 trở đi: "Bộ Công an"
                // Trước đó: "Cục CSQLHC về TTXH"
                if (ngayCap >= ngayCat)
                {
                    cbNoicap.Text = "Bộ Công an";
                }
                else
                {
                    cbNoicap.Text = "Cục CSQLHC về TTXH";
                }
            }
            catch { }

            // Tính lại thời hạn CCCD khi ngày cấp thay đổi (vì nhóm tuổi dựa trên ngày cấp)
            try
            {
                if (dateNgaysinh != null && dateNgaysinh.CustomFormat != " ")
                    DateNgaysinh_ValueChanged(dateNgaysinh, EventArgs.Empty);
            }
            catch { }
        }

        // Xác thực tất cả các DateTimePicker để ngăn ngày tương lai (dùng cho các picker không có checkbox)
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

        // Validate DateTimePicker có ShowCheckBox=true: chỉ check sau khi rời ô và đã được tích chọn
        private void DatePickerChecked_Leave(object sender, EventArgs e)
        {
            try
            {
                var picker = sender as DateTimePicker;
                if (picker == null) return;

                // Bỏ qua nếu ô chưa được tích chọn
                if (!picker.Checked) return;

                // Chỉ validate khi đã nhập đủ ngày/tháng/năm (năm >= 1900)
                if (picker.Value.Year < 1900) return;

                if (picker.Value.Date > DateTime.Today)
                {
                    picker.Value = DateTime.Today;
                    MessageBox.Show("Ngày không được lớn hơn ngày hiện tại.", "Lỗi nhập liệu", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch { }
        }

        /// <summary>
        /// Khi user click vào DateTimePicker đang trống → hiện ngày để nhập
        /// </summary>
        private void DatePicker_Enter(object sender, EventArgs e)
        {
            try
            {
                var picker = sender as DateTimePicker;
                if (picker == null) return;
                if (picker.CustomFormat == " ")
                {
                    picker.Format = DateTimePickerFormat.Custom;
                    picker.CustomFormat = "dd/MM/yyyy";
                    picker.Value = DateTime.Today;
                }
            }
            catch { }
        }

        /// <summary>
        /// Khi dateLaphs tick checkbox → hiện ngày lập hồ sơ; bỏ tick → ẩn ngày
        /// </summary>
        private void DateLaphs_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (dateLaphs == null) return;
                dateLaphs.Format = DateTimePickerFormat.Custom;
                dateLaphs.CustomFormat = dateLaphs.Checked ? "dd/MM/yyyy" : " ";
                if (dateLaphs.Checked && dateLaphs.Value.Year < 1900)
                    dateLaphs.Value = DateTime.Today;
            }
            catch { }
        }

        /// <summary>
        /// Khi datentk tick checkbox → hiện ngày; bỏ tick → ẩn ngày
        /// </summary>
        private void DateNtk_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                var picker = sender as DateTimePicker;
                if (picker == null) return;
                picker.Format = DateTimePickerFormat.Custom;
                picker.CustomFormat = picker.Checked ? "dd/MM/yyyy" : " ";
            }
            catch { }
        }

        /// <summary>
        /// Tính ngày hết hạn CCCD dựa trên tuổi HÀNH CHÍNH (năm cấp − năm sinh) tại thời điểm cấp.
        /// - Tuổi cấp &lt; 25: hết hạn khi đủ 25 tuổi
        /// - Tuổi cấp 25–39: hết hạn khi đủ 40 tuổi
        /// - Tuổi cấp 40–59: hết hạn khi đủ 60 tuổi
        /// - Tuổi cấp ≥ 60: không thời hạn → trả về DateTime.MinValue
        /// Nếu mốc tính được cách ngày cấp dưới 2 năm → nhảy sang mốc kế tiếp
        /// (Bộ Công an không cấp CCCD có thời hạn còn lại dưới 2 năm kể từ ngày cấp).
        /// </summary>
        private DateTime CalcThoiHanCCCD(DateTime ngaysinh, DateTime ngaycap = default(DateTime))
        {
            if (ngaysinh == DateTime.MinValue) return DateTime.MinValue;
            // Dùng ngày cấp để xác định nhóm tuổi; nếu chưa có thì dùng ngày hiện tại
            var refDate = (ngaycap != default(DateTime) && ngaycap != DateTime.MinValue)
                          ? ngaycap.Date : DateTime.Today;

            // Tuổi hành chính VN = năm cấp − năm sinh (không điều chỉnh theo tháng/ngày sinh)
            int ageAtIssue = refDate.Year - ngaysinh.Year;

            // Chọn mốc kế tiếp: 25, 40, 60
            int[] milestones = { 25, 40, 60 };
            foreach (var m in milestones)
            {
                if (ageAtIssue < m)
                {
                    // Hết hạn đúng ngày tháng sinh, năm = năm sinh + mốc
                    int targetYear = ngaysinh.Year + m;
                    int day = ngaysinh.Day;
                    int month = ngaysinh.Month;
                    // Xử lý 29/02 trên năm không nhuận
                    if (month == 2 && day == 29 && !DateTime.IsLeapYear(targetYear))
                        day = 28;
                    var expiryDate = new DateTime(targetYear, month, day);

                    // Nếu mốc cách ngày cấp dưới 2 năm → nhảy sang mốc kế tiếp
                    // (Bộ Công an không cấp CCCD có thời hạn còn lại dưới 2 năm)
                    if ((expiryDate - refDate).TotalDays < 730)
                        continue;

                    return expiryDate;
                }
            }
            // Từ 60 tuổi trở lên (hoặc không còn mốc nào đủ xa) → không thời hạn
            return DateTime.MinValue;
        }

        /// <summary>
        /// Khi thay đổi ngày sinh → tự động tính và hiển thị thời hạn CCCD.
        /// </summary>
        private void DateNgaysinh_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                if (dateNgaysinh == null || datendhcccd == null) return;
                if (dateNgaysinh.CustomFormat == " ") return;

                var ngaysinh = dateNgaysinh.Value.Date;
                // Lấy ngày cấp hiện tại (nếu đã nhập) để tính đúng nhóm tuổi
                var ngaycap = (dateNgaycapCCCD != null && dateNgaycapCCCD.CustomFormat != " ")
                              ? dateNgaycapCCCD.Value.Date : DateTime.Today;
                var thoihan = CalcThoiHanCCCD(ngaysinh, ngaycap);

                datendhcccd.Format = DateTimePickerFormat.Custom;
                if (thoihan == DateTime.MinValue)
                {
                    // Từ 60 tuổi trở lên: không thời hạn — ẩn picker, hiển thị text
                    datendhcccd.CustomFormat = "'không thời hạn'";
                    datendhcccd.Value = new DateTime(ngaysinh.Year + 60, ngaysinh.Month,
                        (ngaysinh.Month == 2 && ngaysinh.Day == 29 && !DateTime.IsLeapYear(ngaysinh.Year + 60)) ? 28 : ngaysinh.Day);
                }
                else
                {
                    datendhcccd.CustomFormat = "dd/MM/yyyy";
                    datendhcccd.Value = thoihan;
                }
                datendhcccd.Enabled = false; // khoá, chỉ đọc
            }
            catch { }
        }

        // Validate thời hạn CCCD: cho phép ngày tương lai nhưng validate định dạng
        private void DateThoihanCCCD_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                var picker = sender as DateTimePicker;
                if (picker == null) return;

                var selectedDate = picker.Value;

                // Validate năm (không quá 5 chữ số)
                if (selectedDate.Year > 99999 || selectedDate.Year < 1)
                {
                    MessageBox.Show(
                        "Năm không hợp lệ. Vui lòng nhập năm từ 1 đến 99999.",
                        "Lỗi nhập liệu",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    picker.Value = DateTime.Today.AddYears(10);  // Mặc định 10 năm sau
                    return;
                }

                // Validate tháng (1-12) - đã tự động bởi DateTimePicker
                // Validate ngày (1-31) - đã tự động bởi DateTimePicker

                // Validate đặc biệt cho tháng 2: không quá 29 ngày
                if (selectedDate.Month == 2 && selectedDate.Day > 29)
                {
                    MessageBox.Show(
                        "Tháng 2 không thể có quá 29 ngày.",
                        "Lỗi nhập liệu",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    // Reset về ngày hợp lệ (29/02)
                    picker.Value = new DateTime(selectedDate.Year, 2, 29);
                    return;
                }

                // Validate năm nhuận cho ngày 29/02
                if (selectedDate.Month == 2 && selectedDate.Day == 29)
                {
                    if (!DateTime.IsLeapYear(selectedDate.Year))
                    {
                        MessageBox.Show(
                            $"Năm {selectedDate.Year} không phải năm nhuận. Tháng 2 chỉ có 28 ngày.",
                            "Lỗi nhập liệu",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        // Reset về 28/02
                        picker.Value = new DateTime(selectedDate.Year, 2, 28);
                        return;
                    }
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

        // Kiểm tra tổng cbSotien1 + cbSotien2 + cbSotien3 == cbSotien
        private void CbSotienSum_Changed(object sender, EventArgs e)
        {
            if (suppressMoneyChange) return;
            try
            {
                long total = ParseMoneyStringToLong(cbSotien.Text);
                long s1    = ParseMoneyStringToLong(cbSotien1.Text);
                long s2    = ParseMoneyStringToLong(cbSotien2.Text);
                long s3    = (cbSotien3 != null ? ParseMoneyStringToLong(cbSotien3.Text) : 0);
                bool anyDetail = s1 > 0 || s2 > 0 || s3 > 0;
                bool match  = anyDetail && (s1 + s2 + s3 == total);
                var okColor   = System.Drawing.Color.FromArgb(198, 239, 206);  // xanh nhạt
                var errColor  = System.Drawing.Color.FromArgb(255, 199, 206);  // đỏ nhạt
                var defColor  = System.Drawing.SystemColors.Window;
                var detailColor = anyDetail ? (match ? okColor : errColor) : defColor;
                cbSotien1.BackColor = detailColor;
                cbSotien2.BackColor = detailColor;
                if (cbSotien3 != null) cbSotien3.BackColor = detailColor;
                cbSotien.BackColor  = anyDetail ? (match ? okColor : errColor) : defColor;
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
            if (include03) list.Add(Get03DSTemplateName(c));
            // Lưu ý: GUQ chỉ nên được xuất khi người dùng bấm btnGUQ rõ ràng
            return list;
         }

        // Chọn template 03 DS phù hợp theo chương trình để tạo file khác nội dung, khác tên
        private string Get03DSTemplateName(Customer c)
        {
            var ct = (c?.Chuongtrinh ?? string.Empty).Trim();
            if (IsSxkdChuongtrinh(ct)) return "03 DS SXKD.docx";
            if (IsGqvlChuongtrinh(ct)) return "03 DS GQVL.docx";
            return "03 DS.docx";
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

                // Kiểm tra cụm từ đầy đủ "Hỗ trợ tạo việc làm duy trì và mở rộng việc làm"
                if (n.Contains("ho tro tao viec lam duy tri") ||
                    n.Contains("ho tro tao viec lam") && n.Contains("duy tri") && n.Contains("mo rong"))
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

        // ========== NÚT TẠO TOÀN BỘ HỒ SƠ ==========
        private async void Btnall_Click(object sender, EventArgs e)
        {
            // Kiểm tra đầy đủ thông tin trước khi xuất
            if (!ValidateRequiredFields()) return;
            if (!ValidateDuplicateCccdSdt()) return;

            try
            {
                var customer = ReadForm();
                string existingFile = null;
                if (customers != null && editingIndex >= 0 && editingIndex < customers.Count)
                {
                    try { existingFile = customers[editingIndex]._fileName; } catch { existingFile = null; }
                }

                if (!string.IsNullOrEmpty(existingFile)) customer._fileName = existingFile;

                SaveCustomerToFile(customer);
                UpsertCustomerInList(customer);

                // Danh sách file đã tạo
                var allCreatedFiles = new List<string>();

                // Tạo tất cả các mẫu
                await Task.Run(() =>
                {
                    try
                    {
                        var files01 = CreateProfileFromTemplate(customer, include03: false);
                        if (files01 != null) allCreatedFiles.AddRange(files01);
                    }
                    catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"Lỗi tạo mẫu 01: {ex.Message}"); }

                    try
                    {
                        var file03 = ExportSpecificTemplate(customer, "03 DS.docx");
                        if (!string.IsNullOrEmpty(file03)) allCreatedFiles.Add(file03);
                    }
                    catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"Lỗi tạo mẫu 03: {ex.Message}"); }

                    if (HasNguoiUyQuyen(customer))
                    {
                        try
                        {
                            var fileGUQ = ExportSpecificTemplate(customer, "GUQ.docx");
                            if (!string.IsNullOrEmpty(fileGUQ)) allCreatedFiles.Add(fileGUQ);
                        }
                        catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"Lỗi tạo mẫu GUQ: {ex.Message}"); }
                    }

                    try
                    {
                        var file01TGTV = ExportSpecificTemplate(customer, "01TGTV.docx");
                        if (!string.IsNullOrEmpty(file01TGTV)) allCreatedFiles.Add(file01TGTV);
                    }
                    catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"Lỗi tạo mẫu 01TGTV: {ex.Message}"); }

                    try
                    {
                        var fileBIA = ExportSpecificTemplate(customer, "BIA.docx");
                        if (!string.IsNullOrEmpty(fileBIA)) allCreatedFiles.Add(fileBIA);
                    }
                    catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"Lỗi tạo bìa hồ sơ: {ex.Message}"); }
                });

                BindGrid();
                ClearForm();

                // Hiển thị thông báo
                var guqLine = HasNguoiUyQuyen(customer)
                    ? "- Mẫu GUQ\n"
                    : "- Mẫu GUQ (bỏ qua - không có người uỷ quyền)\n";
                MessageBox.Show(
                    $"✅ Đã tạo toàn bộ hồ sơ thành công!\n\n" +
                    $"📄 Khách hàng: {customer.Hoten}\n" +
                    $"📁 Số file tạo: {allCreatedFiles.Count}\n\n" +
                    $"Bao gồm:\n" +
                    $"- Bìa hồ sơ (BIA)\n" +
                    $"- Mẫu 01 (TD/SXKD/GQVL)\n" +
                    $"- Mẫu 03 DS\n" +
                    guqLine +
                    $"- Mẫu 01TGTV",
                    "✅ Tạo toàn bộ hồ sơ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );

                // Mở tất cả file vừa tạo
                OpenCreatedFiles(allCreatedFiles);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"❌ Lỗi khi tạo toàn bộ hồ sơ:\n\n{ex.Message}",
                    "❌ Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        // Mở một file Word
        private void OpenFile(string filePath)
        {
            try
            {
                if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
                {
                    System.Diagnostics.Process.Start(filePath);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Không thể mở file {filePath}: {ex.Message}");
            }
        }

        // Mở nhiều file Word
        private void OpenCreatedFiles(List<string> filePaths)
        {
            if (filePaths == null || filePaths.Count == 0) return;

            try
            {
                foreach (var filePath in filePaths)
                {
                    OpenFile(filePath);
                    // Delay nhỏ để tránh mở quá nhiều cùng lúc
                    System.Threading.Thread.Sleep(500);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Lỗi khi mở files: {ex.Message}");
            }
        }



        private void toolTip1_Popup(object sender, PopupEventArgs e)
        {

        }

        // ============================================
        // TỰ ĐỘNG TÍNH NGÀY ĐẾN HẠN TỪ NGÀY GIẢI NGÂN + THỜI HẠN VAY
        // ============================================

        private bool _suppressDenHanCalc = false;

        /// <summary>
        /// dateDH bị khoá hoàn toàn — handler này không làm gì
        /// </summary>
        private void dateDH_ValueChanged(object sender, EventArgs e) { }

        /// <summary>
        /// Khi dateGn tick/bỏ tick hoặc ngày thay đổi → tính lại ngày đến hạn
        /// </summary>
        private void dateGn_ValueChanged(object sender, EventArgs e)
        {
            if (_suppressDenHanCalc) return;
            try
            {
                if (dateGn?.Checked == true)
                {
                    CalculateNgayDenHan();
                }
                else
                {
                    // Bỏ tick → ẩn ngày đến hạn
                    if (dateDH != null)
                        dateDH.CustomFormat = "          ";
                }
            }
            catch { }
        }

        /// <summary>
        /// Khi thời hạn vay thay đổi → cập nhật MaxDate của dateGn và tính lại
        /// </summary>
        private void CbThoihanvay_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                int months = ParseThoiHanThang(cbThoihanvay?.Text);
                if (months > 0 && dateGn != null)
                    dateGn.MaxDate = DateTime.Today.AddMonths(months);
            }
            catch { }
            CalculateNgayDenHan();
        }

        /// <summary>
        /// Tính ngày đến hạn = ngày giải ngân + số tháng thời hạn vay.
        /// Chỉ tính khi dateGn đang được tick.
        /// </summary>
        private void CalculateNgayDenHan()
        {
            if (_suppressDenHanCalc) return;
            try
            {
                if (dateGn?.Checked != true) return;
                if (dateDH == null) return;

                int months = ParseThoiHanThang(cbThoihanvay?.Text);
                if (months <= 0) return;

                var denHan = dateGn.Value.AddMonths(months);

                _suppressDenHanCalc = true;
                dateDH.CustomFormat = "dd/MM/yyyy";
                dateDH.Value = denHan;
            }
            catch { }
            finally { _suppressDenHanCalc = false; }
        }

        /// <summary>
        /// Trích xuất số tháng từ chuỗi thời hạn vay (vd: "60 tháng" → 60, "12" → 12)
        /// </summary>
        private int ParseThoiHanThang(string thoihanvay)
        {
            if (string.IsNullOrWhiteSpace(thoihanvay)) return 0;
            var match = Regex.Match(thoihanvay, @"\d+");
            if (match.Success && int.TryParse(match.Value, out int months))
                return months;
            return 0;
        }

        // ============================================
        // XỬ LÝ BẢNG KÊ TIỀN
        // ============================================

        /// <summary>
        /// Khởi tạo dgvTotruong để hiển thị danh sách tổ trưởng
        /// </summary>
        private void InitializeDgvTotruong()
        {
            try
            {
                if (dgvTotruong == null) return;

                dgvTotruong.AutoGenerateColumns = false;
                dgvTotruong.AllowUserToAddRows = false;
                dgvTotruong.AllowUserToDeleteRows = false;
                dgvTotruong.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                dgvTotruong.MultiSelect = false;
                dgvTotruong.ReadOnly = true;

                // Xóa cột cũ
                dgvTotruong.Columns.Clear();

                // Thêm cột Tổ trưởng
                dgvTotruong.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = "Totruong",
                    HeaderText = "Tổ trưởng",
                    DataPropertyName = "Totruong",
                    Width = 150
                });

                // Thêm cột Tổng tiền
                dgvTotruong.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = "TongTien",
                    HeaderText = "Tổng tiền (VNĐ)",
                    DataPropertyName = "TongTien",
                    Width = 150,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Alignment = DataGridViewContentAlignment.MiddleRight,
                        Format = "N0"
                    }
                });

                // Thêm cột Ngày tạo
                dgvTotruong.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = "NgayTao",
                    HeaderText = "Ngày tạo",
                    DataPropertyName = "NgayTao",
                    Width = 120,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Format = "dd/MM/yyyy HH:mm"
                    }
                });

                // Bind data
                dgvTotruong.DataSource = bangKeList;

                // Đăng ký sự kiện click để load bảng kê
                dgvTotruong.CellClick += DgvTotruong_CellClick;
            }
            catch { }
        }

        /// <summary>
        /// Xử lý click vào dgvTotruong - Load bảng kê của tổ trưởng đã chọn
        /// </summary>
        private void DgvTotruong_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (dgvTotruong == null || e.RowIndex < 0) return;

                var bangKe = dgvTotruong.Rows[e.RowIndex].DataBoundItem as BangKeData;
                if (bangKe == null) return;

                // Load tên tổ trưởng
                if (txtTotruong != null)
                {
                    txtTotruong.Text = bangKe.Totruong;
                }

                // Load chi tiết vào dgvbangke1 (bao gồm cả số tiền sổ sách)
                if (dgvbangke1 != null)
                {
                    BangKeTien.LoadDataFromBangKe(dgvbangke1, bangKe);
                }

                MessageBox.Show(
                    $"Đã load bảng kê của: {bangKe.Totruong}\n" +
                    $"Tổng tiền mặt: {bangKe.TongTien:N0} VNĐ\n" +
                    $"Sổ sách: {bangKe.SoTienSoSach:N0} VNĐ\n" +
                    $"Chênh lệch: {bangKe.ChenhLech:N0} VNĐ",
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch { }
        }

        /// <summary>
        /// Load danh sách bảng kê từ file JSON
        /// </summary>
        private void LoadBangKeFromFiles()
        {
            try
            {
                bangKeList.Clear();

                var loadedList = BangKeTien.LoadAllFromFiles();
                foreach (var bangKe in loadedList)
                {
                    bangKeList.Add(bangKe);
                }
            }
            catch { }
        }

        /// <summary>
        /// Lưu bảng kê tiền theo tổ trưởng
        /// </summary>
        private void BtnLuubangke_Click(object sender, EventArgs e)
        {
            try
            {
                // Kiểm tra tên tổ trưởng
                if (txtTotruong == null || string.IsNullOrWhiteSpace(txtTotruong.Text))
                {
                    MessageBox.Show("Vui lòng nhập tên Tổ trưởng!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string totruong = txtTotruong.Text.Trim();

                // Lấy dữ liệu từ dgvbangke1
                if (dgvbangke1 == null)
                {
                    MessageBox.Show("Không tìm thấy bảng kê tiền!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var chiTiet = BangKeTien.GetChiTiet(dgvbangke1);
                long tongTien = BangKeTien.GetTongThanhTien(dgvbangke1);
                long soTienSoSach = BangKeTien.GetSoTienSoSach(dgvbangke1);
                long chenhLech = BangKeTien.GetChenhLech(dgvbangke1);

                if (tongTien == 0)
                {
                    MessageBox.Show("Bảng kê chưa có dữ liệu (Tổng = 0)!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Kiểm tra xem tổ trưởng đã có trong danh sách chưa
                var existing = bangKeList.FirstOrDefault(b => string.Equals(b.Totruong, totruong, StringComparison.OrdinalIgnoreCase));

                if (existing != null)
                {
                    // Cập nhật bảng kê cũ
                    var result = MessageBox.Show($"Tổ trưởng '{totruong}' đã có bảng kê.\n\nBạn có muốn cập nhật?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        existing.ChiTiet = chiTiet;
                        existing.TongTien = tongTien;
                        existing.SoTienSoSach = soTienSoSach;
                        existing.ChenhLech = chenhLech;
                        existing.NgayTao = DateTime.Now;

                        // Lưu vào file JSON
                        BangKeTien.SaveToFile(existing);

                        // Refresh grid
                        if (dgvTotruong != null)
                        {
                            dgvTotruong.Refresh();
                        }

                        MessageBox.Show($"Đã cập nhật bảng kê cho: {totruong}\nTổng tiền mặt: {tongTien:N0} VNĐ\nSổ sách: {soTienSoSach:N0} VNĐ\nChênh lệch: {chenhLech:N0} VNĐ", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    // Thêm mới
                    var bangKe = new BangKeData
                    {
                        Totruong = totruong,
                        ChiTiet = chiTiet,
                        TongTien = tongTien,
                        SoTienSoSach = soTienSoSach,
                        ChenhLech = chenhLech,
                        NgayTao = DateTime.Now
                    };

                    // Lưu vào file JSON
                    BangKeTien.SaveToFile(bangKe);

                    // Thêm vào danh sách
                    bangKeList.Add(bangKe);

                    MessageBox.Show($"Đã lưu bảng kê cho: {totruong}\nTổng tiền mặt: {tongTien:N0} VNĐ\nSổ sách: {soTienSoSach:N0} VNĐ\nChênh lệch: {chenhLech:N0} VNĐ", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi lưu bảng kê: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Xóa bảng kê của tổ trưởng đã chọn trong dgvTotruong
        /// </summary>
        private void BtnXoabangke_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvTotruong == null || dgvTotruong.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Vui lòng chọn tổ trưởng cần xóa trong danh sách!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var bangKe = dgvTotruong.SelectedRows[0].DataBoundItem as BangKeData;
                if (bangKe == null) return;

                var result = MessageBox.Show($"Bạn có chắc muốn xóa bảng kê của '{bangKe.Totruong}'?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    // Xóa file JSON
                    BangKeTien.DeleteFile(bangKe);

                    // Xóa khỏi danh sách
                    bangKeList.Remove(bangKe);

                    MessageBox.Show($"Đã xóa bảng kê của: {bangKe.Totruong}", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi xóa bảng kê: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Tạo mới bảng kê - Reset về 0 để nhập mới
        /// </summary>
        private void BtnTaobangke_Click(object sender, EventArgs e)
        {
            try
            {
                var result = MessageBox.Show("Bạn có chắc muốn tạo bảng kê mới?\n\nDữ liệu hiện tại sẽ bị xóa (chưa lưu).", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    // Reset bảng kê
                    if (dgvbangke1 != null)
                    {
                        BangKeTien.ResetAll(dgvbangke1);
                    }

                    // Xóa tên tổ trưởng
                    if (txtTotruong != null)
                    {
                        txtTotruong.Text = "";
                    }

                    MessageBox.Show("Đã tạo bảng kê mới!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi tạo bảng kê mới: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ============================================
        // XỬ LÝ AUTO-UPDATE
        // ============================================

        /// <summary>
        /// Kiểm tra và CHẠY cập nhật tự động khi khởi động app
        /// Nếu có phiên bản mới, BẮT BUỘC cập nhật trước khi dùng app
        /// </summary>
        private async void CheckForUpdateOnStartup()
        {
            try
            {
                var updateInfo = await AutoUpdater.CheckForUpdateAsync(silent: true);

                if (updateInfo != null)
                {
                    bool updated = await AutoUpdater.ShowUpdateDialogAsync(updateInfo);

                    if (!updated)
                    {
                        Application.Exit();
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("✅ Không có cập nhật mới hoặc chưa có release.");
                }
            }
            catch
            {
                // Im lặng nếu có lỗi — app tiếp tục chạy bình thường
            }
        }

        /// <summary>
        /// Ghi đè thông tin khách hàng đang được gọi lên (btnUpdate)
        /// </summary>
        private void BtnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                if (editingIndex < 0 || customers == null || editingIndex >= customers.Count)
                {
                    MessageBox.Show("Vui lòng chọn một khách hàng từ danh sách trước khi cập nhật.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var customer = ReadForm();
                if (string.IsNullOrWhiteSpace(customer.Hoten))
                {
                    MessageBox.Show("Vui lòng nhập Họ và tên.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Giữ nguyên tên file JSON cũ để ghi đè đúng khách
                customer._fileName = customers[editingIndex]._fileName;

                SaveCustomerToFile(customer);
                customers[editingIndex] = customer;
                BindGrid();

                MessageBox.Show($"✅ Đã cập nhật thông tin khách hàng '{customer.Hoten}' thành công.", "Cập nhật", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi cập nhật khách hàng:\n\n{ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Xoá toàn bộ dữ liệu trên form để nhập lại (btnClear)
        /// </summary>
        private void BtnClear_Click(object sender, EventArgs e)
        {
            ClearForm();
        }

        private void btnLuubangke_Click_1(object sender, EventArgs e)
        {

        }

        private void label45_Click(object sender, EventArgs e)
        {

        }

        // ============================================
        // CBTINHFIX → CBPGDFIX → DGV1
        // ============================================

        private void CbTinhfix_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                _editTinhModel = null;
                if (cbpgdfix != null) cbpgdfix.Items.Clear();
                if (cbXafix != null) cbXafix.Items.Clear();

                string tinh = cbTinhfix?.Text?.Trim() ?? "";
                if (string.IsNullOrWhiteSpace(tinh)) return;

                _editTinhModel = TinhHelper.LoadTinhModel(tinh);
                if (_editTinhModel?.pgds == null) return;

                foreach (var pgd in _editTinhModel.pgds)
                    if (!string.IsNullOrWhiteSpace(pgd.pgd))
                        cbpgdfix.Items.Add(pgd.pgd);

                if (cbpgdfix.Items.Count > 0) cbpgdfix.SelectedIndex = 0;
            }
            catch { }
        }

        private void CbXafix_PopulateFromPgd(object sender, EventArgs e)
        {
            try
            {
                if (cbXafix == null) return;
                cbXafix.Items.Clear();
                xinManEditor?.FilterByCommune("");

                if (_editTinhModel?.pgds == null || cbpgdfix == null) return;
                if (string.IsNullOrWhiteSpace(cbpgdfix.Text)) return;

                var pgdEntry = _editTinhModel.pgds.FirstOrDefault(p =>
                    string.Equals(p.pgd, cbpgdfix.Text.Trim(), StringComparison.OrdinalIgnoreCase));

                if (pgdEntry?.communes == null) return;

                foreach (var commune in pgdEntry.communes)
                    if (!string.IsNullOrWhiteSpace(commune.name))
                        cbXafix.Items.Add(commune.name);
            }
            catch { }
        }

        private void CbXafix_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                var xa = cbXafix?.Text?.Trim() ?? "";
                xinManEditor?.FilterByCommune(xa);
            }
            catch { }
        }

        private void CbpgdfixEditor_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (_editTinhModel?.pgds == null || cbpgdfix == null || xinManEditor == null) return;
                if (string.IsNullOrWhiteSpace(cbpgdfix.Text)) return;

                var pgdEntry = _editTinhModel.pgds.FirstOrDefault(p =>
                    string.Equals(p.pgd, cbpgdfix.Text.Trim(), StringComparison.OrdinalIgnoreCase));

                if (pgdEntry != null)
                    xinManEditor.LoadFromPgdEntry(pgdEntry, _editTinhModel.tinh);
            }
            catch { }
        }

        // Thứ tự Tab theo luồng nhập liệu thực tế
        private static readonly string[] _inputTabOrder = new[]
        {
            "txtHoten",
            "dateNgaysinh", "cbGioitinh", "cbDantoc",
            "txtSocccd", "dateNgaycapCCCD",
            "txtSdt", "txtNhankhau", "cbNhandang",
            "cbTinh", "cbPGD", "cbXa", "cbThon", "cbHoi", "cbTo",
            "cbChuongtrinh", "cbVtc", "cbPhuongan",
            "cbmucdich1", "cbmucdich2", "cbDoituong1", "cbDoituong2",
            "cbThoihanvay", "cbPhanky",
            "cbSotien", "cbSotien1", "cbSotien2",
            "dateLaphs", "dateGn",
            "txtntk1", "datentk1", "txtcccd1", "cbqh1",
            "txtntk2", "datentk2", "txtcccd2", "cbqh2",
            "txtntk3", "datentk3", "txtcccd3", "cbqh3"
        };

        private List<System.Windows.Forms.Control> GetInputTabOrder()
        {
            var list = new List<System.Windows.Forms.Control>();
            foreach (var name in _inputTabOrder)
            {
                var found = this.Controls.Find(name, true);
                if (found.Length > 0 && found[0].Enabled && found[0].Visible)
                    list.Add(found[0]);
            }
            return list;
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Tab || keyData == (Keys.Tab | Keys.Shift))
            {
                var active = ActiveControl;

                if (active is DataGridView)
                    return base.ProcessCmdKey(ref msg, keyData);
                if (active is RichTextBox rtb && rtb.AcceptsTab)
                    return base.ProcessCmdKey(ref msg, keyData);
                if (active is TextBox tb && tb.AcceptsTab && tb.Multiline)
                    return base.ProcessCmdKey(ref msg, keyData);

                var controls = GetInputTabOrder();
                int idx = controls.IndexOf(active);
                if (idx >= 0)
                {
                    bool fwd = keyData == Keys.Tab;
                    int next = fwd ? (idx + 1) % controls.Count : (idx - 1 + controls.Count) % controls.Count;
                    controls[next].Focus();
                    return true;
                }

                SelectNextControl(active, keyData == Keys.Tab, true, true, true);
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void btnall_Click_1(object sender, EventArgs e)
        {

        }

        private void cbThoihanvay_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void label87_Click(object sender, EventArgs e)
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
    public string Sotien3 { get; set; }
    public string Soluong1 { get; set; }
    public string Soluong2 { get; set; }
    public string Soluong3 { get; set; }
    public string Sotientong { get; set; }
    public string Sotienchu { get; set; }
    public string Mucdich1 { get; set; }
    public string Mucdich2 { get; set; }
    public string Mucdich3 { get; set; }
    public string Doituong1 { get; set; }
    public string Doituong2 { get; set; }
    public DateTime Ngaylaphs { get; set; }
    public DateTime Ngaydenhan { get; set; }
    public DateTime Ngaygiaingaan { get; set; }
    public DateTime Thoihancccd { get; set; }
    /// <summary>"không thời hạn" nếu > 60 tuổi, ngày dd/MM/yyyy nếu còn hạn. Tính từ Ngaysinh.</summary>
    public string ThoihancccdText { get; set; }
    public string PGD { get; set; }
    public string Tinh { get; set; }

    public string Dantoc { get; set; }
    public string Sdt = "";
    public string Nhankhau = "";

    public string Ntk1 = "";
    public string Ntk2 = "";
    public string Ntk3 = "";
    public string Ntk4 = "";
    public string CccdNtk1 = "";
    public string CccdNtk2 = "";
    public string CccdNtk3 = "";
    public string CccdNtk4 = "";
    public string Namsinh1 = "";
    public string Namsinh2 = "";
    public string Namsinh3 = "";
    public string Namsinh4 = "";

    public string Qh1 = "";
    public string Qh2 = "";
    public string Qh3 = "";
    public string Qh4 = "";

    [JsonIgnore]
    public string _fileName { get; set; }
}

internal class XinManModel { public string pgd { get; set; } public List<Commune> communes { get; set; } = new List<Commune>(); }
internal class Commune { public string name { get; set; } public List<Association> associations { get; set; } = new List<Association>(); public List<Village> villages { get; set; } = new List<Village>(); }
internal class Association { public string name { get; set; } public string code { get; set; } public List<Village> villages { get; set; } = new List<Village>(); public List<string> managedVillages { get; set; } = new List<string>(); }
internal class Village { public string name { get; set; } public List<string> groups { get; set; } = new List<string>(); }
internal class TinhPgdEntry { public string pgd { get; set; } public List<Commune> communes { get; set; } = new List<Commune>(); }
internal class TinhModel { public string tinh { get; set; } public List<TinhPgdEntry> pgds { get; set; } = new List<TinhPgdEntry>(); }

#endregion
}
