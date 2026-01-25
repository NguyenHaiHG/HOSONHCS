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
        private const string TemplatesFolder = "Templates";
        private const string EmbeddedTemplateFileName = "01 HN.docx"; // embedded in project
        private BindingList<Customer> customers;
        private int editingIndex = -1;
        // XinMan model và instance trong bộ nhớ
        private XinManModel xinmanModel;
        // editor for xinman.json displayed in tabPage3
        private XinManEditor xinManEditor;
        private bool suppressComboChanged = false;
        // suppress flag for namsinh textbox text-change handlers
        private bool suppressNamsinhChanged = false;

        // cache for resolved template paths to avoid repeated extraction from resources
        private static readonly Dictionary<string, string> templatePathCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // guard for money formatting to avoid recursive TextChanged
        private bool suppressMoneyChange = false;

        public Form1()
        {
            InitializeComponent();
            InitializeApp();
        }

        private void InitializeApp()
        {
            // gắn sự kiện
            try { btn01.Click += BtnSave_Click; } catch { }
            try { btnDelete.Click += BtnDelete_Click; } catch { }

            // nute mới có thể chưa có trong designer cũ -> dùng try-catch
            try { btn03.Click += Btn03_Click; } catch { /* bỏ qua */ }
            try { btn03Group.Click += Btn03Group_Click; } catch { /* bỏ qua */ }
            try { btnGUQ.Click += BtnGUQ_Click; } catch { /* bỏ qua */ }

            // namsinh textbox: chỉ nhập số và định dạng dd/MM/yyyy
            try { txtnamsinh1.KeyPress += TxtNamsinh_KeyPress; txtnamsinh1.TextChanged += TxtNamsinh_TextChanged; } catch { }
            try { txtnamsinh2.KeyPress += TxtNamsinh_KeyPress; txtnamsinh2.TextChanged += TxtNamsinh_TextChanged; } catch { }
            try { txtnamsinh3.KeyPress += TxtNamsinh_KeyPress; txtnamsinh3.TextChanged += TxtNamsinh_TextChanged; } catch { }
            // CCCD fields: digits only and length validation
            try { txtSocccd.KeyPress += TxtDigitsOnly_KeyPress; txtSocccd.TextChanged += TxtCccd_TextChanged; txtSocccd.Leave += TxtCccd_Leave; } catch { }
            try { txtcccd1.KeyPress += TxtDigitsOnly_KeyPress; txtcccd1.TextChanged += TxtCccd_TextChanged; txtcccd1.Leave += TxtCccd_Leave; } catch { }
            try { txtcccd2.KeyPress += TxtDigitsOnly_KeyPress; txtcccd2.TextChanged += TxtCccd_TextChanged; txtcccd2.Leave += TxtCccd_Leave; } catch { }
            try { txtcccd3.KeyPress += TxtDigitsOnly_KeyPress; txtcccd3.TextChanged += TxtCccd_TextChanged; txtcccd3.Leave += TxtCccd_Leave; } catch { }

            // dateNgaycapCCCD: auto-select cbNoicap based on cutoff 01/07/2024
            try { dateNgaycapCCCD.ValueChanged += DateNgaycapCCCD_ValueChanged; } catch { }

            try { cbPGD.SelectedIndexChanged += CbPGD_SelectedIndexChanged; } catch { }
            try { cbXa.SelectedIndexChanged += CbXa_SelectedIndexChanged; } catch { }
            try { cbThon.SelectedIndexChanged += CbThon_SelectedIndexChanged; } catch { }
            try { cbHoi.SelectedIndexChanged += CbHoi_SelectedIndexChanged; } catch { }

            // money combo: only digits allowed and formatted with '.' thousands separator
            try { cbSotien.KeyPress += CbMoney_KeyPress; cbSotien.TextChanged += CbMoney_TextChanged; } catch { }

            try
            {
                // select full rows and allow multi-select for group export
                dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                dgv.MultiSelect = true;
                dgv.AutoGenerateColumns = true;
                dgv.ReadOnly = true;
                dgv.AllowUserToAddRows = false;
                dgv.AllowUserToDeleteRows = false;
                dgv.CellDoubleClick += Dgv_CellDoubleClick;
                dgv.CellClick += Dgv_CellClick;
                dgv.EditMode = DataGridViewEditMode.EditOnEnter;
            }
            catch { }

            // load dữ liệu và bind
            LoadXinManData();
            LoadCustomersFromFiles();
            BindGrid();

            // Attach XinManEditor to controls on tabPage3 (dgv1, login controls, search)
            try
            {
                xinManEditor = new XinManEditor();
                xinManEditor.AttachControls(dgv1, txtUsername, txtPassword, btnLogin, txtSearch);
            }
            catch { }
        }

        // -------------------------
        // Lưu / tải khách hàng (mỗi khách 1 file JSON) + DataGrid
        // -------------------------

        private string GetCustomersFolderPath()
        {
            // Thư mục Customers bên cạnh exe
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Customers");
        }

        private void EnsureCustomersFolder()
        {
            var folder = GetCustomersFolderPath();
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);
        }

        private void LoadCustomersFromFiles()
        {
            customers = new BindingList<Customer>();
            try
            {
                EnsureCustomersFolder();
                var folder = GetCustomersFolderPath();
                foreach (var file in Directory.EnumerateFiles(folder, "*.json"))
                {
                    try
                    {
                        var json = File.ReadAllText(file, Encoding.UTF8);
                        var c = JsonConvert.DeserializeObject<Customer>(json);
                        if (c != null)
                        {
                            c._fileName = Path.GetFileName(file);
                            customers.Add(c);
                        }
                    }
                    catch
                    {
                        // bỏ qua file lỗi
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load customer files: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            customers.ListChanged += Customers_ListChanged;
        }

        private void Customers_ListChanged(object sender, ListChangedEventArgs e)
        {
            // không tự động lưu; lưu khi người dùng bấm nút
        }

        private void BindGrid()
        {
            try
            {
                dgv.DataSource = null;
                dgv.DataSource = customers;

                // make all generated columns readonly; selection is by row(s)
                foreach (DataGridViewColumn col in dgv.Columns)
                {
                    col.ReadOnly = true;
                }

                // optionally show only columns you want. Ensure Hoten visible first
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

            if (editingIndex < 0)
            {
                var baseFolder = folder;
                int i = 1;
                while (Directory.Exists(folder))
                {
                    folder = baseFolder + "_" + i;
                    i++;
                }
            }
            return folder;
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

                    // OpenXML-only replacement
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
            // ensure Sotienchu is filled from numeric value for template 01
            try { EnsureSotienchuFromNumeric(c, docPath); } catch { }
             // Use OpenXML replacement helper
             var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
             {
                { "{{hoten}}", c.Hoten },
                { "{{socccd}}", c.Socccd },
                { "{{cccd}}", c.Socccd },
                { "{{gioitinh}}", c.GioiTinh },
                { "{{dantoc}}", c.Dantoc },
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
                // If a Doituong (combo) is selected show it in the template's 'mucdich' highlighted areas; otherwise use the free-text Mucdich fields
                { "{{mucdich1}}", !string.IsNullOrWhiteSpace(c.Doituong1) ? c.Doituong1 : (c.Mucdich1 ?? "") },
                { "{{mucdich2}}", !string.IsNullOrWhiteSpace(c.Doituong2) ? c.Doituong2 : (c.Mucdich2 ?? "") },
                { "{{doituong1}}", c.Doituong1 ?? "" },
                { "{{doituong2}}", c.Doituong2 ?? "" },
                { "{{doituong}}", !string.IsNullOrWhiteSpace(c.Doituong1) ? c.Doituong1 : (c.Doituong2 ?? "") },
                { "{{ngaylaphs}}", c.Ngaylaphs == DateTime.MinValue ? "" : c.Ngaylaphs.ToString("dd/MM/yyyy") },
                { "{{ngaysinh}}", c.Ngaysinh == DateTime.MinValue ? "" : c.Ngaysinh.ToString("dd/MM/yyyy") },
                { "{{phanky}}", c.Phanky },
                { "{{pgd}}", c.PGD },

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

            // For 03 DS template ensure specific placeholders use the values from the current customer (read from controls)
            try
            {
                if (Is03DS(docPath))
                {
                    // Map exactly as requested: Hoten -> {{hoten}}, Doituong (combo) -> {{doituong}}, Phuongan -> {{phuongan}}, Sotien (combo) -> {{sotien}}, Thoihanvay -> {{thoihanvay}}
                    replacements["{{hoten}}"] = c.Hoten ?? (replacements.ContainsKey("{{hoten}}") ? replacements["{{hoten}}"] : "");
                    // prefer Doituong1 (from cbDoituong) if present otherwise fallback to existing doituong replacement
                    var doituongVal = !string.IsNullOrWhiteSpace(c.Doituong1) ? c.Doituong1 : (!string.IsNullOrWhiteSpace(c.Doituong2) ? c.Doituong2 : "");
                    replacements["{{doituong}}"] = doituongVal;
                    replacements["{{phuongan}}"] = c.Phuongan ?? (replacements.ContainsKey("{{phuongan}}") ? replacements["{{phuongan}}"] : "");
                    replacements["{{sotien}}"] = c.Sotien ?? (replacements.ContainsKey("{{sotien}}") ? replacements["{{sotien}}"] : "");
                    replacements["{{thoihanvay}}"] = c.Thoihanvay ?? (replacements.ContainsKey("{{thoihanvay}}") ? replacements["{{thoihanvay}}"] : "");
                }
            }
            catch { }

            try { ReplacePlaceholdersUsingOpenXml(docPath, replacements, c); }
            catch (Exception ex) { MessageBox.Show("Error replacing placeholders (OpenXML): " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
         }

        // Try to compute `Sotienchu` from the numeric `Sotien` (or `Sotientong`) when producing templates
        private void EnsureSotienchuFromNumeric(Customer c, string docPath)
        {
            try
            {
                if (c == null) return;
                // only apply for template 01 (embedded or file names that contain '01')
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

        // Parse fully numeric string with dots/commas removed already in caller; wrapper left for clarity
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

        // Convert integer number to Vietnamese words (supports up to trillion-range reasonably)
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
                    // still need to append unit when inner groups exist (to keep place) only if some lower non-zero exists
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
                    // if hundreds == 0 and there is tens/units, when this group is not the highest, we may need 'không trăm'
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
            // Cleanup multiple spaces
            result = Regex.Replace(result, @"\s+", " ").Trim();
            // lowercase
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

        private void ReplacePlaceholdersUsingOpenXml(string docPath, Dictionary<string, string> replacements, Customer c = null, bool isGroup = false)
        {
            if (string.IsNullOrWhiteSpace(docPath) || !File.Exists(docPath)) throw new FileNotFoundException("Document not found: " + docPath);

            using (var wordDoc = WordprocessingDocument.Open(docPath, true))
            {
                var mainPart = wordDoc.MainDocumentPart;
                if (mainPart == null) return;

                try
                {
                    if (Is03DS(docPath))
                    {
                        try
                        {
                            var numberPlaceholderRegex = new Regex(@"\{\{\s*([a-zA-Z_]+)(\d+)\s*\}\}", RegexOptions.IgnoreCase);
                            foreach (var table in mainPart.Document.Descendants<Table>())
                            {
                                var rows = table.Elements<TableRow>().ToList();
                                foreach (var row in rows)
                                {
                                    try
                                    {
                                        var rowText = string.Concat(row.Descendants<Text>().Select(t => t.Text ?? ""));
                                        var m = numberPlaceholderRegex.Matches(rowText);
                                        bool remove = false;
                                        foreach (Match mm in m)
                                        {
                                            if (mm.Groups.Count >= 3)
                                            {
                                                if (int.TryParse(mm.Groups[2].Value, out int idx))
                                                {
                                                    if (idx > 1) { remove = true; break; }
                                                  }
                                                }
                                              }
                                              if (remove) row.Remove();
                                    }
                                    catch { }
                                }
                            }
                            mainPart.Document.Save();
                        }
                        catch { }
                    }
                }
                catch { }

                try
                {
                    if (ShouldShowFullNamsinh(docPath) && c != null)
                    {
                        var addressPlaceholders = new[] { "{{thon}}", "{{xa}}", "{{hoi}}" };
                        foreach (var table in mainPart.Document.Descendants<Table>())
                        {
                            try
                            {
                                foreach (var row in table.Elements<TableRow>())
                                {
                                    var rowText = string.Concat(row.Descendants<Text>().Select(t => t.Text ?? ""));
                                    if (rowText.IndexOf("{{ntk1}}", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        bool fill = !string.IsNullOrWhiteSpace(c.Ntk1);
                                        foreach (var text in row.Descendants<Text>())
                                        {
                                            foreach (var ph in addressPlaceholders)
                                            {
                                                if (text.Text != null && text.Text.IndexOf(ph, StringComparison.OrdinalIgnoreCase) >= 0)
                                                {
                                                    var value = "";
                                                    if (fill)
                                                    {
                                                        if (ph == "{{xa}}") value = c.Xa ?? "";
                                                        else if (ph == "{{thon}}") value = c.Thon ?? "";
                                                        else if (ph == "{{hoi}}") value = c.Hoi ?? "";
                                                    }
                                                    text.Text = ReplaceIgnoreCase(text.Text, ph, value);
                                                }
                                            }
                                        }
                                        continue;
                                    }

                                    if (rowText.IndexOf("{{ntk2}}", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        bool fill = !string.IsNullOrWhiteSpace(c.Ntk2);
                                        foreach (var text in row.Descendants<Text>())
                                        {
                                            foreach (var ph in addressPlaceholders)
                                            {
                                                if (text.Text != null && text.Text.IndexOf(ph, StringComparison.OrdinalIgnoreCase) >= 0)
                                                {
                                                    var value = "";
                                                    if (fill)
                                                    {
                                                        if (ph == "{{xa}}") value = c.Xa ?? "";
                                                        else if (ph == "{{thon}}") value = c.Thon ?? "";
                                                        else if (ph == "{{hoi}}") value = c.Hoi ?? "";
                                                    }
                                                    text.Text = ReplaceIgnoreCase(text.Text, ph, value);
                                                }
                                            }
                                        }
                                        continue;
                                    }

                                    if (rowText.IndexOf("{{ntk3}}", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        bool fill = !string.IsNullOrWhiteSpace(c.Ntk3);
                                        foreach (var text in row.Descendants<Text>())
                                        {
                                            foreach (var ph in addressPlaceholders)
                                            {
                                                if (text.Text != null && text.Text.IndexOf(ph, StringComparison.OrdinalIgnoreCase) >= 0)
                                                {
                                                    var value = "";
                                                    if (fill)
                                                        if (ph == "{{xa}}") value = c.Xa ?? "";
                                                        else if (ph == "{{thon}}") value = c.Thon ?? "";
                                                        else if (ph == "{{hoi}}") value = c.Hoi ?? "";
                                                    text.Text = ReplaceIgnoreCase(text.Text, ph, value);
                                                }
                                            }
                                        }
                                        continue;
                                    }
                                }
                            }
                            catch { }
                        }
                        mainPart.Document.Save();
                    }
                }
                catch { }

                // Parts to process: main, headers, footers, footnotes, endnotes, comments
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
                        // Try to replace placeholders that may be split across multiple runs/text nodes
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
                                    newText = ReplaceIgnoreCase(newText, kv.Key, kv.Value ?? "");
                                }
                            }
                            if (newText != t.Text) t.Text = newText;
                        }
                        part.RootElement.Save();
                    }
                    catch { }
                }

                // After performing replacements, clean up 03 DS tables: remove rows that still contain numbered placeholders (>1) or are empty
                try
                {
                    if (Is03DS(docPath))
                    {
                        var numberPlaceholderRegex = new Regex(@"\{\{\s*([a-zA-Z_]+)(\d+)\s*\}\}", RegexOptions.IgnoreCase);
                        foreach (var table in mainPart.Document.Descendants<Table>())
                        {
                            var rows = table.Elements<TableRow>().ToList();
                            foreach (var row in rows)
                            {
                                try
                                {
                                    var rowText = string.Concat(row.Descendants<Text>().Select(t => t.Text ?? ""));

                                    // If row still contains numbered placeholders with index > 1 remove it
                                    var m = numberPlaceholderRegex.Matches(rowText);
                                    bool remove = false;
                                    foreach (Match mm in m)
                                    {
                                        if (mm.Groups.Count >= 3)
                                        {
                                            if (int.TryParse(mm.Groups[2].Value, out int idx))
                                            {
                                                if (idx > 1) { remove = true; break; }
                                            }
                                        }
                                    }

                                    if (remove)
                                    {
                                        row.Remove();
                                        continue;
                                    }

                                    // If row is empty or has only whitespace after replacements, remove it to avoid merged/concatenated cells
                                    var cleaned = Regex.Replace(rowText ?? "", "\\s+", "");
                                    if (string.IsNullOrEmpty(cleaned))
                                    {
                                        row.Remove();
                                        continue;
                                    }
                                }
                                catch { }
                            }

                            // Fix concatenation between STT cell (digits) and next cell starting with letters: insert a space at start of next cell's first Text node
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
                                                    // prepend a space to the first text node of the right cell
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

                // Save main document
                mainPart.Document.Save();
            }
        }

        // Attempt to find and replace placeholders that are split across multiple Text nodes (runs)
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

                    // normalize key token: if key is like "{{name}}" extract inner token "name"
                    string token = rawKey;
                    var m = Regex.Match(rawKey, "^\\s*\\{\\{\\s*(.*?)\\s*\\}\\}\\s*$");
                    if (m.Success && m.Groups.Count > 1) token = m.Groups[1].Value;
                    if (string.IsNullOrEmpty(token)) continue;

                    // build regex to find placeholder allowing optional spaces inside braces
                    var pattern = "\\{\\{\\s*" + Regex.Escape(token) + "\\s*\\}\\}";

                    int i = 0;
                    while (i < texts.Count)
                    {
                        var sb = new StringBuilder();
                        int j = i;
                        // build up to a reasonable window
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

                            // Ensure we don't accidentally concatenate neighboring content without a space.
                            // If prefix ends with non-whitespace and replacement starts with non-whitespace, add a space.
                            var finalReplacement = replacement ?? string.Empty;
                            if (!string.IsNullOrEmpty(prefix) && !char.IsWhiteSpace(prefix[prefix.Length - 1]) && finalReplacement.Length > 0 && !char.IsWhiteSpace(finalReplacement[0]))
                                finalReplacement = " " + finalReplacement;

                            // If replacement ends with non-whitespace and suffix starts with non-whitespace, insert a space after replacement
                            if (finalReplacement.Length > 0 && !char.IsWhiteSpace(finalReplacement[finalReplacement.Length - 1]) && !string.IsNullOrEmpty(suffix) && !char.IsWhiteSpace(suffix[0]))
                                finalReplacement = finalReplacement + " ";

                            texts[startNode].Text = prefix + finalReplacement + suffix;

                            for (int k = startNode + 1; k <= endNode; k++)
                            {
                                texts[k].Text = string.Empty;
                            }

                            // refresh texts collection and continue after replaced node
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
                // If oldValue is a double-brace placeholder like {{name}}, allow optional spaces inside braces in the document
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

        // Designer expects this exact-cased handler name in several places
        private void cbPGD_SelectedIndexChanged(object sender, EventArgs e)
        {
            CbPGD_SelectedIndexChanged(sender, e);
        }

        // Designer stubs referenced in Designer.cs
        private void richTextBox1_TextChanged(object sender, EventArgs e) { }
        private void Form1_Load(object sender, EventArgs e) { }

         // -------------------------
         // Xử lý UI / CRUD
         // -------------------------

         private void Dgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
         {
             // no-op: checkbox column removed; handler kept to avoid designer errors
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
                 MessageBox.Show("Lưu thành công.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
             }
             catch (Exception ex)
             {
                 MessageBox.Show("Lỗi khi lưu: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

             try { if (txtnamsinh1 != null) namsinh1 = txtnamsinh1.Text.Trim(); } catch { }
             try { if (txtnamsinh2 != null) namsinh2 = txtnamsinh2.Text.Trim(); } catch { }
             try { if (txtnamsinh3 != null) namsinh3 = txtnamsinh3.Text.Trim(); } catch { }

             try { if (cbqh1 != null) qh1 = cbqh1.Text.Trim(); } catch { }
             try { if (cbqh2 != null) qh2 = cbqh2.Text.Trim(); } catch { }
             try { if (cbqh3 != null) qh3 = cbqh3.Text.Trim(); } catch { }

             return new Customer
             {
                 Hoten = ToTitleCase(txtHoten.Text.Trim()),
                 Socccd = txtSocccd.Text,
                 GioiTinh = (cbGioitinh != null ? cbGioitinh.Text : ""),
                 Nhandang = cbNhandang.Text,
                 Ngaycap = dateNgaycapCCCD.Value.Date,
                 Ngaysinh = (dateNgaysinh != null ? dateNgaysinh.Value.Date : DateTime.MinValue),
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
                 Ngaylaphs = dateLaphs.Value.Date,
                 Ngaydenhan = (dateDH != null ? dateDH.Value.Date : DateTime.MinValue),
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
             txtHoten.Text = c.Hoten; txtSocccd.Text = c.Socccd ?? ""; cbNhandang.Text = c.Nhandang; dateNgaycapCCCD.Value = c.Ngaycap == DateTime.MinValue ? DateTime.Today : c.Ngaycap;
             try { if (dateNgaysinh != null) dateNgaysinh.Value = c.Ngaysinh == DateTime.MinValue ? DateTime.Today : c.Ngaysinh; } catch { }
             cbNoicap.Text = c.Noicap;
             suppressComboChanged = true;
             try
             {
                 cbPGD.Text = c.PGD; LoadXinManData();
                 if (cbXa.Items.Count == 0 && xinmanModel != null) foreach (var com in xinmanModel.communes) if (!string.IsNullOrWhiteSpace(com.name) && !cbXa.Items.Contains(com.name)) cbXa.Items.Add(com.name);
                 try { if (!string.IsNullOrEmpty(c.Xa)) { if (!cbXa.Items.Cast<object>().Any(x => string.Equals((x ?? "").ToString(), c.Xa, StringComparison.OrdinalIgnoreCase))) cbXa.Items.Add(c.Xa); cbXa.Text = c.Xa; } else cbXa.Text = ""; } catch { }
                 try { if (!string.IsNullOrEmpty(c.Thon)) { if (!cbThon.Items.Cast<object>().Any(x => string.Equals((x ?? "").ToString(), c.Thon, StringComparison.OrdinalIgnoreCase))) cbThon.Items.Add(c.Thon); cbThon.Text = c.Thon; } else cbThon.Text = ""; } catch { }
                 try { if (!string.IsNullOrEmpty(c.Hoi)) { if (!cbHoi.Items.Cast<object>().Any(x => string.Equals((x ?? "").ToString(), c.Hoi, StringComparison.OrdinalIgnoreCase))) cbHoi.Items.Add(c.Hoi); cbHoi.Text = c.Hoi; } else cbHoi.Text = ""; } catch { }
                 try { var toVal = !string.IsNullOrEmpty(c.Totruong) ? c.Totruong : (!string.IsNullOrEmpty(c.To) ? c.To : ""); if (!string.IsNullOrEmpty(toVal)) { if (!cbTo.Items.Cast<object>().Any(x => string.Equals((x ?? "").ToString(), toVal, StringComparison.OrdinalIgnoreCase))) cbTo.Items.Add(toVal); cbTo.Text = toVal; } else cbTo.Text = ""; } catch { }
                 cbSotien1.Text = c.Sotien1 ?? ""; cbSotien2.Text = c.Sotien2 ?? "";
                 try { if (cbVtc != null) cbVtc.Text = c.Vtc ?? ""; } catch { }
             }
             finally { suppressComboChanged = false; }

             cbChuongtrinh.Text = c.Chuongtrinh; cbThoihanvay.Text = c.Thoihanvay; cbSotien.Text = c.Sotien; txtMucdich1.Text = c.Mucdich1; txtMucdich2.Text = c.Mucdich2;
             dateLaphs.Value = c.Ngaylaphs == DateTime.MinValue ? DateTime.Today : c.Ngaylaphs;
             try { if (txtntk1 != null) txtntk1.Text = c.Ntk1 ?? ""; } catch { }
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

        private void UpsertCustomerInList(Customer customer)
        {
            if (customers == null) return;
            try { if (editingIndex >= 0 && editingIndex < customers.Count) customers[editingIndex] = customer; else { customers.Add(customer); editingIndex = customers.IndexOf(customer); } }
            catch { try { customers.Add(customer); editingIndex = customers.IndexOf(customer); } catch { } }
        }

        private void TxtNamsinh_KeyPress(object sender, KeyPressEventArgs e) { if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '/' && e.KeyChar != '-') e.Handled = true; }
        private void TxtNamsinh_TextChanged(object sender, EventArgs e) { if (suppressNamsinhChanged) return; try { var tb = sender as TextBox; if (tb == null) return; var originalText = tb.Text ?? string.Empty; var txt = originalText.Trim(); if (string.IsNullOrEmpty(txt)) return; var origSel = tb.SelectionStart; string digits = new string(txt.Where(char.IsDigit).ToArray()); DateTime dt; if (digits.Length == 6) { var d = int.Parse(digits.Substring(0, 2)); var m = int.Parse(digits.Substring(2, 2)); var yy = int.Parse(digits.Substring(4, 2)); int year = (yy >= 50) ? 1900 + yy : 2000 + yy; if (d >= 1 && d <= 31 && m >= 1 && m <= 12) { dt = new DateTime(year, m, d); var formatted = dt.ToString("dd/MM/yyyy"); if (!string.Equals(formatted, originalText, StringComparison.Ordinal)) { suppressNamsinhChanged = true; tb.Text = formatted; tb.SelectionStart = formatted.Length; suppressNamsinhChanged = false; } return; } } if (digits.Length == 8) { var d = int.Parse(digits.Substring(0, 2)); var m = int.Parse(digits.Substring(2, 2)); var yyyy = int.Parse(digits.Substring(4, 4)); if (d >= 1 && d <= 31 && m >= 1 && m <= 12) { dt = new DateTime(yyyy, m, d); var formatted = dt.ToString("dd/MM/yyyy"); if (!string.Equals(formatted, originalText, StringComparison.Ordinal)) { suppressNamsinhChanged = true; tb.Text = formatted; tb.SelectionStart = formatted.Length; suppressNamsinhChanged = false; } return; } } var formats = new[] { "d/M/yyyy", "dd/MM/yyyy", "d-M-yyyy", "dd-MM-yyyy" }; if (DateTime.TryParseExact(txt, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt)) { var formatted = dt.ToString("dd/MM/yyyy"); if (!string.Equals(formatted, originalText, StringComparison.Ordinal)) { suppressNamsinhChanged = true; tb.Text = formatted; int delta = formatted.Length - originalText.Length; int newSel = origSel + delta; if (newSel < 0) newSel = 0; if (newSel > formatted.Length) newSel = formatted.Length; tb.SelectionStart = newSel; suppressNamsinhChanged = false; } return; } } catch { } }

        private DateTime ParseDateTextOrFallback(string text) { if (string.IsNullOrWhiteSpace(text)) return DateTime.MinValue; try { var digits = new string((text ?? "").Where(char.IsDigit).ToArray()); if (digits.Length == 6) { var d = int.Parse(digits.Substring(0, 2)); var m = int.Parse(digits.Substring(2, 2)); var yy = int.Parse(digits.Substring(4, 2)); int year = (yy >= 50) ? 1900 + yy : 2000 + yy; return new DateTime(year, m, d); } if (digits.Length == 8) { var d = int.Parse(digits.Substring(0, 2)); var m = int.Parse(digits.Substring(2, 2)); var yyyy = int.Parse(digits.Substring(4, 4)); return new DateTime(yyyy, m, d); } if (digits.Length == 4) return DateTime.MinValue; DateTime dt; var formats = new[] { "d/M/yyyy", "dd/MM/yyyy", "yyyy", "yyyyMMdd", "d-M-yyyy", "dd-MM-yyyy" }; if (DateTime.TryParseExact(text, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt)) return dt; if (DateTime.TryParse(text, out dt)) return dt; } catch { } return DateTime.MinValue; }

        private string ToTitleCase(string input) { if (string.IsNullOrWhiteSpace(input)) return input; try { var culture = new CultureInfo("vi-VN"); var cleaned = Regex.Replace(input.Trim(), "\\s+", " ").ToLower(culture); return culture.TextInfo.ToTitleCase(cleaned); } catch { return input; } }

        private string FormatNamsinhStringForDoc(string namsinh, string docPath) { if (string.IsNullOrWhiteSpace(namsinh)) return ""; try { DateTime dt; var formats = new[] { "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "d-M-yyyy", "yyyyMMdd", "yyyy" }; if (DateTime.TryParseExact(namsinh.Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt)) return ShouldShowFullNamsinh(docPath) ? dt.ToString("dd/MM/yyyy") : dt.Year.ToString(); if (namsinh.Length == 8) { var d = int.Parse(namsinh.Substring(0, 2)); var m = int.Parse(namsinh.Substring(2, 2)); var yyyy = int.Parse(namsinh.Substring(4, 4)); dt = new DateTime(yyyy, m, d); return ShouldShowFullNamsinh(docPath) ? dt.ToString("dd/MM/yyyy") : dt.Year.ToString(); } if (namsinh.Length == 6) { var d = int.Parse(namsinh.Substring(0, 2)); var m = int.Parse(namsinh.Substring(2, 2)); var yy = int.Parse(namsinh.Substring(4, 2)); int year = (yy >= 50) ? 1900 + yy : 2000 + yy; dt = new DateTime(year, m, d); return ShouldShowFullNamsinh(docPath) ? dt.ToString("dd/MM/yyyy") : dt.Year.ToString(); } var ds = new string(namsinh.Where(char.IsDigit).ToArray()); if (ds.Length >= 4) { var ypart = ds.Substring(ds.Length - 4); if (int.TryParse(ypart, out int yv)) return yv.ToString(); } } catch { } return namsinh; }

        private string ExtractYearString(string namsinh) { if (string.IsNullOrWhiteSpace(namsinh)) return null; try { var digits = new string(namsinh.Where(char.IsDigit).ToArray()); if (digits.Length >= 4) { var yearPart = digits.Substring(digits.Length - 4); if (int.TryParse(yearPart, out int year)) return year.ToString(); } } catch { } return null; }

        private bool ShouldShowFullNamsinh(string docPath) { if (string.IsNullOrWhiteSpace(docPath)) return false; var name = Path.GetFileName(docPath) ?? ""; return name.IndexOf("GUQ", StringComparison.OrdinalIgnoreCase) >= 0; }
        private bool Is03DS(string docPath) { if (string.IsNullOrWhiteSpace(docPath)) return false; var name = Path.GetFileName(docPath) ?? ""; return (name.IndexOf("03", StringComparison.OrdinalIgnoreCase) >= 0 && name.IndexOf("DS", StringComparison.OrdinalIgnoreCase) >= 0) || name.IndexOf("03 DS", StringComparison.OrdinalIgnoreCase) >= 0; }

        // Resolve template path by checking several locations and embedded resources; caches results
        private string ResolveTemplatePath(string templateFileName)
        {
            if (string.IsNullOrWhiteSpace(templateFileName)) throw new FileNotFoundException("Template name is empty.");
            lock (templatePathCache)
            {
                if (templatePathCache.TryGetValue(templateFileName, out var cached) && !string.IsNullOrWhiteSpace(cached) && File.Exists(cached))
                    return cached;
            }

            // 1) Check Templates folder next to exe
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var candidate = Path.Combine(baseDir, TemplatesFolder, templateFileName);
            if (File.Exists(candidate))
            {
                lock (templatePathCache) { templatePathCache[templateFileName] = candidate; }
                return candidate;
            }

            // 2) Check baseDir root
            candidate = Path.Combine(baseDir, templateFileName);
            if (File.Exists(candidate)) { lock (templatePathCache) { templatePathCache[templateFileName] = candidate; } return candidate; }

            // 3) Recursive search under baseDir
            try
            {
                var found = Directory.EnumerateFiles(baseDir, templateFileName, SearchOption.AllDirectories).FirstOrDefault();
                if (!string.IsNullOrEmpty(found)) { lock (templatePathCache) { templatePathCache[templateFileName] = found; } return found; }
            }
            catch { }

            // 4) Embedded resource extraction
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

        // Very small validation: check file exists and is a .docx; try opening with OpenXML if possible
        private bool IsDocxFile(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return false;
                if (!string.Equals(Path.GetExtension(path), ".docx", StringComparison.OrdinalIgnoreCase)) return false;
                // Try opening as WordprocessingDocument to ensure it's a valid package
                try
                {
                    using (var w = WordprocessingDocument.Open(path, false)) { /* success */ }
                    return true;
                }
                catch { return false; }
            }
            catch { return false; }
        }

        // --- Added missing handlers and helpers ---

        // Designer referenced empty handlers
        private void tabPage1_Click(object sender, EventArgs e) { }
        private void groupBox2_Enter(object sender, EventArgs e) { }
        private void txtnamsinh2_TextChanged(object sender, EventArgs e) { TxtNamsinh_TextChanged(sender, e); }
        private void textBox1_TextChanged(object sender, EventArgs e) { }
        private void label14_Click(object sender, EventArgs e) { }
        private void cbXa_SelectedIndexChanged_1(object sender, EventArgs e) { CbXa_SelectedIndexChanged(sender, e); }
        private void label2_Click(object sender, EventArgs e) { }

        // CCCD helpers
        private void TxtDigitsOnly_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar)) e.Handled = true;
        }
        private void TxtCccd_TextChanged(object sender, EventArgs e)
        {
            // optional: enforce length limits or formatting; keep simple
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
            // no-op for now; placeholder for validation on leave
        }

        private void DateNgaycapCCCD_ValueChanged(object sender, EventArgs e)
        {
            // optional behavior: adjust Noicap based on a cutoff date (01/07/2024). Keep minimal: no-op
        }

        // Money formatting
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

        // Compute Sotientong and Sotienchu if possible
        private void UpdateComputedFields(Customer c)
        {
            if (c == null) return;
            try
            {
                // Compute total as Vốn tự có (Vtc) + Vốn vay (Sotien)
                if (string.IsNullOrWhiteSpace(c.Sotientong))
                {
                    long loan = ParseMoneyStringToLong(c.Sotien);
                    long own = ParseMoneyStringToLong(c.Vtc);
                    long total = loan + own;
                    if (total > 0)
                    {
                        // format with thousands separator using '.' as thousands separator
                        var formatted = string.Format(CultureInfo.InvariantCulture, "{0:N0}", total).Replace(",", ".");
                        c.Sotientong = formatted;
                    }
                }

                // generate Sotienchu if missing and have numeric value in Sotientong
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
            // If the selected program indicates "sản xuất kinh doanh" (SXKD) use the specific 01 SXKD template
            try
            {
                var ct = (c?.Chuongtrinh ?? "").Trim();
                if (!string.IsNullOrEmpty(ct) && IsSxkdChuongtrinh(ct))
                {
                    list.Add("01 SXKD.docx");
                }
                else if (!string.IsNullOrEmpty(ct) && IsGqvlChuongtrinh(ct))
                {
                    // Use GQVL variant when program indicates GQVL
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
            // Note: GUQ should only be exported when the user explicitly clicks btnGUQ
            return list;
         }

        // Detect common variants indicating GQVL program
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
                // check for common shorthand/variants for GQVL
                if (n.Contains("gqvl") || n.Contains("gq vl") || n.Contains("gq-vl") || n.Contains("gq_vl"))
                    return true;
            }
            catch { }
            return false;
        }

        // Detect common variants indicating "Sản xuất kinh doanh" (SXKD)
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
                // check keywords without diacritics
                if (n.Contains("san xuat kinh doanh") || n.Contains("sxkd") || n.Contains("san xuat") || n.Contains("kinh doanh"))
                    return true;
            }
            catch { }
            return false;
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
