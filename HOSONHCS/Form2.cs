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
    public partial class Form2 : Form
    {
        // store selected customers when opened as group
        private List<Customer> selectedCustomers = new List<Customer>();

        public Form2()
        {
            InitializeComponent();
            btn03to.Click += Btn03to_Click;
            btnxoa.Click += BtnXoa_Click;

            // wire money input handlers for cbtien1..5 (format with '.' thousands separator)
            try { cbtien1.KeyPress += CbMoney_KeyPress; cbtien1.TextChanged += CbMoney_TextChanged; } catch { }
            try { cbtien2.KeyPress += CbMoney_KeyPress; cbtien2.TextChanged += CbMoney_TextChanged; } catch { }
            try { cbtien3.KeyPress += CbMoney_KeyPress; cbtien3.TextChanged += CbMoney_TextChanged; } catch { }
            try { cbtien4.KeyPress += CbMoney_KeyPress; cbtien4.TextChanged += CbMoney_TextChanged; } catch { }
            try { cbtien5.KeyPress += CbMoney_KeyPress; cbtien5.TextChanged += CbMoney_TextChanged; } catch { }

            // wire letter-only handlers for name/location fields
            try { txtxa.KeyPress += TextLettersOnly_KeyPress; } catch { }
            try { txttotruong.KeyPress += TextLettersOnly_KeyPress; } catch { }
            try { txtkh1.KeyPress += TextLettersOnly_KeyPress; txtkh2.KeyPress += TextLettersOnly_KeyPress; txtkh3.KeyPress += TextLettersOnly_KeyPress; txtkh4.KeyPress += TextLettersOnly_KeyPress; txtkh5.KeyPress += TextLettersOnly_KeyPress; } catch { }
            // txtnd1..txtnd5 do not exist in this designer; skip wiring for them
        }

        // Constructor used when opening Form2 for a selected group from Form1
        public Form2(List<Customer> selected) : this()
        {
            try
            {
                if (selected != null)
                {
                    // Instead of copying every selected customer into the grid, group by Totruong
                    // and show one row per Totruong (the group leader). This prevents showing
                    // multiple member rows from Form1 inside the DataGridView.
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
                            if (string.IsNullOrWhiteSpace(leaderName))
                                continue; // skip entries without a Totruong value

                            var first = g.FirstOrDefault();
                            var summary = new Customer
                            {
                                // show the group leader's name in the grid's primary name column
                                Hoten = leaderName,
                                Totruong = first?.Totruong ?? leaderName,
                                Xa = first?.Xa ?? "",
                                Chuongtrinh = first?.Chuongtrinh ?? ""
                            };
                            selectedCustomers.Add(summary);
                        }

                        try { dataGridView1.DataSource = null; dataGridView1.DataSource = selectedCustomers; } catch { }
                    }
                    catch { }

                    // Do NOT fill txtkh1..txtkh5 from the selected members passed from Form1
                    // The form should show only the organizer info by default. Prefill only
                    // the top-level common fields from the first selected customer if present.
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

        // ================= CLICK XUẤT WORD =================
        private async void Btn03to_Click(object sender, EventArgs e)
        {
            try
            {
                // -------- THÔNG TIN CHUNG --------
                string totruong = Clean(txttotruong.Text);
                string xa = Clean(txtxa.Text);
                string chuongtrinh = Clean(cbctr.Text);

                // -------- TỔ VIÊN --------
                string[] hoten = {
                    "",
                    Clean(txtkh1.Text),
                    Clean(txtkh2.Text),
                    Clean(txtkh3.Text),
                    Clean(txtkh4.Text),
                    Clean(txtkh5.Text)
                };

                string[] sotien = {
                    "",
                    Money(cbtien1.Text),
                    Money(cbtien2.Text),
                    Money(cbtien3.Text),
                    Money(cbtien4.Text),
                    Money(cbtien5.Text)
                };

                string[] phuongan = {
                    "",
                    Clean(txtmd1.Text),
                    Clean(txtmd2.Text),
                    Clean(txtmd3.Text),
                    Clean(txtmd4.Text),
                    Clean(txtmd5.Text)
                };

                string[] thoihan = {
                    "",
                    Clean(cbtime1.Text),
                    Clean(cbtime2.Text),
                    Clean(cbtime3.Text),
                    Clean(cbtime4.Text),
                    Clean(cbtime5.Text)
                };

                string[] doituong = {
                    "",
                    Clean(cbdt1.Text),
                    Clean(cbdt2.Text),
                    Clean(cbdt3.Text),
                    Clean(cbdt4.Text),
                    Clean(cbdt5.Text)
                };

                int soNguoi = Enumerable.Range(1, 5).Count(i => !string.IsNullOrWhiteSpace(hoten[i]));
                if (soNguoi < 2)
                {
                    MessageBox.Show("Phải có tối thiểu 2 người mới được tạo mẫu.", "Cảnh báo");
                    return;
                }

                // -------- MAP PLACEHOLDER --------
                var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["{{totruong}}"] = totruong,
                    ["{{xa}}"] = xa,
                    ["{{chuongtrinh}}"] = chuongtrinh
                };

                for (int i = 1; i <= 5; i++)
                {
                    map[$"{{{{hoten{i}}}}}"] = hoten[i];
                    map[$"{{{{sotien{i}}}}}"] = sotien[i];
                    map[$"{{{{phuongan{i}}}}}"] = phuongan[i];
                    map[$"{{{{thoihanvay{i}}}}}"] = thoihan[i];
                    map[$"{{{{doituong{i}}}}}"] = doituong[i];
                }

                // Add entered people to the grid (selectedCustomers)
                try
                {
                    for (int i = 1; i <=5; i++)
                    {
                        if (!string.IsNullOrWhiteSpace(hoten[i]))
                        {
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
                    try { dataGridView1.DataSource = null; dataGridView1.DataSource = selectedCustomers; } catch { }
                }
                catch { }

                await Task.Run(() => ExportWord(map, hoten, sotien, totruong, xa));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xuất Word: " + ex.Message);
            }
        }

        // ================= EXPORT WORD =================
        private void ExportWord(Dictionary<string, string> map, string[] hoten, string[] sotien, string totruong, string xa)
        {
            string template = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Templates",
                "03 DS GROUP.docx"
            );

            // If not in Templates under output, try recursive search under base directory
            if (!File.Exists(template))
            {
                try
                {
                    var asm = Assembly.GetExecutingAssembly();
                    var found = Directory.EnumerateFiles(AppDomain.CurrentDomain.BaseDirectory, "03 DS GROUP.docx", SearchOption.AllDirectories).FirstOrDefault();
                    if (!string.IsNullOrEmpty(found)) template = found;

                    // If still not found, search upward from the BaseDirectory (project source root may be above bin folder)
                    if (!File.Exists(template))
                    {
                        var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
                        for (int up = 0; up < 6 && dir.Parent != null; up++)
                        {
                            dir = dir.Parent;
                            var upCandidate = Path.Combine(dir.FullName, "03 DS GROUP.docx");
                            if (File.Exists(upCandidate)) { template = upCandidate; break; }
                        }
                    }

                    // fallback: try embedded resources
                    if (!File.Exists(template))
                    {
                        var res = asm.GetManifestResourceNames().FirstOrDefault(n => n.EndsWith("03 DS GROUP.docx", StringComparison.OrdinalIgnoreCase));
                        if (res != null)
                        {
                            var temp = Path.Combine(Path.GetTempPath(), "template_03ds_" + Guid.NewGuid().ToString("N") + ".docx");
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

            // fallback: prompt user to choose file (must run on UI thread)
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
                            dlg.Title = "Chọn file mẫu 03 DS GROUP.docx";
                            if (dlg.ShowDialog(this) == DialogResult.OK) userPick = dlg.FileName;
                        }
                    }));

                    if (!string.IsNullOrEmpty(userPick) && File.Exists(userPick)) template = userPick;
                }
                catch { }
            }

            if (!File.Exists(template))
                throw new FileNotFoundException("Không tìm thấy file mẫu Word.");

            string folder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "Hồ sơ NHCS",
                Safe(totruong)
            );
            Directory.CreateDirectory(folder);

            string output = Path.Combine(folder, Safe(totruong) + "-" + Safe(xa) + "-" + DateTime.Now.ToString("MM-yyyy") + ".docx");
            File.Copy(template, output, true);

            using (var doc = WordprocessingDocument.Open(output, true))
            {
                // Remove rows related to empty customers first so they don't leave blank rows
                RemoveUnusedRows(doc, hoten);

                // Replace placeholders including those split across runs
                ReplacePlaceholdersAcrossRuns(doc, map);

                // Fallback simple replacement for any remaining placeholders entirely inside a single Text
                ReplacePlaceholdersPreserveFormatting(doc, map);

                // Fill total from sotien placeholders
                FillTongTien(doc, sotien);

                doc.MainDocumentPart.Document.Save();
            }

            System.Diagnostics.Process.Start(output);
        }

        // Replace placeholders by iterating Text elements in main and related parts to preserve formatting
        private void ReplacePlaceholdersPreserveFormatting(WordprocessingDocument doc, Dictionary<string, string> map)
        {
            if (doc?.MainDocumentPart == null) return;
            var mainPart = doc.MainDocumentPart;

            var parts = new List<OpenXmlPart> { mainPart };
            parts.AddRange(mainPart.HeaderParts);
            parts.AddRange(mainPart.FooterParts);
            if (mainPart.FootnotesPart != null) parts.Add(mainPart.FootnotesPart);
            if (mainPart.EndnotesPart != null) parts.Add(mainPart.EndnotesPart);
            if (mainPart.WordprocessingCommentsPart != null) parts.Add(mainPart.WordprocessingCommentsPart);

            foreach (var part in parts.Distinct())
            {
                try
                {
                    var texts = part.RootElement.Descendants<Text>();
                    foreach (var t in texts)
                    {
                        if (string.IsNullOrEmpty(t.Text)) continue;
                        string newText = t.Text;
                        foreach (var kv in map)
                        {
                            if (string.IsNullOrEmpty(kv.Key)) continue;
                            if (newText.IndexOf(kv.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                newText = newText.Replace(kv.Key, kv.Value ?? "");
                            }
                        }
                        if (newText != t.Text) t.Text = newText;
                    }
                    part.RootElement.Save();
                }
                catch { }
            }
        }

        // Replace placeholders that may be split across multiple Text nodes (runs)
        private void ReplacePlaceholdersAcrossRuns(WordprocessingDocument doc, Dictionary<string, string> map)
        {
            if (doc?.MainDocumentPart == null) return;

            var parts = new List<OpenXmlPart> { doc.MainDocumentPart };
            parts.AddRange(doc.MainDocumentPart.HeaderParts);
            parts.AddRange(doc.MainDocumentPart.FooterParts);
            if (doc.MainDocumentPart.FootnotesPart != null) parts.Add(doc.MainDocumentPart.FootnotesPart);
            if (doc.MainDocumentPart.EndnotesPart != null) parts.Add(doc.MainDocumentPart.EndnotesPart);
            if (doc.MainDocumentPart.WordprocessingCommentsPart != null) parts.Add(doc.MainDocumentPart.WordprocessingCommentsPart);

            // Order placeholders by descending length to avoid partial matches
            var placeholders = map.Keys.Where(k => !string.IsNullOrEmpty(k)).OrderByDescending(k => k.Length).ToList();

            foreach (var part in parts.Distinct())
            {
                try
                {
                    // Group Text nodes by their containing paragraph to keep replacements local
                    var paragraphs = part.RootElement.Descendants<Paragraph>().ToList();
                    bool partChanged = false;
                    foreach (var p in paragraphs)
                    {
                        var texts = p.Descendants<Text>().ToList();
                        if (texts.Count == 0) continue;

                        // Build full string for the paragraph
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

                                int matchEndPos = idx + ph.Length; // exclusive
                                // locate end text node
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

                                // compute left and right fragments
                                var left = (texts[startNode].Text ?? "").Substring(0, startOffset);
                                var right = (texts[endNode].Text ?? "").Substring(endOffsetExclusive);

                                // set start node text to left + replacement + right
                                texts[startNode].Text = left + (replacement ?? "") + right;

                                // clear intermediate nodes
                                for (int k = startNode + 1; k <= endNode; k++)
                                {
                                    if (k == startNode + 1)
                                    {
                                        // if endNode == startNode+1 and right already moved into start node, clear this
                                        texts[k].Text = "";
                                    }
                                    else
                                    {
                                        texts[k].Text = "";
                                    }
                                }

                                changed = true;

                                // rebuild full and continue after the replacement
                                full = string.Concat(texts.Select(t => t.Text ?? ""));
                                searchStart = (left + replacement).Length + texts.Take(startNode).Sum(t => (t.Text ?? "").Length);
                            }
                        }

                        if (changed) partChanged = true;
                    }
                    if (partChanged)
                    {
                        try { part.RootElement.Save(); } catch { }
                    }
                }
                catch { }
            }
        }

        // ================= XÓA DÒNG NẾU CHỨA PLACEHOLDER CỦA KHÁCH HÀNG KHÔNG CÓ TÊN =================
        private void RemoveUnusedRows(WordprocessingDocument doc, string[] hoten)
        {
            if (doc?.MainDocumentPart == null) return;
            var rows = doc.MainDocumentPart.Document
                .Descendants<TableRow>()
                .ToList();

            foreach (var row in rows)
            {
                var rowText = string.Concat(row.Descendants<Text>().Select(t => t.Text ?? ""));
                for (int i = 1; i <= 5; i++)
                {
                    if (string.IsNullOrWhiteSpace(hoten[i]) && rowText.IndexOf($"{{{{hoten{i}}}}}", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        row.Remove();
                        break;
                    }
                }
            }
        }

        // ================= TÍNH {{cong}} =================
        private void FillTongTien(WordprocessingDocument doc, string[] sotien)
        {
            long tong = 0;
            for (int i = 1; i <= 5; i++)
            {
                var digits = new string((sotien[i] ?? "").Where(char.IsDigit).ToArray());
                if (long.TryParse(digits, out long v)) tong += v;
            }

            ReplacePlaceholdersPreserveFormatting(doc, new Dictionary<string, string>
            {
                ["{{cong}}"] = tong > 0 ? tong.ToString("N0", new CultureInfo("vi-VN")) : ""
            });
        }

        // ================= TIỆN ÍCH =================
        private string Clean(string s)
        {
            return (s ?? "").Trim().Replace("{", "").Replace("}", "");
        }

        private string Money(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            var d = new string(s.Where(char.IsDigit).ToArray());
            if (long.TryParse(d, out long v))
                return v.ToString("N0", new CultureInfo("vi-VN"));
            return s;
        }

        // Input handlers for money fields (combo boxes) - restrict to digits and format with thousands separator
        private void CbMoney_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void CbMoney_TextChanged(object sender, EventArgs e)
        {
            try
            {
                var cb = sender as ComboBox;
                if (cb == null) return;
                var txt = cb.Text ?? "";
                var digits = new string(txt.Where(char.IsDigit).ToArray());
                if (string.IsNullOrEmpty(digits))
                {
                    if (!string.IsNullOrEmpty(txt)) cb.Text = "";
                    return;
                }
                if (long.TryParse(digits, out var v))
                {
                    var formatted = v.ToString("N0", new CultureInfo("vi-VN"));
                    if (cb.Text != formatted)
                    {
                        cb.Text = formatted;
                        try { cb.SelectionStart = cb.Text.Length; cb.SelectionLength = 0; } catch { }
                    }
                }
            }
            catch { }
        }

        // Allow only letters and whitespace (and a few punctuation chars) in specified text fields; block digits
        private void TextLettersOnly_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsControl(e.KeyChar)) return;
            if (char.IsWhiteSpace(e.KeyChar)) return;
            if (char.IsLetter(e.KeyChar)) return;
            // allow common name punctuation
            if (e.KeyChar == '-' || e.KeyChar == '\'' || e.KeyChar == '.' ) return;
            e.Handled = true;
        }

        // Designer-referenced stub handlers (no-op)
        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            try { /* intentionally left blank to satisfy designer */ } catch { }
        }

        private void label18_Click(object sender, EventArgs e)
        {
            try { /* intentionally left blank to satisfy designer */ } catch { }
        }

        // Remove selected customers from grid
        private void BtnXoa_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1 == null) return;
                var toRemove = new List<Customer>();
                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    try { var item = row.DataBoundItem as Customer; if (item != null) toRemove.Add(item); } catch { }
                }
                foreach (var r in toRemove)
                {
                    selectedCustomers.Remove(r);
                }
                // refresh grid
                try { dataGridView1.DataSource = null; dataGridView1.DataSource = selectedCustomers; } catch { }
            }
            catch { }
        }

        // ================= TIỆN ÍCH PHỤ =================
        private string Safe(string s)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
                s = s.Replace(c, '_');
            return s.Trim();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }
    }
}
