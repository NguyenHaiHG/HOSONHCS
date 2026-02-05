using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace HOSONHCS
{
    // Helper class to provide xinman.json editing UI logic without modifying Form1 heavily
    public class XinManEditor
    {
        private DataGridView dgv;
        private TextBox txtUsername;
        private TextBox txtPassword;
        private Button btnLogin;
        private Button btnSave; // created if not provided (disabled now)
        private TextBox txtSearch;
        private Button btnNextSearch;
        private Button btnPrevSearch;
        private BindingSource source = new BindingSource();
        private DataTable table;
        private XinManModel model;
        private string xinmanPath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "xinman.json");

        // simple hardcoded credential (per user request)
        private const string validUser = "haihg";
        private const string validPass = "Haihg23";

        // search state
        private List<int> searchMatches = new List<int>();
        private int currentMatchIndex = -1;

        public void AttachControls(DataGridView dgv, TextBox txtUsername, TextBox txtPassword, Button btnLogin, TextBox txtSearch, Button btnSave = null)
        {
            this.dgv = dgv ?? throw new ArgumentNullException(nameof(dgv));
            this.txtUsername = txtUsername;
            this.txtPassword = txtPassword;
            this.btnLogin = btnLogin;
            this.txtSearch = txtSearch;
            this.btnSave = btnSave; // keep reference but do not auto-create

            if (this.btnLogin != null) this.btnLogin.Click += BtnLogin_Click;
            if (this.txtSearch != null) this.txtSearch.TextChanged += TxtSearch_TextChanged;

            // REMOVE suspicious runtime-created save button if present in same parent
            try
            {
                if (this.btnLogin != null && this.btnLogin.Parent != null)
                {
                    var parent = this.btnLogin.Parent;
                    var toRemove = parent.Controls.OfType<Button>().FirstOrDefault(b => b != this.btnLogin && b != this.btnSave && string.Equals(b.Text, "Lưu xinman.json", StringComparison.Ordinal));
                    if (toRemove != null)
                    {
                        try { parent.Controls.Remove(toRemove); toRemove.Dispose(); } catch { }
                    }
                }
            }
            catch { }

            // Detect if form already contains user-provided Next/Prev buttons (common names)
            Control form = dgv?.FindForm() ?? btnLogin?.FindForm();
            Button userNext = null, userPrev = null;
            if (form != null)
            {
                // check common name variants
                userNext = form.Controls.Find("btnNext", true).FirstOrDefault() as Button
                           ?? form.Controls.Find("btnNextSearch", true).FirstOrDefault() as Button
                           ?? form.Controls.Find("btnNextBtn", true).FirstOrDefault() as Button
                           ?? form.Controls.Find("btnNextButton", true).FirstOrDefault() as Button;

                userPrev = form.Controls.Find("btnPre", true).FirstOrDefault() as Button
                           ?? form.Controls.Find("btnPrev", true).FirstOrDefault() as Button
                           ?? form.Controls.Find("btnPrevSearch", true).FirstOrDefault() as Button
                           ?? form.Controls.Find("btnPreSave", true).FirstOrDefault() as Button;
            }

            // If user created Next/Prev buttons, wire them; we'll not create duplicates
            try
            {
                if (userNext != null)
                {
                    try { userNext.Click -= BtnNextSearch_Click; } catch { }
                    userNext.Click += BtnNextSearch_Click;
                    btnNextSearch = userNext;
                }
                if (userPrev != null)
                {
                    try { userPrev.Click -= BtnPrevSearch_Click; } catch { }
                    userPrev.Click += BtnPrevSearch_Click;
                    btnPrevSearch = userPrev;
                }
            }
            catch { }

            // Create dynamic next/prev search buttons near txtSearch only if the form does not already provide them
            try
            {
                if (this.txtSearch != null && this.txtSearch.Parent != null && (btnNextSearch == null || btnPrevSearch == null))
                {
                    var parent = this.txtSearch.Parent;
                    // try find existing dynamic buttons placed previously in same parent
                    var existingPrev = parent.Controls.OfType<Button>().FirstOrDefault(b => b != null && b.Tag != null && b.Tag.ToString() == "XinManPrev");
                    var existingNext = parent.Controls.OfType<Button>().FirstOrDefault(b => b != null && b.Tag != null && b.Tag.ToString() == "XinManNext");

                    if (btnPrevSearch == null && existingPrev != null) btnPrevSearch = existingPrev;
                    if (btnNextSearch == null && existingNext != null) btnNextSearch = existingNext;

                    // only create missing ones and only if user buttons not present
                    if (btnPrevSearch == null)
                    {
                        btnPrevSearch = new Button() { Width = 24, Height = this.txtSearch.Height, Text = "<", Tag = "XinManPrev" };
                        btnPrevSearch.Left = this.txtSearch.Left + this.txtSearch.Width + 4;
                        btnPrevSearch.Top = this.txtSearch.Top;
                        parent.Controls.Add(btnPrevSearch);
                        btnPrevSearch.Click += BtnPrevSearch_Click;
                    }

                    if (btnNextSearch == null)
                    {
                        btnNextSearch = new Button() { Width = 24, Height = this.txtSearch.Height, Text = ">", Tag = "XinManNext" };
                        btnNextSearch.Left = this.txtSearch.Left + this.txtSearch.Width + 4 + (btnPrevSearch?.Width ?? 24) + 4;
                        btnNextSearch.Top = this.txtSearch.Top;
                        parent.Controls.Add(btnNextSearch);
                        btnNextSearch.Click += BtnNextSearch_Click;
                    }
                }
            }
            catch { }

            if (this.btnSave != null) this.btnSave.Click += BtnSave_Click; // if designer provided, keep it

            // initial state: load model and populate grid, but keep read-only until login
            LoadModel();
            PopulateGridFromModel();
            SetEditable(false);
        }

        private void BtnPrevSearch_Click(object sender, EventArgs e)
        {
            if (searchMatches == null || searchMatches.Count == 0) return;
            currentMatchIndex = (currentMatchIndex - 1 + searchMatches.Count) % searchMatches.Count;
            SelectMatchAtCurrentIndex();
        }

        private void BtnNextSearch_Click(object sender, EventArgs e)
        {
            if (searchMatches == null || searchMatches.Count == 0) return;
            currentMatchIndex = (currentMatchIndex + 1) % searchMatches.Count;
            SelectMatchAtCurrentIndex();
        }

        private void SelectMatchAtCurrentIndex()
        {
            try
            {
                if (currentMatchIndex < 0 || currentMatchIndex >= searchMatches.Count) return;
                var rowIndex = searchMatches[currentMatchIndex];
                if (dgv == null) return;
                if (dgv.IsHandleCreated && dgv.InvokeRequired)
                {
                    dgv.Invoke((Action)(() => { TrySelectRow(rowIndex); }));
                }
                else
                {
                    TrySelectRow(rowIndex);
                }
            }
            catch { }
        }

        private void TrySelectRow(int i)
        {
            try
            {
                if (i < 0 || i >= dgv.Rows.Count) return;
                dgv.ClearSelection();
                try { dgv.FirstDisplayedScrollingRowIndex = i; } catch { }
                dgv.Rows[i].Selected = true;
            }
            catch { }
        }

        private void TxtSearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                var q = txtSearch?.Text?.Trim();
                searchMatches.Clear(); currentMatchIndex = -1;
                if (string.IsNullOrEmpty(q))
                {
                    // clear selection
                    try { if (dgv != null) { if (dgv.InvokeRequired) dgv.Invoke((Action)(() => dgv.ClearSelection())); else dgv.ClearSelection(); } } catch { }
                    return;
                }

                q = q.ToLowerInvariant();

                // find all matching row indices
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    var row = table.Rows[i];
                    bool match = false;
                    foreach (DataColumn col in table.Columns)
                    {
                        var s = (row[col] ?? "").ToString();
                        if (!string.IsNullOrEmpty(s) && s.ToLowerInvariant().Contains(q)) { match = true; break; }
                    }
                    if (match) searchMatches.Add(i);
                }

                if (searchMatches.Count > 0)
                {
                    currentMatchIndex = 0;
                    SelectMatchAtCurrentIndex();
                }
            }
            catch { }
        }

        private void BtnLogin_Click(object sender, EventArgs e)
        {
            var u = txtUsername?.Text?.Trim() ?? "";
            var p = txtPassword?.Text ?? "";
            if (string.Equals(u, validUser, StringComparison.OrdinalIgnoreCase) && p == validPass)
            {
                MessageBox.Show("Đăng nhập thành công.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
                SetEditable(true);
            }
            else
            {
                MessageBox.Show("Tên đăng nhập hoặc mật khẩu không đúng.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SetEditable(false);
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                // Commit any pending edits in the DataGridView before saving
                try
                {
                    if (dgv != null)
                    {
                        dgv.EndEdit();
                        dgv.CurrentCell = null; // Force commit of the current cell edit
                    }
                    if (source != null)
                    {
                        source.EndEdit();
                    }
                }
                catch { }

                // rebuild model from table
                var newModel = new XinManModel();
                newModel.pgd = model?.pgd ?? "";

                var communes = new Dictionary<string, Commune>(StringComparer.OrdinalIgnoreCase);

                foreach (DataRow r in table.Rows)
                {
                    var comm = (r["Commune"] ?? "").ToString().Trim();
                    var assoc = (r["Association"] ?? "").ToString().Trim();
                    var village = (r["Village"] ?? "").ToString().Trim();
                    var groupsRaw = (r["Groups"] ?? "").ToString().Trim();

                    if (string.IsNullOrEmpty(comm)) continue;

                    if (!communes.TryGetValue(comm, out var comObj))
                    {
                        comObj = new Commune { name = comm, associations = new List<Association>(), villages = new List<Village>() };
                        communes[comm] = comObj;
                    }

                    if (!string.IsNullOrEmpty(assoc))
                    {
                        var assocObj = comObj.associations.FirstOrDefault(a => string.Equals(a.name, assoc, StringComparison.OrdinalIgnoreCase));
                        if (assocObj == null)
                        {
                            assocObj = new Association { name = assoc, code = "", villages = new List<Village>(), managedVillages = new List<string>() };
                            comObj.associations.Add(assocObj);
                        }

                        if (!string.IsNullOrEmpty(village))
                        {
                            var villageObj = assocObj.villages.FirstOrDefault(v => string.Equals(v.name, village, StringComparison.OrdinalIgnoreCase));
                            if (villageObj == null)
                            {
                                villageObj = new Village { name = village, groups = new List<string>() };
                                assocObj.villages.Add(villageObj);
                            }

                            // parse groups by ; or ,
                            var groups = SplitGroups(groupsRaw);
                            foreach (var g in groups)
                                if (!villageObj.groups.Contains(g, StringComparer.OrdinalIgnoreCase)) villageObj.groups.Add(g);
                        }
                    }
                    else
                    {
                        // no association: treat as commune-level village
                        if (!string.IsNullOrEmpty(village))
                        {
                            var villageObj = comObj.villages.FirstOrDefault(v => string.Equals(v.name, village, StringComparison.OrdinalIgnoreCase));
                            if (villageObj == null)
                            {
                                villageObj = new Village { name = village, groups = new List<string>() };
                                comObj.villages.Add(villageObj);
                            }
                            var groups = SplitGroups(groupsRaw);
                            foreach (var g in groups)
                                if (!villageObj.groups.Contains(g, StringComparer.OrdinalIgnoreCase)) villageObj.groups.Add(g);
                        }
                    }
                }

                newModel.communes = communes.Values.ToList();

                // write backup and atomic save
                var json = JsonConvert.SerializeObject(newModel, Formatting.Indented);
                var temp = xinmanPath + ".tmp";
                File.WriteAllText(temp, json, Encoding.UTF8);

                // backup
                try
                {
                    if (File.Exists(xinmanPath))
                    {
                        var bak = xinmanPath + ".bak." + DateTime.Now.ToString("yyyyMMddHHmmss");
                        File.Copy(xinmanPath, bak);
                    }
                }
                catch { }

                // replace
                File.Copy(temp, xinmanPath, true);
                try { File.Delete(temp); } catch { }

                MessageBox.Show("Lưu xinman.json thành công.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // reload model to ensure UI reflects any normalization
                LoadModel();
                PopulateGridFromModel();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi lưu xinman.json: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private IEnumerable<string> SplitGroups(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) yield break;
            var parts = raw.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var p in parts)
            {
                var t = p.Trim();
                if (!string.IsNullOrEmpty(t)) yield return t;
            }
        }

        public void SetEditable(bool enabled)
        {
            try
            {
                if (dgv == null) return;
                dgv.ReadOnly = !enabled;
                if (btnSave != null) btnSave.Enabled = enabled;
                // visual cue
                dgv.BackgroundColor = enabled ? System.Drawing.SystemColors.Window : System.Drawing.SystemColors.Control;
            }
            catch { }
        }

        // Logout method to disable editing and clear credentials
        public void Logout()
        {
            try
            {
                // Disable editing
                SetEditable(false);

                // Clear username and password fields
                try { if (txtUsername != null) txtUsername.Text = ""; } catch { }
                try { if (txtPassword != null) txtPassword.Text = ""; } catch { }

                // Clear search
                try { if (txtSearch != null) txtSearch.Text = ""; } catch { }

                // Reload model to discard any unsaved changes and restore original state
                LoadModel();
                PopulateGridFromModel();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi đăng xuất: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadModel()
        {
            model = null;
            try
            {
                if (File.Exists(xinmanPath))
                {
                    var json = File.ReadAllText(xinmanPath, Encoding.UTF8);
                    model = JsonConvert.DeserializeObject<XinManModel>(json);
                }
            }
            catch { model = null; }

            if (model == null) model = new XinManModel { pgd = "", communes = new List<Commune>() };
        }

        private void PopulateGridFromModel()
        {
            try
            {
                table = new DataTable();
                table.Columns.Add("Commune");
                table.Columns.Add("Association");
                table.Columns.Add("Village");
                table.Columns.Add("Groups");

                foreach (var c in model.communes ?? Enumerable.Empty<Commune>())
                {
                    if (c.associations != null)
                    {
                        foreach (var a in c.associations)
                        {
                            if (a.villages != null)
                            {
                                foreach (var v in a.villages)
                                {
                                    var groups = v.groups != null ? string.Join("; ", v.groups) : "";
                                    table.Rows.Add(c.name ?? "", a.name ?? "", v.name ?? "", groups);
                                }
                            }
                            // if association has managed villages as names
                            if (a.managedVillages != null && a.managedVillages.Count > 0)
                            {
                                foreach (var vn in a.managedVillages)
                                    table.Rows.Add(c.name ?? "", a.name ?? "", vn ?? "", "");
                            }
                        }
                    }
                    if (c.villages != null)
                    {
                        foreach (var v in c.villages)
                        {
                            var groups = v.groups != null ? string.Join("; ", v.groups) : "";
                            table.Rows.Add(c.name ?? "", "", v.name ?? "", groups);
                        }
                    }
                }

                source.DataSource = table;

                // Safely assign DataSource to DataGridView without invoking before handle exists
                try
                {
                    if (dgv != null)
                    {
                        if (dgv.IsHandleCreated)
                        {
                            if (dgv.InvokeRequired)
                                dgv.Invoke((Action)(() => { dgv.DataSource = source; }));
                            else
                                dgv.DataSource = source;
                        }
                        else
                        {
                            // If handle not yet created, setting DataSource directly is fine when on UI thread
                            dgv.DataSource = source;
                        }
                    }
                }
                catch
                {
                    // fallback: attempt direct assignment
                    try { if (dgv != null) dgv.DataSource = source; } catch { }
                }

                // adjust columns
                try
                {
                    if (dgv != null && dgv.Columns.Count > 0)
                    {
                        if (dgv.IsHandleCreated && dgv.InvokeRequired)
                        {
                            dgv.Invoke((Action)(() =>
                            {
                                try
                                {
                                    dgv.Columns["Commune"].Width = 180;
                                    dgv.Columns["Association"].Width = 180;
                                    dgv.Columns["Village"].Width = 180;
                                    dgv.Columns["Groups"].Width = 250;
                                }
                                catch { }
                            }));
                        }
                        else
                        {
                            try
                            {
                                dgv.Columns["Commune"].Width = 180;
                                dgv.Columns["Association"].Width = 180;
                                dgv.Columns["Village"].Width = 180;
                                dgv.Columns["Groups"].Width = 250;
                            }
                            catch { }
                        }
                    }
                }
                catch { }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi nạp dữ liệu vào grid: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
