using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace HOSONHCS
{
    public partial class Form1
    {
        // ========== FIELDS GHI CHÚ ==========
        private ListView    lvNotes;
        private RichTextBox rtbNoteContent;
        private TextBox     txtNoteTitle;
        private TextBox     txtNoteSearch;
        private ComboBox    cbNoteCategory;
        private Label       lblNoteInfo;
        private Button      btnNoteNew;
        private Button      btnNoteSave;
        private Button      btnNoteDelete;
        private Button      btnNotePin;
        private SplitContainer splitNotes;

        private List<NoteItem> notesList         = new List<NoteItem>();
        private NoteItem       currentNote       = null;
        private bool           suppressNoteSearch = false;

        private string GetNotesFolderPath()
            => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Notes");

        // =====================================================================
        // KHỞI TẠO UI — Gọi từ InitializeApp()
        // =====================================================================
        private void InitializeGhiChuTab()
        {
            if (tabPage5 == null) return;

            tabPage5.Controls.Clear();
            tabPage5.BackColor = UIStyler.BgMain;
            tabPage5.Padding   = new Padding(0);

            // ── 1. TOOLBAR (Dock Top) ─────────────────────────────────────────
            var pnlToolbar = new Panel
            {
                Dock      = DockStyle.Top,
                Height    = 46,
                BackColor = UIStyler.BgPanel,
                Padding   = new Padding(6, 7, 6, 5)
            };

            txtNoteSearch = new TextBox
            {
                Location    = new Point(6, 10),
                Width       = 210,
                Height      = 26,
                Font        = new Font("Segoe UI", 10.5f),
                ForeColor   = UIStyler.TextHint,
                BackColor   = UIStyler.BgInput,
                BorderStyle = BorderStyle.FixedSingle,
                Text        = "Tìm kiếm tiêu đề..."
            };
            txtNoteSearch.Enter += (s, e) =>
            {
                if (txtNoteSearch.Text == "Tìm kiếm tiêu đề...")
                { txtNoteSearch.Text = ""; txtNoteSearch.ForeColor = UIStyler.TextMain; }
            };
            txtNoteSearch.Leave += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(txtNoteSearch.Text))
                { txtNoteSearch.Text = "Tìm kiếm tiêu đề..."; txtNoteSearch.ForeColor = UIStyler.TextHint; }
            };
            txtNoteSearch.TextChanged += TxtNoteSearch_TextChanged;

            btnNoteNew    = MakeToolBtn("+ Mới",  224, UIStyler.BtnGreen);
            btnNoteSave   = MakeToolBtn("Lưu",    314, UIStyler.BtnBlue);
            btnNotePin    = MakeToolBtn("Ghim",   404, UIStyler.BtnOrange);
            btnNoteDelete = MakeToolBtn("Xóa",    494, UIStyler.BtnRed);

            btnNoteNew.Click    += BtnNoteNew_Click;
            btnNoteSave.Click   += BtnNoteSave_Click;
            btnNotePin.Click    += BtnNotePin_Click;
            btnNoteDelete.Click += BtnNoteDelete_Click;

            pnlToolbar.Controls.AddRange(new Control[]
                { txtNoteSearch, btnNoteNew, btnNoteSave, btnNotePin, btnNoteDelete });

            // ── 2. PANEL TRÁI: ListView ───────────────────────────────────────
            var pnlLeft = new Panel
            {
                Dock      = DockStyle.Left,
                Width     = 270,
                BackColor = UIStyler.BgPanel
            };

            var lblListHdr = new Label
            {
                Text      = "  DANH SÁCH GHI CHÚ",
                Dock      = DockStyle.Top,
                Height    = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                Font      = new Font("Segoe UI", 10f, FontStyle.Bold),
                BackColor = UIStyler.BgCard,
                ForeColor = UIStyler.Primary
            };

            lvNotes = new ListView
            {
                Dock          = DockStyle.Fill,
                View          = View.Details,
                FullRowSelect = true,
                GridLines     = true,
                HideSelection = false,
                Font          = new Font("Segoe UI", 10.5f),
                BorderStyle   = BorderStyle.None,
                BackColor     = UIStyler.BgPanel,
                ForeColor     = UIStyler.TextMain
            };
            lvNotes.Columns.Add("Tiêu đề", 182);
            lvNotes.Columns.Add("Ngày",     62);
            lvNotes.SelectedIndexChanged += LvNotes_SelectedIndexChanged;

            var sepLine = new Panel
            {
                Dock      = DockStyle.Right,
                Width     = 1,
                BackColor = UIStyler.BorderColor
            };

            pnlLeft.Controls.Add(lvNotes);
            pnlLeft.Controls.Add(lblListHdr);
            pnlLeft.Controls.Add(sepLine);

            // ── 3. PANEL PHẢI: Editor ─────────────────────────────────────────
            var pnlRight = new Panel
            {
                Dock      = DockStyle.Fill,
                BackColor = UIStyler.BgMain,
                Padding   = new Padding(0)
            };

            var pnlHeader = new Panel
            {
                Dock      = DockStyle.Top,
                Height    = 72,
                BackColor = UIStyler.BgPanel,
                Padding   = new Padding(8, 6, 8, 4)
            };

            var lblTL = new Label
            {
                Text      = "Tiêu đề:",
                Location  = new Point(8, 12),
                AutoSize  = true,
                Font      = new Font("Segoe UI", 10f, FontStyle.Bold),
                ForeColor = UIStyler.TextMain,
                BackColor = Color.Transparent
            };
            txtNoteTitle = new TextBox
            {
                Location    = new Point(72, 8),
                Width       = 370,
                Height      = 28,
                Font        = new Font("Segoe UI", 12f, FontStyle.Bold),
                BackColor   = UIStyler.BgInput,
                ForeColor   = UIStyler.TextMain,
                BorderStyle = BorderStyle.FixedSingle
            };

            var lblCL = new Label
            {
                Text      = "Danh mục:",
                Location  = new Point(455, 12),
                AutoSize  = true,
                Font      = new Font("Segoe UI", 10f),
                ForeColor = UIStyler.TextMain,
                BackColor = Color.Transparent
            };
            cbNoteCategory = new ComboBox
            {
                Location      = new Point(540, 8),
                Width         = 160,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font          = new Font("Segoe UI", 10f),
                BackColor     = UIStyler.BgInput,
                ForeColor     = UIStyler.TextMain,
                FlatStyle     = FlatStyle.Flat
            };
            cbNoteCategory.Items.AddRange(new object[]
                { "Chung", "Hẹn gặp", "Cần làm", "Hộ nghèo", "SXKD", "GQVL", "Quan trọng", "Lưu ý" });
            cbNoteCategory.SelectedIndex = 0;

            lblNoteInfo = new Label
            {
                Location  = new Point(8, 44),
                Width     = 720,
                Height    = 20,
                Font      = new Font("Segoe UI", 9.5f),
                ForeColor = UIStyler.TextSub,
                BackColor = Color.Transparent,
                Text      = "Bấm '+ Mới' để tạo ghi chú hoặc chọn từ danh sách bên trái"
            };

            pnlHeader.Controls.AddRange(new Control[]
                { lblTL, txtNoteTitle, lblCL, cbNoteCategory, lblNoteInfo });

            // Separator dưới header
            var sepHeader = new Panel
            {
                Dock      = DockStyle.Top,
                Height    = 1,
                BackColor = UIStyler.BorderColor
            };

            rtbNoteContent = new RichTextBox
            {
                Dock        = DockStyle.Fill,
                Font        = new Font("Segoe UI", 12.5f),
                BorderStyle = BorderStyle.None,
                BackColor   = UIStyler.BgCard,
                ForeColor   = UIStyler.TextMain,
                ScrollBars  = RichTextBoxScrollBars.Vertical,
                Padding     = new Padding(10)
            };

            // Thứ tự Add quan trọng: Fill trước, Top sau (WinForms dock ngược)
            pnlRight.Controls.Add(rtbNoteContent);
            pnlRight.Controls.Add(sepHeader);
            pnlRight.Controls.Add(pnlHeader);

            // ── 4. Gắn tất cả vào tabPage5 ───────────────────────────────────
            // Fill trước, Left và Top sau
            tabPage5.Controls.Add(pnlRight);
            tabPage5.Controls.Add(pnlLeft);
            tabPage5.Controls.Add(pnlToolbar);

            // ── 5. Load dữ liệu ──────────────────────────────────────────────
            LoadNotesFromFiles();
            RefreshNotesList();
        }

        // ── Helper tạo nút toolbar ──
        private Button MakeToolBtn(string text, int x, Color color)
        {
            var btn = new Button
            {
                Text      = text,
                Location  = new Point(x, 8),
                Width     = 82,
                Height    = 28,
                BackColor = color,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 10f, FontStyle.Bold),
                Cursor    = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        // =====================================================================
        // LOAD / SAVE / DELETE
        // =====================================================================
        private void LoadNotesFromFiles()
        {
            notesList.Clear();
            try
            {
                var folder = GetNotesFolderPath();
                if (!Directory.Exists(folder)) return;
                foreach (var file in Directory.EnumerateFiles(folder, "note_*.json"))
                {
                    try
                    {
                        var note = JsonConvert.DeserializeObject<NoteItem>(
                            File.ReadAllText(file, Encoding.UTF8));
                        if (note != null) notesList.Add(note);
                    }
                    catch { }
                }
                SortNotes();
            }
            catch { }
        }

        private void SaveNoteToFile(NoteItem note)
        {
            try
            {
                var folder = GetNotesFolderPath();
                if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
                File.WriteAllText(
                    Path.Combine(folder, "note_" + note.Id + ".json"),
                    JsonConvert.SerializeObject(note, Formatting.Indented),
                    Encoding.UTF8);
            }
            catch { }
        }

        private void DeleteNoteFile(NoteItem note)
        {
            try
            {
                var path = Path.Combine(GetNotesFolderPath(), "note_" + note.Id + ".json");
                if (File.Exists(path)) File.Delete(path);
            }
            catch { }
        }

        private void SortNotes()
        {
            notesList = notesList
                .OrderByDescending(n => n.IsPinned)
                .ThenByDescending(n => n.UpdatedAt)
                .ToList();
        }

        // =====================================================================
        // REFRESH LISTVIEW
        // =====================================================================
        private void RefreshNotesList(string filter = "")
        {
            if (lvNotes == null) return;
            try
            {
                lvNotes.BeginUpdate();
                lvNotes.Items.Clear();

                var list = string.IsNullOrWhiteSpace(filter)
                    ? notesList
                    : notesList.Where(n =>
                        (n.Title    ?? "").IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        (n.Content  ?? "").IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        (n.Category ?? "").IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0
                      ).ToList();

                foreach (var note in list)
                {
                    var item = new ListViewItem(
                        (note.IsPinned ? "[P] " : "") + (note.Title ?? "(Chưa đặt tên)"));
                    item.SubItems.Add(note.UpdatedAt.ToString("dd/MM"));
                    item.Tag = note;

                    // Nền tối + chữ sáng để nhìn rõ trên teal theme
                    switch (note.Category)
                    {
                        case "Quan trọng": item.BackColor = Color.FromArgb(140, 35, 50);  break; // đỏ tối
                        case "Hẹn gặp":   item.BackColor = Color.FromArgb(130, 100, 20); break; // vàng tối
                        case "Cần làm":   item.BackColor = Color.FromArgb(25,  80, 140); break; // xanh dương tối
                        case "Hộ nghèo":  item.BackColor = Color.FromArgb(25,  100, 45); break; // xanh lá tối
                        default:          item.BackColor = UIStyler.BgCard;              break; // teal card
                    }
                    item.ForeColor = UIStyler.TextMain; // chữ luôn sáng
                    lvNotes.Items.Add(item);
                }
            }
            finally { lvNotes.EndUpdate(); }
        }

        // =====================================================================
        // EVENT HANDLERS
        // =====================================================================
        private void TxtNoteSearch_TextChanged(object sender, EventArgs e)
        {
            if (suppressNoteSearch) return;
            var q = txtNoteSearch.Text;
            if (q == "Tìm kiếm tiêu đề...") q = "";
            RefreshNotesList(q);
        }

        private void LvNotes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvNotes == null || lvNotes.SelectedItems.Count == 0) return;
            var note = lvNotes.SelectedItems[0].Tag as NoteItem;
            if (note == null) return;
            currentNote = note;
            LoadNoteToEditor(note);
        }

        private void LoadNoteToEditor(NoteItem note)
        {
            txtNoteTitle.Text   = note.Title   ?? "";
            rtbNoteContent.Text = note.Content ?? "";
            var idx = cbNoteCategory.Items.IndexOf(note.Category ?? "Chung");
            cbNoteCategory.SelectedIndex = idx >= 0 ? idx : 0;
            lblNoteInfo.Text =
                "Tạo: " + note.CreatedAt.ToString("dd/MM/yyyy HH:mm") +
                "   |   Cập nhật: " + note.UpdatedAt.ToString("dd/MM/yyyy HH:mm") +
                "   |   ID: " + note.Id.Substring(0, 8);
        }

        // ── Tạo mới ──────────────────────────────────────────────────────────
        private void BtnNoteNew_Click(object sender, EventArgs e)
        {
            currentNote         = null;
            txtNoteTitle.Text   = "";
            rtbNoteContent.Text = "";
            if (cbNoteCategory != null) cbNoteCategory.SelectedIndex = 0;
            lblNoteInfo.Text    = "Ghi chú mới — chưa lưu";
            if (lvNotes != null) lvNotes.SelectedItems.Clear();
            txtNoteTitle.Focus();
        }

        // ── Lưu ──────────────────────────────────────────────────────────────
        private void BtnNoteSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtNoteTitle.Text))
            {
                MessageBox.Show("Vui lòng nhập tiêu đề ghi chú!",
                    "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtNoteTitle.Focus();
                return;
            }

            if (currentNote == null)
            {
                currentNote = new NoteItem();
                notesList.Insert(0, currentNote);
            }

            currentNote.Title     = txtNoteTitle.Text.Trim();
            currentNote.Content   = rtbNoteContent.Text;
            currentNote.Category  = cbNoteCategory.Text;
            currentNote.UpdatedAt = DateTime.Now;

            SaveNoteToFile(currentNote);
            SortNotes();

            var q = txtNoteSearch.Text == "Tìm kiếm tiêu đề..." ? "" : txtNoteSearch.Text;
            RefreshNotesList(q);

            lblNoteInfo.Text =
                "Tạo: " + currentNote.CreatedAt.ToString("dd/MM/yyyy HH:mm") +
                "   |   Cập nhật: " + currentNote.UpdatedAt.ToString("dd/MM/yyyy HH:mm");

            MessageBox.Show("Da luu: \"" + currentNote.Title + "\"",
                "Thanh cong", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ── Ghim / Bỏ ghim ───────────────────────────────────────────────────
        private void BtnNotePin_Click(object sender, EventArgs e)
        {
            if (currentNote == null)
            {
                MessageBox.Show("Chon ghi chu can ghim!",
                    "Canh bao", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            currentNote.IsPinned  = !currentNote.IsPinned;
            currentNote.UpdatedAt = DateTime.Now;
            SaveNoteToFile(currentNote);
            SortNotes();
            RefreshNotesList(txtNoteSearch.Text == "Tìm kiếm tiêu đề..." ? "" : txtNoteSearch.Text);
            MessageBox.Show(currentNote.IsPinned ? "Da ghim len dau!" : "Da bo ghim.",
                "Thong bao", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ── Xóa ──────────────────────────────────────────────────────────────
        private void BtnNoteDelete_Click(object sender, EventArgs e)
        {
            if (currentNote == null)
            {
                MessageBox.Show("Chon ghi chu can xoa!",
                    "Canh bao", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show("Xoa: \"" + currentNote.Title + "\"?",
                "Xac nhan", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            DeleteNoteFile(currentNote);
            notesList.Remove(currentNote);
            currentNote         = null;
            txtNoteTitle.Text   = "";
            rtbNoteContent.Text = "";
            lblNoteInfo.Text    = "Chua co ghi chu nao duoc chon";
            RefreshNotesList(txtNoteSearch.Text == "Tìm kiếm tiêu đề..." ? "" : txtNoteSearch.Text);
        }
    }
}
