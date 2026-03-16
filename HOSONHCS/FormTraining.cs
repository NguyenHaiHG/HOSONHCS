using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace HOSONHCS
{
    public class FormTraining : Form
    {
        // ── Controls nhập liệu ──
        private ComboBox cbCategory;
        private ComboBox cbPriority;
        private TextBox txtKeywords;
        private TextBox txtQuestion;
        private RichTextBox rtbAnswer;

        // ── Nút thao tác ──
        private Button btnAdd;
        private Button btnUpdate;
        private Button btnDelete;
        private Button btnTrain;     // Nút HUẤN LUYỆN - lưu tất cả
        private Button btnClear;

        // ── Bảng danh sách Q&A ──
        private DataGridView dgvQA;

        // ── Nhãn thống kê ──
        private Label lblStats;

        // ── Dữ liệu ──
        private List<KnowledgeItem> items = new List<KnowledgeItem>();
        private int editingId = -1; // chỉ số đang sửa (-1 = đang thêm mới)

        public FormTraining()
        {
            InitializeLayout();
            LoadData();
        }

        // ================================================================
        // KHỞI TẠO GIAO DIỆN
        // ================================================================
        private void InitializeLayout()
        {
            this.Text = "🎓 Huấn luyện Chatbot NHCSXH";
            this.Size = new Size(900, 680);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Segoe UI", 9.5f);

            // ── Panel nhập liệu (bên trên) ──
            var pnlInput = new Panel
            {
                Dock = DockStyle.Top,
                Height = 260,
                Padding = new Padding(10)
            };

            // Danh mục
            var lblCat = new Label { Text = "Danh mục:", Location = new Point(10, 12), AutoSize = true };
            cbCategory = new ComboBox
            {
                Location = new Point(90, 9),
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDown
            };
            cbCategory.Items.AddRange(new object[]
            {
                "Hộ nghèo", "Hộ cận nghèo", "Hộ mới thoát nghèo",
                "SXKD", "GQVL", "NS&VSMT", "Thủ tục hồ sơ",
                "Lãi suất", "Thời hạn vay", "Chung"
            });

            // Độ ưu tiên
            var lblPri = new Label { Text = "Ưu tiên:", Location = new Point(310, 12), AutoSize = true };
            cbPriority = new ComboBox
            {
                Location = new Point(370, 9),
                Width = 80,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            for (int i = 10; i >= 1; i--)
                cbPriority.Items.Add(i);
            cbPriority.SelectedIndex = 4; // mặc định = 6

            // Từ khóa
            var lblKw = new Label { Text = "Từ khóa:", Location = new Point(10, 45), AutoSize = true };
            txtKeywords = new TextBox
            {
                Location = new Point(90, 42),
                Width = 500
            };

            // Câu hỏi
            var lblQ = new Label
            {
                Text = "Câu hỏi:",
                Location = new Point(10, 78),
                AutoSize = true,
                ForeColor = Color.DarkBlue,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };
            txtQuestion = new TextBox
            {
                Location = new Point(10, 98),
                Width = 860,
                Height = 50,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical
            };

            // Câu trả lời
            var lblA = new Label
            {
                Text = "Câu trả lời:",
                Location = new Point(10, 158),
                AutoSize = true,
                ForeColor = Color.DarkGreen,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };
            rtbAnswer = new RichTextBox
            {
                Location = new Point(10, 178),
                Width = 860,
                Height = 60,
                ScrollBars = RichTextBoxScrollBars.Vertical
            };

            pnlInput.Controls.AddRange(new Control[]
            {
                lblCat, cbCategory, lblPri, cbPriority,
                lblKw, txtKeywords,
                lblQ, txtQuestion,
                lblA, rtbAnswer
            });

            // ── Panel nút thao tác ──
            var pnlButtons = new Panel
            {
                Dock = DockStyle.Top,
                Height = 45,
                Padding = new Padding(10, 5, 10, 5)
            };

            btnAdd = new Button
            {
                Text = "➕ Thêm",
                Location = new Point(10, 8),
                Width = 100,
                Height = 30,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            btnUpdate = new Button
            {
                Text = "✏️ Cập nhật",
                Location = new Point(120, 8),
                Width = 110,
                Height = 30,
                BackColor = Color.FromArgb(255, 140, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Enabled = false
            };

            btnDelete = new Button
            {
                Text = "🗑️ Xóa",
                Location = new Point(240, 8),
                Width = 100,
                Height = 30,
                BackColor = Color.FromArgb(200, 50, 50),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Enabled = false
            };

            btnClear = new Button
            {
                Text = "🔄 Nhập mới",
                Location = new Point(350, 8),
                Width = 110,
                Height = 30,
                BackColor = Color.Gray,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            // Nút HUẤN LUYỆN - nổi bật nhất
            btnTrain = new Button
            {
                Text = "🎓 HUẤN LUYỆN - LƯU TẤT CẢ",
                Location = new Point(490, 8),
                Width = 260,
                Height = 30,
                BackColor = Color.FromArgb(34, 139, 34),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };

            pnlButtons.Controls.AddRange(new Control[]
            {
                btnAdd, btnUpdate, btnDelete, btnClear, btnTrain
            });

            // ── DataGridView danh sách Q&A ──
            dgvQA = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(240, 248, 255)
                }
            };

            dgvQA.Columns.Add(new DataGridViewTextBoxColumn { Name = "colId",       HeaderText = "#",         DataPropertyName = "Id",       Width = 60 });
            dgvQA.Columns.Add(new DataGridViewTextBoxColumn { Name = "colCat",      HeaderText = "Danh mục",  DataPropertyName = "Category", Width = 120 });
            dgvQA.Columns.Add(new DataGridViewTextBoxColumn { Name = "colQ",        HeaderText = "Câu hỏi",   DataPropertyName = "Question", Width = 280, DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True } });
            dgvQA.Columns.Add(new DataGridViewTextBoxColumn { Name = "colKw",       HeaderText = "Từ khóa",   DataPropertyName = "Keywords", Width = 160 });
            dgvQA.Columns.Add(new DataGridViewTextBoxColumn { Name = "colPri",      HeaderText = "Ưu tiên",   DataPropertyName = "Priority", Width = 65 });
            dgvQA.Columns.Add(new DataGridViewCheckBoxColumn { Name = "colActive",  HeaderText = "Bật",       DataPropertyName = "IsActive", Width = 45 });

            // ── Label thống kê ──
            lblStats = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 24,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = Color.DarkSlateGray,
                Padding = new Padding(6, 0, 0, 0)
            };

            // ── Gắn events ──
            btnAdd.Click    += BtnAdd_Click;
            btnUpdate.Click += BtnUpdate_Click;
            btnDelete.Click += BtnDelete_Click;
            btnClear.Click  += BtnClear_Click;
            btnTrain.Click  += BtnTrain_Click;
            dgvQA.CellClick += DgvQA_CellClick;

            // ── Thêm vào Form (thứ tự: Bottom → Fill → Top) ──
            this.Controls.Add(dgvQA);
            this.Controls.Add(lblStats);
            this.Controls.Add(pnlButtons);
            this.Controls.Add(pnlInput);
        }

        // ================================================================
        // TẢI DỮ LIỆU
        // ================================================================
        private void LoadData()
        {
            items = KnowledgeBaseManager.LoadAll();
            RefreshGrid();
        }

        private void RefreshGrid()
        {
            dgvQA.DataSource = null;
            dgvQA.DataSource = items;
            lblStats.Text = KnowledgeBaseManager.GetStats(items);
        }

        // ================================================================
        // NÚT THÊM — Đọc từ TextBox/RichTextBox → thêm vào danh sách
        // ================================================================
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ValidateInput()) return;

            var item = ReadForm();
            items.Add(item);
            RefreshGrid();
            ClearForm();

            lblStats.Text = $"✅ Đã thêm! | {KnowledgeBaseManager.GetStats(items)}";
        }

        // ================================================================
        // NÚT CẬP NHẬT
        // ================================================================
        private void BtnUpdate_Click(object sender, EventArgs e)
        {
            if (editingId < 0 || editingId >= items.Count) return;
            if (!ValidateInput()) return;

            var updated = ReadForm();
            updated.Id = items[editingId].Id;
            updated.CreatedDate = items[editingId].CreatedDate;
            items[editingId] = updated;

            RefreshGrid();
            ClearForm();

            lblStats.Text = $"✏️ Đã cập nhật! | {KnowledgeBaseManager.GetStats(items)}";
        }

        // ================================================================
        // NÚT XÓA
        // ================================================================
        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (editingId < 0 || editingId >= items.Count) return;

            var confirm = MessageBox.Show(
                $"Xóa câu hỏi:\n\"{items[editingId].Question}\"?",
                "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (confirm == DialogResult.Yes)
            {
                items.RemoveAt(editingId);
                RefreshGrid();
                ClearForm();
            }
        }

        // ================================================================
        // NÚT NHẬP MỚI
        // ================================================================
        private void BtnClear_Click(object sender, EventArgs e) => ClearForm();

        // ================================================================
        // NÚT HUẤN LUYỆN — Lưu toàn bộ danh sách ra file JSON
        // ================================================================
        private void BtnTrain_Click(object sender, EventArgs e)
        {
            if (items.Count == 0)
            {
                MessageBox.Show("Chưa có Q&A nào để huấn luyện!\n\nVui lòng thêm câu hỏi và câu trả lời trước.",
                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                KnowledgeBaseManager.SaveAll(items);

                MessageBox.Show(
                    $"🎓 Huấn luyện thành công!\n\n" +
                    $"✅ Đã lưu {items.Count} cặp Q&A\n" +
                    $"📁 File: ChatBot\\knowledge\\knowledge_base.json\n\n" +
                    $"Chatbot đã sẵn sàng trả lời!",
                    "✅ Huấn luyện hoàn tất",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                lblStats.Text = $"🎓 Đã huấn luyện lúc {DateTime.Now:HH:mm} | {KnowledgeBaseManager.GetStats(items)}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi khi lưu:\n{ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ================================================================
        // CLICK VÀO DÒNG TRONG GRID → Load lên form để sửa
        // ================================================================
        private void DgvQA_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= items.Count) return;

            editingId = e.RowIndex;
            var item = items[editingId];

            cbCategory.Text  = item.Category ?? "";
            txtKeywords.Text = item.Keywords ?? "";
            txtQuestion.Text = item.Question ?? "";
            rtbAnswer.Text   = item.Answer ?? "";

            // Chọn đúng mức ưu tiên
            cbPriority.SelectedItem = item.Priority;

            btnUpdate.Enabled = true;
            btnDelete.Enabled = true;
            btnAdd.Enabled    = false;
        }

        // ================================================================
        // HELPERS
        // ================================================================
        private KnowledgeItem ReadForm()
        {
            return new KnowledgeItem
            {
                Category    = cbCategory.Text.Trim(),
                Question    = txtQuestion.Text.Trim(),
                Keywords    = txtKeywords.Text.Trim(),
                Answer      = rtbAnswer.Text.Trim(),
                Priority    = cbPriority.SelectedItem != null ? (int)cbPriority.SelectedItem : 5,
                IsActive    = true,
                UpdatedDate = DateTime.Now
            };
        }

        private void ClearForm()
        {
            txtQuestion.Text = "";
            rtbAnswer.Text   = "";
            txtKeywords.Text = "";
            cbCategory.Text  = "";
            cbPriority.SelectedIndex = 4;

            editingId = -1;
            btnUpdate.Enabled = false;
            btnDelete.Enabled = false;
            btnAdd.Enabled    = true;

            txtQuestion.Focus();
        }

        private bool ValidateInput()
        {
            if (string.IsNullOrWhiteSpace(txtQuestion.Text))
            {
                MessageBox.Show("Vui lòng nhập câu hỏi!", "Thiếu thông tin",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtQuestion.Focus();
                return false;
            }
            if (string.IsNullOrWhiteSpace(rtbAnswer.Text))
            {
                MessageBox.Show("Vui lòng nhập câu trả lời!", "Thiếu thông tin",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                rtbAnswer.Focus();
                return false;
            }
            return true;
        }
    }
}
