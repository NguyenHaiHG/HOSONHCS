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
        // =====================================================================
        // FIELDS
        // =====================================================================
        // -- Chat panel (trái) --
        private RichTextBox rtbChat;
        private TextBox     txtChatInput;
        private Button      btnChatSend;
        private Button      btnChatNew;
        private Button      btnChatClearView;

        // -- Training panel (phải) --
        private ListView    lvKnowledge;
        private TextBox     txtKnowQuestion;
        private TextBox     txtKnowKeywords;
        private RichTextBox rtbKnowAnswer;
        private ComboBox    cbKnowCategory;
        private Button      btnKnowSave;
        private Button      btnKnowDelete;
        private Button      btnKnowReset;
        private KnowledgeItem editingKnowledge = null;

        // -- Lịch sử chat --
        private ListView lvHistory;

        // -- Data --
        private List<KnowledgeItem> knowledgeList  = new List<KnowledgeItem>();
        private ChatSession         currentSession;

        // -- Paths --
        private string GetKnowledgePath()
            => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ChatBot", "knowledge");
        private string GetHistoryPath()
            => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ChatBot", "history");

        // =====================================================================
        // KHỞI TẠO UI — Gọi từ InitializeApp()
        // =====================================================================
        private void InitializeChatbotTab()
        {
            if (tabPage2 == null) return;

            tabPage2.Controls.Clear();
            tabPage2.BackColor = UIStyler.BgMain;
            tabPage2.Padding   = new Padding(0);

            // ── PANEL PHẢI: Huấn luyện (Dock Right, 295px) ───────────────────
            var pnlRight = new Panel
            {
                Dock      = DockStyle.Right,
                Width     = 295,
                BackColor = UIStyler.BgPanel
            };

            var lblTrainHdr = new Label
            {
                Text      = "  HUẤN LUYỆN CHATBOT",
                Dock      = DockStyle.Top,
                Height    = 26,
                Font      = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                BackColor = UIStyler.BgCard,
                ForeColor = UIStyler.Primary,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding   = new Padding(6, 0, 0, 0)
            };

            var pnlForm = new Panel
            {
                Dock        = DockStyle.Top,
                Height      = 220,
                BackColor   = UIStyler.BgPanel,
                Padding     = new Padding(5, 4, 5, 4),
                AutoScroll  = true
            };
            BuildTrainingForm(pnlForm);

            var sep = new Panel { Dock = DockStyle.Top, Height = 1, BackColor = UIStyler.BorderColor };

            var lblListHdr = new Label
            {
                Text      = "  Danh sách Q&A đã lưu",
                Dock      = DockStyle.Top,
                Height    = 24,
                Font      = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                BackColor = UIStyler.BgCard,
                ForeColor = UIStyler.Primary,
                TextAlign = ContentAlignment.MiddleLeft
            };

            lvKnowledge = new ListView
            {
                Dock          = DockStyle.Fill,
                View          = View.Details,
                FullRowSelect = true,
                GridLines     = true,
                Font          = new Font("Segoe UI", 8.5f),
                BorderStyle   = BorderStyle.None,
                BackColor     = UIStyler.BgPanel,
                ForeColor     = UIStyler.TextMain
            };
            lvKnowledge.Columns.Add("Câu hỏi",   185);
            lvKnowledge.Columns.Add("Danh mục",   100);
            lvKnowledge.SelectedIndexChanged += LvKnowledge_SelectedIndexChanged;
            lvKnowledge.DoubleClick          += LvKnowledge_DoubleClick;

            pnlRight.Controls.Add(lvKnowledge);
            pnlRight.Controls.Add(lblListHdr);
            pnlRight.Controls.Add(sep);
            pnlRight.Controls.Add(pnlForm);
            pnlRight.Controls.Add(lblTrainHdr);

            var vline = new Panel { Dock = DockStyle.Right, Width = 1, BackColor = UIStyler.BorderColor };

            var pnlLeft = new Panel { Dock = DockStyle.Fill, BackColor = UIStyler.BgMain };

            var pnlChatHdr = new Panel
            {
                Dock      = DockStyle.Top,
                Height    = 38,
                BackColor = UIStyler.BgPanel,
                Padding   = new Padding(8, 6, 8, 6)
            };
            var lblChatTitle = new Label
            {
                Text      = "NHCSXH Chatbot",
                Location  = new Point(8, 9),
                AutoSize  = true,
                Font      = new Font("Segoe UI", 10.5f, FontStyle.Bold),
                ForeColor = UIStyler.Primary,
                BackColor = Color.Transparent
            };
            btnChatNew       = MakeChatBtn("+ Phiên mới",  UIStyler.BtnGreen);
            btnChatClearView = MakeChatBtn("Xóa màn hình", UIStyler.BtnGray);
            btnChatNew.Click       += BtnChatNew_Click;
            btnChatClearView.Click += (s, e) => { rtbChat.Clear(); AppendBotMsg("Màn hình đã được xóa. Lịch sử vẫn được lưu."); };
            pnlChatHdr.Controls.AddRange(new Control[] { lblChatTitle, btnChatNew, btnChatClearView });
            pnlChatHdr.Resize += (s, e) =>
            {
                btnChatClearView.Location = new Point(pnlChatHdr.Width - btnChatClearView.Width - 6, 7);
                btnChatNew.Location       = new Point(pnlChatHdr.Width - btnChatNew.Width - btnChatClearView.Width - 12, 7);
            };

            rtbChat = new RichTextBox
            {
                Dock        = DockStyle.Fill,
                ReadOnly    = true,
                BorderStyle = BorderStyle.None,
                BackColor   = UIStyler.BgMain,
                ForeColor   = UIStyler.TextMain,
                Font        = new Font("Segoe UI", 11f),
                ScrollBars  = RichTextBoxScrollBars.Vertical
            };

            var pnlInput = new Panel
            {
                Dock      = DockStyle.Bottom,
                Height    = 46,
                BackColor = UIStyler.BgPanel,
                Padding   = new Padding(6, 8, 6, 8)
            };
            txtChatInput = new TextBox
            {
                Location  = new Point(6, 10),
                Height    = 26,
                Font      = new Font("Segoe UI", 10f),
                BorderStyle = BorderStyle.FixedSingle
            };
            txtChatInput.KeyDown += TxtChatInput_KeyDown;

            btnChatSend = new Button
            {
                Text      = "Gửi  ▶",
                Height    = 28,
                Width     = 90,
                BackColor = Color.FromArgb(0, 123, 255),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Cursor    = Cursors.Hand
            };
            btnChatSend.FlatAppearance.BorderSize = 0;
            btnChatSend.Click += BtnChatSend_Click;

            // Resize để fill đúng
            pnlInput.Resize += (s, e) =>
            {
                btnChatSend.Size     = new Size(90, 28);
                btnChatSend.Location = new Point(pnlInput.Width - 96, 9);
                txtChatInput.Width   = pnlInput.Width - 104;
                txtChatInput.Location = new Point(6, 10);
            };

            pnlInput.Controls.Add(txtChatInput);
            pnlInput.Controls.Add(btnChatSend);

            // ── Panel lịch sử (Dock Bottom, 160px) ──────────────────────────
            var pnlHistory = new Panel
            {
                Dock      = DockStyle.Bottom,
                Height    = 165,
                BackColor = UIStyler.BgPanel
            };

            var lblHistHdr = new Label
            {
                Text      = "  LỊCH SỬ CÁC PHIÊN CHAT",
                Dock      = DockStyle.Top,
                Height    = 24,
                Font      = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                BackColor = UIStyler.BgCard,
                ForeColor = UIStyler.Primary,
                TextAlign = ContentAlignment.MiddleLeft
            };

            var btnHistDelete = new Button
            {
                Text      = "Xóa phiên",
                Dock      = DockStyle.Bottom,
                Height    = 24,
                BackColor = UIStyler.BtnRed,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 8.5f),
                Cursor    = Cursors.Hand
            };
            btnHistDelete.FlatAppearance.BorderSize = 0;
            btnHistDelete.Click += (s, e) =>
            {
                if (lvHistory.SelectedItems.Count == 0) return;
                var path = lvHistory.SelectedItems[0].Tag as string;
                if (path != null && File.Exists(path))
                {
                    File.Delete(path);
                    RefreshHistoryList();
                }
            };

            lvHistory = new ListView
            {
                Dock          = DockStyle.Fill,
                View          = View.Details,
                FullRowSelect = true,
                GridLines     = false,
                Font          = new Font("Segoe UI", 9f),
                BorderStyle   = BorderStyle.None,
                BackColor     = UIStyler.BgPanel,
                ForeColor     = UIStyler.TextMain,
                HideSelection = false,
                HeaderStyle   = ColumnHeaderStyle.None
            };
            lvHistory.Columns.Add("", -2);  // tự fill toàn bộ chiều rộng
            lvHistory.Resize += (s, e) =>
            {
                if (lvHistory.Columns.Count > 0)
                    lvHistory.Columns[0].Width = lvHistory.ClientSize.Width;
            };
            lvHistory.DoubleClick += (s, e) =>
            {
                if (lvHistory.SelectedItems.Count == 0) return;
                var path = lvHistory.SelectedItems[0].Tag as string;
                if (path == null || !File.Exists(path)) return;
                try
                {
                    var sess = JsonConvert.DeserializeObject<ChatSession>(
                        File.ReadAllText(path, Encoding.UTF8));
                    if (sess == null) return;
                    rtbChat.Clear();
                    AppendBotMsg("── Xem lại phiên: " + sess.Date.ToString("dd/MM/yyyy HH:mm") + " ──");
                    foreach (var msg in sess.Messages)
                    {
                        if (msg.Role == "user") AppendUserMsg(msg.Content);
                        else                    AppendBotMsg(msg.Content);
                    }
                }
                catch { }
            };

            var sepHist = new Panel { Dock = DockStyle.Top, Height = 1, BackColor = UIStyler.BorderColor };

            pnlHistory.Controls.Add(lvHistory);
            pnlHistory.Controls.Add(btnHistDelete);
            pnlHistory.Controls.Add(lblHistHdr);
            pnlHistory.Controls.Add(sepHist);

            // Thứ tự add vào pnlLeft
            pnlLeft.Controls.Add(rtbChat);
            pnlLeft.Controls.Add(pnlHistory);
            pnlLeft.Controls.Add(pnlInput);
            pnlLeft.Controls.Add(pnlChatHdr);

            // ── Gắn vào tabPage2 ──────────────────────────────────────────────
            tabPage2.Controls.Add(pnlLeft);
            tabPage2.Controls.Add(vline);
            tabPage2.Controls.Add(pnlRight);

            // ── Load data & khởi động phiên ──────────────────────────────────
            LoadKnowledge();
            RefreshKnowledgeList();
            SeedDefaultKnowledge();
            StartNewSession();
        }

        // Build form nhập Q&A trong pnlForm
        private void BuildTrainingForm(Panel p)
        {
            int y = 4;

            // Danh mục
            var lblCat = new Label { Text = "Danh mục:", Location = new Point(0, y), AutoSize = true, Font = new Font("Segoe UI", 8.5f), ForeColor = UIStyler.TextMain, BackColor = Color.Transparent };
            cbKnowCategory = new ComboBox
            {
                Location      = new Point(62, y - 2),
                Width         = 218,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font          = new Font("Segoe UI", 8.5f),
                BackColor     = UIStyler.BgInput,
                ForeColor     = UIStyler.TextMain,
                FlatStyle     = FlatStyle.Flat
            };
            cbKnowCategory.Items.AddRange(new object[]
                { "Chung", "Lãi suất", "Thủ tục", "Đối tượng", "Hồ sơ", "Mức vay", "Hộ nghèo", "SXKD", "GQVL", "NS&VSMT" });
            cbKnowCategory.SelectedIndex = 0;
            y += 22;

            // Câu hỏi
            var lblQ = new Label { Text = "Câu hỏi:", Location = new Point(0, y), AutoSize = true, Font = new Font("Segoe UI", 8.5f, FontStyle.Bold), ForeColor = UIStyler.Primary, BackColor = Color.Transparent };
            y += 16;
            txtKnowQuestion = new TextBox
            {
                Location    = new Point(0, y),
                Width       = 280,
                Height      = 38,
                Multiline   = true,
                ScrollBars  = ScrollBars.Vertical,
                Font        = new Font("Segoe UI", 8.5f),
                BackColor   = UIStyler.BgInput,
                ForeColor   = UIStyler.TextMain,
                BorderStyle = BorderStyle.FixedSingle
            };
            y += 42;

            var lblKw = new Label { Text = "Từ khóa (dấu phẩy):", Location = new Point(0, y), AutoSize = true, Font = new Font("Segoe UI", 8f), ForeColor = UIStyler.TextSub, BackColor = Color.Transparent };
            y += 16;
            txtKnowKeywords = new TextBox { Location = new Point(0, y), Width = 280, Font = new Font("Segoe UI", 8.5f), BackColor = UIStyler.BgInput, ForeColor = UIStyler.TextMain, BorderStyle = BorderStyle.FixedSingle };
            y += 22;

            var lblA = new Label { Text = "Câu trả lời:", Location = new Point(0, y), AutoSize = true, Font = new Font("Segoe UI", 8.5f, FontStyle.Bold), ForeColor = UIStyler.Primary, BackColor = Color.Transparent };
            y += 16;
            rtbKnowAnswer = new RichTextBox
            {
                Location    = new Point(0, y),
                Width       = 280,
                Height      = 42,
                BorderStyle = BorderStyle.FixedSingle,
                Font        = new Font("Segoe UI", 8.5f),
                BackColor   = UIStyler.BgInput,
                ForeColor   = UIStyler.TextMain,
                ScrollBars  = RichTextBoxScrollBars.Vertical
            };
            y += 46;

            // Buttons
            btnKnowSave   = MakeSmBtn("Luu", 0,   UIStyler.BtnBlue);
            btnKnowDelete = MakeSmBtn("Xoa", 76,  UIStyler.BtnRed);
            btnKnowReset  = MakeSmBtn("Moi", 152, UIStyler.BtnGray);
            btnKnowSave.Location   = new Point(0,   y);
            btnKnowDelete.Location = new Point(76,  y);
            btnKnowReset.Location  = new Point(152, y);
            btnKnowSave.Click   += BtnKnowSave_Click;
            btnKnowDelete.Click += BtnKnowDelete_Click;
            btnKnowReset.Click  += (s, e) => ClearTrainingForm();

            p.Controls.AddRange(new Control[]
            {
                lblCat, cbKnowCategory,
                lblQ, txtKnowQuestion,
                lblKw, txtKnowKeywords,
                lblA, rtbKnowAnswer,
                btnKnowSave, btnKnowDelete, btnKnowReset
            });
        }

        private Button MakeChatBtn(string text, Color color)
        {
            var b = new Button { Text = text, Width = 105, Height = 24, BackColor = color, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8f), Cursor = Cursors.Hand };
            b.FlatAppearance.BorderSize = 0;
            return b;
        }
        private Button MakeSmBtn(string text, int x, Color color)
        {
            var b = new Button { Text = text, Location = new Point(x, 0), Width = 70, Height = 24, BackColor = color, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5f), Cursor = Cursors.Hand };
            b.FlatAppearance.BorderSize = 0;
            return b;
        }

        // =====================================================================
        // PHIÊN CHAT
        // =====================================================================
        private void StartNewSession()
        {
            currentSession = new ChatSession();
            rtbChat.Clear();
            AppendBotMsg("Xin chào! Tôi là Chatbot NHCSXH. Bạn có thể hỏi tôi về lãi suất, thủ tục vay vốn, đối tượng vay, hồ sơ cần thiết...");
            AppendBotMsg("Gõ câu hỏi vào ô bên dưới và nhấn Enter hoặc bấm nút Gửi.");
        }

        private void BtnChatNew_Click(object sender, EventArgs e)
        {
            SaveSession();
            StartNewSession();
        }

        // =====================================================================
        // HIỂN THỊ CHAT
        // =====================================================================
        private void AppendUserMsg(string text)
        {
            AppendLine("[" + DateTime.Now.ToString("HH:mm") + "] Bạn: ",
                       Color.FromArgb(100, 180, 255), FontStyle.Bold);
            AppendLine(text + "\n", Color.FromArgb(160, 210, 255), FontStyle.Regular);
            AddChatToSession("user", text);
            rtbChat.ScrollToCaret();
        }

        private void AppendBotMsg(string text)
        {
            AppendLine("[" + DateTime.Now.ToString("HH:mm") + "] Bot: ",
                       UIStyler.Primary, FontStyle.Bold);
            AppendLine(text + "\n\n", Color.FromArgb(160, 230, 228), FontStyle.Regular);
            if (currentSession != null) AddChatToSession("bot", text);
            rtbChat.ScrollToCaret();
        }

        private void AppendLine(string text, Color color, FontStyle style)
        {
            rtbChat.SelectionStart  = rtbChat.TextLength;
            rtbChat.SelectionLength = 0;
            rtbChat.SelectionColor  = color;
            rtbChat.SelectionFont   = new Font(rtbChat.Font, style);
            rtbChat.AppendText(text);
            rtbChat.SelectionColor  = rtbChat.ForeColor;
            // Cuộn xuống cuối — hoạt động đúng với ReadOnly RichTextBox
            rtbChat.SelectionStart  = rtbChat.TextLength;
            rtbChat.ScrollToCaret();
        }

        private void AddChatToSession(string role, string content)
        {
            if (currentSession == null) return;
            currentSession.Messages.Add(new ChatMessage { Role = role, Content = content });
        }

        // =====================================================================
        // XỬ LÝ GỬI TIN
        // =====================================================================
        private void TxtChatInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && !e.Shift)
            {
                e.SuppressKeyPress = true;
                SendChatMessage();
            }
        }

        private void BtnChatSend_Click(object sender, EventArgs e) => SendChatMessage();

        private void SendChatMessage()
        {
            var input = (txtChatInput.Text ?? "").Trim();
            if (string.IsNullOrWhiteSpace(input)) return;

            txtChatInput.Clear();
            AppendUserMsg(input);

            var answer = GetBotAnswer(input);
            AppendBotMsg(answer);

            SaveSession();
        }

        // =====================================================================
        // THUẬT TOÁN MATCHING
        // =====================================================================
        private string GetBotAnswer(string input)
        {
            if (knowledgeList.Count == 0)
                return "Tôi chưa có dữ liệu. Vui lòng huấn luyện tôi bằng cách thêm Q&A ở bảng bên phải!";

            var best = FindBestMatch(input);
            if (best != null) return best.Answer;

            return "Xin lỗi, tôi chưa có thông tin về câu hỏi này.\n" +
                   "Bạn có thể bổ sung kiến thức cho tôi bằng cách nhập Q&A vào bảng \"Huấn luyện\" bên phải.";
        }

        private KnowledgeItem FindBestMatch(string input)
        {
            var normInput = RemoveVietnameseTones(input.ToLowerInvariant());
            var words     = normInput.Split(new[] { ' ', ',', '.', '?', '!', ';' },
                                            StringSplitOptions.RemoveEmptyEntries);

            KnowledgeItem best      = null;
            int           bestScore = 0;

            foreach (var item in knowledgeList.Where(k => k.IsActive))
            {
                int score = 0;

                // 1. Khớp từ khóa (quan trọng nhất)
                if (item.Keywords != null)
                {
                    foreach (var kw in item.Keywords)
                    {
                        var nkw = RemoveVietnameseTones((kw ?? "").ToLowerInvariant().Trim());
                        if (string.IsNullOrWhiteSpace(nkw)) continue;

                        if (normInput.Contains(nkw))
                            score += 12;
                        else
                        {
                            var kwWords = nkw.Split(' ');
                            score += kwWords.Count(w => words.Contains(w)) * 4;
                        }
                    }
                }

                // 2. Khớp câu hỏi
                var normQ  = RemoveVietnameseTones(item.Question.ToLowerInvariant());
                var qWords = normQ.Split(new[] { ' ', ',', '.', '?', '!' },
                                         StringSplitOptions.RemoveEmptyEntries);
                score += words.Intersect(qWords).Count() * 2;

                // 3. Ưu tiên
                score += item.Priority;

                if (score > bestScore && score >= 10)
                {
                    bestScore = score;
                    best      = item;
                }
            }
            return best;
        }

        // =====================================================================
        // KNOWLEDGE BASE — CRUD
        // =====================================================================
        private void LoadKnowledge()
        {
            knowledgeList.Clear();
            try
            {
                var folder = GetKnowledgePath();
                if (!Directory.Exists(folder)) return;
                foreach (var f in Directory.EnumerateFiles(folder, "qa_*.json"))
                {
                    try
                    {
                        var item = JsonConvert.DeserializeObject<KnowledgeItem>(
                            File.ReadAllText(f, Encoding.UTF8));
                        if (item != null) knowledgeList.Add(item);
                    }
                    catch { }
                }
                knowledgeList = knowledgeList.OrderByDescending(k => k.Priority)
                                             .ThenBy(k => k.Category).ToList();
            }
            catch { }
        }

        private void SaveKnowledgeItem(KnowledgeItem item)
        {
            try
            {
                var folder = GetKnowledgePath();
                if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
                File.WriteAllText(
                    Path.Combine(folder, "qa_" + item.Id + ".json"),
                    JsonConvert.SerializeObject(item, Formatting.Indented),
                    Encoding.UTF8);
            }
            catch { }
        }

        private void DeleteKnowledgeFile(KnowledgeItem item)
        {
            try
            {
                var path = Path.Combine(GetKnowledgePath(), "qa_" + item.Id + ".json");
                if (File.Exists(path)) File.Delete(path);
            }
            catch { }
        }

        private void RefreshKnowledgeList()
        {
            if (lvKnowledge == null) return;
            lvKnowledge.BeginUpdate();
            lvKnowledge.Items.Clear();
            foreach (var item in knowledgeList)
            {
                var lvi = new ListViewItem(
                    (item.Question.Length > 40 ? item.Question.Substring(0, 37) + "..." : item.Question));
                lvi.SubItems.Add(item.Category);
                lvi.Tag = item;
                if (!item.IsActive) lvi.ForeColor = Color.Gray;
                lvKnowledge.Items.Add(lvi);
            }
            lvKnowledge.EndUpdate();
        }

        // Seed dữ liệu mặc định nếu chưa có
        private void SeedDefaultKnowledge()
        {
            if (knowledgeList.Count > 0) return;

            foreach (var item in KnowledgeSeed.GetAll())
            {
                knowledgeList.Add(item);
                SaveKnowledgeItem(item);
            }
            RefreshKnowledgeList();
        }
                // =====================================================================
                // EVENT HANDLERS — Training
                // =====================================================================
                private void LvKnowledge_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvKnowledge.SelectedItems.Count == 0) return;
            var item = lvKnowledge.SelectedItems[0].Tag as KnowledgeItem;
            if (item == null) return;
            editingKnowledge = item;
            LoadKnowledgeToForm(item);
        }

        private void LvKnowledge_DoubleClick(object sender, EventArgs e)
        {
            // Double-click: load lên form để sửa
            LvKnowledge_SelectedIndexChanged(sender, e);
        }

        private void LoadKnowledgeToForm(KnowledgeItem item)
        {
            txtKnowQuestion.Text  = item.Question;
            txtKnowKeywords.Text  = item.Keywords != null ? string.Join(", ", item.Keywords) : "";
            rtbKnowAnswer.Text    = item.Answer;
            var idx = cbKnowCategory.Items.IndexOf(item.Category);
            cbKnowCategory.SelectedIndex = idx >= 0 ? idx : 0;
        }

        private void BtnKnowSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtKnowQuestion.Text))
            {
                MessageBox.Show("Vui lòng nhập câu hỏi!", "Cảnh báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(rtbKnowAnswer.Text))
            {
                MessageBox.Show("Vui lòng nhập câu trả lời!", "Cảnh báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (editingKnowledge == null) editingKnowledge = new KnowledgeItem();

            editingKnowledge.Question  = txtKnowQuestion.Text.Trim();
            editingKnowledge.Answer    = rtbKnowAnswer.Text.Trim();
            editingKnowledge.Category  = cbKnowCategory.Text;
            editingKnowledge.Keywords  = txtKnowKeywords.Text
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(k => k.Trim())
                .Where(k => !string.IsNullOrWhiteSpace(k))
                .ToArray();

            SaveKnowledgeItem(editingKnowledge);

            if (!knowledgeList.Any(k => k.Id == editingKnowledge.Id))
                knowledgeList.Add(editingKnowledge);

            RefreshKnowledgeList();
            ClearTrainingForm();

            MessageBox.Show("Da luu Q&A!", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnKnowDelete_Click(object sender, EventArgs e)
        {
            if (editingKnowledge == null)
            {
                MessageBox.Show("Chon dong can xoa trong danh sach!",
                    "Canh bao", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show("Xoa Q&A nay?", "Xac nhan",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

            DeleteKnowledgeFile(editingKnowledge);
            knowledgeList.Remove(editingKnowledge);
            editingKnowledge = null;
            RefreshKnowledgeList();
            ClearTrainingForm();
        }

        private void ClearTrainingForm()
        {
            editingKnowledge     = null;
            txtKnowQuestion.Text = "";
            txtKnowKeywords.Text = "";
            rtbKnowAnswer.Text   = "";
            if (cbKnowCategory != null) cbKnowCategory.SelectedIndex = 0;
            if (lvKnowledge != null) lvKnowledge.SelectedItems.Clear();
        }

        // =====================================================================
        // LỊCH SỬ CHAT — Save / Load
        // =====================================================================
        private void SaveSession()
        {
            if (currentSession == null || currentSession.Messages.Count == 0) return;
            try
            {
                var folder = GetHistoryPath();
                if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
                var path = Path.Combine(folder,
                    "session_" + currentSession.Date.ToString("yyyy-MM-dd_HHmmss") + ".json");
                File.WriteAllText(path,
                    JsonConvert.SerializeObject(currentSession, Formatting.Indented),
                    Encoding.UTF8);
                RefreshHistoryList();
            }
            catch { }
        }

        private void RefreshHistoryList()
        {
            if (lvHistory == null) return;
            try
            {
                lvHistory.BeginUpdate();
                lvHistory.Items.Clear();
                // Cột tự co giãn theo chiều rộng
                if (lvHistory.Columns.Count > 0)
                    lvHistory.Columns[0].Width = lvHistory.ClientSize.Width > 0
                        ? lvHistory.ClientSize.Width : 400;

                var folder = GetHistoryPath();
                if (!Directory.Exists(folder)) { lvHistory.EndUpdate(); return; }

                var files = Directory.GetFiles(folder, "session_*.json")
                                     .OrderByDescending(f => f);
                foreach (var file in files)
                {
                    try
                    {
                        var sess = JsonConvert.DeserializeObject<ChatSession>(
                            File.ReadAllText(file, Encoding.UTF8));
                        if (sess == null) continue;

                        // Lấy câu hỏi đầu tiên của user
                        var firstUser = sess.Messages
                            .FirstOrDefault(m => m.Role == "user");
                        var preview = firstUser != null
                            ? firstUser.Content.Replace("\n", " ").Trim()
                            : "(phiên trống)";
                        // Rút gọn nếu quá dài
                        if (preview.Length > 60)
                            preview = preview.Substring(0, 57) + "...";

                        // Prefix ngày giờ ngắn
                        var label = sess.Date.ToString("dd/MM HH:mm") + "  " + preview;

                        var lvi = new ListViewItem(label);
                        lvi.Tag       = file;
                        lvi.ForeColor = UIStyler.TextMain;
                        lvHistory.Items.Add(lvi);
                    }
                    catch { }
                }
            }
            finally { lvHistory.EndUpdate(); }
        }
    }
}
