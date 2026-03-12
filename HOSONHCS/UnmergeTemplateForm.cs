using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace HOSONHCS
{
    /// <summary>
    /// Form đơn giản để tách gộp ô trong file template "01 SXKD.docx"
    /// Chạy một lần rồi có thể xóa
    /// </summary>
    public class UnmergeTemplateForm : Form
    {
        private Button btnUnmerge;
        private Label lblInfo;
        private TextBox txtLog;

        public UnmergeTemplateForm()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            // Form settings
            this.Text = "Tách gộp ô Template - 01 SXKD.docx";
            this.Size = new Size(600, 400);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            // Label thông tin
            lblInfo = new Label
            {
                Text = "Công cụ này sẽ tách gộp tất cả các ô trong file template \"01 SXKD.docx\"\n\n" +
                       "⚠️ Lưu ý: File gốc sẽ được backup với tên \"01 SXKD_BACKUP_MERGED.docx\"",
                Location = new Point(20, 20),
                Size = new Size(540, 60),
                Font = new Font("Segoe UI", 10F)
            };

            // Button tách gộp
            btnUnmerge = new Button
            {
                Text = "TÁCH GỘP Ô - 01 SXKD.docx",
                Location = new Point(150, 100),
                Size = new Size(280, 50),
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnUnmerge.Click += BtnUnmerge_Click;

            // TextBox log
            txtLog = new TextBox
            {
                Location = new Point(20, 170),
                Size = new Size(540, 160),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                Font = new Font("Consolas", 9F)
            };

            // Add controls
            this.Controls.Add(lblInfo);
            this.Controls.Add(btnUnmerge);
            this.Controls.Add(txtLog);
        }

        private void BtnUnmerge_Click(object sender, EventArgs e)
        {
            txtLog.Clear();
            AppendLog("Đang tìm file template...");

            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string templateName = "01 SXKD.docx";
            string templatePath = null;

            // Thử tìm trong thư mục Templates
            string path1 = Path.Combine(baseDir, "Templates", templateName);
            if (File.Exists(path1))
            {
                templatePath = path1;
                AppendLog($"✓ Tìm thấy: {path1}");
            }
            else
            {
                AppendLog($"✗ Không tìm thấy: {path1}");
                
                // Thử tìm trong thư mục gốc
                string path2 = Path.Combine(baseDir, templateName);
                if (File.Exists(path2))
                {
                    templatePath = path2;
                    AppendLog($"✓ Tìm thấy: {path2}");
                }
                else
                {
                    AppendLog($"✗ Không tìm thấy: {path2}");
                }
            }

            if (string.IsNullOrEmpty(templatePath))
            {
                AppendLog("\n❌ KHÔNG TÌM THẤY FILE TEMPLATE!");
                MessageBox.Show(
                    $"Không tìm thấy file '{templateName}'",
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            AppendLog($"\nĐang tách gộp ô...");
            
            try
            {
                UnmergeTemplateUtil.UnmergeCells(templatePath);
                AppendLog("\n✓ Hoàn tất!");
            }
            catch (Exception ex)
            {
                AppendLog($"\n❌ Lỗi: {ex.Message}");
            }
        }

        private void AppendLog(string message)
        {
            txtLog.AppendText(message + Environment.NewLine);
        }

        /// <summary>
        /// Hiển thị form tách gộp ô
        /// </summary>
        public static void ShowUnmergeForm()
        {
            using (var form = new UnmergeTemplateForm())
            {
                form.ShowDialog();
            }
        }
    }
}
