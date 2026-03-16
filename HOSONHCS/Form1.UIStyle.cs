using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace HOSONHCS
{
    public partial class Form1
    {
        private void ApplyModernStyle() { }

        // =====================================================================
        // TAB 6 — Giới thiệu phần mềm (kiểu IT professional)
        // =====================================================================
        private void InitializeTab6()
        {
            try
            {
                if (tabPage6 == null) return;

                // Ẩn richTextBox5 gốc — ta dựng lại toàn bộ bằng code
                try { richTextBox5.Visible = false; } catch { }

                tabPage6.Controls.Clear();
                tabPage6.BackColor = UIStyler.BgMain;
                tabPage6.Padding   = new Padding(0);

                // ── SCROLLABLE CONTAINER ─────────────────────────────────────
                var scroll = new Panel
                {
                    Dock       = DockStyle.Fill,
                    AutoScroll = true,
                    BackColor  = UIStyler.BgMain
                };

                // ── NỘI DUNG bên trong scroll ────────────────────────────────
                var content = new Panel
                {
                    Width     = 860,
                    BackColor = UIStyler.BgMain,
                    Padding   = new Padding(32, 24, 32, 32)
                };

                int y = 24;

                // ── Logo / Badge ──────────────────────────────────────────────
                var badge = new Label
                {
                    Text      = "NHCSXH",
                    Location  = new Point(32, y),
                    Size      = new Size(90, 32),
                    Font      = new Font("Consolas", 11f, FontStyle.Bold),
                    ForeColor = UIStyler.BgMain,
                    BackColor = UIStyler.Primary,
                    TextAlign = ContentAlignment.MiddleCenter
                };
                content.Controls.Add(badge);

                var appName = new Label
                {
                    Text      = "Hồ Sơ NHCSXH",
                    Location  = new Point(134, y),
                    AutoSize  = true,
                    Font      = new Font("Segoe UI", 18f, FontStyle.Bold),
                    ForeColor = UIStyler.TextMain,
                    BackColor = Color.Transparent
                };
                content.Controls.Add(appName);
                y += 46;

                // ── Version / build info ──────────────────────────────────────
                var ver = Assembly.GetExecutingAssembly().GetName().Version;
                var verLabel = new Label
                {
                    Text      = $"v{ver.Major}.{ver.Minor}.{ver.Build}  •  .NET Framework 4.7.2  •  Windows Forms",
                    Location  = new Point(32, y),
                    AutoSize  = true,
                    Font      = new Font("Consolas", 9f),
                    ForeColor = UIStyler.TextSub,
                    BackColor = Color.Transparent
                };
                content.Controls.Add(verLabel);
                y += 28;

                // ── Đường kẻ phân cách ────────────────────────────────────────
                y = AddSep(content, y);

                // ── SECTION: Mô tả ────────────────────────────────────────────
                y = AddSection(content, y, "›  TỔNG QUAN", new[]
                {
                    "Phần mềm quản lý hồ sơ vay vốn Ngân hàng Chính sách Xã hội (NHCSXH),",
                    "hỗ trợ Hội, tổ cán bộ tín dụng lập hồ sơ, xuất mẫu biểu Word chuẩn Bộ,",
                    "quản lý danh sách khách hàng và tra cứu thông tin nhanh."
                });

                // ── SECTION: Tính năng ────────────────────────────────────────
                y = AddSection(content, y, "›  TÍNH NĂNG CHÍNH", new[]
                {
                    "  [01]  Nhập & lưu hồ sơ khách hàng vay vốn (JSON per file)",
                    "  [02]  Xuất mẫu biểu Word: 01/TD, 03/DS, GUQ, 01-TGTV, Bìa",
                    "  [03]  Quản lý & chỉnh sửa danh sách",
                    "  [04]  Lập bảng kê tiền giao dịch theo tổ trưởng, khách hàng",
                    "  [05]  Ghi chú nội bộ với danh mục & tìm kiếm",
                    "  [06]  Chatbot NHCSXH — hỏi đáp nghiệp vụ ngân hàng",
                    "  [07]  Cập nhật tự động qua GitHub Releases"
                });

                // ── SECTION: Stack kỹ thuật ───────────────────────────────────
                y = AddSection(content, y, "›  STACK KỸ THUẬT", new[]
                {
                    "  Language   C# 7.3",
                    "  Framework  .NET Framework 4.7.2",
                    "  UI         Windows Forms (WinForms)",
                    "  Storage    JSON (Newtonsoft.Json)  /  *.json per record",
                    "  Export     OpenXML (DocumentFormat.OpenXml)  →  *.docx",
                    "  Update     GitHub Releases API  +  Updater.exe",
                    "  IDE        Visual Studio 2022/2026"
                });

                // ── SECTION: Cấu trúc dữ liệu ────────────────────────────────
                y = AddSection(content, y, "›  CẤU TRÚC LƯU TRỮ", new[]
                {
                    "  <AppDir>\\",
                    "  ├── Customers\\         # hồ sơ khách hàng  (*.json)",
                    "  ├── ChatBot\\",
                    "  │   ├── knowledge\\     # Q&A chatbot       (qa_*.json)",
                    "  │   └── history\\       # lịch sử chat      (session_*.json)",
                    "  ├── Notes\\             # ghi chú nội bộ    (note_*.json)",
                    "  └── *.json              # danh sách ..."
                });

                // ── SECTION: Tác giả / liên hệ ───────────────────────────────
                y = AddSection(content, y, "›  TÁC GIẢ & LIÊN HỆ", new[]
                {
                    "  Developer   Nguyễn Hải",
                    "  Email       nxhaihg@gmail.com",
                    "  License     Internal use — NHCSXH"
                });

                // ── Footer ────────────────────────────────────────────────────
                y += 8;
                var footer = new Label
                {
                    Text      = $"Build {DateTime.Now.Year}  •  All rights reserved",
                    Location  = new Point(32, y),
                    AutoSize  = true,
                    Font      = new Font("Consolas", 8f),
                    ForeColor = UIStyler.TextHint,
                    BackColor = Color.Transparent
                };
                content.Controls.Add(footer);
                y += 30;

                content.Height = y;

                scroll.Controls.Add(content);

                // Căn giữa content theo chiều ngang khi resize
                scroll.Resize += (s, e) =>
                {
                    content.Left = Math.Max(0, (scroll.ClientSize.Width - content.Width) / 2);
                };

                tabPage6.Controls.Add(scroll);
            }
            catch { }
        }

        // ── Helpers render ────────────────────────────────────────────────────
        private static int AddSep(Control parent, int y)
        {
            var sep = new Panel
            {
                Location  = new Point(32, y),
                Size      = new Size(parent.Width - 64 > 0 ? parent.Width - 64 : 760, 1),
                BackColor = UIStyler.BorderColor
            };
            parent.Controls.Add(sep);
            return y + 18;
        }

        private static int AddSection(Control parent, int y, string title, string[] lines)
        {
            // Tiêu đề section
            var lblTitle = new Label
            {
                Text      = title,
                Location  = new Point(32, y),
                AutoSize  = true,
                Font      = new Font("Consolas", 9.5f, FontStyle.Bold),
                ForeColor = UIStyler.Primary,
                BackColor = Color.Transparent
            };
            parent.Controls.Add(lblTitle);
            y += 24;

            // Các dòng nội dung
            foreach (var line in lines)
            {
                var lbl = new Label
                {
                    Text      = line,
                    Location  = new Point(48, y),
                    AutoSize  = true,
                    Font      = new Font("Consolas", 9f),
                    ForeColor = UIStyler.TextMain,
                    BackColor = Color.Transparent
                };
                parent.Controls.Add(lbl);
                y += 20;
            }

            y += 8;
            return AddSep(parent, y);
        }
    }
}

