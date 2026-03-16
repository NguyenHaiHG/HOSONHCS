using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace HOSONHCS
{
    /// <summary>Teal theme — nền xanh teal, dễ nhìn, không chói.</summary>
    public static class UIStyler
    {
        // ── Nền Teal (3 cấp độ, từ tối → sáng dần) ───────────────────────────
        public static readonly Color BgMain    = Color.FromArgb(13,  75,  75);   // nền chính
        public static readonly Color BgPanel   = Color.FromArgb(18,  95,  92);   // panel / sidebar
        public static readonly Color BgCard    = Color.FromArgb(24, 112, 108);   // card / nội dung
        public static readonly Color BgInput   = Color.FromArgb(16,  85,  82);   // ô nhập liệu
        public static readonly Color BgHover   = Color.FromArgb(32, 128, 124);   // hover / focus

        // ── Màu nhấn (accent sáng trên nền teal) ─────────────────────────────
        public static readonly Color Primary      = Color.FromArgb(100, 235, 228); // teal sáng
        public static readonly Color PrimaryDark  = Color.FromArgb(0,   185, 178); // pressed
        public static readonly Color PrimaryLight = Color.FromArgb(10,   60,  58); // teal rất tối

        // ── Chữ ───────────────────────────────────────────────────────────────
        public static readonly Color TextMain  = Color.FromArgb(232, 250, 248); // chữ chính (gần trắng)
        public static readonly Color TextSub   = Color.FromArgb(168, 218, 215); // chữ phụ (teal nhạt)
        public static readonly Color TextHint  = Color.FromArgb(110, 168, 165); // placeholder

        // ── Viền ──────────────────────────────────────────────────────────────
        public static readonly Color BorderColor = Color.FromArgb(28, 118, 114);

        // ── Compat ────────────────────────────────────────────────────────────
        public static readonly Color BgWhite    = BgMain;
        public static readonly Color TextDark   = TextMain;
        public static readonly Color TextMuted  = TextSub;

        // ── Màu nút (đủ tương phản trên nền teal) ─────────────────────────────
        public static readonly Color BtnBlue   = Color.FromArgb(50,  130, 255);
        public static readonly Color BtnGreen  = Color.FromArgb(45,  195,  85);
        public static readonly Color BtnRed    = Color.FromArgb(235,  70,  85);
        public static readonly Color BtnOrange = Color.FromArgb(248, 148,  55);
        public static readonly Color BtnPurple = Color.FromArgb(148, 100, 225);
        public static readonly Color BtnGold   = Color.FromArgb(232, 182,  55);
        public static readonly Color BtnTeal   = Color.FromArgb(0,   155, 150);
        public static readonly Color BtnDark   = Color.FromArgb(10,   55,  54);
        public static readonly Color BtnGray   = Color.FromArgb(85,  138, 135);

        // ── Font ──────────────────────────────────────────────────────────────
        public static Font Fn(float size = 8.5f, bool bold = false)
            => new Font("Segoe UI", size, bold ? FontStyle.Bold : FontStyle.Regular);

        // =====================================================================
        // API CÔNG KHAI
        // =====================================================================

        /// <summary>Áp dụng dark style cho toàn bộ controls trong container.</summary>
        public static void Apply(Control.ControlCollection controls,
                                 Dictionary<string, Color>  btnColors = null)
        {
            foreach (Control c in controls)
                ApplyControl(c, btnColors);
        }

        /// <summary>Áp dụng dark style cho DataGridView.</summary>
        public static void StyleDgv(DataGridView dgv)
        {
            if (dgv == null) return;
            try
            {
                dgv.BackgroundColor = BgMain;
                dgv.BorderStyle     = BorderStyle.None;
                dgv.GridColor       = BorderColor;
                dgv.Font            = Fn(8.5f);
                dgv.RowHeadersVisible              = false;
                dgv.EnableHeadersVisualStyles      = false;
                dgv.ColumnHeadersBorderStyle       = DataGridViewHeaderBorderStyle.Single;
                dgv.ColumnHeadersHeight            = 28;
                dgv.ColumnHeadersHeightSizeMode    =
                    DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

                dgv.ColumnHeadersDefaultCellStyle.BackColor          = BgPanel;
                dgv.ColumnHeadersDefaultCellStyle.ForeColor          = Primary;
                dgv.ColumnHeadersDefaultCellStyle.Font               = Fn(8.5f, true);
                dgv.ColumnHeadersDefaultCellStyle.SelectionBackColor = BgHover;
                dgv.ColumnHeadersDefaultCellStyle.Alignment          =
                    DataGridViewContentAlignment.MiddleLeft;
                dgv.ColumnHeadersDefaultCellStyle.Padding            = new Padding(4, 0, 0, 0);

                dgv.DefaultCellStyle.BackColor          = BgMain;
                dgv.DefaultCellStyle.ForeColor          = TextMain;
                dgv.DefaultCellStyle.SelectionBackColor = PrimaryDark;
                dgv.DefaultCellStyle.SelectionForeColor = TextMain;
                dgv.DefaultCellStyle.Font               = Fn(8.5f);
                dgv.DefaultCellStyle.Padding            = new Padding(2, 0, 2, 0);

                dgv.AlternatingRowsDefaultCellStyle.BackColor = BgCard;
                dgv.AlternatingRowsDefaultCellStyle.ForeColor = TextMain;
            }
            catch { }
        }

        /// <summary>Áp dụng dark style cho một Button.</summary>
        public static void StyleBtn(Button btn, Color bg,
                                    float fontSize = 8.5f, bool bold = true)
        {
            if (btn == null) return;
            try
            {
                btn.FlatStyle   = FlatStyle.Flat;
                btn.FlatAppearance.BorderSize     = 0;
                btn.FlatAppearance.MouseOverBackColor =
                    Color.FromArgb(Math.Min(bg.R + 25, 255),
                                   Math.Min(bg.G + 25, 255),
                                   Math.Min(bg.B + 25, 255));
                btn.BackColor   = bg;
                btn.ForeColor   = TextMain;
                btn.Font        = Fn(fontSize, bold);
                btn.Cursor      = Cursors.Hand;
                btn.UseVisualStyleBackColor = false;
            }
            catch { }
        }

        // =====================================================================
        // NỘI BỘ
        // =====================================================================
        private static void ApplyControl(Control c, Dictionary<string, Color> btnColors)
        {
            if (c == null) return;
            try
            {
                // Font chung
                if (!(c is DataGridView) && !(c is RichTextBox))
                    try { c.Font = Fn(8.5f); } catch { }

                if (c is Button btn)
                {
                    var color = BtnTeal;
                    if (btnColors != null && btnColors.ContainsKey(btn.Name))
                        color = btnColors[btn.Name];
                    StyleBtn(btn, color);
                }
                else if (c is DataGridView dgv)
                {
                    StyleDgv(dgv);
                }
                else if (c is GroupBox gb)
                {
                    gb.ForeColor = Primary;
                    gb.BackColor = BgPanel;
                    gb.Font      = Fn(8.5f, true);
                    Apply(gb.Controls, btnColors);
                }
                else if (c is Panel pnl)
                {
                    // Chỉ đổi bg của panel gốc / chứa form — panel toolbar/header giữ nguyên
                    if (pnl.BackColor == Color.White ||
                        pnl.BackColor == SystemColors.Control ||
                        pnl.BackColor == SystemColors.Window)
                        pnl.BackColor = BgPanel;
                    Apply(pnl.Controls, btnColors);
                }
                else if (c is Label lbl)
                {
                    lbl.Font = Fn(8.5f);
                    // Giữ màu đỏ (dấu *), chỉ đổi màu đen/mặc định → TextMain
                    if (lbl.ForeColor == SystemColors.ControlText ||
                        lbl.ForeColor == Color.Black              ||
                        lbl.ForeColor == Color.FromArgb(0, 0, 0) ||
                        lbl.ForeColor == Color.White)
                        lbl.ForeColor = TextMain;
                    // Label trên nền sáng → đổi luôn BackColor
                    if (lbl.BackColor == Color.White ||
                        lbl.BackColor == SystemColors.Control)
                        lbl.BackColor = Color.Transparent;
                }
                else if (c is TextBox tb)
                {
                    tb.Font        = Fn(8.5f);
                    tb.BorderStyle = BorderStyle.FixedSingle;
                    tb.BackColor   = BgInput;
                    tb.ForeColor   = TextMain;
                }
                else if (c is ComboBox cb)
                {
                    cb.Font      = Fn(8.5f);
                    cb.BackColor = BgInput;
                    cb.ForeColor = TextMain;
                    cb.FlatStyle = FlatStyle.Flat;
                }
                else if (c is DateTimePicker dtp)
                {
                    dtp.Font        = Fn(8.5f);
                    dtp.CalendarFont = Fn(8.5f);
                    dtp.CalendarMonthBackground = BgCard;
                    dtp.CalendarForeColor        = TextMain;
                    dtp.CalendarTitleBackColor   = BgPanel;
                    dtp.CalendarTitleForeColor   = Primary;
                }
                else if (c is RichTextBox rtb)
                {
                    if (rtb.BackColor == Color.White ||
                        rtb.BackColor == SystemColors.Window)
                    {
                        rtb.BackColor = BgCard;
                        rtb.ForeColor = TextMain;
                    }
                }
                else if (c is TabControl tc)
                {
                    tc.Font = Fn(9f, true);
                    foreach (TabPage tp in tc.TabPages)
                    {
                        tp.BackColor = BgMain;
                        Apply(tp.Controls, btnColors);
                    }
                }
                else if (c is TabPage tp2)
                {
                    tp2.BackColor = BgMain;
                    Apply(tp2.Controls, btnColors);
                }

                // Đệ quy các container khác
                if (c.HasChildren
                    && !(c is GroupBox) && !(c is Panel)
                    && !(c is TabControl) && !(c is TabPage))
                    Apply(c.Controls, btnColors);
            }
            catch { }
        }
    }
}
