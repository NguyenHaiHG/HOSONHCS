using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

namespace HOSONHCS
{
    /// <summary>
    /// 🎨 THEME MANAGER - Quản lý tập trung tất cả theme styling
    /// Tách riêng khỏi Form1/Form2 để dễ maintain và customize
    /// </summary>
    public static class ThemeManager
    {
        // ========================================================================
        // APPLY THEME CHO FORM
        // ========================================================================
        
        /// <summary>
        /// Áp dụng MacBook theme cho toàn bộ form
        /// </summary>
        public static void ApplyMacBookTheme(Form form)
        {
            if (form == null) return;
            
            try
            {
                // Form background
                form.BackColor = AppTheme.MacBackground;

                // Apply theme cho tất cả controls
                ApplyThemeToAllControls(form);
                
                // Bo góc cho tất cả controls (delay để đảm bảo layout xong)
                form.BeginInvoke((Action)(() =>
                {
                    ApplyRoundedCornersToAllControls(form);
                }));
            }
            catch { }
        }

        // ========================================================================
        // STYLE CÁC CONTROLS
        // ========================================================================

        private static void ApplyThemeToAllControls(Control container)
        {
            try
            {
                foreach (Control ctrl in GetAllControls(container))
                {
                    // TabControl
                    if (ctrl is TabControl tabControl)
                    {
                        tabControl.Font = new Font("Segoe UI", 9.5F, FontStyle.Regular);
                        
                        // Style all tab pages
                        foreach (TabPage page in tabControl.TabPages)
                        {
                            page.BackColor = AppTheme.MacBackground;
                        }
                    }
                    // GroupBox
                    else if (ctrl is GroupBox groupBox)
                    {
                        groupBox.BackColor = AppTheme.MacCardBackground;
                        groupBox.ForeColor = AppTheme.MacTextPrimary;
                        groupBox.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
                    }
                    // Label (except special labels like marquee headers)
                    else if (ctrl is Label label)
                    {
                        // Skip labels named "label14" or with special tags
                        if (label.Name != "label14" && label.Tag?.ToString() != "marquee")
                        {
                            label.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
                            label.ForeColor = Color.Black;
                        }
                    }
                    // TextBox
                    else if (ctrl is TextBox textBox)
                    {
                        StyleTextBox(textBox);
                    }
                    // RichTextBox
                    else if (ctrl is RichTextBox richTextBox)
                    {
                        StyleRichTextBox(richTextBox);
                    }
                    // ComboBox
                    else if (ctrl is ComboBox comboBox)
                    {
                        StyleComboBox(comboBox);
                    }
                    // DataGridView
                    else if (ctrl is DataGridView dataGridView)
                    {
                        StyleDataGridView(dataGridView);
                    }
                }
            }
            catch { }
        }

        // ========================================================================
        // BUTTON STYLING
        // ========================================================================

        /// <summary>
        /// Style button theo MacOS với màu và icon
        /// </summary>
        public static void StyleButton(Button btn, Color color, string icon = "", string text = "")
        {
            if (btn == null) return;

            try
            {
                btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderSize = 0;
                btn.BackColor = color;
                btn.ForeColor = Color.White;
                btn.Font = new Font("Segoe UI", 9.5F, FontStyle.Regular);
                btn.Cursor = Cursors.Hand;
                btn.Height = 36;
                
                // ========== FIX: TỰ ĐỘNG TĂNG WIDTH ĐỂ FIT TEXT ==========
                // Đo text để tính width cần thiết
                if (!string.IsNullOrEmpty(icon) && !string.IsNullOrEmpty(text))
                {
                    btn.Text = $"{icon} {text}";
                    
                    // Tăng width để đủ chỗ cho text
                    using (var g = btn.CreateGraphics())
                    {
                        var textSize = g.MeasureString(btn.Text, btn.Font);
                        int requiredWidth = (int)textSize.Width + 30; // +30 cho padding
                        if (btn.Width < requiredWidth)
                        {
                            btn.Width = requiredWidth;
                        }
                    }
                }
                else if (!string.IsNullOrEmpty(text))
                {
                    btn.Text = text;
                }

                // Hover effect
                Color originalColor = color;
                Color hoverColor = GetHoverColor(color);

                btn.MouseEnter += (s, e) => { btn.BackColor = hoverColor; };
                btn.MouseLeave += (s, e) => { btn.BackColor = originalColor; };
            }
            catch { }
        }

        private static Color GetHoverColor(Color color)
        {
            if (color.ToArgb() == AppTheme.MacGreen.ToArgb()) return AppTheme.MacGreenHover;
            if (color.ToArgb() == AppTheme.MacRed.ToArgb()) return AppTheme.MacRedHover;
            if (color.ToArgb() == AppTheme.MacOrange.ToArgb()) return AppTheme.MacOrangeHover;
            if (color.ToArgb() == AppTheme.MacTeal.ToArgb()) return AppTheme.MacTealHover;
            if (color.ToArgb() == AppTheme.MacPurple.ToArgb()) return AppTheme.MacPurpleHover;
            return AppTheme.MacBlueHover;
        }

        // ========================================================================
        // TEXTBOX STYLING
        // ========================================================================

        private static void StyleTextBox(TextBox txt)
        {
            try
            {
                txt.Font = new Font("Segoe UI", 8.5F, FontStyle.Regular);
                txt.BackColor = AppTheme.MacInputBackground;
                txt.ForeColor = AppTheme.MacTextPrimary;
                txt.BorderStyle = BorderStyle.None;

                // Focus effect
                txt.Enter += (s, e) => { ((TextBox)s).BackColor = AppTheme.MacInputBackgroundFocus; };
                txt.Leave += (s, e) => { ((TextBox)s).BackColor = AppTheme.MacInputBackground; };
            }
            catch { }
        }

        private static void StyleRichTextBox(RichTextBox rtb)
        {
            try
            {
                rtb.Font = new Font("Segoe UI", 8.5F, FontStyle.Regular);
                rtb.BackColor = AppTheme.MacInputBackground;
                rtb.ForeColor = AppTheme.MacTextPrimary;
                rtb.BorderStyle = BorderStyle.None;

                // Focus effect
                rtb.Enter += (s, e) => { ((RichTextBox)s).BackColor = AppTheme.MacInputBackgroundFocus; };
                rtb.Leave += (s, e) => { ((RichTextBox)s).BackColor = AppTheme.MacInputBackground; };
            }
            catch { }
        }

        // ========================================================================
        // COMBOBOX STYLING
        // ========================================================================

        private static void StyleComboBox(ComboBox cb)
        {
            try
            {
                cb.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
                cb.BackColor = AppTheme.MacInputBackground;
                cb.ForeColor = Color.Black;
                cb.FlatStyle = FlatStyle.Popup;  // Popup để text không bị ẩn
                
                if (cb.Height < 24) cb.Height = 24;
            }
            catch { }
        }

        // ========================================================================
        // DATAGRIDVIEW STYLING
        // ========================================================================

        private static void StyleDataGridView(DataGridView dgv)
        {
            try
            {
                dgv.BorderStyle = BorderStyle.None;
                dgv.BackgroundColor = AppTheme.MacCardBackground;
                dgv.GridColor = AppTheme.MacBorderLight;
                dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;

                dgv.DefaultCellStyle.BackColor = Color.White;
                dgv.DefaultCellStyle.ForeColor = AppTheme.MacTextPrimary;
                dgv.DefaultCellStyle.Font = new Font("Segoe UI", 8.5F);
                dgv.DefaultCellStyle.SelectionBackColor = AppTheme.MacBlue;
                dgv.DefaultCellStyle.SelectionForeColor = Color.White;
                dgv.DefaultCellStyle.Padding = new Padding(6, 3, 6, 3);

                dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(249, 249, 251);

                dgv.EnableHeadersVisualStyles = false;
                dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
                dgv.ColumnHeadersDefaultCellStyle.BackColor = AppTheme.MacHeaderGradient1;
                dgv.ColumnHeadersDefaultCellStyle.ForeColor = AppTheme.MacTextPrimary;
                dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
                dgv.ColumnHeadersDefaultCellStyle.Padding = new Padding(8, 6, 8, 6);
                dgv.ColumnHeadersHeight = 36;

                dgv.RowTemplate.Height = 32;
            }
            catch { }
        }

        // ========================================================================
        // ROUNDED CORNERS (BO GÓC)
        // ========================================================================

        private static void ApplyRoundedCornersToAllControls(Control container)
        {
            try
            {
                foreach (Control ctrl in GetAllControls(container))
                {
                    if (!ctrl.Visible || ctrl.Width <= 0 || ctrl.Height <= 0) continue;

                    if (ctrl is TextBox || ctrl is RichTextBox)
                    {
                        ApplyRoundedCorners(ctrl, AppTheme.CornerRadius);
                    }
                    else if (ctrl is ComboBox)
                    {
                        ApplyRoundedCorners(ctrl, AppTheme.CornerRadius);
                    }
                    else if (ctrl is DateTimePicker)
                    {
                        ApplyRoundedCorners(ctrl, AppTheme.CornerRadius);
                    }
                    else if (ctrl is Button)
                    {
                        ApplyRoundedCorners(ctrl, 10); // Button bo góc lớn hơn
                    }
                    else if (ctrl is Panel || ctrl is GroupBox)
                    {
                        ApplyRoundedCorners(ctrl, 12); // Panel bo góc lớn hơn
                    }
                }
            }
            catch { }
        }

        private static void ApplyRoundedCorners(Control ctrl, int radius)
        {
            try
            {
                if (ctrl == null || ctrl.Width <= 0 || ctrl.Height <= 0) return;

                int maxRadius = Math.Min(ctrl.Width, ctrl.Height) / 2;
                if (radius > maxRadius) radius = maxRadius;

                GraphicsPath path = new GraphicsPath();
                path.AddArc(0, 0, radius * 2, radius * 2, 180, 90);
                path.AddArc(ctrl.Width - radius * 2, 0, radius * 2, radius * 2, 270, 90);
                path.AddArc(ctrl.Width - radius * 2, ctrl.Height - radius * 2, radius * 2, radius * 2, 0, 90);
                path.AddArc(0, ctrl.Height - radius * 2, radius * 2, radius * 2, 90, 90);
                path.CloseFigure();

                ctrl.Region = new Region(path);

                // Handle resize
                ctrl.Resize -= Control_ResizeRounded;
                ctrl.Resize += Control_ResizeRounded;
                ctrl.Tag = $"RoundedRadius:{radius}";
            }
            catch { }
        }

        private static void Control_ResizeRounded(object sender, EventArgs e)
        {
            try
            {
                var ctrl = sender as Control;
                if (ctrl == null) return;

                int radius = AppTheme.CornerRadius;
                if (ctrl.Tag is string tag && tag.StartsWith("RoundedRadius:"))
                {
                    int.TryParse(tag.Replace("RoundedRadius:", ""), out radius);
                }

                ctrl.Resize -= Control_ResizeRounded;
                ApplyRoundedCorners(ctrl, radius);
                ctrl.Resize += Control_ResizeRounded;
            }
            catch { }
        }

        // ========================================================================
        // HELPER METHODS
        // ========================================================================

        private static IEnumerable<Control> GetAllControls(Control container)
        {
            foreach (Control ctrl in container.Controls)
            {
                yield return ctrl;
                foreach (Control child in GetAllControls(ctrl))
                {
                    yield return child;
                }
            }
        }
    }
}
