using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace HOSONHCS
{
    internal static class ResponsiveLayoutManager
    {
        private const int DesignedWidth = 1526;
        private const int DesignedHeight = 848;
        private const int ScreenMargin = 24;
        private const int MinZoom = 60;
        private const int MaxZoom = 130;
        private const int DefaultZoom = 100;

        private static readonly Dictionary<Control, LayoutInfo> OriginalLayouts = new Dictionary<Control, LayoutInfo>();
        private static bool captured;

        public static void Apply(Form form, TabControl tabControl)
        {
            if (form == null)
                return;

            EnableFormResize(form);
            ConfigureTabScrolling(tabControl);
            CaptureOriginalLayouts(tabControl);
            AddZoomBar(form, tabControl);

            form.Shown -= Form_Shown;
            form.Shown += Form_Shown;
        }

        private static void EnableFormResize(Form form)
        {
            form.FormBorderStyle = FormBorderStyle.Sizable;
            form.MaximizeBox = true;
            form.SizeGripStyle = SizeGripStyle.Show;

            Rectangle workArea = Screen.FromControl(form).WorkingArea;
            form.MinimumSize = new Size(Math.Min(900, workArea.Width - ScreenMargin), Math.Min(560, workArea.Height - ScreenMargin));
        }

        private static void Form_Shown(object sender, EventArgs e)
        {
            FitFormToScreen(sender as Form);
        }

        private static void FitFormToScreen(Form form)
        {
            if (form == null)
                return;

            Rectangle workArea = Screen.FromControl(form).WorkingArea;
            int maxWidth = Math.Max(640, workArea.Width - ScreenMargin);
            int maxHeight = Math.Max(480, workArea.Height - ScreenMargin);

            if (DesignedWidth > maxWidth || DesignedHeight > maxHeight)
            {
                form.WindowState = FormWindowState.Maximized;
                return;
            }

            form.Size = new Size(DesignedWidth, DesignedHeight);
            form.Location = new Point(
                workArea.Left + (workArea.Width - form.Width) / 2,
                workArea.Top + (workArea.Height - form.Height) / 2);
        }

        private static void ConfigureTabScrolling(TabControl tabControl)
        {
            if (tabControl == null)
                return;

            tabControl.Dock = DockStyle.Fill;

            foreach (TabPage tabPage in tabControl.TabPages)
            {
                tabPage.AutoScroll = true;
                tabPage.AutoScrollMinSize = GetContentSize(tabPage);
            }
        }

        private static void AddZoomBar(Form form, TabControl tabControl)
        {
            Panel existing = form.Controls["pnlResponsiveZoom"] as Panel;
            if (existing != null)
                return;

            var panel = new Panel
            {
                Name = "pnlResponsiveZoom",
                Dock = DockStyle.Bottom,
                Height = 34,
                BackColor = SystemColors.Control
            };

            var label = new Label
            {
                Name = "lblResponsiveZoom",
                AutoSize = true,
                Location = new Point(10, 9),
                Text = "Thu/phóng giao diện: 100%"
            };

            var trackBar = new TrackBar
            {
                Name = "trkResponsiveZoom",
                Minimum = MinZoom,
                Maximum = MaxZoom,
                Value = DefaultZoom,
                TickFrequency = 10,
                SmallChange = 5,
                LargeChange = 10,
                Width = 230,
                Height = 30,
                Location = new Point(170, 2)
            };

            var fitButton = new Button
            {
                Name = "btnResponsiveFit",
                Text = "Vừa màn hình",
                Width = 110,
                Height = 25,
                Location = new Point(410, 4)
            };

            trackBar.ValueChanged += (s, e) =>
            {
                label.Text = "Thu/phóng giao diện: " + trackBar.Value + "%";
                ApplyZoom(tabControl, trackBar.Value / 100f);
            };

            fitButton.Click += (s, e) =>
            {
                trackBar.Value = CalculateFitZoom(form);
            };

            panel.Controls.Add(label);
            panel.Controls.Add(trackBar);
            panel.Controls.Add(fitButton);
            form.Controls.Add(panel);
            panel.BringToFront();
        }

        private static int CalculateFitZoom(Form form)
        {
            Rectangle workArea = Screen.FromControl(form).WorkingArea;
            float widthScale = (workArea.Width - ScreenMargin) / (float)DesignedWidth;
            float heightScale = (workArea.Height - ScreenMargin - 34) / (float)DesignedHeight;
            int percent = (int)Math.Floor(Math.Min(widthScale, heightScale) * 100f);
            return Math.Max(MinZoom, Math.Min(MaxZoom, percent));
        }

        private static void CaptureOriginalLayouts(TabControl tabControl)
        {
            if (tabControl == null || captured)
                return;

            foreach (TabPage tabPage in tabControl.TabPages)
            {
                CaptureChildren(tabPage);
            }

            captured = true;
        }

        private static void CaptureChildren(Control parent)
        {
            foreach (Control control in parent.Controls)
            {
                if (!OriginalLayouts.ContainsKey(control))
                    OriginalLayouts.Add(control, new LayoutInfo(control.Bounds, control.Font));

                CaptureChildren(control);
            }
        }

        private static void ApplyZoom(TabControl tabControl, float scale)
        {
            if (tabControl == null)
                return;

            tabControl.SuspendLayout();
            foreach (TabPage tabPage in tabControl.TabPages)
            {
                tabPage.SuspendLayout();
                ScaleChildren(tabPage, scale);
                tabPage.AutoScrollMinSize = GetContentSize(tabPage);
                tabPage.ResumeLayout();
            }
            tabControl.ResumeLayout();
        }

        private static void ScaleChildren(Control parent, float scale)
        {
            foreach (Control control in parent.Controls)
            {
                LayoutInfo info;
                if (OriginalLayouts.TryGetValue(control, out info))
                {
                    control.Bounds = new Rectangle(
                        (int)Math.Round(info.Bounds.X * scale),
                        (int)Math.Round(info.Bounds.Y * scale),
                        Math.Max(1, (int)Math.Round(info.Bounds.Width * scale)),
                        Math.Max(1, (int)Math.Round(info.Bounds.Height * scale)));

                    float fontSize = Math.Max(6f, info.Font.Size * scale);
                    if (Math.Abs(control.Font.Size - fontSize) > 0.1f)
                        control.Font = new Font(info.Font.FontFamily, fontSize, info.Font.Style);
                }

                ScaleChildren(control, scale);
            }
        }

        private static Size GetContentSize(TabPage tabPage)
        {
            int right = 0;
            int bottom = 0;

            foreach (Control control in tabPage.Controls)
            {
                right = Math.Max(right, control.Right);
                bottom = Math.Max(bottom, control.Bottom);
            }

            return new Size(right + 24, bottom + 24);
        }

        private sealed class LayoutInfo
        {
            public LayoutInfo(Rectangle bounds, Font font)
            {
                Bounds = bounds;
                Font = font;
            }

            public Rectangle Bounds { get; private set; }
            public Font Font { get; private set; }
        }
    }
}
