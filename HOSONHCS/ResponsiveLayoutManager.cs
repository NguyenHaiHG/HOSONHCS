using System;
using System.Drawing;
using System.Windows.Forms;

namespace HOSONHCS
{
    internal static class ResponsiveLayoutManager
    {
        private const int DesignedWidth = 1526;
        private const int DesignedHeight = 848;
        private const int ScreenMargin = 24;

        public static void Apply(Form form, TabControl tabControl)
        {
            if (form == null)
                return;

            EnableFormResize(form);
            ConfigureTabScrolling(tabControl);

            form.Shown -= Form_Shown;
            form.Shown += Form_Shown;
        }

        private static void EnableFormResize(Form form)
        {
            form.FormBorderStyle = FormBorderStyle.Sizable;
            form.MaximizeBox = true;
            form.SizeGripStyle = SizeGripStyle.Show;

            Rectangle workArea = Screen.FromControl(form).WorkingArea;
            form.MinimumSize = new Size(
                Math.Min(1000, Math.Max(640, workArea.Width - ScreenMargin)),
                Math.Min(650, Math.Max(480, workArea.Height - ScreenMargin)));
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
    }
}
