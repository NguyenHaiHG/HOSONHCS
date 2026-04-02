using System;
using System.IO;
using System.Windows.Forms;

namespace HOSONHCS
{
    public partial class Form1
    {
        // Đặt file .rtf vào cùng thư mục với file .exe
        private const string DocFileName = "mau10c.rtf";

        private void InitializeTab7()
        {
            if (tabPage7 == null) return;

            var rtb = new RichTextBox
            {
                Dock      = DockStyle.Fill,
                ReadOnly  = true,
                BackColor = System.Drawing.Color.White,
                BorderStyle = BorderStyle.None,
                ScrollBars = RichTextBoxScrollBars.Vertical
            };

            tabPage7.Controls.Add(rtb);

            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, DocFileName);
            if (File.Exists(path))
                rtb.LoadFile(path, RichTextBoxStreamType.RichText);
        }
    }
}
