using System;
using System.Drawing;
using System.Windows.Forms;

namespace HOSONHCS
{
    /// <summary>
    /// Popup hiển thị thông báo ngày lễ / sự kiện đặc biệt
    /// Thiết kế theo theme Teal của UIStyler, không cần file Designer
    /// </summary>
    public class HolidayNoticeForm : Form
    {
        private readonly HolidayNoticeMessage _notice;

        // ── kéo form bằng header ───────────────────────────────────────────────
        private bool   _dragging;
        private Point  _dragStart;

        public HolidayNoticeForm(HolidayNoticeMessage notice)
        {
            _notice = notice;
            BuildUI();
        }

        private void BuildUI()
        {
            const int W  = 500;
            const int HH = 105;   // header height
            const int SH = 2;     // separator height
            const int FH = 58;    // footer height
            const int CH = 200;   // content height
            const int H  = HH + SH + CH + FH;

            // ── FORM ──────────────────────────────────────────────────────────
            this.Text            = "Thông Báo";
            this.Size            = new Size(W, H);
            this.StartPosition   = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.None;
            this.BackColor       = UIStyler.BgMain;
            this.TopMost         = true;
            this.KeyPreview      = true;
            this.KeyDown        += (s, e) => { if (e.KeyCode == Keys.Escape || e.KeyCode == Keys.Return) Close(); };

            // ── HEADER ────────────────────────────────────────────────────────
            var pHeader = new Panel
            {
                Bounds    = new Rectangle(0, 0, W, HH),
                BackColor = UIStyler.BgPanel
            };

            var lblEmoji = new Label
            {
                Text      = _notice.Emoji ?? "🎉",
                Font      = new Font("Segoe UI Emoji", 38, FontStyle.Regular, GraphicsUnit.Point),
                ForeColor = UIStyler.Primary,
                Bounds    = new Rectangle(16, 14, 74, 74),
                TextAlign = ContentAlignment.MiddleCenter
            };

            var lblTitle = new Label
            {
                Text      = _notice.Title ?? "Thông báo",
                Font      = UIStyler.Fn(13.5f, bold: true),
                ForeColor = UIStyler.Primary,
                Bounds    = new Rectangle(98, 14, W - 114, 74),
                TextAlign = ContentAlignment.MiddleLeft
            };

            var lblDate = new Label
            {
                Text      = $"📅  {_notice.StartDate}  →  {_notice.EndDate}",
                Font      = UIStyler.Fn(7.5f),
                ForeColor = UIStyler.TextHint,
                Bounds    = new Rectangle(0, HH - 22, W - 10, 20),
                TextAlign = ContentAlignment.MiddleRight
            };

            // nút X đóng popup
            var btnClose = new Button
            {
                Text      = "✕",
                Font      = UIStyler.Fn(9, bold: true),
                ForeColor = UIStyler.TextSub,
                BackColor = UIStyler.BgPanel,
                FlatStyle = FlatStyle.Flat,
                Bounds    = new Rectangle(W - 36, 4, 30, 26),
                Cursor    = Cursors.Hand
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += (s, e) => Close();

            pHeader.Controls.AddRange(new Control[] { lblEmoji, lblTitle, lblDate, btnClose });
            pHeader.MouseDown += Header_MouseDown;
            pHeader.MouseMove += Header_MouseMove;
            pHeader.MouseUp   += (s, e) => _dragging = false;
            lblEmoji.MouseDown += Header_MouseDown;
            lblEmoji.MouseMove += Header_MouseMove;
            lblEmoji.MouseUp   += (s, e) => _dragging = false;
            lblTitle.MouseDown += Header_MouseDown;
            lblTitle.MouseMove += Header_MouseMove;
            lblTitle.MouseUp   += (s, e) => _dragging = false;

            // ── SEPARATOR ─────────────────────────────────────────────────────
            var pSep = new Panel
            {
                Bounds    = new Rectangle(0, HH, W, SH),
                BackColor = UIStyler.Primary
            };

            // ── CONTENT ───────────────────────────────────────────────────────
            var pContent = new Panel
            {
                Bounds    = new Rectangle(0, HH + SH, W, CH),
                BackColor = UIStyler.BgCard,
                Padding   = new Padding(20)
            };

            var rtbContent = new RichTextBox
            {
                Text        = _notice.Content ?? "",
                Font        = UIStyler.Fn(10.5f),
                ForeColor   = UIStyler.TextMain,
                BackColor   = UIStyler.BgCard,
                BorderStyle = BorderStyle.None,
                ReadOnly    = true,
                ScrollBars  = RichTextBoxScrollBars.Vertical,
                Bounds      = new Rectangle(20, 14, W - 40, CH - 28)
            };
            pContent.Controls.Add(rtbContent);

            // ── FOOTER ────────────────────────────────────────────────────────
            var pFooter = new Panel
            {
                Bounds    = new Rectangle(0, HH + SH + CH, W, FH),
                BackColor = UIStyler.BgPanel
            };

            var btnOk = new Button
            {
                Text      = "✔  Đã Hiểu",
                Font      = UIStyler.Fn(10, bold: true),
                ForeColor = Color.White,
                BackColor = UIStyler.BtnGreen,
                FlatStyle = FlatStyle.Flat,
                Size      = new Size(140, 36),
                Cursor    = Cursors.Hand
            };
            btnOk.FlatAppearance.BorderSize = 0;
            btnOk.Location = new Point((W - btnOk.Width) / 2, (FH - btnOk.Height) / 2);
            btnOk.Click   += (s, e) => Close();
            pFooter.Controls.Add(btnOk);

            // ── GẮN VÀO FORM ─────────────────────────────────────────────────
            this.Controls.AddRange(new Control[] { pHeader, pSep, pContent, pFooter });
        }

        // ── KÉO FORM ──────────────────────────────────────────────────────────
        private void Header_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                _dragging  = true;
                _dragStart = e.Location;
                // Điều chỉnh toạ độ tương đối với form
                if (sender is Control ctrl && ctrl != this)
                    _dragStart = ctrl.PointToScreen(e.Location) - (Size)this.Location;
            }
        }

        private void Header_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_dragging) return;
            Point screenPos = (sender is Control ctrl && ctrl != this)
                ? ctrl.PointToScreen(e.Location)
                : this.PointToScreen(e.Location);
            this.Location = screenPos - (Size)_dragStart;
        }
    }
}
