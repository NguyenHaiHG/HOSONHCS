using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace HOSONHCS
{
    /// <summary>
    /// 🎨 MODERN UI COMPONENTS - Toast, Loading, Progress
    /// Components đẹp, nhẹ, dễ dùng cho app
    /// </summary>

    // ========================================================================
    // 🔔 TOAST NOTIFICATION - Thông báo đẹp thay MessageBox
    // ========================================================================
    /// <summary>
    /// Toast notification - Thông báo nhỏ tự động biến mất
    /// Giống như thông báo trên điện thoại
    /// </summary>
    public class ToastNotification : Form
    {
        private Timer fadeInTimer;
        private Timer displayTimer;
        private Timer fadeOutTimer;
        
        /// <summary>
        /// Tạo toast notification với message và type
        /// </summary>
        /// <param name="message">Nội dung thông báo</param>
        /// <param name="type">Loại: success, error, warning, info</param>
        /// <param name="duration">Thời gian hiển thị (ms), mặc định 3000ms</param>
        public ToastNotification(string message, ToastType type = ToastType.Info, int duration = 3000)
        {
            // Cấu hình form
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.Manual;
            this.Size = new Size(350, 80);
            this.ShowInTaskbar = false;
            this.TopMost = true;
            
            // Màu background theo type
            this.BackColor = GetColorByType(type);
            
            // Bo góc
            this.Region = GetRoundedRegion(this.ClientRectangle, 10);
            
            // Vị trí: góc phải dưới màn hình
            var screen = Screen.PrimaryScreen.WorkingArea;
            this.Location = new Point(
                screen.Width - this.Width - 20,
                screen.Height - this.Height - 20
            );
            
            // Icon + Message label
            var iconLabel = new Label
            {
                Text = GetIconByType(type),
                Font = new Font("Segoe UI", 20F),
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(15, 25)
            };
            
            var messageLabel = new Label
            {
                Text = message,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                ForeColor = Color.White,
                AutoSize = false,
                Size = new Size(280, 60),
                Location = new Point(60, 10),
                TextAlign = ContentAlignment.MiddleLeft
            };
            
            this.Controls.Add(iconLabel);
            this.Controls.Add(messageLabel);
            
            // Close button (X)
            var closeBtn = new Label
            {
                Text = "✕",
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(this.Width - 30, 5),
                Cursor = Cursors.Hand
            };
            closeBtn.Click += (s, e) => this.Close();
            this.Controls.Add(closeBtn);
            
            // Animations
            SetupAnimations(duration);
        }
        
        private Color GetColorByType(ToastType type)
        {
            switch (type)
            {
                case ToastType.Success: return AppTheme.MacGreen;
                case ToastType.Error: return AppTheme.MacRed;
                case ToastType.Warning: return AppTheme.MacOrange;
                case ToastType.Info: return AppTheme.MacBlue;
                default: return AppTheme.MacBlue;
            }
        }
        
        private string GetIconByType(ToastType type)
        {
            switch (type)
            {
                case ToastType.Success: return "✓";
                case ToastType.Error: return "✕";
                case ToastType.Warning: return "⚠";
                case ToastType.Info: return "ℹ";
                default: return "ℹ";
            }
        }
        
        private Region GetRoundedRegion(Rectangle rect, int radius)
        {
            var path = new GraphicsPath();
            path.AddArc(rect.X, rect.Y, radius * 2, radius * 2, 180, 90);
            path.AddArc(rect.Right - radius * 2, rect.Y, radius * 2, radius * 2, 270, 90);
            path.AddArc(rect.Right - radius * 2, rect.Bottom - radius * 2, radius * 2, radius * 2, 0, 90);
            path.AddArc(rect.X, rect.Bottom - radius * 2, radius * 2, radius * 2, 90, 90);
            path.CloseFigure();
            return new Region(path);
        }
        
        private void SetupAnimations(int duration)
        {
            // Fade in (0 → 1 trong 200ms)
            this.Opacity = 0;
            fadeInTimer = new Timer { Interval = 10 };
            fadeInTimer.Tick += (s, e) =>
            {
                if (this.Opacity < 1)
                    this.Opacity += 0.05;
                else
                    fadeInTimer.Stop();
            };
            fadeInTimer.Start();
            
            // Display duration
            displayTimer = new Timer { Interval = duration };
            displayTimer.Tick += (s, e) =>
            {
                displayTimer.Stop();
                // Start fade out
                fadeOutTimer.Start();
            };
            displayTimer.Start();
            
            // Fade out (1 → 0 trong 200ms)
            fadeOutTimer = new Timer { Interval = 10 };
            fadeOutTimer.Tick += (s, e) =>
            {
                if (this.Opacity > 0)
                    this.Opacity -= 0.05;
                else
                {
                    fadeOutTimer.Stop();
                    this.Close();
                }
            };
        }
    }
    
    public enum ToastType
    {
        Success,
        Error,
        Warning,
        Info
    }

    // ========================================================================
    // ⏳ LOADING INDICATOR - Spinner khi export Word
    // ========================================================================
    /// <summary>
    /// Loading overlay với spinner xoay
    /// Hiển thị khi đang xử lý task nặng (export Word, etc.)
    /// </summary>
    public class LoadingOverlay : Form
    {
        private Timer spinnerTimer;
        private int angle = 0;
        private string message;
        
        public LoadingOverlay(string message = "Đang xử lý...")
        {
            this.message = message;
            
            // Cấu hình form
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Size = new Size(200, 150);
            this.BackColor = Color.FromArgb(230, 0, 0, 0); // Semi-transparent dark
            this.ShowInTaskbar = false;
            this.TopMost = true;
            
            // Spinner animation
            spinnerTimer = new Timer { Interval = 50 }; // 20 FPS
            spinnerTimer.Tick += (s, e) =>
            {
                angle = (angle + 10) % 360;
                this.Invalidate(); // Trigger repaint
            };
            spinnerTimer.Start();
        }
        
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            
            // Vẽ spinner (arc xoay)
            int centerX = this.Width / 2;
            int centerY = (this.Height / 2) - 10;
            int radius = 25;
            
            e.Graphics.TranslateTransform(centerX, centerY);
            e.Graphics.RotateTransform(angle);
            
            using (var pen = new Pen(Color.White, 4))
            {
                e.Graphics.DrawArc(pen, -radius, -radius, radius * 2, radius * 2, 0, 300);
            }
            
            e.Graphics.ResetTransform();
            
            // Vẽ message
            using (var brush = new SolidBrush(Color.White))
            {
                var format = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                };
                
                e.Graphics.DrawString(message,
                    new Font("Segoe UI", 10F),
                    brush,
                    new RectangleF(0, this.Height - 40, this.Width, 30),
                    format);
            }
        }
        
        public new void Close()
        {
            spinnerTimer?.Stop();
            spinnerTimer?.Dispose();
            base.Close();
        }
    }

    // ========================================================================
    // 📊 MODERN PROGRESS BAR - Progress bar với % text
    // ========================================================================
    /// <summary>
    /// Progress bar hiện đại với % hiển thị
    /// Dùng khi export nhiều files
    /// </summary>
    public class ModernProgressBar : ProgressBar
    {
        public ModernProgressBar()
        {
            this.SetStyle(ControlStyles.UserPaint, true);
            this.Height = 30;
        }
        
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            
            // Background
            using (var bgBrush = new SolidBrush(AppTheme.MacInputBackground))
            {
                e.Graphics.FillRectangle(bgBrush, 0, 0, this.Width, this.Height);
            }
            
            // Progress bar
            if (this.Value > 0)
            {
                int progressWidth = (int)((double)this.Value / this.Maximum * this.Width);
                
                using (var progressBrush = new LinearGradientBrush(
                    new Rectangle(0, 0, progressWidth, this.Height),
                    AppTheme.MacBlue,
                    AppTheme.MacBlueDark,
                    LinearGradientMode.Horizontal))
                {
                    e.Graphics.FillRectangle(progressBrush, 0, 0, progressWidth, this.Height);
                }
            }
            
            // Border
            using (var borderPen = new Pen(AppTheme.MacBorder, 1))
            {
                e.Graphics.DrawRectangle(borderPen, 0, 0, this.Width - 1, this.Height - 1);
            }
            
            // Text %
            string percentText = $"{(int)((double)this.Value / this.Maximum * 100)}%";
            using (var textBrush = new SolidBrush(AppTheme.MacTextPrimary))
            {
                var format = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                };
                
                e.Graphics.DrawString(percentText,
                    new Font("Segoe UI", 10F, FontStyle.Bold),
                    textBrush,
                    new RectangleF(0, 0, this.Width, this.Height),
                    format);
            }
        }
    }
}
