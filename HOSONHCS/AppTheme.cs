using System.Drawing;

namespace HOSONHCS
{
    public static class AppTheme
    {
        // Background colors - NHCSXH Blue Theme (Màu xanh NGÂN HÀNG rõ ràng)
        public static readonly Color MacBackground = Color.FromArgb(200, 225, 240);      // XANH NHẠT RÕ
        public static readonly Color MacCardBackground = Color.FromArgb(235, 245, 252);  // XANH RẤT NHẠT
        public static readonly Color MacPanelBackground = Color.FromArgb(220, 235, 245); // XANH PANEL
        public static readonly Color MacSidebarGray = Color.FromArgb(210, 230, 242);

        // Màu xanh chính NHCSXH
        public static readonly Color NHCSXHBlue = Color.FromArgb(0, 102, 178);          // Primary bank blue
        public static readonly Color NHCSXHBlueDark = Color.FromArgb(0, 77, 122);       // Dark blue
        public static readonly Color NHCSXHBlueLight = Color.FromArgb(30, 136, 229);    // Light blue

        // Màu ô nhập liệu - sắc xanh nhạt
        public static readonly Color MacInputBackground = Color.FromArgb(230, 240, 248);
        public static readonly Color MacInputBackgroundFocus = Color.FromArgb(255, 255, 255);

        // Màu nút bấm
        public static readonly Color MacBlue = Color.FromArgb(0, 122, 255);
        public static readonly Color MacBlueHover = Color.FromArgb(0, 102, 204);
        public static readonly Color MacBlueDark = Color.FromArgb(0, 64, 221);

        public static readonly Color MacGreen = Color.FromArgb(52, 199, 89);
        public static readonly Color MacGreenHover = Color.FromArgb(40, 180, 70);

        public static readonly Color MacRed = Color.FromArgb(255, 59, 48);
        public static readonly Color MacRedHover = Color.FromArgb(220, 50, 40);

        public static readonly Color MacOrange = Color.FromArgb(255, 149, 0);
        public static readonly Color MacOrangeHover = Color.FromArgb(230, 130, 0);

        public static readonly Color MacPurple = Color.FromArgb(175, 82, 222);
        public static readonly Color MacPurpleHover = Color.FromArgb(150, 70, 190);

        public static readonly Color MacTeal = Color.FromArgb(90, 200, 250);
        public static readonly Color MacTealHover = Color.FromArgb(70, 180, 230);

        // Màu tiêu đề chạy (marquee) - Màu hiện đại sống động
        public static readonly Color MarqueeCyan = Color.FromArgb(0, 150, 255);        // Vibrant cyan
        public static readonly Color MarqueePurple = Color.FromArgb(138, 43, 226);    // Blue-violet
        public static readonly Color MarqueePink = Color.FromArgb(255, 20, 147);      // Deep pink
        public static readonly Color MarqueeOrange = Color.FromArgb(255, 140, 0);     // Dark orange
        public static readonly Color MarqueeNeon = Color.FromArgb(57, 255, 20);       // Neon green

        // Màu chữ - ĐEN cho labels
        public static readonly Color MacTextPrimary = Color.FromArgb(0, 0, 0);          // Pure black
        public static readonly Color MacTextSecondary = Color.FromArgb(100, 100, 100);  // Dark gray
        public static readonly Color MacTextLight = Color.White;

        // Màu viền
        public static readonly Color MacBorder = Color.FromArgb(180, 200, 220);
        public static readonly Color MacBorderLight = Color.FromArgb(200, 215, 230);
        public static readonly Color MacBorderFocus = Color.FromArgb(0, 122, 255);

        // Bóng đổ
        public static readonly Color MacShadow = Color.FromArgb(50, 0, 0, 0);
        public static readonly Color MacHeaderGradient1 = Color.FromArgb(220, 235, 245);
        public static readonly Color MacHeaderGradient2 = Color.FromArgb(235, 245, 252);
    }
}
