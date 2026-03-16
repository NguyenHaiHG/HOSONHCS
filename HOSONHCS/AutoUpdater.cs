using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace HOSONHCS
{
    /// <summary>
    /// Class quản lý auto-update cho ứng dụng
    /// Kiểm tra phiên bản mới từ GitHub Releases
    /// </summary>
    public class AutoUpdater
    {
        // ============================================
        // CẤU HÌNH
        // ============================================

        /// <summary>
        /// URL để kiểm tra phiên bản mới (GitHub Releases API)
        /// </summary>
        private const string VERSION_CHECK_URL = "https://api.github.com/repos/NguyenHaiHG/HOSONHCS/releases/latest";

        /// <summary>
        /// GitHub Personal Access Token (nếu repository là private)
        /// KHÔNG CẦN nếu chỉ release là public
        /// </summary>
        private const string GITHUB_TOKEN = ""; // Để trống - không cần token!

        /// <summary>
        /// Phiên bản hiện tại của app
        /// Lấy từ Assembly Version
        /// </summary>
        public static string CurrentVersion
        {
            get
            {
                var version = Assembly.GetExecutingAssembly().GetName().Version;
                return $"{version.Major}.{version.Minor}.{version.Build}";
            }
        }

        // ============================================
        // KIỂM TRA CẬP NHẬT
        // ============================================

        /// <summary>
        /// Kiểm tra có phiên bản mới không (async)
        /// </summary>
        /// <param name="silent">Nếu true, không hiện MessageBox khi không có update</param>
        /// <returns>UpdateInfo hoặc null nếu không có update</returns>
        public static async Task<UpdateInfo> CheckForUpdateAsync(bool silent = false)
        {
            try
            {
                // Tạo HTTP request đến GitHub API
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                var request = (HttpWebRequest)WebRequest.Create(VERSION_CHECK_URL);
                request.Method = "GET";
                request.UserAgent = "HOSONHCS-AutoUpdater";
                request.Accept = "application/vnd.github.v3+json";
                request.Timeout = 10000; // 10 seconds

                // Lấy response
                using (var response = (HttpWebResponse)await request.GetResponseAsync())
                using (var stream = response.GetResponseStream())
                using (var reader = new StreamReader(stream))
                {
                    var json = await reader.ReadToEndAsync();
                    var releaseInfo = JsonConvert.DeserializeObject<GitHubRelease>(json);

                    if (releaseInfo == null || string.IsNullOrEmpty(releaseInfo.tag_name))
                    {
                        // Không hiện thông báo khi không có update
                        return null;
                    }

                    // So sánh version (bỏ chữ 'v' ở đầu tag)
                    string latestVersion = releaseInfo.tag_name.TrimStart('v');

                    if (IsNewerVersion(latestVersion, CurrentVersion))
                    {
                        // Tìm file .zip trong assets
                        string downloadUrl = null;
                        foreach (var asset in releaseInfo.assets)
                        {
                            if (asset.name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                            {
                                downloadUrl = asset.browser_download_url;
                                break;
                            }
                        }

                        if (string.IsNullOrEmpty(downloadUrl))
                        {
                            if (!silent)
                                MessageBox.Show("Không tìm thấy file cập nhật.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return null;
                        }

                        return new UpdateInfo
                        {
                            Version = latestVersion,
                            DownloadUrl = downloadUrl,
                            Changelog = releaseInfo.body ?? "Không có mô tả",
                            PublishedDate = releaseInfo.published_at
                        };
                    }
                    else
                    {
                        // Không hiện thông báo khi không có update
                        return null;
                    }
                }
            }
            catch (WebException ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AutoUpdater] ❌ WebException: {ex.Message}");

                // Kiểm tra nếu là lỗi 404 (chưa có release)
                if (ex.Response is HttpWebResponse response && response.StatusCode == HttpStatusCode.NotFound)
                {
                    // 404 = Repository chưa có release nào
                    System.Diagnostics.Debug.WriteLine("[AutoUpdater] 404 - Repository chưa có release nào");

                    // HIỆN THÔNG BÁO NGAY CẢ KHI SILENT (để debug)
                    MessageBox.Show(
                        "⚠️ Repository chưa có bản phát hành (release)\n\n" +
                        "Vui lòng tạo release trên GitHub:\n" +
                        "https://github.com/NguyenHaiHG/HOSONHCS/releases/new",
                        "Chưa có release",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return null;
                }

                // Các lỗi khác (network, timeout, v.v.)
                System.Diagnostics.Debug.WriteLine($"[AutoUpdater] Lỗi khác: {ex.Status}");
                if (!silent)
                {
                    MessageBox.Show(
                        $"Không thể kết nối đến server cập nhật.\n\n{ex.Message}\n\n" +
                        "Có thể do:\n" +
                        "- Không có kết nối Internet\n" +
                        "- Server GitHub đang bảo trì\n" +
                        "- Repository chưa có bản phát hành (release)",
                        "Lỗi kết nối",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                }
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AutoUpdater] ❌ Exception: {ex.Message}");
                if (!silent)
                    MessageBox.Show($"Lỗi khi kiểm tra cập nhật:\n\n{ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        /// <summary>
        /// So sánh 2 version string
        /// </summary>
        /// <returns>true nếu newVersion > currentVersion</returns>
        private static bool IsNewerVersion(string newVersion, string currentVersion)
        {
            try
            {
                var newParts = newVersion.Split('.');
                var currentParts = currentVersion.Split('.');

                for (int i = 0; i < Math.Max(newParts.Length, currentParts.Length); i++)
                {
                    int newPart = i < newParts.Length ? int.Parse(newParts[i]) : 0;
                    int currentPart = i < currentParts.Length ? int.Parse(currentParts[i]) : 0;

                    if (newPart > currentPart) return true;
                    if (newPart < currentPart) return false;
                }

                return false; // Bằng nhau
            }
            catch
            {
                return false;
            }
        }

        // ============================================
        // TẢI VÀ CÀI ĐẶT CẬP NHẬT
        // ============================================

        /// <summary>
        /// Tải và cài đặt bản cập nhật
        /// </summary>
        /// <param name="updateInfo">Thông tin bản cập nhật</param>
        /// <param name="progressCallback">Callback để báo tiến độ (0-100)</param>
        public static async Task<bool> DownloadAndInstallAsync(UpdateInfo updateInfo, IProgress<int> progressCallback = null)
        {
            try
            {
                // Tạo folder tạm để tải về
                string tempFolder = Path.Combine(Path.GetTempPath(), "HOSONHCS_Update");
                if (Directory.Exists(tempFolder))
                    Directory.Delete(tempFolder, true);
                Directory.CreateDirectory(tempFolder);

                string zipPath = Path.Combine(tempFolder, "update.zip");
                string extractPath = Path.Combine(tempFolder, "extracted");

                // Tải file zip
                using (var client = new WebClient())
                {
                    client.DownloadProgressChanged += (s, e) =>
                    {
                        progressCallback?.Report(e.ProgressPercentage / 2); // 0-50% cho download
                    };

                    await client.DownloadFileTaskAsync(updateInfo.DownloadUrl, zipPath);
                }

                progressCallback?.Report(50);

                // Giải nén (dùng Shell.Application thay vì ZipFile để tránh cần reference)
                ExtractZipFile(zipPath, extractPath);
                progressCallback?.Report(75);

                // Gọi Updater.exe để thực hiện update
                string currentExePath = Assembly.GetExecutingAssembly().Location;
                string currentFolder = Path.GetDirectoryName(currentExePath);
                string updaterExe = Path.Combine(currentFolder, "Updater.exe");
                if (!File.Exists(updaterExe))
                {
                    MessageBox.Show($"Không tìm thấy {updaterExe}. Hãy chắc chắn rằng Updater.exe nằm cùng thư mục với ứng dụng.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                // Tham số: <processName> <sourceFolder> <targetFolder> <mainExeName>
                var psi = new ProcessStartInfo
                {
                    FileName = updaterExe,
                    Arguments = $"HOSONHCS.exe \"{extractPath}\" \"{currentFolder}\" HOSONHCS.exe",
                    UseShellExecute = true,
                    CreateNoWindow = false,
                    WindowStyle = ProcessWindowStyle.Normal
                };

                Process.Start(psi);

                progressCallback?.Report(100);

                // Thoát ứng dụng để cho Updater.exe update
                Application.Exit();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi cài đặt bản cập nhật:\n\n{ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// Giải nén file zip (không cần System.IO.Compression.FileSystem)
        /// </summary>
        private static void ExtractZipFile(string zipPath, string extractPath)
        {
            try
            {
                // Tạo folder đích
                if (!Directory.Exists(extractPath))
                    Directory.CreateDirectory(extractPath);

                // Sử dụng Shell.Application (COM) để giải nén
                var shellApplication = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));
                dynamic zipFile = shellApplication.GetType().InvokeMember(
                    "NameSpace",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    shellApplication,
                    new object[] { zipPath });

                dynamic destinationFolder = shellApplication.GetType().InvokeMember(
                    "NameSpace",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    shellApplication,
                    new object[] { extractPath });

                dynamic items = zipFile.Items();

                // Copy items (16 = no progress dialog)
                destinationFolder.GetType().InvokeMember(
                    "CopyHere",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    destinationFolder,
                    new object[] { items, 16 });

                // Đợi giải nén xong
                System.Threading.Thread.Sleep(1000);
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi khi giải nén file: {ex.Message}");
            }
        }

        // ============================================
        // HIỂN THỊ DIALOG CẬP NHẬT
        // ============================================

        /// <summary>
        /// Hiển thị dialog hỏi user có muốn cập nhật không
        /// </summary>
        public static async Task ShowUpdateDialogAsync(UpdateInfo updateInfo)
        {
            if (updateInfo == null) return;

            var message = $"Có phiên bản mới: {updateInfo.Version}\n" +
                         $"Phiên bản hiện tại: {CurrentVersion}\n" +
                         $"Ngày phát hành: {updateInfo.PublishedDate:dd/MM/yyyy}\n\n" +
                         $"Thay đổi:\n{updateInfo.Changelog}\n\n" +
                         $"Bạn có muốn cập nhật ngay bây giờ không?";

            var result = MessageBox.Show(message, "Cập nhật mới", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

            if (result == DialogResult.Yes)
            {
                // Tạo form progress đẹp, bo góc, màu nền xanh lá, label font hiện đại, progress bar custom màu xanh dương
                var progressForm = new Form
                {
                    Text = "Đang tải cập nhật...",
                    Width = 420,
                    Height = 170,
                    FormBorderStyle = FormBorderStyle.None,
                    StartPosition = FormStartPosition.CenterScreen,
                    MaximizeBox = false,
                    MinimizeBox = false,
                    BackColor = System.Drawing.Color.FromArgb(30, 130, 76), // xanh lá đậm
                    ShowInTaskbar = false
                };

                // Bo góc form (chỉ hiệu quả trên Win10+)
                progressForm.Load += (s, e) => {
                    try {
                        progressForm.Region = System.Drawing.Region.FromHrgn(
                            NativeMethods.CreateRoundRectRgn(0, 0, progressForm.Width, progressForm.Height, 18, 18));
                    } catch { }
                };

                var label = new Label
                {
                    Text = "Đang tải xuống...",
                    Dock = DockStyle.Top,
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                    Height = 38,
                    Font = new System.Drawing.Font("Segoe UI", 13F, System.Drawing.FontStyle.Bold),
                    ForeColor = System.Drawing.Color.White,
                    BackColor = System.Drawing.Color.Transparent
                };

                // Custom Progress Bar với hiệu ứng 3D và màu hồng cánh sen
                var progressBar = new CustomProgressBar
                {
                    Dock = DockStyle.Bottom,
                    Height = 40,
                    Value = 0,
                    Maximum = 100
                };

                // Panel nền bo góc
                var panel = new Panel
                {
                    Dock = DockStyle.Fill,
                    BackColor = System.Drawing.Color.Transparent,
                    Padding = new Padding(24, 18, 24, 18)
                };
                panel.Controls.Add(label);
                panel.Controls.Add(progressBar);
                progressForm.Controls.Add(panel);

                // Progress callback
                var progress = new Progress<int>(percent =>
                {
                    if (progressBar.InvokeRequired)
                    {
                        progressBar.Invoke(new Action(() => progressBar.Value = percent));
                    }
                    else
                    {
                        progressBar.Value = percent;
                    }

                    if (percent < 50)
                        label.Text = $"Đang tải xuống... {percent}%";
                    else if (percent < 75)
                        label.Text = "Đang giải nén...";
                    else
                        label.Text = "Đang cài đặt...";
                });

                progressForm.Show();

                await DownloadAndInstallAsync(updateInfo, progress);

                progressForm.Close();
            }
        }
    }

    // ============================================
    // MODELS
    // ============================================

    /// <summary>
    /// Thông tin bản cập nhật
    /// </summary>
    public class UpdateInfo
    {
        public string Version { get; set; }
        public string DownloadUrl { get; set; }
        public string Changelog { get; set; }
        public DateTime PublishedDate { get; set; }
    }

    /// <summary>
    /// Model cho GitHub Release API
    /// </summary>
    internal class GitHubRelease
    {
        public string tag_name { get; set; }
        public string name { get; set; }
        public string body { get; set; }
        public DateTime published_at { get; set; }
        public GitHubAsset[] assets { get; set; }
    }

    internal class GitHubAsset
    {
        public string name { get; set; }
        public string browser_download_url { get; set; }
    }
}

// Custom Progress Bar với hiệu ứng 3D và màu hồng cánh sen
public class CustomProgressBar : System.Windows.Forms.ProgressBar
{
    public CustomProgressBar()
    {
        this.SetStyle(System.Windows.Forms.ControlStyles.UserPaint, true);
        this.SetStyle(System.Windows.Forms.ControlStyles.OptimizedDoubleBuffer, true);
        this.SetStyle(System.Windows.Forms.ControlStyles.AllPaintingInWmPaint, true);
    }

    protected override void OnPaint(System.Windows.Forms.PaintEventArgs e)
    {
        System.Drawing.Rectangle rect = e.ClipRectangle;
        rect.Width = (int)(rect.Width * ((double)Value / Maximum)) - 4;
        rect.Height -= 4;

        // Màu hồng cánh sen (Lotus Pink)
        var lotusPink1 = System.Drawing.Color.FromArgb(255, 182, 193); // Hồng nhạt
        var lotusPink2 = System.Drawing.Color.FromArgb(255, 105, 180); // Hồng đậm
        var lotusPink3 = System.Drawing.Color.FromArgb(255, 20, 147);  // Hồng sẫm

        // Vẽ nền xám nhạt cho phần chưa hoàn thành
        using (var bgBrush = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(240, 240, 240)))
        {
            e.Graphics.FillRectangle(bgBrush, 0, 0, e.ClipRectangle.Width, e.ClipRectangle.Height);
        }

        // Vẽ viền ngoài
        using (var borderPen = new System.Drawing.Pen(System.Drawing.Color.FromArgb(200, 200, 200), 2))
        {
            e.Graphics.DrawRectangle(borderPen, 1, 1, e.ClipRectangle.Width - 2, e.ClipRectangle.Height - 2);
        }

        if (rect.Width > 0)
        {
            // Tạo gradient 3D như ống chất lỏng
            var progressRect = new System.Drawing.Rectangle(2, 2, rect.Width, rect.Height);

            using (var gradient = new System.Drawing.Drawing2D.LinearGradientBrush(
                progressRect,
                lotusPink1,
                lotusPink3,
                System.Drawing.Drawing2D.LinearGradientMode.Vertical))
            {
                // Tạo blend pattern cho hiệu ứng 3D
                var blend = new System.Drawing.Drawing2D.ColorBlend(5);
                blend.Colors = new System.Drawing.Color[]
                {
                    lotusPink1,           // Top highlight
                    lotusPink2,           // Upper middle
                    lotusPink3,           // Center (darkest)
                    lotusPink2,           // Lower middle
                    System.Drawing.Color.FromArgb(255, 160, 180) // Bottom
                };
                blend.Positions = new float[] { 0.0f, 0.2f, 0.5f, 0.8f, 1.0f };
                gradient.InterpolationColors = blend;

                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                e.Graphics.FillRectangle(gradient, progressRect);
            }

            // Thêm highlight ở trên để tạo hiệu ứng bóng
            var highlightRect = new System.Drawing.Rectangle(progressRect.X, progressRect.Y, progressRect.Width, progressRect.Height / 3);
            using (var highlightBrush = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(100, 255, 255, 255)))
            {
                e.Graphics.FillRectangle(highlightBrush, highlightRect);
            }

            // Thêm shadow ở dưới
            var shadowRect = new System.Drawing.Rectangle(progressRect.X, progressRect.Y + (int)(progressRect.Height * 0.7f), progressRect.Width, (int)(progressRect.Height * 0.3f));
            using (var shadowBrush = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(60, 0, 0, 0)))
            {
                e.Graphics.FillRectangle(shadowBrush, shadowRect);
            }
        }
    }
}

// Native methods for bo góc và custom màu progress bar
internal static class NativeMethods
{
    [System.Runtime.InteropServices.DllImport("gdi32.dll", SetLastError = true)]
    public static extern IntPtr CreateRoundRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect, int nWidthEllipse, int nHeightEllipse);

    [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
}
