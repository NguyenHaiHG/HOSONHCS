using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace Updater
{
    class Program
    {
        // Win32 API để ẩn console window
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;

        // args: <processName> <sourceFolder> <targetFolder> <mainExeName>
        [STAThread]
        static int Main(string[] args)
        {
            // Ẩn cửa sổ console ngay lập tức
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);

            if (args.Length < 4)
            {
                Console.WriteLine("Usage: Updater.exe <processName> <sourceFolder> <targetFolder> <mainExeName>");
                return 1;
            }

            string processName = args[0];
            string sourceFolder = args[1];
            string targetFolder = args[2];
            string mainExeName = args[3];

            string logPath = Path.Combine(targetFolder, "update_log.txt");
            void Log(string msg)
            {
                try { File.AppendAllText(logPath, DateTime.Now + ": " + msg + Environment.NewLine); } catch { }
            }
            try
            {
                Log($"Bắt đầu update. processName={processName}, sourceFolder={sourceFolder}, targetFolder={targetFolder}, mainExeName={mainExeName}");
                // 1. Đợi process chính thoát
                WaitForProcessExit(processName);
                Log("Đã tắt process chính.");

                // 2. Nếu sourceFolder chỉ có 1 folder con và không có file lẻ, copy nội dung của folder con đó ra targetFolder
                var dirs = Directory.GetDirectories(sourceFolder);
                var files = Directory.GetFiles(sourceFolder);
                if (dirs.Length == 1 && files.Length == 0)
                {
                    string innerFolder = dirs[0];
                    Log($"Chỉ có 1 folder con: {innerFolder}. Sẽ copy toàn bộ nội dung của nó ra app gốc.");
                    foreach (var file in Directory.GetFiles(innerFolder, "*.*", SearchOption.AllDirectories))
                    {
                        try
                        {
                            string relPath = GetRelativePath(innerFolder, file);
                            string destPath = Path.Combine(targetFolder, relPath);
                            string destDir = Path.GetDirectoryName(destPath);
                            if (!Directory.Exists(destDir)) Directory.CreateDirectory(destDir);
                            File.Copy(file, destPath, true);
                            Log($"Đã copy {relPath} -> {destPath}");
                        }
                        catch (Exception ex)
                        {
                            Log($"Lỗi copy file {file}: {ex.Message}");
                        }
                    }
                }
                else
                {
                    // Copy tất cả file/folder như cũ
                    foreach (var file in Directory.GetFiles(sourceFolder, "*.*", SearchOption.AllDirectories))
                    {
                        try
                        {
                            string relPath = GetRelativePath(sourceFolder, file);
                            string destPath = Path.Combine(targetFolder, relPath);
                            string destDir = Path.GetDirectoryName(destPath);
                            if (!Directory.Exists(destDir)) Directory.CreateDirectory(destDir);
                            File.Copy(file, destPath, true);
                            Log($"Đã copy {relPath} -> {destPath}");
                        }
                        catch (Exception ex)
                        {
                            Log($"Lỗi copy file {file}: {ex.Message}");
                        }
                    }
                }

                // 3. Thông báo thành công và khởi động lại app chính
                string mainExePath = Path.Combine(targetFolder, mainExeName);
                bool notified = false;
                try
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    MessageBox.Show("Cập nhật thành công! Ứng dụng sẽ được mở lại.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    notified = true;
                    Log("Đã hiện MessageBox thông báo thành công.");
                }
                catch (Exception ex)
                {
                    Log($"Không hiện được MessageBox: {ex.Message}");
                }
                if (!notified)
                {
                    Console.WriteLine("Cập nhật thành công! Ứng dụng sẽ được mở lại.");
                    Log("Đã ghi thông báo ra Console.");
                }


                bool started = false;
                if (File.Exists(mainExePath))
                {
                    try
                    {
                        var psi = new ProcessStartInfo
                        {
                            FileName = mainExePath,
                            WorkingDirectory = Path.GetDirectoryName(mainExePath),
                            UseShellExecute = true
                        };
                        var proc = Process.Start(psi);
                        started = proc != null;
                        Log($"Đã gọi Process.Start cho {mainExePath} với WorkingDirectory và UseShellExecute");
                        // Đợi 3s kiểm tra process đã chạy chưa
                        Thread.Sleep(3000);
                        var running = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(mainExeName));
                        if (running.Length == 0)
                        {
                            Log($"CẢNH BÁO: Không phát hiện process {mainExeName} sau khi gọi Process.Start!");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Không thể khởi động lại ứng dụng: {ex.Message}");
                        Log($"Không thể khởi động lại ứng dụng: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Không tìm thấy file {mainExePath} để khởi động lại.");
                    Log($"Không tìm thấy file {mainExePath} để khởi động lại.");
                }

                // 4. Xóa thư mục tạm nếu có thể
                try
                {
                    if (Directory.Exists(sourceFolder))
                    {
                        Directory.Delete(sourceFolder, true);
                        Log($"Đã xóa thư mục tạm: {sourceFolder}");
                    }
                }
                catch (Exception ex)
                {
                    Log($"Không thể xóa thư mục tạm {sourceFolder}: {ex.Message}");
                }

                Log("Kết thúc update thành công.");
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi cập nhật: {ex.Message}");
                Log($"Lỗi cập nhật: {ex.Message}");
                return 2;
            }
        }

        static void WaitForProcessExit(string processName)
        {
            while (true)
            {
                var running = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(processName));
                if (running.Length == 0)
                    break;
                Thread.Sleep(1000);
            }
        }

        static string GetRelativePath(string baseDir, string fullPath)
        {
            if (!baseDir.EndsWith("\\") && !baseDir.EndsWith("/"))
                baseDir += Path.DirectorySeparatorChar;
            Uri baseUri = new Uri(baseDir, UriKind.Absolute);
            Uri fileUri = new Uri(fullPath, UriKind.Absolute);
            return Uri.UnescapeDataString(baseUri.MakeRelativeUri(fileUri).ToString().Replace('/', Path.DirectorySeparatorChar));
        }
    }
}
