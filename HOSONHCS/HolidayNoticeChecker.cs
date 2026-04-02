using System;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace HOSONHCS
{
    /// <summary>
    /// Fetch và hiển thị thông báo ngày lễ từ file JSON trên GitHub.
    ///
    /// Cách dùng (admin):
    ///   1. Chỉnh sửa file holiday_notice.json trong repo GitHub.
    ///   2. Commit &amp; push → người dùng sẽ thấy popup ngay khi mở app.
    ///
    /// URL fetch:
    ///   https://raw.githubusercontent.com/NguyenHaiHG/HOSONHCS/master/holiday_notice.json
    /// </summary>
    public static class HolidayNoticeChecker
    {
        private const string NOTICE_URL =
            "https://raw.githubusercontent.com/NguyenHaiHG/HOSONHCS/master/holiday_notice.json";

        /// <summary>File lưu ID thông báo đã hiện hôm nay (tránh hiện lại khi mở lại app)</summary>
        private static string ShownTodayFilePath =>
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".notice_shown");

        /// <summary>
        /// Fetch thông báo từ GitHub và hiển thị popup nếu hôm nay có thông báo hiệu lực.
        /// Gọi bất đồng bộ từ Form1 — không chặn UI.
        /// </summary>
        public static async Task CheckAndShowAsync(Form parentForm)
        {
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                var request = (HttpWebRequest)WebRequest.Create(NOTICE_URL);
                request.Method    = "GET";
                request.UserAgent = "HOSONHCS-App";
                request.Timeout   = 8000;

                string json;
                using (var response = (HttpWebResponse)await request.GetResponseAsync())
                using (var stream   = response.GetResponseStream())
                using (var reader   = new StreamReader(stream))
                    json = await reader.ReadToEndAsync();

                var data = JsonConvert.DeserializeObject<HolidayNoticeData>(json);
                if (data?.Notices == null) return;

                string alreadyShownId = ReadShownTodayId();

                foreach (var notice in data.Notices)
                {
                    if (!notice.IsActiveToday()) continue;
                    // Nếu đã hiện thông báo này hôm nay thì bỏ qua
                    if (alreadyShownId == notice.Id) continue;

                    // Hiện popup trên UI thread
                    parentForm.Invoke(new Action(() =>
                    {
                        using (var popup = new HolidayNoticeForm(notice))
                            popup.ShowDialog(parentForm);
                    }));

                    SaveShownTodayId(notice.Id);
                    break; // chỉ hiện 1 thông báo mỗi lần mở app
                }
            }
            catch
            {
                // Không làm phiền user khi mạng offline hoặc lỗi fetch
            }
        }

        // ── ĐỌC / GHI ID ĐÃ HIỆN HÔM NAY ────────────────────────────────────

        /// <summary>
        /// Đọc ID thông báo đã được hiện hôm nay.
        /// File .notice_shown lưu dạng: "yyyy-MM-dd|noticeId"
        /// </summary>
        private static string ReadShownTodayId()
        {
            try
            {
                if (!File.Exists(ShownTodayFilePath)) return null;
                var content = File.ReadAllText(ShownTodayFilePath).Trim();
                var parts   = content.Split('|');
                if (parts.Length == 2 && parts[0] == DateTime.Today.ToString("yyyy-MM-dd"))
                    return parts[1];
            }
            catch { }
            return null;
        }

        /// <summary>Lưu ID thông báo vừa hiện để tránh hiện lại trong ngày</summary>
        private static void SaveShownTodayId(string noticeId)
        {
            try
            {
                File.WriteAllText(ShownTodayFilePath,
                    $"{DateTime.Today:yyyy-MM-dd}|{noticeId}");
            }
            catch { }
        }
    }
}
