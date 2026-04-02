using System;
using System.Globalization;
using Newtonsoft.Json;

namespace HOSONHCS
{
    /// <summary>
    /// Model một thông báo ngày lễ / sự kiện đặc biệt
    /// </summary>
    public class HolidayNoticeMessage
    {
        /// <summary>ID duy nhất để tránh hiện lại trong cùng một ngày</summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>Tiêu đề thông báo</summary>
        [JsonProperty("title")]
        public string Title { get; set; }

        /// <summary>Nội dung chi tiết (hỗ trợ xuống dòng \n)</summary>
        [JsonProperty("content")]
        public string Content { get; set; }

        /// <summary>Emoji hiển thị ở header (vd: 🎉 🌸 🎆)</summary>
        [JsonProperty("emoji")]
        public string Emoji { get; set; }

        /// <summary>Ngày bắt đầu hiển thị (định dạng yyyy-MM-dd)</summary>
        [JsonProperty("startDate")]
        public string StartDate { get; set; }

        /// <summary>Ngày kết thúc hiển thị (định dạng yyyy-MM-dd)</summary>
        [JsonProperty("endDate")]
        public string EndDate { get; set; }

        /// <summary>Có bật thông báo này không (false = tắt)</summary>
        [JsonProperty("enabled")]
        public bool Enabled { get; set; }

        /// <summary>Kiểm tra thông báo có đang hiệu lực hôm nay không</summary>
        public bool IsActiveToday()
        {
            if (!Enabled) return false;
            var today = DateTime.Today;
            if (DateTime.TryParseExact(StartDate, "yyyy-MM-dd", CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out var start) &&
                DateTime.TryParseExact(EndDate, "yyyy-MM-dd", CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out var end))
            {
                return today >= start && today <= end;
            }
            return false;
        }
    }

    /// <summary>
    /// Root object của file JSON holiday_notice.json trên GitHub
    /// </summary>
    public class HolidayNoticeData
    {
        [JsonProperty("notices")]
        public HolidayNoticeMessage[] Notices { get; set; }
    }
}
