using System;

namespace HOSONHCS
{
    public class NoteItem
    {
        public string Id        { get; set; } = Guid.NewGuid().ToString("N");
        public string Title     { get; set; } = "";
        public string Content   { get; set; } = "";
        public string Category  { get; set; } = "Chung";
        public bool   IsPinned  { get; set; } = false;
        public DateTime CreatedAt { get; set; } = DateTime.Now;
        public DateTime UpdatedAt { get; set; } = DateTime.Now;
    }
}
