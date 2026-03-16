using System;

namespace HOSONHCS
{
    public class KnowledgeItem
    {
        public string   Id        { get; set; } = Guid.NewGuid().ToString("N");
        public string   Category  { get; set; } = "Chung";
        public string   Question  { get; set; } = "";
        public string[] Keywords  { get; set; } = new string[0];
        public string   Answer    { get; set; } = "";
        public int      Priority  { get; set; } = 5;
        public bool     IsActive  { get; set; } = true;
        public DateTime CreatedAt { get; set; } = DateTime.Now;
    }
}
