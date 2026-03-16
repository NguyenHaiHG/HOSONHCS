using System;
using System.Collections.Generic;

namespace HOSONHCS
{
    public class ChatMessage
    {
        public string   Role      { get; set; }  // "user" | "bot"
        public string   Content   { get; set; }
        public DateTime Timestamp { get; set; } = DateTime.Now;
    }

    public class ChatSession
    {
        public string            Id       { get; set; } = Guid.NewGuid().ToString("N");
        public DateTime          Date     { get; set; } = DateTime.Now;
        public List<ChatMessage> Messages { get; set; } = new List<ChatMessage>();
    }
}
