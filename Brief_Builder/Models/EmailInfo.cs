using System;

namespace Brief_Builder.Models
{
    public class EmailInfo
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public string From { get; set; }
        public string To { get; set; }
        public string Body { get; set; }
    }
}
