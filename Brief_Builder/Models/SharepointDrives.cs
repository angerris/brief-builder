using System.Collections.Generic;

namespace Brief_Builder.Models
{
    public class SharepointDrives
    {
        public List<Drive> Value { get; set; }
    }

    public class Drive
    {
        public string Id { get; set; }
        public string Name { get; set; }
    }
}
