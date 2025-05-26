using System.Collections.Generic;

namespace Brief_Builder.Models
{
    public sealed class SharepointDrives
    {
        public List<Drive> Value { get; set; }
    }

    public sealed class Drive
    {
        public string Id { get; set; }
        public string Name { get; set; }
    }
}
