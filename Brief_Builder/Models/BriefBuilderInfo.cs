using System.Collections.Generic;
using Newtonsoft.Json;

namespace Brief_Builder.Models
{
    public sealed class BriefBuilderInfo
    {
        [JsonProperty("recordId")]
        public string RecordId { get; set; }

        [JsonProperty("emailIds")]
        public List<string> EmailIds { get; set; }

        [JsonProperty("claims")]
        public List<Dictionary<string, string>> Claims { get; set; }

        [JsonProperty("sharepointFiles")]
        public List<Dictionary<string, string>> SharePointFiles { get; set; }
    }
    public sealed class ImportedFile
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public byte[] Content { get; set; }
    }


}
