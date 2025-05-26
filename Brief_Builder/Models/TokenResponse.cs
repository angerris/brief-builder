using Newtonsoft.Json;

namespace Brief_Builder.Models
{
    public sealed class TokenResponse
    {
        [JsonProperty("access_token")]
        public string AccessToken { get; set; }
    }
}
