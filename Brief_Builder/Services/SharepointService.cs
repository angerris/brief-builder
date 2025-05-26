using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Brief_Builder.Models;
using Newtonsoft.Json;

namespace Brief_Builder.Services
{
    public class SharepointService
    {
        private readonly HttpClient _client;
        private readonly string _accessToken;
        private readonly string _siteId = Environment.GetEnvironmentVariable("spSiteID");
        private readonly string _tenantId = Environment.GetEnvironmentVariable("spTenantID");
        private readonly string _clientId = Environment.GetEnvironmentVariable("spClientID");
        private readonly string _clientSecret = Environment.GetEnvironmentVariable("spSecret");

        public SharepointService()
        {
            _client = new HttpClient();
            var tokenResp = GetTokenResponse().GetAwaiter().GetResult();
            _accessToken = tokenResp.AccessToken;
            _client.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", _accessToken);
        }

        private async Task<TokenResponse> GetTokenResponse()
        {
            var authority = $"https://login.microsoftonline.com/{_tenantId}";
            var tokenEndpoint = $"{authority}/oauth2/v2.0/token";

            using var tempClient = new HttpClient();
            var form = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string,string>("grant_type", "client_credentials"),
                new KeyValuePair<string,string>("client_id", _clientId),
                new KeyValuePair<string,string>("client_secret", _clientSecret),
                new KeyValuePair<string,string>("scope", "https://graph.microsoft.com/.default")
            });

            var resp = await tempClient.PostAsync(tokenEndpoint, form);
            resp.EnsureSuccessStatusCode();
            var body = await resp.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<TokenResponse>(body);
        }

        public async Task<string> GetClaimDriveId()
        {
            const string driveName = "Claim";
            var apiUrl = $"https://graph.microsoft.com/v1.0/sites/{_siteId}/drives";

            var resp = await _client.GetAsync(apiUrl);
            resp.EnsureSuccessStatusCode();
            var json = await resp.Content.ReadAsStringAsync();
            var root = JsonConvert.DeserializeObject<SharepointDrives>(json);
            return root.Value
                .FirstOrDefault(d =>
                    string.Equals(d.Name, driveName, StringComparison.OrdinalIgnoreCase))?.Id;
        }

        public async Task<byte[]> DownloadDocumentAsPDFFromSharePoint(string driveId, string itemId)
        {
            var apiUrl =
                $"https://graph.microsoft.com/v1.0/sites/{_siteId}" +
                $"/drives/{driveId}/items/{itemId}/content?format=pdf";

            var resp = await _client.GetAsync(apiUrl);
            resp.EnsureSuccessStatusCode();
            return await resp.Content.ReadAsByteArrayAsync();
        }

        public async Task<byte[]> DownloadDocumentFromSharePoint(string driveId, string fileId)
        {
            var apiUrl =
                $"https://graph.microsoft.com/v1.0/sites/{_siteId}" +
                $"/drives/{driveId}/items/{fileId}/content";

            var resp = await _client.GetAsync(apiUrl);
            resp.EnsureSuccessStatusCode();
            return await resp.Content.ReadAsByteArrayAsync();
        }

        public async Task<string> UploadDocumentToSharePoint(string driveId, string folderPath, string fileName, byte[] fileContent)
        {
            var apiUrl =
                $"https://graph.microsoft.com/v1.0/sites/{_siteId}" +
                $"/drives/{driveId}/root:/{folderPath}/{fileName}:/content" +
                "?@microsoft.graph.conflictBehavior=replace";

            using var content = new ByteArrayContent(fileContent);
            content.Headers.ContentType =
                new MediaTypeHeaderValue("application/octet-stream");

            var resp = await _client.PutAsync(apiUrl, content);
            resp.EnsureSuccessStatusCode();
            var json = await resp.Content.ReadAsStringAsync();
            dynamic obj = JsonConvert.DeserializeObject(json);
            return (string)obj.id;
        }

        public async Task<List<ImportedFile>> GetImportedFilesAsync(IEnumerable<Dictionary<string, string>> fileRefs)
        {
            var imported = new List<ImportedFile>();
            if (fileRefs == null || !fileRefs.Any()) return imported;

            var driveId = await GetClaimDriveId();

            foreach (var sp in fileRefs)
            {
                sp.TryGetValue("id", out var idValue);
                sp.TryGetValue("name", out var nameValue);

                if (string.IsNullOrEmpty(idValue))
                    continue;


                    var bytes = await DownloadDocumentFromSharePoint(driveId, idValue);

                    imported.Add(new ImportedFile
                    {
                        Id = idValue,
                        Name = nameValue,
                        Content = bytes
                    });
            }

            return imported;
        }
    }
}
