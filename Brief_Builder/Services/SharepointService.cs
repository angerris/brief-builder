using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using Brief_Builder.Models;
using Newtonsoft.Json;

namespace Brief_Builder.Services
{
    public static class SharepointService
    {
        private static readonly string _clientId = Environment.GetEnvironmentVariable("spClientID");
        private static readonly string _clientSecret = Environment.GetEnvironmentVariable("spSecret");
        private static readonly string _siteId = Environment.GetEnvironmentVariable("spSiteID");
        private static readonly string _tenantId = Environment.GetEnvironmentVariable("spTenantID");

        public static TokenResponse GetTokenResponse()
        {
            var authority = $"https://login.microsoftonline.com/{_tenantId}";
            var tokenEndpoint = $"{authority}/oauth2/v2.0/token";

            using var client = new HttpClient();
            var form = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string,string>("grant_type",    "client_credentials"),
                new KeyValuePair<string,string>("client_id",     _clientId),
                new KeyValuePair<string,string>("client_secret", _clientSecret),
                new KeyValuePair<string,string>("scope",         "https://graph.microsoft.com/.default")
            });

            var resp = client.PostAsync(tokenEndpoint, form).GetAwaiter().GetResult();
            resp.EnsureSuccessStatusCode();
            var body = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();
            return JsonConvert.DeserializeObject<TokenResponse>(body);
        }

        public static string SiteId => _siteId;

        public static string GetClaimDriveId(string accessToken)
        {
            const string driveName = "Claim";
            var apiUrl = $"https://graph.microsoft.com/v1.0/sites/{_siteId}/drives";

            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", accessToken);

            var resp = client.GetAsync(apiUrl).GetAwaiter().GetResult();
            resp.EnsureSuccessStatusCode();
            var json = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();

            var root = JsonConvert.DeserializeObject<SharepointDrives>(json);
            var match = root.Value.FirstOrDefault(d =>
                string.Equals(d.Name, driveName, StringComparison.OrdinalIgnoreCase));

            return match.Id;
        }

        public static byte[] DownloadDocumentFromSharePoint(
            string driveId,
            string fileId,
            string accessToken)
        {
            var apiUrl =
                $"https://graph.microsoft.com/v1.0/sites/{_siteId}" +
                $"/drives/{driveId}/items/{fileId}/content";

            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", accessToken);

            var resp = client.GetAsync(apiUrl).GetAwaiter().GetResult();
            resp.EnsureSuccessStatusCode();
            return resp.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
        }

        public static void UploadDocumentToSharePoint(
            string driveId,
            string folderPath,
            string fileName,
            byte[] fileContent,
            string accessToken)
        {
            var apiUrl =
                $"https://graph.microsoft.com/v1.0/sites/{_siteId}" +
                $"/drives/{driveId}/root:/{folderPath}/{fileName}:/content";

            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", accessToken);

            using var content = new ByteArrayContent(fileContent);
            content.Headers.ContentType =
                new MediaTypeHeaderValue("application/octet-stream");

            var resp = client.PutAsync(apiUrl, content).GetAwaiter().GetResult();
            resp.EnsureSuccessStatusCode();
        }

        public static string GetFileName(
            string driveId,
            string itemId,
            string accessToken)
        {
            var apiUrl =
              $"https://graph.microsoft.com/v1.0/sites/{_siteId}" +
              $"/drives/{driveId}/items/{itemId}" +
              "?$select=name";

            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", accessToken);

            var resp = client.GetAsync(apiUrl).GetAwaiter().GetResult();
            resp.EnsureSuccessStatusCode();
            var json = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();

            dynamic obj = JsonConvert.DeserializeObject(json);
            return (string)obj.name;
        }
    }
}
