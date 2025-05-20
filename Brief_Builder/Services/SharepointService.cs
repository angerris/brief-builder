using Brief_Builder.Models;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;

namespace Brief_Builder.Services
{
    public static class SharepointService
    {
        private static readonly string _clientId = Environment.GetEnvironmentVariable("spClientID");
        private static readonly string _clientSecret = Environment.GetEnvironmentVariable("spSecret");
        private static readonly string _siteId = Environment.GetEnvironmentVariable("spSiteID");
        private static readonly string _tenantId = Environment.GetEnvironmentVariable("spTenantID");

        public static string GetDriveId(string displayName, string siteId, string accessToken)
        {
            var apiUrl = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drives";
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            var resp = client.GetAsync(apiUrl).GetAwaiter().GetResult();
            resp.EnsureSuccessStatusCode();

            var json = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();
            var root = JsonSerializer.Deserialize<SharepointDrives>(json);
            var drive = root.Value.FirstOrDefault(d => d.Name == displayName)
                        ?? throw new InvalidPluginExecutionException("Drive not found");
            return drive.Id;
        }

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
            var body = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();

            Console.WriteLine($"[Token] {resp.StatusCode}: {body}");

            resp.EnsureSuccessStatusCode();
            return JsonSerializer.Deserialize<TokenResponse>(body);
        }


        public static (string accessToken, string siteId) GetSPTokenAndSiteId()
        {
            var token = GetTokenResponse().Access_token;
            return (token, _siteId);
        }

        public static string UploadDocumentToSharePoint(
            string siteId,
            string driveId,
            string folderName,
            string fileName,
            byte[] fileContent,
            string accessToken)
        {
            var apiUrl = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drives/{driveId}/root:/{folderName}/{fileName}:/content";
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            using var content = new ByteArrayContent(fileContent);
            content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

            var resp = client.PutAsync(apiUrl, content).GetAwaiter().GetResult();
            resp.EnsureSuccessStatusCode();

            var json = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();
            using var doc = JsonDocument.Parse(json);
            var eTag = doc.RootElement.GetProperty("eTag").GetString();
            return eTag.Trim('{', '}');
        }
    }
}
