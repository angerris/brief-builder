using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Brief_Builder.Models;
using Brief_Builder.Services; 
using Brief_Builder.Utils;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;

namespace Brief_Builder
{
    public class CreateWordDocument
    {
        private readonly IOrganizationService _crmService;
        private readonly DataverseService _dataverse;

        public CreateWordDocument(IOrganizationService crmService)
        {
            _crmService = crmService;
            _dataverse = new DataverseService(crmService);
        }

        [FunctionName("CreateWordDocument")]
        public async Task<HttpResponseMessage> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post")] HttpRequestMessage req,
            ILogger log)
        {
            log.LogInformation("Processing CreateWordDocument request.");

            var json = await req.Content.ReadAsStringAsync();
            var data = JsonConvert.DeserializeObject<BriefBuilderInfo>(json);
            if ((data?.EmailIds?.Any() != true) && (data?.Claims?.Any() != true))
                return req.CreateErrorResponse(HttpStatusCode.BadRequest,
                    "Provide one or more EmailIds or Claims.");

            if (data.Claims?.Any() == true)
            {
                log.LogInformation("Logging {0} claim(s)", data.Claims.Count);
                foreach (var claim in data.Claims)
                    foreach (var kv in claim)
                        log.LogInformation("{0} = {1}", kv.Key, kv.Value);
            }

            var emailInfos = new List<EmailInfo>();
            if (data.EmailIds != null)
            {
                foreach (var id in data.EmailIds)
                {
                    if (!Guid.TryParse(id, out var emailId))
                    {
                        log.LogWarning("Skipping invalid GUID: {0}", id);
                        continue;
                    }
                    var email = _dataverse.RetrieveEmailRecord(emailId);
                    var slot = email.GetAttributeValue<string>("pace_slot_display_name") ?? "";
                    var from = ExtractParty(email.GetAttributeValue<EntityCollection>("from"));
                    var to = ExtractParty(email.GetAttributeValue<EntityCollection>("to"));
                    var body = HtmlHelper.StripHtml(
                                    email.GetAttributeValue<string>("description") ?? "").Trim();
                    if (body.Length > 200) body = body.Substring(0, 200) + "…";

                    emailInfos.Add(new EmailInfo
                    {
                        Id = emailId,
                        Name = slot,
                        From = from,
                        To = to,
                        Body = body
                    });
                }
            }

            var wordBytes = WordHelper.CreateDoc(
                claims: data.Claims?.SelectMany(d => d)
                       ?? Enumerable.Empty<KeyValuePair<string, string>>(),
                emails: emailInfos);

            log.LogInformation("Generated Word document: {0} bytes", wordBytes.Length);

            if (!string.IsNullOrEmpty(data.RecordId) &&
                Guid.TryParse(data.RecordId, out var claimId))
            {
                var loc = _dataverse.GetClaimDocumentLocation(claimId);
                var folderPath = loc?.GetAttributeValue<string>("relativeurl");
                if (!string.IsNullOrEmpty(folderPath))
                {
                    var (token, siteId) = SharepointService.GetSPTokenAndSiteId();

                    var driveId = SharepointService.GetClaimDriveId(token);

                    var fileName = $"Brief_{claimId}.docx";
                    SharepointService.UploadDocumentToSharePoint(
                        siteId, driveId, folderPath, fileName, wordBytes, token);

                    log.LogInformation("Uploaded to Claim drive folder='{0}'", folderPath);
                }
                else
                {
                    log.LogWarning("No DocumentLocation.relativeurl found for Claim {0}", claimId);
                }
            }

            var resp = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new ByteArrayContent(wordBytes)
            };
            resp.Content.Headers.ContentType = MediaTypeHeaderValue.Parse(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            resp.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = "BriefBuilderReport.docx"
            };
            return resp;
        }

        private static string ExtractParty(EntityCollection parties)
        {
            var arr = parties?.Entities
                .Select(p => p.GetAttributeValue<string>("addressused")
                           ?? p.GetAttributeValue<EntityReference>("partyid")?.Name)
                .Where(v => !string.IsNullOrEmpty(v))
                .ToArray();
            return (arr == null || arr.Length == 0)
                ? "<none>"
                : string.Join(", ", arr);
        }
    }
}
