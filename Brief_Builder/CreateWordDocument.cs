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
        //public async Task<HttpResponseMessage> Run(
        //    [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestMessage req,
        //    ILogger log)
        public async Task<HttpResponseMessage> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post")] HttpRequestMessage req,
            ILogger log)
        {
            log.LogInformation("Processing CreateWordDocument request.");

            var json = await req.Content.ReadAsStringAsync();
            log.LogInformation("Request Body: {0}", json);
            var data = JsonConvert.DeserializeObject<BriefBuilderInfo>(json);
            if ((data?.EmailIds?.Any() != true)
             && (data?.Claims?.Any() != true))
            {
                return req.CreateErrorResponse(
                    HttpStatusCode.BadRequest,
                    "Please provide one or more emailIds or claims in the request body.");
            }

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
                foreach (var idString in data.EmailIds)
                {
                    if (!Guid.TryParse(idString, out var emailId))
                    {
                        log.LogWarning("Invalid GUID skipped: {0}", idString);
                        continue;
                    }

                        var email = _dataverse.RetrieveEmailRecord(emailId);
                        var slot = email.GetAttributeValue<string>("pace_slot_display_name") ?? string.Empty;
                        var from = ExtractParty(email.GetAttributeValue<EntityCollection>("from"));
                        var to = ExtractParty(email.GetAttributeValue<EntityCollection>("to"));
                        var body = HtmlHelper.StripHtml(
                                        email.GetAttributeValue<string>("description") ?? string.Empty)
                                    .Trim();
                        if (body.Length > 200)
                            body = body.Substring(0, 200) + "…";

                        emailInfos.Add(new EmailInfo
                        {
                            Id = emailId,
                            Name = slot,
                            From = from,
                            To = to,
                            Body = body
                        });

                        log.LogInformation("Prepared EmailInfo for {0}", emailId);
                }
            }

            var wordBytes = WordHelper.CreateDoc(
                claims: data.Claims?.SelectMany(d => d)
                       ?? Enumerable.Empty<KeyValuePair<string, string>>(),
                emails: emailInfos);

            log.LogInformation("Generated Word document: {0} bytes", wordBytes.Length);

            // --------------------------------------------------------
            // If you want to upload to SharePoint, uncomment this block
            // --------------------------------------------------------
            /*
            if (!string.IsNullOrEmpty(data.RecordId) &&
                Guid.TryParse(data.RecordId, out var claimId))
            {
                var locEntity = _dataverse.GetClaimDocumentLocation(claimId);
                var docLoc    = locEntity?.GetAttributeValue<string>("relativeurl");
                if (!string.IsNullOrEmpty(docLoc))
                {
                    var parts   = docLoc.Split(new[] {'/'}, 2);
                    var library = parts[0];
                    var folder  = parts.Length > 1 ? parts[1] : string.Empty;

                    var (token, siteId) = SharepointService.GetSPTokenAndSiteId();
                    var driveId = SharepointService.GetDriveId(library, siteId, token);

                    var fileName = $"Brief_{claimId}.docx";
                    var itemId   = SharepointService.UploadDocumentToSharePoint(
                        siteId, driveId, folder, fileName, wordBytes, token);

                    log.LogInformation("Uploaded to SharePoint, itemId={0}", itemId);
                }
                else
                {
                    log.LogWarning("No DocumentLocation found for Claim {0}", claimId);
                }
            }
            */
            // --------------------------------------------------------

            // 5) Return the .docx for download
            var response = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new ByteArrayContent(wordBytes)
            };
            response.Content.Headers.ContentType = MediaTypeHeaderValue.Parse(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = "BriefBuilderReport.docx"
            };
            return response;
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
