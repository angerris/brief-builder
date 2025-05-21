using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Brief_Builder.Models;
using Brief_Builder.Services;
using Brief_Builder.Utils;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
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
        public async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestMessage req)
        {
            var data = await ParseRequest(req);

            var emailInfos = BuildEmailInfos(data);

            var wordBytes = GenerateWordDocument(data, emailInfos);

            UploadToSharepoint(data, wordBytes);

            var response = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new ByteArrayContent(wordBytes)
            };
         
            return response;
        }

        private async Task<BriefBuilderInfo> ParseRequest(HttpRequestMessage req)
        {
            var json = await req.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<BriefBuilderInfo>(json);
        }

        private List<EmailInfo> BuildEmailInfos(BriefBuilderInfo data)
        {
            var list = new List<EmailInfo>();
            if (data.EmailIds == null) return list;

            foreach (var id in data.EmailIds)
            {
                var emailId = Guid.Parse(id);
                var email = _dataverse.RetrieveEmailRecord(emailId);

                list.Add(new EmailInfo
                {
                    Id = emailId,
                    Name = email.GetAttributeValue<string>("pace_slot_display_name") ?? "",
                    From = ExtractParty(email.GetAttributeValue<EntityCollection>("from")),
                    To = ExtractParty(email.GetAttributeValue<EntityCollection>("to")),
                    Body = HtmlHelper.StripHtml(
                                email.GetAttributeValue<string>("description") ?? "")
                });
            }

            return list;
        }

        private byte[] GenerateWordDocument(
            BriefBuilderInfo data,
            List<EmailInfo> emailInfos)
        {
            return WordHelper.CreateDoc(
                claims: data.Claims?.SelectMany(d => d)
                       ?? Enumerable.Empty<KeyValuePair<string, string>>(),
                emails: emailInfos);
        }

        private void UploadToSharepoint(BriefBuilderInfo data, byte[] wordBytes)
        {

            var loc = _dataverse.GetClaimDocumentLocation(Guid.Parse(data.RecordId));
            var folderPath = loc?.GetAttributeValue<string>("relativeurl");

            var token = SharepointService.GetTokenResponse().AccessToken;
            var driveId = SharepointService.GetClaimDriveId(token);
            var fileName = $"Brief_Report.docx";

            SharepointService.UploadDocumentToSharePoint(driveId, folderPath, fileName, wordBytes, token);
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
