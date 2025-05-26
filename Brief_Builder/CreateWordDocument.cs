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
        private readonly DataverseService _dataverse;
        private readonly SharepointService _spService;
        private readonly WordHelper _wordHelper;
        public CreateWordDocument(IOrganizationService crmService)
        {
            _dataverse = new DataverseService(crmService);
            _spService = new SharepointService();
            _wordHelper = new WordHelper();

        }

        [FunctionName("CreateWordDocument")]
        public async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestMessage req)
        {
            var data = await ParseRequest(req);
            var emailInfos = _dataverse.BuildEmailInfos(data);
            var imported = await _spService.GetImportedFilesAsync(data.SharePointFiles);
            var wordBytes = GenerateWordDocument(data, emailInfos, imported);
            UploadToSharepoint(data, wordBytes);

            return new HttpResponseMessage(HttpStatusCode.OK);
        }

        private async Task<BriefBuilderInfo> ParseRequest(HttpRequestMessage req)
        {
            var json = await req.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<BriefBuilderInfo>(json);
        }

        private  byte[] GenerateWordDocument(BriefBuilderInfo data, List<EmailInfo> emailInfos, List<ImportedFile> imported)
        {
            return _wordHelper.CreateDoc(
                data.Claims?.SelectMany(d => d) ?? Enumerable.Empty<KeyValuePair<string, string>>(),
                 emailInfos, imported);
        }

        private async void UploadToSharepoint(BriefBuilderInfo data, byte[] wordBytes)
        {
            var loc = _dataverse.GetClaimDocumentLocation(Guid.Parse(data.RecordId));
            var folderPath = loc?.GetAttributeValue<string>("relativeurl") ?? "";
            var driveId = _spService.GetClaimDriveId();

            var docxFileName = $"Brief_Report_{DateTime.UtcNow:yyyyMMddHHmmss}.docx";
            var docxItemId = _spService.UploadDocumentToSharePoint(
              await driveId, folderPath, docxFileName, wordBytes);

            var pdfBytes = _spService.DownloadDocumentAsPDFFromSharePoint(await driveId, await docxItemId);

            var pdfFileName = $"Brief_Report_{DateTime.UtcNow:yyyyMMddHHmmss}.pdf";
                    _spService.UploadDocumentToSharePoint(
               await driveId, folderPath, pdfFileName, await pdfBytes);
        }
    }
}
