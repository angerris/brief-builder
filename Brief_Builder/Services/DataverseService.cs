using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Brief_Builder.Models;
using Brief_Builder.Utils;

namespace Brief_Builder.Services
{
    public class DataverseService
    {
        private readonly IOrganizationService _service;
        private readonly HtmlHelper _htmlHelper;

        public DataverseService(IOrganizationService service)
        {
            _service = service;
            _htmlHelper = new HtmlHelper();
        }

        private Entity RetrieveEmailRecord(Guid emailId)
        {
            var fields = new ColumnSet("pace_slot_display_name", "from", "to", "description");
            return _service.Retrieve("email", emailId, fields);
        }

        public Entity GetClaimDocumentLocation(Guid claimId)
        {
            var fetch = $@"
                <fetch top='1'>
                  <entity name='sharepointdocumentlocation'>
                    <attribute name='relativeurl'/>
                    <filter>
                      <condition attribute='regardingobjectid'
                            operator='eq'
                            value='{claimId}'/>
                    </filter>
                  </entity>
                </fetch>";
            var results = _service.RetrieveMultiple(new FetchExpression(fetch));
            return results.Entities.FirstOrDefault();
        }

        public List<EmailInfo> BuildEmailInfos(BriefBuilderInfo data)
        {
            var list = new List<EmailInfo>();
            if (data.EmailIds == null) return list;

            foreach (var id in data.EmailIds)
            {
                var emailId = Guid.Parse(id);
                var email = RetrieveEmailRecord(emailId);
          
                list.Add(new EmailInfo
                {
                    Id = emailId,
                    Name = email.GetAttributeValue<string>("pace_slot_display_name") ?? "",
                    From = ExtractParty(email.GetAttributeValue<EntityCollection>("from")),
                    To = ExtractParty(email.GetAttributeValue<EntityCollection>("to")),
                    Body = _htmlHelper.StripHtml(
                                email.GetAttributeValue<string>("description") ?? "")
                });
            }
            return list;
        }

        private string ExtractParty(EntityCollection parties)
        {
            var arr = parties?.Entities
                .Select(p => p.GetAttributeValue<string>("addressused")
                           ?? p.GetAttributeValue<EntityReference>("partyid")?.Name)
                .Where(s => !string.IsNullOrEmpty(s))
                .ToArray();
            return (arr == null || arr.Length == 0)
                ? "<none>"
                : string.Join(", ", arr);
        }
    }
}
