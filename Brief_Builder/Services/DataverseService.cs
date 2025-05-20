using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Brief_Builder.Services
{
    public class DataverseService
    {
        private readonly IOrganizationService _service;

        public DataverseService(IOrganizationService service)
        {
            _service = service;
        }

        public Entity RetrieveEmailRecord(Guid emailId)
        {
            var fields = new ColumnSet("pace_slot_display_name", "from", "to", "description");
            return _service.Retrieve("email", emailId, fields);
        }

        public Entity GetClaimDocumentLocation(Guid claimId)
        {
            var fetch = $@"
                <fetch top='1'>
                  <entity name='sharepointdocumentlocation'>
                    <attribute name='relativeurl' />
                    <filter>
                      <condition attribute='regardingobjectid' operator='eq' value='{claimId}' />
                    </filter>
                  </entity>
                </fetch>";
            var results = _service.RetrieveMultiple(new FetchExpression(fetch));
            return results.Entities.FirstOrDefault();
        }
    }
}
