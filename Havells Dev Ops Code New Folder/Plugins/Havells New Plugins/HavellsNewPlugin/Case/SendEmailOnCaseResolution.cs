using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.Case
{
    internal class SendEmailOnCaseResolution
    {
        private readonly IOrganizationService service;
        private readonly IPluginExecutionContext context;
        internal SendEmailOnCaseResolution(IOrganizationService _service)
        {
            service = _service;
        }
        internal void SendEmail(Entity Case)
        {

            string CaseNumber = Case.Contains("ticketnumber") ? Case.GetAttributeValue<string>("ticketnumber") : null;
            string caseTitle = Case.Contains("title") ? Case.GetAttributeValue<string>("title") : null;
            string resolutoinDate = Case.Contains("adx_resolutiondate") ? Case.GetAttributeValue<DateTime>("adx_resolutiondate").ToString() : null;

            string recordURL = $"https://havells.crm8.dynamics.com/main.aspx?appid=668eb624-0610-e911-a94e-000d3af06a98&forceUCI=1&pagetype=entityrecord&etn=incident&id={Case.Id}";

            Entity fromActivityParty = new Entity("activityparty");
            Entity toActivityParty = new Entity("activityparty");

            List<EntityReference> owners = Owners(Case.Id);

            if (owners.Count > 0)
            {
                fromActivityParty["partyid"] = new EntityReference("queue", new Guid("9b0480a8-e30f-e911-a94e-000d3af06a98"));
                toActivityParty["partyid"] = owners[0];

                Entity[] cc = new Entity[owners.Count - 1];
                for (int i = 1; i < owners.Count; i++)
                {
                    Entity tempActivityParty = new Entity("activityparty");
                    tempActivityParty["partyid"] = owners[i];
                    cc[i - 1] = tempActivityParty;
                }
                Entity email = new Entity("email");
                email["from"] = new Entity[] { fromActivityParty };
                email["to"] = new Entity[] { toActivityParty };
                if (owners.Count - 1 > 0) email["cc"] = cc;

                email["regardingobjectid"] = Case.ToEntityReference();
                email["subject"] = $"Resolutoin of case with case nubmer {CaseNumber}";
                email["description"] = $"Dear Team, <br/><br/> The case with case-ID {CaseNumber} regarding {caseTitle} has been resolved on {resolutoinDate}. <br/><br/> To open the case <a href={recordURL}>Click Here</a> <br/> <br/> Regards Team CRM";
                email["directioncode"] = true;
                Guid emailId = service.Create(email);

                // Use the SendEmail message to send an e-mail message.
                SendEmailRequest sendEmailRequest = new SendEmailRequest
                {
                    EmailId = emailId,
                    TrackingToken = "",
                    IssueSend = true
                };

                SendEmailResponse sendEmailresp = (SendEmailResponse)service.Execute(sendEmailRequest);
            }
        }
        private List<EntityReference> Owners(Guid CaseId)
        {
            List<EntityReference> ownersList = new List<EntityReference>();
            QueryExpression query = new QueryExpression("hil_grievancehandlingactivity");
            query.ColumnSet = new ColumnSet("ownerid");
            query.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, CaseId);
            query.Criteria.AddCondition("hil_activitytype", ConditionOperator.NotNull);
            query.AddOrder("createdon", OrderType.Ascending);

            EntityCollection ownerColl = service.RetrieveMultiple(query);
            if (ownerColl.Entities.Count > 0)
            {
                foreach (var entity in ownerColl.Entities)
                {
                    EntityReference tempOwner = entity.GetAttributeValue<EntityReference>("ownerid");
                    ownersList.Add(tempOwner);
                }
            }

            return ownersList;
        }
    }
}
