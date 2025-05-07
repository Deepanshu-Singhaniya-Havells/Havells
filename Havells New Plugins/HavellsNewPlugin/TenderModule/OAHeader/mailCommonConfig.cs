using System;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.OAHeader
{
    public static class mailCommonConfig
    {
        public static Entity getTeamUser(EntityReference team, IOrganizationService service)
        {
            ColumnSet coll = new ColumnSet("hil_tenderemailersetupid",
                                            "hil_name",
                                            "createdon",
                                            "hil_ppc",
                                            "hil_inspectionteam2",
                                            "hil_inspectionteam1",
                                            "hil_floorhead4",
                                            "hil_floorhead3",
                                            "hil_floorhead2",
                                            "hil_floorhead1",
                                            "hil_cmt");
            Entity emailr = service.Retrieve(team.LogicalName, team.Id, coll);
            return emailr;
        }

        public static Entity getUserConfiguartion(EntityReference owner, IOrganizationService service)
        {
            QueryExpression _query = new QueryExpression("hil_userbranchmapping");
            _query.ColumnSet = new ColumnSet("hil_name", "hil_zonalhead", "hil_user", "hil_salesoffice", "hil_buhead", "hil_branchproducthead");
            _query.Criteria = new FilterExpression(LogicalOperator.And);
            _query.Criteria.AddCondition(new ConditionExpression("hil_user", ConditionOperator.Equal, owner.Id));
            _query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));

            EntityCollection userMapingColl = service.RetrieveMultiple(_query);
            if (userMapingColl.Entities.Count > 0)
                return userMapingColl[0];
            else
                return new Entity();
        }
        public static void sendEmal(EntityReference approver, EntityCollection copyto, EntityReference regarding, string mailbody, string subject, IOrganizationService service)
        {
            try
            {
                Entity entEmail = new Entity("email");

                Entity entFrom = new Entity("activityparty");
                entFrom["partyid"] = new EntityReference("queue", new Guid("a81ac51b-c9df-eb11-bacb-0022486ea537"));
                Entity[] entFromList = { entFrom };
                entEmail["from"] = entFromList;

                EntityReference to = approver;
                Entity toActivityParty = new Entity("activityparty");
                toActivityParty["partyid"] = to;
                entEmail["to"] = new Entity[] { toActivityParty };

                
                Entity ccActivityParty = new Entity("activityparty");
                ccActivityParty["partyid"] = copyto;
                entEmail["cc"] = copyto;

                entEmail["subject"] = subject;
                entEmail["description"] = mailbody;

                entEmail["regardingobjectid"] = regarding;

                Guid emailId = service.Create(entEmail);

                SendEmailRequest sendEmailReq = new SendEmailRequest()
                {
                    EmailId = emailId,
                    IssueSend = true
                };
                SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error: " + ex.Message);
            }
        }
    }
}
