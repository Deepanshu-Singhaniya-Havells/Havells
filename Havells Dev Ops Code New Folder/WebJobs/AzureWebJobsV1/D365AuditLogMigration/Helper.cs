using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Crm.Sdk.Messages;

namespace D365AuditLogMigration
{
    public class Helper
    {

        public static Guid GetGuidbyName(String sEntityName, String sFieldName, String sFieldValue, IOrganizationService service, int iStatusCode = 0)
        {
            Guid fsResult = Guid.Empty;
            try
            {
                QueryExpression qe = new QueryExpression(sEntityName);
                qe.Criteria.AddCondition(sFieldName, ConditionOperator.Equal, sFieldValue);
                qe.AddOrder("createdon", OrderType.Descending);
                if (iStatusCode >= 0)
                {
                    qe.Criteria.AddCondition("statecode", ConditionOperator.Equal, iStatusCode);
                }
                EntityCollection enColl = service.RetrieveMultiple(qe);
                if (enColl.Entities.Count > 0)
                {
                    fsResult = enColl.Entities[0].Id;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Helper.GetGuidbyName" + ex.Message);
            }
            return fsResult;
        }

        public static Guid GetGuidbyNameInvoiceOnly(String sEntityName, String sFieldName, String sFieldValue, IOrganizationService service)
        {
            Guid fsResult = Guid.Empty;
            try
            {
                QueryExpression qe = new QueryExpression(sEntityName);
                qe.Criteria.AddCondition(sFieldName, ConditionOperator.Equal, sFieldValue);
                qe.AddOrder("createdon", OrderType.Descending);
                EntityCollection enColl = service.RetrieveMultiple(qe);
                if (enColl.Entities.Count > 0)
                {
                    fsResult = enColl.Entities[0].Id;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Helper.GetGuidbyName" + ex.Message);
            }
            return fsResult;
        }

        public static void Assign(String AssigneeLogicalName, String TargetLogicalName, Guid AssigneeId, Guid TargetId, IOrganizationService service)
        {
            try
            {
                AssignRequest assign = new AssignRequest();
                assign.Assignee = new EntityReference(AssigneeLogicalName, AssigneeId); //User or team
                assign.Target = new EntityReference(TargetLogicalName, TargetId); //Record to be assigned
                service.Execute(assign);

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Helper.Assign" + ex.Message);
            }
        }

        public static OptionSetValue GetAccountType(Guid fsUser, IOrganizationService service)
        {
            OptionSetValue result = null;
            try
            {
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    var obj = from _User in orgContext.CreateQuery<SystemUser>()
                              join _Account in orgContext.CreateQuery<Account>() on _User.hil_Account.Id equals _Account.Id
                              where _User.SystemUserId.Value == fsUser
                              select new { _Account.CustomerTypeCode };
                    foreach (var iobj in obj)
                    {
                        result = iobj.CustomerTypeCode;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Helper.GetAccountType" + ex.Message);
            }
            return result;
        }

        public static void SharetoAccessTeam(EntityReference enRef, Guid fsUserId, String sAccessTeamName, IOrganizationService service)
        {
            try
            {
                Guid fsAccessTeamTemplate = Helper.GetGuidbyName(TeamTemplate.EntityLogicalName, "teamtemplatename", sAccessTeamName, service, -1);
                if (fsAccessTeamTemplate != Guid.Empty)
                {
                    AddUserToRecordTeamRequest addReq = new AddUserToRecordTeamRequest();
                    addReq.Record = enRef;
                    addReq.SystemUserId = fsUserId;
                    addReq.TeamTemplateId = fsAccessTeamTemplate;
                    service.Execute(addReq);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Helper.SharetoAccessTeam" + ex.Message);
            }
        }
        public static void SendSMS(IOrganizationService service, string MobileNum, string Message, string Subject, Guid ContactId)
        {
            hil_smsconfiguration SMS = new hil_smsconfiguration();
            SMS.hil_Direction = new OptionSetValue(1);
            SMS.hil_Message = Message;
            SMS.hil_MobileNumber = MobileNum;
            SMS.Subject = "KKG Code SMS Jobs Create";
            SMS.hil_contact = new EntityReference(Contact.EntityLogicalName, ContactId);
            service.Create(SMS);
        }
    }
}
