using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.WebJob.BusinessManagement.AutoMailer
{
    public class Helper
    {
        public static Integration IntegrationConfiguration(IOrganizationService service, string Param)
        {
            Integration output = new Integration();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                if (integrationConfiguration.Contains("hil_username") && integrationConfiguration.Contains("hil_password"))
                {
                    output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + ":" + integrationConfiguration.GetAttributeValue<string>("hil_password");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:- " + ex.Message);
            }
            return output;
        }
        public static EntityReference getSender(string queName, IOrganizationService service)
        {
            EntityReference sender = new EntityReference();
            QueryExpression _queQuery = new QueryExpression("queue");
            _queQuery.ColumnSet = new ColumnSet("name");
            _queQuery.Criteria = new FilterExpression(LogicalOperator.And);
            _queQuery.Criteria.AddCondition("name", ConditionOperator.Equal, queName);
            EntityCollection queueColl = service.RetrieveMultiple(_queQuery);
            if (queueColl.Entities.Count == 1)
            {
                sender = queueColl[0].ToEntityReference();
            }
            return sender;
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
        static public EntityCollection retriveTeamMembers(IOrganizationService service, string _teamName, EntityReferenceCollection _materialGroup, EntityReference _department,
           EntityReference _plant, EntityCollection extTeamMembers)
        {

            try
            {
                List<Guid> materialGuids = new List<Guid>();
                foreach (EntityReference entityReference in _materialGroup)
                {
                    materialGuids.Add(entityReference.Id);
                }
                QueryExpression _query = new QueryExpression("hil_bdteam");
                _query.ColumnSet = new ColumnSet("hil_name", "hil_materialgroup", "hil_department", "hil_plant");
                _query.Criteria = new FilterExpression(LogicalOperator.And);
                if (_teamName != null)
                    _query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, _teamName));
                if (_department != null)
                    _query.Criteria.AddCondition(new ConditionExpression("hil_department", ConditionOperator.Equal, _department.Id));
                EntityCollection bdteamCol = service.RetrieveMultiple(_query);
                if (bdteamCol.Entities.Count > 0)
                {
                    _query = new QueryExpression("hil_bdteam");
                    _query.ColumnSet = new ColumnSet(false);
                    _query.Criteria = new FilterExpression(LogicalOperator.And);
                    if (_teamName != null && bdteamCol[0].Contains("hil_name"))
                        _query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, _teamName));
                    if (materialGuids.Count > 0 && bdteamCol[0].Contains("hil_materialgroup"))
                        _query.Criteria.AddCondition(new ConditionExpression("hil_materialgroup", ConditionOperator.In, materialGuids.ToArray()));
                    if (_department != null && bdteamCol[0].Contains("hil_department"))
                        _query.Criteria.AddCondition(new ConditionExpression("hil_department", ConditionOperator.Equal, _department.Id));
                    if (_plant != null && bdteamCol[0].Contains("hil_plant"))
                        _query.Criteria.AddCondition(new ConditionExpression("hil_plant", ConditionOperator.Equal, _plant.Id));
                    bdteamCol = service.RetrieveMultiple(_query);
                    if (bdteamCol.Entities.Count > 0)
                    {
                        Console.WriteLine("bdteamCol count " + bdteamCol.Entities.Count);
                        QueryExpression _querymem = new QueryExpression("hil_bdteammember");
                        _querymem.ColumnSet = new ColumnSet("emailaddress");
                        _querymem.Criteria = new FilterExpression(LogicalOperator.And);
                        _querymem.Criteria.AddCondition(new ConditionExpression("hil_team", ConditionOperator.Equal, bdteamCol.Entities[0].Id));
                        EntityCollection bdteammemCol = service.RetrieveMultiple(_querymem);
                        EntityCollection entTOList = new EntityCollection();
                        Console.WriteLine("Team Members count" + entTOList.Entities.Count);
                        if (bdteammemCol.Entities.Count > 0)
                        {
                            foreach (Entity entity in bdteammemCol.Entities)
                            {
                                extTeamMembers.Entities.Add(entity);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error in Retriving Team Members : " + ex.Message);
            }
            return extTeamMembers;
        }
        public static void sendEmail(string mailBody, string subject, EntityReference regarding, EntityCollection to, EntityCollection cc, EntityReference senderID, IOrganizationService service)
        {
            try
            {
                Console.WriteLine("sending email.");
                Entity entEmail = new Entity("email");
                entEmail["subject"] = subject;
                entEmail["description"] = mailBody;
                entEmail["to"] = to;
                if (cc != null)
                    if (cc.Entities.Count > 0)
                        entEmail["cc"] = cc;

                Entity entFrom = new Entity("activityparty");
                entFrom["partyid"] = senderID;
                Entity[] entFromList = { entFrom };
                entEmail["from"] = entFromList;
                if (regarding != null)
                    entEmail["regardingobjectid"] = regarding;

                Guid emailId = service.Create(entEmail);
                Console.WriteLine("email is created with GUID:- " + emailId);
                SendEmailRequest sendEmailReq = new SendEmailRequest()
                {
                    EmailId = emailId,
                    IssueSend = true
                };
               // SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);
                Console.WriteLine("email is sended");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
    public class Integration
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }

}
