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
    public class BOEReminder
    {
        public static void SendBOEReminder(IOrganizationService service)
        {
            Console.WriteLine("********** BOE Reminder Mail Send  Program Started **********");
            string FetchQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='hil_oaheader'>
                                <attribute name='hil_oaheaderid' />
                                <attribute name='hil_name' />
                                <attribute name='hil_billlodgmentdate' />
                                <attribute name='hil_lcissuingbank' />
                                <attribute name='hil_lcnumber' />
                                <attribute name='hil_lcdate' />
                                <attribute name='hil_boevalue' />
                                <attribute name='hil_duedate' />
                                <attribute name='hil_orderchecklistid' />
                                <attribute name='hil_omsemailersetup' />
                                <attribute name='hil_department' />
                               hil_department
                                <order attribute='hil_name' descending='false' />
                                <filter type='and'>
                                  <condition attribute='hil_boeacceptance' operator='ne' value='1' />
                                  <condition attribute='hil_boevalue' operator='not-null' />
                                  <condition attribute='hil_lcnumber' operator='not-null' />
                                   </filter>
                                <link-entity name='hil_orderchecklist' from='hil_orderchecklistid' to='hil_orderchecklistid' link-type='inner' alias='ab'>
                                    <attribute name='hil_zonalhead' />
                                    <attribute name='hil_rm' />
                                    <attribute name='ownerid' />
                                    <attribute name='hil_buhead' />
                                  <filter type='and'>
                                    <condition attribute='hil_paymentterms' operator='eq' value='1' />
                                    <condition attribute='statecode' operator='eq' value='0' />
                                  </filter>
                                </link-entity>
                              </entity>
                            </fetch>";

            EntityCollection colBoeReminder = service.RetrieveMultiple(new FetchExpression(FetchQuery));
            if (colBoeReminder.Entities.Count > 0)
            {
                int counter = 1;
                int totalCount = colBoeReminder.Entities.Count;
                Console.WriteLine("Today BOE Reminder are " + totalCount);
                foreach (Entity entity in colBoeReminder.Entities)
                {
                    EntityReference Ownerid = null;
                    EntityReference ZonalHead = null;
                    EntityReference RM = null;
                    EntityReference BuHead = null;
                    // if (entity.Contains("ab.hil_zonalhead") && entity.GetAttributeValue<string>("hil_name") == "0009357355")
                    //{
                    ZonalHead = (EntityReference)entity.GetAttributeValue<AliasedValue>("ab.hil_zonalhead").Value;
                    if (entity.Contains("ab.ownerid"))
                    {
                        Ownerid = (EntityReference)entity.GetAttributeValue<AliasedValue>("ab.ownerid").Value;
                    }
                    if (entity.Contains("ab.hil_rm"))
                    {
                        RM = (EntityReference)entity.GetAttributeValue<AliasedValue>("ab.hil_rm").Value;
                    }
                    if (entity.Contains("ab.hil_buhead"))
                    {
                        BuHead = (EntityReference)entity.GetAttributeValue<AliasedValue>("ab.hil_buhead").Value;
                    }
                    string bodtText = mailBody(entity, service);
                    EntityCollection TeamMembers = new EntityCollection();
                    EntityReference department = entity.GetAttributeValue<EntityReference>("hil_department");
                    TeamMembers = retriveTeamMembers(service, "BOE Reminder", department, TeamMembers);
                    EntityCollection entCCList = new EntityCollection();
                    EntityCollection entTOList = new EntityCollection();
                    Entity entTO = new Entity("activityparty");
                    entTO["partyid"] = RM;
                    entTOList.Entities.Add(entTO);
                    entTO = new Entity("activityparty");
                    entTO["partyid"] = Ownerid;
                    entTOList.Entities.Add(entTO);

                    Entity entCC = new Entity("activityparty");
                    entCC["partyid"] = ZonalHead;
                    entCCList.Entities.Add(entCC);
                    entCC = new Entity("activityparty");
                    entCC["partyid"] = BuHead;
                    entCCList.Entities.Add(entCC);
                    foreach (Entity ccEntity in TeamMembers.Entities)
                    {
                        Entity entccc = new Entity("activityparty");
                        entccc["partyid"] = ccEntity.ToEntityReference();
                        entCCList.Entities.Add(entccc);
                    }
                    string subject = @"Accepatance of BOE document pending in Terms of LC";
                    EntityReference oaheaderregarding = entity.ToEntityReference();
                    sendEmail(bodtText, subject, oaheaderregarding, entTOList, entCCList, service);
                    //}
                    Console.WriteLine("BOE Reminder Mail Send  " + counter +"/"+ totalCount);
                    counter++;
                }
            }
            Console.WriteLine("********* BOE Reminder Mail Send  Progerm End *********");
        }
        static public EntityCollection retriveTeamMembers(IOrganizationService service, string _teamName, EntityReference _department, EntityCollection extTeamMembers)
        {
            try
            {
                QueryExpression _query = new QueryExpression("hil_bdteam");
                _query.ColumnSet = new ColumnSet("hil_name", "hil_materialgroup", "hil_department", "hil_plant", "hil_bdteamid");
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
                    if (_department != null && bdteamCol[0].Contains("hil_department"))
                        _query.Criteria.AddCondition(new ConditionExpression("hil_department", ConditionOperator.Equal, _department.Id));
                    bdteamCol = service.RetrieveMultiple(_query);
                    if (bdteamCol.Entities.Count > 0)
                    {

                        QueryExpression _querymem = new QueryExpression("hil_bdteammember");
                        _querymem.ColumnSet = new ColumnSet("emailaddress");
                        _querymem.Criteria = new FilterExpression(LogicalOperator.And);
                        _querymem.Criteria.AddCondition(new ConditionExpression("hil_team", ConditionOperator.Equal, bdteamCol.Entities[0].Id));
                        _querymem.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                        EntityCollection bdteammemCol = service.RetrieveMultiple(_querymem);
                        EntityCollection entTOList = new EntityCollection();

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
        public static void sendEmail(string mailBody, string subject, EntityReference regarding, EntityCollection to, EntityCollection cc, IOrganizationService _service)
        {
            try
            {
                Entity entFrom = new Entity("activityparty");
                entFrom["partyid"] = new EntityReference("queue", new Guid("b6f0037d-3e93-ec11-b400-6045bdaad0b5"));
                Entity[] entFromList = { entFrom };

                Entity entEmail = new Entity("email");
                entEmail["subject"] = subject;
                entEmail["description"] = mailBody;
                entEmail["to"] = to;
                if (cc != null)
                    if (cc.Entities.Count > 0)
                        entEmail["cc"] = cc;
                entEmail["from"] = entFromList;
                if (regarding.Id != Guid.Empty)
                    entEmail["regardingobjectid"] = regarding;
                Guid emailId = _service.Create(entEmail);
                SendEmailRequest sendEmailReq = new SendEmailRequest()
                {
                    EmailId = emailId,
                    IssueSend = true
                };
                SendEmailResponse sendEmailRes = (SendEmailResponse)_service.Execute(sendEmailReq);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static string mailBody(Entity entBoe, IOrganizationService service)
        {
            QueryExpression tenderURL = new QueryExpression("hil_integrationconfiguration");
            tenderURL.ColumnSet = new ColumnSet("hil_url");
            tenderURL.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "TenderAppURL");
            EntityCollection tenderURLCol = service.RetrieveMultiple(tenderURL);
            string URL = tenderURLCol.Entities[0].GetAttributeValue<string>("hil_url");

            StringBuilder sbtopBody = new StringBuilder();
            StringBuilder sbtable = new StringBuilder();
            StringBuilder sbbelowBody = new StringBuilder();
            sbtopBody.Append("<Div style='width:703.9pt;margin-left:-.15pt;font-weight:bold'>Dear Sales Team,</Div>");
            sbtopBody.Append("<Div><p style='width:703.9pt;margin-left:-.15pt;margin-bottom:5pt;font-weight:bold'>Acceptance against below given LC documents submited to the Bank is yet to receive. You are requested to kindly arrange the Acceptance from the customer banker to our bank for our further needful.</ p></Div>");
            sbtable.Append("<Div>");
            sbtable.Append("<table border=1 cellspacing=0 cellpadding=0 width=0 style='width:703.9pt;margin-left:-.15pt;border-collapse:collapse'>");
            sbtable.Append("<tr style='height:24.0pt;font-weight:bold;'><th> OA Number </th><th>OCL No</th><th>LC Number</th><th>LC Date</th><th>BOE Value</th><th>Bill Lodgment Date </th><th>LC Issuing Bank </th><th>Due Date </th></tr>");

            string partialUrl = URL + entBoe.LogicalName + "&id=" + entBoe.Id;
            string OaNumber = entBoe.GetAttributeValue<string>("hil_name");
            string OCLNo = entBoe.GetAttributeValue<EntityReference>("hil_orderchecklistid").Name;
            string LCNo = "";
            string LCDate = "";
            string BOEValue = "";
            string BillLodgmentDate = "";
            string LCIssuingbank = "";
            string Duedate = "";
            if (entBoe.Contains("hil_lcnumber"))
            {
                LCNo = entBoe.GetAttributeValue<string>("hil_lcnumber");
            }
            if (entBoe.Contains("hil_lcdate"))
            {
                LCDate = dateFormatter(entBoe.GetAttributeValue<DateTime>("hil_lcdate").AddMinutes(330)).ToString();
            }
            if (entBoe.Contains("hil_boevalue"))
            {
                BOEValue = entBoe.GetAttributeValue<Money>("hil_boevalue").Value.ToString("0.00");
            }
            if (entBoe.Contains("hil_billlodgmentdate"))
            {
                BillLodgmentDate = dateFormatter(entBoe.GetAttributeValue<DateTime>("hil_lcdate").AddMinutes(330)).ToString();
            }
            if (entBoe.Contains("hil_lcissuingbank"))
            {
                LCIssuingbank = entBoe.GetAttributeValue<string>("hil_lcissuingbank").ToString();
            }
            if (entBoe.Contains("hil_duedate"))
            {
                Duedate = dateFormatter(entBoe.GetAttributeValue<DateTime>("hil_duedate").AddMinutes(330)).ToString();
            }
            sbtable.Append("<tr style='height:24.0pt;text-align:center'><td>" + OaNumber + "</td><td>" + OCLNo + "</td><td>" + LCNo + "</td><td>" + LCDate + "</td><td>" + BOEValue + "</td><td>" + BillLodgmentDate + "</td><td>" + LCIssuingbank + "</td><td>" + Duedate + "</td></tr>");

            sbtable.Append("</table></Div>");
            sbbelowBody.Append("<Div style='margin-top:20pt;font-weight:bold'><p>Kind Regards </p></Div>");
            sbbelowBody.Append("<Div style='margin-top:20pt;font-weight:bold'><p>SMS </p></Div>");
            return sbtopBody.ToString() + sbtable.ToString() + sbbelowBody.ToString();
        }
        public static string dateFormatter(DateTime date)
        {
            string _enquiryDatetime = string.Empty;
            if (date.Year.ToString().PadLeft(4, '0') != "0001")
                _enquiryDatetime = date.Year.ToString() + "-" + date.Month.ToString().PadLeft(2, '0') + "-" + date.Day.ToString().PadLeft(2, '0');
            return _enquiryDatetime;
        }
    }

}
