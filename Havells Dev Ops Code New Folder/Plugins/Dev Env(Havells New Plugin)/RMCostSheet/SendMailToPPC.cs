using System;
using System.Net.NetworkInformation;
using System.Text;
using System.Web.Configuration;
using System.Web.UI.WebControls;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.RMCostSheet
{

    public class SendMailToPPC : IPlugin 
    {
        public static ITracingService tracingService = null;
        static public EntityReference sender = new EntityReference();
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                /* This plugin will be Triggerd on RM CostSheet Entity Onchange of Cogs Status fields If cogs status is Equal 1 Then send mail to PPC Team else if Cogs Status == 2 then send mail to Design Team */
                if (context.MessageName.ToLower() == "update" && context.InputParameters["Target"] is Entity
                    && context.InputParameters.Contains("Target") && context.Depth == 1)
                    tracingService.Trace("Plugin Execution Start");
                Entity entity = (Entity)context.InputParameters["Target"];
                int CogsStatus = entity.GetAttributeValue<OptionSetValue>("hil_cogsstatus").Value;
                string fetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                          <entity name='hil_rmcostsheet'>
                            <attribute name='hil_rmcostsheetid' />
                            <attribute name='hil_name' />
                            <attribute name='createdon' />
                            <attribute name='hil_email' />
                            <attribute name='hil_tenderid' />
                           <attribute name='hil_cogsstatus' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                              <condition attribute='statecode' operator='eq' value='0' />
                              <condition attribute='hil_rmcostsheetid' operator='eq' value='{entity.Id}' />
                            </filter>
                            <link-entity name='hil_rmcostsheetline' from='hil_rmcostsheetline' to='hil_rmcostsheetid' link-type='inner' alias='aa'>
                              <filter type='and'>
                                <condition attribute='hil_rmtype' operator='eq' value='2' />
                                <condition attribute='statecode' operator='eq' value='0' />
                              </filter>
                            </link-entity>
                          </entity>
                        </fetch>";

                EntityCollection rmsheetColl = service.RetrieveMultiple(new FetchExpression(fetch));
                if (rmsheetColl.Entities.Count > 0)
                {
                    tracingService.Trace("RM Cost Sheet" + rmsheetColl.Entities[0].Id.ToString());
                    EntityCollection entTOList = new EntityCollection();
                    EntityCollection entCCList = new EntityCollection();
                    tracingService.Trace("Retrive PPC & DesignTeam Member Function Started");
                    EntityCollection PPCTeamMembers = new EntityCollection();
                    EntityCollection DesignTeamMembers = new EntityCollection();

                    if (CogsStatus == 1) //PPC
                    {
                        PPCTeamMembers = retriveTeamMembers(service, "PPC", new EntityReference("hil_enquirydepartment", new Guid("ce8b92cb-e64c-ec11-8f8e-6045bd733e10"))); // Cable Department Guid
                        tracingService.Trace("PPC Team Members count : " + PPCTeamMembers.Entities.Count.ToString());

                        foreach (Entity toEntity in PPCTeamMembers.Entities)
                        {
                            Entity entTo = new Entity("activityparty");
                            entTo["partyid"] = toEntity.ToEntityReference();
                            entTOList.Entities.Add(entTo);

                        }
                    }
                    else if (CogsStatus == 2) //Design
                    {
                        Guid tenderId = rmsheetColl.Entities[0].GetAttributeValue<EntityReference>("hil_tenderid").Id;
                        Entity EntSalesOffice = service.Retrieve("hil_tender", tenderId, new ColumnSet("hil_salesoffice"));
                        EntityReference salsesOffice = EntSalesOffice.Contains("hil_salesoffice") ? EntSalesOffice.GetAttributeValue<EntityReference>("hil_salesoffice") : null;
                        if (salsesOffice != null)
                        {
                            QueryExpression query = new QueryExpression("hil_designteambranchmapping");
                            query.ColumnSet = new ColumnSet("hil_user");
                            query.Criteria = new FilterExpression(LogicalOperator.And);
                            query.Criteria.AddCondition(new ConditionExpression("hil_salesoffice", ConditionOperator.Equal, salsesOffice.Id));
                            query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));//Active
                            EntityCollection designteambranchmapping = service.RetrieveMultiple(query);
                            if (designteambranchmapping.Entities.Count > 0)
                            {
                                foreach (Entity toEntity in designteambranchmapping.Entities)
                                {
                                    EntityReference positionRef = toEntity.GetAttributeValue<EntityReference>("hil_user");
                                    Entity entity1 = service.Retrieve(positionRef.LogicalName, positionRef.Id, new ColumnSet(false));
                                    DesignTeamMembers.Entities.Add(entity1);
                                }
                            }
                            foreach (Entity toEntity in DesignTeamMembers.Entities)
                            {
                                Entity entTo = new Entity("activityparty");
                                entTo["partyid"] = toEntity.ToEntityReference();
                                entTOList.Entities.Add(entTo);
                            }
                        }
                    }

                    if (PPCTeamMembers.Entities.Count > 0 || DesignTeamMembers.Entities.Count > 0)
                    {
                        tracingService.Trace("Check condition for PPC and Design Team Members");
                        //EntityReference owner = rmsheetColl.Entities[0].GetAttributeValue<EntityReference>("ownerid");
                        string mailBody = BodyText(rmsheetColl.Entities[0].Id.ToString());
                        string subject = "RM Cost Sheet Email Testing By Anil Yadav";
                        EntityReference rmCostSheetRegarding = rmsheetColl.Entities[0].GetAttributeValue<EntityReference>("hil_tenderid");
                        tracingService.Trace("Enter send Email Function");
                        sendEmail(entTOList, entCCList, rmCostSheetRegarding, mailBody, subject, service);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error :- " + ex.Message);
            }
        }
        public string BodyText(string rmcostsheetId)
        {
            return $@"<!DOCTYPE html>
                    <html lang='en'>
                    <head>
                        <meta charset='UTF-8'>
                        <meta name='viewport' content='width=device-width, initial-scale=1.0'>
                        <title>RM Costsheet</title>
                    </head>
                    <body>
                        Dear PPC Team, Kindly provide Rm Cost sheet Rates of below items 
                        For more Information Please <a href='https://havellscrmdev1.crm8.dynamics.com/main.aspx?appid=675ffa44-b6d0-ea11-a813-000d3af05d7b&forceUCI=1&pagetype=entityrecord&etn=hil_rmcostsheet&id={rmcostsheetId}'> (Click Here)</a><br>
                    Regards,
                    System 
                    </body>
                    </html>";
        }
        private static void sendEmail(EntityCollection too, EntityCollection copyto, EntityReference regarding, string mailbody, string subject, IOrganizationService service)
        {
            try
            {
                Entity entEmail = new Entity("email");
                Entity entFrom = new Entity("activityparty");
                entFrom["partyid"] = new EntityReference("queue", new Guid("a81ac51b-c9df-eb11-bacb-0022486ea537")); // EMS Queue Guid
                Entity[] entFromList = { entFrom };
                entEmail["from"] = entFromList;
                Entity toActivityParty = new Entity("activityparty");
                toActivityParty["partyid"] = too;
                entEmail["to"] = too;
                tracingService.Trace("sendEmail if condition started");
                if (copyto.Entities.Count > 0)
                {
                    Entity ccActivityParty = new Entity("activityparty");
                    ccActivityParty["partyid"] = copyto;
                    entEmail["cc"] = copyto;
                }

                entEmail["subject"] = subject;
                entEmail["description"] = mailbody;
                entEmail["regardingobjectid"] = regarding;
                Guid emailId = service.Create(entEmail);
                tracingService.Trace("Record is Created Sucessfully on EmailMessage Entity");
                SendEmailRequest sendEmailReq = new SendEmailRequest()
                {
                    EmailId = emailId,
                    IssueSend = true
                };
                SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);
                tracingService.Trace("Email sent Sucessfully to PPC or Design TeamMembers check your outlook Inbox " + sendEmailReq);

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error :- " + ex.Message);
            }
        }
        static public EntityCollection retriveTeamMembers(IOrganizationService service, string _teamName, EntityReference _department)
        {
            EntityCollection extTeamMembers = new EntityCollection();
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
                        tracingService.Trace("bdteamCol count " + bdteamCol.Entities.Count);
                        tracingService.Trace("bdteamCol.Entities[0].Id " + bdteamCol.Entities[0].Id);
                        QueryExpression _querymem = new QueryExpression("hil_bdteammember");
                        _querymem.ColumnSet = new ColumnSet("emailaddress");
                        _querymem.Criteria = new FilterExpression(LogicalOperator.And);
                        _querymem.Criteria.AddCondition(new ConditionExpression("hil_team", ConditionOperator.Equal, bdteamCol.Entities[0].Id));
                        _querymem.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                        EntityCollection bdteammemCol = service.RetrieveMultiple(_querymem);
                        tracingService.Trace("bdteammemCol count" + bdteammemCol.Entities.Count);
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

    }
}