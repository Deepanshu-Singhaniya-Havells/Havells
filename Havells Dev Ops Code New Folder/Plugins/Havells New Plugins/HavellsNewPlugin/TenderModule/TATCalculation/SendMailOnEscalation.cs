using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using Microsoft.Crm.Sdk.Messages;
using System.Collections.Generic;

namespace HavellsNewPlugin.TenderModule.TATCalculation
{
    public class SendMailOnEscalation : IPlugin
    {
        public static ITracingService tracingService = null;

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
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {


                    Entity entity = (Entity)context.InputParameters["Target"];
                    Entity EscalationMatrix = service.Retrieve(entity.GetAttributeValue<EntityReference>("hil_escalationmatrix").LogicalName, entity.GetAttributeValue<EntityReference>("hil_escalationmatrix").Id, new ColumnSet("hil_tomailpoition", "hil_copytoposition", "hil_subjectline", "hil_mailbody"));
                    Entity TatOwnerShip = service.Retrieve(entity.GetAttributeValue<EntityReference>("hil_bdtatownership").LogicalName, entity.GetAttributeValue<EntityReference>("hil_bdtatownership").Id, new ColumnSet("regardingobjectid", "hil_department"));
                    Entity tender = service.Retrieve(TatOwnerShip.GetAttributeValue<EntityReference>("regardingobjectid").LogicalName, TatOwnerShip.GetAttributeValue<EntityReference>("regardingobjectid").Id, new ColumnSet("ownerid"));
                    EntityReference toownerid = tender.GetAttributeValue<EntityReference>("ownerid");

                    string subject = EscalationMatrix.GetAttributeValue<string>("hil_subjectline");
                    string body = EscalationMatrix.GetAttributeValue<string>("hil_mailbody");
                    string mailTO = EscalationMatrix.GetAttributeValue<string>("hil_tomailpoition");
                    string copymail = EscalationMatrix.GetAttributeValue<string>("hil_copytoposition");

                    sendEmail(service, mailTO, copymail, body, subject, toownerid, tender, tracingService);

                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error :- " + ex.Message);
            }
        }
        static public EntityCollection retriveTeamMembers(IOrganizationService service, string _teamName,
           EntityReference _department, EntityCollection extTeamMembers)
        {
            string Trace = string.Empty;
            try
            {

                QueryExpression _query = new QueryExpression("hil_bdteam");
                _query.ColumnSet = new ColumnSet("hil_name", "hil_department");
                _query.Criteria = new FilterExpression(LogicalOperator.And);
                if (_teamName != null)
                    _query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, _teamName));
                if (_department != null)
                    _query.Criteria.AddCondition(new ConditionExpression("hil_department", ConditionOperator.Equal, _department.Id));
                EntityCollection bdteamCol = service.RetrieveMultiple(_query);

                if (bdteamCol.Entities.Count > 0)
                {
                    Trace = Trace + bdteamCol.Entities.Count.ToString();
                    _query = new QueryExpression("hil_bdteam");
                    _query.ColumnSet = new ColumnSet(false);
                    _query.Criteria = new FilterExpression(LogicalOperator.And);
                    if (_teamName != null && bdteamCol[0].Contains("hil_name"))
                        _query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, _teamName));
                    if (_department != null && bdteamCol[0].Contains("hil_department"))
                        _query.Criteria.AddCondition(new ConditionExpression("hil_department", ConditionOperator.Equal, _department.Id));
                    bdteamCol = service.RetrieveMultiple(_query);
                    Trace = Trace + " bdteamCol.Entities.Count.ToString() " + bdteamCol.Entities.Count.ToString();
                    if (bdteamCol.Entities.Count > 0)
                    {
                        tracingService.Trace("bdteamCol count " + bdteamCol.Entities.Count);
                        QueryExpression _querymem = new QueryExpression("hil_bdteammember");
                        _querymem.ColumnSet = new ColumnSet("emailaddress");
                        _querymem.Criteria = new FilterExpression(LogicalOperator.And);
                        _querymem.Criteria.AddCondition(new ConditionExpression("hil_team", ConditionOperator.Equal, bdteamCol.Entities[0].Id));
                        //  _querymem.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 0));
                        EntityCollection bdteammemCol = service.RetrieveMultiple(_querymem);
                        EntityCollection entTOList = new EntityCollection();
                        tracingService.Trace("Team Members count" + entTOList.Entities.Count);
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
                throw new InvalidPluginExecutionException("Error in Retriving Team Members : " + ex.Message + "Trace - " + Trace);
            }
            return extTeamMembers;
        }
        public static void sendEmail(IOrganizationService service, string to, string cc, string mailBody, string mailsubject, EntityReference toownerid, Entity tender, ITracingService tracingService)
        {
            EntityCollection entCCList = new EntityCollection();
            EntityCollection entToList = new EntityCollection();
            EntityReferenceCollection materialGroup = new EntityReferenceCollection();

            EntityCollection toTeamMembers = new EntityCollection();
            EntityCollection ccTeamMembers = new EntityCollection();
            EntityReference department = null;

            QueryExpression _querybm = new QueryExpression("hil_userbranchmapping");
            _querybm.ColumnSet = new ColumnSet("hil_name", "hil_department", "hil_zonalhead", "hil_user", "hil_scm", "hil_salesoffice", "hil_buhead", "hil_branchproducthead", "hil_zonerepresentor", "hil_zdc");
            _querybm.Criteria = new FilterExpression(LogicalOperator.And);
            _querybm.Criteria.AddCondition(new ConditionExpression("hil_user", ConditionOperator.Equal, toownerid.Id));
            _querybm.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            EntityCollection userMapingColl = service.RetrieveMultiple(_querybm);

            department = userMapingColl.Entities[0].GetAttributeValue<EntityReference>("hil_department");


            string[] toTeam = to.Split(';');
            string[] ccTeam = cc.Split(';');

            tracingService.Trace("toTeam count" + toTeam.Length);

            foreach (string totype in toTeam)
            {
                if (totype.Contains("t.") || totype.Contains("T."))
                {
                    string teamName = (totype.Replace("t.", "")).Replace("T.", "");
                    tracingService.Trace("team name" + teamName);
                    toTeamMembers = retriveTeamMembers(service, teamName, department, toTeamMembers);
                }
                else if (totype.Contains("p.") || totype.Contains("P."))
                {
                    string position = (totype.Replace("p.", "")).Replace("P.", "");
                    tracingService.Trace("position name" + position);
                    EntityReference positionRef = null;
                    if (position.ToLower() == "Zonal Head".ToLower())
                    {
                        positionRef = userMapingColl.Entities[0].GetAttributeValue<EntityReference>("hil_zonalhead");
                    }
                    else if (position.ToLower() == "scm".ToLower())
                    {
                        positionRef = userMapingColl.Entities[0].GetAttributeValue<EntityReference>("hil_scm");
                    }
                    else if (position.ToLower() == "BU Head".ToLower())
                    {
                        positionRef = userMapingColl.Entities[0].GetAttributeValue<EntityReference>("hil_buhead");
                    }
                    else if (position.ToLower() == "Branch Product Head".ToLower())
                    {
                        positionRef = userMapingColl.Entities[0].GetAttributeValue<EntityReference>("hil_branchproducthead");
                    }
                    else if (position.ToLower() == "Enquiry Creator".ToLower())
                    {
                        positionRef = userMapingColl.Entities[0].GetAttributeValue<EntityReference>("hil_user");
                    }
                    else if (position.ToLower() == "zonal representor".ToLower())
                    {
                        positionRef = userMapingColl.Entities[0].GetAttributeValue<EntityReference>("hil_zonerepresentor");
                    }
                    else if (position.ToLower() == "Design Team".ToLower())
                    {
                        EntityReference salsesOffice = userMapingColl.Entities[0].GetAttributeValue<EntityReference>("hil_salesoffice");
                        QueryExpression query = new QueryExpression("hil_designteambranchmapping");
                        query.ColumnSet = new ColumnSet("hil_user");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition(new ConditionExpression("hil_salesoffice", ConditionOperator.Equal, salsesOffice.Id));
                        // query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 0));
                        EntityCollection design = service.RetrieveMultiple(query);
                        if (design.Entities.Count > 0)
                            positionRef = design[0].GetAttributeValue<EntityReference>("hil_user");
                    }
                    Entity entity = service.Retrieve(positionRef.LogicalName, positionRef.Id, new ColumnSet(false));
                    toTeamMembers.Entities.Add(entity);
                }
            }
            tracingService.Trace("ccTeam count" + ccTeam.Length);
            foreach (string totype in ccTeam)
            {
                if (totype.Contains("t.") || totype.Contains("T."))
                {
                    string teamName = (totype.Replace("t.", "")).Replace("T.", "");
                    tracingService.Trace("team name" + teamName);
                    ccTeamMembers = retriveTeamMembers(service, teamName, department, ccTeamMembers);
                }
                else if (totype.Contains("p.") || totype.Contains("P."))
                {
                    string position = (totype.Replace("p.", "")).Replace("P.", "");
                    tracingService.Trace("position name" + position);
                    EntityReference positionRef = null;
                    if (position.ToLower() == "Zonal Head".ToLower())
                    {
                        positionRef = userMapingColl.Entities[0].GetAttributeValue<EntityReference>("hil_zonalhead");
                    }
                    else if (position.ToLower() == "scm".ToLower())
                    {
                        positionRef = userMapingColl.Entities[0].GetAttributeValue<EntityReference>("hil_scm");
                    }
                    else if (position.ToLower() == "BU Head".ToLower())
                    {
                        positionRef = userMapingColl.Entities[0].GetAttributeValue<EntityReference>("hil_buhead");
                    }
                    else if (position.ToLower() == "Branch Product Head".ToLower())
                    {
                        positionRef = userMapingColl.Entities[0].GetAttributeValue<EntityReference>("hil_branchproducthead");
                    }
                    else if (position.ToLower() == "Enquiry Creator".ToLower())
                    {
                        positionRef = userMapingColl.Entities[0].GetAttributeValue<EntityReference>("hil_user");
                    }
                    else if (position.ToLower() == "zonal representor".ToLower())
                    {
                        positionRef = userMapingColl.Entities[0].GetAttributeValue<EntityReference>("hil_zonerepresentor");
                    }
                    else if (position.ToLower() == "Design Team".ToLower())
                    {
                        EntityReference salsesOffice = userMapingColl.Entities[0].GetAttributeValue<EntityReference>("hil_salesoffice");
                        QueryExpression query = new QueryExpression("hil_designteambranchmapping");
                        query.ColumnSet = new ColumnSet("hil_user");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition(new ConditionExpression("hil_salesoffice", ConditionOperator.Equal, salsesOffice.Id));
                        EntityCollection design = service.RetrieveMultiple(query);
                        if (design.Entities.Count > 0)
                            positionRef = design[0].GetAttributeValue<EntityReference>("hil_user");
                    }

                    Entity entity = service.Retrieve(positionRef.LogicalName, positionRef.Id, new ColumnSet(false));
                    ccTeamMembers.Entities.Add(entity);
                }
            }
            String URL = @"https://havells.crm8.dynamics.com/main.aspx?appid=675ffa44-b6d0-ea11-a813-000d3af05d7b&forceUCI=1&pagetype=entityrecord&etn=";
            string recordURL = URL + tender.LogicalName + "&id=" + tender.Id;
            if (mailBody.Contains("#URL"))
            {
                mailBody = mailBody.Replace("#URL", recordURL);
            }
            tracingService.Trace("toTeamMembers count" + toTeamMembers.Entities.Count);

            foreach (Entity ccEntity in toTeamMembers.Entities)
            {
                Entity entCC = new Entity("activityparty");
                entCC["partyid"] = ccEntity.ToEntityReference();
                entToList.Entities.Add(entCC);
            }
            tracingService.Trace("entCCList count" + entCCList.Entities.Count);
            foreach (Entity ccEntity in ccTeamMembers.Entities)
            {
                Entity entTo = new Entity("activityparty");
                entTo["partyid"] = ccEntity.ToEntityReference();
                entCCList.Entities.Add(entTo);
            }
            Entity entFrom = new Entity("activityparty");
            entFrom["partyid"] = new EntityReference("queue", new Guid("a81ac51b-c9df-eb11-bacb-0022486ea537"));
            Entity[] entFromList = { entFrom };

            tracingService.Trace("1");

            Entity email = new Entity("email");
            email["from"] = entFromList;
            email["to"] = entToList;
            email["cc"] = entCCList;
            email["description"] = mailBody;
            email["subject"] = mailsubject;
            email["regardingobjectid"] = tender.ToEntityReference();
            Guid emailId = service.Create(email);

            SendEmailRequest sendEmailReq = new SendEmailRequest()
            {
                EmailId = emailId,
                IssueSend = true
            };
            SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);

        }
    }
}
