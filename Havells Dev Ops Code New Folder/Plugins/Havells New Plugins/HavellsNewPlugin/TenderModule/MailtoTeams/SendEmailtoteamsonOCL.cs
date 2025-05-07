using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using Microsoft.Crm.Sdk.Messages;

namespace HavellsNewPlugin.TenderModule.MailtoTeams
{
    public class SendEmailtoteamsonOCL : CodeActivity
    {
        public static ITracingService tracingService = null;

        [RequiredArgument]
        [Input("MailBody")]

        public InArgument<String> mailBody { get; set; }

        [RequiredArgument]
        [Input("MailSubject")]

        public InArgument<String> mailsubject { get; set; }

        [RequiredArgument]
        [Input("To")]
        public InArgument<String> to { get; set; }

        [Input("CC")]
        public InArgument<String> cc { get; set; }

        [RequiredArgument]
        [Input("Regarding")]
        [ReferenceTarget("hil_orderchecklist")]
        public InArgument<EntityReference> regarding { get; set; }

        [RequiredArgument]
        [Input("From")]
        [ReferenceTarget("queue")]
        public InArgument<EntityReference> from { get; set; }
        protected override void Execute(CodeActivityContext executionContext)
        {
            try
            {
                EntityCollection entCCList = new EntityCollection();
                EntityCollection entToList = new EntityCollection();

                EntityCollection toTeamMembers = new EntityCollection();
                EntityCollection ccTeamMembers = new EntityCollection();
                EntityReference department = null;
                var context = executionContext.GetExtension<IWorkflowContext>();
                var serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
                tracingService = executionContext.GetExtension<ITracingService>();
                EntityReference fromRef = this.from.Get((ActivityContext)executionContext);
                EntityReference regardingRef = this.regarding.Get((ActivityContext)executionContext);
                string to = this.to.Get((ActivityContext)executionContext);
                string cc = this.cc.Get((ActivityContext)executionContext);
                string mailBody = this.mailBody.Get((ActivityContext)executionContext);
                string mailsubject = this.mailsubject.Get((ActivityContext)executionContext);


                Entity _regardingEntity = service.Retrieve(regardingRef.LogicalName, regardingRef.Id,
                    new ColumnSet("hil_department", "hil_buhead", "ownerid", "hil_rm", "hil_zonalhead", "hil_zdc", "hil_scm", "hil_zonerepresentor", "hil_salesoffice"));
                if (_regardingEntity.Contains("_regardingEntity"))
                {
                    department = _regardingEntity.GetAttributeValue<EntityReference>("hil_department");
                    tracingService.Trace("Department is " + department.Name);
                }



                string[] toTeam = to.Split(';');
                string[] ccTeam = cc.Split(';');

                tracingService.Trace("toTeam count" + toTeam.Length);

                foreach (string totype in toTeam)
                {
                    if (totype.Contains("t.") || totype.Contains("T."))
                    {
                        string teamName = (totype.Replace("t.", "")).Replace("T.", "");
                        tracingService.Trace("team name" + teamName);
                        toTeamMembers = retriveTeamMembers(service, teamName, null, department, null, toTeamMembers);

                    }
                    else if (totype.Contains("p.") || totype.Contains("P."))
                    {
                        string position = (totype.Replace("p.", "")).Replace("P.", "");
                        tracingService.Trace("position name" + position);
                        EntityReference positionRef = null;
                        if (position.ToLower() == "Zonal Head".ToLower())
                        {
                            positionRef = _regardingEntity.GetAttributeValue<EntityReference>("hil_zonalhead");
                        }
                        else if (position.ToLower() == "scm".ToLower())
                        {
                            positionRef = _regardingEntity.GetAttributeValue<EntityReference>("hil_scm");
                        }
                        else if (position.ToLower() == "BU Head".ToLower())
                        {
                            positionRef = _regardingEntity.GetAttributeValue<EntityReference>("hil_buhead");
                        }
                        else if (position.ToLower() == "Branch Product Head".ToLower())
                        {
                            positionRef = _regardingEntity.GetAttributeValue<EntityReference>("hil_rm");
                        }
                        else if (position.ToLower() == "Enquiry Creator".ToLower())
                        {
                            positionRef = _regardingEntity.GetAttributeValue<EntityReference>("ownerid");
                        }
                        else if (position.ToLower() == "zonal representor".ToLower())
                        {
                            positionRef = _regardingEntity.GetAttributeValue<EntityReference>("hil_zonerepresentor");
                        }
                        else if (position.ToLower() == "Design Team".ToLower())
                        {
                            EntityReference salsesOffice = _regardingEntity.GetAttributeValue<EntityReference>("hil_salesoffice");
                            QueryExpression query = new QueryExpression("hil_designteambranchmapping");
                            query.ColumnSet = new ColumnSet("hil_user");
                            query.Criteria = new FilterExpression(LogicalOperator.And);
                            query.Criteria.AddCondition(new ConditionExpression("hil_salesoffice", ConditionOperator.Equal, salsesOffice.Id));
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
                        ccTeamMembers = retriveTeamMembers(service, teamName, null, department, null, ccTeamMembers);
                    }
                    else if (totype.Contains("p.") || totype.Contains("P."))
                    {
                        string position = (totype.Replace("p.", "")).Replace("P.", "");
                        tracingService.Trace("position name" + position);
                        EntityReference positionRef = null;
                        if (position.ToLower() == "Zonal Head".ToLower())
                        {
                            positionRef = _regardingEntity.GetAttributeValue<EntityReference>("hil_zonalhead");
                        }
                        else if (position.ToLower() == "scm".ToLower())
                        {
                            positionRef = _regardingEntity.GetAttributeValue<EntityReference>("hil_scm");
                        }
                        else if (position.ToLower() == "BU Head".ToLower())
                        {
                            positionRef = _regardingEntity.GetAttributeValue<EntityReference>("hil_buhead");
                        }
                        else if (position.ToLower() == "Branch Product Head".ToLower())
                        {
                            positionRef = _regardingEntity.GetAttributeValue<EntityReference>("hil_rm");
                        }
                        else if (position.ToLower() == "Enquiry Creator".ToLower())
                        {
                            positionRef = _regardingEntity.GetAttributeValue<EntityReference>("ownerid");
                        }
                        else if (position.ToLower() == "zonal representor".ToLower())
                        {
                            positionRef = _regardingEntity.GetAttributeValue<EntityReference>("hil_zonerepresentor");
                        }
                        else if (position.ToLower() == "Design Team".ToLower())
                        {
                            EntityReference salsesOffice = _regardingEntity.GetAttributeValue<EntityReference>("hil_salesoffice");
                            QueryExpression query = new QueryExpression("hil_designteambranchmapping");
                            query.ColumnSet = new ColumnSet("hil_user");
                            query.Criteria = new FilterExpression(LogicalOperator.And);
                            query.Criteria.AddCondition(new ConditionExpression("hil_salesoffice", ConditionOperator.Equal, salsesOffice.Id));
                            EntityCollection design = service.RetrieveMultiple(query);
                            if (design.Entities.Count > 0)
                                positionRef = design[0].GetAttributeValue<EntityReference>("hil_user");
                        }
                        tracingService.Trace("User Name " + positionRef.Name);
                        Entity entity = service.Retrieve(positionRef.LogicalName, positionRef.Id, new ColumnSet(false));
                        ccTeamMembers.Entities.Add(entity);
                        tracingService.Trace("User Name Added");
                    }
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
                entFrom["partyid"] = new EntityReference(fromRef.LogicalName, fromRef.Id);
                Entity[] entFromList = { entFrom };
                String URL = @"https://havells.crm8.dynamics.com/main.aspx?appid=675ffa44-b6d0-ea11-a813-000d3af05d7b&forceUCI=1&pagetype=entityrecord&etn=";
                string recordURL = URL + _regardingEntity.LogicalName + "&id=" + _regardingEntity.Id;
                if (mailBody.Contains("#URL"))
                {
                    mailBody = mailBody.Replace("#URL", recordURL);
                }
                tracingService.Trace("1");

                Entity email = new Entity("email");
                email["from"] = entFromList;
                email["to"] = entToList;
                email["cc"] = entCCList;
                email["description"] = mailBody;
                email["subject"] = emailSubjectForOCL(service, regardingRef) + "-" + mailsubject;
                email["regardingobjectid"] = _regardingEntity.ToEntityReference();
                Guid emailId = service.Create(email);

                SendEmailRequest sendEmailReq = new SendEmailRequest()
                {
                    EmailId = emailId,
                    IssueSend = true
                };
                SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error :- " + ex.Message);// "HavellsNewPlugin.TenderModule.OrderCheckList.OrderCheckListPostCreate.Execute Error " + ex.Message);
            }
        }
        public EntityCollection retriveTeamMembers(IOrganizationService service, string _teamName, EntityReference _materialGroup, EntityReference _department,
            EntityReference _plant, EntityCollection extTeamMembers)
        {
            try
            {
                QueryExpression _query = new QueryExpression("hil_bdteam");
                _query.ColumnSet = new ColumnSet(false);
                _query.Criteria = new FilterExpression(LogicalOperator.And);
                _query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1));
                if (_teamName != null)
                    _query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, _teamName));
                if (_materialGroup != null)
                    _query.Criteria.AddCondition(new ConditionExpression("hil_materialgroup", ConditionOperator.Equal, _materialGroup.Id));
                if (_department != null)
                    _query.Criteria.AddCondition(new ConditionExpression("hil_department", ConditionOperator.Equal, _department.Id));
                if (_plant != null)
                    _query.Criteria.AddCondition(new ConditionExpression("hil_plant", ConditionOperator.Equal, _plant.Id));
                EntityCollection bdteamCol = service.RetrieveMultiple(_query);
                if (bdteamCol.Entities.Count > 0)
                {
                    tracingService.Trace("bdteamCol count " + bdteamCol.Entities.Count);
                    QueryExpression _querymem = new QueryExpression("hil_bdteammember");
                    _querymem.ColumnSet = new ColumnSet("emailaddress");
                    _querymem.Criteria = new FilterExpression(LogicalOperator.And);
                    _querymem.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1));
                    _querymem.Criteria.AddCondition(new ConditionExpression("hil_team", ConditionOperator.Equal, bdteamCol.Entities[0].Id));
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
            catch (Exception ex)
            {
                tracingService.Trace("Error in Retriving Team Members : " + ex.Message);
            }
            return extTeamMembers;
        }
        public string emailSubjectForOCL(IOrganizationService service, EntityReference regarding)
        {
            string _emailSubject = null;
            Entity tenderEntity = service.Retrieve(regarding.LogicalName, regarding.Id,
                new ColumnSet("hil_tenderno", "hil_projectname", "hil_name", "hil_nameofclientcustomercode"));
            if (tenderEntity.Contains("hil_tenderno") && tenderEntity["hil_tenderno"] != null)
            {
                _emailSubject = "Tender No " + tenderEntity.GetAttributeValue<EntityReference>("hil_tenderno").Name + " ";
            }
            if (tenderEntity.Contains("hil_name") && tenderEntity["hil_name"] != null)
            {
                _emailSubject = _emailSubject + tenderEntity.GetAttributeValue<string>("hil_name") + " ";
            }
            if (tenderEntity.Contains("hil_projectname") && tenderEntity["hil_projectname"] != null)
            {
                _emailSubject = _emailSubject + "Project " + tenderEntity.GetAttributeValue<string>("hil_projectname") + " ";
            }
            if (tenderEntity.Contains("hil_nameofclientcustomercode") && tenderEntity["hil_nameofclientcustomercode"] != null)
            {
                _emailSubject = _emailSubject + "customer " + tenderEntity.GetAttributeValue<EntityReference>("hil_nameofclientcustomercode").Name;
            }
            tracingService.Trace("Email full subject:" + _emailSubject);
            return _emailSubject;

        }
    }
}
