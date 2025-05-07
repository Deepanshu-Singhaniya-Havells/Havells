using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using Microsoft.Crm.Sdk.Messages;
using System.Collections.Generic;
using System.Web.UI.HtmlControls;

namespace HavellsNewPlugin.CustomWorkflow
{
    public class ClearForDispatch : CodeActivity
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
        [ReferenceTarget("hil_oaheader")]
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
                EntityReference ocl = null;
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
                    new ColumnSet("hil_department", "ownerid", "hil_orderchecklistid", "hil_zdc", "hil_scm", "hil_zonerepresentor", "hil_salesoffice"));
                if (_regardingEntity.Contains("hil_department"))
                {
                    department = _regardingEntity.GetAttributeValue<EntityReference>("hil_department");
                }
                if (_regardingEntity.Contains("hil_orderchecklistid"))
                {
                    ocl = _regardingEntity.GetAttributeValue<EntityReference>("hil_orderchecklistid");
                }

                Entity _oclEntity = service.Retrieve(ocl.LogicalName, ocl.Id, new ColumnSet("hil_department", "hil_despatchpoint", "hil_buhead", "ownerid", "hil_rm", "hil_zonalhead", "hil_zdc", "hil_scm", "hil_zonerepresentor", "hil_salesoffice"));

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
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_zonalhead");
                        }
                        else if (position.ToLower() == "scm".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_scm");
                        }
                        else if (position.ToLower() == "BU Head".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_buhead");
                        }
                        else if (position.ToLower() == "Branch Product Head".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_rm");
                        }
                        else if (position.ToLower() == "Enquiry Creator".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("ownerid");
                        }
                        else if (position.ToLower() == "zonal representor".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_zonerepresentor");
                        }
                        else if (position.ToLower() == "Design Team".ToLower())
                        {
                            EntityReference salsesOffice = _oclEntity.GetAttributeValue<EntityReference>("hil_salesoffice");
                            QueryExpression query = new QueryExpression("hil_designteambranchmapping");
                            query.ColumnSet = new ColumnSet("hil_user");
                            query.Criteria = new FilterExpression(LogicalOperator.And);
                            query.Criteria.AddCondition(new ConditionExpression("hil_salesoffice", ConditionOperator.Equal, salsesOffice.Id));
                            query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 0));
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
                        ccTeamMembers = retriveTeamMembers(service, teamName, department, toTeamMembers);
                    }
                    else if (totype.Contains("p.") || totype.Contains("P."))
                    {
                        string position = (totype.Replace("p.", "")).Replace("P.", "");
                        tracingService.Trace("position name" + position);
                        EntityReference positionRef = null;
                        if (position.ToLower() == "Zonal Head".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_zonalhead");
                        }
                        else if (position.ToLower() == "scm".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_scm");
                        }
                        else if (position.ToLower() == "BU Head".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_buhead");
                        }
                        else if (position.ToLower() == "Branch Product Head".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_rm");
                        }
                        else if (position.ToLower() == "Enquiry Creator".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("ownerid");
                        }
                        else if (position.ToLower() == "zonal representor".ToLower())
                        {
                            positionRef = _oclEntity.GetAttributeValue<EntityReference>("hil_zonerepresentor");
                        }
                        else if (position.ToLower() == "Design Team".ToLower())
                        {
                            EntityReference salsesOffice = _oclEntity.GetAttributeValue<EntityReference>("hil_salesoffice");
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

                QueryExpression tenderURL = new QueryExpression("hil_integrationconfiguration");
                tenderURL.ColumnSet = new ColumnSet("hil_url");
                tenderURL.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "TenderAppURL");
                EntityCollection tenderURLCol = service.RetrieveMultiple(tenderURL);
                string URL = tenderURLCol.Entities[0].GetAttributeValue<string>("hil_url");

                string recordURL = URL + _regardingEntity.LogicalName + "&id=" + _regardingEntity.Id;
                tracingService.Trace("recordURL " + recordURL);
                string clickhereContent = "For more details please &nbsp; <a target='_blank' href=" + recordURL + "> Click Here </a>";
                if (mailBody.Contains("#Clickhere"))
                {
                    tracingService.Trace("_regardingEntity.Id " + _regardingEntity.Id);
                    QueryExpression query = new QueryExpression("hil_attachment");
                    query.ColumnSet = new ColumnSet(true);
                    query.Distinct = true;
                    query.Criteria.AddCondition("hil_isdeleted", ConditionOperator.Equal, false);
                    query.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, _regardingEntity.Id);
                    EntityCollection oaheaderCol = service.RetrieveMultiple(query);
                    tracingService.Trace("oaheaderCol.Entities.Count  " + oaheaderCol.Entities.Count);
                    if (oaheaderCol.Entities.Count > 0)
                    {
                        mailBody = mailBody.Replace("#Clickhere", clickhereContent);
                    }
                    else
                    {
                        mailBody = mailBody.Replace("#Clickhere", "");
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

                tracingService.Trace("1");

                Entity email = new Entity("email");
                email["from"] = entFromList;
                email["to"] = entToList;
                email["cc"] = entCCList;
                email["description"] = mailBody;
                email["subject"] = mailsubject;
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
                throw new InvalidPluginExecutionException("Error :- " + ex.Message);
            }
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
                        tracingService.Trace("bdteamCol count " + bdteamCol.Entities.Count);
                        tracingService.Trace("bdteamCol.Entities[0].Id " + bdteamCol.Entities[0].Id);
                        QueryExpression _querymem = new QueryExpression("hil_bdteammember");
                        _querymem.ColumnSet = new ColumnSet("emailaddress");
                        _querymem.Criteria = new FilterExpression(LogicalOperator.And);
                        _querymem.Criteria.AddCondition(new ConditionExpression("hil_team", ConditionOperator.Equal, bdteamCol.Entities[0].Id));
                        _querymem.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                        EntityCollection bdteammemCol = service.RetrieveMultiple(_querymem);
                        EntityCollection entTOList = new EntityCollection();
                        tracingService.Trace("Team Members count" + entTOList.Entities.Count);
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