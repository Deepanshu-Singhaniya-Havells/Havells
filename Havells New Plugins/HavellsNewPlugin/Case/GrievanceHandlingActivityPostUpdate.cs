using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.AccessControl;

namespace HavellsNewPlugin.Case
{
    public class GrievanceHandlingActivityPostUpdate : IPlugin
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
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                        && context.PrimaryEntityName.ToLower() == "hil_grievancehandlingactivity" && context.Depth == 1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if (entity.Contains("statecode"))
                    {
                        if (entity.GetAttributeValue<OptionSetValue>("statecode").Value == 1)//Completed
                        {
                            //Entity entGHA = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("actualstart"));
                            //DateTime _actualStart = entGHA.GetAttributeValue<DateTime>("actualstart").AddMinutes(330);
                            //DateTime _actualEnd = DateTime.Now.AddMinutes(330);
                            //TimeSpan ts = _actualEnd - _actualStart;
                            //Entity entUpdateGHA = new Entity("hil_grievancehandlingactivity", entity.Id);
                            //entUpdateGHA["actualend"] = _actualEnd;
                            //entUpdateGHA["actualdurationminutes"] = ts.TotalMinutes;
                            //service.Update(entUpdateGHA);

                            Entity entGHA = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("actualstart", "regardingobjectid", "hil_caseassignmentmatrixlineid"));
                            EntityReference caseRef = entGHA.Contains("regardingobjectid") ? entGHA.GetAttributeValue<EntityReference>("regardingobjectid") : null;
                            EntityReference assignmentMatrixLineRef = entGHA.GetAttributeValue<EntityReference>("hil_caseassignmentmatrixlineid");
                            Entity assignmentMatrixLine = service.Retrieve(assignmentMatrixLineRef.LogicalName, assignmentMatrixLineRef.Id, new ColumnSet("hil_level"));
                            int level = assignmentMatrixLine.Contains("hil_level") ? assignmentMatrixLine.GetAttributeValue<int>("hil_level") : -1;

                            if (caseRef != null && level != -1) UpdatePendingStatus(caseRef, level,service);

                            SendEmailOnActivityComplete(entity, service);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error : " + ex.Message);
            }
        }
        private void SendEmailOnActivityComplete(Entity CaseGrievanceAct,IOrganizationService service)
        {
            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='hil_grievancehandlingactivity'>
                    <attribute name='regardingobjectid' />
                    <attribute name='ownerid' />
                    <order attribute='subject' descending='false' />
                    <filter type='and'>
                      <condition attribute='activityid' operator='eq' value='{CaseGrievanceAct.Id}' />
                    </filter>
                    <link-entity name='hil_caseassignmentmatrixline' from='hil_caseassignmentmatrixlineid' to='hil_caseassignmentmatrixlineid' visible='false' link-type='outer' alias='amt'>
                      <attribute name='hil_level' />
                      <attribute name='hil_caseassignmentmatrixid' />
                    </link-entity>
                    <link-entity name='incident' from='incidentid' to='regardingobjectid' visible='false' link-type='outer' alias='case'>
                      <attribute name='title' />
                      <attribute name='ticketnumber' />
                    </link-entity>
                  </entity>
                </fetch>";
            EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
            if (_entCol.Entities.Count > 0) {
                int Level = Convert.ToInt32(_entCol.Entities[0].GetAttributeValue<AliasedValue>("amt.hil_level").Value);
                Guid CaseAssignmentMatrixId = ((EntityReference)_entCol.Entities[0].GetAttributeValue<AliasedValue>("amt.hil_caseassignmentmatrixid").Value).Id;
                EntityReference _entCase = _entCol.Entities[0].GetAttributeValue<EntityReference>("regardingobjectid");
                EntityReference _entActivityOwner = _entCol.Entities[0].GetAttributeValue<EntityReference>("ownerid");
                string CaseNumber = _entCol.Entities[0].Contains("case.ticketnumber") ? _entCol.Entities[0].GetAttributeValue<AliasedValue>("case.ticketnumber").Value.ToString() : null;
                string caseTitle = _entCol.Entities[0].Contains("case.title") ? _entCol.Entities[0].GetAttributeValue<AliasedValue>("case.title").Value.ToString() : null;
                string recordURL = $"https://havells.crm8.dynamics.com/main.aspx?appid=668eb624-0610-e911-a94e-000d3af06a98&forceUCI=1&pagetype=entityrecord&etn=incident&id={_entCase.Id}";

                QueryExpression query = new QueryExpression("hil_caseassignmentmatrixline");
                query.ColumnSet = new ColumnSet("hil_assigneeuser");
                query.Criteria.Filters.Add(new FilterExpression(LogicalOperator.And));
                query.Criteria.AddCondition("hil_level", ConditionOperator.Equal, (Level + 1));
                query.Criteria.AddCondition("hil_caseassignmentmatrixid", ConditionOperator.Equal, CaseAssignmentMatrixId);

                EntityCollection coll = service.RetrieveMultiple(query);
                if (coll.Entities.Count > 0)
                {
                    EntityReference EscUser = coll.Entities[0].GetAttributeValue<EntityReference>("hil_assigneeuser");

                    Entity fromActivityParty = new Entity("activityparty");
                    Entity toActivityParty = new Entity("activityparty");

                    fromActivityParty["partyid"] = new EntityReference("queue", new Guid("9b0480a8-e30f-e911-a94e-000d3af06a98"));
                    toActivityParty["partyid"] = EscUser;

                    Entity email = new Entity("email");
                    email["from"] = new Entity[] { fromActivityParty };
                    email["to"] = new Entity[] { toActivityParty };
                    email["regardingobjectid"] = _entCase;
                    email["subject"] = $"Activity mark completed for case nubmer {CaseNumber}";
                    email["description"] = $"Dear Team, <br/><br/>Activity mark completed by {_entActivityOwner.Name} for the case with case-ID {CaseNumber} regarding {caseTitle}. <br/><br/> To open the case <a href={recordURL}>Click Here</a> <br/> <br/> Regards Team CRM";
                    email["directioncode"] = true;
                    Guid emailId = service.Create(email);
                    //Use the SendEmail message to send an e-mail message.
                    SendEmailRequest sendEmailRequest = new SendEmailRequest
                    {
                        EmailId = emailId,
                        TrackingToken = "",
                        IssueSend = true
                    };
                    SendEmailResponse sendEmailresp = (SendEmailResponse)service.Execute(sendEmailRequest);
                }
            }
        }

        private void UpdatePendingStatus(EntityReference caseRef, int level, IOrganizationService service)
        {
            QueryExpression query = new QueryExpression("hil_grievancehandlingactivity");
            query.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, caseRef.Id);
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); // Open
            query.Criteria.AddCondition("hil_activitytype", ConditionOperator.NotEqual, 4); // Activity not assigned to SPOC 
            EntityCollection pendingAcivitiesColl = service.RetrieveMultiple(query);
            if (pendingAcivitiesColl.Entities.Count == 0)
            {
                query = new QueryExpression("hil_grievancehandlingactivity");
                query.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, caseRef.Id);
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); // Open
                query.Criteria.AddCondition("hil_activitytype", ConditionOperator.Equal, 4); //Activity not assigned to SPOC 
                pendingAcivitiesColl = service.RetrieveMultiple(query);
                Entity caseUpdate = new Entity(caseRef.LogicalName, caseRef.Id);
                if (pendingAcivitiesColl.Entities.Count > 0)
                {
                    caseUpdate["hil_pendinglevel"] = new OptionSetValue(7);//"Pending for closure by SPOC"; 
                }
                else
                {
                    caseUpdate["hil_pendinglevel"] = new OptionSetValue(8);//"Pending for closure by Assginee";   
                }
                service.Update(caseUpdate);
            }
            else
            {
                Entity caseUpdate = new Entity(caseRef.LogicalName, caseRef.Id);
                int pendingLevel = 6;
                if (level == 1)
                {
                    pendingLevel = 4; // Pendint at level 2
                }
                else if (level == 2)
                {
                    pendingLevel = 5; // Pending at level 3; 
                }
                caseUpdate["hil_pendinglevel"] = new OptionSetValue(pendingLevel);
                service.Update(caseUpdate);
            }

        }
    }
}
