using System;
using HavellsNewPlugin.Approval;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.Escalations
{
  public  class UpdateStatus : IPlugin
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
            tracingService.Trace("Depth " + context.Depth);
            EntityReference regardingRef = null;
            tracingService.Trace("Message :- " + context.MessageName);
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    tracingService.Trace("1");
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("2");
                    entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));
                    int SLAStatus = entity.GetAttributeValue<OptionSetValue>("status").Value;
                    regardingRef = entity.GetAttributeValue<EntityReference>("regarding");
                    if (regardingRef.LogicalName== "hil_escalation")
                    {
                        Entity esalation = new Entity(regardingRef.LogicalName, regardingRef.Id);
                        esalation["hil_slastatus"] = entity["status"];
                        service.Update(esalation);
                        if (SLAStatus == 1)
                        {
                            sendEmail(service, regardingRef);
                            createNewEscalation(service, regardingRef);
                        }
                    }
                    

                    tracingService.Trace("CreateTask plugin started  ");
                }
            }
            catch (Exception ex)
            {
                Entity esalation = new Entity(regardingRef.LogicalName, regardingRef.Id);
                esalation["hil_workflowfailed"] = true;
                esalation["hil_exception"] = ex.Message.ToString();
                service.Update(esalation);
                //throw new InvalidPluginExecutionException("Error : " + ex.Message);
            }
        }
        void sendEmail(IOrganizationService service, EntityReference escalationRef)
        {
            Entity escalation = service.Retrieve(escalationRef.LogicalName, escalationRef.Id, new ColumnSet(true));
            tracingService.Trace("2");

            EntityReference taskRef = escalation.GetAttributeValue<EntityReference>("hil_task");
            Entity task = service.Retrieve("task", taskRef.Id, new ColumnSet("regardingobjectid", "subject"));
            Entity tender = service.Retrieve(task.GetAttributeValue<EntityReference>("regardingobjectid").LogicalName, 
                task.GetAttributeValue<EntityReference>("regardingobjectid").Id, new ColumnSet(true));
            QueryExpression query2 = new QueryExpression("hil_escalationmatrix");
            query2.ColumnSet = new ColumnSet("hil_name", "hil_tomailpoition", "hil_copytoposition", "hil_purpose", "hil_entityname", "hil_subjectline", "hil_mailbody");
            query2.Criteria = new FilterExpression(LogicalOperator.And);
            query2.Criteria.AddCondition("hil_purpose", ConditionOperator.Equal, task.GetAttributeValue<string>("subject"));
            query2.Criteria.AddCondition("hil_entityname", ConditionOperator.Equal, tender.LogicalName);
            query2.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            EntityCollection entcol2 = service.RetrieveMultiple(query2);
            if (entcol2.Entities.Count > 0)
            {
                tracingService.Trace("3");
                string _primaryField = string.Empty;
                ApprovalHelper.GetPrimaryIdFieldName(tender.LogicalName, service, out _primaryField);
                string subjectLine = entcol2.Entities[0].GetAttributeValue<string>("hil_subjectline");
                string mailBody = entcol2.Entities[0].GetAttributeValue<string>("hil_mailbody");
                ColumnSet collSet = ApprovalHelper.findEntityColl(mailBody, _primaryField, service, tracingService);
                string Mailsubject = null;
                if (subjectLine.Contains("{"))
                {
                    Mailsubject = ApprovalHelper.createEmailSubject(tender, _primaryField, subjectLine, service, tracingService);
                }
                else
                {
                    Mailsubject = subjectLine;
                }
                string mailbodyText = ApprovalHelper.createEmailBody(tender, mailBody, "", collSet, service, tracingService);

                EntityReference _regarding = task.GetAttributeValue<EntityReference>("regardingobjectid");
                tracingService.Trace("4");
                EntityReference tomailposition = tender.GetAttributeValue<EntityReference>("ownerid");
                EntityReference targetOwner = tender.GetAttributeValue<EntityReference>("ownerid");
                EntityCollection entToList = new EntityCollection();
                entToList.EntityName = "systemuser";
                string copyto = string.Empty;
                if (entcol2[0].Contains("hil_copytoposition"))
                {
                    copyto = entcol2[0].GetAttributeValue<string>("hil_copytoposition");
                    QueryExpression _query = new QueryExpression("hil_userbranchmapping");
                    _query.ColumnSet = new ColumnSet("hil_name", "hil_zonalhead", "hil_user", "hil_salesoffice", "hil_buhead", "hil_branchproducthead");
                    _query.Criteria = new FilterExpression(LogicalOperator.And);
                    _query.Criteria.AddCondition(new ConditionExpression("hil_user", ConditionOperator.Equal, targetOwner.Id));
                    _query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                    EntityCollection userMapingColl = service.RetrieveMultiple(_query);
                    if (copyto != string.Empty)
                        entToList = ApprovalHelper.getCopyToData(copyto, userMapingColl, service);

                }
                tracingService.Trace("5");
                ApprovalHelper.sendEmal(tomailposition, entToList, _regarding, mailbodyText, Mailsubject, service);
            }
        }
        static public void createNewEscalation(IOrganizationService service, EntityReference escalationRef)
        {
            tracingService.Trace("createNewEscalation");
            Entity escalation = service.Retrieve(escalationRef.LogicalName, escalationRef.Id, new ColumnSet(true));
            Entity escalationEntity = new Entity(escalation.LogicalName);
            escalationEntity["hil_name"] = escalation["hil_name"];
            escalationEntity["hil_task"] = escalation["hil_task"];
            escalationEntity["slaid"] = escalation["slaid"];
            escalationEntity["hil_escalationmatrix"] = escalation["hil_escalationmatrix"];
            escalationEntity["hil_slastatus"] = new OptionSetValue(0);
            Guid escID = service.Create(escalationEntity);
            tracingService.Trace("created New Escalation with id = " + escID);
        }
    }
}
