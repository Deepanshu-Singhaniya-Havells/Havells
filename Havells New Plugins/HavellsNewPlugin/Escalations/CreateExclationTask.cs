using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.Escalations
{
    public class CreateExclationTask : IPlugin
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
                tracingService.Trace("CreateTask plugin started  ");
                var entityIds = context.InputParameters["EntityId"].ToString();
                var entityName = context.InputParameters["EntityName"].ToString();
                var purpose = context.InputParameters["Purpose"].ToString();
                createTask(service, entityIds, entityName, purpose);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("CreateExclationTask Error " + ex.Message);
            }
        }
        void createTask(IOrganizationService service, string Guids, string entityName, string purpose)
        {
            try
            {
                QueryExpression tendquery = new QueryExpression(entityName);
                tendquery.ColumnSet = new ColumnSet(true);
                tendquery.Criteria = new FilterExpression(LogicalOperator.And);
                tendquery.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, Guids));
                EntityCollection tenderentity = service.RetrieveMultiple(tendquery);
                if (tenderentity.Entities.Count < 0)
                {
                    throw new InvalidPluginExecutionException("********************test");
                }
                tracingService.Trace("11");

                QueryExpression _query = new QueryExpression("hil_escalationmatrix");
                _query.ColumnSet = new ColumnSet(true);
                _query.Criteria = new FilterExpression(LogicalOperator.And);
                _query.Criteria.AddCondition(new ConditionExpression("hil_entityname", ConditionOperator.Equal, entityName));
                _query.Criteria.AddCondition(new ConditionExpression("hil_purpose", ConditionOperator.Equal, purpose));
                EntityCollection escalationMatrixCheck = service.RetrieveMultiple(_query);
                if (escalationMatrixCheck.Entities.Count > 0)
                {
                    tracingService.Trace("2");
                    Guid taskid = Guid.Empty;
                    QueryExpression query = new QueryExpression("task");
                    query.ColumnSet = new ColumnSet(true);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, tenderentity.Entities[0].Id);
                    query.Criteria.AddCondition("actualend", ConditionOperator.Null);
                    query.AddOrder("createdon", OrderType.Descending);

                    EntityCollection enttask = service.RetrieveMultiple(query);
                    tracingService.Trace("22");
                    if (enttask.Entities.Count > 0)
                    {
                        tracingService.Trace("22TaskUpdate");
                        foreach (Entity tas in enttask.Entities)
                        {
                            Entity ttask = new Entity("task");
                            ttask.Id = tas.Id;
                            ttask["actualend"] = DateTime.Now;
                            service.Update(ttask);
                            query = new QueryExpression("hil_escalation");
                            query.ColumnSet = new ColumnSet(false);
                            query.Criteria = new FilterExpression(LogicalOperator.And);
                            query.Criteria.AddCondition("hil_task", ConditionOperator.Equal, tas.Id);
                            query.AddOrder("createdon", OrderType.Descending);
                            EntityCollection entEsc = service.RetrieveMultiple(query);
                            if (entEsc.Entities.Count > 0)
                            {
                                Entity UpEsc = new Entity(entEsc[0].LogicalName);
                                UpEsc.Id = entEsc[0].Id;
                                UpEsc["hil_issuccess"] = true;
                                service.Update(UpEsc);
                            }
                        }
                        tracingService.Trace("3");
                    }
                    Entity task = new Entity("task");
                    task["subject"] = escalationMatrixCheck.Entities[0].GetAttributeValue<string>("hil_purpose");
                    task["hil_tasktype"] = new OptionSetValue(2); // 2 for esclation
                    tracingService.Trace("222");
                    task["regardingobjectid"] = new EntityReference(tenderentity.Entities[0].LogicalName, tenderentity.Entities[0].Id);
                    task["actualstart"] = DateTime.Now;
                    string toPosition = escalationMatrixCheck.Entities[0].GetAttributeValue<string>("hil_tomailpoition");
                    EntityReference escalToPosition = new EntityReference();
                    if (entityName == "hil_tender")
                    {
                        tracingService.Trace("4");
                        escalToPosition = tenderentity.Entities[0].GetAttributeValue<EntityReference>("ownerid");
                    }
                    task["owninguser"] = new EntityReference(tenderentity.Entities[0].LogicalName, escalToPosition.Id);
                    taskid = service.Create(task);

                    //create esclation
                    Entity esclation = new Entity("hil_escalation");
                    esclation["hil_escalationmatrix"] = new EntityReference("hil_escalationmatrix", escalationMatrixCheck.Entities[0].Id);
                    esclation["hil_name"] = escalationMatrixCheck.Entities[0].GetAttributeValue<string>("hil_purpose");
                    esclation["hil_task"] = new EntityReference(task.LogicalName, taskid);
                    esclation["hil_slastatus"] = new OptionSetValue(0);
                    esclation["slaid"] = new EntityReference("sla", escalationMatrixCheck.Entities[0].GetAttributeValue<EntityReference>("hil_sla").Id);
                    service.Create(esclation);
                    tracingService.Trace("Complete");

                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error in createTask: " + ex.Message);
            }//Entity tenderentity = service.Retrieve(entityName, new Guid(Guids), new ColumnSet(true));

        }
    }
}
