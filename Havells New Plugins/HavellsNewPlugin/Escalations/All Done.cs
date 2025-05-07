using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.Escalations
{
    public class All_Done : IPlugin
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

                updateStatus(service, entityIds, entityName);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("CreateExclationTask Error " + ex.Message);
            }
        }
        void updateStatus(IOrganizationService service, string Guids, string entityName)
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

                QueryExpression query = new QueryExpression("task");
                query.ColumnSet = new ColumnSet(true);
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, tenderentity.Entities[0].Id);
                query.Criteria.AddCondition("actualend", ConditionOperator.Null);
                query.AddOrder("createdon", OrderType.Descending);

                EntityCollection enttask = service.RetrieveMultiple(query);

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
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error in createTask: " + ex.Message);
            }//Entity tenderentity = service.Retrieve(entityName, new Guid(Guids), new ColumnSet(true));

        }
    }
}
