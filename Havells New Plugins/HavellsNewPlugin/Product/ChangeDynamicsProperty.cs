using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace HavellsNewPlugin.Product
{
    public class ChangeDynamicsProperty : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            Entity entity = (Entity)context.InputParameters["Target"];
            if (entity.GetAttributeValue<bool>("hil_reviseproperties"))
            {
                QueryExpression Query = new QueryExpression("dynamicproperty");
                Query.ColumnSet = new ColumnSet(false);
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, entity.Id);
                EntityCollection Found = service.RetrieveMultiple(Query);
                foreach (Entity entity1 in Found.Entities)
                {
                    SetStateRequest req = new SetStateRequest();
                    req.State = new OptionSetValue(1);
                    req.Status = new OptionSetValue(0);
                    req.EntityMoniker = entity1.ToEntityReference();// new EntityReference("dynamicproperty", new Guid("f4fd7c49-6f9c-ed11-aad1-6045bdac5a1d"));
                    var res = (SetStateResponse)service.Execute(req);
                }
                Entity entity2 = new Entity(entity.LogicalName, entity.Id);
                entity2["hil_reviseproperties"] = false;
                service.Update(entity2);
            }
        }
    }
}
