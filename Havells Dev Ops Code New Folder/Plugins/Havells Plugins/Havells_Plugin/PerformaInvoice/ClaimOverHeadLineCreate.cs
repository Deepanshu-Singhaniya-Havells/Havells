using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.PerformaInvoice
{
    public class ClaimOverHeadLineCreate : IPlugin
    {
        public static ITracingService tracingService = null;
        Guid PerformaInvoiceID = Guid.Empty;
        public void Execute(IServiceProvider serviceProvider)
        {

            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                tracingService.Trace("1");
                Entity entity = (Entity)context.InputParameters["Target"];
                tracingService.Trace("2");
                entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_performainvoice"));
                if (entity.Contains("hil_performainvoice"))
                {
                    PerformaInvoiceID = ((EntityReference)entity["hil_performainvoice"]).Id;
                    tracingService.Trace("_claimOverHeadLineColl");
                    if (!entity.Contains("hil_callsubtype"))
                    {
                        QueryExpression query = new QueryExpression("hil_callsubtype");
                        query.ColumnSet = new ColumnSet(false);
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, "Breakdown"));
                        EntityCollection deliveryCol = service.RetrieveMultiple(query);
                        if (deliveryCol.Entities.Count > 0)
                            entity["hil_callsubtype"] = deliveryCol[0].ToEntityReference();
                    }
                }
            }
        }
    }
}
