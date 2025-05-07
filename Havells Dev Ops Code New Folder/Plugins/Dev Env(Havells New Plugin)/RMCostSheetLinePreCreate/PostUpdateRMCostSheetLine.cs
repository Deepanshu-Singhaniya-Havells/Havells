using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.RMCostSheetLinePreCreate
{
    public class PostUpdateRMCostSheetLine : IPlugin
    {
        public static ITracingService tracingService = null;
        IPluginExecutionContext context;
        public void Execute(IServiceProvider serviceProvider)
        {
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "hil_rmcostsheetline" && context.Depth == 1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    Entity _entRMCostSheetLine = service.Retrieve("hil_rmcostsheetline", entity.Id, new ColumnSet("hil_rmcostsheet"));
                    tracingService.Trace("Step-1 " + entity.Id.ToString());
                    UpdatePackgExpensesonRMCostsheetHeader(service, _entRMCostSheetLine.GetAttributeValue<EntityReference>("hil_rmcostsheet").Id);
                    tracingService.Trace("step-1.1 " + _entRMCostSheetLine.GetAttributeValue<EntityReference>("hil_rmcostsheet").Id.ToString());
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("RMCostSheet.PostUpdateRMCostSheetLine.Execute: " + ex.Message);
            }
        }
        private void UpdatePackgExpensesonRMCostsheetHeader(IOrganizationService service, Guid _rmCostSheetId)
        {
            tracingService.Trace("step-1.2");
            string _fetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_rmcostsheetline'>
                                <attribute name='hil_rmcostsheetlineid'/>
                                <attribute name='hil_rmcode'/>
                                <attribute name='hil_rate'/>
                                <filter type='and'>
                                <condition attribute='hil_rmcode' operator='eq' uiname='Packing Charge' uitype='product' value='{{36F6268F-FE31-EF11-8E4F-7C1E520EB873}}'/>
                                <condition attribute='statecode' operator='eq' value='0'/>
                                <condition attribute='hil_rmcostsheet' operator='eq' value='{_rmCostSheetId}'/>
                                </filter>
                                </entity>
                                </fetch>";
            tracingService.Trace("step-1.3 " + _rmCostSheetId.ToString());
            EntityCollection ecoll = service.RetrieveMultiple(new FetchExpression(_fetch));
            if (ecoll.Entities.Count > 0)
            {
                tracingService.Trace("step-2");
                decimal _rmRate = ecoll.Entities[0].GetAttributeValue<Money>("hil_rate").Value;
                Entity CostSheet = new Entity("hil_rmcostsheet", _rmCostSheetId);
                CostSheet["hil_packagingexpenseper"] = _rmRate;
                service.Update(CostSheet);
            }
        }
    }
}
