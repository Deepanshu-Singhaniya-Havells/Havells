using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.RMCostSheet
{
    public class RMCostSheetPostUpdate : IPlugin
    {
        private ITracingService tracingService = null;
        private IOrganizationService service;
        private void UpdateCogsonTender(EntityReference tenderProductRef, decimal Cogs, OptionSetValue cogsstatus)
        {
            tracingService.Trace("in the tedner update method");
            Entity tenderProduct = service.Retrieve(tenderProductRef.LogicalName, tenderProductRef.Id, new ColumnSet("hil_cogs"));
            if (cogsstatus.Value == 3)
            {
                tenderProduct["hil_cogs"] = Cogs;
                service.Update(tenderProduct);
                tracingService.Trace("Updated cogs on tender");
            }
        }
        private void UpdateCOGS(Entity entity)
        {
            Entity rmCostSheet = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_rmcost", "hil_cogs", "hil_totalcost", "hil_tenderproductid", "hil_cogsstatus"));
            tracingService.Trace("Fetches Data");
            decimal totalCost = rmCostSheet.Contains("hil_totalcost") ? rmCostSheet.GetAttributeValue<Money>("hil_totalcost").Value : 0;
            OptionSetValue cogsstatus = rmCostSheet.Contains("hil_cogsstatus") ? rmCostSheet.GetAttributeValue<OptionSetValue>("hil_cogsstatus") : null;
            tracingService.Trace("Total Cost is present");
            tracingService.Trace($"RM Cost={rmCostSheet.GetAttributeValue<Money>("hil_rmcost").Value.ToString()} Total Exp={totalCost.ToString()}");
            if (rmCostSheet.Contains("hil_rmcost") && totalCost != 0)
            {
                tracingService.Trace($"2 RM Cost={rmCostSheet.GetAttributeValue<Money>("hil_rmcost").Value.ToString()} Total Exp={totalCost.ToString()}");
                decimal rmCost = rmCostSheet.GetAttributeValue<Money>("hil_rmcost").Value;
                tracingService.Trace("Rm cost sheet Printed" + rmCost);
                rmCostSheet["hil_cogs"] = totalCost + rmCost;
                service.Update(rmCostSheet);

                EntityReference tenderProduct = rmCostSheet.Contains("hil_tenderproductid") ? rmCostSheet.GetAttributeValue<EntityReference>("hil_tenderproductid") : null;
                tracingService.Trace("TendPrd id " + tenderProduct.ToString());
                if (tenderProduct != null)
                    UpdateCogsonTender(tenderProduct, totalCost + rmCost, cogsstatus);
            }
        }
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                tracingService.Trace("RMCostSheetPostUpdate Started " + context.Depth);

                if (context.MessageName == "Update" && context.InputParameters["Target"] is Entity && context.InputParameters.Contains("Target") && context.Depth <= 2)

                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace(entity.Id.ToString());
                    UpdateCOGS(entity);

                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("HavellsNewPlugin.RMCostSheetLine.RMCostSheetPostUpdate.Execute Error " + ex.Message);
            }
        }
    }
}