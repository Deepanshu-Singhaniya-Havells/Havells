using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreValidateInventoryAdjustment : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion


            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                && context.PrimaryEntityName.ToLower() == "hil_inventoryspareadjustment" && (context.MessageName.ToUpper() == "CREATE"))
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    if (!entity.Contains("hil_adjustmenttype"))
                        throw new InvalidPluginExecutionException("Adjustment Type is required.");
                    if (!entity.Contains("hil_franchise"))
                        throw new InvalidPluginExecutionException("Franchise/DSE is required.");
                    if (!entity.Contains("hil_productdivision"))
                        throw new InvalidPluginExecutionException("Product Division is required.");
                    if (!entity.Contains("hil_salesoffice"))
                        throw new InvalidPluginExecutionException("Sales Office is required.");
                    if (!entity.Contains("hil_warehouse"))
                        throw new InvalidPluginExecutionException("Warehouse is required.");
                    if (!entity.Contains("hil_adjustmentstatus"))
                        throw new InvalidPluginExecutionException("Adjustment Status is required.");
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }
    }
}
