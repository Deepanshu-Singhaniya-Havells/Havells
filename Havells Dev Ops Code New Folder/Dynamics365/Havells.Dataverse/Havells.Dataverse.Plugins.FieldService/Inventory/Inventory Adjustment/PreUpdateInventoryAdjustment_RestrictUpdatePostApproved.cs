using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreUpdateInventoryAdjustment_RestrictUpdatePostApproved : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "hil_inventoryspareadjustment" && (context.MessageName.ToUpper() == "UPDATE"))
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                Entity entityRecord = service.Retrieve("hil_inventoryspareadjustment", entity.Id, new ColumnSet("hil_adjustmentstatus"));
                OptionSetValue _adjustmentStatus = entityRecord.GetAttributeValue<OptionSetValue>("hil_adjustmentstatus");
                if (_adjustmentStatus.Value == 3)
                {
                    throw new InvalidPluginExecutionException("Updation is not allowed once Adjustment is approved.");
                }
            }
        }
    }
}
