using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.FieldService.RMA.InventoryAdjustment
{
    public class PostUpdate : InventoryAdjustment_Helper, IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if (entity.Contains("hil_adjustmentstatus"))
                    {
                        InventoryAdjustmentStatus hil_adjustmentstatus = (InventoryAdjustmentStatus)entity.GetAttributeValue<OptionSetValue>("hil_adjustmentstatus").Value;
                        if(hil_adjustmentstatus == InventoryAdjustmentStatus.Approved)
                        {
                            UpdateQuantityOnInventoryAdjustmentProduct(service, entity.ToEntityReference());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error !!\n " + ex.Message);
            }

        }
    }
}