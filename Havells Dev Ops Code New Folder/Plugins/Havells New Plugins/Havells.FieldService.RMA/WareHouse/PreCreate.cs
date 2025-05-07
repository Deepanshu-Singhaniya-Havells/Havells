using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.FieldService.RMA.WareHouse
{
    public class PreCreate : IPlugin
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
                    EntityReference ownerRef = entity.GetAttributeValue<EntityReference>("ownerid");
                    string accountNUmber = WareHouse_Helper.GetAccountNumber(service, ownerRef);
                    string warehouseType = entity.FormattedValues["hil_warehousetype"];
                    entity["msdyn_name"] = accountNUmber + "-" + warehouseType;
                }

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error !!\n " + ex.Message);
            }

        }
    }
}