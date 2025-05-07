using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IdentityModel.Metadata;

namespace HavellsNewPlugin.SalesOrder
{
    public class PreCreate_SalesOrder : IPlugin
    {
        private static ITracingService tracingService = null;
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
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.PrimaryEntityName.ToLower() == "salesorder")
                {
                    Entity salesOrder = (Entity)context.InputParameters["Target"];

                    if (salesOrder.Contains("hil_serviceaddress"))
                    {
                        EntityReference addresReference = (EntityReference)salesOrder["hil_serviceaddress"];
                        Entity orderheaderaddress = service.Retrieve("hil_address", addresReference.Id, new ColumnSet("hil_businessgeo"));
                        Entity orderheaderbranch = service.Retrieve("hil_businessmapping", orderheaderaddress.GetAttributeValue<EntityReference>("hil_businessgeo").Id, new ColumnSet("hil_branch"));

                        if (orderheaderbranch != null)
                        {
                            salesOrder["hil_branch"] = orderheaderbranch.Contains("hil_branch") ? new EntityReference("hil_branch", orderheaderbranch.GetAttributeValue<EntityReference>("hil_branch").Id) : null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("HavellsNewPlugin.SalesOrderPreCreate.Execute Error" + ex.Message);
            }
        }
        
    }
}

