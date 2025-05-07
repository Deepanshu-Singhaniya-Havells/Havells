using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;

namespace HavellsNewPlugin.TenderModule.OrderCheckListProduct
{
    public class PreUpdate : IPlugin
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
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    //Entity _entTender = service.Retrieve("hil_tender", entity.GetAttributeValue<EntityReference>("hil_tenderid").Id, new ColumnSet("hil_pricelist", "hil_name"));

                    if (entity.Contains("hil_product"))
                    {
                        EntityReference _product = entity.GetAttributeValue<EntityReference>("hil_product");
                        EntityReference _hsnCode;
                        decimal taxValue = OrderCheckListProductPreCreate.getHSNValueBasedOnProduct(service, _product, out _hsnCode);
                        entity["hil_tax"] = taxValue;
                        entity["hil_hsncode"] = _hsnCode;
                    }
                    else if (entity.Contains("hil_hsncode"))
                    {
                        decimal taxValue = OrderCheckListProductPreCreate.getHSNValueBasedOnHSN(service, entity.GetAttributeValue<EntityReference>("hil_hsncode"));
                        entity["hil_tax"] = taxValue;
                    }
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error " + ex.Message);
                throw new InvalidPluginExecutionException("HavellsNewPlugin.TenderModule.PostUpdateTenderProduct.Execute Error " + ex.Message);
            }
        }
    }
}
