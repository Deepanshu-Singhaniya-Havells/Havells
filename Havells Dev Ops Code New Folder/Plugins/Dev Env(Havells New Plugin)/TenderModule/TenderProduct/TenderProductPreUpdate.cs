using System;
using HavellsNewPlugin.TenderModule.OrderCheckListProduct;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.TenderProduct
{

    public class TenderProductPreUpdate : IPlugin
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
                    if (entity.Contains("hil_product"))
                    {
                        Guid _productId = entity.GetAttributeValue<EntityReference>("hil_product").Id;
                        EntityReference _hsnCode;
                        decimal taxValue = OrderCheckListProductPreCreate.getHSNValueBasedOnProduct(service, entity.GetAttributeValue<EntityReference>("hil_product"), out _hsnCode);
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
