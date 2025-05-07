using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreValidateUpdatePurchaseOrder : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "hil_inventorypurchaseorder" && (context.MessageName.ToUpper() == "UPDATE"))
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    if (entity.Contains("hil_ordertype"))
                        throw new InvalidPluginExecutionException("Order Type can't be changed.");
                    if (entity.Contains("hil_franchise"))
                        throw new InvalidPluginExecutionException("Franchise can't be changed.");
                    if (entity.Contains("hil_warehouse"))
                        throw new InvalidPluginExecutionException("Warehouse can't be changed.");
                    if (entity.Contains("hil_salesoffice"))
                        throw new InvalidPluginExecutionException("Sales Office can't be changed.");

                    Entity _entPO = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_postatus"));
                    OptionSetValue _poStatus = _entPO.GetAttributeValue<OptionSetValue>("hil_postatus");
                    if (_poStatus.Value != 1)
                    {
                        if (entity.Contains("hil_brand"))
                            throw new InvalidPluginExecutionException("Brand can't be changed.");
                        if (entity.Contains("hil_productdivision"))
                            throw new InvalidPluginExecutionException("Product Division can't be changed.");
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }
    }
}