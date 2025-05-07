using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PostCreateInventoryJournal_UpdateProductInventory : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "hil_inventoryproductjournal" && (context.MessageName.ToUpper() == "CREATE"))
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                try
                {
                    ProcessRequest(entity, _tracingService, service);
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }
        private void ProcessRequest(Entity entity, ITracingService _tracingService, IOrganizationService service)
        {
            try
            {
                Entity entityJournal = service.Retrieve("hil_inventoryproductjournal", entity.Id, new ColumnSet("hil_franchise", "hil_warehouse", "hil_partcode", "hil_quantity", "ownerid"));

                InventoryServices _invServices = new InventoryServices();
                _invServices.UpdateProductInventory(service, new ProductInventoryDTO()
                {
                    channelPartner = entityJournal.GetAttributeValue<EntityReference>("hil_franchise"),
                    warehouse = entityJournal.GetAttributeValue<EntityReference>("hil_warehouse"),
                    quantity = entityJournal.GetAttributeValue<int>("hil_quantity"),
                    partCode = entityJournal.GetAttributeValue<EntityReference>("hil_partcode"),
                    owner = entityJournal.GetAttributeValue<EntityReference>("ownerid")
                });
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
