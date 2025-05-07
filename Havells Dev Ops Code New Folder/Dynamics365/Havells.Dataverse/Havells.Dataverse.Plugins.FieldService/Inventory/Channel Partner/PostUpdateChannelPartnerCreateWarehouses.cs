using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PostUpdateChannelPartnerCreateWarehouses : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "account" && (context.MessageName.ToUpper() == "UPDATE"))
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                if (entity.Contains("hil_spareinventoryenabled"))
                {
                    bool _inventoryEnabled = entity.GetAttributeValue<bool>("hil_spareinventoryenabled");
                    if (_inventoryEnabled)
                        ProcessRequest(entity, _tracingService, service);
                }
            }
        }

        private void ProcessRequest(Entity entity, ITracingService _tracingService, IOrganizationService service)
        {
            try
            {
                Entity _entAccount = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_spareinventoryenabled", "ownerid", "accountnumber"));
                if (_entAccount != null)
                {
                    if (!_entAccount.Contains("accountnumber"))
                    {
                        throw new InvalidPluginExecutionException("Data Validation Failed.\n Account Number is not Set on the Channel Partner.");
                    }
                    else
                    {
                        CreateWarehouse(_entAccount, service, 1);//Fresh Warehouse
                        CreateWarehouse(_entAccount, service, 2);//Defective Warehouse
                    }
                }
                else
                {
                    throw new InvalidPluginExecutionException("Not allowed.\n Warehouse Setup is only applicable for Franchise/DSE Channel Partners.");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private void CreateWarehouse(Entity entity, IOrganizationService service, int warehouseType)
        {
            try
            {
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_inventorywarehouse'>
                <attribute name='hil_inventorywarehouseid' />
                <filter type='and'>
                    <condition attribute='hil_franchise' operator='eq' value='{entity.Id}' />
                    <condition attribute='hil_type' operator='eq' value='{warehouseType}' />
                    <condition attribute='statecode' operator='eq' value='0' />
                </filter>
                </entity>
                </fetch>";
                EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (_entCol.Entities.Count == 0)
                {
                    Entity _entWarehouse = new Entity("hil_inventorywarehouse");
                    _entWarehouse["hil_franchise"] = entity.ToEntityReference();
                    _entWarehouse["hil_type"] = new OptionSetValue(warehouseType);
                    service.Create(_entWarehouse);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
