using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PostUpsertProductInventoryCheckForNegativeStock : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "hil_inventorysummary")
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    Process(service, entity);
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }

        public void Process(IOrganizationService service, Entity entProdInventory)
        {
            try
            {
                string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='hil_inventorysummary'>
                    <attribute name='hil_quantity'/>
                    <attribute name='hil_partcode'/>
                    <attribute name='hil_warehouse'/>
                    <attribute name='hil_inventorysummaryid'/>
                    <filter type='and'>
                        <condition attribute='hil_inventorysummaryid' operator='eq' value='{entProdInventory.Id}'/>
                    </filter>
                    </entity>
                    </fetch>";

                EntityCollection entColl = service.RetrieveMultiple(new FetchExpression(fetchXML));
                if (entColl.Entities.Count > 0)
                {
                    int _quantity = entColl.Entities[0].GetAttributeValue<int>("hil_quantity");
                    string _partCode = entColl.Entities[0].GetAttributeValue<EntityReference>("hil_partcode").Name;
                    string _warehouse = entColl.Entities[0].GetAttributeValue<EntityReference>("hil_warehouse").Name;
                    if (_quantity < 0) {
                        throw new Exception($"Stock is not available.{System.Environment.NewLine} Part Code: {_partCode}, {System.Environment.NewLine} Warehouse: {_warehouse}");
                    }
                }
            }
            catch(Exception ex)
            {
                throw new Exception($"ERROR! {ex.Message}");
            }
        }
    }
}
