using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PostUpdateInventoryAdjustment : IPlugin
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
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if (entity.Contains("hil_adjustmentstatus"))
                    {
                        OptionSetValue _adjStatus = entity.GetAttributeValue<OptionSetValue>("hil_adjustmentstatus");
                        if (_adjStatus.Value == 3)
                        {
                            Entity _entHeader = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("ownerid", "hil_franchise", "hil_warehouse"));
                            if (_entHeader != null)
                            {
                                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='hil_inventoryspareadjustmentline'>
                                    <attribute name='hil_inventoryspareadjustmentlineid' />
                                    <attribute name='hil_partcode' />
                                    <attribute name='hil_ajustmentquantity' />
                                    <order attribute='hil_name' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='hil_adjustmentnumber' operator='eq' value='{entity.Id}' />
                                      <condition attribute='statecode' operator='eq' value='0' />
                                    </filter>
                                  </entity>
                                </fetch>";
                                EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                foreach (Entity _entLine in entCol.Entities) {
                                    InventoryJournalDTO _IJData = new InventoryJournalDTO()
                                    {
                                        channelPartner = _entHeader.GetAttributeValue<EntityReference>("hil_franchise"),
                                        warehouse = _entHeader.GetAttributeValue<EntityReference>("hil_warehouse"),
                                        partCode = _entLine.GetAttributeValue<EntityReference>("hil_partcode"),
                                        quantity = _entLine.GetAttributeValue<int>("hil_ajustmentquantity"),
                                        isRevert = false,
                                        adjustmentLine = _entLine.ToEntityReference(),
                                        owner= _entHeader.GetAttributeValue<EntityReference>("ownerid"),
                                    };
                                    InventoryServices _invServices= new InventoryServices();
                                    _invServices.CreateInventoryJournal(service, _IJData);
                                }
                            }
                        }
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
