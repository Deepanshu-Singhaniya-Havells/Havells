using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PostUpdatePurchaseReceipt : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "hil_inventorypurchaseorderreceipt" && (context.MessageName.ToUpper() == "UPDATE"))
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if (entity.Contains("hil_receiptstatus"))
                    {
                        OptionSetValue _adjStatus = entity.GetAttributeValue<OptionSetValue>("hil_receiptstatus");
                        if (_adjStatus.Value == 2)//Posted
                        {
                            Entity _entHeader = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("ownerid", "hil_franchise", "hil_warehouse"));
                            if (_entHeader != null)
                            {
                                EntityReference _entRefDefectiveWH = null;
                                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='hil_inventorywarehouse'>
                                        <attribute name='hil_inventorywarehouseid' />
                                        <order attribute='hil_name' descending='false' />
                                        <filter type='and'>
                                          <condition attribute='hil_franchise' operator='eq' value='{_entHeader.GetAttributeValue<EntityReference>("hil_franchise").Id}' />
                                          <condition attribute='hil_type' operator='eq' value='2' />
                                          <condition attribute='statecode' operator='eq' value='0' />
                                        </filter>
                                      </entity>
                                    </fetch>";
                                EntityCollection entColWH = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                if (entColWH.Entities.Count > 0)
                                {
                                    _entRefDefectiveWH = entColWH.Entities[0].ToEntityReference();
                                }
                                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
                                  <entity name='hil_inventorypurchaseorderreceiptline'>
                                    <attribute name='hil_inventorypurchaseorderreceiptlineid' />
                                    <attribute name='hil_name' />
                                    <attribute name='createdon' />
                                    <order attribute='hil_name' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='hil_receiptnumber' operator='eq' value='{entity.Id}' />
                                      <condition attribute='hil_damagequantity' operator='gt' value='0' />
                                      <condition attribute='statecode' operator='eq' value='0' />
                                    </filter>
                                  </entity>
                                </fetch>";
                                EntityCollection entColChecks = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                if (entColChecks.Entities.Count > 0 && _entRefDefectiveWH == null) {
                                    throw new InvalidPluginExecutionException("Defective Warehouse is not defined in Channel Partner.");
                                }
                                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='hil_inventorypurchaseorderreceiptline'>
                                    <attribute name='hil_inventorypurchaseorderreceiptlineid' />
                                    <attribute name='hil_partcode' />
                                    <attribute name='hil_freshquantity' />
                                    <attribute name='hil_damagequantity' />
                                    <order attribute='hil_name' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='hil_receiptnumber' operator='eq' value='{entity.Id}' />
                                      <condition attribute='statecode' operator='eq' value='0' />
                                    </filter>
                                  </entity>
                                </fetch>";
                                EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                int _qty = 0;
                                foreach (Entity _entLine in entCol.Entities) {
                                    _qty = 0;
                                    if (_entLine.Contains("hil_freshquantity"))
                                    {
                                        _qty = _entLine.GetAttributeValue<int>("hil_freshquantity");
                                        if (_qty > 0)
                                        {
                                            InventoryJournalDTO _IJData = new InventoryJournalDTO()
                                            {
                                                channelPartner = _entHeader.GetAttributeValue<EntityReference>("hil_franchise"),
                                                warehouse = _entHeader.GetAttributeValue<EntityReference>("hil_warehouse"),
                                                partCode = _entLine.GetAttributeValue<EntityReference>("hil_partcode"),
                                                quantity = _qty,
                                                isRevert = false,
                                                receiptLine = _entLine.ToEntityReference(),
                                                owner = _entHeader.GetAttributeValue<EntityReference>("ownerid"),
                                            };
                                            InventoryServices _invServices = new InventoryServices();
                                            _invServices.CreateInventoryJournal(service, _IJData);
                                        }
                                    }
                                    if (_entLine.Contains("hil_damagequantity"))
                                    {
                                        _qty = _entLine.GetAttributeValue<int>("hil_damagequantity");
                                        if (_qty > 0)
                                        {
                                            InventoryJournalDTO _IJData = new InventoryJournalDTO()
                                            {
                                                channelPartner = _entHeader.GetAttributeValue<EntityReference>("hil_franchise"),
                                                warehouse = _entRefDefectiveWH,
                                                partCode = _entLine.GetAttributeValue<EntityReference>("hil_partcode"),
                                                quantity = _qty,
                                                isRevert = false,
                                                receiptLine = _entLine.ToEntityReference(),
                                                owner = _entHeader.GetAttributeValue<EntityReference>("ownerid"),
                                            };
                                            InventoryServices _invServices = new InventoryServices();
                                            _invServices.CreateInventoryJournal(service, _IJData);
                                        }
                                    }
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
