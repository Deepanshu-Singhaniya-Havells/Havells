using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PostCreateSpareSalesInvoice : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "hil_inventorysparebills" && (context.MessageName.ToUpper() == "CREATE"))
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                ProcessRequest(entity, _tracingService, service);
            }
        }

        private void ProcessRequest(Entity entity, ITracingService _tracingService, IOrganizationService service)
        {
            try
            {
                Entity _salesInvoice = service.Retrieve("hil_inventorysparebills", entity.Id, new ColumnSet("hil_billnumber"));

                string _query = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_inventorypurchaseorderreceipt'>
                <attribute name='hil_inventorypurchaseorderreceiptid' />
                <filter type='and'>
                    <condition attribute='hil_invoicenumber' operator='eq' value='{_salesInvoice.GetAttributeValue<string>("hil_billnumber")}' />
                </filter>
                </entity>
                </fetch>";

                EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_query));
                if (_entCol.Entities.Count > 0)
                {
                    if (_entCol.Entities[0].Contains("billqty"))
                    {
                        int _totalBillQty = (int)_entCol.Entities[0].GetAttributeValue<AliasedValue>("billqty").Value;
                        EntityReference _poline = (EntityReference)_entCol.Entities[0].GetAttributeValue<AliasedValue>("poline").Value;
                        Entity _entUpdate = new Entity(_poline.LogicalName, _poline.Id);
                        _entUpdate["hil_suppliedquantity"] = _totalBillQty;
                        service.Update(_entUpdate);

                        Entity purchaseOrderLine = service.Retrieve(_poline.LogicalName, _poline.Id, new ColumnSet("hil_pendingquantity", "hil_suppliedquantity", "hil_workorder"));
                        int pendingQuantity = 0;
                        int suppliedQuantity = 0;
                        EntityReference _job = null;

                        if (purchaseOrderLine.Contains("hil_workorder"))
                            _job = purchaseOrderLine.GetAttributeValue<EntityReference>("hil_workorder");

                        if (purchaseOrderLine.Contains("hil_suppliedquantity"))
                            suppliedQuantity = purchaseOrderLine.GetAttributeValue<int>("hil_suppliedquantity");
                        if (purchaseOrderLine.Contains("hil_pendingquantity"))
                            pendingQuantity = purchaseOrderLine.GetAttributeValue<int>("hil_pendingquantity");

                        if (suppliedQuantity > 0 && pendingQuantity == 0)
                        {
                            purchaseOrderLine["hil_partstatus"] = new OptionSetValue(3); //Dispatched
                        }
                        else if (suppliedQuantity > 0 && pendingQuantity != 0)
                        {
                            purchaseOrderLine["hil_partstatus"] = new OptionSetValue(2); //Partially Dispatched
                        }
                        service.Update(purchaseOrderLine);

                        if (_job != null) {
                            Entity _entJob = service.Retrieve(_job.LogicalName, _job.Id, new ColumnSet("msdyn_substatus"));
                            if (_entJob.Contains("msdyn_substatus"))
                            {
                                EntityReference _jobSubStatus = _entJob.GetAttributeValue<EntityReference>("msdyn_substatus");
                                if (_jobSubStatus.Id == new Guid("1b27fa6c-fa0f-e911-a94e-000d3af060a1")) // Part PO Created 
                                {
                                    _query = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='hil_inventorypurchaseorderline'>
                                        <attribute name='hil_inventorypurchaseorderlineid' />
                                        <filter type='and'>
                                          <condition attribute='hil_workorder' operator='eq' value='{_entJob.Id}' />
                                          <condition attribute='hil_partstatus' operator='not-in'>
                                            <value>3</value>
                                            <value>5</value>
                                          </condition>
                                        </filter>
                                      </entity>
                                    </fetch>";
                                    EntityCollection _entColPOLine = service.RetrieveMultiple(new FetchExpression(_query));
                                    if (_entColPOLine.Entities.Count == 0)
                                    {
                                        InventoryServices _inventoryService = new InventoryServices();
                                        _inventoryService.UpdateWorkOrderToWorkInitiated(service, _entJob.ToEntityReference());
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
