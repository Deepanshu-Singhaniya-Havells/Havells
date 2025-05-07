using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class InventoryServices
    {
        private readonly EntityReference WORKINITIATED = new EntityReference("msdyn_workordersubstatus", new Guid("1b27fa6c-fa0f-e911-a94e-000d3af060a1"));
        private readonly EntityReference PARTPOCREATED = new EntityReference("msdyn_workordersubstatus", new Guid("2b27fa6c-fa0f-e911-a94e-000d3af060a1"));
        public void CreateInventoryJournal(IOrganizationService service, InventoryJournalDTO data) {
            try
            {
                Entity _entIJ = new Entity("hil_inventoryproductjournal");
                _entIJ["hil_franchise"] = data.channelPartner;
                _entIJ["hil_warehouse"] = data.warehouse;
                _entIJ["hil_partcode"] = data.partCode;
                _entIJ["hil_quantity"] = data.quantity;
                _entIJ["hil_isrevert"] = data.isRevert;
                _entIJ["hil_receiptline"] = data.receiptLine;
                _entIJ["hil_adjustmentline"] = data.adjustmentLine;
                _entIJ["hil_jobproduct"] = data.jobProduct;
                _entIJ["hil_rmaline"] = data.rmaLine;
                _entIJ["hil_rma"] = data.rma;
                _entIJ["hil_rmatype"] = data.returnType;
                _entIJ["hil_isused"] = false;
                int _transactionType = data.receiptLine != null ? 1 : data.adjustmentLine != null ? 2 : data.jobProduct != null ? 3 : 4;
                _entIJ["hil_transactiontype"] = new OptionSetValue(_transactionType);
                _entIJ["ownerid"] = data.owner;
                service.Create(_entIJ);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void UpdateProductInventory(IOrganizationService service, ProductInventoryDTO data)
        {
            try
            {
                string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_inventorysummary'>
                        <attribute name='hil_inventorysummaryid'/>
                        <attribute name='hil_quantity'/>
                        <order attribute='hil_partcode' descending='false'/>
                        <filter type='and'>
                            <condition attribute='hil_warehouse' operator='eq' value='{data.warehouse.Id}'/>
                            <condition attribute='hil_partcode' operator='eq' value='{data.partCode.Id}'/>
                            <condition attribute='hil_franchise' operator='eq' value='{data.channelPartner.Id}'/>
                            <condition attribute='statecode' operator='eq' value='0'/>
                        </filter>
                        </entity>
                        </fetch>";

                EntityCollection entColl = service.RetrieveMultiple(new FetchExpression(fetchXML));
                if (entColl.Entities.Count > 0)
                {
                    Entity entInventorySummary = new Entity(entColl.Entities[0].LogicalName, entColl.Entities[0].Id);
                    entInventorySummary["hil_quantity"] = entColl.Entities[0].GetAttributeValue<int>("hil_quantity") + data.quantity;
                    service.Update(entInventorySummary);
                }
                else
                {
                    Entity entInventorySummary = new Entity("hil_inventorysummary");
                    entInventorySummary["hil_warehouse"] = data.warehouse;
                    entInventorySummary["hil_partcode"] = data.partCode;
                    entInventorySummary["hil_franchise"] = data.channelPartner;
                    entInventorySummary["hil_quantity"] = data.quantity;
                    entInventorySummary["ownerid"] = data.owner;
                    service.Create(entInventorySummary);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void UpdateWorkOrderToWorkInitiated(IOrganizationService service, EntityReference workOrderRef)
        {
            Entity woObj = service.Retrieve(workOrderRef.LogicalName, workOrderRef.Id, new ColumnSet("msdyn_substatus"));
            EntityReference _jobSubStatus = woObj.GetAttributeValue<EntityReference>("msdyn_substatus");

            if (_jobSubStatus.Id == PARTPOCREATED.Id)//Part PO Created
            {
                Entity workOrder = new Entity("msdyn_workorder", workOrderRef.Id);
                workOrder["msdyn_substatus"] = WORKINITIATED;
                service.Update(workOrder);
            }
        }
        private EntityReference CreatePurchaseReceipt(IOrganizationService service, SaleInvoiceDTO data)
        {
            EntityReference _retValue = null;
            try
            {
                Entity _entPR = new Entity("hil_inventorypurchasereceipt");
                _entPR["hil_franchise"] = data.channelPartner;
                var _retObj = this.GetChannelPartnerWarehouse(service, new GetWarehouseDTO { channelPartner = data.channelPartner, warehouseType = WarehouseType.Fresh });
                if (_retObj.Item1 == null) {
                    throw new Exception($"ERROR!{this.GetType().FullName} {_retObj.Item3}");
                }
                _entPR["hil_warehouse"] = _retObj.Item1;
                _entPR["hil_invoicenumber"] = data.billNumber;
                _entPR["hil_receiptstatus"] = new OptionSetValue(1);//Draft
                _entPR["ownerid"] = this.GetChannelPartnerOwner(service, data.channelPartner);
                Guid _prGuid = service.Create(_entPR);
                _retValue = new EntityReference("hil_inventorypurchasereceipt", _prGuid);
            }
            catch (Exception ex)
            {
                throw new Exception($"ERROR!{this.GetType().FullName} {ex.Message} {ex.StackTrace}");
            }
            return _retValue;
        }
        private Tuple<EntityReference,bool,string> GetChannelPartnerWarehouse (IOrganizationService service, GetWarehouseDTO _data)
        {
            EntityReference _warehouse = null;
            string _error = string.Empty;
            bool _retValue = false;
            try
            {
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_inventorywarehouse'>
                <attribute name='hil_inventorywarehouseid' />
                <attribute name='hil_name' />
                <filter type='and'>
                    <condition attribute='hil_franchise' operator='eq' value='{_data.channelPartner.Id}' />
                    <condition attribute='hil_type' operator='eq' value='{_data.warehouseType}' />
                    <condition attribute='statecode' operator='eq' value='0' />
                </filter>
                </entity>
                </fetch>";
                EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (_entCol.Entities.Count > 0)
                {
                    _warehouse = _entCol.Entities[0].ToEntityReference();
                    _retValue = true;
                }
                else
                {
                    _error = "Warehouse Setup is not defined. Please conenct with HO Team.";
                }
            }
            catch (Exception ex)
            {
                _error = $"{ex.Message} {ex.StackTrace}"; 
            }
            return Tuple.Create(_warehouse, _retValue, _error);
        }

        private EntityReference GetChannelPartnerOwner(IOrganizationService service, EntityReference _channelPartner)
        {
            EntityReference _owner = null;
            try
            {
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='account'>
                <attribute name='ownerid' />
                <filter type='and'>
                    <condition attribute='accountid' operator='eq' value='{_channelPartner.Id}' />
                </filter>
                </entity>
                </fetch>";
                EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (_entCol.Entities.Count > 0)
                {
                    _owner = _entCol.Entities[0].GetAttributeValue<EntityReference>("ownerid");
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"{ex.Message} {ex.StackTrace}");
            }
            return _owner;
        }
    }
    public class InventoryJournalDTO { 
        public EntityReference channelPartner { get;set; }
        public EntityReference warehouse { get; set; }
        public EntityReference partCode { get; set; }
        public int quantity { get; set; }
        public OptionSetValue transactionType { get; set; }
        public bool isRevert { get; set; }
        public EntityReference receiptLine { get; set; }
        public EntityReference adjustmentLine { get; set; }
        public EntityReference jobProduct { get; set; }
        public EntityReference rmaLine { get; set; }
        public EntityReference rma { get; set; }
        public EntityReference owner { get; set; }
        public EntityReference returnType { get; set; }
    }
    public class ProductInventoryDTO
    {
        public EntityReference channelPartner { get; set; }
        public EntityReference warehouse { get; set; }
        public EntityReference partCode { get; set; }
        public int quantity { get; set; }
        public EntityReference owner { get; set; }
    }

    public class SaleInvoiceDTO {
        public string billNumber { get; set; }
        public string salesOrderNumber { get; set; }
        public EntityReference purchaseOrder { get; set; }
        public EntityReference channelPartner { get; set; }
        public EntityReference partCode { get; set; }
        public int billedQuantity { get; set; }
    }
    public class GetWarehouseDTO
    {
        public EntityReference channelPartner { get; set; }
        public WarehouseType warehouseType { get; set; }
    }
    public enum WarehouseType
    {
        Fresh = 1,
        Defective = 2
    }
}
