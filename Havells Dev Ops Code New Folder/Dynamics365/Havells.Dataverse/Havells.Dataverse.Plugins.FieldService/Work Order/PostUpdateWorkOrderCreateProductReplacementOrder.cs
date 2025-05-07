using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Havells.Dataverse.Plugins.FieldService.Work_Order
{
    public class PostUpdateWorkOrderCreateProductReplacementOrder : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "msdyn_workorderincident" && (context.MessageName.ToUpper() == "UPDATE") && context.Depth == 1)
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                if (entity.Contains("hil_flagpo"))
                    ProcessRequest(entity, _tracingService, service);
            }
        }
        public void ProcessRequest(Entity entity, ITracingService _tracingService, IOrganizationService service)
        {
            try
            {
                int _isFlagPO = entity.GetAttributeValue<OptionSetValue>("hil_productreplacement").Value;
                if (_isFlagPO == 1)
                {
                    Entity workorderincident = service.Retrieve("msdyn_workorderincident", entity.Id, new ColumnSet("hil_modelcode", "msdyn_workorder", "hil_quantity", "hil_warrantystatus"));
                    Guid workorderId = workorderincident.Contains("msdyn_workorder") ? workorderincident.GetAttributeValue<EntityReference>("msdyn_workorder").Id : Guid.Empty;
                    JobDataPR data = new JobDataPR();
                    data.WorkOrder = service.Retrieve("msdyn_workorder", workorderId, new ColumnSet("msdyn_substatus", "hil_owneraccount", "hil_productcategory", "hil_salesoffice", "hil_brand"));
                    data.Franchise = data.WorkOrder.GetAttributeValue<EntityReference>("hil_owneraccount");
                    Entity franchise = service.Retrieve(data.Franchise.LogicalName, data.Franchise.Id, new ColumnSet("ownerid", "hil_spareinventoryenabled"));
                    bool isInventoryEnabled = franchise.GetAttributeValue<Boolean>("hil_spareinventoryenabled");
                    if (isInventoryEnabled)
                    {
                        data.SubStatus = data.WorkOrder.GetAttributeValue<EntityReference>("msdyn_substatus");
                        data.ProductCategory = data.WorkOrder.GetAttributeValue<EntityReference>("hil_productcategory");
                        data.SalesOffice = data.WorkOrder.GetAttributeValue<EntityReference>("hil_salesoffice");
                        data.Brand = data.WorkOrder.GetAttributeValue<OptionSetValue>("hil_brand");
                        EntityReference franchiseOwner = franchise.Contains("ownerid") ? franchise.GetAttributeValue<EntityReference>("ownerid") : null;
                        Entity warehouse = GetFreshWarehouse(data.Franchise.Id, service);
                        if (warehouse != null)
                        {
                            // Create purchase order
                            Guid purchaseOrder = DoesOrderExists(data, warehouse, service);
                            if (purchaseOrder == Guid.Empty)
                            {
                                purchaseOrder = CreatePurchaseOrder(data, warehouse, franchiseOwner, service);
                            }
                            // Create purchase order lines.
                            EntityReference product = workorderincident.Contains("hil_modelcode") ? workorderincident.GetAttributeValue<EntityReference>("hil_modelcode") : null;
                            int orderQuantity = workorderincident.Contains("hil_quantity") ? workorderincident.GetAttributeValue<int>("hil_quantity") : 0;
                            int warrantyStatus = workorderincident.Contains("hil_warrantystatus") ? workorderincident.GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value : 1;
                            CreateOrderLine(workorderincident, purchaseOrder, product, orderQuantity, warrantyStatus, franchiseOwner, service);
                            //Update CRMAdmin as approver and approve the Purchase Order.
                            UpdateApproverOnPurchaseOrder(purchaseOrder, service);
                            ChangeWorkOrderSubstatus(entity.Id, service);
                        }
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException($"You are not authorised to create Emergency Order. Please contact with HO.");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException($"ERROR.{ex.Message}");
            }
        }
        private bool DoesOrderLineExists(Entity jobProduct, Guid purchaseOrder, EntityReference product, double orderQuantity, int warrantyStatus, IOrganizationService service)
        {
            QueryExpression query = new QueryExpression("hil_inventorypurchaseorderline");
            query.Criteria.AddCondition("hil_jobproduct", ConditionOperator.Equal, jobProduct.Id);
            query.Criteria.AddCondition("hil_ponumber", ConditionOperator.Equal, purchaseOrder);
            query.Criteria.AddCondition("hil_partcode", ConditionOperator.Equal, product.Id);
            EntityCollection tempColl = service.RetrieveMultiple(query);
            if (tempColl.Entities.Count > 0)
            {
                return true;
            }
            return false;
        }

        private Guid CreateOrderLine(Entity jobincident, Guid purchaseOrder, EntityReference product, double orderQuantity, int warrantyStatus, EntityReference owner, IOrganizationService service)
        {
            Entity newPurchaseOrderLine = new Entity("hil_inventorypurchaseorderline");
            newPurchaseOrderLine["hil_jobincident"] = jobincident.ToEntityReference();
            newPurchaseOrderLine["hil_ponumber"] = new EntityReference("hil_inventorypurchaseorder", purchaseOrder);
            newPurchaseOrderLine["hil_partcode"] = product;
            newPurchaseOrderLine["ownerid"] = owner;
            if (orderQuantity != -1) newPurchaseOrderLine["hil_orderquantity"] = (int)orderQuantity;
            if (warrantyStatus != -1) newPurchaseOrderLine["hil_warrantystatus"] = new OptionSetValue(warrantyStatus);

            return service.Create(newPurchaseOrderLine);
        }

        private Guid DoesOrderExists(JobDataPR data, Entity warehouse, IOrganizationService service)
        {
            Guid _purchaseOrderId = Guid.Empty;
            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='hil_inventorypurchaseorder'>
            <attribute name='hil_inventorypurchaseorderid' />
            <attribute name='hil_name' />
            <order attribute='hil_name' descending='false' />
            <filter type='and'>
                <condition attribute='hil_jobid' operator='eq' value='{data.WorkOrder.Id}' />
                <condition attribute='hil_ordertype' operator='eq' value='1' />
                <condition attribute='hil_postatus' operator='ne' value='4' />
                <condition attribute='statecode' operator='eq' value='0' />
            </filter>
            </entity>
            </fetch>";
            EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
            if (entCol.Entities.Count > 0)
            {
                _purchaseOrderId = entCol.Entities[0].Id;
            }
            return _purchaseOrderId;
        }

        private void UpdateApproverOnPurchaseOrder(Guid purchaseOrder, IOrganizationService service)
        {
            Guid CrmAdmin = new Guid("5190416c-0782-e911-a959-000d3af06a98");
            Entity updatePurchaseOrder = new Entity("hil_inventorypurchaseorder", purchaseOrder);
            updatePurchaseOrder["hil_approver"] = new EntityReference("systemuser", CrmAdmin);
            updatePurchaseOrder["hil_approvedby"] = new EntityReference("systemuser", CrmAdmin);
            updatePurchaseOrder["hil_approvedon"] = DateTime.Now;
            updatePurchaseOrder["hil_postatus"] = new OptionSetValue(3); //Approved
            service.Update(updatePurchaseOrder);
        }

        private void ChangeWorkOrderSubstatus(Guid workOrderId, IOrganizationService service)
        {
            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='msdyn_workordersubstatus'>
                <attribute name='msdyn_systemstatus' />
                <attribute name='msdyn_workordersubstatusid' />
                <filter type='and'>
                    <condition attribute='msdyn_name' operator='eq' value='Part PO Created' />
                </filter>
                </entity>
                </fetch>";
            EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
            if (entCol.Entities.Count > 0)
            {
                Entity entUpdate = new Entity("msdyn_workorder", workOrderId);
                entUpdate["msdyn_substatus"] = entCol.Entities[0].ToEntityReference();
                entUpdate["msdyn_systemstatus"] = entCol.Entities[0].GetAttributeValue<OptionSetValue>("msdyn_systemstatus");
                service.Update(entUpdate);
            }
        }
        private Guid CreatePurchaseOrder(JobDataPR data, Entity warehouse, EntityReference owner, IOrganizationService service)
        {
            Entity newPurchaseOrder = new Entity("hil_inventorypurchaseorder");
            newPurchaseOrder["hil_jobid"] = new EntityReference(data.WorkOrder.LogicalName, data.WorkOrder.Id);
            newPurchaseOrder["hil_productdivision"] = data.ProductCategory;
            newPurchaseOrder["hil_salesoffice"] = data.SalesOffice;
            newPurchaseOrder["hil_franchise"] = data.Franchise;
            newPurchaseOrder["hil_warehouse"] = warehouse.ToEntityReference();
            newPurchaseOrder["ownerid"] = owner;
            newPurchaseOrder["hil_postatus"] = new OptionSetValue(1);
            newPurchaseOrder["hil_ordertype"] = new OptionSetValue(1);
            newPurchaseOrder["hil_brand"] = data.Brand;
            return service.Create(newPurchaseOrder);
        }

        private Entity GetFreshWarehouse(Guid franchise, IOrganizationService service)
        {
            QueryExpression query = new QueryExpression("hil_inventorywarehouse");
            query.Criteria.AddCondition("hil_franchise", ConditionOperator.Equal, franchise);
            query.Criteria.AddCondition("hil_type", ConditionOperator.Equal, 1); // Fresh

            EntityCollection warehouseColl = service.RetrieveMultiple(query);
            if (warehouseColl.Entities.Count > 0)
            {
                return warehouseColl.Entities[0];
            }
            return null;
        }
    }

    class JobDataPR
    {
        public Entity WorkOrder { get; set; }
        public EntityReference SubStatus { get; set; }
        public EntityReference SalesOffice { get; set; }
        public EntityReference ProductCategory { get; set; }
        public EntityReference Franchise { get; set; }
        public OptionSetValue Brand { get; set; }
    }
}
