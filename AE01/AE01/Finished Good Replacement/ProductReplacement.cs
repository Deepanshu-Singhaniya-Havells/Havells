using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Runtime.CompilerServices;
using NLog.Layouts;

namespace AE01.Finished_Good_Replacement
{
    public class ProductReplacement
    {
        private readonly IOrganizationService service;
        public ProductReplacement(IOrganizationService service)
        {
            this.service = service; 
        }
        internal void Program()
        {
            Entity jobIncident = service.Retrieve("msdyn_workorderincident", new Guid("7acb1914-b2e6-ef11-9342-6045bd729d2f"), new ColumnSet(false));
            CreatePurchaseOrder(service, jobIncident.Id);
        }

        private void CreateSAWActivity(JobData jobData)
        {
            Guid productReplacementCategory = new Guid("c679c5fa-7fee-ef11-9341-6045bdc64f83");
            Entity sawActivity = new Entity("hil_sawactivity");
            sawActivity["hil_sawcategory"] = new EntityReference("hil_serviceactionwork", productReplacementCategory);
            sawActivity["hil_relatedchannelpartner"] = jobData.Franchise;
            sawActivity["hil_jobid"] = jobData.WorkOrder.ToEntityReference();
            sawActivity["hil_approvalstatus"] = new OptionSetValue(1); //Requested
            Guid sawActivityId = service.Create(sawActivity);
        }

        private void CreatePurchaseOrder(IOrganizationService service, Guid jobIncidentId)
        {   
            JobData jobData = new();
            AssetData assetData = new();
            
            jobData.JobIncident = service.Retrieve("msdyn_workorderincident", jobIncidentId, new ColumnSet("msdyn_workorder", "hil_modelcode", "hil_warrantystatus", "hil_reasonofreplacement", "hil_replacewithsamemodel", "hil_replacedquantity", "hil_replacewithmodel", "msdyn_customerasset"));
            assetData.Asset = service.Retrieve("msdyn_customerasset", jobData.JobIncident.GetAttributeValue<EntityReference>("msdyn_customerasset").Id, new ColumnSet("hil_invoicedate", "hil_warrantystatus", "hil_warrantysubstatus")); 
            jobData.WorkOrder = service.Retrieve("msdyn_workorder", ((EntityReference)jobData.JobIncident.Attributes["msdyn_workorder"]).Id, new ColumnSet("hil_owneraccount", "hil_brand", "hil_productcategory", "hil_salesoffice", "ownerid", "hil_productcatsubcatmapping", "hil_callsubtype", "createdon", "msdyn_substatus"));
            jobData.Franchise = jobData.WorkOrder.GetAttributeValue<EntityReference>("hil_owneraccount");
            jobData.CustomerAssetRef = jobData.JobIncident.GetAttributeValue<EntityReference>("msdyn_customerasset"); 
            if (jobData.JobIncident.GetAttributeValue<bool>("hil_replacewithsamemodel"))
            {
                jobData.PartCode = jobData.JobIncident.GetAttributeValue<EntityReference>("hil_modelcode");
            }
            else
            {
                jobData.PartCode = jobData.JobIncident.GetAttributeValue<EntityReference>("hil_replacewithmodel");
            }

            Entity franchise = service.Retrieve(jobData.Franchise.LogicalName, jobData.Franchise.Id, new ColumnSet("ownerid", "hil_spareinventoryenabled"));
            bool isInventoryEnabled = franchise.GetAttributeValue<Boolean>("hil_spareinventoryenabled");
            if (isInventoryEnabled)
            {   
                jobData.CallSubType = jobData.WorkOrder.GetAttributeValue<EntityReference>("hil_callsubtype");  
                jobData.SubStatus = jobData.WorkOrder.GetAttributeValue<EntityReference>("msdyn_substatus");
                jobData.ProductCategory = jobData.WorkOrder.GetAttributeValue<EntityReference>("hil_productcategory");
                jobData.ProductSubCategory = jobData.WorkOrder.GetAttributeValue<EntityReference>("hil_productcatsubcatmapping");
                jobData.SalesOffice = jobData.WorkOrder.GetAttributeValue<EntityReference>("hil_salesoffice");
                jobData.Brand = jobData.WorkOrder.GetAttributeValue<OptionSetValue>("hil_brand");
                jobData.JobOwner = jobData.WorkOrder.GetAttributeValue<EntityReference>("ownerid");
                jobData.Quantity = jobData.JobIncident.Contains("hil_quantity") ? jobData.JobIncident.GetAttributeValue<int>("hil_quantity") : -1;
                jobData.FranchiseOwner = franchise.Contains("ownerid") ? franchise.GetAttributeValue<EntityReference>("ownerid") : null;
                jobData.Warehouse = GetFreshWarehouse(jobData.Franchise.Id, service);
                if (jobData.Warehouse != null && jobData.FranchiseOwner != null)
                {
                    // If not create one 
                    // else Order already exists
                    // Check whether the Activity already exist
                    if (!DoesActivityExists(ref jobData, service))
                    {
                        // Create SAW Activity  
                        jobData.SAWActivity = CreateActivity(ref jobData, service);

                        // Create SAW Approvals 
                        assetData.WarrantyStatus = assetData.Asset.GetAttributeValue<OptionSetValue>("hil_warrantystatus");
                        assetData.WarrantySubStatus = assetData.Asset.GetAttributeValue<OptionSetValue>("hil_warrantysubstatus");
                        assetData.InvoiceDate = assetData.Asset.GetAttributeValue<DateTime>("hil_invoicedate");
                        assetData.Aging = (jobData.WorkOrder.GetAttributeValue<DateTime>("createdon") - assetData.InvoiceDate).Days;  
                        CreateSAWApprovals(ref jobData, ref assetData);
                    }
                    
                    bool lineExists = DoesOrderLineExists(jobData, service);
                    if (!lineExists)
                    {
                        // Create purchase order
                        jobData.PurchaseOrder = DoesOrderExists(jobData, service);
                        if (jobData.PurchaseOrder == Guid.Empty) jobData.PurchaseOrder = CreateOrder(jobData, service);
                        // Check whether purchase Order Line Already Exists or not 
                        CreateOrderLine(jobData, service);
                    }
                    else
                    {
                        // Order already exists
                    }
                    // Change job Sub Status
                    UpdateJobSubStatus(jobData, service);
                    // SAW Activity Approvalls will automatically be created.  
                }
            }
            else
            {
                // Inventory is not enabled for the particular franchise.
            }
        }

        private void CreateSAWApprovals(ref JobData jobData, ref AssetData assetData)
        {
            // Fetch the SAW Approvals for the the corresponding incident
            EntityCollection sawApprovalMatrixCollection = FetchSAWApprovals(jobData, assetData);

            QueryExpression fetchSBUMapping = new QueryExpression("hil_sbubranchmapping");
            fetchSBUMapping.ColumnSet = new ColumnSet("hil_nsh", "hil_nph", "hil_branchheaduser");
            fetchSBUMapping.Criteria = new FilterExpression(LogicalOperator.And);
            fetchSBUMapping.Criteria.AddCondition("hil_salesoffice", ConditionOperator.Equal, jobData.SalesOffice.Id);
            fetchSBUMapping.Criteria.AddCondition("hil_productdivision", ConditionOperator.Equal, jobData.ProductCategory.Id);
            EntityCollection sbuMappingColl = service.RetrieveMultiple(fetchSBUMapping);

            foreach (Entity sawApprovalMatrix in sawApprovalMatrixCollection.Entities)
            {
                EntityReference approver = FetchApprover(sawApprovalMatrix, sbuMappingColl);
                Entity entSAWActivity = new Entity("hil_sawactivityapproval");
                entSAWActivity["hil_sawactivity"] = new EntityReference("hil_sawactivity", jobData.SAWActivity);
                entSAWActivity["hil_jobid"] = jobData.WorkOrder;
                entSAWActivity["hil_level"] = sawApprovalMatrix.GetAttributeValue<OptionSetValue>("hil_level");
                if (sawApprovalMatrix.GetAttributeValue<OptionSetValue>("hil_level").Value == 1)
                {
                    entSAWActivity["hil_isenabled"] = true;
                }
                entSAWActivity["hil_approver"] = approver;
                entSAWActivity["hil_approvalstatus"] = new OptionSetValue(1); //requested
                Guid sawActivityId = service.Create(entSAWActivity);
            }
        }

        private EntityReference FetchApprover(Entity approvalMatrix, EntityCollection sbuMappingColl)
        {
            EntityReference user = null, position; 
            if (approvalMatrix.Attributes.Contains("hil_picuser"))
            {
                user = approvalMatrix.GetAttributeValue<EntityReference>("hil_picuser");
            }
            else
            {
                position = approvalMatrix.GetAttributeValue<EntityReference>("hil_picposition");
                if (position.Name.ToUpper() == "BSH")
                {
                    if (sbuMappingColl.Entities.Count > 0)
                    {
                        user = sbuMappingColl.Entities[0].GetAttributeValue<EntityReference>("hil_branchheaduser");
                    }
                }
                else if (position.Name.ToUpper() == "NSH")
                {
                    user = sbuMappingColl.Entities[0].GetAttributeValue<EntityReference>("hil_nsh");

                }
                else if (position.Name.ToUpper() == "NPH")
                {
                    user = sbuMappingColl.Entities[0].GetAttributeValue<EntityReference>("hil_nph");
                }
            }

            return user;

        }

        private EntityCollection FetchSAWApprovals(JobData jobData, AssetData assetData)
        {
            string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='hil_fgreplacementapprovalmatrix'>
                                    <attribute name='hil_fgreplacementapprovalmatrixid'/>
                                    <attribute name='hil_warrantysubstatus'/>
                                    <attribute name='hil_warrantystatus'/>
                                    <attribute name='hil_replacementreason'/>
                                    <attribute name='hil_productsubcategory'/>
                                    <attribute name='hil_productcategory'/>
                                    <attribute name='hil_levelofapproval'/>
                                    <attribute name='hil_level'/>
                                    <attribute name='hil_callsubtype'/>
                                    <attribute name='hil_approverposition'/>
                                    <attribute name='hil_approver'/>
                                    <attribute name='hil_agingto'/>
                                    <attribute name='hil_agingfrom'/>
                                    <order attribute='hil_name' descending='false'/>
                                    <filter type='and'>
                                    <condition attribute='hil_productcategory' operator='eq' uiname='WATER PURIFIERS' uitype='product' value='{jobData.ProductCategory}'/>
                                    <condition attribute='hil_productsubcategory' operator='eq' uiname='HAVELLS Digiplus' uitype='product' value='{jobData.ProductSubCategory}'/>
                                    <condition attribute='hil_callsubtype' operator='eq' uiname='AMC Call' uitype='hil_callsubtype' value='{jobData.CallSubType}'/>
                                    <condition attribute='hil_warrantystatus' operator='eq' value='{assetData.WarrantyStatus}'/>
                                    <condition attribute='hil_warrantysubstatus' operator='eq' value='{assetData.WarrantySubStatus}'/>
                                    <filter type='or'>
                                    <condition attribute='hil_agingfrom' operator='le' value='{assetData.Aging}'/>
                                    <filter type='and'>
                                    <condition attribute='hil_agingfrom' operator='le' value='{assetData.Aging}'/>
                                    <condition attribute='hil_agingto' operator='ge' value='{assetData.Aging}'/>
                                    </filter>
                                    </filter>
                                    </filter>
                                    </entity>
                                    </fetch>";

            return service.RetrieveMultiple(new FetchExpression(fetchXml));

        }

        private Guid CreateActivity(ref JobData jobData, IOrganizationService service)
        {
            Entity newActivity = new Entity("hil_sawactivity");
            newActivity["hil_jobid"] = jobData.WorkOrder.ToEntityReference();
            newActivity["hil_jobincident"] = jobData.JobIncident.ToEntityReference();
            newActivity["hil_relatedchannelpartner"] = jobData.Franchise;
            newActivity["hil_sawcategory"] = new EntityReference("hil_serviceactionwork", new Guid("c679c5fa-7fee-ef11-9341-6045bdc64f83")); //Product Replacement 
            newActivity["hil_approvalstatus"] = new OptionSetValue(1); // Requested
            newActivity["ownerid"] = jobData.JobOwner;
            return service.Create(newActivity);
        }

        private bool DoesActivityExists(ref JobData jobData, IOrganizationService service)
        {   
            QueryExpression query = new QueryExpression("hil_sawactivity");
            query.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, jobData.WorkOrder.Id);
            query.Criteria.AddCondition("hil_jobincident", ConditionOperator.Equal, jobData.JobIncident.Id);    
            query.Criteria.AddCondition("hil_relatedchannelpartner", ConditionOperator.Equal, jobData.Franchise.Id);
            query.Criteria.AddCondition("hil_sawcategory", ConditionOperator.Equal, new Guid("c679c5fa-7fee-ef11-9341-6045bdc64f83")); // Product Replacement    
            query.Criteria.AddCondition("hil_approvalstatus", ConditionOperator.In, 1, 2, 3); //1: Requested 3: Approved   
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); // Active
            EntityCollection tempColl = service.RetrieveMultiple(query);
            if (tempColl.Entities.Count > 0)
            {
                return true;
            }
            return false;
        }

        private void UpdateJobSubStatus(JobData jobData, IOrganizationService service)
        {
            Entity updateJob = new Entity(jobData.WorkOrder.LogicalName, jobData.WorkOrder.Id);
            updateJob["msdyn_substatus"] = new EntityReference("msdyn_workordersubstatus", new Guid("")); //Create new record (Replacement in Progress)
        }

        private bool DoesOrderLineExists(JobData jobData, IOrganizationService service)
        {
            QueryExpression query = new QueryExpression("hil_inventorypurchaseorderline");
            query.Criteria.AddCondition("hil_partcode", ConditionOperator.Equal, jobData.PartCode.Id);
            query.Criteria.AddCondition("hil_jobincident", ConditionOperator.Equal, jobData.JobIncident.Id);
            query.Criteria.AddCondition("hil_workorder", ConditionOperator.Equal, jobData.WorkOrder.Id);
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
            EntityCollection tempColl = service.RetrieveMultiple(query);
            if (tempColl.Entities.Count > 0)
            {
                return true;
            }
            return false;
        }

        private Guid CreateOrderLine(JobData jobData, IOrganizationService service)
        {
            int warrantyStatus = jobData.JobIncident.Contains("hil_warrantystatus") ? jobData.JobIncident.GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value : -1;
            Entity newPurchaseOrderLine = new Entity("hil_inventorypurchaseorderline");
            newPurchaseOrderLine["hil_jobincident"] = jobData.JobIncident.ToEntityReference();
            newPurchaseOrderLine["hil_ponumber"] = new EntityReference("hil_inventorypurchaseorder", jobData.PurchaseOrder);
            newPurchaseOrderLine["hil_partcode"] = jobData.PartCode;
            newPurchaseOrderLine["hil_workorder"] = jobData.WorkOrder;
            newPurchaseOrderLine["ownerid"] = jobData.Franchise;
            if (jobData.Quantity != -1) newPurchaseOrderLine["hil_orderquantity"] = jobData.Quantity;
            //if (warrantyStatus != -1) newPurchaseOrderLine["hil_warrantystatus"] = new OptionSetValue(warrantyStatus);
            return service.Create(newPurchaseOrderLine);
        }

        private Guid DoesOrderExists(JobData jobData, IOrganizationService service)
        {
            //->// Need to check for the order type, currenty taken as emergency
            Guid purchaseOrderId = Guid.Empty;
            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_inventorypurchaseorder'>
                <attribute name='hil_inventorypurchaseorderid' />
                <filter type='and'>
                    <condition attribute='hil_jobid' operator='eq' value='{jobData.WorkOrder.Id}' />
                    <condition attribute='hil_ordertype' operator='eq' value='4' /> 
                    <condition attribute='hil_postatus' operator='ne' value='4' />
                    <condition attribute='statecode' operator='eq' value='0' />
                </filter>
                </entity>
                </fetch>";
            EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
            if (entCol.Entities.Count > 0)
            {
                purchaseOrderId = entCol.Entities[0].Id;
            }
            return purchaseOrderId;
        }

        private Guid CreateOrder(JobData jobData, IOrganizationService service)
        {
            Entity newPurchaseOrder = new Entity("hil_inventorypurchaseorder");
            newPurchaseOrder["hil_jobid"] = new EntityReference(jobData.WorkOrder.LogicalName, jobData.WorkOrder.Id);
            newPurchaseOrder["hil_productdivision"] = jobData.ProductCategory;
            newPurchaseOrder["hil_salesoffice"] = jobData.SalesOffice;
            newPurchaseOrder["hil_franchise"] = jobData.Franchise;
            newPurchaseOrder["hil_warehouse"] = jobData.Warehouse.ToEntityReference();
            newPurchaseOrder["ownerid"] = jobData.FranchiseOwner;
            newPurchaseOrder["hil_postatus"] = new OptionSetValue(1); // 
            newPurchaseOrder["hil_ordertype"] = new OptionSetValue(4);  //Product Replacements
            newPurchaseOrder["hil_brand"] = jobData.Brand;
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

        class JobData
        {
            public Entity WorkOrder { get; set; }
            public EntityReference SubStatus { get; set; }
            public EntityReference SalesOffice { get; set; }
            public EntityReference PartCode { get; set; }
            public EntityReference ProductCategory { get; set; }
            public EntityReference ProductSubCategory { get; set; }
            public EntityReference Franchise { get; set; }
            public EntityReference FranchiseOwner { get; set; }
            public Entity JobIncident { get; set; }
            public Entity Warehouse { get; set; }
            public OptionSetValue Brand { get; set; }
            public Guid PurchaseOrder { get; set; }
            public int Quantity { get; set; }
            public DateTime CreatedOn { get; set; }
            public EntityReference CustomerAssetRef { get; set; }
            public EntityReference JobOwner { get; set; }
            public Guid SAWActivity { get; set; }
            public EntityReference CallSubType { get; set; }
        }

        class AssetData
        {   
            public Entity Asset { get; set; }
            public OptionSetValue WarrantyStatus { get; set; }
            public OptionSetValue WarrantySubStatus { get; set; }
            public DateTime InvoiceDate { get; set; }
            public int Aging { get; set; }
        }
    }
}
