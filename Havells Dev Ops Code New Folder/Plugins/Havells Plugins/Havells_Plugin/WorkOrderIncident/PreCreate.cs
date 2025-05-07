using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.WorkOrderIncident
{
    public class PreCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            try
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                msdyn_workorderincident _entWOInc = entity.ToEntity<msdyn_workorderincident>();
                tracingService.Trace("1");
                if (_entWOInc.msdyn_CustomerAsset == null || _entWOInc.hil_ProductCategory == null || _entWOInc.hil_ProductSubCategory == null)
                {
                    throw new InvalidPluginExecutionException("Customer Assest/Product Category/Product Sub-Category/Model Cannot be Empty");
                }
                msdyn_customerasset cus_Assest = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, _entWOInc.msdyn_CustomerAsset.Id, new ColumnSet(new string[] { "msdyn_product", "hil_productcategory", "hil_invoicedate", "hil_invoiceno", "hil_productsubcategorymapping" }));
                tracingService.Trace("2");
                if (cus_Assest.hil_ProductCategory == null || cus_Assest.msdyn_Product == null || cus_Assest.hil_productsubcategorymapping == null
                    || cus_Assest.hil_ProductCategory.Id != _entWOInc.hil_ProductCategory.Id)
                {
                    throw new InvalidPluginExecutionException("The Customer Asset category combination should match with Job Incident.");
                }
                tracingService.Trace("3");
                hil_stagingdivisonmaterialgroupmapping sdmMapping = (hil_stagingdivisonmaterialgroupmapping)service.Retrieve(
                            hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, cus_Assest.GetAttributeValue<EntityReference>("hil_productsubcategorymapping").Id, new ColumnSet(new string[] { "hil_productsubcategorymg" }));
                tracingService.Trace("4");
                if (sdmMapping.hil_ProductSubCategoryMG == null
                    || sdmMapping.hil_ProductSubCategoryMG.Id != _entWOInc.hil_ProductSubCategory.Id)
                {
                    throw new InvalidPluginExecutionException("The Customer Asset category combination should match with Job Incident. Please contact to Administrator");
                }
                tracingService.Trace("5");
                Product prod_Category = (Product)service.Retrieve(Product.EntityLogicalName, _entWOInc.hil_ProductSubCategory.Id, new ColumnSet(new string[] { "hil_isserialized" }));
                if (prod_Category.hil_IsSerialized != null && prod_Category.hil_IsSerialized.Value == 1
                    && (_entWOInc.hil_Quantity == null || _entWOInc.hil_Quantity != 1))
                {
                    throw new InvalidPluginExecutionException("Quantity should be equals to 1 for serialized Product");
                }
                _entWOInc["hil_modelcode"] = new EntityReference(Product.EntityLogicalName, cus_Assest.msdyn_Product.Id);
                tracingService.Trace("6");
                CheckWorkOrderIncident(service, _entWOInc, (prod_Category.hil_IsSerialized != null ? prod_Category.hil_IsSerialized.Value : 2));
                tracingService.Trace("7");
                CheckIfAnyOtherJobPendingForThisAsset(service, _entWOInc);
                AddQuantityInJob(service, _entWOInc);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrderIncident.PreCreate" + ex.Message);
            }
        }
        public static void CheckWorkOrderIncident(IOrganizationService service, msdyn_workorderincident _entWOInc, int isSerialized)
        {
            //if (_entWOInc.msdyn_WorkOrder != null && _entWOInc.hil_observation != null)
            //{
            //    QueryByAttribute Query = new QueryByAttribute(msdyn_workorderincident.EntityLogicalName);
            //    Query.ColumnSet = new ColumnSet(new string[] { "hil_quantity", "msdyn_customerasset" });
            //    Query.AddAttributeValue("msdyn_workorder", _entWOInc.msdyn_WorkOrder.Id);
            //    Query.AddAttributeValue("hil_observation", _entWOInc.hil_observation.Id);
            //    Query.AddAttributeValue("msdyn_customerasset", _entWOInc.msdyn_CustomerAsset.Id);
            //    Query.AddAttributeValue("statecode", 0);
            //    EntityCollection Found = service.RetrieveMultiple(Query);
            //
            //    if (Found.Entities.Count > 0)
            //    {
            //        foreach (msdyn_workorderincident WoInc in Found.Entities)
            //        {
            //            if (WoInc.hil_observation.Id == _entWOInc.hil_observation.Id
            //                && WoInc.msdyn_CustomerAsset.Id == _entWOInc.msdyn_CustomerAsset.Id)
            //            {
            //                throw new InvalidPluginExecutionException("xxxxxxxxxxxxxxxxxxxxxxxxxxxxx INCIDENT WITH THIS OBSERVATION ALREADY EXISTS xxxxxxxxxxxxxxxxxxxxxxxxxxxxx");
            //            }
            //        }
            //    }
            //}
            msdyn_workorder WO = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, _entWOInc.msdyn_WorkOrder.Id, new ColumnSet(new string[] { "hil_quantity" }));

            if (WO.hil_quantity == null)
            {
                throw new InvalidPluginExecutionException("Job Quantity cannot be Empty");
            }

            QueryExpression WoIncQuery = new QueryExpression()
            {
                EntityName = msdyn_workorderincident.EntityLogicalName,
                ColumnSet = new ColumnSet(new string[] { "hil_quantity", "msdyn_customerasset", "hil_observation" }),
                Criteria =
                            {
                                Conditions =
                                {
                                    new ConditionExpression("msdyn_workorder",ConditionOperator.Equal,_entWOInc.msdyn_WorkOrder.Id),
                                    new ConditionExpression("statecode",ConditionOperator.Equal,0)
                                }
                            },
                NoLock = true
            };

            EntityCollection WoIncColl = service.RetrieveMultiple(WoIncQuery);
            HashSet<Guid> AssestCollection = new HashSet<Guid>();

            AssestCollection.Add(_entWOInc.msdyn_CustomerAsset.Id); //Adding Context record Assest to collection as context record is not fetch in pre operation
            int QuantitySum = _entWOInc.hil_Quantity != null ? _entWOInc.hil_Quantity.Value : 0; //Adding context record quantity

            if (WoIncColl.Entities != null && WoIncColl.Entities.Count > 0)
            {
                foreach (msdyn_workorderincident WoInc in WoIncColl.Entities)
                {
                    if (WoInc.hil_observation.Id == _entWOInc.hil_observation.Id
                            && WoInc.msdyn_CustomerAsset.Id == _entWOInc.msdyn_CustomerAsset.Id)
                    {
                        throw new InvalidPluginExecutionException("xxxxxxxxxxxxxxxxxxxxxxxxxxxxx INCIDENT WITH THIS OBSERVATION ALREADY EXISTS xxxxxxxxxxxxxxxxxxxxxxxxxxxxx");
                    }
                    if (WoInc.msdyn_CustomerAsset != null)
                    {
                        AssestCollection.Add(WoInc.msdyn_CustomerAsset.Id);
                    }
                    if (WoInc.hil_Quantity != null)
                    {
                        QuantitySum += WoInc.hil_Quantity.Value;
                    }
                }
            }
            if (isSerialized == 1 && AssestCollection.Count > WO.hil_quantity.Value)
            {
                throw new InvalidPluginExecutionException("For Serialized Product Job Quantity must be Equals to Total No. of Assests.");
            }
            else if (isSerialized == 2 && AssestCollection.Count > WO.hil_quantity.Value)
            {
                throw new InvalidPluginExecutionException("For Serialized Product Job Quantity must be Equals to or less than Total No. of Assests.");
            }
            else if (isSerialized == 2 && QuantitySum > WO.hil_quantity.Value)
            {
                throw new InvalidPluginExecutionException("Total Incident Quantity must be equals to Job Quantity");
            }
            else
            {
                return;
            }
        }
        public static void CheckIfAnyOtherJobPendingForThisAsset(IOrganizationService service, msdyn_workorderincident _entWOInc)
        {
            if (_entWOInc.msdyn_CustomerAsset != null)
            {
                QueryExpression Query = new QueryExpression(msdyn_workorderincident.EntityLogicalName);
                Query.ColumnSet = new ColumnSet(false);
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression("msdyn_customerasset", ConditionOperator.Equal, _entWOInc.msdyn_CustomerAsset.Id));
                EntityCollection Found = service.RetrieveMultiple(Query);
                if (Found.Entities.Count > 0)
                {
                    foreach (msdyn_workorderincident iCase in Found.Entities)
                    {
                        if (iCase.msdyn_WorkOrder != null && _entWOInc.msdyn_WorkOrder != null)
                        {
                            if (_entWOInc.msdyn_WorkOrder != iCase.msdyn_WorkOrder)
                            {
                                msdyn_workorder iJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, iCase.msdyn_WorkOrder.Id, new ColumnSet("msdyn_substatus"));
                                if (iJob.msdyn_SubStatus.Name != "Work Done" || iJob.msdyn_SubStatus.Name != "Canceled" ||
                                    iJob.msdyn_SubStatus.Name != "Closed")
                                {
                                    throw new InvalidPluginExecutionException("More than One Jobs are Currently Active for this Customer Asset");
                                }
                            }
                        }
                    }
                }
            }
        }
        public static void AddQuantityInJob(IOrganizationService service, msdyn_workorderincident _entWOInc)
        {
            Int32 iQuantity = 0;
            if (_entWOInc.msdyn_WorkOrder != null)
            {
                msdyn_workorder iJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, _entWOInc.msdyn_WorkOrder.Id, new ColumnSet("hil_incidentquantity"));
                msdyn_workorder iUpdateJob = new msdyn_workorder();
                iUpdateJob.Id = iJob.Id;
                if (iJob.Attributes.Contains("hil_incidentquantity"))
                {
                    iQuantity = (Int32)iJob["hil_incidentquantity"];
                    if (_entWOInc.hil_Quantity != null)
                    {
                        iQuantity = iQuantity + Convert.ToInt32(_entWOInc.hil_Quantity);
                        iUpdateJob["hil_incidentquantity"] = iQuantity;
                    }
                }
                else
                {
                    iUpdateJob["hil_incidentquantity"] = Convert.ToInt32(_entWOInc.hil_Quantity);
                }
                service.Update(iUpdateJob);
            }
        }
    }
}
