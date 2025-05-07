using Havells.FieldService.RMA.Helper;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.FieldService.RMA.PurchaseOrder
{
    public class CreatePO : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_owneraccount", "ownerid", "hil_warrantystatus"));
                    RetriveBookableResourceResponse retriveBookableResourceResponse = InventoryHelper.RetriveBookableResource(service, (EntityReference)entity["ownerid"]);
                    if (retriveBookableResourceResponse.FreshWareHouse != null)
                    {
                        PurchaseOrderModel purchaseOrderModel = new PurchaseOrderModel();
                        purchaseOrderModel.msdyn_vendor = (EntityReference)entity["hil_owneraccount"];
                        purchaseOrderModel.hil_potype = new EntityReference("hil_inventorytransactiontype", PurchaseOrderType.Emergency);
                        purchaseOrderModel.msdyn_requestedbyresource = retriveBookableResourceResponse.BooableResource;
                        purchaseOrderModel.msdyn_requestedbyresource = retriveBookableResourceResponse.BooableResource;
                        purchaseOrderModel.msdyn_receivetowarehouse = retriveBookableResourceResponse.FreshWareHouse;
                        purchaseOrderModel.msdyn_purchaseorderdate = DateTime.Now;
                        purchaseOrderModel.msdyn_orderedby = (EntityReference)entity["ownerid"];
                        purchaseOrderModel.IV_Usages = ((OptionSetValue)entity["hil_warrantystatus"]).Value == 1 ? "EM" : "S";
                        purchaseOrderModel.ownerid = (EntityReference)entity["ownerid"];
                        purchaseOrderModel.msdyn_workorder = entity.ToEntityReference();
                        createPO(service, purchaseOrderModel);
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException("Error in Create PO Plugin " + ex.Message);
                }
            }
        }
        public static void createPO(IOrganizationService service, PurchaseOrderModel purchaseOrderTypeModel)
        {
            Guid POID = Guid.Empty;
            EntityCollection JobPrdColl = null;
            try
            {
                QueryExpression query = new QueryExpression("msdyn_purchaseorder");
                query.ColumnSet = new ColumnSet(false);
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, purchaseOrderTypeModel.msdyn_workorder.Id);
                query.Criteria.AddCondition("hil_sapsalesorderno", ConditionOperator.Null);
                query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                EntityCollection _entitys = service.RetrieveMultiple(query);
                if (_entitys.Entities.Count != 0)
                {
                    POID = _entitys.Entities[0].Id;
                    JobPrdColl = RetrivePOProductToCreate(service, purchaseOrderTypeModel.msdyn_workorder, POID);
                }
                else
                {
                    JobPrdColl = RetrivePOProductToCreate(service, purchaseOrderTypeModel.msdyn_workorder, POID);
                    if (JobPrdColl.Entities.Count > 0)
                    {
                        Entity entity = new Entity("msdyn_purchaseorder");
                        entity["hil_potype"] = purchaseOrderTypeModel.hil_potype;
                        entity["msdyn_vendor"] = purchaseOrderTypeModel.msdyn_vendor;
                        entity["msdyn_receivetowarehouse"] = purchaseOrderTypeModel.msdyn_receivetowarehouse;
                        entity["msdyn_purchaseorderdate"] = purchaseOrderTypeModel.msdyn_purchaseorderdate;
                        entity["msdyn_requestedbyresource"] = purchaseOrderTypeModel.msdyn_requestedbyresource;
                        entity["ownerid"] = purchaseOrderTypeModel.ownerid;
                        entity["msdyn_orderedby"] = purchaseOrderTypeModel.msdyn_orderedby;
                        entity["msdyn_workorder"] = purchaseOrderTypeModel.msdyn_workorder;
                        entity["hil_ivusage"] = purchaseOrderTypeModel.IV_Usages;
                        POID = service.Create(entity);
                    }
                }
                if (POID != Guid.Empty)
                    CreatePOProducts(service, new EntityReference("msdyn_purchaseorder", POID), JobPrdColl, purchaseOrderTypeModel);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        public static void CreatePOProducts(IOrganizationService service, EntityReference PORef, EntityCollection JobPrdColl, PurchaseOrderModel purchaseOrderTypeModel)
        {
            foreach (Entity item in JobPrdColl.Entities)
            {
                Entity _POProduct = new Entity("msdyn_purchaseorderproduct");
                _POProduct["msdyn_product"] = item["hil_replacedpart"];
                _POProduct["msdyn_quantity"] = item["msdyn_quantity"];
                _POProduct["msdyn_unit"] = (EntityReference)item.GetAttributeValue<AliasedValue>("prod.defaultuomid").Value;
                _POProduct["msdyn_purchaseorder"] = PORef;// new EntityReference("msdyn_purchaseorder", POID);
                _POProduct["msdyn_associatetowarehouse"] = purchaseOrderTypeModel.msdyn_receivetowarehouse;
                _POProduct["msdyn_associatetoworkorder"] = purchaseOrderTypeModel.msdyn_workorder;
                _POProduct["hil_associatetoworkorderproduct"] = item.ToEntityReference();
                service.Create(_POProduct);
            }
        }
        public static EntityCollection RetrivePOProductToCreate(IOrganizationService service, EntityReference _jobID, Guid POID)
        {
            EntityCollection _entityCollection = new EntityCollection();
            try
            {
                {
                    string fetch = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                                  <entity name=""msdyn_workorderproduct"">
                                    <attribute name=""createdon"" />
                                    <attribute name=""msdyn_product"" />
                                    <attribute name=""msdyn_linestatus"" />
                                    <attribute name=""hil_replacedpart"" />
                                    <attribute name=""msdyn_quantity"" />
                                    <attribute name=""msdyn_workorderproductid"" />
                                    <order attribute=""msdyn_product"" descending=""false"" />
                                    <filter type=""and"">
                                      <condition attribute=""msdyn_workorder"" operator=""eq"" value=""{_jobID.Id}"" />
                                      <condition attribute=""hil_replacedpart"" operator=""not-null"" />
                                      <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                                      <condition attribute=""hil_availabilitystatus"" operator=""ne"" value=""1"" />
                                    </filter>
                                    <link-entity name=""product"" from=""productid"" to=""hil_replacedpart"" visible=""false"" link-type=""outer"" alias=""prod"">
                                      <attribute name=""defaultuomid"" />
                                    </link-entity>
                                  </entity>
                                </fetch>";

                    EntityCollection JobPrdColl = service.RetrieveMultiple(new FetchExpression(fetch));
                    foreach (Entity item in JobPrdColl.Entities)
                    {
                        QueryExpression query = new QueryExpression("msdyn_purchaseorderproduct");
                        query.ColumnSet = new ColumnSet(false);
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_associatetoworkorderproduct", ConditionOperator.Equal, item.Id);
                        query.Criteria.AddCondition("msdyn_associatetoworkorder", ConditionOperator.Equal, _jobID.Id);
                        query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                        EntityCollection _entitysPrd = service.RetrieveMultiple(query);
                        if (_entitysPrd.Entities.Count == 0)
                        {
                            _entityCollection.Entities.Add(item);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            return _entityCollection;
        }
    }
}
