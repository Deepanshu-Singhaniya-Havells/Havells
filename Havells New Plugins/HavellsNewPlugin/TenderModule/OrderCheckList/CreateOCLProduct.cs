using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.OrderCheckList
{
    public class CreateOCLProduct : IPlugin
    {
        public static ITracingService tracingService = null;
        public static EntityReference Department = null;
        public void Execute(IServiceProvider serviceProvider)
        {

            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            tracingService.Trace("Plugin Started...");
            try
            {
                Guid logdinUser = context.UserId;
                getUserMapping(service, logdinUser);
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    Entity ocl = new Entity(entity.LogicalName);
                    ocl["hil_department"] = Department;
                    ocl.Id = entity.Id;
                    service.Update(ocl);
                    if (entity.Contains("hil_tenderno"))
                    {
                        EntityCollection _tenderProductColl = retriveAllTenderProduct(service, entity);
                        tracingService.Trace("1");
                        tracingService.Trace("_tenderProductColl " + _tenderProductColl.Entities.Count);
                        EntityCollection _OCLColl = retriveExistingOCL(service, entity);
                        tracingService.Trace("_OCLColl " + _OCLColl.Entities.Count);
                        if (_OCLColl.Entities.Count > 1)
                        {
                            tracingService.Trace("Multiple order CheckList");
                            foreach (Entity tenderProduct in _tenderProductColl.Entities)
                            {
                                EntityCollection _OCLCollforTenderProd = retriveAllOCLAginstTenderProduct(service, tenderProduct.ToEntityReference());
                                if (_OCLCollforTenderProd.Entities.Count > 0)
                                {
                                    tracingService.Trace("multiple Product");
                                    Decimal totalPOQut = getTotalPOQuantity(service, tenderProduct.Id, tracingService);
                                    decimal remainingPOQty = tenderProduct.GetAttributeValue<decimal>("hil_quantity") - totalPOQut;
                                    if (remainingPOQty > 0)
                                    {
                                        createOCLProduct(service, tenderProduct, entity, remainingPOQty);
                                    }
                                }
                                else if (_OCLCollforTenderProd.Entities.Count == 0)
                                {
                                    createOCLProduct(service, tenderProduct, entity, null);
                                }
                            }
                        }
                        else if (_OCLColl.Entities.Count == 1)
                        {
                            foreach (Entity tenderProduct in _tenderProductColl.Entities)
                            {

                                createOCLProduct(service, tenderProduct, entity, null);
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error : " + ex.Message);
            }
        }
        void getUserMapping(IOrganizationService service, Guid logdinUser)
        {
            try
            {
                QueryExpression query = new QueryExpression("hil_userbranchmapping");
                query.ColumnSet = new ColumnSet("hil_department");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition(new ConditionExpression("hil_user", ConditionOperator.Equal, logdinUser));
                query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                EntityCollection userMapping = service.RetrieveMultiple(query);
                if (userMapping.Entities.Count > 0)
                {
                    Department = userMapping[0].GetAttributeValue<EntityReference>("hil_department");
                }
                else
                {
                    throw new InvalidPluginExecutionException("User Configuration is not Created : ");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        EntityCollection retriveExistingOCL(IOrganizationService service, Entity entity)
        {
            try
            {
                Guid tenderID = entity.GetAttributeValue<EntityReference>("hil_tenderno").Id;
                QueryExpression query = new QueryExpression("hil_orderchecklist");
                query.ColumnSet = new ColumnSet(false);
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition(new ConditionExpression("hil_tenderno", ConditionOperator.Equal, tenderID));
                query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0)); // added for active OCL
                return service.RetrieveMultiple(query);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error in retriveExistingOCL : " + ex.Message);
            }
        }
        EntityCollection retriveAllTenderProduct(IOrganizationService service, Entity entity)
        {
            try
            {
                Guid tenderID = entity.GetAttributeValue<EntityReference>("hil_tenderno").Id;
                QueryExpression query = new QueryExpression("hil_tenderproduct");
                query.ColumnSet = new ColumnSet(true);
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition(new ConditionExpression("hil_tenderid", ConditionOperator.Equal, tenderID));
                query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                return service.RetrieveMultiple(query);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error in retriveAllTenderProduct : " + ex.Message);
            }
        }
        EntityCollection retriveAllOCLAginstTenderProduct(IOrganizationService service, EntityReference tenderProduct)
        {
            try
            {
                QueryExpression query = new QueryExpression("hil_orderchecklistproduct");
                query.ColumnSet = new ColumnSet(false);
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition(new ConditionExpression("hil_tenderproductid", ConditionOperator.Equal, tenderProduct.Id));
                query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                return service.RetrieveMultiple(query);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error in retriveAllOCLAginstTenderProduct : " + ex.Message);
            }
        }
        public static Decimal getTotalPOQuantity(IOrganizationService service, Guid tenderPrdID, ITracingService tracingService)
        {
            try
            {

                tracingService.Trace("getTotalPOQuantity start...");
                string _OCLProductfetch = @"<fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                            <entity name='hil_orderchecklistproduct'>
                            <attribute name='hil_tenderproductid' alias='tenderPrdNo' groupby='true' />
                            <attribute name='hil_poqty' alias='amount' aggregate='sum' />
                            <filter type='and'>
                                <condition attribute='hil_tenderproductid' operator='eq' value='" + tenderPrdID + @"' />
                                <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            </entity>
                        </fetch>";
                tracingService.Trace("Fetch " + _OCLProductfetch);
                EntityCollection _tenderproductColl = service.RetrieveMultiple(new FetchExpression(_OCLProductfetch));
                tracingService.Trace("_tenderproductColl.count" + _tenderproductColl.Entities.Count);
                if (_tenderproductColl.Entities.Count > 0)
                {
                    Decimal totalPOQuantity = (Decimal)((AliasedValue)_tenderproductColl.Entities[0]["amount"]).Value;
                    //throw new InvalidPluginExecutionException("totalPOQuantity" + totalPOQuantity);
                    return totalPOQuantity;
                }
                else return 0;

            }
            catch (Exception ex)
            {

                throw new InvalidPluginExecutionException("Error in getTotalPOQuantity : " + ex.Message);
            }
        }
        void createOCLProduct(IOrganizationService service, Entity tenderProduct, Entity entity, decimal? poQty)
        {
            try
            {
                Entity _oclProd = new Entity("hil_orderchecklistproduct");
                if (tenderProduct.Contains("hil_product"))
                    _oclProd["hil_product"] = tenderProduct["hil_product"];// new EntityReference();

                _oclProd["hil_department"] = Department;

                tracingService.Trace("product: " + ((EntityReference)tenderProduct["hil_product"]).Name);

                if (tenderProduct.Contains("hil_productdescription"))
                    _oclProd["hil_productdescription"] = tenderProduct["hil_productdescription"];//"";

                if (tenderProduct.Contains("hil_lprsmtr"))
                    _oclProd["hil_lprsmtr"] = tenderProduct["hil_lprsmtr"];//new Money();

                if (tenderProduct.Contains("hil_quantity"))
                    _oclProd["hil_quantity"] = tenderProduct["hil_quantity"];//33.3;

                if (poQty == null)
                    _oclProd["hil_poqty"] = tenderProduct["hil_quantity"];//33.3;
                else
                    _oclProd["hil_poqty"] = poQty;//33.3;

                tracingService.Trace("sss tenderProduct.Id " + tenderProduct.Id);
                QueryExpression query = new QueryExpression("hil_orderchecklistproduct");
                query.ColumnSet = new ColumnSet("hil_porate");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition(new ConditionExpression("hil_tenderproductid", ConditionOperator.Equal, tenderProduct.Id));
                query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                query.AddOrder("createdon", OrderType.Ascending);
                EntityCollection userMapping = service.RetrieveMultiple(query);
                if (userMapping.Entities.Count > 0)
                {
                    tracingService.Trace("sssaqqq " + userMapping[0].Id);
                    _oclProd["hil_porate"] = userMapping[0]["hil_porate"];
                }
                else
                {
                    tracingService.Trace("sssaaa");
                    _oclProd["hil_porate"] = tenderProduct["hil_basicpriceinrsmtr"];
                }

                tracingService.Trace("sssssssss");
                _oclProd["hil_name"] = entity["hil_name"] + CreateOCLPrdAutoNo(service, entity.ToEntityReference());

                _oclProd["hil_tenderproductid"] = tenderProduct.ToEntityReference();

                _oclProd["hil_orderchecklistid"] = entity.ToEntityReference();

                if (tenderProduct.Contains("hil_selectproduct"))
                    _oclProd["hil_selectproduct"] = tenderProduct["hil_selectproduct"];//true;

                if (tenderProduct.Contains("hil_hodiscper"))
                    _oclProd["hil_hodiscper"] = tenderProduct["hil_hodiscper"];//true;

                if (tenderProduct.Contains("hil_marginaddedonhoprice"))
                    _oclProd["hil_marginaddedonhoprice"] = tenderProduct["hil_marginaddedonhoprice"];//true;

                if (tenderProduct.Contains("hil_hopricespecialconstructions"))
                    _oclProd["hil_hopricespecialconstructions"] = tenderProduct["hil_hopricespecialconstructions"];//true;

                if (tenderProduct.Contains("hil_basicpriceinrsmtr"))
                    _oclProd["hil_finaloffervalue"] = tenderProduct["hil_basicpriceinrsmtr"];//new Money();

                bool inspection = entity.Contains("hil_inspection") ? entity.GetAttributeValue<bool>("hil_inspection") : false;

                decimal QtyTolerance = entity.GetAttributeValue<decimal>("hil_overall");

                _oclProd["hil_tolerancelowerlimit"] = QtyTolerance;
                _oclProd["hil_toleranceupperlimit"] = QtyTolerance;

                if (inspection == true)
                {
                    QueryExpression QueryInspectionType = new QueryExpression("hil_inspectiontype");
                    QueryInspectionType.ColumnSet = new ColumnSet("hil_name");
                    QueryInspectionType.NoLock = true;
                    QueryInspectionType.Criteria = new FilterExpression(LogicalOperator.And);
                    QueryInspectionType.Criteria.AddCondition("hil_code", ConditionOperator.Equal, "01");
                    QueryInspectionType.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                    EntityCollection entCol2 = service.RetrieveMultiple(QueryInspectionType);
                    if (entCol2.Entities.Count > 0)
                    {
                        _oclProd["hil_inspectiontype"] = entCol2[0].ToEntityReference();
                    }
                }
                else
                {
                    QueryExpression QueryInspectionType = new QueryExpression("hil_inspectiontype");
                    QueryInspectionType.ColumnSet = new ColumnSet("hil_name");
                    QueryInspectionType.NoLock = true;
                    QueryInspectionType.Criteria = new FilterExpression(LogicalOperator.And);
                    QueryInspectionType.Criteria.AddCondition("hil_code", ConditionOperator.Equal, "02");
                    QueryInspectionType.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                    EntityCollection entCol2 = service.RetrieveMultiple(QueryInspectionType);
                    if (entCol2.Entities.Count > 0)
                    {
                        _oclProd["hil_inspectiontype"] = entCol2[0].ToEntityReference();
                    }
                }
                if (entity.Contains("hil_despatchpoint"))
                    _oclProd["hil_plantcode"] = entity["hil_despatchpoint"];//new EntityReference();
                else
                    throw new InvalidPluginExecutionException("Dispatch point is Mandatory..");

                service.Create(_oclProd);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error in Product Creation : " + ex.Message);
            }

        }

        string CreateOCLPrdAutoNo(IOrganizationService service, EntityReference orderCheckListId)
        {
            try
            {
                QueryExpression query = new QueryExpression("hil_orderchecklistproduct");
                query.ColumnSet = new ColumnSet("hil_name");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition(new ConditionExpression("hil_orderchecklistid", ConditionOperator.Equal, orderCheckListId.Id));
                query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                query.Orders.Add(new OrderExpression("createdon", OrderType.Descending));
                query.TopCount = 1;
                EntityCollection entColl = service.RetrieveMultiple(query);
                tracingService.Trace("2");
                int count = 1;
                if (entColl.Entities.Count > 0)
                {
                    string _lastTend = entColl[0].GetAttributeValue<string>("hil_name");
                    string[] number = _lastTend.Split('_');
                    count = int.Parse(number[1]);
                    count++;
                }
                String autoNumber = "_" + count.ToString().PadLeft(3, '0');
                return autoNumber;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error in CreateOCLPrdAutoNo : " + ex.Message);
            }
        }
    }
}
