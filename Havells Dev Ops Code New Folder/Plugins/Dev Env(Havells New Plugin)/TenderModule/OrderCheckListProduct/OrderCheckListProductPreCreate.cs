using System;
using HavellsNewPlugin.TenderModule.OrderCheckList;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.OrderCheckListProduct
{
    public class OrderCheckListProductPreCreate : IPlugin
    {
        public static ITracingService tracingService = null;
        bool inspection = false;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "hil_orderchecklistproduct"
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("0");

                    if (entity.Contains("hil_orderchecklistid"))
                    {
                        tracingService.Trace("1");
                        EntityReference ocl = entity.GetAttributeValue<EntityReference>("hil_orderchecklistid");
                        QueryExpression query = new QueryExpression("hil_orderchecklistproduct");
                        query.ColumnSet = new ColumnSet("hil_name");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition(new ConditionExpression("hil_orderchecklistid", ConditionOperator.Equal,
                            entity.GetAttributeValue<EntityReference>("hil_orderchecklistid").Id));
                        query.Orders.Add(new OrderExpression("createdon", OrderType.Descending));
                        query.TopCount = 1;
                        EntityCollection entColl = service.RetrieveMultiple(query);
                        tracingService.Trace("2ccddd " + entColl.Entities.Count);

                        int count = 1;
                        if (entColl.Entities.Count > 0)
                        {
                            string _lastTend = entColl[0].GetAttributeValue<string>("hil_name");
                            tracingService.Trace("_lastTend" + _lastTend);
                            string[] number = _lastTend.Split('_');
                            count = int.Parse(number[1]);
                            count++;


                        }
                        tracingService.Trace("_lastTend");
                        string _ocl = service.Retrieve("hil_orderchecklist", entity.GetAttributeValue<EntityReference>("hil_orderchecklistid").Id,
                            new ColumnSet("hil_name")).GetAttributeValue<string>("hil_name");
                        entity["hil_name"] = _ocl + "_" + count.ToString().PadLeft(3, '0');
                        tracingService.Trace("2 ss " + _ocl);
                        Entity orderCheckListEntity = service.Retrieve(ocl.LogicalName, ocl.Id, new ColumnSet("hil_inspection", "hil_tenderno", "hil_department"));
                        if (!orderCheckListEntity.Contains("hil_tenderno"))
                        {
                            tracingService.Trace("With out Tender");
                            inspection = orderCheckListEntity.GetAttributeValue<bool>("hil_inspection");

                            tracingService.Trace("1");

                            if (inspection == true)
                            {
                                QueryExpression QueryInspectionType = new QueryExpression("hil_inspectiontype");
                                QueryInspectionType.ColumnSet = new ColumnSet("hil_name");
                                QueryInspectionType.NoLock = true;
                                QueryInspectionType.Criteria = new FilterExpression(LogicalOperator.And);
                                QueryInspectionType.Criteria.AddCondition("hil_code", ConditionOperator.Equal, "01");
                                EntityCollection entCol2 = service.RetrieveMultiple(QueryInspectionType);
                                if (entCol2.Entities.Count > 0)
                                {
                                    entity["hil_inspectiontype"] = entCol2[0].ToEntityReference();
                                }
                            }
                            else
                            {
                                QueryExpression QueryInspectionType = new QueryExpression("hil_inspectiontype");
                                QueryInspectionType.ColumnSet = new ColumnSet("hil_name");
                                QueryInspectionType.NoLock = true;
                                QueryInspectionType.Criteria = new FilterExpression(LogicalOperator.And);
                                QueryInspectionType.Criteria.AddCondition("hil_code", ConditionOperator.Equal, "02");
                                EntityCollection entCol2 = service.RetrieveMultiple(QueryInspectionType);
                                if (entCol2.Entities.Count > 0)
                                {
                                    entity["hil_inspectiontype"] = entCol2[0].ToEntityReference();
                                }
                            }
                            tracingService.Trace("4");
                        }
                        else if (orderCheckListEntity.Contains("hil_tenderno"))
                        {
                            if (entity.Contains("hil_tenderproductid"))
                            {
                                tracingService.Trace("POValidation(service, entity, tracingService);");

                                POValidation(service, entity, tracingService);
                            }
                        }

                        entity["hil_department"] = orderCheckListEntity["hil_department"];
                        if (entity.Contains("hil_product"))
                        {
                            Guid _productId = entity.GetAttributeValue<EntityReference>("hil_product").Id;
                            Entity _prodEnt = service.Retrieve("product", _productId, new ColumnSet("description", "hil_amount"));
                            if (_prodEnt.Contains("description"))
                            {
                                entity["hil_productdescription"] = _prodEnt.GetAttributeValue<string>("description");
                            }
                            //if (_prodEnt.Contains("hil_amount"))
                            //{
                            //    entity["hil_lprsmtr"] = new Money(_prodEnt.GetAttributeValue<Money>("hil_amount").Value / 1000);
                            //}

                            QueryExpression qryPrdLstAmt = new QueryExpression("productpricelevel");
                            qryPrdLstAmt.ColumnSet = new ColumnSet("uomid", "amount", "pricelevelid");
                            qryPrdLstAmt.NoLock = true;
                            qryPrdLstAmt.Criteria = new FilterExpression(LogicalOperator.And);
                            qryPrdLstAmt.Criteria.AddCondition("productid", ConditionOperator.Equal, _productId);
                            EntityCollection entColPrdLstAmt = service.RetrieveMultiple(qryPrdLstAmt);


                            if (entColPrdLstAmt.Entities.Count > 0)
                            {
                                if (entColPrdLstAmt.Entities[0].Contains("amount"))
                                {
                                    Guid PriceListId = entColPrdLstAmt.Entities[0].GetAttributeValue<EntityReference>("pricelevelid").Id;
                                    entity["hil_pricelist"] = new EntityReference("pricelevel", PriceListId);
                                    tracingService.Trace("amount " + entColPrdLstAmt.Entities[0].GetAttributeValue<Money>("amount").Value);

                                    if (entColPrdLstAmt.Entities[0].Contains("uomid"))
                                    {
                                        tracingService.Trace("uomid " + entColPrdLstAmt.Entities[0].GetAttributeValue<EntityReference>("uomid").Id);
                                        Guid unitId = entColPrdLstAmt.Entities[0].GetAttributeValue<EntityReference>("uomid").Id;
                                        tracingService.Trace("5 ");
                                        entity["hil_unit"] = new EntityReference("uom", unitId);
                                        tracingService.Trace("6 ");
                                        entity["hil_lprsmtr"] = new Money(entColPrdLstAmt.Entities[0].GetAttributeValue<Money>("amount").Value);

                                    }
                                }
                            }
                            tracingService.Trace("end");
                        }
                    }
                    if (entity.Contains("hil_product"))
                    {
                        EntityReference hil_hsncode;
                        decimal taxValue = getHSNValueBasedOnProduct(service, (EntityReference)entity["hil_product"], out hil_hsncode);
                        entity["hil_tax"] = taxValue;
                        entity["hil_hsncode"] = hil_hsncode;
                    }
                    else if (entity.Contains("hil_hsncode"))
                    {
                        decimal taxValue = OrderCheckListProductPreCreate.getHSNValueBasedOnHSN(service, entity.GetAttributeValue<EntityReference>("hil_hsncode"));
                        entity["hil_tax"] = taxValue;
                    }
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error " + ex.Message);
                throw new InvalidPluginExecutionException("OrderCheckListProductPreCreate Error " + ex.Message);
            }
        }
        static void getPORate(IOrganizationService service, EntityReference tendPrd, Entity entity, ITracingService tracingService)
        {
            QueryExpression query = new QueryExpression("hil_orderchecklistproduct");
            query.ColumnSet = new ColumnSet("hil_porate");
            query.Criteria = new FilterExpression(LogicalOperator.And);
            query.Criteria.AddCondition(new ConditionExpression("hil_tenderproductid", ConditionOperator.Equal, tendPrd.Id));
            query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 0));
            query.AddOrder("createdon", OrderType.Ascending);
            EntityCollection userMapping = service.RetrieveMultiple(query);
            if (userMapping.Entities.Count > 0)
            {
                decimal porate = userMapping[0].GetAttributeValue<Money>("hil_porate").Value;
                tracingService.Trace("porate " + porate);
                if (entity.Contains("hil_porate"))
                {
                    tracingService.Trace("porate 1");
                    decimal currentPORate = entity.GetAttributeValue<Money>("hil_porate").Value;
                    tracingService.Trace("currentPORate " + currentPORate);

                    if (porate != currentPORate)
                    {
                        throw new InvalidPluginExecutionException("PO Rate should be " + String.Format("{0:0.00}", porate) + " in this Item");
                    }
                }

            }
            else
            {

            }
        }
        public static void POValidation(IOrganizationService service, Entity entity, ITracingService tracingService)
        {
            if (entity.Contains("hil_tenderproductid"))
            {
                tracingService.Trace("With Tender");
                Guid TenderProduct = entity.GetAttributeValue<EntityReference>("hil_tenderproductid").Id;
                tracingService.Trace("TenderProduct  " + TenderProduct.ToString());
                decimal existingPOQty = OrderCheckList.CreateOCLProduct.getTotalPOQuantity(service, TenderProduct, tracingService);
                tracingService.Trace("existingPOQty " + existingPOQty);

                decimal currentPOQty = entity.GetAttributeValue<decimal>("hil_poqty");
                tracingService.Trace("currentPOQty " + currentPOQty);
                decimal totalQty = entity.GetAttributeValue<decimal>("hil_quantity");
                tracingService.Trace("totalQty " + totalQty);
                decimal remainingPOQty = existingPOQty - currentPOQty;
                remainingPOQty = totalQty - remainingPOQty;
                tracingService.Trace("totalPOQty " + totalQty);

                if ((totalQty - existingPOQty) < 0)
                {
                    //throw new InvalidPluginExecutionException("Quantity " + String.Format("{0:0.00}", currentPOQty) + "is not avaliable. Please use " + String.Format("{0:0.00}", remainingPOQty) + " quantity.");
                }
                getPORate(service, entity.GetAttributeValue<EntityReference>("hil_tenderproductid"), entity, tracingService);
            }

        }
        public static decimal getHSNValueBasedOnProduct(IOrganizationService service, EntityReference productCode, out EntityReference hsnCode)
        {
            decimal tax = 0;
            Entity productEnt = service.Retrieve(productCode.LogicalName, productCode.Id, new ColumnSet("hil_hsncode"));
            if (!productEnt.Contains("hil_hsncode"))
            {
                hsnCode = null;
                return tax;
                //throw new InvalidPluginExecutionException("Tax Master Setup is not activated for this product.");
            }
            else
            {
                hsnCode = productEnt.GetAttributeValue<EntityReference>("hil_hsncode");
                Entity hsnEntity = service.Retrieve(hsnCode.LogicalName, hsnCode.Id, new ColumnSet("hil_taxpercentage", "hil_efffromdate", "hil_efftodate"));
                //DateTime hil_efffromdate = hsnEntity.GetAttributeValue<DateTime>("hil_efffromdate");
                //DateTime hil_efftodate = hsnEntity.GetAttributeValue<DateTime>("hil_efftodate");
                //DateTime today = DateTime.Today;
                if (hsnEntity.Contains("hil_taxpercentage"))
                {
                    tax = hsnEntity.GetAttributeValue<decimal>("hil_taxpercentage");
                    return tax;
                }
            }
            //else
            //    throw new InvalidPluginExecutionException("Tax percentage is not maintaned.");
            return tax;
        }
        public static decimal getHSNValueBasedOnHSN(IOrganizationService service, EntityReference hsnCode)
        {
            decimal tax = 0;
           Entity hsnEntity = service.Retrieve(hsnCode.LogicalName, hsnCode.Id, new ColumnSet("hil_taxpercentage", "hil_efffromdate", "hil_efftodate"));
            //DateTime hil_efffromdate = hsnEntity.GetAttributeValue<DateTime>("hil_efffromdate");
            //DateTime hil_efftodate = hsnEntity.GetAttributeValue<DateTime>("hil_efftodate");
            //DateTime today = DateTime.Today;
            if (hsnEntity.Contains("hil_taxpercentage"))
            {
                tax = hsnEntity.GetAttributeValue<decimal>("hil_taxpercentage");
                return tax;
            }
            //else
            //    throw new InvalidPluginExecutionException("Tax percentage is not maintaned.");
            return tax;
        }
    }
}
