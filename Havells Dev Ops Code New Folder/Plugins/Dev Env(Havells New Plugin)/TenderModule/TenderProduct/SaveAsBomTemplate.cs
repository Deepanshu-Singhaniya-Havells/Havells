using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.TenderProduct
{
    public class SaveAsBomTemplate : IPlugin
    {
        public static ITracingService tracingService = null;
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
                if (context.InputParameters.Contains("BomLineId") && context.InputParameters["BomLineId"] is string && context.Depth == 1)
                {
                    tracingService.Trace("1");
                    var BomLineId = context.InputParameters["BomLineId"].ToString();
                    var ProductId = context.InputParameters["PrductId"].ToString();
                    string[] ArrBomLineId = BomLineId.Split(';');

                    QueryExpression query1 = new QueryExpression("productassociation");
                    query1.ColumnSet = new ColumnSet("productid", "associatedproduct", "quantity", "productisrequired", "uomid", "productassociationid");
                    query1.Criteria = new FilterExpression(LogicalOperator.And);
                    query1.Criteria.AddCondition(new ConditionExpression("productid", ConditionOperator.Equal, new Guid(ProductId)));
                    EntityCollection entCol2 = service.RetrieveMultiple(query1); //here we get all kit product of final offer product id
                    if (entCol2.Entities.Count > 0)
                    {
                        tracingService.Trace("2");
                        tracingService.Trace("Count " + entCol2.Entities.Count);
                        foreach (Entity entAssociate in entCol2.Entities)
                        {
                            DeleteAssociatedProduct(service, entAssociate); //here we delete all associate product
                        }
                    }
                    foreach (string guid in ArrBomLineId)
                    {
                        QueryExpression query = new QueryExpression("hil_tenderbomlineitem");
                        query.ColumnSet = new ColumnSet("hil_bomcategory", "hil_bomsubcategory", "hil_tenderproduct", "hil_skucode", "hil_skudescription", "hil_quantity", "hil_unit", "hil_scope");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition(new ConditionExpression("hil_tenderbomlineitemid", ConditionOperator.Equal, new Guid(guid)));
                        EntityCollection entColl = service.RetrieveMultiple(query);// in entColl We collect tenderbomlineitem detail

                        Entity tenderproductitem = service.Retrieve(entColl.Entities[0].GetAttributeValue<EntityReference>("hil_tenderproduct").LogicalName, entColl.Entities[0].GetAttributeValue<EntityReference>("hil_tenderproduct").Id, new ColumnSet("hil_product"));
                        tracingService.Trace("4");
                        EntityReference _Product = tenderproductitem.GetAttributeValue<EntityReference>("hil_product"); //here we get final offer product's productid


                        if (entColl.Entities.Count > 0)
                        {//here we insert bom line item in as kit product 
                            tracingService.Trace("5");
                            foreach (Entity entbomline in entColl.Entities)
                            {
                                Entity AssociateProduct = new Entity("productassociation");
                                AssociateProduct["productid"] = _Product;
                                if (entbomline.Contains("hil_skucode"))
                                {
                                    AssociateProduct["associatedproduct"] = entbomline["hil_skucode"];
                                }
                                if (entbomline.Contains("hil_quantity"))
                                {
                                    AssociateProduct["quantity"] = Convert.ToDecimal(entbomline.GetAttributeValue<Int32>("hil_quantity"));
                                }
                                if (entbomline.Contains("hil_unit"))
                                {
                                    AssociateProduct["uomid"] = entbomline["hil_unit"];
                                }
                                Entity _productdetail = service.Retrieve("product", entbomline.GetAttributeValue<EntityReference>("hil_skucode").Id, new ColumnSet("description", "hil_bommaterialcategory", "hil_materialgroup2", "hil_scope"));
                                Entity ProductBase = new Entity("product");
                                ProductBase.Id = entbomline.GetAttributeValue<EntityReference>("hil_skucode").Id;
                                if (entbomline.Contains("hil_skudescription"))
                                {
                                    ProductBase["description"] = entbomline.GetAttributeValue<string>("hil_skudescription");
                                }
                                if (entbomline.Contains("hil_bomcategory"))
                                {
                                    ProductBase["hil_bommaterialcategory"] = new EntityReference("hil_materialgroup", entbomline.GetAttributeValue<EntityReference>("hil_bomcategory").Id);
                                }
                                if (entbomline.Contains("hil_bomsubcategory"))
                                {
                                    ProductBase["hil_materialgroup2"] = new EntityReference("hil_materialgroup2", entbomline.GetAttributeValue<EntityReference>("hil_bomsubcategory").Id);
                                }
                                if (entbomline.Contains("hil_scope"))
                                {
                                    ProductBase["hil_scope"] = new OptionSetValue(entbomline.GetAttributeValue<OptionSetValue>("hil_scope").Value);
                                }

                                service.Update(ProductBase);
                                service.Create(AssociateProduct);

                                tracingService.Trace("6");
                            }
                        }


                    }
                    context.OutputParameters["message"] = "Template Save Successfully";
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error " + ex.Message);
                throw new InvalidPluginExecutionException("HavellsNewPlugin.TenderModule.TenderProduct.SaveAsBomTemplate.Execute Error " + ex.Message);
            }
        }
        void DeleteAssociatedProduct(IOrganizationService service, Entity ent)
        {
            tracingService.Trace("3");
            service.Delete("productassociation", ent.Id);
            tracingService.Trace("Deleted " + ent.Id);
        }
    }
}
