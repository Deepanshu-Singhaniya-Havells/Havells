using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.TenderProduct
{
    public class TenderProductPostCreate : IPlugin
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
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.PrimaryEntityName.ToLower() == "hil_tenderproduct" && context.MessageName.ToUpper() == "UPDATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    tracingService.Trace("DEPTH : " + context.Depth.ToString());
                    Entity preentity = (Entity)context.PreEntityImages["Pre_ProductId"];
                    Entity entity = (Entity)context.InputParameters["Target"];

                    entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));
                    Entity _entTenderDept = service.Retrieve("hil_tender", entity.GetAttributeValue<EntityReference>("hil_tenderid").Id, new ColumnSet("hil_department"));
                    string department = _entTenderDept.GetAttributeValue<EntityReference>("hil_department").Name;
                    Guid user = context.UserId;
                    if (entity.Contains("hil_product") && department.ToLower() == "solar") // Bom insert
                    {
                        EntityReference NewProductId = entity.GetAttributeValue<EntityReference>("hil_product");
                        if (preentity.Contains("hil_product"))
                        {
                            QueryExpression querydel = new QueryExpression("hil_tenderbomlineitem");
                            querydel.ColumnSet = new ColumnSet(true);
                            querydel.Criteria = new FilterExpression(LogicalOperator.And);
                            querydel.Criteria.AddCondition(new ConditionExpression("hil_tenderproduct", ConditionOperator.Equal, preentity.Id));
                            EntityCollection entColldel = service.RetrieveMultiple(querydel);
                            if (entColldel.Entities.Count > 0)
                            {
                                tracingService.Trace("Count " + entColldel.Entities.Count);
                                foreach (Entity ent in entColldel.Entities)
                                {
                                    tracingService.Trace("DELETE");
                                    DeleteEnquiryBomLineItem(service, ent);
                                }
                            }
                        }
                        string prodname = entity.GetAttributeValue<EntityReference>("hil_product").Name;
                        QueryExpression query = new QueryExpression("productassociation");
                        query.ColumnSet = new ColumnSet("productid", "associatedproduct", "quantity", "productisrequired", "uomid", "productassociationid");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition(new ConditionExpression("productid", ConditionOperator.Equal, entity.GetAttributeValue<EntityReference>("hil_product").Id));
                        EntityCollection entColl = service.RetrieveMultiple(query);
                        if (entColl.Entities.Count > 0)
                        {
                            tracingService.Trace("Count " + entColl.Entities.Count);
                            foreach (Entity ent in entColl.Entities)
                            {
                                CreateEnquiryBomLineItem(service, ent, entity);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error " + ex.Message);
                throw new InvalidPluginExecutionException("HavellsNewPlugin.TenderModule.TenderProduct.TenderProductPostCreate.Execute Error " + ex.Message);
            }
        }
        void DeleteEnquiryBomLineItem(IOrganizationService service, Entity BonEntity)
        {
            service.Delete("hil_tenderbomlineitem", BonEntity.Id);
        }
        void CreateEnquiryBomLineItem(IOrganizationService service, Entity AssociateProduct, Entity TenderProduct)
        {
            try
            {
                Entity _EnqBomItem = new Entity("hil_tenderbomlineitem");
                _EnqBomItem["hil_tenderproduct"] = new EntityReference(TenderProduct.LogicalName, TenderProduct.Id);
                _EnqBomItem["hil_tender"] = new EntityReference("hil_tender", TenderProduct.GetAttributeValue<EntityReference>("hil_tenderid").Id);
                _EnqBomItem["hil_skucode"] = AssociateProduct["associatedproduct"];
                if (AssociateProduct.Contains("quantity"))
                {
                    _EnqBomItem["hil_quantity"] = Convert.ToInt32(AssociateProduct.GetAttributeValue<decimal>("quantity"));
                }
                if (AssociateProduct.Contains("uomid"))
                {
                    _EnqBomItem["hil_unit"] = AssociateProduct["uomid"];
                }
                Entity _product = service.Retrieve("product", AssociateProduct.GetAttributeValue<EntityReference>("associatedproduct").Id, new ColumnSet("description", "hil_bommaterialcategory", "hil_materialgroup2", "hil_scope"));
                if (_product.Contains("description"))
                {
                    _EnqBomItem["hil_skudescription"] = _product.GetAttributeValue<string>("description");
                }
                if (_product.Contains("hil_bommaterialcategory"))
                {
                    _EnqBomItem["hil_bomcategory"] = new EntityReference("hil_materialgroup", _product.GetAttributeValue<EntityReference>("hil_bommaterialcategory").Id);
                }
                if (_product.Contains("hil_materialgroup2"))
                {
                    _EnqBomItem["hil_bomsubcategory"] = new EntityReference("hil_materialgroup2", _product.GetAttributeValue<EntityReference>("hil_materialgroup2").Id);
                }
                if (_product.Contains("hil_scope"))
                {
                    _EnqBomItem["hil_scope"] = new OptionSetValue(_product.GetAttributeValue<OptionSetValue>("hil_scope").Value);
                }

                service.Create(_EnqBomItem);
                tracingService.Trace("Complete");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error in BOM Line Creation :" + ex.Message);
            }
        }

    }
}