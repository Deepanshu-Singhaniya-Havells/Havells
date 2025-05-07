using System;
using System.Globalization;
using HavellsNewPlugin.TenderModule.MailtoTeams;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.PriceListItem
{
    public class UpdateCreatePriceListItem : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_amount", "hil_division"));
                createUpdatePriceListItem(entity, service);

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error " + ex.Message);
            }
        }
        private static void createUpdatePriceListItem(Entity pro, IOrganizationService _service)
        {
            try
            {
               

                if (pro.Contains("hil_amount"))
                {
                    Guid DivisionId = pro.GetAttributeValue<EntityReference>("hil_division").Id;
                    Guid ProductId = pro.Id;
                    if (ProductId != Guid.Empty && DivisionId != Guid.Empty && pro.Contains("hil_amount"))
                    {
                        decimal pamount = pro.GetAttributeValue<Money>("hil_amount").Value;
                        // tracingService.Trace("Amount  " + pamount);
                        Guid iPriceList = Guid.Empty;
                        Guid Uom = Guid.Empty;
                        QueryExpression Query = new QueryExpression("hil_enquirydepartment");
                        Query.ColumnSet = new ColumnSet("hil_pricelist", "hil_uom", "hil_conversionfactor");
                        Query.Criteria = new FilterExpression(LogicalOperator.And);
                        Query.Criteria.AddCondition("hil_productdivision", ConditionOperator.Equal, DivisionId);
                        EntityCollection Found = _service.RetrieveMultiple(Query);
                        if (Found.Entities.Count == 1)
                        {
                            Entity pricelistItem = new Entity("productpricelevel");
                            int conversionFactor = Found.Entities[0].GetAttributeValue<int>("hil_conversionfactor");
                            Decimal caMoney = (pro.GetAttributeValue<Money>("hil_amount").Value) / conversionFactor;
                            iPriceList = Found.Entities[0].GetAttributeValue<EntityReference>("hil_pricelist").Id;
                            Uom = Found.Entities[0].GetAttributeValue<EntityReference>("hil_uom").Id;

                            // tracingService.Trace("caMoney  " + caMoney);
                            pricelistItem["amount"] = new Money(caMoney);
                            pricelistItem["pricelevelid"] = new EntityReference("pricelevel", iPriceList);
                            pricelistItem["uomid"] = new EntityReference("uom", Uom);
                            pricelistItem["quantitysellingcode"] = new OptionSetValue(1);
                            pricelistItem["productid"] = new EntityReference("product", ProductId);

                            QueryExpression query = new QueryExpression("productpricelevel");
                            query.ColumnSet = new ColumnSet(false);
                            query.Distinct = true;
                            query.Criteria.AddCondition("productid", ConditionOperator.Equal, ProductId);
                            query.Criteria.AddCondition("pricelevelid", ConditionOperator.Equal, iPriceList);
                            EntityCollection productColl = _service.RetrieveMultiple(query);
                            if (productColl.Entities.Count == 1)
                            {
                                if (pamount > 0)
                                {
                                    pricelistItem.Id = productColl.Entities[0].Id;
                                    _service.Update(pricelistItem);
                                   // throw new InvalidPluginExecutionException("Price List Item is Update " + done + "/" + totlal);

                                }
                                else
                                {
                                    _service.Delete("productpricelevel", productColl.Entities[0].Id);
                                   // throw new InvalidPluginExecutionException("Price List Item is Deleted" + done + "/" + totlal);
                                }
                            }
                            else
                            {
                                _service.Create(pricelistItem);
                             //   throw new InvalidPluginExecutionException("Price List Item is Created" + done + "/" + totlal);
                            }

                        }
                    }
                }

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Exception " + ex.Message);
            }
        }
    }
}
