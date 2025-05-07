using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.TenderProduct
{
    public class RefreshLP : IPlugin
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
                if (entity.Contains("hil_refreshlp") && entity.GetAttributeValue<bool>("hil_refreshlp"))
                {
                    QueryExpression query = new QueryExpression("hil_tenderproduct");
                    query.ColumnSet = new ColumnSet("hil_basicpriceinrsmtr", "hil_product", "hil_lprsmtr", "hil_hodiscper", "hil_discount", "hil_pricelist", "hil_hopricespecialconstructions");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition(new ConditionExpression("hil_tenderid", ConditionOperator.Equal, entity.Id));
                    query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0)); // Active
                    EntityCollection tendPrdColl = service.RetrieveMultiple(query);
                    foreach (Entity tendPrd in tendPrdColl.Entities)
                    {
                        try
                        {
                            decimal hil_basicpriceinrsmtr = tendPrd.GetAttributeValue<Money>("hil_basicpriceinrsmtr").Value;
                            if (tendPrd.Contains("hil_product"))
                            {
                                tracingService.Trace("Product Contains");
                                Guid _productId = tendPrd.GetAttributeValue<EntityReference>("hil_product").Id;

                                QueryExpression qryPrdLstAmt = new QueryExpression("productpricelevel");
                                qryPrdLstAmt.ColumnSet = new ColumnSet("uomid", "amount", "pricelevelid","hil_cogs");
                                qryPrdLstAmt.NoLock = true;
                                qryPrdLstAmt.Criteria = new FilterExpression(LogicalOperator.And);
                                qryPrdLstAmt.Criteria.AddCondition("productid", ConditionOperator.Equal, _productId);
                                qryPrdLstAmt.Criteria.AddCondition("pricelevelid", ConditionOperator.Equal, tendPrd.GetAttributeValue<EntityReference>("hil_pricelist").Id);

                                EntityCollection entColPrdLstAmt = service.RetrieveMultiple(qryPrdLstAmt);
                                decimal hil_amount = 0;
                                decimal hil_cogs = 0;
                                if (entColPrdLstAmt.Entities.Count > 0)
                                {
                                    hil_amount = entColPrdLstAmt.Entities[0].GetAttributeValue<Money>("amount").Value;
                                    hil_cogs = entColPrdLstAmt.Entities[0].GetAttributeValue<Money>("hil_cogs").Value;
                                }
                                decimal newPer = (1 - (hil_basicpriceinrsmtr / hil_amount)) * 100;
                                if (newPer < 0)
                                {
                                    newPer = 0;
                                }
                                tendPrd["hil_lprsmtr"] = hil_amount;
                                tendPrd["hil_cogs"] = hil_cogs;

                                if (tendPrd.Contains("hil_hopricespecialconstructions"))
                                {
                                    decimal hil_hopricespecialconstructions = tendPrd.GetAttributeValue<Money>("hil_hopricespecialconstructions").Value;
                                    tracingService.Trace("hil_hopricespecialconstructions " + hil_hopricespecialconstructions);
                                    if (hil_hopricespecialconstructions > 0)
                                    {
                                        tendPrd["hil_hodiscper"] = newPer;
                                    }
                                    else
                                        tendPrd["hil_discount"] = newPer;
                                }
                                else
                                {
                                    tendPrd["hil_discount"] = newPer;
                                }
                                service.Update(tendPrd);
                            }
                            else
                            {
                                tracingService.Trace("Product Doesn't Contains");
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error RefreshLP: " + ex.Message);
            }
        }
    }
}
