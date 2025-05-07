using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule
{
    public class TenderProductPostUpdate_UpdateFinalOfferCost : IPlugin
    {
        public static ITracingService tracingService = null;
        Guid tenderID;
        public void Execute(IServiceProvider serviceProvider)
        {
            EntityReference tenderRef = null;
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            tracingService.Trace("Message :- " + context.MessageName);
            try
            {
                decimal amt = 0;
                if (context.MessageName.ToLower() == "create" || context.MessageName.ToLower() == "update")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if (entity.Contains("hil_tenderid"))
                        tenderRef = ((EntityReference)entity["hil_tenderid"]);
                    else
                    {
                        Entity TenderPrd = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_tenderid"));
                        tenderRef = ((EntityReference)TenderPrd["hil_tenderid"]);
                    }
                }
                else if (context.MessageName.ToLower() == "delete")
                {
                    EntityReference entityref = (EntityReference)context.InputParameters["Target"];
                    Entity oclPrd = service.Retrieve(entityref.LogicalName, entityref.Id, new ColumnSet("hil_tenderid", "hil_totaltaxamount"));
                    tenderRef = ((EntityReference)oclPrd["hil_tenderid"]);
                    if (oclPrd.Contains("hil_totaltaxamount"))
                    {
                        amt = oclPrd.GetAttributeValue<Money>("hil_totaltaxamount").Value;
                    }
                }
                tenderID = tenderRef.Id;
                CalculationOfTax(service, amt);
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error " + ex.Message);
                throw new InvalidPluginExecutionException("HavellsNewPlugin.TenderModule.TenderProductPostUpdate_UpdateFinalOfferCost.Execute Error " + ex.Message);
            }

        }
        public void CalculationOfTax(IOrganizationService service, decimal amt = 0)
        {
            string _tenderProductfetch = @"<fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                                                  <entity name='hil_tenderproduct'>
                                                    <attribute name='hil_tenderid' alias='tenderNo' groupby='true' />
                                                    <attribute name='hil_totalvalueinrs' alias='amount' aggregate='sum' />
                                                    <attribute name='hil_totaltaxamount' alias='TaxAmount' aggregate='sum' />
                                                    <filter type='and'>
                                                      <condition attribute='hil_tenderid' operator='eq' value='" + tenderID + @"' />
                                                      <condition attribute='statecode' operator='eq' value='0' />
                                                    </filter>
                                                  </entity>
                                                </fetch>";
            EntityCollection _tenderproductColl = service.RetrieveMultiple(new FetchExpression(_tenderProductfetch));
            tracingService.Trace("_tenderproductColl.count " + _tenderproductColl.Entities.Count);
            if (_tenderproductColl.Entities.Count > 0)
            {
                tracingService.Trace("_tenderproductColl");
                if (_tenderproductColl[0].Contains("amount") && (_tenderproductColl[0].GetAttributeValue<AliasedValue>("amount")).Value != null)
                {
                    Money FinalAmount = ((Money)((AliasedValue)_tenderproductColl[0]["amount"]).Value);

                    tracingService.Trace("_tenderproductColl.count" + FinalAmount.Value);
                    Entity tender = new Entity("hil_tender");
                    tender["hil_tendercost"] = FinalAmount;
                    tender.Id = tenderID;

                    if (_tenderproductColl[0].Contains("TaxAmount") && (_tenderproductColl[0].GetAttributeValue<AliasedValue>("TaxAmount")).Value != null)
                    {
                        Money TotalTaxAmount = (Money)((AliasedValue)_tenderproductColl[0]["TaxAmount"]).Value;
                        tracingService.Trace("_tenderproductColl.count" + TotalTaxAmount);
                        decimal totalTax = TotalTaxAmount.Value - amt;
                        tender["hil_totaltax"] = new Money(totalTax);
                    }
                    service.Update(tender);
                    tracingService.Trace("Tender Uploaded");
                }
            }
            else
            {
                Entity tender = new Entity("hil_tender");
                tender["hil_tendercost"] = new Money(0);
                tender.Id = tenderID;
                service.Update(tender);
                tracingService.Trace("Tender Uploaded");
            }
            //throw new Exception("dd");
        }

    }
}
