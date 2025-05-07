using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.OrderCheckListProduct
{
    public class UpdateQuantityAndAmountOnOCl : IPlugin
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
            EntityReference OCLRef = null;
            tracingService.Trace("Message :- " + context.MessageName);
            //throw new InvalidPluginExecutionException("Message :- "+context.MessageName);
            decimal amt = 0;
            try
            {
                if (context.MessageName.ToLower() == "create" || context.MessageName.ToLower() == "update")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if (entity.Contains("hil_orderchecklistid"))
                        OCLRef = ((EntityReference)entity["hil_orderchecklistid"]);
                    else
                    {
                        Entity oclPrd = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_orderchecklistid"));
                        OCLRef = ((EntityReference)oclPrd["hil_orderchecklistid"]);
                    }
                }
                else if (context.MessageName.ToLower() == "delete")
                {
                    EntityReference entityref = (EntityReference)context.InputParameters["Target"];
                    Entity oclPrd = service.Retrieve(entityref.LogicalName, entityref.Id, new ColumnSet("hil_orderchecklistid", "hil_taxamount"));
                    OCLRef = ((EntityReference)oclPrd["hil_orderchecklistid"]);
                    if (oclPrd.Contains("hil_taxamount"))
                    {
                        amt = oclPrd.GetAttributeValue<Money>("hil_taxamount").Value;
                    }
                }   
                CalculationOfTax(service, OCLRef, amt);
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error " + ex.Message);
                throw new InvalidPluginExecutionException("HavellsNewPlugin.TenderModule.UpdateQuantityAndAmountOnOCl.Execute Error " + ex.Message);
            }
        }
        public void CalculationOfTax(IOrganizationService service, EntityReference oclRef, decimal amt = 0)
        {
            Entity ocl = service.Retrieve(oclRef.LogicalName, oclRef.Id, new ColumnSet("hil_typeoforder"));
            string _tenderProductfetch = null;
            string _amountFieldName = null;
            string _qtyFieldName = null;
            if (ocl.GetAttributeValue<OptionSetValue>("hil_typeoforder").Value != 1)
            {
                _amountFieldName = "hil_poamount";
                _qtyFieldName = "hil_poqty";
            }
            else
            {
                _amountFieldName = "hil_totalvalueinrs";
                _qtyFieldName = "hil_quantity";
            }
            _tenderProductfetch = @"<fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                <entity name='hil_orderchecklistproduct'>
                <attribute name='hil_orderchecklistid' alias='tenderNo' groupby='true' />
                <attribute name='" + _amountFieldName + @"' alias='amount' aggregate='sum' />
                <attribute name='" + _qtyFieldName + @"' alias='qty' aggregate='sum' />
                <attribute name='hil_taxamount' alias='tax' aggregate='sum' />
                <filter type='and'>
                    <condition attribute='hil_orderchecklistid' operator='eq' value='" + oclRef.Id + @"' />
                    <condition attribute='statecode' operator='eq' value='0' />
                </filter>
                </entity>
            </fetch>";


            EntityCollection _tenderproductColl = service.RetrieveMultiple(new FetchExpression(_tenderProductfetch));
            tracingService.Trace("_tenderproductColl.count " + _tenderproductColl.Entities.Count);
            if (_tenderproductColl.Entities.Count > 0)
            {
                tracingService.Trace("_tenderproductColl");
                if (_tenderproductColl[0].Contains("amount"))
                {
                    Money FinalAmount = ((Money)((AliasedValue)_tenderproductColl[0]["amount"]).Value);
                    decimal qty = ((decimal)((AliasedValue)_tenderproductColl[0]["qty"]).Value);
                    tracingService.Trace("FinalAmount " + FinalAmount.Value);
                    tracingService.Trace("qty " + qty);
                    Entity hil_orderchecklist = new Entity("hil_orderchecklist");
                   
                    if (_tenderproductColl[0].Contains("tax") && (_tenderproductColl[0].GetAttributeValue<AliasedValue>("tax")).Value != null)
                    {
                        tracingService.Trace("tax " + _tenderproductColl[0].GetAttributeValue<AliasedValue>("tax").ToString());
                        Money TotalTaxAmount = (Money)((AliasedValue)_tenderproductColl[0]["tax"]).Value;
                        tracingService.Trace("tax123 " + TotalTaxAmount.Value);
                        decimal totalTax = TotalTaxAmount.Value - amt;
                        hil_orderchecklist["hil_totaltax"] = new Money(totalTax);
                    }

                    hil_orderchecklist["hil_totalpoamount"] = FinalAmount;
                    hil_orderchecklist["hil_totalpoquantity"] = qty;
                    hil_orderchecklist.Id = oclRef.Id;
                    service.Update(hil_orderchecklist);
                    tracingService.Trace("Tender Uploaded");
                }
            }
            else
            {
                Entity hil_orderchecklist = new Entity("hil_orderchecklist");
                hil_orderchecklist["hil_totalpoamount"] = new Money(0);
                hil_orderchecklist["hil_totalpoquantity"] = (double)0;
                hil_orderchecklist["hil_totaltax"] = new Money(0);
                hil_orderchecklist.Id = oclRef.Id;
                service.Update(hil_orderchecklist);
                tracingService.Trace("Tender Uploaded");
            }
        }
    }
}