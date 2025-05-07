using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.TenderPaymentDetails
{
    public class TenderPaymentDetailsUpdate : IPlugin
    {
        public static ITracingService tracingService = null;
        Guid tenderID;
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
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    tracingService.Trace("1");
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("2");
                    entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_tenderid"));
                    if (entity.Contains("hil_tenderid"))
                    {
                        
                        
                        tenderID = ((EntityReference)service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_tenderid"))["hil_tenderid"]).Id;
                        string _tenderProductfetch = @"<fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                            <entity name='hil_tenderpaymentdetail'>
                            <attribute name='hil_tenderid' alias='tenderNo' groupby='true' />
                            <attribute name='hil_paymentamount' alias='amount' aggregate='sum' />
                            <filter type='and'>
                                <condition attribute='hil_tenderid' operator='eq' value='" + tenderID + @"' />
                                <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            </entity>
                        </fetch>";
                        EntityCollection _tenderproductColl = service.RetrieveMultiple(new FetchExpression(_tenderProductfetch));
                        tracingService.Trace("_tenderproductColl.count" + _tenderproductColl.Entities.Count);
                        if (_tenderproductColl.Entities.Count > 0)
                        {
                            tracingService.Trace("_tenderproductColl");
                            if (_tenderproductColl.Entities[0].Attributes.Contains("amount"))
                            {
                                tracingService.Trace("_tenderproductCollIF");
                                tracingService.Trace("_tenderproductColl_Amount" + ((Money)((AliasedValue)_tenderproductColl.Entities[0]["amount"]).Value).Value.ToString());
                                Money FinalAmount = ((Money)((AliasedValue)_tenderproductColl.Entities[0]["amount"]).Value);

                                tracingService.Trace("_tenderproductColl.FinalAmount" + FinalAmount.Value);
                                Entity tender = new Entity("hil_tender");
                                tender["hil_totalamountreceived"] = FinalAmount;
                                tender.Id = tenderID;
                                service.Update(tender);
                                tracingService.Trace("Tender Uploaded");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error " + ex.Message);
                throw new InvalidPluginExecutionException("HavellsNewPlugin.TenderModule.TenderPaymentDetails.TenderPaymentDetailsUpdate.Execute Error " + ex.Message);
            }
        }
    }
}

