using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Havells.Dataverse.Plugins.FieldService.Payment_Receipt
{
    public class PreCreatePaymentRecieptValidateSOLines : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
              && context.PrimaryEntityName.ToLower() == "hil_paymentreceipt" && (context.MessageName.ToLower() == "create"))
            {
                try
                {
                    string fetchXml = string.Empty;
                    Entity paymentReceipt = (Entity)context.InputParameters["Target"];
                    if (paymentReceipt.Contains("hil_orderid"))
                    {
                        EntityReference salesOrderRef = paymentReceipt.GetAttributeValue<EntityReference>("hil_orderid");// ?? throw new InvalidPluginExecutionException("Payment Receipt must be associated with a Sales Order.");
                        fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                        <entity name='salesorder'>
                        <attribute name='salesorderid' />
                        <link-entity name='salesorderdetail' from='salesorderid' to='salesorderid' link-type='inner' alias='sod'>
                        <filter type='and'>
                        <condition attribute='salesorderid' operator='eq' value='{salesOrderRef.Id}' />
                        </filter>
                        </link-entity>
                        </entity>
                        </fetch>";
                        EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetchXml));

                        if (result.Entities.Count == 0)
                        {
                            throw new InvalidPluginExecutionException("Please create Order Line to make payment.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }
    }
}
