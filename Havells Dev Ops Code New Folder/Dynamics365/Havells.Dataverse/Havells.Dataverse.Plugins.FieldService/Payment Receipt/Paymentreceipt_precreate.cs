using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.Payment_Receipt
{
    public class Paymentreceipt_precreate : IPlugin
    {
        public static ITracingService tracingService = null;
        IPluginExecutionContext context;
        public void Execute(IServiceProvider serviceProvider)
        {
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "hil_paymentreceipt" && context.MessageName.ToUpper() == "CREATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    Guid _orderId = Guid.Empty;
                    if (entity.Contains("hil_orderid"))
                        _orderId = entity.GetAttributeValue<EntityReference>("hil_orderid").Id;

                    Entity EntityOrder = service.Retrieve(entity.GetAttributeValue<EntityReference>("hil_orderid").LogicalName, _orderId, new ColumnSet("name"));
                    entity["hil_transactionid"] = $"D365{EntityOrder.GetAttributeValue<string>("name")}{GetCounter(service, entity, _orderId)}";
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("HavellsNewPlugin.Payment_Receipt.Paymentreceipt_precreate : " + ex.Message);
            }
        }
        private string GetCounter(IOrganizationService service, Entity entity, Guid _orderId)
        {
            string _counter = "01";

            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='hil_paymentreceipt'>
            <attribute name='hil_paymentreceiptid' />
            <filter type='and'>
                <condition attribute='hil_orderid' operator='eq' value='{_orderId}' />
            </filter>
            </entity>
            </fetch>";
            EntityCollection entColl = service.RetrieveMultiple(new FetchExpression(_fetchXML));
            if (entColl.Entities.Count > 0)
            {
                _counter = (entColl.Entities.Count + 1).ToString().PadLeft(2, '0');
            }
            return _counter;
        }
    }
}