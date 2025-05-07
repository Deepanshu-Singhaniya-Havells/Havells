using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
namespace Havells.Dataverse.CustomConnector.SalesOrder

{
    public class CancelSalesOrder : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region Setup
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion
            try
            {
                Guid _orderType = new Guid("1f9e3353-0769-ef11-a670-0022486e4abb"); //AMC Sale Guid
                if (!context.InputParameters.Contains("OrderID") || !(context.InputParameters["OrderID"] is string orderIdStr) || string.IsNullOrWhiteSpace(orderIdStr))
                {
                    context.OutputParameters["Message"] = "Sales Order ID is required.";
                    context.OutputParameters["Status"] = false;
                    return;
                }
                if (!context.InputParameters.Contains("Remarks") || !(context.InputParameters["Remarks"] is string remarks) || string.IsNullOrWhiteSpace(remarks))
                {
                    context.OutputParameters["Message"] = "Cancellation Remark is required.";
                    context.OutputParameters["Status"] = false;
                    return;
                }
                if (remarks.Length < 10)
                {
                    context.OutputParameters["Message"] = "Cancellation Remark must be at least 10 characters long.";
                    context.OutputParameters["Status"] = false;
                    return;
                }
                if (!Guid.TryParse(orderIdStr, out Guid orderId))
                {
                    context.OutputParameters["Message"] = "Invalid Sales Order ID format.";
                    context.OutputParameters["Status"] = false;
                    return;
                }
                Entity salesOrder = service.Retrieve("salesorder", orderId, new ColumnSet("hil_paymentstatus", "name", "statecode", "hil_ordertype"));
                // Check if the order Type is AMC Sale
                if (!salesOrder.Contains("hil_ordertype") || salesOrder.GetAttributeValue<EntityReference>("hil_ordertype").Id != _orderType)
                {
                    context.OutputParameters["Message"] = "You are not allowed to cancel this order.";
                    context.OutputParameters["Status"] = false;
                    return;
                }
                string orderNumber = salesOrder.GetAttributeValue<string>("name");
                OptionSetValue paymentStatus = salesOrder.GetAttributeValue<OptionSetValue>("hil_paymentstatus");
                OptionSetValue orderStatus = salesOrder.GetAttributeValue<OptionSetValue>("statecode");
                if (orderStatus.Value == 2) // Already Canceled
                {
                    context.OutputParameters["Message"] = "Sales Order is already canceled.";
                    context.OutputParameters["Status"] = false;
                    return;
                }
                if (paymentStatus != null && paymentStatus.Value == 2) // Success
                {
                    context.OutputParameters["Message"] = "Sales Order cannot be canceled as the payment status is 'Success'.";
                    context.OutputParameters["Status"] = false;
                    return;
                }
                // Fetch Payment Receipt to check hil_tokenexpireson
                string _fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_paymentreceipt'>
                        <attribute name='hil_tokenexpireson' />
                        <filter type='and'>
                        <condition attribute='hil_orderid' operator='eq' value='{orderId}' />
                        </filter>
                        </entity>
                        </fetch>";
                EntityCollection paymentReceipts = service.RetrieveMultiple(new FetchExpression(_fetchXml));
                if (paymentReceipts.Entities.Count > 0)
                {
                    DateTime tokenExpiresOn = paymentReceipts.Entities[0].GetAttributeValue<DateTime>("hil_tokenexpireson");
                    if (tokenExpiresOn > DateTime.UtcNow)
                    {
                        context.OutputParameters["Message"] = "Payment link has not expired, so the order cannot be canceled.";
                        context.OutputParameters["Status"] = false;
                        return;
                    }
                }
                Entity user = service.Retrieve("systemuser", context.UserId, new ColumnSet("firstname", "lastname"));
                string firstName = user.GetAttributeValue<string>("firstname");
                string lastName = user.GetAttributeValue<string>("lastname");
                string description = $"Cancelled By : {firstName} {lastName}\nRemarks : {remarks}";
                CancelSalesOrderRequest request = new CancelSalesOrderRequest
                {
                    OrderClose = new Entity("orderclose")
                    {
                        Attributes =
                        {
                            { "salesorderid", new EntityReference("salesorder", orderId) },
                            { "description", description },
                            { "subject", orderNumber },
                            { "actualend", DateTime.Now }
                        }
                    },
                    Status = new OptionSetValue(4)
                };
                service.Execute(request);
                context.OutputParameters["Message"] = "Sales Order canceled successfully.";
                context.OutputParameters["Status"] = true;
            }
            catch (Exception ex)
            {
                context.OutputParameters["Message"] = $"D365 Internal Server Error: {ex.Message.ToUpper()}";
                context.OutputParameters["Status"] = false;
            }
        }
    }
}