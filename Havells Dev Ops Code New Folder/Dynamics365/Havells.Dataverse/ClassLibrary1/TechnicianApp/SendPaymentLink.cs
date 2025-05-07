using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Havells.Dataverse.CustomConnector.TechnicianApp
{
    public class SendPaymentLink : IPlugin
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

            if (context.InputParameters.Contains("SalesOrderId") && context.InputParameters["SalesOrderId"] is string
                && context.InputParameters.Contains("MobileNumber") && context.InputParameters["MobileNumber"] is string
                && context.Depth == 1)
            {
                tracingService.Trace("Execution Start");
                string Salesorderid = context.InputParameters["SalesOrderId"].ToString();
                tracingService.Trace("Salesorder Record ID " + Salesorderid);
                string MobileNumber = context.InputParameters["MobileNumber"].ToString();
                if (string.IsNullOrWhiteSpace(Salesorderid))
                {
                    context.OutputParameters["Message"] = "Invalid SalesOrder Number";
                    context.OutputParameters["Status"] = "false";
                }
                else if (string.IsNullOrWhiteSpace(MobileNumber))
                {
                    context.OutputParameters["Message"] = "Invalid Mobile Number";
                    context.OutputParameters["Status"] = "false";
                }
                else
                {
                    SendSMSResponse sendSMSResponse = SendSMS(Salesorderid, MobileNumber, service);
                    context.OutputParameters["Message"] = sendSMSResponse.Message;
                    context.OutputParameters["Status"] = sendSMSResponse.Status;
                    context.OutputParameters["PaymentLink"] = sendSMSResponse.PaymentLink;
                }
                tracingService.Trace("Execution End");
            }
        }
        public SendSMSResponse SendSMS(string Salesorderid, string MobileNumber, IOrganizationService service)
        {
            SendSMSResponse sendSMSResponse = new SendSMSResponse();
            try
            {
                Entity salesOrder = service.Retrieve("salesorder", new Guid(Salesorderid), new ColumnSet(true));
                EntityReference customerRef = salesOrder.GetAttributeValue<EntityReference>("customerid");
                Entity customer = service.Retrieve(customerRef.LogicalName, customerRef.Id, new ColumnSet("mobilephone"));
                string mobileNumber = customer.GetAttributeValue<string>("mobilephone");
                string paymentURL = string.Empty;

                QueryExpression previousReceipts = new QueryExpression("hil_paymentreceipt");
                previousReceipts.ColumnSet = new ColumnSet("hil_paymentstatus", "hil_paymenturl", "hil_tokenexpireson");
                previousReceipts.AddOrder("createdon", OrderType.Descending);
                previousReceipts.Criteria.AddCondition("hil_orderid", ConditionOperator.Equal, salesOrder.Id);
                previousReceipts.TopCount = 1;
                EntityCollection prevReceipts = service.RetrieveMultiple(previousReceipts);
                if (prevReceipts.Entities.Count > 0)
                {
                    DateTime tokenexpireson = prevReceipts.Entities[0].GetAttributeValue<DateTime>("hil_tokenexpireson").AddMinutes(330);

                    OptionSetValue paymentStatus = prevReceipts.Entities[0].GetAttributeValue<OptionSetValue>("hil_paymentstatus");
                    if (prevReceipts.Entities[0].Contains("hil_paymenturl"))
                        paymentURL = prevReceipts.Entities[0].GetAttributeValue<string>("hil_paymenturl");
                    if (paymentStatus.Value == 4)
                    {
                        sendSMSResponse.Message = "Payment has been already done.";
                        sendSMSResponse.PaymentLink = "";
                        sendSMSResponse.Status = "true";
                        return sendSMSResponse;
                    }
                    else if (paymentStatus.Value == 1 || paymentStatus.Value == 3)
                    {
                        if (tokenexpireson > DateTime.Now.AddMinutes(330))
                        {
                            sendSMSResponse.Message = $"Payment has been already initiated.";
                            sendSMSResponse.PaymentLink = paymentURL;
                            sendSMSResponse.Status = "true";
                            return sendSMSResponse;
                        }
                    }
                }
                string zip = string.Empty;
                string _mamorandumCode = string.Empty;
                string state = string.Empty;
                string city = string.Empty;
                Entity address = service.Retrieve("hil_address", salesOrder.GetAttributeValue<EntityReference>("hil_serviceaddress").Id, new ColumnSet("hil_street1", "hil_businessgeo"));
                string fulladdress = address.Contains("hil_street1") ? address.GetAttributeValue<string>("hil_street1").ToString() : string.Empty;
                string businesmapping = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_businessmapping'>
                        <attribute name='hil_businessmappingid' />
                        <attribute name='hil_branch' />
                        <attribute name='hil_state' />
                        <attribute name='hil_city' />
                        <order attribute='hil_branch' descending='false' />
                        <filter type='and'>
                        <condition attribute='hil_businessmappingid' operator='eq' value='{address.GetAttributeValue<EntityReference>("hil_businessgeo").Id}' />
                        </filter>
                        <link-entity name='hil_pincode' from='hil_pincodeid' to='hil_pincode' link-type='outer' alias='ab'>
                        <attribute name='hil_name' />
                        </link-entity>
                        <link-entity name='hil_branch' from='hil_branchid' to='hil_branch' link-type='outer' alias='ac'>
                        <attribute name='hil_mamorandumcode' />
                        </link-entity>
                        </entity>
                        </fetch>";

                EntityCollection businessmappingcoll = service.RetrieveMultiple(new FetchExpression(businesmapping));

                if (businessmappingcoll.Entities.Count != 0)
                {
                    foreach (Entity bmc in businessmappingcoll.Entities)
                    {
                        zip = bmc.Contains("ab.hil_name") ? bmc.GetAttributeValue<AliasedValue>("ab.hil_name").Value.ToString() : string.Empty;
                        _mamorandumCode = bmc.Contains("ac.hil_mamorandumcode") ? bmc.GetAttributeValue<AliasedValue>("ac.hil_mamorandumcode").Value.ToString() : string.Empty;
                        state = bmc.Contains("hil_state") ? bmc.GetAttributeValue<EntityReference>("hil_state").Name.ToString() : string.Empty;
                        city = bmc.Contains("hil_city") ? bmc.GetAttributeValue<EntityReference>("hil_city").Name.ToString() : string.Empty;
                    }
                }
                SendPaymentUrlRequest req = new SendPaymentUrlRequest();
                string comm = "create_invoice";
                req.PROJECT = "D365";
                req.command = comm.Trim();
                RemotePaymentLinkDetails remotePaymentLinkDetails = new RemotePaymentLinkDetails();
                decimal amount;

                if (salesOrder.Contains("hil_receiptamount"))
                {
                    amount = salesOrder.GetAttributeValue<Money>("hil_receiptamount").Value;

                    if (amount == 0)
                    {
                        sendSMSResponse.Message = "Receript amount cannot be 0";
                        sendSMSResponse.Status = "false";
                        return sendSMSResponse;
                    }
                }
                else
                {
                    sendSMSResponse.Message = "Please enter the receipt amount";
                    sendSMSResponse.Status = "false";
                    return sendSMSResponse;
                }

                remotePaymentLinkDetails.amount = amount.ToString();
                string _txnId = string.Empty;
                Entity contact = service.Retrieve("contact", salesOrder.GetAttributeValue<EntityReference>("customerid").Id, new ColumnSet("mobilephone", "emailaddress1", "firstname"));

                string firstName = contact.Contains("firstname") ? contact.GetAttributeValue<string>("firstname").ToString() : string.Empty;
                string email = contact.Contains("emailaddress1") ? contact.GetAttributeValue<string>("emailaddress1").ToString() : "abcd@gmail.com";

                Entity paymentReceipt = new Entity("hil_paymentreceipt");
                paymentReceipt["hil_orderid"] = new EntityReference("salesorder", salesOrder.Id);
                paymentReceipt["hil_email"] = email;
                paymentReceipt["hil_mobilenumber"] = mobileNumber;
                paymentReceipt["hil_amount"] = new Money(amount);
                paymentReceipt["hil_memorandumcode"] = _mamorandumCode;
                paymentReceipt["hil_paymentstatus"] = new OptionSetValue(1);
                paymentReceipt["hil_tokenexpireson"] = DateTime.Now.AddMinutes(330).AddDays(1);//Datetime.Now (330)+24 hr
                Guid receiptId = service.Create(paymentReceipt);

                paymentReceipt = service.Retrieve(paymentReceipt.LogicalName, receiptId, new ColumnSet("hil_transactionid"));

                _txnId = paymentReceipt.GetAttributeValue<string>("hil_transactionid").ToString();// + Counter.ToString().PadLeft(3, '0');
                remotePaymentLinkDetails.txnid = _txnId;
                remotePaymentLinkDetails.firstname = contact.Contains("firstname") ? contact.GetAttributeValue<string>("firstname").ToString() : string.Empty;
                remotePaymentLinkDetails.email = contact.Contains("emailaddress1") ? contact.GetAttributeValue<String>("emailaddress1").ToString() : "abcd@gmail.com";
                remotePaymentLinkDetails.phone = mobileNumber;
                remotePaymentLinkDetails.address1 = Regex.Replace((fulladdress.Length > 99 ? fulladdress.Substring(0, 99) : fulladdress), "[^a-zA-Z0-9]", ""); //fulladdress.Length > 99 ? fulladdress.Substring(0, 99) : fulladdress;
                remotePaymentLinkDetails.state = state;
                remotePaymentLinkDetails.city = city;
                remotePaymentLinkDetails.country = "India";
                remotePaymentLinkDetails.template_id = "1";
                remotePaymentLinkDetails.productinfo = _mamorandumCode;
                remotePaymentLinkDetails.validation_period = "24";
                remotePaymentLinkDetails.send_email_now = "1";
                remotePaymentLinkDetails.send_sms = "1";
                remotePaymentLinkDetails.time_unit = "H";
                remotePaymentLinkDetails.zipcode = zip;
                req.RemotePaymentLinkDetails = remotePaymentLinkDetails;

                IntegrationConfiguration inconfig = GetIntegrationConfiguration(service, "Send Payment Link");
                var data = new StringContent(JsonSerializer.Serialize(req), Encoding.UTF8, "application/json");
                HttpClient client = new HttpClient();
                var byteArray = Encoding.ASCII.GetBytes(inconfig.userName + ":" + inconfig.password);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                HttpResponseMessage response = client.PostAsync(inconfig.url, data).Result;

                var obj = JsonSerializer.Deserialize<SendPaymentUrlResponse>(response.Content.ReadAsStringAsync().Result);

                Entity paymentReceiptUpdate = new Entity("hil_paymentreceipt", receiptId);
                if (obj.msg == null)
                {
                    paymentReceiptUpdate["hil_paymenturl"] = obj.URL;
                    paymentReceiptUpdate["hil_paymentstatus"] = new OptionSetValue(1);
                    paymentReceiptUpdate["hil_response"] = obj.Status;
                    service.Update(paymentReceiptUpdate);
                    sendSMSResponse.Message = obj.Status;
                    sendSMSResponse.PaymentLink = obj.URL;
                    sendSMSResponse.Status = "true";
                }
                else
                {
                    paymentReceiptUpdate["hil_response"] = obj.msg;
                    service.Update(paymentReceiptUpdate);
                    sendSMSResponse.Message = obj.msg;
                    sendSMSResponse.PaymentLink = null;
                    sendSMSResponse.Status = "false";
                }
                return sendSMSResponse;
            }
            catch (Exception ex)
            {
                return new SendSMSResponse
                {
                    Message = "D365 Internal Error " + ex.Message,
                    PaymentLink = null,
                    Status = "false"
                };
            }
        }
        private static IntegrationConfiguration GetIntegrationConfiguration(IOrganizationService _service, string name)
        {
            IntegrationConfiguration inconfig = new IntegrationConfiguration();
            QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
            qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
            qsCType.NoLock = true;
            qsCType.Criteria = new FilterExpression(LogicalOperator.And);
            qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, name);
            Entity integrationConfiguration = _service.RetrieveMultiple(qsCType)[0];
            inconfig.url = integrationConfiguration.GetAttributeValue<string>("hil_url");
            inconfig.userName = integrationConfiguration.GetAttributeValue<string>("hil_username");
            inconfig.password = integrationConfiguration.GetAttributeValue<string>("hil_password");
            return inconfig;
        }
        public class IntegrationConfiguration
        {
            public string url { get; set; }
            public string userName { get; set; }
            public string password { get; set; }
        }

        public class SendSMSResponse
        {
            public string Status { get; set; }
            public string Message { get; set; }
            public string PaymentLink { get; set; }
        }
        public class SendPaymentUrlResponse
        {
            public string URL { get; set; }
            public string Status { get; set; }
            public string msg { get; set; }
        }
        public class SendPaymentUrlRequest
        {
            public string PROJECT { get; set; }
            public string command { get; set; }
            public RemotePaymentLinkDetails RemotePaymentLinkDetails { get; set; }
        }
        public class RemotePaymentLinkDetails
        {
            public string amount { get; set; }
            public string txnid { get; set; }
            public string productinfo { get; set; }
            public string firstname { get; set; }
            public string email { get; set; }
            public string phone { get; set; }
            public string address1 { get; set; }
            public string city { get; set; }
            public string state { get; set; }
            public string country { get; set; }
            public string zipcode { get; set; }
            public string template_id { get; set; }
            public string validation_period { get; set; }
            public string send_email_now { get; set; }
            public string send_sms { get; set; }
            public string time_unit { get; set; }
        }
    }
}

