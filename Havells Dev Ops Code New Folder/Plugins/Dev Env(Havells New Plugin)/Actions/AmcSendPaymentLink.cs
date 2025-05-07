using HavellsNewPlugin.AMC_OmniChannel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RestSharp;
using System;
using System.Text;
using System.Text.RegularExpressions;

namespace HavellsNewPlugin.Actions
{
    public class AmcSendPaymentLink : IPlugin
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

                QueryExpression previousReceipts = new QueryExpression("hil_paymentreceipt");
                previousReceipts.ColumnSet = new ColumnSet("hil_paymentstatus", "hil_paymenturl");
                previousReceipts.AddOrder("createdon", OrderType.Descending);
                previousReceipts.Criteria.AddCondition("hil_paymentstatus", ConditionOperator.In, 1, 3);
                previousReceipts.Criteria.AddCondition("hil_orderid", ConditionOperator.Equal, salesOrder.Id);
                previousReceipts.Criteria.AddCondition("hil_tokenexpireson", ConditionOperator.GreaterThan, DateTime.Now);

                EntityCollection prevReceipts = service.RetrieveMultiple(previousReceipts);
                if (prevReceipts.Entities.Count > 0)
                {
                    string paymentUrl = prevReceipts.Entities[0].GetAttributeValue<string>("hil_paymenturl");

                    sendSMSResponse.Message = "Payment link has already been sent to the customer";
                    sendSMSResponse.PaymentLink = paymentUrl;
                    sendSMSResponse.Status = "true";
                    return sendSMSResponse;
                }
                string zip = string.Empty;
                string _mamorandumCode = string.Empty;
                string state = string.Empty;

                Entity address = service.Retrieve("hil_address", salesOrder.GetAttributeValue<EntityReference>("hil_serviceaddress").Id, new ColumnSet("hil_street1", "hil_businessgeo"));
                string fulladdress = address.Contains("hil_street1") ? address.GetAttributeValue<string>("hil_street1").ToString() : string.Empty;

                string businesmapping = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                <entity name=""hil_businessmapping"">
                <attribute name=""hil_businessmappingid"" />
                <attribute name=""hil_branch"" />
                <attribute name=""hil_state"" />
                <order attribute=""hil_branch"" descending=""false"" />
                <filter type=""and"">
                <condition attribute=""hil_businessmappingid"" operator=""eq"" value=""{address.GetAttributeValue<EntityReference>("hil_businessgeo").Id}"" />
                </filter>
                <link-entity name=""hil_pincode"" from=""hil_pincodeid"" to=""hil_pincode"" link-type=""outer"" alias=""ab"">
                <attribute name=""hil_name"" />
                </link-entity>
                <link-entity name=""hil_branch"" from=""hil_branchid"" to=""hil_branch"" link-type=""outer"" alias=""ac"">
                <attribute name=""hil_mamorandumcode"" />
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
                    amount =  salesOrder.GetAttributeValue<Money>("hil_receiptamount").Value;

                    if(amount == 0)
                    {
                        sendSMSResponse.Message = "Receript amount cannot be 0";
                        sendSMSResponse.Status = "false";
                        return sendSMSResponse;
                    }
                }else
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
                Guid receiptId = service.Create(paymentReceipt);

                paymentReceipt = service.Retrieve(paymentReceipt.LogicalName, receiptId, new ColumnSet("hil_transactionid"));

                _txnId = paymentReceipt.GetAttributeValue<string>("hil_transactionid").ToString();// + Counter.ToString().PadLeft(3, '0');
                remotePaymentLinkDetails.txnid = _txnId;
                remotePaymentLinkDetails.firstname = contact.Contains("firstname") ? contact.GetAttributeValue<string>("firstname").ToString() : string.Empty;
                remotePaymentLinkDetails.email = contact.Contains("emailaddress1") ? contact.GetAttributeValue<String>("emailaddress1").ToString() : string.Empty;
                remotePaymentLinkDetails.phone = mobileNumber;
                remotePaymentLinkDetails.address1 = fulladdress.Length > 99 ? fulladdress.Substring(0, 99) : fulladdress;
                remotePaymentLinkDetails.state = state;
                remotePaymentLinkDetails.country = "India";
                remotePaymentLinkDetails.template_id = "1";
                remotePaymentLinkDetails.productinfo = _mamorandumCode;
                remotePaymentLinkDetails.validation_period = "24";
                remotePaymentLinkDetails.send_email_now = "1";
                remotePaymentLinkDetails.send_sms = "1";
                remotePaymentLinkDetails.time_unit = "H";
                remotePaymentLinkDetails.zipcode = zip;
                req.RemotePaymentLinkDetails = remotePaymentLinkDetails;
                var client = new RestClient("https://middlewareqa.havells.com:50001/RESTAdapter/PayUBiz/TransactionAPI");
                var request = new RestRequest();
                string authInfo = "D365_HAVELLS" + ":" + "QAD365@1234";
                authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
                request.AddHeader("Authorization", authInfo);
                request.AddHeader("Content-Type", "application/json");
                request.AddHeader("Cookie", "saplb_*=(J2EE2717920)2717950; JSESSIONID=7fOj-tgnbYBRVBihJMBX9THzyTG3dgH-eCkA_SAPa_yX9TL_PrH5RR_PrxfO7kbO");
                request.AddParameter("application/json", Newtonsoft.Json.JsonConvert.SerializeObject(req), ParameterType.RequestBody);

                RestResponse response = (RestResponse)client.Execute(request, Method.Post);

                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<SendPaymentUrlResponse>(response.Content);

                if (obj.msg == null)
                {
                    paymentReceipt["hil_paymenturl"] = obj.URL;
                    paymentReceipt["hil_paymentstatus"] = new OptionSetValue(1);
                    paymentReceipt["hil_response"] = obj.Status;
                    paymentReceipt["hil_tokenexpireson"] = DateTime.Now.AddHours(24).AddHours(5.5);
                    service.Update(paymentReceipt);


                    //Console.WriteLine(obj.Status);
                    //Console.WriteLine(obj.URL);
                    //Console.WriteLine("True");
                    sendSMSResponse.Message = obj.Status;
                    sendSMSResponse.PaymentLink = obj.URL;
                    sendSMSResponse.Status = "true";

                }
                else
                {
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
        private string ReplaceSpecialCharacters(string _inputStr)
        {

            return Regex.Replace(_inputStr, @"[.^0-9a-zA-Z]+", "");
        }
    }
    public class SendSMSResponse
    {
        public string Status { get; set; }
        public string Message { get; set; }
        public string PaymentLink { get; set; }
    }
}

