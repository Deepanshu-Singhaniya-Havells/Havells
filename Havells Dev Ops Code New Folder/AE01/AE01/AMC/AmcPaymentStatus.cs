using Microsoft.OpenApi.Models;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RestSharp;
using System.Text;
using System.Text.RegularExpressions;

namespace AE01.AMC
{
    internal class AmcPaymentStatus
    {
        private IOrganizationService service;
        public AmcPaymentStatus(IOrganizationService _service)
        {
            this.service = _service;
        }
        internal void SendPaymentLink()
        {

            Entity salesOrder = service.Retrieve("salesorder", new Guid("1c15578b-b570-ef11-a671-7c1e520f5a4e"), new ColumnSet(true));
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
                Console.WriteLine("Payment link has already been sent to the customer");
                Console.WriteLine("Payment Link " + paymentUrl);
                return;
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


            decimal amount = salesOrder.GetAttributeValue<Money>("totallineitemamount").Value;
            remotePaymentLinkDetails.amount = amount.ToString();
            string _txnId = string.Empty;
            Entity contact = service.Retrieve("contact", salesOrder.GetAttributeValue<EntityReference>("customerid").Id, new ColumnSet("mobilephone", "emailaddress1", "firstname"));

            string firstName = contact.Contains("firstname") ? contact.GetAttributeValue<string>("firstname").ToString() : string.Empty;
            string email = contact.Contains("emailaddress1") ? contact.GetAttributeValue<string>("emailaddress1").ToString() : string.Empty;

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
            var client = new RestClient("https://middlewaredev.havells.com:50001/RESTAdapter/PayUBiz/TransactionAPI");
            var request = new RestRequest();
            string authInfo = "D365_HAVELLS" + ":" + "DEVD365@1234";
            authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
            request.AddHeader("Authorization", authInfo);
            request.AddHeader("Content-Type", "application/json");
            request.AddHeader("Cookie", "saplb_*=(J2EE2717920)2717950; JSESSIONID=7fOj-tgnbYBRVBihJMBX9THzyTG3dgH-eCkA_SAPa_yX9TL_PrH5RR_PrxfO7kbO");
            request.AddParameter("application/json", Newtonsoft.Json.JsonConvert.SerializeObject(req), ParameterType.RequestBody);

            RestResponse response = (RestResponse)client.Execute(request, Method.Post);

            var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<SendPaymentUrlResponse>(response.Content);

            if (obj.msg == null)
            {
                //paymentReceipt["hil_transactionid"] = obj.Transaction_Id;
                paymentReceipt["hil_paymenturl"] = obj.URL;
                //paymentReceipt["hil_memorandumcode"] = _mamorandumCode;
                paymentReceipt["hil_paymentstatus"] = new OptionSetValue(1);
                //paymentReceipt["hil_mobilenumber"] = obj.Phone;
                //paymentReceipt["hil_email"] = obj.Email_Id;
                paymentReceipt["hil_response"] = obj.Status;
                paymentReceipt["hil_tokenexpireson"] = DateTime.Now.AddHours(24);
                //paymentReceipt["hil_amount"] = amount;
                //paymentReceipt["hil_orderid"] = new EntityReference("salesorder", salesOrder.Id);
                service.Update(paymentReceipt);

                Console.WriteLine(obj.Status);
                Console.WriteLine(obj.URL);
                Console.WriteLine("True");
                //sendSMSResponse.Message = obj.Status;
                //sendSMSResponse.PaymentLink = obj.URL;
                //sendSMSResponse.Status = "true";

            }
            else
            {
                Console.WriteLine(obj.Status);
                Console.WriteLine("no Link");
                Console.WriteLine("False");
                //sendSMSResponse.Message = obj.msg;
                //sendSMSResponse.PaymentLink = null;
                //sendSMSResponse.Status = "false";
            }
        }


        internal void getPaymentStatus()
        {
            Entity entity = service.Retrieve("salesorder", new Guid("6a4a9295-b770-ef11-a671-7c1e520f5a4e"), new ColumnSet(true));

            // entity name, entity id, 
            // check for the transaction id 

            var Query1 = new QueryExpression("hil_paymentreceipt");
            Query1.TopCount = 1;
            Query1.ColumnSet.AddColumns("hil_transactionid", "hil_paymentstatus");
            Query1.Criteria.AddCondition("hil_orderid", ConditionOperator.Equal, entity.Id);
            Query1.AddOrder("createdon", OrderType.Descending);

            EntityCollection _Paymentreceipt = service.RetrieveMultiple(Query1);

            if (_Paymentreceipt.Entities.Count > 0 && _Paymentreceipt.Entities[0].GetAttributeValue<OptionSetValue>("hil_paymentstatus").Value == 4)
            {

                Console.WriteLine("Successs");
                //context.OutputParameters["Status"] = "Success";
                //context.OutputParameters["Message"] = "Payment Paid";
                return;
            }
            else if (_Paymentreceipt.Entities.Count > 0 && _Paymentreceipt.Entities[0].GetAttributeValue<OptionSetValue>("hil_paymentstatus").Value == 2)
            {
                Console.WriteLine("Failed");
                //context.OutputParameters["Status"] = "Failed";
                //context.OutputParameters["Message"] = "Payment Failed";
                return;
            }
            else if (_Paymentreceipt.Entities.Count == 0)
            {
                Console.WriteLine("Not found");
                //context.OutputParameters["Status"] = "Not Found";
                //context.OutputParameters["Message"] = "Transaction ID Not Found";
                return;
            }
            else
            {
                string transactionId = _Paymentreceipt.Entities[0].GetAttributeValue<string>("hil_transactionid");


                StatusRequest reqParm = new StatusRequest();
                reqParm.PROJECT = "D365";
                reqParm.command = "verify_payment";
                reqParm.var1 = transactionId;

                var client = new RestClient("https://middlewareqa.havells.com:50001/RESTAdapter/PayUBiz/TransactionAPI");
                var request = new RestRequest();
                request.AddHeader("Authorization", "Basic RDM2NV9IQVZFTExTOlFBRDM2NUAxMjM0");
                request.AddHeader("Content-Type", "application/json");
                request.AddHeader("Cookie", "saplb_*=(J2EE2717920)2717950; JSESSIONID=7fOj-tgnbYBRVBihJMBX9THzyTG3dgH-eCkA_SAPa_yX9TL_PrH5RR_PrxfO7kbO");
                request.AddParameter("application/json", Newtonsoft.Json.JsonConvert.SerializeObject(reqParm), ParameterType.RequestBody);
                RestResponse response = (RestResponse)client.Execute(request, Method.Post);

                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<StatusResponse>(response.Content);
                String StatusPay = "";
                foreach (var item in obj.transaction_details)
                {
                    Entity Paymentreceipt = new Entity("hil_paymentreceipt");
                    Paymentreceipt["hil_paymentreceiptid"] = _Paymentreceipt[0].Id;

                    if (obj.transaction_details[0].bank_ref_num != null)
                    {
                        Paymentreceipt["hil_bankreferenceid"] = obj.transaction_details[0].bank_ref_num.ToString();
                    }
                    if (obj.transaction_details[0].addedon != null)
                    {
                        Paymentreceipt["hil_receiptdate"] = DateTime.Parse(obj.transaction_details[0].addedon);
                    }
                    if (obj.transaction_details[0].amt != null)
                    {
                        Paymentreceipt["hil_amount"] = Decimal.Parse(obj.transaction_details[0].amt);
                    }
                    if (obj.transaction_details[0].error_Message != null)
                    {
                        Paymentreceipt["hil_response"] = obj.transaction_details[0].error_Message.ToString();
                    }
                    string status = obj.transaction_details[0].status;
                    if (status == "not initiated")
                    {
                        Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(1);
                    }
                    else if (status == "success")
                    {
                        Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(4);
                    }
                    else if (status == "pending")
                    {
                        Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(3);
                    }
                    else if (status == "Not Found")
                    {
                        Console.WriteLine("Transaction not found");
                    }
                    else
                    {
                        Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(2);
                    }

                    service.Update(Paymentreceipt);
                    StatusPay = item.status.ToLower();

                    Console.WriteLine(obj.transaction_details[0].status);
                    Console.WriteLine(obj.transaction_details[0].status);
                    //context.OutputParameters["Status"] = obj.transaction_details[0].status;
                    //context.OutputParameters["Message"] = obj.transaction_details[0].status;
                }
            }




        }

    }

    public class TransactionDetail
    {

        public string mihpayid { get; set; }

        public string request_id { get; set; }

        public string bank_ref_num { get; set; }

        public string amt { get; set; }

        public string transaction_amount { get; set; }

        public string txnid { get; set; }

        public string additional_charges { get; set; }

        public string productinfo { get; set; }

        public string firstname { get; set; }

        public string bankcode { get; set; }

        public string udf1 { get; set; }

        public string udf3 { get; set; }

        public string udf4 { get; set; }

        public string udf5 { get; set; }

        public string field2 { get; set; }

        public string field9 { get; set; }

        public string error_code { get; set; }

        public string addedon { get; set; }

        public string payment_source { get; set; }

        public string card_type { get; set; }

        public string error_Message { get; set; }

        public string net_amount_debit { get; set; }

        public string disc { get; set; }

        public string mode { get; set; }

        public string PG_TYPE { get; set; }

        public string card_no { get; set; }

        public string udf2 { get; set; }

        public string status { get; set; }

        public string unmappedstatus { get; set; }

        public string Merchant_UTR { get; set; }

        public string Settled_At { get; set; }
    }
    public class StatusResponse
    {

        public int status { get; set; }

        public string msg { get; set; }

        public List<TransactionDetail> transaction_details { get; set; }
    }
    public class StatusRequest
    {

        public string PROJECT { get; set; }

        public string command { get; set; }

        public string var1 { get; set; }

    }
    public class SendPaymentUrlResponse
    {

        public string Email_Id { get; set; }

        public string Transaction_Id { get; set; }

        public string URL { get; set; }

        public string Status { get; set; }

        public string Phone { get; set; }

        public string StatusCode { get; set; }

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