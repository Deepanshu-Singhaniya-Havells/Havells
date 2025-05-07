using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SyncStatusForSalesOrderPayments
{
    class SyncPayments
    {
        private IOrganizationService service;

        public SyncPayments(IOrganizationService _service)
        {
            service = _service;
        }
        internal void FetchPayments()
        {
            string fetchPayments = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='hil_paymentreceipt'>
                                        <attribute name='hil_paymentreceiptid' />
                                        <attribute name='hil_transactionid' />
                                        <attribute name='createdon' />
                                        <order attribute='hil_transactionid' descending='false' />
                                        <filter type='and'>
                                            <condition attribute='hil_paymentstatus' operator='in'>
                                            <value>3</value>
                                            <value>1</value>
                                            </condition>
                                            <condition attribute='statecode' operator='eq' value='0' />
                                            <condition attribute='createdon' operator='last-x-days' value='2' />
                                        </filter>
                                        </entity>
                                    </fetch>";


            EntityCollection pendingPayments = service.RetrieveMultiple(new FetchExpression(fetchPayments));


            foreach (Entity payment in pendingPayments.Entities)
            {


                Console.WriteLine("Processing the payment with id :" + payment.Id);
                string transactionId = payment.GetAttributeValue<string>("hil_transactionid");


                Console.WriteLine("Transaction Id :" + transactionId);


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
                    Paymentreceipt["hil_paymentreceiptid"] = payment.Id;

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
                        continue;
                    }
                    else
                    {
                        Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(2);
                    }

                    service.Update(Paymentreceipt);
                    StatusPay = item.status.ToLower();

                }
            }
        }

        public class StatusResponse
        {

            public int status { get; set; }

            public string msg { get; set; }

            public List<TransactionDetail> transaction_details { get; set; }
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

        public class StatusRequest
        {

            public string PROJECT { get; set; }

            public string command { get; set; }

            public string var1 { get; set; }

        }
    }
}
