using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RestSharp;
using System.Text;

namespace PaymentStatus
{
    public class ClsSOUpdatePaymentStatus
    {
        private readonly ServiceClient service;
        public ClsSOUpdatePaymentStatus(ServiceClient _service)
        {
            service = _service;
        }
        public void UpdatePaymentStatus()
        {
            try
            {
                //Payment Receipt.Paymentstatus: 1,3 (Initiated and InProgress)
                //Order.Paymentstatus: 4 (Sucess)
                Console.WriteLine($"Batch Starts: {DateTime.Now}");
                string fetchPaymentReceipts = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='hil_paymentreceipt'>
                        <attribute name='hil_paymentreceiptid' />
                        <attribute name='hil_transactionid' />
                        <attribute name='hil_orderid' />
                        <attribute name='hil_paymentstatus' />
                        <attribute name='hil_paymenturl' />
                        <attribute name='hil_tokenexpireson' />
                        <order attribute='createdon' descending='false' />
                        <filter type='and'>
                            <condition attribute='statecode' operator='eq' value='0' />
                            <condition attribute='hil_paymentstatus' operator='in'>
                            <value>3</value>
                            <value>2</value>
                            <value>1</value>
                            </condition>
                            <condition attribute='createdon' operator='last-x-hours' value='25' />
                        </filter>
                    <link-entity name='salesorder' from='salesorderid' to='hil_orderid' link-type='inner' alias='ab'>
                    <filter type='and'>
                        <condition attribute='hil_paymentstatus' operator='not-in'>
                            <value>2</value>
                        </condition>
                    </filter>
                    </link-entity>
                    </entity>
                    </fetch>";
                EntityCollection paymentReceipts = service.RetrieveMultiple(new FetchExpression(fetchPaymentReceipts));
                for (int i = 0; i < paymentReceipts.Entities.Count; i++)
                {
                    string transactionId = paymentReceipts.Entities[i].GetAttributeValue<string>("hil_transactionid");
                    string paymenturl = paymentReceipts.Entities[i].Contains("hil_paymenturl") ? paymentReceipts.Entities[i].GetAttributeValue<string>("hil_paymenturl") : "";
                    string paymetStatus = paymentReceipts.Entities[i].GetAttributeValue<OptionSetValue>("hil_paymentstatus").Value.ToString();
                    IntegrationConfig intConfig = IntegrationConfiguration("Send Payment Link");
                    string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));

                    DateTime tokenExpire = paymentReceipts.Entities[i].GetAttributeValue<DateTime>("hil_tokenexpireson").AddMinutes(5);
                    Console.WriteLine($"Payment Request#:{transactionId}| Token Expire:{tokenExpire}| Current Date:{DateTime.Now}");

                    StatusRequest reqParm = new StatusRequest();
                    reqParm.PROJECT = "D365";
                    reqParm.command = "verify_payment";
                    reqParm.var1 = transactionId;

                    var client = new RestClient(intConfig.uri);
                    var request = new RestRequest();
                    request.AddHeader("Authorization", authInfo);
                    request.AddHeader("Content-Type", "application/json");
                    request.AddHeader("Cookie", "saplb_*=(J2EE2717920)2717950; JSESSIONID=7fOj-tgnbYBRVBihJMBX9THzyTG3dgH-eCkA_SAPa_yX9TL_PrH5RR_PrxfO7kbO");
                    request.AddParameter("application/json", Newtonsoft.Json.JsonConvert.SerializeObject(reqParm), ParameterType.RequestBody);
                    RestResponse response = client.Execute(request, Method.Post);

                    var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<StatusResponse>(response.Content ?? string.Empty);
                    if (obj?.transaction_details != null)
                    {
                        var item = obj.transaction_details.FirstOrDefault();
                        if (item != null)
                        {
                            Console.WriteLine($"Payment Request# :{transactionId}| Payment Status:{item.status}");
                            Entity Paymentreceipt = new Entity("hil_paymentreceipt", paymentReceipts.Entities[i].Id);
                            if (item.bank_ref_num != null)
                            {
                                Paymentreceipt["hil_bankreferenceid"] = Convert.ToString(item.bank_ref_num);
                            }
                            if (item.mode != null)
                            {
                                Paymentreceipt["hil_paymentmode"] = Convert.ToString(item.mode);
                            }
                            if (item.addedon != null)
                            {
                                Paymentreceipt["hil_receiptdate"] = DateTime.Parse(item.addedon);
                            }
                            if (item.amt != null)
                            {
                                Paymentreceipt["hil_amount"] = Decimal.Parse(item.amt);
                            }
                            if (item.error_Message != null)
                            {
                                Paymentreceipt["hil_response"] = Convert.ToString(item.error_Message);
                            }

                            string status = item.status;
                            if (status == "not initiated")
                            {
                                if (DateTime.Now > tokenExpire)
                                    Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(2); //Failed
                                else
                                    Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(1);//Payment Initiated 
                            }
                            else if (status == "Not Found")
                            {
                                if (DateTime.Now > tokenExpire || paymenturl == "")
                                    Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(2); //Failed
                                else
                                    continue;
                            }
                            else if (status == "success")
                            {
                                Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(4); //Paid
                                Entity Updateorder = new Entity("salesorder", paymentReceipts.Entities[i].GetAttributeValue<EntityReference>("hil_orderid").Id);
                                Updateorder["hil_paymentstatus"] = new OptionSetValue(2);
                                service.Update(Updateorder);
                            }
                            else if (status == "pending")
                            {
                                if (DateTime.Now > tokenExpire)
                                    Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(2); //Failed
                                else
                                    Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(3); //InProgress
                            }
                            else
                            {
                                Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(2);//Failed
                            }
                            service.Update(Paymentreceipt);
                        }
                    }
                }
                Console.WriteLine($"Batch Ends: {DateTime.Now}");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        private IntegrationConfig IntegrationConfiguration(string Param)
        {
            IntegrationConfig output = new IntegrationConfig();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + ":" + integrationConfiguration.GetAttributeValue<string>("hil_password");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            return output;
        }
        public class StatusResponse
        {
            public int status { get; set; }
            public string msg { get; set; }
            public List<TransactionDetail> transaction_details { get; set; }
        }
        public class TransactionDetail
        {
            public object bank_ref_num { get; set; }

            public string amt { get; set; }

            public string addedon { get; set; }

            public string error_Message { get; set; }

            public string mode { get; set; }

            public string status { get; set; }

        }
        public class StatusRequest
        {
            public string PROJECT { get; set; }
            public string command { get; set; }
            public string var1 { get; set; }

        }
        public class IntegrationConfig
        {
            public string uri { get; set; }
            public string Auth { get; set; }
        }

    }
}