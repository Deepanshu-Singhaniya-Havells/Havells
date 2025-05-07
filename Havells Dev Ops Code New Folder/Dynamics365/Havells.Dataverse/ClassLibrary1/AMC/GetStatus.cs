using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.AMC
{
    public class GetStatus : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion


            string LoginUserId = Convert.ToString(context.InputParameters["LoginUserId"]);
            string UserToken = Convert.ToString(context.InputParameters["UserToken"]);

            string jsonString = Convert.ToString(context.InputParameters["reqdata"]);
            var data = JsonSerializer.Deserialize<PaymentStatusParam>(jsonString);
            string SourceType = data.SourceType;
            string InvoiceID = data.InvoiceID.ToString();

            if (!APValidate.IsvalidGuid(InvoiceID))
            {
                string msg = string.IsNullOrWhiteSpace(InvoiceID) ? "InvoiceID is required." : "Invalid InvoiceID.";
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = msg });
                context.OutputParameters["data"] = responnse;
                return;
            }
            if (string.IsNullOrWhiteSpace(SourceType))
            {
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = "Source type is required." });
                context.OutputParameters["data"] = responnse;
                return;
            }
            if (SourceType != "6")
            {
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = "Please enter valid Source Type." });
                context.OutputParameters["data"] = responnse;
                return;
            }
            PaymentStatusParam objparam = new PaymentStatusParam();
            objparam.InvoiceID = InvoiceID.ToString();
            objparam.SourceType = SourceType;
            var response = GetInvoiceStatus(service, objparam, LoginUserId);
            dynamic result;

            if (response.Item2.StatusCode == (int)HttpStatusCode.OK)
                result = JsonSerializer.Serialize(response.Item1);
            else
                result = JsonSerializer.Serialize(response.Item2);
            context.OutputParameters["data"] = result;
            return;
        }
        public (PaymentStatusRes, RequestStatus) GetInvoiceStatus(IOrganizationService _crmService, PaymentStatusParam PaymentStatusParam, string LoginUserId)
        {
            PaymentStatusRes paymentStatusRes = new PaymentStatusRes();
            try
            {
                Guid InvoiceID = Guid.Empty;
                Guid PaymentstatusId = Guid.Empty;
                StatusRequest reqParm = new StatusRequest();
                string TxnId = string.Empty;
                if (_crmService != null)
                {
                    string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' top='1' distinct='false'>
                                      <entity name='hil_paymentreceipt'>
                                        <attribute name='hil_paymentreceiptid' />
                                        <attribute name='hil_transactionid' />
                                        <attribute name='createdon' />
                                        <attribute name='hil_receiptdate' />
                                        <attribute name='hil_paymenturl' />
                                        <attribute name='hil_paymentstatus' />
                                        <attribute name='hil_paymentmode' />
                                        <attribute name='hil_orderid' />
                                        <attribute name='hil_mobilenumber' />
                                        <attribute name='hil_memorandumcode' />
                                        <attribute name='hil_email' />
                                        <attribute name='hil_bankreferenceid' />
                                        <attribute name='hil_amount' />
                                        <order attribute='createdon' descending='true' />
                                        <filter type='and'>
                                          <condition attribute='hil_orderid' operator='eq' value='{PaymentStatusParam.InvoiceID}' />
                                          <condition attribute='statecode' operator='eq' value='0' />
                                        </filter>
                                      </entity>
                                    </fetch>";
                    EntityCollection entpaymentstatus = _crmService.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (entpaymentstatus.Entities.Count == 0)
                    {
                        paymentStatusRes.StatusCode = (int)HttpStatusCode.NotFound;
                        return (paymentStatusRes, new RequestStatus
                        {
                            StatusCode = (int)HttpStatusCode.OK,
                            Message = "Order not found."
                        });
                    }
                    else
                    {
                        int paymentstatus = entpaymentstatus.Entities[0].Contains("hil_paymentstatus") ? entpaymentstatus.Entities[0].GetAttributeValue<OptionSetValue>("hil_paymentstatus").Value : 1;
                        if (paymentstatus == 1 || paymentstatus == 3)//Payment Initiated || In Progress
                        {
                            TxnId = entpaymentstatus.Entities[0].GetAttributeValue<string>("hil_transactionid");
                            string Status = getTransactionStatus(_crmService, entpaymentstatus.Entities[0].Id, TxnId, new Guid(PaymentStatusParam.InvoiceID), "Send Payment Link AMC QA");

                            if (Status == "Success" || Status == "Failed" || Status == "Pending")
                            {
                                paymentStatusRes.PaymentStatus = Status;
                                paymentStatusRes.StatusCode = (int)HttpStatusCode.OK;
                                return (paymentStatusRes, new RequestStatus
                                {
                                    StatusCode = (int)HttpStatusCode.OK
                                });
                            }
                            else
                            {
                                return (paymentStatusRes, new RequestStatus()
                                {
                                    StatusCode = (int)HttpStatusCode.BadRequest,
                                    Message = "D365 internal server error : " + Status
                                });
                            }
                        }
                        else if (paymentstatus == 4)//Paid
                        {
                            paymentStatusRes.PaymentStatus = "Success";
                            paymentStatusRes.StatusCode = (int)HttpStatusCode.OK;
                            return (paymentStatusRes, new RequestStatus
                            {
                                StatusCode = (int)HttpStatusCode.OK
                            });
                        }
                        else
                        {
                            paymentStatusRes.PaymentStatus = "Failed";
                            paymentStatusRes.StatusCode = (int)HttpStatusCode.OK;
                            return (paymentStatusRes, new RequestStatus
                            {
                                StatusCode = (int)HttpStatusCode.OK
                            });
                        }
                    }
                }
                else
                {
                    return (paymentStatusRes, new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.ServiceUnavailable,
                        Message = "D365 service unavailable."
                    });
                }
            }
            catch (Exception ex)
            {
                return (paymentStatusRes, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = "D365 internal server error : " + ex.Message.ToUpper()
                });
            }
        }
        public static string getTransactionStatus(IOrganizationService _service, Guid entPaymentReceiptId, string TxnId, Guid InvoiceId, string _SendPaymentLink)
        {
            string Status = "Failed";
            StatusRequest reqParm = new StatusRequest();
            reqParm.PROJECT = "D365";
            reqParm.command = "verify_payment";
            reqParm.var1 = TxnId;
            try
            {
                IntegrationConfiguration inconfig = GetIntegrationConfiguration(_service, _SendPaymentLink);
                var data = new StringContent(JsonSerializer.Serialize(reqParm), Encoding.UTF8, "application/json");
                HttpClient client = new HttpClient();
                var byteArray = Encoding.ASCII.GetBytes(inconfig.userName + ":" + inconfig.password);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                HttpResponseMessage response = client.PostAsync(inconfig.url, data).Result;
                if (response.IsSuccessStatusCode)
                {
                    var obj = JsonSerializer.Deserialize<StatusResponse>(response.Content.ReadAsStringAsync().Result);
                    foreach (var item in obj.transaction_details)
                    {
                        Entity Paymentreceipt = new Entity("hil_paymentreceipt", entPaymentReceiptId);
                        if (item.bank_ref_num != null)
                        {
                            Paymentreceipt["hil_bankreferenceid"] = Convert.ToString(item.bank_ref_num);
                        }
                        if (!string.IsNullOrWhiteSpace(item.addedon))
                        {
                            Paymentreceipt["hil_receiptdate"] = DateTime.Parse(item.addedon);
                        }
                        if (item.mode != null)
                        {
                            Paymentreceipt["hil_paymentmode"] = item.mode.ToString();
                        }
                        if (!string.IsNullOrWhiteSpace(item.amt))
                        {
                            Paymentreceipt["hil_amount"] = Decimal.Parse(item.amt);
                        }
                        if (!string.IsNullOrWhiteSpace(item.error_Message))
                        {
                            Paymentreceipt["hil_response"] = item.error_Message;
                        }
                        if (item.status.ToLower() == "not initiated" || item.status.ToLower() == "not found")
                        {
                            Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(1);
                            Status = "Pending";
                        }
                        else if (item.status.ToLower() == "pending")
                        {
                            Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(3);
                            Status = "Pending";
                        }
                        else if (item.status.ToLower() == "success")
                        {
                            Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(4);
                            Status = "Success";
                        }
                        else //failure
                        {
                            Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(2);
                        }
                        _service.Update(Paymentreceipt);
                    }
                }
                return Status;
            }
            catch (Exception ex)
            {
                return ex.Message;
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

        #region Get Payment Status
        public class PaymentStatusParam
        {
            public string InvoiceID { get; set; }
            public string SourceType { get; set; }
        }
        public class PaymentStatusRes : TokenExpires
        {
            public string PaymentStatus { get; set; }
        }
        public class StatusRequest
        {
            public string PROJECT { get; set; }
            public string command { get; set; }
            public string var1 { get; set; }
        }
        public class StatusResponse
        {
            public string status { get; set; }
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
        public class IntegrationConfiguration
        {
            public string url { get; set; }
            public string userName { get; set; }
            public string password { get; set; }
        }
        #endregion
    }

}
