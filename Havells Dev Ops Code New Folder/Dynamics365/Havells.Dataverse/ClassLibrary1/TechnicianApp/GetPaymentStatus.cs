using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.TechnicianApp
{
    public class GetPaymentStatus : IPlugin
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

            try
            {
                tracingService.Trace("Execution Start");
                string EntityID = context.InputParameters["EntityID"].ToString();
                string EntityName = context.InputParameters["EntityName"].ToString();
                Entity entity = service.Retrieve(EntityName, new Guid(EntityID), new ColumnSet(true));

                var Query1 = new QueryExpression("hil_paymentreceipt");
                Query1.Criteria.AddCondition("hil_orderid", ConditionOperator.Equal, entity.Id);
                Query1.Criteria.AddCondition("hil_paymentstatus", ConditionOperator.Equal, 4);

                EntityCollection _Paymentreceipt = service.RetrieveMultiple(Query1);

                if (_Paymentreceipt.Entities.Count > 0)
                {
                    context.OutputParameters["Status"] = "Success";
                    context.OutputParameters["Message"] = "Payment Paid";
                    return;
                }

                Query1 = new QueryExpression("hil_paymentreceipt");
                Query1.TopCount = 1;
                Query1.ColumnSet.AddColumns("hil_transactionid", "hil_paymentstatus", "hil_tokenexpireson");
                Query1.Criteria.AddCondition("hil_orderid", ConditionOperator.Equal, entity.Id);
                Query1.AddOrder("createdon", OrderType.Descending);
                _Paymentreceipt = service.RetrieveMultiple(Query1);
                if (_Paymentreceipt.Entities.Count > 0)
                {
                    string transactionId = _Paymentreceipt.Entities[0].GetAttributeValue<string>("hil_transactionid");
                    DateTime tokenExpire = _Paymentreceipt.Entities[0].GetAttributeValue<DateTime>("hil_tokenexpireson");

                    StatusRequest reqParm = new StatusRequest();
                    reqParm.PROJECT = "D365";
                    reqParm.command = "verify_payment";
                    reqParm.var1 = transactionId;

                    IntegrationConfiguration inconfig = GetIntegrationConfiguration(service, "Send Payment Link");
                    var data = new StringContent(JsonSerializer.Serialize(reqParm), Encoding.UTF8, "application/json");
                    HttpClient client = new HttpClient();
                    var byteArray = Encoding.ASCII.GetBytes(inconfig.userName + ":" + inconfig.password);
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                    HttpResponseMessage response = client.PostAsync(inconfig.url, data).Result;

                    if (response.IsSuccessStatusCode)
                    {
                        var obj = JsonSerializer.Deserialize<StatusResponse>(response.Content.ReadAsStringAsync().Result);
                        Entity Paymentreceipt = new Entity("hil_paymentreceipt");
                        Paymentreceipt["hil_paymentreceiptid"] = _Paymentreceipt[0].Id;

                        if (obj.transaction_details[0].mode != null)
                        {
                            Paymentreceipt["hil_paymentmode"] = obj.transaction_details[0].mode.ToString();
                        }
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
                        string Message = obj.transaction_details[0].status;

                        if (status == "not initiated")
                        {
                            if (DateTime.Now > tokenExpire)
                            {
                                Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(2); //Failed
                                status = "Failed";
                                Message = "Payment Failed";
                            }
                            else
                            {
                                Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(1);//Payment Initiated 
                            }
                        }
                        else if (status == "success")
                        {
                            Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(4); //Failed
                            status = "Success";
                            Message = "Payment Paid";
                        }
                        else if (status == "pending")
                        {
                            if (DateTime.Now > tokenExpire.AddHours(24))
                                Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(2); //Failed
                            else
                                Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(3); //InProgress
                        }
                        else if (status == "Not Found")
                        {
                            context.OutputParameters["Status"] = "Not Found"; ;
                            context.OutputParameters["Message"] = "Transaction ID Not Found";
                            return;
                        }
                        else
                        {
                            Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(2);
                        }
                        service.Update(Paymentreceipt);
                        context.OutputParameters["Status"] = status;
                        context.OutputParameters["Message"] = Message;
                    }
                }
                else
                {
                    context.OutputParameters["Status"] = "Not Found";
                    context.OutputParameters["Message"] = "Transaction ID Not Found";
                    return;
                }
            }
            catch (Exception ex)
            {
                context.OutputParameters["Status"] = "Error !";
                context.OutputParameters["Message"] = "D365 Internal Error " + ex.Message;
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

        public class StatusResponse
        {
            public int status { get; set; }
            public string msg { get; set; }
            public List<TransactionDetail> transaction_details { get; set; }
        }
        public class IntegrationConfiguration
        {
            public string url { get; set; }
            public string userName { get; set; }
            public string password { get; set; }
        }
        public class TransactionDetail
        {
            public string request_id { get; set; }
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
    }
}
