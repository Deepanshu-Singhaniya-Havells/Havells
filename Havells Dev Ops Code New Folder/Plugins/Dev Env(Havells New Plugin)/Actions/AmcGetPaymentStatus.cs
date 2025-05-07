using HavellsNewPlugin.AMC_OmniChannel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Util;

namespace HavellsNewPlugin.Actions
{
    public class AmcGetPaymentStatus : IPlugin
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
                //string recivedAmount = "0";
             
                Entity entity = service.Retrieve(EntityName, new Guid(EntityID), new ColumnSet(true));

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
                    context.OutputParameters["Status"] = "Success";
                    context.OutputParameters["Message"] = "Payment Paid";
                    return;
                }
                else if (_Paymentreceipt.Entities.Count > 0 && _Paymentreceipt.Entities[0].GetAttributeValue<OptionSetValue>("hil_paymentstatus").Value == 2)
                {
                    context.OutputParameters["Status"] = "Failed";
                    context.OutputParameters["Message"] = "Payment Failed";
                    return;
                }
                else if (_Paymentreceipt.Entities.Count == 0)
                {
                    context.OutputParameters["Status"] = "Not Found";
                    context.OutputParameters["Message"] = "Transaction ID Not Found";
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
                        }else if(status == "Not Found"){
                            context.OutputParameters["Status"] = "Not Found";
                            context.OutputParameters["Message"] = "Transaction ID Not Found";
                            return; 
                        }
                        else
                        {
                            Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(2);
                        }

                        service.Update(Paymentreceipt);
                        StatusPay = item.status.ToLower();
                        context.OutputParameters["Status"] = obj.transaction_details[0].status;
                        context.OutputParameters["Message"] = obj.transaction_details[0].status;
                    }
                }                
            }
            catch (Exception ex)
            {
                context.OutputParameters["Status"] = "Error !";
                context.OutputParameters["Message"] = "D365 Internal Error " + ex.Message;
            }
        }
        public string getTransactionID(Entity entity, IOrganizationService service)
        {
            if (entity.LogicalName == "invoice")
            {
                return getTransactionID_FGSales(service, entity.Id, entity.GetAttributeValue<string>("name"));
            }
            else if (entity.LogicalName == "msdyn_paymentdetail")
            {
                return entity.GetAttributeValue<string>("msdyn_name");
            }
            else if (entity.LogicalName == "msdyn_workorder")
            {
                return "D365_" + entity.GetAttributeValue<string>("msdyn_name");
            }
            else if (entity.LogicalName == "salesorder")
            {
                var Query1 = new QueryExpression("hil_paymentreceipt");
                Query1.TopCount = 1;
                Query1.ColumnSet.AddColumns("hil_transactionid", "hil_paymentstatus");
                Query1.Criteria.AddCondition("hil_orderid", ConditionOperator.Equal, entity.Id);
                Query1.AddOrder("createdon", OrderType.Descending);

                EntityCollection _Paymentreceipt = service.RetrieveMultiple(Query1);

                if(_Paymentreceipt.Entities.Count > 0){
                    return _Paymentreceipt.Entities[0].GetAttributeValue<string>("hil_transactionid");
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }
        public static string getTransactionID_FGSales(IOrganizationService service, Guid InvoiceId, string InvoiceNum)
        {
            string transactionID = null;
            QueryExpression Query = new QueryExpression("msdyn_paymentdetail");
            Query.ColumnSet = new ColumnSet("msdyn_name", "statuscode");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("msdyn_invoice", ConditionOperator.Equal, InvoiceId);
            //Query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 910590002);
            Query.AddOrder("createdon", OrderType.Descending);
            EntityCollection FoundPaymentDetails = service.RetrieveMultiple(Query);
            if (FoundPaymentDetails.Entities.Count > 0)
            {
                if (FoundPaymentDetails[0].GetAttributeValue<OptionSetValue>("statuscode").Value != 910590001)//Failed
                    transactionID = FoundPaymentDetails[0].GetAttributeValue<string>("msdyn_name");
                else if (FoundPaymentDetails[0].GetAttributeValue<OptionSetValue>("statuscode").Value == 910590001)
                    transactionID = "S365_" + InvoiceNum + "_" + FoundPaymentDetails.Entities.Count;
            }
            else
                transactionID = "S365_" + InvoiceNum + "_" + 0;
            return transactionID;
        }
    }
}
