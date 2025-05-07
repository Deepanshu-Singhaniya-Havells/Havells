using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.AMC_OmniChannel
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
                if (context.InputParameters.Contains("InvoiceId") && context.InputParameters["InvoiceId"] is string && context.Depth == 1)
                {
                    var InvoiceId = context.InputParameters["InvoiceId"].ToString();
                    try
                    {
                        if (InvoiceId == null)
                        {
                            context.OutputParameters["Status"] = "Failed !!";
                            context.OutputParameters["Message"] = "Invalid Invoice GUID";
                        }

                        Entity Invoice = service.Retrieve("invoice", new Guid(InvoiceId), new ColumnSet("name"));
                        string transID = SendPaymentLink.getTransactionID(service, new Guid(InvoiceId), Invoice.GetAttributeValue<string>("name"));
                        StatusRequest reqParm = new StatusRequest();
                        reqParm.PROJECT = "D365";
                        reqParm.command = "verify_payment";
                        reqParm.var1 = transID;
                        QueryExpression Query1 = new QueryExpression("hil_paymentstatus");
                        Query1.ColumnSet = new ColumnSet(false);
                        Query1.Criteria = new FilterExpression(LogicalOperator.And);
                        Query1.Criteria.AddCondition("hil_name", ConditionOperator.Equal, transID);
                        EntityCollection Paymentstatus = service.RetrieveMultiple(Query1);
                        if (Paymentstatus.Entities.Count == 0)
                        {
                            context.OutputParameters["Status"] = "Payment Status !";
                            context.OutputParameters["Message"] = "Payment Failed";
                            return;
                        }
                        IntegrationConfiguration inconfig = SendPaymentLink.GetIntegrationConfiguration(service, "Send Payment Link");


                        #region logrequest
                        Entity intigrationTrace = new Entity("hil_integrationtrace");
                        intigrationTrace["hil_entityname"] = Invoice.LogicalName;
                        intigrationTrace["hil_entityid"] = Invoice.Id.ToString();
                        intigrationTrace["hil_request"] = Newtonsoft.Json.JsonConvert.SerializeObject(reqParm);
                        intigrationTrace["hil_name"] = Invoice.GetAttributeValue<string>("name");
                        Guid intigrationTraceID = service.Create(intigrationTrace);
                        #endregion logrequest

                        var client = new RestClient(inconfig.url);
                        client.Timeout = -1;
                        var request = new RestRequest(Method.POST);
                        string authInfo = inconfig.userName + ":" + inconfig.password;
                        authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
                        request.AddHeader("Authorization", authInfo);
                        request.AddHeader("Content-Type", "application/json");
                        request.AddParameter("application/json", Newtonsoft.Json.JsonConvert.SerializeObject(reqParm), ParameterType.RequestBody);
                        IRestResponse response = client.Execute(request);

                        var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<StatusResponse>(response.Content);
                        String StatusPay = "";
                        #region logresponse
                        Entity intigrationTraceUp = new Entity("hil_integrationtrace");
                        intigrationTraceUp["hil_response"] = response.Content == "" ? response.ErrorMessage : response.Content;
                        intigrationTraceUp.Id = intigrationTraceID;
                        service.Update(intigrationTraceUp);
                        #endregion logresponse
                        foreach (var item in obj.transaction_details)
                        {
                            Entity statusPayment = new Entity("hil_paymentstatus");
                            statusPayment.Id = Paymentstatus[0].Id;
                            statusPayment["hil_mihpayid"] = item.mihpayid;
                            statusPayment["hil_request_id"] = item.request_id;
                            statusPayment["hil_bank_ref_num"] = item.bank_ref_num;
                            statusPayment["hil_amt"] = item.amt;
                            statusPayment["hil_transaction_amount"] = item.transaction_amount;
                            //statusPayment.txnid = item.txnid;
                            statusPayment["hil_additional_charges"] = item.additional_charges;
                            statusPayment["hil_productinfo"] = item.productinfo;
                            statusPayment["hil_firstname"] = item.firstname;
                            statusPayment["hil_bankcode"] = item.bankcode;
                            statusPayment["hil_udf1"] = item.udf1;
                            statusPayment["hil_udf2"] = item.udf2;
                            statusPayment["hil_udf3"] = item.udf3;
                            statusPayment["hil_udf4"] = item.udf4;
                            statusPayment["hil_udf5"] = item.udf5;
                            statusPayment["hil_field2"] = item.field2;
                            statusPayment["hil_field9"] = item.field9;
                            statusPayment["hil_error_code"] = item.error_code;
                            statusPayment["hil_addedon"] = item.addedon;
                            statusPayment["hil_payment_source"] = item.payment_source;
                            statusPayment["hil_card_type"] = item.card_type;
                            statusPayment["hil_error_message"] = item.error_Message;
                            statusPayment["hil_net_amount_debit"] = item.net_amount_debit;
                            statusPayment["hil_disc"] = item.disc;
                            statusPayment["hil_mode"] = item.mode;
                            statusPayment["hil_pg_type"] = item.PG_TYPE;
                            statusPayment["hil_card_no"] = item.card_no;
                            statusPayment["hil_paymentstatus"] = item.status;
                            statusPayment["hil_unmappedstatus"] = item.unmappedstatus;
                            statusPayment["hil_merchant_utr"] = item.Merchant_UTR;
                            statusPayment["hil_settled_at"] = item.Settled_At;
                            service.Update(statusPayment);
                            StatusPay = item.status.ToLower();
                        }
                        if (StatusPay.ToLower() == "success".ToLower())
                        {
                            context.OutputParameters["Status"] = "Payment Status !";
                            context.OutputParameters["Message"] = "Payment is sucessfully received";
                            SetStateRequest req = new SetStateRequest();
                            req.State = new OptionSetValue(2);
                            req.Status = new OptionSetValue(100001);
                            req.EntityMoniker = Invoice.ToEntityReference();
                            var res = (SetStateResponse)service.Execute(req);

                            QueryExpression Query = new QueryExpression("msdyn_paymentdetail");
                            Query.ColumnSet = new ColumnSet(false);
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, transID);
                            EntityCollection FoundPaymentDetails = service.RetrieveMultiple(Query);

                            req = new SetStateRequest();
                            req.State = new OptionSetValue(0);
                            req.Status = new OptionSetValue(910590000);
                            req.EntityMoniker = FoundPaymentDetails[0].ToEntityReference();
                            res = (SetStateResponse)service.Execute(req);
                        }
                        else
                        {
                            context.OutputParameters["Status"] = "Payment Status !";
                            context.OutputParameters["Message"] = "Payment is " + StatusPay;
                            if (StatusPay.ToLower() == "failure".ToLower())
                            {
                                QueryExpression Query = new QueryExpression("msdyn_paymentdetail");
                                Query.ColumnSet = new ColumnSet(false);
                                Query.Criteria = new FilterExpression(LogicalOperator.And);
                                Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, transID);
                                EntityCollection FoundPaymentDetails = service.RetrieveMultiple(Query);

                                Entity entity = new Entity(FoundPaymentDetails[0].LogicalName, FoundPaymentDetails[0].Id);
                                entity["msdyn_paymentamount"] = new Money(0);
                                service.Update(entity);

                                SetStateRequest req = new SetStateRequest();
                                req.State = new OptionSetValue(0);
                                req.Status = new OptionSetValue(910590001);
                                req.EntityMoniker = FoundPaymentDetails[0].ToEntityReference();
                                service.Execute(req);
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        context.OutputParameters["Status"] = "Error !";
                        context.OutputParameters["Message"] = "D365 Internal Error " + ex.Message;
                    }
                }
            }
            catch (Exception ex)
            {
                context.OutputParameters["Status"] = "Error !";
                context.OutputParameters["Message"] = "D365 Internal Error " + ex.Message;
            }
        }
    }
}
