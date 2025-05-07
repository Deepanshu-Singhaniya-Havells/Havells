using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Havells.CRM.AMCInvoiceSync.Model;
using System.Web.Util;
using System.Security.Policy;
using Microsoft.Crm.Sdk.Messages;
using System.Runtime.Remoting.Contexts;
using System.Windows.Controls.Primitives;

namespace Havells.CRM.AMCInvoiceSync
{
    public class GetPendingPaymentStatus
    {
        public static void getPaymentStatus(IOrganizationService service)
        {
            IntegrationConfig intConfig = Program.IntegrationConfiguration(service, "Send Payment Link");
            //intConfig.uri = "https://p90ci.havells.com:50001/RESTAdapter/PayUBiz/TransactionAPI";
            string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));

            QueryExpression Query = new QueryExpression("hil_paymentstatus");
            Query.ColumnSet = new ColumnSet("hil_job", "hil_name");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_paymentstatus", ConditionOperator.NotEqual, "success");
            // Query.Criteria.AddCondition("hil_paymentstatus", ConditionOperator.NotEqual, "failure");
            Query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, DateTime.Today.AddDays(-7));
            Query.Criteria.AddCondition("hil_job", ConditionOperator.NotNull);
            Query.Orders.Add(new OrderExpression("createdon", OrderType.Ascending));
            EntityCollection paymentCollection = service.RetrieveMultiple(Query);

            int total = paymentCollection.Entities.Count;
            Console.WriteLine("Total Payment Collection is " + total);
            int count = 0;
            foreach (Entity entity in paymentCollection.Entities)
            {
                count++;
                string Paystatus = null;
                string transactionID = entity.GetAttributeValue<string>("hil_name");                
                Console.WriteLine(count + "/" + total + " Status Updated of transactionid " + transactionID +" Status is "+ updatePaymentStatusforJob(transactionID, service, intConfig.uri, authInfo));
            }

            Console.WriteLine("-------------------------------------------Done----------------------------------------");



        }
        public static void getPaymentStatusofInvoice(IOrganizationService service)
        {
            string _paymentEffectiveFrom = DateTime.Now.AddDays(-3).ToString("yyyy-MM-dd");
            IntegrationConfig intConfig = Program.IntegrationConfiguration(service, "Send Payment Link");
            string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));

            QueryExpression Query = new QueryExpression("msdyn_paymentdetail");
            Query.ColumnSet = new ColumnSet("msdyn_name", "msdyn_invoice");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);//Active
            Query.Criteria.AddCondition("statuscode", ConditionOperator.NotEqual, 910590000);//Not Received
            Query.Criteria.AddCondition("msdyn_invoice", ConditionOperator.NotNull);
            Query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, _paymentEffectiveFrom);
            //Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Like, "%INV-2023-003091%");
            Query.Orders.Add(new OrderExpression("createdon", OrderType.Descending));
            EntityCollection paymentCollection = service.RetrieveMultiple(Query);

            int total = paymentCollection.Entities.Count;
            Console.WriteLine("Total Payment Collection is " + total);
            int count = 0;
            foreach (Entity entity in paymentCollection.Entities)
            {
                count++;
                string transactionID = entity.GetAttributeValue<string>("msdyn_name");
                EntityReference Invoice = entity.GetAttributeValue<EntityReference>("msdyn_invoice");
                updatePaymentStatusforInvoice(transactionID, Invoice, service, intConfig.uri, authInfo);
                Console.WriteLine(count + "/" + total + " Status Updated of transactionid " + transactionID);
            }
            Console.WriteLine("-------------------------------------------Done----------------------------------------");
        }
        static void updatePaymentStatusforInvoice(String transactionID, EntityReference Invoice, IOrganizationService service, string URL, string authInfo)
        {
            try
            {
                String StatusPay = "";
                StatusRequest req = new StatusRequest();
                req.PROJECT = "D365";
                req.command = "verify_payment";
                req.var1 = transactionID;
                var client = new RestClient(URL);
                client.Timeout = -1;
                var request = new RestRequest(Method.POST);

                request.AddHeader("Authorization", authInfo);
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter("application/json", Newtonsoft.Json.JsonConvert.SerializeObject(req), ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);

                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<StatusResponse>(response.Content);
                QueryExpression Query = new QueryExpression("hil_paymentstatus");
                Query.ColumnSet = new ColumnSet("hil_job", "hil_name");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, transactionID);
                EntityCollection Found = service.RetrieveMultiple(Query);
                foreach (var item in obj.transaction_details)
                {
                    Entity statusPayment = new Entity("hil_paymentstatus");
                    statusPayment.Id = Found[0].Id;
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
                Console.WriteLine(StatusPay + " Payment Status of transactionid " + transactionID);
                if (StatusPay.ToLower() == "success".ToLower())
                {
                    SetStateRequest req2 = new SetStateRequest();
                    req2.State = new OptionSetValue(2);
                    req2.Status = new OptionSetValue(100001);
                    req2.EntityMoniker = Invoice;
                    var res = (SetStateResponse)service.Execute(req2);

                    Query = new QueryExpression("msdyn_paymentdetail");
                    Query.ColumnSet = new ColumnSet(false);
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, transactionID);
                    EntityCollection FoundPaymentDetails = service.RetrieveMultiple(Query);

                    req2 = new SetStateRequest();
                    req2.State = new OptionSetValue(0);
                    req2.Status = new OptionSetValue(910590000);
                    req2.EntityMoniker = FoundPaymentDetails[0].ToEntityReference();
                    res = (SetStateResponse)service.Execute(req2);
                }
                else if (StatusPay.ToLower() == "failure".ToLower())
                {
                    Query = new QueryExpression("msdyn_paymentdetail");
                    Query.ColumnSet = new ColumnSet(false);
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, transactionID);
                    EntityCollection FoundPaymentDetails = service.RetrieveMultiple(Query);

                    Entity entity = new Entity(FoundPaymentDetails[0].LogicalName, FoundPaymentDetails[0].Id);
                    entity["msdyn_paymentamount"] = new Money(0);
                    service.Update(entity);

                    SetStateRequest req1 = new SetStateRequest();
                    req1.State = new OptionSetValue(0);
                    req1.Status = new OptionSetValue(910590001);
                    req1.EntityMoniker = FoundPaymentDetails[0].ToEntityReference();
                    service.Execute(req1);
                }
                else if (StatusPay.ToLower() == "not found".ToLower())
                {
                    Query = new QueryExpression("msdyn_paymentdetail");
                    Query.ColumnSet = new ColumnSet(false);
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, transactionID);
                    //EntityCollection FoundPaymentDetails = service.RetrieveMultiple(Query);
                    //SetStateRequest req1 = new SetStateRequest();
                    //req1.State = new OptionSetValue(0);
                    //req1.Status = new OptionSetValue(910590001);
                    //req1.EntityMoniker = FoundPaymentDetails[0].ToEntityReference();
                    //service.Execute(req1);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("D365 Internal Error " + ex);
            }
        }

        static string updatePaymentStatusforJob(String transactionID, IOrganizationService service, string URL, string authInfo)
        {
            string payStatus = string.Empty;
            try
            {
                StatusRequest req = new StatusRequest();
                req.PROJECT = "D365";
                req.command = "verify_payment";
                req.var1 = transactionID;
                var client = new RestClient(URL);
                client.Timeout = -1;
                var request = new RestRequest(Method.POST);

                request.AddHeader("Authorization", authInfo);
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter("application/json", Newtonsoft.Json.JsonConvert.SerializeObject(req), ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);

                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<StatusResponse>(response.Content);
                QueryExpression Query = new QueryExpression("hil_paymentstatus");
                Query.ColumnSet = new ColumnSet("hil_job", "hil_name");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, transactionID);
                EntityCollection Found = service.RetrieveMultiple(Query);
                string bank_ref_number = string.Empty;
                if (Found.Entities.Count > 0)
                {
                    Entity statusPayment = new Entity("hil_paymentstatus");
                    statusPayment.Id = Found[0].Id;
                    if (obj.transaction_details[0].mihpayid != null)
                    {
                        statusPayment["hil_mihpayid"] = obj.transaction_details[0].mihpayid.ToString();
                    }
                    if (obj.transaction_details[0].request_id != null)
                    {
                        statusPayment["hil_request_id"] = obj.transaction_details[0].request_id;
                    }
                    if (obj.transaction_details[0].bank_ref_num != null)
                    {
                        statusPayment["hil_bank_ref_num"] = obj.transaction_details[0].bank_ref_num;
                        bank_ref_number = obj.transaction_details[0].bank_ref_num;
                    }
                    if (obj.transaction_details[0].amt != null)
                    {
                        statusPayment["hil_amt"] = obj.transaction_details[0].amt;
                    }
                    if (obj.transaction_details[0].transaction_amount != null)
                    {
                        statusPayment["hil_transaction_amount"] = obj.transaction_details[0].transaction_amount;
                    }
                    if (obj.transaction_details[0].additional_charges != null)
                    {
                        statusPayment["hil_additional_charges"] = obj.transaction_details[0].additional_charges;
                    }
                    if (obj.transaction_details[0].productinfo != null)
                    {
                        statusPayment["hil_productinfo"] = obj.transaction_details[0].productinfo;
                    }
                    if (obj.transaction_details[0].firstname != null)
                    {
                        statusPayment["hil_firstname"] = obj.transaction_details[0].firstname;
                    }
                    if (obj.transaction_details[0].bankcode != null)
                    {
                        statusPayment["hil_bankcode"] = obj.transaction_details[0].bankcode;
                    }
                    if (obj.transaction_details[0].udf1 != null)
                    {
                        statusPayment["hil_udf1"] = obj.transaction_details[0].udf1;
                    }
                    if (obj.transaction_details[0].udf2 != null)
                    {
                        statusPayment["hil_udf2"] = obj.transaction_details[0].udf2;
                    }
                    if (obj.transaction_details[0].udf3 != null)
                    {
                        statusPayment["hil_udf3"] = obj.transaction_details[0].udf3;
                    }
                    if (obj.transaction_details[0].udf4 != null)
                    {
                        statusPayment["hil_udf4"] = obj.transaction_details[0].udf4;
                    }
                    if (obj.transaction_details[0].udf5 != null)
                    {
                        statusPayment["hil_udf5"] = obj.transaction_details[0].udf5;
                    }
                    if (obj.transaction_details[0].field2 != null)
                    {
                        statusPayment["hil_field2"] = obj.transaction_details[0].field2;
                    }
                    if (obj.transaction_details[0].field9 != null)
                    {
                        statusPayment["hil_field9"] = obj.transaction_details[0].field9.Length > 100 ? obj.transaction_details[0].field9.Substring(0, 99) : obj.transaction_details[0].field9;
                    }
                    if (obj.transaction_details[0].error_code != null)
                    {
                        statusPayment["hil_error_code"] = obj.transaction_details[0].error_code;
                    }
                    if (obj.transaction_details[0].addedon != null)
                    {
                        statusPayment["hil_addedon"] = obj.transaction_details[0].addedon;
                    }
                    if (obj.transaction_details[0].payment_source != null)
                    {
                        statusPayment["hil_payment_source"] = obj.transaction_details[0].payment_source;
                    }
                    if (obj.transaction_details[0].card_type != null)
                    {
                        statusPayment["hil_card_type"] = obj.transaction_details[0].card_type;
                    }
                    if (obj.transaction_details[0].error_Message != null)
                    {
                        statusPayment["hil_error_message"] = obj.transaction_details[0].error_Message;
                    }
                    if (obj.transaction_details[0].net_amount_debit != null)
                    {
                        statusPayment["hil_net_amount_debit"] = obj.transaction_details[0].net_amount_debit;
                    }
                    if (obj.transaction_details[0].disc != null)
                    {
                        statusPayment["hil_disc"] = obj.transaction_details[0].disc;
                    }
                    if (obj.transaction_details[0].mode != null)
                    {
                        statusPayment["hil_mode"] = obj.transaction_details[0].mode;
                    }
                    if (obj.transaction_details[0].PG_TYPE != null)
                    {
                        statusPayment["hil_pg_type"] = obj.transaction_details[0].PG_TYPE;
                    }
                    if (obj.transaction_details[0].card_no != null)
                    {
                        statusPayment["hil_card_no"] = obj.transaction_details[0].card_no;
                    }
                    if (obj.transaction_details[0].status != null)
                    {
                        statusPayment["hil_paymentstatus"] = obj.transaction_details[0].status;
                    }
                    if (obj.transaction_details[0].unmappedstatus != null)
                    {
                        statusPayment["hil_unmappedstatus"] = obj.transaction_details[0].unmappedstatus;
                    }
                    if (obj.transaction_details[0].Merchant_UTR != null)
                    {
                        statusPayment["hil_merchant_utr"] = obj.transaction_details[0].Merchant_UTR;
                    }
                    if (obj.transaction_details[0].Settled_At != null)
                    {
                        statusPayment["hil_settled_at"] = obj.transaction_details[0].Settled_At;
                    }
                    service.Update(statusPayment);

                    if (Found[0].Contains("hil_job"))
                    {
                        int paymentstatus = 0;

                        if (obj.transaction_details[0].status == "Not Found")
                        {
                            paymentstatus = (1);
                        }
                        else if (obj.transaction_details[0].status == "success")
                        {
                            paymentstatus = (2);
                        }
                        else if (obj.transaction_details[0].status == "pending")
                        {
                            paymentstatus = (3);
                        }
                        else
                            paymentstatus = (4);
                        payStatus = obj.transaction_details[0].status;
                        updatePaymentinJob(service, Found[0].GetAttributeValue<EntityReference>("hil_job").Id, paymentstatus, bank_ref_number);
                    }
                }
            
            }
            catch (Exception ex)
            {
                Console.WriteLine("D365 Internal Error " + ex);
            }
            return payStatus;
        }
        static void updatePaymentinJob(IOrganizationService service, Guid JobId, int paymentstatus, string bank_ref_number)
        {
            #region Updating Job Payment Status field

            Entity _updateJob = new Entity(msdyn_workorder.EntityLogicalName, JobId);
            _updateJob["hil_paymentstatus"] = new OptionSetValue(paymentstatus);
            _updateJob["hil_receiptnumber"] = bank_ref_number;
            service.Update(_updateJob);

            #endregion
        }
    }
}
