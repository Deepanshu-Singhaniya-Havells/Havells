using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services.Description;
using Newtonsoft.Json;
using System.Runtime.Remoting.Contexts;

namespace TestAPP
{
    public class paystats
    {
        static public void getstat(IOrganizationService service)
        {

            try
            {
                string EntityID = "715ec489-5607-ee11-8f6e-6045bdac5a1d";// context.InputParameters["EntityID"].ToString();
                string EntityName = "msdyn_paymentdetail";// context.InputParameters["EntityName"].ToString();
                string recivedAmount = "0";
                Entity entity = service.Retrieve(EntityName, new Guid(EntityID), new ColumnSet(true));
                string transactionId = getTransactionID(entity, service);
                StatusRequest reqParm = new StatusRequest();
                reqParm.PROJECT = "D365";
                reqParm.command = "verify_payment";
                reqParm.var1 = transactionId;
                string bank_ref_number = null;
                QueryExpression Query1 = new QueryExpression("hil_paymentstatus");
                Query1.ColumnSet = new ColumnSet(false);
                Query1.Criteria = new FilterExpression(LogicalOperator.And);
                Query1.Criteria.AddCondition("hil_name", ConditionOperator.Equal, transactionId);
                EntityCollection Paymentstatus = service.RetrieveMultiple(Query1);
                if (Paymentstatus.Entities.Count == 0)
                {
                    Console.WriteLine("Error");
                    Console.WriteLine("Message Payment Details Record Not Found");
                    return;
                }
                IntegrationConfiguration inconfig = SendPaymentLink.GetIntegrationConfiguration(service, "Send Payment Link");
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
                foreach (var item in obj.transaction_details)
                {
                    Entity statusPayment = new Entity("hil_paymentstatus");
                    statusPayment.Id = Paymentstatus[0].Id;
                    statusPayment["hil_mihpayid"] = item.mihpayid;
                    statusPayment["hil_request_id"] = item.request_id;
                    statusPayment["hil_bank_ref_num"] = item.bank_ref_num;
                    bank_ref_number = item.bank_ref_num;
                    statusPayment["hil_amt"] = item.amt;
                    recivedAmount = item.amt;
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
                if (entity.LogicalName == "invoice")
                {
                    if (StatusPay.ToLower() == "success".ToLower())
                    {
                        Console.WriteLine("Payment Status !");
                        Console.WriteLine("Message Payment received successfully.");
                        SetStateRequest req = new SetStateRequest();
                        req.State = new OptionSetValue(2);
                        req.Status = new OptionSetValue(100001);
                        req.EntityMoniker = entity.ToEntityReference();
                        var res = (SetStateResponse)service.Execute(req);

                        QueryExpression Query = new QueryExpression("msdyn_paymentdetail");
                        Query.ColumnSet = new ColumnSet(false);
                        Query.Criteria = new FilterExpression(LogicalOperator.And);
                        Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, transactionId);
                        EntityCollection FoundPaymentDetails = service.RetrieveMultiple(Query);

                        req = new SetStateRequest();
                        req.State = new OptionSetValue(0);
                        req.Status = new OptionSetValue(910590000);
                        req.EntityMoniker = FoundPaymentDetails[0].ToEntityReference();
                        res = (SetStateResponse)service.Execute(req);
                    }
                    else
                    {
                        Console.WriteLine("Payment Status !");
                        Console.WriteLine("Payment is " + StatusPay);
                        if (StatusPay.ToLower() == "failure".ToLower())
                        {
                            QueryExpression Query = new QueryExpression("msdyn_paymentdetail");
                            Query.ColumnSet = new ColumnSet(false);
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, transactionId);
                            EntityCollection FoundPaymentDetails = service.RetrieveMultiple(Query);

                            Entity entity1 = new Entity(FoundPaymentDetails[0].LogicalName, FoundPaymentDetails[0].Id);
                            entity1["msdyn_paymentamount"] = new Money(0);
                            service.Update(entity1);

                            SetStateRequest req = new SetStateRequest();
                            req.State = new OptionSetValue(0);
                            req.Status = new OptionSetValue(910590001);
                            req.EntityMoniker = FoundPaymentDetails[0].ToEntityReference();
                            service.Execute(req);
                        }
                    }

                }
                else if (entity.LogicalName == "msdyn_paymentdetail")
                {
                    EntityReference invoice = entity.GetAttributeValue<EntityReference>("msdyn_invoice");
                    if (StatusPay.ToLower() == "success".ToLower())
                    {
                        Console.WriteLine("Payment Status !");
                        Console.WriteLine("Payment received successfully.");
                        SetStateRequest req2 = new SetStateRequest();
                        req2.State = new OptionSetValue(2);
                        req2.Status = new OptionSetValue(100001);
                        req2.EntityMoniker = invoice;
                        var res = (SetStateResponse)service.Execute(req2);

                        QueryExpression query = new QueryExpression("msdyn_paymentdetail");
                        query.ColumnSet = new ColumnSet("msdyn_payment");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("msdyn_invoice", ConditionOperator.Equal, invoice.Id);
                        query.Criteria.AddCondition("msdyn_paymentdetailid", ConditionOperator.NotEqual, entity.Id);
                        EntityCollection entityCollection = service.RetrieveMultiple(query);
                        foreach (Entity entity2 in entityCollection.Entities)
                        {
                            try
                            {
                                Entity entity3 = new Entity(entity2.LogicalName, entity2.Id);
                                entity3["msdyn_paymentamount"] = new Money(0);
                                service.Update(entity3);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }

                        }
                        Entity entity1 = new Entity(entity.LogicalName, entity.Id);
                        entity1["msdyn_paymentamount"] = new Money(decimal.Parse(recivedAmount));
                        service.Update(entity1);

                        req2 = new SetStateRequest();
                        req2.State = new OptionSetValue(0);
                        req2.Status = new OptionSetValue(910590000);
                        req2.EntityMoniker = entity.ToEntityReference();
                        res = (SetStateResponse)service.Execute(req2);

                    }
                    else if (StatusPay.ToLower() == "failure".ToLower() || StatusPay.ToLower() == "not found")
                    {
                        Console.WriteLine("Payment Status !");
                        Console.WriteLine("Payment is " + StatusPay);

                        Entity entity1 = new Entity(entity.LogicalName, entity.Id);
                        entity1["msdyn_paymentamount"] = new Money(0);
                        service.Update(entity1);

                        SetStateRequest req1 = new SetStateRequest();
                        req1.State = new OptionSetValue(0);
                        req1.Status = new OptionSetValue(910590001);
                        req1.EntityMoniker = entity.ToEntityReference();
                        service.Execute(req1);
                    }
                }
                else if (entity.LogicalName == "msdyn_workorder")
                {
                    if (StatusPay.ToLower() == "success".ToLower())
                    {
                        Entity jobs = new Entity(entity.LogicalName, entity.Id);
                        jobs["hil_receiptnumber"] = bank_ref_number;
                        service.Update(jobs);
                        Console.WriteLine("Payment Status !");
                        Console.WriteLine("Payment received successfully.");
                    }
                    else if (StatusPay.ToLower() == "failure".ToLower())
                    {
                        Console.WriteLine("Payment Status !");
                        Console.WriteLine("Payment is " + StatusPay);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error !");
                Console.WriteLine("D365 Internal Error " + ex.Message);
            }
            Console.WriteLine("c");
        }
        static public string getTransactionID(Entity entity, IOrganizationService service)
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
    public class SendPaymentLink
    {
        public static ITracingService tracingService = null;
        public static IntegrationConfiguration GetIntegrationConfiguration(IOrganizationService _service, string name)
        {
            try
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
            catch (Exception ex)
            {
                throw new Exception("Error : " + ex.Message);
            }
        }
    }
    public class IntegrationConfiguration
    {
        public string url { get; set; }
        public string userName { get; set; }
        public string password { get; set; }
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
    public class SendURLD365Request
    {

        public string InvoiceId { get; set; }

        public string mobile { get; set; }

        public string Amount { get; set; }
    }
    public class StatusRequest
    {

        public string PROJECT { get; set; }

        public string command { get; set; }

        public string var1 { get; set; }

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
    public class PaymentStatusD365Response
    {

        public string Status { get; set; }
    }
}
