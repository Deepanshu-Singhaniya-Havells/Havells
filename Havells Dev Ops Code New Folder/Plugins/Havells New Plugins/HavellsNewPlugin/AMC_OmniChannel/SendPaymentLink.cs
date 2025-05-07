using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.AMC_OmniChannel
{
    enum Operation
    {
        sendPayment,
        GetSatus
    }
    public class SendPaymentLink : IPlugin
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
                if (context.InputParameters.Contains("InvoiceId") && context.InputParameters["InvoiceId"] is string)
                {
                    var InvoiceId = context.InputParameters["InvoiceId"].ToString();
                    decimal amount = 0;
                    SendURLD365Request reqParm = new SendURLD365Request();
                    reqParm.InvoiceId = InvoiceId;
                    SendPaymentUrlResponse sendPaymentUrlResponse = new SendPaymentUrlResponse();
                    try
                    {
                        if (reqParm.InvoiceId == null)
                        {
                            context.OutputParameters["Status"] = "Failed !!";
                            context.OutputParameters["Message"] = "Invalid Invoice GUID";
                        }
                        else
                        {
                            Entity FoundInvoice = service.Retrieve("invoice", new Guid(InvoiceId), new ColumnSet(true));
                            EntityReference customerref = FoundInvoice.GetAttributeValue<EntityReference>("customerid");

                            Entity customer = service.Retrieve(customerref.LogicalName, customerref.Id, new ColumnSet("mobilephone", "emailaddress1"));

                            string mobile = customer.Contains("mobilephone") ? customer.GetAttributeValue<String>("mobilephone").ToString() : null;
                            string email = customer.Contains("emailaddress1") ? customer.GetAttributeValue<String>("emailaddress1").ToString() : null;

                            if (mobile == null)
                            {
                                context.OutputParameters["Status"] = "Data Validation Alert !";
                                context.OutputParameters["Message"] = "Customer Mobile Number does not exist in Customer Master.";
                                return;
                            }
                            else
                            {
                                reqParm.Amount = FoundInvoice.Contains("hil_receiptamount") ? FoundInvoice.GetAttributeValue<Money>("hil_receiptamount").Value.ToString() : "0";
                                if (decimal.Parse(reqParm.Amount) < 1)
                                {
                                    context.OutputParameters["Status"] = "Data Validation Alert !";
                                    context.OutputParameters["Message"] = "Receipt Amount must be greater than 0.";
                                    return;
                                }
                                else
                                {
                                    QueryExpression Query = new QueryExpression("msdyn_paymentdetail");
                                    Query.ColumnSet = new ColumnSet(false);
                                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                                    Query.Criteria.AddCondition("msdyn_invoice", ConditionOperator.Equal, new Guid(InvoiceId));
                                    Query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 910590002);
                                    EntityCollection FoundPaymentDetails = service.RetrieveMultiple(Query);
                                    if (FoundPaymentDetails.Entities.Count > 0)
                                    {
                                        context.OutputParameters["Status"] = "Data Validation Alert !";
                                        context.OutputParameters["Message"] = "Privious payment link is not yet expired.";
                                        return;
                                    }

                                    SendPaymentUrlRequest req = new SendPaymentUrlRequest();
                                    String comm = "create_invoice";
                                    req.PROJECT = "D365";
                                    req.command = comm.Trim();

                                    RemotePaymentLinkDetails remotePaymentLinkDetails = new RemotePaymentLinkDetails();

                                    string _txnId = getTransactionID(service, new Guid(InvoiceId), FoundInvoice.GetAttributeValue<string>("name"));


                                    EntityReference AddressId = FoundInvoice.GetAttributeValue<EntityReference>("hil_address");
                                    Entity AddressCol = service.Retrieve(AddressId.LogicalName, AddressId.Id, new ColumnSet("hil_state", "hil_businessgeo", "hil_pincode", "hil_branch"));

                                    String state = AddressCol.Contains("hil_state") ? AddressCol.GetAttributeValue<EntityReference>("hil_state").Name.ToString() : string.Empty;
                                    String zip = string.Empty;

                                    string address = AddressCol.Contains("hil_businessgeo") ? AddressCol.GetAttributeValue<EntityReference>("hil_businessgeo").Name.ToString() : string.Empty;
                                    zip = AddressCol.Contains("hil_pincode") ? AddressCol.GetAttributeValue<EntityReference>("hil_pincode").Name.ToString() : string.Empty;

                                    string city = string.Empty;
                                    decimal amt = Convert.ToDecimal(reqParm.Amount);
                                    remotePaymentLinkDetails.amount = Math.Round(amt, 2).ToString();
                                    amount = decimal.Parse(remotePaymentLinkDetails.amount);
                                    Entity ent = service.Retrieve("hil_branch", AddressCol.GetAttributeValue<EntityReference>("hil_branch").Id, new ColumnSet("hil_mamorandumcode"));

                                    string _mamorandumCode = "";
                                    if (ent.Attributes.Contains("hil_mamorandumcode"))
                                    {
                                        _mamorandumCode = ent.GetAttributeValue<string>("hil_mamorandumcode");
                                    }


                                    //_txnId = "S365_" + FoundInvoice.GetAttributeValue<string>("name");
                                    remotePaymentLinkDetails.txnid = _txnId;
                                    remotePaymentLinkDetails.firstname = FoundInvoice.Contains("customerid") ? FoundInvoice.GetAttributeValue<EntityReference>("customerid").Name.ToString() : string.Empty;
                                    remotePaymentLinkDetails.email = customer.Contains("emailaddress1") ? customer.GetAttributeValue<String>("emailaddress1").ToString() : "abc@gmail.com";
                                    remotePaymentLinkDetails.phone = customer.Contains("mobilephone") ? customer.GetAttributeValue<String>("mobilephone").ToString() : "";

                                    remotePaymentLinkDetails.address1 = address.Length > 99 ? address.Substring(0, 99) : address;
                                    remotePaymentLinkDetails.state = state;
                                    remotePaymentLinkDetails.country = "India";
                                    remotePaymentLinkDetails.template_id = "1";
                                    remotePaymentLinkDetails.productinfo = _mamorandumCode; //"B2C_PAYUBIZ_TEST_SMS";
                                    remotePaymentLinkDetails.validation_period = "24";
                                    remotePaymentLinkDetails.send_email_now = "1";
                                    remotePaymentLinkDetails.send_sms = "1";
                                    remotePaymentLinkDetails.time_unit = "H";
                                    remotePaymentLinkDetails.zipcode = zip;
                                    req.RemotePaymentLinkDetails = remotePaymentLinkDetails;

                                    #region logrequest             
                                    Entity intigrationTrace = new Entity("hil_integrationtrace");
                                    intigrationTrace["hil_entityname"] = FoundInvoice.LogicalName;
                                    intigrationTrace["hil_entityid"] = FoundInvoice.Id.ToString();
                                    intigrationTrace["hil_request"] = JsonConvert.SerializeObject(req);
                                    intigrationTrace["hil_name"] = FoundInvoice.GetAttributeValue<string>("name");
                                    Guid intigrationTraceID = service.Create(intigrationTrace);
                                    #endregion logrequest

                                    IntegrationConfiguration inconfig = GetIntegrationConfiguration(service, "Send Payment Link");

                                    var client = new RestClient(inconfig.url);
                                    client.Timeout = -1;
                                    var request = new RestRequest(Method.POST);

                                    string authInfo = inconfig.userName + ":" + inconfig.password;
                                    authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
                                    request.AddHeader("Authorization", authInfo);
                                    request.AddHeader("Content-Type", "application/json");
                                    request.AddParameter("application/json", JsonConvert.SerializeObject(req), ParameterType.RequestBody);
                                    IRestResponse response = client.Execute(request);

                                    dynamic obj = JsonConvert.DeserializeObject<SendPaymentUrlResponse>(response.Content);
                                    //throw new InvalidPluginExecutionException("responseJson " + response.Content);
                                    #region logresponse
                                    Entity intigrationTraceUp = new Entity("hil_integrationtrace");
                                    intigrationTraceUp["hil_response"] = response.Content == "" ? response.ErrorMessage : response.Content;
                                    intigrationTraceUp.Id = intigrationTraceID;
                                    service.Update(intigrationTraceUp);
                                    #endregion logresponse

                                    if (obj.msg == null)
                                    {
                                        string url = obj.URL;
                                        string[] invoicenumber = url.Split('=');
                                        Entity statusPayment = new Entity("hil_paymentstatus");
                                        statusPayment["hil_name"] = obj.Transaction_Id;
                                        statusPayment["hil_url"] = obj.URL;
                                        statusPayment["hil_statussendurl"] = obj.Status;
                                        statusPayment["hil_email_id"] = obj.Email_Id;
                                        statusPayment["hil_phone"] = obj.Phone;
                                        statusPayment["hil_invoiceid"] = invoicenumber[1];
                                        service.Create(statusPayment);

                                        Entity entinvoice = new Entity(FoundInvoice.LogicalName, FoundInvoice.Id);
                                        entinvoice["statuscode"] = new OptionSetValue(4);
                                        service.Update(entinvoice);
                                        createPaymentDetailsRecord(service, obj.Transaction_Id, FoundInvoice.ToEntityReference(), amount, FoundInvoice.GetAttributeValue<EntityReference>("ownerid"));
                                        context.OutputParameters["Status"] = "Confirmation !";
                                        context.OutputParameters["Message"] = "Payment link has been sent sucessfully.";
                                    }
                                    else
                                    {
                                        context.OutputParameters["Status"] = "Error !";
                                        context.OutputParameters["Message"] = obj.msg;
                                    }
                                    //}
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        context.OutputParameters["Status"] = "Failed !!";
                        context.OutputParameters["Message"] = "D365 Internal Error " + ex.Message;
                    }
                }
            }
            catch (Exception ex)
            {
                context.OutputParameters["Status"] = "Failed !!";
                context.OutputParameters["Message"] = "D365 Internal Error " + ex.Message;
            }
        }
        void createPaymentDetailsRecord(IOrganizationService service, string Transaction_Id, EntityReference invoice, Decimal _paymentamount, EntityReference owner)
        {
            try
            {
                Entity payment = new Entity("msdyn_payment");
                payment["msdyn_name"] = Transaction_Id;
                payment["msdyn_amount"] = new Money(_paymentamount);
                payment["msdyn_paymenttype"] = new OptionSetValue(690970003);
                payment["ownerid"] = owner;
                Entity _PaymantDetails = new Entity("msdyn_paymentdetail");
                _PaymantDetails["msdyn_name"] = Transaction_Id;
                _PaymantDetails["msdyn_invoice"] = invoice;
                _PaymantDetails["msdyn_paymentamount"] = new Money(_paymentamount);
                _PaymantDetails["statuscode"] = new OptionSetValue(910590002);
                _PaymantDetails["ownerid"] = owner;
                _PaymantDetails["msdyn_payment"] = new EntityReference("msdyn_payment", service.Create(payment));
                //_PaymantDetails["transactioncurrencyid"] = new EntityReference("transactioncurrencies", new Guid("68A6A9CA-6BEB-E811-A96C-000D3AF05828"));
                service.Create(_PaymantDetails);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        public static string getTransactionID(IOrganizationService service, Guid InvoiceId, string InvoiceNum)
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

            //else if (FoundPaymentDetails.Entities.Count == 0)
            //{
            //    Query = new QueryExpression("msdyn_paymentdetail");
            //    Query.ColumnSet = new ColumnSet("statuscode");
            //    Query.Criteria = new FilterExpression(LogicalOperator.And);
            //    Query.Criteria.AddCondition("msdyn_invoice", ConditionOperator.Equal, InvoiceId);
            //    Query.AddOrder("createdon", OrderType.Descending);
            //    //Query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 910590001);
            //    FoundPaymentDetails = service.RetrieveMultiple(Query);
            //    if (FoundPaymentDetails.Entities.Count == 0)
            //    {
            //        Query = new QueryExpression("msdyn_paymentdetail");
            //        Query.ColumnSet = new ColumnSet(false);
            //        Query.Criteria = new FilterExpression(LogicalOperator.And);
            //        Query.Criteria.AddCondition("msdyn_invoice", ConditionOperator.Equal, InvoiceId);
            //        FoundPaymentDetails = service.RetrieveMultiple(Query);
            //        transactionID = "S365_" + InvoiceNum + "_" + FoundPaymentDetails.Entities.Count;
            //    }
            //    else
            //        throw new InvalidPluginExecutionException("Please check status on Payment Details Section.");
            //}
            //else if (FoundPaymentDetails.Entities.Count > 1)
            //{
            //    throw new InvalidPluginExecutionException("More than one Payment Details found. Please contact to Administrator.");
            //}
            return transactionID;
        }
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
