using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RestSharp;
using RestSharp.Serialization.Json;
using System.Runtime.Serialization;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using System.Runtime.Remoting.Contexts;
using System.Runtime.Remoting.Services;

namespace Havells.Crm.WebJob
{

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

        public string jobId { get; set; }

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

    public class RemotePay
    {
        private const string URL = "https://middleware.havells.com:50001/RESTAdapter/PayUBiz/TransactionAPI";
        public SendPaymentUrlResponse SendSMS(SendURLD365Request reqParm)
        {
            SendPaymentUrlResponse sendPaymentUrlResponse = new SendPaymentUrlResponse();
            try
            {
                IOrganizationService service = Program.CreateCRMConnection();
                SendPaymentUrlRequest req = new SendPaymentUrlRequest();
                String comm = "create_invoice";
                req.PROJECT = "D365";
                req.command = comm.Trim();
                RemotePaymentLinkDetails remotePaymentLinkDetails = new RemotePaymentLinkDetails();
                if (service != null)
                {
                    //HOLD Remotepay
                    //sendPaymentUrlResponse.StatusCode = "HOLD";
                    //return sendPaymentUrlResponse;

                    if (reqParm.jobId == null)
                    {
                        sendPaymentUrlResponse.StatusCode = "Invalid Job GUID";
                    }
                    else
                    {
                        Entity job = service.Retrieve(msdyn_workorder.EntityLogicalName, new Guid(reqParm.jobId), new ColumnSet(true));

                        QueryExpression Query = new QueryExpression("hil_paymentstatus");
                        Query.ColumnSet = new ColumnSet("hil_url");
                        Query.Criteria = new FilterExpression(LogicalOperator.And);
                        Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, job.GetAttributeValue<string>("msdyn_name"));
                        EntityCollection Found = service.RetrieveMultiple(Query);
                        if (Found.Entities.Count > 0)
                        {
                            sendPaymentUrlResponse.StatusCode = "SMS/Email is Allready send to Customer";
                            sendPaymentUrlResponse.URL = Found[0].GetAttributeValue<string>("hil_url");
                        }

                        //decimal abc = job.GetAttributeValue<Money>("hil_receiptamount").Value;

                        String state = job.Contains("hil_state") ? job.GetAttributeValue<EntityReference>("hil_state").Name.ToString() : string.Empty;
                        String zip = string.Empty;

                        string address = job.Contains("hil_fulladdress") ? job.GetAttributeValue<String>("hil_fulladdress").ToString() : string.Empty;
                        zip = job.Contains("hil_pincode") ? job.GetAttributeValue<EntityReference>("hil_pincode").Name.ToString() : string.Empty;
                        address = ReplaceSpecialCharacters(address);
                        //if (address != string.Empty)
                        //{
                        //    zip = address.Substring(address.Length - 6);
                        //}
                        //string city = string.Empty;
                        //if (job.Contains("hil_city"))
                        //{
                        //    if (state.ToLower() == "delhi")
                        //    {
                        //        city = "Delhi";
                        //    }
                        //    else
                        //    {
                        //        String cit = job.GetAttributeValue<EntityReference>("hil_city").Name.ToString();
                        //        city = cit.Substring(0, cit.Length - 3);
                        //    }
                        //    city = ReplaceSpecialCharacters(city);
                        //}
                        remotePaymentLinkDetails.amount = reqParm.Amount;

                        Entity ent = service.Retrieve("hil_branch", job.GetAttributeValue<EntityReference>("hil_branch").Id, new ColumnSet("hil_mamorandumcode"));
                        string _txnId = string.Empty;
                        string _mamorandumCode = "";
                        if (ent.Attributes.Contains("hil_mamorandumcode"))
                        {
                            _mamorandumCode = ent.GetAttributeValue<string>("hil_mamorandumcode");
                        }

                        _txnId = "D365_" + job.GetAttributeValue<string>("msdyn_name");
                        remotePaymentLinkDetails.txnid = _txnId;
                        remotePaymentLinkDetails.firstname = job.Contains("hil_customerref") ? job.GetAttributeValue<EntityReference>("hil_customerref").Name.ToString() : string.Empty;
                        remotePaymentLinkDetails.email = job.Contains("hil_email") ? job.GetAttributeValue<String>("hil_email").ToString() : "abc@gmail.com";
                        remotePaymentLinkDetails.phone = reqParm.mobile;
                        //remotePaymentLinkDetails.city = city;//job.Contains("hil_city") ? ((state.ToLower() == "delhi") ? "Delhi" : job.GetAttributeValue<EntityReference>("hil_city").Name.ToString()) : string.Empty;
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
                        //var client = new RestClient("https://middleware.havells.com:50001/RESTAdapter/PayUBiz/TransactionAPI");
                        var client = new RestClient("https://middlewareqa.havells.com:50001/RESTAdapter/PayUBiz/TransactionAPI");
                        client.Timeout = -1;

                        //string authInfo = "D365_Havells" + ":" + "PRDD365@1234";
                        string authInfo = "D365_Havells" + ":" + "QAD365@1234";
                        authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));

                        var request = new RestRequest(Method.POST);
                        request.AddHeader("Authorization", authInfo);
                        //request.AddHeader("Authorization", "Basic RDM2NV9IQVZFTExTOlFBRDM2NUAxMjM0");
                        request.AddHeader("Content-Type", "application/json");
                        request.AddHeader("Cookie", "saplb_*=(J2EE2717920)2717950; JSESSIONID=7fOj-tgnbYBRVBihJMBX9THzyTG3dgH-eCkA_SAPa_yX9TL_PrH5RR_PrxfO7kbO");
                        request.AddParameter("application/json", Newtonsoft.Json.JsonConvert.SerializeObject(req), ParameterType.RequestBody);
                        IRestResponse response = client.Execute(request);

                        var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<SendPaymentUrlResponse>(response.Content);
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
                            statusPayment["hil_job"] = new EntityReference(msdyn_workorder.EntityLogicalName, new Guid(reqParm.jobId));
                            service.Create(statusPayment);

                            #region Updating Job Payment Link Sent field
                            Entity _updateJob = new Entity(msdyn_workorder.EntityLogicalName, new Guid(reqParm.jobId));
                            _updateJob["hil_paymentlinksent"] = true;
                            service.Update(_updateJob);
                            #endregion
                            sendPaymentUrlResponse.StatusCode = obj.Status;
                            sendPaymentUrlResponse.URL = obj.URL;
                        }
                        else
                            sendPaymentUrlResponse.StatusCode = obj.msg;
                    }
                }
            }
            catch (Exception ex)
            {
                sendPaymentUrlResponse.StatusCode = "D365 Internal Error " + ex.Message;
            }
            return sendPaymentUrlResponse;
        }
        public PaymentStatusD365Response getPaymentStatus(String jobId)
        {
            PaymentStatusD365Response var = new PaymentStatusD365Response();
            try
            {
                IOrganizationService service = Program.CreateCRMConnection();
                jobId = "D365_" + jobId;
                StatusRequest req = new StatusRequest();
                req.PROJECT = "D365";
                req.command = "verify_payment";
                req.var1 = jobId;
                //var client = new RestClient("https://middleware.havells.com:50001/RESTAdapter/PayUBiz/TransactionAPI");
                var client = new RestClient("https://middlewareqa.havells.com:50001/RESTAdapter/PayUBiz/TransactionAPI");
                client.Timeout = -1;
                var request = new RestRequest(Method.POST);

                //string authInfo = "D365_Havells" + ":" + "PRDD365@1234";
                string authInfo = "D365_Havells" + ":" + "QAD365@1234";
                authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));

                request.AddHeader("Authorization", authInfo);
                //request.AddHeader("Authorization", "Basic RDM2NV9IQVZFTExTOlFBRDM2NUAxMjM0");
                request.AddHeader("Content-Type", "application/json");
                request.AddHeader("Cookie", "saplb_*=(J2EE2717920)2717950; JSESSIONID=7fOj-tgnbYBRVBihJMBX9THzyTG3dgH-eCkA_SAPa_yX9TL_PrH5RR_PrxfO7kbO");
                request.AddParameter("application/json", Newtonsoft.Json.JsonConvert.SerializeObject(req), ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                //Console.WriteLine(response.Content);
                //Console.ReadLine();

                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<StatusResponse>(response.Content);
                QueryExpression Query = new QueryExpression("hil_paymentstatus");
                Query.ColumnSet = new ColumnSet("hil_job");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, jobId);
                EntityCollection Found = service.RetrieveMultiple(Query);
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
                        statusPayment["hil_field9"] = obj.transaction_details[0].field9;
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

                    if (obj.transaction_details[0].status == "Not Found")
                    {
                        var.Status = "not initiated";
                    }
                    else
                        var.Status = obj.transaction_details[0].status;

                    #region Updating Job Payment Status field
                    Entity _updateJob = new Entity(msdyn_workorder.EntityLogicalName, Found.Entities[0].GetAttributeValue<EntityReference>("hil_job").Id);
                    OptionSetValue _opValue = null;
                    if (var.Status == "not initiated")
                    {
                        _opValue = new OptionSetValue(1);
                    }
                    else if (var.Status == "success")
                    {
                        _opValue = new OptionSetValue(2);
                    }
                    else if (var.Status == "pending")
                    {
                        _opValue = new OptionSetValue(3);
                    }
                    else
                        _opValue = new OptionSetValue(4);

                    _updateJob["hil_paymentstatus"] = _opValue;
                    service.Update(_updateJob);
                    #endregion
                }
                else
                {
                    var.Status = "Transaction ID not found.";
                }
            }
            catch (Exception ex)
            {
                var.Status = "D365 Internal Error " + ex;
            }
            return var;
        }
        private string ReplaceSpecialCharacters(string _inputStr)
        {
            return Regex.Replace(_inputStr, @"[^0-9a-zA-Z]+", "");
        }
      

    }
}
