using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RestSharp;
using RestSharp.Serialization.Json;
using System.Runtime.Serialization;
using System.Text;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class SendPaymentUrlRequest
    {
        [DataMember]
        public string PROJECT { get; set; }
        [DataMember]
        public string command { get; set; }
        [DataMember]
        public RemotePaymentLinkDetails RemotePaymentLinkDetails { get; set; }
    }
    [DataContract]
    public class RemotePaymentLinkDetails
    {
        [DataMember]
        public string amount { get; set; }
        [DataMember]
        public string txnid { get; set; }
        [DataMember]
        public string productinfo { get; set; }
        [DataMember]
        public string firstname { get; set; }
        [DataMember]
        public string email { get; set; }
        [DataMember]
        public string phone { get; set; }
        [DataMember]
        public string address1 { get; set; }
        [DataMember]
        public string city { get; set; }
        [DataMember]
        public string state { get; set; }
        [DataMember]
        public string country { get; set; }
        [DataMember]
        public string zipcode { get; set; }
        [DataMember]
        public string template_id { get; set; }
        [DataMember]
        public string validation_period { get; set; }
        [DataMember]
        public string send_email_now { get; set; }
        [DataMember]
        public string send_sms { get; set; }
        [DataMember]
        public string time_unit { get; set; }
    }
    [DataContract]
    public class SendPaymentUrlResponse
    {
        [DataMember]
        public string Email_Id { get; set; }
        [DataMember]
        public string Transaction_Id { get; set; }
        [DataMember]
        public string URL { get; set; }
        [DataMember]
        public string Status { get; set; }
        [DataMember]
        public string Phone { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string msg { get; set; }
    }
    [DataContract]
    public class SendURLD365Request
    {
        [DataMember]
        public string jobId { get; set; }
        [DataMember]
        public string mobile { get; set; }
        [DataMember]
        public string Amount { get; set; }
    }
    [DataContract]
    public class StatusRequest
    {
        [DataMember]
        public string PROJECT { get; set; }
        [DataMember]
        public string command { get; set; }
        [DataMember]
        public string var1 { get; set; }

    }
    [DataContract]
    public class TransactionDetail
    {
        [DataMember]
        public string mihpayid { get; set; }
        [DataMember]
        public string request_id { get; set; }
        [DataMember]
        public string bank_ref_num { get; set; }
        [DataMember]
        public string amt { get; set; }
        [DataMember]
        public string transaction_amount { get; set; }
        [DataMember]
        public string txnid { get; set; }
        [DataMember]
        public string additional_charges { get; set; }
        [DataMember]
        public string productinfo { get; set; }
        [DataMember]
        public string firstname { get; set; }
        [DataMember]
        public string bankcode { get; set; }
        [DataMember]
        public string udf1 { get; set; }
        [DataMember]
        public string udf3 { get; set; }
        [DataMember]
        public string udf4 { get; set; }
        [DataMember]
        public string udf5 { get; set; }
        [DataMember]
        public string field2 { get; set; }
        [DataMember]
        public string field9 { get; set; }
        [DataMember]
        public string error_code { get; set; }
        [DataMember]
        public string addedon { get; set; }
        [DataMember]
        public string payment_source { get; set; }
        [DataMember]
        public string card_type { get; set; }
        [DataMember]
        public string error_Message { get; set; }
        [DataMember]
        public string net_amount_debit { get; set; }
        [DataMember]
        public string disc { get; set; }
        [DataMember]
        public string mode { get; set; }
        [DataMember]
        public string PG_TYPE { get; set; }
        [DataMember]
        public string card_no { get; set; }
        [DataMember]
        public string udf2 { get; set; }
        [DataMember]
        public string status { get; set; }
        [DataMember]
        public string unmappedstatus { get; set; }
        [DataMember]
        public string Merchant_UTR { get; set; }
        [DataMember]
        public string Settled_At { get; set; }
    }
    [DataContract]
    public class StatusResponse
    {
        [DataMember]
        public int status { get; set; }
        [DataMember]
        public string msg { get; set; }
        [DataMember]
        public List<TransactionDetail> transaction_details { get; set; }
    }
    [DataContract]
    public class PaymentStatusD365Response
    {
        [DataMember]
        public string Status { get; set; }
    }
    [DataContract]
    public class RemotePay
    {
        private const string URL = "https://middlewareqa.havells.com:50001/RESTAdapter/PayUBiz/TransactionAPI";
        public SendPaymentUrlResponse SendSMS(SendURLD365Request reqParm)
        {
            SendPaymentUrlResponse sendPaymentUrlResponse = new SendPaymentUrlResponse();
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                SendPaymentUrlRequest req = new SendPaymentUrlRequest();
                String comm = "create_invoice";
                req.PROJECT = "D365";
                req.command = comm.Trim();
                RemotePaymentLinkDetails remotePaymentLinkDetails = new RemotePaymentLinkDetails();
                if (service != null)
                {
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

                        //if (address != string.Empty)
                        //{
                        //    zip = address.Substring(address.Length - 6);
                        //}
                        string city = string.Empty;
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
                        //}
                        remotePaymentLinkDetails.amount = reqParm.Amount;

                        Entity ent = service.Retrieve("hil_branch", job.GetAttributeValue<EntityReference>("hil_branch").Id, new ColumnSet("hil_mamorandumcode"));
                        string _txnId = string.Empty;
                        string _mamorandumCode = "";
                        if (ent.Attributes.Contains("hil_mamorandumcode"))
                        {
                            _mamorandumCode = ent.GetAttributeValue<string>("hil_mamorandumcode");
                        }
                        
                        _txnId = "D365" + job.GetAttributeValue<string>("msdyn_name").Replace("-","");
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
                        var client = new RestClient("https://middleware.havells.com:50001/RESTAdapter/PayUBiz/TransactionAPI");
                        client.Timeout = -1;
                        var request = new RestRequest(Method.POST);
                        string authInfo = "D365_HAVELLS" + ":" + "PRDD365@1234";
                        authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
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
                IOrganizationService service = ConnectToCRM.GetOrgService();
                StatusRequest req = new StatusRequest();
                jobId = "D365" + jobId;
                req.PROJECT = "D365";
                req.command = "verify_payment";
                req.var1 = jobId;
                var client = new RestClient("https://middlewareqa.havells.com:50001/RESTAdapter/PayUBiz/TransactionAPI");
                client.Timeout = -1;
                var request = new RestRequest(Method.POST);
                request.AddHeader("Authorization", "Basic RDM2NV9IQVZFTExTOlFBRDM2NUAxMjM0");
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
                string bank_refnum = null;
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
                        bank_refnum = obj.transaction_details[0].bank_ref_num;
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
                    _updateJob["hil_receiptnumber"] = bank_refnum;
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
        
    }
}

