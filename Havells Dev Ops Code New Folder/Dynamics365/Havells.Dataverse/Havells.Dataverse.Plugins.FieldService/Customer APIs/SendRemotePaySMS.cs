
using Newtonsoft.Json;
using System;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Web.Script.Serialization;


namespace Havells.Dataverse.Plugins.FieldService.Customer_APIs
{
    public class SendRemotePaySMS : IPlugin
    {
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            if (context.InputParameters.Contains("JobId") && context.InputParameters["JobId"] is string
            && context.InputParameters.Contains("MobileNumber") && context.InputParameters["MobileNumber"] is string && context.InputParameters.Contains("Amount") && context.InputParameters["Amount"] is string)

            {
                tracingService.Trace("Execution Start");
                string JobId = context.InputParameters["JobId"].ToString();
                tracingService.Trace("Job Record ID " + JobId);
                string MobileNumber = context.InputParameters["MobileNumber"].ToString();
                string Amount = context.InputParameters["Amount"].ToString();
                if (string.IsNullOrWhiteSpace(JobId))
                {
                    context.OutputParameters["Response"] = "Invalid Job GUID";
                    context.OutputParameters["Status"] = false;
                }
                else if (string.IsNullOrWhiteSpace(MobileNumber))
                {
                    context.OutputParameters["Response"] = "Invalid Mobile Number";
                    context.OutputParameters["Status"] = false;
                }
                else if (string.IsNullOrWhiteSpace(Amount))
                {
                    context.OutputParameters["Response"] = "Invalid Amount";
                    context.OutputParameters["Status"] = false;
                }

                else
                {
                    SendSMS(JobId, MobileNumber, Amount, context, service);
                }
            }
        }
        public void SendSMS(string JobId, string MobileNumber, string Amount, IPluginExecutionContext context, IOrganizationService service)
        {
            try
            {
                tracingService.Trace("step-1");
                SendPaymentUrlRequest req = new SendPaymentUrlRequest();
                String comm = "create_invoice";
                req.PROJECT = "D365";
                req.command = comm.Trim();
                RemotePaymentLinkDetails remotePaymentLinkDetails = new RemotePaymentLinkDetails();

                tracingService.Trace("step-2");
                Entity job = service.Retrieve("msdyn_workorder", new Guid(JobId), new ColumnSet(true));
                QueryExpression Query = new QueryExpression("hil_paymentstatus");
                Query.ColumnSet = new ColumnSet("hil_url");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, job.GetAttributeValue<string>("msdyn_name"));
                EntityCollection Found = service.RetrieveMultiple(Query);
                if (Found.Entities.Count > 0)
                {
                    context.OutputParameters["Message"] = "Payment link Already send to Customer";
                    context.OutputParameters["PaymentLink"] = Found.Entities[0].GetAttributeValue<string>("hil_url");
                    tracingService.Trace("step-3");
                    context.OutputParameters["Status"] = false;
                }

                string state = job.Contains("hil_state") ? job.GetAttributeValue<EntityReference>("hil_state").Name.ToString() : string.Empty;
                string jobId = job.Contains("msdyn_workorder") ? job.GetAttributeValue<EntityReference>("msdyn_workorder").Id.ToString() : string.Empty;
                string zip = string.Empty;

                string address = job.Contains("hil_fulladdress") ? job.GetAttributeValue<String>("hil_fulladdress").ToString() : string.Empty;
                zip = job.Contains("hil_pincode") ? job.GetAttributeValue<EntityReference>("hil_pincode").Name.ToString() : string.Empty;
                address = ReplaceSpecialCharacters(address);
                //remotePaymentLinkDetails.amount = job.Contains("hil_receiptamount") ? decimal.Round(job.GetAttributeValue<Money>("hil_receiptamount").Value, 2).ToString() : "0.00";
                remotePaymentLinkDetails.amount = Amount;
                Entity ent = service.Retrieve("hil_branch", job.GetAttributeValue<EntityReference>("hil_branch").Id, new ColumnSet("hil_mamorandumcode"));
                string _mamorandumCode = ent.Attributes.Contains("hil_mamorandumcode") ? ent.GetAttributeValue<string>("hil_mamorandumcode") : "";
                string _txnId = "D365_" + job.GetAttributeValue<string>("msdyn_name");
                remotePaymentLinkDetails.txnid = _txnId;
                remotePaymentLinkDetails.firstname = job.Contains("hil_customerref") ? job.GetAttributeValue<EntityReference>("hil_customerref").Name.ToString() : string.Empty;
                remotePaymentLinkDetails.email = job.Contains("hil_email") ? job.GetAttributeValue<String>("hil_email").ToString() : "abc@gmail.com";
                remotePaymentLinkDetails.phone = MobileNumber.Substring(MobileNumber.Length - 10);
                remotePaymentLinkDetails.address1 = address.Length > 99 ? address.Substring(0, 99) : address;
                remotePaymentLinkDetails.state = state;
                remotePaymentLinkDetails.country = "India";
                remotePaymentLinkDetails.template_id = "1";
                remotePaymentLinkDetails.productinfo = _mamorandumCode;
                remotePaymentLinkDetails.validation_period = "24";
                remotePaymentLinkDetails.send_email_now = "0";
                remotePaymentLinkDetails.send_sms = "1";
                remotePaymentLinkDetails.time_unit = "H";
                remotePaymentLinkDetails.zipcode = zip;
                req.RemotePaymentLinkDetails = remotePaymentLinkDetails;
                var json = JsonConvert.SerializeObject(req);
                var data = new StringContent(json, Encoding.UTF8, "application/json");
                string DLCredential = "D365_Havells" + ":" + "PRDD365@1234";
                string apiUrl = "https://middleware.havells.com:50001/RESTAdapter/PayUBiz/TransactionAPI";
                string _authInfo = Convert.ToBase64String(Encoding.ASCII.GetBytes(DLCredential));
                HttpClient client = new HttpClient();
                var byteArray = Encoding.ASCII.GetBytes(DLCredential);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                HttpResponseMessage response = client.PostAsync(apiUrl, data).Result;
                tracingService.Trace("step-4");
                tracingService.Trace(JsonConvert.SerializeObject(response.Content));
                SendPaymentUrlResponse obj = (new JavaScriptSerializer()).Deserialize<SendPaymentUrlResponse>(response.Content.ReadAsStringAsync().Result);

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
                    statusPayment["hil_job"] = new EntityReference("msdyn_workorder", new Guid(JobId));
                    service.Create(statusPayment);
                    #region Updating Job Payment Link Sent field
                    Entity _updateJob = new Entity("msdyn_workorder", new Guid(JobId));
                    _updateJob["hil_paymentlinksent"] = true;
                    service.Update(_updateJob);
                    #endregion

                    context.OutputParameters["Message"] = obj.Status;
                    context.OutputParameters["PaymentLink"] = obj.URL;
                    context.OutputParameters["Status"] = true;
                }
                else
                {
                    context.OutputParameters["Message"] = obj.msg;
                    context.OutputParameters["PaymentLink"] = null;
                    context.OutputParameters["Status"] = false;
                }
            }
            catch (Exception ex)
            {
                context.OutputParameters["Message"] = "D365 Internal Error " + ex.Message;
                context.OutputParameters["PaymentLink"] = null;
                context.OutputParameters["Status"] = false;
            }
        }
        private string ReplaceSpecialCharacters(string _inputStr)
        {
            return Regex.Replace(_inputStr, @"[^0-9a-zA-Z]+", "");
        }
    }
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



