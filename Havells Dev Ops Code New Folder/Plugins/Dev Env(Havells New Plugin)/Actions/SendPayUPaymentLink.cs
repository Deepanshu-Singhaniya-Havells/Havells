using HavellsNewPlugin.AMC_OmniChannel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.Serialization;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace HavellsNewPlugin.Actions
{
    public class SendPayUPaymentLink : IPlugin
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

            if (context.InputParameters.Contains("JobId") && context.InputParameters["JobId"] is string
                && context.InputParameters.Contains("MobileNumber") && context.InputParameters["MobileNumber"] is string
                && context.Depth == 1)
            {
                tracingService.Trace("Execution Start");
                string JobId = context.InputParameters["JobId"].ToString();
                tracingService.Trace("Job Record ID " + JobId);
                string MobileNumber = context.InputParameters["MobileNumber"].ToString();
                if (string.IsNullOrWhiteSpace(JobId) || string.IsNullOrWhiteSpace(MobileNumber))
                {
                    context.OutputParameters["Message"] = "Invalid Job Number";
                    context.OutputParameters["Status"] = false;
                }
                else if (string.IsNullOrWhiteSpace(MobileNumber))
                {
                    context.OutputParameters["Message"] = "Invalid Mobile Number";
                    context.OutputParameters["Status"] = false;
                }
                else { SendSMS(JobId, MobileNumber, context, service); }
                tracingService.Trace("Execution End");
            }
        }

        public void SendSMS(string JobId, string MobileNumber, IPluginExecutionContext context, IOrganizationService service)
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
                    tracingService.Trace("step-3");
                    context.OutputParameters["Status"] = false;
                }

                string state = job.Contains("hil_state") ? job.GetAttributeValue<EntityReference>("hil_state").Name.ToString() : string.Empty;
                string jobId = job.Contains("msdyn_workorder") ? job.GetAttributeValue<EntityReference>("msdyn_workorder").Id.ToString() : string.Empty;
                string zip = string.Empty;

                string address = job.Contains("hil_fulladdress") ? job.GetAttributeValue<String>("hil_fulladdress").ToString() : string.Empty;
                zip = job.Contains("hil_pincode") ? job.GetAttributeValue<EntityReference>("hil_pincode").Name.ToString() : string.Empty;
                address = ReplaceSpecialCharacters(address);
                remotePaymentLinkDetails.amount = job.Contains("hil_receiptamount") ? decimal.Round(job.GetAttributeValue<Money>("hil_receiptamount").Value, 2).ToString() : "0.00";

                Entity ent = service.Retrieve("hil_branch", job.GetAttributeValue<EntityReference>("hil_branch").Id, new ColumnSet("hil_mamorandumcode"));
                string _mamorandumCode = ent.Attributes.Contains("hil_mamorandumcode") ? ent.GetAttributeValue<string>("hil_mamorandumcode") : "";
                string _txnId = "D365_" + job.GetAttributeValue<string>("msdyn_name");
                remotePaymentLinkDetails.txnid = _txnId;
                remotePaymentLinkDetails.firstname = job.Contains("hil_customerref") ? job.GetAttributeValue<EntityReference>("hil_customerref").Name.ToString() : string.Empty;
                remotePaymentLinkDetails.email = job.Contains("hil_email") ? job.GetAttributeValue<String>("hil_email").ToString() : "abc@gmail.com";
                //remotePaymentLinkDetails.phone = job.Contains("hil_mobilenumber") ? job.GetAttributeValue<string>("hil_mobilenumber") : string.Empty;
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
                    context.OutputParameters["Status"] = true;
                }
                else
                {
                    context.OutputParameters["Message"] = obj.msg;
                    context.OutputParameters["Status"] = false;
                }
            }
            catch (Exception ex)
            {
                context.OutputParameters["Message"] = "D365 Internal Error " + ex.Message;
                context.OutputParameters["Status"] = false;
            }
        }
        private string ReplaceSpecialCharacters(string _inputStr)
        {
            return Regex.Replace(_inputStr, @"[^0-9a-zA-Z]+", "");
        }
    }
}
