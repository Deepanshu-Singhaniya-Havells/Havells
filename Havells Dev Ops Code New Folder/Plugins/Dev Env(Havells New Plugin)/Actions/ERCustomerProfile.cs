using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Web.Configuration;

namespace HavellsPlugIn.Actions
{
    public class ERCustomerProfile : IPlugin
    {
        private static ITracingService tracingService = null;
        private static IOrganizationService service;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                tracingService.Trace("Step 1");
                //if (context.InputParameters.Contains("CustomerGuid") && context.InputParameters["CustomerGuid"] is string
                //        && context.InputParameters.Contains("MobileNumber") && context.InputParameters["MobileNumber"] is string && context.Depth == 1)
                if (context.InputParameters.Contains("CustomerID") && context.InputParameters["CustomerID"] is string && context.Depth == 1)
                {
                    string MobileNumber = "";
                    string CustomerGuid = (string)context.InputParameters["CustomerID"];
                    tracingService.Trace(CustomerGuid);

                    if (!string.IsNullOrEmpty(CustomerGuid))
                    {
                        tracingService.Trace("if condition");
                        MobileNumber = GetMobileNumber(new Guid(CustomerGuid), service);
                        tracingService.Trace(MobileNumber);
                        string URL = GetProfileUrl(MobileNumber, service);
                        if (!string.IsNullOrEmpty(URL))
                        {
                            tracingService.Trace("URL");
                            context.OutputParameters["ReturnMessage"] = "Success";
                            context.OutputParameters["WidgetURL"] = URL;
                        }
                        else
                        {
                            tracingService.Trace("URL ELSE");
                            context.OutputParameters["ReturnMessage"] = "Failed";
                            context.OutputParameters["WidgetURL"] = "";
                        }
                    }
                    else
                    {
                        tracingService.Trace("else condition");
                        context.OutputParameters["ReturnMessage"] = "Failed";
                        context.OutputParameters["WidgetURL"] = "";
                    }
                }
            }
            catch (Exception ex)
            {
                context.OutputParameters["ReturnMessage"] = "D365 Internal Server Error : " + ex.Message;
                context.OutputParameters["WidgetURL"] = "";
            }
        }
        public string GetMobileNumber(Guid customerGuid, IOrganizationService service)
        {
            tracingService.Trace("GetMobileNumber");
            string MobileNumber = "";
            Entity data = service.Retrieve("contact", customerGuid, new ColumnSet("mobilephone"));
            MobileNumber = data.Contains("mobilephone") ? data.GetAttributeValue<string>("mobilephone") : "8287910060";
            return MobileNumber;

        }
        public string GetProfileUrl(string mobileNumber, IOrganizationService service)
        {
            tracingService.Trace("GetProfileFunction");
            IntegrationConfiguration inconfig = GetIntegrationConfiguration("Work flow Widget", service);//->// To pass the name
            WidgetRequest objWidgetRequest = new WidgetRequest { WidgetCode = "VC" };
            var Paramdata = new StringContent(JsonConvert.SerializeObject(objWidgetRequest), System.Text.Encoding.UTF8, "application/json");
            var byteArray = Encoding.ASCII.GetBytes(inconfig.userName + ":" + inconfig.password);
            tracingService.Trace("Work flow Widget");
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray)).ToString());
            client.DefaultRequestHeaders.Add("LoginUserId", mobileNumber);
            client.DefaultRequestHeaders.Add("OperationToken", "");
            HttpResponseMessage Result = client.PostAsync(inconfig.url, Paramdata).Result;
            if (Result.IsSuccessStatusCode)
            {
                tracingService.Trace("Response message");
                WidgetResponse obj = JsonConvert.DeserializeObject<WidgetResponse>(Result.Content.ReadAsStringAsync().Result);
                if (!string.IsNullOrEmpty(obj.WidgetURL))
                {
                    return obj.WidgetURL;
                }
            }
            return "";
        }
        private IntegrationConfiguration GetIntegrationConfiguration(string name, IOrganizationService service)
        {
            try
            {
                IntegrationConfiguration inconfig = new IntegrationConfiguration();
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, name);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
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
    public class WidgetRequest
    {
        public string WidgetCode { get; set; }
    }
    public class WidgetResponse
    {
        public string WidgetURL { get; set; }
        public string ReturnMessage { get; set; }
        public string ReturnCode { get; set; }
    }
    public class IntegrationConfiguration
    {
        public string url { get; set; }
        public string userName { get; set; }
        public string password { get; set; }
    }
}