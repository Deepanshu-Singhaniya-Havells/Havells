using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using Newtonsoft.Json;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;

namespace Havells.Dataverse.CustomConnector.SalesOrder
{
    public class PushAMCDocumentsOnWA : CodeActivity
    {
        [Input("MediaGalleryRef")]
        [RequiredArgument]
        [ReferenceTarget("hil_mediagallery")]
        public InArgument<EntityReference> MediaGalleryRef { get; set; }
        public static ITracingService tracingService = null;
        public static readonly string[] MediaTypes = new string[] { "e3589095-4301-ef11-9f89-6045bdac6fcc", "df589095-4301-ef11-9f89-6045bdac6fcc", "e5589095-4301-ef11-9f89-6045bdac6fcc" }; //AMC Certificate, Payment Receipt, Sale Invoice
        protected override void Execute(CodeActivityContext executionContext)
        {
            var context = executionContext.GetExtension<IWorkflowContext>();
            var serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            tracingService = executionContext.GetExtension<ITracingService>();
            EntityReference mediaGalleryRef = MediaGalleryRef.Get(executionContext);
            try
            {
                Entity MediaGallery = service.Retrieve("hil_mediagallery", mediaGalleryRef.Id, new ColumnSet("hil_salesorder", "hil_url", "hil_mediatype", "hil_consumer"));
                EntityReference MediaTypeRef = MediaGallery.GetAttributeValue<EntityReference>("hil_mediatype");
                if (MediaTypes.Contains(MediaTypeRef.Id.ToString()))
                {
                    Entity customer = service.Retrieve("contact", MediaGallery.GetAttributeValue<EntityReference>("hil_consumer").Id, new ColumnSet("fullname", "mobilephone"));
                    string UserName = customer.Contains("fullname") ? FixName(customer.GetAttributeValue<string>("fullname")) : "";
                    string MobileNumber = customer.Contains("mobilephone") ? GetLast10Characters(customer.GetAttributeValue<string>("mobilephone")) : null;
                    string MediaUrl = MediaGallery.GetAttributeValue<string>("hil_url");
                    if (!string.IsNullOrWhiteSpace(MobileNumber))
                    {
                        AMCRoot ParamModelData = new AMCRoot();
                        ParamModelData.phoneNumber = MobileNumber;
                        ParamModelData.callbackData = "AMC document";
                        string triggerName = "amc_docs";
                        AMCTemplate amctemplate = GetTemplatebyTriggerName(triggerName, UserName, MediaUrl);
                        ParamModelData.template = amctemplate;

                        var RESULT = UpdateWhatsAppUser(ParamModelData, service);
                        if (RESULT.result == "true")
                        {
                            var ParamData = new TrackEvents();
                            ParamData.traits = new trait();
                            ParamData.phoneNumber = MobileNumber;
                            ParamData.traits.name = UserName;
                            ParamData.countryCode = "+91";
                            ParamData.@event = triggerName;

                            var resultMsg = SendMessage(ParamData, service);
                            if (resultMsg.result == true)
                            {
                                Entity MediaGallaryUpdate = new Entity("hil_mediagallery", mediaGalleryRef.Id);
                                MediaGallaryUpdate["hil_iswactasend"] = true;
                                service.Update(MediaGallaryUpdate);
                            }
                            else
                            {
                                throw new InvalidPluginExecutionException($"Error: From Whatsapp api side {mediaGalleryRef.Id}: {resultMsg.message}");
                            }
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException($"Error: While Whatsapp Invite sent on mobile number {mediaGalleryRef.Id}: {RESULT.result}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells.Dataverse.CustomConnector.SalesOrder.PushAMCDocumentsOnWA.CustomWorkflow Error " + ex.Message);
            }
        }
        private AMCTemplate GetTemplatebyTriggerName(string triggerName, string UserName, string MediaUrl)
        {
            AMCTemplate amctemplate = new AMCTemplate();
            switch (triggerName)
            {
                case "amc_docs":
                    {
                        amctemplate.name = triggerName;
                        amctemplate.headerValues = new List<string> { MediaUrl };
                        amctemplate.bodyValues = new List<string> { UserName };
                    }
                    break;
            }
            return amctemplate;
        }
        private UpdateAppUserRes UpdateWhatsAppUser(AMCRoot userDetails, IOrganizationService service)
        {
            IntegrationConfig intConfig = IntegrationConfiguration("WhatsApp_SendMedia", service);
            string url = intConfig.uri;
            string authInfo = intConfig.Auth;
            string data = JsonConvert.SerializeObject(userDetails);
            using (var client = new HttpClient())
            {
                HttpContent objcontent = new StringContent(data);
                client.BaseAddress = new Uri(url);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Add("Authorization", authInfo);
                var result = client.PostAsync(url, objcontent).Result;
                return JsonConvert.DeserializeObject<UpdateAppUserRes>(result.Content.ReadAsStringAsync().Result);
            }
        }
        private IntegrationConfig IntegrationConfiguration(string Param, IOrganizationService service)
        {
            IntegrationConfig output = new IntegrationConfig();
            QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
            qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
            qsCType.NoLock = true;
            qsCType.Criteria = new FilterExpression(LogicalOperator.And);
            qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
            EntityCollection integrationConfiguration = service.RetrieveMultiple(qsCType);
            if (integrationConfiguration.Entities.Count > 0)
            {
                output.uri = integrationConfiguration.Entities[0].GetAttributeValue<string>("hil_url");
                output.Auth = integrationConfiguration.Entities[0].GetAttributeValue<string>("hil_username") + " " + integrationConfiguration.Entities[0].GetAttributeValue<string>("hil_password");
            }
            return output;
        }
        private sendMesgRes SendMessage(TrackEvents campaignDetails, IOrganizationService service)
        {
            IntegrationConfig intConfig = IntegrationConfiguration("WhatsApp Campaign Details", service);
            string url = intConfig.uri;
            string authInfo = intConfig.Auth;
            string Campaigndata = JsonConvert.SerializeObject(campaignDetails);
            using (var client = new HttpClient())
            {
                HttpContent objCampaign = new StringContent(Campaigndata);
                client.BaseAddress = new Uri(url);
                client.DefaultRequestHeaders.Add("Authorization", authInfo);
                var result = client.PostAsync(url, objCampaign).Result;
                return JsonConvert.DeserializeObject<sendMesgRes>(result.Content.ReadAsStringAsync().Result);
            }
        }
        private string FixName(string input)
        {
            input = input.Trim();
            if (input.StartsWith("."))
            {
                input = input.Substring(1);
            }
            if (input.EndsWith("."))
            {
                input = input.Substring(0, input.Length - 1);
            }
            input = Regex.Replace(input, @"[.]{2,}", ".");
            input = Regex.Replace(input, @"[ ]{2,}", " ");
            input = Regex.Replace(input, @"[. ]{2}", "");
            input = Regex.Replace(input, @"\.\s", ".");
            input = Regex.Replace(input, @"\s\.", " ");
            input = Regex.Replace(input, @"[^A-Za-z. ]", "");
            return input;
        }
        private string GetLast10Characters(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return string.Empty;
            }
            return input.Length <= 10 ? input : input.Substring(input.Length - 10);
        }
    }
    public class AMCRoot
    {
        public string countryCode { get; set; } = "+91";
        public string phoneNumber { get; set; }
        public string callbackData { get; set; }
        public string type { get; set; } = "Template";
        public AMCTemplate template { get; set; }
    }
    public class AMCTemplate
    {
        public string name { get; set; }
        public List<string> headerValues { get; set; }
        public List<string> bodyValues { get; set; }
        public string languageCode { get; set; } = "en";
    }
    public class UpdateAppUserRes
    {
        public string result { get; set; }
    }
    public class TrackEvents
    {
        public trait traits { get; set; }
        public string phoneNumber { get; set; }
        public string countryCode { get; set; }
        public string @event { get; set; }
    }
    public class trait
    {
        public string name { get; set; }
    }
    public class sendMesgRes
    {
        public bool result { get; set; }
        public string message { get; set; }
    }
    public class IntegrationConfig
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }
}