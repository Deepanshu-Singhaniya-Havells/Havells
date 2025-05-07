using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;

namespace Havells.Dataverse.CustomConnector.Home_Advisory
{
    public class GetUserTimeSlots : IPlugin
    {
        public static ITracingService tracingService = null;
        IPluginExecutionContext context;
        public void Execute(IServiceProvider serviceProvider)
        {
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            List<GetUserTimeSlotsRoot> GetUserTimeSlotsData = new List<GetUserTimeSlotsRoot>();
            string JsonResponse = string.Empty;
            try
            {
                string EnquiryTypeCode = Convert.ToString(context.InputParameters["EnquiryTypeCode"]);
                string SlotDate = Convert.ToString(context.InputParameters["SlotDate"]);
                string UserCode = Convert.ToString(context.InputParameters["UserCode"]);

                GetUserTimeSlotsRequest requestParm = new GetUserTimeSlotsRequest
                {
                    EnquiryTypeCode = EnquiryTypeCode,
                    SlotDate = SlotDate,
                    UserCode = UserCode
                };

                QueryExpression qe = new QueryExpression("hil_integrationconfiguration");
                qe.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qe.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "GetUserTimeSlots");
                Entity enColl = service.RetrieveMultiple(qe)[0];
                String URL = enColl.GetAttributeValue<string>("hil_url");
                String Auth = enColl.GetAttributeValue<string>("hil_username") + ":" + enColl.GetAttributeValue<string>("hil_password");

                var data = new StringContent(JsonConvert.SerializeObject(requestParm), Encoding.UTF8, "application/json");
                HttpClient client = new HttpClient();
                var byteArray = Encoding.ASCII.GetBytes(Auth);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));

                HttpResponseMessage response = client.PostAsync(URL, data).Result;
                if (response.IsSuccessStatusCode)
                {
                    var result = response.Content.ReadAsStringAsync().Result;
                    GetUserTimeSlotsRoot apiResponse = JsonConvert.DeserializeObject<GetUserTimeSlotsRoot>(result);

                    GetUserTimeSlotsData.Add(apiResponse);
                    JsonResponse = JsonConvert.SerializeObject(GetUserTimeSlotsData);
                    context.OutputParameters["data"] = JsonResponse;
                }
            }
            catch (Exception ex)
            {
                GetUserTimeSlotsData.Add(new GetUserTimeSlotsRoot() { IsSuccess = false, Message = "D365 Internal Server Error : " + ex.Message });
                JsonResponse = JsonConvert.SerializeObject(GetUserTimeSlotsData);
                context.OutputParameters["data"] = JsonResponse;
                return;
            }
        }
        public class GetUserTimeSlotsRequest
        {
            public string UserCode { get; set; }
            public string SlotDate { get; set; }
            public string EnquiryTypeCode { get; set; }
        }
        public class GetUserTimeSlotsDatum
        {
            public string SlotStart { get; set; }
            public string SlotEnd { get; set; }
            public int IsAvailable { get; set; }
        }
        public class GetUserTimeSlotsRoot
        {
            public List<GetUserTimeSlotsDatum> Data { get; set; }
            public bool IsSuccess { get; set; }
            public int ResponseCode { get; set; }
            public string Message { get; set; }
        }
    }
}
