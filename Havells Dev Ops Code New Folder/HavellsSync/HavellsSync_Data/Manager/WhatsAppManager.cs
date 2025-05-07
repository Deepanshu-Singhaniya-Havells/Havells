using HavellsSync_Data.IManager;
using HavellsSync_ModelData.AMC;
using HavellsSync_ModelData.Common;
using Microsoft.Extensions.Configuration;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System.Net;

namespace HavellsSync_Data.Manager
{
    public class WhatsAppManager : IWhatsAppManager
    {
        private readonly ICrmService _service;
        public WhatsAppManager(ICrmService service, IConfiguration configuration)
        {
            Check.Argument.IsNotNull(nameof(service), service);
            this._service = service;
        }
        public async Task<(WhatsappConnectRes, RequestStatus)> WhatsappConnect(WhatsAppModel whatsAppdata)
        {
            WhatsappConnectRes objResult = new WhatsappConnectRes();
            try
            {
                string ConsumerName = "";
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='contact'>
                                    <attribute name='fullname' />
                                    <attribute name='telephone1' />
                                    <attribute name='contactid' />
                                    <order attribute='fullname' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='mobilephone' operator='eq' value='{whatsAppdata.MobileNumber}' />
                                    </filter>
                                  </entity>
                                </fetch>";
                EntityCollection entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entcoll.Entities.Count > 0)
                {
                    ConsumerName = entcoll.Entities[0].Contains("fullname") ? entcoll.Entities[0].GetAttributeValue<string>("fullname") : "";
                }
                UserDetail userDetails = new UserDetail();
                Traits objtraits = new Traits();
                objtraits.name = ConsumerName;
                userDetails.phoneNumber = whatsAppdata.MobileNumber;
                userDetails.countryCode = "+91";
                userDetails.traits = objtraits;
                userDetails.tags = new List<object>();

                objResult.UpdateAppUserRes = UpdateWhatsAppUser(userDetails);

                CampaignDetails campaignDetails = new CampaignDetails();
                campaignDetails.phoneNumber = whatsAppdata.MobileNumber;
                campaignDetails.countryCode = "+91";
                switch (whatsAppdata.Event)
                {
                    case "1":
                        campaignDetails.@event = "user_calldrop_notify";
                        break;
                    case "2":
                        campaignDetails.@event = "user_callAbendent_notify";
                        break;
                    case "3":
                        campaignDetails.@event = "user_ivroption_notify";
                        break;
                    case "4":
                        campaignDetails.@event = "cc_queue_event";
                        break;
                    case "5":
                        campaignDetails.@event = "cc_rep_event";
                        break;
                    case "6":
                        campaignDetails.@event = "CC_holiday_aug";
                        break;
                    case "7":
                        campaignDetails.@event = "CC_hol_gandhijay";
                        break;
<<<<<<< HEAD
=======
                    case "8":
                        campaignDetails.@event = "CC_hol_26rep";
                        break;
>>>>>>> 2bcdff781305e7e5eeb0eea84306b44cb57324c8
                    default:
                        campaignDetails.@event = "NA";
                        break;
                }
                //campaignDetails.@event = whatsAppdata.Event == "1" ? "user_calldrop_notify" : whatsAppdata.Event == "2" ? "user_callAbendent_notify" : whatsAppdata.Event == "3" ?"user_ivroption_notify": whatsAppdata.Event == "4" ? "cc_queue_event": "user_ivroption_notify";

                campaignDetails.traits = new Traits();
                objResult.SendMesgRes = SendMessage(campaignDetails);
                return (objResult, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.OK,
                    Message = CommonMessage.SuccessMsg
                });
            }
            catch (Exception ex)
            {
                return (objResult, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }
        public sendMesgRes SendMessage(CampaignDetails campaignDetails)
        {
            IntegrationConfig intConfig = IntegrationConfiguration(_service, "WhatsApp Campaign Details");
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
        public UpdateAppUserRes UpdateWhatsAppUser(UserDetail userDetails)
        {
            IntegrationConfig intConfig = IntegrationConfiguration(_service, "Update WhatsApp User");
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
        public static IntegrationConfig IntegrationConfiguration(ICrmService service, string Param)
        {
            IntegrationConfig output = new IntegrationConfig();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + " " + integrationConfiguration.GetAttributeValue<string>("hil_password");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:- " + ex.Message);
            }
            return output;
        }

    }
}
