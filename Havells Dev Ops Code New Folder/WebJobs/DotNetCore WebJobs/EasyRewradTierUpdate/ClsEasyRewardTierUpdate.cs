using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace EasyRewradTierUpdate
{
    internal class ClsEasyRewardTierUpdate
    {
        private readonly ServiceClient _service;
        public ClsEasyRewardTierUpdate(ServiceClient service)
        {
            _service = service;
        }
        public void EasyRewardTierUpdate()
        {
            try
            {
                if (_service.IsReady)
                {
                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='hil_easyrewardloyaltyprogram'>
                                        <attribute name='hil_easyrewardloyaltyprogramid'/>
                                        <attribute name='hil_name'/>
                                        <attribute name='createdon'/>
                                        <attribute name='hil_mobilenumber'/>
                                        <order attribute='hil_name' descending='false'/>
                                        <filter type='and'>
                                        <condition attribute='hil_productsyned' operator='eq' value='1'/>
                                       <condition attribute='hil_tiersynced' operator='eq' value='0'/>
                                        </filter>
                                        </entity>
                                        </fetch>";
                    Console.WriteLine("Getting Customer Information for Tier Update loyalty Programs");
                    EntityCollection entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    Console.WriteLine("Getting Customer Information Tier Update Count: {0} for loyalty Programs", entcoll.Entities.Count);
                    if (entcoll.Entities.Count > 0)
                    {
                        foreach (var c in entcoll.Entities)
                        {
                            string MobileNumber = c.Contains("hil_mobilenumber") ? c.Attributes["hil_mobilenumber"].ToString() : null;

                            QueryExpression qe = new QueryExpression("hil_integrationconfiguration");
                            qe.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                            qe.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "Easy Reward Get Tier Details");
                            Entity enColl = _service.RetrieveMultiple(qe)[0];
                            String URL = enColl.GetAttributeValue<string>("hil_url");
                            String Auth = enColl.GetAttributeValue<string>("hil_username") + ":" + enColl.GetAttributeValue<string>("hil_password");
                            Console.WriteLine("Tier API Details from D365 integration configuration");

                            /////////////////////// ER Tier Update API Call ///////////////////////
                            TierModel tierModel = new TierModel();
                            var data1 = new StringContent(JsonConvert.SerializeObject(tierModel), System.Text.Encoding.UTF8, "application/json");
                            data1.Headers.Add("LoginUserId", MobileNumber.ToString());
                            data1.Headers.Add("OperationToken", "");
                            HttpClient client1 = new HttpClient();
                            var byteArray = Encoding.ASCII.GetBytes(Auth);
                            client1.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                            Console.WriteLine("Get customer Tier data from ER API");
                            HttpResponseMessage responsetier = client1.PostAsync(URL, data1).Result;
                            if (responsetier.IsSuccessStatusCode)
                            {
                                Console.WriteLine("Tier API response: {0}", responsetier.StatusCode);
                                var resulttier = responsetier.Content.ReadAsStringAsync().Result;
                                dynamic TierResponse = JsonConvert.DeserializeObject<dynamic>(resulttier);

                                string? Tier = TierResponse != null ? TierResponse["TierName"] : null;
                                Console.WriteLine("Customer Tier: {0} and MobileNumber: {1}", Tier, MobileNumber);
                                if (!String.IsNullOrWhiteSpace(Tier))
                                {
                                    int TierValue = 0;
                                    switch (Tier.ToUpper())
                                    {
                                        case "PLUS":
                                        case "BASE":
                                            TierValue = 1;
                                            break;
                                        case "PREMIUM":
                                            TierValue = 2;
                                            break;
                                        case "PRIVILEDGE":
                                            TierValue = 3;
                                            break;
                                        default:
                                            TierValue = 0;
                                            break;
                                    }
                                    Console.WriteLine("Getting Ready Update Customer Tier in Contact table");
                                    QueryExpression query;
                                    query = new QueryExpression("contact");
                                    query.ColumnSet = new ColumnSet(false);
                                    query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, MobileNumber);
                                    EntityCollection sourceEntColl = _service.RetrieveMultiple(query);
                                    if (sourceEntColl.Entities.Count > 0)
                                    {
                                        Guid CustomerGuid = sourceEntColl.Entities[0].Id;
                                        Entity UpdateTier = new Entity("contact", CustomerGuid);
                                        UpdateTier["hil_loyaltyprogramtier"] = new OptionSetValue(TierValue);
                                        _service.Update(UpdateTier);
                                        Console.WriteLine("Customer Tier Updated Guid:{0}", CustomerGuid);

                                        Entity TierStatus = new Entity("hil_easyrewardloyaltyprogram", c.Id);
                                        TierStatus["hil_tiersynced"] = true;
                                        _service.Update(TierStatus);
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Customer Tier Not Found in ER API");
                                }
                            }
                            Console.WriteLine("Completed.");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Error: D365 Service is unavailable!");
                }
            }
            catch (Exception ex)
            {
                Console.Write("Error: " + ex.Message);
            }
        }
    }
    public class TierModel
    {
        public string CountryCode { get; set; } = "91";
    }
}
