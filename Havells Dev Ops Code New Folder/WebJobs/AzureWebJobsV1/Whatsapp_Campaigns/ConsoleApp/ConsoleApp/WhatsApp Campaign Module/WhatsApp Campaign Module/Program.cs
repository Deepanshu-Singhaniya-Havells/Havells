using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using RestSharp;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System.Configuration;
using Havells_Plugin;

namespace WhatsApp_Campaign_Module
{
    internal class Program
    {

        static IOrganizationService _service;
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
            var CrmURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();
            string finalString = string.Format(connStr, CrmURL);
            _service = HavellsConnection.CreateConnection.createConnection(finalString);


            if (_service != null)
            {
                Console.WriteLine("Connection established");
                Console.ReadLine();


                //Outer Loop to get the Active Campaigns
                string _effectiveOn = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString().PadLeft(2, '0') + "-" + DateTime.Now.Day.ToString().PadLeft(2, '0');
                string fetch_campaigns = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='campaign'>
                    <attribute name='name' />
                    <attribute name='istemplate' />
                    <attribute name='statuscode' />
                    <attribute name='createdon' />
                    <attribute name='campaignid' />
                    <attribute name='createdby' />
                    <attribute name='actualstart' />
                    <attribute name='actualend' />
                    <order attribute='name' descending='true' />
                    <filter type='and'>
                    <condition attribute='hil_campaignsource' operator='eq' value='7' />
                    <condition attribute='statuscode' operator='eq' value='2' />
                    <condition attribute='actualstart' operator='on-or-before' value='{_effectiveOn}' />
                    <condition attribute='actualend' operator='on-or-after' value='{_effectiveOn}' />
                    </filter>
                    </entity>
                    </fetch>";

                EntityCollection campaignsr = _service.RetrieveMultiple(new FetchExpression(fetch_campaigns));
                foreach (Entity camp in campaignsr.Entities)
                {
                    string camp_name = camp.GetAttributeValue<string>("name").ToString();
                    string camp_code = camp.GetAttributeValue<string>("campcode").ToString();

                    string _fetchList = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                      <entity name='list'>
                        <attribute name='query' />
                        <attribute name='type' />
                        <order attribute='listname' descending='true' />
                        <link-entity name='campaignitem' from='entityid' to='listid' visible='false' intersect='true'>
                          <link-entity name='campaign' from='campaignid' to='campaignid' alias='ad'>
                            <filter type='and'>
                              <condition attribute='campaignid' operator='eq' value='{camp.Id}' />
                            </filter>
                          </link-entity>
                        </link-entity>
                      </entity>
                    </fetch>";
                    
                    EntityCollection entityCollection = _service.RetrieveMultiple(new FetchExpression(_fetchList));
                    if (entityCollection.Entities.Count > 0)
                    {
                        EntityCollection _entcoll1 = _service.RetrieveMultiple(new FetchExpression(entityCollection.Entities[0].GetAttributeValue<string>("query")));
                        
                        string 
                        foreach(Entity it in  entityCollection.Entities)
                        {

                        }
                        Console.WriteLine(_entcoll1.Entities.Count.ToString());
                        
                    }
                }
            }
        }



        public void UpdateWhatsAppUser(UserDetails userDetails)
        {
            IntegrationConfig intConfig = IntegrationConfiguration(_service, "Update WhatsApp User");
            string url = intConfig.uri;
            string authInfo = intConfig.Auth;
            string data = JsonConvert.SerializeObject(userDetails);
            var client = new RestClient();
            var request = new RestRequest(url, Method.Post);

            request.AddHeader("Authorization", authInfo);
            request.AddHeader("Content-Type", "application/json");

            request.AddParameter("application/json", data, ParameterType.RequestBody);
            client.Execute(request);
        }
        public void SendMessage(CampaignDetails campaignDetails)
        {
            IntegrationConfig intConfig = IntegrationConfiguration(_service, "WhatsApp Campaign Details");
            string url = intConfig.uri;
            string authInfo = intConfig.Auth;
            string data = JsonConvert.SerializeObject(campaignDetails);
            var client = new RestClient();

            var request = new RestRequest(url, Method.Post);

            request.AddHeader("Authorization", authInfo);
            request.AddHeader("Content-Type", "application/json");





            request.AddParameter("application/json", data, ParameterType.RequestBody);
            client.Execute(request);
        }
        public static IntegrationConfig IntegrationConfiguration(IOrganizationService service, string Param)
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

    
    
    
    public class IntegrationConfig
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }
    public class UserDetails
    {
        public string phoneNumber { get; set; }
        public string countryCode { get; set; }
        public Traits traits { get; set; }
        public List<object> tags { get; set; }
    }
    public class CampaignDetails
    {
        public string phoneNumber { get; set; }
        public string countryCode { get; set; }
        public string @event { get; set; }
        public Traits traits { get; set; }
    }
    public class Traits
    {
        public string name { get; set; }
        public string expire_on { get; set; }
        public string prd_cat { get; set; }
        public string product_model { get; set; }
        public string product_serial_number { get; set; }
        public string registration_date { get; set; }
    }
}
