using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Navigation;
using QueryExpression = Microsoft.Xrm.Sdk.Query.QueryExpression;

namespace ConsoleApp
{
    public class WhatsAppPromotions
    {
        private readonly IOrganizationService service;
        public WhatsAppPromotions(IOrganizationService _service)
        {
            service = _service;
        }
        public void retriveJobs(string templateName, int dayDif, string ProductCategoryGUID)
        {
            // string templateName = e;
            // string ProductCategoryGUID = "D51EDD9D-16FA-E811-A94C-000D3AF0694E";
            // int dayDif = ex;
            try
            {
                string dateString = DateTime.Today.AddDays(dayDif).Year + "-" + DateTime.Today.AddDays(dayDif).Month + "-" + DateTime.Today.AddDays(dayDif).Day;

                //query for products registration
                int pageNumber = 1;
                EntityCollection cuatomerAssetColl = new EntityCollection();
                bool moreRecord = true;
                while (moreRecord)
                {
                    string fetchCUstomerAsset = $@"<fetch version=""1.0"" page=""{pageNumber}"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                                                  <entity name=""msdyn_customerasset"">
                                                    <attribute name=""msdyn_customerassetid"" />
                                                    <attribute name=""msdyn_name"" />
                                                    <attribute name=""hil_productcategory"" />
                                                    <attribute name=""hil_productsubcategorymapping"" />
                                                    <attribute name=""msdyn_product"" />
                                                    <attribute name=""hil_customer"" />
                                                    <order attribute=""msdyn_name"" descending=""false"" />
                                                    <filter type=""and"">
                                                      <condition attribute=""createdon"" operator=""on"" value=""{dateString}"" />
                                                      <condition attribute=""hil_productcategory"" operator=""eq"" value=""{ProductCategoryGUID}"" />
                                                    </filter>
                                                    <link-entity name=""contact"" from=""contactid"" to=""hil_customer"" visible=""false"" link-type=""outer"" alias=""consumer"">
                                                      <attribute name=""mobilephone"" />
                                                    </link-entity>
                                                  </entity>
                                                </fetch>";
                    EntityCollection jobsColl = service.RetrieveMultiple(new FetchExpression(fetchCUstomerAsset));
                    if(jobsColl.Entities.Count > 0)
                    {
                        cuatomerAssetColl.Entities.AddRange(jobsColl.Entities);
                        pageNumber++;
                    }
                    else
                    {
                        moreRecord = false;
                    }
                }


                //query for products installation
                /*
                //string fetch = $@"<fetch version=""1.0"" page=""1"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                //                  <entity name=""msdyn_workorder"">
                //                    <attribute name=""msdyn_name"" />
                //                    <attribute name=""createdon"" />
                //                    <attribute name=""msdyn_timeclosed"" />
                //                    <attribute name=""hil_productsubcategory"" />
                //                    <attribute name=""hil_mobilenumber"" />
                //                    <attribute name=""hil_customerref"" />
                //                    <attribute name=""hil_callsubtype"" />
                //                    <attribute name=""msdyn_workorderid"" />
                //                    <order attribute=""msdyn_name"" descending=""false"" />
                //                    <filter type=""and"">
                //                      <condition attribute=""hil_productcategory"" operator=""in"">
                //                        <value uitype=""product"">{ProductCategoryGUID}</value>
                //                      </condition>
                //                      <condition attribute=""msdyn_timeclosed"" operator=""on"" value=""{dateString}"" />
                //                      <condition attribute=""hil_callsubtype"" operator=""eq"" value=""{{E3129D79-3C0B-E911-A94E-000D3AF06CD4}}"" />
                //                    </filter>
                //                    <link-entity name=""msdyn_customerasset"" from=""msdyn_customerassetid"" to=""msdyn_customerasset"" visible=""false"" link-type=""inner"" alias=""ca"">
                //                        <attribute name=""msdyn_name"" />
                //                        <attribute name=""msdyn_product"" />
                //                    </link-entity>
                //                  </entity>
                //                </fetch>";*/


                string customerName = "Customer";
                string customerMobileNumber = string.Empty;
                string productsubcategory = string.Empty;
                string productmodel = string.Empty;
                string serialNumber = string.Empty;
                string installation_date = string.Empty;

                Console.WriteLine("Total Jobs Found " + cuatomerAssetColl.Entities.Count);
                int done = 1;
                int error = 1;
                foreach (Entity job in cuatomerAssetColl.Entities)
                {
                    try
                    {
                        customerName = job.Contains("hil_customer") ? job.GetAttributeValue<EntityReference>("hil_customer").Name : "";
                        customerMobileNumber = job.GetAttributeValue<AliasedValue>("consumer.mobilephone").Value.ToString();
                        productsubcategory = job.GetAttributeValue<EntityReference>("hil_productsubcategorymapping").Name;
                        productmodel = job.GetAttributeValue<EntityReference>("msdyn_product").Name;
                        serialNumber = job.GetAttributeValue<String>("msdyn_name").ToString();
                        installation_date = job.GetAttributeValue<DateTime>("createdon").ToString("yyyy-MM-dd");

                        Console.WriteLine(customerMobileNumber);

                        UserDetails userDetails = new UserDetails();
                        Traits objtraits = new Traits();
                        objtraits.name = customerName;

                        userDetails.phoneNumber = "9050673956";// customerMobileNumber;
                        userDetails.countryCode = "+91";
                        userDetails.traits = objtraits;
                        userDetails.tags = new List<object>();
                        UpdateWhatsAppUser(userDetails);

                        CampaignDetails campaignDetails = new CampaignDetails();
                        campaignDetails.phoneNumber = "9050673956";//  customerMobileNumber;
                        campaignDetails.countryCode = "+91";
                        campaignDetails.@event = templateName;


                        Traits campaignTraits = new Traits();
                        campaignTraits.expire_on = " ";
                        if(templateName.Contains("_WP_"))// == "5D_WP_AMC"|| templateName == "21D_WP_AMC")
                        {
                            campaignTraits.prd_cat = "Havells Aqua";
                        }
                        if(templateName.Contains("_AC_"))// == "5D_AC_AMC" || templateName == "21D_AC_AMC")
                        {
                            campaignTraits.prd_cat = "Lloyd Air Conditioner";
                        }
                        campaignTraits.product_model = productsubcategory;
                        campaignTraits.product_serial_number = serialNumber;
                        campaignTraits.registration_date = installation_date;


                        campaignDetails.traits = campaignTraits;
                        SendMessage(campaignDetails);
                        Console.WriteLine(done + "/" + cuatomerAssetColl.Entities.Count + "Msg Send to " + customerMobileNumber);
                        done++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("***** " + error + "/" + cuatomerAssetColl.Entities.Count + "Error Occured " + ex.Message);
                        error++;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("!!!!!!!!!!!!!!! Error " + ex.Message);
            }
        }
        public void UpdateWhatsAppUser(UserDetails userDetails)
        {
            IntegrationConfig intConfig = IntegrationConfiguration(service, "Update WhatsApp User");
            string url = intConfig.uri;
            string authInfo = intConfig.Auth;
            string data = JsonConvert.SerializeObject(userDetails);
            var client = new RestClient(url);
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);

            request.AddHeader("Authorization", authInfo);
            request.AddHeader("Content-Type", "application/json");

            request.AddParameter("application/json", data, ParameterType.RequestBody);
            client.Execute(request);
        }
        public void SendMessage(CampaignDetails campaignDetails)
        {
            IntegrationConfig intConfig = IntegrationConfiguration(service, "WhatsApp Campaign Details");
            string url = intConfig.uri;
            string authInfo = intConfig.Auth;
            string data = JsonConvert.SerializeObject(campaignDetails);
            var client = new RestClient(url);
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);

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
