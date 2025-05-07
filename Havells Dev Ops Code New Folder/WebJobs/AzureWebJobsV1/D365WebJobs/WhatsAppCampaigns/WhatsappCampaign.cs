using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace D365WebJobs.WhatsAppCampaigns
{
    public class WhatsappCampaign
    {
        private readonly IOrganizationService service;
        public WhatsappCampaign(IOrganizationService _service)
        {
            service = _service;
        }
        public void retriveJobs(ModeOfCommunication _ModeOfComm, string templateName, int dayDif, string ProductCategoryGUID)
        {
            try
            {
                string dateString = DateTime.Today.AddDays(dayDif).Year + "-" + DateTime.Today.AddDays(dayDif).Month + "-" + DateTime.Today.AddDays(dayDif).Day;
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
                    <attribute name='createdon' />
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
                    if (jobsColl.Entities.Count > 0)
                    {
                        cuatomerAssetColl.Entities.AddRange(jobsColl.Entities);
                        pageNumber++;
                    }
                    else
                    {
                        moreRecord = false;
                    }
                }

                string customerName = "Customer";
                string customerMobileNumber = string.Empty;
                string productsubcategory = string.Empty;
                string productmodel = string.Empty;
                string serialNumber = string.Empty;
                string installation_date = string.Empty;
                string _message = string.Empty;
                EntityCollection smsTempEntCol = null;
                Console.WriteLine("Total Jobs Found " + cuatomerAssetColl.Entities.Count);
                int done = 1;
                int error = 1;
                if (_ModeOfComm == ModeOfCommunication.SMS)
                {
                    QueryExpression _qryExp = new QueryExpression("hil_smstemplates");
                    _qryExp.ColumnSet = new ColumnSet(false);
                    _qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    _qryExp.Criteria.AddCondition("hil_templateid", ConditionOperator.Equal, templateName.Trim());
                    smsTempEntCol = service.RetrieveMultiple(_qryExp);
                }
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
                        EntityReference _customer = job.GetAttributeValue<EntityReference>("hil_customer");
                        EntityReference _productDiv = job.GetAttributeValue<EntityReference>("hil_productcategory");

                        if (_ModeOfComm == ModeOfCommunication.Whatsapp)
                        {
                            UserDetails userDetails = new UserDetails();
                            Traits objtraits = new Traits();
                            objtraits.name = customerName;

                            userDetails.phoneNumber = customerMobileNumber;
                            userDetails.countryCode = "+91";
                            userDetails.traits = objtraits;
                            userDetails.tags = new List<object>();

                            UpdateWhatsAppUser(userDetails);

                            CampaignDetails campaignDetails = new CampaignDetails();
                            campaignDetails.phoneNumber = customerMobileNumber;
                            campaignDetails.countryCode = "+91";
                            campaignDetails.@event = templateName;

                            Traits campaignTraits = new Traits();
                            campaignTraits.expire_on = " ";
                            if (templateName.Contains("_WP_"))// == "5D_WP_AMC"|| templateName == "21D_WP_AMC")
                            {
                                campaignTraits.prd_cat = "Havells Aqua";
                            }
                            if (templateName.Contains("_AC_"))// == "5D_AC_AMC" || templateName == "21D_AC_AMC")
                            {
                                campaignTraits.prd_cat = "Lloyd Air Conditioner";
                            }
                            campaignTraits.product_model = productsubcategory;
                            campaignTraits.product_serial_number = serialNumber;
                            campaignTraits.registration_date = installation_date;

                            campaignDetails.traits = campaignTraits;
                            SendMessage(campaignDetails);
                            Console.WriteLine("Executing Whatsapp Campaign# " + templateName + " -> " + done.ToString() + "/" + cuatomerAssetColl.Entities.Count + " to " + customerMobileNumber);
                        }
                        else {
                            if (templateName.Trim() == "1107168654225056811")
                            {
                                _message = string.Format("Hi {0},Protect your new Lloyd AC with Havells assured AMC plan. Special price starts at Rs.2499. TnC.  Buy Now https://bit.ly/3py9qIY -Havells", customerName);
                            }
                            else if (templateName.Trim() == "1107168654305357842")
                            {
                                _message = string.Format("Hi {0},Buy Havells AMC plan for your new Havells Water Purifier at flat 20%25 off. TnC. Visit https://bit.ly/3XzawAJ - Havells", customerName);
                            }
                            else {
                                continue;
                            }
                            sendSMS(smsTempEntCol.Entities[0].ToEntityReference(), customerMobileNumber, _message, _customer, dayDif, _productDiv);
                        }
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
        public void sendSMS(EntityReference  _templateId, string _mobileNumber,string _message,EntityReference _customer,int days,EntityReference _prodDivision) {
            Entity _sms = new Entity("hil_smsconfiguration");
            _sms["hil_contact"] = _customer;
            _sms["subject"] = "Whatsapp Campaign: " + _prodDivision.Name + " Product Registration " + days.ToString() + " days";
            _sms["hil_message"] = _message;
            _sms["hil_mobilenumber"] = _mobileNumber;
            _sms["hil_smstemplate"] = _templateId;
            _sms["hil_direction"] = new OptionSetValue(2);
            service.Create(_sms);
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
    public enum ModeOfCommunication
    {
        SMS = 0,
        Whatsapp =1
    }
}
