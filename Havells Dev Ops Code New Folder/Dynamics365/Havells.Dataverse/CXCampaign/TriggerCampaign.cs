using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using RestSharp;
using static CXCampaign.CXCampaignModels;
using System.Xml.Serialization;

namespace CXCampaign
{
    internal class TriggerCampaign : DataverseServiceFactory
    {
        private string eventURL { get; set; }
        private IntegrationConfig intConfig { get; set; }

        internal protected IntegrationConfig IntegrationConfiguration(string Param)
        {
            IntegrationConfig output = new IntegrationConfig();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + " " + integrationConfiguration.GetAttributeValue<string>("hil_password");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:- " + ex.Message);
            }
            return output;
        }

        internal void ExecuteCamapigns(ModeOfCommunication _modeOfCommunication)
        {
            try
            {
                Console.WriteLine(DateTime.Now + " Query execution starts for Campaign Triggers: " + _modeOfCommunication);

                int count = 0;
                QueryExpression query = new QueryExpression("hil_campaigndata");
                query.ColumnSet = new ColumnSet(true);
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                query.Criteria.AddCondition("createdon", ConditionOperator.Today);
                query.Criteria.AddCondition("hil_communicationmode", ConditionOperator.Equal, ((int)_modeOfCommunication));

                if (_modeOfCommunication == ModeOfCommunication.SMS)
                    query.Criteria.AddCondition("hil_response", ConditionOperator.Null);

                query.PageInfo = new PagingInfo();
                query.PageInfo.Count = 100;
                query.PageInfo.PageNumber = 1;
                query.PageInfo.ReturnTotalRecordCount = true;
                EntityCollection entCol = RetrieveMultiple(query);
                int i = 1;
                count = entCol.Entities.Count;

                if (entCol.Entities.Count == 0)
                {
                    Console.WriteLine(" No Record Found for Campaign Triggers: " + _modeOfCommunication);
                }
                do
                {
                    foreach (Entity entity in entCol.Entities)
                    {
                        try
                        {
                            string customerName = entity.GetAttributeValue<string>("hil_customername");
                            string mobileNumber = entity.GetAttributeValue<string>("hil_mobilenumber");
                            string templateName = entity.GetAttributeValue<string>("hil_templateid");
                            string productCategory = entity.Contains("hil_productcat") ? entity.GetAttributeValue<EntityReference>("hil_productcat").Name : "";
                            string serialNumber = entity.Contains("hil_serialnumber") ? entity.GetAttributeValue<EntityReference>("hil_serialnumber").Name : null;
                            string model = entity.Contains("hil_model") ? entity.GetAttributeValue<EntityReference>("hil_model").Name : "";
                            string installation_date = entity.GetAttributeValue<DateTime>("hil_registrationdate").AddMinutes(330).Date.ToString("yyyy-MM-dd");
                            string sms = entity.Contains("hil_message") ? entity.GetAttributeValue<string>("hil_message") : "";

                            ModeOfCommunication modeOfCommunication = (ModeOfCommunication)entity.GetAttributeValue<OptionSetValue>("hil_communicationmode").Value;

                            if (modeOfCommunication == ModeOfCommunication.Whatsapp)
                            {
                                RegisterUserOnWhatsApp(customerName, mobileNumber);
                                SendWhatsEvent(mobileNumber, templateName, productCategory, serialNumber, model, installation_date);
                            }
                            else
                            {
                                sendCampaignSMS(templateName, mobileNumber, sms);
                            }
                            SetStateRequest req = new SetStateRequest();
                            req.State = new OptionSetValue(1);
                            req.Status = new OptionSetValue(2);
                            req.EntityMoniker = entity.ToEntityReference();
                            var res = (SetStateResponse)Execute(req);
                            i += 1;
                            Console.WriteLine("Processing... " + i.ToString() + "/" + count.ToString());
                        }
                        catch (Exception ex)
                        {
                            try
                            {
                                Entity entity1 = new Entity(entity.LogicalName, entity.Id);
                                entity1["hil_response"] = ex.Message;
                                Update(entity1);
                            }
                            catch (Exception exe)
                            {
                                Console.WriteLine(exe.Message);
                            }
                        }
                    }
                    query.PageInfo.PageNumber += 1;
                    query.PageInfo.PagingCookie = entCol.PagingCookie;
                    entCol = RetrieveMultiple(query);
                    count = count + entCol.Entities.Count;
                }
                while (entCol.MoreRecords);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void RegisterUserOnWhatsApp(string customerName, string customerMobileNumber)
        {
            intConfig = IntegrationConfiguration("WhatsApp_Campaign_API");

            CXCampaignModels.UserDetails userDetails = new CXCampaignModels.UserDetails();
            Traits objtraits = new Traits();

            RegexOptions options = RegexOptions.None;
            Regex regex = new Regex("[ ]{2,}", options);
            customerName = regex.Replace(customerName, " ");

            objtraits.name = customerName;

            userDetails.phoneNumber = customerMobileNumber;
            userDetails.countryCode = "+91";
            userDetails.traits = objtraits;
            userDetails.tags = new List<object>();


            string data = JsonConvert.SerializeObject(userDetails);
            var client = new RestClient(string.Format(intConfig.uri, "users"));
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);

            request.AddHeader("Authorization", intConfig.Auth);
            request.AddHeader("Content-Type", "application/json");

            request.AddParameter("application/json", data, ParameterType.RequestBody);
            var response = client.Execute(request);
            var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<UserResponse>(response.Content);
            if (((int)response.StatusCode) != 200 && ((int)response.StatusCode) != 202)
            {
                throw new Exception(obj.message);
            }
            else
                Console.WriteLine(obj.message);
        }
        private void SendWhatsEvent(string customerMobileNumber, string templateName, string productCategory, string serialNumber, string model, string installation_date)
        {
            try
            {
                CampaignDetails campaignDetails = new CampaignDetails();
                campaignDetails.phoneNumber = customerMobileNumber;
                campaignDetails.countryCode = "+91";
                campaignDetails.@event = templateName;

                Traits campaignTraits = new Traits();
                campaignTraits.expire_on = " ";
                campaignTraits.prd_cat = productCategory;
                campaignTraits.product_model = model;
                campaignTraits.product_serial_number = serialNumber;
                campaignTraits.registration_date = installation_date;

                campaignDetails.traits = campaignTraits;

                string data = JsonConvert.SerializeObject(campaignDetails);
                var client = new RestClient(string.Format(intConfig.uri, "events"));
                client.Timeout = -1;
                var request = new RestRequest(Method.POST);

                request.AddHeader("Authorization", intConfig.Auth);
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter("application/json", data, ParameterType.RequestBody);
                var response = client.Execute(request);

                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<UserResponse>(response.Content);
                if (obj != null)
                {
                    if (((int)response.StatusCode) != 200 && ((int)response.StatusCode) != 201 && ((int)response.StatusCode) != 202)
                    {
                        throw new Exception(obj.message);
                    }
                    else
                        Console.WriteLine(obj.message);
                }
            }
            catch (Exception ex)
            {

                throw new Exception(ex.Message);
            }
        }
        private void sendCampaignSMS(string _templateId, string _mobileNumber, string _message)
        {
            string _api = "https://japi.instaalerts.zone/failsafe/HttpLink?aid=640990&pin=w~7Xg)9V&mnumber=" + _mobileNumber + "&signature=HAVELL&message=" + _message + "&dlt_entity_id=110100001483&dlt_template_id=" + _templateId + "&cust_ref=";

            WebRequest request = WebRequest.Create(_api);
            request.Method = "POST";
            WebResponse response = null;
            string IfOkay = string.Empty;
            string responseFromServer = string.Empty;
            try
            {
                response = request.GetResponse();
                Stream dataStream = Stream.Null;
                IfOkay = ((HttpWebResponse)response).StatusDescription;
                dataStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream);
                responseFromServer = reader.ReadToEnd();
            }
            catch (WebException ex)
            {
                throw new Exception(ex.Message);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}
