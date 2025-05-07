using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Havells.CRM.WhatsAppPromotions
{
    public class SendCampaign : Program
    {
        static string eventURL = null;
        static IntegrationConfig intConfig = null;
        static SendCampaign()
        {
            intConfig = IntegrationConfiguration(service, "WhatsApp_Campaign_API");
        }
        public static void RetriveCamapign()
        {
            try
            {
                int count = 0;
                QueryExpression query = new QueryExpression("hil_campaigndata");
                query.ColumnSet = new ColumnSet(true);
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, "2023-12-29");
                //query.Criteria.AddCondition("hil_response", ConditionOperator.Null);
                query.PageInfo = new PagingInfo();
                query.PageInfo.Count = 100;
                query.PageInfo.PageNumber = 1;
                query.PageInfo.ReturnTotalRecordCount = true;
                EntityCollection entCol = service.RetrieveMultiple(query);
                int i = 1;
                do
                {
                    entCol = service.RetrieveMultiple(query);
                    count += entCol.Entities.Count;
                    foreach (Entity entity in entCol.Entities)
                    {
                        try
                        {
                            string customerName = entity.GetAttributeValue<string>("hil_customername");
                            string mobileNumber = entity.GetAttributeValue<string>("hil_mobilenumber");
                            string templateName = entity.GetAttributeValue<string>("hil_templateid");
                            string productCategory = entity.Contains("hil_productcat") ? entity.GetAttributeValue<EntityReference>("hil_productcat").Name : "";
                            string serialNumber = entity.Contains("hil_serialnumber") ? entity.GetAttributeValue<EntityReference>("hil_serialnumber").Name : entity.GetAttributeValue<EntityReference>("hil_jobid").Name;
                            string model = entity.Contains("hil_model") ? entity.GetAttributeValue<EntityReference>("hil_model").Name : "";
                            string installation_date = entity.GetAttributeValue<DateTime>("hil_registrationdate").ToString("yyyy-MM-dd");
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
                            var res = (SetStateResponse)service.Execute(req);
                            i += 1;
                            Console.WriteLine("Processing... " + i.ToString() + "/" + count.ToString());
                        }
                        catch (Exception ex)
                        {
                            try
                            {
                                Entity entity1 = new Entity(entity.LogicalName, entity.Id);
                                entity1["hil_response"] = ex.Message;
                                service.Update(entity1);
                            }
                            catch (Exception exe)
                            {
                                Console.WriteLine(exe.Message);
                            }
                        }
                    }
                    query.PageInfo.PageNumber += 1;
                    query.PageInfo.PagingCookie = entCol.PagingCookie;
                    entCol = service.RetrieveMultiple(query);
                    count = count + entCol.Entities.Count;
                }
                while (entCol.MoreRecords);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static void RegisterUserOnWhatsApp(string customerName, string customerMobileNumber)
        {
            UserDetails userDetails = new UserDetails();
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
        public static void SendWhatsEvent(string customerMobileNumber, string templateName, string productCategory, string serialNumber, string model, string installation_date)
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
                if (((int)response.StatusCode) != 200 && ((int)response.StatusCode) != 201 && ((int)response.StatusCode) != 202)
                {
                    throw new Exception(obj.message);
                }
                else
                    Console.WriteLine(obj.message);
            }
            catch (Exception ex)
            {

                throw new Exception(ex.Message);
            }
        }
        public static void sendCampaignSMS(string _templateId, string _mobileNumber, string _message)
        {
            string _api = "https://japi.instaalerts.zone/failsafe/HttpLink?aid=640990&pin=w~7Xg)9V&mnumber=" + _mobileNumber + "&signature=HAVELL&message=" + _message + "&dlt_entity_id=110100001483&dlt_template_id=" + _templateId + "&cust_ref=";

            WebRequest request = WebRequest.Create(_api);
            request.Method = "GET";
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
