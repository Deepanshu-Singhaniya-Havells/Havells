using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Security.Cryptography;
using System.Security.Policy;
using System.ServiceModel.Channels;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services.Description;

namespace Havells.CRM.WhatsAppPromotions
{
    internal class Program
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";

        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        static string AquaGuid = "72981D83-16FA-E811-A94C-000D3AF0694E";
        static string ACGUID = "D51EDD9D-16FA-E811-A94C-000D3AF0694E";
        static string messageID = "a28dd97c-1ffb-4fcf-99f1-0b557ed381da";
        static long fromMob = 917428935151;

        static string AC5DayTemplateName = "ac_amc_5d";
        static string callbackData = "D365 Auto AMC Trigger";
        static string buttonPArameter = "Book Water Purifier AMC";
        #endregion

        static string message = @"{
                                    ""messages"": [
                                        {
                                            ""from"": 91{1},
                                            ""to"": 91{2},
                                            ""messageId"": ""{3}"",
                                            ""content"": {
                                                ""templateName"": ""{4}"",
                                                ""templateData"": {
                                                    ""body"": {
                                                        ""placeholders"": []
                                                    },
                                                    ""header"": {
                                                        ""type"": ""TEXT"",
						
                                                        ""placeholder"": ""{5}""
                                                    },
                                                    ""buttons"": [
                                                        {
                                                            ""type"": ""QUICK_REPLY"",
                                                            ""parameter"": ""{6}""
                                                        }
                                                    ]
                                                },
                                                ""language"": ""en""
                                            },
                                            ""callbackData"": ""{7}""
                                        }
                                    ]
                                }";
        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));

            ////retriveJobs(AquaGuid, -5, AC5DayTemplateName, 917428935151, messageID, callbackData, buttonPArameter);
            ////retriveJobs(AquaGuid, -30, AC5DayTemplateName, 917428935151, messageID, callbackData, buttonPArameter);

            retriveJobs(ACGUID, -5, "ac_amc_5d", 917428935151, messageID, callbackData, buttonPArameter);
            retriveJobs(ACGUID, -30, "ac_amc_30d", 917428935151, messageID, callbackData, buttonPArameter);
            retriveJobs(AquaGuid, -5, "amc_wp_5d", 917428935151, messageID, callbackData, buttonPArameter);
            retriveJobs(AquaGuid, -30, "amc_wp_30d", 917428935151, messageID, callbackData, buttonPArameter);
        }

        static void retriveJobs(string ProductCategoryGUID, int dayDif, string templateName, long fromNumber, string messageID, string callbackData, string buttonPArameter)
        {
            try
            {
                string dateString = DateTime.Today.AddDays(dayDif).Year + "-" + DateTime.Today.AddDays(dayDif).Month + "-" + DateTime.Today.AddDays(dayDif).Day;
                string fetch = $@"<fetch version=""1.0"" page=""1"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                    <entity name=""msdyn_workorder"">
                    <attribute name=""msdyn_name"" />
                    <attribute name=""createdon"" />
                    <attribute name=""hil_productsubcategory"" />
                    <attribute name=""hil_mobilenumber"" />
                    <attribute name=""hil_customerref"" />
                    <attribute name=""hil_callsubtype"" />
                    <attribute name=""msdyn_workorderid"" />
                    <order attribute=""msdyn_name"" descending=""false"" />
                    <filter type=""and"">
                        <condition attribute=""hil_productcategory"" operator=""in"">
                        <value>{ProductCategoryGUID}</value>
                        </condition>
                        <condition attribute=""msdyn_timeclosed"" operator=""on"" value=""{dateString}"" />
                        <condition attribute=""hil_callsubtype"" operator=""eq"" value=""{{E3129D79-3C0B-E911-A94E-000D3AF06CD4}}"" />
                    </filter>
                    </entity>
                </fetch>";
                EntityCollection jobsColl = _service.RetrieveMultiple(new FetchExpression(fetch));
                string customerName = string.Empty;
                string customerMobileNumber = string.Empty;
                Console.WriteLine("Total Jobs Found " + jobsColl.Entities.Count);
                int done = 1;
                int error = 1;
                foreach (Entity job in jobsColl.Entities)
                {
                    if (done > 1) { 
                        break; 
                    }
                    try
                    {
                        customerName = job.GetAttributeValue<EntityReference>("hil_customerref").Name;
                        customerMobileNumber = "91" + job.GetAttributeValue<string>("hil_mobilenumber");
                        Console.WriteLine(customerMobileNumber);
                        MsgSend msgSend = new MsgSend();
                        Message message = new Message();
                        message.from = fromNumber;// 917428935151;
                        message.to = 919050673956;//long.Parse(customerMobileNumber);
                        message.messageId = messageID;
                        message.callbackData = callbackData;// "D365 Auto AMC Trigger";

                        Content content = new Content()
                        {
                            templateName = templateName,
                            language = "en"
                        };
                        Header header = new Header()
                        {
                            type = "TEXT",
                            placeholder = customerName
                        };
                        Body body = new Body() { placeholders = new List<object>() };

                        Button button = new Button()
                        {
                            type = "QUICK_REPLY",
                            parameter = buttonPArameter
                        };
                        TemplateData templateData = new TemplateData()
                        {
                            header = header
                        };

                        List<Button> lstbtn = new List<Button>();
                        lstbtn.Add(button);
                        templateData.buttons = lstbtn;
                        templateData.body = body;

                        content.templateData = templateData;

                        message.content = content;

                        List<Message> lstMsg = new List<Message>();
                        lstMsg.Add(message);

                        msgSend.messages= lstMsg;
                        SendMessage(msgSend);
                        Console.WriteLine(done + "/" + jobsColl.Entities.Count + "Msg Send to " + customerMobileNumber);
                        done++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("***** " + error + "/" + jobsColl.Entities.Count + "Error Occured " + ex.Message);
                        error++;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("!!!!!!!!!!!!!!! Error " + ex.Message);
            }
        }
        static void SendMessage(MsgSend msgSend)
        {
            IntegrationConfig intConfig = IntegrationConfiguration(_service, "WhatsApp Promotions");
            string url = intConfig.uri;
            string authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));
            string data = JsonConvert.SerializeObject(msgSend);
            var client = new RestClient(url);
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);

            request.AddHeader("Authorization", "Basic " + authInfo);
            request.AddHeader("Content-Type", "application/json");

            request.AddParameter("application/json", data, ParameterType.RequestBody);
            IRestResponse response = client.Execute(request);

            //Console.WriteLine(response.Content == "" ? response.ErrorMessage : response.Content);
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
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + ":" + integrationConfiguration.GetAttributeValue<string>("hil_password");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:- " + ex.Message);
            }
            return output;
        }

        #region CRM Connection
        static IOrganizationService ConnectToCRM(string connectionString)
        {
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);
                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    Console.WriteLine(((CrmServiceClient)service).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }
            return service;

        }
        #endregion
    }
}
