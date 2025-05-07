using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System.Text.RegularExpressions;


namespace WACTA
{
    public class ClsWACTA
    {
        private readonly ServiceClient service;
        public ClsWACTA(ServiceClient _service)
        {
            service = _service;
        }
        public void SentDataForNewWACTA(string triggerName)
        {
            try
            {
                if (service.IsReady)
                {
                    EntityCollection entcoll = null;
                    int _rowCount = 0;
                    int _totalCount = 0;
                    string _fetchXML = string.Empty;
                    int PageNumber = 1;
                    do
                    {
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' page='{PageNumber++}'>
                            <entity name='msdyn_workorder'>
                            <attribute name='msdyn_name' />
                            <attribute name='hil_productsubcategory' />
                            <attribute name='hil_mobilenumber' />
                            <attribute name='hil_customerref' />
                            <attribute name='msdyn_customerasset' />
                            <attribute name ='hil_productcategory' />
                            <attribute name ='createdon' />
                            <filter type='and'>
                            <condition attribute='msdyn_substatus' operator='eq' value='{{2927FA6C-FA0F-E911-A94E-000D3AF060A1}}' />
                            <condition attribute='hil_callsubtype' operator='ne' value='{{55A71A52-3C0B-E911-A94E-000D3AF06CD4}}' />
                            <condition attribute='hil_kkgcode' operator='null' />
                            <condition attribute='hil_jobclosuredon' operator='yesterday' />
                            </filter>
                            <link-entity name='hil_wacta' from='hil_jobid' to='msdyn_workorderid' link-type='outer' alias='aa' />
                            <filter type='and'>
                                <condition entityname='aa' attribute='hil_jobid' operator='null' />
                            </filter>
                            <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='msdyn_customerasset' link-type='inner' alias='al'>
                            <link-entity name='annotation' from='objectid' to='msdyn_customerassetid' link-type='outer' alias='am' />
                            </link-entity>
                            <filter type='and'>
                            <condition entityname='am' attribute='objectid' operator='null' />
                            </filter>
                            <link-entity name='contact' from='contactid' to='hil_customerref' visible='false' link-type='inner' alias='cust'>
                                <attribute name='fullname' />
                            </link-entity>
                            </entity>
                            </fetch>";
                        entcoll = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entcoll.Entities.Count > 0)
                        {
                            _totalCount += entcoll.Entities.Count;
                            Root ParamModelData = null;
                            Template template = null;
                            foreach (Entity ent in entcoll.Entities)
                            {
                                try
                                {
                                    ++_rowCount;
                                    string? UserName = ent.Contains("cust.fullname") ? FixName(ent.GetAttributeValue<AliasedValue>("cust.fullname").Value.ToString()) : null;
                                    string? MobileNumber = ent.Contains("hil_mobilenumber") ? GetLast10Characters(ent.GetAttributeValue<string>("hil_mobilenumber")) : null;
                                    string JobNumber = ent.GetAttributeValue<string>("msdyn_name");
                                    Console.WriteLine($"Processing: {_rowCount}/{_totalCount} for Job#{JobNumber} Mobile Number: {MobileNumber}");
                                    Console.WriteLine($"Checking for previous WA CTA Trigger.");
                                    if (!ExistJobIdInWACTA(ent.Id) && UserName != null && MobileNumber != null)
                                    {
                                        ParamModelData = new Root();
                                        ParamModelData.phoneNumber = MobileNumber;
                                        ParamModelData.callbackData = "kkg_audit";

                                        template = GetTemplatebyTriggerName(triggerName, JobNumber, UserName);
                                        ParamModelData.template = template;

                                        var RESULT = UpdateWhatsAppUser(ParamModelData);
                                        if (RESULT.result == "true")
                                        {
                                            Console.WriteLine($"Registered Consumer's mobile number {MobileNumber} on Whatsapp.");

                                            var ParamData = new TrackEvents();
                                            ParamData.traits = new trait();
                                            ParamData.phoneNumber = MobileNumber;
                                            ParamData.traits.name = UserName;
                                            ParamData.countryCode = "+91";
                                            ParamData.@event = triggerName;
                                            var resultMsg = SendMessage(ParamData);

                                            if (resultMsg.result == true)
                                            {
                                                DateTime _validUpto = DateTime.Now.AddDays(1).Date;
                                                Entity _entCreate = new Entity("hil_wacta");
                                                _entCreate["hil_jobid"] = ent.ToEntityReference();
                                                _entCreate["hil_triggercount"] = 1;
                                                _entCreate["hil_validupto"] = new DateTime(_validUpto.Year, _validUpto.Month, _validUpto.Day, 07, 30, 00);
                                                _entCreate["hil_wastatusreason"] = new OptionSetValue(1);//Open
                                                _entCreate["hil_watriggername"] = triggerName;
                                                service.Create(_entCreate);
                                            }
                                            else
                                            {
                                                Console.WriteLine($"Error: While Whatsapp Invite sent on mobile number {MobileNumber}: {resultMsg.message}");
                                            }
                                        }
                                        else
                                        {
                                            Console.WriteLine($"Error: While Registering  Consumer's mobile number {MobileNumber} on Whatsapp: {RESULT.message}");
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error: {ex.Message}");
                                }
                            }
                        }
                    }
                    while (entcoll.MoreRecords);
                    Console.WriteLine("Batch Ends.");
                }
                else
                {
                    Console.WriteLine("Error: D365 Service is unavailable.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
        public void SentDataForReminderWACTA(string triggerName)
        {
            try
            {
                if (service.IsReady)
                {
                    DateTime _validUpto = DateTime.Now.Date;
                    DateTime _proccessDate = new DateTime(_validUpto.Year, _validUpto.Month, _validUpto.Day, 23, 59, 59);

                    Console.WriteLine($"Batch Starts: Proccess Date: {_proccessDate}");
                    EntityCollection entcoll = null;
                    int _rowCount = 1;
                    int _totalCount = 0;
                    string _fetchXML = string.Empty;
                    int PageNumber = 1;
                    do
                    {
                        //<condition attribute='hil_validupto' operator='on-or-before' value='{ProccessDate}' /> 

                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' page='{PageNumber++}'>
                            <entity name='hil_wacta'>
                            <attribute name='hil_wactaid'/>
                            <attribute name='hil_watriggername'/>
                            <attribute name='createdon'/>
                            <attribute name='hil_whatsappctaresponse'/>
                            <attribute name='hil_wastatusreason'/>
                            <attribute name='hil_validupto'/>
                            <attribute name='hil_triggercount'/>
                            <attribute name='hil_jobid'/>
                            <order attribute='hil_watriggername' descending='false'/>
                            <filter type='and'>
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_whatsappctaresponse' operator='null' />
                                <condition attribute='hil_validupto' operator='lt' value='{_proccessDate.ToString("yyyy-MM-dd hh:mm:ss")}' /> 
                                <condition attribute='hil_triggercount' operator='eq' value='1' />
                                <condition attribute='hil_watriggername' operator='eq' value='{triggerName}' />
                                <condition attribute='hil_wastatusreason' operator='eq' value='1' />
                            </filter>
                            </entity>
                            </fetch>";
                        entcoll = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        Console.WriteLine($"entcoll.Entities.Count: {entcoll.Entities.Count}");

                        if (entcoll.Entities.Count > 0)
                        {
                            Console.WriteLine($"Total WhatsApp CTA Reminders found to be sent: {entcoll.Entities.Count}");
                            _totalCount += entcoll.Entities.Count;
                            foreach (Entity ent in entcoll.Entities)
                            {
                                try
                                {
                                    int triggerCount = ent.GetAttributeValue<int>("hil_triggercount");
                                    DateTime validUpto = ent.GetAttributeValue<DateTime>("hil_validupto").AddMinutes(330);

                                    Entity Jobs = service.Retrieve("msdyn_workorder", ent.GetAttributeValue<EntityReference>("hil_jobid").Id, new ColumnSet("hil_mobilenumber", "hil_customerref", "msdyn_name", "hil_productcategory"));
                                    string MobileNumber = Jobs.GetAttributeValue<string>("hil_mobilenumber");
                                    string JobNumber = Jobs.GetAttributeValue<string>("msdyn_name");
                                    EntityReference Customer = Jobs.GetAttributeValue<EntityReference>("hil_customerref");
                                    string UserName = Customer.Name;
                                    Console.WriteLine($"Record# {_rowCount++}/{_totalCount} Job# {JobNumber} Mobile Number# {MobileNumber} validUpto: {validUpto}");

                                    if (UserName != null && MobileNumber != null)
                                    {
                                        Root ParamModelData = new Root();
                                        ParamModelData.phoneNumber = MobileNumber;

                                        Template template = GetTemplatebyTriggerName(triggerName, JobNumber, UserName);
                                        ParamModelData.template = template;

                                        var RESULT = UpdateWhatsAppUser(ParamModelData);
                                        if (RESULT.result == "true")
                                        {
                                            Console.WriteLine($"Registered Consumer's mobile number {MobileNumber} on Whatsapp.");

                                            var ParamData = new TrackEvents();
                                            ParamData.traits = new trait();
                                            ParamData.phoneNumber = MobileNumber;
                                            ParamData.traits.name = UserName;
                                            ParamData.countryCode = "+91";
                                            ParamData.@event = triggerName;

                                            var resultMsg = SendMessage(ParamData);
                                            if (resultMsg.result == true)
                                            {
                                                Console.WriteLine($"Current datetime: {DateTime.Now}");

                                                triggerCount = triggerCount + 1;
                                                ent["hil_triggercount"] = triggerCount;
                                                ent["hil_validupto"] = validUpto.AddHours(24);
                                                Console.WriteLine($"Reminder: {validUpto.AddHours(24)}");
                                                service.Update(ent);
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error: {ex.Message}");
                                }
                            }
                        }
                    }
                    while (entcoll.MoreRecords);
                    Console.WriteLine("Batch Ends.");
                }
                else
                {
                    Console.WriteLine("Error: D365 Service is unavailable.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
        private string GetLast10Characters(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return string.Empty;
            }
            return input.Length <= 10 ? input : input.Substring(input.Length - 10);
        }
        private string FixName(string input)
        {
            // Remove leading and trailing spaces
            input = input.Trim();

            // Remove leading dot
            if (input.StartsWith("."))
            {
                input = input.Substring(1);
            }

            // Remove trailing dot
            if (input.EndsWith("."))
            {
                input = input.Substring(0, input.Length - 1);
            }

            // Replace multiple dots or spaces with a single one
            input = Regex.Replace(input, @"[.]{2,}", ".");
            input = Regex.Replace(input, @"[ ]{2,}", " ");

            // Remove dot-space and space-dot combinations
            input = Regex.Replace(input, @"[. ]{2}", "");

            // Ensure no dot-space or space-dot combinations
            input = Regex.Replace(input, @"\.\s", ".");
            input = Regex.Replace(input, @"\s\.", " ");

            // Remove any invalid characters
            input = Regex.Replace(input, @"[^A-Za-z. ]", "");

            return input;
        }
        private UpdateAppUserRes UpdateWhatsAppUser(Root userDetails)
        {
            IntegrationConfig intConfig = IntegrationConfiguration("Send WA Trigger");
            string url = intConfig.uri;
            string authInfo = intConfig.Auth;
            string data = JsonConvert.SerializeObject(userDetails);
            Console.WriteLine(data);
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
        private sendMesgRes SendMessage(TrackEvents campaignDetails)
        {
            IntegrationConfig intConfig = IntegrationConfiguration("WhatsApp Campaign Details");
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
        private IntegrationConfig IntegrationConfiguration(string Param)
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
        private bool ExistJobIdInWACTA(Guid jobId)
        {
            try
            {
                if (service.IsReady)
                {
                    string fetchXML = $@"
                        <fetch top='1'>
                        <entity name='hil_wacta'>
                        <attribute name='hil_jobid' />
                            <filter>
                                 <condition attribute='hil_jobid' operator='eq' value='{jobId}' />
                            </filter>
                        </entity>
                        </fetch>";

                    EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetchXML));
                    return result.Entities.Count > 0;
                }
                else
                {
                    Console.WriteLine("Error: D365 Service is unavailable.");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return false;
            }
        }
        private Template GetTemplatebyTriggerName(string triggerName, string JobNumber, string UserName)
        {
            Template template = new Template();
            switch (triggerName)
            {
                case "cc_b2c_kkg":
                    {
                        template.name = triggerName;
                        template.headerValues = new List<string> { UserName };
                        template.bodyValues = new List<string> { JobNumber };
                        template.buttonPayload = new Dictionary<string, List<string>>
                        {
                            { "0", new List<string>  { $"Yes, KKG Audit SR: {JobNumber} [{triggerName}]" } } ,
                            { "1", new List<string> { $"No, KKG Audit SR: {JobNumber} [{triggerName}]" } }
                        };
                    }
                    break;
            }
            return template;
        }
    }
    public class UpdateAppUserRes
    {
        public string result { get; set; }
        public string message { get; set; }
    }
    public class IntegrationConfig
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }
    public class TrackEvents
    {
        public string phoneNumber { get; set; }
        public string countryCode { get; set; }
        public string @event { get; set; }
        public trait traits { get; set; }
    }
    public class trait
    {
        public string name { get; set; }
        public string category { get; set; }
        public string prdSerial { get; set; }
    }
    public class sendMesgRes
    {
        public bool result { get; set; }
        public string message { get; set; }
        public string id { get; set; }
    }

    public class Root
    {
        public string countryCode { get; set; } = "+91";
        public string phoneNumber { get; set; }
        public string callbackData { get; set; } = "kkg audit";
        public string type { get; set; } = "Template";
        public Template template { get; set; }
    }
    public class Template
    {
        public string name { get; set; }
        public string languageCode { get; set; } = "en";
        public List<string> headerValues { get; set; }
        public List<string> bodyValues { get; set; }
        public Dictionary<string, List<string>> buttonPayload { get; set; }
    }
}
