using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Net.Http.Headers;
using System.Text;
using System.Xml.Linq;

namespace KKGCloseJobsForMoengage
{
    internal class ClsKKGCloseJobsData
    {
        private readonly ServiceClient _service;
        public ClsKKGCloseJobsData(ServiceClient service)
        {
            _service = service;
        }
        public void ClsKKGCloseJobs()
        {
            try
            {
                if (_service.IsReady)
                {
                    Root root = new Root();
                    root.elements = new List<Element>();
                    Element element = new Element();
                    element.attributes = new Attributes();
                    element.attributes.platforms = new List<Platform> { new Platform() };
                    element.actions = new List<Action>();

                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='msdyn_workorder'>
                                        <attribute name='msdyn_name'/>
                                        <attribute name='createdon'/>
                                        <attribute name='hil_productsubcategory'/>
                                        <attribute name='hil_customerref'/>
                                        <attribute name='msdyn_workorderid'/>
                                        <attribute name='msdyn_timeclosed'/>
                                        <attribute name='hil_mobilenumber'/>
                                        <attribute name='hil_productcategory'/>
                                        <attribute name='msdyn_customerasset'/>
                                        <order attribute='msdyn_name' descending='false'/>
                                        <filter type='and'>
                                        <condition attribute='hil_isocr' operator='ne' value='1'/>
                                        <condition attribute='msdyn_substatus' operator='eq' uiname='Closed' uitype='msdyn_workordersubstatus' value='{{1727FA6C-FA0F-E911-A94E-000D3AF060A1}}'/>
                                        <condition attribute='msdyn_timeclosed' operator='last-x-hours' value='1'/>
                                        <condition attribute='hil_kkgcode' operator='not-null'/>
                                        </filter>
                                        <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='msdyn_customerasset' visible='false' link-type='outer' alias='CA'>
                                        <attribute name='msdyn_product'/>
                                        <attribute name='hil_modelname'/>
                                        </link-entity>
                                        <link-entity name='hil_jobsextension' from='hil_jobsextensionid' to='hil_jobextension' link-type='inner' alias='ae'>
                                        <attribute name='hil_jobsextensionid'/>                                       
                                        <filter type='and'>
                                        <condition attribute='hil_moengageeventsent' operator='eq' value='0'/>
                                        </filter>
                                        </link-entity>
                                        </entity>
                                        </fetch>";
                    Console.WriteLine("Getting KKG Jobs details for Moengage");
                    EntityCollection entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entcoll.Entities.Count > 0)
                    {
                        Console.WriteLine("Getting Jobs Count: {0}", entcoll.Entities.Count);
                        foreach (var c in entcoll.Entities)
                        {
                            try
                            {
                                var mobilenumber = c.Contains("hil_mobilenumber") ? c.Attributes["hil_mobilenumber"].ToString() : null;
                                var Name = c.Contains("hil_customerref") ? c.GetAttributeValue<EntityReference>("hil_customerref").Name : null;
                                var category = c.Contains("hil_productcategory") ? c.GetAttributeValue<EntityReference>("hil_productcategory").Name : null;
                                var Subcategory = c.Contains("hil_productsubcategory") ? c.GetAttributeValue<EntityReference>("hil_productsubcategory").Name : null;
                                var SerialNumber = c.Contains("msdyn_customerasset") ? c.GetAttributeValue<EntityReference>("msdyn_customerasset").Name : null;
                                var ModelCode = c.Contains("CA.msdyn_product") ? ((EntityReference)(c.GetAttributeValue<AliasedValue>("CA.msdyn_product").Value)).Name : null;
                                var Productname = c.Contains("CA.hil_modelname") ? c.GetAttributeValue<AliasedValue>("CA.hil_modelname").Value.ToString() : null;
                                var jobId = c.Contains("msdyn_name") ? c.GetAttributeValue<string>("msdyn_name").ToString() : null;
                                var ticket_raised_date = c.Contains("createdon") ? c.GetAttributeValue<DateTime>("createdon").ToString("yyyy-MM-ddTHH:mm:ssZ") : null;
                                var ticket_closed_date = c.Contains("msdyn_timeclosed") ? c.GetAttributeValue<DateTime>("msdyn_timeclosed").ToString("yyyy-MM-ddTHH:mm:ssZ") : null;
                                Guid jobsextensionid = c.Contains("ae.hil_jobsextensionid") ? ((Guid)c.GetAttributeValue<AliasedValue>("ae.hil_jobsextensionid").Value) : Guid.Empty;
                                if (mobilenumber == null)
                                {
                                    Console.WriteLine("Customer Mobile Number not found");
                                    continue;
                                }
                                else if (jobId == null)
                                {
                                    Console.WriteLine("Jobid not found.");
                                    continue;
                                }
                                Console.WriteLine("Detail value for Jobid: {0} and Mobileno.: {1}", jobId, mobilenumber);

                                Attributes _attribute = new Attributes();
                                Platform _platform = new Platform();

                                root.type = "transition";
                                element.type = "customer";
                                element.customer_id = "91" + mobilenumber;
                                element.attributes = new Attributes();
                                element.attributes.mobile_number = "91" + mobilenumber;
                                element.attributes.name = Name;
                                element.attributes.platforms = new List<Platform>();
                                _platform.platform = "D365";
                                _platform.active = "true";
                                element.attributes.platforms.Add(_platform);
                                element.type = "event";
                                element.customer_id = "91" + mobilenumber;

                                element.actions = new List<Action>();
                                Action _action = new Action();

                                _action.action = "Consumer_KKGCode";
                                _action.attributes = new Attributes();
                                _action.attributes.product_category = category;
                                _action.attributes.product_name = Productname;
                                _action.attributes.product_sub_category = Subcategory;
                                _action.attributes.serial_number = SerialNumber;
                                _action.attributes.model_code = ModelCode;
                                _action.attributes.sr_id = jobId;
                                _action.attributes.ticket_raised_date = ticket_raised_date;
                                _action.attributes.ticket_closed_date = ticket_closed_date;

                                _action.platform = "D365";
                                _action.app_version = "1.0";
                                _action.user_timezone_offset = 0;
                                _action.current_time = DateTime.UtcNow.AddMinutes(330).ToString("yyyy-MM-ddTHH:mm:ssZ");
                                element.actions.Add(_action);

                                root.elements.Add(element);

                                Console.WriteLine("Moengage API call");
                                var Paramdata = new StringContent(JsonConvert.SerializeObject(root), System.Text.Encoding.UTF8, "application/json");
                                QueryExpression qe = new QueryExpression("hil_integrationconfiguration");
                                qe.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                                qe.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "KKGCloseJobsForMoengage");
                                Entity enColl = _service.RetrieveMultiple(qe)[0];
                                String URL = enColl.GetAttributeValue<string>("hil_url");
                                String User = enColl.GetAttributeValue<string>("hil_username");
                                string Pass = enColl.GetAttributeValue<string>("hil_password");
                                HttpClient Reqclient = new HttpClient();
                                Reqclient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue(User, Pass);
                                HttpResponseMessage Result = Reqclient.PostAsync(URL, Paramdata).Result;
                                Console.WriteLine("Getting Moengage API Result");
                                if (Result.IsSuccessStatusCode)
                                {
                                    Console.WriteLine("Getting Moengage API response {0}", Result.IsSuccessStatusCode);
                                    var obj = JsonConvert.DeserializeObject<Response>(Result.Content.ReadAsStringAsync().Result);
                                    if (obj.status == "success")
                                    {
                                        Console.WriteLine("Update jobsextension entity for moengageeventsent");
                                        Entity entinvoice = new Entity("hil_jobsextension", jobsextensionid);
                                        entinvoice["hil_moengageeventsent"] = true;
                                        _service.Update(entinvoice);
                                    }
                                }
                            }
                            catch (Exception)
                            {

                                continue;
                            }

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






    public class Root
    {
        public string type { get; set; }
        public List<Element> elements { get; set; }
    }
    public class Element
    {
        public string type { get; set; }
        public string customer_id { get; set; }
        public Attributes attributes { get; set; }
        public List<Action> actions { get; set; }
    }
    public class Action
    {
        public string action { get; set; }
        public Attributes attributes { get; set; }
        public string platform { get; set; }
        public string app_version { get; set; }
        public int user_timezone_offset { get; set; }
        public string current_time { get; set; }
    }

    public class Attributes
    {
        public string mobile_number { get; set; }
        public string name { get; set; }
        public List<Platform> platforms { get; set; }
        public string product_category { get; set; }
        public string product_name { get; set; }
        public string product_sub_category { get; set; }
        public string serial_number { get; set; }
        public string model_code { get; set; }

        [JsonProperty("sr_id ")]
        public string sr_id { get; set; }
        public string ticket_raised_date { get; set; }
        public string ticket_closed_date { get; set; }
    }
    public class Platform
    {
        public string platform { get; set; }
        public string active { get; set; }
    }

    public class Response
    {
        public string status { get; set; }
        public string message { get; set; }
    }

}