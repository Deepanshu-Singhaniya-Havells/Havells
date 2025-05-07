using HavellsNewPlugin.Helper;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;

namespace HavellsNewPlugin.PhoneCall
{
    public class PreCreatePhoneCall : IPlugin
    {
        public static ITracingService tracingService = null;
        public string JobID;
        public int preference;
        private bool isPreference = false;
        public string callingNumber;
        IPluginExecutionContext context;
        public void Execute(IServiceProvider serviceProvider)
        {
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #region MainRegion
            tracingService.Trace("On the top");
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.PrimaryEntityName.ToLower() == "phonecall" && context.MessageName.ToUpper() == "CREATE")
                {
                    tracingService.Trace("I am here");
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entityPhoneCall = (Entity)context.InputParameters["Target"];
                    Guid _ownerId = (entityPhoneCall["ownerid"] as EntityReference).Id;
                    tracingService.Trace(_ownerId.ToString());

                    Entity entity = service.Retrieve(entityPhoneCall.LogicalName, entityPhoneCall.Id, new ColumnSet("ownerid", "regardingobjectid", "hil_contactpreference", "hil_callingnumber"));

                    if (entity.Contains("regardingobjectid"))
                    {
                        JobID = entity.GetAttributeValue<EntityReference>("regardingobjectid").Name;
                    }
                    if (entity.Contains("ownerid"))
                    {
                        _ownerId = entity.GetAttributeValue<EntityReference>("ownerid").Id;
                    }
                    if (entity.Contains("hil_contactpreference"))
                    {
                        preference = entity.GetAttributeValue<OptionSetValue>("hil_contactpreference").Value;
                        isPreference = true;
                    }
                    if (entity.Contains("hil_callingnumber"))
                    {
                        callingNumber = entity.GetAttributeValue<string>("hil_callingnumber");
                    }

                    if (!string.IsNullOrEmpty(JobID) && entity.GetAttributeValue<EntityReference>("regardingobjectid").LogicalName == "msdyn_workorder" && isPreference)
                    {
                        string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                          <entity name='systemuser'>
                            <attribute name='systemuserid' />
                            <filter type='and'>
                              <condition attribute='systemuserid' operator='eq' value='{_ownerId}' />
                            </filter>
                            <link-entity name='systemuserroles' from='systemuserid' to='systemuserid' visible='false' intersect='true'>
                              <link-entity name='role' from='roleid' to='roleid' alias='ae'>
                                <filter type='and'>
                                  <condition attribute='name' operator='eq' value='{HelperClass._callMaskingRoleName}' />
                                </filter>
                              </link-entity>
                            </link-entity>
                          </entity>
                        </fetch>";
                        EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (_entCol.Entities.Count > 0)
                        {
                            string result = C2CAPI(service, entity);
                        }
                        else
                        {

                        }
                    }
                    else
                    {

                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            #endregion
        }
        public string C2CAPI(IOrganizationService service, Entity entityCallMaking)
        {
            try
            {
                Entity entity = new Entity(entityCallMaking.LogicalName, entityCallMaking.Id);

                tracingService.Trace("JobID : " + JobID + " , preference : " + preference + " , callingNumber : " + callingNumber);
                if (string.IsNullOrEmpty(callingNumber))
                {
                    throw new InvalidPluginExecutionException("Calling Nubmer cannot be blank");
                }
                IntegrationConfig intConfig = IntegrationConfiguration(service, "Airtel IQ C2C");
                JObject details = JObject.Parse(intConfig.description);
                string apiEndPoint = details.Value<string>("apiEndPoint");
                string authorization = details.Value<string>("authorization");
                string callerId = details.Value<string>("callerId");
                string mergingStrategy = details.Value<string>("mergingStrategy");
                Root root = new Root();
                root.callFlowId = details.Value<string>("callFlowId");
                root.customerId = details.SelectToken("customerId").ToString();
                root.callType = "OUTBOUND";
                CallFlowConfiguration callflow = new CallFlowConfiguration();
                InitiateCall1 ini = new InitiateCall1();
                ini.callerId = callerId;
                ini.mergingStrategy = mergingStrategy;
                ini.maxTime = 0;
                List<Participant> par1 = new List<Participant>();
                Participant p1 = new Participant();
                p1.participantAddress = callingNumber;
                p1.callerId = callerId;
                p1.participantName = service.Retrieve("systemuser", context.UserId, new ColumnSet("fullname")).GetAttributeValue<string>("fullname");
                p1.maxRetries = 1;
                p1.maxTime = 0;
                par1.Add(p1);
                ini.participants = par1;
                callflow.initiateCall_1 = ini;
                List<Participant> par2 = new List<Participant>();
                Participant p2 = new Participant();
                AddParticipant1 addpar = new AddParticipant1();
                addpar.mergingStrategy = mergingStrategy;
                addpar.maxTime = 0;

                Tuple<string, string> tocall = getPhoneNumber(service, JobID, preference);
                if (tocall.Item1 == "Error")
                {
                    throw new InvalidPluginExecutionException("Deepanshu_Plugin.Appointment.C2CAPI.getPhoneNumber: " + tocall.Item2);
                }

                p2.participantAddress = tocall.Item2;
                p2.participantName = tocall.Item1;
                p2.maxRetries = 1;
                p2.maxTime = 0;
                par2.Add(p2);
                addpar.participants = par2;
                callflow.addParticipant_1 = addpar;
                root.callFlowConfiguration = callflow;
                tracingService.Trace("Calling API");

                StatusResponse obj = new StatusResponse();
                using (HttpClient client = new HttpClient())
                {
                    var data = new StringContent(JsonConvert.SerializeObject(root), Encoding.UTF8, "application/json");
                    client.DefaultRequestHeaders.Add("Authorization", authorization);//.Authorization = new AuthenticationHeaderValue(authInfo);
                    HttpResponseMessage response = client.PostAsync(apiEndPoint, data).Result;
                    if (response.IsSuccessStatusCode)
                    {
                        obj = JsonConvert.DeserializeObject<StatusResponse>(response.Content.ReadAsStringAsync().Result);
                        tracingService.Trace(obj.status);
                        tracingService.Trace(obj.correlationId);

                        tracingService.Trace("Testing");
                        entity["subject"] = obj.correlationId;
                        entity["phonenumber"] = callerId;
                        entity["hil_calledtonum"] = p2.participantAddress;
                        entity["directioncode"] = true;
                        entity["hil_callingnumber"] = p1.participantAddress;
                        if (p1.participantAddress.StartsWith("+91"))
                        {
                            p1.participantAddress = p1.participantAddress.Substring(2);
                        }
                        tracingService.Trace("some of the fields are updated");
                        QueryExpression query = new QueryExpression("systemuser");
                        query.ColumnSet = new ColumnSet("mobilephone");
                        query.Criteria.AddCondition("mobilephone", ConditionOperator.EndsWith, p1.participantAddress);
                        EntityCollection usersCollection = service.RetrieveMultiple(query);
                        if (usersCollection.Entities.Count > 0)
                        {
                            tracingService.Trace("1");
                            Entity toParty = new Entity("activityparty");
                            toParty["partyid"] = new EntityReference("systemuser", usersCollection.Entities[0].Id);
                            entity["from"] = new Entity[] { toParty };
                        }
                        tracingService.Trace("set the from ");

                        if (p2.participantAddress.StartsWith("+91"))
                        {
                            p2.participantAddress = p2.participantAddress.Substring(2);
                        }
                        query = new QueryExpression("contact");
                        query.ColumnSet = new ColumnSet("mobilephone");
                        query.Criteria.AddCondition("mobilephone", ConditionOperator.EndsWith, p2.participantAddress);
                        usersCollection = service.RetrieveMultiple(query);
                        if (usersCollection.Entities.Count > 0)
                        {
                            tracingService.Trace("2");
                            Entity fromParty = new Entity("activityparty");
                            fromParty["partyid"] = new EntityReference("contact", usersCollection.Entities[0].Id);
                            entity["to"] = new Entity[] { fromParty };
                        }
                        service.Update(entity);
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException("Something went wrong in the API. Error Code : " + (int)response.StatusCode);
                    }
                }
                return "Phone Call is initiated Successfully";
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("ERROR!!! " + ex.Message);
            }
        }
        public static Tuple<string, string> getPhoneNumber(IOrganizationService service, string JobID, int pref)
        {
            try
            {
                QueryExpression query = new QueryExpression("msdyn_workorder");
                query.ColumnSet = new ColumnSet("hil_mobilenumber", "hil_callingnumber", "hil_alternate", "hil_customername");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, JobID);
                Entity result = service.RetrieveMultiple(query).Entities[0];
                string regNumber = result.GetAttributeValue<string>("hil_mobilenumber");
                string callingNumber = result.GetAttributeValue<string>("hil_callingnumber");
                string alternateNumber = result.GetAttributeValue<string>("hil_alternate");
                string customerName = result.GetAttributeValue<string>("hil_customername");
                string b = string.Empty;
                if (pref == 0)
                {
                    b = regNumber;
                }
                else if (pref == 1)
                {
                    b = callingNumber;
                }
                else if (pref == 2)
                {
                    b = alternateNumber;
                }
                return new Tuple<string, string>(customerName, b);
            }
            catch (Exception ex)
            {
                return new Tuple<string, string>("Error", ex.Message);
            }
        }
        public static IntegrationConfig IntegrationConfiguration(IOrganizationService service, string Param)
        {
            IntegrationConfig output = new IntegrationConfig();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password", "hil_description");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                output.description = integrationConfiguration.GetAttributeValue<string>("hil_description");
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error:- " + ex.Message);
            }
            return output;
        }
        public class IntegrationConfig
        {
            public string uri { get; set; }
            public string description { get; set; }
        }
        public class Root
        {
            public string callFlowId { get; set; }
            public string customerId { get; set; }
            public string callType { get; set; }
            public CallFlowConfiguration callFlowConfiguration { get; set; }
        }
        public class CallFlowConfiguration
        {
            public InitiateCall1 initiateCall_1 { get; set; }
            public AddParticipant1 addParticipant_1 { get; set; }
        }
        public class InitiateCall1
        {
            public string callerId { get; set; }
            public string mergingStrategy { get; set; }
            public List<Participant> participants { get; set; }
            public int maxTime { get; set; }
        }
        public class AddParticipant1
        {
            public string mergingStrategy { get; set; }
            public int maxTime { get; set; }
            public List<Participant> participants { get; set; }
        }
        public class Participant
        {
            public string participantAddress { get; set; }
            public string callerId { get; set; }
            public string participantName { get; set; }
            public int maxRetries { get; set; }
            public int maxTime { get; set; }
            public bool enableEarlyMedia { get; set; } = true;
        }
        public class StatusResponse
        {
            public string status { get; set; }
            public string correlationId { get; set; }
        }
    }
}
