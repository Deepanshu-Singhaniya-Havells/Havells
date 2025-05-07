using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using D365WebJobs.WhatsAppCampaigns;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Net;
using System.Text;
using System.IO;
using Microsoft.Xrm.Sdk.Query;
using System.Text.RegularExpressions;
using NLog;
using System.Linq;

namespace D365WebJobs
{
    public class Program
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";
        //private const string connStr = "AuthType=ClientSecret;url={0};ClientId=bc6676d6-1387-4dc4-be89-ba13b08ceb4e;ClientSecret=73P7Q~sWxupzl4j8-B55y5g3QNosxhkjkV6Q2";

        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        #endregion

        private static Logger logger = LogManager.GetCurrentClassLogger();

        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady) {
                //BulkRecordDeletion.BulkDataDeletionJobProduct(_service);
                //BulkRecordDeletion.BulkDataDeletionJobService(_service);
                //BulkRecordDeletion.BulkDataDeletionJobProductBetweenRange(_service);
                //BulkRecordDeletion.BulkDataDeletionJobServiceBetweenRange(_service);

                string Msg, Mob,_templateId;
                hil_smsconfiguration Conf =new hil_smsconfiguration();
                Msg = $@"Dear Mr. P R Sajimon ,Your KKG Code for Service Request ID ""22042429529985"" is ""<KKGCode>"".Share KKG Code to Technician only after Satisfactory Completion of your Service Request. - Havells";
                Mob = "8285906486";
                _templateId = "1107161191448698079";

                SendSMSViaAirtelIQ(Msg, Mob, Conf, _templateId);

            }

            #region Backup
                //logger.Error("This is an error message");
                //Console.Read();
                //_service = ConnectToCRM(string.Format(connStr, "https://orga23838be.crm11.dynamics.com"));
                //ServiceBOMResponse obj = JobServiceProductAction.PopulateSpareParts(new Guid("8af90e5e-0b73-ee11-8179-6045bdac55a8"), _service);
                //ServiceBOMResponse obj = JobServiceProductAction.PopulateSpareParts(new Guid("54cdd144-985e-ed11-9562-6045bdac5292"), _service);
                //
                //Console.WriteLine(obj._partList.Count.ToString());
                //Process(_service, new Guid("646b4f51-0140-ee11-bdf3-6045bd0e4156"));
                //if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
                //{
                //    WhatsappCampaign obj = new WhatsappCampaign(_service);

                //    //obj.retriveJobs(ModeOfCommunication.Whatsapp, "14D_AC_AMC", -14, "D51EDD9D-16FA-E811-A94C-000D3AF0694E"); // Whatsapp Campaign
                //    obj.retriveJobs(ModeOfCommunication.Whatsapp, "14D_WP_AMC", -14, "72981d83-16fa-e811-a94c-000d3af0694e"); // Whatsapp Campaign

                //    obj.retriveJobs(ModeOfCommunication.SMS, "1107168654225056811 ", -30, "D51EDD9D-16FA-E811-A94C-000D3AF0694E");// SMS
                //    obj.retriveJobs(ModeOfCommunication.SMS, "1107168654305357842", -30, "72981d83-16fa-e811-a94c-000d3af0694e");// SMS
                //}
                //string Msg = @"Dear  M.Ramulu ,Your KKG Code for Service Request ID 13052321472655 is < KKGCode > .Share KKG Code to Technician only after Satisfactory Completion of your Service Request. - Havells";
                //string Mob = "8285906486";
                //string _templateId = "1107161191448698079";
                //hil_smsconfiguration Conf = new hil_smsconfiguration();
                //SendSMSViaAirtelIQResponse _retObj = SendSMSViaAirtelIQ(Msg, Mob, Conf, _templateId);
                //MFRServiceJobs _retObj = new MFRServiceJobs(_service);
                //_retObj.CreateServiceCallRequest(new JobRequestDTO()
                //{
                //    customer_firstname = "Kuldeep Khare",
                //    customer_mobileno = "8285906486",
                //    address_line1="Sector 4",
                //    pincode="201304",
                //    callertype="Dealer",
                //    callsubtype= "Installation",
                //    dealercode="Dummy Dealer Code for Testing",
                //    natureofcomplaint= "Installation",
                //    productsubcategory= "STORAGE WATER HEATER HAVELLS",
                //    expecteddeliverydate= "25-12-2022"
                //});

                //var cardNumber = "8285906486";
                //var firstDigits = cardNumber.Substring(0, 6);
                //var lastDigits = cardNumber.Substring(cardNumber.Length - 4, 4);
                //var requiredMask = new String('X', cardNumber.Length - firstDigits.Length - lastDigits.Length);
                //var maskedString = string.Concat(firstDigits, requiredMask, lastDigits);
                //var maskedCardNumberWithSpaces = Regex.Replace(maskedString, ".{4}", "$0 ");
                //Console.WriteLine(maskedCardNumberWithSpaces);
                #endregion
        }
        //private static UpdateRentalBookingHeader() {
        //    QueryExpression query = null;
        //    EntityCollection entcoll = null;

        //    query = new QueryExpression("ogre_rentalorderline");
        //    query.ColumnSet = new ColumnSet("ogre_rentalorder", "ogre_vehiclecategory", "ogre_vehicleproperties", "ogre_bodytype", "ogre_unit", "ogre_rate","");
        //    query.Criteria = new FilterExpression(LogicalOperator.And);
        //    query.Criteria.AddCondition("ogre_rentalorderlineid", ConditionOperator.Equal, new Guid("4482D3E4-FC0D-EE11-8F6E-002248C6F659"));
        //    entcoll = _service.RetrieveMultiple(query);
        //}
        private static SendSMSViaAirtelIQResponse SendSMSViaAirtelIQ(string Msg, string Mob, hil_smsconfiguration Conf, string _templateId)
        {
            SendSMSViaAirtelIQResponse _retObj = new SendSMSViaAirtelIQResponse() { ifOkay = "NotOkay", responseFromServer = "ERROR" };
            AirtelIQRequest airtelIQRequest = new AirtelIQRequest()
            {
                customerId = "Havells",
                destinationAddress = new List<string>() { Mob },
                message = Msg,
                sourceAddress = "HAVELL",
                messageType = "SERVICE_IMPLICIT",
                dltTemplateId = _templateId,
                metaData = new MetaData { Key = "Value" }
            };

            string BaseUrl = "https://openapi.airtel.in/gateway/airtel-iq-sms-utility/sendSms";
            airtelIQRequest.entityId = "110100001483";
            var json = JsonConvert.SerializeObject(airtelIQRequest);
            WebRequest request = null;
            WebResponse response = null;
            string IfOkay = string.Empty;
            string responseFromServer = string.Empty;
            try
            {
                request = WebRequest.Create(BaseUrl);
                request.Method = "POST";
                byte[] data = Encoding.ASCII.GetBytes(json);
                request.ContentType = "application/json";
                request.ContentLength = data.Length;
                request.Headers.Add("Authorization", "Basic UG9saWN5QmF6YWFyOkZtQThGSFYzQEJ3RHUmNms=");
                Stream requestStream = request.GetRequestStream();
                requestStream.Write(data, 0, data.Length);
                response = (HttpWebResponse)request.GetResponse();
                IfOkay = ((HttpWebResponse)response).StatusCode.ToString();
                Stream responseStream = response.GetResponseStream();
                StreamReader myStreamReader = new StreamReader(responseStream, Encoding.Default);
                responseFromServer = myStreamReader.ReadToEnd();

                myStreamReader.Close();
                responseStream.Close();
                response.Close();
                _retObj = new SendSMSViaAirtelIQResponse()
                {
                    ifOkay = IfOkay,
                    responseFromServer = responseFromServer
                };
            }
            catch (WebException we)
            {
                Conf.StatusCode = new OptionSetValue(910590001);
                Conf["hil_responsefromserver"] = "GATEWAY BUSY API Url " + BaseUrl + " PayLoad: " + json;
            }
            return _retObj;
        }
        #region App Setting Load/CRM Connection
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
        static IOrganizationService ConnectToOGRECRM()
        {
            IOrganizationService service = null;
            const string connStr = "AuthType=ClientSecret;url=https://ogre-rental-dev.crm11.dynamics.com;ClientId=bc6676d6-1387-4dc4-be89-ba13b08ceb4e;ClientSecret=73P7Q~sWxupzl4j8-B55y5g3QNosxhkjkV6Q2";

            try
            {
                service = new CrmServiceClient(connStr);
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

        static void Process(IOrganizationService service, Guid jobId)
        {
            ITracingService tracingService;
            Entity _job = service.Retrieve("ogre_workorder", jobId,
                    new ColumnSet("ogre_name", "ogre_jobtype", "ogre_assetid", "ogre_resourceassigned", "ogre_jobstartdate", "ogre_jobduration"));

            QueryExpression queryExp = new QueryExpression("ogre_jobtypeactivity");
            queryExp.ColumnSet = new ColumnSet("ogre_executionindex", "ogre_jobtype", "ogre_jobstart", "ogre_jobend", "ogre_activity", "ogre_inspectionapplicable", "ogre_drivingapplicable");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("ogre_jobtype", ConditionOperator.Equal, _job.GetAttributeValue<EntityReference>("ogre_jobtype").Id);
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); // Active
            queryExp.AddOrder("ogre_executionindex", OrderType.Ascending);
            EntityCollection entCol = service.RetrieveMultiple(queryExp);

            if (entCol.Entities.Count > 0)
            {
                int totalJobDuration = 0;

                foreach (Entity jobActivity in entCol.Entities)
                {

                    Entity inspectionTemplate = null;
                    bool needInspection = false;

                    var activity = jobActivity.GetAttributeValue<EntityReference>("ogre_activity");
                    var asset = _job.GetAttributeValue<EntityReference>("ogre_assetid");

                    var executionIndex = jobActivity.GetAttributeValue<Int32>("ogre_executionindex");

                    if (jobActivity.Contains("ogre_inspectionapplicable"))
                        needInspection = jobActivity.GetAttributeValue<bool>("ogre_inspectionapplicable");

                    if (needInspection)
                    {
                        inspectionTemplate = GetInspectionTemplates(service, activity.Id, asset.Id);
                    }

                    var jobActivityLine = new Entity("ogre_workorderactivity");
                    jobActivityLine["ogre_name"] = _job.GetAttributeValue<string>("ogre_name") + "_" + executionIndex.ToString().PadLeft(3 - executionIndex.ToString().Length, '0');
                    jobActivityLine["ogre_workorderid"] = _job.ToEntityReference();

                    jobActivityLine["ogre_activity"] = activity;

                    jobActivityLine["ogre_resourceassigned"] = _job.GetAttributeValue<EntityReference>("ogre_resourceassigned");
                    jobActivityLine["ogre_jobstart"] = jobActivity.GetAttributeValue<bool>("ogre_jobstart");
                    jobActivityLine["ogre_jobend"] = jobActivity.GetAttributeValue<bool>("ogre_jobend");
                    jobActivityLine["ogre_registrationnumber"] = asset;

                    //string activityName = activity.Name.ToUpper();
                    //if (activityName == "WALKAROUND")
                    //    jobActivityLine["ogre_activitytype"] = new OptionSetValue(1);
                    //else if (activityName == "DRIVING")
                    //    jobActivityLine["ogre_activitytype"] = new OptionSetValue(2);
                    //else if (activityName == "CHECKSHEET")
                    //    jobActivityLine["ogre_activitytype"] = new OptionSetValue(3);
                    //else if (activityName == "RIDE")
                    //    jobActivityLine["ogre_activitytype"] = new OptionSetValue(4);

                    if (inspectionTemplate != null)
                    {
                        var duration = inspectionTemplate.GetAttributeValue<Int32>("ogre_duration");

                        jobActivityLine["ogre_duration"] = duration;

                        totalJobDuration += duration;

                        jobActivityLine["ogre_inspectiontemplate"] = inspectionTemplate.ToEntityReference();
                    }


                    service.Create(jobActivityLine);
                }

                var plannedJobDurationCalculator = new PlannedJobDurationCalculator(service, jobId);
                plannedJobDurationCalculator.UpdateJobDuration();
            }
        }

        static Entity GetInspectionTemplates(IOrganizationService service, Guid _inspectionTypeId, Guid _assetId)
        {

            Entity _entRef = null;

            try
            {
                var query = new QueryExpression("ogre_vehicle");
                query.ColumnSet = new ColumnSet("ogre_name", "ogre_assetcategory", "ogre_vehicletype");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("ogre_vehicleid", ConditionOperator.Equal, _assetId);
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                var vehicleCollection = service.RetrieveMultiple(query);

                if (vehicleCollection != null && vehicleCollection.Entities != null && vehicleCollection.Entities.Count > 0)
                {
                    var vehicle = vehicleCollection.Entities[0];

                    var typeRef = vehicle.GetAttributeValue<EntityReference>("ogre_vehicletype");
                    if (typeRef == null) throw new NullReferenceException("ogre_vehicletype field on ogre_vehicle table is null");

                    var categoryRef = vehicle.GetAttributeValue<EntityReference>("ogre_assetcategory");
                    if (categoryRef == null) throw new NullReferenceException("ogre_assetcategory field on ogre_vehicle table is null");

                    string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                            <entity name='ogre_operationalchecksetup'>
                                                <attribute name='ogre_operationalchecksetupid' />
                                                <attribute name='ogre_duration' />
                                                <attribute name='ogre_vehiclenumber' />
                                                <attribute name='ogre_assetcategory' />
                                                <attribute name='ogre_vehicletype' />
                                                <filter type='and'>
                                                    <filter type='or'>
                                                        <condition attribute='ogre_vehiclenumber' operator='eq' value='{" + _assetId + @"}' />
                                                        <condition attribute='ogre_assetcategory' operator='eq' value='{" + categoryRef.Id + @"}' />
                                                        <condition attribute='ogre_vehicletype' operator='eq' value='{" + typeRef.Id + @"}' />
                                                    </filter>
                                                    <condition attribute='statecode' operator='eq' value='0' />
                                                    <condition attribute='ogre_activity' operator='eq' value='{" + _inspectionTypeId + @"}' />
                                                </filter>
                                            </entity>
                                         </fetch>";

                    var collection = service.RetrieveMultiple(new FetchExpression(_fetchXML));

                    switch (collection.Entities.Count)
                    {
                        case 0:
                            break;
                        case 1:
                            _entRef = collection.Entities[0];
                            break;
                        default:
                            {
                                Entity template = null;

                                foreach (var item in collection.Entities)
                                {
                                    var asset = item.GetAttributeValue<EntityReference>("ogre_vehiclenumber");
                                    var assetId = asset == null ? string.Empty : asset.Id.ToString();

                                    var ac = item.GetAttributeValue<EntityReference>("ogre_assetcategory");
                                    var acId = ac == null ? string.Empty : ac.Id.ToString();

                                    var vt = item.GetAttributeValue<EntityReference>("ogre_vehicletype");
                                    var vtId = vt == null ? string.Empty : vt.Id.ToString();

                                }

                                // Asset
                                template = collection.Entities.
                                                FirstOrDefault
                                                (
                                                    t =>
                                                        t.GetAttributeValue<EntityReference>("ogre_vehiclenumber") != null &&
                                                        t.GetAttributeValue<EntityReference>("ogre_vehiclenumber").Id.Equals(_assetId)
                                                );

                                // Category
                                if (template == null)
                                {
                                    template = collection.Entities.
                                                FirstOrDefault
                                                (
                                                    t =>
                                                        t.GetAttributeValue<EntityReference>("ogre_assetcategory") != null &&
                                                        t.GetAttributeValue<EntityReference>("ogre_assetcategory").Id.Equals(categoryRef.Id)
                                                );
                                }

                                // Type
                                if (template == null)
                                {
                                    template = collection.Entities.
                                                FirstOrDefault
                                                (
                                                    t =>
                                                        t.GetAttributeValue<EntityReference>("ogre_vehicletype") != null &&
                                                        t.GetAttributeValue<EntityReference>("ogre_vehicletype").Id.Equals(typeRef.Id)
                                                );
                                }

                                if (template == null)
                                {
                                    //tracingService.Trace($"Inspection Template not found for asset type {typeRef.Id}");
                                }

                                _entRef = template;
                            }

                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error: " + ex.ToString());
            }

            //tracingService.Trace($"End of GetInspectionTemplates");

            return _entRef;
        }
    }
    public class AirtelIQRequest1
    {
        public string customerId { get; set; }
        public List<string> destinationAddress { get; set; }
        public string message { get; set; }
        public string sourceAddress { get; set; }
        public string messageType { get; set; }
        public string dltTemplateId { get; set; }
        public MetaData metaData { get; set; }
        public string entityId { get; set; }
    }
    public class MetaData1
    {
        public string Key { get; set; }
        public string subAccountId { get; set; }
        public object createdBy { get; set; }
        public RbacSubAccount rbacSubAccount { get; set; }
        public string mdrCategory { get; set; }
    }
    public class RbacSubAccount1
    {
        public string accountId { get; set; }
        public string status { get; set; }
        public Services services { get; set; }
    }
    public class Services1
    {
        public ServiceSMS SMS { get; set; }
    }

    public class ServiceSMS
    {
        public bool creditFlag { get; set; }
        public string serviceStatus { get; set; }
        public int creditsAllotted { get; set; }
    }
    public class AirtelIQResponse
    {
        public string customerId { get; set; }
        public string messageRequestId { get; set; }
        public List<object> incorrectNum { get; set; }
        public string sourceAddress { get; set; }
        public string message { get; set; }
        public string messageType { get; set; }
        public string dltTemplateId { get; set; }
        public string entityId { get; set; }
        public List<string> destinationAddress { get; set; }
        public MetaData metaData { get; set; }
    }
    public class SendSMSViaAirtelIQResponse
    {
        public string ifOkay { get; set; }
        public string responseFromServer { get; set; }
    }
}
 
