using Microsoft.Xrm.Sdk;
using System;
using System.Net;
using System.Text;
using Microsoft.Xrm.Sdk.Query;
using System.IO;
using Microsoft.Xrm.Tooling.Connector;
using RestSharp;
using System.Threading;
using System.Collections.Generic;
using Newtonsoft.Json;
using Microsoft.Crm.Sdk.Messages;
using System.Security.Cryptography;

namespace SMSScheduler
{
    public class Program
    {
        private static Random generator = new Random();
        private static string encryptKey = "AEs0PeR @T!0N";
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";
        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        #endregion
        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {
                //SendSMS();
                //DeleteSMS();
                SendSMSViaAirtel();
                //Entity _entSMS = new Entity("hil_smsconfiguration", new Guid("1424b9de-2af9-ed11-8f6e-6045bdac5098"));
                
            }
        }

        static void CheckIncomingSMS() {
            Entity InSMS = new Entity();
            InSMS["hil_message"] = "201304 Fan is not working";
            InSMS["hil_mobilenumber"] = "8285906486";
            InSMS["hil_direction"] = new OptionSetValue(1);
            InSMS["hil_tomobile"] = "9212110303";
            InSMS["hil_jobtype"] = new OptionSetValue(1);
            _service.Create(InSMS);
        }
        static void SendSMSViaAirtel()
        {
            EntityCollection entcoll;
            string _fetchXML;
            string _mobileNumber = string.Empty;
            string _createdOn = string.Empty;
            string _smsBody = string.Empty;
            string _templateId = string.Empty;
            string _requestType = string.Empty;
            string _kkgOTPOriginal = string.Empty;
            string _kkgOTP = string.Empty;
            while (true)
            {
                //<condition attribute='createdon' operator='last-x-hours' value='1' />

                _fetchXML = @"<fetch top='1000'>
                  <entity name='hil_smsconfiguration'>
                    <attribute name='hil_message' />
                    <attribute name='hil_mobilenumber' />
                    <attribute name='hil_requesttype' />
                    <attribute name='hil_smstemplate' />
                    <attribute name='regardingobjectid' />
                    <attribute name='statuscode' />
                    <attribute name='hil_responsefromserver' />
                    <attribute name='createdon' />
                    <order attribute='createdon' descending='true' />
                    <filter type='and'>
                        <condition attribute='createdon' operator='gt' value='05/24/2023 00:00:00' />
                        <condition attribute='hil_direction' operator='eq' value='2' />
                        <condition attribute='regardingobjectid' operator='not-null' />
                        <condition attribute='hil_mobilenumber' operator='not-null' />
                        <condition attribute='hil_smstemplate' operator='not-null' />
                        <condition attribute='statuscode' operator='not-in'>
                            <value>2</value>
                            <value>910590000</value>
                        </condition>
                        <condition attribute='hil_smstemplate' operator='in'>
                            <value uiname='V1_D365_SendSatisfactionKKGCodetoCustomer' uitype='hil_smstemplates'>{0436EDEF-6862-EB11-A812-0022486E907F}</value>
                            <value uiname='V1_D365_ResendKKGCodetoCustomer' uitype='hil_smstemplates'>{D2955151-7262-EB11-A812-0022486E907F}</value>
                        </condition>
                    </filter>
                    <link-entity name='hil_smstemplates' from='hil_smstemplatesid' to='hil_smstemplate' visible='false' link-type='outer' alias='st'>
                      <attribute name='hil_templateid' />
                    </link-entity>
                  </entity>
                </fetch>";

                string _KKGCode;
                Entity _entSMS;
                entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                int i = 1, reccount = entcoll.Entities.Count;
                foreach (Entity entObj in entcoll.Entities)
                {
                    _entSMS = new Entity("hil_smsconfiguration", entObj.Id);
                    try
                    {
                        if (entObj.Attributes.Contains("createdon"))
                        {
                            _createdOn = entObj.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString();
                        }
                        if (entObj.Attributes.Contains("hil_mobilenumber"))
                        {
                            _mobileNumber = entObj.GetAttributeValue<string>("hil_mobilenumber");
                        }
                        if (entObj.Attributes.Contains("hil_message"))
                        {
                            _smsBody = entObj.GetAttributeValue<string>("hil_message");
                        }
                        if (entObj.Attributes.Contains("hil_smstemplate"))
                        {
                            _templateId = entObj.GetAttributeValue<AliasedValue>("st.hil_templateid").Value.ToString();
                            EntityReference _entSMSTemplate = entObj.GetAttributeValue<EntityReference>("hil_smstemplate");
                        }

                        if (entObj.Attributes.Contains("hil_requesttype"))
                        {
                            if (_requestType == "KKGCode" || _smsBody.IndexOf("<KKGCode>") > 0)
                            {
                                _KKGCode = GenerateKKGCodeHash(entObj.GetAttributeValue<EntityReference>("regardingobjectid").Name, _service);
                                _smsBody = _smsBody.Replace("<KKGCode>", _KKGCode);
                            }
                        }
                        _smsBody = _smsBody.Replace("#", "%23").Replace("&", "%26").Replace("+", "%2B");



                        AirtelIQRequest airtelIQRequest = new AirtelIQRequest()
                        {
                            customerId = "Havells",
                            destinationAddress = new List<string>() { "8285906486" },
                            message = _smsBody,
                            sourceAddress = "HAVELL",
                            messageType = "SERVICE_IMPLICIT",
                            dltTemplateId = _templateId,
                            metaData = new MetaData { Key = "Value" }
                        };

                        string BaseUrl = "https://openapi.airtel.in/gateway/airtel-iq-sms-utility/sendSms";
                        airtelIQRequest.entityId = "110100001483";
                        var json = JsonConvert.SerializeObject(airtelIQRequest);

                        try
                        {
                            WebRequest webRequest = WebRequest.Create(BaseUrl);
                            webRequest.Method = "POST";
                            byte[] data = Encoding.ASCII.GetBytes(json);
                            webRequest.ContentType = "application/json";
                            webRequest.ContentLength = data.Length;
                            webRequest.Headers.Add("Authorization", "Basic UG9saWN5QmF6YWFyOkZtQThGSFYzQEJ3RHUmNms=");
                            Stream requestStream = webRequest.GetRequestStream();
                            requestStream.Write(data, 0, data.Length);
                            WebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();
                            string IfOkay = ((HttpWebResponse)webResponse).StatusCode.ToString();
                            Stream responseStream = webResponse.GetResponseStream();
                            StreamReader myStreamReader = new StreamReader(responseStream, Encoding.Default);
                            string responseFromServer = myStreamReader.ReadToEnd();

                            if (IfOkay == "OK")
                            {
                                _entSMS["statuscode"] = new OptionSetValue(910590000); //Sent
                                _entSMS["hil_responsefromserver"] = "SUCCESS";
                            }
                            else
                            {
                                _entSMS["statuscode"] = new OptionSetValue(910590001); //Not Sent
                                _entSMS["hil_responsefromserver"] = responseFromServer;
                            }
                            myStreamReader.Close();
                            responseStream.Close();
                            webResponse.Close();
                        }
                        catch (WebException we)
                        {
                            _entSMS["statuscode"] = new OptionSetValue(910590001);
                            _entSMS["hil_responsefromserver"] = "GATEWAY BUSY API Url " + BaseUrl + " PayLoad: " + json;
                        }
                        Console.WriteLine("SMS Sent:" + i.ToString() + "/" + reccount.ToString() + "/" + _createdOn);
                        i++;
                    }
                    catch (Exception ex)
                    {
                        _entSMS["statuscode"] = new OptionSetValue(910590001); //Not Sent
                        _entSMS["hil_responsefromserver"] = ex.Message.ToString();
                    }
                    _service.Update(_entSMS);
                }
            }
        }
        public static string GenerateKKGCodeHash(string _jobId, IOrganizationService service)
        {
            string _KKGCode = null;
            QueryExpression query = new QueryExpression("hil_jobsauth");
            query.ColumnSet = new ColumnSet(false);
            query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, _jobId));
            query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1));
            EntityCollection entColl = service.RetrieveMultiple(query);

            foreach (Entity ent in entColl.Entities)
            {
                SetStateRequest deactivateRequest = new SetStateRequest
                {
                    EntityMoniker = ent.ToEntityReference(),
                    State = new OptionSetValue(1),
                    Status = new OptionSetValue(2)
                };
                service.Execute(deactivateRequest);
            }
            string Checksum = getChecksum(_jobId).ToString();
            _KKGCode = GenerateRandomNo();
            Entity entJobAuth = new Entity("hil_jobsauth");
            string _salt = getSalt(Checksum);
            entJobAuth["hil_checksum"] = Checksum;
            entJobAuth["hil_salt"] = _salt;
            entJobAuth["hil_hash"] = getHash(_salt + _KKGCode);
            entJobAuth["hil_name"] = _jobId;
            service.Create(entJobAuth);
            return _KKGCode;
        }
        private static string GenerateRandomNo()
        {
            return generator.Next(0, 999999).ToString("D6");
        }
        private static int getChecksum(string _jobId)
        {
            int _checkSum = 0;
            for (int i = 0; i < 3; i++)
            {
                int _randomNum = generator.Next(1, _jobId.Length);
                _checkSum += Convert.ToInt32(_jobId.Substring(_randomNum - 1, 1));
            }
            return _checkSum;
        }
        private static string getSalt(string Checksum)
        {
            byte[] bytes = new byte[256 / 8]; // 32bit
            using (var keyGenerator = RandomNumberGenerator.Create())
            {
                keyGenerator.GetBytes(bytes);
                return BitConverter.ToString(bytes).Replace("-", "").ToLower() + Checksum;
            }
        }
        private static string getHash(string text)
        {
            using (var sha512 = SHA512.Create())
            {
                var hashedBytes = sha512.ComputeHash(Encoding.UTF8.GetBytes(text));
                return BitConverter.ToString(hashedBytes).Replace("-", "").ToLower();
            }
        }
        static void SendSMS()
        {
            QueryExpression queryExp;
            EntityCollection entcoll;
            string _fetchXML;
            string _mobileNumber = string.Empty;
            string _createdOn = string.Empty;
            string _smsBody = string.Empty;
            string _templateId = string.Empty;
            string _requestType = string.Empty;
            string _kkgOTPOriginal = string.Empty;
            string _kkgOTP = string.Empty;
            while (true)
            {
                _fetchXML = @"<fetch top='1000'>
                  <entity name='hil_smsconfiguration'>
                    <attribute name='hil_message' />
                    <attribute name='hil_mobilenumber' />
                    <attribute name='hil_requesttype' />
                    <attribute name='hil_smstemplate' />
                    <attribute name='regardingobjectid' />
                    <attribute name='statuscode' />
                    <attribute name='hil_responsefromserver' />
                    <attribute name='createdon' />
                    <order attribute='createdon' descending='true' />
                    <filter type='and'>
                        <condition attribute='createdon' operator='gt' value='04/25/2021 16:00:00' />
                        <condition attribute='hil_direction' operator='eq' value='2' />
                        <condition attribute='hil_mobilenumber' operator='not-null' />
                        <condition attribute='hil_smstemplate' operator='not-null' />
                        <condition attribute='statuscode' operator='eq' value='910590001' />
                        <condition attribute='createdon' operator='last-x-hours' value='1' />
                    </filter>
                    <link-entity name='hil_smstemplates' from='hil_smstemplatesid' to='hil_smstemplate' visible='false' link-type='outer' alias='st'>
                      <attribute name='hil_templateid' />
                    </link-entity>
                  </entity>
                </fetch>";

                entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                int i = 1, reccount = entcoll.Entities.Count;
                Entity entUpdate = null;
                foreach (Entity entObj in entcoll.Entities)
                {
                    if (entObj.Attributes.Contains("createdon"))
                    {
                        _createdOn = entObj.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString();
                    }
                    if (entObj.Attributes.Contains("hil_mobilenumber"))
                    {
                        _mobileNumber = entObj.GetAttributeValue<string>("hil_mobilenumber");
                    }
                    if (entObj.Attributes.Contains("hil_message"))
                    {
                        _smsBody = entObj.GetAttributeValue<string>("hil_message");
                    }
                    if (entObj.Attributes.Contains("hil_smstemplate"))
                    {
                        _templateId = entObj.GetAttributeValue<AliasedValue>("st.hil_templateid").Value.ToString();
                    }
                    if (entObj.Attributes.Contains("hil_requesttype"))
                    {
                        _requestType = entObj.GetAttributeValue<string>("hil_requesttype");
                        if (_requestType == "KKGCode" || _smsBody.IndexOf("<KKGCode>") > 0)
                        {
                            _kkgOTPOriginal = GetKKGOTP(_service, entObj.GetAttributeValue<EntityReference>("regardingobjectid").Id);
                            _kkgOTP = Base64Decode(_kkgOTPOriginal);

                            if (_kkgOTP.Trim().Length == 4)
                            {
                                _smsBody = _smsBody.Replace("<KKGCode>", _kkgOTP);
                            }
                            else
                            {
                                _smsBody = _smsBody.Replace("<KKGCode>", _kkgOTPOriginal);
                            }
                        }
                    }

                    entUpdate = new Entity("hil_smsconfiguration", entObj.Id);
                    string _api = string.Empty;

                    if (_templateId == "1107161191448698079" || _templateId == "1107161191438154934") //KKG Code Send and Resend Templates
                    {
                        _api = "https://digimate.airtel.in:15443/BULK_API/SendMessage?loginID=havells_htu2&password=havells@123&mobile=" + _mobileNumber + "&text=" + _smsBody + "&senderid=HAVELL&DLT_TM_ID=1001096933494158&DLT_CT_ID=" + _templateId + "&DLT_PE_ID=110100001483&route_id=DLT_SERVICE_IMPLICT&Unicode=0&camp_name=havells_u";
                    }
                    else {
                        _api = "https://japi.instaalerts.zone/failsafe/HttpLink?aid=640990&pin=w~7Xg)9V&mnumber=" + _mobileNumber + "&signature=HAVELL&message=" + _smsBody + "&dlt_entity_id=110100001483&dlt_template_id=" + _templateId;
                    }
                    WebRequest request = WebRequest.Create(_api);

                    if (_templateId != "1107161191448698079" && _templateId != "1107161191438154934") //KKG Code Send and Resend Templates
                    {
                        string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes("HAVELLSSERVICE" + ":" + "w~7Xg)9V"));
                        request.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
                    }

                    request.Method = "GET";
                    WebResponse response = null;
                    string IfOkay = string.Empty;
                    string responseFromServer = string.Empty;
                    try
                    {
                        response = request.GetResponse();
                        Stream dataStream = Stream.Null;
                        //Console.WriteLine(((HttpWebResponse)response).StatusDescription);
                        IfOkay = ((HttpWebResponse)response).StatusDescription;
                        dataStream = response.GetResponseStream();
                        StreamReader reader = new StreamReader(dataStream);
                        responseFromServer = reader.ReadToEnd();
                    }
                    catch (WebException ex)
                    {
                        entUpdate["statuscode"] = new OptionSetValue(910590001);
                        entUpdate["hil_responsefromserver"] = "GATEWAY BUSY APIPayLoad: " + _api;
                    }
                    finally
                    {
                        entUpdate["hil_responsefromserver"] = responseFromServer;
                        if (IfOkay == "OK")
                            entUpdate["statuscode"] = new OptionSetValue(910590000); //Sent
                        else
                            entUpdate["Statuscode"] = new OptionSetValue(910590001); //Not Sent
                    }
                    _service.Update(entUpdate);
                    Console.WriteLine("SMS Sent:" + i.ToString() + "/" + reccount.ToString() + "/" + _createdOn);
                    i++;
                }
            }
        }
        static string ReplaceGSMCharFromSMS(string _msg) {
            try
            {
                string[,] arrayGSM = new string[10, 2]
                {
                    { "!", "%21" },
                    { "#", "%23" },
                    { "$", "%24" },
                    { "%", "%25" },
                    { "%", "%25" },
                    { "&", "%26" },
                    { "‘", "%27" },
                    { "/", "%2F" },
                    { "~", "%7E*" },
                    { "@", "%40" }
                };
                for (int i = 0; i <= arrayGSM.Length; i++)
                {
                    _msg.Replace(arrayGSM[i, 0], arrayGSM[i, 1]);
                }
            }
            catch { }
            return _msg;
        }
        static void DeleteSMS()
        {
            QueryExpression queryExp;
            EntityCollection entcoll;
            string _fetchXML;
            string _createdOn = string.Empty;
            while (true)
            {
                //<condition attribute='createdon' operator='gt' value='03/20/2021 22:00:00' />
                //<condition attribute='createdon' operator='last-x-hours' value='1' />

                _fetchXML = @"<fetch top='1000'>
                  <entity name='hil_smsconfiguration'>
                    <attribute name='createdon' />
                    <order attribute='createdon' descending='true' />
                    <filter type='and'>
                      <condition attribute='createdon' operator='lt' value='04/05/2022 12:10:00' />
                      <condition attribute='hil_direction' operator='eq' value='2' />
                      <condition attribute='hil_mobilenumber' operator='not-null' />
                      <condition attribute='hil_smstemplate' operator='not-null' />
                      <condition attribute='statuscode' operator='eq' value='910590001' />
                    </filter>
                  </entity>
                </fetch>";

                entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                int i = 1, reccount = entcoll.Entities.Count;
                foreach (Entity entObj in entcoll.Entities)
                {
                    if (entObj.Attributes.Contains("createdon"))
                    {
                        _createdOn = entObj.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString();
                    }
                    _service.Delete("hil_smsconfiguration", entObj.Id);
                    Console.WriteLine("SMS Sent:" + i.ToString() + "/" + reccount.ToString() + "/" + _createdOn);
                    i++;
                }
            }
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

        public static string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }
        public static string GetKKGOTP(IOrganizationService service, Guid workOrderId)
        {
            string _kkgOPT = string.Empty;
            try
            {
                QueryExpression Query = new QueryExpression("msdyn_workorder");
                Query.ColumnSet = new ColumnSet("hil_kkgotp");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression("msdyn_workorderid", ConditionOperator.Equal, workOrderId));
                EntityCollection enCol = service.RetrieveMultiple(Query);
                if (enCol.Entities.Count > 0)
                {
                    _kkgOPT = enCol.Entities[0].GetAttributeValue<string>("hil_kkgotp");
                }
            }
            catch (Exception ex)
            {
                throw new Exception("SMSScheduler.GetKKGOTP" + ex.Message);
            }
            return _kkgOPT;
        }
        #endregion
    }

    //public class MetaData
    //{
    //    public string Key { get; set; }
    //}

    public class Root
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

    public class AirtelIQRequest
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

    public class MetaData
    {
        public string Key { get; set; }
        public string subAccountId { get; set; }
        public object createdBy { get; set; }
        public RbacSubAccount rbacSubAccount { get; set; }
        public string mdrCategory { get; set; }
    }

    public class RbacSubAccount
    {
        public string accountId { get; set; }
        public string status { get; set; }
        public Services services { get; set; }
    }
    public class Services
    {
        public AirtelIQSMS SMS { get; set; }
    }

    public class AirtelIQSMS
    {
        public bool creditFlag { get; set; }
        public string serviceStatus { get; set; }
        public int creditsAllotted { get; set; }
    }
}
