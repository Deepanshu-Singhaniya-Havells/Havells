using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Runtime.Serialization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;
using System.Threading.Tasks;
using System.Web.Services.Description;

namespace Havells.CRM.DataEncription
{
    internal class Program
    {
        public static string EncriptionKey = "12s2s121sasfdasdf45346fwrt3w56fw";
        static void Main(string[] args)
        {

            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
            var CrmPrdURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();
            string finalPrdString = string.Format(connStr, CrmPrdURL);
            IOrganizationService service = createConnection(finalPrdString);

            AuthenticateConsumer authenticateConsumerRequest = new AuthenticateConsumer();
            //authenticateConsumerRequest.SourceType = "101";
            authenticateConsumerRequest.SourceType= "5";
            authenticateConsumerRequest.LoginUserId = "9738058501";
            var asn = AuthenticateConsumerAMC(authenticateConsumerRequest, service);

            ValidateSessionRequest validateSessionRequest = new ValidateSessionRequest();
            validateSessionRequest.SessionId = asn.SessionId;
            validateSessionRequest.KeepSessionLive = true;
            var dddd = ValidateSessionDetails(validateSessionRequest);

            //var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
            //var CrmPrdURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();
            //string finalPrdString = string.Format(connStr, CrmPrdURL);
            //IOrganizationService service = createConnection(finalPrdString);

            //Entity entity = service.Retrieve("msdyn_workorder", new Guid("3af317d9-e5e6-ec11-bb3c-6045bdad276f"), new ColumnSet(true));
            //jobPreValidate(service, entity);
            ////var key = "12s2s121sasfdasdf45346fwrt3w56fw";
            //EncriptionFunction(service);
            ////Console.WriteLine("Please enter a secret key for the symmetric algorithm.");  
            ////var key = Console.ReadLine();  

            //Console.WriteLine("Please enter a string for encryption");
            //var str = Console.ReadLine();
            //var encryptedString = EncryptString(key, str);
            //Console.WriteLine($"encrypted string = {encryptedString}");

            //var decryptedString = DecryptString(key, encryptedString);
            //Console.WriteLine($"decrypted string = {decryptedString}");

            Console.ReadKey();
        }
        static void mainFunction(IOrganizationService service)
        {
            var entityId = "c49f9507-6774-ed11-81ac-6045bdad9a22";
            if (entityId == null)
            {
                Console.WriteLine("Status Failed !!");
                //Console.WriteLine["Message"] = "Invalid SMS GUID";
            }
            else
            {
                Entity _oldSms = service.Retrieve("hil_smsconfiguration", new Guid(entityId), new ColumnSet(true));
                EntityReference _smsTemplateID = _oldSms.GetAttributeValue<EntityReference>("hil_smstemplate");
                Entity _smsTemplate = service.Retrieve(_smsTemplateID.LogicalName, _smsTemplateID.Id, new ColumnSet("hil_encryptsms"));
                Entity _newSMS = new Entity(_oldSms.LogicalName);
                if (_oldSms.GetAttributeValue<bool>("hil_encrypted"))
                {
                    _newSMS["hil_message"] = DecryptString(EncriptionKey, (string)_oldSms["hil_message"]);
                }
                else
                    _newSMS["hil_message"] = _oldSms["hil_message"];
                var attributes = _oldSms.Attributes.Keys;

                foreach (string name in attributes)
                {
                    // if (name != "adx_partnervisible")
                    if (name != "modifiedby" && name != "createdby" && name != "ownerid" && name != "owninguser" && name != "statecode"
                        && name != "statuscode" && name != "hil_responsefromserver" && name != "hil_message" && name != "activityid"
                        && name != "createdon" && name != "modifiedon" && name != "owningbusinessunit" && name != "timezoneruleversionnumber"
                        && name != "activitytypecode" && name != "instancetypecode" && name != "isworkflowcreated" && name != "processid")
                    {
                        _newSMS[name] = _oldSms[name];
                    }
                }
                Guid _newSMSID = service.Create(_newSMS);
            }
        }
        static void EncriptionFunction(IOrganizationService service)
        {
            var entityId = "c49f9507-6774-ed11-81ac-6045bdad9a22";
            if (entityId == null)
            {
                Console.WriteLine("Status Failed !!");
                //Console.WriteLine["Message"] = "Invalid SMS GUID";
            }
            else
            {
                Entity _oldSms = service.Retrieve("hil_smsconfiguration", new Guid(entityId), new ColumnSet(true));
                EntityReference _smsTemplateID = _oldSms.GetAttributeValue<EntityReference>("hil_smstemplate");
                Entity _smsTemplate = service.Retrieve(_smsTemplateID.LogicalName, _smsTemplateID.Id, new ColumnSet("hil_encryptsms"));
                if (_smsTemplate.GetAttributeValue<bool>("hil_encryptsms"))
                {
                    Entity _newSMS = new Entity(_oldSms.LogicalName, _oldSms.Id);
                    _newSMS["hil_message"] = EncryptString(EncriptionKey, (string)_oldSms["hil_message"]);
                    service.Update(_newSMS);
                }

            }
        }
        public static string EncryptString(string key, string plainText)
        {
            byte[] iv = new byte[16];
            byte[] array;

            using (Aes aes = Aes.Create())
            {
                aes.Key = Encoding.UTF8.GetBytes(key);
                aes.IV = iv;

                ICryptoTransform encryptor = aes.CreateEncryptor(aes.Key, aes.IV);

                using (MemoryStream memoryStream = new MemoryStream())
                {
                    using (CryptoStream cryptoStream = new CryptoStream((Stream)memoryStream, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter streamWriter = new StreamWriter((Stream)cryptoStream))
                        {
                            streamWriter.Write(plainText);
                        }

                        array = memoryStream.ToArray();
                    }
                }
            }

            return Convert.ToBase64String(array);
        }
        public static string DecryptString(string key, string cipherText)
        {
            byte[] iv = new byte[16];
            byte[] buffer = Convert.FromBase64String(cipherText);

            using (Aes aes = Aes.Create())
            {
                aes.Key = Encoding.UTF8.GetBytes(key);
                aes.IV = iv;
                ICryptoTransform decryptor = aes.CreateDecryptor(aes.Key, aes.IV);

                using (MemoryStream memoryStream = new MemoryStream(buffer))
                {
                    using (CryptoStream cryptoStream = new CryptoStream((Stream)memoryStream, decryptor, CryptoStreamMode.Read))
                    {
                        using (StreamReader streamReader = new StreamReader((Stream)cryptoStream))
                        {
                            return streamReader.ReadToEnd();
                        }
                    }
                }
            }
        }
        public static IOrganizationService createConnection(string connectionString)
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
        static void jobPreValidate(IOrganizationService _service, Entity _jobEntity)
        {
            Guid _productDivisionId = _jobEntity.GetAttributeValue<EntityReference>("hil_productcategory").Id;
            string _mobileNumber = _jobEntity.GetAttributeValue<string>("hil_mobilenumber");
            QueryExpression qryExp = new QueryExpression("hil_jobsquantitymatrix");
            qryExp.ColumnSet = new ColumnSet("hil_quantity", "hil_frequency");
            qryExp.Criteria.AddCondition(new ConditionExpression("hil_divisionname", ConditionOperator.Equal, _productDivisionId));
            qryExp.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            EntityCollection entCol = _service.RetrieveMultiple(qryExp);
            if (entCol.Entities.Count > 0)
            {
                Entity _entObj = entCol.Entities[0];
                if (_entObj != null)
                {
                    int _frequency = _entObj.GetAttributeValue<int>("hil_frequency");
                    int _quantity = _entObj.GetAttributeValue<int>("hil_quantity");
                    qryExp = new QueryExpression("msdyn_workorder");
                    qryExp.ColumnSet = new ColumnSet(false);
                    qryExp.Criteria.AddCondition(new ConditionExpression("hil_productcategory", ConditionOperator.Equal, _productDivisionId));
                    qryExp.Criteria.AddCondition(new ConditionExpression("hil_mobilenumber", ConditionOperator.Equal, _mobileNumber));
                    qryExp.Criteria.AddCondition(new ConditionExpression("createdon", ConditionOperator.LastXHours, _frequency));
                    EntityCollection _entColObj = _service.RetrieveMultiple(qryExp);
                    if (_entColObj.Entities.Count >= _quantity)
                    {
                        Console.WriteLine("Data Duplicacy Rule !!! Concurrent Jobs for Mobile Number# " + _mobileNumber);
                    }
                }
                else
                {
                    Console.WriteLine("Data Duplicacy Rule does not exist for Division " + _productDivisionId + " has been traced.");
                }
            }
            else
            {
                Console.WriteLine("Data Duplicacy Rule does not exist for Division " + _productDivisionId + " has been traced.");
            }
        }

        public static AuthenticateConsumer AuthenticateConsumerAMC(AuthenticateConsumer requestParam, IOrganizationService _service)
        {
            //IOrganizationService _service = ConnectToCRM.GetOrgService();
            AuthenticateConsumer _retValue = new AuthenticateConsumer()
            {
                LoginUserId = requestParam.LoginUserId,
                SourceType = requestParam.SourceType
            };
            Entity _userSession = null;
            try
            {
                if (requestParam.SourceType == string.Empty && requestParam.SourceType == null)// != "5")
                {
                    _retValue.StatusCode = "400";
                    _retValue.StatusDescription = "Access Denied!!! Source Type is mandatory";
                    return _retValue;
                }
                if (requestParam.LoginUserId == string.Empty || requestParam.LoginUserId == null)
                {
                    _retValue.StatusCode = "400";
                    _retValue.StatusDescription = "Access Denied!!! Source Code or Mobile Number is mandatory";
                    return _retValue;
                }
                if (_service != null)
                {
                    int expTime = 0;
                    string _portalSURL = string.Empty;
                    if (!CheckSourceType(_service, requestParam.SourceType, out expTime, out _portalSURL))
                    {
                        _retValue.StatusCode = "400";
                        _retValue.StatusDescription = "Access Denied!!! API is not extended to Source Type: " + requestParam.SourceType;
                        return _retValue;
                    }
                    else
                    {
                        QueryExpression queryExp = new QueryExpression("hil_consumerloginsession");
                        queryExp.ColumnSet = new ColumnSet("hil_expiredon", "hil_consumerloginsessionid");
                        ConditionExpression condExp = new ConditionExpression("hil_name", ConditionOperator.Equal, requestParam.LoginUserId);
                        queryExp.Criteria.AddCondition(condExp);
                        condExp = new ConditionExpression("statecode", ConditionOperator.Equal, 0);
                        queryExp.Criteria.AddCondition(condExp);
                        queryExp.AddOrder("hil_expiredon", OrderType.Descending);

                        EntityCollection entCol = _service.RetrieveMultiple(queryExp);
                        if (entCol.Entities.Count > 0) //No Active session found
                        {
                            foreach (Entity entity in entCol.Entities)
                            {
                                SetStateRequest setStateRequest = new SetStateRequest()
                                {
                                    EntityMoniker = new EntityReference
                                    {
                                        Id = entCol.Entities[0].Id,
                                        LogicalName = "hil_consumerloginsession",
                                    },
                                    State = new OptionSetValue(1), //Inactive
                                    Status = new OptionSetValue(2) //Inactive
                                };
                                _service.Execute(setStateRequest);
                            }
                        }
                        _userSession = new Entity("hil_consumerloginsession");
                        _userSession["hil_name"] = requestParam.LoginUserId;
                        _userSession["hil_origin"] = requestParam.SourceType;
                        expTime = expTime + 330;
                        _userSession["hil_expiredon"] = DateTime.Now.AddMinutes(expTime);
                        _retValue.SessionId = _service.Create(_userSession).ToString();

                        string _postfixURL = EncryptAES256("SessionId=" + _retValue.SessionId + "&MobileNumber=" + requestParam.LoginUserId + "&SourceOrigin=" + requestParam.SourceType);
                        _retValue.RedirectUrl = System.Web.HttpUtility.UrlEncoder( _portalSURL + _postfixURL, UTF8Encoding.UTF32);
                        _retValue.StatusDescription = "Sucess";
                        _retValue.StatusCode = "200";
                    }
                }
                else
                {
                    _retValue.StatusCode = "503";
                    _retValue.StatusDescription = "D365 Service is unavailable.";
                }
            }
            catch (Exception ex)
            {
                _retValue.StatusCode = "500";
                _retValue.StatusDescription = "D365 Internal Server Error : " + ex.Message;
            }
            return _retValue;
        }
        private static string EncryptAES256(string plainText)
        {
            string Key = "DklsdvkfsDlkslsdsdnv234djSDAjkd1";
            byte[] key32 = Encoding.UTF8.GetBytes(Key);
            byte[] IV16 = Encoding.UTF8.GetBytes(Key.Substring(0, 16)); 
            if (plainText == null || plainText.Length <= 0)
            {
                throw new ArgumentNullException("plainText");
            }
            byte[] encrypted;
            using (AesManaged aesAlg = new AesManaged())
            {
                aesAlg.KeySize = 256;
                aesAlg.Padding = PaddingMode.PKCS7;
                aesAlg.IV = IV16;
                aesAlg.Key = key32;
                ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);
                using (MemoryStream msEncrypt = new MemoryStream())
                {
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter swEncrypt = new StreamWriter(csEncrypt))
                        {
                            //Write all data to the stream.
                            swEncrypt.Write(plainText);
                        }
                        encrypted = msEncrypt.ToArray();
                    }
                }
            }
            System.Text.Encoding.UTF8.GetString(encrypted);
            char[] characters = encrypted.Select(b => (char)b).ToArray();
            string ddd =  new string(characters);
            return Convert.ToBase64String(encrypted);
        }

        public static bool CheckSourceType(IOrganizationService service, string sourceOrigin, out int expTime, out string portalURL)
        {
            QueryExpression qryExp = new QueryExpression("hil_integrationsource");
            qryExp.ColumnSet = new ColumnSet("hil_deeplinkingallowed", "hil_sessiontimeout", "hil_amcportalurl");
            qryExp.Criteria.AddCondition(new ConditionExpression("hil_code", ConditionOperator.Equal, sourceOrigin));
            qryExp.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            EntityCollection entCol = service.RetrieveMultiple(qryExp);
            if (entCol.Entities.Count != 1)
            {
                expTime = 0;
                portalURL = string.Empty;
                return false;
            }
            else
            {
                expTime = entCol[0].GetAttributeValue<int>("hil_sessiontimeout");
                portalURL = entCol[0].GetAttributeValue<string>("hil_amcportalurl");
                return entCol[0].GetAttributeValue<bool>("hil_deeplinkingallowed");
            }
        }
        public static ValidateSessionResponse ValidateSessionDetails(ValidateSessionRequest requestParam)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
            var CrmPrdURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();
            string finalPrdString = string.Format(connStr, CrmPrdURL);
            IOrganizationService _service = createConnection(finalPrdString);

            //IOrganizationService _service = ConnectToCRM.GetOrgService();
            ValidateSessionResponse _retValue = new ValidateSessionResponse();
            if (requestParam.SessionId == string.Empty && requestParam.SessionId == null)// != "5")
            {
                _retValue.StatusCode = "400";
                _retValue.StatusDescription = "Access Denied!!! Source Type is mandatory";
                return _retValue;
            }
            QueryExpression queryExp = new QueryExpression("hil_consumerloginsession");
            queryExp.ColumnSet = new ColumnSet("hil_expiredon", "hil_consumerloginsessionid");
            ConditionExpression condExp = new ConditionExpression("hil_consumerloginsessionid", ConditionOperator.Equal, requestParam.SessionId);
            queryExp.Criteria.AddCondition(condExp);
            condExp = new ConditionExpression("statecode", ConditionOperator.Equal, 0);
            queryExp.Criteria.AddCondition(condExp);

            EntityCollection entCol = _service.RetrieveMultiple(queryExp);

            if (entCol.Entities.Count == 1)
            {
                DateTime expDate = entCol[0].GetAttributeValue<DateTime>("hil_expiredon").AddMinutes(330);
                if (expDate > DateTime.Now)
                {
                    _retValue.StatusCode = "200";
                    _retValue.StatusDescription = "Session is Valid";
                    return _retValue;
                }
                else
                {
                    _retValue.StatusCode = "400";
                    _retValue.StatusDescription = "Access Denied!!! Session Expire";
                    return _retValue;
                }
            }
            else
            {
                _retValue.StatusCode = "400";
                _retValue.StatusDescription = "Invalid Session";
                return _retValue;
            }
        }

        public class AuthenticateConsumer
        {
           
            public string LoginUserId { get; set; }
           
            public string SourceType { get; set; }

           
            public string SessionId { get; set; }
           
            public string RedirectUrl { get; set; }
           
            public string StatusCode { get; set; }
           
            public string StatusDescription { get; set; }
        }

        public class ValidateSessionRequest
        {
           
            public string JWTToken { get; set; }
           
            public string MobileNumber { get; set; }
           
            public string SourceType { get; set; }
           
            public string SourceCode { get; set; }
           
            public string SessionId { get; set; }
           
            public bool KeepSessionLive { get; set; }
        }
        public class ValidateSessionResponse
        {
           
            public string JWTToken { get; set; }
           
            public string MobileNumber { get; set; }
           
            public string SourceType { get; set; }
           
            public string SourceCode { get; set; }
           
            public string SessionId { get; set; }
           
            public string StatusCode { get; set; }
           
            public string StatusDescription { get; set; }
        }
        //public class AuthenticateConsumer
        //{
        //   
        //    public string JWTToken { get; set; }
        //   
        //    public string MobileNumber { get; set; }
        //   
        //    public string SourceType { get; set; }
        //   
        //    public string SourceCode { get; set; }
        //   
        //    public string SessionId { get; set; }
        //   
        //    public string StatusCode { get; set; }
        //   
        //    public string StatusDescription { get; set; }
        //}
        public class AuthenticateConsumerRequest
        {
           
            public string MobileNumber { get; set; }
           
            public string SourceType { get; set; }
           
            public string SourceCode { get; set; }
           
            public string TockenType { get; set; }
        }

    }
}