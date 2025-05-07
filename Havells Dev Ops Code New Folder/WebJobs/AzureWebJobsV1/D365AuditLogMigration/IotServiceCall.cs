using System;
using System.Collections.Generic;
using System.Configuration;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Runtime.Serialization;
using Microsoft.Crm.Sdk.Messages;
using System.Text;
using System.Security.Cryptography;
using System.IO;

namespace D365AuditLogMigration
{

    public class IotServiceCall
    {
        public Guid CustomerGuid { get; set; }

        public AuthenticateConsumer AuthenticateConsumerAMC(AuthenticateConsumer requestParam, IOrganizationService _service)
        {
            //IOrganizationService _service = ConnectToCRM.GetOrgServiceProd();
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
                    string _portalURL = string.Empty;
                    if (!CheckSourceType(_service, requestParam.SourceType, out expTime, out _portalURL))
                    {
                        _retValue.StatusCode = "400";
                        _retValue.StatusDescription = "Access Denied!!! API is not extended to Source Type: " + requestParam.SourceType;
                        return _retValue;
                    }
                    else
                    {
                        string Key = ConfigurationManager.AppSettings["EncryptionKeyAES256"];
                        QueryExpression queryExp = new QueryExpression("hil_consumerloginsession");
                        queryExp.ColumnSet = new ColumnSet("hil_expiredon", "hil_consumerloginsessionid");
                        ConditionExpression condExp = new ConditionExpression("hil_name", ConditionOperator.Equal, requestParam.LoginUserId);
                        queryExp.Criteria.AddCondition(condExp);
                        condExp = new ConditionExpression("statecode", ConditionOperator.Equal, 0);
                        queryExp.Criteria.AddCondition(condExp);
                        queryExp.AddOrder("hil_expiredon", OrderType.Descending); EntityCollection entCol = _service.RetrieveMultiple(queryExp);
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

                        _retValue.PortalURL = _portalURL + _postfixURL;
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

        public ValidateSessionResponse ValidateSessionDetails(ValidateSessionRequest requestParam, IOrganizationService _service)
        {
            ValidateSessionResponse _retValue = new ValidateSessionResponse();
            try
            {
                //IOrganizationService _service = ConnectToCRM.GetOrgServiceProd();

                if (requestParam.SessionId == string.Empty && requestParam.SessionId == null)
                {
                    _retValue.StatusCode = "400";
                    _retValue.StatusDescription = "Access Denied!!! Please input Session Id.";
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
                        if (requestParam.KeepSessionLive)
                        {
                            Entity _userSession = new Entity("hil_consumerloginsession", entCol.Entities[0].Id);
                            _userSession["hil_expiredon"] = DateTime.Now.AddMinutes(20);
                            _service.Update(_userSession);
                        }
                        _retValue.StatusCode = "200";
                        _retValue.StatusDescription = "Session Id is Valid";
                    }
                    else
                    {
                        _retValue.StatusCode = "400";
                        _retValue.StatusDescription = "Access Denied!!! Session has been Expired";
                    }
                }
                else
                {
                    _retValue.StatusCode = "400";
                    _retValue.StatusDescription = "Invalid Session Id.";
                }
            }
            catch (Exception ex)
            {
                _retValue.StatusCode = "400";
                _retValue.StatusDescription = "ERROR!!! " + ex.Message;
            }
            return _retValue;
        }

        public ValidateSessionResponse CreateSession(AuthenticateConsumer requestParam, IOrganizationService _service)
        {
            ValidateSessionResponse _retValue = new ValidateSessionResponse();
            try
            {
                //IOrganizationService _service = ConnectToCRM.GetOrgServiceProd();

                QueryExpression queryExp = new QueryExpression("hil_consumerloginsession");
                queryExp.ColumnSet = new ColumnSet("hil_expiredon", "hil_consumerloginsessionid");
                ConditionExpression condExp = new ConditionExpression("hil_name", ConditionOperator.Equal, requestParam.LoginUserId);
                ConditionExpression condExp1 = new ConditionExpression("statecode", ConditionOperator.Equal, 1);

                queryExp.Criteria.AddCondition(condExp);
                queryExp.Criteria.AddCondition(condExp1);
                EntityCollection entCol = _service.RetrieveMultiple(queryExp);
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
                }
                Entity _userSession = new Entity("hil_consumerloginsession");
                _userSession["hil_name"] = requestParam.LoginUserId;
                _userSession["hil_expiredon"] = DateTime.Now.AddMinutes(350);
                _userSession["hil_origin"] = requestParam.SourceType;
                _retValue.SessionId = _service.Create(_userSession).ToString();
                _retValue.StatusCode = "200";
                _retValue.StatusDescription = "OK";
                return _retValue;
            }
            catch (Exception ex)
            {
                _retValue.StatusCode = "400";
                _retValue.StatusDescription = "ERROR!!! " + ex.Message;
            }
            return _retValue;
        }
        public bool CheckSourceType(IOrganizationService service, string sourceOrigin, out int expTime, out string portalURL)
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

        private string EncryptAES256(string plainText)
        {
            string Key = "DklsdvkfsDlkslsdsdnv234djSDAjkd1";
            byte[] key32 = Encoding.UTF8.GetBytes(Key);
            byte[] IV16 = Encoding.UTF8.GetBytes(Key.Substring(0, 16)); if (plainText == null || plainText.Length <= 0)
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
            return Convert.ToBase64String(encrypted);
        }
    }

    public class IoTServiceCallRegistration
    {
        public string SerialNumber { get; set; }
        public string ProductModelNumber { get; set; }
        public Guid NOCGuid { get; set; }
        public string NOCName { get; set; }
        public Guid ProductCategoryGuid { get; set; }

        public Guid ProductSubCategoryGuid { get; set; }


        public string ChiefComplaint { get; set; }

        public Guid AddressGuid { get; set; }

        public Guid AssetGuid { get; set; }

        public string CustomerMobleNo { get; set; }

        public Guid CustomerGuid { get; set; }

        public Guid JobGuid { get; set; }

        public string JobId { get; set; }

        public string ImageBase64String { get; set; }

        public int ImageType { get; set; }

        public int SourceOfJob { get; set; }

        public string PreferredDate { get; set; }

        public int PreferredPartOfDay { get; set; }

        public string StatusCode { get; set; }

        public string StatusDescription { get; set; }

        public string DealerCode { get; set; }

        public string DealerName { get; set; }

        public string CustomerName { get; set; }

        public string AddressLine1 { get; set; }

        public string Pincode { get; set; }

        public string PreferredLanguage { get; set; }
    }

    public class IoTServiceCallResult
    {

        public string JobId { get; set; }


        public Guid JobGuid { get; set; }


        public string CallSubType { get; set; }


        public string JobLoggedon { get; set; }


        public string JobStatus { get; set; }


        public string JobAssignedTo { get; set; }


        public string CustomerAsset { get; set; }


        public string ProductCategory { get; set; }


        public string NatureOfComplaint { get; set; }


        public string JobClosedOn { get; set; }


        public string CustomerName { get; set; }


        public string ServiceAddress { get; set; }


        public string Product { get; set; }


        public string ChiefComplaint { get; set; }

        public string PreferredDate { get; set; }

        public int PreferredPartOfDay { get; set; }

        public string PreferredPartOfDayName { get; set; }

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


    public class IoTNatureofComplaint
    {

        public string SerialNumber { get; set; }


        public Guid ProductSubCategoryId { get; set; }


        public string Name { get; set; }

        public Guid Guid { get; set; }

        public string StatusCode { get; set; }

        public string StatusDescription { get; set; }


        public string Source { get; set; }
    }

    public class CancelJobRequest
    {

        public string JobGuid { get; set; }
    }

    public class CancelJobResponse
    {

        public string StatusCode { get; set; }

        public string StatusDescription { get; set; }
    }

    public class AuthenticateConsumer
    {
        public string LoginUserId { get; set; }
        public string SourceType { get; set; }
        public string SessionId { get; set; }
        public string PortalURL { get; set; }
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
    }
}
