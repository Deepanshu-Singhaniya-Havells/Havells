using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

using System.Threading.Tasks;

namespace Havells.Dataverse.CustomConnector.AMC
{
    public class AuthServiceForAMC : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            try
            {
                var request = new AuthenticateConsumer
                {
                    LoginUserId = Convert.ToString(context.InputParameters["LoginUserId"]),
                    SourceType = Convert.ToString(context.InputParameters["SourceType"])
                    //SessionId = Convert.ToString(context.InputParameters["SessionId"]),
                    //RedirectUrl = Convert.ToString(context.InputParameters["RedirectUrl"]),
                    //StatusCode = Convert.ToString(context.InputParameters["StatusCode"]),
                    //StatusDescription = Convert.ToString(context.InputParameters["StatusDescription"])
                };

                if (!AuthenticateConsumerRequestValidator.Validate(request, out string validationMessage))
                {
                    var paramResponse = JsonSerializer.Serialize(new { StatusCode = "400", StatusDescription = validationMessage });

                    context.OutputParameters["data"] = paramResponse;
                    return;
                }
                AuthenticateConsumer authenticateConsumer = AuthenticateConsumerAMC(request, service);
                var serializedResponse = JsonSerializer.Serialize(authenticateConsumer);
                context.OutputParameters["data"] = serializedResponse;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Plugin Error: " + ex.Message);
            }

        }
        public AuthenticateConsumer AuthenticateConsumerAMC(AuthenticateConsumer requestParam, IOrganizationService _service)
        {
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
                    if (!CheckSourceType(_service,requestParam.SourceType, out expTime, out _portalSURL))
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
                        _retValue.RedirectUrl = _portalSURL + _postfixURL;
                        _retValue.StatusDescription = "Success";
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
        private bool CheckSourceType(IOrganizationService service,string sourceOrigin, out int expTime, out string portalURL)
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
    public class AuthenticateConsumer
    {
        public string LoginUserId { get; set; }
        public string SourceType { get; set; }
        public string SessionId { get; set; }
        public string RedirectUrl { get; set; }
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
    }
    public class AuthenticateConsumerRequestValidator
    {
        public static bool Validate(AuthenticateConsumer request, out string validationMessage)
        {
            if (string.IsNullOrEmpty(request.LoginUserId))
            {
                validationMessage = "Please enter LoginUserId";
                return false;
            }

            if (string.IsNullOrEmpty(request.SourceType))
            {
                validationMessage = "Please enter SourceType";
                return false;
            }
            validationMessage = string.Empty;
            return true;
        }
    }
}
