using HavellsSync_ModelData.Common;
using Microsoft.Extensions.Configuration;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using HavellsSync_Data.IManager;
using Microsoft.Crm.Sdk.Messages;
using System.Reflection.Metadata;
using System.ServiceModel;
using HavellsSync_ModelData.ICommon;
using System.Net;
using System.ServiceModel.Channels;

namespace HavellsSync_Data.Manager
{
    public class AuthenticationManager : IAuthenticationManager
    {
        private readonly IConfiguration _configuration;
        private readonly ICrmService _service;
        private readonly string ScourceType;
        private readonly IAES256 _AES256;
        public AuthenticationManager(IConfiguration configuration, ICrmService crmService, IAES256 aES256)
        {
            _configuration = configuration;
            Check.Argument.IsNotNull(nameof(crmService), crmService);
            _service = crmService;
            _AES256 = aES256;
        }

        public AuthResponse AuthenticateUser(AuthModel objAuth)
        {
            AuthResponse _retValue = new AuthResponse()
            {
                LoginUserId = objAuth.LoginUserId,
                SourceType = objAuth.SourceType
            };
            Entity _userSession;
            try
            {
                if (objAuth.SourceType == string.Empty && objAuth.SourceType == null)// != "5")
                {
                    _retValue.StatusCode = 400;
                    _retValue.Message = CommonMessage.AccessdeniedMsg + CommonMessage.SourcetypeMsg;
                    return _retValue;
                }
                if (objAuth.LoginUserId == string.Empty || objAuth.LoginUserId == null)
                {
                    _retValue.StatusCode = 400;
                    _retValue.Message = CommonMessage.AccessdeniedMsg + CommonMessage.LoginUserIdMsg;
                    //_retValue.Message = "Access denied!!! Login User Id is mandatory.";
                    return _retValue;
                }
                if (objAuth.SourceType != "6")
                {
                    _retValue.StatusCode = 400;
                    _retValue.Message = CommonMessage.AccessdeniedMsg + CommonMessage.InvalidSourcetypeMsg;
                    //_retValue.Message = "Access denied!!! Invalid Source type.";
                    return _retValue;
                }
                if (_service != null)
                {
                    int expTime = 0;
                    string _portalSURL = string.Empty;
                    QueryExpression queryExp;
                    EntityCollection entCol;
                    queryExp = new QueryExpression("contact");
                    queryExp.ColumnSet = new ColumnSet(false);
                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    queryExp.Criteria.AddCondition(new ConditionExpression("mobilephone", ConditionOperator.Equal, objAuth.LoginUserId));
                    queryExp.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));//Active
                    entCol = _service.RetrieveMultiple(queryExp);
                    if (entCol.Entities.Count <= 0) //No Active User found
                    {
                        _retValue.StatusCode = 400;
                        _retValue.Message = CommonMessage.AccessdeniedMsg + CommonMessage.UserdoesnotexistMsg;
                        //  _retValue.Message = "Access denied!!! User does not exist.";
                        return _retValue;

                    }
                    if (!CheckSourceType(_service, objAuth.SourceType, out expTime, out _portalSURL))
                    {
                        _retValue.StatusCode = 400;
                        _retValue.Message = CommonMessage.AccessdeniedMsg + CommonMessage.ExtSourceTypeMsg + " " + objAuth.SourceType;
                        //  _retValue.Message = "Access denied!!! API is not extended to Source Type: " + objAuth.SourceType;
                        return _retValue;
                    }
                    else
                    {
                        queryExp = new QueryExpression("hil_consumerloginsession");
                        queryExp.ColumnSet = new ColumnSet("hil_expiredon", "hil_consumerloginsessionid");
                        ConditionExpression condExp = new ConditionExpression("hil_name", ConditionOperator.Equal, objAuth.LoginUserId);
                        queryExp.Criteria.AddCondition(condExp);
                        condExp = new ConditionExpression("statecode", ConditionOperator.Equal, 0);//Active
                        condExp = new ConditionExpression("hil_origin", ConditionOperator.Equal, objAuth.SourceType);//Active
                        queryExp.Criteria.AddCondition(condExp);
                        queryExp.TopCount = 1;
                        queryExp.AddOrder("hil_expiredon", OrderType.Descending);
                        bool isUpdated = false;
                        entCol = _service.RetrieveMultiple(queryExp);
                        if (entCol.Entities.Count > 0) //No Active session found
                        {
                            DateTime expiration = entCol.Entities[0].Contains("hill_expiration") ? entCol.Entities[0].GetAttributeValue<DateTime>("hill_expiration").AddMinutes(330) : DateTime.Now;
                            if (expiration > DateTime.Now)
                            {
                                _userSession = new Entity("hil_consumerloginsession", entCol.Entities[0].Id);
                                _userSession["hil_expiredon"] = expiration.AddMinutes(expTime);
                                _service.Update(_userSession);
                                _retValue.AccessToken = entCol.Entities[0].Id.ToString();
                                _retValue.TokenExpiresAt = _userSession["hil_expiredon"].ToString();
                                isUpdated = true;
                            }
                            else
                            {
                                foreach (Entity entity in entCol.Entities)
                                {
                                    entity["statecode"] = new OptionSetValue(1); //Inactive
                                    entity["statuscode"] = new OptionSetValue(2); //Inactive
                                    _service.Update(entity);
                                }
                            }
                        }
                        if (!isUpdated)
                        {
                            _userSession = new Entity("hil_consumerloginsession");
                            _userSession["hil_name"] = objAuth.LoginUserId;
                            _userSession["hil_origin"] = objAuth.SourceType;

                            _userSession["hil_expiredon"] = DateTime.Now.AddMinutes(expTime + 330);
                            _retValue.AccessToken = _service.Create(_userSession).ToString();
                            _retValue.TokenExpiresAt = DateTime.Now.AddMinutes(expTime).ToString();// _userSession["hil_expiredon"].ToString()
                        }
                    }

                    //_retValue.Message = "Success";
                    _retValue.Message = CommonMessage.SuccessMsg;
                    _retValue.StatusCode = 200;
                }
                else
                {
                    _retValue.StatusCode = 503;
                    _retValue.Message = CommonMessage.ServiceUnavailableMsg;
                    //_retValue.Message = "D365 service is unavailable.";
                }
            }
            catch (Exception ex)
            {
                _retValue.StatusCode = 500;
                _retValue.Message = CommonMessage.InternalServerErrorMsg + ex.Message;
                //_retValue.Message = "D365 internal server error : " + ex.Message;
            }
            return _retValue;
        }
        public validatesessionResponse ValidateSessionDetails(ValidateSession requestParam, string LoginUserId)
        {
            validatesessionResponse _retValue = new validatesessionResponse();
            try
            {
                requestParam.SessionId = _AES256.DecryptAES256(requestParam.SessionId);
                LoginUserId = _AES256.DecryptAES256(LoginUserId);

                if (string.IsNullOrEmpty(requestParam.SessionId) || string.IsNullOrEmpty(LoginUserId))
                {
                    _retValue.StatusCode = (int)HttpStatusCode.BadRequest;
                    _retValue.Message = _AES256.EncryptAES256(JsonConvert.SerializeObject(new { StatusCode = (int)HttpStatusCode.BadRequest, Message = CommonMessage.AccessdeniedMsg + CommonMessage.InvalidsessionMsg }));
                    return _retValue;
                }

                int expTime = 0; string _portalSURL;
                string SourceType = "";
                QueryExpression queryExp = new QueryExpression("hil_consumerloginsession");
                queryExp.ColumnSet = new ColumnSet("hil_expiredon", "hil_consumerloginsessionid", "hil_origin");
                ConditionExpression condExp = new ConditionExpression("hil_consumerloginsessionid", ConditionOperator.Equal, requestParam.SessionId);
                queryExp.Criteria.AddCondition(condExp);
                condExp = new ConditionExpression("statecode", ConditionOperator.Equal, 0);
                queryExp.Criteria.AddCondition(condExp);
                EntityCollection entCol = _service.RetrieveMultiple(queryExp);
                if (entCol.Entities.Count > 0)
                {
                    SourceType = entCol[0].GetAttributeValue<string>("hil_origin");
                }
                if (!CheckSourceType(_service, SourceType, out expTime, out _portalSURL))
                {
                    _retValue.StatusCode = (int)HttpStatusCode.BadRequest;

                    _retValue.Message = _AES256.EncryptAES256(JsonConvert.SerializeObject(new
                    {
                        StatusCode = (int)HttpStatusCode.BadRequest,
                        Message = CommonMessage.AccessdeniedMsg + CommonMessage.InvalidsessionMsg
                        //Message = "Access denied!!! Invalid access token."
                    }));
                    return _retValue;
                }
                if (entCol.Entities.Count == 1)
                {
                    DateTime expDate = entCol[0].GetAttributeValue<DateTime>("hil_expiredon").AddMinutes(330);
                    if (expDate > DateTime.Now)
                    {
                        if (requestParam.KeepSessionLive)
                        {
                            Entity _userSession = new Entity("hil_consumerloginsession", entCol.Entities[0].Id);
                            _userSession["hil_expiredon"] = DateTime.Now.AddMinutes(330 + expTime);
                            _service.Update(_userSession);
                            _retValue.TokenExpiresAt = _userSession["hil_expiredon"].ToString();
                        }
                        _retValue.StatusCode = (int)HttpStatusCode.OK;
                        _retValue.Message = CommonMessage.SuccessMsg;
                        //_retValue.Message = "Success";
                        return _retValue;
                    }
                    else
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

                        _retValue.StatusCode = (int)HttpStatusCode.BadRequest;
                        _retValue.Message = _AES256.EncryptAES256(JsonConvert.SerializeObject(new
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = CommonMessage.AccessdeniedMsg + CommonMessage.InvalidsessionMsg
                            // Message = "Unauthorization to access!!!"
                        }));
                    }
                }
                else
                {
                    _retValue.StatusCode = (int)HttpStatusCode.BadRequest;
                    _retValue.Message = _AES256.EncryptAES256(JsonConvert.SerializeObject(new
                    {
                        StatusCode = (int)HttpStatusCode.BadRequest,
                        Message = CommonMessage.AccessdeniedMsg + CommonMessage.InvalidsessionMsg
                        // Message = "Invalid input session id."
                    }));
                }
            }
            catch (Exception ex)
            {
                _retValue.StatusCode = (int)HttpStatusCode.BadRequest;
                _retValue.Message = _AES256.EncryptAES256(JsonConvert.SerializeObject(new
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message
                    //Message = "Error!!! " + ex.Message
                }));
            }
            return _retValue;
        }
        private bool CheckSourceType(ICrmService service, string sourceOrigin, out int expTime, out string portalURL)
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
    }
}
