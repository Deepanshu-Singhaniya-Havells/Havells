using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Text.Json;


namespace Havells.Dataverse.CustomConnector.AMC
{
    public class ValidateAMCSession : IPlugin
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
                #region Class Object
                ValidateSessionResponse validateSessionResponse = new ValidateSessionResponse();
                ValidateSessionRequest requestParam = new ValidateSessionRequest();
                #endregion
                #region extract 

                string SessionId = Convert.ToString(context.InputParameters["SessionId"]);
                bool KeepSessionLive = Convert.ToBoolean(context.InputParameters["KeepSessionLive"]);

                #endregion
                #region Validation
                if (string.IsNullOrWhiteSpace(SessionId))
                {
                    validateSessionResponse.StatusCode = "204";
                    validateSessionResponse.StatusDescription = "SessionId is required";
                    string jsonResultResponse = JsonSerializer.Serialize(validateSessionResponse);
                    context.OutputParameters["data"] = jsonResultResponse;
                    return;
                }
                Guid SessionIdGuid;
                if (Guid.TryParse(SessionId, out SessionIdGuid))
                {
                    if (APValidate.IsvalidGuid(SessionId))
                    {
                        requestParam.SessionId = SessionIdGuid.ToString();
                    }
                }
                else
                {
                    validateSessionResponse.StatusCode = "204";
                    validateSessionResponse.StatusDescription = "SessionId is not valid";
                    string jsonResultResponse = JsonSerializer.Serialize(validateSessionResponse);
                    context.OutputParameters["data"] = jsonResultResponse;
                    return;
                }
                if (string.IsNullOrEmpty(KeepSessionLive.ToString()))
                {
                    validateSessionResponse.StatusCode = "204";
                    validateSessionResponse.StatusDescription = "KeepSessionLive is required in boolean{True/False}";
                    string jsonResultKeepSessionLive = JsonSerializer.Serialize(validateSessionResponse);
                    context.OutputParameters["data"] = jsonResultKeepSessionLive;
                    return;
                }
                else
                {
                    requestParam.KeepSessionLive = KeepSessionLive;
                }
                #endregion
                ValidateSessionResponse result = ValidateSessionDetails(requestParam, service);
                string jsonResult = JsonSerializer.Serialize(result);
                context.OutputParameters["data"] = jsonResult;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Plugin Error: " + ex.Message);
            }
        }
        public ValidateSessionResponse ValidateSessionDetails(ValidateSessionRequest requestParam, IOrganizationService _service)
        {
            ValidateSessionResponse _retValue = new ValidateSessionResponse();
            try
            {

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
                    DateTime expDate = entCol[0].GetAttributeValue<DateTime>("hil_expiredon");
                    if (expDate > DateTime.Now)
                    {
                        if (requestParam.KeepSessionLive)
                        {
                            Entity _userSession = new Entity("hil_consumerloginsession", entCol.Entities[0].Id);
                            _userSession["hil_expiredon"] = DateTime.Now.AddMinutes(350);
                            _service.Update(_userSession);
                        }
                        _retValue.StatusCode = "200";
                        _retValue.StatusDescription = "Session Id is Valid";
                    }
                    else
                    {
                        _retValue.StatusCode = "400";
                        _retValue.StatusDescription = "Access Denied!!! Session has been expired";
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
}
