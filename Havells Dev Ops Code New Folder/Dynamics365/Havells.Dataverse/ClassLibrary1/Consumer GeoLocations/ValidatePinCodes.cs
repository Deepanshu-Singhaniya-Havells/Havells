using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Consumer_GeoLocations
{
    public class ValidatePinCodes : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion
            string JsonResponse = string.Empty;
            string PINCode = string.Empty;
            IoTPINCodes objPINCode = new IoTPINCodes();
            if (context.InputParameters.Contains("PINCode") && context.InputParameters["PINCode"] is string)
            {
                PINCode = context.InputParameters["PINCode"].ToString();
                _tracingService.Trace("PINCode: " + PINCode);
                if (string.IsNullOrWhiteSpace(PINCode))
                {
                    objPINCode.StatusCode = "204";
                    objPINCode.StatusDescription = "PIN Code is required.";
                    JsonResponse = JsonSerializer.Serialize(objPINCode);
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                else
                {
                    JsonResponse = JsonSerializer.Serialize(IoTValidatePINCode(PINCode, service));
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
            }
            else
            {
                _tracingService.Trace("PINCode not found in input parameters.");
                objPINCode.StatusCode = "400";
                objPINCode.StatusDescription = "Invalid input. PIN Code is missing.";
                JsonResponse = JsonSerializer.Serialize(objPINCode);
                context.OutputParameters["data"] = JsonResponse;
                return;
            }
        }
        public IoTPINCodes IoTValidatePINCode(string PINCode, IOrganizationService service)
        {
            IoTPINCodes objPINCode = new IoTPINCodes();
            try
            {
                if (string.IsNullOrWhiteSpace(PINCode))
                {
                    objPINCode.StatusCode = "204";
                    objPINCode.StatusDescription = "PIN Code is required.";
                    return objPINCode;
                }
                if (service != null)
                {
                    QueryExpression query = new QueryExpression("hil_pincode")
                    {
                        ColumnSet = new ColumnSet("hil_pincodeid", "hil_name"),
                        Criteria = new FilterExpression(LogicalOperator.And)
                    };

                    query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, PINCode);
                    EntityCollection entcoll = service.RetrieveMultiple(query);
                    if (entcoll.Entities.Count == 0)
                    {
                        objPINCode.StatusCode = "204";
                        objPINCode.StatusDescription = "No PIN Code found.";
                    }
                    else
                    {
                        objPINCode.PINCode = entcoll.Entities[0].GetAttributeValue<string>("hil_name");
                        objPINCode.PINCodeGuid = entcoll.Entities[0].GetAttributeValue<Guid>("hil_pincodeid");
                        objPINCode.StatusCode = "200";
                        objPINCode.StatusDescription = "OK.";
                    }
                }
                else
                {
                    objPINCode.StatusCode = "503";
                    objPINCode.StatusDescription = "D365 Service Unavailable.";
                }
            }
            catch (Exception ex)
            {
                objPINCode.StatusCode = "500";
                objPINCode.StatusDescription = "D365 Internal Server Error: " + ex.Message;
            }
            return objPINCode;
        }
        public class IoTPINCodes
        {
            public Guid PINCodeGuid { get; set; }
            public string PINCode { get; set; }
            public string StatusCode { get; set; }
            public string StatusDescription { get; set; }
        }
    }
}
