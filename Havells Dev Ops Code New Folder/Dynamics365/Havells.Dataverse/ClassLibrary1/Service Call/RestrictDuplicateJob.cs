using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace Havells.Dataverse.CustomConnector.Service_Call
{
    public class RestrictDuplicateJob : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion
            var response = new Response();
            #region Extract value from Input parameter
            var customerId = context.InputParameters["CustomerId"] as string;
            var mobileNumber = context.InputParameters["Mobilenumber"] as string;
            var productCategoryGuid = context.InputParameters["ProductCategoryGuid"] as string;
            var callSubType = context.InputParameters["CallSubType"] as string;
            var productSubCategoryGuid = context.InputParameters["ProductSubCategoryGuid"] as string;
            var nocGuid = context.InputParameters["NOCGuid"] as string;
            #endregion
            #region Validation Check
            if (string.IsNullOrWhiteSpace(customerId) || !APValidate.IsvalidGuid(customerId))
            {
                SetErrorResponse(context, response, "Valid CustomerId is required.");
                return;
            }
            if (string.IsNullOrWhiteSpace(mobileNumber) || !APValidate.IsValidMobileNumber(mobileNumber))
            {
                SetErrorResponse(context, response, "Valid Mobile Number is required.");
                return;
            }
            if (string.IsNullOrWhiteSpace(productCategoryGuid) || !APValidate.IsvalidGuid(productCategoryGuid))
            {
                SetErrorResponse(context, response, "Valid Product Category Guid is required.");
                return;
            }
            if (string.IsNullOrWhiteSpace(productSubCategoryGuid) || !APValidate.IsvalidGuid(productSubCategoryGuid))
            {
                SetErrorResponse(context, response, "Valid Product Sub Category Guid is required.");
                return;
            }
            if (string.IsNullOrWhiteSpace(callSubType) && string.IsNullOrWhiteSpace(nocGuid))
            {
                SetErrorResponse(context, response, "Either Call Sub Type OR NOC Guid is required.");
                return;
            }
            if (!string.IsNullOrWhiteSpace(nocGuid))
            {
                if (!APValidate.IsvalidGuid(nocGuid))
                {
                    SetErrorResponse(context, response, "Valid NOC Guid is required.");
                    return;
                }
            }

            #endregion
            try
            {
                var duplicateJobData = GetDuplicateJobs(service, customerId, mobileNumber, productCategoryGuid, productSubCategoryGuid, callSubType, nocGuid);
                context.OutputParameters["data"] = !string.IsNullOrEmpty(duplicateJobData) ? duplicateJobData : null;
            }
            catch (Exception ex)
            {
                response.Message = $"D365 Internal Server Error: {ex.Message}";
                response.IsDuplicate = false;
                response.Duplicatejob = null;
                context.OutputParameters["data"] = JsonConvert.SerializeObject(response);
            }
        }
        private void SetErrorResponse(IPluginExecutionContext context, Response response, string message)
        {
            response.IsDuplicate = false;
            response.Duplicatejob = null;
            response.Message = message;
            context.OutputParameters["data"] = JsonConvert.SerializeObject(response);
        }
        private string GetDuplicateJobs(IOrganizationService service, string customerId, string mobileNumber, string productCategoryGuid, string productSubCategoryGuid, string callSubType, string nocGuid)
        {
            var response = new Response
            {
                IsDuplicate = false,
                Duplicatejob = new List<Duplicate_job>()
            };
            var callSubTypeName = string.Empty;
            Guid _nocGuid = string.IsNullOrWhiteSpace(nocGuid) ? Guid.Empty : new Guid(nocGuid);
            if (_nocGuid != Guid.Empty)
            {
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='hil_natureofcomplaint'>
                    <attribute name='hil_callsubtype' />
                    <filter type='and'>
                    <condition attribute='hil_natureofcomplaintid' operator='eq' value='{_nocGuid}' />
                    </filter>
                    </entity>
                    </fetch>";

                var nocRef = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (nocRef.Entities.Count > 0)
                {
                    EntityReference CallSubTypeRef = nocRef.Entities[0].GetAttributeValue<EntityReference>("hil_callsubtype");
                    callSubTypeName = CallSubTypeRef.Id.ToString();
                }
            }
            else
                callSubTypeName = GetCallSubTypeName(service, callSubType);


            if (!string.IsNullOrWhiteSpace(callSubTypeName))
            {
                try
                {
                    var fetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='msdyn_workorder'>
                                        <attribute name='msdyn_name'/>
                                        <filter type='and'>
                                            <condition attribute='hil_mobilenumber' operator='eq' value='{mobileNumber}'/>
                                            <condition attribute='hil_productsubcategory' operator='eq' value='{productSubCategoryGuid}'/>
                                            <condition attribute='hil_customerref' operator='eq' value='{customerId}'/>
                                            <condition attribute='hil_productcategory' operator='eq' value='{productCategoryGuid}'/>
                                            <condition attribute='hil_callsubtype' operator='eq' value='{callSubTypeName}'/>
                                            <condition attribute='statecode' operator='eq' value='0'/>
                                            <condition attribute='msdyn_substatus' operator='not-in'>
                                            <value uiname='Canceled' uitype='msdyn_workordersubstatus'>{{1527FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                                            <value uiname='Closed' uitype='msdyn_workordersubstatus'>{{1727FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                                            <value uiname='KKG Audit Failed' uitype='msdyn_workordersubstatus'>{{6C8F2123-5106-EA11-A811-000D3AF057DD}}</value>
                                            <value uiname='Work Done' uitype='msdyn_workordersubstatus'>{{2927FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                                            </condition>
                                        </filter>
                                    </entity>
                                  </fetch>";

                    var duplicateJobs = service.RetrieveMultiple(new FetchExpression(fetch));
                    if (duplicateJobs.Entities.Count > 0)
                    {
                        response.IsDuplicate = true;
                        foreach (var entity in duplicateJobs.Entities)
                        {
                            response.Duplicatejob.Add(new Duplicate_job
                            {
                                jobnumber = entity.GetAttributeValue<string>("msdyn_name")
                            });
                        }
                    }
                    else
                    {
                        response.IsDuplicate = false;
                        response.Duplicatejob = null;
                        response.Message = "Duplicate Job Not Found";
                    }
                }
                catch (Exception ex)
                {
                    response.IsDuplicate = false;
                    response.Message = ex.Message;
                }
            }
            else
            {
                response.IsDuplicate = false;
                response.Duplicatejob = null;
                response.Message = "Call Sub Type Not Found";
            }
            return JsonConvert.SerializeObject(response);
        }
        private string GetCallSubTypeName(IOrganizationService service, string callSubType)
        {
            var fetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_callsubtype'>
                                <attribute name='hil_callsubtypeid'/>
                                <attribute name='hil_name'/>
                                <filter type='and'>
                                    <condition attribute='hil_name' operator='eq' value='{callSubType}'/>
                                </filter>
                            </entity>
                          </fetch>";

            var callSubTypeCollection = service.RetrieveMultiple(new FetchExpression(fetch));
            return callSubTypeCollection.Entities.Count > 0 ? callSubTypeCollection[0].Id.ToString() : string.Empty;
        }
    }
    public class Response
    {
        public bool IsDuplicate { get; set; }
        public string Message { get; set; }
        public List<Duplicate_job> Duplicatejob { get; set; }
    }

    public class Duplicate_job
    {
        public string jobnumber { get; set; }
    }
}
