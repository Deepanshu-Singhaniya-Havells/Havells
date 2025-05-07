using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Net;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Grievance_Escalation
{
    public class GetOpenJobs : IPlugin
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
                string CustomerGuid = Convert.ToString(context.InputParameters["CustomerGuid"]);
                if (string.IsNullOrWhiteSpace(CustomerGuid))
                {

                    var jobIdResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = (int)HttpStatusCode.BadRequest, Message = "CustomerGuid required" });
                    context.OutputParameters["data"] = jobIdResponse;
                    return;

                }
                if (!string.IsNullOrWhiteSpace(CustomerGuid))
                {
                    if (!APValidate.IsvalidGuid(CustomerGuid))
                    {
                        var jobIdResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = (int)HttpStatusCode.BadRequest, Message = "CustomerGuid is not Valid" });
                        context.OutputParameters["data"] = jobIdResponse;
                        return;
                    }

                }

                var _response = GetOpenJob(service, CustomerGuid);
                if (_response.Item2.StatusCode == (int)HttpStatusCode.OK)
                {
                    context.OutputParameters["data"] = JsonSerializer.Serialize(_response.Item1);
                }
                else
                {
                    var RequestStatus = new
                    {
                        StatusCode = _response.Item2.StatusCode,
                        Message = _response.Item2.Message
                    };
                    context.OutputParameters["data"] = JsonSerializer.Serialize(RequestStatus);
                }
            }
            catch (Exception ex)
            {
                var TechnicianScoreValidResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = (int)HttpStatusCode.InternalServerError, Message = ex.Message });
                context.OutputParameters["data"] = TechnicianScoreValidResponse;
                return;
            }
        }

        public (OpenJobsResponse, RequestStatus) GetOpenJob(IOrganizationService _CrmService, string CustomerGuid)
        {

            OpenJobsResponse res = new OpenJobsResponse();
            res.ServiceRequests = new List<OpenJobs_ServiceRequest>();
            try
            {
                string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='msdyn_workorder'>
                                <attribute name='msdyn_name' />
                                <attribute name='msdyn_workorderid' />
                                <attribute name='msdyn_substatus' />
                                <attribute name='createdon' />
                                <attribute name='hil_productcatsubcatmapping' />
                                <order attribute='msdyn_name' descending='false' />
                                <filter type='and'>
                                  <condition attribute='hil_customerref' operator='eq' value='{CustomerGuid}' />
                                  <condition attribute='msdyn_substatus' operator='not-in'>
                                    <value uiname='Canceled' uitype='msdyn_workordersubstatus'>{{1527FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                                    <value uiname='Closed' uitype='msdyn_workordersubstatus'>{{1727FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                                    <value uiname='Work Done SMS' uitype='msdyn_workordersubstatus'>{{7E85074C-9C54-E911-A951-000D3AF0677F}}</value>
                                    <value uiname='Work Done' uitype='msdyn_workordersubstatus'>{{2927FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                                  </condition>
                                </filter>
                              </entity>
                            </fetch>";

                EntityCollection result = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml));

                if (result.Entities.Count > 0)
                {
                    foreach (var entity in result.Entities)
                    {
                        OpenJobs_ServiceRequest tempComplaint = new OpenJobs_ServiceRequest();

                        tempComplaint.ServiceRequestId = entity.Contains("msdyn_name") ? entity.GetAttributeValue<string>("msdyn_name") : "";
                        tempComplaint.ServiceRequestGUID = entity.Id.ToString();
                        tempComplaint.Status = entity.Contains("msdyn_substatus") ? entity.GetAttributeValue<EntityReference>("msdyn_substatus").Name : "";
                        tempComplaint.Created_On = entity.Contains("createdon") ? entity.GetAttributeValue<DateTime>("createdon").Date.ToString("d") : "";
                        tempComplaint.Product_SubCategory = entity.Contains("hil_productcatsubcatmapping") ? entity.GetAttributeValue<EntityReference>("hil_productcatsubcatmapping").Name : "";
                        res.ServiceRequests.Add(tempComplaint);

                    }
                }
                else
                {
                    return (res, new RequestStatus
                    {
                        StatusCode = (int)HttpStatusCode.NotFound,
                        Message = "Record Not found"
                    });
                }


                return (res, new RequestStatus
                {
                    StatusCode = (int)HttpStatusCode.OK
                });
            }
            catch (Exception ex)
            {
                return (res, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = "D365 internal server error : " + ex.Message.ToUpper()
                });
            }
        }
    }

    public class OpenJobsRequest
    {
        public string CustomerGuid { get; set; }
    }

    public class OpenJobs_ServiceRequest
    {
        public string ServiceRequestId { get; set; }

        public string ServiceRequestGUID { get; set; }

        public string Status { get; set; }

        public string Created_On { get; set; }

        public string Product_SubCategory { get; set; }

    }
    public class OpenJobsResponse
    {
        public string Error { get; set; }
        public List<OpenJobs_ServiceRequest> ServiceRequests { get; set; }
    }

    public class RequestStatus
    {
        public int StatusCode { get; set; }
        public string Message { get; set; }
    }
}
