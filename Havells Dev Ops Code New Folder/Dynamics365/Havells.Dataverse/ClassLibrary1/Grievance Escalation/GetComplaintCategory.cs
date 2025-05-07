using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Net;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.NPS

{
    public class GetComplaintCategory : IPlugin
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

                var _response = GetComplaintCategoryList(service);
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
                var TechnicianScoreValidResponse = JsonSerializer.Serialize(new LitmusFeedBackResponse { Status = false, Message = ex.Message });
                context.OutputParameters["data"] = TechnicianScoreValidResponse;
                return;
            }
        }

        public (Response, RequestStatus) GetComplaintCategoryList(IOrganizationService _CrmService)
        {

            Response res = new Response();
            res.ComplaintCategories = new List<ComplaintCategory>();
            try
            {
                string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='hil_casecategory'>
                                    <attribute name='hil_name' />
                                    <attribute name='hil_casedepartment' />
                                    <order attribute='hil_name' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='statecode' operator='eq' value='0' />
                                      <condition attribute='hil_casedepartment' operator='ne' value='{{7BF1705A-3764-EE11-8DF0-6045BDAA91C3}}'/>
                                    </filter>
                                  </entity>
                                </fetch>";
                EntityCollection result = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml));
                if (result.Entities.Count > 0)
                {
                    foreach (var entity in result.Entities)
                    {
                        ComplaintCategory tempComplaint = new ComplaintCategory();
                        tempComplaint.CategoryGUID = entity.Id.ToString();
                        tempComplaint.CategoryName = entity.Contains("hil_name") ? entity.GetAttributeValue<string>("hil_name") : "";
                        tempComplaint.Department = entity.Contains("hil_casedepartment") ? entity.GetAttributeValue<EntityReference>("hil_casedepartment").Name : "";
                        res.ComplaintCategories.Add(tempComplaint);

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

    public class ComplaintCategory
    {
        public string CategoryGUID { get; set; }
        public string CategoryName { get; set; }
        public string Department { get; set; }

    }
    public class Response
    {
        public List<ComplaintCategory> ComplaintCategories { get; set; }
    }

    public class RequestStatus
    {
        public int StatusCode { get; set; }
        public string Message { get; set; }
    }
}