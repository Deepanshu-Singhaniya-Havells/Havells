using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Net;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Grievance_Escalation
{
    public class RaiseRemindercs : IPlugin
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
                string ComplaintGuid = Convert.ToString(context.InputParameters["ComplaintGuid"]);
                string ReminderRemarks = Convert.ToString(context.InputParameters["ReminderRemarks"]);
                if (string.IsNullOrWhiteSpace(ComplaintGuid))
                {

                    var jobIdResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = (int)HttpStatusCode.BadRequest, Message = "ComplaintGuid required" });
                    context.OutputParameters["data"] = jobIdResponse;
                    return;

                }
                if (!string.IsNullOrWhiteSpace(ComplaintGuid))
                {
                    if (!APValidate.IsvalidGuid(ComplaintGuid))
                    {
                        var jobIdResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = (int)HttpStatusCode.BadRequest, Message = "ComplaintGuid is not Valid" });
                        context.OutputParameters["data"] = jobIdResponse;
                        return;
                    }

                }
                RaiseRemindercs obj = new RaiseRemindercs();
                var _response = obj.RaiseReminders(service, new RaiseReminder_Request { ComplaintGuid = ComplaintGuid, ReminderRemarks = ReminderRemarks });
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
        public (RaiseReminder_Response, RequestStatus) RaiseReminders(IOrganizationService _CrmService, RaiseReminder_Request obj)
        {

            RaiseReminder_Response res = new RaiseReminder_Response();         
            Entity incident = _CrmService.Retrieve("incident", new Guid(obj.ComplaintGuid), new ColumnSet("hil_isreminderset", "hil_reminderremarks", "hil_reminderdate"));
            try
            {
                incident["hil_isreminderset"] = true;
                incident["hil_reminderremarks"] = obj.ReminderRemarks;
                incident["hil_reminderdate"] = DateTime.Now.AddMinutes(330);
                _CrmService.Update(incident);
                res.Status = true;
                res.Message = "reminder raised successfully";
                return (res, new RequestStatus
                {
                    StatusCode = (int)HttpStatusCode.OK,
                });
            }
            catch (Exception ex)
            {
                return (res, new RequestStatus
                {
                    StatusCode = (int)HttpStatusCode.InternalServerError,
                    Message = "error: " + ex.Message
                });
            }
        }
    }
    public class RaiseReminder_Response
    {
        public bool Status { get; set; }
        public string Message { get; set; }
    }
    public class RaiseReminder_Request
    {
        public string ComplaintGuid { get; set; }
        public string ReminderRemarks { get; set; }
    }
}