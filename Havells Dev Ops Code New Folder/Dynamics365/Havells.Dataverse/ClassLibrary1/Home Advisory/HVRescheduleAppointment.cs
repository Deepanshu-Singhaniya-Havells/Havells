using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Havells.Dataverse.CustomConnector.Home_Advisory
{
    public class HVRescheduleAppointment : IPlugin
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
                string jsonResponse = string.Empty;
                ReschduleAppointment request = new ReschduleAppointment
                {
                    AdvisorylineId = Convert.ToString(context.InputParameters["AdvisorylineId"]),
                    scheduledstart = Convert.ToString(context.InputParameters["scheduledstart"]),
                    scheduledEnd = Convert.ToString(context.InputParameters["scheduledEnd"]),
                    appointmenturl = Convert.ToString(context.InputParameters["appointmenturl"]),
                    appintmentId = Convert.ToString(context.InputParameters["appintmentId"]),
                    appintmentType = Convert.ToString(context.InputParameters["appintmentType"]),
                    Remarks = Convert.ToString(context.InputParameters["Remarks"]),
                    AssignedUsercode = Convert.ToString(context.InputParameters["AssignedUsercode"]),

                };

                if (string.IsNullOrWhiteSpace(request.AssignedUsercode))
                {
                    jsonResponse = JsonSerializer.Serialize(new Response { Status = false, Message = "AssignedUsercode is required." });
                    context.OutputParameters["data"] = jsonResponse;
                    return;
                }
                if (string.IsNullOrWhiteSpace(request.AdvisorylineId))
                {
                    jsonResponse = JsonSerializer.Serialize(new Response { Status = false, Message = "AdvisorylineId is required." });
                    context.OutputParameters["data"] = jsonResponse;
                    return;
                }
                if (!APValidate.NumericValue(request.AssignedUsercode))
                {
                    context.OutputParameters["data"] = JsonSerializer.Serialize(new Response { Status = false, Message = "Invalid AssignedUsercode." });
                    return;
                }
                if (request.AssignedUsercode.Length > 6)
                {
                    context.OutputParameters["data"] = JsonSerializer.Serialize(new Response { Status = false, Message = "Invalid AssignedUsercode." });
                    return;
                }
                if (!string.IsNullOrWhiteSpace(request.AdvisorylineId))
                {
                    if (!Regex.IsMatch(request.AdvisorylineId, @"^[0-9\-]+$"))
                    {
                        context.OutputParameters["data"] = JsonSerializer.Serialize(new Response
                        {
                            Status = false,
                            Message = "Invalid AdvisorylineId format."
                        });
                        return;
                    }
                    if (request.AdvisorylineId.Length > 19)
                    {
                        context.OutputParameters["data"] = JsonSerializer.Serialize(new Response { Status = false, Message = "Invalid AdvisorylineId." });
                        return;
                    }
                }
                if (string.IsNullOrWhiteSpace(request.scheduledstart) || (string.IsNullOrWhiteSpace(request.scheduledEnd)))
                {
                    jsonResponse = JsonSerializer.Serialize(new Response { Status = false, Message = "ScheduledStart and ScheduledEnd are required and must be in the format {yyyy-MM-dd}." });
                    context.OutputParameters["data"] = jsonResponse;
                    return;

                }
                if (!string.IsNullOrWhiteSpace(request.scheduledstart) || (!string.IsNullOrWhiteSpace(request.scheduledEnd)))
                {
                    if (!APValidate.IsvalidDate(request.scheduledstart) || (!APValidate.IsvalidDate(request.scheduledEnd)))
                    {
                        jsonResponse = JsonSerializer.Serialize(new Response { Status = false, Message = "Invalid date. Please try with this format (yyyy-MM-dd)" });
                        context.OutputParameters["data"] = jsonResponse;
                        return;
                    }
                    if (DateTime.TryParse(request.scheduledstart, out DateTime startDate) && DateTime.TryParse(request.scheduledEnd, out DateTime endDate))
                    {
                        if (startDate > endDate)
                        {
                            jsonResponse = JsonSerializer.Serialize(new Response
                            {
                                Status = false,
                                Message = "Scheduled start must be less than or equal to scheduled end."
                            });
                            context.OutputParameters["data"] = jsonResponse;
                            return;
                        }
                    }
                }
                if (!string.IsNullOrWhiteSpace(request.appintmentType))
                {
                    if (!ValidateApointmentType(request.appintmentType))
                    {
                        jsonResponse = JsonSerializer.Serialize(new Response { Status = false, Message = "Invalid appointmentType" });
                        context.OutputParameters["data"] = jsonResponse;
                        return;
                    }
                }

                Response response = RescheduleAppointment(request, service);
                jsonResponse = JsonSerializer.Serialize(response);
                context.OutputParameters["data"] = jsonResponse;
            }
            catch (Exception ex)
            {
                var Response = new Response
                {
                    Status = false,
                    Message = "D365 Internal Server Error: " + ex.Message
                };
                context.OutputParameters["data"] = JsonSerializer.Serialize(Response);
            }
        }
        public Response RescheduleAppointment(ReschduleAppointment req, IOrganizationService service)
        {
            Response res = new Response();
            try
            {
                string fetchAdvisoryline = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='hil_homeadvisoryline'>
                    <attribute name='hil_homeadvisorylineid' />
                    <attribute name='hil_advisoryenquery' />
                    <attribute name='hil_name' />
                    <attribute name='createdon' />
                    <order attribute='hil_name' descending='false' />
                    <filter type='and'>
                    <condition attribute='statecode' operator='eq' value='0' />
                    <condition attribute='hil_name' operator='eq' value='" + req.AdvisorylineId + @"' />
                    </filter>
                    </entity>
                    </fetch>";
                EntityCollection AdvisoryLineColl = service.RetrieveMultiple(new FetchExpression(fetchAdvisoryline));
                if (AdvisoryLineColl.Entities.Count == 0)
                {
                    res.Message = "Advisory Enquery Line Not Found in Dynamics";
                    res.Status = false;
                    return res;
                }
                Entity Advisoryline = AdvisoryLineColl.Entities[0];

                QueryExpression Query = new QueryExpression("hil_advisoryenquiry");
                Query.ColumnSet = new ColumnSet("hil_customer");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Advisoryline.GetAttributeValue<EntityReference>("hil_advisoryenquery").Name);
                EntityCollection Enquirycoll = service.RetrieveMultiple(Query);
                string fetchappointment = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='appointment'>
                        <attribute name='subject' />
                        <attribute name='statecode' />
                        <attribute name='scheduledstart' />
                        <attribute name='scheduledend' />
                        <attribute name='createdby' />
                        <attribute name='regardingobjectid' />
                        <attribute name='activityid' />
                        <attribute name='instancetypecode' />
                        <order attribute='subject' descending='false' />
                        <filter type='and'>
                            <condition attribute='statecode' operator='in'>
                            <value>0</value>
                            <value>3</value>
                            </condition>
                        </filter>
                        <link-entity name='hil_homeadvisoryline' from='hil_homeadvisorylineid' to='regardingobjectid' link-type='inner' alias='ab'>
                            <filter type='and'>
                            <condition attribute='hil_homeadvisorylineid' operator='eq' uiname='AdvEnq-000060' uitype='hil_homeadvisoryline' value='" + AdvisoryLineColl.Entities[0].Id + @"' />
                            </filter>
                        </link-entity>
                        </entity>
                    </fetch>";
                EntityCollection AdvlineAppointmentColl = service.RetrieveMultiple(new FetchExpression(fetchappointment));
                if (AdvlineAppointmentColl.Entities.Count > 0)
                {
                    Entity _app = new Entity("appointment");
                    _app.Id = AdvlineAppointmentColl.Entities[0].Id;
                    _app["statecode"] = new OptionSetValue(1);
                    _app["statuscode"] = new OptionSetValue(3);
                    _app["description"] = req.Remarks != null ? req.Remarks : "";
                    service.Update(_app);
                }
                Entity _appointment = new Entity("appointment");
                Entity from = new Entity("activityparty");
                EntityReference cust = Enquirycoll.Entities[0].GetAttributeValue<EntityReference>("hil_customer");
                from["partyid"] = cust;
                _appointment["requiredattendees"] = new Entity[] { from };
                _appointment["subject"] = "Meeting with Havells for Advisory";
                _appointment["location"] = "Teams";
                _appointment["regardingobjectid"] = new EntityReference("hil_homeadvisoryline", AdvisoryLineColl.Entities[0].Id);
                _appointment["scheduledstart"] = Convert.ToDateTime(req.scheduledstart);
                _appointment["scheduledend"] = Convert.ToDateTime(req.scheduledEnd);
                _appointment["hil_appointmenturl"] = req.appointmenturl.ToString();
                service.Create(_appointment);

                Entity _advLine = new Entity("hil_homeadvisoryline");
                _advLine.Id = AdvisoryLineColl.Entities[0].Id;
                if (req.appintmentId != null)
                    _advLine["hil_appointmentid"] = req.appintmentId;
                if (req.appointmenturl != null || req.appointmenturl != "")
                    _advLine["hil_videocallurl"] = req.appointmenturl;
                _advLine["hil_appointmentstatus"] = new OptionSetValue(6);

                _advLine["hil_appointmentdate"] = Convert.ToDateTime(req.scheduledstart);
                _advLine["hil_appointmentenddate"] = Convert.ToDateTime(req.scheduledEnd);

                _advLine["hil_appointmenttypes"] = req.appintmentType == "Video" ? new OptionSetValue(2) : new OptionSetValue(1);
                if (req.AssignedUsercode != null)
                {
                    QueryExpression advmasterQuery = new QueryExpression("hil_advisormaster");
                    advmasterQuery.ColumnSet = new ColumnSet(false);
                    advmasterQuery.Criteria = new FilterExpression(LogicalOperator.And);
                    advmasterQuery.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    advmasterQuery.Criteria.AddCondition("hil_code", ConditionOperator.Equal, req.AssignedUsercode.Trim());
                    EntityCollection Advisorymastercoll = service.RetrieveMultiple(advmasterQuery);
                    if (Advisorymastercoll.Entities.Count == 0)
                    {
                        res.Message = "Advisor Not Found in Dynamics";
                        res.Status = false;
                        return res;
                    }
                    _advLine["hil_assignedadvisor"] = new EntityReference("hil_advisormaster", Advisorymastercoll[0].Id);
                }

                service.Update(_advLine);
                res.Message = "Appointment reschdule sucessfully";
                res.Status = true;
                return res;
            }
            catch (Exception ex)
            {
                res.Message = (ex.Message);
                res.Status = false;
                return res;
            }
        }
        private bool ValidateApointmentType(string appointmentType)
        {
            return appointmentType.Trim() == "1" || appointmentType.Trim() == "2";
        }
        public class Response
        {
            public string Message { get; set; }
            public bool Status { get; set; }
        }
        public class ReschduleAppointment
        {
            public string AdvisorylineId { get; set; }
            public string scheduledstart { get; set; }
            public string scheduledEnd { get; set; }
            public string appointmenturl { get; set; }
            public string appintmentId { get; set; }
            public string appintmentType { get; set; }
            public string Remarks { get; set; }
            public string AssignedUsercode { get; set; }
        }
    }
}