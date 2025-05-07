using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.AMC.Airtel_IQ
{
    public class AIQPushCDR : IPlugin
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
                var validationErrors = new List<string>();
                CDR_Request request = new CDR_Request
                {
                    Caller_Id = Convert.ToString(context.InputParameters["Caller_Id"]),
                    Caller_Name = Convert.ToString(context.InputParameters["Caller_Name"]),
                    Caller_Number = Convert.ToString(context.InputParameters["Caller_Number"]),
                    Call_Type = Convert.ToString(context.InputParameters["Call_Type"]),
                    Caller_Status = Convert.ToString(context.InputParameters["Caller_Status"]),
                    Conversation_Duration = Convert.ToString(context.InputParameters["Conversation_Duration"]),
                    Correlation_ID = Convert.ToString(context.InputParameters["Correlation_ID"]),
                    Date = Convert.ToString(context.InputParameters["Date"]),
                    Destination_Name = Convert.ToString(context.InputParameters["Destination_Name"]),
                    Destination_Number = Convert.ToString(context.InputParameters["Destination_Number"]),
                    Destination_Status = Convert.ToString(context.InputParameters["Destination_Status"]),
                    Overall_Call_Duration = Convert.ToString(context.InputParameters["Overall_Call_Duration"]),
                    Overall_Call_Status = Convert.ToString(context.InputParameters["Overall_Call_Status"]),
                    Recording = Convert.ToString(context.InputParameters["Recording"]),
                    Time = Convert.ToString(context.InputParameters["Time"])

                };

                if (string.IsNullOrWhiteSpace(request.Conversation_Duration))
                {
                    context.OutputParameters["data"] = JsonSerializer.Serialize(new CDR_Response { ResultStatus = false, ResultMessage = "Conversation_Duration is required." });
                    return;
                }
                if (!IsValidTimeFormat(request.Conversation_Duration))
                {
                    context.OutputParameters["data"] = JsonSerializer.Serialize(new CDR_Response { ResultStatus = false, ResultMessage = "Invalid Conversation_Duration format." });
                    return;
                }
                if (string.IsNullOrWhiteSpace(request.Overall_Call_Duration))
                {
                    context.OutputParameters["data"] = JsonSerializer.Serialize(new CDR_Response { ResultStatus = false, ResultMessage = "Overall_Call_Duration is required." });
                    return;
                }
                if (!IsValidTimeFormat(request.Overall_Call_Duration))
                {
                    context.OutputParameters["data"] = JsonSerializer.Serialize(new CDR_Response { ResultStatus = false, ResultMessage = "Invalid Overall_Call_Duration format." });
                    return;
                }
                if (string.IsNullOrWhiteSpace(request.Date))
                {
                    context.OutputParameters["data"] = JsonSerializer.Serialize(new CDR_Response { ResultStatus = false, ResultMessage = "Date is required." });
                    return;
                }
                if (string.IsNullOrWhiteSpace(request.Time))
                {
                    context.OutputParameters["data"] = JsonSerializer.Serialize(new CDR_Response { ResultStatus = false, ResultMessage = "Time is required." });
                    return;
                }
                if (!IsValidTimeFormat(request.Time))
                {
                    context.OutputParameters["data"] = JsonSerializer.Serialize(new CDR_Response { ResultStatus = false, ResultMessage = "Invalid Time format." });
                    return;
                }
                if (!IsValidDateFormat(request.Date))
                {
                    context.OutputParameters["data"] = JsonSerializer.Serialize(new CDR_Response { ResultStatus = false, ResultMessage = "Invalid date. Please try with this format (dd/MM/yyyy)." });
                    return;
                }


                CDR_Response response = PushCDRToD365(request, service);
                context.OutputParameters["data"] = JsonSerializer.Serialize(response);
            }
            catch (Exception ex)
            {
                var errorResponse = new CDR_Response
                {
                    ResultStatus = false,
                    ResultMessage = "D365 Internal Server Error: " + ex.Message
                };
                context.OutputParameters["data"] = JsonSerializer.Serialize(errorResponse);
            }
        }
        public CDR_Response PushCDRToD365(CDR_Request request, IOrganizationService service)
        {
            CDR_Response response = new CDR_Response();
            response.ResultMessage = "Failed";
            try
            {
                if ((service != null))
                {
                    string cdr_report = Newtonsoft.Json.JsonConvert.SerializeObject(request);

                    QueryExpression query = new QueryExpression("phonecall");
                    query.ColumnSet = new ColumnSet("description", "actualstart", "actualend", "hil_disposition", "phonenumber", "hil_calledtonum", "hil_callingnumber", "subject", "hil_alternatenumber1", "scheduledstart", "directioncode", "from", "to");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("subject", ConditionOperator.Equal, request.Correlation_ID);
                    EntityCollection phonecalls = service.RetrieveMultiple(query);
                    int count = phonecalls.Entities.Count();
                    TimeSpan endTime = TimeSpan.Parse(request.Time).Add(TimeSpan.Parse(request.Overall_Call_Duration));
                    DateTime start;
                    DateTime.TryParse(request.Date.Split('/')[2] + "-" + request.Date.Split('/')[1] + "-" + request.Date.Split('/')[0] + " " + request.Time, out start);
                    DateTime end = start.Add(TimeSpan.Parse(request.Overall_Call_Duration));
                    int status = 0;
                    if (request.Overall_Call_Status == "Missed") status = 8;
                    else if (request.Overall_Call_Status == "Answered") status = 9;

                    string convertation_duration = TimeSpan.Parse(request.Conversation_Duration).Seconds.ToString();

                    response.ResultStatus = true;
                    response.ResultMessage = "Success";
                    if (count == 0)
                    {

                        Entity phonecall = new Entity("phonecall");

                        phonecall["subject"] = request.Correlation_ID;
                        phonecall["actualstart"] = start;
                        phonecall["actualend"] = end;
                        phonecall["phonenumber"] = request.Caller_Id;
                        phonecall["hil_calledtonum"] = request.Destination_Number;
                        phonecall["hil_callingnumber"] = request.Caller_Number;
                        phonecall["description"] = "Caller: " + request.Caller_Number + " " + request.Caller_Name + ", Destination: " + request.Destination_Name + ", Caller Status: " + request.Caller_Status + ", Destination Status: " + request.Destination_Status;
                        phonecall["directioncode"] = false;
                        phonecall["hil_alternatenumber1"] = convertation_duration + " seconds";

                        if (!string.IsNullOrEmpty(request.Caller_Number))
                        {
                            request.Caller_Number = request.Caller_Number.Substring(Math.Max(0, request.Caller_Number.Length - 10));
                        }
                        else
                        {
                            request.Caller_Number = "";
                        }
                        if (!string.IsNullOrEmpty(request.Destination_Number))
                        {
                            request.Destination_Number = request.Destination_Number.Substring(Math.Max(0, request.Destination_Number.Length - 10));
                        }
                        else
                        {
                            request.Destination_Number = "";
                        }

                        phonecall["hil_disposition"] = new OptionSetValue(status);
                        Guid PhoneGuid = service.Create(phonecall);
                        CreateRecording(service, request, PhoneGuid);
                        response.ResultMessage = "Record Created Successfully";
                    }
                    else
                    {

                        Entity phonecall = new Entity(phonecalls.Entities[0].LogicalName, phonecalls.Entities[0].Id);
                        phonecall["actualstart"] = start.AddHours(-5.5);
                        phonecall["actualend"] = end.AddMinutes(-330);
                        phonecall["hil_alternatenumber1"] = convertation_duration + " seconds";
                        phonecall["description"] = "Caller: " + request.Caller_Number + " " + request.Caller_Name + ", Destination: " + request.Destination_Name + ", Caller Status: " + request.Caller_Status + ", Destination Status: " + request.Destination_Status; ;
                        phonecall["hil_disposition"] = new OptionSetValue(status);
                        service.Update(phonecall);
                        CreateRecording(service, request, phonecalls.Entities[0].Id);
                        response.ResultMessage = "Record Created Successfully";

                    }
                }
                else
                {
                    response.ResultStatus = false;
                    response.ResultMessage = "D365 Service is unavailable.";
                }
            }
            catch (Exception ex)
            {
                response.ResultStatus = false;
                response.ResultMessage = "Something went wrong ! " + ex.Message;
            }
            return response;
        }
        private void CreateRecording(IOrganizationService service, CDR_Request req, Guid ActivityId)
        {
            EntityCollection tempCollection = CheckExistingRecording(service, ActivityId);
            if (tempCollection.Entities.Count == 0)
            {
                Entity recording = new Entity("msdyn_recording");
                recording["msdyn_ci_url"] = req.Recording;
                recording["msdyn_ci_transcript_json"] = Newtonsoft.Json.JsonConvert.SerializeObject(req);
                recording["msdyn_phone_call_activity"] = new EntityReference("phonecall", ActivityId);
                service.Create(recording);
            }
            else
            {
                tempCollection.Entities[0]["msdyn_ci_url"] = req.Recording;
                tempCollection.Entities[0]["msdyn_ci_transcript_json"] = Newtonsoft.Json.JsonConvert.SerializeObject(req);
                tempCollection.Entities[0]["msdyn_phone_call_activity"] = new EntityReference("phonecall", ActivityId);
                service.Update(tempCollection.Entities[0]);
            }
        }
        private EntityCollection CheckExistingRecording(IOrganizationService service, Guid ActivityID)
        {
            QueryExpression query = new QueryExpression("msdyn_recording");
            query.ColumnSet = new ColumnSet("msdyn_ci_url", "msdyn_ci_transcript_json", "msdyn_phone_call_activity");
            query.Criteria.AddCondition("msdyn_phone_call_activity", ConditionOperator.Equal, ActivityID);
            return service.RetrieveMultiple(query);
        }
        private bool IsValidTimeFormat(string input)
        {
            string[] formats = { "HH:mm", "HH:mm:ss", "hh:mm tt", "hh:mm:ss tt" };
            return DateTime.TryParseExact(input, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out _);
        }
        private bool IsValidDateFormat(string input)
        {
            return DateTime.TryParseExact(input, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out _);
        }
        public class CDR_Response
        {
            public string ResultMessage { get; set; }
            public bool ResultStatus { get; set; }
        }
        public class CDR_Request
        {
            public string Caller_Id { get; set; }
            public string Caller_Name { get; set; }
            public string Caller_Number { get; set; }
            public string Call_Type { get; set; }
            public string Caller_Status { get; set; }
            public string Conversation_Duration { get; set; }
            public string Correlation_ID { get; set; }
            public string Date { get; set; }
            public string Destination_Name { get; set; }
            public string Destination_Number { get; set; }
            public string Destination_Status { get; set; }
            public string Overall_Call_Duration { get; set; }
            public string Overall_Call_Status { get; set; }
            public string Recording { get; set; }
            public string Time { get; set; }
        }
    }
}
