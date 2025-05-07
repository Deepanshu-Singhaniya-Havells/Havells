using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Home_Advisory
{
    public class HVCreateAppointment : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion
            string _response = string.Empty;
            try
            {
                CRMRequest _request = new CRMRequest
                {
                    SlotStart = Convert.ToString(context.InputParameters["SlotStart"]),
                    SlotEnd = Convert.ToString(context.InputParameters["SlotEnd"]),
                    SlotDate = Convert.ToString(context.InputParameters["SlotDate"]),
                    RecordID = Convert.ToString(context.InputParameters["RecordID"]),
                    IsVideoMeeting = Convert.ToBoolean(context.InputParameters["IsVideoMeeting"])
                };
                _response = JsonSerializer.Serialize(CreateAppointmentD365(_request, service));
                context.OutputParameters["data"] = _response;
            }
            catch (Exception ex)
            {
                CreateUserMeetingResponse resp = new CreateUserMeetingResponse
                {
                    IsSuccess = false,
                    Message = "D365 Internal Server Error: " + ex.Message
                };
                context.OutputParameters["data"] = JsonSerializer.Serialize(resp);
            }
        }
        public Response CreateAppointmentD365(CRMRequest req, IOrganizationService service)
        {
            Response res = new Response();
            CreateUserMeetingResponse resp = new CreateUserMeetingResponse();
            try
            {
                if (req.SlotDate == string.Empty || req.SlotDate == null)
                {
                    res.Message = ("SlotDate is Null");
                    res.Status = false;
                    return res;
                }
                if (req.SlotEnd == string.Empty || req.SlotEnd == null)
                {
                    res.Message = ("SlotEnd is Null");
                    res.Status = false;
                    return res;
                }
                if (req.SlotStart == string.Empty || req.SlotStart == null)
                {
                    res.Message = ("SlotStart is Null");
                    res.Status = false;
                    return res;
                }
                if (req.RecordID == string.Empty || req.RecordID == null)
                {
                    res.Message = ("RecordID is Null");
                    res.Status = false;
                    return res;
                }
                #region get Privious Appointment

                #endregion
                #region
                QueryExpression Query = new QueryExpression("hil_homeadvisoryline");
                Query.ColumnSet = new ColumnSet("hil_appointmentdate", "hil_advisoryenquery", "hil_name", "hil_typeofenquiiry", "hil_typeofproduct", "hil_assignedadvisor", "hil_appointmentid");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_homeadvisorylineid", ConditionOperator.Equal, req.RecordID);

                LinkEntity EntityA = new LinkEntity("hil_homeadvisoryline", "hil_advisoryenquiry", "hil_advisoryenquery", "hil_advisoryenquiryid", JoinOperator.LeftOuter);
                EntityA.Columns = new ColumnSet("hil_customer", "hil_emailid");
                EntityA.EntityAlias = "PEnq";
                Query.LinkEntities.Add(EntityA);

                LinkEntity Entityb = new LinkEntity("hil_homeadvisoryline", "hil_advisormaster", "hil_assignedadvisor", "hil_advisormasterid", JoinOperator.LeftOuter);
                Entityb.Columns = new ColumnSet("hil_code");
                Entityb.EntityAlias = "Advisor";
                Query.LinkEntities.Add(Entityb);
                #endregion

                EntityCollection Found = service.RetrieveMultiple(Query);
                if (Found.Entities.Count == 1)
                {
                    String SlotEnd = string.Empty;
                    String SlotStart = string.Empty;
                    String EnquirerEmailId = string.Empty;
                    String EnquirerName = string.Empty;
                    string AppointmentID = string.Empty;
                    string EnquiryId = string.Empty;
                    string EnquiryType = string.Empty;
                    string AdvisoryType = string.Empty;

                    DateTime SlotDate = new DateTime(Convert.ToInt32(req.SlotDate.Substring(0, 4)), Convert.ToInt32(req.SlotDate.Substring(4, 2)), Convert.ToInt32(req.SlotDate.Substring(6, 2)));

                    String UserCode = string.Empty;
                    String ddymmyyyy = string.Empty;
                    EntityReference cust = new EntityReference();
                    Guid regarding = new Guid();
                    foreach (Entity line in Found.Entities)
                    {
                        SlotEnd = req.SlotEnd;
                        SlotStart = req.SlotStart;
                        EnquirerEmailId = line.Contains("PEnq.hil_emailid") ? (line["PEnq.hil_emailid"] as AliasedValue).Value.ToString() : throw new Exception("EnquirerEmailId not found");
                        EnquirerName = line.Contains("PEnq.hil_customer") ? ((EntityReference)((AliasedValue)line["PEnq.hil_customer"]).Value).Name.ToString() : throw new Exception("EnquirerName not found");

                        EnquiryId = line.Contains("hil_name") ? line.GetAttributeValue<String>("hil_name") : throw new Exception("EnquiryID not found");
                        EnquiryType = line.Contains("hil_typeofenquiiry") ? line.GetAttributeValue<EntityReference>("hil_typeofenquiiry").Name : throw new Exception("typeofenquiiry not found");
                        AdvisoryType = line.Contains("hil_typeofproduct") ? line.GetAttributeValue<EntityReference>("hil_typeofproduct").Name : throw new Exception("typeofproduct not found");

                        UserCode = line.Contains("Advisor.hil_code") ? (line["Advisor.hil_code"] as AliasedValue).Value.ToString() : throw new Exception("UserCode not found"); ;
                        cust = (EntityReference)((AliasedValue)line["PEnq.hil_customer"]).Value;
                        regarding = line.Id;

                        AppointmentID = line.Contains("hil_appointmentid") ? line.GetAttributeValue<String>("hil_appointmentid") : string.Empty;
                        if (AppointmentID == string.Empty)
                        {
                            String fetct = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">" +
                                              "<entity name=\"appointment\">" +
                                                "<attribute name=\"subject\" />" +
                                                "<attribute name=\"statecode\" />" +
                                                "<attribute name=\"scheduledstart\" />" +
                                                "<attribute name=\"scheduledend\" />" +
                                                "<attribute name=\"createdby\" />" +
                                                "<attribute name=\"regardingobjectid\" />" +
                                                "<attribute name=\"activityid\" />" +
                                                "<attribute name=\"instancetypecode\" />" +
                                                "<order attribute=\"subject\" descending=\"false\" />" +
                                                "<filter type=\"and\">" +
                                                  "<condition attribute=\"regardingobjectid\" operator=\"eq\" uiname=\"\" uitype=\"hil_homeadvisoryline\" value=\"" + regarding + "\" />" +
                                                  "<condition attribute=\"hil_appointmenturl\" operator=\"not-null\" />" +
                                                  "<condition attribute=\"statecode\" operator=\"in\">" +
                                                    "<value>0</value>" +
                                                     "<value>3</value>" +
                                                    "</condition>" +
                                                  "</filter>" +
                                               "</entity>" +
                                              "</fetch>";
                            EntityCollection _app = service.RetrieveMultiple(new FetchExpression(fetct));

                            if (_app.Entities.Count > 0)
                            {
                                res.Message = "Privious Appointment not Completed or Cancled";
                                res.Status = false;
                                return res;
                            }
                        }
                    }
                    if (AppointmentID == string.Empty)
                    {
                        if (resp.IsSuccess)
                        {
                            Entity _appointment = new Entity("appointment");
                            Entity from = new Entity("activityparty");
                            from["partyid"] = cust;
                            _appointment["requiredattendees"] = new Entity[] { from };
                            _appointment["subject"] = "Meeting with Havells for Advisory";
                            if (req.IsVideoMeeting)
                                _appointment["location"] = "Teams";
                            else
                                _appointment["location"] = "Audio Call";
                            _appointment["regardingobjectid"] = new EntityReference("hil_homeadvisoryline", regarding);
                            _appointment["scheduledstart"] = new DateTime(Convert.ToInt32(req.SlotDate.Substring(0, 4)), Convert.ToInt32(req.SlotDate.Substring(4, 2)), Convert.ToInt32(req.SlotDate.Substring(6, 2)), Convert.ToInt32(req.SlotStart.Substring(0, 2)), Convert.ToInt32(req.SlotStart.Substring(3, 2)), 0);
                            _appointment["scheduledend"] = new DateTime(Convert.ToInt32(req.SlotDate.Substring(0, 4)), Convert.ToInt32(req.SlotDate.Substring(4, 2)), Convert.ToInt32(req.SlotDate.Substring(6, 2)), Convert.ToInt32(req.SlotEnd.Substring(0, 2)), Convert.ToInt32(req.SlotEnd.Substring(3, 2)), 0);
                            if (resp.Data.MeetingURL != null && resp.Data.MeetingURL != string.Empty && resp.Data.MeetingURL != "")
                            {
                                _appointment["hil_appointmenturl"] = resp.Data.MeetingURL;
                            }
                            service.Create(_appointment);

                            Entity _enqLine = new Entity(Found.EntityName);
                            _enqLine.Id = Found.Entities[0].Id;

                            _enqLine["hil_appointmentid"] = resp.Data.TransactionId;
                            if (resp.Data.MeetingURL != "" || resp.Data.MeetingURL != null)
                                _enqLine["hil_videocallurl"] = resp.Data.MeetingURL;
                            _enqLine["hil_appointmentstatus"] = new OptionSetValue(5);
                            _enqLine["hil_appointmentdate"] = new DateTime(Convert.ToInt32(req.SlotDate.Substring(0, 4)), Convert.ToInt32(req.SlotDate.Substring(4, 2)), Convert.ToInt32(req.SlotDate.Substring(6, 2)), Convert.ToInt32(req.SlotStart.Substring(0, 2)), Convert.ToInt32(req.SlotStart.Substring(3, 2)), 0);
                            _enqLine["hil_appointmentenddate"] = new DateTime(Convert.ToInt32(req.SlotDate.Substring(0, 4)), Convert.ToInt32(req.SlotDate.Substring(4, 2)), Convert.ToInt32(req.SlotDate.Substring(6, 2)), Convert.ToInt32(req.SlotEnd.Substring(0, 2)), Convert.ToInt32(req.SlotEnd.Substring(3, 2)), 0);
                            _enqLine["hil_appointmenttypes"] = req.IsVideoMeeting ? new OptionSetValue(2) : new OptionSetValue(1);
                            service.Update(_enqLine);
                            res.Message = resp.Message;
                            res.Status = true;
                        }
                        else
                        {
                            res.Message = resp.Message;
                            return res;
                        }
                    }
                    else
                    {
                        resp = CreateMeeting(AppointmentID, SlotEnd, SlotStart, EnquirerEmailId, EnquirerName, req.SlotDate.Substring(4, 2) + "/" + req.SlotDate.Substring(6, 2) + "/" + (req.SlotDate.Substring(0, 4)), UserCode, req.IsVideoMeeting, EnquiryId, EnquiryType, AdvisoryType, service);
                        if (resp.IsSuccess)
                        {
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
                                                <condition attribute='hil_homeadvisorylineid' operator='eq' uiname='' uitype='hil_homeadvisoryline' value='" + req.RecordID + @"' />
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
                                service.Update(_app);
                            }
                            Entity _appointment = new Entity("appointment");
                            Entity from = new Entity("activityparty");
                            from["partyid"] = cust;
                            _appointment["requiredattendees"] = new Entity[] { from };
                            _appointment["subject"] = "Meeting with Havells for Advisory";
                            if (req.IsVideoMeeting)
                                _appointment["location"] = "Teams";
                            else
                                _appointment["location"] = "Audio Call";
                            _appointment["regardingobjectid"] = new EntityReference("hil_homeadvisoryline", regarding);

                            DateTime SlotDateStartTime = new DateTime(Convert.ToInt32(req.SlotDate.Substring(0, 4)), Convert.ToInt32(req.SlotDate.Substring(4, 2)), Convert.ToInt32(req.SlotDate.Substring(6, 2)), Convert.ToInt32(req.SlotStart.Substring(0, 2)), Convert.ToInt32(req.SlotStart.Substring(3, 2)), 0);
                            DateTime SlotDateEndTime = new DateTime(Convert.ToInt32(req.SlotDate.Substring(0, 4)), Convert.ToInt32(req.SlotDate.Substring(4, 2)), Convert.ToInt32(req.SlotDate.Substring(6, 2)), Convert.ToInt32(req.SlotEnd.Substring(0, 2)), Convert.ToInt32(req.SlotEnd.Substring(3, 2)), 0);

                            _appointment["scheduledstart"] = SlotDateStartTime;
                            _appointment["scheduledend"] = SlotDateEndTime;
                            if (resp.Data.MeetingURL != null && resp.Data.MeetingURL != string.Empty && resp.Data.MeetingURL != "")
                            {
                                _appointment["hil_appointmenturl"] = resp.Data.MeetingURL;
                            }
                            service.Create(_appointment);

                            Entity _enqLine = new Entity(Found.EntityName);
                            _enqLine.Id = Found.Entities[0].Id;

                            _enqLine["hil_appointmentid"] = resp.Data.TransactionId;
                            if (resp.Data.MeetingURL != "" || resp.Data.MeetingURL != null)
                                _enqLine["hil_videocallurl"] = resp.Data.MeetingURL;
                            _enqLine["hil_appointmentstatus"] = new OptionSetValue(6);
                            _enqLine["hil_appointmentdate"] = SlotDateStartTime;
                            _enqLine["hil_appointmentenddate"] = SlotDateEndTime;
                            _enqLine["hil_appointmenttypes"] = req.IsVideoMeeting ? new OptionSetValue(2) : new OptionSetValue(1);
                            _enqLine["hil_enquirystauts"] = new OptionSetValue(2);
                            service.Update(_enqLine);
                            res.Message = resp.Message;
                            res.Status = true;
                        }
                        else
                        {
                            res.Message = resp.Message;
                            return res;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                res.Message = (ex.Message);
                res.Status = false;
                return res;
            }
            return res;
        }
        CreateUserMeetingResponse CreateMeeting(String AppointmentID, String SlotEnd, String SlotStart, String EnquirerEmailId, String EnquirerName, String SlotDate, String UserCode, bool IsVideo, string EnquiryID, string EnquiryType, string AdvisoryType, IOrganizationService service)
        {
            CreateUserMeetingResponse obj = new CreateUserMeetingResponse();
            try
            {
                Slot slot = new Slot();
                slot.SlotEnd = SlotEnd;
                slot.SlotStart = SlotStart;

                CreateUserMeetingRequest reqParm = new CreateUserMeetingRequest();
                reqParm.EnquirerEmailId = EnquirerEmailId;
                reqParm.EnquirerName = EnquirerName;
                reqParm.SlotDate = SlotDate;
                reqParm.UserCode = int.Parse(UserCode);
                reqParm.Slot = slot;
                reqParm.IsVideoMeeting = IsVideo;
                reqParm.TransactionId = AppointmentID;
                reqParm.EnquiryId = EnquiryID;
                reqParm.EnquiryType = EnquiryType;
                reqParm.AdvisoryType = AdvisoryType;
                Integration integration = IntegrationConfiguration(service, "CreateUserMeeting");
                string _authInfo = integration.Auth;
                string sUrl = integration.uri;
                HttpClient client = new HttpClient();
                var byteArray = Encoding.ASCII.GetBytes(_authInfo);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                string jsonContent = JsonSerializer.Serialize(reqParm);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");
                HttpResponseMessage response = client.PostAsync(sUrl, content).Result;
                string responseContent = response.Content.ReadAsStringAsync().Result;
                obj = JsonSerializer.Deserialize<CreateUserMeetingResponse>(responseContent);

            }
            catch
            {
                obj.Message = "Failed";
            }
            return obj;
        }

        private static Integration IntegrationConfiguration(IOrganizationService service, string Param)
        {
            Integration output = new Integration();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + ":" + integrationConfiguration.GetAttributeValue<string>("hil_password");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:- " + ex.Message);
            }
            return output;
        }
        public class Integration
        {
            public string uri { get; set; }
            public string Auth { get; set; }
        }
        public class CRMRequest
        {
            public string SlotDate { get; set; }
            public string RecordID { get; set; }
            public string SlotStart { get; set; }
            public string SlotEnd { get; set; }
            public bool IsVideoMeeting { get; set; }
        }
        public class CreateUserMeetingResponse
        {
            public Data Data { get; set; }
            public bool IsSuccess { get; set; }
            public string Message { get; set; }
        }
        public class Data
        {

            public string TransactionId { get; set; }

            public string MeetingURL { get; set; }
        }
        public class CreateUserMeetingRequest
        {
            public int UserCode { get; set; }
            public string EnquirerEmailId { get; set; }
            public string EnquirerName { get; set; }
            public string SlotDate { get; set; }
            public Slot Slot { get; set; }
            public bool IsVideoMeeting { get; set; }
            public String TransactionId { get; set; }
            public String EnquiryId { get; set; }
            public String EnquiryType { get; set; }
            public String AdvisoryType { get; set; }
        }
        public class Slot
        {
            public string SlotStart { get; set; }
            public string SlotEnd { get; set; }
        }
    }
}
