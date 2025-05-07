using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Havells.Dataverse.CustomConnector.Home_Advisory
{
    public class HVCancleAppointment : IPlugin
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
                CancelAppointmentRequest cancelAppointmentRequest = new CancelAppointmentRequest();
                string advisorylineGuid = Convert.ToString(context.InputParameters["AdvisorylineGuid"]);
                string advisorylineId = Convert.ToString(context.InputParameters["AdvisorylineId"]);
                string appointmentRemarks = Convert.ToString(context.InputParameters["AppointmentRemarks"]);
                string appointmentStatus = Convert.ToString(context.InputParameters["AppointmentStatus"]);
                bool isEnquiryClosed = Convert.ToBoolean(context.InputParameters["IsEnquiryClosed"]);
                string enquiryRemarks = Convert.ToString(context.InputParameters["EnquiryRemarks"]);
                string enquiryStatus = Convert.ToString(context.InputParameters["EnquiryStatus"]);
                string enquiryCloseReason = Convert.ToString(context.InputParameters["EnquiryCloseReason"]);

                if (string.IsNullOrWhiteSpace(advisorylineGuid) && string.IsNullOrWhiteSpace(advisorylineId))
                {
                    string msg = "Either AdvisorylineGuid OR AdvisorylineId is Required";
                    var paramPesponse = JsonSerializer.Serialize(new Response { Status = false, Message = msg });
                    context.OutputParameters["data"] = paramPesponse;
                    return;
                }

                if (!string.IsNullOrWhiteSpace(advisorylineGuid))
                {
                    if (APValidate.IsvalidGuid(advisorylineGuid))
                    {
                        cancelAppointmentRequest.AdvisorylineGuid = advisorylineGuid;
                    }
                    else
                    {
                        string msg = "AdvisorylineGuid is not valid";
                        var paramPesponse = JsonSerializer.Serialize(new Response { Status = false, Message = msg });
                        context.OutputParameters["data"] = paramPesponse;
                        return;
                    }
                }

                if (!string.IsNullOrWhiteSpace(advisorylineId))
                {
                    string pattern = "^[0-9-_]+$";
                    bool isMatch = Regex.IsMatch(advisorylineId, pattern);
                    if (isMatch)
                    {
                        cancelAppointmentRequest.AdvisorylineId = advisorylineId;
                    }

                    else
                    {
                        string msg = "AdvisorylineId is not valid";
                        var paramPesponse = JsonSerializer.Serialize(new Response { Status = false, Message = msg });
                        context.OutputParameters["data"] = paramPesponse;
                        return;
                    }
                }

                if (!string.IsNullOrEmpty(appointmentRemarks))
                {
                    if (APValidate.isAlphaNumeric(appointmentRemarks) || APValidate.IsValidString(appointmentRemarks))
                    {
                        cancelAppointmentRequest.AppointmentRemarks = appointmentRemarks;
                    }
                }
                else
                {
                    string msg = "Please enter AppointmentRemarks";
                    var paramPesponse = JsonSerializer.Serialize(new Response { Status = false, Message = msg });
                    context.OutputParameters["data"] = paramPesponse;
                    _tracingService.Trace("Validation failed: AppointmentRemarks is required.");
                    return;
                }

                if (!string.IsNullOrEmpty(appointmentStatus))
                {
                    if (APValidate.IsNumeric(appointmentStatus) && appointmentStatus.Length < 9)
                    {
                        cancelAppointmentRequest.AppointmentStatus = appointmentStatus;
                    }
                    else
                    {
                        string msg = "AppointmentStatus is not valid";
                        var paramPesponse = JsonSerializer.Serialize(new Response { Status = false, Message = msg });
                        context.OutputParameters["data"] = paramPesponse;
                        return;
                    }
                }
                else
                {
                    string msg = "Please enter AppointmentStatus";
                    var paramPesponse = JsonSerializer.Serialize(new Response { Status = false, Message = msg });
                    context.OutputParameters["data"] = paramPesponse;
                    _tracingService.Trace("Validation failed: AppointmentStatus is required.");
                    return;
                }

                if (!string.IsNullOrEmpty(enquiryRemarks))
                {
                    if (APValidate.isAlphaNumeric(enquiryRemarks) || APValidate.IsValidString(enquiryRemarks))
                    {
                        cancelAppointmentRequest.EnquiryRemarks = enquiryRemarks;
                    }

                }
                else
                {
                    string msg = "Please enter EnquiryRemarks";
                    var paramPesponse = JsonSerializer.Serialize(new Response { Status = false, Message = msg });
                    context.OutputParameters["data"] = paramPesponse;
                    _tracingService.Trace("Validation failed: EnquiryRemarks is required.");
                    return;
                }

                if (!string.IsNullOrEmpty(enquiryStatus))
                {
                    if (APValidate.IsNumeric(enquiryStatus) && enquiryStatus.Length < 9)
                    {
                        cancelAppointmentRequest.EnquiryStatus = enquiryStatus;
                    }
                    else
                    {
                        string msg = "EnquiryStatus is not valid";
                        var paramPesponse = JsonSerializer.Serialize(new Response { Status = false, Message = msg });
                        context.OutputParameters["data"] = paramPesponse;
                        return;

                    }

                }
                else
                {
                    string msg = "Please enter EnquiryStatus";
                    var paramPesponse = JsonSerializer.Serialize(new Response { Status = false, Message = msg });
                    context.OutputParameters["data"] = paramPesponse;
                    _tracingService.Trace("Validation failed: EnquiryStatus is required.");
                    return;
                }

                if (!string.IsNullOrEmpty(isEnquiryClosed.ToString()))
                {
                    if (APValidate.IsValidboolen(isEnquiryClosed.ToString()))
                    {
                        cancelAppointmentRequest.IsEnquiryClosed = isEnquiryClosed;
                    }
                }
                Response Response = CancleAppointments(cancelAppointmentRequest, service);
                var serializedResponse = JsonSerializer.Serialize(Response);
                context.OutputParameters["data"] = serializedResponse;
            }
            catch (Exception ex)
            {
                _tracingService.Trace("Execute method error: " + ex.ToString());
                throw new InvalidPluginExecutionException("Plugin Error: " + ex.Message);
            }
        }
        public Response CancleAppointments(CancelAppointmentRequest req, IOrganizationService service)
        {
            Response resp = new Response();
            if (!string.IsNullOrWhiteSpace(req.AdvisorylineGuid))
            {
                resp = CancleTeamAppointment(new Guid(req.AdvisorylineGuid), req.AppointmentStatus, req.AppointmentRemarks, service);
                return resp;
            }
            else
            {
                try
                {
                    string fetchAdvisoryline = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='hil_homeadvisoryline'>
                    <attribute name='hil_homeadvisorylineid' />
                    <attribute name='hil_name' />
                    <attribute name='createdon' />
                    <order attribute='hil_name' descending='false' />
                    <filter type='and'>
                    <condition attribute='statecode' operator='eq' value='0' />
                    <condition attribute='hil_name' operator='eq' value='" + req.AdvisorylineId.Trim() + @"' />
                    </filter>
                    </entity>
                    </fetch>";
                    EntityCollection AdvisoryLineColl = service.RetrieveMultiple(new FetchExpression(fetchAdvisoryline));
                    if (AdvisoryLineColl.Entities.Count == 0)
                    {
                        resp.Message = "Advisory Enquery Line Not Found in Dynamics";
                        resp.Status = false;
                        return resp;
                    }
                    Entity Advisoryline = AdvisoryLineColl.Entities[0];
                    if (req.IsEnquiryClosed)
                    {
                        EntityReference cancleReAON = new EntityReference();
                        QueryExpression query = new QueryExpression("hil_cancellationreason");
                        String _fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                            <entity name='hil_cancellationreason'>
                                            <attribute name='hil_cancellationreasonid' />
                                            <attribute name='hil_name' />
                                            <order attribute='hil_name' descending='false' />
                                            <filter type='and'>
                                            <condition attribute='hil_name' operator='eq' value='" + req.EnquiryRemarks + @"' />
                                            </filter>
                                           </entity>
                                            </fetch>";
                        EntityCollection reasonColl = service.RetrieveMultiple(new FetchExpression(_fetch));
                        if (reasonColl.Entities.Count > 0)
                        {
                            cancleReAON = new EntityReference(reasonColl.EntityName, reasonColl[0].Id);
                        }
                        else
                        {
                            resp.Message = "Cancelation Reasion Not found in Dynamics.";
                            resp.Status = false;
                            return resp;
                        }
                        Entity _advLine = new Entity(Advisoryline.LogicalName);
                        _advLine.Id = Advisoryline.Id;
                        _advLine["hil_enquirystauts"] = new OptionSetValue(Convert.ToInt32(req.EnquiryStatus));
                        _advLine["hil_closingreasion"] = new EntityReference(cancleReAON.LogicalName, cancleReAON.Id);
                        bool _apSt = CloseAppointment(Advisoryline.Id.ToString(), req.AppointmentStatus, req.AppointmentRemarks, service);
                        if (_apSt)
                        {
                            _advLine["hil_appointmentstatus"] = new OptionSetValue(3);
                            service.Update(_advLine);
                            resp.Message = "Enquery and Appointment cancled Sucessfully";
                            resp.Status = true;
                            return resp;
                        }
                        else
                        {
                            service.Update(_advLine);
                            resp.Message = "Enquery Closed Sucessfully";
                            resp.Status = true;
                            return resp;
                        }
                    }
                    else
                    {
                        bool _apSt = CloseAppointment(Advisoryline.Id.ToString(), req.AppointmentStatus, req.AppointmentRemarks, service);
                        if (_apSt)
                        {
                            Entity _advLine = new Entity(Advisoryline.LogicalName);
                            _advLine.Id = Advisoryline.Id;
                            _advLine["hil_appointmentstatus"] = new OptionSetValue(4); // 3- Completed, 4- Cancel
                            service.Update(_advLine);
                            resp.Message = "Appointment Closed Sucessfully";
                            resp.Status = true;
                            return resp;
                        }
                        else
                        {
                            resp.Message = "Appointment Not Closed ";
                            resp.Status = false;
                            return resp;
                        }
                    }
                }
                catch (Exception ex)
                {
                    resp.Message = "D365 Internal Server Error: " + ex.Message;
                    resp.Status = false;
                    return resp;
                }
            }
        }
        public Response CancleTeamAppointment(Guid req, string AppoitmentStatus, string Remarks, IOrganizationService service)
        {
            Response resp = new Response();

            try
            {
                CancelEvent obj = new CancelEvent();
                QueryExpression Query = new QueryExpression("hil_homeadvisoryline");
                Query.ColumnSet = new ColumnSet("hil_appointmentid");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_homeadvisorylineid", ConditionOperator.Equal, req);
                EntityCollection _entitys = service.RetrieveMultiple(Query);
                if (_entitys.Entities.Count == 1)
                {
                    if (_entitys.Entities[0].Contains("hil_appointmentid"))
                    {
                        string trancId = _entitys.Entities[0].GetAttributeValue<String>("hil_appointmentid");

                        Integration integration = IntegrationConfiguration(service, "CancelEvent");
                        string _authInfo = integration.Auth;
                        _authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));
                        String sUrl = integration.uri;

                        using (var client = new HttpClient())
                        {
                            client.DefaultRequestHeaders.Add("Authorization", _authInfo);
                            client.DefaultRequestHeaders.Add("Content-Type", "application/json");
                            var requestBody = new StringContent(
                                "{\r\n    \"TransactionId\": \"" + trancId + "\"\r\n}", Encoding.UTF8, "application/json");
                            HttpResponseMessage response = client.PostAsync(sUrl, requestBody).Result;
                            string content = response.Content.ReadAsStringAsync().Result;
                            obj = Newtonsoft.Json.JsonConvert.DeserializeObject<CancelEvent>(content);
                        }
                        if (obj.IsSuccess)
                        {
                            Entity _advLine = new Entity("hil_homeadvisoryline");
                            _advLine.Id = _entitys.Entities[0].Id;
                            _advLine["hil_appointmentid"] = String.Empty;
                            _advLine["hil_videocallurl"] = String.Empty;
                            _advLine["hil_appointmentstatus"] = new OptionSetValue(4);

                            service.Update(_advLine);

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
                                                  "<condition attribute=\"regardingobjectid\" operator=\"eq\" uiname=\"\" uitype=\"hil_homeadvisoryline\" value=\"" + _entitys.Entities[0].Id + "\" />" +
                                                  "<condition attribute=\"hil_appointmenturl\" operator=\"not-null\" />" +
                                                  "<condition attribute=\"statecode\" operator=\"in\">" +
                                                    "<value>0</value>" +
                                                     "<value>3</value>" +
                                                    "</condition>" +
                                                  "</filter>" +
                                                "</entity>" +
                                              "</fetch>";
                            EntityCollection _apps = service.RetrieveMultiple(new FetchExpression(fetct));
                            var i = _apps.Entities.Count;
                            Entity _app = new Entity(_apps.EntityName);
                            _app.Id = _apps.Entities[0].Id;
                            _app["statecode"] = new OptionSetValue(2);
                            _app["statuscode"] = new OptionSetValue(4);

                            service.Update(_app);
                            resp.Message = obj.Message;

                        }
                        else
                        {
                            resp.Message = ("Invalid Transaction Id");
                        }
                    }
                    else
                    {
                        resp.Message = ("Meeting Not found for this Advisory");
                    }

                }
                else
                {
                    resp.Message = ("Record Not Found");
                }
            }
            catch (Exception ex)
            {
                resp.Message = ex.Message;
            }
            return resp;
        }
        private bool CloseAppointment(string _enqueryLine, string AppoitmentStatus, string Remarks, IOrganizationService service)
        {
            try
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
                            <condition attribute='hil_homeadvisorylineid' operator='eq' uiname='' uitype='hil_homeadvisoryline' value='" + _enqueryLine + @"' />
                            </filter>
                        </link-entity>
                        </entity>
                    </fetch>";
                EntityCollection AdvlineAppointmentColl = service.RetrieveMultiple(new FetchExpression(fetchappointment));
                if (AdvlineAppointmentColl.Entities.Count > 0)
                {

                    Entity _app = new Entity("appointment");
                    _app.Id = AdvlineAppointmentColl.Entities[0].Id;
                    _app["statecode"] = AppoitmentStatus == "3" ? new OptionSetValue(1) : new OptionSetValue(2);
                    _app["statuscode"] = AppoitmentStatus == "3" ? new OptionSetValue(3) : new OptionSetValue(4);
                    _app["description"] = Remarks != null ? Remarks : "";
                    service.Update(_app);
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch
            {
                return false;
            }
        }
        static Integration IntegrationConfiguration(IOrganizationService service, string Param)
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
    }
    public class Integration
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }
    public class CancelAppointmentRequest
    {
        public string AdvisorylineGuid { get; set; }
        public string AdvisorylineId { get; set; }
        public string AppointmentRemarks { get; set; }
        public string AppointmentStatus { get; set; }
        public bool IsEnquiryClosed { get; set; }
        public string EnquiryRemarks { get; set; }
        public string EnquiryStatus { get; set; }
        public string EnquiryCloseReason { get; set; }

    }
    public class Response
    {
        public string Message { get; set; }
        public bool Status { get; set; }
    }
    public class CancelEvent
    {
        public bool IsSuccess { get; set; }
        public string Message { get; set; }
    }
}

