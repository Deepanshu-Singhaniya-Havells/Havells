using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Net;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Grievance_Escalation
{
    public class GetComplaints : IPlugin
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

                var _response = GetComplaint(service, CustomerGuid);
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

        public (Complaints_Response, RequestStatus) GetComplaint(IOrganizationService _CrmService, string CustomerGuid)
        {
            Complaints_Response res = new Complaints_Response();
            res.Complaints = new List<Complaints_Complaint>();
            try
            {
                string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='incident'>
                                <attribute name='ticketnumber' />
                                <attribute name='incidentid' />
                                <attribute name='statuscode' />
                                <attribute name='hil_casecategory' />
                                <attribute name='createdon' />
                                <attribute name='hil_job' />
                                <attribute name='title' />
                                <attribute name='hil_reminderremarks' />
                                <attribute name='hil_isreminderset' />
                                <attribute name='hil_caseassignedon' />
                                <attribute name='hil_firstresponsesenton' />
                                <attribute name='adx_resolutiondate' />
                                <attribute name='description' />
                                <attribute name='escalatedon' />
                                <attribute name='hil_reminderdate' />
                                <attribute name='hil_assignmentlevel' />
                                <order attribute='createdon' descending='true' />
                                <filter type='and'>
                                  <condition attribute='customerid' operator='eq' value='{CustomerGuid}' />
                                </filter>
                                <link-entity name='hil_address' from='hil_addressid' to='hil_address' visible='false' link-type='outer' alias='address_detail'>
                                  <attribute name='hil_street2' />
                                  <attribute name='hil_street1' />
                                  <attribute name='hil_state' />
                                  <attribute name='hil_pincode' />
                                  <attribute name='hil_city' />
                                </link-entity>
                              </entity>
                            </fetch>";

                EntityCollection result = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml));

                if (result.Entities.Count > 0)
                {
                    foreach (var entity in result.Entities)
                    {
                        Complaints_Complaint tempComplaint = new Complaints_Complaint();
                        tempComplaint.CompliantGuid = entity.Id.ToString();
                        tempComplaint.ComplaintId = entity.Contains("ticketnumber") ? entity.GetAttributeValue<string>("ticketnumber") : "";

                        int statuscode = entity.GetAttributeValue<OptionSetValue>("statuscode").Value;
                        string Status_code;
                        switch (statuscode)
                        {
                            case 5:
                                Status_code = "Problem Solved";
                                break;
                            case 1000:
                                Status_code = "Information Provided";
                                break;
                            case 4:
                                Status_code = "Researching";
                                break;
                            case 6:
                                Status_code = "Cancelled";
                                break;
                            default:
                                Status_code = "In Progress";
                                break;
                        }
                        tempComplaint.Status = Status_code;

                        int EscalationLevel = entity.Contains("hil_assignmentlevel") ? entity.GetAttributeValue<OptionSetValue>("hil_assignmentlevel").Value : 5; // 5 chosen as to move the switch to the default case
                        string Escalation_Level;
                        switch (EscalationLevel)
                        {
                            case 1:
                                Escalation_Level = "Level 1";
                                break;
                            case 2:
                                Escalation_Level = "Level 2";
                                break;
                            case 3:
                                Escalation_Level = "Level 3";
                                break;
                            case 4:
                                Escalation_Level = "Escalated";
                                break;
                            default:
                                Escalation_Level = "Pending for assignment";
                                break;
                        }
                        tempComplaint.Escalation_Level = Escalation_Level;

                        tempComplaint.ComplaintCategory = entity.Contains("hil_casecategory") ? entity.GetAttributeValue<EntityReference>("hil_casecategory").Name : "";
                        tempComplaint.Created_On = entity.Contains("createdon") ? entity.GetAttributeValue<DateTime>("createdon").Date.ToString("d") : "";
                        tempComplaint.Service_RequestID = entity.Contains("hil_job") ? entity.GetAttributeValue<EntityReference>("hil_job").Name : "";
                        tempComplaint.Title = entity.Contains("title") ? entity.GetAttributeValue<string>("title") : "";
                        tempComplaint.Description = entity.Contains("description") ? entity.GetAttributeValue<string>("description") : "";
                        tempComplaint.ReminderRemarks = entity.Contains("hil_reminderremarks") ? entity.GetAttributeValue<string>("hil_reminderremarks") : "";
                        tempComplaint.IsReminderSet = entity.Contains("hil_isreminderset") ? entity.GetAttributeValue<Boolean>("hil_isreminderset") : false;
                        tempComplaint.Reminder_RaisedOn = entity.Contains("hil_reminderdate") ? entity.GetAttributeValue<DateTime>("hil_reminderdate").AddHours(5.5).ToString() : "";
                        tempComplaint.EscalatedOn = entity.Contains("escalatedon") ? entity.GetAttributeValue<DateTime>("escalatedon").ToString() : "";
                        tempComplaint.AssignedOn = entity.Contains("hil_caseassignedon") ? entity.GetAttributeValue<DateTime>("hil_caseassignedon").ToString() : "";
                        tempComplaint.FirstResponseSentOn = entity.Contains("hil_firstresponsesenton") ? entity.GetAttributeValue<DateTime>("hil_firstresponsesenton").ToString() : "";
                        tempComplaint.ResolvedOn = entity.Contains("adx_resolutiondate") ? entity.GetAttributeValue<DateTime>("adx_resolutiondate").ToString() : "";


                        if (entity.Contains("address_detail.hil_street1"))
                        {
                            tempComplaint.AddressLine1 = (string)((AliasedValue)entity["address_detail.hil_street1"]).Value;
                        }
                        if (entity.Contains("address_detail.hil_street2"))
                        {
                            tempComplaint.AddressLine2 = (string)((AliasedValue)entity["address_detail.hil_street2"]).Value;
                        }
                        if (entity.Contains("address_detail.hil_state"))
                        {
                            tempComplaint.AddressState = ((EntityReference)((AliasedValue)entity["address_detail.hil_state"]).Value).Name;
                        }
                        if (entity.Contains("address_detail.hil_pincode"))
                        {
                            tempComplaint.AddressPinCode = ((EntityReference)((AliasedValue)entity["address_detail.hil_pincode"]).Value).Name;
                        }
                        if (entity.Contains("address_detail.hil_city"))
                        {
                            tempComplaint.AddressCity = ((EntityReference)((AliasedValue)entity["address_detail.hil_city"]).Value).Name;
                        }

                        QueryExpression mediaGallery = new QueryExpression("hil_mediagallery");
                        mediaGallery.ColumnSet = new ColumnSet("hil_url");
                        mediaGallery.Criteria.AddCondition("hil_case", ConditionOperator.Equal, entity.Id);
                        EntityCollection mediaGalleryColl = _CrmService.RetrieveMultiple(mediaGallery);

                        if (mediaGalleryColl.Entities.Count > 0)
                        {
                            tempComplaint.Attachment_URL = mediaGalleryColl.Entities[0].Contains("hil_url") ? mediaGalleryColl.Entities[0].GetAttributeValue<string>("hil_url") : "";
                        }
                        res.Complaints.Add(tempComplaint);
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
    public class Complaints_Complaint
    {
        public string ComplaintId { get; set; }
        public string CompliantGuid { get; set; }
        public string Status { get; set; }
        public string ComplaintCategory { get; set; }
        public string Created_On { get; set; }
        public string Service_RequestID { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public string AddressLine1 { get; set; }
        public string AddressLine2 { get; set; }
        public string AddressCity { get; set; }
        public string AddressState { get; set; }
        public string AddressPinCode { get; set; }
        public string Attachment_URL { get; set; }
        public string ReminderRemarks { get; set; }
        public bool IsReminderSet { get; set; }
        public string Reminder_RaisedOn { get; set; }
        public string EscalatedOn { get; set; }
        public string Escalation_Level { get; set; }
        public string AssignedOn { get; set; }
        public string FirstResponseSentOn { get; set; }
        public string ResolvedOn { get; set; }
    }
    public class Complaints_Response
    {
        public string Error { get; set; }
        public List<Complaints_Complaint> Complaints { get; set; }
    }
}
