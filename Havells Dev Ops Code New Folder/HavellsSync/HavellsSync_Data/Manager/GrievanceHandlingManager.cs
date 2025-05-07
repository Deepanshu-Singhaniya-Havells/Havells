using HavellsSync_Data.IManager;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.GrievanceHandling;
using Microsoft.Extensions.Configuration;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System.Net;
using System;
using HavellsSync_ModelData.ServiceAlaCarte;
using Azure.Storage.Blobs;
using System.Security.Policy;

namespace HavellsSync_Data.Manager
{
    public class GrievanceHandlingManager : IGrievanceHandlingManager
    {
        private IConfiguration configuration;
        private ICrmService _CrmService;

        public GrievanceHandlingManager(ICrmService crmService, IConfiguration configuration)
        {
            Check.Argument.IsNotNull(nameof(crmService), crmService);
            _CrmService = crmService;
            this.configuration = configuration;

        }
        public async Task<(Response, RequestStatus)> GetComplaintCategory()
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
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }
        public async Task<(OpenJobsResponse, RequestStatus)> GetOpenJobs(OpenJobsRequest obj)
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
                                  <condition attribute='hil_customerref' operator='eq' value='{obj.CustomerGuid}' />
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
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }
        public async Task<(OpenComplaints_Response, RequestStatus)> GetOpenComplaints(OpenComplaints_Request obj)
        {
            OpenComplaints_Response res = new OpenComplaints_Response();
            res.Complaints = new List<OpenComplaints_Complaint>();
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
                                  <condition attribute='customerid' operator='eq' value='{obj.CustomerGuid}' />
                                  <condition attribute='statuscode' operator='eq' value='1' />
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
                        OpenComplaints_Complaint tempComplaint = new OpenComplaints_Complaint();
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

                        int EscalationLevel = entity.Contains("hil_assignmentlevel") ? entity.GetAttributeValue<OptionSetValue>("hil_assignmentlevel").Value : 5; // 5 chosen as to move the switch case to the default case
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
                        mediaGallery.Criteria.AddCondition("cr991_case", ConditionOperator.Equal, entity.Id);   // to be changed to hil_case for Production

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
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }
        public async Task<(Complaints_Response, RequestStatus)> GetComplaints(Complaints_Request obj)
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
                                  <condition attribute='customerid' operator='eq' value='{obj.CustomerGuid}' />
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
                        mediaGallery.Criteria.AddCondition("cr991_case", ConditionOperator.Equal, entity.Id);
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
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }
        public async Task<(RaiseReminder_Response, RequestStatus)> RaiseReminder(RaiseReminder_Request obj)
        {

            RaiseReminder_Response res = new RaiseReminder_Response();
            if (string.IsNullOrEmpty(obj.ComplaintGuid))
            {
                res.Message = "complaint guid cannot be blank";
                res.Status = false;
                return (res, new RequestStatus());
            }
            Entity incident = _CrmService.Retrieve("incident", new Guid(obj.ComplaintGuid), new ColumnSet("hil_isreminderset", "hil_reminderremarks", "hil_reminderdate"));
            try
            {
                incident["hil_isreminderset"] = true;
                incident["hil_reminderremarks"] = obj.ReminderRemarks;
                incident["hil_reminderdate"] = DateTime.Now.AddMinutes(330);
                _CrmService.Update(incident);

                res.Status = true;
                res.Message = "reminder raised successfully";
                return (res, new RequestStatus());
            }
            catch (Exception ex)
            {
                res.Status = false;
                res.Message = "error: " + ex.Message;
                return (res, new RequestStatus());
            }
        }
        public async Task<(CreateCase_Response, RequestStatus)> CreateCase(CreateCase_Request obj)
        {

            CreateCase_Response res = new CreateCase_Response();
            res.ServiceRequests = new List<CreateCase_ServiceRequest>();
            try
            {

                OptionSetValue CaseType;
                EntityReference CaseDepartment;
                EntityReference Branch;
                EntityReference Product;

                if (obj.ComplaintCategoryId != null)
                {
                    Entity caseCategory = _CrmService.Retrieve("hil_casecategory", new Guid(obj.ComplaintCategoryId), new ColumnSet("hil_casedepartment", "hil_casetype"));
                    CaseType = caseCategory.GetAttributeValue<OptionSetValue>("hil_casetype");
                    CaseDepartment = caseCategory.GetAttributeValue<EntityReference>("hil_casedepartment");
                }
                else
                {

                    return (res, new RequestStatus());
                }
                if (string.IsNullOrEmpty(obj.ServiceRequestGuid) && CaseDepartment.Id.ToString() == "ab3dbc3d-4e6e-ee11-8179-6045bdac526a") // service
                {
                    return (res, new RequestStatus
                    {
                        StatusCode = (int)HttpStatusCode.NotFound,
                        Message = "ServiceRequestID is Required when Case Department is 'Service'."
                    });
                }
                if (!string.IsNullOrEmpty(obj.ServiceRequestGuid))
                {
                    Entity Jobs = _CrmService.Retrieve("msdyn_workorder", new Guid(obj.ServiceRequestGuid), new ColumnSet("hil_branch", "hil_productcategory"));
                    Branch = Jobs.GetAttributeValue<EntityReference>("hil_branch");
                    Product = Jobs.GetAttributeValue<EntityReference>("hil_productcategory");
                }
                else
                {
                    if (!string.IsNullOrEmpty(obj.AddressGuid))
                    {
                        Entity Address = _CrmService.Retrieve("hil_address", new Guid(obj.AddressGuid), new ColumnSet("hil_branch"));
                        Branch = Address.GetAttributeValue<EntityReference>("hil_branch");
                        Product = null;
                    }
                    else
                    {

                        return (res, new RequestStatus());
                    }
                }

                var complaintEntity = new Entity("incident");

                complaintEntity["customerid"] = new EntityReference("contact", new Guid(obj.CustomerGuid));
                complaintEntity["hil_address"] = new EntityReference("hil_address", new Guid(obj.AddressGuid));
                if (!string.IsNullOrEmpty(obj.ServiceRequestGuid))
                {
                    complaintEntity["hil_job"] = new EntityReference("msdyn_workorder", new Guid(obj.ServiceRequestGuid));
                }
                complaintEntity["hil_casecategory"] = new EntityReference("hil_casecategory", new Guid(obj.ComplaintCategoryId));
                complaintEntity["title"] = obj.ComplaintTitle;
                complaintEntity["description"] = obj.ComplaintDescription;
                complaintEntity["caseorigincode"] = new OptionSetValue(2); // Havells One Website
                complaintEntity["casetypecode"] = new OptionSetValue(CaseType.Value);
                complaintEntity["productid"] = Product;
                complaintEntity["hil_branch"] = Branch;
                complaintEntity["hil_casedepartment"] = CaseDepartment; // _CrmService
                Guid complaintId = _CrmService.Create(complaintEntity);

                Entity Case = _CrmService.Retrieve("incident", complaintId, new ColumnSet("ticketnumber"));
                string ComplaintGuid = Case.GetAttributeValue<string>("ticketnumber");

                string url = "";
                // to get the url from the Azure connection file
                if (!string.IsNullOrEmpty(obj.Attachment) && !string.IsNullOrEmpty(obj.FileName))
                {
                     url = await StoreFileInAzureBlob(obj.Attachment, obj.FileName, complaintId.ToString());

                    if (!string.IsNullOrEmpty(url))
                    {
                        var MediaGallery = new Entity("hil_mediagallery");
                        MediaGallery["hil_name"] = complaintId.ToString() + "_Attachment";
                        MediaGallery["hil_mediatype"] = new EntityReference("hil_mediatype", new Guid("b787ea2b-688c-ef11-8a6a-7c1e523d7332")); // Grievance Attachment
                        MediaGallery["hil_consumer"] = new EntityReference("contact", new Guid(obj.CustomerGuid));
                        MediaGallery["hil_url"] = url;
                        if (!string.IsNullOrEmpty(obj.ServiceRequestGuid))
                        {
                            MediaGallery["hil_job"] = new EntityReference("contact", new Guid(obj.ServiceRequestGuid));
                        }
                        MediaGallery["cr991_case"] = new EntityReference("contact", complaintId);  // have to change the name for Production to  - hil_cases
                        
                        Guid mediaGalleryGuid = _CrmService.Create(MediaGallery);

                    }
                }
                CreateCase_ServiceRequest tempComplaint = new CreateCase_ServiceRequest();
                tempComplaint.ComplaintId = ComplaintGuid;
                tempComplaint.ComplaintGuid = complaintId.ToString();
                tempComplaint.Attachment_URL = url;
                res.ServiceRequests.Add(tempComplaint);

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
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });

            }

        }
        public async Task<string> StoreFileInAzureBlob(string fileBase64, string fileName, string documentName)
        {

            const string connectionString = "DefaultEndpointsProtocol=https;AccountName=d365storagesa;AccountKey=6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==;EndpointSuffix=core.windows.net";

            string extension = Path.GetExtension(fileName);

            BlobServiceClient blobServiceClient = new BlobServiceClient(connectionString);

            string container = "images";

            // Create the container and return a container client object
            BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(container);

            // Get a reference to a blob
            BlobClient blobClient = containerClient.GetBlobClient(documentName + extension);

            // Convert base64 string to byte array
            byte[] imageBytes = Convert.FromBase64String(fileBase64);

            // Create a memory stream from the byte array
            using MemoryStream uploadStream = new MemoryStream(imageBytes);

            // Upload the stream to the blob
            await blobClient.UploadAsync(uploadStream, true);



            return blobClient.Uri.ToString();

        }
    }
}
