using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Havells.Dataverse.CustomConnector.Service_Call
{
    public class GetJobs : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            StringBuilder errorMessage = new StringBuilder();
            Regex regexDate = new Regex(@"^\d{4}\-(0[1-9]|1[012])\-(0[1-9]|[12][0-9]|3[01])$");
            bool IsValidRequest = true;
            string JsonResponse = "";
            string MobileNumber = Convert.ToString(context.InputParameters["MobileNumber"]);
            string DealerCode = Convert.ToString(context.InputParameters["DealerCode"]);
            string Job_ID = Convert.ToString(context.InputParameters["Job_ID"]);
            string FromDate = Convert.ToString(context.InputParameters["FromDate"]);
            string ToDate = Convert.ToString(context.InputParameters["ToDate"]);
            string SourceOfJob = Convert.ToString(context.InputParameters["SourceOfJob"]);

            if (string.IsNullOrWhiteSpace(DealerCode) && string.IsNullOrWhiteSpace(Job_ID) && string.IsNullOrWhiteSpace(MobileNumber))
            {
                errorMessage.AppendLine("Invalid Request");
                IsValidRequest = false;
            }
            else if (!string.IsNullOrWhiteSpace(DealerCode) || !string.IsNullOrWhiteSpace(FromDate) || !string.IsNullOrWhiteSpace(ToDate))
            {
                if (!string.IsNullOrWhiteSpace(DealerCode))
                {
                    if (!APValidate.isAlphaNumeric(DealerCode))
                    {
                        errorMessage.AppendLine("Invalid Dealer Code.");
                        IsValidRequest = false;
                    }
                    if (DealerCode.Length < 5 || DealerCode.Length > 10)
                    {
                        errorMessage.AppendLine("Dealer code length should be between 5 to 10.");
                        IsValidRequest = false;
                    }
                    if (string.IsNullOrWhiteSpace(FromDate))
                    {
                        errorMessage.AppendLine("From date(Format:yyyy-MM-dd) is required.");
                        IsValidRequest = false;
                    }
                    else if (!regexDate.IsMatch(FromDate))
                    {
                        errorMessage.AppendLine("Invalid date format. It should be (yyyy-MM-dd)");
                        IsValidRequest = false;
                    }
                    if (string.IsNullOrWhiteSpace(ToDate))
                    {
                        errorMessage.AppendLine("To date(Format:yyyy-MM-dd) is required.");
                        IsValidRequest = false;
                    }
                    else if (!regexDate.IsMatch(ToDate))
                    {
                        errorMessage.AppendLine("Invalid date format. It should be (yyyy-MM-dd)");
                        IsValidRequest = false;
                    }
                    DateTime dtFrom;
                    DateTime dtToDate;
                    DateTime.TryParseExact(FromDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dtFrom);
                    DateTime.TryParseExact(ToDate, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dtToDate);
                    string DateValidationMessage = APValidate.ValidateTwoDates(Convert.ToDateTime(dtToDate), Convert.ToDateTime(dtFrom));
                    if (!string.IsNullOrWhiteSpace(DateValidationMessage))
                    {
                        errorMessage.AppendLine(DateValidationMessage);
                        IsValidRequest = false;
                    }
                }
                else
                {
                    errorMessage.AppendLine("Dealer code is required.");
                    IsValidRequest = false;
                }
                if (!IsValidRequest)
                {
                    JsonResponse = JsonSerializer.Serialize(new
                    {
                        StatusCode = 204,
                        Message = errorMessage.ToString(),
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
            }
            else if (!string.IsNullOrWhiteSpace(Job_ID))
            {
                if (Job_ID.Length > 16)
                {
                    errorMessage.AppendLine("Invalid Job Id");
                    IsValidRequest = false;
                }
                if (!APValidate.NumericValue(Job_ID))
                {
                    errorMessage.AppendLine("Invalid Job Id");
                    IsValidRequest = false;
                }
                if (!string.IsNullOrWhiteSpace(MobileNumber))
                {
                    if (!APValidate.IsValidMobileNumber(MobileNumber))
                    {
                        errorMessage.AppendLine("Invalid Mobile Number");
                        IsValidRequest = false;
                    }
                }
            }
            else if (!string.IsNullOrWhiteSpace(MobileNumber))
            {
                if (!APValidate.IsValidMobileNumber(MobileNumber))
                {
                    errorMessage.AppendLine("Invalid Mobile Number");
                    IsValidRequest = false;
                }
            }
            if (!IsValidRequest)
            {
                JsonResponse = JsonSerializer.Serialize(new
                {
                    status_code = 204,
                    status_description = errorMessage.ToString(),
                });
                context.OutputParameters["data"] = JsonResponse;
                return;
            }
            if (!string.IsNullOrEmpty(DealerCode) && !string.IsNullOrEmpty(FromDate) && !string.IsNullOrEmpty(ToDate))
            {
                WorkOrderRequest objOrderReq = new WorkOrderRequest();
                objOrderReq.DealerCode = DealerCode;
                objOrderReq.FromDate = FromDate;
                objOrderReq.ToDate = ToDate;
                JsonResponse = JsonSerializer.Serialize(GetWorkOrdersStatus(service, objOrderReq));
                context.OutputParameters["data"] = JsonResponse;
            }
            else if (!string.IsNullOrWhiteSpace(Job_ID) || !string.IsNullOrWhiteSpace(MobileNumber))
            {
                if (!string.IsNullOrWhiteSpace(MobileNumber) || SourceOfJob == "9")
                {
                    if (MobileNumber.Length > 10)
                    {
                        MobileNumber = MobileNumber.Substring(MobileNumber.Length - 10, 10);
                    }
                    var result = GetJobData(service, Job_ID, MobileNumber);
                    if (result.Count == 0)
                    {
                        JsonResponse = JsonSerializer.Serialize(new List<JobOutput>()
                        {
                            //status_code = 204,
                            //status_description = "No Record Found.",
                        });
                    }
                    else
                    {
                        JsonResponse = JsonSerializer.Serialize(result);
                    }
                    context.OutputParameters["data"] = JsonResponse;
                }
                else
                {
                    JobStatusDTO jobRequest = new JobStatusDTO();
                    jobRequest.job_id = Job_ID;
                    JsonResponse = JsonSerializer.Serialize(GetJobstatus(jobRequest, service));
                    context.OutputParameters["data"] = JsonResponse;
                }
            }
        }
        public List<JobOutput> GetJobData(IOrganizationService service, string Job_ID, string MobileNumber)
        {
            List<JobOutput> jobList = new List<JobOutput>();
            QueryExpression query = new QueryExpression()
            {
                EntityName = "msdyn_workorder",
                ColumnSet = new ColumnSet(true)
            };
            FilterExpression filterExpression = new FilterExpression(LogicalOperator.And);
            if (!String.IsNullOrWhiteSpace(Job_ID))
            {
                filterExpression.Conditions.Add(new ConditionExpression("msdyn_name", ConditionOperator.Equal, Job_ID));
            }
            if (!String.IsNullOrWhiteSpace(MobileNumber))
            {
                filterExpression.Conditions.Add(new ConditionExpression("hil_mobilenumber", ConditionOperator.Equal, MobileNumber));
            }
            query.Criteria.AddFilter(filterExpression);
            EntityCollection collection = service.RetrieveMultiple(query);

            if (collection.Entities != null && collection.Entities.Count > 0)
            {
                foreach (Entity item in collection.Entities)
                {
                    JobOutput jobObj = new JobOutput();
                    if (item.Attributes.Contains("msdyn_customerasset"))
                    {
                        jobObj.Job_Asset = item.GetAttributeValue<EntityReference>("msdyn_customerasset").Name;
                    }
                    if (item.Attributes.Contains("msdyn_name"))
                    {
                        jobObj.Job_ID = item.GetAttributeValue<string>("msdyn_name");
                    }
                    if (item.Attributes.Contains("msdyn_substatus"))
                    {
                        jobObj.Job_Status = item.GetAttributeValue<EntityReference>("msdyn_substatus").Name;
                    }
                    if (item.Attributes.Contains("hil_productcategory"))
                    {
                        jobObj.Job_Category = item.GetAttributeValue<EntityReference>("hil_productcategory").Name;
                    }
                    if (item.Attributes.Contains("createdon"))
                    {
                        jobObj.Job_Loggedon = item.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString();
                    }
                    if (item.Attributes.Contains("hil_jobclosuredon"))
                    {
                        jobObj.Job_ClosedOn = item.GetAttributeValue<DateTime>("hil_jobclosuredon").AddMinutes(330).ToString();
                    }
                    if (item.Attributes.Contains("hil_mobilenumber"))
                    {
                        jobObj.MobileNumber = item.GetAttributeValue<string>("hil_mobilenumber");
                    }
                    if (item.Attributes.Contains("hil_fulladdress"))
                    {
                        jobObj.Customer_Address = item.GetAttributeValue<string>("hil_fulladdress");
                    }
                    if (item.Attributes.Contains("hil_customerref"))
                    {
                        jobObj.Customer_name = item.GetAttributeValue<EntityReference>("hil_customerref").Name;
                    }
                    if (item.Attributes.Contains("hil_owneraccount"))
                    {
                        jobObj.Job_AssignedTo = item.GetAttributeValue<EntityReference>("hil_owneraccount").Name;
                    }
                    if (item.Attributes.Contains("hil_callsubtype"))
                    {
                        jobObj.CallSubType = item.GetAttributeValue<EntityReference>("hil_callsubtype").Name;
                    }
                    if (item.Attributes.Contains("hil_customercomplaintdescription"))
                    {
                        jobObj.ChiefComplaint = item.GetAttributeValue<string>("hil_customercomplaintdescription");
                    }
                    if (item.Attributes.Contains("msdyn_customerasset"))
                    {
                        Entity ec = service.Retrieve("msdyn_customerasset", item.GetAttributeValue<EntityReference>("msdyn_customerasset").Id, new ColumnSet("hil_modelname"));
                        if (ec != null)
                        {
                            jobObj.Product = ec.GetAttributeValue<string>("hil_modelname");
                        }
                    }
                    if (item.Attributes.Contains("hil_productcategory"))
                    {
                        jobObj.ProductCategoryGuid = item.GetAttributeValue<EntityReference>("hil_productcategory").Id;
                        jobObj.ProductCategoryName = item.GetAttributeValue<EntityReference>("hil_productcategory").Name;
                    }
                    jobList.Add(jobObj);
                }
            }
            return jobList;
        }
        public JobStatusDTO GetJobstatus(JobStatusDTO _jobRequest, IOrganizationService service)
        {
            try
            {
                string _fetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorder'>
                        <attribute name='msdyn_name' />
                        <attribute name='createdon' />
                        <attribute name='hil_productcategory' />
                        <attribute name='hil_productsubcategory' />
                        <attribute name='hil_customerref' />
                        <attribute name='hil_callsubtype' />
                        <attribute name='msdyn_workorderid' />
                        <attribute name='hil_webclosureremarks' />
                        <attribute name='hil_typeofassignee' />
                        <attribute name='msdyn_substatus' />
                        <attribute name='hil_productcategory' />
                        <attribute name='ownerid' />
                        <attribute name='hil_mobilenumber' />
                        <attribute name='hil_jobcancelreason' />
                        <attribute name='hil_customercomplaintdescription' />
                        <attribute name='hil_closureremarks' />
                        <attribute name='msdyn_timeclosed' />
                        <attribute name='msdyn_customerasset' />
                        <order attribute='msdyn_name' descending='false' />
                        <filter type='and'>
                            <condition attribute='msdyn_name' operator='eq' value='{_jobRequest.job_id}' />
                        </filter>
                        </entity>
                        </fetch>";

                EntityCollection _jobDetailsColl = service.RetrieveMultiple(new FetchExpression(_fetchQuery));
                if (_jobDetailsColl.Entities.Count > 0)
                {
                    foreach (Entity entity in _jobDetailsColl.Entities)
                    {
                        _jobRequest.mobile_number = entity.Contains("hil_mobilenumber") ? entity.GetAttributeValue<string>("hil_mobilenumber") : "";
                        _jobRequest.job_id = entity.Contains("msdyn_name") ? entity.GetAttributeValue<string>("msdyn_name") : "";
                        _jobRequest.serial_number = entity.Contains("msdyn_customerasset") ? entity.GetAttributeValue<EntityReference>("msdyn_customerasset").Name : "";
                        _jobRequest.product_category = entity.Contains("hil_productcategory") ? entity.GetAttributeValue<EntityReference>("hil_productcategory").Name : "";
                        _jobRequest.product_subcategory = entity.Contains("hil_productsubcategory") ? entity.GetAttributeValue<EntityReference>("hil_productsubcategory").Name : "";
                        _jobRequest.call_type = entity.Contains("hil_callsubtype") ? entity.GetAttributeValue<EntityReference>("hil_callsubtype").Name : "";
                        _jobRequest.customer_complaint = entity.Contains("hil_customercomplaintdescription") ? entity.GetAttributeValue<string>("hil_customercomplaintdescription") : "";
                        _jobRequest.assigned_resource = entity.Contains("ownerid") ? entity.GetAttributeValue<EntityReference>("ownerid").Name : "";
                        _jobRequest.job_substatus = entity.Contains("msdyn_substatus") ? entity.GetAttributeValue<EntityReference>("msdyn_substatus").Name : "";


                        if (entity.Contains("msdyn_timeclosed"))
                        {
                            _jobRequest.closed_on = entity.Contains("msdyn_timeclosed") ? entity.GetAttributeValue<DateTime>("msdyn_timeclosed").AddMinutes(330).ToString("yyyy-MM-dd hh:mm:ss tt") : "";
                        }
                        if (entity.Attributes.Contains("hil_jobcancelreason"))
                        {
                            if (entity.FormattedValues.Contains("hil_jobcancelreason"))
                                _jobRequest.cancel_reason = entity.FormattedValues["hil_jobcancelreason"];
                        }
                        _jobRequest.webclosure_remarks = entity.Contains("hil_webclosureremarks") ? entity.GetAttributeValue<string>("hil_webclosureremarks") : "";
                        _jobRequest.closure_remarks = entity.Contains("hil_closureremarks") ? entity.GetAttributeValue<string>("hil_closureremarks") : "";
                        string IncFetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='msdyn_workorderproduct'>
                                <attribute name='msdyn_product' />
                                <filter type='and'>
                                <condition attribute='msdyn_workorder' operator='eq' value='{entity.Id}' />
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_markused' operator='eq' value='1' />
                                </filter>
                                <link-entity name='product' from='productid' to='hil_replacedpart' visible='false' link-type='outer' alias='pm'>
                                    <attribute name='description' />
                                </link-entity>
                                <link-entity name='msdyn_workorderincident' from='msdyn_workorderincidentid' to='msdyn_workorderincident' visible='false' link-type='outer' alias='wi'>
                                    <attribute name='msdyn_description' />
                                </link-entity>
                                </entity>
                                </fetch>";
                        EntityCollection JobIncDetailsColl = service.RetrieveMultiple(new FetchExpression(IncFetchQuery));
                        StringBuilder _sparePart = new StringBuilder();

                        _jobRequest.spare_parts = new List<JobProductDTO>();
                        if (JobIncDetailsColl.Entities.Count > 0)
                        {
                            _jobRequest.technician_remarks = JobIncDetailsColl.Entities[0].Contains("wi.msdyn_description") ? JobIncDetailsColl.Entities[0].GetAttributeValue<AliasedValue>("wi.msdyn_description").Value.ToString() : "";
                            int i = 1;
                            foreach (Entity ent in JobIncDetailsColl.Entities)
                            {
                                _jobRequest.spare_parts.Add(new JobProductDTO()
                                {
                                    index = i++.ToString(),
                                    product_code = ent.Contains("msdyn_product") ? ent.GetAttributeValue<EntityReference>("msdyn_product").Name : "",
                                    product_description = ent.Contains("pm.description") ? ent.GetAttributeValue<AliasedValue>("pm.description").Value.ToString() : ""
                                });
                            }
                        }
                    }
                    _jobRequest.status_code = "200";
                    _jobRequest.status_description = "Sucess";
                }
                else
                {
                    _jobRequest.status_code = "204";
                    _jobRequest.status_description = "Invalid Job Id.";

                }
            }
            catch (Exception ex)
            {
                _jobRequest.status_code = "500";
                _jobRequest.status_description = "D365 Internal Server Error : " + ex.Message.ToUpper();

            }

            return _jobRequest;
        }
        public WorkOrderResponse GetWorkOrdersStatus(IOrganizationService service, WorkOrderRequest objreq)
        {
            WorkOrderResponse workOrderResponse = new WorkOrderResponse();
            try
            {
                string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='msdyn_workorder'>
                            <attribute name='msdyn_name' />
                            <attribute name='createdon' />
                            <attribute name='hil_productsubcategory' />
                            <attribute name='hil_customerref' />
                            <attribute name='hil_callsubtype' />
                            <attribute name='msdyn_workorderid' />
                            <attribute name='hil_productcategory' />
                            <attribute name='msdyn_substatus' />
                            <attribute name='hil_owneraccount' />
                            <order attribute='msdyn_name' descending='false' />
                            <filter type='and'>
                                <condition attribute='hil_newserialnumber' operator='not-null' />
                                <condition attribute='createdon' operator='on-or-after' value='{objreq.FromDate}' />
                                <condition attribute='hil_sourceofjob' operator='eq' value='6' />
                                <condition attribute='hil_newserialnumber' operator='like' value='%{objreq.DealerCode}%' />
                                <condition attribute='createdon' operator='on-or-before' value='{objreq.ToDate}' />
                                <condition attribute='createdon' operator='last-x-days' value='60' />
                            </filter>
                            </entity>
                            </fetch>";

                EntityCollection results = service.RetrieveMultiple(new FetchExpression(fetchXml));
                List<WorkOrderInfo> workOrderInfos = new List<WorkOrderInfo>();
                if (results.Entities.Count > 0)
                {
                    foreach (var entity in results.Entities)
                    {
                        WorkOrderInfo workOrderInfo = new WorkOrderInfo
                        {
                            JobId = entity.Contains("msdyn_name") ? entity.GetAttributeValue<string>("msdyn_name") : "",
                            Substatus = entity.Contains("msdyn_substatus") ? entity.GetAttributeValue<EntityReference>("msdyn_substatus").Name : "",
                            Createdon = entity.Contains("createdon") ? entity.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString() : "",
                            Productsubcategory = entity.Contains("hil_productsubcategory") ? entity.GetAttributeValue<EntityReference>("hil_productsubcategory").Name : "",
                            Productcategory = entity.Contains("hil_productcategory") ? entity.GetAttributeValue<EntityReference>("hil_productcategory").Name : "",
                            Callsubtype = entity.Contains("hil_callsubtype") ? entity.GetAttributeValue<EntityReference>("hil_callsubtype").Name : "",
                            Customer = entity.Contains("hil_customerref") ? entity.GetAttributeValue<EntityReference>("hil_customerref").Name : "",
                            Owner = entity.Contains("hil_owneraccount") ? entity.GetAttributeValue<EntityReference>("hil_owneraccount").Name : "",
                        };
                        workOrderInfos.Add(workOrderInfo);
                    }
                    workOrderResponse.workOrderInfos = workOrderInfos;
                    workOrderResponse.StatusCode = "200";
                    workOrderResponse.Message = "Success";
                }
                else
                {
                    workOrderResponse.StatusCode = "204";
                    workOrderResponse.Message = "No Record Found";
                }
            }
            catch (Exception ex)
            {
                workOrderResponse.StatusCode = "500";
                workOrderResponse.Message = "D365 Internal Server Error : " + ex.Message.ToUpper();
            }
            return workOrderResponse;
        }
    }

    public class JobStatusDTO
    {

        public string job_id { get; set; }
        public string mobile_number { get; set; }
        public string serial_number { get; set; }
        public string product_category { get; set; }
        public string product_subcategory { get; set; }
        public string call_type { get; set; }
        public string customer_complaint { get; set; }
        public string assigned_resource { get; set; }
        public string job_substatus { get; set; }
        public string technician_remarks { get; set; }
        public string closed_on { get; set; }
        public string cancel_reason { get; set; }
        public string closure_remarks { get; set; }
        public string webclosure_remarks { get; set; }
        public List<JobProductDTO> spare_parts { get; set; }
        public string status_code { get; set; }
        public string status_description { get; set; }
    }
    public class JobProductDTO
    {
        public string index { get; set; }
        public string product_code { get; set; }
        public string product_description { get; set; }
    }
    public class JobDetails
    {

        public string JOB_ID { get; set; }

        public string JOB_STATUS { get; set; }

        public string ASSIGNED_TO { get; set; }

        public string TYPE_OF_OWNER { get; set; }

        public DateTime CREATED_ON { get; set; }

        public string PROD_SUB_CAT { get; set; }

        public string CALL_STYPE { get; set; }

        public string CONSUMER_NAME { get; set; }

        public string MOBILE_NO { get; set; }

        public string CONSUMER_ADDRESS { get; set; }
    }
    public class JobOutput
    {
        public string Job_ID { get; set; }
        public string MobileNumber { get; set; }
        public string CallSubType { get; set; }
        public string Job_Loggedon { get; set; }
        public string Job_Status { get; set; }
        public string Job_AssignedTo { get; set; }
        public string Job_Asset { get; set; }
        public string Job_Category { get; set; }
        public string Job_NOC { get; set; }
        public string Job_ClosedOn { get; set; }
        public string Customer_name { get; set; }
        public string Customer_Address { get; set; }
        public string Product { get; set; }
        public Guid ProductCategoryGuid { get; set; }
        public string ProductCategoryName { get; set; }
        public string ChiefComplaint { get; set; }
    }
    public class WorkOrderResponse
    {
        public string StatusCode { get; set; }
        public string Message { get; set; }
        public List<WorkOrderInfo> workOrderInfos { get; set; }

    }

    public class WorkOrderInfo
    {
        public string JobId { get; set; }

        public string Substatus { get; set; }

        public string Createdon { get; set; }

        public string Productsubcategory { get; set; }

        public string Productcategory { get; set; }

        public string Callsubtype { get; set; }

        public string Customer { get; set; }

        public string Owner { get; set; }

    }

    public class WorkOrderRequest
    {
        public string DealerCode { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
    }

}