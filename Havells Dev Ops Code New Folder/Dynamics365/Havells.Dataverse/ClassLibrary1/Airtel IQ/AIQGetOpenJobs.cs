using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.AMC.Airtel_IQ
{
    public class AIQGetOpenJobs : IPlugin
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
                ResposeDataCallMasking returnObj = new ResposeDataCallMasking();
                List<JobsData> _JobList = new List<JobsData>();
                string mobileNumber = Convert.ToString(context.InputParameters["MobileNumber"]);
                string jobNumber = Convert.ToString(context.InputParameters["JobNumber"]);
                if (!string.IsNullOrWhiteSpace(mobileNumber))
                {
                    if (!APValidate.IsValidMobileNumber(mobileNumber))
                    {
                        context.OutputParameters["data"] = JsonSerializer.Serialize(new ResposeDataCallMasking { ResultStatus = false, ResultMessage = "Invalid Mobile Number." });
                        return;
                    }
                }

                if (!string.IsNullOrWhiteSpace(jobNumber))
                {
                    if (jobNumber.Length > 16 || jobNumber.Length < 4)
                    {
                        context.OutputParameters["data"] = JsonSerializer.Serialize(new ResposeDataCallMasking { ResultStatus = false, ResultMessage = "Invalid Job Id." });
                        return;
                    }
                    if (!APValidate.NumericValue(jobNumber))
                    {
                        context.OutputParameters["data"] = JsonSerializer.Serialize(new ResposeDataCallMasking { ResultStatus = false, ResultMessage = "Invalid Job Id format." });
                        return;
                    }
                }
                returnObj = GetOpenJobs(new RequestDataCallMasking { MobileNumber = mobileNumber, JobNumber = jobNumber }, service);
                jsonResponse = JsonSerializer.Serialize(returnObj);
                context.OutputParameters["data"] = jsonResponse;
                return;
            }
            catch (Exception ex)
            {
                var errorResponse = new ResposeDataCallMasking
                {
                    ResultStatus = false,
                    ResultMessage = "D365 Internal Server Error: " + ex.Message
                };
            }
        }
        public ResposeDataCallMasking GetOpenJobs(RequestDataCallMasking _requestData, IOrganizationService service)
        {
            string requestData = JsonSerializer.Serialize(_requestData);
            ResposeDataCallMasking returnObj = new ResposeDataCallMasking();
            try
            {
                if (service != null)
                {
                    string fetchjobs = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorder'>
                        <attribute name='msdyn_name' />
                        <attribute name='createdon' />
                        <attribute name='msdyn_workorderid' />
                        <attribute name='hil_customername' />
                        <order attribute='createdon' descending='true' />
                        <filter type='and'>
                        <condition attribute='msdyn_substatus' operator='not-in'>
                            <value uiname='Closed' uitype='msdyn_workordersubstatus'>{1727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                            <value uiname='Canceled' uitype='msdyn_workordersubstatus'>{1527FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                            <value uiname='KKGÂ AuditÂ Failed' uitype='msdyn_workordersubstatus'>{6C8F2123-5106-EA11-A811-000D3AF057DD}</value>
                        </condition>
                        <condition attribute='hil_isocr' operator='ne' value='1' />
                        <condition attribute='createdon' operator='last-x-days' value='90' />
                        </filter>
                        <link-entity name='systemuser' from='systemuserid' to='owninguser' visible='false' link-type='outer' alias='user'>
                        <attribute name='mobilephone' />
                        <attribute name='address1_telephone1' />
                        </link-entity>
                        <link-entity name='contact' from='contactid' to='hil_customerref' link-type='inner' alias='ab'>
                        <filter type='and'>
                        <condition attribute='mobilephone' operator='eq' value='" + _requestData.MobileNumber + @"' />
                        </filter>
                        </link-entity>
                        </entity>
                        </fetch>";

                    EntityCollection openjobs = service.RetrieveMultiple(new FetchExpression(fetchjobs));

                    returnObj.MobileNumber = _requestData.MobileNumber;
                    returnObj.OpenJobs = openjobs.Entities.Count();
                    returnObj.ResultStatus = true;
                    List<JobsData> _JobList = new List<JobsData>();
                    foreach (Entity Job in openjobs.Entities)
                    {
                        JobsData details = new JobsData();
                        details.Id = Job.GetAttributeValue<string>("msdyn_name");
                        //details.TechnicianNumber = Job.Attributes.Contains("user.mobilephone") ? Job.GetAttributeValue<AliasedValue>("user.mobilephone").Value.ToString() : Job.GetAttributeValue<AliasedValue>("user.address1_telephone1").Value.ToString();
                        details.TechnicianNumber = (Job.Attributes.Contains("user.mobilephone") && Job.GetAttributeValue<AliasedValue>("user.mobilephone")?.Value != null) ? Job.GetAttributeValue<AliasedValue>("user.mobilephone").Value.ToString() : (Job.GetAttributeValue<AliasedValue>("user.address1_telephone1")?.Value != null) ? Job.GetAttributeValue<AliasedValue>("user.address1_telephone1").Value.ToString() : string.Empty;
                        details.CreatedOn = Job.GetAttributeValue<DateTime>("createdon").ToString(); ;
                        _JobList.Add(details);
                    }
                    _JobList.Sort((x, y) => DateTime.Compare(DateTime.Parse(y.CreatedOn), DateTime.Parse(x.CreatedOn)));
                    QueryExpression query = new QueryExpression("phonecall");
                    query.ColumnSet = new ColumnSet("createdon", "hil_calledtonum", "hil_callingnumber");
                    query.Criteria.AddCondition("hil_calledtonum", ConditionOperator.EndsWith, _requestData.MobileNumber);
                    query.Criteria.AddCondition("statecode", ConditionOperator.NotEqual, 2); // Cancelled 
                    query.AddOrder("createdon", OrderType.Descending);
                    LinkEntity joblink = query.AddLink("msdyn_workorder", "regardingobjectid", "msdyn_workorderid", JoinOperator.Inner);
                    joblink.LinkCriteria.AddCondition("msdyn_substatus", ConditionOperator.NotEqual, "1727FA6C-FA0F-E911-A94E-000D3AF060A1");
                    query.LinkEntities.Add(joblink);
                    EntityCollection convColl = service.RetrieveMultiple(query);
                    int countConv = convColl.Entities.Count;
                    if (_JobList.Count > 0)
                    {
                        returnObj.IsJobFound = true;
                        returnObj.CustomerName = openjobs.Entities[0].GetAttributeValue<string>("hil_customername");

                        if (!string.IsNullOrEmpty(_requestData.JobNumber))
                        {

                            _JobList = GetJobsByLastFourDigits(_JobList, _requestData.JobNumber);
                            if (_JobList.Count > 0)
                            {

                                query = new QueryExpression("phonecall");
                                query.ColumnSet = new ColumnSet("createdon", "hil_calledtonum", "hil_callingnumber", "regardingobjectid");
                                query.Criteria = new FilterExpression(LogicalOperator.And);
                                query.Criteria.AddCondition("hil_calledtonum", ConditionOperator.EndsWith, _requestData.MobileNumber);
                                query.Criteria.AddCondition("regardingobjectidname", ConditionOperator.EndsWith, _requestData.JobNumber);
                                query.Criteria.AddCondition("statecode", ConditionOperator.NotEqual, 2); // Cancelled 
                                query.AddOrder("createdon", OrderType.Descending);

                                EntityCollection jobConv = service.RetrieveMultiple(query);

                                if (jobConv.Entities.Count > 0 && jobConv.Entities[0].Contains("hil_callingnumber"))
                                {
                                    returnObj.ResultMessage = "Conversation found regarding the job number: " + _requestData.JobNumber;
                                    returnObj.TechnicianMobileNo = jobConv.Entities[0].GetAttributeValue<string>("hil_callingnumber");
                                }
                                else
                                {
                                    returnObj.ResultMessage = "No Conversation found regarding the job number: " + _requestData.JobNumber;
                                    returnObj.TechnicianMobileNo = _JobList[0].TechnicianNumber;
                                }

                                returnObj.OpenJobs = 1;

                            }
                            else
                            {
                                returnObj.IsJobFound = false;
                                returnObj.OpenJobs = 0;
                                returnObj.ResultMessage = "No open job found against the job Id ending with " + _requestData.JobNumber;
                                returnObj.TechnicianMobileNo = convColl.Entities.Count > 0 && convColl.Entities[0].Contains("hil_callingnumber") ? convColl.Entities[0].GetAttributeValue<string>("hil_callingnumber") : string.Empty;
                            }
                        }
                        else
                        {
                            returnObj.IsJobFound = true;
                            if (countConv > 0 && convColl.Entities[0].Contains("hil_callingnumber"))
                            {
                                returnObj.ResultMessage = "Conversation found regarding the mobile number: " + _requestData.MobileNumber;
                                returnObj.TechnicianMobileNo = convColl.Entities.Count > 0 && convColl.Entities[0].Contains("hil_callingnumber") ? convColl.Entities[0].GetAttributeValue<string>("hil_callingnumber") : string.Empty;


                            }
                            else
                            {
                                returnObj.ResultMessage = "No Conversation found regarding the mobile number: " + _requestData.MobileNumber;
                                returnObj.TechnicianMobileNo = _JobList[0].TechnicianNumber;
                            }
                            returnObj.OpenJobs = _JobList.Count;
                        }
                    }
                    else
                    {
                        returnObj.IsJobFound = false;
                        returnObj = new ResposeDataCallMasking { ResultStatus = false, ResultMessage = "No Open job found" };
                    }

                }
                else
                {
                    returnObj = new ResposeDataCallMasking { ResultStatus = false, ResultMessage = "D365 Service is not available. : " };
                }
            }
            catch (Exception ex)
            {
                returnObj = new ResposeDataCallMasking { ResultStatus = false, ResultMessage = "D365 Internal Server Error : " + ex.Message };
            }
            return returnObj;
        }
        private List<JobsData> GetJobsByLastFourDigits(List<JobsData> jobs, string lastFourDigits)
        {
            return jobs.Where(job => job.Id.EndsWith(lastFourDigits)).ToList();
        }
        public class RequestDataCallMasking
        {
            public string MobileNumber { get; set; }
            public string JobNumber { get; set; }
        }
        public class ResposeDataCallMasking
        {
            public bool IsJobFound { get; set; }
            public string ResultMessage { get; set; }
            public bool ResultStatus { get; set; }
            public string CustomerName { get; set; }
            public string MobileNumber { get; set; }
            public int OpenJobs { get; set; }
            public string TechnicianMobileNo { get; set; }
        }
        public class JobsData
        {
            public string Id { get; set; }
            public string TechnicianNumber { get; set; }
            public string CreatedOn { get; set; }
        }
    }
}