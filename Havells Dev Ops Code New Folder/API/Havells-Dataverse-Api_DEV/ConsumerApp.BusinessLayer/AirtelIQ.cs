using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace ConsumerApp.BusinessLayer
{
    public class AirtelIQ
    {
        public ResposeDataCallMasking GetOpenJobs(RequestDataCallMasking _requestData)
        {
            string requestData = JsonConvert.SerializeObject(_requestData);
            ResposeDataCallMasking returnObj = new ResposeDataCallMasking();

            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgServiceDev1();
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

                    //returnObj.IsJobFound = true;
                    returnObj.ResultStatus = true;

                    List<JobsData> _JobList = new List<JobsData>();


                    foreach (Entity Job in openjobs.Entities)
                    {
                        JobsData details = new JobsData();
                        details.Id = Job.GetAttributeValue<string>("msdyn_name");
                        details.TechnicianNumber = Job.Attributes.Contains("user.mobilephone") ? Job.GetAttributeValue<AliasedValue>("user.mobilephone").Value.ToString() : Job.GetAttributeValue<AliasedValue>("user.address1_telephone1").Value.ToString();
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

                                if (jobConv.Entities.Count > 0)
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
                                returnObj.TechnicianMobileNo = convColl.Entities[0].GetAttributeValue<string>("hil_callingnumber");
                            }

                        }
                        else
                        {
                            returnObj.IsJobFound = true;
                            if (countConv > 0)
                            {
                                returnObj.ResultMessage = "Conversation found regarding the mobile number: " + _requestData.MobileNumber;
                                returnObj.TechnicianMobileNo = convColl.Entities[0].GetAttributeValue<string>("hil_callingnumber");

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

        internal void createRecording(IOrganizationService service, CDR_Request req, Guid ActivityId)
        {
            EntityCollection tempCollection = checkExistingRecording(service, ActivityId);
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

        internal EntityCollection checkExistingRecording(IOrganizationService service, Guid ActivityID)
        {
            QueryExpression query = new QueryExpression("msdyn_recording");
            query.ColumnSet = new ColumnSet("msdyn_ci_url", "msdyn_ci_transcript_json", "msdyn_phone_call_activity");
            query.Criteria.AddCondition("msdyn_phone_call_activity", ConditionOperator.Equal, ActivityID);
            return service.RetrieveMultiple(query);
        }

        public CDR_Response PushCDRToD365(CDR_Request request)
        {
            CDR_Response response = new CDR_Response();
            response.ResultMessage = "Failed";
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgServiceDev1();

                string cdr_report = Newtonsoft.Json.JsonConvert.SerializeObject(request);
                QueryExpression query = new QueryExpression("phonecall");
                query.ColumnSet = new ColumnSet("description", "actualstart", "actualend", "hil_disposition", "phonenumber", "hil_calledtonum", "hil_callingnumber", "subject", "hil_alternatenumber1", "scheduledstart", "directioncode", "from", "to");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("subject", ConditionOperator.Equal, request.Correlation_ID);

                EntityCollection phonecalls = service.RetrieveMultiple(query);

                int count = phonecalls.Entities.Count();

                //request.Overall_Call_Duration = "00:" + request.Overall_Call_Duration;
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

                    request.Caller_Number = request.Caller_Number.Substring(Math.Max(0, request.Caller_Number.Length - 10));

                    query = new QueryExpression("contact");
                    query.ColumnSet = new ColumnSet("mobilephone");
                    query.Criteria.AddCondition("mobilephone", ConditionOperator.EndsWith, request.Caller_Number);

                    EntityCollection usersCollection = service.RetrieveMultiple(query);
                    if (usersCollection.Entities.Count > 0)
                    {
                        Entity fromParty = new Entity("activityparty");
                        fromParty["partyid"] = new EntityReference("contact", usersCollection.Entities[0].Id);
                        phonecall["from"] = new Entity[] { fromParty };
                    }

                    request.Destination_Number = request.Destination_Number.Substring(Math.Max(0, request.Destination_Number.Length - 10));

                    query = new QueryExpression("systemuser");
                    query.ColumnSet = new ColumnSet("mobilephone");
                    query.Criteria.AddCondition("mobilephone", ConditionOperator.EndsWith, request.Destination_Number);

                    usersCollection = service.RetrieveMultiple(query);
                    if (usersCollection.Entities.Count > 0)
                    {
                        Entity toParty = new Entity("activityparty");
                        toParty["partyid"] = new EntityReference("systemuser", usersCollection.Entities[0].Id);
                        phonecall["to"] = new Entity[] { toParty };

                    }

                    phonecall["hil_disposition"] = new OptionSetValue(status);

                    Guid PhoneGuid = service.Create(phonecall);

                    //Entity recording = new Entity("msdyn_recording");
                    //recording["msdyn_ci_url"] = request.Recording;

                    //recording["msdyn_ci_transcript_json"] = Newtonsoft.Json.JsonConvert.SerializeObject(request);

                    //recording["msdyn_phone_call_activity"] = new EntityReference("phonecall", PhoneGuid);

                    //service.Create(recording);

                    createRecording(service, request, PhoneGuid);
                    response.ResultMessage = "Record Created Successfully";
                }
                else
                {

                    Entity phonecall = new Entity(phonecalls.Entities[0].LogicalName, phonecalls.Entities[0].Id);
                    phonecall["actualstart"] = start.AddHours(-5.5);
                    phonecall["actualend"] = end.AddMinutes(-330);
                    phonecall["hil_alternatenumber1"] = convertation_duration + " seconds";
                    //phonecall["hil_calledtonum"] = request.Destination_Number +  " " +request.Destination_Name;
                    //phonecall["hil_callingnumber"] = request.Caller_Number + " " + request.Caller_Name;
                    phonecall["description"] = "Caller: " + request.Caller_Number + " " + request.Caller_Name + ", Destination: " + request.Destination_Name + ", Caller Status: " + request.Caller_Status + ", Destination Status: " + request.Destination_Status; ;
                    phonecall["hil_disposition"] = new OptionSetValue(status);
                    service.Update(phonecall);

                    //Entity recording = new Entity("msdyn_recording");
                    //recording["msdyn_ci_url"] = request.Recording;

                    //recording["msdyn_ci_transcript_json"] = Newtonsoft.Json.JsonConvert.SerializeObject(request);

                    //recording["msdyn_phone_call_activity"] = phonecalls.Entities[0].ToEntityReference();


                    //service.Create(recording);

                    createRecording(service, request, phonecalls.Entities[0].Id);

                    response.ResultMessage = "Record Created Successfully";

                }
            }
            catch (Exception ex)
            {
                response.ResultStatus = false;
                response.ResultMessage = "Something went wrong";

            }
            return response;
        }
    }

    [DataContract]
    public class RequestDataCallMasking
    {
        [DataMember]
        public string MobileNumber { get; set; }
        [DataMember]
        public string JobNumber { get; set; }
    }
    [DataContract]
    public class ResposeDataCallMasking
    {
        [DataMember]
        public bool IsJobFound { get; set; }
        [DataMember]
        public string ResultMessage { get; set; }
        [DataMember]
        public bool ResultStatus { get; set; }
        [DataMember]
        public string CustomerName { get; set; }
        [DataMember]
        public string MobileNumber { get; set; }
        [DataMember]
        public int OpenJobs { get; set; }
        [DataMember]
        public string TechnicianMobileNo { get; set; }
    }
    [DataContract]
    public class JobsData
    {
        [DataMember]
        public string Id { get; set; }
        [DataMember]
        public string TechnicianNumber { get; set; }
        [DataMember]
        public string CreatedOn { get; set; }
    }
    [DataContract]
    public class CDR_Response
    {
        [DataMember]
        public string ResultMessage { get; set; }
        [DataMember]
        public bool ResultStatus { get; set; }
    }
    [DataContract]
    public class CDR_Request
    {
        [DataMember]
        public string Caller_Id { get; set; }
        [DataMember]
        public string Caller_Name { get; set; }
        [DataMember]
        public string Caller_Number { get; set; }
        [DataMember]
        public string Call_Type { get; set; }
        [DataMember]
        public string Caller_Status { get; set; }
        [DataMember]
        public string Conversation_Duration { get; set; }
        [DataMember]
        public string Correlation_ID { get; set; }
        [DataMember]
        public string Date { get; set; }
        [DataMember]
        public string Destination_Name { get; set; }
        [DataMember]
        public string Destination_Number { get; set; }
        [DataMember]
        public string Destination_Status { get; set; }
        [DataMember]
        public string Overall_Call_Duration { get; set; }
        [DataMember]
        public string Overall_Call_Status { get; set; }
        [DataMember]
        public string Recording { get; set; }
        [DataMember]
        public string Time { get; set; }
    }

}