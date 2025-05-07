using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Text;
using Newtonsoft.Json;
using System.Text.RegularExpressions;
using System.Linq;


namespace D365WebJobs
{
    public class CallMasking
    {
        public ResposeDataCallMasking GetCustomerOpenJobs(IOrganizationService service,RequestDataCallMasking _requestData)
        {
            #region Variable declaration
            string requestData = JsonConvert.SerializeObject(_requestData);

            ResposeDataCallMasking returnObj = new ResposeDataCallMasking();
            returnObj.ResultMessage = "SUCCESS";
            returnObj.ResultStatus = true;
            #endregion

            try
            {
                //IOrganizationService service = ConnectToCRM.GetOrgServiceProd();
                if (service != null)
                {
                    StringBuilder _fetchXML = new StringBuilder();

                    _fetchXML.Append(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='msdyn_workorder'>
                    <attribute name='msdyn_name' />
                    <attribute name='createdon' />
                    <attribute name='hil_customername' />
                    <attribute name='msdyn_workorderid' />
                    <order attribute='createdon' descending='true' />
                    <filter type='and'>
                        <condition attribute='hil_mobilenumber' operator='eq' value='" + _requestData.MobileNumber + @"' />
                        <condition attribute='msdyn_substatus' operator='not-in'>
                        <value uiname='Closed' uitype='msdyn_workordersubstatus'>{1727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                        <value uiname='KKG Audit Failed' uitype='msdyn_workordersubstatus'>{6C8F2123-5106-EA11-A811-000D3AF057DD}</value>
                        <value uiname='Canceled' uitype='msdyn_workordersubstatus'>{1527FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                        </condition>
                        <condition attribute='hil_isocr' operator='ne' value='1' /> 
                    </filter>
                    <link-entity name='systemuser' from='systemuserid' to='owninguser' visible='false' link-type='outer' alias='user'>
                        <attribute name='mobilephone' />
                        <attribute name='address1_telephone1' />
                    </link-entity>
                    </entity>
                    </fetch>");
                    EntityCollection openjobs = service.RetrieveMultiple(new FetchExpression(_fetchXML.ToString()));
                    returnObj.MobileNumber = _requestData.MobileNumber;
                    returnObj.OpenJobs = openjobs.Entities.Count();
                    List<JobsData> _JobList = new List<JobsData>();
                    string techno = string.Empty;
                    foreach (Entity ent in openjobs.Entities)
                    {
                        _JobList.Add(new JobsData()
                        {
                            Id = ent.GetAttributeValue<string>("msdyn_name"),
                            TechnicianNumber = ent.Attributes.Contains("user.mobilephone") ? ent.GetAttributeValue<AliasedValue>("user.mobilephone").Value.ToString() : ent.GetAttributeValue<AliasedValue>("user.address1_telephone1").Value.ToString()
                        });
                    }
                    if (returnObj.OpenJobs > 0)
                    {
                        returnObj.CustomerName = openjobs.Entities[0].GetAttributeValue<string>("hil_customername");
                        if (!string.IsNullOrEmpty(_requestData.JobNumber))
                        {
                            _JobList = GetJobsByLastFourDigits(_JobList, _requestData.JobNumber);
                            if (_JobList.Count != 1)
                            {
                                returnObj.ResultMessage = "No Open job found with Job Id end with " + _requestData.JobNumber;
                            }
                            returnObj.OpenJobs = 1;
                        }
                        returnObj.TechnicianMobileNo = _JobList[0].TechnicianNumber;
                    }
                    else
                    {
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
        public static List<JobsData> GetJobsByLastFourDigits(List<JobsData> jobs, string lastFourDigits)
        {
            List<JobsData> _retObj = jobs.Where(job => job.Id.EndsWith(lastFourDigits)).ToList();

            return _retObj.Count == 0 ? jobs : _retObj;
        }
        public static CDR_Response PushCDR(IOrganizationService service,CDR_Request request)
        {
            CDR_Response returnObj = new CDR_Response();
            try
            {
                var CrmURL = "https://havells.crm8.dynamics.com/";
                string finalString = string.Format("AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=", CrmURL);
                
                if (service != null)
                {
                    string cdr_report = Newtonsoft.Json.JsonConvert.SerializeObject(request);
                    QueryExpression query = new QueryExpression("phonecall");
                    query.ColumnSet = new ColumnSet("description", "actualstart", "actualend", "hil_disposition", "phonenumber", "hil_calledtonum", "hil_callingnumber", "subject", "category");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("subject", ConditionOperator.Equal, request.Correlation_ID);

                    EntityCollection phonecalls = service.RetrieveMultiple(query);

                    int count = phonecalls.Entities.Count();

                    request.Overall_Call_Duration = "00:" + request.Overall_Call_Duration;
                    TimeSpan endTime = TimeSpan.Parse(request.Time).Add(TimeSpan.Parse(request.Overall_Call_Duration));
                    DateTime start;
                    DateTime.TryParse(request.Date.Split('/')[2] + "-" + request.Date.Split('/')[1] + "-" + request.Date.Split('/')[0] + " " + request.Time, out start);
                    DateTime end = start.Add(TimeSpan.Parse(request.Overall_Call_Duration));
                    int status = 0;
                    if (request.Overall_Call_Status == "Missed") status = 8;
                    else if (request.Overall_Call_Status == "Answered") status = 9;

                    if (count == 0)
                    {
                        Entity phonecall = new Entity("phonecall");
                        phonecall["subject"] = request.Correlation_ID;
                        phonecall["actualstart"] = start;
                        phonecall["actualend"] = end;
                        phonecall["phonenumber"] = request.Caller_Id;
                        phonecall["hil_calledtonum"] = request.Destination_Number + " " + request.Destination_Name;
                        phonecall["hil_callingnumber"] = request.Caller_Number + " " + request.Caller_Name;
                        phonecall["description"] = request.Recording;
                        phonecall["directioncode"] = request.Call_Type.ToUpper() == "OUTBOUND" ? true : false;
                        phonecall["hil_disposition"] = new OptionSetValue(status);

                        service.Create(phonecall);
                    }
                    else
                    {
                        Entity phonecall = new Entity(phonecalls.Entities[0].LogicalName, phonecalls.Entities[0].Id);
                        phonecall["actualstart"] = start;
                        phonecall["actualend"] = end;
                        phonecall["phonenumber"] = request.Caller_Id;
                        //phonecall["hil_calledtonum"] = request.Destination_Number +  " " +request.Destination_Name;
                        //phonecall["hil_callingnumber"] = request.Caller_Number + " " + request.Caller_Name;
                        phonecall["description"] = request.Recording;
                        phonecall["directioncode"] = request.Call_Type.ToUpper() == "OUTBOUND" ? new OptionSetValue(1) : new OptionSetValue(0);
                        phonecall["hil_disposition"] = new OptionSetValue(status);

                        service.Update(phonecall);

                        returnObj.ResultMessage = "Success";
                        returnObj.ResultStatus = true;
                    }
                }
                else
                {
                    returnObj = new CDR_Response { ResultStatus = false, ResultMessage = "D365 Service is not available. : "};
                }
            }
            catch (Exception ex)
            {
                returnObj = new CDR_Response { ResultStatus = false, ResultMessage = "D365 Internal Server Error : " + ex.Message };

            }
            return returnObj;
        }
    }
    
    public class RequestDataCallMasking
    {
        public string MobileNumber { get; set; }
        public string JobNumber { get; set; }
    }
    
    public class ResposeDataCallMasking
    {
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
    
    public class CDR_Request
    {
        public string Caller_Id { get; set; }
        public string Correlation_ID { get; set; }
        public string Overall_Call_Status { get; set; }
        public string Time { get; set; }
        public string Date { get; set; }
        public string Overall_Call_Duration { get; set; }
        public string Caller_Number { get; set; }
        public string Destination_Number { get; set; }
        public string Caller_Name { get; set; }
        public string Destination_Name { get; set; }
        public string Recording { get; set; }
        public string Call_Type { get; set; }
    }

    
    public class CDR_Response
    {
        public string ResultMessage { get; set; }
        public bool ResultStatus { get; set; }
    }
}
