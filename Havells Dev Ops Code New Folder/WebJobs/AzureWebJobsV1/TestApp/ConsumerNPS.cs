using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestApp
{
    public class ConsumerNPS
    {
        public static ConsumerSurveyDTO GetConsumerSurvey(IOrganizationService _service, ConsumerSurveyDTO _consumerSurveyData)
        {
            QueryExpression _queryExp;
            EntityCollection _entcoll;
            Guid _jobId = Guid.Empty;
            EntityReference _erJob = null;
            ConsumerSurveyDTO _consumerNPSResult = new ConsumerSurveyDTO();
            try
            {
                _queryExp = new QueryExpression("msdyn_workorder");
                _queryExp.ColumnSet = new ColumnSet("msdyn_workorderid", "hil_customerref", "hil_brand");
                _queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                _queryExp.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, _consumerSurveyData.JobId);
                _entcoll = _service.RetrieveMultiple(_queryExp);
                if (_entcoll.Entities.Count > 0)
                {
                    _erJob = _entcoll.Entities[0].ToEntityReference();
                    _consumerNPSResult.JobId = _consumerSurveyData.JobId;
                    if (_entcoll.Entities[0].Attributes.Contains("hil_customerref")) {
                        _consumerNPSResult.CustomerName = _entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_customerref").Name;
                    }
                    if (_entcoll.Entities[0].Attributes.Contains("hil_brand"))
                    {
                        int _brand = _entcoll.Entities[0].GetAttributeValue<OptionSetValue>("hil_brand").Value;
                        _consumerNPSResult.BrandName = _brand == 2 ? "Lloyd" : "Havells";
                    }
                }
                else
                {

                }
                _queryExp = new QueryExpression("hil_consumernpssurvey");
                _queryExp.ColumnSet = new ColumnSet(true);
                _queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                _queryExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _erJob.Id);
                _entcoll = _service.RetrieveMultiple(_queryExp);

                if (_entcoll.Entities.Count > 0)
                {
                    if (_entcoll.Entities[0].Attributes.Contains("hil_passivesresponse"))
                    {
                        _consumerNPSResult.PassivesResponse = _entcoll.Entities[0].GetAttributeValue<string>("hil_passivesresponse");
                    }
                    if (_entcoll.Entities[0].Attributes.Contains("hil_detractorsresponse"))
                    {
                        _consumerNPSResult.DetractorsResponse = _entcoll.Entities[0].GetAttributeValue<string>("hil_detractorsresponse");
                    }
                    if (_entcoll.Entities[0].Attributes.Contains("hil_promotersresponse"))
                    {
                        _consumerNPSResult.PromotersResponse = _entcoll.Entities[0].GetAttributeValue<string>("hil_promotersresponse");
                    }
                    if (_entcoll.Entities[0].Attributes.Contains("hil_feedback"))
                    {
                        _consumerNPSResult.Feedback = _entcoll.Entities[0].GetAttributeValue<string>("hil_feedback");
                    }
                    if (_entcoll.Entities[0].Attributes.Contains("hil_npsvalue"))
                    {
                        _consumerNPSResult.NPSValue = _entcoll.Entities[0].GetAttributeValue<string>("hil_npsvalue");
                    }
                    if (_entcoll.Entities[0].Attributes.Contains("hil_serviceengineerrating"))
                    {
                        _consumerNPSResult.ServiceEngineerRating = _entcoll.Entities[0].GetAttributeValue<string>("hil_serviceengineerrating");
                    }
                    if (_entcoll.Entities[0].Attributes.Contains("hil_submitstatus"))
                    {
                        _consumerNPSResult.SubmitStatus = _entcoll.Entities[0].GetAttributeValue<bool>("hil_submitstatus");
                    }
                    _consumerNPSResult.Result = true;
                    _consumerNPSResult.ResultMessage = "Success";
                }
                else
                {
                    _consumerNPSResult.Result = false;
                    _consumerNPSResult.ResultMessage = "No Survey found";
                }
            }
            catch (Exception ex)
            {

            }
            return _consumerNPSResult;
        }
        public static ConsumerSurveyDTO CaptureConsumerNPS(IOrganizationService _service, ConsumerSurveyDTO _consumerSurveyData) {
            QueryExpression _queryExp;
            EntityCollection _entcoll;
            Guid _jobId = Guid.Empty;
            EntityReference _erJob = null;
            ConsumerSurveyDTO _consumerNPSResult = new ConsumerSurveyDTO();
            try
            {
                _queryExp = new QueryExpression("msdyn_workorder");
                _queryExp.ColumnSet = new ColumnSet("msdyn_workorderid");
                _queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                _queryExp.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, _consumerSurveyData.JobId);
                _entcoll = _service.RetrieveMultiple(_queryExp);
                if (_entcoll.Entities.Count > 0)
                {
                    _erJob = _entcoll.Entities[0].ToEntityReference();
                }
                else { 
                
                }
                Entity entObj = new Entity("hil_consumernpssurvey");

                if (_consumerSurveyData.PassivesResponse.Trim().Length > 0 && _consumerSurveyData.PassivesResponse != null) {
                    entObj.Attributes["hil_passivesresponse"] = _consumerSurveyData.PassivesResponse;
                }
                if (_consumerSurveyData.DetractorsResponse.Trim().Length > 0 && _consumerSurveyData.DetractorsResponse != null)
                {
                    entObj.Attributes["hil_detractorsresponse"] = _consumerSurveyData.DetractorsResponse;
                }
                if (_consumerSurveyData.Feedback.Trim().Length > 0 && _consumerSurveyData.Feedback != null)
                {
                    entObj.Attributes["hil_feedback"] = _consumerSurveyData.Feedback;
                }
                if (_consumerSurveyData.NPSValue.Trim().Length > 0 && _consumerSurveyData.NPSValue != null)
                {
                    entObj.Attributes["hil_npsvalue"] = _consumerSurveyData.NPSValue;
                }
                if (_consumerSurveyData.PromotersResponse.Trim().Length > 0 && _consumerSurveyData.PromotersResponse != null)
                {
                    entObj.Attributes["hil_promotersresponse"] = _consumerSurveyData.PromotersResponse;
                }
                if (_consumerSurveyData.ServiceEngineerRating.Trim().Length > 0 && _consumerSurveyData.ServiceEngineerRating != null)
                {
                    entObj.Attributes["hil_serviceengineerrating"] = _consumerSurveyData.ServiceEngineerRating;
                }
                entObj.Attributes["hil_jobid"] = _erJob;
                entObj.Attributes["hil_submitstatus"] = _consumerSurveyData.SubmitStatus;

                _queryExp = new QueryExpression("hil_consumernpssurvey");
                _queryExp.ColumnSet = new ColumnSet("hil_submitstatus");
                _queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                _queryExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _erJob.Id);
                _entcoll = _service.RetrieveMultiple(_queryExp);

                if (_entcoll.Entities.Count > 0)
                {
                    if (_entcoll.Entities[0].GetAttributeValue<bool>("hil_submitstatus"))
                    {
                        _consumerNPSResult.Result = false;
                        _consumerNPSResult.ResultMessage = "Already Submitted";
                    }
                    else
                    {
                        entObj.Id = _entcoll.Entities[0].Id;
                        _service.Update(entObj);
                        _consumerNPSResult.Result = true;
                        _consumerNPSResult.ResultMessage = "Success";
                    }
                }
                else
                {
                    _service.Create(entObj);
                    _consumerNPSResult.Result = true;
                    _consumerNPSResult.ResultMessage = "Success";
                }
            }
            catch (Exception ex)
            {
                
            }
            return _consumerNPSResult;
        }
        public static void CaptureTechnicianRating(IOrganizationService _service, ConsumerSurveyDTO _consumerSurveyData)
        {
            QueryExpression _queryExp;
            EntityCollection _entcoll;
            Guid _jobId = Guid.Empty;
            EntityReference _erJob = null;
            ConsumerSurveyDTO _consumerNPSResult = new ConsumerSurveyDTO();
            try
            {
                _queryExp = new QueryExpression("msdyn_workorder");
                _queryExp.ColumnSet = new ColumnSet("msdyn_workorderid");
                _queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                _queryExp.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, _consumerSurveyData.JobId);
                _entcoll = _service.RetrieveMultiple(_queryExp);
                if (_entcoll.Entities.Count > 0)
                {
                    _erJob = _entcoll.Entities[0].ToEntityReference();
                }
                else
                {

                }
                Entity entObj = new Entity("hil_consumernpssurvey");
                if (_consumerSurveyData.ServiceEngineerRating.Trim().Length > 0 && _consumerSurveyData.ServiceEngineerRating != null)
                {
                    entObj.Attributes["hil_serviceengineerrating"] = _consumerSurveyData.ServiceEngineerRating;
                }
                entObj.Attributes["hil_jobid"] = _erJob;

                _queryExp = new QueryExpression("hil_consumernpssurvey");
                _queryExp.ColumnSet = new ColumnSet("hil_submitstatus");
                _queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                _queryExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _erJob.Id);
                _entcoll = _service.RetrieveMultiple(_queryExp);

                if (_entcoll.Entities.Count > 0)
                {
                    if (_entcoll.Entities[0].GetAttributeValue<bool>("hil_submitstatus"))
                    {
                        _consumerNPSResult.Result = false;
                        _consumerNPSResult.ResultMessage = "Survey already Submitted";
                    }
                    else
                    {
                        entObj.Id = _entcoll.Entities[0].Id;
                        _service.Update(entObj);
                        _consumerNPSResult.Result = true;
                        _consumerNPSResult.ResultMessage = "Success";
                    }
                }
                else
                {
                    _service.Create(entObj);
                    _consumerNPSResult.Result = true;
                    _consumerNPSResult.ResultMessage = "Success";
                }
            }
            catch (Exception ex)
            {

            }
        }
    }

    public class ConsumerSurveyDTO {
        public string JobId { get; set; }
        public string CustomerName { get; set; }
        public string BrandName { get; set; }
        public string NPSValue { get; set; }
        public string ServiceEngineerRating { get; set; }
        public string DetractorsResponse { get; set; }
        public string PassivesResponse { get; set; }
        public string PromotersResponse { get; set; }
        public string Feedback { get; set; }
        public bool SubmitStatus { get; set; }
        public bool Result { get; set; }
        public string ResultMessage { get; set; }
    }
}
