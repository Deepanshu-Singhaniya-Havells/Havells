using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class BotSpeechTransScript
    {
        [DataMember]
        public string Transcript { get; set; }
        [DataMember]
        public string MobileNo { get; set; }
        [DataMember]
        public string ConsumerGuid { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }

        public BotSpeechTransScript InsertBotSpeechTransScript(BotSpeechTransScript reqParams)
        {
            BotSpeechTransScript _retObj = null;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (string.IsNullOrWhiteSpace(reqParams.MobileNo))
                    {
                        _retObj = new BotSpeechTransScript() { StatusCode = "204", StatusDescription = "Mobile Number is required." };
                        return _retObj;
                    }
                    if (string.IsNullOrWhiteSpace(reqParams.Transcript))
                    {
                        _retObj = new BotSpeechTransScript() { StatusCode = "204", StatusDescription = "Transcript is required." };
                        return _retObj;
                    }
                    else {
                        reqParams.Transcript.Replace("'s", "`s").Replace("'", "\"");
                    }
                    try
                    {
                        QueryExpression query = new QueryExpression("contact");
                        query.ColumnSet = new ColumnSet(false);
                        query.Distinct = true;
                        query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, reqParams.MobileNo);
                        EntityCollection _ecolConsumer = service.RetrieveMultiple(query);
                        if (_ecolConsumer.Entities.Count == 0)
                        {
                            return new BotSpeechTransScript() { StatusCode = "204", StatusDescription = "Mobile Number does not exist." };
                        }
                        else
                        {
                            Entity obj_task = new Entity("hil_speechtranscript");
                            obj_task.Attributes["description"] =   reqParams.Transcript;
                            obj_task.Attributes["subject"] = "Bot Speech Transcript dt: " + DateTime.Now.AddMinutes(330).ToString();
                            obj_task.Attributes["regardingobjectid"] = _ecolConsumer.Entities[0].ToEntityReference();
                            service.Create(obj_task);
                            _retObj = new BotSpeechTransScript() {ConsumerGuid= _ecolConsumer.Entities[0].Id.ToString(), MobileNo = reqParams.MobileNo, Transcript = reqParams.Transcript, StatusCode = "200", StatusDescription = "Success" };
                            return _retObj;
                        }
                    }
                    catch (Exception ex)
                    {
                        _retObj = new BotSpeechTransScript() { MobileNo = reqParams.MobileNo, StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                        return _retObj;
                    }
                }
                else
                {
                    _retObj = new BotSpeechTransScript() { MobileNo = reqParams.MobileNo, StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                _retObj = new BotSpeechTransScript() { MobileNo = reqParams.MobileNo, StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return _retObj;
        }
    }
    [DataContract]
    public class CallbackRequest
    {
        [DataMember]
        public string MobileNumber { get; set; }
        [DataMember]
        public string CallDescription { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }

        public CallbackRequest InsertCallbackRequest(CallbackRequest reqParams)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            CallbackRequest _retObj = null;
            try
            {
                if (service != null)
                {
                    if (reqParams.MobileNumber.Trim().Length == 0)
                    {
                        _retObj = new CallbackRequest() { StatusCode = "204", StatusDescription = "Consumer Mobile Number is required." };
                        return _retObj;
                    }
                    try
                    {
                        EntityReference erConsumer = null;
                        QueryExpression query = new QueryExpression("contact");
                        query.ColumnSet = new ColumnSet(false);
                        query.Distinct = true;
                        query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, reqParams.MobileNumber);
                        EntityCollection _ecolConsumer = service.RetrieveMultiple(query);
                        if (_ecolConsumer.Entities.Count == 0)
                        {
                            Entity entConsumer = new Entity("contact");
                            entConsumer["mobilephone"] = reqParams.MobileNumber;
                            entConsumer["firstname"] = "Call back-" + reqParams.MobileNumber;
                            entConsumer["hil_consumersource"] = new OptionSetValue(9); // Voice Bot
                            erConsumer = new EntityReference("contact", service.Create(entConsumer));
                        }
                        else {
                            erConsumer = _ecolConsumer.Entities[0].ToEntityReference();
                        }
                        Entity obj_task = new Entity("hil_callbackrequest");
                        obj_task.Attributes["subject"] = "Call back request -" + reqParams.MobileNumber;
                        if (reqParams.CallDescription != null && reqParams.CallDescription != string.Empty)
                        {
                            obj_task.Attributes["description"] = reqParams.CallDescription;
                        }
                        obj_task.Attributes["regardingobjectid"] = new EntityReference("contact", erConsumer.Id);
                        Guid CallbackRequestId = service.Create(obj_task);
                        _retObj = new CallbackRequest() {CallDescription= reqParams.CallDescription, MobileNumber = reqParams.MobileNumber, StatusCode = "200", StatusDescription = "Success" };
                        return _retObj;
                    }
                    catch (Exception ex)
                    {
                        _retObj = new CallbackRequest() { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                        return _retObj;
                    }
                }
                else
                {
                    _retObj = new CallbackRequest() { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                _retObj = new CallbackRequest() { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return _retObj;
        }

    }
    [DataContract]
    public class Escalations
    {
        [DataMember]
        public string JobId { get; set; }
        [DataMember]
        public List<EscalationsReqRes> EscalationDetails { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }

        public Escalations GetEscalations(Escalations JobData)
        {
            Escalations objEscalations = null;
            EntityCollection entcoll;
            IOrganizationService service = ConnectToCRM.GetOrgService();
            try
            {
                QueryExpression query = new QueryExpression("msdyn_workorder");
                query.ColumnSet = new ColumnSet(false);
                query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, JobData.JobId);
                EntityCollection _ecolJob = service.RetrieveMultiple(query);
                if (_ecolJob.Entities.Count == 0)
                {
                    objEscalations = new Escalations { StatusCode = "204", StatusDescription = "Job Id doesn't exist." };
                    return objEscalations;
                }
                else
                {
                    string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='task'>
                                <attribute name='subject' />
                                <attribute name='prioritycode' />
                                <attribute name='createdby' />
                                <attribute name='activityid' />
                                <attribute name='hil_tasktype' />
                                <attribute name='description' />
                                <attribute name='createdon' />
                                <order attribute='subject' descending='false' />
                                <filter type='and'>
                                  <condition attribute='regardingobjectid' operator='eq' value='{_ecolJob.Entities[0].Id}' />
                                </filter>
                              </entity>
                            </fetch>";

                    entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (entcoll.Entities.Count == 0)
                    {
                        objEscalations = new Escalations {JobId= JobData.JobId, StatusCode = "204", StatusDescription = "No Escalations found against given Job ID" };
                        return objEscalations;
                    }
                    else
                    {
                        List<EscalationsReqRes> lstExacalationsResponse = new List<EscalationsReqRes>();

                        foreach (Entity ent in entcoll.Entities)
                        {
                            EscalationsReqRes objEscalationResponse = new EscalationsReqRes();
                            objEscalationResponse.JobId = JobData.JobId;
                            objEscalationResponse.Subject = ent.GetAttributeValue<string>("subject");
                            if (ent.Attributes.Contains("createdby"))
                            {
                                objEscalationResponse.CreatedBy = ent.GetAttributeValue<EntityReference>("createdby").Name;
                            }

                            if (ent.Attributes.Contains("hil_tasktype"))
                            {
                                objEscalationResponse.TaskType = ent.FormattedValues["hil_tasktype"].ToString();
                            }

                            objEscalationResponse.CreatedOn = ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString("dd-mm-yyyy hh:mm:ss");
                            objEscalationResponse.Description = ent.GetAttributeValue<string>("description");
                            lstExacalationsResponse.Add(objEscalationResponse);
                        }
                        objEscalations = new Escalations { EscalationDetails = lstExacalationsResponse,JobId= JobData.JobId, StatusCode = "200", StatusDescription = "Success" };
                        return objEscalations;
                    }
                }
            }
            catch (Exception ex)
            {
                objEscalations = new Escalations { JobId = JobData.JobId, StatusCode = "500", StatusDescription = ex.Message };
                return objEscalations;
            }
        }

        public Escalations InsertEscalations(EscalationsReqRes reqParams)
        {
            Escalations _retObj = null;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (VoiceBot.IsNull(reqParams.JobId))
                    {
                        _retObj = new Escalations() {JobId= reqParams.JobId, StatusCode = "204", StatusDescription = "Job Id is required." };
                        return _retObj;
                    }
                    if (VoiceBot.IsNull(reqParams.TaskType))
                    {
                        _retObj = new Escalations() { JobId = reqParams.JobId, StatusCode = "204", StatusDescription = "Escalation Type is required." };
                        return _retObj;
                    }
                    else if (reqParams.TaskType!="1" && reqParams.TaskType != "2")
                    {
                        _retObj = new Escalations() { JobId = reqParams.JobId, StatusCode = "204", StatusDescription = "Invalid escalation type. Valid Values will be {1-Reminder,2-Escalation}" };
                        return _retObj;
                    }
                    if (VoiceBot.IsNull(reqParams.Description))
                    {
                        _retObj = new Escalations() { JobId = reqParams.JobId, StatusCode = "204", StatusDescription = "Escalation description is required." };
                        return _retObj;
                    }
                    try
                    {
                        EntityReference _efJob = null;

                        QueryExpression query = new QueryExpression("msdyn_workorder");
                        query.ColumnSet = new ColumnSet(false);
                        query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, reqParams.JobId);
                        EntityCollection _ecolJob = service.RetrieveMultiple(query);
                        if (_ecolJob.Entities.Count == 0)
                        {
                            _retObj = new Escalations() { JobId = reqParams.JobId, StatusCode = "204", StatusDescription = "Job Id doesn't exist."};
                            return _retObj;
                        }
                        else {
                            _efJob = _ecolJob.Entities[0].ToEntityReference();
                        }
                        string _subject = string.Empty;
                        query = new QueryExpression("task");
                        query.ColumnSet = new ColumnSet(false);
                        query.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, _efJob.Id);
                        query.Criteria.AddCondition("hil_tasktype", ConditionOperator.Equal, Convert.ToInt32(reqParams.TaskType));
                        _ecolJob = service.RetrieveMultiple(query);
                        if (_ecolJob.Entities.Count == 0) {
                            _subject = reqParams.TaskType == "1" ? "Reminder 1" : "Escalation 1";
                        }
                        else {
                            _subject = (reqParams.TaskType == "1" ? "Reminder " : "Escalation ").ToString() + (_ecolJob.Entities.Count + 1).ToString();
                        }

                        Entity obj_task = new Entity("task");
                        obj_task.Attributes["subject"] = _subject;
                        obj_task.Attributes["prioritycode"] = new OptionSetValue(2); // High
                        obj_task.Attributes["hil_tasktype"] = new OptionSetValue(Convert.ToInt32(reqParams.TaskType));
                        if (!VoiceBot.IsNull(reqParams.Description))
                            obj_task.Attributes["description"] = reqParams.Description;

                        obj_task.Attributes["regardingobjectid"] = _efJob;
                        Guid JobId = service.Create(obj_task);
                        _retObj = new Escalations() { JobId = reqParams.JobId, StatusCode = "200", StatusDescription = "Success"};
                        return _retObj;
                    }
                    catch (Exception ex)
                    {
                        _retObj = new Escalations() { JobId = reqParams.JobId, StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                        return _retObj;
                    }
                }
                else
                {
                    _retObj = new Escalations() { JobId = reqParams.JobId, StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                _retObj = new Escalations() { JobId = reqParams.JobId, StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return _retObj;
        }
    }
    [DataContract]
    public class ProductDivision
    {
        [DataMember]
        public List<ProductHierarchyResponse> ProductDivisionDetails { get; set; }

        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }

        public ProductDivision GetProductDivisions()
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            ProductDivision _objProductDivision = null;
            EntityCollection entcoll;
            try
            {
                string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='hil_whatsappproductdivisionconfig'>
                    <attribute name='hil_productmaterialgroup' />
                    <attribute name='hil_productdivision' />
                    <attribute name='hil_whatsappproductdivisionconfigid' />
                    <order attribute='hil_productdivision' descending='false' />
                    <filter type='and'>
                        <condition attribute='statecode' operator='eq' value='0' />
                        <condition attribute='hil_productdivision' operator='not-null' />
                        <condition attribute='hil_productmaterialgroup' operator='not-null' />
                    </filter>
                    <link-entity name='product' from='productid' to='hil_productdivision' link-type='inner' alias='ab'>
                        <filter type='and'>
                            <condition attribute='hil_brandidentifier' operator='eq' value='2' />
                        </filter>
                    </link-entity>
                    </entity>
                    </fetch>";

                entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                if (entcoll.Entities.Count == 0)
                {
                    _objProductDivision = new ProductDivision { StatusCode = "204", StatusDescription = "No Product divisions found." };
                    return _objProductDivision;
                }
                else
                {
                    List<ProductHierarchyResponse> lstResponse = new List<ProductHierarchyResponse>();

                    foreach (Entity ent in entcoll.Entities)
                    {
                        ProductHierarchyResponse ObjResponse = new ProductHierarchyResponse();

                        ObjResponse.CategoryId = ent.GetAttributeValue<EntityReference>("hil_productdivision").Id.ToString();
                        ObjResponse.CategoryName = ent.GetAttributeValue<EntityReference>("hil_productdivision").Name;

                        lstResponse.Add(new ProductHierarchyResponse()
                        {
                            CategoryId = ent.GetAttributeValue<EntityReference>("hil_productdivision").Id.ToString(),
                            CategoryName = ent.GetAttributeValue<EntityReference>("hil_productdivision").Name,
                            SubCategoryId = ent.GetAttributeValue<EntityReference>("hil_productmaterialgroup").Id.ToString(),
                            SubCategoryName = ent.GetAttributeValue<EntityReference>("hil_productmaterialgroup").Name
                        });
                    }
                    _objProductDivision = new ProductDivision { ProductDivisionDetails = lstResponse, StatusCode = "200", StatusDescription = "Success" };
                    return _objProductDivision;
                }
            }
            catch (Exception ex)
            {
                _objProductDivision = new ProductDivision { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message };
                return _objProductDivision;
            }
        }
    }
    [DataContract]
    public class ProductHierarchyResponse
    {
        [DataMember]
        public string CategoryId { get; set; }
        [DataMember]
        public string CategoryName { get; set; }
        [DataMember]
        public string SubCategoryId { get; set; }
        [DataMember]
        public string SubCategoryName { get; set; }

    }
    [DataContract]
    public class EscalationsReqRes
    {
        [DataMember]
        public string JobId { get; set; }
        [DataMember]
        public string Subject { get; set; }
        [DataMember]
        public string CreatedBy { get; set; }
        [DataMember]
        public string TaskType { get; set; }
        [DataMember]
        public string CreatedOn { get; set; }
        [DataMember]
        public string Description { get; set; }
    }

    public class VoiceBot {
        public static bool IsNull(Object _value)
        {
            return _value == null ? true : _value.ToString() == string.Empty ? true : false;
        }
    }
}
