using System;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class CommunicationNPS
    {
        [DataMember]
        public RequestDTO action { get; set; }

        public ResponseDTO SendCommunication(RequestDTO _data)
        {
            ResponseDTO _responseData = new ResponseDTO() { action_name = _data.name, data = new ResponseDataDTO() };
            if (_data.parameters.mobile_number == null || _data.parameters.mobile_number.Trim().Length == 0)
            {
                _responseData.data.error_code = "2";
                _responseData.data.error_description = "Customer's Mobile number is required.";
                return _responseData;
            }
            if (_data.parameters.message_body == null || _data.parameters.message_body.Trim().Length == 0)
            {
                _responseData.data.error_code = "2";
                _responseData.data.error_description = "Message Body is required.";
                return _responseData;
            }
            if (_data.name != "api_send_sms" && _data.name != "api_send_email")
            {
                _responseData.data.error_code = "2";
                _responseData.data.error_description = "Invalid action name.";
                return _responseData;
            }
            if (_data.parameters.cust_ref == null || _data.parameters.cust_ref.Trim().Length == 0)
            {
                _responseData.data.error_code = "2";
                _responseData.data.error_description = "Customer Reference is required.";
                return _responseData;
            }
            if (_data.name == "api_send_sms" && (_data.parameters.templateId == null || _data.parameters.templateId.Trim().Length == 0))
            {
                _responseData.data.error_code = "2";
                _responseData.data.error_description = "SMS Template Id is required.";
                return _responseData;
            }
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                EntityReference _entrefCust = null;
                QueryExpression query = new QueryExpression()
                {
                    EntityName = "contact",
                    ColumnSet = new ColumnSet(false)
                };
                FilterExpression filterExpression = new FilterExpression(LogicalOperator.And);
                filterExpression.Conditions.Add(new ConditionExpression("mobilephone", ConditionOperator.Equal, _data.parameters.mobile_number));
                query.Criteria.AddFilter(filterExpression);

                EntityCollection _entCol = service.RetrieveMultiple(query);
                if (_entCol.Entities.Count > 0)
                {
                    _entrefCust = _entCol.Entities[0].ToEntityReference();
                    if (_data.name == "api_send_sms")
                    {
                        try
                        {
                            query = new QueryExpression()
                            {
                                EntityName = "hil_smstemplates",
                                ColumnSet = new ColumnSet(false)
                            };
                            filterExpression = new FilterExpression(LogicalOperator.And);
                            filterExpression.Conditions.Add(new ConditionExpression("hil_templateid", ConditionOperator.Equal, _data.parameters.templateId));
                            query.Criteria.AddFilter(filterExpression);
                            _entCol = service.RetrieveMultiple(query);
                            if (_entCol.Entities.Count > 0)
                            {
                                Entity _ent = new Entity("hil_smsconfiguration");
                                _ent["hil_direction"] = new OptionSetValue(2);
                                _ent["subject"] = _data.parameters.subject;
                                _ent["hil_mobilenumber"] = _data.parameters.mobile_number;
                                _ent["hil_contact"] = _entrefCust;
                                _ent["hil_smstemplate"] = _entCol.Entities[0].ToEntityReference();
                                _ent["hil_message"] = _data.parameters.message_body;
                                _ent["regardingobjectid"] = _entrefCust;
                                _ent["hil_requesttype"] = "NPS";
                                _ent["hil_jobtype"] = new OptionSetValue(2);
                                _ent["activityadditionalparams"] = _data.parameters.cust_ref;
                                service.Create(_ent);
                                _responseData.data.data = "Success";
                            }
                            else
                            {
                                _responseData.data.error_code = "4";
                                _responseData.data.error_description = "SMS Template does not exist.";
                            }
                        }
                        catch (Exception ex)
                        {
                            _responseData.data.error_code = "3";
                            _responseData.data.error_description = "Error while sending SMS !!! " + ex.Message;
                        }
                    }
                    else if (_data.name == "api_send_email")
                    {
                        try
                        {
                            Console.WriteLine("sending email.");
                            Entity entEmail = new Entity("email");
                            entEmail["subject"] = _data.parameters.subject;
                            entEmail["description"] = _data.parameters.message_body;

                            Entity entTo = new Entity("activityparty");
                            entTo["partyid"] = _entrefCust;
                            Entity[] entToList = { entTo };
                            entEmail["to"] = entToList;

                            Entity entFrom = new Entity("activityparty");
                            entFrom["partyid"] = getSender("Havells Connect", service);
                            Entity[] entFromList = { entFrom };
                            entEmail["from"] = entFromList;

                            entEmail["regardingobjectid"] = _entrefCust;
                            Guid emailId = service.Create(entEmail);
                            SendEmailRequest sendEmailReq = new SendEmailRequest()
                            {
                                EmailId = emailId,
                                IssueSend = true
                            };
                            SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);
                            _responseData.data.data = "Success";
                        }
                        catch (Exception ex)
                        {
                            _responseData.data.error_code = "4";
                            _responseData.data.error_description = "Error while sending Email !!! " + ex.Message;
                        }
                    }
                }
                else
                {
                    _responseData.data.error_code = "3";
                    _responseData.data.error_description = "Mobile Number does not exist.";
                }
            }
            catch (Exception ex)
            {
                _responseData.data.error_code = "4";
                _responseData.data.error_description = "ERROR !!! " + ex.Message;
            }
            return _responseData;
        }

        public  EntityReference getSender(string queName, IOrganizationService service)
        {
            QueryExpression _queQuery = new QueryExpression("queue");
            _queQuery.ColumnSet = new ColumnSet("name");
            _queQuery.Criteria = new FilterExpression(LogicalOperator.And);
            _queQuery.Criteria.AddCondition("name", ConditionOperator.Equal, queName);
            EntityCollection queueColl = service.RetrieveMultiple(_queQuery);
            if (queueColl.Entities.Count == 1)
            {
                return queueColl[0].ToEntityReference();
            }
            else {

                return null;
            }
        }

    }
    [DataContract]
    public class RequestDTO
    {
        [DataMember]
        public string name { get; set; }
        [DataMember]
        public ParametersDTO parameters { get; set; }
    }

    [DataContract]
    public class ParametersDTO
    {
        [DataMember]
        public string subject { get; set; }
        [DataMember]
        public string mobile_number { get; set; }
        [DataMember]
        public string message_body { get; set; }
        [DataMember]
        public string templateId { get; set; }
        [DataMember]
        public string cust_ref { get; set; }
    }
    [DataContract]
    public class ResponseDTO
    {
        [DataMember]
        public string action_name { get; set; }
        [DataMember]
        public ResponseDataDTO data { get; set; }
    }

    [DataContract]
    public class ResponseDataDTO
    {
        [DataMember]
        public string data { get; set; }
        [DataMember]
        public string error_code { get; set; }
        [DataMember]
        public string error_description { get; set; }
    }

}
