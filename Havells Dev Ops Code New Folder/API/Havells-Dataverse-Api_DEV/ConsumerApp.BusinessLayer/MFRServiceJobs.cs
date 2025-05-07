using Microsoft.WindowsAzure.Storage.Blob.Protocol;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Text.RegularExpressions;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class MFRServiceJobs
    {
        public JobRequestDTO CreateServiceCallRequest(JobRequestDTO _jobRequest)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            Guid businessmappingId = Guid.Empty;
            string[] call_type = new string[] { "INSTALLATION", "BREAKDOWN" };

            try
            {
                if (service != null)
                {
                    if (string.IsNullOrWhiteSpace(_jobRequest.customer_mobileno))
                    {
                        _jobRequest.status_code = "204";
                        _jobRequest.status_description = "Customer Mobile Number is required.";
                        return _jobRequest;
                    }
                    else if (!Regex.IsMatch(_jobRequest.customer_mobileno, @"^[6-9]\d{9}$"))
                    {
                        _jobRequest.status_code = "204";
                        _jobRequest.status_description = "Incorrect Mobile Number";
                        return _jobRequest;
                    }
                    if (string.IsNullOrWhiteSpace(_jobRequest.customer_firstname))
                    {
                        _jobRequest.status_code = "204";
                        _jobRequest.status_description = "Customer First Name is required.";
                        return _jobRequest;
                    }
                    if (string.IsNullOrWhiteSpace(_jobRequest.address_line1))
                    {
                        _jobRequest.status_code = "204";
                        _jobRequest.status_description = "Address Line 1 is required.";
                        return _jobRequest;
                    }
                    if (string.IsNullOrWhiteSpace(_jobRequest.pincode))
                    {
                        _jobRequest.status_code = "204";
                        _jobRequest.status_description = "Pincode is required.";
                        return _jobRequest;
                    }
                    else if (!Regex.IsMatch(_jobRequest.pincode, @"^\d{6}$"))
                    {
                        _jobRequest.status_code = "204";
                        _jobRequest.status_description = "Invalid Pincode.";
                        return _jobRequest;
                    }
                    else
                    {
                        QueryExpression query = new QueryExpression("hil_pincode");
                        query.ColumnSet = new ColumnSet(false);
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, _jobRequest.pincode);
                        query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        EntityCollection entcollpincode = service.RetrieveMultiple(query);
                        if (entcollpincode.Entities.Count > 0)
                        {
                            query = new QueryExpression("hil_businessmapping");
                            query.TopCount = 1;
                            query.ColumnSet.AddColumns("hil_businessmappingid", "hil_pincode");
                            query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, entcollpincode.Entities[0].Id);
                            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            EntityCollection businessmapping = service.RetrieveMultiple(query);

                            if (businessmapping.Entities.Count > 0)
                            {
                                businessmappingId = businessmapping.Entities[0].Id;
                            }
                        }
                        else
                        {
                            _jobRequest.status_code = "204";
                            _jobRequest.status_description = "Not a valid Pincode.";
                            return _jobRequest;
                        }
                    }
                    if (string.IsNullOrWhiteSpace(_jobRequest.call_type))
                    {
                        _jobRequest.status_code = "204";
                        _jobRequest.status_description = "Call type is required.";
                        return _jobRequest;
                    }
                    else
                    {
                        if (!call_type.Contains(_jobRequest.call_type.ToUpper()))
                        {
                            _jobRequest.status_code = "204";
                            _jobRequest.status_description = "Call type is not valid.";
                            return _jobRequest;
                        }
                    }
                    if (string.IsNullOrWhiteSpace(_jobRequest.product_subcategory))
                    {
                        _jobRequest.status_code = "204";
                        _jobRequest.status_description = "Product Subcategory is required.";
                        return _jobRequest;
                    }
                    if (string.IsNullOrWhiteSpace(_jobRequest.caller_type))
                    {
                        _jobRequest.status_description = "Caller Type is required.";
                        _jobRequest.status_code = "204";
                        return _jobRequest;
                    }
                    else
                    {
                        if (_jobRequest.caller_type.ToUpper() != "DEALER")
                        {
                            _jobRequest.status_description = "Invalid Caller Type.";
                            _jobRequest.status_code = "204";
                            return _jobRequest;
                        }
                    }
                    if (string.IsNullOrWhiteSpace(_jobRequest.dealer_code))
                    {
                        _jobRequest.status_code = "204";
                        _jobRequest.status_description = "Dealer Code is required.";
                        return _jobRequest;
                    }
                    DateTime preferreddate;
                    if (string.IsNullOrWhiteSpace(_jobRequest.expected_delivery_date))
                    {
                        _jobRequest.status_code = "204";
                        _jobRequest.status_description = "Expected Delivery Date is required.";
                        return _jobRequest;
                    }
                    else
                    {
                        if (!DateTime.TryParseExact(_jobRequest.expected_delivery_date, "MM-dd-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out preferreddate))
                        {
                            _jobRequest.status_code = "204";
                            _jobRequest.status_description = "Expected Delivery Date is not in the correct format (MM-dd-yyyy)";
                            return _jobRequest;
                        }
                    }
                    #region Create Service Call
                    Entity enWorkorder = new Entity("msdyn_workorder");

                    #region Customer_Creation
                    Guid contactId = Guid.Empty;
                    QueryExpression Query = new QueryExpression("contact");
                    Query.ColumnSet = new ColumnSet("fullname", "emailaddress1", "mobilephone");
                    Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, _jobRequest.customer_mobileno);
                    EntityCollection entcoll = service.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count > 0)
                    {
                        contactId = entcoll.Entities[0].Id;
                        enWorkorder["hil_customername"] = entcoll.Entities[0].Contains("fullname") ? entcoll.Entities[0].GetAttributeValue<string>("fullname") : "";
                        enWorkorder["hil_mobilenumber"] = entcoll.Entities[0].Contains("mobilephone") ? entcoll.Entities[0].GetAttributeValue<string>("mobilephone") : "";
                        enWorkorder["hil_email"] = entcoll.Entities[0].Contains("emailaddress1") ? entcoll.Entities[0].GetAttributeValue<string>("emailaddress1") : "";
                    }
                    else
                    {
                        Entity entConsumer = new Entity("contact");
                        entConsumer["mobilephone"] = _jobRequest.customer_mobileno;
                        entConsumer["firstname"] = _jobRequest.customer_firstname;
                        entConsumer["lastname"] = _jobRequest.customer_lastname ?? "";
                        entConsumer["address1_telephone3"] = _jobRequest.alternate_number ?? "";
                        entConsumer["hil_consumersource"] = new OptionSetValue(21); //MFR
                        entConsumer["hil_subscribeformessagingservice"] = true;
                        contactId = service.Create(entConsumer);
                    }
                    #endregion

                    if (contactId != Guid.Empty)
                    {
                        enWorkorder["hil_customerref"] = new EntityReference("contact", contactId);
                        enWorkorder["hil_mobilenumber"] = _jobRequest.customer_mobileno;
                        enWorkorder["hil_alternate"] = _jobRequest.alternate_number;
                    }
                    if (!string.IsNullOrWhiteSpace(_jobRequest.expected_delivery_date))
                    {
                        enWorkorder["hil_preferreddate"] = preferreddate;
                    }

                    #region Address_Creation
                    Entity address = new Entity("hil_address");
                    address["hil_customer"] = new EntityReference("contact", contactId);
                    address["hil_street1"] = _jobRequest.address_line1;
                    address["hil_street2"] = _jobRequest.address_line2;
                    address["hil_street3"] = _jobRequest.landmark;
                    address["hil_addresstype"] = new OptionSetValue(1);
                    address["hil_businessgeo"] = new EntityReference("hil_businessmapping", businessmappingId);
                    Guid Addressid = service.Create(address);
                    #endregion

                    if (Addressid != Guid.Empty)
                    {
                        enWorkorder["hil_address"] = new EntityReference("hil_address", Addressid);
                    }
                    Query = new QueryExpression("product");
                    Query.ColumnSet = new ColumnSet("hil_division");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("name", ConditionOperator.Equal, _jobRequest.product_subcategory);
                    Query.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 3);
                    EntityCollection entCol = service.RetrieveMultiple(Query);
                    if (entCol.Entities.Count > 0)
                    {
                        enWorkorder["hil_productsubcategory"] = new EntityReference("product", entCol.Entities[0].Id);

                        Query = new QueryExpression("hil_stagingdivisonmaterialgroupmapping");
                        Query.ColumnSet = new ColumnSet("hil_stagingdivisonmaterialgroupmappingid");
                        Query.Criteria = new FilterExpression(LogicalOperator.And);
                        Query.Criteria.AddCondition("hil_productcategorydivision", ConditionOperator.Equal, entCol.Entities[0].GetAttributeValue<EntityReference>("hil_division").Id);
                        Query.Criteria = new FilterExpression(LogicalOperator.And);
                        Query.Criteria.AddCondition("hil_productsubcategorymg", ConditionOperator.Equal, entCol.Entities[0].Id);
                        EntityCollection ec = service.RetrieveMultiple(Query);
                        if (ec.Entities.Count > 0)
                        {
                            enWorkorder["hil_productcatsubcatmapping"] = ec.Entities[0].ToEntityReference();
                        }
                        string _fetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='hil_natureofcomplaint'>
                                        <attribute name='hil_name' />
                                        <attribute name='hil_natureofcomplaintid' />
                                        <attribute name='hil_callsubtype' />
                                        <order attribute='hil_name' descending='false' />
                                        <filter type='and'>
                                            <condition attribute='hil_relatedproduct' operator='eq' value='{entCol.Entities[0].Id}' />
                                            <condition attribute='statecode' operator='eq' value='0' />
                                        </filter>
                                        <link-entity name='hil_callsubtype' from='hil_callsubtypeid' to='hil_callsubtype' link-type='inner' alias='ad'>
                                        <filter type='and'>
                                            <condition attribute='hil_name' operator='eq' value='{_jobRequest.call_type}' />
                                        </filter>
                                        </link-entity>
                                        </entity>
                                        </fetch>";

                        EntityCollection _natureofcomplaintColl = service.RetrieveMultiple(new FetchExpression(_fetchQuery));
                        if (_natureofcomplaintColl.Entities.Count > 0)
                        {
                            enWorkorder["hil_natureofcomplaint"] = _natureofcomplaintColl.Entities[0].ToEntityReference();
                            enWorkorder["hil_callsubtype"] = _natureofcomplaintColl.Entities[0].GetAttributeValue<EntityReference>("hil_callsubtype");
                        }
                    }
                    enWorkorder["hil_consumertype"] = new EntityReference("hil_consumertype", new Guid("484897de-2abd-e911-a957-000d3af0677f")); //B2C
                    enWorkorder["hil_consumercategory"] = new EntityReference("hil_consumercategory", new Guid("baa52f5e-2bbd-e911-a957-000d3af0677f")); //End User
                    enWorkorder["hil_quantity"] = 1;
                    enWorkorder["hil_customercomplaintdescription"] = _jobRequest.chief_complaint;
                    enWorkorder["hil_callertype"] = new OptionSetValue(910590000);//  Dealer
                    enWorkorder["hil_newserialnumber"] = _jobRequest.dealer_code;
                    enWorkorder["hil_sourceofjob"] = new OptionSetValue(19); //MFR

                    enWorkorder["msdyn_serviceaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {ServiceAccount:"Dummy Account"}
                    enWorkorder["msdyn_billingaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {BillingAccount:"Dummy Account"}

                    Guid serviceCallGuid = service.Create(enWorkorder);
                    if (serviceCallGuid != Guid.Empty)
                    {
                        _jobRequest.job_id = service.Retrieve("msdyn_workorder", serviceCallGuid, new ColumnSet("msdyn_name")).GetAttributeValue<string>("msdyn_name");
                        _jobRequest.status_code = "200";
                        _jobRequest.status_description = "Service call request registered successfully.";
                    }
                    else
                    {
                        _jobRequest.status_code = "204";
                        _jobRequest.status_description = "FAILURE !!! Something went wrong";
                    }
                    #endregion
                }
                else
                {
                    _jobRequest.status_code = "503";
                    _jobRequest.status_description = "D365 Service Unavailable.";
                    return _jobRequest;
                }
            }
            catch (Exception ex)
            {
                _jobRequest.status_code = "500";
                _jobRequest.status_description = "D365 Internal Server Error : " + ex.Message.ToUpper();
                return _jobRequest;
            }
            return _jobRequest;
        }
        public JobStatusDTO GetJobstatus(JobStatusDTO _jobRequest)
        {
            _jobRequest.status_code = "200";
            _jobRequest.status_description = "OK.";

            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (string.IsNullOrWhiteSpace(_jobRequest.job_id))
                    {
                        _jobRequest.status_code = "204";
                        _jobRequest.status_description = "Invalid is required.";
                        return _jobRequest;
                    }
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
                    }
                    else
                    {
                        _jobRequest.status_code = "200";
                        _jobRequest.status_description = "Invalid Job Id.";
                        return _jobRequest;
                    }
                }
                else
                {
                    _jobRequest.status_code = "503";
                    _jobRequest.status_description = "D365 Service Unavailable.";
                    return _jobRequest;
                }
            }
            catch (Exception ex)
            {
                _jobRequest.status_code = "500";
                _jobRequest.status_description = "D365 Internal Server Error : " + ex.Message.ToUpper();
                return _jobRequest;
            }
            return _jobRequest;
        }
        public WorkOrderResponse GetWorkOrdersStatus(WorkOrderRequest objreq)
        {
            try
            {
                WorkOrderResponse workOrderResponse = new WorkOrderResponse();
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (string.IsNullOrWhiteSpace(objreq.DealerCode))
                {
                    return new WorkOrderResponse
                    {
                        StatusCode = 204,
                        Message = "Dealer Code required."
                    };
                }
                if (service == null)
                {
                    return new WorkOrderResponse
                    {
                        StatusCode = 503,
                        Message = "D365 Service Unavailable."
                    };
                }
                string fromDate = objreq.FromDate;
                string toDate = objreq.ToDate;
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
                            Createdon = entity.Contains("createdon") ? entity.GetAttributeValue<DateTime>("createdon").ToString() : "",
                            Productsubcategory = entity.Contains("hil_productsubcategory") ? entity.GetAttributeValue<EntityReference>("hil_productsubcategory").Name : "",
                            Productcategory = entity.Contains("hil_productcategory") ? entity.GetAttributeValue<EntityReference>("hil_productcategory").Name : "",
                            Callsubtype = entity.Contains("hil_callsubtype") ? entity.GetAttributeValue<EntityReference>("hil_callsubtype").Name : "",
                            Customer = entity.Contains("hil_customerref") ? entity.GetAttributeValue<EntityReference>("hil_customerref").Name : "",
                            Owner = entity.Contains("hil_owneraccount") ? entity.GetAttributeValue<EntityReference>("hil_owneraccount").Name : "",
                        };
                        workOrderInfos.Add(workOrderInfo);
                    }
                }
                return new WorkOrderResponse
                {
                    workOrderInfos = workOrderInfos,
                    StatusCode = 200,
                    Message = "success"
                };
            }
            catch (Exception ex)
            {
                return new WorkOrderResponse
                {
                    StatusCode = 500,
                    Message = "D365 Internal Server Error : " + ex.Message.ToUpper()
                };
            }
        }
    }

    [DataContract]
    public class JobRequestDTO
    {
        [DataMember]
        public string customer_firstname { get; set; }
        [DataMember]
        public string customer_lastname { get; set; }
        [DataMember]
        public string customer_mobileno { get; set; }
        [DataMember]
        public string alternate_number { get; set; }
        [DataMember]
        public string address_line1 { get; set; }
        [DataMember]
        public string address_line2 { get; set; }
        [DataMember]
        public string landmark { get; set; }
        [DataMember]
        public string pincode { get; set; }
        [DataMember]
        public string call_type { get; set; }
        [DataMember]
        public string product_subcategory { get; set; }
        [DataMember]
        public string caller_type { get; set; }
        [DataMember]
        public string dealer_code { get; set; }
        [DataMember]
        public string expected_delivery_date { get; set; }
        [DataMember]
        public string status_code { get; set; }
        [DataMember]
        public string status_description { get; set; }
        [DataMember]
        public string chief_complaint { get; set; }
        [DataMember]
        public string job_id { get; set; }

    }

    [DataContract]
    public class JobStatusDTO
    {
        [DataMember]
        public string job_id { get; set; }
        [DataMember]
        public string mobile_number { get; set; }
        [DataMember]
        public string serial_number { get; set; }
        [DataMember]
        public string product_category { get; set; }
        [DataMember]
        public string product_subcategory { get; set; }
        [DataMember]
        public string call_type { get; set; }
        [DataMember]
        public string customer_complaint { get; set; }
        [DataMember]
        public string assigned_resource { get; set; }
        [DataMember]
        public string job_substatus { get; set; }
        [DataMember]
        public string technician_remarks { get; set; }
        [DataMember]
        public string closed_on { get; set; }
        [DataMember]
        public string cancel_reason { get; set; }
        [DataMember]
        public string closure_remarks { get; set; }
        [DataMember]
        public string webclosure_remarks { get; set; }
        [DataMember]
        public List<JobProductDTO> spare_parts { get; set; }

        [DataMember]
        public string status_code { get; set; }
        [DataMember]
        public string status_description { get; set; }
    }

    [DataContract]
    public class JobProductDTO
    {
        [DataMember]
        public string index { get; set; }
        [DataMember]
        public string product_code { get; set; }
        [DataMember]
        public string product_description { get; set; }
    }

    [DataContract]
    public class WorkOrderResponse
    {
        [DataMember]
        public List<WorkOrderInfo> workOrderInfos { get; set; }
        [DataMember]
        public int StatusCode { get; set; }
        [DataMember]
        public string Message { get; set; }
    }

    [DataContract]
    public class WorkOrderInfo
    {
        [DataMember]
        public string JobId { get; set; }
        [DataMember]
        public string Substatus { get; set; }
        [DataMember]
        public string Createdon { get; set; }
        [DataMember]
        public string Productsubcategory { get; set; }
        [DataMember]
        public string Productcategory { get; set; }
        [DataMember]
        public string Callsubtype { get; set; }
        [DataMember]
        public string Customer { get; set; }
        [DataMember]
        public string Owner { get; set; }

    }

    [DataContract]
    public class WorkOrderRequest
    {
        [DataMember]
        public string DealerCode { get; set; }
        [DataMember]
        public string FromDate { get; set; }
        [DataMember]
        public string ToDate { get; set; }
    }
}
