using HavellsSync_Data.IManager;
using HavellsSync_ModelData.AMC;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.Consumer;
using HavellsSync_ModelData.EasyReward;
using HavellsSync_ModelData.ServiceAlaCarte;
using Microsoft.Extensions.Configuration;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System.Globalization;
using System.Net;
using System.Text;

namespace HavellsSync_Data.Manager
{
    public class ConsumerManager : IConsumerManager
    {
        private IConfiguration configuration;
        private ICrmService _CrmService;
        public ConsumerManager(ICrmService crmService, IConfiguration configuration)
        {
            Check.Argument.IsNotNull(nameof(crmService), crmService);
            _CrmService = crmService;
            this.configuration = configuration;
        }
        public async Task<ConsumerResponse> ConsumersAppRating(string MobileNumber, string SourceType, string Rating, string Review)
        {
            ConsumerResponse ConResponse = new ConsumerResponse();
            QueryExpression query;
            try
            {
                query = new QueryExpression("hil_consumerrating");
                query.ColumnSet = new ColumnSet(false);
                query.Criteria.AddCondition("hil_mobilenumber", ConditionOperator.Equal, MobileNumber);
                EntityCollection Info = _CrmService.RetrieveMultiple(query);
                if (Info.Entities.Count > 0)
                {
                    Entity RatingUpdate = new Entity("hil_consumerrating", Info.Entities[0].Id);
                    RatingUpdate["hil_rating"] = Convert.ToInt32(Rating);
                    RatingUpdate["hil_remarks"] = Review;
                    _CrmService.Update(RatingUpdate);
                    ConResponse.Response = "Success";

                }
                else
                {
                    Entity RatingAdd = new Entity("hil_consumerrating");
                    RatingAdd["hil_mobilenumber"] = MobileNumber;
                    RatingAdd["hil_rating"] = Convert.ToInt32(Rating);
                    RatingAdd["hil_remarks"] = Review;
                    _CrmService.Create(RatingAdd);
                    ConResponse.Response = "Success";
                }
            }
            catch (Exception ex)
            {
                ConResponse.Response = "Fail! Some Internal Issue";
                ex.Message.ToString();
            }
            return ConResponse;


        }
        public async Task<InvoiceResponse> InvoiceDetails(string FromDate, string ToDate, string OrderNumber)
        {
            InvoiceResponse InvResponse = new InvoiceResponse();
            QueryExpression query;
            try
            {

                if (!string.IsNullOrWhiteSpace(OrderNumber))
                {
                    string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='invoice'>
                                        <attribute name='name'/>
                                        <attribute name='customerid'/>
                                        <attribute name='statuscode'/>
                                        <attribute name='totalamount'/>
                                        <attribute name='invoiceid'/>
                                        <attribute name='hil_sapsonumber'/>
                                        <attribute name='hil_sapinvoicenumber'/>
                                        <attribute name='hil_sapsyncmessage'/>
                                        <attribute name='hil_orderid'/>
                                        <attribute name='createdon'/>
                                        <order attribute='name' descending='false'/>
                                    <filter type='and'>
                                        <condition attribute='hil_amcsellingsource' operator='eq' value='OneWebsite|22' />
                                        <condition attribute='hil_orderid' operator='like' value='{OrderNumber}%' />
                                    </filter>
                                    </entity>
                                    </fetch>";

                    EntityCollection Info1 = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (Info1.Entities.Count > 0)
                    {
                        List<InvoiceDetailsResponse> InvcResponse = new List<InvoiceDetailsResponse>();
                        foreach (Entity ent in Info1.Entities)
                        {
                            InvoiceDetailsResponse IDR = new InvoiceDetailsResponse();
                            string OrderId = ent.Contains("hil_orderid") ? ent.GetAttributeValue<string>("hil_orderid") : "";
                            if (!string.IsNullOrWhiteSpace(OrderId))
                            {
                                OrderId = OrderId.Split("-")[0];
                            }
                            IDR.OrderNumber = OrderId;
                            IDR.SAPOrderNumber = ent.Contains("hil_sapsonumber") ? ent.GetAttributeValue<string>("hil_sapsonumber") : null;
                            IDR.SAPInvoiceNumber = ent.Contains("hil_sapinvoicenumber") ? ent.GetAttributeValue<string>("hil_sapinvoicenumber") : null;
                            IDR.SAPSyncMessage = ent.Contains("hil_sapsyncmessage") ? ent.GetAttributeValue<string>("hil_sapsyncmessage") : null;
                            InvcResponse.Add(IDR);
                        }
                        InvResponse.Data = InvcResponse;
                        InvResponse.Response = "Success";

                    }
                }
                else if (!string.IsNullOrWhiteSpace(FromDate) && !string.IsNullOrWhiteSpace(ToDate))
                {
                    string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='invoice'>
                                        <attribute name='name'/>
                                        <attribute name='customerid'/>
                                        <attribute name='statuscode'/>
                                        <attribute name='totalamount'/>
                                        <attribute name='invoiceid'/>
                                        <attribute name='hil_sapsonumber'/>
                                        <attribute name='hil_sapinvoicenumber'/>
                                        <attribute name='hil_sapsyncmessage'/>
                                        <attribute name='hil_orderid'/>
                                        <attribute name='createdon'/>
                                        <order attribute='name' descending='false'/>
                                    <filter type='and'>
                                    <condition attribute='createdon' operator='on-or-after' value='{FromDate}'/>
                                    <condition attribute='createdon' operator='on-or-before' value='{ToDate}'/>
                                    <condition attribute='hil_amcsellingsource' operator='eq' value='OneWebsite|22' />                                   
                                    </filter>
                                    </entity>
                                    </fetch>";

                    EntityCollection Info1 = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (Info1.Entities.Count > 0)
                    {
                        List<InvoiceDetailsResponse> InvcResponse = new List<InvoiceDetailsResponse>();
                        foreach (Entity ent in Info1.Entities)
                        {
                            InvoiceDetailsResponse IDR = new InvoiceDetailsResponse();
                            string OrderId = ent.Contains("hil_orderid") ? ent.GetAttributeValue<string>("hil_orderid") : "";
                            if (!string.IsNullOrWhiteSpace(OrderId))
                            {
                                OrderId = OrderId.Split("-")[0];
                            }
                            IDR.OrderNumber = OrderId; IDR.SAPOrderNumber = ent.Contains("hil_sapsonumber") ? ent.GetAttributeValue<string>("hil_sapsonumber") : null;
                            IDR.SAPInvoiceNumber = ent.Contains("hil_sapinvoicenumber") ? ent.GetAttributeValue<string>("hil_sapinvoicenumber") : null;
                            IDR.SAPSyncMessage = ent.Contains("hil_sapsyncmessage") ? ent.GetAttributeValue<string>("hil_sapsyncmessage") : null;
                            InvcResponse.Add(IDR);
                        }
                        InvResponse.Data = InvcResponse;
                        InvResponse.Response = "Success";

                    }
                }
            }
            catch (Exception ex)
            {
                InvResponse.Response = "Fail Some Internal Issue";
                ex.Message.ToString();
            }
            return InvResponse;


        }


        public async Task<RequestStatus> PriceList(List<PriceListParam> objPricelist)
        {
            string fetchXml = "";
            try
            {
                foreach (var item in objPricelist)
                {

                    fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='product'>
                          <attribute name='name'/>
                          <attribute name='productnumber'/>
                          <attribute name='description'/>
                          <attribute name='statecode'/>
                          <attribute name='productstructure'/>
                          <attribute name='productid'/>
                          <attribute name='pricelevelid'/>
                          <order attribute='productnumber' descending='false'/>
                          <filter type='and'>
                          <condition attribute='name' operator='eq' value='{item.MATNR}'/>                                       
                          </filter>
                          </entity>
                          </fetch>";
                    EntityCollection Info = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (Info.Entities.Count > 0)
                    {
                        if (item.KSCHL == "ZPR0" || item.KSCHL == "ZPRO" || item.KSCHL == "ZDAM")
                        {
                            Entity PriceListitem = new Entity("hil_stagingpricingmapping");
                            PriceListitem["hil_name"] = item.MATNR;
                            PriceListitem["hil_price"] = Convert.ToInt32(Convert.ToDecimal(item.KBETR));
                            PriceListitem["hil_datestart"] = Convert.ToDateTime(item.DATAB);
                            PriceListitem["hil_dateend"] = Convert.ToDateTime(item.DATBI);
                            if (item.KSCHL == "ZPR0" || item.KSCHL == "ZPRO")
                            {
                                PriceListitem["hil_type"] = true;
                            }
                            else if (item.KSCHL == "ZDAM")
                            {
                                PriceListitem["hil_type"] = false;
                            }
                            PriceListitem["hil_salesoffice"] = item.KSCHL;
                            PriceListitem["hil_product"] = new EntityReference("product", Info.Entities[0].Id);
                            var obj = _CrmService.Create(PriceListitem);
                        }
                    }
                }
                return (new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.OK,
                    Message = CommonMessage.SuccessMsg,
                });
            }

            catch (Exception ex)
            {
                return (new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }

        public async Task<(WorkOrderResponse, RequestStatus)> GetWorkOrdersStatus(WorkOrderRequest objreq)
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

                EntityCollection results = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml));
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
                }

                return (workOrderResponse, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.OK,
                    Message = CommonMessage.SuccessMsg,
                });
            }
            catch (Exception ex)
            {
                return (workOrderResponse, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }

        public async Task<JobStatusDTO> GetJobstatus(JobStatusDTO _jobRequest)
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

                EntityCollection _jobDetailsColl = _CrmService.RetrieveMultiple(new FetchExpression(_fetchQuery));
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
                        EntityCollection JobIncDetailsColl = _CrmService.RetrieveMultiple(new FetchExpression(IncFetchQuery));
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
                    _jobRequest.status_description = "OK.";
                }
                else
                {
                    _jobRequest.status_code = "200";
                    _jobRequest.status_description = "Invalid Job Id.";                  
                    
                }
            }
            catch (Exception ex)
            {
                _jobRequest.status_code = "400";
                _jobRequest.status_description = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper();
               
            }

            return _jobRequest;
        }

        public async Task<JobRequestDTO> CreateServiceCallRequest(JobRequestDTO _jobRequest)
        {
            try
            {
                DateTime preferreddate;
                if (!DateTime.TryParseExact(_jobRequest.expected_delivery_date, "MM-dd-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out preferreddate))
                {
                    _jobRequest.status_code = "204";
                    _jobRequest.status_description = "Expected Delivery Date is not in the correct format (MM-dd-yyyy)";
                    return _jobRequest;
                }
                Guid businessmappingId = Guid.Empty;                
                QueryExpression query = new QueryExpression("hil_pincode");
                query.ColumnSet = new ColumnSet(false);
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, _jobRequest.pincode);
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                EntityCollection entcollpincode = _CrmService.RetrieveMultiple(query);
                if (entcollpincode.Entities.Count > 0)
                {
                    query = new QueryExpression("hil_businessmapping");
                    query.TopCount = 1;
                    query.ColumnSet.AddColumns("hil_businessmappingid", "hil_pincode");
                    query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, entcollpincode.Entities[0].Id);
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    EntityCollection businessmapping = _CrmService.RetrieveMultiple(query);

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
                #region Create Service Call
                Entity enWorkorder = new Entity("msdyn_workorder");

                #region Customer_Creation
                Guid contactId = Guid.Empty;
                QueryExpression Query = new QueryExpression("contact");
                Query.ColumnSet = new ColumnSet("fullname", "emailaddress1", "mobilephone");
                Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, _jobRequest.customer_mobileno);
                EntityCollection entcoll = _CrmService.RetrieveMultiple(Query);
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
                    contactId = _CrmService.Create(entConsumer);
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

                    enWorkorder["hil_preferreddate"] = preferreddate; // Convert.ToDateTime(_jobRequest.expected_delivery_date);
                }

                #region Address_Creation
                Entity address = new Entity("hil_address");
                address["hil_customer"] = new EntityReference("contact", contactId);
                address["hil_street1"] = _jobRequest.address_line1;
                address["hil_street2"] = _jobRequest.address_line2;
                address["hil_street3"] = _jobRequest.landmark;
                address["hil_addresstype"] = new OptionSetValue(1);
                address["hil_businessgeo"] = new EntityReference("hil_businessmapping", businessmappingId);
                Guid Addressid = _CrmService.Create(address);
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
                EntityCollection entCol = _CrmService.RetrieveMultiple(Query);
                if (entCol.Entities.Count > 0)
                {
                    enWorkorder["hil_productsubcategory"] = new EntityReference("product", entCol.Entities[0].Id);

                    Query = new QueryExpression("hil_stagingdivisonmaterialgroupmapping");
                    Query.ColumnSet = new ColumnSet("hil_stagingdivisonmaterialgroupmappingid");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_productcategorydivision", ConditionOperator.Equal, entCol.Entities[0].GetAttributeValue<EntityReference>("hil_division").Id);
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_productsubcategorymg", ConditionOperator.Equal, entCol.Entities[0].Id);
                    EntityCollection ec = _CrmService.RetrieveMultiple(Query);
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

                    EntityCollection _natureofcomplaintColl = _CrmService.RetrieveMultiple(new FetchExpression(_fetchQuery));
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

                Guid serviceCallGuid = _CrmService.Create(enWorkorder);
                if (serviceCallGuid != Guid.Empty)
                {
                    _jobRequest.job_id = _CrmService.Retrieve("msdyn_workorder", serviceCallGuid, new ColumnSet("msdyn_name")).GetAttributeValue<string>("msdyn_name");
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
            catch (Exception ex)
            {
                _jobRequest.status_code = "500";
                _jobRequest.status_description = "D365 Internal Server Error : " + ex.Message.ToUpper();

            }

            return _jobRequest;
        }

        public async Task<List<JobOutput>> GetJobs(Job job)
        {            
            List<JobOutput> jobList = new List<JobOutput>();

            if (String.IsNullOrWhiteSpace(job.MobileNumber))
            {
                return jobList;
            }           

            QueryExpression query = new QueryExpression()
            {
                EntityName = "msdyn_workorder",
                ColumnSet = new ColumnSet(true)
            };
            FilterExpression filterExpression = new FilterExpression(LogicalOperator.And);
            if (!String.IsNullOrEmpty(job.Job_ID))
            {
                filterExpression.Conditions.Add(new ConditionExpression("msdyn_name", ConditionOperator.Equal, job.Job_ID));
            }
                filterExpression.Conditions.Add(new ConditionExpression("hil_mobilenumber", ConditionOperator.Equal, job.MobileNumber));

            query.Criteria.AddFilter(filterExpression);

            EntityCollection collection = _CrmService.RetrieveMultiple(query);

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
                        Entity ec = _CrmService.Retrieve("msdyn_customerasset", item.GetAttributeValue<EntityReference>("msdyn_customerasset").Id, new ColumnSet("hil_modelname"));
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

        public async Task<List<IoTServiceCallResult>> IoTGetServiceCalls(IotServiceCall job)
        {
            List<IoTServiceCallResult> list = new List<IoTServiceCallResult>();
            try
            {
                IoTServiceCallResult item;
                if (job.CustomerGuid.ToString().Trim().Length == 0)
                {
                    item = new IoTServiceCallResult
                    {
                        StatusCode = "204",
                        StatusDescription = "Customer GUID is required."
                    };
                    list.Add(item);
                    return list;
                }

               
                if (_CrmService != null)
                {
                    QueryExpression queryExpression = new QueryExpression
                    {
                        EntityName = "msdyn_workorder",
                        ColumnSet = new ColumnSet(allColumns: true)
                    };
                    FilterExpression filterExpression = new FilterExpression(LogicalOperator.And);
                    filterExpression.Conditions.Add(new ConditionExpression("hil_customerref", ConditionOperator.Equal, job.CustomerGuid));
                    queryExpression.Criteria.AddFilter(filterExpression);
                    queryExpression.AddOrder("createdon", OrderType.Descending);
                    EntityCollection entityCollection = _CrmService.RetrieveMultiple(queryExpression);
                    if (entityCollection.Entities != null && entityCollection.Entities.Count > 0)
                    {
                        foreach (Entity entity2 in entityCollection.Entities)
                        {
                            IoTServiceCallResult ioTServiceCallResult = new IoTServiceCallResult();
                            if (entity2.Attributes.Contains("msdyn_name"))
                            {
                                ioTServiceCallResult.JobId = entity2.GetAttributeValue<string>("msdyn_name");
                            }

                            if (entity2.Attributes.Contains("msdyn_name"))
                            {
                                ioTServiceCallResult.JobGuid = entity2.Id;
                            }

                            if (entity2.Attributes.Contains("hil_callsubtype"))
                            {
                                ioTServiceCallResult.CallSubType = entity2.GetAttributeValue<EntityReference>("hil_callsubtype").Name;
                            }

                            if (entity2.Attributes.Contains("createdon"))
                            {
                                ioTServiceCallResult.JobLoggedon = entity2.GetAttributeValue<DateTime>("createdon").AddMinutes(330.0).ToString();
                            }

                            if (entity2.Attributes.Contains("msdyn_substatus"))
                            {
                                ioTServiceCallResult.JobStatus = entity2.GetAttributeValue<EntityReference>("msdyn_substatus").Name;
                            }

                            if (entity2.Attributes.Contains("hil_owneraccount"))
                            {
                                ioTServiceCallResult.JobAssignedTo = entity2.GetAttributeValue<EntityReference>("hil_owneraccount").Name;
                            }

                            if (entity2.Attributes.Contains("msdyn_customerasset"))
                            {
                                ioTServiceCallResult.CustomerAsset = entity2.GetAttributeValue<EntityReference>("msdyn_customerasset").Name;
                            }

                            if (entity2.Attributes.Contains("hil_productcategory"))
                            {
                                ioTServiceCallResult.ProductCategory = entity2.GetAttributeValue<EntityReference>("hil_productcategory").Name;
                            }

                            if (entity2.Attributes.Contains("hil_natureofcomplaint"))
                            {
                                ioTServiceCallResult.NatureOfComplaint = entity2.GetAttributeValue<EntityReference>("hil_natureofcomplaint").Name;
                            }

                            if (entity2.Attributes.Contains("hil_jobclosuredon"))
                            {
                                ioTServiceCallResult.JobClosedOn = entity2.GetAttributeValue<DateTime>("hil_jobclosuredon").AddMinutes(330.0).ToString();
                            }

                            if (entity2.Attributes.Contains("hil_customerref"))
                            {
                                ioTServiceCallResult.CustomerName = entity2.GetAttributeValue<EntityReference>("hil_customerref").Name;
                            }

                            if (entity2.Attributes.Contains("hil_fulladdress"))
                            {
                                ioTServiceCallResult.ServiceAddress = entity2.GetAttributeValue<string>("hil_fulladdress");
                            }

                            if (entity2.Attributes.Contains("msdyn_customerasset"))
                            {
                                Entity entity = _CrmService.Retrieve("msdyn_customerasset", entity2.GetAttributeValue<EntityReference>("msdyn_customerasset").Id, new ColumnSet("hil_modelname"));
                                if (entity != null)
                                {
                                    ioTServiceCallResult.Product = entity.GetAttributeValue<string>("hil_modelname");
                                }
                            }

                            if (entity2.Attributes.Contains("hil_customercomplaintdescription"))
                            {
                                ioTServiceCallResult.ChiefComplaint = entity2.GetAttributeValue<string>("hil_customercomplaintdescription");
                            }

                            if (entity2.Attributes.Contains("hil_preferredtime"))
                            {
                                ioTServiceCallResult.PreferredPartOfDay = entity2.GetAttributeValue<OptionSetValue>("hil_preferredtime").Value;
                                ioTServiceCallResult.PreferredPartOfDayName = entity2.FormattedValues["hil_preferredtime"].ToString();
                            }

                            if (entity2.Attributes.Contains("hil_preferreddate"))
                            {
                                ioTServiceCallResult.PreferredDate = entity2.GetAttributeValue<DateTime>("hil_preferreddate").AddMinutes(330.0).ToShortDateString();
                            }

                            ioTServiceCallResult.StatusCode = "200";
                            ioTServiceCallResult.StatusDescription = "OK";
                            list.Add(ioTServiceCallResult);
                        }
                    }

                    return list;
                }

                item = new IoTServiceCallResult
                {
                    StatusCode = "503",
                    StatusDescription = "D365 Service Unavailable"
                };
                list.Add(item);
            }
            catch (Exception ex)
            {
                IoTServiceCallResult item = new IoTServiceCallResult
                {
                    StatusCode = "500",
                    StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper()
                };
                list.Add(item);
            }

            return list;
        }

        public async Task<List<IoTRegisteredProducts>> IoTRegisteredProducts(IoTRegisteredProducts registeredProduct)
        {
              List<IoTRegisteredProducts> list = new List<IoTRegisteredProducts>();
        try
        {

                if (_CrmService != null)
                {
                if (registeredProduct.CustomerGuid.ToString().Trim().Length == 0)
                {
                    IoTRegisteredProducts item = new IoTRegisteredProducts
                    {
                        StatusCode = "204",
                        StatusDescription = "Customer GUID is required."
                    };
                    list.Add(item);
                    return list;
                }

                QueryExpression queryExpression = new QueryExpression("msdyn_customerasset");
                queryExpression.ColumnSet = new ColumnSet("hil_warrantytilldate", "hil_warrantysubstatus", "hil_warrantystatus", "hil_modelname", "hil_product", "hil_retailerpincode", "hil_purchasedfrom", "hil_invoicevalue", "hil_invoiceno", "hil_invoicedate", "hil_invoiceavailable", "hil_batchnumber", "hil_pincode", "msdyn_customerassetid", "createdon", "msdyn_product", "msdyn_name", "hil_productsubcategorymapping", "hil_productsubcategory", "hil_productcategory");
                queryExpression.Criteria = new FilterExpression(LogicalOperator.And);
                queryExpression.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, registeredProduct.CustomerGuid);
                queryExpression.AddOrder("hil_invoicedate", OrderType.Descending);
                queryExpression.TopCount = 30;
                EntityCollection entityCollection = _CrmService.RetrieveMultiple(queryExpression);
                if (entityCollection.Entities.Count == 0)
                {
                    IoTRegisteredProducts item = new IoTRegisteredProducts
                    {
                        StatusCode = "204",
                        StatusDescription = "Customer Product does not exist."
                    };
                    list.Add(item);
                }
                else
                {
                    foreach (Entity entity in entityCollection.Entities)
                    {
                        IoTRegisteredProducts item = new IoTRegisteredProducts();
                        item.DealerPinCode = "";
                        item.BatchNumber = "";
                        item.CustomerGuid = registeredProduct.CustomerGuid;
                        item.RegisteredProductGuid = entity.GetAttributeValue<Guid>("msdyn_customerassetid");
                        if (entity.Attributes.Contains("hil_productcategory"))
                        {
                            item.ProductCategory = entity.GetAttributeValue<EntityReference>("hil_productcategory").Name;
                            item.ProductCategoryId = entity.GetAttributeValue<EntityReference>("hil_productcategory").Id;
                        }

                        if (entity.Attributes.Contains("msdyn_product"))
                        {
                            item.ProductCode = entity.GetAttributeValue<EntityReference>("msdyn_product").Name;
                            item.ProductId = entity.GetAttributeValue<EntityReference>("msdyn_product").Id;
                        }

                        if (entity.Attributes.Contains("hil_modelname"))
                        {
                            item.ProductName = entity.GetAttributeValue<string>("hil_modelname");
                        }

                        if (entity.Attributes.Contains("hil_productsubcategory"))
                        {
                            item.ProductSubCategory = entity.GetAttributeValue<EntityReference>("hil_productsubcategory").Name;
                            item.ProductSubCategoryId = entity.GetAttributeValue<EntityReference>("hil_productsubcategory").Id;
                        }

                        if (entity.Attributes.Contains("msdyn_name"))
                        {
                            item.SerialNumber = entity.GetAttributeValue<string>("msdyn_name");
                        }

                        if (entity.Attributes.Contains("hil_batchnumber"))
                        {
                            item.BatchNumber = entity.GetAttributeValue<string>("hil_batchnumber");
                        }

                        if (entity.Attributes.Contains("hil_invoiceavailable"))
                        {
                            item.InvoiceAvailable = entity.GetAttributeValue<bool>("hil_invoiceavailable");
                        }

                        if (entity.Attributes.Contains("hil_invoiceno"))
                        {
                            item.InvoiceNumber = entity.GetAttributeValue<string>("hil_invoiceno");
                        }

                        if (entity.Attributes.Contains("hil_invoicedate"))
                        {
                            item.InvoiceDate = entity.GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330.0).ToShortDateString();
                        }

                        if (entity.Attributes.Contains("hil_invoicevalue"))
                        {
                            item.InvoiceValue = entity.GetAttributeValue<decimal>("hil_invoicevalue");
                        }

                        if (entity.Attributes.Contains("hil_purchasedfrom"))
                        {
                            item.PurchasedFrom = entity.GetAttributeValue<string>("hil_purchasedfrom");
                        }

                        if (entity.Attributes.Contains("hil_retailerpincode"))
                        {
                            item.PurchasedFromLocation = entity.GetAttributeValue<string>("hil_retailerpincode");
                        }

                        if (entity.Attributes.Contains("hil_product"))
                        {
                            item.InstalledLocationEnum = entity.GetAttributeValue<OptionSetValue>("hil_product").Value;
                        }

                        if (entity.Attributes.Contains("hil_product"))
                        {
                            item.InstalledLocation = entity.FormattedValues["hil_product"].ToString();
                        }

                        if (entity.Attributes.Contains("hil_warrantystatus"))
                        {
                            OptionSetValue attributeValue = entity.GetAttributeValue<OptionSetValue>("hil_warrantystatus");
                            item.WarrantyStatus = ((attributeValue.Value == 1) ? "In Warranty" : "Out Of Warranty");
                        }
                        else
                        {
                            item.WarrantyStatus = "Pending for Approval";
                            item.WarrantySubStatus = "";
                            item.WarrantyEndDate = "";
                        }

                        if (entity.Attributes.Contains("hil_warrantysubstatus"))
                        {
                            OptionSetValue attributeValue2 = entity.GetAttributeValue<OptionSetValue>("hil_warrantysubstatus");
                            item.WarrantySubStatus = ((attributeValue2.Value == 1) ? "Standard" : ((attributeValue2.Value == 2) ? "Extended" : ((attributeValue2.Value == 3) ? "Special Scheme" : "AMC")));
                        }

                        if (entity.Attributes.Contains("hil_warrantytilldate"))
                        {
                            item.WarrantyEndDate = entity.GetAttributeValue<DateTime>("hil_warrantytilldate").AddMinutes(330.0).ToShortDateString();
                        }

                        string query = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>\r\n                                <entity name='hil_unitwarranty'>\r\n                                    <attribute name='hil_warrantystartdate' />\r\n                                    <attribute name='hil_warrantyenddate' />\r\n                                    <order attribute='hil_warrantystartdate' descending='false' />\r\n                                    <filter type='and'>\r\n                                      <condition attribute='hil_customerasset' operator='eq' value='" + item.RegisteredProductGuid.ToString() + "' />\r\n                                      <condition attribute='statecode' operator='eq' value='0' />\r\n                                    </filter>\r\n                                    <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' visible='false' link-type='outer' alias='wrt'>\r\n                                      <attribute name='hil_description' />\r\n                                      <attribute name='hil_type' />\r\n                                    </link-entity>\r\n                                  </entity>\r\n                                </fetch>";
                        EntityCollection entityCollection2 = _CrmService.RetrieveMultiple(new FetchExpression(query));
                        item.ProductWarranty = new List<IoTProductWarranty>();
                        if (entityCollection2.Entities.Count > 0)
                        {
                            OptionSetValue optionSetValue = new OptionSetValue();
                            string empty = string.Empty;
                            string empty2 = string.Empty;
                            string warrantySpecifications = string.Empty;
                            foreach (Entity entity2 in entityCollection2.Entities)
                            {
                                optionSetValue = (OptionSetValue)entity2.GetAttributeValue<AliasedValue>("wrt.hil_type").Value;
                                DateTime dateTime = entity2.GetAttributeValue<DateTime>("hil_warrantystartdate").AddMinutes(330.0);
                                empty = dateTime.Day.ToString().PadLeft(2, '0') + "/" + dateTime.Month.ToString().PadLeft(2, '0') + "/" + dateTime.Year;
                                dateTime = entity2.GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330.0);
                                empty2 = dateTime.Day.ToString().PadLeft(2, '0') + "/" + dateTime.Month.ToString().PadLeft(2, '0') + "/" + dateTime.Year;
                                if (entity2.Attributes.Contains("wrt.hil_description"))
                                {
                                    warrantySpecifications = entity2.GetAttributeValue<AliasedValue>("wrt.hil_description").Value.ToString();
                                }

                                item.ProductWarranty.Add(new IoTProductWarranty
                                {
                                    WarrantyType = ((optionSetValue.Value == 1) ? "Standard" : ((optionSetValue.Value == 2) ? "Extended" : ((optionSetValue.Value == 3) ? "AMC" : "Special Scheme"))),
                                    WarrantyStartDate = empty,
                                    WarrantyEndDate = empty2,
                                    WarrantySpecifications = warrantySpecifications
                                });
                            }
                        }

                        item.StatusCode = "200";
                        item.StatusDescription = "OK";
                        list.Add(item);
                    }
                }
            }
            else
            {
                IoTRegisteredProducts item = new IoTRegisteredProducts
                {
                    StatusCode = "503",
                    StatusDescription = "D365 Service Unavailable"
                };
                list.Add(item);
            }
        }
        catch (Exception ex)
        {
            IoTRegisteredProducts item = new IoTRegisteredProducts
            {
                StatusCode = "500",
                StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper()
            };
            list.Add(item);
        }

        return list;
        }

        public async Task<ReturnResult> IoTRegisterConsumer(IoT_RegisterConsumer consumer)
        {
            ReturnResult returnResult = new ReturnResult();
            Guid guid = Guid.Empty;
            EntityReference entityReference = null;
            try
            {
               if (consumer.MobileNumber == null || consumer.MobileNumber.Trim().Length == 0)
                {
                    returnResult.StatusCode = "204";
                    returnResult.StatusDescription = "No Content : Mobile Number is required.";
                    return returnResult;
                }

                if (consumer.PreferredLanguage != null)
                {
                    QueryExpression queryExpression = new QueryExpression("hil_preferredlanguageforcommunication");
                    queryExpression.ColumnSet = new ColumnSet(allColumns: false);
                    ConditionExpression condition = new ConditionExpression("hil_code", ConditionOperator.Equal, consumer.PreferredLanguage.Trim());
                    queryExpression.Criteria.AddCondition(condition);
                    EntityCollection entityCollection = _CrmService.RetrieveMultiple(queryExpression);
                    if (entityCollection.Entities.Count == 0)
                    {
                        returnResult.StatusCode = "204";
                        returnResult.StatusDescription = "No Content : Preferred Language does not exist.";
                        return returnResult;
                    }

                    entityReference = entityCollection.Entities[0].ToEntityReference();
                }

                if (!consumer.SourceOfCreation.HasValue)
                {
                    returnResult.StatusCode = "204";
                    returnResult.StatusDescription = "No Content : Source of Registration is required. Please pass <4> for Whatsapp <5> for IoT Platform <7> for eCommerce<8> for Chatbot";
                    return returnResult;
                }

                if (_CrmService != null)
                {
                    LinkEntity linkEntity = new LinkEntity();
                    linkEntity.Columns = new ColumnSet("hil_code");
                    linkEntity.EntityAlias = "lang";
                    linkEntity.LinkFromEntityName = "contact";
                    linkEntity.LinkFromAttributeName = "hil_preferredlanguageforcommunication";
                    linkEntity.LinkToEntityName = "hil_preferredlanguageforcommunication";
                    linkEntity.LinkToAttributeName = "hil_preferredlanguageforcommunicationid";
                    linkEntity.JoinOperator = JoinOperator.LeftOuter;
                    LinkEntity item = linkEntity;
                    QueryExpression queryExpression2 = new QueryExpression("contact");
                    queryExpression2.ColumnSet = new ColumnSet("contactid", "fullname", "emailaddress1", "hil_salutation", "hil_consent", "hil_subscribeformessagingservice", "hil_preferredlanguageforcommunication");
                    ConditionExpression condition2 = new ConditionExpression("mobilephone", ConditionOperator.Equal, consumer.MobileNumber);
                    queryExpression2.Criteria.AddCondition(condition2);
                    queryExpression2.LinkEntities.Add(item);
                    EntityCollection entityCollection2 = _CrmService.RetrieveMultiple(queryExpression2);
                    if (entityCollection2.Entities.Count > 0)
                    {
                        guid = entityCollection2.Entities[0].Id;
                        returnResult.MobileNumber = consumer.MobileNumber;
                        returnResult.CustomerName = entityCollection2.Entities[0].GetAttributeValue<string>("fullname");
                        returnResult.EmailId = entityCollection2.Entities[0].GetAttributeValue<string>("emailaddress1");
                        if (entityCollection2.Entities[0].Attributes.Contains("hil_consent"))
                        {
                            returnResult.Consent = entityCollection2.Entities[0].GetAttributeValue<bool>("hil_consent");
                        }
                        else
                        {
                            returnResult.Consent = false;
                        }

                        if (entityCollection2.Entities[0].Attributes.Contains("hil_subscribeformessagingservice"))
                        {
                            returnResult.SubscribeForMsgService = entityCollection2.Entities[0].GetAttributeValue<bool>("hil_subscribeformessagingservice");
                        }
                        else
                        {
                            returnResult.SubscribeForMsgService = false;
                        }

                        if (entityCollection2.Entities[0].Attributes.Contains("hil_preferredlanguageforcommunication"))
                        {
                            returnResult.PreferredLanguage = entityCollection2.Entities[0].GetAttributeValue<AliasedValue>("lang.hil_code").Value.ToString().Trim();
                        }

                        if (consumer.SourceOfCreation == 11)
                        {
                            QueryExpression queryExpression3 = new QueryExpression("hil_address");
                            queryExpression3.ColumnSet = new ColumnSet("hil_fulladdress", "hil_pincode");
                            queryExpression3.Criteria.AddCondition(new ConditionExpression("hil_customer", ConditionOperator.Equal, entityCollection2.Entities[0].Id));
                            queryExpression3.TopCount = 1;
                            queryExpression3.AddOrder("modifiedon", OrderType.Descending);
                            EntityCollection entityCollection3 = _CrmService.RetrieveMultiple(queryExpression3);
                            if (entityCollection3.Entities.Count > 0)
                            {
                                returnResult.PINCode = entityCollection3.Entities[0].GetAttributeValue<EntityReference>("hil_pincode").Name;
                                returnResult.Address = entityCollection3.Entities[0].GetAttributeValue<string>("hil_fulladdress");
                            }
                            else
                            {
                                returnResult.PINCode = "NA";
                            }
                        }

                        returnResult.StatusCode = "208";
                        returnResult.StatusDescription = "Already Reported";
                        returnResult.CustomerGuid = guid;
                    }

                    if (guid == Guid.Empty)
                    {
                        if ((consumer.SourceOfCreation == 4 || consumer.SourceOfCreation == 7 || consumer.SourceOfCreation == 8 || consumer.SourceOfCreation == 9 || consumer.SourceOfCreation == 10 || consumer.SourceOfCreation == 11 || consumer.SourceOfCreation == 17) && (consumer.FirstName == null || consumer.FirstName.Trim().Length == 0))
                        {
                            returnResult.StatusCode = "204";
                            returnResult.StatusDescription = "No Content : Customer Name is required.";
                        }
                        else
                        {
                            Entity entity = new Entity("contact");
                            entity["mobilephone"] = consumer.MobileNumber;
                            if (consumer.Salutation.HasValue)
                            {
                                new List<HashTableDTO>();
                                if (!new IoTCommonLib(_CrmService).GetSalutationEnum().Exists((HashTableDTO x) => x.Value == consumer.Salutation))
                                {
                                    returnResult.StatusCode = "204";
                                    returnResult.StatusDescription = "Salutation not found in D365. Please Call <GetSalutationEnum> API and pass proper Salutation.";
                                    return returnResult;
                                }

                                entity["hil_salutation"] = new OptionSetValue(consumer.Salutation.Value);
                            }

                            if (consumer.FirstName == null)
                            {
                                entity["firstname"] = "UNDEF-IoT-" + consumer.MobileNumber;
                            }
                            else
                            {
                                string[] array = consumer.FirstName.Split(' ');
                                if (array.Length >= 1)
                                {
                                    entity["firstname"] = array[0];
                                    if (array.Length == 3)
                                    {
                                        entity["middlename"] = array[1];
                                        entity["lastname"] = array[2];
                                    }

                                    if (array.Length == 2)
                                    {
                                        entity["lastname"] = array[1];
                                    }
                                }
                                else
                                {
                                    entity["firstname"] = consumer.FirstName;
                                }
                            }

                            if (consumer.Email != null && consumer.Email.Trim().Length > 0)
                            {
                                entity["emailaddress1"] = consumer.Email;
                            }

                            entity["hil_consumersource"] = new OptionSetValue(consumer.SourceOfCreation.Value);
                            if (consumer.Consent.HasValue)
                            {
                                entity["hil_consent"] = consumer.Consent;
                            }

                            if (consumer.SubscribeForMsgService.HasValue)
                            {
                                entity["hil_subscribeformessagingservice"] = consumer.SubscribeForMsgService;
                            }

                            if (entityReference != null)
                            {
                                entity["hil_preferredlanguageforcommunication"] = entityReference;
                            }

                            guid = _CrmService.Create(entity);
                            returnResult.CustomerGuid = guid;
                            returnResult.StatusCode = "200";
                            returnResult.StatusDescription = "OK";
                        }
                    }
                }
                else
                {
                    returnResult.StatusCode = "503";
                    returnResult.StatusDescription = "D365 Service Unavailable";
                }
            }
            catch (Exception ex)
            {
                returnResult.StatusCode = "500";
                returnResult.StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper();
            }

            return returnResult;
        }

        public class IoTCommonLib
        {

            private ICrmService _CrmService;

            public IoTCommonLib(ICrmService CrmService)
            {
                this._CrmService = CrmService;
            }
            public List<HashTableDTO> GetSalutationEnum()
            {
                HashTableDTO objIoTSalutationEnum;
                List<HashTableDTO> lstIoTSalutationEnum = new List<HashTableDTO>();
                try
                {

                    if (_CrmService != null)
                    {
                        var attributeRequest = new RetrieveAttributeRequest
                        {
                            EntityLogicalName = "contact",
                            LogicalName = "hil_salutation",
                            RetrieveAsIfPublished = true
                        };

                        var attributeResponse = (RetrieveAttributeResponse)_CrmService.Execute(attributeRequest);
                        var attributeMetadata = (EnumAttributeMetadata)attributeResponse.AttributeMetadata;

                        var optionList = (from o in attributeMetadata.OptionSet.Options
                                          select new { Value = o.Value, Text = o.Label.UserLocalizedLabel.Label }).ToList();
                        foreach (var option in optionList)
                        {
                            lstIoTSalutationEnum.Add(new HashTableDTO() { Value = option.Value, Label = option.Text, StatusCode = "200", StatusDescription = "OK" });
                        }
                    }
                    else
                    {
                        objIoTSalutationEnum = new HashTableDTO { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                        lstIoTSalutationEnum.Add(objIoTSalutationEnum);
                    }
                }
                catch (Exception ex)
                {
                    objIoTSalutationEnum = new HashTableDTO { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                    lstIoTSalutationEnum.Add(objIoTSalutationEnum);
                }
                return lstIoTSalutationEnum;
            }
        }

        public async Task<List<IoTNatureofComplaint>> IoTNatureOfComplaintByProdSubcategory(IoTNatureofComplaint natureOfComplaint)
        {
            List<IoTNatureofComplaint> list = new List<IoTNatureofComplaint>();
            try
            {               
                IoTNatureofComplaint item;
                if (_CrmService != null)
                {
                    if (natureOfComplaint.ProductSubCategoryId == Guid.Empty)
                    {
                        item = new IoTNatureofComplaint
                        {
                            StatusCode = "204",
                            StatusDescription = "Product Subcategory is required."
                        };
                        list.Add(item);
                        return list;
                    }

                    string empty = string.Empty;
                    empty = ((natureOfComplaint.Source != null) ? ("<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>\r\n                            <entity name='hil_natureofcomplaint'>\r\n                            <attribute name='hil_name' />\r\n                            <attribute name='hil_natureofcomplaintid' />\r\n                            <order attribute='hil_name' descending='false' />\r\n                            <filter type='and'>\r\n                                <condition attribute='statecode' operator='eq' value='0' />\r\n                                <condition attribute='hil_relatedproduct' operator='eq' value='{" + natureOfComplaint.ProductSubCategoryId.ToString() + "}' />\r\n                                <condition attribute='hil_callsubtype' operator='in'>\r\n                                <value uiname='AMC\u00a0Call' uitype='hil_callsubtype'>{55A71A52-3C0B-E911-A94E-000D3AF06CD4}</value>\r\n                                <value uiname='Breakdown' uitype='hil_callsubtype'>{6560565A-3C0B-E911-A94E-000D3AF06CD4}</value>\r\n                                <value uiname='Demo' uitype='hil_callsubtype'>{AE1B2B71-3C0B-E911-A94E-000D3AF06CD4}</value>\r\n                                <value uiname='Installation' uitype='hil_callsubtype'>{E3129D79-3C0B-E911-A94E-000D3AF06CD4}</value>\r\n                                <value uiname='PMS' uitype='hil_callsubtype'>{E2129D79-3C0B-E911-A94E-000D3AF06CD4}</value>\r\n                                </condition>\r\n                            </filter>\r\n                            </entity>\r\n                            </fetch>") : ("<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>\r\n                            <entity name='hil_natureofcomplaint'>\r\n                            <attribute name='hil_name' />\r\n                            <attribute name='hil_natureofcomplaintid' />\r\n                            <order attribute='hil_name' descending='false' />\r\n                            <filter type='and'>\r\n                                <condition attribute='statecode' operator='eq' value='0' />\r\n                                <condition attribute='hil_relatedproduct' operator='eq' value='{" + natureOfComplaint.ProductSubCategoryId.ToString() + "}' />\r\n                            </filter>\r\n                            </entity>\r\n                            </fetch>"));
                    EntityCollection entityCollection = _CrmService.RetrieveMultiple(new FetchExpression(empty));
                    if (entityCollection.Entities.Count == 0)
                    {
                        item = new IoTNatureofComplaint
                        {
                            StatusCode = "204",
                            StatusDescription = "No Nature of Complaint is mapped with Serial Number."
                        };
                        list.Add(item);
                    }
                    else
                    {
                        foreach (Entity entity in entityCollection.Entities)
                        {
                            item = new IoTNatureofComplaint();
                            item.Guid = entity.GetAttributeValue<Guid>("hil_natureofcomplaintid");
                            item.Name = entity.GetAttributeValue<string>("hil_name");
                            item.SerialNumber = natureOfComplaint.SerialNumber;
                            item.ProductSubCategoryId = natureOfComplaint.ProductSubCategoryId;
                            item.StatusCode = "200";
                            item.StatusDescription = "OK";
                            list.Add(item);
                        }
                    }

                    return list;
                }

                item = new IoTNatureofComplaint
                {
                    StatusCode = "503",
                    StatusDescription = "D365 Service Unavailable"
                };
                list.Add(item);
                return list;
            }
            catch (Exception ex)
            {
                IoTNatureofComplaint item = new IoTNatureofComplaint
                {
                    StatusCode = "500",
                    StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper()
                };
                list.Add(item);
                return list;
            }
        }

        public async Task<List<NatureOfComplaint>> NatureOfComplaint(NatureOfComplaint natureOfComplaint)
        {
            List<NatureOfComplaint> list = new List<NatureOfComplaint>();
            try
            {               
                if (natureOfComplaint.SerialNumber.Trim().Length == 0)
                {
                    NatureOfComplaint item = new NatureOfComplaint
                    {
                        ResultStatus = false,
                        ResultMessage = "Product Serial Number is required."
                    };
                    list.Add(item);
                    return list;
                }

                QueryExpression queryExpression = new QueryExpression("msdyn_customerasset");
                queryExpression.ColumnSet = new ColumnSet("msdyn_name");
                queryExpression.Criteria = new FilterExpression(LogicalOperator.And);
                queryExpression.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, natureOfComplaint.SerialNumber);
                queryExpression.TopCount = 1;
                EntityCollection entityCollection = _CrmService.RetrieveMultiple(queryExpression);
                if (entityCollection.Entities.Count == 0)
                {
                    NatureOfComplaint item = new NatureOfComplaint
                    {
                        ResultStatus = false,
                        ResultMessage = "Product Serial Number does not exist."
                    };
                    list.Add(item);
                    return list;
                }

                string query = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>\r\n                    <entity name='hil_natureofcomplaint'>\r\n                        <attribute name='hil_name' />\r\n                        <attribute name='hil_natureofcomplaintid' />\r\n                        <order attribute='hil_name' descending='false' />\r\n                        <filter type='and'>\r\n                          <condition attribute='statecode' operator='eq' value='0' />\r\n                        </filter>\r\n                        <link-entity name='product' from='productid' to='hil_relatedproduct' link-type='inner' alias='ae'>\r\n                            <link-entity name='msdyn_customerasset' from='hil_productsubcategory' to='productid' link-type='inner' alias='af'>\r\n                                <filter type='and'>\r\n                                    <condition attribute='msdyn_name' operator='eq' value='" + natureOfComplaint.SerialNumber + "' />\r\n                                </filter>\r\n                            </link-entity>\r\n                        </link-entity>\r\n                    </entity>\r\n                    </fetch>";
                entityCollection = _CrmService.RetrieveMultiple(new FetchExpression(query));
                if (entityCollection.Entities.Count == 0)
                {
                    NatureOfComplaint item = new NatureOfComplaint
                    {
                        ResultStatus = false,
                        ResultMessage = "No Nature of Complaint is mapped with Serial Number."
                    };
                    list.Add(item);
                }
                else
                {
                    foreach (Entity entity in entityCollection.Entities)
                    {
                        NatureOfComplaint item = new NatureOfComplaint();
                        item.Guid = entity.GetAttributeValue<Guid>("hil_natureofcomplaintid");
                        item.Name = entity.GetAttributeValue<string>("hil_name");
                        item.SerialNumber = natureOfComplaint.SerialNumber;
                        item.ResultStatus = true;
                        item.ResultMessage = "Success";
                        list.Add(item);
                    }
                }

                return list;
            }
            catch (Exception ex)
            {
                NatureOfComplaint item = new NatureOfComplaint
                {
                    ResultStatus = false,
                    ResultMessage = ex.Message
                };
                list.Add(item);
                return list;
            }
        }

        public async Task<List<NatureOfComplaint>> AllNatureOfComplaints()
        {
            List<NatureOfComplaint> list = new List<NatureOfComplaint>();
            try
            {
               
                string query = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                    <entity name='hil_natureofcomplaint'>
                    <attribute name='hil_name' />
                    <attribute name='hil_natureofcomplaintid' />
                    <attribute name='hil_relatedproduct' />
                    <order attribute='hil_name' descending='false' /><filter type='and'>
                    <condition attribute='statecode' operator='eq' value='0' />
                    </filter></entity></fetch>";
                EntityCollection entityCollection = _CrmService.RetrieveMultiple(new FetchExpression(query));
                if (entityCollection.Entities.Count == 0)
                {
                    NatureOfComplaint item = new NatureOfComplaint
                    {
                        ResultStatus = false,
                        ResultMessage = "No Nature of Complaint is mapped with Serial Number."
                    };
                    list.Add(item);
                }
                else
                {
                    foreach (Entity entity in entityCollection.Entities)
                    {
                        NatureOfComplaint item = new NatureOfComplaint();
                        item.Guid = entity.GetAttributeValue<Guid>("hil_natureofcomplaintid");
                        item.Name = entity.GetAttributeValue<string>("hil_name");
                        if (entity.Contains("hil_relatedproduct") || entity.Attributes.Contains("hil_relatedproduct"))
                        {
                            item.ProductCategoryGuid = entity.GetAttributeValue<EntityReference>("hil_relatedproduct").Id;
                            item.ProductCategoryName = entity.GetAttributeValue<EntityReference>("hil_relatedproduct").Name;
                        }

                        item.ResultStatus = true;
                        item.ResultMessage = "Success";
                        list.Add(item);
                    }
                }

                return list;
            }
            catch (Exception ex)
            {
                NatureOfComplaint item = new NatureOfComplaint
                {
                    ResultStatus = false,
                    ResultMessage = ex.Message
                };
                list.Add(item);
                return list;
            }
        }

        public async Task<(OCLDetailsResponse, RequestStatus)> GetOCLDetails(OCLDetailsParam _oclRequest)
        {
            OCLDetailsResponse _oclResponse = new OCLDetailsResponse();
            try
            {
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='hil_oaheader'>
                        <attribute name='hil_name'/>
                        <filter type='and'>
                            <condition attribute='hil_department' operator='eq' value='CE8B92CB-E64C-EC11-8F8E-6045BD733E10'/>
                            <condition attribute='hil_name' operator='eq' value='{_oclRequest.OrderNumber}'/>
                        </filter>
                          <link-entity name='hil_orderchecklist' from='hil_orderchecklistid' to='hil_orderchecklistid' link-type='inner' alias='ocl'>
                              <attribute name='hil_approveddatasheetgtp'/>
                              <attribute name='hil_drumlengthschedule'/>
                              <attribute name='hil_individual'/>
                              <attribute name='hil_inspection'/>
                              <attribute name='hil_markingdetails'/>
                              <attribute name='hil_name'/>
                              <attribute name='hil_overall'/>
                              <attribute name='hil_typeofdrum'/>
                          </link-entity>
                      </entity>
                    </fetch>";

                EntityCollection _oclDetailsColl = _CrmService.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (_oclDetailsColl.Entities.Count > 0)
                {
                    _oclResponse.OCLNumber = _oclDetailsColl[0].Contains("ocl.hil_name") ? _oclDetailsColl[0].GetAttributeValue<AliasedValue>("ocl.hil_name").Value.ToString() : null;


                    bool ApprovedDataSheet = (bool)_oclDetailsColl[0].GetAttributeValue<AliasedValue>("ocl.hil_approveddatasheetgtp").Value;
                    if (!ApprovedDataSheet)
                    {
                        _oclResponse.ApprovedDataSheet= "Not Involved";
                    }
                    else
                    {
                        _oclResponse.ApprovedDataSheet = "Involved";
                    }

                    bool DrumLengthSchedule =(bool)_oclDetailsColl[0].GetAttributeValue<AliasedValue>("ocl.hil_drumlengthschedule").Value;
                    if (!DrumLengthSchedule)
                    {
                        _oclResponse.DrumLengthSchedule = "Standard";
                    }
                    else
                    {
                        _oclResponse.DrumLengthSchedule = "Specific";
                    }

                    _oclResponse.StdDrumLength = _oclDetailsColl[0].Contains("ocl.hil_individual") ? _oclDetailsColl[0].GetAttributeValue<AliasedValue>("ocl.hil_individual").Value.ToString():null;
                    _oclResponse.QtyTolerance =  _oclDetailsColl[0].Contains("ocl.hil_overall") ? _oclDetailsColl[0].GetAttributeValue<AliasedValue>("ocl.hil_overall").Value.ToString():null;


                    OptionSetValue typeofdrumoptionset = (OptionSetValue)_oclDetailsColl[0].GetAttributeValue<AliasedValue>("ocl.hil_typeofdrum").Value;
                    if (typeofdrumoptionset.Value == 910590000)
                    {
                        _oclResponse.TypeOfDrum = "Wooden";
                    }
                    else if (typeofdrumoptionset.Value == 910590001)
                    {
                        _oclResponse.TypeOfDrum = "Steel";

                    }

                    bool Inspection = (bool)_oclDetailsColl[0].GetAttributeValue<AliasedValue>("ocl.hil_inspection").Value;
                    if (!Inspection)
                    {
                        _oclResponse.Inspection = "Not Applicable";
                    }
                    else
                    {
                        _oclResponse.Inspection = "Applicable";
                    }


                    bool MarkingDetails = (bool)_oclDetailsColl[0].GetAttributeValue<AliasedValue>("ocl.hil_markingdetails").Value;
                    if (!MarkingDetails)
                    {
                        _oclResponse.MarkingDetails = "Not Applicable";
                    }
                    else
                    {
                        _oclResponse.MarkingDetails = "Applicable";
                    }


                }
                else
                {

                    return (_oclResponse, new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.NotFound,
                        Message = "Invalid Order Number"
                    });
                }
                return (_oclResponse, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.OK
                });
            }
            catch (Exception ex)
            {
                return (_oclResponse, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()

                });
            }
        }

    }

  
}
