using System;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;
using RestSharp;
using Microsoft.Xrm.Sdk;
using System.Text.RegularExpressions;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using System.IO;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System.Linq;
using System.Runtime.Serialization;
using Microsoft.Crm.Sdk.Messages;
using System.Text;

namespace ConsumerApp.BusinessLayer
{
    public class HomeAdvisory
    {
        public Response CreateAppointmentD365(CRMRequest req)
        {
            Response res = new Response();
            CreateUserMeetingResponse resp = new CreateUserMeetingResponse();
            try
            {
                if (req.SlotDate == string.Empty || req.SlotDate == null)
                {
                    res.Message = ("SlotDate is Null");
                    res.Status = false;
                    return res;
                }
                if (req.SlotEnd == string.Empty || req.SlotEnd == null)
                {
                    res.Message = ("SlotEnd is Null");
                    res.Status = false;
                    return res;
                }
                if (req.SlotStart == string.Empty || req.SlotStart == null)
                {
                    res.Message = ("SlotStart is Null");
                    res.Status = false;
                    return res;
                }
                if (req.RecordID == string.Empty || req.RecordID == null)
                {
                    res.Message = ("RecordID is Null");
                    res.Status = false;
                    return res;
                }
                #region get Privious Appointment

                #endregion
                IOrganizationService service = ConnectToCRM.GetOrgService();
                #region
                QueryExpression Query = new QueryExpression("hil_homeadvisoryline");
                Query.ColumnSet = new ColumnSet("hil_appointmentdate", "hil_advisoryenquery", "hil_name", "hil_typeofenquiiry", "hil_typeofproduct", "hil_assignedadvisor", "hil_appointmentid");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_homeadvisorylineid", ConditionOperator.Equal, req.RecordID);

                LinkEntity EntityA = new LinkEntity("hil_homeadvisoryline", "hil_advisoryenquiry", "hil_advisoryenquery", "hil_advisoryenquiryid", JoinOperator.LeftOuter);
                EntityA.Columns = new ColumnSet("hil_customer", "hil_emailid");
                EntityA.EntityAlias = "PEnq";
                Query.LinkEntities.Add(EntityA);

                LinkEntity Entityb = new LinkEntity("hil_homeadvisoryline", "hil_advisormaster", "hil_assignedadvisor", "hil_advisormasterid", JoinOperator.LeftOuter);
                Entityb.Columns = new ColumnSet("hil_code");
                Entityb.EntityAlias = "Advisor";
                Query.LinkEntities.Add(Entityb);
                #endregion

                EntityCollection Found = service.RetrieveMultiple(Query);// new FetchExpression(FetchXML));
                if (Found.Entities.Count == 1)
                {

                    String SlotEnd = string.Empty;
                    String SlotStart = string.Empty;
                    String EnquirerEmailId = string.Empty;
                    String EnquirerName = string.Empty;
                    string AppointmentID = string.Empty;
                    string EnquiryId = string.Empty;
                    string EnquiryType = string.Empty;
                    string AdvisoryType = string.Empty;

                    DateTime SlotDate = new DateTime(Convert.ToInt32(req.SlotDate.Substring(0, 4)), Convert.ToInt32(req.SlotDate.Substring(4, 2)), Convert.ToInt32(req.SlotDate.Substring(6, 2)));

                    String UserCode = string.Empty;
                    String ddymmyyyy = string.Empty;
                    EntityReference cust = new EntityReference();
                    Guid regarding = new Guid();
                    foreach (Entity line in Found.Entities)
                    {
                        SlotEnd = req.SlotEnd;
                        SlotStart = req.SlotStart;
                        EnquirerEmailId = line.Contains("PEnq.hil_emailid") ? (line["PEnq.hil_emailid"] as AliasedValue).Value.ToString() : throw new Exception("EnquirerEmailId not found");
                        EnquirerName = line.Contains("PEnq.hil_customer") ? ((EntityReference)((AliasedValue)line["PEnq.hil_customer"]).Value).Name.ToString() : throw new Exception("EnquirerName not found");

                        EnquiryId = line.Contains("hil_name") ? line.GetAttributeValue<String>("hil_name") : throw new Exception("EnquiryID not found");
                        EnquiryType = line.Contains("hil_typeofenquiiry") ? line.GetAttributeValue<EntityReference>("hil_typeofenquiiry").Name : throw new Exception("typeofenquiiry not found");
                        AdvisoryType = line.Contains("hil_typeofproduct") ? line.GetAttributeValue<EntityReference>("hil_typeofproduct").Name : throw new Exception("typeofproduct not found");

                        UserCode = line.Contains("Advisor.hil_code") ? (line["Advisor.hil_code"] as AliasedValue).Value.ToString() : throw new Exception("UserCode not found"); ;
                        cust = (EntityReference)((AliasedValue)line["PEnq.hil_customer"]).Value;
                        regarding = line.Id;

                        AppointmentID = line.Contains("hil_appointmentid") ? line.GetAttributeValue<String>("hil_appointmentid") : string.Empty;
                        if (AppointmentID == string.Empty)
                        {
                            String fetct = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">" +
                                              "<entity name=\"appointment\">" +
                                                "<attribute name=\"subject\" />" +
                                                "<attribute name=\"statecode\" />" +
                                                "<attribute name=\"scheduledstart\" />" +
                                                "<attribute name=\"scheduledend\" />" +
                                                "<attribute name=\"createdby\" />" +
                                                "<attribute name=\"regardingobjectid\" />" +
                                                "<attribute name=\"activityid\" />" +
                                                "<attribute name=\"instancetypecode\" />" +
                                                "<order attribute=\"subject\" descending=\"false\" />" +
                                                "<filter type=\"and\">" +
                                                  "<condition attribute=\"regardingobjectid\" operator=\"eq\" uiname=\"\" uitype=\"hil_homeadvisoryline\" value=\"" + regarding + "\" />" +
                                                  "<condition attribute=\"hil_appointmenturl\" operator=\"not-null\" />" +
                                                  "<condition attribute=\"statecode\" operator=\"in\">" +
                                                    "<value>0</value>" +
                                                     "<value>3</value>" +
                                                    "</condition>" +
                                                  "</filter>" +
                                               "</entity>" +
                                              "</fetch>";
                            EntityCollection _app = service.RetrieveMultiple(new FetchExpression(fetct));

                            if (_app.Entities.Count > 0)
                            {
                                res.Message = "Privious Appointment not Completed or Cancled";
                                res.Status = false;
                                return res;
                            }
                        }
                    }
                    if (AppointmentID == string.Empty)
                    {
                        resp = CreateMeeting(AppointmentID, SlotEnd, SlotStart, EnquirerEmailId, EnquirerName, req.SlotDate.Substring(4, 2) + "/" + req.SlotDate.Substring(6, 2) + "/" + (req.SlotDate.Substring(0, 4)), UserCode, req.IsVideoMeeting, EnquiryId, EnquiryType, AdvisoryType);

                        if (resp.IsSuccess)
                        {
                            Entity _appointment = new Entity("appointment");
                            Entity from = new Entity("activityparty");
                            from["partyid"] = cust;
                            _appointment["requiredattendees"] = new Entity[] { from };
                            _appointment["subject"] = "Meeting with Havells for Advisory";
                            if (req.IsVideoMeeting)
                                _appointment["location"] = "Teams";
                            else
                                _appointment["location"] = "Audio Call";
                            _appointment["regardingobjectid"] = new EntityReference("hil_homeadvisoryline", regarding);
                            _appointment["scheduledstart"] = new DateTime(Convert.ToInt32(req.SlotDate.Substring(0, 4)), Convert.ToInt32(req.SlotDate.Substring(4, 2)), Convert.ToInt32(req.SlotDate.Substring(6, 2)), Convert.ToInt32(req.SlotStart.Substring(0, 2)), Convert.ToInt32(req.SlotStart.Substring(3, 2)), 0);
                            _appointment["scheduledend"] = new DateTime(Convert.ToInt32(req.SlotDate.Substring(0, 4)), Convert.ToInt32(req.SlotDate.Substring(4, 2)), Convert.ToInt32(req.SlotDate.Substring(6, 2)), Convert.ToInt32(req.SlotEnd.Substring(0, 2)), Convert.ToInt32(req.SlotEnd.Substring(3, 2)), 0);
                            if (resp.Data.MeetingURL != null && resp.Data.MeetingURL != string.Empty && resp.Data.MeetingURL != "")
                            {
                                _appointment["hil_appointmenturl"] = resp.Data.MeetingURL;
                            }
                            service.Create(_appointment);

                            Entity _enqLine = new Entity(Found.EntityName);
                            _enqLine.Id = Found.Entities[0].Id;

                            _enqLine["hil_appointmentid"] = resp.Data.TransactionId;
                            if (resp.Data.MeetingURL != "" || resp.Data.MeetingURL != null)
                                _enqLine["hil_videocallurl"] = resp.Data.MeetingURL;
                            _enqLine["hil_appointmentstatus"] = new OptionSetValue(5);
                            _enqLine["hil_appointmentdate"] = new DateTime(Convert.ToInt32(req.SlotDate.Substring(0, 4)), Convert.ToInt32(req.SlotDate.Substring(4, 2)), Convert.ToInt32(req.SlotDate.Substring(6, 2)), Convert.ToInt32(req.SlotStart.Substring(0, 2)), Convert.ToInt32(req.SlotStart.Substring(3, 2)), 0);
                            _enqLine["hil_appointmentenddate"] = new DateTime(Convert.ToInt32(req.SlotDate.Substring(0, 4)), Convert.ToInt32(req.SlotDate.Substring(4, 2)), Convert.ToInt32(req.SlotDate.Substring(6, 2)), Convert.ToInt32(req.SlotEnd.Substring(0, 2)), Convert.ToInt32(req.SlotEnd.Substring(3, 2)), 0);
                            _enqLine["hil_appointmenttypes"] = req.IsVideoMeeting ? new OptionSetValue(2) : new OptionSetValue(1);
                            service.Update(_enqLine);
                            res.Message = resp.Message;
                            res.Status = true;
                        }
                        else
                        {
                            res.Message = resp.Message;
                            return res;
                        }
                    }
                    else
                    {
                        resp = CreateMeeting(AppointmentID, SlotEnd, SlotStart, EnquirerEmailId, EnquirerName, req.SlotDate.Substring(4, 2) + "/" + req.SlotDate.Substring(6, 2) + "/" + (req.SlotDate.Substring(0, 4)), UserCode, req.IsVideoMeeting, EnquiryId, EnquiryType, AdvisoryType);
                        if (resp.IsSuccess)
                        {
                            string fetchappointment = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='appointment'>
                                            <attribute name='subject' />
                                            <attribute name='statecode' />
                                            <attribute name='scheduledstart' />
                                            <attribute name='scheduledend' />
                                            <attribute name='createdby' />
                                            <attribute name='regardingobjectid' />
                                            <attribute name='activityid' />
                                            <attribute name='instancetypecode' />
                                            <order attribute='subject' descending='false' />
                                            <filter type='and'>
                                              <condition attribute='statecode' operator='in'>
                                                <value>0</value>
                                                <value>3</value>
                                              </condition>
                                            </filter>
                                            <link-entity name='hil_homeadvisoryline' from='hil_homeadvisorylineid' to='regardingobjectid' link-type='inner' alias='ab'>
                                              <filter type='and'>
                                                <condition attribute='hil_homeadvisorylineid' operator='eq' uiname='' uitype='hil_homeadvisoryline' value='" + req.RecordID + @"' />
                                              </filter>
                                            </link-entity>
                                          </entity>
                                        </fetch>";
                            EntityCollection AdvlineAppointmentColl = service.RetrieveMultiple(new FetchExpression(fetchappointment));
                            if (AdvlineAppointmentColl.Entities.Count > 0)
                            {
                                Entity _app = new Entity("appointment");
                                _app.Id = AdvlineAppointmentColl.Entities[0].Id;
                                _app["statecode"] = new OptionSetValue(1);
                                _app["statuscode"] = new OptionSetValue(3);
                                //_app["description"] = req.Remarks != null ? req.Remarks : "";
                                service.Update(_app);
                            }
                            Entity _appointment = new Entity("appointment");
                            Entity from = new Entity("activityparty");
                            from["partyid"] = cust;
                            _appointment["requiredattendees"] = new Entity[] { from };
                            _appointment["subject"] = "Meeting with Havells for Advisory";
                            if (req.IsVideoMeeting)
                                _appointment["location"] = "Teams";
                            else
                                _appointment["location"] = "Audio Call";
                            _appointment["regardingobjectid"] = new EntityReference("hil_homeadvisoryline", regarding);

                            DateTime SlotDateStartTime = new DateTime(Convert.ToInt32(req.SlotDate.Substring(0, 4)), Convert.ToInt32(req.SlotDate.Substring(4, 2)), Convert.ToInt32(req.SlotDate.Substring(6, 2)), Convert.ToInt32(req.SlotStart.Substring(0, 2)), Convert.ToInt32(req.SlotStart.Substring(3, 2)), 0);
                            DateTime SlotDateEndTime = new DateTime(Convert.ToInt32(req.SlotDate.Substring(0, 4)), Convert.ToInt32(req.SlotDate.Substring(4, 2)), Convert.ToInt32(req.SlotDate.Substring(6, 2)), Convert.ToInt32(req.SlotEnd.Substring(0, 2)), Convert.ToInt32(req.SlotEnd.Substring(3, 2)), 0);

                            _appointment["scheduledstart"] = SlotDateStartTime;
                            _appointment["scheduledend"] = SlotDateEndTime;
                            if (resp.Data.MeetingURL != null && resp.Data.MeetingURL != string.Empty && resp.Data.MeetingURL != "")
                            {
                                _appointment["hil_appointmenturl"] = resp.Data.MeetingURL;
                            }
                            service.Create(_appointment);

                            Entity _enqLine = new Entity(Found.EntityName);
                            _enqLine.Id = Found.Entities[0].Id;

                            _enqLine["hil_appointmentid"] = resp.Data.TransactionId;
                            if (resp.Data.MeetingURL != "" || resp.Data.MeetingURL != null)
                                _enqLine["hil_videocallurl"] = resp.Data.MeetingURL;
                            _enqLine["hil_appointmentstatus"] = new OptionSetValue(6);
                            _enqLine["hil_appointmentdate"] = SlotDateStartTime;
                            _enqLine["hil_appointmentenddate"] = SlotDateEndTime;
                            _enqLine["hil_appointmenttypes"] = req.IsVideoMeeting ? new OptionSetValue(2) : new OptionSetValue(1);
                            _enqLine["hil_enquirystauts"] = new OptionSetValue(2);
                            service.Update(_enqLine);
                            res.Message = resp.Message;
                            res.Status = true;
                        }
                        else
                        {
                            res.Message = resp.Message;
                            return res;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                res.Message = (ex.Message);
                res.Status = false;
                return res;
            }
            return res;
        }
        public HopmeAdvisoryResult CreateAdvisory(HomeAdvisoryRequest reqParm)
        {
            HopmeAdvisoryResult homeadvresults = new HopmeAdvisoryResult();
            try
            {
                HomeAdvisoryRequest homeadv = new HomeAdvisoryRequest();
                IOrganizationService service = ConnectToCRM.GetOrgService();
                #region Variables...
                QueryExpression query = new QueryExpression();
                EntityCollection entCol = new EntityCollection();
                EntityReference _customerType = null;
                EntityReference _propertyType = null;
                EntityReference _constructionType = null;
                EntityReference _city = null;
                EntityReference _state = null;
                EntityReference _pinCode = null;
                EntityReference _contact = null;
                EntityReference _enquirytype = null;
                EntityReference _productType = null;
                String enqProduct = string.Empty;
                String customerName = string.Empty;

                #endregion

                if (service != null)
                {

                    Entity _AdvisoryEnquiry = new Entity("hil_advisoryenquiry");
                    if (reqParm.Salutation == string.Empty || reqParm.Salutation == null || reqParm.Salutation == "") {
                        reqParm.Salutation = "1";
                    }
                    if (reqParm.CustomerType != string.Empty && reqParm.CustomerType != null)
                    {
                        #region hil_typeofcustomer...
                        query = new QueryExpression("hil_typeofcustomer");
                        query.ColumnSet = new ColumnSet("hil_name", "hil_typeofcustomerid");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, reqParm.CustomerType);
                        entCol = service.RetrieveMultiple(query);
                        if (entCol.Entities.Count > 0)
                        {
                            _customerType = new EntityReference(entCol.EntityName, entCol[0].Id);
                        }
                        else
                        {
                            homeadvresults.statusCode = "204";
                            homeadvresults.statusDiscription = "CustomerType does not Exist";
                            return homeadvresults;
                        }
                        #endregion
                    }
                    if (reqParm.PropertyType != string.Empty && reqParm.PropertyType != null)
                    {
                        #region PropertyType fetch...
                        query = new QueryExpression("hil_propertytype");
                        query.ColumnSet = new ColumnSet("hil_name", "hil_propertytypeid");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, reqParm.PropertyType);
                        entCol = service.RetrieveMultiple(query);
                        if (entCol.Entities.Count > 0)
                        {
                            _propertyType = new EntityReference(entCol.EntityName, entCol[0].Id);
                        }
                        else
                        {
                            homeadvresults.statusCode = "204";
                            homeadvresults.statusDiscription = "PropertyType does not Exist";
                            return homeadvresults;
                        }
                        #endregion
                    }
                    if (reqParm.ConsturctionType != string.Empty && reqParm.ConsturctionType != null)
                    {
                        #region constructionType Fetch
                        query = new QueryExpression("hil_constructiontype");
                        query.ColumnSet = new ColumnSet("hil_name", "hil_constructiontypeid");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, reqParm.ConsturctionType);
                        entCol = service.RetrieveMultiple(query);
                        if (entCol.Entities.Count > 0)
                        {
                            _constructionType = new EntityReference(entCol.EntityName, entCol[0].Id);
                            //_AdvisoryEnquiry.Attributes["hil_constructiontype"] = _constructionType;
                        }
                        else
                        {
                            homeadvresults.statusCode = "204";
                            homeadvresults.statusDiscription = "ConstructionType does not Exist";
                            return homeadvresults;
                        }
                        #endregion
                    }
                    if (reqParm.city != string.Empty && reqParm.city != null)
                    {
                        #region CityFetch
                        query = new QueryExpression("hil_city");
                        query.ColumnSet = new ColumnSet("hil_name", "hil_cityid");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        query.Criteria.AddCondition("hil_citycode", ConditionOperator.Equal, reqParm.city);
                        entCol = service.RetrieveMultiple(query);
                        if (entCol.Entities.Count > 0)
                        {
                            _city = new EntityReference(entCol.EntityName, entCol[0].Id);
                        }
                        else
                        {
                            homeadvresults.statusCode = "204";
                            homeadvresults.statusDiscription = "City does not Exist";
                            return homeadvresults;
                        }
                        #endregion
                    }
                    
                    if (reqParm.state != string.Empty && reqParm.state != null)
                    {
                        #region StateFetch
                        query = new QueryExpression("hil_state");
                        query.ColumnSet = new ColumnSet("hil_name", "hil_stateid");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        query.Criteria.AddCondition("hil_statecode", ConditionOperator.Equal, reqParm.state);
                        entCol = service.RetrieveMultiple(query);
                        if (entCol.Entities.Count > 0)
                        {
                            _state = new EntityReference(entCol.EntityName, entCol[0].Id);
                        }
                        else
                        {
                            homeadvresults.statusCode = "204";
                            homeadvresults.statusDiscription = "State does not Exist";
                            return homeadvresults;
                        }
                        #endregion
                    }
                    
                    if (reqParm.pincode != string.Empty && reqParm.pincode != null)
                    {
                        #region PincodeFetch
                        query = new QueryExpression("hil_businessmapping");
                        query.ColumnSet = new ColumnSet("hil_state", "hil_city", "hil_pincode");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        query.Criteria.AddCondition("hil_pincodename", ConditionOperator.Like, "%" + reqParm.pincode + "%");
                        entCol = service.RetrieveMultiple(query);
                        if (entCol.Entities.Count > 0)
                        {
                            _pinCode = entCol[0].GetAttributeValue<EntityReference>("hil_pincode");
                            _city = entCol[0].GetAttributeValue<EntityReference>("hil_city");
                            _state = entCol[0].GetAttributeValue<EntityReference>("hil_state");
                        }
                        else
                        {
                            homeadvresults.statusCode = "204";
                            homeadvresults.statusDiscription = "Pincode does not Exist";
                            return homeadvresults;
                        }
                        #endregion
                    }
                    else
                    {
                        homeadvresults.statusCode = "204";
                        homeadvresults.statusDiscription = "Pincode is Empty";
                        return homeadvresults;
                    }
                    if (reqParm.mobilenumber != string.Empty && reqParm.mobilenumber != null)
                    {
                        #region customerFetch
                        query = new QueryExpression("contact");
                        query.ColumnSet = new ColumnSet("contactid", "fullname", "emailaddress1", "hil_salutation");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, reqParm.mobilenumber);
                        entCol = service.RetrieveMultiple(query);
                        if (entCol.Entities.Count > 0)
                        {
                            _contact = new EntityReference(entCol.EntityName, entCol[0].Id);
                            customerName = entCol[0].GetAttributeValue<string>("fullname");
                        }
                        else
                        {
                            Entity entConsumer = new Entity("contact");
                            entConsumer["mobilephone"] = reqParm.mobilenumber;
                            if (reqParm.Salutation != string.Empty || reqParm.Salutation != null)
                                entConsumer["hil_salutation"] = new OptionSetValue(int.Parse(reqParm.Salutation));
                            if (reqParm.CustomerName != string.Empty || reqParm.CustomerName != null)
                            {
                                string[] consumerName = reqParm.CustomerName.Split(' ');
                                if (consumerName.Length >= 1)
                                {
                                    entConsumer["firstname"] = consumerName[0];
                                    if (consumerName.Length == 3)
                                    {
                                        entConsumer["middlename"] = consumerName[1];
                                        entConsumer["lastname"] = consumerName[2];
                                    }
                                    if (consumerName.Length == 2)
                                    {
                                        entConsumer["lastname"] = consumerName[1];
                                    }
                                }
                                else
                                {
                                    entConsumer["firstname"] = reqParm.CustomerName;
                                }
                            }

                            if (reqParm.emailid != string.Empty || reqParm.emailid != null)
                                entConsumer["emailaddress1"] = reqParm.emailid;

                            Guid consumerGuId = service.Create(entConsumer);
                            _contact = new EntityReference(entConsumer.LogicalName, consumerGuId);
                            customerName = reqParm.CustomerName;
                        }
                        #endregion
                    }
                    else
                    {
                        homeadvresults.statusCode = "204";
                        homeadvresults.statusDiscription = "Mobile Number is Empty";
                        return homeadvresults;
                    }
                    if (_customerType != null)
                    {
                        _AdvisoryEnquiry.Attributes["hil_typeofcustomer"] = _customerType;
                    }
                    _AdvisoryEnquiry.Attributes["hil_customer"] = _contact;
                    if (reqParm.Area != null)
                    {
                        _AdvisoryEnquiry.Attributes["hil_areasqrt"] = reqParm.Area.ToString();
                    }
                    if (_propertyType != null)
                    {
                        _AdvisoryEnquiry.Attributes["hil_propertytype"] = _propertyType;
                    }
                    if (_constructionType != null)
                    {
                        _AdvisoryEnquiry.Attributes["hil_constructiontype"] = _constructionType;
                    }
                    if (reqParm.rooftop != null)
                    {
                        _AdvisoryEnquiry.Attributes["hil_rooftop"] = reqParm.rooftop;
                    }
                    if (reqParm.assettype != null)
                    {
                        _AdvisoryEnquiry.Attributes["hil_assettype"] = reqParm.assettype.ToString();
                    }
                    if (reqParm.mobilenumber != null)
                    {
                        _AdvisoryEnquiry.Attributes["hil_mobilenumber"] = reqParm.mobilenumber.ToString();
                    }
                    else
                    {
                        homeadvresults.statusDiscription = "mobilenumber is mandatory";
                        homeadvresults.statusCode = "204";
                        return homeadvresults;
                    }
                    if (reqParm.emailid != null)
                    {
                        _AdvisoryEnquiry.Attributes["hil_emailid"] = reqParm.emailid.ToString();
                    }
                    if (reqParm.tds != null)
                    {
                        _AdvisoryEnquiry.Attributes["hil_tds"] = reqParm.tds.ToString();
                    }
                    if (reqParm.CustomerRemarks != null && reqParm.CustomerRemarks != "")
                    {
                        _AdvisoryEnquiry.Attributes["hil_customerremarks"] = reqParm.CustomerRemarks;
                    }
                    if (reqParm.SourceofCreation != null && reqParm.SourceofCreation != "")
                    {
                        _AdvisoryEnquiry.Attributes["hil_sourceofcreation"] = new OptionSetValue(int.Parse(reqParm.SourceofCreation));
                    }
                    else
                    {
                        homeadvresults.statusDiscription = "Source of Creation does not Exist";
                        homeadvresults.statusCode = "204";
                        return homeadvresults;
                    }

                    if (_city != null)
                        _AdvisoryEnquiry.Attributes["hil_city"] = _city;
                    if (_state != null)
                        _AdvisoryEnquiry.Attributes["hil_state"] = _state;

                    _AdvisoryEnquiry.Attributes["hil_pincode"] = _pinCode;
                    OptionSetValueCollection productColl = new OptionSetValueCollection();
                    for (int i = 0; i < reqParm.advisoryenquryline.Count; i++)
                    {
                        string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_typeofproduct'>
                            <attribute name='hil_index' />
                            <filter type='and'>
                                <condition attribute='hil_code' operator='eq' value='" + reqParm.advisoryenquryline[i].ProductType + @"' />
                            </filter>
                            </entity>
                            </fetch>";
                        EntityCollection entColProdType = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entColProdType.Entities.Count > 0)
                        {
                            productColl.Add(new OptionSetValue(Int32.Parse(entColProdType[0].GetAttributeValue<String>("hil_index"))));
                        }
                    }
                    if (reqParm.advisoryenquryline[0].EnquryTypecode != null)
                    {
                        #region EnquryTypeFetch
                        query = new QueryExpression("hil_enquirytype");
                        query.ColumnSet = new ColumnSet(false);
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        query.Criteria.AddCondition("hil_enquirytypecode", ConditionOperator.Equal, reqParm.advisoryenquryline[0].EnquryTypecode);
                        entCol = service.RetrieveMultiple(query);
                        if (entCol.Entities.Count > 0)
                        {
                            _enquirytype = new EntityReference(entCol.EntityName, entCol[0].Id);
                        }
                        else
                        {
                            homeadvresults.statusCode = "204";
                            homeadvresults.statusDiscription = "EnquryType does not Exist";
                            return homeadvresults;
                        }
                        #endregion
                    }
                    _AdvisoryEnquiry["hil_typeofenquiry"] = _enquirytype;//588,21,588,27
                    _AdvisoryEnquiry["hil_typeofproduct"] = productColl;
                    Guid homeadvisory = service.Create(_AdvisoryEnquiry);
                    homeadvresults.enquiryGuId = homeadvisory.ToString();
                    homeadvresults.enquiryid = service.Retrieve("hil_advisoryenquiry", homeadvisory, new ColumnSet("hil_name")).GetAttributeValue<string>("hil_name").ToString();

                    for (int i = 0; i < reqParm.advisoryenquryline.Count; i++)
                    {

                        Entity AdvisoryEnquryLine = new Entity("hil_homeadvisoryline");
                        AdvisoryEnquryLine.Attributes["hil_advisoryenquery"] = new EntityReference("hil_advisoryenquiry", homeadvisory);

                        if (reqParm.advisoryenquryline[i].ProductType != null)
                        {
                            #region ProductTypeFetch
                            query = new QueryExpression("hil_typeofproduct");
                            query.ColumnSet = new ColumnSet("hil_name");
                            query.Criteria = new FilterExpression(LogicalOperator.And);
                            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, reqParm.advisoryenquryline[i].ProductType);
                            entCol = service.RetrieveMultiple(query);
                            if (entCol.Entities.Count > 0)
                            {
                                _productType = new EntityReference(entCol.EntityName, entCol[0].Id);
                                enqProduct = enqProduct != string.Empty ? enqProduct + "," + entCol[0].GetAttributeValue<string>("hil_name") : entCol[0].GetAttributeValue<string>("hil_name");
                            }
                            else
                            {
                                homeadvresults.statusCode = "204";
                                homeadvresults.statusDiscription = "ProductType does not Exist";
                                return homeadvresults;
                            }
                            #endregion
                        }
                        if (reqParm.advisoryenquryline[i].Enqurystatus != null && reqParm.advisoryenquryline[i].Enqurystatus != "")
                        {
                            AdvisoryEnquryLine.Attributes["hil_enquirystauts"] = new OptionSetValue(int.Parse(reqParm.advisoryenquryline[0].Enqurystatus));
                        }
                        else
                        {
                            AdvisoryEnquryLine.Attributes["hil_enquirystauts"] = new OptionSetValue(1);
                        }
                        if (reqParm.advisoryenquryline[0].ApointmentType != null)
                        {
                            AdvisoryEnquryLine.Attributes["hil_appointmenttypes"] = new OptionSetValue(int.Parse(reqParm.advisoryenquryline[i].ApointmentType));
                        }
                        else
                        {
                            //AdvisoryEnquryLine.Attributes["hil_appointmenttypes"] = new OptionSetValue(910590001);
                        }
                        if (_enquirytype.Id != Guid.Empty)
                        {
                            AdvisoryEnquryLine.Attributes["hil_typeofenquiiry"] = _enquirytype;
                        }
                        if (_productType.Id != Guid.Empty)
                        {
                            AdvisoryEnquryLine.Attributes["hil_typeofproduct"] = _productType;
                        }
                        AdvisoryEnquryLine.Attributes["hil_mobilenumber"] = reqParm.mobilenumber;
                        AdvisoryEnquryLine.Attributes["hil_pincode"] = _pinCode;
                        Guid _lineID = service.Create(AdvisoryEnquryLine);
                        
                        if (reqParm.advisoryenquryline[i].Attachments != null)
                            for (int j = 0; j < reqParm.advisoryenquryline[i].Attachments.Count; j++)
                            {
                                if (reqParm.advisoryenquryline[i].Attachments[j].Subject == null || reqParm.advisoryenquryline[i].Attachments[j].Subject == "" || 
                                    reqParm.advisoryenquryline[i].Attachments[j].FileName == null || reqParm.advisoryenquryline[i].Attachments[j].FileName == "" ||
                                    reqParm.advisoryenquryline[i].Attachments[j].FileSize == null || reqParm.advisoryenquryline[i].Attachments[j].FileSize == "" ||
                                    reqParm.advisoryenquryline[i].Attachments[j].FileString == null || reqParm.advisoryenquryline[i].Attachments[j].FileString == "" ||
                                    reqParm.advisoryenquryline[i].Attachments[j].DocumentType == null || reqParm.advisoryenquryline[i].Attachments[j].DocumentType == "")
                                {
                                    ////
                                }
                                else
                                {
                                    string blobURL = string.Empty;
                                    try
                                    {
                                        EntityReference _doctypee = new EntityReference(); //getOptionSetValue(reqParm.DocumentType, "hil_attachment", "hil_doctype", service);
                                        Guid loginUserGuid = ((WhoAmIResponse)service.Execute(new WhoAmIRequest())).UserId;
                                        query = new QueryExpression("hil_attachmentdocumenttype");
                                        query.ColumnSet = new ColumnSet("hil_containername");
                                        query.Criteria = new FilterExpression(LogicalOperator.And);
                                        query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                        query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, reqParm.advisoryenquryline[i].Attachments[j].DocumentType);
                                        entCol = service.RetrieveMultiple(query);
                                        String containerName = string.Empty;
                                        if (entCol.Entities.Count == 0)
                                        {
                                            ///
                                        }
                                        else
                                        {
                                            _doctypee = new EntityReference(entCol[0].LogicalName, entCol[0].Id);
                                            containerName = entCol[0].Contains("hil_containername") ? entCol[0].GetAttributeValue<string>("hil_containername") : "devanduat";
                                        }
                                        string[] ext = reqParm.advisoryenquryline[i].Attachments[j].FileName.Split('.');
                                        String fileName = reqParm.advisoryenquryline[i].Attachments[j].RegardingGuid + "_" + reqParm.advisoryenquryline[i].Attachments[j].DocumentType + "_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + "." + ext[ext.Length - 1];
                                        byte[] fileContent = Convert.FromBase64String(reqParm.advisoryenquryline[i].Attachments[j].FileString);



                                        blobURL = Upload(fileName, fileContent, containerName);



                                        Entity _attachment = new Entity("hil_attachment");
                                        _attachment["subject"] = reqParm.advisoryenquryline[i].Attachments[j].FileName;



                                        _attachment["hil_documenttype"] = _doctypee;
                                        _attachment["hil_docurl"] = blobURL;
                                        _attachment["regardingobjectid"] = new EntityReference("hil_homeadvisoryline", _lineID);
                                        _attachment["hil_sourceofdocument"] = new OptionSetValue(2);
                                        try
                                        {
                                            service.Create(_attachment);
                                        }
                                        catch (Exception ex)
                                        {
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                    }



                                }
                            }
                    }
                    homeadvresults.statusCode = "200";
                    homeadvresults.statusDiscription = "Success";
                    homeadvresults.message = "Record Created Successfully";
                    return homeadvresults;
                }
            }
            catch (Exception ex)
            {
                homeadvresults.statusCode = "505";
                homeadvresults.statusDiscription = "D365 Internal Error " + ex.Message;
                return homeadvresults;
            }
            return homeadvresults;
        }
        public Response RescheduleAppointment(ReschduleAppointment req)
        {
            Response res = new Response();
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();


                string fetchAdvisoryline = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='hil_homeadvisoryline'>
                    <attribute name='hil_homeadvisorylineid' />
                    <attribute name='hil_advisoryenquery' />
                    <attribute name='hil_name' />
                    <attribute name='createdon' />
                    <order attribute='hil_name' descending='false' />
                    <filter type='and'>
                    <condition attribute='statecode' operator='eq' value='0' />
                    <condition attribute='hil_name' operator='eq' value='" + req.AdvisorylineId + @"' />
                    </filter>
                    </entity>
                    </fetch>";
                EntityCollection AdvisoryLineColl = service.RetrieveMultiple(new FetchExpression(fetchAdvisoryline));
                if (AdvisoryLineColl.Entities.Count == 0)
                {
                    res.Message = "Advisory Enquery Line Not Found in Dynamics";
                    res.Status = false;
                    return res;
                }
                Entity Advisoryline = AdvisoryLineColl.Entities[0];

                QueryExpression Query = new QueryExpression("hil_advisoryenquiry");
                Query.ColumnSet = new ColumnSet("hil_customer");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Advisoryline.GetAttributeValue<EntityReference>("hil_advisoryenquery").Name);
                EntityCollection Enquirycoll = service.RetrieveMultiple(Query);

                string fetchappointment = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='appointment'>
                                            <attribute name='subject' />
                                            <attribute name='statecode' />
                                            <attribute name='scheduledstart' />
                                            <attribute name='scheduledend' />
                                            <attribute name='createdby' />
                                            <attribute name='regardingobjectid' />
                                            <attribute name='activityid' />
                                            <attribute name='instancetypecode' />
                                            <order attribute='subject' descending='false' />
                                            <filter type='and'>
                                              <condition attribute='statecode' operator='in'>
                                                <value>0</value>
                                                <value>3</value>
                                              </condition>
                                            </filter>
                                            <link-entity name='hil_homeadvisoryline' from='hil_homeadvisorylineid' to='regardingobjectid' link-type='inner' alias='ab'>
                                              <filter type='and'>
                                                <condition attribute='hil_homeadvisorylineid' operator='eq' uiname='AdvEnq-000060' uitype='hil_homeadvisoryline' value='" + AdvisoryLineColl.Entities[0].Id + @"' />
                                              </filter>
                                            </link-entity>
                                          </entity>
                                        </fetch>";
                EntityCollection AdvlineAppointmentColl = service.RetrieveMultiple(new FetchExpression(fetchappointment));
                if (AdvlineAppointmentColl.Entities.Count > 0)
                {

                    Entity _app = new Entity("appointment");
                    _app.Id = AdvlineAppointmentColl.Entities[0].Id;
                    _app["statecode"] = new OptionSetValue(1);
                    _app["statuscode"] = new OptionSetValue(3);
                    _app["description"] = req.Remarks != null ? req.Remarks : "";

                    service.Update(_app);
                }
                Entity _appointment = new Entity("appointment");
                Entity from = new Entity("activityparty");
                EntityReference cust = Enquirycoll.Entities[0].GetAttributeValue<EntityReference>("hil_customer");
                from["partyid"] = cust;
                _appointment["requiredattendees"] = new Entity[] { from };
                _appointment["subject"] = "Meeting with Havells for Advisory";
                _appointment["location"] = "Teams";
                _appointment["regardingobjectid"] = new EntityReference("hil_homeadvisoryline", AdvisoryLineColl.Entities[0].Id);
                _appointment["scheduledstart"] = Convert.ToDateTime(req.scheduledstart);
                _appointment["scheduledend"] = Convert.ToDateTime(req.scheduledEnd);
                _appointment["hil_appointmenturl"] = req.appointmenturl.ToString();
                service.Create(_appointment);

                Entity _advLine = new Entity("hil_homeadvisoryline");
                _advLine.Id = AdvisoryLineColl.Entities[0].Id;
                if (req.appintmentId != null)
                    _advLine["hil_appointmentid"] = req.appintmentId;
                if (req.appointmenturl != null || req.appointmenturl != "")
                    _advLine["hil_videocallurl"] = req.appointmenturl;
                _advLine["hil_appointmentstatus"] = new OptionSetValue(6);

                _advLine["hil_appointmentdate"] = Convert.ToDateTime(req.scheduledstart);
                _advLine["hil_appointmentenddate"] = Convert.ToDateTime(req.scheduledEnd);

                _advLine["hil_appointmenttypes"] = req.appintmentType == "Video" ? new OptionSetValue(2) : new OptionSetValue(1);
                if (req.AssignedUsercode != null)
                {
                    QueryExpression advmasterQuery = new QueryExpression("hil_advisormaster");
                    advmasterQuery.ColumnSet = new ColumnSet(false);
                    advmasterQuery.Criteria = new FilterExpression(LogicalOperator.And);
                    advmasterQuery.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    advmasterQuery.Criteria.AddCondition("hil_code", ConditionOperator.Equal, req.AssignedUsercode.Trim());
                    EntityCollection Advisorymastercoll = service.RetrieveMultiple(advmasterQuery);
                    if (Advisorymastercoll.Entities.Count == 0)
                    {
                        res.Message = "Advisor Not Found in Dynamics";
                        res.Status = false;
                        return res;
                    }
                    _advLine["hil_assignedadvisor"] = new EntityReference("hil_advisormaster", Advisorymastercoll[0].Id);
                }

                service.Update(_advLine);
                res.Message = "Appointment reschdule sucessfully";
                res.Status = true;
                return res;
            }
            catch (Exception ex)
            {
                res.Message = (ex.Message);
                res.Status = false;
                return res;
            }
            return res;
        }
        public Response CancleAppointment(CancelAppointmentRequest req)
        {
            Response resp = new Response();
            IOrganizationService service = ConnectToCRM.GetOrgService();
            if (req.AdvisorylineGuid != null)
            {
                resp = CancleTeamAppointment(new Guid(req.AdvisorylineGuid), req.AppointmentStatus, req.AppointmentRemarks, service);
                return resp;
            }
            else
            {
                try
                {
                    string fetchAdvisoryline = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='hil_homeadvisoryline'>
                    <attribute name='hil_homeadvisorylineid' />
                    <attribute name='hil_name' />
                    <attribute name='createdon' />
                    <order attribute='hil_name' descending='false' />
                    <filter type='and'>
                    <condition attribute='statecode' operator='eq' value='0' />
                    <condition attribute='hil_name' operator='eq' value='" + req.AdvisorylineId.Trim() + @"' />
                    </filter>
                    </entity>
                    </fetch>";
                    EntityCollection AdvisoryLineColl = service.RetrieveMultiple(new FetchExpression(fetchAdvisoryline));
                    if (AdvisoryLineColl.Entities.Count == 0)
                    {
                        resp.Message = "Advisory Enquery Line Not Found in Dynamics";
                        resp.Status = false;
                        return resp;
                    }
                    Entity Advisoryline = AdvisoryLineColl.Entities[0];
                    if (req.IsEnquiryClosed)
                    {
                        EntityReference cancleReAON = new EntityReference();
                        QueryExpression query = new QueryExpression("hil_cancellationreason");
                        String _fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                            <entity name='hil_cancellationreason'>
                                            <attribute name='hil_cancellationreasonid' />
                                            <attribute name='hil_name' />
                                            <order attribute='hil_name' descending='false' />
                                            <filter type='and'>
                                            <condition attribute='hil_name' operator='eq' value='" + req.EnquiryRemarks + @"' />
                                            </filter>
                                           </entity>
                                            </fetch>";
                        EntityCollection reasonColl = service.RetrieveMultiple(new FetchExpression(_fetch));
                        if (reasonColl.Entities.Count > 0)
                        {
                            cancleReAON = new EntityReference(reasonColl.EntityName, reasonColl[0].Id);
                        }
                        else
                        {
                            resp.Message = "Cancelation Reasion Not found in Dynamics.";
                            resp.Status = false;
                            return resp;
                        }
                        Entity _advLine = new Entity(Advisoryline.LogicalName);
                        _advLine.Id = Advisoryline.Id;
                        _advLine["hil_enquirystauts"] = new OptionSetValue(int.Parse(req.EnquiryStatus));
                        //_advLine["hil_customerremark"] = req.EnquiryRemarks;
                        _advLine["hil_closingreasion"] = cancleReAON;
                        //service.Update(_advLine);
                        bool _apSt = closeAppointment(Advisoryline.Id.ToString(), req.AppointmentStatus, req.AppointmentRemarks, service);
                        if (_apSt)
                        {
                            _advLine["hil_appointmentstatus"] = new OptionSetValue(3);
                            service.Update(_advLine);
                            resp.Message = "Enquery and Appointment cancled Sucessfully";
                            resp.Status = true;
                            return resp;
                        }
                        else
                        {
                            service.Update(_advLine);
                            resp.Message = "Enquery Closed Sucessfully";
                            resp.Status = true;
                            return resp;
                        }
                    }
                    else
                    {
                        bool _apSt = closeAppointment(Advisoryline.Id.ToString(), req.AppointmentStatus, req.AppointmentRemarks, service);
                        if (_apSt)
                        {
                            Entity _advLine = new Entity(Advisoryline.LogicalName);
                            _advLine.Id = Advisoryline.Id;
                            _advLine["hil_appointmentstatus"] = new OptionSetValue(4); // 3- Completed, 4- Cancel
                            service.Update(_advLine);
                            resp.Message = "Appointment Closed Sucessfully";
                            resp.Status = true;
                            return resp;
                        }
                        else
                        {
                            resp.Message = "Appointment Not Closed ";
                            resp.Status = false;
                            return resp;
                        }
                    }
                }
                catch (Exception ex)
                {
                    resp.Message = "D365 Internal Server Error: " + ex.Message;
                    resp.Status = false;
                    return resp;
                }
            }
        }

        public GetEnquiry GetEnquery(Enquiry req)
        {
            GetEnquiry _retResponse = new GetEnquiry();
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (req.AdvisoryID != null)
                {
                    string enquiryfetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='hil_advisoryenquiry'>
                                            <attribute name='hil_advisoryenquiryid' />
                                            <attribute name='hil_name' />
                                            <attribute name='createdon' />
                                            <attribute name='hil_typeofcustomer' />
                                            <attribute name='hil_tds' />
                                            <attribute name='hil_state' />
                                            <attribute name='hil_rooftop' />
                                            <attribute name='hil_propertytype' />
                                            <attribute name='hil_pincode' />
                                            <attribute name='hil_mobilenumber' />
                                            <attribute name='hil_emailid' />
                                            <attribute name='hil_customer' />
                                            <attribute name='hil_constructiontype' />
                                            <attribute name='hil_city' />
                                            <attribute name='hil_assettype' />
                                            <attribute name='hil_areasqrt' />
                                            <attribute name='hil_customerremarks' />
                                            <order attribute='hil_name' descending='false' />
                                            <filter type='and'>
                                              <condition attribute='hil_name' operator='eq' value='" + req.AdvisoryID + @"' />
                                              <condition attribute='statecode' operator='eq' value='0' />
                                            </filter>
                                          </entity>
                                        </fetch>";
                    EntityCollection _addvisorycoll = service.RetrieveMultiple(new FetchExpression(enquiryfetch));
                    if (_addvisorycoll.Entities.Count > 0)
                    {
                        List<GetHomeAdvisoryEnquiry> homeadvisoryenqList = new List<GetHomeAdvisoryEnquiry>();
                        GetHomeAdvisoryEnquiry homeadvisory = new GetHomeAdvisoryEnquiry();
                        if (_addvisorycoll[0].Contains("hil_advisoryenquiryid"))
                            homeadvisory.AdvisoryGuid = _addvisorycoll[0].Id.ToString();
                        homeadvisory.AdvisoryID = _addvisorycoll[0].Contains("hil_name") ? _addvisorycoll[0].GetAttributeValue<string>("hil_name") : "";
                        homeadvisory.Area = _addvisorycoll[0].Contains("hil_areasqrt") ? _addvisorycoll[0].GetAttributeValue<string>("hil_areasqrt") : "";
                        homeadvisory.CustomerType = _addvisorycoll[0].Contains("hil_typeofcustomer") ? _addvisorycoll[0].GetAttributeValue<EntityReference>("hil_typeofcustomer").Name : "";
                        homeadvisory.TDS = _addvisorycoll[0].Contains("hil_tds") ? _addvisorycoll[0].GetAttributeValue<string>("hil_tds") : "";
                        homeadvisory.CustomerRemarks = _addvisorycoll[0].Contains("hil_customerremarks") ? _addvisorycoll[0].GetAttributeValue<string>("hil_customerremarks") : "";
                        homeadvisory.State = _addvisorycoll[0].Contains("hil_state") ? _addvisorycoll[0].GetAttributeValue<EntityReference>("hil_state").Name : "";
                        homeadvisory.RoofTop = _addvisorycoll[0].GetAttributeValue<bool>("hil_rooftop") ? "Yes" : "No";
                        homeadvisory.PropertyType = _addvisorycoll[0].Contains("hil_propertytype") ? _addvisorycoll[0].GetAttributeValue<EntityReference>("hil_propertytype").Name : "";
                        homeadvisory.ConsturctionType = _addvisorycoll[0].Contains("hil_constructiontype") ? _addvisorycoll[0].GetAttributeValue<EntityReference>("hil_constructiontype").Name : "";
                        homeadvisory.AssetType = _addvisorycoll[0].Contains("hil_assettype") ? _addvisorycoll[0].GetAttributeValue<string>("hil_assettype") : "";
                        homeadvisory.Customer = _addvisorycoll[0].Contains("hil_customer") ? _addvisorycoll[0].GetAttributeValue<EntityReference>("hil_customer").Name : "";
                        homeadvisory.MobileNumber = _addvisorycoll[0].Contains("hil_mobilenumber") ? _addvisorycoll[0].GetAttributeValue<string>("hil_mobilenumber") : "";
                        homeadvisory.EmailID = _addvisorycoll[0].Contains("hil_emailid") ? _addvisorycoll[0].GetAttributeValue<string>("hil_emailid") : "";
                        homeadvisory.City = _addvisorycoll[0].Contains("hil_city") ? _addvisorycoll[0].GetAttributeValue<EntityReference>("hil_city").Name : "";
                        homeadvisory.PinCode = _addvisorycoll[0].Contains("hil_pincode") ? _addvisorycoll[0].GetAttributeValue<EntityReference>("hil_pincode").Name : "";
                        List<GetAdvisoryEnquiryLine> enqlineList = getEnqueryLine(service, _addvisorycoll[0].Id);
                        homeadvisory.AdvisoryEnquiryLine = enqlineList;
                        homeadvisoryenqList.Add(homeadvisory);
                        _retResponse.StatusCode = "200";
                        _retResponse.StatusMessage = "Record Retrived";
                        _retResponse.Success = true;
                        _retResponse.Enquiry = homeadvisoryenqList;
                        return _retResponse;
                    }
                    else
                    {
                        _retResponse.StatusCode = "400";
                        _retResponse.Success = false;
                        _retResponse.StatusMessage = "Enquiry Not found.";
                        return _retResponse;
                    }

                }
                else if (req.MobileNo != null)
                {
                    #region fetch customer data...
                    QueryExpression query = new QueryExpression("contact");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, req.MobileNo.Trim());
                    EntityCollection customercoll = service.RetrieveMultiple(query);

                    if (customercoll.Entities.Count > 0)
                    {
                        string enquiryfetchxml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                  <entity name='hil_advisoryenquiry'>
                                                    <attribute name='hil_advisoryenquiryid' />
                                                    <attribute name='hil_name' />
                                                    <attribute name='createdon' />
                                                    <attribute name='hil_typeofcustomer' />
                                                    <attribute name='hil_tds' />
                                                    <attribute name='hil_state' />
                                                    <attribute name='hil_rooftop' />
                                                    <attribute name='hil_propertytype' />
                                                    <attribute name='hil_pincode' />
                                                    <attribute name='hil_mobilenumber' />
                                                    <attribute name='hil_emailid' />
                                                    <attribute name='hil_customer' />
                                                    <attribute name='hil_constructiontype' />
                                                    <attribute name='hil_city' />
                                                    <attribute name='hil_assettype' />
                                                    <attribute name='hil_areasqrt' />
                                                    <order attribute='hil_name' descending='false' />
                                                    <filter type='and'>
                                                      <condition attribute='statecode' operator='eq' value='0' />
                                                      <condition attribute='hil_customer' operator='eq' uiname='' uitype='contact' value='" + customercoll.Entities[0].Id + @"' />
                                                    </filter>
                                                  </entity>
                                                </fetch>";
                        EntityCollection _enqlinecoll = service.RetrieveMultiple(new FetchExpression(enquiryfetchxml));
                        if (_enqlinecoll.Entities.Count > 0)
                        {
                            List<GetHomeAdvisoryEnquiry> homeadvisoryenqList = new List<GetHomeAdvisoryEnquiry>();


                            foreach (Entity adv in _enqlinecoll.Entities)
                            {
                                GetHomeAdvisoryEnquiry homeadvisory = new GetHomeAdvisoryEnquiry();
                                if (adv.Contains("hil_advisoryenquiryid"))
                                    homeadvisory.AdvisoryGuid = adv.Id.ToString();
                                homeadvisory.AdvisoryID = adv.Contains("hil_name") ? adv.GetAttributeValue<string>("hil_name") : "";
                                homeadvisory.Area = adv.Contains("hil_areasqrt") ? adv.GetAttributeValue<string>("hil_areasqrt") : "";
                                homeadvisory.CustomerType = adv.Contains("hil_typeofcustomer") ? adv.GetAttributeValue<EntityReference>("hil_typeofcustomer").Name : "";
                                homeadvisory.TDS = adv.Contains("hil_tds") ? adv.GetAttributeValue<string>("hil_tds") : "";
                                homeadvisory.State = adv.Contains("hil_state") ? adv.GetAttributeValue<EntityReference>("hil_state").Name : "";
                                homeadvisory.RoofTop = adv.GetAttributeValue<bool>("hil_rooftop") ? "Yes" : "No";
                                homeadvisory.PropertyType = adv.Contains("hil_propertytype") ? adv.GetAttributeValue<EntityReference>("hil_propertytype").Name : "";
                                homeadvisory.ConsturctionType = adv.Contains("hil_constructiontype") ? adv.GetAttributeValue<EntityReference>("hil_constructiontype").Name : "";
                                homeadvisory.AssetType = adv.Contains("hil_assettype") ? adv.GetAttributeValue<string>("hil_assettype") : "";
                                homeadvisory.Customer = adv.Contains("hil_customer") ? adv.GetAttributeValue<EntityReference>("hil_customer").Name : "";
                                homeadvisory.MobileNumber = adv.Contains("hil_mobilenumber") ? adv.GetAttributeValue<string>("hil_mobilenumber") : "";
                                homeadvisory.EmailID = adv.Contains("hil_emailid") ? adv.GetAttributeValue<string>("hil_emailid") : "";
                                homeadvisory.City = adv.Contains("hil_city") ? adv.GetAttributeValue<EntityReference>("hil_city").Name : "";
                                homeadvisory.PinCode = adv.Contains("hil_pincode") ? adv.GetAttributeValue<EntityReference>("hil_pincode").Name : "";
                                List<GetAdvisoryEnquiryLine> enqlineList = getEnqueryLine(service, adv.Id);
                                homeadvisory.AdvisoryEnquiryLine = enqlineList;
                                homeadvisoryenqList.Add(homeadvisory);
                            }
                            _retResponse.StatusCode = "200";
                            _retResponse.StatusMessage = "Record Retrived";
                            _retResponse.Success = true;
                            _retResponse.Enquiry = homeadvisoryenqList;
                            return _retResponse;
                        }
                        else
                        {
                            _retResponse.StatusCode = "400";
                            _retResponse.Success = true;
                            _retResponse.StatusMessage = "Record Not Found";
                            return _retResponse;
                        }
                    }
                    else
                    {
                        _retResponse.StatusCode = "400";
                        _retResponse.StatusMessage = "Contact Not Found";
                        _retResponse.Success = true;
                        return _retResponse;
                    }
                    #endregion
                }
            }
            catch (Exception ex)
            {
                _retResponse.StatusCode = "505";
                _retResponse.Success = false;
                _retResponse.StatusMessage = "Dynamics 365 Internal Server Error : " + ex.Message;
                return _retResponse;
            }
            return _retResponse;
        }

        public GetEnquiry GetEnqueryStatus(EnquiryStatus req)
        {
            GetEnquiry _retResponse = new GetEnquiry();
            string _datefilter = string.Empty;

            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (req.FromDate != null && req.ToDate != null)
                {
                    _datefilter = @"<condition attribute='createdon' operator='on-or-after' value='" + req.FromDate + @"' />
                    <condition attribute='createdon' operator='on-or-before' value='" + req.ToDate + @"' />";
                }
                if (req.EnquiryId != null)
                {
                    string enquiryfetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='hil_advisoryenquiry'>
                                            <attribute name='hil_advisoryenquiryid' />
                                            <attribute name='hil_name' />
                                            <attribute name='createdon' />
                                            <attribute name='hil_typeofcustomer' />
                                            <attribute name='hil_tds' />
                                            <attribute name='hil_state' />
                                            <attribute name='hil_rooftop' />
                                            <attribute name='hil_propertytype' />
                                            <attribute name='hil_pincode' />
                                            <attribute name='hil_mobilenumber' />
                                            <attribute name='hil_emailid' />
                                            <attribute name='hil_customer' />
                                            <attribute name='hil_constructiontype' />
                                            <attribute name='hil_city' />
                                            <attribute name='hil_assettype' />
                                            <attribute name='hil_areasqrt' />
                                            <attribute name='hil_customerremarks' />
                                            <order attribute='hil_name' descending='false' />
                                            <filter type='and'>
                                              <condition attribute='hil_name' operator='eq' value='" + req.EnquiryId + @"' />
                                              <condition attribute='statecode' operator='eq' value='0' />";
                    if (_datefilter != string.Empty)
                    {
                        enquiryfetch += _datefilter;
                    }
                    enquiryfetch += @"</filter>
                                    </entity>
                                </fetch>";

                    EntityCollection _addvisorycoll = service.RetrieveMultiple(new FetchExpression(enquiryfetch));
                    if (_addvisorycoll.Entities.Count > 0)
                    {
                        List<GetHomeAdvisoryEnquiry> homeadvisoryenqList = new List<GetHomeAdvisoryEnquiry>();
                        GetHomeAdvisoryEnquiry homeadvisory = new GetHomeAdvisoryEnquiry();
                        if (_addvisorycoll[0].Contains("hil_advisoryenquiryid"))
                            homeadvisory.AdvisoryGuid = _addvisorycoll[0].Id.ToString();
                        homeadvisory.AdvisoryID = _addvisorycoll[0].Contains("hil_name") ? _addvisorycoll[0].GetAttributeValue<string>("hil_name") : "";
                        homeadvisory.Area = _addvisorycoll[0].Contains("hil_areasqrt") ? _addvisorycoll[0].GetAttributeValue<string>("hil_areasqrt") : "";
                        homeadvisory.CustomerType = _addvisorycoll[0].Contains("hil_typeofcustomer") ? _addvisorycoll[0].GetAttributeValue<EntityReference>("hil_typeofcustomer").Name : "";
                        homeadvisory.TDS = _addvisorycoll[0].Contains("hil_tds") ? _addvisorycoll[0].GetAttributeValue<string>("hil_tds") : "";
                        homeadvisory.CustomerRemarks = _addvisorycoll[0].Contains("hil_customerremarks") ? _addvisorycoll[0].GetAttributeValue<string>("hil_customerremarks") : "";
                        homeadvisory.State = _addvisorycoll[0].Contains("hil_state") ? _addvisorycoll[0].GetAttributeValue<EntityReference>("hil_state").Name : "";
                        homeadvisory.RoofTop = _addvisorycoll[0].GetAttributeValue<bool>("hil_rooftop") ? "Yes" : "No";
                        homeadvisory.PropertyType = _addvisorycoll[0].Contains("hil_propertytype") ? _addvisorycoll[0].GetAttributeValue<EntityReference>("hil_propertytype").Name : "";
                        homeadvisory.ConsturctionType = _addvisorycoll[0].Contains("hil_constructiontype") ? _addvisorycoll[0].GetAttributeValue<EntityReference>("hil_constructiontype").Name : "";
                        homeadvisory.AssetType = _addvisorycoll[0].Contains("hil_assettype") ? _addvisorycoll[0].GetAttributeValue<string>("hil_assettype") : "";
                        homeadvisory.Customer = _addvisorycoll[0].Contains("hil_customer") ? _addvisorycoll[0].GetAttributeValue<EntityReference>("hil_customer").Name : "";
                        homeadvisory.MobileNumber = _addvisorycoll[0].Contains("hil_mobilenumber") ? _addvisorycoll[0].GetAttributeValue<string>("hil_mobilenumber") : "";
                        homeadvisory.EmailID = _addvisorycoll[0].Contains("hil_emailid") ? _addvisorycoll[0].GetAttributeValue<string>("hil_emailid") : "";
                        homeadvisory.City = _addvisorycoll[0].Contains("hil_city") ? _addvisorycoll[0].GetAttributeValue<EntityReference>("hil_city").Name : "";
                        homeadvisory.PinCode = _addvisorycoll[0].Contains("hil_pincode") ? _addvisorycoll[0].GetAttributeValue<EntityReference>("hil_pincode").Name : "";
                        List<GetAdvisoryEnquiryLine> enqlineList = getEnqueryLine(service, _addvisorycoll[0].Id);
                        homeadvisory.AdvisoryEnquiryLine = enqlineList;
                        homeadvisoryenqList.Add(homeadvisory);
                        _retResponse.StatusCode = "200";
                        _retResponse.StatusMessage = "Record Retrieved";
                        _retResponse.Success = true;
                        _retResponse.Enquiry = homeadvisoryenqList;
                        return _retResponse;
                    }
                    else
                    {
                        _retResponse.StatusCode = "400";
                        _retResponse.Success = false;
                        _retResponse.StatusMessage = "Enquiry Not found.";
                        return _retResponse;
                    }

                }
                else if (req.MobileNo != null)
                {
                    #region fetch customer data...
                    QueryExpression query = new QueryExpression("contact");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, req.MobileNo.Trim());
                    EntityCollection customercoll = service.RetrieveMultiple(query);

                    if (customercoll.Entities.Count > 0)
                    {
                        string enquiryfetchxml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                  <entity name='hil_advisoryenquiry'>
                                                    <attribute name='hil_advisoryenquiryid' />
                                                    <attribute name='hil_name' />
                                                    <attribute name='createdon' />
                                                    <attribute name='hil_typeofcustomer' />
                                                    <attribute name='hil_tds' />
                                                    <attribute name='hil_state' />
                                                    <attribute name='hil_rooftop' />
                                                    <attribute name='hil_propertytype' />
                                                    <attribute name='hil_pincode' />
                                                    <attribute name='hil_mobilenumber' />
                                                    <attribute name='hil_emailid' />
                                                    <attribute name='hil_customer' />
                                                    <attribute name='hil_constructiontype' />
                                                    <attribute name='hil_city' />
                                                    <attribute name='hil_assettype' />
                                                    <attribute name='hil_areasqrt' />
                                                    <order attribute='hil_name' descending='false' />
                                                    <filter type='and'>
                                                      <condition attribute='statecode' operator='eq' value='0' />
                                                      <condition attribute='hil_customer' operator='eq' uiname='' uitype='contact' value='" + customercoll.Entities[0].Id + @"' />";
                        if (_datefilter != string.Empty)
                        {
                            enquiryfetchxml += _datefilter;
                        }
                        enquiryfetchxml += @"</filter>
                                    </entity>
                                </fetch>";
                        EntityCollection _enqlinecoll = service.RetrieveMultiple(new FetchExpression(enquiryfetchxml));
                        if (_enqlinecoll.Entities.Count > 0)
                        {
                            List<GetHomeAdvisoryEnquiry> homeadvisoryenqList = new List<GetHomeAdvisoryEnquiry>();
                            foreach (Entity adv in _enqlinecoll.Entities)
                            {
                                GetHomeAdvisoryEnquiry homeadvisory = new GetHomeAdvisoryEnquiry();
                                if (adv.Contains("hil_advisoryenquiryid"))
                                    homeadvisory.AdvisoryGuid = adv.Id.ToString();
                                homeadvisory.AdvisoryID = adv.Contains("hil_name") ? adv.GetAttributeValue<string>("hil_name") : "";
                                homeadvisory.Area = adv.Contains("hil_areasqrt") ? adv.GetAttributeValue<string>("hil_areasqrt") : "";
                                homeadvisory.CustomerType = adv.Contains("hil_typeofcustomer") ? adv.GetAttributeValue<EntityReference>("hil_typeofcustomer").Name : "";
                                homeadvisory.TDS = adv.Contains("hil_tds") ? adv.GetAttributeValue<string>("hil_tds") : "";
                                homeadvisory.State = adv.Contains("hil_state") ? adv.GetAttributeValue<EntityReference>("hil_state").Name : "";
                                homeadvisory.RoofTop = adv.GetAttributeValue<bool>("hil_rooftop") ? "Yes" : "No";
                                homeadvisory.PropertyType = adv.Contains("hil_propertytype") ? adv.GetAttributeValue<EntityReference>("hil_propertytype").Name : "";
                                homeadvisory.ConsturctionType = adv.Contains("hil_constructiontype") ? adv.GetAttributeValue<EntityReference>("hil_constructiontype").Name : "";
                                homeadvisory.AssetType = adv.Contains("hil_assettype") ? adv.GetAttributeValue<string>("hil_assettype") : "";
                                homeadvisory.Customer = adv.Contains("hil_customer") ? adv.GetAttributeValue<EntityReference>("hil_customer").Name : "";
                                homeadvisory.MobileNumber = adv.Contains("hil_mobilenumber") ? adv.GetAttributeValue<string>("hil_mobilenumber") : "";
                                homeadvisory.EmailID = adv.Contains("hil_emailid") ? adv.GetAttributeValue<string>("hil_emailid") : "";
                                homeadvisory.City = adv.Contains("hil_city") ? adv.GetAttributeValue<EntityReference>("hil_city").Name : "";
                                homeadvisory.PinCode = adv.Contains("hil_pincode") ? adv.GetAttributeValue<EntityReference>("hil_pincode").Name : "";
                                List<GetAdvisoryEnquiryLine> enqlineList = getEnqueryLine(service, adv.Id);
                                homeadvisory.AdvisoryEnquiryLine = enqlineList;
                                homeadvisoryenqList.Add(homeadvisory);
                            }
                            _retResponse.StatusCode = "200";
                            _retResponse.StatusMessage = "Record Retrieved";
                            _retResponse.Success = true;
                            _retResponse.Enquiry = homeadvisoryenqList;
                            return _retResponse;
                        }
                        else
                        {
                            _retResponse.StatusCode = "400";
                            _retResponse.Success = true;
                            _retResponse.StatusMessage = "Record Not Found";
                            return _retResponse;
                        }
                    }
                    else
                    {
                        _retResponse.StatusCode = "400";
                        _retResponse.StatusMessage = "Contact Not Found";
                        _retResponse.Success = true;
                        return _retResponse;
                    }
                    #endregion
                }
            }
            catch (Exception ex)
            {
                _retResponse.StatusCode = "505";
                _retResponse.Success = false;
                _retResponse.StatusMessage = "Dynamics 365 Internal Server Error : " + ex.Message;
                return _retResponse;
            }
            return _retResponse;
        }
        //functions
        bool closeAppointment(string _enqueryLine, string AppoitmentStatus, string Remarks, IOrganizationService service)
        {
            try
            {
                string fetchappointment = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='appointment'>
                                            <attribute name='subject' />
                                            <attribute name='statecode' />
                                            <attribute name='scheduledstart' />
                                            <attribute name='scheduledend' />
                                            <attribute name='createdby' />
                                            <attribute name='regardingobjectid' />
                                            <attribute name='activityid' />
                                            <attribute name='instancetypecode' />
                                            <order attribute='subject' descending='false' />
                                            <filter type='and'>
                                              <condition attribute='statecode' operator='in'>
                                                <value>0</value>
                                                <value>3</value>
                                              </condition>
                                            </filter>
                                            <link-entity name='hil_homeadvisoryline' from='hil_homeadvisorylineid' to='regardingobjectid' link-type='inner' alias='ab'>
                                              <filter type='and'>
                                                <condition attribute='hil_homeadvisorylineid' operator='eq' uiname='' uitype='hil_homeadvisoryline' value='" + _enqueryLine + @"' />
                                              </filter>
                                            </link-entity>
                                          </entity>
                                        </fetch>";
                EntityCollection AdvlineAppointmentColl = service.RetrieveMultiple(new FetchExpression(fetchappointment));
                if (AdvlineAppointmentColl.Entities.Count > 0)
                {

                    Entity _app = new Entity("appointment");
                    _app.Id = AdvlineAppointmentColl.Entities[0].Id;
                    _app["statecode"] = AppoitmentStatus == "3" ? new OptionSetValue(1) : new OptionSetValue(2);
                    _app["statuscode"] = AppoitmentStatus == "3" ? new OptionSetValue(3) : new OptionSetValue(4);
                    _app["description"] = Remarks != null ? Remarks : "";
                    service.Update(_app);
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch
            {
                return false;
            }
        }
        Response CancleTeamAppointment(Guid req, string AppoitmentStatus, string Remarks, IOrganizationService service)
        {
            Response resp = new Response();

            try
            {
                CancelEvent obj = new CancelEvent();
                
                //IOrganizationService service = ConnectToCRM.GetOrgService();
                QueryExpression Query = new QueryExpression("hil_homeadvisoryline");
                Query.ColumnSet = new ColumnSet("hil_appointmentid");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_homeadvisorylineid", ConditionOperator.Equal, req);
                EntityCollection _entitys = service.RetrieveMultiple(Query);
                if (_entitys.Entities.Count == 1)
                {
                    if (_entitys.Entities[0].Contains("hil_appointmentid"))
                    {
                        string trancId = _entitys.Entities[0].GetAttributeValue<String>("hil_appointmentid");

                        Integration integration = IntegrationConfiguration(service, "CancelEvent");
                        string _authInfo = integration.Auth;
                        _authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));
                        String sUrl = integration.uri;

                        var client = new RestClient(sUrl);
                        client.Timeout = -1;
                        var request = new RestRequest(Method.POST);
                        request.AddHeader("Authorization", _authInfo);
                        request.AddHeader("Content-Type", "application/json");
                        request.AddParameter("application/json", "{\r\n    \"TransactionId\": \"" + trancId + "\"\r\n}", ParameterType.RequestBody);
                        IRestResponse response = client.Execute(request);
                        Console.WriteLine(response.Content);
                        obj = Newtonsoft.Json.JsonConvert.DeserializeObject<CancelEvent>(response.Content);
                        if (obj.IsSuccess)
                        {
                            Entity _advLine = new Entity("hil_homeadvisoryline");
                            _advLine.Id = _entitys.Entities[0].Id;
                            _advLine["hil_appointmentid"] = String.Empty;
                            _advLine["hil_videocallurl"] = String.Empty;
                            _advLine["hil_appointmentstatus"] = new OptionSetValue(4);

                            service.Update(_advLine);

                            String fetct = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">" +
                                              "<entity name=\"appointment\">" +
                                                "<attribute name=\"subject\" />" +
                                                "<attribute name=\"statecode\" />" +
                                                "<attribute name=\"scheduledstart\" />" +
                                                "<attribute name=\"scheduledend\" />" +
                                                "<attribute name=\"createdby\" />" +
                                                "<attribute name=\"regardingobjectid\" />" +
                                                "<attribute name=\"activityid\" />" +
                                                "<attribute name=\"instancetypecode\" />" +
                                                "<order attribute=\"subject\" descending=\"false\" />" +
                                                "<filter type=\"and\">" +
                                                  "<condition attribute=\"regardingobjectid\" operator=\"eq\" uiname=\"\" uitype=\"hil_homeadvisoryline\" value=\"" + _entitys.Entities[0].Id + "\" />" +
                                                  "<condition attribute=\"hil_appointmenturl\" operator=\"not-null\" />" +
                                                  "<condition attribute=\"statecode\" operator=\"in\">" +
                                                    "<value>0</value>" +
                                                     "<value>3</value>" +
                                                    "</condition>" +
                                                  "</filter>" +
                                                "</entity>" +
                                              "</fetch>";
                            EntityCollection _apps = service.RetrieveMultiple(new FetchExpression(fetct));
                            var i = _apps.Entities.Count;
                            Entity _app = new Entity(_apps.EntityName);
                            _app.Id = _apps.Entities[0].Id;
                            _app["statecode"] = new OptionSetValue(2);
                            _app["statuscode"] = new OptionSetValue(4);

                            service.Update(_app);
                            resp.Message = obj.Message;

                        }
                        else
                        {
                            resp.Message = ("Invalid Transaction Id");
                        }
                    }
                    else
                    {
                        resp.Message = ("Meeting Not found for this Advisory");
                    }

                }
                else
                {
                    resp.Message = ("Record Not Found");
                }
            }
            catch (Exception ex)
            {
                resp.Message = ex.Message;
            }
            return resp;
        }
        List<GetAdvisoryEnquiryLine> getEnqueryLine(IOrganizationService service, Guid advId)
        {
            List<GetAdvisoryEnquiryLine> enqlineList = new List<GetAdvisoryEnquiryLine>();
            try
            {
                string fetcheqline = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                              <entity name='hil_homeadvisoryline'>
                                                <attribute name='hil_homeadvisorylineid' />
                                                <attribute name='hil_name' />
                                                <attribute name='createdon' />
                                                <attribute name='hil_videocallurl' />
                                                <attribute name='hil_typeofproduct' />
                                                <attribute name='hil_typeofenquiiry' />
                                                <attribute name='hil_pincode' />
                                                <attribute name='hil_enquirystauts' />
                                                <attribute name='hil_customerremark' />
                                                <attribute name='hil_assignedadvisor' />
                                                <attribute name='hil_appointmentid' />
                                                <attribute name='hil_appointmenttypes' />
                                                <attribute name='hil_appointmentstatus' />
                                                <attribute name='hil_appointmentdate' />
                                                <attribute name='hil_advisoryenquery' />
                                                <order attribute='hil_name' descending='false' />
                                                <filter type='and'>
                                                  <condition attribute='statecode' operator='eq' value='0' />
                                                  <condition attribute='hil_advisoryenquery' operator='eq' uiname='advenq-00015' uitype='hil_advisoryenquiry' value='" + advId + @"' />
                                                </filter>
                                              </entity>
                                            </fetch>";
                EntityCollection _addvisorylinecoll = service.RetrieveMultiple(new FetchExpression(fetcheqline));

                if (_addvisorylinecoll.Entities.Count > 0)
                {
                    foreach (Entity advline in _addvisorylinecoll.Entities)
                    {
                        GetAdvisoryEnquiryLine enqline = new GetAdvisoryEnquiryLine();
                        enqline.EnquiryLineGuid = advline.Id.ToString();
                        enqline.EnquiryLineID = advline.Contains("hil_name") ? advline.GetAttributeValue<string>("hil_name") : "";
                        enqline.AssignAdvisor = advline.Contains("hil_advisoryenquery") ? advline.GetAttributeValue<EntityReference>("hil_advisoryenquery").Name : "";
                        enqline.AppointmentId = advline.Contains("hil_appointmentid") ? advline.GetAttributeValue<string>("hil_name") : "";
                        enqline.AppointmentDate = advline.Contains("hil_appointmentdate") ? advline.GetAttributeValue<DateTime>("hil_appointmentdate").AddMinutes(330).ToString() : "";
                        enqline.AdvisoryDate = advline.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString();
                        enqline.AppointmentStatus = advline.Contains("hil_appointmentstatus") ? advline.FormattedValues["hil_appointmentstatus"].ToString() : "";
                        enqline.EnquiryStatus = advline.Contains("hil_enquirystauts") ? advline.FormattedValues["hil_enquirystauts"].ToString() : "";
                        enqline.AppointmentType = advline.Contains("hil_appointmenttypes") ? advline.FormattedValues["hil_appointmenttypes"] : "";
                        enqline.TypeofEnquiry = advline.Contains("hil_typeofenquiiry") ? advline.GetAttributeValue<EntityReference>("hil_typeofenquiiry").Name : "";
                        enqline.CustomerReamark = advline.Contains("hil_customerremark") ? advline.GetAttributeValue<String>("hil_customerremark") : "";
                        enqline.PinCode = advline.Contains("hil_pincode") ? advline.GetAttributeValue<EntityReference>("hil_pincode").Name : "";
                        enqline.VideoCallUrl = advline.Contains("hil_videocallurl") ? advline.GetAttributeValue<string>("hil_videocallurl") : "";
                        enqline.TypeofProduct = advline.Contains("hil_typeofproduct") ? advline.GetAttributeValue<EntityReference>("hil_typeofproduct").Name : "";
                        enqline.Documents = getEnqueryLineDocuments(service, advline.Id);
                        enqlineList.Add(enqline);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return enqlineList;
        }
        List<GetDocument> getEnqueryLineDocuments(IOrganizationService service, Guid advEnqId)
        {
            List<GetDocument> _getDocument = new List<GetDocument>();
            try
            {
                string fetcheqline = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='hil_attachment'>
                                        <attribute name='activityid' />
                                        <attribute name='subject' />
                                        <attribute name='hil_docurl' />
                                        <attribute name='hil_doctype' />
                                        <attribute name='hil_docsize' />
                                        <order attribute='subject' descending='false' />
                                        <filter type='and'>
                                          <condition attribute='regardingobjectid' operator='eq' uiname='' uitype='hil_homeadvisoryline' value='" + advEnqId + @"' />
                                        </filter>
                                      </entity>
                                    </fetch>";
                EntityCollection _DocColl = service.RetrieveMultiple(new FetchExpression(fetcheqline));

                foreach (Entity docline in _DocColl.Entities)
                {
                    GetDocument _Doc = new GetDocument();
                    _Doc.Subject = docline.Contains("subject") ? docline.GetAttributeValue<string>("subject") : "";
                    _Doc.DocumentURL = docline.Contains("hil_docurl") ? docline.GetAttributeValue<string>("hil_docurl") : "";
                    _Doc.DocumentType = docline.Contains("hil_doctype") ? docline.FormattedValues["hil_doctype"] : "";
                    _Doc.DocumentSize = docline.Contains("hil_docsize") ? docline.GetAttributeValue<Double>("hil_docsize").ToString() : "";
                    _getDocument.Add(_Doc);
                }

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return _getDocument;
        }
        CreateUserMeetingResponse CreateMeeting(String AppointmentID, String SlotEnd, String SlotStart, String EnquirerEmailId, String EnquirerName, String SlotDate, String UserCode, bool IsVideo, string EnquiryID, string EnquiryType, string AdvisoryType)
        {
            CreateUserMeetingResponse obj = new CreateUserMeetingResponse();
            try
            {
                Slot slot = new Slot();
                slot.SlotEnd = SlotEnd;
                slot.SlotStart = SlotStart;

                CreateUserMeetingRequest reqParm = new CreateUserMeetingRequest();
                reqParm.EnquirerEmailId = EnquirerEmailId;
                reqParm.EnquirerName = EnquirerName;
                reqParm.SlotDate = SlotDate;
                reqParm.UserCode = int.Parse(UserCode);
                reqParm.Slot = slot;
                reqParm.IsVideoMeeting = IsVideo;
                reqParm.TransactionId = AppointmentID;
                reqParm.EnquiryId = EnquiryID;
                reqParm.EnquiryType = EnquiryType;
                reqParm.AdvisoryType = AdvisoryType;

                IOrganizationService service = ConnectToCRM.GetOrgService();

                Integration integration = IntegrationConfiguration(service, "CreateUserMeeting");

                string _authInfo = integration.Auth;
                _authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));
                String sUrl = integration.uri;

                var client = new RestClient(sUrl);
                client.Timeout = -1;
                var request = new RestRequest(Method.POST);

                request.AddHeader("Authorization", _authInfo);
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter("application/json", Newtonsoft.Json.JsonConvert.SerializeObject(reqParm), ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                obj = Newtonsoft.Json.JsonConvert.DeserializeObject<CreateUserMeetingResponse>(response.Content);
            }
            catch {
                obj.Message = "Failed";
            }
            return obj;
        }
        public void Assign(EntityReference _Assignee, EntityReference _Targetd, IOrganizationService service)
        {
            try
            {
                AssignRequest assign = new AssignRequest();
                assign.Assignee = _Assignee;
                assign.Target = _Targetd;
                service.Execute(assign);
            }
            catch
            {
            }
        }
        public ResponseUpload UploadAttachment(UploadAttachment reqParm)
        {
            ResponseUpload resParm = new ResponseUpload();
            String blobURL = string.Empty;
            IOrganizationService service = ConnectToCRM.GetOrgServiceProd();
            if (reqParm.DocGuid != string.Empty && reqParm.DocGuid != null)
            {
                if (!reqParm.IsDeleted)
                {
                    resParm.Message = "Parameter Mismatch to Delete Existing Document";
                    resParm.Status = false;
                    return resParm;
                }
                else
                {
                    Entity ent = new Entity("hil_attachment", new Guid(reqParm.DocGuid));
                    ent["hil_isdeleted"] = true;
                    service.Update(ent);
                    resParm.Message = "Success";
                    resParm.Status = true;
                    return resParm;
                }
            }
            else
            {
                if (reqParm.FileName == null || reqParm.FileName == "" ||
                reqParm.FileSize == null || reqParm.FileSize == "" ||
                reqParm.FileString == null || reqParm.FileString == "" ||
                reqParm.DocumentType == null || reqParm.DocumentType == "")
                {
                    resParm.Message = "Some Required feild is missing";
                    resParm.Status = false;
                    return resParm;
                }
                else
                {
                    try
                    {
                        EntityReference _erDepartment = null;
                        QueryExpression query = null;
                        EntityCollection entCol = null;

                        EntityReference _doctypee = new EntityReference(); //getOptionSetValue(reqParm.DocumentType, "hil_attachment", "hil_doctype", service);
                        Guid loginUserGuid = ((WhoAmIResponse)service.Execute(new WhoAmIRequest())).UserId;
                        query = new QueryExpression("hil_attachmentdocumenttype");
                        query.ColumnSet = new ColumnSet("hil_containername", "hil_department");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, reqParm.DocumentType.Trim());

                        entCol = service.RetrieveMultiple(query);
                        String containerName = string.Empty;

                        if (entCol.Entities.Count == 0)
                        {
                            resParm.Message = "Given Document Type is not Defined in Dynamics.";
                            resParm.Status = false;
                            return resParm;
                        }
                        else
                        {
                            _doctypee = new EntityReference(entCol[0].LogicalName, entCol[0].Id);
                            containerName = entCol[0].Contains("hil_containername") ? entCol[0].GetAttributeValue<string>("hil_containername") : "devanduat";
                            if (entCol.Entities[0].Contains("hil_department"))
                                _erDepartment = entCol.Entities[0].GetAttributeValue<EntityReference>("hil_department");
                        }
                        string[] ext = reqParm.FileName.Split('.');
                        String fileName = string.Empty;

                        if (reqParm.RegardingGuid != null && reqParm.RegardingGuid != string.Empty)
                        {
                            if (reqParm.ObjectType == "hil_tender")
                            {
                                fileName = service.Retrieve(reqParm.ObjectType, new Guid(reqParm.RegardingGuid), new ColumnSet("hil_name")).GetAttributeValue<string>("hil_name");
                            }
                            else
                            {
                                fileName = reqParm.RegardingGuid;
                            }
                            fileName = fileName + "_" + reqParm.DocumentType + "_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + "." + ext[ext.Length - 1];
                        }
                        else
                        {
                            fileName = reqParm.DocumentType + "_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + "." + ext[ext.Length - 1];
                        }
                        //String fileName = reqParm.RegardingGuid + "_" + reqParm.DocumentType + "_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + "." + ext[ext.Length - 1];
                        byte[] fileContent = Convert.FromBase64String(reqParm.FileString);
                        try
                        {
                            blobURL = Upload(fileName, fileContent, containerName);  // reqParm.ContainerName); ;
                        }
                        catch (Exception ex)
                        {
                            resParm.Message = ex.Message;
                            resParm.Status = false;
                            return resParm;
                        }

                        Entity _attachment = new Entity("hil_attachment");
                        if (reqParm.Subject == string.Empty || reqParm.Subject.Trim().Length == 0)
                        {
                            int _rowCount = 0;
                            int pageNumber = 1;
                            query = new QueryExpression("hil_attachment");
                            query.ColumnSet = new ColumnSet(false);
                            query.PageInfo = new PagingInfo();
                            query.PageInfo.Count = 5000;
                            query.PageInfo.PageNumber = pageNumber;
                            query.PageInfo.PagingCookie = null;
                            while (true)
                            {
                                EntityCollection entColFile = service.RetrieveMultiple(query);
                                _rowCount += entColFile.Entities.Count;
                                if (entColFile.MoreRecords)
                                {
                                    // Increment the page number to retrieve the next page.
                                    query.PageInfo.PageNumber++;
                                    // Set the paging cookie to the paging cookie returned from current results.
                                    query.PageInfo.PagingCookie = entColFile.PagingCookie;
                                }
                                else
                                {
                                    // If no more records are in the result nodes, exit the loop.
                                    break;
                                }
                            }
                            _rowCount += 1;
                            _attachment["subject"] = "HIL_" + _rowCount.ToString().PadLeft(8, '0');
                            _attachment["description"] = reqParm.Description;
                        }
                        else
                        {
                            _attachment["subject"] = reqParm.Subject;
                            _attachment["description"] = reqParm.Description;
                        }

                        if (_erDepartment != null) { _attachment["hil_department"] = _erDepartment; }

                        _attachment["hil_documenttype"] = _doctypee;
                        _attachment["hil_docurl"] = blobURL;
                        resParm.BlobURL = blobURL;
                        _attachment["hil_docsize"] = double.Parse(reqParm.FileSize);
                        if (reqParm.ValidFrom != null && reqParm.ValidFrom != "")
                        {
                            DateTime fromDate = new DateTime(Convert.ToInt32(reqParm.ValidFrom.Substring(0, 4)),
                                Convert.ToInt32(reqParm.ValidFrom.Substring(4, 2)),
                                Convert.ToInt32(reqParm.ValidFrom.Substring(6, 2)));

                            _attachment["scheduledstart"] = fromDate.Date;
                        }
                        if (reqParm.ValidTill != null && reqParm.ValidTill != "")
                        {
                            DateTime toDate = new DateTime(Convert.ToInt32(reqParm.ValidTill.Substring(0, 4)),
                               Convert.ToInt32(reqParm.ValidTill.Substring(4, 2)),
                               Convert.ToInt32(reqParm.ValidTill.Substring(6, 2)));
                            _attachment["scheduledend"] = toDate.Date;
                        }

                        if (reqParm.RegardingGuid != null && reqParm.RegardingGuid != string.Empty)
                        {

                            if (reqParm.ObjectType != "systemuser")
                                _attachment["regardingobjectid"] = new EntityReference(reqParm.ObjectType, new Guid(reqParm.RegardingGuid));
                            else
                                _attachment["hil_regardinguser"] = new EntityReference(reqParm.ObjectType, new Guid(reqParm.RegardingGuid));
                        }

                        _attachment["hil_sourceofdocument"] = new OptionSetValue(int.Parse(reqParm.Source));
                        try
                        {
                            service.Create(_attachment);
                            if (reqParm.RegardingGuid != null && reqParm.RegardingGuid != string.Empty)
                            {
                                try
                                {
                                    Entity _entityReg = service.Retrieve(reqParm.ObjectType, new Guid(reqParm.RegardingGuid), new ColumnSet("ownerid"));
                                    if (_entityReg.Contains("ownerid"))
                                    {
                                        EntityReference owner = _entityReg.GetAttributeValue<EntityReference>("ownerid");
                                        if (loginUserGuid != owner.Id)
                                        {
                                            Assign(owner, new EntityReference(reqParm.ObjectType, new Guid(reqParm.RegardingGuid)), service);
                                        }
                                    }
                                }
                                catch { }
                            }
                        }
                        catch (Exception ex)
                        {
                            resParm.Message = "Failed to Create Record : " + ex.Message;
                            resParm.Status = false;
                            return resParm;
                        }
                        resParm.Message = "File Uplaoded Sucessfully";
                        resParm.Status = true;
                        return resParm;
                    }
                    catch (Exception ex)
                    {
                        resParm.Message = "Failed to Create Record : " + ex.Message;
                        resParm.Status = false;
                        return resParm;
                    }
                }
            }
        }
        //public ResponseUpload UploadAttachment(UploadAttachment reqParm)
        //{
        //    ResponseUpload resParm = new ResponseUpload();
        //    String blobURL = string.Empty;
        //    IOrganizationService service = ConnectToCRM.GetOrgService();
        //    if (reqParm.DocGuid != string.Empty && reqParm.DocGuid != null)
        //    {
        //        if (!reqParm.IsDeleted)
        //        {
        //            resParm.Message = "Parameter Mismatch to Delete Existing Document";
        //            resParm.Status = false;
        //            return resParm;
        //        }
        //        else
        //        {
        //            Entity ent = new Entity("hil_attachment", new Guid(reqParm.DocGuid));
        //            ent["hil_isdeleted"] = true;
        //            service.Update(ent);
        //            resParm.Message = "Success";
        //            resParm.Status = true;
        //            return resParm;
        //        }
        //    }
        //    else
        //    {
        //        if (reqParm.Subject == null || reqParm.Subject == "" ||
        //            reqParm.FileName == null || reqParm.FileName == "" ||
        //            reqParm.FileSize == null || reqParm.FileSize == "" ||
        //            reqParm.FileString == null || reqParm.FileString == "" ||
        //            reqParm.DocumentType == null || reqParm.DocumentType == "")
        //        {
        //            resParm.Message = "Some Required feild is missing";
        //            resParm.Status = false;
        //            return resParm;
        //        }
        //        else
        //        {
        //            try
        //            {
        //                EntityReference _doctypee = new EntityReference(); //getOptionSetValue(reqParm.DocumentType, "hil_attachment", "hil_doctype", service);
        //                Guid loginUserGuid = ((WhoAmIResponse)service.Execute(new WhoAmIRequest())).UserId;
        //                QueryExpression query = new QueryExpression("hil_attachmentdocumenttype");
        //                query.ColumnSet = new ColumnSet("hil_containername");
        //                query.Criteria = new FilterExpression(LogicalOperator.And);
        //                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
        //                query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, reqParm.DocumentType.Trim());
        //                EntityCollection entCol = service.RetrieveMultiple(query);
        //                String containerName = string.Empty;

        //                if (entCol.Entities.Count == 0)
        //                {
        //                    resParm.Message = "Given Document Type is not Defined in Dynamics.";
        //                    resParm.Status = false;
        //                    return resParm;
        //                }
        //                else
        //                {
        //                    _doctypee = new EntityReference(entCol[0].LogicalName, entCol[0].Id);
        //                    containerName = entCol[0].Contains("hil_containername") ? entCol[0].GetAttributeValue<string>("hil_containername") : "devanduat";
        //                }
        //                string[] ext = reqParm.FileName.Split('.');
        //                String fileName = string.Empty;

        //                if (reqParm.RegardingGuid != null && reqParm.RegardingGuid != string.Empty)
        //                {
        //                    if (reqParm.ObjectType == "hil_tender")
        //                    {
        //                        fileName = service.Retrieve(reqParm.ObjectType, new Guid(reqParm.RegardingGuid), new ColumnSet("hil_name")).GetAttributeValue<string>("hil_name");
        //                    }
        //                    else {
        //                        fileName = reqParm.RegardingGuid;
        //                    }
        //                    fileName = fileName + "_" + reqParm.DocumentType + "_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + "." + ext[ext.Length - 1];
        //                }
        //                else
        //                {
        //                    fileName = reqParm.DocumentType + "_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + "." + ext[ext.Length - 1];
        //                }
        //                byte[] fileContent = Convert.FromBase64String(reqParm.FileString);
        //                try
        //                {
        //                    blobURL = Upload(fileName, fileContent, containerName);  // reqParm.ContainerName); ;
        //                }
        //                catch (Exception ex)
        //                {
        //                    resParm.Message = ex.Message;
        //                    resParm.Status = false;
        //                    return resParm;
        //                }

        //                Entity _attachment = new Entity("hil_attachment");
        //                _attachment["subject"] = reqParm.Subject;

        //                _attachment["hil_documenttype"] = _doctypee;
        //                _attachment["hil_docurl"] = blobURL;
        //                resParm.BlobURL = blobURL;
        //                _attachment["hil_docsize"] = double.Parse(reqParm.FileSize);
        //                if (reqParm.ValidFrom != null && reqParm.ValidFrom != "")
        //                {
        //                    DateTime fromDate = new DateTime(Convert.ToInt32(reqParm.ValidFrom.Substring(0, 4)),
        //                        Convert.ToInt32(reqParm.ValidFrom.Substring(4, 2)),
        //                        Convert.ToInt32(reqParm.ValidFrom.Substring(6, 2)));

        //                    _attachment["scheduledstart"] = fromDate.Date;
        //                }
        //                if (reqParm.ValidTill != null && reqParm.ValidTill != "")
        //                {
        //                    DateTime toDate = new DateTime(Convert.ToInt32(reqParm.ValidTill.Substring(0, 4)),
        //                       Convert.ToInt32(reqParm.ValidTill.Substring(4, 2)),
        //                       Convert.ToInt32(reqParm.ValidTill.Substring(6, 2)));
        //                    _attachment["scheduledend"] = toDate.Date;
        //                }

        //                if (reqParm.RegardingGuid != null && reqParm.RegardingGuid != string.Empty)
        //                {

        //                    if (reqParm.ObjectType != "systemuser")
        //                        _attachment["regardingobjectid"] = new EntityReference(reqParm.ObjectType, new Guid(reqParm.RegardingGuid));
        //                    else
        //                        _attachment["hil_regardinguser"] = new EntityReference(reqParm.ObjectType, new Guid(reqParm.RegardingGuid));
        //                }

        //                _attachment["hil_sourceofdocument"] = new OptionSetValue(int.Parse(reqParm.Source));
        //                try
        //                {
        //                    resParm.DocGuid = service.Create(_attachment).ToString();
        //                    if (reqParm.RegardingGuid != null && reqParm.RegardingGuid != string.Empty)
        //                    {
        //                        try
        //                        {
        //                            Entity _entityReg = service.Retrieve(reqParm.ObjectType, new Guid(reqParm.RegardingGuid), new ColumnSet("ownerid"));
        //                            if (_entityReg.Contains("ownerid"))
        //                            {
        //                                EntityReference owner = _entityReg.GetAttributeValue<EntityReference>("ownerid");
        //                                if (loginUserGuid != owner.Id)
        //                                {
        //                                    Assign(owner, new EntityReference(reqParm.ObjectType, new Guid(reqParm.RegardingGuid)), service);
        //                                }
        //                            }
        //                        }
        //                        catch { }
        //                    }
        //                }
        //                catch (Exception ex)
        //                {
        //                    resParm.Message = "Failed to Create Record : " + ex.Message;
        //                    resParm.Status = false;
        //                    return resParm;
        //                }
        //                resParm.Message = "File Uplaoded Sucessfully";
        //                resParm.Status = true;
        //                return resParm;
        //            }
        //            catch (Exception ex)
        //            {
        //                resParm.Message = "Failed to Create Record : " + ex.Message;
        //                resParm.Status = false;
        //                return resParm;
        //            }
        //        }
        //    }
        //}
        string ToURLSlug(string s)
        {
            return Regex.Replace(s, @"[^a-z0-9]+", "-", RegexOptions.IgnoreCase)
                .Trim(new char[] { '-' })
                .ToLower();
        }
        string Upload(string fileName, byte[] fileContent, string containerName)
        {
            string _blobURI = string.Empty;
            try
            {
                fileName = Regex.Replace(fileName, @"\s+", String.Empty);
                //byte[] fileContent = Convert.FromBase64String(noteBody);
                string ConnectionSting = "DefaultEndpointsProtocol=https;AccountName=d365storagesa;AccountKey=6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==;EndpointSuffix=core.windows.net";
                // create object of storage account
                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(ConnectionSting);

                // create client of storage account
                CloudBlobClient client = storageAccount.CreateCloudBlobClient();

                // create the reference of your storage account
                CloudBlobContainer container = client.GetContainerReference(ToURLSlug(containerName));

                // check if the container exists or not in your account
                var isCreated = container.CreateIfNotExists();

                // set the permission to blob type
                container.SetPermissionsAsync(new BlobContainerPermissions
                { PublicAccess = BlobContainerPublicAccessType.Blob });

                // create the memory steam which will be uploaded
                using (MemoryStream memoryStream = new MemoryStream(fileContent))
                {
                    // set the memory stream position to starting
                    memoryStream.Position = 0;

                    // create the object of blob which will be created
                    // Test-log.txt is the name of the blob, pass your desired name
                    CloudBlockBlob blob = container.GetBlockBlobReference(fileName);

                    // get the mime type of the file
                    string mimeType = "application/unknown";
                    string ext = (fileName.Contains(".")) ?
                                System.IO.Path.GetExtension(fileName).ToLower() : "." + fileName;
                    Microsoft.Win32.RegistryKey regKey =
                                Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(ext);
                    if (regKey != null && regKey.GetValue("Content Type") != null)
                        mimeType = regKey.GetValue("Content Type").ToString();

                    // set the memory stream position to zero again
                    // this is one of the important stage, If you miss this step, 
                    // your blob will be created but size will be 0 byte
                    memoryStream.ToArray();
                    memoryStream.Seek(0, SeekOrigin.Begin);

                    // set the mime type
                    blob.Properties.ContentType = mimeType;

                    // upload the stream to blob
                    blob.UploadFromStream(memoryStream);
                    _blobURI = blob.Uri.AbsoluteUri;
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Failed to Upload File: " + ex.Message);
            }
            return _blobURI;
        }
        static int getOptionSetValue(string lable, string _entityLogicalName, string entAttr_LogicalName, IOrganizationService service)
        {
            int _value = 100;
            var retrieveAttributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = _entityLogicalName,
                LogicalName = entAttr_LogicalName,
                RetrieveAsIfPublished = true
            };

            var attributeResponse = (RetrieveAttributeResponse)service.Execute(retrieveAttributeRequest);
            var attributeMetadata = (EnumAttributeMetadata)attributeResponse.AttributeMetadata;
            var optionList = (from o in attributeMetadata.OptionSet.Options select new { Value = o.Value, Text = o.Label.UserLocalizedLabel.Label }).ToList();

            int? value = optionList.Where(o => o.Text.ToLower() == lable.ToLower())
                                        .Select(o => o.Value)
                                        .FirstOrDefault();
            if (value != null)
            {
                _value = int.Parse(value.ToString());
            }
            return _value;
        }
        static Integration IntegrationConfiguration(IOrganizationService service, string Param)
        {
            Integration output = new Integration();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + ":" + integrationConfiguration.GetAttributeValue<string>("hil_password");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:- " + ex.Message);
            }
            return output;
        }
        public GetUserTimeSlotsRoot GetUserTimeSlots(GetUserTimeSlotsRequest requestParm)
        {
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();

                Integration intConFig = IntegrationConfiguration(service, "GetUserTimeSlots");
                string _authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConFig.Auth));
                var client = new RestClient(intConFig.uri);
                client.Timeout = -1;
                var request = new RestRequest(Method.POST);
                request.AddHeader("Authorization", _authInfo);
                request.AddHeader("Content-Type", "application/json");
                String body = Newtonsoft.Json.JsonConvert.SerializeObject(requestParm);
                request.AddParameter("application/json", body, ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                Console.WriteLine(response.Content);
                GetUserTimeSlotsRoot rootObject = Newtonsoft.Json.JsonConvert.DeserializeObject<GetUserTimeSlotsRoot>(response.Content);
                return rootObject;
            }
            catch (Exception ex)
            {
                return new GetUserTimeSlotsRoot() { IsSuccess = false, Message = ex.Message, Data = null };
            }
        }
    }
    public class Integration
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }
    [DataContract]
    public class ResponseUpload
    {
        [DataMember]
        public string Message { get; set; }
        [DataMember]
        public bool Status { get; set; }
        [DataMember]
        public String BlobURL { get; set; }
        [DataMember]
        public string DocGuid { get; set; }
    }
    [DataContract]
    public class UploadAttachment
    {
        [DataMember]
        public string Subject { get; set; }
        [DataMember]
        public string Description { get; set; }
        [DataMember]
        public string FileSize { get; set; }
        [DataMember]
        public string FileName { get; set; }
        [DataMember]
        public string FileString { get; set; }
        [DataMember]
        public string ObjectType { get; set; }
        [DataMember]
        public string RegardingGuid { get; set; }
        [DataMember]
        public string Source { get; set; }
        [DataMember]
        public string DocumentType { get; set; }
        [DataMember]
        public string ValidFrom { get; set; }
        [DataMember]
        public string ValidTill { get; set; }
        [DataMember]
        public bool IsDeleted { get; set; }
        [DataMember]
        public string DocGuid { get; set; }
        [DataMember]
        public string Department { get; set; }
    }
    [DataContract]
    public class Enquiry
    {
        [DataMember]
        public string AdvisoryID { get; set; }
        [DataMember]
        public string MobileNo { get; set; }

    }

    [DataContract]
    public class EnquiryStatus
    {
        [DataMember]
        public string EnquiryId { get; set; }
        [DataMember]
        public string MobileNo { get; set; }
        [DataMember]
        public string FromDate { get; set; }
        [DataMember]
        public string ToDate { get; set; }
    }
    [DataContract]
    public class GetEnquiry
    {
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusMessage { get; set; }
        [DataMember]
        public bool Success { get; set; }
        [DataMember]
        public List<GetHomeAdvisoryEnquiry> Enquiry { get; set; }

    }
    [DataContract]
    public class GetHomeAdvisoryEnquiry
    {

        [DataMember]
        public string AdvisoryID { get; set; }
        [DataMember]
        public string AdvisoryGuid { get; set; }
        [DataMember]
        public string CustomerType { get; set; }
        [DataMember]
        public string Area { get; set; }
        [DataMember]
        public string PropertyType { get; set; }
        [DataMember]
        public string ConsturctionType { get; set; }
        [DataMember]
        public String RoofTop { get; set; }
        [DataMember]
        public string Customer { get; set; }
        [DataMember]
        public string AssetType { get; set; }
        [DataMember]
        public string MobileNumber { get; set; }
        [DataMember]
        public string EmailID { get; set; }
        [DataMember]
        public string City { get; set; }
        [DataMember]
        public string State { get; set; }
        [DataMember]
        public string PinCode { get; set; }
        [DataMember]
        public string TDS { get; set; }

        [DataMember]
        public string CustomerRemarks { get; set; }

        [DataMember]
        public List<GetAdvisoryEnquiryLine> AdvisoryEnquiryLine { get; set; }

    }
    [DataContract]
    public class GetAdvisoryEnquiryLine
    {

        [DataMember]
        public string EnquiryLineGuid { get; set; }
        [DataMember]
        public string EnquiryLineID { get; set; }
        [DataMember]
        public string TypeofEnquiry { get; set; }
        [DataMember]
        public string EnquiryStatus { get; set; }
        [DataMember]
        public string TypeofProduct { get; set; }
        [DataMember]
        public string AppointmentId { get; set; }
        [DataMember]
        public String AppointmentType { get; set; }
        [DataMember]
        public string AppointmentDate { get; set; }
        [DataMember]
        public string AppointmentStatus { get; set; }
        [DataMember]
        public string CustomerReamark { get; set; }
        [DataMember]
        public string AssignAdvisor { get; set; }
        [DataMember]
        public string PinCode { get; set; }
        [DataMember]
        public string VideoCallUrl { get; set; }

        [DataMember]
        public string AdvisoryDate { get; set; }
        [DataMember]
        public List<GetDocument> Documents { get; set; }
    }
    [DataContract]
    public class GetDocument
    {
        [DataMember]
        public string Subject { get; set; }
        [DataMember]
        public string DocumentURL { get; set; }
        [DataMember]
        public string DocumentType { get; set; }
        [DataMember]
        public string DocumentSize { get; set; }
    }
    //[DataContract]
    //public class UploadAttachment
    //{
    //    [DataMember]
    //    public string Subject { get; set; }
    //    [DataMember]
    //    public string FileSize { get; set; }
    //    [DataMember]
    //    public string FileName { get; set; }
    //    [DataMember]
    //    public string FileString { get; set; }
    //    [DataMember]
    //    public string ObjectType { get; set; }
    //    [DataMember]
    //    public string RegardingGuid { get; set; }
    //    [DataMember]
    //    public string Source { get; set; }
    //    [DataMember]
    //    public string DocumentType { get; set; }
    //    [DataMember]
    //    public bool IsDeleted { get; set; }
    //    [DataMember]
    //    public string DocGuid { get; set; }
    //    [DataMember]
    //    public string ValidFrom { get; set; }
    //    [DataMember]
    //    public string ValidTill { get; set; }
    //}
    [DataContract]
    public class ReschduleAppointment
    {
        [DataMember]
        public string AdvisorylineId { get; set; }
        [DataMember]
        public string scheduledstart { get; set; }
        [DataMember]
        public string scheduledEnd { get; set; }
        [DataMember]
        public string appointmenturl { get; set; }
        [DataMember]
        public string appintmentId { get; set; }
        [DataMember]
        public string appintmentType { get; set; }
        [DataMember]
        public string Remarks { get; set; }
        [DataMember]
        public string AssignedUsercode { get; set; }
    }
    [DataContract]
    public class CancelAppointmentRequest
    {
        [DataMember]
        public string AdvisorylineGuid { get; set; }
        [DataMember]
        public string AdvisorylineId { get; set; }
        [DataMember]
        public string AppointmentRemarks { get; set; }
        [DataMember]
        public string AppointmentStatus { get; set; }
        [DataMember]
        public bool IsEnquiryClosed { get; set; }
        [DataMember]
        public string EnquiryRemarks { get; set; }
        [DataMember]
        public string EnquiryStatus { get; set; }
        [DataMember]
        public string EnquiryCloseReason { get; set; }

    }
    [DataContract]
    public class AssignEnqeryLine
    {
        [DataMember]
        public string EnquiryId { get; set; }
        [DataMember]
        public string AdvisorylineId { get; set; }
    }
    [DataContract]
    public class HopmeAdvisoryResult
    {
        [DataMember]
        public string statusCode { get; set; }
        [DataMember]
        public string statusDiscription { get; set; }
        [DataMember]
        public string message { get; set; }
        [DataMember]
        public string enquiryid { get; set; }

        [DataMember]
        public string enquiryGuId { get; set; }
    }
    [DataContract]
    public class HomeAdvisoryRequest
    {
        [DataMember]
        public string AdvisoryGuid { get; set; }
        [DataMember]
        public string AdvisoryID { get; set; }
        [DataMember]
        public string CustomerType { get; set; }
        [DataMember]
        public string Area { get; set; }
        [DataMember]
        public string PropertyType { get; set; }
        [DataMember]
        public string ConsturctionType { get; set; }
        [DataMember]
        public bool? rooftop { get; set; }
        [DataMember]
        public string Salutation { get; set; }
        [DataMember]
        public string CustomerName { get; set; }

        [DataMember]
        public string assettype { get; set; }
        [DataMember]
        public string mobilenumber { get; set; }
        [DataMember]
        public string emailid { get; set; }
        [DataMember]
        public string city { get; set; }
        [DataMember]
        public string state { get; set; }
        [DataMember]
        public string pincode { get; set; }
        [DataMember]
        public string tds { get; set; }

        [DataMember]
        public string CustomerRemarks { get; set; }

        [DataMember]
        public string SourceofCreation { get; set; }
        [DataMember]
        public List<AdvisoryEnquiryLine> advisoryenquryline { get; set; }
    }
    [DataContract]
    public class AdvisoryEnquiryLine
    {
        [DataMember]
        public string ProductType { get; set; }
        [DataMember]
        public string Enqurystatus { get; set; }
        [DataMember]
        public string ApointmentType { get; set; }
        [DataMember]
        public string EnquryTypecode { get; set; }
        [DataMember]
        public List<UploadAttachment> Attachments { get; set; }
    }
    [DataContract]
    public class CancelEvent
    {
        [DataMember]
        public bool IsSuccess { get; set; }
        [DataMember]
        public string Message { get; set; }
    }
    [DataContract]
    public class CancleAppointmentResuest
    {
        [DataMember]
        public string RecordID { get; set; }
    }
    [DataContract]
    public class Slot
    {
        [DataMember]
        public string SlotStart { get; set; }
        [DataMember]
        public string SlotEnd { get; set; }
    }
    [DataContract]
    public class Response
    {
        [DataMember]
        public string Message { get; set; }
        [DataMember]
        public bool Status { get; set; }
    }
    //[DataContract]
    //public class ResponseUpload
    //{
    //    [DataMember]
    //    public string Message { get; set; }
    //    [DataMember]
    //    public bool Status { get; set; }
    //    [DataMember]
    //    public String BlobURL { get; set; }

    //    [DataMember]
    //    public bool IsDeleted { get; set; }
    //    [DataMember]
    //    public string DocGuid { get; set; }
    //}
    [DataContract]
    public class CreateUserMeetingRequest
    {
        [DataMember]
        public int UserCode { get; set; }
        [DataMember]
        public string EnquirerEmailId { get; set; }
        [DataMember]
        public string EnquirerName { get; set; }
        [DataMember]
        public string SlotDate { get; set; }
        [DataMember]
        public Slot Slot { get; set; }
        [DataMember]
        public bool IsVideoMeeting { get; set; }
        [DataMember]
        public String TransactionId { get; set; }
        [DataMember]
        public String EnquiryId { get; set; }
        [DataMember]
        public String EnquiryType { get; set; }
        [DataMember]
        public String AdvisoryType { get; set; }
    }
    [DataContract]
    public class Data
    {
        [DataMember]
        public string TransactionId { get; set; }
        [DataMember]
        public string MeetingURL { get; set; }
    }
    [DataContract]
    public class CreateUserMeetingResponse
    {
        [DataMember]
        public Data Data { get; set; }
        [DataMember]
        public bool IsSuccess { get; set; }
        [DataMember]
        public string Message { get; set; }
    }
    [DataContract]
    public class CRMRequest
    {
        [DataMember]
        public string SlotDate { get; set; }
        [DataMember]
        public string RecordID { get; set; }
        [DataMember]
        public string SlotStart { get; set; }
        [DataMember]
        public string SlotEnd { get; set; }
        [DataMember]
        public bool IsVideoMeeting { get; set; }
    }
    [DataContract]
    public class GetUserTimeSlotsRequest
    {
        [DataMember]
        public string UserCode { get; set; }
        [DataMember]
        public string SlotDate { get; set; }
        [DataMember]
        public string EnquiryTypeCode { get; set; }
    }
    [DataContract]
    public class GetUserTimeSlotsDatum
    {
        [DataMember]
        public string SlotStart { get; set; }
        [DataMember]
        public string SlotEnd { get; set; }
        [DataMember]
        public int IsAvailable { get; set; }
    }

    [DataContract]
    public class GetUserTimeSlotsRoot
    {
        [DataMember]
        public List<GetUserTimeSlotsDatum> Data { get; set; }
        [DataMember]
        public bool IsSuccess { get; set; }
        [DataMember]
        public int ResponseCode { get; set; }
        [DataMember]
        public string Message { get; set; }
    }
}
