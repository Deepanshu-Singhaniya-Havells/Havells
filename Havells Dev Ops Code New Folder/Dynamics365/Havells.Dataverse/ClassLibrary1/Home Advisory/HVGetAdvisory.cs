using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Home_Advisory
{
    public class HVGetAdvisory : IPlugin
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
                Enquiry request = new Enquiry
                {
                    MobileNo = Convert.ToString(context.InputParameters["MobileNo"]),
                    AdvisoryID = Convert.ToString(context.InputParameters["AdvisoryID"]),
                };
                if (string.IsNullOrWhiteSpace(request.MobileNo) && string.IsNullOrWhiteSpace(request.AdvisoryID))
                {
                    context.OutputParameters["data"] = JsonSerializer.Serialize(new GetEnquiry { StatusCode = "204", StatusMessage = "MobileNo / AdvisoryID is required." });
                    return;
                }
                if (!string.IsNullOrWhiteSpace(request.MobileNo))
                {
                    if (!APValidate.IsValidMobileNumber(request.MobileNo))
                    {
                        context.OutputParameters["data"] = JsonSerializer.Serialize(new GetEnquiry { StatusCode = "204", StatusMessage = "Invalid MobileNo number." });
                        return;
                    }
                }
                if (!string.IsNullOrWhiteSpace(request.AdvisoryID))
                {
                    if (request.AdvisoryID.Length > 16)
                    {
                        context.OutputParameters["data"] = JsonSerializer.Serialize(new GetEnquiry { StatusCode = "204", StatusMessage = "Invalid AdvisoryID." });
                        return;
                    }
                    if (!APValidate.NumericValue(request.AdvisoryID))
                    {
                        context.OutputParameters["data"] = JsonSerializer.Serialize(new GetEnquiry { StatusCode = "204", StatusMessage = "Invalid AdvisoryID." });
                        return;
                    }
                }
                GetEnquiry response = GetEnquery(request, service);
                jsonResponse = JsonSerializer.Serialize(response);
                context.OutputParameters["data"] = jsonResponse;
            }
            catch (Exception ex)
            {
                var GetEnquiry = new GetEnquiry
                {
                    StatusCode = "503",
                    StatusMessage = "D365 Internal Server Error: " + ex.Message
                };
                context.OutputParameters["data"] = JsonSerializer.Serialize(GetEnquiry);
            }
        }
        public GetEnquiry GetEnquery(Enquiry req, IOrganizationService service)
        {
            GetEnquiry _retResponse = new GetEnquiry();
            try
            {
                if (!string.IsNullOrWhiteSpace(req.AdvisoryID))
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
                        List<GetAdvisoryEnquiryLine> enqlineList = GetEnqueryLine(service, _addvisorycoll[0].Id);
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
                else if (!string.IsNullOrWhiteSpace(req.MobileNo))
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
                                List<GetAdvisoryEnquiryLine> enqlineList = GetEnqueryLine(service, adv.Id);
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
                            _retResponse.StatusCode = "404";
                            _retResponse.Success = true;
                            _retResponse.StatusMessage = "Record Not Found";
                            return _retResponse;
                        }
                    }
                    else
                    {
                        _retResponse.StatusCode = "404";
                        _retResponse.StatusMessage = "Contact Not Found";
                        _retResponse.Success = true;
                        return _retResponse;
                    }
                    #endregion
                }
            }
            catch (Exception ex)
            {
                _retResponse.StatusCode = "503";
                _retResponse.Success = false;
                _retResponse.StatusMessage = "Dynamics 365 Internal Server Error : " + ex.Message;
                return _retResponse;
            }
            return _retResponse;
        }
        List<GetAdvisoryEnquiryLine> GetEnqueryLine(IOrganizationService service, Guid advId)
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
                        enqline.Documents = GetEnqueryLineDocuments(service, advline.Id);
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
        List<GetDocument> GetEnqueryLineDocuments(IOrganizationService service, Guid advEnqId)
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
        public class GetEnquiry
        {
            public string StatusCode { get; set; }
            public string StatusMessage { get; set; }
            public bool Success { get; set; }
            public List<GetHomeAdvisoryEnquiry> Enquiry { get; set; }
        }
        public class GetHomeAdvisoryEnquiry
        {
            public string AdvisoryID { get; set; }
            public string AdvisoryGuid { get; set; }
            public string CustomerType { get; set; }
            public string Area { get; set; }
            public string PropertyType { get; set; }
            public string ConsturctionType { get; set; }
            public String RoofTop { get; set; }
            public string Customer { get; set; }
            public string AssetType { get; set; }
            public string MobileNumber { get; set; }
            public string EmailID { get; set; }
            public string City { get; set; }
            public string State { get; set; }
            public string PinCode { get; set; }
            public string TDS { get; set; }
            public string CustomerRemarks { get; set; }
            public List<GetAdvisoryEnquiryLine> AdvisoryEnquiryLine { get; set; }
        }
        public class GetAdvisoryEnquiryLine
        {
            public string EnquiryLineGuid { get; set; }
            public string EnquiryLineID { get; set; }
            public string TypeofEnquiry { get; set; }
            public string EnquiryStatus { get; set; }
            public string TypeofProduct { get; set; }
            public string AppointmentId { get; set; }
            public String AppointmentType { get; set; }
            public string AppointmentDate { get; set; }
            public string AppointmentStatus { get; set; }
            public string CustomerReamark { get; set; }
            public string AssignAdvisor { get; set; }
            public string PinCode { get; set; }
            public string VideoCallUrl { get; set; }
            public string AdvisoryDate { get; set; }
            public List<GetDocument> Documents { get; set; }
        }
        public class Enquiry
        {
            public string AdvisoryID { get; set; }
            public string MobileNo { get; set; }
        }
        public class GetDocument
        {
            public string Subject { get; set; }
            public string DocumentURL { get; set; }
            public string DocumentType { get; set; }
            public string DocumentSize { get; set; }
        }
    }
}

