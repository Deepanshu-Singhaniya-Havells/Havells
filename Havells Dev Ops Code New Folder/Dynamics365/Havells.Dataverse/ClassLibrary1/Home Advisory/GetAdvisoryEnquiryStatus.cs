using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Havells.Dataverse.CustomConnector.Home_Advisory
{
    public class GetAdvisoryEnquiryStatus : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            string _datefilter = string.Empty;
            string JsonResponse = "";
            GetEnquiry _retResponse = new GetEnquiry();
            string dateRegex = @"^(0[1-9]|1[0-2])/(0[1-9]|[12][0-9]|3[01])/\d{4}$";
            try
            {
                string EnquiryId = Convert.ToString(context.InputParameters["EnquiryId"]);
                string FromDate = Convert.ToString(context.InputParameters["FromDate"]);
                string ToDate = Convert.ToString(context.InputParameters["ToDate"]);
                string MobileNo = Convert.ToString(context.InputParameters["MobileNo"]);

                if (string.IsNullOrWhiteSpace(EnquiryId) && string.IsNullOrWhiteSpace(MobileNo))
                {
                    _retResponse.StatusCode = "204";
                    _retResponse.StatusMessage = "EnquiryId or MobileNo is required";
                    JsonResponse = JsonSerializer.Serialize(_retResponse);
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                if (!string.IsNullOrWhiteSpace(EnquiryId))
                {
                    if (string.IsNullOrWhiteSpace(FromDate) || string.IsNullOrWhiteSpace(ToDate))
                    {
                        _retResponse.StatusCode = "204";
                        _retResponse.StatusMessage = "FromDate and ToDate are required";
                        JsonResponse = JsonSerializer.Serialize(_retResponse);
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }
                }
                if (!string.IsNullOrWhiteSpace(MobileNo))
                {
                    if (!APValidate.IsValidMobileNumber(MobileNo))
                    {
                        _retResponse.StatusCode = "204";
                        _retResponse.StatusMessage = "Invalid mobile number.";
                        JsonResponse = JsonSerializer.Serialize(_retResponse);
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }
                    if (string.IsNullOrWhiteSpace(FromDate) || string.IsNullOrWhiteSpace(ToDate))
                    {
                        _retResponse.StatusCode = "204";
                        _retResponse.StatusMessage = "FromDate and ToDate are required";
                        JsonResponse = JsonSerializer.Serialize(_retResponse);
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }
                }
                if (!string.IsNullOrWhiteSpace(FromDate))
                {
                    if (!Regex.IsMatch(FromDate, dateRegex))
                    {
                        _retResponse.StatusCode = "204";
                        _retResponse.StatusMessage = "Invalid FromDate format. It should be {MM/dd/yyyy}.";
                        JsonResponse = JsonSerializer.Serialize(_retResponse);
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }
                }
                if (!string.IsNullOrWhiteSpace(ToDate))
                {
                    if (!Regex.IsMatch(ToDate, dateRegex))
                    {
                        _retResponse.StatusCode = "204";
                        _retResponse.StatusMessage = "Invalid ToDate format. It should be {MM/dd/yyyy}.";
                        JsonResponse = JsonSerializer.Serialize(_retResponse);
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }
                }
                DateTime parsedFromDate = DateTime.ParseExact(FromDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                string formattedFromDate = parsedFromDate.ToString("yyyy-MM-dd");
                DateTime parsedToDate = DateTime.ParseExact(ToDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                string formattedToDate = parsedToDate.ToString("yyyy-MM-dd");
                if (FromDate != null && ToDate != null)
                {
                    _datefilter = @"<condition attribute='createdon' operator='on-or-after' value='" + formattedFromDate + @"' />
                    <condition attribute='createdon' operator='on-or-before' value='" + formattedToDate + @"' />";
                }
                if (!string.IsNullOrWhiteSpace(EnquiryId))
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
                                              <condition attribute='hil_name' operator='eq' value='" + EnquiryId + @"' />
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
                        List<GetAdvisoryEnquiryLine> enqlineList = getEnqueryLine(service, _addvisorycoll[0].Id, context);
                        homeadvisory.AdvisoryEnquiryLine = enqlineList;
                        homeadvisoryenqList.Add(homeadvisory);

                        _retResponse.StatusCode = "204";
                        _retResponse.StatusMessage = "Record Retrieved";
                        _retResponse.Success = true;
                        _retResponse.Enquiry = homeadvisoryenqList;

                        JsonResponse = JsonSerializer.Serialize(_retResponse);
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }
                    else
                    {
                        _retResponse.StatusCode = "400";
                        _retResponse.StatusMessage = "Enquiry Not found.";
                        JsonResponse = JsonSerializer.Serialize(_retResponse);
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }

                }
                else if (!string.IsNullOrWhiteSpace(MobileNo))
                {
                    #region fetch customer data...
                    QueryExpression query = new QueryExpression("contact");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, MobileNo.Trim());
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
                                List<GetAdvisoryEnquiryLine> enqlineList = getEnqueryLine(service, adv.Id, context);
                                homeadvisory.AdvisoryEnquiryLine = enqlineList;
                                homeadvisoryenqList.Add(homeadvisory);
                            }
                            _retResponse.StatusCode = "204";
                            _retResponse.StatusMessage = "Record Retrieved";
                            _retResponse.Success = true;
                            _retResponse.Enquiry = homeadvisoryenqList;

                            JsonResponse = JsonSerializer.Serialize(_retResponse);
                            context.OutputParameters["data"] = JsonResponse;
                            return;
                        }
                        else
                        {
                            _retResponse.StatusCode = "400";
                            _retResponse.StatusMessage = "Enquiry Not found.";
                            JsonResponse = JsonSerializer.Serialize(_retResponse);
                            context.OutputParameters["data"] = JsonResponse;
                            return;
                        }
                    }
                    else
                    {
                        _retResponse.StatusCode = "400";
                        _retResponse.StatusMessage = "Enquiry Not found.";
                        JsonResponse = JsonSerializer.Serialize(_retResponse);
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }
                    #endregion
                }
            }
            catch (Exception ex)
            {
                _retResponse.StatusCode = "505";
                _retResponse.StatusMessage = $"Dynamics 365 Internal Server Error : {ex.Message}";
                JsonResponse = JsonSerializer.Serialize(_retResponse);
                context.OutputParameters["data"] = JsonResponse;
                return;
            }
        }
        List<GetAdvisoryEnquiryLine> getEnqueryLine(IOrganizationService service, Guid advId, IPluginExecutionContext context)
        {
            List<GetAdvisoryEnquiryLine> enqlineList = new List<GetAdvisoryEnquiryLine>();
            string JsonResponse = "";
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
                        enqline.Documents = getEnqueryLineDocuments(service, advline.Id, context);
                        enqlineList.Add(enqline);
                    }
                }
            }
            catch (Exception ex)
            {
                GetEnquiry errorResponse = new GetEnquiry
                {
                    StatusCode = "505",
                    StatusMessage = $"Error in getEnqueryLine: {ex.Message}",
                    Success = false,
                    Enquiry = new List<GetHomeAdvisoryEnquiry>()
                };

                string jsonResponse = JsonSerializer.Serialize(errorResponse);
                context.OutputParameters["data"] = JsonResponse;
                return new List<GetAdvisoryEnquiryLine>();
            }
            return enqlineList;
        }
        List<GetDocument> getEnqueryLineDocuments(IOrganizationService service, Guid advEnqId, IPluginExecutionContext context)
        {
            List<GetDocument> _getDocument = new List<GetDocument>();
            string JsonResponse = "";
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
                GetEnquiry errorResponse = new GetEnquiry
                {
                    StatusCode = "505",
                    StatusMessage = $"Error in getEnqueryLineDocuments: {ex.Message}",
                    Success = false,
                    Enquiry = new List<GetHomeAdvisoryEnquiry>()
                };

                string jsonResponse = JsonSerializer.Serialize(errorResponse);
                context.OutputParameters["data"] = JsonResponse;
                return new List<GetDocument>();
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
            public string RoofTop { get; set; }
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
            public string AppointmentType { get; set; }
            public string AppointmentDate { get; set; }
            public string AppointmentStatus { get; set; }
            public string CustomerReamark { get; set; }
            public string AssignAdvisor { get; set; }
            public string PinCode { get; set; }
            public string VideoCallUrl { get; set; }
            public string AdvisoryDate { get; set; }
            public List<GetDocument> Documents { get; set; }
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
