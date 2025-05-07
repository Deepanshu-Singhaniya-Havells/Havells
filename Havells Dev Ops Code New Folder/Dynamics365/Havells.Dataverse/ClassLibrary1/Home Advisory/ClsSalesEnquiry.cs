using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Havells.Dataverse.CustomConnector.Home_Advisory
{
    public class ClsSalesEnquiry : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            List<EnquiryDetails> enquiryDetails = new List<EnquiryDetails>();
            EnquiryStatus enquiryStatus = new EnquiryStatus();

            if (context.InputParameters.Contains("EnquiryCategory") && (!string.IsNullOrWhiteSpace(context.InputParameters["EnquiryCategory"].ToString())))
            {
                string _EnquiryCategory = context.InputParameters["EnquiryCategory"].ToString();
                if (APValidate.IsNumeric(_EnquiryCategory) && _EnquiryCategory.Length <= 9)
                {
                    enquiryStatus.EnquiryId = _EnquiryCategory;
                }
                else
                {
                    enquiryDetails.Add(new EnquiryDetails() { ErrorCode = "204", Remarks = "EnquiryCategory is not valid." });
                    string jsonErrorResult = JsonSerializer.Serialize(enquiryDetails);
                    context.OutputParameters["data"] = jsonErrorResult;
                    return;
                }
            }
            if (context.InputParameters.Contains("MobileNo") && (!string.IsNullOrWhiteSpace(context.InputParameters["MobileNo"].ToString())))
            {
                string _MobileNo = context.InputParameters["MobileNo"].ToString();
                if (APValidate.IsValidMobileNumber(_MobileNo))
                {
                    enquiryStatus.MobileNo = _MobileNo;
                }
                else
                {
                    enquiryDetails.Add(new EnquiryDetails() { ErrorCode = "204", Remarks = "MobileNo is not valid." });
                    string jsonErrorResult = JsonSerializer.Serialize(enquiryDetails);
                    context.OutputParameters["data"] = jsonErrorResult;
                    return;
                }
            }
            if (context.InputParameters.Contains("FromDate") && (!string.IsNullOrWhiteSpace(context.InputParameters["FromDate"].ToString())))
            {
                string _FromDate = context.InputParameters["FromDate"].ToString();
                string pattern = @"^(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[012])/\d{4}$";
                bool isMatch = Regex.IsMatch(_FromDate, pattern);
                if (isMatch)
                {
                    //enquiryStatus.FromDate = _FromDate;
                    DateTime fromDate = DateTime.ParseExact(_FromDate, "dd/MM/yyyy", CultureInfo.InvariantCulture); 
                    enquiryStatus.FromDate = fromDate.ToString("yyyy-MM-dd");

                }
                else
                {
                    enquiryDetails.Add(new EnquiryDetails() { ErrorCode = "204", Remarks = "FromDate is not valid format.please try with {dd/MM/yyyy}." });
                    string jsonErrorResult = JsonSerializer.Serialize(enquiryDetails);
                    context.OutputParameters["data"] = jsonErrorResult;
                    return;
                }
            }
            else
            {
                enquiryDetails.Add(new EnquiryDetails() { ErrorCode = "204", Remarks = "FromDate is required with {dd/MM/yyyy}." });
                string jsonErrorResult = JsonSerializer.Serialize(enquiryDetails);
                context.OutputParameters["data"] = jsonErrorResult;
                return;
            }

            if (context.InputParameters.Contains("ToDate") && (!string.IsNullOrWhiteSpace(context.InputParameters["ToDate"].ToString())))
            {
                string _ToDate = context.InputParameters["ToDate"].ToString();
                string pattern = @"^(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[012])/\d{4}$";
                bool isMatch = Regex.IsMatch(_ToDate, pattern);
                if (isMatch) //if (APValidate.IsvalidDate(_ToDate))
                {
                    //enquiryStatus.ToDate = _ToDate;
                    DateTime toDate = DateTime.ParseExact(_ToDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    enquiryStatus.ToDate = toDate.ToString("yyyy-MM-dd");
                }
                else
                {
                    enquiryDetails.Add(new EnquiryDetails() { ErrorCode = "204", Remarks = "ToDate is not valid format.please try with {dd/MM/yyyy}." });
                    string jsonErrorResult = JsonSerializer.Serialize(enquiryDetails);
                    context.OutputParameters["data"] = jsonErrorResult;
                    return;
                }
            }
            else
            {
                enquiryDetails.Add(new EnquiryDetails() { ErrorCode = "204", Remarks = "ToDate is is required with {dd/MM/yyyy}." });
                string jsonErrorResult = JsonSerializer.Serialize(enquiryDetails);
                context.OutputParameters["data"] = jsonErrorResult;
                return;
            }
            if (DateTime.Parse(enquiryStatus.ToDate) == DateTime.Parse(enquiryStatus.FromDate))
            {
                enquiryDetails.Add(new EnquiryDetails() { ErrorCode = "204", Remarks = "FromDate is not equal to ToDate. " });
                string jsonErrorResult = JsonSerializer.Serialize(enquiryDetails);
                context.OutputParameters["data"] = jsonErrorResult;
                return;
            }
            if (DateTime.Parse(enquiryStatus.ToDate) < DateTime.Parse(enquiryStatus.FromDate))
            {
                enquiryDetails.Add(new EnquiryDetails() { ErrorCode = "204", Remarks = "ToDate should be greater than From Date. " });
                string jsonErrorResult = JsonSerializer.Serialize(enquiryDetails);
                context.OutputParameters["data"] = jsonErrorResult;
                return;
            }

            List<EnquiryDetails> result = GetSalesEnquiry(enquiryStatus, service);

            string jsonResult = JsonSerializer.Serialize(result);
            context.OutputParameters["data"] = jsonResult;
        }
        public List<EnquiryDetails> GetSalesEnquiry(EnquiryStatus _enquiryStatus, IOrganizationService service)
        {
            List<EnquiryDetails> lstEnquiryDetails = new List<EnquiryDetails>();
            try
            {
                bool _paramCheck = false;

                string _filterStr = @"<filter type='and'><condition attribute='statecode' operator='eq' value='0' />";
                if (_enquiryStatus.FromDate != null && _enquiryStatus.ToDate != null)
                {
                    _filterStr += @"<condition attribute='createdon' operator='on-or-after' value='" + _enquiryStatus.FromDate + @"' />
                        <condition attribute='createdon' operator='on-or-before' value='" + _enquiryStatus.ToDate + @"' />";
                    _paramCheck = true;
                }
                if (_enquiryStatus.MobileNo != null && _enquiryStatus.MobileNo != "" && _enquiryStatus.MobileNo.Trim().Length > 0)
                {
                    _filterStr += @"<condition attribute='mobilephone' operator='eq' value='" + _enquiryStatus.MobileNo + @"' />";
                    _paramCheck = true;
                }
                if (_enquiryStatus.EnquiryId != null && _enquiryStatus.EnquiryId != "" && _enquiryStatus.EnquiryId.Trim().Length > 0)
                {
                    _filterStr += @"<condition attribute='hil_ticketnumber' operator='eq' value='" + _enquiryStatus.EnquiryId + @"' />";
                    _paramCheck = true;
                }
                _filterStr += @"</filter>";
                if (!_paramCheck)
                {
                    lstEnquiryDetails.Add(new EnquiryDetails() { ErrorCode = "204", Remarks = "ERROR!!! Please input From Date/ToDate/Mobileno/EnquiryId." });
                    return lstEnquiryDetails;
                }
                string _fetchxml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='lead'>
                        <attribute name='fullname' />
                        <attribute name='createdon' />
                        <attribute name='hil_selecteddivisionsname' />
                        <attribute name='hil_productdivision' />
                        <attribute name='leadid' />
                        <attribute name='hil_ticketnumber' />
                        <attribute name='hil_leadtype' />
                        <attribute name='description' />
                        <attribute name='hil_pincode' />
                        <attribute name='hil_leadstatus' />
                        <order attribute='createdon' descending='true' />" + _filterStr + @"
                      </entity>
                    </fetch>";

                EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchxml));
                if (entCol.Entities.Count > 0)
                {
                    foreach (Entity ent in entCol.Entities)
                    {
                        lstEnquiryDetails.Add(new EnquiryDetails()
                        {
                            CreatedOn = ent.Contains("createdon") ? ent.GetAttributeValue<DateTime>("createdon").ToString() : "",
                            CustomerName = ent.Contains("fullname") ? ent.GetAttributeValue<string>("fullname").ToString() : "",
                            DivisionName = ent.Contains("hil_productdivision") ? ent.GetAttributeValue<EntityReference>("hil_productdivision").Name.ToString() : "",
                            //DivisionName = ent.Contains("hil_selecteddivisionsname") ? ent.GetAttributeValue<string>("hil_selecteddivisionsname").ToString() : "",
                            Remarks = ent.Contains("description") ? ent.GetAttributeValue<string>("description").ToString() : "",
                            EnquiryStatus = ent.Contains("hil_leadstatus") ? ent.FormattedValues["hil_leadstatus"].ToString() : "",
                            EnquiryType = ent.Contains("hil_leadtype") ? ent.FormattedValues["hil_leadtype"].ToString() : "",
                            PinCode = ent.Contains("hil_pincode") ? ent.GetAttributeValue<EntityReference>("hil_pincode").Name : "",
                            EnquiryId = ent.Contains("hil_ticketnumber") ? ent.GetAttributeValue<string>("hil_ticketnumber").ToString() : "",
                            ErrorCode = "Success"
                        });
                    }
                }
                else
                {
                    lstEnquiryDetails.Add(new EnquiryDetails() { ErrorCode = "204", Remarks = "No Data found." });
                }
                return lstEnquiryDetails;
            }
            catch (Exception ex)
            {
                lstEnquiryDetails.Add(new EnquiryDetails { ErrorCode = "204", Remarks = "D365 Internal Server Error : " + ex.Message });
                return lstEnquiryDetails;
            }
        }
    }
    public class EnquiryStatus
    {
        public string EnquiryId { get; set; }
        public string MobileNo { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
    }
    public class EnquiryDetails
    {
        public string CustomerName { get; set; }
        public string CreatedOn { get; set; }
        public string DivisionName { get; set; }
        public string Remarks { get; set; }
        public string EnquiryType { get; set; }
        public string PinCode { get; set; }
        public string EnquirySource { get; set; }
        public string EnquiryStatus { get; set; }
        public string ErrorCode { get; set; }
        public string EnquiryId { get; set; }
    }
}
