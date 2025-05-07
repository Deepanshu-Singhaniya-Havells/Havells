using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Text;
using Newtonsoft.Json;
using System.Text.RegularExpressions;
using System.IO;
using System.Net;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class SFA_ServiceCall
    {
        [DataMember]
        public string SalesTrackingID { get; set; }
        [DataMember]
        public string InvoiceDate { get; set; }
        [DataMember]
        public string InvoiceNo { get; set; }
        [DataMember]
        public decimal? InvoiceValue { get; set; }
        [DataMember]
        public string PurchaseLocationPINCode { get; set; }
        [DataMember]
        public string Attachment { get; set; }
        [DataMember]
        public string FileName { get; set; }
        [DataMember]
        public string PurchaseFrom { get; set; }
        [DataMember]
        public string CustomerMobileNo { get; set; }
        [DataMember]
        public string CustomerName { get; set; }
        [DataMember]
        public string CustomerEmailId { get; set; }
        [DataMember]
        public Guid CustomerGuid { get; set; }
        [DataMember]
        public Guid AddressGuid { get; set; }
        [DataMember]
        public string Address { get; set; }
        [DataMember]
        public string PINCode { get; set; }
        [DataMember]
        public string AreaCode { get; set; }
        [DataMember]
        public string Source { get; set; }
        [DataMember]
        public bool ResultStatus { get; set; }
        [DataMember]
        public string ResultMessage { get; set; }
        [DataMember]
        public List<SFA_ServiceCallData> ServiceCallData { get; set; }

        public SFA_ServiceCall()
        {
            ServiceCallData = new List<SFA_ServiceCallData>();
            ResultStatus = false;
        }
        public SFA_ServiceCallResult SFA_CreateServiceCall(SFA_ServiceCall serviceCallData)
        {
            #region Variable declaration
            int _source = 7;
            string _requestData = string.Empty;
            DateTime _startDate = DateTime.Now;
            string _error = string.Empty;
            SFA_ServiceCallResult returnObj = new SFA_ServiceCallResult();
            returnObj.ResultMessage = "SUCCESS";
            returnObj.ResultStatus = true;
            returnObj.ServiceCallLogData = new List<SFA_ServiceCallLogData>();
            StringBuilder headerRemark = new StringBuilder();
            Guid _serviceCallRequestGuid;
            EntityCollection entcoll;
            QueryExpression query;
            IOrganizationService service = null;
            bool flag = true;
            Regex Regex_MobileNo = new Regex("^[6-9]\\d{9}$");
            Regex Regex_PinCode = new Regex("^[1-9]([0-9]){5}$");
            #endregion
            try
            {
                _requestData = JsonConvert.SerializeObject(serviceCallData);
                var CrmURL = "https://havellscrmdev1.crm8.dynamics.com/";
                string finalString = string.Format("AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=", CrmURL);
                service = HavellsConnection.CreateConnection.createConnection(finalString);

                #region Creating Service Call Request
                try
                {
                    if (string.IsNullOrWhiteSpace(serviceCallData.SalesTrackingID))
                    {
                        headerRemark.AppendLine("SalesTrackingID is required.");
                        flag = false;
                    }
                    if (!string.IsNullOrWhiteSpace(serviceCallData.Source))
                    {
                        try
                        {
                            _source = Convert.ToInt32(serviceCallData.Source);
                        }
                        catch (Exception)
                        {
                            headerRemark.AppendLine("Source must be a Number.");
                            flag = false;
                        }
                    }
                    if (serviceCallData.CustomerGuid == null || serviceCallData.CustomerGuid == Guid.Empty)
                    {
                        if (string.IsNullOrWhiteSpace(serviceCallData.CustomerMobileNo))
                        {
                            headerRemark.AppendLine("Customer Mobile No. is required.");
                            flag = false;
                        }
                        else if (!Regex_MobileNo.IsMatch(serviceCallData.CustomerMobileNo))
                        {
                            headerRemark.AppendLine("Invalid Customer Mobile No.");
                            flag = false;
                        }
                        if (string.IsNullOrWhiteSpace(serviceCallData.CustomerName))
                        {
                            headerRemark.AppendLine("Customer Name is required.");
                            flag = false;
                        }
                    }
                    else
                    {
                        Entity entityContact = service.Retrieve("contact", serviceCallData.CustomerGuid, new ColumnSet(false));
                        if (entityContact == null)
                        {
                            headerRemark.AppendLine("Invalid CustomerGuid!! CustomerGuid is not found in D365.");
                            flag = false;
                        }
                    }
                    if (string.IsNullOrWhiteSpace(serviceCallData.InvoiceDate))
                    {
                        headerRemark.AppendLine("Invoice Date is required.");
                        flag = false;
                    }
                    else
                    {
                        try
                        {
                            Convert.ToDateTime(serviceCallData.InvoiceDate);
                        }
                        catch (Exception)
                        {
                            headerRemark.AppendLine("Invalid Invoice Date!! Date format must be YYYY-MM-DD");
                            flag = false;
                        }
                    }
                    if (string.IsNullOrWhiteSpace(serviceCallData.InvoiceNo))
                    {
                        headerRemark.AppendLine("Invoice No. is required.");
                        flag = false;
                    }
                    if (serviceCallData.InvoiceValue == null)
                    {
                        headerRemark.AppendLine("Invoice Value is required.");
                        flag = false;
                    }
                    else if (serviceCallData.InvoiceValue <= 0)
                    {
                        headerRemark.AppendLine("Invalid Invoice Value.");
                        flag = false;
                    }
                    if (string.IsNullOrWhiteSpace(serviceCallData.PurchaseFrom))
                    {
                        headerRemark.AppendLine("Purchase From is required.");
                        flag = false;
                    }
                    if (string.IsNullOrWhiteSpace(serviceCallData.PurchaseLocationPINCode))
                    {
                        headerRemark.AppendLine("Purchase Location PIN Code is required.");
                        flag = false;
                    }
                    else if (!Regex_PinCode.IsMatch(serviceCallData.PurchaseLocationPINCode))
                    {
                        headerRemark.AppendLine("Invalid Purchase Location PIN Code.");
                        flag = false;
                    }
                    else
                    {
                        query = new QueryExpression("hil_pincode");
                        query.ColumnSet = new ColumnSet("hil_pincodeid");
                        query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, serviceCallData.PurchaseLocationPINCode);
                        entcoll = service.RetrieveMultiple(query);
                        if (entcoll.Entities.Count == 0)
                        {
                            headerRemark.AppendLine("Purchase Location PIN Code is not found in D365.");
                            flag = false;
                        }
                    }
                    if (string.IsNullOrWhiteSpace(serviceCallData.Attachment))
                    {
                        headerRemark.AppendLine("Invoice Copy is required.");
                        flag = false;
                    }
                    if (string.IsNullOrWhiteSpace(serviceCallData.FileName))
                    {
                        headerRemark.AppendLine("FileName is required.");
                        flag = false;
                    }
                    if (serviceCallData.ServiceCallData.Count == 0)
                    {
                        headerRemark.AppendLine("Model/Serial Number detail is missing.");
                        flag = false;
                    }
                    else
                    {
                        int i = 1;
                        foreach (SFA_ServiceCallData objServiceCallLine in serviceCallData.ServiceCallData)
                        {
                            if (_source == 15)
                            {
                                objServiceCallLine.Address = serviceCallData.Address;
                                objServiceCallLine.AddressGuid = serviceCallData.AddressGuid;
                                objServiceCallLine.AreaCode = serviceCallData.AreaCode;
                                objServiceCallLine.PINCode = serviceCallData.PINCode;
                            }
                            if (string.IsNullOrWhiteSpace(objServiceCallLine.CallType))
                            {
                                headerRemark.AppendLine("Line " + i.ToString() + ": Call Type is missing.");
                                flag = false;
                            }
                            else
                            {
                                if (objServiceCallLine.CallType != "P")
                                {
                                    if (string.IsNullOrWhiteSpace(objServiceCallLine.ServiceDate))
                                    {
                                        headerRemark.AppendLine("Line " + i.ToString() + ": Service Date is missing.");
                                        flag = false;
                                    }
                                    if (objServiceCallLine.AddressGuid == null || objServiceCallLine.AddressGuid == Guid.Empty)
                                    {
                                        if (string.IsNullOrWhiteSpace(objServiceCallLine.Address))
                                        {
                                            headerRemark.AppendLine("Line " + i.ToString() + ": Address Line is required.");
                                            flag = false;
                                        }
                                        if (string.IsNullOrWhiteSpace(objServiceCallLine.PINCode))
                                        {
                                            headerRemark.AppendLine("Line " + i.ToString() + ": PIN Code is required.");
                                            flag = false;
                                        }
                                        else if (!Regex_PinCode.IsMatch(objServiceCallLine.PINCode))
                                        {
                                            headerRemark.AppendLine("Line " + i.ToString() + ": Invalid PIN Code.");
                                            flag = false;
                                        }
                                        else
                                        {
                                            query = new QueryExpression("hil_pincode");
                                            query.ColumnSet = new ColumnSet("hil_pincodeid");
                                            query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, objServiceCallLine.PINCode);
                                            entcoll = service.RetrieveMultiple(query);
                                            if (entcoll.Entities.Count == 0)
                                            {
                                                headerRemark.AppendLine("Line " + i.ToString() + ": PIN Code is not found in D365.");
                                                flag = false;
                                            }
                                        }
                                        if (string.IsNullOrWhiteSpace(objServiceCallLine.AreaCode))
                                        {
                                            headerRemark.AppendLine("Line " + i.ToString() + ": Job Address Area Code is missing.");
                                            flag = false;
                                        }
                                        else
                                        {
                                            query = new QueryExpression("hil_area");
                                            query.ColumnSet = new ColumnSet("hil_areaid");
                                            query.Criteria = new FilterExpression(LogicalOperator.And);
                                            query.Criteria.AddCondition("hil_areacode", ConditionOperator.Equal, objServiceCallLine.AreaCode);
                                            entcoll = service.RetrieveMultiple(query);
                                            if (entcoll.Entities.Count == 0)
                                            {
                                                headerRemark.AppendLine("Line " + i.ToString() + ": Area Code is not found in D365.");
                                                flag = false;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Entity entityAddress = service.Retrieve("hil_address", objServiceCallLine.AddressGuid, new ColumnSet(false));
                                        if (entityAddress == null)
                                        {
                                            headerRemark.AppendLine("Line " + i.ToString() + ": Invalid AddressGuid!! AddressGuid is not found in D365.");
                                            flag = false;
                                        }
                                    }
                                }
                            }
                            if (objServiceCallLine.Qty <= 0)
                            {
                                headerRemark.AppendLine("Line " + i.ToString() + ": Qty must be greater-than 0");
                                flag = false;
                            }
                            if (string.IsNullOrWhiteSpace(objServiceCallLine.SerialNumber))
                            {
                                headerRemark.AppendLine("Line " + i.ToString() + ": Serial Number is missing.");
                                flag = false;
                            }
                            if (string.IsNullOrWhiteSpace(objServiceCallLine.ModelNumber))
                            {
                                headerRemark.AppendLine("Line " + i.ToString() + ": Model Number is missing.");
                                flag = false;
                            }
                            else
                            {
                                query = new QueryExpression("product");
                                query.ColumnSet = new ColumnSet("description");
                                query.Criteria = new FilterExpression(LogicalOperator.And);
                                query.Criteria.AddCondition("productnumber", ConditionOperator.Equal, objServiceCallLine.ModelNumber);
                                EntityCollection ec1 = service.RetrieveMultiple(query);
                                if (ec1.Entities.Count == 0)
                                {
                                    headerRemark.AppendLine("Line " + i.ToString() + ": Model Number does not exist in D365.");
                                    flag = false;
                                }
                            }
                            i += 1;
                        }
                    }
                    if (!flag)
                    {
                        return new SFA_ServiceCallResult { ResultStatus = false, ResultMessage = headerRemark.ToString(), ServiceCallLogData = new List<SFA_ServiceCallLogData>() };
                    }

                    Entity enServiceCallRequest = new Entity("hil_servicecallrequest");
                    enServiceCallRequest["hil_name"] = serviceCallData.SalesTrackingID;
                    if (serviceCallData.CustomerGuid != null && serviceCallData.CustomerGuid != Guid.Empty)
                    {
                        enServiceCallRequest["hil_consumer"] = new EntityReference("contact", serviceCallData.CustomerGuid);
                    }
                    if (!string.IsNullOrWhiteSpace(serviceCallData.CustomerEmailId))
                    {
                        enServiceCallRequest["hil_customeremailid"] = serviceCallData.CustomerEmailId;
                    }
                    if (!string.IsNullOrWhiteSpace(serviceCallData.CustomerMobileNo))
                    {
                        enServiceCallRequest["hil_customermobilenumber"] = serviceCallData.CustomerMobileNo;
                    }
                    if (!string.IsNullOrWhiteSpace(serviceCallData.CustomerName))
                    {
                        enServiceCallRequest["hil_customername"] = serviceCallData.CustomerName;
                    }
                    if (serviceCallData.AddressGuid != null && serviceCallData.AddressGuid != Guid.Empty)
                    {
                        enServiceCallRequest["hil_address"] = new EntityReference("hil_address", serviceCallData.AddressGuid);
                    }
                    if (!string.IsNullOrWhiteSpace(serviceCallData.Address))
                    {
                        enServiceCallRequest["hil_fulladdress"] = serviceCallData.Address;
                    }
                    if (!string.IsNullOrWhiteSpace(serviceCallData.PINCode))
                    {
                        enServiceCallRequest["hil_pincode"] = serviceCallData.PINCode;
                    }
                    if (!string.IsNullOrWhiteSpace(serviceCallData.AreaCode))
                    {
                        enServiceCallRequest["hil_areacode"] = serviceCallData.AreaCode;
                    }
                    enServiceCallRequest["hil_source"] = new OptionSetValue(_source);
                    enServiceCallRequest["hil_invoicedate"] = Convert.ToDateTime(serviceCallData.InvoiceDate);//YYYY-MM-DD
                    enServiceCallRequest["hil_invoicenumber"] = serviceCallData.InvoiceNo;
                    enServiceCallRequest["hil_invoicevalue"] = new Money(Convert.ToDecimal(serviceCallData.InvoiceValue));
                    enServiceCallRequest["hil_filename"] = serviceCallData.FileName;
                    enServiceCallRequest["hil_purchasefrom"] = serviceCallData.PurchaseFrom;
                    enServiceCallRequest["hil_purchaselocation"] = serviceCallData.PurchaseLocationPINCode;
                    enServiceCallRequest["hil_image"] = Convert.FromBase64String(serviceCallData.Attachment);

                    QueryExpression servicecallrequestquery = new QueryExpression("hil_servicecallrequest");
                    servicecallrequestquery.ColumnSet = new ColumnSet(false);
                    servicecallrequestquery.Criteria = new FilterExpression(LogicalOperator.And);
                    servicecallrequestquery.Criteria.AddCondition("hil_name", ConditionOperator.Equal, serviceCallData.SalesTrackingID);
                    servicecallrequestquery.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);//Active
                    EntityCollection servicecallrequestcoll = service.RetrieveMultiple(servicecallrequestquery);
                    if (servicecallrequestcoll.Entities.Count > 0)
                    {
                        _serviceCallRequestGuid = servicecallrequestcoll.Entities[0].Id;
                    }
                    else
                    {
                        _serviceCallRequestGuid = service.Create(enServiceCallRequest);
                    }

                    if (_serviceCallRequestGuid != Guid.Empty)
                    {
                        Entity enServiceCallRequestDetails = null;
                        foreach (SFA_ServiceCallData objServiceCall in serviceCallData.ServiceCallData)
                        {
                            enServiceCallRequestDetails = new Entity("hil_servicecallrequestdetail");

                            if (objServiceCall.AddressGuid != null && objServiceCall.AddressGuid != Guid.Empty)
                            {
                                enServiceCallRequestDetails["hil_address"] = new EntityReference("hil_address", objServiceCall.AddressGuid);
                            }
                            if (!string.IsNullOrWhiteSpace(objServiceCall.Address))
                            {
                                enServiceCallRequestDetails["hil_fulladdress"] = objServiceCall.Address;
                            }
                            if (!string.IsNullOrWhiteSpace(objServiceCall.AreaCode))
                            {
                                enServiceCallRequestDetails["hil_areacode"] = objServiceCall.AreaCode;
                            }
                            if (!string.IsNullOrWhiteSpace(objServiceCall.PINCode))
                            {
                                enServiceCallRequestDetails["hil_pincode"] = objServiceCall.PINCode;
                            }
                            if (objServiceCall.CallType == "D")
                            {
                                enServiceCallRequestDetails["hil_callsubtype"] = new EntityReference("hil_callsubtype", new Guid("AE1B2B71-3C0B-E911-A94E-000D3AF06CD4"));
                            }
                            else if (objServiceCall.CallType == "I")
                            {
                                enServiceCallRequestDetails["hil_callsubtype"] = new EntityReference("hil_callsubtype", new Guid("E3129D79-3C0B-E911-A94E-000D3AF06CD4"));
                            }
                            else if (objServiceCall.CallType == "B")
                            {
                                enServiceCallRequestDetails["hil_callsubtype"] = new EntityReference("hil_callsubtype", new Guid("13254938-A783-EA11-A811-000D3AF055B6"));
                            }
                            else
                            {
                                enServiceCallRequestDetails["hil_callsubtype"] = new EntityReference("hil_callsubtype", new Guid("27630B63-3C0B-E911-A94E-000D3AF06CD4"));
                            }
                            if (objServiceCall.ModelNumber != null)
                            {
                                query = new QueryExpression("product");
                                query.ColumnSet = new ColumnSet("description");
                                query.Criteria = new FilterExpression(LogicalOperator.And);
                                query.Criteria.AddCondition("productnumber", ConditionOperator.Equal, objServiceCall.ModelNumber);
                                EntityCollection ec1 = service.RetrieveMultiple(query);
                                if (ec1.Entities.Count > 0)
                                {
                                    enServiceCallRequestDetails["hil_product"] = ec1.Entities[0].ToEntityReference();
                                }
                            }
                            enServiceCallRequestDetails["hil_qty"] = objServiceCall.Qty;
                            enServiceCallRequestDetails["hil_serialnumber"] = objServiceCall.SerialNumber;
                            enServiceCallRequestDetails["hil_name"] = objServiceCall.ModelNumber;
                            enServiceCallRequestDetails["hil_servicedate"] = Convert.ToDateTime(objServiceCall.ServiceDate);//YYYY-MM-DD
                            enServiceCallRequestDetails["hil_servicecallrequest"] = new EntityReference("hil_servicecallrequest", _serviceCallRequestGuid);
                            service.Create(enServiceCallRequestDetails);
                        }
                    }
                }
                catch (Exception ex)
                {
                    return new SFA_ServiceCallResult { ResultStatus = false, ResultMessage = headerRemark.ToString() + ex.Message, ServiceCallLogData = new List<SFA_ServiceCallLogData>() };
                }
                #endregion
            }
            catch (Exception ex)
            {
                _error = ex.Message;
                returnObj = new SFA_ServiceCallResult { ResultStatus = false, ResultMessage = "D365 Internal Server Error : " + ex.Message, ServiceCallLogData = new List<SFA_ServiceCallLogData>() };
            }
            finally
            {
                string _responseData = JsonConvert.SerializeObject(returnObj);
                DateTime _endDate = DateTime.Now;
                APIExceptionLog.LogAPIExecution(string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", "SFA App", "SFACreateServiceCall", _error, _requestData.Replace(",", "~"), _responseData.Replace(",", "~"), _startDate.ToString(), _endDate.ToString(), (_endDate - _startDate).TotalSeconds.ToString() + " sec"),"SFAServiceCall");
            }
            return returnObj;
        }
    }
    public class SFA_ServiceCallData
    {
        public Guid AddressGuid { get; set; }
        public string SerialNumber { get; set; }
        public string Address { get; set; }
        public string PINCode { get; set; }
        public string AreaCode { get; set; }
        public string CallType { get; set; }
        public string ModelNumber { get; set; }
        public int Qty { get; set; }
        public Guid PinCodeGuid { get; set; }
        public Guid AreaCodeGuid { get; set; }
        public Guid BizGeoGuid { get; set; }
        public string ServiceDate { get; set; }
    }

    public class SFA_ServiceCallResult
    {
        public bool ResultStatus { get; set; }
        public string ResultMessage { get; set; }
        public List<SFA_ServiceCallLogData> ServiceCallLogData { get; set; }
        public SFA_ServiceCallResult()
        {
            ServiceCallLogData = new List<SFA_ServiceCallLogData>();
            ResultStatus = false;
        }
    }

    public class SFA_ServiceCallLogData
    {
        public string CallType { get; set; }
        public string ModelNumber { get; set; }
        public string ServiceCallNo { get; set; }
        public Guid ServiceCallGuid { get; set; }
    }
}
