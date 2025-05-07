using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Text.RegularExpressions;

namespace MRF_API
{
    public class MRFAPI
    {
        private readonly IOrganizationService service;

        public MRFAPI(IOrganizationService _service)
        {
            service = _service;
        }
        public SFA_ServiceCallResult SFA_CreateServiceCall(SFA_ServiceCall serviceCallData)
        {
            #region Variable declaration
            int _source = 7;
            string _requestData = JsonConvert.SerializeObject(serviceCallData);
            string _responseData = string.Empty;
            DateTime _startDate = DateTime.Now;
            string _error = string.Empty;

            SFA_ServiceCallResult returnObj = new SFA_ServiceCallResult();
            returnObj.ResultMessage = "SUCCESS";
            returnObj.ResultStatus = true;
            returnObj.ServiceCallLogData = new List<SFA_ServiceCallLogData>();
            StringBuilder headerRemark = new StringBuilder();
            StringBuilder lineRemark = new StringBuilder();

            Guid _customerGuid = serviceCallData.CustomerGuid;
            Guid _addressGuid = serviceCallData.AddressGuid;
            if (serviceCallData.Source != null)
            {
                _source = Convert.ToInt32(serviceCallData.Source);
            }

            //serviceCallData.CustomerName = serviceCallData.CustomerFirstName + " " + serviceCallData.CustomerLastName;

            Guid _serviceCallRequestGuid;
            EntityCollection entcoll;
            QueryExpression query;
            IOrganizationService service = null;
            string customerFullName = serviceCallData.CustomerName;
            string customerMobileNumber = serviceCallData.CustomerMobileNo;
            string customerEmail = serviceCallData.CustomerEmailId;
            bool flag = true;
            string InvoiceDate = string.Empty;
            string InvoiceNo = string.Empty;
            string InvoiceValue = string.Empty;
            string PurchaseLocationPINCode = string.Empty;
            string Attachment = string.Empty;
            string PurchaseFrom = string.Empty;
            string FileName = string.Empty;
            #endregion

            try
            {
                var CrmURL = "https://havells.crm8.dynamics.com/";
                string finalString = string.Format("AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=", CrmURL);
                service = HavellsConnection.CreateConnection.createConnection(finalString);

                #region Creating Service Call Request
                try
                {

                    if (_customerGuid == Guid.Empty)
                    {
                        if (serviceCallData.CustomerMobileNo.Trim().Length == 0 || serviceCallData.CustomerMobileNo == null)
                        {
                            headerRemark.AppendLine("Customer Mobile No. is required.");
                            flag = false;
                        }
                        if (string.IsNullOrEmpty(serviceCallData.CustomerName))// == null || serviceCallData.CustomerName.Trim().Length == 0 )
                        {
                            if (string.IsNullOrEmpty(serviceCallData.CustomerFirstName))//.Trim().Length == 0 || serviceCallData.CustomerName == null)
                            {
                                headerRemark.AppendLine("Customer Name is required.");
                                flag = false;
                            }
                            //headerRemark.AppendLine("Customer Name is required.");
                            //flag = false;
                        }
                    }

                    if (serviceCallData.InvoiceDate.Trim().Length == 0 || serviceCallData.InvoiceDate == null)
                    {
                        headerRemark.AppendLine("Invoice Date is required.");
                        flag = false;
                    }

                    if (string.IsNullOrWhiteSpace(serviceCallData.InvoiceNo) )
                    {
                        headerRemark.AppendLine("Invalid InvoiceNo");



                        flag = false;

                    }
                    else
                    {
                        Regex rex = new Regex("^([0])");

                        if (rex.IsMatch(serviceCallData.InvoiceNo))
                        {
                            headerRemark.AppendLine("Please enter a valid Invoice No");
                        }

                        Regex re = new Regex("^([0-9a-zA-Z-_]){1,30}$");

                      
                        if (!re.IsMatch(serviceCallData.InvoiceNo))
                        {
                            headerRemark.AppendLine("Please enter a valid Invoice No");
                            flag = false;
                        }
                    }



                    if (serviceCallData.InvoiceValue == null || serviceCallData.InvoiceValue == 0)
                    {
                        
                        headerRemark.AppendLine("Invalid Invoice Value");
                        flag = false;
                    }
                    else
                    {
                        Regex re = new Regex("^[0-9.]+$");

                        if (!re.IsMatch(serviceCallData.InvoiceValue.ToString()))
                        {
                            headerRemark.AppendLine("Please enter a valid Invoice Value");
                            flag = false;
                        }
                    }

                    if (string.IsNullOrEmpty(serviceCallData.PurchaseFrom))//.Trim().Length == 0 || serviceCallData.PurchaseFrom == null)
                    {
                        if (string.IsNullOrEmpty(serviceCallData.PurchaseFromCode) || string.IsNullOrEmpty(serviceCallData.PurchaseFromName))//.Trim().Length == 0 || serviceCallData.PurchaseFrom == null)
                        {
                            headerRemark.AppendLine("Purchase From is required.");
                            flag = false;
                        }
                        //headerRemark.AppendLine("Purchase From is required.");
                        //flag = false;
                    }
                    else
                    {
                        Regex purchasecode = new Regex("^([0-9a-zA-Z]){1,100}$");

                        if (!purchasecode.IsMatch(serviceCallData.PurchaseFromCode))
                        {
                            headerRemark.AppendLine("Invalid Purchase From Code");
                            flag = false;
                        }
                        Regex purchasename = new Regex("^([0-9a-zA-Z]){1,100}$");

                        if (!purchasename.IsMatch(serviceCallData.PurchaseFromName))
                        {
                            headerRemark.AppendLine("Invalid Purchase From Name");
                            flag = false;
                        }
                    }

                    if (serviceCallData.PurchaseLocationPINCode.Trim().Length == 0 || serviceCallData.PurchaseLocationPINCode == null)
                    {
                        headerRemark.AppendLine("Purchase Location PIN Code is required.");
                        flag = false;
                    }
                    else
                    {
                      
                        Regex re = new Regex("^[1-9]([0-9]){5}$");

                        if (!re.IsMatch(serviceCallData.PurchaseLocationPINCode))
                        {
                            headerRemark.AppendLine("Please enter a valid PIN Code");
                            flag = false;
                        }
                    }

                    if (serviceCallData.Attachment.Trim().Length == 0 || serviceCallData.Attachment == null)
                    {
                        headerRemark.AppendLine("Invoice Copy is required.");
                        flag = false;
                    }
                    else
                    {
                        if(!TryGetFromBase64String(serviceCallData.Attachment)){
                            headerRemark.AppendLine("Invalid Attachment.");
                            flag = false;   
                        }
                    }
                    
                    {
                        Regex re = new Regex("^(?:[\\w]\\:|\\\\)(\\\\[a-z_\\-\\s0-9\\.]+)+\\.(png|gif|jpg|jpeg)$\r\n");

                        if (!re.IsMatch(serviceCallData.PurchaseLocationPINCode))
                        {
                            headerRemark.AppendLine("Please enter a valid PIN Code");
                            flag = false;
                        }
                    }
                    if (serviceCallData.FileName.Trim().Length == 0 || serviceCallData.FileName == null)
                    {
                        headerRemark.AppendLine("FileName is required.");
                        flag = false;
                    }
                    else
                    {
                        Regex re = new Regex("^(?:[\\w]\\:|\\\\)(\\\\[a-z_\\-\\s0-9\\.]+)+\\.(png|pdf|jpg|jpeg)$\r\n");

                        if (!re.IsMatch(serviceCallData.PurchaseLocationPINCode))
                        {
                            headerRemark.AppendLine("Invalid valid File Name");
                            flag = false;
                        }
                    }

                    if (_addressGuid == Guid.Empty)
                    {
                        if (string.IsNullOrEmpty(serviceCallData.Address))
                        {
                            if (string.IsNullOrEmpty(serviceCallData.Address1))
                            {
                                headerRemark.AppendLine("Address:Address Line is required.");
                                flag = false;
                            }
                            //headerRemark.AppendLine("Address:Address Line is required.");
                            //flag = false;
                        }
                        if (serviceCallData.PINCode.Trim().Length == 0 || serviceCallData.PINCode == null)
                        {
                            headerRemark.AppendLine("Address:PIN Code is required.");
                            flag = false;
                        }
                        else
                        {
                            query = new QueryExpression("hil_pincode");
                            query.ColumnSet = new ColumnSet("hil_pincodeid");
                            query.Criteria = new FilterExpression(LogicalOperator.And);
                            query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, serviceCallData.PINCode);
                            entcoll = service.RetrieveMultiple(query);
                            if (entcoll.Entities.Count == 0)
                            {
                                headerRemark.AppendLine("Address:PIN Code is not found in D365 MDM.");
                                flag = false;
                            }
                        }
                        if (serviceCallData.Source == "19")
                        {
                            query = new QueryExpression("hil_businessmapping");
                            query.ColumnSet = new ColumnSet("hil_area");
                            query.Criteria = new FilterExpression(LogicalOperator.And);
                            query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 1);
                            query.Criteria.AddCondition("hil_pincodename", ConditionOperator.Like, serviceCallData.PINCode);
                            entcoll = service.RetrieveMultiple(query);
                            if (entcoll.Entities.Count > 0)
                            {
                                serviceCallData.AreaCode = service.Retrieve("hil_area", entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_area").Id, new ColumnSet("hil_areacode")).GetAttributeValue<string>("hil_areacode");
                            }
                        }
                        else
                        {
                            if (serviceCallData.AreaCode.Trim().Length == 0 || serviceCallData.AreaCode == null)
                            {
                                headerRemark.AppendLine("Address:Area Code is required.");
                                flag = false;
                            }
                            else
                            {
                                query = new QueryExpression("hil_area");
                                query.ColumnSet = new ColumnSet("hil_areaid");
                                query.Criteria = new FilterExpression(LogicalOperator.And);
                                query.Criteria.AddCondition("hil_areacode", ConditionOperator.Equal, serviceCallData.AreaCode);
                                entcoll = service.RetrieveMultiple(query);
                                if (entcoll.Entities.Count == 0)
                                {
                                    headerRemark.AppendLine("Address:Area Code is not found in D365.");
                                    flag = false;
                                }
                            }
                        }
                    }

                    if (serviceCallData.ServiceCallData.Count == 0)
                    {
                        headerRemark.AppendLine("Model/Serial Number detail is missing");
                        flag = false;
                    }
                    else
                    {
                        int i = 1;
                        foreach (SFA_ServiceCallData objServiceCallLine in serviceCallData.ServiceCallData)
                        {
                            if (_source != 19) // If Source Code is not "MFR"
                            {
                                if (_source == 15)
                                {
                                    objServiceCallLine.Address = serviceCallData.Address;
                                    objServiceCallLine.AddressGuid = serviceCallData.AddressGuid;
                                    objServiceCallLine.AreaCode = serviceCallData.AreaCode;
                                    objServiceCallLine.PINCode = serviceCallData.PINCode;
                                }
                                if (objServiceCallLine.Address == null && serviceCallData.AddressGuid == Guid.Empty && objServiceCallLine.AddressGuid == Guid.Empty)
                                {
                                    headerRemark.Append("Line " + i.ToString() + ": Job Address Full OR Addresss ID is missing,");
                                    flag = false;
                                }
                                if (objServiceCallLine.PINCode == null && serviceCallData.AddressGuid == Guid.Empty && objServiceCallLine.AddressGuid == Guid.Empty)
                                {
                                    headerRemark.Append("Line " + i.ToString() + ": Job Address Pincode is missing,");
                                    flag = false;
                                }
                                if (objServiceCallLine.AreaCode == null && serviceCallData.AddressGuid == Guid.Empty && objServiceCallLine.AddressGuid == Guid.Empty)
                                {
                                    headerRemark.Append("Line " + i.ToString() + ": Job Address Area Code is missing,");
                                    flag = false;
                                }

                                if (serviceCallData.AddressGuid == Guid.Empty && objServiceCallLine.PINCode != null)
                                {
                                    query = new QueryExpression("hil_pincode");
                                    query.ColumnSet = new ColumnSet("hil_pincodeid");
                                    query.Criteria = new FilterExpression(LogicalOperator.And);
                                    query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, objServiceCallLine.PINCode);
                                    entcoll = service.RetrieveMultiple(query);
                                    if (entcoll.Entities.Count == 0)
                                    {
                                        headerRemark.AppendLine("Line " + i.ToString() + ": PIN Code is not found in D365.");
                                        flag = false;
                                    }
                                }
                                if (serviceCallData.AddressGuid == Guid.Empty && objServiceCallLine.AreaCode != null)
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
                            if (objServiceCallLine.ServiceDate == null)
                            {
                                headerRemark.Append("Line " + i.ToString() + ": Srvice Date is missing,");
                                flag = false;
                            }
                            if (objServiceCallLine.ModelNumber == null)
                            {
                                headerRemark.Append("Line " + i.ToString() + ": Model Number is missing,");
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
                                    headerRemark.Append("Line " + i.ToString() + ": Model Number does not exist in D365,");
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
                    if (_addressGuid != Guid.Empty)
                    {
                        enServiceCallRequest["hil_address"] = new EntityReference("hil_address", _addressGuid);
                    }
                    enServiceCallRequest["hil_areacode"] = serviceCallData.AreaCode;
                    if (_customerGuid != Guid.Empty)
                    {
                        enServiceCallRequest["hil_consumer"] = new EntityReference("contact", _customerGuid);
                    }
                    if (serviceCallData.CustomerEmailId != null)
                    {
                        enServiceCallRequest["hil_customeremailid"] = serviceCallData.CustomerEmailId;
                    }
                    if (serviceCallData.CustomerMobileNo != null)
                    {
                        enServiceCallRequest["hil_customermobilenumber"] = serviceCallData.CustomerMobileNo;
                    }
                    if (serviceCallData.CustomerName != null)
                    {
                        enServiceCallRequest["hil_customername"] = serviceCallData.CustomerName;
                    }
                    else
                    {
                        enServiceCallRequest["hil_customername"] = serviceCallData.CustomerFirstName + " " + serviceCallData.CustomerLastName;
                    }
                    if (serviceCallData.Address != null)
                    {
                        enServiceCallRequest["hil_fulladdress"] = serviceCallData.Address;
                    }
                    else
                    {
                        enServiceCallRequest["hil_fulladdress"] = serviceCallData.Address1 + ", " + serviceCallData.Address3 + ", " + serviceCallData.Address3;
                    }
                    if (serviceCallData.PINCode != null)
                    {
                        enServiceCallRequest["hil_pincode"] = serviceCallData.PINCode;
                    }
                    if (serviceCallData.AreaCode != null)
                    {
                        enServiceCallRequest["hil_areacode"] = serviceCallData.AreaCode;
                    }
                    if (serviceCallData.Source != null)
                    {
                        enServiceCallRequest["hil_source"] = new OptionSetValue(Convert.ToInt32(serviceCallData.Source));
                    }
                    else
                    {
                        enServiceCallRequest["hil_source"] = new OptionSetValue(7);
                    }
                    if (serviceCallData.FileName != null)
                    {
                        enServiceCallRequest["hil_filename"] = serviceCallData.AreaCode;
                    }
                    if (serviceCallData.InvoiceDate != null)
                    {
                        enServiceCallRequest["hil_invoicedate"] = Convert.ToDateTime(serviceCallData.InvoiceDate);//YYYY-MM-DD
                    }
                    if (serviceCallData.InvoiceNo != null)
                    {
                        enServiceCallRequest["hil_invoicenumber"] = serviceCallData.InvoiceNo;
                    }
                    if (serviceCallData.InvoiceValue != null)
                    {
                        Regex re = new Regex("");

                        enServiceCallRequest["hil_invoicevalue"] = new Money(Convert.ToDecimal(serviceCallData.InvoiceValue));
                    }
                    else
                    {
                        headerRemark.AppendLine("Please enter valid InvoiceValue");
                    }
                    if (serviceCallData.SalesTrackingID != null)
                    {
                        if (serviceCallData.SalesTrackingID.Length <= 20)
                        {
                            Regex re = new Regex("^([0-9a-zA-Z]){1,20}$");
                            if (re.IsMatch(serviceCallData.SalesTrackingID))
                            {
                                headerRemark.AppendLine("Please enter a valid Sales Tracking Id");
                            }
                            else
                            {
                                headerRemark.AppendLine("False");
                            }
                            enServiceCallRequest["hil_name"] = serviceCallData.SalesTrackingID;
                        }
                        else
                        {
                            headerRemark.AppendLine(" The length of the SalesTrackingID exceeded the maximum allowed length of 20");
                        }

                    }
                    else
                    {
                        Console.Write("Sales Trackign Id cannot be blank");
                    }
                    if (serviceCallData.PurchaseFrom != null)
                    {
                        enServiceCallRequest["hil_purchasefrom"] = serviceCallData.PurchaseFrom;
                    }
                    else
                    {
                        enServiceCallRequest["hil_purchasefrom"] = serviceCallData.PurchaseFromName + " || " + serviceCallData.PurchaseFromCode;
                    }
                    if (serviceCallData.PurchaseLocationPINCode != null)
                    {
                        enServiceCallRequest["hil_purchaselocation"] = serviceCallData.PurchaseLocationPINCode;
                    }
                    string jsonString = serviceCallData.Attachment;
                    if (serviceCallData.Attachment != null)
                    {
                        enServiceCallRequest["hil_image"] = Convert.FromBase64String(jsonString);
                        serviceCallData.Attachment = string.Empty;
                    }
                    //enServiceCallRequest["hil_requestjson"] = JsonConvert.SerializeObject(serviceCallData);

                    _serviceCallRequestGuid = service.Create(enServiceCallRequest);

                    if (_serviceCallRequestGuid != Guid.Empty)
                    {
                        Entity enServiceCallRequestDetails = null;
                        foreach (SFA_ServiceCallData objServiceCall in serviceCallData.ServiceCallData)
                        {
                            enServiceCallRequestDetails = new Entity("hil_servicecallrequestdetail");

                            if (objServiceCall.AddressGuid != Guid.Empty)
                            {
                                enServiceCallRequestDetails["hil_address"] = new EntityReference("hil_address", objServiceCall.AddressGuid);
                            }
                            if (objServiceCall.Address != null)
                            {
                                enServiceCallRequestDetails["hil_fulladdress"] = objServiceCall.Address;
                            }
                            if (objServiceCall.AreaCode != null)
                            {
                                enServiceCallRequestDetails["hil_areacode"] = objServiceCall.AreaCode;
                            }
                            if (objServiceCall.PINCode != null)
                            {
                                enServiceCallRequestDetails["hil_pincode"] = objServiceCall.PINCode;
                            }
                            if (objServiceCall.Qty == 0)
                            {
                                lineRemark.Append("Quantity must not be zero,");
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
                _responseData = JsonConvert.SerializeObject(returnObj);
                DateTime _endDate = DateTime.Now;
                //APIExceptionLog.LogAPIExecution(string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", "SFA App", "SFACreateServiceCall", _error, _requestData.Replace(",", "~"), _responseData.Replace(",", "~"), _startDate.ToString(), _endDate.ToString(), (_endDate - _startDate).TotalSeconds.ToString() + " sec"));
            }
            return returnObj;
        }
        public static bool TryGetFromBase64String(string input)
        {            
            try
            {
               Convert.FromBase64String(input);
                return true;
            }
            catch (FormatException ex)
            {
                return false;
            }
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

    public class SFA_ServiceCall
    {
        public string SalesTrackingID { get; set; }

        public string InvoiceDate { get; set; }

        public string InvoiceNo { get; set; }

        public decimal InvoiceValue { get; set; }

        public string PurchaseLocationPINCode { get; set; }

        public string PurchaseFromCode { get; set; }

        public string PurchaseFromName { get; set; }

        public string Attachment { get; set; }

        public string FileName { get; set; }

        public string PurchaseFrom { get; set; }

        public string CustomerMobileNo { get; set; }

        public string CustomerName { get; set; }

        public string CustomerFirstName { get; set; }

        public string CustomerLastName { get; set; }

        public string CustomerEmailId { get; set; }

        public Guid CustomerGuid { get; set; }

        public Guid AddressGuid { get; set; }

        public string Address { get; set; }

        public string Address1 { get; set; }

        public string Address2 { get; set; }

        public string Address3 { get; set; }

        public string PINCode { get; set; }

        public string AreaCode { get; set; }

        public string Source { get; set; }

        public bool ResultStatus { get; set; }

        public string ResultMessage { get; set; }

        public List<SFA_ServiceCallData> ServiceCallData { get; set; }

    }
}
