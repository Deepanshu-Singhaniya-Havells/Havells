using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Text;
using Newtonsoft.Json;

namespace TestApp
{
    
    public class SFA_ServiceCall
    {
        
        public string CustomerMobileNo { get; set; }
        
        public string CustomerName { get; set; }
        
        public string CustomerEmailId { get; set; }
        
        public Guid CustomerGuid { get; set; }
        
        public Guid AddressGuid { get; set; }
        
        public string Address { get; set; }
        
        public string PINCode { get; set; }
        
        public string AreaCode { get; set; }
        
        public string Source { get; set; }
        
        
        public bool ResultStatus { get; set; }
        
        public string ResultMessage { get; set; }
        
        public List<SFA_ServiceCallData> ServiceCallData { get; set; }

        public SFA_ServiceCall() {
            ServiceCallData = new List<SFA_ServiceCallData>();
            ResultStatus = false;
        }

        public SFA_ServiceCallResult SFA_CreateServiceCall(SFA_ServiceCall serviceCallData,IOrganizationService service)
        {
            #region Variable declaration
            int _source = 7;
            SFA_ServiceCallResult returnObj = new SFA_ServiceCallResult();
            returnObj.ResultMessage = "SUCCESS";
            returnObj.ResultStatus = true;
            returnObj.ServiceCallLogData = new List<SFA_ServiceCallLogData>();
            StringBuilder headerRemark =new StringBuilder();
            StringBuilder lineRemark = new StringBuilder();

            Guid _customerGuid = serviceCallData.CustomerGuid;
            //headerRemark.AppendLine("Cust :" + _customerGuid.ToString());
            Guid _addressGuid = serviceCallData.AddressGuid;
            //headerRemark.AppendLine("Add :" + _addressGuid.ToString());
            if (serviceCallData.Source != null)
            {
                _source = Convert.ToInt32(serviceCallData.Source);
            }

            Guid _serviceCallRequestGuid;
            EntityCollection entcoll;
            QueryExpression query;
            //IOrganizationService service = null;
            EntityReference erPINCode = null;
            EntityReference erArea = null;
            EntityReference erBusinessGeo = null;
            int addressType = 1;
            string customerFullName = serviceCallData.CustomerName;
            string customerMobileNumber = serviceCallData.CustomerMobileNo;
            string customerEmail = serviceCallData.CustomerEmailId;
            EntityReference erProductCategory = null;
            EntityReference erProductsubcategory = null;
            EntityReference erConsumertype = new EntityReference("hil_consumertype", new Guid("484897de-2abd-e911-a957-000d3af0677f")); //B2C
            EntityReference erConsumercategory = new EntityReference("hil_consumercategory", new Guid("baa52f5e-2bbd-e911-a957-000d3af0677f")); //End User
            EntityReference erNatureOfComplaint = null;
            EntityReference erCallSubTypeDemo = new EntityReference("hil_callsubtype", new Guid("ae1b2b71-3c0b-e911-a94e-000d3af06cd4")); // Demo
            EntityReference erCallSubTypeInstallation = new EntityReference("hil_callsubtype", new Guid("e3129d79-3c0b-e911-a94e-000d3af06cd4")); //Installation
            EntityReference erCallSubTypeBoth = new EntityReference("hil_callsubtype", new Guid("13254938-a783-ea11-a811-000d3af055b6")); //Both
            bool flag = true;
            
            #endregion

            try
            {
                //service = ConnectToCRM.GetOrgServiceQA();
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
                        if (serviceCallData.CustomerName.Trim().Length == 0 || serviceCallData.CustomerName == null)
                        {
                            headerRemark.AppendLine("Customer Name is required.");
                            flag = false;
                        }
                    }

                    if (_addressGuid == Guid.Empty)
                    {
                        if (serviceCallData.Address.Trim().Length == 0 || serviceCallData.Address == null)
                        {
                            headerRemark.AppendLine("Address:Address Line is required.");
                            flag = false;
                        }
                        if (serviceCallData.PINCode.Trim().Length == 0 || serviceCallData.PINCode == null)
                        {
                            headerRemark.AppendLine("Address:PIN Code is required.");
                            flag = false;
                        }
                        #region Check for PIN Code
                        query = new QueryExpression("hil_pincode");
                        query.ColumnSet = new ColumnSet("hil_pincodeid");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, serviceCallData.PINCode);
                        entcoll = service.RetrieveMultiple(query);
                        if (entcoll.Entities.Count == 0)
                        {
                            headerRemark.AppendLine("Address:PIN Code is not found in D365.");
                            flag = false;
                        }
                        else
                        {
                            erPINCode = entcoll.Entities[0].ToEntityReference();
                        }

                        #endregion
                        if (serviceCallData.AreaCode == null)
                        {
                            //headerRemark.AppendLine("Address:Area Code is required.");
                            //flag = false;
                            LinkEntity lnkEntArea = new LinkEntity
                            {
                                LinkFromEntityName = "hil_businessmapping",
                                LinkToEntityName = hil_area.EntityLogicalName,
                                LinkFromAttributeName = "hil_area",
                                LinkToAttributeName = "hil_areaid",
                                Columns = new ColumnSet("hil_areacode"),
                                EntityAlias = "area",
                                JoinOperator = JoinOperator.LeftOuter
                            };
                            query = new QueryExpression("hil_businessmapping");
                            query.ColumnSet = new ColumnSet("hil_area");
                            query.Criteria = new FilterExpression(LogicalOperator.And);
                            query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, erPINCode.Id);
                            query.LinkEntities.Add(lnkEntArea);
                            query.AddOrder("createdon", OrderType.Ascending);
                            query.TopCount = 1;
                            entcoll = service.RetrieveMultiple(query);
                            if (entcoll.Entities.Count > 0)
                            {
                                if (entcoll.Entities[0].Attributes.Contains("hil_area"))
                                {
                                    erArea = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_area");
                                }
                                if (entcoll.Entities[0].Attributes.Contains("area.hil_areacode"))
                                {
                                    serviceCallData.AreaCode = entcoll.Entities[0].GetAttributeValue<AliasedValue>("area.hil_areacode").Value.ToString();
                                }
                            }
                            else
                            {
                                headerRemark.AppendLine("Address: No Area Code is defined with PIN Code you entered.");
                                flag = false;
                            }
                        }
                        else {
                            #region Check for Area Code
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
                            else
                            {
                                erArea = entcoll.Entities[0].ToEntityReference();
                            }
                            #endregion
                        }

                        //headerRemark.Append("Service: " + service.ToString());
                        #region Check for Business Geo Mapping
                        query = new QueryExpression("hil_businessmapping");
                        query.ColumnSet = new ColumnSet("hil_pincode");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, erPINCode.Id);
                        query.Criteria.AddCondition("hil_area", ConditionOperator.Equal, erArea.Id);
                        query.AddOrder("createdon", OrderType.Ascending);
                        query.TopCount = 1;
                        entcoll = service.RetrieveMultiple(query);
                        if (entcoll.Entities.Count == 0)
                        {
                            headerRemark.AppendLine("Address:Business Geo Mapping is not found in D365.");
                            flag = false;
                        }
                        else
                        {
                            erBusinessGeo = entcoll.Entities[0].ToEntityReference();
                        }
                        #endregion
                    }

                    if (serviceCallData.ServiceCallData.Count == 0)
                    {
                        headerRemark.AppendLine("Model/CallsubType detail is missing");
                        flag = false;
                    }
                    else
                    {
                        int i = 1;
                        foreach (SFA_ServiceCallData objServiceCall in serviceCallData.ServiceCallData)
                        {
                            if (_source == 15) {
                                objServiceCall.Address = serviceCallData.Address;
                                objServiceCall.AddressGuid = serviceCallData.AddressGuid;
                                objServiceCall.AreaCode = serviceCallData.AreaCode;
                                objServiceCall.PINCode = serviceCallData.PINCode;
                            }
                            if (objServiceCall.Address == null && objServiceCall.AddressGuid == Guid.Empty)
                            {
                                headerRemark.Append("Line " + i.ToString() + ": Job Address Full OR Addresss ID is missing,");
                                flag = false;
                            }
                            if (objServiceCall.PINCode == null && objServiceCall.AddressGuid == Guid.Empty)
                            {
                                headerRemark.Append("Line " + i.ToString() + ": Job Address Pincode is missing,");
                                flag = false;
                            }
                            if (objServiceCall.AreaCode == null && objServiceCall.AddressGuid == Guid.Empty)
                            {
                                headerRemark.Append("Line " + i.ToString() + ": Job Address Area Code is missing,");
                                flag = false;
                            }
                            if (objServiceCall.ServiceDate == null)
                            {
                                headerRemark.Append("Line " + i.ToString() + ": Srvice Date is missing,");
                                flag = false;
                            }
                            if (objServiceCall.AddressGuid == Guid.Empty && objServiceCall.PINCode != null)
                            {
                                query = new QueryExpression("hil_pincode");
                                query.ColumnSet = new ColumnSet("hil_pincodeid");
                                query.Criteria = new FilterExpression(LogicalOperator.And);
                                query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, objServiceCall.PINCode);
                                entcoll = service.RetrieveMultiple(query);
                                if (entcoll.Entities.Count == 0)
                                {
                                    headerRemark.AppendLine("Line " + i.ToString() + ": PIN Code is not found in D365.");
                                    flag = false;
                                }
                                else {
                                    objServiceCall.PinCodeGuid = entcoll.Entities[0].Id;
                                }
                            }
                            if (objServiceCall.AddressGuid == Guid.Empty && objServiceCall.AreaCode != null) 
                            {
                                query = new QueryExpression("hil_area");
                                query.ColumnSet = new ColumnSet("hil_areaid");
                                query.Criteria = new FilterExpression(LogicalOperator.And);
                                query.Criteria.AddCondition("hil_areacode", ConditionOperator.Equal, objServiceCall.AreaCode);
                                entcoll = service.RetrieveMultiple(query);
                                if (entcoll.Entities.Count == 0)
                                {
                                    headerRemark.AppendLine("Line " + i.ToString() + ": Area Code is not found in D365.");
                                    flag = false;
                                }
                                else {
                                    objServiceCall.AreaCodeGuid = entcoll.Entities[0].Id;
                                }
                            }
                            if (objServiceCall.AddressGuid == Guid.Empty && objServiceCall.AreaCodeGuid != Guid.Empty && objServiceCall.PinCodeGuid != Guid.Empty)
                            {
                                query = new QueryExpression("hil_businessmapping");
                                query.ColumnSet = new ColumnSet("hil_pincode");
                                query.Criteria = new FilterExpression(LogicalOperator.And);
                                query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, objServiceCall.PinCodeGuid);
                                query.Criteria.AddCondition("hil_area", ConditionOperator.Equal, objServiceCall.AreaCodeGuid);
                                query.AddOrder("createdon", OrderType.Ascending);
                                query.TopCount = 1;
                                entcoll = service.RetrieveMultiple(query);
                                if (entcoll.Entities.Count == 0)
                                {
                                    headerRemark.AppendLine("Address:Business Geo Mapping is not found in D365.");
                                    flag = false;
                                }
                                else
                                {
                                    objServiceCall.BizGeoGuid = entcoll.Entities[0].Id;
                                }
                            }
                            i += 1;
                        }
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
                    if (serviceCallData.Address != null)
                    {
                        enServiceCallRequest["hil_fulladdress"] = serviceCallData.Address;
                    }
                    if (headerRemark != null)
                    {
                        enServiceCallRequest["hil_remarks"] = headerRemark.ToString();
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
                    else {
                        enServiceCallRequest["hil_source"] = new OptionSetValue(7);
                    }

                    string jsonString = JsonConvert.SerializeObject(serviceCallData);
                    enServiceCallRequest["hil_requestjson"] = jsonString;

                    //hil_source
                    _serviceCallRequestGuid = service.Create(enServiceCallRequest);

                    if (_serviceCallRequestGuid != Guid.Empty)
                    {
                        foreach (SFA_ServiceCallData objServiceCall in serviceCallData.ServiceCallData)
                        {
                            lineRemark = new StringBuilder();

                            Entity enServiceCallRequestDetails = new Entity("hil_servicecallrequestdetail");
                            
                            if (objServiceCall.AddressGuid != Guid.Empty)
                            {
                                enServiceCallRequest["hil_address"] = new EntityReference("hil_address", objServiceCall.AddressGuid);
                            }
                            if (objServiceCall.Address != null)
                            {
                                enServiceCallRequest["hil_fulladdress"] = objServiceCall.Address;
                            }
                            if (objServiceCall.AreaCode != null)
                            {
                                enServiceCallRequest["hil_areacode"] = objServiceCall.AreaCode;
                            }
                            if (objServiceCall.PINCode  != null)
                            {
                                enServiceCallRequest["hil_pincode"] = objServiceCall.PINCode;
                            }
                            if (objServiceCall.Qty == 0)
                            {
                                lineRemark.Append("Quantity must not be zero,");
                            }
                            if (objServiceCall.CallType == "D")
                            {
                                enServiceCallRequestDetails["hil_callsubtype"] = erCallSubTypeDemo;
                            }
                            else if (objServiceCall.CallType == "I")
                            {
                                enServiceCallRequestDetails["hil_callsubtype"] = erCallSubTypeInstallation;
                            }
                            else if (objServiceCall.CallType == "B")
                            {
                                enServiceCallRequestDetails["hil_callsubtype"] = erCallSubTypeBoth;
                            }
                            else
                            {
                                lineRemark.Append("Invalid Call Sub Type,");
                                flag = false;
                            }

                            if (objServiceCall.ModelNumber == null || objServiceCall.ModelNumber.Trim().Length == 0)
                            {
                                lineRemark.Append("Product Model Number is required,");
                                flag = false;
                            }
                            else
                            {
                                enServiceCallRequestDetails["hil_name"] = objServiceCall.ModelNumber;
                                query = new QueryExpression("product");
                                query.ColumnSet = new ColumnSet("description");
                                query.Criteria = new FilterExpression(LogicalOperator.And);
                                query.Criteria.AddCondition("productnumber", ConditionOperator.Equal, objServiceCall.ModelNumber);
                                EntityCollection ec1 = service.RetrieveMultiple(query);
                                if (ec1.Entities.Count == 0)
                                {
                                    lineRemark.Append("Model Number does not exist in D365,");
                                    flag = false;
                                }
                                else
                                {
                                    enServiceCallRequestDetails["hil_product"] = ec1.Entities[0].ToEntityReference();
                                }
                            }
                            enServiceCallRequestDetails["hil_productremarks"] = lineRemark.ToString();
                            enServiceCallRequestDetails["hil_qty"] = objServiceCall.Qty;
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

                #region Data Validation
                if (!flag)
                {
                    return new SFA_ServiceCallResult { ResultStatus = false, ResultMessage = headerRemark.ToString(), ServiceCallLogData = new List<SFA_ServiceCallLogData>() };
                }
                #endregion

                if (service != null)
                {
                    #region Register new Consumer
                    if (_customerGuid == Guid.Empty)
                    {
                        query = new QueryExpression("contact");
                        query.ColumnSet = new ColumnSet("contactid", "fullname", "emailaddress1", "mobilephone");
                        ConditionExpression condExp = new ConditionExpression("mobilephone", ConditionOperator.Equal, serviceCallData.CustomerMobileNo);
                        query.Criteria.AddCondition(condExp);
                        EntityCollection entColConsumer = service.RetrieveMultiple(query);
                        if (entColConsumer.Entities.Count > 0)
                        {
                            customerMobileNumber = entColConsumer.Entities[0].GetAttributeValue<string>("mobilephone");
                            customerFullName = entColConsumer.Entities[0].GetAttributeValue<string>("fullname");
                            customerEmail = entColConsumer.Entities[0].GetAttributeValue<string>("emailaddress1");
                            _customerGuid = entColConsumer.Entities[0].Id;
                        }
                        else
                        {
                            Entity entConsumer = new Entity("contact");
                            entConsumer["mobilephone"] = serviceCallData.CustomerMobileNo;
                            entConsumer["hil_salutation"] = new OptionSetValue(6);
                            string[] consumerName = serviceCallData.CustomerName.Split(' ');
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
                                entConsumer["firstname"] = serviceCallData.CustomerName;
                            }

                            if (serviceCallData.CustomerEmailId != null && serviceCallData.CustomerEmailId.Trim().Length > 0)
                            {
                                entConsumer["emailaddress1"] = serviceCallData.CustomerEmailId;
                            }
                            entConsumer["hil_consumersource"] = new OptionSetValue(6); // SFA Mobile App
                            _customerGuid = service.Create(entConsumer);
                        }
                    }
                    else
                    {
                        query = new QueryExpression("contact");
                        query.ColumnSet = new ColumnSet("contactid", "fullname", "emailaddress1", "mobilephone");
                        ConditionExpression condExp = new ConditionExpression("contactid", ConditionOperator.Equal, _customerGuid);
                        query.Criteria.AddCondition(condExp);
                        EntityCollection entColConsumer = service.RetrieveMultiple(query);
                        if (entColConsumer.Entities.Count > 0)
                        {
                            customerMobileNumber = entColConsumer.Entities[0].GetAttributeValue<string>("mobilephone");
                            customerFullName = entColConsumer.Entities[0].GetAttributeValue<string>("fullname");
                            customerEmail = entColConsumer.Entities[0].GetAttributeValue<string>("emailaddress1");
                        }
                    }
                    #endregion

                    #region Register Consumer Address
                    if (_addressGuid == Guid.Empty && _customerGuid != Guid.Empty)
                    {
                        query = new QueryExpression("hil_address");
                        query.ColumnSet = new ColumnSet("hil_addressid");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, _customerGuid);
                        entcoll = service.RetrieveMultiple(query);
                        if (entcoll.Entities.Count > 0) { addressType = 2; }

                        hil_address entObj = new hil_address();
                        entObj.hil_Street1 = serviceCallData.Address;

                        entObj.hil_Customer = new EntityReference("contact", _customerGuid);
                        if (erBusinessGeo != null)
                        {
                            entObj.hil_BusinessGeo = erBusinessGeo;
                            entObj.hil_District = service.Retrieve("hil_businessmapping", erBusinessGeo.Id, new ColumnSet("hil_district")).GetAttributeValue<EntityReference>("hil_district");
                        }
                        entObj.hil_AddressType = new OptionSetValue(addressType);
                        _addressGuid = service.Create(entObj);
                    }
                    #endregion

                    #region Update Customer and Address on Service Call Request 
                    if (_addressGuid == Guid.Empty || _customerGuid == Guid.Empty)
                    {
                        Entity entTemp = new Entity("hil_servicecallrequest");
                        entTemp.Id = _serviceCallRequestGuid;
                        if (serviceCallData.CustomerGuid == Guid.Empty)
                        {
                            entTemp["hil_consumer"] = new EntityReference("contact", _customerGuid);
                        }
                        if (serviceCallData.AddressGuid == Guid.Empty)
                        {
                            entTemp["hil_address"] = new EntityReference("hil_address", _addressGuid);
                        }
                        service.Update(entTemp);
                    }
                    #endregion

                    #region Create Service Call
                    returnObj = new SFA_ServiceCallResult();
                    returnObj.ResultStatus = false;
                    Guid _JobId = Guid.Empty;
                    returnObj.ServiceCallLogData = new List<SFA_ServiceCallLogData>();
                    foreach (SFA_ServiceCallData objServiceCall in serviceCallData.ServiceCallData)
                    {
                        if (objServiceCall.AddressGuid == Guid.Empty && _customerGuid != Guid.Empty)
                        {
                            query = new QueryExpression("hil_address");
                            query.ColumnSet = new ColumnSet("hil_addressid");
                            query.Criteria = new FilterExpression(LogicalOperator.And);
                            query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, _customerGuid);
                            entcoll = service.RetrieveMultiple(query);
                            if (entcoll.Entities.Count > 0) { addressType = 2; }

                            hil_address entObj = new hil_address();
                            entObj.hil_Street1 = serviceCallData.Address;

                            entObj.hil_Customer = new EntityReference("contact", _customerGuid);
                            if (objServiceCall.BizGeoGuid != Guid.Empty)
                            {
                                entObj.hil_BusinessGeo = new EntityReference("hil_businessmapping", objServiceCall.BizGeoGuid);
                                entObj.hil_District = service.Retrieve("hil_businessmapping", objServiceCall.BizGeoGuid, new ColumnSet("hil_district")).GetAttributeValue<EntityReference>("hil_district");
                            }
                            entObj.hil_AddressType = new OptionSetValue(addressType);
                            _addressGuid = service.Create(entObj);
                        }
                        else
                        {
                            _addressGuid = objServiceCall.AddressGuid;
                        }

                        string[] callSubTypeArr;
                        var s = new List<String>();
                        int maxQty = 1;
                        int ServQty = objServiceCall.Qty;

                        if (objServiceCall.CallType == "D" || objServiceCall.CallType == "B")
                        {
                            s.Add("Demo");
                        }
                        if (objServiceCall.CallType == "I" || objServiceCall.CallType == "B")
                        {
                            s.Add("Installation");
                        }
                        callSubTypeArr = s.ToArray();

                        query = new QueryExpression("product");
                        query.ColumnSet = new ColumnSet("hil_division", "hil_materialgroup", "description");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("productnumber", ConditionOperator.Equal, objServiceCall.ModelNumber);
                        EntityCollection ec = service.RetrieveMultiple(query);
                        if (ec.Entities.Count > 0)
                        {
                            erProductCategory = ec.Entities[0].GetAttributeValue<EntityReference>("hil_division");
                            erProductsubcategory = ec.Entities[0].GetAttributeValue<EntityReference>("hil_materialgroup");
                            Entity prod_Category = (Entity)service.Retrieve("product", erProductsubcategory.Id, new ColumnSet(new string[] { "hil_maximumthreshold" }));
                            if (prod_Category.Attributes.Contains("hil_maximumthreshold"))
                            {
                                maxQty = prod_Category.GetAttributeValue<int>("hil_maximumthreshold");
                            }
                        }
                        if (ServQty <= maxQty)
                        {
                            maxQty = ServQty;
                        }
                        if (objServiceCall.Qty > 0 && erProductCategory != null && erProductsubcategory != null)
                        {
                            while (ServQty >= maxQty)
                            {
                                foreach (var _callSubType in callSubTypeArr)
                                {
                                    SFA_ServiceCallLogData _serviceCallLogData = new SFA_ServiceCallLogData();
                                    EntityReference erCallSubType = null;
                                    try
                                    {
                                        _serviceCallLogData.CallType = _callSubType;
                                        _serviceCallLogData.ModelNumber = objServiceCall.ModelNumber;

                                        Entity enWorkorder = new Entity("msdyn_workorder");
                                        enWorkorder["hil_customerref"] = new EntityReference("contact", _customerGuid);
                                        enWorkorder["hil_customername"] = customerFullName;
                                        enWorkorder["hil_mobilenumber"] = customerMobileNumber;
                                        enWorkorder["hil_email"] = customerEmail;
                                        enWorkorder["hil_address"] = new EntityReference("hil_address", _addressGuid);
                                        enWorkorder["hil_modelname"] = objServiceCall.ModelNumber;

                                        enWorkorder["hil_preferreddate"] = Convert.ToDateTime(objServiceCall.ServiceDate);

                                        if (erProductCategory != null)
                                        {
                                            enWorkorder["hil_productcategory"] = erProductCategory;
                                        }
                                        if (erProductsubcategory != null)
                                        {
                                            enWorkorder["hil_productsubcategory"] = erProductsubcategory;
                                        }

                                        query = new QueryExpression("hil_stagingdivisonmaterialgroupmapping");
                                        query.ColumnSet = new ColumnSet("hil_stagingdivisonmaterialgroupmappingid");
                                        query.Criteria = new FilterExpression(LogicalOperator.And);
                                        query.Criteria.AddCondition("hil_productcategorydivision", ConditionOperator.Equal, erProductCategory.Id);
                                        query.Criteria = new FilterExpression(LogicalOperator.And);
                                        query.Criteria.AddCondition("hil_productsubcategorymg", ConditionOperator.Equal, erProductsubcategory.Id);
                                        ec = service.RetrieveMultiple(query);
                                        if (ec.Entities.Count > 0)
                                        {
                                            enWorkorder["hil_productcatsubcatmapping"] = ec.Entities[0].ToEntityReference();
                                        }

                                        enWorkorder["hil_consumertype"] = erConsumertype;
                                        enWorkorder["hil_consumercategory"] = erConsumercategory;

                                        if (_callSubType == "Demo")
                                        {
                                            erCallSubType = erCallSubTypeDemo;
                                        }
                                        else if (_callSubType == "Installation")
                                        {
                                            erCallSubType = erCallSubTypeInstallation;
                                        }
                                        enWorkorder["hil_callsubtype"] = erCallSubType;

                                        query = new QueryExpression("hil_natureofcomplaint");
                                        query.ColumnSet = new ColumnSet("hil_natureofcomplaintid");
                                        query.Criteria = new FilterExpression(LogicalOperator.And);
                                        query.Criteria.AddCondition("hil_callsubtype", ConditionOperator.Equal, erCallSubType.Id);
                                        query.Criteria = new FilterExpression(LogicalOperator.And);
                                        query.Criteria.AddCondition("hil_relatedproduct", ConditionOperator.Equal, erProductsubcategory.Id);
                                        ec = service.RetrieveMultiple(query);
                                        if (ec.Entities.Count > 0)
                                        {
                                            erNatureOfComplaint = ec.Entities[0].ToEntityReference();
                                        }

                                        if (erNatureOfComplaint != null)
                                        {
                                            enWorkorder["hil_natureofcomplaint"] = erNatureOfComplaint;

                                            enWorkorder["hil_quantity"] = maxQty;
                                            enWorkorder["hil_callertype"] = new OptionSetValue(910590001);// {SourceofJob:"Customer"}
                                            enWorkorder["hil_sourceofjob"] = new OptionSetValue(_source); // SourceofJob:[{"7": "SFA"},{"15": "eCommerce"}]
                                            enWorkorder["msdyn_serviceaccount"] = new EntityReference("account", new Guid("b8168e04-7d0a-e911-a94f-000d3af00f43")); // {ServiceAccount:"Dummy Customer"}
                                            enWorkorder["msdyn_billingaccount"] = new EntityReference("account", new Guid("b8168e04-7d0a-e911-a94f-000d3af00f43")); // {BillingAccount:"Dummy Customer"}
                                            _JobId = service.Create(enWorkorder);

                                            _serviceCallLogData.ServiceCallGuid = _JobId;
                                            _serviceCallLogData.ServiceCallNo = service.Retrieve(msdyn_workorder.EntityLogicalName, _JobId, new ColumnSet("msdyn_name")).GetAttributeValue<string>("msdyn_name");
                                            returnObj.ResultStatus = true;
                                            returnObj.ResultMessage = "SUCCESS";
                                        }
                                        else {
                                            _serviceCallLogData.ServiceCallNo = "Nature of Complaint is not defined for this Product Sub Category.";
                                        }
                                        returnObj.ServiceCallLogData.Add(_serviceCallLogData);
                                    }
                                    catch (Exception ex)
                                    {
                                        _serviceCallLogData = new SFA_ServiceCallLogData() { CallType = _callSubType, ModelNumber = objServiceCall.ModelNumber, ServiceCallNo = ex.Message };
                                        returnObj.ServiceCallLogData.Add(_serviceCallLogData);
                                    }
                                }
                                ServQty = ServQty - maxQty;
                            }
                        }
                    }
                    return returnObj;
                    #endregion
                }
                else
                {
                    returnObj = new SFA_ServiceCallResult { ResultStatus = false, ResultMessage = "D365 Service Unavailable", ServiceCallLogData = new List<SFA_ServiceCallLogData>() };
                }
            }
            catch (Exception ex)
            {
                returnObj = new SFA_ServiceCallResult { ResultStatus = false, ResultMessage = "D365 Internal Server Error : " + ex.Message, ServiceCallLogData = new List<SFA_ServiceCallLogData>() };
            }
            return returnObj;
        }
    }
    public class SFA_ServiceCallData
    {

        public Guid AddressGuid { get; set; }

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
