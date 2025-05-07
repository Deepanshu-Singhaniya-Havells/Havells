using Havells_Plugin.AccountEn;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Net.Mime;
using System.Net.NetworkInformation;
using System.Runtime.Serialization;
using System.Text;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class MRNEntry
    {
        [DataMember]
        public string CustomerSearchString { get; set; }
        [DataMember]
        public Guid Technician { get; set; }
        [DataMember]
        public string PdiNumber { get; set; }
        [DataMember]
        public Guid CustomerGuId { get; set; }
        [DataMember]
        public string SerialNumber { get; set; }
        [DataMember]
        public string ProductCode { get; set; }
        [DataMember]
        public string DIVISION { get; set; }
        [DataMember]
        public int? BRAND { get; set; }
        [DataMember]
        public string InvoiceDate { get; set; }
        [DataMember]
        public string InvoiceNumber { get; set; }
        [DataMember]
        public bool? AssetStatus { get; set; } //{True:Success, False: Failed} 
        [DataMember]
        public string Latitude { get; set; }
        [DataMember]
        public string Longitude { get; set; }
        [DataMember]
        public string TempReferenceNumber { get; set; }
        [DataMember]
        public string Is_Type { get; set; }
        [DataMember]
        public bool ResultStatus { get; set; }
        [DataMember]
        public string ResultMessage { get; set; }

        public List<CustomerInfo> GetCustomerDetails(string CustomerSearchString)
        {
            IOrganizationService service = null;
            List<CustomerInfo> _retObj = new List<CustomerInfo>();
            try
            {
                if (CustomerSearchString == string.Empty || CustomerSearchString.Trim().Length == 0)
                {
                    _retObj.Add(new CustomerInfo() { ResultStatus = false, ResultMessage = "Search String is required." });
                    return _retObj;
                }
                service = ConnectToCRM.GetOrgServiceQA();
                if (service != null)
                {
                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='account'>
                    <attribute name='name' />
                    <attribute name='accountnumber' />
                    <attribute name='accountid' />
                    <order attribute='name' descending='false' />
                    <filter type='and'>
                        <filter type='or'>
                        <condition attribute='accountnumber' operator='like' value='%" + CustomerSearchString + @"%' />
                        <condition attribute='name' operator='like' value='%" + CustomerSearchString + @"%' />
                        </filter>
                        <condition attribute='customertypecode' operator='in'>
                            <value>1</value>
                            <value>2</value>
                        </condition>
                    </filter>
                    </entity>
                    </fetch>";
                    EntityCollection entColConsumer = service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (entColConsumer.Entities.Count > 0)
                    {
                        foreach (Entity ent in entColConsumer.Entities)
                        {
                            if (ent.Attributes.Contains("accountnumber") && ent.Attributes.Contains("name"))
                            {
                                _retObj.Add(new CustomerInfo() { CustomerCode = ent.GetAttributeValue<string>("accountnumber"), CustomerName = ent.GetAttributeValue<string>("name"), CustomerGuId = ent.Id, ResultStatus = true, ResultMessage = "OK" });
                            }
                        }
                    }
                    else
                    {
                        _retObj.Add(new CustomerInfo() { CustomerCode = "", CustomerName = "", CustomerGuId = new Guid(), ResultStatus = true, ResultMessage = "No Record found." });
                    }
                    return _retObj;
                }
                else
                {
                    _retObj.Add(new CustomerInfo() { ResultStatus = false, ResultMessage = "D365 Service Unavailable" });
                    return _retObj;
                }
            }
            catch (Exception ex)
            {
                _retObj.Add(new CustomerInfo() { ResultStatus = false, ResultMessage = ex.Message });
                return _retObj;
            }
        }

        public MRNHeader AddtoViewList(MRNEntry _productInfo)
        {
            IOrganizationService service = null;
            MRNHeader _retObj = null;
            EntityReference erProduct = null;
            AssignRequest assign = null;
            QueryExpression qrExp;
            EntityCollection entCol;

            try
            {
                service = ConnectToCRM.GetOrgServiceQA();
                if (service != null)
                {
                    if (_productInfo.Technician == Guid.Empty)
                    {
                        return new MRNHeader() { ResultStatus = false, ResultMessage = "Technician GuId is required." };
                    }
                    else if (_productInfo.CustomerGuId == Guid.Empty)
                    {
                        return new MRNHeader() { ResultStatus = false, ResultMessage = "Customer GuId is required." };
                    }
                    else if (_productInfo.ProductCode == string.Empty || _productInfo.ProductCode.Trim().Length == 0)
                    {
                        return new MRNHeader() { ResultStatus = false, ResultMessage = "Product Code is required." };
                    }
                    else if (_productInfo.DIVISION == string.Empty || _productInfo.DIVISION.Trim().Length == 0)
                    {
                        return new MRNHeader() { ResultStatus = false, ResultMessage = "Division Code is required." };
                    }
                    else if (_productInfo.BRAND == null || _productInfo.BRAND == 0)
                    {
                        return new MRNHeader() { ResultStatus = false, ResultMessage = "Brand is required." };
                    }
                    else if (_productInfo.SerialNumber == string.Empty || _productInfo.SerialNumber.Trim().Length == 0)
                    {
                        return new MRNHeader() { ResultStatus = false, ResultMessage = "Serial Number is required." };
                    }
                    else if (_productInfo.InvoiceNumber == string.Empty || _productInfo.InvoiceNumber.Trim().Length == 0)
                    {
                        return new MRNHeader() { ResultStatus = false, ResultMessage = "Invoice Number is required." };
                    }
                    else if (_productInfo.InvoiceDate == null)
                    {
                        return new MRNHeader() { ResultStatus = false, ResultMessage = "Invoice Date is required." };
                    }
                    else if (_productInfo.AssetStatus == null)
                    {
                        return new MRNHeader() { ResultStatus = false, ResultMessage = "Asset Status is required." };
                    }
                    else if (_productInfo.PdiNumber == null)
                    {
                        return new MRNHeader() { ResultStatus = false, ResultMessage = "PDI Call is required." };
                    }
                    else
                    {
                        qrExp = new QueryExpression("product");
                        qrExp.ColumnSet = new ColumnSet("productid");
                        ConditionExpression condExp = new ConditionExpression("name", ConditionOperator.Equal, _productInfo.ProductCode);
                        qrExp.Criteria.AddCondition(condExp);
                        entCol = service.RetrieveMultiple(qrExp);
                        if (entCol.Entities.Count == 0)
                        {
                            return new MRNHeader() { ResultStatus = false, ResultMessage = "Product Code does not exist." };
                        }
                        else
                        {
                            erProduct = entCol.Entities[0].ToEntityReference();
                        }
                    }


                    DateTime _inVoiceDate = new DateTime(Convert.ToInt16(_productInfo.InvoiceDate.Substring(0, 4)), Convert.ToInt16(_productInfo.InvoiceDate.Substring(4, 2)), Convert.ToInt16(_productInfo.InvoiceDate.Substring(6, 2)));

                    Guid _headerGuId = Guid.Empty;
                    qrExp = new QueryExpression("hil_dealerstockverificationheader");
                    qrExp.ColumnSet = new ColumnSet("hil_dealerstockverificationheaderid");
                    qrExp.Criteria.AddCondition(new ConditionExpression("ownerid", ConditionOperator.Equal, _productInfo.Technician)); //Technician
                    qrExp.Criteria.AddCondition(new ConditionExpression("hil_customercode", ConditionOperator.Equal, _productInfo.CustomerGuId)); //Customer
                    qrExp.Criteria.AddCondition(new ConditionExpression("hil_mrnstatus", ConditionOperator.Equal, 1)); // Draft
                    qrExp.Criteria.AddCondition(new ConditionExpression("createdon", ConditionOperator.Today)); //Today
                    entCol = service.RetrieveMultiple(qrExp);
                    if (entCol.Entities.Count == 0)
                    {
                        Entity ent = new Entity("hil_dealerstockverificationheader");
                        ent["hil_name"] = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') + DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');
                        ent["hil_customercode"] = new EntityReference("account", _productInfo.CustomerGuId);
                        ent["hil_mrnstatus"] = new OptionSetValue(1); //Draft
                        ent["hil_pdinumber"] = _productInfo.PdiNumber;

                        _headerGuId = service.Create(ent);

                        assign = new AssignRequest();
                        assign.Assignee = new EntityReference("systemuser", _productInfo.Technician); //User or team
                        assign.Target = new EntityReference("hil_dealerstockverificationheader", _headerGuId);
                        service.Execute(assign);
                    }
                    else
                    {
                        _headerGuId = entCol.Entities[0].Id;
                        qrExp = new QueryExpression("hil_materialreturn");
                        qrExp.ColumnSet = new ColumnSet("hil_materialreturnid");
                        qrExp.Criteria.AddCondition(new ConditionExpression("hil_dealerstockverificationheader", ConditionOperator.Equal, _headerGuId));
                        qrExp.Criteria.AddCondition(new ConditionExpression("hil_brand", ConditionOperator.NotEqual, _productInfo.BRAND));
                        entCol = service.RetrieveMultiple(qrExp);
                        if (entCol.Entities.Count > 0) {
                            return new MRNHeader() { ResultStatus = false, ResultMessage = "More than One Brand's Product not allowed in one bucket." };
                        }
                    }
                    qrExp = new QueryExpression("hil_materialreturn");
                    qrExp.ColumnSet = new ColumnSet("hil_materialreturnid");
                    qrExp.Criteria.AddCondition(new ConditionExpression("hil_dealerstockverificationheader", ConditionOperator.Equal, _headerGuId));
                    qrExp.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, _productInfo.SerialNumber));
                    entCol = service.RetrieveMultiple(qrExp);
                    if (entCol.Entities.Count == 0)
                    {
                        Entity ent = new Entity("hil_materialreturn");
                        ent["hil_dealerstockverificationheader"] = new EntityReference("hil_dealerstockverificationheader", _headerGuId);
                        ent["hil_invoicedate"] = _inVoiceDate;
                        ent["hil_invoicenumber"] = _productInfo.InvoiceNumber;
                        ent["hil_latitude"] = _productInfo.Latitude;
                        ent["hil_longitude"] = _productInfo.Longitude;
                        ent["hil_name"] = _productInfo.SerialNumber;
                        ent["hil_productcode"] = _productInfo.ProductCode;
                        ent["hil_brand"] = new OptionSetValue(Convert.ToInt16(_productInfo.BRAND));
                        ent["hil_divisioncode"] = _productInfo.DIVISION;

                        ACInOutScanResult result = ValidateACInOutScanSequence(service, _headerGuId, _productInfo);
                        if (result.ResultStatus == false)
                        {
                            return new MRNHeader() { ResultStatus = result.ResultStatus, ResultMessage = result.ResultMessage };
                        }

                        ent["hil_istype"] = _productInfo.Is_Type;

                        if (erProduct != null)
                        {
                            ent["hil_product"] = erProduct;
                        }
                        ent["hil_warrantystatus"] = _productInfo.AssetStatus;
                        Guid _lineGuId = service.Create(ent);

                        assign = new AssignRequest();
                        assign.Assignee = new EntityReference("systemuser", _productInfo.Technician); //User or team
                        assign.Target = new EntityReference("hil_materialreturn", _lineGuId);
                        service.Execute(assign);

                        string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='hil_materialreturn'>
                            <attribute name='hil_materialreturnid' />
                            <attribute name='hil_name' />
                            <attribute name='createdon' />
                            <attribute name='hil_warrantystatus' />
                            <attribute name='ownerid' />
                            <attribute name='hil_dealerstockverificationheader' />
                            <attribute name='hil_productcode' />
                            <attribute name='hil_product' />
                            <attribute name='hil_longitude' />
                            <attribute name='hil_latitude' />
                            <attribute name='hil_brand' />
                            <attribute name='hil_divisioncode' />
                            <attribute name='hil_invoicenumber' />
                            <attribute name='hil_invoicedate' />
                            <order attribute='createdon' descending='false' />
                            <filter type='and'>
                              <condition attribute='hil_dealerstockverificationheader' operator='eq' value='" + _headerGuId.ToString() + @"' />
                            </filter>
                            <link-entity name='hil_dealerstockverificationheader' from='hil_dealerstockverificationheaderid' to='hil_dealerstockverificationheader' link-type='inner' alias='header'>
                              <attribute name='hil_name' />
                              <attribute name='hil_mrnstatus' />
                              <attribute name='hil_customercode' />
                              <link-entity name='account' from='accountid' to='hil_customercode' link-type='inner' alias='cust'>
	                                <attribute name='accountnumber' />
                              </link-entity>
                            </link-entity>
                          </entity>
                        </fetch>";

                        entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entCol.Entities.Count > 0)
                        {
                            EntityReference entRef = entCol.Entities[0].GetAttributeValue<EntityReference>("hil_dealerstockverificationheader");
                            _retObj = new MRNHeader();
                            _retObj.MRNGuId = entRef.Id;
                            _retObj.Technician = entCol.Entities[0].GetAttributeValue<EntityReference>("ownerid").Id;
                            if (entCol.Entities[0].Attributes.Contains("header.hil_mrnstatus"))
                            {
                                OptionSetValue _headerStatus = (OptionSetValue)entCol.Entities[0].GetAttributeValue<AliasedValue>("header.hil_mrnstatus").Value;
                                _retObj.MRNStatus = _headerStatus.Value;

                                if (entCol.Entities[0].FormattedValues.Contains("header.hil_mrnstatus"))
                                    _retObj.MRNStatusName = entCol.Entities[0].FormattedValues["header.hil_mrnstatus"];
                            }
                            if (entCol.Entities[0].Attributes.Contains("cust.accountnumber"))
                            {
                                _retObj.CustomerCode = entCol.Entities[0].GetAttributeValue<AliasedValue>("cust.accountnumber").Value.ToString();
                                _retObj.CustomerName = ((EntityReference)entCol.Entities[0].GetAttributeValue<AliasedValue>("header.hil_customercode").Value).Name;
                                _retObj.CustomerGuId = ((EntityReference)entCol.Entities[0].GetAttributeValue<AliasedValue>("header.hil_customercode").Value).Id;
                            }
                            if (entCol.Entities[0].Attributes.Contains("header.hil_name"))
                            {
                                _retObj.ReferenceNumber = entCol.Entities[0].GetAttributeValue<AliasedValue>("header.hil_name").Value.ToString();
                            }

                            MRNLine _line = null;
                            string _invDate = string.Empty;
                            DateTime dtTemp = DateTime.Now;
                            _retObj.ProductLines = new List<MRNLine>();
                            foreach (Entity entLine in entCol.Entities)
                            {
                                dtTemp = entLine.GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330);
                                _invDate = dtTemp.Year.ToString() + dtTemp.Month.ToString().PadLeft(2, '0') + dtTemp.Day.ToString().PadLeft(2, '0');
                                _line = new MRNLine()
                                {
                                    MRNLineGuId = entLine.Id,
                                    CustomerCode = _retObj.CustomerCode,
                                    CustomerName = _retObj.CustomerName,
                                    SerialNumber = entLine.GetAttributeValue<string>("hil_name"),
                                    ProductCode = entLine.GetAttributeValue<string>("hil_productcode"),
                                    ProductGuId = entLine.GetAttributeValue<EntityReference>("hil_product").Id,
                                    InvoiceDate = _invDate,
                                    InvoiceNumber = entLine.GetAttributeValue<string>("hil_invoicenumber"),
                                    Latitude = entLine.GetAttributeValue<string>("hil_latitude"),
                                    Longitude = entLine.GetAttributeValue<string>("hil_longitude"),
                                    DIVISION = entLine.GetAttributeValue<string>("hil_divisioncode"),
                                    BRAND = entLine.GetAttributeValue<OptionSetValue>("hil_brand").Value,
                                    BRANDNAME = entLine.GetAttributeValue<OptionSetValue>("hil_brand").Value == 1 ? "Havells" : entLine.GetAttributeValue<OptionSetValue>("hil_brand").Value == 2 ? "Lloyd" : "WaterPurifier",
                                    AssetStatus = entLine.GetAttributeValue<bool>("hil_warrantystatus"),
                                    AssetStatusName = entLine.GetAttributeValue<bool>("hil_warrantystatus") ? "Success" : "Failed"
                                };
                                _retObj.ProductLines.Add(_line);
                            }
                            _retObj.ResultStatus = true;
                            _retObj.ResultMessage = "OK";
                        }
                    }
                    else
                    {
                        _retObj = new MRNHeader() { ResultStatus = false, ResultMessage = "Serial Number already exist." };
                    }
                }
                else
                {
                    _retObj = new MRNHeader() { ResultStatus = false, ResultMessage = "D365 Service Unavailable." };
                }
            }
            catch (Exception ex)
            {
                _retObj = new MRNHeader() { ResultStatus = false, ResultMessage = ex.Message };
            }
            return _retObj;
        }
        private ACInOutScanResult ValidateACInOutScanSequence(IOrganizationService service, Guid _headerGuId, MRNEntry _productInfo)
        {
            try
            {
                ACInOutScanResult result = new ACInOutScanResult();
                QueryExpression qrExp;
                EntityCollection entCol;
                qrExp = new QueryExpression("hil_materialreturn");
                qrExp.ColumnSet = new ColumnSet("hil_istype", "hil_productcode", "hil_name");
                qrExp.Criteria.AddCondition(new ConditionExpression("hil_dealerstockverificationheader", ConditionOperator.Equal, _headerGuId));
                qrExp.AddOrder("createdon", OrderType.Descending);
                qrExp.TopCount = 1;
                entCol = service.RetrieveMultiple(qrExp);
                if (entCol.Entities.Count > 0)
                {
                    string IsType = entCol.Entities[0].GetAttributeValue<string>("hil_istype");
                    string ModelCode = entCol.Entities[0].GetAttributeValue<string>("hil_productcode");
                    string SerialNo = entCol.Entities[0].GetAttributeValue<string>("hil_name");

                    if (_productInfo.ProductCode == ModelCode)
                    {
                        if (IsType.ToUpper() == "I" || IsType.ToUpper() == "O")
                        {
                            if (SerialNo == _productInfo.SerialNumber)
                            {
                                return new ACInOutScanResult() { ResultStatus = false, ResultMessage = "Duplicate Scan" };
                            }
                            else if (IsType.ToUpper() == _productInfo.Is_Type.ToUpper())
                            {
                                if (SerialNo == _productInfo.SerialNumber)
                                {
                                    return new ACInOutScanResult() { ResultStatus = true, ResultMessage = "Success" };
                                }
                                else
                                {
                                    return new ACInOutScanResult() { ResultStatus = false, ResultMessage = "Please Scan " + (IsType.ToUpper() == "I" ? "Out door Unit " : "In Door Unit ") + " of Product " + ModelCode };
                                }
                            }
                            else
                            {
                                return new ACInOutScanResult() { ResultStatus = true, ResultMessage = "Success" };
                            }
                        }
                        else
                        {
                            if (SerialNo == _productInfo.SerialNumber)
                            {
                                return new ACInOutScanResult() { ResultStatus = false, ResultMessage = "Duplicate Scan" };
                            }
                            else
                            {
                                return new ACInOutScanResult() { ResultStatus = true, ResultMessage = "Success" };
                            }
                        }
                    }
                    else
                    {
                        if (IsType.ToUpper() == "I" || IsType.ToUpper() == "O")
                        {
                            qrExp = new QueryExpression("hil_materialreturn");
                            qrExp.ColumnSet = new ColumnSet(false);
                            qrExp.Criteria = new FilterExpression(LogicalOperator.And);

                            FilterExpression filter1 = new FilterExpression(LogicalOperator.And);
                            filter1.AddCondition(new ConditionExpression("hil_dealerstockverificationheader", ConditionOperator.Equal, _headerGuId));
                            filter1.AddCondition(new ConditionExpression("hil_productcode", ConditionOperator.Equal, ModelCode));

                            FilterExpression filter2 = new FilterExpression(LogicalOperator.Or);
                            filter2.AddCondition(new ConditionExpression("hil_istype", ConditionOperator.Equal, "I"));
                            filter2.AddCondition(new ConditionExpression("hil_istype", ConditionOperator.Equal, "O"));

                            qrExp.Criteria.AddFilter(filter1);
                            qrExp.Criteria.AddFilter(filter2);
                            entCol = service.RetrieveMultiple(qrExp);
                            if (entCol.Entities.Count == 2)
                            {
                                return new ACInOutScanResult() { ResultStatus = true, ResultMessage = "Success" };
                            }
                            else
                            {
                                if (_productInfo.Is_Type.ToUpper() != "I" || _productInfo.Is_Type.ToUpper() != "O")
                                {
                                    return new ACInOutScanResult() { ResultStatus = true, ResultMessage = "Success" };
                                }
                                else
                                {
                                    return new ACInOutScanResult() { ResultStatus = false, ResultMessage = "Please Scan " + (IsType.ToUpper() == "I" ? "Out door Unit " : "In Door Unit ") + " of Product " + ModelCode };
                                }
                            }
                        }
                        else
                        {
                            qrExp = new QueryExpression("hil_materialreturn");
                            qrExp.ColumnSet = new ColumnSet("hil_istype", "hil_productcode", "hil_name");
                            qrExp.Criteria.AddCondition(new ConditionExpression("hil_dealerstockverificationheader", ConditionOperator.Equal, _headerGuId));
                            qrExp.Criteria.AddCondition(new ConditionExpression("hil_productcode", ConditionOperator.Equal, _productInfo.ProductCode));
                            qrExp.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, _productInfo.SerialNumber));
                            qrExp.Criteria.AddCondition(new ConditionExpression("hil_istype", ConditionOperator.Equal, _productInfo.Is_Type));
                            qrExp.TopCount = 1;
                            entCol = service.RetrieveMultiple(qrExp);
                            if (entCol.Entities.Count > 0)
                            {
                                return new ACInOutScanResult() { ResultStatus = false, ResultMessage = "Duplicate Scan" };
                            }
                            else
                            {
                                return new ACInOutScanResult() { ResultStatus = true, ResultMessage = "Success" };
                            }
                        }
                    }
                }
                else
                {
                    return new ACInOutScanResult() { ResultStatus = true, ResultMessage = "Success" };
                }
            }
            catch (Exception ex)
            {
                return new ACInOutScanResult() { ResultStatus = false, ResultMessage = "ERROR! " + ex.Message };
            }
        }

        public MRNHeader SubmitViewList(MRNHeader _viewList)
        {
            IOrganizationService service = null;
            MRNHeader _retObj = null;
            QueryExpression qrExp;
            EntityCollection entCol;

            try
            {
                if (_viewList.MRNGuId == Guid.Empty)
                {
                    return new MRNHeader() { ResultStatus = false, ResultMessage = "View Bucket GuId is required." };
                }
                //else if (_viewList.MRNStatus > 1)
                //{
                //    return new MRNHeader() { ResultStatus = false, ResultMessage = "Valid Status(1-Draft) is required." };
                //}

                service = ConnectToCRM.GetOrgServiceQA();
                if (service != null)
                {
                    qrExp = new QueryExpression("hil_dealerstockverificationheader");
                    qrExp.ColumnSet = new ColumnSet("hil_dealerstockverificationheaderid", "hil_name");
                    qrExp.Criteria.AddCondition(new ConditionExpression("hil_dealerstockverificationheaderid", ConditionOperator.Equal, _viewList.MRNGuId));
                    qrExp.Criteria.AddCondition(new ConditionExpression("hil_mrnstatus", ConditionOperator.Equal, 1));
                    entCol = service.RetrieveMultiple(qrExp);
                    if (entCol.Entities.Count > 0)
                    {
                        Entity ent = new Entity("hil_dealerstockverificationheader", _viewList.MRNGuId);
                        ent["hil_mrnstatus"] = new OptionSetValue(2);
                        service.Update(ent);
                        _retObj = new MRNHeader() { ResultStatus = true, ResultMessage = "OK", MRNGuId = _viewList.MRNGuId, ReferenceNumber = entCol.Entities[0].GetAttributeValue<string>("hil_name") };
                    }
                    else
                    {
                        _retObj = new MRNHeader() { ResultStatus = false, ResultMessage = "View list already submitted.", MRNGuId = _viewList.MRNGuId };
                    }
                }
                else
                {
                    _retObj = new MRNHeader() { ResultStatus = false, ResultMessage = "D365 Service Unavailable.", MRNGuId = _viewList.MRNGuId };
                }
            }
            catch (Exception ex)
            {
                _retObj = new MRNHeader() { ResultStatus = false, ResultMessage = ex.Message, MRNGuId = _viewList.MRNGuId };
            }
            return _retObj;
        }

        public MRNSummary GetViewList(ViewListSearch _searchCondition)
        {
            IOrganizationService service = null;
            MRNSummary _retObj = null;
            EntityCollection entCol;
            EntityCollection entColLine;
            MRNHeader _header = null;
            MRNLine _line = null;
            try
            {
                service = ConnectToCRM.GetOrgServiceQA();
                if (service != null)
                {
                    if (_searchCondition.TechnicianGuId == Guid.Empty)
                    {
                        return new MRNSummary() { ResultStatus = false, ResultMessage = "Technician GuId is required." };
                    }
                    else if (_searchCondition.CustomerGuId == Guid.Empty)
                    {
                        return new MRNSummary() { ResultStatus = false, ResultMessage = "Customer GuId is required." };
                    }

                    string _fetchLineXML = string.Empty;
                    string _invDate = string.Empty;
                    DateTime dtTemp = DateTime.Now;
                    string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='hil_dealerstockverificationheader'>
                    <attribute name='hil_dealerstockverificationheaderid' />
                    <attribute name='hil_name' />
                    <attribute name='hil_mrnstatus' />
                    <attribute name='hil_customercode' />
                    <attribute name='ownerid' />
                    <attribute name='hil_pdinumber' />
                    <attribute name='createdon' />
                    <order attribute='hil_name' descending='false' />
                    <filter type='and'>";
                    if (_searchCondition.Date != null)
                    {
                        if (_searchCondition.Date.Trim().Length > 0)
                        {
                            _fetchXML += @"<condition attribute='createdon' operator='on' value='" + _searchCondition.Date + @"' />";
                        }
                    }
                    if (_searchCondition.ReferenceNumber != null)
                    {
                        if (_searchCondition.ReferenceNumber.Trim().Length > 0)
                        {
                            _fetchXML += @"<condition attribute='hil_name' operator='eq' value='" + _searchCondition.ReferenceNumber + @"' />";
                        }
                    }
                    _fetchXML += @"<condition attribute='ownerid' operator='eq' value='" + _searchCondition.TechnicianGuId.ToString() + @"' />
                        <condition attribute='hil_customercode' operator='eq' value='" + _searchCondition.CustomerGuId.ToString() + @"' />
                    </filter>
                    <link-entity name='account' from='accountid' to='hil_customercode' link-type='inner' alias='cust'>
                        <attribute name='accountnumber' />
                    </link-entity>
                    </entity>
                    </fetch>";

                    entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entCol.Entities.Count > 0)
                    {
                        EntityReference entRef = entCol.Entities[0].GetAttributeValue<EntityReference>("hil_dealerstockverificationheader");
                        _retObj = new MRNSummary();
                        _retObj.MRNHeader = new List<MRNHeader>();
                        foreach (Entity entHeader in entCol.Entities)
                        {
                            _header = new MRNHeader();
                            _header.MRNGuId = entHeader.Id;
                            _header.Technician = entHeader.GetAttributeValue<EntityReference>("ownerid").Id;
                            _header.PdiNumber = entHeader.GetAttributeValue<string>("hil_pdinumber");

                            dtTemp = entHeader.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                            _invDate = dtTemp.Year.ToString() + dtTemp.Month.ToString().PadLeft(2, '0') + dtTemp.Day.ToString().PadLeft(2, '0');
                            _header.CreatedOn = _invDate;
                            if (entHeader.Attributes.Contains("hil_mrnstatus"))
                            {
                                OptionSetValue _headerStatus = entHeader.GetAttributeValue<OptionSetValue>("hil_mrnstatus");
                                _header.MRNStatus = _headerStatus.Value;

                                if (entHeader.FormattedValues.Contains("hil_mrnstatus"))
                                    _header.MRNStatusName = entHeader.FormattedValues["hil_mrnstatus"];
                            }
                            if (entHeader.Attributes.Contains("cust.accountnumber"))
                            {
                                _header.CustomerCode = entHeader.GetAttributeValue<AliasedValue>("cust.accountnumber").Value.ToString();
                                _header.CustomerName = entHeader.GetAttributeValue<EntityReference>("hil_customercode").Name;
                                _header.CustomerGuId = entHeader.GetAttributeValue<EntityReference>("hil_customercode").Id;
                            }
                            if (entHeader.Attributes.Contains("hil_name"))
                            {
                                _header.ReferenceNumber = entHeader.GetAttributeValue<string>("hil_name").ToString();
                            }
                            _header.ProductLines = new List<MRNLine>();
                            _fetchLineXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_materialreturn'>
                            <attribute name='hil_materialreturnid' />
                            <attribute name='hil_name' />
                            <attribute name='createdon' />
                            <attribute name='hil_warrantystatus' />
                            <attribute name='ownerid' />
                            <attribute name='hil_dealerstockverificationheader' />
                            <attribute name='hil_productcode' />
                            <attribute name='hil_product' />
                            <attribute name='hil_longitude' />
                            <attribute name='hil_latitude' />
                            <attribute name='hil_brand' />
                            <attribute name='hil_divisioncode' />
                            <attribute name='hil_invoicenumber' />
                            <attribute name='hil_invoicedate' />
                            <attribute name='hil_istype' />
                            <order attribute='createdon' descending='false' />
                            <filter type='and'>
                                <condition attribute='hil_dealerstockverificationheader' operator='eq' value='" + entHeader.Id.ToString() + @"' />
                            </filter>
                            </entity>
                            </fetch>";
                            entColLine = service.RetrieveMultiple(new FetchExpression(_fetchLineXML));
                            foreach (Entity entLine in entColLine.Entities)
                            {
                                dtTemp = entLine.GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330);
                                _invDate = dtTemp.Year.ToString() + dtTemp.Month.ToString().PadLeft(2, '0') + dtTemp.Day.ToString().PadLeft(2, '0');
                                _line = new MRNLine()
                                {
                                    MRNLineGuId = entLine.Id,
                                    CustomerCode= _header.CustomerCode,
                                    CustomerName= _header.CustomerName,
                                    SerialNumber = entLine.GetAttributeValue<string>("hil_name"),
                                    ProductCode = entLine.GetAttributeValue<string>("hil_productcode"),
                                    ProductGuId = entLine.GetAttributeValue<EntityReference>("hil_product").Id,
                                    InvoiceDate = _invDate,
                                    InvoiceNumber = entLine.GetAttributeValue<string>("hil_invoicenumber"),
                                    Latitude = entLine.GetAttributeValue<string>("hil_latitude"),
                                    Longitude = entLine.GetAttributeValue<string>("hil_longitude"),
                                    DIVISION = entLine.GetAttributeValue<string>("hil_divisioncode"),
                                    BRAND = entLine.GetAttributeValue<OptionSetValue>("hil_brand").Value,
                                    BRANDNAME = entLine.GetAttributeValue<OptionSetValue>("hil_brand").Value == 1 ? "Havells" : entLine.GetAttributeValue<OptionSetValue>("hil_brand").Value == 2 ? "Lloyd" : "WaterPurifier",
                                    AssetStatus = entLine.GetAttributeValue<bool>("hil_warrantystatus"),
                                    AssetStatusName = entLine.GetAttributeValue<bool>("hil_warrantystatus") ? "Success" : "Failed",
                                    Is_Type = entLine.GetAttributeValue<string>("hil_istype"),
                                };
                                _header.ProductLines.Add(_line);
                            }
                            _retObj.MRNHeader.Add(_header);
                        }
                        _retObj.ResultStatus = true;
                        _retObj.ResultMessage = "OK";
                    }
                }
                else
                {
                    _retObj = new MRNSummary() { ResultStatus = false, ResultMessage = "D365 Service Unavailable." };
                }
            }
            catch (Exception ex)
            {
                _retObj = new MRNSummary() { ResultStatus = false, ResultMessage = ex.Message };
            }
            return _retObj;
        }

        public MRNSummary GetDefectiveStockNoteForSAP(ViewListSearch _searchCondition)
        {
            IOrganizationService service = null;
            MRNSummary _retObj = null;
            EntityCollection entCol;
            EntityCollection entColLine;
            MRNHeader _header = null;
            MRNLine _line = null;
            try
            {
                service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (_searchCondition.FromDate == null ||_searchCondition.FromDate == string.Empty)
                    {
                        return new MRNSummary() { ResultStatus = false, ResultMessage = "Created From Date is required." };
                    }

                    string _fetchLineXML = string.Empty;
                    string _invDate = string.Empty;
                    DateTime dtTemp = DateTime.Now;
                    string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='hil_dealerstockverificationheader'>
                    <attribute name='hil_dealerstockverificationheaderid' />
                    <attribute name='hil_name' />
                    <attribute name='hil_mrnstatus' />
                    <attribute name='hil_customercode' />
                    <attribute name='ownerid' />
                    <attribute name='hil_pdinumber' />
                    <attribute name='createdon' />
                    <order attribute='hil_name' descending='false' />
                    <filter type='and'>";

                    if (_searchCondition.FromDate.Trim().Length > 0)
                    {
                        _fetchXML += @"<condition attribute='createdon' operator='on-or-after' value='" + _searchCondition.FromDate + @"' />";
                    }

                    if (_searchCondition.ToDate.Trim().Length > 0)
                    {
                        _fetchXML += @"<condition attribute='createdon' operator='on-or-before' value='" + _searchCondition.ToDate + @"' />";
                    }
                    //if (_searchCondition.ReferenceNumber.Trim().Length > 0)
                    //{
                    //    _fetchXML += @"<condition attribute='hil_name' operator='eq' value='" + _searchCondition.ReferenceNumber + @"' />";
                    //}
                    _fetchXML += @"</filter>
                    <link-entity name='account' from='accountid' to='hil_customercode' link-type='inner' alias='cust'>
                        <attribute name='accountnumber' />
                    </link-entity>
                    <link-entity name='systemuser' from='systemuserid' to='owninguser' visible='false' link-type='outer' alias='emp'>
                        <attribute name='internalemailaddress' />
                        <attribute name='fullname' />
                        <attribute name='hil_employeecode' />
                    </link-entity>
                    </entity>
                    </fetch>";

                    entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entCol.Entities.Count > 0)
                    {
                        EntityReference entRef = entCol.Entities[0].GetAttributeValue<EntityReference>("hil_dealerstockverificationheader");
                        _retObj = new MRNSummary();
                        _retObj.MRNHeader = new List<MRNHeader>();
                        foreach (Entity entHeader in entCol.Entities)
                        {
                            _header = new MRNHeader();
                            _header.MRNGuId = entHeader.Id;
                            _header.PdiNumber = entHeader.GetAttributeValue<string>("hil_pdinumber");
                            _header.Technician = entHeader.GetAttributeValue<EntityReference>("ownerid").Id;
                            dtTemp = entHeader.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                            _invDate = dtTemp.Year.ToString() + dtTemp.Month.ToString().PadLeft(2, '0') + dtTemp.Day.ToString().PadLeft(2, '0');
                            _header.CreatedOn = _invDate;

                            if (entHeader.Attributes.Contains("hil_mrnstatus"))
                            {
                                OptionSetValue _headerStatus = entHeader.GetAttributeValue<OptionSetValue>("hil_mrnstatus");
                                _header.MRNStatus = _headerStatus.Value;

                                if (entHeader.FormattedValues.Contains("hil_mrnstatus"))
                                    _header.MRNStatusName = entHeader.FormattedValues["hil_mrnstatus"];
                            }
                            if (entHeader.Attributes.Contains("cust.accountnumber"))
                            {
                                _header.CustomerCode = entHeader.GetAttributeValue<AliasedValue>("cust.accountnumber").Value.ToString();
                                _header.CustomerName = entHeader.GetAttributeValue<EntityReference>("hil_customercode").Name;
                                _header.CustomerGuId = entHeader.GetAttributeValue<EntityReference>("hil_customercode").Id;
                            }
                            if (entHeader.Attributes.Contains("emp.internalemailaddress"))
                            {
                                _header.TechnicianEMail = entHeader.GetAttributeValue<AliasedValue>("emp.internalemailaddress").Value.ToString();
                            }
                            if (entHeader.Attributes.Contains("emp.fullname"))
                            {
                                _header.TechnicianName = entHeader.GetAttributeValue<AliasedValue>("emp.fullname").Value.ToString();
                            }
                            if (entHeader.Attributes.Contains("emp.hil_employeecode"))
                            {
                                _header.TechnicianCode = entHeader.GetAttributeValue<AliasedValue>("emp.hil_employeecode").Value.ToString();
                            }
                            if (entHeader.Attributes.Contains("hil_name"))
                            {
                                _header.ReferenceNumber = entHeader.GetAttributeValue<string>("hil_name").ToString();
                            }
                            _header.ProductLines = new List<MRNLine>();
                            _fetchLineXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_materialreturn'>
                            <attribute name='hil_materialreturnid' />
                            <attribute name='hil_name' />
                            <attribute name='createdon' />
                            <attribute name='hil_warrantystatus' />
                            <attribute name='ownerid' />
                            <attribute name='hil_dealerstockverificationheader' />
                            <attribute name='hil_productcode' />
                            <attribute name='hil_product' />
                            <attribute name='hil_longitude' />
                            <attribute name='hil_latitude' />
                            <attribute name='hil_brand' />
                            <attribute name='hil_divisioncode' />
                            <attribute name='hil_invoicenumber' />
                            <attribute name='hil_invoicedate' />
                            <attribute name='hil_istype' />
                            <order attribute='createdon' descending='false' />
                            <filter type='and'>
                                <condition attribute='hil_dealerstockverificationheader' operator='eq' value='" + entHeader.Id.ToString() + @"' />
                            </filter>
                            </entity>
                            </fetch>";
                            entColLine = service.RetrieveMultiple(new FetchExpression(_fetchLineXML));
                            foreach (Entity entLine in entColLine.Entities)
                            {
                                dtTemp = entLine.GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330);
                                _invDate = dtTemp.Year.ToString() + dtTemp.Month.ToString().PadLeft(2, '0') + dtTemp.Day.ToString().PadLeft(2, '0');
                                _line = new MRNLine()
                                {
                                    MRNLineGuId = entLine.Id,
                                    SerialNumber = entLine.GetAttributeValue<string>("hil_name"),
                                    ProductCode = entLine.GetAttributeValue<string>("hil_productcode"),
                                    ProductGuId = entLine.GetAttributeValue<EntityReference>("hil_product").Id,
                                    InvoiceDate = _invDate,
                                    InvoiceNumber = entLine.GetAttributeValue<string>("hil_invoicenumber"),
                                    Latitude = entLine.GetAttributeValue<string>("hil_latitude"),
                                    Longitude = entLine.GetAttributeValue<string>("hil_longitude"),
                                    DIVISION = entLine.GetAttributeValue<string>("hil_divisioncode"),
                                    BRAND = entLine.GetAttributeValue<OptionSetValue>("hil_brand").Value,
                                    BRANDNAME = entLine.GetAttributeValue<OptionSetValue>("hil_brand").Value == 1 ? "Havells" : entLine.GetAttributeValue<OptionSetValue>("hil_brand").Value == 2 ? "Lloyd" : "WaterPurifier",
                                    AssetStatus = entLine.GetAttributeValue<bool>("hil_warrantystatus"),
                                    AssetStatusName = entLine.GetAttributeValue<bool>("hil_warrantystatus") ? "Success" : "Failed",
                                    Is_Type = entLine.GetAttributeValue<string>("hil_istype")
                                };
                                _header.ProductLines.Add(_line);
                            }
                            _retObj.MRNHeader.Add(_header);
                        }
                        _retObj.ResultStatus = true;
                        _retObj.ResultMessage = "OK";
                    }
                    else {
                        _retObj = new MRNSummary() { ResultStatus = true, ResultMessage = "No record found." };
                    }
                }
                else
                {
                    _retObj = new MRNSummary() { ResultStatus = false, ResultMessage = "D365 Service Unavailable." };
                }
            }
            catch (Exception ex)
            {
                _retObj = new MRNSummary() { ResultStatus = false, ResultMessage = ex.Message };
            }
            return _retObj;
        }

        public List<PDOICalls> GetPDICalls(MRNEntry job)
        {
            List<PDOICalls> jobList = new List<PDOICalls>();
            IOrganizationService service = ConnectToCRM.GetOrgService();
            string FetchQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='msdyn_workorder'>
                <attribute name='msdyn_name' />
                <order attribute='msdyn_name' descending='false' />
                <filter type='and'>
                    <condition attribute='ownerid' operator='eq' uiname='anonymous' uitype='systemuser' value='{" + job.Technician + @"}' />
                    <condition attribute='hil_callsubtype' operator='eq' uiname='PDI' uitype='hil_callsubtype' value='{CE45F586-3C0B-E911-A94E-000D3AF06CD4}' />
                    <condition attribute='msdyn_substatus' operator='not-in'>
                        <value uiname='Work Done' uitype='msdyn_workordersubstatus'>{2927FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                        <value uiname='Work Done SMS' uitype='msdyn_workordersubstatus'>{7E85074C-9C54-E911-A951-000D3AF0677F}</value>
                        <value uiname='Canceled' uitype='msdyn_workordersubstatus'>{1527FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                        <value uiname='Closed' uitype='msdyn_workordersubstatus'>{1727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                        <value uiname='KKG Audit Failed' uitype='msdyn_workordersubstatus'>{6C8F2123-5106-EA11-A811-000D3AF057DD}</value>
                    </condition>
                </filter>
                </entity>
            </fetch>";
            EntityCollection colAssignedJob = service.RetrieveMultiple(new FetchExpression(FetchQuery));
            if (colAssignedJob.Entities.Count > 0)
            {

                foreach (Entity item in colAssignedJob.Entities)
                {
                    jobList.Add(new PDOICalls()
                    {
                        Job_ID = item.Id.ToString(),
                        Job_NO = item.GetAttributeValue<string>("msdyn_name")
                    });
                }
            }
            return jobList;
        }
    }

    [DataContract]
    public class ViewListSearch
    {
        [DataMember]
        public string Date { get; set; }
        [DataMember]
        public string FromDate { get; set; }
        [DataMember]
        public string ToDate { get; set; }
        [DataMember]
        public Guid CustomerGuId { get; set; }
        [DataMember]
        public Guid TechnicianGuId { get; set; }
        [DataMember]
        public string ReferenceNumber { get; set; }
    }

    [DataContract]
    public class CustomerInfo
    {
        [DataMember]
        public string CustomerCode { get; set; }
        [DataMember]
        public string CustomerName { get; set; }
        [DataMember]
        public Guid CustomerGuId { get; set; }
        [DataMember]
        public bool ResultStatus { get; set; }
        [DataMember]
        public string ResultMessage { get; set; }
    }

    [DataContract]
    public class MRNHeader
    {
        [DataMember]
        public Guid MRNGuId { get; set; }
        [DataMember]
        public string PdiNumber { get; set; }
        [DataMember]
        public string CreatedOn { get; set; }
        [DataMember]
        public string TechnicianEMail { get; set; }
        [DataMember]
        public string TechnicianCode { get; set; }
        [DataMember]
        public string TechnicianName { get; set; }
        [DataMember]
        public Guid Technician { get; set; }
        [DataMember]
        public Guid CustomerGuId { get; set; }
        [DataMember]
        public string CustomerCode { get; set; }
        [DataMember]
        public string CustomerName { get; set; }
        [DataMember]
        public string ReferenceNumber { get; set; }
        /// <summary>
        /// 1,2,3
        /// </summary>
        [DataMember]
        public int MRNStatus { get; set; }
        /// <summary>
        /// Draft, Submitted, Posted
        /// </summary>
        [DataMember]
        public string MRNStatusName { get; set; }
        [DataMember]
        public bool ResultStatus { get; set; }
        [DataMember]
        public string ResultMessage { get; set; }
        [DataMember]
        public List<MRNLine> ProductLines { get; set; }
    }

    [DataContract]
    public class MRNLine
    {
        [DataMember]
        public Guid MRNLineGuId { get; set; }
        [DataMember]
        public string CustomerCode { get; set; }
        [DataMember]
        public string CustomerName { get; set; }
        [DataMember]
        public string SerialNumber { get; set; }
        [DataMember]
        public string ProductCode { get; set; }
        [DataMember]
        public string DIVISION { get; set; }
        [DataMember]
        public int? BRAND { get; set; }
        [DataMember]
        public string BRANDNAME { get; set; }
        [DataMember]
        public Guid ProductGuId { get; set; }
        [DataMember]
        public string InvoiceDate { get; set; }
        [DataMember]
        public string InvoiceNumber { get; set; }
        [DataMember]
        public string Latitude { get; set; }
        [DataMember]
        public string Longitude { get; set; }
        [DataMember]
        public bool AssetStatus { get; set; }
        [DataMember]
        public string AssetStatusName { get; set; }
        [DataMember]
        public string Is_Type { get; set; }
    }

    [DataContract]
    public class MRNSummary
    {
        [DataMember]
        public List<MRNHeader> MRNHeader { get; set; }
        [DataMember]
        public bool ResultStatus { get; set; }
        [DataMember]
        public string ResultMessage { get; set; }
    }

    [DataContract]
    public class PDOICalls
    {
        [DataMember]
        public string Job_ID { get; set; }

        [DataMember]
        public string Job_NO { get; set; }
    }
    [DataContract]
    public class ACInOutScanResult
    {
        [DataMember]
        public bool ResultStatus { get; set; }

        [DataMember]
        public string ResultMessage { get; set; }
    }

}
