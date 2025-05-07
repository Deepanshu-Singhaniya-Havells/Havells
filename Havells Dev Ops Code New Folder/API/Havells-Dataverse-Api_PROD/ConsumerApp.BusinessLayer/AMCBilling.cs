
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.Serialization;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class AMCBilling
    {
        [DataMember]
        public Guid JobId { get; set; }

        [DataMember]
        public Decimal? ReceiptAmount { get; set; }

        [DataMember]
        public int? SourceCode { get; set; }

        [DataMember]
        public bool ResultStatus { get; set; }
        [DataMember]
        public string ResultMessage { get; set; }
        [DataMember]
        public string ResultMessageType { get; set; }

        public AMCBilling ValidateAMCReceiptAmount(AMCBilling _reqData)
        {
            IOrganizationService service;
            AMCBilling _retObj = new AMCBilling();
            string _fetchXML = string.Empty;
            DateTime _invoiceDate;

            try
            {
                if (_reqData.JobId == Guid.Empty)
                {
                    _retObj = new AMCBilling() { ResultStatus = false, ResultMessage = "Job Id is required." };
                    return _retObj;
                }
                if (_reqData.ReceiptAmount == null || _reqData.ReceiptAmount == 0)
                {
                    _retObj = new AMCBilling() { ResultStatus = false, ResultMessage = "Receipt Amount is required." };
                    return _retObj;
                }
                service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    //if (_reqData.SourceCode == null)
                    //{
                    _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='msdyn_workorderincident'>
                        <attribute name='msdyn_name' />
                        <filter type='and'>
                          <condition attribute='msdyn_workorder' operator='eq' value='" + _reqData.JobId + @"' />
                            <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='msdyn_customerasset' visible='false' link-type='outer' alias='ca'>
                            <attribute name='hil_invoicedate' />
                            </link-entity>
                            <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorder' visible='false' link-type='outer' alias='wo'>
                            <attribute name='createdon' />
                            <attribute name='hil_actualcharges' />
                            <attribute name='hil_callsubtype' />
                            <attribute name='hil_productcategory' />
                            </link-entity>
                            </entity>
                            </fetch>";

                    EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entCol.Entities.Count > 0)
                    {
                        if (!entCol.Entities[0].Attributes.Contains("wo.hil_callsubtype"))
                        {
                            _retObj = new AMCBilling() { ResultStatus = false, ResultMessage = "Call Subtype is required." };
                            return _retObj;
                        }
                        if (!entCol.Entities[0].Attributes.Contains("ca.hil_invoicedate"))
                        {
                            _invoiceDate = new DateTime(1900, 1, 1);
                        }
                        else
                        {
                            _invoiceDate = (DateTime)(entCol.Entities[0].GetAttributeValue<AliasedValue>("ca.hil_invoicedate").Value);
                        }
                        EntityReference entTemp = (EntityReference)entCol.Entities[0].GetAttributeValue<AliasedValue>("wo.hil_callsubtype").Value;
                        EntityReference entProdCatg = (EntityReference)entCol.Entities[0].GetAttributeValue<AliasedValue>("wo.hil_productcategory").Value;

                        if (entTemp.Id != new Guid("55A71A52-3C0B-E911-A94E-000D3AF06CD4"))
                        {
                            _retObj = new AMCBilling() { ResultStatus = true, ResultMessageType = "INFO", ResultMessage = "OK." };
                            return _retObj;
                        }

                        decimal _payableAmount = 0;
                        if (entCol.Entities[0].Attributes.Contains("wo.hil_actualcharges"))
                        {
                            _payableAmount = ((Money)(entCol.Entities[0].GetAttributeValue<AliasedValue>("wo.hil_actualcharges").Value)).Value;
                        }
                        else
                        {
                            _retObj = new AMCBilling() { ResultStatus = true, ResultMessageType = "INFO", ResultMessage = "OK" };
                        }
                        DateTime _jobDate = (DateTime)entCol.Entities[0].GetAttributeValue<AliasedValue>("wo.createdon").Value;
                        //_asOn Definition :: AMC Job Create date is concidered for Applying Discount rate becoz Product ageing also calculated from AMC Job Create Date
                        string _asOn = _jobDate.Year.ToString() + "-" + _jobDate.Month.ToString().PadLeft(2, '0') + "-" + _jobDate.Day.ToString().PadLeft(2, '0');
                        int _dayDiff = Convert.ToInt32(Math.Round((_jobDate - _invoiceDate).TotalDays, 0));
                        if (_dayDiff < 0)
                        {
                            return new AMCBilling() { ResultStatus = false, ResultMessageType = "WARNING", ResultMessage = "Product Age is -4 days. Job is created prior to Asset Invoice Date." };
                        }
                        decimal _stdDiscPer = 0;
                        decimal _spcDiscPer = 0;
                        decimal _stdDiscAmount = 0;
                        decimal _spcDiscAmount = 0;
                        //03B5A2D6-CC64-ED11-9562-6045BDAC526A - AMC Sale - FSM (Source)
                        _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_amcdiscountmatrix'>
                                <attribute name='hil_amcdiscountmatrixid' />
                                <attribute name='hil_name' />
                                <attribute name='createdon' />
                                <attribute name='hil_discounttype' />
                                <attribute name='hil_discper' />
                                <order attribute='hil_name' descending='false' />
                                <filter type='and'>
                                    <condition attribute='statecode' operator='eq' value='0' />
                                    <condition attribute='hil_appliedto' operator='eq' value='{03B5A2D6-CC64-ED11-9562-6045BDAC526A}' />
                                    <condition attribute='hil_productaegingstart' operator='le' value='" + _dayDiff.ToString() + @"' />
                                    <condition attribute='hil_productageingend' operator='ge' value='" + _dayDiff.ToString() + @"' />
                                    <condition attribute='hil_validfrom' operator='on-or-before' value='" + _asOn + @"' />
                                    <condition attribute='hil_validto' operator='on-or-after' value='" + _asOn + @"' />
                                    <condition attribute='hil_productcategory' operator='eq' value='{" + entProdCatg.Id + @"}' />
                                </filter>
                                </entity>
                                </fetch>";

                        EntityCollection entCol1 = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entCol1.Entities.Count > 0)
                        {
                            foreach (Entity ent in entCol1.Entities)
                            {
                                if (ent.GetAttributeValue<OptionSetValue>("hil_discounttype").Value == 1)
                                {
                                    _stdDiscPer = ent.GetAttributeValue<decimal>("hil_discper");
                                }
                                else
                                {
                                    _spcDiscPer = ent.GetAttributeValue<decimal>("hil_discper");
                                }
                            }
                            _stdDiscAmount = Math.Round((_payableAmount - (_payableAmount * _stdDiscPer) / 100), 2); //Max Limit (90)
                            _spcDiscAmount = Math.Round(_payableAmount - (_payableAmount * (_stdDiscPer + _spcDiscPer)) / 100, 2); //Min Limit (85)
                            if (_reqData.ReceiptAmount >= _spcDiscAmount && _reqData.ReceiptAmount < _stdDiscAmount)
                            {
                                decimal _additionaldisc = Math.Round(_stdDiscAmount - Convert.ToDecimal(_reqData.ReceiptAmount), 2);
                                _retObj = new AMCBilling() { ResultStatus = false, ResultMessageType = "CONFIRMATION", ResultMessage = "To offer additional discount (Rs. " + _additionaldisc.ToString() + ") above Standard Discount, you need to take BSH approval. Click 'Yes' if approval already taken Or Click 'No'." };
                            }
                            else if (_reqData.ReceiptAmount < _spcDiscAmount)
                            {
                                _retObj = new AMCBilling() { ResultStatus = false, ResultMessageType = "WARNING", ResultMessage = "As per AMC discount policy, you are allowed to collect minimum Rs. " + _stdDiscAmount.ToString() + "." };
                            }
                            else
                            {
                                _retObj = new AMCBilling() { ResultStatus = true, ResultMessageType = "INFO", ResultMessage = "OK" };
                            }
                        }
                        else
                        {
                            if (_payableAmount != _reqData.ReceiptAmount)
                            {
                                _retObj = new AMCBilling() { ResultStatus = false, ResultMessageType = "WARNING", ResultMessage = "No AMC Discount Policy is defined in System !!! Receipt Amount can't be less than Payable Amount." };
                            }
                            else
                            {
                                _retObj = new AMCBilling() { ResultStatus = true, ResultMessageType = "INFO", ResultMessage = "OK" };
                            }
                        }
                    }
                    else
                    {
                        _retObj = new AMCBilling() { ResultStatus = false, ResultMessageType = "WARNING", ResultMessage = "No Work Order Incident found." };
                    }
                    return _retObj;
                    //}
                    //else
                    //{
                    //    _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    //        <entity name='hil_integrationsource'>
                    //        <attribute name='hil_integrationsourceid' />
                    //        <filter type='and'>
                    //        <condition attribute='hil_code' operator='eq' value='" + _reqData.SourceCode + @"' />
                    //        </filter>
                    //        </entity>
                    //        </fetch>";

                    //    EntityCollection entCol1 = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    //    if (entCol1.Entities.Count > 0)
                    //    {
                    //        string _asOn = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString().PadLeft(2, '0') + "-" + DateTime.Now.Day.ToString().PadLeft(2, '0');
                    //        decimal _stdDiscPer = 0;
                    //        decimal _spcDiscPer = 0;
                    //        decimal _stdDiscAmount = 0;
                    //        decimal _spcDiscAmount = 0;
                    //        _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    //            <entity name='hil_amcdiscountmatrix'>
                    //            <attribute name='hil_amcdiscountmatrixid' />
                    //            <attribute name='hil_name' />
                    //            <attribute name='createdon' />
                    //            <attribute name='hil_discounttype' />
                    //            <attribute name='hil_discper' />
                    //            <order attribute='hil_name' descending='false' />
                    //            <filter type='and'>
                    //                <condition attribute='statecode' operator='eq' value='0' />
                    //                <condition attribute='hil_appliedto' operator='eq' value='{03B5A2D6-CC64-ED11-9562-6045BDAC526A}' />
                    //                <condition attribute='hil_validfrom' operator='on-or-before' value='" + _asOn + @"' />
                    //                <condition attribute='hil_validto' operator='on-or-after' value='" + _asOn + @"' />
                    //            </filter>
                    //            </entity>
                    //            </fetch>";

                    //        entCol1 = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    //        if (entCol1.Entities.Count > 0)
                    //        {
                    //            foreach (Entity ent in entCol1.Entities)
                    //            {
                    //                if (ent.GetAttributeValue<OptionSetValue>("hil_discounttype").Value == 1)
                    //                {
                    //                    _stdDiscPer = ent.GetAttributeValue<decimal>("hil_discper");
                    //                }
                    //                else
                    //                {
                    //                    _spcDiscPer = ent.GetAttributeValue<decimal>("hil_discper");
                    //                }
                    //            }
                    //            _stdDiscAmount = Math.Round((_payableAmount - (_payableAmount * _stdDiscPer) / 100), 2); //Max Limit (90)
                    //            _spcDiscAmount = Math.Round(_payableAmount - (_payableAmount * (_stdDiscPer + _spcDiscPer)) / 100, 2); //Min Limit (85)
                    //            if (_reqData.ReceiptAmount >= _spcDiscAmount && _reqData.ReceiptAmount < _stdDiscAmount)
                    //            {
                    //                decimal _additionaldisc = Math.Round(_stdDiscAmount - Convert.ToDecimal(_reqData.ReceiptAmount), 2);
                    //                _retObj = new AMCBilling() { ResultStatus = false, ResultMessageType = "CONFIRMATION", ResultMessage = "To offer additional discount (Rs. " + _additionaldisc.ToString() + ") above Standard Discount, you need to take BSH approval. Click 'Yes' if approval already taken Or Click 'No'." };
                    //            }
                    //            else if (_reqData.ReceiptAmount < _spcDiscAmount)
                    //            {
                    //                _retObj = new AMCBilling() { ResultStatus = false, ResultMessageType = "WARNING", ResultMessage = "As per AMC discount policy, you are allowed to collect minimum Rs. " + _stdDiscAmount.ToString() + "." };
                    //            }
                    //            else
                    //            {
                    //                _retObj = new AMCBilling() { ResultStatus = true, ResultMessageType = "INFO", ResultMessage = "OK" };
                    //            }
                    //        }
                    //        else
                    //        {
                    //            if (_payableAmount != _reqData.ReceiptAmount)
                    //            {
                    //                _retObj = new AMCBilling() { ResultStatus = false, ResultMessageType = "WARNING", ResultMessage = "No AMC Discount Policy is defined in System !!! Receipt Amount can't be less than Payable Amount." };
                    //            }
                    //            else
                    //            {
                    //                _retObj = new AMCBilling() { ResultStatus = true, ResultMessageType = "INFO", ResultMessage = "OK" };
                    //            }
                    //        }
                    //    }
                    //}
                }
                else
                {
                    _retObj = new AMCBilling() { ResultStatus = false, ResultMessageType = "WARNING", ResultMessage = "D365 Service Unavailable" };
                    return _retObj;
                }
            }
            catch (Exception ex)
            {
                _retObj = new AMCBilling() { ResultStatus = false, ResultMessageType = "WARNING", ResultMessage = ex.Message };
                return _retObj;
            }
        }

        public ServiceResponseData GetOutstandingAMCs(ReqestData reqestData)
        {
            ServiceResponseData responseData = new ServiceResponseData();
            List<InvoiceInfo> lstinvoiceInfo = new List<InvoiceInfo>();
            string[] formats = { "d/MM/yyyy", "dd/MM/yyyy", "d-MM-yyyy", "dd-MM-yyyy" };
            DateTime DOP;
            Guid ModelGuid = Guid.Empty;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (!DateTime.TryParseExact(reqestData.DOP, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DOP))
                    {
                        responseData.result = new ServiceResult { ResultStatus = false, ResultMessage = "No Content : Invalid DOP. Please Provide Date in <dd/MM/yyyy>." };
                        return responseData;
                    }

                    QueryExpression query = new QueryExpression("product");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria.AddCondition("name", ConditionOperator.Equal, reqestData.ModelNumber);
                    EntityCollection ProductColl = service.RetrieveMultiple(query);
                    if (ProductColl.Entities.Count > 0)
                    {
                        DateTime _DOPTo = DOP.AddDays(7);
                        DateTime _DOPFrom = DOP.AddDays(-7);
                        ModelGuid = ProductColl.Entities[0].Id;
                        query = new QueryExpression("invoice");
                        query.ColumnSet = new ColumnSet("invoiceid", "createdon", "hil_productcode");
                        query.Criteria = new FilterExpression(LogicalOperator.And);

                        FilterExpression filter1 = new FilterExpression(LogicalOperator.And);
                        filter1.AddCondition("customerid", ConditionOperator.Equal, reqestData.CustomerId);
                        filter1.AddCondition("msdyn_invoicedate", ConditionOperator.OnOrAfter, _DOPFrom);
                        filter1.AddCondition("msdyn_invoicedate", ConditionOperator.OnOrBefore, _DOPTo);
                        filter1.AddCondition("hil_customerasset", ConditionOperator.Null);

                        FilterExpression filter2 = new FilterExpression(LogicalOperator.Or);
                        filter2.AddCondition("hil_modelcode", ConditionOperator.Equal, ModelGuid);
                        filter2.AddCondition("hil_newserialnumber", ConditionOperator.Equal, reqestData.SerialNumber);

                        query.Criteria.AddFilter(filter1);
                        query.Criteria.AddFilter(filter2);
                        EntityCollection InvoiceColl = service.RetrieveMultiple(query);

                        if (InvoiceColl.Entities.Count > 0)
                        {
                            foreach (Entity ent in InvoiceColl.Entities)
                            {
                                InvoiceInfo invoiceInfo = new InvoiceInfo();
                                invoiceInfo.InvoiceId = ent.Id;
                                invoiceInfo.CreatedOn = ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString();
                                invoiceInfo.PlanName = ent.GetAttributeValue<EntityReference>("hil_productcode").Name;
                                lstinvoiceInfo.Add(invoiceInfo);
                            }
                            responseData.InvoiceInfo = lstinvoiceInfo;
                            responseData.result = new ServiceResult { ResultStatus = true, ResultMessage = "Success" };
                        }
                        else
                        {
                            responseData.InvoiceInfo = lstinvoiceInfo;
                            responseData.result = new ServiceResult { ResultStatus = false, ResultMessage = "No data found." };
                        }
                    }
                    else
                    {
                        responseData.result = new ServiceResult { ResultStatus = false, ResultMessage = "No Content : Please Provide Valid Model Number of Product." };
                        return responseData;
                    }
                }
                else
                {
                    responseData.result = new ServiceResult { ResultStatus = false, ResultMessage = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                responseData.result = new ServiceResult { ResultStatus = false, ResultMessage = ex.Message };
                return responseData;
            }
            return responseData;
        }
    }
    [DataContract]
    public class ReqestData
    {
        [DataMember]
        public Guid CustomerId { get; set; }
        [DataMember]
        public string ModelNumber { get; set; }
        [DataMember]
        public string SerialNumber { get; set; }
        [DataMember]
        public string DOP { get; set; }
    }
    [DataContract]
    public class ServiceResult
    {
        [DataMember]
        public bool ResultStatus { get; set; }

        [DataMember]
        public string ResultMessage { get; set; }
    }
    [DataContract]
    public class ServiceResponseData
    {
        [DataMember]
        public List<InvoiceInfo> InvoiceInfo { get; set; }
        [DataMember]
        public ServiceResult result { get; set; }
    }
    [DataContract]
    public class InvoiceInfo
    {
        [DataMember]
        public Guid InvoiceId { get; set; }
        [DataMember]
        public string PlanName { get; set; }
        [DataMember]
        public string CreatedOn { get; set; }
    }
}
