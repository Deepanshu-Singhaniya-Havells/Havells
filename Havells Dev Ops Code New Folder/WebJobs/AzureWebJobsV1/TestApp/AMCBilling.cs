using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Serialization;

namespace TestApp
{
    public class AMCBilling
    {
        public Guid JobId { get; set; }
        public Decimal? ReceiptAmount { get; set; }
        public bool ResultStatus { get; set; }
        public string ResultMessage { get; set; }
        
        public AMCBilling ValidateAMCReceiptAmount(AMCBilling _reqData, IOrganizationService service)
        {
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
                //service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
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

                        if (entTemp.Id != new Guid("55A71A52-3C0B-E911-A94E-000D3AF06CD4"))
                        {
                            _retObj = new AMCBilling() { ResultStatus = true, ResultMessage = "OK." };
                            return _retObj;
                        }
                        decimal _payableAmount = ((Money)(entCol.Entities[0].GetAttributeValue<AliasedValue>("wo.hil_actualcharges").Value)).Value;
                        DateTime _jobDate = (DateTime)entCol.Entities[0].GetAttributeValue<AliasedValue>("wo.createdon").Value;

                        //_asOn Definition :: AMC Job Create date is concidered for Applying Discount rate becoz Product ageing also calculated from AMC Job Create Date
                        string _asOn = _jobDate.Year.ToString() + "-" + _jobDate.Month.ToString().PadLeft(2, '0') + "-" + _jobDate.Day.ToString().PadLeft(2, '0');
                        int _dayDiff = Convert.ToInt32(Math.Round((_jobDate - _invoiceDate).TotalDays, 0));
                        decimal _stdDiscPer = 0;
                        decimal _spcDiscPer = 0;
                        decimal _stdDiscAmount = 0;
                        decimal _spcDiscAmount = 0;

                        _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_amcdiscountmatrix'>
                            <attribute name='hil_amcdiscountmatrixid' />
                            <attribute name='hil_name' />
                            <attribute name='createdon' />
                            <attribute name='hil_discounttype' />
                            <attribute name='hil_discper' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                                <condition attribute='hil_productaegingstart' operator='le' value='" + _dayDiff.ToString() + @"' />
                                <condition attribute='hil_productageingend' operator='ge' value='" + _dayDiff.ToString() + @"' />
                                <condition attribute='hil_validfrom' operator='on-or-before' value='" + _asOn + @"' />
                                <condition attribute='hil_validto' operator='on-or-after' value='" + _asOn + @"' />
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
                            _stdDiscAmount = Math.Round((_payableAmount - (_payableAmount * _stdDiscPer) / 100), 2);
                            _spcDiscAmount = Math.Round(_payableAmount - (_payableAmount * (_stdDiscPer + _spcDiscPer)) / 100, 2);

                            if (_reqData.ReceiptAmount < _spcDiscAmount)
                            {
                                _retObj = new AMCBilling() { ResultStatus = false, ResultMessage = "As Per AMC Discount Policy minimum allowed receipt amount is " + _spcDiscAmount.ToString() };
                            }
                            else
                            {
                                _retObj = new AMCBilling() { ResultStatus = true, ResultMessage = "OK" };
                            }
                        }
                        else
                        {
                            if (_payableAmount != _reqData.ReceiptAmount)
                            {
                                _retObj = new AMCBilling() { ResultStatus = false, ResultMessage = "No AMC Discount Policy is defined in System !!! Receipt Amount can't be less than Payable Amount." };
                            }
                            else
                            {
                                _retObj = new AMCBilling() { ResultStatus = true, ResultMessage = "OK" };
                            }
                        }
                    }
                    else
                    {
                        _retObj = new AMCBilling() { ResultStatus = false, ResultMessage = "No Work Order Incident found." };
                    }
                    return _retObj;
                }
                else
                {
                    _retObj = new AMCBilling() { ResultStatus = false, ResultMessage = "D365 Service Unavailable" };
                    return _retObj;
                }
            }
            catch (Exception ex)
            {
                _retObj = new AMCBilling() { ResultStatus = false, ResultMessage = ex.Message };
                return _retObj;
            }
        }
    }
}
