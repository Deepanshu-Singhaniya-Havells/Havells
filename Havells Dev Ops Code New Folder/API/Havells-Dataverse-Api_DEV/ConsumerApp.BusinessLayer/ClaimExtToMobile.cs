using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Text;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class ClaimExtToMobile
    {
        [DataMember]
        public Guid WorkOrderId { get; set; }

        public List<WOSchemesResult> GetSchemeCodes(ClaimExtToMobile _reqParam)
        {
            IOrganizationService service = null;
            List<WOSchemesResult> _retObj = new List<WOSchemesResult>();
            try
            {
                #region Validate API IN Params
                if (_reqParam.WorkOrderId == Guid.Empty)
                {
                    _retObj.Add(new WOSchemesResult() { ResultStatus = false, ResultMessage = "Work Order GuId is required." });
                    return _retObj;
                }
                #endregion
                DateTime _PurchaseDate;
                Guid _CallSubType = Guid.Empty;
                Guid _SalesOffice = Guid.Empty;
                Guid _ProdSubCatg = Guid.Empty;
                OptionSetValue _CallerType = null;
                DateTime _CreatedOn;

                service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    #region Get Work Order Details
                    Entity entWO = service.Retrieve(msdyn_workorder.EntityLogicalName, _reqParam.WorkOrderId, new ColumnSet("hil_purchasedate", "hil_callsubtype", "createdon", "hil_salesoffice", "hil_productsubcategory", "hil_callertype"));
                    if (entWO != null)
                    {
                        if (entWO.Attributes.Contains("hil_callertype"))
                        {
                            _CallerType = entWO.GetAttributeValue<OptionSetValue>("hil_callertype");
                        }
                        if (entWO.Attributes.Contains("hil_purchasedate"))
                        {
                            _PurchaseDate = entWO.GetAttributeValue<DateTime>("hil_purchasedate").AddMinutes(330);
                        }
                        else
                        {
                            _PurchaseDate = new DateTime(1900, 1, 1);
                        }
                        if (entWO.Attributes.Contains("hil_callsubtype"))
                        {
                            _CallSubType = entWO.GetAttributeValue<EntityReference>("hil_callsubtype").Id;
                        }
                        else
                        {
                            _retObj.Add(new WOSchemesResult() { ResultStatus = false, ResultMessage = "Call Sub Type is not defined in Work Order." });
                            return _retObj;
                        }
                        if (entWO.Attributes.Contains("hil_salesoffice"))
                        {
                            _SalesOffice = entWO.GetAttributeValue<EntityReference>("hil_salesoffice").Id;
                        }
                        else
                        {
                            _retObj.Add(new WOSchemesResult() { ResultStatus = false, ResultMessage = "Sales Office is not defined in Work Order." });
                            return _retObj;
                        }
                        if (entWO.Attributes.Contains("hil_productsubcategory"))
                        {
                            _ProdSubCatg = entWO.GetAttributeValue<EntityReference>("hil_productsubcategory").Id;
                        }
                        else
                        {
                            _retObj.Add(new WOSchemesResult() { ResultStatus = false, ResultMessage = "Product Sub Category is not defined in Work Order." });
                            return _retObj;
                        }
                        _CreatedOn = entWO.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                        string _purchaseDateValue = _PurchaseDate.Year.ToString() + "-" + _PurchaseDate.Month.ToString().PadLeft(2, '0') + "-" + _PurchaseDate.Day.ToString().PadLeft(2, '0');
                        string _createdOnValue = _CreatedOn.Year.ToString() + "-" + _CreatedOn.Month.ToString().PadLeft(2, '0') + "-" + _CreatedOn.Day.ToString().PadLeft(2, '0');

                        string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_schemeincentive'>
                        <attribute name='hil_schemeincentiveid' />
                        <attribute name='hil_name' />
                        <order attribute='hil_name' descending='false' />
                        <filter type='and'>
                            <condition attribute='hil_schemeexpirydate' operator='on-or-after' value='" + _createdOnValue + @"' />
                            <condition attribute='hil_fromdate' operator='on-or-before' value='" + _purchaseDateValue + @"' />
                            <condition attribute='hil_todate' operator='on-or-after' value='" + _purchaseDateValue + @"' />
                            <condition attribute='hil_callsubtype' operator='eq' value='{" + _CallSubType + @"}' />
                            <condition attribute='hil_productsubcategory' operator='eq' value='{" + _ProdSubCatg + @"}' />
                            <condition attribute='hil_salesoffice' operator='in'>
                            <value >{" + _SalesOffice + @"}</value>
                            </condition>
                            <condition attribute='statecode' operator='eq' value='0' />";
                        if (_CallerType != null)
                            fetchXML = fetchXML + @"<condition attribute='hil_callertype' operator='eq' value='{" + _CallerType.Value.ToString() + @"}' />";
                        fetchXML = fetchXML + @"</filter></entity></fetch>";

                        EntityCollection entColConsumer = service.RetrieveMultiple(new FetchExpression(fetchXML));
                        if (entColConsumer.Entities.Count > 0)
                        {
                            foreach (Entity ent in entColConsumer.Entities)
                            {
                                if (ent.Attributes.Contains("hil_name"))
                                {
                                    _retObj.Add(new WOSchemesResult() { SchemeId = ent.Id, SchemeName = ent.GetAttributeValue<string>("hil_name"), ResultStatus = true, ResultMessage = "OK" });
                                }
                            }
                        }
                        else
                        {

                            fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_schemeincentive'>
                                <attribute name='hil_schemeincentiveid' />
                                <attribute name='hil_name' />
                                <order attribute='hil_name' descending='false' />
                                <filter type='and'>
                                <condition attribute='hil_schemeexpirydate' operator='on-or-after' value='" + _createdOnValue + @"' />
                                <condition attribute='hil_fromdate' operator='on-or-before' value='" + _purchaseDateValue + @"' />
                                <condition attribute='hil_todate' operator='on-or-after' value='" + _purchaseDateValue + @"' />
                                <condition attribute='hil_callsubtype' operator='eq' value='{" + _CallSubType + @"}' />
                                <condition attribute='hil_productsubcategory' operator='eq' value='{" + _ProdSubCatg + @"}' />
                                <condition attribute='hil_salesoffice' operator='in'>
                                <value >{90503976-8FD1-EA11-A813-000D3AF0563C}</value>
                                </condition>
                                <condition attribute='statecode' operator='eq' value='0' />";
                            if (_CallerType != null)
                                fetchXML = fetchXML + @"<condition attribute='hil_callertype' operator='eq' value='{" + _CallerType.Value.ToString() + @"}' />";
                            fetchXML = fetchXML + @"</filter></entity></fetch>";

                            entColConsumer = service.RetrieveMultiple(new FetchExpression(fetchXML));
                            if (entColConsumer.Entities.Count > 0)
                            {
                                foreach (Entity ent in entColConsumer.Entities)
                                {
                                    if (ent.Attributes.Contains("hil_name"))
                                    {
                                        _retObj.Add(new WOSchemesResult() { SchemeId = ent.Id, SchemeName = ent.GetAttributeValue<string>("hil_name"), ResultStatus = true, ResultMessage = "OK" });
                                    }
                                }
                            }
                            else
                            {
                                _retObj.Add(new WOSchemesResult() { ResultStatus = false, ResultMessage = "No Record found." });
                            }
                        }
                    }
                    else
                    {
                        _retObj.Add(new WOSchemesResult() { ResultStatus = false, ResultMessage = "Work Order Id does not exist !!! Something went wrong." });
                    }
                    return _retObj;
                    #endregion
                }
                else
                {
                    _retObj.Add(new WOSchemesResult() { ResultStatus = false, ResultMessage = "D365 Service Unavailable" });
                    return _retObj;
                }
            }
            catch (Exception ex)
            {
                _retObj.Add(new WOSchemesResult() { ResultStatus = false, ResultMessage = ex.Message });
                return _retObj;
            }
        }
    }

    [DataContract]
    public class WOSchemesResult
    {
        [DataMember]
        public Guid SchemeId { get; set; }
        [DataMember]
        public string SchemeName { get; set; }
        [DataMember]
        public bool ResultStatus { get; set; }
        [DataMember]
        public string ResultMessage { get; set; }
    }
}
