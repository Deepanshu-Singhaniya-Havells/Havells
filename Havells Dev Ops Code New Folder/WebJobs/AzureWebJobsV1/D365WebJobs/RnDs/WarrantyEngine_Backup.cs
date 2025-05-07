using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace D365WebJobs
{
    public class WarrantyEngineBackup
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";

        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        #endregion
        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {
                try
                {
                    


                    Entity _entCustomerAsset = _service.Retrieve("msdyn_customerasset", new Guid("52f52c9b-3b01-ef11-9f89-6045bdac2e5b"), new ColumnSet("hil_invoiceavailable", "hil_invoicedate", "hil_productsubcategory", "hil_productcategory", "msdyn_product", "msdyn_name", "statuscode", "hil_branchheadapprovalstatus", "hil_customer"));
                    if (_entCustomerAsset != null)
                    {
                        DateTime? _invoiceDate = null;
                        string _invoiceDateStr = string.Empty;
                        string _serialNumer = _entCustomerAsset.Contains("msdyn_name") ? _entCustomerAsset.GetAttributeValue<string>("msdyn_name") : string.Empty;
                        bool _invoiceAvailable = _entCustomerAsset.Contains("hil_invoiceavailable") ? _entCustomerAsset.GetAttributeValue<bool>("hil_invoiceavailable") : false;
                        string _pinCode = _entCustomerAsset.Contains("hil_retailerpincode") ? _entCustomerAsset.GetAttributeValue<string>("hil_retailerpincode") : string.Empty;

                        if (_entCustomerAsset.Contains("hil_invoicedate"))
                        {
                            _invoiceDate = _entCustomerAsset.GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330).Date;
                            _invoiceDateStr = Convert.ToDateTime(_invoiceDate).ToString("yyyy-MM-dd");
                        }
                        if (!_invoiceAvailable)
                        {
                            UpdateCustomerAssetWarrantyStatus(
                                new WarrantyStatusDTO()
                                {
                                    customerAssetsId = _entCustomerAsset.Id,
                                    warrantyEndDate = new DateTime(1900, 1, 1),
                                    warrantyStatus = new OptionSetValue(2),//{key:"Out Warranty",Value:2}
                                    warrantySubStatus = null
                                });
                            return;
                        }

                        EntityReference _productCategory = _entCustomerAsset.Contains("hil_productcategory") ? _entCustomerAsset.GetAttributeValue<EntityReference>("hil_productcategory") : null;
                        EntityReference _productSubcategory = _entCustomerAsset.Contains("hil_productsubcategory") ? _entCustomerAsset.GetAttributeValue<EntityReference>("hil_productsubcategory") : null;
                        EntityReference _modelNumer = _entCustomerAsset.Contains("msdyn_product") ? _entCustomerAsset.GetAttributeValue<EntityReference>("msdyn_product") : null;
                        EntityReference _customer = _entCustomerAsset.Contains("hil_customer") ? _entCustomerAsset.GetAttributeValue<EntityReference>("hil_customer") : null;

                        bool _isAssetApproved = false;
                        if (_entCustomerAsset.Contains("statuscode"))
                        {
                            if (_entCustomerAsset.GetAttributeValue<OptionSetValue>("statuscode").Value == 910590001)
                            {
                                _isAssetApproved = true;
                            }
                        }
                        if (!_isAssetApproved && _entCustomerAsset.Contains("hil_branchheadapprovalstatus"))
                        {
                            if (_entCustomerAsset.GetAttributeValue<OptionSetValue>("hil_branchheadapprovalstatus").Value == 1)
                            {
                                _isAssetApproved = true;
                            }
                        }

                        DateTime? endDate;
                        if (_productSubcategory == null || _productCategory == null || _modelNumer == null || _customer == null)
                        {
                            UpdateCustomerAssetWarrantyStatus(
                                new WarrantyStatusDTO()
                                {
                                    customerAssetsId = _entCustomerAsset.Id,
                                    warrantyEndDate = null,
                                    warrantyStatus = new OptionSetValue(3),//{key:"Warranty Void",Value:3}
                                    warrantySubStatus = null
                                });
                            InactiveSystemWarranties(_service, _entCustomerAsset);
                            return;
                        }
                        else if (!_isAssetApproved)
                        {
                            UpdateCustomerAssetWarrantyStatus(
                                new WarrantyStatusDTO()
                                {
                                    customerAssetsId = _entCustomerAsset.Id,
                                    warrantyEndDate = null,
                                    warrantyStatus = new OptionSetValue(4),//{key:"Approval Pending",Value:4}
                                    warrantySubStatus = null
                                });
                            InactiveSystemWarranties(_service, _entCustomerAsset);
                            return;
                        }
                        else if (!_invoiceAvailable)
                        {
                            UpdateCustomerAssetWarrantyStatus(
                                new WarrantyStatusDTO()
                                {
                                    customerAssetsId = _entCustomerAsset.Id,
                                    warrantyEndDate = null,
                                    warrantyStatus = new OptionSetValue(3),//{key:"Warranty Void",Value:3}
                                    warrantySubStatus = null
                                });
                            InactiveSystemWarranties(_service, _entCustomerAsset);
                        }
                        else
                        {
                            string _fetchXMLQuery = string.Empty;
                            EntityCollection entCol = null;
                            bool _warrantyExecuted = false;
                            string[] _applicableOn = new string[] { WarrantyApplicableOn.serialNumber, WarrantyApplicableOn.model, WarrantyApplicableOn.productSubcategory };
                            string _filterCondWarrantyHeader = string.Empty;
                            string _filterCondWarrantyLine = string.Empty;

                            foreach (string _element in _applicableOn)
                            {
                                if (_element == WarrantyApplicableOn.serialNumber)
                                {
                                    _filterCondWarrantyLine = $@"<link-entity name='hil_warrantytemplateline' from='hil_warrantytemplate' to='hil_warrantytemplateid' link-type='inner' alias='ad'>
                                      <filter type='and'>
                                        <condition attribute='hil_name' operator='eq' value='{_serialNumer}' />
                                      </filter>
                                    </link-entity>
                                    <link-entity name='hil_warrantytype' from='hil_warrantytypeid' to='hil_warrantytypeindex' visible='false' link-type='outer' alias='wt'>
                                      <attribute name='hil_warrantygenerationorder' />
                                      <order attribute='hil_warrantygenerationorder' descending='false' />
                                    </link-entity>";
                                }
                                else if (_element == WarrantyApplicableOn.model)
                                {
                                    _filterCondWarrantyLine = $@"<link-entity name='hil_warrantytemplateline' from='hil_warrantytemplate' to='hil_warrantytemplateid' link-type='inner' alias='ad'>
                                      <filter type='and'>
                                        <condition attribute='hil_model' operator='eq' value='{_modelNumer.Id}' />
                                      </filter>
                                    </link-entity>
                                    <link-entity name='hil_warrantytype' from='hil_warrantytypeid' to='hil_warrantytypeindex' visible='false' link-type='outer' alias='wt'>
                                      <attribute name='hil_warrantygenerationorder' />
                                      <order attribute='hil_warrantygenerationorder' descending='false' />
                                    </link-entity>";
                                }
                                else if (_element == WarrantyApplicableOn.productSubcategory)
                                {
                                    _filterCondWarrantyHeader = $@"<condition attribute='hil_product' operator='eq' value='{_productSubcategory.Id}' />";
                                    _filterCondWarrantyLine = $@"<link-entity name='hil_warrantytype' from='hil_warrantytypeid' to='hil_warrantytypeindex' visible='false' link-type='outer' alias='wt'>
                                      <attribute name='hil_warrantygenerationorder' />
                                      <order attribute='hil_warrantygenerationorder' descending='false' />
                                    </link-entity>";
                                }

                                int[] _warrantyTypes = new int[] { WarrantyType.zeroWarranty, WarrantyType.standard, WarrantyType.specialScheme, WarrantyType.extended };
                                endDate = _invoiceDate;
                                foreach (int _warrantyType in _warrantyTypes)
                                {
                                    _fetchXMLQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true' top='1'>
                                      <entity name='hil_warrantytemplate'>
                                        <attribute name='hil_name' />
                                        <attribute name='createdon' />
                                        <attribute name='hil_warrantyperiod' />
                                        <attribute name='hil_type' />
                                        <attribute name='hil_product' />
                                        <attribute name='hil_category' />
                                        <attribute name='hil_warrantytemplateid' />
                                        <order attribute='modifiedon' descending='true' />
                                        <filter type='and'>
                                          <condition attribute='statecode' operator='eq' value='0' />
                                          <condition attribute='hil_templatestatus' operator='eq' value='2' />
                                          <condition attribute='hil_validfrom' operator='on-or-before' value='{_invoiceDateStr}' />
                                          <condition attribute='hil_validto' operator='on-or-after' value='{_invoiceDateStr}' />
                                          <condition attribute='hil_applicableon' operator='eq' value='{_element}' />
                                          <condition attribute='hil_type' operator='eq' value='{_warrantyType}' />{_filterCondWarrantyHeader}
                                        </filter>{_filterCondWarrantyLine}
                                      </entity>
                                    </fetch>";
                                    entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXMLQuery));
                                    if (entCol.Entities.Count > 0)
                                    {
                                        foreach (Entity entWrty in entCol.Entities)
                                        {
                                            if (_warrantyType == WarrantyType.zeroWarranty)
                                            {
                                                UpdateCustomerAssetWarrantyStatus(
                                                new WarrantyStatusDTO()
                                                {
                                                    customerAssetsId = _entCustomerAsset.Id,
                                                    warrantyEndDate = null,
                                                    warrantyStatus = new OptionSetValue(3),//{key:"Warranty Void",Value:3}
                                                    warrantySubStatus = null
                                                });
                                                InactiveSystemWarranties(_service, _entCustomerAsset);
                                                CreateZeroUnitWarrantyLine(_service, _customer, _entCustomerAsset.ToEntityReference(), _productCategory, _productSubcategory, _modelNumer, _modelNumer.Name, entCol.Entities[0].ToEntityReference(), _warrantyType);
                                            }
                                            else if (_warrantyType == WarrantyType.specialScheme)//Special Scheme
                                            {
                                                bool _retValue = ValidateWarrantySchemeLines(_service, entWrty.ToEntityReference(), _pinCode, endDate);
                                                if (_retValue)
                                                {
                                                    endDate = CreateUnitWarrantyLine(_service, entWrty.GetAttributeValue<int>("hil_warrantyperiod"), 1, _customer, _entCustomerAsset.ToEntityReference(), _productCategory, _productSubcategory, _modelNumer, _modelNumer.Name, endDate, entWrty.ToEntityReference(), _warrantyType);
                                                    endDate = Convert.ToDateTime(endDate).AddDays(1);
                                                    _warrantyExecuted = true;
                                                }
                                            }
                                            else //Standard|Part Warranty
                                            {
                                                endDate = CreateUnitWarrantyLine(_service, entWrty.GetAttributeValue<int>("hil_warrantyperiod"), 1, _customer, _entCustomerAsset.ToEntityReference(), _productCategory, _productSubcategory, _modelNumer, _modelNumer.Name, endDate, entWrty.ToEntityReference(), _warrantyType);
                                                endDate = Convert.ToDateTime(endDate).AddDays(1);
                                                _warrantyExecuted = true;
                                            }
                                        }
                                    }
                                }
                                if (_warrantyExecuted) { break; }
                            }
                        }
                        CalculateAssetWarrantyStatus(_service, _entCustomerAsset);
                    }

                }
                catch (Exception ex)
                {
                   // Create an Email Record and send it to CRM Admin - MDM
                }
            }
        }
        static bool CheckForAssetWarrantyLines(IOrganizationService service, Entity entCA)
        {
            bool _retValue = false;
            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_unitwarranty'>
                <attribute name='hil_warrantyenddate' />
                <filter type='and'>
                    <condition attribute='hil_customerasset' operator='eq' value='{entCA.Id}' />
                    <condition attribute='statecode' operator='eq' value='0' />
                </filter>
                </entity>
                </fetch>";
            try
            {
                EntityCollection entcol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entcol.Entities.Count > 0) { _retValue= true; }
            }
            catch {}
            return _retValue;
        }
        static void CalculateAssetWarrantyStatus(IOrganizationService service, Entity entCA)
        {
            string _processDate = DateTime.Now.ToString("yyyy-MM-dd");
            Entity _ent = new Entity(entCA.LogicalName, entCA.Id);

            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_unitwarranty'>
                <attribute name='hil_warrantyenddate' />
                <filter type='and'>
                    <condition attribute='hil_customerasset' operator='eq' value='{entCA.Id}' />
                    <condition attribute='statecode' operator='eq' value='0' />
                    <condition attribute='hil_warrantyenddate' operator='on-or-after' value='{_processDate}' />
                </filter>
                <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='wt'>
                    <attribute name='hil_type' />
                    <link-entity name='hil_warrantytype' from='hil_warrantytypeid' to='hil_warrantytypeindex' visible='false' link-type='outer' alias='wtt'>
                        <attribute name='hil_executionindex' />
                        <order attribute='hil_executionindex' descending='false' />
                    </link-entity>
                </link-entity>
                </entity>
                </fetch>";

            EntityCollection entcoll3 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));

            if (entcoll3.Entities.Count == 0)
            {
                //Checking for Zero Warranty Line
                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_unitwarranty'>
                        <attribute name='hil_unitwarrantyid' />
                        <filter type='and'>
                            <condition attribute='hil_customerasset' operator='eq' value='{entCA.Id}' />
                            <condition attribute='statecode' operator='eq' value='0' />
                        </filter>
                        <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='ab'>
                            <filter type='and'>
                            <condition attribute='hil_type' operator='eq' value='8' />
                            </filter>
                        </link-entity>
                        </entity>
                    </fetch>";

                EntityCollection entcoll4 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entcoll4.Entities.Count > 0)
                {
                    _ent["hil_warrantystatus"] = new OptionSetValue(3); //Warranty Void
                }
                else
                {
                    _ent["hil_warrantystatus"] = new OptionSetValue(2); //Out Warranty
                }
                _ent["hil_warrantysubstatus"] = null;

                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
                    <entity name='hil_unitwarranty'>
                    <attribute name='hil_warrantyenddate' /> 
                    <order attribute='hil_warrantyenddate' descending='true' />
                    <filter type='and'>
                        <condition attribute='hil_customerasset' operator='eq' value='{entCA.Id}' />
                        <condition attribute='statecode' operator='eq' value='0' />
                    </filter>
                    </entity>
                    </fetch>";

                entcoll4 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entcoll4.Entities.Count > 0)
                    _ent["hil_warrantytilldate"] = entcoll4.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330).Date;
                else
                    _ent["hil_warrantytilldate"] = null;
            }
            else
            {
                DateTime _warrantyEndDate = entcoll3.Entities.Max(x => x.GetAttributeValue<DateTime>("hil_warrantyenddate")).AddMinutes(330).Date;
                OptionSetValue _warrantyType = (OptionSetValue)entcoll3.Entities[0].GetAttributeValue<AliasedValue>("wt.hil_type").Value;

                _ent["hil_warrantystatus"] = new OptionSetValue(1); //In Warranty

                int _warrantySubStatus = _warrantyType.Value;

                if (_warrantySubStatus == 5) { _warrantySubStatus = 2; }
                else if (_warrantySubStatus == 7) { _warrantySubStatus = 3; }
                else if (_warrantySubStatus == 3) { _warrantySubStatus = 4; }

                _ent["hil_warrantysubstatus"] = new OptionSetValue(_warrantySubStatus);
                _ent["hil_warrantytilldate"] = _warrantyEndDate;
            }
            _service.Update(_ent);
        }
        static void UpdateCustomerAssetWarrantyStatus(WarrantyStatusDTO _assetWarrantyStatus) {
            try
            {
                Entity entCustAsset = new Entity("msdyn_customerasset");
                entCustAsset.Id = _assetWarrantyStatus.customerAssetsId;
                entCustAsset["hil_warrantytilldate"] = _assetWarrantyStatus.warrantyEndDate;
                //entCustAsset["hil_extendedwarrantyenddate"] = _assetWarrantyStatus.partWarrantyEndDate;
                entCustAsset["hil_warrantystatus"] = _assetWarrantyStatus.warrantyStatus;
                entCustAsset["hil_warrantysubstatus"] = _assetWarrantyStatus.warrantySubStatus;
                _service.Update(entCustAsset);
            }
            catch (Exception ex)
            {
                
            }
        }
        static void InactiveSystemWarranties(IOrganizationService service, Entity entAsset)
        {
            try
            {
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='hil_unitwarranty'>
                    <attribute name='hil_unitwarrantyid' />
                    <order attribute='hil_name' descending='false' />
                    <filter type='and'>
                        <condition attribute='statecode' operator='eq' value='0' />
                        <condition attribute='hil_customerasset' operator='eq' value='{entAsset.Id}' />
                    </filter>
                    <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='ab'>
                        <filter type='and'>
                            <condition attribute='hil_type' operator='in'>
                                <value>1</value>
                                <value>7</value>
                                <value>2</value>
                            </condition>
                        </filter>
                    </link-entity>
                  </entity>
                </fetch>";
                EntityCollection entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entCol.Entities.Count > 0)
                {
                    Entity _entUpdate = null;
                    foreach (Entity entLine in entCol.Entities)
                    {
                        _entUpdate = new Entity(entLine.LogicalName, entLine.Id);
                        _entUpdate["statecode"] = new OptionSetValue(1);
                        _entUpdate["statuscode"] = new OptionSetValue(2);
                        _service.Update(_entUpdate);
                    }
                }
            }
            catch
            {

            }
        }
        static void CreateZeroUnitWarrantyLine(IOrganizationService service, EntityReference erCustomer, EntityReference erCustomerasset, EntityReference erProductCatg, EntityReference erProductSubCatg, EntityReference erProductModel, string partdescription, EntityReference erWarrantytemplate, int _warrantyType)
        {
            try
            {
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='hil_unitwarranty'>
                    <attribute name='hil_unitwarrantyid' />
                    <order attribute='hil_name' descending='false' />
                    <filter type='and'>
                      <condition attribute='statecode' operator='eq' value='0' />
                      <condition attribute='hil_customerasset' operator='eq' value='{erCustomerasset.Id}' />
                    </filter>
                    <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='wrt'>
                      <attribute name='hil_type' />
                      <filter type='and'>
                        <condition attribute='hil_type' operator='eq' value='{_warrantyType}' />
                      </filter>
                    </link-entity>
                  </entity>
                </fetch>";
                EntityCollection entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entCol.Entities.Count == 0)
                {
                    hil_unitwarranty iSchWarranty = new hil_unitwarranty();
                    iSchWarranty.hil_CustomerAsset = erCustomerasset;
                    iSchWarranty.hil_productmodel = erProductCatg;
                    iSchWarranty.hil_productitem = erProductSubCatg;
                    iSchWarranty.hil_warrantystartdate = null;
                    iSchWarranty.hil_warrantyenddate = null;
                    iSchWarranty.hil_WarrantyTemplate = erWarrantytemplate;
                    iSchWarranty.hil_ProductType = new OptionSetValue(1);
                    if (erProductModel != null && erProductModel.Id != Guid.Empty)
                    {
                        iSchWarranty.hil_Part = erProductModel;
                    }
                    iSchWarranty["hil_partdescription"] = partdescription;
                    if (erCustomer != null && erCustomer.Id != Guid.Empty)
                    {
                        iSchWarranty.hil_customer = erCustomer;
                    }
                    service.Create(iSchWarranty);
                }
            }
            catch {
                
            }
        }

        static DateTime? CreateUnitWarrantyLine(IOrganizationService service, int period, int producttype, EntityReference erCustomer, EntityReference erCustomerasset, EntityReference erProductCatg, EntityReference erProductSubCatg, EntityReference erProductModel, string partdescription, DateTime? warrantystartdate, EntityReference erWarrantytemplate,int _warrantyType)
        {
            DateTime? WarrantyEnd = null;
            try
            {
                DateTime StartDate = new DateTime();
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='hil_unitwarranty'>
                    <attribute name='hil_warrantytemplate' />
                    <attribute name='hil_warrantystartdate' />
                    <attribute name='hil_warrantyenddate' />
                    <attribute name='hil_unitwarrantyid' />
                    <order attribute='hil_name' descending='false' />
                    <filter type='and'>
                      <condition attribute='statecode' operator='eq' value='0' />
                      <condition attribute='hil_customerasset' operator='eq' value='{erCustomerasset.Id}' />
                    </filter>
                    <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='wrt'>
                      <attribute name='hil_type' />
                      <filter type='and'>
                        <condition attribute='hil_type' operator='eq' value='{_warrantyType}' />
                      </filter>
                    </link-entity>
                  </entity>
                </fetch>";
                EntityCollection entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entCol.Entities.Count > 0)
                {
                    Guid _warrantyTemplateId = entCol.Entities[0].GetAttributeValue<EntityReference>("hil_warrantytemplate").Id;
                    DateTime _uwlStartDate = entCol.Entities[0].GetAttributeValue<DateTime>("hil_warrantystartdate").AddMinutes(330).Date;
                    DateTime _uwlEndDate = entCol.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330).Date;
                    if (_warrantyTemplateId == erWarrantytemplate.Id)
                    {
                        StartDate = Convert.ToDateTime(warrantystartdate).Date;
                        WarrantyEnd = StartDate.AddMonths(period).AddDays(-1);
                        if (_uwlStartDate != StartDate || _uwlEndDate != WarrantyEnd)
                        {
                            Entity _entUWL = new Entity(entCol.Entities[0].LogicalName, entCol.Entities[0].Id);
                            _entUWL["hil_warrantystartdate"] = StartDate;
                            _entUWL["hil_warrantyenddate"] = WarrantyEnd;
                            _service.Update(_entUWL);
                        }
                        return WarrantyEnd;
                    }
                    else
                    {
                        Entity _entUWL = new Entity(entCol.Entities[0].LogicalName, entCol.Entities[0].Id);
                        _entUWL["statecode"] = new OptionSetValue(1);
                        _entUWL["statuscode"] = new OptionSetValue(2);
                        _service.Update(_entUWL);
                    }
                }
                //QueryExpression qryExp = new QueryExpression(hil_unitwarranty.EntityLogicalName);
                //qryExp.ColumnSet = new ColumnSet("hil_warrantyenddate");
                //qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                //qryExp.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, erCustomerasset.Id);
                //qryExp.Criteria.AddCondition("hil_warrantytemplate", ConditionOperator.Equal, erWarrantytemplate.Id);
                //qryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                //entCol = service.RetrieveMultiple(qryExp);
                //if (entCol.Entities.Count == 0)
                //{

                hil_unitwarranty iSchWarranty = new hil_unitwarranty();
                iSchWarranty.hil_CustomerAsset = erCustomerasset;
                iSchWarranty.hil_productmodel = erProductCatg;
                iSchWarranty.hil_productitem = erProductSubCatg;
                iSchWarranty.hil_warrantystartdate = Convert.ToDateTime(warrantystartdate);
                StartDate = Convert.ToDateTime(warrantystartdate);
                WarrantyEnd = StartDate.AddMonths(period).AddDays(-1);
                iSchWarranty.hil_warrantyenddate = WarrantyEnd;
                iSchWarranty.hil_WarrantyTemplate = erWarrantytemplate;
                iSchWarranty.hil_ProductType = new OptionSetValue(1);
                if (erProductModel != null && erProductModel.Id != Guid.Empty)
                {
                    iSchWarranty.hil_Part = erProductModel;
                }
                iSchWarranty["hil_partdescription"] = partdescription;
                if (erCustomer != null && erCustomer.Id != Guid.Empty)
                {
                    iSchWarranty.hil_customer = erCustomer;
                }
                service.Create(iSchWarranty);
                //}
                //else
                //{
                //    WarrantyEnd = entCol.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate");
                //}
                return WarrantyEnd;
            }
            catch (Exception ex)
            {
                return WarrantyEnd;
            }
        }
        static bool ValidateWarrantySchemeLines(IOrganizationService service, EntityReference erWarrantytemplate, string _pinCode, DateTime? invDate)
        {
            bool retValue =false;
            DateTime invDateDT = Convert.ToDateTime(invDate);
            try
            {
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='hil_businessmapping'>
                    <attribute name='hil_salesoffice' />
                    <filter type='and'>
                        <condition attribute='hil_pincodename' operator='like' value='%{_pinCode}%' />
                        <condition attribute='statecode' operator='eq' value='0' />
                    </filter>
                    </entity>
                    </fetch>";
                EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entCol.Entities.Count > 0)
                {
                    EntityReference _entSO = entCol.Entities[0].GetAttributeValue<EntityReference>("hil_salesoffice");
                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_schemeline'>
                        <attribute name='hil_salesoffice' />
                        <attribute name='hil_isincluded' />
                        <filter type='and'>
                            <condition attribute='statecode' operator='eq' value='0' />
                            <condition attribute='hil_warrantytemplate' operator='eq' value='{erWarrantytemplate.Id}' />
                            <condition attribute='hil_fromdate' operator='on-or-before' value='{invDateDT.ToString("yyyy-MM-dd")}' />
                            <condition attribute='hil_todate' operator='on-or-after' value='{invDateDT.ToString("yyyy-MM-dd")}' />
                        </filter>
                        </entity>
                        </fetch>";
                    EntityCollection entColSO = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entColSO.Entities.Count > 0) {
                        bool _flag = entColSO.Entities[0].GetAttributeValue<bool>("hil_isincluded");
                        List<Entity> _lines = entColSO.Entities.ToList().Where(so => so.GetAttributeValue<EntityReference>("hil_salesoffice").Id.Equals(_entSO.Id)).ToList();
                        if (_lines.Count == 0) {//All Sales Office
                            _lines = entColSO.Entities.ToList().Where(so => so.GetAttributeValue<EntityReference>("hil_salesoffice").Id.Equals(new Guid("90503976-8fd1-ea11-a813-000d3af0563c"))).ToList();
                        }
                        retValue = (_lines == null ? !_flag : _flag);
                    }
                }
            }
            catch { }
            return retValue;
        }
        static DateTime? CreateSchemeUnitWarrantyLine(IOrganizationService service, int period, int producttype, EntityReference erCustomer, EntityReference erCustomerasset, EntityReference erProductCatg, EntityReference erProductSubCatg, EntityReference erProductModel, string partdescription, DateTime? warrantystartdate, EntityReference erWarrantytemplate, DateTime invDate,string _pinCode)
        {
            DateTime? WarrantyEnd = warrantystartdate;
            EntityReference erSalesOffice = null;

            try
            {
                QueryExpression qryExp = new QueryExpression(hil_unitwarranty.EntityLogicalName);
                qryExp.ColumnSet = new ColumnSet("hil_warrantyenddate");
                qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                qryExp.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, erCustomerasset.Id);
                qryExp.Criteria.AddCondition("hil_warrantytemplate", ConditionOperator.Equal, erWarrantytemplate.Id);
                qryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                EntityCollection entCol = service.RetrieveMultiple(qryExp);
                if (entCol.Entities.Count == 0)
                {
                    qryExp = new QueryExpression(msdyn_workorder.EntityLogicalName);
                    qryExp.ColumnSet = new ColumnSet("hil_salesoffice");
                    qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    qryExp.Criteria.AddCondition("hil_customerref", ConditionOperator.Equal, erCustomer.Id);
                    qryExp.AddOrder("createdon", OrderType.Descending);
                    qryExp.TopCount = 1;
                    entCol = service.RetrieveMultiple(qryExp);
                    if (entCol.Entities.Count > 0)
                    {
                        erSalesOffice = entCol.Entities[0].GetAttributeValue<EntityReference>("hil_salesoffice");
                        qryExp = new QueryExpression(hil_schemeline.EntityLogicalName);
                        qryExp.ColumnSet = new ColumnSet("hil_salesoffice");
                        qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qryExp.Criteria.AddCondition("hil_warrantytemplate", ConditionOperator.Equal, erWarrantytemplate.Id);
                        qryExp.Criteria.AddCondition("hil_salesoffice", ConditionOperator.Equal, erSalesOffice.Id);
                        qryExp.AddOrder("modifiedon", OrderType.Descending);
                        qryExp.TopCount = 1;
                        entCol = service.RetrieveMultiple(qryExp);
                        if (entCol.Entities.Count == 0)
                        {
                            erSalesOffice = null;
                        }
                    }
                    if (erSalesOffice == null)
                    {
                        qryExp = new QueryExpression(hil_address.EntityLogicalName);
                        qryExp.ColumnSet = new ColumnSet("hil_salesoffice");
                        qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qryExp.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, erCustomer.Id);
                        qryExp.AddOrder("modifiedon", OrderType.Descending);
                        qryExp.TopCount = 1;
                        entCol = service.RetrieveMultiple(qryExp);
                        if (entCol.Entities.Count > 0)
                        {
                            erSalesOffice = entCol.Entities[0].GetAttributeValue<EntityReference>("hil_salesoffice");
                        }
                    }
                    if (erSalesOffice != null)
                    {
                        qryExp = new QueryExpression(hil_schemeline.EntityLogicalName);
                        qryExp.ColumnSet = new ColumnSet("hil_salesoffice", "hil_fromdate", "hil_todate");
                        qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qryExp.Criteria.AddCondition("hil_warrantytemplate", ConditionOperator.Equal, erWarrantytemplate.Id);
                        qryExp.Criteria.AddCondition("hil_salesoffice", ConditionOperator.Equal, erSalesOffice.Id);
                        qryExp.AddOrder("modifiedon", OrderType.Descending);
                        qryExp.TopCount = 1;
                        entCol = service.RetrieveMultiple(qryExp);
                        if (entCol.Entities.Count > 0)
                        {
                            if (invDate >= entCol.Entities[0].GetAttributeValue<DateTime>("hil_fromdate") && invDate <= entCol.Entities[0].GetAttributeValue<DateTime>("hil_todate"))
                            {
                                DateTime StartDate = new DateTime();
                                hil_unitwarranty iSchWarranty = new hil_unitwarranty();
                                iSchWarranty.hil_CustomerAsset = erCustomerasset;
                                iSchWarranty.hil_productmodel = erProductCatg;
                                iSchWarranty.hil_productitem = erProductSubCatg;
                                iSchWarranty.hil_warrantystartdate = Convert.ToDateTime(warrantystartdate);
                                StartDate = Convert.ToDateTime(warrantystartdate);
                                WarrantyEnd = StartDate.AddMonths(period).AddDays(-1);
                                iSchWarranty.hil_warrantyenddate = WarrantyEnd;
                                iSchWarranty.hil_WarrantyTemplate = erWarrantytemplate;
                                iSchWarranty.hil_ProductType = new OptionSetValue(1);
                                if (erProductModel != null && erProductModel.Id != Guid.Empty)
                                {
                                    iSchWarranty.hil_Part = erProductModel;
                                }
                                iSchWarranty["hil_partdescription"] = partdescription;
                                if (erCustomer != null && erCustomer.Id != Guid.Empty)
                                {
                                    iSchWarranty.hil_customer = erCustomer;
                                }
                                service.Create(iSchWarranty);
                            }
                        }
                    }
                }
                else
                {
                    WarrantyEnd = entCol.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate");
                }
                return WarrantyEnd;
            }
            catch (Exception ex)
            {
                return WarrantyEnd;
            }
        }
        static IOrganizationService ConnectToCRM(string connectionString)
        {
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);
                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    Console.WriteLine(((CrmServiceClient)service).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }
            return service;

        }

        static ApplicableWarrantyDTO ApplicationOfCustomerAssetWarranty(IOrganizationService _service, ApplicableWarrantyDTO _inputParam)
        {
            ApplicableWarrantyDTO _outputParam = new ApplicableWarrantyDTO();
            EntityReference _warrantyTemplate = null;
            DateTime _unitWarrStartDate;
            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_unitwarranty'>
                <attribute name='hil_warrantystartdate' />
                <attribute name='hil_warrantyenddate' />
                <attribute name='hil_warrantytemplate' />
                <filter type='and'>
                    <condition attribute='hil_customerasset' operator='eq' value='{_inputParam.customerAssetId}' />
                    <condition attribute='statecode' operator='eq' value='0' />
                    <condition attribute='hil_warrantyenddate' operator='on-or-after' value='{_inputParam.jobCreatedOn}' />
                </filter>
                <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='wt'>
                    <attribute name='hil_type' />
                    <link-entity name='hil_warrantytype' from='hil_warrantytypeid' to='hil_warrantytypeindex' visible='false' link-type='outer' alias='wtt'>
                        <attribute name='hil_executionindex' />
                        <order attribute='hil_executionindex' descending='false' />
                    </link-entity>
                </link-entity>
                </entity>
                </fetch>";

            EntityCollection entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));

            if (entcoll.Entities.Count == 0)
            {
                _outputParam.AsssetWarrantyStatus = new OptionSetValue(2); //OUT Warranty
                _outputParam.LaborWarrantyStatus = new OptionSetValue(2); //OUT Warranty
                _outputParam.SparePartWarrantyStatus = new OptionSetValue(2); //OUT Warranty
            }
            else
            {
                DateTime _warrantyEndDate = entcoll.Entities.Max(x => x.GetAttributeValue<DateTime>("hil_warrantyenddate")).Date;
                OptionSetValue _warrantyType = (OptionSetValue)entcoll.Entities[0].GetAttributeValue<AliasedValue>("wt.hil_type").Value;
                _warrantyTemplate = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_warrantytemplate");
                _unitWarrStartDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_warrantystartdate");

                int _warrantyTypeInt = _warrantyType.Value;

                if (_warrantyTypeInt == 1 || _warrantyTypeInt == 7 || _warrantyTypeInt == 3)//{Warranty Type: Standard, Special Scheme, AMC}
                {
                    _outputParam.AsssetWarrantyStatus = new OptionSetValue(1); //IN Warranty
                    _outputParam.LaborWarrantyStatus = new OptionSetValue(1); //IN Warranty
                    _outputParam.SparePartWarrantyStatus = new OptionSetValue(1); //IN Warranty
                    _outputParam.warrantyEndDate = _warrantyEndDate;
                }
                else //{Warranty Type: Extended}
                {
                    _outputParam.AsssetWarrantyStatus = new OptionSetValue(2); //OUT Warranty
                    _outputParam.LaborWarrantyStatus = new OptionSetValue(2); //OUT Warranty

                    if (_inputParam.replacesPartId == Guid.Empty)
                    {
                        _outputParam.SparePartWarrantyStatus = new OptionSetValue(2); //OUT Warranty
                    }
                    else
                    {
                        string _spareFamilyCond = string.Empty;
                        Entity _entSpareFamily = _service.Retrieve("product", _inputParam.replacesPartId, new ColumnSet("hil_sparepartfamily"));
                        Guid _spareFamily = Guid.Empty;
                        if (_entSpareFamily != null)
                        {
                            _spareFamily = _entSpareFamily.GetAttributeValue<EntityReference>("hil_sparepartfamily").Id;
                            _spareFamilyCond = $"<condition attribute='hil_partfamily' operator='eq' value='{_spareFamily}' />";
                        }
                        TimeSpan difference = (_inputParam.jobCreatedOn - _unitWarrStartDate);
                        double _jobMonth = Math.Round((difference.Days * 1.0 / 30.42), 0);

                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
                            <entity name='hil_part'>
                            <attribute name='hil_includedinwarranty' />
                            <order attribute='hil_partcode' descending='true' />
                            <order attribute='hil_partfamily' descending='false' />
                            <filter type='and'>
                                <condition attribute='hil_warrantytemplateid' operator='eq' value='{_warrantyTemplate.Id}' />
                                <filter type='or'>{_spareFamilyCond}
                                    <condition attribute='hil_partcode' operator='eq' value='{_inputParam.replacesPartId}' />
                                </filter>
                                <condition attribute='hil_validfrommonths' operator='ge' value='{_jobMonth}' />
                                <condition attribute='hil_validtomonths' operator='le' value='{_jobMonth}' />
                                <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            </entity>
                            </fetch>";

                        entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entcoll.Entities.Count > 0)
                        {
                            _outputParam.SparePartWarrantyStatus = new OptionSetValue(1); //IN Warranty
                        }
                    }
                }
            }
            return _outputParam;
        }
    }
    class WarrantyStatusDTO
    {
        public Guid customerAssetsId { get; set; }
        public OptionSetValue warrantyStatus { get; set; }
        public OptionSetValue warrantySubStatus { get; set; }
        public DateTime? warrantyEndDate { get; set; }
        //public DateTime? partWarrantyEndDate { get; set; }
    }
    static class WarrantyApplicableOn
    {
        public static string productSubcategory { get; } = "910590000";
        public static string model { get; } = "910590001";
        public static string serialNumber { get; } = "910590002";
    }
    class ApplicableWarrantyDTO
    {
        public DateTime jobCreatedOn { get; set; }
        public Guid customerAssetId { get; set; }
        public Guid replacesPartId { get; set; }
        public OptionSetValue AsssetWarrantyStatus { get; set; }
        public OptionSetValue LaborWarrantyStatus { get; set; }
        public OptionSetValue SparePartWarrantyStatus { get; set; }
        public DateTime? warrantyEndDate { get; set; }
    }
    static class WarrantyType
    {
        public static int standard { get; } = 1;
        public static int extended { get; } = 2;
        public static int amc { get; } = 3;
        public static int specialScheme { get; } = 7;
        public static int zeroWarranty { get; } = 8;
    }
}
