using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;
using Microsoft.Xrm.Sdk.Messages;

namespace AE01.Miscellaneous.Production_Support
{
    public class Warranty
    {

        private IOrganizationService? service;
        public Warranty(IOrganizationService _service)
        {
            this.service = _service;
        }
        public void ToTest()
        {
            string fetch = @"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                              <entity name=""msdyn_customerasset"">
                                <attribute name=""createdon"" />
                                <attribute name=""msdyn_product"" />
                                <attribute name=""msdyn_name"" />
                                <attribute name=""hil_productsubcategorymapping"" />
                                <attribute name=""hil_productcategory"" />
                                <attribute name=""msdyn_customerassetid"" />
                                <order attribute=""createdon"" descending=""true"" />
                                <filter type=""and"">
                                  <condition attribute=""createdon"" operator=""today"" />
                                  <condition attribute=""hil_invoiceavailable"" operator=""eq"" value=""0"" />
                                </filter>
                              </entity>
                            </fetch>";
            EntityCollection tofix = service.RetrieveMultiple(new FetchExpression(fetch));
            foreach (Entity it in tofix.Entities)
            {

                ExecuteWarrantyEngine(service, it.Id);
            }

            Guid cusstomerAssetId = new Guid("af0d21bf-cf11-ef11-9f89-7c1e520eb873");
        }

        private void ExecuteWarrantyEngine(IOrganizationService _service, Guid entityId)
        {
            try
            {
                Entity _entCustomerAsset = _service.Retrieve("msdyn_customerasset", entityId, new ColumnSet("hil_invoiceavailable", "hil_invoicedate", "hil_productsubcategory", "hil_productcategory", "msdyn_product", "msdyn_name", "statuscode", "hil_branchheadapprovalstatus", "hil_customer", "hil_retailerpincode"));
                if (_entCustomerAsset != null)
                {
                    DateTime? _invoiceDate = null;
                    string _invoiceDateStr = string.Empty;
                    string _serialNumer = _entCustomerAsset.Contains("msdyn_name") ? _entCustomerAsset.GetAttributeValue<string>("msdyn_name") : string.Empty;
                    bool _invoiceAvailable = _entCustomerAsset.Contains("hil_invoiceavailable") ? _entCustomerAsset.GetAttributeValue<bool>("hil_invoiceavailable") : false;
                    string _pinCode = _entCustomerAsset.Contains("hil_retailerpincode") ? _entCustomerAsset.GetAttributeValue<string>("hil_retailerpincode") : string.Empty;

                    if (_entCustomerAsset.Contains("hil_invoicedate"))
                    {
                        _invoiceDate = _entCustomerAsset.GetAttributeValue<DateTime>("hil_invoicedate").Date;
                        _invoiceDateStr = Convert.ToDateTime(_invoiceDate).ToString("yyyy-MM-dd");
                    }
                    if (!_invoiceAvailable)
                    {
                        //Entity _custAsset = new Entity(_entCustomerAsset.LogicalName, _entCustomerAsset.Id);
                        //_custAsset["hil_invoiceavailable"] = true;
                        //_service.Update(_custAsset);
                        //_invoiceAvailable = true;

                        UpdateCustomerAssetWarrantyStatus(_service,
                                            new WarrantyStatusDTO()
                                            {
                                                customerAssetsId = _entCustomerAsset.Id,
                                                warrantyEndDate = new DateTime(1900, 1, 1),
                                                warrantyStatus = new OptionSetValue(3),//{key:"Warranty Void",Value:3}
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
                    if (!_invoiceAvailable || !_isAssetApproved || _productSubcategory == null || _productCategory == null || _modelNumer == null || _customer == null)
                    {
                        InactiveSystemWarranties(_service, _entCustomerAsset);
                    }
                    else
                    {
                        string _fetchXMLQuery = string.Empty;
                        EntityCollection entCol = null;

                        string[] _applicableOn = new string[] { WarrantyApplicableOn.serialNumber, WarrantyApplicableOn.model, WarrantyApplicableOn.productSubcategory };

                        string _filterCondWarrantyHeader = string.Empty;
                        string _filterCondWarrantyLine = string.Empty;

                        endDate = _invoiceDate;

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

                            int[] _warrantyTypes = new int[] { WarrantyType.standard, WarrantyType.specialScheme, WarrantyType.extended };
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
                                    //endDate = _invoiceDate;
                                    foreach (Entity entWrty in entCol.Entities)
                                    {
                                        //OptionSetValue _warrantyType = entWrty.GetAttributeValue<OptionSetValue>("hil_type");
                                        if (_warrantyType == WarrantyType.zeroWarranty)//Zero Warranty
                                        {
                                            UpdateCustomerAssetWarrantyStatus(_service,
                                            new WarrantyStatusDTO()
                                            {
                                                customerAssetsId = _entCustomerAsset.Id,
                                                warrantyEndDate = null,
                                                warrantyStatus = new OptionSetValue(3),//{key:"Warranty Void",Value:3}
                                                warrantySubStatus = null
                                            });
                                            InactiveSystemWarranties(_service, _entCustomerAsset);
                                            CreateZeroUnitWarrantyLine(_service, entWrty.GetAttributeValue<int>("hil_warrantyperiod"), _customer, _entCustomerAsset.ToEntityReference(), _productCategory, _productSubcategory, _modelNumer, _modelNumer.Name, endDate, entCol.Entities[0].ToEntityReference(), _warrantyType);
                                        }
                                        else
                                        {
                                            if (_warrantyType == WarrantyType.specialScheme)//Special Scheme
                                            {

                                            }
                                            else
                                            {
                                                endDate = CreateUnitWarrantyLine(_service, entWrty.GetAttributeValue<int>("hil_warrantyperiod"), 1, _customer, _entCustomerAsset.ToEntityReference(), _productCategory, _productSubcategory, _modelNumer, _modelNumer.Name, endDate, entWrty.ToEntityReference(), _warrantyType);
                                                endDate = Convert.ToDateTime(endDate).AddDays(1);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    CalculateAssetWarrantyStatus(_service, _entCustomerAsset, _isAssetApproved);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }

        }

        private void InactiveSystemWarranties(IOrganizationService _service, Entity entAsset)
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

        private void UpdateCustomerAssetWarrantyStatus(IOrganizationService _service, WarrantyStatusDTO _assetWarrantyStatus)
        {
            try
            {
                Entity entCustAsset = new Entity("msdyn_customerasset");
                entCustAsset.Id = _assetWarrantyStatus.customerAssetsId;
                entCustAsset["hil_warrantytilldate"] = _assetWarrantyStatus.warrantyEndDate;
                entCustAsset["hil_warrantystatus"] = _assetWarrantyStatus.warrantyStatus;
                entCustAsset["hil_warrantysubstatus"] = _assetWarrantyStatus.warrantySubStatus;
                _service.Update(entCustAsset);
            }
            catch (Exception ex)
            {

            }
        }

        private void CreateZeroUnitWarrantyLine(IOrganizationService _service, int period, EntityReference erCustomer, EntityReference erCustomerasset, EntityReference erProductCatg, EntityReference erProductSubCatg, EntityReference erProductModel, string partdescription, DateTime? warrantystartdate, EntityReference erWarrantytemplate, int _warrantyType)
        {
            DateTime? WarrantyEnd = null;
            try
            {
                DateTime StartDate = new DateTime();
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

                    //hil_unitwarranty iSchWarranty = new hil_unitwarranty();
                    //iSchWarranty.hil_CustomerAsset = erCustomerasset;
                    //iSchWarranty.hil_productmodel = erProductCatg;
                    //iSchWarranty.hil_productitem = erProductSubCatg;
                    //iSchWarranty.hil_warrantystartdate = Convert.ToDateTime(warrantystartdate);
                    //StartDate = Convert.ToDateTime(warrantystartdate);
                    //WarrantyEnd = StartDate.AddMonths(period).AddDays(-1);
                    //iSchWarranty.hil_warrantyenddate = WarrantyEnd;
                    //iSchWarranty.hil_WarrantyTemplate = erWarrantytemplate;
                    //iSchWarranty.hil_ProductType = new OptionSetValue(1);
                    //if (erProductModel != null && erProductModel.Id != Guid.Empty)
                    //{
                    //    iSchWarranty.hil_Part = erProductModel;
                    //}
                    //iSchWarranty["hil_partdescription"] = partdescription;
                    //if (erCustomer != null && erCustomer.Id != Guid.Empty)
                    //{
                    //    iSchWarranty.hil_customer = erCustomer;
                    //}
                    //_service.Create(iSchWarranty);
                }
            }
            catch
            {

            }
        }

        private DateTime? CreateUnitWarrantyLine(IOrganizationService _service, int period, int producttype, EntityReference erCustomer, EntityReference erCustomerasset, EntityReference erProductCatg, EntityReference erProductSubCatg, EntityReference erProductModel, string partdescription, DateTime? warrantystartdate, EntityReference erWarrantytemplate, int _warrantyType)
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
                    DateTime _uwlStartDate = entCol.Entities[0].GetAttributeValue<DateTime>("hil_warrantystartdate").Date;
                    DateTime _uwlEndDate = entCol.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate").Date;
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
                //hil_unitwarranty iSchWarranty = new hil_unitwarranty();
                //iSchWarranty.hil_CustomerAsset = erCustomerasset;
                //iSchWarranty.hil_productmodel = erProductCatg;
                //iSchWarranty.hil_productitem = erProductSubCatg;
                //iSchWarranty.hil_warrantystartdate = Convert.ToDateTime(warrantystartdate);
                //StartDate = Convert.ToDateTime(warrantystartdate);
                //WarrantyEnd = StartDate.AddMonths(period).AddDays(-1);
                //iSchWarranty.hil_warrantyenddate = WarrantyEnd;
                //iSchWarranty.hil_WarrantyTemplate = erWarrantytemplate;
                //iSchWarranty.hil_ProductType = new OptionSetValue(1);
                //if (erProductModel != null && erProductModel.Id != Guid.Empty)
                //{
                //    iSchWarranty.hil_Part = erProductModel;
                //}
                //iSchWarranty["hil_partdescription"] = partdescription;
                //if (erCustomer != null && erCustomer.Id != Guid.Empty)
                //{
                //    iSchWarranty.hil_customer = erCustomer;
                //}
                //_service.Create(iSchWarranty);
                return WarrantyEnd;
            }
            catch (Exception ex)
            {
                return WarrantyEnd;
            }
        }

        static void CalculateAssetWarrantyStatus(IOrganizationService _service, Entity entCA, bool _isAssetApproved)
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
                    if (_isAssetApproved)
                        _ent["hil_warrantystatus"] = new OptionSetValue(2); //Out Warranty
                    else
                        _ent["hil_warrantystatus"] = new OptionSetValue(4); //Approval Pending
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
    }
    public class WarrantyStatusDTO
    {
        public Guid customerAssetsId { get; set; }
        public OptionSetValue warrantyStatus { get; set; }
        public OptionSetValue warrantySubStatus { get; set; }
        public DateTime? warrantyEndDate { get; set; }
        //public DateTime? partWarrantyEndDate { get; set; }
    }
    public class ApplicableWarrantyDTO
    {
        public DateTime jobCreatedOn { get; set; }
        public Guid customerAssetId { get; set; }
        public Guid replacesPartId { get; set; }
        public OptionSetValue AsssetWarrantyStatus { get; set; }
        public OptionSetValue LaborWarrantyStatus { get; set; }
        public OptionSetValue SparePartWarrantyStatus { get; set; }
        public DateTime? warrantyEndDate { get; set; }
    }
    public class WarrantyApplicableOn
    {
        public static string productSubcategory { get; } = "910590000";
        public static string model { get; } = "910590001";
        public static string serialNumber { get; } = "910590002";
    }
    public class WarrantyType
    {
        public static int standard { get; } = 1;
        public static int extended { get; } = 2;
        public static int amc { get; } = 3;
        public static int specialScheme { get; } = 7;
        public static int zeroWarranty { get; } = 8;
    }






}


//using Microsoft.Xrm.Sdk.Query;
//using Microsoft.Xrm.Sdk;
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using Microsoft.Xrm.Sdk.Messages;

namespace BookMASTER.ACTIONS
{
    public class RMAAssociateAndLoadInventoryJournal : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            bool _statusCode = false;

            try
            {
                string _RMAId = context.InputParameters.Contains("RMA_Guid") ? context.InputParameters["RMA_Guid"].ToString() : null;
                if (string.IsNullOrEmpty(_RMAId))
                {
                    context.OutputParameters["StatusMessage"] = "First Save RMA Record to generate RMA_ID";
                    context.OutputParameters["StatusCode"] = _statusCode;
                    return;
                }
                else
                {
                    Entity _InventoryRMA = service.Retrieve("hil_inventoryrma", new Guid(_RMAId), new ColumnSet("hil_franchise", "hil_warehouse", "hil_returntype"));
                    if (_InventoryRMA != null)
                    {
                        EntityReference _franchise = _InventoryRMA.GetAttributeValue<EntityReference>("hil_franchise");
                        EntityReference _warehouse = _InventoryRMA.GetAttributeValue<EntityReference>("hil_warehouse");
                        EntityReference _returntype = _InventoryRMA.GetAttributeValue<EntityReference>("hil_returntype");

                        string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                              <entity name='hil_inventoryproductjournal'>
                                                <attribute name='hil_inventoryproductjournalid' />
                                                <attribute name='hil_name' />
                                                <attribute name='createdon' />
                                                <order attribute='hil_name' descending='false' />
                                                <filter type='and'>
                                                  <condition attribute='statecode' operator='eq' value='0' />
                                                  <condition attribute='hil_transactiontype' operator='eq' value='3' />
                                                  <condition attribute='hil_franchise' operator='eq' value='{_franchise.Id}' />
                                                  <condition attribute='hil_warehouse' operator='eq' value='{_warehouse.Id}' />
                                                  <condition attribute='hil_rmatype' operator='eq' value='{_returntype.Id}' />
                                                  <condition attribute='hil_rma' operator='null' />
                                                </filter>
                                              </entity>
                                            </fetch>";
                        EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entCol.Entities.Count > 0)
                        {
                            int batchSize = 1000;

                            for (int i = 0; i < entCol.Entities.Count; i += batchSize)
                            {
                                var batch = entCol.Entities.Skip(i).Take(batchSize).ToList(); 

                                ExecuteMultipleRequest requestWithResults = new ExecuteMultipleRequest()
                                {
                                    // Assign settings that define execution behavior: continue on error, return responses. 
                                    Settings = new ExecuteMultipleSettings()
                                    {
                                        ContinueOnError = false,
                                        ReturnResponses = true
                                    },
                                    // Create an empty organization request collection.
                                    Requests = new OrganizationRequestCollection()
                                };
                                Entity InventoryJournal = null;
                                foreach (var entity in batch)
                                {
                                    InventoryJournal = new Entity("hil_inventoryproductjournal", entity.Id);
                                    InventoryJournal["hil_rma"] = new EntityReference("hil_inventoryrma", new Guid(_RMAId));
                                    UpdateRequest updateRequest = new UpdateRequest() { Target = InventoryJournal };
                                    requestWithResults.Requests.Add(updateRequest);
                                }
                                try
                                {
                                    ExecuteMultipleResponse responseWithResults = (ExecuteMultipleResponse)service.Execute(requestWithResults);
                                    _statusCode = true;
                                    context.OutputParameters["StatusMessage"] = "Success - Inventory Journal RMA updated";
                                }
                                catch (Exception ex)
                                {
                                    context.OutputParameters["StatusMessage"] = $"ERROR! - {ex.Message}";
                                }
                            }
                        }
                        else
                        {
                            context.OutputParameters["StatusMessage"] = "NO Active Inventory Journal Lines Found";
                        }
                    }
                }
                context.OutputParameters["StatusCode"] = _statusCode;
            }
            catch (Exception ex)
            {
                context.OutputParameters["StatusMessage"] = ex.Message;
            }
        }
    }
}