using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.Configuration;
using System.Net;
using System.Text;
using System.ServiceModel.Description;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using Microsoft.Xrm.Sdk.Query;
using System.IO;
using Newtonsoft.Json;

namespace CustomerAssetWarrantyRefresh
{
    public class Program
    {
        #region Global Varialble declaration
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";
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
                string _prodCatgId = ConfigurationManager.AppSettings["ProductCatg"].ToString();
                string _prodSubcatgId = ConfigurationManager.AppSettings["ProductSubCatg"].ToString();
                RefreshAssetWarranty(_prodCatgId, _prodSubcatgId);
            }
        }

        static void RefreshAssetWarranty(string _productCatg, string _productSubCatg)
        {
            Guid CustomerAssetId = Guid.Empty;
            try
            {
                int _rowCount = 1, _totalRowCount = 0;
                string _serialNumber = string.Empty;
                int _pageSize = 1000;
                EntityCollection entcoll = null;
                while (true)
                {
                    string _condition = string.Empty;

                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='{_pageSize}'>
                          <entity name='msdyn_customerasset'>
                            <attribute name='createdon' />
                            <attribute name='msdyn_product' />
                            <attribute name='msdyn_name' />
                            <attribute name='hil_productsubcategory' />
                            <attribute name='hil_productcategory' />
                            <attribute name='msdyn_customerassetid' />
                            <attribute name='statuscode' />
                            <attribute name='hil_branchheadapprovalstatus' />
                            <attribute name='hil_invoicedate' />
                            <attribute name='hil_invoiceavailable' />
                            <order attribute='msdyn_name' descending='false' />
                            <filter type='and'>
                              <condition attribute='msdyn_name' operator='not-null' />
                              <condition attribute='hil_modeofpayment' operator='ne' value='WR' />
                              <condition attribute='hil_productcategory' operator='eq' value='{_productCatg}' />
                              <condition attribute='hil_productsubcategory' operator='eq' value='{_productSubCatg}' />{_condition}
                            </filter>
                          </entity>
                        </fetch>";
                    entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entcoll.Entities.Count == 0) { Console.WriteLine("Batch Completed... "); break; }
                    _totalRowCount += entcoll.Entities.Count;
                    foreach (Entity entCA in entcoll.Entities)
                    {
                        try
                        {
                            Entity _ent = new Entity(entCA.LogicalName, entCA.Id);

                            _serialNumber = entCA.GetAttributeValue<string>("msdyn_name");
                            Console.WriteLine($"Asset# {_serialNumber} Row Count: {_rowCount++}/{_totalRowCount}");
                            string _processDate = DateTime.Now.ToString("yyyy-MM-dd");
                            _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
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
                                _ent["hil_warrantystatus"] = new OptionSetValue(2); //Out Warranty
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

                                EntityCollection entcoll4 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                if (entcoll4.Entities.Count > 0)
                                    _ent["hil_warrantytilldate"] = entcoll4.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330);
                                else
                                    _ent["hil_warrantytilldate"] = null;
                            }
                            else
                            {
                                DateTime _warrantyEndDate = entcoll3.Entities.Max(x => x.GetAttributeValue<DateTime>("hil_warrantyenddate")).AddMinutes(330);
                                OptionSetValue _warrantyType = (OptionSetValue)entcoll3.Entities[0].GetAttributeValue<AliasedValue>("wt.hil_type").Value;
                                _ent["hil_warrantystatus"] = new OptionSetValue(1); //In Warranty
                                int _warrantySubStatus = _warrantyType.Value;

                                if (_warrantySubStatus == 5) { _warrantySubStatus = 2; }
                                else if (_warrantySubStatus == 7) { _warrantySubStatus = 3; }
                                else if (_warrantySubStatus == 3) { _warrantySubStatus = 4; }

                                _ent["hil_warrantysubstatus"] = new OptionSetValue(_warrantySubStatus);
                                _ent["hil_warrantytilldate"] = _warrantyEndDate;
                            }
                            _ent["hil_modeofpayment"] = "WR";
                            _service.Update(_ent);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(CustomerAssetId.ToString() + " : " + ex.Message);
            }
        }
        static void RefreshAssetWarrantyStatus()
        {
            int totalRecords = 0;
            int recordCnt = 0;
            try
            {
                QueryExpression Query = new QueryExpression("msdyn_customerasset");
                Query.ColumnSet = new ColumnSet("hil_warrantystatus", "msdyn_customerassetid");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression("hil_warrantystatus", ConditionOperator.Equal, 1)); //IN Warranty
                Query.Criteria.AddCondition(new ConditionExpression("hil_warrantytilldate", ConditionOperator.OnOrBefore, DateTime.Now.AddDays(-1))); //AMC Call
                Query.AddOrder("createdon", OrderType.Descending);
                EntityCollection ec = _service.RetrieveMultiple(Query);
                if (ec.Entities.Count > 0)
                {
                    totalRecords = ec.Entities.Count;
                    recordCnt = 1;
                    foreach (Entity ent in ec.Entities)
                    {
                        Console.WriteLine("Looping : " + recordCnt.ToString() + "/" + totalRecords.ToString());
                        ent["hil_warrantystatus"] = new OptionSetValue(2);
                        _service.Update(ent);
                        recordCnt += 1;
                    }
                }
                else
                {
                    Console.WriteLine("CustomerAssetWarrantyRefresh.Program.Main.RefreshAssetWarrantyStatus :: No record found to sync");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ACustomerAssetWarrantyRefresh.Program.Main.RefreshAssetWarrantyStatus :: Error While Loading App Settings:" + ex.Message.ToString());
            }
        }

        static void UpdateAssetWarrantyStatus()
        {
            int totalRecords = 0;
            int recordCnt = 0;
            try
            {
                QueryExpression Query = new QueryExpression(msdyn_customerasset.EntityLogicalName);
                Query.ColumnSet = new ColumnSet("statuscode", "hil_warrantytilldate","msdyn_name", "hil_invoicedate", "msdyn_customerassetid", "hil_productsubcategorymapping", "hil_productsubcategory");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 910590001)); 
                Query.Criteria.AddCondition(new ConditionExpression("hil_warrantytilldate", ConditionOperator.Null)); 
                Query.AddOrder("createdon", OrderType.Descending);
                EntityCollection ec = _service.RetrieveMultiple(Query);
                if (ec.Entities.Count > 0)
                {
                    totalRecords = ec.Entities.Count;
                    recordCnt = 1;
                    Entity entTemp;
                    foreach (Entity ent in ec.Entities)
                    {
                        if (ent.Attributes.Contains("hil_invoicedate"))
                        {
                            ent["hil_createwarranty"] = false;
                            _service.Update(ent);

                            entTemp = _service.Retrieve("hil_stagingdivisonmaterialgroupmapping", ent.GetAttributeValue<EntityReference>("hil_productsubcategorymapping").Id, new ColumnSet("hil_productsubcategorymg"));
                            if (entTemp != null) {
                                if (entTemp.GetAttributeValue<EntityReference>("hil_productsubcategorymg").Id != ent.GetAttributeValue<EntityReference>("hil_productsubcategory").Id) {
                                    ent["hil_productsubcategory"] = entTemp.GetAttributeValue<EntityReference>("hil_productsubcategorymg");
                                }
                            }
                            ent["hil_createwarranty"] = true;
                            _service.Update(ent);
                        }
                        else {
                            ent["hil_warrantytilldate"] = new DateTime(1900, 1, 1);
                            ent["hil_warrantystatus"] = new OptionSetValue(2); //OUT
                            _service.Update(ent);
                        }
                        Console.WriteLine(ent.GetAttributeValue<string>("msdyn_name").ToString() + " : " + recordCnt.ToString() + "/" + totalRecords.ToString());
                        recordCnt += 1;
                    }
                }
                else
                {
                    Console.WriteLine("CustomerAssetWarrantyRefresh.Program.Main.RefreshAssetWarrantyStatus :: No record found to sync");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ACustomerAssetWarrantyRefresh.Program.Main.RefreshAssetWarrantyStatus :: Error While Loading App Settings:" + ex.Message.ToString());
            }
        }

        static void RefreshAssetUnitWarranty() {
            string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='msdyn_customerasset'>
            <attribute name='hil_invoicedate' />
            <attribute name='hil_invoiceavailable' />
            <attribute name='statuscode' />
            <attribute name='hil_productcategory' />
            <attribute name='hil_productsubcategory' />
            <attribute name='msdyn_product' />
            <attribute name='hil_modelname' />
            <attribute name='msdyn_name' />
            <attribute name='hil_customer' />
                <order attribute='createdon' descending='true' />
                <filter type='and'>
                <condition attribute='hil_productcategory' operator='eq' value='{A7A5049B-16FA-E811-A94C-000D3AF06091}' />
                <condition attribute='hil_warrantysubstatus' operator='null' />
                <condition attribute='createdon' operator='on-or-after' value='2020-01-01' />
                <condition attribute='createdon' operator='on-or-before' value='2020-01-31' />
                <filter type='or'>
                <condition attribute='statuscode' operator='eq' value='910590001' />
                <condition attribute='hil_branchheadapprovalstatus' operator='eq' value='1' />
                </filter>
                <condition attribute='hil_invoiceavailable' operator='eq' value='1' />
                </filter>
              </entity>
            </fetch>";
            int rec = 0;
            int Total =0;
            while (true)
            {
                EntityCollection ec = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                Total += ec.Entities.Count;
                if (ec.Entities.Count == 0) { break; }
                foreach (Entity CustAsst in ec.Entities)
                {
                    EntityReference erProdCgry = new EntityReference("product");
                    EntityReference erProdSubCategory = new EntityReference("product");
                    EntityReference erProdModelCode = new EntityReference("product");
                    EntityReference erCustomer = new EntityReference("contact");
                    string pdtModelName = string.Empty;
                    DateTime? invDate = null;
                    DateTime? endDate = null;
                    bool invoiceAvailable = false;
                    int? warrantySubstatus = null;

                    if (CustAsst.Attributes.Contains("hil_invoiceavailable") && CustAsst.Attributes.Contains("hil_invoicedate") && CustAsst.Attributes.Contains("hil_productcategory"))
                    {
                        invoiceAvailable = (bool)CustAsst["hil_invoiceavailable"];
                        if (invoiceAvailable)
                        {
                            QueryExpression queryExp1 = new QueryExpression("hil_unitwarranty");
                            queryExp1.ColumnSet = new ColumnSet("hil_unitwarrantyid");
                            queryExp1.Criteria = new FilterExpression(LogicalOperator.And);
                            queryExp1.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, CustAsst.Id);
                            queryExp1.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                            queryExp1.Criteria.AddCondition("hil_customer", ConditionOperator.Null); //Old Unit Warranties
                            EntityCollection entCol1 = _service.RetrieveMultiple(queryExp1);
                            if (entCol1.Entities.Count > 0)
                            {
                                foreach (Entity ent in entCol1.Entities)
                                {
                                    SetStateRequest setStateRequest = new SetStateRequest()
                                    {
                                        EntityMoniker = new EntityReference
                                        {
                                            Id = ent.Id,
                                            LogicalName = "hil_unitwarranty",
                                        },
                                        State = new OptionSetValue(1), //Inactive
                                        Status = new OptionSetValue(2) //Inactive
                                    };
                                    _service.Execute(setStateRequest);
                                }
                            }

                            invDate = (DateTime)CustAsst["hil_invoicedate"];
                            erProdCgry = (EntityReference)CustAsst["hil_productcategory"];

                            if (CustAsst.Attributes.Contains("hil_productsubcategory"))
                            {
                                erProdSubCategory = (EntityReference)CustAsst["hil_productsubcategory"];
                            }
                            if (CustAsst.Attributes.Contains("msdyn_product"))
                            {
                                erProdModelCode = (EntityReference)CustAsst["msdyn_product"];
                            }
                            if (CustAsst.Attributes.Contains("hil_modelname"))
                            {
                                pdtModelName = CustAsst["hil_modelname"].ToString();
                            }
                            if (CustAsst.Attributes.Contains("hil_customer"))
                            {
                                erCustomer = (EntityReference)CustAsst["hil_customer"];
                            }
                            #region Standard Warranty Template
                            QueryExpression queryExp = new QueryExpression("hil_warrantytemplate");
                            queryExp.ColumnSet = new ColumnSet("hil_validfrom", "hil_validto", "hil_warrantyperiod", "hil_category");
                            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                            queryExp.Criteria.AddCondition("hil_product", ConditionOperator.Equal, erProdSubCategory.Id);
                            queryExp.Criteria.AddCondition("hil_validfrom", ConditionOperator.NotNull);
                            queryExp.Criteria.AddCondition("hil_validto", ConditionOperator.NotNull);
                            queryExp.Criteria.AddCondition("hil_type", ConditionOperator.Equal, 1); //Standard Warranty
                            queryExp.Criteria.AddCondition("hil_templatestatus", ConditionOperator.Equal, 2); //Approved
                            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                            EntityCollection entCol = _service.RetrieveMultiple(queryExp);
                            if (entCol.Entities.Count > 0)
                            {
                                foreach (Entity ent in entCol.Entities)
                                {
                                    if (invDate >= ent.GetAttributeValue<DateTime>("hil_validfrom") && invDate <= ent.GetAttributeValue<DateTime>("hil_validto"))
                                    {
                                        endDate = CreateUnitWarrantyLine(_service, ent.GetAttributeValue<int>("hil_warrantyperiod"), 1, erCustomer, CustAsst.ToEntityReference(), erProdCgry, erProdSubCategory, erProdModelCode, pdtModelName, invDate, ent.ToEntityReference());
                                        warrantySubstatus = 1;
                                        break;
                                    }
                                }

                            }
                            #endregion

                            #region Extended Warranty Template
                            if (endDate != null)
                            {
                                queryExp = new QueryExpression("hil_warrantytemplate");
                                queryExp.ColumnSet = new ColumnSet("hil_validfrom", "hil_validto", "hil_warrantyperiod", "hil_category");
                                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                queryExp.Criteria.AddCondition("hil_product", ConditionOperator.Equal, erProdSubCategory.Id);
                                queryExp.Criteria.AddCondition("hil_validfrom", ConditionOperator.NotNull);
                                queryExp.Criteria.AddCondition("hil_validto", ConditionOperator.NotNull);
                                queryExp.Criteria.AddCondition("hil_type", ConditionOperator.Equal, 2); //Extended Warranty
                                queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                                queryExp.Criteria.AddCondition("hil_templatestatus", ConditionOperator.Equal, 2); //Approved
                                entCol = _service.RetrieveMultiple(queryExp);
                                foreach (Entity ent in entCol.Entities)
                                {
                                    if (invDate >= ent.GetAttributeValue<DateTime>("hil_validfrom") && invDate <= ent.GetAttributeValue<DateTime>("hil_validto"))
                                    {
                                        endDate = CreateUnitWarrantyLine(_service, ent.GetAttributeValue<int>("hil_warrantyperiod"), 2, erCustomer, CustAsst.ToEntityReference(), erProdCgry, erProdSubCategory, erProdModelCode, pdtModelName, endDate.Value.AddDays(1), ent.ToEntityReference());
                                        warrantySubstatus = 2;
                                        break;
                                    }
                                }
                            }
                            #endregion

                            #region Special Scheme Warranty Template
                            if (endDate != null)
                            {
                                queryExp = new QueryExpression("hil_warrantytemplate");
                                queryExp.ColumnSet = new ColumnSet("hil_validfrom", "hil_validto", "hil_warrantyperiod", "hil_category");
                                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                queryExp.Criteria.AddCondition("hil_product", ConditionOperator.Equal, erProdSubCategory.Id);
                                queryExp.Criteria.AddCondition("hil_validfrom", ConditionOperator.NotNull);
                                queryExp.Criteria.AddCondition("hil_validto", ConditionOperator.NotNull);
                                //queryExp.Criteria.AddCondition("hil_validfrom", ConditionOperator.OnOrAfter, new DateTime(invDate.Value.Year, invDate.Value.Month, invDate.Value.Day));
                                //queryExp.Criteria.AddCondition("hil_validto", ConditionOperator.OnOrBefore, new DateTime(invDate.Value.Year, invDate.Value.Month, invDate.Value.Day));
                                queryExp.Criteria.AddCondition("hil_type", ConditionOperator.Equal, 7); //Scheme Warranty
                                queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                                queryExp.Criteria.AddCondition("hil_templatestatus", ConditionOperator.Equal, 2); //Approved
                                entCol = _service.RetrieveMultiple(queryExp);
                                foreach (Entity ent in entCol.Entities)
                                {
                                    if (invDate >= ent.GetAttributeValue<DateTime>("hil_validfrom") && invDate <= ent.GetAttributeValue<DateTime>("hil_validto"))
                                    {
                                        endDate = CreateSchemeUnitWarrantyLine(_service, ent.GetAttributeValue<int>("hil_warrantyperiod"), 7, erCustomer, CustAsst.ToEntityReference(), erProdCgry, erProdSubCategory, erProdModelCode, pdtModelName, endDate.Value.AddDays(1), ent.ToEntityReference());
                                        warrantySubstatus = 3;
                                        break;
                                    }
                                }
                            }
                            queryExp = new QueryExpression(hil_unitwarranty.EntityLogicalName);
                            queryExp.ColumnSet = new ColumnSet("hil_warrantyenddate", "hil_warrantytemplate");
                            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                            queryExp.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, CustAsst.Id);
                            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            queryExp.AddOrder("hil_warrantyenddate", OrderType.Descending);
                            queryExp.TopCount = 1;
                            entCol = _service.RetrieveMultiple(queryExp);

                            Entity entCustAsset = new Entity(msdyn_customerasset.EntityLogicalName);
                            entCustAsset.Id = CustAsst.Id;

                            if (entCol.Entities.Count > 0)
                            {
                                endDate = entCol.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(530);
                                entCustAsset["hil_warrantytilldate"] = endDate;
                                if (endDate >= DateTime.Now)
                                {
                                    entCustAsset["hil_warrantystatus"] = new OptionSetValue(1); //InWarranty

                                    queryExp = new QueryExpression(hil_unitwarranty.EntityLogicalName);
                                    queryExp.ColumnSet = new ColumnSet("hil_warrantystartdate", "hil_warrantyenddate", "hil_warrantytemplate");
                                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                    queryExp.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, CustAsst.Id);
                                    queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                    queryExp.AddOrder("hil_warrantyenddate", OrderType.Ascending);
                                    entCol = _service.RetrieveMultiple(queryExp);
                                    if (entCol.Entities.Count > 0)
                                    {
                                        foreach (Entity ent in entCol.Entities)
                                        {
                                            if (DateTime.Now >= ent.GetAttributeValue<DateTime>("hil_warrantystartdate") && DateTime.Now <= ent.GetAttributeValue<DateTime>("hil_warrantyenddate"))
                                            {
                                                Entity entSubStatus = _service.Retrieve("hil_warrantytemplate", ent.GetAttributeValue<EntityReference>("hil_warrantytemplate").Id, new ColumnSet("hil_type"));
                                                if (entSubStatus.GetAttributeValue<OptionSetValue>("hil_type").Value == 1)
                                                {
                                                    entCustAsset["hil_warrantysubstatus"] = new OptionSetValue(1); //InWarranty-Standard
                                                }
                                                else if (entSubStatus.GetAttributeValue<OptionSetValue>("hil_type").Value == 2)
                                                {
                                                    entCustAsset["hil_warrantysubstatus"] = new OptionSetValue(2); //InWarranty-Extended
                                                }
                                                else if (entSubStatus.GetAttributeValue<OptionSetValue>("hil_type").Value == 7)
                                                {
                                                    entCustAsset["hil_warrantysubstatus"] = new OptionSetValue(3); //InWarranty-Special Scheme
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    entCustAsset["hil_warrantystatus"] = new OptionSetValue(2); //OutWarranty
                                    entCustAsset["hil_warrantysubstatus"] = null;
                                }
                            }
                            else
                            {
                                entCustAsset["hil_warrantytilldate"] = null;
                                entCustAsset["hil_warrantystatus"] = new OptionSetValue(2); //OutWarranty
                                entCustAsset["hil_warrantysubstatus"] = null;
                            }
                            _service.Update(entCustAsset);
                            #endregion
                        }
                    }
                    Console.WriteLine(rec++.ToString() + "/" + Total.ToString());
                }
            }
        }

        static DateTime? CreateUnitWarrantyLine(IOrganizationService service, int period, int producttype, EntityReference erCustomer, EntityReference erCustomerasset, EntityReference erProductCatg, EntityReference erProductSubCatg, EntityReference erProductModel, string partdescription, DateTime? warrantystartdate, EntityReference erWarrantytemplate)
        {
            DateTime? WarrantyEnd = null;
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
                    DateTime StartDate = new DateTime();
                    int i = 0;
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
                    iSchWarranty.hil_Part = erProductModel;
                    iSchWarranty["hil_partdescription"] = partdescription;
                    iSchWarranty.hil_customer = erCustomer;
                    service.Create(iSchWarranty);
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

        static DateTime? CreateSchemeUnitWarrantyLine(IOrganizationService service, int period, int producttype, EntityReference erCustomer, EntityReference erCustomerasset, EntityReference erProductCatg, EntityReference erProductSubCatg, EntityReference erProductModel, string partdescription, DateTime? warrantystartdate, EntityReference erWarrantytemplate)
        {
            DateTime? WarrantyEnd = null;
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
                        qryExp.ColumnSet = new ColumnSet("hil_salesoffice");
                        qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qryExp.Criteria.AddCondition("hil_warrantytemplate", ConditionOperator.Equal, erWarrantytemplate.Id);
                        qryExp.Criteria.AddCondition("hil_salesoffice", ConditionOperator.Equal, erSalesOffice.Id);
                        qryExp.AddOrder("modifiedon", OrderType.Descending);
                        qryExp.TopCount = 1;
                        entCol = service.RetrieveMultiple(qryExp);
                        if (entCol.Entities.Count > 0)
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
                            iSchWarranty.hil_Part = erProductModel;
                            iSchWarranty["hil_partdescription"] = partdescription;
                            iSchWarranty.hil_customer = erCustomer;
                            service.Create(iSchWarranty);
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

        #region App Setting Load/CRM Connection
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
        #endregion
    }
}
