using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsCustomerAssetWarrantyRefresh
{
    class Program
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
                string _customerAsset = ConfigurationManager.AppSettings["CustomerAsset"].ToString();
                string _prodCatgArr = ConfigurationManager.AppSettings["ProductCatgArr"].ToString();
                //RefreshAssetWarrantyBulk(_prodCatgId, _prodSubcatgId);
                //CustomerAssetWarrantyDailyRefresh(_prodCatgArr);
                //RefreshAssetWarranty(_prodCatgId, _prodSubcatgId, _customerAsset);
                AMCSAPInvoiceProcess();
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

        static void RefreshAssetWarranty(string _productCatg, string _productSubCatg, string _customerAsset)
        {
            Guid CustomerAssetId = Guid.Empty;
            try
            {
                string _serialNumber = string.Empty;
                EntityCollection entcoll = null;
                string _conditionFilter = string.Empty;

                string _condition = string.Empty;
                //if (!string.IsNullOrWhiteSpace(_productCatg))
                //{
                //    _conditionFilter = $"<condition attribute='hil_productcategory' operator='eq' value='{_productCatg}' /><condition attribute='hil_modeofpayment' operator='ne' value='WR' />";
                //}
                //if (!string.IsNullOrWhiteSpace(_productSubCatg))
                //{
                //    _conditionFilter += $"<condition attribute='hil_productsubcategory' operator='eq' value='{_productSubCatg}' />";
                //}
                if (!string.IsNullOrWhiteSpace(_customerAsset))
                {
                    _conditionFilter += $"<condition attribute='msdyn_customerassetid' operator='eq' value='{_customerAsset}' />";
                }
                //
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
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
                              {_conditionFilter}
                            </filter>
                          </entity>
                        </fetch>";
                entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                foreach (Entity entCA in entcoll.Entities)
                {
                    try
                    {
                        DeleteDuplicateLines(entCA.Id);
                        Entity _ent = new Entity(entCA.LogicalName, entCA.Id);

                        _serialNumber = entCA.GetAttributeValue<string>("msdyn_name");
                        Console.WriteLine($"Asset# {_serialNumber}");

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
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(CustomerAssetId.ToString() + " : " + ex.Message);
            }
        }

        static void RefreshAssetWarrantyBulk(string _productCatg, string _productSubCatg)
        {
            Guid CustomerAssetId = Guid.Empty;
            try
            {
                int _rowCount = 1, _totalRowCount = 0;
                string _serialNumber = string.Empty;
                int _pageSize = 1000;
                EntityCollection entcoll = null;
                string _conditionFilter = string.Empty;

                while (true)
                {
                    string _condition = string.Empty;
                    if (!string.IsNullOrWhiteSpace(_productCatg))
                    {
                        _conditionFilter = $"<condition attribute='hil_productcategory' operator='eq' value='{_productCatg}' />";
                    }
                    if (!string.IsNullOrWhiteSpace(_productSubCatg))
                    {
                        _conditionFilter += $"<condition attribute='hil_productsubcategory' operator='eq' value='{_productSubCatg}' />";
                    }
                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
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
                              {_conditionFilter}
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
                            DeleteDuplicateLines(entCA.Id);
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
        static void DeleteDuplicateLines(Guid _customerAssetId) {

            string _xml = $@"<fetch distinct='false' mapping='logical' aggregate='true'>
                <entity name='hil_unitwarranty'>
                <attribute name='hil_unitwarrantyid' alias='uwl' aggregate='count'/>
                <filter type='and'>
                    <condition attribute='statecode' operator='eq' value='0' />
                    <condition attribute='hil_customerasset' operator='eq' value='{_customerAssetId}' />
                </filter>
                <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='ac'>
                    <attribute name='hil_type' groupby='true' alias='wtype'/>
                    <filter type='and'>
                        <condition attribute='hil_type' operator='ne' value='3' />
                    </filter>
                </link-entity>
                </entity>
            </fetch>";

            List<OptionSetValue> _warrantyType = new List<OptionSetValue>();
            EntityCollection entColAMC3 = _service.RetrieveMultiple(new FetchExpression(_xml));
            foreach (Entity _entType in entColAMC3.Entities)
            {
                int _count = (int)_entType.GetAttributeValue<AliasedValue>("uwl").Value;
                OptionSetValue _wtType = (OptionSetValue)_entType.GetAttributeValue<AliasedValue>("wtype").Value;
                if (_count > 1)
                    _warrantyType.Add(_wtType);
            }
            foreach (OptionSetValue os in _warrantyType)
            {
                string _fetchXML1 = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='hil_unitwarranty'>
                    <attribute name='hil_unitwarrantyid' />
                    <order attribute='createdon' descending='true' />
                    <filter type='and'>
                        <condition attribute='statecode' operator='eq' value='0' />
                        <condition attribute='hil_customerasset' operator='eq' value='{_customerAssetId}' />
                    </filter>
                    <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='ai'>
                        <filter type='and'>
                        <condition attribute='hil_type' operator='eq' value='{os.Value}' />
                        </filter>
                    </link-entity>
                    </entity>
                </fetch>";
                EntityCollection entColAMC1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML1));
                int _rowCount = 1;
                foreach (Entity ent in entColAMC1.Entities)
                {
                    if (_rowCount++ == 1)
                    {
                        continue;
                    }
                    Entity _entInactive = new Entity(ent.LogicalName, ent.Id);
                    _entInactive["statecode"] = new OptionSetValue(1);// Inactive
                    _entInactive["statuscode"] = new OptionSetValue(2);// Inactive
                    _service.Update(_entInactive);
                }
            }
        }
        static void CustomerAssetWarrantyDailyRefresh(string _productCatg)
        {
            Guid CustomerAssetId = Guid.Empty;
            try
            {
                string _batchStartDatetime = DateTime.Now.AddMinutes(330).ToString("yyyy-MM-dd HH:mm:ss");
                int _rowCount = 1, _totalRowCount = 0;
                string _serialNumber = string.Empty;
                EntityCollection entcoll = null;

                string[] _productCatgArray = _productCatg.Split(',');
                string _condition = string.Empty;
                string _warrantyTillDate = DateTime.Now.AddMinutes(330).AddDays(-1).ToString("yyyy-MM-dd");
                //<condition attribute='modifiedon' operator='lt' value='{_batchStartDatetime}' />
                //<condition attribute='hil_productcategory' operator='eq' value='{_prodCatg}' />
                Console.WriteLine("Batch Starts " + _batchStartDatetime);
                Console.WriteLine("Warranty Run Date " + _warrantyTillDate);
                foreach (string _prodCatg in _productCatgArray)
                {
                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='msdyn_customerasset'>
                            <attribute name='msdyn_name' />
                            <attribute name='msdyn_customerassetid' />
                            <order attribute='msdyn_name' descending='false' />
                            <filter type='and'>
                                <condition attribute='msdyn_name' operator='not-null' />
                                <condition attribute='hil_productcategory' operator='eq' value='{_prodCatg}' />
                            </filter>
                            <link-entity name='hil_unitwarranty' from='hil_customerasset' to='msdyn_customerassetid' link-type='inner' alias='ab'>
                                <filter type='and'>
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_warrantyenddate' operator='on' value='{_warrantyTillDate}' />
                                </filter>
                            </link-entity>
                            </entity>
                            </fetch>";
                    entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    _totalRowCount += entcoll.Entities.Count;
                    foreach (Entity entCA in entcoll.Entities)
                    {
                        try
                        {
                            Entity _ent = new Entity(entCA.LogicalName, entCA.Id);

                            _serialNumber = entCA.GetAttributeValue<string>("msdyn_name");
                            Console.WriteLine($"Asset# {_serialNumber} Row Count: {_rowCount++}/{_totalRowCount}");

                            string _processDate = DateTime.Now.AddMinutes(330).ToString("yyyy-MM-dd");
                            Console.WriteLine("Process Date " + _processDate);
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

        static void AMCSAPInvoiceProcess()
        {
            int Total = 0;
            string _error = string.Empty;

            while (true)
            {
                //QueryExpression queryExpTemp = new QueryExpression("hil_amcstaging");
                //queryExpTemp.ColumnSet = new ColumnSet(true);
                //queryExpTemp.Criteria = new FilterExpression(LogicalOperator.And);
                //queryExpTemp.Criteria.AddCondition("hil_amcstagingstatus", ConditionOperator.Equal, false); //Draft
                //queryExpTemp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                //queryExpTemp.Criteria.AddCondition("hil_serailnumber", ConditionOperator.Equal, "67JHG61R00986");
                //EntityCollection entColAMC = _service.RetrieveMultiple(queryExpTemp);

                QueryExpression queryExpTemp = new QueryExpression("hil_amcstaging");
                queryExpTemp.ColumnSet = new ColumnSet(true);
                queryExpTemp.Criteria = new FilterExpression(LogicalOperator.And);
                //queryExpTemp.Criteria.AddCondition("createdby", ConditionOperator.Equal, new Guid("43A82D38-FCEE-E811-A949-000D3AF03089")); //Draft
                //queryExpTemp.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, "2024-03-01");
                //queryExpTemp.Criteria.AddCondition("modifiedon", ConditionOperator.OnOrBefore, "2024-03-21");
                //queryExpTemp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                queryExpTemp.Criteria.AddCondition("hil_serailnumber", ConditionOperator.Equal, "67HFG60102825");
                queryExpTemp.Criteria.AddCondition("hil_amcstagingstatus", ConditionOperator.Equal, false); //Draft
                EntityCollection entColAMC = _service.RetrieveMultiple(queryExpTemp);

                int rec = 1;
                Total += entColAMC.Entities.Count;
                if (entColAMC.Entities.Count == 0) { break; }
                foreach (Entity entAMC in entColAMC.Entities)
                {
                    Entity _updateAMCStg = new Entity(entAMC.LogicalName, entAMC.Id);

                    Console.WriteLine(rec.ToString() + "/" + Total.ToString() + " Serial # " + entAMC.GetAttributeValue<string>("hil_serailnumber") + "/" + entAMC.GetAttributeValue<string>("hil_name") + "/" + entAMC.GetAttributeValue<DateTime>("hil_warrantystartdate").AddMinutes(330).Date + "/" + entAMC.GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330).Date);

                    #region Checking if SAP Invoice Number is exist in Unit Warranty Line
                    string _fetchXMLUWL = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_unitwarranty'>
                        <attribute name='hil_name' />
                        <filter type='and'>
                            <condition attribute='hil_amcbillingdocnum' operator='eq' value='{entAMC.GetAttributeValue<string>("hil_name")}' />
                            <condition attribute='statecode' operator='eq' value='0' />
                        </filter>
                        </entity>
                    </fetch>";
                    EntityCollection _entUWL = _service.RetrieveMultiple(new FetchExpression(_fetchXMLUWL));
                    if (_entUWL.Entities.Count > 0) {
                        _error = "SAP Invoice Number is already exist in D365 Unit Warranty Line " + _entUWL.Entities[0].GetAttributeValue<string>("hil_name");
                        Console.WriteLine(_error);
                        _updateAMCStg["hil_amcstagingstatus"] = true;
                        _updateAMCStg["hil_description"] = _error;
                        _service.Update(_updateAMCStg);
                        rec++;
                        continue;
                    }
                    #endregion

                    
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
                    <attribute name='createdon' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='msdyn_name' operator='eq' value='" + entAMC.GetAttributeValue<string>("hil_serailnumber") + @"' />
                    </filter>
                    </entity>
                    </fetch>";

                    EntityCollection ecCA = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (ecCA.Entities.Count > 0)
                    {
                        try
                        {
                            Entity _customerAsset = ecCA.Entities[0];

                            #region Checking if Unit Warranty Line is already created
                            string _startDate = entAMC.GetAttributeValue<DateTime>("hil_warrantystartdate").AddMinutes(330).Date.ToString("yyyy-MM-dd");
                            string _endDate = entAMC.GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330).Date.ToString("yyyy-MM-dd");

                            _fetchXMLUWL = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_unitwarranty'>
                                <attribute name='hil_name' />
                                <order attribute='hil_name' descending='false' />
                                <filter type='and'>
                                    <condition attribute='hil_amcbillingdocnum' operator='eq' value='{entAMC.GetAttributeValue<string>("hil_name")}' />
                                    <condition attribute='statecode' operator='eq' value='0' />
                                    <condition attribute='hil_customerasset' operator='eq' value='{_customerAsset.Id}' />
                                    <condition attribute='hil_warrantystartdate' operator='on' value='{_startDate}' />
                                    <condition attribute='hil_warrantyenddate' operator='on' value='{_endDate}' />
                                </filter>
                                </entity>
                                </fetch>";
                            EntityCollection _entValidate = _service.RetrieveMultiple(new FetchExpression(_fetchXMLUWL));
                            if (_entValidate.Entities.Count > 0)
                            {
                                _error = "Unit Warranty Line is already created " + _entUWL.Entities[0].GetAttributeValue<string>("hil_name");
                                Console.WriteLine(_error);
                                _updateAMCStg["hil_amcstagingstatus"] = true;
                                _updateAMCStg["hil_description"] = _error;
                                _service.Update(_updateAMCStg);
                                rec++;
                                continue;
                            }
                            #endregion

                            EntityReference erProdCgry = new EntityReference("product");
                            EntityReference erProdSubCategory = new EntityReference("product");
                            EntityReference erProdModelCode = new EntityReference("product");
                            EntityReference erCustomer = new EntityReference("contact");
                            Entity _warrantyTemplate = null;
                            string pdtModelName = string.Empty;

                            QueryExpression queryExp = new QueryExpression("hil_warrantytemplate");
                            queryExp.ColumnSet = new ColumnSet("hil_warrantytemplateid");
                            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                            queryExp.Criteria.AddCondition(new ConditionExpression("hil_amcplan", ConditionOperator.Equal, entAMC.GetAttributeValue<EntityReference>("hil_amcplan").Id));
                            queryExp.Criteria.AddCondition(new ConditionExpression("hil_templatestatus", ConditionOperator.Equal, 2)); //Approved
                            queryExp.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1)); //Active
                            EntityCollection entCollTemp = _service.RetrieveMultiple(queryExp);

                            if (entCollTemp.Entities.Count > 0)
                            {
                                _warrantyTemplate = entCollTemp.Entities[0];
                            }
                            if (_customerAsset != null && _warrantyTemplate != null)
                            {
                                Entity hil_unitwarranty = new Entity("hil_unitwarranty");
                                hil_unitwarranty["hil_customerasset"] = _customerAsset.ToEntityReference();
                                if (_customerAsset.Attributes.Contains("hil_productcategory"))
                                {
                                    hil_unitwarranty["hil_productmodel"] = _customerAsset.GetAttributeValue<EntityReference>("hil_productcategory");
                                }

                                if (_customerAsset.Attributes.Contains("hil_productsubcategory"))
                                {
                                    hil_unitwarranty["hil_productitem"] = _customerAsset.GetAttributeValue<EntityReference>("hil_productsubcategory");
                                }

                                hil_unitwarranty["hil_warrantystartdate"] = entAMC.GetAttributeValue<DateTime>("hil_warrantystartdate").AddMinutes(330).Date;
                                hil_unitwarranty["hil_warrantyenddate"] = entAMC.GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330).Date;

                                hil_unitwarranty["hil_warrantytemplate"] = _warrantyTemplate.ToEntityReference();

                                hil_unitwarranty["hil_producttype"] = new OptionSetValue(1);

                                if (_customerAsset.Attributes.Contains("msdyn_product"))
                                {
                                    hil_unitwarranty["hil_part"] = _customerAsset.GetAttributeValue<EntityReference>("msdyn_product");
                                }

                                if (_customerAsset.Attributes.Contains("hil_modelname"))
                                {
                                    hil_unitwarranty["hil_partdescription"] = _customerAsset.GetAttributeValue<string>("hil_modelname");
                                }

                                if (_customerAsset.Attributes.Contains("hil_customer"))
                                {
                                    hil_unitwarranty["hil_customer"] = _customerAsset.GetAttributeValue<EntityReference>("hil_customer");
                                }
                                hil_unitwarranty["hil_amcbillingdocdate"] = entAMC.GetAttributeValue<DateTime>("hil_sapbillingdate").AddMinutes(330);
                                hil_unitwarranty["hil_amcbillingdocnum"] = entAMC.GetAttributeValue<string>("hil_name");
                                hil_unitwarranty["hil_amcbillingdocurl"] = entAMC.GetAttributeValue<string>("hil_sapbillingdocpath");
                                _service.Create(hil_unitwarranty);
                                RefreshAssetWarranty(string.Empty, string.Empty, _customerAsset.Id.ToString());
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("ERROR!!! "+ entAMC.GetAttributeValue<string>("hil_serailnumber") +" :: " +ex.Message + " / " + rec.ToString() + "/" + Total.ToString());
                        }
                        _updateAMCStg["hil_amcstagingstatus"] = true;
                        _updateAMCStg["hil_description"] = "SUCCESS";
                        _service.Update(_updateAMCStg);
                    }
                    else {
                        _error = "Asset Serial Number Doesn't exist in D365.";
                        Console.WriteLine(_error);
                        _updateAMCStg["hil_amcstagingstatus"] = true;
                        _updateAMCStg["hil_description"] = _error;
                        _service.Update(_updateAMCStg);
                    }
                    rec++;
                }
            }
        }
    }
}
