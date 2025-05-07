using System;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Office.Interop.Excel;
using excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;

namespace D365WebJobs.WarrantyDataCorrection
{
    public class CustomerAssetWarrantyDataCorrection
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";
        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        #endregion

        [STAThread]
        //static void Main(string[] args)
        //{
        //    _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
        //    if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
        //    {
        //        var _prodCatgId = ConfigurationManager.AppSettings["ProductCatg"].ToString();
        //        var _prodSubcatgId = ConfigurationManager.AppSettings["ProductSubCatg"].ToString();

        //        string[] _prodSubCatgArr = new string[] { _prodSubcatgId };
        //        int _rowCount = 1;
        //        string _serialNumber = string.Empty;
        //        EntityCollection entcoll = null;
        //        foreach (string _prodSubcatg in _prodSubCatgArr)
        //        {
        //            _rowCount = 1;
        //            while (true)
        //            {
        //                //<condition attribute='hil_productsubcategory' operator='eq' value='{_prodSubcatg}' />
        //                //<condition attribute='hil_warrantystatus' operator='eq' value='2' />
        //                //<condition attribute='hil_warrantytilldate' operator='null' />
        //                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='50'>
        //                  <entity name='msdyn_customerasset'>
        //                    <attribute name='createdon' />
        //                    <attribute name='msdyn_product' />
        //                    <attribute name='msdyn_name' />
        //                    <attribute name='hil_productsubcategory' />
        //                    <attribute name='hil_productcategory' />
        //                    <attribute name='msdyn_customerassetid' />
        //                    <attribute name='statuscode' />
        //                    <attribute name='hil_branchheadapprovalstatus' />
        //                    <attribute name='hil_invoicedate' />
        //                    <attribute name='hil_invoiceavailable' />
        //                    <order attribute='msdyn_name' descending='false' />
        //                    <filter type='and'>
        //                      <condition attribute='hil_modeofpayment' operator='ne' value='R' />
        //                      <condition attribute='hil_productcategory' operator='eq' value='{_prodCatgId}' />
        //                      <condition attribute='hil_productsubcategory' operator='eq' value='{_prodSubcatg}' />
        //                    </filter>
        //                  </entity>
        //                </fetch>";
        //                entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                if (entcoll.Entities.Count == 0) { Console.WriteLine("Batch Completed... "); break; }
        //                foreach (Entity entCA in entcoll.Entities)
        //                {
        //                    try
        //                    {
        //                        Entity _ent = new Entity(entCA.LogicalName, entCA.Id);
        //                        _serialNumber = entCA.GetAttributeValue<string>("msdyn_name");

        //                        Console.WriteLine("Record Updated.. " + _rowCount++.ToString() + "|" + _serialNumber);

        //                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
        //                            <entity name='hil_unitwarranty'>
        //                            <attribute name='hil_warrantyenddate' />
        //                            <order attribute='hil_warrantyenddate' descending='true' />
        //                                <filter type='and'>
        //                                    <condition attribute='statecode' operator='eq' value='0' />
        //                                    <condition attribute='hil_customerasset' operator='eq' value='{entCA.Id}' />
        //                                </filter>
        //                                <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='wt'>
        //                                    <attribute name='hil_type' />
        //                                </link-entity>
        //                                </entity>
        //                            </fetch>";

        //                        EntityCollection entcoll1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                        if (entcoll1.Entities.Count == 0)
        //                        {
        //                            _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
        //                              <entity name='hil_unitwarranty'>
        //                                <attribute name='hil_warrantyenddate' />
        //                                <order attribute='hil_warrantyenddate' descending='true' />
        //                                <filter type='and'>
        //                                  <condition attribute='hil_customerasset' operator='eq' value='{entCA.Id}' />
        //                                  <condition attribute='statecode' operator='eq' value='1' />
        //                                </filter>
        //                                <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='wt'>
        //                                  <attribute name='hil_type' />
        //                                  <filter type='and'>
        //                                    <condition attribute='hil_type' operator='eq' value='1' />
        //                                  </filter>
        //                                </link-entity>
        //                              </entity>
        //                            </fetch>";

        //                            entcoll1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                            if (entcoll1.Entities.Count == 0) { continue; }
        //                            _service.Update(new Entity()
        //                            {
        //                                LogicalName = entcoll1.Entities[0].LogicalName,
        //                                Id = entcoll1.Entities[0].Id,
        //                                Attributes = new AttributeCollection() { new System.Collections.Generic.KeyValuePair<string, object>("statecode", new OptionSetValue(0)), new System.Collections.Generic.KeyValuePair<string, object>("statuscode", new OptionSetValue(1)) }
        //                            });
        //                        }

        //                        DateTime _warrantyEndDate = entcoll1.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330).Date;
        //                        OptionSetValue _warrantyType = (OptionSetValue)entcoll1.Entities[0].GetAttributeValue<AliasedValue>("wt.hil_type").Value;
        //                        DateTime _dateNow = DateTime.Now.Date;

        //                        if (_warrantyEndDate >= _dateNow)
        //                        {
        //                            _ent["hil_warrantystatus"] = new OptionSetValue(1); //In Warranty
        //                            if (_warrantyType.Value == 3)
        //                                _ent["hil_warrantysubstatus"] = new OptionSetValue(4); //Under AMC
        //                            else
        //                                _ent["hil_warrantysubstatus"] = new OptionSetValue(1); //Standard
        //                        }
        //                        else
        //                        {
        //                            _ent["hil_warrantystatus"] = new OptionSetValue(2); //Out Warranty
        //                        }

        //                        _ent["hil_extendedwarrantyenddate"] = null;
        //                        _ent["hil_warrantytilldate"] = _warrantyEndDate;
        //                        _ent["hil_modeofpayment"] = "R";
        //                        _service.Update(_ent);

        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        Console.WriteLine(ex);
        //                    }

        //                }
        //            }
        //        }
        //    }
        //}

        static void RefreshAssetWarranty(string _productCatg, string _productSubCatg, string _customerAsset)
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

                    if (!string.IsNullOrWhiteSpace(_customerAsset))
                        _condition = $@"<condition attribute='msdyn_customerassetid' operator='eq' value='{_customerAsset}' />";

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
        static void DeleteDuplicateWarrantyLines(string _productCatg, string _productSubCatg)
        {
            int Total = 0;
            while (true)
            {
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='msdyn_customerasset'>
                    <attribute name='createdon' />
                    <attribute name='msdyn_product' />
                    <attribute name='msdyn_name' />
                    <attribute name='hil_productsubcategorymapping' />
                    <attribute name='hil_productcategory' />
                    <attribute name='msdyn_customerassetid' />
                    <order attribute='createdon' descending='true' />
                    <filter type='and'>
                      <condition attribute='hil_productcategory' operator='eq' value='{_productCatg}' />
                      <condition attribute='hil_productsubcategory' operator='eq' value='{_productSubCatg}' />
                      <condition attribute='hil_modeofpayment' operator='ne' value='D' />
                    </filter>
                  </entity>
                </fetch>";

                EntityCollection entColAMC = _service.RetrieveMultiple(new FetchExpression(_fetchXML));

                int rec = 1;
                Total += entColAMC.Entities.Count;
                if (entColAMC.Entities.Count == 0) { break; }
                foreach (Entity entAMC in entColAMC.Entities)
                {
                    Console.WriteLine("Asset# " + entAMC.GetAttributeValue<string>("msdyn_name") + " Record# " + rec + "/" + Total);

                    string _xml = $@"<fetch distinct='false' mapping='logical' aggregate='true'>
                          <entity name='hil_unitwarranty'>
                            <attribute name='hil_unitwarrantyid' alias='uwl' aggregate='count'/>
                            <filter type='and'>
                              <condition attribute='statecode' operator='eq' value='0' />
                              <condition attribute='hil_customerasset' operator='eq' value='{entAMC.Id}' />
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
                              <condition attribute='hil_customerasset' operator='eq' value='{entAMC.Id}' />
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

                            Console.WriteLine("Inactivated .. " + os.Value);
                        }
                    }
                    rec++;

                    Entity _entUpdateStatus = new Entity(entAMC.LogicalName, entAMC.Id);
                    _entUpdateStatus["hil_modeofpayment"] = "WD";
                    _service.Update(_entUpdateStatus);
                }
            }
        }
        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {
                string _prodCatgId = ConfigurationManager.AppSettings["ProductCatg"].ToString();
                string _prodSubcatgId = ConfigurationManager.AppSettings["ProductSubCatg"].ToString();
                //DeleteDuplicateWarrantyLines(_prodCatgId, _prodSubcatgId);
                string _customerAsset = string.Empty;
                //string _customerAsset = "31761d78-5e8b-ee11-8178-000d3a3e4feb";
                RefreshAssetWarranty(_prodCatgId, _prodSubcatgId, _customerAsset);
            }
        }

        //static void Main(string[] args)
        //{
        //    _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
        //    if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
        //    {
        //        string filePath = @"C:\Kuldeep khare\KKGFailed.xlsx";
        //        string conn = string.Empty;
        //        Application excelApp = new Application();
        //        if (excelApp != null)
        //        {
        //            Workbook excelWorkbook = excelApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
        //            Worksheet excelWorksheet = (Worksheet)excelWorkbook.Sheets[1];
        //            Range excelRange = excelWorksheet.UsedRange;
        //            Range range;

        //            string JobId = string.Empty;
        //            DateTime WrtDate;
        //            string WrtDateStr = null;
        //            for (int i = 2; i <= excelRange.Rows.Count; i++)
        //            {
        //                try
        //                {
        //                    #region reading Values from Excel file and declaration of local variables 
        //                    range = (excelWorksheet.Cells[i, 1] as Range);
        //                    JobId = range.Value.ToString();

        //                    #endregion
        //                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                    <entity name='msdyn_workorder'>
        //                    <attribute name='msdyn_name' />
        //                    <attribute name='msdyn_workorderid' />
        //                    <attribute name='msdyn_customerasset' />
        //                    <order attribute='msdyn_name' descending='false' />
        //                    <filter type='and'>
        //                        <condition attribute='msdyn_customerasset' operator='null' />
        //                        <condition attribute='msdyn_name' operator='eq' value='{JobId}' />
        //                    </filter>
        //                    </entity>
        //                    </fetch>";

        //                    EntityCollection invColl = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                    int j = 0;
        //                    foreach (Entity ent in invColl.Entities)
        //                    {
        //                        string _fetchXML1 = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                          <entity name='msdyn_workorderincident'>
        //                            <attribute name='msdyn_customerasset' />
        //                            <attribute name='msdyn_workorderincidentid' />
        //                            <order attribute='msdyn_customerasset' descending='false' />
        //                            <filter type='and'>
        //                              <condition attribute='msdyn_workorder' operator='eq' value='{invColl.Entities[0].Id}' />
        //                              <condition attribute='statecode' operator='eq' value='0' />
        //                            </filter>
        //                          </entity>
        //                        </fetch>";
        //                        EntityCollection invColl1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML1));
        //                        if (invColl1.Entities.Count > 0)
        //                        {
        //                            Entity _entJob = new Entity(ent.LogicalName, ent.Id);
        //                            _entJob["msdyn_customerasset"] = invColl1.Entities[0].GetAttributeValue<EntityReference>("msdyn_customerasset");
        //                            _service.Update(_entJob);
        //                        }
        //                        Console.WriteLine("Processing.. " + i.ToString() + "/" + j.ToString());
        //                    }
        //                }
        //                catch (Exception ex)
        //                {
        //                    Console.WriteLine(ex);
        //                }
        //                Console.WriteLine("Record Updated.. " + i.ToString() + "|" + JobId);
        //            }
        //        }
        //    }
        //}
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
