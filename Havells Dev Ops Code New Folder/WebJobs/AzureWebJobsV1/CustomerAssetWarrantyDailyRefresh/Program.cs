using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System;
using Microsoft.Xrm.Sdk.Query;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Xrm.Tooling.Connector;
using System.Collections.Generic;

namespace CustomerAssetWarrantyDailyRefresh
{
    public class Program
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
                //DailyRefreshCustomerAssetWarranty();
                //UpdateJobUpcountry(_service);
                //RefreshAssetUnitWarrantyLLOYD();
                //RefreshAssetUnitWarranty();
                //DeleteAssetUnitWarranty();
                //DeactivateUnitWarrantyLine();
                //UpdateCustomerCategory(_service);
                //UpdateClaimParameterOnJob();
                //RepeatRepairSAW();
                //RefreshClaimParameters();
                //CreateSAWActivityApprovals();
                //RefreshSAWActivity();
                //RefreshRRSAWActivity();
                //UpdateCustomerCategory(_service);

                // -------- **** Live Webjobs Code 18/Apr/2022 **** --------
                //RefreshAssetUnitWarranty();
                //RefreshCustomerAssetWarrantyMissed(); //Acivate this line of code for missed Warranty Lines Webjob

                //RefreshAssetUnitWarrantyAMC(); //Acivate this line of code for AMC Warranty Lines Webjob


                RefreshHavellsWaterPurifierUnitWarranty(string.Empty); //LLOYD AMC Upload
                //RefreshClaimLinesOnceAMCInvoiceImported();
                //RefreshJobsLabourInWaranty();

                //DeleteDuplicateWarrantyLines();

                //RefreshLLOYDWarrantyLines_Inactive();
                //CancelBulkJobs(_service);
                //CancelBulkJobs(_service);

                //RefreshLLOYDAMCUWLs();
            }
        }

        static void RefreshLLOYDAMCUWLs()
        {
            int Total = 0;
            string _fetchXMLAMC = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_amcstaging'>
                <attribute name='hil_name' />
                <attribute name='hil_serailnumber' />
                <attribute name='hil_amcplan' />
                <filter type='and'>
                    <filter type='or'>
                    <condition attribute='hil_amcplannameslt' operator='like' value='EW%' />
                    <condition attribute='hil_amcplannameslt' operator='like' value='AMC%' />
                    </filter>
                    <condition attribute='hil_amcstagingstatus' operator='eq' value='1' />
                    <condition attribute='hil_description' operator='like' value='%DONE%' />
                </filter>
                </entity>
                </fetch>";
            int i = 1;
            while (true)
            {
                EntityCollection entAMCBill = _service.RetrieveMultiple(new FetchExpression(_fetchXMLAMC));
                if (entAMCBill.Entities.Count == 0) { break; }
                Entity _entCustAsset = null;
                EntityReference _entAMCPlan = null;
                foreach (Entity entAMC in entAMCBill.Entities)
                {
                    string _serialNum = entAMC.GetAttributeValue<string>("hil_serailnumber");
                    string _invoiceNum = entAMC.GetAttributeValue<string>("hil_name");
                    _entAMCPlan = entAMC.GetAttributeValue<EntityReference>("hil_amcplan");
                    Console.WriteLine("Processing... (" + _serialNum + ")" + i++.ToString());

                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_customerasset'>
                        <attribute name='msdyn_customerassetid' />
                        <filter type='and'>
                            <condition attribute='msdyn_name' operator='eq' value='{_serialNum}' />
                        </filter>
                        </entity>
                        </fetch>";
                    EntityCollection entTemp = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entTemp.Entities.Count > 0) {
                        _entCustAsset = new Entity(entTemp.Entities[0].LogicalName, entTemp.Entities[0].Id);
                    }
                    
                    if (_entAMCPlan != null && _entCustAsset != null) {
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='hil_unitwarranty'>
                            <attribute name='hil_unitwarrantyid' />
                            <attribute name='hil_name' />
                            <filter type='and'>
                              <condition attribute='hil_customerasset' operator='eq' value='{_entCustAsset.Id}' />
                              <condition attribute='hil_amcbillingdocnum' operator='eq' value='{_invoiceNum}' />
                              <condition attribute='statecode' operator='eq' value='1' />
                            </filter>
                            <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='wt'>
                              <filter type='and'>
                                <condition attribute='hil_amcplan' operator='eq' value='{_entAMCPlan.Id}' />
                              </filter>
                            </link-entity>
                          </entity>
                        </fetch>";
                        entTemp = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entTemp.Entities.Count > 0) {
                            Console.WriteLine($"Found UWL:{entTemp.Entities[0].GetAttributeValue<string>("hil_name")}");
                            Entity _entUWL = new Entity(entTemp.Entities[0].LogicalName, entTemp.Entities[0].Id);
                            _entUWL["statecode"] = new OptionSetValue(0);
                            _entUWL["statuscode"] = new OptionSetValue(1);
                            _service.Update(_entUWL);

                            Entity _entCust = new Entity(_entCustAsset.LogicalName, _entCustAsset.Id);
                            _entCust["hil_createwarranty"] = false;
                            _service.Update(_entCust);
                            _entCust = new Entity(_entCustAsset.LogicalName, _entCustAsset.Id);
                            _entCust["hil_createwarranty"] = true;
                            _service.Update(_entCust);
                        }
                    }
                    Entity _entAMCRecord = new Entity(entAMC.LogicalName, entAMC.Id);
                    _entAMCRecord["hil_description"] = "Refreshed";
                    _service.Update(_entAMCRecord);
                }
            }
        }
        static void RefreshLLOYDWarrantyLines_Inactive() {
            string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                  <entity name='msdyn_customerasset'>
                    <attribute name='createdon' />
                    <attribute name='msdyn_product' />
                    <attribute name='msdyn_name' />
                    <attribute name='hil_productsubcategorymapping' />
                    <attribute name='hil_productcategory' />
                    <attribute name='msdyn_customerassetid' />
                    <order attribute='createdon' descending='true' />
                    <filter type='and'>
                      <condition attribute='hil_productcategory' operator='in'>
                        <value uiname='LLOYD AIR CONDITIONER' uitype='product'>{D51EDD9D-16FA-E811-A94C-000D3AF0694E}</value>
                        <value uiname='LLOYD LED TELEVISION' uitype='product'>{A7A5049B-16FA-E811-A94C-000D3AF06091}</value>
                        <value uiname='LLOYD REFRIGERATORS' uitype='product'>{2DD99DA1-16FA-E811-A94C-000D3AF06091}</value>
                        <value uiname='LLOYD WASHING MACHIN' uitype='product'>{2FD99DA1-16FA-E811-A94C-000D3AF06091}</value>
                      </condition>
                    </filter>
                    <link-entity name='hil_unitwarranty' from='hil_customerasset' to='msdyn_customerassetid' link-type='inner' alias='ac'>
                      <filter type='and'>
                        <condition attribute='statecode' operator='eq' value='0' />
                      </filter>
                      <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='ad'>
                        <filter type='and'>
                          <condition attribute='hil_product' operator='eq' value='{4AABBB57-A85E-EA11-A811-000D3AF057DD}' />
                          <condition attribute='statecode' operator='eq' value='1' />
                        </filter>
                      </link-entity>
                    </link-entity>
                  </entity>
                </fetch>";

            EntityCollection entCol = null;
            int i = 1, totalCount = 0;
            while (true) {
                entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entCol.Entities.Count == 0) { break; }
                totalCount += entCol.Entities.Count;
                foreach (Entity ent in entCol.Entities) {
                    EntityReference _warrantyPlan = null;
                    EntityReference _productSubCatg = null;
                    EntityReference _warrantyTemplate = null;
                    Console.WriteLine($"Processing {i++.ToString()} /{totalCount.ToString()}  Serial Number {ent.Attributes["msdyn_name"] }");
                    string _fetchWarrantyInfo = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='hil_unitwarranty'>
                        <attribute name='hil_warrantytemplate' />
                        <attribute name='hil_warrantystartdate' />
                        <attribute name='hil_warrantyenddate' />
                        <attribute name='hil_customerasset' />
                        <attribute name='hil_unitwarrantyid' />
                        <attribute name='hil_productitem' />
                        <order attribute='hil_customerasset' descending='false' />
                        <filter type='and'>
                          <condition attribute='hil_customerasset' operator='eq' value='{ent.Id}' />
                          <condition attribute='statecode' operator='eq' value='0' />
                        </filter>
                        <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='ab'>
                          <attribute name='hil_amcplan' />
                          <filter type='and'>
                            <condition attribute='hil_product' operator='eq' uiname='OTHERS' uitype='product' value='{{4AABBB57-A85E-EA11-A811-000D3AF057DD}}' />
                            <condition attribute='statecode' operator='eq' value='1' />
                          </filter>
                        </link-entity>
                      </entity>
                    </fetch>";
                    EntityCollection entCol1 = _service.RetrieveMultiple(new FetchExpression(_fetchWarrantyInfo));
                    foreach (Entity entWty in entCol1.Entities)
                    {
                        _warrantyPlan = (EntityReference)entWty.GetAttributeValue<AliasedValue>("ab.hil_amcplan").Value;
                        _productSubCatg = entWty.GetAttributeValue<EntityReference>("hil_productitem");
                        _warrantyTemplate = entWty.GetAttributeValue<EntityReference>("hil_warrantytemplate");

                        string _fetchWarrantyInfo1 = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='hil_unitwarranty'>
                            <attribute name='hil_name' />
                            <attribute name='hil_unitwarrantyid' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                              <condition attribute='statecode' operator='eq' value='0' />
                              <condition attribute='hil_customerasset' operator='eq' value='{ent.Id}' />
                            </filter>
                            <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='ae'>
                              <filter type='and'>
                                <condition attribute='hil_amcplan' operator='eq' value='{_warrantyPlan.Id}' />
                                <condition attribute='statecode' operator='eq' value='0' />
                              </filter>
                            </link-entity>
                          </entity>
                        </fetch>";
                        EntityCollection entCol2 = _service.RetrieveMultiple(new FetchExpression(_fetchWarrantyInfo1));
                        if (entCol2.Entities.Count > 0)
                        {
                            Entity entUWL = new Entity(entWty.LogicalName, entWty.Id);
                            entUWL["statecode"] = new OptionSetValue(1);// Inactive
                            entUWL["statuscode"] = new OptionSetValue(2);// Inactive
                            _service.Update(entUWL);
                        }
                        else {
                            string _newWarranty = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_warrantytemplate'>
                            <attribute name='hil_warrantytemplateid' />
                            <filter type='and'>
                                <condition attribute='hil_product' operator='eq' value='{_productSubCatg.Id}' />
                                <condition attribute='hil_amcplan' operator='eq' value='{_warrantyPlan.Id}' />
                                <condition attribute='hil_templatestatus' operator='eq' value='2' />
                                <condition attribute='hil_validfrom' operator='on-or-before' value='2024-03-19' />
                                <condition attribute='hil_validto' operator='on-or-after' value='2024-03-19' />
                                <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            </entity>
                            </fetch>";
                            EntityCollection entCol3 = _service.RetrieveMultiple(new FetchExpression(_newWarranty));
                            if (entCol3.Entities.Count > 0)
                            {
                                Entity entUWL = new Entity(entWty.LogicalName, entWty.Id);
                                entUWL["hil_warrantytemplate"] = entCol3.Entities[0].ToEntityReference();
                                _service.Update(entUWL);
                            }
                            else {
                                Entity entWarrantyTempUpdate = new Entity(_warrantyTemplate.LogicalName, _warrantyTemplate.Id);
                                entWarrantyTempUpdate["hil_product"] = _warrantyPlan;
                                entWarrantyTempUpdate["hil_templatestatus"] = new OptionSetValue(2);//Approved
                                entWarrantyTempUpdate["statecode"] = new OptionSetValue(0);
                                entWarrantyTempUpdate["statuscode"] = new OptionSetValue(1);
                                _service.Update(entWarrantyTempUpdate);

                            }
                        }
                    }

                    Entity entCust = new Entity(ent.LogicalName, ent.Id);
                    entCust["hil_modeofpayment"] = "WP";
                    _service.Update(entCust);
                }
            }
        }

        static void DeleteDuplicateWarrantyLines()
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
                      <condition attribute='hil_productcategory' operator='eq' value='{{72981D83-16FA-E811-A94C-000D3AF0694E}}' />
                      <condition attribute='hil_productsubcategory' operator='eq' value='{{851B7022-410B-E911-A94F-000D3AF00F43}}' />
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
                            if (_rowCount++ == 1) { 
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
                    _entUpdateStatus["hil_modeofpayment"] = "D";
                    _service.Update(_entUpdateStatus);
                }
            }
        }
        private static void RefreshJobsLabourInWaranty()
        {
            string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='hil_amcstaging'>
                    <attribute name='hil_serailnumber' />
                    <filter type='and'>
                      <condition attribute='createdby' operator='eq' value='{0DC1D827-DC64-E911-A96C-000D3AF03089}' />
                      <condition attribute='createdon' operator='on-or-after' value='2024-03-19' />
                      <condition attribute='hil_description' operator='like' value='%DONE%' />
                    </filter>
                  </entity>
                </fetch>";

            int rec = 0;
            int Total = 0;
            while (true)
            {
                EntityCollection ec = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                Total += ec.Entities.Count;
                if (ec.Entities.Count == 0) { break; }
                string _serialNumber = "";
                foreach (Entity CustAsst in ec.Entities)
                {
                    try
                    {
                        _serialNumber = CustAsst.GetAttributeValue<string>("hil_serailnumber");
                        string _fetchJob = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='msdyn_workorder'>
                            <attribute name='msdyn_name' />
                            <attribute name='hil_generateclaim' />
                            <order attribute='msdyn_name' descending='false' />
                            <filter type='and'>
                              <condition attribute='hil_fiscalmonth' operator='eq' value='{{50DE3041-22CE-EE11-904C-000D3A3E3D4E}}' />
                              <condition attribute='hil_laborinwarranty' operator='ne' value='1' />
                              <condition attribute='hil_generateclaim' operator='eq' value='1' />
                            </filter>
                            <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='msdyn_customerasset' link-type='inner' alias='ae'>
                              <filter type='and'>
                                <condition attribute='msdyn_name' operator='eq' value='{_serialNumber}' />
                              </filter>
                            </link-entity>
                          </entity>
                        </fetch>";
                        EntityCollection entColJobs = _service.RetrieveMultiple(new FetchExpression(_fetchJob));
                        foreach (Entity ent in entColJobs.Entities) {
                            Entity entJob = new Entity(ent.LogicalName, ent.Id);
                            entJob["hil_generateclaim"] = false;
                            entJob["hil_fiscalmonth"] = null;
                            entJob["hil_reporttext"] = "KK";
                            _service.Update(entJob);
                            Console.WriteLine("Processing... " + ent.GetAttributeValue<string>("msdyn_name"));
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    Console.WriteLine(_serialNumber + " / " + rec++.ToString() + "/" + Total.ToString());
                }
            }
        }

        static void RefreshCustomerAssetWarrantyMissed()
        {
            string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
            <entity name='msdyn_customerasset'>
            <attribute name='msdyn_name' />
            <attribute name='hil_createwarranty' />
            <attribute name='hil_productsubcategorymapping' />
            <attribute name='msdyn_customerassetid' />
            <attribute name='hil_invoicedate' />
            <order attribute='createdon' descending='false' />
            <filter type='and'>
                <condition attribute='createdon' operator='last-x-hours' value='100' />
                <condition attribute='hil_invoiceavailable' operator='eq' value='1' />
                <condition attribute='hil_invoicedate' operator='not-null' />
                <condition attribute='hil_productsubcategorymapping' operator='not-null' />
                <filter type='or'>
                    <condition attribute='hil_branchheadapprovalstatus' operator='eq' value='1' />
                    <condition attribute='statuscode' operator='eq' value='910590001' />
                </filter>
            </filter>
            <link-entity name='hil_unitwarranty' from='hil_customerasset' to='msdyn_customerassetid' link-type='outer' alias='ac' />
            <filter type='and'>
                <condition entityname='ac' attribute='hil_customerasset' operator='null' />
            </filter>
            </entity>
            </fetch>";
            int Total = 0;
            int rec = 1;
            try
            {
                EntityCollection ec = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                Total = ec.Entities.Count;
                Console.WriteLine("Job Start / " + DateTime.Now.ToString());
                Entity _custEnt;
                foreach (Entity CustAsst in ec.Entities)
                {
                    Console.WriteLine(CustAsst.GetAttributeValue<DateTime>("createdon").ToString() + " / " + rec++.ToString() + "/" + Total.ToString());
                    if (CustAsst.Attributes.Contains("hil_productsubcategorymapping"))
                    {
                        hil_stagingdivisonmaterialgroupmapping Map = (hil_stagingdivisonmaterialgroupmapping)_service.Retrieve(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, CustAsst.GetAttributeValue<EntityReference>("hil_productsubcategorymapping").Id, new ColumnSet("hil_productsubcategorymg"));
                        if (Map.hil_ProductSubCategoryMG != null)
                        {
                            QueryExpression queryExpTemp = new QueryExpression("hil_warrantytemplate");
                            queryExpTemp.ColumnSet = new ColumnSet("hil_warrantyperiod");
                            queryExpTemp.Criteria = new FilterExpression(LogicalOperator.And);
                            queryExpTemp.Criteria.AddCondition("hil_product", ConditionOperator.Equal, Map.hil_ProductSubCategoryMG.Id); //Product Sub Category
                            queryExpTemp.Criteria.AddCondition("hil_validfrom", ConditionOperator.OnOrBefore, CustAsst.GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330)); //Valid From
                            queryExpTemp.Criteria.AddCondition("hil_validto", ConditionOperator.OnOrAfter, CustAsst.GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330)); //Valid To
                            queryExpTemp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active Record
                            EntityCollection entColWrtyTemplate = _service.RetrieveMultiple(queryExpTemp);

                            if (entColWrtyTemplate.Entities.Count > 0)
                            {
                                _custEnt = new Entity("msdyn_customerasset");
                                _custEnt.Id = CustAsst.Id;
                                _custEnt["hil_createwarranty"] = false;
                                _custEnt["hil_productsubcategory"] = Map.hil_ProductSubCategoryMG;
                                try
                                {
                                    _service.Update(_custEnt);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex.Message + " / " + rec.ToString() + "/" + Total.ToString());
                                }
                                _custEnt = new Entity("msdyn_customerasset");
                                _custEnt.Id = CustAsst.Id;
                                _custEnt["hil_createwarranty"] = true;
                                try
                                {
                                    _service.Update(_custEnt);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex.Message + " / " + rec.ToString() + "/" + Total.ToString());
                                }
                            }
                        }
                    }
                }
                Console.WriteLine("Job End / " + DateTime.Now.ToString());
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + " / " + rec.ToString() + "/" + Total.ToString());
            }
        }
        static void CreateSAWActivityApprovals() {
            string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
            <entity name='hil_sawactivity'>
                <attribute name='hil_sawactivityid' />
                <attribute name='hil_name' />
                <attribute name='createdon' />
            <order attribute='createdon' descending='false' />
            <link-entity name='hil_sawactivityapproval' from='hil_sawactivity' to='hil_sawactivityid' link-type='outer' alias='ab' />
            <filter type='and'>
            <condition entityname='ab' attribute='hil_sawactivity' operator='null' />
            </filter>
            </entity>
            </fetch>";
            int i = 0;
            while (true)
            {
                EntityCollection entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entCol.Entities.Count == 0) { break; }
                foreach (Entity ent in entCol.Entities)
                {
                    CommonLib cmnLib = new CommonLib();
                    cmnLib.CreateSAWActivityApprovals(ent.Id, _service);
                    Console.WriteLine(i++.ToString());
                }
            }
        }

        static void RefreshSAWActivity()
        {
            string _strCreatedOn = string.Empty;
            DateTime createdOn;
            EntityReference erCustomerAsset = null;
            //Guid JobId = new Guid("B13EFFB5-6EE3-EA11-A813-000D3AF05D7B");
            var moreRecords = false;
            int page = 1;
            var cookie = string.Empty;

            //string _fetchXML = @"<fetch {0}>
            //    <entity name='msdyn_workorder'>
            //    <attribute name='msdyn_name' />
            //    <attribute name='createdon' />
            //    <attribute name='msdyn_customerasset' />
            //    <attribute name='hil_laborinwarranty' />
            //    <attribute name='hil_typeofassignee' />
            //    <attribute name='msdyn_workorderid' />
            //    <attribute name='hil_isgascharged' />
            //    <attribute name='hil_callsubtype' />
            //    <attribute name='hil_productsubcategory' />
            //    <attribute name='hil_customerref' />
            //    <attribute name='hil_claimstatus' />
            //    <attribute name='hil_isocr' />
            //    <order attribute='msdyn_name' descending='false' />
            //    <filter type='and'>
            //        <filter type='or'>
            //        <condition attribute='msdyn_substatus' operator='in'>
            //            <value uiname='Work Done' uitype='msdyn_workordersubstatus'>{1}</value>
            //            <value uiname='Closed' uitype='msdyn_workordersubstatus'>{2}</value>
            //        </condition>
            //        <condition attribute='hil_isgascharged' operator='eq' value='1' />
            //        </filter>
            //        <condition attribute='msdyn_timeclosed' operator='on' value='2020-09-18' />
            //        <condition attribute='hil_isocr' operator='ne' value='1' />
            //    </filter>
            //    </entity>
            //    </fetch>";
            string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='msdyn_workorder'>
                    <attribute name='msdyn_name' />
                    <attribute name='createdon' />
                    <attribute name='msdyn_customerasset' />
                    <attribute name='hil_laborinwarranty' />
                    <attribute name='hil_typeofassignee' />
                    <attribute name='msdyn_workorderid' />
                    <attribute name='hil_isgascharged' />
                    <attribute name='hil_callsubtype' />
                    <attribute name='hil_productsubcategory' />
                    <attribute name='hil_customerref' />
                    <attribute name='hil_claimstatus' />
                    <attribute name='hil_isocr' />
                <order attribute='msdyn_name' descending='false' />
                <filter type='and'>
                    <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                    <condition attribute='hil_isgascharged' operator='eq' value='1' />
                </filter>
                </entity>
                </fetch>";

            int i = 0;
            do
            {
                //var xml = string.Format(_fetchxml, cookie, "2927FA6C-FA0F-E911-A94E-000D3AF060A1", "1727FA6C-FA0F-E911-A94E-000D3AF060A1");
                //EntityCollection entCol = _service.RetrieveMultiple(new FetchExpression(xml));
                EntityCollection entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                moreRecords = entCol.MoreRecords;
                foreach (Entity _enJob in entCol.Entities)
                {
                    bool _underReview = false;
                    bool _laborinwarranty = false;
                    bool _isOCR = false;
                    bool _gasCharged = false;
                    bool _laborInWarranty = false;

                    bool _gaschargepostaudit = false;
                    bool _gaschargepreauditpastinstallationhistory = false;
                    bool _gaschargepreauditpastrepeathistory = false;
                    bool _repeatrepair = false;

                    if (_enJob != null)
                    {
                        if (_enJob.Attributes.Contains("hil_laborinwarranty"))
                        {
                            _laborinwarranty = _enJob.GetAttributeValue<bool>("hil_laborinwarranty");
                        }
                        if (_enJob.Attributes.Contains("createdon"))
                        {
                            createdOn = _enJob.GetAttributeValue<DateTime>("createdon").AddDays(-90);
                            _strCreatedOn = createdOn.Year.ToString() + "-" + createdOn.Month.ToString().PadLeft(2, '0') + "-" + createdOn.Day.ToString().PadLeft(2, '0');
                        }
                        if (_enJob.Attributes.Contains("msdyn_customerasset"))
                        {
                            erCustomerAsset = _enJob.GetAttributeValue<EntityReference>("msdyn_customerasset");
                        }
                        if (_enJob.Attributes.Contains("hil_isgascharged"))
                        {
                            _gasCharged = (bool)_enJob["hil_isgascharged"];
                        }
                    }
                    if (_laborinwarranty && _gasCharged)
                    {
                        _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorder'>
                        <attribute name='msdyn_name' />
                        <filter type='and'>
                            <condition attribute='msdyn_customerasset' operator='eq' value='" + erCustomerAsset.Id + @"' />
                            <condition attribute='hil_callsubtype' operator='eq' value='{E3129D79-3C0B-E911-A94E-000D3AF06CD4}' />
                            <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strCreatedOn + @"' />
                            <condition attribute='msdyn_substatus' operator='in'>
                            <value >{2927FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                            <value >{7E85074C-9C54-E911-A951-000D3AF0677F}</value>
                            <value >{1727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                        </condition>
                        </filter>
                        </entity>
                        </fetch>";

                        EntityCollection entCol1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entCol1.Entities.Count > 0)
                        {
                            string _remarks = "Old Job# " + entCol1.Entities[0].GetAttributeValue<string>("msdyn_name");
                            CommonLib obj = new CommonLib();
                            CommonLib objReturn = obj.CreateSAWActivity(_enJob.Id, 0, SAWCategoryConst._GasChargePreAuditPastInstallationHistory, _service, _remarks, entCol1.Entities[0].ToEntityReference());
                            if (objReturn.statusRemarks == "OK")
                            {
                                _underReview = true;
                                _gaschargepreauditpastinstallationhistory = true;
                            }
                        }

                        _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='msdyn_workorder'>
                        <attribute name='msdyn_name' />
                        <filter type='and'>
                          <condition attribute='msdyn_customerasset' operator='eq' value='" + erCustomerAsset.Id + @"' />
                          <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strCreatedOn + @"' />
                          <condition attribute='msdyn_workorderid' operator='ne' value='" + _enJob.Id + @"' />
                          <condition attribute='msdyn_substatus' operator='in'>
                            <value >{2927FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                            <value >{7E85074C-9C54-E911-A951-000D3AF0677F}</value>
                            <value >{1727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                          </condition>
                          <condition attribute='hil_isgascharged' operator='eq' value='1' />
                        </filter>
                      </entity>
                    </fetch>";

                        EntityCollection entCol2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entCol2.Entities.Count > 0)
                        {
                            string _remarks = "Old Job# " + entCol2.Entities[0].GetAttributeValue<string>("msdyn_name");
                            CommonLib obj = new CommonLib();
                            CommonLib objReturn = obj.CreateSAWActivity(_enJob.Id, 0, SAWCategoryConst._GasChargePreAuditPastRepeatHistory, _service, _remarks, entCol2.Entities[0].ToEntityReference());
                            if (objReturn.statusRemarks == "OK")
                            {
                                _underReview = true;
                                _gaschargepreauditpastrepeathistory = true;
                            }
                        }
                        if (_underReview)
                        {
                            Entity Ticket = new Entity("msdyn_workorder");
                            Ticket.Id = _enJob.Id;
                            Ticket["hil_claimstatus"] = new OptionSetValue(1); //Claim Under Review
                            _service.Update(Ticket);
                        }
                    }

                    #region Post Closure SAW Activity Approvals
                    OptionSetValue optVal = null;
                    EntityReference erTypeOfAssignee = null;

                    if (_enJob.Attributes.Contains("hil_typeofassignee"))
                    {
                        erTypeOfAssignee = _enJob.GetAttributeValue<EntityReference>("hil_typeofassignee");
                    }
                    if (_enJob.Attributes.Contains("hil_claimstatus"))
                    {
                        optVal = _enJob.GetAttributeValue<OptionSetValue>("hil_claimstatus");
                    }
                    if (_enJob.Attributes.Contains("hil_isocr"))
                    {
                        _isOCR = (bool)_enJob["hil_isocr"];
                    }
                    if (_enJob.Attributes.Contains("hil_isgascharged"))
                    {
                        _gasCharged = (bool)_enJob["hil_isgascharged"];
                    }
                    if (_enJob.Attributes.Contains("hil_laborinwarranty"))
                    {
                        _laborInWarranty = (bool)_enJob["hil_laborinwarranty"];
                    }
                    if (!_isOCR)
                    {
                        if (_laborInWarranty && erTypeOfAssignee.Id != new Guid("7D1ECBAB-1208-E911-A94D-000D3AF0694E")) // LaborInWarranty and Type of Assignee !=DSE
                        {
                            #region Gas Charged Approval
                            if (_gasCharged)
                            {
                                CommonLib obj = new CommonLib();
                                CommonLib objReturn = obj.CreateSAWActivity(_enJob.Id, 0, SAWCategoryConst._GasChargePostAudit, _service, "", null);
                                if (objReturn.statusRemarks == "OK")
                                {
                                    _underReview = true;
                                    _gaschargepostaudit = true;
                                }
                            }
                            #endregion

                            #region RepeatRepair Approval
                            //DateTime _createdOn = _enJob.GetAttributeValue<DateTime>("createdon").AddDays(-15);
                            //DateTime _ClosedOn = DateTime.Now.AddDays(-15);
                            //_strCreatedOn = _createdOn.Year.ToString() + "-" + _createdOn.Month.ToString() + "-" + _createdOn.Day.ToString();
                            //string _strClosedOn = _ClosedOn.Year.ToString() + "-" + _ClosedOn.Month.ToString() + "-" + _ClosedOn.Day.ToString();
                            //EntityCollection entCol3;
                            //_fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            //<entity name='msdyn_workorder'>
                            //<attribute name='msdyn_name' />
                            //<attribute name='createdon' />
                            //<attribute name='hil_productsubcategory' />
                            //<attribute name='hil_customerref' />
                            //<attribute name='hil_callsubtype' />
                            //<attribute name='msdyn_workorderid' />
                            //<attribute name='msdyn_timeclosed' />
                            //<attribute name='msdyn_closedby' />
                            //<order attribute='msdyn_name' descending='false' />
                            //<filter type='and'>
                            //    <condition attribute='hil_isocr' operator='ne' value='1' />
                            //    <condition attribute='hil_typeofassignee' operator='ne' value='{7D1ECBAB-1208-E911-A94D-000D3AF0694E}' />
                            //    <condition attribute='msdyn_workorderid' operator='ne' value='" + _enJob.Id + @"' />
                            //    <condition attribute='hil_customerref' operator='eq' value='" + _enJob.GetAttributeValue<EntityReference>("hil_customerref").Id + @"' />
                            //    <condition attribute='hil_callsubtype' operator='eq' value='" + _enJob.GetAttributeValue<EntityReference>("hil_callsubtype").Id + @"' />
                            //    <condition attribute='hil_callsubtype' operator='ne' value='{8D80346B-3C0B-E911-A94E-000D3AF06CD4}' />
                            //    <condition attribute='hil_productsubcategory' operator='eq' value='" + _enJob.GetAttributeValue<EntityReference>("hil_productsubcategory").Id + @"' />
                            //    <filter type='or'>
                            //        <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strCreatedOn + @"' />
                            //        <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strClosedOn + @"' />
                            //    </filter>
                            //    <condition attribute='msdyn_substatus' operator='in'>
                            //    <value>{1727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                            //    <value>{2927FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                            //    <value>{7E85074C-9C54-E911-A951-000D3AF0677F}</value>
                            //    </condition>
                            //</filter>
                            //</entity>
                            //</fetch>";
                            //entCol3 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            //if (entCol3.Entities.Count > 0)
                            //{
                            //    string _remarks = "Old Job# " + entCol3.Entities[0].GetAttributeValue<string>("msdyn_name");
                            //    CommonLib obj = new CommonLib();
                            //    CommonLib objReturn = obj.CreateSAWActivity(_enJob.Id, 0, SAWCategoryConst._RepeatRepair, _service, _remarks, entCol3.Entities[0].ToEntityReference());
                            //    if (objReturn.statusRemarks == "OK")
                            //    {
                            //        _underReview = true;
                            //        _repeatrepair = true;
                            //    }
                            //}
                            #endregion

                            if (_underReview)
                            {
                                if (optVal != null && optVal.Value != 3)
                                {
                                    Entity Ticket = new Entity("msdyn_workorder");
                                    Ticket.Id = _enJob.Id;
                                    Ticket["hil_claimstatus"] = new OptionSetValue(1); //Claim Under Review
                                    _service.Update(Ticket);
                                }
                            }
                            else
                            {
                                QueryExpression Query;
                                EntityCollection entcoll;
                                Query = new QueryExpression("hil_sawactivity");
                                Query.ColumnSet = new ColumnSet("hil_sawactivityid");
                                Query.Criteria = new FilterExpression(LogicalOperator.And);
                                Query.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _enJob.Id);
                                entcoll = _service.RetrieveMultiple(Query);
                                if (entcoll.Entities.Count == 0)
                                {
                                    Entity Ticket = new Entity("msdyn_workorder");
                                    Ticket.Id = _enJob.Id;
                                    Ticket["hil_claimstatus"] = new OptionSetValue(4); //Claim Approved
                                    _service.Update(Ticket);
                                }
                            }
                        }
                    }
                    #endregion
                    Console.WriteLine(i++.ToString() + "/" + entCol.Entities.Count.ToString());
                }
                if (moreRecords)
                {
                    page++;
                    cookie = string.Format("paging-cookie='{0}' page='{1}'", System.Security.SecurityElement.Escape(entCol.PagingCookie), page);
                }
            } while (moreRecords);
        }

        static void RefreshRRSAWActivity()
        {
            string _strCreatedOn = string.Empty;
            DateTime createdOn;
            EntityReference erCustomerAsset = null;
            //Guid JobId = new Guid("B13EFFB5-6EE3-EA11-A813-000D3AF05D7B");
            var moreRecords = false;
            int page = 1;
            var cookie = string.Empty;

            string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='hil_refreshjobs'>
                <attribute name='hil_refreshjobsid' />
                <attribute name='hil_name' />
                <attribute name='createdon' />
                <order attribute='hil_name' descending='false' />
                <filter type='and'>
                  <condition attribute='hil_repeatrepair' operator='ne' value='1' />
                </filter>
              </entity>
            </fetch>";
            int i = 0;
            do
            {
                var xml = string.Format(_fetchXML, cookie, "2927FA6C-FA0F-E911-A94E-000D3AF060A1", "1727FA6C-FA0F-E911-A94E-000D3AF060A1");
                EntityCollection entCol = _service.RetrieveMultiple(new FetchExpression(xml));
                moreRecords = entCol.MoreRecords;
                foreach (Entity _enJob in entCol.Entities)
                {
                //    string _fetchXML = @"<fetch {0}>
                //<entity name='msdyn_workorder'>
                //<attribute name='msdyn_name' />
                //<attribute name='createdon' />
                //<attribute name='msdyn_customerasset' />
                //<attribute name='hil_laborinwarranty' />
                //<attribute name='hil_typeofassignee' />
                //<attribute name='msdyn_workorderid' />
                //<attribute name='hil_isgascharged' />
                //<attribute name='hil_callsubtype' />
                //<attribute name='hil_productsubcategory' />
                //<attribute name='hil_customerref' />
                //<attribute name='hil_claimstatus' />
                //<attribute name='hil_isocr' />
                //<order attribute='msdyn_name' descending='false' />
                //<filter type='and'>
                //    <filter type='or'>
                //    <condition attribute='msdyn_substatus' operator='in'>
                //        <value uiname='Work Done' uitype='msdyn_workordersubstatus'>{1}</value>
                //        <value uiname='Closed' uitype='msdyn_workordersubstatus'>{2}</value>
                //    </condition>
                //    <condition attribute='hil_isgascharged' operator='eq' value='1' />
                //    </filter>
                //    <condition attribute='msdyn_timeclosed' operator='on' value='2020-09-19' />
                //    <condition attribute='hil_isocr' operator='ne' value='1' />
                //</filter>
                //</entity>
                //</fetch>";

                    bool _underReview = false;
                    bool _laborinwarranty = false;
                    bool _isOCR = false;
                    bool _gasCharged = false;
                    bool _laborInWarranty = false;

                    if (_enJob != null)
                    {
                        if (_enJob.Attributes.Contains("hil_laborinwarranty"))
                        {
                            _laborinwarranty = _enJob.GetAttributeValue<bool>("hil_laborinwarranty");
                        }
                        if (_enJob.Attributes.Contains("createdon"))
                        {
                            createdOn = _enJob.GetAttributeValue<DateTime>("createdon").AddDays(-90);
                            _strCreatedOn = createdOn.Year.ToString() + "-" + createdOn.Month.ToString().PadLeft(2, '0') + "-" + createdOn.Day.ToString().PadLeft(2, '0');
                        }
                        if (_enJob.Attributes.Contains("msdyn_customerasset"))
                        {
                            erCustomerAsset = _enJob.GetAttributeValue<EntityReference>("msdyn_customerasset");
                        }
                    }
                    #region Post Closure SAW Activity Approvals
                    OptionSetValue optVal = null;
                    EntityReference erTypeOfAssignee = null;

                    if (_enJob.Attributes.Contains("hil_typeofassignee"))
                    {
                        erTypeOfAssignee = _enJob.GetAttributeValue<EntityReference>("hil_typeofassignee");
                    }
                    if (_enJob.Attributes.Contains("hil_claimstatus"))
                    {
                        optVal = _enJob.GetAttributeValue<OptionSetValue>("hil_claimstatus");
                    }
                    if (_enJob.Attributes.Contains("hil_isocr"))
                    {
                        _isOCR = (bool)_enJob["hil_isocr"];
                    }
                    if (_enJob.Attributes.Contains("hil_isgascharged"))
                    {
                        _gasCharged = (bool)_enJob["hil_isgascharged"];
                    }
                    if (_enJob.Attributes.Contains("hil_laborinwarranty"))
                    {
                        _laborInWarranty = (bool)_enJob["hil_laborinwarranty"];
                    }
                    if (!_isOCR)
                    {
                        if (_laborInWarranty && erTypeOfAssignee.Id != new Guid("7D1ECBAB-1208-E911-A94D-000D3AF0694E")) // LaborInWarranty and Type of Assignee !=DSE
                        {
                            #region RepeatRepair Approval
                            DateTime _createdOn = _enJob.GetAttributeValue<DateTime>("createdon").AddDays(-15);
                            DateTime _ClosedOn = DateTime.Now.AddDays(-15);
                            _strCreatedOn = _createdOn.Year.ToString() + "-" + _createdOn.Month.ToString() + "-" + _createdOn.Day.ToString();
                            string _strClosedOn = _ClosedOn.Year.ToString() + "-" + _ClosedOn.Month.ToString() + "-" + _ClosedOn.Day.ToString();
                            EntityCollection entCol3;
                            _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='msdyn_workorder'>
                            <attribute name='msdyn_name' />
                            <attribute name='createdon' />
                            <attribute name='hil_productsubcategory' />
                            <attribute name='hil_customerref' />
                            <attribute name='hil_callsubtype' />
                            <attribute name='msdyn_workorderid' />
                            <attribute name='msdyn_timeclosed' />
                            <attribute name='msdyn_closedby' />
                            <order attribute='msdyn_name' descending='false' />
                            <filter type='and'>
                                <condition attribute='hil_isocr' operator='ne' value='1' />
                                <condition attribute='hil_typeofassignee' operator='ne' value='{7D1ECBAB-1208-E911-A94D-000D3AF0694E}' />
                                <condition attribute='msdyn_workorderid' operator='ne' value='" + _enJob.Id + @"' />
                                <condition attribute='hil_customerref' operator='eq' value='" + _enJob.GetAttributeValue<EntityReference>("hil_customerref").Id + @"' />
                                <condition attribute='hil_callsubtype' operator='eq' value='" + _enJob.GetAttributeValue<EntityReference>("hil_callsubtype").Id + @"' />
                                <condition attribute='hil_callsubtype' operator='ne' value='{8D80346B-3C0B-E911-A94E-000D3AF06CD4}' />
                                <condition attribute='hil_productsubcategory' operator='eq' value='" + _enJob.GetAttributeValue<EntityReference>("hil_productsubcategory").Id + @"' />
                                <filter type='or'>
                                    <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strCreatedOn + @"' />
                                    <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strClosedOn + @"' />
                                </filter>
                                <condition attribute='msdyn_substatus' operator='in'>
                                <value>{1727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                                <value>{2927FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                                <value>{7E85074C-9C54-E911-A951-000D3AF0677F}</value>
                                </condition>
                            </filter>
                            </entity>
                            </fetch>";
                            entCol3 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (entCol3.Entities.Count > 0)
                            {
                                string _remarks = "Old Job# " + entCol3.Entities[0].GetAttributeValue<string>("msdyn_name");
                                CommonLib obj = new CommonLib();
                                CommonLib objReturn = obj.CreateSAWActivity(_enJob.Id, 0, SAWCategoryConst._RepeatRepair, _service, _remarks, entCol3.Entities[0].ToEntityReference());
                                if (objReturn.statusRemarks == "OK")
                                {
                                    _underReview = true;
                                    //_repeatrepair = true;
                                }
                            }
                            #endregion

                            if (_underReview)
                            {
                                if (optVal != null && optVal.Value != 3)
                                {
                                    Entity Ticket = new Entity("msdyn_workorder");
                                    Ticket.Id = _enJob.Id;
                                    Ticket["hil_claimstatus"] = new OptionSetValue(1); //Claim Under Review
                                    _service.Update(Ticket);
                                }
                            }
                            else
                            {
                                QueryExpression Query;
                                EntityCollection entcoll;
                                Query = new QueryExpression("hil_sawactivity");
                                Query.ColumnSet = new ColumnSet("hil_sawactivityid");
                                Query.Criteria = new FilterExpression(LogicalOperator.And);
                                Query.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _enJob.Id);
                                entcoll = _service.RetrieveMultiple(Query);
                                if (entcoll.Entities.Count == 0)
                                {
                                    Entity Ticket = new Entity("msdyn_workorder");
                                    Ticket.Id = _enJob.Id;
                                    Ticket["hil_claimstatus"] = new OptionSetValue(4); //Claim Approved
                                    _service.Update(Ticket);
                                }
                            }
                        }
                    }
                    #endregion
                    Console.WriteLine(i++.ToString() + "/" + entCol.Entities.Count.ToString());
                }
                if (moreRecords)
                {
                    page++;
                    cookie = string.Format("paging-cookie='{0}' page='{1}'", System.Security.SecurityElement.Escape(entCol.PagingCookie), page);
                }
            } while (moreRecords);
        }
        static void RefreshClaimParametersBulk() {
            Guid _jobGuId = new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"); // Closed Call
            Guid _subStatus = new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"); // Closed Call
            EntityReference _unitWarranty = null;
            Guid _warrantyTemplateId = Guid.Empty;
            int ClosureCount = 0;
            bool SparePartUsed = false;
            DateTime _fromDate, _toDate;
            string _fetchXML = string.Empty;
            QueryExpression qrExp;
            EntityCollection entCol;
            bool laborInWarranty = false;
            int _jobWarrantyStatus = 2; //OutWarranty
            int _jobWarrantySubStatus = 0;
            int _warrantyTempType = 0;
            DateTime _unitWarrStartDate = new DateTime(1900, 1, 1);
            double _jobMonth = 0;


            if (DateTime.Now.Day >= 21)
            {
                _fromDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 21);
                _toDate = _fromDate.AddDays(30).AddDays(1).AddMinutes(-1);
            }
            else
            {
                _toDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 20).AddDays(1).AddMinutes(-1);
                _fromDate = _toDate.AddDays(-30).AddDays(-1);
            }

            qrExp = new QueryExpression("hil_claimperiod");
            qrExp.ColumnSet = new ColumnSet("hil_claimperiodid");
            qrExp.Criteria = new FilterExpression(LogicalOperator.And);
            qrExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _jobGuId);
            qrExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, 0);
            EntityCollection entColCP = _service.RetrieveMultiple(qrExp);

            //2008206063037
            _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='msdyn_workorder'>
                    <attribute name='msdyn_name' />
                    <attribute name='createdon' />
                    <attribute name='msdyn_timeclosed' />
                    <attribute name='hil_owneraccount' />
                    <attribute name='msdyn_customerasset' />
                    <attribute name='hil_purchasedate' />
                    <order attribute='msdyn_timeclosed' descending='false' />
                    <filter type='and'>
                        <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                        <condition attribute='hil_isocr' operator='ne' value='1' />
                        <condition attribute='msdyn_timeclosed' operator='on-or-after' value='2020-08-21' />
                        <condition attribute='msdyn_timeclosed' operator='on-or-before' value='2020-09-18' />
                        <condition attribute='hil_systemremarks' operator='not-like' value='%CPARAM%' />
                        <condition attribute='msdyn_name' operator='eq' value='2008206063037' />
                    </filter>
                    <link-entity name='account' from='accountid' to='hil_owneraccount' visible='false' link-type='outer' alias='custCatg'>
                    <attribute name='hil_category' />
                    </link-entity>
                  </entity>
                </fetch>";

            while (true)
            {
                entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entCol.Entities.Count == 0) { break; }
                int i = 1;
                foreach (Entity ent in entCol.Entities)
                {
                    _jobGuId = ent.Id;
                    if (ent != null)
                    {
                        _unitWarranty = null;
                        _warrantyTemplateId = Guid.Empty;
                        SparePartUsed = false;
                        laborInWarranty = false;
                        _jobWarrantyStatus = 2; //OutWarranty
                        _jobWarrantySubStatus = 0;
                        _warrantyTempType = 0;
                        _unitWarrStartDate = new DateTime(1900, 1, 1);

                        DateTime _jobCreatedOn = ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330);

                        DateTime _assetPurchaseDate = ent.GetAttributeValue<DateTime>("hil_purchasedate");
                        if (ent.Attributes.Contains("hil_purchasedate")) { }

                        DateTime _jobClosedOn = ent.GetAttributeValue<DateTime>("msdyn_timeclosed").AddMinutes(330);
                        if (_jobCreatedOn < _assetPurchaseDate)
                        {
                            _jobCreatedOn = _assetPurchaseDate;
                        }
                        qrExp = new QueryExpression("msdyn_workorderincident");
                        qrExp.ColumnSet = new ColumnSet("msdyn_workorderincidentid");
                        qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qrExp.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, _jobGuId);
                        qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        qrExp.Criteria.AddCondition("hil_warrantystatus", ConditionOperator.Equal, 3); // Warranty Void
                        EntityCollection entCol1 = _service.RetrieveMultiple(qrExp);
                        if (entCol1.Entities.Count == 0)
                        {
                            qrExp = new QueryExpression("hil_unitwarranty");
                            qrExp.ColumnSet = new ColumnSet("hil_name", "hil_warrantytemplate", "hil_warrantystartdate", "hil_warrantyenddate");
                            qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                            qrExp.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("msdyn_customerasset").Id);
                            qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            EntityCollection entCol2 = _service.RetrieveMultiple(qrExp);
                            if (entCol2.Entities.Count > 0)
                            {
                                foreach (Entity Wt in entCol2.Entities)
                                {
                                    DateTime iValidTo = ((DateTime)Wt["hil_warrantyenddate"]).AddMinutes(330);
                                    DateTime iValidFrom = ((DateTime)Wt["hil_warrantystartdate"]).AddMinutes(330);
                                    if (_jobCreatedOn >= iValidFrom && _jobCreatedOn <= iValidTo)
                                    {
                                        _jobWarrantyStatus = 1;
                                        _unitWarranty = Wt.ToEntityReference();
                                        _warrantyTemplateId = Wt.GetAttributeValue<EntityReference>("hil_warrantytemplate").Id;
                                        Entity _entTemp = _service.Retrieve(hil_warrantytemplate.EntityLogicalName, _warrantyTemplateId, new ColumnSet("hil_name", "hil_type"));
                                        if (_entTemp != null)
                                        {
                                            _warrantyTempType = _entTemp.GetAttributeValue<OptionSetValue>("hil_type").Value;
                                            if (_warrantyTempType == 1) { _jobWarrantySubStatus = 1; }
                                            else if (_warrantyTempType == 2) { _jobWarrantySubStatus = 2; }
                                            else if (_warrantyTempType == 7) { _jobWarrantySubStatus = 3; }
                                            else if (_warrantyTempType == 3) { _jobWarrantySubStatus = 4; }
                                        }
                                        _unitWarrStartDate = Wt.GetAttributeValue<DateTime>("hil_warrantystartdate").AddMinutes(330);
                                        TimeSpan difference = (_jobCreatedOn - _unitWarrStartDate);
                                        _jobMonth = Math.Ceiling((difference.Days * 1.0 / 30.42));
                                        qrExp = new QueryExpression("hil_labor");
                                        qrExp.ColumnSet = new ColumnSet("hil_laborid", "hil_includedinwarranty", "hil_validtomonths", "hil_validfrommonths");
                                        qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                                        qrExp.Criteria.AddCondition("hil_warrantytemplateid", ConditionOperator.Equal, _warrantyTemplateId);
                                        qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                        EntityCollection entCol3 = _service.RetrieveMultiple(qrExp);
                                        if (entCol3.Entities.Count == 0) { laborInWarranty = true; }
                                        else
                                        {
                                            if (_jobMonth >= entCol3.Entities[0].GetAttributeValue<int>("hil_validfrommonths") && _jobMonth <= entCol3.Entities[0].GetAttributeValue<int>("hil_validtomonths"))
                                            {
                                                OptionSetValue _laborType = entCol3.Entities[0].GetAttributeValue<OptionSetValue>("hil_includedinwarranty");
                                                laborInWarranty = _laborType.Value == 1 ? true : false;
                                            }
                                            else
                                            {
                                                OptionSetValue _laborType = entCol3.Entities[0].GetAttributeValue<OptionSetValue>("hil_includedinwarranty");
                                                laborInWarranty = !(_laborType.Value == 1 ? true : false);
                                            }
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                        else
                        {
                            _jobWarrantyStatus = 3;
                        }

                        qrExp = new QueryExpression("msdyn_workorderproduct");
                        qrExp.ColumnSet = new ColumnSet("msdyn_workorderproductid");
                        qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qrExp.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, _jobGuId);
                        qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        qrExp.Criteria.AddCondition("hil_markused", ConditionOperator.Equal, true);
                        EntityCollection entCol4 = _service.RetrieveMultiple(qrExp);
                        if (entCol4.Entities.Count > 0) { SparePartUsed = true; } else { SparePartUsed = false; }

                        bool _claimstatus = false;
                        OptionSetValue _claimStatusVal = new OptionSetValue(4);

                        qrExp = new QueryExpression("hil_sawactivity");
                        qrExp.ColumnSet = new ColumnSet("hil_sawactivityid");
                        qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qrExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _jobGuId);
                        EntityCollection entCol5 = _service.RetrieveMultiple(qrExp);
                        if (entCol5.Entities.Count == 0)
                        {
                            _claimStatusVal = new OptionSetValue(4);
                        }
                        else
                        {
                            qrExp = new QueryExpression("hil_sawactivityapproval");
                            qrExp.ColumnSet = new ColumnSet("hil_sawactivityapprovalid");
                            qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                            qrExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _jobGuId);
                            qrExp.Criteria.AddCondition("hil_approvalstatus", ConditionOperator.NotIn, new object[] { 3, 4 });
                            EntityCollection entCol6 = _service.RetrieveMultiple(qrExp);
                            if (entCol6.Entities.Count > 0)
                            {
                                _claimStatusVal = new OptionSetValue(1); //Under Review
                            }
                            else
                            {
                                qrExp = new QueryExpression("hil_sawactivityapproval");
                                qrExp.ColumnSet = new ColumnSet("hil_sawactivityapprovalid");
                                qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                                qrExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _jobGuId);
                                qrExp.Criteria.AddCondition("hil_approvalstatus", ConditionOperator.Equal, 3);
                                EntityCollection entCol7 = _service.RetrieveMultiple(qrExp);
                                if (entCol7.Entities.Count > 0)
                                {
                                    _claimStatusVal = new OptionSetValue(4); //Approved
                                }
                                else
                                {
                                    _claimStatusVal = new OptionSetValue(3); //Rejected
                                }
                            }
                        }
                        Entity entJobUpdate = new Entity(msdyn_workorder.EntityLogicalName, _jobGuId);

                        entJobUpdate["hil_warrantystatus"] = new OptionSetValue(_jobWarrantyStatus);
                        if (_jobWarrantyStatus == 1)
                        {
                            entJobUpdate["hil_warrantysubstatus"] = new OptionSetValue(_jobWarrantySubStatus);
                        }
                        else
                        {
                            entJobUpdate["hil_warrantysubstatus"] = null;
                        }
                        if (_unitWarranty != null)
                        {
                            entJobUpdate["hil_unitwarranty"] = _unitWarranty;
                        }
                        else { entJobUpdate["hil_unitwarranty"] = null; }

                        entJobUpdate["hil_laborinwarranty"] = laborInWarranty;
                        if (_assetPurchaseDate.Year != 1900 && _assetPurchaseDate.Year != 1)
                        {
                            entJobUpdate["hil_purchasedate"] = _assetPurchaseDate;
                        }
                        else
                        {
                            entJobUpdate["hil_purchasedate"] = null;
                        }

                        entJobUpdate["hil_sparepartuse"] = SparePartUsed;

                        entJobUpdate["hil_fiscalmonth"] = new EntityReference("hil_claimperiod", new Guid("ad387612-46e8-ea11-a817-000d3af05a4b"));
                        if (ent.Attributes.Contains("custCatg.hil_category"))
                        {
                            OptionSetValue opVal = (OptionSetValue)((AliasedValue)ent.Attributes["custCatg.hil_category"]).Value;
                            entJobUpdate["hil_channelpartnercategory"] = new OptionSetValue(opVal.Value);
                        }
                        entJobUpdate["hil_systemremarks"] = "CPARAM";
                        entJobUpdate["hil_claimstatus"] = _claimStatusVal;
                        try
                        {
                            _service.Update(entJobUpdate);
                        }
                        catch {}
                        
                        Console.WriteLine(i++.ToString() + "/" + entCol.Entities.Count.ToString());
                    }
                }
            }
        }

        static void RefreshClaimParameters()
        {
            Guid _jobGuId = new Guid("2AF03FB1-CBE2-EA11-A813-000D3AF05D7B");
            Guid _subStatus = new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"); // Closed Call
            EntityReference _unitWarranty = null;
            Guid _warrantyTemplateId = Guid.Empty;
            bool SparePartUsed = false;
            string _fetchXML = string.Empty;
            QueryExpression qrExp;
            bool laborInWarranty = false;
            int _jobWarrantyStatus = 2; //OutWarranty
            int _jobWarrantySubStatus = 0;
            int _warrantyTempType = 0;
            DateTime _unitWarrStartDate = new DateTime(1900, 1, 1);
            double _jobMonth = 0;
            EntityReference _claimPeriod = null;

            Entity ent = _service.Retrieve("msdyn_workorder", _jobGuId, new ColumnSet("hil_isocr", "msdyn_name", "createdon", "msdyn_timeclosed", "hil_owneraccount", "msdyn_customerasset", "hil_purchasedate"));

            if (ent != null)
            {
                bool _isOCR = false;
                EntityReference erCustomerAsset = null;

                if (ent.Attributes.Contains("msdyn_customerasset"))
                {
                    erCustomerAsset = ent.GetAttributeValue<EntityReference>("msdyn_customerasset");
                }

                if (ent.Attributes.Contains("hil_isocr"))
                {
                    _isOCR = ent.GetAttributeValue<bool>("hil_isocr");
                }
                if (!_isOCR && erCustomerAsset != null)
                {
                    DateTime _jobClosedOn = ent.GetAttributeValue<DateTime>("msdyn_timeclosed").AddMinutes(330);
                    DateTime _tempDate = new DateTime(_jobClosedOn.Year, _jobClosedOn.Month, _jobClosedOn.Day);
                    qrExp = new QueryExpression("hil_claimperiod");
                    qrExp.ColumnSet = new ColumnSet("hil_claimperiodid");
                    qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                    qrExp.Criteria.AddCondition("hil_fromdate", ConditionOperator.OnOrBefore, _tempDate);
                    qrExp.Criteria.AddCondition("hil_todate", ConditionOperator.OnOrAfter, _tempDate);
                    EntityCollection entColCP = _service.RetrieveMultiple(qrExp);
                    if (entColCP != null)
                    {
                        _claimPeriod = entColCP.Entities[0].ToEntityReference();
                        _unitWarranty = null;
                        _warrantyTemplateId = Guid.Empty;
                        SparePartUsed = false;
                        laborInWarranty = false;
                        _jobWarrantyStatus = 2; //OutWarranty
                        _jobWarrantySubStatus = 0;
                        _warrantyTempType = 0;
                        _unitWarrStartDate = new DateTime(1900, 1, 1);

                        DateTime _jobCreatedOn = ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                        DateTime _assetPurchaseDate = new DateTime(1900, 1, 1);

                        Entity entCustAsset = _service.Retrieve("msdyn_customerasset", ent.GetAttributeValue<EntityReference>("msdyn_customerasset").Id, new ColumnSet("hil_invoicedate", "hil_invoiceavailable"));
                        if (entCustAsset != null) {
                            if (entCustAsset.Attributes.Contains("hil_invoicedate")) {
                                _assetPurchaseDate = entCustAsset.GetAttributeValue<DateTime>("hil_invoicedate");
                            }
                        }

                        if (_jobCreatedOn < _assetPurchaseDate)
                        {
                            _jobCreatedOn = _assetPurchaseDate;
                        }

                        qrExp = new QueryExpression("msdyn_workorderincident");
                        qrExp.ColumnSet = new ColumnSet("msdyn_workorderincidentid");
                        qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qrExp.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, _jobGuId);
                        qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        qrExp.Criteria.AddCondition("hil_warrantystatus", ConditionOperator.Equal, 3); // Warranty Void
                        EntityCollection entCol1 = _service.RetrieveMultiple(qrExp);
                        if (entCol1.Entities.Count == 0)
                        {
                            qrExp = new QueryExpression("hil_unitwarranty");
                            qrExp.ColumnSet = new ColumnSet("hil_name", "hil_warrantytemplate", "hil_warrantystartdate", "hil_warrantyenddate");
                            qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                            qrExp.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("msdyn_customerasset").Id);
                            qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            EntityCollection entCol2 = _service.RetrieveMultiple(qrExp);
                            if (entCol2.Entities.Count > 0)
                            {
                                foreach (Entity Wt in entCol2.Entities)
                                {
                                    DateTime iValidTo = ((DateTime)Wt["hil_warrantyenddate"]).AddMinutes(330);
                                    DateTime iValidFrom = ((DateTime)Wt["hil_warrantystartdate"]).AddMinutes(330);
                                    if (_jobCreatedOn >= iValidFrom && _jobCreatedOn <= iValidTo)
                                    {
                                        _jobWarrantyStatus = 1;
                                        _unitWarranty = Wt.ToEntityReference();
                                        _warrantyTemplateId = Wt.GetAttributeValue<EntityReference>("hil_warrantytemplate").Id;
                                        Entity _entTemp = _service.Retrieve(hil_warrantytemplate.EntityLogicalName, _warrantyTemplateId, new ColumnSet("hil_name", "hil_type"));
                                        if (_entTemp != null)
                                        {
                                            _warrantyTempType = _entTemp.GetAttributeValue<OptionSetValue>("hil_type").Value;
                                            if (_warrantyTempType == 1) { _jobWarrantySubStatus = 1; }
                                            else if (_warrantyTempType == 2) { _jobWarrantySubStatus = 2; }
                                            else if (_warrantyTempType == 7) { _jobWarrantySubStatus = 3; }
                                            else if (_warrantyTempType == 3) { _jobWarrantySubStatus = 4; }
                                        }
                                        _unitWarrStartDate = Wt.GetAttributeValue<DateTime>("hil_warrantystartdate").AddMinutes(330);
                                        TimeSpan difference = (_jobCreatedOn - _unitWarrStartDate);
                                        _jobMonth = Math.Ceiling((difference.Days * 1.0 / 30.42));
                                        qrExp = new QueryExpression("hil_labor");
                                        qrExp.ColumnSet = new ColumnSet("hil_laborid", "hil_includedinwarranty", "hil_validtomonths", "hil_validfrommonths");
                                        qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                                        qrExp.Criteria.AddCondition("hil_warrantytemplateid", ConditionOperator.Equal, _warrantyTemplateId);
                                        qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                        EntityCollection entCol3 = _service.RetrieveMultiple(qrExp);
                                        if (entCol3.Entities.Count == 0) { laborInWarranty = true; }
                                        else
                                        {
                                            if (_jobMonth >= entCol3.Entities[0].GetAttributeValue<int>("hil_validfrommonths") && _jobMonth <= entCol3.Entities[0].GetAttributeValue<int>("hil_validtomonths"))
                                            {
                                                OptionSetValue _laborType = entCol3.Entities[0].GetAttributeValue<OptionSetValue>("hil_includedinwarranty");
                                                laborInWarranty = _laborType.Value == 1 ? true : false;
                                            }
                                            else
                                            {
                                                OptionSetValue _laborType = entCol3.Entities[0].GetAttributeValue<OptionSetValue>("hil_includedinwarranty");
                                                laborInWarranty = !(_laborType.Value == 1 ? true : false);
                                            }
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                        else
                        {
                            _jobWarrantyStatus = 3;
                        }

                        qrExp = new QueryExpression("msdyn_workorderproduct");
                        qrExp.ColumnSet = new ColumnSet("msdyn_workorderproductid");
                        qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qrExp.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, _jobGuId);
                        qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        qrExp.Criteria.AddCondition("hil_markused", ConditionOperator.Equal, true);
                        EntityCollection entCol4 = _service.RetrieveMultiple(qrExp);
                        if (entCol4.Entities.Count > 0) { SparePartUsed = true; } else { SparePartUsed = false; }

                        //bool _claimstatus = false;
                        OptionSetValue _claimStatusVal = new OptionSetValue(4);

                        qrExp = new QueryExpression("hil_sawactivity");
                        qrExp.ColumnSet = new ColumnSet("hil_sawactivityid");
                        qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qrExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _jobGuId);
                        EntityCollection entCol5 = _service.RetrieveMultiple(qrExp);
                        if (entCol5.Entities.Count == 0)
                        {
                            _claimStatusVal = new OptionSetValue(4); // Claim Approved
                        }
                        else
                        {
                            qrExp = new QueryExpression("hil_sawactivityapproval");
                            qrExp.ColumnSet = new ColumnSet("hil_sawactivityapprovalid");
                            qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                            qrExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _jobGuId);
                            qrExp.Criteria.AddCondition("hil_approvalstatus", ConditionOperator.Equal, 4);// Rejected
                            EntityCollection entCol6 = _service.RetrieveMultiple(qrExp);
                            if (entCol6.Entities.Count > 0)
                            {
                                _claimStatusVal = new OptionSetValue(3); //Rejected
                            }
                            else
                            {
                                qrExp = new QueryExpression("hil_sawactivityapproval");
                                qrExp.ColumnSet = new ColumnSet("hil_sawactivityapprovalid");
                                qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                                qrExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _jobGuId);
                                qrExp.Criteria.AddCondition("hil_approvalstatus", ConditionOperator.NotEqual, 3); //Approved
                                EntityCollection entCol7 = _service.RetrieveMultiple(qrExp);
                                if (entCol7.Entities.Count == 0)
                                {
                                    _claimStatusVal = new OptionSetValue(4); //Approved
                                }
                                else
                                {
                                    _claimStatusVal = new OptionSetValue(1); //Under Review
                                }
                            }
                        }
                        Entity entJobUpdate = new Entity(msdyn_workorder.EntityLogicalName, _jobGuId);

                        entJobUpdate["hil_warrantystatus"] = new OptionSetValue(_jobWarrantyStatus);
                        if (_jobWarrantyStatus == 1)
                        {
                            entJobUpdate["hil_warrantysubstatus"] = new OptionSetValue(_jobWarrantySubStatus);
                        }
                        else
                        {
                            entJobUpdate["hil_warrantysubstatus"] = null;
                        }
                        if (_unitWarranty != null)
                        {
                            entJobUpdate["hil_unitwarranty"] = _unitWarranty;
                        }
                        else { entJobUpdate["hil_unitwarranty"] = null; }

                        entJobUpdate["hil_laborinwarranty"] = laborInWarranty;

                        if (_assetPurchaseDate.Year != 1900 && _assetPurchaseDate.Year != 1)
                        {
                            entJobUpdate["hil_purchasedate"] = _assetPurchaseDate;
                        }
                        else
                        {
                            entJobUpdate["hil_purchasedate"] = null;
                        }

                        entJobUpdate["hil_sparepartuse"] = SparePartUsed;

                        entJobUpdate["hil_fiscalmonth"] = _claimPeriod;
                        if (ent.Attributes.Contains("custCatg.hil_category"))
                        {
                            OptionSetValue opVal = (OptionSetValue)((AliasedValue)ent.Attributes["custCatg.hil_category"]).Value;
                            entJobUpdate["hil_channelpartnercategory"] = new OptionSetValue(opVal.Value);
                        }
                        entJobUpdate["hil_claimstatus"] = _claimStatusVal;
                        try
                        {
                            _service.Update(entJobUpdate);
                        }
                        catch { }
                    }
                }
            }
        }
        static void DailyRefreshCustomerAssetWarranty()
        {
            int totalRecords = 0;
            int recordCnt = 0;
            try
            {
                //Skip Aqua Division
                DateTime wefDate = DateTime.Now.AddDays(-1);
                QueryExpression Query = new QueryExpression(msdyn_customerasset.EntityLogicalName);
                Query.ColumnSet = new ColumnSet("msdyn_name", "hil_warrantystatus");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression("hil_warrantytilldate", ConditionOperator.OnOrBefore, new DateTime(wefDate.Year, wefDate.Month, wefDate.Day)));

                Query.AddOrder("hil_warrantytilldate", OrderType.Descending);
                EntityCollection ec = _service.RetrieveMultiple(Query);
                if (ec.Entities.Count > 0)
                {
                    totalRecords = ec.Entities.Count;
                    recordCnt = 1;
                    foreach (Entity ent in ec.Entities)
                    {
                        ent["hil_warrantystatus"] = new OptionSetValue(2); //OUT
                        _service.Update(ent);
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

        static void UpdateJobUpcountry(IOrganizationService service)
        {
            try
            {
                QueryExpression Query;
                EntityCollection entcoll;
                int recordCnt = 0;
                int totalRecords = 0;
                while (true)
                {
                    Query = new QueryExpression("msdyn_workorder");
                    Query.ColumnSet = new ColumnSet("msdyn_name","hil_productcategory", "hil_pincode", "hil_owneraccount", "hil_callsubtype", "hil_brand");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_brand", ConditionOperator.Equal, 2);
                    Query.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrAfter, new DateTime(2020, 7, 21))); //
                    Query.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrBefore, new DateTime(2020, 8, 07))); //
                    Query.Criteria.AddCondition(new ConditionExpression("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"))); //
                    Query.Criteria.AddCondition(new ConditionExpression("hil_termsconditions", ConditionOperator.Null)); //
                    entcoll = service.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count == 0) {
                        break;
                    }
                    totalRecords += entcoll.Entities.Count;
                    foreach (Entity ent in entcoll.Entities)
                    {
                        Entity Ticket = new Entity("msdyn_workorder");
                        Ticket.Id = ent.Id;
                        if (ent.GetAttributeValue<OptionSetValue>("hil_brand").Value == 2) //LLOYD
                        {
                            if (ent.Attributes.Contains("hil_productcategory") && ent.Attributes.Contains("hil_pincode") && ent.Attributes.Contains("hil_owneraccount") && ent.Attributes.Contains("hil_callsubtype"))
                            {
                                QueryExpression Query1;
                                EntityCollection entcoll1;
                                Query1 = new QueryExpression("hil_assignmentmatrix");
                                Query1.ColumnSet = new ColumnSet("hil_upcountry");
                                Query1.Criteria = new FilterExpression(LogicalOperator.And);
                                Query1.Criteria.AddCondition("hil_division", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_productcategory").Id);
                                Query1.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_pincode").Id);
                                Query1.Criteria.AddCondition("hil_franchiseedirectengineer", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_owneraccount").Id);
                                Query1.Criteria.AddCondition("hil_callsubtype", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_callsubtype").Id);
                                Query1.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                entcoll1 = service.RetrieveMultiple(Query1);
                                if (entcoll1.Entities.Count > 0)
                                {
                                    if (entcoll1.Entities[0].Attributes.Contains("hil_upcountry"))
                                    {
                                        bool flag = entcoll1.Entities[0].GetAttributeValue<bool>("hil_upcountry");
                                        Ticket["hil_countryclassification"] = flag == true ? new OptionSetValue(2) : new OptionSetValue(1);
                                    }
                                    else
                                    {
                                        Ticket["hil_countryclassification"] = new OptionSetValue(1); // Local
                                    }
                                }
                                else
                                {
                                    Ticket["hil_countryclassification"] = new OptionSetValue(2); // Upcountry
                                }
                                Ticket["hil_termsconditions"] = "Done";
                                service.Update(Ticket);
                                Console.WriteLine(ent.GetAttributeValue<string>("msdyn_name").ToString() + " : " + recordCnt.ToString() + "/" + totalRecords.ToString());
                                recordCnt += 1;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Havells_Plugin.WorkOrder.PostUpdate.UpdateJobclosureSourceforMobileApp: " + ex.Message);
            }
        }

        static void RefreshAssetUnitWarranty()
        {

            #region Backup
            //string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
            //<entity name='msdyn_customerasset'>
            //<attribute name='hil_invoicedate' />
            //<attribute name='msdyn_customerassetid' />
            //<attribute name='hil_invoiceavailable' />
            //<attribute name='statuscode' />
            //<attribute name='hil_productcategory' />
            //<attribute name='hil_productsubcategory' />
            //<attribute name='msdyn_product' />
            //<attribute name='hil_modelname' />
            //<attribute name='msdyn_name' />
            //<attribute name='hil_customer' />
            //<order attribute='createdon' descending='true' />
            //<filter type='and'>
            //<condition attribute='hil_customernm' operator='ne' value='DONE' />
            //</filter>
            //<link-entity name='hil_unitwarranty' from='hil_customerasset' to='msdyn_customerassetid' link-type='inner' alias='ad'>
            //    <filter type='and'>
            //    <condition attribute='hil_customer' operator='not-null' />
            //    </filter>
            //    <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='ae'>
            //    <filter type='and'>
            //        <condition attribute='modifiedby' operator='eq' uiname='Rahul Panchal' uitype='systemuser' value='{FAC2D349-2AFD-E811-A94C-000D3AF060A1}' />
            //        <condition attribute='modifiedon' operator='on-or-after' value='2020-08-10' />
            //        <condition attribute='createdon' operator='on-or-before' value='2020-08-09' />
            //    </filter>
            //    <link-entity name='product' from='productid' to='hil_product' link-type='inner' alias='af'>
            //        <filter type='and'>
            //        <condition attribute='hil_division' operator='eq' uiname='LLOYD LED TELEVISION' uitype='product' value='{A7A5049B-16FA-E811-A94C-000D3AF06091}' />
            //        </filter>
            //    </link-entity>
            //    </link-entity>
            //</link-entity>
            //</entity>
            //</fetch>";

            //string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            //<entity name='msdyn_customerasset'>
            //<attribute name='hil_invoicedate' />
            //<attribute name='hil_invoiceavailable' />
            //<attribute name='statuscode' />
            //<attribute name='hil_productcategory' />
            //<attribute name='hil_productsubcategory' />
            //<attribute name='msdyn_product' />
            //<attribute name='hil_modelname' />
            //<attribute name='msdyn_name' />
            //<attribute name='hil_customer' />
            //<order attribute='createdon' descending='true' />
            //<filter type='and'>
            //    <condition attribute='hil_customernm' operator='ne' value='DONE' />
            //    <condition attribute='hil_productcategory' operator='eq' value='{d51edd9d-16fa-e811-a94c-000d3af0694e}' />
            //    <condition attribute='hil_warrantysubstatus' operator='null' />
            //    <condition attribute='createdon' operator='on-or-after' value='2020-07-01' />
            //    <condition attribute='createdon' operator='on-or-before' value='2020-08-01' />
            //    <filter type='or'>
            //    <condition attribute='statuscode' operator='eq' value='910590001' />
            //    <condition attribute='hil_branchheadapprovalstatus' operator='eq' value='1' />
            //    </filter>
            //    <condition attribute='hil_invoiceavailable' operator='eq' value='1' />
            //</filter>
            //</entity>
            //</fetch>";

            //string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            //<entity name='msdyn_customerasset'>
            //<attribute name='hil_invoicedate' />
            //<attribute name='hil_invoiceavailable' />
            //<attribute name='statuscode' />
            //<attribute name='hil_productcategory' />
            //<attribute name='hil_productsubcategory' />
            //<attribute name='msdyn_product' />
            //<attribute name='hil_modelname' />
            //<attribute name='msdyn_name' />
            //<attribute name='hil_customer' />
            //<attribute name='createdon' />
            //<order attribute='createdon' descending='false' />
            //<filter type='and'>
            //<condition attribute='hil_productcategory' operator='in'>
            //<value uiname='LLOYD WASHING MACHIN' uitype='product'>{2FD99DA1-16FA-E811-A94C-000D3AF06091}</value>
            //<value uiname='LLOYD LED TELEVISION' uitype='product'>{A7A5049B-16FA-E811-A94C-000D3AF06091}</value>
            //<value uiname='LLOYD REFRIGERATORS' uitype='product'>{2DD99DA1-16FA-E811-A94C-000D3AF06091}</value>
            //<value uiname='LLOYD AIR CONDITIONER' uitype='product'>{D51EDD9D-16FA-E811-A94C-000D3AF0694E}</value>
            //</condition>
            //<filter type='or'>
            //<condition attribute='statuscode' operator='eq' value='910590001' />
            //<condition attribute='hil_branchheadapprovalstatus' operator='eq' value='1' />
            //</filter>
            //<condition attribute='hil_invoiceavailable' operator='eq' value='1' />
            //<condition attribute='hil_invoicedate' operator='not-null' />
            //<condition attribute='hil_warrantysubstatus' operator='null' />
            //<condition attribute='hil_warrantystatus' operator='ne' value='2' />
            //</filter>
            //</entity>
            //</fetch>";

            #endregion

            //string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            //<entity name='msdyn_customerasset'>
            //<attribute name='hil_invoicedate' />
            //<attribute name='hil_invoiceavailable' />
            //<attribute name='statuscode' />
            //<attribute name='hil_productcategory' />
            //<attribute name='hil_productsubcategory' />
            //<attribute name='msdyn_product' />
            //<attribute name='hil_modelname' />
            //<attribute name='msdyn_name' />
            //<attribute name='hil_customer' />
            //<attribute name='createdon' />
            //<order attribute='createdon' descending='false' />
            //<filter type='and'>
            //    <condition attribute='hil_customernm' operator='ne' value='DONE' />
            //    <condition attribute='createdon' operator='on-or-after' value='2020-08-01' />
            //    <filter type='or'>
            //        <condition attribute='statuscode' operator='eq' value='910590001' />
            //        <condition attribute='hil_branchheadapprovalstatus' operator='eq' value='1' />
            //    </filter>
            //    <condition attribute='hil_invoiceavailable' operator='eq' value='1' />
            //    <condition attribute='hil_invoicedate' operator='not-null' />
            //</filter>
            //<link-entity name='product' from='productid' to='hil_productsubcategory' link-type='inner' alias='av'>
            //    <filter type='and'>
            //        <condition attribute='vendorpartnumber' operator='eq' value='REFRESH' />
            //    </filter>
            //</link-entity>
            //</entity>
            //</fetch>";

            //string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            //<entity name='msdyn_customerasset'>
            //<attribute name='hil_invoicedate' />
            //<attribute name='hil_invoiceavailable' />
            //<attribute name='statuscode' />
            //<attribute name='hil_productcategory' />
            //<attribute name='hil_productsubcategory' />
            //<attribute name='msdyn_product' />
            //<attribute name='hil_modelname' />
            //<attribute name='msdyn_name' />
            //<attribute name='hil_customer' />
            //<attribute name='createdon' />
            //<order attribute='createdon' descending='false' />
            //<filter type='and'>
            //    <filter type='or'>
            //        <condition attribute='statuscode' operator='eq' value='910590001' />
            //        <condition attribute='hil_branchheadapprovalstatus' operator='eq' value='1' />
            //    </filter>
            //    <condition attribute='hil_invoiceavailable' operator='eq' value='1' />
            //    <condition attribute='hil_invoicedate' operator='not-null' />
            //</filter>
            //</entity>
            //</fetch>";

            /*
                <condition attribute='hil_warrantytilldate' operator='on-or-after' value='2022-07-01' />
                <condition attribute='hil_warrantytilldate' operator='on-or-before' value='2022-07-11' />
             */
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
                    <filter type='or'>
                    <condition attribute='statuscode' operator='eq' value='910590001' />
                    <condition attribute='hil_branchheadapprovalstatus' operator='eq' value='1' />
                    </filter>
                    <condition attribute='hil_invoiceavailable' operator='eq' value='1' />
                    <condition attribute='hil_invoicedate' operator='not-null' />
                    <condition attribute='msdyn_name' operator='eq' value='67KDG60400539' />
                </filter>
                <link-entity name='product' from='productid' to='hil_productcategory' link-type='inner' alias='ai'>
                    <filter type='and'>
                        <condition attribute='hil_brandidentifier' operator='eq' value='3' />
                    </filter>
                </link-entity>
                </entity>
                </fetch>";

            int Total = 0;
            int rec = 1;
            //Console.WriteLine(rec.ToString() + "/" + Total.ToString() + " Serial # " + entAMC.GetAttributeValue<string>("hil_serailnumber"));
            while (true)
            {
                EntityCollection ec = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                Total += ec.Entities.Count;
                if (ec.Entities.Count == 0) { break; }
                foreach (Entity CustAsst in ec.Entities)
                {
                    try
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
                        DateTime? stdWrtyEndDate = null;
                        DateTime? spcWrtyEndDate = null;
                        DateTime? extWrtyEndDate = null;

                        if (CustAsst.Attributes.Contains("hil_invoiceavailable") && CustAsst.Attributes.Contains("hil_invoicedate") && CustAsst.Attributes.Contains("hil_productcategory"))
                        {
                            invoiceAvailable = (bool)CustAsst["hil_invoiceavailable"];
                            if (invoiceAvailable)
                            {
                                //QueryExpression queryExp1 = new QueryExpression("hil_unitwarranty");
                                //queryExp1.ColumnSet = new ColumnSet("hil_unitwarrantyid", "hil_customer");
                                //queryExp1.Criteria = new FilterExpression(LogicalOperator.And);
                                //queryExp1.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, CustAsst.Id);
                                //queryExp1.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                                //queryExp1.Criteria.AddCondition("hil_customer", ConditionOperator.Null); //Old Unit Warranties
                                //EntityCollection entCol1 = _service.RetrieveMultiple(queryExp1);
                                //if (entCol1.Entities.Count > 0)
                                //{
                                //    foreach (Entity ent in entCol1.Entities)
                                //    {
                                //        SetStateRequest setStateRequest = new SetStateRequest()
                                //        {
                                //            EntityMoniker = new EntityReference
                                //            {
                                //                Id = ent.Id,
                                //                LogicalName = "hil_unitwarranty",
                                //            },
                                //            State = new OptionSetValue(1), //Inactive
                                //            Status = new OptionSetValue(2) //Inactive
                                //        };
                                //        _service.Execute(setStateRequest);
                                //    }
                                //}

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
                                        if (invDate >= ent.GetAttributeValue<DateTime>("hil_validfrom").AddMinutes(330) && invDate <= ent.GetAttributeValue<DateTime>("hil_validto").AddMinutes(330))
                                        {
                                            endDate = CreateUnitWarrantyLine(_service, ent.GetAttributeValue<int>("hil_warrantyperiod"), 1, erCustomer, CustAsst.ToEntityReference(), erProdCgry, erProdSubCategory, erProdModelCode, pdtModelName, invDate, ent.ToEntityReference());
                                            stdWrtyEndDate = endDate;
                                            warrantySubstatus = 1;
                                            break;
                                        }
                                    }
                                }
                                #endregion

                                #region AMC Warranty Template
                                bool _amcWarranty = false;
                                Entity _warrantyTemplate = null;
                                Entity _customerAsset = null;

                                queryExp = new QueryExpression("hil_amcstaging");
                                queryExp.ColumnSet = new ColumnSet("hil_name", "hil_warrantystartdate", "hil_warrantyenddate", "hil_serailnumber", "hil_sapbillingdocpath", "hil_sapbillingdate", "hil_amcstagingstatus", "hil_amcplan");
                                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                queryExp.Criteria.AddCondition("hil_serailnumber", ConditionOperator.Equal, CustAsst.GetAttributeValue<string>("msdyn_name"));
                                queryExp.Criteria.AddCondition("hil_amcstagingstatus", ConditionOperator.Equal, false); //Draft
                                queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                                queryExp.AddOrder("hil_warrantystartdate", OrderType.Ascending);
                                entCol = _service.RetrieveMultiple(queryExp);

                                if (entCol.Entities.Count > 0)
                                {
                                    _amcWarranty = true;
                                    EntityCollection entCollTemp;

                                    foreach (Entity ent in entCol.Entities)
                                    {
                                        _customerAsset = null;
                                        _warrantyTemplate = null;

                                        queryExp = new QueryExpression("msdyn_customerasset");
                                        queryExp.ColumnSet = new ColumnSet("hil_customer", "msdyn_customerassetid", "msdyn_product", "hil_productcategory", "hil_productsubcategory", "hil_modelname");
                                        queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                        queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_name", ConditionOperator.Equal, ent.GetAttributeValue<string>("hil_serailnumber")));
                                        entCollTemp = _service.RetrieveMultiple(queryExp);
                                        if (entCollTemp.Entities.Count > 0)
                                        {
                                            _customerAsset = entCollTemp.Entities[0];
                                        }
                                        queryExp = new QueryExpression("hil_warrantytemplate");
                                        queryExp.ColumnSet = new ColumnSet("hil_warrantytemplateid");
                                        queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                        queryExp.Criteria.AddCondition(new ConditionExpression("hil_amcplan", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_amcplan").Id));
                                        queryExp.Criteria.AddCondition(new ConditionExpression("hil_templatestatus", ConditionOperator.Equal, 2)); //Approved
                                        queryExp.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1)); //Active
                                        entCollTemp = _service.RetrieveMultiple(queryExp);
                                        if (entCollTemp.Entities.Count > 0)
                                        {
                                            _warrantyTemplate = entCollTemp.Entities[0];
                                        }
                                        if (_customerAsset != null && _warrantyTemplate != null)
                                        {
                                            hil_unitwarranty iSchWarranty = new hil_unitwarranty();
                                            iSchWarranty.hil_CustomerAsset = _customerAsset.ToEntityReference();
                                            if (_customerAsset.Attributes.Contains("hil_productcategory"))
                                            {
                                                iSchWarranty.hil_productmodel = _customerAsset.GetAttributeValue<EntityReference>("hil_productcategory");
                                            }

                                            if (_customerAsset.Attributes.Contains("hil_productsubcategory"))
                                            {
                                                iSchWarranty.hil_productitem = _customerAsset.GetAttributeValue<EntityReference>("hil_productsubcategory");
                                            }

                                            iSchWarranty.hil_warrantystartdate = ent.GetAttributeValue<DateTime>("hil_warrantystartdate").AddMinutes(330);
                                            iSchWarranty.hil_warrantyenddate = ent.GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330);

                                            iSchWarranty.hil_WarrantyTemplate = _warrantyTemplate.ToEntityReference();

                                            iSchWarranty.hil_ProductType = new OptionSetValue(1);

                                            if (_customerAsset.Attributes.Contains("msdyn_product"))
                                            {
                                                iSchWarranty.hil_Part = _customerAsset.GetAttributeValue<EntityReference>("msdyn_product");
                                            }

                                            if (_customerAsset.Attributes.Contains("hil_modelname"))
                                            {
                                                iSchWarranty["hil_partdescription"] = _customerAsset.GetAttributeValue<string>("hil_modelname");
                                            }

                                            if (_customerAsset.Attributes.Contains("hil_customer"))
                                            {
                                                iSchWarranty.hil_customer = _customerAsset.GetAttributeValue<EntityReference>("hil_customer");
                                            }
                                            iSchWarranty["hil_amcbillingdocdate"] = ent.GetAttributeValue<DateTime>("hil_sapbillingdate").AddMinutes(330);
                                            iSchWarranty["hil_amcbillingdocnum"] = ent.GetAttributeValue<string>("hil_name");
                                            iSchWarranty["hil_amcbillingdocurl"] = ent.GetAttributeValue<string>("hil_sapbillingdocpath");
                                            _service.Create(iSchWarranty);
                                        }
                                    }
                                }

                                #endregion

                                #region Special Scheme Warranty Template
                                bool _specialSchemeApplied = false;
                                if (stdWrtyEndDate != null && !_amcWarranty)
                                {
                                    queryExp = new QueryExpression("hil_warrantytemplate");
                                    queryExp.ColumnSet = new ColumnSet("hil_validfrom", "hil_validto", "hil_warrantyperiod", "hil_category");
                                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                    queryExp.Criteria.AddCondition("hil_product", ConditionOperator.Equal, erProdSubCategory.Id);
                                    queryExp.Criteria.AddCondition("hil_validfrom", ConditionOperator.NotNull);
                                    queryExp.Criteria.AddCondition("hil_validto", ConditionOperator.NotNull);
                                    queryExp.Criteria.AddCondition("hil_type", ConditionOperator.Equal, 7); //Scheme Warranty
                                    queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                                    queryExp.Criteria.AddCondition("hil_templatestatus", ConditionOperator.Equal, 2); //Approved
                                    entCol = _service.RetrieveMultiple(queryExp);
                                    foreach (Entity ent in entCol.Entities)
                                    {
                                        if (invDate >= ent.GetAttributeValue<DateTime>("hil_validfrom").AddMinutes(330) && invDate <= ent.GetAttributeValue<DateTime>("hil_validto").AddMinutes(330))
                                        {
                                            endDate = CreateSchemeUnitWarrantyLine(_service, ent.GetAttributeValue<int>("hil_warrantyperiod"), 7, erCustomer, CustAsst.ToEntityReference(), erProdCgry, erProdSubCategory, erProdModelCode, pdtModelName, stdWrtyEndDate.Value.AddDays(1), ent.ToEntityReference(), Convert.ToDateTime(invDate));
                                            if (endDate == stdWrtyEndDate.Value.AddDays(1))
                                            {
                                                endDate = stdWrtyEndDate;
                                            }
                                            else
                                            {
                                                spcWrtyEndDate = endDate;
                                                warrantySubstatus = 3;
                                                _specialSchemeApplied = true;
                                            }
                                            break;
                                        }
                                    }
                                }
                                #endregion

                                #region Extended Warranty Template
                                if (endDate != null && !_specialSchemeApplied && !_amcWarranty)
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
                                        if (invDate >= ent.GetAttributeValue<DateTime>("hil_validfrom").AddMinutes(330) && invDate <= ent.GetAttributeValue<DateTime>("hil_validto").AddMinutes(330))
                                        {
                                            endDate = CreateUnitWarrantyLine(_service, ent.GetAttributeValue<int>("hil_warrantyperiod"), 2, erCustomer, CustAsst.ToEntityReference(), erProdCgry, erProdSubCategory, erProdModelCode, pdtModelName, endDate.Value.AddDays(1), ent.ToEntityReference());
                                            warrantySubstatus = 2;
                                            break;
                                        }
                                    }
                                }
                                #endregion

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
                                    if (entCol.Entities[0].Attributes.Contains("hil_warrantyenddate"))
                                    {
                                        endDate = entCol.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330);
                                    }
                                    else
                                    {
                                        endDate = new DateTime(1900, 1, 1);
                                    }
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
                                                if (DateTime.Now >= ent.GetAttributeValue<DateTime>("hil_warrantystartdate").AddMinutes(330) && DateTime.Now <= ent.GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330))
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
                                                    else if (entSubStatus.GetAttributeValue<OptionSetValue>("hil_type").Value == 3)
                                                    {
                                                        entCustAsset["hil_warrantysubstatus"] = new OptionSetValue(4); //InWarranty-AMC
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
                                    entCustAsset["hil_warrantytilldate"] = new DateTime(1900, 1, 1);
                                    entCustAsset["hil_warrantystatus"] = new OptionSetValue(2); //OutWarranty
                                    entCustAsset["hil_warrantysubstatus"] = null;
                                }
                                //hil_customernm
                                //entCustAsset["hil_customernm"] = "DONE";
                                _service.Update(entCustAsset);
                            }
                            else
                            {
                                Entity entCustAsset = new Entity(msdyn_customerasset.EntityLogicalName);
                                entCustAsset.Id = CustAsst.Id;
                                entCustAsset["hil_warrantytilldate"] = new DateTime(1900, 1, 1);
                                entCustAsset["hil_warrantystatus"] = new OptionSetValue(2); //OutWarranty
                                entCustAsset["hil_warrantysubstatus"] = null;
                                _service.Update(entCustAsset);
                            }
                        }
                        else
                        {
                            Entity entCustAsset = new Entity(msdyn_customerasset.EntityLogicalName);
                            entCustAsset.Id = CustAsst.Id;
                            entCustAsset["hil_warrantytilldate"] = new DateTime(1900, 1, 1);
                            entCustAsset["hil_warrantystatus"] = new OptionSetValue(2); //OutWarranty
                            entCustAsset["hil_warrantysubstatus"] = null;
                            _service.Update(entCustAsset);
                        }
                    }
                    catch (Exception ex)
                    {
                       // Console.WriteLine(CustAsst.GetAttributeValue<string>("msdyn_name").ToString() + " / " + ex.Message + " / " + rec.ToString() + "/" + Total.ToString());
                    }
                    Console.WriteLine(CustAsst.GetAttributeValue<DateTime>("createdon").ToString() + " / " + rec++.ToString() + "/" + Total.ToString());
                }
            }
        }

        static void RefreshClaimLinesOnceAMCInvoiceImported()
        {
            int Total = 0;
            string _fetchXMLAMC = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_amcstaging'>
                <attribute name='hil_amcstagingid' />
                <attribute name='createdon' />
                <attribute name='hil_serailnumber' />
                <order attribute='createdon' descending='false' />
                <filter type='and'>
                    <condition attribute='createdon' operator='today' />
                    <condition attribute='hil_description' operator='like' value='%DONE%' />
                </filter>
                </entity>
                </fetch>";
            int i = 1;
            while (true)
            {
                EntityCollection entAMCBill = _service.RetrieveMultiple(new FetchExpression(_fetchXMLAMC));
                if (entAMCBill.Entities.Count == 0) { break; }
                foreach (Entity entAMC in entAMCBill.Entities)
                {
                    string _serialNum = entAMC.GetAttributeValue<string>("hil_serailnumber");
                    Console.WriteLine("Processing... (" + _serialNum + ")" + i++.ToString());

                    string fetchXML = $@"<fetch version='1.0' top='1000' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='msdyn_customerasset'>
                    <attribute name='hil_invoicedate' />
                    <attribute name='hil_invoiceavailable' />
                    <attribute name='statuscode' />
                    <attribute name='hil_branchheadapprovalstatus' />
                    <attribute name='hil_productcategory' />
                    <attribute name='hil_productsubcategory' />
                    <attribute name='msdyn_product' />
                    <attribute name='hil_modelname' />
                    <attribute name='msdyn_name' />
                    <attribute name='hil_customer' />
                    <attribute name='createdon' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='msdyn_name' operator='eq' value='{_serialNum}' />
                    </filter>
                    </entity>
                    </fetch>";
                    int rec = 1;
                    EntityCollection ec = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    Total += ec.Entities.Count;
                    if (ec.Entities.Count == 0) { continue; }
                    foreach (Entity CustAsst in ec.Entities)
                    {
                        Console.WriteLine("Processing... " + rec.ToString() + "/" + Total.ToString() + " Serial # " + CustAsst.GetAttributeValue<string>("msdyn_name"));
                        string _fetchXMLJob = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorder'>
                        <attribute name='msdyn_name' />
                        <attribute name='hil_generateclaim' />
                        <order attribute='msdyn_name' descending='false' />
                        <filter type='and'>
                            <condition attribute='hil_fiscalmonth' operator='eq' value='{{56DE3041-22CE-EE11-904C-000D3A3E3D4E}}' />
                            <condition attribute='msdyn_customerasset' operator='eq' value='{CustAsst.Id}' />
                        </filter>
                        </entity>
                        </fetch>";
                        EntityCollection ec1 = _service.RetrieveMultiple(new FetchExpression(_fetchXMLJob));
                        foreach (Entity entJob in ec1.Entities)
                        {
                            string _fetchXMLClaim = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='hil_claimline'>
                                <attribute name='hil_claimlineid' />
                                <attribute name='statuscode' />
                                <attribute name='statecode' />
                                <order attribute='hil_name' descending='false' />
                                <filter type='and'>
                                  <condition attribute='hil_claimperiod' operator='eq' value='{{56DE3041-22CE-EE11-904C-000D3A3E3D4E}}' />
                                  <condition attribute='hil_jobid' operator='eq' value='{entJob.Id}' />
                                  <condition attribute='statecode' operator='eq' value='0' />
                                </filter>
                              </entity>
                            </fetch>";
                            EntityCollection ec2 = _service.RetrieveMultiple(new FetchExpression(_fetchXMLClaim));
                            foreach (Entity entClaim in ec2.Entities) {
                                Entity _entClaimUpdate = new Entity(entClaim.LogicalName, entClaim.Id);
                                _entClaimUpdate["statecode"] = new OptionSetValue(1);
                                _entClaimUpdate["statuscode"] = new OptionSetValue(2);
                                try
                                {
                                    _service.Update(_entClaimUpdate);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("ERROR!!! " + ex.Message);
                                }
                            }
                            Entity _entJobUpdate = new Entity(entJob.LogicalName, entJob.Id);
                            _entJobUpdate["hil_generateclaim"] = false;
                            _service.Update(_entJobUpdate);
                        }
                        rec += 1;
                    }
                    //entAMC["hil_description"] = "DONE";
                    //_service.Update(entAMC);
                    Console.WriteLine("DONE:: " + _serialNum);
                }
            }
        }
        static void RefreshHavellsWaterPurifierUnitWarranty(string _assetSerialNum)
        {
            int Total = 0;
            //Manual Upload -LLOYD
            //72981D83-16FA-E811-A94C-000D3AF0694E - WP
            //D51EDD9D-16FA-E811-A94C-000D3AF0694E - AC
            //<condition attribute='hil_serailnumber' operator='eq' value='MIJGG41R01285' />
            string _fetchXMLAMC = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_amcstaging'>
                <attribute name='hil_serailnumber' />
                <attribute name='hil_description' />
                <filter type='and'>
                    <condition attribute='hil_description' operator='like' value='%Manual Upload -LLOYD%' />
                    <condition attribute='hil_amcstagingstatus' operator='eq' value='0' />
                </filter>
                </entity>
                </fetch>";
            int i = 1;
            while (true)
            {
                EntityCollection entAMCBill = _service.RetrieveMultiple(new FetchExpression(_fetchXMLAMC));
                if (entAMCBill.Entities.Count == 0) { break; }
                foreach (Entity entAMC in entAMCBill.Entities)
                {
                    string _serialNum = entAMC.GetAttributeValue<string>("hil_serailnumber");
                    Console.WriteLine("Processing... (" + _serialNum  + ")"+ i++.ToString());
                    //<condition attribute='hil_productcategory' operator='eq' value='{D51EDD9D-16FA-E811-A94C-000D3AF0694E}' />

                    string fetchXML = $@"<fetch version='1.0' top='1000' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='msdyn_customerasset'>
                    <attribute name='hil_invoicedate' />
                    <attribute name='hil_invoiceavailable' />
                    <attribute name='statuscode' />
                    <attribute name='hil_branchheadapprovalstatus' />
                    <attribute name='hil_productcategory' />
                    <attribute name='hil_productsubcategory' />
                    <attribute name='msdyn_product' />
                    <attribute name='hil_modelname' />
                    <attribute name='msdyn_name' />
                    <attribute name='hil_customer' />
                    <attribute name='createdon' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='msdyn_name' operator='eq' value='{_serialNum}' />";
                        //<condition attribute='hil_uniquekey' operator='ne' value='..' />";
                    if (!string.IsNullOrEmpty(_assetSerialNum))
                        fetchXML = fetchXML + @"<condition attribute='msdyn_name' operator='eq' value='" + _assetSerialNum + @"' />";

                    fetchXML = fetchXML + @"</filter>
                    </entity>
                    </fetch>";
                    int rec = 1;
                    //while (true)
                    //{
                        EntityCollection ec = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                        Total += ec.Entities.Count;
                        if (ec.Entities.Count == 0) { continue; }
                        foreach (Entity CustAsst in ec.Entities)
                        {
                            Console.WriteLine("Processing... " + rec.ToString() + "/" + Total.ToString() + " Serial # " + CustAsst.GetAttributeValue<string>("msdyn_name"));
                            if (!CustAsst.Attributes.Contains("msdyn_name"))
                            {
                                Entity entCustAssetTemp = new Entity(CustAsst.LogicalName, CustAsst.Id);
                                entCustAssetTemp["hil_uniquekey"] = ".";
                                entCustAssetTemp["hil_warrantytilldate"] = new DateTime(1900, 1, 1);
                                entCustAssetTemp["hil_warrantystatus"] = new OptionSetValue(2); //OutWarranty
                                entCustAssetTemp["hil_warrantysubstatus"] = null;
                                _service.Update(entCustAssetTemp);
                                continue;
                            }
                            try
                            {
                                EntityReference erProdCgry = new EntityReference("product");
                                EntityReference erProdSubCategory = new EntityReference("product");
                                EntityReference erProdModelCode = new EntityReference("product");
                                EntityReference erCustomer = new EntityReference("contact");
                                string pdtModelName = string.Empty;
                                DateTime? invDate = null;
                                DateTime? endDate = null;
                                DateTime? stdWrtyEndDate = null;
                                bool isApproved = false;
                                QueryExpression queryExp = null;
                                EntityCollection entCol = null;

                                if (CustAsst.Attributes.Contains("hil_invoicedate"))
                                {
                                    invDate = ((DateTime)CustAsst["hil_invoicedate"]).AddMinutes(330);
                                }
                                if (CustAsst.Attributes.Contains("statuscode"))
                                {
                                    isApproved = ((OptionSetValue)CustAsst["statuscode"]).Value == 910590001;
                                }
                                if (CustAsst.Attributes.Contains("hil_branchheadapprovalstatus"))
                                {
                                    if (!isApproved)
                                    {
                                        isApproved = ((OptionSetValue)CustAsst["hil_branchheadapprovalstatus"]).Value == 1;
                                    }
                                }

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

                                if (isApproved && invDate != null)
                                {
                                    #region Standard Warranty Template
                                    queryExp = new QueryExpression("hil_warrantytemplate");
                                    queryExp.ColumnSet = new ColumnSet("hil_validfrom", "hil_validto", "hil_warrantyperiod", "hil_category");
                                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                    queryExp.Criteria.AddCondition("hil_product", ConditionOperator.Equal, erProdSubCategory.Id);
                                    queryExp.Criteria.AddCondition("hil_validfrom", ConditionOperator.NotNull);
                                    queryExp.Criteria.AddCondition("hil_validto", ConditionOperator.NotNull);
                                    queryExp.Criteria.AddCondition("hil_type", ConditionOperator.Equal, 1); //Standard Warranty
                                    queryExp.Criteria.AddCondition("hil_templatestatus", ConditionOperator.Equal, 2); //Approved
                                    queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                                    entCol = _service.RetrieveMultiple(queryExp);
                                    if (entCol.Entities.Count > 0)
                                    {
                                        foreach (Entity ent in entCol.Entities)
                                        {
                                            if (invDate >= ent.GetAttributeValue<DateTime>("hil_validfrom").AddMinutes(330) && invDate <= ent.GetAttributeValue<DateTime>("hil_validto").AddMinutes(330))
                                            {
                                                endDate = CreateUnitWarrantyLine(_service, ent.GetAttributeValue<int>("hil_warrantyperiod"), 1, erCustomer, CustAsst.ToEntityReference(), erProdCgry, erProdSubCategory, erProdModelCode, pdtModelName, invDate, ent.ToEntityReference());
                                                stdWrtyEndDate = endDate;
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
                                        queryExp.Criteria.AddCondition("hil_type", ConditionOperator.Equal, 7); //Scheme Warranty
                                        queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                                        queryExp.Criteria.AddCondition("hil_templatestatus", ConditionOperator.Equal, 2); //Approved
                                        entCol = _service.RetrieveMultiple(queryExp);
                                        foreach (Entity ent in entCol.Entities)
                                        {
                                            if (invDate >= ent.GetAttributeValue<DateTime>("hil_validfrom") && invDate <= ent.GetAttributeValue<DateTime>("hil_validto"))
                                            {
                                                endDate = CreateSchemeUnitWarrantyLine(_service, ent.GetAttributeValue<int>("hil_warrantyperiod"), 7, erCustomer, CustAsst.ToEntityReference(), erProdCgry, erProdSubCategory, erProdModelCode, pdtModelName, stdWrtyEndDate.Value.AddDays(1), ent.ToEntityReference(), Convert.ToDateTime(invDate));
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
                                                CreateUnitWarrantyLine(_service, ent.GetAttributeValue<int>("hil_warrantyperiod"), 2, erCustomer, CustAsst.ToEntityReference(), erProdCgry, erProdSubCategory, erProdModelCode, pdtModelName, endDate.Value.AddDays(1), ent.ToEntityReference());
                                                break;
                                            }
                                        }
                                    }
                                    #endregion
                                }

                                #region Inactivating Existing AMC Warranties
                                string _fetchXMLUWL = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='hil_unitwarranty'>
                                    <attribute name='hil_unitwarrantyid' />
                                    <order attribute='hil_name' descending='false' />
                                    <filter type='and'>
                                        <condition attribute='hil_customerasset' operator='eq' value='{CustAsst.Id}' />
                                        <condition attribute='statecode' operator='eq' value='0' />
                                    </filter>
                                    <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='ad'>
                                        <filter type='and'>
                                        <condition attribute='hil_type' operator='eq' value='3' />
                                        </filter>
                                    </link-entity>
                                    </entity>
                                    </fetch>";

                                EntityCollection entCol1 = _service.RetrieveMultiple(new FetchExpression(_fetchXMLUWL));
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
                                #endregion

                                #region Processing AMC SAP Invoices
                                DateTime _amcStartDate = DateTime.Now;
                                queryExp = new QueryExpression("hil_amcstaging");
                                queryExp.ColumnSet = new ColumnSet("hil_name", "hil_warrantystartdate", "hil_warrantyenddate", "hil_serailnumber", "hil_sapbillingdocpath", "hil_sapbillingdate", "hil_amcstagingstatus", "hil_amcplan");
                                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                queryExp.Criteria.AddCondition("hil_serailnumber", ConditionOperator.Equal, CustAsst.GetAttributeValue<string>("msdyn_name"));
                                queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                                queryExp.AddOrder("hil_warrantystartdate", OrderType.Ascending);

                                EntityCollection entColAMC = _service.RetrieveMultiple(queryExp);

                                if (entColAMC.Entities.Count > 0)
                                {
                                    EntityCollection entCollTemp;
                                    foreach (Entity ent in entColAMC.Entities)
                                    {
                                        DateTime _sapInvoiceDate = ent.GetAttributeValue<DateTime>("hil_sapbillingdate").AddMinutes(330);
                                        int _compare = DateTime.Compare(Convert.ToDateTime(endDate), _sapInvoiceDate);
                                        if (_compare >= 0) //Warranrty Enddate is later than AMC Invloice Date
                                        {
                                            _amcStartDate = Convert.ToDateTime(endDate).AddDays(1);
                                        }
                                        else
                                        {
                                            _amcStartDate = _sapInvoiceDate;
                                        }
                                        Entity _warrantyTemplate = null;
                                        int _period = 0;
                                        queryExp = new QueryExpression("hil_warrantytemplate");
                                        queryExp.ColumnSet = new ColumnSet("hil_warrantytemplateid", "hil_warrantyperiod");
                                        queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                        queryExp.Criteria.AddCondition(new ConditionExpression("hil_amcplan", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_amcplan").Id));
                                        queryExp.Criteria.AddCondition(new ConditionExpression("hil_templatestatus", ConditionOperator.Equal, 2)); //Approved
                                        queryExp.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1)); //Active
                                        entCollTemp = _service.RetrieveMultiple(queryExp);
                                        if (entCollTemp.Entities.Count > 0)
                                        {
                                            _warrantyTemplate = entCollTemp.Entities[0];
                                            _period = entCollTemp.Entities[0].GetAttributeValue<int>("hil_warrantyperiod");
                                        }
                                        if (_warrantyTemplate != null)
                                        {
                                            hil_unitwarranty iSchWarranty = new hil_unitwarranty();
                                            iSchWarranty.hil_CustomerAsset = CustAsst.ToEntityReference();
                                            iSchWarranty.hil_productmodel = erProdCgry;
                                            iSchWarranty.hil_productitem = erProdSubCategory;

                                            endDate = _amcStartDate.AddMonths(_period).AddDays(-1);
                                            iSchWarranty.hil_warrantystartdate = _amcStartDate;//ent.GetAttributeValue<DateTime>("hil_warrantystartdate").AddMinutes(330);
                                            iSchWarranty.hil_warrantyenddate = endDate; //ent.GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330);

                                            iSchWarranty.hil_WarrantyTemplate = _warrantyTemplate.ToEntityReference();

                                            iSchWarranty.hil_ProductType = new OptionSetValue(1);

                                            iSchWarranty.hil_Part = erProdModelCode;

                                            iSchWarranty["hil_partdescription"] = pdtModelName;

                                            iSchWarranty.hil_customer = erCustomer;
                                            iSchWarranty["hil_amcbillingdocdate"] = ent.GetAttributeValue<DateTime>("hil_sapbillingdate").AddMinutes(330);
                                            iSchWarranty["hil_amcbillingdocnum"] = ent.GetAttributeValue<string>("hil_name");
                                            iSchWarranty["hil_amcbillingdocurl"] = ent.GetAttributeValue<string>("hil_sapbillingdocpath");
                                            _service.Create(iSchWarranty);
                                        }
                                    }
                                }
                            #endregion

                            #region Updating Customer Asset Warranty Status
                            string _currentDate = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString().PadLeft(2, '0') + "-" + DateTime.Now.Day.ToString().PadLeft(2, '0');
                                string _fetchXMLW = $@"<fetch top='1' version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='hil_unitwarranty'>
                                    <attribute name='hil_name' />
                                    <attribute name='hil_warrantyenddate' />
                                    <attribute name='hil_unitwarrantyid' />
                                    <order attribute='hil_warrantyenddate' descending='true' />
                                    <filter type='and'>
                                        <condition attribute='statecode' operator='eq' value='0' />
                                        <condition attribute='hil_customerasset' operator='eq' value='{CustAsst.Id}' />
                                        <condition attribute='hil_warrantystartdate' operator='on-or-before' value='{_currentDate}' />
                                        <condition attribute='hil_warrantyenddate' operator='on-or-after' value='{_currentDate}' />
                                    </filter>
                                    <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='wt'>
                                      <attribute name='hil_warrantyperiod' />
                                      <attribute name='hil_type' />
                                      <filter type='and'>
                                        <condition attribute='hil_type' operator='ne' value='2' />
                                      </filter>
                                    </link-entity>
                                  </entity>
                                </fetch>";

                                entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXMLW));

                                Entity entCustAsset = new Entity(msdyn_customerasset.EntityLogicalName);
                                entCustAsset.Id = CustAsst.Id;
                                int _warrantyType = 0;
                                if (entCol.Entities.Count > 0)
                                {
                                    if (entCol.Entities[0].Attributes.Contains("hil_warrantyenddate"))
                                    {
                                        endDate = entCol.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330);
                                        _warrantyType = ((OptionSetValue)entCol.Entities[0].GetAttributeValue<AliasedValue>("wt.hil_type").Value).Value;
                                    }
                                    else
                                    {
                                        endDate = new DateTime(1900, 1, 1);
                                    }
                                    entCustAsset["hil_warrantytilldate"] = endDate;
                                    if (endDate >= DateTime.Now)
                                    {
                                        entCustAsset["hil_warrantystatus"] = new OptionSetValue(1); //InWarranty

                                        if (_warrantyType == 1)
                                        {
                                            entCustAsset["hil_warrantysubstatus"] = new OptionSetValue(1); //InWarranty-Standard
                                        }
                                        else if (_warrantyType == 2)
                                        {
                                            entCustAsset["hil_warrantysubstatus"] = new OptionSetValue(2); //InWarranty-Extended
                                        }
                                        else if (_warrantyType == 7)
                                        {
                                            entCustAsset["hil_warrantysubstatus"] = new OptionSetValue(3); //InWarranty-Special Scheme
                                        }
                                        else if (_warrantyType == 3)
                                        {
                                            entCustAsset["hil_warrantysubstatus"] = new OptionSetValue(4); //InWarranty-AMC
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
                                    entCustAsset["hil_warrantytilldate"] = new DateTime(1900, 1, 1);
                                    entCustAsset["hil_warrantystatus"] = new OptionSetValue(2); //OutWarranty
                                    entCustAsset["hil_warrantysubstatus"] = null;
                                }
                                //Calculate Extended Warranty End Date
                                _fetchXMLW = $@"<fetch top='1' version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='hil_unitwarranty'>
                                    <attribute name='hil_name' />
                                    <attribute name='hil_warrantyenddate' />
                                    <attribute name='hil_unitwarrantyid' />
                                    <order attribute='hil_warrantyenddate' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='statecode' operator='eq' value='0' />
                                      <condition attribute='hil_customerasset' operator='eq' value='{CustAsst.Id}' />
                                    </filter>
                                    <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='wt'>
                                      <attribute name='hil_warrantyperiod' />
                                      <attribute name='hil_type' />
                                      <filter type='and'>
                                        <condition attribute='hil_type' operator='eq' value='2' />
                                      </filter>
                                    </link-entity>
                                  </entity>
                                </fetch>";
                                entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXMLW));
                                if (entCol.Entities.Count > 0)
                                {
                                    entCustAsset["hil_extendedwarrantyenddate"] = entCol.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330);
                                }
                                entCustAsset["hil_uniquekey"] = ".";
                                _service.Update(entCustAsset);
                                #endregion
                            }
                            catch (Exception ex)
                            {
                                Entity ent1 = new Entity(CustAsst.LogicalName, CustAsst.Id);
                                ent1["hil_uniquekey"] = ".";
                                ent1["hil_warrantytilldate"] = new DateTime(1900, 1, 1);
                                ent1["hil_warrantystatus"] = new OptionSetValue(2); //OutWarranty
                                ent1["hil_warrantysubstatus"] = null;
                                _service.Update(ent1);
                                Console.WriteLine(CustAsst.GetAttributeValue<string>("msdyn_name").ToString() + " / " + ex.Message + " / " + rec.ToString() + "/" + Total.ToString());
                            }
                            Console.WriteLine(CustAsst.GetAttributeValue<DateTime>("createdon").ToString() + " / " + rec.ToString() + "/" + Total.ToString());
                            rec += 1;
                        }
                    //}
                    entAMC["hil_description"] = "DONE";
                    _service.Update(entAMC);
                    Console.WriteLine("DONE:: " + _serialNum);
                }
            }
        }

        static void RefreshAssetUnitWarrantyAMC()
        {
            int Total = 0;
            while (true)
            {
                QueryExpression queryExpTemp = new QueryExpression("hil_amcstaging");
                queryExpTemp.ColumnSet = new ColumnSet("hil_serailnumber", "hil_amcstagingstatus", "hil_sapbillingdocpath");
                queryExpTemp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExpTemp.Criteria.AddCondition("hil_amcstagingstatus", ConditionOperator.Equal, false); //Draft
                queryExpTemp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                queryExpTemp.Criteria.AddCondition("hil_serailnumber", ConditionOperator.Equal, "67JHG61R00986");
                EntityCollection entColAMC = _service.RetrieveMultiple(queryExpTemp);

                int rec = 1;
                Total += entColAMC.Entities.Count;
                if (entColAMC.Entities.Count == 0) { break; }
                foreach (Entity entAMC in entColAMC.Entities)
                {
                    Console.WriteLine(rec.ToString() + "/" + Total.ToString() + " Serial # " + entAMC.GetAttributeValue<string>("hil_serailnumber"));
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

                    EntityCollection ec = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    foreach (Entity CustAsst in ec.Entities)
                    {
                        try
                        {
                            EntityReference erProdCgry = new EntityReference("product");
                            EntityReference erProdSubCategory = new EntityReference("product");
                            EntityReference erProdModelCode = new EntityReference("product");
                            EntityReference erCustomer = new EntityReference("contact");
                            string pdtModelName = string.Empty;
                            DateTime? invDate = null;
                            DateTime? endDate = null;
                            bool invoiceAvailable = false;
                            DateTime? stdWrtyEndDate = null;

                            if (CustAsst.Attributes.Contains("hil_invoiceavailable") && CustAsst.Attributes.Contains("hil_invoicedate") && CustAsst.Attributes.Contains("hil_productcategory"))
                            {
                                invoiceAvailable = (bool)CustAsst["hil_invoiceavailable"];

                                if (invoiceAvailable)
                                {
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
                                            if (invDate >= ent.GetAttributeValue<DateTime>("hil_validfrom").AddMinutes(330) && invDate <= ent.GetAttributeValue<DateTime>("hil_validto").AddMinutes(330))
                                            {
                                                endDate = CreateUnitWarrantyLine(_service, ent.GetAttributeValue<int>("hil_warrantyperiod"), 1, erCustomer, CustAsst.ToEntityReference(), erProdCgry, erProdSubCategory, erProdModelCode, pdtModelName, invDate, ent.ToEntityReference());
                                                stdWrtyEndDate = endDate;
                                                break;
                                            }
                                        }
                                    }
                                    #endregion

                                    #region AMC Warranty Template
                                    bool _amcWarranty = false;
                                    Entity _warrantyTemplate = null;
                                    Entity _customerAsset = null;

                                    queryExp = new QueryExpression("hil_amcstaging");
                                    queryExp.ColumnSet = new ColumnSet("hil_name", "hil_warrantystartdate", "hil_warrantyenddate", "hil_serailnumber", "hil_sapbillingdocpath", "hil_sapbillingdate", "hil_amcstagingstatus", "hil_amcplan");
                                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                    queryExp.Criteria.AddCondition("hil_serailnumber", ConditionOperator.Equal, CustAsst.GetAttributeValue<string>("msdyn_name"));
                                    queryExp.Criteria.AddCondition("hil_amcstagingstatus", ConditionOperator.Equal, false); //Draft
                                    queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                    queryExp.AddOrder("hil_warrantystartdate", OrderType.Ascending);
                                    entCol = _service.RetrieveMultiple(queryExp);

                                    if (entCol.Entities.Count > 0)
                                    {
                                        _amcWarranty = true;
                                        EntityCollection entCollTemp;

                                        foreach (Entity ent in entCol.Entities)
                                        {
                                            QueryExpression queryExpUWL = new QueryExpression("hil_unitwarranty");
                                            queryExpUWL.ColumnSet = new ColumnSet(false);
                                            queryExpUWL.Criteria = new FilterExpression(LogicalOperator.And);
                                            queryExpUWL.Criteria.AddCondition("hil_amcbillingdocnum", ConditionOperator.Equal, ent.GetAttributeValue<string>("hil_name"));
                                            queryExpUWL.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                                            EntityCollection entColUWL = _service.RetrieveMultiple(queryExpUWL);
                                            if (entColUWL.Entities.Count > 0)
                                            {
                                                Console.WriteLine(CustAsst.GetAttributeValue<string>("msdyn_name").ToString() + " Duplicate SAP Invoice #" + ent.GetAttributeValue<string>("hil_name"));
                                                break;
                                            }

                                            _customerAsset = null;
                                            _warrantyTemplate = null;

                                            queryExp = new QueryExpression("msdyn_customerasset");
                                            queryExp.ColumnSet = new ColumnSet("hil_customer", "msdyn_customerassetid", "msdyn_product", "hil_productcategory", "hil_productsubcategory", "hil_modelname");
                                            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                            queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_name", ConditionOperator.Equal, ent.GetAttributeValue<string>("hil_serailnumber")));
                                            entCollTemp = _service.RetrieveMultiple(queryExp);
                                            if (entCollTemp.Entities.Count > 0)
                                            {
                                                _customerAsset = entCollTemp.Entities[0];
                                            }
                                            queryExp = new QueryExpression("hil_warrantytemplate");
                                            queryExp.ColumnSet = new ColumnSet("hil_warrantytemplateid");
                                            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                            queryExp.Criteria.AddCondition(new ConditionExpression("hil_amcplan", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_amcplan").Id));
                                            queryExp.Criteria.AddCondition(new ConditionExpression("hil_templatestatus", ConditionOperator.Equal, 2)); //Approved
                                            queryExp.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1)); //Active
                                            entCollTemp = _service.RetrieveMultiple(queryExp);
                                            if (entCollTemp.Entities.Count > 0)
                                            {
                                                _warrantyTemplate = entCollTemp.Entities[0];
                                            }
                                            if (_customerAsset != null && _warrantyTemplate != null)
                                            {
                                                hil_unitwarranty iSchWarranty = new hil_unitwarranty();
                                                iSchWarranty.hil_CustomerAsset = _customerAsset.ToEntityReference();
                                                if (_customerAsset.Attributes.Contains("hil_productcategory"))
                                                {
                                                    iSchWarranty.hil_productmodel = _customerAsset.GetAttributeValue<EntityReference>("hil_productcategory");
                                                }

                                                if (_customerAsset.Attributes.Contains("hil_productsubcategory"))
                                                {
                                                    iSchWarranty.hil_productitem = _customerAsset.GetAttributeValue<EntityReference>("hil_productsubcategory");
                                                }

                                                iSchWarranty.hil_warrantystartdate = ent.GetAttributeValue<DateTime>("hil_warrantystartdate").AddMinutes(330);
                                                iSchWarranty.hil_warrantyenddate = ent.GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330);

                                                iSchWarranty.hil_WarrantyTemplate = _warrantyTemplate.ToEntityReference();

                                                iSchWarranty.hil_ProductType = new OptionSetValue(1);

                                                if (_customerAsset.Attributes.Contains("msdyn_product"))
                                                {
                                                    iSchWarranty.hil_Part = _customerAsset.GetAttributeValue<EntityReference>("msdyn_product");
                                                }

                                                if (_customerAsset.Attributes.Contains("hil_modelname"))
                                                {
                                                    iSchWarranty["hil_partdescription"] = _customerAsset.GetAttributeValue<string>("hil_modelname");
                                                }

                                                if (_customerAsset.Attributes.Contains("hil_customer"))
                                                {
                                                    iSchWarranty.hil_customer = _customerAsset.GetAttributeValue<EntityReference>("hil_customer");
                                                }
                                                iSchWarranty["hil_amcbillingdocdate"] = ent.GetAttributeValue<DateTime>("hil_sapbillingdate").AddMinutes(330);
                                                iSchWarranty["hil_amcbillingdocnum"] = ent.GetAttributeValue<string>("hil_name");
                                                iSchWarranty["hil_amcbillingdocurl"] = ent.GetAttributeValue<string>("hil_sapbillingdocpath");
                                                _service.Create(iSchWarranty);

                                            }
                                        }
                                    }

                                    #endregion

                                    #region Special Scheme Warranty Template
                                    //bool _specialSchemeApplied = false;
                                    //if (stdWrtyEndDate != null && !_amcWarranty)
                                    //{
                                    //    queryExp = new QueryExpression("hil_warrantytemplate");
                                    //    queryExp.ColumnSet = new ColumnSet("hil_validfrom", "hil_validto", "hil_warrantyperiod", "hil_category");
                                    //    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                    //    queryExp.Criteria.AddCondition("hil_product", ConditionOperator.Equal, erProdSubCategory.Id);
                                    //    queryExp.Criteria.AddCondition("hil_validfrom", ConditionOperator.NotNull);
                                    //    queryExp.Criteria.AddCondition("hil_validto", ConditionOperator.NotNull);
                                    //    queryExp.Criteria.AddCondition("hil_type", ConditionOperator.Equal, 7); //Scheme Warranty
                                    //    queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                                    //    queryExp.Criteria.AddCondition("hil_templatestatus", ConditionOperator.Equal, 2); //Approved
                                    //    entCol = _service.RetrieveMultiple(queryExp);
                                    //    foreach (Entity ent in entCol.Entities)
                                    //    {
                                    //        if (invDate >= ent.GetAttributeValue<DateTime>("hil_validfrom").AddMinutes(330) && invDate <= ent.GetAttributeValue<DateTime>("hil_validto").AddMinutes(330))
                                    //        {
                                    //            endDate = CreateSchemeUnitWarrantyLine(_service, ent.GetAttributeValue<int>("hil_warrantyperiod"), 7, erCustomer, CustAsst.ToEntityReference(), erProdCgry, erProdSubCategory, erProdModelCode, pdtModelName, stdWrtyEndDate.Value.AddDays(1), ent.ToEntityReference(), Convert.ToDateTime(invDate));
                                    //            if (endDate == stdWrtyEndDate.Value.AddDays(1))
                                    //            {
                                    //                endDate = stdWrtyEndDate;
                                    //            }
                                    //            else
                                    //            {
                                    //                spcWrtyEndDate = endDate;
                                    //                warrantySubstatus = 3;
                                    //                _specialSchemeApplied = true;
                                    //            }
                                    //            break;
                                    //        }
                                    //    }
                                    //}
                                    #endregion

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
                                        if (entCol.Entities[0].Attributes.Contains("hil_warrantyenddate"))
                                        {
                                            endDate = entCol.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330);
                                        }
                                        else
                                        {
                                            endDate = new DateTime(1900, 1, 1);
                                        }
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
                                                    if (DateTime.Now >= ent.GetAttributeValue<DateTime>("hil_warrantystartdate").AddMinutes(330) && DateTime.Now <= ent.GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330))
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
                                                        else if (entSubStatus.GetAttributeValue<OptionSetValue>("hil_type").Value == 3)
                                                        {
                                                            entCustAsset["hil_warrantysubstatus"] = new OptionSetValue(4); //InWarranty-AMC
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
                                        entCustAsset["hil_warrantytilldate"] = new DateTime(1900, 1, 1);
                                        entCustAsset["hil_warrantystatus"] = new OptionSetValue(2); //OutWarranty
                                        entCustAsset["hil_warrantysubstatus"] = null;
                                    }
                                    _service.Update(entCustAsset);
                                }
                                else
                                {
                                    Entity entCustAsset = new Entity(msdyn_customerasset.EntityLogicalName);
                                    entCustAsset.Id = CustAsst.Id;
                                    entCustAsset["hil_warrantytilldate"] = new DateTime(1900, 1, 1);
                                    entCustAsset["hil_warrantystatus"] = new OptionSetValue(2); //OutWarranty
                                    entCustAsset["hil_warrantysubstatus"] = null;
                                    _service.Update(entCustAsset);
                                }
                            }
                            else
                            {
                                Entity entCustAsset = new Entity(msdyn_customerasset.EntityLogicalName);
                                entCustAsset.Id = CustAsst.Id;
                                entCustAsset["hil_warrantytilldate"] = new DateTime(1900, 1, 1);
                                entCustAsset["hil_warrantystatus"] = new OptionSetValue(2); //OutWarranty
                                entCustAsset["hil_warrantysubstatus"] = null;
                                _service.Update(entCustAsset);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(CustAsst.GetAttributeValue<string>("msdyn_name").ToString() + " / " + ex.Message + " / " + rec.ToString() + "/" + Total.ToString());
                        }
                        Console.WriteLine(CustAsst.GetAttributeValue<DateTime>("createdon").ToString() + " / " + rec.ToString() + "/" + Total.ToString());
                    }
                    entAMC["hil_amcstagingstatus"] = true;
                    _service.Update(entAMC);
                    rec++;
                }
            }
        }
        static void RefreshAssetUnitWarrantyLLOYD()
        {

            #region Backup
            //string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
            //<entity name='msdyn_customerasset'>
            //<attribute name='hil_invoicedate' />
            //<attribute name='msdyn_customerassetid' />
            //<attribute name='hil_invoiceavailable' />
            //<attribute name='statuscode' />
            //<attribute name='hil_productcategory' />
            //<attribute name='hil_productsubcategory' />
            //<attribute name='msdyn_product' />
            //<attribute name='hil_modelname' />
            //<attribute name='msdyn_name' />
            //<attribute name='hil_customer' />
            //<order attribute='createdon' descending='true' />
            //<filter type='and'>
            //<condition attribute='hil_customernm' operator='ne' value='DONE' />
            //</filter>
            //<link-entity name='hil_unitwarranty' from='hil_customerasset' to='msdyn_customerassetid' link-type='inner' alias='ad'>
            //    <filter type='and'>
            //    <condition attribute='hil_customer' operator='not-null' />
            //    </filter>
            //    <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='ae'>
            //    <filter type='and'>
            //        <condition attribute='modifiedby' operator='eq' uiname='Rahul Panchal' uitype='systemuser' value='{FAC2D349-2AFD-E811-A94C-000D3AF060A1}' />
            //        <condition attribute='modifiedon' operator='on-or-after' value='2020-08-10' />
            //        <condition attribute='createdon' operator='on-or-before' value='2020-08-09' />
            //    </filter>
            //    <link-entity name='product' from='productid' to='hil_product' link-type='inner' alias='af'>
            //        <filter type='and'>
            //        <condition attribute='hil_division' operator='eq' uiname='LLOYD LED TELEVISION' uitype='product' value='{A7A5049B-16FA-E811-A94C-000D3AF06091}' />
            //        </filter>
            //    </link-entity>
            //    </link-entity>
            //</link-entity>
            //</entity>
            //</fetch>";

            //string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            //<entity name='msdyn_customerasset'>
            //<attribute name='hil_invoicedate' />
            //<attribute name='hil_invoiceavailable' />
            //<attribute name='statuscode' />
            //<attribute name='hil_productcategory' />
            //<attribute name='hil_productsubcategory' />
            //<attribute name='msdyn_product' />
            //<attribute name='hil_modelname' />
            //<attribute name='msdyn_name' />
            //<attribute name='hil_customer' />
            //<order attribute='createdon' descending='true' />
            //<filter type='and'>
            //    <condition attribute='hil_customernm' operator='ne' value='DONE' />
            //    <condition attribute='hil_productcategory' operator='eq' value='{d51edd9d-16fa-e811-a94c-000d3af0694e}' />
            //    <condition attribute='hil_warrantysubstatus' operator='null' />
            //    <condition attribute='createdon' operator='on-or-after' value='2020-07-01' />
            //    <condition attribute='createdon' operator='on-or-before' value='2020-08-01' />
            //    <filter type='or'>
            //    <condition attribute='statuscode' operator='eq' value='910590001' />
            //    <condition attribute='hil_branchheadapprovalstatus' operator='eq' value='1' />
            //    </filter>
            //    <condition attribute='hil_invoiceavailable' operator='eq' value='1' />
            //</filter>
            //</entity>
            //</fetch>";

            //string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            //<entity name='msdyn_customerasset'>
            //<attribute name='hil_invoicedate' />
            //<attribute name='hil_invoiceavailable' />
            //<attribute name='statuscode' />
            //<attribute name='hil_productcategory' />
            //<attribute name='hil_productsubcategory' />
            //<attribute name='msdyn_product' />
            //<attribute name='hil_modelname' />
            //<attribute name='msdyn_name' />
            //<attribute name='hil_customer' />
            //<attribute name='createdon' />
            //<order attribute='createdon' descending='false' />
            //<filter type='and'>
            //<condition attribute='hil_productcategory' operator='in'>
            //<value uiname='LLOYD WASHING MACHIN' uitype='product'>{2FD99DA1-16FA-E811-A94C-000D3AF06091}</value>
            //<value uiname='LLOYD LED TELEVISION' uitype='product'>{A7A5049B-16FA-E811-A94C-000D3AF06091}</value>
            //<value uiname='LLOYD REFRIGERATORS' uitype='product'>{2DD99DA1-16FA-E811-A94C-000D3AF06091}</value>
            //<value uiname='LLOYD AIR CONDITIONER' uitype='product'>{D51EDD9D-16FA-E811-A94C-000D3AF0694E}</value>
            //</condition>
            //<filter type='or'>
            //<condition attribute='statuscode' operator='eq' value='910590001' />
            //<condition attribute='hil_branchheadapprovalstatus' operator='eq' value='1' />
            //</filter>
            //<condition attribute='hil_invoiceavailable' operator='eq' value='1' />
            //<condition attribute='hil_invoicedate' operator='not-null' />
            //<condition attribute='hil_warrantysubstatus' operator='null' />
            //<condition attribute='hil_warrantystatus' operator='ne' value='2' />
            //</filter>
            //</entity>
            //</fetch>";

            #endregion

            string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='hil_unitwarranty'>
                <attribute name='hil_name' />
                <attribute name='createdon' />
                <attribute name='hil_warrantytemplate' />
                <attribute name='hil_warrantystartdate' />
                <attribute name='hil_warrantyenddate' />
                <attribute name='hil_producttype' />
                <attribute name='hil_customerasset' />
                <attribute name='hil_unitwarrantyid' />
            <order attribute='hil_name' descending='false' />
            <filter type='and'>
            <condition attribute='hil_customer' operator='not-null' />
            <condition attribute='hil_productmodel' operator='in'>
                <value>{D51EDD9D-16FA-E811-A94C-000D3AF0694E}</value>
                <value>{A7A5049B-16FA-E811-A94C-000D3AF06091}</value>
                <value>{2DD99DA1-16FA-E811-A94C-000D3AF06091}</value>
                <value>{2FD99DA1-16FA-E811-A94C-000D3AF06091}</value>
            </condition>
            <condition attribute='statecode' operator='eq' value='0' />
            </filter>
            <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='ai'>
            <filter type='and'>
            <condition attribute='hil_type' operator='eq' value='7' />
            </filter>
            </link-entity>
            <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='hil_customerasset' link-type='inner' alias='aj'>
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
            <filter type='and'>
                <condition attribute='hil_customernm' operator='ne' value='REDONE' />
                <condition attribute='createdon' operator='on-or-after' value='2020-08-01' />
                <condition attribute='createdon' operator='on-or-before' value='2020-09-04' />
            </filter>
            </link-entity>
            </entity>
            </fetch>";

            int rec = 0;
            int Total = 0;
            while (true)
            {
                EntityCollection ec = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                Total += ec.Entities.Count;
                if (ec.Entities.Count == 0) { break; }
                foreach (Entity CustAsst in ec.Entities)
                {
                    try
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

                        if (CustAsst.Attributes.Contains("aj.hil_invoiceavailable") && CustAsst.Attributes.Contains("aj.hil_invoicedate") && CustAsst.Attributes.Contains("aj.hil_productcategory"))
                        {
                            invoiceAvailable = (bool)((AliasedValue)CustAsst["aj.hil_invoiceavailable"]).Value;
                            EntityReference _customerAssetER = CustAsst.GetAttributeValue<EntityReference>("hil_customerasset");
                            if (invoiceAvailable)
                            {
                                QueryExpression queryExp1 = new QueryExpression("hil_unitwarranty");
                                queryExp1.ColumnSet = new ColumnSet("hil_unitwarrantyid", "hil_customer", "hil_warrantytemplate");
                                queryExp1.Criteria = new FilterExpression(LogicalOperator.And);
                                queryExp1.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, _customerAssetER.Id);
                                queryExp1.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                                EntityCollection entCol1 = _service.RetrieveMultiple(queryExp1);
                                if (entCol1.Entities.Count > 0)
                                {
                                    foreach (Entity ent in entCol1.Entities)
                                    {
                                        Entity entTemp = _service.Retrieve("hil_warrantytemplate", ent.GetAttributeValue<EntityReference>("hil_warrantytemplate").Id,new ColumnSet("hil_type"));
                                        if (entTemp != null) {
                                            if (entTemp.GetAttributeValue<OptionSetValue>("hil_type").Value == 2 || entTemp.GetAttributeValue<OptionSetValue>("hil_type").Value == 7)
                                            {
                                                _service.Delete("hil_unitwarranty", ent.Id);
                                            }
                                        }
                                    }
                                }

                                invDate = (DateTime)((AliasedValue)CustAsst["aj.hil_invoicedate"]).Value;
                                erProdCgry = (EntityReference)((AliasedValue)CustAsst["aj.hil_productcategory"]).Value;

                                if (CustAsst.Attributes.Contains("aj.hil_productsubcategory"))
                                {
                                    erProdSubCategory = (EntityReference)((AliasedValue)CustAsst["aj.hil_productsubcategory"]).Value;
                                }
                                if (CustAsst.Attributes.Contains("aj.msdyn_product"))
                                {
                                    erProdModelCode = (EntityReference)((AliasedValue)CustAsst["aj.msdyn_product"]).Value;
                                }
                                if (CustAsst.Attributes.Contains("aj.hil_modelname"))
                                {
                                    pdtModelName = ((AliasedValue)CustAsst["aj.hil_modelname"]).Value.ToString();
                                }
                                if (CustAsst.Attributes.Contains("aj.hil_customer"))
                                {
                                    erCustomer = (EntityReference)((AliasedValue)CustAsst["aj.hil_customer"]).Value;
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
                                        if (invDate >= ent.GetAttributeValue<DateTime>("hil_validfrom").AddMinutes(330) && invDate <= ent.GetAttributeValue<DateTime>("hil_validto").AddMinutes(330))
                                        {
                                            endDate = CreateUnitWarrantyLine(_service, ent.GetAttributeValue<int>("hil_warrantyperiod"), 1, erCustomer, _customerAssetER, erProdCgry, erProdSubCategory, erProdModelCode, pdtModelName, invDate, ent.ToEntityReference());
                                            warrantySubstatus = 1;
                                            break;
                                        }
                                    }
                                }
                                #endregion

                                #region Special Scheme Warranty Template
                                bool _specialSchemeApplied = false;
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
                                        if (invDate >= ent.GetAttributeValue<DateTime>("hil_validfrom").AddMinutes(330) && invDate <= ent.GetAttributeValue<DateTime>("hil_validto").AddMinutes(330))
                                        {
                                            endDate = CreateSchemeUnitWarrantyLine(_service, ent.GetAttributeValue<int>("hil_warrantyperiod"), 7, erCustomer, _customerAssetER, erProdCgry, erProdSubCategory, erProdModelCode, pdtModelName, endDate.Value.AddDays(1), ent.ToEntityReference(), Convert.ToDateTime(invDate));
                                            warrantySubstatus = 3;
                                            _specialSchemeApplied = true;
                                            break;
                                        }
                                    }
                                }
                                #endregion

                                #region Extended Warranty Template
                                if (endDate != null && !_specialSchemeApplied)
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
                                        if (invDate >= ent.GetAttributeValue<DateTime>("hil_validfrom").AddMinutes(330) && invDate <= ent.GetAttributeValue<DateTime>("hil_validto").AddMinutes(330))
                                        {
                                            endDate = CreateUnitWarrantyLine(_service, ent.GetAttributeValue<int>("hil_warrantyperiod"), 2, erCustomer, _customerAssetER, erProdCgry, erProdSubCategory, erProdModelCode, pdtModelName, endDate.Value.AddDays(1), ent.ToEntityReference());
                                            warrantySubstatus = 2;
                                            break;
                                        }
                                    }
                                }
                                #endregion

                                queryExp = new QueryExpression(hil_unitwarranty.EntityLogicalName);
                                queryExp.ColumnSet = new ColumnSet("hil_warrantyenddate", "hil_warrantytemplate");
                                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                queryExp.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, _customerAssetER.Id);
                                queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                queryExp.AddOrder("hil_warrantyenddate", OrderType.Descending);
                                queryExp.TopCount = 1;
                                entCol = _service.RetrieveMultiple(queryExp);

                                Entity entCustAsset = new Entity(msdyn_customerasset.EntityLogicalName);
                                entCustAsset.Id = _customerAssetER.Id;

                                if (entCol.Entities.Count > 0)
                                {
                                    if (entCol.Entities[0].Attributes.Contains("hil_warrantyenddate"))
                                    {
                                        endDate = entCol.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330);
                                    }
                                    else
                                    {
                                        endDate = new DateTime(1900, 1, 1);
                                    }
                                    entCustAsset["hil_warrantytilldate"] = endDate;
                                    if (endDate >= DateTime.Now)
                                    {
                                        entCustAsset["hil_warrantystatus"] = new OptionSetValue(1); //InWarranty

                                        queryExp = new QueryExpression(hil_unitwarranty.EntityLogicalName);
                                        queryExp.ColumnSet = new ColumnSet("hil_warrantystartdate", "hil_warrantyenddate", "hil_warrantytemplate");
                                        queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                        queryExp.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, _customerAssetER.Id);
                                        queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                        queryExp.AddOrder("hil_warrantyenddate", OrderType.Ascending);
                                        entCol = _service.RetrieveMultiple(queryExp);
                                        if (entCol.Entities.Count > 0)
                                        {
                                            foreach (Entity ent in entCol.Entities)
                                            {
                                                if (DateTime.Now >= ent.GetAttributeValue<DateTime>("hil_warrantystartdate").AddMinutes(330) && DateTime.Now <= ent.GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330))
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
                                    entCustAsset["hil_warrantytilldate"] = new DateTime(1900, 1, 1);
                                    entCustAsset["hil_warrantystatus"] = new OptionSetValue(2); //OutWarranty
                                    entCustAsset["hil_warrantysubstatus"] = null;
                                }
                                //hil_customernm
                                entCustAsset["hil_customernm"] = "REDONE";
                                _service.Update(entCustAsset);
                            }
                            else
                            {
                                Entity entCustAsset = new Entity(msdyn_customerasset.EntityLogicalName);
                                entCustAsset.Id = CustAsst.Id;
                                entCustAsset["hil_warrantytilldate"] = new DateTime(1900, 1, 1);
                                entCustAsset["hil_warrantystatus"] = new OptionSetValue(2); //OutWarranty
                                entCustAsset["hil_warrantysubstatus"] = null;
                                entCustAsset["hil_customernm"] = "REDONE";
                                _service.Update(entCustAsset);
                            }
                        }
                        else
                        {
                            Entity entCustAsset = new Entity(msdyn_customerasset.EntityLogicalName);
                            entCustAsset.Id = CustAsst.Id;
                            entCustAsset["hil_warrantytilldate"] = new DateTime(1900, 1, 1);
                            entCustAsset["hil_warrantystatus"] = new OptionSetValue(2); //OutWarranty
                            entCustAsset["hil_warrantysubstatus"] = null;
                            entCustAsset["hil_customernm"] = "REDONE";
                            _service.Update(entCustAsset);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(((AliasedValue)CustAsst["aj.msdyn_name"]).Value.ToString() + " / " + ex.Message + " / " + rec++.ToString() + "/" + Total.ToString());
                    }
                    Console.WriteLine(((AliasedValue)CustAsst["aj.createdon"]).Value.ToString() + " / " + rec++.ToString() + "/" + Total.ToString());
                }
            }
        }
        static void DeactivateUnitWarrantyLine() {

            string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='hil_unitwarranty'>
                <attribute name='hil_name' />
                <attribute name='createdon' />
                <attribute name='hil_unitwarrantyid' />
                <order attribute='hil_name' descending='false' />
                <filter type='and'>
                  <condition attribute='statecode' operator='eq' value='0' />
                  <condition attribute='createdon' operator='on-or-after' value='2020-05-01' />
                  <condition attribute='createdon' operator='on-or-before' value='2020-09-01' />
                </filter>
                <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='ab'>
                  <filter type='and'>
                    <condition attribute='hil_type' operator='not-in'>
                      <value>1</value>
                      <value>7</value>
                      <value>2</value>
                    </condition>
                  </filter>
                </link-entity>
              </entity>
            </fetch>";

            int rec = 0;
            int Total = 0;
            while (true)
            {
                EntityCollection ec = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                Total += ec.Entities.Count;
                if (ec.Entities.Count == 0) { break; }
                foreach (Entity ent in ec.Entities)
                {
                    try
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
                    catch (Exception ex)
                    {
                        Console.WriteLine(ent.GetAttributeValue<string>("hil_name").ToString() + " / " + ex.Message + " / " + rec++.ToString() + "/" + Total.ToString());
                    }
                    Console.WriteLine(ent.GetAttributeValue<DateTime>("createdon").ToString() + " / " + rec++.ToString() + "/" + Total.ToString());
                }
            }
        }
        static void DeleteAssetUnitWarranty()
        {
            string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='hil_unitwarranty'>
            <attribute name='hil_unitwarrantyid' />
            <order attribute='hil_name' descending='false' />
            <filter type='and'>
                <condition attribute='hil_productmodel' operator='eq' value='{2FD99DA1-16FA-E811-A94C-000D3AF06091}' />
                <condition attribute='createdon' operator='last-x-hours' value='3' />
                <condition attribute='statecode' operator='eq' value='0' />
            </filter>
            <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='af'>
            <filter type='and'>
            <condition attribute='hil_type' operator='in'>
                <value>2</value>
                <value>1</value>
                <value>7</value>
            </condition>
            <condition attribute='hil_templatestatus' operator='eq' value='2' />
            </filter>
            </link-entity>
            </entity>
            </fetch>";

            int rec = 0;
            int Total = 0;
            while (true)
            {
                EntityCollection ec = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                Total += ec.Entities.Count;
                if (ec.Entities.Count == 0) { break; }
                foreach (Entity CustAsst in ec.Entities)
                {
                    _service.Delete("hil_unitwarranty", CustAsst.Id);
                    Console.WriteLine(rec++.ToString() + "/" + Total.ToString());
                }
            }
        }
        static void CancelBulkJobs(IOrganizationService service)
        {
            string filePath = @"C:\Kuldeep khare\BulkJobCancel12Apr2024.xlsx";
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();

            Microsoft.Office.Interop.Excel.Application excelApp = new Excel.Application();
            if (excelApp != null)
            {
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[1];

                Excel.Range excelRange = excelWorksheet.UsedRange;
                Excel.Range range;
                string _mobileNo;
                EntityCollection entcoll = null;
                Entity entJob = null;
                string _fetchXML;

                for (int i = 2; i <= 1606; i++)
                {
                    range = (excelWorksheet.Cells[i, 1] as Excel.Range);
                    _mobileNo = range.Value.ToString();

                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorder'>
                        <attribute name='msdyn_name' />
                        <order attribute='createdon' descending='true' />
                        <filter type='and'>
                            <condition attribute='hil_mobilenumber' operator='eq' value='{_mobileNo}' />
                            <condition attribute='msdyn_substatus' operator='not-in'>
                                <value uiname='Canceled' uitype='msdyn_workordersubstatus'>{{1527FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                                <value uiname='Closed' uitype='msdyn_workordersubstatus'>{{1727FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                                <value uiname='KKG Audit Failed' uitype='msdyn_workordersubstatus'>{{6C8F2123-5106-EA11-A811-000D3AF057DD}}</value>
                                <value uiname='Work Done' uitype='msdyn_workordersubstatus'>{{2927FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                                <value uiname='Work Done SMS' uitype='msdyn_workordersubstatus'>{{7E85074C-9C54-E911-A951-000D3AF0677F}}</value>
                            </condition>
                            <condition attribute='hil_isocr' operator='ne' value='1' />
                        </filter>
                        <link-entity name='hil_productrequest' from='hil_job' to='msdyn_workorderid' link-type='outer' alias='ad' />
                        <filter type='and'>
                            <condition entityname='ad' attribute='hil_job' operator='null' />
                        </filter>
                        </entity>
                    </fetch>";
                    entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entcoll.Entities.Count > 0)
                    {
                        int j = 0;
                        foreach (Entity ent in entcoll.Entities)
                        {
                            j += 1;
                            if (j == 1) { continue; }
                            entJob = new Entity("msdyn_workorder", ent.Id);
                            entJob["hil_webclosureremarks"] = "Consumer created duplicate Jobs from multiple Channels. Action taken on Dt: 15Apr2024.";
                            entJob["hil_closureremarks"] = "Consumer created duplicate Jobs from multiple Channels. Action taken on Dt: 15Apr2024.";
                            entJob["hil_isocr"] = true;
                            try
                            {
                                _service.Update(entJob);
                                Console.WriteLine($"Processing... Mobile Number {_mobileNo} /Job Number {ent.GetAttributeValue<string>("msdyn_name")} Record# {i.ToString()}/{j.ToString()}");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Processing... Mobile Number {_mobileNo} /Job Number {ent.GetAttributeValue<string>("msdyn_name")} Record# {i.ToString()}/{j.ToString()} ERROR! " + ex.Message);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine(_mobileNo + " does not exist.");
                    }
                }
            }
        }
        //static void CancelBulkJobs(IOrganizationService service)
        //{
        //    string filePath = @"C:\Kuldeep khare\BulkJobCancel12Apr2024.xlsx";
        //    string conn = string.Empty;
        //    DataTable dtexcel = new DataTable();

        //    Microsoft.Office.Interop.Excel.Application excelApp = new Excel.Application();
        //    if (excelApp != null)
        //    {
        //        Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
        //        Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[1];

        //        Excel.Range excelRange = excelWorksheet.UsedRange;
        //        int rowCount = 1;
        //        Excel.Range range;
        //        string jobId;
        //        QueryExpression qryExp;
        //        EntityCollection entcoll= null;
        //        Entity entJob = null;
        //        string _fetchXML;

        //        for (int i = 2; i <= 1090; i++)
        //        {
        //            range = (excelWorksheet.Cells[i, 1] as Excel.Range);
        //            jobId = range.Value.ToString();

        //            _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                  <entity name='msdyn_workorder'>
        //                    <attribute name='msdyn_workorderid' />
        //                    <order attribute='msdyn_name' descending='false' />
        //                    <filter type='and'>
        //                      <condition attribute='msdyn_name' operator='eq' value='{jobId}' />
        //                    </filter>
        //                  </entity>
        //                </fetch>";
        //            entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //            if (entcoll.Entities.Count > 0)
        //            {
        //                entJob = new Entity("msdyn_workorder", entcoll.Entities[0].Id);
        //                entJob["hil_webclosureremarks"] = "Duplicate job through excel upload.";
        //                entJob["msdyn_substatus"] = new EntityReference("msdyn_workordersubstatus", new Guid("1527FA6C-FA0F-E911-A94E-000D3AF060A1"));
        //                entJob["hil_jobcancelreason"] = new OptionSetValue(4);//Duplicate Request
        //                entJob["hil_closureremarks"] = "Duplicate job through excel upload.";
        //                entJob["hil_cancelticket"] = true;
        //                try
        //                {
        //                    _service.Update(entJob);
        //                    Console.WriteLine("Processing... " + jobId + " / " + i.ToString());
        //                }
        //                catch (Exception ex)
        //                {
        //                    Console.WriteLine("Processing... " + jobId + " / " + i.ToString() + " /" + ex.Message);
        //                }
        //            }
        //            else {
        //                Console.WriteLine(jobId + " does not exist.");
        //            }

        //        }
        //    }
        //}
        static void UpdateCustomerCategory(IOrganizationService service)
        {
            string filePath = @"C:\Kuldeep khare\ChannelPartnersCategory.xls";
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();

            Microsoft.Office.Interop.Excel.Application excelApp = new Excel.Application();
            if (excelApp != null)
            {
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[1];

                Excel.Range excelRange = excelWorksheet.UsedRange;
                int rowCount = 1;
                Excel.Range range;
                string customerCode, CategoryCode;
                QueryExpression Query1;
                EntityCollection entcoll1;

                for (int i = 2; i <= 58; i++)
                {
                    range = (excelWorksheet.Cells[i, 1] as Excel.Range);
                    customerCode = range.Value.ToString();
                    range = (excelWorksheet.Cells[i, 2] as Excel.Range);
                    CategoryCode = range.Value.ToString();

                    Query1 = new QueryExpression("account");
                    Query1.ColumnSet = new ColumnSet("accountid", "hil_category");
                    Query1.Criteria = new FilterExpression(LogicalOperator.And);
                    Query1.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, customerCode);
                    entcoll1 = service.RetrieveMultiple(Query1);
                    if (entcoll1.Entities.Count > 0)
                    {
                        OptionSetValue opValue = null;
                        foreach (Entity ent in entcoll1.Entities)
                        {
                            if (CategoryCode == "A")
                                opValue = new OptionSetValue(910590000);
                            else if (CategoryCode == "B")
                                opValue = new OptionSetValue(910590001);
                            else if (CategoryCode == "C")
                                opValue = new OptionSetValue(910590002);
                            else if (CategoryCode == "D")
                                opValue = new OptionSetValue(910590003);

                            ent["hil_category"] = opValue;
                            service.Update(ent);
                            Console.WriteLine(i.ToString() + "/" + rowCount.ToString());
                        }
                    }

                }
            }
        }
        static void UpdateCustomerVendorCode(IOrganizationService service)
        {
            string filePath = @"C:\MyWorkSpace\CustomerData1.xls";
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();

            Microsoft.Office.Interop.Excel.Application excelApp = new Excel.Application();
            if (excelApp != null)
            {
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[1];

                Excel.Range excelRange = excelWorksheet.UsedRange;
                int rowCount = 1;
                Excel.Range range;
                string customerCode, vendorCode;
                QueryExpression Query1;
                EntityCollection entcoll1;

                for (int i = 2; i <= 1000; i++)
                {
                    range = (excelWorksheet.Cells[i, 1] as Excel.Range);
                    customerCode = range.Value.ToString();
                    range = (excelWorksheet.Cells[i, 2] as Excel.Range);
                    vendorCode = range.Value.ToString();

                    Query1 = new QueryExpression("account");
                    Query1.ColumnSet = new ColumnSet("accountid", "hil_vendorcode");
                    Query1.Criteria = new FilterExpression(LogicalOperator.And);
                    Query1.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, customerCode);
                    entcoll1 = service.RetrieveMultiple(Query1);
                    if (entcoll1.Entities.Count > 0)
                    {
                        foreach (Entity ent in entcoll1.Entities)
                        {
                            ent["hil_vendorcode"] = vendorCode;
                            service.Update(ent);
                            Console.WriteLine(i.ToString() + "/" + rowCount.ToString());
                        }
                    }

                }
            }
        }
        static void UpdateClaimParameterOnJob() {
            Guid _jobGuId = Guid.Empty;
            EntityReference _unitWarranty = null;
            Guid _warrantyTemplateId = Guid.Empty;
            bool SparePartUsed = false;
            string _fetchXML = string.Empty;
            QueryExpression qrExp;
            EntityCollection entCol;
            bool laborInWarranty = false;
            int _jobWarrantyStatus = 2; //OutWarranty
            int _jobWarrantySubStatus = 0;
            int _warrantyTempType = 0;
            DateTime _unitWarrStartDate = new DateTime(1900, 1, 1);
            double _jobMonth = 0;

            _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='msdyn_workorder'>
                <attribute name='msdyn_name' />
                <attribute name='createdon' />
                <attribute name='msdyn_timeclosed' />
                <attribute name='hil_owneraccount' />
                <attribute name='msdyn_customerasset' />
                <attribute name='hil_purchasedate' />
                <order attribute='msdyn_timeclosed' descending='false' />
                <filter type='and'>
                    <condition attribute='msdyn_timeclosed' operator='on-or-after' value='2020-09-08' />
                    <condition attribute='msdyn_timeclosed' operator='on-or-before' value='2020-09-09' />
                    <condition attribute='hil_isocr' operator='ne' value='1' />
                    <condition attribute='msdyn_substatus' operator='in'>
                    <value>{2927FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                    <value>{1727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                    </condition>
                    <condition attribute='hil_systemremarks' operator='null'/>
                    <condition attribute='msdyn_customerasset' operator='not-null'/>
                </filter>
                </entity>
            </fetch>";

            while (true)
            {
                entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entCol.Entities.Count == 0) { break; }
                int i = 1;
                foreach (Entity ent in entCol.Entities)
                {
                    try
                    {
                        _jobGuId = ent.Id;
                        if (ent != null)
                        {
                            DateTime _jobCreatedOn = ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330);

                            DateTime _assetPurchaseDate = ent.GetAttributeValue<DateTime>("hil_purchasedate").AddMinutes(330);
                            if (ent.Attributes.Contains("hil_purchasedate")) { }

                            DateTime _jobClosedOn = ent.GetAttributeValue<DateTime>("msdyn_timeclosed").AddMinutes(330);

                            qrExp = new QueryExpression("msdyn_workorderincident");
                            qrExp.ColumnSet = new ColumnSet("msdyn_workorderincidentid");
                            qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                            qrExp.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, _jobGuId);
                            qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            qrExp.Criteria.AddCondition("hil_warrantystatus", ConditionOperator.Equal, 3); // Warranty Void
                            entCol = _service.RetrieveMultiple(qrExp);
                            if (entCol.Entities.Count == 0)
                            {
                                qrExp = new QueryExpression("hil_unitwarranty");
                                qrExp.ColumnSet = new ColumnSet("hil_warrantytemplate", "hil_warrantystartdate", "hil_warrantyenddate");
                                qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                                qrExp.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("msdyn_customerasset").Id);
                                qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                entCol = _service.RetrieveMultiple(qrExp);
                                if (entCol.Entities.Count > 0)
                                {
                                    foreach (Entity Wt in entCol.Entities)
                                    {
                                        DateTime iValidTo = ((DateTime)Wt["hil_warrantyenddate"]).AddMinutes(330);
                                        DateTime iValidFrom = ((DateTime)Wt["hil_warrantystartdate"]).AddMinutes(330);
                                        if (_jobCreatedOn >= iValidFrom && _jobCreatedOn <= iValidTo)
                                        {
                                            _jobWarrantyStatus = 1;
                                            _unitWarranty = Wt.ToEntityReference();
                                            _warrantyTemplateId = Wt.GetAttributeValue<EntityReference>("hil_warrantytemplate").Id;
                                            Entity _entTemp = _service.Retrieve(hil_warrantytemplate.EntityLogicalName, _warrantyTemplateId, new ColumnSet("hil_type"));
                                            if (_entTemp != null)
                                            {
                                                _warrantyTempType = _entTemp.GetAttributeValue<OptionSetValue>("hil_type").Value;
                                                if (_warrantyTempType == 1) { _jobWarrantySubStatus = 1; }
                                                else if (_warrantyTempType == 2) { _jobWarrantySubStatus = 2; }
                                                else if (_warrantyTempType == 7) { _jobWarrantySubStatus = 3; }
                                                else if (_warrantyTempType == 3) { _jobWarrantySubStatus = 4; }
                                            }
                                            _unitWarrStartDate = Wt.GetAttributeValue<DateTime>("hil_warrantystartdate").AddMinutes(330);
                                            TimeSpan difference = (_jobCreatedOn - _unitWarrStartDate);
                                            _jobMonth = Math.Round((difference.Days * 1.0 / 30.42), 0);
                                            qrExp = new QueryExpression("hil_labor");
                                            qrExp.ColumnSet = new ColumnSet("hil_laborid", "hil_includedinwarranty", "hil_validtomonths", "hil_validfrommonths");
                                            qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                                            qrExp.Criteria.AddCondition("hil_warrantytemplateid", ConditionOperator.Equal, _warrantyTemplateId);
                                            qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                            entCol = _service.RetrieveMultiple(qrExp);
                                            if (entCol.Entities.Count == 0) { laborInWarranty = true; }
                                            else
                                            {
                                                if (_jobMonth >= entCol.Entities[0].GetAttributeValue<int>("hil_validfrommonths") && _jobMonth <= entCol.Entities[0].GetAttributeValue<int>("hil_validtomonths"))
                                                {
                                                    OptionSetValue _laborType = entCol.Entities[0].GetAttributeValue<OptionSetValue>("hil_includedinwarranty");
                                                    laborInWarranty = _laborType.Value == 1 ? true : false;
                                                }
                                                else
                                                {
                                                    OptionSetValue _laborType = entCol.Entities[0].GetAttributeValue<OptionSetValue>("hil_includedinwarranty");
                                                    laborInWarranty = !(_laborType.Value == 1 ? true : false);
                                                }
                                            }
                                            break;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                _jobWarrantyStatus = 3;
                            }
                            qrExp = new QueryExpression("msdyn_workorderproduct");
                            qrExp.ColumnSet = new ColumnSet("msdyn_workorderproductid");
                            qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                            qrExp.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, _jobGuId);
                            qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            qrExp.Criteria.AddCondition("hil_markused", ConditionOperator.Equal, true);
                            entCol = _service.RetrieveMultiple(qrExp);
                            if (entCol.Entities.Count > 0) { SparePartUsed = true; }

                            bool _claimstatus = false;
                            qrExp = new QueryExpression("hil_sawactivity");
                            qrExp.ColumnSet = new ColumnSet("hil_sawactivityid");
                            qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                            qrExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _jobGuId);
                            entCol = _service.RetrieveMultiple(qrExp);
                            if (entCol.Entities.Count == 0)
                            {
                                _claimstatus = true;
                            }

                            Entity entJobUpdate = new Entity(msdyn_workorder.EntityLogicalName, _jobGuId);

                            entJobUpdate["hil_warrantystatus"] = new OptionSetValue(_jobWarrantyStatus);
                            if (_jobWarrantyStatus == 1)
                            {
                                entJobUpdate["hil_warrantysubstatus"] = new OptionSetValue(_jobWarrantySubStatus);
                            }
                            if (_unitWarranty != null)
                            {
                                entJobUpdate["hil_unitwarranty"] = _unitWarranty;
                            }
                            entJobUpdate["hil_laborinwarranty"] = laborInWarranty;
                            if (_assetPurchaseDate.Year != 1900 && _assetPurchaseDate.Year != 1)
                            {
                                entJobUpdate["hil_purchasedate"] = _assetPurchaseDate;
                            }

                            entJobUpdate["hil_sparepartuse"] = SparePartUsed;
                            entJobUpdate["hil_systemremarks"] = "DONE";
                            if (_claimstatus)
                            {
                                entJobUpdate["hil_claimstatus"] = new OptionSetValue(4); //Claim Approved
                            }
                            _service.Update(entJobUpdate);
                            Console.WriteLine(i++.ToString());

                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
        }
        static void RepeatRepairSAW() {
            QueryExpression qrExp;
            EntityCollection entCol;
            int i = 0;
            int j = 0;
            bool _underReview = false;
            string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='msdyn_workorder'>
                <attribute name='msdyn_name' />
                <attribute name='createdon' />
                <attribute name='hil_productsubcategory' />
                <attribute name='hil_customerref' />
                <attribute name='hil_callsubtype' />
                <attribute name='msdyn_workorderid' />
                <order attribute='msdyn_name' descending='false' />
                <filter type='and'>
                    <condition attribute='msdyn_timeclosed' operator='on-or-after' value='2020-08-21' />
                    <condition attribute='msdyn_timeclosed' operator='on-or-before' value='2020-09-07' />
                    <condition attribute='hil_claimstatus' operator='ne' value='3' />
                    <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                    <condition attribute='hil_laborinwarranty' operator='eq' value='1' />
                    <condition attribute='msdyn_bookingsummary' operator='not-like' value='%RR%' />
                    <condition attribute='hil_callsubtype' operator='not-null' />
                    <condition attribute='msdyn_customerasset' operator='not-null' />
                    <condition attribute='hil_productsubcategory' operator='not-null' />
                </filter>
              </entity>
            </fetch>";
            while (true)
            {
                entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                j += entCol.Entities.Count;
                if (entCol.Entities.Count == 0) { break; }
                foreach (Entity ent in entCol.Entities)
                {
                    try
                    {
                        i++;
                        #region RepeatRepair Approval
                        DateTime _createdOn = ent.GetAttributeValue<DateTime>("createdon").AddDays(-15);
                        DateTime _ClosedOn = DateTime.Now.AddDays(-15);
                        string _strCreatedOn = _createdOn.Year.ToString() + "-" + _createdOn.Month.ToString() + "-" + _createdOn.Day.ToString();
                        string _strClosedOn = _ClosedOn.Year.ToString() + "-" + _ClosedOn.Month.ToString() + "-" + _ClosedOn.Day.ToString();

                        _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='msdyn_workorder'>
                    <attribute name='msdyn_name' />
                    <attribute name='createdon' />
                    <attribute name='hil_productsubcategory' />
                    <attribute name='hil_customerref' />
                    <attribute name='hil_callsubtype' />
                    <attribute name='msdyn_workorderid' />
                    <attribute name='msdyn_timeclosed' />
                    <attribute name='msdyn_closedby' />
                    <order attribute='msdyn_name' descending='false' />
                    <filter type='and'>
                        <condition attribute='msdyn_workorderid' operator='ne' value='" + ent.Id + @"' />
                        <condition attribute='hil_customerref' operator='eq' value='" + ent.GetAttributeValue<EntityReference>("hil_customerref").Id + @"' />
                        <condition attribute='hil_callsubtype' operator='eq' value='" + ent.GetAttributeValue<EntityReference>("hil_callsubtype").Id + @"' />
                        <condition attribute='hil_callsubtype' operator='ne' value='{8D80346B-3C0B-E911-A94E-000D3AF06CD4}' />
                        <condition attribute='hil_productsubcategory' operator='eq' value='" + ent.GetAttributeValue<EntityReference>("hil_productsubcategory").Id + @"' />
                        <filter type='or'>
                            <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strCreatedOn + @"' />
                            <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strClosedOn + @"' />
                        </filter>
                        <condition attribute='msdyn_substatus' operator='in'>
                        <value>{1727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                        <value>{2927FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                        <value>{7E85074C-9C54-E911-A951-000D3AF0677F}</value>
                        </condition>
                    </filter>
                    </entity>
                    </fetch>";
                        entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entCol.Entities.Count > 0)
                        {
                            string _remarks = "Old Job# " + entCol.Entities[0].GetAttributeValue<string>("msdyn_name");
                            CommonLib obj = new CommonLib();
                            CommonLib objReturn = obj.CreateSAWActivity(ent.Id, 0, SAWCategoryConst._RepeatRepair, _service, _remarks, entCol.Entities[0].ToEntityReference());
                            if (objReturn.statusRemarks == "OK")
                            {
                                _underReview = true;
                            }
                            if (_underReview)
                            {
                                Entity Ticket = new Entity("msdyn_workorder");
                                Ticket.Id = ent.Id;
                                Ticket["hil_claimstatus"] = new OptionSetValue(1); //Claim Under Review
                                Ticket["msdyn_bookingsummary"] = "RR";
                                _service.Update(Ticket);
                            }
                            else
                            {
                                Entity Ticket = new Entity("msdyn_workorder");
                                Ticket.Id = ent.Id;
                                Ticket["hil_claimstatus"] = new OptionSetValue(1); //Claim Under Review
                                Ticket["msdyn_bookingsummary"] = "RR";
                                _service.Update(Ticket);
                            }
                        }
                        else
                        {
                            Entity Ticket = new Entity("msdyn_workorder");
                            Ticket.Id = ent.Id;
                            Ticket["msdyn_bookingsummary"] = "RR";
                            _service.Update(Ticket);
                        }
                        Console.WriteLine(i.ToString() + "/" + j.ToString());
                        #endregion
                    }
                    catch { }
                }
            }
        }
        static DateTime? CreateUnitWarrantyLine(IOrganizationService service, int period, int producttype, EntityReference erCustomer, EntityReference erCustomerasset, EntityReference erProductCatg, EntityReference erProductSubCatg, EntityReference erProductModel, string partdescription, DateTime? warrantystartdate, EntityReference erWarrantytemplate)
        {
            DateTime? WarrantyEnd = null;
            try
            {
                DateTime StartDate = Convert.ToDateTime(warrantystartdate);
                WarrantyEnd = StartDate.AddMonths(period).AddDays(-1);

                QueryExpression qryExp = new QueryExpression(hil_unitwarranty.EntityLogicalName);
                qryExp.ColumnSet = new ColumnSet("hil_warrantyenddate", "hil_warrantystartdate", "hil_warrantyenddate");
                qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                qryExp.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, erCustomerasset.Id);
                qryExp.Criteria.AddCondition("hil_warrantytemplate", ConditionOperator.Equal, erWarrantytemplate.Id);
                qryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                EntityCollection entCol = service.RetrieveMultiple(qryExp);
                if (entCol.Entities.Count == 0)
                {
                    //int i = 0;
                    hil_unitwarranty iSchWarranty = new hil_unitwarranty();
                    iSchWarranty.hil_CustomerAsset = erCustomerasset;

                    iSchWarranty.hil_productmodel = erProductCatg;

                    iSchWarranty.hil_productitem = erProductSubCatg;
                    iSchWarranty.hil_warrantystartdate = Convert.ToDateTime(warrantystartdate);
                    
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
                else
                {
                    DateTime _WarrantyEntDate = entCol.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330);
                    //if (WarrantyEnd != _WarrantyEntDate)
                    //{
                    //    Entity _entUWL = new Entity("hil_unitwarranty", entCol.Entities[0].Id);
                    //    _entUWL["hil_warrantyenddate"] = WarrantyEnd;
                    //    _service.Update(_entUWL);
                    //}
                    WarrantyEnd = entCol.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330);
                }
                return WarrantyEnd;
            }
            catch {
                return WarrantyEnd;
            }
        }

        static DateTime? CreateSchemeUnitWarrantyLine(IOrganizationService service, int period, int producttype, EntityReference erCustomer, EntityReference erCustomerasset, EntityReference erProductCatg, EntityReference erProductSubCatg, EntityReference erProductModel, string partdescription, DateTime? warrantystartdate, EntityReference erWarrantytemplate,DateTime invDate)
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
                            if (invDate >= entCol.Entities[0].GetAttributeValue<DateTime>("hil_fromdate").AddMinutes(330) && invDate <= entCol.Entities[0].GetAttributeValue<DateTime>("hil_todate").AddMinutes(330))
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
                    else {
                        WarrantyEnd = warrantystartdate;
                    }
                }
                else
                {
                    WarrantyEnd = entCol.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate");
                }
                return WarrantyEnd;
            }
            catch {
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

    public class CommonLib
    {
        public string status { get; set; }
        public string statusRemarks { get; set; }

        public CommonLib CreateSAWActivity(Guid _jobId, decimal _amount, string _sawCategory, IOrganizationService _service, string _remarks, EntityReference _repeatRefjobId)
        {
            try
            {
                QueryExpression qryExp = new QueryExpression(msdyn_workorder.EntityLogicalName);
                qryExp.ColumnSet = new ColumnSet("hil_owneraccount", "ownerid");
                qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                qryExp.Criteria.AddCondition("msdyn_workorderid", ConditionOperator.Equal, _jobId);
                EntityCollection entCol = _service.RetrieveMultiple(qryExp);
                if (entCol.Entities.Count > 0)
                {
                    if (!CheckIfSAWActivityExist(_jobId, _service, _sawCategory))
                    {
                        Entity entSAWActivity = new Entity("hil_sawactivity");
                        entSAWActivity["hil_sawcategory"] = new EntityReference("hil_serviceactionwork", new Guid(_sawCategory));
                        entSAWActivity["hil_relatedchannelpartner"] = entCol.Entities[0].GetAttributeValue<EntityReference>("hil_owneraccount");
                        entSAWActivity["hil_jobid"] = new EntityReference("msdyn_workorder", _jobId);
                        entSAWActivity["hil_description"] = _remarks;
                        if (_repeatRefjobId != null)
                        {
                            entSAWActivity["hil_repeatreferencejob"] = _repeatRefjobId;
                        }
                        entSAWActivity["hil_amount"] = new Money(_amount);
                        entSAWActivity["hil_approvalstatus"] = new OptionSetValue(1); //requested
                        Guid sawActivityId = _service.Create(entSAWActivity);
                        CreateSAWActivityApprovals(sawActivityId, _service);
                    }
                    return new CommonLib() { status = "200", statusRemarks = "OK" };
                }
                else
                {
                    return new CommonLib() { status = "204", statusRemarks = "something went wrong." };
                }
            }
            catch (Exception ex)
            {
                return new CommonLib() { status = "204", statusRemarks = ex.Message };
            }
        }

        public CommonLib CreateSAWActivityApprovals(Guid _sawActivityId, IOrganizationService _service)
        {
            try
            {
                Entity ent = _service.Retrieve("hil_sawactivity", _sawActivityId, new ColumnSet("hil_sawcategory", "hil_jobid"));
                if (ent != null)
                {
                    EntityReference _salesOffice = null;
                    EntityReference _productCatg = null;
                    EntityReference _job = null;
                    EntityReference _picUser = null;
                    EntityReference _picPosition = null;
                    Entity entJob = _service.Retrieve("msdyn_workorder", ent.GetAttributeValue<EntityReference>("hil_jobid").Id, new ColumnSet("hil_salesoffice", "hil_productcategory"));
                    if (entJob != null)
                    {
                        if (entJob.Attributes.Contains("hil_salesoffice"))
                        {
                            _salesOffice = entJob.GetAttributeValue<EntityReference>("hil_salesoffice");
                            _productCatg = entJob.GetAttributeValue<EntityReference>("hil_productcategory");
                            _job = ent.GetAttributeValue<EntityReference>("hil_jobid");
                        }
                    }
                    if (_salesOffice != null && _productCatg != null)
                    {
                        QueryExpression qrySBU = new QueryExpression("hil_sbubranchmapping");
                        qrySBU.ColumnSet = new ColumnSet("hil_nsh", "hil_nph", "hil_branchheaduser");
                        qrySBU.Criteria = new FilterExpression(LogicalOperator.And);
                        qrySBU.Criteria.AddCondition("hil_salesoffice", ConditionOperator.Equal, _salesOffice.Id);
                        qrySBU.Criteria.AddCondition("hil_productdivision", ConditionOperator.Equal, _productCatg.Id);
                        EntityCollection entColSBU = _service.RetrieveMultiple(qrySBU);

                        QueryExpression qryExp = new QueryExpression("hil_sawcategoryapprovals");
                        qryExp.ColumnSet = new ColumnSet("hil_picuser", "hil_picposition", "hil_level");
                        qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qryExp.Criteria.AddCondition("hil_sawcategoryid", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_sawcategory").Id);
                        EntityCollection entCol = _service.RetrieveMultiple(qryExp);
                        if (entCol.Entities.Count > 0)
                        {
                            foreach (Entity entTemp in entCol.Entities)
                            {
                                if (entTemp.Attributes.Contains("hil_picuser"))
                                {
                                    _picUser = entTemp.GetAttributeValue<EntityReference>("hil_picuser");
                                }
                                else
                                {
                                    _picPosition = entTemp.GetAttributeValue<EntityReference>("hil_picposition");
                                    if (_picPosition.Name.ToUpper() == "BSH")
                                    {
                                        if (entColSBU.Entities.Count > 0)
                                        {
                                            _picUser = entColSBU.Entities[0].GetAttributeValue<EntityReference>("hil_branchheaduser");
                                        }
                                    }
                                    else if (_picPosition.Name.ToUpper() == "NSH")
                                    {
                                        _picUser = entColSBU.Entities[0].GetAttributeValue<EntityReference>("hil_nsh");

                                    }
                                    else if (_picPosition.Name.ToUpper() == "NPH")
                                    {
                                        _picUser = entColSBU.Entities[0].GetAttributeValue<EntityReference>("hil_nph");
                                    }
                                }
                                Entity entSAWActivity = new Entity("hil_sawactivityapproval");
                                entSAWActivity["hil_sawactivity"] = new EntityReference("hil_sawactivity", _sawActivityId);
                                entSAWActivity["hil_jobid"] = _job;
                                entSAWActivity["hil_level"] = entTemp.GetAttributeValue<OptionSetValue>("hil_level");
                                if (entTemp.GetAttributeValue<OptionSetValue>("hil_level").Value == 1)
                                {
                                    entSAWActivity["hil_isenabled"] = true;
                                }
                                entSAWActivity["hil_approver"] = _picUser;
                                entSAWActivity["hil_approvalstatus"] = new OptionSetValue(1); //requested
                                Guid sawActivityId = _service.Create(entSAWActivity);
                            }
                            return new CommonLib() { status = "200", statusRemarks = "OK" };
                        }
                        else
                        {
                            return new CommonLib() { status = "204", statusRemarks = "SAW Activity approvals not found." };
                        }
                    }
                    else
                    {
                        return new CommonLib() { status = "204", statusRemarks = "Sales Office not found in Job." };
                    }
                }
                else
                {
                    return new CommonLib() { status = "204", statusRemarks = "Something went wrong." };
                }
            }
            catch (Exception ex)
            {
                return new CommonLib() { status = "204", statusRemarks = ex.Message };
            }
        }


        public bool CheckIfSAWActivityExist(Guid _jobId, IOrganizationService _service, string _sawCategory)
        {
            try
            {
                QueryExpression qryExp = new QueryExpression("hil_sawactivity");
                qryExp.ColumnSet = new ColumnSet("hil_jobid");
                qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                qryExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _jobId);
                qryExp.Criteria.AddCondition("hil_sawcategory", ConditionOperator.Equal, new EntityReference("hil_serviceactionwork", new Guid(_sawCategory)).Id);
                EntityCollection entCol = _service.RetrieveMultiple(qryExp);
                if (entCol.Entities.Count > 0) { return true; } else { return false; }
            }
            catch
            {
                return false;
            }
        }
    }

    public class SAWCategoryConst
    {
        public const string _GasChargePostAudit = "f5784d62-4bdd-ea11-a813-000d3af0563c";
        public const string _GasChargePreAuditPastInstallationHistory = "9d01bc73-c5db-ea11-a813-000d3af055b6";
        public const string _GasChargePreAuditPastRepeatHistory = "a8e50e38-4bdd-ea11-a813-000d3af0563c";
        public const string _KKGFailureReview = "e123bd08-8add-ea11-a813-000d3af05a4b";
        public const string _LocalPurchase = "d0a3babb-f3d0-ea11-a813-000d3af05a4b";
        public const string _LocalRepair = "d577a3c8-f3d0-ea11-a813-000d3af05a4b";
        public const string _ProductTransportation = "db074fae-f3d0-ea11-a813-000d3af05a4b";
        public const string _RepeatRepair = "b0918d74-44ed-ea11-a815-000d3af05d7b";
        public const string _OneTimeLaborException = "ad96a922-0aee-ea11-a815-000d3af057dd";
    }
}
