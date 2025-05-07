using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Configuration;
using Microsoft.Xrm.Sdk.Query;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Net.Http.Headers;
using Microsoft.Crm.Sdk.Messages;
using System.Data;
using System.Linq;
using System.Runtime.Remoting.Contexts;

namespace Havells.CRM.DataMigration
{
    class Program : HelperClass
    {
        public static IOrganizationService _servicePrd = null;
        public static IOrganizationService _serviceDev = null;
        private static EntityReference systemAdminRef = null;
        static Program()
        {
            systemAdminRef = new EntityReference("systemuser", new Guid(SystemAdmin));
            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
            var CrmDevURL = ConfigurationManager.AppSettings["CRMDevUrl"].ToString();
            var CrmPrdURL = ConfigurationManager.AppSettings["CRMPrdUrl"].ToString();
            string finalDevString = string.Format(connStr, CrmDevURL);
            string finalPrdString = string.Format(connStr, CrmPrdURL);

            _servicePrd = createConnection(finalPrdString);
            _serviceDev = createConnection(finalDevString);
        }
        static void Main(string[] args)
        {

            try
            {

                //WarrantyTemplateMigration.warrantyTemplateMigration(_servicePrd, _serviceDev);
                string entityId = "a59b4a18-080b-ec11-b6e6-6045bd72f2f7";// context.InputParameters["EntityID"].ToString();
                _serviceDev.Create(_servicePrd.Retrieve("hil_approval", new Guid(entityId), new ColumnSet(true)));

            }
            catch (Exception ex)
            {
                //context.OutputParameters["ErrorOccured"] = true;
                //context.OutputParameters["Error"] = ex.Message;
            }
            //try
            //{
            //    //MigrateServiceChannelPartner.migrateServiceChannelPartner(_servicePrd, _serviceDev);
            //    //MigrateAssingmentMatrix.AssingmentMatrixgetCount(_servicePrd);
            //    // MigrateAssingmentMatrix._MigrateAssingmentMatrix(_servicePrd, _serviceDev);
            //    // MigrateBulkRecords(_servicePrd, _serviceDev, "hil_upcountrytravelcharge", systemAdminRef, true);
            //    //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_minimumstocklevel", systemAdminRef, true);
            //    //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_channelpartnercountryclassification", systemAdminRef, true);
            //    //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_partnerdivisionmapping", systemAdminRef, true);
            //    //  MigrateBulkRecords(_servicePrd, _serviceDev, "hil_integrationconfiguration", systemAdminRef,true);
            //    //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_subterritory", systemAdminRef, true);
            //    //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_integrationsource", systemAdminRef, true);
            //    //  MigrateBulkRecords(_servicePrd, _serviceDev, "hil_smstemplates", systemAdminRef, true);
            //    //MigrateBusinessMapping.MigrateData_hil_businessmapping(_servicePrd, _serviceDev);

            //    //  MigrateBusinessMapping.UpdateMigratedBusMapping(_servicePrd, _serviceDev);
            //    //  MigrateBusinessMapping.migrateBusMapping(_servicePrd, _serviceDev);
            //    //  MigrateMaterialGroups.migrateMaterialGropus(_servicePrd, _serviceDev);
            //    //  MigratePriceList.migratePriceList(_servicePrd, _serviceDev);
            //    //MigrateProducts.migrateProducts(_servicePrd, _serviceDev);
            //    //MigrateProducts.UpdateMigratedProducts(_servicePrd, _serviceDev);
            //    //  MigrateUserSetup.migrateUserSetup(_servicePrd, _serviceDev);
            //    //  TenderMasterMigration.tenderMasterMigration(_servicePrd, _serviceDev);
            //    //  HomeAdvisoryMasterMigration.homeAdvisoryMasterMigration(_servicePrd, _serviceDev);
            //    //  ServiceMasterMigration.serviceMasterMigration(_servicePrd, _serviceDev);
            //    //string[] entityList = {
            //    //"hil_campaigndivisions",
            //    //"hil_campaignenquirytypes",
            //    //"hil_campaignwebsitesetup",
            //    // "hil_cancellationreason",
            //    // "hil_caseassignmentmatrix",
            //    //"hil_casecategory",
            //    // "hil_casecategorymapping",
            //    // "hil_discountmatrix",
            //    // "hil_distributionchannel",
            //    // "hil_enquirydepartmentdivisionmapping",
            //    // "hil_enquirydocumenttype",
            //    //  "hil_enquirylostreason",
            //    //   "hil_enquiryproductsegment",
            //    // "hil_enquirysegmentdcmapping",
            //    //  "hil_joberrorcode",
            //    //"hil_jobreassignreason",
            //    //"hil_jobsquantitymatrix",
            //    //"hil_jobtat",
            //    // "hil_jobtatcategory",
            //    // "hil_mobileappbanner",
            //    //  "hil_npssetup",
            //    // "hil_partnerdepartment",
            //    //"hil_plantmaster",
            //    //"hil_plantordertypesetup",
            //    // "hil_pmsconfigurationlines",
            //    //  "hil_pmsscheduleconfiguration",
            //    // "hil_preferredlanguageforcommunication",
            //    // "hil_propertytype",
            //    //"hil_schemedistrictexclusion",
            //    //   "hil_statustransitionmatrix",
            //    //"hil_tolerencesetup",
            //    //  "hil_typeofcustomer",
            //    //  "hil_usertype",
            //    //"hil_voltage",
            //    // "hil_warrantyperiod",
            //    // "hil_whatsappproductdivisionconfig",
            //    //  "hil_wrongclosurepenalty",
            //    //"hil_amcdiscountmatrix",
            //    //"hil_jobbasecharge",
            //    //"hil_natureofcomplaint",
            //    //"hil_sbubranchmapping",
            //    //    "hil_schemeincentive",
            //    //"hil_specialincentive",
            //    //"hil_tatachievementslabmaster",
            //    //"hil_tatbreachpenalty",
            //    //  "hil_tatincentive",
            //    //"hil_tenderattachmentdoctype",
            //    //"hil_upcountrydataupdate",
            //    //"hil_usage",
            //    //"hil_warrantyvoidreason",
            //    //"account"
            //    //};
            //    //Parallel.ForEach(entityList, entityName =>
            //    //{
            //    // MigrateBulkRecords(_servicePrd, _serviceDev, entityName, systemAdminRef, true);
            //    //});
            //    // WarrantyTemplateMigration.warrantyTemplateMigration(_servicePrd, _serviceDev);
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            //}
        }

        //        static void InsertPriceListDate()
        //        {
        //            string productCode = "ACMC-ALBUS;ACMC-ALBUS 2YR;SCMC-ALBUS;SCMC-ALBUS 2YR;ACMC-AQUAS;ACMC-AQUAS-2YRS;SCMC-AQUAS;SCMC-AQUAS-2YRS;ACMC-DELITE;ACMC-DELITE 2YRS;SCMC-DELITE;SCMC-DELITE 2YRS;ACMC-DELITE ALK HR;ACMC-DELITE-ALK2HR;SCMC-DELITE ALK HR;SCMC-DELITE-ALK2HR;ACMC-DELITE ALK;ACMC-DELITE ALK-2Y;SCMC-DELITE ALK;SCMC-DELITE ALK-2Y;ACMC-DELITE DX;ACMC-DELITE DX-2YR;SCMC-DELITE DX;SCMC-DELITE DX-2YR;ACMC-DELITE APLUS;ACMC-DELITE APLUS2;SCMC-DELITE APLUS;SCMC-DELITE APLUS2;ACMC-ENTICER ALK;ACMC-ENTICER ALK 2;SCMC-ENTICER ALK;SCMC-ENTICER ALK2Y;ACMC-FEST RO+UF;ACMC-FEST RO+UF2YR;SCMC-FEST RO+UF;SCMC-FEST RO+UF2YR;ACMC-GRACIA ALK;ACMC-GRACIA ALK2YR;SCMC-GRACIA ALK;SCMC-GRACIA ALK2YR;ACMC-FAB;ACMC-FAB-2YRS;SCMC-FAB;SCMC-FAB-2YRS;ACMC-FAB ALK;ACMC-FAB ALK2;SCMC-FAB ALK;SCMC-FAB-ALK2;ACMC-MAX;ACMC-MAX-2YRS;SCMC-MAX;SCMC-MAX-2YRS;ACMC-MAXALKALINE;ACMC-MAXALKALINE-2;SCMC-MAXALKALINE;SCMC-MAXALKALINE-2;ACMC-PRO ALKALINE;ACMC-PRO ALKLINE-2;SCMC-PRO ALKALINE;SCMC-PRO ALKLINE-2;ACMC-PRO;ACMC-PRO-2YRS;SCMC-PRO;SCMC-PRO-2YRS;ACMC-PRO DX;ACMC-PRO DX-2YR;SCMC-PRO DX;SCMC-PRO DX-2YR;ACMC-UTC ALKALINE;ACMC-UTC ALKLINE-2;SCMC-UTC ALKALINE;SCMC-UTC ALKLINE-2;ACMC-UTC ADX;ACMC-UTC ADX-2YRS;SCMC-UTC ADX;SCMC-UTC ADX-2YRS;ACMC-UTC;ACMC-UTC-2YRS;SCMC-UTC;SCMC-UTC-2YRS;ACMC-25LPH;ACMC-25LPH-2YRS;SCMC-25LPH;SCMC-25LPH-2YRS;ACMC-DIGITOUCH ALK;ACMC-DIGITOUCHALK2;SCMC-DIGITOUCH ALK;SCMC-DIGITOUCHALK2;ACMC-DIGITOUCH;ACMC-DIGITOUCH-2YR;SCMC-DIGITOUCH;SCMC-DIGITOUCH-2YR;ACMC-DIGIPLUS;ACMC-DIGIPLUS-2YRS;ACMC-DISIPLUS;SCMC-DIGIPLUS;SCMC-DIGIPLUS-2YRS;ACMC-DIGIPLUS ALK;ACMC-DIGIPLUS ALK2;SCMC-DGIPLUS ALK-2;SCMC-DIGIPLUS ALK;ACMC-ACTIVE;ACMC-ACTIVE-2YRS;SCMC-ACTIVE;SCMC-ACTIVE-2YRS;ACMC-ACT PLUS BST2;ACMC-ACT PLUS BSTR;SCMC-ACT PLUS BST2;SCMC-ACT PLUS BSTR;ACMC-ACTIVE PLUS;ACMC-ACTIVE PLUS-2;SCMC-ACTIVE PLUS;SCMC-ACTIVE PLUS-2;ACMC-ACTIVE TOUCH;ACMC-ACTIVE TOUCH2;SCMC-ACTIVE TOUCH;SCMC-ACTIVE TOUCH2;ACMC-UVPLUS;ACMC-UVPLUS-2YRS;SCMC-UVPLUS;SCMC-UVPLUS-2YRS;ACMC-UVFABS;ACMC-UVFABS-2YRS;SCMC-UVFABS;SCMC-UVFABS-2YRS";
        //            string amount = "5450;9810;4000;7200;5450;9810;4000;7200;6100;10980;4675;8415;6500;12000;4550;8190;6450;11610;5000;9000;6100;10980;4675;8415;6450;11610;5000;9000;6450;11610;5000;9000;5450;9810;4000;7200;7500;13500;5500;9900;5450;9810;4000;7200;5799;10438;4350;7830;5450;9810;4000;7200;5799;10438;4350;7830;5799;10438;4350;7830;5450;9810;4000;7200;5450;9810;4000;6570;6450;11610;5000;9000;6450;11610;5000;9000;6100;10980;4675;8415;14250;25650;10500;18900;6450;11610;5000;9000;6100;10980;4675;8415;6100;10980;6100;4675;8415;6450;11610;9000;5000;2250;4050;1250;2250;5580;3100;4410;2450;2250;4050;1450;2610;3100;5580;2450;4410;3100;5580;2450;4410;3100;5580;2450;4410";
        //            string[] productCodes = productCode.Split(';');
        //            string[] Amount = amount.Split(';');
        //            Guid PriceList = new Guid("c05e7837-8fc4-ed11-b597-6045bdac5e78");
        //            Guid UomId = new Guid("0359d51b-d7cf-43b1-87f6-fc13a2c1dec8");
        //            int i = 0;
        //            foreach (string code in productCodes)
        //            {
        //                try
        //                {
        //                    Guid productId = getProductCode(code);
        //                    QueryExpression query = new QueryExpression("productpricelevel");
        //                    query.ColumnSet = new ColumnSet(false);
        //                    query.Distinct = true;
        //                    query.Criteria.AddCondition("productid", ConditionOperator.Equal, productId);
        //                    query.Criteria.AddCondition("pricelevelid", ConditionOperator.Equal, PriceList);
        //                    EntityCollection productColl = _servicePrd.RetrieveMultiple(query);

        //                    Entity pricelistItem = new Entity("productpricelevel");
        //                    pricelistItem["amount"] = new Money(Convert.ToDecimal(Amount[0]));
        //                    pricelistItem["pricelevelid"] = new EntityReference("pricelevel", PriceList);
        //                    pricelistItem["uomid"] = new EntityReference("uom", UomId);
        //                    pricelistItem["quantitysellingcode"] = new OptionSetValue(1);
        //                    pricelistItem["productid"] = new EntityReference("product", productId);
        //                    if (productColl.Entities.Count == 1)
        //                    {
        //                        pricelistItem.Id = (productColl.Entities[0].Id);
        //                        _servicePrd.Update(pricelistItem);
        //                    }
        //                    else
        //                        _servicePrd.Create(pricelistItem);

        //                }
        //                catch (Exception ex)
        //                {
        //                    Console.WriteLine(code + " || " + ex.Message);
        //                }
        //                i++;
        //                Console.WriteLine(code + " || Done ");
        //            }
        //            Console.WriteLine("Done");

        //            foreach (string code in productCodes)
        //            {
        //                SetStateRequest state = new SetStateRequest();
        //                state.State = new OptionSetValue(0);
        //                state.Status = new OptionSetValue(1);
        //                state.EntityMoniker = new EntityReference("product", getProductCode(code));
        //                SetStateResponse stateSet = (SetStateResponse)_servicePrd.Execute(state);
        //            }
        //        }
        //        static Guid getProductCode(string productCode)
        //        {
        //            QueryExpression Query = new QueryExpression("product");
        //            Query.ColumnSet = new ColumnSet(false);
        //            Query.Criteria = new FilterExpression(LogicalOperator.And);
        //            Query.Criteria.AddCondition("productnumber", ConditionOperator.Equal, productCode);
        //            EntityCollection Found = _servicePrd.RetrieveMultiple(Query);
        //            if (Found.Entities.Count != 1)
        //            {
        //                throw new Exception("Product Not Found");
        //            }
        //            else
        //            {
        //                return Found.Entities[0].Id;
        //            }
        //        }

        //        static void deleteActionCard(IOrganizationService _service)
        //        {
        //            QueryExpression query = new QueryExpression("actioncard");
        //            query.ColumnSet = new ColumnSet(true);

        //            query.AddOrder("createdon", OrderType.Ascending);
        //            query.PageInfo = new PagingInfo();
        //            query.PageInfo.Count = 5000;
        //            query.PageInfo.PageNumber = 1;
        //            query.PageInfo.ReturnTotalRecordCount = true;
        //            int count = 0;
        //            int error = 0;
        //            try
        //            {
        //                EntityCollection entCol = _service.RetrieveMultiple(query);
        //                Console.WriteLine("Record Count " + entCol.Entities.Count);
        //                do
        //                {

        //                    foreach (Entity entity1 in entCol.Entities)
        //                    {
        //                        try
        //                        {
        //                            _service.Delete(entity1.LogicalName, entity1.Id);
        //                            count++;
        //                            Console.WriteLine("Deleted " + count);
        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            Console.WriteLine("Error " + ex.Message);
        //                        }
        //                    }
        //                    query.PageInfo.PageNumber += 1;
        //                    query.PageInfo.PagingCookie = entCol.PagingCookie;
        //                    entCol = _service.RetrieveMultiple(query);
        //                }
        //                while (entCol.MoreRecords);
        //                Console.WriteLine("Count !!! " + count);
        //                Console.WriteLine("****************************************** Entity is ended.****************************************** ");
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine("ERROR !!! " + ex.Message);
        //            }
        //        }
        //        static void FSMDataMigration()
        //        {
        //            try
        //            {
        //                int i = 0;
        //                int j = 0;
        //                int pagenumber = 1;
        //                while (true)
        //                {
        //                    string fetchXML = $@"<fetch version='1.0' page='{pagenumber}' output-format='xml-platform' mapping='logical' distinct='true'>
        //                    <entity name='product'>
        //                    <attribute name='name' />
        //                    <attribute name='productnumber' />
        //                    <attribute name='description' />
        //                    <attribute name='statecode' />
        //                    <attribute name='productstructure' />
        //                    <attribute name='productid' />
        //                    <attribute name='hil_materialgroup' />
        //                    <attribute name='hil_division' />
        //                    <attribute name='hil_amount' />
        //                    <order attribute='productnumber' descending='false' />
        //                    <link-entity name='hil_servicebom' from='hil_productcategory' to='productid' link-type='inner' alias='ac' />
        //                    </entity>
        //                    </fetch>";
        //                    EntityCollection EntityList = _servicePrd.RetrieveMultiple(new FetchExpression(fetchXML));
        //                    if (EntityList.Entities.Count == 0) { break; }
        //                    j = j + EntityList.Entities.Count;
        //                    Entity ent = null;
        //                    foreach (var record in EntityList.Entities)
        //                    {
        //                        try
        //                        {
        //                            if (record.Contains("hil_materialgroup"))
        //                            {
        //                                ent = new Entity(record.LogicalName, record.Id);
        //                                ent["name"] = record.GetAttributeValue<string>("name");
        //                                ent["productnumber"] = record.GetAttributeValue<string>("productnumber");
        //                                ent["description"] = record.GetAttributeValue<string>("description");
        //                                ent["price"] = record.GetAttributeValue<Money>("hil_amount");
        //                                ent["producttypecode"] = new OptionSetValue(5);//MODEL
        //                                ent["hil_assetcategory"] = new EntityReference("msdyn_customerassetcategory", record.GetAttributeValue<EntityReference>("hil_materialgroup").Id);
        //                                ent["defaultuomscheduleid"] = new EntityReference("uomschedule", new Guid("ca3d0dcc-1332-4b87-87af-dd4d495c9fd6"));
        //                                ent["defaultuomid"] = new EntityReference("uom", new Guid("8368d3d7-1e1a-4db4-ac6a-a9ae6fac531a"));
        //                                ent["pricelevelid"] = new EntityReference("pricelevel", new Guid("3c111099-1b8c-ed11-81ac-6045bdaaae69"));
        //                                ent["quantitydecimal"] = 0;
        //                                ent["msdyn_fieldserviceproducttype"] = new OptionSetValue(690970000);
        //                                Guid recID = _serviceDev.Create(ent);

        //                                Entity updateProd = new Entity(record.LogicalName, recID);
        //                                updateProd["pricelevelid"] = new EntityReference("pricelevel", new Guid("3c111099-1b8c-ed11-81ac-6045bdaaae69"));
        //                                _serviceDev.Update(updateProd);

        //                                _serviceDev.Execute(new SetStateRequest
        //                                {
        //                                    EntityMoniker = new EntityReference(record.LogicalName, record.Id),
        //                                    State = new OptionSetValue(0), //Status
        //                                    Status = new OptionSetValue(1) //Status reason
        //                                });
        //                            }
        //                        }
        //                        catch (Exception ex) { Console.WriteLine(ex.Message); }
        //                        i += 1;
        //                        Console.WriteLine(record.GetAttributeValue<string>("productnumber") + ": " + i.ToString() + "/" + j.ToString());
        //                    }
        //                    pagenumber = pagenumber + 1;
        //                }
        //            }
        //            catch (Exception err)
        //            {
        //                Console.WriteLine(err.Message);
        //            }
        //        }
        //        static private string alphanums = "0123456789abcdefghijklmnopqrstuvwxyz";
        //        static private int codeLen = 6; //Length of coded string. Must be at least 4

        //        static public string EncodeNumber(int num)
        //        {
        //            if (num < 1 || num > 999999) //or throw an exception
        //                return "";
        //            int[] nums = new int[codeLen];
        //            int pos = 0;

        //            while (!(num == 0))
        //            {
        //                nums[pos] = num % alphanums.Length;
        //                num /= alphanums.Length;
        //                pos += 1;
        //            }

        //            string result = "";
        //            foreach (int numIndex in nums)
        //                result = alphanums[numIndex].ToString() + result;

        //            return result;
        //        }
        //        static void DataMigrationNew()
        //        {
        //            int count = 1;
        //            try
        //            {
        //                string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
        //                                      <entity name='product'>
        //	                                      <all-attributes />
        //                                        <order attribute='productnumber' descending='false' />
        //                                        <link-entity name='hil_servicebom' from='hil_product' to='productid' link-type='inner' alias='ac'>
        //                                          <link-entity name='product' from='productid' to='hil_productcategory' link-type='inner' alias='ad'>
        //                                            <filter type='and'>
        //                                              <condition attribute='hil_division' operator='eq' uiname='HAVELLS AQUA' uitype='product' value='{72981D83-16FA-E811-A94C-000D3AF0694E}' />
        //                                            </filter>
        //                                          </link-entity>
        //                                        </link-entity>
        //                                      </entity>
        //                                    </fetch>";
        //                EntityCollection entCol = _servicePrd.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                Console.WriteLine("Record Count " + entCol.Entities.Count);
        //                //do
        //                //{
        //                foreach (Entity entity in entCol.Entities)
        //                {
        //                    try
        //                    {
        //                        Guid uomId = getBaseUoM(entity.GetAttributeValue<EntityReference>("defaultuomid").Id);

        //                        bool conmt = false;
        //                        var attributes = entity.Attributes.Keys;
        //                        Entity _createEntity = new Entity(entity.LogicalName, entity.Id);
        //                        foreach (string name in attributes)
        //                        {
        //                            if (name != "modifiedby" && name != "createdby")
        //                            {
        //                                if (entity[name].GetType().Name == "EntityReference")
        //                                {
        //                                    if (entity.GetAttributeValue<EntityReference>(name).LogicalName != "systemuser")
        //                                    {
        //                                        //Console.WriteLine(name);
        //                                        Guid productId = entity.GetAttributeValue<EntityReference>(name).Id;
        //                                        if (productId != entity.Id)
        //                                        {
        //                                            Guid prdId = getRecord(entity.GetAttributeValue<EntityReference>(name).LogicalName, productId);
        //                                            if (prdId != Guid.Empty)
        //                                                _createEntity[name] = entity[name];
        //                                            else
        //                                                conmt = true;
        //                                        }
        //                                    }
        //                                    else if (entity.GetAttributeValue<EntityReference>(name).LogicalName == "systemuser")
        //                                    {
        //                                        Guid userid = RetriveSystemUser(_servicePrd, _serviceDev, entity.GetAttributeValue<EntityReference>(name).Id);
        //                                        if (userid == Guid.Empty)
        //                                            userid = SystemAdmin.Id;
        //                                        _createEntity[name] = new EntityReference("systemuser", userid);
        //                                    }
        //                                    else
        //                                        _createEntity[name] = entity[name];
        //                                }
        //                                else
        //                                    _createEntity[name] = entity[name];
        //                            }
        //                        }
        //                        if (conmt)
        //                            continue;
        //                        _createEntity["statecode"] = new OptionSetValue(2);
        //                        _createEntity["statuscode"] = new OptionSetValue(0);
        //                        _createEntity["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
        //                        _createEntity["defaultuomscheduleid"] = new EntityReference("uomschedule", getBaseUoMSchedule(uomId));
        //                        _createEntity["defaultuomid"] = new EntityReference("uom", uomId);
        //                        _serviceDev.Create(_createEntity);
        //                        count++;
        //                        Console.WriteLine("Product With Name " + _createEntity.GetAttributeValue<string>("name") + " is created");
        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        Console.WriteLine("Error " + ex.Message);
        //                    }
        //                }
        //                //query.PageInfo.PageNumber += 1;
        //                //query.PageInfo.PagingCookie = entCol.PagingCookie;
        //                //entCol = _servicePrd.RetrieveMultiple(query);
        //                //}
        //                //while (entCol.MoreRecords);
        //                Console.WriteLine("Count !!! " + count);
        //                Console.WriteLine("****************************************** Entity is ended.****************************************** ");
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine("ERROR !!! " + ex.Message);
        //            }
        //        }
        //        static Guid getRecord(string entityLogicalName, Guid productid)
        //        {
        //            bool conmt = false;
        //            try
        //            {
        //                QueryExpression query = new QueryExpression(entityLogicalName);
        //                query.ColumnSet = new ColumnSet(true);
        //                query.Criteria.AddCondition(entityLogicalName + "id", ConditionOperator.Equal, productid);
        //                EntityCollection entityCollection = _serviceDev.RetrieveMultiple(query);
        //                if (entityCollection.Entities.Count == 1)
        //                {
        //                    return entityCollection[0].Id;
        //                }
        //                else
        //                {
        //                    Entity entity = _servicePrd.Retrieve(entityLogicalName, productid, new ColumnSet(true));


        //                    var attributes = entity.Attributes.Keys;
        //                    Entity _createEntity = new Entity(entity.LogicalName, entity.Id);
        //                    bool conti = false;
        //                    foreach (string name in attributes)
        //                    {
        //                        if (name != "modifiedby" && name != "createdby")
        //                        {
        //                            if (entity[name].GetType().Name == "EntityReference")
        //                            {
        //                                if (entity.GetAttributeValue<EntityReference>(name).LogicalName != "systemuser")
        //                                {
        //                                    //Console.WriteLine(name);
        //                                    Guid productId = entity.GetAttributeValue<EntityReference>(name).Id;
        //                                    if (productId != entity.Id)
        //                                    {
        //                                        Guid prdId = getRecord(entity.LogicalName, productId);
        //                                        if (prdId != Guid.Empty)
        //                                            _createEntity[name] = entity[name];
        //                                        else
        //                                            conmt = true;
        //                                    }
        //                                }
        //                                else if (entity.GetAttributeValue<EntityReference>(name).LogicalName == "systemuser")
        //                                {
        //                                    Guid userid = RetriveSystemUser(_servicePrd, _serviceDev, entity.GetAttributeValue<EntityReference>(name).Id);
        //                                    if (userid == Guid.Empty)
        //                                        userid = SystemAdmin.Id;
        //                                    _createEntity[name] = new EntityReference("systemuser", userid);
        //                                }
        //                                else
        //                                    _createEntity[name] = entity[name];
        //                            }
        //                            else
        //                                _createEntity[name] = entity[name];
        //                        }
        //                    }
        //                    if (!conti)
        //                    {

        //                        _createEntity["statecode"] = new OptionSetValue(2);
        //                        _createEntity["statuscode"] = new OptionSetValue(0);
        //                        _createEntity["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
        //                        if (entity.Contains("defaultuomid"))
        //                        {
        //                            Guid uomId = getBaseUoM(entity.GetAttributeValue<EntityReference>("defaultuomid").Id);
        //                            _createEntity["defaultuomscheduleid"] = new EntityReference("uomschedule", getBaseUoMSchedule(uomId));
        //                            _createEntity["defaultuomid"] = new EntityReference("uom", uomId);
        //                        }
        //                        Guid recordid = _serviceDev.Create(_createEntity);
        //                        Console.WriteLine("Product With Name " + _createEntity.GetAttributeValue<string>("name") + " is created");
        //                        return recordid;
        //                    }
        //                    else
        //                        return Guid.Empty;
        //                }

        //            }
        //            catch (Exception ex)
        //            {
        //                Console.Write("___________Error : " + ex.Message);
        //                return Guid.Empty;
        //            }
        //        }
        //        static void MigrateBussGeoMaping()
        //        {
        //            string entityName = "hil_businessmapping";
        //            Console.WriteLine("****************************************** Entity " + entityName + " is started.****************************************** ");
        //            QueryExpression query = new QueryExpression(entityName.ToLower());
        //            query.ColumnSet = new ColumnSet(true);
        //            query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 1);
        //            query.Criteria.AddCondition("hil_state", ConditionOperator.Equal, new Guid("27fcb4df-bbf7-e811-a94c-000d3af06091"));

        //            query.AddOrder("createdon", OrderType.Ascending);
        //            query.PageInfo = new PagingInfo();
        //            query.PageInfo.Count = 5000;
        //            query.PageInfo.PageNumber = 1;
        //            query.PageInfo.ReturnTotalRecordCount = true;
        //            int count = 0;
        //            int error = 0;
        //            try
        //            {
        //                EntityCollection entCol = _servicePrd.RetrieveMultiple(query);
        //                Console.WriteLine("Record Count " + entCol.Entities.Count);
        //                do
        //                {

        //                    foreach (Entity entity1 in entCol.Entities)
        //                    {
        //                    Cont:
        //                        Entity _createEntity = new Entity(entity1.LogicalName, entity1.Id);
        //                        try
        //                        {
        //                            foreach (string name in entity1.Attributes.Keys)
        //                            {
        //                                if (name != "modifiedby" && name != "createdby")
        //                                {
        //                                    if (entity1[name].GetType().Name == "EntityReference")
        //                                    {
        //                                        if (entity1.GetAttributeValue<EntityReference>(name).LogicalName == "systemuser")
        //                                        {
        //                                            Console.WriteLine(name);
        //                                            _createEntity[name] = new EntityReference("systemuser", RetriveSystemUser(_servicePrd, _serviceDev, entity1.GetAttributeValue<EntityReference>(name).Id));
        //                                        }
        //                                        else if (entity1.GetAttributeValue<EntityReference>(name).LogicalName == "uom")
        //                                        {
        //                                            _createEntity[name] = new EntityReference("uom", getBaseUoM(entity1.GetAttributeValue<EntityReference>(name).Id));
        //                                        }
        //                                        else if (entity1.GetAttributeValue<EntityReference>(name).LogicalName == "transactioncurrency")
        //                                        {
        //                                            _createEntity[name] = new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
        //                                        }
        //                                        else
        //                                        {
        //                                            _createEntity[name] = new EntityReference(entity1.GetAttributeValue<EntityReference>(name).LogicalName,
        //                                                CreateRecordIfNotExist(entity1.GetAttributeValue<EntityReference>(name)));
        //                                        }
        //                                    }
        //                                    else
        //                                    {
        //                                        _createEntity[name] = entity1[name];
        //                                    }
        //                                }
        //                                else
        //                                {
        //                                    Console.WriteLine("exc " + name);
        //                                }
        //                            }

        //                            var attributes = entity1;
        //                            _createEntity["createdby"] = SystemAdmin;
        //                            if (entity1.Contains("ownerid"))
        //                                _createEntity["ownerid"] = SystemAdmin;
        //                            _createEntity["modifiedby"] = SystemAdmin;
        //                            _serviceDev.Create(_createEntity);
        //                            count++;
        //                            Console.WriteLine("Done " + count + " Record Created On " + entity1["createdon"]);
        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            if (ex.Message != "Cannot insert duplicate key.")
        //                            {
        //                                Console.WriteLine("Error " + ex.Message);
        //                                if (ex.Message.Contains("Does Not Exist"))
        //                                {
        //                                    string[] arr = ex.Message.Split(' ');
        //                                    string entityNameRef = arr[0];
        //                                    string entityIDRef = arr[4];
        //                                    migrateData(entityNameRef, entityIDRef);
        //                                    goto Cont;
        //                                }
        //                                error++;
        //                                Console.WriteLine("Error Count " + error);
        //                            }

        //                            Console.WriteLine("Error " + ex.Message);
        //                        }
        //                    }
        //                    query.PageInfo.PageNumber += 1;
        //                    query.PageInfo.PagingCookie = entCol.PagingCookie;
        //                    entCol = _servicePrd.RetrieveMultiple(query);
        //                }
        //                while (entCol.MoreRecords);
        //                Console.WriteLine("Count !!! " + count);
        //                Console.WriteLine("****************************************** Entity " + entityName + " is ended.****************************************** ");
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine("ERROR !!! " + ex.Message);
        //            }
        //        }
        //        static Guid CreateRecordIfNotExist(EntityReference entityReference)
        //        {
        //        Cont1:
        //            try
        //            {
        //                QueryExpression query = new QueryExpression(entityReference.LogicalName);
        //                query.ColumnSet = new ColumnSet(false);
        //                query.Criteria.AddCondition(entityReference.LogicalName + "id", ConditionOperator.Equal, entityReference.Id);
        //                EntityCollection entityCollection = _serviceDev.RetrieveMultiple(query);
        //                if (entityCollection.Entities.Count != 1)
        //                {
        //                    Entity entity = _servicePrd.Retrieve(entityReference.LogicalName, entityReference.Id, new ColumnSet(true));
        //                    _serviceDev.Create(entity);
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine(ex.Message);
        //                if (ex.Message != "Cannot insert duplicate key.")
        //                {
        //                    Console.WriteLine("Error " + ex.Message);
        //                    if (ex.Message.Contains("Does Not Exist"))
        //                    {
        //                        string[] arr = ex.Message.Split(' ');
        //                        string entityNameRef = arr[1];
        //                        string entityIDRef = arr[5];
        //                        if (!ex.Message.Contains("Product"))
        //                        {
        //                            migrateData(entityNameRef, entityIDRef);
        //                            goto Cont1;
        //                        }
        //                    }
        //                }
        //            }
        //            return entityReference.Id;
        //        }

        //        static void ProductDataMigration()
        //        {
        //            string entity = "product";
        //            //SBUProductDataMigration(entity);
        //            //DivisionProductDataMigration(entity);
        //            //MGProductDataMigration(entity);
        //            //MaterialProductDataMigration(entity);
        //            PartProductDataMigration(entity);
        //        }
        //        static void MaterialProductDataMigration(string entityName)
        //        {

        //            //QueryExpression query = new QueryExpression(entityName);
        //            //query.ColumnSet = new ColumnSet(true);
        //            //query.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 5);
        //            //query.AddOrder("createdon", OrderType.Ascending);
        //            //query.PageInfo = new PagingInfo();
        //            //query.PageInfo.Count = 5000;
        //            //query.PageInfo.PageNumber = 1;
        //            //query.PageInfo.ReturnTotalRecordCount = true;
        //            int count = 0;
        //            int error = 0;
        //            try
        //            {
        //                string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
        //                  <entity name='product'>
        //                    <all-attributes />
        //                    <order attribute='productnumber' descending='false' />
        //                    <link-entity name='hil_servicebom' from='hil_productcategory' to='productid' link-type='inner' alias='as'>
        //                      <link-entity name='product' from='productid' to='hil_productcategory' link-type='inner' alias='at'>
        //                        <filter type='and'>
        //                          <condition attribute='hil_division' operator='eq' uiname='HAVELLS AQUA' uitype='product' value='{72981D83-16FA-E811-A94C-000D3AF0694E}' />
        //                        </filter>
        //                      </link-entity>
        //                    </link-entity>
        //                  </entity>
        //                </fetch>";
        //                EntityCollection entCol = _servicePrd.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                Console.WriteLine("Record Count " + entCol.Entities.Count);
        //                //do
        //                //{
        //                foreach (Entity entity in entCol.Entities)
        //                {
        //                    try
        //                    {
        //                        Guid uomId = getBaseUoM(entity.GetAttributeValue<EntityReference>("defaultuomid").Id);

        //                        bool conmt = false;
        //                        var attributes = entity.Attributes.Keys;
        //                        Entity _createEntity = new Entity(entity.LogicalName, entity.Id);
        //                        foreach (string name in attributes)
        //                        {
        //                            if (name != "modifiedby" && name != "createdby")
        //                            {
        //                                if (entity[name].GetType().Name == "EntityReference")
        //                                {
        //                                    if (entity.GetAttributeValue<EntityReference>(name).LogicalName == "product")
        //                                    {
        //                                        //Console.WriteLine(name);
        //                                        Guid productId = entity.GetAttributeValue<EntityReference>(name).Id;
        //                                        if (productId != entity.Id)
        //                                        {
        //                                            Guid prdId = getPrdouct(productId);
        //                                            if (prdId != Guid.Empty)
        //                                                _createEntity[name] = entity[name];
        //                                            else
        //                                                conmt = true;
        //                                        }
        //                                    }
        //                                    else
        //                                        _createEntity[name] = entity[name];
        //                                }
        //                                else
        //                                    _createEntity[name] = entity[name];
        //                            }
        //                        }
        //                        if (conmt)
        //                            continue;
        //                        _createEntity["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
        //                        _createEntity["defaultuomscheduleid"] = new EntityReference("uomschedule", getBaseUoMSchedule(uomId));
        //                        _createEntity["defaultuomid"] = new EntityReference("uom", uomId);
        //                        _serviceDev.Create(_createEntity);
        //                        count++;
        //                        Console.WriteLine("Product With Name " + _createEntity.GetAttributeValue<string>("name") + " is created");
        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        Console.WriteLine("Error " + ex.Message);
        //                    }
        //                }
        //                //query.PageInfo.PageNumber += 1;
        //                //query.PageInfo.PagingCookie = entCol.PagingCookie;
        //                //entCol = _servicePrd.RetrieveMultiple(query);
        //                //}
        //                //while (entCol.MoreRecords);
        //                Console.WriteLine("Count !!! " + count);
        //                Console.WriteLine("****************************************** Entity " + entityName + " is ended.****************************************** ");
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine("ERROR !!! " + ex.Message);
        //            }

        //        }
        //        static void PartProductDataMigration(string entityName)
        //        {

        //            //QueryExpression query = new QueryExpression(entityName);
        //            //query.ColumnSet = new ColumnSet(true);
        //            //query.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 5);
        //            //query.AddOrder("createdon", OrderType.Ascending);
        //            //query.PageInfo = new PagingInfo();
        //            //query.PageInfo.Count = 5000;
        //            //query.PageInfo.PageNumber = 1;
        //            //query.PageInfo.ReturnTotalRecordCount = true;
        //            int count = 0;
        //            int error = 0;
        //            try
        //            {
        //                string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
        //  <entity name='product'>
        //	  <all-attributes />
        //    <order attribute='productnumber' descending='false' />
        //    <link-entity name='hil_natureofcomplaint' from='hil_relatedproduct' to='productid' link-type='inner' alias='ao'>
        //      <link-entity name='product' from='productid' to='hil_relatedproduct' link-type='inner' alias='ap'>
        //        <filter type='and'>
        //          <condition attribute='hil_division' operator='eq' uiname='HAVELLS AQUA' uitype='product' value='{72981D83-16FA-E811-A94C-000D3AF0694E}' />
        //        </filter>
        //      </link-entity>
        //    </link-entity>
        //  </entity>
        //</fetch>";
        //                EntityCollection entCol = _servicePrd.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                Console.WriteLine("Record Count " + entCol.Entities.Count);
        //                //do
        //                //{
        //                foreach (Entity entity in entCol.Entities)
        //                {
        //                    try
        //                    {
        //                        Guid uomId = getBaseUoM(entity.GetAttributeValue<EntityReference>("defaultuomid").Id);

        //                        bool conmt = false;
        //                        var attributes = entity.Attributes.Keys;
        //                        Entity _createEntity = new Entity(entity.LogicalName, entity.Id);
        //                        foreach (string name in attributes)
        //                        {
        //                            if (name != "modifiedby" && name != "createdby")
        //                            {
        //                                if (entity[name].GetType().Name == "EntityReference")
        //                                {
        //                                    if (entity.GetAttributeValue<EntityReference>(name).LogicalName == "product")
        //                                    {
        //                                        //Console.WriteLine(name);
        //                                        Guid productId = entity.GetAttributeValue<EntityReference>(name).Id;
        //                                        if (productId != entity.Id)
        //                                        {
        //                                            Guid prdId = getPrdouct(productId);
        //                                            if (prdId != Guid.Empty)
        //                                                _createEntity[name] = entity[name];
        //                                            else
        //                                                conmt = true;
        //                                        }
        //                                    }
        //                                    else
        //                                        _createEntity[name] = entity[name];
        //                                }
        //                                else
        //                                    _createEntity[name] = entity[name];
        //                            }
        //                        }
        //                        if (conmt)
        //                            continue;
        //                        _createEntity["statecode"] = new OptionSetValue(2);
        //                        _createEntity["statuscode"] = new OptionSetValue(0);
        //                        _createEntity["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
        //                        _createEntity["defaultuomscheduleid"] = new EntityReference("uomschedule", getBaseUoMSchedule(uomId));
        //                        _createEntity["defaultuomid"] = new EntityReference("uom", uomId);
        //                        _serviceDev.Create(_createEntity);
        //                        count++;
        //                        Console.WriteLine("Product With Name " + _createEntity.GetAttributeValue<string>("name") + " is created");
        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        Console.WriteLine("Error " + ex.Message);
        //                    }
        //                }
        //                //query.PageInfo.PageNumber += 1;
        //                //query.PageInfo.PagingCookie = entCol.PagingCookie;
        //                //entCol = _servicePrd.RetrieveMultiple(query);
        //                //}
        //                //while (entCol.MoreRecords);
        //                Console.WriteLine("Count !!! " + count);
        //                Console.WriteLine("****************************************** Entity " + entityName + " is ended.****************************************** ");
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine("ERROR !!! " + ex.Message);
        //            }

        //        }
        //        static Guid getPrdouct(Guid productid)
        //        {
        //            try
        //            {
        //                QueryExpression query = new QueryExpression("product");
        //                query.ColumnSet = new ColumnSet(true);
        //                query.Criteria.AddCondition("productid", ConditionOperator.Equal, productid);
        //                EntityCollection entityCollection = _serviceDev.RetrieveMultiple(query);
        //                if (entityCollection.Entities.Count == 1)
        //                {
        //                    return entityCollection[0].Id;
        //                }
        //                else
        //                {
        //                    Entity product = _servicePrd.Retrieve("product", productid, new ColumnSet(true));


        //                    var attributes = product.Attributes.Keys;
        //                    Entity _createEntity = new Entity(product.LogicalName, product.Id);
        //                    bool conti = false;
        //                    foreach (string name in attributes)
        //                    {
        //                        if (name != "modifiedby" && name != "createdby")
        //                        {
        //                            if (product[name].GetType().Name == "EntityReference")
        //                            {
        //                                if (product.GetAttributeValue<EntityReference>(name).LogicalName == "product")
        //                                {
        //                                    //Console.WriteLine(name);
        //                                    Guid productId = product.GetAttributeValue<EntityReference>(name).Id;
        //                                    if (productId != product.Id)
        //                                    {
        //                                        Guid prdId = getPrdouct(productId);
        //                                        if (prdId != Guid.Empty)
        //                                            _createEntity[name] = product[name];
        //                                        else
        //                                            conti = true;
        //                                    }
        //                                }
        //                                else
        //                                    _createEntity[name] = product[name];
        //                            }
        //                            else
        //                                _createEntity[name] = product[name];
        //                        }
        //                    }
        //                    if (!conti)
        //                    {

        //                        _createEntity["statecode"] = new OptionSetValue(2);
        //                        _createEntity["statuscode"] = new OptionSetValue(0);
        //                        _createEntity["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
        //                        if (product.Contains("defaultuomid"))
        //                        {
        //                            Guid uomId = getBaseUoM(product.GetAttributeValue<EntityReference>("defaultuomid").Id);
        //                            _createEntity["defaultuomscheduleid"] = new EntityReference("uomschedule", getBaseUoMSchedule(uomId));
        //                            _createEntity["defaultuomid"] = new EntityReference("uom", uomId);
        //                        }
        //                        Guid recordid = _serviceDev.Create(_createEntity);
        //                        Console.WriteLine("Product With Name " + _createEntity.GetAttributeValue<string>("name") + " is created");
        //                        return recordid;
        //                    }
        //                    else
        //                        return Guid.Empty;
        //                }

        //            }
        //            catch (Exception ex)
        //            {
        //                Console.Write("___________Error : " + ex.Message);
        //                return Guid.Empty;
        //            }
        //        }
        //        static void MGProductDataMigration(string entityName)
        //        {

        //            QueryExpression query = new QueryExpression(entityName);
        //            query.ColumnSet = new ColumnSet(true);
        //            query.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 3);
        //            query.AddOrder("createdon", OrderType.Ascending);
        //            query.PageInfo = new PagingInfo();
        //            query.PageInfo.Count = 5000;
        //            query.PageInfo.PageNumber = 1;
        //            query.PageInfo.ReturnTotalRecordCount = true;
        //            int count = 0;
        //            int error = 0;
        //            try
        //            {
        //                EntityCollection entCol = _servicePrd.RetrieveMultiple(query);
        //                Console.WriteLine("Record Count " + entCol.Entities.Count);
        //                do
        //                {

        //                    foreach (Entity entity in entCol.Entities)
        //                    {
        //                    Cont:
        //                        try
        //                        {
        //                            Guid uomId = getBaseUoM(entity.GetAttributeValue<EntityReference>("defaultuomid").Id);
        //                            entity["defaultuomscheduleid"] = new EntityReference("uomschedule", getBaseUoMSchedule(uomId));
        //                            entity["defaultuomid"] = new EntityReference("uom", uomId);
        //                            entity["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
        //                            _serviceDev.Create(entity);
        //                            count++;
        //                            Console.WriteLine("Done " + count + " Record Created On " + entity["createdon"]);

        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            Console.WriteLine("Error " + ex.Message);
        //                        }
        //                    }
        //                    query.PageInfo.PageNumber += 1;
        //                    query.PageInfo.PagingCookie = entCol.PagingCookie;
        //                    entCol = _servicePrd.RetrieveMultiple(query);
        //                }
        //                while (entCol.MoreRecords);
        //                Console.WriteLine("Count !!! " + count);
        //                Console.WriteLine("****************************************** Entity " + entityName + " is ended.****************************************** ");
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine("ERROR !!! " + ex.Message);
        //            }

        //        }
        //        static void DivisionProductDataMigration(string entityName)
        //        {

        //            QueryExpression query = new QueryExpression(entityName);
        //            query.ColumnSet = new ColumnSet(true);
        //            query.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 2);
        //            query.AddOrder("createdon", OrderType.Ascending);
        //            query.PageInfo = new PagingInfo();
        //            query.PageInfo.Count = 5000;
        //            query.PageInfo.PageNumber = 1;
        //            query.PageInfo.ReturnTotalRecordCount = true;
        //            int count = 0;
        //            int error = 0;
        //            try
        //            {
        //                EntityCollection entCol = _servicePrd.RetrieveMultiple(query);
        //                Console.WriteLine("Record Count " + entCol.Entities.Count);
        //                do
        //                {

        //                    foreach (Entity entity in entCol.Entities)
        //                    {
        //                    Cont:
        //                        try
        //                        {
        //                            Guid uomId = getBaseUoM(entity.GetAttributeValue<EntityReference>("defaultuomid").Id);
        //                            entity["defaultuomscheduleid"] = new EntityReference("uomschedule", getBaseUoMSchedule(uomId));
        //                            entity["defaultuomid"] = new EntityReference("uom", uomId);
        //                            entity["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
        //                            _serviceDev.Create(entity);
        //                            count++;
        //                            Console.WriteLine("Done " + count + " Record Created On " + entity["createdon"]);

        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            Console.WriteLine("Error " + ex.Message);
        //                        }
        //                    }
        //                    query.PageInfo.PageNumber += 1;
        //                    query.PageInfo.PagingCookie = entCol.PagingCookie;
        //                    entCol = _servicePrd.RetrieveMultiple(query);
        //                }
        //                while (entCol.MoreRecords);
        //                Console.WriteLine("Count !!! " + count);
        //                Console.WriteLine("****************************************** Entity " + entityName + " is ended.****************************************** ");
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine("ERROR !!! " + ex.Message);
        //            }

        //        }
        //        static void SBUProductDataMigration(string entityName)
        //        {
        //            #region SBU
        //            QueryExpression query = new QueryExpression(entityName);
        //            query.ColumnSet = new ColumnSet(true);
        //            query.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 1);
        //            query.AddOrder("createdon", OrderType.Ascending);
        //            query.PageInfo = new PagingInfo();
        //            query.PageInfo.Count = 5000;
        //            query.PageInfo.PageNumber = 1;
        //            query.PageInfo.ReturnTotalRecordCount = true;
        //            int count = 0;
        //            int error = 0;
        //            try
        //            {
        //                EntityCollection entCol = _servicePrd.RetrieveMultiple(query);
        //                Console.WriteLine("Record Count " + entCol.Entities.Count);
        //                do
        //                {

        //                    foreach (Entity entity in entCol.Entities)
        //                    {
        //                    Cont:
        //                        try
        //                        {
        //                            Guid uomId = getBaseUoM(entity.GetAttributeValue<EntityReference>("defaultuomid").Id);
        //                            entity["defaultuomscheduleid"] = new EntityReference("uomschedule", getBaseUoMSchedule(uomId));
        //                            entity["defaultuomid"] = new EntityReference("uom", uomId);
        //                            entity["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
        //                            _serviceDev.Create(entity);
        //                            count++;
        //                            Console.WriteLine("Done " + count + " Record Created On " + entity["createdon"]);

        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            Console.WriteLine("Error " + ex.Message);
        //                        }
        //                    }
        //                    query.PageInfo.PageNumber += 1;
        //                    query.PageInfo.PagingCookie = entCol.PagingCookie;
        //                    entCol = _servicePrd.RetrieveMultiple(query);
        //                }
        //                while (entCol.MoreRecords);
        //                Console.WriteLine("Count !!! " + count);
        //                Console.WriteLine("****************************************** Entity " + entityName + " is ended.****************************************** ");
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine("ERROR !!! " + ex.Message);
        //            }
        //            #endregion
        //        }
        //        static void deleteData(string entityName)
        //        {
        //            QueryExpression query = new QueryExpression(entityName.ToLower());
        //            query.ColumnSet = new ColumnSet(true);
        //            query.AddOrder("createdon", OrderType.Ascending);
        //            query.PageInfo = new PagingInfo();
        //            query.PageInfo.Count = 5000;
        //            query.PageInfo.PageNumber = 1;
        //            query.PageInfo.ReturnTotalRecordCount = true;
        //            EntityCollection entCol = _serviceDev.RetrieveMultiple(query);
        //            Console.WriteLine("Record Count " + entCol.Entities.Count);
        //            do
        //            {

        //                foreach (Entity entity1 in entCol.Entities)
        //                {
        //                    try
        //                    {
        //                        Console.WriteLine(entity1.GetAttributeValue<string>("name"));
        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        Console.WriteLine(ex.Message);
        //                    }
        //                }
        //                query.PageInfo.PageNumber += 1;
        //                query.PageInfo.PagingCookie = entCol.PagingCookie;
        //                entCol = _servicePrd.RetrieveMultiple(query);
        //            }
        //            while (entCol.MoreRecords);
        //        }
        //        static void migrate(string entityName, string ID)
        //        {
        //            int count = 0;
        //            int error = 0;
        //            try
        //            {
        //                string fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //  <entity name='hil_pmsconfiguration'>
        //    <all-attributes />
        //    <order attribute='hil_name' descending='false' />
        //    <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='ap'>
        //      <link-entity name='product' from='productid' to='hil_product' link-type='inner' alias='aq'>
        //        <filter type='and'>
        //          <condition attribute='hil_division' operator='eq' uiname='HAVELLS AQUA' uitype='product' value='{72981D83-16FA-E811-A94C-000D3AF0694E}' />
        //        </filter>
        //      </link-entity>
        //    </link-entity>
        //  </entity>
        //</fetch>";
        //                EntityCollection entCol = _servicePrd.RetrieveMultiple(new FetchExpression(fetch));
        //                Console.WriteLine("Record Count " + entCol.Entities.Count);
        //                do
        //                {

        //                    foreach (Entity entity1 in entCol.Entities)
        //                    {
        //                        Guid UomId = Guid.Empty;
        //                    Cont:
        //                        Entity _createEntity = new Entity(entity1.LogicalName, entity1.Id);
        //                        try
        //                        {
        //                            var attributes = entity1.Attributes.Keys;

        //                            foreach (string name in attributes)
        //                            {
        //                                // if (name != "adx_partnervisible")
        //                                if (name != "modifiedby" && name != "createdby" && name != "ownerid" && name != "owninguser" && name != "statecode" && name != "statuscode")
        //                                {
        //                                    if (entity1[name].GetType().Name == "EntityReference")
        //                                    {
        //                                        if (entity1.GetAttributeValue<EntityReference>(name).LogicalName == "systemuser")
        //                                        {
        //                                            Console.WriteLine(name);
        //                                            Guid userid = RetriveSystemUser(_servicePrd, _serviceDev, entity1.GetAttributeValue<EntityReference>(name).Id);
        //                                            if (userid == Guid.Empty)
        //                                                userid = SystemAdmin.Id;
        //                                            _createEntity[name] = new EntityReference("systemuser", userid);
        //                                        }
        //                                        else if (entity1.GetAttributeValue<EntityReference>(name).LogicalName == "uom")
        //                                        {
        //                                            UomId = getBaseUoM(entity1.GetAttributeValue<EntityReference>(name).Id);
        //                                            _createEntity[name] = new EntityReference("uom", UomId);
        //                                        }
        //                                        else if (entity1.GetAttributeValue<EntityReference>(name).LogicalName.ToUpper() == "uomschedule".ToUpper())
        //                                        {
        //                                            if (UomId != Guid.Empty)
        //                                                _createEntity[name] = new EntityReference("uomschedule", getBaseUoMSchedule(UomId));
        //                                        }
        //                                        else if (entity1.GetAttributeValue<EntityReference>(name).LogicalName == "transactioncurrency")
        //                                        {
        //                                            _createEntity[name] = new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
        //                                        }
        //                                        else
        //                                        {
        //                                            _createEntity[name] = entity1[name];
        //                                        }
        //                                    }
        //                                    else
        //                                    {
        //                                        _createEntity[name] = entity1[name];
        //                                    }
        //                                }
        //                                else
        //                                {
        //                                    //Console.WriteLine("exc " + name);
        //                                }
        //                            }
        //                            //_createEntity["hil_name"] = "New To";
        //                            //_createEntity["createdby"] = SystemAdmin;

        //                            // _createEntity["isdefaulttheme"] = false;
        //                            if (entity1.Contains("ownerid"))
        //                                _createEntity["ownerid"] = SystemAdmin;
        //                            //_createEntity["modifiedby"] = SystemAdmin;
        //                            //string nameUoM = entity1.Contains("name") ? entity1.GetAttributeValue<string>("name") : "";
        //                            //if (nameUoM != "Primary Unit" && nameUoM != "Meter" && nameUoM != "Foot" && nameUoM != "Yard" && nameUoM != "Kilometer" && nameUoM != "Mile" && nameUoM != "Hour" && nameUoM != "Meter" && nameUoM != "Unit")
        //                            _serviceDev.Create(_createEntity);
        //                            count++;
        //                            Console.WriteLine("@@@@@@@@@@@@@@@@@@@@Done " + count + " Record Created On " + entity1["createdon"]);
        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            if (ex.Message != "Cannot insert duplicate key.")
        //                            {
        //                                Console.WriteLine("Error " + ex.Message);
        //                                if (ex.Message.Contains("Does Not Exist"))
        //                                {
        //                                    string[] arr = ex.Message.Split(' ');
        //                                    string entityNameRef = arr[0];
        //                                    string entityIDRef = arr[4];
        //                                    if (!ex.Message.Contains("Product"))
        //                                    {
        //                                        migrateData(entityNameRef, entityIDRef);
        //                                        goto Cont;
        //                                    }
        //                                }
        //                                error++;
        //                                Console.WriteLine("Error Count " + error);
        //                            }

        //                            Console.WriteLine("Error " + ex.Message);
        //                        }
        //                    }
        //                    //query.PageInfo.PageNumber += 1;
        //                    //query.PageInfo.PagingCookie = entCol.PagingCookie;
        //                    //entCol = _servicePrd.RetrieveMultiple(query);
        //                }
        //                while (entCol.MoreRecords);
        //                Console.WriteLine("Count !!! " + count);
        //                Console.WriteLine("****************************************** Entity " + entityName + " is ended.****************************************** ");
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine("ERROR !!! " + ex.Message);
        //            }
        //        }
        //        static void migrateNoc()
        //        {
        //            string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //  <entity name='hil_stagingdivisonmaterialgroupmapping'>
        //   <all-attributes />
        //    <order attribute='hil_name' descending='false' />
        //  </entity>
        //</fetch>";
        //            int count = 1;
        //            EntityCollection entCol = _servicePrd.RetrieveMultiple(new FetchExpression(_fetchXML));
        //            Console.WriteLine("Record Count " + entCol.Entities.Count);
        //            //do
        //            //{
        //            foreach (Entity entity in entCol.Entities)
        //            {
        //            Cont123:
        //                try
        //                {
        //                    //    if(entity.Contains("msdyn_unit"))
        //                    //    Guid uomId = getBaseUoM(entity.GetAttributeValue<EntityReference>("msdyn_unit").Id);

        //                    bool conmt = false;
        //                    var attributes = entity.Attributes.Keys;
        //                    Entity _createEntity = new Entity(entity.LogicalName, entity.Id);
        //                    foreach (string name in attributes)
        //                    {
        //                        if (name != "modifiedby" && name != "createdby" && name != "ownerid" && name != "owninguser" && name != "statecode" && name != "statuscode")
        //                        {
        //                            if (entity[name].GetType().Name == "EntityReference")
        //                            {
        //                                if (entity.GetAttributeValue<EntityReference>(name).LogicalName == "product")
        //                                {
        //                                    //Console.WriteLine(name);
        //                                    Guid productId = entity.GetAttributeValue<EntityReference>(name).Id;
        //                                    if (productId != entity.Id)
        //                                    {
        //                                        Guid prdId = getPrdouct(productId);
        //                                        if (prdId != Guid.Empty)
        //                                            _createEntity[name] = entity[name];
        //                                        else
        //                                            conmt = true;
        //                                    }
        //                                }
        //                                else if (entity.GetAttributeValue<EntityReference>(name).LogicalName == "systemuser")
        //                                {
        //                                    Console.WriteLine(name);

        //                                }
        //                                else if (entity.GetAttributeValue<EntityReference>(name).LogicalName == "msdyn_incidenttypeservice")
        //                                {
        //                                    migrateData(entity.GetAttributeValue<EntityReference>(name).LogicalName, entity.GetAttributeValue<EntityReference>(name).Id.ToString());
        //                                    _createEntity[name] = entity[name];

        //                                }
        //                                else
        //                                    _createEntity[name] = entity[name];
        //                            }
        //                            else
        //                                _createEntity[name] = entity[name];
        //                        }
        //                    }
        //                    if (conmt)
        //                        continue;
        //                    //_createEntity["statecode"].GetType();
        //                    //if()
        //                    //_createEntity["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
        //                    //_createEntity["msdyn_unit"] = new EntityReference("uomschedule", getBaseUoMSchedule(uomId));
        //                    //_createEntity["msdyn_unit"] = new EntityReference("uom", uomId);
        //                    _serviceDev.Create(_createEntity);
        //                    count++;
        //                    Console.WriteLine(count + " is created ================================");

        //                }
        //                catch (Exception ex)
        //                {
        //                    if (ex.Message != "Cannot insert duplicate key.")
        //                        Console.WriteLine("Error " + ex.Message);
        //                    if (ex.Message.Contains("Does Not Exist"))
        //                    {
        //                        string[] arr = ex.Message.Split(' ');
        //                        string entityNameRef = arr[0];
        //                        string entityIDRef = arr[4];
        //                        if (!ex.Message.Contains("Product"))
        //                        {
        //                            migrateData(entityNameRef, entityIDRef);
        //                            goto Cont123;
        //                        }
        //                    }
        //                }
        //            }
        //            //query.PageInfo.PageNumber += 1;
        //            //query.PageInfo.PagingCookie = entCol.PagingCookie;
        //            //entCol = _servicePrd.RetrieveMultiple(query);
        //            //}
        //            //while (entCol.MoreRecords);
        //            Console.WriteLine("Count !!! " + count);
        //            // Console.WriteLine("****************************************** Entity " + entityName + " is ended.****************************************** ");

        //        }
        //        static Guid getBaseUoMSchedule(Guid UomId)
        //{
        //    Entity prdUoM = _serviceDev.Retrieve("uom", UomId, new ColumnSet("uomscheduleid"));

        //    if (!prdUoM.Contains("uomscheduleid"))
        //        return Guid.Empty;
        //    else
        //        return prdUoM.GetAttributeValue<EntityReference>("uomscheduleid").Id;

        //}
        public static IOrganizationService createConnection(string connectionString)
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
    }
}
