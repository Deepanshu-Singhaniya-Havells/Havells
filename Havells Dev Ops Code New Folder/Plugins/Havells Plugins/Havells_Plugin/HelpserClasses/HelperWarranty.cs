using System;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;

namespace Havells_Plugin
{
    
    public class HelperWarranty
    {
        public static void Warranty_InitFun(Entity entity, IOrganizationService service)
        {
            try
            {
                Entity CustAsst = new Entity("msdyn_customerasset");
                CustAsst = service.Retrieve("msdyn_customerasset", entity.Id, new ColumnSet(true));
                string Slno = string.Empty;
                EntityReference PdtCgry = new EntityReference("product");
                EntityReference Pdt = new EntityReference("product");
                EntityReference PdtSubCategory = new EntityReference("product");
                EntityReference PdtCode = new EntityReference("product");
                DateTime InvDate = new DateTime();
                DateTime RegDate = new DateTime();
                if ((CustAsst.Attributes.Contains("hil_invoicedate")) && (CustAsst.Attributes.Contains("hil_productcategory")))
                {
                    InvDate = (DateTime)CustAsst["hil_invoicedate"];
                    PdtCgry = (EntityReference)CustAsst["hil_productcategory"];
                    if ((CustAsst.Attributes.Contains("msdyn_name")) || (CustAsst.Attributes.Contains("msdyn_product")))
                    {
                        if (CustAsst.Attributes.Contains("msdyn_name"))
                        {
                            Slno = (string)CustAsst["msdyn_name"];
                        }
                        if (CustAsst.Attributes.Contains("msdyn_product"))
                        {
                            Pdt = (EntityReference)CustAsst["msdyn_product"];
                        }
                        if (CustAsst.Attributes.Contains("hil_registrationdate"))
                        {
                            RegDate = (DateTime)CustAsst["hil_registrationdate"];
                        }
                        if (CustAsst.Attributes.Contains("hil_productsubcategory"))
                        {
                            PdtSubCategory = (EntityReference)CustAsst["hil_productsubcategory"];
                        }
                        if (CustAsst.Attributes.Contains("hil_productcode"))
                        {
                            PdtCode = (EntityReference)CustAsst["hil_productcode"];
                        }
                        DateTime WarEndDate = LookForWarrantyTemplate(service, PdtCgry.Id, CustAsst, Pdt, PdtSubCategory.Id);
                        Product iDivision = (Product)service.Retrieve(Product.EntityLogicalName, PdtCgry.Id, new ColumnSet("hil_brandidentifier"));
                        if (iDivision.hil_BrandIdentifier != null && iDivision.hil_BrandIdentifier.Value == 2)
                        {
                            WarEndDate = CreateComprehensiveWarranty(service, PdtCgry.Id, CustAsst, Pdt, PdtSubCategory.Id, WarEndDate);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("---Havells_Plugin.HelperWarranty.Warranty_InitFun---" + ex.Message);
            }
        }
        #region Create Comprehensive Warranty
        public static DateTime CreateComprehensiveWarranty(IOrganizationService service, Guid DivisionId, Entity CustAsset, EntityReference Prod, Guid SubCategory, DateTime ITill)
        {
            DateTime WarrantyEnd = new DateTime();
            DateTime StartDate = new DateTime();
            Int32 Period = 0;
            msdyn_customerasset Asset = CustAsset.ToEntity<msdyn_customerasset>();
            if (Asset.hil_InvoiceDate != null)
            {
                DateTime InvDate = Convert.ToDateTime(Asset.hil_InvoiceDate);
                hil_unitwarranty iSchWarranty = new hil_unitwarranty();
                QueryExpression Query1 = new QueryExpression(hil_warrantyscheme.EntityLogicalName);
                Query1.ColumnSet = new ColumnSet(false);
                Query1.Criteria = new FilterExpression(LogicalOperator.And);
                Query1.Criteria.AddCondition("hil_productcategory", ConditionOperator.Equal, SubCategory);
                //Query1.Criteria.AddCondition("hil_warrantytemplate", ConditionOperator.NotNull);
                EntityCollection Found = service.RetrieveMultiple(Query1);
                if (Found.Entities.Count > 0)
                {
                    //hil_warrantyscheme iScheme = Found.Entities[0].ToEntity<hil_warrantyscheme>();
                    foreach (hil_warrantyscheme iScheme in Found.Entities)
                    {
                        QueryExpression Query2 = new QueryExpression(hil_schemeline.EntityLogicalName);
                        Query2.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate", "hil_warrantytemplate");
                        Query2.Criteria = new FilterExpression(LogicalOperator.And);
                        Query2.Criteria.AddCondition("hil_warrantyschemeid", ConditionOperator.Equal, iScheme.Id);
                        Query2.Criteria.AddCondition("hil_fromdate", ConditionOperator.NotNull);
                        Query2.Criteria.AddCondition("hil_todate", ConditionOperator.NotNull);
                        Query2.Criteria.AddCondition("hil_warrantytemplate", ConditionOperator.NotNull);
                        //Query2.AddOrder("hil_salesoffice", OrderType.Descending);
                        EntityCollection Found1 = service.RetrieveMultiple(Query2);
                        if (Found1.Entities.Count > 0)
                        {
                            foreach (hil_schemeline iSch in Found1.Entities)
                            {
                                if (InvDate >= Convert.ToDateTime(iSch.hil_FromDate) && InvDate <= Convert.ToDateTime(iSch.hil_ToDate))
                                {
                                    hil_warrantytemplate _iTemp = (hil_warrantytemplate)service.Retrieve(hil_warrantytemplate.EntityLogicalName, iSch.hil_WarrantyTemplate.Id, new ColumnSet("hil_warrantyperiod", "hil_type"));
                                    if(_iTemp.hil_type != null)
                                    {
                                        if(_iTemp.hil_type.Value == 5)
                                        {
                                            WarrantyEnd = ITill;
                                            CreateComponentWarranty(service, DivisionId, SubCategory, _iTemp, CustAsset.Id, ITill);
                                        }
                                        else if(_iTemp.hil_type.Value == 4)
                                        {
                                            Period = _iTemp.hil_WarrantyPeriod.Value;
                                            iSchWarranty.hil_CustomerAsset = new EntityReference(msdyn_customerasset.EntityLogicalName, CustAsset.Id);
                                            //iSchWarranty.hil_ProductType = new OptionSetValue(1);
                                            //iSchWarranty.hil_Part = new EntityReference(Product.EntityLogicalName, Prod.Id);
                                            iSchWarranty.hil_productmodel = new EntityReference(Product.EntityLogicalName, DivisionId);
                                            iSchWarranty.hil_productitem = new EntityReference(Product.EntityLogicalName, SubCategory);
                                            iSchWarranty.hil_warrantystartdate = ITill.AddDays(1);
                                            StartDate = ITill.AddDays(1);
                                            iSchWarranty.hil_warrantyenddate = StartDate.AddMonths(Period);
                                            iSchWarranty.hil_WarrantyTemplate = new EntityReference(hil_warrantytemplate.EntityLogicalName, _iTemp.hil_warrantytemplateId.Value);
                                            iSchWarranty.hil_pmscount = new OptionSetValue(910590006);
                                            iSchWarranty.hil_ProductType = new OptionSetValue(1);
                                            iSchWarranty.hil_Part = new EntityReference(Product.EntityLogicalName, Prod.Id);
                                            service.Create(iSchWarranty);
                                            WarrantyEnd = StartDate.AddMonths(Period);
                                            msdyn_customerasset iAsset = new msdyn_customerasset();
                                            iAsset.Id = CustAsset.Id;
                                            iAsset["hil_warrantytilldate"] = WarrantyEnd;
                                            service.Update(iAsset);
                                            //Console.WriteLine("Warranty Created");
                                            break;
                                        }
                                    }
                                    //if (_iTemp.hil_type != null && _iTemp.hil_type.Value == 5)//Component Warranty
                                    //{
                                    //    QueryExpression Query3 = new QueryExpression(hil_part.EntityLogicalName);
                                    //    Query3.ColumnSet = new ColumnSet("hil_period", "hil_partcode", "hil_warrantytemplateid");
                                    //    Query3.Criteria.AddCondition("hil_warrantytemplateid", ConditionOperator.Equal, _iTemp.Id);
                                    //    Query3.Criteria.AddCondition("hil_partcode", ConditionOperator.NotNull);
                                    //    Query3.Criteria.AddCondition("hil_period", ConditionOperator.NotNull);
                                    //    EntityCollection Found2 = service.RetrieveMultiple(Query3);
                                    //    if(Found2.Entities.Count > 0)
                                    //    {
                                    //        foreach(hil_part enPart in Found2.Entities)
                                    //        {
                                    //            Period = Convert.ToInt32(enPart.hil_Period);
                                    //            iSchWarranty.hil_CustomerAsset = new EntityReference(msdyn_customerasset.EntityLogicalName, CustAsset.Id);
                                    //            iSchWarranty.hil_ProductType = new OptionSetValue(2);
                                    //            iSchWarranty.hil_Part = new EntityReference(Product.EntityLogicalName, enPart.hil_PartCode.Id);
                                    //            iSchWarranty.hil_productmodel = new EntityReference(Product.EntityLogicalName, DivisionId);
                                    //            iSchWarranty.hil_productitem = new EntityReference(Product.EntityLogicalName, SubCategory);
                                    //            iSchWarranty.hil_warrantystartdate = ITill.AddDays(1);
                                    //            StartDate = ITill.AddDays(1);
                                    //            iSchWarranty.hil_warrantyenddate = StartDate.AddMonths(Period);
                                    //            iSchWarranty.hil_WarrantyTemplate = new EntityReference(hil_warrantytemplate.EntityLogicalName, _iTemp.hil_warrantytemplateId.Value);
                                    //            iSchWarranty.hil_pmscount = new OptionSetValue(910590006);
                                    //            iSchWarranty.hil_ProductType = new OptionSetValue(1);
                                    //            iSchWarranty.hil_Part = new EntityReference(Product.EntityLogicalName, Prod.Id);
                                    //            WarrantyEnd = StartDate.AddMonths(Period);
                                    //            service.Create(iSchWarranty);
                                    //            WarrantyEnd = ITill;
                                    //        }
                                    //    }
                                    //}
                                    //else if (_iTemp.hil_type != null && _iTemp.hil_type.Value == 4)// Comprehensive Warranty
                                    //{
                                        
                                }
                            }
                        }
                    }
                }
            }
            return WarrantyEnd;
        }
        #endregion
        #region Component Warranty
        public static void CreateComponentWarranty(IOrganizationService service, Guid Division, Guid SubCategory, hil_warrantytemplate enTemplate, Guid AssetId, DateTime TillDate)
        {
            hil_unitwarranty iSchWarranty = new hil_unitwarranty();
            Int32 Period = 0;
            DateTime StartDate = new DateTime();
            DateTime EndDate = new DateTime();
            QueryExpression Query = new QueryExpression(hil_part.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("hil_period", "hil_partcode", "hil_warrantytemplateid");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_warrantytemplateid", ConditionOperator.Equal, enTemplate.Id);
            Query.Criteria.AddCondition("hil_period", ConditionOperator.NotNull);
            Query.Criteria.AddCondition("hil_partcode", ConditionOperator.NotNull);
            EntityCollection Found1 = service.RetrieveMultiple(Query);
            if (Found1.Entities.Count > 0)
            {
                foreach(hil_part enPart in Found1.Entities)
                {
                    Period = Convert.ToInt32(enPart.hil_Period.Value);
                    iSchWarranty.hil_CustomerAsset = new EntityReference(msdyn_customerasset.EntityLogicalName, AssetId);
                    iSchWarranty.hil_productmodel = new EntityReference(Product.EntityLogicalName, Division);
                    iSchWarranty.hil_productitem = new EntityReference(Product.EntityLogicalName, SubCategory);
                    iSchWarranty.hil_warrantystartdate = TillDate.AddDays(1);
                    StartDate = TillDate.AddDays(1);
                    iSchWarranty.hil_warrantyenddate = StartDate.AddMonths(Period);
                    iSchWarranty.hil_WarrantyTemplate = new EntityReference(hil_warrantytemplate.EntityLogicalName, enPart.hil_WarrantyTemplateId.Id);
                    iSchWarranty.hil_pmscount = new OptionSetValue(910590006);
                    iSchWarranty.hil_ProductType = new OptionSetValue(2);
                    iSchWarranty.hil_Part = new EntityReference(Product.EntityLogicalName, enPart.hil_PartCode.Id);
                    service.Create(iSchWarranty);
                }
            }
        }
        #endregion
        #region Look For Warranty Schemes
        static void LookForWarrantySchemes(IOrganizationService service, Guid PdtCgryId, Entity CustAsst, EntityReference Pdt, Guid PdtSubCat, Guid PdtCode)
        {
            try
            {
                QueryExpression Qry = new QueryExpression();
                Qry.EntityName = hil_warrantyscheme.EntityLogicalName;
                ColumnSet Col = new ColumnSet("hil_warrantytemplate", "hil_productcategory");
                Qry.ColumnSet = Col;
                Qry.Criteria = new FilterExpression(LogicalOperator.Or);
                Qry.Criteria.AddCondition("hil_productcategory", ConditionOperator.Equal, PdtCgryId);
                Qry.Criteria.AddCondition("hil_productcategory", ConditionOperator.Equal, PdtSubCat);
                Qry.Criteria.AddCondition("hil_productcategory", ConditionOperator.Equal, PdtCode);
                Qry.Criteria.AddCondition("hil_productcategory", ConditionOperator.Equal, Pdt.Id);
                //QueryByAttribute Query = new QueryByAttribute("hil_warrantyscheme");
                //Query.AddAttributeValue("hil_productcategory", PdtCgryId);
                //ColumnSet Columns = new ColumnSet("hil_warrantytemplate", "hil_productcategory");
                //Query.ColumnSet = Columns;
                EntityCollection Found = service.RetrieveMultiple(Qry);
                if (Found.Entities.Count >= 1)
                {
                    foreach (Entity WSch in Found.Entities)
                    {
                        bool IfSchemeExist = CheckIfSchemeLinesPresent(service, CustAsst, WSch, PdtCgryId, PdtSubCat, Pdt);
                        if (IfSchemeExist == false)
                        {
                            LookForWarrantyTemplate(service, PdtCgryId, CustAsst, Pdt, PdtSubCat);
                        }
                    }
                }
                else
                {
                    LookForWarrantyTemplate(service, PdtCgryId, CustAsst, Pdt, PdtSubCat);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("---Havells_Plugin.HelperWarranty.LookForWarrantySchemes---" + ex.Message);
            }
        }
        #endregion
        #region Look For Warranty Template
        static DateTime LookForWarrantyTemplate(IOrganizationService service, Guid PdtCgryId, Entity CustAsst, EntityReference Pdt, Guid PdtSubCat)
        {
            DateTime WarEndDate = new DateTime();
            try
            {
                QueryByAttribute Query = new QueryByAttribute("hil_warrantytemplate");
                Query.AddAttributeValue("hil_product", PdtSubCat);
                Query.AddAttributeValue("hil_type", 1);
                ColumnSet Columns = new ColumnSet(true);
                Query.ColumnSet = Columns;
                EntityCollection Found = service.RetrieveMultiple(Query);
                if (Found.Entities.Count >= 1)
                {
                    foreach (Entity Wt in Found.Entities)
                    {
                        DateTime iPurchase = (DateTime)CustAsst["hil_invoicedate"];
                        if (Wt.Attributes.Contains("hil_validto") && Wt.Attributes.Contains("hil_validfrom"))
                        {
                            DateTime iValidTo = (DateTime)Wt["hil_validto"];
                            DateTime iValidFrom = (DateTime)Wt["hil_validfrom"];
                            if (iPurchase >= iValidFrom && iPurchase <= iValidTo)
                            {
                                if (Wt.Attributes.Contains("hil_warrantyperiod"))
                                {
                                    int WTPeriod;
                                    Product _SubCat = (Product)service.Retrieve(Product.EntityLogicalName, PdtSubCat, new ColumnSet("description"));
                                    string Desc = string.Empty;
                                    if (_SubCat.Description != null)
                                    {
                                        Desc = _SubCat.Description;
                                    }
                                    WTPeriod = (int)Wt["hil_warrantyperiod"];
                                    WarEndDate = CreateUnitWarrantyRecords(service, CustAsst, Wt.Id, WTPeriod, PdtCgryId, 1, Pdt, PdtSubCat, Desc);
                                    CreateUnitWarrantyForParts(service, Wt.Id, CustAsst, 2, Pdt, PdtSubCat, PdtCgryId);
                                }
                                else
                                {
                                    CreateUnitWarrantyForParts(service, Wt.Id, CustAsst, 2, Pdt, PdtSubCat, PdtCgryId);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("---Havells_Plugin.HelperWarranty.LookForWarrantyTemplate---" + ex.Message);
            }
            return WarEndDate;
        }
        #endregion
        #region Unit Warranty For Parts
        static void CreateUnitWarrantyForParts(IOrganizationService service, Guid WtId, Entity CustAsst, int ForParts, EntityReference Pdt, Guid ProductSubCategory, Guid ProductCategory)
        {
            try
            {
                int PartPeriod = 0;
                string PartDesc = string.Empty;
                EntityReference PartCode = new EntityReference("product");
                QueryByAttribute Query = new QueryByAttribute("hil_part");
                Query.AddAttributeValue("hil_warrantytemplateid", WtId);
                ColumnSet Col = new ColumnSet("hil_period", "hil_partcode");
                Query.ColumnSet = Col;
                EntityCollection Found = service.RetrieveMultiple(Query);
                if (Found.Entities.Count >= 1)
                {
                    foreach (Entity Prt in Found.Entities)
                    {
                        if (Prt.Attributes.Contains("hil_partcode"))
                        {
                            PartCode = (EntityReference)Prt["hil_partcode"];
                            Product _Prt = (Product)service.Retrieve(Product.EntityLogicalName, PartCode.Id, new ColumnSet("description"));
                            if (_Prt.Description != null)
                                PartDesc = _Prt.Description;
                            if (Prt.Attributes.Contains("hil_period"))
                            {
                                PartPeriod = (int)Prt["hil_period"];
                                CreateUnitWarrantyRecords(service, CustAsst, WtId, PartPeriod, ProductCategory, ForParts, PartCode, ProductSubCategory, PartDesc);
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("---Havells_Plugin.HelperWarranty.CreateUnitWarrantyForParts---" + ex.Message);
            }
        }
        #endregion
        #region Check If Scheme Lines Present
        static Boolean CheckIfSchemeLinesPresent(IOrganizationService service, Entity CustAsst, Entity WrtySch, Guid PdtCgryId, Guid PdtSubCategory, EntityReference Pdt)
        {
            bool IfPresent = false;
            try
            {

                EntityReference Region = new EntityReference("hil_region");
                DateTime RegDate = new DateTime();
                if (CustAsst.Attributes.Contains("hil_invoicedate"))
                {
                    RegDate = (DateTime)CustAsst["hil_invoicedate"];
                }
                EntityReference Customer = new EntityReference("contact");
                if (CustAsst.Attributes.Contains("hil_customer"))
                {
                    Customer = (EntityReference)CustAsst["hil_customer"];
                }
                Entity Contact = new Entity("contact");
                if (Customer.LogicalName == "contact")
                {
                    //Contact = service.Retrieve("contact", Customer.Id, new ColumnSet("hil_region"));
                    //if (Contact.Attributes.Contains("hil_region"))
                    //{
                    //    Region = (EntityReference)Contact["hil_region"];
                    //}
                    //else
                    //{
                    //    throw new InvalidPluginExecutionException("Region can't be null in contat");
                    //}
                    QueryByAttribute Qry = new QueryByAttribute(hil_address.EntityLogicalName);
                    Qry.ColumnSet = new ColumnSet("hil_region");
                    Qry.AddAttributeValue("hil_customer", Customer.Id);
                    Qry.AddAttributeValue("hil_addresstype", 1);
                    EntityCollection Found1 = service.RetrieveMultiple(Qry);
                    if (Found1.Entities.Count > 0)
                    {
                        foreach (hil_address Add in Found1.Entities)
                        {
                            if (Add.Attributes.Contains("hil_region"))
                            {
                                Region = (EntityReference)Add["hil_region"];
                            }
                            else
                            {
                                throw new InvalidPluginExecutionException("Region can't be null");
                            }
                        }
                    }
                }//Revisit_ What if Account is Selected?
                QueryExpression Query = new QueryExpression("hil_schemeline");
                ConditionExpression condition1 = new ConditionExpression("hil_warrantyschemeid", ConditionOperator.Equal, WrtySch.Id);
                ConditionExpression condition2 = new ConditionExpression("hil_fromdate", ConditionOperator.GreaterEqual, RegDate);
                ConditionExpression condition3 = new ConditionExpression("hil_todate", ConditionOperator.LessEqual, RegDate);
                ConditionExpression condition4 = new ConditionExpression("hil_region", ConditionOperator.Equal, Region.Id);
                FilterExpression filter1 = new FilterExpression(LogicalOperator.And);
                filter1.AddCondition(condition1);
                filter1.AddCondition(condition2);
                filter1.AddCondition(condition3);
                filter1.AddCondition(condition4);
                Query.ColumnSet.AddColumns("hil_warrantyschemeid", "hil_fromdate", "hil_todate", "hil_region", "hil_warrantytemplate");
                Query.Criteria.AddFilter(filter1);
                EntityCollection Found = service.RetrieveMultiple(Query);
                if (Found.Entities.Count >= 1)
                {
                    IfPresent = true;
                    foreach (Entity SchLn in Found.Entities)
                    {
                        EntityReference WarrantyTemp = new EntityReference("hil_warrantytemplate");
                        if (SchLn.Attributes.Contains("hil_warrantytemplate"))
                        {
                            WarrantyTemp = (EntityReference)SchLn["hil_warrantytemplate"];
                            Entity Wt = new Entity("hil_warrantytemplate");
                            Wt = service.Retrieve("hil_warrantytemplate", WarrantyTemp.Id, new ColumnSet("hil_warrantyperiod"));
                            if (Wt.Attributes.Contains("hil_warrantyperiod"))
                            {
                                Product _SubCat = (Product)service.Retrieve(Product.EntityLogicalName, PdtSubCategory, new ColumnSet("description"));
                                string Desc = string.Empty;
                                if (_SubCat.Description != null)
                                {
                                    Desc = _SubCat.Description;
                                }
                                int WTPeriod = (int)Wt["hil_warrantyperiod"];
                                CreateUnitWarrantyRecords(service, CustAsst, Wt.Id, WTPeriod, PdtCgryId, 1, Pdt, PdtSubCategory, Desc);
                                CreateUnitWarrantyForParts(service, Wt.Id, CustAsst, 2, Pdt, PdtSubCategory, PdtCgryId);
                            }
                            else
                            {
                                CreateUnitWarrantyForParts(service, Wt.Id, CustAsst, 2, Pdt, PdtSubCategory, PdtCgryId);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("---Havells_Plugin.HelperWarranty.CheckIfSchemeLinesPresent---" + ex.Message);
            }
            return IfPresent;
        }
        #endregion
        #region Create Unit Warranty Records
        static DateTime CreateUnitWarrantyRecords(IOrganizationService service, Entity CustAsst, Guid WrtyTmplt, int Duration, Guid PrdCategory, int ProductType, EntityReference Pdt, Guid ProductSubCategory, string Desc)
        {
            DateTime WarEndDt = new DateTime();
            try
            {
                Entity UtWrty = new Entity("hil_unitwarranty");
                DateTime TimeNow = DateTime.Now.ToLocalTime();
                UtWrty["hil_customerasset"] = new EntityReference("msdyn_customerasset", CustAsst.Id);
                UtWrty["hil_warrantytemplate"] = new EntityReference("hil_warrantytemplate", WrtyTmplt);
                UtWrty["hil_warrantystartdate"] = (DateTime)CustAsst["hil_invoicedate"];
                UtWrty["hil_producttype"] = new OptionSetValue(ProductType);
                DateTime WrtyStDt = (DateTime)CustAsst["hil_invoicedate"];
                WarEndDt = WrtyStDt.AddMonths(Duration);
                WarEndDt = WarEndDt.AddDays(-1);
                UtWrty["hil_warrantyenddate"] = (DateTime)WarEndDt;
                UtWrty["hil_productitem"] = new EntityReference("product", ProductSubCategory);
                UtWrty["hil_productmodel"] = new EntityReference("product", PrdCategory);
                if (Desc != null)
                    UtWrty["hil_partdescription"] = Desc;
                if (Pdt.Id != Guid.Empty)
                {
                    UtWrty["hil_part"] = Pdt;
                }
                if (TimeNow >= WrtyStDt && TimeNow <= WarEndDt)
                {
                    UtWrty["statuscode"] = new OptionSetValue(1);
                }
                else
                {
                    UtWrty["statuscode"] = new OptionSetValue(910590000);
                }
                Guid WtyId = service.Create(UtWrty);
                if (WtyId != Guid.Empty && ProductType == 1)
                {
                    Entity iAsst = new Entity(msdyn_customerasset.EntityLogicalName);
                    iAsst.Id = CustAsst.Id;
                    iAsst["hil_warrantystatus"] = new OptionSetValue(1);
                    iAsst["hil_warrantytilldate"] = WarEndDt;
                    service.Update(iAsst);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("---Havells_Plugin.HelperWarranty.CreateUnitWarrantyRecords---" + ex.Message);
            }
            return WarEndDt;
        }
        #endregion
        public static hil_inventorytype GetWarrantyStatus(IOrganizationService service, Guid CustAsset)
        {
            try
            {
                hil_inventorytype result = new hil_inventorytype();
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    var obj = (from _UtWrty in orgContext.CreateQuery<hil_unitwarranty>()
                               where _UtWrty.hil_CustomerAsset.Id == CustAsset //&& //_UtWrty.hil_productitem.Id == ProductCode &&
                                && _UtWrty.hil_ProductType.Value == 1                                               //_UtWrty.hil_warrantystartdate <= DateTime.UtcNow && _UtWrty.hil_warrantyenddate >= DateTime.UtcNow
                               select new
                               {
                                   _UtWrty.hil_CustomerAsset,
                                   _UtWrty.hil_productitem,
                                   _UtWrty.hil_warrantyenddate,
                                   _UtWrty.hil_warrantystartdate,
                               }).Take(1);
                    if (Enumerable.Count(obj) > 0)
                    {
                        foreach (var iobj in obj)
                        {
                            if (iobj.hil_warrantystartdate != null && iobj.hil_warrantyenddate != null)
                            {
                                DateTime today = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.ToUniversalTime().Day);
                                //DateTime startTime = iobj.hil_warrantystartdate.Value.AddMinutes(330);
                                //DateTime endTime = iobj.hil_warrantyenddate.Value.AddMinutes(330);
                                DateTime startTime = iobj.hil_warrantystartdate.Value;
                                DateTime endTime = iobj.hil_warrantyenddate.Value;

                                if (startTime <= today && endTime >= today)
                                {
                                    result = hil_inventorytype.InWarranty;
                                }
                                else
                                {
                                    result = hil_inventorytype.OutWarranty;
                                }
                            }
                            else
                            {
                                result = hil_inventorytype.NANotfound;
                            }
                        }
                    }
                    else
                    {
                        result = hil_inventorytype.NANotfound;
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("---Havells_Plugin.HelperWarranty.GetWarrantyStatus" + ex.Message);
            }
        }

        public static DateTime GetWarrantyEndDate(IOrganizationService service, Guid CustAsset)
        {
            try
            {
                DateTime result = DateTime.MinValue;
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    var obj = (from _UtWrty in orgContext.CreateQuery<hil_unitwarranty>()
                               where _UtWrty.hil_CustomerAsset.Id == CustAsset //&& //_UtWrty.hil_productitem.Id == ProductCode &&
                                 && _UtWrty.hil_ProductType.Value == 1                                              //_UtWrty.hil_warrantystartdate <= DateTime.UtcNow && _UtWrty.hil_warrantyenddate >= DateTime.UtcNow
                               select new
                               {
                                   _UtWrty.hil_CustomerAsset,
                                   _UtWrty.hil_productitem,
                                   _UtWrty.hil_warrantyenddate,
                                   _UtWrty.hil_warrantystartdate,
                               }).Take(1);
                    if (Enumerable.Count(obj) > 0)
                    {
                        foreach (var iobj in obj)
                        {
                            if (iobj.hil_warrantystartdate != null && iobj.hil_warrantyenddate != null)
                            {
                                DateTime today = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.ToUniversalTime().Day);
                                DateTime startTime = iobj.hil_warrantystartdate.Value.AddMinutes(330);
                                DateTime endTime = iobj.hil_warrantyenddate.Value.AddMinutes(330);
                                result = endTime;
                            }
                        }
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("---Havells_Plugin.HelperWarranty.GetWarrantyEnd" + ex.Message);
            }
        }
    }
}