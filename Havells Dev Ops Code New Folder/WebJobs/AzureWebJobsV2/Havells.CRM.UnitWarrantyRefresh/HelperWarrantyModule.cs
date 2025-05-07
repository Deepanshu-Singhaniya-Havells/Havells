using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IdentityModel.Metadata;

namespace Havells.CRM.UnitWarrantyRefresh
{
    public class HelperWarrantyModule
    {
        public static void UnitWarranty_AMC(Entity CustAsst, IOrganizationService service)
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

                //if (CustAsst.Attributes.Contains("hil_invoiceavailable") && CustAsst.Attributes.Contains("hil_invoicedate") && CustAsst.Attributes.Contains("hil_productcategory"))
                //{
                //    invoiceAvailable = (bool)CustAsst["hil_invoiceavailable"];
                bool skip = false;
                DateTime lastEndDate = DateTime.MinValue;
                LinkEntity EntityA = new LinkEntity("hil_unitwarranty", "hil_warrantytemplate", "hil_warrantytemplate", "hil_warrantytemplateid", JoinOperator.LeftOuter);
                EntityA.Columns = new ColumnSet("hil_type");
                EntityA.EntityAlias = "WarrantyTemplate";

                WarrentyType lastType = new WarrentyType();
                QueryExpression queryExp1 = new QueryExpression("hil_unitwarranty");
                queryExp1.ColumnSet = new ColumnSet("hil_unitwarrantyid", "hil_warrantyenddate");
                queryExp1.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp1.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, CustAsst.Id);
                queryExp1.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                queryExp1.Orders.Add(new OrderExpression("hil_warrantyenddate", OrderType.Ascending));
                queryExp1.LinkEntities.Add(EntityA);
                EntityCollection entCol1 = service.RetrieveMultiple(queryExp1);
                if (entCol1.Entities.Count > 0)
                {
                    foreach (Entity ent in entCol1.Entities)
                    {
                        if (lastType != (WarrentyType)((OptionSetValue)ent.GetAttributeValue<AliasedValue>("WarrantyTemplate.hil_type").Value).Value
                            || lastEndDate != ent.GetAttributeValue<DateTime>("hil_warrantyenddate"))
                        {
                            skip = false;
                        }
                        if (!skip)
                        {
                            lastEndDate = ent.GetAttributeValue<DateTime>("hil_warrantyenddate");
                            lastType = (WarrentyType)((OptionSetValue)ent.GetAttributeValue<AliasedValue>("WarrantyTemplate.hil_type").Value).Value;
                            skip = true;
                            Console.WriteLine("Skiped: " + lastType.ToString() + "|EndDate " + lastEndDate);
                        }
                        else
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
                            service.Execute(setStateRequest);
                            Console.WriteLine("Deactiveated: " + lastType.ToString() + "|EndDate " + lastEndDate);
                        }
                    }
                }

                //}
            }
            catch (Exception ex)
            {

            }
        }
        public static void Init_Warranty(Entity CustAsst, IOrganizationService _service)
        {
            try
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
                            bool skip = false;

                            LinkEntity EntityA = new LinkEntity("hil_unitwarranty", "hil_warrantytemplate", "hil_warrantytemplate", "hil_warrantytemplateid", JoinOperator.LeftOuter);
                            EntityA.Columns = new ColumnSet("hil_type");
                            EntityA.EntityAlias = "WarrantyTemplate";

                            WarrentyType lastType = new WarrentyType();
                            QueryExpression queryExp1 = new QueryExpression("hil_unitwarranty");
                            queryExp1.ColumnSet = new ColumnSet("hil_unitwarrantyid", "hil_warrantyenddate");
                            queryExp1.Criteria = new FilterExpression(LogicalOperator.And);
                            queryExp1.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, CustAsst.Id);
                            queryExp1.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                            queryExp1.Orders.Add(new OrderExpression("hil_warrantyenddate", OrderType.Ascending));
                            queryExp1.LinkEntities.Add(EntityA);
                            EntityCollection entCol1 = _service.RetrieveMultiple(queryExp1);
                            if (entCol1.Entities.Count > 0)
                            {
                                foreach (Entity ent in entCol1.Entities)
                                {
                                    if (lastType != (WarrentyType)((OptionSetValue)ent.GetAttributeValue<AliasedValue>("WarrantyTemplate.hil_type").Value).Value)
                                    {
                                        skip = false;
                                    }
                                    if (!skip)
                                    {
                                        lastType = (WarrentyType)((OptionSetValue)ent.GetAttributeValue<AliasedValue>("WarrantyTemplate.hil_type").Value).Value;
                                        skip = true;
                                    }
                                    else
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
                                    Console.WriteLine(lastType.ToString());
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
                            string serialNumber = CustAsst.GetAttributeValue<string>("msdyn_name");

                            queryExp = new QueryExpression("hil_amcstaging");
                            queryExp.ColumnSet = new ColumnSet("hil_name", "hil_warrantystartdate", "hil_warrantyenddate", "hil_serailnumber", "hil_sapbillingdocpath", "hil_sapbillingdate", "hil_amcstagingstatus", "hil_amcplan");
                            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                            queryExp.Criteria.AddCondition("hil_serailnumber", ConditionOperator.Equal, serialNumber);
                            queryExp.Criteria.AddCondition("hil_amcstagingstatus", ConditionOperator.Equal, false); //posted
                            // queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
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
                            if (stdWrtyEndDate != null)
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
                            if (endDate != null && !_specialSchemeApplied)
                            {
                                queryExp = new QueryExpression("hil_warrantytemplate");
                                queryExp.ColumnSet = new ColumnSet("hil_validfrom", "hil_validto", "hil_warrantyperiod", "hil_category");
                                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                queryExp.Criteria.AddCondition("hil_product", ConditionOperator.Equal, erProdSubCategory.Id);
                                queryExp.Criteria.AddCondition("hil_validfrom", ConditionOperator.NotNull);
                                queryExp.Criteria.AddCondition("hil_validto", ConditionOperator.NotNull);
                                queryExp.Criteria.AddCondition("hil_type", ConditionOperator.Equal, (int)WarrentyType.Extended); //Extended Warranty
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
                                    endDate = entCol.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate");
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
                                entCustAsset["hil_warrantytilldate"] = new DateTime(1900, 1, 1);
                                entCustAsset["hil_warrantystatus"] = new OptionSetValue(2); //OutWarranty
                                entCustAsset["hil_warrantysubstatus"] = null;
                            }
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
                            //entCustAsset["hil_customernm"] = "DONE";
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
                        //entCustAsset["hil_customernm"] = "DONE";
                        _service.Update(entCustAsset);
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(" *** Havells_Plugin.HelperWarrantyModule.Init_Warranty *** " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(" *** Havells_Plugin.HelperWarrantyModule.Init_Warranty *** " + ex.Message);
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
                qryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
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

                    SetStateRequest setStateRequest = new SetStateRequest()
                    {
                        EntityMoniker = new EntityReference
                        {
                            Id = entCol[0].Id,
                            LogicalName = "hil_unitwarranty",
                        },
                        State = new OptionSetValue(0), //Inactive
                        Status = new OptionSetValue(1) //Inactive
                    };
                    service.Execute(setStateRequest);
                    DateTime StartDate = new DateTime();
                    int i = 0;
                    hil_unitwarranty iSchWarranty = new hil_unitwarranty();
                    iSchWarranty.Id = entCol[0].Id;
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
                    service.Update(iSchWarranty);
                }
                return WarrantyEnd;
            }
            catch (Exception ex)
            {
                return WarrantyEnd;
            }
        }
        static DateTime? CreateSchemeUnitWarrantyLine(IOrganizationService service, int period, int producttype, EntityReference erCustomer, EntityReference erCustomerasset, EntityReference erProductCatg, EntityReference erProductSubCatg, EntityReference erProductModel, string partdescription, DateTime? warrantystartdate, EntityReference erWarrantytemplate, DateTime invDate)
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

        public static void CreateUnitWarrantyLines(IOrganizationService service, Entity CustAsst)
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
                Entity template = null;
                if (CustAsst.Contains("statuscode"))
                {
                    int _statusCode = CustAsst.GetAttributeValue<OptionSetValue>("statuscode").Value;
                    if (_statusCode == (int)AssetStatusCode.ProductApproved)
                    {
                        DactivateUnitWarrantyLins(service, CustAsst.ToEntityReference());
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

                        //WarrentyType == Standard
                        template = RetriveWarrantyTemplate(service, erProdSubCategory, WarrentyType.Standard, (DateTime)invDate);
                        if (template != null)
                        {
                            CreateUpdateUnitWarrantyLine(service, template.GetAttributeValue<int>("hil_warrantyperiod"), erCustomer, CustAsst.ToEntityReference(), erProdCgry, erProdSubCategory, erProdModelCode, pdtModelName, invDate, template.ToEntityReference());
                        }
                        //WarrentyType == Extended
                        template = RetriveWarrantyTemplate(service, erProdSubCategory, WarrentyType.Extended, (DateTime)invDate);
                        if (template != null)
                        {
                            CreateUpdateUnitWarrantyLine(service, template.GetAttributeValue<int>("hil_warrantyperiod"), erCustomer, CustAsst.ToEntityReference(), erProdCgry, erProdSubCategory, erProdModelCode, pdtModelName, invDate, template.ToEntityReference());
                        }
                        ////WarrentyType == SpecialScheme
                        //template = RetriveWarrantyTemplate(service, erProdSubCategory, WarrentyType.SpecialScheme, (DateTime)invDate);
                        //if (template != null)
                        //{
                        //    CreateUpdateUnitWarrantyLine(service, template.GetAttributeValue<int>("hil_warrantyperiod"), erCustomer, CustAsst.ToEntityReference(), erProdCgry, erProdSubCategory, erProdModelCode, pdtModelName, invDate, template.ToEntityReference());
                        //}
                    }
                    else
                    {
                        UpdateWarrantyStatusOnAsset(service, CustAsst.ToEntityReference(), 0, (int)WarrentyStatus.NAforWarranty, null);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in Creating Unit Warranty Line || " + ex.Message);
            }

        }
        static void DactivateUnitWarrantyLins(IOrganizationService service, EntityReference CustAsst)
        {
            try
            {
                QueryExpression queryExp1 = new QueryExpression("hil_unitwarranty");
                queryExp1.ColumnSet = new ColumnSet("hil_unitwarrantyid", "hil_warrantyenddate");
                queryExp1.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp1.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, CustAsst.Id);
                queryExp1.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                EntityCollection entCol1 = service.RetrieveMultiple(queryExp1);
                if (entCol1.Entities.Count > 0)
                {
                    foreach (Entity ent in entCol1.Entities)
                    {
                        SetStateRequest setStateRequest = new SetStateRequest()
                        {
                            EntityMoniker = ent.ToEntityReference(),
                            State = new OptionSetValue(1), //Inactive
                            Status = new OptionSetValue(2) //Inactive
                        };
                        service.Execute(setStateRequest);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error in Deactivating Unit Warranty Lines of Asset || " + ex.Message);
            }
        }
        static void UpdateWarrantyStatusOnAsset(IOrganizationService service, EntityReference entityReference, int _subStatus, int _Status, DateTime? _warrantyTillDate)
        {
            try
            {
                Entity _CustAsset = new Entity(entityReference.LogicalName, entityReference.Id);
                _CustAsset["hil_warrantysubstatus"] = _subStatus != 0 ? new OptionSetValue(_subStatus) : null;
                _CustAsset["hil_warrantystatus"] = new OptionSetValue(_Status);
                _CustAsset["hil_warrantytilldate"] = _warrantyTillDate;
                service.Update(_CustAsset);
            }
            catch (Exception ex)
            {
                throw new Exception("Error in Updating Warranty Status of Asset || " + ex.Message);
            }
        }
        static Entity RetriveWarrantyTemplate(IOrganizationService service, EntityReference erProdSubCategory, WarrentyType warrentyType, DateTime invDate)
        {
            Entity _WarrantyTemplate = null;
            try
            {
                QueryExpression queryExp = new QueryExpression("hil_warrantytemplate");
                queryExp.ColumnSet = new ColumnSet("hil_validfrom", "hil_validto", "hil_warrantyperiod", "hil_category");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("hil_product", ConditionOperator.Equal, erProdSubCategory.Id);
                queryExp.Criteria.AddCondition("hil_validfrom", ConditionOperator.NotNull);
                queryExp.Criteria.AddCondition("hil_validto", ConditionOperator.NotNull);
                queryExp.Criteria.AddCondition("hil_type", ConditionOperator.Equal, (int)warrentyType);
                queryExp.Criteria.AddCondition("hil_templatestatus", ConditionOperator.Equal, 2); //Approved
                queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                EntityCollection entCol = service.RetrieveMultiple(queryExp);
                if (entCol.Entities.Count > 0)
                {
                    foreach (Entity ent in entCol.Entities)
                    {
                        if (invDate >= ent.GetAttributeValue<DateTime>("hil_validfrom") && invDate <= ent.GetAttributeValue<DateTime>("hil_validto"))
                        {
                            _WarrantyTemplate = ent;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error in Retrieving Warranty Template of Asset || " + ex.Message);
            }
            return _WarrantyTemplate;
        }
        static void CreateUpdateUnitWarrantyLine(IOrganizationService service, int period, EntityReference erCustomer, EntityReference erCustomerasset, EntityReference erProductCatg, EntityReference erProductSubCatg, EntityReference erProductModel, string partdescription, DateTime? warrantystartdate, EntityReference erWarrantytemplate)
        {
            DateTime? WarrantyEnd = null;
            try
            {
                QueryExpression qryExp = new QueryExpression(hil_unitwarranty.EntityLogicalName);
                qryExp.ColumnSet = new ColumnSet("hil_warrantyenddate");
                qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                qryExp.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, erCustomerasset.Id);
                qryExp.Criteria.AddCondition("hil_warrantytemplate", ConditionOperator.Equal, erWarrantytemplate.Id);
                qryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
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

                    SetStateRequest setStateRequest = new SetStateRequest()
                    {
                        EntityMoniker = new EntityReference
                        {
                            Id = entCol[0].Id,
                            LogicalName = "hil_unitwarranty",
                        },
                        State = new OptionSetValue(0), //Inactive
                        Status = new OptionSetValue(1) //Inactive
                    };
                    service.Execute(setStateRequest);
                    DateTime StartDate = new DateTime();
                    int i = 0;
                    hil_unitwarranty iSchWarranty = new hil_unitwarranty();
                    iSchWarranty.Id = entCol[0].Id;
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
                    service.Update(iSchWarranty);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in Creating Unit Warranty Line || " + ex.Message);
            }
        }
        static void AMCUnitWarrantyLines(IOrganizationService service, string serialNumber)
        {
            QueryExpression queryExp = new QueryExpression("hil_amcstaging");
            queryExp.ColumnSet = new ColumnSet("hil_name", "hil_warrantystartdate", "hil_warrantyenddate", "hil_serailnumber", "hil_sapbillingdocpath", "hil_sapbillingdate", "hil_amcstagingstatus", "hil_amcplan");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("hil_serailnumber", ConditionOperator.Equal, serialNumber);
            // queryExp.Criteria.AddCondition("hil_amcstagingstatus", ConditionOperator.Equal, false); //posted
            queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
            queryExp.AddOrder("hil_warrantystartdate", OrderType.Ascending);
            EntityCollection entCol = service.RetrieveMultiple(queryExp);

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
        }
    }
}
