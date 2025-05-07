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
using Microsoft.Crm.Sdk.Messages;

namespace Havells_Plugin
{
    public class HelperWarrantyModule
    {
        public static void Init_Warranty(Guid entityId, IOrganizationService _service)
        {
            try
            {
                QueryExpression qryExp = new QueryExpression("msdyn_customerasset");
                qryExp.ColumnSet = new ColumnSet("msdyn_name","hil_invoiceavailable", "hil_productcategory", "hil_invoicedate", "hil_productsubcategory", "msdyn_product", "hil_modelname", "hil_customer");
                qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                qryExp.Criteria.AddCondition("msdyn_customerassetid", ConditionOperator.Equal, entityId);
                EntityCollection ec = _service.RetrieveMultiple(qryExp);
                foreach (Entity CustAsst in ec.Entities)
                {
                    try
                    {
                        EntityReference erProdCgry = new EntityReference("product");
                        EntityReference erProdSubCategory = new EntityReference("product");
                        EntityReference erProdModelCode = new EntityReference("product");
                        EntityReference erCustomer = new EntityReference("contact");
                        string pdtModelName = string.Empty;
                        string _serialNumber = string.Empty;
                        DateTime? invDate = null;
                        DateTime? endDate = null;
                        bool invoiceAvailable = false;
                        int? warrantySubstatus = null;
                        DateTime? stdWrtyEndDate = null;
                        DateTime? spcWrtyEndDate = null;
                        DateTime? extWrtyEndDate = null;
                        string _fetchXML = string.Empty;
                        EntityCollection entCol = null;
                        if (CustAsst.Contains("msdyn_name"))
                        {
                            

                            if (CustAsst.Attributes.Contains("hil_invoiceavailable") && CustAsst.Attributes.Contains("hil_invoicedate") && CustAsst.Attributes.Contains("hil_productcategory"))
                            {
                                invoiceAvailable = (bool)CustAsst["hil_invoiceavailable"];
                                if (invoiceAvailable)
                                {
                                    invDate = (DateTime)CustAsst["hil_invoicedate"];
                                    string _invDateStr = invDate.Value.Year.ToString() + "-" + invDate.Value.Month.ToString().PadLeft(2, '0') + "-" + invDate.Value.Day.ToString().PadLeft(2, '0');

                                    if (CustAsst.Attributes.Contains("msdyn_name"))
                                    {
                                        _serialNumber = (string)CustAsst["msdyn_name"];
                                    }

                                    QueryExpression queryExp1 = new QueryExpression("hil_unitwarranty");
                                    queryExp1.ColumnSet = new ColumnSet("hil_unitwarrantyid");
                                    queryExp1.Criteria = new FilterExpression(LogicalOperator.And);
                                    queryExp1.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, CustAsst.Id);
                                    queryExp1.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                                    queryExp1.Criteria.AddCondition("hil_customer", ConditionOperator.NotNull); //Old Unit Warranties
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
                                    QueryExpression queryExp = null;
                                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                      <entity name='hil_warrantytemplate'>
                                        <attribute name='hil_validfrom' />
                                        <attribute name='hil_validto' />
                                        <attribute name='hil_warrantyperiod' />
                                        <attribute name='hil_warrantytemplateid' />
                                        <attribute name='hil_category' />
                                        <order attribute='hil_name' descending='false' />
                                        <filter type='and'>
                                          <condition attribute='hil_validfrom' operator='on-or-before' value='{_invDateStr}' />
                                          <condition attribute='hil_validto' operator='on-or-after' value='{_invDateStr}' />
                                          <condition attribute='hil_type' operator='eq' value='1' />
                                          <condition attribute='hil_templatestatus' operator='eq' value='2' />
                                          <condition attribute='statecode' operator='eq' value='0' />
                                          <condition attribute='hil_applicableon' operator='eq' value='910590001' />
                                        </filter>
                                        <link-entity name='hil_warrantytemplateline' from='hil_warrantytemplate' to='hil_warrantytemplateid' link-type='inner' alias='ac'>
                                          <filter type='and'>
                                            <condition attribute='hil_model' operator='eq' value='{erProdModelCode.Id}' />
                                          </filter>
                                        </link-entity>
                                      </entity>
                                    </fetch>";
                                    entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                    if (entCol.Entities.Count > 0)
                                    {
                                        //EntityReference _warrantyTemp = new EntityReference("hil_warrantytemplate", entCol.Entities[0].Id);
                                        endDate = CreateUnitWarrantyLine(_service, entCol.Entities[0].GetAttributeValue<int>("hil_warrantyperiod"), 1, erCustomer, CustAsst.ToEntityReference(), erProdCgry, erProdSubCategory, erProdModelCode, pdtModelName, invDate, entCol.Entities[0].ToEntityReference());
                                        stdWrtyEndDate = endDate;
                                        warrantySubstatus = 1;
                                    }
                                    else
                                    {
                                        queryExp = new QueryExpression("hil_warrantytemplate");
                                        queryExp.ColumnSet = new ColumnSet("hil_validfrom", "hil_validto", "hil_warrantyperiod", "hil_category");
                                        queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                        queryExp.Criteria.AddCondition("hil_product", ConditionOperator.Equal, erProdSubCategory.Id);
                                        queryExp.Criteria.AddCondition("hil_validfrom", ConditionOperator.NotNull);
                                        queryExp.Criteria.AddCondition("hil_validto", ConditionOperator.NotNull);
                                        queryExp.Criteria.AddCondition("hil_type", ConditionOperator.Equal, 1); //Standard Warranty
                                        queryExp.Criteria.AddCondition("hil_templatestatus", ConditionOperator.Equal, 2); //Approved
                                        queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                                        queryExp.Criteria.AddCondition("hil_applicableon", ConditionOperator.Equal, 910590000); //Standard Warranty
                                        entCol = _service.RetrieveMultiple(queryExp);
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

                                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                      <entity name='hil_warrantytemplate'>
                                        <attribute name='hil_validfrom' />
                                        <attribute name='hil_validto' />
                                        <attribute name='hil_warrantyperiod' />
                                        <attribute name='hil_warrantytemplateid' />
                                        <attribute name='hil_category' />
                                        <order attribute='hil_name' descending='false' />
                                        <filter type='and'>
                                          <condition attribute='hil_validfrom' operator='on-or-before' value='{_invDateStr}' />
                                          <condition attribute='hil_validto' operator='on-or-after' value='{_invDateStr}' />
                                          <condition attribute='hil_type' operator='eq' value='2' />
                                          <condition attribute='hil_templatestatus' operator='eq' value='2' />
                                          <condition attribute='statecode' operator='eq' value='0' />
                                          <condition attribute='hil_applicableon' operator='eq' value='910590001' />
                                        </filter>
                                        <link-entity name='hil_warrantytemplateline' from='hil_warrantytemplate' to='hil_warrantytemplateid' link-type='inner' alias='ac'>
                                          <filter type='and'>
                                            <condition attribute='hil_model' operator='eq' value='{erProdModelCode.Id}' />
                                          </filter>
                                        </link-entity>
                                      </entity>
                                    </fetch>";
                                    entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                    if (entCol.Entities.Count > 0)
                                    {
                                        //EntityReference _warrantyTemp = new EntityReference("hil_warrantytemplate", entCol.Entities[0].Id);
                                        endDate = CreateUnitWarrantyLine(_service, entCol.Entities[0].GetAttributeValue<int>("hil_warrantyperiod"), 2, erCustomer, CustAsst.ToEntityReference(), erProdCgry, erProdSubCategory, erProdModelCode, pdtModelName, endDate.Value.AddDays(1), entCol.Entities[0].ToEntityReference());
                                        stdWrtyEndDate = endDate;
                                        warrantySubstatus = 1;
                                    }
                                    else
                                    {
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
                                            queryExp.Criteria.AddCondition("hil_applicableon", ConditionOperator.Equal, 910590000); //Standard Warranty
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
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidPluginExecutionException(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
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
                    WarrantyEnd = entCol.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate");
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

        
    }
}
