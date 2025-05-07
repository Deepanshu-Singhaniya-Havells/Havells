using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Text.RegularExpressions;
using Havells_Plugin;

namespace Havells_Plugin.AMC_Staging
{
    public class PreCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            Entity entity = (Entity)context.InputParameters["Target"];
            Guid pmsJobId = Guid.Empty;
            try
            {
                EntityReference _amcPlan = entity.GetAttributeValue<EntityReference>("hil_amcplan");
                string _sapBillingDoc = entity.GetAttributeValue<string>("hil_name");
                DateTime _sapBillingDate = entity.GetAttributeValue<DateTime>("hil_sapbillingdate");
                string _sapBillingDocPath = entity.GetAttributeValue<string>("hil_sapbillingdocpath");
                string _serailNumber = entity.GetAttributeValue<string>("hil_serailnumber");
                DateTime _warrantyStartDate = entity.GetAttributeValue<DateTime>("hil_warrantystartdate");
                DateTime _warrantyEndDate = entity.GetAttributeValue<DateTime>("hil_warrantyenddate");
                Entity _warrantyTemplate = null;
                Entity _customerAsset = null;
                QueryExpression qryExp = null;
                EntityCollection enCol = null;

                qryExp = new QueryExpression("msdyn_customerasset");
                qryExp.ColumnSet = new ColumnSet("hil_customer", "msdyn_customerassetid", "msdyn_product", "hil_productcategory", "hil_productsubcategory", "hil_modelname");
                qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                qryExp.Criteria.AddCondition(new ConditionExpression("msdyn_name", ConditionOperator.Equal, _serailNumber));
                enCol = service.RetrieveMultiple(qryExp);
                if (enCol.Entities.Count > 0)
                {
                    _customerAsset = enCol.Entities[0];
                }
                else
                {
                    entity["hil_description"] = "Serial Number does not exist.";
                    entity["hil_amcstagingstatus"] = false;
                }

                if (_customerAsset != null)
                {
                    qryExp = new QueryExpression("hil_warrantytemplate");
                    qryExp.ColumnSet = new ColumnSet("hil_warrantytemplateid");
                    qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    qryExp.Criteria.AddCondition(new ConditionExpression("hil_amcplan", ConditionOperator.Equal, _amcPlan.Id));
                    qryExp.Criteria.AddCondition(new ConditionExpression("hil_product", ConditionOperator.Equal, _customerAsset.GetAttributeValue<EntityReference>("hil_productsubcategory").Id));
                    qryExp.Criteria.AddCondition(new ConditionExpression("hil_validfrom", ConditionOperator.OnOrBefore, _sapBillingDate));
                    qryExp.Criteria.AddCondition(new ConditionExpression("hil_validto", ConditionOperator.OnOrAfter, _sapBillingDate));
                    qryExp.Criteria.AddCondition(new ConditionExpression("hil_templatestatus", ConditionOperator.Equal, 2)); //Approved
                    qryExp.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1)); //Active
                    enCol = service.RetrieveMultiple(qryExp);
                    if (enCol.Entities.Count > 0)
                    {
                        _warrantyTemplate = enCol.Entities[0];
                    }
                    else
                    {
                        entity["hil_description"] = "Warranty Template does not exist for selected AMC Plan.";
                        entity["hil_amcstagingstatus"] = false;
                    }
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

                    iSchWarranty.hil_warrantystartdate = _warrantyStartDate;
                    iSchWarranty.hil_warrantyenddate = _warrantyEndDate;

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
                        iSchWarranty.hil_customer = _customerAsset.GetAttributeValue<EntityReference>("hil_customer"); ;
                    }

                    iSchWarranty["hil_amcbillingdocdate"] = _sapBillingDate;
                    iSchWarranty["hil_amcbillingdocnum"] = _sapBillingDoc;
                    iSchWarranty["hil_amcbillingdocurl"] = _sapBillingDocPath;

                    service.Create(iSchWarranty);

                    #region Update Customer Asset Warranty Status
                    Entity entCustAsset = new Entity(msdyn_customerasset.EntityLogicalName);
                    entCustAsset.Id = _customerAsset.Id;

                    if (_warrantyEndDate >= DateTime.Now)
                    {
                        entCustAsset["hil_warrantystatus"] = new OptionSetValue(1); //InWarranty
                        entCustAsset["hil_warrantytilldate"] = _warrantyEndDate;
                        qryExp = new QueryExpression(hil_unitwarranty.EntityLogicalName);
                        qryExp.ColumnSet = new ColumnSet("hil_warrantystartdate", "hil_warrantyenddate", "hil_warrantytemplate");
                        qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qryExp.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, _customerAsset.Id);
                        qryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        qryExp.AddOrder("hil_warrantyenddate", OrderType.Ascending);
                        enCol = service.RetrieveMultiple(qryExp);
                        if (enCol.Entities.Count > 0)
                        {
                            foreach (Entity ent in enCol.Entities)
                            {
                                if (DateTime.Now >= ent.GetAttributeValue<DateTime>("hil_warrantystartdate").AddMinutes(330) && DateTime.Now <= ent.GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330))
                                {
                                    Entity entSubStatus = service.Retrieve("hil_warrantytemplate", ent.GetAttributeValue<EntityReference>("hil_warrantytemplate").Id, new ColumnSet("hil_type"));
                                    if (entSubStatus.GetAttributeValue<OptionSetValue>("hil_type").Value == 1)
                                    {
                                        entCustAsset["hil_warrantysubstatus"] = new OptionSetValue(1); //InWarranty-Standard
                                    }
                                    else if (entSubStatus.GetAttributeValue<OptionSetValue>("hil_type").Value == 3)
                                    {
                                        entCustAsset["hil_warrantysubstatus"] = new OptionSetValue(4); //InWarranty-AMC
                                    }
                                    service.Update(entCustAsset);
                                }
                            }
                        }
                        else
                        {
                            entCustAsset["hil_warrantystatus"] = new OptionSetValue(2); //OutWarranty
                            entCustAsset["hil_warrantytilldate"] = new DateTime(1900, 1, 1);
                            entCustAsset["hil_warrantysubstatus"] = null;
                            service.Update(entCustAsset);
                        }
                    }
                    else
                    {
                        entCustAsset["hil_warrantystatus"] = new OptionSetValue(2); //OutWarranty
                        entCustAsset["hil_warrantysubstatus"] = null;
                        service.Update(entCustAsset);
                    }
                    #endregion
                    entity["hil_description"] = "AMC Unit Warranty created successfully.";
                    entity["hil_amcstagingstatus"] = true;
                }
                //else
                //{
                //    entity["hil_description"] = "Something went wrong.";
                //    entity["hil_amcstagingstatus"] = false;
                //}
            }
            catch (Exception ex)
            {
                entity["hil_description"] = ex.Message;
                entity["hil_amcstagingstatus"] = false;
                throw new InvalidPluginExecutionException("  **Havells_Plugin.AMC_Staging.PreCreate.Execute***  " + ex.Message);
            }
        }
    }
}
