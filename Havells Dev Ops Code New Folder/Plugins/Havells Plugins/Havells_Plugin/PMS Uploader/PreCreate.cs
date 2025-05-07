using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Text.RegularExpressions;
using Havells_Plugin;

namespace Havells_Plugin.PMS_Uploader
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
                tracingService.Trace("1");
                OptionSetValue callerType = entity.GetAttributeValue<OptionSetValue>("hil_callertype");
                EntityReference callSubType = entity.GetAttributeValue<EntityReference>("hil_callsubtype");
                EntityReference customerAsset = entity.GetAttributeValue<EntityReference>("hil_associatedcustomerproduct");
                msdyn_customerasset customerAssetInfo = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, customerAsset.Id, new ColumnSet("hil_productcategory", "hil_productsubcategorymapping", "hil_productsubcategory"));
                EntityReference consumerCategory = entity.GetAttributeValue<EntityReference>("hil_consumercategory");
                EntityReference consumerType = entity.GetAttributeValue<EntityReference>("hil_consumertype");
                string customerMobileNo = entity.GetAttributeValue<string>("hil_customermobileno");
                string empAspDealerName = entity.GetAttributeValue<string>("hil_empaspdealername");
                //EntityReference natureOfCategory = entity.GetAttributeValue<EntityReference>("hil_natureofcomplaint");
                EntityReference natureOfCategory;
                string natureOfCategoryStr = entity.GetAttributeValue<string>("hil_natureofcomplaint");
                Int32 pmsCount = entity.GetAttributeValue<Int32>("hil_pmscount");
                DateTime pmsDueDate = entity.GetAttributeValue<DateTime>("hil_pmsduedate");
                tracingService.Trace("3");
                #region Checks & Validations
                #endregion

                #region Creating PMS Job
                msdyn_workorder enPMSWorkorder = new msdyn_workorder();
                QueryExpression Query = new QueryExpression("contact");
                Query.ColumnSet = new ColumnSet("fullname", "emailaddress1", "contactid");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression("mobilephone", ConditionOperator.Equal, customerMobileNo));
                EntityCollection enCol = service.RetrieveMultiple(Query);
                if (enCol.Entities.Count == 1)
                {
                    tracingService.Trace("4");
                    enPMSWorkorder.hil_CustomerRef = enCol.Entities[0].ToEntityReference();
                    if (enCol.Entities[0].Attributes.Contains("fullname"))
                    { enPMSWorkorder.hil_customername = enCol.Entities[0].GetAttributeValue<string>("fullname"); }
                    enPMSWorkorder.hil_mobilenumber = customerMobileNo;
                    if (enCol.Entities[0].Attributes.Contains("emailaddress1"))
                    { enPMSWorkorder.hil_Email = enCol.Entities[0].GetAttributeValue<string>("emailaddress1"); }
                    tracingService.Trace("5");
                    enPMSWorkorder.hil_Address = GetCustomerAddress(service, enCol.Entities[0].Id);

                    enPMSWorkorder.msdyn_CustomerAsset = customerAsset; // new EntityReference("msdyn_customerasset", customerAsset.Id);
                    enPMSWorkorder.hil_Productcategory = customerAssetInfo.hil_ProductCategory;
                    enPMSWorkorder.hil_ProductCatSubCatMapping = customerAssetInfo.hil_productsubcategorymapping;

                    enPMSWorkorder["hil_consumertype"] = consumerType;
                    enPMSWorkorder["hil_consumercategory"] = consumerCategory;
                    natureOfCategory = GetNatureOfComplaint(service, natureOfCategoryStr, customerAssetInfo.hil_ProductSubcategory.Id);
                    if (natureOfCategory != null)
                    {
                        enPMSWorkorder.hil_natureofcomplaint = natureOfCategory;
                        enPMSWorkorder["hil_newserialnumber"] = empAspDealerName;
                        enPMSWorkorder.hil_CallSubType = callSubType;
                        enPMSWorkorder.hil_quantity = 1;
                        enPMSWorkorder["hil_callertype"] = callerType;
                        enPMSWorkorder.hil_SourceofJob = new OptionSetValue(10); // {SourceofJob:"System Generated"}
                        enPMSWorkorder.msdyn_ServiceAccount = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {ServiceAccount:"Dummy Account"}
                        enPMSWorkorder.msdyn_BillingAccount = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {BillingAccount:"Dummy Account"}
                        enPMSWorkorder["hil_pmscount"] = pmsCount;
                        enPMSWorkorder["hil_pmsdate"] = pmsDueDate;
                        pmsJobId = service.Create(enPMSWorkorder);
                        tracingService.Trace("6");
                        entity["hil_pmsjobid"] = new EntityReference("msdyn_workorder", pmsJobId);
                        entity["hil_pmsjobstatus"] = true;
                        entity["hil_description"] = "PMS Job created successfully.";
                        entity["hil_pmsjobstatus"] = true;
                    }
                    else
                    {
                        entity["hil_description"] = "Nature of Complaint is not mapped with Related Product (" + customerAssetInfo.hil_ProductSubcategory.Name + ")";
                        entity["hil_pmsjobstatus"] = false;
                    }
                }
                else if (enCol.Entities.Count > 1)
                {
                    entity["hil_description"] = "Multiple customers found with Mobile No. " + customerMobileNo;
                    entity["hil_pmsjobstatus"] = false;
                }
                else
                {
                    entity["hil_description"] = "No customer found with Mobile No. " + customerMobileNo;
                    entity["hil_pmsjobstatus"] = false;
                }
                #endregion  
            }
            catch (Exception ex)
            {
                entity["hil_description"] = ex.Message;
                entity["hil_pmsjobstatus"] = false;
                throw new InvalidPluginExecutionException("  ***Havells_Plugin.PMS_Uploader.PreCreate.Execute***  " + ex.Message);
            }
        }
        public EntityReference GetCustomerAddress(IOrganizationService service, Guid customerGuid)
        {
            EntityReference retValue = null;
            QueryExpression Query;
            EntityCollection enCol;
            try
            {
                Query = new QueryExpression("msdyn_workorder")
                {
                    ColumnSet = new ColumnSet("hil_address"),
                    Criteria = new FilterExpression(LogicalOperator.And)
                };
                Query.Criteria.AddCondition(new ConditionExpression("hil_customerref", ConditionOperator.Equal, customerGuid));
                Query.TopCount = 1;
                Query.AddOrder("createdon", OrderType.Descending);
                enCol = service.RetrieveMultiple(Query);
                if (enCol.Entities.Count > 0)
                {
                    retValue = enCol.Entities[0].GetAttributeValue<EntityReference>("hil_address");
                }
                else
                {
                    Query = new QueryExpression("hil_address")
                    {
                        ColumnSet = new ColumnSet("hil_fulladdress"),
                        Criteria = new FilterExpression(LogicalOperator.And)
                    };
                    Query.Criteria.AddCondition(new ConditionExpression("hil_customer", ConditionOperator.Equal, customerGuid));
                    Query.TopCount = 1;
                    Query.AddOrder("createdon", OrderType.Descending);
                    enCol = service.RetrieveMultiple(Query);
                    if (enCol.Entities.Count > 0)
                    {
                        retValue = enCol.Entities[0].ToEntityReference();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("  ***Havells_Plugin.PMS_Uploader.PreCreate.GetCustomerAddress***  " + ex.Message);
            }
            return retValue;
        }
        public EntityReference GetNatureOfComplaint(IOrganizationService service, string noc, Guid relatedProduct)
        {
            EntityReference retValue = null;
            QueryExpression Query;
            EntityCollection enCol;
            try
            {
                Query = new QueryExpression("hil_natureofcomplaint")
                {
                    ColumnSet = new ColumnSet("hil_natureofcomplaintid"),
                    Criteria = new FilterExpression(LogicalOperator.And)
                };
                Query.Criteria.AddCondition(new ConditionExpression("hil_relatedproduct", ConditionOperator.Equal, relatedProduct));
                Query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, noc));
                Query.TopCount = 1;
                enCol = service.RetrieveMultiple(Query);
                if (enCol.Entities.Count > 0)
                {
                    retValue = enCol.Entities[0].ToEntityReference();
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("  ***Havells_Plugin.PMS_Uploader.PreCreate.GetNatureOfComplaint***  " + ex.Message);
            }
            return retValue;
        }
    }
}
