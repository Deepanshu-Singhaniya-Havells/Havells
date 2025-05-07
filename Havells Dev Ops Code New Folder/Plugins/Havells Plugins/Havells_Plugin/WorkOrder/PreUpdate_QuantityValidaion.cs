using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;


namespace Havells_Plugin.WorkOrder
{
    public class PreUpdate_QuantityValidaion : IPlugin
    {
        static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "msdyn_workorder" && context.MessageName.ToUpper() == "UPDATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    msdyn_workorder enWorkorder = entity.ToEntity<msdyn_workorder>();
                    msdyn_workorder PreImageWO = ((Entity)context.PreEntityImages["image"]).ToEntity<msdyn_workorder>();

                    Guid productCategory = enWorkorder.hil_Productcategory != null ? enWorkorder.hil_Productcategory.Id : PreImageWO.hil_Productcategory != null ? PreImageWO.hil_Productcategory.Id : throw new InvalidPluginExecutionException("Product Category is Empty");
                    Guid productsubCatergory = enWorkorder.hil_ProductCatSubCatMapping != null ? enWorkorder.hil_ProductCatSubCatMapping.Id : PreImageWO.hil_ProductCatSubCatMapping != null ? PreImageWO.hil_ProductCatSubCatMapping.Id : throw new InvalidPluginExecutionException("Product Sub Category is Empty");

                    #region Refresh Product Category Brand on Job {Added by Kuldeep Khare on 10/Feb/2020}
                    OptionSetValue jobBrand = enWorkorder.hil_Brand != null ? enWorkorder.hil_Brand : PreImageWO.hil_Brand != null ? PreImageWO.hil_Brand : throw new InvalidPluginExecutionException("Product Brand is Empty");
                    OptionSetValue productCatgBrand = service.Retrieve(Product.EntityLogicalName, productCategory, new ColumnSet("hil_brandidentifier")).GetAttributeValue<OptionSetValue>("hil_brandidentifier");

                    if (jobBrand.Value != productCatgBrand.Value)
                    {
                        entity["hil_brand"] = productCatgBrand;
                    }
                    #endregion

                    hil_stagingdivisonmaterialgroupmapping sdmMapping = (hil_stagingdivisonmaterialgroupmapping)service.Retrieve(
                        hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, productsubCatergory, new ColumnSet(new string[] { "hil_productsubcategorymg", "hil_productcategorydivision" }));

                    if (entity.Attributes.Contains("hil_calculatecharges") && entity.GetAttributeValue<Boolean>("hil_calculatecharges"))
                    {
                        tracingService.Trace("1");
                        int jobQuantity = enWorkorder.Contains("hil_quantity") ? enWorkorder.hil_quantity.Value
                            : PreImageWO.Contains("hil_quantity") ? PreImageWO.hil_quantity.Value
                            : throw new InvalidPluginExecutionException("Job Quantity cannot be Empty");
                        tracingService.Trace("2");
                        if (!PreImageWO.Contains("hil_productcatsubcatmapping"))
                        {
                            throw new InvalidPluginExecutionException("Product Sub Category cannot be Empty");
                        }
                        if (PreImageWO.hil_Productcategory == null)
                        {
                            throw new InvalidPluginExecutionException("Product Category cannot be Empty");
                        }
                        tracingService.Trace("3");
                        //hil_stagingdivisonmaterialgroupmapping sdmMapping = (hil_stagingdivisonmaterialgroupmapping)service.Retrieve(
                        //    hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, PreImageWO.GetAttributeValue<EntityReference>("hil_productcatsubcatmapping").Id, new ColumnSet(new string[] { "hil_productsubcategorymg" }));
                        //
                        //if (sdmMapping.hil_ProductSubCategoryMG == null)
                        //{
                        //    throw new InvalidPluginExecutionException("Product Sub Category cannot be Empty. Please contact to Administrator");
                        //}
                        tracingService.Trace("4");
                        Product prod_Category = (Product)service.Retrieve(Product.EntityLogicalName, sdmMapping.hil_ProductSubCategoryMG.Id, new ColumnSet(new string[] { "hil_isserialized" }));

                        int isSerialized = 2;
                        if (prod_Category.hil_IsSerialized != null && prod_Category.hil_IsSerialized.Value == 1)
                        {
                            isSerialized = 1;
                        }
                        tracingService.Trace("5");
                        QueryExpression WoIncQuery = new QueryExpression()
                        {
                            EntityName = msdyn_workorderincident.EntityLogicalName,
                            ColumnSet = new ColumnSet(new string[] { "hil_quantity", "msdyn_customerasset" }),
                            Criteria =
                            {
                                Conditions =
                                {
                                    new ConditionExpression("msdyn_workorder",ConditionOperator.Equal,enWorkorder.Id),
                                    new ConditionExpression("statecode",ConditionOperator.Equal,0)
                                }
                            },
                            NoLock = true
                        };
                        EntityCollection WoIncColl = service.RetrieveMultiple(WoIncQuery);
                        tracingService.Trace("6");
                        HashSet<Guid> AssestCollection = new HashSet<Guid>();
                        int QuantitySum = 0;

                        if (WoIncColl.Entities.Count == 0)
                        {
                            throw new InvalidPluginExecutionException("At least 1 Job Incident is required to close the job.");
                        }

                        if (WoIncColl.Entities != null && WoIncColl.Entities.Count > 0)
                        {
                            foreach (msdyn_workorderincident WoInc in WoIncColl.Entities)
                            {
                                if (WoInc.msdyn_CustomerAsset != null)
                                {
                                    AssestCollection.Add(WoInc.msdyn_CustomerAsset.Id);
                                }
                                if (WoInc.hil_Quantity != null)
                                {
                                    QuantitySum += WoInc.hil_Quantity.Value;
                                }
                            }
                        }
                        if (isSerialized == 1 && jobQuantity != AssestCollection.Count)
                        {
                            throw new InvalidPluginExecutionException("For Serialized Product Job Quantity must be Equals to Total No. of Assests.");
                        }
                        else if (isSerialized == 2 && AssestCollection.Count > jobQuantity)
                        {
                            throw new InvalidPluginExecutionException("For Serialized Product Job Quantity must be Equals to or less than Total No. of Assests.");
                        }
                        else if (isSerialized == 2 && QuantitySum != jobQuantity)
                        {
                            throw new InvalidPluginExecutionException("Total Incident Quantity must be equals to Job Quantity.");
                        }
                        else
                        {
                            foreach (Guid Asset in AssestCollection)
                            {
                                msdyn_customerasset cus_Assest = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, Asset, new ColumnSet(new string[] { "msdyn_product", "hil_productcategory", "hil_invoicedate", "hil_invoiceno", "hil_productsubcategorymapping" }));
                                tracingService.Trace("2");
                                if (cus_Assest.hil_ProductCategory == null || cus_Assest.msdyn_Product == null || cus_Assest.hil_productsubcategorymapping == null
                                    || cus_Assest.hil_ProductCategory.Id != PreImageWO.hil_Productcategory.Id)
                                {
                                    throw new InvalidPluginExecutionException("The Customer Asset Category combination should match with Job Incident.");
                                }
                                tracingService.Trace("3");
                                hil_stagingdivisonmaterialgroupmapping sdmMappingInc = (hil_stagingdivisonmaterialgroupmapping)service.Retrieve(
                                            hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, cus_Assest.GetAttributeValue<EntityReference>("hil_productsubcategorymapping").Id, new ColumnSet(new string[] { "hil_productsubcategorymg" }));
                                tracingService.Trace("4");
                                if (sdmMappingInc.hil_ProductSubCategoryMG == null
                                    || sdmMappingInc.hil_ProductSubCategoryMG.Id != sdmMapping.hil_ProductSubCategoryMG.Id)
                                {
                                    throw new InvalidPluginExecutionException("The Customer Asset category combination should match with Job. Please contact to Administrator");
                                }
                            }
                        }
                        Entity _jobEntity = service.Retrieve("msdyn_workorder", entity.Id, new ColumnSet("hil_callsubtype"));
                        if (_jobEntity != null)
                        {
                            string _callsubType = _jobEntity.GetAttributeValue<EntityReference>("hil_callsubtype").Id.ToString();
                            if (_callsubType == JobCallSubType.AMCCall)
                            {
                                //Added by Kuldeep Khare 26/Oct/2023 to validate AMC Product must used in AMC Job 
                                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='msdyn_workorderproduct'>
                            <attribute name='createdon' />
                            <attribute name='msdyn_product' />
                            <attribute name='msdyn_linestatus' />
                            <attribute name='msdyn_description' />
                            <attribute name='msdyn_workorderproductid' />
                            <order attribute='msdyn_product' descending='false' />
                            <filter type='and'>
                              <condition attribute='msdyn_workorder' operator='eq' value='{entity.Id}' />
                              <condition attribute='hil_replacedpart' operator='not-null' />
                              <condition attribute='hil_markused' operator='eq' value='1' />
                              <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            <link-entity name='product' from='productid' to='hil_replacedpart' link-type='inner' alias='ag'>
                              <filter type='and'>
                                <condition attribute='hil_hierarchylevel' operator='eq' value='910590001' />
                              </filter>
                            </link-entity>
                          </entity>
                        </fetch>";
                                EntityCollection entColAMCProd = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                if (entColAMCProd.Entities.Count == 0)
                                {
                                    throw new InvalidPluginExecutionException("Please consume atleast one AMC Product.");
                                }
                            }
                        }
                    }
                    //Added by Kuldeep khare 30/Oct/2019 {KKG OTP Decode}
                    // && PreImageWO.hil_KKGOTP != null
                    if (enWorkorder.hil_kkgcode != null)
                    {
                        //if ((enWorkorder.hil_kkgcode != PreImageWO.hil_KKGOTP) && (enWorkorder.hil_kkgcode != Havells_Plugin.WorkOrder.Common.Base64Decode(PreImageWO.hil_KKGOTP)))
                        //{
                        //    throw new InvalidPluginExecutionException(" ***Given KKG code does not match Job KKG OTP.*** ");
                        //}
                        bool? validated = KKGCodeHashing.GenerateKKGCodeHashVerification(PreImageWO.msdyn_name, enWorkorder.hil_kkgcode, service);
                        if (validated == null || validated == false)
                        {
                            //if ((enWorkorder.hil_kkgcode != PreImageWO.hil_KKGOTP) && (enWorkorder.hil_kkgcode != Havells_Plugin.WorkOrder.Common.Base64Decode(PreImageWO.hil_KKGOTP)))
                            //{
                            throw new InvalidPluginExecutionException(" ***Given KKG code does not match Job KKG OTP.*** ");
                            //}
                        }
                        //else if (validated == false)
                        //{
                        //    throw new InvalidPluginExecutionException(" ***!!! Given KKG code does not match Job KKG OTP.*** ");
                        //}
                    }

                    if (entity.Attributes.Contains("hil_closeticket") && entity.GetAttributeValue<Boolean>("hil_closeticket"))
                    {
                        // Checking for Mandatory Fields
                        Common.PreJobClosureValidations(service, entity.Id);

                        string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_sawactivityapproval'>
                        <attribute name='hil_sawactivityapprovalid' />
                        <attribute name='hil_name' />
                        <attribute name='hil_jobid' />
                        <order attribute='hil_name' descending='false' />
                        <filter type='and'>
                            <condition attribute='hil_jobid' operator='eq' value='{" + entity.Id + @"}' />
                            <condition attribute='hil_approvalstatus' operator='not-in'>
                            <value>3</value>
                            <value>4</value>
                            </condition>
                        </filter>
                        <link-entity name='hil_sawactivity' from='hil_sawactivityid' to='hil_sawactivity' link-type='inner' alias='ao'>
                            <link-entity name='hil_serviceactionwork' from='hil_serviceactionworkid' to='hil_sawcategory' link-type='inner' alias='ap'>
                            <filter type='and'>
                                <condition attribute='hil_mandatoryforjobclosure' operator='eq' value='1' />
                            </filter>
                            </link-entity>
                        </link-entity>
                        </entity>
                        </fetch>";
                        EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entCol.Entities.Count > 0)
                        {
                            string _approvalIds = string.Empty;
                            foreach (Entity ent in entCol.Entities)
                            {
                                _approvalIds += ent.GetAttributeValue<string>("hil_name") + ",";
                            }
                            throw new InvalidPluginExecutionException(" ***Please get your following open SAW reviewed before closing the Job. \n " + _approvalIds + " *** ");
                        }
                    }

                    if (entity.Attributes.Contains("hil_kkgcode_sms"))
                    {
                        OptionSetValue _optValue = entity.GetAttributeValue<OptionSetValue>("hil_kkgcode_sms");
                        if (_optValue.Value == 100000001) //Work not done (Reopen-New complain no.)
                        {
                            string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_sawactivityapproval'>
                            <attribute name='hil_sawactivityapprovalid' />
                            <attribute name='hil_name' />
                            <attribute name='hil_jobid' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                                <condition attribute='hil_jobid' operator='eq' value='{" + entity.Id + @"}' />
                                <condition attribute='hil_approvalstatus' operator='not-in'>
                                <value>3</value>
                                <value>4</value>
                                </condition>
                            </filter>
                            <link-entity name='hil_sawactivity' from='hil_sawactivityid' to='hil_sawactivity' link-type='inner' alias='ao'>
                                <link-entity name='hil_serviceactionwork' from='hil_serviceactionworkid' to='hil_sawcategory' link-type='inner' alias='ap'>
                                <filter type='and'>
                                    <condition attribute='hil_mandatoryforjobclosure' operator='eq' value='1' />
                                </filter>
                                </link-entity>
                            </link-entity>
                            </entity>
                            </fetch>";
                            EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (entCol.Entities.Count > 0)
                            {
                                string _approvalIds = string.Empty;
                                foreach (Entity ent in entCol.Entities)
                                {
                                    _approvalIds += ent.GetAttributeValue<string>("hil_name") + ",";
                                }
                                throw new InvalidPluginExecutionException(" ***Please get your following open SAW reviewed before closing the Job. \n " + _approvalIds + " *** ");
                            }
                        }
                    }

                    #region Added by Saurabh Tripathi on 13/Jan/2021 Validate AMC Receipt Amount
                    if (entity.Contains("hil_receiptamount"))
                    {
                        Decimal reciptAmount = entity.GetAttributeValue<Money>("hil_receiptamount").Value;
                        Common.ValidateAMCReceiptAmount(service, entity.Id, reciptAmount);
                    }
                    if (entity.Contains("hil_isgascharged"))
                    {
                        Entity _jobEnt = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("msdyn_substatus", "hil_callsubtype", "hil_productcategory"));
                        var gasChargeStatus = entity.GetAttributeValue<Boolean>("hil_isgascharged");
                        var subStatus = _jobEnt.GetAttributeValue<EntityReference>("msdyn_substatus");
                        var callSubType = _jobEnt.GetAttributeValue<EntityReference>("hil_callsubtype");
                        var prodCategory = _jobEnt.GetAttributeValue<EntityReference>("hil_productcategory");
                        tracingService.Trace("***************************");
                        tracingService.Trace("Job Id " + entity.Id);

                        tracingService.Trace("Status " + subStatus.Name + " callSubType " + callSubType.Id + " prd Cat " + prodCategory.Id + " Gas Char " + gasChargeStatus);

                        if ((gasChargeStatus) && subStatus.Name == "Work Done" && (callSubType.Id == new Guid(JobCallSubType.PMS) ||
                            callSubType.Id == new Guid(JobCallSubType.Breakdown) || callSubType.Id == new Guid(JobCallSubType.DealerStockRepair))
                            && (prodCategory.Id == new Guid(JobProductCategory.LLOYDAIRCONDITIONER) || prodCategory.Id == new Guid(JobProductCategory.LLOYDREFRIGERATORS)))
                        {
                            // do nothing
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException(" ***Gas Charge Not Allowed*** ");
                        }
                    }

                    #endregion
                }
            }
            catch (InvalidPluginExecutionException e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
            #endregion
        }
        public void checkForBrandProdCategory(IOrganizationService service, Guid productCategory, msdyn_workorder enWorkorder)
        {
            QueryExpression Query = new QueryExpression("product");
            Query.ColumnSet = new ColumnSet("hil_brandidentifier");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("productid", ConditionOperator.Equal, productCategory));
            EntityCollection Found1 = service.RetrieveMultiple(Query);
            if (Found1.Entities.Count > 0)
            {
                foreach (Entity ent in Found1.Entities)
                {
                    if (((OptionSetValue)enWorkorder.hil_Brand).Value != ((OptionSetValue)ent.Attributes["hil_brandidentifier"]).Value)
                    {
                        throw new InvalidPluginExecutionException("Job Brand cant not be differenct from Product Category Brand. Please contact to Administrator");
                    }
                }
            }
        }
    }
}
