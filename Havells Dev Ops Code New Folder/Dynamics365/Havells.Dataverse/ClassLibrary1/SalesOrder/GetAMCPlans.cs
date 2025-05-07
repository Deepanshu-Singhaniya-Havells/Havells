using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace Havells.Dataverse.CustomConnector.SalesOrder
{
    public class GetAMCPlans : IPlugin
    {
        private IOrganizationService service;
        private ITracingService tracingService;
        private static Guid PriceLevelForFGsale = new Guid("c05e7837-8fc4-ed11-b597-6045bdac5e78");//AMC Ominichannnel
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion

            try
            {
                string customerAssetGuid = Convert.ToString(context.InputParameters["CustomerAssetGuid"]);
                string customerAddressGuid = Convert.ToString(context.InputParameters["CustomerAddressGuid"]);
                string orderTypeGuid = Convert.ToString(context.InputParameters["OrderTypeGuid"]);
                int sourceCode = Convert.ToInt32(context.InputParameters["SourceCode"]);
                string invoiceDate = Convert.ToString(context.InputParameters["InvoiceDate"]);
                string modelCode = Convert.ToString(context.InputParameters["ModelCode"]);

                GetProductByOrderTypeData data = new GetProductByOrderTypeData();
                data.CustomerAssetGuid = APValidate.IsvalidGuid(customerAssetGuid) ? new Guid(customerAssetGuid) : Guid.Empty;
                data.CustomerAddressGuid = APValidate.IsvalidGuid(customerAddressGuid) ? new Guid(customerAddressGuid) : Guid.Empty;
                data.OrderTypeGuid = APValidate.IsvalidGuid(orderTypeGuid) ? new Guid(orderTypeGuid) : Guid.Empty;
                data.SourceCode = sourceCode.ToString();

                GetProductByOrderTypeResponse response = GetAMCPlan(data, service);

                context.OutputParameters["Data"] = response.Data;
                context.OutputParameters["StatusCode"] = response.StatusCode;
                context.OutputParameters["StatusMessage"] = response.StatusMessage;
            }
            catch (Exception ex)
            {
                context.OutputParameters["Data"] = null;
                context.OutputParameters["StatusCode"] = false;
                context.OutputParameters["StatusMessage"] = "An error occurred: " + ex.Message;
            }

        }
        public GetProductByOrderTypeResponse GetAMCPlan(GetProductByOrderTypeData AMCPlanParam, IOrganizationService _crmService)
        {
            GetProductByOrderTypeResponse AMCPlanDtls = new GetProductByOrderTypeResponse();
            try
            {
                QueryExpression query;
                Guid ProductCategoryId = Guid.Empty;
                EntityReference Model = null;
                Guid ProductSubcategoryId = Guid.Empty;
                DateTime InvoiceDate = new DateTime(1900, 1, 1);
                int ProductAgeing = 0;

                List<AMCPlanInfo> lstAMCPlanInfo = new List<AMCPlanInfo>();
                if (_crmService != null)
                {
                    query = new QueryExpression("hil_integrationsource");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, AMCPlanParam.SourceCode);
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);//Active
                    EntityCollection sourceEntColl = _crmService.RetrieveMultiple(query);
                    if (sourceEntColl.Entities.Count == 0)
                    {
                        AMCPlanDtls = new GetProductByOrderTypeResponse() { Data = null, StatusMessage = "Invalid Souce Code.", StatusCode = false };
                        return AMCPlanDtls;
                    }
                    Guid sourceId = sourceEntColl.Entities[0].Id;
                    Entity entityCustomerAsset = _crmService.Retrieve("msdyn_customerasset", AMCPlanParam.CustomerAssetGuid, new ColumnSet("msdyn_product", "hil_productcategory", "hil_invoicedate", "hil_productsubcategory"));
                    if (entityCustomerAsset != null)
                    {
                        ProductCategoryId = entityCustomerAsset.Contains("hil_productcategory") ? entityCustomerAsset.GetAttributeValue<EntityReference>("hil_productcategory").Id : Guid.Empty;
                        Model = entityCustomerAsset.Contains("msdyn_product") ? entityCustomerAsset.GetAttributeValue<EntityReference>("msdyn_product") : null;
                        InvoiceDate = entityCustomerAsset.Contains("hil_invoicedate") ? entityCustomerAsset.GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330) : new DateTime(1900, 1, 1);
                        ProductSubcategoryId = entityCustomerAsset.Contains("hil_productsubcategory") ? entityCustomerAsset.GetAttributeValue<EntityReference>("hil_productsubcategory").Id : Guid.Empty;
                    }
                    if (ProductCategoryId == Guid.Empty || Model == null || ProductSubcategoryId == Guid.Empty)
                    {
                        AMCPlanDtls = new GetProductByOrderTypeResponse() { Data = null, StatusMessage = "Category/Sub-Category/Model is missing.", StatusCode = false };
                        return AMCPlanDtls;
                    }
                    if (AMCPlanParam.CustomerAddressGuid == Guid.Empty)
                    {
                        AMCPlanDtls = new GetProductByOrderTypeResponse() { Data = null, StatusMessage = "Address is required.", StatusCode = false };
                        return AMCPlanDtls;
                    }
                    Entity entCustomerAddress = _crmService.Retrieve("hil_address", AMCPlanParam.CustomerAddressGuid, new ColumnSet("hil_state", "hil_salesoffice"));
                    Guid stateId = entCustomerAddress.Contains("hil_state") ? entCustomerAddress.GetAttributeValue<EntityReference>("hil_state").Id : Guid.Empty;
                    Guid salesofficeId = entCustomerAddress.Contains("hil_salesoffice") ? entCustomerAddress.GetAttributeValue<EntityReference>("hil_salesoffice").Id : Guid.Empty;

                    ProductAgeing = (int)DateTime.Now.Date.Subtract(InvoiceDate).TotalDays;

                    string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                    <entity name='hil_amcplansetup'>
                                    <attribute name='hil_amcplansetupid' />
                                    <attribute name='hil_amcplan' />
                                    <order attribute='hil_amcplan' descending='false' />
                                    <filter type='and'>
                                    <condition attribute='statecode' operator='eq' value='0' />
                                    <condition attribute='hil_model' operator='eq' value='{Model.Id}' />
                                    <filter type='or'>
                                    <condition attribute='hil_applicablesource' operator='eq' value='{sourceId}' />
                                    <condition attribute='hil_applicablesource' operator='null' />
                                    </filter>
                                    <filter type='or'>
                                    <condition attribute='hil_ageingstart' operator='null' />
                                    <filter type='and'>
                                    <condition attribute='hil_ageingstart' operator='le' value='{ProductAgeing}' />
                                    <condition attribute='hil_ageingend' operator='ge' value='{ProductAgeing}' />
                                    </filter>
                                    </filter>
                                    </filter>
                                    <link-entity name='product' from='productid' to='hil_model' link-type='inner' alias='be'>
                                    <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='bf'>
                                    <filter type='and'>
                                    <condition attribute='hil_eligibleforamc' operator='eq' value='1' />
                                    </filter>
                                    </link-entity>
                                    </link-entity>
                                    <link-entity name='product' from='productid' to='hil_amcplan' link-type='inner' alias='bg'>
                                    <filter type='and'>
                                    <condition attribute='hil_hierarchylevel' operator='eq' value='910590001' />
                                    </filter>
                                    <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='pc'>
                                    <attribute name='hil_name' />   
                                    <attribute name='createdon' />   
                                    <attribute name='hil_plantclink' />   
                                    <attribute name='hil_planperiod' />   
                                    <attribute name='hil_notcovered' />
                                    <attribute name='hil_coverage' />
                                    <attribute name='hil_amctandc' />
                                    </link-entity>
                                    <link-entity name='productpricelevel' from='productid' to='productid' link-type='inner' alias='pricelist'>
                                    <attribute name='amount' />
                                    <filter type='and'>
                                    <condition attribute='pricelevelid' operator='eq' value='{PriceLevelForFGsale}' />
                                    </filter>
                                    </link-entity>
                                    </link-entity>
                                    </entity>
                                    </fetch>";

                    EntityCollection entCollProduct = _crmService.RetrieveMultiple(new FetchExpression(fetchXML));

                    if (entCollProduct.Entities.Count > 0)
                    {
                        var objAssetAggingValue = GetAssetWarrentyAging(AMCPlanParam.CustomerAssetGuid, ProductAgeing, _crmService); //InvoiceDate

                        foreach (Entity entProduct in entCollProduct.Entities)
                        {
                            AMCPlanInfo objAMCPlanInfo = new AMCPlanInfo();

                            if (entProduct.Contains("hil_amcplan"))
                            {
                                objAMCPlanInfo.ProductCode = Model.Name;
                                objAMCPlanInfo.PlanGuid = entProduct.GetAttributeValue<EntityReference>("hil_amcplan").Id;
                                objAMCPlanInfo.DiscountPercentage = GetDiscountValue(_crmService, Model.Id, sourceId, ProductAgeing, ProductCategoryId, ProductSubcategoryId, objAMCPlanInfo.PlanGuid, stateId, salesofficeId, objAssetAggingValue);
                                objAMCPlanInfo.MRP = decimal.Round((entProduct.Contains("pricelist.amount") ? ((Money)entProduct.GetAttributeValue<AliasedValue>("pricelist.amount").Value).Value : 0), 2);
                                objAMCPlanInfo.PlanName = entProduct.Contains("pc.hil_name") ? entProduct.GetAttributeValue<AliasedValue>("pc.hil_name").Value.ToString() : "";
                                objAMCPlanInfo.DiscountAmount = objAMCPlanInfo.DiscountPercentage > 0 ? Math.Round((objAMCPlanInfo.MRP * objAMCPlanInfo.DiscountPercentage) / 100, 2) : 0;
                                objAMCPlanInfo.LP = objAMCPlanInfo.MRP - objAMCPlanInfo.DiscountAmount;
                            }
                            lstAMCPlanInfo.Add(objAMCPlanInfo);
                        }
                    }
                    AMCPlanDtls = new GetProductByOrderTypeResponse()
                    {
                        Data = JsonConvert.SerializeObject(lstAMCPlanInfo),
                        StatusCode = true,
                        StatusMessage = "Success"
                    };
                    return AMCPlanDtls;
                }
                else
                {
                    AMCPlanDtls = new GetProductByOrderTypeResponse() { Data = null, StatusMessage = "D365 Service unavailable.", StatusCode = false };
                    return AMCPlanDtls;
                }
            }
            catch (Exception ex)
            {
                AMCPlanDtls = new GetProductByOrderTypeResponse() { Data = null, StatusMessage = "D365 Error: " + ex.Message, StatusCode = false };
                return AMCPlanDtls;
            }
        }
        public decimal GetDiscountValue(IOrganizationService service, Guid ModelId, Guid sourceId, int ProductAgeing, Guid ProductCategoryId, Guid ProductSubcategoryId, Guid Planid, Guid entState, Guid entSalesOffice, WarrentyAgeingDeails ObjAssetAgeing)
        {
            decimal DiscPer = 0.00M;
            string _applicableOn = DateTime.Now.ToString("yyyy-MM-dd");
            string fetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
                            <entity name='hil_amcdiscountmatrix'>
                            <attribute name='hil_discper' />
                            <order attribute='hil_applicableon' descending='false' />
                            <order attribute='hil_product' descending='true' />
                            <order attribute='hil_salesoffice' descending='true' />
                            <order attribute='hil_state' descending='true' />
                            <order attribute='hil_model' descending='true' />
                            <order attribute='hil_productcategory' descending='true' />
                            <order attribute='hil_productsubcategory' descending='true' />
                            <filter type='and'>
                               <condition attribute='statecode' operator='eq' value='0' />
                               <condition attribute='hil_validfrom' operator='on-or-before' value='{_applicableOn}' />
                               <condition attribute='hil_validto' operator='on-or-after' value='{_applicableOn}' />
                               <condition attribute='hil_appliedto' operator='eq' value='{sourceId}' />
                             <filter type='or'>
                              <filter type='and'>
                               <condition attribute='hil_applicableon' operator='eq' value='1' />
                               <condition attribute='hil_productaegingstart' operator='le' value='{ProductAgeing}' />
                               <condition attribute='hil_productageingend' operator='ge' value='{ProductAgeing}' />
                              </filter>
                              <filter type='and'>
                                <condition attribute='hil_applicableon' operator='eq' value='{ObjAssetAgeing.ApplicableOn}' />
                                <condition attribute='hil_productaegingstart' operator='le' value='{ObjAssetAgeing.AssetWarrentyAgeing}' />
                                <condition attribute='hil_productageingend' operator='ge' value='{ObjAssetAgeing.AssetWarrentyAgeing}' />
                              </filter>
                              </filter>   
                              <filter type='or'>
                              <condition attribute='hil_productcategory' operator='eq' value='{ProductCategoryId}' />
                              <condition attribute='hil_productcategory' operator='null' />
                              </filter>
                              <filter type='or'>
                              <condition attribute='hil_productsubcategory' operator='eq' value='{ProductSubcategoryId}' />
                              <condition attribute='hil_productsubcategory' operator='null' />
                              </filter>
                              <filter type='or'>
                              <condition attribute='hil_model' operator='eq' value='{ModelId}' />
                              <condition attribute='hil_model' operator='null' />
                              </filter>
                              <filter type='or'>
                              <condition attribute='hil_state' operator='null' />
                              <condition attribute='hil_state' operator='eq' value='{entState}' />
                              </filter>
                              <filter type='or'>
                              <condition attribute='hil_salesoffice' operator='null' />
                              <condition attribute='hil_salesoffice' operator='eq' value='{{90503976-8FD1-EA11-A813-000D3AF0563C}}' />
                              <condition attribute='hil_salesoffice' operator='eq' value='{entSalesOffice}' />
                              </filter>
                              <filter type='or'>
                              <condition attribute='hil_product' operator='eq' value='{Planid}' />
                              <condition attribute='hil_product' operator='null' />
                              </filter>
                              </filter>
                              </entity>
                              </fetch>";

            EntityCollection entamcdiscountmatrix = service.RetrieveMultiple(new FetchExpression(fetchQuery));
            if (entamcdiscountmatrix.Entities.Count > 0)
            {
                if (entamcdiscountmatrix.Entities[0].Contains("hil_discper"))
                {
                    DiscPer = entamcdiscountmatrix.Entities[0].GetAttributeValue<Decimal>("hil_discper");
                }
            }
            return Math.Round(DiscPer, 2);
        }
        #region Get Asset Warrenty Aging
        private WarrentyAgeingDeails GetAssetWarrentyAging(Guid CusAssetId, int ProductAgeing, IOrganizationService _crmService)
        {
            var objAgeing = new WarrentyAgeingDeails();
            objAgeing.ApplicableOn = 1;
            objAgeing.AssetWarrentyAgeing = ProductAgeing;
            string xmquery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
                    <entity name='hil_unitwarranty'>
                    <attribute name='hil_name'/>
                    <attribute name='hil_warrantytemplate'/>
                    <attribute name='hil_warrantystartdate'/>
                    <attribute name='hil_warrantyenddate'/>
                    <attribute name='hil_producttype'/>
                    <attribute name='hil_customerasset'/>
                    <attribute name='hil_unitwarrantyid'/>
                    <order attribute='hil_warrantyenddate' descending='true'/>
                    <filter type='and'>
                    <condition attribute='hil_customerasset' operator='eq' value='{CusAssetId}'/>
                    </filter>
                    <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='ac'>
                    <attribute name='hil_warrantytypeindex'/>
                    </link-entity>
                    </entity>
                    </fetch>";
            EntityCollection EntColl = _crmService.RetrieveMultiple(new FetchExpression(xmquery));
            if (EntColl.Entities.Count > 0)
            {
                DateTime EndWarrentydate = EntColl.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330);
                int AssetWarrentyAgeing = (int)DateTime.Now.Date.Subtract(EndWarrentydate).TotalDays;
                EntityReference Entywarrantytype = (EntityReference)EntColl.Entities[0].GetAttributeValue<AliasedValue>("ac.hil_warrantytypeindex").Value; // Warranty Type( Standrad Warranty, AMC Warranty)
                if (Entywarrantytype.Name.ToUpper() == "AMC" || Entywarrantytype.Name.ToUpper() == "STANDARD")//amc 
                {
                    if (AssetWarrentyAgeing <= 0)
                    {
                        objAgeing.ApplicableOn = 2; //Pre AMC Expiry
                        objAgeing.AssetWarrentyAgeing = Math.Abs(AssetWarrentyAgeing);
                    }
                    else
                    {
                        objAgeing.ApplicableOn = 3;  //Post AMC Expiry
                        objAgeing.AssetWarrentyAgeing = Math.Abs(AssetWarrentyAgeing);
                    }
                }
            }
            return objAgeing;
        }
        #endregion

        public class GetProductByOrderTypeData
        {
            public Guid CustomerAssetGuid { get; set; }
            public Guid CustomerAddressGuid { get; set; }
            public Guid OrderTypeGuid { get; set; }
            public string SourceCode { get; set; }
        }
        public class GetProductByOrderTypeResponse
        {
            public bool StatusCode { get; set; }
            public string StatusMessage { get; set; }
            public string Data { get; set; }
        }
        public class AMCPlanInfo
        {
            public Guid PlanGuid { get; set; }
            public string PlanName { get; set; }
            public string ProductCode { get; set; }
            public decimal MRP { get; set; }
            public decimal DiscountPercentage { get; set; }
            public decimal DiscountAmount { get; set; }
            public decimal LP { get; set; }
        }
     
        #region Model WarrentyDiscount
        public class WarrentyAgeingDeails
        {
            public DateTime EndWarrentydate { get; set; }
            public int ApplicableOn { get; set; }
            public int AssetWarrentyAgeing { get; set; }
            public string WarrantyType { get; set; }
        }

        #endregion
    }
}
