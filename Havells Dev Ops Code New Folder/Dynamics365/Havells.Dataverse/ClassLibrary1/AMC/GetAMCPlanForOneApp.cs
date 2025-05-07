using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.AMC
{
    public class GetAMCPlanForOneApp : IPlugin
    {
        private static Guid PriceLevelForFGsale = new Guid("c05e7837-8fc4-ed11-b597-6045bdac5e78");//AMC Ominichannnel
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion            

            string[] Source = { "6", "16" }; //OneApp & One Website

            string LoginUserId = Convert.ToString(context.InputParameters["LoginUserId"]);
            string UserToken = Convert.ToString(context.InputParameters["UserToken"]);

            string jsonString = Convert.ToString(context.InputParameters["reqdata"]);
            var data = JsonSerializer.Deserialize<AMCPlanParam>(jsonString);
            string SourceType = data.SourceType;
            string CustomerAssestId = data.CustomerAssestId;
            string ModelNumber = data.ModelNumber;
            string AddressId = data.AddressId;
            if (!APValidate.IsvalidGuid(CustomerAssestId))
            {
                string msg = string.IsNullOrWhiteSpace(CustomerAssestId) ? "Customer Assest Id required" : "Invalid Customer Assest Id.";
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = msg });
                context.OutputParameters["data"] = responnse;
                return;
            }
            if (string.IsNullOrWhiteSpace(ModelNumber))
            {
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = "Model number required." });
                context.OutputParameters["data"] = responnse;
                return;
            }
            if (!APValidate.isAlphaNumeric(ModelNumber))
            {
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = "Please enter valid Model Number." });
                context.OutputParameters["data"] = responnse;
                return;
            }
            if (!APValidate.IsvalidGuid(AddressId))
            {
                string msg = string.IsNullOrWhiteSpace(AddressId) ? "Address Guid required" : "Invalid Address Guid";
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = msg });
                context.OutputParameters["data"] = responnse;
                return;
            }
            if (!Source.Contains(SourceType))
            {
                string msg = string.IsNullOrWhiteSpace(SourceType) ? "Source Type is required." : "Invalid Source Type.";
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = msg });
                context.OutputParameters["data"] = responnse;
                return;
            }
            AMCPlanParam objAMCplan = new AMCPlanParam();
            objAMCplan.AddressId = AddressId;
            objAMCplan.ModelNumber = ModelNumber;
            objAMCplan.CustomerAssestId = CustomerAssestId;
            objAMCplan.SourceType = SourceType;

            var response = GetAMCPlans(service, objAMCplan, LoginUserId);
            dynamic result;

            if (response.Item2.StatusCode == (int)HttpStatusCode.OK)
                result = JsonSerializer.Serialize(response.Item1);
            else
                result = JsonSerializer.Serialize(response.Item2);
            context.OutputParameters["data"] = result;
            return;
        }
        public (AMCPlanRes, RequestStatus) GetAMCPlans(IOrganizationService _crmService, AMCPlanParam AMCPlanParam, string LoginUserId)
        {
            AMCPlanRes AMCPlanDtls = new AMCPlanRes();
            try
            {
                QueryExpression query;
                Guid ProductCategoryId = Guid.Empty;
                Guid Model = Guid.Empty;
                Guid ProductSubcategoryId = Guid.Empty;
                DateTime InvoiceDate = new DateTime(1900, 1, 1);
                int ProductAgeing = 0;

                List<AMCPlanInfo> lstAMCPlanInfo = new List<AMCPlanInfo>();
                if (_crmService != null)
                {
                    query = new QueryExpression("hil_integrationsource");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, AMCPlanParam.SourceType);
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);//Active
                    EntityCollection sourceEntColl = _crmService.RetrieveMultiple(query);
                    if (sourceEntColl.Entities.Count == 0)
                    {
                        return (AMCPlanDtls, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = "Invalid Source Type."
                        });
                    }
                    Guid sourceId = sourceEntColl.Entities[0].Id;
                    Entity entityCustomerAsset = _crmService.Retrieve("msdyn_customerasset", new Guid(AMCPlanParam.CustomerAssestId), new ColumnSet("msdyn_product", "hil_productcategory", "hil_invoicedate", "hil_productsubcategory"));
                    if (entityCustomerAsset != null)
                    {
                        ProductCategoryId = entityCustomerAsset.Contains("hil_productcategory") ? entityCustomerAsset.GetAttributeValue<EntityReference>("hil_productcategory").Id : Guid.Empty;
                        Model = entityCustomerAsset.Contains("msdyn_product") ? entityCustomerAsset.GetAttributeValue<EntityReference>("msdyn_product").Id : Guid.Empty;
                        InvoiceDate = entityCustomerAsset.Contains("hil_invoicedate") ? entityCustomerAsset.GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330) : new DateTime(1900, 1, 1);
                        ProductSubcategoryId = entityCustomerAsset.Contains("hil_productsubcategory") ? entityCustomerAsset.GetAttributeValue<EntityReference>("hil_productsubcategory").Id : Guid.Empty;
                    }
                    if (ProductCategoryId == Guid.Empty || Model == Guid.Empty || ProductSubcategoryId == Guid.Empty)
                    {
                        return (AMCPlanDtls, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = "Category/Sub-Category/Model is missing."
                        });
                    }
                    Guid AddressId = Guid.Empty;
                    if (!Guid.TryParse(AMCPlanParam.AddressId, out AddressId))
                    {
                        return (AMCPlanDtls, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = "Address is required."
                        });
                    }
                    Entity entCustomerAddress = _crmService.Retrieve("hil_address", AddressId, new ColumnSet("hil_state", "hil_salesoffice"));
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
                                    <condition attribute='hil_model' operator='eq' value='{Model}' />
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
                        var objAssetAggingValue = GetAssetWarrentyAging(new Guid(AMCPlanParam.CustomerAssestId), ProductAgeing, _crmService); //InvoiceDate
                        foreach (Entity entProduct in entCollProduct.Entities)
                        {
                            AMCPlanInfo objAMCPlanInfo = new AMCPlanInfo();

                            if (entProduct.Contains("hil_amcplan"))
                            {
                                objAMCPlanInfo.PlanId = entProduct.GetAttributeValue<EntityReference>("hil_amcplan").Id;
                                objAMCPlanInfo.DiscountPercent = GetDiscountValue(_crmService, Model, sourceId, ProductAgeing, ProductCategoryId, ProductSubcategoryId, objAMCPlanInfo.PlanId, stateId, salesofficeId, objAssetAggingValue).ToString();
                                objAMCPlanInfo.MRP = decimal.Round((entProduct.Contains("pricelist.amount") ? ((Money)entProduct.GetAttributeValue<AliasedValue>("pricelist.amount").Value).Value : 0), 2).ToString();

                                objAMCPlanInfo.Coverage = entProduct.Contains("pc.hil_coverage") ? entProduct.GetAttributeValue<AliasedValue>("pc.hil_coverage").Value.ToString() : "";
                                objAMCPlanInfo.NonCoverage = entProduct.Contains("pc.hil_notcovered") ? entProduct.GetAttributeValue<AliasedValue>("pc.hil_notcovered").Value.ToString() : "";
                                objAMCPlanInfo.PlanName = entProduct.Contains("pc.hil_name") ? entProduct.GetAttributeValue<AliasedValue>("pc.hil_name").Value.ToString() : "";
                                objAMCPlanInfo.PlanPeriod = entProduct.Contains("pc.hil_planperiod") ? entProduct.GetAttributeValue<AliasedValue>("pc.hil_planperiod").Value.ToString() : "";
                                objAMCPlanInfo.PlanTCLink = entProduct.Contains("pc.hil_plantclink") ? entProduct.GetAttributeValue<AliasedValue>("pc.hil_plantclink").Value.ToString() : "";
                            }
                            lstAMCPlanInfo.Add(objAMCPlanInfo);
                        }
                    }
                    AMCPlanDtls.ModelNumber = AMCPlanParam.ModelNumber;
                    AMCPlanDtls.AMCPlanInfo = lstAMCPlanInfo;
                    RequestStatus requestStatus = new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.OK,
                        Message = "Success",
                    };
                    AMCPlanDtls.StatusCode = requestStatus.StatusCode;
                    return (AMCPlanDtls, requestStatus);
                }
                else
                {
                    return (AMCPlanDtls, new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.ServiceUnavailable,
                        Message = "D365 service unavailable."
                    });
                }
            }
            catch (Exception ex)
            {
                return (AMCPlanDtls, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = "D365 internal server error : " + ex.Message.ToUpper()
                });
            }
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
        #region Get AMC Plan
        public class AMCPlanParam
        {
            public string ModelNumber { get; set; }
            public string SourceType { get; set; }
            public string AddressId { get; set; }
            public string CustomerAssestId { get; set; }
        }
        public class AMCPlanInfo
        {
            public Guid PlanId { get; set; }
            public string PlanName { get; set; }
            public string PlanPeriod { get; set; }
            public string MRP { get; set; }
            public string DiscountPercent { get; set; }
            public string EffectivePrice
            {
                get
                {
                    if (DiscountPercent != null)
                    {
                        return decimal.Round(Convert.ToDecimal(MRP) - ((Convert.ToDecimal(MRP) * Convert.ToDecimal(DiscountPercent)) / 100), 2).ToString();
                    }
                    else
                    {
                        return decimal.Round(Convert.ToDecimal(MRP), 2).ToString();
                    }
                }
            }
            public string Coverage { get; set; }
            public string NonCoverage { get; set; }
            public string PlanTCLink { get; set; }
        }
        public class AMCPlanRes : TokenExpires
        {
            public string ModelNumber { get; set; }
            public List<AMCPlanInfo> AMCPlanInfo { get; set; }
        }

        #endregion
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
