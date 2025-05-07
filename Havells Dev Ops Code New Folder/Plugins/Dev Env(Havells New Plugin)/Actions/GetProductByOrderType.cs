using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Security.Principal;


namespace HavellsNewPlugin.Actions
{
    public class GetProductByOrderType : IPlugin
    {
        private IOrganizationService service;
        private ITracingService tracingService;
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
                string customerAssetGuid = context.InputParameters.ContainsKey("CustomerAssetGuid") ? context.InputParameters["CustomerAssetGuid"]?.ToString() : null;
                string customerAddressGuid = context.InputParameters.ContainsKey("CustomerAddressGuid") ? context.InputParameters["CustomerAddressGuid"]?.ToString() : null;
                string orderTypeGuid = context.InputParameters.ContainsKey("OrderTypeGuid") ? context.InputParameters["OrderTypeGuid"]?.ToString() : null;
                int sourceCode = context.InputParameters.ContainsKey("SourceCode") ? Convert.ToInt32(context.InputParameters["SourceCode"]) : 0;

                if (context.Depth == 1)
                {
                    GetProductByOrderTypeData data = new GetProductByOrderTypeData();
                    data.CustomerAssetGuid = string.IsNullOrEmpty(customerAssetGuid) ? Guid.Empty : new Guid(customerAssetGuid);
                    data.CustomerAddressGuid = string.IsNullOrEmpty(customerAddressGuid) ? Guid.Empty : new Guid(customerAddressGuid);
                    data.OrderTypeGuid = string.IsNullOrEmpty(orderTypeGuid) ? Guid.Empty : new Guid(orderTypeGuid);
                    data.SourceCode = sourceCode.ToString();


                    GetProductByOrderTypeResponse response = GetProductByOrder(data);

                    context.OutputParameters["Data"] = response.Data;
                    context.OutputParameters["StatusCode"] = response.StatusCode;
                    context.OutputParameters["StatusMessage"] = response.StatusMessage;
                }
                else
                {
                    context.OutputParameters["Data"] = null;
                    context.OutputParameters["StatusCode"] = false;
                    context.OutputParameters["StatusMessage"] = "Invalid depth";
                }
            }
            catch (Exception ex)
            {
                context.OutputParameters["Data"] = null;
                context.OutputParameters["StatusCode"] = false;
                context.OutputParameters["StatusMessage"] = "An error occurred: " + ex.Message;
            }

        }
        public static decimal getDiscountPercentage(EntityCollection entCol, Guid AMCPlanId)
        {
            decimal discountPer = 0;
            discountPer = entCol.Entities.Where(m => m.GetAttributeValue<EntityReference>("hil_product")?.Id == AMCPlanId).Select(m => m.GetAttributeValue<decimal>("hil_discper")).FirstOrDefault();

            if (discountPer == 0)
            {
                discountPer = entCol.Entities.OrderBy(m => m.GetAttributeValue<EntityReference>("hil_product")?.Id).Select(m => m.GetAttributeValue<decimal>("hil_discper")).FirstOrDefault();
            }
            return discountPer;
        }
        public GetProductByOrderTypeResponse GetProductByOrder(GetProductByOrderTypeData _reqData)
        {
            GetProductByOrderTypeResponse _retObj = new GetProductByOrderTypeResponse();
            string _fetchXML = string.Empty;

            if (_reqData.CustomerAssetGuid == Guid.Empty)
            {
                _retObj = new GetProductByOrderTypeResponse() { Data = null, StatusMessage = "Customer Asset Guid is required.", StatusCode = false };
                return _retObj;
            }
            if (_reqData.CustomerAddressGuid == Guid.Empty)
            {
                _retObj = new GetProductByOrderTypeResponse() { Data = null, StatusMessage = "Customer Address Guid is required.", StatusCode = false };
                return _retObj;
            }
            if (_reqData.OrderTypeGuid == Guid.Empty)
            {
                _retObj = new GetProductByOrderTypeResponse() { Data = null, StatusMessage = "Order Type Guid is required.", StatusCode = false };
                return _retObj;
            }
            if (_reqData.SourceCode == string.Empty)
            {
                _retObj = new GetProductByOrderTypeResponse() { Data = null, StatusMessage = "Source Code is required.", StatusCode = false };
                return _retObj;
            }

            Entity entityOrderType = service.Retrieve("hil_ordertype", _reqData.OrderTypeGuid, new ColumnSet("hil_ordertype"));
            string OrderType = entityOrderType.Contains("hil_ordertype") ? entityOrderType.GetAttributeValue<string>("hil_ordertype").ToLower() : "";

            if (OrderType == "amc sale")
            {
                decimal discountPercent = 0;

                int invoiceDateDaysFromToday = 0;



                Entity entityCustomer = service.Retrieve("msdyn_customerasset", _reqData.CustomerAssetGuid, new ColumnSet("hil_productcategory", "hil_productsubcategory", "msdyn_product", "hil_invoicedate"));
                Guid productCategoryId = entityCustomer.Contains("hil_productcategory") ? entityCustomer.GetAttributeValue<EntityReference>("hil_productcategory").Id : Guid.Empty;
                Guid productSubCategoryId = entityCustomer.Contains("hil_productsubcategory") ? entityCustomer.GetAttributeValue<EntityReference>("hil_productsubcategory").Id : Guid.Empty;
                EntityReference modelCode = entityCustomer.Contains("msdyn_product") ? entityCustomer.GetAttributeValue<EntityReference>("msdyn_product") : null;
                DateTime? invoiceDate = entityCustomer.GetAttributeValue<DateTime?>("hil_invoicedate");
                if (invoiceDate.HasValue)
                {
                    invoiceDateDaysFromToday = (DateTime.Now - invoiceDate.Value).Days;
                }

                Entity entityAddress = service.Retrieve("hil_address", _reqData.CustomerAddressGuid, new ColumnSet("hil_salesoffice", "hil_state"));
                EntityReference SalesOffice = entityAddress.Contains("hil_salesoffice") ? entityAddress.GetAttributeValue<EntityReference>("hil_salesoffice") : null;
                EntityReference State = entityAddress.Contains("hil_state") ? entityAddress.GetAttributeValue<EntityReference>("hil_state") : null;

                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='hil_amcdiscountmatrix'>
                                    <attribute name='hil_discper'/>
                                    <attribute name = 'hil_product'/>
                                    <order attribute='hil_name' descending='false'/>
                                    <filter type='and'>
                                        <condition attribute='hil_model' operator='eq' value='{modelCode.Id}'/>
                                        <condition attribute='hil_productcategory' operator='eq' value='{productCategoryId}'/>
                                        <condition attribute='hil_productsubcategory' operator='eq' value='{productSubCategoryId}'/>
                                        <condition attribute='hil_state' operator='eq' value='{State.Id}'/>
                                        <condition attribute='hil_salesoffice' operator='eq' value='{SalesOffice.Id}'/>
                                        <condition attribute='hil_validfrom' operator='on-or-before' value='{DateTime.Now.ToString("yyyy-MM-dd")}'/>
                                        <condition attribute='hil_validto' operator='on-or-after' value='{DateTime.Now.ToString("yyyy-MM-dd")}'/>
                                        <condition attribute='hil_productaegingstart' operator='le' value='{invoiceDateDaysFromToday}'/>
                                        <condition attribute='hil_productageingend' operator='ge' value='{invoiceDateDaysFromToday}' />
                                        <condition attribute='statecode' operator='eq' value='0'/>
                                    </filter>
                                    </entity>
                                    </fetch>";
                EntityCollection entColl = service.RetrieveMultiple(new FetchExpression(_fetchXML));


                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='hil_servicebom'>
                                    <attribute name='hil_name'/>
                                    <attribute name='createdon'/>
                                    <attribute name='hil_product'/>
                                    <attribute name='hil_servicebomid'/>
                                    <attribute name='hil_productcategory'/>
                                    <order attribute='hil_name' descending='false'/>
                                    <filter type='and'>
                                        <condition attribute='hil_productcategory' operator='eq' value='{modelCode.Id}'/>
                                        <condition attribute='statecode' operator='eq' value='0'/>
                                    </filter>
                                    <link-entity name='product' from='productid' to='hil_product' link-type='inner' alias='ac'>
                                    <filter type='and'>
                                        <condition attribute='hil_hierarchylevel' operator='eq' value='910590001'/>
                                        <condition attribute='statecode' operator='eq' value='0'/>
                                    </filter>
                                    </link-entity>
                                    </entity>
                                    </fetch>";
                EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entCol.Entities.Count > 0)
                {
                    List<AMCPlanInfo> AMCPlanInfos = new List<AMCPlanInfo>();
                    foreach (Entity item in entCol.Entities)
                    {
                        EntityReference AMCPlan = item.Contains("hil_product") ? item.GetAttributeValue<EntityReference>("hil_product") : null;

                        Entity AMCPlanMRP = service.Retrieve("product", AMCPlan.Id, new ColumnSet("hil_amount"));
                        decimal MRP = AMCPlanMRP.Contains("hil_amount") ? AMCPlanMRP.GetAttributeValue<Money>("hil_amount").Value : 0;
                        if (MRP > 0)
                        {
                            discountPercent = getDiscountPercentage(entColl, AMCPlan.Id);
                            decimal discountAmount = MRP * (discountPercent / 100);
                            decimal ListPrice = MRP - discountAmount;

                            AMCPlanInfos.Add(new AMCPlanInfo
                            {
                                ProductCode = modelCode.Name,
                                PlanGuid = AMCPlan.Id,
                                PlanName = AMCPlan.Name,
                                MRP = Math.Round(MRP, 2),
                                DiscountPercentage = Math.Round(discountPercent, 2),
                                DiscountAmount = Math.Round(discountAmount, 2),
                                LP = Math.Round(ListPrice, 2)
                            });
                        }
                    }
                    _retObj = new GetProductByOrderTypeResponse()
                    {
                        Data = JsonConvert.SerializeObject(AMCPlanInfos),
                        StatusCode = true,
                        StatusMessage = "Success"
                    };
                }
                else
                {
                    _retObj = new GetProductByOrderTypeResponse() { Data = null, StatusMessage = "No Plans Available!!", StatusCode = false };
                    return _retObj;
                }


            }
            else
            {
                _retObj = new GetProductByOrderTypeResponse() { Data = null, StatusCode = false, StatusMessage = "Order Type is not AMC Sale" };
            }
            return _retObj;
        }
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
    }
}
