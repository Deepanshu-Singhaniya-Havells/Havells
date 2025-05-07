using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;

namespace Havells.Dataverse.CustomConnector.TechnicianApp
{
    public class ValidateAMCDiscount : IPlugin
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
                string orderId = context.InputParameters.ContainsKey("OrderId") ? context.InputParameters["OrderId"]?.ToString() : null;
                string ReceiptAmount = context.InputParameters.ContainsKey("ReceiptAmount") ? context.InputParameters["ReceiptAmount"]?.ToString() : null;


                if (context.Depth == 1)
                {
                    ValidateAMCDiscountData data = new ValidateAMCDiscountData();
                    data.OrderId = string.IsNullOrEmpty(orderId) ? Guid.Empty : new Guid(orderId);
                    data.ReceiptAmount = string.IsNullOrEmpty(ReceiptAmount) ? 0 : decimal.Parse(ReceiptAmount);



                    ValidateAMCDiscountResponse response = ValidateDiscount(data);

                    context.OutputParameters["Status"] = response.Status;
                    context.OutputParameters["Message"] = response.Message;
                    context.OutputParameters["MessageType"] = response.MessageType;
                }
                else
                {
                    context.OutputParameters["Status"] = false;
                    context.OutputParameters["Message"] = "Invalid Depth";
                    context.OutputParameters["MessageType"] = "Error";
                }
            }
            catch (Exception ex)
            {
                context.OutputParameters["Status"] = false;
                context.OutputParameters["Message"] = ex.Message;
                context.OutputParameters["MessageType"] = "Error";
            }

        }


        public ValidateAMCDiscountResponse ValidateDiscount(ValidateAMCDiscountData _reqData)
        {
            ValidateAMCDiscountResponse _retObj = new ValidateAMCDiscountResponse();
            string _fetchXML = string.Empty;
            DateTime _invoiceDate = new DateTime(1900, 1, 1);
            int invoiceDateAgeing = 0;
            Guid productCategoryId = Guid.Empty;
            Guid productSubCategoryId = Guid.Empty;
            EntityReference modelCode = null;
            decimal discountPercent = 0;
            EntityReference PlanName = null;
            string CustomerAssetId = "";
            if (_reqData.OrderId == Guid.Empty)
            {
                _retObj = new ValidateAMCDiscountResponse() { Status = false, Message = "Order Id is required." };
                return _retObj;
            }
            if (_reqData.ReceiptAmount == null || _reqData.ReceiptAmount == 0)
            {
                _retObj = new ValidateAMCDiscountResponse() { Status = false, Message = "Receipt Amount is required." };
                return _retObj;
            }

            _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='salesorder'>
                        <attribute name='salesorderid' />
                        <attribute name='hil_productdivision' />
                        <attribute name='totallineitemamount' />
                        <attribute name='createdon' />
                        <attribute name='hil_serviceaddress' />
                        <attribute name='hil_sellingsource' />
                        <attribute name='hil_paymentstatus' />
                        <order attribute='name' descending='false' />
                        <link-entity name='hil_address' from='hil_addressid' to='hil_serviceaddress' link-type='inner' alias='ab'>
                            <attribute name='hil_state'/>
                            <attribute name='hil_salesoffice'/>
                        </link-entity>
                        <filter type='and'>
                          <condition attribute='salesorderid' operator='eq' value='{_reqData.OrderId}'/>
                        </filter>
                      </entity>
                    </fetch>";

            EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
            if (entCol.Entities.Count > 0)
            {
                EntityReference sellingsource = entCol.Entities[0].GetAttributeValue<EntityReference>("hil_sellingsource");
                EntityReference entProdCatg = entCol.Entities[0].GetAttributeValue<EntityReference>("hil_productdivision");
                EntityReference State = ((EntityReference)entCol.Entities[0].GetAttributeValue<AliasedValue>("ab.hil_state").Value);
                EntityReference SalesOffice = ((EntityReference)entCol.Entities[0].GetAttributeValue<AliasedValue>("ab.hil_salesoffice").Value);
                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='salesorderdetail'>
                        <attribute name='hil_invoicedate' />
                        <attribute name='hil_customerasset' />
                        <attribute name='productid' />
                        <order attribute='productid' descending='false' />
                        <filter type='and'>
                            <condition attribute='salesorderid' operator='eq' value='{_reqData.OrderId}'/>
                        </filter>
                        <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='hil_customerasset' visible='false' link-type='outer' alias='ca'>
                            <attribute name='hil_invoicedate' />
                            <attribute name='hil_productcategory' />
                            <attribute name='hil_productsubcategory' />
                            <attribute name='msdyn_product' />  
                        </link-entity>
                        </entity>
                        </fetch>";
                EntityCollection entColLine = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entColLine.Entities.Count > 0)
                {
                    _invoiceDate = (DateTime)(entColLine.Entities[0].GetAttributeValue<AliasedValue>("ca.hil_invoicedate").Value);
                    invoiceDateAgeing = (DateTime.Now - _invoiceDate).Days;
                    productCategoryId = ((EntityReference)entColLine.Entities[0].GetAttributeValue<AliasedValue>("ca.hil_productcategory").Value).Id;
                    productSubCategoryId = ((EntityReference)entColLine.Entities[0].GetAttributeValue<AliasedValue>("ca.hil_productsubcategory").Value).Id;
                    modelCode = ((EntityReference)entColLine.Entities[0].GetAttributeValue<AliasedValue>("ca.msdyn_product").Value);
                    PlanName = entColLine.Entities[0].GetAttributeValue<EntityReference>("productid");
                    CustomerAssetId = entColLine.Entities[0].GetAttributeValue<EntityReference>("hil_customerasset").Id.ToString();
                }
                WarrentyAgeingDeails ObjAssetAgeing = GetAssetWarrentyAging(new Guid(CustomerAssetId), invoiceDateAgeing, service);
                decimal _payableAmount = 0;
                if (entCol.Entities[0].Attributes.Contains("totallineitemamount"))
                {
                    _payableAmount = entCol.Entities[0].GetAttributeValue<Money>("totallineitemamount").Value;
                    _payableAmount = Math.Round(_payableAmount, 0);
                }
                decimal _stdDiscAmount = 0;

                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
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
                               <condition attribute='hil_validfrom' operator='on-or-before' value='{DateTime.Now.ToString("yyyy-MM-dd")}' />
                               <condition attribute='hil_validto' operator='on-or-after' value='{DateTime.Now.ToString("yyyy-MM-dd")}' />
                               <condition attribute='hil_appliedto' operator='eq' value='{sellingsource.Id}' />
                             <filter type='or'>
                              <filter type='and'>
                               <condition attribute='hil_applicableon' operator='eq' value='1' />
                               <condition attribute='hil_productaegingstart' operator='le' value='{invoiceDateAgeing}' />
                               <condition attribute='hil_productageingend' operator='ge' value='{invoiceDateAgeing}' />
                              </filter>
                              <filter type='and'>
                                <condition attribute='hil_applicableon' operator='eq' value='{ObjAssetAgeing.ApplicableOn}' />
                                <condition attribute='hil_productaegingstart' operator='le' value='{ObjAssetAgeing.AssetWarrentyAgeing}' />
                                <condition attribute='hil_productageingend' operator='ge' value='{ObjAssetAgeing.AssetWarrentyAgeing}' />
                              </filter>
                              </filter>
                             <filter type='or'>
                             <condition attribute='hil_model' operator='eq' value='{modelCode.Id}' />
                             <condition attribute='hil_productsubcategory' operator='eq' value='{productSubCategoryId}' />
                             <condition attribute='hil_productcategory' operator='eq' value='{productCategoryId}' />
                             </filter>
                             <filter type='or'>
                             <condition attribute='hil_state' operator='null' />
                             <condition attribute='hil_state' operator='eq' value='{State.Id}' />
                             </filter>
                             <filter type='or'>
                             <condition attribute='hil_salesoffice' operator='null' />
                             <condition attribute='hil_salesoffice' operator='eq' value='{{90503976-8FD1-EA11-A813-000D3AF0563C}}' />
                             <condition attribute='hil_salesoffice' operator='eq' value='{SalesOffice.Id}' />
                             </filter>
                             <filter type='or'>
                             <condition attribute='hil_product' operator='null'/>
                             <condition attribute='hil_product' operator='eq'  value='{PlanName.Id}'/>
                             </filter>
                             </filter>
                             </entity>
                             </fetch>";

                EntityCollection entCol1 = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entCol1.Entities.Count > 0)
                {
                    foreach (Entity ent in entCol1.Entities)
                    {
                        discountPercent = entCol1.Entities[0].Contains("hil_discper") ? entCol1.Entities[0].GetAttributeValue<decimal>("hil_discper") : 0;
                    }
                    _stdDiscAmount = Math.Round((_payableAmount - (_payableAmount * discountPercent) / 100), 0); //Max Limit (90)

                    if (_reqData.ReceiptAmount >= _stdDiscAmount && _reqData.ReceiptAmount <= _payableAmount)
                    {
                        _retObj = new ValidateAMCDiscountResponse() { Status = true, MessageType = "INFO", Message = "OK" };
                    }
                    else
                    {
                        _retObj = new ValidateAMCDiscountResponse() { Status = false, MessageType = "WARNING", Message = "As per AMC discount policy, you are allowed to collect between Rs." + _stdDiscAmount.ToString() + " - " + "Rs." + _payableAmount.ToString() };
                        Console.WriteLine(JsonConvert.SerializeObject(_retObj));
                    }

                }
                else
                {
                    if (_payableAmount > _reqData.ReceiptAmount)
                    {
                        _retObj = new ValidateAMCDiscountResponse() { Status = false, MessageType = "WARNING", Message = "No AMC Discount Policy is defined in System !!! Receipt Amount can't be less than Payable Amount." };
                    }
                    else if (_payableAmount < _reqData.ReceiptAmount)
                    {
                        _retObj = new ValidateAMCDiscountResponse() { Status = false, MessageType = "WARNING", Message = "No AMC Discount Policy is defined in System !!! Receipt Amount can't be greater than Payable Amount." };

                    }
                    else
                    {
                        _retObj = new ValidateAMCDiscountResponse() { Status = true, MessageType = "INFO", Message = "OK" };
                    }
                }
            }
            else
            {
                _retObj = new ValidateAMCDiscountResponse() { Status = false, MessageType = "WARNING", Message = "Order not found." };
            }
            return _retObj;

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
        public class ValidateAMCDiscountData
        {
            public Guid OrderId { get; set; }
            public Decimal? ReceiptAmount { get; set; }
        }

        public class ValidateAMCDiscountResponse
        {
            public bool Status { get; set; }
            public string MessageType { get; set; }
            public string Message { get; set; }
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
