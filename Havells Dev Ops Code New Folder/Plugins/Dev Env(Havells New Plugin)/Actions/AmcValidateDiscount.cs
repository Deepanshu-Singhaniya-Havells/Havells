
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace HavellsNewPlugin.Actions
{


    public class AmcValidateDiscount : IPlugin
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
                if (context.InputParameters.Contains("OrderId") && context.InputParameters["OrderId"] is string
                    && context.InputParameters.Contains("ReceiptAmount") && context.InputParameters["ReceiptAmount"] is string
                    && context.Depth == 1)
                {
                    string salesOrderId = context.InputParameters["OrderId"].ToString();
                    string receiptAmount = context.InputParameters["ReceiptAmount"].ToString();
                    
                    
                    ValidateDiscountData data = new ValidateDiscountData();
                    data.OrderId = new Guid(salesOrderId); 
                    data.ReceiptAmount = decimal.Parse(receiptAmount);

                    ValidateDiscountResponse response = ValidateDiscount(data); 

                    context.OutputParameters["Status"] = response.ResultStatus;
                    context.OutputParameters["Message"] = response.ResultMessage;
                    context.OutputParameters["MessageType"] = response.ResultMessageType;

                }
            }
            catch (Exception ex)
            {
                context.OutputParameters["Status"] = false;
                context.OutputParameters["Message"] = ex.Message;
                context.OutputParameters["MessageType"] = "Error";
            }

        }


        public ValidateDiscountResponse ValidateDiscount(ValidateDiscountData _reqData)
        {
            ValidateDiscountResponse _retObj = new ValidateDiscountResponse();
            string _fetchXML = string.Empty;
            DateTime _invoiceDate = new DateTime(1900, 1, 1);

            if (_reqData.OrderId == Guid.Empty)
            {
                _retObj = new ValidateDiscountResponse() { ResultStatus = false, ResultMessage = "Order Id is required." };
                return _retObj;
            }
            if (_reqData.ReceiptAmount == null || _reqData.ReceiptAmount == 0)
            {
                _retObj = new ValidateDiscountResponse() { ResultStatus = false, ResultMessage = "Receipt Amount is required." };
                return _retObj;
            }
            _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='salesorder'>
                <attribute name='salesorderid' />
                <attribute name='hil_productdivision' />
                <attribute name='totallineitemamount' />
                <attribute name='createdon' />
                <attribute name='hil_paymentstatus' />
                <order attribute='name' descending='false' />
                <filter type='and'>
                  <condition attribute='salesorderid' operator='eq' value='{_reqData.OrderId}' />
                </filter>
              </entity>
            </fetch>";

            EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
            if (entCol.Entities.Count > 0)
            {
                EntityReference entProdCatg = entCol.Entities[0].GetAttributeValue<EntityReference>("hil_productdivision");
                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='salesorderdetail'>
                    <attribute name='hil_invoicedate' />
                    <attribute name='hil_customerasset' />
                    <order attribute='productid' descending='false' />
                    <filter type='and'>
                        <condition attribute='salesorderid' operator='eq' value='{_reqData.OrderId}' />
                    </filter>
                    <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='hil_customerasset' visible='false' link-type='outer' alias='ca'>
                        <attribute name='hil_invoicedate' />
                    </link-entity>
                    </entity>
                    </fetch>";
                EntityCollection entColLine = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entColLine.Entities[0].Attributes.Contains("hil_customerasset") && entColLine.Entities[0].Attributes.Contains("ca.hil_invoicedate"))
                    _invoiceDate = (DateTime)(entColLine.Entities[0].GetAttributeValue<AliasedValue>("ca.hil_invoicedate").Value);
                else if (entColLine.Entities[0].Attributes.Contains("hil_invoicedate"))
                    _invoiceDate = entColLine.Entities[0].GetAttributeValue<DateTime>("hil_invoicedate");


                decimal _payableAmount = 0;
                if (entCol.Entities[0].Attributes.Contains("totallineitemamount"))
                {
                    _payableAmount = entCol.Entities[0].GetAttributeValue<Money>("totallineitemamount").Value;
                }

                DateTime _orderDate = entCol.Entities[0].GetAttributeValue<DateTime>("createdon").AddMinutes(330);

                string _asOn = _orderDate.Year.ToString() + "-" + _orderDate.Month.ToString().PadLeft(2, '0') + "-" + _orderDate.Day.ToString().PadLeft(2, '0');
                int _dayDiff = Convert.ToInt32(Math.Round((_orderDate - _invoiceDate).TotalDays, 0));
                if (_dayDiff < 0)
                {
                    return new ValidateDiscountResponse() { ResultStatus = false, ResultMessageType = "WARNING", ResultMessage = "Product Age is -4 days. Job is created prior to Asset Invoice Date." };
                }
                decimal _stdDiscPer = 0;
                decimal _spcDiscPer = 0;
                decimal _stdDiscAmount = 0;
                decimal _spcDiscAmount = 0;

                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_amcdiscountmatrix'>
                <attribute name='hil_amcdiscountmatrixid' />
                <attribute name='hil_name' />
                <attribute name='createdon' />
                <attribute name='hil_discounttype' />
                <attribute name='hil_discper' />
                <order attribute='hil_name' descending='false' />
                <filter type='and'>
                    <condition attribute='statecode' operator='eq' value='0' />
                    <condition attribute='hil_appliedto' operator='eq' value='{03B5A2D6-CC64-ED11-9562-6045BDAC526A}' />
                    <condition attribute='hil_productaegingstart' operator='le' value='" + _dayDiff.ToString() + @"' />
                    <condition attribute='hil_productageingend' operator='ge' value='" + _dayDiff.ToString() + @"' />
                    <condition attribute='hil_validfrom' operator='on-or-before' value='" + _asOn + @"' />
                    <condition attribute='hil_validto' operator='on-or-after' value='" + _asOn + @"' />
                    <condition attribute='hil_productcategory' operator='eq' value='{" + entProdCatg.Id + @"}' />
                </filter>
                </entity>
                </fetch>";

                EntityCollection entCol1 = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entCol1.Entities.Count > 0)
                {
                    foreach (Entity ent in entCol1.Entities)
                    {
                        if (ent.GetAttributeValue<OptionSetValue>("hil_discounttype").Value == 1)
                        {
                            _stdDiscPer = ent.GetAttributeValue<decimal>("hil_discper");
                        }
                        else
                        {
                            _spcDiscPer = ent.GetAttributeValue<decimal>("hil_discper");
                        }
                    }
                    _stdDiscAmount = Math.Round((_payableAmount - (_payableAmount * _stdDiscPer) / 100), 0); //Max Limit (90)
                    _spcDiscAmount = Math.Round(_payableAmount - (_payableAmount * (_stdDiscPer + _spcDiscPer)) / 100, 0); //Min Limit (85)
                    if (_reqData.ReceiptAmount >= _spcDiscAmount && _reqData.ReceiptAmount < _stdDiscAmount)
                    {
                        decimal _additionaldisc = Math.Round(_stdDiscAmount - Convert.ToDecimal(_reqData.ReceiptAmount), 2);
                        _retObj = new ValidateDiscountResponse() { ResultStatus = false, ResultMessageType = "CONFIRMATION", ResultMessage = "To offer additional discount (Rs. " + _additionaldisc.ToString() + ") above Standard Discount, you need to take BSH approval. Click 'Yes' if approval already taken Or Click 'No'." };
                    }
                    else if (_reqData.ReceiptAmount < _spcDiscAmount)
                    {
                        _retObj = new ValidateDiscountResponse() { ResultStatus = false, ResultMessageType = "WARNING", ResultMessage = "As per AMC discount policy, you are allowed to collect minimum Rs. " + _stdDiscAmount.ToString() + "." };
                    }
                    else
                    {
                        _retObj = new ValidateDiscountResponse() { ResultStatus = true, ResultMessageType = "INFO", ResultMessage = "OK" };
                    }
                }
                else
                {
                    if (_payableAmount != _reqData.ReceiptAmount)
                    {
                        _retObj = new ValidateDiscountResponse() { ResultStatus = false, ResultMessageType = "WARNING", ResultMessage = "No AMC Discount Policy is defined in System !!! Receipt Amount can't be less than Payable Amount." };
                    }
                    else
                    {
                        _retObj = new ValidateDiscountResponse() { ResultStatus = true, ResultMessageType = "INFO", ResultMessage = "OK" };
                    }
                }
            }
            else
            {
                _retObj = new ValidateDiscountResponse() { ResultStatus = false, ResultMessageType = "WARNING", ResultMessage = "Order not found." };
            }
            return _retObj;
        }


        public class ValidateDiscountData
        {
            public Guid OrderId { get; set; }
            public Decimal? ReceiptAmount { get; set; }
            public int? SourceCode { get; set; }
        }

        public class ValidateDiscountResponse
        {
            public bool ResultStatus { get; set; }
            public string ResultMessage { get; set; }
            public string ResultMessageType { get; set; }
        }
    }
}
