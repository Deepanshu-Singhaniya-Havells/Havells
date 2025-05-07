using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Havells.Dataverse.CustomConnector.AMC
{
    public class CreateSalesOrder : IPlugin
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

            string AMCSourceType = (string)context.InputParameters["AMC_Selling_Source"];
            if (AMCSourceType == "22") //One Website
            {
                InvoiceDelailInfo requestParm = new InvoiceDelailInfo();
                requestParm.Mobile_Number = (string)context.InputParameters["Mobile_Number"];
                requestParm.Address_Line_1 = (string)context.InputParameters["Address_Line_1"];
                requestParm.Address_Line_2 = (string)context.InputParameters["Address_Line_2"];
                requestParm.Pincode = (string)context.InputParameters["Pincode"];
                requestParm.Order_Date = (string)context.InputParameters["Order_Date"];
                requestParm.Serial_Number = (string)context.InputParameters["Serial_Number"];
                requestParm.Model_Code = (string)context.InputParameters["Model_Code"];
                requestParm.AMC_Plan = (string)context.InputParameters["AMC_Plan"];
                requestParm.AMC_Selling_Source = (string)context.InputParameters["AMC_Selling_Source"];
                requestParm.Payable_Amount = (Decimal)context.InputParameters["Payable_Amount"];
                requestParm.Payment_Status = (string)context.InputParameters["Payment_Status"];
                requestParm.Bank_Reference_Number = (string)context.InputParameters["Bank_Reference_Number"];
                requestParm.Discount = (Decimal)context.InputParameters["Discount"];
                requestParm.OrderNo = (string)context.InputParameters["OrderNo"];
                requestParm.OrderLineItem = (string)context.InputParameters["OrderLineItem"];
                requestParm.TransactionId = (string)context.InputParameters["TransactionId"];
                requestParm.MRP = (decimal)context.InputParameters["MRP"];
                context.OutputParameters["data"] = JsonSerializer.Serialize(CreateAMCSalesOrder(requestParm, service));
                return;
            }
            if (AMCSourceType == "25") // Technitian App
            {
                OrderDeatils objOrder = new OrderDeatils();
                objOrder.Customerid = (string)context.InputParameters["Customerid"];
                objOrder.Addressid = (string)context.InputParameters["Addressid"];
                objOrder.Assetid = (string)context.InputParameters["Assetid"];
                objOrder.Ordertype = (string)context.InputParameters["Ordertype"];
                objOrder.TotalPayableamount = (string)context.InputParameters["Payable_Amount"];
                objOrder.Technicianid = (string)context.InputParameters["Technicianid"];
                objOrder.Orderlinedata = (string)context.InputParameters["Orderlinedata"];
                objOrder.orderlines = JsonSerializer.Deserialize<List<Orderline>>(objOrder.Orderlinedata);
                context.OutputParameters["data"] = JsonSerializer.Serialize(CreateAMCSalesOrderFromFSM(objOrder, service));
                return;
            }
            else
            {
                context.OutputParameters["data"] = JsonSerializer.Serialize(new InvoiceResponse { status = false, message = "Invalid AMC Selling Source Type" });
                return;
            }
        }
        public InvoiceResponse CreateAMCSalesOrder(InvoiceDelailInfo requestParm, IOrganizationService _service)
        {
            try
            {
                #region Variable Declaration
                Guid OrderId = Guid.Empty;
                QueryExpression query;
                EntityCollection entCol;
                Guid AMCPlanId = Guid.Empty;
                Guid CustomerAssetId = Guid.Empty;
                Guid PincodeId = Guid.Empty;
                Guid CustomerId = Guid.Empty;
                Guid ModelCodeId = Guid.Empty;
                DateTime OrderDate;
                string EmailId = string.Empty;
                decimal DiscPer = 0.00M;
                Money baseAmount = new Money(); Money pAmount = new Money();
                Guid enAddress = new Guid();
                string[] formats = { "d/MM/yyyy", "dd/MM/yyyy", "d-MM-yyyy", "dd-MM-yyyy" };
                Regex Regex_MobileNo = new Regex("^[6-9]\\d{9}$");
                Regex Regex_PinCode = new Regex("^[1-9]([0-9]){5}$");
                EntityReference sellingSource = new EntityReference("hil_integrationsource", new Guid("608e899b-a8a3-ed11-aad1-6045bdad27a7"));//Havells - One Website
                #endregion
                if (_service != null)
                {
                    #region Check Validations
                    if (requestParm.AMC_Selling_Source != "22")
                    {
                        return new InvoiceResponse { status = false, message = "Invalid AMC selling source." };
                    }
                    if (string.IsNullOrWhiteSpace(requestParm.Mobile_Number))
                    {
                        return new InvoiceResponse { status = false, message = "Mobile Number is required." };
                    }
                    else if (!Regex_MobileNo.IsMatch(requestParm.Mobile_Number))
                    {
                        return new InvoiceResponse { status = false, message = "Invalid Mobile Number." };
                    }
                    else
                    {
                        query = new QueryExpression("contact");
                        query.ColumnSet = new ColumnSet(true);
                        query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, requestParm.Mobile_Number);
                        entCol = _service.RetrieveMultiple(query);
                        if (entCol.Entities.Count > 0)
                        {
                            EmailId = entCol.Entities[0].GetAttributeValue<string>("emailaddress1");
                            CustomerId = entCol.Entities[0].Id;
                        }
                        else
                        {
                            return new InvoiceResponse { status = false, message = "Customer not found in D365." };
                        }
                    }
                    if (string.IsNullOrWhiteSpace(requestParm.Pincode))
                    {
                        return new InvoiceResponse { status = false, message = "Pincode is required." };
                    }
                    else if (!Regex_PinCode.IsMatch(requestParm.Pincode))
                    {
                        return new InvoiceResponse { status = false, message = "Invalid Pincode." };
                    }
                    else
                    {
                        query = new QueryExpression("hil_pincode");
                        query.ColumnSet = new ColumnSet("hil_pincodeid");
                        query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, requestParm.Pincode);
                        entCol = _service.RetrieveMultiple(query);
                        if (entCol.Entities.Count > 0)
                        {
                            PincodeId = entCol.Entities[0].Id;
                        }
                        else
                        {
                            return new InvoiceResponse { status = false, message = "Pincode is not found in D365." };
                        }
                    }
                    if (string.IsNullOrWhiteSpace(requestParm.Address_Line_1))
                    {
                        return new InvoiceResponse { status = false, message = "Address Line 1 is required." };
                    }
                    if (!string.IsNullOrWhiteSpace(requestParm.Serial_Number))
                    {
                        query = new QueryExpression("msdyn_customerasset");
                        query.ColumnSet = new ColumnSet(false);
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, requestParm.Serial_Number);
                        query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, CustomerId);
                        entCol = _service.RetrieveMultiple(query);
                        if (entCol.Entities.Count > 0)
                        {
                            CustomerAssetId = entCol.Entities[0].Id;
                        }
                        else
                        {
                            return new InvoiceResponse { status = false, message = "Asset registered with another customer." };
                        }
                    }
                    if (requestParm.Payable_Amount < 0)
                    {
                        return new InvoiceResponse { status = false, message = "Payable Amount should be geater than 0." };
                    }
                    if (requestParm.Payment_Status != "2")//Paid
                    {
                        return new InvoiceResponse { status = false, message = "Invalid Payment Status." };
                    }
                    if (string.IsNullOrWhiteSpace(requestParm.Bank_Reference_Number))
                    {
                        return new InvoiceResponse { status = false, message = "Bank Reference Number is required." };
                    }
                    if (string.IsNullOrWhiteSpace(requestParm.Model_Code))
                    {
                        return new InvoiceResponse { status = false, message = "Model Code is required." };
                    }
                    else
                    {
                        query = new QueryExpression("product");
                        query.ColumnSet = new ColumnSet(false);
                        query.Criteria.AddCondition("name", ConditionOperator.Equal, requestParm.Model_Code);
                        entCol = _service.RetrieveMultiple(query);
                        if (entCol.Entities.Count > 0)
                        {
                            ModelCodeId = entCol.Entities[0].Id;
                        }
                        else
                        {
                            return new InvoiceResponse { status = false, message = "Invalid Model Code." };
                        }
                    }
                    if (string.IsNullOrWhiteSpace(requestParm.Order_Date))
                    {
                        return new InvoiceResponse { status = false, message = "Order Date is required." };
                    }
                    else if (!DateTime.TryParseExact(requestParm.Order_Date, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out OrderDate))
                    {
                        return new InvoiceResponse { status = false, message = "Date should be in the format (dd-MM-YYYY)" };
                    }
                    #endregion

                    try
                    {
                        Entity SOEntity = new Entity("salesorder");

                        query = new QueryExpression("hil_address");
                        query.ColumnSet = new ColumnSet("hil_addressid");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, CustomerId);
                        query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, PincodeId);
                        query.AddOrder("createdon", OrderType.Descending);
                        query.TopCount = 1;
                        entCol = _service.RetrieveMultiple(query);
                        if (entCol.Entities.Count > 0)
                        {
                            enAddress = entCol.Entities[0].Id;
                            SOEntity["hil_serviceaddress"] = new EntityReference("hil_address", enAddress);
                        }
                        else
                        {
                            Entity entObj = new Entity("hil_address");

                            query = new QueryExpression("hil_businessmapping");
                            query.ColumnSet = new ColumnSet("hil_pincode");
                            query.Criteria = new FilterExpression(LogicalOperator.And);
                            query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, PincodeId);
                            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                            query.AddOrder("createdon", OrderType.Descending);
                            query.TopCount = 1;
                            entCol = _service.RetrieveMultiple(query);
                            if (entCol.Entities.Count > 0)
                            {
                                entObj["hil_businessgeo"] = entCol.Entities[0].ToEntityReference();
                            }
                            entObj["hil_street1"] = requestParm.Address_Line_1;
                            if (!string.IsNullOrWhiteSpace(requestParm.Address_Line_2))
                            {
                                entObj["hil_street2"] = requestParm.Address_Line_2;
                            }
                            entObj["hil_customer"] = new EntityReference("contact", CustomerId);
                            entObj["hil_addresstype"] = new OptionSetValue(3); //Other
                            enAddress = _service.Create(entObj);
                            SOEntity["hil_serviceaddress"] = new EntityReference("hil_address", enAddress);
                        }
                        Entity customerasset = _service.Retrieve("msdyn_customerasset", CustomerAssetId, new ColumnSet("hil_invoicedate", "hil_invoiceno", "hil_invoicevalue", "hil_purchasedfrom", "hil_retailerpincode", "hil_productcategory", "hil_productsubcategory", "msdyn_product"));
                        Entity AddressCol = _service.Retrieve("hil_address", enAddress, new ColumnSet("hil_state", "hil_businessgeo", "hil_pincode", "hil_branch", "hil_salesoffice"));
                        Entity entBranch = _service.Retrieve("hil_branch", AddressCol.GetAttributeValue<EntityReference>("hil_branch").Id, new ColumnSet("hil_mamorandumcode"));
                        string _mamorandumCode = "";
                        if (entBranch.Attributes.Contains("hil_mamorandumcode"))
                        {
                            _mamorandumCode = entBranch.GetAttributeValue<string>("hil_mamorandumcode");
                        }
                        if (string.IsNullOrWhiteSpace(requestParm.AMC_Plan))
                        {
                            return new InvoiceResponse { status = false, message = "AMC Plan is required." };
                        }
                        else
                        {
                            Entity AMCPlan = _service.Retrieve("product", new Guid(requestParm.AMC_Plan), new ColumnSet(false));
                            if (AMCPlan == null)
                            {
                                return new InvoiceResponse { status = false, message = "Invalid AMC Plan." };
                            }
                            else
                            {
                                AMCPlanId = AMCPlan.Id;
                                QueryExpression AMCPlanquery = new QueryExpression("productpricelevel");
                                AMCPlanquery.ColumnSet = new ColumnSet(true);
                                AMCPlanquery.Criteria = new FilterExpression(LogicalOperator.And);
                                AMCPlanquery.Criteria.AddCondition("productid", ConditionOperator.Equal, AMCPlanId);
                                AMCPlanquery.Criteria.AddCondition("pricelevelid", ConditionOperator.Equal, PriceLevelForFGsale);
                                EntityCollection entityPricelevel = _service.RetrieveMultiple(AMCPlanquery);
                                if (entityPricelevel.Entities.Count > 0)
                                {
                                    baseAmount = entityPricelevel.Entities[0].Contains("amount_base") ? entityPricelevel.Entities[0].GetAttributeValue<Money>("amount_base") : new Money();
                                    pAmount = entityPricelevel.Entities[0].Contains("amount") ? entityPricelevel.Entities[0].GetAttributeValue<Money>("amount") : new Money();
                                    decimal MRP = decimal.Round(pAmount.Value, 2);
                                    //if (requestParm.AMC_Selling_Source == "22")
                                    // DiscPer = GetDiscountValueBySource(sellingSource.Id, _service);
                                    //else
                                    DateTime InvoiceDate = customerasset.Contains("hil_invoicedate") ? customerasset.GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330) : new DateTime(1900, 1, 1);
                                    int ProductAgeing = (int)DateTime.Now.Date.Subtract(InvoiceDate).TotalDays;
                                    var objAssetAggingValue = GetAssetWarrentyAging(CustomerAssetId, ProductAgeing, _service); //InvoiceDate
                                    DiscPer = GetDiscountValue(requestParm.Model_Code, sellingSource, customerasset, AddressCol, new Guid(requestParm.AMC_Plan), objAssetAggingValue, _service);

                                    decimal DiscountValue = Convert.ToDecimal((MRP * DiscPer) / 100);
                                    decimal Paya_Amount = decimal.Round(Convert.ToDecimal((MRP - (MRP * DiscPer) / 100)), 2);

                                    if (!Decimal.Equals(MRP, requestParm.MRP))
                                    {
                                        return new InvoiceResponse { status = false, message = "Invalid MRP amount of AMC plan." };
                                    }
                                    if (!Decimal.Equals(Paya_Amount, requestParm.Payable_Amount))
                                    {
                                        return new InvoiceResponse { status = false, message = "Invalid Payable amount of AMC plan." };
                                    }
                                    if (!Decimal.Equals(DiscountValue, requestParm.Discount))
                                    {
                                        return new InvoiceResponse { status = false, message = "Invalid Discount amount of AMC plan." };
                                    }
                                }
                                else
                                {
                                    return new InvoiceResponse { status = false, message = "Invalid AMC Plan." };
                                }
                            }

                        }

                        SOEntity["customerid"] = new EntityReference("contact", CustomerId);
                        SOEntity["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("68a6a9ca-6beb-e811-a96c-000d3af05828"));
                        SOEntity["msdyn_account"] = new EntityReference("account", new Guid("d166ba69-65da-ec11-a7b5-6045bdad2a19")); //Dummey Account
                        SOEntity["totalamount"] = new Money(Convert.ToDecimal(requestParm.MRP));
                        SOEntity["totallineitemamount"] = new Money(Convert.ToDecimal(requestParm.MRP));
                        SOEntity["hil_receiptamount"] = new Money(Convert.ToDecimal(requestParm.Payable_Amount));
                        SOEntity["hil_ordertype"] = new EntityReference("hil_ordertype", new Guid("1f9e3353-0769-ef11-a670-0022486e4abb")); //new OptionSetValue(3);//AMC Sale Ominichannel
                        SOEntity["pricelevelid"] = new EntityReference("pricelevel", PriceLevelForFGsale);// FG sale //AMC Ominichannnel
                        SOEntity["hil_source"] = new OptionSetValue(Convert.ToInt32(requestParm.AMC_Selling_Source));  //"OneWebsite|22";//requestParm.AMC_Selling_Source
                        SOEntity["msdyn_ordertype"] = new OptionSetValue(690970002);
                        if (requestParm.AMC_Selling_Source == "22")
                        {
                            SOEntity["hil_modeofpayment"] = new OptionSetValue(1);
                            SOEntity["hil_paymentstatus"] = new OptionSetValue(Convert.ToInt32(requestParm.Payment_Status));
                            SOEntity["hil_sourcereferencecode"] = requestParm.OrderNo;
                            SOEntity["hil_sellingsource"] = sellingSource;
                        }
                        OrderId = _service.Create(SOEntity);
                        ColumnSet columnSalesOrder = new ColumnSet("ordernumber", "createdon", "msdyn_psastatusreason", "name");
                        Entity entitySalesOrder = _service.Retrieve("salesorder", OrderId, columnSalesOrder);

                        if (OrderId != Guid.Empty)
                        {
                            Entity soitem = new Entity("salesorderdetail");
                            soitem["salesorderid"] = new EntityReference("salesorder", OrderId);
                            soitem["productid"] = new EntityReference("product", AMCPlanId);
                            soitem["hil_product"] = new EntityReference("product", AMCPlanId);
                            soitem["quantity"] = decimal.Round(Convert.ToDecimal(1), 2);
                            soitem["uomid"] = new EntityReference("uom", new Guid("0359d51b-d7cf-43b1-87f6-fc13a2c1dec8"));
                            soitem["ownerid"] = new EntityReference("systemuser", CustomerId);
                            soitem["hil_customerasset"] = new EntityReference("msdyn_customerasset", CustomerAssetId);
                            soitem["hil_invoicedate"] = customerasset.GetAttributeValue<DateTime>("hil_invoicedate");
                            soitem["hil_invoicenumber"] = customerasset.Contains("hil_invoiceno") ? customerasset.GetAttributeValue<string>("hil_invoiceno") : null;
                            soitem["hil_invoicevalue"] = new Money(customerasset.GetAttributeValue<decimal>("hil_invoicevalue"));
                            soitem["hil_purchasefrom"] = customerasset.Contains("hil_purchasedfrom") ? customerasset.GetAttributeValue<string>("hil_purchasedfrom") : null;
                            soitem["hil_purchasefromlocation"] = customerasset.Contains("hil_retailerpincode") ? customerasset.GetAttributeValue<string>("hil_retailerpincode") : null;
                            if (requestParm.Discount > 0)
                            {
                                soitem["hil_eligiblediscount"] = new Money(decimal.Round(Convert.ToDecimal(requestParm.Discount), 2));
                            }
                            soitem["hil_assetserialnumber"] = requestParm.Serial_Number;
                            soitem["hil_assetmodelcode"] = new EntityReference("product", ModelCodeId);
                            soitem["msdyn_quoteline"] = requestParm.OrderLineItem; // USE TO SAVE nero nimbus (NN) OrderLine Value
                            Guid orderLineID = _service.Create(soitem);

                            Entity paymentReceipt = new Entity("hil_paymentreceipt");
                            paymentReceipt["hil_mobilenumber"] = requestParm.Mobile_Number;
                            paymentReceipt["hil_email"] = EmailId;
                            paymentReceipt["hil_orderid"] = new EntityReference("salesorder", OrderId);
                            paymentReceipt["hil_amount"] = new Money(Convert.ToDecimal(requestParm.Payable_Amount));
                            paymentReceipt["hil_memorandumcode"] = _mamorandumCode;
                            paymentReceipt["hil_receiptdate"] = OrderDate;
                            paymentReceipt["hil_paymentmode"] = "Online";
                            if (requestParm.Payment_Status == "2")
                                paymentReceipt["hil_paymentstatus"] = new OptionSetValue(4);
                            else if (requestParm.Payment_Status == "3")
                                paymentReceipt["hil_paymentstatus"] = new OptionSetValue(3);
                            else
                                paymentReceipt["hil_paymentstatus"] = new OptionSetValue(2);
                            paymentReceipt["hil_bankreferenceid"] = requestParm.Bank_Reference_Number;
                            Guid receiptId = _service.Create(paymentReceipt);
                            if (requestParm.AMC_Selling_Source == "22")
                            {
                                Entity entity = new Entity("hil_paymentreceipt", receiptId);
                                entity["hil_transactionid"] = requestParm.TransactionId;   // NN TransactionId value save 
                                _service.Update(entity);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        return new InvoiceResponse { status = false, message = ex.Message };
                    }
                }
                else
                {
                    return new InvoiceResponse { status = false, message = "D365 Service Unavailable." };
                }
                return new InvoiceResponse { InvoiceGuid = OrderId.ToString(), status = true, message = "Success" };
            }
            catch (Exception ex)
            {
                return new InvoiceResponse { status = false, message = "D365 Internal Server Error : " + ex.Message };
            }
        }
        public OrderResponse CreateAMCSalesOrderFromFSM(OrderDeatils objOrder, IOrganizationService service)
        {
            OrderResponse objresposne = new OrderResponse();
            #region Validation check i.e., Inputparamters

            if (string.IsNullOrEmpty(objOrder.Customerid))
            {
                objresposne.Status = false;
                objresposne.Statusmessage = "Customerid is required.";
                return objresposne;
            }
            if (string.IsNullOrEmpty(objOrder.Addressid))
            {
                objresposne.Status = false;
                objresposne.Statusmessage = "Addressid is required.";
                return objresposne;
            }
            if (string.IsNullOrEmpty(objOrder.Assetid))
            {
                objresposne.Status = false;
                objresposne.Statusmessage = "Assetid is required.";
                return objresposne;
            }
            if (string.IsNullOrEmpty(objOrder.Ordertype))
            {
                objresposne.Status = false;
                objresposne.Statusmessage = "Ordertype is required.";
                return objresposne;
            }
            if (string.IsNullOrEmpty(objOrder.Technicianid))
            {
                objresposne.Status = false;
                objresposne.Statusmessage = "Technicianid is required.";
                return objresposne;
            }
            #endregion

            #region Salesorder&orderline Creation

            Entity SOEntity = new Entity("salesorder");
            SOEntity["customerid"] = new EntityReference("contact", new Guid(objOrder.Customerid));
            SOEntity["msdyn_psastatusreason"] = new OptionSetValue(192350000);
            SOEntity["transactioncurrencyid"] = new EntityReference("pricelevel", new Guid("68a6a9ca-6beb-e811-a96c-000d3af05828"));
            SOEntity["ownerid"] = new EntityReference("systemuser", new Guid(objOrder.Technicianid));
            SOEntity["msdyn_ordertype"] = new OptionSetValue(690970002);
            SOEntity["msdyn_account"] = new EntityReference("account", new Guid("d166ba69-65da-ec11-a7b5-6045bdad2a19")); // DummyAccount
            #region 08/11/2024 Source: TechnicianApp added  as requested by Mobile app team (Sahil Rajput)
            SOEntity["hil_source"] = new OptionSetValue(25);
            SOEntity["hil_sellingsource"] = new EntityReference("hil_integrationsource", new Guid("03b5a2d6-cc64-ed11-9562-6045bdac526a"));//Technitian App

            #endregion
            SOEntity["hil_serviceaddress"] = new EntityReference("hil_address", new Guid(objOrder.Addressid));

            #region OrderType validation

            QueryExpression ordertypequery = new QueryExpression("hil_ordertype");
            ordertypequery.ColumnSet.AddColumns("hil_ordertype", "hil_ordertypeid", "hil_pricelist");
            ordertypequery.Criteria.AddCondition("hil_ordertypeid", ConditionOperator.Equal, objOrder.Ordertype);
            EntityCollection ordertypecollection = service.RetrieveMultiple(ordertypequery);

            if (ordertypecollection.Entities.Count > 0)
            {
                string ordertype_value = ordertypecollection[0].Id.ToString();
                SOEntity["hil_ordertype"] = new EntityReference("hil_ordertype", new Guid(ordertype_value));
                SOEntity["pricelevelid"] = ordertypecollection[0].GetAttributeValue<EntityReference>("hil_pricelist");
            }
            else
            {
                objresposne.Status = false;
                objresposne.Statusmessage = "Invalid Order type.";
                return objresposne;
            }
            #endregion

            SOEntity["totalamount"] = new Money(Convert.ToDecimal(objOrder.TotalPayableamount));

            Guid orderID = service.Create(SOEntity);

            if (orderID != null)
            {
                ColumnSet columnSalesOrder = new ColumnSet("ordernumber", "createdon", "msdyn_psastatusreason", "name");
                Entity entitySalesOrder = service.Retrieve("salesorder", orderID, columnSalesOrder);
                Entity customerasset = service.Retrieve("msdyn_customerasset", new Guid(objOrder.Assetid), new ColumnSet("hil_invoicedate", "hil_invoiceno", "hil_invoicevalue", "hil_purchasedfrom", "hil_retailerpincode"));

                foreach (var line in objOrder.orderlines)
                {
                    Entity soitem = new Entity("salesorderdetail");
                    soitem["salesorderid"] = new EntityReference("salesorder", orderID);
                    soitem["productid"] = new EntityReference("product", new Guid(line.ProductID));
                    soitem["hil_product"] = new EntityReference("product", new Guid(line.ProductID));
                    soitem["quantity"] = Convert.ToDecimal(line.Quantity);
                    soitem["priceperunit"] = new Money(Convert.ToDecimal(line.PricePerUnit));
                    soitem["baseamount"] = new Money(Convert.ToDecimal(line.Amount));
                    soitem["uomid"] = new EntityReference("uom", new Guid("0359d51b-d7cf-43b1-87f6-fc13a2c1dec8"));
                    soitem["ownerid"] = new EntityReference("systemuser", new Guid(objOrder.Technicianid));
                    soitem["hil_customerasset"] = new EntityReference("msdyn_customerasset", new Guid(objOrder.Assetid));
                    soitem["hil_invoicedate"] = customerasset.GetAttributeValue<DateTime>("hil_invoicedate");
                    soitem["hil_invoicenumber"] = customerasset.Contains("hil_invoiceno") ? customerasset.GetAttributeValue<string>("hil_invoiceno") : null;
                    soitem["hil_invoicevalue"] = new Money(customerasset.GetAttributeValue<decimal>("hil_invoicevalue"));
                    soitem["hil_purchasefrom"] = customerasset.Contains("hil_purchasedfrom") ? customerasset.GetAttributeValue<string>("hil_purchasedfrom") : null;
                    soitem["hil_purchasefromlocation"] = customerasset.Contains("hil_retailerpincode") ? customerasset.GetAttributeValue<string>("hil_retailerpincode") : null;

                    Guid orderLineID = service.Create(soitem);

                    objresposne.Status = true;
                    objresposne.OrderID = orderID.ToString();
                    objresposne.Ordernumber = entitySalesOrder.GetAttributeValue<string>("name");
                    objresposne.Statusmessage = "Order Created.";
                }
            }
            else
            {
                objresposne.Status = false;
                objresposne.Statusmessage = "Failed to create Order.";
            }
            return objresposne;
            #endregion
        }
        public decimal GetDiscountValueBySource(Guid sourceId, IOrganizationService service)
        {
            decimal DiscPer = 0;
            string fetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
                              <entity name='hil_amcdiscountmatrix'>
                                 <attribute name='hil_discper' />|
                                  <order attribute='modifiedon' descending='true' />
                                  <filter type='and'>
                                    <condition attribute='statecode' operator='eq' value='0'/>
                                    <condition attribute='hil_appliedto' operator='eq' value='{sourceId}' />
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
            return DiscPer;
        }
        private WarrentyAgeingDeails GetAssetWarrentyAging(Guid CusAssetId, int ProductAgeing, IOrganizationService service)
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
            EntityCollection EntColl = service.RetrieveMultiple(new FetchExpression(xmquery));
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

        #region amc discount matrix
        public decimal GetDiscountValue(string modelName, EntityReference sellingsource, Entity entCust, Entity CustDeatils, Guid Planid, WarrentyAgeingDeails ObjAssetAgeing, IOrganizationService service)
        {
            //string PlanTypequery = string.Empty;
            //if (ModeOfPayment == 2 || ModeOfPayment == 1)
            //{
            //    PlanTypequery = $@"<condition attribute='hil_modeofpayment' operator='eq' value='{ModeOfPayment}' />";
            //}

            Guid ProductCategoryId = Guid.Empty;
            int ProductAgeing = 0;
            Guid ProductSubcategoryId = Guid.Empty;
            Guid ModelId = Guid.Empty;
            DateTime DateofPurchase;
            decimal DiscPer = 0;
            EntityReference entState = new EntityReference();
            EntityReference entSalesOffice = new EntityReference();
            if (CustDeatils.Contains("hil_state"))
                entState = CustDeatils.GetAttributeValue<EntityReference>("hil_state");
            if (CustDeatils.Contains("hil_salesoffice"))
                entSalesOffice = CustDeatils.GetAttributeValue<EntityReference>("hil_salesoffice");
            if (entCust.Contains("hil_productcategory"))
                ProductCategoryId = entCust.GetAttributeValue<EntityReference>("hil_productcategory").Id;
            if (entCust.Contains("hil_productsubcategory"))
                ProductSubcategoryId = entCust.GetAttributeValue<EntityReference>("hil_productsubcategory").Id;
            if (entCust.Contains("msdyn_product"))
                ModelId = entCust.GetAttributeValue<EntityReference>("msdyn_product").Id;

            string _applicableOn = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString().PadLeft(2, '0') + "-" + DateTime.Now.Day.ToString().PadLeft(2, '0');
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
                                          <order attribute='hil_discper' descending='true' />
                                          <filter type='and'>
                                             <condition attribute='statecode' operator='eq' value='0' />
                                             <condition attribute='hil_validfrom' operator='on-or-before' value='{_applicableOn}' />
                                             <condition attribute='hil_validto' operator='on-or-after' value='{_applicableOn}' />
                                             <condition attribute='hil_appliedto' operator='eq' value='{sellingsource.Id}' />
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
                                               <condition attribute='hil_state' operator='eq' value='{entState.Id}' />
                                             </filter>
                                             <filter type='or'>
                                               <condition attribute='hil_salesoffice' operator='null' />
                                               <condition attribute='hil_salesoffice' operator='eq' value='{{90503976-8FD1-EA11-A813-000D3AF0563C}}' />
                                               <condition attribute='hil_salesoffice' operator='eq' value='{entSalesOffice.Id}' />
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
            return DiscPer;
        }
        #endregion

        #region Models
        public class WarrentyAgeingDeails
        {
            public DateTime EndWarrentydate { get; set; }
            public int ApplicableOn { get; set; }
            public int AssetWarrentyAgeing { get; set; }
            public string WarrantyType { get; set; }
        }
        public class InvoiceDelailInfo
        {
            public string Mobile_Number { get; set; }
            public string Address_Line_1 { get; set; }
            public string Address_Line_2 { get; set; }
            public string Pincode { get; set; }
            public string Order_Date { get; set; }
            public string Serial_Number { get; set; }
            public string Model_Code { get; set; }
            public string AMC_Plan { get; set; }
            public string AMC_Selling_Source { get; set; }
            public decimal Payable_Amount { get; set; }
            public string Payment_Status { get; set; }
            public string Bank_Reference_Number { get; set; }
            public decimal Discount { get; set; }
            public string OrderNo { get; set; }
            public string OrderLineItem { get; set; }
            public string TransactionId { get; set; }
            public Decimal MRP { get; set; }
        }
        public class InvoiceResponse
        {
            public string InvoiceGuid { get; set; }
            public string message { get; set; }
            public bool status { get; set; }
        }
        public class OrderDeatils
        {
            public string Customerid { get; set; }
            public string Addressid { get; set; }
            public string Assetid { get; set; }
            public string Ordertype { get; set; }
            public string TotalPayableamount { get; set; }
            public string Technicianid { get; set; }
            public string Orderlinedata { get; set; }
            public List<Orderline> orderlines { get; set; }

        }
        public class Orderline
        {
            public string Amount { get; set; }
            public string PricePerUnit { get; set; }
            public string ProductID { get; set; }
            public string Quantity { get; set; }
        }
        public class OrderResponse
        {
            public bool Status { get; set; }
            public string OrderID { get; set; }
            public string Ordernumber { get; set; }
            public string Statusmessage { get; set; }
        }

        #endregion
    }
}
