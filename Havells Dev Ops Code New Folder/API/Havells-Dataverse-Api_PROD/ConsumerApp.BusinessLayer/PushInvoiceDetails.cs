using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
//using System.Transactions;
using System.Web.Services.Description;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class PushInvoiceDetails
    {
        private static Guid PriceLevelForFGsale = new Guid("c05e7837-8fc4-ed11-b597-6045bdac5e78");//AMC Ominichannnel
        public InvoiceResponse InsertInvoiceDetail(InvoiceDelailInfo requestParm)
        {
            Guid invoiceID = Guid.Empty;
            try
            {
                #region Variable Declaration
                QueryExpression query;
                EntityCollection entCol;
                EntityReference businessGeo = null;
                Guid AMCPlanId = Guid.Empty;
                Guid CustomerAssetId = Guid.Empty;
                Guid PincodeId = Guid.Empty;
                Guid CustomerId = Guid.Empty;
                Guid ModelCodeId = Guid.Empty;
                DateTime OrderDate;
                decimal DiscPer = 0.00M;
                string[] formats = { "d/MM/yyyy", "dd/MM/yyyy", "d-MM-yyyy", "dd-MM-yyyy" };
                Regex Regex_MobileNo = new Regex("^[6-9]\\d{9}$");
                Regex Regex_PinCode = new Regex("^[1-9]([0-9]){5}$");
                #endregion
                IOrganizationService _service = ConnectToCRM.GetOrgService();
                if (_service != null)
                {
                    #region Check Validations


                    if (string.IsNullOrWhiteSpace(requestParm.Mobile_Number))
                    {
                        return new InvoiceResponse { status = false, message = "Mobile Number is required." };
                    }
                    else if (!Regex_MobileNo.IsMatch(requestParm.Mobile_Number))
                    {
                        return new InvoiceResponse { status = false, message = "Mobile Number is Invalid." };
                    }
                    else
                    {
                        query = new QueryExpression("contact");
                        query.ColumnSet = new ColumnSet(false);
                        query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, requestParm.Mobile_Number);
                        entCol = _service.RetrieveMultiple(query);
                        if (entCol.Entities.Count > 0)
                        {
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
                            query = new QueryExpression("productpricelevel");
                            query.ColumnSet = new ColumnSet("amount");
                            query.Criteria.AddCondition("productid", ConditionOperator.Equal, AMCPlanId);
                            query.Criteria.AddCondition("pricelevelid", ConditionOperator.Equal, PriceLevelForFGsale);
                            Entity pAmount = _service.RetrieveMultiple(query).Entities[0];
                            decimal MRP = decimal.Round((pAmount.Contains("amount") ? (pAmount.GetAttributeValue<Money>("amount").Value) : 0), 2);
                            DiscPer = Math.Round((MRP - requestParm.Payable_Amount) / 100, 2);
                        }
                        //query = new QueryExpression("product");
                        //query.ColumnSet = new ColumnSet(false);
                        //query.Criteria = new FilterExpression(LogicalOperator.And);
                        //query.Criteria.AddCondition("name", ConditionOperator.Equal, requestParm.AMC_Plan);
                        //query.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 910590001);//AMC
                        //entCol = _service.RetrieveMultiple(query);
                        //if (entCol.Entities.Count == 0)
                        //{
                        //    return new InvoiceResponse { status = false, message = "Invalid AMC Plan." };
                        //}
                        //else
                        //{
                        //    AMCPlanId = entCol.Entities[0].Id;
                        //    query = new QueryExpression("productpricelevel");
                        //    query.ColumnSet = new ColumnSet("amount");
                        //    query.Criteria.AddCondition("productid", ConditionOperator.Equal, AMCPlanId);
                        //    query.Criteria.AddCondition("pricelevelid", ConditionOperator.Equal, PriceLevelForFGsale);
                        //    Entity pAmount = _service.RetrieveMultiple(query).Entities[0];
                        //    decimal MRP = decimal.Round((pAmount.Contains("amount") ? (pAmount.GetAttributeValue<Money>("amount").Value) : 0), 2);
                        //    DiscPer = Math.Round((MRP - requestParm.Payable_Amount) / 100, 2);
                        //}
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
                    //using (var txscope = new TransactionScope(TransactionScopeOption.RequiresNew))
                    //{
                    try
                    {
                        Entity invoiceentity = new Entity("invoice");

                        query = new QueryExpression("hil_address");
                        query.ColumnSet = new ColumnSet("hil_addressid");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, CustomerId);
                        query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, PincodeId);
                        entCol = _service.RetrieveMultiple(query);
                        if (entCol.Entities.Count > 0)
                        {
                            invoiceentity["hil_address"] = new EntityReference("hil_address", entCol.Entities[0].Id);
                        }
                        else
                        {
                            query = new QueryExpression("hil_businessmapping");
                            query.ColumnSet = new ColumnSet("hil_pincode");
                            query.Criteria = new FilterExpression(LogicalOperator.And);
                            query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, PincodeId);
                            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                            query.AddOrder("createdon", OrderType.Ascending);
                            query.TopCount = 1;
                            entCol = _service.RetrieveMultiple(query);
                            if (entCol.Entities.Count > 0)
                            {
                                businessGeo = entCol.Entities[0].ToEntityReference();
                            }
                            Entity entObj = new Entity("hil_address");
                            entObj["hil_street1"] = requestParm.Address_Line_1;
                            if (!string.IsNullOrWhiteSpace(requestParm.Address_Line_2))
                            {
                                entObj["hil_street2"] = requestParm.Address_Line_2;
                            }
                            entObj["hil_customer"] = new EntityReference("contact", CustomerId);
                            if (businessGeo != null)
                            {
                                entObj["hil_businessgeo"] = businessGeo;
                            }
                            entObj["hil_addresstype"] = new OptionSetValue(3); //Other
                            invoiceentity["hil_address"] = new EntityReference("hil_address", _service.Create(entObj));
                        }

                        if (!string.IsNullOrWhiteSpace(requestParm.Serial_Number))
                        {
                            invoiceentity["hil_customerasset"] = new EntityReference("msdyn_customerasset", CustomerAssetId);
                            invoiceentity["hil_newserialnumber"] = requestParm.Serial_Number;
                        }

                        invoiceentity["msdyn_invoicedate"] = OrderDate;
                        invoiceentity["customerid"] = new EntityReference("contact", CustomerId);
                        invoiceentity["hil_modelcode"] = new EntityReference("product", ModelCodeId);
                        invoiceentity["hil_productcode"] = new EntityReference("product", AMCPlanId);
                        invoiceentity["hil_mobilenumber"] = requestParm.Mobile_Number;
                        invoiceentity["hil_receiptamount"] = new Money(requestParm.Payable_Amount);
                        invoiceentity["hil_salestype"] = new OptionSetValue(3);//AMC Sale Ominichannel
                        invoiceentity["pricelevelid"] = new EntityReference("pricelevel", PriceLevelForFGsale);// FG sale //AMC Ominichannnel
                        invoiceentity["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("68a6a9ca-6beb-e811-a96c-000d3af05828"));
                        invoiceentity["hil_amcsellingsource"] = "OneWebsite|22";//requestParm.AMC_Selling_Source
                        invoiceentity["statecode"] = new OptionSetValue(2); // Paid requestParm.Payment_Status
                        invoiceentity["statuscode"] = new OptionSetValue(100001); // Complete
                        invoiceentity["hil_orderid"] = requestParm.OrderNo + "-" + requestParm.OrderLineItem;

                        invoiceID = _service.Create(invoiceentity);

                        Entity _entInvoiceInfo = _service.Retrieve("invoice", invoiceID, new ColumnSet("name"));
                        string _invoiceNo = string.Empty;

                        if (_entInvoiceInfo != null)
                        {
                            if (_entInvoiceInfo.Contains("name"))
                                _invoiceNo = _entInvoiceInfo.GetAttributeValue<string>("name");
                        }

                        Entity paymententity = new Entity("msdyn_payment");
                        paymententity["msdyn_name"] = requestParm.TransactionId; //_invoiceNo;
                        paymententity["msdyn_amount"] = new Money(requestParm.Payable_Amount);
                        paymententity["msdyn_paymenttype"] = new OptionSetValue(690970003);//
                        paymententity["msdyn_date"] = OrderDate;
                        Guid paymentId = _service.Create(paymententity);

                        Entity paymentdetailentity = new Entity("msdyn_paymentdetail");
                        paymentdetailentity["msdyn_invoice"] = new EntityReference("invoice", invoiceID);
                        paymentdetailentity["statuscode"] = new OptionSetValue(910590000); // Received
                        paymentdetailentity["msdyn_paymentamount"] = new Money(requestParm.Payable_Amount);
                        paymentdetailentity["msdyn_name"] = requestParm.TransactionId; //_invoiceNo;
                        paymentdetailentity["msdyn_payment"] = new EntityReference("msdyn_payment", paymentId);
                        paymentdetailentity["hil_paymentmode"] = new OptionSetValue(1);//Online
                        _service.Create(paymentdetailentity);

                        Entity statusPayment = new Entity("hil_paymentstatus");
                        statusPayment["hil_bank_ref_num"] = requestParm.Bank_Reference_Number;
                        statusPayment["hil_amt"] = requestParm.Payable_Amount.ToString();
                        statusPayment["hil_transaction_amount"] = requestParm.Payable_Amount.ToString();
                        statusPayment["hil_name"] = requestParm.TransactionId; //_invoiceNo;
                        statusPayment["hil_disc"] = requestParm.Discount.ToString();
                        statusPayment["hil_phone"] = requestParm.Mobile_Number;
                        statusPayment["hil_productinfo"] = requestParm.AMC_Plan;
                        statusPayment["hil_net_amount_debit"] = requestParm.Payable_Amount.ToString();
                        statusPayment["hil_paymentstatus"] = "success";

                        if (_entInvoiceInfo != null)
                        {
                            if (_entInvoiceInfo.Contains("name"))
                                statusPayment["hil_invoiceid"] = _entInvoiceInfo.GetAttributeValue<string>("name");
                        }
                        _service.Create(statusPayment);
                    }
                    catch (Exception ex)
                    {
                        return new InvoiceResponse { status = false, message = ex.Message };
                    }
                    Entity UpdateInvoice = new Entity("invoice", invoiceID);
                    UpdateInvoice["discountpercentage"] = requestParm.Discount;
                    //UpdateInvoice["name"] = requestParm.OrderNo;
                    UpdateInvoice["statecode"] = new OptionSetValue(2); // Paid requestParm.Payment_Status
                    UpdateInvoice["statuscode"] = new OptionSetValue(100001); // Complete
                    _service.Update(UpdateInvoice);
                }
                else
                {
                    return new InvoiceResponse { status = false, message = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                return new InvoiceResponse { status = false, message = "D365 Internal Server Error : " + ex.Message };
            }
            return new InvoiceResponse { InvoiceGuid = invoiceID.ToString(), status = true, message = "Success" };
        }

        public InvoiceResponse InsertInvoiceDetail(global::InvoiceDetail request)
        {
            throw new NotImplementedException();
        }

        #region Sync Product/AMC Plan
        public ModelDetailsList SyncProductList(string syncDateTime)
        {
            ModelDetailsList ModelDetailsList = new ModelDetailsList();
            List<ModelDetails> lstModelDetails = new List<ModelDetails>();
            string _error = string.Empty;
            DateTime _startDate = DateTime.Now;
            string _requestData = JsonConvert.SerializeObject(syncDateTime);
            string _responseData = string.Empty;
            DateTime _syncDateTime;
            IOrganizationService _service = ConnectToCRM.GetOrgService();

            if (string.IsNullOrEmpty(syncDateTime))
            {
                ModelDetailsList.Result = new ResResult { ResultStatus = false, ResultMessage = "No Content : Sync Date time is required." };
                return ModelDetailsList;
            }
            if (!DateTime.TryParseExact(syncDateTime, "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out _syncDateTime))
            {
                ModelDetailsList.Result = new ResResult { ResultStatus = false, ResultMessage = "No Content : Invalid Sync Datetime format. Required format : yyyy-MM-dd HH:mm" };
                return ModelDetailsList;
            }
            syncDateTime = _syncDateTime.ToString("yyyy-MM-dd HH:mm");
            try
            {
                string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true' {{0}}>
                            <entity name='product'>
                            <attribute name='name' />
                            <attribute name='productnumber' />
                            <attribute name='description' />
                            <attribute name='statecode' />
                            <attribute name='modifiedon'/>
                            <attribute name='productstructure' />
                            <attribute name='productid' />
                            <attribute name='hil_materialgroup' />
                            <attribute name='hil_division' />
                            <order attribute='productnumber' descending='false' />
                            <filter type='and'>
                                <condition attribute='modifiedon' operator='gt' value='{syncDateTime}' />
                            </filter>
                                <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='al'>
                                <filter type='and'>
                                    <condition attribute='hil_eligibleforamc' operator='eq' value='1' />
                                </filter>
                                </link-entity>
                            </entity>
                            </fetch>";

                //EntityCollection entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                EntityCollection entcoll = RetrieveAllRecords(_service, fetchXml);
                if (entcoll.Entities.Count == 0)
                {
                    ModelDetailsList.Result = new ResResult { ResultStatus = false, ResultMessage = "Not Any Eligible AMC Product found" };
                    return ModelDetailsList;
                }
                else
                {
                    foreach (Entity ent in entcoll.Entities)
                    {
                        ModelDetails objModelDetails = new ModelDetails();
                        if (ent.Attributes.Contains("name"))
                        {
                            objModelDetails.ModelNumber = ent.GetAttributeValue<string>("name");
                            objModelDetails.ModelId = ent.Id;
                        }
                        if (ent.Attributes.Contains("description"))
                        {
                            objModelDetails.ModelName = ent.GetAttributeValue<string>("description");
                        }
                        if (ent.Attributes.Contains("hil_division"))
                        {
                            objModelDetails.ProductCategory = ent.GetAttributeValue<EntityReference>("hil_division").Name;
                            objModelDetails.ProductCategoryId = ent.GetAttributeValue<EntityReference>("hil_division").Id;
                        }
                        if (ent.Attributes.Contains("hil_materialgroup"))
                        {
                            objModelDetails.ProductSubcategory = ent.GetAttributeValue<EntityReference>("hil_materialgroup").Name;
                            objModelDetails.ProductSubcategoryId = ent.GetAttributeValue<EntityReference>("hil_materialgroup").Id;
                        }
                        lstModelDetails.Add(objModelDetails);
                    }

                    ModelDetailsList.ModelDetails = lstModelDetails;
                    ModelDetailsList.Result = new ResResult { ResultStatus = true, ResultMessage = "Success" };
                    return ModelDetailsList;
                }
            }
            catch (Exception ex)
            {
                _error = ex.Message;
                ModelDetailsList.Result = new ResResult { ResultStatus = false, ResultMessage = ex.Message };
                return ModelDetailsList;
            }
            finally
            {
                _responseData = JsonConvert.SerializeObject(ModelDetailsList);
                DateTime _endDate = DateTime.Now;
                APIExceptionLog.LogAPIExecution(string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", "AMC OmniChannel", "SyncProductList", _error, _requestData.Replace(",", "~"), _responseData.Replace(",", "~"), _startDate.ToString(), _endDate.ToString(), (_endDate - _startDate).TotalSeconds.ToString() + " sec"), "OneWebsite");
            }
        }
        public AMCPlanDetailsList SyncAMCPlanDetails(string syncDateTime)
        {
            //var valueAsString = HttpContext.Current.Request.QueryString[syncDateTime];
            AMCPlanDetailsList objAMCPlanDetailsList = new AMCPlanDetailsList();
            string _requestData = JsonConvert.SerializeObject(syncDateTime);
            string _responseData = string.Empty;
            string _error = string.Empty;
            DateTime _syncDateTime;
            DateTime _startDate = DateTime.Now;
            List<AMCPlanDetails> lstAMCPlanDetails = new List<AMCPlanDetails>();
            IOrganizationService _service = ConnectToCRM.GetOrgService();

            try
            {
                if (string.IsNullOrEmpty(syncDateTime))
                {
                    objAMCPlanDetailsList.Result = new ResResult { ResultStatus = false, ResultMessage = "No Content : Sync Date time is required." };
                    return objAMCPlanDetailsList;
                }
                if (!DateTime.TryParseExact(syncDateTime, "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out _syncDateTime))
                {
                    objAMCPlanDetailsList.Result = new ResResult { ResultStatus = false, ResultMessage = "No Content : Invalid Sync Datetime format. Required format : yyyy-MM-dd HH:mm" };
                    return objAMCPlanDetailsList;
                }
                syncDateTime = _syncDateTime.ToString("yyyy-MM-dd HH:mm");
                string fetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true' {{0}}>
                              <entity name='hil_servicebom'>
                                <attribute name='hil_name' />
                                <attribute name='hil_productcategory' />
                                <attribute name='modifiedon'/>
                                <attribute name='hil_product' />
                                <attribute name='hil_servicebomid' />
	                            <filter type='and'>
			                            <condition attribute='modifiedon' operator='gt' value='{syncDateTime}' />
		                         </filter>
                                <order attribute='hil_name' descending='false' />
                                <link-entity name='product' from='productid' to='hil_productcategory' link-type='inner' alias='aq'>
                                    <attribute name='name' />
                                    <attribute name='description' />
                                  <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='ar'>
                                    <filter type='and'>
                                      <condition attribute='hil_eligibleforamc' operator='eq' value='1' />
                                    </filter>
                                  </link-entity>
                                </link-entity>
                                <link-entity name='product' from='productid' to='hil_product' link-type='inner' alias='as'>                                  
                                  <filter type='and'>
                                    <condition attribute='hil_hierarchylevel' operator='eq' value='910590001' />
                                  </filter>
                                  <link-entity name='productpricelevel' from='productid' to='productid' link-type='inner' alias='at' />
                                </link-entity>
                              </entity>
                            </fetch>";

                //EntityCollection entCollProduct = service.RetrieveMultiple(new FetchExpression(fetch));
                EntityCollection entCollProduct = RetrieveAllRecords(_service, fetch);

                if (entCollProduct.Entities.Count > 0)
                {
                    foreach (Entity entProduct in entCollProduct.Entities)
                    {
                        AMCPlanDetails objAMCPlanInfo = new AMCPlanDetails();
                        objAMCPlanInfo.ModelId = entProduct.GetAttributeValue<EntityReference>("hil_productcategory").Id;
                        objAMCPlanInfo.ModelName = entProduct.Contains("aq.description") ? entProduct.GetAttributeValue<AliasedValue>("aq.description").Value.ToString() : "";
                        objAMCPlanInfo.ModelNumber = entProduct.Contains("aq.name") ? entProduct.GetAttributeValue<AliasedValue>("aq.name").Value.ToString() : "";
                        if (entProduct.Contains("hil_product"))
                        {
                            Guid sPart = entProduct.GetAttributeValue<EntityReference>("hil_product").Id;
                            objAMCPlanInfo.PlanId = sPart;

                            QueryExpression query = new QueryExpression(ProductPriceLevel.EntityLogicalName);
                            query.ColumnSet = new ColumnSet("amount");
                            query.Criteria.AddCondition("productid", ConditionOperator.Equal, sPart);
                            query.Criteria.AddCondition("pricelevelid", ConditionOperator.Equal, PriceLevelForFGsale);// FG sale <for dev env>"8a4789cd-e78b-ed11-81ad-6045bdac5013"
                            ProductPriceLevel pAmount = _service.RetrieveMultiple(query).Entities[0].ToEntity<ProductPriceLevel>();

                            objAMCPlanInfo.MRP = decimal.Round((pAmount.Contains("amount") ? (pAmount.GetAttributeValue<Money>("amount").Value) : 0), 2);

                            string fetchPC = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'> 
                                              <entity name='hil_productcatalog'>
                                                <attribute name='hil_productcatalogid' />
                                                <attribute name='hil_name' />
                                                <attribute name='createdon' />
                                                <attribute name='hil_plantclink' />
                                                <attribute name='hil_planperiod' />
                                                <attribute name='hil_notcovered' />
                                                <attribute name='hil_coverage' />
                                                <attribute name='hil_amctandc' />
                                                <order attribute='hil_name' descending='false' /> 
                                                <filter type='and'>
                                                    <condition attribute='statecode' operator='eq' value='0' />
                                                    <condition attribute='hil_productcode' operator='eq' value='{sPart}' />
                                                </filter>
                                            </entity>
                                        </fetch>";

                            EntityCollection entdynamicproperty = _service.RetrieveMultiple(new FetchExpression(fetchPC));

                            if (entdynamicproperty.Entities.Count > 0)
                            {
                                foreach (Entity entdynamic in entdynamicproperty.Entities)
                                {
                                    objAMCPlanInfo.Coverage = entdynamic.GetAttributeValue<string>("hil_coverage");
                                    objAMCPlanInfo.NonCoverage = entdynamic.GetAttributeValue<string>("hil_notcovered");
                                    objAMCPlanInfo.PlanName = entdynamic.GetAttributeValue<string>("hil_name");
                                    objAMCPlanInfo.PlanPeriod = entdynamic.GetAttributeValue<string>("hil_planperiod");
                                    objAMCPlanInfo.PlanTCLink = entdynamic.GetAttributeValue<string>("hil_plantclink");
                                }
                            }
                        }
                        lstAMCPlanDetails.Add(objAMCPlanInfo);
                    }
                    objAMCPlanDetailsList.AMCPlanDetails = lstAMCPlanDetails;
                    objAMCPlanDetailsList.Result = new ResResult { ResultStatus = true, ResultMessage = "Success" };
                }
                return objAMCPlanDetailsList;
            }
            catch (Exception ex)
            {
                _error = ex.Message;
                objAMCPlanDetailsList.Result = new ResResult { ResultStatus = false, ResultMessage = ex.Message };
                return objAMCPlanDetailsList;
            }
            finally
            {
                _responseData = JsonConvert.SerializeObject(objAMCPlanDetailsList);
                DateTime _endDate = DateTime.Now;
                APIExceptionLog.LogAPIExecution(string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", "AMC OmniChannel", "SyncAMCPlanDetails", _error, _requestData.Replace(",", "~"), _responseData.Replace(",", "~"), _startDate.ToString(), _endDate.ToString(), (_endDate - _startDate).TotalSeconds.ToString() + " sec"), "OneWebsite");
            }
        }
        public static EntityCollection RetrieveAllRecords(IOrganizationService service, string fetchXML)
        {
            var moreRecords = false;
            int page = 1;
            var cookie = string.Empty;
            var entityCollection = new EntityCollection();
            do
            {
                var xml = string.Format(fetchXML, cookie);
                var collection = service.RetrieveMultiple(new FetchExpression(xml));

                if (collection.Entities.Count > 0) entityCollection.Entities.AddRange(collection.Entities);

                moreRecords = collection.MoreRecords;
                if (moreRecords)
                {
                    page++;
                    cookie = string.Format("paging-cookie='{0}' page='{1}'", System.Security.SecurityElement.Escape(collection.PagingCookie), page);
                }
                Console.WriteLine(entityCollection.Entities.Count);
            } while (moreRecords);

            return entityCollection;
        }

        #endregion
    }
    public class InvoiceDelailInfo
    {
        [DataMember]
        public string Mobile_Number { get; set; }
        [DataMember]
        public string Address_Line_1 { get; set; }
        [DataMember]
        public string Address_Line_2 { get; set; }
        [DataMember]
        public string Pincode { get; set; }
        [DataMember]
        public string Order_Date { get; set; }
        [DataMember]
        public string Serial_Number { get; set; }
        [DataMember]
        public string Model_Code { get; set; }
        [DataMember]
        public string AMC_Plan { get; set; }
        [DataMember]
        public string AMC_Selling_Source { get; set; }
        [DataMember]
        public decimal Payable_Amount { get; set; }
        [DataMember]
        public string Payment_Status { get; set; }
        [DataMember]
        public string Bank_Reference_Number { get; set; }
        [DataMember]
        public decimal Discount { get; set; }
        [DataMember]
        public string OrderNo { get; set; }
        [DataMember]
        public string OrderLineItem { get; set; }
        [DataMember]
        public string TransactionId { get; set; }
    }
    public class InvoiceResponse
    {
        [DataMember]
        public string InvoiceGuid { get; set; }
        [DataMember]
        public string message { get; set; }
        [DataMember]
        public bool status { get; set; }
    }
    public class AMCPlanDetails
    {
        public Guid ModelId { get; set; }
        public string ModelNumber { get; set; }
        public string ModelName { get; set; }
        public Guid PlanId { get; set; }
        public string PlanName { get; set; }
        public string PlanPeriod { get; set; }
        public decimal MRP { get; set; }
        public string Coverage { get; set; }
        public string NonCoverage { get; set; }
        public string PlanTCLink { get; set; }
    }

    public class AMCPlanDetailsList
    {
        public List<AMCPlanDetails> AMCPlanDetails { get; set; }
        public ResResult Result { get; set; }
    }
    public class ModelDetailsList
    {
        public List<ModelDetails> ModelDetails { get; set; }
        public ResResult Result { get; set; }
    }
    public class ModelDetails
    {
        public string ProductCategory { get; set; }
        public Guid ProductCategoryId { get; set; }
        public string ProductSubcategory { get; set; }
        public Guid ProductSubcategoryId { get; set; }
        public Guid ModelId { get; set; }
        public string ModelNumber { get; set; }
        public string ModelName { get; set; }
    }
    public class ResResult
    {
        public bool ResultStatus { get; set; }
        public string ResultMessage { get; set; }
    }
}
