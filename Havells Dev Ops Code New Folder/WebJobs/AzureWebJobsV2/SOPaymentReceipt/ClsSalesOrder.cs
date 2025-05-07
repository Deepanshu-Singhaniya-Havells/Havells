using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Workflow.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;

namespace SOPaymentReceipt
{
    public class ClsSalesOrder
    {
        private readonly IOrganizationService _service;
        public ClsSalesOrder(IOrganizationService service)
        {
            _service = service;
        }
        public void PushSalesOrderToSAP()
        {
            IntegrationConfig intConfig = IntegrationConfiguration("amc_data_v1");
           // intConfig.uri = "https://p90ci.havells.com:50001/RESTAdapter/dynamics/v1/amc_data";
            //intConfig.Auth = "D365_Havells:QAD365@1234";
            string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));
            int totalRecords = 0;
            int recordCnt = 0;
            int entOrderlineIndex = 0;
            try
            {
                string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='salesorder'>
                        <attribute name='name' />
                        <attribute name='customerid' />
                        <attribute name='totallineitemamount' />
                        <attribute name='hil_receiptamount' />
                        <attribute name='salesorderid' />
                        <attribute name='pricelevelid' />
                        <attribute name='hil_ordertype' />
                        <attribute name='hil_sourcereferencecode' />
                        <attribute name='hil_sellingsource' />
                        <attribute name='hil_branch' />
                        <attribute name='hil_serviceaddress' />
                        <attribute name='discountamount' />
                        <attribute name='ownerid' />
                        <attribute name='createdon' />
                        <order attribute='createdon' />
                        <filter type='and'>
                            <filter type='or'>
                                <condition attribute='hil_issynctosap' operator='ne' value='1' />
                                <condition attribute='hil_issynctosap' operator='null' />
                            </filter>
                            <condition attribute='customerid' operator='not-null' />
                            <condition attribute='pricelevelid' operator='not-null' />
                            <condition attribute='hil_ordertype' operator='not-null' />
                            <condition attribute='hil_branch' operator='not-null' />
                            <condition attribute='hil_serviceaddress' operator='not-null' />
                            <condition attribute='hil_paymentstatus' operator='eq' value='2' />
                            <condition attribute='totallineitemamount' operator='gt' value='0' />
                            <condition attribute='hil_receiptamount' operator='gt' value='1' />
                        </filter>
                        <link-entity name='hil_branch' from='hil_branchid' to='hil_branch' visible='false' link-type='outer' alias='br'>
                          <attribute name='hil_mamorandumcode' />
                        </link-entity>
                        <link-entity name='hil_ordertype' from='hil_ordertypeid' to='hil_ordertype' visible='false' link-type='outer' alias='ot'>
                            <attribute name='hil_refordertype' />
                        </link-entity>
                      </entity>
                    </fetch>";
                EntityCollection ec = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                Console.WriteLine("Sync Started:");
                if (ec.Entities.Count > 0)
                {
                    totalRecords = ec.Entities.Count;
                    foreach (Entity ent in ec.Entities)
                    {
                        recordCnt += 1;
                        DateTime PaymentReceiptDate = DateTime.Now;
                        fetchXML = $@"<fetch version='1.0' output-format='xml-platform' top='1' mapping='logical' distinct='false'>
                                    <entity name='hil_paymentreceipt'>
                                    <attribute name='hil_paymentreceiptid' />
                                    <attribute name='hil_transactionid' />
                                    <attribute name='hil_paymentstatus' />
                                    <attribute name='hil_bankreferenceid' />
									<attribute name='hil_receiptdate' />
									<attribute name='createdon' />
                                    <order attribute='createdon' descending='true' />
                                    <filter type='and'>
                                        <condition attribute='statecode' operator='eq' value='0' />
                                        <condition attribute='hil_paymentstatus' operator='eq' value='4' />
                                        <condition attribute='hil_orderid' operator='eq' value='{ent.Id}' />
                                    </filter>
                                    </entity>
                                </fetch>";
                        EntityCollection entPaymentreceipt = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                        if (entPaymentreceipt.Entities.Count == 0)
                        {
                            Console.WriteLine("Payment not yet received against order: {0}", ent.GetAttributeValue<string>("name"));
                            continue;
                        }
                        PaymentReceiptDate = entPaymentreceipt.Entities[0].GetAttributeValue<DateTime>("hil_receiptdate").AddMinutes(330);

                        fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
								  <entity name='salesorderdetail'>
									<attribute name='productid' />
									<attribute name='productdescription' />
									<attribute name='priceperunit' />
									<attribute name='hil_eligiblediscount' />
									<attribute name='quantity' />
									<attribute name='extendedamount' />
									<attribute name='salesorderdetailid' />
									<attribute name='hil_assetserialnumber' />
									<attribute name='hil_purchasefromlocation' />
									<attribute name='hil_purchasefrom' />
									<attribute name='hil_prefereddatetime' />
									<attribute name='hil_prefereddate' />
									<attribute name='hil_paymenttype' />
									<attribute name='salesorderid' />
									<attribute name='hil_assetmodelcode' />
									<attribute name='hil_job' />
									<attribute name='hil_invoicevalue' />
									<attribute name='hil_invoicenumber' />
									<attribute name='hil_invoicedate' />
									<attribute name='hil_discountpercent' />
									<attribute name='createdon' />
									<attribute name='hil_customerasset' />
									<order attribute='productid' descending='false' />
									<filter type='and'>
										<condition attribute='salesorderid' operator='eq' value='{ent.Id}' />
									</filter>
								  </entity>
								</fetch>";
                        EntityCollection entOrderline = _service.RetrieveMultiple(new FetchExpression(fetchXML));

                        if (entOrderline.Entities.Count > 0)
                        {
                            AMCInvoiceRequest requestData = new AMCInvoiceRequest();
                            requestData.LT_TABLE = new List<AMCInvoiceData>();
                            entOrderlineIndex = 0;
                            foreach (Entity orderline in entOrderline.Entities)
                            {
                                if (orderline.Contains("productid"))
                                {
                                    entOrderlineIndex += 1;
                                    AMCInvoiceData invoiceData = new AMCInvoiceData();
                                    EntityReference ConsumerRef = ent.GetAttributeValue<EntityReference>("customerid");
                                    Entity Consumer = _service.Retrieve(ConsumerRef.LogicalName, ConsumerRef.Id, new ColumnSet("fullname", "hil_salutation",
                                        "emailaddress1", "mobilephone"));
                                    if (Consumer.Contains("fullname"))
                                    {
                                        invoiceData.CUSTOMER_NAME = Consumer.GetAttributeValue<string>("fullname");
                                    }
                                    if (Consumer.Contains("hil_salutation"))
                                    {
                                        invoiceData.TITLE = Consumer.FormattedValues["hil_salutation"].ToString();
                                    }
                                    if (Consumer.Contains("emailaddress1"))
                                    {
                                        invoiceData.EMAIL = Consumer.GetAttributeValue<string>("emailaddress1");
                                    }
                                    if (Consumer.Contains("mobilephone"))
                                    {
                                        invoiceData.PHONE = Consumer.GetAttributeValue<string>("mobilephone");
                                    }
                                    EntityReference AddressRef = ent.GetAttributeValue<EntityReference>("hil_serviceaddress");
                                    Entity Address = _service.Retrieve(AddressRef.LogicalName, AddressRef.Id, new ColumnSet("hil_street1", "hil_street2", "hil_street3",
                                        "hil_pincode", "hil_district", "hil_city", "hil_state"));
                                    if (Address.Contains("hil_street1"))
                                    {
                                        invoiceData.STREET = Address.GetAttributeValue<string>("hil_street1");
                                    }
                                    if (Address.Contains("hil_street2"))
                                    {
                                        invoiceData.HOUSE_NO = Address.GetAttributeValue<string>("hil_street2");
                                    }
                                    if (Address.Contains("hil_street3"))
                                    {
                                        invoiceData.STREET4 = Address.GetAttributeValue<string>("hil_street3");
                                    }
                                    if (Address.Contains("hil_pincode"))
                                    {
                                        invoiceData.POSTAL_CODE = Address.GetAttributeValue<EntityReference>("hil_pincode").Name;
                                    }
                                    if (Address.Contains("hil_district"))
                                    {
                                        invoiceData.DISTRICT = Address.GetAttributeValue<EntityReference>("hil_district").Name;
                                    }
                                    if (Address.Contains("hil_city"))
                                    {
                                        invoiceData.CITY = Address.GetAttributeValue<EntityReference>("hil_city").Name;
                                    }
                                    if (Address.Contains("hil_state"))
                                    {
                                        Entity entTemp = _service.Retrieve("hil_state", Address.GetAttributeValue<EntityReference>("hil_state").Id, new ColumnSet("hil_sapstatecode"));
                                        if (entTemp.Attributes.Contains("hil_sapstatecode"))
                                        {
                                            invoiceData.REGION_CODE = entTemp.GetAttributeValue<string>("hil_sapstatecode").ToString();
                                        }
                                    }
                                    Guid WarrentyTemplateId = Guid.Empty;
                                    EntityReference entRefProduct = orderline.GetAttributeValue<EntityReference>("productid");
                                    string Orddrtype = ent.Contains("ot.hil_refordertype") ? ent.GetAttributeValue<AliasedValue>("ot.hil_refordertype").Value.ToString() : "";
                                    invoiceData.REMARKS = Orddrtype;
                                    if (Orddrtype != "A La Carte Sale")
                                    {
                                        Guid CustomerassetId = orderline.Contains("hil_customerasset") ? orderline.GetAttributeValue<EntityReference>("hil_customerasset").Id : Guid.Empty;
                                        if (CustomerassetId == Guid.Empty)
                                        {
                                            continue;
                                        }
                                        Entity CustomerAsset = _service.Retrieve("msdyn_customerasset", CustomerassetId, new ColumnSet("hil_warrantytilldate", "hil_invoicedate", "hil_productcategory", "hil_productsubcategory"));
                                        string AMCStartDate = GetWarrantyStartDate(CustomerAsset.Id, PaymentReceiptDate.Date);
                                        string AMCEndDate = GetWarrantyEndDate(entRefProduct.Id, Convert.ToDateTime(AMCStartDate), ref WarrentyTemplateId);
                                        if (AMCEndDate == string.Empty)
                                        {
                                            Console.WriteLine("AMC End Date is NUll: Sales Order# " + entRefProduct.Name + " / " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("name"));
                                            continue;
                                        }
                                        CreateUnitWarrantyLine(ConsumerRef, CustomerAsset, entRefProduct, WarrentyTemplateId, AMCStartDate, AMCEndDate);
                                        invoiceData.WARNTY_TILL_DATE = AMCEndDate;
                                        invoiceData.PRODUCT_DOP = CustomerAsset.Contains("hil_invoicedate") ? CustomerAsset.GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330).ToString("yyyy-MM-dd") : PaymentReceiptDate.ToString("yyyy-MM-dd");
                                        invoiceData.PRODUCT_SERIAL_NO = orderline.GetAttributeValue<EntityReference>("hil_customerasset").Name;
                                        Console.WriteLine("Start Date " + AMCStartDate + " || EndDate " + AMCEndDate);
                                        invoiceData.START_DATE = AMCStartDate;
                                        invoiceData.END_DATE = AMCEndDate;
                                        decimal totalamount = ent.GetAttributeValue<Money>("totallineitemamount").Value;
                                        decimal receiptamount = ent.Contains("hil_receiptamount") ? ent.GetAttributeValue<Money>("hil_receiptamount").Value : 0;
                                        decimal discountamount = Math.Round(totalamount - receiptamount, 2);
                                        if (discountamount < 0)
                                        {
                                            Console.WriteLine("Discount Amount is negative: Sales Order# " + entRefProduct.Name + " / " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("name"));
                                            continue;
                                        }
                                        invoiceData.CS_SERVICE_PRICE = totalamount.ToString();
                                        invoiceData.ZDS6_AMOUNT = discountamount.ToString();
                                        invoiceData.ITEM_NO = "1";
                                    }
                                    else
                                    {
                                        invoiceData.CS_SERVICE_PRICE = orderline.GetAttributeValue<Money>("priceperunit").Value.ToString();
                                        invoiceData.ZDS6_AMOUNT = "0";
                                        invoiceData.ITEM_NO = entOrderlineIndex.ToString();
                                    }
                                    invoiceData.DISCOUNT_FLAG = "X";
                                    string transId = entPaymentreceipt.Entities[0].GetAttributeValue<string>("hil_transactionid");
                                    invoiceData.CALL_ID = transId.Replace("D365", "");
                                    invoiceData.PAYMENT_REF_NO = entPaymentreceipt.Entities[0].Contains("hil_bankreferenceid") ? entPaymentreceipt.Entities[0].GetAttributeValue<string>("hil_bankreferenceid") : "";

                                    if (ent.Attributes.Contains("hil_branch"))
                                    {
                                        EntityReference BranchRef = ent.GetAttributeValue<EntityReference>("hil_branch");
                                        invoiceData.BRANCH_NAME = BranchRef.Name;
                                        invoiceData.SHIP_TO_PA_NAME = BranchRef.Name;
                                        string hil_mamorandumcode = "";
                                        if (ent.Attributes.Contains("br.hil_mamorandumcode"))
                                        {
                                            hil_mamorandumcode = ent.GetAttributeValue<AliasedValue>("br.hil_mamorandumcode").Value.ToString();
                                            invoiceData.SHIP_TO_PARTY = hil_mamorandumcode;
                                            invoiceData.SOLD_TO_PARTY = hil_mamorandumcode;
                                        }
                                    }
                                    invoiceData.CLOSED_ON = PaymentReceiptDate.ToString("yyyy-MM-dd");
                                    invoiceData.PRICING_DATE = ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString("yyyy-MM-dd");
                                    invoiceData.WARNTY_STATUS = "OUT";
                                    invoiceData.MATERIAL_CODE = entRefProduct.Name;

                                    string SourceReferencecode = ent.Attributes.Contains("hil_sourcereferencecode") ? ent.GetAttributeValue<string>("hil_sourcereferencecode") : "";
                                    string SellingSource = ent.Attributes.Contains("hil_sellingsource") ? ent.GetAttributeValue<EntityReference>("hil_sellingsource").Name : "";
                                    Entity entUser = _service.Retrieve("systemuser", ent.GetAttributeValue<EntityReference>("ownerid").Id, new ColumnSet("hil_employeecode"));

                                    if (entUser.Id == new Guid("5190416c-0782-e911-a959-000d3af06a98"))//CRM Admin
                                    {
                                        string[] _campaignCodeArr = SourceReferencecode.Split('|');
                                        if (_campaignCodeArr.Length >= 1)
                                            invoiceData.EMPLOYEE_ID = _campaignCodeArr[0].ToUpper();
                                        if (_campaignCodeArr.Length > 1)
                                            invoiceData.EMP_TYPE = _campaignCodeArr[1].ToUpper();
                                    }
                                    else
                                    {
                                        invoiceData.EMPLOYEE_ID = entUser.Contains("hil_employeecode") ? entUser.GetAttributeValue<string>("hil_employeecode") : SellingSource;
                                    }
                                    requestData.LT_TABLE.Add(invoiceData);
                                    var Json = JsonConvert.SerializeObject(requestData);
                                    Console.WriteLine(Json);
                                    WebRequest request = WebRequest.Create(intConfig.uri);
                                    request.Headers[HttpRequestHeader.Authorization] = authInfo;
                                    request.Method = "POST";
                                    byte[] byteArray = Encoding.UTF8.GetBytes(Json);
                                    request.ContentType = "application/x-www-form-urlencoded";
                                    request.ContentLength = byteArray.Length;
                                    Stream dataStream = request.GetRequestStream();
                                    dataStream.Write(byteArray, 0, byteArray.Length);
                                    dataStream.Close();
                                    WebResponse response = request.GetResponse();
                                    Console.WriteLine(((HttpWebResponse)response).StatusDescription);
                                    dataStream = response.GetResponseStream();
                                    StreamReader reader = new StreamReader(dataStream);
                                    string responseFromServer = reader.ReadToEnd();
                                    AMCInvoiceRequest resp = JsonConvert.DeserializeObject<AMCInvoiceRequest>(responseFromServer);
                                    Console.WriteLine(responseFromServer);
                                    Entity entUpdate = new Entity(ent.LogicalName, ent.Id);
                                    entUpdate["hil_sapsyncmessage"] = resp.LT_TABLE[0].MESSAGE;
                                    entUpdate["hil_issynctosap"] = true;

                                    if (resp.LT_TABLE[0].STATUS == "S")
                                    {
                                        entUpdate["statecode"] = new OptionSetValue(4);//Invoiced
                                        entUpdate["statuscode"] = new OptionSetValue(100003);//Invoiced
                                        entUpdate["hil_sapinvoicenumber"] = resp.LT_TABLE[0].BILLING_NUMBER;
                                        entUpdate["hil_sapsonumber"] = resp.LT_TABLE[0].ORDER_NUMBER;
                                        requestData.LT_TABLE[0].MESSAGE = resp.LT_TABLE[0].MESSAGE;
                                        requestData.LT_TABLE[0].ORDER_NUMBER = resp.LT_TABLE[0].ORDER_NUMBER;
                                        requestData.LT_TABLE[0].BILLING_NUMBER = resp.LT_TABLE[0].BILLING_NUMBER;

                                        if (ent.GetAttributeValue<EntityReference>("hil_ordertype").Id.ToString() == "1f9e3353-0769-ef11-a670-0022486e4abb")  // Order Type - AMC
                                            SaveAMCSAPInvoice(requestData.LT_TABLE[0]);
                                        CreateMediaGallary(ent, resp.LT_TABLE[0].BILLING_NUMBER);
                                    }
                                    else
                                    {
                                        Console.WriteLine("Not Done " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("name"));
                                    }
                                    _service.Update(entUpdate);
                                }
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("SOSyncToSAP.ClsSalesOrder.Program.Main.PushSalesOrderToSAP :: Error: No record found to sync");
                }
                Console.WriteLine("Sync Ended:");
            }
            catch (Exception ex)
            {
                Console.WriteLine("SOSyncToSAP.ClsSalesOrder.Program.Main.PushSalesOrderToSAP :: Error: " + ex.Message.ToString());
            }
        }

        private void CreateUnitWarrantyLine(EntityReference Customer, Entity Customerasset, EntityReference Product, Guid WarrentyTemplateId, string AMCStartDate, string AMCEndDate)
        {
            DateTime warrantyStartdate = Convert.ToDateTime(AMCStartDate);
            try
            {
                Entity entProduct = _service.Retrieve(Product.LogicalName, Product.Id, new ColumnSet("description"));
                if (entProduct != null)
                {
                    //Create Unit warranty line
                    Entity iSchWarranty = new Entity("hil_unitwarranty");
                    iSchWarranty["hil_customerasset"] = Customerasset.ToEntityReference();
                    iSchWarranty["hil_productmodel"] = Customerasset.GetAttributeValue<EntityReference>("hil_productcategory");
                    iSchWarranty["hil_productitem"] = Customerasset.GetAttributeValue<EntityReference>("hil_productsubcategory");
                    iSchWarranty["hil_warrantystartdate"] = warrantyStartdate;
                    iSchWarranty["hil_warrantyenddate"] = Convert.ToDateTime(AMCEndDate);
                    iSchWarranty["hil_warrantytemplate"] = new EntityReference("hil_warrantytemplate", WarrentyTemplateId);
                    iSchWarranty["hil_producttype"] = new OptionSetValue(1);
                    iSchWarranty["hil_part"] = entProduct.ToEntityReference();
                    iSchWarranty["hil_partdescription"] = entProduct.GetAttributeValue<string>("description");
                    iSchWarranty["hil_customer"] = Customerasset.GetAttributeValue<EntityReference>("hil_customer");
                    _service.Create(iSchWarranty);

                    //Refresh customer asset for current warranty
                    bool createwarranty = Customerasset.GetAttributeValue<bool>("hil_createwarranty");
                    Entity entCustomerAsset = new Entity("msdyn_customerasset", Customerasset.Id);
                    entCustomerAsset["hil_createwarranty"] = !createwarranty;
                    _service.Update(entCustomerAsset);

                    entCustomerAsset["hil_createwarranty"] = createwarranty;
                    _service.Update(entCustomerAsset);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        #region AMC start and end date calculation
        private string GetWarrantyStartDate(Guid AssetID, DateTime _purchaseDate)
        {
            string WarrantyStartDate = _purchaseDate.ToString("yyyy-MM-dd");
            try
            {
                LinkEntity lnkEntInvoice = new LinkEntity
                {
                    LinkFromEntityName = "hil_unitwarranty",
                    LinkToEntityName = "hil_warrantytemplate",
                    LinkFromAttributeName = "hil_warrantytemplate",
                    LinkToAttributeName = "hil_warrantytemplateid",
                    Columns = new ColumnSet("hil_type"),
                    EntityAlias = "invoice",
                    JoinOperator = JoinOperator.Inner
                };
                QueryExpression Query = new QueryExpression("hil_unitwarranty");
                Query.ColumnSet = new ColumnSet("hil_name", "hil_warrantyenddate", "hil_warrantytemplate");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression("hil_customerasset", ConditionOperator.Equal, AssetID));
                Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));  //Active
                Query.AddOrder("hil_warrantyenddate", OrderType.Descending);
                Query.LinkEntities.Add(lnkEntInvoice);
                EntityCollection ec = _service.RetrieveMultiple(Query);
                if (ec.Entities.Count >= 1)
                {
                    int WarrantyType = ((OptionSetValue)ec.Entities[0].GetAttributeValue<AliasedValue>("invoice.hil_type").Value).Value;
                    DateTime _warrantyTempDate = ec.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330).Date;
                    if (WarrantyType == 1 || WarrantyType == 3)
                    {
                        if (_warrantyTempDate >= _purchaseDate)
                        {
                            WarrantyStartDate = _warrantyTempDate.AddDays(1).ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            WarrantyStartDate = _purchaseDate.ToString("yyyy-MM-dd");
                        }
                        return WarrantyStartDate;
                    }
                    else
                    {
                        foreach (Entity entity in ec.Entities)
                        {
                            string fetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_labor'>
                                <attribute name='hil_laborid' />
                                <attribute name='hil_name' />
                                <attribute name='createdon' />
                                <order attribute='hil_name' descending='false' />
                                <filter type='and'>
                                    <condition attribute='hil_includedinwarranty' operator='eq' value='2' />
                                </filter>
                                <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplateid' link-type='inner' alias='aa'>
                                <filter type='and'>
                                    <condition attribute='hil_warrantytemplateid' operator='eq' value='{entity.GetAttributeValue<EntityReference>("hil_warrantytemplate").Id}'/>
                                </filter>
                                </link-entity>
                                </entity>
                                </fetch>";
                            EntityCollection ec1 = _service.RetrieveMultiple(new FetchExpression(fetch));
                            if (ec1.Entities.Count == 0)
                            {
                                _warrantyTempDate = entity.GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330).Date;
                                if (_warrantyTempDate >= _purchaseDate)
                                {
                                    WarrantyStartDate = _warrantyTempDate.AddDays(1).ToString("yyyy-MM-dd");
                                }
                                else
                                {
                                    WarrantyStartDate = _purchaseDate.ToString("yyyy-MM-dd");
                                }
                                return WarrantyStartDate;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return WarrantyStartDate;
        }
        private string GetWarrantyEndDate(Guid _AMCPlaneID, DateTime StartDate, ref Guid WarrentyTemplateId)
        {
            string WarrantyEndDate = string.Empty;
            QueryExpression Query = new QueryExpression("hil_warrantytemplate");
            Query.ColumnSet = new ColumnSet("hil_warrantyperiod");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("hil_amcplan", ConditionOperator.Equal, _AMCPlaneID));
            Query.Criteria.AddCondition(new ConditionExpression("hil_templatestatus", ConditionOperator.Equal, 2)); //Approved
            Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));  //Active
            Query.TopCount = 1;
            Query.AddOrder("createdon", OrderType.Descending);
            EntityCollection ec = _service.RetrieveMultiple(Query);
            if (ec.Entities.Count == 1)
            {
                WarrentyTemplateId = ec.Entities[0].Id;
                WarrantyEndDate = StartDate.AddMonths(ec[0].GetAttributeValue<int>("hil_warrantyperiod")).AddDays(-1).ToString("yyyy-MM-dd");
            }
            return WarrantyEndDate;
        }
        #endregion
        #region Save SAP AMC Invoices
        private ActionResponse SaveAMCSAPInvoice(AMCInvoiceData objamcsapdetails)
        {
            ActionResponse objresult = new ActionResponse();
            try
            {
                QueryExpression query = new QueryExpression("hil_amcstaging");
                query.ColumnSet = new ColumnSet("hil_serailnumber");
                query.NoLock = true;
                query.Criteria = new FilterExpression(LogicalOperator.And);
                if (!string.IsNullOrWhiteSpace(objamcsapdetails.PRODUCT_SERIAL_NO))
                    query.Criteria.AddCondition("hil_serailnumber", ConditionOperator.Equal, objamcsapdetails.PRODUCT_SERIAL_NO);
                query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, objamcsapdetails.BILLING_NUMBER);
                EntityCollection objEntity = _service.RetrieveMultiple(query);
                if (objEntity.Entities.Count <= 0)
                {
                    QueryExpression qsCType = new QueryExpression("product");
                    qsCType.ColumnSet = new ColumnSet("name");
                    qsCType.NoLock = true;
                    qsCType.Criteria.AddCondition("name", ConditionOperator.Equal, objamcsapdetails.MATERIAL_CODE);
                    EntityCollection AmcplanGuid = _service.RetrieveMultiple(qsCType);

                    Entity entAMCUpdate = new Entity("hil_amcstaging");
                    entAMCUpdate["hil_callid"] = objamcsapdetails.CALL_ID;
                    entAMCUpdate["hil_serailnumber"] = objamcsapdetails.PRODUCT_SERIAL_NO;
                    entAMCUpdate["hil_mobilenumber"] = objamcsapdetails.PHONE;
                    entAMCUpdate["hil_warrantystartdate"] = Convert.ToDateTime(objamcsapdetails.START_DATE).AddMinutes(330);
                    entAMCUpdate["hil_warrantyenddate"] = Convert.ToDateTime(objamcsapdetails.END_DATE).AddMinutes(330);
                    entAMCUpdate["hil_amcplannameslt"] = objamcsapdetails.MATERIAL_CODE;
                    entAMCUpdate["hil_name"] = objamcsapdetails.BILLING_NUMBER;
                    entAMCUpdate["hil_sapbillingdate"] = DateTime.Now;
                    entAMCUpdate["hil_salesordernumber"] = objamcsapdetails.ORDER_NUMBER;
                    entAMCUpdate["hil_branchmemorandumcode"] = objamcsapdetails.SOLD_TO_PARTY;
                    if (AmcplanGuid.Entities.Count > 0)
                    {
                        if (AmcplanGuid.Entities[0].Id != Guid.Empty)
                            entAMCUpdate["hil_amcplan"] = new EntityReference("product", AmcplanGuid.Entities[0].Id);

                    }
                    string EncryptBIllingNumber = new ClsEncryptDecrypt().EncryptAES256URL(objamcsapdetails.BILLING_NUMBER);
                    entAMCUpdate["hil_sapbillingdocpath"] = ConfigurationManager.AppSettings["Docurl"].ToString() + EncryptBIllingNumber;
                    _service.Create(entAMCUpdate);
                    objresult.Is_Successful = true;
                    objresult.Message = "Success";
                    return objresult;
                }
                else
                {
                    objresult.Is_Successful = false;
                    objresult.Message = "Failed";
                    return objresult;
                }
            }
            catch (Exception ex)
            {
                objresult.Is_Successful = false;
                objresult.Message = "Failed";
                return objresult;
            }
        }
        #endregion
        private ActionResponse CreateMediaGallary(Entity SalesOrder, string BILLING_NUMBER)
        {
            ActionResponse objresult = new ActionResponse();
            try
            {
                Entity mediagallery = new Entity("hil_mediagallery");
                mediagallery["hil_salesorder"] = new EntityReference("salesorder", SalesOrder.Id);
                mediagallery["hil_name"] = SalesOrder.GetAttributeValue<string>("name") + "_Invoice";
                mediagallery["hil_consumer"] = SalesOrder.GetAttributeValue<EntityReference>("customerid");
                mediagallery["hil_mediatype"] = new EntityReference("hil_mediatype", new Guid("e5589095-4301-ef11-9f89-6045bdac6fcc"));//Sales Invoice
                string EncryptBIllingNumber = new ClsEncryptDecrypt().EncryptAES256URL(BILLING_NUMBER);
                mediagallery["hil_url"] = ConfigurationManager.AppSettings["Docurl"].ToString() + EncryptBIllingNumber;
                _service.Create(mediagallery);
                objresult.Is_Successful = true;
                return objresult;
            }
            catch (Exception ex)
            {
                objresult.Is_Successful = false;
                objresult.Message = ex.Message;
                return objresult;
            }
        }
        #region Supporting Methods
        private IntegrationConfig IntegrationConfiguration(string Param)
        {
            IntegrationConfig output = new IntegrationConfig();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = _service.RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + ":" + integrationConfiguration.GetAttributeValue<string>("hil_password");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            return output;
        }
        #endregion
    }
}
