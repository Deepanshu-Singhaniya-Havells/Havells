using Microsoft.Office.Interop.Excel;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Globalization;
using System.Net;
using System.Web.Services.Description;
using System.Xml.Linq;

namespace CustomerAssetWarrantyDailyRefresh
{
    public class RefreshUnitWarranty
    {
        private readonly ServiceClient _service;
        public RefreshUnitWarranty(ServiceClient service)
        {
            _service = service;
        }
        public void RefreshAssetUnitWarranty()
        {
            QueryExpression query = new QueryExpression("hil_amcstaging");
            query.ColumnSet = new ColumnSet("hil_name", "hil_warrantystartdate", "hil_warrantyenddate", "hil_serailnumber", "hil_sapbillingdocpath", "hil_sapbillingdate", "hil_amcstagingstatus", "hil_amcplan");
            query.Criteria.AddCondition("hil_amcstagingstatus", ConditionOperator.Equal, false);
            // query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "6001268725");
            query.Criteria.AddCondition("createdon", ConditionOperator.Today);
            EntityCollection entCollAMCStaging = RetrieveAllRecords(query);
            if (entCollAMCStaging.Entities.Count > 0)
            {
                try
                {
                    int i = 1;
                    int Total = entCollAMCStaging.Entities.Count;
                    Guid amcPlanGuid = Guid.Empty;
                    string amcPlanDesc = string.Empty;
                    foreach (Entity entAMC in entCollAMCStaging.Entities)
                    {
                        Entity entAMCStaging = new Entity("hil_amcstaging", entAMC.Id);
                        string serialNumber = entAMC.GetAttributeValue<string>("hil_serailnumber");
                        string name = entAMC.GetAttributeValue<string>("hil_name");
                        DateTime warrantyStartdate = entAMC.GetAttributeValue<DateTime>("hil_warrantystartdate").AddMinutes(330);
                        DateTime warrantyEnddate = entAMC.GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330);
                        DateTime sapbillingdate = entAMC.GetAttributeValue<DateTime>("hil_sapbillingdate").AddMinutes(330);
                        string amcPlanName = entAMC.GetAttributeValue<EntityReference>("hil_amcplan").Name;
                        Console.WriteLine("Processing... " + i.ToString() + "/" + Total.ToString() + " Serial # " + serialNumber);

                        string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='msdyn_customerasset'>
                                    <attribute name='hil_productcategory' />
                                    <attribute name='hil_productsubcategory' />
                                    <attribute name='hil_createwarranty' />
                                    <attribute name='msdyn_product' />
                                    <attribute name='hil_modelname' />
                                    <attribute name='hil_customer' />
                                    <attribute name='msdyn_name' />
                                    <attribute name='createdon' />
                                    <order attribute='createdon' descending='false' />
                                    <filter type='and'>
                                        <condition attribute='msdyn_name' operator='eq' value='{serialNumber}' />
					                </filter>
                                    </entity>
                                    </fetch>";
                        EntityCollection CustomerAsset = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                        if (CustomerAsset.Entities.Count > 0)
                        {
                            fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                              <entity name='product'>
                                                <attribute name='name' />
                                                <attribute name='productnumber' />
                                                <attribute name='description' />
                                                <attribute name='statecode' />
                                                <attribute name='productstructure' />
                                                <attribute name='productid' />
                                                <order attribute='productnumber' descending='false' />
                                                <filter type='and'>
                                                  <condition attribute='hil_hierarchylevel' operator='eq' value='910590001' />
                                                  <condition attribute='name' operator='eq' value='{amcPlanName}' />
                                                </filter>
                                              </entity>
                                            </fetch>";

                            EntityCollection entCollProduct = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                            if (entCollProduct.Entities.Count > 0)
                            {
                                amcPlanGuid = entCollProduct.Entities[0].Id;
                                amcPlanDesc = entCollProduct.Entities[0].GetAttributeValue<string>("description");
                            }
                            string _fetchXMLW = $@"<fetch top='1' version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='hil_unitwarranty'>
                                    <attribute name='hil_name' />
                                    <attribute name='hil_warrantyenddate' />
                                    <attribute name='hil_unitwarrantyid' />
                                    <order attribute='hil_warrantyenddate' descending='true' />
                                    <filter type='and'>
                                        <condition attribute='statecode' operator='eq' value='0' />
                                        <condition attribute='hil_customerasset' operator='eq' value='{CustomerAsset.Entities[0].Id}' />
                                        <condition attribute='hil_amcbillingdocnum' operator='eq' value='{name}' />
                                    </filter>                                   
                                  </entity>
                                </fetch>";

                            EntityCollection entColl = _service.RetrieveMultiple(new FetchExpression(_fetchXMLW));
                            if (entColl.Entities.Count > 0)
                            {
                                Console.WriteLine("Unit Warranty already exist for " + i.ToString() + "/" + Total.ToString() + " Serial # " + serialNumber);
                                entAMCStaging["hil_amcstagingstatus"] = true;
                                entAMCStaging["hil_description"] = "Unit Warranty already exist";
                                _service.Update(entAMCStaging);
                                continue;
                            }
                            else
                            {
                                string _currentDate = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString().PadLeft(2, '0') + "-" + DateTime.Now.Day.ToString().PadLeft(2, '0');
                                _fetchXMLW = $@"<fetch top='1' version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='hil_unitwarranty'>
                                    <attribute name='hil_name' />
                                    <attribute name='hil_warrantyenddate' />
                                    <attribute name='hil_unitwarrantyid' />
                                    <order attribute='hil_warrantyenddate' descending='true' />
                                    <filter type='and'>
                                        <condition attribute='statecode' operator='eq' value='0' />
                                        <condition attribute='hil_customerasset' operator='eq' value='{CustomerAsset.Entities[0].Id}' />
                                        <condition attribute='hil_warrantystartdate' operator='on-or-before' value='{_currentDate}' />
                                        <condition attribute='hil_warrantyenddate' operator='on-or-after' value='{_currentDate}' />
                                    </filter>
                                    <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='wt'>
                                      <attribute name='hil_warrantyperiod' />
                                      <attribute name='hil_type' />
                                      <filter type='and'>
                                        <condition attribute='hil_type' operator='in'>
                                            <value>3</value>
                                            <value>1</value>
                                          </condition>
                                      </filter>
                                    </link-entity>
                                  </entity>
                                </fetch>";

                                EntityCollection entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXMLW));
                                if (entCol.Entities.Count > 0)
                                {
                                    if (entCol.Entities[0].Attributes.Contains("hil_warrantyenddate"))
                                    {
                                        warrantyStartdate = entCol.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330).AddDays(1);
                                    }
                                }

                                Entity warrantyTemplate = null;
                                int period = 0;
                                QueryExpression queryExp = new QueryExpression("hil_warrantytemplate");
                                queryExp.ColumnSet = new ColumnSet("hil_warrantytemplateid", "hil_warrantyperiod");
                                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                queryExp.Criteria.AddCondition(new ConditionExpression("hil_amcplan", ConditionOperator.Equal, amcPlanGuid));
                                queryExp.Criteria.AddCondition(new ConditionExpression("hil_templatestatus", ConditionOperator.Equal, 2)); //Approved
                                queryExp.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1)); //Active
                                EntityCollection entCollTemp = _service.RetrieveMultiple(queryExp);
                                if (entCollTemp.Entities.Count > 0)
                                {
                                    warrantyTemplate = entCollTemp.Entities[0];
                                    period = entCollTemp.Entities[0].GetAttributeValue<int>("hil_warrantyperiod");
                                }
                                if (warrantyTemplate != null)
                                {
                                    //Create Unit warranty line
                                    Entity iSchWarranty = new Entity("hil_unitwarranty");
                                    iSchWarranty["hil_customerasset"] = CustomerAsset.Entities[0].ToEntityReference();
                                    iSchWarranty["hil_productmodel"] = CustomerAsset.Entities[0].GetAttributeValue<EntityReference>("hil_productcategory");
                                    iSchWarranty["hil_productitem"] = CustomerAsset.Entities[0].GetAttributeValue<EntityReference>("hil_productsubcategory");
                                    iSchWarranty["hil_warrantystartdate"] = warrantyStartdate;
                                    iSchWarranty["hil_warrantyenddate"] = warrantyEnddate;
                                    iSchWarranty["hil_warrantytemplate"] = warrantyTemplate.ToEntityReference();
                                    iSchWarranty["hil_producttype"] = new OptionSetValue(1);
                                    iSchWarranty["hil_part"] = entCollProduct.Entities[0].ToEntityReference();
                                    iSchWarranty["hil_partdescription"] = amcPlanDesc;
                                    iSchWarranty["hil_customer"] = CustomerAsset.Entities[0].GetAttributeValue<EntityReference>("hil_customer");
                                    iSchWarranty["hil_amcbillingdocnum"] = name;
                                    iSchWarranty["hil_amcbillingdocdate"] = sapbillingdate;
                                    _service.Create(iSchWarranty);
                                    Console.WriteLine("Unit Warranty line created for " + i.ToString() + "/" + Total.ToString() + " Serial # " + serialNumber);

                                    //Update AMC Staging after process
                                    entAMCStaging["hil_amcstagingstatus"] = true;
                                    entAMCStaging["hil_description"] = "Done";

                                    //Refresh customer asset for current warranty
                                    bool createwarranty = CustomerAsset.Entities[0].GetAttributeValue<bool>("hil_createwarranty");
                                    Entity entCustomerAsset = new Entity("msdyn_customerasset", CustomerAsset.Entities[0].Id);
                                    entCustomerAsset["hil_createwarranty"] = !createwarranty;
                                    _service.Update(entCustomerAsset);

                                    entCustomerAsset["hil_createwarranty"] = createwarranty;
                                    _service.Update(entCustomerAsset);
                                }
                                else
                                {
                                    entAMCStaging["hil_amcstagingstatus"] = false;
                                    entAMCStaging["hil_description"] = "Warranty Template not found";
                                }
                            }
                        }
                        else
                        {
                            entAMCStaging["hil_amcstagingstatus"] = false;
                            entAMCStaging["hil_description"] = "Serial No not found";
                        }
                        //Update AMC Staging  
                        _service.Update(entAMCStaging);
                        i++;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error : " + ex.Message);
                }
            }
        }
        public void RefreshAssetAMCUnitWarranty()
        {
            try
            {
                string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='salesorder'>
                        <attribute name='name' />
                        <attribute name='customerid' />
                        <attribute name='totalamount' />
                        <attribute name='salesorderid' />
                        <attribute name='pricelevelid' />
                        <attribute name='hil_receiptamount' />
                        <attribute name='hil_ordertype' />
                        <attribute name='hil_source' />
                        <attribute name='hil_source' />
                        <attribute name='hil_serviceaddress' />
                        <attribute name='discountamount' />
                        <attribute name='createdon' />
                        <order attribute='createdon' descending='true' />
                        <filter type='and'>
                        <condition attribute='createdon' operator='on-or-after' value='2025-01-22' />
                        <condition attribute='customerid' operator='not-null' />
                            <condition attribute='pricelevelid' operator='not-null' />
                            <condition attribute='hil_serviceaddress' operator='not-null' />
                            <condition attribute='hil_paymentstatus' operator='eq' value='2' />
                            <condition attribute='totalamount' operator='gt' value='0' />
                            <condition attribute='hil_ordertype' operator='eq' value='{{1F9E3353-0769-EF11-A670-0022486E4ABB}}' />
                        </filter>
                    </entity>
                    </fetch>";

                EntityCollection ec = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                if (ec.Entities.Count > 0)
                {
                    int OrderCount = 0;

                    foreach (Entity ent in ec.Entities)
                    {
                        OrderCount++;
                        string SalesOrderNumber = ent.GetAttributeValue<string>("name");
                        fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='salesorderdetail'>
                                <attribute name='productid' />
                                <attribute name='productdescription' />
                                <attribute name='salesorderdetailid' />
                                <attribute name='hil_assetserialnumber' />
                                <attribute name='hil_paymenttype' />
                                <attribute name='salesorderid' />
                                <attribute name='hil_assetmodelcode' />
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
                            foreach (Entity orderline in entOrderline.Entities)
                            {
                                try
                                {
                                    Entity CustomerAsset = _service.Retrieve("msdyn_customerasset", orderline.GetAttributeValue<EntityReference>("hil_customerasset").Id,
                                        new ColumnSet("hil_productcategory", "hil_productsubcategory", "hil_createwarranty", "msdyn_product", "hil_modelname", "hil_customer", "msdyn_name", "createdon"));

                                    EntityReference ConsumerRef = ent.GetAttributeValue<EntityReference>("customerid");
                                    Entity Consumer = _service.Retrieve(ConsumerRef.LogicalName, ConsumerRef.Id, new ColumnSet("fullname", "hil_salutation","emailaddress1", "mobilephone"));

                                    EntityReference RefAMCPlane = orderline.GetAttributeValue<EntityReference>("productid");

                                    #region Acknowledgement, Warrenty Start & End Date 

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
                                                      <condition attribute='hil_orderid' operator='eq' value='{ent.Id}' />
                                                    </filter>
                                                  </entity>
                                                </fetch>";

                                    EntityCollection entPaymentreceipt = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                                    if (entPaymentreceipt.Entities.Count > 0)
                                    {
                                        int paymentstatus = entPaymentreceipt.Entities[0].GetAttributeValue<OptionSetValue>("hil_paymentstatus").Value;
                                        if (paymentstatus != 4)//success
                                        {
                                            Console.WriteLine($"Payment not yet received for Order No {0}", SalesOrderNumber);
                                            continue;
                                        }
                                    }
                                    Guid WarrentyTemplateId = Guid.Empty;
                                    string AMCStartDate = GetWarrantyStartDate(orderline.GetAttributeValue<EntityReference>("hil_customerasset").Id, entPaymentreceipt.Entities[0].GetAttributeValue<DateTime>("hil_receiptdate").AddMinutes(330));
                                    string AMCEndDate = GetWarrantyEndDate(orderline.GetAttributeValue<EntityReference>("productid").Id, Convert.ToDateTime(AMCStartDate), ref WarrentyTemplateId);

                                    if (AMCEndDate == string.Empty)
                                    {
                                        Console.WriteLine("AMC End Date is NUll: Sales Order# " + SalesOrderNumber);
                                        continue;
                                    }
                                    CreateUnitWarrantyLine(ConsumerRef, CustomerAsset, RefAMCPlane, WarrentyTemplateId, AMCStartDate, AMCEndDate);
                                    Console.WriteLine("Sales Order#{0}: {1} ", SalesOrderNumber, OrderCount.ToString() + "/" + ec.Entities.Count.ToString());
                                    #endregion
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("Error For Sales Order#{0}: {1}", SalesOrderNumber, ex.Message);
                                    continue;
                                }
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : " + ex.Message);
            }
        }
        private string GetWarrantyStartDate(Guid AssetID, DateTime _purchaseDate)
        {
            DateTime WarrantyStartDate = _purchaseDate;
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
                    DateTime _warrantyTempDate = ec.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330);
                    if (WarrantyType == 1 || WarrantyType == 3)
                    {
                        if (_warrantyTempDate >= _purchaseDate)
                        {
                            WarrantyStartDate = _warrantyTempDate.AddDays(1);
                        }
                        return WarrantyStartDate.ToString("dd-MM-yyyy");
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
                                _warrantyTempDate = entity.GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330);
                                if (_warrantyTempDate >= _purchaseDate)
                                {
                                    WarrantyStartDate = _warrantyTempDate.AddDays(1);
                                }
                                return WarrantyStartDate.ToString("dd-MM-yyyy");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return WarrantyStartDate.ToString("dd-MM-yyyy");
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
            else
            {
                Console.WriteLine("Warranty Template not found");
            }
            return WarrantyEndDate;
        }
        private void CreateUnitWarrantyLine(EntityReference Customer, Entity Customerasset, EntityReference Product, Guid WarrentyTemplateId, string AMCStartDate, string AMCEndDate)
        {
            DateTime warrantyStartdate = Convert.ToDateTime(AMCStartDate);
            try
            {
                string _currentDate = DateTime.Now.ToString("yyyy-MM-dd");
                string _fetchXMLW = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='hil_unitwarranty'>
                                    <attribute name='hil_name' />
                                    <attribute name='hil_warrantyenddate' />
                                    <attribute name='hil_unitwarrantyid' />
                                    <attribute name='createdon' />
                                    <order attribute='createdon' descending='false' />
                                    <filter type='and'>
                                        <condition attribute='statecode' operator='eq' value='0' />
                                        <condition attribute='hil_customerasset' operator='eq' value='{Customerasset.Id}' />
                                        <condition attribute='hil_warrantytemplate' operator='eq' value='{WarrentyTemplateId}' />
                                        <condition attribute='hil_part' operator='eq' value='{Product.Id}' />
                                        <condition attribute='createdon' operator='on-or-after' value='2025-01-24' />
                                    </filter>
                                  </entity>
                                </fetch>";

                EntityCollection entColl = _service.RetrieveMultiple(new FetchExpression(_fetchXMLW));
                if (entColl.Entities.Count == 1)
                {
                    return;
                }
                _fetchXMLW = $@"<fetch top='1' version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='hil_unitwarranty'>
                                    <attribute name='hil_name' />
                                    <attribute name='hil_warrantyenddate' />
                                    <attribute name='hil_unitwarrantyid' />
                                    <attribute name='createdon' />
                                    <order attribute='hil_warrantyenddate' descending='true' />
                                    <filter type='and'>
                                        <condition attribute='statecode' operator='eq' value='0' />
                                        <condition attribute='hil_customerasset' operator='eq' value='{Customerasset.Id}' />
                                        <condition attribute='hil_warrantystartdate' operator='on-or-before' value='{_currentDate}' />
                                        <condition attribute='hil_warrantyenddate' operator='on-or-after' value='{_currentDate}' />
                                    </filter>
                                    <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='wt'>
                                      <attribute name='hil_warrantyperiod' />
                                      <attribute name='hil_type' />
                                      <filter type='and'>
                                        <condition attribute='hil_type' operator='in'>
                                            <value>3</value>
                                            <value>1</value>
                                          </condition>
                                      </filter>
                                    </link-entity>
                                  </entity>
                                </fetch>";

                EntityCollection entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXMLW));
                if (entCol.Entities.Count > 0)
                {
                    if (entCol.Entities[0].Attributes.Contains("hil_warrantyenddate"))
                    {
                        warrantyStartdate = entCol.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330).AddDays(1);
                    }
                }
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

                    //Update AMC Staging after process

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
        private EntityCollection RetrieveAllRecords(QueryExpression query, int count = 5000)
        {
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = count;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;
            EntityCollection entityCollection = _service.RetrieveMultiple(query);
            EntityCollection ecFinal = new EntityCollection();
            foreach (Entity i in entityCollection.Entities)
            {
                ecFinal.Entities.Add(i);
            }
            do
            {
                query.PageInfo.PageNumber += 1;
                query.PageInfo.PagingCookie = entityCollection.PagingCookie;
                entityCollection = _service.RetrieveMultiple(query);
                foreach (Entity i in entityCollection.Entities)
                    ecFinal.Entities.Add(i);
            }
            while (entityCollection.MoreRecords);
            return ecFinal;
        }
    }
}
