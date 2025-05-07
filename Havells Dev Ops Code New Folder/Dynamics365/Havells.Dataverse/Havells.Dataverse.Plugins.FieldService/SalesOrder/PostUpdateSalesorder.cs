using System;
using System.IdentityModel.Metadata;
using System.Xml.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.SalesOrder
{
    public class PostUpdateSalesorder : IPlugin
    {
        private static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.PrimaryEntityName.ToLower() == "salesorder" && context.MessageName.ToUpper() == "UPDATE")
                {
                    Entity salesOrder = (Entity)context.InputParameters["Target"];
                    #region Update Discount Amount for A La Carte
                    if (salesOrder.Contains("hil_paymentstatus"))
                    {
                        salesOrder = service.Retrieve("salesorder", salesOrder.Id, new ColumnSet("hil_modeofpayment", "hil_paymentstatus", "hil_ordertype", "hil_receiptamount"));

                        EntityReference Ordertype = salesOrder.GetAttributeValue<EntityReference>("hil_ordertype");
                        int Paymentstatus = salesOrder.GetAttributeValue<OptionSetValue>("hil_paymentstatus").Value;
                        decimal receiptamount = salesOrder.Contains("hil_receiptamount") ? salesOrder.GetAttributeValue<Money>("hil_receiptamount").Value : 0;
                        if (Paymentstatus == 2)//Payment Status = Success
                        {
                            if (Ordertype.Id == new Guid("019f761c-1669-ef11-a670-000d3a3e636d"))// {OrderType = "A La Carte"}
                            {
                                if (receiptamount > 1)
                                {
                                    string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' aggregate='true' >
                                <entity name='salesorderdetail'>
                                <attribute name='salesorderid' alias='salesorderid' groupby='true'/>
                                <attribute name='baseamount' alias='totalamount' aggregate='sum' />
                                <filter type='and'>
                                    <condition attribute='salesorderid' operator='eq' value='{salesOrder.Id}' />
                                </filter>
                                </entity>
                                </fetch>";
                                    EntityCollection entColl = service.RetrieveMultiple(new FetchExpression(fetchXML));
                                    if (entColl.Entities.Count > 0)
                                    {
                                        decimal totalamount = ((Money)entColl.Entities[0].GetAttributeValue<AliasedValue>("totalamount").Value).Value;
                                        decimal DiscountPercentage = totalamount != 0 ? Math.Round(((totalamount - receiptamount) / totalamount) * 100, 2) : 0;

                                        fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' >
                                        <entity name='salesorderdetail'>
                                        <attribute name='salesorderid' />
                                        <attribute name='salesorderdetailid' />
                                        <attribute name='baseamount' />
                                        <filter type='and'>
                                            <condition attribute='salesorderid' operator='eq' value='{salesOrder.Id}' />
                                        </filter>
                                        </entity>
                                        </fetch>";
                                        EntityCollection OrderLineColl = service.RetrieveMultiple(new FetchExpression(fetchXML));
                                        if (OrderLineColl.Entities.Count > 0)
                                        {
                                            foreach (var orderline in OrderLineColl.Entities)
                                            {
                                                decimal baseamount = orderline.Contains("baseamount") ? orderline.GetAttributeValue<Money>("baseamount").Value : 0;
                                                Entity updateOrderline = new Entity(orderline.LogicalName, orderline.Id);
                                                updateOrderline["hil_eligiblediscount"] = new Money((baseamount * DiscountPercentage) / 100);
                                                updateOrderline["manualdiscountamount"] = new Money(0);
                                                service.Update(updateOrderline);
                                            }
                                        }
                                        Entity updateOrder = new Entity(salesOrder.LogicalName, salesOrder.Id);
                                        updateOrder["totallineitemamount"] = new Money(totalamount);
                                        updateOrder["discountamount"] = new Money(totalamount - receiptamount);
                                        service.Update(updateOrder);
                                    }
                                }
                            }
                            //else if (Ordertype.Id == new Guid("1f9e3353-0769-ef11-a670-0022486e4abb"))// {OrderType = "AMC Order"}
                            //{
                            //    ProcessUnitWarrantyLine(salesOrder, service, tracingService);
                            //}
                        }
                    }
                    #endregion

                    #region AutoAssign AMC Order
                    //if (salesOrder.Contains("hil_productdivision"))
                    //{
                    //    salesOrder = service.Retrieve("salesorder", salesOrder.Id, new ColumnSet("hil_modeofpayment", "hil_productdivision", "hil_serviceaddress", "hil_ordertype", "hil_sellingsource"));
                    //    EntityReference _entRefOrderType = salesOrder.GetAttributeValue<EntityReference>("hil_ordertype");
                    //    EntityReference _entRefSellingSource = salesOrder.GetAttributeValue<EntityReference>("hil_sellingsource");
                    //    OptionSetValue _opModeofPayment = salesOrder.GetAttributeValue<OptionSetValue>("hil_modeofpayment");
                    //    if (_entRefOrderType.Id == new Guid("1f9e3353-0769-ef11-a670-0022486e4abb"))
                    //    {
                    //        if (_opModeofPayment.Value != 1 || _entRefSellingSource.Id != new Guid("03b5a2d6-cc64-ed11-9562-6045bdac526a"))//AMC Order && Selling Course != FSM
                    //            ProcessCallAssignment(salesOrder, service, tracingService);
                    //    }
                    //}
                    if (salesOrder.Contains("hil_productdivision"))
                    {
                        salesOrder = service.Retrieve("salesorder", salesOrder.Id, new ColumnSet("hil_productdivision", "hil_serviceaddress", "hil_ordertype", "hil_sellingsource"));
                        EntityReference _entRefOrderType = salesOrder.GetAttributeValue<EntityReference>("hil_ordertype");
                        EntityReference _entRefSellingSource = salesOrder.GetAttributeValue<EntityReference>("hil_sellingsource");

                        if (_entRefOrderType.Id == new Guid("1f9e3353-0769-ef11-a670-0022486e4abb") && _entRefSellingSource.Id != new Guid("03b5a2d6-cc64-ed11-9562-6045bdac526a")) //AMC Order && Selling Course != FSM
                            ProcessCallAssignment(salesOrder, service, tracingService);
                    }
                    #endregion

                    #region Sync AMC Order to Loyalty Program for Point Awarding [Havells Loyalty Program]
                    if (salesOrder.Contains("statecode") && salesOrder.GetAttributeValue<OptionSetValue>("statecode")?.Value == 4)
                    {
                        string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                              <entity name='salesorder'>
                                <attribute name='name' />
                                <attribute name='hil_receiptamount' />
                                <attribute name='customerid' />
                                <attribute name='salesorderid' />
                                <attribute name='createdon' />
                                <attribute name='hil_ordertype' />
                                <attribute name='hil_source' />
                                <attribute name='statecode' />
                                <attribute name='hil_paymentstatus' />
                                <attribute name='hil_modeofpayment' />
                                <attribute name='hil_productdivision' />
                                <order attribute='createdon' descending='true' />
                                <filter type='and'>
                                  <condition attribute='hil_paymentstatus' operator='eq' value='2' />
                                    <condition attribute='statecode' operator='eq' value='4' />
                                  <condition attribute='hil_pushforloyaltyprograms' operator='ne' value='1' />
                                  <condition attribute='hil_ordertype' operator='eq' value='1f9e3353-0769-ef11-a670-0022486e4abb' />
                                  <condition attribute='salesorderid' operator='eq' value='{salesOrder.Id}' />                                 
                                </filter>
                                <link-entity name='contact' from='contactid' to='customerid' link-type='inner' alias='co'>
                                    <attribute name='mobilephone'/>     
                                    <filter type='and'>
                                        <condition attribute='hil_isloyaltyprogramenabled' operator='eq' value='1' />
                                    </filter>
                                </link-entity>
                                <link-entity name='salesorderdetail' from='salesorderid' to='salesorderid' link-type='inner' alias='az'>
                                    <link-entity name='product' from='productid' to='productid' link-type='outer' alias='pr'>
                                        <attribute name='hil_materialgroup' />
                                    </link-entity>                                    
                                    <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='hil_customerasset' link-type='inner' alias='CA'>
                                        <attribute name='hil_invoicedate' />  
                                    </link-entity>
                                </link-entity>  
                                <link-entity name='product' from='productid' to='hil_productdivision' link-type='inner' alias='ak'>
                                <attribute name='hil_division' />
                                <attribute name='hil_sapcode' />
                                <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='bc'>
                                <filter type='and'>
                                <condition attribute='hil_eligibleforloyaltyprograms' operator='eq' value='1' />
                                </filter>
                                </link-entity>
                                </link-entity>
 
                              </entity>
                            </fetch>";
                        EntityCollection entcoll = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entcoll.Entities.Count > 0)
                        {
                            foreach (var c in entcoll.Entities)
                            {
                                string MobileNumber = c.Contains("co.mobilephone") ? c.GetAttributeValue<AliasedValue>("co.mobilephone").Value.ToString() : null;
                                string Materialgroup = c.Contains("pr.hil_materialgroup") ? ((EntityReference)((AliasedValue)c["pr.hil_materialgroup"]).Value).Name : null;
                                string Division = c.Contains("ak.hil_sapcode") ? c.GetAttributeValue<AliasedValue>("ak.hil_sapcode").Value.ToString() : null;
                                string OrderNumber = c.Attributes["name"].ToString();
                                decimal receiptamount = Math.Round(c.Contains("hil_receiptamount") ? c.GetAttributeValue<Money>("hil_receiptamount").Value : 0, 2);
                                string Source = c.Contains("hil_source") ? c.GetAttributeValue<OptionSetValue>("hil_source").Value.ToString() : "";
                                string invoiceDate = c.Contains("createdon") ? ((DateTime)(c["createdon"])).ToString("dd MMM yyyy") : string.Empty;
                                string CustomerName = c.Contains("customerid") ? c.GetAttributeValue<EntityReference>("customerid").Name : null;
                                string Category = c.Contains("hil_productdivision") ? c.GetAttributeValue<EntityReference>("hil_productdivision").Name : null;

                                Entity ERCreate = new Entity("hil_easyrewardloyaltyprogram");
                                ERCreate["hil_division"] = Division;
                                ERCreate["hil_invoicevalue"] = receiptamount;
                                ERCreate["hil_customerasset"] = OrderNumber;
                                ERCreate["hil_mobilenumber"] = MobileNumber;
                                ERCreate["hil_materialgroup"] = Materialgroup;
                                ERCreate["hil_invoicedate"] = Convert.ToDateTime(invoiceDate);
                                ERCreate["hil_source"] = Source;
                                ERCreate["hil_productsyned"] = false;
                                ERCreate["hil_syncstatus"] = new OptionSetValue(1);
                                ERCreate["hil_name"] = CustomerName;
                                ERCreate["hil_synccount"] = 0;
                                ERCreate["hil_category"] = "AMC" + " " + Category;
                                var res = service.Create(ERCreate);

                                if (res != null)
                                {
                                    Entity entinvoice = new Entity("salesorder", salesOrder.Id);
                                    entinvoice["hil_pushforloyaltyprograms"] = true;
                                    service.Update(entinvoice);
                                }
                            }
                        }
                    }
                    #endregion
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("HavellsNewPlugin.SalesOrderPostUpdate.Execute Error" + ex.Message);
            }
        }
        public static void ProcessCallAssignment(Entity entity, IOrganizationService service, ITracingService tracingService)
        {
            try
            {
                if (entity.Attributes.Contains("hil_productdivision") && entity.Attributes.Contains("hil_serviceaddress"))
                {
                    Guid _saleOfficeId = Guid.Empty;
                    Guid _pinCodeId = Guid.Empty;

                    Guid _divisionId = entity.GetAttributeValue<EntityReference>("hil_productdivision").Id;
                    Entity _entAddress = service.Retrieve("hil_address", ((EntityReference)entity["hil_serviceaddress"]).Id, new ColumnSet("hil_pincode", "hil_salesoffice"));
                    if (_entAddress.Contains("hil_salesoffice"))
                        _saleOfficeId = ((EntityReference)_entAddress["hil_salesoffice"]).Id;

                    if (_entAddress.Contains("hil_pincode"))
                        _pinCodeId = ((EntityReference)_entAddress["hil_pincode"]).Id;

                    Guid _callSubtypeId = new Guid("55a71a52-3c0b-e911-a94e-000d3af06cd4"); //AMC Call

                    if (_saleOfficeId != Guid.Empty && _pinCodeId != Guid.Empty)
                    {
                        EntityReference _assignTo = null;
                        EntityReference _bshUser = null;
                        EntityReference _assignmentMatrix = null;
                        string _fetchXML = string.Empty;
                        EntityCollection _entCol = null;
                        int _matrixRowCount = 0;

                        OptionSetValue _partnerType = null;
                        OptionSetValue _brand = null;

                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='product'>
                            <attribute name='productid' />
                            <attribute name='hil_brandidentifier' />
                            <filter type='and'>
                                <condition attribute='productid' operator='eq' value='{_divisionId}' />
                            </filter>
                            </entity>
                            </fetch>";
                        _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (_entCol.Entities.Count > 0)
                            _brand = _entCol.Entities[0].GetAttributeValue<OptionSetValue>("hil_brandidentifier");

                        #region Querying SBU Branch Mapping for Fallback Assignee [BSH]
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_sbubranchmapping'>
                            <attribute name='hil_sbubranchmappingid' />
                            <attribute name='hil_branchheaduser' />
                            <filter type='and'>
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_productdivision' operator='eq' value='{_divisionId}' />
                                <condition attribute='hil_salesoffice' operator='eq' value='{_saleOfficeId}' />
                                <condition attribute='hil_branchheaduser' operator='not-null' />
                            </filter>
                            </entity>
                            </fetch>";
                        _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (_entCol.Entities.Count > 0)
                        {
                            _bshUser = _entCol.Entities[0].GetAttributeValue<EntityReference>("hil_branchheaduser");
                        }
                        else
                        {
                            QueryExpression Query = new QueryExpression("hil_integrationconfiguration");
                            Query.ColumnSet = new ColumnSet("hil_brand", "hil_approvername");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition(new ConditionExpression("hil_brand", ConditionOperator.Equal, _brand.Value));
                            Query.Criteria.AddCondition(new ConditionExpression("hil_approvername", ConditionOperator.NotNull));
                            _entCol = service.RetrieveMultiple(Query);
                            if (_entCol.Entities.Count > 0)
                            {
                                _bshUser = _entCol.Entities[0].GetAttributeValue<EntityReference>("hil_approvername");
                            }
                        }
                        #endregion

                        #region Querying Assignment Matrix to get Assignee [DSE/Franchise]
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_assignmentmatrix'>
                            <attribute name='hil_assignmentmatrixid' />
                            <attribute name='ownerid' />
                            <attribute name='hil_franchiseedirectengineer' />
                            <filter type='and'>
                                <condition attribute='hil_division' operator='eq' value='{_divisionId}' />
                                <condition attribute='hil_pincode' operator='eq' value='{_pinCodeId}' />
                                <condition attribute='hil_callsubtype' operator='eq' value='{_callSubtypeId}' />
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_franchiseedirectengineer' operator='not-null' />
                            </filter>
                            </entity>
                            </fetch>";
                        _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        _matrixRowCount = _entCol.Entities.Count;
                        if (_entCol.Entities.Count > 0)
                        {
                            _assignmentMatrix = _entCol.Entities[0].ToEntityReference();
                            EntityReference _dseFranchiseCP = _entCol.Entities[0].GetAttributeValue<EntityReference>("hil_franchiseedirectengineer");
                            Entity _entChannelPartner = service.Retrieve("account", _dseFranchiseCP.Id, new ColumnSet("ownerid", "customertypecode"));
                            _partnerType = _entChannelPartner.Contains("customertypecode") ? _entChannelPartner.GetAttributeValue<OptionSetValue>("customertypecode") : new OptionSetValue(6);
                            _assignTo = _entChannelPartner.GetAttributeValue<EntityReference>("ownerid");
                        }
                        #endregion
                        if (_matrixRowCount == 0 || _matrixRowCount > 1)
                        {
                            _assignTo = _bshUser;
                        }
                        else //Only One Assignee is found in Matrix
                        {
                            if (_partnerType.Value == 9)//DSE
                            {
                                QueryExpression Query = new QueryExpression("msdyn_timeoffrequest");
                                Query.ColumnSet = new ColumnSet(false);
                                Query.Criteria = new FilterExpression(LogicalOperator.And);
                                Query.Criteria.AddCondition(new ConditionExpression("ownerid", ConditionOperator.Equal, _assignTo.Id));
                                Query.Criteria.AddCondition(new ConditionExpression("msdyn_starttime", ConditionOperator.Today));
                                EntityCollection Found = service.RetrieveMultiple(Query);
                                if (Found.Entities.Count == 0)
                                {
                                    _assignTo = _bshUser;
                                }
                            }
                        }

                        Entity iUser = service.Retrieve("systemuser", _assignTo.Id, new ColumnSet("isdisabled", "businessunitid"));
                        bool _isDisabled = iUser.GetAttributeValue<bool>("isdisabled");
                        Entity _orderUpdate = new Entity(entity.LogicalName, entity.Id);
                        if (_isDisabled == false)
                        {
                            _orderUpdate["ownerid"] = _assignTo;
                            _orderUpdate["owningbusinessunit"] = iUser.GetAttributeValue<EntityReference>("businessunitid");
                            service.Update(_orderUpdate);
                        }
                        else
                        {
                            iUser = service.Retrieve("systemuser", _bshUser.Id, new ColumnSet("isdisabled", "businessunitid"));
                            _isDisabled = iUser.GetAttributeValue<bool>("isdisabled");
                            if (_isDisabled == false)
                            {
                                _orderUpdate["ownerid"] = _bshUser;
                                _orderUpdate["owningbusinessunit"] = iUser.GetAttributeValue<EntityReference>("businessunitid");
                                service.Update(_orderUpdate);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        public static void ProcessUnitWarrantyLine(Entity entity, IOrganizationService service, ITracingService tracingService)
        {
            try
            {
                string _fetchXMLW = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_unitwarranty'>
                        <attribute name='hil_name' />
                        <attribute name='hil_warrantyenddate' />
                        <attribute name='hil_unitwarrantyid' />
                        <order attribute='hil_warrantyenddate' descending='true' />
                        <filter type='and'>
                            <condition attribute='statecode' operator='eq' value='0' />
                            <condition attribute='hil_order' operator='eq' value='{entity.Id}' />
                        </filter>                                   
                        </entity>
                    </fetch>";
                EntityCollection entColl = service.RetrieveMultiple(new FetchExpression(_fetchXMLW));
                if (entColl.Entities.Count == 0)
                {
                    string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='salesorderdetail'>
                    <attribute name='productid' />
                    <attribute name='salesorderid' />
                    <attribute name='hil_customerasset' />
                    <filter type='and'>
                        <condition attribute='hil_customerasset' operator='not-null' />
                        <condition attribute='productid' operator='not-null' />
                        <condition attribute='salesorderid' operator='eq' value='{entity.Id}' />
                    </filter>
                    </entity>
                    </fetch>";
                    EntityCollection entOrderline = service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (entOrderline.Entities.Count > 0)
                    {
                        foreach (var orderline in entOrderline.Entities)
                        {
                            string SalesOrder = orderline.GetAttributeValue<AliasedValue>("so.name").Value.ToString();
                            Guid CustomerAssetId = orderline.GetAttributeValue<EntityReference>("hil_customerasset").Id;
                            Entity CustomerAsset = service.Retrieve("msdyn_customerasset", CustomerAssetId,
                                        new ColumnSet("hil_productcategory", "hil_productsubcategory", "hil_createwarranty", "msdyn_product", "hil_modelname", "hil_customer", "msdyn_name", "createdon"));
                            EntityReference AMCPlaneRef = orderline.GetAttributeValue<EntityReference>("productid");

                            DateTime PaymentReceiptDate = DateTime.Now;
                            fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_paymentreceipt'>
                                <attribute name='hil_paymentreceiptid' />
                                <attribute name='hil_receiptdate' />
                                <filter type='and'>
                                    <condition attribute='hil_paymentstatus' operator='eq' value='4' />
                                    <condition attribute='statecode' operator='eq' value='0' />
                                    <condition attribute='hil_orderid' operator='eq' value='{entity.Id}' />
                                </filter>
                            </entity>
                            </fetch>";

                            EntityCollection entPaymentreceipt = service.RetrieveMultiple(new FetchExpression(fetchXML));
                            if (entPaymentreceipt.Entities.Count > 0)
                            {
                                PaymentReceiptDate = entPaymentreceipt.Entities[0].GetAttributeValue<DateTime>("hil_receiptdate").AddMinutes(330);
                                Guid WarrentyTemplateId = Guid.Empty;
                                string AMCStartDate = GetWarrantyStartDate(service, CustomerAssetId, PaymentReceiptDate.Date);
                                string AMCEndDate = GetWarrantyEndDate(service, AMCPlaneRef.Id, Convert.ToDateTime(AMCStartDate), ref WarrentyTemplateId);
                                if (AMCEndDate != string.Empty)
                                {
                                    CreateUnitWarrantyLine(service, CustomerAsset, AMCPlaneRef, WarrentyTemplateId, AMCStartDate, AMCEndDate, entity.Id);
                                }
                                else {
                                    Entity entOrderUpdate = new Entity("salesorder", entity.Id);
                                    entOrderUpdate["hil_sapsyncmessage"] = $"Warranty Template does not exist AMC Plan: {AMCPlaneRef.Name}";
                                    entOrderUpdate["hil_issynctosap"] = false;
                                    service.Update(entOrderUpdate);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        private static string GetWarrantyStartDate(IOrganizationService _service, Guid AssetID, DateTime _purchaseDate)
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
                Query.TopCount = 1;
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
        private static string GetWarrantyEndDate(IOrganizationService _service, Guid _AMCPlaneID, DateTime StartDate, ref Guid WarrentyTemplateId)
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
        private static void CreateUnitWarrantyLine(IOrganizationService _service, Entity Customerasset, EntityReference Product, Guid WarrentyTemplateId, string AMCStartDate, string AMCEndDate, Guid orderId)
        {
            DateTime warrantyStartdate = Convert.ToDateTime(AMCStartDate);
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
                iSchWarranty["hil_order"] = new EntityReference("salesorder", orderId);
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

    }
}
