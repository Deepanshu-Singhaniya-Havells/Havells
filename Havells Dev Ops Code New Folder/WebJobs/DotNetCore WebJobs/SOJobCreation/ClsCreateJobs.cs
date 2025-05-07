using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace SOJobCreation
{
    public class ClsCreateJobs
    {
        private readonly ServiceClient _service;
        public ClsCreateJobs(ServiceClient service)
        {
            _service = service;
        }
        public void CreateJobs()
        {
            try
            {
                String fetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='salesorderdetail'>
                            <attribute name='productid'/>
                            <attribute name='productdescription'/>
                            <attribute name='priceperunit'/>
                            <attribute name='quantity'/>
                            <attribute name='extendedamount'/>
                            <attribute name='salesorderdetailid'/>
                            <attribute name='salesorderid'/>  
                            <attribute name='hil_eligiblediscount' />
                            <attribute name='baseamount' />
                            <order attribute='productid' descending='false'/>
                            <filter type='and'>
                                <condition attribute='hil_job' operator='null'/>
                                <condition attribute='baseamount' operator='gt' value='0' />
                            </filter>
                            <link-entity name='salesorder' from='salesorderid' to='salesorderid' link-type='inner' alias='ae'>
                                <filter type='and'>
                                    <condition attribute='hil_ordertype' operator='eq' value='{{019F761C-1669-EF11-A670-000D3A3E636D}}'/>
                                    <filter type='or'>
                                    <condition attribute='hil_paymentstatus' operator='eq' value='2'/>
                                    <condition attribute='hil_modeofpayment' operator='in'>
                                        <value>3</value>
                                        <value>2</value>
                                    </condition>
                                    </filter>
                                </filter>
                            </link-entity>
                            </entity>
                            </fetch>";

                EntityCollection linecoll = _service.RetrieveMultiple(new FetchExpression(fetch));
                Console.WriteLine($"Orderline Count No : {linecoll.Entities.Count}");
                if (linecoll.Entities.Count > 0)
                {
                    int i = 0;
                    foreach (Entity orderline in linecoll.Entities)
                    {
                        try
                        {
                            i += 1;
                            Console.WriteLine($"#{i}/{linecoll.Entities.Count} Processing Order line : {orderline.Id}");

                            Entity orderso = _service.Retrieve("salesorder", orderline.GetAttributeValue<EntityReference>("salesorderid").Id, new ColumnSet("customerid", "hil_serviceaddress", "requestdeliveryby", "hil_preferreddaytime", "hil_ordertype", "hil_modeofpayment", "hil_paymentstatus", "hil_receiptamount"));
                            Entity contact = _service.Retrieve("contact", orderso.GetAttributeValue<EntityReference>("customerid").Id, new ColumnSet("contactid", "firstname", "mobilephone", "address1_telephone2", "emailaddress1"));
                            Entity product = _service.Retrieve("product", orderline.GetAttributeValue<EntityReference>("productid").Id, new ColumnSet("hil_division", "hil_materialgroup"));

                            #region Job Creation

                            Entity Workorder = new Entity("msdyn_workorder");
                            Workorder["hil_customerref"] = new EntityReference("contact", orderso.GetAttributeValue<EntityReference>("customerid").Id);
                            Workorder["hil_customername"] = contact.Contains("firstname") ? contact.GetAttributeValue<string>("firstname") : null;
                            Workorder["hil_mobilenumber"] = contact.Contains("mobilephone") ? contact.GetAttributeValue<string>("mobilephone") : null;
                            Workorder["hil_callingnumber"] = contact.Contains("address1_telephone2") ? contact.GetAttributeValue<string>("address1_telephone2") : null;
                            Workorder["hil_alternate"] = contact.Contains("address1_telephone3") ? contact.GetAttributeValue<string>("address1_telephone3") : null;
                            Workorder["hil_email"] = contact.Contains("emailaddress1") ? contact.GetAttributeValue<string>("emailaddress1") : null;
                            Workorder["hil_preferredtime"] = orderso.Contains("hil_preferreddaytime") ? orderso.GetAttributeValue<OptionSetValue>("hil_preferreddaytime") : null;
                            Workorder["hil_preferreddate"] = orderso.Contains("requestdeliveryby") ? orderso.GetAttributeValue<DateTime>("requestdeliveryby").AddMinutes(330) : (DateTime?)null;
                            if (orderso.Contains("hil_serviceaddress"))
                            {
                                Entity _serviceAddress = _service.Retrieve("hil_address", orderso.GetAttributeValue<EntityReference>("hil_serviceaddress").Id, new ColumnSet("hil_pincode", "hil_salesoffice", "hil_branch"));
                                Workorder["hil_address"] = orderso.GetAttributeValue<EntityReference>("hil_serviceaddress");
                                if (_serviceAddress.Contains("hil_pincode") && _serviceAddress.Contains("hil_salesoffice") && _serviceAddress.Contains("hil_branch"))
                                {
                                    Workorder["hil_pincode"] = _serviceAddress.GetAttributeValue<EntityReference>("hil_pincode");
                                    Workorder["hil_salesoffice"] = _serviceAddress.GetAttributeValue<EntityReference>("hil_salesoffice");
                                    Workorder["hil_branch"] = _serviceAddress.GetAttributeValue<EntityReference>("hil_branch");
                                }
                                else
                                {
                                    Entity entSOLine = new Entity(orderline.LogicalName, orderline.Id);
                                    entSOLine["shipto_line3"] = "Customer service address Pincode/SalesOffice not found.";
                                    _service.Update(entSOLine);
                                    continue;
                                }
                            }
                            else
                            {
                                Entity entSOLine = new Entity(orderline.LogicalName, orderline.Id);
                                entSOLine["shipto_line3"] = "Customer service address not found.";
                                _service.Update(entSOLine);
                                continue;
                            }
                            Workorder["hil_productcategory"] = product.Contains("hil_division") ? new EntityReference("product", product.GetAttributeValue<EntityReference>("hil_division").Id) : null;
                            Workorder["hil_sourceofjob"] = new OptionSetValue(24);// A La Carte
                            #region Condition check for Productcategory 
                            QueryExpression query = new QueryExpression("hil_stagingdivisonmaterialgroupmapping");
                            query.ColumnSet.AddColumn("hil_name");
                            query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, product.GetAttributeValue<EntityReference>("hil_materialgroup").Name.ToString());
                            EntityCollection productcon = _service.RetrieveMultiple(query);
                            if (productcon.Entities.Count > 0)
                            {
                                Workorder["hil_productcatsubcatmapping"] = new EntityReference("hil_stagingdivisonmaterialgroupmapping", productcon[0].Id);
                            }
                            else
                            {
                                Entity entSOLine = new Entity(orderline.LogicalName, orderline.Id);
                                entSOLine["shipto_line3"] = "Material Group Configuration is mising.";
                                _service.Update(entSOLine);
                                continue;
                            }
                            #endregion

                            Workorder["hil_consumertype"] = new EntityReference("hil_consumertype", new Guid("484897de-2abd-e911-a957-000d3af0677f"));
                            Workorder["hil_consumercategory"] = new EntityReference("hil_consumercategory", new Guid("baa52f5e-2bbd-e911-a957-000d3af0677f"));
                            Workorder["hil_quantity"] = 1;
                            Workorder["hil_callertype"] = new OptionSetValue(910590001);
                            Workorder["msdyn_serviceaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43"));
                            Workorder["msdyn_billingaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43"));

                            #region condition check for natureofcomplaint & callsubtype

                            query = new QueryExpression("hil_natureofcomplaint");
                            query.TopCount = 1;
                            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            query.AddOrder("createdon", OrderType.Descending);
                            query.ColumnSet.AddColumns("hil_callsubtype", "hil_productcode", "hil_applicationservice", "hil_relatedproduct");
                            query.Criteria.AddCondition("hil_applicationservice", ConditionOperator.Equal, orderline.GetAttributeValue<EntityReference>("productid").Id); // overrided discussed with kuldder sir 

                            EntityCollection noccoll = _service.RetrieveMultiple(query);

                            if (noccoll.Entities.Count > 0)
                            {
                                Workorder["hil_natureofcomplaint"] = new EntityReference("hil_natureofcomplaint", noccoll[0].Id);
                                Workorder["hil_callsubtype"] = new EntityReference("hil_callsubtype", noccoll[0].GetAttributeValue<EntityReference>("hil_callsubtype").Id);
                            }
                            else
                            {
                                Entity entSOLine = new Entity(orderline.LogicalName, orderline.Id);
                                entSOLine["shipto_line3"] = "Nature of complaint not found related to Product on order line.";
                                _service.Update(entSOLine);
                                continue;
                            }
                            #endregion

                            int PaymentStatus = orderso.GetAttributeValue<OptionSetValue>("hil_paymentstatus").Value;
                            #region Payment Mode/Status on Job
                            if (PaymentStatus == 2)
                            {
                                Workorder["hil_modeofpayment"] = new OptionSetValue(6);//Online-PayU 
                                Workorder["hil_paymentstatus"] = new OptionSetValue(2);//Success
                            }
                            else
                            {
                                Workorder["hil_modeofpayment"] = new OptionSetValue(2);//Cash
                                Workorder["hil_paymentstatus"] = new OptionSetValue(3);//Pending
                            }
                            #endregion

                            Guid workorderid = _service.Create(Workorder);

                            #endregion

                            if (workorderid != Guid.Empty)
                            {
                                #region Update Payment Details on Job
                                Money _baseAmount = orderline.GetAttributeValue<Money>("baseamount");
                                Money _discountAmount = orderline.Contains("hil_eligiblediscount") ? orderline.GetAttributeValue<Money>("hil_eligiblediscount") : new Money(0);
                                Entity UpdateWorkorder = new Entity("msdyn_workorder", workorderid);
                                UpdateWorkorder["msdyn_subtotalamount"] = _baseAmount; //Payable Charges
                                if (PaymentStatus == 2)
                                {
                                    UpdateWorkorder["msdyn_totalsalestax"] = _discountAmount;//Discount
                                    UpdateWorkorder["msdyn_totalamount"] = new Money(_baseAmount.Value - _discountAmount.Value); // Payment Receipt
                                }
                                _service.Update(UpdateWorkorder);
                                #endregion

                                orderline["hil_job"] = new EntityReference("msdyn_workorder", workorderid);
                                orderline["shipto_line3"] = "";
                                orderline["salesorderdetailid"] = orderline.Id;
                                _service.Update(orderline);
                            }
                            Console.WriteLine($"#{i}/{linecoll.Entities.Count} Completed Order line : {orderline.Id}");
                        }
                        catch (Exception ex)
                        {
                            Entity entSOLine = new Entity(orderline.LogicalName, orderline.Id);
                            entSOLine["shipto_line3"] = ex.Message;
                            _service.Update(entSOLine);
                            Console.WriteLine($"#{i}/{linecoll.Entities.Count} Error Occured for Order line : {orderline.Id} : {ex.Message} ");
                            continue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
