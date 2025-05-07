using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System.Text;

namespace EPOSCreateServiceCall
{
    internal class ClsCreateJobs
    {
        private readonly ServiceClient service;
        public ClsCreateJobs(ServiceClient _service)
        {
            service = _service;
        }

        public void getAllActiveSFAServiceRequest()
        {
            QueryExpression query = new QueryExpression("hil_servicecallrequest");
            query.ColumnSet = new ColumnSet(true);
            query.Criteria = new FilterExpression(LogicalOperator.And);
            query.Criteria.AddCondition("hil_syncstatus", ConditionOperator.NotEqual, true);
            query.Criteria.AddCondition("createdon", ConditionOperator.Today);//, new DateTime(2022, 12, 13));
            query.Criteria.AddCondition("hil_source", ConditionOperator.Equal, 23);
            query.Criteria.AddCondition("hil_remarks", ConditionOperator.Null);
            EntityCollection ActiveSFAServiceRequestColl = service.RetrieveMultiple(query);
            Console.WriteLine("Total Record Found " + ActiveSFAServiceRequestColl.Entities.Count);
            EntityReference _CustomerID = null;
            EntityReference _AddressID = null;
            String fullname = null;
            string mobileNo = null;
            string EmailID = null;
            string areaCode = null;
            string fulladdress = null;
            string pincode = null;
            var fs = "";
            OptionSetValue source = new OptionSetValue(6);
            foreach (Entity entity in ActiveSFAServiceRequestColl.Entities)
            {
                try
                {
                    fullname = null;
                    fulladdress = null;
                    pincode = null;
                    areaCode = null;
                    _AddressID = null;
                    mobileNo = null;
                    _CustomerID = null;
                    bool consent = true;
                    bool gender = true;
                    EmailID = null;
                    fs = "";
                    DateTime dateofbirth = new DateTime(1900, 1, 1);
                    int num = entity.Contains("hil_source") ? entity.GetAttributeValue<OptionSetValue>("hil_source").Value : 0;
                    if (num == 23)
                    {
                        source = new OptionSetValue(num);
                    }
                    else
                    {
                        Entity entity123 = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_image"));
                        string photoName = entity["hil_name"] + ".jpg" as string;
                        string entityimage_url = entity["hil_image_url"] as string;
                        String file = String.Format("{0}", photoName);

                        byte[] image = entity123["hil_image"] as byte[];
                        fs = Convert.ToBase64String(image, 0, image.Length);
                    }

                    EmailID = null;
                    if (entity.Contains("hil_consumer"))
                        _CustomerID = entity.GetAttributeValue<EntityReference>("hil_consumer");
                    else
                    {

                        if (entity.Contains("hil_customername"))
                            fullname = (string)entity["hil_customername"];
                        if (entity.Contains("hil_customermobilenumber"))
                            mobileNo = (string)entity["hil_customermobilenumber"];
                        if (entity.Contains("hil_customeremailid"))
                            EmailID = (string)entity["hil_customeremailid"];
                        if (entity.Contains("hil_dateofbirth"))
                            dateofbirth = (DateTime)entity["hil_dateofbirth"];
                        if (entity.Contains("hil_consent"))
                            consent = (bool)entity["hil_consent"];
                        if (entity.Contains("hil_gender"))
                            gender = (bool)entity["hil_gender"];
                        _CustomerID = new EntityReference("contact", createCustomer(fullname, mobileNo, EmailID, source, service));
                    }
                    EntityReference address;
                    if (entity.Contains("hil_address"))
                    {
                        address = entity.GetAttributeValue<EntityReference>("hil_address");
                    }
                    else
                    {
                        if (entity.Contains("hil_pincode"))
                            pincode = (string)entity["hil_pincode"];

                        if (entity.Contains("hil_areacode"))
                        {
                            areaCode = (string)entity["hil_areacode"];
                        }
                        else
                        {
                            QueryExpression query2 = new QueryExpression("hil_pincode");
                            query2.ColumnSet = new ColumnSet("hil_pincodeid");
                            query2.Criteria.AddCondition("hil_name", ConditionOperator.Equal, pincode);
                            EntityCollection entity3 = service.RetrieveMultiple(query2);
                            if (entity3.Entities.Count == 0)
                                throw new Exception("Pin Code Not Found");

                            QueryExpression query3 = new QueryExpression("hil_businessmapping");
                            query3.ColumnSet = new ColumnSet("hil_area", "hil_stagingarea");
                            query3.Criteria = new FilterExpression(LogicalOperator.And);
                            query3.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, entity3[0].Id);
                            query3.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 1);
                            EntityCollection entityCollection2 = service.RetrieveMultiple((QueryBase)query3);
                            if (entityCollection2.Entities.Count == 0)
                                throw new Exception("Area Code Not Found");
                            areaCode = Convert.ToString(entityCollection2[0].GetAttributeValue<string>("hil_stagingarea"));
                        }
                        if (entity.Contains("hil_fulladdress"))
                            fulladdress = (string)entity["hil_fulladdress"];
                        address = new EntityReference("hil_address", createCustomerAddress(_CustomerID, fulladdress, areaCode, pincode, service));
                    }
                    //getServiceCallDetails(service, entity, _CustomerID, _AddressID, new OptionSetValue(8), fs);
                    getServiceCallDetails(entity, _CustomerID, _AddressID, source, fs);
                    Entity entity1 = new Entity(entity.LogicalName, entity.Id);
                    entity1["hil_syncstatus"] = true;
                    entity1["hil_remarks"] = "Done";
                    service.Update(entity1);
                }
                catch (Exception ex)
                {
                    Entity entity1 = new Entity(entity.LogicalName, entity.Id);
                    entity1["hil_remarks"] = ex.Message;
                    service.Update(entity1);
                }
            }
        }
        private void AttachNotes(EntityReference customerAsset, string attachment)
        {
            try
            {
                string title = string.Empty;
                string fileName = string.Empty;
                string mimeType = "image/jpeg";
                string extension = ".jpg";

                fileName = customerAsset.Id.ToString() + "InvoiceImage" + extension;

                Entity An = new Entity("annotation");
                An["documentbody"] = attachment;
                An["mimetype"] = mimeType;
                An["filename"] = fileName;
                An["objectid"] = customerAsset;
                An["objectidtypecode"] = customerAsset.LogicalName;
                service.Create(An);

            }
            catch (Exception ex)
            {
                throw new Exception("Error in Attachment Creation " + ex.Message);
            }
        }
        private void getServiceCallDetails(Entity ServiceCallRequestId, EntityReference customer, EntityReference address, OptionSetValue Source, string attachment)
        {
            try
            {
                QueryExpression query = new QueryExpression("hil_servicecallrequestdetail");
                query.ColumnSet = new ColumnSet(true);
                query.Criteria = new FilterExpression(LogicalOperator.And);
                //query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 1);
                query.Criteria.AddCondition("hil_syncstatus", ConditionOperator.NotEqual, true);
                query.Criteria.AddCondition("hil_servicecallrequest", ConditionOperator.Equal, ServiceCallRequestId.Id);
                query.Criteria.AddCondition("hil_syncremarks", ConditionOperator.Null);
                EntityCollection entityColl = service.RetrieveMultiple(query);
                EntityReference product = null;
                string areaCode = null;
                string fulladdress = null;
                string pincode = null;
                string serialNo = null;
                string InvoiceNo = null;
                string PurchasedFrom = null;
                string modelName = null;
                string PurchasedFromLocation = null;
                decimal InvoiceValue = 0;
                EntityReference customerAsset = null;
                string callSubType = null;
                DateTime dtInvoice = new DateTime(1900, 01, 01);
                DateTime PreferredDateofService;
                OptionSetValue PreferredTimeofService;
                foreach (Entity entity in entityColl.Entities)
                {
                    try
                    {


                        callSubType = null;
                        modelName = null;
                        product = null;
                        InvoiceNo = null;
                        areaCode = null;
                        fulladdress = null;
                        pincode = null;
                        serialNo = null;
                        PurchasedFrom = null;
                        PurchasedFromLocation = null;
                        dtInvoice = new DateTime(1900, 01, 01);
                        InvoiceValue = 0;
                        PreferredDateofService = new DateTime(1900, 1, 1);
                        PreferredTimeofService = new OptionSetValue(1);
                        if (entity.Contains("hil_address"))
                            address = entity.GetAttributeValue<EntityReference>("hil_address");
                        else if (entity.Contains("hil_areacode") && entity.Contains("hil_pincode") && entity.Contains("hil_fulladdress"))
                        {
                            areaCode = (string)entity["hil_areacode"];
                            pincode = (string)entity["hil_pincode"];
                            fulladdress = (string)entity["hil_fulladdress"];
                            address = new EntityReference("hil_address", createCustomerAddress(customer, fulladdress, areaCode, pincode, service));
                        }
                        serialNo = entity.GetAttributeValue<string>("hil_serialnumber");
                        dtInvoice = ServiceCallRequestId.GetAttributeValue<DateTime>("hil_invoicedate");
                        InvoiceNo = ServiceCallRequestId.GetAttributeValue<string>("hil_invoicenumber");
                        InvoiceValue = ServiceCallRequestId.Contains("hil_invoicevalue") ? ServiceCallRequestId.GetAttributeValue<Money>("hil_invoicevalue").Value : 0;
                        PurchasedFrom = ServiceCallRequestId.GetAttributeValue<string>("hil_purchasefrom");
                        PurchasedFromLocation = ServiceCallRequestId.GetAttributeValue<string>("hil_purchaselocation");
                        product = entity.GetAttributeValue<EntityReference>("hil_product");
                        PreferredDateofService = entity.GetAttributeValue<DateTime>("hil_servicedate").AddMinutes(330);
                        PreferredTimeofService = entity.GetAttributeValue<OptionSetValue>("hil_preferredtime");
                        modelName = entity.GetAttributeValue<string>("hil_name");


                        if (entity.Contains("hil_callsubtype"))
                            callSubType = entity.GetAttributeValue<EntityReference>("hil_callsubtype").Name;
                        if (Source.Value == 23)
                        {
                            createServiceCallePos(service, customer, address, callSubType, PreferredDateofService, PreferredTimeofService, modelName);
                        }
                        else
                        {
                            customerAsset = new EntityReference("msdyn_customerasset", CreateCustomerAsset(customer, serialNo, Source, dtInvoice, InvoiceNo, InvoiceValue, PurchasedFrom, PurchasedFromLocation, product, modelName, service));
                            AttachNotes(customerAsset, attachment);
                            if (callSubType != null && callSubType != "Product Registration")
                            {
                                if (callSubType != "Both")
                                {
                                    createServiceCall(customer, address, customerAsset, callSubType);
                                }
                                if (callSubType == "Both")
                                {
                                    createServiceCall(customer, address, customerAsset, "Installation");
                                    createServiceCall(customer, address, customerAsset, "Demo");
                                }
                            }
                        }
                        Entity entity1 = new Entity(entity.LogicalName, entity.Id);
                        entity1["hil_syncstatus"] = true;
                        entity1["hil_syncremarks"] = "Done";
                        service.Update(entity1);
                    }
                    catch (Exception ex)
                    {
                        Entity entity1 = new Entity(entity.LogicalName, entity.Id);
                        entity1["hil_syncremarks"] = ex.Message;
                        service.Update(entity1);
                        throw new Exception(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        private void createServiceCall(EntityReference customer, EntityReference address, EntityReference erCustomerAsset, string callSubType)
        {
            try
            {
                #region Create Service Call
                //  Int32 warrantyStatus = 2;
                Entity contatcEntity = service.Retrieve(customer.LogicalName, customer.Id, new ColumnSet("mobilephone", "emailaddress1", "fullname"));
                Entity enWorkorder = new Entity("msdyn_workorder");
                enWorkorder["hil_customerref"] = customer;
                if (contatcEntity.Contains("fullname"))
                    enWorkorder["hil_customername"] = contatcEntity.GetAttributeValue<string>("fullname");
                if (contatcEntity.Contains("mobilephone"))
                    enWorkorder["hil_mobilenumber"] = contatcEntity.GetAttributeValue<string>("mobilephone");
                if (contatcEntity.Contains("emailaddress1"))
                    enWorkorder["hil_email"] = contatcEntity.GetAttributeValue<string>("emailaddress1");
                enWorkorder["hil_address"] = address;

                enWorkorder["msdyn_customerasset"] = erCustomerAsset;

                Entity _AssetEntity = service.Retrieve(erCustomerAsset.LogicalName, erCustomerAsset.Id, new ColumnSet("hil_modelname", "hil_productcategory",
                    "hil_productsubcategory", "hil_productsubcategorymapping"));

                if (_AssetEntity.Contains("hil_modelname"))
                {
                    enWorkorder["hil_modelname"] = _AssetEntity.GetAttributeValue<string>("hil_modelname");
                }
                if (_AssetEntity.Contains("hil_productcategory"))
                {
                    enWorkorder["hil_productcategory"] = _AssetEntity.GetAttributeValue<EntityReference>("hil_productcategory");
                }
                if (_AssetEntity.Contains("hil_productsubcategory"))
                {
                    enWorkorder["hil_productsubcategory"] = _AssetEntity.GetAttributeValue<EntityReference>("hil_productsubcategory");
                }
                if (_AssetEntity.Contains("hil_productsubcategorymapping"))
                {
                    enWorkorder["hil_productcatsubcatmapping"] = _AssetEntity.GetAttributeValue<EntityReference>("hil_productsubcategorymapping");
                }

                EntityCollection entCol;
                //enWorkorder["hil_warrantystatus"] = new OptionSetValue(warrantyStatus);
                QueryExpression Query = new QueryExpression("hil_consumertype");
                Query.ColumnSet = new ColumnSet("hil_consumertypeid");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "B2C");
                entCol = service.RetrieveMultiple(Query);
                if (entCol.Entities.Count > 0)
                {
                    enWorkorder["hil_consumertype"] = entCol.Entities[0].ToEntityReference();
                }

                Query = new QueryExpression("hil_consumercategory");
                Query.ColumnSet = new ColumnSet("hil_consumercategoryid");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "End User");
                entCol = service.RetrieveMultiple(Query);
                if (entCol.Entities.Count > 0)
                {
                    enWorkorder["hil_consumercategory"] = entCol.Entities[0].ToEntityReference();
                }
                Query = new QueryExpression("hil_natureofcomplaint");
                Query.ColumnSet = new ColumnSet("hil_callsubtype");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, callSubType);
                Query.Criteria.AddCondition("hil_relatedproduct", ConditionOperator.Equal, _AssetEntity.GetAttributeValue<EntityReference>("hil_productsubcategory").Id);
                entCol = service.RetrieveMultiple(Query);

                if (entCol.Entities.Count > 0)
                {
                    enWorkorder["hil_natureofcomplaint"] = entCol[0].ToEntityReference();
                    enWorkorder["hil_callsubtype"] = entCol[0].GetAttributeValue<EntityReference>("hil_callsubtype");
                }
                else
                {
                    throw new Exception("NoC and Call Sub Type not Define");
                }

                enWorkorder["hil_quantity"] = 1;
                enWorkorder["hil_callertype"] = new OptionSetValue(910590001);// {SourceofJob:"Customer"}

                enWorkorder["hil_sourceofjob"] = new OptionSetValue(7); // SourceofJob:[{"7": "SFA"}]

                enWorkorder["msdyn_serviceaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {ServiceAccount:"Dummy Account"}
                enWorkorder["msdyn_billingaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {BillingAccount:"Dummy Account"}
                                                                                                                                        //enWorkorder["msdyn_primaryincidenttype"] = new EntityReference("msdyn_incidenttype", new Guid("0F5E8009-3BFD-E811-A94C-000D3AF0694E")); // {Primary Incident Type:"Installation -Decorative FAN CF"}

                Guid serviceCallGuid = service.Create(enWorkorder);

                #endregion
            }
            catch (Exception ex)
            {
                throw new Exception("Error in Job Creation " + ex.Message);
            }
        }
        private Guid CreateCustomerAsset(EntityReference customer, string serialNo, OptionSetValue Source, DateTime dtInvoice, string InvoiceNumber,
            decimal InvoiceValue, string PurchasedFrom, string PurchasedFromLocation, EntityReference product, string _ModelName, IOrganizationService service)
        {
            Guid customerAssetId = Guid.Empty;
            try
            {
                Entity entCustomerAsset = new Entity("msdyn_customerasset");
                entCustomerAsset["hil_customer"] = customer;
                entCustomerAsset["msdyn_name"] = serialNo;
                entCustomerAsset["hil_source"] = Source;
                entCustomerAsset["hil_invoiceavailable"] = true;
                entCustomerAsset["hil_retailerpincode"] = "0";
                entCustomerAsset["hil_purchasedfrom"] = "Registered From SFA";
                entCustomerAsset["hil_invoicedate"] = dtInvoice;
                entCustomerAsset["hil_invoiceno"] = InvoiceNumber;
                entCustomerAsset["hil_invoicevalue"] = InvoiceValue;
                entCustomerAsset["hil_purchasedfrom"] = PurchasedFrom;
                entCustomerAsset["hil_retailerpincode"] = PurchasedFromLocation;
                entCustomerAsset["msdyn_product"] = product;
                entCustomerAsset["hil_modelname"] = _ModelName;

                Entity _ProductEntity = service.Retrieve(product.LogicalName, product.Id, new ColumnSet("hil_division", "hil_materialgroup"));
                entCustomerAsset["hil_productcategory"] = (EntityReference)_ProductEntity["hil_division"];
                entCustomerAsset["hil_productsubcategory"] = (EntityReference)_ProductEntity["hil_materialgroup"];
                QueryExpression Query = new QueryExpression("hil_stagingdivisonmaterialgroupmapping");
                Query.ColumnSet = new ColumnSet("hil_productcategorydivision", "hil_name", "hil_productsubcategorymg", "statecode");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_productcategorydivision", ConditionOperator.Equal, ((EntityReference)_ProductEntity["hil_division"]).Id);
                Query.Criteria.AddCondition("hil_productsubcategorymg", ConditionOperator.Equal, ((EntityReference)_ProductEntity["hil_materialgroup"]).Id);
                EntityCollection Found = service.RetrieveMultiple(Query);
                if (Found.Entities.Count > 0)
                {
                    entCustomerAsset["hil_productsubcategorymapping"] = Found.Entities[0].ToEntityReference();
                }

                entCustomerAsset["statuscode"] = new OptionSetValue(910590000); // Pending for Approval
                customerAssetId = service.Create(entCustomerAsset);
                Console.WriteLine("Customer Asset is Created.");
            }
            catch (Exception ex)
            {

                throw new Exception("Error in Customer Asset Creation || " + ex.Message);
            }
            return customerAssetId;
        }
        private Guid createCustomerAddress(EntityReference consumer, string Line1, string areaCode, string PinCode, IOrganizationService service)
        {

            Guid _AddressID = Guid.Empty;
            EntityReference _pincode = null;
            EntityReference _area = null;
            EntityReference _BussMapp = null;
            EntityReference _district = null;
            try
            {

                QueryExpression query = new QueryExpression("hil_pincode");
                query.ColumnSet = new ColumnSet(false);
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, PinCode);
                query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 1);
                EntityCollection pincodeColl = service.RetrieveMultiple(query);
                if (pincodeColl.Entities.Count == 0)
                {
                    throw new Exception("PinCode Not Found");
                }
                else
                {
                    _pincode = pincodeColl[0].ToEntityReference();
                }
                query = new QueryExpression("hil_area");
                query.ColumnSet = new ColumnSet(false);
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("hil_areacode", ConditionOperator.Equal, areaCode);
                query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 1);
                EntityCollection branchColl = service.RetrieveMultiple(query);
                if (branchColl.Entities.Count == 0)
                {
                    throw new Exception("Area Not Found");
                }
                else
                {
                    _area = branchColl[0].ToEntityReference();
                }
                query = new QueryExpression("hil_businessmapping");
                query.ColumnSet = new ColumnSet("hil_district");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, _pincode.Id);
                query.Criteria.AddCondition("hil_area", ConditionOperator.Equal, _area.Id);
                query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 1);
                EntityCollection BussMappColl = service.RetrieveMultiple(query);
                if (BussMappColl.Entities.Count == 0)
                {
                    throw new Exception("Bussiness Mapping Not Found");
                }
                else
                {
                    _BussMapp = BussMappColl[0].ToEntityReference();
                    _district = BussMappColl[0].GetAttributeValue<EntityReference>("hil_district");
                }

                Entity entity = new Entity("hil_address");
                entity["hil_customer"] = consumer;
                entity["hil_addresstype"] = new OptionSetValue(1);
                entity["hil_street1"] = Line1;
                entity["hil_businessgeo"] = _BussMapp;
                entity["hil_district"] = _district;
                _AddressID = service.Create(entity);
                Console.WriteLine("Customer Address is Created.");
            }
            catch (Exception ex)
            {
                throw new Exception("Error in Address Createtion || " + ex.Message);
            }

            return _AddressID;
        }
        private Guid createCustomer(string fullName, string mobileNo, string EmailID, OptionSetValue source, IOrganizationService service)
        {

            Guid _CustomerID = Guid.Empty;
            try
            {
                QueryExpression Query = new QueryExpression("contact");
                Query.ColumnSet = new ColumnSet(false);
                Query.Criteria = new FilterExpression(LogicalOperator.Or);
                if (EmailID != null)
                    Query.Criteria.AddCondition("emailaddress1", ConditionOperator.Equal, EmailID);
                Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, mobileNo);
                EntityCollection Found = service.RetrieveMultiple(Query);
                if (Found.Entities.Count > 0)
                {
                    _CustomerID = Found[0].Id;
                }
                else
                {
                    Entity entity = new Entity("contact");
                    string[] consumerName = fullName.Split(' ');
                    if (consumerName.Length >= 1)
                    {
                        entity["firstname"] = consumerName[0];
                        if (consumerName.Length == 3)
                        {
                            entity["middlename"] = consumerName[1];
                            entity["lastname"] = consumerName[2];
                        }
                        if (consumerName.Length == 2)
                        {
                            entity["lastname"] = consumerName[1];
                        }
                    }
                    else
                    {
                        entity["firstname"] = fullName;
                    }
                    entity["fullname"] = fullName;

                    entity["mobilephone"] = mobileNo;
                    if (EmailID != null)
                        entity["emailaddress1"] = EmailID;
                    entity["hil_consumersource"] = new OptionSetValue(6); //{6,SFA MobileApp}
                    _CustomerID = service.Create(entity);
                }
                Console.WriteLine("Customer is Created.");
            }
            catch (Exception ex)
            {
                throw new Exception("Error in Consumer Createtion || " + ex.Message);
            }

            return _CustomerID;
        }
        private void createServiceCallePos(IOrganizationService service, EntityReference customer, EntityReference address, string callSubType, DateTime PreferredDateofService, OptionSetValue PreferredTimeofService, string modelName)
        {
            try
            {
                Guid ProductCategory = Guid.Empty;
                Guid ProductSubcategory = Guid.Empty;
                EntityCollection entityColProduct = null;
                Entity enCustomer = service.Retrieve(customer.LogicalName, customer.Id, new ColumnSet("mobilephone", "emailaddress1", "fullname"));
                Entity enWorkorder = new Entity("msdyn_workorder");
                enWorkorder["hil_customerref"] = customer;
                if (enCustomer.Contains("fullname"))
                    enWorkorder["hil_customername"] = enCustomer.GetAttributeValue<string>("fullname");
                if (enCustomer.Contains("mobilephone"))
                    enWorkorder["hil_mobilenumber"] = enCustomer.GetAttributeValue<string>("mobilephone");
                if (enCustomer.Contains("emailaddress1"))
                    enWorkorder["hil_email"] = enCustomer.GetAttributeValue<string>("emailaddress1");
                enWorkorder["hil_address"] = address;
                QueryExpression query1 = new QueryExpression("hil_consumertype");
                query1.ColumnSet = new ColumnSet("hil_consumertypeid");
                query1.Criteria = new FilterExpression(LogicalOperator.And);
                query1.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "B2C");
                EntityCollection entityCollection1 = service.RetrieveMultiple((QueryBase)query1);
                if (entityCollection1.Entities.Count > 0)
                    enWorkorder["hil_consumertype"] = entityCollection1.Entities[0].ToEntityReference();

                QueryExpression query2 = new QueryExpression("hil_consumercategory");
                query2.ColumnSet = new ColumnSet("hil_consumercategoryid");
                query2.Criteria = new FilterExpression(LogicalOperator.And);
                query2.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "End User");
                EntityCollection entityCollection2 = service.RetrieveMultiple((QueryBase)query2);
                if (entityCollection2.Entities.Count > 0)
                    enWorkorder["hil_consumercategory"] = entityCollection2.Entities[0].ToEntityReference();

                if (!string.IsNullOrEmpty(modelName))
                {
                    string fetchXmlQuery = $@"<fetch top='1' version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='product'>
                        <attribute name='name' />
                        <attribute name='hil_materialgroup' />
                        <attribute name='hil_division' />
                        <order attribute='productnumber' descending='false' />
                        <filter type='and'>
                          <condition attribute='name' operator='eq' value='{modelName}' />
                        </filter>
                      </entity>
                    </fetch>";
                    entityColProduct = service.RetrieveMultiple(new FetchExpression(fetchXmlQuery));
                    if (entityColProduct.Entities.Count > 0)
                    {
                        ProductCategory = entityColProduct.Entities[0].Contains("hil_division") ? entityColProduct.Entities[0].GetAttributeValue<EntityReference>("hil_division").Id : Guid.Empty;
                        ProductSubcategory = entityColProduct.Entities[0].Contains("hil_materialgroup") ? entityColProduct.Entities[0].GetAttributeValue<EntityReference>("hil_materialgroup").Id : Guid.Empty;
                    }
                }
                enWorkorder["hil_productcategory"] = new EntityReference("product", ProductCategory);

                string fetchXmlQuery2 = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='hil_stagingdivisonmaterialgroupmapping'>
                                    <attribute name='hil_stagingdivisonmaterialgroupmappingid' />
                                    <attribute name='hil_name' />
                                    <attribute name='createdon' />
                                    <order attribute='hil_name' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='hil_productsubcategorymg' operator='eq'  uitype='product' value='{ProductSubcategory}' />
                                      <condition attribute='statecode' operator='eq' value='0' />
                                      <condition attribute='hil_productcategorydivision' operator='eq' uitype='product' value='{ProductCategory}' />
                                    </filter>
                                  </entity>
                                </fetch>";
                entityColProduct = service.RetrieveMultiple(new FetchExpression(fetchXmlQuery2));
                if (entityColProduct.Entities.Count > 0)
                {
                    enWorkorder["hil_productcatsubcatmapping"] = entityColProduct.Entities[0].ToEntityReference();
                }

                string fetchXmlQuery3 = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='hil_natureofcomplaint'>
                                    <attribute name='hil_name'/>
                                    <attribute name='hil_callsubtype'/>
                                    <attribute name='hil_relatedproduct'/>
                                    <attribute name='hil_natureofcomplaintid'/>
                                    <order attribute='hil_name' descending='false'/>
                                    <filter type='and'>
                                    <condition attribute='hil_name' operator='eq' value='{callSubType}'/>
                                    <condition attribute='hil_relatedproduct' operator='eq' uitype='product' value='{ProductSubcategory}'/>
                                    </filter>
                                    </entity>
                                    </fetch>";
                entityColProduct = service.RetrieveMultiple(new FetchExpression(fetchXmlQuery3));
                if (entityColProduct.Entities.Count > 0)
                {
                    enWorkorder["hil_natureofcomplaint"] = entityColProduct[0].ToEntityReference();
                    enWorkorder["hil_callsubtype"] = entityColProduct[0].GetAttributeValue<EntityReference>("hil_callsubtype");
                }
                else
                {
                    throw new Exception("NoC and Call Sub Type not Define");
                }
                enWorkorder["hil_quantity"] = 1;
                enWorkorder["hil_callertype"] = new OptionSetValue(910590001);
                enWorkorder["hil_sourceofjob"] = new OptionSetValue(23);
                enWorkorder["hil_preferreddate"] = PreferredDateofService;
                enWorkorder["hil_preferredtime"] = PreferredTimeofService;
                enWorkorder["msdyn_serviceaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43"));
                enWorkorder["msdyn_billingaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43"));
                service.Create(enWorkorder);

                ParamResponse Response = new ParamResponse();
                CustomerDetails customerDetails = new CustomerDetails();
                customerDetails.Customer = new List<Customer>();
                using (HttpClient Client = new HttpClient())
                {
                    QueryExpression query4 = new QueryExpression("hil_integrationconfiguration");
                    query4.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                    query4.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "ePosTokenGeneration");
                    Entity entity3 = service.RetrieveMultiple(query4)[0];
                    string URL = entity3.GetAttributeValue<string>("hil_url");
                    Client.DefaultRequestHeaders.Accept.Clear();
                    Client.DefaultRequestHeaders.Add("SERVICE_METHODNAME", "GetToken");
                    Client.DefaultRequestHeaders.Add("Username", entity3.GetAttributeValue<string>("hil_username"));
                    Client.DefaultRequestHeaders.Add("Password", entity3.GetAttributeValue<string>("hil_password"));
                    HttpResponseMessage paramResponse = Client.PostAsync(URL, null).Result;
                    if (paramResponse.IsSuccessStatusCode)
                    {
                        Response = JsonConvert.DeserializeObject<ParamResponse>(paramResponse.Content.ReadAsStringAsync().Result);
                        if (Response.Response.Result == "SUCCESS")
                        {
                            Console.WriteLine("Access Token Generated");
                        }
                    }
                    customerDetails.Customer.Add(new Customer()
                    {
                        MobileNumber = enCustomer.GetAttributeValue<string>("mobilephone"),
                        AlternateCustomerCode = Convert.ToString(customer.Id)
                    });

                    HttpClient Client1 = new HttpClient();
                    var Paramdata = new StringContent(JsonConvert.SerializeObject(customerDetails), Encoding.UTF8, "application/json");
                    QueryExpression query5 = new QueryExpression("hil_integrationconfiguration");
                    query5.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                    query5.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "ePosProcessData");
                    string URL1 = service.RetrieveMultiple((QueryBase)query5)[0].GetAttributeValue<string>("hil_url");
                    Client1.DefaultRequestHeaders.Add("SERVICE_METHODNAME", "SetCustomerAlternateCode");
                    Client1.DefaultRequestHeaders.Add("Authorization", Response.Response.Access_Token);
                    HttpResponseMessage responsetier = Client1.PostAsync(URL1, Paramdata).Result;
                    if (responsetier.IsSuccessStatusCode)
                    {
                        var resulttier = responsetier.Content.ReadAsStringAsync().Result;
                        dynamic Response1 = JsonConvert.DeserializeObject<dynamic>(resulttier);
                        string ePosResopnse = Response1["Response"]["Result"];
                        if (ePosResopnse == "SUCCESS")
                        {
                            Console.WriteLine("SUCCESS");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error in Job Creation " + ex.Message);
            }
        }
    }

    public class ParamResult
    {
        public string Result { get; set; }

        public string Access_Token { get; set; }
    }

    public class ParamResponse
    {
        public ParamResult Response { get; set; }
    }

    public class Customer
    {
        public string MobileNumber { get; set; }

        public string AlternateCustomerCode { get; set; }
    }

    public class CustomerDetails
    {
        public List<Customer> Customer { get; set; }
    }

    public class Response
    {
        public string Result { get; set; }

    }
}
