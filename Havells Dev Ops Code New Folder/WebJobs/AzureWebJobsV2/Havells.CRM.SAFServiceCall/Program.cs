using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.ModelBinding;
using System.Web.Services.Description;
using System.Xml.Schema;

namespace Havells.CRM.SAFServiceCall
{
    public class Program
    {
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
            var CrmURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();
            string finalString = string.Format(connStr, CrmURL);
            IOrganizationService service = createConnection(finalString);
            getAllActiveSFAServiceRequest(service);
        }

        static void getJobs()
        {
            for (int i = 0; i < 10000; i++)
            {
                Thread.Sleep(1000);
                Console.WriteLine("Service Request Retrived");
                Thread.Sleep(1000);
                Console.WriteLine("Address Created");
                Thread.Sleep(1000);
                Console.WriteLine("Customer Created");
                Thread.Sleep(1000);
                Console.WriteLine("Asset Created");
                Thread.Sleep(1000);
                Console.WriteLine("Job Created Created");
                Thread.Sleep(1000);
                Console.WriteLine("Job Created Created");
                Thread.Sleep(2000);
                Console.WriteLine("****************Done*********************");
            }
        }
        static void getAllActiveSFAServiceRequest(IOrganizationService service)
        {
            QueryExpression query = new QueryExpression("hil_servicecallrequest");
            query.ColumnSet = new ColumnSet(true);
            query.Criteria = new FilterExpression(LogicalOperator.And);
            //query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 1);

            query.Criteria.AddCondition("hil_syncstatus", ConditionOperator.NotEqual, true);
            query.Criteria.AddCondition("createdon", ConditionOperator.Today);//, new DateTime(2022, 12, 13));

            //query.Criteria.AddCondition("hil_servicecallrequestid", ConditionOperator.Equal, new Guid("b1f1cbb3-ba97-ed11-aad1-6045bdac5a1d"));

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
            OptionSetValue _ConsumerSource = new OptionSetValue(6);
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
                    Entity entity123 = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_image"));
                    string photoName = entity["hil_name"] + ".jpg" as string;
                    string entityimage_url = entity["hil_image_url"] as string;
                    String file = String.Format("{0}", photoName);

                    byte[] image = entity123["hil_image"] as byte[];
                    var fs = Convert.ToBase64String(image, 0, image.Length);


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
                        _CustomerID = new EntityReference("contact", createCustomer(fullname, mobileNo, EmailID, _ConsumerSource, service));
                    }
                    if (entity.Contains("hil_address"))
                        _AddressID = entity.GetAttributeValue<EntityReference>("hil_address");
                    else
                    {

                        if (entity.Contains("hil_areacode"))
                            areaCode = (string)entity["hil_areacode"];
                        if (entity.Contains("hil_pincode"))
                            pincode = (string)entity["hil_pincode"];
                        if (entity.Contains("hil_fulladdress"))
                            fulladdress = (string)entity["hil_fulladdress"];


                        _AddressID = new EntityReference("hil_address", createCustomerAddress(_CustomerID, fulladdress, areaCode, pincode, service));
                    }
                    //getServiceCallDetails(service, entity, _CustomerID, _AddressID, new OptionSetValue(8), fs);
                    getServiceCallDetails(service, entity, _CustomerID, _AddressID, new OptionSetValue(9), fs);
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
        public static void AttachNotes(IOrganizationService service, EntityReference customerAsset, string attachment)
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
        static void getServiceCallDetails(IOrganizationService service, Entity ServiceCallRequestId, EntityReference customer, EntityReference address, OptionSetValue Source, string attachment)
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
                        InvoiceValue = ServiceCallRequestId.GetAttributeValue<Money>("hil_invoicevalue").Value;
                        PurchasedFrom = ServiceCallRequestId.GetAttributeValue<string>("hil_purchasefrom");
                        PurchasedFromLocation = ServiceCallRequestId.GetAttributeValue<string>("hil_purchaselocation");
                        product = entity.GetAttributeValue<EntityReference>("hil_product");
                        modelName = entity.GetAttributeValue<string>("hil_name");
                        customerAsset = new EntityReference("msdyn_customerasset", CreateCustomerAsset(customer, serialNo, Source, dtInvoice, InvoiceNo, InvoiceValue, PurchasedFrom, PurchasedFromLocation, product, modelName, service));
                        AttachNotes(service, customerAsset, attachment);
                        if (entity.Contains("hil_callsubtype"))
                            callSubType = entity.GetAttributeValue<EntityReference>("hil_callsubtype").Name;
                        if (callSubType != null && callSubType != "Product Registration")
                        {
                            if (callSubType != "Both")
                            {
                                createServiceCall(service, customer, address, customerAsset, callSubType);
                            }
                            if (callSubType == "Both")
                            {
                                createServiceCall(service, customer, address, customerAsset, "Installation");
                                createServiceCall(service, customer, address, customerAsset, "Demo");
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
        static void createServiceCall(IOrganizationService service, EntityReference customer, EntityReference address, EntityReference erCustomerAsset, string callSubType)
        {
            try
            {
                #region Create Service Call
                Int32 warrantyStatus = 2;
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
        static Guid CreateCustomerAsset(EntityReference customer, string serialNo, OptionSetValue Source, DateTime dtInvoice, string InvoiceNumber,
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
        static Guid createCustomerAddress(EntityReference consumer, string Line1, string areaCode, string PinCode, IOrganizationService service)
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
        static Guid createCustomer(string fullName, string mobileNo, string EmailID, OptionSetValue source, IOrganizationService service)
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
        public static IOrganizationService createConnection(string connectionString)
        {
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);
                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    Console.WriteLine(((CrmServiceClient)service).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }
            return service;
        }
    }
}
