using Havells_Plugin.HelperIntegration;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class AMCOmniChannelSale
    {
        //#region Consumer Verification
        //public ConsumerVerification ConsumerVerification(ConsumerInfo consumerInfo)
        //{
        //    ConsumerVerification objConsumerVerification = new ConsumerVerification();
            

        //    string fullname = string.Empty;
        //    string emailaddress1 = string.Empty;
        //    string hil_state = string.Empty;
        //    string hil_pincode = string.Empty;
        //    string hil_fulladdress = string.Empty;
        //    string hil_city = string.Empty;
        //    string hil_district = string.Empty;

        //    try
        //    {
        //        IOrganizationService service = ConnectToCRM.GetOrgService();
        //        ConsumerInfo objConsumerInfo = new ConsumerInfo();
        //        objConsumerInfo.AddressInfo = new AddressInfo();
        //        List<AssetWarrantyInfo> lstAssetWaarrantyInfo = new List<AssetWarrantyInfo>();
        //        Guid customerGuid = Guid.Empty;
        //        if (string.IsNullOrWhiteSpace(consumerInfo.MobileNumber) || consumerInfo.MobileNumber.Length < 10)
        //        {
        //            objConsumerVerification.Result = new Result { ResultStatus = "200", ResultMessage = "Please provide valid Mobile Number" };
        //            return objConsumerVerification;
        //        }
        //        else
        //        {
        //            QueryExpression Query = new QueryExpression("contact");
        //            Query.ColumnSet = new ColumnSet("firstname", "middlename", "lastname", "fullname", "emailaddress1");
        //            Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, consumerInfo.MobileNumber);
        //            EntityCollection entColl = service.RetrieveMultiple(Query);

        //            if (entColl.Entities.Count > 0)
        //            {
        //                objConsumerInfo.ConsumerID = entColl.Entities[0].Id;

        //                if (entColl.Entities[0].Contains("fullname"))
        //                    objConsumerInfo.Name = entColl.Entities[0].GetAttributeValue<string>("fullname");

        //                if (entColl.Entities[0].Contains("emailaddress1"))
        //                    objConsumerInfo.Email = entColl.Entities[0].GetAttributeValue<string>("emailaddress1");

        //                Query = new QueryExpression("hil_address");
        //                Query.ColumnSet = new ColumnSet("hil_state", "hil_pincode", "hil_fulladdress", "hil_district", "hil_city");
        //                Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, customerGuid);
        //                Query.AddOrder("createdon", OrderType.Descending);
        //                Query.TopCount = 1;
        //                Entity entAddressColl = service.RetrieveMultiple(Query).Entities[0];

        //                if (entAddressColl.Contains("hil_state"))
        //                    objConsumerInfo.AddressInfo.State = entAddressColl.GetAttributeValue<EntityReference>("hil_state").Name;

        //                if (entAddressColl.Contains("hil_pincode"))
        //                    objConsumerInfo.AddressInfo.Pincode = entAddressColl.GetAttributeValue<EntityReference>("hil_pincode").Name;

        //                if (entAddressColl.Contains("hil_fulladdress"))
        //                    objConsumerInfo.AddressInfo.Address = entAddressColl.GetAttributeValue<string>("hil_fulladdress");

        //                if (entAddressColl.Contains("hil_district"))
        //                    objConsumerInfo.AddressInfo.District = entAddressColl.GetAttributeValue<EntityReference>("hil_district").Name;

        //                if (entAddressColl.Contains("hil_city"))
        //                    objConsumerInfo.AddressInfo.City = entAddressColl.GetAttributeValue<EntityReference>("hil_city").Name;

        //                string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
        //                              <entity name='msdyn_customerasset'>
        //                                <attribute name='hil_modelname' />
        //                                <attribute name='hil_productcategory' />
        //                                <attribute name='msdyn_name' />
        //                                <attribute name='hil_productsubcategory' />
        //                                <attribute name='hil_purchasedate' />
        //                                <attribute name='msdyn_product' />
        //                                <attribute name='hil_invoiceno' />
        //                                <attribute name='hil_invoicevalue' />
        //                                <attribute name='hil_purchasedfrom' />
        //                                <attribute name='hil_retailerpincode' />
        //                                <order attribute='createdon' descending='true' />
        //                                <filter type='and'>
        //                                  <condition attribute='hil_customer' operator='eq' value='{customerGuid}' />
        //                                  <condition attribute='statecode' operator='eq' value='0' />
        //                                </filter>
        //                                <link-entity name='product' from='productid' to='msdyn_product' link-type='inner' alias='ag'>
        //                                  <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='ah'>
        //                                    <filter type='and'>
        //                                      <condition attribute='hil_eligibleforamc' operator='eq' value='1' />
        //                                    </filter>
        //                                  </link-entity>
        //                                </link-entity>
        //                              </entity>
        //                            </fetch>";

        //                EntityCollection entCollProduct = service.RetrieveMultiple(new FetchExpression(fetchXML));

        //                if (entCollProduct.Entities.Count > 0)
        //                {
        //                    foreach (Entity ent in entCollProduct.Entities)
        //                    {
        //                        AssetWarrantyInfo objAssetWaarrantyInfo = new AssetWarrantyInfo();
        //                        AssetInfo objAssetInfo = new AssetInfo();
        //                        List<WarrantyLineInfo> lstWarrantyLineInfo = new List<WarrantyLineInfo>();

        //                        Guid ProductGuid = ent.Id;

        //                        string hil_modelname = string.Empty;
        //                        if (ent.Contains("hil_modelname"))
        //                            hil_modelname = ent.GetAttributeValue<string>("hil_modelname");

        //                        string hil_productcategory = string.Empty;
        //                        if (ent.Contains("hil_productcategory"))
        //                            hil_productcategory = ent.GetAttributeValue<EntityReference>("hil_productcategory").Name;

        //                        string hil_productsubcategory = string.Empty;
        //                        if (ent.Contains("hil_productsubcategory"))
        //                            hil_productsubcategory = ent.GetAttributeValue<EntityReference>("hil_productsubcategory").Name;

        //                        string hil_productserialno = string.Empty;
        //                        if (ent.Contains("msdyn_name"))
        //                            hil_productserialno = ent.GetAttributeValue<string>("msdyn_name");

        //                        string modelCode = string.Empty;
        //                        if (ent.Contains("msdyn_product"))
        //                            modelCode = ent.GetAttributeValue<EntityReference>("msdyn_product").Name;

        //                        DateTime hil_purchasedate = DateTime.Now;
        //                        if (ent.Contains("hil_purchasedate"))
        //                            hil_purchasedate = ent.GetAttributeValue<DateTime>("hil_purchasedate");

        //                        string hil_invoiceno = string.Empty;
        //                        if (ent.Contains("hil_invoiceno"))
        //                            hil_invoiceno = ent.GetAttributeValue<string>("hil_invoiceno");

        //                        string hil_invoicevalue = string.Empty;
        //                        if (ent.Contains("hil_invoicevalue"))
        //                            hil_invoicevalue = ent.GetAttributeValue<string>("hil_invoicevalue");

        //                        string hil_purchasedfrom = string.Empty;
        //                        if (ent.Contains("hil_purchasedfrom"))
        //                            hil_purchasedfrom = ent.GetAttributeValue<string>("hil_purchasedfrom");

        //                        string hil_retailerpincode = string.Empty;
        //                        if (ent.Contains("hil_retailerpincode"))
        //                            hil_retailerpincode = ent.GetAttributeValue<string>("hil_retailerpincode");

        //                        objAssetInfo.ProductCategory = hil_productcategory;
        //                        objAssetInfo.ProductSubcategory = hil_productsubcategory;
        //                        objAssetInfo.ModelName = hil_modelname;
        //                        objAssetInfo.ModelNumber = modelCode;
        //                        objAssetInfo.SerialNumber = hil_productserialno;
        //                        objAssetInfo.DOP = hil_purchasedate.ToString();
        //                        objAssetInfo.InvoiceNumber = hil_invoiceno;
        //                        objAssetInfo.InvoiceValue = hil_invoicevalue;
        //                        objAssetInfo.PurchaseFrom = hil_purchasedfrom;
        //                        objAssetInfo.PurchaseLocation = hil_retailerpincode;

        //                        Query = new QueryExpression("hil_unitwarranty");
        //                        Query.ColumnSet = new ColumnSet("hil_warrantystartdate", "hil_warrantyenddate", "hil_warrantytemplate");
        //                        Query.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, ProductGuid);
        //                        EntityCollection entCollUnitwarranty = service.RetrieveMultiple(Query);

        //                        if (entColl.Entities.Count > 0)
        //                        {
        //                            foreach (Entity entwarranty in entCollUnitwarranty.Entities)
        //                            {
        //                                WarrantyLineInfo objWarrantyLineInfo = new WarrantyLineInfo();

        //                                DateTime hil_warrantystartdate = DateTime.Now;
        //                                if (entwarranty.Contains("hil_warrantystartdate"))
        //                                    hil_warrantystartdate = entwarranty.GetAttributeValue<DateTime>("hil_warrantystartdate");

        //                                DateTime hil_warrantyenddate = DateTime.Now;
        //                                if (entwarranty.Contains("hil_warrantyenddate"))
        //                                    hil_warrantyenddate = entwarranty.GetAttributeValue<DateTime>("hil_warrantyenddate");

        //                                objWarrantyLineInfo.StartDate = hil_warrantystartdate.ToString();
        //                                objWarrantyLineInfo.EndDate = hil_warrantyenddate.ToString();

        //                                Guid Guidwarrantytemplate = Guid.Empty;
        //                                if (entwarranty.Contains("hil_warrantytemplate"))
        //                                    Guidwarrantytemplate = entwarranty.GetAttributeValue<EntityReference>("hil_warrantytemplate").Id;

        //                                if (Guidwarrantytemplate != Guid.Empty)
        //                                {
        //                                    Entity entwarrantytemplate = service.Retrieve("hil_warrantytemplate", Guidwarrantytemplate, new ColumnSet("hil_warrantyperiod", "hil_type", "hil_description"));

        //                                    if (entwarrantytemplate.Attributes.Count > 0)
        //                                    {
        //                                        string WarrantyPeriod = string.Empty;
        //                                        if (entwarrantytemplate.Contains("hil_warrantyperiod"))
        //                                            WarrantyPeriod = entwarrantytemplate.GetAttributeValue<Int32>("hil_warrantyperiod").ToString();

        //                                        string WarrantyDescription = string.Empty;
        //                                        if (entwarrantytemplate.Contains("hil_description"))
        //                                            WarrantyDescription = entwarrantytemplate.GetAttributeValue<string>("hil_description");

        //                                        string WarrantyType = string.Empty;
        //                                        if (entwarrantytemplate.Contains("hil_type"))
        //                                        {
        //                                            WarrantyType = entwarrantytemplate.FormattedValues["hil_type"].ToString();
        //                                        }
        //                                        objWarrantyLineInfo.WarrantyPeriod = WarrantyPeriod;
        //                                        objWarrantyLineInfo.WarrantyCoverage = WarrantyDescription;
        //                                        objWarrantyLineInfo.WarrantyType = WarrantyType;
        //                                        //objWarrantyLineInfo.WarrantyTCLink = "";
        //                                        lstWarrantyLineInfo.Add(objWarrantyLineInfo);
        //                                    }
        //                                }
        //                            }
        //                        }
        //                        objAssetWaarrantyInfo.AssetInfo = objAssetInfo;
        //                        objAssetWaarrantyInfo.WarrantyLineInfo = lstWarrantyLineInfo;
        //                        lstAssetWaarrantyInfo.Add(objAssetWaarrantyInfo);
        //                    }
        //                }
        //                objConsumerInfo.ConsumerID = customerGuid;
        //                objConsumerInfo.MobileNumber = consumerInfo.MobileNumber;
        //                objConsumerInfo.Email = emailaddress1;
        //                objConsumerInfo.Name = fullname;
        //                objConsumerInfo.AddressInfo = new AddressInfo
        //                {
        //                    Address = hil_fulladdress,
        //                    Pincode = hil_pincode,
        //                    State = hil_state,
        //                    District = hil_district,
        //                    City = hil_city
        //                };
        //                objConsumerVerification = new ConsumerVerification
        //                {
        //                    ConsumerInfo = objConsumerInfo,
        //                    AssetWaarrantyInfo = lstAssetWaarrantyInfo,
        //                    Result = new Result { ResultStatus = "200", ResultMessage = "Success" },
        //                    SourceType = consumerInfo.SourceType,
        //                    SourceCode = consumerInfo.SourceCode
        //                };

        //                string output = JsonConvert.SerializeObject(objConsumerVerification);
        //                Console.WriteLine(output);
        //            }
        //            else
        //            {
        //                objConsumerVerification.Result = new Result { ResultStatus = "200", ResultMessage = "Mobile Number Not Registered in D365" };
        //                return objConsumerVerification;
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        objConsumerVerification.Result = new Result { ResultStatus = "Error", ResultMessage = ex.Message };
        //        return objConsumerVerification;
        //    }
        //    return objConsumerVerification;
        //}
        //#endregion Consumer Verification

        //#region Customer Registration
        //public ConsumerInfo CustomerRegistration(ConsumerInfo ConsumerInfo, Char Flag = 'I', IOrganizationService service)
        //{
        //    //ConsumerInfo ConsumerInfo = new ConsumerInfo();
        //    Guid consumerGuId = Guid.Empty;
        //    Guid PincodeGuid = Guid.Empty;
        //    try
        //    {
        //        if (service != null)
        //        {
        //            if (string.IsNullOrWhiteSpace(ConsumerInfo.MobileNumber))
        //            {
        //                ConsumerInfo.Result = new Result { ResultStatus = "204", ResultMessage = "No Content : Mobile Number is required." };
        //                return ConsumerInfo;
        //            }
        //            if (ConsumerInfo.SourceType == null || ConsumerInfo.SourceCode == null)
        //            {
        //                ConsumerInfo.Result = new Result { ResultStatus = "204", ResultMessage = "No Content : Source of Registration is required. Please pass <4> for Whatsapp <5> for IoT Platform <7> for eCommerce<8> for Chatbot" };
        //                return ConsumerInfo;
        //            }
        //            if (string.IsNullOrWhiteSpace(ConsumerInfo.AddressInfo.Pincode))
        //            {
        //                ConsumerInfo.Result = new Result { ResultStatus = "204", ResultMessage = "PIN Code is required." };
        //                return ConsumerInfo;
        //            }
        //            else
        //            {
        //                QueryExpression query = new QueryExpression("hil_pincode");
        //                query.ColumnSet = new ColumnSet("hil_pincodeid", "hil_name");
        //                query.Criteria = new FilterExpression(LogicalOperator.And);
        //                query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, ConsumerInfo.AddressInfo.Pincode);
        //                EntityCollection entcoll = service.RetrieveMultiple(query);
        //                if (entcoll.Entities.Count > 0)
        //                {
        //                    PincodeGuid = entcoll.Entities[0].GetAttributeValue<Guid>("hil_pincodeid");
        //                }
        //                else
        //                {
        //                    ConsumerInfo.Result = new Result { ResultStatus = "204", ResultMessage = "No PIN Code found." };
        //                    return ConsumerInfo;
        //                }
        //            }

        //            //Checking Mobile Number already exist
        //            QueryExpression qsContact = new QueryExpression("contact");
        //            qsContact.ColumnSet = new ColumnSet("contactid", "fullname", "emailaddress1", "hil_salutation");
        //            ConditionExpression condExp = new ConditionExpression("mobilephone", ConditionOperator.Equal, ConsumerInfo.MobileNumber);
        //            qsContact.Criteria.AddCondition(condExp);

        //            EntityCollection entColConsumer = service.RetrieveMultiple(qsContact);

        //            if (Flag == 'U')
        //            {
        //                Entity entConsumer = new Entity("contact");
        //                entConsumer.Id = entColConsumer.Entities[0].Id;
        //                entConsumer["hil_consent"] = ConsumerInfo.TC == true ? true : false;
        //                entConsumer["hil_subscribeformessagingservice"] = ConsumerInfo.SocialMediaConsent == true ? true : false;
        //                service.Update(entConsumer);
        //            }
        //            else
        //            {
        //                if (entColConsumer.Entities.Count > 0) //Consumer already Exists in D365 Database
        //                {
        //                    consumerGuId = entColConsumer.Entities[0].Id;
        //                    ConsumerInfo.ConsumerID = consumerGuId;
        //                    ConsumerInfo.MobileNumber = ConsumerInfo.MobileNumber;
        //                    ConsumerInfo.Name = entColConsumer.Entities[0].GetAttributeValue<string>("fullname");
        //                    ConsumerInfo.Email = entColConsumer.Entities[0].GetAttributeValue<string>("emailaddress1");

        //                    ConsumerInfo.Result = new Result { ResultStatus = "208", ResultMessage = "Already Reported" };
        //                }
        //                if (consumerGuId == Guid.Empty) //Creating Consumer in D365 Database
        //                {
        //                    Entity entConsumer = new Entity("contact");
        //                    entConsumer["mobilephone"] = ConsumerInfo.MobileNumber;

        //                    if (string.IsNullOrWhiteSpace(ConsumerInfo.Name))
        //                    {
        //                        entConsumer["firstname"] = "UNDEF-IoT-" + ConsumerInfo.MobileNumber;
        //                    }
        //                    else
        //                    {
        //                        string[] consumerName = ConsumerInfo.Name.Split(' ');
        //                        if (consumerName.Length >= 1)
        //                        {
        //                            entConsumer["firstname"] = consumerName[0];
        //                            if (consumerName.Length == 3)
        //                            {
        //                                entConsumer["middlename"] = consumerName[1];
        //                                entConsumer["lastname"] = consumerName[2];
        //                            }
        //                            if (consumerName.Length == 2)
        //                            {
        //                                entConsumer["lastname"] = consumerName[1];
        //                            }
        //                        }
        //                        else
        //                        {
        //                            entConsumer["firstname"] = ConsumerInfo.Name;
        //                        }
        //                    }

        //                    if (!string.IsNullOrWhiteSpace(ConsumerInfo.SourceCode))
        //                    {
        //                        entConsumer["hil_channel"] = ConsumerInfo.SourceCode;
        //                    }

        //                    if (!string.IsNullOrWhiteSpace(ConsumerInfo.Email))
        //                    {
        //                        entConsumer["emailaddress1"] = ConsumerInfo.Email;
        //                    }
        //                    entConsumer["hil_consent"] = ConsumerInfo.TC == true ? true : false;

        //                    entConsumer["hil_subscribeformessagingservice"] = ConsumerInfo.SocialMediaConsent == true ? true : false;

        //                    //if (consumer.SourceOfCreation != null)
        //                    //{
        //                    //    entConsumer["hil_consumersource"] = new OptionSetValue(consumer.SourceOfCreation.Value);
        //                    //}

        //                    entConsumer["hil_customertype"] = new OptionSetValue(1); //<1> for Consumer 

        //                    consumerGuId = service.Create(entConsumer);

        //                    Result result = CreateAddress(consumerGuId, PincodeGuid, service);

        //                    ConsumerInfo.Result = result;
        //                }
        //            }
        //        }
        //        else
        //        {
        //            ConsumerInfo.Result = new Result { ResultStatus = "503", ResultMessage = "D365 Service Unavailable" };
        //            return ConsumerInfo;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        ConsumerInfo.Result = new Result { ResultStatus = "500", ResultMessage = "D365 Internal Server Error : " + ex.Message.ToUpper() };
        //        return ConsumerInfo;
        //    }
        //    return ConsumerInfo;
        //}
        //public Result CreateAddress(Guid consumerGuId, Guid pincodeGuid, IOrganizationService service)
        //{
        //    Result result = new Result();
        //    QueryExpression query;
        //    EntityCollection entcoll;
        //    EntityReference businessGeo = null;
        //    EntityReference district = null;

        //    try
        //    {

        //        if (service != null)
        //        {
        //            int addressType;
        //            query = new QueryExpression("hil_address");
        //            query.ColumnSet = new ColumnSet(false);
        //            query.Criteria = new FilterExpression(LogicalOperator.And);
        //            query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, consumerGuId);
        //            entcoll = service.RetrieveMultiple(query);
        //            if (entcoll.Entities.Count > 0)
        //            {
        //                addressType = 2;
        //            }
        //            else
        //            {
        //                addressType = 1;
        //            }

        //            query = new QueryExpression("hil_businessmapping");
        //            query.ColumnSet = new ColumnSet("hil_pincode", "hil_district");
        //            query.Criteria = new FilterExpression(LogicalOperator.And);
        //            query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, pincodeGuid);
        //            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
        //            query.AddOrder("createdon", OrderType.Ascending);
        //            query.TopCount = 1;

        //            entcoll = service.RetrieveMultiple(query);
        //            if (entcoll.Entities.Count > 0)
        //            {
        //                businessGeo = entcoll.Entities[0].ToEntityReference();
        //                district = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_district");
        //            }
        //            else
        //            {
        //                result = new Result { ResultStatus = "204", ResultMessage = "PIN Code does not exist." };
        //                return result;
        //            }
        //            hil_address entObj = new hil_address();
        //            entObj.hil_Customer = new EntityReference("contact", consumerGuId);
        //            if (businessGeo != null)
        //            {
        //                entObj.hil_BusinessGeo = businessGeo;
        //            }
        //            if (district != null)
        //            {
        //                entObj.hil_District = district;
        //            }
        //            entObj.hil_AddressType = new OptionSetValue(addressType);
        //            Guid AddressGuid = service.Create(entObj);
        //            if (AddressGuid != Guid.Empty)
        //            {
        //                result = new Result { ResultStatus = "200", ResultMessage = "OK" };
        //            }
        //            else
        //            {
        //                result = new Result { ResultStatus = "204", ResultMessage = "Something went wrong." };
        //            }
        //        }
        //        else
        //        {
        //            result = new Result { ResultStatus = "503", ResultMessage = "D365 Service Unavailable" };
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        result = new Result { ResultStatus = "500", ResultMessage = "D365 Internal Server Error : " + ex.Message.ToUpper() };
        //    }
        //    return result;
        //}

        //#endregion Customer Registration

        //#region Product Hierarchy
        //public ModelInfoList ProductHierarchy(SourceInfo sourceInfo, IOrganizationService service)
        //{
        //    ModelInfoList ModelInfoList = new ModelInfoList();
        //    ModelInfoList.SourceCode = sourceInfo.SourceCode;
        //    ModelInfoList.SourceType = sourceInfo.SourceType;

        //    List<ModelInfo> lstModelInfo = new List<ModelInfo>();

        //    EntityCollection entcoll;
        //    try
        //    {
        //        string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
        //                      <entity name='product'>
        //                        <attribute name='name' />
        //                        <attribute name='productnumber' />
        //                        <attribute name='description' />
        //                        <attribute name='statecode' />
        //                        <attribute name='productstructure' />
        //                        <attribute name='productid' />
        //                        <attribute name='hil_materialgroup' />
        //                        <attribute name='hil_division' />
        //                        <order attribute='productnumber' descending='false' />
        //                        <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='al'>
        //                          <filter type='and'>
        //                            <condition attribute='hil_eligibleforamc' operator='eq' value='1' />
        //                          </filter>
        //                        </link-entity>
        //                      </entity>
        //                    </fetch>";

        //        entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
        //        if (entcoll.Entities.Count == 0)
        //        {
        //            ModelInfoList.Result = new Result { ResultStatus = "204", ResultMessage = "Not Any Eligible AMC Product found" };
        //            return ModelInfoList;
        //        }
        //        else
        //        {
        //            foreach (Entity ent in entcoll.Entities)
        //            {
        //                ModelInfo objModelInfo = new ModelInfo();
        //                //objAMCProduct.ModelId = ent.Id;
        //                if (ent.Attributes.Contains("name"))
        //                {
        //                    objModelInfo.ModelNumber = ent.GetAttributeValue<string>("name");
        //                }
        //                if (ent.Attributes.Contains("description"))
        //                {
        //                    objModelInfo.ModelName = ent.GetAttributeValue<string>("description");
        //                }
        //                if (ent.Attributes.Contains("hil_division"))
        //                {
        //                    objModelInfo.ProductCategory = ent.GetAttributeValue<EntityReference>("hil_division").Name;
        //                    //objAMCProduct.ProductCategoryId = ent.GetAttributeValue<EntityReference>("hil_division").Id;
        //                }
        //                if (ent.Attributes.Contains("hil_materialgroup"))
        //                {
        //                    objModelInfo.ProductSubcategory = ent.GetAttributeValue<EntityReference>("hil_materialgroup").Name;
        //                    // objAMCProduct.ProductSubcategoryId = ent.GetAttributeValue<EntityReference>("hil_materialgroup").Id;
        //                }
        //                lstModelInfo.Add(objModelInfo);
        //            }

        //            ModelInfoList.ModelInfo = lstModelInfo;
        //            ModelInfoList.Result = new Result { ResultStatus = "200", ResultMessage = "Success" };
        //            return ModelInfoList;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        ModelInfoList.Result = new Result { ResultStatus = "500", ResultMessage = ex.Message };
        //        return ModelInfoList;
        //    }
        //}
        //#endregion Product Hierarchy

        //#region Validate Serial Number
        //public AMCPlanInfoList CustomerAssetValidation(SerialNumberInfo SlNoInfo, IOrganizationService service)
        //{
        //    AMCPlanInfoList objAMCPlanInfoList = new AMCPlanInfoList();
        //    objAMCPlanInfoList.SourceCode = SlNoInfo.SourceCode;
        //    objAMCPlanInfoList.SourceType = SlNoInfo.SourceType;
        //    try
        //    {
        //        string ModelCode = IsSerialNumberExistInSAP(SlNoInfo.SerialNumber, service);
        //        Guid customerGuid = getCustomerGuid(SlNoInfo.MobileNumber, service);

        //        if (ModelCode == null)
        //        {
        //            objAMCPlanInfoList.Result = new Result { ResultStatus = "200", ResultMessage = "Serial Number Not Exist in SAP." };
        //            return objAMCPlanInfoList;
        //        }

        //        if (customerGuid == Guid.Empty)
        //        {
        //            objAMCPlanInfoList.Result = new Result { ResultStatus = "200", ResultMessage = "Mobile Number Not Exist in D365." };
        //            return objAMCPlanInfoList;
        //        }

        //        string fetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //              <entity name='msdyn_customerasset'>
        //                <attribute name='createdon' />
        //                <attribute name='msdyn_product' />
        //                <attribute name='msdyn_name' />
        //                <attribute name='hil_productsubcategorymapping' />
        //                <attribute name='hil_productcategory' />
        //                <attribute name='msdyn_customerassetid' />
        //                 <attribute name='msdyn_product' />
        //                <attribute name='hil_modelname' />
        //                <order attribute='createdon' descending='true' />
        //                <filter type='and'>
        //                  <condition attribute='hil_customer' operator='eq' uiname='Nanhe Siddique' uitype='contact' value='{ customerGuid }' />
        //                  <condition attribute='msdyn_name' operator='eq' value='{SlNoInfo.SerialNumber }' />
        //                </filter>
        //              </entity>
        //            </fetch>";
        //        EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(fetchQuery));
        //        if (entCol.Entities.Count > 0)
        //        {
        //            EntityReference msdyn_product = entCol.Entities[0].GetAttributeValue<EntityReference>("msdyn_product");
        //            List<AMCPlanInfo> AMCPlanInfo = GetAMCPlanDetails(msdyn_product.Id, service);
        //            objAMCPlanInfoList.AMCPlanInfo = AMCPlanInfo;
        //        }
        //        else
        //        {
        //            string fetchQuery2 = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                          <entity name='msdyn_customerasset'>
        //                            <attribute name='createdon' />
        //                            <attribute name='msdyn_product' />
        //                            <attribute name='msdyn_name' />
        //                            <attribute name='hil_productsubcategorymapping' />
        //                            <attribute name='hil_productcategory' />
        //                            <attribute name='msdyn_customerassetid' />
        //                            <order attribute='createdon' descending='true' />
        //                            <filter type='and'>
        //                              <condition attribute='hil_customer' operator='ne' uiname='Nanhe Siddique' uitype='contact' value='{ customerGuid }' />
        //                              <condition attribute='msdyn_name' operator='eq' value='{ SlNoInfo.SerialNumber }' />
        //                            </filter>
        //                          </entity>
        //                        </fetch>";
        //            EntityCollection entColne = service.RetrieveMultiple(new FetchExpression(fetchQuery2));
        //            if (entColne.Entities.Count > 0)
        //            {
        //                objAMCPlanInfoList.Result = new Result { ResultStatus = "200", ResultMessage = "Serail Number is Registered with another consumer." };
        //                return objAMCPlanInfoList;
        //            }
        //        }

        //        return objAMCPlanInfoList;
        //    }
        //    catch (Exception ex)
        //    {
        //        objAMCPlanInfoList.Result = new Result { ResultStatus = "Error", ResultMessage = ex.Message };
        //        return objAMCPlanInfoList;
        //    }
        //}
        //public Guid getCustomerGuid(string MobileNo, IOrganizationService service)
        //{
        //    Guid customer = Guid.Empty;
        //    string FetchContact = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                          <entity name='contact'>
        //                            <attribute name='fullname' />
        //                            <attribute name='telephone1' />
        //                            <attribute name='contactid' />
        //                            <order attribute='fullname' descending='false' />
        //                            <filter type='and'>
        //                              <condition attribute='mobilephone' operator='eq' value='{MobileNo}' />
        //                            </filter>
        //                          </entity>
        //                        </fetch>";
        //    EntityCollection colCustomer = service.RetrieveMultiple(new FetchExpression(FetchContact));
        //    if (colCustomer.Entities.Count > 0)
        //    {
        //        customer = colCustomer.Entities[0].ToEntityReference().Id;
        //    }
        //    return customer;
        //}
        //public string IsSerialNumberExistInSAP(string SerialNumber, IOrganizationService service)
        //{
        //    string ModelCode = null;
        //    using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
        //    {
        //        #region Credentials
        //        String sUserName = String.Empty;
        //        String sPassword = String.Empty;
        //        var obj2 = from _IConfig in orgContext.CreateQuery<hil_integrationconfiguration>()
        //                   where _IConfig.hil_name == "Credentials"
        //                   select new { _IConfig };
        //        foreach (var iobj2 in obj2)
        //        {
        //            if (iobj2._IConfig.hil_Username != String.Empty)
        //                sUserName = iobj2._IConfig.hil_Username;
        //            if (iobj2._IConfig.hil_Password != String.Empty)
        //                sPassword = iobj2._IConfig.hil_Password;
        //        }
        //        #endregion

        //        var obj = from _IConfig in orgContext.CreateQuery<hil_integrationconfiguration>()
        //                  where _IConfig.hil_name == "SerialNumberValidation"
        //                  select new { _IConfig.hil_Url };
        //        foreach (var iobj in obj)
        //        {
        //            if (iobj.hil_Url != null)
        //            {
        //                //String sUrl = iobj.hil_Url + SerialNumber;
        //                String sUrl = $"https://p90ci.havells.com:50001/RESTAdapter/Service/ProductDetails?IM_SERIAL_NO={SerialNumber}";
        //                WebClient webClient = new WebClient();
        //                string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(sUserName + ":" + sPassword));
        //                webClient.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
        //                webClient.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
        //                var jsonData = webClient.DownloadData(sUrl);

        //                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(SerialNumberValidation));
        //                SerialNumberValidation rootObject = (SerialNumberValidation)ser.ReadObject(new MemoryStream(jsonData));

        //                if (rootObject.EX_PRD_DET != null)//valid
        //                {
        //                    ModelCode = rootObject.EX_PRD_DET.MATNR;
        //                }
        //            }
        //        }
        //    }
        //    return ModelCode;
        //}

        //#endregion

        //#region Validate Model Number
        //public AMCPlanInfoList GetAMCDetailsBasedOnModel(ModelInfo modelInfo, IOrganizationService service)
        //{
        //    AMCPlanInfoList objAMCPlanInfoList = new AMCPlanInfoList();
        //    AssetWarrantyInfo objAssetWaarrantyInfo = new AssetWarrantyInfo();
        //    try
        //    {
        //        if (string.IsNullOrEmpty(modelInfo.ModelNumber))
        //        {
        //            objAMCPlanInfoList.Result = new Result { ResultStatus = "200", ResultMessage = "Model Number is Required" };
        //            return objAMCPlanInfoList;
        //        }
        //        string fetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //              <entity name='product'>
        //                <attribute name='name' />
        //                <attribute name='productnumber' />
        //                <attribute name='description' />
        //                <attribute name='statecode' />
        //                <attribute name='productstructure' />
        //                <attribute name='productid' />
        //                <attribute name='hil_materialgroup' />
        //                <attribute name='hil_division' />
        //                <order attribute='productnumber' descending='false' />
        //                <filter type='and'>
        //                  <condition attribute='name' operator='eq' value='{modelInfo.ModelNumber}' />
        //                </filter>
        //              </entity>
        //            </fetch>";

        //        EntityCollection entityColAMC = service.RetrieveMultiple(new FetchExpression(fetchQuery));
        //        if (entityColAMC.Entities.Count > 0)
        //        {
        //            Guid productGUID = entityColAMC.Entities[0].Id;
        //            modelInfo.ModelNumber = entityColAMC.Entities[0].GetAttributeValue<string>("name");
        //            modelInfo.ModelName = entityColAMC.Entities[0].GetAttributeValue<string>("description");

        //            if (entityColAMC.Entities[0].Contains("hil_division"))
        //            {
        //                modelInfo.ProductCategory = entityColAMC.Entities[0].GetAttributeValue<EntityReference>("hil_division").Name;
        //            }
        //            if (entityColAMC.Entities[0].Contains("hil_materialgroup"))
        //            {
        //                modelInfo.ProductSubcategory = entityColAMC.Entities[0].GetAttributeValue<EntityReference>("hil_materialgroup").Name;
        //            }
        //            List<AMCPlanInfo> AMCPlanInfo = GetAMCPlanDetails(productGUID, service);
        //            objAMCPlanInfoList.AMCPlanInfo = AMCPlanInfo;
        //        }
        //        else
        //        {
        //            objAMCPlanInfoList.Result = new Result { ResultStatus = "200", ResultMessage = "Model Number not found." };
        //            return objAMCPlanInfoList;
        //        }

        //        return objAMCPlanInfoList;
        //    }
        //    catch (Exception ex)
        //    {
        //        objAMCPlanInfoList.Result = new Result { ResultStatus = "Error", ResultMessage = ex.Message };
        //        return objAMCPlanInfoList;
        //    }
        //}
        //public List<AMCPlanInfo> GetAMCPlanDetails(Guid ProductGuid, IOrganizationService service)
        //{
        //    List<AMCPlanInfo> lstAMCPlanInfo = new List<AMCPlanInfo>();
        //    QueryExpression query;
        //    decimal? DiscPer = null;

        //    string fetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                <entity name='hil_amcdiscountmatrix'>
        //                <attribute name='hil_discper' />
        //                <filter type='and'>
        //                <condition attribute='statecode' operator='eq' value='0' />
        //                <condition attribute='hil_appliedto' operator='eq' value='E5A41AE9-CC64-ED11-9562-6045BDAC526A' />
        //                </filter>
        //                </entity>
        //                </fetch>";
        //    EntityCollection entamcdiscountmatrix = service.RetrieveMultiple(new FetchExpression(fetchQuery));
        //    if (entamcdiscountmatrix.Entities.Count > 0)
        //    {
        //        if (entamcdiscountmatrix.Entities[0].Contains("hil_discper"))
        //        {
        //            DiscPer = Convert.ToDecimal(entamcdiscountmatrix.Entities[0].GetAttributeValue<Money>("hil_discper"));
        //        }
        //    }

        //    query = new QueryExpression("hil_servicebom");
        //    query.ColumnSet = new ColumnSet("hil_product");
        //    query.Criteria.AddCondition(new ConditionExpression("hil_productcategory", ConditionOperator.Equal, ProductGuid));
        //    EntityCollection entCollProduct = service.RetrieveMultiple(query);

        //    if (entCollProduct.Entities.Count > 0)
        //    {
        //        foreach (Entity entProduct in entCollProduct.Entities)
        //        {
        //            AMCPlanInfo objAMCPlanInfo = new AMCPlanInfo();

        //            if (entProduct.Contains("hil_product"))
        //            {
        //                Guid sPart = entProduct.GetAttributeValue<EntityReference>("hil_product").Id;

        //                query = new QueryExpression("dynamicproperty");
        //                query.ColumnSet = new ColumnSet(true);
        //                query.Criteria.AddCondition(new ConditionExpression("regardingobjectid", ConditionOperator.Equal, sPart));
        //                EntityCollection entdynamicproperty = service.RetrieveMultiple(query);
        //                if (entdynamicproperty.Entities.Count > 0)
        //                {
        //                    foreach (Entity entdynamic in entdynamicproperty.Entities)
        //                    {
        //                        if (entdynamic.Contains("name"))
        //                        {
        //                            string name = entdynamic.GetAttributeValue<string>("name");
        //                            switch (name)
        //                            {
        //                                case "Coverage":
        //                                    objAMCPlanInfo.Coverage = entdynamic.GetAttributeValue<string>("description").ToString();
        //                                    break;
        //                                case "Nov-Covered":
        //                                    objAMCPlanInfo.NonCoverage = entdynamic.GetAttributeValue<string>("description").ToString();
        //                                    break;
        //                                case "PlanName":
        //                                    objAMCPlanInfo.PlanName = entdynamic.GetAttributeValue<string>("description").ToString();
        //                                    break;
        //                                case "PlanPeriod":
        //                                    objAMCPlanInfo.PlanPeriod = entdynamic.GetAttributeValue<string>("description").ToString();
        //                                    break;
        //                                case "PlanTCLink":
        //                                    objAMCPlanInfo.PlanTCLink = entdynamic.GetAttributeValue<string>("description").ToString();
        //                                    break;
        //                            }
        //                        }
        //                    }

        //                }
        //            }
        //            objAMCPlanInfo.DiscountPercent = DiscPer;
        //            lstAMCPlanInfo.Add(objAMCPlanInfo);
        //        }
        //    }
        //    return lstAMCPlanInfo;
        //}
        //public List<WarrantyLineInfo> GetWarrantyDetails(Guid ProductGuid, IOrganizationService service)
        //{
        //    List<WarrantyLineInfo> lstWarrantyLineInfo = new List<WarrantyLineInfo>();
        //    EntityCollection entColl = new EntityCollection();
        //    QueryExpression Query = new QueryExpression("hil_unitwarranty");
        //    Query.ColumnSet = new ColumnSet("hil_warrantystartdate", "hil_warrantyenddate", "hil_warrantytemplate");
        //    Query.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, ProductGuid);
        //    EntityCollection entCollUnitwarranty = service.RetrieveMultiple(Query);

        //    if (entColl.Entities.Count > 0)
        //    {
        //        foreach (Entity entwarranty in entCollUnitwarranty.Entities)
        //        {
        //            WarrantyLineInfo objWarrantyLineInfo = new WarrantyLineInfo();

        //            DateTime hil_warrantystartdate = DateTime.Now;
        //            if (entwarranty.Contains("hil_warrantystartdate"))
        //                hil_warrantystartdate = entwarranty.GetAttributeValue<DateTime>("hil_warrantystartdate");

        //            DateTime hil_warrantyenddate = DateTime.Now;
        //            if (entwarranty.Contains("hil_warrantyenddate"))
        //                hil_warrantyenddate = entwarranty.GetAttributeValue<DateTime>("hil_warrantyenddate");

        //            objWarrantyLineInfo.StartDate = hil_warrantystartdate.ToString();
        //            objWarrantyLineInfo.EndDate = hil_warrantyenddate.ToString();

        //            Guid Guidwarrantytemplate = Guid.Empty;
        //            if (entwarranty.Contains("hil_warrantytemplate"))
        //                Guidwarrantytemplate = entwarranty.GetAttributeValue<EntityReference>("hil_warrantytemplate").Id;

        //            if (Guidwarrantytemplate != Guid.Empty)
        //            {
        //                Entity entwarrantytemplate = service.Retrieve("hil_warrantytemplate", Guidwarrantytemplate, new ColumnSet("hil_warrantyperiod", "hil_type", "hil_description"));

        //                if (entwarrantytemplate.Attributes.Count > 0)
        //                {
        //                    string WarrantyPeriod = string.Empty;
        //                    if (entwarrantytemplate.Contains("hil_warrantyperiod"))
        //                        WarrantyPeriod = entwarrantytemplate.GetAttributeValue<Int32>("hil_warrantyperiod").ToString();

        //                    string WarrantyDescription = string.Empty;
        //                    if (entwarrantytemplate.Contains("hil_description"))
        //                        WarrantyDescription = entwarrantytemplate.GetAttributeValue<string>("hil_description");

        //                    string WarrantyType = string.Empty;
        //                    if (entwarrantytemplate.Contains("hil_type"))
        //                    {
        //                        WarrantyType = entwarrantytemplate.FormattedValues["hil_type"].ToString();
        //                    }
        //                    objWarrantyLineInfo.WarrantyPeriod = WarrantyPeriod;
        //                    objWarrantyLineInfo.WarrantyCoverage = WarrantyDescription;
        //                    objWarrantyLineInfo.WarrantyType = WarrantyType;
        //                    //objWarrantyLineInfo.WarrantyTCLink ="";
        //                    lstWarrantyLineInfo.Add(objWarrantyLineInfo);

        //                }
        //            }
        //        }
        //    }
        //    return lstWarrantyLineInfo;
        //}

        //#endregion
    }

    [DataContract]
    public class ConsumerInfo : SourceInfo
    {
        [DataMember]
        public Guid ConsumerID { get; set; }
        [DataMember]
        public string MobileNumber { get; set; }
        [DataMember]
        public string Email { get; set; }
        [DataMember]
        public string Name { get; set; }
        [DataMember]
        public AddressInfo AddressInfo { get; set; }
        [DataMember]
        public bool TCConsent { get; set; }
        [DataMember]
        public bool SocialMediaCommConsent { get; set; }
        [DataMember]
        public Result Result { get; set; }
    }

    public class SerialNumberInfo : SourceInfo
    {
        [DataMember]
        public string MobileNumber { get; set; }

        [DataMember]
        public string SerialNumber { get; set; }
        [DataMember]
        public Result Result { get; set; }
    }


    [DataContract]
    public class AddressInfo
    {
        [DataMember]
        public string Address { get; set; }
        [DataMember]
        public string Pincode { get; set; }
        [DataMember]
        public string State { get; set; }
        [DataMember]
        public string District { get; set; }
        [DataMember]
        public string City { get; set; }
    }
    [DataContract]
    public class ModelInfo
    {
        [DataMember]
        public string ProductCategory { get; set; }
        [DataMember]
        public string ProductSubcategory { get; set; }
        [DataMember]
        public string ModelNumber { get; set; }
        [DataMember]
        public string ModelName { get; set; }
    }
    [DataContract]
    public class AssetInfo
    {
        [DataMember]
        public string SerialNumber { get; set; }
        [DataMember]
        public string ProductCategory { get; set; }
        [DataMember]
        public string ProductSubcategory { get; set; }
        [DataMember]
        public string ModelNumber { get; set; }
        [DataMember]
        public string ModelName { get; set; }
        [DataMember]
        public string DOP { get; set; }
        [DataMember]
        public string InvoiceNumber { get; set; }
        [DataMember]
        public string InvoiceValue { get; set; }
        [DataMember]
        public string PurchaseFrom { get; set; }
        [DataMember]
        public string PurchaseLocation { get; set; }
    }
    [DataContract]
    public class WarrantyLineInfo
    {
        [DataMember]
        public string WarrantyType { get; set; }
        [DataMember]
        public string WarrantyPeriod { get; set; }
        [DataMember]
        public string WarrantyCoverage { get; set; }
        [DataMember]
        public string WarrantyTCLink { get; set; }
        [DataMember]
        public string StartDate { get; set; }
        [DataMember]
        public string EndDate { get; set; }
    }

    [DataContract]
    public class AssetWarrantyInfo
    {
        [DataMember]
        public AssetInfo AssetInfo { get; set; }
        [DataMember]
        public List<WarrantyLineInfo> WarrantyLineInfo { get; set; }
    }

    [DataContract]
    public class ConsumerVerification : SourceInfo
    {
        [DataMember]
        public ConsumerInfo ConsumerInfo { get; set; }

        [DataMember]
        public List<AssetWarrantyInfo> AssetWarrantyInfo { get; set; }

        [DataMember]
        public Result Result { get; set; }
    }

    [DataContract]
    public class SourceInfo
    {
        [DataMember]
        public string SourceCode { get; set; }

        [DataMember]
        public string SourceType { get; set; }
    }

    [DataContract]
    public class Result
    {
        [DataMember]
        public string ResultStatus { get; set; }

        [DataMember]
        public string ResultMessage { get; set; }
    }

    [DataContract]
    public class ModelInfoList : SourceInfo
    {
        [DataMember]
        public List<ModelInfo> ModelInfo { get; set; }
        [DataMember]
        public Result Result { get; set; }
    }

    [DataContract]
    public class AMCPlanInfo
    {
        [DataMember]
        public string PlanName { get; set; }
        [DataMember]
        public string PlanPeriod { get; set; }
        [DataMember]
        public decimal MRP { get; set; }
        [DataMember]
        public decimal? DiscountPercent { get; set; }
        [DataMember]
        public decimal EffectivePrice
        {
            get
            {
                if (DiscountPercent != null)
                {
                    return (MRP * DiscountPercent.Value) / 100;
                }
                else
                {
                    return MRP;
                }
            }
        }
        [DataMember]
        public string Coverage { get; set; }
        [DataMember]
        public string NonCoverage { get; set; }

        [DataMember]
        public string PlanTCLink { get; set; }
        [DataMember]
        public string EffectiveFromDate { get; set; }
        [DataMember]
        public string EffectiveToDate { get; set; }

    }

    [DataContract]
    public class AMCPlanInfoList : SourceInfo
    {
        [DataMember]
        public List<AMCPlanInfo> AMCPlanInfo { get; set; }
        [DataMember]
        public Result Result { get; set; }

    }
}
