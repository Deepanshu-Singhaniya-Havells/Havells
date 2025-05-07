using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.Serialization;
using System.Web.Services.Description;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class IoTAddressBookV1
    {
        [DataMember]
        public Guid CustomerGuid { get; set; }
        [DataMember]
        public string Pincode { get; set; }
        [DataMember]
        public string MobileNumber { get; set; }
        [DataMember]
        public bool IsDefault { get; set; }

        public List<IoTAddressBookResultV1> GetIoTAddressBook(IoTAddressBookV1 address)
        {
            List<IoTAddressBookResultV1> addressList = new List<IoTAddressBookResultV1>();
            IoTAddressBookResultV1 objAddress;
            QueryExpression query;
            EntityCollection collection;
            EntityReference erPincode = null;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (address.CustomerGuid == Guid.Empty && address.MobileNumber == null)
                {
                    objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "Customer GUID/Mobile Number is required." };
                    addressList.Add(objAddress);
                    return addressList;
                }
                if (service != null)
                {
                    if (address.Pincode != null)
                    {
                        if (address.Pincode.Trim().Length == 6)
                        {
                            query = new QueryExpression("hil_pincode");
                            query.ColumnSet = new ColumnSet(false);
                            query.Criteria = new FilterExpression(LogicalOperator.And);
                            query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, address.Pincode);
                            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            EntityCollection entcoll = service.RetrieveMultiple(query);
                            if (entcoll.Entities.Count > 0)
                            {
                                erPincode = entcoll.Entities[0].ToEntityReference();
                            }
                            else
                            {
                                objAddress = new IoTAddressBookResultV1 { PINCode = address.Pincode, StatusCode = 204, StatusDescription = "Invalid Pincode." };
                                addressList.Add(objAddress);
                                return addressList;
                            }
                        }
                        else
                        {
                            objAddress = new IoTAddressBookResultV1 { PINCode = address.Pincode, StatusCode = 204, StatusDescription = "Invalid Pincode." };
                            addressList.Add(objAddress);
                            return addressList;
                        }
                    }
                    query = new QueryExpression("hil_address");
                    query.ColumnSet = new ColumnSet("hil_customer", "hil_addressid", "hil_street3", "hil_street2",
                        "hil_street1", "hil_fulladdress", "hil_pincode", "hil_area", "hil_businessgeo", "hil_addresstype", "hil_isdefault");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);//InActive
                    if (address.CustomerGuid != Guid.Empty)
                    {
                        query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, address.CustomerGuid);
                    }
                    if (erPincode != null)
                    {
                        query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, erPincode.Id);
                    }
                    if (address.MobileNumber != null)
                    {
                        LinkEntity lnk = new LinkEntity("hil_address", "contact", "hil_customer", "contactid", JoinOperator.Inner);
                        lnk.Columns = new ColumnSet("fullname", "emailaddress1", "mobilephone", "contactid");
                        lnk.LinkCriteria.AddCondition("mobilephone", ConditionOperator.Equal, address.MobileNumber);
                        query.LinkEntities.Add(lnk);
                    }
                    //query.AddOrder("modifiedon", OrderType.Descending);
                    //query.TopCount = 10;
                    collection = service.RetrieveMultiple(query);

                    if (collection.Entities.Count == 0)
                    {
                        objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "No Address Book found." };
                        addressList.Add(objAddress);
                        return addressList;
                    }
                    if (collection.Entities != null && collection.Entities.Count > 0)
                    {
                        foreach (Entity item in collection.Entities)
                        {
                            objAddress = new IoTAddressBookResultV1();
                            if (item.Attributes.Contains("hil_street1"))
                            {
                                objAddress.AddressLine1 = item.GetAttributeValue<string>("hil_street1");
                            }
                            if (item.Attributes.Contains("hil_street2"))
                            {
                                objAddress.AddressLine2 = item.GetAttributeValue<string>("hil_street2");
                            }
                            if (item.Attributes.Contains("hil_street3"))
                            {
                                objAddress.AddressLine3 = item.GetAttributeValue<string>("hil_street3");
                            }
                            if (item.Attributes.Contains("hil_fulladdress"))
                            {
                                objAddress.FullAddress = item.GetAttributeValue<string>("hil_fulladdress");
                            }
                            if (item.Attributes.Contains("hil_addressid"))
                            {
                                objAddress.AddressGuid = item.Id;
                            }
                            if (item.Attributes.Contains("hil_businessgeo"))
                            {
                                objAddress.BizGeoGuid = item.GetAttributeValue<EntityReference>("hil_businessgeo").Id;
                            }
                            if (item.Attributes.Contains("hil_businessgeo"))
                            {
                                objAddress.BizGeoName = item.GetAttributeValue<EntityReference>("hil_businessgeo").Name;
                            }
                            if (item.Attributes.Contains("hil_pincode"))
                            {
                                objAddress.PINCodeGuid = item.GetAttributeValue<EntityReference>("hil_pincode").Id;
                            }
                            if (item.Attributes.Contains("hil_pincode"))
                            {
                                objAddress.PINCode = item.GetAttributeValue<EntityReference>("hil_pincode").Name;
                            }
                            if (item.Attributes.Contains("hil_area"))
                            {
                                objAddress.AreaGuid = item.GetAttributeValue<EntityReference>("hil_area").Id;
                            }
                            if (item.Attributes.Contains("hil_area"))
                            {
                                objAddress.Area = item.GetAttributeValue<EntityReference>("hil_area").Name;
                            }
                            if (item.Attributes.Contains("hil_addresstype"))
                            {
                                objAddress.AddressType = item.FormattedValues["hil_addresstype"].ToString();
                            }
                            if (item.Attributes.Contains("hil_addresstype"))
                            {
                                objAddress.AddressTypeEnum = item.GetAttributeValue<OptionSetValue>("hil_addresstype").Value;
                            }
                            if (item.Attributes.Contains("hil_isdefault"))
                            {
                                objAddress.IsDefault = item.GetAttributeValue<bool>("hil_isdefault");
                            }
                            if (item.Attributes.Contains("contact1.fullname"))
                            {
                                objAddress.CustomerName = item.GetAttributeValue<AliasedValue>("contact1.fullname").Value.ToString();
                            }
                            if (item.Attributes.Contains("contact1.emailaddress1"))
                            {
                                objAddress.EmailAddress = item.GetAttributeValue<AliasedValue>("contact1.emailaddress1").Value.ToString();
                            }
                            if (item.Attributes.Contains("contact1.contactid"))
                            {
                                objAddress.CustomerGuid = (Guid)(item.GetAttributeValue<AliasedValue>("contact1.contactid").Value);
                            }
                            if (item.Attributes.Contains("contact1.mobilephone"))
                            {
                                objAddress.MobileNumber = item.GetAttributeValue<AliasedValue>("contact1.mobilephone").Value.ToString();
                            }
                            objAddress.StatusCode = 200;
                            objAddress.StatusDescription = "OK";
                            addressList.Add(objAddress);
                        }
                    }
                    return addressList;
                }
                else
                {
                    objAddress = new IoTAddressBookResultV1 { StatusCode = 503, StatusDescription = "D365 Service Unavailable" };
                    addressList.Add(objAddress);
                }
            }
            catch (Exception ex)
            {
                objAddress = new IoTAddressBookResultV1 { StatusCode = 500, StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                addressList.Add(objAddress);
            }
            return addressList;
        }
        //public IoTPINCodes IoTValidatePINCode(IoTPINCodes pinCode)
        //{
        //    IoTPINCodes objPINCode = null;
        //    try
        //    {
        //        IOrganizationService service = ConnectToCRM.GetOrgService();
        //        if (pinCode.PINCode.Trim().Length == 0)
        //        {
        //            objPINCode = new IoTPINCodes { StatusCode = 204, StatusDescription = "PIN Code is required." };
        //            return objPINCode;
        //        }
        //        if (service != null)
        //        {
        //            QueryExpression query = new QueryExpression("hil_pincode");
        //            query.ColumnSet = new ColumnSet("hil_pincodeid", "hil_name");
        //            query.Criteria = new FilterExpression(LogicalOperator.And);
        //            query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, pinCode.PINCode);
        //            EntityCollection entcoll = service.RetrieveMultiple(query);
        //            if (entcoll.Entities.Count == 0)
        //            {
        //                objPINCode = new IoTPINCodes { StatusCode = 204, StatusDescription = "No PIN Code found." };
        //                return objPINCode;
        //            }
        //            else
        //            {
        //                objPINCode = new IoTPINCodes { PINCode = entcoll.Entities[0].GetAttributeValue<string>("hil_name"), PINCodeGuid = entcoll.Entities[0].GetAttributeValue<Guid>("hil_pincodeid"), StatusCode = 200, StatusDescription = "OK." };
        //                return objPINCode;
        //            }
        //        }
        //        else
        //        {
        //            objPINCode = new IoTPINCodes { StatusCode = 503, StatusDescription = "D365 Service Unavailable" };
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        objPINCode = new IoTPINCodes { StatusCode = 500, StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
        //    }
        //    return objPINCode;
        //}
        //public List<IoTAreas> IoTGetAreasByPinCode(IoTPINCodes pinCode)
        //{
        //    IoTAreas objArea = null;
        //    List<IoTAreas> lstAreas = null;
        //    try
        //    {
        //        IOrganizationService service = ConnectToCRM.GetOrgService();
        //        if (service != null)
        //        {
        //            QueryExpression query = new QueryExpression("hil_businessmapping");
        //            query.ColumnSet = new ColumnSet("hil_pincode", "hil_area");
        //            query.Criteria = new FilterExpression(LogicalOperator.And);
        //            query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, pinCode.PINCodeGuid);
        //            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
        //            EntityCollection entcoll = service.RetrieveMultiple(query);
        //            if (entcoll.Entities.Count == 0)
        //            {
        //                objArea = new IoTAreas { PINCodeGuid = pinCode.PINCodeGuid, StatusCode = 204, StatusDescription = "No Area found." };
        //                lstAreas = new List<IoTAreas>();
        //                lstAreas.Add(objArea);
        //            }
        //            else
        //            {
        //                lstAreas = new List<IoTAreas>();
        //                foreach (Entity ent in entcoll.Entities)
        //                {
        //                    lstAreas.Add(new IoTAreas { PINCodeGuid = pinCode.PINCodeGuid, AreaName = ent.GetAttributeValue<EntityReference>("hil_area").Name, AreaGuid = ent.GetAttributeValue<EntityReference>("hil_area").Id, StatusCode = 200, StatusDescription = "OK." });
        //                }
        //            }
        //        }
        //        else
        //        {
        //            objArea = new IoTAreas { StatusCode = 503, StatusDescription = "D365 Service Unavailable" };
        //            lstAreas = new List<IoTAreas>();
        //            lstAreas.Add(objArea);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        objArea = new IoTAreas { StatusCode = 500, StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
        //        lstAreas = new List<IoTAreas>();
        //        lstAreas.Add(objArea);
        //    }
        //    return lstAreas;
        //}
        public IoTAddressBookResultV1 IoTCreateAddress(IoTAddressBookResultV1 addressData)
        {
            IoTAddressBookResultV1 objAddress = null;
            QueryExpression query;
            EntityCollection entcoll;
            EntityReference businessGeo = null;
            EntityReference district = null;

            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (addressData.CustomerGuid == Guid.Empty)
                {
                    objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "Customer GUID is required." };
                    return objAddress;
                }
                if (addressData.AddressLine1 == null)
                {
                    objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "Address Line 1 is required." };
                    return objAddress;
                }
                if (addressData.AddressLine1.Trim().Length == 0)
                {
                    objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "Address Line 1 is required." };
                    return objAddress;
                }
                if (addressData.PINCodeGuid == Guid.Empty)
                {
                    objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "PIN Code is required." };
                    return objAddress;
                }

                if (service != null)
                {
                    query = new QueryExpression("hil_businessmapping");
                    query.ColumnSet = new ColumnSet("hil_pincode", "hil_district");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, addressData.PINCodeGuid);
                    query.Criteria.AddCondition("hil_district", ConditionOperator.NotNull);
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                    if (addressData.AreaGuid != Guid.Empty)
                    {
                        query.Criteria.AddCondition("hil_area", ConditionOperator.Equal, addressData.AreaGuid);
                    }
                    query.AddOrder("createdon", OrderType.Ascending);
                    query.TopCount = 1;

                    entcoll = service.RetrieveMultiple(query);
                    if (entcoll.Entities.Count > 0)
                    {
                        businessGeo = entcoll.Entities[0].ToEntityReference();
                        district = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_district");
                    }
                    else
                    {
                        objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "PIN Code or Area does not exist." };
                        return objAddress;
                    }
                    hil_address entObj = new hil_address();
                    entObj.hil_Street1 = addressData.AddressLine1;
                    if (addressData.AddressLine2 != null)
                    {
                        entObj.hil_Street2 = addressData.AddressLine2;
                    }
                    if (addressData.AddressLine3 != null)
                    {
                        entObj.hil_Street3 = addressData.AddressLine3;
                    }
                    entObj.hil_Customer = new EntityReference("contact", addressData.CustomerGuid);
                    if (businessGeo != null)
                    {
                        entObj.hil_BusinessGeo = businessGeo;
                    }
                    if (district != null)
                    {
                        entObj.hil_District = district;
                    }
                    int[] intAddressType = { 1, 2, 3 };
                    int AddressTypeEnum = intAddressType.Contains(addressData.AddressTypeEnum) ? addressData.AddressTypeEnum : 3;
                    entObj.hil_AddressType = new OptionSetValue(AddressTypeEnum);
                    if (addressData.IsDefault)
                    {
                        RemoveDefaultAddress(service, addressData.CustomerGuid);
                        entObj["hil_isdefault"] = addressData.IsDefault;
                    }
                    objAddress = new IoTAddressBookResultV1();
                    objAddress.AddressGuid = service.Create(entObj);
                    if (objAddress.AddressGuid != Guid.Empty)
                    {
                        objAddress.CustomerGuid = addressData.CustomerGuid;
                        objAddress.AddressLine1 = addressData.AddressLine1;
                        objAddress.AddressLine2 = addressData.AddressLine2;
                        objAddress.AddressLine3 = addressData.AddressLine3;
                        objAddress.PINCodeGuid = addressData.PINCodeGuid;
                        objAddress.AddressTypeEnum = AddressTypeEnum;
                        if (AddressTypeEnum == 1)
                        {
                            objAddress.AddressType = "Home";
                        }
                        else if (AddressTypeEnum == 2)
                        {
                            objAddress.AddressType = "Office";
                        }
                        else
                        {
                            objAddress.AddressType = "Other";
                        }
                        objAddress.IsDefault = addressData.IsDefault;
                        objAddress.StatusCode = 200;
                        objAddress.StatusDescription = "OK.";
                    }
                    else
                    {
                        objAddress.CustomerGuid = addressData.CustomerGuid;
                        objAddress.StatusCode = 204;
                        objAddress.StatusDescription = "Something went wrong.";
                    }
                }
                else
                {
                    objAddress = new IoTAddressBookResultV1 { StatusCode = 503, StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                objAddress = new IoTAddressBookResultV1 { StatusCode = 500, StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return objAddress;
        }
        public IoTAddressBookResultV1 IoTUpdateAddress(IoTAddressBookResultV1 addressData)
        {
            IoTAddressBookResultV1 objAddress = null;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (addressData.AddressGuid == Guid.Empty)
                {
                    objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "Address GUID is required." };
                    return objAddress;
                }
                if (addressData.CustomerGuid == Guid.Empty)
                {
                    objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "Customer GUID is required." };
                    return objAddress;
                }
                if (addressData.AddressLine1.Trim().Length == 0)
                {
                    objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "Address Line 1 is required." };
                    return objAddress;
                }

                if (service != null)
                {
                    Entity address = service.Retrieve("hil_address", addressData.AddressGuid, new ColumnSet("hil_customer"));

                    if (address.Contains("hil_customer"))
                    {
                        if (address.GetAttributeValue<EntityReference>("hil_customer").Id != addressData.CustomerGuid)
                        {
                            objAddress = new IoTAddressBookResultV1 { AddressGuid = addressData.AddressGuid, StatusCode = 204, StatusDescription = "Address ID does not belong to Customer." };
                            return objAddress;
                        }
                    }
                    else
                    {
                        objAddress = new IoTAddressBookResultV1 { AddressGuid = addressData.AddressGuid, StatusCode = 204, StatusDescription = "Address ID does not belong to Customer." };
                        return objAddress;
                    }
                    //***changed by Saurabh ends here.
                    hil_address entObj = new hil_address();
                    entObj.Id = addressData.AddressGuid;
                    entObj.hil_Street1 = addressData.AddressLine1;
                    if (addressData.AddressLine2 != null)
                    {
                        entObj.hil_Street2 = addressData.AddressLine2;
                    }
                    if (addressData.AddressLine3 != null)
                    {
                        entObj.hil_Street3 = addressData.AddressLine3;
                    }
                    int[] AddressType = { 1, 2, 3 };
                    int AddressTypeEnum = AddressType.Contains(addressData.AddressTypeEnum) ? addressData.AddressTypeEnum : 3;
                    entObj.hil_AddressType = new OptionSetValue(AddressTypeEnum);
                    if (addressData.IsDefault)
                    {
                        RemoveDefaultAddress(service, addressData.CustomerGuid);
                    }
                    entObj["hil_isdefault"] = addressData.IsDefault;
                    service.Update(entObj);

                    objAddress = new IoTAddressBookResultV1();
                    objAddress.AddressGuid = addressData.AddressGuid;
                    objAddress.CustomerGuid = addressData.CustomerGuid;
                    objAddress.AddressLine1 = addressData.AddressLine1;
                    objAddress.AddressLine2 = addressData.AddressLine2;
                    objAddress.AddressLine3 = addressData.AddressLine3;
                    objAddress.AddressTypeEnum = AddressTypeEnum;
                    if (AddressTypeEnum == 1)
                    {
                        objAddress.AddressType = "Home";
                    }
                    else if (AddressTypeEnum == 2)
                    {
                        objAddress.AddressType = "Office";
                    }
                    else
                    {
                        objAddress.AddressType = "Other";
                    }
                    objAddress.IsDefault = addressData.IsDefault;
                    objAddress.StatusCode = 200;
                    objAddress.StatusDescription = "OK.";
                }
                else
                {
                    objAddress = new IoTAddressBookResultV1 { StatusCode = 503, StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                objAddress = new IoTAddressBookResultV1 { StatusCode = 500, StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return objAddress;
        }
        //public IoTAddressBookResultV1 ECommerceBulkCustomerData(IoTAddressBookResultV1 addressData)
        //{
        //    IoTAddressBookResultV1 objAddress = new IoTAddressBookResultV1();
        //    Guid consumerGuId = Guid.Empty;
        //    QueryExpression query;
        //    EntityCollection entcoll;
        //    EntityReference erBusinessGeo = null;
        //    EntityReference erPinCode = null;
        //    EntityReference erArea = null;
        //    try
        //    {
        //        IOrganizationService service = ConnectToCRM.GetOrgService();
        //        if (addressData.MobileNumber == null || addressData.MobileNumber.Trim().Length == 0)
        //        {
        //            objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "Customer Mobile Number is required." };
        //            return objAddress;
        //        }
        //        if (addressData.CustomerName == null || addressData.CustomerName.Trim().Length == 0)
        //        {
        //            objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "Customer Name is required." };
        //            return objAddress;
        //        }
        //        if (addressData.AddressLine1 == null)
        //        {
        //            objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "Address Line 1 is required." };
        //            return objAddress;
        //        }
        //        if (addressData.AddressLine1.Trim().Length == 0)
        //        {
        //            objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "Address Line 1 is required." };
        //            return objAddress;
        //        }
        //        if (addressData.PINCode == null || addressData.PINCode.Trim().Length == 0)
        //        {
        //            objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "PIN Code is required." };
        //            return objAddress;
        //        }
        //        if (service != null)
        //        {
        //            query = new QueryExpression("hil_pincode");
        //            query.ColumnSet = new ColumnSet("hil_name");
        //            query.Criteria = new FilterExpression(LogicalOperator.And);
        //            query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, addressData.PINCode);
        //            entcoll = service.RetrieveMultiple(query);
        //            if (entcoll.Entities.Count > 0)
        //            {
        //                erPinCode = entcoll.Entities[0].ToEntityReference();
        //            }
        //            if (erPinCode == null)
        //            {
        //                objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "PIN Code does not exist." };
        //                return objAddress;
        //            }
        //            query = new QueryExpression("hil_businessmapping");
        //            query.ColumnSet = new ColumnSet("hil_area");
        //            query.Criteria = new FilterExpression(LogicalOperator.And);

        //            query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, erPinCode.Id);
        //            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

        //            query.AddOrder("createdon", OrderType.Ascending);
        //            query.TopCount = 1;

        //            entcoll = service.RetrieveMultiple(query);
        //            if (entcoll.Entities.Count == 0)
        //            {
        //                objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "Business Geo Mapping does not exist for this PIN Code." };
        //                return objAddress;
        //            }

        //            QueryExpression qsContact = new QueryExpression("contact");
        //            qsContact.ColumnSet = new ColumnSet("contactid", "fullname", "emailaddress1");
        //            ConditionExpression condExp = new ConditionExpression("mobilephone", ConditionOperator.Equal, addressData.MobileNumber);
        //            qsContact.Criteria.AddCondition(condExp);
        //            EntityCollection entColConsumer = service.RetrieveMultiple(qsContact);
        //            if (entColConsumer.Entities.Count > 0) //Consumer already Exists in D365 Database
        //            {
        //                consumerGuId = entColConsumer.Entities[0].Id;
        //                objAddress.CustomerGuid = consumerGuId;
        //                objAddress.MobileNumber = addressData.MobileNumber;
        //                if (entColConsumer.Entities[0].Contains("emailaddress1"))
        //                {
        //                    objAddress.EmailAddress = entColConsumer.Entities[0].GetAttributeValue<string>("emailaddress1");
        //                }
        //                if (entColConsumer.Entities[0].Contains("fullname"))
        //                {
        //                    objAddress.CustomerName = entColConsumer.Entities[0].GetAttributeValue<string>("fullname");
        //                }
        //            }
        //            else
        //            {
        //                Entity entConsumer = new Entity("contact");
        //                entConsumer["mobilephone"] = addressData.MobileNumber;
        //                entConsumer["hil_salutation"] = new OptionSetValue(2);

        //                string[] consumerName = addressData.CustomerName.Split(' ');
        //                if (consumerName.Length >= 1)
        //                {
        //                    entConsumer["firstname"] = consumerName[0];
        //                    if (consumerName.Length == 3)
        //                    {
        //                        entConsumer["middlename"] = consumerName[1];
        //                        entConsumer["lastname"] = consumerName[2];
        //                    }
        //                    if (consumerName.Length == 2)
        //                    {
        //                        entConsumer["lastname"] = consumerName[1];
        //                    }
        //                }
        //                else
        //                {
        //                    entConsumer["firstname"] = addressData.CustomerName;
        //                }

        //                if (addressData.EmailAddress != null && addressData.EmailAddress.Trim().Length > 0)
        //                {
        //                    entConsumer["emailaddress1"] = addressData.EmailAddress;
        //                }
        //                entConsumer["hil_consumersource"] = new OptionSetValue(7);
        //                consumerGuId = service.Create(entConsumer);
        //                objAddress.CustomerGuid = consumerGuId;
        //            }
        //            query = new QueryExpression("hil_businessmapping");
        //            query.ColumnSet = new ColumnSet("hil_area");
        //            query.Criteria = new FilterExpression(LogicalOperator.And);
        //            query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, erPinCode.Id);
        //            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
        //            query.AddOrder("createdon", OrderType.Ascending);
        //            query.TopCount = 1;

        //            entcoll = service.RetrieveMultiple(query);
        //            if (entcoll.Entities.Count > 0)
        //            {
        //                erBusinessGeo = entcoll.Entities[0].ToEntityReference();
        //                erArea = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_area");
        //                objAddress.PINCode = addressData.PINCode;
        //                objAddress.PINCodeGuid = erPinCode.Id;
        //                objAddress.Area = erArea.Name;
        //                objAddress.AreaGuid = erArea.Id;
        //            }
        //            hil_address entObj = new hil_address();
        //            entObj.hil_Street1 = addressData.AddressLine1;
        //            if (addressData.AddressLine2 != null)
        //            {
        //                entObj.hil_Street2 = addressData.AddressLine2;
        //            }
        //            if (addressData.AddressLine3 != null)
        //            {
        //                entObj.hil_Street3 = addressData.AddressLine3;
        //            }
        //            entObj.hil_Customer = new EntityReference("contact", objAddress.CustomerGuid);
        //            if (erBusinessGeo != null)
        //            {
        //                entObj.hil_BusinessGeo = erBusinessGeo;
        //            }
        //            int[] AddressType = { 1, 2, 3 };
        //            int AddressTypeEnum = AddressType.Contains(addressData.AddressTypeEnum) ? addressData.AddressTypeEnum : 3;
        //            entObj.hil_AddressType = new OptionSetValue(AddressTypeEnum);
        //            if (objAddress.IsDefault)
        //            {
        //                RemoveDefaultAddress(service, objAddress.CustomerGuid);
        //                entObj["hil_isdefault"] = addressData.IsDefault;
        //            }
        //            objAddress.AddressGuid = service.Create(entObj);
        //            objAddress.AddressTypeEnum = AddressTypeEnum;
        //            if (AddressTypeEnum == 1)
        //            {
        //                objAddress.AddressType = "Home";
        //            }
        //            else if (AddressTypeEnum == 2)
        //            {
        //                objAddress.AddressType = "Office";
        //            }
        //            else
        //            {
        //                objAddress.AddressType = "Other";
        //            }
        //            objAddress.StatusCode = 200;
        //            objAddress.StatusDescription = "OK";
        //            return objAddress;
        //        }
        //        else
        //        {
        //            objAddress = new IoTAddressBookResultV1 { StatusCode = 503, StatusDescription = "D365 Service Unavailable." };
        //            return objAddress;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        objAddress = new IoTAddressBookResultV1 { StatusCode = 500, StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
        //        return objAddress;
        //    }
        //}
        public IoTAddressBookResultV1 IoTDeleteAddress(IoTAddressBookResultV1 addressData)
        {
            IoTAddressBookResultV1 objAddress = null;
            QueryExpression query;
            EntityCollection entcoll;

            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (addressData.AddressGuid == Guid.Empty)
                {
                    objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "Address GUID is required." };
                    return objAddress;
                }

                if (service != null)
                {
                    query = new QueryExpression("hil_address");
                    query.ColumnSet = new ColumnSet("statecode");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_addressid", ConditionOperator.Equal, addressData.AddressGuid);
                    entcoll = service.RetrieveMultiple(query);
                    if (entcoll.Entities.Count == 0)
                    {
                        objAddress = new IoTAddressBookResultV1 { AddressGuid = addressData.AddressGuid, StatusCode = 204, StatusDescription = "Address does not exist in D365." };
                        return objAddress;
                    }
                    else
                    {
                        if (entcoll.Entities[0].GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                        {
                            objAddress = new IoTAddressBookResultV1 { AddressGuid = addressData.AddressGuid, StatusCode = 204, StatusDescription = "Address is already deleted." };
                            return objAddress;
                        }
                    }
                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorder'>
                            <attribute name='msdyn_name' />
                            <filter type='and'>
                                <condition attribute='hil_serviceaddress' operator='eq' value='{addressData.AddressGuid}' />
                            </filter>
                        </entity>
                        </fetch>";
                    entcoll = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entcoll.Entities.Count > 0)
                    {
                        objAddress = new IoTAddressBookResultV1 { AddressGuid = addressData.AddressGuid, StatusCode = 204, StatusDescription = $"Address is in use against JobId:{entcoll.Entities[0].GetAttributeValue<string>("msdyn_name")}.Can't be deleted." };
                        return objAddress;
                    }
                    else
                    {
                        Entity entity = new Entity("hil_address", addressData.AddressGuid);
                        entity["statecode"] = new OptionSetValue(1); //InActive
                        entity["statuscode"] = new OptionSetValue(2); //InActive
                        service.Update(entity);
                    }
                    objAddress = new IoTAddressBookResultV1 { AddressGuid = addressData.AddressGuid, StatusCode = 200, StatusDescription = "Success" };
                }
                else
                {
                    objAddress = new IoTAddressBookResultV1 { StatusCode = 503, StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                objAddress = new IoTAddressBookResultV1 { StatusCode = 500, StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return objAddress;
        }
        public void RemoveDefaultAddress(IOrganizationService service, Guid CustomerGuid)
        {
            QueryExpression query = new QueryExpression("hil_address");
            query.ColumnSet = new ColumnSet("hil_customer", "hil_addressid", "hil_addresstype", "hil_isdefault");
            query.Criteria = new FilterExpression(LogicalOperator.And);
            query.Criteria.AddCondition("hil_isdefault", ConditionOperator.Equal, true);
            query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, CustomerGuid);
            EntityCollection collection = service.RetrieveMultiple(query);

            if (collection.Entities.Count > 0)
            {
                foreach (var item in collection.Entities)
                {
                    Entity entity = new Entity("hil_address", item.Id);
                    entity["hil_isdefault"] = false;
                    service.Update(entity);
                }
            }
        }
    }

    [DataContract]
    public class IoTAddressBookResultV1
    {
        [DataMember]
        public Guid CustomerGuid { get; set; }

        [DataMember]
        public Guid AddressGuid { get; set; }

        [DataMember]
        public string MobileNumber { get; set; }

        [DataMember]
        public string CustomerName { get; set; }

        [DataMember]
        public string EmailAddress { get; set; }

        [DataMember]
        public string AddressLine1 { get; set; }

        [DataMember]
        public string AddressLine2 { get; set; }

        [DataMember]
        public string AddressLine3 { get; set; }

        [DataMember]
        public string AddressPhone { get; set; }

        [DataMember]
        public Guid BizGeoGuid { get; set; }

        [DataMember]
        public string BizGeoName { get; set; }

        [DataMember]
        public Guid PINCodeGuid { get; set; }

        [DataMember]
        public string PINCode { get; set; }

        [DataMember]
        public string Area { get; set; }

        [DataMember]
        public Guid AreaGuid { get; set; }

        [DataMember]
        public string FullAddress { get; set; }

        [DataMember]
        public string AddressType { get; set; }

        [DataMember]
        public int AddressTypeEnum { get; set; } = 3;
        [DataMember]
        public bool IsDefault { get; set; }
        [DataMember]
        public int StatusCode { get; set; }

        [DataMember]
        public string StatusDescription { get; set; }
    }

    //[DataContract]
    //public class IoTPINCodes
    //{
    //    [DataMember]
    //    public Guid PINCodeGuid { get; set; }
    //    [DataMember]
    //    public string PINCode { get; set; }
    //    [DataMember]
    //    public int StatusCode { get; set; }
    //    [DataMember]
    //    public string StatusDescription { get; set; }
    //}

    //[DataContract]
    //public class IoTAreas
    //{
    //    [DataMember]
    //    public Guid PINCodeGuid { get; set; }
    //    [DataMember]
    //    public Guid AreaGuid { get; set; }
    //    [DataMember]
    //    public string AreaName { get; set; }
    //    [DataMember]
    //    public int StatusCode { get; set; }
    //    [DataMember]
    //    public string StatusDescription { get; set; }
    //}
}
