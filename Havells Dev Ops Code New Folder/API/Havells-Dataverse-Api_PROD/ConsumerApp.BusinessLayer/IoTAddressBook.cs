using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class IoTAddressBook
    {
        [DataMember]
        public Guid CustomerGuid { get; set; }
        [DataMember]
        public string Pincode { get; set; }
        [DataMember]
        public string MobileNumber { get; set; }

        public List<IoTAddressBookResult> GetIoTAddressBook(IoTAddressBook address)
        {
            List<IoTAddressBookResult> addressList = new List<IoTAddressBookResult>();
            IoTAddressBookResult objAddress;
            QueryExpression query;
            EntityCollection collection;
            EntityReference erPincode = null;
            try
            {
                if (address.CustomerGuid == Guid.Empty && address.MobileNumber == null)
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Customer GUID/Mobile Number is required." };
                    addressList.Add(objAddress);
                    return addressList;
                }
                IOrganizationService service = ConnectToCRM.GetOrgService();
                
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
                                objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Invalid Pincode." };
                                addressList.Add(objAddress);
                                return addressList;
                            }
                        }
                        else
                        {
                            objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Invalid Pincode." };
                            addressList.Add(objAddress);
                            return addressList;
                        }
                    }

                    query = new QueryExpression("hil_address");
                    query.ColumnSet = new ColumnSet("hil_customer", "hil_addressid", "hil_street3", "hil_street2", "hil_street1", "hil_fulladdress",
                        "hil_pincode", "hil_area", "hil_businessgeo", "hil_addresstype", "hil_city", "hil_state");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
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
                        lnk.LinkCriteria.AddCondition("mobilephone", ConditionOperator.Equal, address.MobileNumber);
                        query.LinkEntities.Add(lnk);
                    }
                    collection = service.RetrieveMultiple(query);

                    if (collection.Entities.Count == 0)
                    {
                        objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "No Address Book found." };
                        addressList.Add(objAddress);
                        return addressList;
                    }

                    //query = new QueryExpression()
                    //{
                    //    EntityName = hil_address.EntityLogicalName,
                    //    ColumnSet = new ColumnSet("hil_addressid", "hil_street3", "hil_street2", "hil_street1", "hil_fulladdress", "hil_pincode", "hil_area", "hil_businessgeo", "hil_addresstype")
                    //};
                    //FilterExpression filterExpression = new FilterExpression(LogicalOperator.And);
                    //filterExpression.Conditions.Add(new ConditionExpression("hil_customer", ConditionOperator.Equal, address.CustomerGuid));
                    //query.Criteria.AddFilter(filterExpression);

                    //collection = service.RetrieveMultiple(query);

                    if (collection.Entities != null && collection.Entities.Count > 0)
                    {
                        foreach (Entity item in collection.Entities)
                        {
                            objAddress = new IoTAddressBookResult();
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
                                OptionSetValue osv = item.GetAttributeValue<OptionSetValue>("hil_addresstype");
                                objAddress.AddressType = osv.Value == 1 ? "Permanent" : "Alternate";
                            }
                            if (item.Attributes.Contains("hil_addresstype"))
                            {
                                objAddress.AddressTypeEnum = item.GetAttributeValue<OptionSetValue>("hil_addresstype").Value.ToString();
                            }
                            if (item.Attributes.Contains("hil_city"))
                            {
                                objAddress.CityName = item.GetAttributeValue<EntityReference>("hil_city").Name;
                            }
                            if (item.Attributes.Contains("hil_state"))
                            {
                                objAddress.StateName = item.GetAttributeValue<EntityReference>("hil_state").Name;
                            }
                            objAddress.CustomerGuid = address.CustomerGuid;
                            objAddress.MobileNumber = address.MobileNumber;
                            objAddress.StatusCode = "200";
                            objAddress.StatusDescription = "OK";
                            addressList.Add(objAddress);
                        }
                    }
                    return addressList;
                }
                else
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    addressList.Add(objAddress);
                }
            }
            catch (Exception ex)
            {
                objAddress = new IoTAddressBookResult { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                addressList.Add(objAddress);
            }
            return addressList;
        }

        public IoTPINCodes IoTValidatePINCode(IoTPINCodes pinCode)
        {
            IoTPINCodes objPINCode = null;
            try
            {
                if (pinCode.PINCode.Trim().Length == 0)
                {
                    objPINCode = new IoTPINCodes { StatusCode = "204", StatusDescription = "PIN Code is required." };
                    return objPINCode;
                }
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    QueryExpression query = new QueryExpression("hil_pincode");
                    query.ColumnSet = new ColumnSet("hil_pincodeid", "hil_name");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, pinCode.PINCode);
                    EntityCollection entcoll = service.RetrieveMultiple(query);
                    if (entcoll.Entities.Count == 0)
                    {
                        objPINCode = new IoTPINCodes { StatusCode = "204", StatusDescription = "No PIN Code found." };
                        return objPINCode;
                    }
                    else
                    {
                        objPINCode = new IoTPINCodes { PINCode = entcoll.Entities[0].GetAttributeValue<string>("hil_name"), PINCodeGuid = entcoll.Entities[0].GetAttributeValue<Guid>("hil_pincodeid"), StatusCode = "200", StatusDescription = "OK." };
                        return objPINCode;
                    }
                }
                else
                {
                    objPINCode = new IoTPINCodes { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                objPINCode = new IoTPINCodes { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return objPINCode;
        }

        public List<IoTAreas> IoTGetAreasByPinCode(IoTPINCodes pinCode)
        {
            IoTAreas objArea = null;
            List<IoTAreas> lstAreas = null;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    QueryExpression query = new QueryExpression("hil_businessmapping");
                    query.ColumnSet = new ColumnSet("hil_pincode", "hil_area");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, pinCode.PINCodeGuid);
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                    EntityCollection entcoll = service.RetrieveMultiple(query);
                    if (entcoll.Entities.Count == 0)
                    {
                        objArea = new IoTAreas { StatusCode = "204", StatusDescription = "No Area found." };
                        lstAreas = new List<IoTAreas>();
                        lstAreas.Add(objArea);
                    }
                    else
                    {
                        lstAreas = new List<IoTAreas>();
                        foreach (Entity ent in entcoll.Entities)
                        {
                            if (ent.Attributes.Contains("hil_area"))
                            {
                                lstAreas.Add(new IoTAreas { PINCodeGuid = pinCode.PINCodeGuid, AreaName = ent.GetAttributeValue<EntityReference>("hil_area").Name, AreaGuid = ent.GetAttributeValue<EntityReference>("hil_area").Id, StatusCode = "200", StatusDescription = "OK." });
                            }
                        }
                    }
                }
                else
                {
                    objArea = new IoTAreas { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    lstAreas = new List<IoTAreas>();
                    lstAreas.Add(objArea);
                }
            }
            catch (Exception ex)
            {
                objArea = new IoTAreas { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                lstAreas = new List<IoTAreas>();
                lstAreas.Add(objArea);
            }
            return lstAreas;
        }

        public IoTAddressBookResult IoTCreateAddress(IoTAddressBookResult addressData)
        {
            IoTAddressBookResult objAddress = null;
            QueryExpression query;
            EntityCollection entcoll;
            EntityReference businessGeo = null;

            try
            {
                if (addressData.CustomerGuid == Guid.Empty)
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Customer GUID is required." };
                    return objAddress;
                }
                if (addressData.AddressLine1 == null)
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Address Line 1 is required." };
                    return objAddress;
                }
                if (addressData.AddressLine1.Trim().Length == 0)
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Address Line 1 is required." };
                    return objAddress;
                }
                if (addressData.PINCodeGuid == Guid.Empty)
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "PIN Code is required." };
                    return objAddress;
                }
                //if (addressData.AreaGuid == Guid.Empty)
                //{
                //    objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Area is required." };
                //    return objAddress;
                //}

                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    int addressType;
                    if (addressData.AddressTypeEnum == null)
                    {
                        query = new QueryExpression("hil_address");
                        query.ColumnSet = new ColumnSet("hil_addressid");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, addressData.CustomerGuid);
                        entcoll = service.RetrieveMultiple(query);
                        if (entcoll.Entities.Count > 0) { addressType = 2; }
                        else { addressType = 1; }
                    }
                    else
                    {
                        if (addressData.AddressTypeEnum == "1")
                        {
                            addressType = 1;
                        }
                        else
                        {
                            addressType = 2;
                        }
                    }

                    query = new QueryExpression("hil_businessmapping");
                    query.ColumnSet = new ColumnSet("hil_pincode");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, addressData.PINCodeGuid);
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
                    }
                    else
                    {
                        objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "PIN Code or Area does not exist." };
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
                    entObj.hil_AddressType = new OptionSetValue(addressType);
                    objAddress = new IoTAddressBookResult();
                    objAddress.AddressGuid = service.Create(entObj);
                    if (objAddress.AddressGuid != Guid.Empty)
                    {
                        objAddress.CustomerGuid = addressData.CustomerGuid;
                        objAddress.StatusCode = "200";
                        objAddress.StatusDescription = "OK.";
                    }
                    else
                    {
                        objAddress.CustomerGuid = addressData.CustomerGuid;
                        objAddress.StatusCode = "204";
                        objAddress.StatusDescription = "Something went wrong.";
                    }
                }
                else
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                objAddress = new IoTAddressBookResult { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return objAddress;
        }

        public IoTAddressBookResult IoTUpdateAddress(IoTAddressBookResult addressData)
        {
            IoTAddressBookResult objAddress = null;
            QueryExpression query;
            EntityCollection entcoll;
            //EntityReference businessGeo = null;

            try
            {
                if (addressData.AddressGuid == Guid.Empty)
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Address GUID is required." };
                    return objAddress;
                }
                if (addressData.CustomerGuid == Guid.Empty)
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Customer GUID is required." };
                    return objAddress;
                }
                if (addressData.AddressLine1.Trim().Length == 0)
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Address Line 1 is required." };
                    return objAddress;
                }
                //if (addressData.PINCodeGuid == Guid.Empty)
                //{
                //    objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "PIN Code is required." };
                //    return objAddress;
                //}
                //if (addressData.AreaGuid == Guid.Empty)
                //{
                //    objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Area is required." };
                //    return objAddress;
                //}

                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    query = new QueryExpression("hil_address");
                    query.ColumnSet = new ColumnSet("hil_addressid");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, addressData.CustomerGuid);
                    query.Criteria.AddCondition("hil_addressid", ConditionOperator.Equal, addressData.AddressGuid);
                    entcoll = service.RetrieveMultiple(query);
                    if (entcoll.Entities.Count == 0)
                    {
                        objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Address ID does not belong to Customer." };
                        return objAddress;
                    }

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
                    entObj.hil_Customer = new EntityReference("contact", addressData.CustomerGuid);
                    //if (businessGeo != null)
                    //{
                    //    entObj.hil_BusinessGeo = businessGeo;
                    //}
                    if (addressData.AddressTypeEnum == "1")
                    {
                        entObj.hil_AddressType = new OptionSetValue(1); //Permanent
                    }
                    else
                    {
                        entObj.hil_AddressType = new OptionSetValue(2); //Alternate
                    }
                    service.Update(entObj);

                    objAddress = new IoTAddressBookResult();
                    objAddress.AddressGuid = addressData.AddressGuid;
                    objAddress.CustomerGuid = addressData.CustomerGuid;
                    objAddress.StatusCode = "200";
                    objAddress.StatusDescription = "OK.";
                }
                else
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                objAddress = new IoTAddressBookResult { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return objAddress;
        }

        public IoTAddressBookResult ECommerceBulkCustomerData(IoTAddressBookResult addressData)
        {
            IoTAddressBookResult objAddress = new IoTAddressBookResult();
            Guid consumerGuId = Guid.Empty;
            QueryExpression query;
            EntityCollection entcoll;
            EntityReference erBusinessGeo = null;
            EntityReference erPinCode = null;
            EntityReference erArea = null;
            try
            {
                if (addressData.MobileNumber == null || addressData.MobileNumber.Trim().Length == 0)
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Customer Mobile Number is required." };
                    return objAddress;
                }
                if (addressData.CustomerName == null || addressData.CustomerName.Trim().Length == 0)
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Customer Name is required." };
                    return objAddress;
                }
                if (addressData.AddressLine1 == null)
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Address Line 1 is required." };
                    return objAddress;
                }
                if (addressData.AddressLine1.Trim().Length == 0)
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Address Line 1 is required." };
                    return objAddress;
                }
                if (addressData.PINCode == null || addressData.PINCode.Trim().Length == 0)
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "PIN Code is required." };
                    return objAddress;
                }
                if (addressData.AddressTypeEnum == null || addressData.AddressTypeEnum.Trim().Length == 0)
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Address Type is required." };
                    return objAddress;
                }
                IOrganizationService service = ConnectToCRM.GetOrgService(); //get org service obj for connection
                if (service != null)
                {
                    query = new QueryExpression("hil_pincode");
                    query.ColumnSet = new ColumnSet("hil_name");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, addressData.PINCode);
                    entcoll = service.RetrieveMultiple(query);
                    if (entcoll.Entities.Count > 0)
                    {
                        erPinCode = entcoll.Entities[0].ToEntityReference();
                    }
                    if (erPinCode == null)
                    {
                        objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "PIN Code does not exist." };
                        return objAddress;
                    }
                    query = new QueryExpression("hil_businessmapping");
                    query.ColumnSet = new ColumnSet("hil_area");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, erPinCode.Id);
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                    query.AddOrder("createdon", OrderType.Ascending);
                    query.TopCount = 1;

                    entcoll = service.RetrieveMultiple(query);
                    if (entcoll.Entities.Count == 0)
                    {
                        objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Business Geo Mapping does not exist for this PIN Code." };
                        return objAddress;
                    }

                    QueryExpression qsContact = new QueryExpression("contact");
                    qsContact.ColumnSet = new ColumnSet("contactid", "fullname", "emailaddress1");
                    ConditionExpression condExp = new ConditionExpression("mobilephone", ConditionOperator.Equal, addressData.MobileNumber);
                    qsContact.Criteria.AddCondition(condExp);
                    EntityCollection entColConsumer = service.RetrieveMultiple(qsContact);
                    if (entColConsumer.Entities.Count > 0) //Consumer already Exists in D365 Database
                    {
                        consumerGuId = entColConsumer.Entities[0].Id;
                        objAddress.CustomerGuid = consumerGuId;
                        objAddress.MobileNumber = addressData.MobileNumber;
                        if (entColConsumer.Entities[0].Contains("emailaddress1"))
                        {
                            objAddress.EmailAddress = entColConsumer.Entities[0].GetAttributeValue<string>("emailaddress1");
                        }
                        if (entColConsumer.Entities[0].Contains("fullname"))
                        {
                            objAddress.CustomerName = entColConsumer.Entities[0].GetAttributeValue<string>("fullname");
                        }
                    }
                    else
                    {
                        Entity entConsumer = new Entity("contact");
                        entConsumer["mobilephone"] = addressData.MobileNumber;
                        entConsumer["hil_salutation"] = new OptionSetValue(2);

                        string[] consumerName = addressData.CustomerName.Split(' ');
                        if (consumerName.Length >= 1)
                        {
                            entConsumer["firstname"] = consumerName[0];
                            if (consumerName.Length == 3)
                            {
                                entConsumer["middlename"] = consumerName[1];
                                entConsumer["lastname"] = consumerName[2];
                            }
                            if (consumerName.Length == 2)
                            {
                                entConsumer["lastname"] = consumerName[1];
                            }
                        }
                        else
                        {
                            entConsumer["firstname"] = addressData.CustomerName;
                        }

                        if (addressData.EmailAddress != null && addressData.EmailAddress.Trim().Length > 0)
                        {
                            entConsumer["emailaddress1"] = addressData.EmailAddress;
                        }
                        entConsumer["hil_consumersource"] = new OptionSetValue(7);
                        consumerGuId = service.Create(entConsumer);
                        objAddress.CustomerGuid = consumerGuId;
                    }
                    query = new QueryExpression("hil_businessmapping");
                    query.ColumnSet = new ColumnSet("hil_area");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, erPinCode.Id);
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                    query.AddOrder("createdon", OrderType.Ascending);
                    query.TopCount = 1;

                    entcoll = service.RetrieveMultiple(query);
                    if (entcoll.Entities.Count > 0)
                    {
                        erBusinessGeo = entcoll.Entities[0].ToEntityReference();
                        erArea = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_area");
                        objAddress.PINCode = addressData.PINCode;
                        objAddress.PINCodeGuid = erPinCode.Id;
                        objAddress.Area = erArea.Name;
                        objAddress.AreaGuid = erArea.Id;
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
                    entObj.hil_Customer = new EntityReference("contact", objAddress.CustomerGuid);
                    if (erBusinessGeo != null)
                    {
                        entObj.hil_BusinessGeo = erBusinessGeo;
                    }
                    entObj.hil_AddressType = new OptionSetValue(Convert.ToInt32(addressData.AddressTypeEnum));
                    objAddress.AddressGuid = service.Create(entObj);
                    objAddress.StatusCode = "200";
                    objAddress.StatusDescription = "OK";
                    return objAddress;
                }
                else
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "503", StatusDescription = "D365 Service Unavailable." };
                    return objAddress;
                }
            }
            catch (Exception ex)
            {
                objAddress = new IoTAddressBookResult { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                return objAddress;
            }
        }
    }

    [DataContract]
    public class IoTAddressBookResult
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
        public string AddressTypeEnum { get; set; }

        [DataMember]
        public string StatusCode { get; set; }

        [DataMember]
        public string StatusDescription { get; set; }
        [DataMember]
        public string CityName { get; set; }

        [DataMember]
        public string StateName { get; set; }

    }

    [DataContract]
    public class IoTPINCodes
    {
        [DataMember]
        public Guid PINCodeGuid { get; set; }
        [DataMember]
        public string PINCode { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }
    }

    [DataContract]
    public class IoTAreas
    {
        [DataMember]
        public Guid PINCodeGuid { get; set; }
        [DataMember]
        public Guid AreaGuid { get; set; }
        [DataMember]
        public string AreaName { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }
    }
}
