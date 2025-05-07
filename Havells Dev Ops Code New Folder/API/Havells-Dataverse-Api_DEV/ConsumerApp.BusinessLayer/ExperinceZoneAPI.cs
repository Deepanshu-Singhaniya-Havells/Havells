using System;
using Microsoft.Xrm.Sdk;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class ExperinceZoneAPI
    {
        public ExperinceZonePayLoad RegisterConsumer(ExperinceZonePayLoad consumer)
        {
            ExperinceZonePayLoad returnResult = new ExperinceZonePayLoad();
            Entity _contactEntity = new Entity();
            try
            {
                //IOrganizationService _service = null;
                IOrganizationService _service = ConnectToCRM.GetOrgService();
                if (consumer.MobileNumber == null || consumer.MobileNumber.Trim().Length == 0)
                {
                    returnResult.StatusCode = "204";
                    returnResult.StatusDescription = "No Content : Mobile Number is required.";
                    return returnResult;
                }
                else
                {
                    if (_service != null)
                    {
                        QueryExpression query = new QueryExpression("contact");
                        query.ColumnSet = new ColumnSet("contactid", "fullname", "emailaddress1", "mobilephone", "adx_organizationname");
                        ConditionExpression condExp = new ConditionExpression("mobilephone", ConditionOperator.Equal, consumer.MobileNumber.Trim());
                        query.Criteria.AddCondition(condExp);
                        EntityCollection entcoll = _service.RetrieveMultiple(query);
                        if (consumer.Operation.ToLower() == "verify")
                        {
                            if (entcoll.Entities.Count > 0)
                            {
                                returnResult.StatusCode = "200";
                                returnResult.StatusDescription = "Consumer already exist.";
                                returnResult.FullName = entcoll[0].Contains("fullname") ? entcoll[0].GetAttributeValue<string>("fullname") : "";
                                returnResult.MobileNumber = entcoll[0].Contains("mobilephone") ? entcoll[0].GetAttributeValue<string>("mobilephone") : "";
                                returnResult.Email = entcoll[0].Contains("emailaddress1") ? entcoll[0].GetAttributeValue<string>("emailaddress1") : "";
                                returnResult.OrganizationName = entcoll[0].Contains("adx_organizationname") ? entcoll[0].GetAttributeValue<string>("adx_organizationname") : "";
                                returnResult.consumerGuId = entcoll[0].Id.ToString();

                                #region Getting Latest Address
                                QueryExpression queryExp = new QueryExpression("hil_address");
                                queryExp.ColumnSet = new ColumnSet("hil_pincode", "hil_street1", "hil_fulladdress");
                                queryExp.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, entcoll[0].Id);
                                queryExp.AddOrder("createdon", OrderType.Descending);
                                queryExp.TopCount = 1;
                                EntityCollection entCollAddress = _service.RetrieveMultiple(queryExp);
                                if (entCollAddress.Entities.Count > 0)
                                {
                                    returnResult.Address = entCollAddress.Entities[0].Contains("hil_fulladdress") ? entCollAddress.Entities[0].GetAttributeValue<string>("hil_fulladdress") : (entCollAddress.Entities[0].Contains("hil_street1") ? entCollAddress.Entities[0].GetAttributeValue<string>("hil_street1") : null);
                                    returnResult.pinCode = entCollAddress.Entities[0].Contains("hil_pincode") ? entCollAddress.Entities[0].GetAttributeValue<EntityReference>("hil_pincode").Name : null;
                                }
                                returnResult.Operation = consumer.Operation;
                                #endregion
                                return returnResult;
                            }
                            else
                            {
                                returnResult.StatusCode = "200";
                                returnResult.StatusDescription = "Consumer dose not exist.";
                                returnResult.FullName = "";
                                returnResult.MobileNumber = "";
                                returnResult.Email = "";
                                returnResult.OrganizationName = "";
                            }
                        }
                        else
                        {
                            if (entcoll.Entities.Count > 0)
                                _contactEntity = entcoll[0];

                            if (_contactEntity.Id== Guid.Empty)
                            {
                                if (consumer.SourceOfCreation == null)
                                {
                                    returnResult.StatusCode = "204";
                                    returnResult.StatusDescription = "No Content : Source of Registration is required. Please pass <4> for Whatsapp <5> for IoT Platform <7> for eCommerce";
                                    return returnResult;
                                }
                                else if (consumer.pinCode == null)
                                {
                                    returnResult.StatusCode = "204";
                                    returnResult.StatusDescription = "No Content : Pincode  is required.";
                                    return returnResult;
                                }
                                query = new QueryExpression("hil_pincode");
                                query.ColumnSet = new ColumnSet("hil_pincodeid", "hil_name");
                                query.Criteria = new FilterExpression(LogicalOperator.And);
                                query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, consumer.pinCode);
                                entcoll = _service.RetrieveMultiple(query);
                                if (entcoll.Entities.Count != 1)
                                {
                                    returnResult.StatusCode = "204";
                                    returnResult.StatusDescription = "PinCode does not exist in system.";
                                    return returnResult;
                                }

                                query = new QueryExpression("hil_businessmapping");
                                query.ColumnSet = new ColumnSet("hil_pincode", "hil_area", "hil_city", "hil_district", "hil_state", "hil_region", "hil_branch", "hil_salesoffice", "hil_subterritory");
                                query.Criteria = new FilterExpression(LogicalOperator.And);
                                query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, entcoll[0].Id);
                                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                entcoll = _service.RetrieveMultiple(query);
                                if (entcoll.Entities.Count < 0)
                                {
                                    returnResult.StatusCode = "204";
                                    returnResult.StatusDescription = "PinCode does not exist in system.";
                                    return returnResult;
                                }
                                Entity _pin = entcoll[0];

                                Entity _contactEntityCreate = new Entity("contact");
                                if (consumer.FullName != null && consumer.FullName != "")
                                    _contactEntityCreate["firstname"] = consumer.FullName;
                                if (consumer.Email != null && consumer.Email != "")
                                    _contactEntityCreate["emailaddress1"] = consumer.Email;
                                if (consumer.OrganizationName != null && consumer.OrganizationName != "")
                                    _contactEntityCreate["adx_organizationname"] = consumer.OrganizationName;
                                if (consumer.MobileNumber != null && consumer.MobileNumber != "")
                                    _contactEntityCreate["mobilephone"] = consumer.MobileNumber;
                                _contactEntityCreate["hil_consumersource"] = new OptionSetValue((int)consumer.SourceOfCreation);
                                Guid newContactID = _service.Create(_contactEntityCreate);

                                Entity _address = new Entity("hil_address");
                                _address["hil_customer"] = new EntityReference("contact", newContactID);
                                _address["hil_street1"] = consumer.Address;
                                _address["hil_businessgeo"] = (EntityReference)_pin.ToEntityReference();
                                _address["hil_state"] = _pin["hil_state"];
                                _address["hil_city"] = _pin["hil_city"];
                                _address["hil_district"] = _pin["hil_district"];
                                _address["hil_area"] = _pin["hil_area"];
                                _address["hil_branch"] = _pin["hil_branch"];
                                _address["hil_pincode"] = _pin["hil_pincode"];
                                _address["hil_salesoffice"] = _pin["hil_salesoffice"];
                                Guid newAddressID = _service.Create(_address);
                                returnResult.StatusCode = "200";
                                returnResult.consumerGuId = newContactID.ToString();
                                returnResult.StatusDescription = "Consumer Created Sucessfully.";
                                return returnResult;
                            }
                            else
                            {
                                
                                Entity _contactEntityUpdate = new Entity(_contactEntity.LogicalName);
                                if (consumer.FullName != null && consumer.FullName != "")
                                    _contactEntityUpdate["fullname"] = consumer.FullName;
                                if (consumer.Email != null && consumer.Email != "")
                                    _contactEntityUpdate["emailaddress1"] = consumer.Email;
                                if (consumer.OrganizationName != null && consumer.OrganizationName != "")
                                    _contactEntityUpdate["adx_organizationname"] = consumer.OrganizationName;
                                _contactEntityUpdate.Id = _contactEntity.Id;
                                _service.Update(_contactEntityUpdate);
                                returnResult.StatusCode = "200";
                                returnResult.StatusDescription = "Consumer Updated Sucessfully.";
                                returnResult.consumerGuId = _contactEntityUpdate.Id.ToString();
                                return returnResult;
                            }
                        }
                    }
                    else
                    {
                        returnResult.StatusCode = "503";
                        returnResult.StatusDescription = "D365 Service Unavailable";
                        return returnResult;
                    }
                }
            }
            catch (Exception ex)
            {
                returnResult.StatusCode = "500";
                returnResult.StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper();
            }
            return returnResult;
        }
    }

    [DataContract]
    public class ExperinceZonePayLoad
    {
        [DataMember]
        public string Operation { get; set; }
        [DataMember]
        public string FullName { get; set; }
        [DataMember]
        public string Email { get; set; }
        [DataMember]
        public string MobileNumber { get; set; }
        [DataMember]
        public string OrganizationName { get; set; }
        [DataMember]
        public string Address { get; set; }
        [DataMember]
        public string pinCode { get; set; }
        [DataMember]
        public int? SourceOfCreation { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }
        [DataMember]
        public string consumerGuId { get; set; }
    }
}
